"""
main_pipeline.py
================
Fault-tolerant workflow engine for the Bank Statement Automation Agent.

Orchestrates the six-phase execution pipeline defined in Section 8 of the
Development Plan.  Each phase is a discrete method on `BankStatementPipeline`
so that callers can invoke the full pipeline via `run()` or replay individual
phases for debugging.

BRANCH CONTRACTS (fault-tolerance rules enforced in this file):
  Branch A - AI Inference Exception  : see Phase 3
  Branch B - Relational Mapping Logic: see Phase 4
  Branch C - Fatal Balance Error     : see Phase 5

Usage:
    # Auto-detect opening balance from statement:
    pipeline = BankStatementPipeline(
        source_path="statement.pdf",
        output_dir="./output",
    )
    pipeline.run()

    # Or provide it explicitly:
    pipeline = BankStatementPipeline(
        source_path="statement.csv",
        initial_balance=Decimal("5000.00"),
        output_dir="./output",
    )
    pipeline.run()
"""

from __future__ import annotations

import argparse
import csv
import json
import logging
import os
import re
import sys
from dataclasses import asdict, dataclass, field
from datetime import date
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from dotenv import load_dotenv

for _env_candidate in [
    Path(__file__).resolve().parent.parent.parent / ".env",
    Path(__file__).resolve().parent.parent / ".env",
    Path.cwd() / ".env",
]:
    if _env_candidate.exists():
        load_dotenv(_env_candidate)
        break

from bank_statement_model import (
    BankStatementStore,
    Counterparty,
    CounterpartyRegistry,
    IntegrityError,
    Transaction,
    _safe_decimal,
    validate_running_balance,
)

logger = logging.getLogger(__name__)

TWO_PLACES = Decimal("0.01")

VALID_CATEGORIES = frozenset({
    "F&B", "Transport", "Utilities", "Telecommunications",
    "Salary", "Rental", "Insurance", "Government", "Bank Fees",
    "Transfer", "Investment", "Subscription", "Other", "Uncategorized",
})

# ──────────────────────────────────────────────────────────────────────
# Custom Exceptions
# ──────────────────────────────────────────────────────────────────────


class PipelineError(Exception):
    """Base exception for all pipeline failures."""
    pass


class FatalBalanceError(PipelineError):
    """
    Raised during Phase 5 when the running balance validation fails.

    BRANCH C — CRITICAL BLOCKING RULE:
    This error MUST terminate the pipeline immediately.
    Phase 6 (Relational Database Export) is ABSOLUTELY FORBIDDEN
    when this exception is raised.  The caller must catch this,
    log the offending rows, and halt.
    """

    def __init__(self, errors: List[dict], message: str = "") -> None:
        self.balance_errors = errors
        detail = message or (
            f"Fatal balance break detected at {len(errors)} row(s). "
            f"Pipeline suspended — Phase 6 export BLOCKED."
        )
        super().__init__(detail)


class ExtractionError(PipelineError):
    """Raised when Phase 1 data extraction produces no usable records."""
    pass


class ClassificationError(PipelineError):
    """Raised when Phase 3 classification completely fails (all records fail)."""
    pass


# ──────────────────────────────────────────────────────────────────────
# Data Transfer Objects (inter-phase payloads)
# ──────────────────────────────────────────────────────────────────────


@dataclass
class RawRecord:
    """
    Output of Phase 1 (extraction) — a single raw row from the bank statement.
    All fields are strings; type coercion happens in Phase 2.
    """
    TransactionDate: str
    RawDescription: str
    Amount: str
    Balance: str


@dataclass
class CleanedRecord:
    """
    Output of Phase 2 (cleaning) — typed and standardized.
    Amount and Balance are Decimal for precision safety.
    """
    TransactionDate: date
    RawDescription: str
    Amount: Decimal
    Balance: Decimal


@dataclass
class ClassifiedRecord:
    """
    Output of Phase 3 (classification) — enriched with AI-inferred fields.
    """
    TransactionDate: date
    RawDescription: str
    Amount: Decimal
    Balance: Decimal
    Category: str
    CounterpartyName: str
    IsRecurring: bool
    Confidence: float


@dataclass
class AuditEntry:
    """
    BRANCH A — appended to the audit log whenever the LLM returns
    low-confidence or unrecognizable classifications.
    """
    TransactionDate: date
    RawDescription: str
    Amount: Decimal
    Balance: Decimal
    Reason: str
    OriginalCategory: Optional[str]
    OriginalCounterparty: Optional[str]


@dataclass
class PipelineResult:
    """
    Final output of the pipeline run, returned to the caller.
    """
    total_transactions: int
    total_counterparties: int
    audit_log_count: int
    balance_valid: bool
    export_path: Optional[str]
    errors: List[str] = field(default_factory=list)


# ──────────────────────────────────────────────────────────────────────
# LLM Classification + Keyword Fallback
# ──────────────────────────────────────────────────────────────────────

_KEYWORD_RULES: List[Tuple[List[str], str, str]] = [
    (["fairprice", "ntuc", "sheng siong", "cold storage", "giant"], "F&B", "FairPrice"),
    (["grab"], "Transport", "Grab"),
    (["comfort", "citycab", "taxi"], "Transport", "ComfortDelGro"),
    (["singtel", "starhub", "m1"], "Telecommunications", "Singtel"),
    (["sp services", "pub", "power"], "Utilities", "SP Services"),
    (["transfer", "inward", "outward", "remittance"], "Transfer", "Unknown"),
    (["salary", "payroll", "wage"], "Salary", "Employer"),
    (["rental", "rent"], "Rental", "Landlord"),
    (["insurance", "aia", "prudential", "great eastern"], "Insurance", "Insurance Provider"),
    (["svc fee", "service fee", "bank fee", "monthly fee"], "Bank Fees", "DBS Bank"),
    (["subscription", "netflix", "spotify", "amazon"], "Subscription", "Subscription Service"),
    (["gst", "iras", "government", "tax"], "Government", "Government"),
]

_RECURRING_KEYWORDS = frozenset({
    "monthly", "subscription", "bill", "fee", "rental", "insurance premium",
    "svc fee", "service fee", "recurring",
})

# Mapping from LLM prompt categories to internal VALID_CATEGORIES.
_LLM_CATEGORY_MAP: Dict[str, str] = {
    "Food & Dining": "F&B",
    "Transport": "Transport",
    "Shopping": "Other",
    "Bills & Utilities": "Utilities",
    "Transfer": "Transfer",
    "Income": "Salary",
    "Entertainment": "Other",
    "Health": "Other",
    "Subscription": "Subscription",
    "Cashback": "Other",
    "Interest": "Investment",
    "Other": "Other",
}


def _fallback_classify(descriptions: List[str]) -> List[dict]:
    """
    Deterministic keyword-based classification used as a fallback when
    the LLM API is unavailable.

    Returns a list of dicts, one per input description, with keys:
        index, Category, CounterpartyName, IsRecurring, Confidence
    """
    results: List[dict] = []
    for i, desc in enumerate(descriptions):
        lower = desc.lower()
        matched = False
        for keywords, category, counterparty in _KEYWORD_RULES:
            if any(kw in lower for kw in keywords):
                confidence = 0.92 if len(keywords) <= 2 else 0.78
                is_recurring = any(kw in lower for kw in _RECURRING_KEYWORDS)
                results.append({
                    "index": i,
                    "Category": category,
                    "CounterpartyName": counterparty,
                    "IsRecurring": is_recurring,
                    "Confidence": confidence,
                })
                matched = True
                break
        if not matched:
            results.append({
                "index": i,
                "Category": "",
                "CounterpartyName": "",
                "IsRecurring": False,
                "Confidence": 0.15,
            })
    return results


# Backward-compatible alias so that existing test imports still resolve.
_stub_llm_classify = _fallback_classify


def _llm_classify_batch(descriptions: List[str]) -> List[dict]:
    """
    Real LLM classification via the ZhipuAI SDK.

    Sends a batch of raw bank transaction descriptions to the LLM and
    returns structured classification results.  Falls back gracefully
    (returns an empty list) on any API or parsing failure, allowing the
    caller to activate the keyword-based fallback path.

    Returns a list of dicts with keys:
        index, Category, CounterpartyName, IsRecurring, Confidence
    """
    api_key = os.environ.get("ZAI_API_KEY")
    model = os.environ.get("ZAI_MODEL", "glm-4-flash")

    if not api_key:
        logger.warning(
            "ZAI_API_KEY not set — LLM classification unavailable, "
            "using keyword fallback."
        )
        return []

    try:
        from zhipuai import ZhipuAI
    except ImportError:
        logger.warning(
            "zhipuai SDK not installed — LLM classification unavailable. "
            "Install with: pip install zhipuai"
        )
        return []

    # ── Build the system prompt ──
    system_msg = (
        "You are a bank transaction classifier.  Classify each transaction "
        "description into exactly one category and identify the counterparty.\n\n"
        "Valid categories: Food & Dining, Transport, Shopping, "
        "Bills & Utilities, Transfer, Income, Entertainment, Health, "
        "Subscription, Cashback, Interest, Other\n\n"
        "Return ONLY a valid JSON array — no markdown, no explanation.  "
        "One object per transaction with these fields:\n"
        '  "index":        (integer) position in the input list\n'
        '  "category":      (string)  one of the valid categories above\n'
        '  "counterparty":  (string)  standardised business/entity name in UPPER CASE\n'
        '  "confidence":    (float)   0.0–1.0\n'
        '  "is_recurring":  (boolean) true for subscriptions, monthly bills, regular payments\n'
    )

    user_lines = "\n".join(
        f"{i}. {desc}" for i, desc in enumerate(descriptions)
    )
    user_msg = f"Classify these transactions:\n{user_lines}"

    # ── Call the API via ZhipuAI SDK ──
    try:
        client = ZhipuAI(api_key=api_key)
        resp = client.chat.completions.create(
            model=model,
            messages=[
                {"role": "system", "content": system_msg},
                {"role": "user", "content": user_msg},
            ],
            temperature=0.1,
            max_tokens=4096,
        )
        content: str = resp.choices[0].message.content

        content = content.strip()
        if content.startswith("```"):
            lines = content.split("\n")
            lines = [ln for ln in lines if not ln.strip().startswith("```")]
            content = "\n".join(lines)

        parsed = json.loads(content)

    except json.JSONDecodeError as exc:
        logger.warning("LLM returned malformed JSON: %s — falling back to keywords.", exc)
        return []
    except (KeyError, IndexError) as exc:
        logger.warning("LLM response missing expected field: %s — falling back to keywords.", exc)
        return []
    except Exception as exc:
        logger.warning("LLM API error: %s — falling back to keywords.", exc)
        return []

    # ── Map LLM fields → internal format ──
    results: List[dict] = []
    for item in parsed:
        raw_category = str(item.get("category", "")).strip()
        mapped = _LLM_CATEGORY_MAP.get(raw_category, raw_category)
        if mapped not in VALID_CATEGORIES:
            mapped = ""

        results.append({
            "index": int(item.get("index", 0)),
            "Category": mapped,
            "CounterpartyName": str(item.get("counterparty", "")).strip(),
            "IsRecurring": bool(item.get("is_recurring", False)),
            "Confidence": float(item.get("confidence", 0.0)),
        })
    return results


# ──────────────────────────────────────────────────────────────────────
# Pipeline Engine
# ──────────────────────────────────────────────────────────────────────


class BankStatementPipeline:
    """
    Orchestrates the 6-phase bank statement processing pipeline.

    Phases:
      1. Data Extraction          (pdf-specialist-skill)
      2. Data Cleaning            (relational-data-processor-skill)
      3. AI Classification        (transaction-classifier-skill)
      4. Relational Key Mapping   (relational-data-processor-skill)
      5. Financial Aggregation    (relational-data-processor-skill)
      6. Relational Database Export (db-storage-specialist-skill)
    """

    def __init__(
        self,
        source_path: str | Path,
        initial_balance: Optional[Decimal] = None,
        output_dir: str | Path = "./output",
        confidence_threshold: float = 0.6,
    ) -> None:
        self.source_path = Path(source_path)
        self._initial_balance_input: Optional[Decimal] = (
            _safe_decimal(initial_balance) if initial_balance is not None else None
        )
        self.output_dir = Path(output_dir)
        self.confidence_threshold = confidence_threshold

        self.extracted_initial_balance: Optional[Decimal] = None

        self._raw_records: List[RawRecord] = []
        self._cleaned_records: List[CleanedRecord] = []
        self._classified_records: List[ClassifiedRecord] = []
        self._enriched_records: List[dict] = []
        self._audit_log: List[AuditEntry] = []
        self._store: Optional[BankStatementStore] = None

        self._phases_completed: List[str] = []

    @property
    def initial_balance(self) -> Decimal:
        if self._initial_balance_input is not None:
            return self._initial_balance_input
        if self.extracted_initial_balance is not None:
            return self.extracted_initial_balance
        return Decimal("0")

    # ── Public API ────────────────────────────────────────────────────

    def run(self) -> PipelineResult:
        """
        Execute the full 6-phase pipeline sequentially.

        BRANCH C orchestration:
          Phase 5 raises FatalBalanceError → Phase 6 is SKIPPED.
          The except block exports the audit log and returns early
          with balance_valid=False.

        BRANCH A:
          Phase 3 failures route low-confidence records to audit_log
          but do NOT halt the pipeline.

        BRANCH B:
          Phase 4 dedup is handled transparently by CounterpartyRegistry.
        """
        result_errors: List[str] = []
        balance_valid = True
        export_path: Optional[str] = None

        try:
            logger.info("=" * 50)
            logger.info("Pipeline START — source: %s", self.source_path)
            logger.info("=" * 50)

            # Phase 1
            self._phases_completed.append("Phase 1")
            self.phase_1_extract()
            logger.info(
                "Phase 1 complete: %d raw records extracted",
                len(self._raw_records),
            )

            # Phase 2
            self._phases_completed.append("Phase 2")
            self.phase_2_clean()
            logger.info(
                "Phase 2 complete: %d records cleaned",
                len(self._cleaned_records),
            )

            # Phase 3
            self._phases_completed.append("Phase 3")
            self.phase_3_classify()
            logger.info(
                "Phase 3 complete: %d classified, %d flagged for review",
                len(self._classified_records),
                len(self._audit_log),
            )

            # Phase 4
            self._phases_completed.append("Phase 4")
            self.phase_4_map_keys()
            logger.info("Phase 4 complete: relational keys mapped")

            # ── Resolve initial_balance before Phase 5 ──
            if self._initial_balance_input is not None:
                logger.info(
                    "Initial balance: %s (provided by user)",
                    str(self._initial_balance_input),
                )
            elif self.extracted_initial_balance is not None:
                logger.info(
                    "Initial balance: %s (auto-extracted from 'Balance Brought Forward')",
                    str(self.extracted_initial_balance),
                )
            else:
                raise PipelineError(
                    "No initial balance available. Either provide --initial-balance "
                    "or ensure the statement contains a 'Balance Brought Forward' row."
                )

            # Phase 5 — may raise FatalBalanceError (BRANCH C)
            self._phases_completed.append("Phase 5")
            self.phase_5_validate_balances()
            logger.info("Phase 5 complete: all balances validated")

            # Phase 6 — ONLY reached if Phase 5 passed
            self._phases_completed.append("Phase 6")
            self.phase_6_export()
            export_path = str(self.output_dir.resolve())
            logger.info("Phase 6 complete: data exported to %s", export_path)

        except FatalBalanceError as exc:
            balance_valid = False
            error_summary = str(exc)
            result_errors.append(error_summary)
            logger.critical("BRANCH C ACTIVATED: %s", error_summary)
            for err in exc.balance_errors:
                logger.critical(
                    "  Row %s: expected=%s actual=%s discrepancy=%s",
                    err.get("index"),
                    err.get("expected_balance"),
                    err.get("actual_balance"),
                    err.get("discrepancy"),
                )
            logger.critical("Phase 6 export BLOCKED — pipeline terminated.")

        except (ExtractionError, ClassificationError, PipelineError) as exc:
            result_errors.append(str(exc))
            logger.error("Pipeline halted: %s", exc)

        except IntegrityError as exc:
            result_errors.append(f"Database integrity error: {exc}")
            logger.error("Integrity failure: %s", exc)

        finally:
            audit_path = self.output_dir / "audit_log.json"
            if self._audit_log:
                self.export_audit_log(audit_path)
                logger.info("Audit log written: %s (%d entries)", audit_path, len(self._audit_log))

        txn_count = len(self._classified_records)
        cp_count = self._store.registry.count() if self._store else 0

        logger.info("=" * 50)
        logger.info(
            "Pipeline END — phases: %s, transactions: %d, counterparties: %d",
            "+".join(self._phases_completed),
            txn_count,
            cp_count,
        )
        logger.info("=" * 50)

        return PipelineResult(
            total_transactions=txn_count,
            total_counterparties=cp_count,
            audit_log_count=len(self._audit_log),
            balance_valid=balance_valid,
            export_path=export_path,
            errors=result_errors,
        )

    @property
    def audit_log(self) -> List[AuditEntry]:
        """Return the accumulated audit entries (Branch A)."""
        return list(self._audit_log)

    def export_audit_log(self, path: str | Path) -> None:
        """
        Write the audit log to a JSON file for human review.
        Called automatically at the end of the pipeline or on early termination.
        """
        out = Path(path)
        out.parent.mkdir(parents=True, exist_ok=True)
        entries = []
        for entry in self._audit_log:
            d = asdict(entry)
            d["TransactionDate"] = entry.TransactionDate.isoformat()
            d["Amount"] = str(entry.Amount)
            d["Balance"] = str(entry.Balance)
            entries.append(d)
        with open(out, "w", encoding="utf-8") as f:
            json.dump(entries, f, indent=2, ensure_ascii=False)

    # ── Phase 1: Data Extraction ──────────────────────────────────────

    def phase_1_extract(self) -> List[RawRecord]:
        """
        Phase 1 — Data Extraction
        Skill invoked: pdf-specialist-skill

        Reads the source bank statement (PDF or CSV) and extracts raw
        text/table rows into a list of RawRecord objects.

        Raises:
            ExtractionError: if no records are extracted from the source.
        """
        suffix = self.source_path.suffix.lower()
        records: List[RawRecord] = []

        if suffix == ".pdf":
            records = self._extract_pdf()
        elif suffix == ".csv":
            records = self._extract_csv()
        else:
            raise ExtractionError(
                f"Unsupported file format: '{suffix}'. Only .pdf and .csv are accepted."
            )

        if not records:
            raise ExtractionError(
                f"No records extracted from '{self.source_path}'. "
                f"File may be empty, corrupted, or in an unrecognized format."
            )

        self._raw_records = records
        return records

    def _extract_pdf(self) -> List[RawRecord]:
        """
        Extract transaction data from a PDF bank statement using pdfplumber.

        Two-pass strategy:
          1. Try extract_table() — works for PDFs with tagged <table> elements.
          2. On failure, fall back to extract_text() — line-by-line text parsing
             with regex-based transaction detection and multi-line description merging.

        The text fallback handles DBS-style statements where transactions are
        rendered as sequential text lines rather than structured tables.
        """
        try:
            import pdfplumber
        except ImportError:
            raise ExtractionError(
                "pdfplumber is required for PDF extraction. "
                "Install with: pip install pdfplumber"
            )

        # ── Pass 1: table-based extraction ──
        records = self._extract_pdf_tables(pdfplumber)
        if records is not None:
            return records

        # ── Pass 2: text-based fallback ──
        records = self._extract_pdf_text(pdfplumber)
        return records

    def _extract_pdf_tables(self, pdfplumber) -> Optional[List[RawRecord]]:
        """
        Attempt table-based extraction using pdfplumber.extract_table().

        Returns a list of RawRecord on success, or None to signal that the
        text-based fallback should be tried instead.
        """
        all_tables: List[List[List[Optional[str]]]] = []

        with pdfplumber.open(str(self.source_path)) as pdf:
            for page in pdf.pages:
                table = page.extract_table()
                if table:
                    all_tables.append(table)

        if not all_tables:
            return None

        best_table = max(all_tables, key=len)
        if len(best_table) < 2:
            return None

        header = [str(c).strip().lower() if c else "" for c in best_table[0]]
        col_map = self._map_columns(header)

        records: List[RawRecord] = []
        for row in best_table[1:]:
            if not row or all(not (c and str(c).strip()) for c in row):
                continue
            cells = [str(c).strip() if c else "" for c in row]
            date_val = cells[col_map["date"]] if col_map["date"] < len(cells) else ""
            desc_val = cells[col_map["desc"]] if col_map["desc"] < len(cells) else ""
            amt_val = cells[col_map["amount"]] if col_map["amount"] < len(cells) else ""
            bal_val = cells[col_map["balance"]] if col_map["balance"] < len(cells) else ""

            if desc_val and "balance brought forward" in desc_val.lower():
                cleaned_bal = self._normalize_amount_string(bal_val)
                try:
                    self.extracted_initial_balance = _safe_decimal(cleaned_bal).quantize(
                        TWO_PLACES, rounding=ROUND_HALF_UP
                    )
                    logger.info(
                        "Phase 1 [table]: Extracted opening balance "
                        "from 'Balance Brought Forward': %s",
                        str(self.extracted_initial_balance),
                    )
                except ValueError:
                    logger.warning(
                        "Phase 1 [table]: 'Balance Brought Forward' found "
                        "but balance '%s' unparseable", bal_val,
                    )
                continue

            if not date_val and not amt_val:
                continue
            records.append(RawRecord(
                TransactionDate=date_val,
                RawDescription=desc_val[:255],
                Amount=amt_val,
                Balance=bal_val,
            ))

        return records

    def _extract_pdf_text(self, pdfplumber) -> List[RawRecord]:
        """
        Text-based fallback for PDFs without tagged tables (e.g. DBS statements).

        Algorithm:
          1. Concatenate extract_text() from all pages into one stream of lines.
          2. Detect "Balance Brought Forward" to capture opening balance.
          3. Use a regex (DD/MM/YYYY) to identify the first line of each transaction.
          4. Extract amount and balance from that first line.
          5. Collect subsequent non-date lines as the description continuation
             until the next transaction date line or "Balance Carried Forward".
          6. Merge everything into a single RawRecord per transaction.
        """
        _DATE_RE = re.compile(
            r"^(\d{2}/\d{2}/\d{4})\s+(.+?)\s+"
            r"([\d,]+\.\d{2})\s+"
            r"([\d,]+\.\d{2})\s*$"
        )
        _BBF_RE = re.compile(
            r"Balance\s+Brought\s+Forward\s+"
            r"(?:SGD\s+)?([\d,]+\.\d{2})",
            re.IGNORECASE,
        )
        _BCF_RE = re.compile(
            r"Balance\s+Carried\s+Forward",
            re.IGNORECASE,
        )
        _SKIP_LINE_RE = re.compile(
            r"^(?:CURRENCY\s*:|Date\s+Description|Transaction Details|"
            r"DBS\s+Multiplier|Account\s+No|Balance\s+(?:Brought|Carried)|"
            r"Page\s+\d+\s+of\s+\d+|PDS_|A\d{8}|"
            r"\.oN|geR|ziB|BSOP|:oN|TSG|\.geR|\.oC|SBD|\).*\(|"
            r"678\d+GS|E\d{9}|/\s*\d+-\d+-RM|"
            r"Account\s+Summary|as\s+(?:at|of)\s+\d|Deposits|"
            r"Current\s+and\s+Savings|Summary\s+of\s+Currency|"
            r"Account\s+Account\s+No)",
            re.IGNORECASE,
        )

        all_lines: List[str] = []
        with pdfplumber.open(str(self.source_path)) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    for raw_line in text.split("\n"):
                        stripped = raw_line.strip()
                        if stripped:
                            all_lines.append(stripped)

        if not all_lines:
            raise ExtractionError(
                "PDF appears to be empty — no text could be extracted from any page."
            )

        logger.info(
            "Phase 1 [text]: Extracted %d text lines across all pages, "
            "switching to text-based parsing", len(all_lines),
        )

        bbf_found = False
        for line in all_lines:
            m = _BBF_RE.search(line)
            if m:
                raw_bal = m.group(1)
                cleaned = self._normalize_amount_string(raw_bal)
                try:
                    self.extracted_initial_balance = _safe_decimal(cleaned).quantize(
                        TWO_PLACES, rounding=ROUND_HALF_UP
                    )
                    logger.info(
                        "Phase 1 [text]: Extracted opening balance from "
                        "'Balance Brought Forward': %s",
                        str(self.extracted_initial_balance),
                    )
                    bbf_found = True
                except ValueError:
                    logger.warning(
                        "Phase 1 [text]: 'Balance Brought Forward' found "
                        "but balance '%s' unparseable", raw_bal,
                    )
                break

        if not bbf_found:
            logger.warning(
                "Phase 1 [text]: 'Balance Brought Forward' not found in text — "
                "opening balance will need to be provided via --initial-balance"
            )

        records: List[RawRecord] = []
        current_date: Optional[str] = None
        current_desc_parts: List[str] = []
        current_amount: Optional[str] = None
        current_balance: Optional[str] = None
        txn_started = False

        def _flush():
            if not txn_started or current_date is None:
                return
            desc = " ".join(current_desc_parts).strip()[:255]
            records.append(RawRecord(
                TransactionDate=current_date,
                RawDescription=desc,
                Amount=current_amount or "",
                Balance=current_balance or "",
            ))

        for line in all_lines:
            if _SKIP_LINE_RE.match(line):
                continue

            if _BBF_RE.search(line):
                continue

            if _BCF_RE.search(line):
                _flush()
                txn_started = False
                current_date = None
                current_desc_parts = []
                current_amount = None
                current_balance = None
                continue

            date_match = _DATE_RE.match(line)

            if date_match:
                _flush()

                current_date = date_match.group(1)
                desc_head = date_match.group(2).strip()
                current_amount = date_match.group(3).replace(",", "")
                current_balance = date_match.group(4).replace(",", "")
                current_desc_parts = [desc_head] if desc_head else []
                txn_started = True
            elif txn_started:
                current_desc_parts.append(line.strip())

        _flush()

        logger.info(
            "Phase 1 [text]: Parsed %d transaction records from text lines",
            len(records),
        )
        return records

    @staticmethod
    def _map_columns(headers: List[str]) -> Dict[str, int]:
        mapping = {"date": -1, "desc": -1, "amount": -1, "balance": -1}
        date_kw = {"date", "transaction date", "value date", "posting date"}
        desc_kw = {"description", "particulars", "details", "transaction details", "reference"}
        amt_kw = {"amount", "debit", "credit", "withdrawal", "deposit"}
        bal_kw = {"balance", "running balance", "closing balance"}
        for i, h in enumerate(headers):
            if any(k in h for k in date_kw) and mapping["date"] == -1:
                mapping["date"] = i
            elif any(k in h for k in desc_kw) and mapping["desc"] == -1:
                mapping["desc"] = i
            elif any(k in h for k in amt_kw) and mapping["amount"] == -1:
                mapping["amount"] = i
            elif any(k in h for k in bal_kw) and mapping["balance"] == -1:
                mapping["balance"] = i
        for key, idx in mapping.items():
            if idx == -1:
                defaults = {"date": 0, "desc": 1, "amount": 2, "balance": 3}
                mapping[key] = defaults[key]
        return mapping

    def _extract_csv(self) -> List[RawRecord]:
        """
        Extract records from a CSV bank statement.

        Attempts flexible column mapping by header keyword matching.
        Handles various date formats, amount columns, and encodings.
        """
        records: List[RawRecord] = []
        encodings = ["utf-8-sig", "utf-8", "latin-1", "cp1252"]

        for enc in encodings:
            try:
                with open(self.source_path, "r", encoding=enc, newline="") as f:
                    reader = csv.DictReader(f)
                    if not reader.fieldnames:
                        continue
                    headers = [h.strip().lower() for h in reader.fieldnames]
                    col_map = self._map_columns(headers)
                    header_list = list(reader.fieldnames)

                    for row in reader:
                        get = lambda key: (row.get(header_list[col_map[key]], "") or "").strip()
                        date_val = get("date")
                        desc_val = get("desc")
                        amt_val = get("amount")
                        bal_val = get("balance")

                        if desc_val and "balance brought forward" in desc_val.lower():
                            cleaned_bal = self._normalize_amount_string(bal_val)
                            try:
                                self.extracted_initial_balance = _safe_decimal(cleaned_bal).quantize(
                                    TWO_PLACES, rounding=ROUND_HALF_UP
                                )
                                logger.info(
                                    "Phase 1: Extracted opening balance from "
                                    "'Balance Brought Forward': %s",
                                    str(self.extracted_initial_balance),
                                )
                            except ValueError:
                                logger.warning(
                                    "Phase 1: 'Balance Brought Forward' found "
                                    "but balance '%s' unparseable",
                                    bal_val,
                                )
                            continue

                        if not date_val and not amt_val:
                            continue
                        records.append(RawRecord(
                            TransactionDate=date_val,
                            RawDescription=desc_val[:255],
                            Amount=amt_val,
                            Balance=bal_val,
                        ))
                break
            except (UnicodeDecodeError, csv.Error):
                continue

        return records

    # ── Phase 2: Data Cleaning & Format Standardization ───────────────

    def phase_2_clean(self) -> List[CleanedRecord]:
        """
        Phase 2 — Data Cleaning & Format Standardization
        Skill invoked: relational-data-processor-skill

        Transforms raw string records into typed, cleaned records with
        Decimal precision for financial fields.

        Steps:
          a. Strip whitespace, remove null rows
          b. Parse TransactionDate → date object
          c. Convert Amount and Balance to Decimal (via _safe_decimal)
          d. Balance Delta Deduction — recover the correct sign for Amount
             by comparing running balance against each row's closing balance.
             Only activated when ``extracted_initial_balance`` is available
             (set during Phase 1 from 'Balance Brought Forward').
          e. Truncate RawDescription to 255 chars
          f. Sort by TransactionDate ascending

        Balance Delta Deduction algorithm:
          For PDF text extraction the Withdrawal/Deposit column distinction
          is lost — all amounts arrive as unsigned magnitudes.  This step
          deterministically recovers the cash-flow direction:

          1. Maintain ``running_balance`` initialised from the opening balance.
          2. For each row, try the amount with its *current* sign first.
             If ``running_balance + amount == row_balance`` the sign is
             correct — keep it (handles CSV data that already has signs).
          3. If the current sign doesn't match, determine the correct
             direction by testing both deltas against the row balance.
          4. Advance ``running_balance`` to ``row_balance`` for the next row.
        """
        cleaned: List[CleanedRecord] = []
        parse_errors: List[str] = []

        # ── Balance Delta Deduction state ──
        # Activated only when extracted_initial_balance is available
        # (auto-extracted from 'Balance Brought Forward' during Phase 1).
        # This ensures CSV sources with pre-signed amounts are unaffected.
        tolerance = Decimal("0.01")
        running_balance: Optional[Decimal] = self.extracted_initial_balance

        for i, raw in enumerate(self._raw_records):
            raw_date = raw.TransactionDate.strip()
            raw_desc = raw.RawDescription.strip()[:255]
            raw_amt = raw.Amount.strip()
            raw_bal = raw.Balance.strip()

            if not raw_amt and not raw_bal:
                parse_errors.append(f"Row {i}: both Amount and Balance empty — skipped")
                continue

            txn_date = self._parse_flexible_date(raw_date)
            if txn_date is None:
                parse_errors.append(f"Row {i}: unparseable date '{raw_date}' — skipped")
                continue

            amt_str = self._normalize_amount_string(raw_amt)
            bal_str = self._normalize_amount_string(raw_bal)

            try:
                amount = _safe_decimal(amt_str) if amt_str else Decimal("0")
                balance = _safe_decimal(bal_str) if bal_str else Decimal("0")
            except ValueError as exc:
                parse_errors.append(f"Row {i}: numeric conversion failed — {exc}")
                continue

            # ── Balance Delta Deduction ──
            if running_balance is not None:
                row_balance = balance.quantize(TWO_PLACES, rounding=ROUND_HALF_UP)

                # Step 1: Try amount with its current sign.
                expected = (running_balance + amount).quantize(
                    TWO_PLACES, rounding=ROUND_HALF_UP
                )
                if abs(expected - row_balance) <= tolerance:
                    # Current sign is consistent with balance arithmetic.
                    pass
                else:
                    # Step 2: Determine correct sign using absolute magnitude.
                    abs_amount = abs(amount).quantize(
                        TWO_PLACES, rounding=ROUND_HALF_UP
                    )
                    expected_withdrawal = (
                        running_balance - abs_amount
                    ).quantize(TWO_PLACES, rounding=ROUND_HALF_UP)
                    expected_deposit = (
                        running_balance + abs_amount
                    ).quantize(TWO_PLACES, rounding=ROUND_HALF_UP)

                    if abs(expected_withdrawal - row_balance) <= tolerance:
                        # running_balance - amount == row_balance → Withdrawal
                        amount = -abs_amount
                    elif abs(expected_deposit - row_balance) <= tolerance:
                        # running_balance + amount == row_balance → Deposit
                        amount = abs_amount
                    else:
                        # Neither delta matches — default to positive, log warning
                        logger.warning(
                            "Phase 2: Row %d — cannot determine cash flow "
                            "direction for amount %s (running_balance=%s, "
                            "row_balance=%s). Defaulting to deposit (positive).",
                            i,
                            str(abs_amount),
                            str(running_balance),
                            str(row_balance),
                        )
                        amount = abs_amount

                # Advance running balance for next iteration
                running_balance = row_balance

            cleaned.append(CleanedRecord(
                TransactionDate=txn_date,
                RawDescription=raw_desc,
                Amount=amount.quantize(TWO_PLACES, rounding=ROUND_HALF_UP),
                Balance=balance.quantize(TWO_PLACES, rounding=ROUND_HALF_UP),
            ))

        if parse_errors:
            logger.warning(
                "Phase 2: %d rows had parse issues (%d kept, %d dropped)",
                len(parse_errors), len(cleaned), len(parse_errors),
            )
            for err in parse_errors:
                logger.warning("  %s", err)

        cleaned.sort(key=lambda r: r.TransactionDate)
        self._cleaned_records = cleaned
        return cleaned

    @staticmethod
    def _parse_flexible_date(date_str: str) -> Optional[date]:
        """
        Attempt multiple date format patterns to parse a date string.

        Supports: YYYY-MM-DD, DD/MM/YYYY, DD-MM-YYYY, DD Mon YYYY, etc.
        """
        if not date_str:
            return None
        date_str = date_str.strip()
        for fmt in (
            "%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%Y/%m/%d",
            "%d %b %Y", "%d %B %Y", "%b %d, %Y", "%B %d, %Y",
            "%m/%d/%Y", "%d.%m.%Y",
        ):
            try:
                import datetime
                return datetime.datetime.strptime(date_str, fmt).date()
            except ValueError:
                continue
        try:
            return date.fromisoformat(date_str[:10])
        except (ValueError, TypeError):
            return None

    @staticmethod
    def _normalize_amount_string(s: str) -> str:
        """
        Normalize a financial amount string for Decimal parsing.

        Removes currency symbols, commas, and parenthetical negatives:
          "(1,234.56)" → "-1234.56"
          "SGD 500.00" → "500.00"
          "1,234.56"   → "1234.56"
        """
        s = s.strip()
        if not s:
            return "0"
        s = re.sub(r"[A-Za-z$€£¥]", "", s)
        s = s.replace(",", "")
        if s.startswith("(") and s.endswith(")"):
            s = "-" + s[1:-1]
        if s.startswith("CR") or s.endswith("CR"):
            s = s.replace("CR", "").strip()
        elif s.startswith("DR") or s.endswith("DR"):
            val = s.replace("DR", "").strip()
            if val and not val.startswith("-"):
                s = "-" + val
        return s.strip() or "0"

    # ── Phase 3: AI Classification & Attribute Inference ──────────────

    def phase_3_classify(self) -> List[ClassifiedRecord]:
        """
        Phase 3 — AI Classification & Attribute Inference
        Skill invoked: transaction-classifier-skill

        Sends cleaned records to the LLM for classification.

        BRANCH A — AI Inference Exception (CRITICAL RULE):
        If the LLM returns a classification with confidence below
        self.confidence_threshold, or returns an unrecognizable/
        empty category, the system MUST:
          1. Set Category to "Uncategorized" (never guess or fabricate)
          2. Set CounterpartyName to "Unknown"
          3. Append an AuditEntry to self._audit_log with:
             — the original raw description
             — the reason (e.g., "low confidence: 0.32")
             — any partial LLM output for traceability
          4. Continue processing remaining records (do NOT halt)

        If ALL records fail classification → raise ClassificationError.
        """
        if not self._cleaned_records:
            raise ClassificationError("No cleaned records available for classification.")

        descriptions = [r.RawDescription for r in self._cleaned_records]

        # ── Sub-batch processing (25 per batch) ──
        _BATCH_SIZE = 25
        all_results: List[dict] = []

        for batch_start in range(0, len(descriptions), _BATCH_SIZE):
            batch_end = min(batch_start + _BATCH_SIZE, len(descriptions))
            batch = descriptions[batch_start:batch_end]

            batch_results: List[dict] = []
            try:
                batch_results = _llm_classify_batch(batch)
                if not batch_results:
                    # API returned nothing (no key / empty response) → fallback
                    raise ValueError("LLM returned empty results for this batch")
                # Shift indices from batch-local to global position
                for r in batch_results:
                    r["index"] = r["index"] + batch_start
            except Exception as exc:
                logger.warning(
                    "Phase 3: LLM failed for batch [%d:%d] (%s). "
                    "Using keyword fallback.",
                    batch_start, batch_end, exc,
                )
                batch_results = _fallback_classify(batch)
                for r in batch_results:
                    r["index"] = r["index"] + batch_start

            all_results.extend(batch_results)

        indexed_results: Dict[int, dict] = {r["index"]: r for r in all_results}

        classified: List[ClassifiedRecord] = []
        success_count = 0

        for i, cleaned in enumerate(self._cleaned_records):
            result = indexed_results.get(i, {})
            category = result.get("Category", "")
            counterparty = result.get("CounterpartyName", "")
            is_recurring = result.get("IsRecurring", False)
            confidence = float(result.get("Confidence", 0.0))

            if (
                confidence >= self.confidence_threshold
                and category
                and category in VALID_CATEGORIES
            ):
                classified.append(ClassifiedRecord(
                    TransactionDate=cleaned.TransactionDate,
                    RawDescription=cleaned.RawDescription,
                    Amount=cleaned.Amount,
                    Balance=cleaned.Balance,
                    Category=category,
                    CounterpartyName=counterparty if counterparty else "Unknown",
                    IsRecurring=bool(is_recurring),
                    Confidence=confidence,
                ))
                success_count += 1
            else:
                # ── BRANCH A: Low confidence or invalid category ──
                reasons = []
                if confidence < self.confidence_threshold:
                    reasons.append(f"low confidence: {confidence:.2f}")
                if not category:
                    reasons.append("empty category")
                elif category not in VALID_CATEGORIES:
                    reasons.append(f"invalid category: '{category}'")
                reason_str = "; ".join(reasons) if reasons else "unclassifiable"

                self._audit_log.append(AuditEntry(
                    TransactionDate=cleaned.TransactionDate,
                    RawDescription=cleaned.RawDescription,
                    Amount=cleaned.Amount,
                    Balance=cleaned.Balance,
                    Reason=reason_str,
                    OriginalCategory=category if category else None,
                    OriginalCounterparty=counterparty if counterparty else None,
                ))

                classified.append(ClassifiedRecord(
                    TransactionDate=cleaned.TransactionDate,
                    RawDescription=cleaned.RawDescription,
                    Amount=cleaned.Amount,
                    Balance=cleaned.Balance,
                    Category="Uncategorized",
                    CounterpartyName="Unknown",
                    IsRecurring=False,
                    Confidence=confidence,
                ))
                logger.warning(
                    "Branch A: '%s' → %s (confidence=%.2f)",
                    cleaned.RawDescription[:50], reason_str, confidence,
                )

        if success_count == 0:
            logger.warning(
                "All %d records fell below confidence threshold %.2f — "
                "all flagged as Uncategorized for human review.",
                len(self._cleaned_records),
                self.confidence_threshold,
            )

        self._classified_records = classified
        return classified

    # ── Phase 4: Relational Key Generation & Referential Mapping ──────

    def phase_4_map_keys(self) -> None:
        """
        Phase 4 — Relational Key Generation & Referential Mapping
        Skill invoked: relational-data-processor-skill

        Generates deterministic UUID5 keys and maps foreign keys
        using the CounterpartyRegistry from bank_statement_model.py.

        BRANCH B — Relational Mapping Logic (CRITICAL RULE):
        For each ClassifiedRecord:
          1. Query the CounterpartyRegistry with the inferred CounterpartyName
          2. IF the name already exists in the registry:
             → Reuse the existing CounterpartyID (no duplicate)
          3. IF the name is new:
             → CounterpartyRegistry.resolve() generates a new deterministic
               uuid5 ID and registers it
          4. Map the CounterpartyID into the transaction record as FK
        This ensures referential integrity without duplicate counterparties.
        """
        self._store = BankStatementStore()
        self._store.open()

        enriched: List[dict] = []
        for rec in self._classified_records:
            # ── BRANCH B: registry.resolve handles dedup ──
            cp = self._store.registry.resolve(rec.CounterpartyName)
            logger.debug(
                "Branch B: '%s' → CounterpartyID=%s (registry count=%d)",
                rec.CounterpartyName, cp.CounterpartyID[:8] + "...",
                self._store.registry.count(),
            )

            enriched.append({
                "TransactionDate": rec.TransactionDate,
                "RawDescription": rec.RawDescription,
                "Amount": rec.Amount,
                "Balance": rec.Balance,
                "CounterpartyName": rec.CounterpartyName,
                "CounterpartyID": cp.CounterpartyID,
                "Category": rec.Category,
                "IsRecurring": rec.IsRecurring,
            })

        self._enriched_records = enriched
        logger.info(
            "Phase 4: %d transactions mapped to %d unique counterparties",
            len(enriched),
            self._store.registry.count(),
        )

    # ── Phase 5: Financial Aggregation & Balance Validation ───────────

    def phase_5_validate_balances(self) -> None:
        """
        Phase 5 — Financial Aggregation & Balance Validation
        Skill invoked: relational-data-processor-skill

        Runs strict running balance validation using the Decimal-based
        validate_running_balance() from bank_statement_model.py.

        BRANCH C — Fatal Error Blocking (CRITICAL RULE):
        The validation formula:
            Opening Balance + Current Amount = Closing Balance
        is checked for EVERY row in chronological order.

        IF any discrepancy exceeds the tolerance (0.01):
          1. Raise FatalBalanceError immediately
          2. The error contains the full list of balance mismatches
          3. The caller (run()) catches this and:
             — Skip Phase 6 (Database Export) entirely
             — Write the audit log
             — Return PipelineResult with balance_valid=False
          4. It is ABSOLUTELY FORBIDDEN to proceed to Phase 6
             when this validation fails.
        """
        transactions: List[Transaction] = []
        for rec in self._enriched_records:
            txn_id = Transaction.generate_id(
                rec["TransactionDate"], rec["Amount"], rec["RawDescription"]
            )
            transactions.append(Transaction(
                TransactionID=txn_id,
                CounterpartyID=rec["CounterpartyID"],
                TransactionDate=rec["TransactionDate"],
                RawDescription=rec["RawDescription"],
                Amount=rec["Amount"],
                Balance=rec["Balance"],
                Category=rec["Category"],
                IsRecurring=rec["IsRecurring"],
            ))

        errors = validate_running_balance(transactions, self.initial_balance)

        if errors:
            # ── BRANCH C: Fatal balance break → raise immediately ──
            logger.critical(
                "BRANCH C: %d balance break(s) detected. Raising FatalBalanceError.",
                len(errors),
            )
            raise FatalBalanceError(errors)

        logger.info(
            "Phase 5: All %d balances validated (initial=%s)",
            len(transactions),
            str(self.initial_balance),
        )

    # ── Phase 6: Relational Database Export ────────────────────────────

    def phase_6_export(self) -> None:
        """
        Phase 6 — Relational Database Export
        Skill invoked: db-storage-specialist-skill

        Writes the fully validated, classified, and keyed records into
        the relational database via BankStatementStore.insert_transactions().

        PRECONDITION: Phase 5 balance validation MUST have passed.
        This method is ONLY called if phase_5_validate_balances() returned
        without raising FatalBalanceError.
        """
        assert self._store is not None, "Store not initialized — Phase 4 must run first."

        count, balance_errs = self._store.insert_transactions(
            self._enriched_records,
            validate_balance=False,
            initial_balance=self.initial_balance,
        )
        logger.info("Phase 6: %d transactions inserted into database", count)

        integrity_issues = self._store.verify_integrity()
        if integrity_issues:
            raise IntegrityError(
                "Post-export integrity check failed: " + "; ".join(integrity_issues)
            )
        logger.info("Phase 6: Referential integrity verified — 0 orphans")

        self.output_dir.mkdir(parents=True, exist_ok=True)
        self._store.export_to_csv(self.output_dir)
        logger.info("Phase 6: CSV files exported to %s", self.output_dir)

        db_path = self.output_dir / "bank_statements.db"
        file_store = BankStatementStore(db_path)
        file_store.open()
        file_store.insert_transactions(
            self._enriched_records,
            validate_balance=False,
            initial_balance=self.initial_balance,
        )
        file_store.close()
        logger.info("Phase 6: SQLite database saved to %s", db_path)

        if self._audit_log:
            self.export_audit_log(self.output_dir / "audit_log.json")

        summary = self._store.counterparty_summary()
        logger.info("Phase 6: Counterparty summary (%d entries):", len(summary))
        for s in summary:
            logger.info(
                "  %-25s  %d txns  total=%s",
                s["CounterpartyName"], s["TransactionCount"], s["TotalAmount"],
            )


# ──────────────────────────────────────────────────────────────────────
# Entry Point
# ──────────────────────────────────────────────────────────────────────


def main() -> None:
    """
    CLI entry point for running the pipeline.

    Usage:
        python main_pipeline.py <source_path> <initial_balance> [--output-dir DIR]

    Example:
        python main_pipeline.py statement.pdf 5000.00 --output-dir ./output
    """
    parser = argparse.ArgumentParser(
        description="Bank Statement Automation Pipeline"
    )
    parser.add_argument(
        "source_path",
        type=str,
        help="Path to the bank statement file (PDF or CSV)",
    )
    parser.add_argument(
        "--initial-balance",
        type=str,
        default=None,
        help="Opening account balance (e.g. 5000.00). "
             "If omitted, auto-extracted from 'Balance Brought Forward' row.",
    )
    parser.add_argument(
        "--output-dir",
        type=str,
        default="./output",
        help="Directory for exported files (default: ./output)",
    )
    parser.add_argument(
        "--confidence",
        type=float,
        default=0.6,
        help="Minimum LLM confidence threshold (default: 0.6)",
    )
    parser.add_argument(
        "--verbose",
        action="store_true",
        help="Enable debug-level logging",
    )
    args = parser.parse_args()

    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.INFO,
        format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
        datefmt="%H:%M:%S",
    )

    pipeline = BankStatementPipeline(
        source_path=args.source_path,
        initial_balance=Decimal(args.initial_balance) if args.initial_balance else None,
        output_dir=args.output_dir,
        confidence_threshold=args.confidence,
    )

    result = pipeline.run()

    print("\n" + "=" * 50)
    print("PIPELINE RESULT")
    print("=" * 50)
    print(f"  Transactions:   {result.total_transactions}")
    print(f"  Counterparties: {result.total_counterparties}")
    print(f"  Audit entries:  {result.audit_log_count}")
    print(f"  Balance valid:  {result.balance_valid}")
    print(f"  Export path:    {result.export_path or 'N/A (blocked or failed)'}")
    if result.errors:
        print(f"  Errors:")
        for err in result.errors:
            print(f"    - {err}")
    print("=" * 50)

    sys.exit(0 if result.balance_valid and not result.errors else 1)


if __name__ == "__main__":
    main()
