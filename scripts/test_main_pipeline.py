"""
test_main_pipeline.py
======================
End-to-end test suite for the fault-tolerant pipeline engine.

Covers:
  1. Happy path — all 6 phases complete successfully
  2. Branch A — low-confidence LLM output → Uncategorized + audit log
  3. Branch B — duplicate counterparty names → same UUID (idempotent)
  4. Branch C — balance break → FatalBalanceError → Phase 6 BLOCKED
  5. Phase 1 — CSV extraction and date/amount parsing
  6. Re-run idempotency — same data produces same TransactionIDs
"""

from __future__ import annotations

import csv
import json
import sys
from datetime import date
from decimal import Decimal
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from bank_statement_model import (
    BankStatementStore,
    Counterparty,
    Transaction,
    _safe_decimal,
    validate_running_balance,
)
from main_pipeline import (
    AuditEntry,
    BankStatementPipeline,
    ClassificationError,
    ClassifiedRecord,
    CleanedRecord,
    ExtractionError,
    FatalBalanceError,
    PipelineError,
    PipelineResult,
    RawRecord,
    _stub_llm_classify,
)

FIXTURES_DIR = Path(__file__).resolve().parent / "test_fixtures"


class TestResult:
    def __init__(self) -> None:
        self.passed: list[str] = []
        self.failed: list[tuple[str, str]] = []

    def ok(self, name: str) -> None:
        self.passed.append(name)

    def fail(self, name: str, reason: str) -> None:
        self.failed.append((name, reason))

    def summary(self) -> None:
        total = len(self.passed) + len(self.failed)
        print(f"\n{'=' * 60}")
        print(f"RESULTS: {len(self.passed)}/{total} passed")
        if self.failed:
            print("FAILURES:")
            for name, reason in self.failed:
                print(f"  FAIL [{name}] {reason}")
        else:
            print("ALL TESTS PASSED")
        print(f"{'=' * 60}")


def _write_csv(filename: str, rows: list[dict]) -> Path:
    FIXTURES_DIR.mkdir(parents=True, exist_ok=True)
    path = FIXTURES_DIR / filename
    with open(path, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=rows[0].keys())
        w.writeheader()
        w.writerows(rows)
    return path


HAPPY_CSV_ROWS = [
    {"Date": "2026-01-03", "Description": "FAIRPRICE ONLINE  #SG1234", "Amount": "-85.40", "Balance": "4914.60"},
    {"Date": "2026-01-05", "Description": "GRAB RIDE  SGP-9988", "Amount": "-22.50", "Balance": "4892.10"},
    {"Date": "2026-01-06", "Description": "SINGTEL MOBILE BILL JAN", "Amount": "-55.00", "Balance": "4837.10"},
    {"Date": "2026-01-10", "Description": "INCOMING TRANSFER FROM CLIENT A", "Amount": "5000.00", "Balance": "9837.10"},
    {"Date": "2026-01-12", "Description": "FAIRPRICE ONLINE  #SG5678", "Amount": "-112.30", "Balance": "9724.80"},
    {"Date": "2026-01-15", "Description": "GRAB RIDE  SGP-4455", "Amount": "-18.90", "Balance": "9705.90"},
    {"Date": "2026-01-31", "Description": "MONTHLY SVC FEE", "Amount": "-2.00", "Balance": "9703.90"},
]


def test_happy_path_full_pipeline(r: TestResult) -> None:
    name = "Happy Path — Full Pipeline (6 Phases)"
    try:
        csv_path = _write_csv("happy.csv", HAPPY_CSV_ROWS)
        output = FIXTURES_DIR / "happy_output"

        pipeline = BankStatementPipeline(
            source_path=csv_path,
            initial_balance=Decimal("5000.00"),
            output_dir=output,
        )
        result = pipeline.run()

        assert result.total_transactions == 7, f"Expected 7 txns, got {result.total_transactions}"
        assert result.total_counterparties >= 4, f"Expected >= 4 CPs, got {result.total_counterparties}"
        assert result.balance_valid is True, "Balance should be valid"
        assert result.export_path is not None, "Export path should exist"
        assert len(result.errors) == 0, f"Unexpected errors: {result.errors}"

        assert (output / "counterparty.csv").exists(), "counterparty.csv missing"
        assert (output / "transaction.csv").exists(), "transaction.csv missing"

        r.ok(name)
    except (AssertionError, Exception) as e:
        r.fail(name, str(e))


def test_branch_a_low_confidence(r: TestResult) -> None:
    name = "Branch A — Low Confidence → Uncategorized + Audit Log"
    try:
        csv_path = _write_csv("branch_a.csv", HAPPY_CSV_ROWS)
        output = FIXTURES_DIR / "branch_a_output"

        pipeline = BankStatementPipeline(
            source_path=csv_path,
            initial_balance=Decimal("5000.00"),
            output_dir=output,
            confidence_threshold=0.99,
        )
        result = pipeline.run()

        assert result.audit_log_count > 0, f"Expected audit entries, got {result.audit_log_count}"
        assert result.balance_valid is True

        audit_path = output / "audit_log.json"
        assert audit_path.exists(), "audit_log.json should exist"
        with open(audit_path, encoding="utf-8") as f:
            audit_data = json.load(f)
        assert len(audit_data) > 0
        assert audit_data[0]["Reason"] != ""

        classified = pipeline._classified_records
        uncategorized = [c for c in classified if c.Category == "Uncategorized"]
        assert len(uncategorized) > 0, "Should have Uncategorized records at 0.99 threshold"

        r.ok(name)
    except (AssertionError, Exception) as e:
        r.fail(name, str(e))


def test_branch_b_counterparty_dedup(r: TestResult) -> None:
    name = "Branch B — Counterparty Deduplication (Same UUID)"
    try:
        csv_path = _write_csv("branch_b.csv", HAPPY_CSV_ROWS)
        output = FIXTURES_DIR / "branch_b_output"

        pipeline = BankStatementPipeline(
            source_path=csv_path,
            initial_balance=Decimal("5000.00"),
            output_dir=output,
        )
        result = pipeline.run()

        cp_csv = output / "counterparty.csv"
        assert cp_csv.exists()
        with open(cp_csv, encoding="utf-8") as f:
            reader = csv.DictReader(f)
            cp_names = [row["CounterpartyName"] for row in reader]

        name_counts = {}
        for n in cp_names:
            name_counts[n] = name_counts.get(n, 0) + 1
        for n, count in name_counts.items():
            assert count == 1, f"Counterparty '{n}' appears {count} times (expected 1)"

        fairprice_ids = set()
        for rec in pipeline._enriched_records:
            if rec["CounterpartyName"] == "FairPrice":
                fairprice_ids.add(rec["CounterpartyID"])
        assert len(fairprice_ids) == 1, \
            f"FairPrice should have 1 UUID but has {len(fairprice_ids)}: {fairprice_ids}"

        r.ok(name)
    except (AssertionError, Exception) as e:
        r.fail(name, str(e))


def test_branch_c_balance_break(r: TestResult) -> None:
    name = "Branch C — Balance Break → FatalBalanceError → Phase 6 BLOCKED"
    try:
        broken_rows = list(HAPPY_CSV_ROWS)
        broken_rows[3] = {
            "Date": "2026-01-10",
            "Description": "INCOMING TRANSFER FROM CLIENT A",
            "Amount": "5000.00",
            "Balance": "9999.99",
        }

        csv_path = _write_csv("branch_c.csv", broken_rows)
        output = FIXTURES_DIR / "branch_c_output"

        pipeline = BankStatementPipeline(
            source_path=csv_path,
            initial_balance=Decimal("5000.00"),
            output_dir=output,
        )
        result = pipeline.run()

        assert result.balance_valid is False, "Balance should be INVALID"
        assert len(result.errors) > 0, "Should have error messages"
        assert any("Fatal" in e or "balance" in e.lower() for e in result.errors), \
            f"Expected fatal balance error in: {result.errors}"
        assert result.export_path is None, \
            f"Phase 6 should be BLOCKED but export_path={result.export_path}"

        assert not (output / "transaction.csv").exists(), \
            "Phase 6 CSV export should NOT exist after Branch C"

        r.ok(name)
    except (AssertionError, Exception) as e:
        r.fail(name, str(e))


def test_phase2_date_and_amount_parsing(r: TestResult) -> None:
    name = "Phase 2 — Flexible Date & Amount Parsing"
    try:
        rows = [
            {"Date": "03/01/2026", "Description": "TEST A", "Amount": "(100.50)", "Balance": "4,899.50"},
            {"Date": "05-01-2026", "Description": "TEST B", "Amount": "SGD 200.00", "Balance": "4,699.50"},
            {"Date": "2026-01-10", "Description": "TEST C", "Amount": "1,000.00", "Balance": "5,699.50"},
        ]
        csv_path = _write_csv("parsing.csv", rows)
        output = FIXTURES_DIR / "parse_output"

        pipeline = BankStatementPipeline(
            source_path=csv_path,
            initial_balance=Decimal("5000.00"),
            output_dir=output,
        )
        pipeline.phase_1_extract()
        pipeline.phase_2_clean()

        cleaned = pipeline._cleaned_records
        assert len(cleaned) == 3, f"Expected 3 cleaned, got {len(cleaned)}"
        assert cleaned[0].Amount == Decimal("-100.50"), f"Got {cleaned[0].Amount}"
        assert cleaned[0].Balance == Decimal("4899.50"), f"Got {cleaned[0].Balance}"
        assert cleaned[1].Amount == Decimal("200.00")
        assert cleaned[2].Amount == Decimal("1000.00")

        r.ok(name)
    except (AssertionError, Exception) as e:
        r.fail(name, str(e))


def test_rerun_idempotency(r: TestResult) -> None:
    name = "Re-Run Idempotency — Same Data = Same TransactionIDs"
    try:
        csv_path = _write_csv("idempotent.csv", HAPPY_CSV_ROWS)

        pipeline1 = BankStatementPipeline(
            source_path=csv_path,
            initial_balance=Decimal("5000.00"),
            output_dir=FIXTURES_DIR / "run1",
        )
        pipeline1.phase_1_extract()
        pipeline1.phase_2_clean()
        pipeline1.phase_3_classify()
        pipeline1.phase_4_map_keys()

        ids_run1 = [rec["CounterpartyID"] + "|" + rec["RawDescription"][:20]
                     for rec in pipeline1._enriched_records]

        pipeline2 = BankStatementPipeline(
            source_path=csv_path,
            initial_balance=Decimal("5000.00"),
            output_dir=FIXTURES_DIR / "run2",
        )
        pipeline2.phase_1_extract()
        pipeline2.phase_2_clean()
        pipeline2.phase_3_classify()
        pipeline2.phase_4_map_keys()

        ids_run2 = [rec["CounterpartyID"] + "|" + rec["RawDescription"][:20]
                     for rec in pipeline2._enriched_records]

        assert ids_run1 == ids_run2, "TransactionIDs differ between runs — not idempotent!"

        r.ok(name)
    except (AssertionError, Exception) as e:
        r.fail(name, str(e))


def test_empty_csv_raises_extraction_error(r: TestResult) -> None:
    name = "Empty CSV → ExtractionError"
    try:
        FIXTURES_DIR.mkdir(parents=True, exist_ok=True)
        csv_path = FIXTURES_DIR / "empty.csv"
        with open(csv_path, "w", newline="", encoding="utf-8") as f:
            f.write("Date,Description,Amount,Balance\n")
        pipeline = BankStatementPipeline(
            source_path=csv_path,
            initial_balance=Decimal("0"),
            output_dir=FIXTURES_DIR / "empty_out",
        )
        try:
            pipeline.phase_1_extract()
            assert False, "Should have raised ExtractionError"
        except ExtractionError:
            pass

        r.ok(name)
    except (AssertionError, Exception) as e:
        r.fail(name, str(e))


def test_unsupported_format_raises(r: TestResult) -> None:
    name = "Unsupported File Format → ExtractionError"
    try:
        pipeline = BankStatementPipeline(
            source_path="statement.xlsx",
            initial_balance=Decimal("0"),
            output_dir=FIXTURES_DIR / "fmt_out",
        )
        try:
            pipeline.phase_1_extract()
            assert False, "Should have raised ExtractionError"
        except ExtractionError as e:
            assert "Unsupported" in str(e)

        r.ok(name)
    except (AssertionError, Exception) as e:
        r.fail(name, str(e))


def test_auto_extract_balance_brought_forward(r: TestResult) -> None:
    name = "Auto-Extract — 'Balance Brought Forward' Row"
    try:
        rows_with_bbf = [
            {"Date": "", "Description": "Balance Brought Forward", "Amount": "", "Balance": "5,000.00"},
        ] + HAPPY_CSV_ROWS
        csv_path = _write_csv("bbf.csv", rows_with_bbf)
        output = FIXTURES_DIR / "bbf_output"

        pipeline = BankStatementPipeline(
            source_path=csv_path,
            output_dir=output,
        )
        result = pipeline.run()

        assert pipeline.extracted_initial_balance == Decimal("5000.00"), \
            f"Expected 5000.00, got {pipeline.extracted_initial_balance}"
        assert pipeline.initial_balance == Decimal("5000.00"), \
            f"initial_balance property should be 5000.00, got {pipeline.initial_balance}"
        assert result.total_transactions == 7, \
            f"BBF row should NOT be counted as transaction, got {result.total_transactions}"
        assert result.balance_valid is True
        assert len(result.errors) == 0, f"Unexpected errors: {result.errors}"

        r.ok(name)
    except (AssertionError, Exception) as e:
        r.fail(name, str(e))


def test_auto_extract_with_explicit_override(r: TestResult) -> None:
    name = "Explicit --initial-balance Overrides Auto-Extract"
    try:
        rows_with_bbf = [
            {"Date": "", "Description": "Balance Brought Forward", "Amount": "", "Balance": "5,000.00"},
        ] + HAPPY_CSV_ROWS
        csv_path = _write_csv("bbf_override.csv", rows_with_bbf)
        output = FIXTURES_DIR / "bbf_override_output"

        pipeline = BankStatementPipeline(
            source_path=csv_path,
            initial_balance=Decimal("5000.00"),
            output_dir=output,
        )
        result = pipeline.run()

        assert pipeline.extracted_initial_balance == Decimal("5000.00"), \
            "Auto-extract should still capture BBF"
        assert pipeline._initial_balance_input == Decimal("5000.00"), \
            "Explicit input should be preserved"
        assert pipeline.initial_balance == Decimal("5000.00")
        assert result.balance_valid is True

        r.ok(name)
    except (AssertionError, Exception) as e:
        r.fail(name, str(e))


def test_no_balance_at_all_raises_pipeline_error(r: TestResult) -> None:
    name = "No BBF + No --initial-balance → PipelineError"
    try:
        csv_path = _write_csv("no_bbf.csv", HAPPY_CSV_ROWS)
        output = FIXTURES_DIR / "no_bbf_output"

        pipeline = BankStatementPipeline(
            source_path=csv_path,
            output_dir=output,
        )
        result = pipeline.run()

        assert len(result.errors) > 0, "Should have errors"
        assert any("initial balance" in e.lower() for e in result.errors), \
            f"Expected initial balance error in: {result.errors}"

        r.ok(name)
    except (AssertionError, Exception) as e:
        r.fail(name, str(e))


def main() -> None:
    print("=" * 60)
    print("Pipeline Engine — End-to-End Test Suite")
    print("=" * 60)

    result = TestResult()
    tests = [
        test_happy_path_full_pipeline,
        test_branch_a_low_confidence,
        test_branch_b_counterparty_dedup,
        test_branch_c_balance_break,
        test_phase2_date_and_amount_parsing,
        test_rerun_idempotency,
        test_empty_csv_raises_extraction_error,
        test_unsupported_format_raises,
        test_auto_extract_balance_brought_forward,
        test_auto_extract_with_explicit_override,
        test_no_balance_at_all_raises_pipeline_error,
    ]

    for test_fn in tests:
        print(f"\nRunning: {test_fn.__name__} ...", end=" ")
        test_fn(result)
        status = "PASS" if test_fn.__name__ in result.passed else "FAIL"
        print(status)

    result.summary()
    return 0 if not result.failed else 1


if __name__ == "__main__":
    sys.exit(main())
