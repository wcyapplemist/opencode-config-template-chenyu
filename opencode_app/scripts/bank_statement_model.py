"""
bank_statement_model.py
========================
Core data modeling module for the Bank Statement Automation Agent.

Implements the relational schema defined in the Data Dictionary:
  - Counterparty table (CounterpartyID PK, CounterpartyName)
  - Transaction table  (TransactionID PK, CounterpartyID FK, ...)

Key guarantees:
  1. UUID Idempotency  - uuid5 with deterministic seeds (no uuid4)
  2. Decimal Precision - decimal.Decimal for all Amount / Balance arithmetic
  3. Relational Integrity - FK enforcement between Transaction -> Counterparty
"""

from __future__ import annotations

import sqlite3
import uuid
from dataclasses import dataclass, field
from datetime import date
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from pathlib import Path
from typing import Dict, List, Optional, Tuple

COUNTERPARTY_NS = uuid.UUID("b3e1c2a4-5f6d-4a8b-9c7e-1d2f3a4b5c6d")
TRANSACTION_NS = uuid.UUID("d7f8e6c4-2a1b-4d3e-8f5a-6b7c8d9e0f1a")
TWO_PLACES = Decimal("0.01")

COUNTERPARTY_DDL = """
CREATE TABLE IF NOT EXISTS Counterparty (
    CounterpartyID   TEXT    PRIMARY KEY,
    CounterpartyName TEXT    NOT NULL UNIQUE
);
"""

TRANSACTION_DDL = """
CREATE TABLE IF NOT EXISTS BankTransaction (
    TransactionID   TEXT     PRIMARY KEY,
    CounterpartyID  TEXT     NOT NULL,
    TransactionDate TEXT     NOT NULL,
    RawDescription  TEXT,
    Amount          TEXT     NOT NULL,
    Balance         TEXT     NOT NULL,
    Category        TEXT,
    IsRecurring     INTEGER  DEFAULT 0,
    FOREIGN KEY (CounterpartyID) REFERENCES Counterparty(CounterpartyID)
);
"""


def _safe_decimal(value) -> Decimal:
    if isinstance(value, Decimal):
        return value
    try:
        return Decimal(str(value))
    except (InvalidOperation, ValueError) as exc:
        raise ValueError(f"Cannot convert to Decimal: {value!r}") from exc


@dataclass(frozen=True)
class Counterparty:
    CounterpartyID: str
    CounterpartyName: str

    @staticmethod
    def generate_id(name: str) -> str:
        seed = name.strip().lower()
        return str(uuid.uuid5(COUNTERPARTY_NS, seed))


@dataclass(frozen=True)
class Transaction:
    TransactionID: str
    CounterpartyID: str
    TransactionDate: date
    RawDescription: str
    Amount: Decimal
    Balance: Decimal
    Category: str
    IsRecurring: bool

    @staticmethod
    def generate_id(txn_date: date, amount: Decimal, raw_desc: str) -> str:
        seed = f"{txn_date.isoformat()}|{amount:.2f}|{raw_desc.strip()}"
        return str(uuid.uuid5(TRANSACTION_NS, seed))


class CounterpartyRegistry:
    def __init__(self) -> None:
        self._by_name: Dict[str, Counterparty] = {}
        self._by_id: Dict[str, Counterparty] = {}

    def resolve(self, name: str) -> Counterparty:
        key = name.strip().lower()
        display_name = name.strip()
        if key not in self._by_name:
            cp = Counterparty(
                CounterpartyID=Counterparty.generate_id(display_name),
                CounterpartyName=display_name,
            )
            self._by_name[key] = cp
            self._by_id[cp.CounterpartyID] = cp
        return self._by_name[key]

    def get_by_id(self, cp_id: str) -> Optional[Counterparty]:
        return self._by_id.get(cp_id)

    def all(self) -> List[Counterparty]:
        return list(self._by_name.values())

    def count(self) -> int:
        return len(self._by_name)


def validate_running_balance(
    transactions: List[Transaction],
    initial_balance: Decimal = Decimal("0"),
    tolerance: Decimal = Decimal("0.01"),
) -> List[dict]:
    errors: List[dict] = []
    running = _safe_decimal(initial_balance)
    for idx, txn in enumerate(transactions):
        running = (running + txn.Amount).quantize(TWO_PLACES, rounding=ROUND_HALF_UP)
        diff = abs(running - txn.Balance)
        if diff > tolerance:
            errors.append(
                {
                    "index": idx,
                    "transaction_id": txn.TransactionID,
                    "expected_balance": str(running),
                    "actual_balance": str(txn.Balance),
                    "discrepancy": str(diff),
                }
            )
    return errors


class BankStatementStore:
    def __init__(self, db_path: str | Path = ":memory:") -> None:
        self._db_path = str(db_path)
        self._conn: Optional[sqlite3.Connection] = None
        self._registry = CounterpartyRegistry()

    # -- context manager --------------------------------------------------

    def __enter__(self) -> "BankStatementStore":
        self.open()
        return self

    def __exit__(self, *exc) -> None:
        self.close()

    # -- connection management -------------------------------------------

    def open(self) -> None:
        self._conn = sqlite3.connect(self._db_path)
        self._conn.execute("PRAGMA foreign_keys = ON")
        self._conn.executescript(COUNTERPARTY_DDL + TRANSACTION_DDL)
        self._conn.commit()

    def close(self) -> None:
        if self._conn:
            self._conn.close()
            self._conn = None

    # -- counterparty operations -----------------------------------------

    @property
    def registry(self) -> CounterpartyRegistry:
        return self._registry

    def resolve_counterparty(self, name: str) -> Counterparty:
        return self._registry.resolve(name)

    # -- transaction operations ------------------------------------------

    def insert_transactions(
        self,
        records: List[dict],
        validate_balance: bool = True,
        initial_balance: Decimal = Decimal("0"),
    ) -> Tuple[int, List[dict]]:
        if self._conn is None:
            raise RuntimeError("Database not open. Use context manager or .open().")

        transactions: List[Transaction] = []
        for rec in records:
            cp = self._registry.resolve(rec["CounterpartyName"])
            amount = _safe_decimal(rec["Amount"])
            balance = _safe_decimal(rec["Balance"])
            txn_date = (
                rec["TransactionDate"]
                if isinstance(rec["TransactionDate"], date)
                else date.fromisoformat(str(rec["TransactionDate"]))
            )
            txn = Transaction(
                TransactionID=Transaction.generate_id(
                    txn_date, amount, rec["RawDescription"]
                ),
                CounterpartyID=cp.CounterpartyID,
                TransactionDate=txn_date,
                RawDescription=rec["RawDescription"][:255],
                Amount=amount.quantize(TWO_PLACES, rounding=ROUND_HALF_UP),
                Balance=balance.quantize(TWO_PLACES, rounding=ROUND_HALF_UP),
                Category=rec.get("Category", "Other")[:50],
                IsRecurring=bool(rec.get("IsRecurring", False)),
            )
            transactions.append(txn)

        balance_errors: List[dict] = []
        if validate_balance:
            balance_errors = validate_running_balance(transactions, initial_balance)

        try:
            self._write_to_db(transactions)
        except sqlite3.IntegrityError as exc:
            raise IntegrityError(str(exc)) from exc

        return len(transactions), balance_errors

    def _write_to_db(self, transactions: List[Transaction]) -> None:
        assert self._conn is not None
        cur = self._conn.cursor()

        for cp in self._registry.all():
            cur.execute(
                """
                INSERT OR IGNORE INTO Counterparty (CounterpartyID, CounterpartyName)
                VALUES (?, ?)
                """,
                (cp.CounterpartyID, cp.CounterpartyName),
            )

        for txn in transactions:
            cur.execute(
                """
                INSERT OR REPLACE INTO BankTransaction
                    (TransactionID, CounterpartyID, TransactionDate,
                     RawDescription, Amount, Balance, Category, IsRecurring)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    txn.TransactionID,
                    txn.CounterpartyID,
                    txn.TransactionDate.isoformat(),
                    txn.RawDescription,
                    str(txn.Amount),
                    str(txn.Balance),
                    txn.Category,
                    1 if txn.IsRecurring else 0,
                ),
            )

        self._conn.commit()

    # -- query helpers ---------------------------------------------------

    def query_transactions(self, counterparty_name: Optional[str] = None) -> List[dict]:
        assert self._conn is not None
        if counterparty_name:
            rows = self._conn.execute(
                """
                SELECT t.* FROM BankTransaction t
                JOIN Counterparty c ON t.CounterpartyID = c.CounterpartyID
                WHERE c.CounterpartyName = ?
                ORDER BY t.TransactionDate
                """,
                (counterparty_name,),
            ).fetchall()
        else:
            rows = self._conn.execute(
                "SELECT * FROM BankTransaction ORDER BY TransactionDate"
            ).fetchall()

        cols = [
            "TransactionID", "CounterpartyID", "TransactionDate",
            "RawDescription", "Amount", "Balance", "Category", "IsRecurring",
        ]
        return [dict(zip(cols, row)) for row in rows]

    def query_counterparties(self) -> List[dict]:
        assert self._conn is not None
        rows = self._conn.execute(
            "SELECT * FROM Counterparty ORDER BY CounterpartyName"
        ).fetchall()
        return [dict(zip(["CounterpartyID", "CounterpartyName"], row)) for row in rows]

    def counterparty_summary(self) -> List[dict]:
        assert self._conn is not None
        rows = self._conn.execute(
            """
            SELECT c.CounterpartyName,
                   COUNT(*)                        AS txn_count,
                   SUM(CAST(t.Amount AS REAL))     AS total_amount
            FROM BankTransaction t
            JOIN Counterparty c ON t.CounterpartyID = c.CounterpartyID
            GROUP BY c.CounterpartyName
            ORDER BY c.CounterpartyName
            """
        ).fetchall()
        return [
            {
                "CounterpartyName": r[0],
                "TransactionCount": r[1],
                "TotalAmount": str(
                    Decimal(str(round(r[2], 2))).quantize(
                        TWO_PLACES, rounding=ROUND_HALF_UP
                    )
                ),
            }
            for r in rows
        ]

    # -- export ----------------------------------------------------------

    def export_to_csv(self, output_dir: str | Path) -> None:
        import csv

        out = Path(output_dir)
        out.mkdir(parents=True, exist_ok=True)

        counterparties = self.query_counterparties()
        with open(out / "counterparty.csv", "w", newline="", encoding="utf-8") as f:
            w = csv.DictWriter(f, fieldnames=["CounterpartyID", "CounterpartyName"])
            w.writeheader()
            w.writerows(counterparties)

        transactions = self.query_transactions()
        fields = [
            "TransactionID", "CounterpartyID", "TransactionDate",
            "RawDescription", "Amount", "Balance", "Category", "IsRecurring",
        ]
        with open(out / "transaction.csv", "w", newline="", encoding="utf-8") as f:
            w = csv.DictWriter(f, fieldnames=fields)
            w.writeheader()
            w.writerows(transactions)

    def verify_integrity(self) -> List[str]:
        assert self._conn is not None
        issues: List[str] = []
        orphans = self._conn.execute(
            """
            SELECT COUNT(*) FROM BankTransaction t
            LEFT JOIN Counterparty c ON t.CounterpartyID = c.CounterpartyID
            WHERE c.CounterpartyID IS NULL
            """
        ).fetchone()[0]
        if orphans > 0:
            issues.append(f"{orphans} orphaned transaction(s) with invalid CounterpartyID")
        return issues


class IntegrityError(Exception):
    pass
