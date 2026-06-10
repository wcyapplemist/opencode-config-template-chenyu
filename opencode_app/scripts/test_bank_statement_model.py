"""
test_bank_statement_model.py
=============================
Mock-data test suite that verifies:

  1. UUID Idempotency   - uuid5 produces identical keys across re-runs
  2. Decimal Precision  - no floating-point traps in Amount / Balance
  3. Referential Integrity - FK constraints enforced by SQLite
  4. Duplicate Safety   - INSERT OR REPLACE is idempotent on re-insert
  5. Balance Validation - running-balance checker catches discrepancies
"""

from __future__ import annotations

import sys
from datetime import date
from decimal import Decimal
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from bank_statement_model import (
    Counterparty,
    CounterpartyRegistry,
    Transaction,
    BankStatementStore,
    IntegrityError,
    _safe_decimal,
    validate_running_balance,
)

MOCK_RECORDS = [
    {
        "TransactionDate": "2026-01-03",
        "RawDescription": "FAIRPRICE ONLINE  #SG1234",
        "Amount": "-85.40",
        "Balance": "4914.60",
        "CounterpartyName": "FairPrice",
        "Category": "F&B",
        "IsRecurring": False,
    },
    {
        "TransactionDate": "2026-01-05",
        "RawDescription": "GRAB RIDE  SGP-9988",
        "Amount": "-22.50",
        "Balance": "4892.10",
        "CounterpartyName": "Grab",
        "Category": "Transport",
        "IsRecurring": False,
    },
    {
        "TransactionDate": "2026-01-06",
        "RawDescription": "SINGTEL MOBILE BILL JAN",
        "Amount": "-55.00",
        "Balance": "4837.10",
        "CounterpartyName": "Singtel",
        "Category": "Telecommunications",
        "IsRecurring": True,
    },
    {
        "TransactionDate": "2026-01-10",
        "RawDescription": "INCOMING TRANSFER FROM CLIENT A",
        "Amount": "5000.00",
        "Balance": "9837.10",
        "CounterpartyName": "Client A",
        "Category": "Transfer",
        "IsRecurring": False,
    },
    {
        "TransactionDate": "2026-01-12",
        "RawDescription": "FAIRPRICE ONLINE  #SG5678",
        "Amount": "-112.30",
        "Balance": "9724.80",
        "CounterpartyName": "FairPrice",
        "Category": "F&B",
        "IsRecurring": False,
    },
    {
        "TransactionDate": "2026-01-15",
        "RawDescription": "GRAB RIDE  SGP-4455",
        "Amount": "-18.90",
        "Balance": "9705.90",
        "CounterpartyName": "Grab",
        "Category": "Transport",
        "IsRecurring": False,
    },
    {
        "TransactionDate": "2026-01-31",
        "RawDescription": "MONTHLY SVC FEE",
        "Amount": "-2.00",
        "Balance": "9703.90",
        "CounterpartyName": "DBS Bank",
        "Category": "Bank Fees",
        "IsRecurring": True,
    },
]

INITIAL_BALANCE = Decimal("5000.00")


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


def test_uuid_idempotency(r: TestResult) -> None:
    name = "UUID Idempotency"
    try:
        id_a = Counterparty.generate_id("FairPrice")
        id_b = Counterparty.generate_id("FairPrice")
        id_c = Counterparty.generate_id("fairprice")
        assert id_a == id_b, "Same name produced different UUIDs"
        assert id_a == id_c, "Case-insensitive name produced different UUID"

        txn_id_a = Transaction.generate_id(date(2026, 1, 3), Decimal("-85.40"), "FAIRPRICE ONLINE  #SG1234")
        txn_id_b = Transaction.generate_id(date(2026, 1, 3), Decimal("-85.40"), "FAIRPRICE ONLINE  #SG1234")
        assert txn_id_a == txn_id_b, "Same transaction seed produced different UUIDs"

        txn_id_c = Transaction.generate_id(date(2026, 1, 3), Decimal("-85.41"), "FAIRPRICE ONLINE  #SG1234")
        assert txn_id_a != txn_id_c, "Different amount produced same UUID"

        r.ok(name)
    except AssertionError as e:
        r.fail(name, str(e))


def test_decimal_precision(r: TestResult) -> None:
    name = "Decimal Precision Safety"
    try:
        float_sum = 0.1 + 0.2
        assert float_sum != 0.3, "float 0.1+0.2 should NOT equal 0.3 (baseline check)"

        dec_sum = Decimal("0.1") + Decimal("0.2")
        assert dec_sum == Decimal("0.3"), "Decimal 0.1+0.2 MUST equal 0.3"

        val = _safe_decimal(0.1)
        assert isinstance(val, Decimal), "_safe_decimal must return Decimal"

        tricky = _safe_decimal("999999999999.99") + _safe_decimal("0.01")
        assert tricky == Decimal("1000000000000.00"), "Large decimal addition failed"

        neg = _safe_decimal("-33.33") + _safe_decimal("-66.67")
        assert neg == Decimal("-100.00"), "Negative decimal addition failed"

        r.ok(name)
    except (AssertionError, Exception) as e:
        r.fail(name, str(e))


def test_balance_validation_passes(r: TestResult) -> None:
    name = "Balance Validation - Correct Chain"
    try:
        transactions = []
        running = INITIAL_BALANCE
        for rec in MOCK_RECORDS:
            amt = Decimal(rec["Amount"])
            running += amt
            transactions.append(
                Transaction(
                    TransactionID="dummy",
                    CounterpartyID="dummy",
                    TransactionDate=date.fromisoformat(rec["TransactionDate"]),
                    RawDescription=rec["RawDescription"],
                    Amount=amt,
                    Balance=Decimal(rec["Balance"]),
                    Category=rec["Category"],
                    IsRecurring=rec["IsRecurring"],
                )
            )
        errors = validate_running_balance(transactions, INITIAL_BALANCE)
        assert len(errors) == 0, f"Expected 0 balance errors, got {len(errors)}: {errors}"
        r.ok(name)
    except (AssertionError, Exception) as e:
        r.fail(name, str(e))


def test_balance_validation_catches_error(r: TestResult) -> None:
    name = "Balance Validation - Detects Discrepancy"
    try:
        transactions = [
            Transaction(
                TransactionID="t1",
                CounterpartyID="c1",
                TransactionDate=date(2026, 1, 1),
                RawDescription="TEST",
                Amount=Decimal("-100.00"),
                Balance=Decimal("999.00"),
                Category="Test",
                IsRecurring=False,
            ),
        ]
        errors = validate_running_balance(transactions, Decimal("1000.00"))
        assert len(errors) == 1, f"Expected 1 error, got {len(errors)}"
        assert errors[0]["expected_balance"] == "900.00"
        assert errors[0]["actual_balance"] == "999.00"
        r.ok(name)
    except (AssertionError, Exception) as e:
        r.fail(name, str(e))


def test_full_pipeline_and_idempotent_rerun(r: TestResult) -> None:
    name = "Full Pipeline + Idempotent Re-Run"
    try:
        with BankStatementStore() as store:
            count1, errs1 = store.insert_transactions(
                MOCK_RECORDS,
                validate_balance=True,
                initial_balance=INITIAL_BALANCE,
            )
            assert count1 == len(MOCK_RECORDS), f"Expected {len(MOCK_RECORDS)} txns, got {count1}"
            assert len(errs1) == 0, f"Balance errors on first run: {errs1}"

            cp_rows = store.query_counterparties()
            assert len(cp_rows) == 5, f"Expected 5 counterparties, got {len(cp_rows)}"

            cp_names = {c["CounterpartyName"] for c in cp_rows}
            assert "FairPrice" in cp_names
            assert "Grab" in cp_names
            assert "Singtel" in cp_names

            txn_rows = store.query_transactions()
            assert len(txn_rows) == 7

            fairprice_txns = store.query_transactions(counterparty_name="FairPrice")
            assert len(fairprice_txns) == 2, f"Expected 2 FairPrice txns, got {len(fairprice_txns)}"

            integrity = store.verify_integrity()
            assert len(integrity) == 0, f"Integrity issues: {integrity}"

            count2, errs2 = store.insert_transactions(
                MOCK_RECORDS,
                validate_balance=True,
                initial_balance=INITIAL_BALANCE,
            )
            assert count2 == len(MOCK_RECORDS)

            txn_after_rerun = store.query_transactions()
            txn_ids = [t["TransactionID"] for t in txn_after_rerun]
            assert len(txn_ids) == len(set(txn_ids)), "Duplicate TransactionIDs after re-run!"

            cp_after = store.query_counterparties()
            assert len(cp_after) == 5, "Counterparties duplicated on re-run"

            summary = store.counterparty_summary()
            assert len(summary) == 5
            fp = [s for s in summary if s["CounterpartyName"] == "FairPrice"][0]
            assert fp["TransactionCount"] == 2
            expected_fp_total = (Decimal("-85.40") + Decimal("-112.30")).quantize(Decimal("0.01"))
            assert Decimal(fp["TotalAmount"]) == expected_fp_total, \
                f"FairPrice total mismatch: {fp['TotalAmount']} vs {expected_fp_total}"

        r.ok(name)
    except (AssertionError, Exception) as e:
        r.fail(name, str(e))


def test_uuid_no_uuid4(r: TestResult) -> None:
    name = "UUID No uuid4 Usage"
    try:
        import inspect
        from bank_statement_model import Counterparty, Transaction

        src = inspect.getsource(Counterparty.generate_id) + inspect.getsource(Transaction.generate_id)
        assert "uuid4" not in src, "uuid4 must never be used - only uuid5 is allowed"
        assert "uuid5" in src, "uuid5 must be used for deterministic ID generation"
        r.ok(name)
    except (AssertionError, Exception) as e:
        r.fail(name, str(e))


def test_floating_point_trap_prevention(r: TestResult) -> None:
    name = "Floating-Point Trap Prevention"
    try:
        poison_values = [
            ("0.1", "0.2", "0.3"),
            ("0.7", "0.1", "0.8"),
            ("1.1", "2.2", "3.3"),
            ("0.01", "0.02", "0.03"),
            ("99.99", "0.01", "100.00"),
        ]
        for a, b, expected in poison_values:
            result = _safe_decimal(a) + _safe_decimal(b)
            assert result == Decimal(expected), \
                f"Decimal trap: {a} + {b} = {result}, expected {expected}"

        store = BankStatementStore()
        store.open()
        try:
            trap_records = [
                {
                    "TransactionDate": "2026-02-01",
                    "RawDescription": "SMALL PAYMENT A",
                    "Amount": "0.1",
                    "Balance": "1000.1",
                    "CounterpartyName": "Test Corp",
                    "Category": "Other",
                    "IsRecurring": False,
                },
                {
                    "TransactionDate": "2026-02-02",
                    "RawDescription": "SMALL PAYMENT B",
                    "Amount": "0.2",
                    "Balance": "1000.3",
                    "CounterpartyName": "Test Corp",
                    "Category": "Other",
                    "IsRecurring": False,
                },
            ]
            count, errs = store.insert_transactions(
                trap_records,
                validate_balance=True,
                initial_balance=Decimal("1000.00"),
            )
            assert len(errs) == 0, f"Floating-point trap caused balance errors: {errs}"
        finally:
            store.close()

        r.ok(name)
    except (AssertionError, Exception) as e:
        r.fail(name, str(e))


def test_counterparty_registry_dedup(r: TestResult) -> None:
    name = "Counterparty Registry Deduplication"
    try:
        reg = CounterpartyRegistry()
        cp1 = reg.resolve("Grab")
        cp2 = reg.resolve("Grab")
        cp3 = reg.resolve("GRAB")

        assert cp1.CounterpartyID == cp2.CounterpartyID, "Same name should return same ID"
        assert cp1.CounterpartyID == cp3.CounterpartyID, "Case-folded name should return same ID"
        assert reg.count() == 1, "Registry should have exactly 1 counterparty"

        reg.resolve("FairPrice")
        assert reg.count() == 2

        r.ok(name)
    except (AssertionError, Exception) as e:
        r.fail(name, str(e))


def main() -> None:
    print("=" * 60)
    print("Bank Statement Model - Mock Test Suite")
    print("=" * 60)

    result = TestResult()
    tests = [
        test_uuid_idempotency,
        test_decimal_precision,
        test_balance_validation_passes,
        test_balance_validation_catches_error,
        test_full_pipeline_and_idempotent_rerun,
        test_uuid_no_uuid4,
        test_floating_point_trap_prevention,
        test_counterparty_registry_dedup,
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
