---
name: db-storage-specialist-skill
description: Responsible for structured relational data storage, maintaining referential integrity (guaranteeing Foreign Key existence in the Counterparty table), and exporting final results to a local database format or multi-sheet relational layout.
license: Apache-2.0
compatibility: opencode
metadata:
  audience: developers
  workflow: database-export
---

## What I do

- Write processed transaction and counterparty records into structured relational storage
- Maintain referential integrity between Transaction and Counterparty tables via Foreign Key constraints
- Export final results to SQLite database, multi-sheet Excel, or structured CSV files
- Validate all records conform to the Data Dictionary before writing
- Handle upsert operations for existing counterparties and transactions
- Support idempotent writes to prevent duplicate records on re-runs

## When to use me

Use this skill when:
- Writing classified, cleaned bank statement data to a relational database
- Ensuring Foreign Key constraints between Transaction and Counterparty tables
- Exporting processed financial data to SQLite, Excel, or CSV formats
- Validating data integrity before final storage
- Building or updating a persistent local database from multiple bank statement files

Do NOT use for:
- Extracting data from PDF files (use `pdf-specialist-skill`)
- Data cleaning or UUID generation (use `relational-data-processor-skill`)
- Classifying transactions via LLM (use `transaction-classifier-skill`)

## Prerequisites

### Python Libraries

- **sqlite3**: Built-in Python SQLite support for local database operations
- **pandas**: Data export to Excel/CSV formats
- **openpyxl**: Excel file writing engine for pandas

Install dependencies:
```bash
pip install pandas openpyxl
```

## Steps

### Step 1: Validate Input Data

Verify all records conform to the Data Dictionary schema before writing:

```python
import pandas as pd

def validate_before_export(transactions_df, counterparties_df):
    assert transactions_df['TransactionID'].is_unique, "Duplicate TransactionIDs found"
    assert counterparties_df['CounterpartyID'].is_unique, "Duplicate CounterpartyIDs found"
    fk_orphans = transactions_df[
        ~transactions_df['CounterpartyID'].isin(counterparties_df['CounterpartyID'])
    ]
    assert fk_orphans.empty, f"Orphaned Foreign Keys: {len(fk_orphans)} records"
    assert transactions_df['RawDescription'].str.len().le(255).all()
    assert transactions_df['Category'].str.len().le(50).all()
    assert transactions_df['IsRecurring'].isin([0, 1]).all()
    return True
```

### Step 2: Export to SQLite Database

Write records to a local SQLite database with proper schema and constraints:

```python
import sqlite3

def export_to_sqlite(transactions_df, counterparties_df, db_path):
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS Counterparty (
            CounterpartyID TEXT(36) PRIMARY KEY,
            CounterpartyName TEXT(100) NOT NULL
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS Transaction (
            TransactionID TEXT(36) PRIMARY KEY,
            CounterpartyID TEXT(36) NOT NULL,
            TransactionDate DATE NOT NULL,
            RawDescription TEXT(255),
            Amount DECIMAL(12,2),
            Balance DECIMAL(12,2),
            Category TEXT(50),
            IsRecurring INTEGER DEFAULT 0,
            FOREIGN KEY (CounterpartyID) REFERENCES Counterparty(CounterpartyID)
        )
    """)

    counterparties_df.to_sql('Counterparty', conn, if_exists='replace', index=False)
    transactions_df.to_sql('Transaction', conn, if_exists='replace', index=False)

    conn.execute("PRAGMA foreign_keys = ON")
    conn.commit()
    conn.close()
```

### Step 3: Export to Multi-Sheet Excel

Write records to a multi-sheet Excel workbook with relational layout:

```python
def export_to_excel(transactions_df, counterparties_df, excel_path):
    with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
        counterparties_df.to_excel(writer, sheet_name='Counterparty', index=False)
        transactions_df.to_excel(writer, sheet_name='Transaction', index=False)
```

### Step 4: Export to Structured CSV Files

Write each table to a separate CSV file:

```python
def export_to_csv(transactions_df, counterparties_df, output_dir):
    import os
    os.makedirs(output_dir, exist_ok=True)
    counterparties_df.to_csv(
        os.path.join(output_dir, 'counterparty.csv'), index=False
    )
    transactions_df.to_csv(
        os.path.join(output_dir, 'transaction.csv'), index=False
    )
```

### Step 5: Verify Referential Integrity

Post-export validation to ensure data integrity:

```python
def verify_referential_integrity(db_path):
    conn = sqlite3.connect(db_path)
    orphans = conn.execute("""
        SELECT COUNT(*) FROM Transaction t
        LEFT JOIN Counterparty c ON t.CounterpartyID = c.CounterpartyID
        WHERE c.CounterpartyID IS NULL
    """).fetchone()[0]
    assert orphans == 0, f"Found {orphans} orphaned transactions"
    conn.close()
    return True
```

## Best Practices

- Always validate data against the Data Dictionary before writing
- Use `if_exists='replace'` for full refreshes or upsert logic for incremental updates
- Enable `PRAGMA foreign_keys = ON` in SQLite to enforce constraints
- Export to multiple formats (SQLite + Excel) for redundancy
- Verify referential integrity after every write operation
- Use transactions (BEGIN/COMMIT) for atomic writes

## Common Issues

### Foreign Key Constraint Violations

**Issue**: Transaction records reference CounterpartyIDs that don't exist

**Solution**: Ensure Counterparty table is written BEFORE Transaction table. Run `validate_before_export()` to catch orphans before writing.

### Duplicate Records on Re-Run

**Issue**: Running the pipeline multiple times creates duplicate records

**Solution**: Use TransactionID as the primary key with `INSERT OR REPLACE` logic, or clear tables before re-writing.

### SQLite Database Locked

**Issue**: Database is locked by another process

**Solution**: Close all connections before re-opening. Use a timeout in the connection string.

## Verification Commands

```bash
python -c "
import sqlite3
conn = sqlite3.connect('bank_statements.db')
tables = conn.execute(\"SELECT name FROM sqlite_master WHERE type='table'\").fetchall()
print(f'Tables: {[t[0] for t in tables]}')
for table in tables:
    count = conn.execute(f'SELECT COUNT(*) FROM {table[0]}').fetchone()[0]
    print(f'  {table[0]}: {count} rows')
conn.close()
"
```
