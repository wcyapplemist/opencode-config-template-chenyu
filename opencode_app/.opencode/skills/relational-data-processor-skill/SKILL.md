---
name: relational-data-processor-skill
description: Employs deterministic Python scripts (Pandas / SQL) to perform precise data cleaning, exact mathematical aggregations, UUID generation for Primary Keys, and strict type validation for financial records.
license: Apache-2.0
compatibility: opencode
metadata:
  audience: developers
  workflow: data-processing
---

## What I do

- Clean and standardize raw bank statement data (noise removal, date formatting, description trimming)
- Generate deterministic UUIDs for Primary Keys (TransactionID, CounterpartyID)
- Cross-reference inferred CounterpartyName with existing Counterparty table entries
- Establish Foreign Key mappings ensuring referential integrity between Transaction and Counterparty tables
- Perform exact mathematical aggregations and running balance validation using Pandas
- Apply strict type validation for all financial record fields per the Data Dictionary

## When to use me

Use this skill when:
- Processing raw extracted bank statement data into clean, standardized records
- Generating UUIDs for transaction and counterparty primary keys
- Mapping CounterpartyID foreign keys between Transaction and Counterparty tables
- Validating numerical fields (Amount signs, Balance continuity) deterministically
- Splitting and standardizing date fields (TransactionDate, ValueDate)
- Ensuring all data conforms to the Data Dictionary schema constraints

Do NOT use for:
- Extracting data from PDF files (use `pdf-specialist-skill`)
- Classifying transactions or inferring counterparties via LLM (use `transaction-classifier-skill`)
- Writing final results to database storage (use `db-storage-specialist-skill`)

## Prerequisites

### Python Libraries

- **pandas**: Data manipulation, cleaning, and aggregation
- **uuid**: Deterministic UUID generation for primary keys
- **sqlite3**: In-memory relational validation (Python built-in)

Install dependencies:
```bash
pip install pandas
```

## Steps

### Step 1: Data Cleaning & Format Standardization

Remove noise from raw extracted data and standardize all fields:

```python
import pandas as pd
import re

def clean_raw_data(df):
    df.columns = df.columns.str.strip()
    df = df.dropna(how='all')
    df['RawDescription'] = df['RawDescription'].astype(str).str.strip()
    df['TransactionDate'] = pd.to_datetime(df['TransactionDate'], errors='coerce')
    df['Amount'] = pd.to_numeric(df['Amount'], errors='coerce')
    df['Balance'] = pd.to_numeric(df['Balance'], errors='coerce')
    return df
```

### Step 2: UUID Generation for Primary Keys

Generate deterministic UUIDs for TransactionID:

```python
import uuid

def generate_transaction_id(row):
    seed = f"{row['TransactionDate']}_{row['Amount']}_{row['RawDescription']}"
    return str(uuid.uuid5(uuid.NAMESPACE_DNS, seed))

df['TransactionID'] = df.apply(generate_transaction_id, axis=1)
```

### Step 3: Counterparty Cross-Reference

Match or create CounterpartyIDs based on inferred CounterpartyName:

```python
def resolve_counterparty_id(counterparty_name, counterparty_table):
    existing = counterparty_table[
        counterparty_table['CounterpartyName'] == counterparty_name
    ]
    if not existing.empty:
        return existing.iloc[0]['CounterpartyID']
    new_id = str(uuid.uuid4())
    counterparty_table.loc[len(counterparty_table)] = {
        'CounterpartyID': new_id,
        'CounterpartyName': counterparty_name
    }
    return new_id

df['CounterpartyID'] = df['InferredCounterpartyName'].apply(
    lambda name: resolve_counterparty_id(name, counterparties)
)
```

### Step 4: Financial Aggregation & Balance Validation

Validate running balance continuity and format amount signs:

```python
def validate_balances(df):
    errors = []
    for i in range(1, len(df)):
        expected = df.iloc[i-1]['Balance'] + df.iloc[i]['Amount']
        actual = df.iloc[i]['Balance']
        if abs(expected - actual) > 0.01:
            errors.append({
                'row': i,
                'expected': expected,
                'actual': actual,
                'discrepancy': actual - expected
            })
    return errors

def format_amount_signs(df):
    df.loc[df['Amount'] > 0, 'Amount'] = df['Amount'].round(2)
    df.loc[df['Amount'] < 0, 'Amount'] = df['Amount'].round(2)
    return df
```

### Step 5: Type Validation

Ensure all fields conform to the Data Dictionary:

```python
def validate_types(df):
    assert df['TransactionID'].str.len().eq(36).all(), "TransactionID must be 36 chars"
    assert df['CounterpartyID'].str.len().eq(36).all(), "CounterpartyID must be 36 chars"
    assert df['RawDescription'].str.len().le(255).all(), "RawDescription max 255 chars"
    assert df['Amount'].astype(str).str.extract(r'-?\d+\.\d{1,2}').notna().all().all()
    assert df['Category'].str.len().le(50).all(), "Category max 50 chars"
    assert df['IsRecurring'].isin([0, 1]).all(), "IsRecurring must be 0 or 1"
    return True
```

## Best Practices

- Always use deterministic UUID5 with a consistent namespace seed for reproducibility
- Never use LLM for numerical calculations - use Pandas/Python exclusively
- Preserve `RawDescription` exactly as extracted for audit trail
- Validate running balance continuity before writing to storage
- Check for duplicate TransactionIDs before export

## Common Issues

### Balance Validation Failures

**Issue**: Running balance does not match expected values

**Solution**: Check for missing transactions, date ordering, or Amount sign conventions. Re-extract data if necessary.

### Duplicate CounterpartyIDs

**Issue**: Same counterparty appears with different IDs

**Solution**: Normalize CounterpartyName (strip whitespace, standardize case) before cross-referencing.

### UUID Collisions

**Issue**: Two different transactions generate the same TransactionID

**Solution**: Ensure the seed includes enough unique fields (date + amount + description).

## Verification Commands

```bash
python -c "
import pandas as pd
df = pd.read_csv('transactions.csv')
print(f'Rows: {len(df)}')
print(f'Null TransactionIDs: {df[\"TransactionID\"].isna().sum()}')
print(f'Null CounterpartyIDs: {df[\"CounterpartyID\"].isna().sum()}')
print(f'Duplicate TransactionIDs: {df[\"TransactionID\"].duplicated().sum()}')
"
```
