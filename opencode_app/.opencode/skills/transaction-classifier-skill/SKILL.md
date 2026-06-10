---
name: transaction-classifier-skill
description: Integrates LLM capabilities to classify bank raw descriptions into distinct categories and map them to standard counterparties, outputting structured JSON. Evaluates patterns to flag recurring items.
license: Apache-2.0
compatibility: opencode
metadata:
  audience: developers
  workflow: classification
---

## What I do

- Classify bank transaction RawDescription into standardized Categories (e.g., F&B, Transport, Utilities, Salary)
- Infer and normalize CounterpartyName from raw bank descriptions
- Output structured JSON with classification results for each transaction
- Evaluate transaction patterns to flag recurring items (IsRecurring flag)
- Support batch classification of multiple transactions for efficiency
- Maintain a consistent category taxonomy and counterparty name mapping

## When to use me

Use this skill when:
- Raw bank transaction descriptions need to be categorized into industry/expense categories
- Counterparty names need to be normalized from raw bank statement text
- Determining if a transaction is a recurring deduction (subscriptions, monthly bills)
- Producing structured JSON output for downstream relational processing
- Building or updating a Counterparty master list from new transactions

Do NOT use for:
- Extracting data from PDF files (use `pdf-specialist-skill`)
- Data cleaning, UUID generation, or numerical validation (use `relational-data-processor-skill`)
- Writing results to database storage (use `db-storage-specialist-skill`)

## Prerequisites

- Access to LLM via OpenCode (GLM series models)
- Python 3+ with `json` module (built-in)
- Input data must have `RawDescription` field populated

## Steps

### Step 1: Prepare Classification Prompt

Structure the raw descriptions for LLM classification:

```python
import json

def build_classification_prompt(transactions):
    categories = [
        "F&B", "Transport", "Utilities", "Telecommunications",
        "Salary", "Rental", "Insurance", "Government", "Bank Fees",
        "Transfer", "Investment", "Subscription", "Other"
    ]
    prompt = f"""Classify each bank transaction description into a category and infer the standardized counterparty name.

Categories: {json.dumps(categories)}

Rules:
- CounterpartyName must be the official or most recognizable name
- Use "Other" only when no category fits
- Flag recurring items (subscriptions, monthly bills, regular deductions)

Transactions:
{json.dumps(transactions, indent=2)}

Output JSON array with fields: index, Category, CounterpartyName, IsRecurring
"""
    return prompt
```

### Step 2: Invoke LLM Classification

Send the prompt to the LLM and parse the structured response:

```python
def parse_classification_response(response_text):
    try:
        results = json.loads(response_text)
        required_fields = {"index", "Category", "CounterpartyName", "IsRecurring"}
        for item in results:
            if not required_fields.issubset(item.keys()):
                raise ValueError(f"Missing fields in: {item}")
            item["IsRecurring"] = 1 if item["IsRecurring"] else 0
        return results
    except json.JSONDecodeError:
        raise ValueError("LLM response is not valid JSON")
```

### Step 3: Validate Classification Results

Ensure all classifications conform to the expected schema:

```python
def validate_classifications(results, valid_categories):
    errors = []
    for item in results:
        if item["Category"] not in valid_categories:
            errors.append(f"Invalid category '{item['Category']}' at index {item['index']}")
        if not item["CounterpartyName"] or len(item["CounterpartyName"]) > 100:
            errors.append(f"Invalid CounterpartyName at index {item['index']}")
        if item["IsRecurring"] not in (0, 1):
            errors.append(f"Invalid IsRecurring at index {item['index']}")
    return errors
```

### Step 4: Recurring Pattern Detection

Evaluate transaction patterns to identify recurring items:

```python
def detect_recurring_patterns(transactions_df):
    recurring = []
    grouped = transactions_df.groupby('CounterpartyName')
    for name, group in grouped:
        if len(group) >= 2:
            amounts = group['Amount'].abs()
            if amounts.std() < 1.0:
                recurring.extend(group.index.tolist())
    return recurring
```

### Step 5: Merge Results with Source Data

Combine classification results back into the transaction DataFrame:

```python
def merge_classifications(df, classifications):
    class_df = pd.DataFrame(classifications)
    class_df = class_df.set_index('index')
    df = df.join(class_df[['Category', 'CounterpartyName', 'IsRecurring']])
    return df
```

## Best Practices

- Always provide the full category taxonomy to the LLM for consistent classification
- Validate LLM output against the category list - reject or retry invalid categories
- Use deterministic recurring detection (amount + counterparty frequency) as a secondary check
- Normalize CounterpartyName to title case for consistency
- Log any transactions classified as "Other" for manual review

## Common Issues

### Inconsistent Counterparty Names

**Issue**: Same entity classified with different name variations (e.g., "Grab", "GRAB PTE LTD", "GrabPay")

**Solution**: Maintain a mapping dictionary of known aliases. Normalize LLM output against this mapping before saving.

### Invalid JSON from LLM

**Issue**: LLM response contains markdown formatting or extra text

**Solution**: Strip markdown fences and extract JSON array from the response before parsing.

### False Recurring Flags

**Issue**: One-time transactions flagged as recurring

**Solution**: Require minimum 2 occurrences within a defined time window AND similar amounts before flagging.

## Verification Commands

```bash
python -c "
import json
results = json.load(open('classifications.json'))
print(f'Classified transactions: {len(results)}')
categories = set(r['Category'] for r in results)
print(f'Categories used: {categories}')
recurring = sum(1 for r in results if r['IsRecurring'] == 1)
print(f'Recurring items: {recurring}')
"
```
