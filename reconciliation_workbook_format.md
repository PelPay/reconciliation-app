# Reconciliation System

## Overview

A Python system that reconciles **Pelpay** transaction data against **gateway settlement files** (Cybersource, ChoicePay/MPGS) and produces a structured Excel workbook.

## Components

| File | Purpose |
|------|---------|
| `reconcile_core.py` | Core reconciliation engine — importable, no I/O globals. Public API: `run(pelpay_path, settlement_files, settlement_date, output_path)` |
| `app.py` | Streamlit web app — file upload, run reconciliation, download result |
| `reconciliation_workbook_format.md` | This file — system docs |

## Core Module (`reconcile_core.py`)

### Public API

```python
from reconcile_core import run, DEFAULT_SCHEMAS

result = run(
    pelpay_path="path/to/pelpay.xlsx",
    settlement_files=[
        ("CYBERSOURCE", "USD", "path/to/cybersource_usd.xlsx"),
        ("CYBERSOURCE", "NGN", "path/to/cybersource_ngn.xlsx"),
        ("CHOICEPAY", "NGN", "path/to/mpgs_ngn.xlsx"),
        ("CHOICEPAY", "USD", "path/to/mpgs_usd.xlsx"),
    ],
    settlement_date=date(2026, 7, 6),
    output_path="output.xlsx",
)

print(result)
# {'output_path': '...', 'sheets': [...], 'settle_rows': 194, 'matched': 194, 'exceptions': 98}
```

### Parameters

- `pelpay_path` — path to the Pelpay combined .xlsx (all merchants, all currencies)
- `settlement_files` — list of `(gateway_name, currency, file_path)` tuples
- `settlement_date` — `datetime.date` for a single day, or `(start_date, end_date)` tuple for a range
- `output_path` — where to save the result .xlsx
- `schemas` — optional dict of gateway schemas (defaults to `DEFAULT_SCHEMAS`)

### Custom Schemas

```python
from reconcile_core import run

custom_schemas = {
    'NEW_GATEWAY': {
        'ref': 'Transaction ID',
        'amount': 'Amount',
        'merchant_field': 'Merchant ID',
        'currency': 'Currency',
        'status': 'Status',
        'approved_statuses': {'SETTLED'},
        'date': 'batch_date',  # optional; used for daily grouping
    },
}

run(..., schemas=custom_schemas)
```

## Streamlit App (`app.py`)

### Usage

```bash
python -m streamlit run app.py
```

The app opens at `http://localhost:8501`.

### Workflow

1. **Set date range** (From → To) in sidebar
2. **Upload all .xlsx files** — Pelpay + settlement files all at once; the app auto-detects each file's role
3. Review the detection table
4. Click **Run Reconciliation**
5. Download the output workbook

## Output Workbook

The workbook contains up to 5 sheet types:

### 1. SUMMARY
| Col | Header | Description |
|-----|--------|-------------|
| A | Settlement Date | ISO date (batch_date for Cybersource, Order/Transaction Date for others) |
| B | Currency | NGN / USD |
| C | Settlement Rows | Approved settlement row count |
| D | Matched Rows | Matched to Pelpay |
| E | Settlement Total | Sum of settlement amounts |
| F | Matched Pelpay Total | Sum of matched Pelpay amounts |
| G | Difference | Settlement − Pelpay |
| H | Missing (Pelpay) | Successful Pelpay txs not in settlement |
| I | Missing Amount (Pelpay) | Sum of missing amounts |
| J | Coverage | PARTIAL (default) |
| K | Gateway | CYBERSOURCE / CHOICEPAY |

**Layout:** Title → date → merchant sections with per-day rows → section totals → grand totals → Pelpay transaction summary section.

### 2. All received Transactions
Daily blocks of successful Pelpay transactions grouped by merchant → currency.

**Layout per block:** Block title → subtotal → Pelpay field headers → data rows → subtotal.

Pelpay fields (columns B–O): Transaction Date, Payment Reference, Advice Id, Merchant Name, Merchant code, Merchant Reference, Processor Reference, Currency, Channel, Transaction Status, Gross Amount, Amount Collected, Processing Fee Applied, Merchant Settlement.

### 3. Missing from Settlement
Successful Pelpay transactions not found in any settlement file, grouped by merchant → currency.

**Layout per block:** Block title → summary (count + amount) → field headers → data rows → summary.

Fields: received_date, pelpay_processor_reference, pelpay_transaction_date, pelpay_transaction_amount, pelpay_status, pelpay_merchant_reference, pelpay_payment_reference.

### 4. Missing from Pelpay
Settlement rows not found in Pelpay, grouped by merchant → currency.

**Layout per block:** Block title → summary (count + amount) → field headers → data rows → summary.

Fields: settlement_ref, settlement_date, settlement_amount, settlement_currency, merchant_id, gateway, merchant_name.

### 5. Settlement Sheets (one per merchant/currency/gateway)

### 5. Settlement Sheets (one per merchant/currency/gateway)
Matched settlement rows with appended Pelpay columns.

**Layout:** Title → totals row → original settlement headers + appended columns → data rows → totals row.

Appended: pelpay_transaction_date, pelpay_transaction_amount, pelpay_status, pelpay_merchant_reference, pelpay_payment_reference, settlement_minus_pelpay_difference.

## Color Scheme

| Element | Hex |
|---------|-----|
| Sheet titles | bg: 17365D, text: white |
| Column headers | bg: 1F4E78, text: white |
| Total rows | bg: D9EAF7, text: black |
| Matched rows | bg: E2F0D9 |
| Missing/unmatched | bg: FCE4D6 |
| Exception headers | bg: F4B183 |

## Amount Formatting

- All monetary values: `#,##0.00`
- Datetimes: `yyyy-mm-dd hh:mm:ss`
- Counts: General format

## How It Works

1. **Load Pelpay** — reads the combined Pelpay file, indexes transactions by processor reference
2. **Load settlements** — reads each gateway file, filters approved statuses, groups by merchant
3. **Match** — for each settlement row, looks up the matching Pelpay row by processor reference
4. **Missing from Settlement** — successful Pelpay transactions with no corresponding settlement row
5. **Missing from Pelpay** — settlement rows with no matching Pelpay record
6. **Build workbook** — creates all sheet types with styling and formatting

## Dependencies

- Python 3.10+
- `openpyxl`
- `streamlit` (only for the app)

```bash
pip install openpyxl streamlit
```
