# US Quickstart

## Install

```bash
pip install -r requirements.txt
```

## Single Company

```bash
python US/export_three_statements_excel.py \
  --symbol AAPL \
  --cik 0000320193 \
  --out outputs/us_aapl_3statements.xlsx
```

## Cash-Flow CSV

```bash
python US/extract_xbrl_cash_flow.py \
  --cik 0000104169 \
  --report-date 2025-01-31 \
  --symbol WMT \
  --out outputs/us_wmt_cash_flow.csv
```

## Batch Cash-Flow CSV

Prepare a CSV:

```csv
cik,report_date,symbol,form_type
0000104169,2025-01-31,WMT,10-K
0000320193,2024-09-28,AAPL,10-K
```

Run:

```bash
python US/batch_extract.py \
  --input examples/us_companies.csv \
  --out outputs/us_cash_flow_batch.csv \
  --continue-on-error
```

## Finding CIKs

Use the SEC company search page or the SEC company tickers JSON. The CIK should be zero-padded to ten digits.

