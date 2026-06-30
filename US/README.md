# US XBRL Extraction

This module extracts annual financial statement values from SEC XBRL filings and exports CSV or Excel artifacts.

## Basic Usage

Cash-flow CSV extraction:

```bash
python US/extract_xbrl_cash_flow.py \
  --cik 0000104169 \
  --report-date 2025-01-31 \
  --symbol WMT \
  --out outputs/us_wmt_cash_flow.csv
```

Three-statement Excel export:

```bash
python US/export_three_statements_excel.py \
  --symbol AAPL \
  --cik 0000320193 \
  --out outputs/us_aapl_3statements.xlsx
```

## Workflow

1. Fetch SEC submissions metadata for a CIK.
2. Locate the target 10-K filing by fiscal year end date.
3. Download and parse the XBRL instance document.
4. Map US-GAAP tags to canonical fields.
5. Export structured statement values to CSV or Excel.

## Notes

- Amounts are stored as raw USD values, not millions or thousands.
- SEC requests require a User-Agent. Set `SEC_USER_AGENT` if you want a project-specific contact string.
- Some companies use variant US-GAAP tags; update the tag lists in `export_three_statements_excel.py` or `extract_xbrl_cash_flow.py` if needed.
