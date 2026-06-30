# FinReporting

Code release for an ACL Demo accepted project on automated financial statement extraction and LLM-assisted verification across CN, US, and JP filings.

This repository is intentionally data-light. It includes source code, schemas, evaluation utilities, and toy input examples. It does not include API keys, raw annual-report caches, generated model outputs, or human-checked Excel files.

## Contents

- `CN/`: CN annual-report PDF extraction and canonical three-statement export.
- `US/`: SEC/XBRL extraction and canonical three-statement export.
- `JP/`: EDINET/XBRL extraction and canonical three-statement export.
- `eval/`: scripts for rule-only, LLM-only, verify, repair, and ablation workflows.
- `schemas/`: public schema files required by the CN extraction pipeline.
- `examples/`: small toy manifests showing input formats.

## Install

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

Some PDF table extraction workflows use Camelot. Depending on your OS, Camelot may require system packages such as Ghostscript.

## Start Here

Run a no-data smoke test first:

```bash
python run_smoke_test.py
```

Then run one market with your own public filing data:

```bash
python run_example.py cn \
  --symbol 300750 \
  --pdf /path/to/annual_report.pdf
```

```bash
python run_example.py us \
  --symbol AAPL \
  --cik 0000320193
```

```bash
python run_example.py jp \
  --symbol 7203 \
  --company-name "Toyota" \
  --xbrl-zip /path/to/edinet_type1.zip
```

## Environment Variables

Copy `.env.example` to `.env` for local use, then fill in only the keys you need. Do not commit `.env`.

```bash
cp .env.example .env
```

The code can run rule-only extraction without LLM provider keys. LLM verification and repair require a provider API key.

## Quick Examples

The unified runner above wraps the underlying scripts. You can also call each script directly.

CN from a local annual-report PDF:

```bash
python CN/export_three_statements_excel_cn.py \
  --symbol 300750 \
  --pdf /path/to/annual_report.pdf \
  --schema-file schemas/CN_Schemas.xlsx \
  --out outputs/cn_300750_3statements.xlsx
```

US from SEC XBRL:

```bash
python US/export_three_statements_excel.py \
  --symbol AAPL \
  --cik 0000320193 \
  --out outputs/us_aapl_3statements.xlsx
```

JP from a local EDINET ZIP:

```bash
python JP/export_three_statements_excel_jp.py \
  --xbrl-zip /path/to/edinet_type1.zip \
  --symbol 7203 \
  --company-name "Toyota" \
  --out outputs/jp_7203_3statements.xlsx
```

## What Is Not Included

- No API keys or private environment files.
- No raw PDF/HTML/ZIP filing caches.
- No generated Excel outputs.
- No human-checked `check_*.xlsx` or `manual_check_*.xlsx` files.
- No local virtual environment, build artifacts, or UI demo files.

Use the download scripts and examples to recreate the data locally from public sources.
