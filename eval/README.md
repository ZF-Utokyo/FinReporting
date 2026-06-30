# Evaluation Plan (Non-Financial, FY)

This folder defines a practical evaluation baseline for FinReporting.

## Scope
- Markets: `US`, `JP`, `CN`
- Company count: `20` per market
- Report type: annual report / FY only
- Exclusions: banks, insurance, securities brokers, diversified financials

## Files
- `eval/us_non_fin_20.csv`
- `eval/jp_non_fin_20.csv`
- `eval/cn_non_fin_20.csv`
- `eval/run_batch_pipeline.py`
- `eval/generate_manual_check_template.py`
- `eval/compute_accuracy.py`
- `llm_cn_verifier.py`
- `eval/run_llm_only_cn.py`
- `eval/run_llm_only_cn_doc.py`
- `eval/generate_llm_review_template.py`
- `eval/compute_llm_ablation_metrics.py`
- `eval/generate_four_way_review_template.py`
- `eval/compute_four_way_ablation_metrics.py`
- `eval/extra10_cn_challenge/*`

## Manual check workbook naming
- Use: `check_{market}_{symbol}.xlsx`
- Examples:
  - `check_cn_300750.xlsx`
  - `check_us_aapl.xlsx`
  - `check_jp_7203.xlsx`

## Suggested workflow
1. Batch run extraction + check workbook generation:
   - `./venv/bin/python eval/run_batch_pipeline.py --markets us,jp,cn`
2. Generate manual check template from run log:
   - `./venv/bin/python eval/generate_manual_check_template.py --run-log eval/outputs/run_log_*.csv --out eval/manual_check_template.xlsx`
3. Human verifies key fields against filing in `checklist` sheet.
4. Compute coverage / accuracy:
   - `./venv/bin/python eval/compute_accuracy.py --template eval/manual_check_template.xlsx --out-dir eval/reports`
5. Patch mapping and anomaly rules, then rerun.

## LLM Verify/Repair Workflow (Deterministic-first)
Goal: keep rule pipeline as primary output and add a traceable LLM verification layer.

### Guardrails
1. Rule-first: baseline value is always `rule_value`.
2. No evidence, no repair: LLM `REPAIR` is rejected unless evidence references valid candidate codes and supports the proposed value.
3. Fixed generation settings for reproducibility:
   - fixed model id (default `gpt-4o-2024-11-20`)
   - temperature `0`
   - fixed system prompt + JSON schema

### 1) Run LLM verify/repair on CN outputs
Set API key in env first:

```bash
export OPENAI_API_KEY=
```

Run verifier (core fields):

```bash
./venv/bin/python llm_cn_verifier.py \
  --input-dir eval/outputs/cn \
  --out-csv eval/outputs/llm_cn_audit.csv \
  --append-sheet
```

Outputs:
- CSV: `eval/outputs/llm_cn_audit.csv`
- Per workbook sheet: `LLM_AUDIT_CN`
- Key trace columns: `rule_value`, `llm_decision`, `evidence_json`, `final_value`, `final_source`

### 2) Build ablation review template
Merge baseline checklist with LLM outputs:

```bash
./venv/bin/python eval/generate_llm_review_template.py \
  --base-template eval/manual_check_template.xlsx \
  --llm-audit eval/outputs/llm_cn_audit.csv \
  --out eval/manual_check_template_llm.xlsx
```

Human labels:
- `is_match_rule`: correctness of rule baseline
- `is_match_final`: correctness after LLM verify/repair

### 3) Compute ablation metrics
```bash
./venv/bin/python eval/compute_llm_ablation_metrics.py \
  --template eval/manual_check_template_llm.xlsx \
  --out-dir eval/reports_llm
```

Main metrics:
- `accuracy_rule` vs `accuracy_final`
- `workload_rate` and `workload_reduction_rate` (based on `review_required_recommended`)
- `false_repair_rate` (rows with `repair_applied=1` but `is_match_final=0`)

## Four-way Ablation (Recommended for ACL)
Systems:
1. Rule-only
2. Rule + LLM-verify (`--verify-only`)
3. Rule + LLM-verify/repair (guardrailed)
4. LLM-only (recommended: document-level `eval/run_llm_only_cn_doc.py`)

### Build four-way checklist
```bash
./venv/bin/python eval/generate_four_way_review_template.py \
  --base-template eval/manual_check_template.xlsx \
  --verify-csv eval/outputs/llm_verify_cn.csv \
  --repair-csv eval/outputs/llm_repair_cn.csv \
  --llm-only-csv eval/outputs/llm_only_cn_doc.csv \
  --out eval/manual_check_template_four_way.xlsx
```

### Compute four-way metrics
```bash
./venv/bin/python eval/compute_four_way_ablation_metrics.py \
  --template eval/manual_check_template_four_way.xlsx \
  --out-dir eval/reports_four_way
```

## Extra10 Challenge Split
Use `eval/extra10_cn_challenge/` for the additional 10-company CN split.

Quick run:
```bash
bash eval/extra10_cn_challenge/run_four_way.sh
```

## Recommended first-pass check fields
- IS: revenue, operating income, pretax income, net income, EPS
- BS: total assets, total liabilities, equity, cash, receivables, inventories
- CF: CFO, CFI, CFF, net change in cash, ending cash
