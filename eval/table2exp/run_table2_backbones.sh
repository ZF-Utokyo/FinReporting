#!/usr/bin/env bash
set -euo pipefail

# Run from project root:
#   bash eval/table2exp/run_table2_backbones.sh

ROOT_DIR="$(cd "$(dirname "$0")/../.." && pwd)"
cd "$ROOT_DIR"

source ~/.zshrc >/dev/null 2>&1 || true

OUT_ROOT="eval/table2exp/outputs"
REPORT_ROOT="eval/table2exp/reports"
SPLIT_CSV="eval/table2exp/cn_non_fin_table2_10.csv"

mkdir -p "$OUT_ROOT" "$OUT_ROOT/cn" "$REPORT_ROOT"

# 1) Rule pipeline on held-out 10 CN firms (downloads annual PDFs to CN/raw_pdfs)
./venv/bin/python eval/run_batch_pipeline.py \
  --markets cn \
  --list-cn "$SPLIT_CSV" \
  --out-root "$OUT_ROOT"

LATEST_RUN_LOG="$(ls -t "$OUT_ROOT"/run_log_cn_*.csv | head -n 1)"
./venv/bin/python eval/generate_manual_check_template.py \
  --run-log "$LATEST_RUN_LOG" \
  --out "$OUT_ROOT/manual_check_template_table2_base.xlsx"

run_backbone() {
  local provider="$1"
  local model="$2"
  local out_csv="$3"
  local price_in="${4:-}"
  local price_out="${5:-}"

  local extra=()
  if [[ -n "$price_in" && -n "$price_out" ]]; then
    extra+=(--price-input-per-1m "$price_in" --price-output-per-1m "$price_out")
  fi

  ./venv/bin/python llm_cn_verifier.py \
    --provider "$provider" \
    --model "$model" \
    --input-dir "$OUT_ROOT/cn" \
    --out-csv "$out_csv" \
    "${extra[@]}"
}

# Optional pricing env vars (USD / 1M tokens):
# PRICE_GPT52_IN_PER_1M, PRICE_GPT52_OUT_PER_1M, ...
run_backbone openai  "gpt-5.2"                     "$OUT_ROOT/llm_repair_gpt52.csv"      "${PRICE_GPT52_IN_PER_1M:-}"      "${PRICE_GPT52_OUT_PER_1M:-}"
run_backbone gemini  "gemini-2.5-flash"           "$OUT_ROOT/llm_repair_gemini25f.csv"  "${PRICE_GEMINI25F_IN_PER_1M:-}"  "${PRICE_GEMINI25F_OUT_PER_1M:-}"
run_backbone claude  "claude-sonnet-4-20250514"   "$OUT_ROOT/llm_repair_claude_s4.csv"  "${PRICE_CLAUDE_S4_IN_PER_1M:-}"  "${PRICE_CLAUDE_S4_OUT_PER_1M:-}"
run_backbone deepseek "deepseek-chat"              "$OUT_ROOT/llm_repair_deepseek.csv"   "${PRICE_DEEPSEEK_IN_PER_1M:-}"   "${PRICE_DEEPSEEK_OUT_PER_1M:-}"

# 2) Build PDF manifest for human checking
./venv/bin/python eval/table2exp/build_cn_pdf_manifest.py \
  --split-csv "$SPLIT_CSV" \
  --out-dir eval/table2exp/raw_pdfs \
  --manifest eval/table2exp/raw_pdfs/manifest_cn_table2.csv \
  --copy

# 3) Build backbone review template
./venv/bin/python eval/table2exp/generate_backbone_review_template.py \
  --base-template "$OUT_ROOT/manual_check_template_table2_base.xlsx" \
  --gpt52-csv "$OUT_ROOT/llm_repair_gpt52.csv" \
  --gemini25f-csv "$OUT_ROOT/llm_repair_gemini25f.csv" \
  --claude-s4-csv "$OUT_ROOT/llm_repair_claude_s4.csv" \
  --deepseek-csv "$OUT_ROOT/llm_repair_deepseek.csv" \
  --out "$OUT_ROOT/manual_check_template_table2_backbones.xlsx"

# 4) Compute FR/CR now; Acc after filling is_match_* columns
./venv/bin/python eval/table2exp/compute_backbone_metrics.py \
  --template "$OUT_ROOT/manual_check_template_table2_backbones.xlsx" \
  --out-dir "$REPORT_ROOT"

echo "[OK] Table2 experiment artifacts ready:"
echo "  - Outputs: $OUT_ROOT"
echo "  - Human check template: $OUT_ROOT/manual_check_template_table2_backbones.xlsx"
echo "  - Real PDF manifest: eval/table2exp/raw_pdfs/manifest_cn_table2.csv"
echo "  - Metrics: $REPORT_ROOT/table2_backbone_metrics.csv"
