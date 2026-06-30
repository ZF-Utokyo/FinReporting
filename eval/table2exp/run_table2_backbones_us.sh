#!/usr/bin/env bash
set -euo pipefail

# Run from project root:
#   bash eval/table2exp/run_table2_backbones_us.sh

ROOT_DIR="$(cd "$(dirname "$0")/../.." && pwd)"
cd "$ROOT_DIR"

source ~/.zshrc >/dev/null 2>&1 || true

OUT_ROOT="eval/table2exp/outputs_us"
REPORT_ROOT="eval/table2exp/reports_us"
SPLIT_CSV="eval/table2exp/us_non_fin_table2_10.csv"
RAW_PDF_DIR="eval/table2exp/raw_pdfs_us"

mkdir -p "$OUT_ROOT" "$OUT_ROOT/us" "$REPORT_ROOT" "$RAW_PDF_DIR"

# 1) Rule pipeline on held-out 10 US firms
./venv/bin/python eval/run_batch_pipeline.py \
  --markets us \
  --list-us "$SPLIT_CSV" \
  --out-root "$OUT_ROOT"

# 2) Download real 10-K PDFs for manual checking
./venv/bin/python US/download_raw_pdfs.py \
  --input-csv "$SPLIT_CSV" \
  --out-dir "$RAW_PDF_DIR" \
  --manifest "$RAW_PDF_DIR/manifest_us_table2.csv"

LATEST_RUN_LOG="$(ls -t "$OUT_ROOT"/run_log_us_*.csv | head -n 1)"
./venv/bin/python eval/generate_manual_check_template.py \
  --run-log "$LATEST_RUN_LOG" \
  --out "$OUT_ROOT/manual_check_template_table2_base.xlsx"

run_backbone() {
  local key="$1"
  local provider="$2"
  local model="$3"
  local out_csv="$4"
  local price_in="${5:-}"
  local price_out="${6:-}"

  local key_env=""
  case "$provider" in
    openai) key_env="${OPENAI_API_KEY:-}" ;;
    gemini) key_env="${GEMINI_API_KEY:-}" ;;
    deepseek) key_env="${DEEPSEEK_API_KEY:-}" ;;
    claude) key_env="${ANTHROPIC_API_KEY:-}" ;;
    *) key_env="" ;;
  esac
  if [[ -z "$key_env" ]]; then
    echo "[WARN] Skip $key ($provider/$model): missing API key env for provider=$provider"
    return 0
  fi

  local -a extra
  extra=()
  if [[ -n "$price_in" && -n "$price_out" ]]; then
    extra+=(--price-input-per-1m "$price_in" --price-output-per-1m "$price_out")
  fi

  if [[ ${#extra[@]} -gt 0 ]]; then
    ./venv/bin/python eval/table2exp/llm_us_verifier.py \
      --provider "$provider" \
      --model "$model" \
      --input-dir "$OUT_ROOT/us" \
      --out-csv "$out_csv" \
      "${extra[@]}"
  else
    ./venv/bin/python eval/table2exp/llm_us_verifier.py \
      --provider "$provider" \
      --model "$model" \
      --input-dir "$OUT_ROOT/us" \
      --out-csv "$out_csv"
  fi

  MODEL_CSV_ARGS+=(--model-csv "${key}=${out_csv}")
  MODEL_KEYS_DONE+=("$key")
}

declare -a MODEL_CSV_ARGS=()
declare -a MODEL_KEYS_DONE=()

# Optional pricing env vars (USD / 1M tokens):
# PRICE_GPT52_IN_PER_1M, PRICE_GPT52_OUT_PER_1M, ...
# PRICE_GPT5MINI_IN_PER_1M, PRICE_GPT5MINI_OUT_PER_1M, ...
# PRICE_GPT4O_IN_PER_1M, PRICE_GPT4O_OUT_PER_1M, ...
# PRICE_GEMINI25PRO_IN_PER_1M, PRICE_GEMINI25PRO_OUT_PER_1M, ...
# PRICE_GEMINI25F_IN_PER_1M, PRICE_GEMINI25F_OUT_PER_1M, ...
# PRICE_GEMINI25FL_IN_PER_1M, PRICE_GEMINI25FL_OUT_PER_1M, ...
# PRICE_DEEPSEEK_IN_PER_1M, PRICE_DEEPSEEK_OUT_PER_1M, ...

# Strong / mid / weak models to widen performance gap
run_backbone gpt52        openai   "gpt-5.2"                 "$OUT_ROOT/llm_repair_gpt52.csv"         "${PRICE_GPT52_IN_PER_1M:-}"         "${PRICE_GPT52_OUT_PER_1M:-}"
run_backbone gpt5mini     openai   "gpt-5-mini"              "$OUT_ROOT/llm_repair_gpt5mini.csv"      "${PRICE_GPT5MINI_IN_PER_1M:-}"      "${PRICE_GPT5MINI_OUT_PER_1M:-}"
run_backbone gpt4o        openai   "gpt-4o-2024-11-20"       "$OUT_ROOT/llm_repair_gpt4o.csv"         "${PRICE_GPT4O_IN_PER_1M:-}"         "${PRICE_GPT4O_OUT_PER_1M:-}"
run_backbone gemini25pro  gemini   "gemini-2.5-pro"          "$OUT_ROOT/llm_repair_gemini25pro.csv"   "${PRICE_GEMINI25PRO_IN_PER_1M:-}"   "${PRICE_GEMINI25PRO_OUT_PER_1M:-}"
run_backbone gemini25f    gemini   "gemini-2.5-flash"        "$OUT_ROOT/llm_repair_gemini25f.csv"     "${PRICE_GEMINI25F_IN_PER_1M:-}"     "${PRICE_GEMINI25F_OUT_PER_1M:-}"
run_backbone gemini25fl   gemini   "gemini-2.5-flash-lite"   "$OUT_ROOT/llm_repair_gemini25fl.csv"    "${PRICE_GEMINI25FL_IN_PER_1M:-}"    "${PRICE_GEMINI25FL_OUT_PER_1M:-}"
run_backbone deepseek     deepseek "deepseek-chat"           "$OUT_ROOT/llm_repair_deepseek.csv"      "${PRICE_DEEPSEEK_IN_PER_1M:-}"      "${PRICE_DEEPSEEK_OUT_PER_1M:-}"

if [[ ${#MODEL_CSV_ARGS[@]} -eq 0 ]]; then
  echo "[ERROR] No backbone model ran successfully (no provider API keys found)."
  exit 1
fi

MODEL_ORDER_CSV="$(IFS=,; echo "${MODEL_KEYS_DONE[*]}")"

# 3) Build backbone review template
./venv/bin/python eval/table2exp/generate_backbone_review_template.py \
  --base-template "$OUT_ROOT/manual_check_template_table2_base.xlsx" \
  "${MODEL_CSV_ARGS[@]}" \
  --model-order "$MODEL_ORDER_CSV" \
  --out "$OUT_ROOT/manual_check_template_table2_backbones.xlsx"

# 4) Compute FR/CR now; Acc after filling is_match_* columns
./venv/bin/python eval/table2exp/compute_backbone_metrics.py \
  --template "$OUT_ROOT/manual_check_template_table2_backbones.xlsx" \
  --out-dir "$REPORT_ROOT"

echo "[OK] US Table2 experiment artifacts ready:"
echo "  - Outputs: $OUT_ROOT"
echo "  - Human check template: $OUT_ROOT/manual_check_template_table2_backbones.xlsx"
echo "  - Real 10-K PDFs + manifest: $RAW_PDF_DIR"
echo "  - Metrics: $REPORT_ROOT/table2_backbone_metrics.csv"
