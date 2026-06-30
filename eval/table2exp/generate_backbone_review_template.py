#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from __future__ import annotations

import argparse
import re
from pathlib import Path
from typing import Dict, List, Tuple

import pandas as pd


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(description="Generate manual review template for backbone comparison.")
    p.add_argument("--base-template", required=True)
    p.add_argument("--sheet", default="checklist")
    p.add_argument(
        "--model-csv",
        action="append",
        default=[],
        help="Model CSV in key=path form (repeatable), e.g. --model-csv gpt52=eval/.../llm_repair_gpt52.csv",
    )
    p.add_argument(
        "--model-order",
        default="",
        help="Optional comma-separated model key order for output columns, e.g. gpt52,gemini25f,deepseek",
    )
    p.add_argument("--out", default="eval/table2exp/outputs/manual_check_template_table2_backbones.xlsx")
    return p.parse_args()


def normalize_symbol(market: str, symbol_raw: str) -> str:
    s = (symbol_raw or "").strip()
    market = (market or "").strip().lower()
    if market == "cn":
        d = "".join(ch for ch in s if ch.isdigit())
        return d.zfill(6) if d else s
    if market == "jp":
        d = "".join(ch for ch in s if ch.isdigit())
        return d.zfill(4) if d else s
    return s.upper()


def build_idx(df: pd.DataFrame) -> Dict[Tuple[str, str, str, str], dict]:
    out: Dict[Tuple[str, str, str, str], dict] = {}
    for _, r in df.iterrows():
        market = str(r.get("market", "")).strip().lower()
        symbol = normalize_symbol(market, str(r.get("symbol", "")))
        statement = str(r.get("statement", "")).strip().upper()
        field_name = str(r.get("field_name", "")).strip()
        if not (market and symbol and statement and field_name):
            continue
        out[(market, symbol, statement, field_name)] = r.to_dict()
    return out


def parse_model_csv_pairs(items: List[str]) -> List[Tuple[str, Path]]:
    out: List[Tuple[str, Path]] = []
    key_pat = re.compile(r"^[a-z0-9_]+$")
    for raw in items:
        s = str(raw or "").strip()
        if not s or "=" not in s:
            raise ValueError(f"Invalid --model-csv entry: {raw!r}. Expected key=path.")
        key, path = s.split("=", 1)
        key = key.strip()
        path = path.strip()
        if not key_pat.fullmatch(key):
            raise ValueError(f"Invalid model key: {key!r}. Use lowercase letters/numbers/underscore only.")
        if not path:
            raise ValueError(f"Empty path for model key: {key!r}")
        out.append((key, Path(path)))
    if not out:
        raise ValueError("At least one --model-csv key=path is required.")
    return out


def reorder_model_keys(keys: List[str], model_order: str) -> List[str]:
    if not model_order.strip():
        return keys
    order = [x.strip() for x in model_order.split(",") if x.strip()]
    seen = set()
    out: List[str] = []
    for k in order:
        if k in keys and k not in seen:
            out.append(k)
            seen.add(k)
    for k in keys:
        if k not in seen:
            out.append(k)
            seen.add(k)
    return out


def main() -> None:
    args = parse_args()
    pairs = parse_model_csv_pairs(args.model_csv)

    base_df = pd.read_excel(Path(args.base_template), sheet_name=args.sheet)
    model_df = {k: pd.read_csv(p, dtype=str).fillna("") for k, p in pairs}
    model_keys = reorder_model_keys(list(model_df.keys()), args.model_order)
    model_idx = {k: build_idx(v) for k, v in model_df.items()}

    rows = []
    for _, r in base_df.iterrows():
        market = str(r.get("market", "")).strip().lower()
        symbol = normalize_symbol(market, str(r.get("symbol", "")))
        statement = str(r.get("statement", "")).strip().upper()
        field_name = str(r.get("field_name", "")).strip()
        key = (market, symbol, statement, field_name)

        out = dict(r.to_dict())
        out["symbol"] = symbol
        out["rule_value"] = r.get("extracted_value")

        for model_key in model_keys:
            rec = model_idx[model_key].get(key, {})
            out[f"{model_key}_value"] = rec.get("final_value", "")
            out[f"{model_key}_decision"] = rec.get("llm_decision", "")
            out[f"{model_key}_reason"] = rec.get("reason", "")
            out[f"{model_key}_evidence"] = rec.get("evidence_json", "")
            out[f"{model_key}_review_required"] = rec.get("review_required_recommended", "1")
            out[f"{model_key}_provider"] = rec.get("provider", "")
            out[f"{model_key}_model"] = rec.get("model", "")
            out[f"{model_key}_prompt_tokens"] = rec.get("prompt_tokens", "")
            out[f"{model_key}_completion_tokens"] = rec.get("completion_tokens", "")
            out[f"{model_key}_total_tokens"] = rec.get("total_tokens", "")
            out[f"{model_key}_estimated_cost_usd"] = rec.get("estimated_cost_usd", "")
            out[f"is_match_{model_key}"] = ""

        rows.append(out)

    out_df = pd.DataFrame(rows)
    instructions = pd.DataFrame(
        [
            ["How to label", "Fill expected_value, then set is_match_* as 1/0 (or yes/no)."],
            ["FR", "Filled Rate = non-null {model}_value / total rows."],
            ["CR", "Conflict Rate = {model}_review_required=1 / total rows."],
            ["Acc", "Accuracy = correct / reviewed rows from is_match_* labels."],
            ["Cost", "Use {model}_estimated_cost_usd sum (if pricing args were provided while running)."],
        ],
        columns=["Item", "Instruction"],
    )

    out_path = Path(args.out)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        out_df.to_excel(writer, index=False, sheet_name="checklist")
        instructions.to_excel(writer, index=False, sheet_name="instructions")

    print(f"[OK] Backbone review template generated: {out_path}")
    print(f"[INFO] Rows: {len(out_df)}")


if __name__ == "__main__":
    main()
