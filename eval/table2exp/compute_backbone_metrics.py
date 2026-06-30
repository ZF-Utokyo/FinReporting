#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from __future__ import annotations

import argparse
from pathlib import Path
from typing import List, Optional

import pandas as pd


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(description="Compute Table2 backbone metrics.")
    p.add_argument("--template", required=True)
    p.add_argument("--sheet", default="checklist")
    p.add_argument("--out-dir", default="eval/table2exp/reports")
    p.add_argument(
        "--model-keys",
        default="",
        help="Optional comma-separated model keys. If empty, infer from *_value columns.",
    )
    return p.parse_args()


def parse_match(v) -> Optional[int]:
    if pd.isna(v):
        return None
    s = str(v).strip().lower()
    if s in {"1", "y", "yes", "true", "t", "pass", "correct", "ok"}:
        return 1
    if s in {"0", "n", "no", "false", "f", "fail", "wrong", "ng"}:
        return 0
    return None


def parse_bool01(v) -> int:
    if pd.isna(v):
        return 0
    s = str(v).strip().lower()
    if s in {"1", "y", "yes", "true", "t", "on"}:
        return 1
    if s in {"0", "n", "no", "false", "f", "off"}:
        return 0
    try:
        return 1 if int(float(s)) != 0 else 0
    except Exception:
        return 0


def safe_rate(num: int, den: int):
    if den == 0:
        return pd.NA
    return num / den


def fmt(v) -> str:
    if pd.isna(v):
        return "N/A"
    return f"{float(v):.4f}"


def infer_model_keys(df: pd.DataFrame, model_keys_arg: str) -> List[str]:
    if model_keys_arg.strip():
        return [x.strip() for x in model_keys_arg.split(",") if x.strip()]
    keys = []
    for c in df.columns:
        if c.endswith("_value"):
            k = c[: -len("_value")]
            keys.append(k)
    # keep stable order by first appearance
    out = []
    seen = set()
    for k in keys:
        if k not in seen:
            need_cols = [f"{k}_review_required", f"is_match_{k}", f"{k}_estimated_cost_usd"]
            if all(col in df.columns for col in need_cols):
                out.append(k)
                seen.add(k)
    return out


def main() -> None:
    args = parse_args()
    out_dir = Path(args.out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    df = pd.read_excel(Path(args.template), sheet_name=args.sheet)
    total = len(df)
    model_keys = infer_model_keys(df, args.model_keys)
    if not model_keys:
        raise SystemExit("No model keys found. Provide --model-keys or ensure *_value columns exist.")

    rows = []
    for k in model_keys:
        value_col = f"{k}_value"
        review_col = f"{k}_review_required"
        match_col = f"is_match_{k}"
        cost_col = f"{k}_estimated_cost_usd"

        for col in [value_col, review_col, match_col, cost_col]:
            if col not in df.columns:
                raise SystemExit(f"Missing required column for model_key={k}: {col}")

        filled = int(df[value_col].notna().sum())
        fr = safe_rate(filled, total)

        review_required = int(df[review_col].apply(parse_bool01).sum())
        cr = safe_rate(review_required, total)

        match_num = df[match_col].apply(parse_match)
        reviewed = int(match_num.notna().sum())
        correct = int((match_num == 1).sum())
        acc = safe_rate(correct, reviewed)

        cost_total = pd.to_numeric(df[cost_col], errors="coerce").fillna(0.0).sum()
        avg_cost_per_field = cost_total / total if total else 0.0

        rows.append(
            {
                "model_key": k,
                "total_rows": total,
                "filled_rows": filled,
                "FR": fr,
                "review_required_rows": review_required,
                "CR": cr,
                "reviewed_rows": reviewed,
                "correct_rows": correct,
                "Acc": acc,
                "estimated_cost_usd_total": cost_total,
                "estimated_cost_usd_per_field": avg_cost_per_field,
            }
        )

    out_df = pd.DataFrame(rows)
    p_csv = out_dir / "table2_backbone_metrics.csv"
    p_md = out_dir / "table2_backbone_metrics.md"
    out_df.to_csv(p_csv, index=False, encoding="utf-8-sig")

    show = out_df.copy()
    for c in ["FR", "CR", "Acc"]:
        show[c] = show[c].apply(fmt)
    with p_md.open("w", encoding="utf-8") as f:
        f.write("# Table2 Backbone Metrics\n\n")
        f.write(show.to_markdown(index=False))
        f.write("\n")

    print(f"[OK] {p_csv}")
    print(f"[OK] {p_md}")


if __name__ == "__main__":
    main()
