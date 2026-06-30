#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Compute extraction coverage and accuracy from manual check template.

Usage:
  ./venv/bin/python eval/compute_accuracy.py \
    --template eval/manual_check_template.xlsx \
    --out-dir eval/reports
"""

from __future__ import annotations

import argparse
from pathlib import Path
from typing import Optional

import pandas as pd


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(description="Compute accuracy metrics from manual check template.")
    p.add_argument("--template", required=True, help="Manual check template Excel path")
    p.add_argument("--sheet", default="checklist", help="Template sheet name")
    p.add_argument("--out-dir", default="eval/reports", help="Output report directory")
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


def agg_metrics(df: pd.DataFrame, group_cols: list) -> pd.DataFrame:
    g = df.groupby(group_cols, dropna=False)
    out = g.agg(
        total_rows=("field_name", "count"),
        non_null_extracted=("extracted_value", lambda x: int(pd.Series(x).notna().sum())),
        reviewed_rows=("match_num", lambda x: int(pd.Series(x).notna().sum())),
        correct_rows=("match_num", lambda x: int((pd.Series(x) == 1).sum())),
    ).reset_index()
    out["coverage_rate"] = out["non_null_extracted"] / out["total_rows"].where(out["total_rows"] != 0, 1)
    out["accuracy_rate"] = out["correct_rows"] / out["reviewed_rows"].replace({0: pd.NA})
    return out


def main() -> None:
    args = parse_args()
    template = Path(args.template)
    out_dir = Path(args.out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    df = pd.read_excel(template, sheet_name=args.sheet)
    df["match_num"] = df["is_match"].apply(parse_match)

    overall = agg_metrics(df.assign(_all="ALL"), ["_all"])
    by_market = agg_metrics(df, ["market"])
    by_market_stmt = agg_metrics(df, ["market", "statement"])
    by_market_stmt_field = agg_metrics(df, ["market", "statement", "field_name"])

    overall_path = out_dir / "accuracy_overall.csv"
    by_market_path = out_dir / "accuracy_by_market.csv"
    by_market_stmt_path = out_dir / "accuracy_by_market_statement.csv"
    by_field_path = out_dir / "accuracy_by_market_statement_field.csv"

    overall.to_csv(overall_path, index=False, encoding="utf-8-sig")
    by_market.to_csv(by_market_path, index=False, encoding="utf-8-sig")
    by_market_stmt.to_csv(by_market_stmt_path, index=False, encoding="utf-8-sig")
    by_market_stmt_field.to_csv(by_field_path, index=False, encoding="utf-8-sig")

    summary_md = out_dir / "summary.md"
    with summary_md.open("w", encoding="utf-8") as f:
        f.write("# Accuracy Summary\n\n")
        if not overall.empty:
            r = overall.iloc[0]
            acc_v = r["accuracy_rate"]
            acc_s = "N/A" if pd.isna(acc_v) else f"{float(acc_v):.4f}"
            f.write(
                f"- total_rows: {int(r['total_rows'])}\n"
                f"- reviewed_rows: {int(r['reviewed_rows'])}\n"
                f"- correct_rows: {int(r['correct_rows'])}\n"
                f"- coverage_rate: {float(r['coverage_rate']):.4f}\n"
                f"- accuracy_rate: {acc_s}\n"
            )
        f.write("\n## By Market\n\n")
        if not by_market.empty:
            f.write(by_market.to_markdown(index=False))
            f.write("\n")

    print(f"[OK] Reports generated in: {out_dir}")
    print(f"[OK] {overall_path}")
    print(f"[OK] {by_market_path}")
    print(f"[OK] {by_market_stmt_path}")
    print(f"[OK] {by_field_path}")
    print(f"[OK] {summary_md}")


if __name__ == "__main__":
    main()
