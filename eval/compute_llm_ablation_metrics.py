#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Compute Rule-only vs Rule+LLM ablation metrics.

Usage:
  ./venv/bin/python eval/compute_llm_ablation_metrics.py \
    --template eval/manual_check_template_llm.xlsx \
    --out-dir eval/reports_llm
"""

from __future__ import annotations

import argparse
from pathlib import Path
from typing import Optional

import pandas as pd


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(description="Compute ablation metrics for LLM verify/repair.")
    p.add_argument("--template", required=True, help="LLM review template path")
    p.add_argument("--sheet", default="checklist", help="Checklist sheet name")
    p.add_argument("--out-dir", default="eval/reports_llm", help="Output report directory")
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


def fmt_rate(v) -> str:
    if pd.isna(v):
        return "N/A"
    return f"{float(v):.4f}"


def build_summary(df: pd.DataFrame) -> pd.DataFrame:
    total = len(df)
    reviewed_rule = int(df["match_rule_num"].notna().sum())
    correct_rule = int((df["match_rule_num"] == 1).sum())
    reviewed_final = int(df["match_final_num"].notna().sum())
    correct_final = int((df["match_final_num"] == 1).sum())

    review_required = int(df["review_required_num"].sum())
    workload_rate = safe_rate(review_required, total)
    workload_reduction = pd.NA if pd.isna(workload_rate) else (1.0 - workload_rate)

    repaired = df["repair_applied_num"] == 1
    repaired_reviewed = repaired & df["match_final_num"].notna()
    repaired_reviewed_count = int(repaired_reviewed.sum())
    false_repair_count = int((repaired_reviewed & (df["match_final_num"] == 0)).sum())

    return pd.DataFrame(
        [
            {
                "total_rows": total,
                "reviewed_rule": reviewed_rule,
                "correct_rule": correct_rule,
                "accuracy_rule": safe_rate(correct_rule, reviewed_rule),
                "reviewed_final": reviewed_final,
                "correct_final": correct_final,
                "accuracy_final": safe_rate(correct_final, reviewed_final),
                "accuracy_delta": (
                    pd.NA
                    if pd.isna(safe_rate(correct_rule, reviewed_rule)) or pd.isna(safe_rate(correct_final, reviewed_final))
                    else safe_rate(correct_final, reviewed_final) - safe_rate(correct_rule, reviewed_rule)
                ),
                "review_required_rows": review_required,
                "workload_rate": workload_rate,
                "workload_reduction_rate": workload_reduction,
                "repair_applied_rows": int(repaired.sum()),
                "repair_reviewed_rows": repaired_reviewed_count,
                "false_repair_rows": false_repair_count,
                "false_repair_rate": safe_rate(false_repair_count, repaired_reviewed_count),
            }
        ]
    )


def by_market_summary(df: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for market, g in df.groupby("market", dropna=False):
        s = build_summary(g).iloc[0].to_dict()
        s["market"] = market
        rows.append(s)
    out = pd.DataFrame(rows)
    cols = ["market"] + [c for c in out.columns if c != "market"]
    return out[cols]


def false_repair_details(df: pd.DataFrame) -> pd.DataFrame:
    repaired = df[df["repair_applied_num"] == 1].copy()
    repaired = repaired[
        [
            "market",
            "symbol",
            "statement",
            "field_name",
            "rule_value",
            "llm_value",
            "final_value",
            "expected_value",
            "match_final_num",
            "llm_decision",
            "evidence",
            "llm_reason",
        ]
    ]
    return repaired


def main() -> None:
    args = parse_args()
    template = Path(args.template)
    out_dir = Path(args.out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    df = pd.read_excel(template, sheet_name=args.sheet)
    if "is_match_rule" not in df.columns:
        df["is_match_rule"] = df.get("is_match")
    if "is_match_final" not in df.columns:
        df["is_match_final"] = pd.NA

    df["match_rule_num"] = df["is_match_rule"].apply(parse_match)
    df["match_final_num"] = df["is_match_final"].apply(parse_match)
    df["review_required_num"] = df.get("review_required_recommended", 1).apply(parse_bool01)
    df["repair_applied_num"] = df.get("repair_applied", 0).apply(parse_bool01)

    overall = build_summary(df)
    by_market = by_market_summary(df)
    repaired_details = false_repair_details(df)

    p_overall = out_dir / "ablation_overall.csv"
    p_market = out_dir / "ablation_by_market.csv"
    p_repair = out_dir / "ablation_repair_details.csv"
    p_summary = out_dir / "summary.md"

    overall.to_csv(p_overall, index=False, encoding="utf-8-sig")
    by_market.to_csv(p_market, index=False, encoding="utf-8-sig")
    repaired_details.to_csv(p_repair, index=False, encoding="utf-8-sig")

    with p_summary.open("w", encoding="utf-8") as f:
        f.write("# LLM Ablation Summary\n\n")
        r = overall.iloc[0]
        f.write(f"- total_rows: {int(r['total_rows'])}\n")
        f.write(f"- reviewed_rule: {int(r['reviewed_rule'])}\n")
        f.write(f"- reviewed_final: {int(r['reviewed_final'])}\n")
        f.write(f"- accuracy_rule: {fmt_rate(r['accuracy_rule'])}\n")
        f.write(f"- accuracy_final: {fmt_rate(r['accuracy_final'])}\n")
        f.write(f"- accuracy_delta: {fmt_rate(r['accuracy_delta'])}\n")
        f.write(f"- review_required_rows: {int(r['review_required_rows'])}\n")
        f.write(f"- workload_rate: {fmt_rate(r['workload_rate'])}\n")
        f.write(f"- workload_reduction_rate: {fmt_rate(r['workload_reduction_rate'])}\n")
        f.write(f"- repair_applied_rows: {int(r['repair_applied_rows'])}\n")
        f.write(f"- false_repair_rows: {int(r['false_repair_rows'])}\n")
        f.write(f"- false_repair_rate: {fmt_rate(r['false_repair_rate'])}\n")
        f.write("\n## By Market\n\n")
        if not by_market.empty:
            f.write(by_market.to_markdown(index=False))
            f.write("\n")

    print(f"[OK] Reports generated in: {out_dir}")
    print(f"[OK] {p_overall}")
    print(f"[OK] {p_market}")
    print(f"[OK] {p_repair}")
    print(f"[OK] {p_summary}")


if __name__ == "__main__":
    main()
