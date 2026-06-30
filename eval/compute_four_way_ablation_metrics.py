#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Compute four-way ablation metrics:
Rule-only / Rule+Verify / Rule+Repair / LLM-only.
"""

from __future__ import annotations

import argparse
from pathlib import Path
from typing import Optional

import pandas as pd


SYSTEMS = [
    ("rule", "is_match_rule", None),
    ("verify", "is_match_verify", "verify_review_required"),
    ("repair", "is_match_repair", "repair_review_required"),
    ("llm_only", "is_match_llm_only", "llm_only_review_required"),
]


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(description="Compute four-way ablation metrics.")
    p.add_argument("--template", required=True, help="Four-way review template path")
    p.add_argument("--sheet", default="checklist", help="Checklist sheet name")
    p.add_argument("--out-dir", default="eval/reports_four_way", help="Output report directory")
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


def system_metrics(df: pd.DataFrame, system: str, match_num_col: str, review_col: Optional[str]) -> dict:
    total = len(df)
    reviewed = int(df[match_num_col].notna().sum())
    correct = int((df[match_num_col] == 1).sum())
    accuracy = safe_rate(correct, reviewed)

    if system == "rule":
        # Baseline assumption: without LLM triage all rows need manual review.
        review_required_rows = total
    else:
        review_required_rows = int(df[review_col].sum()) if review_col else total
    workload_rate = safe_rate(review_required_rows, total)

    return {
        "system": system,
        "total_rows": total,
        "reviewed_rows": reviewed,
        "correct_rows": correct,
        "accuracy": accuracy,
        "review_required_rows": review_required_rows,
        "workload_rate": workload_rate,
    }


def compute_table(df: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for system, match_col, review_col in SYSTEMS:
        rows.append(system_metrics(df, system, match_col + "_num", review_col))
    out = pd.DataFrame(rows)
    rule_workload = out[out["system"] == "rule"]["workload_rate"].iloc[0] if not out.empty else pd.NA
    if pd.isna(rule_workload) or float(rule_workload) == 0.0:
        out["workload_reduction_vs_rule"] = pd.NA
    else:
        out["workload_reduction_vs_rule"] = 1.0 - (out["workload_rate"] / float(rule_workload))
    rule_acc = out[out["system"] == "rule"]["accuracy"].iloc[0] if not out.empty else pd.NA
    out["accuracy_delta_vs_rule"] = out["accuracy"] - rule_acc if not pd.isna(rule_acc) else pd.NA
    return out


def false_repair_metrics(df: pd.DataFrame) -> dict:
    repaired = df[df["repair_applied_num"] == 1]
    repaired_reviewed = repaired[repaired["is_match_repair_num"].notna()]
    false_repair = repaired_reviewed[repaired_reviewed["is_match_repair_num"] == 0]
    return {
        "repair_applied_rows": int(len(repaired)),
        "repair_reviewed_rows": int(len(repaired_reviewed)),
        "false_repair_rows": int(len(false_repair)),
        "false_repair_rate": safe_rate(int(len(false_repair)), int(len(repaired_reviewed))),
    }


def main() -> None:
    args = parse_args()
    template = Path(args.template)
    out_dir = Path(args.out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    df = pd.read_excel(template, sheet_name=args.sheet)
    for _, match_col, review_col in SYSTEMS:
        if match_col not in df.columns:
            df[match_col] = pd.NA
        df[match_col + "_num"] = df[match_col].apply(parse_match)
        if review_col:
            if review_col not in df.columns:
                df[review_col] = 1
            df[review_col] = df[review_col].apply(parse_bool01)

    if "repair_applied" not in df.columns:
        df["repair_applied"] = 0
    df["repair_applied_num"] = df["repair_applied"].apply(parse_bool01)

    overall = compute_table(df)

    by_market_rows = []
    for market, g in df.groupby("market", dropna=False):
        sub = compute_table(g)
        sub["market"] = market
        by_market_rows.append(sub)
    by_market = pd.concat(by_market_rows, ignore_index=True) if by_market_rows else pd.DataFrame()

    fr = false_repair_metrics(df)
    false_repair_df = pd.DataFrame([fr])

    p_overall = out_dir / "four_way_overall.csv"
    p_market = out_dir / "four_way_by_market.csv"
    p_false = out_dir / "four_way_false_repair.csv"
    p_summary = out_dir / "summary.md"

    overall.to_csv(p_overall, index=False, encoding="utf-8-sig")
    by_market.to_csv(p_market, index=False, encoding="utf-8-sig")
    false_repair_df.to_csv(p_false, index=False, encoding="utf-8-sig")

    with p_summary.open("w", encoding="utf-8") as f:
        f.write("# Four-way Ablation Summary\n\n")
        f.write("## Overall\n\n")
        if not overall.empty:
            show = overall.copy()
            for c in ["accuracy", "workload_rate", "workload_reduction_vs_rule", "accuracy_delta_vs_rule"]:
                show[c] = show[c].apply(fmt_rate)
            f.write(show.to_markdown(index=False))
            f.write("\n")
        f.write("\n## False Repair\n\n")
        f.write(
            f"- repair_applied_rows: {int(fr['repair_applied_rows'])}\n"
            f"- repair_reviewed_rows: {int(fr['repair_reviewed_rows'])}\n"
            f"- false_repair_rows: {int(fr['false_repair_rows'])}\n"
            f"- false_repair_rate: {fmt_rate(fr['false_repair_rate'])}\n"
        )
        f.write("\n## By Market\n\n")
        if not by_market.empty:
            show_m = by_market.copy()
            for c in ["accuracy", "workload_rate", "workload_reduction_vs_rule", "accuracy_delta_vs_rule"]:
                show_m[c] = show_m[c].apply(fmt_rate)
            cols = ["market"] + [c for c in show_m.columns if c != "market"]
            f.write(show_m[cols].to_markdown(index=False))
            f.write("\n")

    print(f"[OK] Reports generated in: {out_dir}")
    print(f"[OK] {p_overall}")
    print(f"[OK] {p_market}")
    print(f"[OK] {p_false}")
    print(f"[OK] {p_summary}")


if __name__ == "__main__":
    main()
