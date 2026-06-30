#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Generate manual review template for Rule-only vs Rule+LLM ablation.

Usage:
  ./venv/bin/python eval/generate_llm_review_template.py \
    --base-template eval/manual_check_template.xlsx \
    --llm-audit eval/outputs/llm_cn_audit.csv \
    --out eval/manual_check_template_llm.xlsx
"""

from __future__ import annotations

import argparse
from pathlib import Path
from typing import Dict, Tuple

import pandas as pd


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(description="Generate LLM ablation review template.")
    p.add_argument("--base-template", required=True, help="Existing checklist template path")
    p.add_argument("--sheet", default="checklist", help="Checklist sheet name")
    p.add_argument("--llm-audit", required=True, help="LLM audit CSV from llm_cn_verifier.py")
    p.add_argument("--out", default="eval/manual_check_template_llm.xlsx", help="Output review template path")
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


def build_audit_index(df: pd.DataFrame) -> Dict[Tuple[str, str, str, str], dict]:
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


def main() -> None:
    args = parse_args()
    base_path = Path(args.base_template)
    llm_path = Path(args.llm_audit)
    out_path = Path(args.out)
    out_path.parent.mkdir(parents=True, exist_ok=True)

    checklist = pd.read_excel(base_path, sheet_name=args.sheet)
    llm_df = pd.read_csv(llm_path, dtype=str).fillna("")
    audit_idx = build_audit_index(llm_df)

    rows = []
    for _, row in checklist.iterrows():
        market = str(row.get("market", "")).strip().lower()
        symbol = normalize_symbol(market, str(row.get("symbol", "")))
        statement = str(row.get("statement", "")).strip().upper()
        field_name = str(row.get("field_name", "")).strip()
        key = (market, symbol, statement, field_name)
        a = audit_idx.get(key, {})

        rule_value = row.get("extracted_value")
        final_value = a.get("final_value", rule_value)
        final_source = a.get("final_source", "rule")
        llm_decision = a.get("llm_decision", "NO_LLM")
        review_required = a.get("review_required_recommended", "1")

        out = dict(row.to_dict())
        out["symbol"] = symbol
        out["rule_value"] = rule_value
        out["llm_decision"] = llm_decision
        out["llm_value"] = a.get("llm_value", "")
        out["llm_status"] = a.get("llm_status", "")
        out["evidence"] = a.get("evidence_json", "")
        out["llm_reason"] = a.get("reason", "")
        out["guard_fail"] = a.get("guard_fail", "")
        out["repair_allowed"] = a.get("repair_allowed", "")
        out["repair_applied"] = a.get("repair_applied", "")
        out["final_value"] = final_value
        out["final_source"] = final_source
        out["review_required_recommended"] = review_required
        out["is_match_rule"] = out.get("is_match", "")
        out["is_match_final"] = ""
        rows.append(out)

    out_df = pd.DataFrame(rows)

    instruction = pd.DataFrame(
        [
            ["How to label", "Fill expected_value, then set is_match_rule and is_match_final as 1/0 (or yes/no)."],
            ["Rule baseline", "Compare expected_value with rule_value (same as extracted_value)."],
            ["LLM final", "Compare expected_value with final_value after guardrailed LLM decision."],
            ["Workload metric", "review_required_recommended=1 means recommended manual review row."],
            ["False repair", "repair_applied=1 and is_match_final=0 counts as false repair."],
        ],
        columns=["Item", "Instruction"],
    )

    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        out_df.to_excel(writer, index=False, sheet_name="checklist")
        instruction.to_excel(writer, index=False, sheet_name="instructions")

    print(f"[OK] LLM review template generated: {out_path}")
    print(f"[INFO] Rows: {len(out_df)}")


if __name__ == "__main__":
    main()
