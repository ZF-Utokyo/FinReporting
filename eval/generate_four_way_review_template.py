#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Generate four-way manual review template:
1) Rule-only
2) Rule + LLM-verify
3) Rule + LLM-verify/repair
4) LLM-only
"""

from __future__ import annotations

import argparse
from pathlib import Path
from typing import Dict, Tuple

import pandas as pd


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(description="Generate four-way ablation review template.")
    p.add_argument("--base-template", required=True, help="Checklist template path")
    p.add_argument("--sheet", default="checklist", help="Checklist sheet name")
    p.add_argument("--verify-csv", required=True, help="CSV from llm_cn_verifier.py --verify-only")
    p.add_argument("--repair-csv", required=True, help="CSV from llm_cn_verifier.py (repair mode)")
    p.add_argument("--llm-only-csv", required=True, help="CSV from eval/run_llm_only_cn.py")
    p.add_argument("--out", default="eval/manual_check_template_four_way.xlsx", help="Output template path")
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


def main() -> None:
    args = parse_args()
    base_df = pd.read_excel(Path(args.base_template), sheet_name=args.sheet)
    verify_df = pd.read_csv(Path(args.verify_csv), dtype=str).fillna("")
    repair_df = pd.read_csv(Path(args.repair_csv), dtype=str).fillna("")
    llm_only_df = pd.read_csv(Path(args.llm_only_csv), dtype=str).fillna("")

    verify_idx = build_idx(verify_df)
    repair_idx = build_idx(repair_df)
    llm_only_idx = build_idx(llm_only_df)

    rows = []
    for _, row in base_df.iterrows():
        market = str(row.get("market", "")).strip().lower()
        symbol = normalize_symbol(market, str(row.get("symbol", "")))
        statement = str(row.get("statement", "")).strip().upper()
        field_name = str(row.get("field_name", "")).strip()
        key = (market, symbol, statement, field_name)

        vr = verify_idx.get(key, {})
        rp = repair_idx.get(key, {})
        lo = llm_only_idx.get(key, {})

        rule_value = row.get("extracted_value")
        verify_value = vr.get("final_value", rule_value)
        repair_value = rp.get("final_value", rule_value)
        llm_only_value = lo.get("final_value", "")

        out = dict(row.to_dict())
        out["symbol"] = symbol

        out["rule_value"] = rule_value
        out["verify_value"] = verify_value
        out["repair_value"] = repair_value
        out["llm_only_value"] = llm_only_value

        out["verify_decision"] = vr.get("llm_decision", "")
        out["repair_decision"] = rp.get("llm_decision", "")
        out["llm_only_decision"] = lo.get("llm_only_decision", "")

        out["verify_reason"] = vr.get("reason", "")
        out["repair_reason"] = rp.get("reason", "")
        out["llm_only_reason"] = lo.get("llm_only_reason", "")

        out["verify_evidence"] = vr.get("evidence_json", "")
        out["repair_evidence"] = rp.get("evidence_json", "")
        out["llm_only_evidence"] = lo.get("evidence_json", "")

        out["verify_review_required"] = vr.get("review_required_recommended", "1")
        out["repair_review_required"] = rp.get("review_required_recommended", "1")
        out["llm_only_review_required"] = lo.get("review_required_recommended", "1")

        out["repair_applied"] = rp.get("repair_applied", "0")
        out["repair_guard_fail"] = rp.get("guard_fail", "")

        out["is_match_rule"] = ""
        out["is_match_verify"] = ""
        out["is_match_repair"] = ""
        out["is_match_llm_only"] = ""
        rows.append(out)

    out_df = pd.DataFrame(rows)
    instructions = pd.DataFrame(
        [
            ["How to label", "Fill expected_value, then set is_match_* as 1/0 (or yes/no)."],
            ["Rule-only", "Compare expected_value vs rule_value."],
            ["Rule+Verify", "Compare expected_value vs verify_value (verify-only mode)."],
            ["Rule+Repair", "Compare expected_value vs repair_value (guardrailed repair mode)."],
            ["LLM-only", "Compare expected_value vs llm_only_value."],
            ["Workload", "Use *_review_required columns for workload by system."],
            ["False repair", "repair_applied=1 and is_match_repair=0 counts as false repair."],
        ],
        columns=["Item", "Instruction"],
    )

    out_path = Path(args.out)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        out_df.to_excel(writer, index=False, sheet_name="checklist")
        instructions.to_excel(writer, index=False, sheet_name="instructions")
    print(f"[OK] Four-way template generated: {out_path}")
    print(f"[INFO] Rows: {len(out_df)}")


if __name__ == "__main__":
    main()
