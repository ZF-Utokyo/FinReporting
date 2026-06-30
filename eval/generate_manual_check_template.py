#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Generate manual verification template from batch run log.

Usage:
  ./venv/bin/python eval/generate_manual_check_template.py \
    --run-log eval/outputs/run_log_us_jp_cn_YYYYMMDD_HHMMSS.csv \
    --out eval/manual_check_template.xlsx
"""

from __future__ import annotations

import argparse
from pathlib import Path
from typing import Dict, List

import pandas as pd


FIELD_MAP = {
    "IS": [
        "total_revenue",
        "operating_income",
        "income_before_income_taxes",
        "net_income",
        "net_income_per_share_basic",
    ],
    "BS": [
        "total_assets",
        "cash_and_cash_equivalents",
        "accounts_receivable",
        "inventories",
        "total_liabilities",
        "total_shareholders_equity",
        "total_liabilities_and_shareholders_equity",
    ],
    "CF": [
        "net_income",
        "net_cash_operating",
        "net_cash_investing",
        "net_cash_financing",
        "net_change_in_cash",
        "cash_end_of_period",
    ],
}


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(description="Generate manual check template from run log.")
    p.add_argument("--run-log", required=True, help="Path to batch run log CSV")
    p.add_argument("--out", default="eval/manual_check_template.xlsx", help="Output Excel template path")
    return p.parse_args()


def normalize_symbol(market: str, symbol_raw: str) -> str:
    s = (symbol_raw or "").strip()
    if market == "cn":
        d = "".join(ch for ch in s if ch.isdigit())
        return d.zfill(6) if d else s
    if market == "jp":
        d = "".join(ch for ch in s if ch.isdigit())
        return d.zfill(4) if d else s
    return s.upper()


def sheet_names_for_market(market: str) -> Dict[str, str]:
    if market == "us":
        return {"IS": "US_FIN_IS", "BS": "US_FIN_BS", "CF": "US_FIN_CF"}
    if market == "jp":
        return {"IS": "JP_FIN_IS", "BS": "JP_FIN_BS", "CF": "JP_FIN_CF"}
    return {"IS": "CN_FIN_IS", "BS": "CN_FIN_BS", "CF": "CN_FIN_CF"}


def load_row0(workbook: Path, sheet: str) -> pd.Series:
    df = pd.read_excel(workbook, sheet_name=sheet)
    if df.empty:
        return pd.Series(dtype=object)
    return df.iloc[0]


def main() -> None:
    args = parse_args()
    run_log = Path(args.run_log)
    out_path = Path(args.out)
    out_path.parent.mkdir(parents=True, exist_ok=True)

    log_df = pd.read_csv(run_log, dtype=str).fillna("")
    ok_df = log_df[log_df["status"] == "success"].copy()

    rows: List[dict] = []
    for _, r in ok_df.iterrows():
        market = str(r.get("market", "")).strip().lower()
        symbol = normalize_symbol(market, str(r.get("symbol", "")))
        company = str(r.get("company", "")).strip()
        raw_path = Path(str(r.get("raw_workbook", "")).strip())
        if not raw_path.exists():
            continue

        sheets = sheet_names_for_market(market)
        try:
            is_row = load_row0(raw_path, sheets["IS"])
            bs_row = load_row0(raw_path, sheets["BS"])
            cf_row = load_row0(raw_path, sheets["CF"])
        except Exception:
            continue

        meta_fye = is_row.get("fiscal_year_end_date")
        meta_filing = is_row.get("filing_date")
        meta_form = is_row.get("form_type")
        meta_source = is_row.get("source_filing_id", "")

        source_map = {"IS": is_row, "BS": bs_row, "CF": cf_row}
        for stmt, fields in FIELD_MAP.items():
            row0 = source_map[stmt]
            for f in fields:
                rows.append(
                    {
                        "market": market,
                        "symbol": symbol,
                        "company": company,
                        "statement": stmt,
                        "field_name": f,
                        "extracted_value": row0.get(f),
                        "expected_value": "",
                        "is_match": "",
                        "reviewer": "",
                        "comment": "",
                        "fiscal_year_end_date": meta_fye,
                        "filing_date": meta_filing,
                        "form_type": meta_form,
                        "source_filing_id": meta_source,
                        "raw_workbook": str(raw_path),
                    }
                )

    checklist = pd.DataFrame(rows)
    instructions = pd.DataFrame(
        [
            ["How to use", "Fill expected_value from filing, then set is_match as 1/0 (or yes/no)."],
            ["Market scope", "US / JP / CN non-financial companies, FY only."],
            ["Core rule", "Compare against consolidated annual statements."],
            ["Field hints", "IS: revenue/profit; BS: assets/liabilities/equity; CF: CFO/CFI/CFF/cash."],
        ],
        columns=["Item", "Instruction"],
    )

    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        checklist.to_excel(writer, index=False, sheet_name="checklist")
        instructions.to_excel(writer, index=False, sheet_name="instructions")

    print(f"[OK] Manual check template generated: {out_path}")
    print(f"[INFO] Rows: {len(checklist)}")


if __name__ == "__main__":
    main()
