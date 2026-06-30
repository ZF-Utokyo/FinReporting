#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
批量提取多个公司的XBRL现金流量表数据

Usage:
    python batch_extract.py --input companies.csv --out output.csv
"""

from __future__ import annotations

import argparse
import csv
import sys
from pathlib import Path
from typing import List, Dict

import pandas as pd

from extract_xbrl_cash_flow import extract_xbrl_cash_flow, XBRLCashFlowRecord


def load_companies(input_file: str) -> List[Dict]:
    """从CSV文件加载公司列表"""
    df = pd.read_csv(input_file)
    
    required_cols = ["cik", "report_date", "symbol"]
    missing = [col for col in required_cols if col not in df.columns]
    if missing:
        raise ValueError(f"Missing required columns: {missing}")
    
    companies = []
    for _, row in df.iterrows():
        companies.append({
            "cik": str(row["cik"]).zfill(10),
            "report_date": str(row["report_date"]),
            "symbol": str(row["symbol"]),
            "form_type": str(row.get("form_type", "10-K")),
        })
    
    return companies


def main():
    parser = argparse.ArgumentParser(
        description="Batch extract US cash flow statements from SEC XBRL"
    )
    parser.add_argument("--input", required=True, help="Input CSV file with columns: cik, report_date, symbol, [form_type]")
    parser.add_argument("--out", help="Output CSV file path")
    parser.add_argument("--continue-on-error", action="store_true", help="Continue processing on errors")
    
    args = parser.parse_args()
    
    # 加载公司列表
    print(f"[INFO] Loading companies from {args.input}...")
    companies = load_companies(args.input)
    print(f"[INFO] Found {len(companies)} companies")
    
    # 批量处理
    records = []
    errors = []
    
    for i, company in enumerate(companies, 1):
        print(f"\n[{i}/{len(companies)}] Processing {company['symbol']} (CIK: {company['cik']})...")
        
        try:
            record = extract_xbrl_cash_flow(
                cik=company["cik"],
                report_date=company["report_date"],
                symbol=company["symbol"],
                form_type=company.get("form_type", "10-K"),
            )
            records.append(record)
            print(f"[OK] Successfully extracted {company['symbol']}")
        
        except Exception as e:
            error_msg = f"Error processing {company['symbol']}: {e}"
            print(f"[ERROR] {error_msg}")
            errors.append({
                "symbol": company["symbol"],
                "cik": company["cik"],
                "report_date": company["report_date"],
                "error": str(e),
            })
            
            if not args.continue_on_error:
                print("\n[FATAL] Stopping due to error (use --continue-on-error to continue)")
                sys.exit(1)
    
    # 汇总结果
    print("\n" + "="*60)
    print(f"Summary: {len(records)} succeeded, {len(errors)} failed")
    print("="*60)
    
    if errors:
        print("\nErrors:")
        for err in errors:
            print(f"  {err['symbol']}: {err['error']}")
    
    # 保存结果
    if records:
        from dataclasses import asdict
        df = pd.DataFrame([asdict(r) for r in records])
        
        if args.out:
            df.to_csv(args.out, index=False)
            print(f"\n[OK] Saved {len(records)} records to {args.out}")
        
if __name__ == "__main__":
    main()
