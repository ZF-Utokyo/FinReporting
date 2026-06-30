#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""Unified entry point for common extraction runs."""

from __future__ import annotations

import argparse
import shlex
import subprocess
import sys
from pathlib import Path
from typing import List


ROOT = Path(__file__).resolve().parent


def default_out(prefix: str, symbol: str) -> str:
    return str(Path("outputs") / f"{prefix}_{symbol.lower()}_3statements.xlsx")


def run_command(cmd: List[str]) -> int:
    print("[RUN] " + shlex.join(cmd))
    return subprocess.run(cmd, cwd=ROOT).returncode


def run_cn(args: argparse.Namespace) -> int:
    out = args.out or default_out("cn", args.symbol)
    Path(ROOT / out).parent.mkdir(parents=True, exist_ok=True)
    cmd = [
        sys.executable,
        "CN/export_three_statements_excel_cn.py",
        "--symbol",
        args.symbol,
        "--pdf",
        args.pdf,
        "--schema-file",
        args.schema_file,
        "--out",
        out,
    ]
    if args.company_name:
        cmd.extend(["--company-name", args.company_name])
    if args.llm_verify:
        cmd.append("--llm-verify")
    return run_command(cmd)


def run_us(args: argparse.Namespace) -> int:
    out = args.out or default_out("us", args.symbol)
    Path(ROOT / out).parent.mkdir(parents=True, exist_ok=True)
    cmd = [
        sys.executable,
        "US/export_three_statements_excel.py",
        "--symbol",
        args.symbol,
        "--cik",
        args.cik,
        "--form-type",
        args.form_type,
        "--out",
        out,
    ]
    return run_command(cmd)


def run_jp(args: argparse.Namespace) -> int:
    out = args.out or default_out("jp", args.symbol or "company")
    Path(ROOT / out).parent.mkdir(parents=True, exist_ok=True)
    cmd = [
        sys.executable,
        "JP/export_three_statements_excel_jp.py",
        "--out",
        out,
    ]
    if args.symbol:
        cmd.extend(["--symbol", args.symbol])
    if args.company_name:
        cmd.extend(["--company-name", args.company_name])
    if args.xbrl_zip:
        cmd.extend(["--xbrl-zip", args.xbrl_zip])
    if args.edinet_code:
        cmd.extend(["--edinet-code", args.edinet_code])
    if args.edinet_key:
        cmd.extend(["--edinet-key", args.edinet_key])
    if args.report_date:
        cmd.extend(["--report-date", args.report_date])
    if args.filing_date:
        cmd.extend(["--filing-date", args.filing_date])
    return run_command(cmd)


def run_smoke(_: argparse.Namespace) -> int:
    return run_command([sys.executable, "run_smoke_test.py"])


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Run FinReporting examples from one entry point."
    )
    sub = parser.add_subparsers(dest="command", required=True)

    smoke = sub.add_parser("smoke", help="Run local repository smoke checks.")
    smoke.set_defaults(func=run_smoke)

    cn = sub.add_parser("cn", help="Run CN extraction from a local annual-report PDF.")
    cn.add_argument("--symbol", required=True, help="A-share symbol, e.g. 300750")
    cn.add_argument("--pdf", required=True, help="Path to local annual-report PDF")
    cn.add_argument("--company-name", default="", help="Optional company name keyword")
    cn.add_argument("--schema-file", default="schemas/CN_Schemas.xlsx")
    cn.add_argument("--out", default="", help="Output Excel path")
    cn.add_argument("--llm-verify", action="store_true", help="Run optional LLM verify/repair")
    cn.set_defaults(func=run_cn)

    us = sub.add_parser("us", help="Run US extraction from SEC XBRL.")
    us.add_argument("--symbol", required=True, help="Ticker, e.g. AAPL")
    us.add_argument("--cik", required=True, help="Zero-padded CIK, e.g. 0000320193")
    us.add_argument("--form-type", default="10-K")
    us.add_argument("--out", default="", help="Output Excel path")
    us.set_defaults(func=run_us)

    jp = sub.add_parser("jp", help="Run JP extraction from local EDINET ZIP or API.")
    jp.add_argument("--symbol", default="", help="Security code, e.g. 7203")
    jp.add_argument("--company-name", default="", help="Optional filer name keyword")
    jp.add_argument("--xbrl-zip", default="", help="Path to local EDINET type=1 ZIP")
    jp.add_argument("--edinet-code", default="", help="Optional EDINET code for API mode")
    jp.add_argument("--edinet-key", default="", help="Optional EDINET API key for API mode")
    jp.add_argument("--report-date", default="", help="Override fiscal year end date")
    jp.add_argument("--filing-date", default="", help="Override filing date")
    jp.add_argument("--out", default="", help="Output Excel path")
    jp.set_defaults(func=run_jp)

    return parser.parse_args()


def main() -> int:
    args = parse_args()
    return args.func(args)


if __name__ == "__main__":
    raise SystemExit(main())

