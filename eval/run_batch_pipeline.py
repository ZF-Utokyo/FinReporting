#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Batch run extraction + collaborator check workbook generation.

Example:
  ./venv/bin/python eval/run_batch_pipeline.py --markets cn --limit 3
  ./venv/bin/python eval/run_batch_pipeline.py --markets us,jp,cn --edinet-key "$EDINET_API_KEY"
"""

from __future__ import annotations

import argparse
import csv
import json
import os
import subprocess
import sys
import time
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import requests


SEC_TICKER_URL = "https://www.sec.gov/files/company_tickers.json"
SEC_TIMEOUT = 45


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(description="Batch run FinReporting extraction + check workbook generation.")
    p.add_argument(
        "--markets",
        default="us,jp,cn",
        help="Comma-separated markets subset: us,jp,cn",
    )
    p.add_argument("--list-us", default=None, help="Optional override list CSV for US")
    p.add_argument("--list-jp", default=None, help="Optional override list CSV for JP")
    p.add_argument("--list-cn", default=None, help="Optional override list CSV for CN")
    p.add_argument("--limit", type=int, default=0, help="Optional per-market limit (0 = all)")
    p.add_argument("--out-root", default="eval/outputs", help="Output root for raw/check files and run logs")
    p.add_argument("--overwrite", action="store_true", help="Overwrite existing raw/check files")
    p.add_argument("--edinet-key", default=os.getenv("EDINET_API_KEY"), help="EDINET API key (for JP API mode)")
    p.add_argument("--form-type", default="10-K", help="US form type, default 10-K")
    p.add_argument("--sleep-seconds", type=float, default=0.0, help="Sleep between companies")
    p.add_argument("--retries", type=int, default=1, help="Retry count for extract/check command failures")
    p.add_argument(
        "--sec-user-agent",
        default=os.getenv("SEC_USER_AGENT", "FinReporting/1.0 zhangoutstanding@hotmail.com"),
        help="User-Agent for SEC requests (recommended: include email)",
    )
    p.add_argument("--dry-run", action="store_true", help="Print planned commands without executing")
    return p.parse_args()


def load_rows(csv_path: Path, limit: int) -> List[dict]:
    with csv_path.open("r", encoding="utf-8-sig", newline="") as f:
        rows = list(csv.DictReader(f))
    if limit > 0:
        return rows[:limit]
    return rows


def run_cmd(cmd: List[str], dry_run: bool = False) -> Tuple[int, str, str, float]:
    if dry_run:
        return 0, "DRY_RUN", "", 0.0
    t0 = time.time()
    proc = subprocess.run(cmd, capture_output=True, text=True)
    dt = time.time() - t0
    return proc.returncode, proc.stdout, proc.stderr, dt


def normalize_market_list(raw: str) -> List[str]:
    items = [x.strip().lower() for x in raw.split(",") if x.strip()]
    valid = {"us", "jp", "cn"}
    out = []
    for m in items:
        if m in valid and m not in out:
            out.append(m)
    return out


def sec_ticker_to_cik_map(cache_path: Path, user_agent: str) -> Dict[str, str]:
    if cache_path.exists():
        try:
            payload = json.loads(cache_path.read_text(encoding="utf-8"))
        except Exception:
            payload = None
    else:
        payload = None

    if payload is None:
        headers = {"User-Agent": user_agent, "Accept": "application/json"}
        r = requests.get(SEC_TICKER_URL, headers=headers, timeout=SEC_TIMEOUT)
        r.raise_for_status()
        payload = r.json()
        cache_path.parent.mkdir(parents=True, exist_ok=True)
        cache_path.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")

    mapping: Dict[str, str] = {}
    if isinstance(payload, dict):
        for _, row in payload.items():
            if not isinstance(row, dict):
                continue
            ticker = str(row.get("ticker", "")).upper().strip()
            cik = row.get("cik_str")
            if not ticker or cik is None:
                continue
            try:
                cik10 = str(int(cik)).zfill(10)
            except Exception:
                continue
            mapping[ticker] = cik10
    return mapping


def company_name(row: dict) -> str:
    return (row.get("company") or "").strip()


def run_cmd_with_retry(cmd: List[str], retries: int, dry_run: bool = False) -> Tuple[int, str, str, float]:
    last = (1, "", "", 0.0)
    for i in range(max(0, retries) + 1):
        code, out, err, dt = run_cmd(cmd, dry_run=dry_run)
        if code == 0:
            return code, out, err, dt
        last = (code, out, err, dt)
        if i < retries:
            time.sleep(1.5 * (i + 1))
    return last


def run_market_cn(
    row: dict,
    py: str,
    out_dir: Path,
    overwrite: bool,
    retries: int,
    dry_run: bool,
) -> dict:
    symbol = str(row.get("symbol", "")).strip()
    cname = company_name(row)
    raw_out = out_dir / f"cn_{symbol}_3statements.xlsx"
    raw_extract_out = out_dir / f"cn_{symbol}_raw_extract.xlsx"
    check_out = out_dir / f"check_cn_{symbol}.xlsx"
    meta = {
        "market": "cn",
        "symbol": symbol,
        "company": cname,
        "raw_workbook": str(raw_out),
        "check_workbook": str(check_out),
        "source_mode": "cninfo_api",
    }

    if not symbol:
        meta.update({"status": "skipped_invalid_symbol", "error": "symbol is empty"})
        return meta

    extract_ran = False
    check_ran = False

    if overwrite or (not raw_out.exists()):
        cmd = [
            py,
            "CN/export_three_statements_excel_cn.py",
            "--symbol",
            symbol,
            "--raw-out",
            str(raw_extract_out),
            "--out",
            str(raw_out),
        ]
        if cname:
            cmd.extend(["--company-name", cname])
        code, out, err, dt = run_cmd_with_retry(cmd, retries=retries, dry_run=dry_run)
        extract_ran = True
        if code != 0:
            meta.update(
                {
                    "status": "extract_failed",
                    "error": (err or out).strip()[:2000],
                    "extract_ran": extract_ran,
                    "check_ran": check_ran,
                }
            )
            return meta

    if overwrite or (not check_out.exists()):
        cmd = [py, "CN/convert_three_statements_readable.py", "--in", str(raw_out), "--out", str(check_out)]
        code, out, err, dt = run_cmd_with_retry(cmd, retries=retries, dry_run=dry_run)
        check_ran = True
        if code != 0:
            meta.update(
                {
                    "status": "check_failed",
                    "error": (err or out).strip()[:2000],
                    "extract_ran": extract_ran,
                    "check_ran": check_ran,
                }
            )
            return meta

    meta.update({"status": "success", "error": "", "extract_ran": extract_ran, "check_ran": check_ran})
    return meta


def run_market_us(
    row: dict,
    py: str,
    out_dir: Path,
    overwrite: bool,
    retries: int,
    dry_run: bool,
    cik_map: Dict[str, str],
    form_type: str,
) -> dict:
    symbol = str(row.get("symbol", "")).strip().upper()
    cname = company_name(row)
    cik = str(row.get("cik", "")).strip() or cik_map.get(symbol, "")
    raw_out = out_dir / f"us_{symbol.lower()}_3statements.xlsx"
    check_out = out_dir / f"check_us_{symbol.lower()}.xlsx"
    meta = {
        "market": "us",
        "symbol": symbol,
        "company": cname,
        "cik": cik,
        "raw_workbook": str(raw_out),
        "check_workbook": str(check_out),
        "source_mode": "sec_xbrl",
    }

    if not symbol:
        meta.update({"status": "skipped_invalid_symbol", "error": "symbol is empty"})
        return meta
    if not cik:
        meta.update({"status": "skipped_missing_cik", "error": f"CIK not found for symbol={symbol}"})
        return meta

    extract_ran = False
    check_ran = False

    if overwrite or (not raw_out.exists()):
        cmd = [
            py,
            "US/export_three_statements_excel.py",
            "--symbol",
            symbol,
            "--cik",
            cik,
            "--form-type",
            form_type,
            "--out",
            str(raw_out),
        ]
        code, out, err, dt = run_cmd_with_retry(cmd, retries=retries, dry_run=dry_run)
        extract_ran = True
        if code != 0:
            meta.update(
                {
                    "status": "extract_failed",
                    "error": (err or out).strip()[:2000],
                    "extract_ran": extract_ran,
                    "check_ran": check_ran,
                }
            )
            return meta

    if overwrite or (not check_out.exists()):
        cmd = [py, "US/convert_three_statements_readable.py", "--in", str(raw_out), "--out", str(check_out)]
        code, out, err, dt = run_cmd_with_retry(cmd, retries=retries, dry_run=dry_run)
        check_ran = True
        if code != 0:
            meta.update(
                {
                    "status": "check_failed",
                    "error": (err or out).strip()[:2000],
                    "extract_ran": extract_ran,
                    "check_ran": check_ran,
                }
            )
            return meta

    meta.update({"status": "success", "error": "", "extract_ran": extract_ran, "check_ran": check_ran})
    return meta


def run_market_jp(
    row: dict,
    py: str,
    out_dir: Path,
    overwrite: bool,
    retries: int,
    dry_run: bool,
    edinet_key: Optional[str],
) -> dict:
    symbol = str(row.get("symbol", "")).strip()
    cname = company_name(row)
    xbrl_zip = str(row.get("xbrl_zip", "")).strip()
    raw_out = out_dir / f"jp_{symbol}_3statements.xlsx"
    check_out = out_dir / f"check_jp_{symbol}.xlsx"
    meta = {
        "market": "jp",
        "symbol": symbol,
        "company": cname,
        "raw_workbook": str(raw_out),
        "check_workbook": str(check_out),
        "source_mode": "edinet_api",
    }

    if not symbol:
        meta.update({"status": "skipped_invalid_symbol", "error": "symbol is empty"})
        return meta

    extract_ran = False
    check_ran = False

    if overwrite or (not raw_out.exists()):
        cmd = [py, "JP/export_three_statements_excel_jp.py", "--symbol", symbol, "--out", str(raw_out)]
        if cname:
            cmd.extend(["--company-name", cname])
        if xbrl_zip:
            cmd.extend(["--xbrl-zip", xbrl_zip])
            meta["source_mode"] = "local_xbrl_zip"
        else:
            if not edinet_key:
                meta.update({"status": "skipped_missing_edinet_key", "error": "EDINET key is required for JP API mode"})
                return meta
            cmd.extend(["--edinet-key", edinet_key])
        code, out, err, dt = run_cmd_with_retry(cmd, retries=retries, dry_run=dry_run)
        extract_ran = True
        if code != 0:
            meta.update(
                {
                    "status": "extract_failed",
                    "error": (err or out).strip()[:2000],
                    "extract_ran": extract_ran,
                    "check_ran": check_ran,
                }
            )
            return meta

    if overwrite or (not check_out.exists()):
        cmd = [py, "JP/convert_three_statements_readable.py", "--in", str(raw_out), "--out", str(check_out)]
        code, out, err, dt = run_cmd_with_retry(cmd, retries=retries, dry_run=dry_run)
        check_ran = True
        if code != 0:
            meta.update(
                {
                    "status": "check_failed",
                    "error": (err or out).strip()[:2000],
                    "extract_ran": extract_ran,
                    "check_ran": check_ran,
                }
            )
            return meta

    meta.update({"status": "success", "error": "", "extract_ran": extract_ran, "check_ran": check_ran})
    return meta


def market_csv_path(market: str) -> Path:
    return Path(f"eval/{market}_non_fin_20.csv")


def market_csv_path_with_override(market: str, args: argparse.Namespace) -> Path:
    if market == "us" and args.list_us:
        return Path(args.list_us)
    if market == "jp" and args.list_jp:
        return Path(args.list_jp)
    if market == "cn" and args.list_cn:
        return Path(args.list_cn)
    return market_csv_path(market)


def main() -> None:
    args = parse_args()
    markets = normalize_market_list(args.markets)
    if not markets:
        raise SystemExit("No valid markets provided. Use subset of us,jp,cn")

    py = sys.executable
    out_root = Path(args.out_root)
    out_root.mkdir(parents=True, exist_ok=True)

    cik_map: Dict[str, str] = {}
    if "us" in markets:
        cache_path = out_root / "cache" / "sec_company_tickers.json"
        try:
            cik_map = sec_ticker_to_cik_map(cache_path, user_agent=args.sec_user_agent)
        except Exception as e:
            print(f"[WARN] Failed to load SEC ticker map: {e}")
            cik_map = {}

    rows_out: List[dict] = []
    started_at = datetime.now()

    for market in markets:
        csv_path = market_csv_path_with_override(market, args)
        if not csv_path.exists():
            print(f"[WARN] Skip {market}: list file not found: {csv_path}")
            continue
        companies = load_rows(csv_path, args.limit)
        out_dir = out_root / market
        out_dir.mkdir(parents=True, exist_ok=True)

        print(f"[INFO] Market={market} companies={len(companies)}")
        for idx, row in enumerate(companies, start=1):
            s = str(row.get("symbol", "")).strip()
            print(f"[INFO] ({market} {idx}/{len(companies)}) symbol={s}")
            item_start = datetime.now()

            if market == "cn":
                result = run_market_cn(
                    row,
                    py=py,
                    out_dir=out_dir,
                    overwrite=args.overwrite,
                    retries=args.retries,
                    dry_run=args.dry_run,
                )
            elif market == "us":
                result = run_market_us(
                    row,
                    py=py,
                    out_dir=out_dir,
                    overwrite=args.overwrite,
                    retries=args.retries,
                    dry_run=args.dry_run,
                    cik_map=cik_map,
                    form_type=args.form_type,
                )
            else:
                result = run_market_jp(
                    row,
                    py=py,
                    out_dir=out_dir,
                    overwrite=args.overwrite,
                    retries=args.retries,
                    dry_run=args.dry_run,
                    edinet_key=args.edinet_key,
                )

            item_end = datetime.now()
            result["start_time"] = item_start.isoformat(timespec="seconds")
            result["end_time"] = item_end.isoformat(timespec="seconds")
            result["duration_sec"] = round((item_end - item_start).total_seconds(), 2)
            rows_out.append(result)

            if args.sleep_seconds > 0:
                time.sleep(args.sleep_seconds)

    ended_at = datetime.now()
    ts = ended_at.strftime("%Y%m%d_%H%M%S")
    log_path = out_root / f"run_log_{'_'.join(markets)}_{ts}.csv"
    fields = [
        "market",
        "symbol",
        "company",
        "cik",
        "status",
        "source_mode",
        "raw_workbook",
        "check_workbook",
        "extract_ran",
        "check_ran",
        "start_time",
        "end_time",
        "duration_sec",
        "error",
    ]
    with log_path.open("w", encoding="utf-8", newline="") as f:
        w = csv.DictWriter(f, fieldnames=fields)
        w.writeheader()
        for r in rows_out:
            w.writerow({k: r.get(k, "") for k in fields})

    total = len(rows_out)
    success = sum(1 for r in rows_out if r.get("status") == "success")
    print(f"[OK] Batch finished. success={success}/{total}")
    print(f"[OK] Run log: {log_path}")
    print(f"[INFO] Started: {started_at.isoformat(timespec='seconds')}")
    print(f"[INFO] Ended:   {ended_at.isoformat(timespec='seconds')}")


if __name__ == "__main__":
    main()
