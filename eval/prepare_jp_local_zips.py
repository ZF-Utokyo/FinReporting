#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Download JP EDINET ZIP files (no API key, web mode) for batch extraction.

Usage:
  ./venv/bin/python eval/prepare_jp_local_zips.py \
    --in-csv eval/jp_non_fin_20.csv \
    --out-csv eval/outputs/jp_non_fin_20_with_zip.csv \
    --zip-dir eval/outputs/jp_zips
"""

from __future__ import annotations

import argparse
import csv
import subprocess
import sys
import time
from pathlib import Path
from typing import List


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(description="Prepare JP local ZIP list by downloading EDINET ZIPs.")
    p.add_argument("--in-csv", default="eval/jp_non_fin_20.csv", help="Input JP symbol list CSV")
    p.add_argument("--out-csv", default="eval/outputs/jp_non_fin_20_with_zip.csv", help="Output CSV with xbrl_zip column")
    p.add_argument("--zip-dir", default="eval/outputs/jp_zips", help="Directory to store downloaded ZIPs")
    p.add_argument("--limit", type=int, default=0, help="Optional limit (0 = all)")
    p.add_argument("--overwrite", action="store_true", help="Re-download existing ZIP files")
    p.add_argument("--retries", type=int, default=2, help="Retry count per symbol")
    p.add_argument("--sleep-seconds", type=float, default=0.5, help="Sleep between symbols")
    return p.parse_args()


def run_cmd(cmd: List[str]) -> tuple[int, str, str]:
    p = subprocess.run(cmd, capture_output=True, text=True)
    return p.returncode, p.stdout, p.stderr


def main() -> None:
    args = parse_args()
    in_csv = Path(args.in_csv)
    out_csv = Path(args.out_csv)
    zip_dir = Path(args.zip_dir)
    zip_dir.mkdir(parents=True, exist_ok=True)
    out_csv.parent.mkdir(parents=True, exist_ok=True)

    with in_csv.open("r", encoding="utf-8-sig", newline="") as f:
        rows = list(csv.DictReader(f))
    if args.limit > 0:
        rows = rows[: args.limit]

    out_rows = []
    py = sys.executable

    for i, row in enumerate(rows, start=1):
        symbol = str(row.get("symbol", "")).strip()
        company = str(row.get("company", "")).strip()
        notes = str(row.get("notes", "")).strip()
        zip_path = zip_dir / f"jp_{symbol}.zip"
        status = "success"
        error = ""

        if not symbol:
            status = "failed"
            error = "symbol empty"
        else:
            if args.overwrite or (not zip_path.exists()):
                cmd = [
                    py,
                    "JP/download_edinet_zip_no_key.py",
                    "--keyword",
                    symbol,
                    "--prefer-asr",
                    "--out",
                    str(zip_path),
                ]

                code = 1
                out = ""
                err = ""
                for k in range(args.retries + 1):
                    code, out, err = run_cmd(cmd)
                    if code == 0 and zip_path.exists():
                        break
                    if k < args.retries:
                        time.sleep(1.0 * (k + 1))

                if code != 0 or (not zip_path.exists()):
                    status = "failed"
                    error = (err or out).strip()[:2000]

        out_rows.append(
            {
                "symbol": symbol,
                "company": company,
                "notes": notes,
                "xbrl_zip": str(zip_path) if status == "success" else "",
                "download_status": status,
                "download_error": error,
            }
        )
        print(f"[INFO] ({i}/{len(rows)}) symbol={symbol} status={status}")
        if args.sleep_seconds > 0:
            time.sleep(args.sleep_seconds)

    with out_csv.open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(
            f,
            fieldnames=["symbol", "company", "notes", "xbrl_zip", "download_status", "download_error"],
        )
        w.writeheader()
        for r in out_rows:
            w.writerow(r)

    ok = sum(1 for r in out_rows if r["download_status"] == "success")
    print(f"[OK] JP ZIP preparation finished: success={ok}/{len(out_rows)}")
    print(f"[OK] Output CSV: {out_csv}")


if __name__ == "__main__":
    main()

