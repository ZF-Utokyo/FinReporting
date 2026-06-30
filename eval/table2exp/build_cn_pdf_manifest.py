#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from __future__ import annotations

import argparse
import csv
import shutil
from pathlib import Path
from typing import List


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(description="Build CN PDF manifest for table2 experiment.")
    p.add_argument("--split-csv", default="eval/table2exp/cn_non_fin_table2_10.csv")
    p.add_argument("--pdf-root", default="CN/raw_pdfs")
    p.add_argument("--out-dir", default="eval/table2exp/raw_pdfs")
    p.add_argument("--manifest", default="eval/table2exp/raw_pdfs/manifest_cn_table2.csv")
    p.add_argument("--copy", action="store_true", help="Copy matched PDFs into out-dir")
    return p.parse_args()


def load_rows(csv_path: Path) -> List[dict]:
    with csv_path.open("r", encoding="utf-8-sig", newline="") as f:
        return list(csv.DictReader(f))


def main() -> None:
    args = parse_args()
    split_csv = Path(args.split_csv)
    pdf_root = Path(args.pdf_root)
    out_dir = Path(args.out_dir)
    manifest = Path(args.manifest)
    out_dir.mkdir(parents=True, exist_ok=True)
    manifest.parent.mkdir(parents=True, exist_ok=True)

    rows = load_rows(split_csv)
    out_rows = []

    for r in rows:
        symbol = str(r.get("symbol", "")).strip()
        company = str(r.get("company", "")).strip()
        matches = sorted(pdf_root.glob(f"{symbol}_*_annual.pdf"))
        if not matches:
            out_rows.append(
                {
                    "symbol": symbol,
                    "company": company,
                    "pdf_path": "",
                    "status": "missing",
                    "note": "no CN/raw_pdfs match",
                }
            )
            continue
        src = matches[-1]
        dst = out_dir / src.name
        if args.copy:
            shutil.copy2(src, dst)
            final_path = dst
        else:
            final_path = src
        out_rows.append(
            {
                "symbol": symbol,
                "company": company,
                "pdf_path": str(final_path),
                "status": "ok",
                "note": "",
            }
        )

    with manifest.open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=["symbol", "company", "pdf_path", "status", "note"])
        w.writeheader()
        w.writerows(out_rows)

    ok = sum(1 for x in out_rows if x["status"] == "ok")
    print(f"[OK] Manifest: {manifest}")
    print(f"[INFO] PDF found: {ok}/{len(out_rows)}")


if __name__ == "__main__":
    main()
