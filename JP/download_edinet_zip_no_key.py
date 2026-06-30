#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Download EDINET XBRL ZIP from web UI without API key.

Example:
  ./venv/bin/python JP/download_edinet_zip_no_key.py \
    --keyword 7203 \
    --prefer-asr \
    --out JP/toyota_asr.zip
"""

from __future__ import annotations

import argparse
import re
import tempfile
import zipfile
from pathlib import Path
from typing import List, Optional

from playwright.sync_api import sync_playwright


SEARCH_URL = "https://disclosure2.edinet-fsa.go.jp/weee0030.aspx"


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(description="Download EDINET XBRL ZIP (no API key, web mode).")
    p.add_argument("--keyword", required=True, help="Keyword for search, e.g. 7203 or トヨタ")
    p.add_argument("--out", required=True, help="Output ZIP path")
    p.add_argument("--result-index", type=int, default=1, help="Fallback result index (1-based) if no ASR match")
    p.add_argument("--prefer-asr", action="store_true", help="Prefer annual securities report (asr) if available")
    p.add_argument("--headless", action="store_true", default=True, help="Run browser in headless mode")
    return p.parse_args()


def extract_tokens(html: str) -> List[str]:
    return list(dict.fromkeys(re.findall(r"Weee0030XbrlClick\('([^']+)'\)", html)))


def xbrl_names_from_zip(path: Path) -> List[str]:
    with zipfile.ZipFile(path, "r") as zf:
        return [n for n in zf.namelist() if n.lower().endswith(".xbrl")]


def download_token_zip(page, token: str, out_path: Path) -> None:
    selector = f"a[onclick=\"javascript:Weee0030XbrlClick('{token}');\"]"
    with page.expect_download(timeout=120000) as dl_info:
        page.click(selector)
    download = dl_info.value
    download.save_as(str(out_path))


def main() -> None:
    args = parse_args()
    out_path = Path(args.out)
    out_path.parent.mkdir(parents=True, exist_ok=True)

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=args.headless)
        page = browser.new_page()
        page.goto(SEARCH_URL, wait_until="domcontentloaded", timeout=120000)
        page.fill("#W0018vD_KEYWORD", args.keyword)
        page.click("#W0018BTNBTN_SEARCH")
        page.wait_for_timeout(5000)
        html = page.content()
        tokens = extract_tokens(html)
        if not tokens:
            browser.close()
            raise SystemExit("No XBRL download link found on search result page.")

        picked_path: Optional[Path] = None
        picked_token: Optional[str] = None
        picked_xbrls: List[str] = []

        with tempfile.TemporaryDirectory(prefix="edinet_dl_") as tmp_dir:
            tmp_dir_path = Path(tmp_dir)

            if args.prefer_asr:
                for i, token in enumerate(tokens, 1):
                    tmp_zip = tmp_dir_path / f"candidate_{i}.zip"
                    download_token_zip(page, token, tmp_zip)
                    xbrls = xbrl_names_from_zip(tmp_zip)
                    if any("-asr-" in n.lower() for n in xbrls):
                        picked_path = tmp_zip
                        picked_token = token
                        picked_xbrls = xbrls
                        break

            if picked_path is None:
                idx = max(1, args.result_index) - 1
                if idx >= len(tokens):
                    idx = 0
                token = tokens[idx]
                tmp_zip = tmp_dir_path / "fallback.zip"
                download_token_zip(page, token, tmp_zip)
                picked_path = tmp_zip
                picked_token = token
                picked_xbrls = xbrl_names_from_zip(tmp_zip)

            out_path.write_bytes(picked_path.read_bytes())

        browser.close()

    print(f"[OK] Downloaded ZIP: {out_path}")
    print(f"[OK] Token: {picked_token}")
    if picked_xbrls:
        print("[OK] XBRL entries:")
        for name in picked_xbrls[:5]:
            print("  -", name)


if __name__ == "__main__":
    main()
