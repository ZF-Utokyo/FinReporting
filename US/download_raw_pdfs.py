#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Download latest US 10-K filing pages and render to PDF for manual checking.

Usage:
  ./venv/bin/python US/download_raw_pdfs.py \
    --input-csv eval/us_non_fin_20.csv \
    --out-dir US/raw_pdfs
"""

from __future__ import annotations

import argparse
import csv
import json
import os
import re
import tempfile
import time
from pathlib import Path
from typing import Dict, List, Optional

import requests
from playwright.sync_api import sync_playwright


SEC_TICKER_URL = "https://www.sec.gov/files/company_tickers.json"
SEC_SUBMISSIONS_URL = "https://data.sec.gov/submissions/CIK{cik}.json"
SEC_TIMEOUT = 45
SEC_BLOCK_PHRASE = "Your Request Originates from an Undeclared Automated Tool"


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(description="Download latest US 10-K filing pages as PDFs.")
    p.add_argument("--input-csv", default="eval/us_non_fin_20.csv", help="Input CSV with symbol/company columns")
    p.add_argument("--out-dir", default="US/raw_pdfs", help="Output directory for PDFs")
    p.add_argument("--manifest", default=None, help="Optional output manifest CSV path")
    p.add_argument("--form-type", default="10-K", help="Target form type, default 10-K")
    p.add_argument("--overwrite", action="store_true", help="Overwrite existing PDFs")
    p.add_argument("--retries", type=int, default=2, help="Retry count per symbol")
    p.add_argument("--sleep-seconds", type=float, default=0.35, help="Sleep between SEC requests")
    p.add_argument(
        "--sec-user-agent",
        default=os.getenv("SEC_USER_AGENT", "FinReporting/1.0 (contact: your-email@example.com)"),
        help="SEC User-Agent (recommended with email)",
    )
    return p.parse_args()


def load_list(path: Path) -> List[dict]:
    with path.open("r", encoding="utf-8-sig", newline="") as f:
        return list(csv.DictReader(f))


def sec_http_get(url: str, user_agent: str, accept: str = "application/json") -> requests.Response:
    headers = {
        "User-Agent": user_agent,
        "Accept": accept,
        "Accept-Language": "en-US,en;q=0.9",
    }
    r = requests.get(url, headers=headers, timeout=SEC_TIMEOUT)
    r.raise_for_status()
    return r


def sec_json_get(url: str, user_agent: str) -> dict:
    return sec_http_get(url, user_agent=user_agent, accept="application/json").json()


def sec_text_get(url: str, user_agent: str) -> str:
    r = sec_http_get(
        url,
        user_agent=user_agent,
        accept="text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    )
    return r.text


def load_ticker_to_cik(user_agent: str, cache_path: Optional[Path] = None) -> Dict[str, str]:
    payload = None
    if cache_path and cache_path.exists():
        try:
            payload = json.loads(cache_path.read_text(encoding="utf-8"))
        except Exception:
            payload = None
    if payload is None:
        payload = sec_json_get(SEC_TICKER_URL, user_agent=user_agent)
        if cache_path:
            cache_path.parent.mkdir(parents=True, exist_ok=True)
            cache_path.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")

    out: Dict[str, str] = {}
    if isinstance(payload, dict):
        for _, row in payload.items():
            if not isinstance(row, dict):
                continue
            t = str(row.get("ticker", "")).upper().strip()
            cik = row.get("cik_str")
            if not t or cik is None:
                continue
            try:
                out[t] = str(int(cik)).zfill(10)
            except Exception:
                continue
    return out


def latest_filing_from_submissions(sub: dict, form_type: str) -> Optional[dict]:
    recent = sub.get("filings", {}).get("recent", {})
    forms = recent.get("form", [])
    report_dates = recent.get("reportDate", [])
    filing_dates = recent.get("filingDate", [])
    accessions = recent.get("accessionNumber", [])
    primary_docs = recent.get("primaryDocument", [])

    n = min(len(forms), len(report_dates), len(filing_dates), len(accessions), len(primary_docs))
    for i in range(n):
        form = str(forms[i] or "")
        if form != form_type:
            continue
        if form.endswith("/A"):
            continue
        report_date = str(report_dates[i] or "").strip()
        filing_date = str(filing_dates[i] or "").strip()
        accession = str(accessions[i] or "").strip()
        primary_doc = str(primary_docs[i] or "").strip()
        if not (report_date and filing_date and accession and primary_doc):
            continue
        return {
            "form_type": form,
            "report_date": report_date,
            "filing_date": filing_date,
            "accession": accession,
            "primary_document": primary_doc,
        }
    return None


def safe_slug(s: str) -> str:
    s = s.strip().lower()
    s = re.sub(r"[^a-z0-9._-]+", "_", s)
    return s.strip("_")


def main() -> None:
    args = parse_args()
    input_csv = Path(args.input_csv)
    out_dir = Path(args.out_dir)
    manifest_path = Path(args.manifest) if args.manifest else out_dir / "manifest_us_10k.csv"
    out_dir.mkdir(parents=True, exist_ok=True)
    manifest_path.parent.mkdir(parents=True, exist_ok=True)

    rows = load_list(input_csv)
    ticker_map = load_ticker_to_cik(
        user_agent=args.sec_user_agent,
        cache_path=Path("eval/outputs/cache/sec_company_tickers.json"),
    )

    results: List[dict] = []

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(user_agent=args.sec_user_agent)
        page = context.new_page()

        for idx, row in enumerate(rows, start=1):
            symbol = str(row.get("symbol", "")).upper().strip()
            company = str(row.get("company", "")).strip()
            cik = ticker_map.get(symbol, "")
            base = {
                "symbol": symbol,
                "company": company,
                "cik": cik,
                "form_type": args.form_type,
                "report_date": "",
                "filing_date": "",
                "accession": "",
                "primary_document": "",
                "primary_url": "",
                "html_path": "",
                "pdf_path": "",
                "status": "",
                "error": "",
            }

            if not symbol:
                base["status"] = "failed"
                base["error"] = "empty symbol"
                results.append(base)
                continue
            if not cik:
                base["status"] = "failed"
                base["error"] = "cik not found"
                results.append(base)
                continue

            ok = False
            last_err = ""
            for attempt in range(args.retries + 1):
                try:
                    sub = sec_json_get(SEC_SUBMISSIONS_URL.format(cik=cik), user_agent=args.sec_user_agent)
                    if args.sleep_seconds > 0:
                        time.sleep(args.sleep_seconds)
                    filing = latest_filing_from_submissions(sub, form_type=args.form_type)
                    if filing is None:
                        raise RuntimeError(f"no {args.form_type} filing found")

                    accession = filing["accession"]
                    accession_nodash = accession.replace("-", "")
                    cik_int = str(int(cik))
                    primary_doc = filing["primary_document"]
                    report_date = filing["report_date"]
                    filing_date = filing["filing_date"]
                    primary_url = (
                        f"https://www.sec.gov/Archives/edgar/data/{cik_int}/{accession_nodash}/{primary_doc}"
                    )

                    name = f"{safe_slug(symbol)}_{report_date.replace('-', '')}_{accession_nodash}.pdf"
                    pdf_path = out_dir / name
                    html_path = out_dir / f"{safe_slug(symbol)}_{report_date.replace('-', '')}_{accession_nodash}.html"

                    if args.overwrite or (not pdf_path.exists()):
                        # Fetch filing HTML via requests (with declared SEC User-Agent),
                        # then render locally to PDF to avoid SEC anti-bot block pages.
                        html_text = sec_text_get(primary_url, user_agent=args.sec_user_agent)
                        if SEC_BLOCK_PHRASE in html_text:
                            raise RuntimeError("SEC returned anti-bot page; check SEC_USER_AGENT and request rate")
                        html_path.write_text(html_text, encoding="utf-8")

                        with tempfile.NamedTemporaryFile("w", suffix=".html", delete=False, encoding="utf-8") as tf:
                            tf.write(html_text)
                            tmp_html_path = Path(tf.name)
                        try:
                            page.goto(tmp_html_path.as_uri(), wait_until="networkidle", timeout=120000)
                            page.pdf(path=str(pdf_path), format="A4", print_background=True)
                        finally:
                            try:
                                tmp_html_path.unlink(missing_ok=True)
                            except Exception:
                                pass

                    base.update(
                        {
                            "report_date": report_date,
                            "filing_date": filing_date,
                            "accession": accession,
                            "primary_document": primary_doc,
                            "primary_url": primary_url,
                            "html_path": str(html_path),
                            "pdf_path": str(pdf_path),
                            "status": "success",
                            "error": "",
                        }
                    )
                    ok = True
                    break
                except Exception as e:
                    last_err = str(e)
                    if attempt < args.retries:
                        time.sleep(1.5 * (attempt + 1))

            if not ok:
                base["status"] = "failed"
                base["error"] = last_err[:2000]

            results.append(base)
            print(f"[INFO] ({idx}/{len(rows)}) {symbol} -> {results[-1]['status']}")

        context.close()
        browser.close()

    with manifest_path.open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(
            f,
            fieldnames=[
                "symbol",
                "company",
                "cik",
                "form_type",
                "report_date",
                "filing_date",
                "accession",
                "primary_document",
                "primary_url",
                "html_path",
                "pdf_path",
                "status",
                "error",
            ],
        )
        w.writeheader()
        for r in results:
            w.writerow(r)

    success = sum(1 for r in results if r["status"] == "success")
    print(f"[OK] US raw PDF download finished: success={success}/{len(results)}")
    print(f"[OK] Output dir: {out_dir}")
    print(f"[OK] Manifest: {manifest_path}")


if __name__ == "__main__":
    main()
