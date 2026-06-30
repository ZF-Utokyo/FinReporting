#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Export JP Income Statement / Balance Sheet / Cash Flow to a single Excel file.

Data source:
  1) EDINET API v2 (requires Subscription-Key), or
  2) Local EDINET XBRL ZIP file (no API key required).

Usage:
  python export_three_statements_excel_jp.py \
    --symbol 7203 \
    --company-name "トヨタ自動車" \
    --edinet-key "$EDINET_API_KEY" \
    --out JP/toyota_3statements.xlsx
"""

from __future__ import annotations

import argparse
import io
import json
import os
import re
import zipfile
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import pandas as pd
import requests
from lxml import etree


API_BASE = "https://disclosure.edinet-fsa.go.jp/api/v2"
REQUEST_TIMEOUT = 45

XBRLI_NS = {"xbrli": "http://www.xbrl.org/2003/instance"}


@dataclass
class FilingInfo:
    doc_id: str
    edinet_code: Optional[str]
    sec_code: Optional[str]
    filer_name: Optional[str]
    form_code: Optional[str]
    doc_type_code: Optional[str]
    doc_description: Optional[str]
    period_start: Optional[str]
    period_end: str
    filing_date: str
    is_amendment: bool


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(description="Export JP IS/BS/CF from EDINET to one Excel workbook.")
    p.add_argument("--symbol", help="Security code, e.g. 7203")
    p.add_argument("--company-name", help="Optional filer name keyword, e.g. トヨタ")
    p.add_argument("--edinet-code", help="Optional EDINET code (e.g. E02144)")
    p.add_argument("--edinet-key", default=os.getenv("EDINET_API_KEY"), help="EDINET Subscription-Key")
    p.add_argument("--as-of-date", default=date.today().isoformat(), help="Search anchor date (YYYY-MM-DD)")
    p.add_argument("--lookback-days", type=int, default=400, help="How many days to look back for annual filing")
    p.add_argument("--xbrl-zip", help="Local EDINET ZIP path (type=1 download) or .xbrl file path")
    p.add_argument("--report-date", help="Override fiscal year end date for local mode (YYYY-MM-DD)")
    p.add_argument("--filing-date", help="Override filing date for local mode (YYYY-MM-DD)")
    p.add_argument("--form-type", help="Override form type text for local mode")
    p.add_argument("--source-filing-id", help="Override source filing id for local mode")
    p.add_argument("--out", default="JP/jp_3statements.xlsx", help="Output Excel path")
    return p.parse_args()


def normalize_sec_code(sec_code: Optional[str]) -> Optional[str]:
    if not sec_code:
        return None
    s = re.sub(r"\D", "", str(sec_code))
    if not s:
        return None
    return s[:4]


def safe_get(d: dict, *keys: str) -> Optional[str]:
    for k in keys:
        if k in d and d[k] not in (None, ""):
            return str(d[k])
    return None


def request_json(session: requests.Session, path: str, *, params: dict, api_key: str) -> dict:
    q = dict(params)
    q["Subscription-Key"] = api_key
    r = session.get(f"{API_BASE}{path}", params=q, timeout=REQUEST_TIMEOUT)
    r.raise_for_status()
    payload = r.json()
    if isinstance(payload, dict) and payload.get("statusCode") not in (None, 200, "200"):
        raise RuntimeError(f"EDINET API error: statusCode={payload.get('statusCode')} message={payload.get('message')}")
    return payload


def request_binary(session: requests.Session, path: str, *, params: dict, api_key: str) -> bytes:
    q = dict(params)
    q["Subscription-Key"] = api_key
    r = session.get(f"{API_BASE}{path}", params=q, timeout=REQUEST_TIMEOUT)
    r.raise_for_status()
    ctype = (r.headers.get("content-type") or "").lower()
    if "json" in ctype:
        payload = r.json()
        raise RuntimeError(f"EDINET API binary download failed: {payload}")
    return r.content


def is_annual_security_report(doc: dict) -> bool:
    form_code = safe_get(doc, "formCode", "form_code") or ""
    desc = safe_get(doc, "docDescription", "doc_description") or ""
    # Primary condition by form code; fallback by Japanese description.
    if form_code in {"030000", "030001"}:
        return True
    return "有価証券報告書" in desc


def is_amendment_doc(doc: dict) -> bool:
    desc = safe_get(doc, "docDescription", "doc_description") or ""
    form_code = safe_get(doc, "formCode", "form_code") or ""
    return ("訂正" in desc) or form_code.endswith("1")


def find_latest_annual_filing(
    session: requests.Session,
    api_key: str,
    symbol: str,
    company_name: Optional[str],
    edinet_code: Optional[str],
    as_of_date: str,
    lookback_days: int,
) -> FilingInfo:
    target_symbol = normalize_sec_code(symbol)
    if not target_symbol:
        raise ValueError(f"Invalid symbol: {symbol}")

    base_dt = datetime.strptime(as_of_date, "%Y-%m-%d").date()
    best: Optional[Tuple[datetime, FilingInfo]] = None

    for i in range(lookback_days + 1):
        d = (base_dt - timedelta(days=i)).isoformat()
        payload = request_json(session, "/documents.json", params={"date": d, "type": "2"}, api_key=api_key)
        rows = payload.get("results") or payload.get("Results") or []
        if not isinstance(rows, list):
            continue

        for row in rows:
            sec = normalize_sec_code(safe_get(row, "secCode", "sec_code"))
            if sec != target_symbol:
                continue
            if edinet_code and safe_get(row, "edinetCode", "edinet_code") != edinet_code:
                continue
            filer_name = safe_get(row, "filerName", "filer_name")
            if company_name and filer_name and company_name not in filer_name:
                continue
            if safe_get(row, "xbrlFlag", "xbrl_flag") not in ("1", "true", "True", "TRUE"):
                continue
            if not is_annual_security_report(row):
                continue

            period_end = safe_get(row, "periodEnd", "period_end")
            submit_dt = safe_get(row, "submitDateTime", "submit_date_time", "submitDate")
            doc_id = safe_get(row, "docID", "doc_id")
            if not (period_end and submit_dt and doc_id):
                continue

            # submitDateTime may contain time; normalize.
            filing_date = submit_dt[:10]
            try:
                rank_dt = datetime.strptime(filing_date, "%Y-%m-%d")
            except ValueError:
                continue

            info = FilingInfo(
                doc_id=doc_id,
                edinet_code=safe_get(row, "edinetCode", "edinet_code"),
                sec_code=sec,
                filer_name=filer_name,
                form_code=safe_get(row, "formCode", "form_code"),
                doc_type_code=safe_get(row, "docTypeCode", "doc_type_code"),
                doc_description=safe_get(row, "docDescription", "doc_description"),
                period_start=safe_get(row, "periodStart", "period_start"),
                period_end=period_end,
                filing_date=filing_date,
                is_amendment=is_amendment_doc(row),
            )

            if (best is None) or (rank_dt > best[0]):
                best = (rank_dt, info)

        # Optimization: once found a very recent match, stop scanning too far.
        if best and i > 90:
            break

    if not best:
        raise RuntimeError(
            f"No annual filing found for symbol={target_symbol} "
            f"within {lookback_days} days before {as_of_date}."
        )
    return best[1]


def pick_instance_xbrl_from_zip(blob: bytes) -> Tuple[str, bytes]:
    with zipfile.ZipFile(io.BytesIO(blob), "r") as zf:
        names = zf.namelist()
        candidates = [n for n in names if n.lower().endswith(".xbrl")]
        if not candidates:
            raise RuntimeError("No .xbrl instance file found inside EDINET ZIP.")

        preferred = [n for n in candidates if "/XBRL/PublicDoc/" in n or "XBRL/PublicDoc/" in n]
        if preferred:
            candidates = preferred

        # Prefer annual securities report naming pattern if available.
        candidates_sorted = sorted(
            candidates,
            key=lambda x: (
                0 if "-asr-" in x.lower() else 1,
                len(x),
                x.lower(),
            ),
        )
        picked = candidates_sorted[0]
        return picked, zf.read(picked)


def load_xbrl_source(path: Path) -> Tuple[str, bytes]:
    if not path.exists():
        raise FileNotFoundError(f"Local XBRL source not found: {path}")
    if path.is_dir():
        raise RuntimeError(f"Local XBRL source must be a file, got directory: {path}")

    suffix = path.suffix.lower()
    if suffix == ".zip":
        blob = path.read_bytes()
        return pick_instance_xbrl_from_zip(blob)
    if suffix in {".xbrl", ".xml"}:
        return path.name, path.read_bytes()
    raise RuntimeError(f"Unsupported local source type: {path}. Use .zip or .xbrl/.xml")


def local_name(tag: str) -> str:
    return tag.split("}")[-1] if "}" in tag else tag


def parse_iso_date(raw: Optional[str]) -> Optional[str]:
    if raw is None:
        return None
    s = str(raw).strip()
    if not s:
        return None

    m = re.search(r"(\d{4}-\d{2}-\d{2})", s)
    if m:
        return m.group(1)
    m = re.search(r"(\d{4})/(\d{1,2})/(\d{1,2})", s)
    if m:
        y, mo, d = int(m.group(1)), int(m.group(2)), int(m.group(3))
        return f"{y:04d}-{mo:02d}-{d:02d}"
    m = re.search(r"(\d{4})年(\d{1,2})月(\d{1,2})日", s)
    if m:
        y, mo, d = int(m.group(1)), int(m.group(2)), int(m.group(3))
        return f"{y:04d}-{mo:02d}-{d:02d}"

    return None


def parse_numeric(text: Optional[str]) -> Optional[float]:
    if text is None:
        return None
    s = str(text).strip()
    if not s:
        return None
    if s in {"-", "—", "–", "―"}:
        return None
    # Japanese filings often use triangle marks for negatives.
    s = s.replace(",", "").replace("△", "-").replace("▲", "-").replace("−", "-").replace("－", "-")
    try:
        return float(s)
    except ValueError:
        return None


def parse_xbrl(xml_bytes: bytes) -> Tuple[List[dict], Dict[str, dict]]:
    root = etree.fromstring(xml_bytes)
    contexts: Dict[str, dict] = {}
    for ctx in root.findall(".//xbrli:context", namespaces=XBRLI_NS):
        cid = ctx.get("id")
        if not cid:
            continue
        start = ctx.find(".//xbrli:period/xbrli:startDate", namespaces=XBRLI_NS)
        end = ctx.find(".//xbrli:period/xbrli:endDate", namespaces=XBRLI_NS)
        instant = ctx.find(".//xbrli:period/xbrli:instant", namespaces=XBRLI_NS)
        segments = ctx.findall(".//xbrli:segment", namespaces=XBRLI_NS)

        start_date = start.text.strip() if start is not None and start.text else None
        end_date = end.text.strip() if end is not None and end.text else None
        instant_date = instant.text.strip() if instant is not None and instant.text else None
        duration_days = None
        if start_date and end_date:
            try:
                duration_days = (datetime.strptime(end_date, "%Y-%m-%d") - datetime.strptime(start_date, "%Y-%m-%d")).days
            except ValueError:
                duration_days = None

        contexts[cid] = {
            "start_date": start_date,
            "end_date": end_date,
            "instant": instant_date,
            "duration_days": duration_days,
            "is_consolidated": len(segments) == 0,
        }

    facts: List[dict] = []
    for el in root.iter():
        if not isinstance(el.tag, str):
            continue
        context_ref = el.get("contextRef")
        if not context_ref:
            continue
        if el.get("{http://www.w3.org/2001/XMLSchema-instance}nil") in {"true", "1"}:
            continue
        value = (el.text or "").strip()
        if value == "":
            continue
        facts.append(
            {
                "tag": local_name(el.tag),
                "contextRef": context_ref,
                "unitRef": el.get("unitRef"),
                "decimals": el.get("decimals"),
                "value": value,
            }
        )

    return facts, contexts


def pick_text_fact(facts: List[dict], contexts: Dict[str, dict], tags: List[str]) -> Optional[str]:
    candidates: List[Tuple[int, str]] = []
    tag_rank = {t: i for i, t in enumerate(tags)}

    for f in facts:
        t = f.get("tag")
        if t not in tag_rank:
            continue
        v = str(f.get("value", "")).strip()
        if not v:
            continue
        score = 100 - tag_rank[t]
        ctx = contexts.get(f.get("contextRef") or "")
        if ctx and ctx.get("is_consolidated"):
            score += 10
        candidates.append((score, v))

    if not candidates:
        return None
    candidates.sort(key=lambda x: x[0], reverse=True)
    return candidates[0][1]


def infer_report_date(contexts: Dict[str, dict]) -> Optional[str]:
    scores: Dict[str, int] = {}
    for ctx in contexts.values():
        end_date = parse_iso_date(ctx.get("end_date"))
        instant_date = parse_iso_date(ctx.get("instant"))
        dd = ctx.get("duration_days")
        is_cons = bool(ctx.get("is_consolidated"))

        if end_date:
            score = 2
            if is_cons:
                score += 8
            if isinstance(dd, int):
                if 300 <= dd <= 380:
                    score += 10
                elif dd > 0:
                    score += 1
            scores[end_date] = scores.get(end_date, 0) + score

        if instant_date:
            score = 1 + (4 if is_cons else 0)
            scores[instant_date] = scores.get(instant_date, 0) + score

    if not scores:
        return None

    return sorted(scores.items(), key=lambda x: (x[1], x[0]), reverse=True)[0][0]


def build_local_filing_info(
    source_path: Path,
    symbol: Optional[str],
    company_name: Optional[str],
    edinet_code: Optional[str],
    report_date: Optional[str],
    filing_date: Optional[str],
    form_type: Optional[str],
    source_filing_id: Optional[str],
    facts: List[dict],
    contexts: Dict[str, dict],
) -> FilingInfo:
    inferred_sec = normalize_sec_code(symbol) or normalize_sec_code(
        pick_text_fact(facts, contexts, ["SecurityCodeDEI", "SecurityCode"])
    )
    inferred_edinet = edinet_code or pick_text_fact(facts, contexts, ["EDINETCodeDEI"])
    inferred_filer = company_name or pick_text_fact(
        facts,
        contexts,
        [
            "FilerNameInJapaneseDEI",
            "FilerNameInEnglishDEI",
            "CompanyNameInJapaneseTextBlock",
            "CompanyNameInEnglishTextBlock",
        ],
    )

    inferred_period_end = (
        parse_iso_date(report_date)
        or parse_iso_date(pick_text_fact(facts, contexts, ["CurrentFiscalYearEndDateDEI", "FiscalYearEnd"]))
        or infer_report_date(contexts)
        or date.today().isoformat()
    )
    inferred_filing_date = (
        parse_iso_date(filing_date)
        or parse_iso_date(pick_text_fact(facts, contexts, ["FilingDateCoverPage", "FilingDateDEI", "DateOfFilingDEI"]))
        or date.today().isoformat()
    )

    doc_id = source_filing_id
    if not doc_id:
        m = re.search(r"(S\d{8,})", source_path.name)
        if m:
            doc_id = m.group(1)
    if not doc_id:
        doc_id = f"LOCAL_{source_path.stem}"

    desc = form_type or "有価証券報告書"
    return FilingInfo(
        doc_id=doc_id,
        edinet_code=inferred_edinet,
        sec_code=inferred_sec,
        filer_name=inferred_filer,
        form_code=None,
        doc_type_code=None,
        doc_description=desc,
        period_start=None,
        period_end=inferred_period_end,
        filing_date=inferred_filing_date,
        is_amendment=("訂正" in desc),
    )


def infer_prior_period_end(contexts: Dict[str, dict], report_date: str) -> Optional[str]:
    candidates: List[Tuple[int, str]] = []
    for ctx in contexts.values():
        if ctx.get("end_date") != report_date:
            continue
        if not ctx.get("start_date"):
            continue
        dd = ctx.get("duration_days")
        if not isinstance(dd, int) or dd <= 0:
            continue
        candidates.append((dd, ctx["start_date"]))
    if not candidates:
        return None
    _, start_date = max(candidates, key=lambda x: x[0])
    try:
        prior = datetime.strptime(start_date, "%Y-%m-%d").date() - timedelta(days=1)
        return prior.isoformat()
    except ValueError:
        return None


def choose_value(
    facts: List[dict],
    contexts: Dict[str, dict],
    report_date: str,
    tags: List[str],
    period_type: str,
    unit_keywords: Optional[List[str]] = None,
) -> Optional[float]:
    candidates: List[Tuple[int, float]] = []
    for f in facts:
        if f["tag"] not in tags:
            continue
        ctx = contexts.get(f["contextRef"])
        if not ctx:
            continue

        if period_type == "duration":
            if ctx.get("end_date") != report_date or not ctx.get("start_date"):
                continue
        elif period_type == "instant":
            if (ctx.get("instant") != report_date) and (ctx.get("end_date") != report_date):
                continue
        else:
            continue

        val = parse_numeric(f["value"])
        if val is None:
            continue

        score = 0
        if ctx.get("is_consolidated"):
            score += 100

        dd = ctx.get("duration_days")
        if period_type == "duration" and isinstance(dd, int):
            if 300 <= dd <= 380:
                score += 20
            else:
                score -= 5

        if unit_keywords:
            unit_ref = (f.get("unitRef") or "").lower()
            if any(k.lower() in unit_ref for k in unit_keywords):
                score += 10

        score += max(0, 5 - tags.index(f["tag"]))
        candidates.append((score, val))

    if not candidates:
        return None
    candidates.sort(key=lambda x: x[0], reverse=True)
    return candidates[0][1]


def choose_value_at_date(
    facts: List[dict],
    contexts: Dict[str, dict],
    target_date: str,
    tags: List[str],
    period_type: str,
    unit_keywords: Optional[List[str]] = None,
) -> Optional[float]:
    return choose_value(facts, contexts, target_date, tags, period_type, unit_keywords)


def pick_first_not_none(values: List[Optional[float]]) -> Optional[float]:
    for v in values:
        if v is not None:
            return v
    return None


def build_is_row(symbol: str, filing: FilingInfo, facts: List[dict], contexts: Dict[str, dict]) -> dict:
    report_date = filing.period_end
    non_op_income = choose_value(facts, contexts, report_date, ["NonOperatingIncome"], "duration", ["jpy"])
    non_op_exp = choose_value(facts, contexts, report_date, ["NonOperatingExpenses"], "duration", ["jpy"])
    non_op_net = None
    if non_op_income is not None and non_op_exp is not None:
        non_op_net = non_op_income - non_op_exp

    return {
        "country": "JP",
        "symbol": symbol,
        "company_id": filing.edinet_code,
        "form_type": filing.doc_description or "有価証券報告書",
        "fiscal_year_end_date": report_date,
        "filing_date": filing.filing_date,
        "fiscal_year": int(report_date[:4]),
        "fiscal_period": "FY",
        "source_filing_id": filing.doc_id,
        "is_amendment": 1 if filing.is_amendment else 0,
        "accounting_standard": "JP_GAAP",
        "taxonomy_version": None,
        "statement_scope": "CONSOLIDATED",
        "currency": "JPY",
        "unit_scale": 1,
        "total_revenue": pick_first_not_none(
            [
                choose_value(facts, contexts, report_date, ["NetSales"], "duration", ["jpy"]),
                choose_value(facts, contexts, report_date, ["OperatingRevenue1"], "duration", ["jpy"]),
                choose_value(facts, contexts, report_date, ["Revenue"], "duration", ["jpy"]),
            ]
        ),
        "cost_of_revenue": pick_first_not_none(
            [
                choose_value(facts, contexts, report_date, ["CostOfSales"], "duration", ["jpy"]),
                choose_value(facts, contexts, report_date, ["CostOfGoodsSold"], "duration", ["jpy"]),
            ]
        ),
        "gross_profit": choose_value(facts, contexts, report_date, ["GrossProfit"], "duration", ["jpy"]),
        # JP filings typically report SG&A instead of a standalone "Operating Expenses" line.
        # Keep canonical key `operating_expenses`, but map it to SG&A concepts.
        "operating_expenses": choose_value(
            facts,
            contexts,
            report_date,
            [
                "SellingGeneralAndAdministrativeExpenses",
                "SellingGeneralAndAdministrativeExpense",
                "SellingAndAdministrativeExpenses",
            ],
            "duration",
            ["jpy"],
        ),
        "operating_income": choose_value(facts, contexts, report_date, ["OperatingIncome"], "duration", ["jpy"]),
        "non_operating_income_expense_net": non_op_net,
        "other_income": pick_first_not_none(
            [
                choose_value(facts, contexts, report_date, ["OtherIncome"], "duration", ["jpy"]),
                choose_value(facts, contexts, report_date, ["OrdinaryIncome"], "duration", ["jpy"]),
            ]
        ),
        "income_before_income_taxes": choose_value(
            facts, contexts, report_date, ["IncomeBeforeIncomeTaxes"], "duration", ["jpy"]
        ),
        "provision_for_income_taxes": pick_first_not_none(
            [
                choose_value(facts, contexts, report_date, ["IncomeTaxes"], "duration", ["jpy"]),
                choose_value(facts, contexts, report_date, ["IncomeTaxesCurrent"], "duration", ["jpy"]),
            ]
        ),
        "net_income": pick_first_not_none(
            [
                choose_value(facts, contexts, report_date, ["ProfitLoss"], "duration", ["jpy"]),
                choose_value(facts, contexts, report_date, ["ProfitLossAttributableToOwnersOfParent"], "duration", ["jpy"]),
            ]
        ),
        "net_income_per_share_basic": choose_value(
            facts, contexts, report_date, ["BasicEarningsPerShare"], "duration", ["jpy/share", "pure"]
        ),
        "net_income_per_share_diluted": choose_value(
            facts, contexts, report_date, ["DilutedEarningsPerShare"], "duration", ["jpy/share", "pure"]
        ),
        "shares_outstanding_basic": choose_value(
            facts, contexts, report_date, ["AverageNumberOfSharesOutstanding"], "duration", ["shares"]
        ),
        "shares_outstanding_diluted": choose_value(
            facts, contexts, report_date, ["AverageNumberOfDilutedSharesOutstanding"], "duration", ["shares"]
        ),
        "anomaly_flag": 0,
        "anomaly_type": None,
        "anomaly_score": None,
        "anomaly_detail": None,
    }


def build_bs_row(symbol: str, filing: FilingInfo, facts: List[dict], contexts: Dict[str, dict]) -> dict:
    report_date = filing.period_end
    cash = pick_first_not_none(
        [
            choose_value(facts, contexts, report_date, ["CashAndDeposits"], "instant", ["jpy"]),
            choose_value(facts, contexts, report_date, ["CashAndCashEquivalents"], "instant", ["jpy"]),
        ]
    )
    short_inv = pick_first_not_none(
        [
            choose_value(facts, contexts, report_date, ["ShortTermInvestmentSecurities"], "instant", ["jpy"]),
            choose_value(facts, contexts, report_date, ["MarketableSecurities"], "instant", ["jpy"]),
            choose_value(facts, contexts, report_date, ["Securities"], "instant", ["jpy"]),
        ]
    )
    # No derived summation: only use explicitly disclosed combined line if present.
    cash_plus_short = pick_first_not_none(
        [
            choose_value(
                facts,
                contexts,
                report_date,
                ["CashAndDepositsAndShortTermInvestmentSecurities"],
                "instant",
                ["jpy"],
            ),
            choose_value(
                facts,
                contexts,
                report_date,
                ["CashAndCashEquivalentsAndShortTermInvestmentSecurities"],
                "instant",
                ["jpy"],
            ),
            choose_value(facts, contexts, report_date, ["CashAndDepositsAndSecurities"], "instant", ["jpy"]),
            choose_value(
                facts,
                contexts,
                report_date,
                ["CashAndCashEquivalentsAndSecurities"],
                "instant",
                ["jpy"],
            ),
        ]
    )

    return {
        "country": "JP",
        "symbol": symbol,
        "company_id": filing.edinet_code,
        "form_type": filing.doc_description or "有価証券報告書",
        "fiscal_year_end_date": report_date,
        "filing_date": filing.filing_date,
        "fiscal_year": int(report_date[:4]),
        "fiscal_period": "FY",
        "source_filing_id": filing.doc_id,
        "is_amendment": 1 if filing.is_amendment else 0,
        "accounting_standard": "JP_GAAP",
        "taxonomy_version": None,
        "statement_scope": "CONSOLIDATED",
        "currency": "JPY",
        "unit_scale": 1,
        "total_assets": choose_value(facts, contexts, report_date, ["Assets"], "instant", ["jpy"]),
        "cash_and_cash_equivalents": cash,
        "short_term_investments": short_inv,
        "total_cash_and_short_term_investments": cash_plus_short,
        "inventories": choose_value(facts, contexts, report_date, ["Inventories"], "instant", ["jpy"]),
        "inventories_finished_goods": pick_first_not_none(
            [
                choose_value(facts, contexts, report_date, ["MerchandiseAndFinishedGoods"], "instant", ["jpy"]),
                choose_value(facts, contexts, report_date, ["FinishedGoods"], "instant", ["jpy"]),
            ]
        ),
        "inventories_work_in_process": choose_value(facts, contexts, report_date, ["WorkInProcess"], "instant", ["jpy"]),
        "inventories_raw_materials_and_supplies": pick_first_not_none(
            [
                choose_value(facts, contexts, report_date, ["RawMaterialsAndSupplies"], "instant", ["jpy"]),
                choose_value(facts, contexts, report_date, ["RawMaterials"], "instant", ["jpy"]),
            ]
        ),
        "accounts_receivable": pick_first_not_none(
            [
                choose_value(
                    facts,
                    contexts,
                    report_date,
                    ["NotesAndAccountsReceivableTradeAndContractAssets"],
                    "instant",
                    ["jpy"],
                ),
                choose_value(
                    facts,
                    contexts,
                    report_date,
                    ["AccountsReceivableTradeAndContractAssets"],
                    "instant",
                    ["jpy"],
                ),
                choose_value(facts, contexts, report_date, ["NotesAndAccountsReceivableTrade"], "instant", ["jpy"]),
                choose_value(facts, contexts, report_date, ["AccountsReceivableTrade"], "instant", ["jpy"]),
            ]
        ),
        "other_current_assets": choose_value(facts, contexts, report_date, ["OtherCurrentAssets"], "instant", ["jpy"]),
        "current_assets_other_jp": choose_value(facts, contexts, report_date, ["OtherCurrentAssets"], "instant", ["jpy"]),
        "total_current_assets": choose_value(facts, contexts, report_date, ["CurrentAssets"], "instant", ["jpy"]),
        "total_noncurrent_assets": pick_first_not_none(
            [
                choose_value(facts, contexts, report_date, ["NoncurrentAssets"], "instant", ["jpy"]),
                choose_value(facts, contexts, report_date, ["FixedAssets"], "instant", ["jpy"]),
            ]
        ),
        "property_plant_and_equipment_net": choose_value(
            facts, contexts, report_date, ["PropertyPlantAndEquipment"], "instant", ["jpy"]
        ),
        "tangible_fixed_assets_total": pick_first_not_none(
            [
                choose_value(facts, contexts, report_date, ["TangibleAssets"], "instant", ["jpy"]),
                choose_value(facts, contexts, report_date, ["PropertyPlantAndEquipment"], "instant", ["jpy"]),
            ]
        ),
        "goodwill": choose_value(facts, contexts, report_date, ["Goodwill"], "instant", ["jpy"]),
        "intangible_assets": choose_value(facts, contexts, report_date, ["IntangibleAssets"], "instant", ["jpy"]),
        "investments_and_other_assets_total": pick_first_not_none(
            [
                choose_value(facts, contexts, report_date, ["InvestmentsAndOtherAssets"], "instant", ["jpy"]),
                choose_value(facts, contexts, report_date, ["InvestmentsOtherAssets"], "instant", ["jpy"]),
            ]
        ),
        "other_noncurrent_assets": choose_value(facts, contexts, report_date, ["OtherNonCurrentAssets"], "instant", ["jpy"]),
        "accounts_payable": pick_first_not_none(
            [
                choose_value(facts, contexts, report_date, ["NotesAndAccountsPayableTrade"], "instant", ["jpy"]),
                choose_value(facts, contexts, report_date, ["AccountsPayableTrade"], "instant", ["jpy"]),
            ]
        ),
        "short_term_debt": pick_first_not_none(
            [
                choose_value(facts, contexts, report_date, ["ShortTermBorrowings"], "instant", ["jpy"]),
                choose_value(facts, contexts, report_date, ["CurrentPortionOfLongTermLoansPayable"], "instant", ["jpy"]),
            ]
        ),
        "other_current_liabilities": choose_value(
            facts, contexts, report_date, ["OtherCurrentLiabilities"], "instant", ["jpy"]
        ),
        "current_liabilities_other_jp": choose_value(
            facts, contexts, report_date, ["OtherCurrentLiabilities"], "instant", ["jpy"]
        ),
        "total_current_liabilities": choose_value(facts, contexts, report_date, ["CurrentLiabilities"], "instant", ["jpy"]),
        "total_noncurrent_liabilities": pick_first_not_none(
            [
                choose_value(facts, contexts, report_date, ["NoncurrentLiabilities"], "instant", ["jpy"]),
                choose_value(facts, contexts, report_date, ["FixedLiabilities"], "instant", ["jpy"]),
            ]
        ),
        "long_term_debt": pick_first_not_none(
            [
                choose_value(facts, contexts, report_date, ["LongTermBorrowings"], "instant", ["jpy"]),
                choose_value(facts, contexts, report_date, ["LongTermLoansPayable"], "instant", ["jpy"]),
            ]
        ),
        "other_noncurrent_liabilities": choose_value(
            facts, contexts, report_date, ["OtherNonCurrentLiabilities"], "instant", ["jpy"]
        ),
        "noncurrent_liabilities_other_jp": choose_value(
            facts, contexts, report_date, ["OtherNonCurrentLiabilities"], "instant", ["jpy"]
        ),
        "total_liabilities": choose_value(facts, contexts, report_date, ["Liabilities"], "instant", ["jpy"]),
        "total_net_assets": choose_value(facts, contexts, report_date, ["NetAssets"], "instant", ["jpy"]),
        "total_shareholders_equity": pick_first_not_none(
            [
                choose_value(facts, contexts, report_date, ["ShareholdersEquity"], "instant", ["jpy"]),
                choose_value(facts, contexts, report_date, ["Equity"], "instant", ["jpy"]),
            ]
        ),
        "accumulated_other_comprehensive_income_or_loss": choose_value(
            facts, contexts, report_date, ["AccumulatedOtherComprehensiveIncome"], "instant", ["jpy"]
        ),
        "share_subscription_rights": pick_first_not_none(
            [
                choose_value(facts, contexts, report_date, ["SubscriptionRightsToShares"], "instant", ["jpy"]),
                choose_value(facts, contexts, report_date, ["ShareSubscriptionRights"], "instant", ["jpy"]),
            ]
        ),
        "common_stock": choose_value(facts, contexts, report_date, ["CapitalStock"], "instant", ["jpy"]),
        "additional_paid_in_capital": choose_value(
            facts, contexts, report_date, ["CapitalSurplus"], "instant", ["jpy"]
        ),
        "retained_earnings": choose_value(facts, contexts, report_date, ["RetainedEarnings"], "instant", ["jpy"]),
        "noncontrolling_interests": choose_value(
            facts, contexts, report_date, ["NonControllingInterests"], "instant", ["jpy"]
        ),
        "total_liabilities_and_shareholders_equity": pick_first_not_none(
            [
                choose_value(facts, contexts, report_date, ["LiabilitiesAndNetAssets"], "instant", ["jpy"]),
                choose_value(facts, contexts, report_date, ["LiabilitiesAndEquity"], "instant", ["jpy"]),
            ]
        ),
        "anomaly_flag": 0,
        "anomaly_type": None,
        "anomaly_score": None,
        "anomaly_detail": None,
    }


def build_cf_row(symbol: str, filing: FilingInfo, facts: List[dict], contexts: Dict[str, dict]) -> dict:
    report_date = filing.period_end
    prior_end = infer_prior_period_end(contexts, report_date)

    net_cash_operating = choose_value(
        facts,
        contexts,
        report_date,
        ["NetCashProvidedByUsedInOperatingActivities", "NetCashProvidedByUsedInOperatingActivitiesIFRS"],
        "duration",
        ["jpy"],
    )
    net_cash_investing = choose_value(
        facts,
        contexts,
        report_date,
        ["NetCashProvidedByUsedInInvestingActivities", "NetCashProvidedByUsedInInvestingActivitiesIFRS"],
        "duration",
        ["jpy"],
    )
    net_cash_financing = choose_value(
        facts,
        contexts,
        report_date,
        ["NetCashProvidedByUsedInFinancingActivities", "NetCashProvidedByUsedInFinancingActivitiesIFRS"],
        "duration",
        ["jpy"],
    )
    fx_effect = choose_value(
        facts,
        contexts,
        report_date,
        [
            "EffectOfExchangeRateChangeOnCashAndCashEquivalents",
            "EffectOfExchangeRateChangesOnCashAndCashEquivalentsIFRS",
        ],
        "duration",
        ["jpy"],
    )
    net_change_cash = choose_value(
        facts,
        contexts,
        report_date,
        ["NetIncreaseDecreaseInCashAndCashEquivalents", "NetIncreaseDecreaseInCashAndCashEquivalentsIFRS"],
        "duration",
        ["jpy"],
    )

    cash_begin = pick_first_not_none(
        [
            choose_value(facts, contexts, report_date, ["CashAndCashEquivalentsAtBeginningOfPeriod"], "duration", ["jpy"]),
            choose_value_at_date(
                facts,
                contexts,
                prior_end,
                ["CashAndCashEquivalents", "CashAndCashEquivalentsIFRS"],
                "instant",
                ["jpy"],
            )
            if prior_end
            else None,
        ]
    )
    cash_end = pick_first_not_none(
        [
            choose_value(facts, contexts, report_date, ["CashAndCashEquivalentsAtEndOfPeriod"], "duration", ["jpy"]),
            choose_value(
                facts,
                contexts,
                report_date,
                ["CashAndCashEquivalents", "CashAndCashEquivalentsIFRS"],
                "instant",
                ["jpy"],
            ),
        ]
    )
    if cash_begin is None and (cash_end is not None) and (net_change_cash is not None):
        cash_begin = cash_end - net_change_cash

    return {
        "country": "JP",
        "symbol": symbol,
        "company_id": filing.edinet_code,
        "form_type": filing.doc_description or "有価証券報告書",
        "fiscal_year_end_date": report_date,
        "filing_date": filing.filing_date,
        "fiscal_year": int(report_date[:4]),
        "fiscal_period": "FY",
        "source_filing_id": filing.doc_id,
        "is_amendment": 1 if filing.is_amendment else 0,
        "accounting_standard": "JP_GAAP",
        "taxonomy_version": None,
        "statement_scope": "CONSOLIDATED",
        "currency": "JPY",
        "unit_scale": 1,
        "net_income": pick_first_not_none(
            [
                choose_value(facts, contexts, report_date, ["ProfitLoss"], "duration", ["jpy"]),
                choose_value(facts, contexts, report_date, ["ProfitLossAttributableToOwnersOfParent"], "duration", ["jpy"]),
            ]
        ),
        "net_cash_operating": net_cash_operating,
        "net_cash_investing": net_cash_investing,
        "net_cash_financing": net_cash_financing,
        "effect_of_exchange_rates_on_cash": fx_effect,
        "net_change_in_cash": net_change_cash,
        "cash_beginning_of_period": cash_begin,
        "cash_end_of_period": cash_end,
        "anomaly_flag": 0,
        "anomaly_type": None,
        "anomaly_score": None,
        "anomaly_detail": None,
    }


def export_three_statements_excel(
    symbol: Optional[str],
    company_name: Optional[str],
    edinet_code: Optional[str],
    api_key: Optional[str],
    as_of_date: str,
    lookback_days: int,
    out_path: Path,
    xbrl_zip_path: Optional[Path] = None,
    report_date: Optional[str] = None,
    filing_date: Optional[str] = None,
    form_type: Optional[str] = None,
    source_filing_id: Optional[str] = None,
) -> Tuple[Path, FilingInfo]:
    xbrl_name = ""
    if xbrl_zip_path is not None:
        xbrl_name, xbrl_bytes = load_xbrl_source(xbrl_zip_path)
        print(f"[INFO] Local XBRL source: {xbrl_zip_path}")
        print(f"[INFO] XBRL instance: {xbrl_name}")
        facts, contexts = parse_xbrl(xbrl_bytes)
        print(f"[INFO] Parsed facts={len(facts)} contexts={len(contexts)}")
        filing = build_local_filing_info(
            source_path=xbrl_zip_path,
            symbol=symbol,
            company_name=company_name,
            edinet_code=edinet_code,
            report_date=report_date,
            filing_date=filing_date,
            form_type=form_type,
            source_filing_id=source_filing_id,
            facts=facts,
            contexts=contexts,
        )
        print(
            f"[INFO] Local filing metadata: docID={filing.doc_id} symbol={filing.sec_code} "
            f"filer={filing.filer_name} period_end={filing.period_end} filing_date={filing.filing_date}"
        )
    else:
        if not api_key:
            raise RuntimeError("EDINET key is required in API mode.")
        if not symbol:
            raise RuntimeError("--symbol is required in API mode.")

        session = requests.Session()
        session.headers.update(
            {
                "User-Agent": "FinReporting-JP/1.0 (research/demo)",
                "Accept-Encoding": "gzip, deflate",
            }
        )

        filing = find_latest_annual_filing(
            session=session,
            api_key=api_key,
            symbol=symbol,
            company_name=company_name,
            edinet_code=edinet_code,
            as_of_date=as_of_date,
            lookback_days=lookback_days,
        )

        print(
            f"[INFO] Filing selected: docID={filing.doc_id} symbol={filing.sec_code} "
            f"filer={filing.filer_name} period_end={filing.period_end} filing_date={filing.filing_date}"
        )
        zip_blob = request_binary(
            session, f"/documents/{filing.doc_id}", params={"type": "1"}, api_key=api_key
        )
        xbrl_name, xbrl_bytes = pick_instance_xbrl_from_zip(zip_blob)
        print(f"[INFO] XBRL instance: {xbrl_name}")
        facts, contexts = parse_xbrl(xbrl_bytes)
        print(f"[INFO] Parsed facts={len(facts)} contexts={len(contexts)}")

    normalized_symbol = normalize_sec_code(symbol) or filing.sec_code or symbol or "UNKNOWN"
    is_row = build_is_row(normalized_symbol, filing, facts, contexts)
    bs_row = build_bs_row(normalized_symbol, filing, facts, contexts)
    cf_row = build_cf_row(normalized_symbol, filing, facts, contexts)

    out_path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        pd.DataFrame([is_row]).to_excel(writer, index=False, sheet_name="JP_FIN_IS")
        pd.DataFrame([bs_row]).to_excel(writer, index=False, sheet_name="JP_FIN_BS")
        pd.DataFrame([cf_row]).to_excel(writer, index=False, sheet_name="JP_FIN_CF")

    return out_path, filing


def main() -> None:
    args = parse_args()
    local_mode = bool(args.xbrl_zip)
    if not local_mode and not args.edinet_key:
        raise SystemExit("EDINET key is required in API mode. Set --edinet-key or EDINET_API_KEY.")
    if not local_mode and not args.symbol:
        raise SystemExit("--symbol is required in API mode.")

    out_path = Path(args.out)
    try:
        exported, filing = export_three_statements_excel(
            symbol=normalize_sec_code(args.symbol) if args.symbol else None,
            company_name=args.company_name,
            edinet_code=args.edinet_code,
            api_key=args.edinet_key,
            as_of_date=args.as_of_date,
            lookback_days=args.lookback_days,
            out_path=out_path,
            xbrl_zip_path=Path(args.xbrl_zip) if args.xbrl_zip else None,
            report_date=args.report_date,
            filing_date=args.filing_date,
            form_type=args.form_type,
            source_filing_id=args.source_filing_id,
        )
    except Exception as e:
        raise SystemExit(f"JP extraction failed: {e}")

    print("[OK] Excel exported:", exported)
    print(
        "[OK] Filing metadata:",
        json.dumps(
            {
                "doc_id": filing.doc_id,
                "filer_name": filing.filer_name,
                "sec_code": filing.sec_code,
                "period_end": filing.period_end,
                "filing_date": filing.filing_date,
                "form_code": filing.form_code,
                "doc_type_code": filing.doc_type_code,
            },
            ensure_ascii=False,
        ),
    )


if __name__ == "__main__":
    main()
