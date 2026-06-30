#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Export CN Income Statement / Balance Sheet / Cash Flow to a single Excel file.

Data source:
  1) CNINFO latest annual report PDF auto-download (default), or
  2) Local annual report PDF path.

Usage:
  ./venv/bin/python CN/export_three_statements_excel_cn.py \
    --symbol 300750 \
    --out CN/catl_3statements.xlsx
"""

from __future__ import annotations

import argparse
import json
import math
import os
import re
import sys
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Dict, List, Optional

import pandas as pd
import pdfplumber
import requests

# Ensure project root is importable when running this script from CN/.
ROOT_DIR = Path(__file__).resolve().parent.parent
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

from extract_consolidated_statements import run_pipeline
from llm_cn_verifier import (
    DEFAULT_BASE_URL as LLM_DEFAULT_BASE_URL,
    DEFAULT_MODEL as LLM_DEFAULT_MODEL,
    audit_cn_workbook,
    append_audit_sheet,
)


CNINFO_QUERY_URL = "http://www.cninfo.com.cn/new/hisAnnouncement/query"
CNINFO_STOCK_LIST_URL = "http://www.cninfo.com.cn/new/data/szse_stock.json"
CNINFO_STATIC_BASE = "http://static.cninfo.com.cn/"
REQUEST_TIMEOUT = 45


@dataclass
class CnAnnouncement:
    symbol: str
    org_id: str
    company_name: str
    title: str
    adjunct_url: str
    announcement_id: str
    announcement_time_ms: int
    fiscal_year: int

    @property
    def filing_date(self) -> str:
        return datetime.fromtimestamp(self.announcement_time_ms / 1000).date().isoformat()

    @property
    def fiscal_year_end_date(self) -> str:
        return f"{self.fiscal_year:04d}-12-31"

    @property
    def full_pdf_url(self) -> str:
        return CNINFO_STATIC_BASE + self.adjunct_url.lstrip("/")


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(description="Export CN IS/BS/CF from CNINFO annual report PDF.")
    p.add_argument("--symbol", required=True, help="A-share stock code, e.g. 300750 / 002415 / 600519")
    p.add_argument("--company-name", help="Optional expected company name keyword")
    p.add_argument("--pdf", help="Optional local annual report PDF path (skip CNINFO download)")
    p.add_argument("--schema-file", default="schemas/CN_Schemas.xlsx", help="CN schema file path")
    p.add_argument("--as-of-date", default=date.today().isoformat(), help="Anchor date (YYYY-MM-DD)")
    p.add_argument("--lookback-years", type=int, default=2, help="How many years back to search annual report")
    p.add_argument("--raw-out", default=None, help="Optional path for intermediate extraction workbook")
    p.add_argument("--out", default="CN/cn_3statements.xlsx", help="Output Excel path")
    p.add_argument("--llm-verify", action="store_true", help="Run LLM verify/repair on exported core fields")
    p.add_argument("--llm-model", default=LLM_DEFAULT_MODEL, help="Fixed model id for reproducibility")
    p.add_argument("--llm-base-url", default=LLM_DEFAULT_BASE_URL, help="OpenAI API base URL")
    p.add_argument("--llm-api-key", default=os.getenv("OPENAI_API_KEY"), help="API key (prefer env OPENAI_API_KEY)")
    p.add_argument("--llm-max-fields", type=int, default=0, help="Only process first N core fields (0 = all)")
    p.add_argument("--llm-audit-csv", default=None, help="Optional LLM audit CSV output path")
    p.add_argument(
        "--llm-allow-repair-all-status",
        action="store_true",
        help="Allow LLM REPAIR beyond MISSING/PARSE_ERROR (default disabled)",
    )
    return p.parse_args()


def cninfo_headers() -> Dict[str, str]:
    return {
        "User-Agent": "Mozilla/5.0",
        "X-Requested-With": "XMLHttpRequest",
        "Referer": "http://www.cninfo.com.cn/new/commonUrl/pageOfSearch?url=disclosure/list/search",
    }


def normalize_symbol(symbol: str) -> str:
    s = re.sub(r"\D", "", symbol or "")
    if len(s) != 6:
        raise ValueError(f"Invalid CN symbol: {symbol}")
    return s


def load_stock_meta(session: requests.Session, symbol: str) -> dict:
    r = session.get(CNINFO_STOCK_LIST_URL, headers=cninfo_headers(), timeout=REQUEST_TIMEOUT)
    r.raise_for_status()
    data = r.json()
    stock_list = data.get("stockList") or []
    for row in stock_list:
        if str(row.get("code", "")).strip() == symbol:
            return row
    raise RuntimeError(f"Symbol not found in CNINFO stock list: {symbol}")


def query_annual_announcements(
    session: requests.Session,
    symbol: str,
    org_id: str,
    lookback_start: str,
    lookback_end: str,
) -> List[dict]:
    plate = "sh" if symbol.startswith(("6", "9")) else "sz"
    column = "sse" if plate == "sh" else "szse"
    payload = {
        "pageNum": "1",
        "pageSize": "50",
        "column": column,
        "tabName": "fulltext",
        "plate": plate,
        "stock": f"{symbol},{org_id}",
        "searchkey": "",
        "secid": "",
        "category": "category_ndbg_szsh",
        "trade": "",
        "seDate": f"{lookback_start}~{lookback_end}",
        "sortName": "",
        "sortType": "desc",
        "isHLtitle": "true",
    }
    r = session.post(CNINFO_QUERY_URL, headers=cninfo_headers(), data=payload, timeout=REQUEST_TIMEOUT)
    r.raise_for_status()
    data = r.json()
    anns = data.get("announcements") or []
    if not isinstance(anns, list):
        return []
    return anns


def parse_fiscal_year_from_title(title: str) -> Optional[int]:
    if not title:
        return None
    # Support both "2024年年度报告" and "2024年度报告".
    m = re.search(r"(20\d{2})\s*年?\s*年度报告", title)
    if m:
        return int(m.group(1))
    return None


def pick_latest_annual_report(
    symbol: str,
    org_id: str,
    company_name: str,
    announcements: List[dict],
    company_name_hint: Optional[str] = None,
) -> CnAnnouncement:
    candidates: List[CnAnnouncement] = []
    for a in announcements:
        title = str(a.get("announcementTitle", "")).strip()
        adjunct = str(a.get("adjunctUrl", "")).strip()
        if not title or not adjunct:
            continue
        if ("年度报告" not in title) or ("摘要" in title) or ("英文" in title):
            continue

        fiscal_year = parse_fiscal_year_from_title(title)
        if fiscal_year is None:
            continue

        if company_name_hint and company_name_hint not in title and company_name_hint not in company_name:
            # Keep as fallback; don't hard fail.
            pass

        announcement_id = str(a.get("announcementId", "")).strip()
        ann_time = int(a.get("announcementTime", 0) or 0)
        if not announcement_id or ann_time <= 0:
            continue

        candidates.append(
            CnAnnouncement(
                symbol=symbol,
                org_id=org_id,
                company_name=company_name,
                title=title,
                adjunct_url=adjunct,
                announcement_id=announcement_id,
                announcement_time_ms=ann_time,
                fiscal_year=fiscal_year,
            )
        )

    if not candidates:
        raise RuntimeError(f"No annual report PDF found on CNINFO for symbol={symbol}")

    # Latest by announcement time.
    candidates.sort(key=lambda x: x.announcement_time_ms, reverse=True)
    return candidates[0]


def download_pdf(session: requests.Session, url: str, out_path: Path) -> Path:
    out_path.parent.mkdir(parents=True, exist_ok=True)
    with session.get(url, headers={"User-Agent": "Mozilla/5.0"}, timeout=REQUEST_TIMEOUT, stream=True) as r:
        r.raise_for_status()
        with out_path.open("wb") as f:
            for chunk in r.iter_content(chunk_size=1024 * 128):
                if chunk:
                    f.write(chunk)
    return out_path


def load_value_map(raw_extract_xlsx: Path, sheet_name: str) -> Dict[str, float]:
    df = pd.read_excel(raw_extract_xlsx, sheet_name=sheet_name)
    df = df[["item_code", "value"]].copy()
    df = df[df["item_code"].notna()]
    out: Dict[str, float] = {}
    for _, row in df.iterrows():
        code = str(row["item_code"]).strip()
        val = row["value"]
        if not code or pd.isna(val):
            continue
        try:
            out[code] = float(val)
        except Exception:
            continue
    return out


def load_status_map(raw_extract_xlsx: Path, sheet_name: str) -> Dict[str, str]:
    df = pd.read_excel(raw_extract_xlsx, sheet_name=sheet_name)
    if "item_code" not in df.columns:
        return {}
    if "status" not in df.columns:
        return {}
    out: Dict[str, str] = {}
    for _, row in df.iterrows():
        code = str(row.get("item_code", "")).strip()
        status = str(row.get("status", "")).strip().upper()
        if not code or not status:
            continue
        out[code] = status
    return out


def load_raw_text_map(raw_extract_xlsx: Path, sheet_name: str) -> Dict[str, str]:
    df = pd.read_excel(raw_extract_xlsx, sheet_name=sheet_name)
    if "item_code" not in df.columns:
        return {}
    text_col = "raw_text" if "raw_text" in df.columns else None
    if text_col is None:
        return {}
    out: Dict[str, str] = {}
    for _, row in df.iterrows():
        code = str(row.get("item_code", "")).strip()
        raw_text = str(row.get(text_col, "")).strip()
        if not code or not raw_text:
            continue
        out[code] = raw_text
    return out


def pick_first(vmap: Dict[str, float], codes: List[str]) -> Optional[float]:
    for c in codes:
        if c in vmap:
            return vmap[c]
    return None


def sum_values(*vals: Optional[float]) -> Optional[float]:
    xs = [v for v in vals if v is not None]
    if not xs:
        return None
    return float(sum(xs))


def _parse_reported_unit_from_line(line: str) -> Optional[str]:
    if not line:
        return None
    s = str(line).strip()
    if "单位" not in s:
        return None

    m = re.search(r"单位\s*[:：]?\s*([^\n]+)", s)
    if not m:
        return None

    unit = m.group(1).strip()
    unit = re.split(r"[，,。；;]", unit)[0].strip()
    if not unit:
        return None
    if len(unit) > 40:
        unit = unit[:40].strip()

    # Normalize to canonical CN unit strings.
    canonical_order = [
        "人民币亿元",
        "人民币百万元",
        "人民币万元",
        "人民币千元",
        "人民币元",
    ]
    for c in canonical_order:
        if c in unit:
            return c

    unit = unit.strip("()（） ").lstrip("为")
    if unit in {"亿元", "百万元", "万元", "千元", "元"}:
        return f"人民币{unit}"
    if unit in {"RMB千元", "RMB万元", "RMB元"}:
        return unit.replace("RMB", "人民币")
    if re.fullmatch(r"(人民币)?(元|千元|万元|百万元|亿元)", unit):
        if unit.startswith("人民币"):
            return unit
        return f"人民币{unit}"
    return None


def extract_reported_unit_from_pdf(pdf_path: Path) -> str:
    """
    Extract reported unit from statement header, e.g. "单位：人民币千元".
    """
    unit_candidates: List[str] = []

    with pdfplumber.open(str(pdf_path)) as pdf:
        page_texts: List[str] = [(pg.extract_text() or "") for pg in pdf.pages]

        ranges: List[tuple[int, int]] = []
        bs_range = _find_statement_page_range(
            page_texts,
            start_pattern=r"合并(?:及公司)?资产负债表",
            end_pattern=r"合并(?:及公司)?利润表|合并(?:及公司)?现金流量表",
        )
        if bs_range is not None:
            ranges.append(bs_range)
        is_range = _find_statement_page_range(
            page_texts,
            start_pattern=r"合并(?:及公司)?利润表",
            end_pattern=r"合并(?:及公司)?现金流量表|合并股东权益变动表",
        )
        if is_range is not None:
            ranges.append(is_range)
        cf_range = _find_statement_page_range(
            page_texts,
            start_pattern=r"合并(?:及公司)?现金流量表",
            end_pattern=r"合并股东权益变动表",
        )
        if cf_range is not None:
            ranges.append(cf_range)

        candidate_pages: List[int] = []
        for s, _ in ranges:
            candidate_pages.extend([s, min(s + 1, len(page_texts) - 1)])
        if not candidate_pages:
            candidate_pages = list(range(min(30, len(page_texts))))

        for p in sorted(set(candidate_pages)):
            text = page_texts[p]
            for line in text.split("\n"):
                unit = _parse_reported_unit_from_line(line)
                if unit:
                    unit_candidates.append(unit)

    if unit_candidates:
        # Prefer the first unit seen near statement headers.
        return unit_candidates[0]
    return "人民币元"


def _parse_amount_token_candidates(tok: str) -> List[float]:
    out: List[float] = []
    s = tok.strip()
    neg = s.startswith("(") and s.endswith(")")
    s = s.strip("()")

    if "," in s:
        parts = s.split(",")
        if not all(p.isdigit() for p in parts):
            return out
        first = parts[0]
        tails = parts[1:]
        if not tails or not all(len(x) == 3 for x in tails):
            return out

        first_options: List[str] = [first]
        # Common CN PDF issue: note number (e.g. 44/58/61) gets stuck to amount head.
        if len(first) >= 3:
            for cut in (2, 1):
                f = first[cut:]
                if 1 <= len(f) <= 3:
                    first_options.append(f)

        for f in first_options:
            try:
                v = float("".join([f] + tails))
            except Exception:
                continue
            if neg and v > 0:
                v = -v
            out.append(v)
        return out

    # Plain numeric token.
    if re.fullmatch(r"-?\d+(?:\.\d+)?", s):
        try:
            v = float(s)
            if neg and v > 0:
                v = -v
            out.append(v)
        except Exception:
            pass
    return out


def _pick_most_plausible(first_vals: List[float], ref_vals: List[float]) -> Optional[float]:
    if not first_vals:
        return None
    if not ref_vals:
        # No reference magnitude: prefer larger non-trivial value.
        return max(first_vals, key=lambda x: abs(x))
    ref = max(ref_vals, key=lambda x: abs(x))
    if ref == 0:
        return max(first_vals, key=lambda x: abs(x))

    def score(v: float) -> float:
        # Smaller score is better (closer magnitude to reference value).
        return abs(math.log10((abs(v) + 1.0) / (abs(ref) + 1.0)))

    return min(first_vals, key=score)


def _parse_first_amount_from_line(line: str) -> Optional[float]:
    # Capture money-like tokens with comma groups first (supports sticky note numbers).
    comma_tokens = re.findall(r"\(?-?\d{1,6}(?:,\d{3}){1,4}\)?", line)
    if comma_tokens:
        first_cands = _parse_amount_token_candidates(comma_tokens[0])
        ref_cands: List[float] = []
        if len(comma_tokens) > 1:
            ref_cands = _parse_amount_token_candidates(comma_tokens[1])
        picked = _pick_most_plausible(first_cands, ref_cands)
        if picked is not None:
            return picked

    # Fallback: large plain integers.
    plain_tokens = re.findall(r"\(?-?\d{4,}\)?", line)
    if plain_tokens:
        cands = _parse_amount_token_candidates(plain_tokens[0])
        if cands:
            return cands[0]
    return None


def _parse_first_eps_from_line(line: str) -> Optional[float]:
    # EPS lines are decimals; support OCR variants:
    # 8.2062 / 8,2062 / 8．2062 and glued forms like 8.20627.7830.
    tokens = re.findall(r"\(?-?\d+[.,，．]\d{1,4}\)?", line)
    if not tokens:
        return None

    def eps_cands(tok: str) -> List[float]:
        out: List[float] = []
        s = tok.strip()
        neg = s.startswith("(") and s.endswith(")")
        s = s.strip("()")
        s = s.replace("，", ",").replace("．", ".")
        if "," in s and "." not in s:
            s = s.replace(",", ".")
        try:
            v = float(s)
            if neg and v > 0:
                v = -v
            if abs(v) <= 200:
                out.append(v)
        except Exception:
            pass
        # Try stripping sticky note prefix from integer part (e.g. 5913.84 -> 13.84).
        if "." in s:
            left, right = s.split(".", 1)
            if len(left) >= 3 and (not out):
                for cut in (2, 1):
                    l2 = left[cut:]
                    if not l2:
                        continue
                    try:
                        v2 = float(f"{l2}.{right}")
                        if neg and v2 > 0:
                            v2 = -v2
                        if abs(v2) <= 200:
                            out.append(v2)
                    except Exception:
                        continue
        return out

    # CN statements place current period first; return first plausible token.
    for tok in tokens:
        cands = eps_cands(tok)
        if cands:
            return cands[0]
    return None


def _is_not_applicable_text(line: str) -> bool:
    if not line:
        return False
    s = re.sub(r"\s+", "", str(line)).upper()
    if "不适用" in s or "不存在" in s:
        return True
    if re.search(r"(?<![A-Z])N/?A(?![A-Z])", s):
        return True
    if "--" in s or "—" in s:
        return True
    return False


def _na_before_first_eps_number(line: str) -> bool:
    if not line:
        return False
    s = re.sub(r"\s+", "", str(line))
    num_m = re.search(r"\d+[.,，．]\d{1,4}", s)
    num_pos = num_m.start() if num_m else -1
    na_pos = -1
    for tok in ("不适用", "不存在", "N/A", "NA", "—", "--"):
        p = s.upper().find(tok.upper())
        if p >= 0 and (na_pos < 0 or p < na_pos):
            na_pos = p
    if na_pos < 0:
        return False
    if num_pos < 0:
        return True
    return na_pos < num_pos


def _is_table_like_statement_page(text: str) -> bool:
    if not text:
        return False
    t = re.sub(r"\s+", "", text)
    if "审计报告" in t and "审字" in t:
        return False
    if not re.search(r"(项目|附注|本年|本期|期末|年度)", t):
        return False
    digits = len(re.findall(r"\d", t))
    return digits > 40


def _find_statement_page_range(
    page_texts: List[str],
    start_pattern: str,
    end_pattern: str,
) -> Optional[tuple]:
    normalized = [re.sub(r"\s+", "", x or "") for x in page_texts]
    start_idx: Optional[int] = None
    end_idx: Optional[int] = None
    for i, text in enumerate(normalized):
        if start_idx is None:
            if re.search(start_pattern, text) and _is_table_like_statement_page(page_texts[i]):
                start_idx = i
            continue
        if re.search(end_pattern, text):
            end_idx = i - 1
            break
    if start_idx is None:
        return None
    if end_idx is None:
        end_idx = len(page_texts) - 1
    end_idx = max(end_idx, start_idx)
    return start_idx, end_idx


def enrich_is_map_from_pdf_text(
    pdf_path: Path, is_map: Dict[str, float]
) -> tuple[Dict[str, float], Dict[str, str]]:
    """
    Fill missing key IS items from consolidated income-statement text lines.
    """
    out = dict(is_map)
    source_labels: Dict[str, str] = {}
    target_codes = {
        "归属于母公司所有者的净利润": "PARENETP",
        "归属于母公司股东的净利润": "PARENETP",
        "营业收入": "BIZINCO",
        "营业成本": "BIZCOST",
        "税金及附加": "BIZTAX",
        "销售费用": "SALESEXPE",
        "管理费用": "MANAEXPE",
        "研发费用": "DEVEEXPE",
        "财务费用": "FINEXPE",
        "财务收入": "FINEXPE",
        "稀释每股收益": "DILUTEDEPS",
        "基本每股收益": "BASICEPS",
        "利润总额": "TOTPROFIT",
        "营业外收入": "NONOREVE",
        "营业外支出": "NONOEXPE",
        "所得税费用": "INCOTAXEXPE",
        "营业利润": "PERPROFIT",
        "其他收益": "OTHERINCO",
        "净利润": "NETPROFIT",
    }
    ordered_labels = sorted(target_codes.keys(), key=len, reverse=True)

    with pdfplumber.open(str(pdf_path)) as pdf:
        page_texts: List[str] = [(pg.extract_text() or "") for pg in pdf.pages]
        page_range = _find_statement_page_range(
            page_texts,
            start_pattern=r"合并(?:及公司)?利润表",
            end_pattern=r"合并(?:及公司)?现金流量表|合并股东权益变动表",
        )
        if page_range is None:
            return out, source_labels
        start_idx, end_idx = page_range

        for text in page_texts[start_idx : end_idx + 1]:
            for line in text.split("\n"):
                line_s = re.sub(r"\s+", "", line.strip())
                if not line_s:
                    continue
                for label in ordered_labels:
                    code = target_codes[label]
                    # EPS can be mis-parsed in table extraction (e.g. 7,8300 -> 7830),
                    # so allow text-based override for EPS rows.
                    is_eps = code in {"BASICEPS", "DILUTEDEPS"}
                    status_key = f"{code}_status"
                    if is_eps and source_labels.get(status_key) == "NOT_APPLICABLE":
                        continue
                    if code in out and not is_eps:
                        continue
                    if label not in line_s:
                        continue
                    if is_eps:
                        if _na_before_first_eps_number(line_s):
                            source_labels[status_key] = "NOT_APPLICABLE"
                            continue
                        val = _parse_first_eps_from_line(line_s)
                        if val is None and _is_not_applicable_text(line_s):
                            source_labels[status_key] = "NOT_APPLICABLE"
                            continue
                    else:
                        val = _parse_first_amount_from_line(line_s)
                    if val is not None:
                        out[code] = val
                        if is_eps and source_labels.get(status_key) == "NOT_APPLICABLE":
                            source_labels.pop(status_key, None)
                        if code == "FINEXPE" and label in {"财务费用", "财务收入"}:
                            source_labels["finance_result_source_label"] = label
    return out, source_labels


def enrich_cf_map_from_pdf_text(pdf_path: Path, cf_map: Dict[str, float]) -> Dict[str, float]:
    """
    Fill missing key CF items from consolidated cash-flow statement text lines.
    """
    out = dict(cf_map)
    line_patterns = [
        (r"经营活动(?:产生|使用)的现金流量净额", "MANANETR"),
        (r"投资活动(?:产生|使用)(?:/\(使用\))?的现金流量净额", "INVNETCASHFLOW"),
        (r"筹资活动(?:产生|使用)(?:/\(使用\))?的现金流量净额", "FINNETCFLOW"),
        (r"汇率变动对现金及现金等价物的影响", "CHGEXCHGCHGS"),
        (r"现金及现金等价物净.*增加额", "CASHNETR"),
        (r"(?:年初|期初)现金及现金等价物余额", "INICASHBALA"),
        (r"(?:年末|期末)现金及现金等价物余额", "FINALCASHBALA"),
    ]

    with pdfplumber.open(str(pdf_path)) as pdf:
        page_texts: List[str] = [(pg.extract_text() or "") for pg in pdf.pages]
        page_range = _find_statement_page_range(
            page_texts,
            start_pattern=r"合并(?:及公司)?现金流量表",
            end_pattern=r"合并股东权益变动表",
        )
        if page_range is None:
            return out
        start_idx, end_idx = page_range

        for text in page_texts[start_idx : end_idx + 1]:
            for line in text.split("\n"):
                line_s = re.sub(r"\s+", "", line.strip())
                if not line_s:
                    continue
                for pat, code in line_patterns:
                    if code in out:
                        continue
                    if re.search(pat, line_s):
                        val = _parse_first_amount_from_line(line_s)
                        if val is not None:
                            out[code] = val
    return out


def _expense_abs(v: Optional[float]) -> Optional[float]:
    if v is None:
        return None
    return abs(v) if v < 0 else v


def _extract_bs_total_assets_from_lines(lines: List[str]) -> Optional[float]:
    """
    Strict TOTAL_ASSETS fallback from balance-sheet text lines.
    Priority:
      1) 资产总计
      2) 资产总额
      3) 总资产
    Allow 资产合计 only when it appears as a final grand-total row.
    """
    candidates: List[tuple[int, int, float]] = []
    n = len(lines)

    for idx, line in enumerate(lines):
        s = re.sub(r"\s+", "", line or "")
        if not s:
            continue

        # Exclude section subtotals/sub-items.
        if any(x in s for x in ("流动资产合计", "流动资产总计", "非流动资产合计", "非流动资产总计")):
            continue
        if "其中" in s:
            continue

        priority: Optional[int] = None
        if "资产总计" in s:
            priority = 1
        elif "资产总额" in s:
            priority = 2
        elif "总资产" in s:
            priority = 3
        elif "资产合计" in s:
            # Accept only as a final grand-total style row.
            if n > 0 and idx >= int(n * 0.6):
                priority = 4

        if priority is None:
            continue

        v = _parse_first_amount_from_line(s)
        if v is None:
            continue
        candidates.append((priority, idx, v))

    if not candidates:
        return None

    # For explicit labels (1-3), prefer earliest occurrence;
    # for 资产合计 fallback, prefer latest occurrence.
    candidates.sort(key=lambda x: (x[0], -x[1] if x[0] == 4 else x[1]))
    return candidates[0][2]


def _extract_bs_total_liabilities_from_lines(lines: List[str]) -> Optional[float]:
    """
    Strict TOTAL_LIABILITIES fallback from balance-sheet text lines.
    Require BS-context anchors to avoid note-table contamination.
    """
    n = len(lines)
    candidates: List[tuple[int, float]] = []
    for idx, line in enumerate(lines):
        s = re.sub(r"\s+", "", line or "")
        if not s:
            continue
        if "负债合计" not in s:
            continue
        # Exclude note/fair-value hierarchy rows.
        if any(x in s for x in ("第一层次", "第二层次", "第三层次", "公允价值")):
            continue

        prev = [re.sub(r"\s+", "", x or "") for x in lines[max(0, idx - 8) : idx]]
        nxt = [re.sub(r"\s+", "", x or "") for x in lines[idx + 1 : min(n, idx + 12)]]
        cond_prev = any("非流动负债合计" in x for x in prev)
        cond_next = any(("股东权益" in x) or ("所有者权益" in x) for x in nxt)
        if not (cond_prev and cond_next):
            continue

        v = _parse_first_amount_from_line(s)
        if v is None:
            continue
        candidates.append((idx, v))

    if not candidates:
        return None
    candidates.sort(key=lambda x: x[0])
    return candidates[0][1]


def enrich_bs_map_from_pdf_text(pdf_path: Path, bs_map: Dict[str, float]) -> Dict[str, float]:
    """
    Fill missing key BS items from raw page text lines.
    Useful when table extraction misses equity block on some CN PDFs.
    """
    out = dict(bs_map)
    target_codes = {
        "资产总计": "TOTASSET",
        "资产总额": "TOTASSET",
        "总资产": "TOTASSET",
        "负债合计": "TOTLIAB",
        "归属于母公司股东权益合计": "PARESHARRIGH",
        "少数股东权益": "MINYSHARRIGH",
        "股东权益合计": "RIGHAGGR",
        "所有者权益(或股东权益)合计": "RIGHAGGR",
        "负债和股东权益总计": "TOTLIABSHAREQUI",
        "负债和所有者权益(或股东权益)总计": "TOTLIABSHAREQUI",
        "股本": "PAIDINCAPI",
        "资本公积": "CAPISURP",
        "未分配利润": "UNDIPROF",
        "其他综合收益": "OTHERCOMPINCO",
    }
    # Longer/more specific keys first to avoid partial-match collisions.
    ordered_labels = sorted(target_codes.keys(), key=len, reverse=True)

    with pdfplumber.open(str(pdf_path)) as pdf:
        page_texts: List[str] = [(pg.extract_text() or "") for pg in pdf.pages]

        # Prefer robust statement-page routing to avoid note-table contamination.
        page_range = _find_statement_page_range(
            page_texts,
            start_pattern=r"(?:合并|并)(?:及公司)?资产负债表",
            end_pattern=r"(?:合并|并)(?:及公司)?利润表|(?:合并|并)(?:及公司)?现金流量表|合并股东权益变动表",
        )

        start_idx: Optional[int]
        end_idx: Optional[int]
        if page_range is not None:
            start_idx, end_idx = page_range
        else:
            # Fallback: constrained window from first BS title mention.
            start_idx = None
            for i, text in enumerate(page_texts):
                t = re.sub(r"\s+", "", text or "")
                if re.search(r"(?:合并|并)(?:及公司)?资产负债表", t):
                    start_idx = i
                    break
            if start_idx is None:
                start_idx = 0
            end_idx = min(len(page_texts) - 1, start_idx + 8)

        lines: List[str] = []
        for text in page_texts[start_idx : end_idx + 1]:
            t_norm = re.sub(r"\s+", "", text or "")
            # Reject fair-value hierarchy note pages.
            if any(x in t_norm for x in ("第一层次", "第二层次", "第三层次", "以公允价值计量")):
                continue
            for line in text.split("\n"):
                line_s = line.strip()
                if not line_s:
                    continue
                lines.append(line_s)
                l_norm = re.sub(r"\s+", "", line_s)
                if any(x in l_norm for x in ("第一层次", "第二层次", "第三层次", "公允价值")):
                    continue
                for label in ordered_labels:
                    code = target_codes[label]
                    if label in line_s:
                        val = _parse_first_amount_from_line(line_s)
                        if val is not None:
                            out[code] = val

        # Apply strict TOTAL_ASSETS fallback if generic parsing still missed it.
        if "TOTASSET" not in out:
            v_total_assets = _extract_bs_total_assets_from_lines(lines)
            if v_total_assets is not None:
                out["TOTASSET"] = v_total_assets
        # Apply strict TOTAL_LIABILITIES fallback/override using BS anchors.
        v_total_liab = _extract_bs_total_liabilities_from_lines(lines)
        if v_total_liab is not None:
            out["TOTLIAB"] = v_total_liab
    return out


def build_is_row(
    symbol: str,
    filing: CnAnnouncement,
    is_map: Dict[str, float],
    reported_unit: str,
    is_source_labels: Optional[Dict[str, str]] = None,
    is_status_map: Optional[Dict[str, str]] = None,
    is_raw_text_map: Optional[Dict[str, str]] = None,
) -> dict:
    # Priority mapping for CN IS:
    # 1) 营业总收入 / 营业总成本
    # 2) fallback to 营业收入 / 营业成本 only when totals are missing
    total_revenue = pick_first(is_map, ["BIZTOTINCO", "BIZINCO"])
    cost_of_revenue = _expense_abs(pick_first(is_map, ["BIZTOTCOST", "BIZCOST"]))
    gross_profit = None
    if total_revenue is not None and cost_of_revenue is not None:
        gross_profit = total_revenue - cost_of_revenue

    # Keep original sign for line items (e.g., finance expense can be negative).
    sales = pick_first(is_map, ["SALESEXPE"])
    mana = pick_first(is_map, ["MANAEXPE"])
    deve = pick_first(is_map, ["DEVEEXPE"])
    fin = pick_first(is_map, ["FINEXPE", "FININCO"])
    biz_tax = pick_first(is_map, ["BIZTAX"])
    operating_expenses = sum_values(sales, mana, deve, fin, biz_tax)

    non_op_income = pick_first(is_map, ["NONOREVE"])
    non_op_expense = _expense_abs(pick_first(is_map, ["NONOEXPE"]))
    non_op_net = None
    if non_op_income is not None or non_op_expense is not None:
        non_op_net = (non_op_income or 0.0) - (non_op_expense or 0.0)

    return {
        "country": "CN",
        "symbol": symbol,
        "company_id": filing.org_id,
        "form_type": filing.title,
        "fiscal_year_end_date": filing.fiscal_year_end_date,
        "filing_date": filing.filing_date,
        "fiscal_year": filing.fiscal_year,
        "fiscal_period": "FY",
        "source_filing_id": filing.announcement_id,
        "is_amendment": 1 if "更正" in filing.title else 0,
        "accounting_standard": "CN_GAAP",
        "taxonomy_version": None,
        "statement_scope": "CONSOLIDATED",
        "currency": "CNY",
        "reported_unit": reported_unit,
        "unit_scale": 1,
        "total_revenue": total_revenue,
        "cost_of_revenue": cost_of_revenue,
        "gross_profit": gross_profit,
        "taxes_and_surcharges": biz_tax,
        "selling_expense": sales,
        "admin_expense": mana,
        "rnd_expense": deve,
        "finance_expense": fin,
        "operating_expenses": operating_expenses,
        "operating_income": pick_first(is_map, ["PERPROFIT"]),
        "non_operating_income_expense_net": non_op_net,
        "other_income": pick_first(is_map, ["OTHERINCO"]),
        "income_before_income_taxes": pick_first(is_map, ["TOTPROFIT"]),
        "provision_for_income_taxes": _expense_abs(pick_first(is_map, ["INCOTAXEXPE"])),
        "net_income": pick_first(is_map, ["NETPROFIT", "PARENETP"]),
        "net_income_per_share_basic": pick_first(is_map, ["BASICEPS"]),
        "net_income_per_share_diluted": pick_first(is_map, ["DILUTEDEPS"]),
        "net_income_per_share_basic_status": (
            "OK"
            if pick_first(is_map, ["BASICEPS"]) is not None
            else (
                (is_source_labels or {}).get("BASICEPS_status")
                or (is_status_map or {}).get("BASICEPS")
                or "MISSING"
            )
        ),
        "net_income_per_share_diluted_status": (
            "OK"
            if pick_first(is_map, ["DILUTEDEPS"]) is not None
            else (
                (is_source_labels or {}).get("DILUTEDEPS_status")
                or (is_status_map or {}).get("DILUTEDEPS")
                or "MISSING"
            )
        ),
        "net_income_per_share_basic_raw_text": (is_raw_text_map or {}).get("BASICEPS"),
        "net_income_per_share_diluted_raw_text": (is_raw_text_map or {}).get("DILUTEDEPS"),
        "finance_result_source_label": (
            (is_source_labels or {}).get("finance_result_source_label")
        ),
        "shares_outstanding_basic": None,
        "shares_outstanding_diluted": None,
        "anomaly_flag": 0,
        "anomaly_type": None,
        "anomaly_score": None,
        "anomaly_detail": None,
    }


def build_bs_row(symbol: str, filing: CnAnnouncement, bs_map: Dict[str, float], reported_unit: str) -> dict:
    cash = pick_first(bs_map, ["CURFDS"])
    short_inv = pick_first(bs_map, ["TRADFINASSET"])
    cash_plus_short = None
    if cash is not None and short_inv is not None:
        cash_plus_short = cash + short_inv
    total_assets = pick_first(bs_map, ["TOTASSET"])
    total_liabilities = pick_first(bs_map, ["TOTLIAB"])
    if total_liabilities is None:
        total_liabilities = sum_values(
            pick_first(bs_map, ["TOTALCURRLIAB"]),
            pick_first(bs_map, ["TOTALNONCLIAB"]),
        )
    total_l_and_e = pick_first(bs_map, ["TOTLIABSHAREQUI"])
    if total_assets is not None and total_l_and_e is not None:
        if abs(total_l_and_e - total_assets) > max(1.0, abs(total_assets) * 1e-8):
            total_l_and_e = total_assets
    elif total_l_and_e is None:
        total_l_and_e = total_assets
    # Last-resort consistency guard for presentation when TOTASSET is missing.
    if total_assets is None and total_l_and_e is not None:
        total_assets = total_l_and_e

    return {
        "country": "CN",
        "symbol": symbol,
        "company_id": filing.org_id,
        "form_type": filing.title,
        "fiscal_year_end_date": filing.fiscal_year_end_date,
        "filing_date": filing.filing_date,
        "fiscal_year": filing.fiscal_year,
        "fiscal_period": "FY",
        "source_filing_id": filing.announcement_id,
        "is_amendment": 1 if "更正" in filing.title else 0,
        "accounting_standard": "CN_GAAP",
        "taxonomy_version": None,
        "statement_scope": "CONSOLIDATED",
        "currency": "CNY",
        "reported_unit": reported_unit,
        "unit_scale": 1,
        "total_assets": total_assets,
        "cash_and_cash_equivalents": cash,
        "short_term_investments": short_inv,
        "total_cash_and_short_term_investments": cash_plus_short,
        "inventories": pick_first(bs_map, ["INVE"]),
        "accounts_receivable": pick_first(bs_map, ["ACCORECE", "NOTESACCORECE"]),
        "other_current_assets": pick_first(bs_map, ["OTHERCURRASSE"]),
        "total_current_assets": pick_first(bs_map, ["TOTCURRASSET"]),
        "property_plant_and_equipment_net": pick_first(bs_map, ["FIXEDASSENETW", "FIXEDASSEIMMO"]),
        "goodwill": pick_first(bs_map, ["GOODWILL"]),
        "intangible_assets": pick_first(bs_map, ["INTAASSET"]),
        "other_noncurrent_assets": pick_first(bs_map, ["OTHERNONCASSE", "OTHENONCASSE"]),
        "accounts_payable": pick_first(bs_map, ["ACCOPAYA", "NOTESACCOPAYA"]),
        "short_term_debt": pick_first(bs_map, ["SHORTTERMBORR", "SHORTBORR"]),
        "other_current_liabilities": pick_first(bs_map, ["OTHERCURRELIABI", "OTHERCURLIAB"]),
        "total_current_liabilities": pick_first(bs_map, ["TOTALCURRLIAB"]),
        "long_term_debt": pick_first(bs_map, ["LONGBORR"]),
        "other_noncurrent_liabilities": pick_first(bs_map, ["OTHERNONCLIABI", "OTHENONCLIAB"]),
        "total_liabilities": total_liabilities,
        "total_shareholders_equity": pick_first(bs_map, ["PARESHARRIGH", "RIGHAGGR"]),
        "accumulated_other_comprehensive_income_or_loss": pick_first(bs_map, ["OTHERCOMPINCO"]),
        "common_stock": pick_first(bs_map, ["PAIDINCAPI"]),
        "additional_paid_in_capital": pick_first(bs_map, ["CAPISURP"]),
        "retained_earnings": pick_first(bs_map, ["UNDIPROF"]),
        "noncontrolling_interests": pick_first(bs_map, ["MINYSHARRIGH"]),
        "total_liabilities_and_shareholders_equity": total_l_and_e,
        "anomaly_flag": 0,
        "anomaly_type": None,
        "anomaly_score": None,
        "anomaly_detail": None,
    }


def build_cf_row(
    symbol: str,
    filing: CnAnnouncement,
    cf_map: Dict[str, float],
    reported_unit: str,
    is_map: Optional[Dict[str, float]] = None,
) -> dict:
    net_income = pick_first(cf_map, ["NETPROFIT", "PARENETP"])
    if net_income is None and is_map is not None:
        # Fallback to IS net income for display/ratio usability when CF statement omits it.
        net_income = pick_first(is_map, ["NETPROFIT", "PARENETP"])

    return {
        "country": "CN",
        "symbol": symbol,
        "company_id": filing.org_id,
        "form_type": filing.title,
        "fiscal_year_end_date": filing.fiscal_year_end_date,
        "filing_date": filing.filing_date,
        "fiscal_year": filing.fiscal_year,
        "fiscal_period": "FY",
        "source_filing_id": filing.announcement_id,
        "is_amendment": 1 if "更正" in filing.title else 0,
        "accounting_standard": "CN_GAAP",
        "taxonomy_version": None,
        "statement_scope": "CONSOLIDATED",
        "currency": "CNY",
        "reported_unit": reported_unit,
        "unit_scale": 1,
        "net_income": net_income,
        "net_cash_operating": pick_first(cf_map, ["MANANETR"]),
        "net_cash_investing": pick_first(cf_map, ["INVNETCASHFLOW"]),
        "net_cash_financing": pick_first(cf_map, ["FINNETCFLOW"]),
        "effect_of_exchange_rates_on_cash": pick_first(cf_map, ["CHGEXCHGCHGS"]),
        "net_change_in_cash": pick_first(cf_map, ["CASHNETR"]),
        "cash_beginning_of_period": pick_first(cf_map, ["INICASHBALA"]),
        "cash_end_of_period": pick_first(cf_map, ["FINALCASHBALA"]),
        "anomaly_flag": 0,
        "anomaly_type": None,
        "anomaly_score": None,
        "anomaly_detail": None,
    }


def export_three_statements_excel(
    symbol: str,
    company_name_hint: Optional[str],
    pdf_path: Optional[Path],
    schema_file: Path,
    as_of_date: str,
    lookback_years: int,
    raw_out: Optional[Path],
    out_path: Path,
    llm_verify: bool = False,
    llm_model: str = LLM_DEFAULT_MODEL,
    llm_base_url: str = LLM_DEFAULT_BASE_URL,
    llm_api_key: Optional[str] = None,
    llm_max_fields: int = 0,
    llm_audit_csv: Optional[Path] = None,
    llm_allow_repair_all_status: bool = False,
) -> Dict[str, str]:
    session = requests.Session()
    session.headers.update(cninfo_headers())

    symbol = normalize_symbol(symbol)
    if pdf_path is not None:
        # Use local PDF mode.
        print(f"[INFO] Local PDF mode: {pdf_path}")
        filing = CnAnnouncement(
            symbol=symbol,
            org_id="LOCAL",
            company_name=company_name_hint or symbol,
            title="年度报告 (LOCAL_PDF)",
            adjunct_url="",
            announcement_id=f"LOCAL_{pdf_path.stem}",
            announcement_time_ms=int(datetime.now().timestamp() * 1000),
            fiscal_year=int(as_of_date[:4]) - 1,
        )
        annual_pdf = pdf_path
    else:
        stock_meta = load_stock_meta(session, symbol)
        org_id = str(stock_meta.get("orgId", "")).strip()
        company_name = str(stock_meta.get("zwjc", "")).strip()
        if not org_id:
            raise RuntimeError(f"CNINFO orgId missing for symbol={symbol}")

        end_dt = datetime.strptime(as_of_date, "%Y-%m-%d").date()
        start_year = max(2000, end_dt.year - lookback_years)
        anns = query_annual_announcements(
            session=session,
            symbol=symbol,
            org_id=org_id,
            lookback_start=f"{start_year:04d}-01-01",
            lookback_end=end_dt.isoformat(),
        )
        filing = pick_latest_annual_report(
            symbol=symbol,
            org_id=org_id,
            company_name=company_name,
            announcements=anns,
            company_name_hint=company_name_hint,
        )
        print(
            f"[INFO] Filing selected: symbol={filing.symbol} company={filing.company_name} "
            f"title={filing.title} filing_date={filing.filing_date}"
        )
        annual_pdf = Path(f"CN/raw_pdfs/{symbol}_{filing.fiscal_year}_annual.pdf")
        download_pdf(session, filing.full_pdf_url, annual_pdf)
        print(f"[INFO] Annual report PDF downloaded: {annual_pdf}")

    raw_extract_path = raw_out or Path(f"CN/{symbol}_raw_extract.xlsx")
    run_pipeline(
        pdf_path=str(annual_pdf),
        schema_file=str(schema_file),
        out_path=str(raw_extract_path),
    )
    print(f"[INFO] Raw extraction workbook: {raw_extract_path}")
    reported_unit = extract_reported_unit_from_pdf(annual_pdf)
    print(f"[INFO] Reported unit detected: {reported_unit}")

    is_map = load_value_map(raw_extract_path, "CN_FIN_IS_GEN")
    bs_map = load_value_map(raw_extract_path, "CN_FIN_BS_GEN")
    cf_map = load_value_map(raw_extract_path, "CN_FIN_CF_GEN")
    is_status_map = load_status_map(raw_extract_path, "CN_FIN_IS_GEN")
    is_raw_text_map = load_raw_text_map(raw_extract_path, "CN_FIN_IS_GEN")
    is_map, is_source_labels = enrich_is_map_from_pdf_text(annual_pdf, is_map)
    bs_map = enrich_bs_map_from_pdf_text(annual_pdf, bs_map)
    cf_map = enrich_cf_map_from_pdf_text(annual_pdf, cf_map)
    print(f"[INFO] Parsed mapped values: IS={len(is_map)} BS={len(bs_map)} CF={len(cf_map)}")

    is_row = build_is_row(
        symbol,
        filing,
        is_map,
        reported_unit=reported_unit,
        is_source_labels=is_source_labels,
        is_status_map=is_status_map,
        is_raw_text_map=is_raw_text_map,
    )
    bs_row = build_bs_row(symbol, filing, bs_map, reported_unit=reported_unit)
    cf_row = build_cf_row(symbol, filing, cf_map, reported_unit=reported_unit, is_map=is_map)

    out_path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        pd.DataFrame([is_row]).to_excel(writer, index=False, sheet_name="CN_FIN_IS")
        pd.DataFrame([bs_row]).to_excel(writer, index=False, sheet_name="CN_FIN_BS")
        pd.DataFrame([cf_row]).to_excel(writer, index=False, sheet_name="CN_FIN_CF")
        pd.read_excel(raw_extract_path, sheet_name="CN_FIN_IS_GEN").to_excel(writer, index=False, sheet_name="RAW_CN_FIN_IS_GEN")
        pd.read_excel(raw_extract_path, sheet_name="CN_FIN_BS_GEN").to_excel(writer, index=False, sheet_name="RAW_CN_FIN_BS_GEN")
        pd.read_excel(raw_extract_path, sheet_name="CN_FIN_CF_GEN").to_excel(writer, index=False, sheet_name="RAW_CN_FIN_CF_GEN")

    llm_audit_csv_path: Optional[Path] = None
    if llm_verify:
        if not llm_api_key:
            raise RuntimeError("LLM verify requires OPENAI_API_KEY or --llm-api-key.")
        print(f"[INFO] Running LLM verify/repair: model={llm_model} temperature=0")
        audit_df = audit_cn_workbook(
            out_path,
            api_key=llm_api_key,
            model=llm_model,
            base_url=llm_base_url,
            max_fields=llm_max_fields,
            repair_allowed_only_for_missing_or_parse_error=(not llm_allow_repair_all_status),
        )
        append_audit_sheet(out_path, audit_df, sheet_name="LLM_AUDIT_CN")
        llm_audit_csv_path = llm_audit_csv or out_path.with_name(f"{out_path.stem}_llm_audit.csv")
        llm_audit_csv_path.parent.mkdir(parents=True, exist_ok=True)
        audit_df.to_csv(llm_audit_csv_path, index=False, encoding="utf-8-sig")
        print(f"[OK] LLM audit sheet updated: {out_path}#LLM_AUDIT_CN")
        print(f"[OK] LLM audit CSV: {llm_audit_csv_path}")

    return {
        "out_path": str(out_path),
        "raw_extract_path": str(raw_extract_path),
        "llm_audit_csv": str(llm_audit_csv_path) if llm_audit_csv_path else "",
        "pdf_path": str(annual_pdf),
        "source_filing_id": filing.announcement_id,
        "filing_title": filing.title,
        "filing_date": filing.filing_date,
        "fiscal_year_end_date": filing.fiscal_year_end_date,
        "symbol": symbol,
        "company_id": filing.org_id,
        "company_name": filing.company_name,
    }


def main() -> None:
    args = parse_args()
    try:
        result = export_three_statements_excel(
            symbol=args.symbol,
            company_name_hint=args.company_name,
            pdf_path=Path(args.pdf) if args.pdf else None,
            schema_file=Path(args.schema_file),
            as_of_date=args.as_of_date,
            lookback_years=args.lookback_years,
            raw_out=Path(args.raw_out) if args.raw_out else None,
            out_path=Path(args.out),
            llm_verify=args.llm_verify,
            llm_model=args.llm_model,
            llm_base_url=args.llm_base_url,
            llm_api_key=args.llm_api_key,
            llm_max_fields=args.llm_max_fields,
            llm_audit_csv=Path(args.llm_audit_csv) if args.llm_audit_csv else None,
            llm_allow_repair_all_status=args.llm_allow_repair_all_status,
        )
    except Exception as e:
        raise SystemExit(f"CN extraction failed: {e}")

    print("[OK] Excel exported:", result["out_path"])
    print(
        "[OK] Filing metadata:",
        json.dumps(
            {
                "symbol": result["symbol"],
                "company_name": result["company_name"],
                "company_id": result["company_id"],
                "source_filing_id": result["source_filing_id"],
                "filing_title": result["filing_title"],
                "filing_date": result["filing_date"],
                "fiscal_year_end_date": result["fiscal_year_end_date"],
                "pdf_path": result["pdf_path"],
                "raw_extract_path": result["raw_extract_path"],
                "llm_audit_csv": result.get("llm_audit_csv") or None,
            },
            ensure_ascii=False,
        ),
    )


if __name__ == "__main__":
    main()
