#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Document-level CN LLM-only baseline.

Unlike candidate-level baseline, this script reads statement text directly from
raw annual-report PDF and asks LLM to extract target fields.
"""

from __future__ import annotations

import argparse
import json
import math
import os
import re
import sys
import time
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd
import pdfplumber
import requests

ROOT_DIR = Path(__file__).resolve().parent.parent
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

from extract_statements import run as locate_statement_pages  # noqa: E402
from llm_cn_verifier import (  # noqa: E402
    DEFAULT_BASE_URL,
    DEFAULT_MODEL,
    FIELD_CONFIG,
    FIXED_TEMPERATURE,
    REQUEST_TIMEOUT,
    find_cn_workbooks,
)


FIELD_ZH_LABEL = {
    "total_revenue": "营业总收入/营业收入",
    "operating_income": "营业利润/经营利润",
    "income_before_income_taxes": "利润总额/税前利润",
    "net_income": "净利润",
    "net_income_per_share_basic": "基本每股收益",
    "net_income_per_share_diluted": "稀释每股收益",
    "total_assets": "资产总计",
    "cash_and_cash_equivalents": "货币资金/现金及现金等价物",
    "accounts_receivable": "应收账款/应收票据及应收账款",
    "inventories": "存货",
    "total_liabilities": "负债合计",
    "total_shareholders_equity": "归属于母公司股东权益合计/所有者权益合计",
    "total_liabilities_and_shareholders_equity": "负债和所有者权益总计",
    "net_cash_operating": "经营活动产生的现金流量净额",
    "net_cash_investing": "投资活动产生的现金流量净额",
    "net_cash_financing": "筹资活动产生的现金流量净额",
    "net_change_in_cash": "现金及现金等价物净增加额",
    "cash_end_of_period": "期末现金及现金等价物余额",
}

STATEMENT_TITLE = {
    "IS": "利润表",
    "BS": "资产负债表",
    "CF": "现金流量表",
}

STMT_NAME_MAP = {
    "利润表": "IS",
    "资产负债表": "BS",
    "现金流量表": "CF",
}

TITLE_PATTERNS = {
    "BS": [re.compile(r"合并(?:及(?:公司|母公司))?资产负债表(?:(?:（续）|\(续\)|[-—－]续))?")],
    "IS": [re.compile(r"合并(?:及(?:公司|母公司))?利润表(?:(?:（续）|\(续\)|[-—－]续))?")],
    "CF": [re.compile(r"合并(?:及(?:公司|母公司))?现金流量表(?:(?:（续）|\(续\)|[-—－]续))?")],
}

STMT_KEYWORDS = {
    "IS": ["营业收入", "营业利润", "利润总额", "净利润", "基本每股收益", "稀释每股收益"],
    "BS": ["资产总计", "货币资金", "应收账款", "存货", "负债合计", "所有者权益合计", "股东权益合计"],
    "CF": ["经营活动产生的现金流量净额", "投资活动产生的现金流量净额", "筹资活动产生的现金流量净额", "现金及现金等价物净增加额"],
}

SYSTEM_PROMPT = """You are a strict CN financial statement extractor.

Task:
- Extract only requested fields from provided statement text.
- Return EXTRACT only with concrete evidence.
- If uncertain, return NEED_REVIEW.

Hard rules:
1) Do NOT fabricate values.
2) Cite page-level quote evidence.
3) Keep numeric values as plain numbers without unit words.
"""


def _to_primitive(v: Any) -> Any:
    if v is None:
        return None
    if isinstance(v, pd.Timestamp):
        return v.to_pydatetime().date().isoformat()
    if pd.isna(v):
        return None
    if hasattr(v, "item"):
        try:
            return v.item()
        except Exception:
            pass
    return v


def _to_float(v: Any) -> Optional[float]:
    vv = _to_primitive(v)
    if vv is None:
        return None
    try:
        x = float(vv)
    except Exception:
        return None
    if math.isnan(x):
        return None
    return x


def _normalize_text(s: str) -> str:
    s = s or ""
    s = re.sub(r"[ \t]+", " ", s)
    s = re.sub(r"\n{3,}", "\n\n", s)
    return s.strip()


def _parse_json_content(raw: str) -> Dict[str, Any]:
    s = (raw or "").strip()
    if s.startswith("```"):
        s = s.strip("`")
        if s.startswith("json"):
            s = s[4:].strip()
    return json.loads(s)


def _read_row0(path: Path, sheet_name: str) -> Dict[str, Any]:
    df = pd.read_excel(path, sheet_name=sheet_name)
    if df.empty:
        return {}
    return {str(k): _to_primitive(v) for k, v in df.iloc[0].to_dict().items()}


def resolve_pdf_path(symbol: str, fiscal_year: int) -> Optional[Path]:
    target = ROOT_DIR / "CN" / "raw_pdfs" / f"{symbol}_{fiscal_year}_annual.pdf"
    if target.exists():
        return target
    alt = sorted((ROOT_DIR / "CN" / "raw_pdfs").glob(f"{symbol}_*_annual.pdf"))
    if alt:
        return alt[-1]
    return None


def _fallback_statement_ranges(page_texts: List[str]) -> Dict[str, Tuple[int, int]]:
    def _compact(s: str) -> str:
        return re.sub(r"\s+", "", s or "")

    def _score(stmt: str, compact_text: str) -> int:
        return sum(1 for kw in STMT_KEYWORDS.get(stmt, []) if kw in compact_text)

    compact_pages = [_compact(t) for t in page_texts]
    total_pages = len(compact_pages)

    def _pick_start(stmt: str) -> Optional[int]:
        title_candidates: List[int] = []
        for idx, txt in enumerate(compact_pages, start=1):
            if any(p.search(txt) for p in TITLE_PATTERNS.get(stmt, [])):
                title_candidates.append(idx)
        # Conservative fallback: without an explicit statement title, do not guess page range.
        if not title_candidates:
            return None
        candidates = title_candidates
        scored: List[Tuple[float, int]] = []
        for p in candidates:
            curr = _score(stmt, compact_pages[p - 1])
            nxt = _score(stmt, compact_pages[p]) if p < total_pages else 0
            title_bonus = 1.0 if p in title_candidates else 0.0
            scored.append((curr + 0.5 * nxt + title_bonus, p))
        scored.sort(key=lambda x: (-x[0], x[1]))
        if not scored:
            return None
        best_score, best_page = scored[0]
        if best_score < 2.0:
            return None
        return best_page

    def _extend_range(stmt: str, start_page: int) -> Tuple[int, int]:
        bare_title = {
            "IS": "利润表",
            "BS": "资产负债表",
            "CF": "现金流量表",
        }[stmt]
        full_title = {
            "IS": "合并利润表",
            "BS": "合并资产负债表",
            "CF": "合并现金流量表",
        }[stmt]
        end_page = start_page
        max_span = 8
        for p in range(start_page + 1, min(total_pages, start_page + max_span - 1) + 1):
            txt = compact_pages[p - 1]
            # Stop when entering standalone-company statement page, e.g. "资产负债表" without "合并".
            if bare_title in txt and full_title not in txt:
                break
            score = _score(stmt, txt)
            has_same_title = any(pat.search(txt) for pat in TITLE_PATTERNS.get(stmt, []))
            if score > 0 or has_same_title:
                end_page = p
                continue
            break
        return (start_page, end_page)

    ranges: Dict[str, Tuple[int, int]] = {}
    for stmt in ("IS", "BS", "CF"):
        sp = _pick_start(stmt)
        if sp is None:
            continue
        ranges[stmt] = _extend_range(stmt, sp)
    return ranges


def get_statement_text_blocks(pdf_path: Path, max_chars_per_stmt: int = 25000) -> Dict[str, Dict[str, Any]]:
    blocks: Dict[str, Dict[str, Any]] = {k: {"pages": [], "text": ""} for k in ("IS", "BS", "CF")}
    page_texts: List[str] = []
    with pdfplumber.open(str(pdf_path)) as pdf:
        page_texts = [_normalize_text(p.extract_text() or "") for p in pdf.pages]
    compact_pages = [re.sub(r"\s+", "", t or "") for t in page_texts]

    ranges: Dict[str, Tuple[int, int]] = {}
    locator_ranges: Dict[str, Tuple[int, int]] = {}
    try:
        loc = locate_statement_pages(str(pdf_path))
        for s in loc.get("statements", []):
            name = str(s.get("name") or "")
            sp = s.get("start_page")
            ep = s.get("end_page")
            key = STMT_NAME_MAP.get(name)
            if key and sp and ep:
                locator_ranges[key] = (int(sp), int(ep))
    except Exception:
        locator_ranges = {}

    def _locator_usable(r: Dict[str, Tuple[int, int]]) -> bool:
        if len(r) < 2:
            return False
        uniq = {(sp, ep) for sp, ep in r.values()}
        if len(uniq) == 1:
            return False
        total = len(compact_pages)
        for stmt, (sp, ep) in r.items():
            if sp < 1 or ep < sp or ep > total:
                return False
            title_on_start = any(p.search(compact_pages[sp - 1]) for p in TITLE_PATTERNS.get(stmt, []))
            if not title_on_start:
                return False
        return True

    if _locator_usable(locator_ranges):
        ranges = locator_ranges

    if len(ranges) < 2:
        ranges = _fallback_statement_ranges(page_texts)

    total_pages = len(page_texts)
    for stmt in ("IS", "BS", "CF"):
        if stmt not in ranges:
            continue
        sp, ep = ranges[stmt]
        sp = max(1, min(total_pages, sp))
        ep = max(sp, min(total_pages, ep))
        lines: List[str] = []
        pages: List[int] = []
        for p in range(sp, ep + 1):
            txt = page_texts[p - 1]
            if not txt:
                continue
            pages.append(p)
            lines.append(f"[PAGE {p}]\n{txt}\n")
        blob = "\n".join(lines).strip()
        if len(blob) > max_chars_per_stmt:
            blob = blob[:max_chars_per_stmt]
        blocks[stmt] = {"pages": pages, "text": blob}
    return blocks


def _response_schema(field_names: List[str]) -> Dict[str, Any]:
    return {
        "type": "json_schema",
        "json_schema": {
            "name": "cn_doc_level_llm_only",
            "strict": True,
            "schema": {
                "type": "object",
                "additionalProperties": False,
                "properties": {
                    "results": {
                        "type": "array",
                        "items": {
                            "type": "object",
                            "additionalProperties": False,
                            "properties": {
                                "field_name": {"type": "string", "enum": field_names},
                                "decision": {"type": "string", "enum": ["EXTRACT", "NEED_REVIEW"]},
                                "proposed_value": {"type": ["number", "null"]},
                                "proposed_status": {"type": "string", "enum": ["OK", "NOT_APPLICABLE", "MISSING", "PARSE_ERROR"]},
                                "evidence": {
                                    "type": "array",
                                    "items": {
                                        "type": "object",
                                        "additionalProperties": False,
                                        "properties": {
                                            "page": {"type": "integer"},
                                            "quote": {"type": "string"},
                                            "why": {"type": "string"},
                                        },
                                        "required": ["page", "quote", "why"],
                                    },
                                },
                                "reason": {"type": "string"},
                            },
                            "required": ["field_name", "decision", "proposed_value", "proposed_status", "evidence", "reason"],
                        },
                    }
                },
                "required": ["results"],
            },
        },
    }


def call_openai_doc_extract(
    *,
    api_key: str,
    model: str,
    base_url: str,
    timeout: int,
    symbol: str,
    statement: str,
    field_names: List[str],
    statement_text: str,
) -> Dict[str, Any]:
    endpoint = base_url.rstrip("/") + "/chat/completions"
    field_defs = [
        {"field_name": f, "label_zh": FIELD_ZH_LABEL.get(f, f)}
        for f in field_names
    ]
    payload_obj = {
        "market": "cn",
        "symbol": symbol,
        "statement": statement,
        "fields": field_defs,
        "statement_text": statement_text,
    }
    payload = {
        "model": model,
        "temperature": FIXED_TEMPERATURE,
        "messages": [
            {"role": "system", "content": SYSTEM_PROMPT},
            {
                "role": "user",
                "content": "Extract requested fields from statement text and return strict JSON.\n\n"
                + json.dumps(payload_obj, ensure_ascii=False),
            },
        ],
        "response_format": _response_schema(field_names),
    }
    headers = {"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"}
    r = requests.post(endpoint, headers=headers, json=payload, timeout=timeout)
    r.raise_for_status()
    body = r.json()
    content = body["choices"][0]["message"]["content"]
    obj = _parse_json_content(content)
    obj["_raw_response"] = body
    return obj


def apply_doc_guardrails(item: Dict[str, Any]) -> Dict[str, Any]:
    decision = str(item.get("decision") or "").upper()
    proposed_status = str(item.get("proposed_status") or "").upper()
    proposed_value = _to_float(item.get("proposed_value"))
    evidence = item.get("evidence") or []
    reason = str(item.get("reason") or "")

    guard_fail = ""
    if decision == "EXTRACT":
        if not evidence:
            guard_fail = "missing_evidence"
        elif proposed_status == "OK":
            if proposed_value is None:
                guard_fail = "missing_proposed_value"
        elif proposed_status == "NOT_APPLICABLE":
            has_na_hint = any("不适用" in str((e or {}).get("quote") or "") for e in evidence)
            if not has_na_hint:
                guard_fail = "unsupported_not_applicable"
        else:
            guard_fail = "invalid_extract_status"

    if guard_fail:
        decision = "NEED_REVIEW"

    final_value = proposed_value if (decision == "EXTRACT" and proposed_status == "OK") else None
    return {
        "decision": decision,
        "proposed_status": proposed_status,
        "proposed_value": proposed_value,
        "evidence": evidence,
        "reason": reason,
        "guard_fail": guard_fail or None,
        "final_value": final_value,
        "final_source": "llm_doc_only" if decision == "EXTRACT" else "llm_doc_only_review",
        "review_required_recommended": int(decision != "EXTRACT"),
    }


def run_doc_llm_only_on_workbook(
    workbook: Path,
    *,
    api_key: str,
    model: str,
    base_url: str,
    timeout: int,
    max_chars_per_stmt: int,
    sleep_seconds: float,
    max_fields: int,
) -> pd.DataFrame:
    is_row = _read_row0(workbook, "CN_FIN_IS")
    bs_row = _read_row0(workbook, "CN_FIN_BS")
    cf_row = _read_row0(workbook, "CN_FIN_CF")
    symbol = str(is_row.get("symbol") or bs_row.get("symbol") or cf_row.get("symbol") or "")
    company_id = _to_primitive(is_row.get("company_id") or bs_row.get("company_id") or cf_row.get("company_id"))
    fye = str(_to_primitive(is_row.get("fiscal_year_end_date") or bs_row.get("fiscal_year_end_date") or cf_row.get("fiscal_year_end_date")) or "")
    fy = None
    if fye and re.match(r"^\d{4}-\d{2}-\d{2}$", fye):
        fy = int(fye[:4])
    else:
        fy = int(_to_primitive(is_row.get("fiscal_year") or bs_row.get("fiscal_year") or cf_row.get("fiscal_year") or 0) or 0)

    pdf_path = resolve_pdf_path(symbol, fy) if symbol and fy else None
    stmt_texts = {k: {"pages": [], "text": ""} for k in ("IS", "BS", "CF")}
    if pdf_path is not None and pdf_path.exists():
        stmt_texts = get_statement_text_blocks(pdf_path, max_chars_per_stmt=max_chars_per_stmt)

    rows: List[Dict[str, Any]] = []
    for stmt, field_pairs in FIELD_CONFIG.items():
        fields = [f for f, _ in field_pairs]
        if max_fields > 0:
            fields = fields[:max_fields]
        text_blob = stmt_texts.get(stmt, {}).get("text", "")
        llm_error = None
        result_by_field: Dict[str, Dict[str, Any]] = {}

        if not pdf_path or not pdf_path.exists():
            llm_error = "pdf_not_found"
        elif not text_blob:
            llm_error = "statement_text_missing"
        else:
            try:
                obj = call_openai_doc_extract(
                    api_key=api_key,
                    model=model,
                    base_url=base_url,
                    timeout=timeout,
                    symbol=symbol,
                    statement=STATEMENT_TITLE.get(stmt, stmt),
                    field_names=fields,
                    statement_text=text_blob,
                )
                for it in obj.get("results", []):
                    fn = str(it.get("field_name") or "")
                    if fn in fields:
                        result_by_field[fn] = apply_doc_guardrails(it)
            except Exception as e:
                llm_error = str(e)

        for fn in fields:
            g = result_by_field.get(fn)
            if g is None:
                g = {
                    "decision": "NEED_REVIEW",
                    "proposed_status": "MISSING",
                    "proposed_value": None,
                    "evidence": [],
                    "reason": "no_result_for_field",
                    "guard_fail": llm_error or "no_result_for_field",
                    "final_value": None,
                    "final_source": "llm_doc_only_review",
                    "review_required_recommended": 1,
                }
            rows.append(
                {
                    "workbook": str(workbook),
                    "market": "cn",
                    "symbol": symbol,
                    "company_id": company_id,
                    "fiscal_year_end_date": fye,
                    "pdf_path": str(pdf_path) if pdf_path else "",
                    "statement": stmt,
                    "field_name": fn,
                    "candidate_codes_json": "[]",
                    "candidates_json": "[]",
                    "llm_only_decision": g["decision"],
                    "llm_only_value": g["proposed_value"],
                    "llm_only_status": g["proposed_status"],
                    "evidence_json": json.dumps(g["evidence"], ensure_ascii=False),
                    "evidence_item_codes_json": "[]",
                    "llm_only_reason": g["reason"],
                    "guard_fail": g["guard_fail"],
                    "final_value": g["final_value"],
                    "final_source": g["final_source"],
                    "review_required_recommended": g["review_required_recommended"],
                    "experiment_mode": "llm_only_doc",
                    "model": model,
                    "temperature": FIXED_TEMPERATURE,
                    "llm_error": llm_error,
                }
            )
        if sleep_seconds > 0:
            time.sleep(sleep_seconds)
    return pd.DataFrame(rows)


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(description="Run document-level CN LLM-only baseline.")
    p.add_argument("--workbook", help="Single CN workbook path")
    p.add_argument("--input-dir", default="eval/outputs/cn", help="Directory containing cn_*_3statements.xlsx")
    p.add_argument("--pattern", default="cn_*_3statements.xlsx", help="Glob pattern under input-dir")
    p.add_argument("--out-csv", default="eval/outputs/llm_only_cn_doc.csv", help="Output CSV path")
    p.add_argument("--model", default=DEFAULT_MODEL, help="OpenAI model")
    p.add_argument("--base-url", default=DEFAULT_BASE_URL, help="OpenAI API base URL")
    p.add_argument("--api-key", default=os.getenv("OPENAI_API_KEY"), help="OpenAI API key (or env OPENAI_API_KEY)")
    p.add_argument("--timeout", type=int, default=REQUEST_TIMEOUT, help="HTTP timeout seconds")
    p.add_argument("--sleep-seconds", type=float, default=0.0, help="Sleep between statement calls")
    p.add_argument("--max-workbooks", type=int, default=0, help="Only process first N workbooks (0 = all)")
    p.add_argument("--max-fields", type=int, default=0, help="Only process first N fields per statement (0 = all)")
    p.add_argument("--max-chars-per-stmt", type=int, default=25000, help="Max chars sent per statement text")
    return p.parse_args()


def main() -> None:
    args = parse_args()
    if not args.api_key:
        raise SystemExit("Missing OpenAI API key. Set OPENAI_API_KEY or pass --api-key.")

    if args.workbook:
        workbooks = [Path(args.workbook)]
    else:
        workbooks = find_cn_workbooks(Path(args.input_dir), pattern=args.pattern)

    if args.max_workbooks > 0:
        workbooks = workbooks[: args.max_workbooks]

    if not workbooks:
        raise SystemExit("No CN workbook found.")

    frames = []
    for wb in workbooks:
        print(f"[INFO] Doc-LLM-only workbook: {wb}")
        df = run_doc_llm_only_on_workbook(
            wb,
            api_key=args.api_key,
            model=args.model,
            base_url=args.base_url,
            timeout=args.timeout,
            max_chars_per_stmt=args.max_chars_per_stmt,
            sleep_seconds=args.sleep_seconds,
            max_fields=args.max_fields,
        )
        frames.append(df)

    out_df = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()
    out_path = Path(args.out_csv)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_df.to_csv(out_path, index=False, encoding="utf-8-sig")
    print(f"[OK] Wrote document-level LLM-only CSV: {out_path}")
    print(f"[INFO] Rows: {len(out_df)}")


if __name__ == "__main__":
    main()
