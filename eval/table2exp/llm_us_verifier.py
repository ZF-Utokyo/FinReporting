#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
US deterministic-first LLM verifier/repair layer.

This script mirrors the CN verifier interface so it can be reused in table2 backbone experiments.
"""

from __future__ import annotations

import argparse
import ast
import json
import time
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Tuple

import pandas as pd

ROOT_DIR = Path(__file__).resolve().parents[2]
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

import llm_cn_verifier as base


# Reuse provider callers/guardrails from base module with US-specific prompt/market payload.
base.MARKET = "us"
base.SYSTEM_PROMPT = """You are a strict financial extraction verifier for US 10-K structured outputs.

Task:
- Review one field case at a time.
- Decide KEEP / REPAIR / NEED_REVIEW.
- Use only provided candidate evidence.

Hard rules:
1) If evidence is weak or missing, choose NEED_REVIEW.
2) Do NOT fabricate numbers.
3) If proposing REPAIR, cite concrete candidate item codes in evidence.
4) If evidence shows value is not separately disclosed for this canonical field, keep status MISSING and use NEED_REVIEW.
5) If candidates contain conflicting numeric values, DO NOT output KEEP. Choose REPAIR (if strongly supported) or NEED_REVIEW.
"""
base.OUTPUT_CONTRACT = """Output contract (must follow exactly):
1) KEEP:
   - proposed_value must equal rule_value.
   - evidence_item_codes must contain at least 1 valid candidate item_code.
   - KEEP is forbidden when candidate numeric values conflict.
2) REPAIR:
   - proposed_value must be a concrete number from provided candidates.
   - evidence_item_codes must include the supporting candidate item_code(s).
3) NEED_REVIEW:
   - proposed_value must be null.
4) Do not use item codes outside provided candidates.
5) Return JSON only, no markdown fences.
"""
base.MAX_CONTRACT_ATTEMPTS = 3

DEFAULT_PROVIDER = base.DEFAULT_PROVIDER
DEFAULT_MODEL = "gpt-5.2"
REQUEST_TIMEOUT = base.REQUEST_TIMEOUT
FIXED_TEMPERATURE = base.FIXED_TEMPERATURE
LLM_CALL_MAX_RETRIES = 4

SHEETS = {
    "IS": "US_FIN_IS",
    "BS": "US_FIN_BS",
    "CF": "US_FIN_CF",
}
RAW_CANDIDATE_SHEET = "RAW_US_FIELD_CANDIDATES"
EXPANDED_CANDIDATE_TOPK = 6

# Same 18-field protocol used in manual checking.
FIELD_CONFIG: Dict[str, List[str]] = {
    "IS": [
        "total_revenue",
        "operating_income",
        "income_before_income_taxes",
        "net_income",
        "net_income_per_share_basic",
    ],
    "BS": [
        "total_assets",
        "cash_and_cash_equivalents",
        "accounts_receivable",
        "inventories",
        "total_liabilities",
        "total_shareholders_equity",
        "total_liabilities_and_shareholders_equity",
    ],
    "CF": [
        "net_income",
        "net_cash_operating",
        "net_cash_investing",
        "net_cash_financing",
        "net_change_in_cash",
        "cash_end_of_period",
    ],
}


@dataclass
class FieldCase:
    workbook: str
    symbol: str
    company_id: Optional[str]
    fiscal_year_end_date: Optional[str]
    statement: str
    field_name: str
    rule_value: Optional[float]
    rule_status: str
    candidate_codes: List[str]
    candidates: List[Dict[str, Any]]
    expanded_reason: str = ""
    num_candidates_to_llm: int = 0


def _to_primitive(v: Any) -> Any:
    return base._to_primitive(v)


def _to_float(v: Any) -> Optional[float]:
    return base._to_float(v)


def _float_close(a: Optional[float], b: Optional[float]) -> bool:
    if a is None or b is None:
        return False
    if abs(a - b) <= 1e-9:
        return True
    scale = max(abs(a), abs(b), 1.0)
    return abs(a - b) / scale <= 1e-6


def _read_row0(path: Path, sheet_name: str) -> Dict[str, Any]:
    df = pd.read_excel(path, sheet_name=sheet_name)
    if df.empty:
        return {}
    return {str(k): _to_primitive(v) for k, v in df.iloc[0].to_dict().items()}


def _read_candidate_pool(path: Path) -> Dict[Tuple[str, str], List[Dict[str, Any]]]:
    try:
        df = pd.read_excel(path, sheet_name=RAW_CANDIDATE_SHEET)
    except Exception:
        return {}
    if df.empty:
        return {}

    required = {"statement", "field_name", "candidate_id", "tag", "value"}
    if not required.issubset(set(df.columns)):
        return {}

    sort_cols = [c for c in ["statement", "field_name", "candidate_rank", "score"] if c in df.columns]
    if sort_cols:
        df = df.sort_values(sort_cols, ascending=[True, True, True, False][: len(sort_cols)])

    out: Dict[Tuple[str, str], List[Dict[str, Any]]] = {}
    for _, r in df.iterrows():
        stmt = str(_to_primitive(r.get("statement")) or "").strip()
        field = str(_to_primitive(r.get("field_name")) or "").strip()
        if not stmt or not field:
            continue
        candidate_id = str(_to_primitive(r.get("candidate_id")) or "").strip()
        tag = str(_to_primitive(r.get("tag")) or "").strip()
        unit_ref = str(_to_primitive(r.get("unit_ref")) or "").strip()
        period_type = str(_to_primitive(r.get("period_type")) or "").strip().lower()
        end_date = str(_to_primitive(r.get("end_date")) or "").strip()
        instant = str(_to_primitive(r.get("instant")) or "").strip()
        start_date = str(_to_primitive(r.get("start_date")) or "").strip()
        duration_days = _to_float(r.get("duration_days"))
        score = _to_float(r.get("score"))
        is_consolidated = int(_to_float(r.get("is_consolidated")) or 0)
        tag_priority = None
        try:
            sb = ast.literal_eval(str(_to_primitive(r.get("score_breakdown")) or ""))
            if isinstance(sb, dict):
                tag_priority = _to_float(sb.get("tag_priority"))
        except Exception:
            pass
        raw_text = (
            f"tag={tag}; context={str(_to_primitive(r.get('context_ref')) or '')}; "
            f"score={str(_to_primitive(r.get('score')) or '')}; "
            f"period={str(_to_primitive(r.get('start_date')) or '')}->{str(_to_primitive(r.get('end_date')) or _to_primitive(r.get('instant')) or '')}"
        )
        val = _to_primitive(r.get("value"))
        item_code = candidate_id or f"{stmt}.{field}.fact.{tag}"
        c = {
            "item_code": item_code,
            "item_name": tag or item_code,
            "value": val,
            "status": "OK" if val is not None else "MISSING",
            "raw_text": raw_text,
            "fixup_reason": None,
            "__unit_ref": unit_ref,
            "__period_type": period_type,
            "__start_date": start_date,
            "__end_date": end_date,
            "__instant": instant,
            "__duration_days": duration_days,
            "__score": score,
            "__is_consolidated": is_consolidated,
            "__tag_priority": tag_priority,
        }
        out.setdefault((stmt, field), []).append(c)
    return out


def _derive_rule_status(rule_value: Any) -> str:
    return "OK" if _to_primitive(rule_value) is not None else "MISSING"


def _candidate(code: str, value: Any, raw_text: str) -> Dict[str, Any]:
    v = _to_primitive(value)
    return {
        "item_code": code,
        "item_name": raw_text,
        "value": v,
        "status": "OK" if v is not None else "MISSING",
        "raw_text": raw_text,
        "fixup_reason": None,
    }


def _report_date_from_rows(rows: Dict[str, Dict[str, Any]]) -> str:
    for stmt in ["IS", "BS", "CF"]:
        v = _to_primitive(rows.get(stmt, {}).get("fiscal_year_end_date"))
        if v:
            return str(v)
    return ""


def _unit_expectations(stmt: str, key: str) -> List[str]:
    if stmt == "IS" and key == "net_income_per_share_basic":
        return ["usd/shares", "pure"]
    return ["usd"]


def _parse_int(v: Any, default: int = 0) -> int:
    try:
        return int(float(v))
    except Exception:
        return default


def _field_needs_expanded_candidates(
    stmt: str,
    key: str,
    rows: Dict[str, Dict[str, Any]],
    candidate_pool: Optional[Dict[Tuple[str, str], List[Dict[str, Any]]]] = None,
) -> Tuple[bool, str]:
    cur = rows.get(stmt, {})
    val = _to_float(cur.get(key))
    if val is None:
        return True, "missing"

    pool = list((candidate_pool or {}).get((stmt, key), []))
    if not pool:
        if key == "net_income":
            is_v = _to_float(rows.get("IS", {}).get("net_income"))
            cf_v = _to_float(rows.get("CF", {}).get("net_income"))
            if is_v is not None and cf_v is not None and not _float_close(is_v, cf_v):
                return True, "net_income_mismatch"
        return False, ""

    pool = sorted(pool, key=lambda c: (_to_float(c.get("__score")) or -10**9), reverse=True)
    top1 = pool[0]
    top2 = pool[1] if len(pool) > 1 else None

    # A) low confidence: top1-top2 gap too small, weak tag priority, or non-consolidated top1 while consolidated exists.
    gap_threshold = 15.0
    tag_priority_floor = 1.0
    s1 = _to_float(top1.get("__score"))
    s2 = _to_float(top2.get("__score")) if top2 is not None else None
    if s1 is not None and s2 is not None and (s1 - s2) < gap_threshold:
        return True, "low_confidence"

    tp1 = _to_float(top1.get("__tag_priority"))
    if tp1 is not None and tp1 <= tag_priority_floor:
        return True, "low_confidence"

    top1_cons = _parse_int(top1.get("__is_consolidated"))
    has_consolidated = any(_parse_int(c.get("__is_consolidated")) == 1 for c in pool)
    if top1_cons == 0 and has_consolidated:
        return True, "low_confidence"

    report_date = _report_date_from_rows(rows)
    period_type = str(top1.get("__period_type") or "").strip().lower()
    if report_date:
        # B) period mismatch: top1 period misaligned but pool has aligned candidate.
        if period_type == "duration":
            top_end = str(top1.get("__end_date") or "").strip()
            top_dur = _to_float(top1.get("__duration_days"))
            top_bad = (top_end != report_date) or (top_dur is None) or (top_dur < 330) or (top_dur > 400)

            def _good_duration(c: Dict[str, Any]) -> bool:
                end = str(c.get("__end_date") or "").strip()
                dur = _to_float(c.get("__duration_days"))
                return end == report_date and dur is not None and 330 <= dur <= 400

            has_good = any(_good_duration(c) for c in pool)
            if top_bad and has_good:
                return True, "period_mismatch"

        if period_type == "instant":
            top_inst = str(top1.get("__instant") or top1.get("__end_date") or "").strip()
            top_bad = top_inst != report_date
            has_good = any(
                str(c.get("__instant") or c.get("__end_date") or "").strip() == report_date
                for c in pool
            )
            if top_bad and has_good:
                return True, "period_mismatch"

    # C) unit mismatch: top1 unit unreasonable while pool has reasonable unit.
    expected_units = _unit_expectations(stmt, key)
    top_unit = str(top1.get("__unit_ref") or "").lower()
    top_ok = any(k in top_unit for k in expected_units)
    has_ok = any(any(k in str(c.get("__unit_ref") or "").lower() for k in expected_units) for c in pool)
    if (not top_ok) and has_ok:
        return True, "unit_mismatch"

    if key == "net_income":
        is_v = _to_float(rows.get("IS", {}).get("net_income"))
        cf_v = _to_float(rows.get("CF", {}).get("net_income"))
        if is_v is not None and cf_v is not None and not _float_close(is_v, cf_v):
            return True, "net_income_mismatch"

    return False, ""


def _build_candidates(
    stmt: str,
    key: str,
    rows: Dict[str, Dict[str, Any]],
    candidate_pool: Optional[Dict[Tuple[str, str], List[Dict[str, Any]]]] = None,
) -> Tuple[List[Dict[str, Any]], str]:
    out: List[Dict[str, Any]] = []
    cur = rows.get(stmt, {})
    use_pool, expanded_reason = _field_needs_expanded_candidates(
        stmt, key, rows, candidate_pool=candidate_pool
    )
    pool_hits = list((candidate_pool or {}).get((stmt, key), [])[:EXPANDED_CANDIDATE_TOPK]) if use_pool else []
    if pool_hits:
        out.extend(pool_hits)
    out.append(_candidate(f"{stmt}.{key}", cur.get(key), f"{stmt} primary field: {key}"))

    # Cross-statement consistency candidate for net_income.
    if key == "net_income":
        if stmt == "IS":
            out.append(_candidate("CF.net_income", rows.get("CF", {}).get("net_income"), "CF net_income"))
        if stmt == "CF":
            out.append(_candidate("IS.net_income", rows.get("IS", {}).get("net_income"), "IS net_income"))

    # Remove fully-empty candidates and de-duplicate item_code.
    dedup: List[Dict[str, Any]] = []
    seen = set()
    for c in out:
        code = str(c.get("item_code") or "").strip()
        if code in seen:
            continue
        seen.add(code)
        if c.get("value") is not None or c.get("status") == "MISSING":
            dedup.append(c)
    return dedup, expanded_reason


def build_cases_from_workbook(path: Path) -> List[FieldCase]:
    is_row = _read_row0(path, SHEETS["IS"])
    bs_row = _read_row0(path, SHEETS["BS"])
    cf_row = _read_row0(path, SHEETS["CF"])
    rows = {"IS": is_row, "BS": bs_row, "CF": cf_row}
    candidate_pool = _read_candidate_pool(path)

    symbol = str(is_row.get("symbol") or bs_row.get("symbol") or cf_row.get("symbol") or "")
    company_id = _to_primitive(is_row.get("company_id") or bs_row.get("company_id") or cf_row.get("company_id"))
    fye = _to_primitive(
        is_row.get("fiscal_year_end_date")
        or bs_row.get("fiscal_year_end_date")
        or cf_row.get("fiscal_year_end_date")
    )

    cases: List[FieldCase] = []
    for stmt, keys in FIELD_CONFIG.items():
        row = rows.get(stmt, {})
        for key in keys:
            rule_value = row.get(key)
            candidates, expanded_reason = _build_candidates(
                stmt, key, rows, candidate_pool=candidate_pool
            )
            candidate_codes = [str(c.get("item_code") or "") for c in candidates if str(c.get("item_code") or "")]
            cases.append(
                FieldCase(
                    workbook=str(path),
                    symbol=symbol,
                    company_id=str(company_id) if company_id is not None else None,
                    fiscal_year_end_date=str(fye) if fye is not None else None,
                    statement=stmt,
                    field_name=key,
                    rule_value=_to_float(rule_value),
                    rule_status=_derive_rule_status(rule_value),
                    candidate_codes=candidate_codes,
                    candidates=candidates,
                    expanded_reason=expanded_reason,
                    num_candidates_to_llm=len(candidates),
                )
            )
    return cases


def find_us_workbooks(input_dir: Path, pattern: str = "us_*_3statements.xlsx") -> List[Path]:
    if not input_dir.exists():
        return []
    return sorted(input_dir.glob(pattern))


def _merge_frames(frames: Iterable[pd.DataFrame]) -> pd.DataFrame:
    xs = [x for x in frames if x is not None and not x.empty]
    if not xs:
        return pd.DataFrame()
    return pd.concat(xs, ignore_index=True)


def audit_us_workbook(
    workbook_path: Path,
    *,
    provider: str,
    api_key: str,
    model: str,
    base_url: str,
    timeout: int,
    sleep_seconds: float,
    disable_repair: bool,
    max_fields: int,
    price_input_per_1m: Optional[float],
    price_output_per_1m: Optional[float],
) -> pd.DataFrame:
    cases = build_cases_from_workbook(workbook_path)
    if max_fields > 0:
        cases = cases[:max_fields]

    rows: List[Dict[str, Any]] = []
    for i, case in enumerate(cases, start=1):
        llm_error = None
        usage: Dict[str, Optional[int]] = {"prompt_tokens": None, "completion_tokens": None, "total_tokens": None}
        try:
            last_exc: Optional[Exception] = None
            llm_obj = None
            for attempt in range(1, LLM_CALL_MAX_RETRIES + 1):
                try:
                    llm_obj = base.call_verifier(
                        base.FieldCase(
                            workbook=case.workbook,
                            symbol=case.symbol,
                            company_id=case.company_id,
                            fiscal_year_end_date=case.fiscal_year_end_date,
                            statement=case.statement,
                            field_name=case.field_name,
                            rule_value=case.rule_value,
                            rule_status=case.rule_status,
                            candidate_codes=case.candidate_codes,
                            candidates=case.candidates,
                        ),
                        provider=provider,
                        api_key=api_key,
                        model=model,
                        base_url=base_url,
                        timeout=timeout,
                    )
                    last_exc = None
                    break
                except Exception as e:
                    last_exc = e
                    if attempt < LLM_CALL_MAX_RETRIES:
                        # Retries help absorb transient DNS/network/API hiccups.
                        time.sleep(min(4.0, float(attempt)))
            if llm_obj is None:
                raise last_exc or RuntimeError("llm_call_failed")
            usage_raw = llm_obj.get("_usage")
            if isinstance(usage_raw, dict):
                usage = {
                    "prompt_tokens": usage_raw.get("prompt_tokens"),
                    "completion_tokens": usage_raw.get("completion_tokens"),
                    "total_tokens": usage_raw.get("total_tokens"),
                }
            guarded = base.apply_repair_guardrails(
                base.FieldCase(
                    workbook=case.workbook,
                    symbol=case.symbol,
                    company_id=case.company_id,
                    fiscal_year_end_date=case.fiscal_year_end_date,
                    statement=case.statement,
                    field_name=case.field_name,
                    rule_value=case.rule_value,
                    rule_status=case.rule_status,
                    candidate_codes=case.candidate_codes,
                    candidates=case.candidates,
                ),
                llm_obj,
                repair_allowed_only_for_missing_or_parse_error=False,
                disable_repair=disable_repair,
            )
        except Exception as e:
            llm_error = str(e)
            guarded = {
                "decision": "NEED_REVIEW",
                "proposed_status": "UNCHANGED",
                "proposed_value": None,
                "reason": "llm_call_failed",
                "evidence": [],
                "evidence_item_codes": [],
                "repair_allowed": False,
                "repair_applied": False,
                "final_value": case.rule_value,
                "final_source": "rule",
                "review_required_recommended": True,
                "guard_fail": "llm_call_failed",
            }

        rows.append(
            {
                "workbook": case.workbook,
                "market": "us",
                "symbol": case.symbol,
                "company_id": case.company_id,
                "fiscal_year_end_date": case.fiscal_year_end_date,
                "statement": case.statement,
                "field_name": case.field_name,
                "rule_value": case.rule_value,
                "rule_status": case.rule_status,
                "candidate_codes_json": json.dumps(case.candidate_codes, ensure_ascii=False),
                "candidates_json": json.dumps(case.candidates, ensure_ascii=False),
                "expanded_reason": case.expanded_reason,
                "num_candidates_to_llm": case.num_candidates_to_llm,
                "llm_decision": guarded["decision"],
                "llm_value": guarded["proposed_value"],
                "llm_status": guarded["proposed_status"],
                "evidence_json": json.dumps(guarded["evidence"], ensure_ascii=False),
                "evidence_item_codes_json": json.dumps(guarded["evidence_item_codes"], ensure_ascii=False),
                "reason": guarded["reason"],
                "repair_allowed": int(bool(guarded["repair_allowed"])),
                "repair_applied": int(bool(guarded["repair_applied"])),
                "guard_fail": guarded["guard_fail"],
                "final_value": guarded["final_value"],
                "final_source": guarded["final_source"],
                "review_required_recommended": int(bool(guarded["review_required_recommended"])),
                "experiment_mode": ("verify_only" if disable_repair else "verify_repair"),
                "provider": provider,
                "model": model,
                "temperature": FIXED_TEMPERATURE,
                "prompt_tokens": usage.get("prompt_tokens"),
                "completion_tokens": usage.get("completion_tokens"),
                "total_tokens": usage.get("total_tokens"),
                "price_input_per_1m": price_input_per_1m,
                "price_output_per_1m": price_output_per_1m,
                "estimated_cost_usd": base._estimate_cost_usd(
                    usage.get("prompt_tokens"),
                    usage.get("completion_tokens"),
                    price_input_per_1m,
                    price_output_per_1m,
                ),
                "llm_error": llm_error,
            }
        )
        if sleep_seconds > 0 and i < len(cases):
            import time as _t
            _t.sleep(sleep_seconds)

    return pd.DataFrame(rows)


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(description="Run US LLM verify/repair layer on workbook(s).")
    p.add_argument("--workbook", help="Single US workbook path")
    p.add_argument("--input-dir", default="eval/outputs/us", help="Directory containing us_*_3statements.xlsx")
    p.add_argument("--pattern", default="us_*_3statements.xlsx", help="Glob pattern under input-dir")
    p.add_argument("--out-csv", default="eval/outputs/llm_us_audit.csv")
    p.add_argument("--provider", default=DEFAULT_PROVIDER, choices=["openai", "gemini", "deepseek", "claude"])
    p.add_argument("--model", default=DEFAULT_MODEL)
    p.add_argument("--base-url", default="")
    p.add_argument("--api-key", default="")
    p.add_argument("--sleep-seconds", type=float, default=0.0)
    p.add_argument("--timeout", type=int, default=REQUEST_TIMEOUT)
    p.add_argument("--max-fields", type=int, default=0)
    p.add_argument("--verify-only", action="store_true")
    p.add_argument("--price-input-per-1m", type=float, default=None)
    p.add_argument("--price-output-per-1m", type=float, default=None)
    return p.parse_args()


def main() -> None:
    args = parse_args()
    provider = (args.provider or DEFAULT_PROVIDER).strip().lower()
    api_key = base.resolve_api_key(provider, args.api_key)
    if not api_key:
        env_name = base.DEFAULT_API_KEY_ENV_BY_PROVIDER.get(provider, "OPENAI_API_KEY")
        raise SystemExit(f"Missing API key for provider={provider}. Set {env_name} or pass --api-key.")
    base_url = base.resolve_base_url(provider, args.base_url)

    if args.workbook:
        workbooks = [Path(args.workbook)]
    else:
        workbooks = find_us_workbooks(Path(args.input_dir), pattern=args.pattern)

    if not workbooks:
        raise SystemExit("No US workbook found.")

    frames = []
    for wb in workbooks:
        print(f"[INFO] Auditing US workbook: {wb}")
        df = audit_us_workbook(
            wb,
            provider=provider,
            api_key=api_key,
            model=args.model,
            base_url=base_url,
            timeout=args.timeout,
            sleep_seconds=args.sleep_seconds,
            disable_repair=args.verify_only,
            max_fields=args.max_fields,
            price_input_per_1m=args.price_input_per_1m,
            price_output_per_1m=args.price_output_per_1m,
        )
        frames.append(df)

    out_df = _merge_frames(frames)
    out_path = Path(args.out_csv)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_df.to_csv(out_path, index=False, encoding="utf-8-sig")
    print(f"[OK] Wrote audit CSV: {out_path}")
    print(f"[INFO] Rows: {len(out_df)}")


if __name__ == "__main__":
    main()
