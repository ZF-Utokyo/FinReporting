#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
CN deterministic-first LLM verifier/repair layer.

Design goals:
1) Rule output is primary source of truth.
2) LLM may only repair values with concrete evidence references.
3) All outputs are traceable: rule_value, llm_decision, final_value.
"""

from __future__ import annotations

import argparse
import json
import math
import os
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Tuple

import pandas as pd
import requests


DEFAULT_PROVIDER = "openai"
DEFAULT_MODEL = "gpt-4o-2024-11-20"
DEFAULT_BASE_URL = "https://api.openai.com/v1"
DEFAULT_BASE_URL_BY_PROVIDER = {
    "openai": "https://api.openai.com/v1",
    "gemini": "https://generativelanguage.googleapis.com/v1beta/openai",
    "deepseek": "https://api.deepseek.com/v1",
    "claude": "https://api.anthropic.com/v1",
}
DEFAULT_API_KEY_ENV_BY_PROVIDER = {
    "openai": "OPENAI_API_KEY",
    "gemini": "GEMINI_API_KEY",
    "deepseek": "DEEPSEEK_API_KEY",
    "claude": "ANTHROPIC_API_KEY",
}
FIXED_TEMPERATURE = 0.0
REQUEST_TIMEOUT = 120

MARKET = "cn"
SHEETS = {
    "IS": "CN_FIN_IS",
    "BS": "CN_FIN_BS",
    "CF": "CN_FIN_CF",
}
RAW_SHEETS = {
    "IS": "RAW_CN_FIN_IS_GEN",
    "BS": "RAW_CN_FIN_BS_GEN",
    "CF": "RAW_CN_FIN_CF_GEN",
}

# Core fields used in current evaluation workflow.
FIELD_CONFIG: Dict[str, List[Tuple[str, List[str]]]] = {
    "IS": [
        ("total_revenue", ["BIZTOTINCO", "BIZINCO"]),
        ("operating_income", ["PERPROFIT"]),
        ("income_before_income_taxes", ["TOTPROFIT"]),
        ("net_income", ["NETPROFIT", "PARENETP"]),
        ("net_income_per_share_basic", ["BASICEPS"]),
        ("net_income_per_share_diluted", ["DILUTEDEPS"]),
    ],
    "BS": [
        ("total_assets", ["TOTASSET"]),
        ("cash_and_cash_equivalents", ["CURFDS"]),
        ("accounts_receivable", ["ACCORECE", "NOTESACCORECE"]),
        ("inventories", ["INVE"]),
        ("total_liabilities", ["TOTLIAB"]),
        ("total_shareholders_equity", ["PARESHARRIGH", "RIGHAGGR"]),
        ("total_liabilities_and_shareholders_equity", ["TOTLIABSHAREQUI"]),
    ],
    "CF": [
        ("net_income", ["NETPROFIT", "PARENETP"]),
        ("net_cash_operating", ["MANANETR"]),
        ("net_cash_investing", ["INVNETCASHFLOW"]),
        ("net_cash_financing", ["FINNETCFLOW"]),
        ("net_change_in_cash", ["CASHNETR"]),
        ("cash_end_of_period", ["FINALCASHBALA"]),
    ],
}


SYSTEM_PROMPT = """You are a strict financial extraction verifier for CN annual reports.

Task:
- Review one field case at a time.
- Decide KEEP / REPAIR / NEED_REVIEW.
- Use only provided candidate evidence.

Hard rules:
1) If evidence is weak or missing, choose NEED_REVIEW.
2) Do NOT fabricate numbers.
3) If proposing REPAIR, cite concrete candidate item codes in evidence.
4) If value is explicitly not applicable, use proposed_status=NOT_APPLICABLE and decision=KEEP.
"""

OUTPUT_CONTRACT = """Output contract (must follow exactly):
1) KEEP:
   - proposed_value must equal rule_value.
   - evidence_item_codes must contain at least 1 valid candidate item_code.
2) REPAIR:
   - proposed_value must be a concrete number from provided candidates.
   - evidence_item_codes must include the supporting candidate item_code(s).
3) NEED_REVIEW:
   - proposed_value must be null.
4) Do not use item codes outside provided candidates.
5) Return JSON only, no markdown fences.
"""

REASON_CODES = {
    "supported_by_candidate",
    "candidate_conflict",
    "missing_evidence",
    "no_supported_candidate",
    "not_applicable",
    "other",
}
VALID_DECISIONS = {"KEEP", "REPAIR", "NEED_REVIEW"}
VALID_PROPOSED_STATUS = {"OK", "MISSING", "PARSE_ERROR", "NOT_APPLICABLE", "UNCHANGED"}
MAX_CONTRACT_ATTEMPTS = 2


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


def _is_missing(v: Any) -> bool:
    return _to_primitive(v) is None


def _normalize_status(s: Any) -> str:
    x = str(_to_primitive(s) or "").strip().upper()
    if x in {"OK", "MISSING", "PARSE_ERROR", "NOT_APPLICABLE"}:
        return x
    return ""


def _derive_rule_status(rule_value: Any, candidates: List[Dict[str, Any]]) -> str:
    if not _is_missing(rule_value):
        return "OK"
    statuses = [_normalize_status(c.get("status")) for c in candidates]
    if "PARSE_ERROR" in statuses:
        return "PARSE_ERROR"
    if "NOT_APPLICABLE" in statuses:
        return "NOT_APPLICABLE"
    if "MISSING" in statuses:
        return "MISSING"
    return "MISSING"


def _read_row0(path: Path, sheet_name: str) -> Dict[str, Any]:
    df = pd.read_excel(path, sheet_name=sheet_name)
    if df.empty:
        return {}
    return {str(k): _to_primitive(v) for k, v in df.iloc[0].to_dict().items()}


def _read_raw_map(path: Path, sheet_name: str) -> Dict[str, Dict[str, Any]]:
    df = pd.read_excel(path, sheet_name=sheet_name)
    if df.empty or "item_code" not in df.columns:
        return {}
    out: Dict[str, Dict[str, Any]] = {}
    for _, r in df.iterrows():
        code = str(_to_primitive(r.get("item_code")) or "").strip()
        if not code:
            continue
        out[code] = {
            "item_code": code,
            "item_name": _to_primitive(r.get("item_name")),
            "value": _to_primitive(r.get("value")),
            "status": _normalize_status(r.get("status")),
            "raw_text": _to_primitive(r.get("raw_text")),
            "fixup_reason": _to_primitive(r.get("fixup_reason")),
        }
    return out


def build_cases_from_workbook(path: Path) -> List[FieldCase]:
    is_row = _read_row0(path, SHEETS["IS"])
    bs_row = _read_row0(path, SHEETS["BS"])
    cf_row = _read_row0(path, SHEETS["CF"])
    rows = {"IS": is_row, "BS": bs_row, "CF": cf_row}

    raw_maps = {
        "IS": _read_raw_map(path, RAW_SHEETS["IS"]),
        "BS": _read_raw_map(path, RAW_SHEETS["BS"]),
        "CF": _read_raw_map(path, RAW_SHEETS["CF"]),
    }

    symbol = str(is_row.get("symbol") or bs_row.get("symbol") or cf_row.get("symbol") or "")
    company_id = _to_primitive(is_row.get("company_id") or bs_row.get("company_id") or cf_row.get("company_id"))
    fye = _to_primitive(
        is_row.get("fiscal_year_end_date")
        or bs_row.get("fiscal_year_end_date")
        or cf_row.get("fiscal_year_end_date")
    )

    cases: List[FieldCase] = []
    for stmt, items in FIELD_CONFIG.items():
        row = rows.get(stmt, {})
        raw_map = raw_maps.get(stmt, {})
        for field_name, codes in items:
            candidates = [raw_map[c] for c in codes if c in raw_map]
            rule_value = row.get(field_name)
            rule_status = _derive_rule_status(rule_value, candidates)
            cases.append(
                FieldCase(
                    workbook=str(path),
                    symbol=symbol,
                    company_id=str(company_id) if company_id is not None else None,
                    fiscal_year_end_date=str(fye) if fye is not None else None,
                    statement=stmt,
                    field_name=field_name,
                    rule_value=_to_float(rule_value),
                    rule_status=rule_status,
                    candidate_codes=list(codes),
                    candidates=candidates,
                )
            )
    return cases


def _response_schema() -> Dict[str, Any]:
    return {
        "type": "json_schema",
        "json_schema": {
            "name": "cn_field_verification",
            "strict": True,
            "schema": {
                "type": "object",
                "additionalProperties": False,
                "properties": {
                    "decision": {"type": "string", "enum": ["KEEP", "REPAIR", "NEED_REVIEW"]},
                    "proposed_value": {"type": ["number", "null"]},
                    "proposed_status": {
                        "type": "string",
                        "enum": ["OK", "MISSING", "PARSE_ERROR", "NOT_APPLICABLE", "UNCHANGED"],
                    },
                    "evidence_item_codes": {
                        "type": "array",
                        "items": {"type": "string"},
                    },
                    "reason_code": {"type": "string"},
                    "reason": {"type": "string"},
                },
                "required": [
                    "decision",
                    "proposed_value",
                    "proposed_status",
                    "evidence_item_codes",
                    "reason_code",
                    "reason",
                ],
            },
        },
    }


def _parse_json_content(raw: str) -> Dict[str, Any]:
    s = (raw or "").strip()
    if s.startswith("```"):
        s = s.strip("`")
        if s.startswith("json"):
            s = s[4:].strip()
    return json.loads(s)


def _usage_from_openai_like(body: Dict[str, Any]) -> Dict[str, Optional[int]]:
    usage = body.get("usage") if isinstance(body, dict) else {}
    prompt_tokens = usage.get("prompt_tokens") if isinstance(usage, dict) else None
    completion_tokens = usage.get("completion_tokens") if isinstance(usage, dict) else None
    total_tokens = usage.get("total_tokens") if isinstance(usage, dict) else None
    return {
        "prompt_tokens": int(prompt_tokens) if prompt_tokens is not None else None,
        "completion_tokens": int(completion_tokens) if completion_tokens is not None else None,
        "total_tokens": int(total_tokens) if total_tokens is not None else None,
    }


def _usage_from_claude(body: Dict[str, Any]) -> Dict[str, Optional[int]]:
    usage = body.get("usage") if isinstance(body, dict) else {}
    in_tok = usage.get("input_tokens") if isinstance(usage, dict) else None
    out_tok = usage.get("output_tokens") if isinstance(usage, dict) else None
    prompt_tokens = int(in_tok) if in_tok is not None else None
    completion_tokens = int(out_tok) if out_tok is not None else None
    total_tokens = None
    if prompt_tokens is not None and completion_tokens is not None:
        total_tokens = prompt_tokens + completion_tokens
    return {
        "prompt_tokens": prompt_tokens,
        "completion_tokens": completion_tokens,
        "total_tokens": total_tokens,
    }


def _estimate_cost_usd(
    prompt_tokens: Optional[int],
    completion_tokens: Optional[int],
    price_input_per_1m: Optional[float],
    price_output_per_1m: Optional[float],
) -> Optional[float]:
    if (
        prompt_tokens is None
        or completion_tokens is None
        or price_input_per_1m is None
        or price_output_per_1m is None
    ):
        return None
    return (prompt_tokens / 1_000_000.0) * price_input_per_1m + (
        completion_tokens / 1_000_000.0
    ) * price_output_per_1m


def _build_user_prompt(case_payload: Dict[str, Any], retry_feedback: Optional[str] = None) -> str:
    out = [
        "Verify this field case and return strict JSON.",
        OUTPUT_CONTRACT,
    ]
    if retry_feedback:
        out.append("Previous output violated contract:")
        out.append(retry_feedback)
        out.append("Re-generate JSON that strictly follows the contract.")
    out.append(json.dumps(case_payload, ensure_ascii=False))
    return "\n\n".join(out)


def _dedupe_codes(codes: List[str]) -> List[str]:
    out: List[str] = []
    seen = set()
    for c in codes:
        cc = str(c or "").strip()
        if not cc or cc in seen:
            continue
        out.append(cc)
        seen.add(cc)
    return out


def _extract_evidence_item_codes(llm_obj: Dict[str, Any]) -> List[str]:
    raw_codes = llm_obj.get("evidence_item_codes")
    if isinstance(raw_codes, list):
        return _dedupe_codes([str(x or "") for x in raw_codes])

    raw_evidence = llm_obj.get("evidence")
    if isinstance(raw_evidence, list):
        codes = [str((e or {}).get("item_code") or "") for e in raw_evidence if isinstance(e, dict)]
        return _dedupe_codes(codes)

    return []


def _build_evidence_from_item_codes(case: FieldCase, codes: List[str], reason: str) -> List[Dict[str, Any]]:
    evidence: List[Dict[str, Any]] = []
    for code in _dedupe_codes(codes):
        c = _find_candidate(case, code)
        if c is None:
            continue
        quote = (
            f"value: {_to_primitive(c.get('value'))}, "
            f"status: {str(c.get('status') or '')}, "
            f"raw_text: {str(c.get('raw_text') or '')}"
        )
        evidence.append(
            {
                "item_code": code,
                "quote": quote,
                "why": reason or "Supported by candidate evidence.",
            }
        )
    return evidence


def _equal_or_both_none(a: Optional[float], b: Optional[float]) -> bool:
    if a is None and b is None:
        return True
    return _float_close(a, b)


def _has_conflicting_candidate_values(case: FieldCase) -> bool:
    vals: List[float] = []
    for c in case.candidates:
        v = _to_float(c.get("value"))
        if v is None:
            continue
        vals.append(v)
    if len(vals) <= 1:
        return False
    base = vals[0]
    for v in vals[1:]:
        if not _float_close(base, v):
            return True
    return False


def _normalize_and_validate_llm_output(
    case: FieldCase,
    llm_obj: Dict[str, Any],
) -> Tuple[Dict[str, Any], List[str]]:
    decision = str(llm_obj.get("decision") or "").upper().strip()
    proposed_status = str(llm_obj.get("proposed_status") or "UNCHANGED").upper().strip()
    proposed_value = _to_float(llm_obj.get("proposed_value"))
    reason = str(llm_obj.get("reason") or "").strip()
    reason_code = str(llm_obj.get("reason_code") or "other").strip().lower()
    if reason_code not in REASON_CODES:
        reason_code = "other"

    evidence_item_codes = _extract_evidence_item_codes(llm_obj)
    evidence = _build_evidence_from_item_codes(case, evidence_item_codes, reason)

    errors: List[str] = []
    if decision not in VALID_DECISIONS:
        errors.append("decision_invalid")
    if proposed_status not in VALID_PROPOSED_STATUS:
        errors.append("proposed_status_invalid")
    if any(_find_candidate(case, code) is None for code in evidence_item_codes):
        errors.append("evidence_item_codes_out_of_candidates")

    if decision in {"KEEP", "REPAIR"} and not evidence_item_codes:
        errors.append("missing_evidence_item_codes")
    if decision == "KEEP" and not _equal_or_both_none(proposed_value, case.rule_value):
        errors.append("keep_value_must_equal_rule_value")
    if decision == "KEEP" and _has_conflicting_candidate_values(case):
        errors.append("keep_not_allowed_with_conflicting_candidates")
    if decision == "REPAIR":
        if proposed_value is None:
            errors.append("repair_requires_proposed_value")
        else:
            supported = False
            for code in evidence_item_codes:
                c = _find_candidate(case, code)
                c_val = _to_float(c.get("value")) if c is not None else None
                if _float_close(c_val, proposed_value):
                    supported = True
                    break
            if not supported:
                errors.append("repair_value_not_supported_by_evidence_item_codes")
    if decision == "NEED_REVIEW" and proposed_value is not None:
        errors.append("need_review_requires_null_proposed_value")

    normalized = {
        "decision": decision,
        "proposed_value": proposed_value,
        "proposed_status": proposed_status,
        "evidence_item_codes": evidence_item_codes,
        "evidence": evidence,
        "reason_code": reason_code,
        "reason": reason or reason_code,
    }
    return normalized, errors


def call_openai_compatible_verifier(
    case: FieldCase,
    *,
    provider: str,
    api_key: str,
    model: str,
    base_url: str,
    timeout: int = REQUEST_TIMEOUT,
    retry_feedback: Optional[str] = None,
) -> Dict[str, Any]:
    endpoint = base_url.rstrip("/") + "/chat/completions"

    case_payload = {
        "market": MARKET,
        "symbol": case.symbol,
        "statement": case.statement,
        "field_name": case.field_name,
        "rule_value": case.rule_value,
        "rule_status": case.rule_status,
        "candidate_codes": case.candidate_codes,
        "candidates": case.candidates,
    }
    user_prompt = _build_user_prompt(case_payload, retry_feedback=retry_feedback)

    payload = {
        "model": model,
        "messages": [
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": user_prompt},
        ],
    }
    # Some OpenAI models (e.g., gpt-5-mini) only accept default temperature.
    model_lc = (model or "").strip().lower()
    if not model_lc.startswith("gpt-5-mini"):
        payload["temperature"] = FIXED_TEMPERATURE
    # DeepSeek currently supports json_object but not json_schema.
    if (provider or "").strip().lower() == "deepseek":
        payload["response_format"] = {"type": "json_object"}
    else:
        payload["response_format"] = _response_schema()
    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json",
    }
    r = requests.post(endpoint, headers=headers, json=payload, timeout=timeout)
    r.raise_for_status()
    body = r.json()
    content = body["choices"][0]["message"]["content"]
    obj = _parse_json_content(content)
    obj["_usage"] = _usage_from_openai_like(body)
    obj["_raw_response"] = body
    return obj


def call_claude_verifier(
    case: FieldCase,
    *,
    api_key: str,
    model: str,
    base_url: str,
    timeout: int = REQUEST_TIMEOUT,
    retry_feedback: Optional[str] = None,
) -> Dict[str, Any]:
    endpoint = base_url.rstrip("/") + "/messages"

    case_payload = {
        "market": MARKET,
        "symbol": case.symbol,
        "statement": case.statement,
        "field_name": case.field_name,
        "rule_value": case.rule_value,
        "rule_status": case.rule_status,
        "candidate_codes": case.candidate_codes,
        "candidates": case.candidates,
    }
    user_prompt = _build_user_prompt(case_payload, retry_feedback=retry_feedback)

    payload = {
        "model": model,
        "temperature": FIXED_TEMPERATURE,
        "max_tokens": 1200,
        "system": SYSTEM_PROMPT,
        "messages": [{"role": "user", "content": user_prompt}],
    }
    headers = {
        "x-api-key": api_key,
        "anthropic-version": "2023-06-01",
        "content-type": "application/json",
    }
    r = requests.post(endpoint, headers=headers, json=payload, timeout=timeout)
    r.raise_for_status()
    body = r.json()
    content_blocks = body.get("content") or []
    text_parts: List[str] = []
    for b in content_blocks:
        if isinstance(b, dict) and b.get("type") == "text":
            text_parts.append(str(b.get("text") or ""))
    content_text = "\n".join(x for x in text_parts if x)
    obj = _parse_json_content(content_text)
    obj["_usage"] = _usage_from_claude(body)
    obj["_raw_response"] = body
    return obj


def call_verifier(
    case: FieldCase,
    *,
    provider: str,
    api_key: str,
    model: str,
    base_url: str,
    timeout: int = REQUEST_TIMEOUT,
) -> Dict[str, Any]:
    p = (provider or DEFAULT_PROVIDER).strip().lower()
    retry_feedback = ""
    last_errors: List[str] = []
    for _ in range(MAX_CONTRACT_ATTEMPTS):
        if p == "claude":
            raw = call_claude_verifier(
                case,
                api_key=api_key,
                model=model,
                base_url=base_url,
                timeout=timeout,
                retry_feedback=retry_feedback or None,
            )
        else:
            raw = call_openai_compatible_verifier(
                case,
                provider=p,
                api_key=api_key,
                model=model,
                base_url=base_url,
                timeout=timeout,
                retry_feedback=retry_feedback or None,
            )

        normalized, errors = _normalize_and_validate_llm_output(case, raw)
        normalized["_usage"] = raw.get("_usage")
        normalized["_raw_response"] = raw.get("_raw_response")
        if not errors:
            return normalized
        last_errors = errors
        retry_feedback = "; ".join(errors)

    raise ValueError(f"llm_output_contract_failed: {', '.join(last_errors)}")


def _find_candidate(case: FieldCase, code: str) -> Optional[Dict[str, Any]]:
    for c in case.candidates:
        if str(c.get("item_code") or "") == code:
            return c
    return None


def _float_close(a: Optional[float], b: Optional[float]) -> bool:
    if a is None or b is None:
        return False
    if abs(a - b) <= 1e-9:
        return True
    scale = max(abs(a), abs(b), 1.0)
    return abs(a - b) / scale <= 1e-6


def apply_repair_guardrails(
    case: FieldCase,
    llm_obj: Dict[str, Any],
    *,
    repair_allowed_only_for_missing_or_parse_error: bool = True,
    disable_repair: bool = False,
) -> Dict[str, Any]:
    decision = str(llm_obj.get("decision") or "").upper()
    proposed_status = str(llm_obj.get("proposed_status") or "UNCHANGED").upper()
    proposed_value = _to_float(llm_obj.get("proposed_value"))
    evidence = llm_obj.get("evidence") or []
    if not evidence:
        evidence = _build_evidence_from_item_codes(case, _extract_evidence_item_codes(llm_obj), str(llm_obj.get("reason") or ""))
    reason = str(llm_obj.get("reason") or "")

    repair_allowed = True
    if repair_allowed_only_for_missing_or_parse_error:
        repair_allowed = case.rule_status in {"MISSING", "PARSE_ERROR"}

    evidence_codes: List[str] = []
    evidence_has_supported_number = False
    for e in evidence:
        code = str((e or {}).get("item_code") or "").strip()
        if not code:
            continue
        evidence_codes.append(code)
        c = _find_candidate(case, code)
        if c is None:
            continue
        c_val = _to_float(c.get("value"))
        if _float_close(c_val, proposed_value):
            evidence_has_supported_number = True

    valid_evidence = bool(evidence_codes)

    guard_fail = ""
    repair_applied = False
    final_value = case.rule_value
    final_source = "rule"
    review_required = decision == "NEED_REVIEW"

    if decision == "REPAIR":
        if disable_repair:
            guard_fail = "verify_only_mode"
        elif not repair_allowed:
            guard_fail = "repair_not_allowed_for_rule_status"
        elif _float_close(proposed_value, case.rule_value):
            guard_fail = "noop_repair_same_as_rule"
        elif not valid_evidence:
            guard_fail = "missing_evidence"
        elif proposed_value is None:
            guard_fail = "missing_proposed_value"
        elif not evidence_has_supported_number:
            guard_fail = "unsupported_repair_value"
        else:
            repair_applied = True
            final_value = proposed_value
            final_source = "llm_repair"
            review_required = False

        if guard_fail:
            decision = "NEED_REVIEW"
            review_required = True

    if decision == "KEEP":
        # KEEP without any evidence should still be sent to human review.
        if not valid_evidence:
            review_required = True

    return {
        "decision": decision,
        "proposed_status": proposed_status,
        "proposed_value": proposed_value,
        "reason": reason,
        "evidence": evidence,
        "evidence_item_codes": evidence_codes,
        "repair_allowed": repair_allowed,
        "repair_applied": repair_applied,
        "final_value": final_value,
        "final_source": final_source,
        "review_required_recommended": review_required,
        "guard_fail": guard_fail or None,
    }


def audit_cn_workbook(
    workbook_path: Path,
    *,
    provider: str = DEFAULT_PROVIDER,
    api_key: str,
    model: str = DEFAULT_MODEL,
    base_url: str = DEFAULT_BASE_URL,
    timeout: int = REQUEST_TIMEOUT,
    sleep_seconds: float = 0.0,
    repair_allowed_only_for_missing_or_parse_error: bool = True,
    disable_repair: bool = False,
    max_fields: int = 0,
    experiment_mode: str = "verify_repair",
    price_input_per_1m: Optional[float] = None,
    price_output_per_1m: Optional[float] = None,
) -> pd.DataFrame:
    cases = build_cases_from_workbook(workbook_path)
    if max_fields > 0:
        cases = cases[:max_fields]

    rows: List[Dict[str, Any]] = []
    for i, case in enumerate(cases, start=1):
        llm_error = None
        llm_obj: Dict[str, Any] = {}
        guarded: Dict[str, Any] = {}
        usage: Dict[str, Optional[int]] = {"prompt_tokens": None, "completion_tokens": None, "total_tokens": None}
        try:
            llm_obj = call_verifier(
                case,
                provider=provider,
                api_key=api_key,
                model=model,
                base_url=base_url,
                timeout=timeout,
            )
            usage_raw = llm_obj.get("_usage")
            if isinstance(usage_raw, dict):
                usage = {
                    "prompt_tokens": usage_raw.get("prompt_tokens"),
                    "completion_tokens": usage_raw.get("completion_tokens"),
                    "total_tokens": usage_raw.get("total_tokens"),
                }
            guarded = apply_repair_guardrails(
                case,
                llm_obj,
                repair_allowed_only_for_missing_or_parse_error=repair_allowed_only_for_missing_or_parse_error,
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
                "market": MARKET,
                "symbol": case.symbol,
                "company_id": case.company_id,
                "fiscal_year_end_date": case.fiscal_year_end_date,
                "statement": case.statement,
                "field_name": case.field_name,
                "rule_value": case.rule_value,
                "rule_status": case.rule_status,
                "candidate_codes_json": json.dumps(case.candidate_codes, ensure_ascii=False),
                "candidates_json": json.dumps(case.candidates, ensure_ascii=False),
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
                "experiment_mode": experiment_mode,
                "provider": provider,
                "model": model,
                "temperature": FIXED_TEMPERATURE,
                "prompt_tokens": usage.get("prompt_tokens"),
                "completion_tokens": usage.get("completion_tokens"),
                "total_tokens": usage.get("total_tokens"),
                "price_input_per_1m": price_input_per_1m,
                "price_output_per_1m": price_output_per_1m,
                "estimated_cost_usd": _estimate_cost_usd(
                    usage.get("prompt_tokens"),
                    usage.get("completion_tokens"),
                    price_input_per_1m,
                    price_output_per_1m,
                ),
                "llm_error": llm_error,
            }
        )
        if sleep_seconds > 0 and i < len(cases):
            time.sleep(sleep_seconds)

    return pd.DataFrame(rows)


def append_audit_sheet(workbook_path: Path, df: pd.DataFrame, sheet_name: str = "LLM_AUDIT_CN") -> None:
    with pd.ExcelWriter(
        workbook_path,
        engine="openpyxl",
        mode="a",
        if_sheet_exists="replace",
    ) as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)


def find_cn_workbooks(input_dir: Path, pattern: str = "cn_*_3statements.xlsx") -> List[Path]:
    if not input_dir.exists():
        return []
    return sorted(input_dir.glob(pattern))


def _merge_frames(frames: Iterable[pd.DataFrame]) -> pd.DataFrame:
    xs = [x for x in frames if x is not None and not x.empty]
    if not xs:
        return pd.DataFrame()
    return pd.concat(xs, ignore_index=True)


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(description="Run CN LLM verify/repair layer on workbook(s).")
    p.add_argument("--workbook", help="Single CN workbook path")
    p.add_argument("--input-dir", default="eval/outputs/cn", help="Directory containing cn_*_3statements.xlsx")
    p.add_argument("--pattern", default="cn_*_3statements.xlsx", help="Glob pattern under input-dir")
    p.add_argument("--out-csv", default="eval/outputs/llm_cn_audit.csv", help="Output CSV path")
    p.add_argument("--append-sheet", action="store_true", help="Write/replace LLM_AUDIT_CN sheet in workbook(s)")
    p.add_argument("--sheet-name", default="LLM_AUDIT_CN", help="Audit sheet name")
    p.add_argument(
        "--provider",
        default=DEFAULT_PROVIDER,
        choices=["openai", "gemini", "deepseek", "claude"],
        help="LLM provider",
    )
    p.add_argument("--model", default=DEFAULT_MODEL, help="LLM model id (fixed per experiment)")
    p.add_argument("--base-url", default="", help="Provider API base URL (empty uses provider default)")
    p.add_argument("--api-key", default="", help="API key (empty uses provider-specific env)")
    p.add_argument("--sleep-seconds", type=float, default=0.0, help="Sleep between field calls")
    p.add_argument("--timeout", type=int, default=REQUEST_TIMEOUT, help="HTTP timeout seconds")
    p.add_argument("--max-fields", type=int, default=0, help="Only process first N fields per workbook (0 = all)")
    p.add_argument("--verify-only", action="store_true", help="Disable all repairs (Rule + LLM-verify mode)")
    p.add_argument(
        "--price-input-per-1m",
        type=float,
        default=None,
        help="Input token price (USD per 1M tokens) for estimated_cost_usd",
    )
    p.add_argument(
        "--price-output-per-1m",
        type=float,
        default=None,
        help="Output token price (USD per 1M tokens) for estimated_cost_usd",
    )
    p.add_argument(
        "--allow-repair-all-status",
        action="store_true",
        help="Allow REPAIR even when rule_status is not MISSING/PARSE_ERROR (default: disabled)",
    )
    return p.parse_args()


def resolve_api_key(provider: str, cli_api_key: str) -> str:
    key = (cli_api_key or "").strip()
    if key:
        return key
    env_name = DEFAULT_API_KEY_ENV_BY_PROVIDER.get(provider, "OPENAI_API_KEY")
    return str(os.getenv(env_name, "") or "").strip()


def resolve_base_url(provider: str, cli_base_url: str) -> str:
    x = (cli_base_url or "").strip()
    if x:
        return x
    return DEFAULT_BASE_URL_BY_PROVIDER.get(provider, DEFAULT_BASE_URL)


def main() -> None:
    args = parse_args()
    provider = (args.provider or DEFAULT_PROVIDER).strip().lower()
    api_key = resolve_api_key(provider, args.api_key)
    if not api_key:
        env_name = DEFAULT_API_KEY_ENV_BY_PROVIDER.get(provider, "OPENAI_API_KEY")
        raise SystemExit(f"Missing API key for provider={provider}. Set {env_name} or pass --api-key.")
    base_url = resolve_base_url(provider, args.base_url)

    if args.workbook:
        workbooks = [Path(args.workbook)]
    else:
        workbooks = find_cn_workbooks(Path(args.input_dir), pattern=args.pattern)

    if not workbooks:
        raise SystemExit("No CN workbook found.")

    all_frames: List[pd.DataFrame] = []
    for wb in workbooks:
        print(f"[INFO] Auditing workbook: {wb}")
        df = audit_cn_workbook(
            wb,
            provider=provider,
            api_key=api_key,
            model=args.model,
            base_url=base_url,
            timeout=args.timeout,
            sleep_seconds=args.sleep_seconds,
            repair_allowed_only_for_missing_or_parse_error=(not args.allow_repair_all_status),
            disable_repair=args.verify_only,
            max_fields=args.max_fields,
            experiment_mode=("verify_only" if args.verify_only else "verify_repair"),
            price_input_per_1m=args.price_input_per_1m,
            price_output_per_1m=args.price_output_per_1m,
        )
        all_frames.append(df)
        if args.append_sheet:
            append_audit_sheet(wb, df, sheet_name=args.sheet_name)
            print(f"[OK] Updated sheet '{args.sheet_name}' in {wb}")

    out_df = _merge_frames(all_frames)
    out_path = Path(args.out_csv)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_df.to_csv(out_path, index=False, encoding="utf-8-sig")
    print(f"[OK] Wrote audit CSV: {out_path}")
    print(f"[INFO] Rows: {len(out_df)}")


if __name__ == "__main__":
    main()
