#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Run US LLM-only extraction baseline on existing workbook candidates.

This baseline intentionally does not use rule_value for decisions.
"""

from __future__ import annotations

import argparse
import importlib.util
import json
import math
import os
import sys
import time
from pathlib import Path
from typing import Any, Dict, List, Optional

import pandas as pd
import requests

ROOT_DIR = Path(__file__).resolve().parent.parent
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

from llm_cn_verifier import DEFAULT_BASE_URL, FIXED_TEMPERATURE, REQUEST_TIMEOUT  # noqa: E402


def _load_us_verifier_module():
    mod_path = ROOT_DIR / "eval" / "table2exp" / "llm_us_verifier.py"
    spec = importlib.util.spec_from_file_location("llm_us_verifier_impl", mod_path)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"Failed to load module spec from {mod_path}")
    mod = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = mod
    spec.loader.exec_module(mod)
    return mod


US_VERIFIER = _load_us_verifier_module()
DEFAULT_MODEL = "gpt-4o-2024-11-20"

SYSTEM_PROMPT = """You are a strict extractor for US annual-report financial fields.

Task:
- Given candidate rows only, decide whether one value can be extracted.
- Return EXTRACT only when evidence is concrete.
- If uncertain, return NEED_REVIEW.

Hard rules:
1) Do not use any external knowledge.
2) Do not fabricate numbers.
3) Cite item_code evidence.
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


def _parse_json_content(raw: str) -> Dict[str, Any]:
    s = (raw or "").strip()
    if s.startswith("```"):
        s = s.strip("`")
        if s.startswith("json"):
            s = s[4:].strip()
    return json.loads(s)


def _response_schema() -> Dict[str, Any]:
    return {
        "type": "json_schema",
        "json_schema": {
            "name": "us_llm_only_extraction",
            "strict": True,
            "schema": {
                "type": "object",
                "additionalProperties": False,
                "properties": {
                    "decision": {"type": "string", "enum": ["EXTRACT", "NEED_REVIEW"]},
                    "proposed_value": {"type": ["number", "null"]},
                    "proposed_status": {"type": "string", "enum": ["OK", "NOT_APPLICABLE", "MISSING", "PARSE_ERROR"]},
                    "evidence": {
                        "type": "array",
                        "items": {
                            "type": "object",
                            "additionalProperties": False,
                            "properties": {
                                "item_code": {"type": "string"},
                                "quote": {"type": "string"},
                                "why": {"type": "string"},
                            },
                            "required": ["item_code", "quote", "why"],
                        },
                    },
                    "reason": {"type": "string"},
                },
                "required": ["decision", "proposed_value", "proposed_status", "evidence", "reason"],
            },
        },
    }


def call_openai_llm_only(
    case_payload: Dict[str, Any],
    *,
    api_key: str,
    model: str,
    base_url: str,
    timeout: int,
) -> Dict[str, Any]:
    endpoint = base_url.rstrip("/") + "/chat/completions"
    payload = {
        "model": model,
        "temperature": FIXED_TEMPERATURE,
        "messages": [
            {"role": "system", "content": SYSTEM_PROMPT},
            {
                "role": "user",
                "content": "Extract one field from candidates and return strict JSON.\n\n"
                + json.dumps(case_payload, ensure_ascii=False),
            },
        ],
        "response_format": _response_schema(),
    }
    headers = {"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"}
    r = requests.post(endpoint, headers=headers, json=payload, timeout=timeout)
    r.raise_for_status()
    body = r.json()
    content = body["choices"][0]["message"]["content"]
    obj = _parse_json_content(content)
    obj["_raw_response"] = body
    return obj


def _candidate_index(candidates: List[Dict[str, Any]]) -> Dict[str, Dict[str, Any]]:
    out: Dict[str, Dict[str, Any]] = {}
    for c in candidates:
        code = str(c.get("item_code") or "").strip()
        if code:
            out[code] = c
    return out


def apply_llm_only_guardrails(
    llm_obj: Dict[str, Any],
    candidates: List[Dict[str, Any]],
) -> Dict[str, Any]:
    decision = str(llm_obj.get("decision") or "").upper()
    proposed_status = str(llm_obj.get("proposed_status") or "").upper()
    proposed_value = _to_float(llm_obj.get("proposed_value"))
    evidence = llm_obj.get("evidence") or []
    reason = str(llm_obj.get("reason") or "")

    guard_fail = ""
    evidence_codes = []
    candidate_map = _candidate_index(candidates)
    supported_value = False
    supports_na = False

    for e in evidence:
        code = str((e or {}).get("item_code") or "").strip()
        if not code:
            continue
        evidence_codes.append(code)
        c = candidate_map.get(code)
        if c is None:
            continue
        c_val = _to_float(c.get("value"))
        if proposed_value is not None and c_val is not None and abs(c_val - proposed_value) <= max(1e-9, abs(c_val) * 1e-6):
            supported_value = True
        st = str(_to_primitive(c.get("status")) or "").upper()
        raw = str(_to_primitive(c.get("raw_text")) or "")
        if st == "NOT_APPLICABLE" or ("not applicable" in raw.lower()):
            supports_na = True

    if decision == "EXTRACT":
        if not evidence_codes:
            guard_fail = "missing_evidence"
            decision = "NEED_REVIEW"
        elif proposed_status == "OK":
            if proposed_value is None:
                guard_fail = "missing_proposed_value"
                decision = "NEED_REVIEW"
            elif not supported_value:
                guard_fail = "unsupported_extract_value"
                decision = "NEED_REVIEW"
        elif proposed_status == "NOT_APPLICABLE":
            if not supports_na:
                guard_fail = "unsupported_not_applicable"
                decision = "NEED_REVIEW"
        else:
            guard_fail = "invalid_extract_status"
            decision = "NEED_REVIEW"

    final_value = proposed_value if (decision == "EXTRACT" and proposed_status == "OK") else None
    review_required = decision != "EXTRACT"
    return {
        "decision": decision,
        "proposed_status": proposed_status,
        "proposed_value": proposed_value,
        "evidence": evidence,
        "evidence_item_codes": evidence_codes,
        "reason": reason,
        "guard_fail": guard_fail or None,
        "final_value": final_value,
        "final_source": "llm_only" if decision == "EXTRACT" else "llm_only_review",
        "review_required_recommended": review_required,
    }


def run_llm_only_on_workbook(
    workbook: Path,
    *,
    api_key: str,
    model: str,
    base_url: str,
    timeout: int,
    sleep_seconds: float,
    max_fields: int,
) -> pd.DataFrame:
    cases = US_VERIFIER.build_cases_from_workbook(workbook)
    if max_fields > 0:
        cases = cases[:max_fields]

    rows: List[Dict[str, Any]] = []
    for i, case in enumerate(cases, start=1):
        llm_error = None
        guarded = {}
        try:
            payload = {
                "market": "us",
                "symbol": case.symbol,
                "statement": case.statement,
                "field_name": case.field_name,
                "candidate_codes": case.candidate_codes,
                "candidates": case.candidates,
            }
            llm_obj = call_openai_llm_only(
                payload,
                api_key=api_key,
                model=model,
                base_url=base_url,
                timeout=timeout,
            )
            guarded = apply_llm_only_guardrails(llm_obj, case.candidates)
        except Exception as e:
            llm_error = str(e)
            guarded = {
                "decision": "NEED_REVIEW",
                "proposed_status": "MISSING",
                "proposed_value": None,
                "evidence": [],
                "evidence_item_codes": [],
                "reason": "llm_call_failed",
                "guard_fail": "llm_call_failed",
                "final_value": None,
                "final_source": "llm_only_review",
                "review_required_recommended": True,
            }

        rows.append(
            {
                "workbook": str(workbook),
                "market": "us",
                "symbol": case.symbol,
                "company_id": case.company_id,
                "fiscal_year_end_date": case.fiscal_year_end_date,
                "statement": case.statement,
                "field_name": case.field_name,
                "candidate_codes_json": json.dumps(case.candidate_codes, ensure_ascii=False),
                "candidates_json": json.dumps(case.candidates, ensure_ascii=False),
                "llm_only_decision": guarded["decision"],
                "llm_only_value": guarded["proposed_value"],
                "llm_only_status": guarded["proposed_status"],
                "evidence_json": json.dumps(guarded["evidence"], ensure_ascii=False),
                "evidence_item_codes_json": json.dumps(guarded["evidence_item_codes"], ensure_ascii=False),
                "llm_only_reason": guarded["reason"],
                "guard_fail": guarded["guard_fail"],
                "final_value": guarded["final_value"],
                "final_source": guarded["final_source"],
                "review_required_recommended": int(bool(guarded["review_required_recommended"])),
                "experiment_mode": "llm_only",
                "model": model,
                "temperature": FIXED_TEMPERATURE,
                "llm_error": llm_error,
            }
        )
        if sleep_seconds > 0 and i < len(cases):
            time.sleep(sleep_seconds)
    return pd.DataFrame(rows)


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(description="Run US LLM-only extraction baseline.")
    p.add_argument("--workbook", help="Single US workbook path")
    p.add_argument("--input-dir", default="eval/outputs/us", help="Directory containing us_*_3statements.xlsx")
    p.add_argument("--pattern", default="us_*_3statements.xlsx", help="Glob pattern under input-dir")
    p.add_argument("--out-csv", default="eval/outputs/llm_only_us.csv", help="Output CSV path")
    p.add_argument("--model", default=DEFAULT_MODEL, help="OpenAI model")
    p.add_argument("--base-url", default=DEFAULT_BASE_URL, help="OpenAI API base URL")
    p.add_argument("--api-key", default=os.getenv("OPENAI_API_KEY"), help="OpenAI API key (or env OPENAI_API_KEY)")
    p.add_argument("--sleep-seconds", type=float, default=0.0, help="Sleep between field calls")
    p.add_argument("--timeout", type=int, default=REQUEST_TIMEOUT, help="HTTP timeout seconds")
    p.add_argument("--max-fields", type=int, default=0, help="Only process first N fields per workbook (0 = all)")
    return p.parse_args()


def main() -> None:
    args = parse_args()
    if not args.api_key:
        raise SystemExit("Missing OpenAI API key. Set OPENAI_API_KEY or pass --api-key.")

    if args.workbook:
        workbooks = [Path(args.workbook)]
    else:
        workbooks = US_VERIFIER.find_us_workbooks(Path(args.input_dir), pattern=args.pattern)
    if not workbooks:
        raise SystemExit("No US workbook found.")

    frames = []
    for wb in workbooks:
        print(f"[INFO] LLM-only workbook: {wb}")
        df = run_llm_only_on_workbook(
            wb,
            api_key=args.api_key,
            model=args.model,
            base_url=args.base_url,
            timeout=args.timeout,
            sleep_seconds=args.sleep_seconds,
            max_fields=args.max_fields,
        )
        frames.append(df)

    out_df = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()
    out_path = Path(args.out_csv)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_df.to_csv(out_path, index=False, encoding="utf-8-sig")
    print(f"[OK] Wrote LLM-only CSV: {out_path}")
    print(f"[INFO] Rows: {len(out_df)}")


if __name__ == "__main__":
    main()
