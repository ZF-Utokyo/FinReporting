#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
JP deterministic-first LLM verifier/repair layer.

Implementation reuses core logic from eval/table2exp/llm_us_verifier.py with
JP-specific workbook schema and market labeling.
"""

from __future__ import annotations

import argparse
import importlib.util
import sys
from pathlib import Path
from typing import Iterable

import pandas as pd

ROOT_DIR = Path(__file__).resolve().parent.parent
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))


def _load_us_core():
    mod_path = ROOT_DIR / "eval" / "table2exp" / "llm_us_verifier.py"
    spec = importlib.util.spec_from_file_location("llm_us_verifier_core_for_jp", mod_path)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"Failed to load module spec from {mod_path}")
    mod = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = mod
    spec.loader.exec_module(mod)
    return mod


core = _load_us_core()

# JP-specific prompt/market settings.
core.base.MARKET = "jp"
core.base.SYSTEM_PROMPT = """You are a strict financial extraction verifier for JP annual-report structured outputs.

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

core.DEFAULT_MODEL = "gpt-4o-2024-11-20"
core.SHEETS = {"IS": "JP_FIN_IS", "BS": "JP_FIN_BS", "CF": "JP_FIN_CF"}
core.RAW_CANDIDATE_SHEET = "RAW_JP_FIELD_CANDIDATES"
core.FIELD_CONFIG = {
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


def _merge_frames(frames: Iterable[pd.DataFrame]) -> pd.DataFrame:
    xs = [x for x in frames if x is not None and not x.empty]
    if not xs:
        return pd.DataFrame()
    return pd.concat(xs, ignore_index=True)


def find_jp_workbooks(input_dir: Path, pattern: str = "jp_*_3statements.xlsx"):
    if not input_dir.exists():
        return []
    return sorted(input_dir.glob(pattern))


def audit_jp_workbook(
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
    price_input_per_1m,
    price_output_per_1m,
) -> pd.DataFrame:
    df = core.audit_us_workbook(
        workbook_path,
        provider=provider,
        api_key=api_key,
        model=model,
        base_url=base_url,
        timeout=timeout,
        sleep_seconds=sleep_seconds,
        disable_repair=disable_repair,
        max_fields=max_fields,
        price_input_per_1m=price_input_per_1m,
        price_output_per_1m=price_output_per_1m,
    )
    if not df.empty and "market" in df.columns:
        df["market"] = "jp"
    return df


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(description="Run JP LLM verify/repair layer on workbook(s).")
    p.add_argument("--workbook", help="Single JP workbook path")
    p.add_argument("--input-dir", default="eval/outputs/jp", help="Directory containing jp_*_3statements.xlsx")
    p.add_argument("--pattern", default="jp_*_3statements.xlsx", help="Glob pattern under input-dir")
    p.add_argument("--out-csv", default="eval/outputs/llm_jp_audit.csv")
    p.add_argument("--provider", default=core.DEFAULT_PROVIDER, choices=["openai", "gemini", "deepseek", "claude"])
    p.add_argument("--model", default=core.DEFAULT_MODEL)
    p.add_argument("--base-url", default="")
    p.add_argument("--api-key", default="")
    p.add_argument("--sleep-seconds", type=float, default=0.0)
    p.add_argument("--timeout", type=int, default=core.REQUEST_TIMEOUT)
    p.add_argument("--max-fields", type=int, default=0)
    p.add_argument("--verify-only", action="store_true")
    p.add_argument("--price-input-per-1m", type=float, default=None)
    p.add_argument("--price-output-per-1m", type=float, default=None)
    return p.parse_args()


def main() -> None:
    args = parse_args()
    provider = (args.provider or core.DEFAULT_PROVIDER).strip().lower()
    api_key = core.base.resolve_api_key(provider, args.api_key)
    if not api_key:
        env_name = core.base.DEFAULT_API_KEY_ENV_BY_PROVIDER.get(provider, "OPENAI_API_KEY")
        raise SystemExit(f"Missing API key for provider={provider}. Set {env_name} or pass --api-key.")
    base_url = core.base.resolve_base_url(provider, args.base_url)

    if args.workbook:
        workbooks = [Path(args.workbook)]
    else:
        workbooks = find_jp_workbooks(Path(args.input_dir), pattern=args.pattern)
    if not workbooks:
        raise SystemExit("No JP workbook found.")

    frames = []
    for wb in workbooks:
        print(f"[INFO] Auditing JP workbook: {wb}")
        df = audit_jp_workbook(
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
