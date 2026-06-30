#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""Repository smoke test.

This script does not download filings or call any LLM provider. It checks that
the public artifact has the expected files, importable dependencies, readable
CN schema, and syntactically valid Python modules.
"""

from __future__ import annotations

import argparse
import importlib.util
import py_compile
import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parent

REQUIRED_FILES = [
    "README.md",
    "requirements.txt",
    "schemas/CN_Schemas.xlsx",
    "extract_statements.py",
    "extract_consolidated_statements.py",
    "CN/export_three_statements_excel_cn.py",
    "US/export_three_statements_excel.py",
    "US/extract_xbrl_cash_flow.py",
    "JP/export_three_statements_excel_jp.py",
    "llm_cn_verifier.py",
]

REQUIRED_PACKAGES = [
    ("pandas", "pandas"),
    ("openpyxl", "openpyxl"),
    ("requests", "requests"),
    ("pdfplumber", "pdfplumber"),
    ("rapidfuzz", "rapidfuzz"),
    ("lxml", "lxml"),
    ("camelot", "camelot-py[cv]"),
]

REQUIRED_CN_SCHEMA_SHEETS = {
    "CN_FIN_BS_GEN",
    "CN_FIN_IS_GEN",
    "CN_FIN_CF_GEN",
}


def ok(message: str) -> None:
    print(f"[OK] {message}")


def fail(message: str) -> None:
    print(f"[FAIL] {message}")


def check_files() -> bool:
    missing = [p for p in REQUIRED_FILES if not (ROOT / p).exists()]
    if missing:
        fail("Missing required files:")
        for path in missing:
            print(f"  - {path}")
        return False
    ok("Required files are present")
    return True


def check_packages() -> bool:
    missing = []
    for import_name, package_name in REQUIRED_PACKAGES:
        if importlib.util.find_spec(import_name) is None:
            missing.append(package_name)
    if missing:
        fail("Missing Python packages. Install with: pip install -r requirements.txt")
        for name in missing:
            print(f"  - {name}")
        return False
    ok("Required Python packages are importable")
    return True


def check_cn_schema() -> bool:
    try:
        import pandas as pd

        schema_path = ROOT / "schemas" / "CN_Schemas.xlsx"
        xl = pd.ExcelFile(schema_path)
        sheets = set(xl.sheet_names)
    except Exception as exc:
        fail(f"Could not read schemas/CN_Schemas.xlsx: {exc}")
        return False

    missing = sorted(REQUIRED_CN_SCHEMA_SHEETS - sheets)
    if missing:
        fail("CN schema is missing required sheets:")
        for sheet in missing:
            print(f"  - {sheet}")
        return False
    ok("CN schema workbook is readable")
    return True


def check_python_syntax() -> bool:
    failed = []
    for path in sorted(ROOT.rglob("*.py")):
        if any(part in {".venv", "venv", "__pycache__"} for part in path.parts):
            continue
        try:
            py_compile.compile(str(path), doraise=True)
        except py_compile.PyCompileError as exc:
            failed.append((path.relative_to(ROOT), exc.msg))

    if failed:
        fail("Python syntax check failed:")
        for rel_path, msg in failed:
            print(f"  - {rel_path}: {msg}")
        return False
    ok("Python syntax check passed")
    return True


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Run repository smoke checks.")
    parser.add_argument(
        "--skip-packages",
        action="store_true",
        help="Skip package import checks.",
    )
    parser.add_argument(
        "--skip-compile",
        action="store_true",
        help="Skip Python syntax compilation checks.",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    checks = [
        check_files(),
        True if args.skip_packages else check_packages(),
        check_cn_schema(),
        True if args.skip_compile else check_python_syntax(),
    ]
    if all(checks):
        print("\nSmoke test passed.")
        return 0
    print("\nSmoke test failed.")
    return 1


if __name__ == "__main__":
    raise SystemExit(main())

