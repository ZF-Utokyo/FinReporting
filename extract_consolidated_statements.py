#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
提取合并财务报表（只提取本年年末数据）

Usage:
  python extract_consolidated_statements.py --pdf path/to/report.pdf --schema-file path/to/CN_Schemas.xlsx --out output.xlsx
"""

from __future__ import annotations

import argparse
import os
import re
import math
from datetime import date
from dataclasses import dataclass
from typing import Any, Dict, List, Optional, Tuple

import camelot
import pandas as pd
from rapidfuzz import fuzz, process

from extract_statements import run as locate_statements
import pdfplumber

try:
    from balance_sheet_anomaly import is_anomaly_marker_row, detect_aggregation_anomaly
except ImportError:
    def is_anomaly_marker_row(item_raw: str, numeric_values: list) -> bool:
        return False
    def detect_aggregation_anomaly(item_raw: str, matched_canonical_codes: list) -> Tuple[bool, dict]:
        return False, {}


# Core CN codes for LLM repair eligibility (evaluation-focused fields).
LLM_CORE_CODES_BY_STMT: Dict[str, set] = {
    "IS": {
        "BIZTOTINCO", "BIZINCO", "PERPROFIT", "TOTPROFIT",
        "NETPROFIT", "PARENETP", "BASICEPS", "DILUTEDEPS",
    },
    "BS": {
        "TOTASSET", "CURFDS", "ACCORECE", "NOTESACCORECE", "INVE",
        "TOTLIAB", "PARESHARRIGH", "RIGHAGGR", "TOTLIABSHAREQUI",
    },
    "CF": {
        "NETPROFIT", "PARENETP", "MANANETR", "INVNETCASHFLOW",
        "FINNETCFLOW", "CASHNETR", "FINALCASHBALA",
    },
}


# -----------------------
# Data structures
# -----------------------

@dataclass
class SchemaItem:
    statement_type: str  # BS/IS/CF
    item_code: str
    item_name_std: str


# -----------------------
# Schema loading
# -----------------------

def load_schemas(schema_file: str) -> Dict[str, List[SchemaItem]]:
    """
    Load schema from CN_Schemas.xlsx.
    Sheets: CN_FIN_BS_GEN, CN_FIN_IS_GEN, CN_FIN_CF_GEN
    """
    items: List[SchemaItem] = []
    
    if not os.path.exists(schema_file):
        raise FileNotFoundError(f"Schema file not found: {schema_file}")
    
    xl_file = pd.ExcelFile(schema_file)
    
    # Map sheet names to statement types (fixed order)
    sheet_to_type = {
        "CN_FIN_BS_GEN": "BS",
        "CN_FIN_IS_GEN": "IS",
        "CN_FIN_CF_GEN": "CF",
    }
    
    # Process sheets in fixed order to maintain schema order
    sheet_order = ["CN_FIN_BS_GEN", "CN_FIN_IS_GEN", "CN_FIN_CF_GEN"]
    
    # Group by statement type (maintain order per type)
    by_stmt: Dict[str, List[SchemaItem]] = {}
    
    for sheet_name in sheet_order:
        if sheet_name not in xl_file.sheet_names:
            continue
        if sheet_name not in sheet_to_type:
            continue
            
        statement_type = sheet_to_type[sheet_name]
        df = pd.read_excel(xl_file, sheet_name=sheet_name, header=None)
        
        # Initialize list for this statement type
        if statement_type not in by_stmt:
            by_stmt[statement_type] = []
        
        # Find header row (contains "Field Code", "Field Name")
        header_row = None
        for i in range(min(10, len(df))):
            row_vals = [str(x).strip() for x in df.iloc[i].tolist()]
            if "Field Code" in row_vals and "Field Name" in row_vals:
                header_row = i
                break
        
        if header_row is None:
            print(f"[WARN] Could not find header row in sheet {sheet_name}")
            continue
        
        # Parse data rows
        for i in range(header_row + 1, len(df)):
            row = df.iloc[i]
            if len(row) < 3:
                continue
                
            item_code = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ""
            item_name = str(row.iloc[2]).strip() if pd.notna(row.iloc[2]) else ""
            
            # Skip if not a valid row
            if not item_code or not item_name:
                continue
            
            # Include ALL fields from schema (including metadata fields like COMPCODE, PUBLISHDATE, etc.)
            # These will have empty values if not extracted from PDF, but should still be in output
            
            # Append directly to the statement type list (maintains order)
            by_stmt[statement_type].append(
                SchemaItem(
                    statement_type=statement_type,
                    item_code=item_code,
                    item_name_std=item_name,
                )
            )
    
    print(f"[INFO] Loaded schemas: BS={len(by_stmt.get('BS', []))}, IS={len(by_stmt.get('IS', []))}, CF={len(by_stmt.get('CF', []))}")
    return by_stmt


# -----------------------
# Table extraction
# -----------------------

def post_clean_table(df: pd.DataFrame) -> pd.DataFrame:
    """Clean extracted table."""
    if df is None or df.empty:
        return df
    
    df = df.copy()

    # Some lattice tables collapse an entire statement section into one giant row:
    # row0 = headers, row1 = newline-separated item/note/value lists.
    # Expand this shape to normal row-wise records before downstream matching.
    if len(df) == 2:
        header = [str(x).strip() if pd.notna(x) else "" for x in df.iloc[0].tolist()]
        body = [str(x) if pd.notna(x) else "" for x in df.iloc[1].tolist()]
        header_join = " ".join(header)
        header_join_norm = re.sub(r"\s+", "", header_join)
        line_counts = [c.replace("\r", "\n").count("\n") + 1 for c in body]
        if ("项目" in header_join_norm or "科目" in header_join_norm) and max(line_counts) >= 12:
            split_cols = [c.replace("\r", "\n").split("\n") for c in body]
            max_len = max(len(col_lines) for col_lines in split_cols)
            if max_len >= 12:
                rows = [header]
                for i in range(max_len):
                    row_i = []
                    for col_lines in split_cols:
                        v = col_lines[i] if i < len(col_lines) else ""
                        row_i.append(str(v).strip())
                    rows.append(row_i)
                df = pd.DataFrame(rows)

    # Remove completely empty rows/columns
    df = df.replace(r"^\s*$", None, regex=True)
    df = df.dropna(how="all").dropna(axis=1, how="all")
    
    if df.empty:
        return df
    
    # Use first row as header if it looks like headers
    first_row = df.iloc[0].tolist()
    first_row_norm = [re.sub(r"\s+", "", str(x)) for x in first_row]
    if any("项目" in x or "科目" in x for x in first_row_norm):
        raw_cols = [str(c).strip() for c in first_row]
        # De-duplicate header names (e.g. repeated "本期" columns) to avoid ambiguous DataFrame selection.
        seen = {}
        uniq_cols = []
        for c in raw_cols:
            k = c if c else "col"
            n = seen.get(k, 0)
            uniq_cols.append(k if n == 0 else f"{k}_{n}")
            seen[k] = n + 1
        df.columns = uniq_cols
        df = df.iloc[1:].reset_index(drop=True)
    else:
        # Rename columns to generic names
        df.columns = [f"col_{i}" for i in range(len(df.columns))]
    
    return df


def extract_statement_tables(
    pdf_path: str,
    page_start: int,
    page_end: int,
    flavor: str = "auto",
) -> List[pd.DataFrame]:
    """
    Extract tables from PDF pages using Camelot.
    flavor:
      - auto: lattice first, stream fallback when lattice finds nothing
      - lattice: only lattice
      - stream: only stream
    """
    pages = ",".join(str(p) for p in range(page_start, page_end + 1))
    tables_list = []
    
    if flavor not in {"auto", "lattice", "stream"}:
        flavor = "auto"

    # Try lattice first (unless stream-only)
    if flavor in {"auto", "lattice"}:
        try:
            tables_l = camelot.read_pdf(pdf_path, pages=pages, flavor="lattice")
            for t in tables_l:
                if t.df is not None and len(t.df) >= 2:
                    cleaned = post_clean_table(t.df)
                    if cleaned is not None and not cleaned.empty:
                        tables_list.append(cleaned)
        except Exception as e:
            print(f"[WARN] Lattice extraction failed: {e}")
    
    # Try stream when explicitly requested, or when auto mode had no usable lattice result.
    should_try_stream = (flavor == "stream") or (flavor == "auto" and len(tables_list) == 0)
    if should_try_stream:
        try:
            tables_s = camelot.read_pdf(pdf_path, pages=pages, flavor="stream")
            for t in tables_s:
                if t.df is not None and len(t.df) >= 2:
                    cleaned = post_clean_table(t.df)
                    if cleaned is not None and not cleaned.empty:
                        tables_list.append(cleaned)
        except Exception as e:
            print(f"[WARN] Stream extraction failed: {e}")
    
    print(f"[INFO] Extracted {len(tables_list)} tables from pages {page_start}-{page_end}")
    return tables_list


# -----------------------
# Text normalization
# -----------------------

def clean_item_name(s: str) -> str:
    """Normalize item name for matching."""
    if s is None:
        return ""
    s = str(s).strip()
    # Remove common prefixes
    s = re.sub(r"^[一二三四五六七八九十]+[、．.]", "", s)  # "一、" "二."
    s = re.sub(r"^[（(]?[一二三四五六七八九十0-9]+[)）]?", "", s)  # "(一)" "1)"
    s = re.sub(r"^(加|减|其中)[：:]", "", s)  # "加：" "减：" "其中："
    s = re.sub(r"\s+", "", s)  # Remove all whitespace
    s = s.replace("：", "").replace(":", "")
    # Remove trailing colon (for category headers)
    s = s.rstrip("：:")
    return s


def parse_number(x) -> Optional[float]:
    """Parse numeric value from cell."""
    if x is None:
        return None
    s = str(x).strip()
    if not s or s in {"—", "-", "–", "nan", "NaN"}:
        return None

    # Prefer money-like tokens (comma-grouped or >=4 digits) and keep
    # parenthesis-negatives. This avoids picking note numbers like "四(49)".
    money_tokens = re.findall(
        r"\(?-?\d{1,3}(?:,\d{3})+(?:\.\d+)?\)?|\(?-?\d{4,}(?:\.\d+)?\)?",
        s,
    )
    if money_tokens:
        tok = money_tokens[0]
        neg = tok.startswith("(") and tok.endswith(")")
        tok = tok.strip("()").replace(",", "").replace("，", "")
        try:
            v = float(tok)
            if neg and v > 0:
                v = -v
            return v
        except ValueError:
            pass

    # EPS-like rows can be small decimals.
    dec_tokens = re.findall(r"\(?-?\d+\.\d+\)?", s)
    if dec_tokens:
        tok = dec_tokens[0]
        neg = tok.startswith("(") and tok.endswith(")")
        tok = tok.strip("()")
        try:
            v = float(tok)
            if neg and v > 0:
                v = -v
            return v
        except ValueError:
            pass

    # Fallback: accept plain numeric-only cells.
    s_plain = s.replace(",", "").replace("，", "").replace(" ", "")
    if re.fullmatch(r"\(?-?\d+(?:\.\d+)?\)?", s_plain):
        neg = s_plain.startswith("(") and s_plain.endswith(")")
        s_plain = s_plain.strip("()")
        try:
            v = float(s_plain)
            if neg and v > 0:
                v = -v
            return v
        except ValueError:
            return None
    return None


def parse_eps_number(x) -> Optional[float]:
    """
    Parse EPS-like values while preserving decimal semantics.
    Supports OCR variants like 8.2062 / 8,2062 / 8．2062 and glued tokens.
    """
    if x is None:
        return None
    s = str(x).strip()
    if not s or s in {"—", "-", "–", "nan", "NaN"}:
        return None
    s = re.sub(r"\s+", "", s)

    tokens = re.findall(r"\(?-?\d+[.,，．]\d{1,4}\)?", s)
    if not tokens:
        return None

    def token_candidates(tok: str) -> List[float]:
        out: List[float] = []
        t = tok.strip()
        neg = t.startswith("(") and t.endswith(")")
        t = t.strip("()")
        t = t.replace("，", ",").replace("．", ".")
        if "," in t and "." not in t:
            t = t.replace(",", ".")
        try:
            v = float(t)
            if neg and v > 0:
                v = -v
            out.append(v)
            # Sticky note prefix in integer part, e.g. 5913.84 -> 13.84.
            if abs(v) > 200 and "." in t:
                left, right = t.split(".", 1)
                if len(left) >= 3:
                    for cut in (2, 1, 3):
                        l2 = left[cut:]
                        if not l2:
                            continue
                        try:
                            v2 = float(f"{l2}.{right}")
                        except Exception:
                            continue
                        if neg and v2 > 0:
                            v2 = -v2
                        out.append(v2)
        except Exception:
            pass
        return out

    for tok in tokens:
        for v in token_candidates(tok):
            if abs(v) <= 200:
                return v
    return None


NA_TOKENS = ("不适用", "N/A", "NA", "—", "--", "无", "不存在")
ALLOW_NOT_APPLICABLE_CODES = {"DILUTEDEPS"}


def _is_not_applicable_text(s: str) -> bool:
    if not s:
        return False
    t = re.sub(r"\s+", "", str(s)).upper()
    if "不适用" in t:
        return True
    if "不存在" in t:
        return True
    if re.search(r"(?<![A-Z])N/?A(?![A-Z])", t):
        return True
    if "--" in t or "—" in t:
        return True
    if t in {"无", "NA", "N/A"}:
        return True
    return False


def _status_rank(status: str) -> int:
    order = {
        "OK": 4,
        "NOT_APPLICABLE": 3,
        "PARSE_ERROR": 2,
        "MISSING": 1,
    }
    return order.get(status, 0)


def _merge_extracted_record(
    extracted: Dict[str, Dict[str, Any]],
    rec: Dict[str, Any],
) -> None:
    code = str(rec.get("item_code") or "").strip()
    if not code:
        return
    old = extracted.get(code)
    if old is None:
        extracted[code] = rec
        return
    if _status_rank(str(rec.get("status") or "")) > _status_rank(str(old.get("status") or "")):
        extracted[code] = rec


def _order_cols_for_current_period(t: pd.DataFrame, item_col: str, cols: List[str]) -> List[str]:
    """
    Sort candidate value columns by current-period priority.
    If no semantic signals are found, fallback to left-to-right order.
    """
    if not cols:
        return []

    col_pos = {c: i for i, c in enumerate([c for c in t.columns if c != item_col])}
    sem: Dict[str, int] = {}
    temporal: Dict[str, date] = {}

    for c in cols:
        blob = re.sub(r"\s+", "", _col_header_blob(t, c, header_rows=3, item_col=item_col))
        score = 0
        if re.search(r"(报告期末|本期末|期末余额|期末数|期末|年末|报告期|本期|本年|当期|本期发生额|本年发生额)", blob):
            score += 100
        if re.search(r"(期初余额|期初数|期初|上年年末|上期期末|上期末|上年|上期|上年度|去年|同期|前期)", blob):
            score -= 100
        if "合并" in blob:
            score += 20
        if "公司" in blob and "合并" not in blob:
            score -= 10
        sem[c] = score

        key = _latest_header_temporal_key(blob)
        if key is not None:
            temporal[c] = key

    if temporal:
        latest = max(temporal.values())
    else:
        latest = None

    def rank(col: str) -> Tuple[int, int]:
        score = sem.get(col, 0)
        if latest is not None:
            score += 50 if temporal.get(col) == latest else -20
        return (-score, col_pos.get(col, 10**9))

    ordered = sorted(cols, key=rank)
    if not sem or max(sem.values()) == 0 and min(sem.values()) == 0 and latest is None:
        ordered = sorted(cols, key=lambda c: col_pos.get(c, 10**9))
    return ordered


def page_has_table_features_local(text: str) -> bool:
    """Lightweight table-like page detector used during statement title search."""
    if not text:
        return False
    if "审计报告" in text and "审字" in text:
        return False
    if re.search(r"(项目|附注|本期|本年|期末|期初)", text):
        return True
    digits = len(re.findall(r"\d", text))
    return len(text) > 0 and digits > 40 and (digits / max(len(text), 1)) > 0.10


def has_likely_statement_title_line(text: str, stmt_name: str) -> bool:
    """
    Check whether a page contains a short, title-like statement line rather than
    a sentence mentioning the statement (e.g., in audit report paragraphs).
    """
    title_map = {
        "资产负债表": ["合并资产负债表", "合并及公司资产负债表"],
        "利润表": ["合并利润表", "合并及公司利润表"],
        "现金流量表": ["合并现金流量表", "合并及公司现金流量表"],
    }
    titles = title_map.get(stmt_name, [])
    if not titles:
        return False

    lines = [re.sub(r"\s+", "", ln) for ln in text.splitlines()[:60]]
    for ln in lines:
        if not ln:
            continue
        if len(ln) > 28:
            continue
        if any(t in ln for t in titles):
            return True
    return False


def _extract_dates_from_text(text: str) -> List[date]:
    out: List[date] = []
    if not text:
        return out
    s = str(text)
    # Supports: YYYY年MM月DD日, YYYY-MM-DD, YYYY/MM/DD, YYYY.MM.DD
    for m in re.finditer(r"(20\d{2})\s*[年./-]\s*(\d{1,2})\s*[月./-]\s*(\d{1,2})\s*日?", s):
        try:
            y = int(m.group(1))
            mm = int(m.group(2))
            dd = int(m.group(3))
            out.append(date(y, mm, dd))
        except Exception:
            continue
    return out


def _extract_years_from_text(text: str) -> List[int]:
    out: List[int] = []
    if not text:
        return out
    s = str(text)
    # Supports: YYYY年度 / YYYY年 / YYYY
    for m in re.finditer(r"(20\d{2})\s*年度", s):
        out.append(int(m.group(1)))
    for m in re.finditer(r"(20\d{2})\s*年", s):
        out.append(int(m.group(1)))
    for m in re.finditer(r"\b(20\d{2})\b", s):
        out.append(int(m.group(1)))
    return out


def _latest_header_temporal_key(text: str) -> Optional[date]:
    """
    Parse the latest year/date from a header blob.
    Priority: full date > year-only.
    """
    ds = _extract_dates_from_text(text)
    if ds:
        return max(ds)
    ys = _extract_years_from_text(text)
    if ys:
        return date(max(ys), 12, 31)
    return None


def _col_header_blob(
    t: pd.DataFrame,
    col: str,
    header_rows: int = 3,
    item_col: Optional[str] = None,
) -> str:
    vals = [str(col)]
    n = len(t)
    if n <= 0:
        return " ".join(vals)

    # Prefer semantic header rows around the row containing "项目/科目".
    # This is more robust than blindly taking the first 1-3 rows because
    # some PDFs insert title/unit lines above the true column header.
    if item_col and item_col in t.columns:
        anchor = None
        probe_n = min(12, n)
        for i in range(probe_n):
            s = str(t.iloc[i].get(item_col, "") or "")
            s = re.sub(r"\s+", "", s)
            if re.search(r"(项目|科目)", s):
                anchor = i
                break
        if anchor is not None:
            start = max(0, anchor - (header_rows - 1))
            end = anchor
            for i in range(start, end + 1):
                v = t.iloc[i].get(col, "")
                if pd.notna(v):
                    vals.append(str(v))
            return " ".join(vals)

    # Fallback to top rows.
    n_head = min(header_rows, n)
    for i in range(n_head):
        v = t.iloc[i].get(col, "")
        if pd.notna(v):
            vals.append(str(v))
    return " ".join(vals)


def _numeric_density(t: pd.DataFrame, col: str, header_rows: int = 3) -> float:
    """
    Rough numeric-density estimate for selecting candidate value columns.
    """
    if col not in t.columns:
        return 0.0
    vals = t[col].tolist()[header_rows:]
    if not vals:
        vals = t[col].tolist()
    total = 0
    hits = 0
    for v in vals:
        s = str(v).strip() if v is not None else ""
        if not s:
            continue
        total += 1
        if parse_number(v) is not None:
            hits += 1
    if total == 0:
        return 0.0
    return hits / total


def choose_best_value_col(t: pd.DataFrame, item_col: str, statement_type: str = "") -> Optional[str]:
    """
    Choose the best current-year amount column.
    Works for both header-rich tables and generic col_* tables.
    """
    candidates = [c for c in t.columns if c != item_col]
    if not candidates:
        return None

    numeric_candidates = [c for c in candidates if _numeric_density(t, c, header_rows=3) >= 0.2]
    if not numeric_candidates:
        numeric_candidates = candidates

    # Rule 1 (date/year-driven): lock to the latest year/date column.
    # Header text is built by concatenating the first 1-3 rows.
    col_temporal: Dict[str, date] = {}
    for c in numeric_candidates:
        blob = _col_header_blob(t, c, header_rows=3, item_col=item_col)
        key = _latest_header_temporal_key(blob)
        if key is not None:
            col_temporal[c] = key
    if col_temporal:
        return max(numeric_candidates, key=lambda c: col_temporal.get(c, date.min))

    # Rule 2 (semantic fallback): current/end-period terms > prior/begin terms.
    sem_score: Dict[str, int] = {}
    for c in numeric_candidates:
        blob = re.sub(r"\s+", "", _col_header_blob(t, c, header_rows=3, item_col=item_col))
        score = 0
        if re.search(r"(报告期末|本期末|期末余额|期末数|期末|年末|报告期|本期|本年|当期|本期发生额|本年发生额)", blob):
            score += 100
        if re.search(r"(期初余额|期初数|期初|上年年末|上期期末|上期末|上年|上期|上年度|去年|同期|前期)", blob):
            score -= 100
        sem_score[c] = score
    if sem_score and max(sem_score.values()) > 0:
        return max(numeric_candidates, key=lambda c: sem_score.get(c, 0))

    best_col = None
    best_score = -10**9

    for col in numeric_candidates:
        col_str = _col_header_blob(t, col, header_rows=3, item_col=item_col)
        score = 0.0

        # Header-based hints.
        if any(x in col_str for x in ["附注", "注释", "注"]):
            score -= 120
        if any(x in col_str for x in ["上年", "上期", "期初", "去年", "上年度"]):
            score -= 80
        if any(x in col_str for x in ["本年", "本期", "期末", "12月31日", "当期"]):
            score += 120
        if "合并" in col_str and any(x in col_str for x in ["本年", "本期", "期末", "金额", "余额"]):
            score += 60

        # Content-based hints.
        vals = []
        raw_cells = []
        for v in t[col].tolist():
            raw = str(v).strip() if v is not None else ""
            if not raw:
                continue
            raw_cells.append(raw)
            n = parse_number(v)
            if n is not None:
                vals.append(n)

        cnt = len(vals)
        if cnt == 0:
            score -= 200
        else:
            abs_vals = [abs(v) for v in vals]
            big_cnt = sum(1 for a in abs_vals if a >= 10000)
            small_cnt = sum(1 for a in abs_vals if a <= 999)
            mag_score = sum(min(12.0, math.log10(a + 1.0)) for a in abs_vals)

            score += cnt * 6
            score += big_cnt * 10
            score -= small_cnt * 2
            score += mag_score

            # Note-like columns often are mostly small integers.
            if big_cnt == 0 and (small_cnt / max(cnt, 1)) >= 0.7:
                score -= 140

            # Penalize columns with many Chinese chars in numeric-looking cells (e.g., "五、54").
            zh_ratio = 0.0
            if raw_cells:
                zh_cells = sum(1 for s in raw_cells if re.search(r"[\u4e00-\u9fff]", s))
                zh_ratio = zh_cells / len(raw_cells)
            if zh_ratio >= 0.2:
                score -= 100

        if score > best_score:
            best_score = score
            best_col = col

    return best_col


def choose_bs_period_end_col(t: pd.DataFrame, item_col: str) -> Optional[str]:
    """
    Strict Step A for CN BS TOTAL_ASSETS:
    - Build header blob from top semantic header rows.
    - Prefer latest date column.
    - Fallback to period-end keywords and avoid prior-period keywords.
    """
    candidates = [c for c in t.columns if c != item_col]
    if not candidates:
        return None

    # BS tables often have multi-row headers (year/date + 合并/公司).
    # Use deeper header skip for density estimation.
    numeric_candidates = [c for c in candidates if _numeric_density(t, c, header_rows=6) >= 0.2]
    if not numeric_candidates:
        numeric_candidates = candidates

    def bs_header_blob(col: str) -> str:
        vals = [str(col)]
        # Collect broader top rows to capture merged header semantics.
        probe_n = min(8, len(t))
        for i in range(probe_n):
            v = t.iloc[i].get(col, "")
            if pd.notna(v):
                vals.append(str(v))
        return " ".join(vals)

    # A1) Temporal + scope semantics.
    # Prefer latest period-end year/date + 合并 column.
    col_meta: Dict[str, Dict[str, Any]] = {}
    for c in numeric_candidates:
        blob = bs_header_blob(c)
        blob_n = re.sub(r"\s+", "", blob)
        ds = _extract_dates_from_text(blob)
        ys = _extract_years_from_text(blob)
        latest_dt = max(ds) if ds else None
        latest_year = max(ys) if ys else None
        col_meta[c] = {
            "blob_n": blob_n,
            "latest_dt": latest_dt,
            "latest_year": latest_year,
            "is_consolidated": ("合并" in blob_n),
            "is_company_only": ("公司" in blob_n and "合并" not in blob_n),
        }

    has_temporal = any((m["latest_dt"] is not None or m["latest_year"] is not None) for m in col_meta.values())
    if has_temporal:
        # Normalize temporal key to date for easy comparison.
        def temporal_key(col: str) -> date:
            m = col_meta[col]
            if m["latest_dt"] is not None:
                return m["latest_dt"]
            if m["latest_year"] is not None:
                return date(int(m["latest_year"]), 12, 31)
            return date.min

        max_key = max(temporal_key(c) for c in numeric_candidates)
        temporal_cols = [c for c in numeric_candidates if temporal_key(c) == max_key]

        # Scope preference: consolidated > unspecified > company-only.
        def scope_rank(col: str) -> int:
            m = col_meta[col]
            if m["is_consolidated"]:
                return 2
            if m["is_company_only"]:
                return 0
            return 1

        best_scope = max(scope_rank(c) for c in temporal_cols)
        scope_cols = [c for c in temporal_cols if scope_rank(c) == best_scope]
        # Keep left-most among ties.
        return scope_cols[0]

    # A2) Semantic period-end fallback.
    pos_keywords = ("期末", "报告期末", "本期末")
    neg_keywords = ("期初", "上期末", "上年年末", "上年末")
    sem_score: Dict[str, int] = {}
    for c in numeric_candidates:
        blob_n = str(col_meta.get(c, {}).get("blob_n") or "")
        score = 0
        if any(k in blob_n for k in pos_keywords):
            score += 100
        if any(k in blob_n for k in neg_keywords):
            score -= 100
        if "合并" in blob_n:
            score += 40
        if "公司" in blob_n and "合并" not in blob_n:
            score -= 20
        sem_score[c] = score
    if sem_score and max(sem_score.values()) > 0:
        return max(numeric_candidates, key=lambda c: sem_score.get(c, 0))

    return choose_best_value_col(t, item_col, statement_type="BS")


def _extract_bs_total_assets_from_table(
    t: pd.DataFrame,
    item_col: str,
) -> Tuple[Optional[float], Optional[str]]:
    """
    Strict Step B/C for CN BS TOTAL_ASSETS:
    - Row label priority: 资产总计 > 资产总额 > 总资产
    - Exclude section subtotals (流动/非流动资产合计等) and clear sub-items.
    - Allow 资产合计 only as final grand-total fallback.
    """
    value_col = choose_bs_period_end_col(t, item_col)
    if value_col is None:
        return None, None
    other_cols = [c for c in t.columns if c not in {item_col, value_col}]

    explicit_candidates: List[Tuple[int, int, float, str]] = []
    asset_heji_candidates: List[Tuple[int, float, str]] = []
    n_rows = len(t)

    for idx, r in t.iterrows():
        raw_item = str(r.get(item_col, "")).strip()
        if not raw_item:
            continue
        label = clean_item_name(raw_item)
        if not label:
            continue

        # Exclusions for subsection subtotals / sub-items.
        if any(x in label for x in ("流动资产合计", "流动资产总计", "非流动资产合计", "非流动资产总计")):
            continue
        if any(x in label for x in ("流动资产", "非流动资产")):
            continue
        if "其中" in label:
            continue

        priority: Optional[int] = None
        if label == "资产总计":
            priority = 1
        elif label == "资产总额":
            priority = 2
        elif label == "总资产":
            priority = 3
        elif label == "资产合计":
            # Only allow as final grand-total fallback.
            if n_rows > 0 and idx >= int(n_rows * 0.6):
                v = parse_number(r.get(value_col))
                if v is None:
                    for c in other_cols:
                        v_alt = parse_number(r.get(c))
                        if v_alt is not None:
                            v = v_alt
                            break
                if v is not None:
                    asset_heji_candidates.append((idx, float(v), raw_item))
            continue

        if priority is None:
            continue

        v = parse_number(r.get(value_col))
        if v is None:
            for c in other_cols:
                v_alt = parse_number(r.get(c))
                if v_alt is not None:
                    v = v_alt
                    break
        if v is None:
            continue
        explicit_candidates.append((priority, idx, float(v), raw_item))

    if explicit_candidates:
        explicit_candidates.sort(key=lambda x: (x[0], x[1]))
        _, _, v, raw_item = explicit_candidates[0]
        return v, raw_item

    if asset_heji_candidates:
        # Prefer the last grand-total style row.
        asset_heji_candidates.sort(key=lambda x: x[0], reverse=True)
        _, v, raw_item = asset_heji_candidates[0]
        return v, raw_item
    return None, None


def is_likely_cf_table(t: pd.DataFrame, item_col: str) -> bool:
    """Heuristic filter to keep real cash-flow tables and drop unrelated noise."""
    anchors = [
        "经营活动产生的现金流量净额",
        "投资活动产生的现金流量净额",
        "筹资活动产生的现金流量净额",
        "现金及现金等价物净增加额",
        "销售商品、提供劳务收到的现金",
    ]
    exclusion_tokens = ["股东权益", "本年增减变动", "所有者权益变动表"]
    text_cells = []
    for v in t[item_col].tolist():
        if v is None:
            continue
        s = str(v).strip()
        if s:
            text_cells.append(s)
    if not text_cells:
        return False

    joined = "\n".join(text_cells[:200])
    if any(tok in joined for tok in exclusion_tokens):
        return False
    return any(a in joined for a in anchors)


# -----------------------
# Item matching
# -----------------------

def build_match_index(schema_items: List[SchemaItem]) -> Dict[str, str]:
    """Build index: normalized_name -> item_code."""
    idx: Dict[str, str] = {}
    for it in schema_items:
        key = clean_item_name(it.item_name_std)
        if key:
            idx[key] = it.item_code
    return idx


def match_item(
    raw_item: str,
    schema_items: List[SchemaItem],
    match_index: Dict[str, str],
    min_score: int = 75,  # Lower threshold to catch more matches
) -> Tuple[Optional[str], float, str]:
    """Match raw item name to schema item."""
    key = clean_item_name(raw_item)
    if not key:
        return None, 0.0, ""
    
    # 1) Exact match
    if key in match_index:
        return match_index[key], 1.0, key
    
    # 2) Try partial match (if key is substring of schema name or vice versa)
    for schema_key, item_code in match_index.items():
        if key in schema_key or schema_key in key:
            if len(key) >= 3 and len(schema_key) >= 3:  # Both should be meaningful
                return item_code, 0.95, schema_key
    
    # 3) Fuzzy match
    candidates = list(match_index.keys())
    if not candidates:
        return None, 0.0, ""
    
    best = process.extractOne(key, candidates, scorer=fuzz.ratio)
    if not best:
        return None, 0.0, ""
    
    matched_name, score, _ = best
    if score >= min_score:
        return match_index[matched_name], score / 100.0, matched_name
    
    return None, score / 100.0, matched_name


# -----------------------
# Extract consolidated statements only
# -----------------------

def search_consolidated_statements(
    pdf_path: str,
    financial_section: Optional[Dict[str, int]],
    found_statements: Dict[str, Tuple[int, int]],
) -> Dict[str, Tuple[int, int]]:
    """Search for consolidated statements in financial section."""
    consolidated_patterns = {
        "资产负债表": re.compile(r"合并(?:及公司)?资产负债表"),
        "利润表": re.compile(r"合并(?:及公司)?利润表"),
        "现金流量表": re.compile(r"合并(?:及公司)?现金流量表"),
    }
    
    page_ranges = found_statements.copy()
    if not isinstance(financial_section, dict):
        financial_section = {}
    start_page = financial_section.get("start_page", 1)
    end_page = financial_section.get("end_page", 9999)
    
    # Extract page texts
    page_texts = []
    with pdfplumber.open(pdf_path) as pdf:
        for i in range(start_page - 1, min(end_page, len(pdf.pages))):
            text = pdf.pages[i].extract_text() or ""
            page_texts.append((i + 1, text))
    
    # Search for consolidated statements
    for stmt_name, pattern in consolidated_patterns.items():
        if stmt_name in page_ranges:
            continue
        
        for page_num, text in page_texts:
            # Reduce false positives from audit text / notes by requiring table-like page features.
            if (
                pattern.search(text)
                and page_has_table_features_local(text)
                and has_likely_statement_title_line(text, stmt_name)
            ):
                # Found consolidated statement, estimate end page.
                # Some reports repeat the same title with "(续)" on next page(s),
                # so only stop when a *different* statement starts.
                end_page_est = min(page_num + 5, end_page)
                for p in range(page_num + 1, min(page_num + 10, end_page + 1)):
                    rel_idx = p - start_page
                    if 0 <= rel_idx < len(page_texts):
                        next_text = page_texts[rel_idx][1]
                        for other_name, pat in consolidated_patterns.items():
                            if other_name == stmt_name:
                                continue
                            if pat.search(next_text):
                                end_page_est = p - 1
                                break
                        if end_page_est == p - 1:
                            break
                page_ranges[stmt_name] = (page_num, end_page_est)
                print(f"[INFO] Found consolidated {stmt_name} at page {page_num}")
                break
    
    return page_ranges


def tighten_statement_page_ranges(
    page_ranges: Dict[str, Tuple[int, int]],
    max_pages_per_statement: int = 10,
) -> Dict[str, Tuple[int, int]]:
    """
    Narrow statement blocks using next statement start and a max-page cap.
    This reduces unnecessary table extraction on long financial sections.
    """
    if not page_ranges:
        return page_ranges

    out = page_ranges.copy()
    stmt_order = ["资产负债表", "利润表", "现金流量表"]
    starts = [(name, out[name][0]) for name in stmt_order if name in out]
    starts.sort(key=lambda x: x[1])

    for i, (name, start_page) in enumerate(starts):
        cur_start, cur_end = out[name]
        refined_end = cur_end

        if i + 1 < len(starts):
            next_start = starts[i + 1][1]
            refined_end = min(refined_end, next_start - 1)

        refined_end = min(refined_end, cur_start + max_pages_per_statement - 1)
        refined_end = max(refined_end, cur_start)
        out[name] = (cur_start, refined_end)

    return out


# -----------------------
# Fill 特殊不平/格式/归并 at schema positions (from section computation)
# -----------------------

# Anchor item names per section (for FORMAT: missing many anchors => 特殊格式). Key = parent label (e.g. 流动资产合计).
BS_ANCHOR_NAMES: Dict[str, List[str]] = {
    "流动资产合计": ["货币资金", "应收票据", "应收账款", "存货", "合同资产", "一年内到期的非流动资产", "其他流动资产"],
    "非流动资产合计": ["长期股权投资", "固定资产", "在建工程", "无形资产", "商誉", "长期待摊费用", "递延所得税资产", "其他非流动资产"],
    "资产总计": [],  # use children only
    "流动负债合计": ["短期借款", "应付票据", "应付账款", "合同负债", "一年内到期的非流动负债", "其他流动负债"],
    "非流动负债合计": ["长期借款", "应付债券", "递延所得税负债", "其他非流动负债"],
    "负债合计": [],
    "归属于母公司股东权益合计": ["实收资本", "资本公积", "盈余公积", "未分配利润"],
    "所有者权益(或股东权益)合计": [],
    "负债和所有者权益(或股东权益)总计": [],
}
FORMAT_ANOMALY_MIN_ANCHORS = 3  # if fewer than this many anchors have value, treat as format anomaly

# BS section config: (parent_item_code, children_item_codes, suneven, sformat, smerger)
# Children = item_codes that sum to parent (excluding parent). Derived from schema order.
def _build_bs_section_config(schema_items: List[SchemaItem]) -> List[Tuple[str, str, List[str], str, str, str]]:
    """Returns list of (label, parent_code, children_codes, suneven_code, sformat_code, smerger_code)."""
    codes = [s.item_code for s in schema_items]
    names = [s.item_name_std for s in schema_items]
    config = []
    section_specs = [
        ("流动资产合计", "TOTCURRASSET", "SUNEVENCURRASSE", "SFORMATCURRASSE", "SMERGERCURRASSE", None),
        ("非流动资产合计", "TOTALNONCASSETS", "SUNEVENNONCASSE", "SFORMATNONCASSE", "SMERGERNONCASSE", None),
        ("资产总计", "TOTASSET", "SUNEVENTOTASSET", "SFORMATTOTASSET", "SMERGERTOTASSET", ["TOTCURRASSET", "TOTALNONCASSETS"]),
        ("流动负债合计", "TOTALCURRLIAB", "SUNEVENCURRELIABI", "SFORMATCURRELIABI", "SMERGERCURRELIABI", None),
        ("非流动负债合计", "TOTALNONCLIAB", "SUNEVENNONCLIAB", "SFORMATNONCLIAB", "SMERGERNONCLIAB", None),
        ("负债合计", "TOTLIAB", "SUNEVENTOTLIAB", "SFORMATTOTLIAB", "SMERGERTOTLIAB", ["TOTALCURRLIAB", "TOTALNONCLIAB"]),
        ("归属于母公司股东权益合计", "PARESHARRIGH", "SUNEVENPARESHARRIGH", "SFORMATPARESHARRIGH", "SMERGERPARESHARRIGH", None),
        ("所有者权益(或股东权益)合计", "RIGHAGGR", "SUNEVENRIGHAGGR", "SFORMATRIGHAGGR", "SMERGERRIGHAGGR", None),
        ("负债和所有者权益(或股东权益)总计", "TOTLIABSHAREQUI", "SUNEVENTOTLIABSHAREQUI", "SFORMATTOTLIABSHAREQUI", "SMERGERTOTLIABSHAREQUI", ["TOTLIAB", "RIGHAGGR"]),
    ]
    start = 0
    for spec in section_specs:
        label, parent_code, suneven, sformat, smerger = spec[0], spec[1], spec[2], spec[3], spec[4]
        explicit_children = spec[5] if len(spec) > 5 else None
        try:
            idx = names.index(label)
        except ValueError:
            continue
        if explicit_children is not None:
            children = explicit_children
        else:
            children = [c for i, c in enumerate(codes[start:idx]) if c != parent_code]
        config.append((label, parent_code, children, suneven, sformat, smerger))
        start = idx + 1
    return config


def _compute_imbalance_value(parent_val: Optional[float], children_vals: List[Optional[float]], tolerance_ratio: float = 0.0001) -> Optional[float]:
    """When imbalance (parent != sum(children)), return parent value to put in 特殊不平_* (流动资产的数值)."""
    if parent_val is None:
        return None
    valid = [v for v in children_vals if v is not None]
    if not valid:
        return None
    children_sum = sum(valid)
    diff = parent_val - children_sum
    tol = max(1.0, abs(parent_val) * tolerance_ratio)
    if abs(diff) <= tol:
        return None  # no imbalance
    # 把流动资产的数值 放在 特殊不平_流动资产 -> use parent value (section total)
    return float(parent_val)


def _anchor_codes_for_section(label: str, schema_items: List[SchemaItem]) -> List[str]:
    """Resolve anchor item_name_std to item_codes from schema."""
    names_to_codes = {s.item_name_std: s.item_code for s in schema_items}
    anchor_names = BS_ANCHOR_NAMES.get(label, [])
    return [names_to_codes[n] for n in anchor_names if n in names_to_codes]


def fill_special_anomaly_values_bs(
    df_all: pd.DataFrame,
    schema_items: List[SchemaItem],
    extracted_raw_map: Optional[Dict[str, str]] = None,
) -> None:
    """
    Fill 特殊不平_* / 特殊格式_* / 特殊归并_* at schema positions.
    - 特殊不平: section value when parent != sum(children).
    - 特殊格式: section value when section has too few anchor items (format anomaly).
    - 特殊归并: section value when any row in section matches merged-item pattern (e.g. 应收票据及应收账款).
    extracted_raw_map: optional item_code -> raw_item from extraction (for AGGREGATION).
    """
    val_map = df_all.set_index("item_code")["value"].to_dict()
    config = _build_bs_section_config(schema_items)
    updates = {}  # item_code -> value
    for label, parent_code, children_codes, suneven_code, sformat_code, smerger_code in config:
        parent_val = val_map.get(parent_code)
        if parent_val is not None and not pd.isna(parent_val):
            parent_val_f = float(parent_val)
            children_vals = [val_map.get(c) for c in children_codes]
            # 特殊不平: put section value when imbalance detected
            v = _compute_imbalance_value(parent_val, children_vals)
            if v is not None:
                updates[suneven_code] = v
        else:
            parent_val_f = None
        # 特殊格式: section has too few anchor items with value
        anchor_codes = _anchor_codes_for_section(label, schema_items)
        if anchor_codes and parent_val_f is not None:
            present = sum(1 for c in anchor_codes if val_map.get(c) is not None and not pd.isna(val_map.get(c)))
            if present < FORMAT_ANOMALY_MIN_ANCHORS:
                updates[sformat_code] = parent_val_f
        # 特殊归并: any extracted row in this section has merge pattern (known merged name or 及/和/与 in raw)
        if extracted_raw_map and parent_val_f is not None:
            section_codes = set(children_codes) | {parent_code}
            for code in section_codes:
                raw = extracted_raw_map.get(code)
                if not raw:
                    continue
                is_agg, _ = detect_aggregation_anomaly(raw, [code])
                if not is_agg and re.search(r"[及和与]", raw):
                    is_agg = True
                if is_agg:
                    updates[smerger_code] = parent_val_f
                    break
    for code, value in updates.items():
        df_all.loc[df_all["item_code"] == code, "value"] = value


# -----------------------
# Extract current year data only
# -----------------------

def extract_current_year_data(
    tables: List[pd.DataFrame],
    statement_type: str,
    schema_items: List[SchemaItem],
) -> pd.DataFrame:
    """
    Extract data from tables, only keep current year (本年/期末) data.
    Returns DataFrame with columns: item_code, item_name, value
    Deduplicates by item_code (keeps first occurrence).
    """
    match_index = build_match_index(schema_items)
    
    extracted_by_code: Dict[str, Dict[str, Any]] = {}
    
    for t in tables:
        if t is None or t.empty:
            continue
        
        t = t.copy()
        # Ensure unique column labels to avoid ambiguous selection.
        raw_cols = [str(c).strip() for c in t.columns]
        seen = {}
        uniq_cols = []
        for c in raw_cols:
            k = c if c else "col"
            n = seen.get(k, 0)
            uniq_cols.append(k if n == 0 else f"{k}_{n}")
            seen[k] = n + 1
        t.columns = uniq_cols
        
        # Find item column
        item_col = None
        for col in t.columns:
            col_lower = str(col).lower()
            if "项目" in col or "科目" in col or col == "col_0":
                item_col = col
                break
        
        if item_col is None and len(t.columns) > 0:
            item_col = t.columns[0]
        
        if item_col is None:
            continue

        # For CF, ignore unrelated tables (e.g., equity movement statement captured in nearby pages).
        if statement_type == "CF" and not is_likely_cf_table(t, item_col):
            continue
        
        # Choose best value column by statement semantics.
        if statement_type == "BS":
            value_col = choose_bs_period_end_col(t, item_col)
        else:
            value_col = choose_best_value_col(t, item_col, statement_type=statement_type)
        if value_col is None:
            continue
        fallback_cols = [c for c in t.columns if c not in {item_col, value_col}]
        current_first_cols = _order_cols_for_current_period(
            t,
            item_col=item_col,
            cols=[value_col] + fallback_cols,
        )

        # CN BS strict TOTAL_ASSETS extraction by row/column rule.
        if statement_type == "BS":
            v_total_assets, raw_item_total_assets = _extract_bs_total_assets_from_table(t, item_col)
            if v_total_assets is not None:
                _merge_extracted_record(
                    extracted_by_code,
                    {
                        "item_code": "TOTASSET",
                        "item_name": "资产总计",
                        "value": float(v_total_assets),
                        "raw_item": raw_item_total_assets or "资产总计",
                        "raw_text": raw_item_total_assets or "资产总计",
                        "status": "OK",
                        "fixup_reason": None,
                    },
                )
        
        for _, r in t.iterrows():
            raw_item = str(r.get(item_col, "")).strip()
            if not raw_item or len(raw_item) < 2:
                continue
            
            # Skip anomaly marker rows (特殊不平_流动资产 etc. — no numeric values, flag only)
            row_values = [r.get(c) for c in t.columns if c != item_col]
            if is_anomaly_marker_row(raw_item, row_values):
                continue
            
            # Skip category headers only (not 合计 — we need 流动资产合计 etc. for section totals and 特殊不平)
            cleaned = clean_item_name(raw_item)
            if cleaned in ["项目", "资产", "负债", "股东权益", "流动资产", "非流动资产", "流动负债", "非流动负债"]:
                continue
            
            item_code, score, matched_name = match_item(raw_item, schema_items, match_index)
            if not item_code:
                continue
            
            old_rec = extracted_by_code.get(item_code)
            # Keep scanning unless we already have a strong OK for this item.
            if old_rec is not None and str(old_rec.get("status") or "") == "OK":
                continue

            rec_status = "MISSING"
            rec_value: Optional[float] = None
            fixup_reason: Optional[str] = None
            candidate_texts: List[str] = []

            if item_code in {"BASICEPS", "DILUTEDEPS"}:
                # EPS: preserve decimal semantics and prioritize current-period column.
                chosen_col: Optional[str] = None
                for c in current_first_cols:
                    raw_cell = r.get(c)
                    cell_text = str(raw_cell).strip() if raw_cell is not None else ""
                    if cell_text:
                        candidate_texts.append(cell_text)
                    v_alt = parse_eps_number(raw_cell)
                    if v_alt is not None:
                        rec_value = float(v_alt)
                        rec_status = "OK"
                        chosen_col = c
                        break
                if rec_status != "OK":
                    row_blob = raw_item + " " + " ".join(candidate_texts)
                    if item_code in ALLOW_NOT_APPLICABLE_CODES and _is_not_applicable_text(row_blob):
                        rec_status = "NOT_APPLICABLE"
                    elif candidate_texts:
                        rec_status = "PARSE_ERROR"
                    else:
                        continue
                elif chosen_col is not None and current_first_cols and chosen_col != current_first_cols[0]:
                    fixup_reason = "COLUMN_ROLE_SWAP"
            else:
                v = parse_number(r.get(value_col))
                if v is not None:
                    rec_value = float(v)
                    rec_status = "OK"
                else:
                    for c in fallback_cols:
                        raw_cell = r.get(c)
                        cell_text = str(raw_cell).strip() if raw_cell is not None else ""
                        if cell_text:
                            candidate_texts.append(cell_text)
                        v_alt = parse_number(raw_cell)
                        if v_alt is not None:
                            rec_value = float(v_alt)
                            rec_status = "OK"
                            break
                if rec_status != "OK":
                    row_blob = raw_item + " " + " ".join(candidate_texts)
                    if item_code in ALLOW_NOT_APPLICABLE_CODES and _is_not_applicable_text(row_blob):
                        rec_status = "NOT_APPLICABLE"
                    elif candidate_texts:
                        rec_status = "PARSE_ERROR"
                    else:
                        continue

            _merge_extracted_record(
                extracted_by_code,
                {
                    "item_code": item_code,
                    "item_name": matched_name if matched_name else raw_item,
                    "value": rec_value,
                    "raw_item": raw_item,
                    "raw_text": raw_item,
                    "status": rec_status,
                    "fixup_reason": fixup_reason,
                },
            )

    return pd.DataFrame(list(extracted_by_code.values()))


# -----------------------
# Main pipeline
# -----------------------

def run_pipeline(
    pdf_path: str,
    schema_file: str,
    out_path: str,
) -> None:
    """
    Main pipeline: extract consolidated statements and output to Excel.
    """
    # Load schemas
    schemas = load_schemas(schema_file)
    
    # Locate statements
    print("\n[INFO] Locating statements...")
    loc_result = locate_statements(pdf_path)
    
    # Find consolidated statements
    page_ranges = {}
    for stmt in loc_result["statements"]:
        if stmt["start_page"] and stmt["end_page"]:
            name = stmt["name"]
            matched_titles = " ".join(stmt.get("matched_titles", []))
            # Only keep consolidated statements
            if "合并" in matched_titles:
                page_ranges[name] = (stmt["start_page"], stmt["end_page"])
            elif "母公司" in matched_titles:
                print(f"[INFO] Skipping parent company statement: {name}")
    
    # Search for consolidated if not found
    if len(page_ranges) < 3:
        print("[INFO] Searching for consolidated statements...")
        financial_section = loc_result.get("financial_section") or {}
        page_ranges = search_consolidated_statements(pdf_path, financial_section, page_ranges)

    page_ranges = tighten_statement_page_ranges(page_ranges)
    print(f"[INFO] Page ranges: {page_ranges}")
    
    # Map statement names
    stmt_map = {
        "资产负债表": "BS",
        "利润表": "IS",
        "现金流量表": "CF",
    }
    
    # Extract data for each statement
    results = {}
    
    for stmt_name, (ps, pe) in page_ranges.items():
        stmt_type = stmt_map.get(stmt_name, "")
        if not stmt_type or stmt_type not in schemas:
            print(f"[WARN] Skipping {stmt_name}: no schema found")
            continue
        
        print(f"\n[INFO] Processing {stmt_name} (pages {ps}-{pe})...")
        
        # Extract with lattice first.
        tables_l = extract_statement_tables(pdf_path, ps, pe, flavor="lattice")
        df_l = extract_current_year_data(tables_l, stmt_type, schemas[stmt_type]) if tables_l else pd.DataFrame()

        # Stream second pass: always run and compare, because some CN reports are
        # parsed much better by stream (or vice versa) depending on layout.
        need_stream = True
        df_s = pd.DataFrame()
        if need_stream:
            print(f"[INFO] Trying stream extraction for {stmt_name}...")
            tables_s = extract_statement_tables(pdf_path, ps, pe, flavor="stream")
            if tables_s:
                df_s = extract_current_year_data(tables_s, stmt_type, schemas[stmt_type])

        # Choose denser result; for BS, merge stream+lattice to recover complementary items.
        df = df_l
        source = "lattice"
        if not df_s.empty and len(df_s) >= len(df):
            df = df_s
            source = "stream"
        if stmt_type == "BS" and not df_l.empty and not df_s.empty:
            merged = pd.concat([df_s, df_l], ignore_index=True)
            merged = merged.drop_duplicates(subset=["item_code"], keep="first")
            if len(merged) > len(df):
                df = merged
                source = "stream+lattice"
        
        if not df.empty:
            print(f"[INFO] Extracted {len(df)} items for {stmt_name} ({source})")
            results[stmt_name] = df
        else:
            print(f"[WARN] No data extracted for {stmt_name}")
    
    # Write to Excel with all schema items (fixed structure)
    with pd.ExcelWriter(out_path, engine='openpyxl') as writer:
        # Map statement names to sheet names
        stmt_to_sheet = {
            "资产负债表": "CN_FIN_BS_GEN",
            "利润表": "CN_FIN_IS_GEN",
            "现金流量表": "CN_FIN_CF_GEN",
        }
        
        for stmt_name in ["资产负债表", "利润表", "现金流量表"]:
            stmt_type = stmt_map.get(stmt_name, "")
            sheet_name = stmt_to_sheet.get(stmt_name, stmt_name)
            
            if stmt_type not in schemas:
                # Create empty sheet with columns only
                df_empty = pd.DataFrame(
                    columns=[
                        "item_code", "item_name", "value", "status",
                        "raw_text", "fixup_reason", "llm_repair_eligible",
                    ]
                )
                df_empty.to_excel(writer, sheet_name=sheet_name, index=False)
                print(f"[WARN] No schema for {stmt_name}, created empty sheet '{sheet_name}'")
                continue
            
            # Get all schema items for this statement type (keep original order)
            schema_items = schemas[stmt_type]
            
            # Create DataFrame with all schema items in schema order
            all_items = []
            for schema_item in schema_items:
                all_items.append({
                    'item_code': schema_item.item_code,
                    'item_name': schema_item.item_name_std,
                    'value': None,  # Default to None (empty)
                    'status': 'MISSING',
                    'raw_text': None,
                    'fixup_reason': None,
                })
            
            df_all = pd.DataFrame(all_items)
            
            # Merge with extracted data
            extracted_raw_map = None
            if stmt_name in results:
                df_extracted = results[stmt_name]
                # Create a lookup dict: item_code -> value
                extracted_dict = dict(zip(df_extracted['item_code'], df_extracted['value']))
                status_series = df_extracted["status"] if "status" in df_extracted.columns else pd.Series(["OK"] * len(df_extracted))
                raw_text_series = df_extracted["raw_text"] if "raw_text" in df_extracted.columns else pd.Series([None] * len(df_extracted))
                fixup_series = df_extracted["fixup_reason"] if "fixup_reason" in df_extracted.columns else pd.Series([None] * len(df_extracted))
                extracted_status = dict(zip(df_extracted["item_code"], status_series))
                extracted_raw_text = dict(zip(df_extracted["item_code"], raw_text_series))
                extracted_fixup = dict(zip(df_extracted["item_code"], fixup_series))
                # Update values where we have extracted data (map preserves order)
                df_all['value'] = df_all['item_code'].map(extracted_dict)
                df_all['status'] = df_all['item_code'].map(extracted_status).fillna('MISSING')
                df_all['raw_text'] = df_all['item_code'].map(extracted_raw_text)
                df_all['fixup_reason'] = df_all['item_code'].map(extracted_fixup)
                matched_count = (df_all['status'] != 'MISSING').sum()
                print(f"[INFO] Matched {matched_count} out of {len(df_all)} schema items for {stmt_name}")
                if stmt_type == "BS" and "raw_item" in df_extracted.columns:
                    extracted_raw_map = dict(zip(df_extracted["item_code"], df_extracted["raw_item"]))
            else:
                print(f"[WARN] No extracted data for {stmt_name}, all values will be empty")
            
            # Fill 特殊不平_* / 特殊格式_* / 特殊归并_* at schema positions
            if stmt_type == "BS":
                fill_special_anomaly_values_bs(df_all, schema_items, extracted_raw_map=extracted_raw_map)
            # Any value filled by post-processing should be treated as OK unless already stronger.
            mask_ok = df_all["value"].notna() & (df_all["status"] == "MISSING")
            df_all.loc[mask_ok, "status"] = "OK"

            # LLM repair layer only targets core fields currently marked as missing/parse_error.
            llm_core_codes = LLM_CORE_CODES_BY_STMT.get(stmt_type, set())
            df_all["llm_repair_eligible"] = (
                df_all["item_code"].isin(llm_core_codes)
                & df_all["status"].isin(["MISSING", "PARSE_ERROR"])
            ).astype(int)
            
            # Ensure order is preserved: reset index explicitly and write
            df_all = df_all.reset_index(drop=True)  # Ensure clean index
            # Keep original schema order (no sorting, no reordering)
            df_all.to_excel(writer, sheet_name=sheet_name, index=False)
            print(f"[OK] Wrote {len(df_all)} rows to sheet '{sheet_name}'")
    
    print(f"\n[OK] Wrote Excel file: {out_path}")


def main():
    import os
    import re
    
    parser = argparse.ArgumentParser(description="Extract consolidated financial statements (current year only)")
    parser.add_argument("--pdf", required=True, help="Path to annual report PDF")
    parser.add_argument("--schema-file", default="schemas/CN_Schemas.xlsx", help="Path to schema Excel file")
    parser.add_argument("--out", default=None, help="Output Excel path (default: auto-generated from PDF filename)")
    args = parser.parse_args()
    
    # Auto-generate output filename from PDF if not provided
    if args.out is None:
        pdf_basename = os.path.basename(args.pdf)
        # Extract stock code from filename (e.g., "002415_Ch.pdf" -> "002415")
        match = re.match(r"^(\d{6})", pdf_basename)
        if match:
            stock_code = match.group(1)
            out_path = f"{stock_code}.xlsx"
        else:
            # Fallback: use PDF basename without extension
            out_path = os.path.splitext(pdf_basename)[0] + ".xlsx"
        print(f"[INFO] Auto-generated output filename: {out_path}")
    else:
        out_path = args.out
    
    run_pipeline(
        pdf_path=args.pdf,
        schema_file=args.schema_file,
        out_path=out_path,
    )


if __name__ == "__main__":
    main()
