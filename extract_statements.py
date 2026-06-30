#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
FinReporting (CN) - Annual report PDF -> TOC -> Financial section -> 3 statements page blocks

Usage:
  python extract_statements.py --pdf path/to/report.pdf --out out.json
"""

from __future__ import annotations

import argparse
import json
import re
from dataclasses import dataclass, asdict
from typing import Dict, List, Optional, Tuple

import pdfplumber


# -----------------------------
# Data structures
# -----------------------------

@dataclass
class TocItem:
    title: str
    page: int  # 1-indexed


@dataclass
class SectionRange:
    start_page: int
    end_page: int


@dataclass
class StatementLoc:
    name: str
    start_page: Optional[int]
    end_page: Optional[int]
    confidence: float
    matched_titles: List[str]


# -----------------------------
# Regex / Heuristics
# -----------------------------

# Common TOC line patterns (very tolerant)
# Examples:
# "第十节 财务报告 ........................................ 152"
# "第八节    财务报告 ………………………………………… 111"
# "第十节 财务报告 .......................... - 148 -"
# Also covers "致股东 ... 1" etc.
TOC_LINE_RE = re.compile(
    r"""
    ^\s*
    (?P<title>.+?)                       # title (lazy)
    \s*
    (?:\.{2,}|…{2,}|·{2,}|_{2,}|-{2,})?  # dotted leader variants
    \s*
    (?P<page>\d{1,4})                    # page number
    \s*
    (?:-\s*\d{1,4}\s*-\s*)?              # optional "- 148 -" style
    \s*$
    """,
    re.VERBOSE,
)

# More targeted for "第X节 ..." lines, helps to get next section boundary
SECTION_HEADER_RE = re.compile(r"^\s*第[一二三四五六七八九十百0-9]+节\s+.+$")

# TOC page detection: must contain "目录" and have leader+page patterns
LEADER_HINT_RE = re.compile(r"(\.{3,}|…{3,})\s*\d{1,4}")

# Statement title patterns
STATEMENT_PATTERNS: Dict[str, re.Pattern] = {
    "balance_sheet": re.compile(r"(合并及公司|合并|母公司|公司)?资产负债表"),
    "income_statement": re.compile(r"(合并及公司|合并|母公司|公司)?利润表"),
    "cash_flow_statement": re.compile(r"(合并及公司|合并|母公司|公司)?现金流量表"),
}
CONSOLIDATED_STATEMENT_PATTERNS: Dict[str, re.Pattern] = {
    "balance_sheet": re.compile(r"合并(?:及公司)?资产负债表"),
    "income_statement": re.compile(r"合并(?:及公司)?利润表"),
    "cash_flow_statement": re.compile(r"合并(?:及公司)?现金流量表"),
}

# Table-ish features that indicate we are still inside a statement table
TABLE_FEATURE_RE = re.compile(r"(项目|本期金额|上期金额|期末余额|期初余额)")
# Another heuristic: numeric density (many digits)
DIGIT_RE = re.compile(r"\d")

# Fast-scan config:
# 1) first pass: only scan first N pages to find TOC
# 2) second pass: extend TOC scan range if first pass fails
# 3) fallback: full scan only if still cannot localize financial section
TOC_INITIAL_SCAN_PAGES = 40
TOC_EXTENDED_SCAN_PAGES = 120


# -----------------------------
# PDF text extraction
# -----------------------------

def normalize_page_text(text: str) -> str:
    """Normalize page text for downstream regex rules."""
    text = text or ""
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n{3,}", "\n\n", text).strip()
    return text


def fill_page_texts_from_pdf(
    pdf,
    page_texts: List[str],
    page_numbers: List[int],
    extracted_pages: set,
) -> Tuple[int, int]:
    """
    Extract selected pages into page_texts in-place.
    Returns (newly_extracted_pages_count, newly_empty_pages_count).
    """
    total_pages = len(pdf.pages)
    added = 0
    empty = 0
    for p in sorted(set(page_numbers)):
        if p < 1 or p > total_pages:
            continue
        if p in extracted_pages:
            continue
        text = normalize_page_text(pdf.pages[p - 1].extract_text() or "")
        page_texts[p - 1] = text
        extracted_pages.add(p)
        added += 1
        if len(text) < 10:
            empty += 1
    return added, empty


def extract_page_texts(pdf_path: str) -> Tuple[List[str], int, bool]:
    """
    Returns:
      page_texts: list of per-page text (1-indexed concept, but list is 0-indexed)
      total_pages: int
      looks_scanned: bool (rough heuristic: too many empty pages)
    """
    page_texts: List[str] = []
    empty = 0

    with pdfplumber.open(pdf_path) as pdf:
        total_pages = len(pdf.pages)
        for i, page in enumerate(pdf.pages):
            text = normalize_page_text(page.extract_text() or "")
            page_texts.append(text)
            if len(text) < 10:
                empty += 1

    # If a large proportion of pages have no text, it might be scanned
    looks_scanned = (empty / max(total_pages, 1)) > 0.6
    return page_texts, total_pages, looks_scanned


# -----------------------------
# TOC detection & parsing
# -----------------------------

def is_toc_page(text: str) -> bool:
    if not text:
        return False
    if "目录" not in text:
        return False

    # at least a few leader+page hints
    hits = len(LEADER_HINT_RE.findall(text))
    # sometimes leader symbols are not extracted; also check many short lines ending with digits
    digit_line_hits = sum(1 for line in text.splitlines() if re.search(r"\d{1,4}\s*$", line))
    return hits >= 2 or digit_line_hits >= 6


def find_toc_pages(page_texts: List[str]) -> List[int]:
    toc_pages: List[int] = []
    for idx, text in enumerate(page_texts, start=1):
        if is_toc_page(text):
            toc_pages.append(idx)

    # Merge consecutive TOC pages (keep all, but we will parse them together)
    return toc_pages


def parse_toc_items(toc_text: str) -> List[TocItem]:
    items: List[TocItem] = []
    for raw_line in toc_text.splitlines():
        line = raw_line.strip()
        if not line:
            continue
        # Skip "目录" title line itself
        if line == "目录":
            continue

        m = TOC_LINE_RE.match(line)
        if not m:
            continue

        title = m.group("title").strip()
        page_str = m.group("page").strip()

        # Basic cleaning: collapse multiple spaces
        title = re.sub(r"\s{2,}", " ", title)

        # Guard: title too short or too numeric
        if len(title) < 2:
            continue
        if title.isdigit():
            continue

        try:
            page = int(page_str)
        except ValueError:
            continue

        # Page number sanity
        if page <= 0 or page > 9999:
            continue

        items.append(TocItem(title=title, page=page))

    # De-duplicate by (title,page)
    uniq: Dict[Tuple[str, int], TocItem] = {}
    for it in items:
        uniq[(it.title, it.page)] = it
    return list(uniq.values())


def get_financial_section(items: List[TocItem], total_pages: int) -> Optional[SectionRange]:
    """
    Find "财务报告" start page and infer end page from next section-like item.
    """
    if not items:
        return None

    # Sort by page
    items_sorted = sorted(items, key=lambda x: x.page)

    # Find financial report item
    idx_fin = None
    for i, it in enumerate(items_sorted):
        if "财务报告" in it.title:
            idx_fin = i
            break
    if idx_fin is None:
        return None

    start_page = items_sorted[idx_fin].page

    # Find next "第X节" item after financial report
    end_page = total_pages
    for j in range(idx_fin + 1, len(items_sorted)):
        # Prefer next section headers (common in CN reports)
        if SECTION_HEADER_RE.match(items_sorted[j].title):
            end_page = max(items_sorted[j].page - 1, start_page)
            break
        # If no section header, we can still use next item with a larger page gap
        # but keep it conservative: if it's far enough ahead, treat it as boundary
        if items_sorted[j].page > start_page + 5:
            end_page = max(items_sorted[j].page - 1, start_page)
            break

    # Clamp
    start_page = max(1, min(start_page, total_pages))
    end_page = max(start_page, min(end_page, total_pages))
    return SectionRange(start_page=start_page, end_page=end_page)


# -----------------------------
# Statement localization
# -----------------------------

def page_has_table_features(text: str) -> bool:
    if not text:
        return False
    # Avoid treating audit opinion pages as table pages.
    if "审计报告" in text and "审字" in text:
        return False
    if TABLE_FEATURE_RE.search(text):
        return True

    # numeric density heuristic
    digits = len(DIGIT_RE.findall(text))
    # Rough: if many digits relative to length, likely table
    if len(text) > 0 and (digits / max(len(text), 1)) > 0.08 and digits > 40:
        return True
    return False


def find_statement_start_pages(
    page_texts: List[str],
    section: SectionRange,
) -> Dict[str, Tuple[int, List[str]]]:
    """
    Returns mapping: statement_key -> (start_page, matched_titles)
    """
    starts: Dict[str, Tuple[int, List[str]]] = {}
    is_consolidated_map: Dict[str, bool] = {}

    # scan within financial section
    for p in range(section.start_page, section.end_page + 1):
        text = page_texts[p - 1]  # 0-indexed list
        if not text:
            continue

        # If multiple statement titles appear in one non-table-like page,
        # it's likely a directory/index page, skip it.
        page_hits = []
        for k, pat in STATEMENT_PATTERNS.items():
            if pat.search(text):
                page_hits.append(k)
        if len(page_hits) >= 2 and not page_has_table_features(text):
            continue

        for key, pat in STATEMENT_PATTERNS.items():
            # Prefer consolidated titles first.
            m_cons = CONSOLIDATED_STATEMENT_PATTERNS[key].search(text)
            m = m_cons if m_cons else pat.search(text)
            if not m:
                continue

            matched = m.group(0)
            cur_is_cons = m_cons is not None

            if key not in starts:
                starts[key] = (p, [matched])
                is_consolidated_map[key] = cur_is_cons
                continue

            prev_is_cons = is_consolidated_map.get(key, False)
            prev_page = starts[key][0]
            # Upgrade from parent/company title to consolidated title when found.
            if (not prev_is_cons and cur_is_cons) or (
                prev_is_cons == cur_is_cons and p < prev_page
            ):
                starts[key] = (p, [matched])
                is_consolidated_map[key] = cur_is_cons

    return starts


def expand_statement_block(
    page_texts: List[str],
    start_page: int,
    max_end: int,
) -> int:
    """
    Expand until two consecutive pages lack table features.
    """
    end_page = start_page
    misses = 0
    for p in range(start_page, max_end + 1):
        text = page_texts[p - 1]
        if page_has_table_features(text):
            end_page = p
            misses = 0
        else:
            misses += 1
            # allow one blank-ish page inside, but stop after 2 consecutive misses
            if misses >= 2 and p > start_page:
                break
    return end_page


def make_confidence(
    start_page: Optional[int],
    end_page: Optional[int],
    matched_titles: List[str],
    page_texts: List[str],
) -> float:
    if start_page is None or end_page is None:
        return 0.0

    conf = 0.0
    if matched_titles:
        conf += 0.6

    # add up to 0.4 based on how many pages look table-like in block (cap 4 pages * 0.1)
    table_pages = 0
    for p in range(start_page, end_page + 1):
        if page_has_table_features(page_texts[p - 1]):
            table_pages += 1
    conf += min(0.4, 0.1 * table_pages)
    return round(min(conf, 1.0), 3)


def locate_statements(
    page_texts: List[str],
    section: SectionRange,
) -> List[StatementLoc]:
    starts = find_statement_start_pages(page_texts, section)

    results: List[StatementLoc] = []
    for key, human_name in [
        ("balance_sheet", "资产负债表"),
        ("income_statement", "利润表"),
        ("cash_flow_statement", "现金流量表"),
    ]:
        if key not in starts:
            results.append(
                StatementLoc(
                    name=human_name,
                    start_page=None,
                    end_page=None,
                    confidence=0.0,
                    matched_titles=[],
                )
            )
            continue

        start_page, matched_titles = starts[key]
        end_page = expand_statement_block(page_texts, start_page, section.end_page)
        conf = make_confidence(start_page, end_page, matched_titles, page_texts)

        results.append(
            StatementLoc(
                name=human_name,
                start_page=start_page,
                end_page=end_page,
                confidence=conf,
                matched_titles=matched_titles,
            )
        )

    return results


# -----------------------------
# Main
# -----------------------------

def run(pdf_path: str) -> Dict:
    # Fast path: avoid reading full PDF text up front.
    with pdfplumber.open(pdf_path) as pdf:
        total_pages = len(pdf.pages)
        page_texts = [""] * total_pages
        extracted_pages = set()
        empty_pages = 0

        # Pass 1: scan first pages to locate TOC quickly.
        p1_end = min(total_pages, TOC_INITIAL_SCAN_PAGES)
        _, p1_empty = fill_page_texts_from_pdf(
            pdf=pdf,
            page_texts=page_texts,
            page_numbers=list(range(1, p1_end + 1)),
            extracted_pages=extracted_pages,
        )
        empty_pages += p1_empty

        toc_pages = find_toc_pages(page_texts)
        toc_text = "\n".join(page_texts[p - 1] for p in toc_pages) if toc_pages else ""
        toc_items = parse_toc_items(toc_text) if toc_text else []
        section = get_financial_section(toc_items, total_pages)

        # Pass 2: extend TOC scan range if needed.
        if section is None:
            p2_end = min(total_pages, TOC_EXTENDED_SCAN_PAGES)
            if p2_end > p1_end:
                _, p2_empty = fill_page_texts_from_pdf(
                    pdf=pdf,
                    page_texts=page_texts,
                    page_numbers=list(range(p1_end + 1, p2_end + 1)),
                    extracted_pages=extracted_pages,
                )
                empty_pages += p2_empty
                toc_pages = find_toc_pages(page_texts)
                toc_text = "\n".join(page_texts[p - 1] for p in toc_pages) if toc_pages else ""
                toc_items = parse_toc_items(toc_text) if toc_text else []
                section = get_financial_section(toc_items, total_pages)

        # Fallback: full scan only when TOC/section still missing.
        if section is None and len(extracted_pages) < total_pages:
            _, p3_empty = fill_page_texts_from_pdf(
                pdf=pdf,
                page_texts=page_texts,
                page_numbers=list(range(1, total_pages + 1)),
                extracted_pages=extracted_pages,
            )
            empty_pages += p3_empty
            toc_pages = find_toc_pages(page_texts)
            toc_text = "\n".join(page_texts[p - 1] for p in toc_pages) if toc_pages else ""
            toc_items = parse_toc_items(toc_text) if toc_text else []
            section = get_financial_section(toc_items, total_pages)

        statements: List[StatementLoc] = []
        if section:
            # Only now extract financial section pages (if still missing).
            _, sec_empty = fill_page_texts_from_pdf(
                pdf=pdf,
                page_texts=page_texts,
                page_numbers=list(range(section.start_page, section.end_page + 1)),
                extracted_pages=extracted_pages,
            )
            empty_pages += sec_empty
            statements = locate_statements(page_texts, section)

    # Scanned-like heuristic over actually extracted pages.
    looks_scanned = (empty_pages / max(len(extracted_pages), 1)) > 0.6

    out = {
        "pdf_path": pdf_path,
        "total_pages": total_pages,
        "extracted_pages_count": len(extracted_pages),
        "looks_scanned": looks_scanned,
        "toc_pages": toc_pages,
        "toc_items": [asdict(x) for x in sorted(toc_items, key=lambda t: t.page)],
        "financial_section": asdict(section) if section else None,
        "statements": [asdict(s) for s in statements],
        "notes": {
            "next_step_if_scanned": (
                "PDF看起来像扫描件(文本提取为空较多)。需要OCR：用pdf2image把指定页转图片，再tesseract/ppocr抽文本后复用本逻辑。"
                if looks_scanned
                else "PDF看起来是文本型，可直接进行表格抽取。"
            )
        },
    }
    return out


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--pdf", required=True, help="Path to annual report PDF")
    parser.add_argument("--out", required=False, help="Output JSON path (optional)")
    args = parser.parse_args()

    result = run(args.pdf)

    print("[INFO] total_pages:", result["total_pages"])
    print("[INFO] looks_scanned:", result["looks_scanned"])
    print("[INFO] toc_pages:", result["toc_pages"])
    print("[INFO] financial_section:", result["financial_section"])
    for s in result["statements"]:
        print(f"[INFO] {s['name']}: {s['start_page']} - {s['end_page']} (conf={s['confidence']})")

    if args.out:
        with open(args.out, "w", encoding="utf-8") as f:
            json.dump(result, f, ensure_ascii=False, indent=2)
        print("[OK] wrote:", args.out)
    else:
        print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
