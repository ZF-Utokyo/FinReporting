#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Balance Sheet extraction with anomaly detection.

Handles 3 anomaly types per section (e.g. 流动资产):
- 特殊不平_* = IMBALANCE (parent != sum(children))
- 特殊格式_* = FORMAT (table structure differs from template)
- 特殊归并_* = AGGREGATION (merged disclosure lines)

Anomaly marker rows (text "特殊xxx_*" with no numeric values) are treated as
flags, not financial items — detected by simple exact/normalized string match.
"""

from __future__ import annotations

import json
import os
import re
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional, Tuple

# ---------------------------------------------------------------------------
# 0) Data model
# ---------------------------------------------------------------------------

@dataclass
class TableItem:
    doc_id: str
    statement_type: str  # consolidated / parent
    period_end: str  # YYYY-MM-DD
    section: str  # e.g. current_assets, noncurrent_assets
    item_raw: str
    item_std: str  # canonical code
    value: Optional[float] = None
    value_prev: Optional[float] = None
    unit: str = "CNY"
    note_ref: Optional[str] = None
    confidence: float = 1.0
    anomaly_flags: List[str] = field(default_factory=list)
    anomaly_details: Dict[str, Any] = field(default_factory=dict)
    row_type: str = "ITEM"  # ITEM | ANOMALY_MARKER


@dataclass
class TableSection:
    doc_id: str
    statement_type: str
    period_end: str
    section: str
    parent_item_std: str
    parent_value: Optional[float] = None
    children_sum: Optional[float] = None
    diff: Optional[float] = None
    anomaly_flags: List[str] = field(default_factory=list)
    confidence: float = 1.0
    anomaly_details: Dict[str, Any] = field(default_factory=dict)


# ---------------------------------------------------------------------------
# Anomaly marker detection (no numeric values → flag rows only)
# ---------------------------------------------------------------------------

# Exact strings for anomaly marker rows (simple match)
ANOMALY_MARKER_PREFIX = "特殊"
ANOMALY_MARKER_PATTERNS = [
    "特殊不平_流动资产",
    "特殊不平_非流动资产",
    "特殊格式_流动资产",
    "特殊格式_非流动资产",
    "特殊归并_流动资产",
    "特殊归并_非流动资产",
]

# Map marker text → section + flag type
ANOMALY_MARKER_TO_FLAG: Dict[str, Tuple[str, str]] = {
    "特殊不平_流动资产": ("current_assets", "IMBALANCE"),
    "特殊不平_非流动资产": ("noncurrent_assets", "IMBALANCE"),
    "特殊格式_流动资产": ("current_assets", "FORMAT"),
    "特殊格式_非流动资产": ("noncurrent_assets", "FORMAT"),
    "特殊归并_流动资产": ("current_assets", "AGGREGATION"),
    "特殊归并_非流动资产": ("noncurrent_assets", "AGGREGATION"),
}


def normalize_for_anomaly_match(text: str) -> str:
    """Normalize text for exact anomaly marker match."""
    if not text:
        return ""
    s = str(text).strip()
    s = re.sub(r"\s+", "", s)
    return s


def is_anomaly_marker_row(item_raw: str, numeric_values: List[Any]) -> bool:
    """
    If row text is exactly one of 特殊不平_流动资产 / 特殊格式_流动资产 / 特殊归并_流动资产
    (or 非流动资产 variants) AND all numeric cells are empty
    => treat as ANOMALY_MARKER, not ITEM. Simple exact/normalized string match.
    """
    if not item_raw or ANOMALY_MARKER_PREFIX not in item_raw:
        return False
    key = normalize_for_anomaly_match(item_raw)
    if not key:
        return False
    # Exact match: must match one of the known marker strings exactly (after normalize)
    matched = False
    for pattern in ANOMALY_MARKER_PATTERNS:
        if key == normalize_for_anomaly_match(pattern):
            matched = True
            break
    if not matched:
        return False
    # No numeric values => anomaly marker row (these rows have no numbers)
    has_any_number = False
    for v in numeric_values:
        if v is None:
            continue
        s = str(v).strip().replace(",", "").replace("，", "")
        if re.search(r"[\d.]", s) and s not in {"—", "-", "–", ""}:
            has_any_number = True
            break
    return not has_any_number


def anomaly_marker_to_section_flag(item_raw: str) -> Optional[Tuple[str, str]]:
    """
    Map "特殊不平_流动资产" -> (section, flag).
    Returns None if not a known marker.
    """
    key = normalize_for_anomaly_match(item_raw)
    for pattern, (section, flag) in ANOMALY_MARKER_TO_FLAG.items():
        if key == normalize_for_anomaly_match(pattern) or pattern in key:
            return (section, flag)
    if key.startswith(ANOMALY_MARKER_PREFIX):
        # Generic: 特殊xxx_流动资产 -> current_assets
        if "流动资产" in key:
            if "不平" in key:
                return ("current_assets", "IMBALANCE")
            if "格式" in key:
                return ("current_assets", "FORMAT")
            if "归并" in key:
                return ("current_assets", "AGGREGATION")
        if "非流动资产" in key:
            if "不平" in key:
                return ("noncurrent_assets", "IMBALANCE")
            if "格式" in key:
                return ("noncurrent_assets", "FORMAT")
            if "归并" in key:
                return ("noncurrent_assets", "AGGREGATION")
    return None


# ---------------------------------------------------------------------------
# FORMAT detection (structural)
# ---------------------------------------------------------------------------

EXPECTED_HEADERS_CURRENT = ["项目", "附注", "本年年末余额", "上年年末余额"]
ANCHOR_ITEMS_CURRENT_ASSETS = [
    "货币资金", "应收票据", "应收账款", "存货", "合同资产",
    "一年内到期的非流动资产", "其他流动资产", "流动资产合计",
]


def detect_format_anomaly(
    headers: List[str],
    section_items_raw: List[str],
    anchor_items: List[str],
) -> Tuple[bool, Dict[str, Any]]:
    """
    Returns (is_format_anomaly, details).
    """
    details: Dict[str, Any] = {}
    # A. Header layout
    expected_keywords = ["项目", "本年年末", "上年年末", "期末", "期初", "附注"]
    found = sum(1 for h in headers for k in expected_keywords if k in str(h))
    if found < 2:
        details["header_abnormal"] = True
        details["headers_found"] = headers
    # C. Missing anchors
    normalized_anchors = set(re.sub(r"\s+", "", a) for a in anchor_items)
    normalized_section = set(normalize_for_anomaly_match(i) for i in section_items_raw)
    missing = [a for a in anchor_items if re.sub(r"\s+", "", a) not in normalized_section]
    if len(missing) >= 3:
        details["missing_anchors"] = missing
    is_anomaly = bool(details.get("header_abnormal") or details.get("missing_anchors"))
    return is_anomaly, details


# ---------------------------------------------------------------------------
# AGGREGATION detection (merged items)
# ---------------------------------------------------------------------------

MERGE_CONNECTORS = re.compile(r"[及和与以及合并合计其中包含、]")
KNOWN_MERGED_PATTERNS: List[Tuple[str, List[str]]] = [
    ("应收票据及应收账款", ["NOTESRECE", "ACCORECE"]),
    ("其他应收款及应收利息", ["OTHERRECE", "INTERECE"]),
]


def detect_aggregation_anomaly(item_raw: str, matched_canonical_codes: List[str]) -> Tuple[bool, Dict[str, Any]]:
    """
    One row matches multiple canonical items or known merged pattern.
    """
    details: Dict[str, Any] = {}
    key = normalize_for_anomaly_match(item_raw)
    for pattern, codes in KNOWN_MERGED_PATTERNS:
        if normalize_for_anomaly_match(pattern) in key or key in normalize_for_anomaly_match(pattern):
            details["merged_candidates"] = codes
            details["known_pattern"] = pattern
            return True, details
    if MERGE_CONNECTORS.search(item_raw) and len(matched_canonical_codes) > 1:
        details["merged_candidates"] = matched_canonical_codes
        return True, details
    return False, details


# ---------------------------------------------------------------------------
# IMBALANCE detection (numeric)
# ---------------------------------------------------------------------------

def compute_imbalance(
    parent_value: Optional[float],
    children_values: List[Optional[float]],
    tolerance_ratio: float = 0.0001,
) -> Tuple[Optional[float], Optional[float], bool]:
    """
    Returns (children_sum, diff, is_imbalance).
    """
    if parent_value is None:
        return None, None, False
    valid_children = [v for v in children_values if v is not None]
    if not valid_children:
        return None, None, True
    children_sum = sum(valid_children)
    diff = parent_value - children_sum
    tol = max(1.0, abs(parent_value) * tolerance_ratio)
    is_imbalance = abs(diff) > tol
    return children_sum, diff, is_imbalance


# ---------------------------------------------------------------------------
# Confidence heuristic
# ---------------------------------------------------------------------------

def apply_confidence_heuristic(
    base: float,
    has_format: bool = False,
    has_aggregation: bool = False,
    has_imbalance: bool = False,
    embedding_only_match: bool = False,
) -> float:
    if has_format:
        base *= 0.7
    if has_aggregation:
        base *= 0.8
    if has_imbalance:
        base *= 0.6
    if embedding_only_match:
        base *= 0.85
    return round(min(1.0, max(0.0, base)), 4)


# ---------------------------------------------------------------------------
# Serialization for JSON/DB
# ---------------------------------------------------------------------------

def table_item_to_dict(t: TableItem) -> Dict[str, Any]:
    return {
        "doc_id": t.doc_id,
        "statement_type": t.statement_type,
        "period_end": t.period_end,
        "section": t.section,
        "item_raw": t.item_raw,
        "item_std": t.item_std,
        "value": t.value,
        "value_prev": t.value_prev,
        "unit": t.unit,
        "note_ref": t.note_ref,
        "confidence": t.confidence,
        "anomaly_flags": t.anomaly_flags,
        "anomaly_details": t.anomaly_details,
        "row_type": t.row_type,
    }


def table_section_to_dict(s: TableSection) -> Dict[str, Any]:
    return {
        "doc_id": s.doc_id,
        "statement_type": s.statement_type,
        "period_end": s.period_end,
        "section": s.section,
        "parent_item_std": s.parent_item_std,
        "parent_value": s.parent_value,
        "children_sum": s.children_sum,
        "diff": s.diff,
        "anomaly_flags": s.anomaly_flags,
        "confidence": s.confidence,
        "anomaly_details": s.anomaly_details,
    }


def build_anomaly_report(
    table_items: List[TableItem],
    table_sections: List[TableSection],
) -> Dict[str, Any]:
    """Minimal anomaly report for demo (ACL friendly)."""
    sections_flagged: Dict[str, List[str]] = {}
    for s in table_sections:
        if s.anomaly_flags:
            sections_flagged.setdefault(s.section, []).extend(s.anomaly_flags)
    for k in sections_flagged:
        sections_flagged[k] = list(dict.fromkeys(sections_flagged[k]))

    return {
        "sections_flagged": sections_flagged,
        "imbalance_diffs": {
            s.section: s.diff for s in table_sections if s.diff is not None
        },
        "aggregation_merged_candidates": [
            (i.item_raw, i.anomaly_details.get("merged_candidates", []))
            for i in table_items
            if "AGGREGATION" in i.anomaly_flags
        ],
        "format_missing_anchors": {
            s.section: s.anomaly_details.get("missing_anchors", [])
            for s in table_sections
            if "FORMAT" in s.anomaly_flags
        },
    }


# ---------------------------------------------------------------------------
# 7) Minimal outputs for demo (ACL demo friendly)
# ---------------------------------------------------------------------------

def save_demo_outputs(
    table_items: List[TableItem],
    table_sections: List[TableSection],
    out_dir: str = ".",
    doc_id: str = "demo",
) -> Tuple[str, str]:
    """
    Write structured JSON (table_items, table_sections) and anomaly report.
    Returns (path_to_json, path_to_report).
    """
    json_path = os.path.join(out_dir, f"{doc_id}_balance_sheet_items.json")
    report_path = os.path.join(out_dir, f"{doc_id}_anomaly_report.json")

    payload = {
        "table_items": [table_item_to_dict(i) for i in table_items],
        "table_sections": [table_section_to_dict(s) for s in table_sections],
    }
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)

    report = build_anomaly_report(table_items, table_sections)
    with open(report_path, "w", encoding="utf-8") as f:
        json.dump(report, f, ensure_ascii=False, indent=2)

    return json_path, report_path
