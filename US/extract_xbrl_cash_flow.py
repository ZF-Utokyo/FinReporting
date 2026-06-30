#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
SEC XBRL Cash Flow Statement Extractor

直接从SEC抓取XBRL数据，解析us-gaap标签，自动生成US_FIN_CF表。

Usage:
    python extract_xbrl_cash_flow.py --cik 0000104169 --report-date 2025-01-31 --symbol WMT
    python extract_xbrl_cash_flow.py --cik 0000104169 --report-date 2025-01-31 --symbol WMT --out us_fin_cf.csv
"""

from __future__ import annotations

import argparse
import re
import json
import time
import os
from datetime import datetime, timedelta
from pathlib import Path
from typing import Dict, List, Optional, Tuple
from decimal import Decimal
from dataclasses import dataclass, asdict

import requests
from lxml import etree
import pandas as pd


# -----------------------------
# SEC API Configuration
# -----------------------------

SEC_USER_AGENT = os.getenv(
    "SEC_USER_AGENT",
    "FinReporting/1.0 (contact: your-email@example.com)",
)

SEC_HEADERS = {
    "User-Agent": SEC_USER_AGENT,
    "Accept-Encoding": "gzip, deflate",
    "Host": "data.sec.gov",
}

ARCH_HEADERS = {
    "User-Agent": SEC_USER_AGENT,
    "Accept-Encoding": "gzip, deflate",
    "Host": "www.sec.gov",
}

# Rate limiting: SEC要求请求间隔至少0.1秒
REQUEST_DELAY = 0.1


# -----------------------------
# Data Structures
# -----------------------------

@dataclass
class FilingInfo:
    form: str
    report_date: str
    filing_date: str
    accession_number: str
    primary_document: str


@dataclass
class XBRLCashFlowRecord:
    """US_FIN_CF表的一条记录"""
    symbol: str
    form_type: str
    fiscal_year_end_date: str
    filing_date: str
    accession_number: str
    net_income: Optional[float] = None
    net_cash_operating: Optional[float] = None
    net_cash_investing: Optional[float] = None
    net_cash_financing: Optional[float] = None
    effect_of_exchange_rates_on_cash: Optional[float] = None
    net_change_in_cash: Optional[float] = None
    cash_beginning_of_period: Optional[float] = None
    cash_end_of_period: Optional[float] = None
    currency: str = "USD"
    created_at: Optional[str] = None


# -----------------------------
# SEC API Functions
# -----------------------------

def get_submissions_json(cik: str) -> dict:
    """获取CIK的submissions JSON"""
    cik10 = cik.zfill(10)
    url = f"https://data.sec.gov/submissions/CIK{cik10}.json"
    
    time.sleep(REQUEST_DELAY)
    r = requests.get(url, headers=SEC_HEADERS, timeout=30)
    r.raise_for_status()
    return r.json()


def find_filing(sub: dict, form_type: str, report_date: str) -> List[FilingInfo]:
    """
    从submissions JSON中找到匹配的filing
    
    Args:
        sub: submissions JSON字典
        form_type: 表单类型，如"10-K"
        report_date: 报告日期，格式"YYYY-MM-DD"
    
    Returns:
        匹配的FilingInfo列表
    """
    r = sub["filings"]["recent"]
    forms = r["form"]
    report_dates = r["reportDate"]
    filing_dates = r["filingDate"]
    accession_numbers = r["accessionNumber"]
    primary_docs = r["primaryDocument"]

    hits = []
    for i, (f, rd, fd) in enumerate(zip(forms, report_dates, filing_dates)):
        # 排除修正版本（10-K/A）
        if f == form_type and rd == report_date and not f.endswith("/A"):
            hits.append(FilingInfo(
                form=f,
                report_date=rd,
                filing_date=fd,
                accession_number=accession_numbers[i],
                primary_document=primary_docs[i],
            ))
    return hits


def accession_to_path(acc: str) -> str:
    """将accession number转换为路径格式"""
    # "0000104169-26-0000xx" -> "0000104169260000xx"
    return acc.replace("-", "")


def download_filing_index(cik: str, accession: str) -> dict:
    """下载filing目录的index.json"""
    cik_no0 = str(int(cik))  # archives path用无前导0的cik
    acc_path = accession_to_path(accession)
    url = f"https://www.sec.gov/Archives/edgar/data/{cik_no0}/{acc_path}/index.json"
    
    time.sleep(REQUEST_DELAY)
    r = requests.get(url, headers=ARCH_HEADERS, timeout=30)
    r.raise_for_status()
    return r.json()


def pick_instance_file(index_json: dict) -> str:
    """
    从index.json中选择XBRL instance文件
    
    优先级：
    1. *_htm.xml (inline XBRL instance)
    2. *.xml (standalone instance，排除schema/cal/def/pre/lab)
    """
    names = [item["name"] for item in index_json["directory"]["item"]]
    
    # 优先 *_htm.xml
    for n in names:
        if n.endswith("_htm.xml"):
            return n
    
    # 再找 .xml 里看起来像 instance（排除schema/cal/def/pre/lab）
    for n in names:
        if n.endswith(".xml") and not re.search(r"(_cal|_def|_pre|_lab|\.xsd)$", n):
            return n
    
    raise RuntimeError("No XBRL instance XML found in index.json")


def download_instance_xml(cik: str, accession: str, filename: str) -> bytes:
    """下载XBRL instance XML文件"""
    cik_no0 = str(int(cik))
    acc_path = accession_to_path(accession)
    url = f"https://www.sec.gov/Archives/edgar/data/{cik_no0}/{acc_path}/{filename}"
    
    time.sleep(REQUEST_DELAY)
    r = requests.get(url, headers=ARCH_HEADERS, timeout=30)
    r.raise_for_status()
    return r.content


# -----------------------------
# XBRL Parsing Functions
# -----------------------------

def parse_xbrl_facts(xml_bytes: bytes) -> Tuple[List[Dict], etree.Element]:
    """
    解析XBRL instance XML，提取所有facts
    
    Returns:
        (facts列表, root元素)
    """
    root = etree.fromstring(xml_bytes)
    facts = []
    
    for el in root.iter():
        if not isinstance(el.tag, str):
            continue
        
        # XBRL fact通常带namespace，形如 "{ns}NetIncomeLoss"
        if el.get("contextRef") and (el.text is not None):
            tag = el.tag.split("}")[-1]  # 去掉namespace
            facts.append({
                "tag": tag,
                "contextRef": el.get("contextRef"),
                "unitRef": el.get("unitRef"),
                "decimals": el.get("decimals"),
                "value": el.text.strip(),
            })
    
    return facts, root


def context_end_dates(root: etree.Element) -> Dict[str, str]:
    """
    建立contextRef -> endDate映射
    
    Returns:
        {context_id: end_date}
    """
    ns = {
        "xbrli": "http://www.xbrl.org/2003/instance",
    }
    ctx_map = {}
    
    for ctx in root.findall(".//xbrli:context", namespaces=ns):
        cid = ctx.get("id")
        end = ctx.find(".//xbrli:period/xbrli:endDate", namespaces=ns)
        inst = ctx.find(".//xbrli:period/xbrli:instant", namespaces=ns)
        
        if end is not None:
            ctx_map[cid] = end.text.strip()
        elif inst is not None:
            ctx_map[cid] = inst.text.strip()
    
    return ctx_map


def context_period_info(root: etree.Element) -> Dict[str, Dict[str, Optional[str]]]:
    """
    建立contextRef -> period信息映射

    Returns:
        {
            context_id: {
                "start_date": str | None,
                "end_date": str | None,
                "instant": str | None,
            }
        }
    """
    ns = {
        "xbrli": "http://www.xbrl.org/2003/instance",
    }
    info_map: Dict[str, Dict[str, Optional[str]]] = {}

    for ctx in root.findall(".//xbrli:context", namespaces=ns):
        cid = ctx.get("id")
        if not cid:
            continue

        start = ctx.find(".//xbrli:period/xbrli:startDate", namespaces=ns)
        end = ctx.find(".//xbrli:period/xbrli:endDate", namespaces=ns)
        inst = ctx.find(".//xbrli:period/xbrli:instant", namespaces=ns)

        info_map[cid] = {
            "start_date": start.text.strip() if start is not None and start.text else None,
            "end_date": end.text.strip() if end is not None and end.text else None,
            "instant": inst.text.strip() if inst is not None and inst.text else None,
        }

    return info_map


def context_is_consolidated(root: etree.Element, context_id: str) -> bool:
    """
    判断context是否为合并口径（consolidated）
    
    规则：如果context没有segments，通常是consolidated
    """
    ns = {
        "xbrli": "http://www.xbrl.org/2003/instance",
    }
    
    for ctx in root.findall(".//xbrli:context", namespaces=ns):
        if ctx.get("id") == context_id:
            segments = ctx.findall(".//xbrli:segment", namespaces=ns)
            # 没有segments通常表示consolidated
            return len(segments) == 0
    
    return False


def infer_prior_period_end_date(
    context_map: Dict[str, Dict[str, Optional[str]]],
    report_date: str,
    root: etree.Element,
) -> Optional[str]:
    """
    从duration context推断上一期期末日期（通常用于期初现金）。

    逻辑：
    1. 找end_date=report_date且有start_date的context
    2. 优先consolidated context
    3. 取start_date - 1天
    """
    duration_contexts: List[Tuple[str, str, int]] = []
    for cid, period in context_map.items():
        start_date = period.get("start_date")
        end_date = period.get("end_date")
        if not start_date or end_date != report_date:
            continue

        try:
            start_dt = datetime.strptime(start_date, "%Y-%m-%d").date()
            end_dt = datetime.strptime(end_date, "%Y-%m-%d").date()
        except ValueError:
            continue

        duration_days = (end_dt - start_dt).days
        # 排除无效/零长度duration context
        if duration_days <= 0:
            continue

        duration_contexts.append((cid, start_date, duration_days))

    if not duration_contexts:
        return None

    consolidated = [
        (cid, start_date, duration_days)
        for cid, start_date, duration_days in duration_contexts
        if context_is_consolidated(root, cid)
    ]
    candidates = consolidated or duration_contexts

    # 选最长duration（通常是年报口径），若相同再取最新start_date
    cid, chosen_start, _ = max(candidates, key=lambda x: (x[2], x[1]))

    try:
        prior_end = datetime.strptime(chosen_start, "%Y-%m-%d").date() - timedelta(days=1)
        return prior_end.isoformat()
    except ValueError:
        return None


def pick_value_for_enddate(
    facts: List[Dict],
    ctx_end_map: Dict[str, str],
    tag_name: str,
    end_date: str,
    root: Optional[etree.Element] = None,
) -> Optional[Dict]:
    """
    选出tag=tag_name且context endDate匹配的fact
    
    优先选择：
    1. Consolidated（无segments）
    2. USD单位的
    """
    candidates = [
        f for f in facts
        if f["tag"] == tag_name
        and ctx_end_map.get(f["contextRef"]) == end_date
    ]
    
    if not candidates:
        return None
    
    # 如果有root，优先选择consolidated（无segments）的
    if root is not None:
        consolidated_candidates = []
        for c in candidates:
            if context_is_consolidated(root, c["contextRef"]):
                consolidated_candidates.append(c)
        
        if consolidated_candidates:
            candidates = consolidated_candidates
    
    # 优先选择USD单位的
    for c in candidates:
        unit_ref = (c.get("unitRef") or "").upper()
        if unit_ref.startswith("USD") or "USD" in unit_ref:
            return c
    
    return candidates[0]


def parse_numeric_value(value_str: str, decimals: Optional[str] = None) -> Optional[float]:
    """
    解析XBRL数值，考虑decimals精度
    
    XBRL中decimals字段的含义：
    - decimals="0" 表示整数，无小数位
    - decimals="-6" 表示数值已经是以百万为单位存储的（即已经除以了10^6）
      在这种情况下，数值已经是真实值，不需要再缩放
    - decimals="2" 表示有2位小数
    
    重要：在XBRL中，decimals="-6"通常意味着数值已经是以百万为单位，
    但实际存储的数值已经是真实值（只是显示时会考虑小数位）。
    我们需要直接使用数值，不需要缩放。
    
    实际上，SEC的XBRL文件中，decimals="-6"的数值已经是真实值（美元），
    不需要乘以10^6。如果乘以了，会导致数值过大。
    """
    try:
        value = float(value_str.replace(",", ""))
        
        # 注意：SEC XBRL中的数值通常已经是真实值（USD）
        # decimals字段主要用于表示显示精度，而不是缩放因子
        # 所以这里直接返回数值，不做缩放处理
        
        return value
    except (ValueError, AttributeError):
        return None


# -----------------------------
# US-GAAP Tag Mapping
# -----------------------------

# 常见us-gaap标签映射（按优先级）
TAG_MAP = {
    "NET_INCOME": [
        "NetIncomeLoss",
        "ProfitLoss",
        "IncomeLossFromContinuingOperations",
    ],
    "NET_CASH_OPERATING": [
        "NetCashProvidedByUsedInOperatingActivities",
        "NetCashProvidedByUsedInOperatingActivitiesContinuingOperations",
    ],
    "NET_CASH_INVESTING": [
        "NetCashProvidedByUsedInInvestingActivities",
        "NetCashProvidedByUsedInInvestingActivitiesContinuingOperations",
    ],
    "NET_CASH_FINANCING": [
        "NetCashProvidedByUsedInFinancingActivities",
        "NetCashProvidedByUsedInFinancingActivitiesContinuingOperations",
    ],
    "EFFECT_OF_EXCHANGE_RATES_ON_CASH": [
        "EffectOfExchangeRateOnCashAndCashEquivalents",
        "EffectOfExchangeRateChangesOnCash",
        "EffectOfExchangeRateChangesOnCashAndCashEquivalents",
        "EffectOfExchangeRateChangesOnCashCashEquivalents",
        "EffectOfExchangeRateChangesOnCashCashEquivalentsAndRestrictedCash",
        "EffectOfExchangeRateChangesOnCashCashEquivalentsAndRestrictedCashEquivalents",
        "EffectOfExchangeRateOnCashCashEquivalentsAndRestrictedCash",
        "EffectOfExchangeRateOnCashCashEquivalentsRestrictedCashAndRestrictedCashEquivalents",
    ],
    "NET_CHANGE_IN_CASH": [
        "CashCashEquivalentsRestrictedCashAndRestrictedCashEquivalentsPeriodIncreaseDecreaseIncludingExchangeRateEffect",
        "CashAndCashEquivalentsPeriodIncreaseDecreaseIncludingExchangeRateEffect",
        "CashCashEquivalentsAndRestrictedCashPeriodIncreaseDecreaseIncludingExchangeRateEffect",
    ],
    "CASH_BEGINNING_OF_PERIOD": [
        "CashCashEquivalentsRestrictedCashAndRestrictedCashEquivalents",
        "CashAndCashEquivalentsAtCarryingValue",
    ],
    "CASH_END_OF_PERIOD": [
        "CashCashEquivalentsRestrictedCashAndRestrictedCashEquivalents",
        "CashAndCashEquivalentsAtCarryingValue",
    ],
}

FX_EFFECT_LABEL_ALIASES = [
    "effect of exchange rate changes on cash",
    "effect of exchange rate changes on cash cash equivalents",
    "effect of exchange rate changes on cash cash equivalents and restricted cash",
]


def _normalize_label_for_match(text: str) -> str:
    s = str(text or "")
    s = re.sub(r"([a-z0-9])([A-Z])", r"\1 \2", s)
    s = s.lower()
    s = re.sub(r"[^a-z0-9]+", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


def find_value_by_label_alias(
    facts: List[Dict],
    ctx_end_map: Dict[str, str],
    label_aliases: List[str],
    end_date: str,
    root: Optional[etree.Element] = None,
) -> Optional[float]:
    aliases = [_normalize_label_for_match(x) for x in label_aliases if str(x).strip()]
    if not aliases:
        return None

    candidates: List[Tuple[int, float]] = []
    for fact in facts:
        if ctx_end_map.get(fact.get("contextRef")) != end_date:
            continue

        value = parse_numeric_value(fact.get("value", ""), fact.get("decimals"))
        if value is None:
            continue

        normalized = _normalize_label_for_match(fact.get("tag", ""))
        if not normalized:
            continue
        if not any(alias in normalized for alias in aliases):
            continue

        score = 0
        if root is not None and context_is_consolidated(root, fact.get("contextRef", "")):
            score += 100
        unit_ref = str(fact.get("unitRef") or "").upper()
        if "USD" in unit_ref:
            score += 20
        if normalized.startswith("effect of exchange rate changes on cash"):
            score += 10
        candidates.append((score, value))

    if not candidates:
        return None
    return sorted(candidates, key=lambda x: x[0], reverse=True)[0][1]


def find_tag_value(
    facts: List[Dict],
    ctx_end_map: Dict[str, str],
    tag_candidates: List[str],
    end_date: str,
    root: Optional[etree.Element] = None,
    period_type: str = "end",
) -> Optional[float]:
    """
    尝试多个tag候选，找到第一个有值的（优先consolidated）
    
    Args:
        period_type: "end" 表示期末值（instant），"begin" 表示期初值
        root: XML root元素，用于判断consolidated
    """
    for tag_name in tag_candidates:
        fact = pick_value_for_enddate(facts, ctx_end_map, tag_name, end_date, root)
        if fact:
            value = parse_numeric_value(fact["value"], fact.get("decimals"))
            if value is not None:
                return value
    
    return None


def extract_cash_flow_data(
    facts: List[Dict],
    ctx_end_map: Dict[str, str],
    context_map: Dict[str, Dict[str, Optional[str]]],
    report_date: str,
    root: etree.Element,
) -> Dict[str, Optional[float]]:
    """
    从XBRL facts中提取现金流量表数据
    
    Returns:
        字段名到数值的映射
    """
    result = {}
    
    # 期末现金（instant）
    result["cash_end_of_period"] = find_tag_value(
        facts, ctx_end_map,
        TAG_MAP["CASH_END_OF_PERIOD"],
        report_date,
        root,
        "end",
    )
    
    # 期初现金：由duration context的startDate推断上一期期末日期后，再取instant值
    prior_period_end = infer_prior_period_end_date(context_map, report_date, root)
    if prior_period_end:
        result["cash_beginning_of_period"] = find_tag_value(
            facts, ctx_end_map,
            TAG_MAP["CASH_BEGINNING_OF_PERIOD"],
            prior_period_end,
            root,
            "begin",
        )
    else:
        result["cash_beginning_of_period"] = None
    
    # 期间现金流（period，endDate=report_date）
    # Consolidated Net Income = NetIncomeLoss + NetIncomeLossAttributableToNoncontrollingInterest
    # 优先选择consolidated net income（无segments的context）
    net_income_wmt = find_tag_value(
        facts, ctx_end_map,
        TAG_MAP["NET_INCOME"],
        report_date,
        root,
        "end",
    )
    
    # 查找少数股东权益的net income
    nci_tags = [
        "NetIncomeLossAttributableToNoncontrollingInterest",
        "NetIncomeLossAttributableToNonredeemableNoncontrollingInterest",
    ]
    net_income_nci = find_tag_value(
        facts, ctx_end_map,
        nci_tags,
        report_date,
        root,
        "end",
    )
    
    # Consolidated Net Income = NetIncomeLoss + NCI
    if net_income_wmt is not None:
        if net_income_nci is not None:
            # 如果NCI是负数，需要加上（因为它是loss）
            result["net_income"] = net_income_wmt + net_income_nci
        else:
            # 如果没有NCI标签，使用NetIncomeLoss（可能已经是consolidated的）
            result["net_income"] = net_income_wmt
    else:
        result["net_income"] = None
    
    result["net_cash_operating"] = find_tag_value(
        facts, ctx_end_map,
        TAG_MAP["NET_CASH_OPERATING"],
        report_date,
        root,
        "end",
    )
    
    result["net_cash_investing"] = find_tag_value(
        facts, ctx_end_map,
        TAG_MAP["NET_CASH_INVESTING"],
        report_date,
        root,
        "end",
    )
    
    result["net_cash_financing"] = find_tag_value(
        facts, ctx_end_map,
        TAG_MAP["NET_CASH_FINANCING"],
        report_date,
        root,
        "end",
    )
    
    result["effect_of_exchange_rates_on_cash"] = find_tag_value(
        facts, ctx_end_map,
        TAG_MAP["EFFECT_OF_EXCHANGE_RATES_ON_CASH"],
        report_date,
        root,
        "end",
    )
    if result["effect_of_exchange_rates_on_cash"] is None:
        # Fallback: label-alias mapping (case-insensitive, punctuation-insensitive).
        result["effect_of_exchange_rates_on_cash"] = find_value_by_label_alias(
            facts,
            ctx_end_map,
            FX_EFFECT_LABEL_ALIASES,
            report_date,
            root,
        )
    
    result["net_change_in_cash"] = find_tag_value(
        facts, ctx_end_map,
        TAG_MAP["NET_CHANGE_IN_CASH"],
        report_date,
        root,
        "end",
    )
    
    return result


# -----------------------------
# Main Pipeline
# -----------------------------

def extract_xbrl_cash_flow(
    cik: str,
    report_date: str,
    symbol: str,
    form_type: str = "10-K",
) -> XBRLCashFlowRecord:
    """
    主流程：从SEC抓取XBRL并提取现金流量表数据
    
    Returns:
        XBRLCashFlowRecord对象
    """
    print(f"[INFO] Fetching submissions JSON for CIK {cik}...")
    sub = get_submissions_json(cik)
    
    print(f"[INFO] Finding {form_type} filing with reportDate={report_date}...")
    hits = find_filing(sub, form_type, report_date)
    
    if not hits:
        raise RuntimeError(f"No matching {form_type} found for reportDate={report_date}")
    
    filing = hits[0]  # 取第一条
    print(f"[INFO] Found filing: {filing.accession_number} (filed on {filing.filing_date})")
    
    print(f"[INFO] Downloading filing index...")
    index_json = download_filing_index(cik, filing.accession_number)
    
    print(f"[INFO] Picking XBRL instance file...")
    instance_file = pick_instance_file(index_json)
    print(f"[INFO] Instance file: {instance_file}")
    
    print(f"[INFO] Downloading XBRL instance XML...")
    xml_bytes = download_instance_xml(cik, filing.accession_number, instance_file)
    
    print(f"[INFO] Parsing XBRL facts...")
    facts, root = parse_xbrl_facts(xml_bytes)
    print(f"[INFO] Found {len(facts)} facts")
    
    print(f"[INFO] Building context end date map...")
    ctx_end_map = context_end_dates(root)
    context_map = context_period_info(root)
    
    print(f"[INFO] Extracting cash flow data...")
    cf_data = extract_cash_flow_data(facts, ctx_end_map, context_map, report_date, root)
    
    # 构建记录
    record = XBRLCashFlowRecord(
        symbol=symbol,
        form_type=form_type,
        fiscal_year_end_date=report_date,
        filing_date=filing.filing_date,
        accession_number=filing.accession_number,
        **cf_data,
    )
    
    return record


def main():
    parser = argparse.ArgumentParser(
        description="Extract US cash flow statement from SEC XBRL data"
    )
    parser.add_argument("--cik", required=True, help="CIK number (e.g., 0000104169)")
    parser.add_argument("--report-date", required=True, help="Fiscal year end date (YYYY-MM-DD)")
    parser.add_argument("--symbol", required=True, help="Stock symbol (e.g., WMT)")
    parser.add_argument("--form-type", default="10-K", help="Form type (default: 10-K)")
    parser.add_argument("--out", help="Output CSV file path")
    
    args = parser.parse_args()
    
    try:
        record = extract_xbrl_cash_flow(
            cik=args.cik,
            report_date=args.report_date,
            symbol=args.symbol,
            form_type=args.form_type,
        )
        
        # 打印结果
        print("\n" + "="*60)
        print("Extracted Cash Flow Data:")
        print("="*60)
        print(f"Symbol: {record.symbol}")
        print(f"Form Type: {record.form_type}")
        print(f"Fiscal Year End: {record.fiscal_year_end_date}")
        print(f"Filing Date: {record.filing_date}")
        print(f"Accession Number: {record.accession_number}")
        print("\nCash Flow Items:")
        print(f"  Consolidated Net Income: {record.net_income}")
        print(f"  Net Cash from Operating: {record.net_cash_operating}")
        print(f"  Net Cash from Investing: {record.net_cash_investing}")
        print(f"  Net Cash from Financing: {record.net_cash_financing}")
        print(f"  Effect of Exchange Rates: {record.effect_of_exchange_rates_on_cash}")
        print(f"  Net Change in Cash: {record.net_change_in_cash}")
        print(f"  Cash Beginning: {record.cash_beginning_of_period}")
        print(f"  Cash End: {record.cash_end_of_period}")
        print("="*60)
        
        # 保存到CSV
        if args.out:
            df = pd.DataFrame([asdict(record)])
            df.to_csv(args.out, index=False)
            print(f"\n[OK] Saved to {args.out}")
        
    except Exception as e:
        print(f"\n[ERROR] {e}", flush=True)
        raise


if __name__ == "__main__":
    main()
