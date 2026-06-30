#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Export US Income Statement / Balance Sheet / Cash Flow to a single Excel file.

Usage:
  python export_three_statements_excel.py --symbol AAPL --cik 0000320193 --out aapl_3statements.xlsx
"""

from __future__ import annotations

import argparse
from dataclasses import asdict
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd
from lxml import etree

from extract_xbrl_cash_flow import (
    TAG_MAP,
    download_filing_index,
    download_instance_xml,
    extract_cash_flow_data,
    get_submissions_json,
    parse_numeric_value,
    parse_xbrl_facts,
    pick_instance_file,
    context_period_info,
)


XBRLI_NS = {"xbrli": "http://www.xbrl.org/2003/instance"}

US_LLM_FIELD_SPECS: Dict[Tuple[str, str], Dict[str, Any]] = {
    ("IS", "total_revenue"): {
        "tags": ["Revenues", "SalesRevenueNet", "RevenueFromContractWithCustomerExcludingAssessedTax"],
        "period_type": "duration",
        "unit_keywords": ["usd"],
    },
    ("IS", "operating_income"): {
        "tags": ["OperatingIncomeLoss"],
        "period_type": "duration",
        "unit_keywords": ["usd"],
    },
    ("IS", "income_before_income_taxes"): {
        "tags": [
            "IncomeBeforeTax",
            "IncomeBeforeTaxExpenseBenefit",
            "IncomeLossFromContinuingOperationsBeforeIncomeTaxes",
            "IncomeLossFromContinuingOperationsBeforeIncomeTaxesExtraordinaryItemsNoncontrollingInterest",
            "IncomeLossFromContinuingOperationsBeforeIncomeTaxesMinorityInterestAndIncomeLossFromEquityMethodInvestments",
            "IncomeLossFromContinuingOperationsAfterNoncontrollingInterestBeforeIncomeTaxes",
            "IncomeLossFromContinuingOperationsIncludingNoncontrollingInterestBeforeIncomeTaxesExtraordinaryItems",
        ],
        "period_type": "duration",
        "unit_keywords": ["usd"],
    },
    ("IS", "net_income"): {
        "tags": ["NetIncomeLoss", "ProfitLoss", "NetIncomeLossAvailableToCommonStockholdersBasic"],
        "period_type": "duration",
        "unit_keywords": ["usd"],
    },
    ("IS", "net_income_per_share_basic"): {
        "tags": ["EarningsPerShareBasic"],
        "period_type": "duration",
        "unit_keywords": ["usd/shares", "pure"],
    },
    ("BS", "total_assets"): {
        "tags": ["Assets"],
        "period_type": "instant",
        "unit_keywords": ["usd"],
    },
    ("BS", "cash_and_cash_equivalents"): {
        "tags": ["CashAndCashEquivalentsAtCarryingValue", "CashCashEquivalentsRestrictedCashAndRestrictedCashEquivalents"],
        "period_type": "instant",
        "unit_keywords": ["usd"],
    },
    ("BS", "accounts_receivable"): {
        "tags": ["AccountsReceivableNetCurrent", "ReceivablesNetCurrent"],
        "period_type": "instant",
        "unit_keywords": ["usd"],
    },
    ("BS", "inventories"): {
        "tags": ["InventoryNet"],
        "period_type": "instant",
        "unit_keywords": ["usd"],
    },
    ("BS", "total_liabilities"): {
        "tags": ["Liabilities"],
        "period_type": "instant",
        "unit_keywords": ["usd"],
    },
    ("BS", "total_shareholders_equity"): {
        "tags": ["StockholdersEquity", "StockholdersEquityIncludingPortionAttributableToNoncontrollingInterest"],
        "period_type": "instant",
        "unit_keywords": ["usd"],
    },
    ("BS", "total_liabilities_and_shareholders_equity"): {
        "tags": ["LiabilitiesAndStockholdersEquity"],
        "period_type": "instant",
        "unit_keywords": ["usd"],
    },
    ("CF", "net_income"): {
        "tags": TAG_MAP["NET_INCOME"],
        "period_type": "duration",
        "unit_keywords": ["usd"],
    },
    ("CF", "net_cash_operating"): {
        "tags": TAG_MAP["NET_CASH_OPERATING"],
        "period_type": "duration",
        "unit_keywords": ["usd"],
    },
    ("CF", "net_cash_investing"): {
        "tags": TAG_MAP["NET_CASH_INVESTING"],
        "period_type": "duration",
        "unit_keywords": ["usd"],
    },
    ("CF", "net_cash_financing"): {
        "tags": TAG_MAP["NET_CASH_FINANCING"],
        "period_type": "duration",
        "unit_keywords": ["usd"],
    },
    ("CF", "net_change_in_cash"): {
        "tags": TAG_MAP["NET_CHANGE_IN_CASH"],
        "period_type": "duration",
        "unit_keywords": ["usd"],
    },
    ("CF", "cash_end_of_period"): {
        "tags": TAG_MAP["CASH_END_OF_PERIOD"],
        "period_type": "instant",
        "unit_keywords": ["usd"],
    },
}


def latest_annual_filing(sub: dict, form_type: str = "10-K") -> Tuple[dict, int]:
    recent = sub["filings"]["recent"]
    forms = recent["form"]
    report_dates = recent["reportDate"]

    for i, form in enumerate(forms):
        if form != form_type:
            continue
        if form.endswith("/A"):
            continue
        if not report_dates[i]:
            continue
        return recent, i

    raise RuntimeError(f"No {form_type} filing found in recent submissions.")


def parse_contexts(root: etree._Element) -> Dict[str, Dict[str, Optional[object]]]:
    contexts: Dict[str, Dict[str, Optional[object]]] = {}
    for ctx in root.findall(".//xbrli:context", namespaces=XBRLI_NS):
        cid = ctx.get("id")
        if not cid:
            continue

        start = ctx.find(".//xbrli:period/xbrli:startDate", namespaces=XBRLI_NS)
        end = ctx.find(".//xbrli:period/xbrli:endDate", namespaces=XBRLI_NS)
        instant = ctx.find(".//xbrli:period/xbrli:instant", namespaces=XBRLI_NS)
        segments = ctx.findall(".//xbrli:segment", namespaces=XBRLI_NS)

        start_date = start.text.strip() if start is not None and start.text else None
        end_date = end.text.strip() if end is not None and end.text else None
        instant_date = instant.text.strip() if instant is not None and instant.text else None
        duration_days = None
        if start_date and end_date:
            try:
                duration_days = (datetime.strptime(end_date, "%Y-%m-%d") - datetime.strptime(start_date, "%Y-%m-%d")).days
            except ValueError:
                duration_days = None

        contexts[cid] = {
            "start_date": start_date,
            "end_date": end_date,
            "instant": instant_date,
            "duration_days": duration_days,
            "is_consolidated": len(segments) == 0,
        }

    return contexts


def choose_value(
    facts: List[Dict],
    contexts: Dict[str, Dict[str, Optional[object]]],
    report_date: str,
    tags: List[str],
    period_type: str,
    unit_keywords: Optional[List[str]] = None,
) -> Optional[float]:
    best_fact = choose_fact(facts, contexts, report_date, tags, period_type, unit_keywords=unit_keywords)
    if not best_fact:
        return None
    return parse_numeric_value(best_fact["value"], best_fact.get("decimals"))


def choose_fact(
    facts: List[Dict],
    contexts: Dict[str, Dict[str, Optional[object]]],
    report_date: str,
    tags: List[str],
    period_type: str,
    unit_keywords: Optional[List[str]] = None,
) -> Optional[Dict]:
    def _score_fact(fact: Dict, ctx: Dict[str, Optional[object]]) -> Tuple[int, Dict[str, int]]:
        score = 0
        breakdown = {
            "consolidated": 0,
            "annual_duration": 0,
            "unit_match": 0,
            "tag_priority": 0,
        }
        if ctx.get("is_consolidated"):
            breakdown["consolidated"] = 100
            score += 100

        duration_days = ctx.get("duration_days")
        if period_type == "duration" and isinstance(duration_days, int):
            if 300 <= duration_days <= 380:
                breakdown["annual_duration"] = 30
                score += 30
            else:
                breakdown["annual_duration"] = -5
                score -= 5

        if unit_keywords:
            unit_ref = (fact.get("unitRef") or "").lower()
            if any(k.lower() in unit_ref for k in unit_keywords):
                breakdown["unit_match"] = 10
                score += 10

        tag_priority = max(0, 5 - tags.index(fact["tag"]))
        breakdown["tag_priority"] = tag_priority
        score += tag_priority
        return score, breakdown

    candidates: List[Tuple[int, Dict]] = []

    for fact in facts:
        if fact["tag"] not in tags:
            continue
        ctx = contexts.get(fact["contextRef"])
        if not ctx:
            continue

        if period_type == "duration":
            if ctx.get("end_date") != report_date or not ctx.get("start_date"):
                continue
        elif period_type == "instant":
            if ctx.get("instant") != report_date and ctx.get("end_date") != report_date:
                continue
        else:
            continue

        score, _ = _score_fact(fact, ctx)
        candidates.append((score, fact))

    if not candidates:
        return None

    best_fact = sorted(candidates, key=lambda x: x[0], reverse=True)[0][1]
    return best_fact


def choose_top_facts(
    facts: List[Dict],
    contexts: Dict[str, Dict[str, Optional[object]]],
    report_date: str,
    tags: List[str],
    period_type: str,
    unit_keywords: Optional[List[str]] = None,
    top_k: int = 10,
) -> List[Dict[str, Any]]:
    scored: List[Dict[str, Any]] = []
    seen = set()
    for fact in facts:
        if fact["tag"] not in tags:
            continue
        ctx = contexts.get(fact["contextRef"])
        if not ctx:
            continue
        if period_type == "duration":
            if ctx.get("end_date") != report_date or not ctx.get("start_date"):
                continue
        elif period_type == "instant":
            if ctx.get("instant") != report_date and ctx.get("end_date") != report_date:
                continue
        else:
            continue

        score = 0
        score_breakdown = {
            "consolidated": 0,
            "annual_duration": 0,
            "unit_match": 0,
            "tag_priority": 0,
        }
        if ctx.get("is_consolidated"):
            score_breakdown["consolidated"] = 100
            score += 100
        duration_days = ctx.get("duration_days")
        if period_type == "duration" and isinstance(duration_days, int):
            if 300 <= duration_days <= 380:
                score_breakdown["annual_duration"] = 30
                score += 30
            else:
                score_breakdown["annual_duration"] = -5
                score -= 5
        if unit_keywords:
            unit_ref = (fact.get("unitRef") or "").lower()
            if any(k.lower() in unit_ref for k in unit_keywords):
                score_breakdown["unit_match"] = 10
                score += 10
        tag_priority = max(0, 5 - tags.index(fact["tag"]))
        score_breakdown["tag_priority"] = tag_priority
        score += tag_priority

        value = parse_numeric_value(fact.get("value", ""), fact.get("decimals"))
        dedupe_key = (str(fact.get("tag") or ""), str(fact.get("contextRef") or ""), value)
        if dedupe_key in seen:
            continue
        seen.add(dedupe_key)
        scored.append(
            {
                "tag": fact.get("tag"),
                "value": value,
                "raw_value": fact.get("value"),
                "decimals": fact.get("decimals"),
                "unitRef": fact.get("unitRef"),
                "contextRef": fact.get("contextRef"),
                "start_date": ctx.get("start_date"),
                "end_date": ctx.get("end_date"),
                "instant": ctx.get("instant"),
                "duration_days": ctx.get("duration_days"),
                "is_consolidated": int(bool(ctx.get("is_consolidated"))),
                "score": score,
                "score_breakdown": score_breakdown,
            }
        )

    scored = sorted(scored, key=lambda x: x.get("score", 0), reverse=True)
    return scored[: max(1, top_k)]


def build_us_candidate_rows(
    report_date: str,
    facts: List[Dict],
    contexts: Dict[str, Dict[str, Optional[object]]],
    is_row: Dict[str, Any],
    bs_row: Dict[str, Any],
    cf_row: Dict[str, Any],
    top_k: int = 10,
) -> List[Dict[str, Any]]:
    out: List[Dict[str, Any]] = []
    rule_rows = {"IS": is_row, "BS": bs_row, "CF": cf_row}
    for (stmt, field_name), spec in US_LLM_FIELD_SPECS.items():
        tags = list(spec["tags"])
        period_type = str(spec["period_type"])
        unit_keywords = list(spec.get("unit_keywords") or [])
        tops = choose_top_facts(
            facts,
            contexts,
            report_date,
            tags,
            period_type,
            unit_keywords=unit_keywords,
            top_k=top_k,
        )
        rule_value = rule_rows.get(stmt, {}).get(field_name)
        for i, c in enumerate(tops, start=1):
            row = {
                "statement": stmt,
                "field_name": field_name,
                "candidate_rank": i,
                "candidate_id": f"{stmt}.{field_name}.fact.{i}",
                "tag": c.get("tag"),
                "value": c.get("value"),
                "raw_value": c.get("raw_value"),
                "decimals": c.get("decimals"),
                "unit_ref": c.get("unitRef"),
                "context_ref": c.get("contextRef"),
                "start_date": c.get("start_date"),
                "end_date": c.get("end_date"),
                "instant": c.get("instant"),
                "duration_days": c.get("duration_days"),
                "is_consolidated": c.get("is_consolidated"),
                "period_type": period_type,
                "score": c.get("score"),
                "score_breakdown": str(c.get("score_breakdown")),
                "rule_value": rule_value,
            }
            out.append(row)
    return out


def choose_value_with_source(
    facts: List[Dict],
    contexts: Dict[str, Dict[str, Optional[object]]],
    report_date: str,
    tags: List[str],
    period_type: str,
    unit_keywords: Optional[List[str]] = None,
) -> Tuple[Optional[float], Optional[str]]:
    best_fact = choose_fact(facts, contexts, report_date, tags, period_type, unit_keywords=unit_keywords)
    if not best_fact:
        return None, None
    return parse_numeric_value(best_fact["value"], best_fact.get("decimals")), best_fact.get("tag")


def build_is_row(
    symbol: str,
    cik: str,
    filing: dict,
    report_date: str,
    facts: List[Dict],
    contexts: Dict[str, Dict[str, Optional[object]]],
) -> Dict:
    return {
        "country": "US",
        "symbol": symbol,
        "company_id": cik,
        "form_type": filing["form_type"],
        "fiscal_year_end_date": report_date,
        "filing_date": filing["filing_date"],
        "fiscal_year": filing["fiscal_year"],
        "fiscal_period": filing["fiscal_period"],
        "source_filing_id": filing["accession_number"],
        "is_amendment": 0,
        "accounting_standard": "US_GAAP",
        "taxonomy_version": None,
        "statement_scope": "CONSOLIDATED",
        "currency": "USD",
        "unit_scale": 1,
        "total_revenue": choose_value(
            facts, contexts, report_date,
            ["Revenues", "SalesRevenueNet", "RevenueFromContractWithCustomerExcludingAssessedTax"],
            "duration",
            ["usd"],
        ),
        "cost_of_revenue": choose_value(
            facts, contexts, report_date,
            ["CostOfRevenue", "CostOfGoodsAndServicesSold", "CostOfProductsSold"],
            "duration",
            ["usd"],
        ),
        "gross_profit": choose_value(facts, contexts, report_date, ["GrossProfit"], "duration", ["usd"]),
        "operating_expenses": choose_value(
            facts, contexts, report_date,
            ["OperatingExpenses", "CostsAndExpenses", "OperatingCostsAndExpenses"],
            "duration",
            ["usd"],
        ),
        "operating_income": choose_value(facts, contexts, report_date, ["OperatingIncomeLoss"], "duration", ["usd"]),
        "non_operating_income_expense_net": choose_value(
            facts, contexts, report_date,
            ["NonoperatingIncomeExpense", "OtherNonoperatingIncomeExpense"],
            "duration",
            ["usd"],
        ),
        "other_income": choose_value(
            facts, contexts, report_date,
            ["OtherIncomeExpenseNet", "OtherNonoperatingIncomeExpense"],
            "duration",
            ["usd"],
        ),
        "income_before_income_taxes": choose_value(
            facts, contexts, report_date,
            [
                "IncomeBeforeTax",
                "IncomeBeforeTaxExpenseBenefit",
                "IncomeLossFromContinuingOperationsBeforeIncomeTaxes",
                "IncomeLossFromContinuingOperationsBeforeIncomeTaxesExtraordinaryItemsNoncontrollingInterest",
                "IncomeLossFromContinuingOperationsBeforeIncomeTaxesMinorityInterestAndIncomeLossFromEquityMethodInvestments",
                "IncomeLossFromContinuingOperationsAfterNoncontrollingInterestBeforeIncomeTaxes",
                "IncomeLossFromContinuingOperationsIncludingNoncontrollingInterestBeforeIncomeTaxesExtraordinaryItems",
            ],
            "duration",
            ["usd"],
        ),
        "provision_for_income_taxes": choose_value(
            facts, contexts, report_date,
            ["IncomeTaxExpenseBenefit", "ProvisionForIncomeTaxes"],
            "duration",
            ["usd"],
        ),
        "net_income": choose_value(
            facts, contexts, report_date,
            ["NetIncomeLoss", "ProfitLoss", "NetIncomeLossAvailableToCommonStockholdersBasic"],
            "duration",
            ["usd"],
        ),
        "net_income_per_share_basic": choose_value(
            facts, contexts, report_date,
            ["EarningsPerShareBasic"],
            "duration",
            ["usd/shares", "pure"],
        ),
        "net_income_per_share_diluted": choose_value(
            facts, contexts, report_date,
            ["EarningsPerShareDiluted"],
            "duration",
            ["usd/shares", "pure"],
        ),
        "shares_outstanding_basic": choose_value(
            facts, contexts, report_date,
            ["WeightedAverageNumberOfSharesOutstandingBasic"],
            "duration",
            ["shares"],
        ),
        "shares_outstanding_diluted": choose_value(
            facts, contexts, report_date,
            ["WeightedAverageNumberOfDilutedSharesOutstanding"],
            "duration",
            ["shares"],
        ),
        "anomaly_flag": 0,
        "anomaly_type": None,
        "anomaly_score": None,
        "anomaly_detail": None,
    }


def build_bs_row(
    symbol: str,
    cik: str,
    filing: dict,
    report_date: str,
    facts: List[Dict],
    contexts: Dict[str, Dict[str, Optional[object]]],
) -> Dict:
    term_debt_current_tags = [
        "LongTermDebtCurrent",
        "CurrentPortionOfLongTermDebt",
        "CurrentMaturitiesOfLongTermDebt",
        "CurrentPortionOfLongTermDebtAndCapitalLeaseObligations",
        "LongTermDebtAndCapitalLeaseObligationsCurrent",
        "DebtCurrent",
        "ShortTermDebt",
        "ShortTermBorrowings",
    ]
    term_debt_noncurrent_strict_tags = [
        "LongTermDebtNoncurrent",
        "LongTermDebtAndCapitalLeaseObligationsNoncurrent",
        "NotesPayableNoncurrent",
        "BondsPayableNoncurrent",
        "SeniorNotesNoncurrent",
    ]
    term_debt_noncurrent_fallback_tags = [
        "LongTermDebt",
        "LongTermDebtAndCapitalLeaseObligations",
        "NotesPayableLongTerm",
        "BondsPayable",
        "SeniorNotes",
    ]
    commercial_paper_current_tags = [
        "CommercialPaper",
        "CommercialPaperCurrent",
    ]

    term_debt_current, term_debt_current_source_tag = choose_value_with_source(
        facts, contexts, report_date, term_debt_current_tags, "instant", ["usd"]
    )
    term_debt_noncurrent, term_debt_noncurrent_source_tag = choose_value_with_source(
        facts, contexts, report_date, term_debt_noncurrent_strict_tags, "instant", ["usd"]
    )
    term_debt_noncurrent_proxy_from_generic_long_term_tag = 0
    if term_debt_noncurrent is None:
        term_debt_noncurrent, term_debt_noncurrent_source_tag = choose_value_with_source(
            facts, contexts, report_date, term_debt_noncurrent_fallback_tags, "instant", ["usd"]
        )
        if term_debt_noncurrent is not None:
            term_debt_noncurrent_proxy_from_generic_long_term_tag = 1

    commercial_paper_current, commercial_paper_current_source_tag = choose_value_with_source(
        facts, contexts, report_date, commercial_paper_current_tags, "instant", ["usd"]
    )

    term_debt_total: Optional[float] = None
    term_debt_total_imputed_current_zero = 0
    term_debt_total_imputed_noncurrent_zero = 0
    if term_debt_current is not None or term_debt_noncurrent is not None:
        if term_debt_current is None:
            term_debt_total_imputed_current_zero = 1
        if term_debt_noncurrent is None:
            term_debt_total_imputed_noncurrent_zero = 1
        term_debt_total = (term_debt_current or 0.0) + (term_debt_noncurrent or 0.0)

    return {
        "country": "US",
        "symbol": symbol,
        "company_id": cik,
        "form_type": filing["form_type"],
        "fiscal_year_end_date": report_date,
        "filing_date": filing["filing_date"],
        "fiscal_year": filing["fiscal_year"],
        "fiscal_period": filing["fiscal_period"],
        "source_filing_id": filing["accession_number"],
        "is_amendment": 0,
        "accounting_standard": "US_GAAP",
        "taxonomy_version": None,
        "statement_scope": "CONSOLIDATED",
        "currency": "USD",
        "unit_scale": 1,
        "total_assets": choose_value(facts, contexts, report_date, ["Assets"], "instant", ["usd"]),
        "cash_and_cash_equivalents": choose_value(
            facts, contexts, report_date,
            ["CashAndCashEquivalentsAtCarryingValue", "CashCashEquivalentsRestrictedCashAndRestrictedCashEquivalents"],
            "instant",
            ["usd"],
        ),
        "marketable_securities_current": choose_value(
            facts,
            contexts,
            report_date,
            [
                "MarketableSecuritiesCurrent",
                "AvailableForSaleSecuritiesCurrent",
                "ShortTermInvestments",
            ],
            "instant",
            ["usd"],
        ),
        "marketable_securities_noncurrent": choose_value(
            facts,
            contexts,
            report_date,
            [
                "MarketableSecuritiesNoncurrent",
                "AvailableForSaleSecuritiesNoncurrent",
                "LongTermMarketableSecurities",
                "AvailableForSaleDebtSecuritiesNoncurrent",
            ],
            "instant",
            ["usd"],
        ),
        "short_term_investments": choose_value(
            facts, contexts, report_date, ["ShortTermInvestments"], "instant", ["usd"]
        ),
        "total_cash_and_short_term_investments": choose_value(
            facts, contexts, report_date,
            ["CashCashEquivalentsAndShortTermInvestments", "CashAndCashEquivalentsAndMarketableSecurities"],
            "instant",
            ["usd"],
        ),
        "inventories": choose_value(facts, contexts, report_date, ["InventoryNet"], "instant", ["usd"]),
        "accounts_receivable": choose_value(
            facts, contexts, report_date,
            ["AccountsReceivableNetCurrent", "ReceivablesNetCurrent"],
            "instant",
            ["usd"],
        ),
        "other_current_assets": choose_value(facts, contexts, report_date, ["OtherAssetsCurrent"], "instant", ["usd"]),
        "total_current_assets": choose_value(facts, contexts, report_date, ["AssetsCurrent"], "instant", ["usd"]),
        "property_plant_and_equipment_net": choose_value(
            facts, contexts, report_date, ["PropertyPlantAndEquipmentNet"], "instant", ["usd"]
        ),
        "goodwill": choose_value(facts, contexts, report_date, ["Goodwill"], "instant", ["usd"]),
        "intangible_assets": choose_value(
            facts, contexts, report_date,
            ["IntangibleAssetsNetExcludingGoodwill", "IntangibleAssetsNetIncludingGoodwill"],
            "instant",
            ["usd"],
        ),
        "other_noncurrent_assets": choose_value(facts, contexts, report_date, ["OtherAssetsNoncurrent"], "instant", ["usd"]),
        "accounts_payable": choose_value(facts, contexts, report_date, ["AccountsPayableCurrent"], "instant", ["usd"]),
        # US debt split follows deterministic current/non-current tags to avoid double counting.
        "short_term_debt": term_debt_current,
        "other_current_liabilities": choose_value(
            facts, contexts, report_date, ["OtherLiabilitiesCurrent"], "instant", ["usd"]
        ),
        "total_current_liabilities": choose_value(facts, contexts, report_date, ["LiabilitiesCurrent"], "instant", ["usd"]),
        "long_term_debt": term_debt_noncurrent,
        "term_debt_total": term_debt_total,
        "commercial_paper_current": commercial_paper_current,
        "short_term_debt_source_tag": term_debt_current_source_tag,
        "long_term_debt_source_tag": term_debt_noncurrent_source_tag,
        "commercial_paper_current_source_tag": commercial_paper_current_source_tag,
        "term_debt_total_imputed_current_zero": term_debt_total_imputed_current_zero,
        "term_debt_total_imputed_noncurrent_zero": term_debt_total_imputed_noncurrent_zero,
        "term_debt_noncurrent_proxy_from_generic_long_term_tag": term_debt_noncurrent_proxy_from_generic_long_term_tag,
        "other_noncurrent_liabilities": choose_value(
            facts, contexts, report_date, ["OtherLiabilitiesNoncurrent"], "instant", ["usd"]
        ),
        "total_liabilities": choose_value(facts, contexts, report_date, ["Liabilities"], "instant", ["usd"]),
        "total_shareholders_equity": choose_value(
            facts, contexts, report_date,
            ["StockholdersEquity", "StockholdersEquityIncludingPortionAttributableToNoncontrollingInterest"],
            "instant",
            ["usd"],
        ),
        "accumulated_other_comprehensive_income_or_loss": choose_value(
            facts, contexts, report_date,
            ["AccumulatedOtherComprehensiveIncomeLossNetOfTax"],
            "instant",
            ["usd"],
        ),
        "common_stock": choose_value(
            facts, contexts, report_date, ["CommonStockValue", "CommonStocksIncludingAdditionalPaidInCapital"], "instant", ["usd"]
        ),
        "additional_paid_in_capital": choose_value(
            facts, contexts, report_date, ["AdditionalPaidInCapital"], "instant", ["usd"]
        ),
        "retained_earnings": choose_value(
            facts, contexts, report_date, ["RetainedEarningsAccumulatedDeficit"], "instant", ["usd"]
        ),
        "noncontrolling_interests": choose_value(
            facts, contexts, report_date, ["NoncontrollingInterestInEquity"], "instant", ["usd"]
        ),
        "total_liabilities_and_shareholders_equity": choose_value(
            facts, contexts, report_date, ["LiabilitiesAndStockholdersEquity"], "instant", ["usd"]
        ),
        "anomaly_flag": 0,
        "anomaly_type": None,
        "anomaly_score": None,
        "anomaly_detail": None,
    }


def build_cf_row(
    symbol: str,
    filing: dict,
    report_date: str,
    facts: List[Dict],
    root: etree._Element,
) -> Dict:
    context_map = context_period_info(root)
    ctx_end_map: Dict[str, str] = {}
    for cid, p in context_map.items():
        if p.get("end_date"):
            ctx_end_map[cid] = p["end_date"]  # type: ignore[index]
        elif p.get("instant"):
            ctx_end_map[cid] = p["instant"]  # type: ignore[index]

    cf_data = extract_cash_flow_data(facts, ctx_end_map, context_map, report_date, root)

    return {
        "symbol": symbol,
        "form_type": filing["form_type"],
        "fiscal_year_end_date": report_date,
        "filing_date": filing["filing_date"],
        "accession_number": filing["accession_number"],
        "net_income": cf_data.get("net_income"),
        "net_cash_operating": cf_data.get("net_cash_operating"),
        "net_cash_investing": cf_data.get("net_cash_investing"),
        "net_cash_financing": cf_data.get("net_cash_financing"),
        "effect_of_exchange_rates_on_cash": cf_data.get("effect_of_exchange_rates_on_cash"),
        "net_change_in_cash": cf_data.get("net_change_in_cash"),
        "cash_beginning_of_period": cf_data.get("cash_beginning_of_period"),
        "cash_end_of_period": cf_data.get("cash_end_of_period"),
        "currency": "USD",
    }


def export_three_statements(symbol: str, cik: str, out_path: Path, form_type: str = "10-K") -> Path:
    print(f"[INFO] Fetching submissions for {symbol} ({cik})...")
    sub = get_submissions_json(cik)
    recent, i = latest_annual_filing(sub, form_type=form_type)

    filing = {
        "form_type": recent["form"][i],
        "report_date": recent["reportDate"][i],
        "filing_date": recent["filingDate"][i],
        "accession_number": recent["accessionNumber"][i],
        "fiscal_year": int(recent["reportDate"][i][:4]),
        "fiscal_period": recent.get("fiscalPeriod", ["FY"] * len(recent["form"]))[i] or "FY",
    }

    report_date = filing["report_date"]
    print(
        "[INFO] Using filing:",
        filing["accession_number"],
        filing["form_type"],
        f"reportDate={report_date}",
        f"filingDate={filing['filing_date']}",
    )

    index_json = download_filing_index(cik, filing["accession_number"])
    instance_file = pick_instance_file(index_json)
    print(f"[INFO] XBRL instance: {instance_file}")

    xml_bytes = download_instance_xml(cik, filing["accession_number"], instance_file)
    facts, root = parse_xbrl_facts(xml_bytes)
    contexts = parse_contexts(root)
    print(f"[INFO] Parsed {len(facts)} facts and {len(contexts)} contexts")

    is_row = build_is_row(symbol, cik, filing, report_date, facts, contexts)
    bs_row = build_bs_row(symbol, cik, filing, report_date, facts, contexts)
    cf_row = build_cf_row(symbol, filing, report_date, facts, root)
    us_candidates = build_us_candidate_rows(
        report_date,
        facts,
        contexts,
        is_row,
        bs_row,
        cf_row,
        top_k=10,
    )

    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        pd.DataFrame([is_row]).to_excel(writer, index=False, sheet_name="US_FIN_IS")
        pd.DataFrame([bs_row]).to_excel(writer, index=False, sheet_name="US_FIN_BS")
        pd.DataFrame([cf_row]).to_excel(writer, index=False, sheet_name="US_FIN_CF")
        pd.DataFrame(us_candidates).to_excel(writer, index=False, sheet_name="RAW_US_FIELD_CANDIDATES")

    print(f"[OK] Excel exported: {out_path}")
    return out_path


def main() -> None:
    parser = argparse.ArgumentParser(description="Export US IS/BS/CF to one Excel workbook.")
    parser.add_argument("--symbol", required=True, help="Ticker symbol, e.g. AAPL")
    parser.add_argument("--cik", required=True, help="10-digit CIK, e.g. 0000320193")
    parser.add_argument("--out", default="aapl_3statements.xlsx", help="Output Excel path")
    parser.add_argument("--form-type", default="10-K", help="Form type, default 10-K")
    args = parser.parse_args()

    export_three_statements(
        symbol=args.symbol.upper(),
        cik=args.cik,
        out_path=Path(args.out),
        form_type=args.form_type,
    )


if __name__ == "__main__":
    main()
