#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Display Cash Flow Statement Results

Usage:
    python show_cash_flow.py test_aapl.csv
    python show_cash_flow.py test_wmt.csv
"""

import sys
import pandas as pd


def format_millions(value):
    """Format value in millions/billions USD"""
    if pd.isna(value):
        return "N/A"
    if value == 0:
        return "$0.00M"
    # Convert to millions
    millions = value / 1_000_000
    if abs(millions) >= 1000:
        billions = millions / 1000
        return f"${billions:.2f}B"
    return f"${millions:.2f}M"


def format_usd(value):
    """Format value as USD"""
    if pd.isna(value):
        return "N/A"
    return f"${value:,.0f}"


def show_cash_flow(csv_file: str):
    """Display cash flow statement from CSV file"""
    df = pd.read_csv(csv_file)
    
    print("=" * 80)
    print(f"{df['symbol'].iloc[0]} - Cash Flow Statement")
    print(f"Extracted from SEC XBRL Data (FY {df['fiscal_year_end_date'].iloc[0]})")
    print("=" * 80)
    
    print(f"\n📋 Basic Information:")
    print(f"  Symbol: {df['symbol'].iloc[0]}")
    print(f"  Form Type: {df['form_type'].iloc[0]}")
    print(f"  Fiscal Year End: {df['fiscal_year_end_date'].iloc[0]}")
    print(f"  Filing Date: {df['filing_date'].iloc[0]}")
    print(f"  Accession Number: {df['accession_number'].iloc[0]}")
    print(f"  Currency: {df['currency'].iloc[0]}")
    
    print(f"\n💰 Cash Flow Statement Data:")
    print(f"  {'-' * 70}")
    print(f"  {'Item':<45} {'Amount (Millions USD)':<25} {'Raw Value (USD)':<25}")
    print(f"  {'-' * 70}")
    print(f"  Consolidated Net Income:                 {format_millions(df['net_income'].iloc[0]):<25} {format_usd(df['net_income'].iloc[0])}")
    print(f"  Net Cash from Operating Activities:      {format_millions(df['net_cash_operating'].iloc[0]):<25} {format_usd(df['net_cash_operating'].iloc[0])}")
    print(f"  Net Cash from Investing Activities:      {format_millions(df['net_cash_investing'].iloc[0]):<25} {format_usd(df['net_cash_investing'].iloc[0])}")
    print(f"  Net Cash from Financing Activities:      {format_millions(df['net_cash_financing'].iloc[0]):<25} {format_usd(df['net_cash_financing'].iloc[0])}")
    print(f"  Effect of Exchange Rates on Cash:        {format_millions(df['effect_of_exchange_rates_on_cash'].iloc[0]):<25} {format_millions(df['effect_of_exchange_rates_on_cash'].iloc[0])}")
    print(f"  Net Change in Cash:                      {format_millions(df['net_change_in_cash'].iloc[0]):<25} {format_usd(df['net_change_in_cash'].iloc[0])}")
    print(f"  Cash Beginning of Period:                {format_millions(df['cash_beginning_of_period'].iloc[0]):<25} {format_millions(df['cash_beginning_of_period'].iloc[0])}")
    print(f"  Cash End of Period:                       {format_millions(df['cash_end_of_period'].iloc[0]):<25} {format_usd(df['cash_end_of_period'].iloc[0])}")
    
    print(f"\n✅ Data Validation:")
    net_income = df['net_income'].iloc[0]
    net_cash_operating = df['net_cash_operating'].iloc[0]
    net_cash_investing = df['net_cash_investing'].iloc[0]
    net_cash_financing = df['net_cash_financing'].iloc[0]
    net_change = df['net_change_in_cash'].iloc[0]
    cash_end = df['cash_end_of_period'].iloc[0]
    
    print(f"  Consolidated Net Income: {format_millions(net_income)}")
    print(f"  Net Cash from Operating: {format_millions(net_cash_operating)}")
    print(f"  Net Cash from Investing: {format_millions(net_cash_investing)}")
    print(f"  Net Cash from Financing: {format_millions(net_cash_financing)}")
    print(f"  Net Change in Cash: {format_millions(net_change)}")
    print(f"  Cash End of Period: {format_millions(cash_end)}")
    
    # Validation: Operating + Investing + Financing ≈ Net Change
    calculated_change = net_cash_operating + net_cash_investing + net_cash_financing
    if pd.notna(df['effect_of_exchange_rates_on_cash'].iloc[0]):
        calculated_change += df['effect_of_exchange_rates_on_cash'].iloc[0]
    
    print(f"\n  Validation Calculation:")
    print(f"    Operating + Investing + Financing = {format_millions(calculated_change)}")
    print(f"    Actual Net Change in Cash = {format_millions(net_change)}")
    diff = abs(calculated_change - net_change)
    if diff < abs(net_change) * 0.01:  # Allow 1% error
        print(f"    ✅ Data consistency check passed (Difference: {format_millions(diff)})")
    else:
        print(f"    ⚠️  Data may have discrepancies (Difference: {format_millions(diff)}, possibly due to FX effects or other factors)")
    
    print("\n" + "=" * 80)
    print("📊 Key Metrics:")
    if pd.notna(net_income) and net_income != 0:
        print(f"  Cash Flow Quality: Operating CF / Consolidated Net Income = {net_cash_operating/net_income:.2f}x")
    if pd.notna(net_cash_operating) and pd.notna(net_cash_investing):
        print(f"  Free Cash Flow: Operating CF + Investing CF = {format_millions(net_cash_operating + net_cash_investing)}")
    print("=" * 80)


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python show_cash_flow.py <csv_file>")
        sys.exit(1)
    
    csv_file = sys.argv[1]
    show_cash_flow(csv_file)
