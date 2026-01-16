#!/usr/bin/env python3
"""
Sensitivity Analysis Skill Executor
Generates sensitivity matrices for valuation stress testing.
"""

from typing import Dict, Any, List
import numpy as np


def calculate_ev(
    base_year_cf: float,
    growth_1_3: float,
    growth_4_5: float,
    discount_rate: float,
    terminal_growth: float
) -> float:
    """Calculate enterprise value for given parameters."""

    # Project cash flows
    cf = [base_year_cf]
    for year in range(1, 6):
        if year <= 3:
            cf.append(cf[-1] * (1 + growth_1_3))
        else:
            cf.append(cf[-1] * (1 + growth_4_5))

    # PV of cash flows
    pv_cf = sum(cf[i] / (1 + discount_rate) ** i for i in range(1, 6))

    # Terminal value
    if discount_rate > terminal_growth:
        terminal_cf = cf[5] * (1 + terminal_growth)
        terminal_value = terminal_cf / (discount_rate - terminal_growth)
        pv_terminal = terminal_value / (1 + discount_rate) ** 5
    else:
        pv_terminal = cf[5] * 10  # Fallback

    return pv_cf + pv_terminal


def execute(
    base_year_cf: float,
    base_growth: float = 0.05,
    base_discount: float = 0.12,
    base_terminal: float = -0.05,
    growth_4_5: float = 0.03
) -> Dict[str, Any]:
    """
    Generate sensitivity analysis matrices.

    Args:
        base_year_cf: Starting cash flow
        base_growth: Base growth rate for years 1-3
        base_discount: Base discount rate
        base_terminal: Base terminal growth rate
        growth_4_5: Growth rate for years 4-5

    Returns:
        Dictionary with sensitivity matrices and analysis
    """

    # Define ranges
    discount_rates = [0.08, 0.10, 0.12, 0.14, 0.16, 0.18]
    growth_rates = [0.00, 0.02, 0.04, 0.06, 0.08, 0.10, 0.12]
    terminal_rates = [-0.10, -0.07, -0.05, -0.03, 0.00, 0.02, 0.03]

    # Sensitivity 1: Discount Rate vs Growth Rate (Years 1-3)
    growth_matrix = []
    for dr in discount_rates:
        row = []
        for gr in growth_rates:
            ev = calculate_ev(
                base_year_cf=base_year_cf,
                growth_1_3=gr,
                growth_4_5=growth_4_5,
                discount_rate=dr,
                terminal_growth=base_terminal
            )
            row.append(ev)
        growth_matrix.append(row)

    # Sensitivity 2: Discount Rate vs Terminal Growth Rate
    terminal_matrix = []
    for dr in discount_rates:
        row = []
        for tr in terminal_rates:
            ev = calculate_ev(
                base_year_cf=base_year_cf,
                growth_1_3=base_growth,
                growth_4_5=growth_4_5,
                discount_rate=dr,
                terminal_growth=tr
            )
            row.append(ev)
        terminal_matrix.append(row)

    # Base case value
    base_value = calculate_ev(
        base_year_cf=base_year_cf,
        growth_1_3=base_growth,
        growth_4_5=growth_4_5,
        discount_rate=base_discount,
        terminal_growth=base_terminal
    )

    # Tornado analysis: Calculate impact of each parameter
    tornado_data = []

    # Growth rate impact
    low_growth_ev = calculate_ev(base_year_cf, 0.00, growth_4_5, base_discount, base_terminal)
    high_growth_ev = calculate_ev(base_year_cf, 0.10, growth_4_5, base_discount, base_terminal)
    tornado_data.append({
        "parameter": "Growth Rate (Yr 1-3)",
        "low_value": low_growth_ev,
        "high_value": high_growth_ev,
        "low_change": (low_growth_ev - base_value) / base_value,
        "high_change": (high_growth_ev - base_value) / base_value,
        "range": high_growth_ev - low_growth_ev
    })

    # Discount rate impact
    low_discount_ev = calculate_ev(base_year_cf, base_growth, growth_4_5, 0.08, base_terminal)
    high_discount_ev = calculate_ev(base_year_cf, base_growth, growth_4_5, 0.16, base_terminal)
    tornado_data.append({
        "parameter": "Discount Rate",
        "low_value": low_discount_ev,
        "high_value": high_discount_ev,
        "low_change": (low_discount_ev - base_value) / base_value,
        "high_change": (high_discount_ev - base_value) / base_value,
        "range": abs(low_discount_ev - high_discount_ev)
    })

    # Terminal growth impact
    low_terminal_ev = calculate_ev(base_year_cf, base_growth, growth_4_5, base_discount, -0.10)
    high_terminal_ev = calculate_ev(base_year_cf, base_growth, growth_4_5, base_discount, 0.02)
    tornado_data.append({
        "parameter": "Terminal Growth Rate",
        "low_value": low_terminal_ev,
        "high_value": high_terminal_ev,
        "low_change": (low_terminal_ev - base_value) / base_value,
        "high_change": (high_terminal_ev - base_value) / base_value,
        "range": abs(high_terminal_ev - low_terminal_ev)
    })

    # Base year CF impact
    low_cf_ev = calculate_ev(base_year_cf * 0.8, base_growth, growth_4_5, base_discount, base_terminal)
    high_cf_ev = calculate_ev(base_year_cf * 1.2, base_growth, growth_4_5, base_discount, base_terminal)
    tornado_data.append({
        "parameter": "Base Year Cash Flow",
        "low_value": low_cf_ev,
        "high_value": high_cf_ev,
        "low_change": (low_cf_ev - base_value) / base_value,
        "high_change": (high_cf_ev - base_value) / base_value,
        "range": high_cf_ev - low_cf_ev
    })

    # Sort tornado by impact range
    tornado_data.sort(key=lambda x: x["range"], reverse=True)

    return {
        "growth_sensitivity": growth_matrix,
        "terminal_sensitivity": terminal_matrix,
        "discount_rates": discount_rates,
        "growth_rates": growth_rates,
        "terminal_rates": terminal_rates,
        "base_value": base_value,
        "tornado_analysis": tornado_data,
        "base_parameters": {
            "base_year_cf": base_year_cf,
            "growth_1_3": base_growth,
            "growth_4_5": growth_4_5,
            "discount_rate": base_discount,
            "terminal_growth": base_terminal
        }
    }
