#!/usr/bin/env python3
"""
DCF Valuation Skill Executor
Implements Discounted Cash Flow analysis for music royalty assets.
"""

from dataclasses import dataclass
from typing import Dict, Optional


@dataclass
class DCFResult:
    """Results from DCF valuation analysis."""
    bear_value: float
    base_value: float
    bull_value: float
    weighted_value: float
    pv_cash_flows: float
    pv_terminal: float
    terminal_value: float
    cash_flows: list
    discount_factors: list
    ev_multiple: float
    payback_years: float


def calculate_dcf(
    base_year_cf: float,
    growth_rate_1_3: float,
    growth_rate_4_5: float,
    discount_rate: float,
    terminal_growth: float
) -> dict:
    """Calculate DCF for a single scenario."""

    # Project cash flows
    cf = [base_year_cf]
    for year in range(1, 6):
        if year <= 3:
            cf.append(cf[-1] * (1 + growth_rate_1_3))
        else:
            cf.append(cf[-1] * (1 + growth_rate_4_5))

    # Calculate discount factors
    discount_factors = [1 / (1 + discount_rate) ** i for i in range(6)]

    # Present value of cash flows (years 1-5)
    pv_cfs = sum(cf[i] * discount_factors[i] for i in range(1, 6))

    # Terminal value using Gordon Growth Model
    year_5_cf = cf[5]
    terminal_cf = year_5_cf * (1 + terminal_growth)

    if discount_rate > terminal_growth:
        terminal_value = terminal_cf / (discount_rate - terminal_growth)
    else:
        # Fallback: use multiple
        terminal_value = year_5_cf * 10

    # Present value of terminal
    pv_terminal = terminal_value * discount_factors[5]

    # Enterprise value
    enterprise_value = pv_cfs + pv_terminal

    return {
        'enterprise_value': enterprise_value,
        'pv_cash_flows': pv_cfs,
        'pv_terminal': pv_terminal,
        'terminal_value': terminal_value,
        'cash_flows': cf,
        'discount_factors': discount_factors
    }


def execute(
    base_year_cf: float,
    growth_rate_1_3: float = 0.05,
    growth_rate_4_5: float = 0.03,
    discount_rate: float = 0.12,
    terminal_growth: float = -0.05,
    scenario_weights: Optional[Dict[str, float]] = None
) -> DCFResult:
    """
    Execute DCF valuation with three-scenario analysis.

    Args:
        base_year_cf: Starting annual cash flow
        growth_rate_1_3: Growth rate for years 1-3
        growth_rate_4_5: Growth rate for years 4-5
        discount_rate: Required rate of return
        terminal_growth: Perpetual growth rate
        scenario_weights: Bear/Base/Bull probability weights

    Returns:
        DCFResult with valuation metrics
    """

    if scenario_weights is None:
        scenario_weights = {'bear': 0.25, 'base': 0.50, 'bull': 0.25}

    # Bear scenario: -10% CF, lower growth, higher discount
    bear = calculate_dcf(
        base_year_cf=base_year_cf * 0.9,
        growth_rate_1_3=growth_rate_1_3 - 0.02,
        growth_rate_4_5=growth_rate_4_5 - 0.01,
        discount_rate=discount_rate + 0.02,
        terminal_growth=terminal_growth - 0.02
    )

    # Base scenario
    base = calculate_dcf(
        base_year_cf=base_year_cf,
        growth_rate_1_3=growth_rate_1_3,
        growth_rate_4_5=growth_rate_4_5,
        discount_rate=discount_rate,
        terminal_growth=terminal_growth
    )

    # Bull scenario: +10% CF, higher growth
    bull = calculate_dcf(
        base_year_cf=base_year_cf * 1.1,
        growth_rate_1_3=growth_rate_1_3 + 0.03,
        growth_rate_4_5=growth_rate_4_5 + 0.02,
        discount_rate=discount_rate,
        terminal_growth=terminal_growth + 0.02
    )

    # Weighted average
    weighted_value = (
        bear['enterprise_value'] * scenario_weights['bear'] +
        base['enterprise_value'] * scenario_weights['base'] +
        bull['enterprise_value'] * scenario_weights['bull']
    )

    return DCFResult(
        bear_value=bear['enterprise_value'],
        base_value=base['enterprise_value'],
        bull_value=bull['enterprise_value'],
        weighted_value=weighted_value,
        pv_cash_flows=base['pv_cash_flows'],
        pv_terminal=base['pv_terminal'],
        terminal_value=base['terminal_value'],
        cash_flows=base['cash_flows'],
        discount_factors=base['discount_factors'],
        ev_multiple=weighted_value / base_year_cf if base_year_cf > 0 else 0,
        payback_years=weighted_value / base_year_cf if base_year_cf > 0 else 0
    )
