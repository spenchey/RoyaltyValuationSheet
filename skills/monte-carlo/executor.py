#!/usr/bin/env python3
"""
Monte Carlo Simulation Skill Executor
Runs probabilistic valuation simulations for music royalty assets.
"""

from typing import Dict, Any
import numpy as np


def execute(
    base_year_cf: float,
    growth_rate: float,
    terminal_rate: float,
    discount_rate: float,
    volatility: float = 0.15,
    n_simulations: int = 1000,
    projection_years: int = 5,
    seed: int = 42
) -> Dict[str, Any]:
    """
    Run Monte Carlo simulation for valuation analysis.

    Args:
        base_year_cf: Starting annual cash flow
        growth_rate: Expected growth rate for years 1-3
        terminal_rate: Terminal growth rate
        discount_rate: Discount rate
        volatility: Historical volatility for parameter variation
        n_simulations: Number of simulation runs
        projection_years: Number of years to project
        seed: Random seed for reproducibility

    Returns:
        Dictionary with simulation statistics and distributions
    """

    np.random.seed(seed)
    valuations = []

    for _ in range(n_simulations):
        # Simulate uncertain parameters
        sim_growth_1_3 = np.random.normal(
            growth_rate,
            volatility * abs(growth_rate) + 0.02
        )
        sim_growth_4_5 = np.random.normal(
            growth_rate * 0.6,
            volatility * abs(growth_rate) + 0.015
        )
        sim_terminal = np.random.normal(terminal_rate, 0.02)
        sim_discount = np.random.normal(discount_rate, 0.01)

        # Apply bounds
        sim_growth_1_3 = np.clip(sim_growth_1_3, -0.20, 0.25)
        sim_growth_4_5 = np.clip(sim_growth_4_5, -0.15, 0.15)
        sim_terminal = np.clip(sim_terminal, -0.15, 0.03)
        sim_discount = np.clip(sim_discount, 0.06, 0.25)

        # Project cash flows
        cf = [base_year_cf]
        for year in range(1, projection_years + 1):
            if year <= 3:
                cf.append(cf[-1] * (1 + sim_growth_1_3))
            else:
                cf.append(cf[-1] * (1 + sim_growth_4_5))

        # Calculate PV of cash flows
        pv_cf = sum(
            cf[i] / (1 + sim_discount) ** i
            for i in range(1, projection_years + 1)
        )

        # Terminal value
        if sim_discount > sim_terminal:
            terminal_cf = cf[-1] * (1 + sim_terminal)
            terminal_value = terminal_cf / (sim_discount - sim_terminal)
            pv_terminal = terminal_value / (1 + sim_discount) ** projection_years
        else:
            pv_terminal = cf[-1] * 10  # Fallback multiple

        total_value = pv_cf + pv_terminal
        valuations.append(total_value)

    valuations = np.array(valuations)

    # Calculate statistics
    percentiles = {
        'p5': float(np.percentile(valuations, 5)),
        'p10': float(np.percentile(valuations, 10)),
        'p25': float(np.percentile(valuations, 25)),
        'p50': float(np.percentile(valuations, 50)),
        'p75': float(np.percentile(valuations, 75)),
        'p90': float(np.percentile(valuations, 90)),
        'p95': float(np.percentile(valuations, 95)),
    }

    return {
        'mean': float(np.mean(valuations)),
        'median': float(np.median(valuations)),
        'std_dev': float(np.std(valuations)),
        'min': float(np.min(valuations)),
        'max': float(np.max(valuations)),
        'percentiles': percentiles,
        'n_simulations': n_simulations,
        'confidence_interval_90': (percentiles['p5'], percentiles['p95']),
        'confidence_interval_80': (percentiles['p10'], percentiles['p90']),
        'confidence_interval_50': (percentiles['p25'], percentiles['p75']),
        'downside_risk': percentiles['p10'],
        'upside_potential': percentiles['p90'],
        'value_at_risk': percentiles['p50'] - percentiles['p10']
    }
