---
name: monte-carlo
description: Monte Carlo simulation for probability-weighted royalty valuations. Runs thousands of simulations with randomized parameters to generate probability distributions, confidence intervals, and risk metrics. Use when analyzing valuation uncertainty, calculating Value at Risk, or presenting probabilistic outcomes.
version: "1.0.0"
author: Royalty Valuation Tool
triggers:
  - monte carlo
  - simulation
  - probability
  - confidence interval
  - var
  - risk analysis
  - uncertainty
dependencies:
  - numpy
---

# Monte Carlo Simulation Skill

## Overview

This skill runs Monte Carlo simulations to generate probability-weighted valuations. It accounts for uncertainty in growth rates, discount rates, and terminal values.

## Methodology

### Parameter Randomization

Each simulation varies key parameters using normal distributions:
- Growth rates: Mean = input, StdDev = volatility * |input| + 0.02
- Terminal rate: Mean = input, StdDev = 0.02
- Discount rate: Mean = input, StdDev = 0.01

### Bounds

Parameters are clipped to reasonable ranges:
- Growth (Yr 1-3): -20% to +25%
- Growth (Yr 4-5): -15% to +15%
- Terminal: -15% to +3%
- Discount: 6% to 25%

## Parameters

| Parameter | Type | Description | Default |
|-----------|------|-------------|---------|
| base_year_cf | float | Starting annual cash flow | Required |
| growth_rate | float | Expected growth rate | Required |
| terminal_rate | float | Terminal growth rate | Required |
| discount_rate | float | Discount rate | Required |
| volatility | float | Historical volatility | 0.15 |
| n_simulations | int | Number of simulations | 1000 |
| projection_years | int | Years to project | 5 |

## Output

Returns a dictionary containing:
- Mean, median, std_dev of valuations
- Percentiles (5th, 10th, 25th, 50th, 75th, 90th, 95th)
- Confidence intervals (90%, 80%, 50%)
- Downside risk and upside potential
