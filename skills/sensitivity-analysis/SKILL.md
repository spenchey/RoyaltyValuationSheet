---
name: sensitivity-analysis
description: Sensitivity analysis tables for royalty valuations. Generates two-way sensitivity matrices showing how enterprise value changes with variations in discount rate, growth rate, and terminal growth. Use when analyzing parameter impact, stress testing assumptions, or presenting valuation ranges.
version: "1.0.0"
author: Royalty Valuation Tool
triggers:
  - sensitivity
  - what-if
  - stress test
  - parameter impact
  - sensitivity table
  - tornado chart
dependencies:
  - numpy
---

# Sensitivity Analysis Skill

## Overview

This skill generates sensitivity analysis matrices that show how enterprise value varies with changes to key assumptions. Useful for stress testing and understanding parameter impact.

## Analysis Types

### 1. Discount Rate vs Growth Rate (Years 1-3)
Shows how valuation changes when varying both:
- Discount rates: 8% to 18%
- Growth rates: 0% to 12%

### 2. Discount Rate vs Terminal Growth Rate
Shows how valuation changes when varying:
- Discount rates: 8% to 18%
- Terminal growth: -10% to +3%

### 3. Tornado Analysis
Ranks parameters by impact on valuation:
- Identifies most sensitive assumptions
- Shows upside/downside for each parameter

## Parameters

| Parameter | Type | Description | Default |
|-----------|------|-------------|---------|
| base_year_cf | float | Starting cash flow | Required |
| base_growth | float | Base growth rate | 0.05 |
| base_discount | float | Base discount rate | 0.12 |
| base_terminal | float | Base terminal growth | -0.05 |
| growth_4_5 | float | Growth rate years 4-5 | 0.03 |

## Output

Returns a dictionary containing:
- growth_sensitivity: 2D matrix (discount x growth)
- terminal_sensitivity: 2D matrix (discount x terminal)
- tornado_analysis: Ranked parameter impacts
- discount_rates: Row labels
- growth_rates: Column labels for growth matrix
- terminal_rates: Column labels for terminal matrix
