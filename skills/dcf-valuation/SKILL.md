---
name: dcf-valuation
description: Discounted Cash Flow (DCF) valuation model for music royalty assets. Calculates enterprise value using multi-phase growth projections, discount rates, and terminal value via Gordon Growth Model. Use when performing royalty valuations, creating financial projections, or analyzing investment opportunities.
version: "1.0.0"
author: Royalty Valuation Tool
triggers:
  - dcf
  - valuation
  - enterprise value
  - discount rate
  - terminal value
  - cash flow projection
dependencies:
  - numpy
---

# DCF Valuation Skill

## Overview

This skill performs Discounted Cash Flow analysis specifically designed for music royalty assets. It implements a two-phase growth model with scenario analysis.

## Methodology

### Two-Phase Growth Model

1. **Phase 1 (Years 1-3)**: Near-term growth rate based on recent performance
2. **Phase 2 (Years 4-5)**: Mature growth rate as catalog stabilizes

### Terminal Value

Uses Gordon Growth Model: `TV = CF_5 * (1 + g) / (r - g)`

Where:
- `CF_5` = Year 5 Cash Flow
- `g` = Terminal growth rate (typically negative for music catalogs)
- `r` = Discount rate

### Scenario Analysis

Three scenarios are modeled:
- **Bear**: Conservative assumptions (-10% base CF, lower growth, higher discount)
- **Base**: Expected case using input assumptions
- **Bull**: Optimistic assumptions (+10% base CF, higher growth)

## Parameters

| Parameter | Type | Description | Default |
|-----------|------|-------------|---------|
| base_year_cf | float | Starting annual cash flow | Required |
| growth_rate_1_3 | float | Growth rate for years 1-3 | 0.05 |
| growth_rate_4_5 | float | Growth rate for years 4-5 | 0.03 |
| discount_rate | float | Required rate of return | 0.12 |
| terminal_growth | float | Perpetual growth rate | -0.05 |
| scenario_weights | dict | Bear/Base/Bull weights | {25, 50, 25} |

## Output

Returns a `DCFResult` object containing:
- Enterprise value for each scenario
- Weighted average valuation
- Present value breakdown (cash flows vs terminal)
- Sensitivity metrics
