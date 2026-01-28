---
name: ai-insights
description: AI-powered narrative generation and intelligent parameter suggestions for royalty valuations. Analyzes historical trends, generates risk factors and opportunities, and creates investment memo narratives. Can optionally use Claude API for enhanced insights. Use when generating analysis reports, suggesting valuation parameters, or creating investment narratives.
version: "1.0.0"
author: Royalty Valuation Tool
triggers:
  - ai analysis
  - narrative
  - insights
  - risk factors
  - opportunities
  - investment memo
  - parameter suggestion
dependencies:
  - anthropic (optional)
---

# AI Insights Skill

## Overview

This skill provides intelligent analysis and narrative generation for music royalty valuations. It works in two modes:

1. **Statistical Mode**: Uses historical data analysis to generate insights (no API required)
2. **AI-Enhanced Mode**: Uses Claude API for richer narratives (requires API key)

## Capabilities

### Trend Analysis
- Calculate CAGR (Compound Annual Growth Rate)
- Measure historical volatility
- Detect trend direction (growing/stable/declining)
- Identify momentum (accelerating/steady/decelerating)

### Parameter Suggestion
- Suggest growth rates based on historical CAGR
- Recommend discount rates based on volatility
- Propose terminal growth rates based on trend

### Risk & Opportunity Analysis
- Generate contextual risk factors
- Identify potential opportunities
- Assess confidence levels

### Narrative Generation
- Create executive summaries
- Generate investment memo content
- Explain valuation methodology

## Parameters

| Parameter | Type | Description | Default |
|-----------|------|-------------|---------|
| yearly_data | dict | Year-to-earnings mapping | Required |
| base_year_cf | float | Starting cash flow | Required |
| trend_analysis | dict | Pre-computed trend analysis | None |
| benchmark_comparison | dict | Decay curve comparison | None |
| monte_carlo | dict | Monte Carlo results | None |
| oauth_token | str | OAuth token (optional) | None |

## Output

Returns a dictionary containing:
- Suggested parameters (growth, discount, terminal rates)
- Confidence score
- Risk factors list
- Opportunities list
- AI narrative summary
- Scenario probabilities
