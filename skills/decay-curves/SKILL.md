---
name: decay-curves
description: Genre-based decay curve benchmarking for music royalty catalogs. Compares actual catalog performance against industry benchmarks by genre (Pop, Rock, Hip-Hop, Country, Electronic, Classical). Use when analyzing catalog health, comparing to industry standards, or selecting appropriate growth assumptions.
version: "1.0.0"
author: Royalty Valuation Tool
triggers:
  - decay
  - benchmark
  - genre
  - industry comparison
  - catalog performance
  - decline rate
dependencies: []
---

# Decay Curve Benchmarking Skill

## Overview

This skill compares royalty catalog decay rates against industry benchmarks by music genre. Different genres exhibit different decay patterns based on listener behavior and catalog characteristics.

## Industry Benchmarks

### Genre-Specific Decay Rates

| Genre | Year 1 | Year 2 | Year 3 | Year 4 | Year 5+ |
|-------|--------|--------|--------|--------|---------|
| Pop | -15% | -12% | -10% | -8% | -5% |
| Rock | -8% | -6% | -5% | -4% | -3% |
| Hip-Hop | -20% | -15% | -12% | -8% | -5% |
| Country | -10% | -8% | -6% | -5% | -3% |
| Electronic | -18% | -14% | -10% | -7% | -4% |
| Classical | -2% | -2% | -2% | -1% | -1% |
| Mixed | -12% | -10% | -8% | -6% | -4% |

### Interpretation

- **Pop**: Fast initial decay, stabilizes after ~3 years
- **Rock**: Evergreen performance, loyal fanbase
- **Hip-Hop**: Fastest decay but streaming keeps catalogs active
- **Country**: Steady, loyal listeners
- **Electronic**: Variable; sync licensing can boost older tracks
- **Classical**: Near-perpetual assets, minimal decay

## Parameters

| Parameter | Type | Description | Default |
|-----------|------|-------------|---------|
| yearly_data | dict | Year-to-earnings mapping | Required |
| genre | str | Music genre for benchmark | "mixed" |

## Output

Returns a dictionary containing:
- Year-by-year comparison (actual vs benchmark)
- Overall assessment (above/at/below benchmark)
- Average decay rates
- Variance from benchmark
