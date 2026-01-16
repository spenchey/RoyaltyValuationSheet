#!/usr/bin/env python3
"""
Decay Curve Benchmarking Skill Executor
Compares catalog performance against industry genre benchmarks.
"""

from typing import Dict, Any, List
import statistics


# Industry benchmark decay curves by genre
GENRE_BENCHMARKS = {
    "pop": {
        "year_1": -0.15,
        "year_2": -0.12,
        "year_3": -0.10,
        "year_4": -0.08,
        "year_5_plus": -0.05,
        "description": "Pop catalogs typically see faster initial decay but stabilize"
    },
    "rock": {
        "year_1": -0.08,
        "year_2": -0.06,
        "year_3": -0.05,
        "year_4": -0.04,
        "year_5_plus": -0.03,
        "description": "Rock catalogs tend to have more stable, evergreen performance"
    },
    "hip_hop": {
        "year_1": -0.20,
        "year_2": -0.15,
        "year_3": -0.12,
        "year_4": -0.08,
        "year_5_plus": -0.05,
        "description": "Hip-hop sees faster decay but streaming keeps catalogs active"
    },
    "country": {
        "year_1": -0.10,
        "year_2": -0.08,
        "year_3": -0.06,
        "year_4": -0.05,
        "year_5_plus": -0.03,
        "description": "Country music has loyal fanbase with steady catalog performance"
    },
    "electronic": {
        "year_1": -0.18,
        "year_2": -0.14,
        "year_3": -0.10,
        "year_4": -0.07,
        "year_5_plus": -0.04,
        "description": "Electronic music varies widely; sync licensing can boost older tracks"
    },
    "classical": {
        "year_1": -0.02,
        "year_2": -0.02,
        "year_3": -0.02,
        "year_4": -0.01,
        "year_5_plus": -0.01,
        "description": "Classical recordings are extremely stable, near-perpetual assets"
    },
    "mixed": {
        "year_1": -0.12,
        "year_2": -0.10,
        "year_3": -0.08,
        "year_4": -0.06,
        "year_5_plus": -0.04,
        "description": "Diversified catalog with blended decay characteristics"
    }
}


def execute(
    yearly_data: Dict[int, float],
    genre: str = "mixed"
) -> Dict[str, Any]:
    """
    Compare catalog decay curve against industry benchmarks.

    Args:
        yearly_data: Dictionary mapping years to earnings
        genre: Music genre for benchmark selection

    Returns:
        Dictionary with benchmark comparison results
    """

    benchmark = GENRE_BENCHMARKS.get(genre.lower(), GENRE_BENCHMARKS["mixed"])

    if len(yearly_data) < 2:
        return {
            "comparison": "insufficient_data",
            "genre": genre,
            "benchmark": benchmark,
            "message": "Need at least 2 years of data for comparison"
        }

    years = sorted(yearly_data.keys())
    values = [yearly_data[y] for y in years]

    # Calculate actual year-over-year decay rates
    actual_decays = []
    for i in range(1, len(values)):
        if values[i - 1] > 0:
            decay = (values[i] - values[i - 1]) / values[i - 1]
            actual_decays.append(decay)

    # Get benchmark rates
    benchmark_rates = [
        benchmark["year_1"],
        benchmark["year_2"],
        benchmark["year_3"],
        benchmark["year_4"],
        benchmark["year_5_plus"]
    ]

    # Year-by-year comparison
    comparisons: List[Dict[str, Any]] = []
    for i, actual in enumerate(actual_decays[:5]):
        expected = benchmark_rates[min(i, len(benchmark_rates) - 1)]
        diff = actual - expected

        if diff > 0.03:
            status = "outperforming"
        elif diff < -0.03:
            status = "underperforming"
        else:
            status = "in_line"

        comparisons.append({
            "year": i + 1,
            "actual": actual,
            "benchmark": expected,
            "difference": diff,
            "status": status
        })

    # Overall assessment
    avg_actual = statistics.mean(actual_decays) if actual_decays else 0
    avg_benchmark = statistics.mean(
        benchmark_rates[:len(actual_decays)]
    ) if actual_decays else 0

    if avg_actual > avg_benchmark + 0.02:
        overall = "above_benchmark"
        assessment = "Catalog is outperforming industry average"
    elif avg_actual < avg_benchmark - 0.02:
        overall = "below_benchmark"
        assessment = "Catalog is underperforming industry average"
    else:
        overall = "at_benchmark"
        assessment = "Catalog is performing in line with industry average"

    return {
        "genre": genre,
        "genre_description": benchmark["description"],
        "year_by_year": comparisons,
        "overall_status": overall,
        "assessment": assessment,
        "avg_actual_decay": avg_actual,
        "avg_benchmark_decay": avg_benchmark,
        "variance_from_benchmark": avg_actual - avg_benchmark,
        "all_benchmarks": GENRE_BENCHMARKS
    }
