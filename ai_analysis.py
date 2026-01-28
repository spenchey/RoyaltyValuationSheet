#!/usr/bin/env python3
"""
AI Analysis Module for Music Royalty Valuation
Integrates with Claude API for intelligent projections, Monte Carlo simulations,
and industry benchmark comparisons.
"""

import os
import json
import numpy as np
from typing import Dict, List, Tuple, Optional
from dataclasses import dataclass
import statistics

# Industry benchmark decay curves by genre (annual decay rates)
# Based on typical music catalog performance patterns
GENRE_DECAY_BENCHMARKS = {
    "pop": {
        "year_1": -0.15,  # 15% decline in year 1
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

# Industry valuation multiples by catalog type
VALUATION_MULTIPLES = {
    "premium_catalog": {"low": 15, "mid": 20, "high": 28, "description": "Iconic artists, consistent earners"},
    "established_catalog": {"low": 10, "mid": 14, "high": 18, "description": "Proven track record, stable income"},
    "developing_catalog": {"low": 6, "mid": 9, "high": 12, "description": "Growing but less predictable"},
    "legacy_catalog": {"low": 8, "mid": 11, "high": 15, "description": "Older catalog, stable but declining"},
}


@dataclass
class AIAnalysisResult:
    """Container for AI analysis results"""
    suggested_growth_rate: float
    suggested_terminal_rate: float
    suggested_discount_rate: float
    confidence_score: float
    genre_classification: str
    decay_curve_comparison: Dict
    monte_carlo_results: Dict
    scenario_probabilities: Dict
    ai_narrative: str
    risk_factors: List[str]
    opportunities: List[str]


def analyze_historical_trend(yearly_data: Dict[int, float]) -> Dict:
    """
    Analyze historical royalty data to identify trends and patterns.
    Returns trend analysis without requiring AI API.
    """
    if len(yearly_data) < 2:
        return {
            "trend": "insufficient_data",
            "cagr": 0,
            "volatility": 0,
            "trend_direction": "stable"
        }

    years = sorted(yearly_data.keys())
    values = [yearly_data[y] for y in years]

    # Calculate year-over-year changes
    yoy_changes = []
    for i in range(1, len(values)):
        if values[i-1] > 0:
            change = (values[i] - values[i-1]) / values[i-1]
            yoy_changes.append(change)

    # Calculate CAGR
    if values[0] > 0 and values[-1] > 0 and len(years) > 1:
        cagr = (values[-1] / values[0]) ** (1 / (len(years) - 1)) - 1
    else:
        cagr = 0

    # Calculate volatility (standard deviation of returns)
    volatility = statistics.stdev(yoy_changes) if len(yoy_changes) > 1 else 0

    # Determine trend direction
    if cagr > 0.05:
        trend_direction = "growing"
    elif cagr < -0.05:
        trend_direction = "declining"
    else:
        trend_direction = "stable"

    # Check for acceleration/deceleration
    if len(yoy_changes) >= 2:
        recent_avg = statistics.mean(yoy_changes[-2:]) if len(yoy_changes) >= 2 else yoy_changes[-1]
        earlier_avg = statistics.mean(yoy_changes[:-2]) if len(yoy_changes) > 2 else yoy_changes[0]
        if recent_avg > earlier_avg + 0.03:
            momentum = "accelerating"
        elif recent_avg < earlier_avg - 0.03:
            momentum = "decelerating"
        else:
            momentum = "steady"
    else:
        momentum = "insufficient_data"

    return {
        "trend": "analyzed",
        "cagr": cagr,
        "volatility": volatility,
        "trend_direction": trend_direction,
        "momentum": momentum,
        "yoy_changes": yoy_changes,
        "years_of_data": len(years)
    }


def suggest_parameters_from_trend(trend_analysis: Dict, base_year_earnings: float) -> Dict:
    """
    Suggest DCF parameters based on trend analysis.
    This is the fallback when AI API is not available.
    """
    cagr = trend_analysis.get("cagr", 0)
    volatility = trend_analysis.get("volatility", 0.1)
    trend_direction = trend_analysis.get("trend_direction", "stable")
    momentum = trend_analysis.get("momentum", "steady")

    # Base growth rate suggestion on historical CAGR with adjustments
    if trend_direction == "growing":
        # Don't project unsustainable growth
        suggested_growth = min(cagr * 0.7, 0.10)  # Cap at 10%
    elif trend_direction == "declining":
        # Assume some mean reversion
        suggested_growth = max(cagr * 0.5, -0.05)  # Floor at -5%
    else:
        suggested_growth = cagr * 0.8  # Slight regression to mean

    # Adjust for momentum
    if momentum == "accelerating":
        suggested_growth += 0.02
    elif momentum == "decelerating":
        suggested_growth -= 0.02

    # Terminal growth rate (always conservative)
    if trend_direction == "growing":
        terminal_rate = -0.03
    elif trend_direction == "declining":
        terminal_rate = -0.07
    else:
        terminal_rate = -0.05

    # Discount rate based on volatility
    base_discount = 0.12
    volatility_premium = min(volatility * 0.5, 0.04)  # Up to 4% premium
    suggested_discount = base_discount + volatility_premium

    # Confidence based on data quality
    years_of_data = trend_analysis.get("years_of_data", 1)
    if years_of_data >= 4:
        confidence = 0.85
    elif years_of_data >= 3:
        confidence = 0.70
    elif years_of_data >= 2:
        confidence = 0.55
    else:
        confidence = 0.40

    # Reduce confidence for high volatility
    confidence *= max(0.6, 1 - volatility)

    return {
        "growth_rate_1_3": round(suggested_growth, 4),
        "growth_rate_4_5": round(suggested_growth * 0.6, 4),
        "terminal_rate": round(terminal_rate, 4),
        "discount_rate": round(suggested_discount, 4),
        "confidence": round(confidence, 2)
    }


def run_monte_carlo_simulation(
    base_year: float,
    growth_rate: float,
    terminal_rate: float,
    discount_rate: float,
    volatility: float = 0.15,
    n_simulations: int = 1000,
    projection_years: int = 5
) -> Dict:
    """
    Run Monte Carlo simulation to generate probability-weighted valuations.
    """
    np.random.seed(42)  # For reproducibility

    valuations = []
    cash_flow_paths = []

    for _ in range(n_simulations):
        # Simulate uncertain parameters
        sim_growth_1_3 = np.random.normal(growth_rate, volatility * abs(growth_rate) + 0.02)
        sim_growth_4_5 = np.random.normal(growth_rate * 0.6, volatility * abs(growth_rate) + 0.015)
        sim_terminal = np.random.normal(terminal_rate, 0.02)
        sim_discount = np.random.normal(discount_rate, 0.01)

        # Ensure reasonable bounds
        sim_growth_1_3 = np.clip(sim_growth_1_3, -0.20, 0.25)
        sim_growth_4_5 = np.clip(sim_growth_4_5, -0.15, 0.15)
        sim_terminal = np.clip(sim_terminal, -0.15, 0.03)
        sim_discount = np.clip(sim_discount, 0.06, 0.25)

        # Project cash flows
        cf = [base_year]
        for year in range(1, projection_years + 1):
            if year <= 3:
                cf.append(cf[-1] * (1 + sim_growth_1_3))
            else:
                cf.append(cf[-1] * (1 + sim_growth_4_5))

        # Calculate PV of cash flows
        pv_cf = sum(cf[i] / (1 + sim_discount) ** i for i in range(1, projection_years + 1))

        # Terminal value
        if sim_discount > sim_terminal:
            terminal_cf = cf[-1] * (1 + sim_terminal)
            terminal_value = terminal_cf / (sim_discount - sim_terminal)
            pv_terminal = terminal_value / (1 + sim_discount) ** projection_years
        else:
            pv_terminal = cf[-1] * 10  # Fallback multiple

        total_value = pv_cf + pv_terminal
        valuations.append(total_value)
        cash_flow_paths.append(cf)

    valuations = np.array(valuations)

    # Calculate statistics
    percentiles = {
        "p5": float(np.percentile(valuations, 5)),
        "p10": float(np.percentile(valuations, 10)),
        "p25": float(np.percentile(valuations, 25)),
        "p50": float(np.percentile(valuations, 50)),
        "p75": float(np.percentile(valuations, 75)),
        "p90": float(np.percentile(valuations, 90)),
        "p95": float(np.percentile(valuations, 95)),
    }

    return {
        "mean": float(np.mean(valuations)),
        "median": float(np.median(valuations)),
        "std_dev": float(np.std(valuations)),
        "min": float(np.min(valuations)),
        "max": float(np.max(valuations)),
        "percentiles": percentiles,
        "n_simulations": n_simulations,
        "confidence_interval_90": (percentiles["p5"], percentiles["p95"]),
        "downside_risk": percentiles["p10"],
        "upside_potential": percentiles["p90"]
    }


def compare_to_industry_benchmark(
    yearly_data: Dict[int, float],
    genre: str = "mixed"
) -> Dict:
    """
    Compare catalog decay curve against industry benchmarks.
    """
    benchmark = GENRE_DECAY_BENCHMARKS.get(genre, GENRE_DECAY_BENCHMARKS["mixed"])

    if len(yearly_data) < 2:
        return {
            "comparison": "insufficient_data",
            "genre": genre,
            "benchmark": benchmark
        }

    years = sorted(yearly_data.keys())
    values = [yearly_data[y] for y in years]

    # Calculate actual decay rates
    actual_decays = []
    for i in range(1, len(values)):
        if values[i-1] > 0:
            decay = (values[i] - values[i-1]) / values[i-1]
            actual_decays.append(decay)

    # Compare to benchmark
    benchmark_rates = [
        benchmark["year_1"],
        benchmark["year_2"],
        benchmark["year_3"],
        benchmark["year_4"],
        benchmark["year_5_plus"]
    ]

    # Calculate how catalog performs vs benchmark
    comparisons = []
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
    avg_benchmark = statistics.mean(benchmark_rates[:len(actual_decays)]) if actual_decays else 0

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
        "variance_from_benchmark": avg_actual - avg_benchmark
    }


def call_claude_api(
    yearly_data: Dict[int, float],
    trend_analysis: Dict,
    oauth_token: Optional[str] = None
) -> Optional[Dict]:
    """
    Call Claude API for advanced analysis and narrative generation.
    Returns None if OAuth token not available.
    Uses OAuth authentication for Claude Max subscribers.
    """
    if not oauth_token:
        oauth_token = os.environ.get("ANTHROPIC_OAUTH_TOKEN")

    if not oauth_token:
        return None

    try:
        import anthropic

        client = anthropic.Anthropic(
            api_key=oauth_token,
            base_url="https://api.anthropic.com/v1"
        )

        # Prepare the analysis prompt
        prompt = f"""Analyze this music royalty catalog data and provide investment insights:

Historical Earnings by Year:
{json.dumps(yearly_data, indent=2)}

Trend Analysis:
- CAGR: {trend_analysis.get('cagr', 0):.1%}
- Volatility: {trend_analysis.get('volatility', 0):.1%}
- Trend Direction: {trend_analysis.get('trend_direction', 'unknown')}
- Momentum: {trend_analysis.get('momentum', 'unknown')}

Please provide:
1. Genre classification (best guess based on earnings pattern)
2. Key risk factors (list 3-5)
3. Potential opportunities (list 2-3)
4. Recommended scenario weights (bear/base/bull as percentages totaling 100%)
5. A brief narrative summary (2-3 sentences) suitable for an investment memo

Format your response as JSON with keys: genre, risk_factors, opportunities, scenario_weights, narrative"""

        message = client.messages.create(
            model="claude-sonnet-4-20250514",
            max_tokens=1024,
            messages=[
                {"role": "user", "content": prompt}
            ]
        )

        response_text = message.content[0].text

        # Try to parse JSON from response
        try:
            # Find JSON in response
            start = response_text.find('{')
            end = response_text.rfind('}') + 1
            if start >= 0 and end > start:
                return json.loads(response_text[start:end])
        except json.JSONDecodeError:
            pass

        return {
            "genre": "mixed",
            "risk_factors": ["Market volatility", "Streaming rate changes", "Catalog aging"],
            "opportunities": ["Sync licensing potential", "Catalog expansion"],
            "scenario_weights": {"bear": 25, "base": 50, "bull": 25},
            "narrative": response_text[:500]
        }

    except ImportError:
        return None
    except Exception as e:
        print(f"Claude API error: {e}")
        return None


def generate_fallback_narrative(
    trend_analysis: Dict,
    benchmark_comparison: Dict,
    monte_carlo: Dict,
    base_year: float
) -> str:
    """Generate analysis narrative without AI API."""

    trend = trend_analysis.get("trend_direction", "stable")
    momentum = trend_analysis.get("momentum", "steady")
    cagr = trend_analysis.get("cagr", 0)
    benchmark_status = benchmark_comparison.get("overall_status", "at_benchmark")

    # Build narrative
    parts = []

    # Trend description
    if trend == "growing":
        parts.append(f"This catalog shows positive growth momentum with a {cagr:.1%} CAGR.")
    elif trend == "declining":
        parts.append(f"This catalog is experiencing decline with a {cagr:.1%} CAGR.")
    else:
        parts.append(f"This catalog shows stable performance with minimal growth variation.")

    # Benchmark comparison
    if benchmark_status == "above_benchmark":
        parts.append("Performance exceeds industry benchmarks, suggesting strong catalog fundamentals.")
    elif benchmark_status == "below_benchmark":
        parts.append("Performance trails industry benchmarks, warranting conservative assumptions.")
    else:
        parts.append("Performance aligns with typical industry patterns.")

    # Monte Carlo insight
    p10 = monte_carlo["percentiles"]["p10"]
    p90 = monte_carlo["percentiles"]["p90"]
    median = monte_carlo["median"]
    parts.append(f"Monte Carlo analysis suggests a median value of ${median:,.0f} with 80% confidence range of ${p10:,.0f} to ${p90:,.0f}.")

    return " ".join(parts)


def generate_risk_factors(trend_analysis: Dict, benchmark_comparison: Dict) -> List[str]:
    """Generate risk factors based on analysis."""
    risks = []

    volatility = trend_analysis.get("volatility", 0)
    trend = trend_analysis.get("trend_direction", "stable")
    momentum = trend_analysis.get("momentum", "steady")
    benchmark_status = benchmark_comparison.get("overall_status", "at_benchmark")

    # Always present risks
    risks.append("Streaming platform rate changes could impact royalty income")

    # Conditional risks
    if volatility > 0.15:
        risks.append(f"High historical volatility ({volatility:.0%}) increases forecast uncertainty")

    if trend == "declining":
        risks.append("Declining trend may continue without new releases or sync placements")

    if momentum == "decelerating":
        risks.append("Recent deceleration suggests momentum may be weakening")

    if benchmark_status == "below_benchmark":
        risks.append("Underperformance vs. industry peers suggests structural challenges")

    risks.append("Catalog concentration risk if earnings depend on few tracks")

    return risks[:5]  # Max 5 risks


def generate_opportunities(trend_analysis: Dict, benchmark_comparison: Dict) -> List[str]:
    """Generate opportunity factors based on analysis."""
    opportunities = []

    trend = trend_analysis.get("trend_direction", "stable")
    momentum = trend_analysis.get("momentum", "steady")
    benchmark_status = benchmark_comparison.get("overall_status", "at_benchmark")

    # Always present opportunities
    opportunities.append("Sync licensing placements could provide uplift to earnings")

    if trend == "growing" or momentum == "accelerating":
        opportunities.append("Positive momentum could sustain with playlist placements")

    if benchmark_status == "above_benchmark":
        opportunities.append("Strong fundamentals support premium valuation multiple")

    opportunities.append("International streaming growth may offset domestic maturation")
    opportunities.append("Social media virality can revive catalog tracks unexpectedly")

    return opportunities[:4]  # Max 4 opportunities


def run_full_analysis(
    yearly_data: Dict[int, float],
    base_year: float,
    genre: str = "mixed",
    oauth_token: Optional[str] = None,
    n_simulations: int = 1000
) -> AIAnalysisResult:
    """
    Run complete AI-powered analysis pipeline.
    Works with or without OAuth token (falls back to statistical analysis).
    """

    # Step 1: Analyze historical trends
    trend_analysis = analyze_historical_trend(yearly_data)

    # Step 2: Get parameter suggestions
    params = suggest_parameters_from_trend(trend_analysis, base_year)

    # Step 3: Run Monte Carlo simulation
    monte_carlo = run_monte_carlo_simulation(
        base_year=base_year,
        growth_rate=params["growth_rate_1_3"],
        terminal_rate=params["terminal_rate"],
        discount_rate=params["discount_rate"],
        volatility=trend_analysis.get("volatility", 0.15),
        n_simulations=n_simulations
    )

    # Step 4: Compare to industry benchmarks
    benchmark_comparison = compare_to_industry_benchmark(yearly_data, genre)

    # Step 5: Try Claude API for enhanced insights
    ai_response = call_claude_api(yearly_data, trend_analysis, oauth_token)

    if ai_response:
        genre_classification = ai_response.get("genre", genre)
        risk_factors = ai_response.get("risk_factors", generate_risk_factors(trend_analysis, benchmark_comparison))
        opportunities = ai_response.get("opportunities", generate_opportunities(trend_analysis, benchmark_comparison))
        scenario_weights = ai_response.get("scenario_weights", {"bear": 25, "base": 50, "bull": 25})
        narrative = ai_response.get("narrative", "")
    else:
        genre_classification = genre
        risk_factors = generate_risk_factors(trend_analysis, benchmark_comparison)
        opportunities = generate_opportunities(trend_analysis, benchmark_comparison)
        scenario_weights = {"bear": 25, "base": 50, "bull": 25}
        narrative = generate_fallback_narrative(trend_analysis, benchmark_comparison, monte_carlo, base_year)

    return AIAnalysisResult(
        suggested_growth_rate=params["growth_rate_1_3"],
        suggested_terminal_rate=params["terminal_rate"],
        suggested_discount_rate=params["discount_rate"],
        confidence_score=params["confidence"],
        genre_classification=genre_classification,
        decay_curve_comparison=benchmark_comparison,
        monte_carlo_results=monte_carlo,
        scenario_probabilities=scenario_weights,
        ai_narrative=narrative,
        risk_factors=risk_factors,
        opportunities=opportunities
    )
