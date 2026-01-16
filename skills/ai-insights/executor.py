#!/usr/bin/env python3
"""
AI Insights Skill Executor
Provides intelligent analysis and narrative generation for royalty valuations.
"""

import os
import json
import statistics
from typing import Dict, Any, List, Optional


def analyze_trend(yearly_data: Dict[int, float]) -> Dict[str, Any]:
    """Analyze historical royalty data for trends and patterns."""

    if len(yearly_data) < 2:
        return {
            "trend": "insufficient_data",
            "cagr": 0,
            "volatility": 0,
            "trend_direction": "stable",
            "years_of_data": len(yearly_data)
        }

    years = sorted(yearly_data.keys())
    values = [yearly_data[y] for y in years]

    # Calculate year-over-year changes
    yoy_changes = []
    for i in range(1, len(values)):
        if values[i - 1] > 0:
            change = (values[i] - values[i - 1]) / values[i - 1]
            yoy_changes.append(change)

    # Calculate CAGR
    if values[0] > 0 and values[-1] > 0 and len(years) > 1:
        cagr = (values[-1] / values[0]) ** (1 / (len(years) - 1)) - 1
    else:
        cagr = 0

    # Calculate volatility
    volatility = statistics.stdev(yoy_changes) if len(yoy_changes) > 1 else 0

    # Determine trend direction
    if cagr > 0.05:
        trend_direction = "growing"
    elif cagr < -0.05:
        trend_direction = "declining"
    else:
        trend_direction = "stable"

    # Check momentum
    if len(yoy_changes) >= 2:
        recent_avg = statistics.mean(yoy_changes[-2:])
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


def suggest_parameters(trend_analysis: Dict[str, Any]) -> Dict[str, Any]:
    """Suggest DCF parameters based on trend analysis."""

    cagr = trend_analysis.get("cagr", 0)
    volatility = trend_analysis.get("volatility", 0.1)
    trend_direction = trend_analysis.get("trend_direction", "stable")
    momentum = trend_analysis.get("momentum", "steady")

    # Growth rate suggestion
    if trend_direction == "growing":
        suggested_growth = min(cagr * 0.7, 0.10)  # Cap at 10%
    elif trend_direction == "declining":
        suggested_growth = max(cagr * 0.5, -0.05)  # Floor at -5%
    else:
        suggested_growth = cagr * 0.8

    # Adjust for momentum
    if momentum == "accelerating":
        suggested_growth += 0.02
    elif momentum == "decelerating":
        suggested_growth -= 0.02

    # Terminal growth rate
    if trend_direction == "growing":
        terminal_rate = -0.03
    elif trend_direction == "declining":
        terminal_rate = -0.07
    else:
        terminal_rate = -0.05

    # Discount rate based on volatility
    base_discount = 0.12
    volatility_premium = min(volatility * 0.5, 0.04)
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


def generate_risk_factors(
    trend_analysis: Dict[str, Any],
    benchmark_comparison: Optional[Dict[str, Any]] = None
) -> List[str]:
    """Generate risk factors based on analysis."""

    risks = []

    volatility = trend_analysis.get("volatility", 0)
    trend = trend_analysis.get("trend_direction", "stable")
    momentum = trend_analysis.get("momentum", "steady")
    benchmark_status = (
        benchmark_comparison.get("overall_status", "at_benchmark")
        if benchmark_comparison else "at_benchmark"
    )

    # Always present risks
    risks.append("Streaming platform rate changes could impact royalty income")

    # Conditional risks
    if volatility > 0.15:
        risks.append(
            f"High historical volatility ({volatility:.0%}) increases forecast uncertainty"
        )

    if trend == "declining":
        risks.append(
            "Declining trend may continue without new releases or sync placements"
        )

    if momentum == "decelerating":
        risks.append("Recent deceleration suggests momentum may be weakening")

    if benchmark_status == "below_benchmark":
        risks.append(
            "Underperformance vs. industry peers suggests structural challenges"
        )

    risks.append("Catalog concentration risk if earnings depend on few tracks")

    return risks[:5]


def generate_opportunities(
    trend_analysis: Dict[str, Any],
    benchmark_comparison: Optional[Dict[str, Any]] = None
) -> List[str]:
    """Generate opportunity factors based on analysis."""

    opportunities = []

    trend = trend_analysis.get("trend_direction", "stable")
    momentum = trend_analysis.get("momentum", "steady")
    benchmark_status = (
        benchmark_comparison.get("overall_status", "at_benchmark")
        if benchmark_comparison else "at_benchmark"
    )

    opportunities.append("Sync licensing placements could provide uplift to earnings")

    if trend == "growing" or momentum == "accelerating":
        opportunities.append("Positive momentum could sustain with playlist placements")

    if benchmark_status == "above_benchmark":
        opportunities.append("Strong fundamentals support premium valuation multiple")

    opportunities.append("International streaming growth may offset domestic maturation")
    opportunities.append("Social media virality can revive catalog tracks unexpectedly")

    return opportunities[:4]


def generate_narrative(
    trend_analysis: Dict[str, Any],
    benchmark_comparison: Optional[Dict[str, Any]] = None,
    monte_carlo: Optional[Dict[str, Any]] = None,
    base_year_cf: float = 0
) -> str:
    """Generate analysis narrative without AI API."""

    trend = trend_analysis.get("trend_direction", "stable")
    momentum = trend_analysis.get("momentum", "steady")
    cagr = trend_analysis.get("cagr", 0)
    benchmark_status = (
        benchmark_comparison.get("overall_status", "at_benchmark")
        if benchmark_comparison else "at_benchmark"
    )

    parts = []

    # Trend description
    if trend == "growing":
        parts.append(f"This catalog shows positive growth momentum with a {cagr:.1%} CAGR.")
    elif trend == "declining":
        parts.append(f"This catalog is experiencing decline with a {cagr:.1%} CAGR.")
    else:
        parts.append("This catalog shows stable performance with minimal growth variation.")

    # Benchmark comparison
    if benchmark_status == "above_benchmark":
        parts.append(
            "Performance exceeds industry benchmarks, suggesting strong catalog fundamentals."
        )
    elif benchmark_status == "below_benchmark":
        parts.append(
            "Performance trails industry benchmarks, warranting conservative assumptions."
        )
    else:
        parts.append("Performance aligns with typical industry patterns.")

    # Monte Carlo insight
    if monte_carlo:
        p10 = monte_carlo["percentiles"]["p10"]
        p90 = monte_carlo["percentiles"]["p90"]
        median = monte_carlo["median"]
        parts.append(
            f"Monte Carlo analysis suggests a median value of ${median:,.0f} "
            f"with 80% confidence range of ${p10:,.0f} to ${p90:,.0f}."
        )

    return " ".join(parts)


def call_claude_api(
    yearly_data: Dict[int, float],
    trend_analysis: Dict[str, Any],
    api_key: Optional[str] = None
) -> Optional[Dict[str, Any]]:
    """Call Claude API for enhanced analysis (optional)."""

    if not api_key:
        api_key = os.environ.get("ANTHROPIC_API_KEY")

    if not api_key:
        return None

    try:
        import anthropic

        client = anthropic.Anthropic(api_key=api_key)

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
            messages=[{"role": "user", "content": prompt}]
        )

        response_text = message.content[0].text

        # Parse JSON from response
        start = response_text.find('{')
        end = response_text.rfind('}') + 1
        if start >= 0 and end > start:
            return json.loads(response_text[start:end])

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


def execute(
    yearly_data: Dict[int, float],
    base_year_cf: float,
    trend_analysis: Optional[Dict[str, Any]] = None,
    benchmark_comparison: Optional[Dict[str, Any]] = None,
    monte_carlo: Optional[Dict[str, Any]] = None,
    api_key: Optional[str] = None
) -> Dict[str, Any]:
    """
    Generate AI-powered insights for royalty valuation.

    Args:
        yearly_data: Dictionary mapping years to earnings
        base_year_cf: Starting cash flow for valuation
        trend_analysis: Pre-computed trend analysis (optional)
        benchmark_comparison: Decay curve comparison results (optional)
        monte_carlo: Monte Carlo simulation results (optional)
        api_key: Claude API key for enhanced insights (optional)

    Returns:
        Dictionary with suggested parameters and analysis
    """

    # Run trend analysis if not provided
    if trend_analysis is None:
        trend_analysis = analyze_trend(yearly_data)

    # Get parameter suggestions
    params = suggest_parameters(trend_analysis)

    # Try Claude API for enhanced insights
    ai_response = call_claude_api(yearly_data, trend_analysis, api_key)

    if ai_response:
        genre = ai_response.get("genre", "mixed")
        risk_factors = ai_response.get(
            "risk_factors",
            generate_risk_factors(trend_analysis, benchmark_comparison)
        )
        opportunities = ai_response.get(
            "opportunities",
            generate_opportunities(trend_analysis, benchmark_comparison)
        )
        scenario_weights = ai_response.get(
            "scenario_weights",
            {"bear": 25, "base": 50, "bull": 25}
        )
        narrative = ai_response.get("narrative", "")
    else:
        genre = "mixed"
        risk_factors = generate_risk_factors(trend_analysis, benchmark_comparison)
        opportunities = generate_opportunities(trend_analysis, benchmark_comparison)
        scenario_weights = {"bear": 25, "base": 50, "bull": 25}
        narrative = generate_narrative(
            trend_analysis, benchmark_comparison, monte_carlo, base_year_cf
        )

    return {
        "suggested_growth_rate": params["growth_rate_1_3"],
        "suggested_growth_rate_4_5": params["growth_rate_4_5"],
        "suggested_terminal_rate": params["terminal_rate"],
        "suggested_discount_rate": params["discount_rate"],
        "confidence_score": params["confidence"],
        "genre_classification": genre,
        "scenario_probabilities": scenario_weights,
        "risk_factors": risk_factors,
        "opportunities": opportunities,
        "ai_narrative": narrative,
        "trend_analysis": trend_analysis,
        "api_enhanced": ai_response is not None
    }
