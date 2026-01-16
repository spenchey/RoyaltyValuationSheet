#!/usr/bin/env python3
"""
Skill-Powered Valuation Orchestrator
Coordinates multiple skills to perform comprehensive royalty valuations.
Inspired by Claude Skills architecture - modular, composable analysis pipeline.
"""

from typing import Dict, Any, Optional
from dataclasses import dataclass, field
from datetime import datetime
import pandas as pd

from skill_loader import SkillRegistry, SkillPipeline


@dataclass
class ValuationResult:
    """Complete valuation result aggregating all skill outputs."""
    royalty_name: str
    base_year_cf: float
    yearly_data: Dict[int, float]

    # DCF Results
    dcf_result: Optional[Dict[str, Any]] = None

    # Monte Carlo Results
    monte_carlo_result: Optional[Dict[str, Any]] = None

    # Decay Curve Analysis
    decay_analysis: Optional[Dict[str, Any]] = None

    # AI Insights
    ai_insights: Optional[Dict[str, Any]] = None

    # Sensitivity Analysis
    sensitivity: Optional[Dict[str, Any]] = None

    # Metadata
    generated_at: str = field(default_factory=lambda: datetime.now().isoformat())
    skills_used: list = field(default_factory=list)


class SkillValuationEngine:
    """
    Orchestrates skill-based valuation analysis.

    Uses modular skills to perform:
    - DCF valuation with scenario analysis
    - Monte Carlo simulations
    - Genre-based decay curve benchmarking
    - AI-powered insights and parameter suggestions
    - Sensitivity analysis
    """

    def __init__(self, skills_dir: Optional[str] = None):
        """Initialize the valuation engine with skill registry."""
        self.registry = SkillRegistry(skills_dir)

    def list_available_skills(self) -> list:
        """List all available valuation skills."""
        return [
            {
                "name": skill.name,
                "description": skill.description,
                "triggers": skill.triggers
            }
            for skill in self.registry.list_skills()
        ]

    def run_full_valuation(
        self,
        royalty_name: str,
        yearly_data: Dict[int, float],
        base_year_cf: Optional[float] = None,
        genre: str = "mixed",
        n_simulations: int = 1000,
        api_key: Optional[str] = None
    ) -> ValuationResult:
        """
        Run complete valuation analysis using all available skills.

        Args:
            royalty_name: Name/ID of the royalty asset
            yearly_data: Historical earnings by year
            base_year_cf: Starting cash flow (defaults to most recent year)
            genre: Music genre for benchmarking
            n_simulations: Number of Monte Carlo simulations
            api_key: Optional Claude API key for enhanced insights

        Returns:
            ValuationResult with comprehensive analysis
        """

        # Determine base year CF
        if base_year_cf is None:
            years = sorted(yearly_data.keys())
            current_year = datetime.now().year
            base_year_cf = yearly_data.get(
                current_year - 1,
                yearly_data.get(max(years), 0) if years else 0
            )

        result = ValuationResult(
            royalty_name=royalty_name,
            base_year_cf=base_year_cf,
            yearly_data=yearly_data
        )

        # Step 1: Get AI insights and suggested parameters
        try:
            ai_insights = self.registry.execute_skill(
                "ai-insights",
                yearly_data=yearly_data,
                base_year_cf=base_year_cf,
                api_key=api_key
            )
            result.ai_insights = ai_insights
            result.skills_used.append("ai-insights")
        except Exception as e:
            print(f"AI Insights skill error: {e}")
            ai_insights = {
                "suggested_growth_rate": 0.05,
                "suggested_discount_rate": 0.12,
                "suggested_terminal_rate": -0.05,
                "confidence_score": 0.5
            }

        # Step 2: Get decay curve analysis
        try:
            decay_analysis = self.registry.execute_skill(
                "decay-curves",
                yearly_data=yearly_data,
                genre=genre
            )
            result.decay_analysis = decay_analysis
            result.skills_used.append("decay-curves")
        except Exception as e:
            print(f"Decay Curves skill error: {e}")

        # Step 3: Run DCF valuation with suggested parameters
        try:
            scenario_weights = ai_insights.get(
                "scenario_probabilities",
                {"bear": 0.25, "base": 0.50, "bull": 0.25}
            )
            # Convert percentages to decimals if needed
            if scenario_weights.get("bear", 0) > 1:
                scenario_weights = {
                    k: v / 100 for k, v in scenario_weights.items()
                }

            dcf_result = self.registry.execute_skill(
                "dcf-valuation",
                base_year_cf=base_year_cf,
                growth_rate_1_3=ai_insights.get("suggested_growth_rate", 0.05),
                growth_rate_4_5=ai_insights.get("suggested_growth_rate_4_5", 0.03),
                discount_rate=ai_insights.get("suggested_discount_rate", 0.12),
                terminal_growth=ai_insights.get("suggested_terminal_rate", -0.05),
                scenario_weights=scenario_weights
            )
            result.dcf_result = {
                "bear_value": dcf_result.bear_value,
                "base_value": dcf_result.base_value,
                "bull_value": dcf_result.bull_value,
                "weighted_value": dcf_result.weighted_value,
                "pv_cash_flows": dcf_result.pv_cash_flows,
                "pv_terminal": dcf_result.pv_terminal,
                "ev_multiple": dcf_result.ev_multiple,
                "cash_flows": dcf_result.cash_flows
            }
            result.skills_used.append("dcf-valuation")
        except Exception as e:
            print(f"DCF Valuation skill error: {e}")

        # Step 4: Run Monte Carlo simulation
        try:
            monte_carlo_result = self.registry.execute_skill(
                "monte-carlo",
                base_year_cf=base_year_cf,
                growth_rate=ai_insights.get("suggested_growth_rate", 0.05),
                terminal_rate=ai_insights.get("suggested_terminal_rate", -0.05),
                discount_rate=ai_insights.get("suggested_discount_rate", 0.12),
                volatility=ai_insights.get("trend_analysis", {}).get("volatility", 0.15),
                n_simulations=n_simulations
            )
            result.monte_carlo_result = monte_carlo_result
            result.skills_used.append("monte-carlo")
        except Exception as e:
            print(f"Monte Carlo skill error: {e}")

        # Step 5: Run sensitivity analysis
        try:
            sensitivity = self.registry.execute_skill(
                "sensitivity-analysis",
                base_year_cf=base_year_cf,
                base_growth=ai_insights.get("suggested_growth_rate", 0.05),
                base_discount=ai_insights.get("suggested_discount_rate", 0.12),
                base_terminal=ai_insights.get("suggested_terminal_rate", -0.05),
                growth_4_5=ai_insights.get("suggested_growth_rate_4_5", 0.03)
            )
            result.sensitivity = sensitivity
            result.skills_used.append("sensitivity-analysis")
        except Exception as e:
            print(f"Sensitivity Analysis skill error: {e}")

        return result

    def run_quick_valuation(
        self,
        base_year_cf: float,
        growth_rate: float = 0.05,
        discount_rate: float = 0.12,
        terminal_growth: float = -0.05
    ) -> Dict[str, Any]:
        """
        Run quick DCF-only valuation without full analysis pipeline.

        Args:
            base_year_cf: Starting cash flow
            growth_rate: Growth rate for years 1-3
            discount_rate: Discount rate
            terminal_growth: Terminal growth rate

        Returns:
            Dictionary with DCF results
        """
        dcf_result = self.registry.execute_skill(
            "dcf-valuation",
            base_year_cf=base_year_cf,
            growth_rate_1_3=growth_rate,
            discount_rate=discount_rate,
            terminal_growth=terminal_growth
        )

        return {
            "bear_value": dcf_result.bear_value,
            "base_value": dcf_result.base_value,
            "bull_value": dcf_result.bull_value,
            "weighted_value": dcf_result.weighted_value,
            "ev_multiple": dcf_result.ev_multiple
        }

    def analyze_csv(
        self,
        csv_path: str,
        genre: str = "mixed",
        n_simulations: int = 1000,
        api_key: Optional[str] = None
    ) -> ValuationResult:
        """
        Analyze a CSV file and run full valuation.

        Args:
            csv_path: Path to CSV file with royalty data
            genre: Music genre for benchmarking
            n_simulations: Number of Monte Carlo simulations
            api_key: Optional Claude API key

        Returns:
            ValuationResult with comprehensive analysis
        """
        # Read CSV
        df = pd.read_csv(csv_path)

        # Find amount column
        amount_col = None
        for col in ['payable_amount', 'amount', 'earnings', 'royalty']:
            if col in df.columns.str.lower().tolist():
                amount_col = [c for c in df.columns if c.lower() == col][0]
                break

        if amount_col is None:
            amount_cols = [c for c in df.columns if 'amount' in c.lower()]
            if amount_cols:
                amount_col = amount_cols[0]
            else:
                raise ValueError("Could not find amount column in CSV")

        # Find year column
        year_col = None
        for col in ['distribution_year', 'year', 'date']:
            if col in df.columns.str.lower().tolist():
                year_col = [c for c in df.columns if c.lower() == col][0]
                break

        if year_col is None:
            raise ValueError("Could not find year column in CSV")

        # Aggregate by year
        yearly = df.groupby(year_col)[amount_col].sum().sort_index()
        yearly_data = {int(k): float(v) for k, v in yearly.items()}

        # Extract royalty name from filename
        import os
        import re
        base_name = os.path.basename(csv_path)
        base_name = os.path.splitext(base_name)[0]

        if 'listing' in base_name.lower():
            match = re.search(r'listing[-_]?(\d+)', base_name, re.IGNORECASE)
            if match:
                royalty_name = f"Listing {match.group(1)}"
            else:
                royalty_name = base_name
        else:
            royalty_name = base_name

        return self.run_full_valuation(
            royalty_name=royalty_name,
            yearly_data=yearly_data,
            genre=genre,
            n_simulations=n_simulations,
            api_key=api_key
        )


def create_pipeline(registry: SkillRegistry) -> SkillPipeline:
    """
    Create a standard valuation pipeline.

    Pipeline flow:
    1. ai-insights -> suggests parameters
    2. decay-curves -> benchmark comparison
    3. dcf-valuation -> scenario analysis
    4. monte-carlo -> probabilistic valuation
    5. sensitivity-analysis -> parameter impact
    """
    pipeline = SkillPipeline(registry)

    # Add steps with parameter mappers
    pipeline.add_step("ai-insights", lambda d: {
        "yearly_data": d["yearly_data"],
        "base_year_cf": d["base_year_cf"],
        "api_key": d.get("api_key")
    })

    pipeline.add_step("decay-curves", lambda d: {
        "yearly_data": d["yearly_data"],
        "genre": d.get("genre", "mixed")
    })

    pipeline.add_step("dcf-valuation", lambda d: {
        "base_year_cf": d["base_year_cf"],
        "growth_rate_1_3": d.get("suggested_growth_rate", 0.05),
        "discount_rate": d.get("suggested_discount_rate", 0.12),
        "terminal_growth": d.get("suggested_terminal_rate", -0.05)
    })

    pipeline.add_step("monte-carlo", lambda d: {
        "base_year_cf": d["base_year_cf"],
        "growth_rate": d.get("suggested_growth_rate", 0.05),
        "terminal_rate": d.get("suggested_terminal_rate", -0.05),
        "discount_rate": d.get("suggested_discount_rate", 0.12),
        "n_simulations": d.get("n_simulations", 1000)
    })

    return pipeline


# Convenience function for quick valuations
def quick_value(base_year_cf: float, **kwargs) -> Dict[str, Any]:
    """Quick valuation helper function."""
    engine = SkillValuationEngine()
    return engine.run_quick_valuation(base_year_cf, **kwargs)


if __name__ == "__main__":
    # Example usage
    engine = SkillValuationEngine()

    print("Available Skills:")
    for skill in engine.list_available_skills():
        print(f"  - {skill['name']}: {skill['description'][:60]}...")

    print("\nRunning quick valuation for $10,000 base year CF:")
    result = quick_value(10000)
    print(f"  Weighted Value: ${result['weighted_value']:,.2f}")
    print(f"  EV Multiple: {result['ev_multiple']:.1f}x")
