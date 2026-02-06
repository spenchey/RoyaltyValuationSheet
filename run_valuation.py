#!/usr/bin/env python3
"""
Royalty Valuation Runner
Usage: python run_valuation.py <csv_file> [listing_name] [genre]

Example:
    python run_valuation.py earnings.csv "Production_Music_2826" mixed
"""

import pandas as pd
import os
import sys
from datetime import datetime

def run_valuation(csv_path, listing_name=None, genre="mixed"):
    """
    Process a Royalty Exchange CSV and generate full valuation Excel.
    
    Args:
        csv_path: Path to the CSV file from Royalty Exchange
        listing_name: Name for the output file (defaults to filename)
        genre: Genre for decay benchmarks (pop, rock, hip_hop, country, electronic, classical, mixed)
    """
    # Read the raw CSV
    raw_df = pd.read_csv(csv_path, low_memory=False)
    print(f"Loaded {len(raw_df)} rows from {csv_path}")
    
    # Detect column format and aggregate by year
    if 'distribution_year' in raw_df.columns and 'payable_amount' in raw_df.columns:
        # Royalty Exchange format
        yearly = raw_df.groupby('distribution_year')['payable_amount'].sum().reset_index()
        yearly.columns = ['year', 'earnings']
    elif 'year' in raw_df.columns and 'earnings' in raw_df.columns:
        # Simple format
        yearly = raw_df.groupby('year')['earnings'].sum().reset_index()
    else:
        print("ERROR: CSV must have either:")
        print("  - 'distribution_year' and 'payable_amount' columns (Royalty Exchange format)")
        print("  - 'year' and 'earnings' columns (simple format)")
        sys.exit(1)
    
    print("\nYearly earnings:")
    for _, row in yearly.iterrows():
        print(f"  {int(row['year'])}: ${row['earnings']:,.2f}")
    
    # Get earnings dict
    earnings = {int(row['year']): row['earnings'] for _, row in yearly.iterrows()}
    
    # Get most recent 5 years
    recent_years = sorted(earnings.keys())[-5:]
    print(f"\nUsing years for model: {recent_years}")
    
    # Prepare data for valuation model (need at least 4 years)
    if len(recent_years) >= 4:
        year_minus_3 = earnings[recent_years[-4]]
        year_minus_2 = earnings[recent_years[-3]]
        year_minus_1 = earnings[recent_years[-2]]
        ytd = earnings[recent_years[-1]]
        base_year = (year_minus_3 + year_minus_2 + year_minus_1) / 3
    elif len(recent_years) >= 3:
        year_minus_3 = earnings[recent_years[-3]]
        year_minus_2 = earnings[recent_years[-2]]
        year_minus_1 = earnings[recent_years[-1]]
        ytd = year_minus_1
        base_year = (year_minus_3 + year_minus_2 + year_minus_1) / 3
    else:
        print("ERROR: Need at least 3 years of data")
        sys.exit(1)
    
    # Import modules (after confirming data is good)
    from web_app import create_valuation_template
    from ai_analysis import run_full_analysis
    
    # Run AI analysis (statistical fallback if no API key)
    yearly_data = {int(k): float(v) for k, v in earnings.items() if int(k) >= recent_years[0]}
    print(f"\nRunning analysis...")
    
    # Check for OAuth token in environment or .env file
    api_key = os.environ.get('ANTHROPIC_OAUTH_TOKEN') or os.environ.get('ANTHROPIC_API_KEY')
    
    ai_result = run_full_analysis(
        yearly_data=yearly_data,
        base_year=base_year,
        genre=genre,
        api_key=api_key,
        n_simulations=1000
    )
    
    # Generate listing name from filename if not provided
    if not listing_name:
        listing_name = os.path.splitext(os.path.basename(csv_path))[0]
    
    # Create valuation template with all sheets
    output_bytes = create_valuation_template(
        royalty_name=listing_name,
        year_minus_3=year_minus_3,
        year_minus_2=year_minus_2,
        year_minus_1=year_minus_1,
        ytd=ytd,
        base_year=base_year,
        ai_analysis=ai_result,
        yearly_data=yearly_data,
        raw_df=raw_df
    )
    
    # Save output
    script_dir = os.path.dirname(os.path.abspath(__file__))
    output_dir = os.path.join(script_dir, "Output Sheets")
    os.makedirs(output_dir, exist_ok=True)
    
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    # Clean listing name for filename
    safe_name = "".join(c if c.isalnum() or c in ('_', '-') else '_' for c in listing_name)
    filename = f"{safe_name}_{timestamp}.xlsx"
    filepath = os.path.join(output_dir, filename)
    
    with open(filepath, 'wb') as f:
        f.write(output_bytes.read())
    
    print(f"\n✓ Saved: {filepath}")
    print(f"\nModel inputs:")
    print(f"  Year -3: ${year_minus_3:,.2f}")
    print(f"  Year -2: ${year_minus_2:,.2f}")
    print(f"  Year -1: ${year_minus_1:,.2f}")
    print(f"  YTD: ${ytd:,.2f}")
    print(f"  Base Year (3yr avg): ${base_year:,.2f}")
    print(f"\nSheets generated:")
    print(f"  - Valuation Model (full DCF with scenarios)")
    print(f"  - AI Analysis (parameters, risks, opportunities)")
    print(f"  - Monte Carlo (simulation results)")
    print(f"  - Decay Benchmarks (industry comparison)")
    print(f"  - Raw Data (complete input file)")
    
    return filepath


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)
    
    csv_path = sys.argv[1]
    listing_name = sys.argv[2] if len(sys.argv) > 2 else None
    genre = sys.argv[3] if len(sys.argv) > 3 else "mixed"
    
    run_valuation(csv_path, listing_name, genre)
