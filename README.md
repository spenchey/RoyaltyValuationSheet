# Music Royalty Valuation Tool

A DCF (Discounted Cash Flow) valuation tool for music royalty catalogs. Upload historical earnings data and get a complete valuation spreadsheet with AI-powered analysis, Monte Carlo simulations, and industry benchmark comparisons.

Available as both a **desktop GUI** (tkinter) and a **web app** (Flask).

## Features

- **DCF Valuation Model** -- Three-scenario (bear/base/bull) discounted cash flow analysis with probability-weighted output
- **AI-Powered Analysis** -- Uses the Claude API to generate intelligent projections, genre-specific decay curves, and narrative insights
- **Monte Carlo Simulations** -- Runs thousands of randomized scenarios to produce confidence intervals around the valuation
- **Sensitivity Analysis** -- Shows how the valuation changes across discount rates and growth assumptions
- **Genre Decay Benchmarks** -- Built-in decay curves for Pop, Rock, Hip-Hop, Country, Electronic, R&B, Latin, and Classical catalogs
- **Modular Skill System** -- Extensible skills/ directory lets you add new analysis modules without modifying core code
- **Excel Output** -- Generates a formatted .xlsx workbook with all tabs, charts, and color-coded inputs

## Skills Directory

The skills/ folder contains self-contained analysis modules loaded at runtime by skill_loader.py:

| Skill | Description |
|-------|-------------|
| ai-insights | Claude-powered narrative analysis and market commentary |
| dcf-valuation | Core discounted cash flow model |
| decay-curves | Genre-specific revenue decay modeling |
| monte-carlo | Monte Carlo simulation engine |
| sensitivity-analysis | Sensitivity tables across key assumptions |

Each skill has a SKILL.md with metadata (name, version, triggers) and a Python executor.

## Requirements

- Python 3.9+
- Dependencies listed in requirements.txt

## Installation

```bash
git clone https://github.com/spenchey/RoyaltyValuationSheet.git
cd RoyaltyValuationSheet
pip install -r requirements.txt
```

## Environment Variables

Copy the example env file and add your Anthropic API key:

```bash
cp .env.example .env
```

Then edit .env:

```
ANTHROPIC_OAUTH_TOKEN=your-oauth-token-here
```

The AI analysis features require a valid Anthropic token. The tool still works without one, but AI-powered insights will be disabled.

## Usage

### Desktop GUI

Double-click the platform launcher or run from the terminal:

```bash
python run_valuation.py
```

- On **Mac**: use Run Valuation Tool (Mac).command
- On **Windows**: use Run Valuation Tool (Windows).bat

A file picker will open. Select a CSV of historical royalty earnings, fill in the prompts, and an Excel workbook will be saved to the Output Sheets/ directory.

### Web App

```bash
python web_app.py
```

Or on Windows, double-click Run Web App (Windows).bat.

Opens a mobile-friendly web interface (default http://localhost:5000). Upload your CSV through the browser and download the generated valuation spreadsheet.

For production deployment, the app includes gunicorn in its dependencies:

```bash
gunicorn web_app:app
```

## Input Format

The tool expects a CSV of historical royalty earnings with date and amount columns. It will auto-detect the column layout and parse yearly totals for the most recent 3+ years.

## Output

A multi-tab Excel workbook containing:

- Valuation Model (bear/base/bull scenarios)
- AI Analysis summary (if enabled)
- Monte Carlo distribution chart
- Sensitivity tables
- Raw data and decay curve comparisons
