# U.S. Economy Dashboard — auto-refreshing weekly

A finance-portfolio Excel workbook (macro + micro analysis, real data) that
regenerates itself every Monday morning via GitHub Actions.

## What's inside

- **`U.S._Economy_Dashboard.xlsx`** — the workbook itself (7 tabs: Cover,
  Macro Dashboard, Labor Market, Inflation & Monetary Policy, a real
  gasoline-elasticity natural-experiment case study, a live Phillips Curve
  regression, and an Automation & Sources log).
- **`workbook_layout.py`** — the single source of truth for the workbook's
  structure (styling, formulas, charts). Both the manual build and the
  automated refresh call `build_workbook(data)` from this file, so they can
  never drift out of sync.
- **`refresh_economics_model.py`** — pulls live data from the FRED API
  (Federal Reserve Bank of St. Louis) and the EIA API (U.S. Energy
  Information Administration), shapes it into the `data` dict
  `workbook_layout.build_workbook()` expects, and re-saves the workbook.
- **`.github/workflows/weekly-refresh.yml`** — a GitHub Actions workflow
  that runs the refresh script every Monday at 12:00 UTC and commits the
  updated workbook back to the repo. No server, no local machine needed to
  stay running — GitHub runs this for you, for free.

## One-time setup (~10 minutes)

1. **Get a free FRED API key**: https://fred.stlouisfed.org/docs/api/api_key.html
2. **Get a free EIA API key** (optional, powers the gasoline tab):
   https://www.eia.gov/opendata/register.php
3. Push this folder to a GitHub repo (e.g. `ajaiupadhyaya/economics-dashboard`).
4. In the repo: **Settings → Secrets and variables → Actions → New repository secret**
   - `FRED_API_KEY` = your key
   - `EIA_API_KEY` = your key (optional)
5. Done. The workflow will run automatically every Monday. You can also
   trigger it manually any time from the **Actions** tab → *Weekly Economics
   Dashboard Refresh* → **Run workflow**.

## Running it locally

```bash
pip install requests openpyxl
export FRED_API_KEY=xxxxx
export EIA_API_KEY=xxxxx   # optional
python3 refresh_economics_model.py
```

## Honest limitations

- This repo's automation runs on GitHub's infrastructure, not Claude's —
  Claude (the assistant that built this) has no ability to run on a
  schedule itself; this is the standard way people achieve genuine
  hands-off weekly automation for something like this.
- A few series (e.g. the exact gas-station-visit index used as a quantity
  proxy in the elasticity tab) aren't available via a free public API and
  remain a manually-dated snapshot even after a refresh. The script only
  automates what has a real, free, machine-readable source — it never
  fabricates a number to make everything look automated.
- FRED and EIA both require free registration for an API key; without one,
  the script fails loudly rather than silently serving stale data.

## Data sources

| Series | Source | ID |
|---|---|---|
| Real GDP growth | FRED | `A191RL1Q225SBEA` |
| Unemployment rate | FRED | `UNRATE` |
| Nonfarm payrolls | FRED | `PAYEMS` |
| Labor force participation | FRED | `CIVPART` |
| JOLTS job openings | FRED | `JTSJOL` |
| CPI / Core CPI YoY | FRED | `CPIAUCSL`, `CPILFESL` |
| PCE / Core PCE YoY | FRED | `PCEPI`, `PCEPILFE` |
| Effective Fed funds rate | FRED | `DFF` |
| Treasury yield curve | FRED | `DGS1MO`...`DGS30` |
| Retail gasoline price | EIA | `petroleum/pri/gnd` |
