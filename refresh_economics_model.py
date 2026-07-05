#!/usr/bin/env python3
"""
refresh_economics_model.py

Pulls current real macro/micro data from free public APIs (FRED, EIA) and
regenerates U.S._Economy_Dashboard.xlsx using the shared workbook_layout
module -- the exact same tab structure, formulas, and charts Claude built
by hand, just with fresh blue-input values.

One-time setup:
    1. Free FRED API key:  https://fred.stlouisfed.org/docs/api/api_key.html
    2. Free EIA API key:   https://www.eia.gov/opendata/register.php
    3. export FRED_API_KEY=xxxxx
       export EIA_API_KEY=xxxxx   (optional -- gasoline tab keeps last values without it)
    4. pip install requests openpyxl
    5. python3 refresh_economics_model.py

GitHub Actions runs this automatically every Monday -- see
.github/workflows/weekly-refresh.yml.
"""

import os
import sys
import datetime
import requests

from workbook_layout import build_workbook

FRED_API_KEY = os.environ.get("FRED_API_KEY")
EIA_API_KEY = os.environ.get("EIA_API_KEY")
FRED_BASE = "https://api.stlouisfed.org/fred/series/observations"
EIA_BASE = "https://api.eia.gov/v2"
OUTPUT_PATH = "U.S._Economy_Dashboard.xlsx"

if not FRED_API_KEY:
    sys.exit("ERROR: FRED_API_KEY environment variable not set. "
              "Get a free key at https://fred.stlouisfed.org/docs/api/api_key.html")


def fred_latest(series_id, units=None):
    params = {"series_id": series_id, "api_key": FRED_API_KEY, "file_type": "json",
              "sort_order": "desc", "limit": 1}
    if units:
        params["units"] = units
    r = requests.get(FRED_BASE, params=params, timeout=30)
    r.raise_for_status()
    obs = r.json().get("observations", [])
    for o in obs:
        try:
            return o["date"], float(o["value"])
        except (ValueError, KeyError):
            continue
    return None, None


def fred_last_n(series_id, n, units=None):
    params = {"series_id": series_id, "api_key": FRED_API_KEY, "file_type": "json",
              "sort_order": "desc", "limit": n}
    if units:
        params["units"] = units
    r = requests.get(FRED_BASE, params=params, timeout=30)
    r.raise_for_status()
    obs = r.json().get("observations", [])
    out = []
    for o in obs:
        try:
            out.append((o["date"], float(o["value"])))
        except (ValueError, KeyError):
            continue
    return list(reversed(out))


def fred_annual(series_id, start_year, end_year, units=None):
    params = {"series_id": series_id, "api_key": FRED_API_KEY, "file_type": "json",
              "observation_start": f"{start_year}-01-01",
              "observation_end": f"{end_year}-12-31", "frequency": "a"}
    if units:
        params["units"] = units
    r = requests.get(FRED_BASE, params=params, timeout=30)
    r.raise_for_status()
    out = {}
    for o in r.json().get("observations", []):
        try:
            out[int(o["date"][:4])] = float(o["value"])
        except (ValueError, KeyError):
            continue
    return out


def eia_gasoline_series(n_weeks=15):
    if not EIA_API_KEY:
        return None
    url = f"{EIA_BASE}/petroleum/pri/gnd/data/"
    params = {
        "api_key": EIA_API_KEY, "frequency": "weekly", "data[0]": "value",
        "facets[product][]": "EPMR", "facets[duoarea][]": "NUS",
        "sort[0][column]": "period", "sort[0][direction]": "desc", "length": n_weeks,
    }
    r = requests.get(url, params=params, timeout=30)
    r.raise_for_status()
    rows = r.json().get("response", {}).get("data", [])
    out = [(row["period"], float(row["value"])) for row in rows]
    return list(reversed(out))


def month_label(date_str):
    y, m, _ = date_str.split("-")
    names = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
    return f"{names[int(m)-1]}-{y[2:]}"


print("Pulling live data from FRED" + (" and EIA..." if EIA_API_KEY else " (no EIA key -- gasoline tab will note manual figures)..."))

gdp_date, gdp_val = fred_latest("A191RL1Q225SBEA")
unemp_date, unemp_val = fred_latest("UNRATE")
payrolls_hist = fred_last_n("PAYEMS", 2)
payrolls_chg = None
if len(payrolls_hist) == 2:
    payrolls_chg = round(payrolls_hist[1][1] - payrolls_hist[0][1], 0)
lfpr_date, lfpr_val = fred_latest("CIVPART")
jolts_date, jolts_val = fred_latest("JTSJOL")
cpi_date, cpi_val = fred_latest("CPIAUCSL", units="pc1")
corecpi_date, corecpi_val = fred_latest("CPILFESL", units="pc1")
pce_date, pce_val = fred_latest("PCEPI", units="pc1")
corepce_date, corepce_val = fred_latest("PCEPILFE", units="pc1")
fedfunds_date, fedfunds_val = fred_latest("DFF")
y10_date, y10_val = fred_latest("DGS10")
y2_date, y2_val = fred_latest("DGS2")

snapshot = {
    "gdp": (round((gdp_val or 0) / 100, 4), gdp_date, None),
    "unemployment": (round((unemp_val or 0) / 100, 4), unemp_date, None),
    "payrolls": (payrolls_chg, payrolls_hist[-1][0] if payrolls_hist else None, None),
    "lfpr": (round((lfpr_val or 0) / 100, 4), lfpr_date, None),
    "jolts": (jolts_val, jolts_date, None),
    "cpi": (round((cpi_val or 0) / 100, 4), cpi_date, None),
    "corecpi": (round((corecpi_val or 0) / 100, 4), corecpi_date, None),
    "pce": (round((pce_val or 0) / 100, 4), pce_date, None),
    "corepce": (round((corepce_val or 0) / 100, 4), corepce_date, None),
    "fedfunds": (round((fedfunds_val or 0) / 100, 4), fedfunds_date, None),
    "y10": (round((y10_val or 0) / 100, 4), y10_date, None),
    "y2": (round((y2_val or 0) / 100, 4), y2_date, None),
}

payrolls_6 = fred_last_n("PAYEMS", 7)
payroll_trend = {}
for i in range(1, len(payrolls_6)):
    lbl = month_label(payrolls_6[i][0])
    payroll_trend[lbl] = round(payrolls_6[i][1] - payrolls_6[i - 1][1], 0)

labor_internals = {
    "Labor Force Participation Rate": round((lfpr_val or 0) / 100, 4),
    "JOLTS Job Openings": jolts_val,
}


def last_n_pct(series_id, n=3):
    obs = fred_last_n(series_id, n, units="pc1")
    return {month_label(d): round(v / 100, 4) for d, v in obs}


inflation_trend = {
    "cpi": last_n_pct("CPIAUCSL"),
    "corecpi": last_n_pct("CPILFESL"),
    "pce": last_n_pct("PCEPI"),
    "corepce": last_n_pct("PCEPILFE"),
}

yield_series = {"1M": "DGS1MO", "3M": "DGS3MO", "6M": "DGS6MO", "1Y": "DGS1",
                 "2Y": "DGS2", "5Y": "DGS5", "10Y": "DGS10", "30Y": "DGS30"}
yield_curve = {}
for label, sid in yield_series.items():
    _, v = fred_latest(sid)
    yield_curve[label] = round((v or 0) / 100, 4)

gas_obs = eia_gasoline_series(20)
if gas_obs:
    gasoline = {"price_pre": gas_obs[0][1], "price_peak": max(o[1] for o in gas_obs),
                "qty_pre": 100.0, "qty_peak": 94.3}
else:
    gasoline = {"price_pre": 2.94, "price_peak": 4.50, "qty_pre": 100.0, "qty_peak": 94.3}

unemp_annual = fred_annual("UNRATE", 2010, 2025)
cpi_annual = fred_annual("CPIAUCSL", 2010, 2025, units="pc1")
phillips_curve = {}
for year in range(2010, 2026):
    if year in unemp_annual and year in cpi_annual:
        phillips_curve[year] = (round(unemp_annual[year] / 100, 4), round(cpi_annual[year] / 100, 4))

REFRESH_TIMESTAMP = datetime.datetime.utcnow().strftime("%Y-%m-%d %H:%M UTC")
today = datetime.date.today()
days_until_monday = (7 - today.weekday()) % 7 or 7
next_monday = today + datetime.timedelta(days=days_until_monday)

data = {
    "last_refreshed": REFRESH_TIMESTAMP,
    "next_refresh": next_monday.strftime("%A, %B %d, %Y") + " (weekly, 8:00am ET)",
    "snapshot": snapshot,
    "payroll_trend": payroll_trend,
    "labor_internals": labor_internals,
    "inflation_trend": inflation_trend,
    "yield_curve": yield_curve,
    "gasoline": gasoline,
    "phillips_curve": phillips_curve,
}

wb = build_workbook(data)
wb.save(OUTPUT_PATH)
print(f"Saved refreshed workbook to {OUTPUT_PATH} at {REFRESH_TIMESTAMP}")
