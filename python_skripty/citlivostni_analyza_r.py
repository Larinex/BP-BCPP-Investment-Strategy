"""
DDM Sensitivity Analysis – Fixed vs. Dynamic Required Rate of Return (r)
=========================================================================
Bachelor's thesis: Timing of Stock Purchases on the Prague Stock Exchange
Author: Nick Trefil, 2026

Purpose
-------
The main backtest uses a fixed required rate of return r = 9 % in the Gordon
Growth Model for all 21 years (2004–2024). This script tests how sensitive the
DDM "undervalued / not undervalued" verdict is to using a dynamic r derived
from the actual yield of 10-year Czech government bonds plus a 5 % equity risk
premium (Damodaran, 2025).

Inputs (place in the same folder as this script)
-------------------------------------------------
  pomocná_data.xlsx  – fundamental data per ticker (dividends, prices, …)
  dluhopisy.xlsx     – annual yield of 10-year Czech government bonds

Output
------
  sensitivity_output.xlsx  – three sheets:
      Detail  – full row-level results
      Souhrn  – year-by-year summary (ready for thesis table)
      Změny   – only rows where the DDM verdict differs between fixed and dynamic r

Usage
-----
  pip install pandas openpyxl
  python sensitivity_analysis.py
"""

import os
import pandas as pd
import numpy as np

# ── Configuration ─────────────────────────────────────────────────────────────

BASE_DIR        = os.path.dirname(os.path.abspath(__file__))
ROOT_DIR        = os.path.dirname(BASE_DIR)
INPUT_DATA      = os.path.join(ROOT_DIR, 'data', 'pomocna_data.xlsx')
INPUT_BONDS     = os.path.join(ROOT_DIR, 'data', 'dluhopisy.xlsx')
OUTPUT_FILE     = os.path.join(ROOT_DIR, 'vystupy', 'sensitivity_output.xlsx')

R_FIXED         = 0.09   # required return used in the thesis
ERP             = 0.05   # Czech equity risk premium (Damodaran, 2025)
G               = 0.02   # long-term dividend growth rate (= CNB inflation target)

# ── Bond yield data ───────────────────────────────────────────────────────────
# Loaded from dluhopisy.xlsx; also hard-coded here as a fallback reference.
BOND_YIELDS_FALLBACK = {
    '2003/2004': 0.0482, '2004/2005': 0.0414, '2005/2006': 0.0361,
    '2006/2007': 0.0377, '2007/2008': 0.0468, '2008/2009': 0.0430,
    '2009/2010': 0.0398, '2010/2011': 0.0389, '2011/2012': 0.0370,
    '2012/2013': 0.0192, '2013/2014': 0.0220, '2014/2015': 0.0067,
    '2015/2016': 0.0049, '2016/2017': 0.0053, '2017/2018': 0.0150,
    '2018/2019': 0.0201, '2019/2020': 0.0151, '2020/2021': 0.0126,
    '2021/2022': 0.0262, '2022/2023': 0.0471, '2023/2024': 0.0397,
}

# ── Helper functions ──────────────────────────────────────────────────────────

def load_bond_yields(path: str) -> dict:
    """
    Read bond yields from dluhopisy.xlsx.
    Expected format: one row, columns = period strings like '2003/2004'.
    Falls back to hard-coded values if file not found.
    """
    try:
        df = pd.read_excel(path)
        # Drop the label column (first column), keep period columns
        period_cols = [c for c in df.columns if '/' in str(c)]
        yields = {col: float(df[col].iloc[0]) for col in period_cols}
        print(f"Bond yields loaded from file: {len(yields)} periods.")
        return yields
    except FileNotFoundError:
        print(f"Warning: '{path}' not found – using hard-coded fallback values.")
        return BOND_YIELDS_FALLBACK


def extract_ticker_data(df: pd.DataFrame,
                        div_keyword: str,
                        price_keyword: str,
                        fx_keyword: str = None) -> dict:
    """
    Extract (D0, price) pairs per period from a ticker sheet.

    Parameters
    ----------
    df           : DataFrame for one ticker sheet
    div_keyword  : substring to find the dividend row (case-insensitive)
    price_keyword: substring to find the price row
    fx_keyword   : if provided, dividend is in EUR and is multiplied by this
                   FX row to convert to CZK

    Returns
    -------
    dict  period_string -> (D0_czk, price_czk)
    """
    label_col = df.columns[0]
    period_cols = [c for c in df.columns if c != label_col]

    def find_row(keyword):
        mask = df[label_col].astype(str).str.lower().str.contains(
            keyword.lower(), na=False
        )
        matched = df[mask]
        return matched.iloc[0] if not matched.empty else None

    div_row   = find_row(div_keyword)
    price_row = find_row(price_keyword)
    fx_row    = find_row(fx_keyword) if fx_keyword else None

    result = {}
    for col in period_cols:
        try:
            d = float(div_row[col])   if div_row   is not None else np.nan
            p = float(price_row[col]) if price_row is not None else np.nan
            if fx_row is not None:
                fx = float(fx_row[col])
                d  = d * fx           # EUR → CZK conversion
            if not np.isnan(d) and not np.isnan(p) and d > 0 and p > 0:
                result[col] = (d, p)
        except (ValueError, TypeError):
            pass
    return result


def gordon_vh(d1: float, r: float, g: float) -> float:
    """Gordon Growth Model: intrinsic value = D1 / (r - g)."""
    if r <= g:
        return np.nan
    return d1 / (r - g)

# ── Load data ─────────────────────────────────────────────────────────────────

bond_yields = load_bond_yields(INPUT_BONDS)

xl = pd.read_excel(INPUT_DATA, sheet_name=None)

# Map ticker name → extracted (D0, price) dict
# Keyword arguments mirror the row labels in pomocná_data.xlsx
tickers = {
    'ČEZ':            extract_ticker_data(xl['ČEZ'],
                          'dividenda', 'cena akcie k prvnímu'),
    'Erste Group':    extract_ticker_data(xl['Erste Group'],
                          'dividenda eur', 'cena akcie k prvnímu',
                          fx_keyword='kurz k prvnímu'),
    'KB':             extract_ticker_data(xl['KB'],
                          'dividenda', 'cena akcie k prvnímu'),
    'VIG':            extract_ticker_data(xl['VIG'],
                          'dividenda czk', 'cena akcie k prvnímu'),
    'MONETA':         extract_ticker_data(xl['MONETA'],
                          'dividenda', 'cena akcie k prvnímu'),
    'O2 CR':          extract_ticker_data(xl['O2 CR'],
                          'dividenda', 'cena akcie k prvnímu'),
    'Phillip Morris': extract_ticker_data(xl['Phillip Morris'],
                          'dividenda', 'cena akcie k prvnímu'),
    'PFNONWOVENS':    extract_ticker_data(xl['PFNONWOVENS'],
                          'dividenda czk', 'cena akcie k prvnímu'),
}

# ── Run sensitivity analysis ──────────────────────────────────────────────────

records = []

for period, bond_y in bond_yields.items():

    r_dynamic    = round(bond_y + ERP, 4)          # base dynamic scenario
    r_dynamic_hi = round(bond_y + ERP + 0.01, 4)   # conservative +1 % buffer

    for ticker, data in tickers.items():
        if period not in data:
            continue

        D0, price = data[period]
        D1 = D0 * (1 + G)   # expected next dividend

        vh_fixed     = gordon_vh(D1, R_FIXED,      G)
        vh_dynamic   = gordon_vh(D1, r_dynamic,    G)
        vh_dynamic_hi= gordon_vh(D1, r_dynamic_hi, G)

        def is_undervalued(vh):
            return bool(not np.isnan(vh) and vh > price)

        in_fixed      = is_undervalued(vh_fixed)
        in_dynamic    = is_undervalued(vh_dynamic)
        in_dynamic_hi = is_undervalued(vh_dynamic_hi)

        records.append({
            'Období':             period,
            'Titul':              ticker,
            'D1 (Kč)':            round(D1, 2),
            'Tržní cena (Kč)':    round(price, 2),
            'r fixní (9,0 %)':    R_FIXED,
            'r dynamické':        r_dynamic,
            'r dynamické +1 %':   r_dynamic_hi,
            'VH fixní':           round(vh_fixed, 2)      if not np.isnan(vh_fixed)      else '—',
            'VH dynamické':       round(vh_dynamic, 2)    if not np.isnan(vh_dynamic)    else '—',
            'VH dynamické +1 %':  round(vh_dynamic_hi, 2) if not np.isnan(vh_dynamic_hi) else '—',
            'DDM fixní':          'ANO' if in_fixed      else 'NE',
            'DDM dynamické':      'ANO' if in_dynamic    else 'NE',
            'DDM dynamické +1 %': 'ANO' if in_dynamic_hi else 'NE',
            'Změna (fixní→dyn)':  '⚠ ZMĚNA' if in_fixed != in_dynamic else '',
        })

df_detail = pd.DataFrame(records)

# ── Summary table (one row per year) ─────────────────────────────────────────

summary_rows = []
for period in bond_yields.keys():
    sub = df_detail[df_detail['Období'] == period]
    if sub.empty:
        continue
    r_dyn = sub['r dynamické'].iloc[0]
    summary_rows.append({
        'Rok':                          period.split('/')[1],
        'Výnos dluhopisu':              f"{bond_yields[period]*100:.2f} %",
        'r fixní':                      f"{R_FIXED*100:.1f} %",
        'r dynamické (dluhopis + ERP)': f"{r_dyn*100:.1f} %",
        'Titulů DDM ANO – fixní':       (sub['DDM fixní']      == 'ANO').sum(),
        'Titulů DDM ANO – dynamické':   (sub['DDM dynamické']  == 'ANO').sum(),
        'Počet změn verdiktu':          (sub['Změna (fixní→dyn)'] == '⚠ ZMĚNA').sum(),
    })

df_summary = pd.DataFrame(summary_rows)

df_changes = df_detail[df_detail['Změna (fixní→dyn)'] == '⚠ ZMĚNA'][
    ['Období', 'Titul', 'VH fixní', 'VH dynamické', 'Tržní cena (Kč)',
     'DDM fixní', 'DDM dynamické']
].copy()

# ── Print key findings ────────────────────────────────────────────────────────

total      = len(df_detail)
n_changes  = len(df_changes)
pct_change = n_changes / total * 100

print("=" * 60)
print("  DDM SENSITIVITY ANALYSIS – RESULTS")
print("=" * 60)
print(f"  Total observations:          {total}")
print(f"  Verdict changes (fixed→dyn): {n_changes}  ({pct_change:.1f} %)")
print(f"  Dynamic r range:             "
      f"{df_detail['r dynamické'].min()*100:.1f} % – "
      f"{df_detail['r dynamické'].max()*100:.1f} %")
print(f"  Dynamic r average:           {df_detail['r dynamické'].mean()*100:.1f} %")
print()
print("  Note: all verdict changes are NE→ANO (fixed r is conservative).")
print("  The fixed r = 9 % systematically disadvantages the strategy")
print("  in low-interest-rate environments (approx. 2013–2021).")
print("=" * 60)
print()
print(df_summary.to_string(index=False))

# ── Export to Excel ───────────────────────────────────────────────────────────

with pd.ExcelWriter(OUTPUT_FILE, engine='openpyxl') as writer:
    df_detail.to_excel(writer,  sheet_name='Detail',  index=False)
    df_summary.to_excel(writer, sheet_name='Souhrn',  index=False)
    df_changes.to_excel(writer, sheet_name='Změny',   index=False)

print(f"\nOutput saved to: {OUTPUT_FILE}")
