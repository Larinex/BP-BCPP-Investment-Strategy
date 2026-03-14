"""
No Look-Ahead Analysis – DDM with Prior-Year Dividend (D1)
==========================================================
Bachelor's thesis: Timing of Stock Purchases on the Prague Stock Exchange
Author: Nick Trefil, 2026

Purpose
-------
The main backtest uses the actual dividend paid in year T as D1 in the Gordon
Growth Model (correct foresight assumption). This script tests how sensitive
the DDM "undervalued / not undervalued" verdict is when D1 is replaced by the
prior-year dividend grown by g – i.e. eliminating the look-ahead assumption.

For EUR-denominated dividends (Erste Group, VIG, PFNONWOVENS), the historical
CZK/EUR exchange rate from the first trading day of each year is used instead
of a fixed rate.

Inputs (place in the same folder as this script)
-------------------------------------------------
  screening.xlsx                      – fundamental screening results per year
  pomocna_data_no_lookahead_analyza.xlsx – historical dividends, prices and
                                          CZK/EUR exchange rates per ticker

Output
------
  no_lookahead_output.xlsx  – two sheets:
      Srovnání    – year-by-year comparison of look-ahead vs. no look-ahead
      Souhrn      – aggregate statistics

Usage
-----
  pip install pandas openpyxl
  python no_lookahead_analyza.py
"""

import pandas as pd
import numpy as np
import warnings
import os
warnings.filterwarnings('ignore')

# ==========================================
# KONFIGURACE
# ==========================================

BASE_DIR         = os.path.dirname(os.path.abspath(__file__))
ROOT_DIR         = os.path.dirname(BASE_DIR)
SCREENING_SOUBOR = os.path.join(ROOT_DIR, 'data', 'screening.xlsx')
POMOCNA_SOUBOR   = os.path.join(ROOT_DIR, 'data', 'pomocna_data_no_lookahead_analyza.xlsx')
OUTPUT_FILE      = os.path.join(ROOT_DIR, 'vystupy', 'no_lookahead_output.xlsx')

PE_LIMIT  = 16.0
PB_LIMIT  = 2.0
R_BASE    = 0.09
G_BASE    = 0.02

ROKY = list(range(2004, 2025))

# Tituly s dividendou v EUR – kurz se načítá z pomocných dat
EUR_TITULY = ['Erste Group', 'VIG', 'PFNonwovens']

# Mapování názvů titulů ze screening.xlsx na jednotné názvy
NAZVY_MAP = {
    'ČEZ':                    'ČEZ',
    'O2':                     'O2 C.R.',
    'O2 C.R.':                'O2 C.R.',
    'Erste Group':            'Erste Group',
    'Komerční banka':         'Komerční banka',
    'Philip Morris':          'Philip Morris',
    'Vienna Insurance Group': 'VIG',
    'VIG':                    'VIG',
    'PFNonwovens':            'PFNonwovens',
    'MONETA':                 'MONETA',
    'Kofola':                 'Kofola',
    'COLTCZ':                 'COLT CZ',
    'COLZCZ':                 'COLT CZ',
}

# Mapování názvů na sheet v pomocných datech
SHEET_MAP = {
    'ČEZ':            'ČEZ',
    'O2 C.R.':        'O2 CR',
    'Erste Group':    'Erste Group',
    'Komerční banka': 'KB',
    'Philip Morris':  'Phillip Morris',
    'VIG':            'VIG',
    'PFNonwovens':    'PFNONWOVENS',
    'MONETA':         'MONETA',
    'Kofola':         'KOFOLA',
    'COLT CZ':        'COLTCZ',
}

# ==========================================
# NAČTENÍ GROUND TRUTH ZE SCREENING.XLSX
# ==========================================

def nacti_screening_ground_truth():
    """
    Načte výsledky screeningu ze screening.xlsx (ground truth z BP).
    Vrátí dict: {rok: {titul: {'pe': ..., 'pb': ..., 'ddm': ...,
                               'cena': ..., 'vysledek': bool}}}
    """
    data = {}
    for rok in ROKY:
        df = pd.read_excel(SCREENING_SOUBOR, sheet_name=str(rok))
        df.columns = ['Titul', 'PE', 'PB', 'DDM', 'Cena', 'Vysledek']
        rok_data = {}
        for _, row in df.iterrows():
            titul_raw = str(row['Titul']).strip()
            if titul_raw in ('nan', '') or titul_raw.startswith('*'):
                continue
            nazev = NAZVY_MAP.get(titul_raw, titul_raw)
            try:    pe_v   = float(row['PE'])
            except: pe_v   = np.nan
            try:    pb_v   = float(row['PB'])
            except: pb_v   = np.nan
            try:    ddm_v  = float(row['DDM'])
            except: ddm_v  = np.nan
            try:    cena_v = float(row['Cena'])
            except: cena_v = np.nan
            vysledek = str(row['Vysledek']).strip().upper() == 'ANO'
            rok_data[nazev] = {
                'pe': pe_v, 'pb': pb_v, 'ddm': ddm_v,
                'cena': cena_v, 'vysledek': vysledek
            }
        data[rok] = rok_data
    return data

ground_truth = nacti_screening_ground_truth()

# Základní (look-ahead) portfolio z BP
portfolio_bp = {
    rok: sorted([t for t, d in ground_truth[rok].items() if d['vysledek']])
    for rok in ROKY
}

# ==========================================
# NAČTENÍ DIVIDEND A KURZŮ Z POMOCNÝCH DAT
# ==========================================

def nacti_data_titulu():
    """
    Pro každý titul načte z pomocných dat:
      - časovou řadu dividend (řádek obsahující 'dividenda')
      - časovou řadu kurzů CZK/EUR (řádek obsahující 'kurz') – jen pro EUR tituly
    Vrátí dict: {nazev: {'dividendy': Series, 'kurzy': Series nebo None}}
    """
    vysledek = {}
    for nazev, sheet in SHEET_MAP.items():
        try:
            df = pd.read_excel(POMOCNA_SOUBOR, sheet_name=sheet)
            df.set_index(df.columns[0], inplace=True)

            div_serie  = None
            kurz_serie = None

            for idx in df.index:
                idx_lower = str(idx).lower()
                # Pro EUR tituly hledáme dividendu v EUR (ne CZK přepočet)
                if nazev in EUR_TITULY:
                    if 'dividenda eur' in idx_lower or \
                       ('dividenda' in idx_lower and 'czk' not in idx_lower and div_serie is None):
                        div_serie = df.loc[idx]
                else:
                    if 'dividenda' in idx_lower and div_serie is None:
                        div_serie = df.loc[idx]

                if 'kurz' in idx_lower and kurz_serie is None:
                    kurz_serie = df.loc[idx]

            vysledek[nazev] = {
                'dividendy': div_serie,
                'kurzy':     kurz_serie if nazev in EUR_TITULY else None
            }
        except Exception as e:
            print(f'  VAROVÁNÍ: Nepodařilo se načíst data pro {nazev}: {e}')
    return vysledek

data_titulu = nacti_data_titulu()


def get_dividenda_prev_czk(nazev, rok):
    """
    Vrátí dividendu z předchozího roku v CZK (no look-ahead varianta).
    Aplikuje pravidlo BP:
      - jednorázový výpadek (D=0 v roce t-1) → použije D z roku t-2 * (1+g)
      - dvojí výpadek → NaN (titul nevhodný)
    Pro EUR tituly převede historickým kurzem z dat.
    """
    if nazev not in data_titulu:
        return np.nan

    div_serie  = data_titulu[nazev]['dividendy']
    kurz_serie = data_titulu[nazev]['kurzy']

    if div_serie is None:
        return np.nan

    col_prev  = f'{rok-2}/{rok-1}'
    col_prev2 = f'{rok-3}/{rok-2}'

    def safe_float(serie, col):
        try:
            v = serie.get(col, np.nan)
            return float(v) if v is not None else np.nan
        except:
            return np.nan

    d0 = safe_float(div_serie, col_prev)

    # Jednorázový výpadek
    if pd.isna(d0) or d0 == 0:
        prev2 = safe_float(div_serie, col_prev2)
        d0 = np.nan if (pd.isna(prev2) or prev2 == 0) else prev2 * (1 + G_BASE)

    if pd.isna(d0) or d0 == 0:
        return np.nan

    # Převod EUR → CZK historickým kurzem
    if nazev in EUR_TITULY and kurz_serie is not None:
        kurz = safe_float(kurz_serie, col_prev)
        if pd.isna(kurz) or kurz == 0:
            # Záloha: kurz z předchozího dostupného roku
            kurz = safe_float(kurz_serie, col_prev2)
        if pd.isna(kurz) or kurz == 0:
            return np.nan
        return d0 * kurz

    return d0

# ==========================================
# NO LOOK-AHEAD PORTFOLIO
# ==========================================

def vypocti_nola_portfolio():
    """
    No look-ahead portfolio:
      - P/E a P/B z ground truth (BP) – nezměněno
      - DDM přepočítáno z dividendy předchozího roku s historickým kurzem
    """
    portfolio = {}
    for rok in ROKY:
        tituly = []
        for nazev, d in ground_truth[rok].items():
            pe_v = d['pe']
            pb_v = d['pb']
            cena = d['cena']

            d0_czk = get_dividenda_prev_czk(nazev, rok)
            if not pd.isna(d0_czk) and d0_czk > 0:
                d1     = d0_czk * (1 + G_BASE)
                ddm_vh = d1 / (R_BASE - G_BASE)
            else:
                ddm_vh = np.nan

            kpe  = not pd.isna(pe_v)  and pe_v > 0 and pe_v < PE_LIMIT
            kpb  = not pd.isna(pb_v)  and pb_v < PB_LIMIT
            kddm = not pd.isna(ddm_vh) and not pd.isna(cena) and ddm_vh > cena

            if sum([kpe, kpb, kddm]) >= 2:
                tituly.append(nazev)
        portfolio[rok] = sorted(tituly)
    return portfolio

portfolio_nola = vypocti_nola_portfolio()

# ==========================================
# TISK VÝSLEDKŮ
# ==========================================

print("=" * 105)
print("LOOK-AHEAD (BP) VS. NO LOOK-AHEAD – srovnání složení portfolia")
print("Základní (BP):  skutečná dividenda daného roku (correct foresight)")
print("No look-ahead:  dividenda z předchozího roku + historický kurz CZK/EUR")
print("=" * 105)
print(f"\n{'Rok':<6} {'Look-ahead (BP)':<45} {'No look-ahead':<45} {'Rozdíl'}")
print("─" * 105)

celkem         = 0
roky_s_rozd    = 0
radky_export   = []

for rok in ROKY:
    s1 = set(portfolio_bp[rok])
    s2 = set(portfolio_nola[rok])
    pridane  = s2 - s1
    odebrany = s1 - s2
    diff     = len(pridane) + len(odebrany)
    celkem  += diff
    if diff > 0:
        roky_s_rozd += 1
    flag     = ' <--' if diff > 0 else ''
    la_str   = ', '.join(sorted(s1)) if s1 else '—'
    nola_str = ', '.join(sorted(s2)) if s2 else '—'
    print(f"{rok:<6} {la_str:<45} {nola_str:<45} +{len(pridane)}/-{len(odebrany)}{flag}")
    radky_export.append({
        'Rok':              rok,
        'Look-ahead (BP)':  la_str,
        'No look-ahead':    nola_str,
        'Přidáno (NLA)':    ', '.join(sorted(pridane))  if pridane  else '—',
        'Odebráno (NLA)':   ', '.join(sorted(odebrany)) if odebrany else '—',
        'Počet změn':       diff,
    })

print(f"\nCelkový počet změněných titulů: {celkem}")
print(f"Průměrná změna za rok:          {celkem / len(ROKY):.2f}")
print(f"Roky s rozdílem:                {roky_s_rozd} z {len(ROKY)}")
print(f"Roky bez rozdílu:               {len(ROKY) - roky_s_rozd} z {len(ROKY)}")

# ==========================================
# EXPORT DO EXCELU
# ==========================================

souhrn = [{
    'Ukazatel':  'Celkový počet změněných titulů',
    'Hodnota':    celkem,
}, {
    'Ukazatel':  'Průměrná změna za rok',
    'Hodnota':    round(celkem / len(ROKY), 2),
}, {
    'Ukazatel':  'Roky s rozdílem',
    'Hodnota':   f'{roky_s_rozd} z {len(ROKY)}',
}, {
    'Ukazatel':  'Roky bez rozdílu',
    'Hodnota':   f'{len(ROKY) - roky_s_rozd} z {len(ROKY)}',
}]

with pd.ExcelWriter(OUTPUT_FILE, engine='openpyxl') as writer:
    pd.DataFrame(radky_export).to_excel(writer, sheet_name='Srovnání', index=False)
    pd.DataFrame(souhrn).to_excel(writer,       sheet_name='Souhrn',   index=False)

print(f"\n✓ Exportováno do: {OUTPUT_FILE}")
print("Hotovo.")
