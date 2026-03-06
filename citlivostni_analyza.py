import pandas as pd
import numpy as np
import warnings
import os
warnings.filterwarnings('ignore')

# ==========================================
# KONFIGURACE
# ==========================================

SOUBOR = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'pomocná data_pro_citlivostni_analyzu.xlsx')
# Look-ahead: skutečná dividenda daného roku (col_cur)
# No look-ahead: dividenda z předchozího roku (col_prev) — stejný soubor, jiný sloupec
SOUBOR_BP   = SOUBOR
SOUBOR_NOLA = SOUBOR
PE_LIMIT = 16.0
PB_LIMIT = 2.0

# Scénáře citlivostní analýzy: (r, g) dvojice
SCENARE = {
    'Základní (r=9%, g=2%)':    (0.09, 0.02),
    'Konzervativní (r=10%, g=2%)': (0.10, 0.02),
    'Optimistický (r=8%, g=2%)':  (0.08, 0.02),
    'Vyšší růst (r=9%, g=3%)':   (0.09, 0.03),
    'Nízké sazby (r=7%, g=2%)':  (0.07, 0.02),
}

# Společnosti s EUR dividendou - kurz EUR/CZK průměr pro přepočet
# Použijeme fixní přepočet jen pro DDM screening (VH vs TC jsou pak srovnatelné)
# Pro Erste a VIG a PFNONWOVENS je dividenda v EUR, cena v CZK
# => DDM vnitřní hodnotu přepočteme kurzem 25 CZK/EUR (konzervativní průměr)
EUR_CZK = 25.0

# Společnosti kde je dividenda v EUR
EUR_DIVIDENDA = ['Erste Group', 'VIG', 'PFNONWOVENS']

# ==========================================
# NAČTENÍ DAT
# ==========================================

sheets = ['ČEZ', 'Erste Group', 'KB', 'VIG', 'MONETA', 'KOFOLA', 'COLTCZ', 'O2 CR', 'Phillip Morris', 'PFNONWOVENS']

# Zkrácené názvy pro tabulku
NAZVY = {
    'ČEZ': 'ČEZ',
    'Erste Group': 'Erste Group',
    'KB': 'Komerční banka',
    'VIG': 'VIG',
    'MONETA': 'MONETA',
    'KOFOLA': 'Kofola',
    'COLTCZ': 'COLT CZ',
    'O2 CR': 'O2 C.R.',
    'Phillip Morris': 'Philip Morris',
    'PFNONWOVENS': 'PFNonwovens',
}

def nacti_data(sheet):
    """Načte data pro jednu společnost a vrátí dict řádkových sérií."""
    df = pd.read_excel(SOUBOR, sheet_name=sheet)
    df.set_index(df.columns[0], inplace=True)
    return df

def najdi_radek(df, klicova_slova):
    """Najde řádek podle klíčových slov (case-insensitive)."""
    for idx in df.index:
        if any(kw.lower() in str(idx).lower() for kw in klicova_slova):
            return df.loc[idx]
    return None

def aplikuj_pravidlo_dividendy(dividendy_serie, g):
    """
    Aplikuje pravidlo z BP pro ošetření nulových/chybějících dividend:
    - Standardní: použij D0 z dat (= skutečně vyplacená dividenda v daném roce)
    - Jednorázový výpadek (1 rok D=0): nahraď D_prev * (1+g)
    - Trvalý výpadek (2+ roky D=0 za sebou): titul vyřadit (vrátí NaN)
    
    Vrací sérii D1 (očekávané dividendy) pro každý rok.
    Pozor: D1 = D0 * (1+g) — dividenda příštího období pro Gordon model.
    """
    roky = dividendy_serie.index.tolist()
    d0_upravene = {}

    for i, rok in enumerate(roky):
        d0 = dividendy_serie[rok]

        if pd.isna(d0) or d0 == 0:
            # Zjisti, zda jde o jednorázový nebo trvalý výpadek
            # Podívej se na předchozí rok
            prev_d = dividendy_serie[roky[i-1]] if i > 0 else None
            next_d = dividendy_serie[roky[i+1]] if i < len(roky)-1 else None

            # Zkontroluj, zda předcházející rok byl také 0
            prev_zero = (prev_d is None or pd.isna(prev_d) or prev_d == 0)

            if prev_zero:
                # Trvalý výpadek (2+ roky) => titul nevhodný
                d0_upravene[rok] = np.nan
            else:
                # Jednorázový výpadek => nahraď prev * (1+g)
                d0_upravene[rok] = prev_d * (1 + g)
        else:
            d0_upravene[rok] = d0

    return pd.Series(d0_upravene)


# ==========================================
# HLAVNÍ VÝPOČET SCREENINGU
# ==========================================

def vypocti_screening(r, g, label):
    """
    Pro dané r a g provede hodnotový screening pro všechny roky a společnosti.
    Vrátí DataFrame s výsledky podobným Tabulce 3 z BP.
    """
    vsechna_data = {}

    for sheet in sheets:
        df = nacti_data(sheet)
        nazev = NAZVY[sheet]

        cena = najdi_radek(df, ['cena akcie'])
        pe = najdi_radek(df, ['p/e'])
        pb = najdi_radek(df, ['p/b'])
        dividenda_raw = najdi_radek(df, ['dividenda'])
        
        if dividenda_raw is None or cena is None:
            continue

        # Ošetři dividendy podle pravidel BP
        div_upravene = aplikuj_pravidlo_dividendy(dividenda_raw, g)

        for rok_key in df.columns:
            # Rok pro který se sestavuje portfolio (druhá část "YYYY/YYYY")
            rok_portfolio = int(str(rok_key).split('/')[1])

            pe_val = pe[rok_key] if pe is not None else np.nan
            pb_val = pb[rok_key] if pb is not None else np.nan
            cena_val = cena[rok_key]
            d0 = div_upravene.get(rok_key, np.nan)

            # Výpočet DDM vnitřní hodnoty
            if pd.isna(d0) or d0 == 0:
                ddm_vh = np.nan
            else:
                d1 = d0 * (1 + g)
                # Pro EUR dividendu přepočti na CZK
                if sheet in EUR_DIVIDENDA:
                    d1_czk = d1 * EUR_CZK
                else:
                    d1_czk = d1
                ddm_vh = d1_czk / (r - g)

            # Převod na číslo (ošetření hodnot jako "X", "N/A" atd.)
            try:
                pe_val = float(pe_val)
            except (ValueError, TypeError):
                pe_val = np.nan
            try:
                pb_val = float(pb_val)
            except (ValueError, TypeError):
                pb_val = np.nan

            # Vyhodnocení kritérií
            krit_pe = (not pd.isna(pe_val)) and (pe_val > 0) and (pe_val < PE_LIMIT)
            krit_pb = (not pd.isna(pb_val)) and (pb_val < PB_LIMIT)
            krit_ddm = (not pd.isna(ddm_vh)) and (not pd.isna(cena_val)) and (ddm_vh > cena_val)

            splneno = sum([krit_pe, krit_pb, krit_ddm])
            vysledek = splneno >= 2

            key = (rok_portfolio, nazev)
            vsechna_data[key] = {
                'Rok': rok_portfolio,
                'Titul': nazev,
                'Cena': round(cena_val, 2) if not pd.isna(cena_val) else np.nan,
                'P/E': round(pe_val, 2) if not pd.isna(pe_val) else np.nan,
                'P/B': round(pb_val, 2) if not pd.isna(pb_val) else np.nan,
                'D0': round(d0, 2) if not pd.isna(d0) else np.nan,
                'DDM_VH': round(ddm_vh, 0) if not pd.isna(ddm_vh) else np.nan,
                'P/E OK': krit_pe,
                'P/B OK': krit_pb,
                'DDM OK': krit_ddm,
                'Splněno': splneno,
                'V portfoliu': vysledek,
            }

    return pd.DataFrame(vsechna_data.values())


# ==========================================
# VÝPOČET PRO VŠECHNY SCÉNÁŘE
# ==========================================

print("=" * 70)
print("CITLIVOSTNÍ ANALÝZA DDM — VLIV PARAMETRŮ r A g NA SLOŽENÍ PORTFOLIA")
print("=" * 70)

# Sestavíme souhrnnou tabulku jako Tabulka 3 z BP
roky = list(range(2004, 2025))
vsechny_tituly = list(NAZVY.values())

vysledky_scenaru = {}

for label, (r, g) in SCENARE.items():
    df_screen = vypocti_screening(r, g, label)
    
    # Pro každý rok sestavíme portfolio
    portfolio_roky = {}
    for rok in roky:
        df_rok = df_screen[df_screen['Rok'] == rok]
        tituly_v_portfoliu = df_rok[df_rok['V portfoliu'] == True]['Titul'].tolist()
        portfolio_roky[rok] = sorted(tituly_v_portfoliu)
    
    vysledky_scenaru[label] = portfolio_roky


# ==========================================
# TISKNUTÍ VÝSLEDKŮ — Tabulka 3 styl
# ==========================================

print("\n" + "=" * 70)
print("SROVNÁNÍ SLOŽENÍ PORTFOLIA PODLE SCÉNÁŘŮ")
print("=" * 70)

for label, (r, g) in SCENARE.items():
    print(f"\n{'─'*70}")
    print(f"SCÉNÁŘ: {label}")
    print(f"{'─'*70}")
    print(f"{'Rok':<6} {'Počet':>5}  Zahrnuté tituly")
    print(f"{'─'*70}")
    
    portfolio = vysledky_scenaru[label]
    for rok in roky:
        tituly = portfolio[rok]
        pocet = len(tituly)
        tituly_str = ', '.join(tituly) if tituly else '— (žádný titul nesplnil kritéria)'
        print(f"{rok:<6} {pocet:>5}  {tituly_str}")


# ==========================================
# SROVNÁVACÍ TABULKA — rozdíly mezi scénáři
# ==========================================

print("\n" + "=" * 70)
print("SROVNÁVACÍ TABULKA — POČET TITULŮ V PORTFOLIU PODLE SCÉNÁŘE")
print("=" * 70)

zakladni_label = 'Základní (r=9%, g=2%)'
header = f"{'Rok':<6}" + "".join(f"{'r9g2':>6}" if i == 0 else f"{'r10g2':>6}" if i == 1 
                                  else f"{'r8g2':>6}" if i == 2 else f"{'r9g3':>6}" if i == 3 
                                  else f"{'r7g2':>6}" for i, _ in enumerate(SCENARE))
print(header)
print("─" * (6 + 6 * len(SCENARE)))

for rok in roky:
    radek = f"{rok:<6}"
    for label in SCENARE:
        pocet = len(vysledky_scenaru[label][rok])
        radek += f"{pocet:>6}"
    print(radek)


# ==========================================
# ROZDÍLY OPROTI ZÁKLADNÍMU SCÉNÁŘI
# ==========================================

print("\n" + "=" * 70)
print("ROKY S ODLIŠNÝM SLOŽENÍM PORTFOLIA OPROTI ZÁKLADNÍMU SCÉNÁŘI")
print("=" * 70)

zakladni = vysledky_scenaru[zakladni_label]

for label, (r, g) in SCENARE.items():
    if label == zakladni_label:
        continue
    
    print(f"\nScénář: {label}")
    odlisne = False
    for rok in roky:
        z = set(zakladni[rok])
        s = set(vysledky_scenaru[label][rok])
        pridane = s - z
        odebrané = z - s
        if pridane or odebrané:
            odlisne = True
            zprava = f"  {rok}: "
            if pridane:
                zprava += f"PŘIDÁNO: {', '.join(sorted(pridane))}  "
            if odebrané:
                zprava += f"ODEBRÁNO: {', '.join(sorted(odebrané))}"
            print(zprava)
    if not odlisne:
        print("  Žádné rozdíly — složení portfolia identické se základním scénářem.")


# ==========================================
# LOOK-AHEAD VS. NO LOOK-AHEAD POROVNÁNÍ
# (pouze pro základní scénář r=9%, g=2%)
# ==========================================

print("\n" + "=" * 70)
print("ROBUSTNOSTNÍ TEST — LOOK-AHEAD VS. NO LOOK-AHEAD (r=9%, g=2%)")
print("Základní scénář používá skutečnou dividendu daného roku (look-ahead).")
print("No look-ahead varianta používá dividendu z předchozího roku.")
print("=" * 70)

def screening_nola(soubor_la, soubor_nola):
    """Screening bez look-ahead biasu.
    - soubor_la:   původní BP data (skutečné dividendy) => look-ahead výsledky
    - soubor_nola: rozšířená data s rokem 2002/2003 => no look-ahead výsledky
    """
    portfolio_la   = {}
    portfolio_nola = {}
    r, g = 0.09, 0.02

    for rok in roky:
        tituly_la   = []
        tituly_nola = []

        for sheet in sheets:
            nazev = NAZVY[sheet]

            # --- LOOK-AHEAD: původní BP soubor, skutečná dividenda daného roku ---
            try:
                df_la = pd.read_excel(soubor_la, sheet_name=sheet)
                df_la.set_index(df_la.columns[0], inplace=True)
                cena_row = najdi_radek(df_la, ['cena akcie'])
                pe_row   = najdi_radek(df_la, ['p/e'])
                pb_row   = najdi_radek(df_la, ['p/b'])
                div_row  = najdi_radek(df_la, ['dividenda'])
                col_cur  = f'{rok-1}/{rok}'
                col_prev = f'{rok-2}/{rok-1}'

                if div_row is not None and cena_row is not None and col_cur in df_la.columns:
                    cena = cena_row.get(col_cur, np.nan)
                    try: pe_v = float(pe_row.get(col_cur, np.nan)) if pe_row is not None else np.nan
                    except: pe_v = np.nan
                    try: pb_v = float(pb_row.get(col_cur, np.nan)) if pb_row is not None else np.nan
                    except: pb_v = np.nan

                    try: d0 = float(div_row.get(col_cur, np.nan))
                    except: d0 = np.nan
                    if pd.isna(d0) or d0 == 0:
                        try: prev_d = float(div_row.get(col_prev, np.nan))
                        except: prev_d = np.nan
                        d0 = np.nan if (pd.isna(prev_d) or prev_d == 0) else prev_d * (1 + g)

                    if not (pd.isna(d0) or d0 == 0):
                        d1 = d0 * (1 + g)
                        if sheet in EUR_DIVIDENDA:
                            d1 = d1 * EUR_CZK
                        ddm_vh = d1 / (r - g)
                    else:
                        ddm_vh = np.nan

                    kpe  = not pd.isna(pe_v)   and pe_v > 0 and pe_v < PE_LIMIT
                    kpb  = not pd.isna(pb_v)   and pb_v < PB_LIMIT
                    kddm = not pd.isna(ddm_vh) and not pd.isna(cena) and ddm_vh > float(cena)
                    if sum([kpe, kpb, kddm]) >= 2:
                        tituly_la.append(nazev)
            except Exception:
                pass

            # --- NO LOOK-AHEAD: rozšířený soubor, dividenda z předchozího roku ---
            try:
                df_nola = pd.read_excel(soubor_nola, sheet_name=sheet)
                df_nola.set_index(df_nola.columns[0], inplace=True)
                cena_row = najdi_radek(df_nola, ['cena akcie'])
                pe_row   = najdi_radek(df_nola, ['p/e'])
                pb_row   = najdi_radek(df_nola, ['p/b'])
                div_row  = najdi_radek(df_nola, ['dividenda'])
                col_cur   = f'{rok-1}/{rok}'
                col_prev  = f'{rok-2}/{rok-1}'
                col_prev2 = f'{rok-3}/{rok-2}'

                if div_row is not None and cena_row is not None and col_cur in df_nola.columns:
                    cena = cena_row.get(col_cur, np.nan)
                    try: pe_v = float(pe_row.get(col_cur, np.nan)) if pe_row is not None else np.nan
                    except: pe_v = np.nan
                    try: pb_v = float(pb_row.get(col_cur, np.nan)) if pb_row is not None else np.nan
                    except: pb_v = np.nan

                    try: d0 = float(div_row.get(col_prev, np.nan))
                    except: d0 = np.nan
                    if pd.isna(d0) or d0 == 0:
                        try: prev_d = float(div_row.get(col_prev2, np.nan))
                        except: prev_d = np.nan
                        d0 = np.nan if (pd.isna(prev_d) or prev_d == 0) else prev_d * (1 + g)

                    if not (pd.isna(d0) or d0 == 0):
                        d1 = d0 * (1 + g)
                        if sheet in EUR_DIVIDENDA:
                            d1 = d1 * EUR_CZK
                        ddm_vh = d1 / (r - g)
                    else:
                        ddm_vh = np.nan

                    kpe  = not pd.isna(pe_v)   and pe_v > 0 and pe_v < PE_LIMIT
                    kpb  = not pd.isna(pb_v)   and pb_v < PB_LIMIT
                    kddm = not pd.isna(ddm_vh) and not pd.isna(cena) and ddm_vh > float(cena)
                    if sum([kpe, kpb, kddm]) >= 2:
                        tituly_nola.append(nazev)
            except Exception:
                pass

        portfolio_la[rok]   = sorted(tituly_la)
        portfolio_nola[rok] = sorted(tituly_nola)

    return portfolio_la, portfolio_nola

nla_la, nla = screening_nola(SOUBOR_BP, SOUBOR_NOLA)
la = nla_la  # look-ahead z BP dat (přepíše základní scénář pro toto srovnání)

print(f"\n{'Rok':<6} {'Look-ahead (BP)':<52} {'No look-ahead':<52} Rozdíl")
print("─" * 120)

celkem_nola = 0
roky_rozdil_nola = 0
nola_radky = []

for rok in roky:
    s1 = set(la[rok])
    s2 = set(nla[rok])
    pridane  = s2 - s1
    odebrané = s1 - s2
    diff = len(pridane) + len(odebrané)
    celkem_nola += diff
    if diff > 0:
        roky_rozdil_nola += 1
    flag = ' <--' if diff > 0 else ''
    print(f"{rok:<6} {str(sorted(s1)):<52} {str(sorted(s2)):<52} +{len(pridane)}/-{len(odebrané)}{flag}")
    nola_radky.append({
        'Rok': rok,
        'Look-ahead (BP)': ', '.join(sorted(s1)) if s1 else '—',
        'No look-ahead':   ', '.join(sorted(s2)) if s2 else '—',
        'Přidáno (NLA)':   ', '.join(sorted(pridane))  if pridane  else '—',
        'Odebráno (NLA)':  ', '.join(sorted(odebrané)) if odebrané else '—',
        'Počet změn':      diff,
    })

print(f"\nCelkový počet změněných titulů: {celkem_nola} přes {len(roky)} let")
print(f"Průměrná změna za rok:          {celkem_nola/len(roky):.2f} titulů")
print(f"Roky s rozdílem:                {roky_rozdil_nola} z {len(roky)}")
print(f"Roky bez rozdílu:               {len(roky)-roky_rozdil_nola} z {len(roky)}")

# ==========================================
# EXPORT DO EXCELU
# ==========================================

vystup = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'citlivostni_analyza_DDM.xlsx')

with pd.ExcelWriter(vystup, engine='openpyxl') as writer:
    
    # List 1: Srovnávací tabulka počtu titulů
    radky = []
    for rok in roky:
        radek = {'Rok': rok}
        for label in SCENARE:
            radek[label] = len(vysledky_scenaru[label][rok])
        radky.append(radek)
    df_pocty = pd.DataFrame(radky)
    df_pocty.to_excel(writer, sheet_name='Počty titulů', index=False)
    
    # List 2-6: Detailní složení portfolia pro každý scénář
    for label, (r, g) in SCENARE.items():
        radky2 = []
        for rok in roky:
            tituly = vysledky_scenaru[label][rok]
            radky2.append({
                'Rok': rok,
                'Počet titulů': len(tituly),
                'Zahrnuté tituly': ', '.join(tituly) if tituly else '—'
            })
        df_detail = pd.DataFrame(radky2)
        sheet_name = label[:31]  # Excel limit 31 znaků
        df_detail.to_excel(writer, sheet_name=sheet_name, index=False)
    
    # List 7: Rozdíly oproti základnímu scénáři
    diff_radky = []
    for label, (r, g) in SCENARE.items():
        if label == zakladni_label:
            continue
        for rok in roky:
            z = set(zakladni[rok])
            s = set(vysledky_scenaru[label][rok])
            pridane = s - z
            odebrané = z - s
            if pridane or odebrané:
                diff_radky.append({
                    'Scénář': label,
                    'Rok': rok,
                    'Přidáno oproti základnímu': ', '.join(sorted(pridane)) if pridane else '—',
                    'Odebráno oproti základnímu': ', '.join(sorted(odebrané)) if odebrané else '—',
                })
    
    if diff_radky:
        df_diff = pd.DataFrame(diff_radky)
        df_diff.to_excel(writer, sheet_name='Rozdíly vs základní', index=False)

    # List 8: Robustnostní test look-ahead vs. no look-ahead
    df_nola = pd.DataFrame(nola_radky)
    df_nola.to_excel(writer, sheet_name='Look-ahead vs No look-ahead', index=False)

print(f"\n✓ Výsledky exportovány do: {vystup}")
print("\nHotovo.")
