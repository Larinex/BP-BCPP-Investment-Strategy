import pandas as pd
import math
import os

cesta = os.path.dirname(os.path.abspath(__file__))

# ==========================================
# KONFIGURACE (0,35 %, min 40 Kč, max 1190 Kč)
# ==========================================
startovni_kapital = 1_000_000
poplatek_procento = 0.0035
min_poplatek = 40
max_poplatek = 1190
vstupni_soubor = os.path.join(cesta, 'obchody.xlsx')

# ==========================================
# NAČTENÍ A OPRAVA DAT
# ==========================================
try:
    df = pd.read_excel(vstupni_soubor)

    # Přejmenování sloupce Podíl -> Podil (kvůli diakritice)
    prejmenovani = {'Podíl': 'Podil'}
    df.rename(columns=prejmenovani, inplace=True)

    # Formátování datumu
    df['Datum'] = pd.to_datetime(df['Datum'], dayfirst=True)
    df = df.sort_values(by='Datum').reset_index(drop=True)

    # Prázdný podíl = 1 = 100 %
    df['Podil'] = df['Podil'].fillna(1.0)

    # Sjednocení textu (Nákup / NAKUP / nakup -> NAKUP)
    df['Typ'] = df['Typ'].str.upper().str.strip()

    print(f"Data načtena: {len(df)} řádků.")
    print(df.head(3))
    print("-" * 50)

except FileNotFoundError:
    print(f"CHYBA: Soubor '{vstupni_soubor}' nebyl nalezen!")
    exit()
except Exception as e:
    print(f"CHYBA při načítání dat: {e}")
    exit()

# ==========================================
# POMOCNÉ FUNKCE
# ==========================================
def vypocitej_poplatek(objem):
    v = objem * poplatek_procento
    return min(max(v, min_poplatek), max_poplatek)

def ocen_portfolio(portfolio, dnesni_pokyny, den):
    """Vrátí celkovou hodnotu akcií v portfoliu k danému dni."""
    hodnota_akcii = 0
    for titul, ks in portfolio.items():
        if ks > 0:
            cena_rows = dnesni_pokyny[dnesni_pokyny['Titul'] == titul]
            if not cena_rows.empty:
                cena = cena_rows.iloc[0]['Cena']
                hodnota_akcii += ks * cena
            else:
                print(f"  VAROVÁNÍ: Chybí cena pro '{titul}' k {den.date()}!")
    return hodnota_akcii

# ==========================================
# SIMULACE
# ==========================================
hotovost = startovni_kapital
portfolio = {}
seznam_transakci = []
rocni_prehled = []      # [{rok, datum_zacatek, majetek_zacatek, datum_konec, majetek_konec}]

# Pomocné proměnné pro párování ZACATEK <-> INFO v rámci roku
zacatek_roku = {}       # {rok: {'datum': ..., 'majetek': ...}}

unikatni_dny = df['Datum'].unique()

for den in unikatni_dny:
    dnesni_pokyny = df[df['Datum'] == den]
    rok = pd.Timestamp(den).year

    # -------------------------------------------------------
    # 1. ZACATEK – ocenění na první obchodní den roku
    #    (žádný nákup/prodej, jen zaznamenání hodnoty)
    # -------------------------------------------------------
    if 'ZACATEK' in dnesni_pokyny['Typ'].values:
        hodnota_akcii = ocen_portfolio(portfolio, dnesni_pokyny, pd.Timestamp(den))
        majetek = round(hotovost + hodnota_akcii, 2)
        zacatek_roku[rok] = {'datum': pd.Timestamp(den).date(), 'majetek': majetek}
        print(f"ZAČÁTEK {rok} ({pd.Timestamp(den).date()}): {majetek:,.0f} Kč")

    # -------------------------------------------------------
    # 2. PRODEJE
    # -------------------------------------------------------
    prodeje = dnesni_pokyny[dnesni_pokyny['Typ'] == 'PRODEJ']

    for index, row in prodeje.iterrows():
        titul = row['Titul']
        cena = row['Cena']
        podil = row['Podil']

        aktualni_pocet = portfolio.get(titul, 0)

        if aktualni_pocet > 0:
            if podil >= 0.99:
                ks_k_prodeji = aktualni_pocet
            else:
                ks_k_prodeji = math.floor(aktualni_pocet * podil)

            if ks_k_prodeji > 0:
                objem = ks_k_prodeji * cena
                poplatek = vypocitej_poplatek(objem)

                hotovost += (objem - poplatek)
                portfolio[titul] = aktualni_pocet - ks_k_prodeji

                typ_txt = "PRODEJ" if podil >= 0.99 else f"PRODEJ ({podil:.2f})"
                seznam_transakci.append({
                    'Datum': pd.Timestamp(den).date(),
                    'Typ': typ_txt,
                    'Titul': titul,
                    'Pocet_ks': ks_k_prodeji,
                    'Cena': cena,
                    'Objem_CZK': round(objem, 2),
                    'Poplatek': round(poplatek, 2),
                    'Zustatek_Hotovosti': round(hotovost, 2)
                })

    # -------------------------------------------------------
    # 3. NÁKUPY – celá dostupná hotovost se rovnoměrně rozdělí
    # -------------------------------------------------------
    nakupy = dnesni_pokyny[dnesni_pokyny['Typ'] == 'NAKUP']

    if not nakupy.empty:
        pocet_titulu = len(nakupy)
        rozpocet_na_akcii = hotovost / pocet_titulu

        for index, row in nakupy.iterrows():
            titul = row['Titul']
            cena = row['Cena']

            odhad_ks = math.floor((rozpocet_na_akcii - min_poplatek) / cena)
            if odhad_ks < 0:
                odhad_ks = 0

            while True:
                objem = odhad_ks * cena
                poplatek = vypocitej_poplatek(objem)
                if (objem + poplatek) <= rozpocet_na_akcii:
                    break
                odhad_ks -= 1
                if odhad_ks <= 0:
                    break

            if odhad_ks > 0:
                hotovost -= (objem + poplatek)
                portfolio[titul] = portfolio.get(titul, 0) + odhad_ks

                seznam_transakci.append({
                    'Datum': pd.Timestamp(den).date(),
                    'Typ': 'NAKUP',
                    'Titul': titul,
                    'Pocet_ks': odhad_ks,
                    'Cena': cena,
                    'Objem_CZK': round(objem, 2),
                    'Poplatek': round(poplatek, 2),
                    'Zustatek_Hotovosti': round(hotovost, 2)
                })

    # -------------------------------------------------------
    # 4. INFO – ocenění na konci roku (výsledek roku)
    # -------------------------------------------------------
    je_info = 'INFO' in dnesni_pokyny['Typ'].values
    je_posledni_den = (den == unikatni_dny[-1])

    if je_info or je_posledni_den:
        hodnota_akcii = ocen_portfolio(portfolio, dnesni_pokyny, pd.Timestamp(den))
        majetek_konec = round(hotovost + hodnota_akcii, 2)

        print(f"KONEC   {rok} ({pd.Timestamp(den).date()}): {majetek_konec:,.0f} Kč  "
              f"(Cash: {hotovost:,.0f} | Akcie: {hodnota_akcii:,.0f})")

        # Výpočet ročního výnosu – porovnání s ZACATEK stejného roku
        if rok in zacatek_roku:
            majetek_zacatek = zacatek_roku[rok]['majetek']
            vynos_pct = round((majetek_konec - majetek_zacatek) / majetek_zacatek * 100, 2)
        else:
            # Pro první rok (2004) porovnáváme se startovním kapitálem
            majetek_zacatek = startovni_kapital
            vynos_pct = round((majetek_konec - startovni_kapital) / startovni_kapital * 100, 2)

        rocni_prehled.append({
            'Rok': rok,
            'Datum_zacatek': zacatek_roku[rok]['datum'] if rok in zacatek_roku else '2004-01-05',
            'Majetek_zacatek': majetek_zacatek,
            'Datum_konec': pd.Timestamp(den).date(),
            'Majetek_konec': majetek_konec,
            'Vynos_Pct': vynos_pct
        })

# ==========================================
# EXPORT
# ==========================================
# 1. Detailní transakce
df_t = pd.DataFrame(seznam_transakci)
vystup_transakce = os.path.join(cesta, "detailni_transakce.xlsx")
df_t.to_excel(vystup_transakce, index=False)
print(f"\nTransakce uloženy: {vystup_transakce}")

# 2. Roční výkonnost
if rocni_prehled:
    df_v = pd.DataFrame(rocni_prehled)

    # Celkové zhodnocení od počátku (vs startovní kapitál)
    df_v['Celkove_zhodnoceni_Pct'] = round(
        (df_v['Majetek_konec'] - startovni_kapital) / startovni_kapital * 100, 2
    )

    vystup_vykonnost = os.path.join(cesta, "vysledna_vykonnost.xlsx")
    df_v.to_excel(vystup_vykonnost, index=False)
    print(f"Výkonnost uložena:  {vystup_vykonnost}")

    print("\n" + "=" * 60)
    print(f"{'Rok':<6} {'Začátek':>14} {'Konec':>14} {'Výnos':>8}")
    print("=" * 60)
    for _, r in df_v.iterrows():
        print(f"{int(r['Rok']):<6} {r['Majetek_zacatek']:>14,.0f} {r['Majetek_konec']:>14,.0f} {r['Vynos_Pct']:>7.2f} %")
    print("=" * 60)
else:
    print("Žádná data pro vyhodnocení.")