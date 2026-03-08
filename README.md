# Backtesting investiční strategie na BCPP (2004–2024)

Tento repozitář obsahuje kompletní metodiku, data a výpočetní algoritmy pro simulaci investiční strategie kombinující fundamentální analýzu s kalendářními anomáliemi na Burze cenných papírů Praha. Repozitář slouží jako příloha k bakalářské práci *Časování nákupu akciových titulů na Burze cenných papírů Praha* (FSE UJEP, 2026).

---

## Struktura repozitáře

```
├── data/                   – vstupní datové soubory
├── python_skripty/         – zdrojové kódy simulací
├── vystupy/                – výsledky připravené ke čtení bez Pythonu
├── README.md
└── LICENSE
```

---

## Popis skriptů

### `algoritmus_strategie_BP.py`
Hlavní simulační skript. Načítá předpřipravené obchodní příkazy a provádí backtest strategie za období 2004–2024.

- **Vstup:** `data/obchody.xlsx`
- **Výstup:** `vystupy/vysledna_vykonnost.xlsx`, `vystupy/detailni_transakce.xlsx`

Skript simuluje správu portfolia včetně rovnoměrného rozdělení kapitálu mezi tituly, rebalancování, reinvestice zisků a transakčních nákladů dle ceníku Fio banky (0,35 %, min. 40 Kč, max. 1 190 Kč). Roční výnos je počítán metodou konec roku N−1 → konec roku N, shodně s metodikou indexu PX.

---

### `citlivostni_analyza_r.py`
Citlivostní analýza parametru požadované výnosové míry *r* v Gordonově růstovém modelu (DDM). Testuje, jak se mění verdikt DDM (podhodnoceno / nepodhodnoceno) při nahrazení fixního *r* = 9 % dynamickým *r* odvozeným od výnosů 10letých českých státních dluhopisů a rizikové prémie dle Damodarana (2025).

- **Vstup:** `data/pomocna_data.xlsx`
- **Výstup:** `vystupy/sensitivity_output.xlsx` (listy: Detail, Souhrn, Změny)

Výnosy dluhopisů jsou hardcoded přímo ve skriptu jako záložní hodnoty – soubor `dluhopisy.xlsx` tedy není ke spuštění nutný.

---

### `no_lookahead_analyza.py`
Analýza no look-ahead (analýza bez předvídání budoucích dat). Ověřuje robustnost strategie vůči předpokladu správné prozíravosti (*correct foresight*) při výpočtu DDM. Místo skutečné dividendy vyplacené v roce T používá dividendu z předchozího roku navýšenou o *g*, čímž eliminuje závislost na budoucích datech. Pro EUR tituly (Erste Group, VIG, PFNonwovens) používá historický kurz CZK/EUR z prvního obchodního dne daného roku.

- **Vstup:** `data/screening.xlsx`, `data/pomocna_data_no_lookahead_analyza.xlsx`
- **Výstup:** `vystupy/no_lookahead_output.xlsx` (listy: Srovnání, Souhrn)

---

## Požadavky a instalace

Skripty vyžadují Python 3.8 nebo novější a následující knihovny:

```
pip install pandas openpyxl
```

---

## Spuštění

Pro správnou funkčnost je nutné stáhnout celý repozitář a zachovat jeho strukturu složek. Každý skript očekává vstupní data ve složce `data/` a výstupy ukládá do složky `vystupy/`, obě relativně ke složce `python_skripty/`.

```bash
# Stažení repozitáře
git clone https://github.com/Larinex/BP-BCPP-Investment-Strategy

# Přechod do složky se skripty
cd BP-BCPP-Investment-Strategy/python_skripty

# Spuštění hlavního algoritmu
python algoritmus_strategie_BP.py

# Spuštění citlivostní analýzy r
python citlivostni_analyza_r.py

# Spuštění no look-ahead analýzy
python no_lookahead_analyza.py
```

Výsledky jsou po spuštění dostupné ve složce `vystupy/`. Pokud nechcete spouštět skripty, předpočítané výstupy jsou již ve složce `vystupy/` k dispozici přímo.

---

## Autor

Nick Trefil – Fakulta sociálně ekonomická, UJEP, 2026
