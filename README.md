# Backtesting investiční strategie na BCPP (2004–2024)

Tento projekt obsahuje kompletní metodiku, data a výpočetní algoritmus pro simulaci investiční strategie kombinující fundamentální analýzu a kalendářní anomálie na pražské burze.

## Přehled strategie
Simulace testuje aktivní přístup k investování na segmentu Prime Market BCPP:
* **Fundamentální kritéria:** Výběr podhodnocených titulů pomocí ukazatelů $P/E < 16,0$, $P/B < 2,0$ a vnitřní hodnoty dle Gordonova růstového modelu (DDM).
* **Časování trhu:** Implementace lednového, květnového (Sell in May) a pondělního efektu pro optimalizaci vstupů a výstupů.
* **Období:** Historický backtest na datech z let 2004–2024.

## Struktura repozitáře
* `BP_kod.py`: Hlavní simulační skript v Pythonu zajišťující exekuci obchodů, rebalancování portfolia a výpočet poplatků.
* `screening.xlsx`: Detailní výsledky fundamentálního screeningu pro jednotlivé roky, určující složení portfolia.
* `pomocná data.xlsx`: Historické časové řady cen a fundamentálních údajů pro klíčové tituly (ČEZ, KB, Erste, Moneta atd.).
* `vysledna_vykonnost.xlsx`: Exportované výstupy simulace včetně ročních výnosů a celkového zhodnocení.

## Technické informace
Algoritmus je postaven na knihovně `pandas` a simuluje reálné tržní podmínky včetně transakčních nákladů (dle ceníku Fio banky). Pro správnou funkčnost skriptu je nutné zachovat relativní cesty k datovým souborům v rámci pracovní složky.

## Autor
Nick Trefil 
FSE UJEP
