# Časování nákupu akciových titulů na Burze cenných papírů Praha

[cite_start]Tento repozitář obsahuje zdrojové kódy a datové podklady pro bakalářskou práci na téma časování trhu na BCPP[cite: 512].

## Popis projektu
[cite_start]Cílem práce je otestovat aktivní investiční strategii, která propojuje fundamentální analýzu s kalendářními anomáliemi[cite: 524]. Strategie využívá:
* [cite_start]**Fundamentální screening:** $P/E < 16,0$, $P/B < 2,0$ a Gordonův růstový model (DDM)[cite: 797, 801, 806].
* [cite_start]**Kalendářní efekty:** Lednový, květnový a pondělní efekt[cite: 819].
* [cite_start]**Backtesting:** Testování probíhá na datech z Prime Marketu BCPP za období 2004–2024[cite: 772, 775].

## Obsah repozitáře
* [cite_start]`simulace_strategie.py`: Hlavní výpočetní skript v jazyce Python[cite: 871].
* [cite_start]`obchody.xlsx`: Vstupní datový soubor s historickými kurzy a fundamenty[cite: 873].
* [cite_start]`vysledna_vykonnost.xlsx`: Exportované výsledky roční výkonnosti a celkového zhodnocení[cite: 1239].

## Technické požadavky
* Python 3.x
* [cite_start]Knihovny: `pandas`, `openpyxl` [cite: 872]

## Autor
[cite_start][Tvé Jméno] (UJEP, Fakulta sociálně ekonomická) [cite: 509, 510]
