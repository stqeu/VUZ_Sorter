# VUZ Sorter – triedenie dokumentov podľa Tag No

Nástroj priradí každý dokument ku **všetkým Tag No, ktoré obsahuje**. Dokument uloží **raz**
a v priečinkoch zariadení vytvorí iba **odkazy**. Nahrádza pôvodných ~14 jednorazových skriptov.

> **Stav:** demo (v0.1). OCR je ďalší krok – pozri [docs/ROADMAP.md](docs/ROADMAP.md).

## Rýchly štart

```powershell
pip install -r requirements.txt
python -m unittest discover testy      # 11 testov, musí byť OK
.\TEST.bat                             # skúšobný beh na demo dátach – nič nevytvorí, iba protokol
.\OSTRO.bat                            # ostrý beh – vystup\_dokumenty + vystup\zariadenia
```

Reálny projekt: skopíruj `config/demo.ini` → `config/rhc.local.ini`, uprav cesty a spusti
`.\TEST.bat --config config\rhc.local.ini`. Súbory `*.local.ini` sa do gitu nenahrávajú.

## Štruktúra repa

```
VUZ_Sorter/
├─ lta_triedenie.py          celý nástroj (MasterData → inventár → text → párovanie → odkazy → Excel)
├─ config/
│  └─ demo.ini               konfigurácia pre demo dáta (vzor pre reálne *.local.ini)
├─ demo_data/                FIKTÍVNE testovacie dáta (MasterData_DEMO.xlsx + neroztriedene/)
│  └─ vytvor_demo_data.py    vygeneruje demo dáta nanovo
├─ testy/
│  ├─ test_parovanie.py      pravidlá párovania 1 – 7
│  └─ test_demo_e2e.py       celý beh na demo dátach (test, ostro, opakovanie)
├─ docs/
│  ├─ PRAVIDLA.md            pravidlá párovania, výstupy, identifikácia dokumentov
│  └─ ROADMAP.md             čo ďalej (OCR, …)
├─ .github/                  automatické testy pri PR + šablóna PR
├─ TEST.bat / OSTRO.bat      spustenie dvojklikom
├─ CONTRIBUTING.md           ako spolu pracujeme (vetvy, PR, review)
└─ CHANGELOG.md
```

## Čo nástroj robí

1. **MasterData** – načíta Tag No z hárkov v `.ini`, iba `Part of scope = Y`.
2. **Inventár** – každý súbor dostane SHA-256 odtlačok obsahu. Rovnaký obsah = jeden dokument.
3. **Text** – PDF, DOCX, XLSX/XLSM, XLS, TXT/CSV, RTF, DOC (cez Word); uloží sa do cache.
   Skeny a obrázky dostanú stav *potrebuje OCR*.
4. **Párovanie** – pravidlá 1 – 7 ([docs/PRAVIDLA.md](docs/PRAVIDLA.md)), v obsahu, názve súboru aj priečinka.
5. **Výstup** – `vystup/_dokumenty` (každý dokument raz), `vystup/zariadenia/<hárok>/<tag>/` (odkazy .lnk),
   `index.xlsx` (index tag ↔ dokument), `kontrola.xlsx` (nejasné zhody na rozhodnutie A/N).

Opakovaný beh nič neduplikuje. Predvolený režim je TEST; ostrý beh len s `--ostro`.

## Dáta – dôležité

Do repa patrí **iba kód a fiktívne demo dáta**. `.gitignore` blokuje dokumenty, Excel, výstupy
a `*.local.ini` mimo `demo_data/`. Pred commitom vždy skontroluj `git status`.
