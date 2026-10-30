# VUZ Sorter – triedenie dokumentov podľa Tag No

Nástroj priradí každý dokument ku **všetkým Tag No, ktoré obsahuje**. Dokument uloží **raz**
a v priečinkoch zariadení vytvorí iba **odkazy**. Nahrádza pôvodných ~14 jednorazových skriptov.

> **Stav:** v0.2 – demo + lokálne OCR (Tesseract, OCRmyPDF). Ďalej pozri [docs/ROADMAP.md](docs/ROADMAP.md).

### 👉 [Ukážka výsledku triedenia – ukazka_vystupu/PREHLAD.md](ukazka_vystupu/PREHLAD.md)

Fiktívne demo dokumenty roztriedené podľa Tag No – dá sa preklikať zo zariadenia na dokument priamo na GitHube.
Vygeneruje sa cez `UKAZKA.bat`.

## Rýchly štart

```powershell
pip install -r requirements.txt
python -m unittest discover testy      # musí byť OK (test skutočného OCR sa bez Tesseractu preskočí)
.\TEST.bat                             # skúšobný beh na demo dátach – nič nevytvorí, iba protokol
.\OSTRO.bat                            # ostrý beh – vystup\_dokumenty + vystup\zariadenia
.\UKAZKA.bat                           # ukážka pre GitHub – ukazka_vystupu\ (README.md + PREHLAD.md)
```

**OCR (voliteľné, odporúčané):** nainštaluj [Tesseract](https://github.com/UB-Mannheim/tesseract/wiki)
s jazykmi Slovak + Czech, Ghostscript (64-bit, AGPL) a `python -m pip install ocrmypdf`.
Kontrola: `& "C:\Program Files\Tesseract-OCR\tesseract.exe" --list-langs` (musí obsahovať `slk`, `ces`, `eng`).

Reálny projekt: skopíruj `config/demo.ini` → `config/rhc.local.ini`, uprav cesty a spusti
`.\TEST.bat --config config\rhc.local.ini`. Súbory `*.local.ini` sa do gitu nenahrávajú.

## Štruktúra repa

```
VUZ_Sorter/
├─ lta_triedenie.py          celý nástroj (MasterData → inventár → text → párovanie → odkazy → Excel)
├─ config/
│  ├─ demo.ini               konfigurácia pre demo dáta (vzor pre reálne *.local.ini)
│  └─ ukazka.ini             ukážka pre GitHub (výstup ukazka_vystupu/, odkazy ako README.md)
├─ demo_data/                FIKTÍVNE testovacie dáta (MasterData_DEMO.xlsx + neroztriedene/, skeny/ na OCR)
│  └─ vytvor_demo_data.py    vygeneruje demo dáta nanovo
├─ testy/
│  ├─ test_parovanie.py      pravidlá párovania 1 – 7
│  ├─ test_ocr.py            OCR: oprava zámen, cache, skutočné OCR (ak je Tesseract)
│  └─ test_demo_e2e.py       celý beh na demo dátach (test, ostro, opakovanie)
├─ ukazka_vystupu/           vygenerovaný výsledok z demo dát (iba fiktívne dáta!)
├─ docs/
│  ├─ PRAVIDLA.md            pravidlá párovania, výstupy, identifikácia dokumentov
│  └─ ROADMAP.md             čo ďalej (OCR, …)
├─ .github/                  automatické testy pri PR + šablóna PR
├─ TEST.bat / OSTRO.bat / UKAZKA.bat   spustenie dvojklikom
├─ CONTRIBUTING.md           ako spolu pracujeme (vetvy, PR, review)
└─ CHANGELOG.md
```

## Čo nástroj robí

1. **MasterData** – načíta Tag No z hárkov v `.ini`, iba `Part of scope = Y`.
2. **Inventár** – každý súbor dostane SHA-256 odtlačok obsahu. Rovnaký obsah = jeden dokument.
3. **Text** – PDF, DOCX, XLSX/XLSM, XLS, TXT/CSV, RTF, DOC (cez Word); uloží sa do cache.
   Skeny a obrázky prejdú **lokálnym OCR** (iba strany bez textu, výsledok v cache).
   Bez Tesseractu dostanú stav *potrebuje OCR*.
4. **Párovanie** – pravidlá 1 – 7 ([docs/PRAVIDLA.md](docs/PRAVIDLA.md)), v obsahu, názve súboru aj priečinka.
5. **Výstup** – `vystup/_dokumenty` (každý dokument raz), `vystup/zariadenia/<hárok>/<tag>/` (odkazy .lnk,
   alebo `README.md` s odkazmi pri `typ_odkazu = markdown`), `PREHLAD.md`, `index.xlsx` (index tag ↔ dokument),
   `kontrola.xlsx` (nejasné zhody na rozhodnutie A/N).

Opakovaný beh nič neduplikuje. Predvolený režim je TEST; ostrý beh len s `--ostro`.

## Dáta – dôležité

Do repa patrí **iba kód a fiktívne demo dáta**. `.gitignore` blokuje dokumenty, Excel, výstupy
a `*.local.ini` mimo `demo_data/`. Pred commitom vždy skontroluj `git status`.
