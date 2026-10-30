# Zmeny

## [0.2.0] – 2026-10-30 – OCR

- Lokálne OCR pre skeny a obrázky: **OCRmyPDF** (PDF, vyrobí aj prehľadávateľné PDF) a **Tesseract**
  (obrázky; záloha pre PDF bez OCRmyPDF – strany sa vykreslia cez PyMuPDF / pypdfium2).
- OCR iba na stranách bez textu (`--skip-text`), obrázky sa neprekomprimujú (`--optimize 0`), otočené strany sa narovnajú.
- Výsledky OCR v cache podľa SHA-256 – každý dokument sa OCR-uje iba raz.
- Odkaz pri tagu vedie na prehľadávateľnú verziu skenu (`…__hash8.ocr.pdf`), originál ostáva vedľa.
- Oprava zámen OCR (O/0, C/0, D/0, I/1, l/1, S/5, B/8, Z/2, G/6) – zhoda typu *OCR oprava* ide vždy na kontrolu;
  ak je na danom mieste správne prečítaný tag, alternatívy sa nenavrhujú.
- Zhody z OCR textu sú v indexe označené *z OCR*; hárok *Dokumenty* má stĺpce OCR strany, OCR motor, Prehľadávateľné PDF.
- Nová sekcia `[ocr]` v konfigurácii. Bez Tesseractu skript beží ďalej (stav *potrebuje OCR*).
- Testy: oprava zámen, zlúčenie a cache OCR (bez Tesseractu), skutočné OCR (iba ak je Tesseract).
- Realistické demo skeny v `demo_data/neroztriedene/skeny/` (2-stranový sken s rozsahom, zmiešané PDF,
  otočený sken, fotka štítku, zámerná chyba `11VI09B`). Tesseract bez OCRmyPDF používa `--psm 1` (otočené strany).
- Zámena `£`/`€` → `E` (typická chyba OCR, napr. `11£101`).
- OCRmyPDF: `--rotate-pages-threshold 2` a `--tesseract-pagesegmode 1` – otočené skeny sa narovnajú aj pri nižšej istote.
- Čítanie PDF: ak PyMuPDF vráti pri strane prázdny text (napr. otočená strana s OCR vrstvou), strana sa prečíta cez pypdf.
- Nový typ odkazov `markdown`: `README.md` v priečinku zariadenia + `PREHLAD.md` – čitateľné na GitHube.
  Súbory sa prepisujú iba pri zmene.
- `config/ukazka.ini` + `UKAZKA.bat` → `ukazka_vystupu/` v repe; skript odmietne do neho zapisovať iné než demo dáta.

## [0.1.0] – 2026-10-30 – demo

- Jeden skript `lta_triedenie.py` + konfigurácia `.ini` namiesto pôvodných jednorazových skriptov.
- MasterData: všetky hárky z konfigurácie, iba `Part of scope = Y`, duplicitné tagy = jeden kľúč.
- Dokumenty identifikované SHA-256 odtlačkom obsahu; rovnaký obsah = jeden dokument.
- Čítanie textu: PDF (PyMuPDF / pypdf), DOCX, XLSX/XLSM, XLS, TXT/CSV, RTF, DOC (Word). Cache v SQLite.
- Párovanie podľa pravidiel 1 – 7 (oddeľovače, hranice, rozsahy, bez predpony, súčasť dlhšieho tagu,
  krátke tagy, potrubia + DN/palce), v obsahu, názve súboru aj priečinka.
- Úložisko `_dokumenty` (každý dokument raz) + odkazy `.lnk` / `.url` / symlink v priečinkoch zariadení.
- Excel protokol (index tag ↔ dokument) a `kontrola.xlsx` s rozhodnutiami A/N zachovanými medzi behmi.
- Stav „potrebuje OCR“ pre skeny a obrázky (OCR zatiaľ nie je).
- Testy: 9 testov pravidiel + end-to-end test demo behu.
