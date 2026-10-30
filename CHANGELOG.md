# Zmeny

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
