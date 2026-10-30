# Pravidlá párovania a výstupy

## Pravidlá párovania

| # | Pravidlo | Príklad | Výsledok |
|---|---|---|---|
| 1 | Oddeľovače sa tolerujú | `11 V-109 A` | 11V109A (*tolerovaná*) |
| 1 | …ale nie malé písmeno za medzerou | `11A101 a potrubie` | iba 11A101 |
| 2 | Hranice tagu | `11V109A`, `11V1090` | iba 11V109A, druhé nič |
| 3 | Rozsahy | `11V109A-C`, `11P110A/B` | A, B, C / A, B (*rozsah*) |
| 4 | Bez predpony jednotky | `V109A` | 11V109A → **na kontrolu** |
| 5 | Tag vnútri dlhšieho tagu | `57DRM2-11E126A` | oba (*súčasť dlhšieho tagu*) |
| 6 | Krátke tagy (`E1`, `P-11`, bez číslic, `PLUG-VALVE-*`) | `E1_kalibracia.txt` / `…E2…` v texte | z názvu automaticky / z obsahu **na kontrolu** |
| 7 | Potrubia: časť pred zátvorkou, DN ↔ palce | `11-DEA-N-1105-DN 50-B1E`, `TBP-0279-DN80-A1B` | 11-DEA-N-1105-DN50-B1E, TBP-0279-3"-A1B |

Typy súborov v `typy_len_nazov` (napr. `.dwg`) sa nečítajú, priradia sa iba podľa názvu.

## Štruktúra výstupu

```
vystup/
  _dokumenty/                      každý dokument raz:  <názov>__<hash8>.<prípona>
  zariadenia/
    Static/11V109A/                Revizna_sprava_2024_odlucovace.pdf.lnk, Supis_zariadeni_jednotka_11.xlsx.lnk
    Pipes/11-DEA-N-1105-DN50-B1E/  ...
  index.xlsx                       protokol (pri teste protokol_TEST.xlsx)
  kontrola.xlsx                    nejasné zhody na rozhodnutie A/N
  posledny_beh.log
  .cache/                          SQLite cache (hash + text), zoznam vytvorených odkazov
```

## OCR

- Na OCR idú iba dokumenty so stavom *potrebuje OCR* / *čiastočne OCR* a iba ich strany bez textu.
- PDF: **OCRmyPDF** → prehľadávateľné PDF (`_dokumenty/<názov>__<hash8>.ocr.pdf`), odkazy pri tagoch vedú naň;
  originál ostáva vedľa. Bez OCRmyPDF: strany sa vykreslia a prečíta ich priamo Tesseract (iba text).
- Obrázky (`.jpg`, `.png`, `.tif`): Tesseract.
- Zhody z OCR textu majú v indexe typ doplnený o *z OCR* a platia pre ne rovnaké pravidlá 1 – 7.
- **Oprava zámen:** ak tag v OCR texte nesedí presne, skúsi sa porovnanie po zjednotení zameniteľných znakov
  (O/0, C/0, D/0, Q/0, I/1, l/1, S/5, B/8, Z/2, G/6, £/E, €/E). Takáto zhoda má typ *OCR oprava* a ide **vždy na kontrolu**.
  Ak sa text dá zjednotiť na viac tagov (napr. `110101` → 11C101 aj 11D101), na kontrolu idú všetci kandidáti.
- Výsledky OCR sú v cache podľa SHA-256; zmena jazykov v `[ocr]` spôsobí nové OCR.

## Markdown výstup (GitHub)

Pri `typ_odkazu = markdown` sa namiesto skratiek vytvorí v každom priečinku zariadenia `README.md`
(tabuľka dokumentov s odkazmi do `_dokumenty`, strany, typ zhody, nájdený text, sekcia *Čaká na kontrolu*)
a vo výstupe `PREHLAD.md` (súhrn, zariadenia, dokumenty, tagy bez dokumentov). `PREHLAD.md` vzniká pri
každom ostrom behu. Ukážka: [ukazka_vystupu/PREHLAD.md](../ukazka_vystupu/PREHLAD.md).

## Identifikácia dokumentov

Každý dokument je identifikovaný **SHA-256 odtlačkom obsahu** – nie názvom.

- rovnaký obsah pod rôznymi názvami / v rôznych priečinkoch = jeden dokument
- rôzny obsah s rovnakým názvom = dva dokumenty (`Protokol__a1b2c3d4.pdf`, `Protokol__9f8e7d6c.pdf`)
- akákoľvek zmena obsahu (aj opätovné uloženie) = nový dokument
- odtlačok je v `index.xlsx` (hárok *Dokumenty*), prvých 8 znakov v názve súboru v `_dokumenty`
- overenie kedykoľvek aj bez nástroja: `Get-FileHash "subor.pdf" -Algorithm SHA256`

## Nastavenia (`config/*.ini`)

- `format_priecinka` – `{tag}` → `11E101`, alebo starý formát `BRHCK_LTA_{tag_bez_jednotky}_r0` → `BRHCK_LTA_E101_r0`
- `ulozisko` – `kopia` (jedna kópia vo výstupe) alebo `zdroj` (odkazy rovno na pôvodné súbory, nič sa nekopíruje)
- `typ_odkazu` – `auto` (.lnk na Windows) | `url` | `symlink`. Ak firma blokuje PowerShell, skript sám prejde na `.url`.
- `[harok:…]` – ktoré hárky MasterData a ktorý stĺpec je tag; `potrubia = ano` zapne pravidlo 7

## Známe obmedzenia demo verzie

- **Rukopis** Tesseract prečíta zle – na rukopisy je v pláne lokálny model (Qwen-VL cez Ollama).
- `.doc` potrebuje nainštalovaný MS Word a `pip install pywin32`.
- Odkazy .lnk obsahujú absolútnu cestu – po presune priečinka `vystup` stačí spustiť skript znova.
- Veľmi dlhé cesty (> 260 znakov, OneDrive) môže Windows odmietnuť – výstup dávaj čo najplytšie.
- `11V109 A potom…` (veľké „A“ ako začiatok vety za tagom) sa vyhodnotí ako 11V109A – zriedkavé, ukáže sa v protokole ako *tolerovaná*.
