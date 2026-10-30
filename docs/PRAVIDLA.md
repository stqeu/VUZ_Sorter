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

- **OCR** zatiaľ nie je – súbory sú v protokole označené „potrebuje OCR“ / „čiastočne OCR“ so zoznamom strán.
- `.doc` potrebuje nainštalovaný MS Word a `pip install pywin32`.
- Odkazy .lnk obsahujú absolútnu cestu – po presune priečinka `vystup` stačí spustiť skript znova.
- Veľmi dlhé cesty (> 260 znakov, OneDrive) môže Windows odmietnuť – výstup dávaj čo najplytšie.
- `11V109 A potom…` (veľké „A“ ako začiatok vety za tagom) sa vyhodnotí ako 11V109A – zriedkavé, ukáže sa v protokole ako *tolerovaná*.
