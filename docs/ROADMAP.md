# Čo ďalej

## 1. OCR (ďalší krok)
- Strany označené *potrebuje OCR* / *čiastočne OCR* spracovať lokálne – nič neodchádza mimo počítača.
- Tlačené skeny: **OCRmyPDF + Tesseract** (`slk+ces+eng`).
- Rukopisy: **Ollama + Qwen2.5-VL 7B / Qwen3-VL 8B** na RTX 3070 – pýtať sa iba na Tag No,
  výsledok overiť voči MasterData, nízka istota → `kontrola.xlsx`.
- OCR text uložiť do cache rovnako ako bežný text (podľa SHA-256), aby sa OCR robilo iba raz.

## 2. Rovnaké názvy, rôzny obsah
- Pri kolízii názvov v jednom priečinku tagu dať čitateľný názov obom odkazom
  (napr. `Protokol (revizie).pdf.lnk`), aby sa mená odkazov nemenili podľa poradia súborov.
- V protokole hárok *Rovnaký názov – rôzny obsah* (verzie dokumentov).
- Prehľad súborov v `_dokumenty`, na ktoré už nič neodkazuje (staré verzie).

## 3. Reálne nasadenie
- `config/rhc.local.ini`, `config/hrp7.local.ini` – reálne cesty (mimo gitu).
- Pilot na vzorke reálnych dokumentov, kontrola falošných zhôd v `kontrola.xlsx`.
- Výkon pri ~6 000 dokumentoch a ~8 700 tagoch.
