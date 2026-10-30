# -*- coding: utf-8 -*-
"""
LTA triedenie dokumentov podľa Tag No
=====================================

Jeden skript namiesto pôvodných ~14. Robí celý postup:

  1. načíta zoznam tagov z MasterData (Excel), iba položky v rozsahu (Part of scope = Y)
  2. prejde zdrojové priečinky, každému súboru vypočíta SHA-256 (rovnaký obsah = jeden dokument)
  3. vytiahne text (PDF, Word, Excel, TXT/CSV, RTF) a uloží ho do cache
  4. nájde v texte, názve súboru a ceste Tag No podľa pravidiel 1 – 7
  5. uloží každý dokument RAZ (vystup/_dokumenty) a pri tagoch vytvorí iba ODKAZY
  6. zapíše Excel protokol (index tag <-> dokument) a hárok na ručnú kontrolu

Použitie:
    python lta_triedenie.py                         -> TEST, nič sa nekopíruje (iba protokol)
    python lta_triedenie.py --ostro                 -> ostrý beh
    python lta_triedenie.py --config config/rhc.local.ini   -> reálny projekt

OCR (lokálne): skeny a obrázky sa prečítajú cez Tesseract / OCRmyPDF, ak sú nainštalované.
Bez nich skript beží ďalej a súbory označí ako „potrebuje OCR“.
"""

from __future__ import annotations

import argparse
import configparser
import fnmatch
import hashlib
import json
import logging
import os
import re
import shutil
import sqlite3
import subprocess
import sys
import tempfile
import unicodedata
from collections import defaultdict
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

VERZIA = "0.2.0-ocr"
VERZIA_EXTRAKCIE = 2  # zvýš, ak sa zmení spôsob čítania textu -> cache sa prepočíta

log = logging.getLogger("lta")


# =============================================================================
# 1. NASTAVENIA
# =============================================================================

@dataclass
class HarokNastavenie:
    nazov: str                   # názov hárku v MasterData
    stlpec_tag: str              # hlavička stĺpca s tagom
    stlpec_rozsah: str | None    # hlavička stĺpca "Part of scope" (None = berie všetko)
    priecinok: str               # podpriečinok vo výstupe
    potrubia: bool = False       # pravidlo 7


@dataclass
class Nastavenia:
    projekt: str
    masterdata: Path
    zdroje: list[Path]
    vystup: Path
    ulozisko: str                     # "kopia" alebo "zdroj"
    typ_odkazu: str                   # auto | lnk | url | symlink
    format_priecinka: str
    vytvorit_prazdne_priecinky: bool
    typy_suborov: set[str]
    typy_len_nazov: set[str]          # obsah sa nečíta, páruje sa iba názov (výkresy, archívy)
    min_znakov_na_stranu: int
    len_v_rozsahu: bool
    predpona_jednotky: str
    max_dlzka_kratkeho_tagu: int
    nejednoznacne_navyse: list[str]
    povolit_bez_predpony: bool
    harky: list[HarokNastavenie] = field(default_factory=list)
    ocr: "OcrNastavenie" = None  # type: ignore[assignment]


def _ano(hodnota: str) -> bool:
    return str(hodnota).strip().lower() in {"ano", "áno", "yes", "true", "1", "y"}


def nacitaj_nastavenia(cesta: Path) -> Nastavenia:
    if not cesta.is_file():
        raise SystemExit(f"Konfiguračný súbor neexistuje: {cesta}")
    cp = configparser.ConfigParser(inline_comment_prefixes=(";", "#"))
    cp.read(cesta, encoding="utf-8")
    zaklad = cesta.parent  # relatívne cesty sú voči umiestneniu .ini

    def cesta_z(text: str) -> Path:
        p = Path(text.strip())
        return p if p.is_absolute() else (zaklad / p).resolve()

    p = cp["projekt"]
    par = cp["parovanie"]
    n = Nastavenia(
        projekt=p.get("nazov", "LTA"),
        masterdata=cesta_z(p["masterdata"]),
        zdroje=[cesta_z(r) for r in p["zdroje"].splitlines() if r.strip()],
        vystup=cesta_z(p["vystup"]),
        ulozisko=p.get("ulozisko", "kopia").strip().lower(),
        typ_odkazu=p.get("typ_odkazu", "auto").strip().lower(),
        format_priecinka=p.get("format_priecinka", "{tag}").strip(),
        vytvorit_prazdne_priecinky=_ano(p.get("vytvorit_prazdne_priecinky", "ano")),
        typy_suborov={t.lower() for t in p.get("typy_suborov", ".pdf .docx .txt").split()},
        typy_len_nazov={t.lower() for t in p.get("typy_len_nazov", "").split()},
        min_znakov_na_stranu=p.getint("min_znakov_na_stranu", 50),
        len_v_rozsahu=_ano(p.get("len_v_rozsahu", "ano")),
        predpona_jednotky=par.get("predpona_jednotky", "11").strip(),
        max_dlzka_kratkeho_tagu=par.getint("max_dlzka_kratkeho_tagu", 5),
        nejednoznacne_navyse=par.get("nejednoznacne_navyse", "").split(),
        povolit_bez_predpony=_ano(par.get("povolit_bez_predpony", "ano")),
    )
    for sekcia in cp.sections():
        if not sekcia.lower().startswith("harok:"):
            continue
        s = cp[sekcia]
        nazov = sekcia.split(":", 1)[1].strip()
        n.harky.append(HarokNastavenie(
            nazov=nazov,
            stlpec_tag=s.get("stlpec_tag", "Tag No"),
            stlpec_rozsah=s.get("stlpec_rozsah", "Part of scope") or None,
            priecinok=s.get("priecinok", nazov),
            potrubia=_ano(s.get("potrubia", "nie")),
        ))
    o = cp["ocr"] if cp.has_section("ocr") else {}
    n.ocr = OcrNastavenie(
        zapnute=_ano(o.get("zapnute", "ano")),
        jazyky=o.get("jazyky", "slk+ces+eng").strip(),
        tesseract=o.get("tesseract", "").strip(),
        pouzit_ocrmypdf=_ano(o.get("ocrmypdf", "ano")),
        prehladavatelne_pdf=_ano(o.get("prehladavatelne_pdf", "ano")),
        oprava_zamien=_ano(o.get("oprava_zamien", "ano")),
        max_stran=int(o.get("max_stran", 200)),
        casovy_limit_s=int(o.get("casovy_limit_s", 900)),
        paralelne=int(o.get("paralelne", 2)),
    )
    if n.ulozisko not in {"kopia", "zdroj"}:
        raise SystemExit("ulozisko musí byť 'kopia' alebo 'zdroj'")
    return n


# =============================================================================
# 2. MASTERDATA – zoznam tagov
# =============================================================================

def kanonicky(text: str) -> str:
    """Kľúč tagu: veľké písmená, bez oddeľovačov. '11 V-109 a' -> '11V109A'."""
    return re.sub(r"[\W_]", "", normalizuj_text(text).upper())


def tag_potrubia(nazov: str) -> str:
    """Pravidlo 7: SAP skracuje text na 40 znakov -> berieme časť pred zátvorkou."""
    return re.sub(r"\s+", " ", nazov.split(" (")[0].split("(")[0]).strip()


@dataclass
class Tag:
    kluc: str                          # kanonický kľúč (unikátny)
    zobrazenie: str                    # ako sa tag zobrazuje (prvý výskyt v MasterData)
    harky: list[str] = field(default_factory=list)
    povodne_nazvy: set[str] = field(default_factory=set)
    potrubie: bool = False
    nejednoznacny: bool = False        # pravidlo 6


def _norm_hlavicka(h) -> str:
    return re.sub(r"\s+", " ", str(h or "")).strip().lower()


def nacitaj_masterdata(n: Nastavenia) -> dict[str, Tag]:
    from openpyxl import load_workbook

    if not n.masterdata.is_file():
        raise SystemExit(f"MasterData neexistuje: {n.masterdata}")
    wb = load_workbook(n.masterdata, read_only=True, data_only=True)
    tagy: dict[str, Tag] = {}
    try:
        for h in n.harky:
            if h.nazov not in wb.sheetnames:
                log.warning("Hárok '%s' v MasterData nie je, preskakujem.", h.nazov)
                continue
            riadky = list(wb[h.nazov].iter_rows(values_only=True))
            # hlavičku hľadáme v prvých 15 riadkoch (MasterData má dvojriadkovú hlavičku)
            i_hl, i_tag, i_roz = None, None, None
            for i, r in enumerate(riadky[:15]):
                hl = [_norm_hlavicka(c) for c in r]
                if _norm_hlavicka(h.stlpec_tag) in hl:
                    i_hl, i_tag = i, hl.index(_norm_hlavicka(h.stlpec_tag))
                    if h.stlpec_rozsah and _norm_hlavicka(h.stlpec_rozsah) in hl:
                        i_roz = hl.index(_norm_hlavicka(h.stlpec_rozsah))
                    break
            if i_hl is None:
                log.warning("Hárok '%s': nenašiel som stĺpec '%s'.", h.nazov, h.stlpec_tag)
                continue
            pocet = 0
            for r in riadky[i_hl + 1:]:
                if i_tag >= len(r) or r[i_tag] in (None, ""):
                    continue
                if n.len_v_rozsahu and i_roz is not None and str(r[i_roz]).strip().upper() != "Y":
                    continue
                povodny = str(r[i_tag]).strip()
                zobraz = tag_potrubia(povodny) if h.potrubia else povodny
                k = kanonicky(zobraz)
                if not k:
                    continue
                t = tagy.get(k)
                if t is None:
                    t = tagy[k] = Tag(kluc=k, zobrazenie=zobraz, potrubie=h.potrubia)
                if h.nazov not in t.harky:
                    t.harky.append(h.nazov)
                t.povodne_nazvy.add(povodny)
                pocet += 1
            log.info("MasterData %-16s %5d riadkov v rozsahu", h.nazov, pocet)
    finally:
        wb.close()

    # pravidlo 6 – krátke / nejednoznačné tagy
    for t in tagy.values():
        bez_cislic = not re.search(r"\d", t.kluc)
        kratky = len(t.kluc) <= n.max_dlzka_kratkeho_tagu
        rucne = any(fnmatch.fnmatch(t.zobrazenie.upper(), vzor.upper()) for vzor in n.nejednoznacne_navyse)
        t.nejednoznacny = (kratky or bez_cislic or rucne)
    log.info("Unikátnych tagov: %d (z toho nejednoznačných: %d)",
             len(tagy), sum(t.nejednoznacny for t in tagy.values()))
    return tagy


# =============================================================================
# 3. PÁROVANIE – pravidlá 1 až 7
# =============================================================================

POMLCKY = dict.fromkeys(map(ord, "\u2010\u2011\u2012\u2013\u2014\u2212"), "-")
UVODZOVKY = dict.fromkeys(map(ord, "\u201c\u201d\u201e\u2033\u02ba"), '"')


def normalizuj_text(text: str) -> str:
    """Zjednotí pomlčky, úvodzovky a medzery. Veľkosť písmen ZACHOVÁ (pravidlo 1 – ochrana pred spojkou 'a')."""
    text = unicodedata.normalize("NFC", text or "")
    return text.translate(POMLCKY).translate(UVODZOVKY).replace("\u00a0", " ")


# DN <-> palce (pravidlo 7)
DN_PALCE = {
    15: '1/2', 20: '3/4', 25: '1', 32: '1 1/4', 40: '1 1/2', 50: '2', 65: '2 1/2', 80: '3',
    100: '4', 125: '5', 150: '6', 200: '8', 250: '10', 300: '12', 350: '14', 400: '16',
    450: '18', 500: '20', 600: '24', 700: '28', 800: '32', 900: '36',
}
PALCE_DN = {v: k for k, v in DN_PALCE.items()}

ZAC = r"(?<![^\W_])"     # pred tagom nesmie byť písmeno/číslica (podčiarkovník je OK)
KON = r"(?![^\W_])"      # za tagom nesmie byť písmeno/číslica
SEP_VOL = r"(?:[ \t]?-?[ \t]?)"                 # pravidlo 1: medzera / pomlčka medzi skupinami
SEP_VLASTNY = r"(?:[ \t]?[-_.,/]?[ \t]?)"       # tag sám obsahuje oddeľovač medzi rôznymi skupinami
SEP_POVINNY = r"(?:[ \t]?[-_.,/][ \t]?|[ \t])"  # 0,6  3.5  6-4_1  -> oddeľovač musí ostať


def _palce_vzor(palce: str) -> str:
    return re.escape(palce).replace(r"\ ", r"[ \t-]?") + r"[ \t]?(?:\"|''|IN\b|ZOLL)"


def _atomy(tag: str) -> list[str]:
    return re.findall(r"\d+|[^\W\d_]+|[^\w\s]|_+|\s+", tag)


def vzor_tagu(tag: str, potrubie: bool, skupiny: bool = True) -> tuple[str, list[str]]:
    """Postaví regex pre jeden tag podľa pravidiel 1, 2 a 7.

    Vráti (regex, mená skupín). Skupiny označujú písmenové časti tagu, ktoré sú v MasterData
    veľkými písmenami – ak ich v texte oddeľuje medzera a sú malými písmenami, nejde o tag
    ('11A101 a potrubie' nie je 11A101A).
    """
    atomy = _atomy(tag)
    casti: list[str] = []
    mena: list[str] = []
    predch_alnum: str | None = None
    medzera_vlastna = False
    i = 0
    while i < len(atomy):
        a = atomy[i]
        if not a.isalnum() or a.startswith("_"):
            if a.strip() == '"':          # palcová značka patrí k číslu (TBP-0279-3"-A1B)
                i += 1
                continue
            medzera_vlastna = True
            i += 1
            continue
        # spojenie s predchádzajúcou skupinou
        if predch_alnum is not None:
            rovnaka_trieda = predch_alnum.isdigit() == a.isdigit()
            if medzera_vlastna:
                casti.append(SEP_POVINNY if rovnaka_trieda else SEP_VLASTNY)
            else:
                casti.append(SEP_VOL)
        # pravidlo 7: DN50 <-> 2"
        nasl = atomy[i + 1] if i + 1 < len(atomy) else ""
        if potrubie and a.upper() == "DN" and nasl.isdigit():
            dn = int(nasl)
            alt = [r"DN[ \t]?-?[ \t]?" + nasl]
            if dn in DN_PALCE:
                alt.append(_palce_vzor(DN_PALCE[dn]))
            casti.append("(?:" + "|".join(alt) + ")")
            predch_alnum, medzera_vlastna = nasl, False
            i += 2
            continue
        if potrubie and a.isdigit() and nasl.strip() == '"':
            alt = [_palce_vzor(a)]
            if a in PALCE_DN:
                alt.append(r"DN[ \t]?-?[ \t]?" + str(PALCE_DN[a]))
            casti.append("(?:" + "|".join(alt) + ")")
            predch_alnum, medzera_vlastna = a, False
            i += 2
            continue
        if skupiny and not a.isdigit() and a.isupper():
            meno = f"L{len(mena)}"
            mena.append(meno)
            casti.append(f"(?P<{meno}>{re.escape(a)})")
        else:
            casti.append(re.escape(a))
        predch_alnum, medzera_vlastna = a, False
        i += 1
    return ZAC + "".join(casti) + KON, mena


def _jadro(tag: str) -> str:
    """Regex tagu bez hraníc a skupín – na vloženie do iného regexu (rozsahy)."""
    return vzor_tagu(tag, False, skupiny=False)[0][len(ZAC):-len(KON)]


def _kotva(tag: str) -> str | None:
    """Najdlhšia číselná skupina – rýchly predfilter (netreba skúšať 8 700 regexov na každý text)."""
    cisla = re.findall(r"\d+", tag)
    return max(cisla, key=len) if cisla else None


# Typické zámeny OCR: znaky v jednej skupine sa pri OCR zhode považujú za rovnaké.
# Text aj tag sa prevedú na „kostru“ rovnakej dĺžky (pozície ostávajú) a porovnajú sa kostry.
ZAMENY_OCR = str.maketrans({"O": "0", "o": "0", "D": "0", "Q": "0", "C": "0",
                            "I": "1", "l": "1", "|": "1", "!": "1",
                            "S": "5", "s": "5", "B": "8", "Z": "2", "z": "2", "G": "6",
                            "£": "E", "€": "E"})


def parovac_zamen(tagy: dict[str, "Tag"], n: "Nastavenia") -> "Parovac":
    """Párovač nad kostrami tagov – na nájdenie '11V1O9A' (OCR) ako 11V109A. Zhody idú na kontrolu."""
    kostry = {k: Tag(k, t.zobrazenie.translate(ZAMENY_OCR), list(t.harky), set(), t.potrubie, t.nejednoznacny)
              for k, t in tagy.items()}
    return Parovac(kostry, n)


@dataclass
class Zhoda:
    kluc: str
    zaciatok: int
    koniec: int
    typ: str          # presná | tolerovaná | rozsah | súčasť dlhšieho tagu | bez predpony | názov súboru | priečinok
    text: str


class Parovac:
    """Hľadá Tag No v texte. Vytvorí sa raz, používa sa na všetky dokumenty."""

    def __init__(self, tagy: dict[str, Tag], n: Nastavenia):
        self.tagy = tagy
        self.vzory: dict[str, tuple[re.Pattern, list[str]]] = {}
        self.podla_kotvy: dict[str | None, list[str]] = defaultdict(list)
        for t in tagy.values():
            vz, mena = vzor_tagu(t.zobrazenie, t.potrubie)
            self.vzory[t.kluc] = (re.compile(vz, re.IGNORECASE), mena)
            self.podla_kotvy[_kotva(t.zobrazenie)].append(t.kluc)

        # pravidlo 3 – rodiny tagov líšiace sa posledným písmenom (11V109A, 11V109B, ...)
        self.rodiny: dict[str, dict[str, str]] = defaultdict(dict)   # základ -> {písmeno: kľúč}
        zaklad_zobraz: dict[str, str] = {}
        for t in tagy.values():
            if t.potrubie:
                continue
            m = re.fullmatch(r"(.*\d)[ \t-]?([A-Z])", t.zobrazenie)
            if m:
                zk = kanonicky(m.group(1))
                self.rodiny[zk][m.group(2)] = t.kluc
                zaklad_zobraz.setdefault(zk, m.group(1))
        self.rozsahy: dict[str, tuple[re.Pattern, re.Pattern]] = {}
        self.rozsahy_podla_kotvy: dict[str | None, list[str]] = defaultdict(list)
        for zk, clenovia in self.rodiny.items():
            if len(clenovia) < 2:
                continue
            zv = _jadro(zaklad_zobraz[zk])
            retaz = ZAC + zv + SEP_VOL + r"([A-Z](?:[ \t]?[-/][ \t]?(?:" + zv + SEP_VOL + r")?[A-Z])+)" + KON
            self.rozsahy[zk] = (re.compile(retaz, re.IGNORECASE), re.compile(zv + SEP_VOL, re.IGNORECASE))
            self.rozsahy_podla_kotvy[_kotva(zaklad_zobraz[zk])].append(zk)

        # pravidlo 4 – tag bez predpony jednotky (V109A pre 11V109A) -> iba na kontrolu
        self.aliasy: dict[str, tuple[re.Pattern, list[str]]] = {}
        self.aliasy_podla_kotvy: dict[str | None, list[str]] = defaultdict(list)
        if n.povolit_bez_predpony and n.predpona_jednotky:
            pred = re.escape(n.predpona_jednotky)
            kandidati: dict[str, list[tuple[str, str]]] = defaultdict(list)
            for t in tagy.values():
                if t.potrubie or not re.match(pred + r"[ -]?[A-Za-z]", t.zobrazenie):
                    continue
                alias = re.sub(r"^" + pred + r"[ -]?", "", t.zobrazenie)
                if len(kanonicky(alias)) >= 4 and re.search(r"\d", alias):
                    kandidati[kanonicky(alias)].append((t.kluc, alias))
            for ak, polozky in kandidati.items():
                if len(polozky) == 1 and ak not in tagy:   # iba jednoznačný alias, ktorý sám nie je tagom
                    k, alias = polozky[0]
                    vz, mena = vzor_tagu(alias, False)
                    self.aliasy[k] = (re.compile(vz, re.IGNORECASE), mena)
                    self.aliasy_podla_kotvy[_kotva(alias)].append(k)

    # ------------------------------------------------------------------
    @staticmethod
    def _platna(m: re.Match, mena: list[str]) -> bool:
        """Pravidlo 1 – ochrana: písmenová časť oddelená medzerou nesmie byť malými písmenami
        ('11A101 a potrubie' nie je 11A101A, 'nádrž 11V109 v prevádzke' nie je 11V109V)."""
        text = m.string
        for meno in mena:
            s = m.start(meno)
            if s > m.start() and text[s - 1].isspace() and m.group(meno).islower():
                return False
        return True

    @staticmethod
    def _vyries_prekryvy(zhody: list[Zhoda]) -> list[Zhoda]:
        """Pravidlo 2 + 5.
        - kratšia zhoda s rovnakým začiatkom ako dlhšia sa zahodí ('11V109 A' -> 11V109A, nie 11V109)
        - zhoda vnútri dlhšieho tagu sa ponechá ako 'súčasť dlhšieho tagu' (57DRM2-11E126A -> 11E126A)
        """
        najdlhsia_od: dict[int, int] = {}
        for z in zhody:
            najdlhsia_od[z.zaciatok] = max(najdlhsia_od.get(z.zaciatok, 0), z.koniec)
        vysledok = []
        max_koniec_skor = -1          # najďalší koniec zhôd, ktoré začali SKÔR
        akt_zaciatok, max_koniec_akt = None, -1
        for z in sorted(zhody, key=lambda z: (z.zaciatok, -z.koniec)):
            if z.zaciatok != akt_zaciatok:
                max_koniec_skor = max(max_koniec_skor, max_koniec_akt)
                akt_zaciatok, max_koniec_akt = z.zaciatok, -1
            if najdlhsia_od[z.zaciatok] > z.koniec:
                continue
            if max_koniec_skor >= z.koniec:
                z.typ = "súčasť dlhšieho tagu"
            vysledok.append(z)
            max_koniec_akt = max(max_koniec_akt, z.koniec)
        return vysledok

    def najdi(self, text: str) -> list[Zhoda]:
        text = normalizuj_text(text)
        if not text.strip():
            return []
        kotvy = set(re.findall(r"\d+", text)) | {None}
        zhody: list[Zhoda] = []

        for kotva in kotvy:
            for k in self.podla_kotvy.get(kotva, ()):
                vzor, mena = self.vzory[k]
                povodny = self.tagy[k].zobrazenie
                for m in vzor.finditer(text):
                    if self._platna(m, mena):
                        s = m.group(0)
                        zhody.append(Zhoda(k, m.start(), m.end(), "presná" if s == povodny else "tolerovaná", s))
        vysledok = self._vyries_prekryvy(zhody)

        # pravidlo 3: rozsahy 11V109A-C, 11V109A/B, 11V109A-11V109D
        for kotva in kotvy:
            for zk in self.rozsahy_podla_kotvy.get(kotva, ()):
                vzor, zaklad = self.rozsahy[zk]
                for m in vzor.finditer(text):
                    zvysok = zaklad.sub("", m.group(1))
                    if re.search(r"[a-z]", zvysok):        # 'A-c' – písmená rozsahu musia byť veľké
                        continue
                    pismena: set[str] = set()
                    posledne, op = None, None
                    for c in re.findall(r"[-/]|[A-Z]", zvysok):
                        if c in "-/":
                            op = c
                            continue
                        if op == "-" and posledne:
                            lo, hi = sorted((ord(posledne), ord(c)))
                            pismena.update(chr(x) for x in range(lo, hi + 1))
                        else:
                            pismena.add(c)
                        posledne, op = c, None
                    for p in sorted(pismena):
                        k = self.rodiny[zk].get(p)
                        if k:
                            vysledok.append(Zhoda(k, m.start(), m.end(), "rozsah", m.group(0)))

        # pravidlo 4: bez predpony (iba ak to nie je súčasť plnej zhody)
        for kotva in kotvy:
            for k in self.aliasy_podla_kotvy.get(kotva, ()):
                vzor, mena = self.aliasy[k]
                for m in vzor.finditer(text):
                    if not self._platna(m, mena):
                        continue
                    if any(y.zaciatok <= m.start() and y.koniec >= m.end() for y in vysledok):
                        continue
                    vysledok.append(Zhoda(k, m.start(), m.end(), "bez predpony", m.group(0)))
        return vysledok


# =============================================================================
# 4. ČÍTANIE TEXTU
# =============================================================================

OBRAZKY = {".jpg", ".jpeg", ".png", ".tif", ".tiff", ".bmp"}


class CitacTextu:
    """Vráti zoznam (strana, text). Strana = číslo strany v PDF, názov hárku v Exceli, '-' inak."""

    def __init__(self):
        self._word = None
        self._excel = None

    def zavri(self):
        for app in (self._word, self._excel):
            if app is not None:
                try:
                    app.Quit()
                except Exception:
                    pass

    def citaj(self, cesta: Path) -> list[tuple[str, str]]:
        ext = cesta.suffix.lower()
        if cesta.stat().st_size == 0:
            raise ValueError("prázdny súbor (0 B)")
        if ext in OBRAZKY:
            return [("1", "")]           # obrázok -> text dodá až OCR
        if ext == ".pdf":
            return self._pdf(cesta)
        if ext == ".docx":
            return [("-", self._docx(cesta))]
        if ext in {".xlsx", ".xlsm"}:
            return self._xlsx(cesta)
        if ext == ".xls":
            return self._xls(cesta)
        if ext in {".txt", ".csv"}:
            return [("-", self._txt(cesta))]
        if ext == ".rtf":
            return [("-", self._rtf(cesta))]
        if ext == ".doc":
            return [("-", self._word_com(cesta))]
        raise ValueError(f"nepodporovaný typ {ext}")

    @staticmethod
    def _pdf(cesta: Path) -> list[tuple[str, str]]:
        try:
            import pymupdf  # type: ignore
        except ImportError:
            try:
                import fitz as pymupdf  # type: ignore
            except ImportError:
                pymupdf = None
        from pypdf import PdfReader
        if pymupdf is None:
            r = PdfReader(str(cesta))
            return [(str(i + 1), p.extract_text() or "") for i, p in enumerate(r.pages)]
        with pymupdf.open(cesta) as doc:
            strany = [(str(i + 1), p.get_text("text", sort=True)) for i, p in enumerate(doc)]
        # PyMuPDF pri niektorých stranách (napr. otočená strana s OCR vrstvou) vráti prázdny text –
        # takú stranu skúsime prečítať aj cez pypdf
        prazdne = [i for i, (_, t) in enumerate(strany) if not t.strip()]
        if prazdne:
            try:
                r = PdfReader(str(cesta))
                for i in prazdne:
                    strany[i] = (strany[i][0], r.pages[i].extract_text() or "")
            except Exception:  # noqa: BLE001 – záloha, pôvodný (prázdny) výsledok ostáva
                pass
        return strany

    @staticmethod
    def _docx(cesta: Path) -> str:
        from docx import Document
        d = Document(str(cesta))
        casti = [p.text for p in d.paragraphs]
        for t in d.tables:
            for r in t.rows:
                casti.extend(c.text for c in r.cells)
        for s in d.sections:
            casti.extend(p.text for p in s.header.paragraphs)
            casti.extend(p.text for p in s.footer.paragraphs)
        return "\n".join(casti)

    @staticmethod
    def _xlsx(cesta: Path) -> list[tuple[str, str]]:
        from openpyxl import load_workbook
        wb = load_workbook(str(cesta), read_only=True, data_only=True)
        try:
            return [(ws.title, "\n".join(" ".join(str(v) for v in r if v is not None)
                                         for r in ws.iter_rows(values_only=True)))
                    for ws in wb.worksheets]
        finally:
            wb.close()

    def _xls(self, cesta: Path) -> list[tuple[str, str]]:
        try:
            import xlrd  # type: ignore
            wb = xlrd.open_workbook(str(cesta), on_demand=True)
            try:
                return [(n, "\n".join(" ".join(str(v) for v in wb.sheet_by_name(n).row_values(i))
                                      for i in range(wb.sheet_by_name(n).nrows)))
                        for n in wb.sheet_names()]
            finally:
                wb.release_resources()
        except ImportError:
            pass
        import win32com.client  # type: ignore  # fallback: Excel cez COM
        if self._excel is None:
            self._excel = win32com.client.DispatchEx("Excel.Application")
            self._excel.Visible = False
            self._excel.DisplayAlerts = False
        wb = self._excel.Workbooks.Open(str(cesta), ReadOnly=True)
        try:
            vys = []
            for ws in wb.Worksheets:
                hodn = ws.UsedRange.Value
                riadky = hodn if isinstance(hodn, tuple) else ((hodn,),)
                vys.append((str(ws.Name), "\n".join(" ".join(str(v) for v in (r if isinstance(r, tuple) else (r,))
                                                             if v is not None) for r in riadky)))
            return vys
        finally:
            wb.Close(False)

    @staticmethod
    def _txt(cesta: Path) -> str:
        data = cesta.read_bytes()
        for kod in ("utf-8-sig", "cp1250", "latin-1"):
            try:
                return data.decode(kod)
            except UnicodeDecodeError:
                continue
        return data.decode("utf-8", errors="ignore")

    def _rtf(self, cesta: Path) -> str:
        surove = cesta.read_bytes().decode("latin-1", errors="ignore")
        # jednoduchý prevod RTF -> text (bez Wordu); \'xx = znak v cp1250
        text = re.sub(r"\\'([0-9a-fA-F]{2})", lambda m: bytes([int(m.group(1), 16)]).decode("cp1250", "ignore"), surove)
        text = re.sub(r"\\par[d]?\b", "\n", text)
        text = re.sub(r"\{\\\*[^{}]*\}|\\[a-zA-Z]+-?\d* ?|[{}]", "", text)
        return text

    def _word_com(self, cesta: Path) -> str:
        try:
            import win32com.client  # type: ignore
        except ImportError:
            raise ValueError(".doc vyžaduje MS Word a balík pywin32 (pip install pywin32)")
        if self._word is None:
            self._word = win32com.client.DispatchEx("Word.Application")
            self._word.Visible = False
            self._word.DisplayAlerts = 0
        d = self._word.Documents.Open(str(cesta), ReadOnly=True, AddToRecentFiles=False)
        try:
            return d.Content.Text
        finally:
            d.Close(False)


# =============================================================================
# 5. CACHE (SQLite) – hash súborov a vytiahnutý text
# =============================================================================

class Cache:
    def __init__(self, cesta: Path):
        cesta.parent.mkdir(parents=True, exist_ok=True)
        self.db = sqlite3.connect(cesta)
        self.db.executescript("""
            CREATE TABLE IF NOT EXISTS subory(cesta TEXT PRIMARY KEY, velkost INT, mtime REAL, hash TEXT);
            CREATE TABLE IF NOT EXISTS texty(hash TEXT, verzia INT, poradie INT, strana TEXT, text TEXT);
            CREATE TABLE IF NOT EXISTS extrakcia(hash TEXT PRIMARY KEY, verzia INT, chyba TEXT);
            CREATE TABLE IF NOT EXISTS ocr_texty(hash TEXT, strana TEXT, text TEXT);
            CREATE TABLE IF NOT EXISTS ocr(hash TEXT PRIMARY KEY, verzia INT, jazyky TEXT, motor TEXT,
                                           chyba TEXT, pdf INT);
        """)

    def hash_suboru(self, cesta: Path) -> str:
        st = cesta.stat()
        r = self.db.execute("SELECT velkost, mtime, hash FROM subory WHERE cesta=?", (str(cesta),)).fetchone()
        if r and r[0] == st.st_size and abs(r[1] - st.st_mtime) < 1e-6:
            return r[2]
        h = hashlib.sha256()
        with cesta.open("rb") as f:
            for blok in iter(lambda: f.read(1 << 20), b""):
                h.update(blok)
        hx = h.hexdigest()
        self.db.execute("INSERT OR REPLACE INTO subory VALUES(?,?,?,?)", (str(cesta), st.st_size, st.st_mtime, hx))
        return hx

    def text(self, hx: str) -> tuple[list[tuple[str, str]], str | None] | None:
        r = self.db.execute("SELECT verzia, chyba FROM extrakcia WHERE hash=?", (hx,)).fetchone()
        if not r or r[0] != VERZIA_EXTRAKCIE:
            return None
        strany = self.db.execute("SELECT strana, text FROM texty WHERE hash=? ORDER BY poradie", (hx,)).fetchall()
        return [(s, t) for s, t in strany], r[1]

    def uloz_text(self, hx: str, strany: list[tuple[str, str]], chyba: str | None):
        self.db.execute("DELETE FROM texty WHERE hash=?", (hx,))
        self.db.executemany("INSERT INTO texty VALUES(?,?,?,?,?)",
                            [(hx, VERZIA_EXTRAKCIE, i, s, t) for i, (s, t) in enumerate(strany)])
        self.db.execute("INSERT OR REPLACE INTO extrakcia VALUES(?,?,?)", (hx, VERZIA_EXTRAKCIE, chyba))

    def ocr(self, hx: str, jazyky: str):
        r = self.db.execute("SELECT verzia, jazyky, motor, chyba, pdf FROM ocr WHERE hash=?", (hx,)).fetchone()
        if not r or r[0] != VERZIA_OCR or r[1] != jazyky:
            return None
        strany = dict(self.db.execute("SELECT strana, text FROM ocr_texty WHERE hash=?", (hx,)).fetchall())
        return strany, r[2], r[3], bool(r[4])

    def uloz_ocr(self, hx: str, jazyky: str, strany: dict, motor: str, chyba: str | None, pdf: bool):
        self.db.execute("DELETE FROM ocr_texty WHERE hash=?", (hx,))
        self.db.executemany("INSERT INTO ocr_texty VALUES(?,?,?)", [(hx, s, t) for s, t in strany.items()])
        self.db.execute("INSERT OR REPLACE INTO ocr VALUES(?,?,?,?,?,?)",
                        (hx, VERZIA_OCR, jazyky, motor, chyba, int(pdf)))

    def commit(self):
        self.db.commit()


# =============================================================================
# 5b. OCR – lokálne (Tesseract, OCRmyPDF). Nič neodchádza mimo počítača.
# =============================================================================

VERZIA_OCR = 3  # zvýš, ak sa zmení spôsob OCR -> výsledky v cache sa prepočítajú

TESSERACT_WINDOWS = [
    Path(r"C:\Program Files\Tesseract-OCR\tesseract.exe"),
    Path(r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe"),
    Path(os.environ.get("LOCALAPPDATA", "")) / "Programs" / "Tesseract-OCR" / "tesseract.exe",
]


@dataclass
class OcrNastavenie:
    zapnute: bool = True
    jazyky: str = "slk+ces+eng"
    tesseract: str = ""               # prázdne = nájde sám
    pouzit_ocrmypdf: bool = True      # PDF cez OCRmyPDF (vyrobí aj prehľadávateľné PDF)
    prehladavatelne_pdf: bool = True  # odkaz pri tagu povedie na prehľadávateľnú verziu skenu
    oprava_zamien: bool = True        # O/0, I/1, S/5, B/8 ... -> zhoda na kontrolu
    max_stran: int = 200
    casovy_limit_s: int = 900
    paralelne: int = 2
    dpi: int = 300


class Ocr:
    """Zistí, čo je nainštalované, a spúšťa OCR. Ak nič nie je, skript beží ďalej bez OCR."""

    def __init__(self, ns: OcrNastavenie):
        self.ns = ns
        self.tesseract: Path | None = None
        self.jazyky = ""
        self.ocrmypdf = False
        self.psm = "3"
        self._varovane_argumenty = False
        self.dovod = ""
        if not ns.zapnute:
            self.dovod = "OCR je vypnuté v nastaveniach"
            return
        self.tesseract = self._najdi_tesseract()
        if not self.tesseract:
            self.dovod = "Tesseract nie je nainštalovaný (alebo ho nenašiel)"
            return
        dostupne = self._jazyky()
        chcene = [j for j in ns.jazyky.split("+") if j]
        ok = [j for j in chcene if j in dostupne]
        chyba = [j for j in chcene if j not in dostupne]
        if chyba:
            log.warning("OCR: v Tesseracte chýbajú jazyky %s – používam %s", "+".join(chyba), "+".join(ok) or "eng")
        self.jazyky = "+".join(ok) or "eng"
        self.psm = "1" if "osd" in dostupne else "3"   # 1 = rozpozná aj otočené strany
        if ns.pouzit_ocrmypdf:
            import importlib.util
            self.ocrmypdf = importlib.util.find_spec("ocrmypdf") is not None

    @property
    def dostupne(self) -> bool:
        return self.tesseract is not None

    @property
    def popis(self) -> str:
        if not self.dostupne:
            return f"nedostupné – {self.dovod}"
        return f"Tesseract ({self.jazyky})" + (" + OCRmyPDF" if self.ocrmypdf else " (bez OCRmyPDF)")

    def _najdi_tesseract(self) -> Path | None:
        if self.ns.tesseract:
            p = Path(self.ns.tesseract)
            return p if p.is_file() else None
        kde = shutil.which("tesseract")
        if kde:
            return Path(kde)
        return next((p for p in TESSERACT_WINDOWS if p.is_file()), None)

    def _env(self) -> dict:
        env = dict(os.environ)
        env["PATH"] = str(self.tesseract.parent) + os.pathsep + env.get("PATH", "")   # aby ho našiel aj OCRmyPDF
        env.setdefault("OMP_THREAD_LIMIT", "1")
        return env

    def _jazyky(self) -> set[str]:
        r = subprocess.run([str(self.tesseract), "--list-langs"], capture_output=True, timeout=60)
        riadky = (r.stdout or b"").decode("utf-8", "ignore").splitlines()
        return {x.strip() for x in riadky[1:] if x.strip()}

    # -------------------------------------------------------------- obrázok
    def obrazok(self, cesta: Path) -> str:
        r = subprocess.run([str(self.tesseract), str(cesta), "stdout", "-l", self.jazyky, "--psm", self.psm],
                           capture_output=True, timeout=self.ns.casovy_limit_s, env=self._env())
        if r.returncode != 0:
            raise RuntimeError(f"Tesseract: {(r.stderr or b'').decode('utf-8', 'ignore').strip()[:300]}")
        return (r.stdout or b"").decode("utf-8", "ignore")

    # -------------------------------------------------------------- PDF
    def pdf(self, cesta: Path, strany_bez_textu: list[str], vystup_pdf: Path) -> tuple[dict[str, str], str, bool]:
        """Vráti ({strana: text}, motor, vytvorene_prehladavatelne_pdf)."""
        if self.ocrmypdf:
            try:
                return self._ocrmypdf(cesta, vystup_pdf), "OCRmyPDF", True
            except Exception as e:  # noqa: BLE001
                log.warning("   OCRmyPDF zlyhal (%s) – skúšam priamo Tesseract", str(e)[:200])
        return self._vykresli_a_rozpoznaj(cesta, strany_bez_textu), "Tesseract", False

    def _ocrmypdf(self, cesta: Path, vystup_pdf: Path) -> dict[str, str]:
        vystup_pdf.parent.mkdir(parents=True, exist_ok=True)
        tmp = vystup_pdf.with_suffix(".tmp.pdf")
        prikaz = [sys.executable, "-m", "ocrmypdf",
                  "-l", self.jazyky,
                  "--skip-text",            # OCR iba na stranách bez textu
                  "--rotate-pages",         # otočené skeny
                  "--rotate-pages-threshold", "2",       # predvolený prah je prísny – otočený sken neotočil
                  "--tesseract-pagesegmode", self.psm,   # 1 = Tesseract sám rozpozná orientáciu
                  "--optimize", "0",        # neprekomprimovať obrázky (archív)
                  "--output-type", "pdf",
                  "--jobs", str(max(1, self.ns.paralelne)),
                  "--quiet",
                  str(cesta), str(tmp)]
        r = subprocess.run(prikaz, capture_output=True, timeout=self.ns.casovy_limit_s, env=self._env())
        if r.returncode != 0 or not tmp.is_file():
            chyba = (r.stderr or b'').decode('utf-8', 'ignore').strip()
            if r.returncode == 2 and not self._varovane_argumenty:   # zlé argumenty = problém verzie OCRmyPDF
                self._varovane_argumenty = True
                log.warning("OCRmyPDF odmietol argumenty – skeny pôjdu priamo cez Tesseract (bez prehľadávateľného PDF).")
            raise RuntimeError(f"kód {r.returncode}: {chyba[-300:]}")
        tmp.replace(vystup_pdf)
        # text čítame z výsledného PDF – zachová čísla strán (pôvodný text + OCR vrstva)
        return {s: t for s, t in CitacTextu._pdf(vystup_pdf)}

    def _vykresli_a_rozpoznaj(self, cesta: Path, strany: list[str]) -> dict[str, str]:
        vys: dict[str, str] = {}
        with tempfile.TemporaryDirectory() as tmp:
            for s in strany[: self.ns.max_stran]:
                obr = Path(tmp) / f"s{s}.png"
                if _vykresli_stranu(cesta, int(s) - 1, obr, self.ns.dpi):
                    vys[s] = self.obrazok(obr)
        return vys


def _vykresli_stranu(pdf: Path, index: int, ciel: Path, dpi: int) -> bool:
    """Strana PDF -> PNG. Skúsi PyMuPDF, potom pypdfium2, nakoniec vložený obrázok cez pypdf."""
    try:
        import pymupdf  # type: ignore
    except ImportError:
        try:
            import fitz as pymupdf  # type: ignore
        except ImportError:
            pymupdf = None
    if pymupdf is not None:
        with pymupdf.open(pdf) as doc:
            doc[index].get_pixmap(dpi=dpi).save(str(ciel))
        return True
    try:
        import pypdfium2 as pdfium  # type: ignore
        doc = pdfium.PdfDocument(str(pdf))
        try:
            doc[index].render(scale=dpi / 72).to_pil().save(ciel)
        finally:
            doc.close()
        return True
    except ImportError:
        pass
    from pypdf import PdfReader
    obr = PdfReader(str(pdf)).pages[index].images
    if not obr:
        return False
    obr[0].image.save(ciel)
    return True


def spusti_ocr(dokumenty: dict[str, Dokument], texty: dict[str, list[tuple[str, str]]],
               cache: Cache, ocr: Ocr, n: Nastavenia) -> int:
    """Doplní text z OCR do `texty` (iba strany bez textu). Vráti počet nových OCR (bez cache)."""
    kandidati = [d for d in dokumenty.values() if d.stav in {"potrebuje OCR", "čiastočne OCR"}]
    if not kandidati:
        return 0
    if not ocr.dostupne:
        log.warning("OCR %s – %d dokumentov ostáva „potrebuje OCR“.", ocr.popis, len(kandidati))
        return 0
    log.info("OCR: %s | dokumentov na OCR: %d", ocr.popis, len(kandidati))
    pdf_dir = n.vystup / ".cache" / "ocr"
    nove = 0
    for i, d in enumerate(kandidati, 1):
        ulozene = cache.ocr(d.hash, ocr.jazyky)
        if ulozene is None:
            log.info("OCR [%d/%d] %s (strán: %d)", i, len(kandidati), d.nazov, len(d.strany_bez_textu))
            strany, motor, chyba, ma_pdf = {}, "", None, False
            try:
                if d.pripona == ".pdf":
                    strany, motor, ma_pdf = ocr.pdf(d.zdroje[0], d.strany_bez_textu, pdf_dir / f"{d.hash}.pdf")
                else:
                    strany, motor = {"1": ocr.obrazok(d.zdroje[0])}, "Tesseract"
            except Exception as e:  # noqa: BLE001
                chyba = f"{type(e).__name__}: {e}"[:500]
            cache.uloz_ocr(d.hash, ocr.jazyky, strany, motor, chyba, ma_pdf)
            cache.commit()
            nove += 1
        else:
            strany, motor, chyba, ma_pdf = ulozene
        d.ocr_motor = motor
        if chyba:
            d.stav, d.chyba = "chyba OCR", chyba
            continue
        # zlúčiť: pôvodný text strany ostáva, ak ho bolo dosť; inak text z OCR
        nove_texty = []
        for s, t in texty.get(d.hash, []) or [(s, "") for s in strany]:
            if s in d.strany_bez_textu and strany.get(s, "").strip():
                nove_texty.append((s, strany[s]))
                d.ocr_strany.add(s)
            else:
                nove_texty.append((s, t))
        texty[d.hash] = nove_texty
        if ma_pdf and n.ocr.prehladavatelne_pdf and (pdf_dir / f"{d.hash}.pdf").is_file():
            d.ocr_pdf = pdf_dir / f"{d.hash}.pdf"
        d.stav = "OK (OCR)" if d.ocr_strany else "OCR bez textu"
    return nove


# =============================================================================
# 6. INVENTÁR, EXTRAKCIA, PÁROVANIE
# =============================================================================

@dataclass
class Dokument:
    hash: str
    nazov: str                         # názov prvého nájdeného súboru
    pripona: str
    zdroje: list[Path] = field(default_factory=list)
    stav: str = ""                     # OK | OK (OCR) | potrebuje OCR | čiastočne OCR | OCR bez textu | chyba
    chyba: str = ""
    strany_bez_textu: list[str] = field(default_factory=list)
    ulozene: Path | None = None        # kam vedú odkazy (pri skene prehľadávateľná verzia)
    ulozene_original: Path | None = None
    ocr_strany: set = field(default_factory=set)
    ocr_motor: str = ""
    ocr_pdf: Path | None = None        # prehľadávateľné PDF z OCRmyPDF (v cache)


@dataclass
class Vazba:
    """Tag <-> dokument (agregované zo zhôd)."""
    kluc: str
    hash: str
    typy: set[str] = field(default_factory=set)
    strany: set[str] = field(default_factory=set)
    texty: set[str] = field(default_factory=set)
    kontext: str = ""
    stav: str = "auto"                 # auto | na kontrolu | potvrdené | zamietnuté


SILNE_TYPY = {"presná", "tolerovaná", "rozsah", "súčasť dlhšieho tagu", "názov súboru", "priečinok"}


def inventar(n: Nastavenia, cache: Cache) -> tuple[dict[str, Dokument], int]:
    dokumenty: dict[str, Dokument] = {}
    ignorovane = 0
    for zdroj in n.zdroje:
        if not zdroj.is_dir():
            log.error("Zdrojový priečinok neexistuje: %s", zdroj)
            continue
        for p in sorted(zdroj.rglob("*")):
            if not p.is_file():
                continue
            if n.vystup in p.parents:
                continue
            if p.name.startswith("~$") or p.suffix.lower() not in (n.typy_suborov | n.typy_len_nazov):
                ignorovane += 1
                continue
            hx = cache.hash_suboru(p)
            d = dokumenty.get(hx)
            if d is None:
                d = dokumenty[hx] = Dokument(hash=hx, nazov=p.name, pripona=p.suffix.lower())
            d.zdroje.append(p)
    cache.commit()
    return dokumenty, ignorovane


def extrahuj(dokumenty: dict[str, Dokument], cache: Cache, n: Nastavenia) -> dict[str, list[tuple[str, str]]]:
    citac = CitacTextu()
    texty: dict[str, list[tuple[str, str]]] = {}
    try:
        for i, d in enumerate(dokumenty.values(), 1):
            if d.pripona in n.typy_len_nazov and d.pripona not in n.typy_suborov:
                d.stav, texty[d.hash] = "iba názov", []
                continue
            ulozene = cache.text(d.hash)
            if ulozene is None:
                log.info("[%d/%d] čítam %s", i, len(dokumenty), d.nazov)
                try:
                    strany, chyba = citac.citaj(d.zdroje[0]), None
                except Exception as e:  # noqa: BLE001 – chceme pokračovať na ďalší súbor
                    strany, chyba = [], f"{type(e).__name__}: {e}"
                cache.uloz_text(d.hash, strany, chyba)
            else:
                strany, chyba = ulozene
            texty[d.hash] = strany
            if chyba:
                d.stav, d.chyba = "chyba", chyba
                continue
            if d.pripona == ".pdf" or d.pripona in OBRAZKY:
                # strana s menej ako N znakmi = pravdepodobne sken -> bude ju treba OCR
                d.strany_bez_textu = [s for s, t in strany if len(t.strip()) < n.min_znakov_na_stranu]
                if strany and len(d.strany_bez_textu) == len(strany):
                    d.stav = "potrebuje OCR"
                elif d.strany_bez_textu:
                    d.stav = "čiastočne OCR"
                else:
                    d.stav = "OK"
            else:
                d.stav = "OK" if any(t.strip() for _, t in strany) else "bez textu"
    finally:
        citac.zavri()
        cache.commit()
    return texty


def _kontext(text: str, z: int, k: int, okolie: int = 45) -> str:
    return re.sub(r"\s+", " ", ("…" if z > okolie else "") + text[max(0, z - okolie):k + okolie]).strip()


def paruj(dokumenty: dict[str, Dokument], texty: dict[str, list[tuple[str, str]]],
          parovac: Parovac, tagy: dict[str, Tag], zdroje: list[Path],
          parovac_ocr: Parovac | None = None) -> dict[tuple[str, str], Vazba]:
    vazby: dict[tuple[str, str], Vazba] = {}

    def pridaj(kluc, hx, typ, strana, najdene, kontext):
        v = vazby.setdefault((kluc, hx), Vazba(kluc, hx))
        v.typy.add(typ)
        if strana:
            v.strany.add(strana)
        v.texty.add(najdene)
        if not v.kontext:
            v.kontext = kontext

    def cesta_v_zdroji(p: Path) -> str:
        for koren in zdroje:
            if koren in p.parents:
                return " / ".join(p.parent.relative_to(koren).parts)
        return ""

    for d in dokumenty.values():
        # názov súboru a priečinky (pravidlo 6: tu aj krátke tagy idú automaticky)
        for zdroj in d.zdroje:
            for typ, retazec in (("názov súboru", zdroj.stem), ("priečinok", cesta_v_zdroji(zdroj))):
                for z in parovac.najdi(retazec):
                    pridaj(z.kluc, d.hash, "bez predpony" if z.typ == "bez predpony" else typ,
                           "", z.text, f"{typ}: {retazec}")
        # obsah
        for strana, text in texty.get(d.hash, []):
            t = normalizuj_text(text)
            z_ocr = strana in d.ocr_strany
            najdene = parovac.najdi(t)
            for z in najdene:
                pridaj(z.kluc, d.hash, z.typ, strana, z.text, _kontext(t, z.zaciatok, z.koniec))
                if z_ocr:
                    vazby[(z.kluc, d.hash)].typy.add("z OCR")
            # OCR zámeny (11V1O9A -> 11V109A): iba na stranách z OCR, vždy na kontrolu
            if z_ocr and parovac_ocr is not None:
                for z in parovac_ocr.najdi(t.translate(ZAMENY_OCR)):
                    # ak tu už je presne prečítaný tag (ten istý alebo iný), nič nenavrhujeme
                    if any((y.kluc == z.kluc and y.zaciatok < z.koniec and z.zaciatok < y.koniec)
                           or (y.zaciatok <= z.zaciatok and y.koniec >= z.koniec) for y in najdene):
                        continue
                    povodny = t[z.zaciatok:z.koniec]
                    pridaj(z.kluc, d.hash, "OCR oprava", strana, povodny, _kontext(t, z.zaciatok, z.koniec))

    # stav väzby
    for v in vazby.values():
        cez_nazov = v.typy & {"názov súboru", "priečinok"}
        silne = v.typy & SILNE_TYPY
        if cez_nazov:
            v.stav = "auto"
        elif tagy[v.kluc].nejednoznacny or not silne:
            v.stav = "na kontrolu"
        else:
            v.stav = "auto"
    return vazby


# =============================================================================
# 7. RUČNÁ KONTROLA (vystup/kontrola.xlsx – rozhodnutia sa zachovávajú medzi behmi)
# =============================================================================

KONTROLA_HLAVICKA = ["ID", "Tag No", "Dokument", "Strany", "Typ zhody", "Nájdený text", "Kontext", "Rozhodnutie (A/N)"]


def nacitaj_rozhodnutia(cesta: Path) -> dict[str, str]:
    if not cesta.is_file():
        return {}
    from openpyxl import load_workbook
    wb = load_workbook(cesta, read_only=True)
    try:
        ws = wb.active
        vys = {}
        for r in ws.iter_rows(min_row=2, values_only=True):
            if r and r[0] and len(r) >= 8 and r[7]:
                vys[str(r[0])] = str(r[7]).strip().upper()[:1]
        return vys
    finally:
        wb.close()


def aplikuj_rozhodnutia(vazby: dict[tuple[str, str], Vazba], rozhodnutia: dict[str, str]):
    for v in vazby.values():
        if v.stav != "na kontrolu":
            continue
        r = rozhodnutia.get(f"{v.kluc}|{v.hash[:16]}")
        if r == "A":
            v.stav = "potvrdené"
        elif r == "N":
            v.stav = "zamietnuté"


# =============================================================================
# 8. VÝSTUP – úložisko, odkazy, Excel
# =============================================================================

ZAKAZANE = str.maketrans({'<': '(', '>': ')', ':': '-', '"': 'in', '/': '-', '\\': '-', '|': '-', '?': '', '*': 'x'})
REZERVOVANE = {"CON", "PRN", "AUX", "NUL"} | {f"COM{i}" for i in range(1, 10)} | {f"LPT{i}" for i in range(1, 10)}


def bezpecny_nazov(nazov: str) -> str:
    nazov = re.sub(r"\s+", " ", nazov.translate(ZAKAZANE)).strip().rstrip(" .") or "BEZ_NAZVU"
    return "_" + nazov if nazov.upper() in REZERVOVANE else nazov


def nazov_priecinka(t: Tag, harok: str, n: Nastavenia) -> str:
    pred = n.predpona_jednotky
    bez = re.sub(r"^" + re.escape(pred) + r"[ -]?", "", t.zobrazenie) if pred and not t.potrubie else t.zobrazenie
    return bezpecny_nazov(n.format_priecinka.format(tag=t.zobrazenie, tag_bez_jednotky=bez, harok=harok))


def ulozisko_cesta(d: Dokument, n: Nastavenia) -> Path:
    if n.ulozisko == "zdroj":
        return d.zdroje[0]
    stem = bezpecny_nazov(Path(d.nazov).stem)[:80]
    return n.vystup / "_dokumenty" / f"{stem}__{d.hash[:8]}{d.pripona}"


def zvol_typ_odkazu(n: Nastavenia) -> str:
    if n.typ_odkazu != "auto":
        return n.typ_odkazu
    if os.name == "nt":
        global _LNK_OK
        if _LNK_OK is None:                  # skúška: povolí firma PowerShell + WScript.Shell?
            try:
                with tempfile.TemporaryDirectory() as tmp:
                    _vytvor_lnk_davkovo([(Path(tmp) / "test.lnk", Path(tmp))])
                    _LNK_OK = (Path(tmp) / "test.lnk").exists()
            except Exception as e:  # noqa: BLE001
                log.warning("Skratky .lnk sa nedajú vytvoriť (%s) – použijem .url odkazy.", e)
                _LNK_OK = False
        return "lnk" if _LNK_OK else "url"
    # Linux/Mac: symbolické odkazy, ak ich súborový systém podporuje (sieťové disky nie vždy)
    try:
        n.vystup.mkdir(parents=True, exist_ok=True)
        with tempfile.TemporaryDirectory(dir=n.vystup) as tmp:
            (Path(tmp) / "t").symlink_to("x")
        return "symlink"
    except OSError:
        return "url"


PRIPONA_ODKAZU = {"lnk": ".lnk", "url": ".url", "symlink": "", "markdown": ""}
_LNK_OK: bool | None = None


def planuj_odkazy(vazby, dokumenty, tagy, n: Nastavenia) -> dict[Path, Path]:
    """Vráti {cesta_odkazu: cieľ}. Iba väzby so stavom auto / potvrdené."""
    typ = zvol_typ_odkazu(n)
    plan: dict[Path, Path] = {}
    priecinky = {h.nazov: h.priecinok for h in n.harky}
    for v in vazby.values():
        if v.stav not in {"auto", "potvrdené"}:
            continue
        t, d = tagy[v.kluc], dokumenty[v.hash]
        for harok in t.harky:
            zlozka = n.vystup / "zariadenia" / priecinky[harok] / nazov_priecinka(t, harok, n)
            meno = bezpecny_nazov(d.nazov)
            odkaz = zlozka / (meno + PRIPONA_ODKAZU[typ])
            if odkaz in plan and plan[odkaz] != d.ulozene:     # dva rôzne dokumenty s rovnakým názvom
                odkaz = zlozka / (f"{Path(meno).stem}__{d.hash[:8]}{Path(meno).suffix}" + PRIPONA_ODKAZU[typ])
            plan[odkaz] = d.ulozene
    return plan


def _vytvor_lnk_davkovo(polozky: list[tuple[Path, Path]]):
    """Windows .lnk cez PowerShell (WScript.Shell) – jedno volanie pre všetky odkazy, bez pywin32."""
    if not polozky:
        return
    with tempfile.TemporaryDirectory() as tmp:
        data = Path(tmp) / "odkazy.json"
        skript = Path(tmp) / "odkazy.ps1"
        data.write_text(json.dumps([{"o": str(o), "c": str(c)} for o, c in polozky], ensure_ascii=False),
                        encoding="utf-8")
        skript.write_text(
            "param([string]$Data)\n"
            "$ws = New-Object -ComObject WScript.Shell\n"
            "$zoznam = Get-Content -LiteralPath $Data -Raw -Encoding UTF8 | ConvertFrom-Json\n"
            "foreach ($p in $zoznam) {\n"
            "  $s = $ws.CreateShortcut($p.o)\n"
            "  $s.TargetPath = $p.c\n"
            "  $s.WorkingDirectory = [System.IO.Path]::GetDirectoryName($p.c)\n"
            "  $s.Save()\n"
            "}\n", encoding="utf-8-sig")
        r = subprocess.run(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(skript),
                            "-Data", str(data)], capture_output=True, text=True)
        if r.returncode != 0:
            raise RuntimeError(f"PowerShell zlyhal: {r.stderr.strip()[:500]}")


def synchronizuj_odkazy(plan: dict[Path, Path], n: Nastavenia, vsetky_priecinky: set[Path]) -> tuple[int, int]:
    """Vytvorí chýbajúce odkazy a zmaže zastarané. Mení iba súbory v vystup/zariadenia, ktoré sám spravuje."""
    typ = zvol_typ_odkazu(n)
    manifest_cesta = n.vystup / ".cache" / "odkazy.json"
    stary = json.loads(manifest_cesta.read_text(encoding="utf-8")) if manifest_cesta.is_file() else {}

    for p in vsetky_priecinky | {o.parent for o in plan}:
        p.mkdir(parents=True, exist_ok=True)

    zmazane = 0
    for o_str, c_str in stary.items():
        o = Path(o_str)
        if plan.get(o) != Path(c_str) and (o.is_symlink() or o.is_file()):
            o.unlink()
            zmazane += 1

    if typ == "markdown":   # odkazy sú v README.md priečinkov (zapis_markdown), súbory sa nevytvárajú
        nove = [(o, c) for o, c in plan.items() if stary.get(str(o)) != str(c)]
    else:
        nove = [(o, c) for o, c in plan.items()
                if stary.get(str(o)) != str(c) or not (o.is_symlink() or o.exists())]
    if typ == "markdown":
        pass
    elif typ == "lnk":
        _vytvor_lnk_davkovo(nove)
    else:
        for o, c in nove:
            if o.is_symlink() or o.exists():
                o.unlink()
            if typ == "symlink":
                o.symlink_to(os.path.relpath(c, o.parent))
            else:  # .url – funguje všade na Windows, netreba PowerShell
                o.write_text(f"[InternetShortcut]\r\nURL={c.resolve().as_uri()}\r\n", encoding="utf-8")

    # prázdne priečinky po zmazaných odkazoch nechávame (tag existuje, len nemá dokumenty)
    manifest_cesta.parent.mkdir(parents=True, exist_ok=True)
    manifest_cesta.write_text(json.dumps({str(o): str(c) for o, c in plan.items()}, ensure_ascii=False, indent=0),
                              encoding="utf-8")
    return len(nove), zmazane


def uloz_dokumenty(dokumenty: dict[str, Dokument], potrebne: set[str], n: Nastavenia, ostro: bool) -> int:
    """Do úložiska ide iba dokument, ktorý má aspoň jednu väzbu. Ostatné ostanú v zdroji (sú v protokole)."""
    nove = 0
    for d in dokumenty.values():
        if d.hash not in potrebne:
            d.ulozene = d.zdroje[0]
            continue
        d.ulozene = d.ulozene_original = ulozisko_cesta(d, n)
        if ostro and n.ulozisko == "kopia" and not d.ulozene.exists():
            d.ulozene.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(d.zdroje[0], d.ulozene)
            nove += 1
        if d.ocr_pdf is not None:
            # prehľadávateľná verzia skenu vedľa originálu; odkazy povedú na ňu
            priecinok = n.vystup / ("_dokumenty" if n.ulozisko == "kopia" else "_ocr")
            ocr_ciel = priecinok / f"{bezpecny_nazov(Path(d.nazov).stem)[:80]}__{d.hash[:8]}.ocr.pdf"
            if ostro and not ocr_ciel.exists():
                ocr_ciel.parent.mkdir(parents=True, exist_ok=True)
                shutil.copy2(d.ocr_pdf, ocr_ciel)
            d.ulozene = ocr_ciel
    return nove



# ---------------------------------------------------------------------------- Markdown (čitateľné na GitHube)

def _md(text) -> str:
    return str(text if text is not None else "").replace("|", "\\|").replace("\n", " ")


def _md_odkaz(text: str, ciel: Path, od: Path) -> str:
    return f"[{_md(text)}](<{_rel(ciel, od).replace(os.sep, '/')}>)"


def _zapis_ak_iny(cesta: Path, obsah: str) -> bool:
    """Zapíše iba pri zmene – opakovaný beh nemení súbory (žiadne zbytočné zmeny v gite)."""
    if cesta.is_file() and cesta.read_text(encoding="utf-8") == obsah:
        return False
    cesta.parent.mkdir(parents=True, exist_ok=True)
    cesta.write_text(obsah, encoding="utf-8")
    return True


def zapis_markdown(n: Nastavenia, tagy, dokumenty, vazby, readme_v_priecinkoch: bool) -> int:
    """PREHLAD.md vo výstupe + (v režime markdown) README.md v každom priečinku tagu. Vráti počet zmenených."""
    priecinky = {h.nazov: h.priecinok for h in n.harky}
    koren = n.vystup
    po_tagu: dict[str, list[Vazba]] = defaultdict(list)
    for v in vazby.values():
        po_tagu[v.kluc].append(v)
    zmenene = 0

    def zlozka(t: Tag, harok: str) -> Path:
        return koren / "zariadenia" / priecinky[harok] / nazov_priecinka(t, harok, n)

    if readme_v_priecinkoch:
        for k, zoznam in po_tagu.items():
            t = tagy[k]
            pripojene = sorted((v for v in zoznam if v.stav in {"auto", "potvrdené"}),
                               key=lambda v: dokumenty[v.hash].nazov.lower())
            kontrola = sorted((v for v in zoznam if v.stav == "na kontrolu"),
                              key=lambda v: dokumenty[v.hash].nazov.lower())
            if not pripojene and not kontrola:
                continue
            for harok in t.harky:
                z = zlozka(t, harok)
                r = [f"# {_md(t.zobrazenie)}", "",
                     f"Hárok **{_md(harok)}** · dokumentov: **{len(pripojene)}** · na kontrolu: {len(kontrola)} · "
                     f"{_md_odkaz('← celý prehľad', koren / 'PREHLAD.md', z)}", ""]
                if pripojene:
                    r += ["Každý dokument je uložený iba raz v `_dokumenty` – tu sú len odkazy.", "",
                          "| Dokument | Strany | Typ zhody | Nájdený text |", "|---|---|---|---|"]
                    for v in pripojene:
                        d = dokumenty[v.hash]
                        r.append(f"| {_md_odkaz(d.nazov, d.ulozene, z)} | {_md(', '.join(sorted(v.strany)))} | "
                                 f"{_md(', '.join(sorted(v.typy)))} | {_md(', '.join(sorted(v.texty)))} |")
                if kontrola:
                    r += ["", "## Čaká na kontrolu", "",
                          "Nejasné zhody – rozhoduje sa v `kontrola.xlsx` (A = priradiť, N = nepriradiť).", "",
                          "| Dokument | Typ zhody | Nájdený text | Kontext |", "|---|---|---|---|"]
                    for v in kontrola:
                        d = dokumenty[v.hash]
                        r.append(f"| {_md_odkaz(d.nazov, d.ulozene or d.zdroje[0], z)} | "
                                 f"{_md(', '.join(sorted(v.typy)))} | {_md(', '.join(sorted(v.texty)))} | "
                                 f"{_md(v.kontext[:120])} |")
                zmenene += _zapis_ak_iny(z / "README.md", "\n".join(r) + "\n")

    # PREHLAD.md
    pripojene_v = [v for v in vazby.values() if v.stav in {"auto", "potvrdené"}]
    suborov = sum(len(d.zdroje) for d in dokumenty.values())
    po_starom = sum(len(dokumenty[v.hash].zdroje) for v in pripojene_v)
    r = [f"# Prehľad – {_md(n.projekt)}", "",
         f"Vygenerované nástrojom VUZ Sorter {VERZIA}. Každý dokument je uložený **raz**, "
         "zariadenia naň iba odkazujú.", "",
         "## Súhrn", "",
         "| Položka | Hodnota |", "|---|---|",
         f"| Súborov v zdrojoch | {suborov} |",
         f"| Unikátnych dokumentov (podľa obsahu) | {len(dokumenty)} |",
         f"| Duplicitné kópie v zdrojoch | {suborov - len(dokumenty)} |",
         f"| Dokumenty prečítané cez OCR | {sum(1 for d in dokumenty.values() if d.ocr_strany)} |",
         f"| Tagov s aspoň jedným dokumentom | {len({v.kluc for v in pripojene_v})} |",
         f"| Pôvodný postup (kópia do každého tagu) | {po_starom} súborov |",
         f"| Teraz: dokumentov v úložisku | {len({v.hash for v in vazby.values()})} |",
         f"| Teraz: väzieb tag ↔ dokument | {len(pripojene_v)} |",
         f"| Na kontrolu | {sum(1 for v in vazby.values() if v.stav == 'na kontrolu')} |", "",
         "## Zariadenia", "",
         "| Hárok | Tag No | Dokumentov | Na kontrolu |", "|---|---|---|---|"]
    for t in sorted(tagy.values(), key=lambda t: (t.harky[0], t.zobrazenie)):
        zoz = po_tagu.get(t.kluc, [])
        pocet = sum(1 for v in zoz if v.stav in {"auto", "potvrdené"})
        kontr = sum(1 for v in zoz if v.stav == "na kontrolu")
        if not pocet and not kontr:
            continue
        for harok in t.harky:
            nazov = t.zobrazenie
            if readme_v_priecinkoch:
                nazov = _md_odkaz(t.zobrazenie, zlozka(t, harok) / "README.md", koren)
            r.append(f"| {_md(harok)} | {nazov if readme_v_priecinkoch else _md(nazov)} | {pocet} | {kontr} |")
    bez = sorted(t.zobrazenie for t in tagy.values()
                 if not any(v.stav in {"auto", "potvrdené"} for v in po_tagu.get(t.kluc, [])))
    r += ["", "## Dokumenty", "", "| Dokument | Stav čítania | OCR | Tagy |", "|---|---|---|---|"]
    for d in sorted(dokumenty.values(), key=lambda d: d.nazov.lower()):
        tagy_d = sorted(tagy[v.kluc].zobrazenie for v in pripojene_v if v.hash == d.hash)
        ciel = d.ulozene or d.zdroje[0]
        r.append(f"| {_md_odkaz(d.nazov, ciel, koren)} | {_md(d.stav)} | {_md(d.ocr_motor)} | "
                 f"{_md(', '.join(tagy_d)) or '–'} |")
    r += ["", f"## Tagy bez dokumentov ({len(bez)})", "", _md(", ".join(bez)) or "–", ""]
    zmenene += _zapis_ak_iny(koren / "PREHLAD.md", "\n".join(r))
    return zmenene

# ---------------------------------------------------------------------------- Excel

def _rel(ciel: Path, od: Path) -> str:
    try:
        return os.path.relpath(ciel, od)
    except ValueError:  # iný disk
        return str(ciel)


def zapis_kontrolu(vazby, dokumenty, tagy, n: Nastavenia, rozhodnutia: dict[str, str]) -> Path | None:
    from openpyxl import Workbook
    from openpyxl.worksheet.datavalidation import DataValidation

    cesta = n.vystup / "kontrola.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "Na kontrolu"
    ws.append(KONTROLA_HLAVICKA)
    polozky = [v for v in vazby.values() if v.stav in {"na kontrolu", "potvrdené", "zamietnuté"}]
    for v in sorted(polozky, key=lambda v: (tagy[v.kluc].zobrazenie, dokumenty[v.hash].nazov)):
        idk = f"{v.kluc}|{v.hash[:16]}"
        d = dokumenty[v.hash]
        ws.append([idk, tagy[v.kluc].zobrazenie, d.nazov, ", ".join(sorted(v.strany)), ", ".join(sorted(v.typy)),
                   ", ".join(sorted(v.texty)), v.kontext, rozhodnutia.get(idk, "")])
        c = ws.cell(ws.max_row, 3)
        c.hyperlink = _rel(d.ulozene or d.zdroje[0], n.vystup)
        c.style = "Hyperlink"
    dv = DataValidation(type="list", formula1='"A,N"', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(f"H2:H{max(ws.max_row, 2)}")
    _formatuj(ws, [28, 22, 40, 10, 22, 22, 70, 18])
    ws.column_dimensions["A"].hidden = True
    try:
        wb.save(cesta)
        return cesta
    except PermissionError:
        log.error("kontrola.xlsx je otvorený v Exceli – zavri ho a spusti znova. Rozhodnutia sa nestratili.")
        return None


def _formatuj(ws, sirky):
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="1F4E78")
        c.alignment = Alignment(wrap_text=True, vertical="center")
    for i, s in enumerate(sirky, 1):
        ws.column_dimensions[get_column_letter(i)].width = s
    ws.freeze_panes = "A2"
    if ws.max_row > 1:
        ws.auto_filter.ref = ws.dimensions


def zapis_protokol(cesta: Path, n: Nastavenia, tagy, dokumenty, vazby, plan, suhrn: list[tuple[str, object]]):
    from openpyxl import Workbook
    from openpyxl.styles import Font

    wb = Workbook()
    ws = wb.active
    ws.title = "Súhrn"
    ws.append(["Položka", "Hodnota"])
    for k, h in suhrn:
        ws.append([k, h])
    _formatuj(ws, [62, 40])
    for r in ws.iter_rows(min_row=2):
        if str(r[0].value).startswith("■"):
            r[0].font = Font(bold=True)

    # Zariadenia
    ws = wb.create_sheet("Zariadenia")
    ws.append(["Tag No", "Hárky", "Dokumentov (auto/potvrdené)", "Na kontrolu", "Priečinok"])
    priecinky = {h.nazov: h.priecinok for h in n.harky}
    po_tagu = defaultdict(lambda: [0, 0])
    for v in vazby.values():
        if v.stav in {"auto", "potvrdené"}:
            po_tagu[v.kluc][0] += 1
        elif v.stav == "na kontrolu":
            po_tagu[v.kluc][1] += 1
    for t in sorted(tagy.values(), key=lambda t: (-po_tagu[t.kluc][0], t.zobrazenie)):
        harok = t.harky[0]
        zl = n.vystup / "zariadenia" / priecinky[harok] / nazov_priecinka(t, harok, n)
        ws.append([t.zobrazenie, ", ".join(t.harky), po_tagu[t.kluc][0], po_tagu[t.kluc][1], nazov_priecinka(t, harok, n)])
        c = ws.cell(ws.max_row, 5)
        c.hyperlink = _rel(zl, cesta.parent)
        c.style = "Hyperlink"
    _formatuj(ws, [42, 28, 26, 13, 45])

    # Zhody (index tag <-> dokument)
    ws = wb.create_sheet("Index tag-dokument")
    ws.append(["Tag No", "Dokument", "Strany", "Typ zhody", "Stav", "Nájdený text", "Kontext"])
    for v in sorted(vazby.values(), key=lambda v: (tagy[v.kluc].zobrazenie, dokumenty[v.hash].nazov)):
        d = dokumenty[v.hash]
        ws.append([tagy[v.kluc].zobrazenie, d.nazov, ", ".join(sorted(v.strany, key=lambda s: (len(s), s))),
                   ", ".join(sorted(v.typy)), v.stav, ", ".join(sorted(v.texty)), v.kontext])
        c = ws.cell(ws.max_row, 2)
        c.hyperlink = _rel(d.ulozene or d.zdroje[0], cesta.parent)
        c.style = "Hyperlink"
    _formatuj(ws, [36, 42, 10, 26, 13, 26, 80])

    # Dokumenty
    ws = wb.create_sheet("Dokumenty")
    ws.append(["Dokument", "Stav čítania", "Strany bez textu", "OCR strany", "OCR motor", "Počet tagov",
               "Kópií v zdrojoch", "Zdrojové cesty", "Uložené ako", "Prehľadávateľné PDF", "Chyba", "SHA-256"])
    tagov_v_dok = defaultdict(int)
    for v in vazby.values():
        if v.stav in {"auto", "potvrdené"}:
            tagov_v_dok[v.hash] += 1
    for d in sorted(dokumenty.values(), key=lambda d: d.nazov.lower()):
        orig = d.ulozene_original or d.ulozene
        ws.append([d.nazov, d.stav, ", ".join(d.strany_bez_textu),
                   ", ".join(sorted(d.ocr_strany, key=lambda s: (len(s), s))), d.ocr_motor,
                   tagov_v_dok[d.hash], len(d.zdroje), "\n".join(str(z) for z in d.zdroje),
                   _rel(orig, cesta.parent) if orig else "",
                   _rel(d.ulozene, cesta.parent) if d.ulozene and d.ulozene != orig else "",
                   d.chyba, d.hash])
    _formatuj(ws, [42, 15, 16, 12, 12, 12, 15, 70, 50, 50, 50, 20])

    # Tagy bez dokumentov / MasterData
    ws = wb.create_sheet("Odkazy")
    ws.append(["Odkaz", "Cieľ"])
    for o, c in sorted(plan.items()):
        ws.append([_rel(o, cesta.parent), _rel(c, cesta.parent)])
    _formatuj(ws, [90, 70])

    try:
        wb.save(cesta)
    except PermissionError:
        alt = cesta.with_name(cesta.stem + datetime.now().strftime("_%H%M%S") + cesta.suffix)
        wb.save(alt)
        log.warning("%s je otvorený, protokol uložený ako %s", cesta.name, alt.name)
        cesta = alt
    return cesta


# =============================================================================
# 9. HLAVNÝ BEH
# =============================================================================

def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description="LTA triedenie dokumentov podľa Tag No")
    ap.add_argument("--config", default=str(Path(__file__).resolve().parent / "config" / "demo.ini"))
    ap.add_argument("--ostro", action="store_true", help="skutočne uloží dokumenty a vytvorí odkazy")
    ap.add_argument("--bez-cache", action="store_true", help="znovu prečíta všetky texty")
    a = ap.parse_args(argv)
    for prud in (sys.stdout, sys.stderr):   # diakritika v konzole aj pri presmerovaní výstupu
        try:
            prud.reconfigure(errors="replace")
        except Exception:  # noqa: BLE001
            pass

    n = nacitaj_nastavenia(Path(a.config).resolve())
    if n.vystup.name == "ukazka_vystupu" and not all(
            "demo_data" in p.parts for p in [n.masterdata, *n.zdroje]):
        raise SystemExit("Priečinok ukazka_vystupu sa nahráva na GitHub – smie obsahovať IBA fiktívne demo_data.")
    n.vystup.mkdir(parents=True, exist_ok=True)
    logging.basicConfig(level=logging.INFO, format="%(message)s",
                        handlers=[logging.StreamHandler(sys.stdout),
                                  logging.FileHandler(n.vystup / "posledny_beh.log", "w", encoding="utf-8")],
                        force=True)
    rezim = "OSTRÝ BEH" if a.ostro else "TEST – nič sa nekopíruje ani nevytvára"
    log.info("LTA triedenie %s | projekt %s | %s", VERZIA, n.projekt, rezim)
    log.info("=" * 78)

    cache_cesta = n.vystup / ".cache" / "lta.sqlite"
    if a.bez_cache and cache_cesta.exists():
        cache_cesta.unlink()
    cache = Cache(cache_cesta)
    try:
        tagy = nacitaj_masterdata(n)
        parovac = Parovac(tagy, n)

        dokumenty, ignorovane = inventar(n, cache)
        suborov = sum(len(d.zdroje) for d in dokumenty.values())
        log.info("Súborov: %d | unikátnych dokumentov: %d | ignorovaných typov: %d", suborov, len(dokumenty), ignorovane)

        texty = extrahuj(dokumenty, cache, n)
        ocr = Ocr(n.ocr)
        nove_ocr = spusti_ocr(dokumenty, texty, cache, ocr, n)
        treba_zameny = n.ocr.oprava_zamien and any(d.ocr_strany for d in dokumenty.values())
        parovac_ocr = parovac_zamen(tagy, n) if treba_zameny else None
        vazby = paruj(dokumenty, texty, parovac, tagy, n.zdroje, parovac_ocr)
        rozhodnutia = nacitaj_rozhodnutia(n.vystup / "kontrola.xlsx")
        aplikuj_rozhodnutia(vazby, rozhodnutia)

        nove_kopie = uloz_dokumenty(dokumenty, {v.hash for v in vazby.values()}, n, a.ostro)
        plan = planuj_odkazy(vazby, dokumenty, tagy, n)

        priecinky = {h.nazov: h.priecinok for h in n.harky}
        vsetky = set()
        if n.vytvorit_prazdne_priecinky:
            vsetky = {n.vystup / "zariadenia" / priecinky[h] / nazov_priecinka(t, h, n) for t in tagy.values() for h in t.harky}

        vytvorene = zmazane = 0
        if a.ostro:
            vytvorene, zmazane = synchronizuj_odkazy(plan, n, vsetky)
            md = zapis_markdown(n, tagy, dokumenty, vazby, zvol_typ_odkazu(n) == "markdown")
            if zvol_typ_odkazu(n) == "markdown":
                vytvorene = md

        # súhrn
        stavy = defaultdict(int)
        for d in dokumenty.values():
            stavy[d.stav] += 1
        vazby_stav = defaultdict(int)
        for v in vazby.values():
            vazby_stav[v.stav] += 1
        po_starom = sum(len(dokumenty[v.hash].zdroje) for v in vazby.values() if v.stav in {"auto", "potvrdené"})
        tagov_s_dok = len({v.kluc for v in vazby.values() if v.stav in {"auto", "potvrdené"}})
        suhrn: list[tuple[str, object]] = [
            ("■ Beh", ""),
            ("Režim", rezim),
            ("Čas", datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
            ("Verzia", VERZIA),
            ("■ MasterData", ""),
            ("Tagov v rozsahu (unikátnych)", len(tagy)),
            ("Z toho nejednoznačných (pravidlo 6)", sum(t.nejednoznacny for t in tagy.values())),
            ("Tagov s aspoň jedným dokumentom", tagov_s_dok),
            ("■ Dokumenty", ""),
            ("Súborov v zdrojoch", suborov),
            ("Unikátnych dokumentov (podľa obsahu)", len(dokumenty)),
            ("Duplicitné kópie v zdrojoch", suborov - len(dokumenty)),
            ("Ignorované (nepodporovaný typ)", ignorovane),
            ("OCR", ocr.popis),
            ("OCR: dokumenty spracované v tomto behu (ostatné z cache)", nove_ocr),
            ("Dokumenty bez tagu (ostali iba v zdroji)", len(dokumenty) - len({v.hash for v in vazby.values()})),
            *[(f"Stav čítania: {k}", v) for k, v in sorted(stavy.items())],
            ("■ Väzby tag ↔ dokument", ""),
            *[(f"Stav: {k}", v) for k, v in sorted(vazby_stav.items())],
            ("■ Referencie namiesto kópií", ""),
            ("Pôvodný postup (kópia do každého tagu): súborov", po_starom),
            ("Teraz: dokumentov v úložisku (každý raz)", len({v.hash for v in vazby.values()})),
            ("Teraz: odkazov", len(plan)),
        ]
        if a.ostro:
            suhrn += [("Nové kópie do úložiska", nove_kopie), ("Vytvorené / aktualizované odkazy", vytvorene),
                      ("Zmazané zastarané odkazy", zmazane)]

        kontrola = zapis_kontrolu(vazby, dokumenty, tagy, n, rozhodnutia)
        protokol = zapis_protokol(n.vystup / ("index.xlsx" if a.ostro else "protokol_TEST.xlsx"),
                                  n, tagy, dokumenty, vazby, plan, suhrn)
        log.info("=" * 78)
        for k, h in suhrn:
            log.info("%s%s", k.replace("■ ", "\n"), f": {h}" if h != "" else "")
        log.info("\nProtokol:  %s", protokol)
        if kontrola:
            log.info("Kontrola:  %s  (vyplň stĺpec Rozhodnutie A/N a spusti znova)", kontrola)
        if not a.ostro:
            log.info("\nTEST – nič sa nevytvorilo. Ostrý beh: python %s --ostro", Path(__file__).name)
        return 0
    finally:
        cache.db.close()
        for h in list(logging.getLogger().handlers):
            if isinstance(h, logging.FileHandler):
                h.close()
                logging.getLogger().removeHandler(h)


if __name__ == "__main__":
    dvojklik = len(sys.argv) == 1
    kod = 1
    try:
        kod = main()
    except SystemExit as e:
        print(e)
    except Exception:
        import traceback
        traceback.print_exc()
    if dvojklik:
        input("\nStlač Enter na zatvorenie...")
    sys.exit(kod)
