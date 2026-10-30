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

OCR zatiaľ nie je súčasťou. Súbory a strany bez textu sa označia ako „potrebuje OCR“.
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

VERZIA = "0.1.0-demo"
VERZIA_EXTRAKCIE = 1  # zvýš, ak sa zmení spôsob čítania textu -> cache sa prepočíta

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
        if pymupdf is not None:
            with pymupdf.open(cesta) as doc:
                return [(str(i + 1), p.get_text("text", sort=True)) for i, p in enumerate(doc)]
        from pypdf import PdfReader
        r = PdfReader(str(cesta))
        return [(str(i + 1), p.extract_text() or "") for i, p in enumerate(r.pages)]

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

    def commit(self):
        self.db.commit()


# =============================================================================
# 6. INVENTÁR, EXTRAKCIA, PÁROVANIE
# =============================================================================

@dataclass
class Dokument:
    hash: str
    nazov: str                         # názov prvého nájdeného súboru
    pripona: str
    zdroje: list[Path] = field(default_factory=list)
    stav: str = ""                     # OK | potrebuje OCR | čiastočne OCR | chyba
    chyba: str = ""
    strany_bez_textu: list[str] = field(default_factory=list)
    ulozene: Path | None = None        # kde leží jediná kópia


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
          parovac: Parovac, tagy: dict[str, Tag], zdroje: list[Path]) -> dict[tuple[str, str], Vazba]:
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
            for z in parovac.najdi(t):
                pridaj(z.kluc, d.hash, z.typ, strana, z.text, _kontext(t, z.zaciatok, z.koniec))

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


PRIPONA_ODKAZU = {"lnk": ".lnk", "url": ".url", "symlink": ""}
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

    nove = [(o, c) for o, c in plan.items()
            if stary.get(str(o)) != str(c) or not (o.is_symlink() or o.exists())]
    if typ == "lnk":
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
        d.ulozene = ulozisko_cesta(d, n)
        if ostro and n.ulozisko == "kopia" and not d.ulozene.exists():
            d.ulozene.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(d.zdroje[0], d.ulozene)
            nove += 1
    return nove


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
    ws.append(["Dokument", "Stav čítania", "Strany bez textu (OCR)", "Počet tagov", "Kópií v zdrojoch",
               "Zdrojové cesty", "Uložené ako", "Chyba", "SHA-256"])
    tagov_v_dok = defaultdict(int)
    for v in vazby.values():
        if v.stav in {"auto", "potvrdené"}:
            tagov_v_dok[v.hash] += 1
    for d in sorted(dokumenty.values(), key=lambda d: d.nazov.lower()):
        ws.append([d.nazov, d.stav, ", ".join(d.strany_bez_textu), tagov_v_dok[d.hash], len(d.zdroje),
                   "\n".join(str(z) for z in d.zdroje), _rel(d.ulozene, cesta.parent) if d.ulozene else "",
                   d.chyba, d.hash])
    _formatuj(ws, [42, 15, 20, 12, 15, 70, 50, 50, 20])

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
        vazby = paruj(dokumenty, texty, parovac, tagy, n.zdroje)
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
