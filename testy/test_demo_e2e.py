# -*- coding: utf-8 -*-
"""End-to-end test celého behu na fiktívnych demo dátach.

Overuje: TEST režim nič nevytvorí, ostrý beh uloží každý dokument raz a vytvorí odkazy,
opakovaný beh nič nezduplikuje.  Spustenie:  python -m unittest discover testy
"""

import configparser
import shutil
import sys
import tempfile
import unittest
from pathlib import Path

KOREN = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(KOREN))
import lta_triedenie as lt  # noqa: E402


def _suhrn(xlsx: Path) -> dict:
    from openpyxl import load_workbook
    wb = load_workbook(xlsx, read_only=True)
    try:
        return {r[0]: r[1] for r in wb["Súhrn"].iter_rows(min_row=2, values_only=True)}
    finally:
        wb.close()


class TestDemoBeh(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.tmp = Path(tempfile.mkdtemp(prefix="vuz_sorter_"))
        shutil.copytree(KOREN / "demo_data", cls.tmp / "demo_data")
        cp = configparser.ConfigParser(inline_comment_prefixes=(";", "#"))
        cp.read(KOREN / "config" / "demo.ini", encoding="utf-8")
        cp["projekt"]["masterdata"] = str(cls.tmp / "demo_data" / "MasterData_DEMO.xlsx")
        cp["projekt"]["zdroje"] = str(cls.tmp / "demo_data" / "neroztriedene")
        cp["projekt"]["vystup"] = str(cls.tmp / "vystup")
        cp["projekt"]["typ_odkazu"] = "url"          # funguje na každom OS bez PowerShellu
        if not cp.has_section("ocr"):
            cp.add_section("ocr")
        cp["ocr"]["zapnute"] = "nie"                 # čísla nezávisia od toho, či je Tesseract
        cls.ini = cls.tmp / "test.ini"
        with cls.ini.open("w", encoding="utf-8") as f:
            cp.write(f)
        cls.vystup = cls.tmp / "vystup"

    @classmethod
    def tearDownClass(cls):
        import logging
        for h in list(logging.getLogger().handlers):    # uvoľniť log súbor (Windows)
            h.close()
            logging.getLogger().removeHandler(h)
        shutil.rmtree(cls.tmp, ignore_errors=True)

    def test_1_test_rezim_nic_nevytvori(self):
        self.assertEqual(lt.main(["--config", str(self.ini)]), 0)
        self.assertTrue((self.vystup / "protokol_TEST.xlsx").is_file())
        self.assertFalse((self.vystup / "_dokumenty").exists())
        self.assertFalse((self.vystup / "zariadenia").exists())
        s = _suhrn(self.vystup / "protokol_TEST.xlsx")
        self.assertEqual(s["Unikátnych dokumentov (podľa obsahu)"], 18)
        self.assertEqual(s["Duplicitné kópie v zdrojoch"], 1)
        self.assertEqual(s["Stav: auto"], 40)
        self.assertEqual(s["Stav: na kontrolu"], 4)

    def test_2_ostry_beh_a_opakovanie(self):
        self.assertEqual(lt.main(["--config", str(self.ini), "--ostro"]), 0)
        s = _suhrn(self.vystup / "index.xlsx")
        self.assertEqual(s["Nové kópie do úložiska"], 13)
        self.assertEqual(s["Vytvorené / aktualizované odkazy"], 41)
        self.assertEqual(len(list((self.vystup / "_dokumenty").iterdir())), 13)
        odkazy = list((self.vystup / "zariadenia").rglob("*.url"))
        self.assertEqual(len(odkazy), 41)
        self.assertEqual({o.name for o in (self.vystup / "zariadenia" / "Static" / "11V109A").iterdir()},
                         {"Revizna_sprava_2024_odlucovace.pdf.url", "Supis_zariadeni_jednotka_11.xlsx.url"})

        # opakovaný beh – nič nové
        self.assertEqual(lt.main(["--config", str(self.ini), "--ostro"]), 0)
        s = _suhrn(self.vystup / "index.xlsx")
        self.assertEqual(s["Nové kópie do úložiska"], 0)
        self.assertEqual(s["Vytvorené / aktualizované odkazy"], 0)
        self.assertEqual(s["Zmazané zastarané odkazy"], 0)



class TestMarkdownUkazka(unittest.TestCase):
    """Výstup pre GitHub: README.md v priečinkoch zariadení + PREHLAD.md, ochrana ukazka_vystupu."""

    def _ini(self, tmp: Path, vystup: Path, masterdata: Path, zdroj: Path) -> Path:
        cp = configparser.ConfigParser(inline_comment_prefixes=(";", "#"))
        cp.read(KOREN / "config" / "ukazka.ini", encoding="utf-8")
        cp["projekt"]["masterdata"] = str(masterdata)
        cp["projekt"]["zdroje"] = str(zdroj)
        cp["projekt"]["vystup"] = str(vystup)
        cp["ocr"]["zapnute"] = "nie"
        ini = tmp / "u.ini"
        with ini.open("w", encoding="utf-8") as f:
            cp.write(f)
        return ini

    def tearDown(self):
        import logging
        for h in list(logging.getLogger().handlers):
            if isinstance(h, logging.FileHandler):
                h.close()
                logging.getLogger().removeHandler(h)

    def test_readme_a_prehlad(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp = Path(tmp)
            shutil.copytree(KOREN / "demo_data", tmp / "demo_data")
            vystup = tmp / "ukazka_vystupu"
            ini = self._ini(tmp, vystup, tmp / "demo_data" / "MasterData_DEMO.xlsx", tmp / "demo_data" / "neroztriedene")
            self.assertEqual(lt.main(["--config", str(ini), "--ostro"]), 0)
            prehlad = (vystup / "PREHLAD.md").read_text(encoding="utf-8")
            self.assertIn("zariadenia/Static/11V109A/README.md", prehlad)
            readme = (vystup / "zariadenia" / "Static" / "11V109A" / "README.md").read_text(encoding="utf-8")
            self.assertIn("../../../_dokumenty/Revizna_sprava_2024_odlucovace__", readme)
            self.assertEqual(list((vystup / "zariadenia").rglob("*.url")), [])      # žiadne súbory odkazov
            # opakovaný beh nič nemení
            pred = {p: p.stat().st_mtime_ns for p in vystup.rglob("*.md")}
            self.assertEqual(lt.main(["--config", str(ini), "--ostro"]), 0)
            self.assertEqual(pred, {p: p.stat().st_mtime_ns for p in vystup.rglob("*.md")})

    def test_ukazka_odmietne_realne_data(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp = Path(tmp)
            realne = tmp / "realne"
            realne.mkdir()
            ini = self._ini(tmp, tmp / "ukazka_vystupu", realne / "MasterData.xlsx", realne)
            with self.assertRaises(SystemExit):
                lt.main(["--config", str(ini)])


if __name__ == "__main__":
    unittest.main()
