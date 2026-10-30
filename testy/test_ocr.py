# -*- coding: utf-8 -*-
"""Testy OCR.  Spustenie:  python -m unittest discover testy

- oprava zámen a zlúčenie textu z OCR sa testujú bez Tesseractu (náhradný OCR objekt)
- skutočné OCR sa testuje iba ak je Tesseract nainštalovaný (inak sa test preskočí)
"""

import sys
import tempfile
import unittest
from pathlib import Path

KOREN = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(KOREN))
sys.path.insert(0, str(Path(__file__).resolve().parent))
import lta_triedenie as lt  # noqa: E402
from test_parovanie import _parovac  # noqa: E402


class TestOpravaZamen(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.p, cls.tagy = _parovac(["11V109A", "11V109B", "11C101", "11D101", "11E101", "11P110B"])
        n = lt.Nastavenia(projekt="t", masterdata=Path("."), zdroje=[], vystup=Path("."), ulozisko="kopia",
                          typ_odkazu="auto", format_priecinka="{tag}", vytvorit_prazdne_priecinky=False,
                          typy_suborov=set(), typy_len_nazov=set(), min_znakov_na_stranu=50, len_v_rozsahu=True,
                          predpona_jednotky="11", max_dlzka_kratkeho_tagu=5, nejednoznacne_navyse=[],
                          povolit_bez_predpony=True)
        cls.pz = lt.parovac_zamen(cls.tagy, n)

    def najdi(self, text):
        t = lt.normalizuj_text(text)
        return {self.tagy[z.kluc].zobrazenie for z in self.pz.najdi(t.translate(lt.ZAMENY_OCR))}

    def test_typicke_zameny(self):
        self.assertIn("11V109A", self.najdi("nádrž 11V1O9A"))      # O namiesto 0
        self.assertIn("11E101", self.najdi("11El0l"))               # l namiesto 1
        self.assertIn("11P110B", self.najdi("11P1108"))             # 8 namiesto B
        self.assertIn("11E101", self.najdi("11£101"))               # £ namiesto E
        self.assertIn("11V109B", self.najdi("11VI09B"))             # I namiesto 1

    def test_nejednoznacne_dostanu_oboch_kandidatov(self):
        self.assertEqual(self.najdi("110101"), {"11C101", "11D101"})  # C/D prečítané ako 0

    def test_spravne_precitany_tag_nema_alternativy(self):
        t = "protokol 11C101"
        najdene = self.p.najdi(t)
        alternativy = [z for z in self.pz.najdi(t.translate(lt.ZAMENY_OCR))
                       if not any(y.zaciatok <= z.zaciatok and y.koniec >= z.koniec for y in najdene)]
        self.assertEqual(alternativy, [])

    def test_bez_ocr_textu_nic(self):
        self.assertEqual(self.najdi("bežný text bez tagov"), set())


class FalosneOcr:
    """Náhrada za Tesseract – vráti vopred daný text."""
    dostupne = True
    jazyky = "slk"
    popis = "falošné OCR"

    def __init__(self, text):
        self.text = text
        self.volania = 0

    def obrazok(self, cesta):
        self.volania += 1
        return self.text

    def pdf(self, cesta, strany, vystup_pdf):
        self.volania += 1
        return {s: self.text for s in strany}, "Tesseract", False


class TestZlucenieACache(unittest.TestCase):
    def test_ocr_text_sa_doplni_a_cachuje(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp = Path(tmp)
            obr = tmp / "stitok.jpg"
            obr.write_bytes(b"x")
            d = lt.Dokument(hash="a" * 64, nazov=obr.name, pripona=".jpg", zdroje=[obr],
                            stav="potrebuje OCR", strany_bez_textu=["1"])
            texty = {d.hash: [("1", "")]}
            n = lt.Nastavenia(projekt="t", masterdata=tmp, zdroje=[tmp], vystup=tmp / "vystup", ulozisko="kopia",
                              typ_odkazu="url", format_priecinka="{tag}", vytvorit_prazdne_priecinky=False,
                              typy_suborov=set(), typy_len_nazov=set(), min_znakov_na_stranu=50,
                              len_v_rozsahu=True, predpona_jednotky="11", max_dlzka_kratkeho_tagu=5,
                              nejednoznacne_navyse=[], povolit_bez_predpony=True)
            n.ocr = lt.OcrNastavenie()
            cache = lt.Cache(n.vystup / ".cache" / "t.sqlite")
            ocr = FalosneOcr("Výrobný štítok 11E101")
            self.assertEqual(lt.spusti_ocr({d.hash: d}, texty, cache, ocr, n), 1)
            self.assertEqual(d.stav, "OK (OCR)")
            self.assertEqual(texty[d.hash], [("1", "Výrobný štítok 11E101")])
            self.assertEqual(d.ocr_strany, {"1"})

            # druhý beh: z cache, OCR sa nevolá
            d2 = lt.Dokument(hash=d.hash, nazov=obr.name, pripona=".jpg", zdroje=[obr],
                             stav="potrebuje OCR", strany_bez_textu=["1"])
            texty2 = {d.hash: [("1", "")]}
            self.assertEqual(lt.spusti_ocr({d.hash: d2}, texty2, cache, ocr, n), 0)
            self.assertEqual(ocr.volania, 1)
            self.assertEqual(texty2[d.hash][0][1], "Výrobný štítok 11E101")
            cache.db.close()

    def test_bez_tesseractu_nic_nespadne(self):
        o = lt.Ocr(lt.OcrNastavenie(tesseract=str(Path(tempfile.gettempdir()) / "neexistuje.exe")))
        self.assertFalse(o.dostupne)
        self.assertIn("nedostupné", o.popis)


@unittest.skipUnless(lt.Ocr(lt.OcrNastavenie()).dostupne, "Tesseract nie je nainštalovaný")
class TestSkutocneOcr(unittest.TestCase):
    def test_obrazok_s_tagom(self):
        from PIL import Image, ImageDraw, ImageFont
        font = None
        for f in ("C:/Windows/Fonts/arial.ttf", "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"):
            if Path(f).exists():
                font = ImageFont.truetype(f, 64)
                break
        if font is None:
            self.skipTest("nenašiel som TTF písmo")
        with tempfile.TemporaryDirectory() as tmp:
            obr = Path(tmp) / "stitok.png"
            img = Image.new("RGB", (1400, 300), "white")
            ImageDraw.Draw(img).text((60, 100), "Zariadenie 11V109A", fill="black", font=font)
            img.save(obr)
            text = lt.Ocr(lt.OcrNastavenie()).obrazok(obr)
        self.assertIn("11V109A", text.replace(" ", ""))


if __name__ == "__main__":
    unittest.main()
