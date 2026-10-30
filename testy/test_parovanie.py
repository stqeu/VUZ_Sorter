# -*- coding: utf-8 -*-
"""Testy pravidiel párovania 1 – 7.  Spustenie:  python -m unittest discover testy"""

import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import lta_triedenie as lt  # noqa: E402


def _parovac(zoznam, potrubia=()):
    n = lt.Nastavenia(projekt="t", masterdata=Path("."), zdroje=[], vystup=Path("."), ulozisko="kopia",
                      typ_odkazu="auto", format_priecinka="{tag}", vytvorit_prazdne_priecinky=False,
                      typy_suborov=set(), typy_len_nazov=set(), min_znakov_na_stranu=50, len_v_rozsahu=True, predpona_jednotky="11",
                      max_dlzka_kratkeho_tagu=5, nejednoznacne_navyse=["PLUG-VALVE-*"], povolit_bez_predpony=True)
    tagy = {}
    for z in zoznam:
        tagy[lt.kanonicky(z)] = lt.Tag(lt.kanonicky(z), z, ["Static"])
    for z in potrubia:
        t = lt.tag_potrubia(z)
        tagy[lt.kanonicky(t)] = lt.Tag(lt.kanonicky(t), t, ["Pipes"], potrubie=True)
    return lt.Parovac(tagy, n), tagy


TAGY = ["11V109", "11V109A", "11V109B", "11V109C", "11E101", "11E102", "11A101", "11A101A", "11A101B",
        "11A101C", "11A101D", "11E126A", "57DRM2-11E126A", "E1", "PLUG-VALVE-1", "11-N2-0,6-1285",
        "FAM6-4_1", "11TE1347A1"]
POTRUBIA = ["11-DEA-N-1105-DN50-B1E (od 11P110A/B do", 'TBP-0279-3"-A1B', "Potrubný rozvod A/e-KSO"]


class TestPravidla(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.p, cls.tagy = _parovac(TAGY, POTRUBIA)

    def najdi(self, text):
        return {(self.tagy[z.kluc].zobrazenie, z.typ) for z in self.p.najdi(text)}

    def tagy_v(self, text):
        return {t for t, _ in self.najdi(text)}

    # pravidlo 1 – oddeľovače
    def test_1_oddelovace(self):
        self.assertEqual(self.tagy_v("nádoba 11 V-109 A bola"), {"11V109A"})
        self.assertEqual(self.tagy_v("výmenník 11-E-101."), {"11E101"})
        self.assertIn(("11V109A", "tolerovaná"), self.najdi("11 V-109 A"))
        self.assertIn(("11V109A", "presná"), self.najdi("11V109A"))

    def test_1_spojka_a_nie_je_suffix(self):
        self.assertEqual(self.tagy_v("zariadenie 11A101 a potrubie"), {"11A101"})
        self.assertEqual(self.tagy_v("nádrž 11V109 v prevádzke"), {"11V109"})

    # pravidlo 2 – hranice
    def test_2_hranice(self):
        self.assertEqual(self.tagy_v("11V109A"), {"11V109A"})            # nie 11V109
        self.assertEqual(self.tagy_v("11V1090 a 11E1011 a X11E101"), set())
        self.assertEqual(self.tagy_v("(11E101)"), {"11E101"})
        self.assertEqual(self.tagy_v("subor_11E101_rev2"), {"11E101"})    # podčiarkovník je hranica

    # pravidlo 3 – rozsahy
    def test_3_rozsahy(self):
        self.assertEqual(self.tagy_v("nádoby 11V109A-C"), {"11V109A", "11V109B", "11V109C"})
        self.assertEqual(self.tagy_v("11A101A - D"), {"11A101A", "11A101B", "11A101C", "11A101D"})
        self.assertEqual(self.tagy_v("11V109A/C"), {"11V109A", "11V109C"})          # lomka = výpočet
        self.assertEqual(self.tagy_v("11V109A-11V109C"), {"11V109A", "11V109B", "11V109C"})
        self.assertIn(("11V109B", "rozsah"), self.najdi("11V109A-C"))
        self.assertEqual(self.tagy_v("11V109A - Nádrž"), {"11V109A"})               # nie rozsah do N

    # pravidlo 4 – bez predpony
    def test_4_bez_predpony(self):
        self.assertEqual(self.najdi("pozri V109A"), {("11V109A", "bez predpony")})
        self.assertEqual(self.najdi("pozri 11-V109A"), {("11V109A", "tolerovaná")})

    # pravidlo 5 – tag vnútri dlhšieho tagu
    def test_5_sucast(self):
        n = self.najdi("vývod 57DRM2-11E126A")
        self.assertIn(("57DRM2-11E126A", "presná"), n)
        self.assertIn(("11E126A", "súčasť dlhšieho tagu"), n)

    # pravidlo 6 – krátke tagy sa nájdu, ale sú označené ako nejednoznačné
    def test_6_kratke(self):
        self.assertEqual(self.tagy_v("prístroj E1 a E-1"), {"E1"})
        self.assertTrue(self.tagy[lt.kanonicky("PLUG-VALVE-1")].nejednoznacny is False)  # flag nastaví MasterData
        self.assertEqual(self.tagy_v("PLUG-VALVE-1 a PLUG-VALVE-12"), {"PLUG-VALVE-1"})

    # pravidlo 7 – potrubia
    def test_7_potrubia(self):
        self.assertEqual(self.tagy_v("linka 11-DEA-N-1105-DN50-B1E"), {"11-DEA-N-1105-DN50-B1E"})
        self.assertEqual(self.tagy_v("linka 11-DEA-N-1105-DN 50-B1E"), {"11-DEA-N-1105-DN50-B1E"})
        self.assertEqual(self.tagy_v('linka 11-DEA-N-1105-2"-B1E'), {"11-DEA-N-1105-DN50-B1E"})
        self.assertEqual(self.tagy_v("TBP-0279-DN80-A1B"), {'TBP-0279-3"-A1B'})
        self.assertEqual(self.tagy_v('TBP-0279-3"-A1B'), {'TBP-0279-3"-A1B'})
        self.assertEqual(self.tagy_v("11-DEA-N-1105-DN80-B1E"), set())             # iné DN
        self.assertEqual(self.tagy_v("Potrubný rozvod A/e-KSO"), {"Potrubný rozvod A/e-KSO"})

    def test_ostatne_tvary(self):
        self.assertEqual(self.tagy_v("dusík 11-N2-0,6-1285"), {"11-N2-0,6-1285"})
        self.assertEqual(self.tagy_v("11-N2-06-1285"), set())                       # 0,6 != 06
        self.assertEqual(self.tagy_v("FAM6-4_1"), {"FAM6-4_1"})
        self.assertEqual(self.tagy_v("teplota 11TE1347A1"), {"11TE1347A1"})
        self.assertEqual(self.tagy_v("11–E–101"), {"11E101"})                       # dlhé pomlčky z Wordu


if __name__ == "__main__":
    unittest.main()
