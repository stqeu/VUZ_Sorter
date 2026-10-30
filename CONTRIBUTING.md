# Ako spolu pracujeme

## Pravidlá

- **Do `main` sa nepushuje priamo.** Každá zmena ide cez vetvu a pull request (PR).
- `main` je vždy funkčný: testy prechádzajú a demo beh funguje.
- PR schvaľuje **ten druhý** (1 review). Zlučuje autor po schválení.
- Do repa nikdy nepatria reálne dokumenty, MasterData ani `*.local.ini`.

## Postup pri zmene

```powershell
git switch main
git pull                                  # mať najnovší main
git switch -c feature/ocr-tesseract       # nová vetva

# ... úpravy ...
python -m unittest discover testy         # musí prejsť
git status                                # skontrolovať, že tam nie sú reálne dáta
git add -A
git commit -m "OCR: Tesseract pre skenované PDF"
git push -u origin feature/ocr-tesseract
```

Potom na GitHube: **Compare & pull request** → vyplniť šablónu → požiadať druhého o review.
Po schválení a zelených testoch: **Squash and merge** a vetvu zmazať.

## Názvy vetiev

| Prefix | Na čo | Príklad |
|---|---|---|
| `feature/` | nová funkcia | `feature/ocr-tesseract` |
| `fix/` | oprava chyby | `fix/kolizia-nazvov-odkazov` |
| `docs/` | iba dokumentácia | `docs/navod-hrp7` |

## Keď sa zmení `main`, kým pracujem na vetve

```powershell
git switch main
git pull
git switch feature/moja-vetva
git merge main          # vyriešiť prípadné konflikty, otestovať, pushnúť
```

## Zmeny v párovaní

Každá zmena pravidiel = nový alebo upravený test v `testy/test_parovanie.py`.
Ak sa menia čísla v demo behu, upraviť aj `testy/test_demo_e2e.py` a v PR vysvetliť prečo.
