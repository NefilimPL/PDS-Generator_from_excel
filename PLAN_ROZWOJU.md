# Plan Rozwoju PDS Generator

Plan roboczy przygotowany na podstawie przegladu repo z dnia 2026-04-26. Ten plik ma byc centralnym backlogiem projektu: po wykonaniu zadania zmieniaj `[ ]` na `[x]` i w razie potrzeby dopisuj date, numer PR albo krotka notatke o zakresie zmian.

## Oznaczenia

- `[ ]` do zrobienia
- `[x]` zrobione
- Priorytet `P0` krytyczne, `P1` wysokie, `P2` srednie, `P3` niski
- Typy: `BUG`, `SEC`, `TECH-DEBT`, `TEST`, `QOL`, `PERF`, `DOCS`, `OPS`, `FEATURE`

## Stan bazowy

- [x] `FEATURE` Aplikacja ma tryb GUI i headless (`pds_gui.pyw`, `pds_headless.py`).
- [x] `FEATURE` Istnieje zapis/odczyt konfiguracji z kopia zapasowa w AppData oraz migracja starszych sciezek (`pds_generator/gui/config_io.py`, `pds_generator/app_paths.py`).
- [x] `FEATURE` Dziala eksport PDF z obsluga obrazow, grup, zaleznosci layoutu i kompresji (`pds_generator/gui/pdf_export.py`, `pds_generator/layout_dependencies.py`, `pds_generator/pdf_settings.py`).
- [x] `FEATURE` Jest podglad PDF i obsluga indeksu obrazow oraz auto-zoom obrazow (`pds_generator/gui/pdf_preview.py`, `pds_generator/image_index.py`, `pds_generator/image_auto_zoom.py`).
- [x] `FEATURE` Jest obsluga aktualizacji z GitHub i launcher dla uruchomienia na Windows (`pds_generator/github_utils.py`, `launcher.py`).
- [x] `SEC` Sekrety e-mail maja warstwe szyfrowania/maskowania i fallback ENV (`pds_generator/gui/mailer.py`, `README.md`).
- [x] `OPS` Jest podstawowe logowanie dla GUI i headless z retencja logow (`pds_gui.pyw`, `pds_headless.py`).
- [x] `TEST` Repo zawiera testy dla kluczowych obszarow: layout dependencies, image auto-zoom, text layout, value sources, fragmenty mailera i parity GUI/headless (`tests/`).

## Sugerowana kolejnosc realizacji

1. Najpierw zamknac tematy `P0`: bezpieczna aktualizacja, prawdziwe blokady plikow i powtarzalne uruchamianie testow.
2. Potem domknac `P1` zwiazane z jakoscia i release: CI, walidacja konfiguracji, lepsza obsluga bledow i jasny support matrix.
3. Dopiero na ustabilizowanej bazie wejsc w wiekszy refactor modulow GUI, mailera i eksportu PDF.
4. Nastepnie realizowac UX/QoL i optymalizacje wydajnosci.

## Stabilnosc i bezpieczenstwo

- [ ] `P0` `BUG/SEC` Zabezpieczyc mechanizm aktualizacji przed nadpisaniem lokalnych zmian i polowicznym overlayem ZIP.
- [ ] `P0` `BUG/TECH-DEBT` Podjac decyzje dla systemu blokad plikow: wlaczyc i przetestowac realne blokowanie albo usunac martwy kod; obecnie `LOCKS_ENABLED = False`.
- [ ] `P1` `BUG` Ujednolicic obsluge wyjatkow i raportowanie bledow, z ograniczeniem szerokich `except Exception` w krytycznych sciezkach.
- [ ] `P1` `SEC` Przejrzec przeplyw sekretow miedzy GUI, `config.json`, AppData i ENV; dopisac testy regresyjne dla precedencji zrodel.
- [ ] `P1` `BUG/QOL` Dodac walidacje konfiguracji przed generowaniem PDF i przed uruchomieniem headless.
- [ ] `P1` `BUG/OPS` Uporzadkowac timeouty, retry i komunikaty dla operacji sieciowych: GitHub, obrazy z URL, Graph API, SMTP.
- [ ] `P2` `BUG` Dodac bezpieczne odzyskiwanie po uszkodzonym `config.json`: rotacja backupow, prompt naprawczy, restore ostatniej poprawnej konfiguracji.
- [ ] `P2` `SEC/OPS` Sprawdzic, czy aktualizacja i auto-installer sa bezpieczne w srodowiskach firmowych z politykami blokujacymi instalatory i skrypty.

## Jakosc i testy

- [ ] `P0` `TEST/OPS` Dodac powtarzalne srodowisko developerskie: `requirements-dev.txt` albo `pyproject.toml`, `pytest`, instrukcje uruchamiania testow.
- [ ] `P1` `TEST/OPS` Dodac CI uruchamiajace testy i lint dla kazdego PR.
- [ ] `P1` `TEST` Rozszerzyc testy o `excel_io.py`, szczegolnie fallback dla formul, typowanie wartosci i edge-case'y arkuszy.
- [ ] `P1` `TEST` Rozszerzyc testy o `config_io.py`, migracje configu, backup configu i bledy odczytu/zapisu.
- [ ] `P1` `TEST` Rozszerzyc testy o `github_utils.py`, `locks.py` oraz `requirements_installer.py`.
- [ ] `P1` `TEST` Dodac test end-to-end dla eksportu headless na przykladowym workbooku z obrazami i warunkami.
- [ ] `P2` `TEST` Dodac smoke testy GUI dla podstawowych przeplywow: otwarcie Excela, zapis configu, podglad, eksport, anulowanie pracy.
- [ ] `P2` `TEST/TECH-DEBT` Wprowadzic statyczna analize (`ruff`) i formatowanie kodu jako standard repo.
- [ ] `P2` `TEST/OPS` Ustalic minimalny prog coverage i raportowanie pokrycia w CI.

## Architektura i utrzymanie

- [ ] `P1` `TECH-DEBT` Rozbic `pds_generator/gui/gui.py` na mniejsze moduly: shell aplikacji, canvas/editor, background jobs, update flow, mail settings.
- [ ] `P1` `TECH-DEBT` Rozbic `pds_generator/gui/mailer.py` na warstwy: konfiguracja, auth, transport, szyfrowanie sekretow, reminder workflow.
- [ ] `P1` `TECH-DEBT` Rozbic `pds_generator/gui/pdf_export.py` na pipeline renderowania, obsluge obrazow, raportowanie i orchestration batcha.
- [ ] `P1` `TECH-DEBT` Oddzielic logike domenowa od `tkinter` i `messagebox`, tak aby GUI bylo cienka warstwa nad serwisami.
- [ ] `P2` `TECH-DEBT` Wprowadzic jawny model konfiguracji (np. dataclass / typed schema) wspolny dla GUI i headless.
- [ ] `P2` `TECH-DEBT` Ujednolicic zarzadzanie zadaniami w tle: kolejki, cancel tokeny, statusy i aktualizacje UI.
- [ ] `P2` `TECH-DEBT` Przejrzec warstwe kompatybilnosci i legacy fallbacki; zostawic tylko te, ktore sa testowane i rzeczywiscie potrzebne.
- [ ] `P3` `TECH-DEBT` Uporzadkowac nazewnictwo, stale i granice odpowiedzialnosci miedzy modulami `elements`, `groups`, `excel_io`, `value_sources`.

## UX i Quality of Life

- [ ] `P1` `QOL` Dodac panel "health check projektu" pokazujacy brakujace kolumny, niedzialajace zaleznosci, brakujace obrazy i problemy z e-mailem.
- [ ] `P2` `QOL` Dodac autosave i snapshoty konfiguracji, z mozliwoscia szybkiego cofniecia do ostatniej stabilnej wersji.
- [ ] `P2` `QOL` Ulepszyc edytor zaleznosci i grup: lepsza wizualizacja zaleznosci, filtrowanie, podglad skutkow zmian.
- [ ] `P2` `QOL` Dodac wyszukiwarke pol, kolumn i grup oraz szybkie przejscie do elementu na canvasie.
- [ ] `P2` `QOL` Rozbudowac skroty klawiaturowe i operacje masowe dla zaznaczonych elementow.
- [ ] `P2` `QOL` Dodac podglad dla wybranego wiersza wraz z czytelnym porownaniem wartosci z Excela i wartosci po transformacjach.
- [ ] `P3` `FEATURE/QOL` Dodac biblioteke szablonow ukladow i presetow dokumentow.
- [ ] `P3` `QOL` Dodac ostatnio otwierane pliki/projekty oraz szybkie wznowienie pracy po starcie aplikacji.

## Wydajnosc i skalowalnosc

- [ ] `P1` `PERF` Zmierzyc realna wydajnosc dla duzych workbookow i duzej liczby obrazow; ustalic benchmarki bazowe.
- [ ] `P2` `PERF` Dodac cache dyskowy dla pobranych obrazow z URL oraz deduplikacje wielokrotnie uzywanych assetow.
- [ ] `P2` `PERF` Zoptymalizowac odswiezanie canvasu i podgladu PDF, aby nie przeliczac calej strony po drobnych zmianach.
- [ ] `P2` `PERF` Zrobic przebudowe indeksu obrazow bardziej przyrostowa, z odswiezaniem tylko zmienionych katalogow.
- [ ] `P3` `PERF/QOL` Dodac tryb wznowienia batcha i generowanie tylko zmienionych / nieprzetworzonych wierszy.
- [ ] `P3` `PERF` Zweryfikowac ustawienia `ProcessPoolExecutor` i limity pamieci przy duzych eksportach.

## Operacje, release i dokumentacja

- [ ] `P1` `OPS/DOCS` Ustalic oficjalny support matrix: Windows-only vs cross-platform, wymagania Pythona, zaleznosci systemowe, ograniczenia trybu GUI/headless.
- [ ] `P1` `OPS` Zdefiniowac proces release: wersjonowanie, changelog, smoke-test przed wydaniem, artefakty do dystrybucji.
- [ ] `P1` `OPS/BUG` Uczynic zrodlo brancha aktualizacji konfigurowalnym; nie polegac na sztywno wpisanym `MAIN`.
- [ ] `P1` `DOCS` Dodac sekcje "Developer setup" i "Troubleshooting" do README.
- [ ] `P2` `DOCS/TEST` Dodac przykladowy workbook, przykladowy `config.json` i dane testowe do szybszych testow manualnych.
- [ ] `P2` `OPS` Rozbudowac `.gitignore` i housekeeping repo o typowe artefakty developerskie, buildy i cache narzedzi.
- [ ] `P2` `OPS` Dodac automatyczne sprawdzenie integralnosci builda launchera i opis procesu budowania artefaktu.
- [ ] `P3` `DOCS` Dodac szablony issue / bug report / feature request oraz prosty changelog produktu.

## Notatki z przegladu repo

- Najwieksze moduly do rozbicia: `pds_generator/gui/gui.py` (~3721 linii), `pds_generator/gui/mailer.py` (~2894), `pds_generator/gui/pdf_export.py` (~2072), `pds_generator/gui/dependencies_editor.py` (~966), `pds_generator/groups.py` (~940).
- Testy sa obecne, ale repo nie ma obecnie jawnej konfiguracji narzedzi developerskich potrzebnych do ich uruchamiania.
- Kod jest wyraznie "Windows-first": AppData, DPAPI, `windows_auth.py`, launcher z prywatnym Python Runtime i build `launcher.exe`.
- W projekcie jest juz sporo dobrych fundamentow produktowych, ale kolejny etap powinien byc nastawiony bardziej na stabilizacje, testowalnosc i rozdzielenie odpowiedzialnosci modulow niz na dopisywanie nowych funkcji.
