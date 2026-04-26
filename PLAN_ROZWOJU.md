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

## Rozwiniecie kluczowych tematow

### 1. Bezpieczna aktualizacja aplikacji

Cel:
Wykluczyc sytuacje, w ktorej update nadpisuje lokalne zmiany, zostawia repo w stanie czesciowo zaktualizowanym albo miesza pliki z roznych wersji.

Zakres:
- Rozdzielic dwa scenariusze: aktualizacja repo git i aktualizacja z paczki ZIP.
- Przed aktualizacja wykrywac lokalne modyfikacje i informowac uzytkownika, czy aktualizacja bedzie bezpieczna.
- Dla wariantu ZIP nie rozpakowywac "na zywo" do katalogu roboczego, tylko najpierw do katalogu tymczasowego, potem wykonac atomowa podmiane wybranych plikow.
- Jawnie zdefiniowac liste plikow i katalogow, ktorych updater nie moze ruszac: np. `config.json`, logi, cache, lokalne artefakty uzytkownika.
- Dodac plan rollbacku, jezeli aktualizacja przerwie sie w polowie.

Definicja done:
- Aktualizacja nie rusza lokalnych danych uzytkownika.
- Aplikacja nie zostaje w stanie "pol wersji".
- Bledy aktualizacji sa czytelne i nie wymagaja recznego sprzatania repo.

Status 2026-04-26:
- [x] Rozdzielono preflight i wykonanie dla `git` oraz `ZIP`; `git` blokuje update przy brudnym checkoutcie i uzywa tylko `fetch` + `merge --ff-only`.
- [x] Dodano manifest instalacji ZIP, liste sciezek chronionych, staging do `.pds-updater/` oraz rollback/recovery po przerwanym update.
- [x] GUI pokazuje przed aktualizacja, czy update jest bezpieczny, a testy regresyjne obejmuja manifest, rollback i komunikaty preflight.

### 2. Realne blokady plikow i odporne zapisy konfiguracji

Cel:
Zapobiec rownoleglym zapisom `config.json`, jednoczesnym uruchomieniom eksportu na tych samych danych i przypadkowemu uszkodzeniu plikow roboczych.

Zakres:
- Zdecydowac, czy locki maja byc funkcja obowiazkowa, czy opcjonalna, ale dzialajaca.
- Jezeli zostaja: wlaczyc je naprawde, przetestowac stale locki, przejmowanie locka i crash recovery.
- Zapis konfiguracji robic przez plik tymczasowy i rename zamiast nadpisywania docelowego pliku wprost.
- Czytelnie rozdzielic lock dla konfiguracji, lock dla arkusza i lock dla procesu eksportu.
- Dla headless i GUI utrzymac te same zasady zachowania.

Definicja done:
- Nie da sie latwo uszkodzic `config.json` przez przerwany zapis.
- Uzytkownik dostaje jasny komunikat, kto trzyma blokade i co moze zrobic.
- Locki dzialaja tak samo w typowych przeplywach GUI i headless.

### 3. Refactor najwiekszych modulow

Cel:
Zmniejszyc koszt utrzymania i ryzyko regresji przez rozbicie modulow, ktore dzis lacza UI, logike biznesowa, IO, siec i obsluge bledow w jednym miejscu.

Zakres:
- `gui.py`: wydzielic shell aplikacji, kontrolery akcji UI, obsluge background jobs i logike canvasu.
- `mailer.py`: rozdzielic konfiguracje, szyfrowanie sekretow, pobieranie tokenow, wysylke SMTP/Graph i reminder workflow.
- `pdf_export.py`: wydzielic render jednego dokumentu, przygotowanie danych, obsluge obrazow, raportowanie wynikow i orchestration batcha.
- Ustalic publiczne API miedzy modulami, zamiast odwolania do wielu atrybutow `app`.
- Tam, gdzie sie da, zastapic bezposrednie `messagebox` wynikiem/wyjatkiem zwracanym do warstwy UI.

Definicja done:
- Kazdy duzy obszar ma mniejsze moduly o jasnej odpowiedzialnosci.
- Testy jednostkowe da sie pisac bez stawiania calego `tkinter`.
- Mniej zmian wymaga dotykania kilku niepowiazanych miejsc naraz.

### 4. Jawny model konfiguracji wspolny dla GUI i headless

Cel:
Usunac rozjazdy miedzy tym, co zapisuje GUI, a tym, co interpretuje tryb headless i eksport PDF.

Zakres:
- Zdefiniowac jeden model konfiguracji: pola, grupy, zaleznosci, ustawienia PDF, mail, image dirs, tracking, feature flags.
- Przy ladowaniu stosowac normalizacje i walidacje w jednym miejscu.
- Wymusic wersjonowanie formatu configu i przygotowac migracje starszych wersji.
- Wszystkie sciezki startowe maja korzystac z tej samej warstwy odczytu i walidacji.

Definicja done:
- Ten sam `config.json` daje przewidywalny wynik w GUI i headless.
- Bledna konfiguracja jest wykrywana przed startem generowania.
- Dodanie nowego pola do configu nie wymaga zmian w pieciu roznych miejscach.

### 5. CI, testy i srodowisko developerskie

Cel:
Sprawic, zeby jakosc projektu nie zalezal od recznego sprawdzania zmian lokalnie.

Zakres:
- Dodac jawny zestaw narzedzi developerskich i sposob uruchamiania testow.
- Wlaczyc automatyczne odpalanie testow przy push/PR.
- Podzielic testy na szybkie jednostkowe, integracyjne i e2e dla headless.
- Dodac fixture z przykladowym Excelem, obrazami i przykladowym configiem.
- Ustalic minimalny standard: testy, lint, ewentualnie coverage dla najwazniejszych modulow.

Definicja done:
- Nowy dev potrafi uruchomic testy bez zgadywania.
- Kazdy PR ma automatyczny sygnal, czy nie zepsul bazowych przeplywow.
- Najbardziej ryzykowne obszary maja testy regresyjne.

### 6. Panel "health check projektu"

Cel:
Zamiast odkrywac problemy dopiero podczas eksportu, pokazac je od razu po wczytaniu projektu.

Zakres:
- Sprawdzanie brakujacych kolumn, pustych mapowan, nieistniejacych pol w zaleznosciach i grupach.
- Weryfikacja obrazow: brak pliku, zly URL, konflikt wielu trafien, uszkodzony format.
- Weryfikacja konfiguracji maila i sekretow bez wysylki produkcyjnej.
- Czytelny raport z poziomami: blad, ostrzezenie, informacja.
- Szybkie przejscie z raportu do problematycznego elementu lub ustawienia.

Definicja done:
- Uzytkownik przed eksportem wie, co na pewno nie zadziala.
- Typowe problemy sa diagnozowane w jednym miejscu.
- Zmniejsza sie liczba bledow odkrywanych dopiero po wygenerowaniu batcha.

### 7. Wydajnosc dla duzych wsadow i obrazow

Cel:
Utrzymac przewidywalny czas generowania przy wiekszych plikach Excel i duzej liczbie assetow graficznych.

Zakres:
- Zmierzyc czasy: ladowanie Excela, przebudowa indeksu obrazow, render jednej strony, batch export.
- Dodac cache dla pobranych obrazow z URL i lepsze reuse juz przetworzonych obrazow lokalnych.
- Ograniczyc pelne przeliczenia podgladu po drobnych zmianach w edytorze.
- Sprawdzic, czy wieloprocesowosc realnie pomaga, czy tylko zwieksza koszt serializacji i zuzycie RAM.

Definicja done:
- Istnieja benchmarki przed/po.
- Da sie wskazac, ktore operacje sa najdrozsze.
- Optymalizacje sa oparte na pomiarach, nie na zgadywaniu.

## Stabilnosc i bezpieczenstwo

- [x] `P0` `BUG/SEC` Zabezpieczono mechanizm aktualizacji przed nadpisaniem lokalnych zmian i polowicznym overlayem ZIP. (2026-04-26: osobne flow `git`/`ZIP`, preflight, staging, rollback, recovery)
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
