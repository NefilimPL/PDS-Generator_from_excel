# PDS-Generator_from_excel

Rozbudowany generator dokumentów **PDS** (PDF) oparty na danych z arkuszy Excel.  
Aplikacja udostępnia graficzny edytor umożliwiający projektowanie układu
strony poprzez przeciąganie pól tekstowych, obrazów oraz elementów grupowych.
Skonfigurowany projekt może zostać wykorzystany do seryjnego tworzenia plików
PDF – dla każdego wiersza arkusza powstaje oddzielny dokument.

## Funkcjonalności
- Wczytywanie wielu arkuszy Excela (`.xlsx`) i przypisywanie kolumn do pól na
  stronie.
- Pola statyczne z własną treścią i możliwością dowolnego formatowania.
- Przeciąganie i skalowanie elementów na siatce z przyciąganiem do kroków
  (domyślnie co 5 pkt) oraz podglądem linii wyrównania.
- Edycja kroju, rozmiaru i stylu czcionki, kolorów tła i tekstu, wyrównania oraz
  warstwy (kolejności rysowania) każdego elementu.
- Obsługa obrazów lokalnych lub zdalnych (URL podany w komórce Excela).
- Grupowanie pól w *obszary* z własnym podglądem i indywidualną konfiguracją,
  w tym warunkowym ukrywaniem elementów zależnie od zawartości innych pól.
- Zapisywanie i wczytywanie konfiguracji do pliku `config.json` w katalogu
  wybranego pliku Excel (tworzona jest także kopia zapasowa w
  `~/.pds_generator/config.json`) – zapamiętywane są m.in. rozmieszczenie
  elementów, ostatnio użyty plik Excel, pola statyczne czy grupy.
- Automatyczne sprawdzanie dostępności nowszej wersji programu w repozytorium
  GitHub oraz możliwość pobrania aktualizacji.
- Automatyczna instalacja wymaganych pakietów przy pierwszym uruchomieniu.

## Struktura projektu
```
├── launcher.py           # uruchomienie bez zainstalowanego Pythona
├── pds_gui.py            # główny punkt startowy aplikacji (GUI)
├── pds_generator/        # logika aplikacji jako paczka Pythona
│   ├── elements.py       # definicja obiektów przeciągalnych na płótnie
│   ├── groups.py         # obsługa obszarów grupujących i edytora grup
│   ├── github_utils.py   # komunikacja z GitHubem i aktualizacje
│   ├── requirements_installer.py  # doinstalowywanie zależności
│   └── gui/
│       ├── gui.py        # klasa głównego okna PDSGeneratorGUI
│       ├── pdf_export.py # generowanie plików PDF (ReportLab)
│       ├── config_io.py  # zapis/odczyt konfiguracji użytkownika
│       └── ui_layout.py  # budowanie interfejsu w Tkinterze
├── requirements.txt     # lista wymaganych bibliotek
└── build_launcher_exe.bat # tworzenie samodzielnego `launcher.exe` (Windows)
```

## Instalacja i uruchomienie
1. **Windows – bez zainstalowanego Pythona**  
   Uruchom `python launcher.py`. Skrypt pobierze przenośną wersję Pythona,
   doinstaluje wymagane pakiety i wystartuje aplikację w trybie graficznym.
2. **System z zainstalowanym Pythonem**  
   Zainstaluj zależności z `requirements.txt` (przy pierwszym uruchomieniu
   robi to automatycznie `pds_gui.py`) i uruchom:
   ```bash
   python pds_gui.py
   ```
3. **Budowanie samodzielnego `launcher.exe` (Windows)**  
   Skrypt `build_launcher_exe.bat` pobierze instalator Pythona, utworzy
   katalog `python_runtime`, zainstaluje PyInstaller i spakuje `launcher.py` w
   pojedynczy plik wykonywalny:
   ```bash
   build_launcher_exe.bat
   ```
   Po zakończeniu w katalogu projektu pojawi się `launcher.exe` wraz z
   katalogiem `python_runtime` zawierającym wbudowany interpreter.

## Podstawowy przepływ pracy
1. Wskaż plik Excel z danymi. Wiele arkuszy traktowane jest jako osobne
   źródła kolumn.
2. Zaznacz kolumny, które mają pojawić się na stronie, lub dodaj pola
   statyczne z własnym tekstem.
3. Ustal rozmiar strony (np. A4/B5 lub parametry własne w punktach) i
   zaprojektuj układ poprzez przeciąganie elementów na płótnie. Dostępne są
   narzędzia formatowania, zmiana warstwy, przybliżanie oraz usuwanie
   elementów klawiszem `Del`.
4. Opcjonalnie twórz *grupy* zawierające zestawy pól. Dla każdej grupy możesz
   ustawić pozycje poszczególnych pól, dodatkowe style oraz warunki
   wyświetlania.
5. Zapisz konfigurację, aby przy kolejnym uruchomieniu wczytać układ i
   ostatnio użyty plik Excel.
6. Wybierz „Generuj PDF”, aby utworzyć dokumenty w katalogu `PDS` obok pliku
   Excel. Nazwy plików bazują na pierwszej kolumnie wiersza – znaki niedozwolone
   są automatycznie usuwane.
7. W komórkach Excela można podawać nazwy plików obrazów (w folderze pliku
   lub jego podfolderach) bądź pełne adresy URL; obrazy zostaną osadzone w
   odpowiednich elementach.

## Wymagane biblioteki
Pakiety instalowane automatycznie (lista w `requirements.txt`):
- `pandas`
- `Pillow`
- `reportlab`
- `requests`
- `openpyxl`

## Aktualizacje
Podczas uruchamiania aplikacja sprawdza dostępność nowszej wersji w
repozytorium GitHub. W razie wykrycia aktualizacji można ją pobrać jednym
kliknięciem; gdy brak lokalnego repozytorium Git, pobierana jest paczka ZIP
z najnowszego kodu.

## Licencja
Projekt udostępniany jest na licencji określonej w pliku `LICENSE`.

