# PDS-Generator_from_excel

## Opis / Description

| Polski | English |
|---|---|
| Rozbudowany generator dokumentów **PDS** (PDF) oparty na danych z arkuszy Excel. | Advanced **PDS** (PDF) document generator based on data from Excel sheets. |
| Aplikacja udostępnia graficzny edytor umożliwiający projektowanie układu strony poprzez przeciąganie pól tekstowych, obrazów oraz elementów grupowych. | The application provides a graphical editor that lets you design the page layout by dragging text fields, images, and group elements. |
| Skonfigurowany projekt może zostać wykorzystany do seryjnego tworzenia plików PDF – dla każdego wiersza arkusza powstaje oddzielny dokument. | A configured project can be used to generate PDF files in batches—creating a separate document for each row in the sheet. |

## Funkcjonalności / Features

| Polski | English |
|---|---|
| Wczytywanie wielu arkuszy Excela (`.xlsx`) i przypisywanie kolumn do pól na stronie. | Loading multiple Excel sheets (`.xlsx`) and assigning columns to fields on the page. |
| Pola statyczne z własną treścią i możliwością dowolnego formatowania. | Static fields with custom content and flexible formatting. |
| Przeciąganie i skalowanie elementów na siatce z przyciąganiem do kroków (domyślnie co 5 pt) oraz podglądem linii wyrównania. | Dragging and scaling elements on a grid with snapping (default every 5 pt) and alignment guide preview. |
| Edycja kroju, rozmiaru i stylu czcionki, kolorów tła i tekstu, wyrównania oraz warstwy (kolejności rysowania) każdego elementu. | Editing font family, size and style, background and text colors, alignment, and layer (drawing order) of each element. |
| Obsługa obrazów lokalnych lub zdalnych (URL podany w komórce Excela). | Support for local or remote images (URL provided in the Excel cell). |
| Grupowanie pól w *obszary* z własnym podglądem i indywidualną konfiguracją, w tym warunkowym ukrywaniem elementów zależnie od zawartości innych pól. | Grouping fields into *areas* with their own preview and configuration, including conditional hiding based on other fields' content. |
| Zapisywanie i wczytywanie konfiguracji do pliku `config.json` w katalogu wybranego pliku Excel (tworzona jest także kopia zapasowa w `~/.pds_generator/config.json`) – zapamiętywane są m.in. rozmieszczenie elementów, ostatnio użyty plik Excel, pola statyczne czy grupy. | Saving and loading configuration to a `config.json` file in the selected Excel file's directory (a backup is also created in `~/.pds_generator/config.json`) – stores element layout, last used Excel file, static fields and groups. |
| Automatyczne sprawdzanie dostępności nowszej wersji programu w repozytorium GitHub oraz możliwość pobrania aktualizacji. | Automatic check for newer versions in the GitHub repository and option to download updates. |
| Automatyczna instalacja wymaganych pakietów przy pierwszym uruchomieniu. | Automatic installation of required packages on first run. |

## Struktura projektu / Project Structure

| Polski | English |
|---|---|
| <pre>├── launcher.py           # uruchomienie bez zainstalowanego Pythona
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
└── build_launcher_exe.bat # tworzenie samodzielnego `launcher.exe` (Windows)</pre> | <pre>├── launcher.py           # launch without installed Python
├── pds_gui.py            # main application entry point (GUI)
├── pds_generator/        # application logic as a Python package
│   ├── elements.py       # definition of draggable objects on the canvas
│   ├── groups.py         # handling group areas and group editor
│   ├── github_utils.py   # communication with GitHub and updates
│   ├── requirements_installer.py  # installing dependencies
│   └── gui/
│       ├── gui.py        # PDSGeneratorGUI main window class
│       ├── pdf_export.py # generating PDF files (ReportLab)
│       ├── config_io.py  # saving/loading user configuration
│       └── ui_layout.py  # building the interface in Tkinter
├── requirements.txt     # list of required libraries
└── build_launcher_exe.bat # creating standalone `launcher.exe` (Windows)</pre> |

## Instalacja i uruchomienie / Installation and Run

| Polski | English |
|---|---|
| 1. **Windows – bez zainstalowanego Pythona**<br>Uruchom `python launcher.py`. Skrypt pobierze przenośną wersję Pythona, doinstaluje wymagane pakiety i wystartuje aplikację w trybie graficznym.<br><br>2. **System z zainstalowanym Pythonem**<br>Zainstaluj zależności z `requirements.txt` (przy pierwszym uruchomieniu robi to automatycznie `pds_gui.py`) i uruchom:<br>`python pds_gui.py`<br><br>3. **Budowanie samodzielnego `launcher.exe` (Windows)**<br>Skrypt `build_launcher_exe.bat` pobierze instalator Pythona, utworzy katalog `python_runtime`, zainstaluje PyInstaller i spakuje `launcher.py` w pojedynczy plik wykonywalny:<br>`build_launcher_exe.bat`<br>Po zakończeniu w katalogu projektu pojawi się `launcher.exe` wraz z katalogiem `python_runtime` zawierającym wbudowany interpreter. | 1. **Windows – without Python installed**<br>Run `python launcher.py`. The script downloads a portable Python, installs required packages, and starts the GUI application.<br><br>2. **System with Python installed**<br>Install dependencies from `requirements.txt` (on first run this is done automatically by `pds_gui.py`) and run:<br>`python pds_gui.py`<br><br>3. **Building standalone `launcher.exe` (Windows)**<br>The `build_launcher_exe.bat` script downloads the Python installer, creates the `python_runtime` directory, installs PyInstaller, and packages `launcher.py` into a single executable:<br>`build_launcher_exe.bat`<br>After completion, `launcher.exe` appears in the project directory along with `python_runtime` containing the embedded interpreter. |

## Podstawowy przepływ pracy / Basic Workflow

| Polski | English |
|---|---|
| 1. Wskaż plik Excel z danymi. Wiele arkuszy traktowane jest jako osobne źródła kolumn.<br>2. Zaznacz kolumny, które mają pojawić się na stronie, lub dodaj pola statyczne z własnym tekstem.<br>3. Ustal rozmiar strony (np. A4/B5 lub parametry własne w punktach) i zaprojektuj układ poprzez przeciąganie elementów na płótnie. Dostępne są narzędzia formatowania, zmiana warstwy, przybliżanie oraz usuwanie elementów klawiszem `Del`.<br>4. Opcjonalnie twórz *grupy* zawierające zestawy pól. Dla każdej grupy możesz ustawić pozycje poszczególnych pól, dodatkowe style oraz warunki wyświetlania.<br>5. Zapisz konfigurację, aby przy kolejnym uruchomieniu wczytać układ i ostatnio użyty plik Excel.<br>6. Wybierz „Generuj PDF”, aby utworzyć dokumenty w katalogu `PDS` obok pliku Excel. Nazwy plików bazują na pierwszej kolumnie wiersza – znaki niedozwolone są automatycznie usuwane.<br>7. W komórkach Excela można podawać nazwy plików obrazów (w folderze pliku lub jego podfolderach) bądź pełne adresy URL; obrazy zostaną osadzone w odpowiednich elementach. | 1. Select the Excel file with data. Multiple sheets are treated as separate column sources.<br>2. Mark the columns that should appear on the page or add static fields with custom text.<br>3. Set the page size (e.g., A4/B5 or custom in points) and design the layout by dragging elements on the canvas. Formatting tools, layer changes, zooming, and deleting elements with the `Del` key are available.<br>4. Optionally create *groups* containing sets of fields. For each group, you can set positions of individual fields, additional styles, and display conditions.<br>5. Save the configuration to reload the layout and last used Excel file next time.<br>6. Choose "Generate PDF" to create documents in the `PDS` directory next to the Excel file. File names are based on the row's first column—invalid characters are automatically removed.<br>7. Excel cells can include image file names (in the file's folder or subfolders) or full URLs; images are embedded in the respective elements. |

## Wymagane biblioteki / Required Libraries

| Polski | English |
|---|---|
| Pakiety instalowane automatycznie (lista w `requirements.txt`):<br>• `pandas`<br>• `Pillow`<br>• `reportlab`<br>• `requests`<br>• `openpyxl` | Packages installed automatically (listed in `requirements.txt`):<br>• `pandas`<br>• `Pillow`<br>• `reportlab`<br>• `requests`<br>• `openpyxl` |

## Aktualizacje / Updates

| Polski | English |
|---|---|
| Podczas uruchamiania aplikacja sprawdza dostępność nowszej wersji w repozytorium GitHub. W razie wykrycia aktualizacji można ją pobrać jednym kliknięciem; gdy brak lokalnego repozytorium Git, pobierana jest paczka ZIP z najnowszego kodu. | On startup, the application checks for a newer version in the GitHub repository. If an update is available, it can be downloaded with one click; if no local Git repository is present, a ZIP archive of the latest code is downloaded. |

## Licencja / License

| Polski | English |
|---|---|
| Projekt udostępniany jest na licencji określonej w pliku `LICENSE`. | The project is distributed under the license specified in the `LICENSE` file. |
