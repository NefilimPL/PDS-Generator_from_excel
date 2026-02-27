# PDS-Generator_from_excel
<img width="1920" height="1040" alt="image" src="https://github.com/user-attachments/assets/252bb043-014d-47ed-bdc0-853664916e82" />

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
| Zapisywanie i wczytywanie konfiguracji do pliku `config.json` w katalogu wybranego pliku Excel (tworzona jest także kopia zapasowa w `%APPDATA%\\PDS Generator\\config.json`) – zapamiętywane są m.in. rozmieszczenie elementów, ostatnio użyty plik Excel, pola statyczne czy grupy. | Saving and loading configuration to a `config.json` file in the selected Excel file's directory (a backup is also created in `%APPDATA%\\PDS Generator\\config.json`) – stores element layout, last used Excel file, static fields and groups. |
| Automatyczne sprawdzanie dostępności nowszej wersji programu w repozytorium GitHub oraz możliwość pobrania aktualizacji. | Automatic check for newer versions in the GitHub repository and option to download updates. |
| Automatyczna instalacja wymaganych pakietów przy pierwszym uruchomieniu. | Automatic installation of required packages on first run. |



## Instalacja i uruchomienie / Installation and Run

| Polski | English |
|---|---|
| 1. **Windows – bez zainstalowanego Pythona**<br>Uruchom `python launcher.py`. Skrypt pobierze przenośną wersję Pythona, doinstaluje wymagane pakiety i wystartuje aplikację w trybie graficznym.<br><br>2. **System z zainstalowanym Pythonem**<br>Zainstaluj zależności z `requirements.txt` (przy pierwszym uruchomieniu robi to automatycznie `pds_gui.py`) i uruchom:<br>`python pds_gui.py`<br><br>3. **Budowanie samodzielnego `launcher.exe` (Windows)**<br>Skrypt `build_launcher_exe.bat` pobierze instalator Pythona, utworzy katalog `python_runtime`, zainstaluje PyInstaller i spakuje `launcher.py` w pojedynczy plik wykonywalny:<br>`build_launcher_exe.bat`<br>Po zakończeniu w katalogu projektu pojawi się `launcher.exe` wraz z katalogiem `python_runtime` zawierającym wbudowany interpreter. | 1. **Windows – without Python installed**<br>Run `python launcher.py`. The script downloads a portable Python, installs required packages, and starts the GUI application.<br><br>2. **System with Python installed**<br>Install dependencies from `requirements.txt` (on first run this is done automatically by `pds_gui.py`) and run:<br>`python pds_gui.py`<br><br>3. **Building standalone `launcher.exe` (Windows)**<br>The `build_launcher_exe.bat` script downloads the Python installer, creates the `python_runtime` directory, installs PyInstaller, and packages `launcher.py` into a single executable:<br>`build_launcher_exe.bat`<br>After completion, `launcher.exe` appears in the project directory along with `python_runtime` containing the embedded interpreter. |

## Podstawowy przepływ pracy / Basic Workflow

| Polski | English |
|---|---|
| 1. Wskaż plik Excel z danymi. Wiele arkuszy traktowane jest jako osobne źródła kolumn.<br>2. Zaznacz kolumny, które mają pojawić się na stronie, lub dodaj pola statyczne z własnym tekstem.<br>3. Ustal rozmiar strony (np. A4/B5 lub parametry własne w punktach) i zaprojektuj układ poprzez przeciąganie elementów na płótnie. Dostępne są narzędzia formatowania, zmiana warstwy, przybliżanie oraz usuwanie elementów klawiszem `Del`.<br>4. Opcjonalnie twórz *grupy* zawierające zestawy pól. Dla każdej grupy możesz ustawić pozycje poszczególnych pól, dodatkowe style oraz warunki wyświetlania.<br>5. Zapisz konfigurację, aby przy kolejnym uruchomieniu wczytać układ i ostatnio użyty plik Excel.<br>6. Wybierz „Generuj PDF”, aby utworzyć dokumenty w katalogu `PDS` obok pliku Excel. Nazwy plików bazują na pierwszej kolumnie wiersza – znaki niedozwolone są automatycznie usuwane.<br>7. W komórkach Excela można podawać nazwy plików obrazów (w folderze pliku lub jego podfolderach) bądź pełne adresy URL; obrazy zostaną osadzone w odpowiednich elementach. | 1. Select the Excel file with data. Multiple sheets are treated as separate column sources.<br>2. Mark the columns that should appear on the page or add static fields with custom text.<br>3. Set the page size (e.g., A4/B5 or custom in points) and design the layout by dragging elements on the canvas. Formatting tools, layer changes, zooming, and deleting elements with the `Del` key are available.<br>4. Optionally create *groups* containing sets of fields. For each group, you can set positions of individual fields, additional styles, and display conditions.<br>5. Save the configuration to reload the layout and last used Excel file next time.<br>6. Choose "Generate PDF" to create documents in the `PDS` directory next to the Excel file. File names are based on the row's first column—invalid characters are automatically removed.<br>7. Excel cells can include image file names (in the file's folder or subfolders) or full URLs; images are embedded in the respective elements. |

## Bezpieczeństwo sekretów / Secrets Security

| Polski | English |
|---|---|
| Podczas zapisu konfiguracji pola wrażliwe (`Hasło SMTP`, `Token Bearer`, `Secret Value`) nie są trzymane jawnie. W systemie Windows zapisywane są w `config.json` jako wartości zaszyfrowane przez DPAPI. | Sensitive fields (`SMTP password`, `Bearer token`, `Secret Value`) are not stored in plaintext. On Windows they are saved in `config.json` as DPAPI-encrypted values. |
| W ustawieniach e-mail możesz kliknąć `Generuj certyfikat` (dla Tenant ID + Client ID). Aplikacja utworzy plik klucza w AppData użytkownika (`.../PDS Generator/secrets/*.key`) i użyje go do szyfrowania sekretów w `config.json`. | In e-mail settings you can click `Generate certificate` (for Tenant ID + Client ID). The app creates a key file in user AppData (`.../PDS Generator/secrets/*.key`) and uses it to encrypt secrets in `config.json`. |
| Każdy komputer z kopią tego pliku klucza w swoim AppData może odszyfrować konfigurację. | Any machine with a copy of that key file in its AppData can decrypt the configuration. |
| W oknie konfiguracji e-mail pola wrażliwe są domyślnie maskowane. | Sensitive fields in e-mail settings are masked by default. |
| Dostęp do okna konfiguracji e-mail można wymusić przez logowanie kontem administratora Windows (natywny monit systemowy). Wymaganie jest domyślnie włączone; można je wyłączyć przez `PDS_REQUIRE_ADMIN_MAIL_SETTINGS=0`. | Access to e-mail settings can be gated by Windows administrator sign-in (native system credential prompt). The requirement is enabled by default; disable with `PDS_REQUIRE_ADMIN_MAIL_SETTINGS=0`. |
| Do uruchomień automatycznych/headless ustaw sekrety jako zmienne środowiskowe: `PDS_SMTP_PASSWORD`, `PDS_ENTRA_TOKEN`, `PDS_ENTRA_CLIENT_SECRET`. | For scheduled/headless runs, provide secrets via environment variables: `PDS_SMTP_PASSWORD`, `PDS_ENTRA_TOKEN`, `PDS_ENTRA_CLIENT_SECRET`. |
| Odszyfrowanie DPAPI działa tylko dla tego samego konta Windows, które zapisało konfigurację. | DPAPI decryption works only for the same Windows account that saved the configuration. |
| Opcjonalnie możesz wymusić konkretny plik klucza przez `PDS_SECRET_KEY_FILE` lub użyć wspólnego klucza ENV `PDS_SHARED_SECRET_KEY` (fallback). | Optionally you can force a specific key file via `PDS_SECRET_KEY_FILE` or use shared ENV key `PDS_SHARED_SECRET_KEY` (fallback). |
| Jeśli wartość jest wpisana w GUI, ma pierwszeństwo; zmienne środowiskowe są używane jako fallback. | Values entered in the GUI take precedence; environment variables are used as fallback. |

## Microsoft Entra API - Tokeny i uprawnienia (PL)

### 1) Wymagane uprawnienia aplikacji w Entra ID
1. Wejdź do **Microsoft Entra admin center** -> **App registrations** -> Twoja aplikacja.
2. W **API permissions** dodaj:
   `Microsoft Graph` -> `Application permissions` -> `Mail.Send`.
3. Kliknij **Grant admin consent** dla tenantu.
4. Bez `Mail.Send` w trybie aplikacyjnym (`client_credentials`) wysyłka przez Graph nie zadziała.

### 2) Wymagane dane w GUI (tryb Entra API)
1. `Tenant ID` = identyfikator dzierżawy.
2. `Client ID` = identyfikator aplikacji.
3. `Secret Value` = **wartość** sekretu klienta (nie Secret ID).
4. `Nadawca (UPN/ID)` = skrzynka, z której Graph ma wysyłać (`/users/{sender}/sendMail`).

### 3) Token - jak działa w aplikacji
1. Aplikacja pobiera token przez endpoint:
   `https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token`
2. Zakres dla `client_credentials`:
   `scope=https://graph.microsoft.com/.default`
3. Możesz też wkleić gotowy `Token Bearer` ręcznie.
4. Dla tokenu ręcznego sprawdź:
   `aud = https://graph.microsoft.com`
   oraz obecność `Mail.Send` (`roles` dla app token, `scp` dla delegated token).

### 4) Ograniczenie dostępu aplikacji do wybranych skrzynek (zalecane)
1. Samo `Mail.Send` (application) domyślnie daje szeroki dostęp.
2. Aby zawęzić zakres skrzynek, skonfiguruj po stronie Exchange Online:
   **RBAC for Applications** (zalecane, nowsze)
   lub **Application Access Policies** (legacy).

### 5) Certyfikat/klucz szyfrowania konfiguracji - gdzie ma trafić
1. Po kliknięciu `Generuj certyfikat` plik klucza trafia do:
   `%APPDATA%\\PDS Generator\\secrets\\<secret_key_id>.key`
   (zwykle: `C:\\Users\\<uzytkownik>\\AppData\\Roaming\\PDS Generator\\secrets\\<secret_key_id>.key`).
2. Ten sam plik `.key` musi być skopiowany na każdy komputer, który ma odszyfrowywać sekrety z `config.json`.
3. Jeżeli klucz ma być poza AppData, ustaw:
   `PDS_SECRET_KEY_FILE=C:\\sciezka\\klucz.key`
4. Jeśli na innym PC utworzysz **nowy** klucz dla tej samej konfiguracji i zapiszesz config, stare `*_enc` mogą przestać się odszyfrowywać na maszynach ze starym kluczem.

### 6) Headless / harmonogram
1. Najbezpieczniej ustawić sekrety jako ENV na serwerze:
   `PDS_SMTP_PASSWORD`, `PDS_ENTRA_TOKEN`, `PDS_ENTRA_CLIENT_SECRET`.
2. Wtedy wysyłka nie zależy od GUI i ręcznego wpisywania danych.

### 7) Oficjalne źródła Microsoft (aktualne)
1. Graph permissions reference (`Mail.Send`):  
   https://learn.microsoft.com/en-us/graph/permissions-reference
2. Client credentials i `/.default`:  
   https://learn.microsoft.com/en-us/entra/identity-platform/scenario-daemon-acquire-token  
   https://learn.microsoft.com/en-us/azure/active-directory/develop/scopes-oidc
3. Endpoint `user: sendMail` (v1.0):  
   https://learn.microsoft.com/en-us/graph/api/user-sendmail?view=graph-rest-1.0
4. Ograniczanie zakresu skrzynek w Exchange Online:  
   RBAC for Applications: https://learn.microsoft.com/en-us/exchange/permissions-exo/application-rbac  
   Application Access Policies (legacy): https://learn.microsoft.com/en-us/exchange/permissions-exo/application-access-policies

## Microsoft Entra API - Tokens and permissions (EN)
1. Required Graph app permission for client credentials: `Mail.Send` (Application) + admin consent.
2. Use `Tenant ID`, `Client ID`, `Secret Value` (secret value, not secret ID), and mailbox sender (`UPN/ID`).
3. Token scope for client credentials: `https://graph.microsoft.com/.default`.
4. If using a pasted bearer token, verify Graph audience and `Mail.Send` claim (`roles` or `scp`).
5. Generated encryption key file location:
   `%APPDATA%\\PDS Generator\\secrets\\<secret_key_id>.key`
   (or override with `PDS_SECRET_KEY_FILE`).
6. To decrypt the same config on multiple machines, copy the same `.key` file to each machine.
7. For mailbox scoping, configure Exchange Online **RBAC for Applications** (recommended) or legacy Application Access Policies.

## Harmonogram zadań (Windows) / Task Scheduler (Windows)

| Polski | English |
|---|---|
| 1. Upewnij się, że masz zapisany układ (`config.json`).<br>2. Uruchom ręcznie `pds_headless.bat` (opcjonalnie z `--excel "C:\\sciezka\\plik.xlsx"` lub `--config "C:\\sciezka\\config.json"`), aby sprawdzić logi w `logs`.<br>3. Otwórz **Harmonogram zadań** → **Utwórz zadanie...**.<br>4. Zakładka **Akcje** → **Nowa**: **Program/skrypt** = pełna ścieżka do `pds_headless.bat`.<br>5. (Opcjonalnie) **Dodaj argumenty**: `--excel "C:\\sciezka\\plik.xlsx"` lub `--config "C:\\sciezka\\config.json"`.<br>6. **Rozpocznij w**: katalog projektu (np. `C:\\_GitHub_\\PDS-Generator_from_excel`).<br>7. Ustaw wyzwalacz (godzina/dni) i zapisz zadanie. | 1. Make sure you saved the layout (`config.json`).<br>2. Run `pds_headless.bat` once (optionally with `--excel "C:\\path\\file.xlsx"` or `--config "C:\\path\\config.json"`) and check logs in `logs`.<br>3. Open **Task Scheduler** → **Create Task...**.<br>4. **Actions** tab → **New**: **Program/script** = full path to `pds_headless.bat`.<br>5. (Optional) **Add arguments**: `--excel "C:\\path\\file.xlsx"` or `--config "C:\\path\\config.json"`.<br>6. **Start in**: project directory (e.g. `C:\\_GitHub_\\PDS-Generator_from_excel`).<br>7. Set the trigger (time/days) and save the task. |

## Wymagane biblioteki / Required Libraries

| Polski | English |
|---|---|
| Pakiety instalowane automatycznie (lista w `requirements.txt`):<br>• `pandas`<br>• `Pillow`<br>• `reportlab`<br>• `requests`<br>• `openpyxl`<br>• `cryptography` | Packages installed automatically (listed in `requirements.txt`):<br>• `pandas`<br>• `Pillow`<br>• `reportlab`<br>• `requests`<br>• `openpyxl`<br>• `cryptography` |

## Aktualizacje / Updates

| Polski | English |
|---|---|
| Podczas uruchamiania aplikacja sprawdza dostępność nowszej wersji w repozytorium GitHub. W razie wykrycia aktualizacji można ją pobrać jednym kliknięciem; gdy brak lokalnego repozytorium Git, pobierana jest paczka ZIP z najnowszego kodu. | On startup, the application checks for a newer version in the GitHub repository. If an update is available, it can be downloaded with one click; if no local Git repository is present, a ZIP archive of the latest code is downloaded. |

## Licencja / License

| Polski | English |
|---|---|
| Projekt udostępniany jest na licencji określonej w pliku `LICENSE`. | The project is distributed under the license specified in the `LICENSE` file. |
