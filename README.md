# PDS-Generator_from_excel

Prosty generator PDS (PDF) na podstawie danych z Excela z opcją dodawania nowych pól dla nowych kolumn.

## Uruchomienie

Jeśli nie masz zainstalowanego środowiska Python, możesz skorzystać ze
skryptu `launcher.py`, który w razie potrzeby pobierze i zainstaluje
lokalną kopię Pythona (na systemie Windows) i uruchomi główną aplikację.
Przy kolejnych uruchomieniach wykorzystana zostanie już pobrana wersja.
Na Windowsie launcher uruchamia `pythonw.exe`, dzięki czemu przy starcie
nie pojawia się dodatkowa konsola:

```bash
python launcher.py
```

Gdy Python jest już zainstalowany, program można uruchomić bezpośrednio:

```bash
python pds_gui.py
```

### Budowanie samodzielnego `launcher.exe`

W repozytorium znajduje się skrypt `build_launcher_exe.bat`, który tworzy
samodzielny plik wykonywalny `launcher.exe`. Skrypt nie wymaga wcześniej
zainstalowanego Pythona – w razie potrzeby pobiera oficjalny instalator,
instaluje Pythona do katalogu `python_runtime` (bez tworzenia skrótów),
po czym instaluje PyInstaller i pakuje `launcher.py` w pojedynczy plik EXE.
Jeżeli katalog `python_runtime` już istnieje, skrypt wykorzysta
zainstalowany tam interpreter, dzięki czemu ponowne budowanie jest
znacznie szybsze.

```bash
build_launcher_exe.bat
```

Po zakończeniu w katalogu projektu pojawi się plik `launcher.exe` oraz
zainstalowany interpreter w podkatalogu `python_runtime`. Dzięki temu
`launcher.exe` może uruchomić aplikację bez dodatkowego pobierania
Pythona.

Po uruchomieniu aplikacji:

1. Wybierz plik Excel zawierający dane.
2. Zaznacz kolumny oraz pola statyczne, które chcesz umieścić na stronie.
3. Wybierz rozmiar strony lub podaj własne w formacie `szerokośćxwysokość` (w punktach).
4. Przeciągnij i zmień rozmiar elementów na polu konfiguracji z wyraźną siatką; elementy przyciągają się do siatki po puszczeniu przycisku.
5. Listy po prawej stronie można przewijać kółkiem myszy. Pola statyczne można dodawać w dowolnej liczbie i dla każdego wpisać własną wartość.
6. Nad polem konfiguracji znajdziesz przyciski formatowania tekstu oraz wybór kolorów tekstu i tła zaznaczonego elementu.
7. Suwak Zoom pozwala powiększać pole konfiguracji, a klawisz `Del` usuwa zaznaczony element i odznacza jego checkbox.
8. Układ dopasowuje się do rozmiaru okna, zachowując proporcje strony. Ostatnio zapisany rozmiar strony jest wczytywany przy kolejnym uruchomieniu.
9. Zapisz konfigurację (zapamiętuje ostatni plik Excel i ustawienia pól) lub wygeneruj pliki PDF dla wszystkich wierszy Excela.
10. W komórkach Excela możesz podać nazwę pliku obrazu (np. `czarny.jpg`); program wyszuka go w folderze pliku Excel oraz jego podfolderach.

Wymagane biblioteki są instalowane automatycznie przy pierwszym uruchomieniu skryptu.

