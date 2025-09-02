# PDS-Generator_from_excel

Prosty generator PDS (PDF) na podstawie danych z Excela z opcją dodawania nowych pól dla nowych kolumn.

## Uruchomienie

```bash
python pds_gui.py
```

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

## Budowanie launchera

Aby przygotować plik `.exe` uruchamiający aplikację bez konsoli:

```bash
python build_launcher.py
```

Polecenie tworzy katalog `dist` z `launcher.exe` oraz kopiami `pds_gui.py` i modułu `pds_generator`. Uruchomiony plik sprawdza obecność Pythona, w razie potrzeby proponuje jego instalację i następnie uruchamia `pds_gui.py`.

