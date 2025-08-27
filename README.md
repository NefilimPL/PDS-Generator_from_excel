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
6. Użyj przycisków formatowania tekstu nad polem konfiguracji, aby zmienić pogrubienie lub wpisać dokładny rozmiar czcionki zaznaczonego elementu; zmiana rozmiaru elementu nie nadpisuje ręcznie ustawionego rozmiaru czcionki.
7. Układ dopasowuje się do rozmiaru okna, zachowując proporcje strony. Ostatnio zapisany rozmiar strony jest wczytywany przy kolejnym uruchomieniu.
8. Zapisz konfigurację (zapamiętuje ostatni plik Excel i ustawienia pól) lub wygeneruj pliki PDF dla wszystkich wierszy Excela.

Wymagane biblioteki są instalowane automatycznie przy pierwszym uruchomieniu skryptu.

