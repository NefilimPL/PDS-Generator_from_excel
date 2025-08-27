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
4. Przeciągnij i zmień rozmiar elementów na polu konfiguracji z widoczną siatką; elementy przyciągają się do siatki.
5. Użyj przycisków formatowania tekstu nad polem konfiguracji, aby zmienić pogrubienie lub wpisać dokładny rozmiar czcionki zaznaczonego elementu.
6. Układ dopasowuje się do rozmiaru okna, zachowując proporcje strony.
7. Zapisz konfigurację (zapamiętuje ostatni plik Excel i ustawienia pól) lub wygeneruj pliki PDF dla wszystkich wierszy Excela.

Wymagane biblioteki są instalowane automatycznie przy pierwszym uruchomieniu skryptu.
