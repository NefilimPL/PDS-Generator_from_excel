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
4. Przeciągnij i zmień rozmiar elementów na białej stronie z drobną siatką; elementy przyciągają się do siatki po puszczeniu przycisku.
5. Listy po prawej stronie można przewijać kółkiem myszy. Pola statyczne można dodawać w dowolnej liczbie i dla każdego wpisać własną wartość.
6. Nad polem konfiguracji znajdziesz przyciski formatowania tekstu oraz wybór kolorów tekstu i tła zaznaczonego elementu. Nowe pola statyczne są tworzone jako puste.
7. Przytrzymaj `Ctrl` i użyj kółka myszy, aby przybliżać widok w miejscu kursora; gęstość siatki dopasowuje się do skali, a aktualne powiększenie w procentach widać w prawym dolnym rogu obok przycisku dopasowania strony. Widok można przesuwać, trzymając wciśnięty środkowy przycisk myszy. Klawisz `Del` usuwa zaznaczony element i odznacza jego checkbox.
8. Układ dopasowuje się do rozmiaru okna, zachowując proporcje strony i ograniczając maksymalne oddalenie tak, by kartka zawsze pozostawała w polu widzenia. Ostatnio zapisany rozmiar strony jest wczytywany przy kolejnym uruchomieniu.
9. Zapisz konfigurację (zapamiętuje ostatni plik Excel i ustawienia pól) lub wygeneruj pliki PDF dla wszystkich wierszy Excela.

Wymagane biblioteki są instalowane automatycznie przy pierwszym uruchomieniu skryptu.

