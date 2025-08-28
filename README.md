# PDS-Generator_from_excel

Prosty generator PDS (PDF) na podstawie danych z Excela z opcją dodawania nowych pól dla nowych kolumn.

## Uruchomienie

```bash
python pds_gui.py
```

Po uruchomieniu aplikacji:

1. Wybierz plik Excel zawierający dane.
2. Zaznacz kolumny oraz pola statyczne, które chcesz umieścić na stronie.
3. W polu rozmiaru strony wpisz nazwę formatu (np. `A4`) albo własne wymiary w punktach w formacie `szerokośćxwysokość` i zatwierdź przyciskiem **Ustaw**.
4. Przeciągnij i zmień rozmiar elementów na białej stronie wyświetlanej na szarym tle z drobną siatką; elementy przyciągają się do siatki po puszczeniu przycisku, a przy górnej i lewej krawędzi widoczna jest linijka w punktach ułatwiająca precyzyjne pozycjonowanie.
5. Listy po prawej stronie można przewijać kółkiem myszy. Pola statyczne można dodawać w dowolnej liczbie i dla każdego wpisać własną wartość.
6. Nad polem konfiguracji znajdują się przyciski formatowania: pogrubienie, zmiana wielkości czcionki, kolory tekstu i tła, wyrównanie lewo/środek/prawo oraz wyśrodkowanie zaznaczonych elementów w pionie lub poziomie. Wartość rozmiaru czcionki można wpisać ręcznie.
7. Można zaznaczać wiele elementów jednocześnie, przytrzymując `Shift` podczas klikania; wybrane elementy można następnie przesuwać razem.
8. Przytrzymaj `Ctrl` i użyj kółka myszy, aby przybliżać widok w miejscu kursora (maksymalnie do 400%); gęstość siatki dopasowuje się do skali, a aktualne powiększenie w procentach widać w prawym dolnym rogu obok przycisku dopasowania strony. Widok można przesuwać, trzymając wciśnięty środkowy przycisk myszy. Klawisz `Del` usuwa zaznaczone elementy i odznacza ich checkboxy.
9. Układ dopasowuje się do rozmiaru okna, zachowując proporcje strony i ograniczając maksymalne oddalenie tak, by kartka pozostawała w polu widzenia. Przycisk **Dopasuj** centruje stronę w oknie i usuwa efekt "gubienia" kartki poza widokiem. Ostatnio zapisany rozmiar strony jest wczytywany przy kolejnym uruchomieniu, a obrazki w PDF-ach są zapisywane w pełnej jakości.
10. Zapisz konfigurację (zapamiętuje ostatni plik Excel i ustawienia pól) lub wygeneruj pliki PDF dla wszystkich wierszy Excela.

Wymagane biblioteki są instalowane automatycznie przy pierwszym uruchomieniu skryptu.

