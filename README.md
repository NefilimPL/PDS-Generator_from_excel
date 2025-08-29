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
7. Wiele elementów można zaznaczać, przeciągając prostokąt zaznaczenia na pustym obszarze canvasu; zaznaczone elementy można następnie przesuwać razem.
8. Przytrzymaj `Ctrl` i użyj kółka myszy, aby przybliżać widok w miejscu kursora (maksymalnie do 400%); gęstość siatki dopasowuje się do skali, a aktualne powiększenie w procentach widać w prawym dolnym rogu obok przycisku dopasowania strony. Widok można przesuwać, trzymając wciśnięty środkowy przycisk myszy. Klawisz `Del` usuwa zaznaczone elementy i odznacza ich checkboxy.
9. Układ dopasowuje się do rozmiaru okna, zachowując proporcje strony i ograniczając maksymalne oddalenie tak, by kartka pozostawała w polu widzenia. Przycisk **Dopasuj** centruje stronę w oknie, a obszar przewijania nie zmienia się przy zmianie powiększenia, dzięki czemu kartka nie znika za niewidzialnymi granicami. Ostatnio zapisany rozmiar strony jest wczytywany przy kolejnym uruchomieniu, a obrazki w PDF-ach są zapisywane w pełnej jakości.
10. Zaznaczone pole może mieć przezroczyste tło (przycisk **Przezroczyste**) lub kolor, dzięki czemu tekst da się nakładać na inne elementy.
11. Zapisz konfigurację (zapamiętuje ostatni plik Excel i ustawienia pól) lub wygeneruj pliki PDF dla wszystkich wierszy Excela; generator próbuje nadpisać istniejące pliki, a gdy są zablokowane, zapisuje nową wersję z inną nazwą.
12. Przyciskiem **Warunki** można zdefiniować zależności: jeżeli wskazane pole jest puste, inne pola zostaną pominięte podczas generowania.
13. **Dodaj grupę** tworzy na stronie półprzezroczyste pole z podglądem zawartości. Dwukrotne kliknięcie otwiera edytor z taką samą siatką i paskiem formatowania jak w głównym oknie; pola można dodawać z listy, przeciągać, zmieniać ich czcionkę, kolory, wyrównanie oraz definiować warunki ukrywania. Przy podglądzie oraz w PDF‑ie elementy są przesuwane w górę, ale zatrzymują się tuż pod innymi polami, dzięki czemu ich rozmiary są respektowane i nie nachodzą na siebie. Podgląd na stronie głównej pokazuje układ bloków ograniczony do obszaru grupy i nie pozwala, by dane wychodziły poza jej granice. Rozmiar grupy zmieniasz, przeciągając jej róg na głównej stronie, bez wpisywania wartości w edytorze. Warunki zdefiniowane w edytorze grupy są niezależne od warunków głównej strony i działają tylko na pola wewnątrz grupy.
    Dodanie pola do grupy nie usuwa go z głównej strony – obie wersje pojawią się w wygenerowanym PDF.
    Ustawienia pól w grupie (rozmiary, kolory, wyrównanie) są niezależne od tych samych pól na głównej stronie i zapisywane w konfiguracji. Lista dostępnych pól pokazuje również wpisany tekst pól statycznych i można ją przewijać kółkiem myszy; próba otwarcia kolejnego edytora przenosi już otwarte okno na wierzch zamiast tworzyć nowe.

14. Skrót `Ctrl+Z` cofa ostatnią zmianę na stronie, a `Ctrl+X` przywraca cofniętą modyfikację.

Wymagane biblioteki są instalowane automatycznie przy pierwszym uruchomieniu skryptu.

