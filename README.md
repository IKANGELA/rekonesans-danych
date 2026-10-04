# Rekonesans danych

Szybki raport jakości pliku **CSV lub Excel** – zanim zaczniesz czyścić dane, od razu widzisz, co jest w nich nie tak.

Działa **w całości w przeglądarce**: Python uruchamia się lokalnie dzięki [Pyodide](https://pyodide.org) (WebAssembly),
a plik nie jest nigdzie wysyłany.

## Co wykrywa

Dla każdej kolumny:

| Grupa | Kontrole |
|---|---|
| **Puste** | puste komórki, wpisy udające brak danych (`NULL`, `brak`, `-`) |
| **Spacje i znaki** | spacje na brzegach i podwójne, twarde spacje, znaki niewidoczne (tabulator, enter, spacja zerowej szerokości) |
| **Kodowanie** | źle odczytane polskie znaki (`KrakÃ³w`, `zako�czone`) |
| **Formaty** | kilka formatów dat i liczb, wartości innego typu niż reszta kolumny, zapis naukowy z Excela (`5,90123E+12`), zgubione zera wiodące (`2-345` zamiast `02-345`), jednostki zapisane różnie (`kg` / `kilogram`, g i kg w jednej kolumnie), ucięte wartości |
| **Daty** | daty niemożliwe (`2025-02-31`), spoza sensownego zakresu, zapisane jako liczby Excela (`45683`), niejednoznaczne (`03/04/2025`) |
| **Warianty** | ta sama wartość zapisana różnie (`różowy` / `Różowy` / `rozowy`), literówki (`Warszwa` → `Warszawa`), tak/nie w wielu zapisach (`tak`, `T`, `1`, `TAK`) |
| **Identyfikatory** | ten sam numer w różnych wierszach, **NIP, PESEL, REGON, IBAN z kontrolą cyfry kontrolnej**, e-mail, telefon i kod pocztowy – poprawność i warianty zapisu |
| **Odstające** | nieliczne, wyraźnie oderwane od reszty duże wartości – typowo literówki z dopisanym zerem |

Dla całego pliku: kodowanie, separator kolumn i dziesiętny, zdublowane wiersze, problemy z nagłówkiem,
wiersze o złej liczbie kolumn, kolumny puste i stałe oraz **sprzeczne daty w wierszu**
(np. data wysyłki wcześniejsza niż data zamówienia).

Synonimów (`różowy` / `pink`) program nie rozpozna sam – dla kolumn z małą liczbą różnych wartości
pokazuje pełną, alfabetyczną listę, w której łatwo je wyłapać.

## Jak to działa

- `rekonesans.py` – rdzeń w czystym Pythonie (bez pandas). Każda wartość jest badana **jako tekst,
  dokładnie tak, jak jest zapisana w pliku** – dzięki temu widać prawdziwe formaty i spacje.
  Moduł nie korzysta z dysku ani sieci, więc działa w przeglądarce i w terminalu.
- `app.js`, `index.html` – interfejs: wybór pliku, raport, szczegóły kolumn.
- Warianty grupowane są po ujednoliceniu (małe litery, bez polskich znaków i nadmiarowych spacji).
- Wartości odstające liczone są w skali logarytmicznej i względem mediany – dobrze łapią literówki
  w ilościach, które w firmach są zwykle mocno zróżnicowane.

## Prywatność i bezpieczeństwo

Plik jest analizowany tylko w pamięci przeglądarki · Pyodide i biblioteki są dołączone do repozytorium
(bez zewnętrznych CDN) · polityka Content-Security-Policy blokuje połączenia z innymi adresami ·
treść pliku wyświetlana jest wyłącznie jako tekst.

## Uruchomienie lokalne

Dwuklik w `podglad.bat` (wymaga Node.js) i otwarcie http://127.0.0.1:8765.
Przykładowy plik z celowymi błędami: `przyklady/przyklad_zamowienia.csv` (dane fikcyjne,
generuje go `narzedzia/zrob_przyklad.py`).

## Oczyszczanie i zapis (wersja 2)

Przy każdym wykrytym problemie jest pole **„popraw”** – zaznaczone domyślnie tam, gdzie poprawka jest bezpieczna.
Przycisk **„Wyczyść i zapisz w formacie:”** zapisuje wynik jako **Excel (.xlsx)** albo **CSV** (średnik lub przecinek).

Poprawki: zdublowane wiersze, puste i stałe kolumny, spacje i znaki niewidoczne, polskie znaki (`KrakÃ³w` → `Kraków`),
`NULL` / `brak` → puste, warianty zapisu (wybierany jest zapis z polskimi znakami, nie WERSALIKAMI), literówki
(domyślnie wyłączone), tak/nie, daty w wybranym formacie (także z liczb Excela), liczby, zgubione zera,
masa w kilogramach, telefony, NIP/PESEL/REGON, IBAN, e-maile, kody pocztowe.

Każda zmiana trafia do **dziennika zmian** (kolumna, operacja, liczba komórek, przykłady „przed → po”).
Czego nie da się bezpiecznie poprawić – literówki, wartości odstające, niemożliwe daty, błędne numery,
powtórzone identyfikatory, sprzeczne daty, zapis naukowy, ucięte wartości – trafia na listę
**„do sprawdzenia”** z numerami wierszy pliku źródłowego.

Excel zawiera arkusze: *Dane* (prawdziwe daty i liczby), *Dziennik zmian*, *Do sprawdzenia*.
Przy CSV dziennik pobiera się osobnym przyciskiem.

## Plany

Własny słownik zamian (np. `pink` → `różowy`, `grey` → `szary`) – synonimów program nie rozpozna sam.

## Autorka

Angelika Berecka · [ikangela.pl](https://ikangela.pl) · [GitHub](https://github.com/IKANGELA)

Licencja: MIT (dołączone biblioteki zachowują własne licencje – zob. `LICENSE`).
