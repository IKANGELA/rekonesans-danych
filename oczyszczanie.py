"""
Oczyszczanie danych i eksport do Excela (.xlsx) lub CSV.

Korzysta z tych samych reguł co rekonesans.py. Każda zmiana trafia do dziennika:
co zmieniono, w której kolumnie, ile komórek i przykłady "przed → po".
Czego nie da się bezpiecznie poprawić automatycznie, trafia na listę "do sprawdzenia"
z numerami wierszy (wiersz 2 = pierwszy wiersz danych, jak w Excelu).
"""

import csv
import io
import re
from collections import Counter, defaultdict
from datetime import date, datetime, timedelta

import rekonesans as rk

MAKS_PRZYKLADOW = 5

FORMATY_DATY = {"RRRR-MM-DD": "%Y-%m-%d", "DD.MM.RRRR": "%d.%m.%Y", "DD/MM/RRRR": "%d/%m/%Y",
                "DD-MM-RRRR": "%d-%m-%Y"}
TAK_NIE = {"tak/nie": ("tak", "nie"), "1/0": ("1", "0"), "TRUE/FALSE": ("TRUE", "FALSE"), "T/N": ("T", "N")}
PRAWDA = {"tak", "t", "y", "yes", "true", "prawda", "1", "x", "✓", "✔"}
JEDNOSTKA_KROTKA = {"kilogram": "kg", "gram": "g", "sztuka": "szt", "metr kwadratowy": "m2",
                    "metr bieżący": "mb", "metr sześcienny": "m3", "litr": "l", "opakowanie": "op", "para": "para"}


# ---------------------------------------------------------------------------
# Sugestie: jakie poprawki proponujemy dla kolumny (na podstawie raportu rekonesansu)
# ---------------------------------------------------------------------------

def sugestie(k):
    """Lista poprawek do wyboru dla kolumny z raportu: [{id, opis, domyslnie}]."""
    p, s = k["problemy"], []

    def dodaj(id_, opis, domyslnie=True):
        s.append({"id": id_, "opis": opis, "domyslnie": domyslnie})

    if p["niewidoczne znaki"]:
        dodaj("niewidoczne", f"usuń znaki niewidoczne ({p['niewidoczne znaki']})")
    if p["spacje"]:
        dodaj("spacje", f"usuń zbędne spacje ({p['spacje']})")
    if p["kodowanie"]:
        dodaj("kodowanie", f"napraw polskie znaki ({p['kodowanie']})")
    if p["ukryte puste"]:
        dodaj("ukryte_puste", f"zamień NULL / brak / - na puste ({p['ukryte puste']})")
    if p["warianty"]:
        dodaj("warianty", f"ujednolić zapis do najczęstszego wariantu ({p['warianty']} grup)")
    if p["literówki"]:
        dodaj("literowki", "popraw możliwe literówki: " +
              ", ".join(f"{x['rzadki']} → {x['czesty']}" for x in k["literowki"][:4]), domyslnie=False)
    if k["logiczne"]:
        dodaj("logiczne", "ujednolić tak / nie")
    if k["typ"] == "data" or p["daty Excela"]:
        dodaj("daty", "zamień daty na wybrany format" +
              (f" (w tym {p['daty Excela']} dat z Excela)" if p["daty Excela"] else ""))
    numer = k["numer"]["rodzaj"] if k["numer"] else None
    if k["typ"] == "liczba" and not numer and not k["identyfikator"] and not p["zapis naukowy"] \
            and not _wyglada_na_kod(k["nazwa"]):
        dodaj("liczby", "zamień na liczby (przecinek → kropka, bez waluty i separatorów tysięcy)")
    if p["zgubione zera"]:
        dodaj("zera", f"uzupełnij zgubione zera na początku ({p['zgubione zera']})")
    if k["jednostki"]:
        if "kilogram" in k["jednostki"].get("różne jednostki w kolumnie", {}) or "kilogram" in k["jednostki"]:
            dodaj("masa_kg", "przelicz masę na liczby w kilogramach (g → kg)")
        else:
            dodaj("jednostki", "ujednolić zapis jednostek")
    if numer == "telefon":
        dodaj("telefon", "ujednolić zapis telefonów")
    if numer in ("NIP", "PESEL", "REGON"):
        dodaj("cyfry", f"{numer}: zostaw same cyfry (bez myślników i spacji)")
    if numer == "IBAN":
        dodaj("iban", "IBAN: wielkie litery, bez spacji")
    if numer == "e-mail":
        dodaj("email", "e-mail: małe litery, bez spacji na brzegach")
    if numer == "kod pocztowy":
        dodaj("kod_pocztowy", "kod pocztowy w formacie 00-000")
    return s


def _wyglada_na_kod(nazwa):
    return re.search(r"(^|[_\s])(nr|numer|id|kod|symbol|sku|ean|gtin|nip|pesel|regon|telefon|tel)($|[_\s])",
                     nazwa.lower()) is not None


# ---------------------------------------------------------------------------
# Pojedyncze poprawki
# ---------------------------------------------------------------------------

POLSKIE = re.compile(r"[ąćęłńóśźżĄĆĘŁŃÓŚŹŻ]")


def wybierz_wzorzec(zapisy):
    """
    Który zapis wariantu zostawić? Najlepiej: z polskimi znakami (jeśli któryś je ma),
    nie WERSALIKAMI, bez spacji na brzegach – a dopiero potem najczęstszy.
    'wysłane' ×69, 'Wysłane' ×62, 'WYSLANE' ×76  ->  'wysłane'.
    """
    ma_polskie = any(POLSKIE.search(z) for z in zapisy)

    def ocena(z):
        return (bool(POLSKIE.search(z)) or not ma_polskie, not z.isupper(), z == z.strip(), zapisy[z])
    return max(zapisy, key=ocena).strip()

def napraw_kodowanie(w):
    """'KrakÃ³w' -> 'Kraków'. Gdy nie da się odwrócić (np. utracony znak), zwraca None."""
    for kodowanie in ("cp1252", "cp1250", "latin-1"):
        try:
            naprawione = w.encode(kodowanie).decode("utf-8")
            if not rk.KRZAKI.search(naprawione):
                return naprawione
        except (UnicodeEncodeError, UnicodeDecodeError):
            continue
    return None


def jako_data_lub_excel(w):
    d, fmt = rk.jako_data(w)
    if isinstance(d, date):
        return d, fmt
    if d is None:
        return None, "niemożliwa"
    s = w.strip().replace(",", ".")
    if re.fullmatch(r"\d{5}(\.\d+)?", s) and 20000 <= float(s) <= 60000:
        return date(1899, 12, 30) + timedelta(days=int(float(s))), "liczba Excela"
    return "brak", None


def jako_liczba_ogolnie(w):
    s = w.strip().replace(" ", " ")
    s = re.sub(r"\s?(zł|pln|eur|€|\$|usd|%)$", "", s, flags=re.I).strip()
    if re.fullmatch(r"[+-]?\d{1,3}([ .]\d{3})+(,\d+)?", s):
        s = s.replace(" ", "").replace(".", "").replace(",", ".")
    elif re.fullmatch(r"[+-]?\d{1,3}(,\d{3})+(\.\d+)?", s):
        s = s.replace(",", "")
    else:
        s = s.replace(" ", "").replace(",", ".")
    try:
        x = float(s)
    except ValueError:
        return None
    return int(x) if x == int(x) and "." not in s and "e" not in s.lower() else x


def masa_w_kg(w):
    m = re.fullmatch(r"\s*([\d\s.,]+)\s*([a-ząćęłńóśźż.]+)\s*", w.lower())
    if not m:
        return None
    jedn = rk.JEDNOSTKA_DLA.get(m.group(2))
    x = jako_liczba_ogolnie(m.group(1))
    if x is None or jedn not in ("kilogram", "gram"):
        return None
    return round(x / 1000, 6) if jedn == "gram" else x


# ---------------------------------------------------------------------------
# Oczyszczanie
# ---------------------------------------------------------------------------

def oczysc(dane, nazwa_pliku, ustawienia):
    """Zwraca słownik: plik wynikowy (bajty), nazwa, dziennik zmian, lista do sprawdzenia, podsumowanie."""
    raport = rk.analizuj(dane, nazwa_pliku)
    if nazwa_pliku.lower().endswith((".xlsx", ".xlsm")) or dane[:2] == b"PK":
        wiersze, _ = rk.wczytaj_xlsx(dane)
    else:
        wiersze, _ = rk.wczytaj_csv(dane)
    naglowek, rekordy = wiersze[0], wiersze[1:]
    szer = len(raport["kolumny_raport"])
    naglowek = naglowek + [""] * (szer - len(naglowek))
    rekordy = [w + [""] * (szer - len(w)) for w in rekordy]
    numery_wierszy = list(range(2, len(rekordy) + 2))          # numer wiersza w pliku źródłowym

    dziennik, do_sprawdzenia = [], []

    def zapisz(kolumna, operacja, zmiany):
        if zmiany:
            dziennik.append({"kolumna": kolumna, "operacja": operacja, "komorki": len(zmiany),
                             "przyklady": [f"{rk.repr_wartosci(str(a))} → {rk.repr_wartosci(str(b))}"
                                           for a, b in list(dict.fromkeys(zmiany))[:MAKS_PRZYKLADOW]]})

    def sprawdz(wiersz, kolumna, problem, wartosc):
        do_sprawdzenia.append({"wiersz": wiersz, "kolumna": kolumna, "problem": problem,
                               "wartosc": rk.repr_wartosci(str(wartosc))})

    # --- 1. całe wiersze: duplikaty
    if ustawienia.get("duplikaty", True):
        widziane, nowe, nowe_nr, usuniete = set(), [], [], []
        for nr, w in zip(numery_wierszy, rekordy):
            t = tuple(w)
            if t in widziane:
                usuniete.append(nr)
                continue
            widziane.add(t)
            nowe.append(w)
            nowe_nr.append(nr)
        rekordy, numery_wierszy = nowe, nowe_nr
        if usuniete:
            dziennik.append({"kolumna": "(cały wiersz)", "operacja": "usunięto zdublowane wiersze",
                             "komorki": len(usuniete),
                             "przyklady": [f"wiersz {n}" for n in usuniete[:MAKS_PRZYKLADOW]]})

    # --- 2. kolumna po kolumnie
    kolumny = [[w[i] for w in rekordy] for i in range(szer)]
    nazwy = [n.strip() or f"kolumna_{i + 1}" for i, n in enumerate(naglowek)]
    if any(n != o for n, o in zip(nazwy, naglowek)) and ustawienia.get("naglowki", True):
        zapisz("(nagłówek)", "usunięto spacje z nazw kolumn",
               [(o, n) for n, o in zip(nazwy, naglowek) if n != o])
    else:
        nazwy = [n or f"kolumna_{i + 1}" for i, n in enumerate(naglowek)]
    typy_wyjscia = ["tekst"] * szer
    wybrane = ustawienia.get("kolumny", {})
    format_daty = ustawienia.get("format_daty", "RRRR-MM-DD")
    tak, nie = TAK_NIE[ustawienia.get("tak_nie", "tak/nie")]
    styl_telefonu = ustawienia.get("telefon", "600 100 200")

    for i, k in enumerate(raport["kolumny_raport"]):
        ops = set(wybrane.get(str(i), [s["id"] for s in sugestie(k) if s["domyslnie"]]))
        nazwa = nazwy[i]
        kol = kolumny[i]

        def zastosuj(opis, funkcja):
            zmiany = []
            for j, w in enumerate(kol):
                nowa = funkcja(w)
                if nowa != w:
                    zmiany.append((w, nowa))
                    kol[j] = nowa
            zapisz(nazwa, opis, zmiany)

        if "niewidoczne" in ops:
            zastosuj("usunięto znaki niewidoczne", lambda w: rk.NIEWIDOCZNE.sub("", w) if isinstance(w, str) else w)
        if "spacje" in ops:
            zastosuj("usunięto zbędne spacje",
                     lambda w: re.sub(r" {2,}", " ", w.replace(" ", " ")).strip() if isinstance(w, str) else w)
        if "kodowanie" in ops:
            for j, w in enumerate(kol):
                if isinstance(w, str) and rk.KRZAKI.search(w) and napraw_kodowanie(w) is None:
                    sprawdz(numery_wierszy[j], nazwa, "błędne kodowanie – nie da się odtworzyć znaków", w)
            zastosuj("naprawiono polskie znaki",
                     lambda w: (napraw_kodowanie(w) or w) if isinstance(w, str) and rk.KRZAKI.search(w) else w)
        if "ukryte_puste" in ops:
            zastosuj("zamieniono wpisy typu NULL / brak na puste",
                     lambda w: "" if isinstance(w, str) and w.strip().lower() in rk.PSEUDO_PUSTE else w)

        if "warianty" in ops or "literowki" in ops:
            licznik = Counter(w for w in kol if isinstance(w, str) and w.strip())
            grupy = defaultdict(Counter)
            for w, n in licznik.items():
                grupy[rk.klucz_wariantu(w)][w] += n
            mapa = {}
            if "warianty" in ops:
                for zapisy in grupy.values():
                    wzor = wybierz_wzorzec(zapisy)
                    for z in zapisy:
                        if z != wzor:
                            mapa[z] = wzor
                zastosuj("ujednolicono zapis do najczęstszego wariantu", lambda w: mapa.get(w, w))
            if "literowki" in ops:
                sumy = {kl: sum(z.values()) for kl, z in grupy.items()}
                mapa_lit = {}
                for kl in sumy:
                    if len(kl) < 4 or re.search(r"\d", kl):
                        continue
                    kand = [c for c in sumy if c != kl and sumy[c] >= 3 * sumy[kl]]
                    traf = rk.difflib.get_close_matches(kl, kand, n=1, cutoff=0.85)
                    if traf:
                        docelowy = wybierz_wzorzec(grupy[traf[0]])
                        for z in grupy[kl]:
                            mapa_lit[z] = docelowy
                zastosuj("poprawiono możliwe literówki", lambda w: mapa_lit.get(w, w))

        if "logiczne" in ops:
            zastosuj(f"ujednolicono tak / nie ({tak} / {nie})",
                     lambda w: (tak if rk.klucz_wariantu(w) in PRAWDA else nie)
                     if isinstance(w, str) and w.strip() and rk.klucz_wariantu(w) in rk.LOGICZNE else w)

        if "daty" in ops:
            zmiany, niejednoznaczne = [], k["problemy"]["niejednoznaczne daty"] > 0
            for j, w in enumerate(kol):
                if not isinstance(w, str) or not w.strip():
                    continue
                d, fmt = jako_data_lub_excel(w)
                if isinstance(d, date):
                    if d.year < 1950 or d.year > 2100:
                        sprawdz(numery_wierszy[j], nazwa, "data poza sensownym zakresem", w)
                    zmiany.append((w, d.strftime(FORMATY_DATY[format_daty])))
                    kol[j] = d
                elif d is None:
                    sprawdz(numery_wierszy[j], nazwa, "data niemożliwa (nie ma jej w kalendarzu)", w)
                elif w.strip():
                    sprawdz(numery_wierszy[j], nazwa, "wartość nie jest datą", w)
            zapisz(nazwa, f"zamieniono daty na format {format_daty}" +
                   (" (daty z ukośnikiem przyjęto jako dzień/miesiąc)" if niejednoznaczne else ""),
                   [(a, b) for a, b in zmiany if a.strip() != b])
            typy_wyjscia[i] = "data"

        if "liczby" in ops:
            zmiany = []
            for j, w in enumerate(kol):
                if not isinstance(w, str) or not w.strip():
                    continue
                x = jako_liczba_ogolnie(w)
                if x is None:
                    sprawdz(numery_wierszy[j], nazwa, "wartość nie jest liczbą", w)
                else:
                    if w.strip() != str(x):
                        zmiany.append((w, x))
                    kol[j] = x
            zapisz(nazwa, "zamieniono tekst na liczby", zmiany)
            typy_wyjscia[i] = "liczba"

        if "masa_kg" in ops:
            zmiany = []
            for j, w in enumerate(kol):
                if isinstance(w, str) and w.strip():
                    x = masa_w_kg(w)
                    if x is None:
                        sprawdz(numery_wierszy[j], nazwa, "nie rozpoznano masy", w)
                    else:
                        zmiany.append((w, x))
                        kol[j] = x
            zapisz(nazwa, "przeliczono masę na liczby w kilogramach", zmiany)
            nazwy[i] = f"{nazwa} [kg]"
            typy_wyjscia[i] = "liczba"

        if "jednostki" in ops:
            def jednolita(w):
                if not isinstance(w, str):
                    return w
                j = rk.jednostka_w_wartosci(w)
                return w.strip()[: len(w.strip()) - len(j[1])].strip() + " " + JEDNOSTKA_KROTKA[j[0]] \
                    if j and w.strip().lower() != j[1].lower() else (JEDNOSTKA_KROTKA[j[0]] if j else w)
            zastosuj("ujednolicono zapis jednostek", jednolita)

        if "zera" in ops:
            dominujaca = Counter(len(w.strip()) for w in kol if isinstance(w, str) and re.fullmatch(r"\d+", w.strip()))
            dl = dominujaca.most_common(1)[0][0] if dominujaca else 0

            def uzupelnij(w):
                if not isinstance(w, str):
                    return w
                s = w.strip()
                if re.fullmatch(r"\d-\d{3}", s):
                    return "0" + s
                if dl and re.fullmatch(r"\d+", s) and len(s) < dl:
                    return s.zfill(dl)
                return w
            zastosuj("uzupełniono zgubione zera na początku", uzupelnij)

        if "kod_pocztowy" in ops:
            def kod(w):
                if not isinstance(w, str):
                    return w
                c = re.sub(r"\D", "", w)
                return f"{c[:2]}-{c[2:]}" if len(c) == 5 else w
            zastosuj("ujednolicono kody pocztowe (00-000)", kod)

        if "telefon" in ops:
            def tel(w):
                if not isinstance(w, str) or not w.strip():
                    return w
                c = re.sub(r"\D", "", w)
                if len(c) == 11 and c.startswith("48"):
                    c = c[2:]
                if len(c) != 9:
                    return w
                if styl_telefonu == "600100200":
                    return c
                grup = f"{c[:3]} {c[3:6]} {c[6:]}"
                return "+48 " + grup if styl_telefonu == "+48 600 100 200" else grup
            zastosuj(f"ujednolicono zapis telefonów ({styl_telefonu})", tel)

        if "cyfry" in ops:
            zastosuj("zostawiono same cyfry",
                     lambda w: re.sub(r"^PL", "", re.sub(r"[\s\-]", "", w), flags=re.I) if isinstance(w, str) else w)
        if "iban" in ops:
            zastosuj("ujednolicono IBAN", lambda w: re.sub(r"\s", "", w).upper() if isinstance(w, str) else w)
        if "email" in ops:
            zastosuj("ujednolicono e-maile", lambda w: w.strip().lower() if isinstance(w, str) else w)

        # --- rzeczy do ręcznego sprawdzenia (bez automatycznej zmiany)
        if k["numer"]:
            rodzaj, poprawny = rk.rozpoznaj_numer(k["nazwa"], [w for w in kol if isinstance(w, str) and w.strip()]) \
                or (None, None)
            if poprawny:
                for j, w in enumerate(kol):
                    if isinstance(w, str) and w.strip() and not poprawny(w):
                        sprawdz(numery_wierszy[j], nazwa, f"niepoprawny {k['numer']['rodzaj']}", w)
        if k["odstajace"]["ile"]:
            podejrzane = {float(x) for x in k["odstajace"]["przyklady"]}
            for j, w in enumerate(kol):
                x = w if isinstance(w, (int, float)) else (rk.jako_liczba(w) if isinstance(w, str) else None)
                if x is not None and float(x) in podejrzane:
                    sprawdz(numery_wierszy[j], nazwa, "wartość nietypowo duża (literówka?)", w)
        for j, w in enumerate(kol):
            if isinstance(w, str) and rk.FORMATY_LICZB[4][1].match(w.strip()):
                sprawdz(numery_wierszy[j], nazwa, "zapis naukowy – cyfry mogły zostać utracone", w)
        if k["uciete"]["ile"]:
            # jedna zbiorcza pozycja – setki wierszy z tym samym problemem zasłoniłyby resztę listy
            dl = k["uciete"]["dlugosc"]
            uciete_wiersze = [numery_wierszy[j] for j, w in enumerate(kol) if isinstance(w, str) and len(w) == dl]
            if uciete_wiersze:
                sprawdz(None, nazwa, f"możliwe ucięcie: {len(uciete_wiersze)} wartości ma dokładnie {dl} znaków "
                                     f"(wiersze {', '.join(map(str, uciete_wiersze[:10]))}"
                                     f"{'…' if len(uciete_wiersze) > 10 else ''})", k["uciete"]["przyklady"][0])

    # --- 3. powtórzone identyfikatory i sprzeczne daty (na danych po oczyszczeniu)
    for i, k in enumerate(raport["kolumny_raport"]):
        if k.get("powtorzone_id"):
            wiersze_dla = defaultdict(list)
            for j, w in enumerate(kolumny[i]):
                if str(w).strip():
                    wiersze_dla[str(w).strip()].append(j)
            for wartosc, idx in wiersze_dla.items():
                if len(idx) > 1 and len({tuple(str(kolumny[c][j]) for c in range(szer)) for j in idx}) > 1:
                    for j in idx:
                        sprawdz(numery_wierszy[j], nazwy[i], "ten sam identyfikator w różnych wierszach", wartosc)
    for para in raport["kolejnosc_dat"]:
        a = next(i for i, k in enumerate(raport["kolumny_raport"]) if k["nazwa"] == para["poczatek"])
        b = next(i for i, k in enumerate(raport["kolumny_raport"]) if k["nazwa"] == para["koniec"])
        for j in range(len(rekordy)):
            x, y = kolumny[a][j], kolumny[b][j]
            x = x if isinstance(x, date) else jako_data_lub_excel(x)[0] if isinstance(x, str) else None
            y = y if isinstance(y, date) else jako_data_lub_excel(y)[0] if isinstance(y, str) else None
            if isinstance(x, date) and isinstance(y, date) and y < x:
                sprawdz(numery_wierszy[j], nazwy[b], f"„{nazwy[b]}” wcześniej niż „{nazwy[a]}”",
                        f"{x.isoformat()} → {y.isoformat()}")

    # --- 4. kolumny puste i stałe
    zostaw = []
    for i, k in enumerate(raport["kolumny_raport"]):
        pusta = not any(str(w).strip() for w in kolumny[i])
        stala = not pusta and len({str(w) for w in kolumny[i]}) == 1
        if pusta and ustawienia.get("puste_kolumny", True):
            dziennik.append({"kolumna": nazwy[i], "operacja": "usunięto pustą kolumnę", "komorki": len(kolumny[i]),
                             "przyklady": []})
        elif stala and ustawienia.get("stale_kolumny", False):
            dziennik.append({"kolumna": nazwy[i], "operacja": "usunięto kolumnę ze stałą wartością",
                             "komorki": len(kolumny[i]), "przyklady": [rk.repr_wartosci(str(kolumny[i][0]))]})
        else:
            zostaw.append(i)

    do_sprawdzenia.sort(key=lambda x: (x["wiersz"] or 0, x["kolumna"]))      # zbiorcze (bez wiersza) na górze
    wynik_naglowek = [nazwy[i] for i in zostaw]
    wynik_typy = [typy_wyjscia[i] for i in zostaw]
    wynik_wiersze = [[kolumny[i][j] for i in zostaw] for j in range(len(rekordy))]

    if ustawienia.get("kolumna_uwag", False):
        uwagi_dla = defaultdict(list)
        for x in do_sprawdzenia:
            uwagi_dla[x["wiersz"]].append(f"{x['kolumna']}: {x['problem']}")
        wynik_naglowek.append("do_sprawdzenia")
        wynik_typy.append("tekst")
        for nr, w in zip(numery_wierszy, wynik_wiersze):
            w.append("; ".join(uwagi_dla.get(nr, [])))

    baza = re.sub(r"\.[^.]+$", "", nazwa_pliku or "dane") + "_oczyszczone"
    if ustawienia.get("format", "xlsx") == "csv":
        plik = do_csv(wynik_naglowek, wynik_wiersze, ustawienia)
        nazwa_wyniku = baza + ".csv"
    else:
        plik = do_xlsx(wynik_naglowek, wynik_typy, wynik_wiersze, dziennik, do_sprawdzenia, ustawienia)
        nazwa_wyniku = baza + ".xlsx"

    return {
        "plik": plik,
        "nazwa_pliku": nazwa_wyniku,
        "dziennik_csv": dziennik_do_csv(dziennik, do_sprawdzenia, ustawienia),
        "nazwa_dziennika": baza + "_dziennik_zmian.csv",
        "dziennik": dziennik,
        "do_sprawdzenia": do_sprawdzenia,
        "podsumowanie": {
            "wiersze_przed": raport["wiersze"], "wiersze_po": len(wynik_wiersze),
            "kolumny_przed": raport["kolumny"], "kolumny_po": len(wynik_naglowek),
            "zmienione_komorki": sum(d["komorki"] for d in dziennik if d["kolumna"] != "(cały wiersz)"
                                     and not d["operacja"].startswith("usunięto pustą")
                                     and not d["operacja"].startswith("usunięto kolumnę")),
            "do_sprawdzenia": len(do_sprawdzenia),
        },
    }


# ---------------------------------------------------------------------------
# Eksport
# ---------------------------------------------------------------------------

def _tekst_csv(w, ustawienia):
    if isinstance(w, date):
        return w.strftime(FORMATY_DATY[ustawienia.get("format_daty", "RRRR-MM-DD")])
    if isinstance(w, float):
        s = f"{w:.10f}".rstrip("0").rstrip(".")          # 259.0 -> "259", 1.25 -> "1.25"
        return s.replace(".", ",") if ustawienia.get("separator", ";") == ";" else s
    return str(w)


def do_csv(naglowek, wiersze, ustawienia):
    """CSV dla polskiego Excela: średnik + przecinek dziesiętny albo przecinek + kropka; UTF-8 z BOM."""
    bufor = io.StringIO()
    zapis = csv.writer(bufor, delimiter=ustawienia.get("separator", ";"), lineterminator="\r\n")
    zapis.writerow(naglowek)
    for w in wiersze:
        zapis.writerow([_tekst_csv(x, ustawienia) for x in w])
    return ("﻿" + bufor.getvalue()).encode("utf-8")


def dziennik_do_csv(dziennik, do_sprawdzenia, ustawienia):
    bufor = io.StringIO()
    zapis = csv.writer(bufor, delimiter=ustawienia.get("separator", ";"), lineterminator="\r\n")
    zapis.writerow(["rodzaj", "kolumna", "operacja / problem", "liczba komórek / wiersz", "przykłady / wartość"])
    for d in dziennik:
        zapis.writerow(["zmiana", d["kolumna"], d["operacja"], d["komorki"], " | ".join(d["przyklady"])])
    for x in do_sprawdzenia:
        zapis.writerow(["do sprawdzenia", x["kolumna"], x["problem"], x["wiersz"], x["wartosc"]])
    return ("﻿" + bufor.getvalue()).encode("utf-8")


def _komorka_tekstowa(ws, wiersz, kolumna, wartosc):
    """Tekst zapisujemy zawsze jako tekst – inaczej '=...' stałoby się formułą Excela."""
    c = ws.cell(wiersz, kolumna, wartosc)
    if isinstance(wartosc, str):
        c.data_type = "s"
    return c


def do_xlsx(naglowek, typy, wiersze, dziennik, do_sprawdzenia, ustawienia):
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter

    fiolet = PatternFill("solid", fgColor="6D3A62")
    bialy = Font(bold=True, color="FFFFFF")
    format_daty = {"RRRR-MM-DD": "yyyy-mm-dd", "DD.MM.RRRR": "dd.mm.yyyy", "DD/MM/RRRR": "dd/mm/yyyy",
                   "DD-MM-RRRR": "dd-mm-yyyy"}[ustawienia.get("format_daty", "RRRR-MM-DD")]

    wb = Workbook()
    ws = wb.active
    ws.title = "Dane"
    for c, n in enumerate(naglowek, 1):
        k = _komorka_tekstowa(ws, 1, c, n)
        k.fill, k.font = fiolet, bialy
    for r, w in enumerate(wiersze, 2):
        for c, x in enumerate(w, 1):
            if x == "":
                continue
            k = _komorka_tekstowa(ws, r, c, x)
            if isinstance(x, date):
                k.number_format = format_daty
    for c, n in enumerate(naglowek, 1):
        dl = max([len(str(n))] + [len(str(w[c - 1])) for w in wiersze[:300]])
        ws.column_dimensions[get_column_letter(c)].width = min(max(10, dl + 2), 50)
    ws.freeze_panes = "A2"
    if naglowek:
        ws.auto_filter.ref = f"A1:{get_column_letter(len(naglowek))}{len(wiersze) + 1}"

    wd = wb.create_sheet("Dziennik zmian")
    for c, n in enumerate(["Kolumna", "Operacja", "Liczba komórek", "Przykłady (przed → po)"], 1):
        k = wd.cell(1, c, n)
        k.fill, k.font = fiolet, bialy
    for r, d in enumerate(dziennik, 2):
        for c, x in enumerate([d["kolumna"], d["operacja"], d["komorki"], "\n".join(d["przyklady"])], 1):
            k = _komorka_tekstowa(wd, r, c, x)
            k.alignment = Alignment(wrap_text=True, vertical="top")
    for kol, sz in zip("ABCD", (22, 48, 16, 60)):
        wd.column_dimensions[kol].width = sz

    wsp = wb.create_sheet("Do sprawdzenia")
    for c, n in enumerate(["Wiersz w pliku źródłowym", "Kolumna", "Problem", "Wartość"], 1):
        k = wsp.cell(1, c, n)
        k.fill, k.font = fiolet, bialy
    for r, x in enumerate(do_sprawdzenia, 2):
        for c, v in enumerate([x["wiersz"], x["kolumna"], x["problem"], x["wartosc"]], 1):
            _komorka_tekstowa(wsp, r, c, v)
    for kol, sz in zip("ABCD", (12, 22, 48, 40)):
        wsp.column_dimensions[kol].width = sz
    wsp.freeze_panes = "A2"

    bufor = io.BytesIO()
    wb.save(bufor)
    return bufor.getvalue()
