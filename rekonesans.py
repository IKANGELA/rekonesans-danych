"""
Rekonesans danych – szybki raport jakości pliku CSV lub Excel.

Moduł nie korzysta z dysku ani sieci: dostaje bajty pliku i zwraca słownik z raportem.
Ten sam kod działa w przeglądarce (Pyodide) i w terminalu.

Wszystkie wartości analizujemy jako TEKST, dokładnie tak, jak są zapisane w pliku.
Dzięki temu widać prawdziwe formaty, spacje i błędy kodowania, zanim cokolwiek je zmieni.
"""

import csv
import difflib
import io
import math
import re
import statistics
import unicodedata
from collections import Counter, defaultdict
from datetime import date, datetime, timedelta

MAKS_PRZYKLADOW = 8

# Grupy problemów pokazywane jako kolumny tabeli podsumowania
GRUPY_PROBLEMOW = [
    ("Puste", ["ukryte puste"]),
    ("Spacje i znaki", ["spacje", "niewidoczne znaki"]),
    ("Kodowanie", ["kodowanie"]),
    ("Formaty", ["formaty", "mieszane typy", "zapis naukowy", "zgubione zera", "jednostki", "ucięte wartości"]),
    ("Daty", ["niemożliwe daty", "daty poza zakresem", "daty Excela", "niejednoznaczne daty"]),
    ("Warianty", ["warianty", "literówki", "tak/nie"]),
    ("Identyfikatory", ["powtórzone identyfikatory", "błędne numery"]),
    ("Odstające", ["odstające"]),
]

# ---------------------------------------------------------------------------
# Wzorce i słowniki
# ---------------------------------------------------------------------------

# Znaki typowe dla "krzaków": tekst UTF-8 odczytany jako Windows-1250/1252 lub odwrotnie
KRZAKI = re.compile(r"Ã.|Å.|Ä.|Ĺ.|Ă.|â€|Â |�")

# Znaki niewidoczne: tabulator, entery, spacje zerowej szerokości, miękki dywiz, BOM
NIEWIDOCZNE = re.compile(r"[\t\r\n​‌‍⁠﻿­]")
NAZWY_NIEWIDOCZNYCH = {"\t": "⟨TAB⟩", "\r": "⟨CR⟩", "\n": "⏎", "​": "⟨ZWSP⟩", "‌": "⟨ZWNJ⟩",
                       "‍": "⟨ZWJ⟩", "⁠": "⟨WJ⟩", "﻿": "⟨BOM⟩", "­": "⟨SHY⟩"}

PSEUDO_PUSTE = {"null", "none", "nan", "n/a", "na", "brak", "-", "--", "?", "#n/a", "b/d"}
LOGICZNE = {"tak", "nie", "t", "n", "y", "yes", "no", "true", "false", "prawda", "fałsz", "falsz",
            "1", "0", "x", "✓", "✔"}

MIESIACE = {
    "sty": 1, "lut": 2, "mar": 3, "kwi": 4, "maj": 5, "cze": 6, "lip": 7, "sie": 8, "wrz": 9,
    "paź": 10, "paz": 10, "lis": 11, "gru": 12, "stycznia": 1, "lutego": 2, "marca": 3,
    "kwietnia": 4, "maja": 5, "czerwca": 6, "lipca": 7, "sierpnia": 8, "września": 9,
    "wrzesnia": 9, "października": 10, "pazdziernika": 10, "listopada": 11, "grudnia": 12,
    "jan": 1, "feb": 2, "apr": 4, "may": 5, "jun": 6, "jul": 7, "aug": 8, "sep": 9, "oct": 10,
    "nov": 11, "dec": 12,
}
_MIES = "|".join(sorted(MIESIACE, key=len, reverse=True))

# (nazwa formatu, wyrażenie, kolejność grup: d/m/r)
FORMATY_DAT = [
    ("RRRR-MM-DD", re.compile(r"^(\d{4})-(\d{1,2})-(\d{1,2})$"), "rmd"),
    ("DD.MM.RRRR", re.compile(r"^(\d{1,2})\.(\d{1,2})\.(\d{4})$"), "dmr"),
    ("DD/MM/RRRR", re.compile(r"^(\d{1,2})/(\d{1,2})/(\d{4})$"), "dmr"),
    ("DD-MM-RRRR", re.compile(r"^(\d{1,2})-(\d{1,2})-(\d{4})$"), "dmr"),
    ("RRRR.MM.DD", re.compile(r"^(\d{4})\.(\d{1,2})\.(\d{1,2})$"), "rmd"),
    ("DD.MM.RR", re.compile(r"^(\d{1,2})\.(\d{1,2})\.(\d{2})$"), "dmr"),
    ("data z godziną", re.compile(r"^(\d{4})-(\d{1,2})-(\d{1,2})[ T]\d{1,2}:\d{2}(?::\d{2}(?:\.\d+)?)?$"), "rmd"),
    ("data z godziną", re.compile(r"^(\d{1,2})\.(\d{1,2})\.(\d{4}) \d{1,2}:\d{2}(?::\d{2})?$"), "dmr"),
    ("DD miesiąc RRRR", re.compile(rf"^(\d{{1,2}})\.? ?({_MIES})\.? (\d{{4}})$", re.I), "dMr"),
]

FORMATY_LICZB = [
    ("liczba całkowita", re.compile(r"^[+-]?\d+$")),
    ("liczba z kropką", re.compile(r"^[+-]?\d+\.\d+$")),
    ("liczba z przecinkiem", re.compile(r"^[+-]?\d+,\d+$")),
    ("liczba z separatorem tysięcy", re.compile(r"^[+-]?\d{1,3}([  .]\d{3})+(,\d+)?$")),
    ("zapis naukowy (np. 1,23E+05)", re.compile(r"^[+-]?\d+([.,]\d+)?[eE][+-]?\d+$")),
    ("kwota z walutą", re.compile(r"^[+-]?[\d  .,]+ ?(zł|pln|eur|€|\$|usd)$", re.I)),
    ("procent", re.compile(r"^[+-]?\d+([.,]\d+)? ?%$")),
]

JEDNOSTKI = {
    "kilogram": ["kg", "kg.", "kilogram", "kilogramy", "kilogramów", "kilo"],
    "gram": ["g", "g.", "gr", "gram", "gramy", "gramów"],
    "sztuka": ["szt", "szt.", "sztuka", "sztuki", "sztuk", "pcs"],
    "metr kwadratowy": ["m2", "m²", "mkw", "m.kw.", "m kw"],
    "metr bieżący": ["mb", "mb.", "m.b.", "metr bieżący"],
    "metr sześcienny": ["m3", "m³", "m.sz."],
    "litr": ["l", "l.", "litr", "litry", "litrów"],
    "opakowanie": ["op", "op.", "opak", "opakowanie"],
    "para": ["para", "pary", "par", "kpl", "kpl."],
}
JEDNOSTKA_DLA = {zapis: nazwa for nazwa, zapisy in JEDNOSTKI.items() for zapis in zapisy}

TYPOWE_LIMITY_DLUGOSCI = {20, 25, 30, 32, 35, 40, 50, 60, 64, 80, 100, 120, 128, 150, 200, 250, 255}


class BladPliku(Exception):
    """Pliku nie da się odczytać jako tabeli."""


# ---------------------------------------------------------------------------
# Odczyt pliku
# ---------------------------------------------------------------------------

def dekoduj(dane):
    """Zwraca (tekst, nazwa_kodowania). Najpierw UTF-8, potem Windows-1250."""
    if dane.startswith(b"\xef\xbb\xbf"):
        return dane[3:].decode("utf-8", errors="replace"), "UTF-8 (z BOM)"
    try:
        return dane.decode("utf-8"), "UTF-8"
    except UnicodeDecodeError:
        return dane.decode("cp1250", errors="replace"), "Windows-1250"


def wczytaj_csv(dane):
    tekst, kodowanie = dekoduj(dane)
    probka = tekst[:20000]
    try:
        separator = csv.Sniffer().sniff(probka, delimiters=";,\t|").delimiter
    except csv.Error:
        separator = max(";,\t|", key=probka.count)
    wiersze = [w for w in csv.reader(io.StringIO(tekst), delimiter=separator) if any(k.strip() for k in w)]
    if len(wiersze) < 2:
        raise BladPliku("Plik ma mniej niż dwa wiersze – brak danych do analizy.")
    nazwy_sep = {";": "średnik ;", ",": "przecinek ,", "\t": "tabulator", "|": "pionowa kreska |"}
    return wiersze, {"format": "CSV", "kodowanie": kodowanie, "separator": nazwy_sep[separator]}


def wczytaj_xlsx(dane):
    from openpyxl import load_workbook
    wb = load_workbook(io.BytesIO(dane), read_only=True, data_only=True)
    ws = wb.worksheets[0]
    wiersze = []
    for w in ws.iter_rows(values_only=True):
        wiersz = []
        for v in w:
            if v is None:
                wiersz.append("")
            elif isinstance(v, datetime):
                wiersz.append(v.strftime("%Y-%m-%d") if v.time() == datetime.min.time() else v.isoformat(" "))
            elif isinstance(v, date):
                wiersz.append(v.isoformat())
            else:
                wiersz.append(str(v))
        if any(k.strip() for k in wiersz):
            wiersze.append(wiersz)
    if len(wiersze) < 2:
        raise BladPliku("Pierwszy arkusz ma mniej niż dwa wiersze – brak danych do analizy.")
    return wiersze, {"format": f"Excel (arkusz „{ws.title}”)", "kodowanie": "nie dotyczy (Excel)",
                     "separator": "nie dotyczy (Excel)"}


# ---------------------------------------------------------------------------
# Pomocnicze: wartości
# ---------------------------------------------------------------------------

def wzorzec(wartosc):
    """Kształt wartości: cyfra -> 9, ciąg liter -> a, reszta bez zmian. '08 maj 2023' -> '99 a 9999'."""
    w = re.sub(r"\d", "9", wartosc.strip())
    return re.sub(r"[^\W\d_]+", "a", w)


def klucz_wariantu(wartosc):
    """Wartość po ujednoliceniu: małe litery, bez polskich znaków i nadmiarowych spacji."""
    w = unicodedata.normalize("NFKD", wartosc.strip().lower().replace("ł", "l"))
    w = "".join(z for z in w if not unicodedata.combining(z))
    w = NIEWIDOCZNE.sub("", w).replace(" ", " ")
    return re.sub(r"\s+", " ", w)


def repr_wartosci(w):
    """Wartość do pokazania w raporcie: spacje na brzegach jako ·, twarde spacje jako ⍽, znaki niewidoczne opisane."""
    lewa = len(w) - len(w.lstrip(" "))
    prawa = len(w) - len(w.rstrip(" "))
    srodek = w.strip(" ").replace(" ", "⍽")
    srodek = re.sub(r"  +", lambda m: "·" * len(m.group()), srodek)
    srodek = NIEWIDOCZNE.sub(lambda m: NAZWY_NIEWIDOCZNYCH[m.group()], srodek)
    return "·" * lewa + srodek + "·" * prawa


def liczba_tekst(x):
    return str(int(x)) if x == int(x) else f"{x:g}"


def rodzaj_wartosci(wartosc):
    w = wartosc.strip()
    for nazwa, wzor, _ in FORMATY_DAT:
        if wzor.match(w):
            return "data", nazwa
    for nazwa, wzor in FORMATY_LICZB:
        if wzor.match(w):
            return "liczba", nazwa
    return "tekst", "tekst"


def jako_liczba(wartosc):
    w = wartosc.strip().replace(" ", "").replace(" ", "")
    if re.match(r"^[+-]?\d+(,\d+)?([eE][+-]?\d+)?$", w):
        w = w.replace(",", ".")
    try:
        return float(w)
    except ValueError:
        return None


def jako_data(wartosc):
    """Zwraca (date | None, format, (pierwsza, druga) liczba dla dat z '/'). None = data niemożliwa."""
    w = wartosc.strip()
    for nazwa, wzor, kolejnosc in FORMATY_DAT:
        m = wzor.match(w)
        if not m:
            continue
        g = m.groups()
        try:
            if kolejnosc == "rmd":
                return date(int(g[0]), int(g[1]), int(g[2])), nazwa
            if kolejnosc == "dMr":
                return date(int(g[2]), MIESIACE[g[1].lower()], int(g[0])), nazwa
            rok = int(g[2]) + (2000 if len(g[2]) == 2 else 0)
            return date(rok, int(g[1]), int(g[0])), nazwa
        except ValueError:
            return None, nazwa
    return "brak", None


def jednostka_w_wartosci(wartosc):
    """Zwraca (jednostka, zapis), jeśli wartość jest jednostką albo kończy się jednostką: '12 kg'."""
    w = wartosc.strip().lower()
    if w in JEDNOSTKA_DLA:
        return JEDNOSTKA_DLA[w], wartosc.strip()
    m = re.match(r"^[\d.,\s]+([a-ząćęłńóśźż²³.]+)$", w)
    if m and m.group(1) in JEDNOSTKA_DLA:
        return JEDNOSTKA_DLA[m.group(1)], m.group(1)
    return None


# ---------------------------------------------------------------------------
# Polskie numery z sumą kontrolną
# ---------------------------------------------------------------------------

def tylko_cyfry(w):
    return re.sub(r"[\s\-]", "", w.strip())


def nip_ok(w):
    c = tylko_cyfry(w)
    if c.upper().startswith("PL"):
        c = c[2:]
    if not re.fullmatch(r"\d{10}", c):
        return False
    s = sum(int(a) * b for a, b in zip(c, [6, 5, 7, 2, 3, 4, 5, 6, 7])) % 11
    return s != 10 and s == int(c[9])


def pesel_ok(w):
    c = tylko_cyfry(w)
    if not re.fullmatch(r"\d{11}", c):
        return False
    s = sum(int(a) * b for a, b in zip(c, [1, 3, 7, 9, 1, 3, 7, 9, 1, 3])) % 10
    return (10 - s) % 10 == int(c[10])


def regon_ok(w):
    c = tylko_cyfry(w)
    wagi = {9: [8, 9, 2, 3, 4, 5, 6, 7], 14: [2, 4, 8, 5, 0, 9, 7, 3, 6, 1, 2, 4, 8]}.get(len(c))
    if not wagi or not c.isdigit():
        return False
    return sum(int(a) * b for a, b in zip(c, wagi)) % 11 % 10 == int(c[-1])


def iban_ok(w):
    c = re.sub(r"\s", "", w.strip()).upper()
    if re.fullmatch(r"\d{26}", c):
        c = "PL" + c
    if not re.fullmatch(r"[A-Z]{2}\d{2}[A-Z0-9]{10,30}", c):
        return False
    przest = c[4:] + c[:4]
    return int("".join(str(int(z, 36)) for z in przest)) % 97 == 1


EMAIL = re.compile(r"^[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}$")
TELEFON = re.compile(r"^(\+?48[\s\-]?)?(\(?\d{2,3}\)?[\s\-]?)?\d{2,3}[\s\-]?\d{2,3}[\s\-]?\d{2,3}$")
KOD_POCZTOWY = re.compile(r"^\d{2}-\d{3}$")


def rozpoznaj_numer(nazwa, wartosci):
    """Rozpoznaje kolumnę z numerami (NIP, PESEL, REGON, IBAN, e-mail, telefon, kod pocztowy)."""
    n = nazwa.lower()
    if not wartosci:
        return None
    udzial = lambda f: sum(1 for w in wartosci if f(w)) / len(wartosci)
    cyfry = [tylko_cyfry(w) for w in wartosci]

    if "mail" in n or udzial(lambda w: "@" in w) > 0.5:
        return "e-mail", lambda w: bool(EMAIL.match(w.strip()))
    if any(s in n for s in ("iban", "konto", "rachunek")) or udzial(lambda w: len(re.sub(r"\D", "", w)) == 26) > 0.6:
        return "IBAN", iban_ok
    if "pesel" in n or (all(len(c) == 11 for c in cyfry) and udzial(pesel_ok) > 0.6):
        return "PESEL", pesel_ok
    if "nip" in n or (udzial(lambda w: len(re.sub(r"\D", "", w)) == 10) > 0.8 and udzial(nip_ok) > 0.6):
        return "NIP", nip_ok
    if "regon" in n:
        return "REGON", regon_ok
    if any(s in n for s in ("kod poczt", "kod_poczt", "kodpoczt", "pocztowy")) or (
            udzial(lambda w: re.fullmatch(r"\d{2}-?\s?\d{3}|\d-\d{3}|\d{4}", w.strip()) is not None) > 0.8
            and udzial(lambda w: KOD_POCZTOWY.match(w.strip()) is not None) > 0.5):
        return "kod pocztowy", lambda w: bool(KOD_POCZTOWY.match(w.strip()))
    if any(s in n for s in ("tel", "phone", "komórk", "komork")) or (
            udzial(lambda w: TELEFON.match(w.strip()) is not None) > 0.8 and
            udzial(lambda w: len(re.sub(r"\D", "", w)) in (9, 11)) > 0.8):
        return "telefon", lambda w: len(re.sub(r"\D", "", w.strip()).removeprefix("48")) == 9
    return None


# ---------------------------------------------------------------------------
# Analiza kolumny
# ---------------------------------------------------------------------------

def analizuj_kolumne(nazwa, wartosci, dzis=None):
    dzis = dzis or date.today()
    n = len(wartosci)
    puste = sum(1 for w in wartosci if w.strip() == "")
    pseudo = Counter(w.strip() for w in wartosci if w.strip().lower() in PSEUDO_PUSTE)
    niepuste = [w for w in wartosci if w.strip() != "" and w.strip().lower() not in PSEUDO_PUSTE]
    unikalne = Counter(niepuste)
    uwagi = []

    def przyklady(lista):
        return [repr_wartosci(w) for w in list(dict.fromkeys(lista))[:MAKS_PRZYKLADOW]]

    # --- spacje i znaki niewidoczne
    spacje_brzegi = [w for w in niepuste if w != w.strip(" ")]
    spacje_podwojne = [w for w in niepuste if "  " in w.strip(" ")]
    twarde_spacje = [w for w in niepuste if " " in w]
    niewidoczne = [w for w in niepuste if NIEWIDOCZNE.search(w)]
    krzaki = [w for w in niepuste if KRZAKI.search(w)]

    # --- rodzaje i formaty
    rodzaje, formaty = Counter(), Counter()
    przyklady_formatow = defaultdict(list)
    for w in niepuste:
        r, f = rodzaj_wartosci(w)
        rodzaje[r] += 1
        formaty[f] += 1
        if len(przyklady_formatow[f]) < 3:
            przyklady_formatow[f].append(repr_wartosci(w))
    typ = rodzaje.most_common(1)[0][0] if rodzaje else "pusta"
    obce_typy = [w for w in niepuste if rodzaj_wartosci(w)[0] != typ] if typ != "tekst" else []
    wzorce = Counter(wzorzec(w) for w in niepuste)

    # --- daty: niemożliwe, poza zakresem, niejednoznaczne, zapisane jako liczby Excela
    niemozliwe, poza_zakresem, niejednoznaczne, daty_excela = [], [], [], []
    daty = []                                        # sparsowane daty (do kontroli kolejności dat)
    nazwa_mowi_data = bool(re.search(r"dat|date|dzie[nń]|termin|kiedy|okres", nazwa.lower()))
    if typ == "data" or nazwa_mowi_data:
        ukosniki = []
        for w in niepuste:
            d, fmt = jako_data(w)
            if d is None:
                niemozliwe.append(w)
            elif d != "brak":
                daty.append(d)
                if d.year < 1950 or d > dzis + timedelta(days=3650):
                    poza_zakresem.append(w)
                if fmt == "DD/MM/RRRR":
                    ukosniki.append(w)
            elif re.fullmatch(r"\d{5}([.,]\d+)?", w.strip()) and 20000 <= float(w.strip().replace(",", ".")) <= 60000:
                daty_excela.append(w)
        obce_typy = [w for w in obce_typy if w not in set(daty_excela)]   # nie liczymy ich podwójnie
        # "03/04/2025" – 3 kwietnia czy 4 marca? Niejednoznaczne, jeśli w kolumnie nic tego nie rozstrzyga
        if ukosniki:
            pary = [tuple(int(x) for x in w.strip().split("/")[:2]) for w in ukosniki]
            rozstrzyga = any(a > 12 for a, b in pary) or any(b > 12 for a, b in pary)
            if not rozstrzyga:
                niejednoznaczne = ukosniki

    # --- liczby: zapis naukowy, zgubione zera, statystyki, odstające
    naukowe = [w for w in niepuste if FORMATY_LICZB[4][1].match(w.strip())]
    zgubione_zera = []
    cyfrowe = [w.strip() for w in niepuste if re.fullmatch(r"\d+", w.strip())]
    if cyfrowe and len(cyfrowe) >= 0.8 * len(niepuste):
        dlugosci = Counter(len(w) for w in cyfrowe)
        dl, ile = dlugosci.most_common(1)[0]
        z_zerem = sum(1 for w in cyfrowe if len(w) == dl and w.startswith("0"))
        if dl >= 3 and ile >= 0.7 * len(cyfrowe) and z_zerem > 0:
            zgubione_zera = [w for w in cyfrowe if len(w) < dl]
    # kody pocztowe bez pierwszego zera: "2-345" zamiast "02-345"
    zgubione_zera += [w for w in niepuste if re.fullmatch(r"\d-\d{3}", w.strip())]

    statystyki, odstajace = None, []
    if typ == "liczba":
        liczby = [x for x in (jako_liczba(w) for w in niepuste) if x is not None]
        if len(liczby) >= 10:
            dodatnie = [x for x in liczby if x > 0]
            if len(dodatnie) >= 10 and len(set(dodatnie)) >= 3:
                q = statistics.quantiles([math.log10(x) for x in dodatnie], n=4)
                gorna = q[2] + 2 * max(q[2] - q[0], 0.1)
                mediana = statistics.median(dodatnie)
                # Podejrzane wartości muszą być:
                #  1) duże: ponad górnym "płotem" w skali log albo co najmniej 8x większe od mediany,
                #  2) NIELICZNE: najwyżej 1% wierszy (literówka to wyjątek, a nie stała grupa),
                #  3) ODERWANE od reszty: o co najmniej 50% większe od największej zwykłej wartości.
                # Dzięki temu "310" przy typowych 5–60 jest zgłaszane, a zamówienia hurtowe
                # na 20–40 sztuk (liczna grupa, płynnie przechodząca w resztę) – już nie.
                duze = [x for x in dodatnie if math.log10(x) > gorna or x >= 8 * mediana]
                zwykle = [x for x in dodatnie if not (math.log10(x) > gorna or x >= 8 * mediana)]
                if duze and zwykle and len(duze) <= max(3, 0.01 * len(dodatnie)):
                    granica = 1.5 * max(zwykle)
                    odstajace = sorted((x for x in duze if x > granica), key=lambda x: -x)
            statystyki = {"min": liczba_tekst(min(liczby)), "mediana": liczba_tekst(statistics.median(liczby)),
                          "max": liczba_tekst(max(liczby)), "ujemne": sum(1 for x in liczby if x < 0)}

    # --- warianty, literówki, tak/nie (kolumny tekstowe z powtarzającymi się wartościami)
    warianty, literowki, logiczne = [], [], {}
    kategoryczna = typ == "tekst" and unikalne and len(unikalne) <= max(50, n // 2)
    if kategoryczna:
        grupy = defaultdict(Counter)
        for w, ile in unikalne.items():
            grupy[klucz_wariantu(w)][w] += ile
        for zapisy in grupy.values():
            if len(zapisy) > 1:
                warianty.append([{"zapis": repr_wartosci(z), "ile": i} for z, i in zapisy.most_common()])
        warianty.sort(key=lambda g: -sum(x["ile"] for x in g))

        # literówki: rzadka wartość bardzo podobna do częstej (np. "Warszwa" i "Warszawa")
        liczebnosc = Counter()
        for klucz, zapisy in grupy.items():
            liczebnosc[klucz] = sum(zapisy.values())
        klucze = [k for k in liczebnosc if len(k) >= 4 and not re.search(r"\d", k)]
        if len(klucze) <= 400:
            czeste = sorted(klucze, key=lambda k: -liczebnosc[k])
            for k in klucze:
                kandydaci = [c for c in czeste if c != k and liczebnosc[c] >= 3 * liczebnosc[k]]
                trafienie = difflib.get_close_matches(k, kandydaci, n=1, cutoff=0.85)
                if trafienie:
                    rzadki = grupy[k].most_common(1)[0][0]
                    czesty = grupy[trafienie[0]].most_common(1)[0][0]
                    literowki.append({"rzadki": repr_wartosci(rzadki), "ile_rzadki": liczebnosc[k],
                                      "czesty": repr_wartosci(czesty), "ile_czesty": liczebnosc[trafienie[0]]})

        # tak/nie zapisane na kilka sposobów
        znormalizowane = {klucz_wariantu(w) for w in unikalne}
        if znormalizowane and znormalizowane <= LOGICZNE and len(unikalne) > 2:
            logiczne = dict(unikalne.most_common())

    # --- jednostki zapisane na kilka sposobów
    jedn = defaultdict(Counter)
    for w in niepuste:
        j = jednostka_w_wartosci(w)
        if j:
            jedn[j[0]][j[1]] += 1
    jednostki = {k: dict(v.most_common()) for k, v in jedn.items() if len(v) > 1}
    if len(jedn) > 1:
        jednostki["różne jednostki w kolumnie"] = {k: sum(v.values()) for k, v in jedn.items()}

    # --- ucięte wartości (dużo wartości o identycznej, "okrągłej" maksymalnej długości)
        # (sprawdzamy tylko tekst dłuższy niż kilka znaków)
    # Ucięcie rozpoznajemy po "ścianie": wiele wartości o okrągłej długości maksymalnej
    # i prawie żadnej o jeden znak krótszej (przy naturalnych długościach rozkład jest gładki).
    uciete = []
    if typ == "tekst" and niepuste:
        dlugosci_tekstu = Counter(len(w) for w in niepuste)
        maks = max(dlugosci_tekstu)
        na_maks = [w for w in niepuste if len(w) == maks]
        if (maks in TYPOWE_LIMITY_DLUGOSCI and len(set(na_maks)) >= 3 and len(na_maks) >= 5
                and len(na_maks) >= 3 * max(1, dlugosci_tekstu.get(maks - 1, 0))):
            uciete = na_maks

    # --- WIELKIE i małe litery wymieszane w całej kolumnie (informacyjnie)
    z_literami = [w for w in niepuste if re.search(r"[^\W\d_]{2,}", w)]
    if z_literami:
        wielkie = sum(1 for w in z_literami if w.isupper())
        if 0.1 <= wielkie / len(z_literami) <= 0.9:
            uwagi.append(f"Wymieszane WIELKIE i małe litery: {wielkie} z {len(z_literami)} wartości "
                         f"zapisano wielkimi literami.")

    # --- numery: NIP, PESEL, REGON, IBAN, e-mail, telefon, kod pocztowy
    numer = None
    rozpoznany = rozpoznaj_numer(nazwa, niepuste) if niepuste and typ != "data" else None
    if rozpoznany:
        rodzaj, poprawny = rozpoznany
        bledne = [w for w in niepuste if not poprawny(w)]
        if rodzaj == "kod pocztowy":            # kody bez zera są już policzone jako "zgubione zera"
            bledne = [w for w in bledne if w not in set(zgubione_zera)]
        numer = {"rodzaj": rodzaj, "poprawne": len(niepuste) - len(bledne), "bledne": len(bledne),
                 "przyklady": przyklady(bledne),
                 "zapisy": [{"wzorzec": p, "ile": i} for p, i in wzorce.most_common(6)]}
        # numer telefonu czy NIP to nie "liczba" – statystyki i odstające nie mają tu sensu;
        # problemem jest za to zapisywanie tego samego rodzaju numeru na kilka sposobów
        statystyki, odstajace, obce_typy = None, [], []
        formaty_numeru = max(0, len(wzorce) - 1) if rodzaj not in ("e-mail",) else 0

    # --- kandydat na identyfikator (do kontroli powtórzeń na poziomie pliku)
    identyfikator = bool(niepuste) and not (numer and numer["rodzaj"] in ("kod pocztowy", "telefon")) \
        and "poczt" not in nazwa.lower() and (
        re.search(r"(^|[_\s])(nr|numer|id|kod|symbol|sku|ean|nip|pesel|regon)($|[_\s])", nazwa.lower()) is not None
        or "numer" in nazwa.lower()
    ) and len(unikalne) >= 0.9 * len(niepuste) and len(niepuste) >= 20

    # --- kolumna stała / pusta
    if not niepuste:
        uwagi.append("Kolumna jest całkowicie pusta – może da się ją usunąć.")
    elif len(unikalne) == 1 and n > 1:
        uwagi.append(f"Wszędzie ta sama wartość („{repr_wartosci(niepuste[0])}”) – kolumna nic nie wnosi do analizy.")

    problemy = {
        "puste": puste + sum(pseudo.values()),
        "ukryte puste": sum(pseudo.values()),
        "spacje": sum(1 for w in niepuste if w != w.strip(" ") or "  " in w.strip(" ") or " " in w),
        "niewidoczne znaki": len(niewidoczne),
        "kodowanie": len(krzaki),
        "formaty": (formaty_numeru if numer else
                    max(0, len([f for f in formaty if f != "tekst"]) - 1) if typ in ("data", "liczba") else 0),
        "mieszane typy": len(obce_typy),
        "zapis naukowy": len(naukowe),
        "zgubione zera": len(zgubione_zera),
        "jednostki": len(jednostki),
        "ucięte wartości": len(uciete),
        "niemożliwe daty": len(niemozliwe),
        "daty poza zakresem": len(poza_zakresem),
        "daty Excela": len(daty_excela),
        "niejednoznaczne daty": len(niejednoznaczne),
        "warianty": len(warianty),
        "literówki": len(literowki),
        "tak/nie": 1 if logiczne else 0,
        "błędne numery": numer["bledne"] if numer else 0,
        "powtórzone identyfikatory": 0,          # uzupełniane na poziomie pliku
        "odstające": len(odstajace),
    }

    return {
        "nazwa": nazwa,
        "typ": typ,
        "wiersze": n,
        "unikalne": len(unikalne),
        "problemy": problemy,
        "uwagi": uwagi,
        "puste": {"puste": puste, "pseudo": dict(pseudo.most_common())},
        "spacje": {"na_brzegach": len(spacje_brzegi), "podwojne": len(spacje_podwojne),
                   "twarde": len(twarde_spacje),
                   "przyklady": przyklady(spacje_brzegi + spacje_podwojne + twarde_spacje)},
        "niewidoczne": {"ile": len(niewidoczne), "przyklady": przyklady(niewidoczne)},
        "kodowanie": {"ile": len(krzaki), "przyklady": przyklady(krzaki)},
        "formaty": [{"format": f, "ile": i, "przyklady": przyklady_formatow[f]} for f, i in formaty.most_common()],
        "obce_typy": {"ile": len(obce_typy), "przyklady": przyklady(obce_typy)},
        "wzorce": [{"wzorzec": w, "ile": i} for w, i in wzorce.most_common(10)],
        "liczba_wzorcow": len(wzorce),
        "daty": {"niemozliwe": przyklady(niemozliwe), "poza_zakresem": przyklady(poza_zakresem),
                 "excel": przyklady(daty_excela), "niejednoznaczne": przyklady(niejednoznaczne),
                 "zakres": [min(daty).isoformat(), max(daty).isoformat()] if daty else None},
        "naukowe": przyklady(naukowe),
        "zgubione_zera": przyklady(zgubione_zera),
        "warianty": warianty[:15],
        "literowki": literowki[:15],
        "logiczne": logiczne,
        "jednostki": jednostki,
        "uciete": {"ile": len(uciete), "dlugosc": max((len(w) for w in uciete), default=0),
                   "przyklady": przyklady(uciete)},
        "numer": numer,
        "identyfikator": identyfikator,
        "statystyki": statystyki,
        "odstajace": {"ile": len(odstajace), "przyklady": [liczba_tekst(x) for x in odstajace[:MAKS_PRZYKLADOW]]},
        "najczestsze": [{"wartosc": repr_wartosci(w), "ile": i} for w, i in unikalne.most_common(6)],
        # pełna lista przy małej liczbie różnych wartości – tu łatwo wyłapać synonimy (różowy / pink)
        "wszystkie_wartosci": ([{"wartosc": repr_wartosci(w), "ile": i}
                                for w, i in sorted(unikalne.items(), key=lambda x: klucz_wariantu(x[0]))]
                               if typ == "tekst" and 1 < len(unikalne) <= 40 else []),
        "_daty": daty_wierszy(wartosci) if (typ == "data" or nazwa_mowi_data) else None,
    }


def daty_wierszy(wartosci):
    """Daty w kolejności wierszy (None, gdy wartość nie jest poprawną datą) – do porównań między kolumnami."""
    wynik = []
    for w in wartosci:
        d, _ = jako_data(w) if w.strip() else (None, None)
        wynik.append(d if isinstance(d, date) else None)
    return wynik


# ---------------------------------------------------------------------------
# Kontrole na poziomie pliku
# ---------------------------------------------------------------------------

POCZATEK = re.compile(r"start|pocz|rozpocz|zamów|zamow|utworz|wystaw|przyj|od$|_od|begin|created|order", re.I)
KONIEC = re.compile(r"koniec|końc|konc|zakończ|zakoncz|wysył|wysyl|dostaw|realiz|termin|do$|_do|end|ship|deliver", re.I)


def kolejnosc_dat(naglowek, kolumny):
    """Szuka par kolumn dat typu początek–koniec i liczy wiersze, w których koniec jest przed początkiem."""
    wyniki = []
    daty = [(i, naglowek[i]) for i, k in enumerate(kolumny) if k.get("_daty")]
    for i, nazwa_a in daty:
        for j, nazwa_b in daty:
            if i == j or not POCZATEK.search(nazwa_a) or not KONIEC.search(nazwa_b) or KONIEC.search(nazwa_a):
                continue
            a, b = kolumny[i]["_daty"], kolumny[j]["_daty"]
            zle = [(x, y) for x, y in zip(a, b) if x and y and y < x]
            if zle:
                wyniki.append({"poczatek": nazwa_a, "koniec": nazwa_b, "ile": len(zle),
                               "przyklady": [f"{x.isoformat()} → {y.isoformat()}" for x, y in zle[:5]]})
    return wyniki


def powtorzone_identyfikatory(naglowek, rekordy, kolumny):
    """Ten sam identyfikator w RÓŻNYCH wierszach (całe zdublowane wiersze liczymy osobno)."""
    wyniki = []
    for i, k in enumerate(kolumny):
        if not k["identyfikator"]:
            continue
        wiersze_dla = defaultdict(set)
        for w in rekordy:
            if w[i].strip():
                wiersze_dla[w[i].strip()].add(tuple(w))
        rozne = {idx: wiersze for idx, wiersze in wiersze_dla.items() if len(wiersze) > 1}
        if rozne:
            k["problemy"]["powtórzone identyfikatory"] = len(rozne)
            k["powtorzone_id"] = [repr_wartosci(x) for x in list(rozne)[:MAKS_PRZYKLADOW]]
            wyniki.append({"kolumna": naglowek[i], "ile": len(rozne)})
    return wyniki


# ---------------------------------------------------------------------------
# Raport całego pliku
# ---------------------------------------------------------------------------

def analizuj(dane, nazwa_pliku="", dzis=None):
    """Bajty pliku CSV/XLSX -> słownik z raportem."""
    if nazwa_pliku.lower().endswith((".xlsx", ".xlsm")) or dane[:2] == b"PK":
        wiersze, info = wczytaj_xlsx(dane)
    else:
        wiersze, info = wczytaj_csv(dane)

    naglowek, rekordy = wiersze[0], wiersze[1:]
    szer = max(len(naglowek), max(len(w) for w in rekordy))
    zla_dlugosc = sum(1 for w in rekordy if len(w) != len(naglowek))
    naglowek = naglowek + [""] * (szer - len(naglowek))
    rekordy = [w + [""] * (szer - len(w)) for w in rekordy]

    duplikaty = sum(i - 1 for i in Counter(tuple(w) for w in rekordy).values() if i > 1)

    uwagi_naglowka = []
    for i, n in enumerate(naglowek):
        if n.strip() == "":
            uwagi_naglowka.append(f"Kolumna nr {i + 1} nie ma nazwy.")
        elif n != n.strip() or NIEWIDOCZNE.search(n):
            uwagi_naglowka.append(f"Nazwa „{repr_wartosci(n)}” ma zbędne spacje lub niewidoczne znaki.")
    for n, ile in Counter(x.strip() for x in naglowek).items():
        if n and ile > 1:
            uwagi_naglowka.append(f"Nazwa „{n}” występuje {ile} razy.")

    kolumny = [analizuj_kolumne(naglowek[i].strip() or f"(kolumna {i + 1})", [w[i] for w in rekordy], dzis)
               for i in range(szer)]
    kolejnosc = kolejnosc_dat([k["nazwa"] for k in kolumny], kolumny)
    powtorzone_identyfikatory([k["nazwa"] for k in kolumny], rekordy, kolumny)
    for k in kolumny:
        k.pop("_daty", None)
        k.setdefault("powtorzone_id", [])
        grupy = {g: sum(k["problemy"].get(p, 0) for p in klucze) for g, klucze in GRUPY_PROBLEMOW}
        k["grupy"] = grupy
        k["suma_problemow"] = sum(grupy.values())

    przecinek = sum(f["ile"] for k in kolumny for f in k["formaty"] if f["format"] == "liczba z przecinkiem")
    kropka = sum(f["ile"] for k in kolumny for f in k["formaty"] if f["format"] == "liczba z kropką")
    info["separator_dziesietny"] = ("przecinek" if przecinek > kropka else
                                    "kropka" if kropka else "brak liczb dziesiętnych")
    info["krzaki_w_pliku"] = sum(k["kodowanie"]["ile"] for k in kolumny)

    return {
        "plik": nazwa_pliku,
        "info": info,
        "wiersze": len(rekordy),
        "kolumny": len(naglowek),
        "duplikaty": duplikaty,
        "zla_liczba_kolumn": zla_dlugosc,
        "uwagi_naglowka": uwagi_naglowka,
        "kolejnosc_dat": kolejnosc,
        "grupy_problemow": [g for g, _ in GRUPY_PROBLEMOW],
        "kolumny_raport": kolumny,
    }
