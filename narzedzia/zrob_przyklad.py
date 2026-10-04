"""
Tworzy przykładowy plik przyklady/przyklad_zamowienia.csv – fikcyjny eksport zamówień
sklepu internetowego z celowo wprowadzonymi problemami jakości danych.
Wszystkie dane (osoby, NIP-y, telefony, adresy) są wymyślone.

    python narzedzia/zrob_przyklad.py
"""

import csv
import random
from datetime import date, timedelta
from pathlib import Path

random.seed(7)
CEL = Path(__file__).resolve().parent.parent / "przyklady" / "przyklad_zamowienia.csv"

PRODUKTY = [("Poduszka dekoracyjna", 79.0, 0.4, "5901234560017"), ("Koc bawełniany", 149.0, 1.2, "5901234560024"),
            ("Pościel satynowa", 259.0, 1.6, "5901234560031"), ("Ręcznik kąpielowy", 59.0, 0.6, "5901234560048"),
            ("Zasłona lniana", 189.0, 0.9, "5901234560055"), ("Narzuta pikowana", 299.0, 2.1, "5901234560062")]
KOLORY = {"różowy": ["różowy", "Różowy", "rozowy", "pink", " różowy"],
          "szary": ["szary", "Szary", "grey", "szary "],
          "biały": ["biały", "Biały", "bialy", "white"],
          "zielony": ["zielony", "butelkowa zieleń"]}
MIASTA = {"Lublin": "20-", "Warszawa": "02-", "Kraków": "30-", "Łódź": "90-", "Gdańsk": "80-",
          "Poznań": "60-", "Wrocław": "50-"}
ULICE = ["ul. Krakowskie Przedmieście", "al. Jana Pawła II", "ul. Generała Władysława Andersa",
         "ul. Kwiatowa", "ul. Lipowa", "ul. Marszałka Józefa Piłsudskiego"]
KLIENCI = ["Anna K.", "Marek W.", "Ewa Z.", "Piotr N.", "Kasia L.", "Jan B."]
STATUSY = ["Wysłane", "wysłane", "WYSLANE", "W realizacji", "Anulowane", "Zwrot"]
FAKTURA = ["tak", "nie", "nie", "nie", "TAK", "T", "N", "0", "1"]


def nip_poprawny():
    while True:
        c = [random.randint(1, 9)] + [random.randint(0, 9) for _ in range(8)]
        s = sum(a * b for a, b in zip(c, [6, 5, 7, 2, 3, 4, 5, 6, 7])) % 11
        if s != 10:
            n = "".join(map(str, c + [s]))
            return random.choice([n, f"{n[:3]}-{n[3:6]}-{n[6:8]}-{n[8:]}"])


def data_tekst(d, i):
    if i % 9 == 0:
        return d.strftime("%d.%m.%Y")
    if i % 13 == 0:
        m = ["sty", "lut", "mar", "kwi", "maj", "cze", "lip", "sie", "wrz", "paź", "lis", "gru"][d.month - 1]
        return f"{d.day:02d} {m} {d.year}"
    return d.isoformat()


def waga_tekst(kg, i):
    if i % 11 == 0:
        return f"{round(kg * 1000)} g"
    if i % 17 == 0:
        return f"{kg:.1f} kilogram".replace(".", ",")
    return f"{kg:.1f} kg".replace(".", ",") if i % 2 else f"{kg:.1f} kg"


def telefon(i):
    c = "".join(str(random.randint(0, 9)) for _ in range(8))
    n = random.choice("5678") + c
    if i in (88, 301):
        return n[:8]                                   # za krótki numer
    return random.choice([n, f"{n[:3]} {n[3:6]} {n[6:]}", f"+48 {n[:3]} {n[3:6]} {n[6:]}", f"{n[:3]}-{n[3:6]}-{n[6:]}"])


wiersze = []
start = date(2025, 1, 2)
for i in range(1, 401):
    nazwa, cena, waga, ean = random.choice(PRODUKTY)
    if i in (12, 77, 140, 250, 333):
        nazwa = nazwa.replace(" ", " ​", 1)       # niewidoczna spacja zerowej szerokości
    if i in (61, 199):
        nazwa = nazwa + "\t"                            # tabulator z kopiowania
    kolor = random.choice(list(KOLORY))
    zapis_koloru = random.choice(KOLORY[kolor]) if random.random() < 0.35 else kolor
    ilosc = random.choice([1, 1, 1, 2, 2, 3, 4])
    if i in (57, 233):
        ilosc = random.choice([2, 3]) * 10              # literówka: dopisane zero
    miasto = random.choice(list(MIASTA))
    kod = MIASTA[miasto] + f"{random.randint(1, 999):03d}"
    if miasto == "Warszawa" and i % 5 == 0:
        kod = kod[1:]                                   # Excel zgubił zero: 2-345 zamiast 02-345
    if i in (15, 98, 260):
        miasto = "Warszwa"                              # literówka
    if i in (44, 310):
        miasto = "Wroclw"
    if i % 23 == 0:
        miasto = miasto.encode("utf-8").decode("cp1252", errors="replace")   # krzaki
    klient = random.choice(KLIENCI)
    if i % 19 == 0:
        klient = klient.replace(" ", " ")          # twarda spacja z Worda
    if i % 29 == 0:
        klient = "  " + klient
    email = klient.lower().replace(" ", ".").replace(" ", ".").strip(" .") + f"{i}@poczta.pl"
    email = email.replace("..", ".")
    if i % 41 == 0:
        email = email.upper()
    if i in (33, 166, 299):
        email = email.replace("@", "@@") if i == 33 else email.replace(".pl", "") if i == 166 else email.replace(".", " ", 1)
    adres = f"{random.choice(ULICE)} {random.randint(1, 120)}/{random.randint(1, 40)}"[:30]   # system ucinał do 30 znaków
    d_zam = start + timedelta(days=i // 3)
    d_wys = d_zam + timedelta(days=random.randint(1, 5))
    d_wys_txt = d_wys.isoformat()
    if i in (20, 140, 205, 380):
        d_wys_txt = (d_zam - timedelta(days=random.randint(2, 6))).isoformat()   # wysyłka przed zamówieniem
    if i == 111:
        d_wys_txt = "2025-02-31"                        # data niemożliwa
    if i in (70, 222, 344):
        d_wys_txt = str((d_wys - date(1899, 12, 30)).days)                       # data jako liczba Excela
    ean_txt = ean if i % 14 else f"{int(ean) / 1e12:.5f}E+12".replace(".", ",")  # Excel: zapis naukowy
    cena_txt = f"{cena:.2f}".replace(".", ",") if i % 3 == 0 else f"{cena:.2f}"
    if i % 31 == 0:
        cena_txt = f"{cena:.2f} zł".replace(".", ",")
    nip = ""
    if random.random() < 0.3:
        nip = nip_poprawny()
        if i in (9, 150, 270) or random.random() < 0.03:
            nip = nip[:-1] + str((int(nip[-1]) + 3) % 10)                      # zła cyfra kontrolna
    nr = f"ZAM-{1000 + i}"
    if i in (180, 320):
        nr = f"ZAM-{1000 + i - 100}"                    # ten sam numer co inne zamówienie
    wiersze.append([nr, data_tekst(d_zam, i), d_wys_txt, nazwa, ean_txt, zapis_koloru, ilosc, cena_txt,
                    waga_tekst(waga * ilosc, i), klient, email, telefon(i), nip, adres, kod, miasto,
                    "Polska", random.choice(FAKTURA), random.choice(STATUSY),
                    random.choice(["", "", "", "brak", "NULL", "dostawa po 16", "prezent – bez paragonu"]), ""])
wiersze += [wiersze[40], wiersze[41], wiersze[120]]     # całe zdublowane wiersze

CEL.parent.mkdir(exist_ok=True)
with open(CEL, "w", newline="", encoding="utf-8") as f:
    w = csv.writer(f, delimiter=";")
    w.writerow(["nr_zamowienia", "data zamówienia", "data wysyłki", "produkt", "ean", "kolor", "ilość", "cena ",
                "waga", "klient", "email", "telefon", "NIP", "adres", "kod pocztowy", "miasto", "kraj",
                "faktura", "status", "uwagi", "notatka systemowa"])
    w.writerows(wiersze)
print("Zapisano", CEL, len(wiersze), "wierszy")
