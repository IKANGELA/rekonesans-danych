// Rekonesans danych – obsługa strony.
// Python (Pyodide), biblioteki i kod są wczytywane z tego samego serwera,
// a polityka CSP w index.html blokuje połączenia z innymi adresami.

const PLIK_PRZYKLADU = "przyklady/przyklad_zamowienia.csv";
const BIBLIOTEKI = [
    "biblioteki/et_xmlfile-2.0.0-py3-none-any.whl",
    "biblioteki/openpyxl-3.1.5-py2.py3-none-any.whl",
];
const MAKS_ROZMIAR = 30 * 1024 * 1024;

const $ = (id) => document.getElementById(id);
const strefa = $("strefa"), wejscie = $("plik"), raport = $("raport");
let pyodide = null;

function ustawStatus(tekst, typ = "") {
    $("status").className = "status " + typ;
    $("status-tekst").textContent = tekst;
}

// Pomocnik do budowania elementów. Treść z pliku trafia na stronę wyłącznie jako tekst.
function el(tag, klasa, tekst) {
    const e = document.createElement(tag);
    if (klasa) e.className = klasa;
    if (tekst !== undefined) e.textContent = tekst;
    return e;
}

const liczba = new Intl.NumberFormat("pl-PL");

// --- 1. Start: Python w przeglądarce + nasz kod ---
async function uruchom() {
    try {
        pyodide = await loadPyodide({ indexURL: new URL("pyodide/", location.href).href });
        await pyodide.loadPackage(BIBLIOTEKI.map((b) => new URL(b, location.href).href));
        for (const modul of ["rekonesans.py", "oczyszczanie.py"]) {
            pyodide.FS.writeFile(modul, await fetch(modul).then((r) => r.text()));
        }
        pyodide.runPython(`
import json, base64, rekonesans, oczyszczanie

def przeanalizuj(sciezka, nazwa):
    with open(sciezka, "rb") as f:
        raport = rekonesans.analizuj(f.read(), nazwa)
    for k in raport["kolumny_raport"]:
        k["sugestie"] = oczyszczanie.sugestie(k)
    return json.dumps(raport, ensure_ascii=False)

def wyczysc_i_zapisz(sciezka, nazwa, ustawienia):
    with open(sciezka, "rb") as f:
        w = oczyszczanie.oczysc(f.read(), nazwa, json.loads(ustawienia))
    return json.dumps({
        "plik": base64.b64encode(w["plik"]).decode(), "nazwa_pliku": w["nazwa_pliku"],
        "dziennik_csv": base64.b64encode(w["dziennik_csv"]).decode(), "nazwa_dziennika": w["nazwa_dziennika"],
        "dziennik": w["dziennik"], "do_sprawdzenia": w["do_sprawdzenia"][:500],
        "podsumowanie": w["podsumowanie"],
    }, ensure_ascii=False)
`);
        strefa.classList.remove("zablokowana");
        wejscie.disabled = false;
        ustawStatus("Gotowe – wybierz plik CSV lub Excel.", "gotowy");
        const jest = await fetch(PLIK_PRZYKLADU, { method: "HEAD" }).then((r) => r.ok).catch(() => false);
        $("przyklad").hidden = !jest;
        $("przyklad").disabled = false;
    } catch (e) {
        console.error(e);
        ustawStatus("Nie udało się uruchomić narzędzia. Odśwież stronę lub spróbuj w innej przeglądarce.", "blad");
    }
}

// --- 2. Analiza pliku ---
async function analizuj(nazwaPliku, bajty) {
    ustawStatus(`Analizuję „${nazwaPliku}”…`);
    raport.classList.remove("widoczny");
    try {
        // plik zostaje w wirtualnej pamięci Pythona do czasu oczyszczenia albo „Wyczyść”
        pyodide.FS.writeFile("/tmp/plik", bajty);
        biezacyPlik = nazwaPliku;
        const funkcja = pyodide.globals.get("przeanalizuj");
        const wynik = JSON.parse(funkcja("/tmp/plik", nazwaPliku));
        funkcja.destroy();
        pokazRaport(wynik);
        ustawStatus("Gotowe. Przejrzyj raport, zaznacz poprawki i na dole wybierz „Wyczyść i zapisz”.", "gotowy");
    } catch (e) {
        console.error(e);
        usunPlik();
        const blad = String(e.message || e).match(/BladPliku: (.*)/);
        ustawStatus(blad ? blad[1] : "Nie udało się odczytać tego pliku. Czy to na pewno CSV lub Excel?", "blad");
    }
}

let biezacyPlik = null;
const adresy = [];

function usunPlik() {
    try { pyodide.FS.unlink("/tmp/plik"); } catch { /* plik mógł nie istnieć */ }
    biezacyPlik = null;
    while (adresy.length) URL.revokeObjectURL(adresy.pop());
}

// --- 3. Raport ---
function pokazRaport(r) {
    $("nazwa-pliku").textContent = r.plik || "Twój plik";

    const kafelki = [
        ["Wiersze", liczba.format(r.wiersze)],
        ["Kolumny", r.kolumny],
        ["Zdublowane wiersze", liczba.format(r.duplikaty), r.duplikaty > 0],
        ["Format", r.info.format],
        ["Kodowanie", r.info.kodowanie],
        ["Separator kolumn", r.info.separator],
        ["Separator dziesiętny", r.info.separator_dziesietny],
        ["Błędy kodowania", liczba.format(r.info.krzaki_w_pliku), r.info.krzaki_w_pliku > 0],
    ];
    $("kafelki").replaceChildren(...kafelki.map(([etykieta, wartosc, uwaga]) => {
        const k = el("div", "kafelek" + (uwaga ? " uwaga" : ""));
        k.append(el("span", "", etykieta), el("strong", "", String(wartosc)));
        return k;
    }));

    // uwagi do całego pliku: nagłówek, liczba kolumn, sprzeczne daty, kolumny puste i stałe
    const uwagi = [...r.uwagi_naglowka];
    if (r.zla_liczba_kolumn) uwagi.push(`${r.zla_liczba_kolumn} wierszy ma inną liczbę kolumn niż nagłówek.`);
    for (const k of r.kolejnosc_dat) {
        uwagi.push(`„${k.koniec}” jest wcześniej niż „${k.poczatek}” w ${k.ile} wierszach, np. ${k.przyklady.slice(0, 3).join(", ")}.`);
    }
    for (const k of r.kolumny_raport) {
        for (const u of k.uwagi) if (/pusta|ta sama wartość/.test(u)) uwagi.push(`Kolumna „${k.nazwa}”: ${u}`);
    }
    $("uwagi").replaceChildren(...uwagi.map((u) => el("li", "", u)));

    $("naglowek-tabeli").replaceChildren(...["Kolumna", "Typ", ...r.grupy_problemow].map((t) => el("th", "", t)));
    $("tabela").replaceChildren(...r.kolumny_raport.map((k, i) => {
        const tr = el("tr");
        const td = el("td");
        const a = el("a", "", k.nazwa);
        a.href = "#szczegoly";
        a.addEventListener("click", () => pokazZakladke(i));
        td.append(a);
        tr.append(td, el("td", "typ", k.typ));
        for (const g of r.grupy_problemow) {
            const n = k.grupy[g] || 0;
            const komorka = el("td");
            komorka.append(n ? el("span", "liczba-problemu", liczba.format(n)) : el("span", "zero", "—"));
            tr.append(komorka);
        }
        return tr;
    }));

    budujZakladki(r);

    // ustawienia całego pliku
    $("opt-duplikaty").checked = r.duplikaty > 0;
    $("opt-duplikaty-opis").textContent = `usuń zdublowane wiersze (${liczba.format(r.duplikaty)})`;
    $("opt-uwagi").checked = false;
    kolumnyZSugestiami = r.kolumny_raport.map((k, i) => (k.sugestie && k.sugestie.length ? i : null)).filter((i) => i !== null);
    $("wynik-czyszczenia").classList.remove("widoczny");
    raport.classList.add("widoczny");
}

let kolumnyZSugestiami = [];

// --- Oczyszczanie i zapis ---
function pobierz(b64, typ) {
    const bajty = Uint8Array.from(atob(b64), (c) => c.charCodeAt(0));
    const url = URL.createObjectURL(new Blob([bajty], { type: typ }));
    adresy.push(url);
    return url;
}

async function wyczyscIZapisz() {
    if (!biezacyPlik) return;
    const wybor = $("opt-format").value;
    const [format, separator] = wybor === "xlsx" ? ["xlsx", ";"] : ["csv", wybor.slice(3)];
    const kolumny = {};
    for (const i of kolumnyZSugestiami) kolumny[i] = [];
    for (const pole of document.querySelectorAll("input.popraw")) {
        if (pole.checked) kolumny[pole.dataset.kol].push(pole.dataset.op);
    }
    const ustawienia = {
        format, separator, kolumny,
        duplikaty: $("opt-duplikaty").checked, puste_kolumny: $("opt-puste").checked,
        stale_kolumny: $("opt-stale").checked, naglowki: $("opt-naglowki").checked,
        kolumna_uwag: $("opt-uwagi").checked, format_daty: $("opt-data").value,
        tak_nie: $("opt-taknie").value, telefon: $("opt-telefon").value,
    };
    ustawStatus("Czyszczę dane…");
    try {
        const funkcja = pyodide.globals.get("wyczysc_i_zapisz");
        const w = JSON.parse(funkcja("/tmp/plik", biezacyPlik, JSON.stringify(ustawienia)));
        funkcja.destroy();
        pokazWynikCzyszczenia(w, format);
        ustawStatus(`Gotowe – zapisano „${w.nazwa_pliku}”.`, "gotowy");
    } catch (e) {
        console.error(e);
        ustawStatus("Nie udało się oczyścić danych. Spróbuj odznaczyć część poprawek.", "blad");
    }
}

function pokazWynikCzyszczenia(w, format) {
    const typ = format === "xlsx" ? "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" : "text/csv";
    const link = $("pobierz-plik");
    link.href = pobierz(w.plik, typ);
    link.download = w.nazwa_pliku;
    const dz = $("pobierz-dziennik");
    dz.href = pobierz(w.dziennik_csv, "text/csv");
    dz.download = w.nazwa_dziennika;
    link.click();                                   // od razu zapisujemy plik

    const p = w.podsumowanie;
    const kafelki = [
        ["Wiersze", `${liczba.format(p.wiersze_przed)} → ${liczba.format(p.wiersze_po)}`],
        ["Kolumny", `${p.kolumny_przed} → ${p.kolumny_po}`],
        ["Zmienione komórki", liczba.format(p.zmienione_komorki)],
        ["Do sprawdzenia ręcznie", liczba.format(p.do_sprawdzenia), p.do_sprawdzenia > 0],
    ];
    $("kafelki-czyszczenia").replaceChildren(...kafelki.map(([e, v, u]) => {
        const k = el("div", "kafelek" + (u ? " uwaga" : ""));
        k.append(el("span", "", e), el("strong", "", v));
        return k;
    }));
    $("dziennik").replaceChildren(...w.dziennik.map((d) => {
        const tr = el("tr");
        tr.append(el("td", "", d.kolumna), el("td", "", d.operacja), el("td", "", liczba.format(d.komorki)),
                  el("td", "", d.przyklady.join("  |  ")));
        return tr;
    }));
    $("do-sprawdzenia").replaceChildren(...w.do_sprawdzenia.map((x) => {
        const tr = el("tr");
        tr.append(el("td", "", x.wiersz == null ? "kilka" : String(x.wiersz)), el("td", "", x.kolumna),
                  el("td", "", x.problem), el("td", "", x.wartosc));
        return tr;
    }));
    $("wynik-czyszczenia").classList.add("widoczny");
    $("wynik-czyszczenia").scrollIntoView({ behavior: "smooth", block: "start" });
}

$("zapisz").addEventListener("click", wyczyscIZapisz);
$("zaznacz-wszystko").addEventListener("click", () => document.querySelectorAll("input.popraw").forEach((p) => { p.checked = true; }));
$("odznacz-wszystko").addEventListener("click", () => document.querySelectorAll("input.popraw").forEach((p) => { p.checked = false; }));

// --- Zakładki kolumn (jak arkusze w Excelu) ---

// Problemy, które zawsze traktujemy poważnie – oznaczają utracone lub błędne dane
const KRYTYCZNE = ["kodowanie", "niemożliwe daty", "daty Excela", "błędne numery", "powtórzone identyfikatory",
                   "zapis naukowy", "odstające", "ucięte wartości"];
const POZIOMY = ["bez uwag", "drobne problemy", "średnie problemy", "poważne problemy"];

// Poziomy:
//   3 – poważne: dane utracone albo błędne (krzaki, niemożliwe daty, błędne numery, literówki w liczbach…)
//   2 – średnie: niespójności, które zafałszują analizę (warianty, formaty, jednostki) w ≥ 2% wierszy
//   1 – drobne: kosmetyka (spacje, NULL / brak) albo niespójności w pojedynczych wierszach
function poziomKolumny(k) {
    if (!k.suma_problemow) return 0;
    const p = k.problemy;
    if (KRYTYCZNE.some((x) => p[x])) return 3;
    const wariantyWiersze = k.warianty.reduce((s, g) => s + pozaNajczestszym(g.map((x) => x.ile)), 0);
    const formatyWiersze = k.formaty.length > 1 && k.typ !== "tekst" ? pozaNajczestszym(k.formaty.map((f) => f.ile)) : 0;
    const istotne = wariantyWiersze + formatyWiersze + (p["mieszane typy"] || 0) + (p["zgubione zera"] || 0)
        + (p["literówki"] || 0) + (p["niejednoznaczne daty"] || 0) + (p["daty poza zakresem"] || 0)
        + (p["niewidoczne znaki"] || 0) + Object.values(k.jednostki).reduce((s, z) => s + pozaNajczestszym(Object.values(z)), 0)
        + (k.logiczne && Object.keys(k.logiczne).length ? Object.values(k.logiczne).sort((a, b) => b - a).slice(2).reduce((s, x) => s + x, 0) : 0);
    return istotne / Math.max(1, k.wiersze) >= 0.02 ? 2 : 1;
}

let kolumnyRaportu = [];
let aktywnaZakladka = 0;
let sortowanie = "plik";

function budujZakladki(r) {
    kolumnyRaportu = r.kolumny_raport.map((k, i) => ({ k, i, poziom: poziomKolumny(k), panel: panelKolumny(k, i) }));
    $("panele").replaceChildren(...kolumnyRaportu.map((x) => x.panel));
    // na start otwieramy kolumnę z największymi problemami
    const najgorsza = [...kolumnyRaportu].sort((a, b) => b.poziom - a.poziom || b.k.suma_problemow - a.k.suma_problemow)[0];
    rysujPasekZakladek();
    pokazZakladke(najgorsza ? najgorsza.i : 0);
    const zPoprawkami = kolumnyRaportu.filter((x) => x.k.suma_problemow).length;
    $("podsumowanie-zakladek").textContent =
        `${zPoprawkami} z ${kolumnyRaportu.length} kolumn ma coś do sprawdzenia · otwarto kolumnę z największymi problemami`;
}

function rysujPasekZakladek() {
    const kolejnosc = [...kolumnyRaportu];
    if (sortowanie === "problemy") {
        kolejnosc.sort((a, b) => b.poziom - a.poziom || b.k.suma_problemow - a.k.suma_problemow || a.i - b.i);
    }
    $("zakladki").replaceChildren(...kolejnosc.map(({ k, i, poziom }) => {
        const z = el("button", `zakladka poziom-${poziom}`);
        z.type = "button";
        z.id = `zakladka-${i}`;
        z.setAttribute("role", "tab");
        z.setAttribute("aria-controls", `panel-${i}`);
        z.setAttribute("aria-selected", String(i === aktywnaZakladka));
        z.tabIndex = i === aktywnaZakladka ? 0 : -1;
        z.title = `${k.nazwa} – ${POZIOMY[poziom]}`;
        z.append(el("span", "nazwa-zakladki", k.nazwa),
                 el("span", "znacznik", k.suma_problemow ? liczba.format(k.suma_problemow) : "✓"));
        z.addEventListener("click", () => pokazZakladke(i));
        return z;
    }));
    for (const p of document.querySelectorAll(".sortowanie button")) {
        p.setAttribute("aria-pressed", String(p.dataset.sort === sortowanie));
    }
}

function pokazZakladke(i) {
    aktywnaZakladka = i;
    for (const { i: j, panel } of kolumnyRaportu) panel.hidden = j !== i;
    for (const z of document.querySelectorAll(".zakladka")) {
        const aktywna = z.id === `zakladka-${i}`;
        z.setAttribute("aria-selected", String(aktywna));
        z.tabIndex = aktywna ? 0 : -1;
        if (aktywna) z.scrollIntoView({ block: "nearest", inline: "nearest" });
    }
}

// klawiatura: strzałki przełączają zakładki, jak w arkuszu
$("zakladki").addEventListener("keydown", (e) => {
    if (!["ArrowLeft", "ArrowRight", "Home", "End"].includes(e.key)) return;
    const przyciski = [...document.querySelectorAll(".zakladka")];
    const teraz = przyciski.findIndex((z) => z.id === `zakladka-${aktywnaZakladka}`);
    const nastepny = e.key === "Home" ? 0 : e.key === "End" ? przyciski.length - 1
        : (teraz + (e.key === "ArrowRight" ? 1 : -1) + przyciski.length) % przyciski.length;
    przyciski[nastepny].click();
    przyciski[nastepny].focus();
    e.preventDefault();
});
for (const p of document.querySelectorAll(".sortowanie button")) {
    p.addEventListener("click", () => { sortowanie = p.dataset.sort; rysujPasekZakladek(); });
}

function chip(tekst, ile) {
    const c = el("span", "chip", tekst);
    if (ile !== undefined) c.append(el("b", "", "×" + liczba.format(ile)));
    return c;
}

function chipy(lista) {
    const d = el("div", "chipy");
    d.append(...lista);
    return d;
}

const lista = (wartosci) => chipy(wartosci.map((w) => chip(w)));

// Blok w panelu kolumny. waga > 0 = problem (z liczbą przy nagłówku), waga = -1 = informacja.
function blok(tytul, waga, ...zawartosc) {
    const b = el("div", "blok" + (waga > 0 ? " problem" : ""));
    const h = el("h4", "", tytul);
    if (waga > 0) h.append(el("span", "waga", liczba.format(waga)));
    b.append(h, ...zawartosc);
    b.dataset.waga = String(waga);
    return b;
}

// ile wierszy trzeba zmienić, żeby ujednolicić grupę wariantów (wszystkie poza najczęstszym)
const pozaNajczestszym = (ilosci) => ilosci.reduce((s, x) => s + x, 0) - Math.max(...ilosci);

function panelKolumny(k, i) {
    const panel = el("div", "panel");
    panel.id = `panel-${i}`;
    panel.setAttribute("role", "tabpanel");
    panel.setAttribute("aria-labelledby", `zakladka-${i}`);
    const p = k.problemy;
    const poziom = poziomKolumny(k);

    // --- nagłówek panelu
    const glowka = el("div", "glowka-panelu");
    glowka.append(el("h3", "nazwa-kolumny", k.nazwa), el("span", `stopien poziom-${poziom}`, POZIOMY[poziom]));
    const meta = [`${k.typ}`, `${liczba.format(k.unikalne)} różnych wartości`];
    if (k.puste.puste) meta.push(`puste komórki: ${liczba.format(k.puste.puste)}`);
    if (k.numer) meta.push(`rozpoznano: ${k.numer.rodzaj}`);
    if (k.identyfikator) meta.push("identyfikator");
    glowka.append(el("p", "etykieta", meta.join(" · ")));
    panel.append(glowka);

    // --- poprawki do wyboru: przy każdym problemie pole "popraw"
    if (k.sugestie && k.sugestie.length) {
        const b = el("div", "blok poprawki");
        b.append(el("h4", "", "Popraw w tej kolumnie"), ...k.sugestie.map((s) => {
            const etykieta = el("label", "opcja");
            const pole = el("input", "popraw");
            pole.type = "checkbox";
            pole.checked = s.domyslnie;
            pole.dataset.kol = String(i);
            pole.dataset.op = s.id;
            const tekst = el("span");
            tekst.append(el("b", "", "popraw: "), document.createTextNode(s.opis));
            etykieta.append(pole, tekst);
            return etykieta;
        }));
        panel.append(b);
    }

    const bloki = [];
    const dodaj = (...b) => bloki.push(blok(...b));

    // --- numery: NIP, PESEL, REGON, IBAN, e-mail, telefon, kod pocztowy
    if (k.numer) {
        const n = k.numer;
        const tresc = [el("p", "", `Poprawne: ${liczba.format(n.poprawne)} · błędne: ${liczba.format(n.bledne)}` +
            (["NIP", "PESEL", "REGON", "IBAN"].includes(n.rodzaj) ? " (sprawdzona cyfra kontrolna)" : ""))];
        if (n.przyklady.length) tresc.push(lista(n.przyklady));
        dodaj(`${n.rodzaj} – kontrola poprawności`, n.bledne || -1, ...tresc);
        if (n.zapisy.length > 1) {
            dodaj(`${n.rodzaj} – sposoby zapisu (9 = cyfra)`, p.formaty || -1, chipy(n.zapisy.map((z) => chip(z.wzorzec, z.ile))));
        }
    }
    if (k.powtorzone_id.length) {
        dodaj("Ten sam identyfikator w różnych wierszach", p["powtórzone identyfikatory"],
            el("p", "", `Identyfikatory użyte w wierszach o różnej treści (to nie są całe zdublowane wiersze): ${p["powtórzone identyfikatory"]}.`),
            lista(k.powtorzone_id));
    }

    // --- formaty i typy
    if (k.formaty.length && k.typ !== "tekst" && !k.numer) {
        dodaj(k.formaty.length > 1 ? "Kilka formatów w jednej kolumnie" : "Format",
            k.formaty.length > 1 ? pozaNajczestszym(k.formaty.map((f) => f.ile)) : -1,
            chipy(k.formaty.map((f) => chip(f.format, f.ile))));
    }
    if (k.obce_typy.ile) {
        dodaj("Wartości innego typu niż reszta kolumny", k.obce_typy.ile,
            el("p", "", `Wartości niepasujące do typu „${k.typ}”: ${liczba.format(k.obce_typy.ile)}.`), lista(k.obce_typy.przyklady));
    }
    if (p["zapis naukowy"]) {
        dodaj("Zapis naukowy – Excel zamienił długi numer na przybliżenie", p["zapis naukowy"],
            el("p", "", `Liczba takich wartości: ${p["zapis naukowy"]}. Ostatnie cyfry numeru mogły zostać bezpowrotnie utracone.`),
            lista(k.naukowe));
    }
    if (p["zgubione zera"]) {
        dodaj("Zgubione zera na początku", p["zgubione zera"],
            el("p", "", `Wartości krótsze niż reszta – wygląda na to, że zniknęło zero z przodu (np. 2-345 zamiast 02-345): ${p["zgubione zera"]}.`),
            lista(k.zgubione_zera));
    }
    if (p["ucięte wartości"]) {
        dodaj("Możliwe ucięte wartości", p["ucięte wartości"],
            el("p", "", `Wartości o długości dokładnie ${k.uciete.dlugosc} znaków – tyle, ile mieści pole w systemie, więc końcówka mogła zostać obcięta: ${p["ucięte wartości"]}.`),
            lista(k.uciete.przyklady));
    }

    // --- daty
    const dt = k.daty;
    if (p["niemożliwe daty"]) dodaj("Niemożliwe daty", p["niemożliwe daty"], el("p", "", `Daty, których nie ma w kalendarzu (np. 31 lutego): ${p["niemożliwe daty"]}.`), lista(dt.niemozliwe));
    if (p["daty poza zakresem"]) dodaj("Daty poza sensownym zakresem", p["daty poza zakresem"], el("p", "", "Przed 1950 rokiem albo ponad 10 lat w przyszłości – często „domyślna” data z systemu:"), lista(dt.poza_zakresem));
    if (p["daty Excela"]) dodaj("Daty zapisane jako liczby Excela", p["daty Excela"], el("p", "", `Liczby dni od 1900 roku – tak Excel przechowuje daty: ${p["daty Excela"]}.`), lista(dt.excel));
    if (p["niejednoznaczne daty"]) dodaj("Niejednoznaczne daty", p["niejednoznaczne daty"], el("p", "", "Nie da się ustalić, czy to dzień/miesiąc czy miesiąc/dzień – np. 03/04/2025:"), lista(dt.niejednoznaczne));
    if (dt.zakres) dodaj("Zakres dat", -1, el("p", "", `od ${dt.zakres[0]} do ${dt.zakres[1]}`));

    // --- warianty, literówki, tak/nie, jednostki
    if (k.warianty.length) {
        dodaj("Ta sama wartość zapisana na kilka sposobów",
            k.warianty.reduce((s, g) => s + pozaNajczestszym(g.map((x) => x.ile)), 0),
            ...k.warianty.map((g) => {
                const w = el("div", "grupa");
                w.append(...g.map((x) => chip(x.zapis, x.ile)));
                return w;
            }));
    }
    if (k.literowki.length) {
        dodaj("Możliwe literówki", k.literowki.reduce((s, l) => s + l.ile_rzadki, 0), ...k.literowki.map((l) => {
            const w = el("div", "grupa");
            w.append(chip(l.rzadki, l.ile_rzadki), el("span", "etykieta", "→ może chodziło o"), chip(l.czesty, l.ile_czesty));
            return w;
        }));
    }
    if (Object.keys(k.logiczne).length) {
        const ilosci = Object.values(k.logiczne).sort((a, b) => b - a);
        dodaj("Tak / nie zapisane na kilka sposobów", ilosci.slice(2).reduce((s, x) => s + x, 0) || 1,
            chipy(Object.entries(k.logiczne).map(([w, n]) => chip(w, n))));
    }
    if (Object.keys(k.jednostki).length) {
        dodaj("Jednostki zapisane na kilka sposobów",
            Object.values(k.jednostki).reduce((s, z) => s + pozaNajczestszym(Object.values(z)), 0),
            ...Object.entries(k.jednostki).map(([nazwa, zapisy]) => {
                const w = el("div", "grupa");
                w.append(el("span", "etykieta", nazwa + ":"), ...Object.entries(zapisy).map(([z, n]) => chip(z, n)));
                return w;
            }));
    }

    // --- spacje, znaki, kodowanie
    if (p.spacje) {
        const o = k.spacje;
        dodaj("Zbędne spacje", p.spacje,
            el("p", "", `Na początku lub końcu: ${o.na_brzegach} · podwójne w środku: ${o.podwojne} · twarde spacje: ${o.twarde}`),
            lista(o.przyklady));
    }
    if (p["niewidoczne znaki"]) {
        dodaj("Znaki niewidoczne", p["niewidoczne znaki"],
            el("p", "", `Wartości ze znakami, których nie widać na ekranie (tabulator, enter, spacja zerowej szerokości), zwykle z kopiowania ze stron lub dokumentów: ${p["niewidoczne znaki"]}.`),
            lista(k.niewidoczne.przyklady));
    }
    if (k.kodowanie.ile) {
        dodaj("Błędy kodowania polskich znaków", k.kodowanie.ile,
            el("p", "", `Wartości ze źle odczytanymi polskimi znakami: ${k.kodowanie.ile}.`), lista(k.kodowanie.przyklady));
    }

    // --- liczby
    if (k.statystyki) {
        const st = k.statystyki;
        const tresc = [el("p", "", `min ${st.min} · mediana ${st.mediana} · max ${st.max}` + (st.ujemne ? ` · ujemne: ${st.ujemne}` : ""))];
        if (k.odstajace.ile) {
            tresc.push(el("p", "", `Wartości nietypowo duże i oderwane od reszty – może to literówka (np. dopisane zero) albo inna jednostka: ${k.odstajace.ile}.`),
                       lista(k.odstajace.przyklady));
        }
        dodaj(k.odstajace.ile ? "Podejrzane wartości liczbowe" : "Liczby", k.odstajace.ile || -1, ...tresc);
    }

    // --- puste
    const pseudo = Object.entries(k.puste.pseudo);
    if (k.puste.puste || pseudo.length) {
        dodaj(pseudo.length ? "Wpisy udające brak danych" : "Puste wartości", p["ukryte puste"] || -1,
            el("p", "", `Puste komórki: ${k.puste.puste}` + (pseudo.length ? " · wpisy udające brak danych:" : "")),
            chipy(pseudo.map(([w, n]) => chip(w, n))));
    }

    // --- informacje
    if (k.uwagi.length) dodaj("Uwagi", -1, ...k.uwagi.map((u) => el("p", "", u)));
    if (!k.numer && (k.typ !== "tekst" || (k.liczba_wzorcow > 1 && k.liczba_wzorcow <= 30))) {
        dodaj("Wzorce wartości (9 = cyfra, a = litery)", -1, chipy(k.wzorce.map((w) => chip(w.wzorzec, w.ile))));
    }
    if (k.wszystkie_wartosci.length) {
        dodaj("Wszystkie wartości (alfabetycznie – łatwo wyłapać synonimy)", -1, chipy(k.wszystkie_wartosci.map((w) => chip(w.wartosc, w.ile))));
    } else if (k.najczestsze.length) {
        dodaj("Najczęstsze wartości", -1, chipy(k.najczestsze.map((w) => chip(w.wartosc, w.ile))));
    }

    // największe problemy na górze, informacje na końcu
    bloki.sort((a, b) => Number(b.dataset.waga) - Number(a.dataset.waga));
    const problemy = bloki.filter((b) => Number(b.dataset.waga) > 0);
    const informacje = bloki.filter((b) => Number(b.dataset.waga) <= 0);
    if (problemy.length) {
        problemy[0].classList.add("najwiekszy");
        panel.append(el("p", "sekcja-panelu", "Problemy – od największego"), ...problemy);
    } else {
        panel.append(el("p", "brak-problemow", "✓ W tej kolumnie nie znaleziono problemów."));
    }
    if (informacje.length) panel.append(el("p", "sekcja-panelu", "Informacje"), ...informacje);
    return panel;
}

// --- 4. Wybór pliku ---
async function wczytaj(plik) {
    if (!plik || !pyodide) return;
    if (plik.size > MAKS_ROZMIAR) {
        ustawStatus("Plik jest za duży (ponad 30 MB).", "blad");
        return;
    }
    analizuj(plik.name, new Uint8Array(await plik.arrayBuffer()));
}

wejscie.addEventListener("change", () => wczytaj(wejscie.files[0]));
strefa.addEventListener("dragover", (e) => { e.preventDefault(); strefa.classList.add("nad"); });
strefa.addEventListener("dragleave", () => strefa.classList.remove("nad"));
strefa.addEventListener("drop", (e) => { e.preventDefault(); strefa.classList.remove("nad"); wczytaj(e.dataTransfer.files[0]); });
$("przyklad").addEventListener("click", async () => {
    const r = await fetch(PLIK_PRZYKLADU);
    analizuj("przyklad_zamowienia.csv", new Uint8Array(await r.arrayBuffer()));
});
$("wyczysc").addEventListener("click", () => {
    raport.classList.remove("widoczny");
    for (const id of ["kafelki", "uwagi", "tabela", "zakladki", "panele", "dziennik", "do-sprawdzenia", "kafelki-czyszczenia"]) $(id).replaceChildren();
    $("wynik-czyszczenia").classList.remove("widoczny");
    usunPlik();
    wejscie.value = "";
    try { pyodide.runPython("import gc; gc.collect()"); } catch { /* nic */ }
    ustawStatus("Raport wyczyszczony. Dane usunięte z pamięci przeglądarki.", "gotowy");
});

// --- 5. Osadzenie na ikangela.pl: wysokość treści dla ramki (iframe) ---
new ResizeObserver(() => {
    window.parent.postMessage({ rekonesansWysokosc: Math.ceil(document.body.getBoundingClientRect().height) }, "*");
}).observe(document.body);

uruchom();
