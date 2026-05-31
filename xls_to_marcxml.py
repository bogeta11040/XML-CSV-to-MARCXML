#!/usr/bin/env python3
import sys
import os
import csv
import re
import subprocess
import tempfile
import unicodedata
import xml.etree.ElementTree as ET
from xml.dom import minidom
from datetime import datetime

# ─────────────────────────────────────────────────────────────────────────────
# MAPE NAZIVA KOLONA  (SVaki niz se ispituje redom; pobedjuje prvi pogodak)
# Srpski, Hrvatski, Slovenački, Engleski, Nemački
# ─────────────────────────────────────────────────────────────────────────────
KOLONE_MAPA = {
    "naslov": [
        "naziv knjige", "naziv", "naslov", "naslova", "knjiga",
        "title", "titel", "titre", "naslov knjige",
        "knjizni naslov", "knjiški naslov",
        "naslov dela", "naslov knjige/članka",
    ],
    "autor": [
        "autor", "autori", "pisac", "pisci", "author", "authors",
        "verfasser", "auteur", "pisatelj", "pisatelji",
        "ime autora", "prezime autora", "ime i prezime autora",
    ],
    "izdavac": [
        "izdavač", "izdavaci", "publisher", "verlag", "editeur",
        "izdanje", "nakladnik", "nakladnici", "založba",
        "izdavačka kuća", "izdavačka kuca",
    ],
    "godina": [
        "godina", "god.", "god", "year", "jahr", "année",
        "godina izdanja", "godina objave", "datum izdanja",
        "year of publication", "erscheinungsjahr",
        "leto izdaje", "leto",
    ],
    "isbn": [
        "isbn", "isbn-10", "isbn-13", "isbn10", "isbn13",
        "međunarodni standardni broj knjige",
    ],
    "mesto": [
        "mjesto izdanja", "mesto izdanja", "grad izdanja",
        "place of publication", "erscheinungsort",
        "kraj", "mjesto", "mesto", "grad",
        "kraj izdaje", "kraj izida",
    ],
    "jezik": [
        "jezik", "language", "sprache", "langue",
        "jezik knjige", "jezik teksta",
    ],
    "serija": [
        "serija", "series", "reihe", "série",
        "biblioteka", "edicija", "zbirka",
    ],
    "opis": [
        "opis", "napomena", "napomene", "beleška", "bilješka",
        "notes", "note", "bemerkung", "komentar", "sadržaj",
        "description", "abstract",
    ],
    "redni_broj": [
        "r. broj", "r.broj", "rbr", "rb", "r. br.", "r.br.",
        "redni broj", "broj", "no.", "no", "number", "#",
        "seq", "id",
    ],
}


def normalizuj(s: str) -> str:
    """Vrati lowercase"""
    s = unicodedata.normalize("NFKD", s)
    s = "".join(c for c in s if not unicodedata.combining(c))
    return s.strip().lower()


def mapuj_kolone(zaglavlja: list[str]) -> dict[str, int]:
    norm_mapa: dict[str, list[str]] = {
        k: [normalizuj(v) for v in vs]
        for k, vs in KOLONE_MAPA.items()
    }

    rezultat: dict[str, int] = {}
    for idx, zaglavlje in enumerate(zaglavlja):
        nz = normalizuj(zaglavlje)
        for polje, kandidati in norm_mapa.items():
            if polje in rezultat:
                continue
            if nz in kandidati:
                rezultat[polje] = idx
                break
            for k in kandidati:
                if k and (k in nz or nz in k):
                    rezultat[polje] = idx
                    break

    return rezultat


# ─────────────────────────────────────────────────────────────────────────────
# ČITANJE ULAZNIH FORMATA
# ─────────────────────────────────────────────────────────────────────────────

def xls_u_redove(putanja: str) -> list[list[str]]:
    ext = os.path.splitext(putanja)[1].lower()

    # --- pokušaj openpyxl ---
    try:
        import openpyxl
        wb = openpyxl.load_workbook(putanja, read_only=True, data_only=True)
        ws = wb.active
        redovi = []
        for red in ws.iter_rows(values_only=True):
            redovi.append([str(c) if c is not None else "" for c in red])
        wb.close()
        return redovi
    except Exception:
        pass

    # --- rezerva: LibreOffice konverzija u CSV ---
    try:
        with tempfile.TemporaryDirectory() as tmpdir:
            subprocess.run(
                ["libreoffice", "--headless", "--convert-to", "csv",
                 putanja, "--outdir", tmpdir],
                check=True, capture_output=True
            )
            csv_fajl = os.path.join(
                tmpdir,
                os.path.splitext(os.path.basename(putanja))[0] + ".csv"
            )
            return csv_u_redove(csv_fajl)
    except Exception as e:
        raise RuntimeError(
            f"Nije moguće otvoriti Excel fajl '{putanja}'. "
            f"Instaliraj openpyxl ili LibreOffice.\n{e}"
        )


def csv_u_redove(putanja: str) -> list[list[str]]:
    """Čita CSV i vraća listu redova"""
    redovi = []
    # Pokušaj nekoliko enkodinzga
    for enc in ("utf-8-sig", "utf-8", "cp1250", "iso-8859-2", "latin-1"):
        try:
            with open(putanja, newline="", encoding=enc) as f:
                reader = csv.reader(f)
                redovi = [row for row in reader]
            break
        except (UnicodeDecodeError, Exception):
            continue
    return redovi


def xml_u_redove(putanja: str) -> list[list[str]]:
    stablo = ET.parse(putanja)
    koreni = stablo.getroot()

    children = list(koreni)
    if not children:
        raise ValueError("XML fajl je prazan ili nema child elemenata u korenu.")

    svi_tagovi: list[str] = []
    seen: set[str] = set()
    for child in children:
        for el in child:
            if el.tag not in seen:
                svi_tagovi.append(el.tag)
                seen.add(el.tag)
        for attr in child.attrib:
            key = f"@{attr}"
            if key not in seen:
                svi_tagovi.append(key)
                seen.add(key)

    if not svi_tagovi:
        for child in children:
            if child.tag not in seen:
                svi_tagovi.append(child.tag)
                seen.add(child.tag)

    redovi: list[list[str]] = [svi_tagovi]  # Zaglavlje

    for child in children:
        red = []
        for tag in svi_tagovi:
            if tag.startswith("@"):
                red.append(child.attrib.get(tag[1:], ""))
            else:
                el = child.find(tag)
                if el is not None:
                    red.append((el.text or "").strip())
                else:
                    if child.tag == tag:
                        red.append((child.text or "").strip())
                    else:
                        red.append("")
        redovi.append(red)

    return redovi


# ─────────────────────────────────────────────────────────────────────────────
# DETEKCIJA ZAGLAVLJA
# ─────────────────────────────────────────────────────────────────────────────

def pronadji_zaglavlje(redovi: list[list[str]], max_pretraga: int = 10
                       ) -> tuple[int, dict[str, int]]:
    obavezni = {"naslov", "autor", "izdavac"}

    for i, red in enumerate(redovi[:max_pretraga]):
        mapa = mapuj_kolone(red)
        if len(obavezni & mapa.keys()) >= 2:
            return i, mapa

    return 0, mapuj_kolone(redovi[0]) if redovi else {}


# ─────────────────────────────────────────────────────────────────────────────
# IZGRADNJA MARCXML
# ─────────────────────────────────────────────────────────────────────────────

NS_MARC = "http://www.loc.gov/MARC21/slim"
SCHEMA  = "http://www.loc.gov/standards/marcxml/schema/MARC21slim.xsd"


def marc_kolekcija() -> ET.Element:
    ET.register_namespace("", NS_MARC)
    ET.register_namespace("xsi", "http://www.w3.org/2001/XMLSchema-instance")
    kolekcija = ET.Element(
        f"{{{NS_MARC}}}collection",
        attrib={
            "{http://www.w3.org/2001/XMLSchema-instance}schemaLocation":
                f"{NS_MARC} {SCHEMA}",
        }
    )
    return kolekcija


def dodaj_kontrolno_polje(zapis: ET.Element, tag: str, vrednost: str):
    cf = ET.SubElement(zapis, f"{{{NS_MARC}}}controlfield", attrib={"tag": tag})
    cf.text = vrednost


def dodaj_polje(zapis: ET.Element, tag: str, ind1: str, ind2: str,
                potpolja: list[tuple[str, str]]):
    df = ET.SubElement(
        zapis, f"{{{NS_MARC}}}datafield",
        attrib={"tag": tag, "ind1": ind1, "ind2": ind2}
    )
    for kod, vrednost in potpolja:
        if vrednost and vrednost.strip():
            sf = ET.SubElement(df, f"{{{NS_MARC}}}subfield", attrib={"code": kod})
            sf.text = vrednost.strip()
    return df


def red_u_marc_zapis(
    red: dict[str, str],
    ctrl_broj: int,
) -> ET.Element:

    naslov   = red.get("naslov", "").strip()
    autor    = red.get("autor", "").strip()
    izdavac  = red.get("izdavac", "").strip()
    godina   = red.get("godina", "").strip()
    isbn     = red.get("isbn", "").strip()
    mesto   = red.get("mesto", "").strip()
    jezik    = red.get("jezik", "").strip()
    serija   = red.get("serija", "").strip()
    opis     = red.get("opis", "").strip()

    if not naslov:
        return None

    zapis = ET.Element(f"{{{NS_MARC}}}record")

    # --- Leader (minimalni) ---
    leader = ET.SubElement(zapis, f"{{{NS_MARC}}}leader")
    leader.text = "00000nam a2200000 i 4500"

    # --- 001 Kontrolni broj ---
    dodaj_kontrolno_polje(zapis, "001", f"{ctrl_broj:06d}")

    # --- 003 Organizacija ---
    dodaj_kontrolno_polje(zapis, "003", "LOCAL")

    # --- 008 Fiksni podaci ---
    datum_danas = datetime.now().strftime("%y%m%d")
    god_marc = (godina[:4] if re.match(r"\d{4}", godina) else "    ")
    god_marc = god_marc.ljust(4)
    lang_kod = jezik[:3].lower() if len(jezik) >= 3 else "srp"
    f008 = f"{datum_danas}s{god_marc}    xx            000 0 {lang_kod} d"
    dodaj_kontrolno_polje(zapis, "008", f008[:40].ljust(40))

    # --- ISBN  (020) ---
    if isbn:
        isbn_cist = re.sub(r"[^0-9Xx]", "", isbn)
        dodaj_polje(zapis, "020", " ", " ", [("a", isbn_cist)])

    # --- Jezik (041) ---
    if jezik:
        dodaj_polje(zapis, "041", "0", " ", [("a", lang_kod)])

    # --- Autor (100 ili 700) ---
    autori = [a.strip() for a in re.split(r"[;|]", autor) if a.strip()]
    if autori:
        # Glavni autor → 100
        dodaj_polje(zapis, "100", "1", " ", [("a", autori[0])])
        # Ostali autori → 700
        for saauto in autori[1:]:
            dodaj_polje(zapis, "700", "1", " ", [("a", saauto)])

    # --- Naslov (245) ---
    # Ako ima autora, ind1=1 (jer je 100 prisutan), inače ind1=0
    ind1_245 = "1" if autori else "0"
    potpolja_245: list[tuple[str, str]] = [("a", naslov)]
    dodaj_polje(zapis, "245", ind1_245, "0", potpolja_245)

    # --- Impresum: mesto, izdavač, godina (264 / 260) ---
    if izdavac or mesto or godina:
        # Pokušaj izvući godinu iz polja "izdavač" ako nije zasebna kolona
        # (npr. "Kultura, 1949")
        god_izlucena = godina
        izd_cist     = izdavac
        if not god_izlucena and izdavac:
            m = re.search(r"\b(1[5-9]\d{2}|20[0-2]\d)\b", izdavac)
            if m:
                god_izlucena = m.group(1)
                izd_cist = re.sub(r",?\s*" + m.group(1), "", izdavac).strip().rstrip(",").strip()

        # Isto za mesto koje je uklopljeno u izdavača ("Kultura – Beograd, 1949")
        mj_cist = mesto
        if not mj_cist and izd_cist:
            separatori = r"[–\-—\/]"
            delovi = re.split(separatori, izd_cist)
            if len(delovi) >= 2:
                mj_cist  = delovi[-1].strip()
                izd_cist = separatori.join(delovi[:-1]).strip() if len(delovi) > 2 else delovi[0].strip()

        potpolja_264 = []
        if mj_cist:
            potpolja_264.append(("a", mj_cist + " :"))
        if izd_cist:
            potpolja_264.append(("b", izd_cist + ","))
        if god_izlucena:
            potpolja_264.append(("c", god_izlucena))
        if potpolja_264:
            dodaj_polje(zapis, "264", " ", "1", potpolja_264)

    # --- Serija (490) ---
    if serija:
        dodaj_polje(zapis, "490", "0", " ", [("a", serija)])

    # --- Opšta napomena (500) ---
    if opis:
        dodaj_polje(zapis, "500", " ", " ", [("a", opis)])

    return zapis


# ─────────────────────────────────────────────────────────────────────────────
# UGLAĐENI XML ISPIS
# ─────────────────────────────────────────────────────────────────────────────

def ugladi_xml(element: ET.Element) -> str:
    grubi = ET.tostring(element, encoding="unicode")
    dom = minidom.parseString(grubi)
    uglađen = dom.toprettyxml(indent="  ", encoding="UTF-8")
    # Ukloni prazne linije koje minidom ubacuje
    linije = [l for l in uglađen.decode("utf-8").splitlines() if l.strip()]
    return "\n".join(linije)


# ─────────────────────────────────────────────────────────────────────────────
# GLAVNI TOK
# ─────────────────────────────────────────────────────────────────────────────

def konvertuj(ulaz: str, izlaz: str | None = None) -> str:
    """
    Konvertuje ulazni fajl (csv/xls/xlsx/xml) u MARCXML.
    Vraća putanju do izlaznog fajla.
    """
    ext = os.path.splitext(ulaz)[1].lower()
    print(f"[+] Ulazni fajl : {ulaz}  (format: {ext})")

    # 1. Učitaj redove prema formatu
    if ext == ".xml":
        print("[+] Otkrivem XML → konvertujem u tabelu …")
        redovi = xml_u_redove(ulaz)
    elif ext == ".csv":
        print("[+] Otkrivem CSV → čitam …")
        redovi = csv_u_redove(ulaz)
    elif ext in (".xls", ".xlsx"):
        print("[+] Otkrivem Excel → čitam …")
        redovi = xls_u_redove(ulaz)
    else:
        raise ValueError(f"Nepodržani format: '{ext}'. Koristite csv/xls/xlsx/xml.")

    if not redovi:
        raise ValueError("Ulazni fajl je prazan.")

    # 2. Pronađi zaglavlje i mapu kolona
    idx_zaglavlja, mapa = pronadji_zaglavlje(redovi)
    print(f"[+] Zaglavlje pronađeno u redu {idx_zaglavlja}: {list(mapa.keys())}")

    if "naslov" not in mapa:
        print("[!] UPOZORENJE: Kolona naslova nije pronađena. "
              "Proveri nazive kolona u fajlu.")

    # 3. Gradi MARC kolekciju
    kolekcija = marc_kolekcija()
    zapisano = 0
    preskoceno = 0

    for br_reda, red in enumerate(redovi[idx_zaglavlja + 1:], start=1):
        # Preskoči prazne redove
        if not any(c.strip() for c in red):
            preskoceno += 1
            continue

        podaci: dict[str, str] = {}
        for polje, idx in mapa.items():
            if idx < len(red):
                podaci[polje] = red[idx]

        zapis = red_u_marc_zapis(podaci, br_reda)
        if zapis is not None:
            kolekcija.append(zapis)
            zapisano += 1
        else:
            preskoceno += 1

    print(f"[+] Konvertovano zapisa : {zapisano}")
    print(f"[+] Preskočeno (prazni) : {preskoceno}")

    # 4. Odredi izlaznu putanju
    if not izlaz:
        baza = os.path.splitext(os.path.basename(ulaz))[0]
        izlaz = os.path.join(os.path.dirname(ulaz) or ".", baza + "_marcxml.xml")

    # 5. Upiši fajl
    xml_sadrzaj = ugladi_xml(kolekcija)
    with open(izlaz, "w", encoding="utf-8") as f:
        f.write(xml_sadrzaj)

    print(f"[+] MARCXML fajl sačuvan: {izlaz}")
    return izlaz


# ─────────────────────────────────────────────────────────────────────────────
# CLI
# ─────────────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(
        description="Konvertor spiska knjiga (XLS/CSV/XML) → MARCXML"
    )
    parser.add_argument("ulaz", help="Ulazni fajl (.csv, .xls, .xlsx, .xml)")
    parser.add_argument(
        "--izlaz", "-o",
        help="Putanja izlaznog MARCXML fajla (opcionalno)",
        default=None,
    )
    args = parser.parse_args()

    try:
        izlazni = konvertuj(args.ulaz, args.izlaz)
        print(f"\n✓ Gotovo! MARCXML fajl: {izlazni}")
    except Exception as e:
        print(f"\n✗ Greška: {e}", file=sys.stderr)
        sys.exit(1)
