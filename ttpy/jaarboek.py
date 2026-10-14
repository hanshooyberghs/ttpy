#!/usr/bin/env python3
"""Maak clubpagina's uit de VTTL-export en voeg ze in het vaste jaarboek in.

Gebruik in de bestaande ttpy-omgeving; vereist pypdf en lxml.
De bewerkbare DOCX wordt uitsluitend met --maak-vaste-docx aangemaakt,
en een bestaand DOCX-bestand wordt daarbij nooit overschreven.
"""

from __future__ import annotations

import argparse
import csv
import hashlib
import json
import os
import re
import shutil
import subprocess
import tempfile
import tomllib
import unicodedata
import zipfile
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path, PureWindowsPath

from lxml import etree as ET
from pypdf import PdfReader, PdfWriter
from pypdf.generic import DecodedStreamObject, DictionaryObject, NameObject


NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
W = "{" + NS["w"] + "}"
MARKER = "CLUBBESTUREN_INVOEGPUNT"
ROLES = (
    ("voorzitter", "Voorzitter"),
    ("secretaris", "Secretaris"),
    ("penningmeester", "Penningmeester"),
    ("interclubSeniors", "Interclub"),
    ("interclubJeugd", "Jeugd"),
)


def clean(value: str | None) -> str:
    value = (value or "").strip()
    return "" if value.casefold() in {"null", "nan", "none"} else value


def read_csv(path: Path) -> list[dict[str, str]]:
    # Zelfde bronformaat als ttpy; latin-1 behoudt letterlijke backslashes.
    with path.open(encoding="latin-1", newline="") as handle:
        return [{k: clean(v) for k, v in row.items() if k} for row in csv.DictReader(handle, delimiter=";")]


def unique(values):
    result = []
    seen = set()
    for value in values:
        value = clean(value)
        if value and value.casefold() not in seen:
            seen.add(value.casefold())
            result.append(value)
    return result


def postal_address(row, prefix="", suffix=""):
    def field(name):
        return row.get(prefix + name + suffix, "")
    street = " ".join(filter(None, [field("Straat"), field("Huisnummer")]))
    box = field("Busnummer") if not prefix else field("busnummer")
    if box:
        street += " bus " + box
    postal = field("Postnummer") if not prefix else field("postnummer")
    city = field("Gemeente") if not prefix else field("gemeente")
    locality = " ".join(filter(None, [postal, city]))
    return ", ".join(filter(None, [street, locality]))


def load_clubs(person_path: Path, club_path: Path):
    people = read_csv(person_path)
    clubs = [r for r in read_csv(club_path) if re.fullmatch(r"A\d{3}", r["AansluitingsNummer"]) and r["AansluitingsNummer"] != "A000" and r.get("Actief?", "1") == "1"]
    codes = [r["AansluitingsNummer"] for r in clubs]
    if not clubs or len(set(codes)) != len(codes):
        raise ValueError("Geen unieke actieve Antwerpse clubs in exportClub.csv.")
    if not people or "Club Funtie (0)" not in people[0]:
        raise ValueError("De persoons-export bevat geen verwachte clubfunctiekolommen.")
    for club in clubs:
        code = club["AansluitingsNummer"]
        club["_venues"] = []
        for index in range(5):
            suffix = f" ({index})"
            name = club.get("Lokaal naam" + suffix, "")
            address = postal_address(club, "Lokaal ", suffix)
            if name or address:
                club["_venues"].append({"naam": name, "adres": address})
        club["_contacts"] = []
        for role, title in ROLES:
            selected = [p for p in people if p.get("Club (0)") == code and role in {v for k, v in p.items() if k.startswith("Club Fun")}]
            seen = set()
            selected.sort(key=lambda p: (p["Naam"].casefold(), p["Voornaam"].casefold(), p["Lidnummer"]))
            for person in selected:
                if person["Lidnummer"] in seen:
                    continue
                seen.add(person["Lidnummer"])
                email_values = unique([person.get("Email"), person.get("Email (CC)")])
                # Behoud beide exportvelden; splits alleen lijsten met ; of ,.
                email_values = unique(part for value in email_values for part in re.split(r"[;,]", value))
                club["_contacts"].append({
                    "functie": role,
                    "titel": title,
                    "naam": " ".join(filter(None, [person["Naam"], person["Voornaam"]])),
                    "telefoon": person.get("GSM") or person.get("Telefoon") or "",
                    "emails": email_values,
                    "adres": postal_address(person) if role == "secretaris" else "",
                })
    clubs.sort(key=lambda c: c["AansluitingsNummer"])
    return clubs, len(people)


# Helvetica-breedtes in duizendsten van een punt; Type 1-fonts zijn onderdeel
# van de PDF-standaard. Er zijn geen browser, netwerk of extra fonts nodig.
WIDTHS = dict(zip(
    " !\"#$%&'()*+,-./0123456789:;<=>?@ABCDEFGHIJKLMNOPQRSTUVWXYZ[\\]^_`abcdefghijklmnopqrstuvwxyz{|}~",
    [278,278,355,556,556,889,667,191,333,333,389,584,278,333,278,278,
     556,556,556,556,556,556,556,556,556,556,278,278,584,584,584,556,1015,
     667,667,722,722,667,611,778,722,278,500,667,556,833,722,778,667,778,722,667,611,722,667,944,667,667,611,
     278,278,278,469,556,333,
     556,556,500,556,556,278,556,556,222,222,500,222,833,556,556,556,556,333,500,278,556,500,722,500,500,500,
     334,260,334,584]
))


def text_width(text, size, bold=False):
    total = 0
    for character in text:
        base = unicodedata.normalize("NFKD", character)[0]
        total += WIDTHS.get(character, WIDTHS.get(base, 1000))
    # Bold glyphs zijn soms breder; houd conservatief extra ruimte vrij.
    return total * size / 1000 * (1.10 if bold else 1)


def wrap(text, width, size, bold=False):
    lines = []
    current = ""
    for word in text.split():
        candidate = (current + " " + word).strip()
        if current and text_width(candidate, size, bold) > width:
            lines.append(current)
            current = ""
        while text_width(word, size, bold) > width:
            cut = 1
            while cut < len(word) and text_width(word[:cut + 1], size, bold) <= width:
                cut += 1
            if current:
                lines.append(current)
                current = ""
            lines.append(word[:cut])
            word = word[cut:]
        current = (current + " " + word).strip()
    if current:
        lines.append(current)
    return lines or [""]


def pdf_string(text):
    encoded = text.encode("cp1252", errors="strict")
    return b"(" + encoded.replace(b"\\", b"\\\\").replace(b"(", b"\\(").replace(b")", b"\\)") + b")"


class ClubPage:
    def __init__(self):
        self.commands = []
        self.lines = []

    def text(self, text, x, y, size=11, bold=False, color=(.12, .17, .23)):
        if not text:
            return
        if x + text_width(text, size, bold) > 550:
            raise ValueError(f"Tekst buiten de paginabreedte: {text!r}")
        r, g, b = color
        font = "F2" if bold else "F1"
        self.commands.append(f"BT {r} {g} {b} rg /{font} {size} Tf 1 0 0 1 {x:.2f} {y:.2f} Tm ".encode() + pdf_string(text) + b" Tj ET\n")
        self.lines.append((text, x, y, size))

    def rule(self, y):
        self.commands.append(f"0.78 0.82 0.86 RG 0.5 w 48 {y:.2f} m 547 {y:.2f} l S\n".encode())

    def paragraph(self, text, x, y, width, size=11, bold=False, leading=14):
        for line in wrap(text, width, size, bold):
            self.text(line, x, y, size, bold)
            y -= leading
        return y


def draw_club(club, season, factor=1.0):
    page = ClubPage()
    size = lambda value: round(value * factor, 2)
    page.text("CLUBBESTUREN ANTWERPEN", 48, 797, 9, True, (.30, .36, .43))
    page.text(season.replace("-", "–"), 480, 797, 9)
    y = page.paragraph(club["Naam"], 48, 762, 450, size(21), True, size(25))
    page.text(club["AansluitingsNummer"], 48, y - 1, 10, False, (.35, .40, .46))
    y -= 21
    page.rule(y)
    y -= 22
    page.text("Lokaal" if len(club["_venues"]) <= 1 else "Lokalen", 48, y, size(11), True)
    y -= size(17)
    if not club["_venues"]:
        y = page.paragraph("Niet vermeld in export", 48, y, 499, size(10), leading=size(13))
    for venue in club["_venues"]:
        value = ": ".join(filter(None, [venue["naam"], venue["adres"]]))
        y = page.paragraph(value, 48, y, 499, size(10.5), leading=size(14))
    y -= size(15)
    for contact in club["_contacts"]:
        page.text(contact["titel"], 48, y, size(11), True)
        y -= size(17)
        y = page.paragraph(contact["naam"], 48, y, 499, size(11), leading=size(14))
        if contact["adres"]:
            y = page.paragraph(contact["adres"], 48, y, 499, size(10), leading=size(13))
        if contact["telefoon"]:
            y = page.paragraph(contact["telefoon"], 48, y, 499, size(10), leading=size(13))
        for email in contact["emails"]:
            y = page.paragraph(email, 48, y, 499, size(10), leading=size(13))
        y -= size(14)
    if y < 66:
        if factor > .84:
            return draw_club(club, season, round(factor - .04, 2))
        raise ValueError(f"Club {club['AansluitingsNummer']} past niet op één pagina.")
    page.rule(48)
    page.text("Clubbesturen Antwerpen", 48, 32, 8, False, (.40, .44, .49))
    page.text(club["AansluitingsNummer"], 514, 32, 8, False, (.40, .44, .49))
    return page


def create_pdf(clubs, season, output):
    writer = PdfWriter()
    fonts = DictionaryObject()
    for label, face in [("F1", "Helvetica"), ("F2", "Helvetica-Bold")]:
        font = DictionaryObject({NameObject("/Type"): NameObject("/Font"), NameObject("/Subtype"): NameObject("/Type1"), NameObject("/BaseFont"): NameObject("/" + face), NameObject("/Encoding"): NameObject("/WinAnsiEncoding")})
        fonts[NameObject("/" + label)] = writer._add_object(font)
    layouts = []
    for index, club in enumerate(clubs):
        layout = draw_club(club, season)
        page = writer.add_blank_page(width=595.276, height=841.890)
        page[NameObject("/Resources")] = DictionaryObject({NameObject("/Font"): fonts})
        stream = DecodedStreamObject()
        stream.set_data(b"".join(layout.commands))
        page[NameObject("/Contents")] = writer._add_object(stream)
        writer.add_outline_item(club["AansluitingsNummer"] + " — " + club["Naam"], index)
        layouts.append(layout)
    writer.add_metadata({"/Title": f"Clubbesturen Antwerpen {season}", "/Author": "Provinciaal Comité Antwerpen"})
    temporary = output.with_suffix(".tmp.pdf")
    with temporary.open("wb") as handle:
        writer.write(handle)
    reader = PdfReader(temporary)
    if len(reader.pages) != len(clubs):
        raise ValueError("Onverwacht aantal clubpagina's.")
    for club, page in zip(clubs, reader.pages):
        text = page.extract_text()
        for expected in [club["AansluitingsNummer"], club["Naam"]] + [c["naam"] for c in club["_contacts"]] + [e for c in club["_contacts"] for e in c["emails"]]:
            if " ".join(expected.split()) not in " ".join(text.split()):
                raise ValueError(f"Tekstcontrole PDF mislukt voor {club['AansluitingsNummer']}: {expected}")
    temporary.replace(output)
    return layouts


def read_docx(source):
    with zipfile.ZipFile(source) as archive:
        return ET.fromstring(archive.read("word/document.xml"))


def paragraph_text(paragraph):
    return "".join(paragraph.itertext()) if paragraph.tag != W + "p" else "".join(paragraph.xpath(".//w:t/text()", namespaces=NS))


def marker_paragraph(text, title=False, first=False):
    paragraph = ET.Element(W + "p")
    properties = ET.SubElement(paragraph, W + "pPr")
    if title:
        ET.SubElement(properties, W + "pStyle").set(W + "val", "Heading1")
    if first:
        ET.SubElement(properties, W + "pageBreakBefore")
    run = ET.SubElement(paragraph, W + "r")
    ET.SubElement(run, W + "t").text = text
    return paragraph


def create_static_docx(source, output):
    if source.resolve() == output.resolve() or output.exists():
        raise FileExistsError(f"DOCX niet overschreven: {output}. Gebruik een andere --vaste-docx om opnieuw te beginnen.")
    root = read_docx(source)
    body = root.find("w:body", NS)
    children = list(body)
    start = [i for i, p in enumerate(children) if p.tag == W + "p" and paragraph_text(p).strip() == "Clubbesturen Antwerpen"]
    end = [i for i, p in enumerate(children) if p.tag == W + "p" and paragraph_text(p).strip() == "Adresgegevens nationale interclub"]
    if len(start) != 1 or len(end) != 1 or end[0] <= start[0]:
        raise ValueError("De clubsectie kan niet eenduidig worden afgebakend.")
    start, end = start[0], end[0]
    removed = children[start:end]
    if any(p.find(".//w:sectPr", NS) is not None for p in removed):
        raise ValueError("Clubsectie bevat een sectie-einde; eerst handmatig beoordelen.")
    before = [ET.tostring(p) for p in children[:start]]
    after = [ET.tostring(p) for p in children[end + 1:]]
    for p in removed:
        body.remove(p)
    placeholders = [
        marker_paragraph("Clubbesturen Antwerpen", title=True, first=True),
        marker_paragraph(MARKER),
        marker_paragraph("Deze pagina wordt bij de PDF-samenvoeging vervangen door de automatisch gegenereerde clubbesturen."),
    ]
    for offset, paragraph in enumerate(placeholders):
        body.insert(start + offset, paragraph)
    next_heading = children[end]
    ppr = next_heading.find("w:pPr", NS)
    if ppr is None:
        ppr = ET.Element(W + "pPr")
        next_heading.insert(0, ppr)
    if ppr.find("w:pageBreakBefore", NS) is None:
        page_break = ET.Element(W + "pageBreakBefore")
        preceding = [i for i, child in enumerate(ppr) if child.tag in {W + "pStyle", W + "keepNext", W + "keepLines"}]
        ppr.insert(max(preceding) + 1 if preceding else 0, page_break)
    final = list(body)
    assert [ET.tostring(p) for p in final[:start]] == before
    assert [ET.tostring(p) for p in final[start + len(placeholders) + 1:]] == after
    xml = ET.tostring(root, encoding="UTF-8", xml_declaration=True, standalone=True)
    with zipfile.ZipFile(source) as original, zipfile.ZipFile(output, "x", zipfile.ZIP_DEFLATED) as target:
        for item in original.infolist():
            target.writestr(item, xml if item.filename == "word/document.xml" else original.read(item.filename))
    # Alle andere onderdelen (afbeeldingen, stijlen, kopteksten enz.) blijven bytegelijk.
    with zipfile.ZipFile(source) as original, zipfile.ZipFile(output) as target:
        assert set(original.namelist()) == set(target.namelist())
        for filename in original.namelist():
            if filename != "word/document.xml":
                assert original.read(filename) == target.read(filename), filename
    return len(removed)


def merge_pdf(static_pdf, club_pdf, output):
    if output.resolve() in {static_pdf.resolve(), club_pdf.resolve()}:
        raise ValueError("De samengevoegde PDF mag geen invoerbestand overschrijven.")
    fixed = PdfReader(static_pdf)
    inserted = PdfReader(club_pdf)
    positions = [i for i, p in enumerate(fixed.pages) if MARKER in (p.extract_text() or "")]
    if len(positions) != 1:
        raise ValueError(f"Verwacht één invoegpagina met {MARKER}; gevonden: {len(positions)}.")
    position = positions[0]
    marker_text = fixed.pages[position].extract_text() or ""
    if "Adresgegevens nationale interclub" in marker_text:
        raise ValueError("De invoegpagina bevat ook vaste inhoud; behoud de pagina-einden in de DOCX.")
    writer = PdfWriter()
    for index, page in enumerate(fixed.pages):
        if index == position:
            for clubpage in inserted.pages:
                writer.add_page(clubpage)
        else:
            writer.add_page(page)
    writer.add_metadata({"/Title": "Jaarboek Provinciaal Comité Antwerpen"})
    temporary = output.with_suffix(".tmp.pdf")
    with temporary.open("wb") as handle:
        writer.write(handle)
    expected = len(fixed.pages) - 1 + len(inserted.pages)
    if len(PdfReader(temporary).pages) != expected:
        raise ValueError("Onverwacht aantal pagina's na samenvoegen.")
    temporary.replace(output)


def convert_static_docx(source):
    """Exporteer met LibreOffice, lokaal of met de Windows-installatie in WSL."""
    if not source.is_file():
        raise FileNotFoundError(f"Bewerkbare DOCX ontbreekt: {source}")
    candidates = [shutil.which("libreoffice"), shutil.which("soffice")]
    if os.name == "nt":
        for base in [os.environ.get("ProgramFiles", "C:/Program Files"), os.environ.get("ProgramFiles(x86)", "C:/Program Files (x86)")]:
            candidates.extend(str(Path(base) / "LibreOffice/program" / name) for name in ["soffice.com", "soffice.exe"])
    else:
        candidates.extend("/mnt/c/Program Files/LibreOffice/program/" + name for name in ["soffice.com", "soffice.exe"])
    executable = next((str(c) for c in candidates if c and Path(c).is_file()), None)
    if not executable:
        raise RuntimeError("Geen LibreOffice gevonden. Exporteer de bewerkbare DOCX naar PDF en geef die op met --vaste-pdf.")
    windows_from_wsl = executable.startswith("/mnt/") and executable.lower().endswith((".exe", ".com"))

    def office_path(path):
        path = path.resolve()
        if windows_from_wsl:
            match = re.fullmatch(r"/mnt/([a-z])/(.*)", str(path))
            if not match:
                raise ValueError(f"Windows LibreOffice kan dit pad niet rechtstreeks openen: {path}")
            return str(PureWindowsPath(match[1].upper() + ":/" + match[2]))
        return str(path)

    with tempfile.TemporaryDirectory(prefix=".jaarboek-conversie-", dir=source.parent) as directory:
        temporary = Path(directory)
        profile = temporary / "profiel"
        profile_uri = PureWindowsPath(office_path(profile)).as_uri() if windows_from_wsl or os.name == "nt" else profile.as_uri()
        command = [executable, "-env:UserInstallation=" + profile_uri, "--headless", "--convert-to", "pdf", "--outdir", office_path(temporary), office_path(source)]
        print("Vaste pagina's omzetten met LibreOffice…", flush=True)
        try:
            result = subprocess.run(command, capture_output=True, text=True, errors="replace", timeout=180, check=False)
        except (OSError, subprocess.TimeoutExpired) as exc:
            raise RuntimeError("DOCX-naar-PDF-conversie kon niet worden uitgevoerd. Exporteer de DOCX naar PDF en gebruik --vaste-pdf.") from exc
        converted = temporary / (source.stem + ".pdf")
        if result.returncode != 0 or not converted.is_file():
            detail = (result.stderr or result.stdout or "Geen PDF ontvangen").strip()[-1200:]
            raise RuntimeError("LibreOffice kon de vaste pagina's niet omzetten: " + detail)
        # Eerst inhoud controleren; een bestaande goede PDF blijft bij fouten behouden.
        reader = PdfReader(converted)
        markers = [p for p in reader.pages if MARKER in (p.extract_text() or "")]
        if len(markers) != 1:
            raise ValueError("De geconverteerde vaste PDF heeft niet precies één herkenbare invoegpagina.")
        output = source.with_suffix(".pdf")
        shutil.copyfile(converted, output)
    return output


def source_details(path):
    return {"bestand": str(path.resolve()), "gewijzigd": datetime.fromtimestamp(path.stat().st_mtime).astimezone().isoformat(), "sha256": hashlib.sha256(path.read_bytes()).hexdigest()}


def publication_manifest(clubs, person_count, season, person_path, club_path):
    return {
        "gemaakt": datetime.now().astimezone().isoformat(),
        "seizoen": season, "clubs": len(clubs), "personen_export": person_count,
        "bronnen": [source_details(person_path), source_details(club_path)],
        "regels": {
            "bron": "Export", "telefoon": "GSM, anders Telefoon",
            "email": "Email en Email (CC), identieke waarden eenmaal",
            "adres": "Alleen secretaris", "interclub": "Alle interclubSeniors-contacten",
            "nationaal": "Geen afzonderlijke nationale interclubcontacten uit jaarboek",
            "jeugd": "Uitsluitend interclubJeugd uit export", "pagina": "Eén pagina per club",
        },
        "clubgegevens": [
            {"clubcode": c["AansluitingsNummer"], "naam": c["Naam"],
             "lokalen": c["_venues"], "contacten": c["_contacts"]}
            for c in clubs
        ],
    }


def main(base=None):
    here = Path(base) if base is not None else Path.cwd()
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--export", type=Path, default=None, help="Map met exportPerson.csv en exportClub.csv, of pad naar exportPerson.csv")
    parser.add_argument("--bron", type=Path, default=here / "Jaarboek_PCA_zonder_clubbesturen.docx")
    parser.add_argument("--seizoen", default=None)
    parser.add_argument("--club-pdf", type=Path, default=here / "Clubbesturen_Antwerpen.pdf")
    parser.add_argument("--maak-vaste-docx", action="store_true")
    parser.add_argument("--vaste-docx", type=Path, default=here / "Jaarboek_PCA_zonder_clubbesturen.docx")
    parser.add_argument("--vaste-pdf", type=Path, help="PDF-export van de bewerkbare DOCX, inclusief invoegpagina")
    parser.add_argument("--volledig", action="store_true", help="Zet ook de bewerkbare DOCX om en maak het volledige jaarboek")
    parser.add_argument("--jaarboek-pdf", type=Path, default=here.parent / "DocumentenPubliek/Jaarboek_PCA.pdf")
    args = parser.parse_args()
    export = args.export or Path(os.environ.get("EXPORT_TT") or here.parent.parent / "Ledenadministratie/Exports")
    person_path = export / "exportPerson.csv" if export.is_dir() else export
    club_path = person_path.with_name("exportClub.csv")
    if args.maak_vaste_docx and args.bron == args.vaste_docx:
        args.bron = here / "Jaarboek_PCA.docx"
    if args.maak_vaste_docx and args.vaste_docx.exists():
        parser.error(f"Bewerkbare DOCX bestaat al en wordt niet overschreven: {args.vaste_docx}")
    season = args.seizoen
    if not season:
        paragraphs = read_docx(args.bron).xpath("//w:body/w:p", namespaces=NS)
        match = re.search(r"(20\d{2})\s*[-–]\s*(20\d{2})", " ".join(paragraph_text(p) for p in paragraphs[:10]))
        if not match:
            parser.error("Geen seizoen in brondocument; geef --seizoen op.")
        season = match[1] + "-" + match[2]
    clubs, person_count = load_clubs(person_path, club_path)
    create_pdf(clubs, season, args.club_pdf)
    print(f"Club-PDF: {args.club_pdf} ({len(clubs)} pagina's, één per club)")
    manifest = publication_manifest(clubs, person_count, season, person_path, club_path)
    if args.maak_vaste_docx:
        removed = create_static_docx(args.bron, args.vaste_docx)
        print(f"Bewerkbare DOCX: {args.vaste_docx} ({removed} clubparagrafen vervangen door één invoegpagina)")
    if args.volledig and not args.vaste_pdf:
        args.vaste_pdf = convert_static_docx(args.vaste_docx)
    if args.vaste_pdf:
        merge_pdf(args.vaste_pdf, args.club_pdf, args.jaarboek_pdf)
        print(f"Samengevoegd jaarboek: {args.jaarboek_pdf}")
    if not args.vaste_pdf:
        manifest["alleen_clubpagina"] = True
    args.club_pdf.with_suffix(".json").write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")


@dataclass(frozen=True)
class YearbookConfig:
    vaste_docx: Path
    club_pdf: Path
    jaarboek_pdf: Path
    export: Path | None = None
    vaste_pdf: Path | None = None
    seizoen: str | None = None


def load_config(path):
    """Lees jaarboekpaden relatief aan de config, niet aan de werkmap."""
    path = Path(path).resolve()
    with path.open("rb") as handle:
        data = tomllib.load(handle)
    if set(data) - {"jaarboek"}:
        raise ValueError("Onbekende secties in check-config: " + ", ".join(sorted(set(data) - {"jaarboek"})))
    if "jaarboek" not in data:
        return None
    section = data["jaarboek"]
    allowed = {"vaste_docx", "club_pdf", "jaarboek_pdf", "export", "vaste_pdf", "seizoen"}
    if not isinstance(section, dict) or set(section) - allowed:
        raise ValueError("Ongeldige of onbekende instellingen in [jaarboek].")
    values = {}
    for name in allowed:
        value = section.get(name)
        if value is None:
            if name in {"vaste_docx", "club_pdf", "jaarboek_pdf"}:
                raise ValueError(f"[jaarboek].{name} is verplicht.")
            continue
        if not isinstance(value, str) or not value.strip():
            raise ValueError(f"[jaarboek].{name} moet een niet-lege string zijn.")
        if name == "seizoen":
            if not re.fullmatch(r"20\d{2}-20\d{2}", value):
                raise ValueError("Gebruik seizoen = '2026-2027'.")
            values[name] = value
        else:
            values[name] = (path.parent / Path(value).expanduser()).resolve()
    config = YearbookConfig(**values)
    outputs = [config.club_pdf, config.club_pdf.with_suffix(".json"), config.jaarboek_pdf]
    inputs = [config.vaste_docx, config.vaste_pdf or config.vaste_docx.with_suffix(".pdf")]
    if config.club_pdf.suffix.lower() != ".pdf" or config.jaarboek_pdf.suffix.lower() != ".pdf":
        raise ValueError("club_pdf en jaarboek_pdf moeten PDF-paden zijn.")
    if len(set(outputs)) != len(outputs) or set(outputs) & set(inputs):
        raise ValueError("Jaarboekuitvoer mag geen invoer of andere uitvoer overschrijven.")
    return config


def update_yearbook(config, *, force=False):
    """Publiceer alleen wijzigingen; sla de vergelijkingsbasis pas na succes op."""
    print("\n--------------------\nJaarboek publiceren\n--------------------\n", flush=True)
    export = config.export or (Path(os.environ["EXPORT_TT"]) if os.environ.get("EXPORT_TT") else None)
    if export is None:
        raise ValueError("Stel EXPORT_TT of [jaarboek].export in.")
    person_path = export / "exportPerson.csv" if export.is_dir() else export
    club_path = person_path.with_name("exportClub.csv")
    protected = {person_path.resolve(), club_path.resolve()}
    if any(path.resolve() in protected for path in (
        config.club_pdf, config.club_pdf.with_suffix(".json"), config.jaarboek_pdf,
    )):
        raise ValueError("Jaarboekuitvoer mag geen exportbestand overschrijven.")
    clubs, person_count = load_clubs(person_path, club_path)
    season = config.seizoen
    if season is None:
        paragraphs = read_docx(config.vaste_docx).xpath("//w:body/w:p", namespaces=NS)
        match = re.search(r"(20\d{2})\s*[-–]\s*(20\d{2})", " ".join(paragraph_text(p) for p in paragraphs[:10]))
        if not match:
            raise ValueError("Geen seizoen in vaste DOCX; stel [jaarboek].seizoen in.")
        season = match[1] + "-" + match[2]
    manifest = publication_manifest(clubs, person_count, season, person_path, club_path)
    manifest_path = config.club_pdf.with_suffix(".json")
    previous = None
    if manifest_path.exists():
        previous = json.loads(manifest_path.read_text(encoding="utf-8"))
        if not isinstance(previous, dict) or not isinstance(previous.get("clubgegevens"), list):
            raise ValueError(f"Ongeldig jaarboek-bronbestand: {manifest_path}")
    if (not force and previous is not None
            and not previous.get("alleen_clubpagina")
            and previous["clubgegevens"] == manifest["clubgegevens"]
            and previous.get("seizoen") == season
            and config.club_pdf.is_file() and config.jaarboek_pdf.is_file()):
        print("Jaarboek: geen wijzigingen in gepubliceerde bestuurs- of clubgegevens.")
        return False
    print("Jaarboek: bestuurs-/clubgegevens gewijzigd, publicatie ontbreekt of handmatige generatie gevraagd.")
    static_pdf = config.vaste_pdf or convert_static_docx(config.vaste_docx)
    for output in (config.club_pdf, config.jaarboek_pdf):
        output.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix=".jaarboek-publicatie-", dir=config.club_pdf.parent) as directory:
        staged_clubs = Path(directory) / "clubs.pdf"
        create_pdf(clubs, season, staged_clubs)
        merge_pdf(static_pdf, staged_clubs, config.jaarboek_pdf)
        shutil.copyfile(staged_clubs, config.club_pdf)
    temporary = manifest_path.with_suffix(".tmp.json")
    temporary.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
    temporary.replace(manifest_path)
    print(f"Samengevoegd jaarboek: {config.jaarboek_pdf} ({len(clubs)} clubpagina's)")
    return True


def run():
    parser = argparse.ArgumentParser(description="Genereer het volledige jaarboek met de check-config.")
    parser.add_argument("--config", type=Path, default=Path(os.environ.get("TTPY_CHECK_CONFIG", "check.toml")))
    args = parser.parse_args()
    config = load_config(args.config)
    if config is None:
        parser.error("De config bevat geen [jaarboek]-sectie.")
    update_yearbook(config, force=True)


if __name__ == "__main__":
    run()
