"""Veilige, expliciete persoonsaanpassingen via de bestaande VTTL-formulieren."""

import argparse
import datetime as dt
import getpass
from html.parser import HTMLParser
import os
import re
import tomllib
import unicodedata
from urllib.parse import parse_qs, urljoin, urlsplit

import requests

from .vttl_download import VTTL_BASE_URL, login_session


class AanpassingFout(Exception):
    """Een veilige foutmelding zonder response bodies of authenticatiegegevens."""


# Publieke Nederlandse namen staan bewust los van technische formuliernamen.
TEXT = {
    "sportdatabank_id": ("sportDatabankId", 50),
    "voornaam": ("firstName", 50), "achternaam": ("lastName", 50),
    "straat": ("straat", 50), "huisnummer": ("huisNummer", 5),
    "busnummer": ("busNummer", 5), "postcode": ("postNummer", 4),
    "gemeente": ("gemeente", 50), "nationaliteit": ("nationaliteit", 2),
    "telefoon": ("telefoonNummer", 15), "fax": ("faxNummer", 15),
    "gsm": ("GSMNummer", 15), "telefoon_werk": ("telefoonNummerWerk", 15),
    "fax_werk": ("faxNummerWerk", 15), "geboortedatum": ("geboorteDatum", 10),
    "email": ("email", 40), "email_cc": ("emailCC", 200),
    "speelgerechtigd_tot": ("speelgerechtigdTot", 10),
}
BOOL = {
    "geen_matchen_heren": "geenMatchenHeren",
    "geen_matchen_dames": "geenMatchenDames",
    "begeleiden_sportende_senioren": "begeleidenSportendeSenioren",
    "jeugdsportcoordinator": "jeugdSportCoordinator",
    "bewegingsanimator": "bewegingsAnimator",
    "begeleider_multimove": "begeleiderMultimove",
    "coachlicentie": "coachLicentie",
    "onrechtmatige_transfer": "onrechtmatigeTransfer",
}
ENUM = {
    "provincie": ("provincie", ("Antwerpen", "Limburg", "Oost-Vlaanderen",
                              "Vlaams-Brabant", "West-Vlaanderen")),
    "geslacht": ("geslacht", ("man", "vrouw")),
    "vts_kwalificatie": ("vtsKwalificatie", ("geen", "aspirantInitiator", "initiator",
                                           "instructeurB", "trainerB", "trainerA")),
    "master_bachelor_lo": ("masterBachelorLO", ("geen", "regentLO", "licentiaatLO")),
    "diploma_scheidsrechter": ("diplomaScheidsrechter",
                              ("geen", "internationaal", "nationaal", "P1", "P2")),
}
GSPORT = ("AU", "FYSR", "FYSS", "VE", "VI", "PSY")
STATUTEN = ("administratiefLid", "competitiefLid", "superAfdeling",
            "dubbeleAansluiting", "dubbeleAansluitingSuper", "recreantReserve",
            "recreatiefLid", "vrijetijdsspeler")
NUMERIC_TEXT = {"sportdatabank_id", "huisnummer", "busnummer", "postcode",
                "telefoon", "fax", "gsm", "telefoon_werk", "fax_werk"}
PERSON_FIELDS = {v[0] for v in TEXT.values()} | set(BOOL.values()) | {
    v[0] for v in ENUM.values()} | {"gSportersCategorieen"}
PUBLIC = {v[0]: k for k, v in TEXT.items()} | {
    v: k for k, v in BOOL.items()} | {v[0]: k for k, v in ENUM.items()} | {
    "gSportersCategorieen": "gsport_categorieen", "statutenAsString": "statuut",
    "klassementHeren": "heren", "klassementDames": "dames", "startDatum": "begindatum"}


def normaliseer(value):
    return " ".join("".join(c for c in unicodedata.normalize("NFKD", value)
                           if not unicodedata.combining(c)).casefold().split())


def leeg(value):
    return value == "" or (type(value) is list and not value)


def datum(value, field):
    if type(value) is not str or not re.fullmatch(r"\d{2}/\d{2}/\d{4}", value):
        raise AanpassingFout(f"{field}: gebruik een datumstring DD/MM/YYYY.")
    try:
        dt.datetime.strptime(value, "%d/%m/%Y")
    except ValueError:
        raise AanpassingFout(f"{field}: ongeldige datum.") from None
    return value


def onbekend(data, allowed, context):
    if type(data) is not dict:
        raise AanpassingFout(f"{context}: verwacht een TOML-tabel.")
    if set(data) - set(allowed):
        # No echo of arbitrary keys: a misplaced secret must never reach output.
        raise AanpassingFout(f"{context}: onbekende veldnaam of sleutel.")


def parse_config(data, today=None):
    onbekend(data, {"leden"}, "config")
    leden = data.get("leden")
    if type(leden) is not list:
        raise AanpassingFout("config: verwacht leden = [] of één of meer [[leden]].")
    result, identities = [], set()
    for index, raw in enumerate(leden, 1):
        context = f"lid {index}"
        onbekend(raw, {"naam", "lidnummer", "wijzigingen", "statuut", "klassement"}, context)
        name = raw.get("naam")
        if type(name) is not str or not normaliseer(name):
            raise AanpassingFout(f"{context}: identificatienaam is verplicht.")
        number = raw.get("lidnummer")
        if number is not None and not leeg(number):
            if type(number) not in (str, int) or not re.fullmatch(r"\d{1,6}", str(number)):
                raise AanpassingFout(f"{context}: lidnummer moet 1–6 cijfers bevatten.")
            number = str(number)
        else:
            number = None
        identity = ("nummer", int(number)) if number is not None else ("naam", normaliseer(name))
        if identity in identities:
            raise AanpassingFout(f"{context}: dubbel lid.")
        identities.add(identity)
        changes = raw.get("wijzigingen", {})
        onbekend(changes, set(TEXT) | set(BOOL) | set(ENUM) | {"gsport_categorieen"}, context)
        parsed = {}
        for field, value in changes.items():
            if leeg(value):
                continue
            if field in BOOL:
                if type(value) is not bool:
                    raise AanpassingFout(f"{context}/{field}: verwacht true of false.")
                parsed[BOOL[field]] = value
            elif field in ENUM:
                target, options = ENUM[field]
                if type(value) is not str or value not in options:
                    raise AanpassingFout(f"{context}/{field}: ongeldige keuzewaarde.")
                parsed[target] = value
            elif field == "gsport_categorieen":
                if (type(value) is not list or any(type(x) is not str or x not in GSPORT for x in value)
                        or len(value) != len(set(value))):
                    raise AanpassingFout(f"{context}/{field}: ongeldige of dubbele categorie.")
                parsed["gSportersCategorieen"] = sorted(value)
            else:
                target, limit = TEXT[field]
                if type(value) is int and field in NUMERIC_TEXT:
                    value = str(value)
                if type(value) is not str or len(value) > limit:
                    raise AanpassingFout(f"{context}/{field}: verwacht tekst van maximaal {limit} tekens.")
                if field in {"geboortedatum", "speelgerechtigd_tot"}:
                    datum(value, field)
                parsed[target] = value
        statute = raw.get("statuut")
        if statute is not None and not leeg(statute):
            if type(statute) is not str or statute not in STATUTEN:
                raise AanpassingFout(f"{context}/statuut: ongeldige keuzewaarde.")
        else:
            statute = None
        rank = raw.get("klassement", {})
        onbekend(rank, {"heren", "dames", "begindatum"}, context + "/klassement")
        rank = {k: v for k, v in rank.items() if not leeg(v)}
        ranking, default_date = {}, False
        if rank:
            if not ({"heren", "dames"} & set(rank)):
                raise AanpassingFout(f"{context}: klassement vereist heren en/of dames.")
            for field in ("heren", "dames"):
                if field in rank:
                    value = rank[field]
                    if type(value) is not str or not value.strip() or len(value) > 4:
                        raise AanpassingFout(f"{context}/klassement/{field}: verwacht 1–4 tekens.")
                    ranking["klassement" + field.capitalize()] = value
            default_date = "begindatum" not in rank
            ranking["startDatum"] = datum(
                rank.get("begindatum", (today or dt.date.today()).strftime("%d/%m/%Y")), "begindatum")
        result.append({"naam": name, "lidnummer": number, "person": parsed,
                       "statute": statute, "rank": ranking, "default_date": default_date})
    return result


def load_config(path):
    try:
        with open(path, "rb") as stream:
            data = tomllib.load(stream)
    except (OSError, tomllib.TOMLDecodeError):
        raise AanpassingFout("Config niet leesbaar of ongeldig TOML (ook dubbele keys zijn verboden).") from None
    return parse_config(data)


class Form:
    def __init__(self, action):
        self.action = action
        self.controls = {}

    def add(self, name, kind, value="", **meta):
        if name in self.controls:
            if urlsplit(self.action).path == "/personList.jspa":
                return  # Read-only lookup; bulk activation controls are never submitted.
            if kind in {"image", "submit", "button", "reset"} and self.controls[name]["kind"] == kind:
                return
            if kind != "radio" or self.controls[name]["kind"] != "radio":
                raise AanpassingFout("Onverwachte dubbele formuliervelden; veilig gestopt.")
            self.controls[name]["options"].append(value)
            if meta.get("checked"):
                if self.controls[name]["value"]:
                    raise AanpassingFout("Dubbele geselecteerde radiowaarde.")
                self.controls[name]["value"] = value
            return
        self.controls[name] = dict(kind=kind, value=value, **meta)
        if kind == "radio":
            self.controls[name]["options"] = [value]
            self.controls[name]["value"] = value if meta.get("checked") else ""

    def values(self):
        return {name: sorted(control["value"]) if control["kind"] == "multiple" else control["value"]
                for name, control in self.controls.items()
                if control["kind"] not in {"hidden", "image", "submit", "button", "reset"}}

    def validate(self, changes):
        for name, value in changes.items():
            control = self.controls.get(name)
            if not control or control.get("disabled") or name not in self.values():
                raise AanpassingFout(f"Veld {PUBLIC.get(name, 'onbekend')} niet bewerkbaar op dit formulier.")
            if control["kind"] == "checkbox":
                if type(value) is not bool:
                    raise AanpassingFout("Onjuist type voor checkbox.")
            elif control["kind"] in {"select", "multiple", "radio"}:
                values = value if type(value) is list else [value]
                if any(x not in control["options"] for x in values):
                    raise AanpassingFout(f"Ongeldige online keuze voor {PUBLIC.get(name, 'veld')}.")
            elif type(value) is not str:
                raise AanpassingFout("Onjuist formulierveldtype.")
            if type(value) is str and control.get("maxlength") and len(value) > control["maxlength"]:
                raise AanpassingFout(f"Online maximumlengte overschreden voor {PUBLIC.get(name, 'veld')}.")

    def payload(self, changes):
        self.validate(changes)
        data = []
        for name, control in self.controls.items():
            kind = control["kind"]
            if kind in {"image", "submit", "button", "reset"} or control.get("disabled"):
                continue
            value = changes.get(name, control["value"])
            if kind == "checkbox":
                if value:
                    data.append((name, control["checked_value"]))
            elif kind == "multiple":
                data.extend((name, x) for x in value)
            elif kind != "radio" or value:
                data.append((name, value))
        data.extend([("save.x", "50"), ("save.y", "10")])
        return data


class Page(HTMLParser):
    """Form-state en alleen zichtbare tabelrijen; scripts en secrets blijven buiten snapshots."""

    def __init__(self, html):
        super().__init__(convert_charrefs=True)
        self.forms, self.links, self.tables = [], [], []
        self.form = self.select = self.option = None
        self.stack = []
        self.ignore = 0
        self.feed(html)
        self.close()

    def handle_starttag(self, tag, attrs):
        a = dict(attrs)
        if tag in {"script", "style"}:
            self.ignore += 1
        if tag == "form":
            self.form = Form(a.get("action", ""))
            self.forms.append(self.form)
        if tag == "input" and self.form and a.get("name"):
            kind, value = a.get("type", "text").lower(), a.get("value", "")
            meta = {"disabled": "disabled" in a}
            if "maxlength" in a:
                try:
                    meta["maxlength"] = int(a["maxlength"])
                except ValueError:
                    raise AanpassingFout("Ongeldige formulierlengte.") from None
            if kind == "checkbox":
                meta["checked_value"] = a.get("value", "on")
                value = "checked" in a
            if kind == "radio":
                meta["checked"] = "checked" in a
            self.form.add(a["name"], kind, value, **meta)
        if tag == "select" and self.form and a.get("name"):
            self.select = {"name": a["name"], "multiple": "multiple" in a,
                           "disabled": "disabled" in a, "options": [], "selected": []}
        if tag == "option" and self.select:
            self.option = {"value": a.get("value"), "selected": "selected" in a, "text": []}
        if tag == "a":
            link = {"href": a.get("href", ""), "text": []}
            self.links.append(link)
            if self.stack and self.stack[-1].get("row") is not None:
                self.stack[-1]["row"]["links"].append(link)
            self.anchor = link
        if tag == "table":
            self.stack.append({"headers": [], "rows": [], "row": None, "cell": None, "header": False})
        if tag == "tr" and self.stack:
            self.stack[-1]["row"] = {"cells": [], "links": []}
        if tag in {"td", "th"} and self.stack and self.stack[-1]["row"] is not None:
            table = self.stack[-1]
            table["cell"] = []
            table["header"] = tag == "th"
            table["row"]["cells"].append(table["cell"])
            if tag == "th":
                table["headers"].append(table["cell"])

    def handle_endtag(self, tag):
        if tag in {"script", "style"}:
            self.ignore = max(0, self.ignore - 1)
        if tag == "form":
            self.form = None
        if tag == "option" and self.option:
            value = self.option["value"]
            if value is None:
                value = "".join(self.option["text"]).strip()
            self.select["options"].append(value)
            if self.option["selected"]:
                self.select["selected"].append(value)
            self.option = None
        if tag == "select" and self.select:
            c = self.select
            selected = c["selected"]
            if not c["multiple"] and len(selected) > 1:
                raise AanpassingFout("Dubbele geselecteerde dropdownwaarde.")
            value = selected if c["multiple"] else (
                selected[0] if selected else (c["options"][0] if c["options"] else ""))
            self.form.add(c["name"], "multiple" if c["multiple"] else "select", value,
                          options=c["options"], disabled=c["disabled"])
            self.select = None
        if tag == "a":
            self.anchor = None
        if tag in {"td", "th"} and self.stack:
            self.stack[-1]["cell"] = None
        if tag == "tr" and self.stack:
            table = self.stack[-1]
            if table["row"] is not None:
                table["rows"].append(table["row"])
            table["row"] = None
        if tag == "table" and self.stack:
            self.tables.append(self.stack.pop())

    def handle_data(self, data):
        if self.ignore:
            return
        if self.option:
            self.option["text"].append(data)
        if getattr(self, "anchor", None):
            self.anchor["text"].append(data)
        if self.stack and self.stack[-1]["cell"] is not None:
            self.stack[-1]["cell"].append(data)

    def form_for(self, route):
        forms = [f for f in self.forms if urlsplit(f.action).path.split(";")[0] == "/" + route]
        if len(forms) != 1:
            raise AanpassingFout("Doelformulier ontbreekt of is dubbel; controleer login en toegang.")
        return forms[0]

    def grids(self):
        result = {}
        for table in self.tables:
            headers = tuple(" ".join("".join(x).split()) for x in table["headers"])
            if not headers:
                continue
            rows = tuple(tuple(" ".join("".join(cell).split()) for cell in row["cells"])
                         for row in table["rows"])
            if headers in result:
                raise AanpassingFout("Dubbele relatietabel; veilig gestopt.")
            result[headers] = rows
        return result


def relevant_params(href):
    # Some JSP links contain template fragments and duplicate query parameters.
    values = {}
    for name, value in re.findall(r"[?&](dispatch|idPerson|id)=([^&#?]*)", href):
        values[name] = value
    return values


class Client:
    def __init__(self, session):
        self.session = session

    def get(self, route, params=None):
        try:
            response = self.session.get(VTTL_BASE_URL + "/" + route, params=params, timeout=30)
            response.raise_for_status()
        except requests.RequestException:
            raise AanpassingFout("Online lezen mislukt; geen response of credentials gelogd.") from None
        return Page(response.text)

    def resolve(self, item):
        # The portal only searches parts of surnames, not "surname firstname",
        # and does not guarantee accent normalization. Scan all result pages for
        # name-only lookup rather than silently excluding normalized duplicates.
        params = {"filterActief": "AAAAActief", "filterProvincie": "AAAAAGeenProvincie",
                  "filterNaamPersoon": "",
                  "filterLidnummer": item["lidnummer"] or ""}
        pages, seen_pages, ids = [1], set(), set()
        queued, seen_rows, row_numbers = {1}, set(), {}
        while pages:
            number = pages.pop(0)
            if number in seen_pages:
                continue
            seen_pages.add(number)
            page = self.get("personList.jspa", dict(params, nextPage=str(number)))
            page.form_for("personList.jspa")
            result_tables = []
            for table in page.tables:
                headers = [" ".join("".join(x).split()) for x in table["headers"]]
                if "Naam" in headers and "Aansl.Nr." in headers:
                    result_tables.append(table)
                    for row in table["rows"]:
                        row_ids = {relevant_params(x["href"]).get("id") for x in row["links"]
                                   if "person.jspa" in x["href"]
                                   and relevant_params(x["href"]).get("dispatch") == "view"}
                        if not row_ids:
                            continue
                        if len(row_ids) != 1 or not next(iter(row_ids), "").isdigit():
                            raise AanpassingFout("Onleesbare persoonsrij; veilig gestopt.")
                        person_id = next(iter(row_ids))
                        if person_id in seen_rows:
                            raise AanpassingFout("Dubbele persoonsrij of veranderde paginatie; veilig gestopt.")
                        seen_rows.add(person_id)
                        if len(row["cells"]) < len(headers):
                            raise AanpassingFout("Onvolledige persoonsrij.")
                        actual = normaliseer("".join(row["cells"][headers.index("Naam")]))
                        number_text = "".join(row["cells"][headers.index("Aansl.Nr.")]).strip()
                        if actual == normaliseer(item["naam"]):
                            if person_id in row_numbers and row_numbers[person_id] != number_text:
                                raise AanpassingFout("Tegenstrijdige dubbele persoonsrij.")
                            row_numbers[person_id] = number_text
                            ids.add(person_id)
            if len(result_tables) != 1:
                raise AanpassingFout("Ledenoverzicht ontbreekt of is dubbel.")
            for link in page.links:
                href = link["href"]
                if "person.jspa" in href and relevant_params(href).get("dispatch") == "view":
                    person_id = relevant_params(href).get("id", "")
                    if not person_id.isdigit():
                        raise AanpassingFout("Onleesbare persoonslink; veilig gestopt.")
                if "personList.jspa?" in href:
                    query = parse_qs(urlsplit(href).query)
                    if "nextPage" in query:
                        value = query["nextPage"][-1]
                        if not value.isdigit() or int(value) < 1:
                            raise AanpassingFout("Onleesbare paginatie; veilig gestopt.")
                        if int(value) not in queued:
                            pages.append(int(value))
                            queued.add(int(value))
            if len(seen_pages) > 10000:
                raise AanpassingFout("Paginatie niet volledig aantoonbaar; veilig gestopt.")
        matches = []
        for person_id in sorted(ids):
            form = self.person(person_id)
            values = form.values()
            actual = normaliseer(values["lastName"] + " " + values["firstName"])
            if actual != normaliseer(item["naam"]):
                continue
            if item["lidnummer"] is not None:
                # The membership number belongs to the search-result row, not hidden person id.
                value = row_numbers.get(person_id, "")
                found = value.isdigit() and int(value) == int(item["lidnummer"])
                if not found:
                    raise AanpassingFout("Lidnummer kon niet onafhankelijk worden bevestigd.")
            matches.append(person_id)
        if len(matches) != 1:
            raise AanpassingFout("Naam/lidnummer ontbreekt of is niet eenduidig (exacte naamcontrole verplicht).")
        return matches[0]

    def person(self, person_id):
        form = self.get("person.jspa", {"dispatch": "update", "id": person_id}).form_for("personSave.jspa")
        self.identity(form, "id", person_id)
        if form.controls.get("actionOnPerson", {}).get("value") != "update":
            raise AanpassingFout("Persoonsformulier staat niet in update-modus.")
        if not PERSON_FIELDS <= set(form.values()):
            raise AanpassingFout("Persoonsformulier is onvolledig; veilig gestopt.")
        return form

    @staticmethod
    def identity(form, key, person_id):
        if form.controls.get(key, {}).get("value") != person_id:
            raise AanpassingFout("Doelidentiteit van formulier klopt niet.")

    def statute(self, person_id):
        form = self.get("statuut.jspa", {"dispatch": "update", "idPerson": person_id}).form_for("statuutSave.jspa")
        self.identity(form, "idPerson", person_id)
        if form.controls.get("actionOnStatuut", {}).get("value") != "update":
            raise AanpassingFout("Statuutformulier staat niet in update-modus.")
        if "statutenAsString" not in form.values():
            raise AanpassingFout("Statuutformulier is onvolledig.")
        return form

    def relations(self, person_id):
        page = self.get("person.jspa", {"dispatch": "view", "id": person_id})
        form = page.form_for("person.jspa")
        self.identity(form, "id", person_id)
        grids = page.grids()
        if (not any("Klassement heren" in h for h in grids)
                or not any("Statuut" in h for h in grids)
                or not any("Naam club" in h for h in grids)):
            raise AanpassingFout("Relatieoverzicht onvolledig; afwezigheid niet aantoonbaar.")
        records = {}
        for link in page.links:
            if "klassement.jspa" not in link["href"]:
                continue
            params = relevant_params(link["href"])
            if params.get("idPerson") != person_id or params.get("dispatch") not in {"view", "update", "delete"}:
                raise AanpassingFout("Onleesbare klassementlink; afwezigheid niet aantoonbaar.")
            if "id" not in params:
                if params["dispatch"] != "update":
                    raise AanpassingFout("Onvolledige klassementlink.")
                continue
            if not params["id"].isdigit():
                raise AanpassingFout("Onleesbaar klassementrecord.")
            records[params["id"]] = params
        rank_headers = [h for h in grids if "Klassement heren" in h]
        row_record_ids = set()
        for table in page.tables:
            headers = [" ".join("".join(x).split()) for x in table["headers"]]
            if "Klassement heren" not in headers:
                continue
            for row in table["rows"]:
                row_ids = {relevant_params(x["href"]).get("id") for x in row["links"]
                           if "klassement.jspa" in x["href"]
                           and relevant_params(x["href"]).get("id") is not None}
                if len(row_ids) > 1 or row_ids & row_record_ids:
                    raise AanpassingFout("Dubbele of tegenstrijdige klassementrecords.")
                row_record_ids.update(row_ids)
        for h in rank_headers:
            data_rows = [r for r in grids[h][1:] if any(r)]
            if data_rows and not records:
                raise AanpassingFout("Klassementrijen zonder recordidentiteit; veilig gestopt.")
        return grids, records

    def rank_create(self, person_id):
        form = self.get("klassement.jspa", {"dispatch": "update", "idPerson": person_id}).form_for("klassementSave.jspa")
        self.identity(form, "idPerson", person_id)
        if (form.controls.get("actionOnKlassement", {}).get("value") != "create"
                or form.controls.get("id", {}).get("value") not in {"", None}):
            raise AanpassingFout("Klassementformulier staat niet in create-modus.")
        if "eindDatum" in form.values():
            raise AanpassingFout("Onverwachte einddatum bij klassementcreatie.")
        return form

    def rank_existing(self, person_id, record_id):
        form = self.get("klassement.jspa", {"dispatch": "update", "idPerson": person_id,
                                          "id": record_id}).form_for("klassementSave.jspa")
        self.identity(form, "idPerson", person_id)
        self.identity(form, "id", record_id)
        if form.controls.get("actionOnKlassement", {}).get("value") != "update":
            raise AanpassingFout("Bestaand klassementformulier staat niet in update-modus.")
        return form

    def snapshot(self, person_id):
        person = self.person(person_id)
        statute = self.statute(person_id)
        relations, records = self.relations(person_id)
        rankings = {}
        for key in records:
            values = self.rank_existing(person_id, key).values()
            # Gender-dependent empty controls may appear/disappear without changing a record.
            values.setdefault("klassementDames", "")
            rankings[key] = values
        return {"person": person.values(), "statute": statute.values(), "relations": relations,
                "rankings": rankings}, {"person": person, "statute": statute}

    def save(self, form, changes):
        action = urljoin(VTTL_BASE_URL, form.action)
        parts = urlsplit(action)
        if parts.scheme != "https" or parts.netloc != urlsplit(VTTL_BASE_URL).netloc or (
                parts.path.split(";")[0] not in {"/personSave.jspa", "/statuutSave.jspa", "/klassementSave.jspa"}):
            raise AanpassingFout("Onveilige formulierbestemming.")
        data = form.payload(changes)
        try:
            response = self.session.post(action, data=data, timeout=30, allow_redirects=False)
            response.raise_for_status()
        except requests.RequestException:
            raise AanpassingFout("Opslagverzoek mislukt; online resultaat mogelijk onzeker.") from None
        # A 200 response is not proof of success: callers always re-read online state.


def preflight(client, items):
    plans, targets = [], set()
    for item in items:
        person_id = client.resolve(item)
        if person_id in targets:
            raise AanpassingFout("Meerdere configitems wijzen naar hetzelfde lid.")
        targets.add(person_id)
        before, forms = client.snapshot(person_id)
        actual_name = before["person"]["lastName"] + " " + before["person"]["firstName"]
        if normaliseer(actual_name) != normaliseer(item["naam"]):
            raise AanpassingFout("Identificatienaam gewijzigd tijdens preflight; veilig gestopt.")
        forms["person"].validate(item["person"])
        gender = item["person"].get("geslacht", before["person"]["geslacht"])
        if gender == "man" and before["person"]["geslacht"] == "vrouw" and any(
                record.get("klassementDames") for record in before["rankings"].values()):
            raise AanpassingFout("Geslachtswijziging maakt bestaand damesklassement ongeschikt/onverifieerbaar.")
        if item["statute"] is not None:
            forms["statute"].validate({"statutenAsString": item["statute"]})
        rank_form = None
        if item["rank"]:
            if before["rankings"]:
                raise AanpassingFout("Klassement aangevraagd maar er bestaat al een (mogelijk historisch) record; alles geblokkeerd.")
            if "klassementDames" in item["rank"] and gender != "vrouw":
                raise AanpassingFout("Damesklassement is uitsluitend geldig voor een vrouwelijke persoon.")
            rank_form = client.rank_create(person_id)
            # A gender change may reveal a field only after person save. Validate all
            # other fields now; the new surface is verified again before creation.
            pending = dict(item["rank"])
            if ("klassementDames" in pending and "klassementDames" not in rank_form.values()
                    and gender == "vrouw" and before["person"]["geslacht"] == "man"):
                del pending["klassementDames"]
            rank_form.validate(pending)
        operations = []
        if any(before["person"][k] != v for k, v in item["person"].items()):
            operations.append(("person", item["person"]))
        if item["statute"] is not None and before["statute"]["statutenAsString"] != item["statute"]:
            operations.append(("statute", {"statutenAsString": item["statute"]}))
        if item["rank"]:
            operations.append(("rank", item["rank"]))
        plans.append({"item": item, "id": person_id, "before": before, "operations": operations})
    return plans


def show_proposal(plans, output=print):
    output("Volledig voorstel (nog niets opgeslagen):")
    for plan in plans:
        item = plan["item"]
        output(f"Doel: {item['naam']} | lidnummer {item['lidnummer'] or '(online opgelost)'} | persoon-id {plan['id']}")
        proposed = [("person", item["person"])]
        if item["statute"] is not None:
            proposed.append(("statute", {"statutenAsString": item["statute"]}))
        if item["rank"]:
            proposed.append(("rank", item["rank"]))
        for section, changes in proposed:
            for field, value in changes.items():
                old = plan["before"].get(section, {}).get(field, "(geen record)")
                label = PUBLIC[field]
                suffix = " [standaard: vandaag]" if field == "startDatum" and item["default_date"] else ""
                if old == value:
                    suffix += " [ongewijzigd]"
                output(f"  {section}/{label}: {old!r} -> {value!r}{suffix}")
        if not plan["operations"]:
            output("  Geen verschil: niets op te slaan.")


def verify_after(before, after, section, changes):
    expected_person = dict(before["person"])
    expected_statute = dict(before["statute"])
    if section == "person":
        expected_person.update(changes)
    if section == "statute":
        expected_statute.update(changes)
    if after["person"] != expected_person or after["statute"] != expected_statute:
        raise AanpassingFout("Online herlezing: bedoelde of ongemoeide persoonsvelden/statuut wijken af.")
    if section == "rank":
        if before["rankings"] or len(after["rankings"]) != 1:
            raise AanpassingFout("Online herlezing: klassementcreatie niet eenduidig bevestigd.")
        record = next(iter(after["rankings"].values()))
        if any(record.get(k) != v for k, v in changes.items()):
            raise AanpassingFout("Online herlezing: klassementwaarden niet bevestigd.")
        if record.get("eindDatum", ""):
            raise AanpassingFout("Online herlezing: onverwachte einddatum.")
    elif after["rankings"] != before["rankings"]:
        raise AanpassingFout("Online herlezing: bestaand klassement is gewijzigd.")
    def unrelated(grids):
        return {h: rows for h, rows in grids.items()
                if "Klassement heren" not in h
                and not (section == "statute" and "Statuut" in h)}
    if unrelated(before["relations"]) != unrelated(after["relations"]):
        raise AanpassingFout("Online herlezing: ongemoeide relaties zijn gewijzigd.")


def execute(client, items, dry_run=False, confirm=input, output=print):
    plans = preflight(client, items)
    show_proposal(plans, output)
    if dry_run:
        output("Dry-run: geen opslagverzoeken en geen bevestiging.")
        return
    if not any(plan["operations"] for plan in plans):
        output("Geen wijzigingen nodig.")
        return
    if confirm("Alle bovenstaande wijzigingen opslaan? Typ JA: ").strip() != "JA":
        output("Geannuleerd; niets opgeslagen.")
        return
    completed, pending = [], None
    try:
        for plan in plans:
            expected = plan["before"]
            for section, changes in plan["operations"]:
                pending = None
                current, forms = client.snapshot(plan["id"])
                if current != expected:
                    raise AanpassingFout("Concurrentie/preconditie: online toestand gewijzigd sinds voorstel; gestopt vóór volgende opslag.")
                form = client.rank_create(plan["id"]) if section == "rank" else forms[section]
                if section == "rank" and current["rankings"]:
                    raise AanpassingFout("Klassement inmiddels aanwezig; gestopt vóór opslag.")
                form.validate(changes)
                pending = f"persoon-id {plan['id']}/{section}"
                client.save(form, changes)
                after, _ = client.snapshot(plan["id"])
                verify_after(expected, after, section, changes)
                completed.append(pending)
                output(f"Opgeslagen en online geverifieerd: {pending}")
                pending = None
                expected = after
    except (AanpassingFout, KeyboardInterrupt, EOFError):
        output("STOP. Geen automatische rollback, retry of verdere opslag.")
        output("Reeds opgeslagen en geverifieerd: " + (", ".join(completed) or "geen"))
        output("Mogelijk opgeslagen/onzeker: " + (pending or "geen opslagverzoek onderweg"))
        raise
    output("Alle aangevraagde wijzigingen opgeslagen en online geverifieerd.")


def run(argv=None):
    parser = argparse.ArgumentParser(description="Pas VTTL-persoonsgegevens expliciet en gecontroleerd aan.")
    parser.add_argument("--config", required=True, help="Pad naar TOML-configuratie")
    parser.add_argument("--dry-run", action="store_true", help="Alleen online lezen en voorstel tonen")
    args = parser.parse_args(argv)
    session = None
    try:
        items = load_config(args.config)
        if not items:
            print("Geen leden geconfigureerd; niets gelezen of opgeslagen.")
            return
        password = os.getenv("PASWOORD_TT") or getpass.getpass("VTTL wachtwoord (PASWOORD_TT): ")
        session = login_session(password)
        execute(Client(session), items, args.dry_run)
    except (AanpassingFout, RuntimeError) as exc:
        parser.exit(1, f"Fout: {exc}\n")
    except (KeyboardInterrupt, EOFError):
        parser.exit(1, "Afgebroken.\n")
    finally:
        if session is not None:
            session.close()


if __name__ == "__main__":
    run()
