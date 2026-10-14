"""Offline portal fixtures: every HTTP call, including saves, is mocked."""

import copy
import datetime as dt
from html import escape
import io
import unittest
from unittest.mock import Mock, patch

import requests

from ttpy import ledenlijst_aanpassing as cli
from ttpy import vttl_download


def inp(name, value="", kind="text", extra=""):
    return f'<input type="{kind}" name="{name}" value="{escape(str(value))}" {extra}>'


def form(action, fields):
    return f'<form action="/{action}">{fields}{inp("save", kind="image")}</form>'


def select(name, values, chosen, multiple=False):
    values = list(values)
    selected = chosen if type(chosen) is list else [chosen]
    options = "".join(f'<option value="{x}" {"selected" if x in selected else ""}>{x}</option>'
                      for x in values)
    return f'<select name="{name}" {"multiple" if multiple else ""}>{options}</select>'


def grid(headers, rows):
    return "<table><tr>" + "".join(f"<th>{x}</th>" for x in headers) + "</tr>" + "".join(
        "<tr>" + "".join(f"<td>{x}</td>" for x in row) + "</tr>" for row in rows) + "</table>"


class Portal:
    """Synthetic state machine behind a requests-compatible mocked session."""

    def __init__(self, count=1):
        self.people = {}
        for number in range(1, count + 1):
            values = {target: "" for target, _ in cli.TEXT.values()}
            values.update({target: False for target in cli.BOOL.values()})
            values.update({target: options[0] for target, options in cli.ENUM.values()})
            values.update(firstName="Voorbeeld", lastName=f"Test{number}", straat="Bestaande straat",
                          gSportersCategorieen=["AU"], geslacht="man")
            self.people[str(number)] = {"person": values, "statute": "administratiefLid",
                                        "ranks": {}, "club": "ONVERANDERDE RELATIE"}
        self.session = Mock()
        self.session.get.side_effect = self.get
        self.session.post.side_effect = self.post
        self.calls = []
        self.failure = None
        self.corrupt = None

    @staticmethod
    def response(text):
        r = Mock()
        r.text = text
        r.raise_for_status.return_value = None
        return r

    def get(self, url, params=None, timeout=None):
        route = url.split("/")[-1]
        params = params or {}
        if route == "personList.jspa":
            rows = []
            for key, person in self.people.items():
                if params.get("filterLidnummer") and int(params["filterLidnummer"]) != int(key):
                    continue
                p = person["person"]
                link = f'<a href="/person.jspa?dispatch=view&amp;id={key}">view</a>'
                rows.append([link, key, p["lastName"] + " " + p["firstName"]])
            return self.response(form("personList.jspa", inp("filterNaamPersoon") + inp("filterLidnummer"))
                                 + grid(["", "Aansl.Nr.", "Naam"], rows))
        key = params.get("idPerson", params.get("id"))
        state = self.people[key]
        if route == "person.jspa" and params["dispatch"] == "update":
            values = state["person"]
            fields = inp("id", key, "hidden") + inp("actionOnPerson", "update", "hidden")
            for _, (target, limit) in cli.TEXT.items():
                fields += inp(target, values[target], extra=f'maxlength="{limit}"')
            for target in cli.BOOL.values():
                fields += inp(target, "true", "checkbox", "checked" if values[target] else "")
            for _, (target, options) in cli.ENUM.items():
                fields += select(target, options, values[target])
            fields += select("gSportersCategorieen", cli.GSPORT, values["gSportersCategorieen"], True)
            fields += inp("password", "HIDDEN_TEST_SECRET", "hidden") + inp("userId", "TEST_AUTH", "hidden")
            return self.response(form("personSave.jspa", fields))
        if route == "statuut.jspa":
            fields = inp("id", "ST1", "hidden") + inp("idPerson", key, "hidden")
            fields += inp("actionOnStatuut", "update", "hidden")
            fields += "".join(inp("statutenAsString", x, "radio",
                                  "checked" if x == state["statute"] else "") for x in cli.STATUTEN)
            return self.response(form("statuutSave.jspa", fields))
        if route == "person.jspa" and params["dispatch"] == "view":
            rank_rows = []
            for rid, rank in state["ranks"].items():
                # Deliberately duplicated query parameters and malformed JSP prefix.
                link = (f'<a href="/klassement.jspa?dispatch=view&amp;idPerson=&lt;%=id%&gt;'
                        f'&amp;id=bad&amp;dispatch=view&amp;idPerson={key}&amp;id={rid}">view</a>')
                rank_rows.append([link, rank.get("klassementHeren", ""), rank["startDatum"],
                                  rank.get("eindDatum", "")])
            create = f'<a href="/klassement.jspa?dispatch=update&amp;idPerson={key}">create</a>'
            return self.response(form("person.jspa", inp("id", key, "hidden")) + create
                                 + grid(["", "Naam club", "Begin", "Einde"], [["", state["club"], "", ""]])
                                 + grid(["", "Klassement heren", "Begin", "Einde"], rank_rows)
                                 + grid(["", "Actief", "Begin", "Einde", "Statuut"],
                                        [["", "true", "", "", state["statute"]]]))
        if route == "klassement.jspa":
            rid = params.get("id", "")
            rank = state["ranks"].get(rid, {})
            fields = inp("id", rid, "hidden") + inp("idPerson", key, "hidden")
            fields += inp("actionOnKlassement", "update" if rid else "create", "hidden")
            fields += inp("klassementHeren", rank.get("klassementHeren", ""), extra='maxlength="4"')
            if state["person"]["geslacht"] == "vrouw":
                fields += inp("klassementDames", rank.get("klassementDames", ""), extra='maxlength="4"')
            fields += inp("startDatum", rank.get("startDatum", ""), extra='maxlength="10"')
            if rid:
                fields += inp("eindDatum", rank.get("eindDatum", ""), extra='maxlength="10"')
            return self.response(form("klassementSave.jspa", fields))
        raise AssertionError("Unexpected GET route")

    def post(self, url, data, timeout, allow_redirects):
        if allow_redirects:
            raise AssertionError("Never redirect/retry a save POST")
        route = url.split("/")[-1]
        self.calls.append(route)
        values = dict(data)
        if self.failure == len(self.calls):
            raise requests.Timeout()
        key = values.get("idPerson", values.get("id"))
        state = self.people[key]
        if route == "personSave.jspa":
            p = state["person"]
            for target, _ in cli.TEXT.values():
                p[target] = values[target]
            for target in cli.BOOL.values():
                p[target] = target in values
            for target, _ in cli.ENUM.values():
                p[target] = values[target]
            p["gSportersCategorieen"] = [v for k, v in data if k == "gSportersCategorieen"]
        elif route == "statuutSave.jspa":
            state["statute"] = values["statutenAsString"]
        elif route == "klassementSave.jspa":
            state["ranks"]["99"] = {k: values[k] for k in
                                   ("klassementHeren", "klassementDames", "startDatum") if k in values}
        else:
            raise AssertionError("Forbidden POST route")
        if self.corrupt:
            self.corrupt(state)
        return self.response("HTTP 200 is not proof of success")

    def client(self):
        return cli.Client(self.session)


def config(**kwargs):
    raw = {"naam": "Test1 Voorbeeld", "lidnummer": 1, "wijzigingen": {"straat": "Nieuwe straat"}}
    raw.update(kwargs)
    return cli.parse_config({"leden": [raw]}, today=dt.date(2026, 10, 11))


class ConfigTests(unittest.TestCase):
    def test_empty_skipped_false_zero_and_no_are_values(self):
        item = config(wijzigingen={"straat": "", "gsport_categorieen": [],
                                   "coachlicentie": False, "huisnummer": 0,
                                   "vts_kwalificatie": "geen"})[0]
        self.assertEqual(item["person"], {"coachLicentie": False, "huisNummer": "0",
                                         "vtsKwalificatie": "geen"})

    def test_identification_is_separate(self):
        item = config(wijzigingen={"voornaam": "Anders", "achternaam": "Nieuw"})[0]
        self.assertEqual(item["naam"], "Test1 Voorbeeld")
        self.assertEqual(item["person"]["firstName"], "Anders")

    def test_default_and_explicit_dates(self):
        item = config(klassement={"heren": "D0", "begindatum": ""})[0]
        self.assertEqual(item["rank"]["startDatum"], "11/10/2026")
        self.assertTrue(item["default_date"])
        item = config(klassement={"dames": "B2", "begindatum": "01/01/2027"})[0]
        self.assertFalse(item["default_date"])

    def test_invalid_values(self):
        cases = [
            {"wijzigingen": {"straat": "x" * 51}},
            {"wijzigingen": {"straat": 0}},
            {"wijzigingen": {"postcode": True}},
            {"wijzigingen": {"coachlicentie": 0}},
            {"wijzigingen": {"geslacht": "anders"}},
            {"wijzigingen": {"geboortedatum": "31/02/2026"}},
            {"wijzigingen": {"geboortedatum": "2026-01-01"}},
            {"wijzigingen": {"gsport_categorieen": ["AU", "AU"]}},
            {"wijzigingen": {"gsport_categorieen": ["INVALID"]}},
            {"wijzigingen": {"gsport_categorieen": "AU"}},
            {"wijzigingen": {"password": "DO_NOT_PRINT"}},
            {"klassement": {"heren": ["D0"]}},
            {"klassement": {"heren": "12345"}},
            {"klassement": {"begindatum": "01/01/2026"}},
            {"klassement": {"einddatum": ""}},
            {"statuut": False}, {"lidnummer": False}, {"lidnummer": "1234567"},
            {"naam": ""}, {"unknown": "DO_NOT_PRINT"},
        ]
        for case in cases:
            with self.subTest(case=case), self.assertRaises(cli.AanpassingFout) as caught:
                config(**case)
            self.assertNotIn("DO_NOT_PRINT", str(caught.exception))

    def test_unknown_even_empty(self):
        with self.assertRaises(cli.AanpassingFout):
            config(wijzigingen={"unknown": ""})

    def test_all_empty_is_noop(self):
        item = config(wijzigingen={"straat": "", "coachlicentie": [], "gsport_categorieen": []},
                      statuut="", klassement={"heren": "", "begindatum": ""})[0]
        self.assertEqual(item["person"], {})
        self.assertIsNone(item["statute"])
        self.assertEqual(item["rank"], {})
        self.assertEqual(cli.parse_config({"leden": []}), [])

    def test_duplicate_members(self):
        raw = {"naam": "Tést Voorbeeld", "wijzigingen": {"straat": "nieuw"}}
        with self.assertRaises(cli.AanpassingFout):
            cli.parse_config({"leden": [raw, dict(raw, naam=" TEST   VOORBEELD ")]})

    def test_toml_duplicates_sanitized(self):
        with patch("builtins.open", return_value=io.BytesIO(b'password="secret"\npassword="secret"\n')):
            with self.assertRaises(cli.AanpassingFout) as caught:
                cli.load_config("not_a_real_file")
        self.assertNotIn("secret", str(caught.exception))


class PortalTests(unittest.TestCase):
    def test_dry_run_never_saves_or_confirms(self):
        portal = Portal()
        confirm = Mock(side_effect=AssertionError("No confirmation in dry run"))
        output = []
        cli.execute(portal.client(), config(statuut="competitiefLid", klassement={"heren": "D0"}),
                    dry_run=True, confirm=confirm, output=output.append)
        portal.session.post.assert_not_called()
        confirm.assert_not_called()
        self.assertIn("standaard: vandaag", "\n".join(output))
        self.assertNotIn("HIDDEN_TEST_SECRET", "\n".join(output))
        self.assertNotIn("TEST_AUTH", "\n".join(output))

    def test_denied_confirmation_saves_nothing(self):
        portal = Portal()
        cli.execute(portal.client(), config(), confirm=lambda _: "nee", output=lambda _: None)
        portal.session.post.assert_not_called()

    def test_one_confirmation_for_all_operations(self):
        portal = Portal()
        confirm = Mock(return_value="JA")
        cli.execute(portal.client(), config(statuut="competitiefLid", klassement={"heren": "D0"}),
                    confirm=confirm, output=lambda _: None)
        self.assertEqual(portal.calls, ["personSave.jspa", "statuutSave.jspa", "klassementSave.jspa"])
        confirm.assert_called_once()
        self.assertEqual(portal.people["1"]["person"]["gSportersCategorieen"], ["AU"])

    def test_preserves_all_unconfigured_fields_and_replaces_list(self):
        portal = Portal()
        portal.people["1"]["person"]["coachLicentie"] = True
        before = copy.deepcopy(portal.people["1"])
        changes = {"huisnummer": 0, "coachlicentie": False, "gsport_categorieen": ["FYSR", "VI"]}
        cli.execute(portal.client(), config(wijzigingen=changes), confirm=lambda _: "JA", output=lambda _: None)
        expected = copy.deepcopy(before)
        expected["person"].update(huisNummer="0", coachLicentie=False, gSportersCategorieen=["FYSR", "VI"])
        self.assertEqual(portal.people["1"], expected)
        sent = dict(portal.session.post.call_args.kwargs["data"])
        self.assertEqual(sent["password"], "HIDDEN_TEST_SECRET")
        self.assertNotIn("coachLicentie", sent)

    def test_multiple_selection_order_is_not_a_change(self):
        portal = Portal()
        cli.execute(portal.client(), config(wijzigingen={"gsport_categorieen": ["VI", "FYSR"]}),
                    confirm=lambda _: "JA", output=lambda _: None)
        self.assertEqual(set(portal.people["1"]["person"]["gSportersCategorieen"]), {"VI", "FYSR"})

    def test_existing_historical_rank_blocks_entire_batch(self):
        portal = Portal(2)
        portal.people["2"]["ranks"]["10"] = {"klassementHeren": "D0", "startDatum": "01/01/2000",
                                             "eindDatum": "01/01/2001"}
        data = {"leden": [
            {"naam": "Test1 Voorbeeld", "lidnummer": 1, "wijzigingen": {"straat": "nieuw"}},
            {"naam": "Test2 Voorbeeld", "lidnummer": 2, "klassement": {"heren": "D2"}}]}
        confirm = Mock()
        with self.assertRaises(cli.AanpassingFout):
            cli.execute(portal.client(), cli.parse_config(data), confirm=confirm, output=lambda _: None)
        portal.session.post.assert_not_called()
        confirm.assert_not_called()

    def test_no_class_config_preserves_existing_record(self):
        portal = Portal()
        portal.people["1"]["ranks"]["10"] = {"klassementHeren": "D0", "startDatum": "01/01/2000",
                                             "eindDatum": "01/01/2001"}
        before = copy.deepcopy(portal.people["1"]["ranks"])
        cli.execute(portal.client(), config(), confirm=lambda _: "JA", output=lambda _: None)
        self.assertEqual(portal.people["1"]["ranks"], before)

    def test_dames_for_man_rejected(self):
        portal = Portal()
        with self.assertRaises(cli.AanpassingFout):
            cli.execute(portal.client(), config(klassement={"dames": "B2"}), dry_run=True)
        portal.session.post.assert_not_called()

    def test_gender_change_then_both_rank_fields(self):
        portal = Portal()
        cli.execute(portal.client(), config(wijzigingen={"geslacht": "vrouw"},
                                           klassement={"heren": "D0", "dames": "B2"}),
                    confirm=lambda _: "JA", output=lambda _: None)
        self.assertEqual(portal.calls, ["personSave.jspa", "klassementSave.jspa"])
        self.assertEqual(portal.people["1"]["ranks"]["99"]["klassementDames"], "B2")

    def test_empty_existing_statute_can_be_set(self):
        portal = Portal()
        portal.people["1"]["statute"] = ""
        cli.execute(portal.client(), config(wijzigingen={}, statuut="competitiefLid"),
                    confirm=lambda _: "JA", output=lambda _: None)
        self.assertEqual(portal.people["1"]["statute"], "competitiefLid")

    def test_partial_failure_reports_verified_and_uncertain(self):
        portal = Portal()
        portal.failure = 2
        output = []
        with self.assertRaises(cli.AanpassingFout):
            cli.execute(portal.client(), config(statuut="competitiefLid", klassement={"heren": "D0"}),
                        confirm=lambda _: "JA", output=output.append)
        self.assertEqual(portal.calls, ["personSave.jspa", "statuutSave.jspa"])
        text = "\n".join(output)
        self.assertIn("Reeds opgeslagen en geverifieerd: persoon-id 1/person", text)
        self.assertIn("Mogelijk opgeslagen/onzeker: persoon-id 1/statute", text)

    def test_http_success_but_ignored_save_is_failure(self):
        portal = Portal()
        portal.corrupt = lambda state: state["person"].update(straat="Bestaande straat")
        with self.assertRaises(cli.AanpassingFout):
            cli.execute(portal.client(), config(), confirm=lambda _: "JA", output=lambda _: None)
        self.assertEqual(len(portal.calls), 1)

    def test_unconfigured_person_statute_class_or_relation_change_fails(self):
        corruptions = [
            lambda s: s["person"].update(email="unexpected"),
            lambda s: s.update(statute="competitiefLid"),
            lambda s: s.update(club="CHANGED"),
            lambda s: s["ranks"].update({"9": {"klassementHeren": "D0", "startDatum": "01/01/2020"}}),
        ]
        for corrupt in corruptions:
            portal = Portal()
            portal.corrupt = corrupt
            with self.subTest(corrupt=corrupt), self.assertRaises(cli.AanpassingFout):
                cli.execute(portal.client(), config(), confirm=lambda _: "JA", output=lambda _: None)
            self.assertEqual(len(portal.calls), 1)

    def test_concurrency_before_first_save_blocks(self):
        portal = Portal()
        client = portal.client()
        original = client.snapshot
        calls = 0
        def snapshot(key):
            nonlocal calls
            calls += 1
            if calls == 2:
                portal.people[key]["person"]["email"] = "concurrent"
            return original(key)
        client.snapshot = snapshot
        with self.assertRaises(cli.AanpassingFout):
            cli.execute(client, config(), confirm=lambda _: "JA", output=lambda _: None)
        portal.session.post.assert_not_called()

    def test_concurrency_before_second_save_stops_after_verified_first(self):
        portal = Portal()
        client = portal.client()
        original = client.snapshot
        calls = 0
        def snapshot(key):
            nonlocal calls
            calls += 1
            if calls == 4:
                portal.people[key]["person"]["email"] = "concurrent"
            return original(key)
        client.snapshot = snapshot
        output = []
        with self.assertRaises(cli.AanpassingFout):
            cli.execute(client, config(statuut="competitiefLid"), confirm=lambda _: "JA", output=output.append)
        self.assertEqual(portal.calls, ["personSave.jspa"])
        self.assertIn("geen opslagverzoek onderweg", "\n".join(output))

    def test_missing_and_ambiguous_names(self):
        portal = Portal(2)
        client = portal.client()
        with self.assertRaises(cli.AanpassingFout):
            client.resolve({"naam": "Niet Bestaand", "lidnummer": None})
        portal.people["2"]["person"].update(lastName="Tést1", firstName="Voorbeeld")
        with self.assertRaises(cli.AanpassingFout):
            client.resolve({"naam": " TEST1   voorbeeld ", "lidnummer": None})
        self.assertEqual(client.resolve({"naam": " TEST1   voorbeeld ", "lidnummer": 1}), "1")
        with self.assertRaises(cli.AanpassingFout):
            client.resolve({"naam": "Verkeerde Naam", "lidnummer": 1})
        portal.session.post.assert_not_called()

    def test_name_pagination_and_inactive_filter(self):
        portal = Portal(2)
        original = portal.get
        def get(url, params=None, timeout=None):
            if url.endswith("personList.jspa") and params.get("nextPage") == "1":
                response = original(url, {"filterLidnummer": 1}, timeout)
                response.text += '<a href="/personList.jspa?nextPage=2">2</a>'
                return response
            if url.endswith("personList.jspa"):
                return original(url, {"filterLidnummer": 2}, timeout)
            return original(url, params, timeout)
        portal.session.get.side_effect = get
        self.assertEqual(portal.client().resolve({"naam": "Test2 Voorbeeld", "lidnummer": None}), "2")
        queries = [x.kwargs["params"] for x in portal.session.get.call_args_list
                   if x.args[0].endswith("personList.jspa")]
        self.assertEqual(len(queries), 2)
        self.assertTrue(all(x["filterActief"] == "AAAAActief" for x in queries))
        portal.session.post.assert_not_called()

    def test_same_target_multiple_config_items_rejected(self):
        portal = Portal()
        items = config() + config(lidnummer="", wijzigingen={"email": "synthetic"})
        with self.assertRaises(cli.AanpassingFout):
            cli.execute(portal.client(), items, dry_run=True)
        portal.session.post.assert_not_called()

    def test_malformed_or_duplicate_class_records_fail_closed(self):
        portal = Portal()
        original = portal.get
        def get(url, params=None, timeout=None):
            r = original(url, params, timeout)
            if url.endswith("person.jspa") and params["dispatch"] == "view":
                r.text += '<a href="/klassement.jspa?dispatch=view&amp;idPerson=1&amp;id=BAD">bad</a>'
            return r
        portal.session.get.side_effect = get
        with self.assertRaises(cli.AanpassingFout):
            cli.preflight(portal.client(), config(klassement={"heren": "D0"}))
        portal.session.post.assert_not_called()

    def test_duplicate_class_rows_fail_closed(self):
        portal = Portal()
        portal.people["1"]["ranks"]["10"] = {"klassementHeren": "D0", "startDatum": "01/01/2000"}
        original = portal.get
        def get(url, params=None, timeout=None):
            r = original(url, params, timeout)
            if url.endswith("person.jspa") and params["dispatch"] == "view":
                link = '<a href="/klassement.jspa?dispatch=view&amp;idPerson=1&amp;id=10">view</a>'
                row = "<tr><td>" + link + "</td><td>D0</td><td>01/01/2000</td><td></td></tr>"
                marker = "</table><table><tr><th></th><th>Actief"
                r.text = r.text.replace(marker, row + marker)
            return r
        portal.session.get.side_effect = get
        with self.assertRaises(cli.AanpassingFout):
            cli.preflight(portal.client(), config())
        portal.session.post.assert_not_called()

    def test_gender_change_cannot_hide_existing_dames_rank(self):
        portal = Portal()
        portal.people["1"]["person"]["geslacht"] = "vrouw"
        portal.people["1"]["ranks"]["10"] = {"klassementHeren": "D0", "klassementDames": "B2",
                                             "startDatum": "01/01/2000"}
        with self.assertRaises(cli.AanpassingFout):
            cli.preflight(portal.client(), config(wijzigingen={"geslacht": "man"}))
        portal.session.post.assert_not_called()

    def test_form_duplicate_and_invalid_option_rejected(self):
        with self.assertRaises(cli.AanpassingFout):
            cli.Page(form("personSave.jspa", inp("straat") + inp("straat")))
        f = cli.Page(form("personSave.jspa", select("geslacht", ["man"], "man"))).forms[0]
        with self.assertRaises(cli.AanpassingFout):
            f.validate({"geslacht": "vrouw"})

    def test_save_action_origin_checked(self):
        portal = Portal()
        f = cli.Form("https://example.invalid/personSave.jspa")
        with self.assertRaises(cli.AanpassingFout):
            portal.client().save(f, {})
        portal.session.post.assert_not_called()

    def test_no_changes_no_save_or_confirmation(self):
        portal = Portal()
        confirm = Mock()
        cli.execute(portal.client(), config(wijzigingen={"straat": "Bestaande straat"}),
                    confirm=confirm, output=lambda _: None)
        confirm.assert_not_called()
        portal.session.post.assert_not_called()


class LoginTests(unittest.TestCase):
    def test_login_only_helper_never_triggers_report(self):
        s = Mock()
        s.headers = {}
        s.cookies = {}
        login = Portal.response("jsessionid=ABCDEF")
        authenticated = Portal.response("Uitloggen")
        s.get.return_value = login
        s.post.return_value = authenticated
        with patch.object(vttl_download.requests, "Session", return_value=s):
            self.assertIs(vttl_download.login_session("SYNTHETIC_PASSWORD"), s)
        self.assertEqual(s.get.call_count, 1)
        self.assertEqual(s.post.call_count, 1)
        self.assertIn("/loginSave.jspa;", s.post.call_args.args[0])

    def test_export_uses_helper_and_retains_report_submission(self):
        s = Mock()
        with patch.object(vttl_download, "login_session", return_value=s) as login, patch("builtins.print"):
            result = vttl_download._trigger_report("SYNTHETIC_PASSWORD")
        login.assert_called_once_with("SYNTHETIC_PASSWORD")
        self.assertIsNotNone(result.tzinfo)
        self.assertTrue(s.get.call_args.args[0].endswith("/reportList.jspa"))
        self.assertTrue(s.post.call_args.args[0].endswith("/report.jspa"))
        self.assertEqual(s.post.call_args.kwargs["files"]["selectedReport"], (None, "Export"))

    def test_login_failure_redacts_transport_details(self):
        s = Mock()
        s.headers = {}
        s.get.side_effect = requests.ConnectionError("SYNTHETIC_SECRET")
        with patch.object(vttl_download.requests, "Session", return_value=s):
            with self.assertRaises(RuntimeError) as caught:
                vttl_download.login_session("SYNTHETIC_SECRET")
        self.assertNotIn("SYNTHETIC_SECRET", str(caught.exception))
        s.close.assert_called_once()

    def test_cli_dry_run_login_only_and_no_mail_prompt(self):
        portal = Portal()
        with patch.object(cli, "load_config", return_value=config()), \
                patch.object(cli, "login_session", return_value=portal.session) as login, \
                patch.dict("os.environ", {"PASWOORD_TT": "SYNTHETIC_PASSWORD"}), \
                patch.object(cli.getpass, "getpass") as prompt, patch("builtins.print"):
            cli.run(["--config", "synthetic.toml", "--dry-run"])
        login.assert_called_once()
        prompt.assert_not_called()
        portal.session.post.assert_not_called()
        portal.session.close.assert_called_once()

    def test_cli_requires_config_and_has_no_yes(self):
        with patch("sys.stderr", new_callable=io.StringIO):
            for args in ([], ["--config", "synthetic.toml", "--yes"]):
                with self.assertRaises(SystemExit) as caught:
                    cli.run(args)
                self.assertEqual(caught.exception.code, 2)

    def test_blank_template_does_not_login(self):
        with patch.object(cli, "load_config", return_value=[]), \
                patch.object(cli, "login_session") as login, patch("builtins.print"):
            cli.run(["--config", "blank.toml"])
        login.assert_not_called()


if __name__ == "__main__":
    unittest.main()
