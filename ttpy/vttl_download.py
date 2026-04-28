"""
vttl_download.py
----------------
Triggert het ledenrapport op leden.vttl.be, wacht op de email met
CSV-bijlagen, slaat ze op en vergelijkt met de vorige export.
"""

import os
import sys
import re
import time
import imaplib
import email
import datetime
import shutil

try:
    import requests
except ImportError:
    sys.exit("Installeer requests:  pip install requests")

try:
    import pandas as pd
except ImportError:
    sys.exit("Installeer pandas:  pip install pandas")

# ── Configuratie ──────────────────────────────────────────────────────────────
VTTL_BASE_URL   = "https://leden.vttl.be"
VTTL_LOGIN      = "hans hooyberghs"

IMAP_SERVER     = "mail.tafeltennisantwerpen.be"
IMAP_PORT       = 993
IMAP_USER       = "secretariaat@tafeltennisantwerpen.be"
IMAP_FOLDER     = "INBOX"
VTTL_MAIL_FROM  = "leden@vttl.be"
VTTL_MAIL_SUBJ  = "Export VTTL-database"

WAIT_SECONDS    = 300
POLL_INTERVAL   = 10
# ─────────────────────────────────────────────────────────────────────────────


def _get_passwords():
    vttl_pwd = os.getenv("PASWOORD_TT")
    mail_pwd = os.getenv("MAIL_PASSWORD") or vttl_pwd
    if not vttl_pwd:
        import getpass
        vttl_pwd = getpass.getpass("VTTL wachtwoord (PASWOORD_TT): ")
    if not mail_pwd:
        import getpass
        mail_pwd = getpass.getpass("IMAP wachtwoord (MAIL_PASSWORD): ")
    return vttl_pwd, mail_pwd


def _trigger_report(vttl_pwd: str) -> datetime.datetime:
    s = requests.Session()
    s.headers["User-Agent"] = "Mozilla/5.0"

    r1 = s.get(f"{VTTL_BASE_URL}/login.jspa?dispatch=view")
    jsid = re.search(r"jsessionid=([A-F0-9]+)", r1.text).group(1)

    r2 = s.post(
        f"{VTTL_BASE_URL}/loginSave.jspa;jsessionid={jsid}",
        data={"userId": VTTL_LOGIN, "password": vttl_pwd,
              "logIn.x": "50", "logIn.y": "10"},
    )
    if "Uitloggen" not in r2.text:
        sys.exit("✗ Login mislukt op leden.vttl.be – controleer PASWOORD_TT.")
    print(f"✓ Ingelogd op leden.vttl.be als '{VTTL_LOGIN}'")

    triggered_at = datetime.datetime.now(datetime.timezone.utc)

    s.get(f"{VTTL_BASE_URL}/reportList.jspa")
    try:
        s.post(
            f"{VTTL_BASE_URL}/report.jspa",
            data={
                "filterProvincie": "", "filterAansluitingsNrClub": "",
                "filterActief": "true", "extraLabelText": "",
                "startLabel": "1", "selectedLabel": "0",
                "selectedReport": "LedenLijst",
                "prepare.x": "50", "prepare.y": "10",
            },
            timeout=15,
        )
    except requests.exceptions.Timeout:
        pass

    print(f"✓ Rapport getriggerd om {triggered_at.strftime('%H:%M:%S')} UTC")
    return triggered_at


def _fetch_from_imap(mail_pwd: str, triggered_at: datetime.datetime) -> dict:
    from email.utils import parsedate_to_datetime

    deadline = time.time() + WAIT_SECONDS
    print(f"⏳ Wachten op email van {VTTL_MAIL_FROM} (max {WAIT_SECONDS}s) …")

    while time.time() < deadline:
        try:
            m = imaplib.IMAP4_SSL(IMAP_SERVER, IMAP_PORT)
            m.login(IMAP_USER, mail_pwd)
            m.select(IMAP_FOLDER)

            status, ids = m.search(
                None,
                "FROM", f'"{VTTL_MAIL_FROM}"',
                "SUBJECT", f'"{VTTL_MAIL_SUBJ}"',
            )
            for mid in reversed(ids[0].split()):
                status, data = m.fetch(mid, "(RFC822)")
                msg = email.message_from_bytes(data[0][1])
                try:
                    mail_dt = parsedate_to_datetime(msg.get("Date", ""))
                    if mail_dt.timestamp() < triggered_at.timestamp() - 5:
                        continue
                except Exception:
                    pass
                found = {}
                for part in msg.walk():
                    fname = part.get_filename()
                    if fname and fname.endswith(".csv"):
                        found[fname] = part.get_payload(decode=True)
                if found:
                    m.logout()
                    return found

            m.logout()
        except Exception as e:
            print(f"  IMAP fout: {e}")

        remaining = int(deadline - time.time())
        if remaining > 0:
            print(f"  Nog geen email — opnieuw over {POLL_INTERVAL}s "
                  f"({remaining}s resterend) …")
            time.sleep(POLL_INTERVAL)

    sys.exit(
        f"✗ Geen nieuwe export-email ontvangen binnen {WAIT_SECONDS} seconden.\n"
        "  Controleer of het rapport correct getriggerd werd op leden.vttl.be."
    )


def _compare_persons(old_path: str, new_bytes: bytes) -> None:
    try:
        old = pd.read_csv(old_path, sep=";", encoding="unicode_escape",
                          dtype=str, low_memory=False)
        new = pd.read_csv(
            __import__("io").BytesIO(new_bytes),
            sep=";", encoding="latin-1", dtype=str, low_memory=False,
        )
    except Exception as e:
        print(f"  (vergelijking overgeslagen: {e})")
        return

    old_ids = set(old["Lidnummer"].dropna().str.strip())
    new_ids = set(new["Lidnummer"].dropna().str.strip())

    added   = new_ids - old_ids
    removed = old_ids - new_ids

    def _label(df, ids):
        rows = df[df["Lidnummer"].str.strip().isin(ids)]
        return [
            f"  {r['Voornaam'].strip()} {r['Naam'].strip()} "
            f"({r.get('Club (0)', '?').strip()})"
            for _, r in rows.iterrows()
        ]

    if added:
        print(f"\n🟢 Nieuw ({len(added)}):")
        for line in _label(new, added):
            print(line)
    else:
        print("\n🟢 Geen nieuwe leden.")

    if removed:
        print(f"\n🔴 Verwijderd ({len(removed)}):")
        for line in _label(old, removed):
            print(line)
    else:
        print("🔴 Geen verwijderde leden.")


def run():
    export_dir = os.getenv("EXPORT_TT")
    if not export_dir:
        sys.exit("✗ Omgevingsvariabele EXPORT_TT is niet ingesteld.")

    # EXPORT_TT kan een bestand zijn (exportPerson.csv) of een map
    if os.path.isdir(export_dir):
        person_path = os.path.join(export_dir, "exportPerson.csv")
        club_path   = os.path.join(export_dir, "exportClub.csv")
    else:
        person_path = export_dir
        club_path   = os.path.join(os.path.dirname(export_dir), "exportClub.csv")

    vttl_pwd, mail_pwd = _get_passwords()
    triggered_at       = _trigger_report(vttl_pwd)
    attachments        = _fetch_from_imap(mail_pwd, triggered_at)

    person_bytes = attachments.get("exportPerson.csv")
    club_bytes   = attachments.get("exportClub.csv")

    if not person_bytes:
        sys.exit("✗ exportPerson.csv niet gevonden in de email-bijlagen.")

    # Vergelijk met vorige versie vóór overschrijven
    if os.path.exists(person_path):
        print("\n📊 Vergelijking met vorige export:")
        _compare_persons(person_path, person_bytes)
    else:
        print("\nℹ Geen vorige export gevonden — vergelijking overgeslagen.")

    # Opslaan
    os.makedirs(os.path.dirname(os.path.abspath(person_path)), exist_ok=True)
    with open(person_path, "wb") as f:
        f.write(person_bytes)
    print(f"\n✓ {person_path}  ({len(person_bytes):,} bytes)")

    if club_bytes:
        with open(club_path, "wb") as f:
            f.write(club_bytes)
        print(f"✓ {club_path}  ({len(club_bytes):,} bytes)")
