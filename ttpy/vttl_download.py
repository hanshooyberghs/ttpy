"""
vttl_download.py
----------------
Triggert het ledenrapport op leden.vttl.be, wacht op de email met
CSV-bijlagen, slaat ze op en vergelijkt met de vorige export.
"""

import os
import sys
import io
import re
import time
import imaplib
import email
import datetime
import shutil
import requests
import pandas as pd

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
RECENT_WINDOW   = 1800   # accept emails up to 30 min old before triggering
# ─────────────────────────────────────────────────────────────────────────────


def _get_passwords():
    """Haal VTTL- en IMAP-wachtwoorden op uit omgevingsvariabelen of interactieve invoer.

    Gebruikt ``PASWOORD_TT`` voor het VTTL-webportaal en ``MAIL_PASSWORD`` voor
    de IMAP-mailbox.  Als ``MAIL_PASSWORD`` niet ingesteld is, valt de functie
    terug op ``PASWOORD_TT`` (wanneer beide hetzelfde zijn).

    Returns:
        tuple[str, str]: ``(vttl_wachtwoord, imap_wachtwoord)``.
    """
    mail_pwd = os.getenv("MAIL_PASSWORD") or vttl_pwd
    if not vttl_pwd:
        import getpass
        vttl_pwd = getpass.getpass("VTTL wachtwoord (PASWOORD_TT): ")
    if not mail_pwd:
        import getpass
        mail_pwd = getpass.getpass("IMAP wachtwoord (MAIL_PASSWORD): ")
    return vttl_pwd, mail_pwd


def _trigger_report(vttl_pwd: str) -> datetime.datetime:
    """Log in op leden.vttl.be en trigger het ledenlijst-rapport.

    Voert een HTTP-sessie uit: login, navigeer naar de rapportenpagina en
    dien het LedenLijst-rapport in voor provincie Antwerpen (alleen actieve leden).

    Args:
        vttl_pwd (str): Wachtwoord voor leden.vttl.be.

    Returns:
        datetime.datetime: UTC-tijdstip waarop het rapport getriggerd werd
        (gebruikt als tijdsanker voor het opwachten van de e-mail).

    Raises:
        SystemExit: Als de login mislukt.
    """
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
                "filterProvincie": "Antwerpen", "filterAansluitingsNrClub": "",
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


def _search_imap(mail_pwd: str, not_before: float) -> dict:
    """Return attachment dict from the most recent matching email after not_before."""
    from email.utils import parsedate_to_datetime
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
            if mail_dt.timestamp() < not_before:
                break   # messages are ordered oldest-first; no point checking further
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
    return {}


def _fetch_from_imap(mail_pwd: str, triggered_at: datetime.datetime) -> dict:
    """Wacht op de export-e-mail en retourneer de CSV-bijlagen.

    Pollt de IMAP-mailbox tot een e-mail van ``VTTL_MAIL_FROM`` met onderwerp
    ``VTTL_MAIL_SUBJ`` en CSV-bijlagen arriveert, of tot de wachttijd
    (``WAIT_SECONDS``) verstreken is.

    Args:
        mail_pwd (str): Wachtwoord voor de IMAP-mailbox.
        triggered_at (datetime.datetime): UTC-tijdstip van het triggeren van het
            rapport.  E-mails die ouder zijn dan ``RECENT_WINDOW`` seconden voor
            dit tijdstip worden genegeerd.

    Returns:
        dict[str, bytes]: Woordenboek ``{bestandsnaam: bestandsinhoud}`` met de
        CSV-bijlagen uit de e-mail.

    Raises:
        SystemExit: Als er na ``WAIT_SECONDS`` seconden geen geldige e-mail is
            ontvangen.
    """
    # Accept emails that arrived up to RECENT_WINDOW seconds before trigger
    not_before = triggered_at.timestamp() - RECENT_WINDOW
    print(f"⏳ Wachten op email van {VTTL_MAIL_FROM} (max {WAIT_SECONDS}s) …")

    while time.time() < deadline:
        try:
            found = _search_imap(mail_pwd, not_before)
            if found:
                return found
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
    """Vergelijk de nieuwe ledenexport met de vorige versie en rapporteer wijzigingen.

    Toont via stdout welke lidnummers nieuw zijn (🟢) of verwijderd zijn (🔴)
    ten opzichte van de opgeslagen export.  Bij een fout bij het inlezen van een
    van de bestanden wordt de vergelijking overgeslagen.

    Args:
        old_path (str): Pad naar het vorige CSV-exportbestand.
        new_bytes (bytes): Inhoud van het nieuwe CSV-exportbestand.
    """
    try:
        old = pd.read_csv(old_path, sep=";", encoding="unicode_escape",
                          dtype=str, low_memory=False)
        new = pd.read_csv(
            io.BytesIO(new_bytes),
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
    """Voer de volledige download-workflow uit voor de VTTL-ledenexport.

    Stappen:
        1. Bepaal de opslaglocatie op basis van omgevingsvariabele ``EXPORT_TT``.
        2. Haal wachtwoorden op (omgevingsvariabelen of interactieve invoer).
        3. Log in op leden.vttl.be en trigger het LedenLijst-rapport.
        4. Wacht op de export-e-mail en download de CSV-bijlagen.
        5. Vergelijk de nieuwe export met de vorige versie (indien aanwezig).
        6. Schrijf ``exportPerson.csv`` (en eventueel ``exportClub.csv``) weg.

    Raises:
        SystemExit: Als ``EXPORT_TT`` niet ingesteld is, de login mislukt,
            er geen e-mail ontvangen wordt, of ``exportPerson.csv`` ontbreekt
            in de bijlagen.
    """
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
