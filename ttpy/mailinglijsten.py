"""Print mailinglijsten op basis van de VTTL-ledenexport."""

import csv
import os
from pathlib import Path


VAST_ADRES = "hans.hooyberghs@vito.be"

ROLES = {
    "Voorzitters": "voorzitter",
    "Secretarissen": "secretaris",
    "Penningmeesters": "penningmeester",
    "Interclubleiders": "interclubSeniors",
    "Jeugdverantwoordelijken": "interclubJeugd",
}


def _lees_export_emails(export_path: Path, functie: str) -> set[str]:
    emails = set()

    with export_path.open(encoding="latin-1") as bestand:
        reader = csv.DictReader(bestand, delimiter=";")
        functie_kolommen = [
            kolom for kolom in (reader.fieldnames or [])
            if kolom.startswith("Club Funtie (")
        ]

        for rij in reader:
            if not any(rij.get(kolom, "").strip() == functie for kolom in functie_kolommen):
                continue

            for kolom in ("Email", "Email (CC)"):
                for email in rij.get(kolom, "").split(","):
                    email = email.strip().lower()
                    if email and "@" in email and email != "null":
                        emails.add(email)

    return emails


def print_mailinglijsten(export_path: str | Path | None = None) -> None:
    """Print unieke e-mailadressen per clubfunctie en gecombineerd."""
    if export_path is None:
        export_path = os.getenv("EXPORT_TT")
    if not export_path:
        raise RuntimeError("Omgevingsvariabele EXPORT_TT is niet ingesteld.")

    export_path = Path(export_path)
    alle_emails = set()

    print("\n------------------")
    print("Mailinglijsten")
    print("------------------")

    for rol, functie in ROLES.items():
        emails = _lees_export_emails(export_path, functie) | {VAST_ADRES}
        alle_emails |= emails
        print(f"\n{rol} ({len(emails)} adressen)")
        print(",".join(sorted(emails)))

    print(f"\nNieuwsbrief - alle functies ({len(alle_emails)} adressen)")
    print(",".join(sorted(alle_emails)))


def run() -> None:
    print_mailinglijsten()


if __name__ == "__main__":
    run()
