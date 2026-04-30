"""
ttpy_mailing_clubs
------------------
Haal mailinglijst op basis van clubnamen.
Config: mailing_clubs.toml (of opgegeven via --config)

Config-formaat:
    clubs    = ["Salamander", "Brasgata"]   # (gedeeltelijke) clubnamen
    functies = ["secretaris", "voorzitter"] # optioneel

Gebruik:
    ttpy_mailing_clubs [--config mailing_clubs.toml]
"""
import argparse
import tomllib
import re
import pandas as pd
from ttpy.mailroutines import getExport

OPTIONS_FUNCTIES = [
    'interclubJeugd', 'interclubSeniors', 'secretaris',
    'voorzitter', 'penningmeester', 'aanspreekpunt',
]


def _parse_list(value):
    """Verwerk een TOML-waarde als lijst: list, newline-gescheiden of komma-gescheiden string."""
    if isinstance(value, list):
        return value
    if '\n' in value:
        return [v.strip() for v in value.splitlines() if v.strip()]
    return [v.strip() for v in value.split(',') if v.strip()]


def strip_prefix(naam: str) -> str:
    """Verwijder gangbare clubprefixen voor fuzzy matching op clubnaam.

    Haalt prefixen zoals ``ttc``, ``kttc``, ``ttk``, ``omni`` en ``geelse``
    weg zodat zoekopdrachten op de kernnaam van de club werken.

    Args:
        naam (str): Volledige clubnaam.

    Returns:
        str: Clubnaam in kleine letters zonder prefix en witruimte.
    """
    return naam.lower().strip()


def run():
    """Lees de config, zoek clubs op en druk de mailinglijst af via stdout.

    Laadt het TOML-configuratiebestand (standaard ``mailing_clubs.toml``),
    zoekt clubs op naam (deelstring, hoofdletterongevoelig, prefix-tolerant)
    en drukt een kommagescheiden lijst van unieke e-mailadressen af.
    Niet-gevonden clubs worden gerapporteerd.

    Command-line argumenten:
        --config (str): Pad naar het TOML-configuratiebestand.
            Standaard ``'mailing_clubs.toml'``.
    """
    parser.add_argument('--config', default='mailing_clubs.toml',
                        help='Pad naar config bestand (default: mailing_clubs.toml)')
    args = parser.parse_args()

    with open(args.config, 'rb') as f:
        config = tomllib.load(f)

    functies = _parse_list(config.get('functies', ['secretaris', 'voorzitter']))
    clubs_invoer = _parse_list(config['clubs'])

    export, clubs = getExport()
    clubs['NaamStripped'] = clubs['NaamClub'].apply(strip_prefix)
    columns_emails = [f"{f} Emails" for f in functies]

    gevonden = []
    niet_gevonden = []
    for zoek in clubs_invoer:
        match = clubs[
            clubs['NaamStripped'].str.contains(zoek.lower(), case=False, regex=False) |
            clubs['NaamClub'].str.contains(zoek, case=False, regex=False)
        ]
        if match.empty:
            niet_gevonden.append(zoek)
        else:
            gevonden.append(match)

    if niet_gevonden:
        print(f"Niet gevonden clubs: {niet_gevonden}")

    if not gevonden:
        print("Geen clubs gevonden.")
        return

    sel = pd.concat(gevonden)
    print(f"\nGevonden clubs:")
    print(sel['NaamClub'].tolist())

    alle_mails = []
    for col in columns_emails:
        if col in sel.columns:
            for m in sel[col].dropna().astype(str):
                alle_mails.extend([x.strip() for x in m.split(',') if x.strip()])

    uniek = sorted(set(alle_mails))
    print(','.join(uniek))


if __name__ == '__main__':
    run()
