"""
ttpy_mailing_naam
-----------------
Haal mailinglijst op basis van een lijst namen.
Config: mailing_naam.toml (of opgegeven via --config)

Config-formaat:
    namen   = ["Achternaam Voornaam", ...]  # of newline-/kommagescheiden string
    functies = ["secretaris", "voorzitter"]  # optioneel, standaard leeg

Gebruik:
    ttpy_mailing_naam [--config mailing_naam.toml]
"""
import argparse
import tomllib
import pandas as pd
import ttpy.mailroutines as mr


def _parse_list(value):
    """Verwerk een TOML-waarde als lijst: list, newline-gescheiden of komma-gescheiden string."""
    if isinstance(value, list):
        return value
    if '\n' in value:
        return [v.strip() for v in value.splitlines() if v.strip()]
    return [v.strip() for v in value.split(',') if v.strip()]


def run():
    """Lees de config, haal namen op en druk de mailinglijst af via stdout.

    Laadt het TOML-configuratiebestand (standaard ``mailing_naam.toml``),
    parseert de namenlijst en optionele functies, en roept
    ``GetMailinglijst_Naam()`` aan.

    Command-line argumenten:
        --config (str): Pad naar het TOML-configuratiebestand.
            Standaard ``'mailing_naam.toml'``.
    """
    parser.add_argument('--config', default='mailing_naam.toml',
                        help='Pad naar config bestand (default: mailing_naam.toml)')
    args = parser.parse_args()

    with open(args.config, 'rb') as f:
        config = tomllib.load(f)

    namen = _parse_list(config['namen'])
    functies = _parse_list(config.get('functies', []))

    mr.GetMailinglijst_Naam(pd.DataFrame(namen, columns=['Naam']), functies=functies)


if __name__ == '__main__':
    run()
