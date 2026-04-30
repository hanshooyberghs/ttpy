"""
ttpy_mails_tornooi
------------------
Haal mailinglijst op van ingeschreven spelers in tornooien.
Config: mails_tornooi.toml (of opgegeven via --config)

Config-formaat:
    tornooien = ["Tornooi A", "Tornooi B"]
    functies  = ["secretaris"]

    # Vast inschrijvingsgeld voor alle tornooien:
    inschrijvingsgeld = 5

    # Of per tornooi:
    # [inschrijvingsgeld]
    # "Tornooi A" = 5
    # "Tornooi B" = 8
"""
import argparse
import tomllib
import ttpy.tournamentroutines as tr
import ttpy.mailroutines as mr


def _parse_list(value):
    """Verwerk een TOML-waarde als lijst: list, newline-gescheiden of komma-gescheiden string."""
    if isinstance(value, list):
        return value
    if '\n' in value:
        return [v.strip() for v in value.splitlines() if v.strip()]
    return [v.strip() for v in value.split(',') if v.strip()]


def run():
    parser = argparse.ArgumentParser(description='Haal mailing op van tornooi-inschrijvingen.')
    parser.add_argument('--config', default='mails_tornooi.toml',
                        help='Pad naar config bestand (default: mails_tornooi.toml)')
    args = parser.parse_args()

    with open(args.config, 'rb') as f:
        config = tomllib.load(f)

    tornooien = _parse_list(config['tornooien'])
    functies = _parse_list(config.get('functies', ['secretaris']))
    inschrijvingsgeld = config.get('inschrijvingsgeld', 0)

    resultaat, _ = tr.GetTournamentEntries(tornooien, inschrijvingsgeld, provincie='A')
    mr.GetMailinglijst_Lidnummer(resultaat[['Lidnummer']], functies=functies)


if __name__ == '__main__':
    run()
