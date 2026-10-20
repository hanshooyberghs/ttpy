"""
ttpy_mails_tornooi
------------------
Haal mailinglijst op van ingeschreven spelers in tornooien.
Config: mails_tornooi.toml (of opgegeven via --config)

Config-formaat:
    tornooien = ["Tornooi A", "Tornooi B"]
    functies  = ["secretaris"]

Inclusief beide dubbelspelers en gratis reeksen via de CSV-export.
Vereist ACCOUNT_TT, PASWOORD_TT en EXPORT_TT.
Toont onderaan spelers en paren die als enige in hun reeks zijn ingeschreven.
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


def _print_enige_inschrijvingen(resultaat):
    """Druk per reeks de enige speler of het enige paar af."""
    print("\nEnige inschrijvingen per reeks:")
    # TeamId bevat tornooi, reeks en inschrijvingsnummer.
    reeks_ids = resultaat['TeamId'].str.rsplit(':', n=1).str[0]
    gevonden = False
    for _, reeks in resultaat.groupby(reeks_ids, sort=False):
        if reeks['TeamId'].nunique() != 1:
            continue
        spelers = reeks.drop_duplicates('Lidnummer')
        namen = [
            f"{speler.Naam} {speler.Voornaam}".strip()
            for speler in spelers.itertuples()
        ]
        print(f"{' / '.join(namen)} ({reeks.iloc[0]['Reeks']})")
        gevonden = True
    if not gevonden:
        print('Geen.')


def run():
    """Lees de config, haal tornooi-inschrijvingen op en druk de mailinglijst af.

    Laadt het TOML-configuratiebestand (standaard ``mails_tornooi.toml``),
    haalt de registraties op via ``GetTournamentSeriesEntries()`` en drukt een
    kommagescheiden mailinglijst af op basis van de ingeschreven spelers.

    Command-line argumenten:
        --config (str): Pad naar het TOML-configuratiebestand.
            Standaard ``'mails_tornooi.toml'``.
    """
    parser = argparse.ArgumentParser(description='Haal mailing op van tornooi-inschrijvingen.')
    parser.add_argument('--config', default='mails_tornooi.toml',
                        help='Pad naar config bestand (default: mails_tornooi.toml)')
    args = parser.parse_args()

    with open(args.config, 'rb') as f:
        config = tomllib.load(f)

    tornooien = _parse_list(config['tornooien'])
    functies = _parse_list(config.get('functies', ['secretaris']))

    # Tel alle inschrijvingen; het provinciefilter geldt enkel voor de mailing.
    resultaat = tr.GetTournamentSeriesEntries(tornooien, provincie=None)
    antwerpen = resultaat[resultaat['Club'].fillna('').str.startswith('A')]
    lidnummers = antwerpen[['Lidnummer']].drop_duplicates()
    if lidnummers.empty:
        print('Geen ingeschreven spelers gevonden voor deze tornooien in provincie A.')
    else:
        mr.GetMailinglijst_Lidnummer(lidnummers, functies=functies)

    _print_enige_inschrijvingen(resultaat)


if __name__ == '__main__':
    run()
