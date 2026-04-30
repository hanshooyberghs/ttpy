"""
ttpy_inschrijving_tornooi
-------------------------
Schrijf spelers in voor een tornooi via de VTTL API.
Config: inschrijving_tornooi.toml (of opgegeven via --config)

Config-formaat:
    tornooi    = "Naam van het tornooi"
    mail       = true
    unregister = false

    [[inschrijvingen]]
    reeks  = "Heren B"
    namen  = "Jansen Piet"                         # enkeling: string
    dubbel = false

    [[inschrijvingen]]
    reeks  = "Dubbel Gemengd -19"
    namen  = ["Michielsen Jente", "Belmans Maithe"] # dubbel: lijst van 2 namen
    dubbel = true
"""
import argparse
import tomllib
import ttpy.tournamentroutines as tr


def run():
    """Lees de config en schrijf spelers in voor het opgegeven tornooi.

    Laadt het TOML-configuratiebestand (standaard
    ``inschrijving_tornooi.toml``), itereert over de inschrijvingen en roept
    ``InschrijvingNaam()`` aan per speler of dubbelpaar.

    Voor elk inschrijvings-item in de config:
        - Een enkelvoudige naam of kommalijst van namen wordt per speler
          afzonderlijk ingeschreven (``dubbel=False``).
        - Een lijst van twee namen met ``dubbel=true`` wordt als dubbelpaar
          ingeschreven.

    Command-line argumenten:
        --config (str): Pad naar het TOML-configuratiebestand.
            Standaard ``'inschrijving_tornooi.toml'``.
    """
    parser = argparse.ArgumentParser(description='Schrijf spelers in voor een tornooi.')
    parser.add_argument('--config', default='inschrijving_tornooi.toml',
                        help='Pad naar config bestand (default: inschrijving_tornooi.toml)')
    args = parser.parse_args()

    with open(args.config, 'rb') as f:
        config = tomllib.load(f)

    tornooi = config['tornooi']
    mail = config.get('mail', True)
    unregister = config.get('unregister', False)

    for item in config.get('inschrijvingen', []):
        reeks = item['reeks']
        namen_raw = item['namen']
        dubbel = item.get('dubbel', False)

        if isinstance(namen_raw, list):
            namen = namen_raw
        elif '\n' in namen_raw:
            namen = [n.strip() for n in namen_raw.splitlines() if n.strip()]
        else:
            namen = [n.strip() for n in namen_raw.split(',') if n.strip()]
            if len(namen) == 1:
                namen = namen[0]  # enkeling als string

        if dubbel:
            tr.InschrijvingNaam(namen, tornooi, reeks=reeks, mail=mail, dubbel=True, unregister=unregister)
        else:
            if isinstance(namen, list):
                for naam in namen:
                    tr.InschrijvingNaam(naam, tornooi, reeks=reeks, mail=mail, dubbel=False, unregister=unregister)
            else:
                tr.InschrijvingNaam(namen, tornooi, reeks=reeks, mail=mail, dubbel=False, unregister=unregister)


if __name__ == '__main__':
    run()
