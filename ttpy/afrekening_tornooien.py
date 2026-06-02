"""
ttpy_afrekening_tornooien
-------------------------
Genereer afrekeningen voor tornooien: inschrijvingsgelden + boetes.
Inschrijvingen en prijzen worden automatisch via de VTTL competitiesite CSV gedownload.

Gebruik:
    ttpy_afrekening_tornooien --cfg afrekening_tornooien.toml
"""
import argparse
import tomllib
from pathlib import Path

from ttpy.tournamentroutines import (
    AddFines, GetTournamentEntries, PrintTotals, TournamentsSaveAndMail
)


def run():
    """Lees de config en genereer de volledige tornooi-afrekening.

    Laadt het TOML-configuratiebestand (opgegeven via ``--cfg``), haalt
    inschrijvingen op via de VTTL competitiesite CSV, voegt boetes toe en
    slaat de resultaten op als Excel-bestanden. Optioneel worden mails
    verstuurd naar spelers en clubs.

    Command-line argumenten:
        --cfg (str): Pad naar het TOML-configuratiebestand.
    """
    parser = argparse.ArgumentParser(
        description='Genereer afrekeningen voor tornooien.')
    parser.add_argument('--cfg', required=True,
                        help='Pad naar het TOML-configuratiebestand')
    args = parser.parse_args()

    with open(args.cfg, 'rb') as f:
        config = tomllib.load(f)

    tornooien    = config['tornooien']
    clubs_direct = config['clubs_direct']
    season       = config.get('season')
    provincie    = config.get('provincie', 'A')
    send_mails   = config.get('send_mails', False)

    file_boetes         = config['file_boetes']
    lege_factuur        = config['lege_factuur']
    folder_out          = config['folder_out']
    formaat_factuur     = config['formaat_factuur']
    startnummer_factuur = config['startnummer_factuur']

    Path(folder_out).mkdir(parents=True, exist_ok=True)
    uitvoer_clubs        = f'{folder_out}/Afrekening_Tornooien_Clubs.xlsx'
    uitvoer_individueel  = f'{folder_out}/Afrekening_Tornooien_Individueel.xlsx'
    uitvoer_totaal       = f'{folder_out}/Afrekening_Tornooien_Totaal.xlsx'

    mail_cfg = config.get('mail', {})
    default_items = {
        'Mailindividueel': mail_cfg.get('mail_individueel', '').strip(),
        'MailClubs':       mail_cfg.get('mail_clubs', '').strip(),
        'UITZENDDATUM':    mail_cfg.get('uitzenddatum', ''),
        'VERVALDATUM':     mail_cfg.get('vervaldatum', ''),
    }

    lijst_inschrijvingen_persoon, lijst_inschrijvingen_totaal = GetTournamentEntries(
        tornooien, season=season, provincie=provincie, csv_dubbels=True)

    lijst_inschrijvingen_persoon = AddFines(
        lijst_inschrijvingen_persoon, file_boetes)

    PrintTotals(
        lijst_inschrijvingen_persoon, lijst_inschrijvingen_totaal, uitvoer_totaal)

    TournamentsSaveAndMail(
        lijst_inschrijvingen_persoon, clubs_direct, default_items,
        uitvoer_clubs, uitvoer_individueel, lege_factuur, tornooien,
        startnummer_factuur=startnummer_factuur,
        formaat_factuur=formaat_factuur,
        functies=['secretaris', 'penningmeester'],
        send_mails=send_mails,
        mail_test=False)


if __name__ == '__main__':
    run()
