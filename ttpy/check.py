"""
ttpy_check
----------
Voer alle ledencontroles uit op de VTTL-ledendata van Provincie Antwerpen.

Controles: statuten, geboortedatums, dames op herensterktelijst en maximaal
één voorzitter, secretaris en penningmeester per club.
Na de controles worden de mailinglijsten per clubfunctie afgedrukt.
Een optionele check-config activeert de jaarboekpublicatie bij wijzigingen.
Zie ``check_routines.py`` voor de details per controle.

Gebruik:
    ttpy_check
"""
import argparse
import os
from pathlib import Path

import ttpy.check_routines as tc
import ttpy.mailinglijsten as ml
import ttpy.vttl_download as vd


def run():
    """Download de export, voer controles uit en print de mailinglijsten."""
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--config', type=Path, default=None,
                        help='Check-config met optionele [jaarboek]-sectie')
    args = parser.parse_args()
    configured = args.config or os.environ.get('TTPY_CHECK_CONFIG')
    config_path = Path(configured) if configured else Path('check.toml')
    yearbook = None
    if configured or config_path.exists():
        from ttpy.jaarboek import load_config
        yearbook = load_config(config_path)
        if yearbook is not None and yearbook.export is not None:
            export = os.environ.get('EXPORT_TT')
            if not export:
                parser.error('EXPORT_TT is vereist voor de download en ledencontroles.')
            def person_path(path):
                path = Path(path).resolve()
                return path / 'exportPerson.csv' if path.is_dir() else path
            if person_path(export) != person_path(yearbook.export):
                parser.error('[jaarboek].export moet dezelfde export gebruiken als EXPORT_TT.')
    vd.run()
    tc.checkall()
    ml.print_mailinglijsten()
    if yearbook is not None:
        from ttpy.jaarboek import update_yearbook
        update_yearbook(yearbook)


if __name__ == '__main__':
    run()
