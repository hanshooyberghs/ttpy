"""
ttpy_check
----------
Voer alle ledencontroles uit op de VTTL-ledendata van Provincie Antwerpen.

Controles: statuten, geboortedatums en dames op herensterktelijst.
Na de controles worden de mailinglijsten per clubfunctie afgedrukt.
Zie ``check_routines.py`` voor de details per controle.

Gebruik:
    ttpy_check
"""
import ttpy.check_routines as tc
import ttpy.mailinglijsten as ml
import ttpy.vttl_download as vd


def run():
    """Download de export, voer controles uit en print de mailinglijsten."""
    vd.run()
    tc.checkall()
    ml.print_mailinglijsten()


if __name__ == '__main__':
    run()
