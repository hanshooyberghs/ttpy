"""
ttpy_check
----------
Voer alle ledencontroles uit op de VTTL-ledendata van Provincie Antwerpen.

Controles: statuten, geboortedatums en dames op herensterktelijst.
Zie ``check_routines.py`` voor de details per controle.

Gebruik:
    ttpy_check
"""
import ttpy.check_routines as tc
import ttpy.vttl_download as vd


def run():
    """Trigger eerst de VTTL-ledenexport, voer daarna alle ledencontroles uit."""
    vd.run()
    tc.checkall()


if __name__ == '__main__':
    run()
