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


def run():
    """Voer alle ledencontroles uit via ``check_routines.checkall()``.

    Roept ``checkall()`` aan, dat achtereenvolgens ``checkStatuten()``,
    ``checkGeboortedatum()`` en ``checkDamesOpHeren()`` uitvoert en de
    resultaten afdrukt via stdout.
    """


if __name__ == '__main__':
    run()
