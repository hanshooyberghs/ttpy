import io
import unittest
from contextlib import redirect_stdout
from unittest.mock import patch

import pandas as pd

from ttpy import check_routines


class ClubfunctiesTests(unittest.TestCase):
    def check_output(self, rows):
        columns = [
            'Lidnummer', 'Naam', 'Voornaam', 'Club (0)', 'NaamClub',
            'Club Funtie (0)', 'Club Funtie (1)', 'Club Funtie (2)',
        ]
        data = pd.DataFrame(rows, columns=columns)
        output = io.StringIO()
        with patch.object(check_routines, 'loadData', return_value=data):
            with redirect_stdout(output):
                check_routines.checkClubfuncties()
        return output.getvalue()

    def test_reports_each_role_in_any_function_column(self):
        for role in ('voorzitter', 'secretaris', 'penningmeester'):
            with self.subTest(role=role):
                output = self.check_output([
                    [1, 'Eerste', 'Lid', 'A001', 'Club Een', None, role, None],
                    [2, 'Tweede', 'Lid', 'A001', 'Club Een', None, None, role],
                    [3, 'Derde', 'Lid', 'A002', 'Club Twee', role, None, None],
                ])
                self.assertIn(f'Club A001: 2 leden met functie {role}', output)
                self.assertIn('Eerste', output)
                self.assertIn('Tweede', output)
                self.assertNotIn('Derde', output)

    def test_member_counts_once_per_club_and_role(self):
        row = [1, 'Eerste', 'Lid', 'A001', 'Club Een', 'voorzitter', 'voorzitter', 'secretaris']
        output = self.check_output([
            row, row,
            [2, 'Tweede', 'Lid', 'A002', 'Club Twee', 'voorzitter', None, None],
            [3, 'Derde', 'Lid', 'A001', 'Club Een', 'interclubJeugd', None, None],
        ])
        self.assertIn('Geen clubs met meerdere', output)

    def test_empty_export_has_no_conflicts(self):
        self.assertIn('Geen clubs met meerdere', self.check_output([]))

    def test_missing_function_columns_is_explicit_error(self):
        with patch.object(check_routines, 'loadData', return_value=pd.DataFrame()):
            with redirect_stdout(io.StringIO()):
                with self.assertRaisesRegex(ValueError, 'clubfunctiekolommen'):
                    check_routines.checkClubfuncties()

    def test_checkall_includes_club_functions(self):
        with patch.object(check_routines, 'checkStatuten'), \
                patch.object(check_routines, 'checkGeboortedatum'), \
                patch.object(check_routines, 'checkDamesOpHeren'), \
                patch.object(check_routines, 'checkClubfuncties') as check:
            check_routines.checkall()
        check.assert_called_once_with()


if __name__ == '__main__':
    unittest.main()
