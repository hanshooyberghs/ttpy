import unittest

from ttpy.tournamentroutines import _parse_tournament_csv


class TournamentCsvTests(unittest.TestCase):
    def test_parser_preserves_series_types_teams_and_free_series(self):
        csv_text = '\n'.join([
            'R;10;Enkel Jeugd;;5;single',
            'R;20;Dubbel Jeugd;;0;double',
            'R;30;Dubbel Gemengd;;5;mixed',
            'I;10;101',
            'I;10;102',
            'I;20;101;103',
            'I;30;102;104',
        ])

        entries = _parse_tournament_csv(
            csv_text, 99, 'Jeugdtornooi', include_free=True,
            strict_structure=True)

        self.assertEqual(len(entries), 6)
        self.assertEqual(
            set(entries['ReeksType']), {'single', 'double', 'mixed'})
        double = entries[entries['Reeks'] == 'Dubbel Jeugd']
        self.assertEqual(double['TeamId'].nunique(), 1)
        self.assertEqual(set(double['Lidnummer']), {101, 103})

    def test_billing_mode_excludes_free_series(self):
        csv_text = '\n'.join([
            'R;10;Enkel Jeugd;;5;single',
            'R;20;Gratis Dubbel;;0;double',
            'I;10;101',
            'I;20;101;102',
        ])

        entries = _parse_tournament_csv(
            csv_text, 99, 'Jeugdtornooi', include_free=False)

        self.assertEqual(entries['Reeks'].tolist(), ['Enkel Jeugd'])

    def test_incomplete_double_is_rejected(self):
        csv_text = '\n'.join([
            'R;20;Dubbel Jeugd;;5;double',
            'I;20;101',
        ])

        with self.assertRaises(ValueError):
            _parse_tournament_csv(
                csv_text, 99, 'Jeugdtornooi', include_free=True,
                strict_structure=True)


if __name__ == '__main__':
    unittest.main()
