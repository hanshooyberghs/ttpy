import csv
import io
import json
import os
import tempfile
import unittest
from contextlib import redirect_stdout
from dataclasses import replace
from pathlib import Path
from unittest.mock import patch

from ttpy import check, jaarboek


class YearbookTests(unittest.TestCase):
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory()
        self.addCleanup(self.directory.cleanup)
        self.base = Path(self.directory.name)
        self.config = jaarboek.YearbookConfig(
            vaste_docx=self.base / 'fixed.docx',
            club_pdf=self.base / 'clubs.pdf',
            jaarboek_pdf=self.base / 'book.pdf',
            export=self.base,
            seizoen='2026-2027',
        )
        self.people = [
            {'Lidnummer': '1', 'Naam': 'Peeters', 'Voornaam': 'Jan',
             'Club (0)': 'A001', 'Club Funtie (0)': 'secretaris',
             'Email': 'jan@example.test', 'Email (CC)': '',
             'GSM': '012345', 'Telefoon': '98765', 'Straat': 'Straat',
             'Huisnummer': '1', 'Busnummer': '', 'Postnummer': '2000',
             'Gemeente': 'Antwerpen', 'Statuut (0)': 'competitief lid'},
            {'Lidnummer': '2', 'Naam': 'Anders', 'Voornaam': 'Lid',
             'Club (0)': 'A001', 'Club Funtie (0)': '',
             'Email': 'ander@example.test'},
        ]
        self.clubs = [
            {'AansluitingsNummer': 'A001', 'Naam': 'Club Een', 'Actief?': '1',
             'Lokaal naam (0)': 'Zaal', 'Lokaal Straat (0)': 'Sportstraat',
             'Lokaal Huisnummer (0)': '2'},
        ]
        self.write_exports()
        self.addCleanup(patch.stopall)
        patch.object(jaarboek, 'convert_static_docx', return_value=self.base / 'fixed.pdf').start()
        patch.object(jaarboek, 'create_pdf', side_effect=lambda clubs, season, output: output.write_bytes(b'clubs')).start()
        self.merge = patch.object(
            jaarboek, 'merge_pdf',
            side_effect=lambda fixed, clubs, output: output.write_bytes(b'book'),
        ).start()

    def write_exports(self):
        for filename, rows in [('exportPerson.csv', self.people), ('exportClub.csv', self.clubs)]:
            columns = list(dict.fromkeys(k for row in rows for k in row))
            with (self.base / filename).open('w', encoding='latin-1', newline='') as handle:
                writer = csv.DictWriter(handle, fieldnames=columns, delimiter=';')
                writer.writeheader()
                writer.writerows(rows)

    def update(self, **kwargs):
        with redirect_stdout(io.StringIO()):
            return jaarboek.update_yearbook(self.config, **kwargs)

    def test_first_publication_and_unchanged_export(self):
        self.assertTrue(self.update())
        self.assertFalse(self.update())
        self.assertEqual(self.merge.call_count, 1)

    def test_public_club_pdf_is_created_updated_and_repaired(self):
        self.config = replace(self.config, jaarboek_pdf=self.base / 'public' / 'book.pdf')
        public_clubs = self.config.jaarboek_pdf.parent / self.config.club_pdf.name
        self.update()
        self.assertEqual(public_clubs.read_bytes(), self.config.club_pdf.read_bytes())
        public_clubs.write_bytes(b'outdated')
        self.assertFalse(self.update())
        self.assertEqual(public_clubs.read_bytes(), self.config.club_pdf.read_bytes())
        public_clubs.unlink()
        self.assertFalse(self.update())
        self.assertTrue(public_clubs.is_file())
        self.assertEqual(self.merge.call_count, 1)
        jaarboek.create_pdf.side_effect = lambda clubs, season, output: output.write_bytes(b'new clubs')
        self.people[0]['Email'] = 'changed@example.test'
        self.write_exports()
        self.assertTrue(self.update())
        self.assertEqual(public_clubs.read_bytes(), b'new clubs')

    def test_public_copy_refuses_to_overwrite_input(self):
        source = self.base / 'source' / 'fixed.pdf'
        source.parent.mkdir()
        source.write_bytes(b'clubs')
        with self.assertRaisesRegex(ValueError, 'overschrijven'):
            jaarboek.publish_club_pdf(source, self.base / 'book.pdf',
                                     protected=(self.base / 'fixed.pdf',))

    def test_output_has_heading_and_whitespace_before_status(self):
        heading = '\n--------------------\nJaarboek publiceren\n--------------------\n\n'
        for expected in ('Jaarboek: bestuurs-/clubgegevens gewijzigd',
                         'Jaarboek: geen wijzigingen'):
            with self.subTest(status=expected):
                output = io.StringIO()
                with redirect_stdout(output):
                    jaarboek.update_yearbook(self.config)
                self.assertTrue(output.getvalue().startswith(heading + expected))

    def test_unrelated_member_and_statute_changes_are_ignored(self):
        self.update()
        self.people[1]['Email'] = 'nieuw@example.test'
        self.people[0]['Statuut (0)'] = 'recreant-reserve'
        self.people.reverse()
        self.write_exports()
        self.assertFalse(self.update())

    def test_contact_fields_and_roles_trigger_publication(self):
        self.update()
        for field, value in [
            ('Email', 'nieuw@example.test'), ('GSM', '123456'),
            ('Straat', 'Nieuwstraat'), ('Naam', 'Maes'),
            ('Email (CC)', 'cc@example.test'), ('Club Funtie (0)', 'voorzitter'),
            ('Club Funtie (0)', ''),
        ]:
            with self.subTest(field=field, value=value):
                self.people[0][field] = value
                self.write_exports()
                self.assertTrue(self.update())

    def test_club_name_venue_and_active_status_trigger_publication(self):
        self.update()
        for field, value in [('Naam', 'Club Twee'), ('Lokaal naam (0)', 'Nieuwe zaal')]:
            self.clubs[0][field] = value
            self.write_exports()
            self.assertTrue(self.update())
        self.clubs.append({'AansluitingsNummer': 'A002', 'Naam': 'Club Drie', 'Actief?': '1'})
        self.write_exports()
        self.assertTrue(self.update())
        self.clubs[1]['Actief?'] = '0'
        self.write_exports()
        self.assertTrue(self.update())

    def test_failure_preserves_snapshot_and_next_run_retries(self):
        self.update()
        snapshot = self.config.club_pdf.with_suffix('.json')
        previous = snapshot.read_bytes()
        self.people[0]['Email'] = 'nieuw@example.test'
        self.write_exports()
        original = self.merge.side_effect
        self.merge.side_effect = RuntimeError('Conversie mislukt')
        with self.assertRaisesRegex(RuntimeError, 'Conversie mislukt'):
            self.update()
        self.assertEqual(snapshot.read_bytes(), previous)
        self.merge.side_effect = original
        self.assertTrue(self.update())
        self.assertFalse(self.update())

    def test_missing_pdf_force_and_club_only_snapshot_rebuild(self):
        self.update()
        self.config.jaarboek_pdf.unlink()
        self.assertTrue(self.update())
        self.assertTrue(self.update(force=True))
        snapshot = self.config.club_pdf.with_suffix('.json')
        manifest = json.loads(snapshot.read_text())
        manifest['alleen_clubpagina'] = True
        snapshot.write_text(json.dumps(manifest))
        self.assertTrue(self.update())

    def test_manual_static_pdf_skips_conversion(self):
        self.config = replace(self.config, vaste_pdf=self.base / 'manual.pdf')
        self.update()
        jaarboek.convert_static_docx.assert_not_called()

    def test_corrupt_snapshot_is_explicit_error(self):
        self.config.club_pdf.with_suffix('.json').write_text('{}')
        with self.assertRaisesRegex(ValueError, 'Ongeldig jaarboek'):
            self.update()


class ConfigTests(unittest.TestCase):
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory()
        self.addCleanup(self.directory.cleanup)
        self.base = Path(self.directory.name)
        self.path = self.base / 'check.toml'

    def write_config(self, extra=''):
        self.path.write_text(
            '[jaarboek]\nvaste_docx = "fixed.docx"\n'
            'club_pdf = "clubs.pdf"\njaarboek_pdf = "book.pdf"\n' + extra
        )

    def test_paths_are_relative_to_config(self):
        self.write_config()
        config = jaarboek.load_config(self.path)
        self.assertEqual(config.club_pdf, self.base / 'clubs.pdf')
        self.assertIsNone(config.export)

    def test_empty_config_disables_yearbook(self):
        self.path.write_text('')
        self.assertIsNone(jaarboek.load_config(self.path))

    def test_invalid_config_is_rejected(self):
        for extra in ['onbekend = true\n', 'seizoen = false\n', 'seizoen = "fout"\n']:
            with self.subTest(extra=extra):
                self.write_config(extra)
                with self.assertRaises(ValueError):
                    jaarboek.load_config(self.path)
        self.path.write_text('[jaarboek]\nclub_pdf = "clubs.pdf"\n')
        with self.assertRaisesRegex(ValueError, 'verplicht'):
            jaarboek.load_config(self.path)
        self.path.write_text(
            '[jaarboek]\nvaste_docx = "fixed.docx"\n'
            'club_pdf = "book.pdf"\njaarboek_pdf = "book.pdf"\n'
        )
        with self.assertRaisesRegex(ValueError, 'overschrijven'):
            jaarboek.load_config(self.path)


class CheckIntegrationTests(unittest.TestCase):
    def test_hook_runs_after_download_checks_and_mailing(self):
        events = []
        config = jaarboek.YearbookConfig(
            Path('fixed.docx'), Path('clubs.pdf'), Path('book.pdf'),
        )
        with patch('sys.argv', ['ttpy_check', '--config', 'check.toml']), \
                patch.object(jaarboek, 'load_config', return_value=config), \
                patch.object(check.vd, 'run', side_effect=lambda: events.append('download')), \
                patch.object(check.tc, 'checkall', side_effect=lambda: events.append('checks')), \
                patch.object(check.ml, 'print_mailinglijsten', side_effect=lambda: events.append('mailing')), \
                patch.object(jaarboek, 'update_yearbook', side_effect=lambda c: events.append(c)):
            check.run()
        self.assertEqual(events, ['download', 'checks', 'mailing', config])

    def test_unconfigured_check_keeps_existing_behavior(self):
        with patch('sys.argv', ['ttpy_check']), \
                patch.dict(os.environ, {}, clear=True), \
                patch.object(check.Path, 'exists', return_value=False), \
                patch.object(check.vd, 'run') as download, \
                patch.object(check.tc, 'checkall') as checks, \
                patch.object(check.ml, 'print_mailinglijsten') as mailing, \
                patch.object(jaarboek, 'update_yearbook') as update:
            check.run()
        download.assert_called_once()
        checks.assert_called_once()
        mailing.assert_called_once()
        update.assert_not_called()

    def test_bad_config_stops_before_downloading(self):
        with patch('sys.argv', ['ttpy_check', '--config', 'missing.toml']), \
                patch.object(jaarboek, 'load_config', side_effect=FileNotFoundError), \
                patch.object(check.vd, 'run') as download:
            with self.assertRaises(FileNotFoundError):
                check.run()
        download.assert_not_called()

    def test_environment_selects_config(self):
        with patch('sys.argv', ['ttpy_check']), \
                patch.dict(os.environ, {'TTPY_CHECK_CONFIG': '/config/check.toml'}), \
                patch.object(jaarboek, 'load_config', return_value=None) as loader, \
                patch.object(check.vd, 'run'), \
                patch.object(check.tc, 'checkall'), \
                patch.object(check.ml, 'print_mailinglijsten'):
            check.run()
        loader.assert_called_once_with(Path('/config/check.toml'))

    def test_mismatched_export_stops_before_download(self):
        config = jaarboek.YearbookConfig(
            Path('fixed.docx'), Path('clubs.pdf'), Path('book.pdf'),
            export=Path('/different/exportPerson.csv'),
        )
        with patch('sys.argv', ['ttpy_check', '--config', 'check.toml']), \
                patch.dict(os.environ, {'EXPORT_TT': '/original/exportPerson.csv'}), \
                patch.object(jaarboek, 'load_config', return_value=config), \
                patch.object(check.vd, 'run') as download, \
                redirect_stdout(io.StringIO()), patch('sys.stderr', io.StringIO()):
            with self.assertRaises(SystemExit):
                check.run()
        download.assert_not_called()

    def test_failed_download_does_not_publish(self):
        config = jaarboek.YearbookConfig(
            Path('fixed.docx'), Path('clubs.pdf'), Path('book.pdf'),
        )
        with patch('sys.argv', ['ttpy_check', '--config', 'check.toml']), \
                patch.object(jaarboek, 'load_config', return_value=config), \
                patch.object(check.vd, 'run', side_effect=RuntimeError('Download mislukt')), \
                patch.object(jaarboek, 'update_yearbook') as update:
            with self.assertRaisesRegex(RuntimeError, 'Download mislukt'):
                check.run()
        update.assert_not_called()


if __name__ == '__main__':
    unittest.main()
