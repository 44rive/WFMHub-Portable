"""Synthetic regressions; also run under the packaged Windows interpreter."""
import csv
import hashlib
import importlib.util
import os
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook

APP = Path(os.environ.get("MR_PROPRE_TEST_APP", Path(__file__).resolve().parents[1] / "mr_propre.py"))
spec = importlib.util.spec_from_file_location("mr_propre", APP)
mr = importlib.util.module_from_spec(spec)
spec.loader.exec_module(mr)


class FilterTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory(prefix="Mr Propre test ")
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.sources = self.root / "sources"
        self.sources.mkdir()
        self.output = self.root / "output"
        self.roster = self.root / "listing.csv"

    def write(self, name, text, encoding="utf-8"):
        path = self.sources / name
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_bytes(text.encode(encoding))
        return path

    def rows(self, path):
        with path.open(encoding="utf-8-sig", newline="") as handle:
            return list(csv.reader(handle))

    def test_one_column_listing_duplicates_blanks_and_text_ids(self):
        self.roster.write_text('Client ID\n00123\n123\n00123\n\n', encoding="utf-8")
        self.assertEqual(mr.load_roster(self.roster), {"00123", "123"})

    def test_membership_ignores_status_dates_and_org(self):
        self.roster.write_text('Client ID,Status,End date,LOB\n123,Inactive,2000-01-01,Other\n', encoding="utf-8")
        source = self.write("data.csv", 'Client ID,Date,Start Time,Value\n123,2026-11-01,09:00,keep\n456,2026-11-01,09:00,drop\n')
        self.assertEqual(mr.filter_file(source, self.output / "data.csv", mr.load_roster(self.roster)), (2, 1))
        self.assertEqual(self.rows(self.output / "data.csv")[1], ['123', '2026-11-01', '09:00', 'keep'])

    def test_original_fields_order_duplicates_multiline_and_sources_unchanged(self):
        source = self.write("data.csv", 'Value,Client ID,Note\n"x,y",00123,"line1\nline2"\n"x,y",00123,"line1\nline2"\nother,123,drop\n')
        before = hashlib.sha256(source.read_bytes()).digest()
        self.assertEqual(mr.filter_file(source, self.output / "data.csv", {"00123"}), (3, 2))
        self.assertEqual(self.rows(self.output / "data.csv"), [['Value', 'Client ID', 'Note'], ['x,y', '00123', 'line1\nline2'], ['x,y', '00123', 'line1\nline2']])
        self.assertEqual(hashlib.sha256(source.read_bytes()).digest(), before)

    def test_utf8_cp1252_utf16_and_delimiters(self):
        for i, (encoding, delimiter) in enumerate([("utf-8", "\t"), ("utf-8-sig", ";"), ("cp1252", "|"), ("utf-16", "\t")]):
            with self.subTest(encoding=encoding):
                source = self.write(f"{i}.txt", f'Data Source IDs{delimiter}Name\n123{delimiter}André\n', encoding)
                target = self.output / f"{i}.csv"
                mr.filter_file(source, target, {"123"})
                self.assertEqual(self.rows(target), [['Data Source IDs', 'Name'], ['123', 'André']])

    def test_late_non_utf8_accent(self):
        source = self.write("late.csv", 'Client ID,Note\n123,' + 'x' * 70000 + '\n123,André\n', "cp1252")
        target = self.output / "late.csv"
        self.assertEqual(mr.filter_file(source, target, {"123"}), (2, 2))
        self.assertEqual(self.rows(target)[-1], ['123', 'André'])

    def test_header_delimiter_not_data_punctuation(self):
        source = self.write("data.csv", 'Client ID;Note\n123;"a,b,c,d,e"\n')
        mr.filter_file(source, self.output / "data.csv", {"123"})
        self.assertEqual(self.rows(self.output / "data.csv")[1], ['123', 'a,b,c,d,e'])

    def test_client_id_precedence_and_ambiguous_aliases(self):
        self.assertEqual(mr.id_column(['Agent ID', 'Client ID']), 1)
        with self.assertRaisesRegex(ValueError, 'Multiple ID'):
            mr.id_column(['Agent ID', 'Data Source IDs'])

    def test_missing_id_and_malformed_rows_leave_no_csv(self):
        for i, text in enumerate(['Name,Value\nSynthetic,1\n', 'Client ID,Value\n123,ok\n456,too,many\n', 'Client ID,Value\n123,"unfinished']):
            with self.subTest(text=text):
                source = self.write(f"{i}.csv", text)
                target = self.output / f"{i}.csv"
                with self.assertRaises((ValueError, csv.Error)):
                    mr.filter_file(source, target, {"123"})
                self.assertFalse(target.exists())
                self.assertFalse(target.with_name(target.name + '.partial').exists())

    def test_never_overwrites_or_deletes_preexisting_partial(self):
        source = self.write('data.csv', 'Client ID\n123\n')
        self.output.mkdir()
        target = self.output / 'data.csv'
        target.write_text('existing')
        with self.assertRaises(ValueError):
            mr.filter_file(source, target, {'123'})
        self.assertEqual(target.read_text(), 'existing')
        partial = self.output / 'next.csv.partial'
        partial.write_text('existing partial')
        with self.assertRaises(FileExistsError):
            mr.filter_file(source, self.output / 'next.csv', {'123'})
        self.assertEqual(partial.read_text(), 'existing partial')

    def test_xlsx_agent_sheet_only_numeric_format_and_text(self):
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = 'Agent'
        sheet.append(['Synthetic listing'])
        sheet.append(['Client ID', 'Status'])
        sheet.append([123, 'Inactive'])
        sheet['A3'].number_format = '00000'
        sheet.append(['00099', 'Leaver'])
        pto = workbook.create_sheet('PTO')
        pto.append(['Client ID'])
        pto.append(['99999'])
        path = self.root / 'FTE Count.xlsx'
        workbook.save(path)
        workbook.close()
        self.assertEqual(mr.load_roster(path), {'00123', '00099'})

    def test_xlsx_ambiguous_sheet_requires_selection(self):
        workbook = Workbook()
        workbook.active.title = 'Roster'
        for sheet in [workbook.active, workbook.create_sheet('Agent')]:
            sheet.append(['Client ID'])
            sheet.append(['123'])
        path = self.root / 'roster.xlsx'
        workbook.save(path)
        workbook.close()
        with self.assertRaisesRegex(ValueError, 'Multiple roster sheets'):
            mr.load_roster(path)
        self.assertEqual(mr.load_roster(path, 'Agent'), {'123'})

    def test_selection_ranges_and_bounds(self):
        self.assertEqual(mr.parse_selection('1, 3-5,3', 5), [0, 2, 3, 4])
        self.assertEqual(mr.parse_selection('A', 3), [0, 1, 2])
        for value in ['', '0', '3-1', '2-1000000000', 'text', '1,']:
            with self.subTest(value=value), self.assertRaises(ValueError):
                mr.parse_selection(value, 5)

    def cli(self, *args, input=None):
        return subprocess.run([sys.executable, str(APP), '--folder', str(self.sources), '--roster', str(self.roster), '--output', str(self.output), *args], input=input, text=True, encoding='utf-8', capture_output=True)

    def test_cli_selected_nested_outputs_and_repeat_refusal(self):
        self.roster.write_text('Client ID\n00123\n')
        self.write('a.csv', 'Client ID,Value\n00123,yes\n123,no\n')
        self.write('sub/b.txt', 'Data Source IDs\tValue\n00123\tyes\n')
        self.write('z.csv', 'Client ID\n00123\n')
        result = self.cli('--select', '1-2')
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
        self.assertTrue((self.output / 'sub/b.csv').exists())
        self.assertFalse((self.output / 'z.csv').exists())
        result = self.cli('--all')
        self.assertEqual(result.returncode, 2)
        self.assertFalse((self.output / 'z.csv').exists())

    def test_cli_collisions_fail_before_writes(self):
        self.roster.write_text('Client ID\n123\n')
        self.write('data.csv', 'Client ID\n123\n')
        self.write('data.txt', 'Client ID\n123\n')
        result = self.cli('--all')
        self.assertEqual(result.returncode, 2)
        self.assertIn('collide', result.stderr)
        self.assertFalse(self.output.exists())

    def test_cli_output_inside_source_rejected(self):
        self.roster.write_text('Client ID\n123\n')
        result = self.cli('--output', str(self.sources / 'out'), '--all')
        self.assertEqual(result.returncode, 2)
        self.assertIn('separate', result.stderr)

    def test_cli_reports_file_failures_and_continues(self):
        self.roster.write_text('Client ID\n123\n')
        self.write('bad.csv', 'Name\nSynthetic\n')
        self.write('good.csv', 'Client ID\n123\n')
        result = self.cli('--all')
        self.assertEqual(result.returncode, 2)
        self.assertIn('1 exported, 1 failed', result.stdout)
        self.assertTrue((self.output / 'good.csv').exists())
        self.assertFalse((self.output / 'bad.csv').exists())

    def test_interactive_folder_listing_and_file_selection(self):
        self.write('a.csv', 'Client ID\n123\n')
        listings = self.root / 'listings'
        listings.mkdir()
        (listings / 'one.csv').write_text('Client ID\n123\n')
        (listings / 'two.csv').write_text('Client ID\n456\n')
        result = self.cli('--roster', str(listings), input='1\n1\n')
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
        self.assertEqual(self.rows(self.output / 'a.csv'), [['Client ID'], ['123']])

    @unittest.skipUnless(os.name == 'nt' and os.environ.get('MR_PROPRE_TEST_LAUNCHER'), 'Packaged Windows launcher check')
    def test_packaged_windows_launcher(self):
        self.roster.write_text('Client ID\n00123\n')
        self.write('a.csv', 'Client ID,Name\n00123,Synthetic Agent\n99999,Excluded Agent\n')
        result = subprocess.run([os.environ['MR_PROPRE_TEST_LAUNCHER'], '--folder', str(self.sources), '--roster', str(self.roster), '--output', str(self.output), '--all'], shell=True, text=True, capture_output=True)
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
        self.assertEqual(self.rows(self.output / 'a.csv'), [['Client ID', 'Name'], ['00123', 'Synthetic Agent']])


if __name__ == '__main__':
    unittest.main(verbosity=2)
