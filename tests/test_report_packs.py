from __future__ import annotations

import unittest
import tempfile
from datetime import datetime
from pathlib import Path
from types import SimpleNamespace

from wfmhub.report_packs import (
    IMPLEMENTED_REPORT_PACK_KEYS,
    REPORT_PACKS,
    build_report_pack,
    publish_report,
    report_current_path,
)


class ReportPackTests(unittest.TestCase):
    def test_independent_report_packs_are_registered(self):
        self.assertEqual(IMPLEMENTED_REPORT_PACK_KEYS, (
            "pcs", "bonus", "service", "staffing", "attendance", "corrections", "absence",
        ))
        self.assertTrue(all(REPORT_PACKS[key].implemented for key in IMPLEMENTED_REPORT_PACK_KEYS))
        self.assertFalse(REPORT_PACKS["intraday"].implemented)
        self.assertEqual(REPORT_PACKS["pcs"].default_folder, "pcs")
        self.assertEqual(REPORT_PACKS["service"].default_folder, "service")
        self.assertEqual(REPORT_PACKS["attendance"].default_folder, "attendance")
        self.assertEqual(REPORT_PACKS["absence"].default_folder, "absence")
        self.assertEqual(REPORT_PACKS["corrections"].default_folder, "corrections")
        # Old command/API keys remain callable for backwards compatibility, but
        # are deliberately absent from the interactive product menu.
        self.assertTrue(REPORT_PACKS["operations"].implemented)
        self.assertTrue(REPORT_PACKS["quality_pcs"].implemented)
        self.assertFalse(REPORT_PACKS["scorecard"].implemented)

    def test_current_reports_are_flat_and_previous_copy_is_archived(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            config = SimpleNamespace(
                reports=root / "Reports",
                system=root / "_system",
            )
            current = report_current_path(config, "attendance")
            current.parent.mkdir(parents=True)
            current.write_bytes(b"old")
            partial = current.with_name("Attendance Callout.partial.xlsx")
            partial.write_bytes(b"new")
            publish_report(
                config, "attendance", partial, current,
                datetime(2026, 9, 4, 10, 30),
            )
            self.assertEqual(current.read_bytes(), b"new")
            archives = list((root / "Reports" / "Archive" / "2026-09-04").glob("*.xlsx"))
            self.assertEqual(len(archives), 1)
            self.assertEqual(archives[0].read_bytes(), b"old")
            self.assertEqual(current.name, "Attendance Callout.xlsx")
            self.assertEqual(report_current_path(config, "service").name, "OEM Flash.xlsx")


if __name__ == "__main__":
    unittest.main()
