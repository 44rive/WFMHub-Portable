from __future__ import annotations

import unittest

from wfmhub.report_packs import IMPLEMENTED_REPORT_PACK_KEYS, REPORT_PACKS, build_report_pack


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


if __name__ == "__main__":
    unittest.main()
