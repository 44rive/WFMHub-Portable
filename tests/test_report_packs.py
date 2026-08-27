from __future__ import annotations

import unittest

from wfmhub.report_packs import IMPLEMENTED_REPORT_PACK_KEYS, REPORT_PACKS, build_report_pack


class ReportPackTests(unittest.TestCase):
    def test_independent_report_packs_are_registered(self):
        self.assertEqual(IMPLEMENTED_REPORT_PACK_KEYS, ("operations", "intraday", "quality_pcs", "absence", "scorecard"))
        self.assertTrue(all(REPORT_PACKS[key].implemented for key in IMPLEMENTED_REPORT_PACK_KEYS))
        self.assertEqual(REPORT_PACKS["intraday"].default_folder, "intraday")
        self.assertEqual(REPORT_PACKS["quality_pcs"].default_folder, "quality_pcs")
        self.assertEqual(REPORT_PACKS["absence"].default_folder, "absence")
        self.assertEqual(REPORT_PACKS["scorecard"].default_folder, "scorecard")


if __name__ == "__main__":
    unittest.main()
