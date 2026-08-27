from __future__ import annotations

import unittest

from wfmhub.report_packs import IMPLEMENTED_REPORT_PACK_KEYS, REPORT_PACKS, build_report_pack


class ReportPackTests(unittest.TestCase):
    def test_operations_is_only_implemented_pack_and_pcs_is_reserved(self):
        self.assertEqual(IMPLEMENTED_REPORT_PACK_KEYS, ("operations",))
        self.assertTrue(REPORT_PACKS["operations"].implemented)
        self.assertFalse(REPORT_PACKS["quality_pcs"].implemented)
        with self.assertRaisesRegex(ValueError, "reserved but not implemented"):
            build_report_pack("quality_pcs", None, None, None, None)


if __name__ == "__main__":
    unittest.main()
