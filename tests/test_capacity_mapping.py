from __future__ import annotations

import shutil
import tempfile
import unittest
from pathlib import Path

from wfmhub.capacity_mapping import ensure_capacity_mapping, load_capacity_mapping


REPO = Path(__file__).resolve().parents[1]


class CapacityMappingTests(unittest.TestCase):
    def test_forecast_staff_types_map_to_planning_grain(self):
        mapping = load_capacity_mapping(REPO / "config" / "default_capacity_mapping.csv")
        checks = (
            ("RSA_BE_FR_VOL&FTE_09-2026.txt", "BE RSA FO FR", ("RSA BE", "RSA BE FR")),
            ("RSA_BE_VL_VOL&FTE_09-2026.txt", "BE RSA FO VL", ("RSA BE", "RSA BE VL")),
            ("FORD_FR_VOL&FTE_09-2026.txt", "FR RSA Ford Level 1", ("OEM", "Ford FR")),
        )
        for filename, staff_type, expected in checks:
            result = mapping.map_forecast(filename, staff_type)
            self.assertEqual((result.management_lob, result.planning_group), expected)
            self.assertEqual(result.status, "MAPPED")

    def test_queue_labels_are_not_accepted_as_staff_types(self):
        mapping = load_capacity_mapping(REPO / "config" / "default_capacity_mapping.csv")
        result = mapping.map_forecast(
            "RSA_BE_FR_VOL&FTE_09-2026.txt",
            "APBN_BRU_MOBILITY_INSURAN_Front_FR",
        )
        self.assertEqual(result.status, "UNMAPPED_STAFF_TYPE")
        self.assertEqual(result.planning_group, "RSA BE FR")

    def test_copy_on_first_run_never_rewrites_user_mapping(self):
        with tempfile.TemporaryDirectory() as folder:
            home = Path(folder)
            (home / "config").mkdir()
            shutil.copy2(
                REPO / "config" / "default_capacity_mapping.csv",
                home / "config" / "default_capacity_mapping.csv",
            )
            target = home / "config" / "capacity_mapping.csv"
            content = (
                "management_lob,planning_group,workforce_lob,forecast_file_prefix,"
                "forecast_staff_type,schedule_assignment,staff_type\n"
                "CUSTOM,Custom Group,Custom LOB,CUSTOM,Custom Staff,Custom Shift,Custom Staff\n"
            )
            target.write_text(content, encoding="utf-8")
            ensure_capacity_mapping(home, target)
            self.assertEqual(target.read_text(encoding="utf-8"), content)


if __name__ == "__main__":
    unittest.main()
