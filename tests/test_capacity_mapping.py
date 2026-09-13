from __future__ import annotations

import shutil
import tempfile
import unittest
from pathlib import Path

from wfmhub.capacity_mapping import ensure_capacity_mapping, load_capacity_mapping
from wfmhub.ingestion import parse_forecast


REPO = Path(__file__).resolve().parents[1]


class CapacityMappingTests(unittest.TestCase):
    def test_every_supplied_forecast_staff_type_has_an_exact_capacity_mapping(self):
        mapping = load_capacity_mapping(REPO / "config" / "default_capacity_mapping.csv")
        files = sorted((REPO / "attachments").glob("*VOL&FTE*.txt"))
        self.assertEqual(len(files), 5)
        seen: set[tuple[str, str, str]] = set()
        for path in files:
            rows = parse_forecast(path, path.name).tables["raw.forecast_interval"]
            for source_staff_type in {str(row["queue_name"]) for row in rows}:
                mapped = mapping.map_forecast(path.name, source_staff_type)
                with self.subTest(file=path.name, staff_type=source_staff_type):
                    self.assertEqual(mapped.status, "MAPPED")
                seen.add((mapped.management_lob, mapped.planning_group, mapped.staff_type))
        self.assertIn(("RSA BE", "RSA BE FR", "BE RSA FO FR"), seen)
        self.assertIn(("RSA BE", "RSA BE VL", "BE RSA FO VL"), seen)
        self.assertIn(("OEM", "Ford FR", "FR RSA Ford Level 1"), seen)

    def test_service_queues_are_not_used_to_map_forecast_staff_types(self):
        mapping = load_capacity_mapping(REPO / "config" / "default_capacity_mapping.csv")

        be_fr = mapping.map_forecast(
            "RSA_BE_FR_VOL&FTE Sep 2026.xlsx", "BE RSA FO FR",
        )
        self.assertEqual(
            (be_fr.management_lob, be_fr.planning_group, be_fr.staff_type, be_fr.status),
            ("RSA BE", "RSA BE FR", "BE RSA FO FR", "MAPPED"),
        )
        be_vl = mapping.map_schedule("RSA VL", "BE RSA Front-office VL")
        self.assertEqual(
            (be_vl.management_lob, be_vl.planning_group, be_vl.staff_type, be_vl.status),
            ("RSA BE", "RSA BE VL", "BE RSA FO VL", "MAPPED"),
        )

        unknown = mapping.map_forecast(
            "RSA_BE_FR_VOL&FTE Sep 2026.xlsx",
            "APBN_BRU_MOBILITY_INSURAN_Front_FR",
        )
        self.assertEqual(unknown.status, "UNMAPPED_STAFF_TYPE")
        self.assertEqual(unknown.planning_group, "RSA BE FR")

    def test_default_merge_preserves_user_rows_and_adds_new_defaults(self):
        with tempfile.TemporaryDirectory() as folder:
            home = Path(folder)
            (home / "config").mkdir()
            shutil.copy2(
                REPO / "config" / "default_capacity_mapping.csv",
                home / "config" / "default_capacity_mapping.csv",
            )
            target = home / "config" / "capacity_mapping.csv"
            target.write_text(
                "management_lob,planning_group,workforce_lob,forecast_file_prefix,forecast_staff_type,schedule_assignment,staff_type\n"
                "CUSTOM,Custom Group,Custom LOB,CUSTOM,Custom Staff,Custom Shift,Custom Staff\n",
                encoding="utf-8",
            )

            ensure_capacity_mapping(home, target)

            text = target.read_text(encoding="utf-8-sig")
            self.assertIn("CUSTOM,Custom Group,Custom LOB", text)
            self.assertIn("RSA BE,RSA BE FR,RSA FR", text)
            backups = list(target.parent.glob("capacity_mapping_pre_default_merge_*.csv"))
            self.assertEqual(len(backups), 1)


if __name__ == "__main__":
    unittest.main()
