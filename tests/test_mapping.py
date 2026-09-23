from __future__ import annotations

import shutil
import tempfile
import unittest
from pathlib import Path

from wfmhub.mapping import ensure_queue_mapping, load_queue_mapping


REPO = Path(__file__).resolve().parents[1]


class QueueMappingTests(unittest.TestCase):
    def test_default_mapping_resolves_current_service_scopes(self):
        mapping = load_queue_mapping(REPO / "config" / "default_queue_mapping.csv")
        self.assertGreaterEqual(len(mapping.queue_rows), 80)
        for filename, scope in {
            "RSA_NL_09-2026.txt": "RSA NL",
            "RSA_BE_09-2026.txt": "RSA BE",
            "FORD_NL_09-2026.txt": "Ford NL",
            "FORD_FR_09-2026.txt": "Ford FR",
        }.items():
            result = mapping.map_forecast(filename, "Combined - All Media")
            self.assertEqual((result.service_scope, result.comparison_scope), (scope, scope))

        belgium = mapping.map_actual(
            "STORM", "APBN_BRU_RSA_INTERNAT_All_FR", None, None,
        )
        self.assertEqual(
            (belgium.service_scope, belgium.comparison_scope),
            ("RSA BE FR", "RSA BE"),
        )
        self.assertEqual(
            mapping.comparison_scopes_for(("RSA BE FR", "RSA BE VL")),
            ("RSA BE",),
        )

        for queue in (
            "APFR_PAR_RSA_CSTRUCTR_FORD_ASSISTANCE_FR",
            "APFR_PAR_RSA_CSTRUCTR_TOYOTA-LEXUS_FR",
            "APFR_PAR_RSA_CHERY_ASSISTANCE_FR",
        ):
            self.assertEqual(
                mapping.map_actual("STORM", queue, None, None).service_scope,
                "Ford FR",
            )

    def test_unknown_queue_is_explicitly_unmapped(self):
        mapping = load_queue_mapping(REPO / "config" / "default_queue_mapping.csv")
        result = mapping.map_actual("STORM", "UNKNOWN_QUEUE", None, None)
        self.assertEqual((result.service_scope, result.status), ("UNMAPPED", "UNMAPPED"))

    def test_copy_on_first_run_never_rewrites_user_mapping(self):
        with tempfile.TemporaryDirectory() as folder:
            home = Path(folder)
            (home / "config").mkdir()
            shutil.copy2(
                REPO / "config" / "default_queue_mapping.csv",
                home / "config" / "default_queue_mapping.csv",
            )
            target = home / "config" / "queue_mapping.csv"
            content = (
                "mapping_type,source_system,source_value,service_scope,designation\n"
                "queue,STORM,CUSTOM_QUEUE,CUSTOM,Custom\n"
            )
            target.write_text(content, encoding="utf-8")
            ensure_queue_mapping(home, target)
            self.assertEqual(target.read_text(encoding="utf-8"), content)
            self.assertEqual(
                load_queue_mapping(target).map_actual(
                    "STORM", "CUSTOM_QUEUE", None, None,
                ).service_scope,
                "CUSTOM",
            )


if __name__ == "__main__":
    unittest.main()
