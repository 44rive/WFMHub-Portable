from __future__ import annotations

import shutil
import tempfile
import unittest
from pathlib import Path

from wfmhub.mapping import ensure_queue_mapping, load_queue_mapping


REPO = Path(__file__).resolve().parents[1]


class QueueMappingTests(unittest.TestCase):
    def test_default_mapping_maps_forecasts_queues_and_rollups(self):
        mapping = load_queue_mapping(REPO / "config" / "default_queue_mapping.csv")
        forecast = mapping.map_forecast("RSA_BE_08-2026.txt", "Combined - All Media")
        self.assertEqual((forecast.service_scope, forecast.comparison_scope), ("RSA BE", "RSA BE"))
        prefixed = mapping.map_forecast("Forecast_RSA_NL_August.txt", "Combined - All Media")
        self.assertEqual((prefixed.service_scope, prefixed.comparison_scope), ("RSA NL", "RSA NL"))
        september = {
            "RSA_NL_09-2026.txt": "RSA NL",
            "RSA_BE_09-2026.txt": "RSA BE",
            "FORD_NL_09-2026.txt": "Ford NL",
            "FORD_FR_09-2026.txt": "Ford FR",
        }
        for filename, scope in september.items():
            result = mapping.map_forecast(filename, "Combined - All Media")
            self.assertEqual(
                (result.service_scope, result.comparison_scope),
                (scope, scope),
            )
        belgium = mapping.map_actual("APBE", "APBN_BRU_RSA_INTERNAT_All_FR", None, "RSA")
        self.assertEqual((belgium.service_scope, belgium.comparison_scope), ("RSA BE FR", "RSA BE"))
        self.assertEqual(
            mapping.comparison_scopes_for(("RSA BE FR", "RSA BE VL")),
            ("RSA BE",),
        )
        ford_nl = mapping.map_actual("APBE", "APBN_AMS_MOBILITY_Ford_Assistance_NL", None, "FORD")
        self.assertEqual((ford_nl.service_scope, ford_nl.comparison_scope), ("Ford NL", "Ford NL"))
        reference_additions = {
            "APBN_AMS_MOBILITY_VARIOUS_VariousAssist_NL": "RSA NL",
            "APBN_AMS_MOBILITY_NIGHT_NightShift_NL": "RSA NL",
            "APBN_AMS_RSA_OEM_All_NL": "RSA NL",
            "APBN_AMS_RSA_PROVIDER_All_NL": "RSA NL",
            "APBN_BRU_MOBILITY_BIKE_Bike_FR": "RSA BE FR",
            "APBN_BRU_MOBILITY_BIKE_Bike_VL": "RSA BE VL",
            "APBN_BRU_MOBILITY_NIGHT_NightShift_FR": "RSA BE FR",
            "APBN_BRU_RSA_INSURAN_AzB_VL": "RSA BE VL",
            "APBN_BRU_RSA_OEM-CONS_All_FR": "RSA BE FR",
            "APBN_BRU_RSA_OEM_All_EN": "RSA BE VL",
            "APBN_BRU_MOBILITY_Ford_Assistance_DE": "Ford NL",
            "APBN_LUX_MOBILITY_Ford_Assistance_DE": "Ford NL",
        }
        for queue, scope in reference_additions.items():
            self.assertEqual(
                mapping.map_actual("STORM", queue, None, None).service_scope,
                scope,
            )

    def test_new_defaults_merge_without_replacing_local_queue_override(self):
        with tempfile.TemporaryDirectory() as folder:
            home = Path(folder)
            (home / "config").mkdir()
            shipped = home / "config" / "default_queue_mapping.csv"
            shutil.copy2(REPO / "config" / "default_queue_mapping.csv", shipped)
            target = home / "config" / "queue_mapping.csv"
            target.write_text(
                "mapping_type,source_system,source_value,service_scope,designation\n"
                "queue,STORM,APBN_AMS_MOBILITY_VARIOUS_VariousAssist_NL,CUSTOM,Custom\n",
                encoding="utf-8",
            )
            ensure_queue_mapping(home, target)
            mapping = load_queue_mapping(target)
            self.assertEqual(
                mapping.map_actual(
                    "STORM", "APBN_AMS_MOBILITY_VARIOUS_VariousAssist_NL", None, None,
                ).service_scope,
                "CUSTOM",
            )
            self.assertEqual(
                mapping.map_actual(
                    "STORM", "APBN_BRU_MOBILITY_Ford_Assistance_DE", None, None,
                ).service_scope,
                "Ford NL",
            )
            self.assertTrue(list((home / "config").glob("queue_mapping_pre_default_merge_*.csv")))

    def test_unlisted_apde_partner_falls_back_to_raw_lob(self):
        mapping = load_queue_mapping(REPO / "config" / "default_queue_mapping.csv")
        result = mapping.map_actual("APDE", "Some Partner", "Some Partner", "RSA_Automotive")
        self.assertEqual((result.service_scope, result.status), ("RSA_Automotive", "FALLBACK_LOB"))


if __name__ == "__main__":
    unittest.main()
