from __future__ import annotations

import unittest
from pathlib import Path

from wfmhub.mapping import load_queue_mapping


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

    def test_unlisted_apde_partner_falls_back_to_raw_lob(self):
        mapping = load_queue_mapping(REPO / "config" / "default_queue_mapping.csv")
        result = mapping.map_actual("APDE", "Some Partner", "Some Partner", "RSA_Automotive")
        self.assertEqual((result.service_scope, result.status), ("RSA_Automotive", "FALLBACK_LOB"))


if __name__ == "__main__":
    unittest.main()
