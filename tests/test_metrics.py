from __future__ import annotations

import tempfile
import unittest
from datetime import date
from pathlib import Path

from wfmhub.metrics import (
    MetricCatalogError,
    diff_metric_catalogs,
    evaluate_metric,
    load_metric_catalog,
    validate_metric_catalog,
)


REPO = Path(__file__).resolve().parents[1]


class MetricCatalogTests(unittest.TestCase):
    def test_default_catalog_evaluates_ratio_and_zero_denominator(self):
        catalog = load_metric_catalog(REPO, REPO / "config" / "default_metrics.toml")
        method = catalog.method_for("service_level", date(2026, 8, 1), {"lob": "FORD"})
        result = evaluate_metric(method, {
            "answered_within_target": 75,
            "offered": 100,
            "short_abandoned": 5,
        })
        self.assertAlmostEqual(result.value, 75 / 95)
        self.assertEqual(result.state, "BELOW_TARGET")
        zero = evaluate_metric(method, {
            "answered_within_target": 0,
            "offered": 5,
            "short_abandoned": 5,
        })
        self.assertIsNone(zero.value)
        self.assertEqual(zero.state, "NO_DATA")

    def test_effective_date_and_scope_priority_select_one_method(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "metrics.toml"
            path.write_text(
                """[catalog]
version = "test"

[[metrics]]
id = "service_level"
method = "default"
label = "Default"
domain = "service"
source_model = "service_interval"
grain = "interval"
unit = "percent"
aggregation = "ratio_of_sums"
numerator = "answered_within_target"
denominator = "offered"
effective_from = 2026-01-01
effective_to = 2026-06-30
priority = 0

[[metrics]]
id = "service_level"
method = "default_v2"
label = "Default v2"
domain = "service"
source_model = "service_interval"
grain = "interval"
unit = "percent"
aggregation = "ratio_of_sums"
numerator = "answered_within_target"
denominator = "offered - short_abandoned"
effective_from = 2026-07-01
priority = 0

[[metrics]]
id = "service_level"
method = "rsa_v2"
label = "RSA v2"
domain = "service"
source_model = "service_interval"
grain = "interval"
unit = "percent"
aggregation = "ratio_of_sums"
numerator = "answered_within_target"
denominator = "offered"
effective_from = 2026-07-01
priority = 10
scope = { lob_contains = ["RSA"] }
""",
                encoding="utf-8",
            )
            catalog = load_metric_catalog(REPO, path)
            validate_metric_catalog(catalog, {
                "service_interval": {"answered_within_target", "offered", "short_abandoned"},
            })
            self.assertEqual(
                catalog.method_for("service_level", date(2026, 6, 30), {"lob": "RSA"}).method_id,
                "default",
            )
            self.assertEqual(
                catalog.method_for("service_level", date(2026, 8, 1), {"lob": "FORD"}).method_id,
                "default_v2",
            )
            self.assertEqual(
                catalog.method_for("service_level", date(2026, 8, 1), {"lob": "BE RSA"}).method_id,
                "rsa_v2",
            )
            ambiguous_path = Path(folder) / "ambiguous.toml"
            ambiguous_path.write_text(
                path.read_text(encoding="utf-8").replace("priority = 10", "priority = 0"),
                encoding="utf-8",
            )
            ambiguous = load_metric_catalog(REPO, ambiguous_path)
            with self.assertRaisesRegex(MetricCatalogError, "Ambiguous equal-priority"):
                validate_metric_catalog(ambiguous, {
                    "service_interval": {"answered_within_target", "offered", "short_abandoned"},
                })

    def test_catalog_diff_reports_method_changes(self):
        current = load_metric_catalog(REPO, REPO / "config" / "default_metrics.toml")
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "changed.toml"
            text = (REPO / "config" / "default_metrics.toml").read_text(encoding="utf-8")
            path.write_text(text.replace("target = 0.80", "target = 0.85", 1), encoding="utf-8")
            changed = load_metric_catalog(REPO, path)
        self.assertTrue(any(line.startswith("CHANGED service_level.adjusted_20")
                            for line in diff_metric_catalogs(current, changed)))


if __name__ == "__main__":
    unittest.main()
