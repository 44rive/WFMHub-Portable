from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from wfmhub.metrics import MetricCatalogError, load_metric_catalog
from wfmhub.rules import (
    RulebookError,
    ensure_rulebook,
    evaluate_formula,
    load_rulebook,
    validate_rulebook,
)


REPO = Path(__file__).resolve().parents[1]


class RulebookTests(unittest.TestCase):
    def test_default_rulebook_validates_and_classifies_real_verint_labels(self):
        rules = load_rulebook(REPO, REPO / "config" / "default_rules.toml")
        self.assertTrue(validate_rulebook(rules))
        self.assertEqual(rules.classify_activity(".AP BEN | Short Sickness").category, "SICKNESS_SHORT")
        self.assertEqual(rules.classify_activity(".AP BEN | Leave - Unpaid").category, "UNPAID_LEAVE")
        self.assertEqual(rules.classify_activity(".AP BEN | NR - Lunch Break").category, "LUNCH")
        self.assertEqual(rules.classify_activity(".AP BEN | BE RSA Front-office FR").category, "PRODUCTION")
        self.assertEqual(rules.classify_status("Disponible").attendance_category, "Productive")
        self.assertEqual(rules.classify_status("Disponible").aux_classification, "Available")
        self.assertEqual(rules.classify_status("Back-Office").qualification_1, "BO productive")
        self.assertEqual(rules.classify_status("Meal").attendance_category, "Lunch")
        self.assertEqual(rules.classify_status("not configured"), None)
        self.assertEqual(len(rules.status_rules), 32)

    def test_formula_engine_calculates_service_availability_and_rejects_code(self):
        self.assertEqual(evaluate_formula("answered / nullif(offered, 0)", {"answered": 9, "offered": 10}), 0.9)
        self.assertIsNone(evaluate_formula("answered / nullif(offered, 0)", {"answered": 9, "offered": 0}))
        self.assertEqual(evaluate_formula("ifelse(offered > 0, answered / offered, 0)", {"answered": 9, "offered": 10}), 0.9)
        with self.assertRaises(RulebookError):
            evaluate_formula("__import__('os').system('whoami')", {})

    def test_rulebook_upgrade_changes_shipped_sla_target_without_losing_local_rules(self):
        with tempfile.TemporaryDirectory() as folder:
            home = Path(folder)
            config = home / "config"
            config.mkdir()
            default = (REPO / "config" / "default_rules.toml").read_text(
                encoding="utf-8",
            )
            (config / "default_rules.toml").write_text(default, encoding="utf-8")
            current = default.replace('version = "2026.09.3"', 'version = "2026.08.3"', 1)
            current = current.split("# Agent Status and AUX reference supplied by Operations.", 1)[0]
            current = current.replace("target_seconds = 30", "target_seconds = 20", 1)
            current = current.replace(
                "Canonical WFM rules learned", "Locally reviewed WFM rules learned", 1,
            )
            (config / "wfm_rules.toml").write_text(current, encoding="utf-8")
            target = ensure_rulebook(home)
            migrated = target.read_text(encoding="utf-8")
            self.assertIn('version = "2026.09.3"', migrated)
            self.assertIn("target_seconds = 30", migrated)
            self.assertIn("Locally reviewed WFM rules learned", migrated)
            self.assertIn('status = "Disponible"', migrated)
            self.assertEqual(len(list(config.glob("wfm_rules_pre_storm_service_*.toml"))), 1)

    def test_invalid_metric_formula_is_rejected_before_refresh(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "bad.toml"
            text = (REPO / "config" / "default_metrics.toml").read_text(encoding="utf-8")
            path.write_text(
                text.replace(
                    'numerator = "answered_within_target"',
                    'numerator = "answered_within_target // offered"',
                    1,
                ),
                encoding="utf-8",
            )
            with self.assertRaisesRegex(MetricCatalogError, "forbidden formula element"):
                load_metric_catalog(REPO, path)


if __name__ == "__main__":
    unittest.main()
