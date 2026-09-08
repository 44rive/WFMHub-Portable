from __future__ import annotations

import unittest
import tempfile
import shutil
from datetime import datetime
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

from wfmhub.cli import SOURCE_GROUPS
from wfmhub.config import ConfigError, load_config
from wfmhub.design import COLORS, REPORT_DESIGN_ID, REPORT_DESIGN_VERSION
from wfmhub.reports import COLORS as REPORT_COLORS
from wfmhub.shared_reports import COLORS as SHARED_REPORT_COLORS
from wfmhub.report_packs import (
    IMPLEMENTED_REPORT_PACK_KEYS,
    REPORT_PACKS,
    ReportPublishError,
    archive_superseded_reports,
    build_report_pack,
    publish_report,
    report_current_path,
)


class ReportPackTests(unittest.TestCase):
    def test_report_design_has_one_palette_and_a_data_free_reference(self):
        repo = Path(__file__).resolve().parents[1]
        self.assertIs(REPORT_COLORS, COLORS)
        self.assertIs(SHARED_REPORT_COLORS, COLORS)
        self.assertEqual((REPORT_DESIGN_ID, REPORT_DESIGN_VERSION), (
            "WFMHUB-DESIGN", "2.0.0",
        ))
        reference = repo / "docs" / "WFMHub Report Design Reference.xlsx"
        self.assertTrue(reference.is_file())
        from openpyxl import load_workbook

        workbook = load_workbook(reference, read_only=True, data_only=False)
        try:
            self.assertEqual(workbook.sheetnames, [
                "DESIGN SYSTEM", "PCS", "RTM", "ATTENDANCE",
            ])
            self.assertEqual(workbook["DESIGN SYSTEM"]["A1"].value, "WFMHUB REPORT SYSTEM")
        finally:
            workbook.close()

    def test_pcs_tracker_window_is_configured_and_validated(self):
        repo = Path(__file__).resolve().parents[1]
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            (root / "config").mkdir()
            for name in (
                "default.toml", "default_rules.toml", "default_metrics.toml",
                "default_analytics.toml", "default_reports.toml",
            ):
                shutil.copy2(repo / "config" / name, root / "config" / name)
            default = root / "config" / "default.toml"
            config = load_config(root, default)
            self.assertEqual(config.pcs_tracker.history_months, 13)
            self.assertEqual(config.pcs_tracker.trend_days, 90)
            invalid = root / "invalid.toml"
            invalid.write_text(
                default.read_text(encoding="utf-8").replace(
                    "history_months = 13", "history_months = 1",
                ),
                encoding="utf-8",
            )
            with self.assertRaises(ConfigError):
                load_config(root, invalid)

    def test_service_refresh_group_includes_flash_actual_and_forecast_sources(self):
        self.assertEqual(
            SOURCE_GROUPS["service"],
            {"fte", "schedule", "lilo", "agent_status", "calls", "forecast"},
        )
        self.assertFalse(
            {"apbe", "apfr", "apde"} & set(SOURCE_GROUPS["intraday"] or ())
        )

    def test_independent_report_packs_are_registered(self):
        self.assertEqual(IMPLEMENTED_REPORT_PACK_KEYS, (
            "pcs", "bonus", "service", "realisations", "staffing",
            "corrections", "absence",
        ))
        self.assertTrue(all(REPORT_PACKS[key].implemented for key in IMPLEMENTED_REPORT_PACK_KEYS))
        self.assertFalse(REPORT_PACKS["intraday"].implemented)
        self.assertEqual(REPORT_PACKS["pcs"].default_folder, "pcs")
        self.assertEqual(REPORT_PACKS["pcs"].current_filename, "PCS Operational Tracker.xlsx")
        self.assertEqual(REPORT_PACKS["service"].default_folder, "service")
        self.assertEqual(REPORT_PACKS["realisations"].default_folder, "realisations")
        self.assertEqual(REPORT_PACKS["attendance"].default_folder, "attendance")
        self.assertEqual(REPORT_PACKS["absence"].default_folder, "absence")
        self.assertEqual(REPORT_PACKS["corrections"].default_folder, "corrections")
        # Old command/API keys remain callable for backwards compatibility, but
        # are deliberately absent from the interactive product menu.
        self.assertTrue(REPORT_PACKS["operations"].implemented)
        self.assertTrue(REPORT_PACKS["quality_pcs"].implemented)
        self.assertFalse(REPORT_PACKS["scorecard"].implemented)

    def test_current_reports_are_flat_and_previous_copy_is_archived(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            config = SimpleNamespace(
                reports=root / "Reports",
                system=root / "_system",
            )
            current = report_current_path(config, "service")
            current.parent.mkdir(parents=True)
            current.write_bytes(b"old")
            partial = current.with_name("RTM Daily Control.partial.xlsx")
            partial.write_bytes(b"new")
            publish_report(
                config, "service", partial, current,
                datetime(2026, 9, 4, 10, 30),
            )
            self.assertEqual(current.read_bytes(), b"new")
            archives = list((root / "Reports" / "Archive" / "2026-09-04").glob("*.xlsx"))
            self.assertEqual(len(archives), 1)
            self.assertEqual(archives[0].read_bytes(), b"old")
            self.assertEqual(current.name, "RTM Daily Control.xlsx")
            self.assertEqual(
                report_current_path(config, "attendance").name,
                "Legacy Attendance Callout.xlsx",
            )
            self.assertEqual(report_current_path(config, "corrections").name, "Attendance Review.xlsx")

    def test_replaced_operational_files_are_archived_not_deleted(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            config = SimpleNamespace(
                reports=root / "Reports",
                system=root / "_system",
            )
            config.reports.mkdir(parents=True)
            for filename in (
                "Attendance Callout.xlsx", "Service Flashes.xlsx", "OEM Flash.xlsx",
            ):
                (config.reports / filename).write_bytes(filename.encode())
            archived = archive_superseded_reports(
                config,
                ("Attendance Callout.xlsx", "Service Flashes.xlsx", "OEM Flash.xlsx"),
                datetime(2026, 9, 6, 16, 0),
            )
            self.assertEqual(len(archived), 3)
            self.assertFalse((config.reports / "Attendance Callout.xlsx").exists())
            self.assertTrue(all(path.is_file() for path in archived))

    def test_locked_workbook_has_an_actionable_publish_error(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            partial = root / "PCS Operational Tracker.partial.xlsx"
            target = root / "PCS Operational Tracker.xlsx"
            partial.write_bytes(b"complete replacement")
            with patch.object(
                Path, "replace", side_effect=PermissionError("access denied"),
            ):
                with self.assertRaisesRegex(
                    ReportPublishError, "open in Excel.*locked by OneDrive",
                ):
                    publish_report(
                        SimpleNamespace(), "custom", partial, target,
                        datetime(2026, 9, 8, 0, 5),
                    )


if __name__ == "__main__":
    unittest.main()
