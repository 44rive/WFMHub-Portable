from __future__ import annotations

import tempfile
import unittest
import zipfile
from contextlib import nullcontext
from datetime import date
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import MagicMock, patch

from openpyxl import load_workbook

from wfmhub.cli import _build_latest_pcs_report, _build_pcs_from_database, refresh
from wfmhub.models import ModelSummary
from wfmhub.pcs_tracker import (
    PCS_TRACKER_FILENAME,
    ensure_pcs_tracker,
    latest_pcs_report,
)


class PCSWorkflowTests(unittest.TestCase):
    def test_live_tracker_is_created_once_and_never_replaced(self):
        with tempfile.TemporaryDirectory() as folder:
            reports = Path(folder) / "Reports"
            config = SimpleNamespace(reports=reports)
            path = ensure_pcs_tracker(config)
            self.assertEqual(path.name, PCS_TRACKER_FILENAME)
            with zipfile.ZipFile(path) as archive:
                self.assertIsNone(archive.testzip())
                self.assertIn("xl/metadata.xml", archive.namelist())
                worksheet_xml = b"".join(
                    archive.read(name) for name in archive.namelist()
                    if name.startswith("xl/worksheets/sheet") and name.endswith(".xml")
                )
                self.assertIn(b"_xlfn._xlws.FILTER", worksheet_xml)
                self.assertNotIn(b"#REF!", worksheet_xml)

            workbook = load_workbook(path)
            sheet = workbook["COACHING"]
            sheet["L10"] = "call-key-1"
            sheet["N10"] = "Completed"
            workbook.save(path)
            workbook.close()
            saved = path.read_bytes()

            self.assertEqual(ensure_pcs_tracker(config), path)
            self.assertEqual(path.read_bytes(), saved)

    def test_latest_report_is_the_fixed_permanent_tracker(self):
        with tempfile.TemporaryDirectory() as folder:
            reports = Path(folder)
            config = SimpleNamespace(reports=reports)
            tracker = reports / PCS_TRACKER_FILENAME
            tracker.write_bytes(b"tracker")
            self.assertEqual(latest_pcs_report(config), tracker)

    def test_latest_build_loads_only_pcs_sources_and_prepares_paste_data(self):
        with tempfile.TemporaryDirectory() as folder:
            home = Path(folder)
            report = home / "Reports" / PCS_TRACKER_FILENAME
            report.parent.mkdir()
            report.write_bytes(b"tracker")
            config = SimpleNamespace(reports=report.parent)
            with (
                patch("wfmhub.cli.refresh", return_value=0) as run,
                patch("wfmhub.cli.load_config", return_value=config),
                patch("wfmhub.cli.latest_pcs_report", return_value=report),
                patch("wfmhub.cli.latest_pcs_paste", return_value=report.parent / "PCS Paste Data.xlsx"),
            ):
                _build_latest_pcs_report(home)
            run.assert_called_once_with(home, None, None, ("pcs",), "pcs", False)

    def test_fast_rebuild_uses_current_database_without_source_ingestion(self):
        with tempfile.TemporaryDirectory() as folder:
            home = Path(folder)
            report = home / "Reports" / PCS_TRACKER_FILENAME
            config = SimpleNamespace(reports=report.parent)
            conn = MagicMock()
            with (
                patch("wfmhub.cli.load_config", return_value=config),
                patch("wfmhub.cli.write_session", return_value=nullcontext(conn)),
                patch(
                    "wfmhub.cli._pcs_mart_period",
                    return_value=(date(2026, 8, 1), date(2026, 9, 8)),
                ),
                patch("wfmhub.cli.build_report_pack", return_value=report) as build,
                patch("wfmhub.cli.latest_pcs_paste", return_value=report.parent / "PCS Paste Data.xlsx"),
            ):
                self.assertEqual(_build_pcs_from_database(home), report)
            build.assert_called_once_with(
                "pcs", conn, config, date(2026, 8, 1), date(2026, 9, 8),
            )

    def test_pcs_refresh_uses_targeted_model_without_feeds_or_excel(self):
        home = Path("/test/wfmhub")
        config = SimpleNamespace(
            business_rules=Path("rules.toml"),
            queue_mapping=Path("queues.csv"),
        )
        conn = MagicMock()
        ingest = SimpleNamespace(
            loaded=0, skipped=2, failed=0, scoped_out=0, errors=[],
        )
        model = ModelSummary(
            start=date(2026, 8, 1), end=date(2026, 9, 8), pcs_rows=12,
        )
        rulebook = SimpleNamespace(version="rules", sha256="a" * 64)
        mapping = SimpleNamespace(sha256="b" * 64)
        report = Path(f"/test/wfmhub/Reports/{PCS_TRACKER_FILENAME}")
        with (
            patch("wfmhub.cli.load_config", return_value=config),
            patch("wfmhub.cli._logging"),
            patch("wfmhub.cli.ProgressBar", return_value=MagicMock()),
            patch("wfmhub.cli.write_session", return_value=nullcontext(conn)),
            patch("wfmhub.cli.ingest_all", return_value=ingest),
            patch("wfmhub.cli.refresh_pcs_models", return_value=model) as pcs_models,
            patch("wfmhub.cli.refresh_models") as all_models,
            patch("wfmhub.cli.publish_shared_feeds") as all_feeds,
            patch("wfmhub.cli.build_report_pack", return_value=report) as build,
            patch("wfmhub.cli.load_rulebook", return_value=rulebook),
            patch("wfmhub.cli.load_queue_mapping", return_value=mapping),
        ):
            result = refresh(home, None, None, ("pcs",), "pcs", False)
        self.assertEqual(result, 0)
        pcs_models.assert_called_once()
        all_models.assert_not_called()
        all_feeds.assert_not_called()
        build.assert_called_once_with(
            "pcs", conn, config, model.start, model.end, service_profile=None,
        )


if __name__ == "__main__":
    unittest.main()
