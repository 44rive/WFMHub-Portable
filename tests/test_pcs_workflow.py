from __future__ import annotations

import os
import tempfile
import unittest
from contextlib import nullcontext
from datetime import date
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import MagicMock, patch

from openpyxl import load_workbook

from wfmhub.cli import _build_latest_pcs_report, _build_pcs_from_database, refresh
from wfmhub.models import ModelSummary
from wfmhub.pcs_report import (
    PCS_COACHING_LOG_FILENAME,
    ensure_coaching_log,
    latest_pcs_report,
)


class PCSWorkflowTests(unittest.TestCase):
    def test_coaching_log_is_created_once_and_never_replaced(self):
        with tempfile.TemporaryDirectory() as folder:
            reports = Path(folder) / "Reports"
            config = SimpleNamespace(reports=reports)
            path = ensure_coaching_log(config)
            self.assertEqual(path.name, PCS_COACHING_LOG_FILENAME)

            workbook = load_workbook(path)
            sheet = workbook["COACHING"]
            sheet["A5"] = "call-key-1"
            sheet["N5"] = "Completed"
            workbook.save(path)
            workbook.close()
            saved = path.read_bytes()

            self.assertEqual(ensure_coaching_log(config), path)
            self.assertEqual(path.read_bytes(), saved)

    def test_latest_report_uses_timestamped_snapshot_files(self):
        with tempfile.TemporaryDirectory() as folder:
            reports = Path(folder)
            config = SimpleNamespace(reports=reports)
            older = reports / "PCS Operational Report - 2026-09-09 090000.xlsx"
            newer = reports / "PCS Operational Report - 2026-09-09 120000.xlsx"
            interrupted = reports / "PCS Operational Report - 2026-09-09 130000.partial.xlsx"
            older.write_bytes(b"old")
            newer.write_bytes(b"new")
            interrupted.write_bytes(b"incomplete")
            os.utime(older, (1, 1))
            os.utime(newer, (2, 2))
            os.utime(interrupted, (3, 3))
            self.assertEqual(latest_pcs_report(config), newer)

    def test_latest_build_loads_only_pcs_sources_and_creates_snapshot(self):
        with tempfile.TemporaryDirectory() as folder:
            home = Path(folder)
            report = home / "Reports" / "PCS Operational Report - test.xlsx"
            report.parent.mkdir()
            report.write_bytes(b"snapshot")
            config = SimpleNamespace(reports=report.parent)
            with (
                patch("wfmhub.cli.refresh", return_value=0) as run,
                patch("wfmhub.cli.load_config", return_value=config),
                patch("wfmhub.cli.latest_pcs_report", return_value=report),
                patch("wfmhub.cli.coaching_log_path", return_value=report.parent / PCS_COACHING_LOG_FILENAME),
            ):
                _build_latest_pcs_report(home)
            run.assert_called_once_with(home, None, None, ("pcs",), "pcs", False)

    def test_fast_rebuild_uses_current_database_without_source_ingestion(self):
        with tempfile.TemporaryDirectory() as folder:
            home = Path(folder)
            report = home / "Reports" / "PCS Operational Report - test.xlsx"
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
                patch("wfmhub.cli.coaching_log_path", return_value=report.parent / PCS_COACHING_LOG_FILENAME),
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
        report = Path("/test/wfmhub/Reports/PCS Operational Report - test.xlsx")
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
