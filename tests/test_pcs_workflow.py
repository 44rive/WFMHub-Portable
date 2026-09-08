from __future__ import annotations

import tempfile
import unittest
from contextlib import nullcontext
from datetime import date
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import MagicMock, patch

from wfmhub.cli import _update_pcs_now, refresh
from wfmhub.models import ModelSummary
from wfmhub.pcs_excel import PCSExcelError, PCSTrackerState, PCS_TEMPLATE_VERSION


class PCSWorkflowTests(unittest.TestCase):
    def _state(
        self,
        workbook: Path,
        *,
        version: str | None = PCS_TEMPLATE_VERSION,
        installed: bool = True,
        problem: str | None = None,
    ) -> PCSTrackerState:
        return PCSTrackerState(
            path=workbook,
            exists=True,
            template_version=version,
            setup_state="YES" if installed else "NO",
            connection_mode="LOCAL",
            query_parts=5 if installed else 0,
            has_connections=installed,
            problem=problem,
        )

    def test_normal_update_refreshes_data_and_feeds_without_rebuilding_tracker(self):
        with tempfile.TemporaryDirectory() as folder:
            home = Path(folder)
            workbook = home / "Reports" / "PCS Operational Tracker.xlsx"
            workbook.parent.mkdir()
            workbook.write_bytes(b"permanent tracker")
            config = SimpleNamespace(feed=home / "Feed")
            state = self._state(workbook)
            with (
                patch("wfmhub.cli.load_config", return_value=config),
                patch("wfmhub.cli.report_current_path", return_value=workbook),
                patch("wfmhub.cli.inspect_pcs_tracker", return_value=state),
                patch("wfmhub.cli.refresh", return_value=0) as refresh,
                patch("wfmhub.cli._rebuild_pcs_tracker_from_marts") as rebuild,
                patch(
                    "wfmhub.cli.run_pcs_excel_action",
                    return_value="PCS WORKBOOK READY",
                ) as excel,
            ):
                _update_pcs_now(home)
            refresh.assert_called_once_with(home, None, None, (), "pcs", False)
            rebuild.assert_not_called()
            excel.assert_called_once_with(
                config, workbook, "Refresh", "LOCAL", open_after=True,
            )

    def test_pcs_source_refresh_uses_targeted_model_and_feed_only(self):
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
        feed = SimpleNamespace(rows=12)
        rulebook = SimpleNamespace(version="rules", sha256="a" * 64)
        mapping = SimpleNamespace(sha256="b" * 64)
        with (
            patch("wfmhub.cli.load_config", return_value=config),
            patch("wfmhub.cli._logging"),
            patch("wfmhub.cli.ProgressBar", return_value=MagicMock()),
            patch("wfmhub.cli.write_session", return_value=nullcontext(conn)),
            patch("wfmhub.cli.ingest_all", return_value=ingest),
            patch("wfmhub.cli.refresh_pcs_models", return_value=model) as pcs_models,
            patch("wfmhub.cli.refresh_models") as all_models,
            patch("wfmhub.cli.publish_pcs_feeds", return_value=feed) as pcs_feed,
            patch("wfmhub.cli.publish_shared_feeds") as all_feeds,
            patch("wfmhub.cli.load_rulebook", return_value=rulebook),
            patch("wfmhub.cli.load_queue_mapping", return_value=mapping),
        ):
            result = refresh(home, None, None, (), "pcs", False)
        self.assertEqual(result, 0)
        pcs_models.assert_called_once()
        all_models.assert_not_called()
        pcs_feed.assert_called_once_with(
            conn, config, date(2026, 8, 1), date(2026, 9, 8),
        )
        all_feeds.assert_not_called()

    def test_unreadable_or_locked_tracker_stops_before_long_source_refresh(self):
        with tempfile.TemporaryDirectory() as folder:
            home = Path(folder)
            workbook = home / "PCS Operational Tracker.xlsx"
            config = SimpleNamespace(feed=home / "Feed")
            state = self._state(workbook, problem="[WinError 5] Access is denied")
            with (
                patch("wfmhub.cli.load_config", return_value=config),
                patch("wfmhub.cli.report_current_path", return_value=workbook),
                patch("wfmhub.cli.inspect_pcs_tracker", return_value=state),
                patch("wfmhub.cli.refresh") as refresh,
            ):
                with self.assertRaisesRegex(PCSExcelError, "open in Excel.*Close it"):
                    _update_pcs_now(home)
            refresh.assert_not_called()

    def test_old_readable_tracker_is_upgraded_before_refresh(self):
        with tempfile.TemporaryDirectory() as folder:
            home = Path(folder)
            workbook = home / "PCS Operational Tracker.xlsx"
            workbook.write_bytes(b"old tracker")
            config = SimpleNamespace(feed=home / "Feed")
            old = self._state(workbook, version="old", installed=False)
            current = self._state(workbook, installed=False)
            with (
                patch("wfmhub.cli.load_config", return_value=config),
                patch("wfmhub.cli.report_current_path", return_value=workbook),
                patch("wfmhub.cli.inspect_pcs_tracker", side_effect=(old, current)),
                patch("wfmhub.cli.refresh", return_value=0) as refresh,
                patch(
                    "wfmhub.cli._rebuild_pcs_tracker_from_marts",
                    return_value=workbook,
                ) as rebuild,
                patch(
                    "wfmhub.cli.run_pcs_excel_action",
                    return_value="PCS WORKBOOK READY",
                ) as excel,
            ):
                _update_pcs_now(home)
            rebuild.assert_called_once_with(home)
            refresh.assert_called_once_with(home, None, None, (), "pcs", False)
            excel.assert_called_once_with(
                config, workbook, "Install", "LOCAL", open_after=True,
            )

    def test_excel_failure_does_not_claim_the_data_refresh_failed(self):
        with tempfile.TemporaryDirectory() as folder:
            home = Path(folder)
            workbook = home / "PCS Operational Tracker.xlsx"
            workbook.write_bytes(b"permanent tracker")
            config = SimpleNamespace(feed=home / "Feed")
            state = self._state(workbook)
            with (
                patch("wfmhub.cli.load_config", return_value=config),
                patch("wfmhub.cli.report_current_path", return_value=workbook),
                patch("wfmhub.cli.inspect_pcs_tracker", return_value=state),
                patch("wfmhub.cli.refresh", return_value=0),
                patch(
                    "wfmhub.cli.run_pcs_excel_action",
                    side_effect=PCSExcelError("read-only workbook"),
                ),
            ):
                with self.assertRaisesRegex(
                    PCSExcelError,
                    "all four clean feeds updated successfully.*Refresh Excel only",
                ):
                    _update_pcs_now(home)


if __name__ == "__main__":
    unittest.main()
