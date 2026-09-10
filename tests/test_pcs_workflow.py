from __future__ import annotations

import tempfile
import unittest
import zipfile
from contextlib import nullcontext
from datetime import date, datetime
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import MagicMock, patch

from openpyxl import load_workbook

from wfmhub.cli import (
    _build_latest_pcs_report,
    _build_pcs_from_database,
    _install_pcs_power_query,
    refresh,
)
from wfmhub.models import ModelSummary
from wfmhub.pcs_tracker import (
    PCS_TRACKER_FILENAME,
    PCS_TRACKER_VERSION,
    _as_date,
    _tracker_contract_version,
    ensure_pcs_tracker,
    latest_pcs_report,
)
from wfmhub.shared_feeds import (
    PCS_AGENT_SCORECARD_HEADERS,
    PCS_COACHING_HEADERS,
    PCS_DAILY_SCORECARD_HEADERS,
    PCS_FILTER_HEADERS,
    PCS_LOB_SCORECARD_HEADERS,
    PCS_RESULTS_HEADERS,
    _pcs_reporting_windows,
)


def _rows() -> dict[str, list[tuple[object, ...]]]:
    refreshed = datetime(2026, 9, 8, 18, 0)
    selection = "Current MTD|All|All|All"
    filters = [
        ("PERIOD", 1, "Current MTD"),
        ("LOB", 1, "All"), ("LOB", 2, "RSA NL"),
        ("TL|All", 1, "All"), ("TL|All", 2, "TL 1"),
        ("AGENT|All|All", 1, "All"),
        ("AGENT|All|All", 2, "Agent One [001]"),
        ("AGENT|All|TL 1", 1, "All"),
        ("AGENT|All|TL 1", 2, "Agent One [001]"),
        ("TL|RSA NL", 1, "All"), ("TL|RSA NL", 2, "TL 1"),
        ("AGENT|RSA NL|All", 1, "All"),
        ("AGENT|RSA NL|All", 2, "Agent One [001]"),
        ("AGENT|RSA NL|TL 1", 1, "All"),
        ("AGENT|RSA NL|TL 1", 2, "Agent One [001]"),
    ]
    lob = (
        f"KPI|{selection}", "KPI", 0, "All", 20, 4.3, 4.1, .2,
        .15, 3, date(2026, 9, 8), refreshed,
    )
    lob_detail = (
        f"LOB|{selection}|1", "LOB", 1, "RSA NL", 20, 4.3, 4.1, .2,
        .15, 3, date(2026, 9, 8), refreshed,
    )
    agent = (
        f"AGENT|{selection}|1", 1, "Agent One", "001", "RSA NL", "TL 1",
        4.3, 4.0, .3, .15, 10, 2, date(2026, 9, 8), refreshed,
    )
    daily = (
        f"DAILY|{selection}|1", 1, date(2026, 9, 8), 4.3, .15, 10, 2,
        date(2026, 9, 8), refreshed,
    )
    result = (
        "Current MTD", date(2026, 9, 1), date(2026, 9, 8), "AGENT",
        "RSA NL", "TL 1", "Agent One [001]", "001", "Agent One", "NL",
        4.3, .15, 10, 60, 9, 2, 8, 120, "OK", date(2026, 9, 8), refreshed,
    )
    coaching = (
        "COACH|Current MTD|All|1", 1, date(2026, 9, 8), "RSA NL", "TL 1",
        "Agent One", "001", 2, "High", "call-1", "Review the explanation",
        "key-1", date(2026, 9, 8), refreshed,
    )
    return {
        "filter_rows": filters,
        "lob_rows": [lob, lob_detail],
        "agent_rows": [agent],
        "daily_rows": [daily],
        "result_rows": [result],
        "coaching_rows": [coaching],
    }


class PCSWorkflowTests(unittest.TestCase):
    def test_text_business_dates_are_parsed_for_sqlite_rows(self):
        self.assertEqual(_as_date("2026-09-08 17:00:00"), date(2026, 9, 8))

    def test_selectable_periods_have_explicit_like_for_like_comparisons(self):
        windows = {
            item.label: item
            for item in _pcs_reporting_windows(
                date(2026, 9, 10), date(2026, 9, 8),
            )
        }
        self.assertEqual(
            (windows["Latest day"].prior_start, windows["Latest day"].prior_end),
            (date(2026, 9, 8), date(2026, 9, 8)),
        )
        self.assertEqual(
            (windows["Current week"].prior_start, windows["Current week"].prior_end),
            (date(2026, 8, 31), date(2026, 9, 3)),
        )
        self.assertEqual(
            (windows["Current MTD"].prior_start, windows["Current MTD"].prior_end),
            (date(2026, 8, 1), date(2026, 8, 10)),
        )

    def test_live_tracker_is_lightweight_and_created_once(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            config = SimpleNamespace(
                reports=root / "Reports",
                feed=root / "Feed",
            )
            path = ensure_pcs_tracker(config, **_rows())
            self.assertEqual(path.name, PCS_TRACKER_FILENAME)
            self.assertEqual(_tracker_contract_version(path), PCS_TRACKER_VERSION)
            with zipfile.ZipFile(path) as archive:
                self.assertIsNone(archive.testzip())
                self.assertNotIn("xl/metadata.xml", archive.namelist())
                self.assertNotIn("xl/connections.xml", archive.namelist())
                worksheet_xml = b"".join(
                    archive.read(name) for name in archive.namelist()
                    if name.startswith("xl/worksheets/sheet") and name.endswith(".xml")
                )
                self.assertNotIn(b"_xlfn", worksheet_xml)
                self.assertNotIn(b"FILTER(", worksheet_xml)
                self.assertNotIn(b"AGGREGATE", worksheet_xml)
                self.assertNotIn(b"SUMPRODUCT", worksheet_xml)
                self.assertNotIn(b"#REF!", worksheet_xml)

            workbook = load_workbook(path, read_only=False, data_only=False)
            try:
                self.assertEqual(workbook.sheetnames, [
                    "OVERVIEW", "PERFORMANCE", "COACHING", "SETUP", "HELP",
                    "_PCS_FILTERS", "_PCS_LOB", "_PCS_AGENT", "_PCS_DAILY",
                    "_PCS_COACH", "_PCS_CALC", "_AUDIT",
                ])
                self.assertNotIn("DATA", workbook.sheetnames)
                self.assertEqual(len(workbook["OVERVIEW"]._charts), 2)
                self.assertEqual(workbook["OVERVIEW"]["A39"].value,
                                 "AGENT PERFORMANCE · CASCADING LOB → TEAM LEADER → AGENT")
                self.assertEqual(len(workbook["OVERVIEW"].data_validations.dataValidation), 4)
                self.assertTrue(workbook["OVERVIEW"]["A6"].value.startswith("=IFERROR(INDEX("))
                self.assertIn("tblPcsPerformance", workbook["PERFORMANCE"].tables)
                self.assertIn("tblCoachingQueue", workbook["COACHING"].tables)
                self.assertIn("tblCoachingActions", workbook["COACHING"].tables)
                self.assertIn("tblPcsFilters", workbook["_PCS_FILTERS"].tables)
                self.assertIn("tblPcsLob", workbook["_PCS_LOB"].tables)
                self.assertIn("tblPcsAgent", workbook["_PCS_AGENT"].tables)
                self.assertIn("tblPcsDaily", workbook["_PCS_DAILY"].tables)
                self.assertIn("tblPcsCoachingView", workbook["_PCS_COACH"].tables)
                self.assertEqual(
                    tuple(cell.value for cell in workbook["_PCS_FILTERS"][4]),
                    PCS_FILTER_HEADERS,
                )
                self.assertEqual(
                    tuple(cell.value for cell in workbook["_PCS_LOB"][4]),
                    PCS_LOB_SCORECARD_HEADERS,
                )
                self.assertEqual(
                    tuple(cell.value for cell in workbook["_PCS_AGENT"][4]),
                    PCS_AGENT_SCORECARD_HEADERS,
                )
                self.assertEqual(
                    tuple(cell.value for cell in workbook["_PCS_DAILY"][4]),
                    PCS_DAILY_SCORECARD_HEADERS,
                )
                self.assertEqual(
                    tuple(cell.value for cell in workbook["PERFORMANCE"][4]),
                    PCS_RESULTS_HEADERS,
                )
                self.assertEqual(
                    tuple(cell.value for cell in workbook["COACHING"][4][:10]),
                    PCS_COACHING_HEADERS[2:12],
                )
                self.assertEqual(
                    tuple(cell.value for cell in workbook["_PCS_COACH"][4]),
                    PCS_COACHING_HEADERS,
                )
            finally:
                workbook.close()

            cached = load_workbook(path, read_only=True, data_only=True)
            try:
                self.assertEqual(cached["OVERVIEW"]["A6"].value, 4.3)
                self.assertEqual(cached["_PCS_CALC"]["B2"].value, 4.3)
                self.assertEqual(cached["COACHING"]["A5"].value, datetime(2026, 9, 8))
            finally:
                cached.close()

            saved = path.read_bytes()
            self.assertEqual(ensure_pcs_tracker(config, **_rows()), path)
            self.assertEqual(path.read_bytes(), saved)

    def test_version_migration_preserves_keyed_actions(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            config = SimpleNamespace(
                reports=root / "Reports",
                feed=root / "Feed",
            )
            with patch("wfmhub.pcs_tracker.PCS_TRACKER_VERSION", "2026.10.9"):
                path = ensure_pcs_tracker(config, **_rows())
            workbook = load_workbook(path)
            sheet = workbook["COACHING"]
            sheet["L5"] = "key-migrate"
            sheet["M5"] = "call-migrate"
            sheet["N5"] = "Completed"
            sheet["O5"] = "TL 1"
            workbook.save(path)
            workbook.close()

            self.assertEqual(ensure_pcs_tracker(config, **_rows()), path)
            self.assertEqual(_tracker_contract_version(path), PCS_TRACKER_VERSION)
            archives = list((config.reports / "Archive").rglob("PCS Live Tracker_pre_*.xlsx"))
            self.assertEqual(len(archives), 1)
            migrated = load_workbook(path, read_only=True, data_only=True)
            try:
                self.assertEqual(migrated["COACHING"]["L5"].value, "key-migrate")
                self.assertEqual(migrated["COACHING"]["M5"].value, "call-migrate")
                self.assertEqual(migrated["COACHING"]["N5"].value, "Completed")
                self.assertEqual(migrated["COACHING"]["O5"].value, "TL 1")
            finally:
                migrated.close()

    def test_latest_report_is_the_fixed_permanent_tracker(self):
        with tempfile.TemporaryDirectory() as folder:
            reports = Path(folder)
            config = SimpleNamespace(reports=reports)
            tracker = reports / PCS_TRACKER_FILENAME
            tracker.write_bytes(b"tracker")
            self.assertEqual(latest_pcs_report(config), tracker)

    def test_latest_update_loads_only_pcs_sources_and_publishes_csv(self):
        with tempfile.TemporaryDirectory() as folder:
            home = Path(folder)
            report = home / "Reports" / PCS_TRACKER_FILENAME
            report.parent.mkdir()
            report.write_bytes(b"tracker")
            config = SimpleNamespace(reports=report.parent, feed=home / "Feed")
            with (
                patch("wfmhub.cli.refresh", return_value=0) as run,
                patch("wfmhub.cli.load_config", return_value=config),
                patch("wfmhub.cli.latest_pcs_report", return_value=report),
            ):
                _build_latest_pcs_report(home)
            run.assert_called_once_with(home, None, None, ("pcs",), "pcs", False)

    def test_fast_update_uses_current_database_without_source_ingestion(self):
        with tempfile.TemporaryDirectory() as folder:
            home = Path(folder)
            report = home / "Reports" / PCS_TRACKER_FILENAME
            config = SimpleNamespace(reports=report.parent, feed=home / "Feed")
            conn = MagicMock()
            with (
                patch("wfmhub.cli.load_config", return_value=config),
                patch("wfmhub.cli.write_session", return_value=nullcontext(conn)),
                patch(
                    "wfmhub.cli._pcs_mart_period",
                    return_value=(date(2026, 8, 1), date(2026, 9, 8)),
                ),
                patch("wfmhub.cli.build_report_pack", return_value=report) as build,
            ):
                self.assertEqual(_build_pcs_from_database(home), report)
            build.assert_called_once_with(
                "pcs", conn, config, date(2026, 8, 1), date(2026, 9, 8),
            )

    def test_install_uses_desktop_excel_helper_once(self):
        with tempfile.TemporaryDirectory() as folder:
            home = Path(folder)
            report = home / "Reports" / PCS_TRACKER_FILENAME
            report.parent.mkdir()
            report.write_bytes(b"tracker")
            config = SimpleNamespace(reports=report.parent)
            with (
                patch("wfmhub.cli.load_config", return_value=config),
                patch("wfmhub.cli.latest_pcs_report", return_value=report),
                patch("wfmhub.cli.run_pcs_excel_action", return_value="ready") as run,
            ):
                _install_pcs_power_query(home)
            run.assert_called_once_with(config, report, "Install", "LOCAL")

    def test_pcs_refresh_uses_targeted_model_without_all_shared_feeds(self):
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
