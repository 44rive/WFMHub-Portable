from __future__ import annotations

import shutil
import tempfile
import unittest
from datetime import date, datetime
from pathlib import Path

from wfmhub.config import ensure_user_config, load_config
from wfmhub.database import migrate, write_session
from wfmhub.ui import DashboardStatus, dashboard_text, load_dashboard_status


REPO = Path(__file__).resolve().parents[1]


class DashboardTests(unittest.TestCase):
    def make_home(self, folder: str, include_sql: bool = True) -> Path:
        home = Path(folder) / "hub"
        (home / "config").mkdir(parents=True)
        for name in (
            "default.toml", "default_rules.toml", "default_metrics.toml",
            "default_analytics.toml", "default_reports.toml",
        ):
            shutil.copy2(REPO / "config" / name, home / "config" / name)
        if include_sql:
            shutil.copytree(REPO / "sql", home / "sql")
        ensure_user_config(home)
        return home

    def test_dashboard_text_is_compact_and_contains_brand_and_latest_data(self):
        text = dashboard_text(DashboardStatus(
            state="READY",
            database_size=183_500_800,
            agents=247,
            sources_healthy=8,
            sources_total=8,
            latest_source_date=date(2026, 8, 27),
            latest_source_family="schedule",
            last_status="SUCCESS",
            last_refresh=datetime(2026, 8, 27, 9, 41),
            period_start=date(2026, 8, 1),
            period_end=date(2026, 8, 27),
            quality_reviews=2,
        ))

        self.assertIn("# # #", text)
        self.assertIn("#####", text)
        self.assertIn("made by Anass ASSRI", text)
        self.assertIn("Latest data: 2026-08-27", text)
        self.assertIn("(schedule)", text)
        self.assertIn("247 agents", text)
        self.assertTrue(all(len(line) <= 78 for line in text.splitlines()))

    def test_status_reads_latest_loaded_source_business_date(self):
        with tempfile.TemporaryDirectory() as folder:
            home = self.make_home(folder)
            config = load_config(home)
            migrate(config)
            with write_session(config) as conn:
                conn.execute(
                    """INSERT INTO meta.refresh_run(
                           run_id, started_at, finished_at, requested_start,
                           requested_end, status
                       ) VALUES (?, ?, ?, ?, ?, 'SUCCESS')""",
                    ["ui", datetime(2026, 8, 27, 9, 40), datetime(2026, 8, 27, 9, 41), date(2026, 8, 1), date(2026, 8, 27)],
                )
                conn.execute("INSERT INTO core.dim_agent(agent_id) VALUES ('100'), ('200')")
                conn.execute(
                    """INSERT INTO mart.source_health(
                           source_family, expected_path, newest_business_date,
                           row_count, rejected_count, status, details,
                           scoped_out_count
                       ) VALUES
                           ('lilo', 'lilo', '2026-08-26', 10, 0, 'SUCCESS', 'ok', 0),
                           ('schedule', 'schedule', '2026-08-27', 10, 0, 'ERROR', 'bad', 0)"""
                )
                conn.execute(
                    """INSERT INTO meta.quality_issue(
                           issue_id, run_id, detected_at, issue_type, severity, details
                       ) VALUES ('issue', 'ui', ?, 'Test', 'ERROR', 'Test error')""",
                    [datetime(2026, 8, 27, 9, 41)],
                )

            status = load_dashboard_status(home)
            self.assertEqual(status.state, "CHECK DATA")
            self.assertEqual(status.agents, 2)
            self.assertEqual((status.sources_healthy, status.sources_total), (1, 2))
            self.assertEqual(str(status.latest_source_date), "2026-08-27")
            self.assertEqual(status.latest_source_family, "schedule")
            self.assertEqual(status.quality_errors, 1)

    def test_missing_database_requests_setup_without_crashing(self):
        with tempfile.TemporaryDirectory() as folder:
            status = load_dashboard_status(self.make_home(folder, include_sql=False))
            self.assertEqual(status.state, "SETUP REQUIRED")
            self.assertEqual(status.last_status, "NEVER")


if __name__ == "__main__":
    unittest.main()
