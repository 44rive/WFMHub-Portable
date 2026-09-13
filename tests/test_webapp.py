from __future__ import annotations

import json
import shutil
import tempfile
import threading
import unittest
import urllib.request
import urllib.error
from datetime import date, datetime
from pathlib import Path

from wfmhub.config import ensure_user_config, load_config
from wfmhub.database import migrate, write_session
from wfmhub.web_data import DashboardData, DashboardFilter
from wfmhub.webapp import ConsoleServer


REPO = Path(__file__).resolve().parents[1]


def make_home(folder: str) -> Path:
    home = Path(folder) / "hub"
    (home / "config").mkdir(parents=True)
    for name in (
        "default.toml", "default_rules.toml", "default_metrics.toml",
        "default_analytics.toml", "default_reports.toml",
        "default_queue_mapping.csv", "default_capacity_mapping.csv",
        "default_service_profiles.toml",
    ):
        shutil.copy2(REPO / "config" / name, home / "config" / name)
    shutil.copytree(REPO / "sql", home / "sql")
    ensure_user_config(home)
    config = load_config(home)
    migrate(config)
    return home


class WebConsoleTests(unittest.TestCase):
    def test_governed_service_ratio_and_meta_are_available(self):
        with tempfile.TemporaryDirectory() as folder:
            home = make_home(folder)
            config = load_config(home)
            with write_session(config) as conn:
                conn.execute(
                    """INSERT INTO meta.refresh_run(
                           run_id, started_at, finished_at, status,
                           files_loaded, files_skipped, files_failed
                       ) VALUES ('web', ?, ?, 'SUCCESS', 2, 3, 0)""",
                    (datetime(2026, 9, 14, 8), datetime(2026, 9, 14, 8, 5)),
                )
                conn.execute(
                    """INSERT INTO mart.call_service_15min(
                           business_date, interval_start, interval_end,
                           source_system, service_scope, comparison_scope,
                           queue, designation, language, offered, answered,
                           abandoned, short_abandoned, abandoned_within_target,
                           answered_within_target, talk_seconds, hold_seconds,
                           wrap_seconds, handled_seconds, service_level,
                           service_availability, abandon_rate, aht_seconds,
                           call_legs, transferred_legs, source_files,
                           mapping_sha256, rule_version, rule_sha256
                       ) VALUES (
                           '2026-09-14', '2026-09-14 08:00:00',
                           '2026-09-14 08:15:00', 'CALL_BY_CALL', 'Ford FR',
                           'OEM', 'APFR_PAR_RSA_CSTRUCTR_FORD_ASSISTANCE_FR',
                           'Ford', 'FR', 100, 90, 10, 3, 2, 75,
                           18000, 0, 0, 18000, .7653, .9, .1, 200,
                           100, 0, 'fixture.csv', 'map', 'rules', 'rulehash'
                       )"""
                )
            service = DashboardData(config)
            meta = service.meta()
            self.assertEqual(meta["latest_date"], "2026-09-14")
            self.assertEqual(meta["last_refresh"]["status"], "SUCCESS")
            result = service.service(DashboardFilter(
                date(2026, 9, 14), date(2026, 9, 14), management_lob="OEM",
            ))
            self.assertAlmostEqual(result["cards"][0]["value"], 75 / 98)
            self.assertEqual(result["cards"][1]["value"], 100)
            self.assertEqual(result["cards"][2]["value"], 90)
            self.assertEqual(result["queues"][0]["handled_in_sl"], 75)
            unscoped = service.service(DashboardFilter(
                date(2026, 9, 14), date(2026, 9, 14),
            ))
            self.assertTrue(unscoped["service_scope_required"])
            self.assertIsNone(unscoped["cards"][0]["value"])
            self.assertEqual(unscoped["by_lob"][0]["service_level"], 75 / 98)

    def test_local_http_server_serves_ui_and_json(self):
        with tempfile.TemporaryDirectory() as folder:
            home = make_home(folder)
            server = ConsoleServer(("127.0.0.1", 0), load_config(home))
            thread = threading.Thread(target=server.serve_forever, daemon=True)
            thread.start()
            root = f"http://127.0.0.1:{server.server_port}"
            try:
                with urllib.request.urlopen(root + "/", timeout=5) as response:
                    html = response.read().decode("utf-8")
                    self.assertIn("WFMHub Operations Console", html)
                    self.assertEqual(response.headers["X-Frame-Options"], "DENY")
                with urllib.request.urlopen(root + "/api/meta", timeout=5) as response:
                    payload = json.load(response)
                    self.assertEqual(payload["database"], "wfm.sqlite3")
                for view in ("today", "staffing", "service", "attendance", "history"):
                    with urllib.request.urlopen(
                        root + f"/api/view/{view}?start=2026-09-14&end=2026-09-14",
                        timeout=5,
                    ) as response:
                        payload = json.load(response)
                        self.assertEqual(payload["view"], view)
                        self.assertEqual(len(payload["cards"]), 4)
                request = urllib.request.Request(
                    root + "/api/refresh", data=b"{}", method="POST",
                )
                with self.assertRaises(urllib.error.HTTPError) as rejected:
                    urllib.request.urlopen(request, timeout=5)
                self.assertEqual(rejected.exception.code, 415)
            finally:
                server.shutdown()
                server.server_close()
                thread.join(timeout=5)

    def test_console_assets_are_offline_and_no_power_bi_dependency_is_used(self):
        html = (REPO / "src" / "wfmhub" / "web" / "index.html").read_text(encoding="utf-8")
        script = (REPO / "src" / "wfmhub" / "web" / "app.js").read_text(encoding="utf-8")
        server = (REPO / "src" / "wfmhub" / "webapp.py").read_text(encoding="utf-8")
        self.assertNotIn("http://", html)
        self.assertNotIn("https://", html)
        self.assertNotIn("powerbi", (script + server).lower())
        self.assertIn('("127.0.0.1", candidate)', server)


if __name__ == "__main__":
    unittest.main()
