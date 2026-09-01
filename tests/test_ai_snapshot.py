from __future__ import annotations

import hashlib
import io
import json
import shutil
import sqlite3
import stat
import tempfile
import unittest
from contextlib import redirect_stdout
from datetime import date, datetime
from pathlib import Path

from wfmhub.ai_snapshot import create_analysis_snapshot
from wfmhub.cli import main
from wfmhub.config import ensure_user_config, load_config, write_source_root
from wfmhub.database import connect, migrate
from wfmhub.mapping import load_queue_mapping
from wfmhub.rules import load_rulebook


REPO = Path(__file__).resolve().parents[1]


class AnalysisSnapshotTests(unittest.TestCase):
    def make_config(self, folder: str):
        home = Path(folder) / "hub"
        (home / "config").mkdir(parents=True)
        shutil.copy2(REPO / "config" / "default.toml", home / "config" / "default.toml")
        shutil.copy2(
            REPO / "config" / "default_rules.toml",
            home / "config" / "default_rules.toml",
        )
        shutil.copy2(
            REPO / "config" / "default_queue_mapping.csv",
            home / "config" / "default_queue_mapping.csv",
        )
        shutil.copytree(REPO / "sql", home / "sql")
        config_file = ensure_user_config(home)
        write_source_root(config_file, Path(folder) / "source")
        config = load_config(home)
        migrate(config)
        return config

    def seed_governed_data(self, config) -> None:
        rulebook = load_rulebook(config.home, config.business_rules)
        mapping = load_queue_mapping(config.queue_mapping)
        conn = connect(config)
        try:
            conn.execute(
                """INSERT INTO meta.refresh_run(
                       run_id, started_at, finished_at, requested_start,
                       requested_end, status, files_loaded
                   ) VALUES (?, ?, ?, ?, ?, 'SUCCESS', 4)""",
                [
                    "refresh-1", datetime(2026, 8, 31, 1), datetime(2026, 8, 31, 2),
                    date(2026, 8, 1), date(2026, 8, 31),
                ],
            )
            conn.execute(
                """INSERT INTO meta.rule_application(
                       run_id, rule_version, rule_sha256, rule_file,
                       effective_from, applied_at
                   ) VALUES (?, ?, ?, ?, ?, ?)""",
                [
                    "model-1", rulebook.version, rulebook.sha256, "wfm_rules.toml",
                    date(2026, 1, 1), datetime(2026, 8, 31, 2),
                ],
            )
            conn.execute(
                """INSERT INTO meta.mapping_application(
                       run_id, mapping_sha256, mapping_file, applied_at
                   ) VALUES (?, ?, ?, ?)""",
                ["model-1", mapping.sha256, "queue_mapping.csv", datetime(2026, 8, 31, 2)],
            )
            conn.execute(
                """INSERT INTO mart.source_health(
                       source_family, expected_path, newest_file, newest_business_date,
                       modified_at, loaded_at, row_count, rejected_count, status,
                       details, scoped_out_count
                   ) VALUES ('apde', '/private/extract/path', 'apde.csv', '2026-08-31',
                             '2026-08-31 01:00:00', '2026-08-31 02:00:00',
                             10, 1, 'SUCCESS', 'Loaded successfully', 2)"""
            )
            for business_date in ("2026-08-15", "2026-07-31"):
                conn.execute(
                    """INSERT INTO mart.service_interval(
                           business_date, interval_start, source_system, lob, language,
                           offered, answered, abandoned, short_abandoned,
                           answered_within_target, handled_seconds, sl_profile,
                           sl_target, mapping_status, rule_version, rule_sha256
                       ) VALUES (?, ?, 'APDE', 'RSA', 'NL', 10, 8, 2, 1, 7,
                                 800, 'default', 0.80, 'MAPPED', ?, ?)""",
                    [business_date, f"{business_date} 09:00:00", rulebook.version, rulebook.sha256],
                )
                conn.execute(
                    """INSERT INTO mart.staffing_interval(
                           business_date, interval_start, interval_end, lob, language,
                           scheduled_agents, observed_agents, productive_agents,
                           auxiliary_agents, scheduled_fte, elapsed_scheduled_fte,
                           observed_fte, productive_fte, staffing_variance_fte,
                           staffing_gap_fte, staffing_state, evidence_basis,
                           evaluation_as_of
                       ) VALUES (?, ?, ?, 'RSA', 'NL', 10, 8, 7, 1, 10, 10,
                                 8, 7, -2, 2, 'GAP', 'AGENT_STATUS', ?)""",
                    [
                        business_date, f"{business_date} 09:00:00",
                        f"{business_date} 09:15:00", f"{business_date} 09:10:00",
                    ],
                )
                conn.execute(
                    """INSERT INTO mart.agent_pcs_day(
                           agent_day_key, business_date, agent_id, team_leader,
                           lob, language, inbound_calls, pcs_enabled_calls,
                           pcs_status_calls, pcs_participation_responses,
                           survey_responses, pcs_score_sum, low_score_responses,
                           top_box_responses, pcs_invalid_responses
                       ) VALUES (?, ?, ?, 'Lead A', 'RSA', 'NL', 10, 9, 8, 7,
                                 6, 24, 2, 4, 1)""",
                    [f"{business_date}|A1", business_date, "A1"],
                )
        finally:
            conn.close()

    def test_snapshot_is_bounded_curated_and_provenanced(self):
        with tempfile.TemporaryDirectory() as folder:
            config = self.make_config(folder)
            self.seed_governed_data(config)
            source = connect(config, read_only=True)
            try:
                source_tables_before = {
                    row[0]
                    for row in source.execute(
                        "SELECT name FROM sqlite_master WHERE type='table'"
                    ).fetchall()
                }
            finally:
                source.close()
            modified_before = config.database.stat().st_mtime_ns
            output = Path(folder) / "analysis-bundle"
            source = connect(config, read_only=True)
            try:
                result = create_analysis_snapshot(
                    source, config, date(2026, 8, 1), date(2026, 8, 31), output
                )
            finally:
                source.close()

            self.assertEqual(result.bundle_dir, output.resolve())
            self.assertEqual(
                result.row_counts,
                {
                    "source_health": 1,
                    "daily_service_lob": 1,
                    "daily_staffing_gaps": 1,
                    "pcs_team_day": 1,
                },
            )
            self.assertFalse(result.database.stat().st_mode & stat.S_IWUSR)
            self.assertEqual(config.database.stat().st_mtime_ns, modified_before)
            source = connect(config, read_only=True)
            try:
                source_tables_after = {
                    row[0]
                    for row in source.execute(
                        "SELECT name FROM sqlite_master WHERE type='table'"
                    ).fetchall()
                }
            finally:
                source.close()
            self.assertEqual(source_tables_after, source_tables_before)

            snapshot = sqlite3.connect(result.database.as_uri() + "?mode=ro", uri=True)
            try:
                tables = {
                    row[0]
                    for row in snapshot.execute(
                        "SELECT name FROM sqlite_master WHERE type='table'"
                    ).fetchall()
                }
                self.assertEqual(
                    tables,
                    {
                        "source_health", "daily_service_lob",
                        "daily_staffing_gaps", "pcs_team_day",
                    },
                )
                self.assertEqual(
                    snapshot.execute(
                        "SELECT business_date, service_level FROM daily_service_lob"
                    ).fetchone(),
                    ("2026-08-15", 7 / 9),
                )
                with self.assertRaises(sqlite3.OperationalError):
                    snapshot.execute("DELETE FROM daily_service_lob")
            finally:
                snapshot.close()

            manifest = json.loads(result.manifest.read_text(encoding="utf-8"))
            self.assertEqual(manifest["period"]["start"], "2026-08-01")
            self.assertEqual(
                manifest["provenance"]["latest_successful_refresh"]["run_id"],
                "refresh-1",
            )
            self.assertEqual(manifest["provenance"]["model_run_id"], "model-1")
            self.assertTrue(manifest["provenance"]["configuration_matches_model"])
            self.assertEqual(
                [dataset["table"] for dataset in manifest["datasets"]],
                ["source_health", "daily_service_lob", "daily_staffing_gaps", "pcs_team_day"],
            )
            self.assertFalse(manifest["safety"]["contains_raw_extracts"])
            self.assertNotIn(
                "/private/extract/path",
                json.dumps(manifest),
            )
            digest = hashlib.sha256(result.database.read_bytes()).hexdigest()
            self.assertEqual(manifest["files"][0]["sha256"], digest)
            self.assertEqual(manifest["files"][0]["bytes"], result.database.stat().st_size)

    def test_snapshot_rejects_reverse_period_and_existing_output(self):
        with tempfile.TemporaryDirectory() as folder:
            config = self.make_config(folder)
            source = connect(config, read_only=True)
            try:
                with self.assertRaisesRegex(ValueError, "start date"):
                    create_analysis_snapshot(
                        source, config, date(2026, 8, 31), date(2026, 8, 1)
                    )
                output = Path(folder) / "existing"
                output.mkdir()
                with self.assertRaisesRegex(FileExistsError, "already exists"):
                    create_analysis_snapshot(
                        source, config, date(2026, 8, 1), date(2026, 8, 31), output
                    )
            finally:
                source.close()

    def test_snapshot_requires_a_query_only_source_connection(self):
        with tempfile.TemporaryDirectory() as folder:
            config = self.make_config(folder)
            source = connect(config)
            try:
                with self.assertRaisesRegex(ValueError, "must be read-only"):
                    create_analysis_snapshot(
                        source, config, date(2026, 8, 1), date(2026, 8, 31)
                    )
            finally:
                source.close()

    def test_cli_creates_the_snapshot_without_refreshing_models(self):
        with tempfile.TemporaryDirectory() as folder:
            config = self.make_config(folder)
            self.seed_governed_data(config)
            output = Path(folder) / "cli-bundle"
            captured = io.StringIO()
            with redirect_stdout(captured):
                exit_code = main([
                    "--home", str(config.home), "analysis-snapshot",
                    "--start", "2026-08-01", "--end", "2026-08-31",
                    "--output", str(output),
                ])
            self.assertEqual(exit_code, 0)
            self.assertTrue((output / "wfmhub_analysis.sqlite3").is_file())
            self.assertIn("opened read-only; no models were rebuilt", captured.getvalue())


if __name__ == "__main__":
    unittest.main()
