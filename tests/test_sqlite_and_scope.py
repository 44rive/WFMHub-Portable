from __future__ import annotations

import csv
import shutil
import sqlite3
import tempfile
import unittest
from datetime import date, datetime
from pathlib import Path
from unittest.mock import patch

from openpyxl import Workbook

from wfmhub.config import ensure_user_config, load_config, write_source_root
from wfmhub.database import DatabaseFormatError, backup_database, connect, migrate, write_session
from wfmhub.ingestion import AgentScope, ingest_all


REPO = Path(__file__).resolve().parents[1]


def make_home(folder: str) -> tuple[Path, Path]:
    root = Path(folder)
    home = root / "hub"
    source = root / "source"
    (home / "config").mkdir(parents=True)
    for name in (
        "default.toml", "default_rules.toml", "default_metrics.toml",
        "default_analytics.toml", "default_reports.toml",
    ):
        shutil.copy2(REPO / "config" / name, home / "config" / name)
    shutil.copytree(REPO / "sql", home / "sql")
    config_file = ensure_user_config(home)
    write_source_root(config_file, source)
    return home, source


def make_fte(path: Path, agents: list[tuple[str, str]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Agent"
    sheet.append(["Client ID", "Status", "Name", "Team leader", "Ops Manager", "LOB", "Market", "Language", "Location", "City", "FTE", "End date if leaver"])
    for agent_id, name in agents:
        sheet.append([agent_id, "Active", name, "TL", "Ops", "LOB", "Market", "EN", "Office", "City", 1, None])
    workbook.save(path)


def make_lilo(path: Path, first_100: str = "2026-08-01 08:00:00") -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.writer(handle)
        writer.writerow(["[Agent]", "[Agent ID]", "[First Log-on Time]", "[Last Log-off Time]"])
        writer.writerow(["Agent 100", "100", first_100, "2026-08-01 16:00:00"])
        writer.writerow(["Agent 200", "200", "2026-08-01 08:00:00", "2026-08-01 16:00:00"])


class AgentScopeTests(unittest.TestCase):
    def test_scope_fingerprint_reloads_unchanged_extract_when_roster_expands(self):
        with tempfile.TemporaryDirectory() as folder:
            home, source = make_home(folder)
            fte = source / "FTE" / "FTE Count.xlsx"
            lilo = source / "Storm" / "LILO" / "AP-Historical-Report---Agent-Login 2026-08-01.csv"
            make_fte(fte, [("100", "Agent 100")])
            make_lilo(lilo)
            config = load_config(home)
            with write_session(config) as conn:
                first = ingest_all(conn, config)
                self.assertEqual(first.scoped_out, 1)
                self.assertEqual(conn.execute("SELECT count(*) FROM raw.lilo r JOIN meta.source_file f ON f.file_id=r.source_file_id AND f.active").fetchone()[0], 1)

                make_fte(fte, [("100", "Agent 100"), ("200", "Agent 200")])
                second = ingest_all(conn, config)
                self.assertEqual(second.failed, 0)
                self.assertEqual(conn.execute("SELECT count(*) FROM raw.lilo r JOIN meta.source_file f ON f.file_id=r.source_file_id AND f.active").fetchone()[0], 2)
                self.assertEqual(conn.execute("SELECT count(*) FROM meta.source_file WHERE source_family='lilo' AND active=true").fetchone()[0], 1)
                self.assertEqual(conn.execute("SELECT count(DISTINCT scope_fingerprint) FROM meta.source_file WHERE source_family='lilo'").fetchone()[0], 2)

    def test_same_path_a_b_a_reactivates_original_immutable_rows(self):
        with tempfile.TemporaryDirectory() as folder:
            home, source = make_home(folder)
            make_fte(source / "FTE" / "FTE Count.xlsx", [("100", "Agent 100")])
            lilo = source / "Storm" / "LILO" / "AP-Historical-Report---Agent-Login 2026-08-01.csv"
            make_lilo(lilo)
            original = lilo.read_bytes()
            config = load_config(home)
            with write_session(config) as conn:
                ingest_all(conn, config)
                make_lilo(lilo, "2026-08-01 09:00:00")
                ingest_all(conn, config)
                lilo.write_bytes(original)
                reverted = ingest_all(conn, config)
                self.assertEqual(reverted.failed, 0)
                active_first = conn.execute(
                    """SELECT r.first_login FROM raw.lilo r
                       JOIN meta.source_file f ON f.file_id=r.source_file_id AND f.active
                       WHERE r.agent_id='100'"""
                ).fetchone()[0]
                self.assertEqual(active_first, datetime(2026, 8, 1, 8, 0))
                self.assertEqual(conn.execute("SELECT count(*) FROM meta.source_file WHERE source_family='lilo' AND active=true").fetchone()[0], 1)

    def test_unique_name_can_scope_a_populated_operational_id(self):
        scope = AgentScope(frozenset({"100"}), {"agent one": "100"}, "fingerprint")
        self.assertEqual(scope.resolve(None, "Agent One"), "100")
        self.assertEqual(scope.resolve("999", "Agent One"), "999")
        self.assertIsNone(scope.resolve("999", "Someone Else"))

    def test_scope_keeps_active_and_leaver_only_through_leave_date(self):
        scope = AgentScope(
            frozenset({"100", "200", "300"}),
            {"active agent": "100", "leaver agent": "200", "transfer agent": "300"},
            "fingerprint",
            {
                "100": ("Active", None),
                "200": ("Leaver", date(2026, 8, 15)),
                "300": ("Transfer", date(2026, 8, 15)),
            },
        )
        self.assertEqual(scope.resolve("100", "Active Agent", date(2026, 9, 1)), "100")
        self.assertEqual(scope.resolve("200", "Leaver Agent", date(2026, 8, 15)), "200")
        self.assertIsNone(scope.resolve("200", "Leaver Agent", date(2026, 8, 16)))
        self.assertIsNone(scope.resolve("300", "Transfer Agent", date(2026, 8, 1)))

    def test_ingestion_applies_fte_status_on_each_business_date(self):
        with tempfile.TemporaryDirectory() as folder:
            home, source = make_home(folder)
            fte = source / "FTE" / "FTE Count.xlsx"
            fte.parent.mkdir(parents=True)
            workbook = Workbook()
            sheet = workbook.active
            sheet.title = "Agent"
            sheet.append([
                "Client ID", "Status", "Name", "Team leader", "Ops Manager",
                "LOB", "Market", "Language", "Location", "City", "FTE",
                "End date if leaver",
            ])
            sheet.append(["100", "Active", "Active Agent", "TL", "Ops", "LOB", "M", "EN", "O", "C", 1, None])
            sheet.append(["200", "Leaver", "Leaver Agent", "TL", "Ops", "LOB", "M", "EN", "O", "C", 1, date(2026, 8, 1)])
            sheet.append(["300", "Transfer", "Transfer Agent", "TL", "Ops", "LOB", "M", "EN", "O", "C", 1, date(2026, 8, 1)])
            workbook.save(fte)
            lilo = source / "Storm" / "LILO" / "LILO 2026-08-01 - 2026-08-02.csv"
            lilo.parent.mkdir(parents=True)
            with lilo.open("w", encoding="utf-8-sig", newline="") as handle:
                writer = csv.writer(handle)
                writer.writerow(["Date", "[Agent]", "[Agent ID]", "[First Log-on Time]", "[Last Log-off Time]"])
                for day in ("2026-08-01", "2026-08-02"):
                    for agent_id, name in (("100", "Active Agent"), ("200", "Leaver Agent"), ("300", "Transfer Agent")):
                        writer.writerow([day, name, agent_id, f"{day} 08:00:00", f"{day} 16:00:00"])
            config = load_config(home)
            with write_session(config) as conn:
                result = ingest_all(conn, config)
                kept = conn.execute(
                    """SELECT r.agent_id, r.extract_date FROM raw.lilo r
                       JOIN meta.source_file f ON f.file_id=r.source_file_id AND f.active
                       ORDER BY r.agent_id, r.extract_date"""
                ).fetchall()
            self.assertEqual(result.failed, 0)
            self.assertEqual(kept, [
                ("100", date(2026, 8, 1)),
                ("100", date(2026, 8, 2)),
                ("200", date(2026, 8, 1)),
            ])


class SQLiteLifecycleTests(unittest.TestCase):
    def test_v051_database_upgrades_additively_to_governed_exports(self):
        with tempfile.TemporaryDirectory() as folder:
            home = Path(folder) / "hub"
            (home / "config").mkdir(parents=True)
            for name in (
                "default.toml", "default_rules.toml", "default_metrics.toml",
                "default_analytics.toml", "default_reports.toml", "default_queue_mapping.csv",
            ):
                shutil.copy2(REPO / "config" / name, home / "config" / name)
            migrations = home / "sql" / "migrations"
            migrations.mkdir(parents=True)
            for name in (
                "001_initial.sql", "002_performance_indexes.sql",
                "003_call_pcs.sql", "004_sota_rules_absence_service.sql",
                "005_observed_attendance_mapping_reconciliation.sql",
            ):
                shutil.copy2(REPO / "sql" / "migrations" / name, migrations / name)
            config = load_config(home)
            migrate(config)
            conn = connect(config)
            try:
                conn.execute(
                    "INSERT INTO core.correction_action VALUES (?, NULL, 'Open', NULL, NULL, NULL, ?, 'v051')",
                    ["keep-v051", datetime.now()],
                )
            finally:
                conn.close()
            shutil.copy2(
                REPO / "sql" / "migrations" / "006_governed_operations_exports.sql",
                migrations / "006_governed_operations_exports.sql",
            )
            self.assertEqual(migrate(config), ["006_governed_operations_exports"])
            upgraded = connect(config, read_only=True)
            try:
                self.assertEqual(
                    upgraded.execute("SELECT count(*) FROM core.correction_action WHERE correction_id='keep-v051'").fetchone()[0],
                    1,
                )
                source_columns = {row[1] for row in upgraded.execute("PRAGMA table_info(meta_source_file)").fetchall()}
                attendance_columns = {row[1] for row in upgraded.execute("PRAGMA table_info(mart_attendance_agent_day)").fetchall()}
                self.assertIn("source_variant", source_columns)
                self.assertIn("call_action", attendance_columns)
                for table in (
                    "mart_staffing_interval", "mart_shift_timeline_segment",
                    "mart_correction_residual_segment",
                    "mart_verint_final_absence_agent_day",
                ):
                    self.assertIsNotNone(
                        upgraded.execute("SELECT name FROM sqlite_master WHERE name=?", [table]).fetchone()
                    )
            finally:
                upgraded.close()
            self.assertEqual(len(list(config.backups.glob("wfm_pre_migration_*.sqlite3"))), 1)

    def test_v040_database_upgrades_to_observed_reconciliation_and_mapping(self):
        with tempfile.TemporaryDirectory() as folder:
            home = Path(folder) / "hub"
            (home / "config").mkdir(parents=True)
            for name in (
                "default.toml", "default_rules.toml", "default_metrics.toml",
                "default_analytics.toml", "default_reports.toml", "default_queue_mapping.csv",
            ):
                shutil.copy2(REPO / "config" / name, home / "config" / name)
            migrations = home / "sql" / "migrations"
            migrations.mkdir(parents=True)
            for name in (
                "001_initial.sql", "002_performance_indexes.sql",
                "003_call_pcs.sql", "004_sota_rules_absence_service.sql",
            ):
                shutil.copy2(REPO / "sql" / "migrations" / name, migrations / name)
            config = load_config(home)
            migrate(config)
            conn = connect(config)
            try:
                conn.execute(
                    "INSERT INTO core.correction_action VALUES (?, NULL, 'Open', NULL, NULL, NULL, ?, 'v040')",
                    ["keep-v040", datetime.now()],
                )
            finally:
                conn.close()
            shutil.copy2(
                REPO / "sql" / "migrations" / "005_observed_attendance_mapping_reconciliation.sql",
                migrations / "005_observed_attendance_mapping_reconciliation.sql",
            )
            self.assertEqual(migrate(config), ["005_observed_attendance_mapping_reconciliation"])
            upgraded = connect(config, read_only=True)
            try:
                self.assertEqual(upgraded.execute("SELECT count(*) FROM core.correction_action WHERE correction_id='keep-v040'").fetchone()[0], 1)
                correction_columns = {row[1] for row in upgraded.execute("PRAGMA table_info(mart_correction_candidate)").fetchall()}
                self.assertIn("verint_reconciliation", correction_columns)
                self.assertIsNotNone(upgraded.execute("SELECT name FROM sqlite_master WHERE name='mart_verint_final_exception'").fetchone())
                indexes = {row[1] for row in upgraded.execute("PRAGMA index_list(mart_forecast_hour)").fetchall()}
                self.assertIn("idx_forecast_hour_grain_v050", indexes)
                self.assertNotIn("idx_forecast_hour_grain", indexes)
            finally:
                upgraded.close()
            self.assertEqual(len(list(config.backups.glob("wfm_pre_migration_*.sqlite3"))), 1)

    def test_v02_database_upgrades_additively_to_call_pcs(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            home = root / "hub"
            (home / "config").mkdir(parents=True)
            for name in (
                "default.toml", "default_rules.toml", "default_metrics.toml",
                "default_analytics.toml", "default_reports.toml",
            ):
                shutil.copy2(REPO / "config" / name, home / "config" / name)
            migrations = home / "sql" / "migrations"
            migrations.mkdir(parents=True)
            shutil.copy2(REPO / "sql" / "migrations" / "001_initial.sql", migrations / "001_initial.sql")
            shutil.copy2(REPO / "sql" / "migrations" / "002_performance_indexes.sql", migrations / "002_performance_indexes.sql")
            config = load_config(home)
            migrate(config)
            conn = connect(config)
            try:
                conn.execute(
                    "INSERT INTO core.correction_action VALUES (?, NULL, 'Open', NULL, NULL, NULL, ?, 'v02')",
                    ["keep-me", datetime.now()],
                )
            finally:
                conn.close()
            shutil.copy2(REPO / "sql" / "migrations" / "003_call_pcs.sql", migrations / "003_call_pcs.sql")
            applied = migrate(config)
            self.assertEqual(applied, ["003_call_pcs"])
            upgraded = connect(config, read_only=True)
            try:
                self.assertEqual(upgraded.execute("SELECT count(*) FROM core.correction_action WHERE correction_id='keep-me'").fetchone()[0], 1)
                self.assertIsNotNone(upgraded.execute("SELECT name FROM sqlite_master WHERE name='raw_call_leg'").fetchone())
                self.assertIsNotNone(upgraded.execute("SELECT name FROM sqlite_master WHERE name='core_clean_call_leg'").fetchone())
            finally:
                upgraded.close()
            self.assertEqual(len(list(config.backups.glob("wfm_pre_migration_*.sqlite3"))), 1)

    def test_failed_file_can_retry_same_fingerprint(self):
        with tempfile.TemporaryDirectory() as folder:
            home, source = make_home(folder)
            make_fte(source / "FTE" / "FTE Count.xlsx", [("100", "Agent 100")])
            config = load_config(home)
            with write_session(config) as conn:
                with patch("wfmhub.ingestion._insert_rows", side_effect=RuntimeError("transient")):
                    failed = ingest_all(conn, config)
                self.assertEqual(failed.failed, 1)
                succeeded = ingest_all(conn, config)
                self.assertEqual(succeeded.failed, 0)
                self.assertEqual(succeeded.loaded, 1)
                self.assertEqual(conn.execute("SELECT count(*) FROM meta.source_file WHERE source_family='fte'").fetchone()[0], 1)
                self.assertEqual(conn.execute("SELECT status FROM meta.source_file WHERE source_family='fte'").fetchone()[0], "SUCCESS")

    def test_wal_backup_read_only_and_application_marker(self):
        with tempfile.TemporaryDirectory() as folder:
            home, _ = make_home(folder)
            config = load_config(home)
            with write_session(config) as conn:
                conn.execute(
                    "INSERT INTO core.correction_action VALUES (?, NULL, 'Open', NULL, NULL, NULL, ?, 'test')",
                    ["decision", datetime.now()],
                )
                self.assertNotEqual(conn.execute("PRAGMA application_id").fetchone()[0], 0)
            backup = backup_database(config)
            copied = sqlite3.connect(backup)
            try:
                self.assertEqual(copied.execute("PRAGMA quick_check").fetchone()[0], "ok")
                self.assertEqual(copied.execute("SELECT count(*) FROM core_correction_action").fetchone()[0], 1)
            finally:
                copied.close()
            read_only = connect(config, read_only=True)
            try:
                with self.assertRaises(sqlite3.OperationalError):
                    read_only.execute("DELETE FROM core.correction_action")
            finally:
                read_only.close()

    def test_bad_migration_rolls_back_ddl_and_version(self):
        with tempfile.TemporaryDirectory() as folder:
            home, _ = make_home(folder)
            config = load_config(home)
            migrate(config)
            (home / "sql" / "migrations" / "003_bad.sql").write_text(
                "CREATE TABLE core.should_rollback(value VARCHAR);\nTHIS IS INVALID;\n",
                encoding="utf-8",
            )
            with self.assertRaises(sqlite3.OperationalError):
                migrate(config)
            conn = connect(config, read_only=True)
            try:
                self.assertIsNone(conn.execute("SELECT name FROM sqlite_master WHERE name='core_should_rollback'").fetchone())
                self.assertIsNone(conn.execute("SELECT version FROM meta.schema_migration WHERE version='003_bad'").fetchone())
            finally:
                conn.close()

    def test_non_sqlite_file_is_refused_and_preserved(self):
        with tempfile.TemporaryDirectory() as folder:
            home, _ = make_home(folder)
            config = load_config(home)
            payload = b"not sqlite and must remain"
            config.database.parent.mkdir(parents=True, exist_ok=True)
            config.database.write_bytes(payload)
            with self.assertRaises(DatabaseFormatError):
                connect(config)
            self.assertEqual(config.database.read_bytes(), payload)

    def test_legacy_config_is_backed_up_and_duckdb_is_untouched(self):
        with tempfile.TemporaryDirectory() as folder:
            home = Path(folder) / "hub"
            (home / "config").mkdir(parents=True)
            default = (REPO / "config" / "default.toml").read_text(encoding="utf-8")
            (home / "config" / "default.toml").write_text(default, encoding="utf-8")
            for name in (
                "default_rules.toml", "default_metrics.toml",
                "default_analytics.toml", "default_reports.toml",
            ):
                shutil.copy2(REPO / "config" / name, home / "config" / name)
            (home / "config" / "wfmhub.toml").write_text(
                default.replace('database = "_system/database/wfm.sqlite3"', 'database = "database/wfm.duckdb"'),
                encoding="utf-8",
            )
            legacy = home / "database" / "wfm.duckdb"
            legacy.parent.mkdir(parents=True)
            legacy.write_bytes(b"duckdb stays")
            ensure_user_config(home)
            self.assertIn('database = "database/wfm.sqlite3"', (home / "config" / "wfmhub.toml").read_text(encoding="utf-8"))
            self.assertEqual(legacy.read_bytes(), b"duckdb stays")
            self.assertEqual(len(list((home / "config").glob("wfmhub_pre_sqlite_*.toml"))), 1)


if __name__ == "__main__":
    unittest.main()
