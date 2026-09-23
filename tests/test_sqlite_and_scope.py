"""Durable SQLite and employee-scope contracts for the 1.0 baseline."""

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

from wfmhub.config import load_config, write_source_root
from wfmhub.database import (
    DatabaseFormatError, adopt_portable_install, backup_database, connect,
    migrate, write_session,
)
from wfmhub.ingestion import AgentScope, ingest_all


ROOT = Path(__file__).resolve().parents[1]


def _home(root: Path, name: str = "hub") -> tuple[Path, Path]:
    home = root / name
    source = root / "source"
    (home / "config").mkdir(parents=True)
    for path in (ROOT / "config").glob("default*"):
        if path.is_file():
            shutil.copy2(path, home / "config" / path.name)
    shutil.copytree(ROOT / "sql", home / "sql")
    config = load_config(home)
    write_source_root(config.file, source)
    return home, source


def _fte(path: Path, agents: list[tuple[str, str]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Agent"
    sheet.append([
        "Client ID", "Status", "Name", "Team leader", "Ops Manager",
        "LOB", "Market", "Language", "Location", "City", "FTE",
        "End date if leaver",
    ])
    for agent_id, name in agents:
        sheet.append([agent_id, "Active", name, "TL", "Ops", "RSA NL", "NL", "NL", "Site", "City", 1, None])
    workbook.save(path)
    workbook.close()


def _lilo(path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.writer(handle)
        writer.writerow(["[Agent]", "[Agent ID]", "[First Log-on Time]", "[Last Log-off Time]"])
        writer.writerow(["Agent One", "001", "2026-08-01 08:00:00", "2026-08-01 16:00:00"])
        writer.writerow(["Agent Two", "002", "2026-08-01 08:00:00", "2026-08-01 16:00:00"])


def test_client_id_scope_keeps_leavers_only_through_leave_date():
    scope = AgentScope(
        frozenset({"001", "002", "003"}),
        {"agent one": "001", "agent two": "002", "agent three": "003"},
        "synthetic", {
            "001": ("Active", None),
            "002": ("Leaver", date(2026, 8, 15)),
            "003": ("Transfer", date(2026, 8, 15)),
        },
    )
    assert scope.resolve("001", "Agent One", date(2026, 9, 1)) == "001"
    assert scope.resolve("002", "Agent Two", date(2026, 8, 15)) == "002"
    assert scope.resolve("002", "Agent Two", date(2026, 8, 16)) is None
    assert scope.resolve("003", "Agent Three", date(2026, 8, 1)) is None


def test_unchanged_extract_fast_path_and_scope_expansion():
    with tempfile.TemporaryDirectory() as folder:
        home, source = _home(Path(folder))
        path = source / "FTE/FTE Count.xlsx"
        _fte(path, [("001", "Agent One")])
        _lilo(source / "Storm/LILO/lilo.csv")
        config = load_config(home)
        with write_session(config) as conn:
            first = ingest_all(conn, config)
            assert first.scoped_out == 1
            with patch("wfmhub.ingestion.file_sha256") as digest:
                repeat = ingest_all(conn, config, {"fte"})
            digest.assert_not_called()
            assert repeat.metadata_skipped == 1
            _fte(path, [("001", "Agent One"), ("002", "Agent Two")])
            expanded = ingest_all(conn, config)
            assert expanded.failed == 0
            assert conn.execute(
                "SELECT count(*) FROM raw.lilo r JOIN meta.source_file f "
                "ON f.file_id=r.source_file_id AND f.active"
            ).fetchone()[0] == 2


def test_fresh_schema_backup_and_portable_adoption_preserve_data():
    with tempfile.TemporaryDirectory() as folder:
        root = Path(folder)
        old_home, _ = _home(root, "old")
        new_home, _ = _home(root, "new")
        old_config = load_config(old_home)
        old_capacity = old_home / "config" / "capacity_mapping.csv"
        old_capacity.write_text(
            old_capacity.read_text(encoding="utf-8")
            + "Synthetic,Synthetic,Synthetic,,,,Synthetic Staff\n",
            encoding="utf-8",
        )
        (new_home / "config" / "capacity_mapping.csv").unlink(missing_ok=True)
        assert migrate(old_config) == ["001_schema"]
        with write_session(old_config) as conn:
            conn.execute(
                "INSERT INTO core.dim_agent (agent_id, canonical_name) VALUES (?,?)",
                ["001", "Synthetic Agent"],
            )
        backup = backup_database(old_config)
        assert sqlite3.connect(backup).execute("PRAGMA quick_check").fetchone()[0] == "ok"
        database, applied = adopt_portable_install(new_home, old_home)
        assert database.is_file()
        assert applied == []
        assert (new_home / "config" / "capacity_mapping.csv").read_text(encoding="utf-8") == old_capacity.read_text(encoding="utf-8")
        copied = connect(load_config(new_home), read_only=True)
        try:
            assert copied.execute("SELECT canonical_name FROM core.dim_agent WHERE agent_id='001'").fetchone()[0] == "Synthetic Agent"
            with unittest.TestCase().assertRaises(sqlite3.OperationalError):
                copied.execute("DELETE FROM core.dim_agent")
        finally:
            copied.close()


def test_bad_migration_rolls_back_and_non_sqlite_file_is_preserved():
    with tempfile.TemporaryDirectory() as folder:
        home, _ = _home(Path(folder))
        config = load_config(home)
        migrate(config)
        bad = home / "sql/migrations/002_bad.sql"
        bad.write_text("CREATE TABLE core.should_rollback(value VARCHAR);\nTHIS IS INVALID;\n", encoding="utf-8")
        with unittest.TestCase().assertRaises(sqlite3.OperationalError):
            migrate(config)
        probe = connect(config, read_only=True)
        try:
            assert probe.execute("SELECT name FROM sqlite_master WHERE name='core_should_rollback'").fetchone() is None
            assert probe.execute("SELECT version FROM meta.schema_migration WHERE version='002_bad'").fetchone() is None
        finally:
            probe.close()
    with tempfile.TemporaryDirectory() as folder:
        home, _ = _home(Path(folder))
        config = load_config(home)
        payload = b"not sqlite and must remain"
        config.database.parent.mkdir(parents=True, exist_ok=True)
        config.database.write_bytes(payload)
        with unittest.TestCase().assertRaises(DatabaseFormatError):
            connect(config)
        assert config.database.read_bytes() == payload


def load_tests(loader, tests, pattern):
    """Include function-style contracts in the dependency-free CI suite."""
    for test in (
        test_client_id_scope_keeps_leavers_only_through_leave_date,
        test_unchanged_extract_fast_path_and_scope_expansion,
        test_fresh_schema_backup_and_portable_adoption_preserve_data,
        test_bad_migration_rolls_back_and_non_sqlite_file_is_preserved,
    ):
        tests.addTest(unittest.FunctionTestCase(test))
    return tests
