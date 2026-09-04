"""Corporate-runtime and local-storage diagnostics with no external installs."""

from __future__ import annotations

import hashlib
import os
import sqlite3
import sys
import tempfile
import tomllib
from pathlib import Path

import _sqlite3
import et_xmlfile
import openpyxl
import xlsxwriter

from . import __version__


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def _database_path(home: Path) -> Path:
    config_file = home / "config" / "wfmhub.toml"
    if not config_file.exists():
        config_file = home / "config" / "default.toml"
    with config_file.open("rb") as handle:
        value = str(tomllib.load(handle).get("paths", {}).get("database", "_system/database/wfm.sqlite3"))
    configured = Path(value)
    return configured.resolve() if configured.is_absolute() else (home / configured).resolve()


def _check_runtime_manifest(home: Path) -> tuple[bool, str]:
    manifest = home / "_system" / "RUNTIME_MANIFEST.sha256"
    if not manifest.exists():
        return True, "development checkout; packaged runtime manifest not present"
    checked = 0
    for line in manifest.read_text(encoding="utf-8").splitlines():
        if not line.strip():
            continue
        expected, relative = line.split("  ", 1)
        path = home / "_system" / "runtime" / relative
        if not path.is_file():
            return False, f"missing reviewed runtime file: {path}"
        actual = _sha256(path)
        if actual != expected:
            return False, f"runtime hash mismatch: {path}"
        checked += 1
    return True, f"{checked} official CPython native files match the reviewed manifest"


def _check_sqlite_and_excel(home: Path) -> tuple[bool, str]:
    with tempfile.TemporaryDirectory(prefix=".wfmhub-doctor-", dir=home) as folder:
        temporary = Path(folder)
        source_path = temporary / "doctor.sqlite3"
        backup_path = temporary / "doctor-backup.sqlite3"
        connection = sqlite3.connect(source_path, isolation_level=None)
        try:
            connection.execute("PRAGMA journal_mode=WAL")
            connection.execute("CREATE TABLE test(value TEXT NOT NULL)")
            connection.execute("INSERT INTO test VALUES ('committed')")
            connection.execute("BEGIN")
            connection.execute("INSERT INTO test VALUES ('must roll back')")
            connection.execute("ROLLBACK")
            if connection.execute("SELECT count(*) FROM test").fetchone()[0] != 1:
                return False, "SQLite rollback test returned the wrong row count"
            backup = sqlite3.connect(backup_path)
            try:
                connection.backup(backup)
            finally:
                backup.close()
        finally:
            connection.close()
        verify = sqlite3.connect(backup_path)
        try:
            if verify.execute("PRAGMA quick_check").fetchone()[0].lower() != "ok":
                return False, "SQLite backup quick-check failed"
            if verify.execute("SELECT value FROM test").fetchone()[0] != "committed":
                return False, "SQLite backup content test failed"
        finally:
            verify.close()

        workbook_path = temporary / "doctor.xlsx"
        workbook = xlsxwriter.Workbook(workbook_path)
        workbook.add_worksheet("CHECK").write("A1", "WFMHub")
        workbook.close()
        loaded = openpyxl.load_workbook(workbook_path, read_only=True, data_only=True)
        try:
            if loaded["CHECK"]["A1"].value != "WFMHub":
                return False, "Excel write/read test returned the wrong value"
        finally:
            loaded.close()
    return True, "create, transaction, WAL backup, XLSX write and XLSX read all passed"


def run_doctor(home: Path) -> bool:
    home = home.resolve()
    home.mkdir(parents=True, exist_ok=True)
    print("\nWFMHUB SYSTEM CHECK")
    print(f"WFMHub     : {__version__}")
    print(f"Python     : {sys.version.split()[0]} ({sys.executable})")
    print(f"SQLite     : {sqlite3.sqlite_version}")
    print(f"_sqlite3   : {Path(_sqlite3.__file__).resolve()}")
    print(f"openpyxl   : {openpyxl.__version__}")
    print(f"XlsxWriter : {xlsxwriter.__version__}")
    print(f"et_xmlfile : {getattr(et_xmlfile, '__version__', 'loaded')}")
    if os.name == "nt":
        sqlite_dll = Path(sys.executable).resolve().parent / "sqlite3.dll"
        print(f"sqlite3.dll: {sqlite_dll} ({'present' if sqlite_dll.is_file() else 'MISSING'})")
        print("App Control: _sqlite3 loaded successfully; Windows accepted the official CPython binary")

    failures: list[str] = []
    if sqlite3.sqlite_version_info < (3, 35, 0):
        failures.append(f"SQLite 3.35+ is required; found {sqlite3.sqlite_version}")

    configured_database = _database_path(home)
    if configured_database.suffix.lower() == ".duckdb":
        print(f"Old setting : {configured_database} (setup will preserve it and select SQLite)")
        database = home / "_system" / "database" / "wfm.sqlite3"
    else:
        database = configured_database
    database.parent.mkdir(parents=True, exist_ok=True)
    print(f"Database   : {database}")
    legacy = home / "database" / "wfm.duckdb"
    if legacy.exists():
        print(f"Legacy DB  : detected and will remain untouched ({legacy})")

    for label, check in (
        ("Runtime", lambda: _check_runtime_manifest(home)),
        ("Storage", lambda: _check_sqlite_and_excel(home)),
    ):
        try:
            ok, details = check()
        except PermissionError as exc:
            ok, details = False, f"write permission blocked: {exc}"
        except (ImportError, OSError, sqlite3.Error) as exc:
            ok, details = False, str(exc)
        print(f"{label:11}: {'PASS' if ok else 'FAIL'} - {details}")
        if not ok:
            failures.append(f"{label}: {details}")

    if failures:
        print("\nSYSTEM CHECK FAILED")
        for failure in failures:
            print(f"- {failure}")
        return False
    print("\nSYSTEM CHECK PASSED")
    return True
