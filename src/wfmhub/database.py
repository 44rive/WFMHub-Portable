"""DuckDB lifecycle, migrations, locking and safe backup helpers."""

from __future__ import annotations

import os
import re
import shutil
import socket
from contextlib import contextmanager
from datetime import datetime
from pathlib import Path
from typing import Iterator

import duckdb

from .config import Config


class HubLockedError(RuntimeError):
    pass


class ProcessLock:
    def __init__(self, path: Path):
        self.path = path
        self.fd: int | None = None

    def __enter__(self) -> "ProcessLock":
        self.path.parent.mkdir(parents=True, exist_ok=True)
        if self.path.exists():
            owner = self.path.read_text(encoding="utf-8", errors="replace")
            match = re.search(r"\bpid=(\d+)\b", owner)
            stale = False
            if match:
                try:
                    os.kill(int(match.group(1)), 0)
                except ProcessLookupError:
                    stale = True
                except PermissionError:
                    stale = False
            if stale:
                self.path.unlink()
        try:
            self.fd = os.open(self.path, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
        except FileExistsError as exc:
            owner = self.path.read_text(encoding="utf-8", errors="replace") if self.path.exists() else "unknown"
            raise HubLockedError(
                "Another WFMHub refresh appears to be running. "
                f"Lock: {self.path}. Owner: {owner.strip()}"
            ) from exc
        payload = f"pid={os.getpid()} host={socket.gethostname()} started={datetime.now().isoformat(timespec='seconds')}\n"
        os.write(self.fd, payload.encode("utf-8"))
        return self

    def __exit__(self, exc_type, exc, tb) -> None:
        if self.fd is not None:
            os.close(self.fd)
        try:
            self.path.unlink()
        except FileNotFoundError:
            pass


def project_sql_dir(config: Config) -> Path:
    direct = config.home / "sql"
    if direct.exists():
        return direct
    packaged = config.home / "app" / "sql"
    if packaged.exists():
        return packaged
    raise FileNotFoundError("Cannot locate the sql directory")


def connect(config: Config, read_only: bool = False) -> duckdb.DuckDBPyConnection:
    config.database.parent.mkdir(parents=True, exist_ok=True)
    return duckdb.connect(str(config.database), read_only=read_only)


def _migration_files(config: Config) -> list[Path]:
    return sorted((project_sql_dir(config) / "migrations").glob("*.sql"))


def _backup_if_migrations_pending(config: Config) -> Path | None:
    if not config.database.exists():
        return None
    present: set[str] = set()
    probe = duckdb.connect(str(config.database), read_only=True)
    try:
        try:
            present = {row[0] for row in probe.execute("SELECT version FROM meta.schema_migration").fetchall()}
        except duckdb.Error:
            present = set()
    finally:
        probe.close()
    pending = [path for path in _migration_files(config) if path.stem not in present]
    if not pending:
        return None
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    target = config.backups / f"wfm_pre_migration_{stamp}.duckdb"
    target.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(config.database, target)
    return target


def migrate(config: Config, connection: duckdb.DuckDBPyConnection | None = None) -> list[str]:
    own = connection is None
    if own:
        _backup_if_migrations_pending(config)
    conn = connection or connect(config)
    applied: list[str] = []
    try:
        conn.execute("CREATE SCHEMA IF NOT EXISTS meta")
        conn.execute(
            "CREATE TABLE IF NOT EXISTS meta.schema_migration "
            "(version VARCHAR PRIMARY KEY, applied_at TIMESTAMP NOT NULL DEFAULT current_timestamp)"
        )
        present = {row[0] for row in conn.execute("SELECT version FROM meta.schema_migration").fetchall()}
        for path in _migration_files(config):
            version = path.stem
            if version in present:
                continue
            conn.execute("BEGIN TRANSACTION")
            try:
                conn.execute(path.read_text(encoding="utf-8"))
                conn.execute("INSERT INTO meta.schema_migration(version) VALUES (?)", [version])
                conn.execute("COMMIT")
            except Exception:
                conn.execute("ROLLBACK")
                raise
            applied.append(version)
        return applied
    finally:
        if own:
            conn.close()


@contextmanager
def write_session(config: Config) -> Iterator[duckdb.DuckDBPyConnection]:
    lock = config.database.with_suffix(config.database.suffix + ".lock")
    with ProcessLock(lock):
        _backup_if_migrations_pending(config)
        conn = connect(config)
        try:
            migrate(config, conn)
            yield conn
        finally:
            conn.close()


def backup_database(config: Config) -> Path:
    if not config.database.exists():
        raise FileNotFoundError(f"Database does not exist yet: {config.database}")
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    target = config.backups / f"wfm_{stamp}.duckdb"
    target.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(config.database, target)
    return target
