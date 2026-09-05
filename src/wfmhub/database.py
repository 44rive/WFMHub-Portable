"""SQLite lifecycle, logical schemas, migrations, locking and safe backups."""

from __future__ import annotations

import os
import re
import socket
import sqlite3
from contextlib import contextmanager
from datetime import date, datetime, time
from pathlib import Path
from typing import Any, Iterator, Mapping, Sequence

from .config import Config, ConfigError


class HubLockedError(RuntimeError):
    pass


class DatabaseFormatError(ConfigError):
    pass


_SCHEMA_REFERENCE = re.compile(r"\b(meta|raw|core|mart)\.", re.IGNORECASE)
_SQLITE_HEADER = b"SQLite format 3\x00"
_APPLICATION_ID = 0x57464D48  # ASCII "WFMH"
_STORAGE_VERSION = 4


def _rewrite_sql(sql: str) -> str:
    """Map stable logical schema names to portable SQLite table prefixes."""
    return _SCHEMA_REFERENCE.sub(lambda match: f"{match.group(1).lower()}_", sql)


def _adapt_date(value: date) -> str:
    return value.isoformat()


def _adapt_datetime(value: datetime) -> str:
    return value.isoformat(sep=" ", timespec="microseconds")


def _adapt_time(value: time) -> str:
    return value.isoformat(timespec="microseconds")


def _convert_date(value: bytes) -> date:
    return date.fromisoformat(value.decode("utf-8"))


def _convert_datetime(value: bytes) -> datetime:
    return datetime.fromisoformat(value.decode("utf-8"))


def _convert_time(value: bytes) -> time:
    return time.fromisoformat(value.decode("utf-8"))


def _convert_boolean(value: bytes) -> bool:
    return value not in {b"0", b"false", b"False", b"FALSE", b""}


sqlite3.register_adapter(date, _adapt_date)
sqlite3.register_adapter(datetime, _adapt_datetime)
sqlite3.register_adapter(time, _adapt_time)
sqlite3.register_converter("DATE", _convert_date)
sqlite3.register_converter("TIMESTAMP", _convert_datetime)
sqlite3.register_converter("TIME", _convert_time)
sqlite3.register_converter("BOOLEAN", _convert_boolean)


class DatabaseConnection:
    """Small DB-API facade retaining backend-neutral logical schema names."""

    def __init__(self, connection: sqlite3.Connection):
        self.raw = connection

    @property
    def in_transaction(self) -> bool:
        return self.raw.in_transaction

    def execute(
        self,
        sql: str,
        parameters: Sequence[Any] | Mapping[str, Any] | None = None,
    ) -> sqlite3.Cursor:
        return self.raw.execute(_rewrite_sql(sql), parameters or ())

    def executemany(self, sql: str, parameters: Sequence[Sequence[Any]]) -> sqlite3.Cursor:
        return self.raw.executemany(_rewrite_sql(sql), parameters)

    def close(self) -> None:
        self.raw.close()

    def commit(self) -> None:
        self.raw.commit()

    def rollback(self) -> None:
        self.raw.rollback()

    @property
    def max_variable_number(self) -> int:
        return self.raw.getlimit(sqlite3.SQLITE_LIMIT_VARIABLE_NUMBER)


DatabaseCursor = sqlite3.Cursor


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
    candidates = (
        config.home / "sql",
        config.home / "app" / "sql",
        config.system / "app" / "sql",
        Path(__file__).resolve().parents[1] / "sql",
    )
    for candidate in candidates:
        if candidate.exists():
            return candidate
    raise FileNotFoundError("Cannot locate the sql directory")


def _validate_database_file(path: Path) -> None:
    if not path.exists() or path.stat().st_size == 0:
        return
    with path.open("rb") as handle:
        header = handle.read(len(_SQLITE_HEADER))
    if header != _SQLITE_HEADER:
        raise DatabaseFormatError(
            f"{path} is not a SQLite database. WFMHub preserved the old file. "
            "Set paths.database in config\\wfmhub.toml to database/wfm.sqlite3."
        )


def connect(config: Config, read_only: bool = False) -> DatabaseConnection:
    config.database.parent.mkdir(parents=True, exist_ok=True)
    _validate_database_file(config.database)
    if read_only:
        uri = config.database.resolve().as_uri() + "?mode=ro"
        raw = sqlite3.connect(
            uri,
            uri=True,
            isolation_level=None,
            detect_types=sqlite3.PARSE_DECLTYPES | sqlite3.PARSE_COLNAMES,
        )
    else:
        raw = sqlite3.connect(
            config.database,
            isolation_level=None,
            detect_types=sqlite3.PARSE_DECLTYPES | sqlite3.PARSE_COLNAMES,
        )
    raw.execute("PRAGMA busy_timeout=30000")
    raw.execute("PRAGMA foreign_keys=ON")
    raw.execute("PRAGMA temp_store=MEMORY")
    raw.execute("PRAGMA cache_size=-65536")
    application_id = raw.execute("PRAGMA application_id").fetchone()[0]
    if application_id not in {0, _APPLICATION_ID}:
        raw.close()
        raise DatabaseFormatError(
            f"{config.database} is SQLite, but it belongs to another application. "
            "WFMHub did not modify it. Choose database/wfm.sqlite3."
        )
    if application_id == 0:
        table_names = {
            row[0]
            for row in raw.execute("SELECT name FROM sqlite_master WHERE type='table'").fetchall()
        }
        if table_names and "meta_schema_migration" not in table_names:
            raw.close()
            raise DatabaseFormatError(
                f"{config.database} is not a WFMHub SQLite database. WFMHub did not modify it."
            )
        if not read_only:
            raw.execute(f"PRAGMA application_id={_APPLICATION_ID}")
    if read_only:
        raw.execute("PRAGMA query_only=ON")
    else:
        journal_mode = raw.execute("PRAGMA journal_mode=WAL").fetchone()[0]
        if str(journal_mode).lower() != "wal":
            raw.close()
            raise RuntimeError(
                "SQLite could not enable WAL mode. Keep WFMHub on a local writable disk, "
                "not a network or sync-managed folder."
            )
        raw.execute("PRAGMA synchronous=FULL")
        raw.execute("PRAGMA wal_autocheckpoint=1000")
        raw.execute("PRAGMA journal_size_limit=67108864")
    return DatabaseConnection(raw)


def _quick_check(conn: DatabaseConnection, label: str) -> None:
    result = conn.execute("PRAGMA quick_check").fetchone()[0]
    if str(result).lower() != "ok":
        raise DatabaseFormatError(f"SQLite quick check failed for {label}: {result}")


def _migration_files(config: Config) -> list[Path]:
    return sorted((project_sql_dir(config) / "migrations").glob("*.sql"))


def _migration_statements(script: str) -> Iterator[str]:
    buffer = ""
    for line in script.splitlines(keepends=True):
        buffer += line
        if not sqlite3.complete_statement(buffer):
            continue
        statement = buffer.strip()
        buffer = ""
        if not statement:
            continue
        if re.match(r"^CREATE\s+SCHEMA\b", statement, re.IGNORECASE):
            continue
        yield statement
    if buffer.strip():
        raise sqlite3.OperationalError("Incomplete SQL statement at the end of a migration")


def _backup_to(config: Config, target: Path) -> Path:
    target.parent.mkdir(parents=True, exist_ok=True)
    source = connect(config, read_only=True)
    destination = sqlite3.connect(target)
    failed = False
    try:
        _quick_check(source, str(config.database))
        source.raw.backup(destination)
        result = destination.execute("PRAGMA quick_check").fetchone()[0]
        if str(result).lower() != "ok":
            raise DatabaseFormatError(f"SQLite quick check failed for backup {target}: {result}")
    except Exception:
        failed = True
        raise
    finally:
        destination.close()
        source.close()
        if failed:
            target.unlink(missing_ok=True)
    return target


def _backup_if_migrations_pending(config: Config) -> Path | None:
    if not config.database.exists() or config.database.stat().st_size == 0:
        return None
    present: set[str] = set()
    probe = connect(config, read_only=True)
    try:
        try:
            present = {row[0] for row in probe.execute("SELECT version FROM meta.schema_migration").fetchall()}
        except sqlite3.Error:
            present = set()
    finally:
        probe.close()
    pending = [path for path in _migration_files(config) if path.stem not in present]
    if not pending:
        return None
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    return _backup_to(config, config.backups / f"wfm_pre_migration_{stamp}.sqlite3")


def migrate(config: Config, connection: DatabaseConnection | None = None) -> list[str]:
    own = connection is None
    if own:
        _backup_if_migrations_pending(config)
    conn = connection or connect(config)
    applied: list[str] = []
    try:
        _quick_check(conn, str(config.database))
        conn.execute(
            "CREATE TABLE IF NOT EXISTS meta.schema_migration "
            "(version VARCHAR NOT NULL PRIMARY KEY, applied_at TIMESTAMP NOT NULL DEFAULT current_timestamp)"
        )
        present = {row[0] for row in conn.execute("SELECT version FROM meta.schema_migration").fetchall()}
        for path in _migration_files(config):
            version = path.stem
            if version in present:
                continue
            conn.execute("BEGIN IMMEDIATE")
            try:
                for statement in _migration_statements(path.read_text(encoding="utf-8")):
                    conn.execute(statement)
                conn.execute("INSERT INTO meta.schema_migration(version) VALUES (?)", [version])
                conn.execute("COMMIT")
            except Exception:
                if conn.in_transaction:
                    conn.execute("ROLLBACK")
                raise
            applied.append(version)
        conn.execute(f"PRAGMA user_version={_STORAGE_VERSION}")
        _quick_check(conn, str(config.database))
        return applied
    finally:
        if own:
            conn.close()


@contextmanager
def write_session(config: Config) -> Iterator[DatabaseConnection]:
    lock = config.database.with_suffix(config.database.suffix + ".lock")
    with ProcessLock(lock):
        _backup_if_migrations_pending(config)
        conn = connect(config)
        try:
            migrate(config, conn)
            yield conn
            conn.execute("PRAGMA optimize")
        finally:
            conn.close()


def backup_database(config: Config) -> Path:
    if not config.database.exists():
        raise FileNotFoundError(f"Database does not exist yet: {config.database}")
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    return _backup_to(config, config.backups / f"wfm_{stamp}.sqlite3")
