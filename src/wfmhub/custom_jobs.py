"""Trusted custom Python and read-only SQL jobs for local extensions."""

from __future__ import annotations

import csv
import runpy
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Any, Mapping, Sequence

from .config import Config
from .database import DatabaseConnection


@dataclass(frozen=True)
class QueryResult:
    headers: list[str]
    rows: list[tuple[Any, ...]]


@dataclass(frozen=True)
class JobResult:
    job: str
    output_dir: Path
    result: Any


class HubContext:
    """Small read-only API passed to custom jobs."""

    def __init__(self, conn: DatabaseConnection, config: Config, start: date, end: date, output_dir: Path):
        self._conn = conn
        self.start = start
        self.end = end
        self.output_dir = output_dir
        self.output_dir.mkdir(parents=True, exist_ok=True)

    def query(
        self,
        sql: str,
        parameters: Sequence[Any] | Mapping[str, Any] | None = None,
        max_rows: int = 100_000,
    ) -> QueryResult:
        cursor = self._conn.execute(sql, parameters)
        headers = [item[0] for item in cursor.description or []]
        rows = cursor.fetchmany(max_rows + 1)
        if len(rows) > max_rows:
            raise ValueError(f"Custom query exceeded {max_rows:,} rows; export a bounded result instead")
        return QueryResult(headers, rows)

    def write_csv(self, name: str, result: QueryResult) -> Path:
        safe = "".join(char if char.isalnum() or char in {"-", "_"} else "_" for char in name).strip("_")
        if not safe:
            raise ValueError("Custom output name is empty")
        path = self.output_dir / f"{safe}.csv"
        with path.open("w", encoding="utf-8-sig", newline="") as handle:
            writer = csv.writer(handle)
            writer.writerow(result.headers)
            writer.writerows(result.rows)
        return path


def _inside(root: Path, path: Path) -> Path:
    root = root.resolve()
    path = path.resolve()
    try:
        path.relative_to(root)
    except ValueError as exc:
        raise ValueError(f"Custom job must stay inside {root}") from exc
    return path


def list_jobs(config: Config, kind: str) -> list[Path]:
    if kind not in {"python", "sql"}:
        raise ValueError("Custom job kind must be python or sql")
    folder = config.custom / ("jobs" if kind == "python" else "sql")
    folder.mkdir(parents=True, exist_ok=True)
    pattern = "*.py" if kind == "python" else "*.sql"
    return [path for path in sorted(folder.glob(pattern)) if not path.name.startswith("_")]


def run_python_job(
    conn: DatabaseConnection,
    config: Config,
    job_path: Path,
    start: date,
    end: date,
) -> JobResult:
    root = config.custom / "jobs"
    path = _inside(root, job_path)
    if path.suffix.lower() != ".py" or not path.is_file():
        raise ValueError(f"Python job does not exist: {path}")
    output_dir = config.output / "custom" / path.stem / f"{start:%Y-%m-%d}_to_{end:%Y-%m-%d}"
    context = HubContext(conn, config, start, end, output_dir)
    namespace = runpy.run_path(str(path), run_name=f"wfmhub_custom_{path.stem}")
    runner = namespace.get("run")
    if not callable(runner):
        raise ValueError(f"{path.name} must define: def run(ctx):")
    return JobResult(path.name, output_dir, runner(context))


def run_sql_job(
    conn: DatabaseConnection,
    config: Config,
    job_path: Path,
    start: date,
    end: date,
) -> JobResult:
    root = config.custom / "sql"
    path = _inside(root, job_path)
    if path.suffix.lower() != ".sql" or not path.is_file():
        raise ValueError(f"SQL job does not exist: {path}")
    sql = path.read_text(encoding="utf-8").strip()
    statement = sql[:-1].rstrip() if sql.endswith(";") else sql
    if ";" in statement or not statement.lstrip().upper().startswith(("SELECT", "WITH")):
        raise ValueError("Custom SQL must contain exactly one read-only SELECT/WITH statement")
    output_dir = config.output / "custom" / path.stem / f"{start:%Y-%m-%d}_to_{end:%Y-%m-%d}"
    context = HubContext(conn, config, start, end, output_dir)
    result = context.query(statement, {"start": start, "end": end})
    output = context.write_csv(path.stem, result)
    return JobResult(path.name, output_dir, output)
