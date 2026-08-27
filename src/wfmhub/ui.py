"""Compact terminal dashboard for the portable WFMHub menu."""

from __future__ import annotations

import os
import sys
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import TextIO

from . import __version__
from .config import load_config
from .database import connect


ASCII_LOGO = (
    " __        ________ __  __    _   _ _   _ ____  ",
    " \\ \\      / /  ____|  \\/  |  | | | | | | |  _ \\ ",
    "  \\ \\ /\\ / /| |__  | |\\/| |  | |_| | | | | |_) |",
    "   \\ V  V / |  __| | |  | |  |  _  | |_| |  _ < ",
    "    \\_/\\_/  |_|    |_|  |_|  |_| |_|\\___/|_| \\_\\",
)
PANEL_INNER_WIDTH = 76


@dataclass(frozen=True)
class DashboardStatus:
    state: str
    database_size: int = 0
    agents: int = 0
    sources_healthy: int = 0
    sources_total: int = 0
    latest_source_date: date | str | None = None
    latest_source_family: str = ""
    quality_errors: int = 0
    quality_reviews: int = 0
    rule_version: str = "not applied"
    rule_sha256: str = ""
    last_status: str = "NEVER"
    last_refresh: datetime | str | None = None
    period_start: date | str | None = None
    period_end: date | str | None = None
    detail: str = ""


def _database_size(path: Path) -> int:
    total = path.stat().st_size if path.exists() else 0
    wal = Path(str(path) + "-wal")
    if wal.exists():
        total += wal.stat().st_size
    return total


def _format_bytes(size: int) -> str:
    value = float(max(0, size))
    for unit in ("B", "KB", "MB", "GB", "TB"):
        if value < 1024 or unit == "TB":
            return f"{value:.0f} {unit}" if unit in {"B", "KB"} else f"{value:.1f} {unit}"
        value /= 1024
    return f"{size} B"


def _format_datetime(value: datetime | str | None) -> str:
    if value is None:
        return "never"
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d %H:%M")
    return str(value).replace("T", " ")[:16]


def _format_period(start: date | str | None, end: date | str | None) -> str:
    if start is None and end is None:
        return "not built"
    return f"{start or '?'} -> {end or '?'}"


def load_dashboard_status(home: Path) -> DashboardStatus:
    """Read a tiny, failure-tolerant snapshot without changing hub data."""
    try:
        config = load_config(home)
    except Exception as exc:
        return DashboardStatus("SETUP REQUIRED", detail=str(exc))
    database_size = _database_size(config.database)
    if not config.database.exists() or database_size == 0:
        return DashboardStatus(
            "SETUP REQUIRED",
            detail="Run SETUP.cmd, then refresh the hub for the first time.",
        )

    conn = None
    try:
        conn = connect(config, read_only=True)
        last = conn.execute(
            """SELECT status, coalesce(finished_at, started_at),
                      requested_start, requested_end
               FROM meta.refresh_run ORDER BY started_at DESC LIMIT 1"""
        ).fetchone()
        agents = conn.execute("SELECT count(*) FROM core.dim_agent").fetchone()[0]
        sources_total, sources_healthy = conn.execute(
            """SELECT count(*),
                      coalesce(sum(CASE WHEN status='SUCCESS' THEN 1 ELSE 0 END), 0)
               FROM mart.source_health"""
        ).fetchone()
        latest_source = conn.execute(
            """SELECT source_family, newest_business_date
               FROM mart.source_health
               WHERE newest_business_date IS NOT NULL
               ORDER BY newest_business_date DESC, source_family
               LIMIT 1"""
        ).fetchone()
        latest_source_family = str(latest_source[0]) if latest_source else ""
        latest_source_date = latest_source[1] if latest_source else None
        quality_errors, quality_reviews = conn.execute(
            """SELECT
                   coalesce(sum(CASE WHEN severity='ERROR' THEN 1 ELSE 0 END), 0),
                   coalesce(sum(CASE WHEN severity<>'ERROR' THEN 1 ELSE 0 END), 0)
               FROM meta.quality_issue"""
        ).fetchone()
        rule = conn.execute(
            """SELECT rule_version, rule_sha256 FROM meta.rule_application
               ORDER BY applied_at DESC LIMIT 1"""
        ).fetchone()

        last_status = str(last[0]).upper() if last else "NEVER"
        last_refresh = last[1] if last else None
        period_start = last[2] if last else None
        period_end = last[3] if last else None
        if period_start is None and period_end is None:
            period_start, period_end = conn.execute(
                "SELECT min(business_date), max(business_date) FROM mart.attendance_agent_day"
            ).fetchone()

        source_problems = max(0, int(sources_total or 0) - int(sources_healthy or 0))
        if last_status == "ERROR" or source_problems or quality_errors:
            state = "CHECK DATA"
        elif last_status == "RUNNING":
            state = "WORKING"
        elif quality_reviews:
            state = "REVIEW"
        elif last_status == "SUCCESS":
            state = "READY"
        else:
            state = "READY TO REFRESH"
        return DashboardStatus(
            state=state,
            database_size=database_size,
            agents=int(agents or 0),
            sources_healthy=int(sources_healthy or 0),
            sources_total=int(sources_total or 0),
            latest_source_date=latest_source_date,
            latest_source_family=latest_source_family,
            quality_errors=int(quality_errors or 0),
            quality_reviews=int(quality_reviews or 0),
            rule_version=str(rule[0]) if rule else "not applied",
            rule_sha256=str(rule[1]) if rule else "",
            last_status=last_status,
            last_refresh=last_refresh,
            period_start=period_start,
            period_end=period_end,
        )
    except Exception as exc:
        return DashboardStatus(
            "DATABASE CHECK",
            database_size=database_size,
            detail=str(exc),
        )
    finally:
        if conn is not None:
            conn.close()


def _panel_line(value: str = "") -> str:
    cleaned = str(value).replace("\r", " ").replace("\n", " ")
    if len(cleaned) > PANEL_INNER_WIDTH - 2:
        cleaned = cleaned[: PANEL_INNER_WIDTH - 5] + "..."
    return f"| {cleaned:<{PANEL_INNER_WIDTH - 2}} |"


def dashboard_text(status: DashboardStatus) -> str:
    """Create a fixed-width ASCII dashboard suitable for standard CMD."""
    border = "+" + "-" * PANEL_INNER_WIDTH + "+"
    lines = [line.center(PANEL_INNER_WIDTH + 2) for line in ASCII_LOGO]
    lines.extend([
        "WORKFORCE MANAGEMENT CONTROL CENTER".center(PANEL_INNER_WIDTH + 2),
        "made by Anass ASSRI".center(PANEL_INNER_WIDTH + 2),
        "",
        border,
        _panel_line(
            f"{status.state}  |  v{__version__}  |  DB {_format_bytes(status.database_size)}"
            f"  |  {status.agents:,} agents"
        ),
        _panel_line(
            f"Last: {status.last_status} {_format_datetime(status.last_refresh)}"
            f"  |  Period {_format_period(status.period_start, status.period_end)}"
        ),
        _panel_line(
            f"Sources: {status.sources_healthy}/{status.sources_total} healthy"
            f"  |  Latest data: {status.latest_source_date or 'none'}"
            f"{f' ({status.latest_source_family})' if status.latest_source_family else ''}"
        ),
        _panel_line(
            f"Quality: {status.quality_errors} errors / {status.quality_reviews} review"
        ),
        _panel_line(
            f"Rules: {status.rule_version}"
            f"{f'  |  {status.rule_sha256[:12]}' if status.rule_sha256 else ''}"
        ),
    ])
    if status.detail:
        lines.append(_panel_line(status.detail))
    lines.append(border)
    return "\n".join(lines)


def render_dashboard(home: Path, stream: TextIO | None = None) -> DashboardStatus:
    stream = stream or sys.stdout
    status = load_dashboard_status(home)
    stream.write(dashboard_text(status) + "\n")
    stream.flush()
    return status


def clear_screen() -> None:
    """Clear only an interactive console; redirected output remains complete."""
    try:
        interactive = sys.stdout.isatty()
    except (AttributeError, OSError):
        interactive = False
    if interactive:
        os.system("cls" if os.name == "nt" else "clear")
