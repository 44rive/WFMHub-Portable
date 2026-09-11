"""Publish stable, collaboration-safe data feeds for long-lived Excel reports.

The files in ``Feed`` are replaceable data products.  The shared Excel
workbooks are not: team comments and action logs must remain under Excel /
SharePoint version control and are never rewritten by this module.
"""

from __future__ import annotations

import csv
import os
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any, Iterable, Sequence

from .config import Config
from .database import DatabaseConnection
from .metrics import load_metric_catalog


PCS_FEED_SCHEMA_VERSION = "6"
PCS_FILTER_HEADERS = (
    "Group Key", "Sort Order", "Value",
)
PCS_COACHING_HEADERS = (
    "View Key", "Rank", "Date", "LOB", "Team Leader", "Agent",
    "Agent ID", "Q1 Score", "Priority", "Call ID", "Customer Comment",
    "Coaching Key", "Data Through", "Feed Refreshed At",
)
PCS_LOB_SCORECARD_HEADERS = (
    "View Key", "Record Type", "Rank", "LOB", "Valid Q1", "PCS",
    "Prior PCS", "Change", "Participation", "Coaching Due",
    "Data Through", "Feed Refreshed At",
)
PCS_RESULTS_HEADERS = (
    "Period View", "Period Start", "Period End", "Scope Level", "LOB",
    "Team Leader", "Agent Selector", "Agent ID", "Agent", "Language",
    "PCS Average", "Participation Rate", "Valid Q1", "PCS Status 1",
    "Q1 Nonblank", "Score <= 3", "Score > 3", "Inbound Call Legs",
    "Sample State", "Data Through", "Feed Refreshed At",
)
PCS_AGENT_SCORECARD_HEADERS = (
    "View Key", "Rank", "Agent", "Agent ID", "LOB", "Team Leader", "PCS",
    "Prior PCS", "Change", "Participation", "Valid Q1", "Coaching Due",
    "Data Through", "Feed Refreshed At",
)
PCS_DAILY_SCORECARD_HEADERS = (
    "View Key", "Rank", "Date", "PCS", "Participation", "Valid Q1",
    "Coaching Due", "Data Through", "Feed Refreshed At",
)


@dataclass(frozen=True)
class SharedFeedResult:
    family: str
    files: tuple[Path, ...]
    rows: int


def _rows(
    conn: DatabaseConnection,
    sql: str,
    parameters: Sequence[Any] = (),
) -> tuple[list[str], list[tuple[Any, ...]]]:
    cursor = conn.execute(sql, list(parameters))
    return [str(item[0]) for item in cursor.description], cursor.fetchall()


def _cell(value: Any) -> Any:
    if isinstance(value, datetime):
        return value.isoformat(sep=" ", timespec="seconds")
    if isinstance(value, date):
        return value.isoformat()
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    return value


def _atomic_csv(path: Path, headers: Sequence[str], rows: Iterable[Sequence[Any]]) -> int:
    """Write a complete UTF-8 CSV, then replace the previous feed in one step."""

    path.parent.mkdir(parents=True, exist_ok=True)
    partial = path.with_name(f".{path.name}.partial")
    count = 0
    try:
        with partial.open("w", encoding="utf-8-sig", newline="") as handle:
            writer = csv.writer(handle, lineterminator="\n")
            writer.writerow(headers)
            for row in rows:
                writer.writerow([_cell(value) for value in row])
                count += 1
            handle.flush()
            os.fsync(handle.fileno())
        partial.replace(path)
    except Exception:
        partial.unlink(missing_ok=True)
        raise
    return count


def _atomic_text(path: Path, content: str) -> None:
    """Replace a small instruction asset without exposing a partial file."""

    path.parent.mkdir(parents=True, exist_ok=True)
    partial = path.with_name(f".{path.name}.partial")
    try:
        partial.write_text(content.rstrip() + "\n", encoding="utf-8")
        partial.replace(path)
    except Exception:
        partial.unlink(missing_ok=True)
        raise


def _available_period(
    conn: DatabaseConnection,
    table: str,
    fallback_start: date,
    fallback_end: date,
) -> tuple[date, date]:
    minimum, maximum = conn.execute(
        f"SELECT min(business_date), max(business_date) FROM {table}",
    ).fetchone()
    def as_date(value: Any, fallback: date) -> date:
        if value is None:
            return fallback
        if isinstance(value, datetime):
            return value.date()
        if isinstance(value, date):
            return value
        return date.fromisoformat(str(value)[:10])

    return as_date(minimum, fallback_start), as_date(maximum, fallback_end)


def _date_value(value: Any, fallback: date) -> date:
    if value is None:
        return fallback
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    return date.fromisoformat(str(value)[:10])


def _pcs_business_dates(
    conn: DatabaseConnection,
    as_of: date,
) -> tuple[date, ...]:
    rows = conn.execute(
        """SELECT DISTINCT business_date FROM mart.agent_pcs_day
           WHERE business_date<=? ORDER BY business_date""",
        (as_of,),
    ).fetchall()
    return tuple(_date_value(value, as_of) for (value,) in rows)


def _manifest(
    folder: Path,
    family: str,
    start: date,
    end: date,
    counts: Sequence[tuple[str, int]],
    *,
    schema_version: str = "1",
    extra: Sequence[tuple[Any, Any, Any]] = (),
) -> Path:
    path = folder / f"{family.upper()}_MANIFEST_CURRENT.csv"
    now = datetime.now()
    rows: list[tuple[Any, ...]] = [
        ("Feed family", family.upper(), "Business report feed"),
        ("Schema version", schema_version, "Workbook compatibility"),
        ("Data from", start, "Earliest included business date"),
        ("Data through", end, "Latest included business date"),
        ("Last refreshed", now, "Local work-machine time"),
    ]
    rows.extend(extra)
    rows.extend(("Rows", count, name) for name, count in counts)
    _atomic_csv(path, ("Item", "Value", "Details"), rows)
    return path


def _power_query_script(
    *,
    filename: str,
    headers: Sequence[str],
    types: Sequence[tuple[str, str]],
    sharepoint: bool,
) -> str:
    """Return copy-ready M for a deliberately simple PCS transport query."""

    required = ", ".join(f'"{value}"' for value in headers)
    type_rows = ",\n            ".join(f'{{"{name}", {kind}}}' for name, kind in types)
    if sharepoint:
        source = f'''    SiteUrl = Setting("SharePoint Site URL"),
    FeedFolder = Text.Lower(Text.Replace(Setting("SharePoint Feed Folder"), "\\", "/")),
    Files = SharePoint.Files(SiteUrl, [ApiVersion = 15]),
    Matches = Table.SelectRows(Files, each [Name] = "{filename}" and Text.Contains(Text.Lower(Text.Replace([Folder Path], "\\", "/")), FeedFolder)),
    Checked = if Table.RowCount(Matches) = 1 then Matches{{0}}[Content] else error Error.Record("PCS feed", "Expected exactly one {filename} in the configured SharePoint folder", [Matches = Table.RowCount(Matches)]),
    Csv = Csv.Document(Checked, [Delimiter = ",", Encoding = 65001, QuoteStyle = QuoteStyle.Csv]),'''
    else:
        source = f'''    FeedFolder = Setting("Local Feed Folder"),
    FilePath = FeedFolder & (if Text.EndsWith(FeedFolder, "\\") then "" else "\\") & "{filename}",
    Csv = Csv.Document(File.Contents(FilePath), [Delimiter = ",", Encoding = 65001, QuoteStyle = QuoteStyle.Csv]),'''
    return f'''// WFMHub PCS schema {PCS_FEED_SCHEMA_VERSION}
// Excel: Data > Get Data > Blank Query > Advanced Editor. Replace everything with this script.
let
    Setup = Excel.CurrentWorkbook(){{[Name="tblSetup"]}}[Content],
    Setting = (Name as text) as text =>
        let
            Matches = Table.SelectRows(Setup, each Text.From([Setting]) = Name),
            Result = if Table.RowCount(Matches) = 1 then Text.Trim(Text.From(Matches{{0}}[Value])) else error Error.Record("PCS setup", "Missing or duplicate SETUP row", [Setting = Name])
        in
            Result,
{source}
    Promoted = Table.PromoteHeaders(Csv, [PromoteAllScalars = true]),
    Required = {{{required}}},
    Missing = List.Difference(Required, Table.ColumnNames(Promoted)),
    CheckedColumns = if List.IsEmpty(Missing) then Promoted else error Error.Record("PCS feed schema", "Missing required columns", [Missing = Text.Combine(Missing, ", ")]),
    Typed = Table.TransformColumnTypes(CheckedColumns, {{
            {type_rows}
        }}, "en-US")
in
    Typed'''


def _pcs_reporting_periods(
    as_of: date,
    available_start: date | None = None,
    available_months: Sequence[date] | None = None,
) -> tuple[tuple[str, date, date], ...]:
    """Return operational presets plus every available PCS calendar month."""

    return tuple(
        (window.label, window.start, window.end)
        for window in _pcs_reporting_windows(
            as_of,
            available_start=available_start,
            available_months=available_months,
        )
    )


def _pcs_scope_aggregates(
    conn: DatabaseConnection,
    start: date,
    end: date,
    level: str,
) -> list[dict[str, Any]]:
    """Aggregate additive PCS counters at one explicit presentation grain."""

    dimensions = {
        "ALL": (
            "'ALL', '', '', '', '', ''",
            "",
            "",
        ),
        "LOB": (
            "coalesce(lob,'Unassigned'), '', '', '', '', ''",
            "GROUP BY coalesce(lob,'Unassigned')",
            "ORDER BY 1",
        ),
        "TEAM": (
            "coalesce(lob,'Unassigned'), coalesce(team_leader,'Unassigned'), '', '', '', ''",
            "GROUP BY coalesce(lob,'Unassigned'), coalesce(team_leader,'Unassigned')",
            "ORDER BY 1, 2",
        ),
        "AGENT": (
            "coalesce(lob,'Unassigned'), coalesce(team_leader,'Unassigned'), "
            "coalesce(agent_name,'Agent') || ' [' || agent_id || ']', agent_id, "
            "coalesce(agent_name,'Agent'), coalesce(language,'')",
            "GROUP BY coalesce(lob,'Unassigned'), coalesce(team_leader,'Unassigned'), "
            "agent_id, coalesce(agent_name,'Agent'), coalesce(language,'')",
            "ORDER BY 1, 2, 5, 4",
        ),
    }
    try:
        select_dimensions, group_by, order_by = dimensions[level]
    except KeyError as exc:
        raise ValueError(f"Unsupported PCS result level: {level}") from exc
    cursor = conn.execute(
        f"""SELECT {select_dimensions},
                   coalesce(sum(pcs_score_sum),0),
                   coalesce(sum(survey_responses),0),
                   coalesce(sum(pcs_participation_responses),0),
                   coalesce(sum(pcs_status_calls),0),
                   coalesce(sum(low_score_responses),0),
                   coalesce(sum(top_box_responses),0),
                   coalesce(sum(inbound_calls),0)
            FROM mart.agent_pcs_day
            WHERE business_date BETWEEN ? AND ?
            {group_by} {order_by}""",
        (start, end),
    )
    output = []
    for row in cursor.fetchall():
        lob, team, selector, agent_id, agent, language = row[:6]
        score_sum, valid, q1_nonblank, eligible, low, positive, inbound = row[6:]
        # SQLite returns one all-zero aggregate row for an empty ungrouped set.
        if level != "ALL" and not any((valid, eligible, inbound)):
            continue
        output.append({
            "lob": lob,
            "team_leader": team,
            "agent_selector": selector,
            "agent_id": agent_id,
            "agent": agent,
            "language": language,
            "pcs_average": float(score_sum) / float(valid) if valid else None,
            "participation_rate": float(q1_nonblank) / float(eligible) if eligible else None,
            "valid_q1": valid,
            "pcs_status_1": eligible,
            "q1_nonblank": q1_nonblank,
            "low_scores": low,
            "positive_scores": positive,
            "inbound_legs": inbound,
        })
    return output


@dataclass(frozen=True)
class PCSPeriodWindow:
    label: str
    start: date
    end: date
    prior_start: date
    prior_end: date
    trend_start: date | None = None


def _month_before(value: date) -> tuple[date, date]:
    end = value.replace(day=1) - timedelta(days=1)
    return end.replace(day=1), end


def _pcs_reporting_windows(
    as_of: date,
    previous_available: date | None = None,
    available_start: date | None = None,
    available_months: Sequence[date] | None = None,
) -> tuple[PCSPeriodWindow, ...]:
    """Return operational presets and complete available-month coverage.

    Dashboard cards and tables use the complete window.  The all-history trend
    is intentionally limited to its latest 31 calendar days so the permanent
    tracker remains lightweight; PERFORMANCE carries every agent/day row.
    """

    month_start = as_of.replace(day=1)
    previous_start, previous_end = _month_before(as_of)
    two_months_start, two_months_end = _month_before(previous_start)
    current_days = (as_of - month_start).days
    previous_mtd_end = min(previous_end, previous_start + timedelta(days=current_days))
    previous_mtd_days = (previous_mtd_end - previous_start).days
    two_months_mtd_end = min(
        two_months_end,
        two_months_start + timedelta(days=previous_mtd_days),
    )
    week_start = as_of - timedelta(days=as_of.weekday())
    latest_prior = previous_available or (as_of - timedelta(days=1))
    operational = (
        PCSPeriodWindow("Latest day", as_of, as_of, latest_prior, latest_prior),
        PCSPeriodWindow(
            "Current week", week_start, as_of,
            week_start - timedelta(days=7), as_of - timedelta(days=7),
        ),
        PCSPeriodWindow(
            "Current MTD", month_start, as_of,
            previous_start, previous_mtd_end,
        ),
        PCSPeriodWindow(
            "Previous MTD same days", previous_start, previous_mtd_end,
            two_months_start, two_months_mtd_end,
        ),
        PCSPeriodWindow(
            "Previous full month", previous_start, previous_end,
            two_months_start, two_months_end,
        ),
    )
    if available_start is None:
        return operational

    first_available = min(available_start, as_of)
    span_days = (as_of - first_available).days
    prior_all_end = first_available - timedelta(days=1)
    all_available = PCSPeriodWindow(
        "All available",
        first_available,
        as_of,
        prior_all_end - timedelta(days=span_days),
        prior_all_end,
        max(first_available, as_of - timedelta(days=30)),
    )

    first_month = first_available.replace(day=1)
    if available_months is None:
        month_values: list[date] = []
        month = month_start
        while month >= first_month:
            month_values.append(month)
            month = _month_before(month)[0]
    else:
        month_values = sorted({
            value.replace(day=1)
            for value in available_months
            if first_month <= value.replace(day=1) <= month_start
        }, reverse=True)

    months: list[PCSPeriodWindow] = []
    for month in month_values:
        next_month = (
            date(month.year + 1, 1, 1)
            if month.month == 12
            else date(month.year, month.month + 1, 1)
        )
        month_end = next_month - timedelta(days=1)
        window_start = max(month, first_available)
        window_end = min(month_end, as_of)
        prior_month_start, prior_month_end = _month_before(month)
        start_offset = (window_start - month).days
        end_offset = (window_end - month).days
        comparable_start = min(
            prior_month_end,
            prior_month_start + timedelta(days=start_offset),
        )
        comparable_end = min(
            prior_month_end,
            prior_month_start + timedelta(days=end_offset),
        )
        months.append(PCSPeriodWindow(
            f"Month {month:%Y-%m}",
            window_start,
            window_end,
            comparable_start,
            comparable_end,
        ))

    return (*operational, all_available, *months)


def _key_part(value: Any) -> str:
    return str(value or "Unassigned").strip().replace("|", "/") or "Unassigned"


def _selection_key(period: str, lob: str, leader: str, agent: str) -> str:
    return "|".join(_key_part(value) for value in (period, lob, leader, agent))


def _pcs_agent_days(
    conn: DatabaseConnection,
    start: date,
    end: date,
) -> list[dict[str, Any]]:
    cursor = conn.execute(
        """SELECT business_date, coalesce(lob,'Unassigned'),
                  coalesce(team_leader,'Unassigned'), agent_id,
                  coalesce(agent_name,'Agent'),
                  coalesce(agent_name,'Agent') || ' [' || agent_id || ']',
                  coalesce(pcs_score_sum,0), coalesce(survey_responses,0),
                  coalesce(pcs_participation_responses,0),
                  coalesce(pcs_status_calls,0),
                  coalesce(low_score_responses,0),
                  coalesce(top_box_responses,0), coalesce(inbound_calls,0)
           FROM mart.agent_pcs_day
           WHERE business_date BETWEEN ? AND ?
             AND trim(coalesce(agent_id,''))<>''
           ORDER BY business_date, lob, team_leader, agent_name, agent_id""",
        (start, end),
    )
    output: list[dict[str, Any]] = []
    for row in cursor.fetchall():
        business_date = row[0]
        if isinstance(business_date, datetime):
            business_date = business_date.date()
        elif not isinstance(business_date, date):
            business_date = date.fromisoformat(str(business_date)[:10])
        output.append({
            "date": business_date,
            "lob": _key_part(row[1]),
            "leader": _key_part(row[2]),
            "agent_id": str(row[3] or "").strip(),
            "agent": str(row[4] or "Agent").strip() or "Agent",
            "selector": _key_part(row[5]),
            "score_sum": float(row[6] or 0),
            "valid": int(row[7] or 0),
            "nonblank": int(row[8] or 0),
            "eligible": int(row[9] or 0),
            "low": int(row[10] or 0),
            "positive": int(row[11] or 0),
            "inbound": int(row[12] or 0),
        })
    return output


def _pcs_dimensions(
    rows: Sequence[dict[str, Any]],
) -> tuple[
    tuple[str, ...], dict[str, tuple[str, ...]],
    dict[tuple[str, str], tuple[str, ...]],
]:
    lobs = tuple(sorted({str(row["lob"]) for row in rows}, key=str.casefold))
    leaders: dict[str, tuple[str, ...]] = {
        "All": tuple(sorted({str(row["leader"]) for row in rows}, key=str.casefold)),
    }
    agents: dict[tuple[str, str], tuple[str, ...]] = {}
    for lob in lobs:
        leaders[lob] = tuple(sorted(
            {str(row["leader"]) for row in rows if row["lob"] == lob},
            key=str.casefold,
        ))
    for lob in ("All", *lobs):
        scoped = list(rows) if lob == "All" else [row for row in rows if row["lob"] == lob]
        agents[(lob, "All")] = tuple(sorted(
            {str(row["selector"]) for row in scoped}, key=str.casefold,
        ))
        for leader in leaders[lob]:
            agents[(lob, leader)] = tuple(sorted(
                {str(row["selector"]) for row in scoped if row["leader"] == leader},
                key=str.casefold,
            ))
    return lobs, leaders, agents


def pcs_filter_rows(
    rows: Sequence[dict[str, Any]],
    windows: Sequence[PCSPeriodWindow],
) -> list[tuple[Any, ...]]:
    """Build contiguous governed dropdown lists for the permanent tracker."""

    lobs, leaders, agents = _pcs_dimensions(rows)
    groups: list[tuple[str, Sequence[str]]] = [
        ("PERIOD", tuple(window.label for window in windows)),
        ("LOB", ("All", *lobs)),
    ]
    for lob in ("All", *lobs):
        groups.append((f"TL|{_key_part(lob)}", ("All", *leaders[lob])))
        for leader in ("All", *leaders[lob]):
            groups.append((
                f"AGENT|{_key_part(lob)}|{_key_part(leader)}",
                ("All", *agents[(lob, leader)]),
            ))
    return [
        (group, rank, value)
        for group, values in groups
        for rank, value in enumerate(values, 1)
    ]


def pcs_dashboard_cache_rows(
    conn: DatabaseConnection,
    as_of: date,
    refreshed: datetime,
) -> tuple[
    list[tuple[Any, ...]], list[tuple[Any, ...]],
    list[tuple[Any, ...]], list[tuple[Any, ...]],
]:
    """Precompute every valid Overview selection; Excel only performs lookups."""

    available_dates = _pcs_business_dates(conn, as_of)
    available_start = available_dates[0] if available_dates else as_of
    previous = available_dates[-2] if len(available_dates) > 1 else None
    windows = _pcs_reporting_windows(
        as_of,
        previous,
        available_start=available_start,
        available_months=available_dates,
    )
    earliest = min(window.prior_start for window in windows)
    rows = _pcs_agent_days(conn, earliest, as_of)
    lobs, leaders, agents = _pcs_dimensions(rows)
    filter_rows = pcs_filter_rows(rows, windows)
    lob_cache: list[tuple[Any, ...]] = []
    agent_cache: list[tuple[Any, ...]] = []
    daily_cache: list[tuple[Any, ...]] = []

    selections = [
        (lob, leader, agent)
        for lob in ("All", *lobs)
        for leader in ("All", *leaders[lob])
        for agent in ("All", *agents[(lob, leader)])
    ]

    def scopes(row: dict[str, Any]) -> set[tuple[str, str, str]]:
        lob = str(row["lob"])
        leader = str(row["leader"])
        agent = str(row["selector"])
        return {
            ("All", "All", "All"), ("All", "All", agent),
            ("All", leader, "All"), ("All", leader, agent),
            (lob, "All", "All"), (lob, "All", agent),
            (lob, leader, "All"), (lob, leader, agent),
        }

    def add(target: dict[Any, list[float]], key: Any, row: dict[str, Any]) -> None:
        values = target.setdefault(key, [0.0, 0.0, 0.0, 0.0, 0.0])
        values[0] += float(row["score_sum"])
        values[1] += int(row["valid"])
        values[2] += int(row["nonblank"])
        values[3] += int(row["eligible"])
        values[4] += int(row["low"])

    def metric(values: Sequence[float] | None) -> dict[str, Any]:
        score_sum, valid, nonblank, eligible, low = values or (0, 0, 0, 0, 0)
        return {
            "pcs": score_sum / valid if valid else None,
            "participation": nonblank / eligible if eligible else None,
            "valid": int(valid), "low": int(low),
        }

    for window in windows:
        current_totals: dict[Any, list[float]] = {}
        prior_totals: dict[Any, list[float]] = {}
        current_lobs: dict[Any, list[float]] = {}
        prior_lobs: dict[Any, list[float]] = {}
        current_agents: dict[Any, list[float]] = {}
        prior_agents: dict[Any, list[float]] = {}
        current_days: dict[Any, list[float]] = {}
        agent_sources: dict[Any, dict[str, Any]] = {}
        scope_lobs: dict[Any, set[str]] = {}
        scope_agents: dict[Any, set[str]] = {}
        for row in rows:
            is_current = window.start <= row["date"] <= window.end
            is_prior = window.prior_start <= row["date"] <= window.prior_end
            if not is_current and not is_prior:
                continue
            for scope in scopes(row):
                if is_current:
                    add(current_totals, scope, row)
                    add(current_lobs, (scope, row["lob"]), row)
                    add(current_agents, (scope, row["agent_id"]), row)
                    add(current_days, (scope, row["date"]), row)
                    agent_sources.setdefault((scope, row["agent_id"]), row)
                    scope_lobs.setdefault(scope, set()).add(str(row["lob"]))
                    scope_agents.setdefault(scope, set()).add(str(row["agent_id"]))
                if is_prior:
                    add(prior_totals, scope, row)
                    add(prior_lobs, (scope, row["lob"]), row)
                    add(prior_agents, (scope, row["agent_id"]), row)

        for scope in selections:
            lob, leader, agent = scope
            selection = _selection_key(window.label, lob, leader, agent)
            current = metric(current_totals.get(scope))
            prior = metric(prior_totals.get(scope))
            change = (
                current["pcs"] - prior["pcs"]
                if current["pcs"] is not None and prior["pcs"] is not None
                else None
            )
            lob_cache.append((
                f"KPI|{selection}", "KPI", 0, lob,
                current["valid"], current["pcs"], prior["pcs"], change,
                current["participation"], current["low"], as_of, refreshed,
            ))
            visible_lobs = sorted(scope_lobs.get(scope, set()), key=str.casefold)
            for rank, visible_lob in enumerate(visible_lobs, 1):
                current_lob = metric(current_lobs.get((scope, visible_lob)))
                prior_lob = metric(prior_lobs.get((scope, visible_lob)))
                lob_change = (
                    current_lob["pcs"] - prior_lob["pcs"]
                    if current_lob["pcs"] is not None and prior_lob["pcs"] is not None
                    else None
                )
                lob_cache.append((
                    f"LOB|{selection}|{rank}", "LOB", rank, visible_lob,
                    current_lob["valid"], current_lob["pcs"],
                    prior_lob["pcs"], lob_change,
                    current_lob["participation"], current_lob["low"],
                    as_of, refreshed,
                ))
            agent_ids = sorted(scope_agents.get(scope, set()), key=lambda agent_id: str(
                agent_sources[(scope, agent_id)]["agent"]
            ).casefold())
            for rank, agent_id in enumerate(agent_ids, 1):
                source = agent_sources[(scope, agent_id)]
                current_agent = metric(current_agents.get((scope, agent_id)))
                prior_agent = metric(prior_agents.get((scope, agent_id)))
                agent_change = (
                    current_agent["pcs"] - prior_agent["pcs"]
                    if current_agent["pcs"] is not None and prior_agent["pcs"] is not None
                    else None
                )
                agent_cache.append((
                    f"AGENT|{selection}|{rank}", rank, source["agent"],
                    agent_id, source["lob"], source["leader"],
                    current_agent["pcs"], prior_agent["pcs"], agent_change,
                    current_agent["participation"], current_agent["valid"],
                    current_agent["low"], as_of, refreshed,
                ))
            trend_start = window.trend_start or window.start
            for rank, offset in enumerate(
                range((window.end - trend_start).days + 1), 1,
            ):
                business_date = trend_start + timedelta(days=offset)
                daily = metric(current_days.get((scope, business_date)))
                daily_cache.append((
                    f"DAILY|{selection}|{rank}", rank, business_date,
                    daily["pcs"], daily["participation"], daily["valid"],
                    daily["low"], as_of, refreshed,
                ))
    return filter_rows, lob_cache, agent_cache, daily_cache


def pcs_result_rows(
    conn: DatabaseConnection,
    as_of: date,
    minimum_sample: int,
    refreshed: datetime,
) -> list[tuple[Any, ...]]:
    """Build filter-ready LOB, team, and agent scorecards for standard periods."""

    available_dates = _pcs_business_dates(conn, as_of)
    available_start = available_dates[0] if available_dates else as_of

    rows: list[tuple[Any, ...]] = []
    for period_label, period_start, period_end in _pcs_reporting_periods(
        as_of,
        available_start,
        available_dates,
    ):
        for level in ("LOB", "TEAM", "AGENT"):
            for item in _pcs_scope_aggregates(conn, period_start, period_end, level):
                valid = int(item["valid_q1"] or 0)
                rows.append((
                    period_label, period_start, period_end, level,
                    item["lob"], item["team_leader"], item["agent_selector"],
                    item["agent_id"], item["agent"], item["language"],
                    item["pcs_average"], item["participation_rate"], valid,
                    int(item["pcs_status_1"] or 0), int(item["q1_nonblank"] or 0),
                    int(item["low_scores"] or 0), int(item["positive_scores"] or 0),
                    int(item["inbound_legs"] or 0),
                    "LOW SAMPLE" if valid < minimum_sample else "OK",
                    as_of, refreshed,
                ))

    # PERFORMANCE is the detailed, native-filter surface.  These rows make
    # every available business date inspectable without exposing raw call legs
    # or adding a new query/table contract.
    cursor = conn.execute(
        """SELECT business_date, coalesce(lob,'Unassigned'),
                  coalesce(team_leader,'Unassigned'),
                  coalesce(agent_name,'Agent') || ' [' || agent_id || ']',
                  agent_id, coalesce(agent_name,'Agent'), coalesce(language,''),
                  coalesce(pcs_score_sum,0), coalesce(survey_responses,0),
                  coalesce(pcs_participation_responses,0),
                  coalesce(pcs_status_calls,0),
                  coalesce(low_score_responses,0),
                  coalesce(top_box_responses,0), coalesce(inbound_calls,0)
           FROM mart.agent_pcs_day
           WHERE business_date BETWEEN ? AND ?
             AND trim(coalesce(agent_id,''))<>''
           ORDER BY business_date DESC, lob, team_leader, agent_name, agent_id""",
        (available_start, as_of),
    )
    for item in cursor.fetchall():
        business_date = item[0]
        if isinstance(business_date, datetime):
            business_date = business_date.date()
        elif not isinstance(business_date, date):
            business_date = date.fromisoformat(str(business_date)[:10])
        (
            lob, team, selector, agent_id, agent, language,
            score_sum, valid, q1_nonblank, eligible, low, positive, inbound,
        ) = item[1:]
        valid = int(valid or 0)
        eligible = int(eligible or 0)
        rows.append((
            "Daily detail", business_date, business_date, "AGENT DAY",
            lob, team, selector, agent_id, agent, language,
            float(score_sum) / valid if valid else None,
            float(q1_nonblank) / eligible if eligible else None,
            valid, eligible, int(q1_nonblank or 0), int(low or 0),
            int(positive or 0), int(inbound or 0),
            "LOW SAMPLE" if valid < minimum_sample else "OK",
            as_of, refreshed,
        ))
    return rows


def pcs_coaching_cache_rows(
    conn: DatabaseConnection,
    config: Config,
    as_of: date,
    refreshed: datetime,
) -> list[tuple[Any, ...]]:
    """Build period/LOB lookup rows for the visible coaching work queue."""

    available_dates = _pcs_business_dates(conn, as_of)
    available_start = available_dates[0] if available_dates else as_of
    windows = _pcs_reporting_windows(
        as_of,
        available_dates[-2] if len(available_dates) > 1 else None,
        available_start=available_start,
        available_months=available_dates,
    )
    earliest = min(window.prior_start for window in windows)
    primary_score = f"question_{config.pcs.primary_score_question}_score"
    allowed_scores = ", ".join(f"{value:g}" for value in config.pcs.allowed_scores)
    cursor = conn.execute(
        f"""SELECT c.business_date, coalesce(d.lob,c.lob,'Unassigned'),
                   coalesce(d.team_leader,'Unassigned'),
                   coalesce(d.canonical_name,c.agent_name,'Agent'), c.agent_id,
                   c.{primary_score},
                   CASE WHEN c.{primary_score}<=2 THEN 'High' ELSE 'Normal' END,
                   c.call_id, c.question_3, c.call_key
            FROM core.clean_call_leg c
            LEFT JOIN core.dim_agent d ON d.agent_id=c.agent_id
            WHERE c.business_date BETWEEN ? AND ?
              AND upper(coalesce(c.call_direction,''))='I'
              AND c.{primary_score} IN ({allowed_scores})
              AND c.{primary_score}<=?
            ORDER BY c.business_date DESC, d.team_leader,
                     coalesce(d.canonical_name,c.agent_name), c.call_start""",
        (earliest, as_of, config.pcs.negative_score_maximum),
    )
    source: list[tuple[Any, ...]] = []
    for raw in cursor.fetchall():
        business_date = raw[0]
        if isinstance(business_date, datetime):
            business_date = business_date.date()
        elif not isinstance(business_date, date):
            business_date = date.fromisoformat(str(business_date)[:10])
        source.append((
            business_date, _key_part(raw[1]), _key_part(raw[2]),
            str(raw[3] or "Agent"), str(raw[4] or ""), raw[5], raw[6],
            str(raw[7] or ""), str(raw[8] or ""), str(raw[9] or ""),
        ))
    lobs = sorted({str(row[1]) for row in source}, key=str.casefold)
    output: list[tuple[Any, ...]] = []
    for window in windows:
        for lob in ("All", *lobs):
            matches = [
                row for row in source
                if window.start <= row[0] <= window.end
                and (lob == "All" or row[1] == lob)
            ]
            for rank, row in enumerate(matches, 1):
                output.append((
                    f"COACH|{_key_part(window.label)}|{_key_part(lob)}|{rank}",
                    rank, *row, as_of, refreshed,
                ))
    return output


def _publish_pcs_power_query_scripts(folder: Path) -> tuple[Path, ...]:
    filter_types = (
        ("Group Key", "type text"), ("Sort Order", "Int64.Type"),
        ("Value", "type text"),
    )
    queue_types = (
        ("View Key", "type text"), ("Rank", "Int64.Type"),
        ("Date", "type date"), ("LOB", "type text"),
        ("Team Leader", "type text"), ("Agent", "type text"),
        ("Agent ID", "type text"),
        ("Q1 Score", "type number"), ("Customer Comment", "type text"),
        ("Priority", "type text"), ("Call ID", "type text"),
        ("Coaching Key", "type text"), ("Data Through", "type date"),
        ("Feed Refreshed At", "type datetime"),
    )
    lob_types = (
        ("View Key", "type text"), ("Record Type", "type text"),
        ("Rank", "Int64.Type"), ("LOB", "type text"),
        ("Valid Q1", "Int64.Type"), ("PCS", "type number"),
        ("Prior PCS", "type number"), ("Change", "type number"),
        ("Participation", "type number"), ("Coaching Due", "Int64.Type"),
        ("Data Through", "type date"),
        ("Feed Refreshed At", "type datetime"),
    )
    result_types = (
        ("Period View", "type text"), ("Period Start", "type date"),
        ("Period End", "type date"), ("Scope Level", "type text"),
        ("LOB", "type text"), ("Team Leader", "type text"),
        ("Agent Selector", "type text"), ("Agent ID", "type text"),
        ("Agent", "type text"), ("Language", "type text"),
        ("PCS Average", "type number"), ("Participation Rate", "type number"),
        ("Valid Q1", "Int64.Type"), ("PCS Status 1", "Int64.Type"),
        ("Q1 Nonblank", "Int64.Type"), ("Score <= 3", "Int64.Type"),
        ("Score > 3", "Int64.Type"), ("Inbound Call Legs", "Int64.Type"),
        ("Sample State", "type text"), ("Data Through", "type date"),
        ("Feed Refreshed At", "type datetime"),
    )
    agent_types = (
        ("View Key", "type text"), ("Rank", "Int64.Type"),
        ("Agent", "type text"), ("Agent ID", "type text"),
        ("LOB", "type text"), ("Team Leader", "type text"),
        ("PCS", "type number"), ("Prior PCS", "type number"),
        ("Change", "type number"), ("Participation", "type number"),
        ("Valid Q1", "Int64.Type"), ("Coaching Due", "Int64.Type"),
        ("Data Through", "type date"),
        ("Feed Refreshed At", "type datetime"),
    )
    daily_types = (
        ("View Key", "type text"), ("Rank", "Int64.Type"),
        ("Date", "type date"), ("PCS", "type number"),
        ("Participation", "type number"), ("Valid Q1", "Int64.Type"),
        ("Coaching Due", "Int64.Type"), ("Data Through", "type date"),
        ("Feed Refreshed At", "type datetime"),
    )
    specifications = (
        ("POWER_QUERY_PCS_FILTERS_SHAREPOINT.txt", "PCS_FILTER_LIST_CURRENT.csv", PCS_FILTER_HEADERS, filter_types, True),
        ("POWER_QUERY_COACHING_QUEUE_SHAREPOINT.txt", "PCS_COACHING_OPPORTUNITY_CURRENT.csv", PCS_COACHING_HEADERS, queue_types, True),
        ("POWER_QUERY_PCS_LOB_SHAREPOINT.txt", "PCS_LOB_SCORECARD_CURRENT.csv", PCS_LOB_SCORECARD_HEADERS, lob_types, True),
        ("POWER_QUERY_PCS_AGENT_SHAREPOINT.txt", "PCS_AGENT_SCORECARD_CURRENT.csv", PCS_AGENT_SCORECARD_HEADERS, agent_types, True),
        ("POWER_QUERY_PCS_DAILY_SHAREPOINT.txt", "PCS_DAILY_SCORECARD_CURRENT.csv", PCS_DAILY_SCORECARD_HEADERS, daily_types, True),
        ("POWER_QUERY_PCS_RESULTS_SHAREPOINT.txt", "PCS_RESULTS_CURRENT.csv", PCS_RESULTS_HEADERS, result_types, True),
        ("POWER_QUERY_PCS_FILTERS_LOCAL.txt", "PCS_FILTER_LIST_CURRENT.csv", PCS_FILTER_HEADERS, filter_types, False),
        ("POWER_QUERY_COACHING_QUEUE_LOCAL.txt", "PCS_COACHING_OPPORTUNITY_CURRENT.csv", PCS_COACHING_HEADERS, queue_types, False),
        ("POWER_QUERY_PCS_LOB_LOCAL.txt", "PCS_LOB_SCORECARD_CURRENT.csv", PCS_LOB_SCORECARD_HEADERS, lob_types, False),
        ("POWER_QUERY_PCS_AGENT_LOCAL.txt", "PCS_AGENT_SCORECARD_CURRENT.csv", PCS_AGENT_SCORECARD_HEADERS, agent_types, False),
        ("POWER_QUERY_PCS_DAILY_LOCAL.txt", "PCS_DAILY_SCORECARD_CURRENT.csv", PCS_DAILY_SCORECARD_HEADERS, daily_types, False),
        ("POWER_QUERY_PCS_RESULTS_LOCAL.txt", "PCS_RESULTS_CURRENT.csv", PCS_RESULTS_HEADERS, result_types, False),
    )
    paths = []
    for script_name, filename, headers, types, sharepoint in specifications:
        path = folder / script_name
        _atomic_text(path, _power_query_script(
            filename=filename, headers=headers, types=types, sharepoint=sharepoint,
        ))
        paths.append(path)
    return tuple(paths)


def publish_pcs_feeds(
    conn: DatabaseConnection,
    config: Config,
    fallback_start: date,
    fallback_end: date,
) -> SharedFeedResult:
    """Publish lightweight, direct-CSV PCS report and coaching feeds."""

    start, end = _available_period(
        conn, "mart.agent_pcs_day", fallback_start, fallback_end,
    )
    folder = config.feed / "PCS"
    files: list[Path] = []
    counts: list[tuple[str, int]] = []

    metric_catalog = load_metric_catalog(config.home, config.metric_catalog)
    method = metric_catalog.method_for("pcs_average", end, {})
    minimum_sample = int(method.minimum_sample) if method is not None else 1
    refreshed = datetime.now()

    filter_rows, lob_rows, agent_rows, daily_rows = pcs_dashboard_cache_rows(
        conn, end, refreshed,
    )
    path = folder / "PCS_FILTER_LIST_CURRENT.csv"
    counts.append((path.name, _atomic_csv(path, PCS_FILTER_HEADERS, filter_rows)))
    files.append(path)

    path = folder / "PCS_LOB_SCORECARD_CURRENT.csv"
    counts.append((path.name, _atomic_csv(path, PCS_LOB_SCORECARD_HEADERS, lob_rows)))
    files.append(path)

    path = folder / "PCS_AGENT_SCORECARD_CURRENT.csv"
    counts.append((path.name, _atomic_csv(path, PCS_AGENT_SCORECARD_HEADERS, agent_rows)))
    files.append(path)

    path = folder / "PCS_DAILY_SCORECARD_CURRENT.csv"
    counts.append((path.name, _atomic_csv(path, PCS_DAILY_SCORECARD_HEADERS, daily_rows)))
    files.append(path)

    result_rows = pcs_result_rows(conn, end, minimum_sample, refreshed)
    path = folder / "PCS_RESULTS_CURRENT.csv"
    counts.append((path.name, _atomic_csv(path, PCS_RESULTS_HEADERS, result_rows)))
    files.append(path)

    rows = pcs_coaching_cache_rows(conn, config, end, refreshed)
    path = folder / "PCS_COACHING_OPPORTUNITY_CURRENT.csv"
    counts.append((path.name, _atomic_csv(path, PCS_COACHING_HEADERS, rows)))
    files.append(path)
    files.extend(_publish_pcs_power_query_scripts(folder))
    files.append(_manifest(
        folder, "PCS", start, end, counts,
        schema_version=PCS_FEED_SCHEMA_VERSION,
        extra=(
            ("Metric catalog version", metric_catalog.version, metric_catalog.sha256),
            ("Workbook load", "Lightweight tables only", "No raw PCS data worksheet"),
        ),
    ))
    return SharedFeedResult("PCS", tuple(files), sum(count for _, count in counts))


def publish_absence_feeds(
    conn: DatabaseConnection,
    config: Config,
    fallback_start: date,
    fallback_end: date,
) -> SharedFeedResult:
    """Publish Activities-final absence, component and completeness feeds."""

    start, end = _available_period(
        conn, "mart.verint_final_absence_agent_day", fallback_start, fallback_end,
    )
    folder = config.feed / "Absenteeism"
    files: list[Path] = []
    counts: list[tuple[str, int]] = []

    _query_headers, rows = _rows(
        conn,
        """SELECT business_date, lob, team_leader,
                  coalesce(agent_name,'Agent') || ' [' || agent_id || ']',
                  agent_id, agent_name, ops_manager, language, location,
                  scheduled_minutes/60.0, planned_net_minutes/60.0,
                  final_absence_minutes/60.0, final_vacation_minutes/60.0,
                  final_unpaid_minutes/60.0, final_shrinkage_minutes/60.0,
                  final_unmapped_minutes/60.0, final_absence_rate,
                  final_absence_day, final_ledger_status, agent_day_key
           FROM mart.verint_final_absence_agent_day
           WHERE business_date BETWEEN ? AND ?
           ORDER BY business_date, lob, team_leader, agent_name, agent_id""",
        (start, end),
    )
    headers = [
        "Date", "LOB", "Team Leader", "Agent Selector", "Agent ID", "Agent",
        "Ops Manager", "Language", "Location", "Scheduled Hours",
        "Planned Net Hours", "Absence Hours", "Vacation Hours", "Unpaid Hours",
        "Shrinkage Hours", "Unmapped Hours", "Absence Rate %",
        "Final Absence Day", "Final Ledger Status", "Case ID",
    ]
    path = folder / "ABSENCE_AGENT_DAY_CURRENT.csv"
    counts.append((path.name, _atomic_csv(path, headers, rows)))
    files.append(path)

    headers, rows = _rows(
        conn,
        """SELECT event_key AS Event_ID, agent_day_key AS Case_ID,
                  business_date AS Date, agent_id AS Agent_ID,
                  agent_name AS Agent_Name,
                  coalesce(agent_name,'Agent') || ' [' || agent_id || ']' AS Agent_Selector,
                  team_leader AS Team_Leader, ops_manager AS Operations_Manager,
                  lob AS LOB, language AS Language, activity AS Activity,
                  category AS Component, event_start AS Start_Time,
                  event_end AS End_Time, minutes AS Minutes, hours AS Hours,
                  counts_as_absence AS Counts_As_Absence,
                  counts_as_vacation AS Counts_As_Vacation,
                  counts_as_unpaid AS Counts_As_Unpaid,
                  counts_as_shrinkage AS Counts_As_Shrinkage,
                  mapped AS Mapped
           FROM mart.verint_final_absence_event
           WHERE business_date BETWEEN ? AND ?
           ORDER BY business_date, lob, team_leader, agent_name, event_start""",
        (start, end),
    )
    path = folder / "ABSENCE_COMPONENT_CURRENT.csv"
    counts.append((path.name, _atomic_csv(path, headers, rows)))
    files.append(path)

    headers, rows = _rows(
        conn,
        """SELECT agent_day_key AS Case_ID, business_date AS Date,
                  agent_id AS Agent_ID, agent_name AS Agent_Name,
                  coalesce(agent_name,'Agent') || ' [' || agent_id || ']' AS Agent_Selector,
                  team_leader AS Team_Leader, ops_manager AS Operations_Manager,
                  lob AS LOB, language AS Language,
                  final_ledger_status AS Result_Status,
                  planned_net_minutes/60.0 AS Planned_Net_Hours,
                  final_absence_minutes/60.0 AS Absence_Hours,
                  final_shrinkage_minutes/60.0 AS Shrinkage_Hours,
                  final_unmapped_minutes/60.0 AS Unmapped_Hours
           FROM mart.verint_final_absence_agent_day
           WHERE business_date BETWEEN ? AND ?
             AND (final_absence_day=true
                  OR final_ledger_status NOT IN ('CLEAR','ABSENCE_RECORDED'))
           ORDER BY business_date, lob, team_leader, agent_name""",
        (start, end),
    )
    path = folder / "ABSENCE_REVIEW_CASE_CURRENT.csv"
    counts.append((path.name, _atomic_csv(path, headers, rows)))
    files.append(path)
    files.append(_manifest(folder, "ABSENCE", start, end, counts))
    return SharedFeedResult("ABSENCE", tuple(files), sum(count for _, count in counts))


def publish_shared_feeds(
    conn: DatabaseConnection,
    config: Config,
    start: date,
    end: date,
) -> tuple[SharedFeedResult, ...]:
    """Refresh every stable collaboration feed after the Hub models finish."""

    # PCS publishes its six focused feeds inside its targeted report builder.
    # Generic full refreshes publish Absenteeism and Power BI together.
    # Local import avoids a module cycle: the Power BI publisher reuses the
    # atomic CSV and manifest primitives defined above.
    from .powerbi import publish_powerbi_feeds

    return (
        publish_absence_feeds(conn, config, start, end),
        publish_powerbi_feeds(conn, config, start, end),
    )
