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
from .rules import load_rulebook


PCS_FEED_SCHEMA_VERSION = "3"
PCS_AGENT_DAY_HEADERS = (
    "LOB", "Team Leader", "Agent Selector", "Agent ID", "Agent", "Date",
    "Ops Manager", "Language", "Inbound Call Legs", "PCS Status 1",
    "Q1 Nonblank", "Valid Q1", "Q1 Score Sum", "PCS Average",
    "Participation Rate", "Score <= 3", "Score > 3", "Invalid Q1",
    "Sample State", "Agent Day Key", "Data Through", "Feed Refreshed At",
    "PCS Rule Version", "PCS Rule SHA-256", "Metric Catalog Version",
    "Metric Catalog SHA-256",
)
PCS_COACHING_HEADERS = (
    "LOB", "Team Leader", "Agent Selector", "Agent", "Agent ID",
    "Priority", "Date", "Call Start", "Q1 Score", "Customer Comment",
    "Call Reference Number", "Language", "Coaching Key",
)
PCS_LOB_SCORECARD_HEADERS = (
    "LOB", "As Of Date", "Latest Day PCS", "Latest Day Participation",
    "Latest Day Valid Q1", "Current MTD PCS", "Prior MTD PCS",
    "MTD Change", "Current MTD Participation", "Prior MTD Participation",
    "Current MTD Valid Q1", "Current MTD PCS Status 1",
    "Current MTD Q1 Nonblank", "Current MTD Score <= 3",
    "Current MTD Score > 3", "Current MTD Inbound Call Legs",
    "Sample State", "Data Through", "Feed Refreshed At",
)
PCS_RESULTS_HEADERS = (
    "Period View", "Period Start", "Period End", "Scope Level", "LOB",
    "Team Leader", "Agent Selector", "Agent ID", "Agent", "Language",
    "PCS Average", "Participation Rate", "Valid Q1", "PCS Status 1",
    "Q1 Nonblank", "Score <= 3", "Score > 3", "Inbound Call Legs",
    "Sample State", "Data Through", "Feed Refreshed At",
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
    """Return copy-ready M for the two deliberately simple PCS queries."""

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


def _pcs_reporting_periods(as_of: date) -> tuple[tuple[str, date, date], ...]:
    """Return the stable novice-facing periods used in the PCS result feed."""

    month_start = as_of.replace(day=1)
    previous_end = month_start - timedelta(days=1)
    previous_start = previous_end.replace(day=1)
    previous_mtd_end = min(
        previous_end,
        previous_start + timedelta(days=as_of.day - 1),
    )
    return (
        ("Latest day", as_of, as_of),
        ("Current week", as_of - timedelta(days=as_of.weekday()), as_of),
        ("Current MTD", month_start, as_of),
        ("Previous MTD same days", previous_start, previous_mtd_end),
        ("Previous full month", previous_start, previous_end),
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


def pcs_lob_scorecard_rows(
    conn: DatabaseConnection,
    as_of: date,
    minimum_sample: int,
    refreshed: datetime,
) -> list[tuple[Any, ...]]:
    """Build one management-ready row per LOB plus a reconciling ALL row."""

    periods = {label: (start, end) for label, start, end in _pcs_reporting_periods(as_of)}
    latest = {
        str(row["lob"]): row
        for row in _pcs_scope_aggregates(conn, *periods["Latest day"], "LOB")
    }
    current = {
        str(row["lob"]): row
        for row in _pcs_scope_aggregates(conn, *periods["Current MTD"], "LOB")
    }
    prior = {
        str(row["lob"]): row
        for row in _pcs_scope_aggregates(conn, *periods["Previous MTD same days"], "LOB")
    }
    for target, label in ((latest, "Latest day"), (current, "Current MTD"), (prior, "Previous MTD same days")):
        target["ALL"] = _pcs_scope_aggregates(conn, *periods[label], "ALL")[0]
    lobs = ["ALL", *sorted((set(latest) | set(current) | set(prior)) - {"ALL"})]
    rows: list[tuple[Any, ...]] = []
    empty: dict[str, Any] = {}
    for lob in lobs:
        day = latest.get(lob, empty)
        mtd = current.get(lob, empty)
        previous = prior.get(lob, empty)
        mtd_score = mtd.get("pcs_average")
        prior_score = previous.get("pcs_average")
        valid = int(mtd.get("valid_q1") or 0)
        rows.append((
            lob, as_of,
            day.get("pcs_average"), day.get("participation_rate"),
            int(day.get("valid_q1") or 0),
            mtd_score, prior_score,
            (mtd_score - prior_score) if mtd_score is not None and prior_score is not None else None,
            mtd.get("participation_rate"), previous.get("participation_rate"),
            valid, int(mtd.get("pcs_status_1") or 0),
            int(mtd.get("q1_nonblank") or 0), int(mtd.get("low_scores") or 0),
            int(mtd.get("positive_scores") or 0), int(mtd.get("inbound_legs") or 0),
            "LOW SAMPLE" if valid < minimum_sample else "OK",
            as_of, refreshed,
        ))
    return rows


def pcs_result_rows(
    conn: DatabaseConnection,
    as_of: date,
    minimum_sample: int,
    refreshed: datetime,
) -> list[tuple[Any, ...]]:
    """Build filter-ready LOB, team, and agent scorecards for standard periods."""

    rows: list[tuple[Any, ...]] = []
    for period_label, period_start, period_end in _pcs_reporting_periods(as_of):
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
    return rows


def _publish_pcs_power_query_scripts(folder: Path) -> tuple[Path, ...]:
    data_types = (
        ("LOB", "type text"), ("Team Leader", "type text"),
        ("Agent Selector", "type text"), ("Agent ID", "type text"),
        ("Agent", "type text"), ("Date", "type date"),
        ("Ops Manager", "type text"), ("Language", "type text"),
        ("Inbound Call Legs", "Int64.Type"), ("PCS Status 1", "Int64.Type"),
        ("Q1 Nonblank", "Int64.Type"), ("Valid Q1", "Int64.Type"),
        ("Q1 Score Sum", "type number"), ("PCS Average", "type number"),
        ("Participation Rate", "type number"), ("Score <= 3", "Int64.Type"),
        ("Score > 3", "Int64.Type"), ("Invalid Q1", "Int64.Type"),
        ("Sample State", "type text"), ("Agent Day Key", "type text"),
        ("Data Through", "type date"), ("Feed Refreshed At", "type datetime"),
        ("PCS Rule Version", "type text"), ("PCS Rule SHA-256", "type text"),
        ("Metric Catalog Version", "type text"), ("Metric Catalog SHA-256", "type text"),
    )
    queue_types = (
        ("LOB", "type text"), ("Team Leader", "type text"),
        ("Agent Selector", "type text"), ("Agent", "type text"),
        ("Agent ID", "type text"), ("Priority", "type text"),
        ("Date", "type date"), ("Call Start", "type datetime"),
        ("Q1 Score", "type number"), ("Customer Comment", "type text"),
        ("Call Reference Number", "type text"), ("Language", "type text"),
        ("Coaching Key", "type text"),
    )
    lob_types = (
        ("LOB", "type text"), ("As Of Date", "type date"),
        ("Latest Day PCS", "type number"),
        ("Latest Day Participation", "type number"),
        ("Latest Day Valid Q1", "Int64.Type"),
        ("Current MTD PCS", "type number"),
        ("Prior MTD PCS", "type number"), ("MTD Change", "type number"),
        ("Current MTD Participation", "type number"),
        ("Prior MTD Participation", "type number"),
        ("Current MTD Valid Q1", "Int64.Type"),
        ("Current MTD PCS Status 1", "Int64.Type"),
        ("Current MTD Q1 Nonblank", "Int64.Type"),
        ("Current MTD Score <= 3", "Int64.Type"),
        ("Current MTD Score > 3", "Int64.Type"),
        ("Current MTD Inbound Call Legs", "Int64.Type"),
        ("Sample State", "type text"), ("Data Through", "type date"),
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
    specifications = (
        ("POWER_QUERY_PCS_DATA_SHAREPOINT.txt", "PCS_AGENT_DAY_CURRENT.csv", PCS_AGENT_DAY_HEADERS, data_types, True),
        ("POWER_QUERY_COACHING_QUEUE_SHAREPOINT.txt", "PCS_COACHING_OPPORTUNITY_CURRENT.csv", PCS_COACHING_HEADERS, queue_types, True),
        ("POWER_QUERY_PCS_LOB_SHAREPOINT.txt", "PCS_LOB_SCORECARD_CURRENT.csv", PCS_LOB_SCORECARD_HEADERS, lob_types, True),
        ("POWER_QUERY_PCS_RESULTS_SHAREPOINT.txt", "PCS_RESULTS_CURRENT.csv", PCS_RESULTS_HEADERS, result_types, True),
        ("POWER_QUERY_PCS_DATA_LOCAL.txt", "PCS_AGENT_DAY_CURRENT.csv", PCS_AGENT_DAY_HEADERS, data_types, False),
        ("POWER_QUERY_COACHING_QUEUE_LOCAL.txt", "PCS_COACHING_OPPORTUNITY_CURRENT.csv", PCS_COACHING_HEADERS, queue_types, False),
        ("POWER_QUERY_PCS_LOB_LOCAL.txt", "PCS_LOB_SCORECARD_CURRENT.csv", PCS_LOB_SCORECARD_HEADERS, lob_types, False),
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
    """Publish agent/day, coaching-opportunity and selector feeds for PCS."""

    start, end = _available_period(
        conn, "mart.agent_pcs_day", fallback_start, fallback_end,
    )
    folder = config.feed / "PCS"
    files: list[Path] = []
    counts: list[tuple[str, int]] = []

    metric_catalog = load_metric_catalog(config.home, config.metric_catalog)
    rulebook = load_rulebook(config.home, config.business_rules)
    method = metric_catalog.method_for("pcs_average", end, {})
    minimum_sample = int(method.minimum_sample) if method is not None else 1
    refreshed = datetime.now()
    _query_headers, rows = _rows(
        conn,
        """SELECT lob, team_leader,
                  coalesce(agent_name,'Agent') || ' [' || agent_id || ']',
                  agent_id, agent_name, business_date, ops_manager, language,
                  inbound_calls, pcs_status_calls, pcs_participation_responses,
                  survey_responses, pcs_score_sum, pcs_average,
                  pcs_participation_rate, low_score_responses,
                  top_box_responses, pcs_invalid_responses,
                  CASE WHEN survey_responses<? THEN 'LOW_SAMPLE' ELSE 'OK' END,
                  agent_id || '|' || business_date, ?, ?, ?, ?, ?, ?
           FROM mart.agent_pcs_day
           WHERE business_date BETWEEN ? AND ?
           ORDER BY business_date, lob, team_leader, agent_name, agent_id""",
        (
            minimum_sample, end, refreshed, rulebook.version, rulebook.sha256,
            metric_catalog.version, metric_catalog.sha256, start, end,
        ),
    )
    path = folder / "PCS_AGENT_DAY_CURRENT.csv"
    counts.append((path.name, _atomic_csv(path, PCS_AGENT_DAY_HEADERS, rows)))
    files.append(path)

    lob_rows = pcs_lob_scorecard_rows(conn, end, minimum_sample, refreshed)
    path = folder / "PCS_LOB_SCORECARD_CURRENT.csv"
    counts.append((path.name, _atomic_csv(path, PCS_LOB_SCORECARD_HEADERS, lob_rows)))
    files.append(path)

    result_rows = pcs_result_rows(conn, end, minimum_sample, refreshed)
    path = folder / "PCS_RESULTS_CURRENT.csv"
    counts.append((path.name, _atomic_csv(path, PCS_RESULTS_HEADERS, result_rows)))
    files.append(path)

    primary = config.pcs.primary_score_question
    primary_score = f"question_{primary}_score"
    allowed_scores = ", ".join(f"{value:g}" for value in config.pcs.allowed_scores)
    _query_headers, rows = _rows(
        conn,
        f"""SELECT coalesce(d.lob,c.lob), d.team_leader,
                   coalesce(d.canonical_name,c.agent_name,'Agent') || ' [' || c.agent_id || ']',
                   coalesce(d.canonical_name,c.agent_name), c.agent_id,
                   CASE WHEN c.{primary_score}<=2 THEN 'High' ELSE 'Normal' END,
                   c.business_date, c.call_start, c.{primary_score}, c.question_3,
                   c.call_reference_number, coalesce(d.language,c.language),
                   c.call_key
            FROM core.clean_call_leg c
            LEFT JOIN core.dim_agent d ON d.agent_id=c.agent_id
            WHERE c.business_date BETWEEN ? AND ?
              AND upper(coalesce(c.call_direction,''))='I'
              AND c.{primary_score} IN ({allowed_scores})
              AND c.{primary_score}<=?
            ORDER BY c.business_date DESC, d.team_leader,
                     coalesce(d.canonical_name,c.agent_name), c.call_start""",
        (start, end, config.pcs.negative_score_maximum),
    )
    path = folder / "PCS_COACHING_OPPORTUNITY_CURRENT.csv"
    counts.append((path.name, _atomic_csv(path, PCS_COACHING_HEADERS, rows)))
    files.append(path)

    _query_headers, rows = _rows(
        conn,
        """SELECT DISTINCT lob AS LOB, team_leader AS Team_Leader,
                  agent_id AS Agent_ID, agent_name AS Agent_Name,
                  coalesce(agent_name,'Agent') || ' [' || agent_id || ']' AS Agent_Selector
           FROM mart.agent_pcs_day
           WHERE business_date BETWEEN ? AND ?
           ORDER BY lob, team_leader, agent_name, agent_id""",
        (start, end),
    )
    headers = ["LOB", "Team Leader", "Agent ID", "Agent", "Agent Selector"]
    path = folder / "PCS_SCOPE_CURRENT.csv"
    counts.append((path.name, _atomic_csv(path, headers, rows)))
    files.append(path)
    files.extend(_publish_pcs_power_query_scripts(folder))
    files.append(_manifest(
        folder, "PCS", start, end, counts,
        schema_version=PCS_FEED_SCHEMA_VERSION,
        extra=(
            ("PCS rule version", rulebook.version, rulebook.sha256),
            ("Metric catalog version", metric_catalog.version, metric_catalog.sha256),
        ),
    ))
    return SharedFeedResult("PCS", tuple(files), sum(count for _, count in counts))


def publish_absence_feeds(
    conn: DatabaseConnection,
    config: Config,
    fallback_start: date,
    fallback_end: date,
) -> SharedFeedResult:
    """Publish reviewed attendance-led absence, component and case feeds."""

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
) -> tuple[SharedFeedResult, SharedFeedResult]:
    """Refresh every stable collaboration feed after the Hub models finish."""

    return (
        publish_pcs_feeds(conn, config, start, end),
        publish_absence_feeds(conn, config, start, end),
    )
