"""Materialize clean dimensions and WFM report marts."""

from __future__ import annotations

import hashlib
from collections import defaultdict
from dataclasses import dataclass
from datetime import date, datetime, time, timedelta
from pathlib import Path
from typing import Any, Iterable

from .config import Config
from .database import DatabaseConnection, DatabaseCursor
from .mapping import QueueMapping, load_queue_mapping
from .progress import ProgressCallback
from .rules import Rulebook, evaluate_formula, load_rulebook
from .utils import clip_intervals, interval_minutes, merge_intervals, subtract_intervals


@dataclass
class ModelSummary:
    start: date
    end: date
    attendance_rows: int = 0
    conformance_rows: int = 0
    correction_rows: int = 0
    rta_rows: int = 0
    forecast_rows: int = 0
    intraday_rows: int = 0
    pcs_rows: int = 0
    quality_rows: int = 0
    absence_rows: int = 0
    absence_event_rows: int = 0
    service_rows: int = 0


def _dicts(cursor: DatabaseCursor) -> list[dict[str, Any]]:
    columns = [item[0] for item in cursor.description]
    return [dict(zip(columns, row)) for row in cursor.fetchall()]


def resolve_period(
    conn: DatabaseConnection,
    config: Config,
    start: date | None,
    end: date | None,
    use_config_period: bool = True,
) -> tuple[date, date]:
    start = start or (config.period_start if use_config_period else None)
    end = end or (config.period_end if use_config_period else None)
    if start and end:
        if start > end:
            raise ValueError("Start date cannot be after end date")
        return start, end
    row = conn.execute(
        """
        SELECT min(d) AS "start_date [DATE]", max(d) AS "end_date [DATE]" FROM (
            SELECT min(schedule_date) AS d FROM raw.schedule_shift r JOIN meta.source_file f ON f.file_id=r.source_file_id AND f.active
            UNION ALL SELECT max(schedule_date) FROM raw.schedule_shift r JOIN meta.source_file f ON f.file_id=r.source_file_id AND f.active
            UNION ALL SELECT min(business_date) FROM raw.queue_actual r JOIN meta.source_file f ON f.file_id=r.source_file_id AND f.active
            UNION ALL SELECT max(business_date) FROM raw.queue_actual r JOIN meta.source_file f ON f.file_id=r.source_file_id AND f.active
            UNION ALL SELECT min(business_date) FROM raw.forecast_interval r JOIN meta.source_file f ON f.file_id=r.source_file_id AND f.active
            UNION ALL SELECT max(business_date) FROM raw.forecast_interval r JOIN meta.source_file f ON f.file_id=r.source_file_id AND f.active
            UNION ALL SELECT min(extract_date) FROM raw.lilo r JOIN meta.source_file f ON f.file_id=r.source_file_id AND f.active
            UNION ALL SELECT max(extract_date) FROM raw.lilo r JOIN meta.source_file f ON f.file_id=r.source_file_id AND f.active
            UNION ALL SELECT min(extract_date) FROM raw.agent_status r JOIN meta.source_file f ON f.file_id=r.source_file_id AND f.active
            UNION ALL SELECT max(extract_date) FROM raw.agent_status r JOIN meta.source_file f ON f.file_id=r.source_file_id AND f.active
            UNION ALL SELECT min(business_date) FROM raw.call_leg r JOIN meta.source_file f ON f.file_id=r.source_file_id AND f.active
            UNION ALL SELECT max(business_date) FROM raw.call_leg r JOIN meta.source_file f ON f.file_id=r.source_file_id AND f.active
        ) x WHERE d IS NOT NULL
        """
    ).fetchone()
    auto_start, auto_end = row if row else (None, None)
    start = start or auto_start
    end = end or auto_end
    if not start or not end:
        raise RuntimeError("No business dates were found. Load extracts first or supply --start and --end.")
    return start, end


def _load_schedules(conn: DatabaseConnection, start: date, end: date) -> list[dict[str, Any]]:
    return _dicts(conn.execute(
        """
        WITH source_choice AS (
            SELECT r.*,
                   f.file_name AS source_file,
                   f.modified_at,
                   dense_rank() OVER (
                       PARTITION BY r.schedule_date, coalesce(r.agent_id, 'NAME|' || upper(coalesce(r.agent_name,'')))
                       ORDER BY f.modified_at DESC NULLS LAST, f.file_name DESC
                   ) AS source_rank
            FROM raw.schedule_shift r
            JOIN meta.source_file f ON f.file_id=r.source_file_id AND f.active AND f.status='SUCCESS'
            WHERE r.schedule_date BETWEEN ? AND ?
        ), dedup AS (
            SELECT *, row_number() OVER (
                PARTITION BY schedule_date, coalesce(agent_id, 'NAME|' || upper(coalesce(agent_name,''))),
                             scheduled_start, scheduled_end, coalesce(assignment,'')
                ORDER BY source_row DESC
            ) AS row_rank
            FROM source_choice WHERE source_rank=1
        )
        SELECT source_file_id, source_row, schedule_date, agent_id_raw, agent_id,
               agent_name, scheduling_period, shift_assignment, assignment,
               assignment_type, scheduled_start, scheduled_end, shift_events,
               parse_ok, source_file
        FROM dedup WHERE row_rank=1
        ORDER BY schedule_date, agent_id, scheduled_start
        """,
        [start, end],
    ))


def _load_events(conn: DatabaseConnection, start: date, end: date) -> list[dict[str, Any]]:
    window_start = datetime.combine(start - timedelta(days=1), time.min)
    window_end = datetime.combine(end + timedelta(days=2), time.min)
    return _dicts(conn.execute(
        """
        SELECT source_file_id, source_row, event_index, schedule_date, agent_id,
               agent_name, activity, activity_type, event_start, event_end,
               parse_ok, source_file
        FROM (
            SELECT r.*, f.file_name AS source_file, f.modified_at,
                   row_number() OVER (
                       PARTITION BY r.agent_id, upper(coalesce(r.activity,'')), r.event_start, r.event_end
                       ORDER BY f.modified_at DESC NULLS LAST, f.file_name DESC, r.source_row DESC
                   ) AS row_rank
            FROM raw.schedule_event r
            JOIN meta.source_file f ON f.file_id=r.source_file_id AND f.active AND f.status='SUCCESS'
            WHERE r.parse_ok AND r.event_end >= ?
              AND r.event_start < ?
        ) x WHERE row_rank=1
        """,
        [window_start, window_end],
    ))


def _load_lilo(conn: DatabaseConnection, start: date, end: date) -> tuple[dict[tuple[date, str], list[dict[str, Any]]], set[date], set[str]]:
    rows = _dicts(conn.execute(
        """
        SELECT r.*, f.file_name AS source_file
        FROM raw.lilo r JOIN meta.source_file f ON f.file_id=r.source_file_id AND f.active AND f.status='SUCCESS'
        WHERE r.extract_date BETWEEN ? AND ?
        """,
        [start - timedelta(days=1), end + timedelta(days=1)],
    ))
    grouped: dict[tuple[date, str], list[dict[str, Any]]] = defaultdict(list)
    loaded_dates: set[date] = set()
    seen: set[str] = set()
    for row in rows:
        loaded_dates.add(row["extract_date"])
        if row["agent_id"]:
            grouped[(row["extract_date"], row["agent_id"])].append(row)
            seen.add(row["agent_id"])
    return grouped, loaded_dates, seen


def _load_statuses(
    conn: DatabaseConnection, start: date, end: date,
) -> tuple[dict[tuple[date, str], list[dict[str, Any]]], set[date]]:
    window_start = datetime.combine(start - timedelta(days=1), time.min)
    window_end = datetime.combine(end + timedelta(days=2), time.min)
    rows = _dicts(conn.execute(
        """
        SELECT source_file_id, source_row, serial_number, extract_date, agent_id,
               agent_name, status, actual_category, status_start, status_end,
               duration_seconds, queue, source_file
        FROM (
            SELECT r.*, f.file_name AS source_file, f.modified_at,
                   row_number() OVER (
                       PARTITION BY r.serial_number
                       ORDER BY f.modified_at DESC NULLS LAST, f.file_name DESC, r.source_row DESC
                   ) AS row_rank
            FROM raw.agent_status r
            JOIN meta.source_file f ON f.file_id=r.source_file_id AND f.active AND f.status='SUCCESS'
            WHERE r.status_start < ?
              AND r.status_end >= ?
        ) x WHERE row_rank=1 AND agent_id IS NOT NULL AND status_end > status_start
        ORDER BY agent_id, status_start
        """,
        [window_end, window_start],
    ))
    grouped: dict[tuple[date, str], list[dict[str, Any]]] = defaultdict(list)
    loaded_dates: set[date] = set()
    for row in rows:
        first_day = row["status_start"].date()
        last_day = (row["status_end"] - timedelta(microseconds=1)).date()
        for offset in range((last_day - first_day).days + 1):
            status_day = first_day + timedelta(days=offset)
            loaded_dates.add(status_day)
            grouped[(status_day, row["agent_id"])].append(row)
    return grouped, loaded_dates


def _statuses_for_shift(
    agent_id: str,
    start: datetime | None,
    end: datetime | None,
    statuses_by_day: dict[tuple[date, str], list[dict[str, Any]]],
) -> list[dict[str, Any]]:
    if not start or not end:
        return []
    rows: dict[tuple[Any, ...], dict[str, Any]] = {}
    first_day = start.date()
    last_day = (end - timedelta(microseconds=1)).date()
    for offset in range((last_day - first_day).days + 1):
        for row in statuses_by_day.get((first_day + timedelta(days=offset), agent_id), []):
            key = (row["serial_number"], row["status_start"], row["status_end"], row["source_file_id"])
            rows[key] = row
    return sorted(rows.values(), key=lambda row: (row["status_start"], row["source_row"]))


def _build_agents(conn: DatabaseConnection) -> dict[str, dict[str, Any]]:
    conn.execute("DELETE FROM core.dim_agent")
    conn.execute(
        """
        WITH roster_ranked AS (
            SELECT r.agent_id, r.agent_name,
                   row_number() OVER (
                       PARTITION BY r.agent_id
                       ORDER BY f.modified_at DESC NULLS LAST, r.source_row DESC
                   ) AS row_rank
            FROM raw.schedule_shift r
            JOIN meta.source_file f ON f.file_id=r.source_file_id AND f.active
            WHERE r.agent_id IS NOT NULL AND r.agent_name IS NOT NULL
        ), roster AS (
            SELECT agent_id, agent_name FROM roster_ranked WHERE row_rank=1
        ), actual_ranked AS (
            SELECT r.agent_id, r.agent_name,
                   row_number() OVER (
                       PARTITION BY r.agent_id
                       ORDER BY f.modified_at DESC NULLS LAST, r.source_row DESC
                   ) AS row_rank
            FROM raw.lilo r
            JOIN meta.source_file f ON f.file_id=r.source_file_id AND f.active
            WHERE r.agent_id IS NOT NULL AND r.agent_name IS NOT NULL
        ), actual_names AS (
            SELECT agent_id, agent_name FROM actual_ranked WHERE row_rank=1
        ), ids AS (
            SELECT agent_id FROM roster UNION SELECT agent_id FROM actual_names
            UNION SELECT agent_id FROM raw.agent_status r JOIN meta.source_file f ON f.file_id=r.source_file_id AND f.active WHERE agent_id IS NOT NULL
            UNION SELECT agent_id FROM raw.call_leg r JOIN meta.source_file f ON f.file_id=r.source_file_id AND f.active WHERE agent_id IS NOT NULL
            UNION SELECT agent_id FROM raw.fte_agent r JOIN meta.source_file f ON f.file_id=r.source_file_id AND f.active WHERE agent_id IS NOT NULL
        ), fte_ranked AS (
                SELECT r.agent_id, r.agent_name, r.employment_status, r.team_leader,
                       r.ops_manager, r.lob, r.market, r.language, r.location,
                       r.city, r.fte, r.end_date,
                       row_number() OVER (
                    PARTITION BY r.agent_id
                    ORDER BY CASE WHEN upper(coalesce(employment_status,''))='ACTIVE' THEN 0 ELSE 1 END,
                             end_date DESC NULLS LAST, f.modified_at DESC NULLS LAST, source_row DESC
                ) row_rank
                FROM raw.fte_agent r JOIN meta.source_file f ON f.file_id=r.source_file_id AND f.active
                WHERE r.agent_id IS NOT NULL
        ), fte AS (
            SELECT agent_id, agent_name, employment_status, team_leader, ops_manager,
                   lob, market, language, location, city, fte
            FROM fte_ranked WHERE row_rank=1
        )
        INSERT INTO core.dim_agent
        SELECT ids.agent_id,
               coalesce(fte.agent_name, roster.agent_name, actual_names.agent_name) AS canonical_name,
               fte.employment_status, fte.team_leader, fte.ops_manager, fte.lob, fte.market,
               fte.language, fte.location, fte.city, fte.fte,
               CASE WHEN fte.agent_id IS NOT NULL THEN 'Agent ID' ELSE 'Unmatched to FTE' END AS match_method
        FROM ids LEFT JOIN roster USING(agent_id) LEFT JOIN actual_names USING(agent_id) LEFT JOIN fte USING(agent_id)
        """
    )
    return {row["agent_id"]: row for row in _dicts(conn.execute("SELECT * FROM core.dim_agent"))}


def _events_by_agent(events: list[dict[str, Any]]) -> dict[str, list[dict[str, Any]]]:
    grouped: dict[str, list[dict[str, Any]]] = defaultdict(list)
    for event in events:
        if event["agent_id"]:
            grouped[event["agent_id"]].append(event)
    return grouped


def _shift_events(shift: dict[str, Any], events_by_agent: dict[str, list[dict[str, Any]]]) -> list[dict[str, Any]]:
    if not shift["scheduled_start"] or not shift["scheduled_end"] or not shift["agent_id"]:
        return []
    return [
        event for event in events_by_agent.get(shift["agent_id"], [])
        if event["event_start"] < shift["scheduled_end"] and event["event_end"] > shift["scheduled_start"]
    ]


def _planned_intervals(events: list[dict[str, Any]], categories: set[str]) -> list[tuple[datetime, datetime]]:
    return merge_intervals(
        (event["event_start"], event["event_end"])
        for event in events if event["activity_type"] in categories
    )


def _lilo_boundaries(
    shift: dict[str, Any],
    lilo: dict[tuple[date, str], list[dict[str, Any]]],
    loaded_dates: set[date],
) -> tuple[datetime | None, datetime | None, bool, bool, str | None]:
    agent_id = shift["agent_id"]
    start = shift["scheduled_start"]
    end = shift["scheduled_end"]
    if start and end:
        required = {start.date() + timedelta(days=n) for n in range((end.date() - start.date()).days + 1)}
    else:
        required = {shift["schedule_date"]}
    rows = [row for day in sorted(required) for row in lilo.get((day, agent_id), [])]
    source_loaded = required <= loaded_dates
    row_present = bool(rows)
    files = "; ".join(sorted({row["source_file"] for row in rows})) or None
    first_candidates = sorted({row["first_login"] for row in rows if row["first_login"]})
    last_candidates = sorted({value for row in rows for value in (row["raw_last_logout"], row["last_logout"]) if value})
    if not start or not end:
        first = min(first_candidates) if first_candidates else None
        last = max(last_candidates) if last_candidates else None
        return first, last, source_loaded, row_present, files
    first_match = [value for value in first_candidates if start - timedelta(hours=4) <= value <= end]
    last_match = [value for value in last_candidates if start <= value <= end + timedelta(hours=4)]
    return (
        min(first_match) if first_match else None,
        max(last_match) if last_match else None,
        source_loaded,
        row_present,
        files,
    )


ATTENDANCE_COLUMNS = [
    "agent_day_key", "business_date", "agent_id", "agent_name", "team_leader", "ops_manager", "lob", "market", "language", "location",
    "scheduled_start", "scheduled_end", "scheduled_minutes", "assignment", "assignment_type", "planned_absence_minutes", "first_login", "last_logout",
    "source_loaded", "lilo_row_present", "seen_in_lilo", "raw_late_minutes", "raw_early_leave_minutes", "uncoded_late_minutes", "uncoded_early_leave_minutes",
    "no_show_minutes", "worked_span_minutes", "attendance_result", "attendance_percent", "schedule_source", "lilo_source",
    "actual_first_seen", "actual_last_seen", "actual_evidence", "status_covered_minutes", "status_source",
]


def _insert_dicts(conn: DatabaseConnection, table: str, columns: list[str], rows: list[dict[str, Any]]) -> None:
    if not rows:
        return
    row_placeholders = "(" + ", ".join("?" for _ in columns) + ")"
    batch_size = max(1, min(500, conn.max_variable_number // len(columns)))
    for offset in range(0, len(rows), batch_size):
        batch = rows[offset : offset + batch_size]
        sql = f"INSERT INTO {table} ({', '.join(columns)}) VALUES " + ", ".join(row_placeholders for _ in batch)
        params = [row.get(column) for row in batch for column in columns]
        conn.execute(sql, params)


def _build_attendance(
    conn: DatabaseConnection,
    rulebook: Rulebook,
    schedules: list[dict[str, Any]],
    events_by_agent: dict[str, list[dict[str, Any]]],
    lilo: dict[tuple[date, str], list[dict[str, Any]]],
    loaded_dates: set[date],
    seen_ids: set[str],
    agents: dict[str, dict[str, Any]],
    statuses_by_day: dict[tuple[date, str], list[dict[str, Any]]],
    status_loaded_dates: set[date],
) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    tolerance = rulebook.late_tolerance_minutes
    for shift in schedules:
        agent_id = shift["agent_id"]
        if not agent_id:
            continue
        start, end = shift["scheduled_start"], shift["scheduled_end"]
        scheduled_minutes = int((end - start).total_seconds() // 60) if start and end and end > start else 0
        events = _shift_events(shift, events_by_agent)
        # Verint activities are the corrected final ledger. They remain useful
        # context, but they must never erase a gap observed in LILO/status.
        planned = _planned_intervals(events, {"Planned absence", "Planned adjustment"})
        planned_absence = scheduled_minutes if shift["assignment_type"] == "Planned absence" else (interval_minutes(start, end, planned) if start and end else 0)
        first, last, source_loaded, row_present, lilo_source = _lilo_boundaries(shift, lilo, loaded_dates)
        statuses = _statuses_for_shift(agent_id, start, end, statuses_by_day)
        if start and end and end > start:
            status_category, status_covered, status_exclusive = _exclusive_category_minutes(start, end, statuses)
            required_status_dates = {
                start.date() + timedelta(days=offset)
                for offset in range(((end - timedelta(microseconds=1)).date() - start.date()).days + 1)
            }
        else:
            status_category, status_covered, status_exclusive = {}, 0, []
            required_status_dates = {shift["schedule_date"]}
        status_source_loaded = required_status_dates <= status_loaded_dates
        status_sources = "; ".join(sorted({row["source_file"] for row in statuses})) or None
        active_status = [
            row for row in status_exclusive
            if row["actual_category"] in {"Productive", "Auxiliary", "Lunch", "Break"}
        ]
        status_first = min((row["interval_start"] for row in active_status), default=None)
        status_last = max((row["interval_end"] for row in active_status), default=None)
        actual_first = min((value for value in (first, status_first) if value is not None), default=None)
        actual_last = max((value for value in (last, status_last) if value is not None), default=None)
        evidence_parts = []
        if first is not None or last is not None or row_present:
            evidence_parts.append("LILO")
        if statuses:
            evidence_parts.append("AGENT_STATUS")
        actual_evidence = "+".join(evidence_parts) or "NONE"
        raw_late = max(0, int((first - start).total_seconds() // 60)) if first and start else 0
        raw_early = max(0, int((end - last).total_seconds() // 60)) if last and end else 0
        usable_pair = bool(actual_first and actual_last and start and end and actual_last >= actual_first and actual_last > start and actual_first < end)
        late_segments = [(start, min(actual_first, end))] if usable_pair and actual_first > start else []
        early_segments = [(max(actual_last, start), end)] if usable_pair and end > actual_last else []
        late = sum(int((b - a).total_seconds() // 60) for a, b in late_segments)
        early = sum(int((b - a).total_seconds() // 60) for a, b in early_segments)
        late = 0 if late <= tolerance else late
        early = 0 if early <= tolerance else early
        no_show = scheduled_minutes if (
            source_loaded and row_present and first is None and last is None
            and not active_status and shift["assignment_type"] != "Off"
        ) else 0
        worked_span = int((actual_last - actual_first).total_seconds() // 60) if actual_first and actual_last and actual_last >= actual_first else 0
        parse_ok = bool(shift["parse_ok"])
        if not parse_ok and shift["assignment_type"] != "Off":
            result = "Schedule parse error"
        elif shift["assignment_type"] == "Off":
            result = "Off"
        elif not source_loaded and not status_source_loaded:
            result = "Data not loaded"
        elif no_show:
            result = "No show"
        elif actual_first is None and actual_last is None:
            result = "Missing actual evidence"
        elif actual_first is None or actual_last is None:
            result = "Incomplete actual evidence"
        elif not usable_pair:
            result = "No schedule overlap"
        elif late and early:
            result = "Late + early leave"
        elif late:
            result = "Late"
        elif early:
            result = "Early leave"
        else:
            result = "Present"
        attendance_pct = None
        trusted_results = {"No show", "Present", "Late", "Early leave", "Late + early leave"}
        if shift["assignment_type"] != "Off" and scheduled_minutes and result in trusted_results:
            attendance_pct = max(0.0, 1 - (no_show + late + early) / scheduled_minutes)
        agent = agents.get(agent_id, {})
        rows.append({
            "agent_day_key": f"{shift['schedule_date']:%Y%m%d}-{agent_id}",
            "business_date": shift["schedule_date"], "agent_id": agent_id,
            "agent_name": agent.get("canonical_name") or shift["agent_name"],
            "team_leader": agent.get("team_leader"), "ops_manager": agent.get("ops_manager"),
            "lob": agent.get("lob"), "market": agent.get("market"), "language": agent.get("language"), "location": agent.get("location"),
            "scheduled_start": start, "scheduled_end": end, "scheduled_minutes": scheduled_minutes,
            "assignment": shift["assignment"], "assignment_type": shift["assignment_type"], "planned_absence_minutes": planned_absence,
            "first_login": first, "last_logout": last, "source_loaded": source_loaded or status_source_loaded, "lilo_row_present": row_present,
            "seen_in_lilo": agent_id in seen_ids, "raw_late_minutes": raw_late, "raw_early_leave_minutes": raw_early,
            "uncoded_late_minutes": late, "uncoded_early_leave_minutes": early, "no_show_minutes": no_show,
            "worked_span_minutes": worked_span, "attendance_result": result, "attendance_percent": attendance_pct,
            "schedule_source": shift["source_file"], "lilo_source": lilo_source,
            "actual_first_seen": actual_first, "actual_last_seen": actual_last,
            "actual_evidence": actual_evidence, "status_covered_minutes": status_covered,
            "status_source": status_sources,
            "_events": events, "_late_segments": late_segments, "_early_segments": early_segments,
            "_status_exclusive": status_exclusive, "_status_category": status_category,
        })
    conn.execute("DELETE FROM mart.attendance_agent_day")
    _insert_dicts(conn, "mart.attendance_agent_day", ATTENDANCE_COLUMNS, rows)
    return rows


def _exclusive_category_minutes(
    shift_start: datetime,
    shift_end: datetime,
    statuses: list[dict[str, Any]],
) -> tuple[dict[str, int], int, list[dict[str, Any]]]:
    clipped: list[dict[str, Any]] = []
    for status in statuses:
        start, end = max(shift_start, status["status_start"]), min(shift_end, status["status_end"])
        if end > start:
            clipped.append({**status, "clip_start": start, "clip_end": end})
    boundaries = sorted({shift_start, shift_end, *(row["clip_start"] for row in clipped), *(row["clip_end"] for row in clipped)})
    minutes: dict[str, int] = defaultdict(int)
    covered_segments: list[tuple[datetime, datetime]] = []
    exclusive: list[dict[str, Any]] = []
    for left, right in zip(boundaries, boundaries[1:]):
        active = [row for row in clipped if row["clip_start"] < right and row["clip_end"] > left]
        if not active:
            continue
        # During minute-rounded transitions, the newest starting state wins.
        chosen = max(active, key=lambda row: (row["status_start"], row["source_row"]))
        segment_minutes = int((right - left).total_seconds() // 60)
        if segment_minutes <= 0:
            continue
        minutes[chosen["actual_category"]] += segment_minutes
        covered_segments.append((left, right))
        exclusive.append({**chosen, "interval_start": left, "interval_end": right})
    covered = int(sum((b - a).total_seconds() for a, b in merge_intervals(covered_segments)) // 60)
    return dict(minutes), covered, exclusive


CONFORMANCE_COLUMNS = [
    "agent_day_key", "business_date", "agent_id", "scheduled_minutes", "scheduled_net_minutes", "planned_absence_minutes", "planned_lunch_minutes", "planned_break_minutes",
    "productive_minutes", "auxiliary_minutes", "break_minutes", "lunch_minutes", "unavailable_minutes", "logged_off_minutes", "status_covered_minutes",
    "status_coverage_percent", "login_span_minutes", "measurement_basis", "worked_minutes", "conformance_percent", "break_overrun_minutes", "lunch_overrun_minutes", "unexplained_minutes",
]


def _build_conformance(
    conn: DatabaseConnection,
    config: Config,
    attendance: list[dict[str, Any]],
    statuses_by_agent: dict[str, list[dict[str, Any]]],
) -> tuple[list[dict[str, Any]], dict[str, list[dict[str, Any]]]]:
    output: list[dict[str, Any]] = []
    exclusive_by_key: dict[str, list[dict[str, Any]]] = {}
    for row in attendance:
        start, end = row["scheduled_start"], row["scheduled_end"]
        if not start or not end or end <= start or row["assignment_type"] == "Off":
            continue
        events = row["_events"]
        planned_absence = row["planned_absence_minutes"]
        lunch_intervals = _planned_intervals(events, {"Lunch"})
        break_intervals = _planned_intervals(events, {"Break"})
        planned_lunch = interval_minutes(start, end, lunch_intervals)
        planned_break = interval_minutes(start, end, break_intervals)
        scheduled_net = max(0, row["scheduled_minutes"] - planned_absence - planned_lunch)
        category, covered, exclusive = _exclusive_category_minutes(start, end, statuses_by_agent.get(row["agent_id"], []))
        coverage = min(1.0, covered / row["scheduled_minutes"]) if row["scheduled_minutes"] else None
        login_span = interval_minutes(start, end, [(row["first_login"], row["last_logout"])]) if row["first_login"] and row["last_logout"] and row["last_logout"] > row["first_login"] else 0
        if coverage is not None and coverage >= config.rules.minimum_status_coverage:
            basis = "Agent Status"
        elif login_span > 0:
            basis = "LILO span"
        else:
            basis = "None"
        productive = category.get("Productive", 0)
        auxiliary = category.get("Auxiliary", 0)
        break_minutes = category.get("Break", 0)
        lunch_minutes = category.get("Lunch", 0)
        unavailable = category.get("Unavailable", 0)
        logged_off = category.get("Logged Off", 0)
        if basis == "Agent Status":
            worked = productive + auxiliary + min(break_minutes, config.rules.break_minutes)
        elif basis == "LILO span":
            worked = max(0, login_span - planned_lunch - planned_absence)
        else:
            worked = None
        conformance = min(1.0, worked / scheduled_net) if worked is not None and scheduled_net else None
        item = {
            "agent_day_key": row["agent_day_key"], "business_date": row["business_date"], "agent_id": row["agent_id"],
            "scheduled_minutes": row["scheduled_minutes"], "scheduled_net_minutes": scheduled_net,
            "planned_absence_minutes": planned_absence, "planned_lunch_minutes": planned_lunch, "planned_break_minutes": planned_break,
            "productive_minutes": productive, "auxiliary_minutes": auxiliary, "break_minutes": break_minutes,
            "lunch_minutes": lunch_minutes, "unavailable_minutes": unavailable, "logged_off_minutes": logged_off,
            "status_covered_minutes": covered, "status_coverage_percent": coverage, "login_span_minutes": login_span,
            "measurement_basis": basis, "worked_minutes": worked, "conformance_percent": conformance,
            "break_overrun_minutes": max(0, break_minutes - config.rules.break_minutes) if basis == "Agent Status" else 0,
            "lunch_overrun_minutes": max(0, lunch_minutes - max(planned_lunch, config.rules.lunch_minutes)) if basis == "Agent Status" else 0,
            "unexplained_minutes": max(0, scheduled_net - covered) if basis == "Agent Status" else 0,
        }
        output.append(item)
        exclusive_by_key[row["agent_day_key"]] = exclusive
    conn.execute("DELETE FROM mart.conformance_agent_day")
    _insert_dicts(conn, "mart.conformance_agent_day", CONFORMANCE_COLUMNS, output)
    return output, exclusive_by_key


def _nearby_merge(intervals: Iterable[tuple[datetime, datetime]], tolerance_minutes: int) -> list[tuple[datetime, datetime]]:
    ordered = sorted(intervals)
    merged: list[list[datetime]] = []
    tolerance = timedelta(minutes=tolerance_minutes)
    for start, end in ordered:
        if not merged or start > merged[-1][1] + tolerance:
            merged.append([start, end])
        else:
            merged[-1][1] = max(merged[-1][1], end)
    return [(a, b) for a, b in merged]


CORRECTION_COLUMNS = [
    "correction_id", "business_date", "agent_id", "agent_name", "team_leader", "ops_manager", "lob", "scheduled_start", "scheduled_end", "first_login", "last_logout",
    "priority", "detected_issue", "gap_start", "gap_end", "gap_minutes", "confidence", "suggested_activity", "source_file",
    "confirmed_activity", "validation_status", "owner", "comment", "injected_date",
    "observed_source", "verint_reconciliation", "verint_activity", "verint_category",
    "verint_overlap_minutes", "verint_source_file",
]

VERINT_EXCEPTION_COLUMNS = [
    "exception_key", "agent_day_key", "business_date", "agent_id", "agent_name",
    "activity", "category", "event_start", "event_end", "minutes",
    "exception_type", "source_file", "rule_version", "rule_sha256",
]


def _correction_id(row: dict[str, Any], issue: str, start: datetime | None, end: datetime | None) -> str:
    clean = "".join(char for char in issue.upper() if char.isalnum())
    start_key = start.strftime("%H%M") if start else "DAY"
    end_key = end.strftime("%H%M") if end else "DAY"
    return f"{row['business_date']:%Y%m%d}-{row['agent_id']}-{start_key}-{end_key}-{clean}"


def _final_verint_events(base: dict[str, Any], rulebook: Rulebook) -> list[dict[str, Any]]:
    """Return post-day Verint activities that can explain an observed gap."""
    candidates: list[dict[str, Any]] = []
    assignment_rule = rulebook.classify_activity(base.get("assignment"))
    if assignment_rule is not None and not assignment_rule.working and assignment_rule.category != "OFF":
        candidates.append({
            "activity": base.get("assignment"), "category": assignment_rule.category,
            "start": base.get("scheduled_start"), "end": base.get("scheduled_end"),
            "source_file": base.get("schedule_source"), "rule": assignment_rule,
        })
    ignored = {"OFF", "LUNCH", "BREAK", "NO_ACTIVITY", "PRODUCTION"}
    for event in base.get("_events", []):
        rule = rulebook.classify_activity(event.get("activity"))
        if rule is None or rule.working or rule.category in ignored:
            continue
        candidates.append({
            "activity": event.get("activity"), "category": rule.category,
            "start": event.get("event_start"), "end": event.get("event_end"),
            "source_file": event.get("source_file"), "rule": rule,
        })
    deduped: dict[tuple[Any, ...], dict[str, Any]] = {}
    for item in candidates:
        if item["start"] and item["end"] and item["end"] > item["start"]:
            deduped[(item["activity"], item["start"], item["end"], item["source_file"])] = item
    return list(deduped.values())


def _build_corrections(
    conn: DatabaseConnection,
    rulebook: Rulebook,
    attendance: list[dict[str, Any]],
) -> list[dict[str, Any]]:
    output: list[dict[str, Any]] = []

    def add(
        base: dict[str, Any], issue: str, start: datetime | None, end: datetime | None,
        minutes: int, priority: int, confidence: str, activity: str,
        observed_source: str, source: str | None = None,
    ) -> None:
        output.append({
            "correction_id": _correction_id(base, issue, start, end), "business_date": base["business_date"], "agent_id": base["agent_id"],
            "agent_name": base["agent_name"], "team_leader": base["team_leader"], "ops_manager": base["ops_manager"], "lob": base["lob"],
            "scheduled_start": base["scheduled_start"], "scheduled_end": base["scheduled_end"], "first_login": base["first_login"], "last_logout": base["last_logout"],
            "priority": priority, "detected_issue": issue, "gap_start": start, "gap_end": end, "gap_minutes": minutes,
            "confidence": confidence, "suggested_activity": activity, "source_file": source or base["schedule_source"],
            "observed_source": observed_source,
        })

    status_tolerance = rulebook.status_gap_tolerance_minutes
    match_tolerance = rulebook.verint_match_tolerance_minutes
    for row in attendance:
        start, end = row["scheduled_start"], row["scheduled_end"]
        if row["attendance_result"] == "No show" and start and end:
            minutes = int((end - start).total_seconds() // 60)
            if minutes > 0:
                source = "; ".join(value for value in (row["lilo_source"], row["status_source"]) if value) or row["schedule_source"]
                add(row, "No show", start, end, minutes, 1, "High", "Review absence reason", row["actual_evidence"], source)
        if row["uncoded_late_minutes"] > 0:
            for gap_start, gap_end in row["_late_segments"]:
                minutes = int((gap_end - gap_start).total_seconds() // 60)
                if minutes > 0:
                    source = "; ".join(value for value in (row["lilo_source"], row["status_source"]) if value) or row["schedule_source"]
                    add(row, "Late", gap_start, gap_end, minutes, 2, "High", "Late", row["actual_evidence"], source)
        if row["uncoded_early_leave_minutes"] > 0:
            for gap_start, gap_end in row["_early_segments"]:
                minutes = int((gap_end - gap_start).total_seconds() // 60)
                if minutes > 0:
                    source = "; ".join(value for value in (row["lilo_source"], row["status_source"]) if value) or row["schedule_source"]
                    add(row, "Early leave", gap_start, gap_end, minutes, 3, "High", "Early Leaving", row["actual_evidence"], source)
        if row["attendance_result"] == "Incomplete actual evidence":
            source = "; ".join(value for value in (row["lilo_source"], row["status_source"]) if value) or row["schedule_source"]
            add(row, "Incomplete actual evidence", None, None, 0, 9, "Review", "Schedule Correction", row["actual_evidence"], source)

        actual_first, actual_last = row["actual_first_seen"], row["actual_last_seen"]
        if not actual_first or not actual_last or actual_last <= actual_first:
            continue
        exclusive = row.get("_status_exclusive", [])
        for category, issue, priority in (("Logged Off", "Mid-shift logged off", 4), ("Unavailable", "Unavailable in shift", 7)):
            raw = [(item["interval_start"], item["interval_end"]) for item in exclusive if item["actual_category"] == category]
            for hit_start, hit_end in clip_intervals(actual_first, actual_last, _nearby_merge(raw, status_tolerance)):
                minutes = int((hit_end - hit_start).total_seconds() // 60)
                if minutes > status_tolerance:
                    confidence = "High" if category == "Logged Off" else "Review"
                    add(row, issue, hit_start, hit_end, minutes, priority, confidence, "General Unavailability", "AGENT_STATUS", row["status_source"])

    by_key = {row["agent_day_key"]: row for row in attendance}
    for item in output:
        base = by_key[item["business_date"].strftime("%Y%m%d") + "-" + item["agent_id"]]
        gap_start, gap_end = item["gap_start"], item["gap_end"]
        matches: list[tuple[int, dict[str, Any]]] = []
        if gap_start and gap_end:
            for event in _final_verint_events(base, rulebook):
                overlap = interval_minutes(gap_start, gap_end, [(event["start"], event["end"])])
                if overlap > 0:
                    matches.append((overlap, event))
        if matches:
            overlap, event = max(matches, key=lambda pair: (pair[0], pair[1]["start"]))
            status = "CORRECTED" if overlap >= max(1, item["gap_minutes"] - match_tolerance) else "PARTIAL"
            item.update({
                "verint_reconciliation": status, "verint_activity": event["activity"],
                "verint_category": event["category"], "verint_overlap_minutes": overlap,
                "verint_source_file": event["source_file"],
            })
        else:
            item.update({
                "verint_reconciliation": "NOT_APPLICABLE" if not gap_start or not gap_end else "NOT_CORRECTED",
                "verint_activity": None, "verint_category": None, "verint_overlap_minutes": 0,
                "verint_source_file": None,
            })

    actions = {row["correction_id"]: row for row in _dicts(conn.execute("SELECT * FROM core.correction_action"))}
    deduped: dict[str, dict[str, Any]] = {}
    for item in output:
        action = actions.get(item["correction_id"], {})
        item.update({
            "confirmed_activity": action.get("confirmed_activity"),
            "validation_status": action.get("validation_status") or "Open",
            "owner": action.get("owner"), "comment": action.get("comment"), "injected_date": action.get("injected_date"),
        })
        deduped[item["correction_id"]] = item
    output = sorted(deduped.values(), key=lambda item: (item["business_date"], item["priority"], -item["gap_minutes"], item["agent_id"]))
    conn.execute("DELETE FROM mart.correction_candidate")
    _insert_dicts(conn, "mart.correction_candidate", CORRECTION_COLUMNS, output)

    exceptions: list[dict[str, Any]] = []
    gaps_by_key: dict[str, list[dict[str, Any]]] = defaultdict(list)
    for item in output:
        gaps_by_key[f"{item['business_date']:%Y%m%d}-{item['agent_id']}"].append(item)
    for base in attendance:
        for event in _final_verint_events(base, rulebook):
            has_observed_overlap = any(
                gap["gap_start"] and gap["gap_end"]
                and interval_minutes(gap["gap_start"], gap["gap_end"], [(event["start"], event["end"])]) > 0
                for gap in gaps_by_key.get(base["agent_day_key"], [])
            )
            if has_observed_overlap:
                continue
            minutes = interval_minutes(base["scheduled_start"], base["scheduled_end"], [(event["start"], event["end"])])
            if minutes <= 0:
                continue
            key = hashlib.sha256(
                f"{base['agent_day_key']}|{event['activity']}|{event['start']}|{event['end']}".encode("utf-8")
            ).hexdigest()
            exceptions.append({
                "exception_key": key, "agent_day_key": base["agent_day_key"],
                "business_date": base["business_date"], "agent_id": base["agent_id"],
                "agent_name": base["agent_name"], "activity": event["activity"],
                "category": event["category"], "event_start": event["start"],
                "event_end": event["end"], "minutes": minutes,
                "exception_type": "VERINT_FINAL_WITHOUT_OBSERVED_GAP",
                "source_file": event["source_file"], "rule_version": rulebook.version,
                "rule_sha256": rulebook.sha256,
            })
    conn.execute("DELETE FROM mart.verint_final_exception")
    _insert_dicts(conn, "mart.verint_final_exception", VERINT_EXCEPTION_COLUMNS, exceptions)
    return output


RTA_COLUMNS = [
    "snapshot_at", "agent_id", "agent_name", "team_leader", "lob", "scheduled_start", "scheduled_end", "planned_activity", "actual_status", "actual_category",
    "status_start", "minutes_in_status", "rta_result", "severity", "freshness", "source_file",
]


def _build_rta(
    conn: DatabaseConnection,
    config: Config,
    schedules: list[dict[str, Any]],
    events_by_agent: dict[str, list[dict[str, Any]]],
    statuses_by_agent: dict[str, list[dict[str, Any]]],
    agents: dict[str, dict[str, Any]],
) -> list[dict[str, Any]]:
    all_statuses = [row for rows in statuses_by_agent.values() for row in rows]
    if not all_statuses:
        conn.execute("DELETE FROM mart.rta_snapshot")
        return []
    snapshot = max(row["status_start"] for row in all_statuses)
    output: list[dict[str, Any]] = []
    for shift in schedules:
        if not shift["agent_id"] or not shift["scheduled_start"] or not shift["scheduled_end"]:
            continue
        if not (shift["scheduled_start"] <= snapshot < shift["scheduled_end"]):
            continue
        events = _shift_events(shift, events_by_agent)
        active_event = next((event for event in events if event["event_start"] <= snapshot < event["event_end"]), None)
        planned = active_event["activity_type"] if active_event else shift["assignment_type"]
        history = [row for row in statuses_by_agent.get(shift["agent_id"], []) if row["status_start"] <= snapshot]
        latest = max(history, key=lambda row: row["status_start"]) if history else None
        minutes = int((snapshot - latest["status_start"]).total_seconds() // 60) if latest else None
        stale = latest is None or minutes is None or minutes > config.rules.rta_stale_minutes
        actual = latest["actual_category"] if latest else None
        expected = {
            "Lunch": {"Lunch"}, "Break": {"Break"}, "Planned absence": {"Logged Off"},
            "Planned adjustment": {"Logged Off", "Unavailable"}, "Work": {"Productive", "Auxiliary"},
            "Non-phone planned": {"Auxiliary", "Productive"}, "Other planned": {"Auxiliary", "Productive"},
        }.get(planned, {"Productive", "Auxiliary"})
        if stale:
            result, severity, freshness = "Stale data", "Review", "Stale"
        elif actual in expected:
            result, severity, freshness = "In adherence", "OK", "Current"
        elif planned == "Planned absence" and actual in {"Productive", "Auxiliary"}:
            result, severity, freshness = "Review", "Review", "Current"
        else:
            result, severity, freshness = "Out of adherence", "High", "Current"
        agent = agents.get(shift["agent_id"], {})
        output.append({
            "snapshot_at": snapshot, "agent_id": shift["agent_id"], "agent_name": agent.get("canonical_name") or shift["agent_name"],
            "team_leader": agent.get("team_leader"), "lob": agent.get("lob"), "scheduled_start": shift["scheduled_start"], "scheduled_end": shift["scheduled_end"],
            "planned_activity": planned, "actual_status": latest["status"] if latest else None, "actual_category": actual,
            "status_start": latest["status_start"] if latest else None, "minutes_in_status": minutes, "rta_result": result,
            "severity": severity, "freshness": freshness, "source_file": latest["source_file"] if latest else None,
        })
    output.sort(key=lambda item: ({"High": 0, "Review": 1, "OK": 2}.get(item["severity"], 9), item["agent_name"] or ""))
    conn.execute("DELETE FROM mart.rta_snapshot")
    _insert_dicts(conn, "mart.rta_snapshot", RTA_COLUMNS, output)
    return output


FORECAST_HOUR_COLUMNS = [
    "business_date", "hour_start", "queue_name", "volume_forecast", "fte_forecast",
    "fte_required", "sl_forecast", "sl_required", "aht_forecast_seconds", "source_file",
    "service_scope", "comparison_scope", "mapping_status", "mapping_sha256",
]

INTRADAY_COLUMNS = [
    "business_date", "interval_start", "hour_start", "source_system", "queue",
    "business_partner", "lob", "language", "offered", "answered", "abandoned",
    "short_calls", "answered_20s", "service_level_20s", "abandon_rate", "asa_seconds",
    "aht_seconds", "source_file", "service_scope", "comparison_scope", "designation", "mapping_status",
    "mapping_sha256",
]


def _build_intraday(
    conn: DatabaseConnection, start: date, end: date, mapping: QueueMapping,
) -> tuple[int, int]:
    conn.execute("DELETE FROM mart.forecast_hour")
    forecast_source = _dicts(conn.execute(
        """
        SELECT business_date, interval_start, queue_name, volume_forecast,
               fte_forecast, fte_required, sl_forecast, sl_required,
               aht_forecast_seconds, source_file
        FROM (
            SELECT r.*, f.file_name AS source_file, f.source_path,
                   row_number() OVER (
                       PARTITION BY f.source_path, queue_name, interval_start, interval_minutes
                       ORDER BY f.modified_at DESC NULLS LAST, f.file_name DESC, source_row DESC
                   ) row_rank
            FROM raw.forecast_interval r JOIN meta.source_file f ON f.file_id=r.source_file_id AND f.active AND f.status='SUCCESS'
            WHERE business_date BETWEEN ? AND ?
        ) x WHERE row_rank=1
        """,
        [start, end],
    ))
    forecast_rows: list[dict[str, Any]] = []
    for row in forecast_source:
        mapped = mapping.map_forecast(row["source_file"], row["queue_name"])
        forecast_rows.append({
            **{key: row[key] for key in (
                "business_date", "queue_name", "volume_forecast", "fte_forecast",
                "fte_required", "sl_forecast", "sl_required", "aht_forecast_seconds", "source_file",
            )},
            "hour_start": row["interval_start"].replace(minute=0, second=0, microsecond=0),
            "service_scope": mapped.service_scope, "comparison_scope": mapped.comparison_scope,
            "mapping_status": mapped.status,
            "mapping_sha256": mapping.sha256,
        })
    _insert_dicts(conn, "mart.forecast_hour", FORECAST_HOUR_COLUMNS, forecast_rows)
    conn.execute("DELETE FROM mart.intraday_queue_interval")
    actual_source = _dicts(conn.execute(
        """
        SELECT business_date, interval_start, hour_start, source_system, queue, business_partner, lob, language,
               offered, answered, abandoned, short_calls, answered_20s,
               asa_seconds, aht_seconds, source_file
        FROM (
            SELECT r.*, f.file_name AS source_file,
                   row_number() OVER (
                       PARTITION BY source_system, business_date, interval_time, coalesce(queue,''),
                                    coalesce(business_partner,''), coalesce(lob,''), coalesce(language,'')
                       ORDER BY f.modified_at DESC NULLS LAST, f.file_name DESC, source_row DESC
                   ) row_rank
            FROM raw.queue_actual r JOIN meta.source_file f ON f.file_id=r.source_file_id AND f.active AND f.status='SUCCESS'
            WHERE business_date BETWEEN ? AND ?
        ) x WHERE row_rank=1
        """,
        [start, end],
    ))
    actual_rows: list[dict[str, Any]] = []
    for row in actual_source:
        mapped = mapping.map_actual(row["source_system"], row["queue"], row["business_partner"], row["lob"])
        offered, answered, abandoned = row["offered"], row["answered"], row["abandoned"]
        actual_rows.append({
            **row,
            "service_level_20s": None if not offered else (row["answered_20s"] or 0) / offered,
            "abandon_rate": None if not offered else (abandoned or 0) / offered,
            "service_scope": mapped.service_scope, "comparison_scope": mapped.comparison_scope,
            "designation": mapped.designation,
            "mapping_status": mapped.status, "mapping_sha256": mapping.sha256,
        })
    _insert_dicts(conn, "mart.intraday_queue_interval", INTRADAY_COLUMNS, actual_rows)
    return len(forecast_rows), len(actual_rows)


def _build_pcs(conn: DatabaseConnection, config: Config, start: date, end: date) -> int:
    """Aggregate deduplicated, FTE-scoped call legs to one agent/day."""
    conn.execute("DELETE FROM mart.agent_pcs_day")
    if not config.modules.get("pcs", True):
        return 0
    minimum = config.pcs.minimum_score
    maximum = config.pcs.maximum_score
    scored = sorted(set(config.pcs.scored_questions))
    valid = {
        number: f"CASE WHEN question_{number}_score BETWEEN {minimum:g} AND {maximum:g} THEN 1 ELSE 0 END"
        for number in range(1, 11)
    }
    score_count = " + ".join(valid[number] for number in scored)
    score_sum = " + ".join(
        f"CASE WHEN question_{number}_score BETWEEN {minimum:g} AND {maximum:g} THEN question_{number}_score ELSE 0 END"
        for number in scored
    )
    comments = sorted(set(config.pcs.comment_questions))
    comment_test = " OR ".join(
        f"coalesce(trim(question_{number}), '') <> ''" for number in comments
    ) or "0"
    sql = f"""
        WITH prepared AS (
            SELECT c.*,
                   d.canonical_name, d.team_leader, d.ops_manager,
                   d.lob AS roster_lob, d.market, d.language AS roster_language,
                   d.location,
                   coalesce(c.talk_seconds,0)+coalesce(c.hold_seconds,0)+coalesce(c.wrap_seconds,0) AS handle_seconds,
                   CASE WHEN upper(coalesce(c.call_direction,''))='I'
                              AND coalesce(c.post_call_survey_mode,'')=? THEN 1 ELSE 0 END AS pcs_eligible,
                   ({score_count}) AS valid_score_count,
                   ({score_sum}) AS valid_score_sum,
                   CASE WHEN question_1_score BETWEEN {minimum:g} AND {maximum:g} THEN question_1_score END AS valid_q1,
                   CASE WHEN question_2_score BETWEEN {minimum:g} AND {maximum:g} THEN question_2_score END AS valid_q2,
                   CASE WHEN {comment_test} THEN 1 ELSE 0 END AS has_comment
            FROM core.clean_call_leg c
            LEFT JOIN core.dim_agent d ON d.agent_id=c.agent_id
            WHERE c.business_date BETWEEN ? AND ? AND c.agent_id IS NOT NULL
        ), scored_calls AS (
            SELECT *,
                   CASE WHEN valid_score_count>0 THEN 1.0*valid_score_sum/valid_score_count END AS survey_score
            FROM prepared
        ), aggregated AS (
            SELECT business_date, agent_id,
                   coalesce(max(canonical_name), max(agent_name)) AS agent_name,
                   max(team_leader) AS team_leader, max(ops_manager) AS ops_manager,
                   coalesce(max(roster_lob), max(lob)) AS lob, max(market) AS market,
                   coalesce(max(roster_language), max(language)) AS language,
                   max(location) AS location,
                   count(*) AS call_legs,
                   sum(CASE WHEN handle_seconds>0 THEN 1 ELSE 0 END) AS handled_calls,
                   sum(CASE WHEN upper(coalesce(call_direction,''))='I' THEN 1 ELSE 0 END) AS inbound_calls,
                   sum(CASE WHEN upper(coalesce(call_direction,''))='O' THEN 1 ELSE 0 END) AS outbound_calls,
                   sum(coalesce(talk_seconds,0)) AS talk_seconds,
                   sum(coalesce(hold_seconds,0)) AS hold_seconds,
                   sum(coalesce(wrap_seconds,0)) AS wrap_seconds,
                   sum(handle_seconds) AS handle_seconds,
                   sum(pcs_eligible) AS pcs_enabled_calls,
                   sum(CASE WHEN pcs_eligible=1 AND survey_score IS NOT NULL THEN 1 ELSE 0 END) AS survey_responses,
                   sum(CASE WHEN pcs_eligible=1 AND valid_q1 IS NOT NULL THEN 1 ELSE 0 END) AS q1_response_count,
                   sum(CASE WHEN pcs_eligible=1 THEN coalesce(valid_q1,0) ELSE 0 END) AS q1_score_sum,
                   sum(CASE WHEN pcs_eligible=1 AND valid_q2 IS NOT NULL THEN 1 ELSE 0 END) AS q2_response_count,
                   sum(CASE WHEN pcs_eligible=1 THEN coalesce(valid_q2,0) ELSE 0 END) AS q2_score_sum,
                   sum(CASE WHEN pcs_eligible=1 AND survey_score IS NOT NULL THEN 1 ELSE 0 END) AS pcs_score_count,
                   sum(CASE WHEN pcs_eligible=1 THEN coalesce(survey_score,0) ELSE 0 END) AS pcs_score_sum,
                   sum(CASE WHEN pcs_eligible=1 AND survey_score >= {config.pcs.top_box_minimum:g} THEN 1 ELSE 0 END) AS top_box_responses,
                   sum(CASE WHEN pcs_eligible=1 AND survey_score <= {config.pcs.low_score_maximum:g} THEN 1 ELSE 0 END) AS low_score_responses,
                   sum(CASE WHEN pcs_eligible=1 THEN has_comment ELSE 0 END) AS comments_count
            FROM scored_calls
            GROUP BY business_date, agent_id
        )
        INSERT INTO mart.agent_pcs_day (
            agent_day_key, business_date, agent_id, agent_name, team_leader,
            ops_manager, lob, market, language, location, call_legs,
            handled_calls, inbound_calls, outbound_calls, talk_seconds,
            hold_seconds, wrap_seconds, handle_seconds, average_talk_seconds,
            average_hold_seconds, average_wrap_seconds, average_handle_seconds,
            pcs_enabled_calls, survey_responses, response_rate,
            q1_response_count, q1_score_sum, q1_average,
            q2_response_count, q2_score_sum, q2_average,
            pcs_score_count, pcs_score_sum, pcs_average,
            top_box_responses, low_score_responses, top_box_percent,
            low_score_percent, comments_count
        )
        SELECT replace(business_date,'-','') || '-' || agent_id,
               business_date, agent_id, agent_name, team_leader, ops_manager,
               lob, market, language, location, call_legs, handled_calls,
               inbound_calls, outbound_calls, talk_seconds, hold_seconds,
               wrap_seconds, handle_seconds,
               CASE WHEN handled_calls>0 THEN 1.0*talk_seconds/handled_calls END,
               CASE WHEN handled_calls>0 THEN 1.0*hold_seconds/handled_calls END,
               CASE WHEN handled_calls>0 THEN 1.0*wrap_seconds/handled_calls END,
               CASE WHEN handled_calls>0 THEN 1.0*handle_seconds/handled_calls END,
               pcs_enabled_calls, survey_responses,
               CASE WHEN pcs_enabled_calls>0 THEN 1.0*survey_responses/pcs_enabled_calls END,
               q1_response_count, q1_score_sum,
               CASE WHEN q1_response_count>0 THEN 1.0*q1_score_sum/q1_response_count END,
               q2_response_count, q2_score_sum,
               CASE WHEN q2_response_count>0 THEN 1.0*q2_score_sum/q2_response_count END,
               pcs_score_count, pcs_score_sum,
               CASE WHEN pcs_score_count>0 THEN 1.0*pcs_score_sum/pcs_score_count END,
               top_box_responses, low_score_responses,
               CASE WHEN survey_responses>0 THEN 1.0*top_box_responses/survey_responses END,
               CASE WHEN survey_responses>0 THEN 1.0*low_score_responses/survey_responses END,
               comments_count
        FROM aggregated
    """
    conn.execute(sql, [config.pcs.survey_mode, start, end])
    return conn.execute("SELECT count(*) FROM mart.agent_pcs_day").fetchone()[0]


ABSENCE_EVENT_COLUMNS = [
    "event_key", "agent_day_key", "business_date", "agent_id", "agent_name",
    "team_leader", "ops_manager", "lob", "market", "language", "location",
    "activity", "category", "event_start", "event_end", "minutes", "hours",
    "planned", "working", "counts_as_absence", "counts_as_vacation",
    "counts_as_unpaid", "counts_as_shrinkage", "mapped", "evidence_type",
    "source_file", "rule_version", "rule_sha256",
    "reconciliation_status", "verint_activity", "verint_category",
    "verint_overlap_minutes", "verint_source_file",
]

ABSENCE_DAY_COLUMNS = [
    "agent_day_key", "business_date", "agent_id", "agent_name", "team_leader",
    "ops_manager", "lob", "market", "language", "location", "scheduled_minutes",
    "break_minutes", "lunch_minutes", "planned_net_minutes", "production_minutes",
    "absence_minutes", "vacation_minutes", "unpaid_minutes", "shrinkage_minutes",
    "late_minutes", "early_leave_minutes", "no_show_minutes", "unmapped_minutes",
    "absence_rate", "vacation_rate", "shrinkage_rate", "absence_day",
    "absence_spell", "absence_spells", "absence_days", "bradford_factor",
    "rule_version", "rule_sha256",
    "unverified_minutes", "corrected_minutes",
]


def _rule_flags(rule, fallback_category: str, **fallback: bool) -> dict[str, Any]:
    if rule is not None:
        return {
            "category": rule.category, "planned": rule.planned, "working": rule.working,
            "absence": rule.absence, "vacation": rule.vacation, "unpaid": rule.unpaid,
            "shrinkage": rule.shrinkage, "mapped": True,
        }
    return {
        "category": fallback_category, "planned": fallback.get("planned", False),
        "working": fallback.get("working", False), "absence": fallback.get("absence", False),
        "vacation": fallback.get("vacation", False), "unpaid": fallback.get("unpaid", False),
        "shrinkage": fallback.get("shrinkage", False), "mapped": False,
    }


def _build_absence(
    conn: DatabaseConnection,
    config: Config,
    rulebook: Rulebook,
    attendance: list[dict[str, Any]],
    corrections: list[dict[str, Any]],
) -> tuple[int, int]:
    """Build absence only from observed LILO/status gaps, then attach Verint final labels."""
    conn.execute("DELETE FROM mart.absence_event")
    conn.execute("DELETE FROM mart.absence_agent_day")
    if not config.modules.get("absence", True):
        return 0, 0
    event_rows: list[dict[str, Any]] = []
    day_rows: list[dict[str, Any]] = []
    standard_day_minutes = max(1, int(round(rulebook.standard_day_hours * 60)))
    corrections_by_key: dict[str, list[dict[str, Any]]] = defaultdict(list)
    for correction in corrections:
        corrections_by_key[f"{correction['business_date']:%Y%m%d}-{correction['agent_id']}"].append(correction)

    for base in attendance:
        shift_start, shift_end = base["scheduled_start"], base["scheduled_end"]
        scheduled_minutes = base["scheduled_minutes"] if shift_start and shift_end else 0
        intervals: dict[str, list[tuple[datetime, datetime]]] = defaultdict(list)

        issue_intervals: dict[str, list[tuple[datetime, datetime]]] = defaultdict(list)
        corrected_overlap_minutes = 0

        def add_event(correction: dict[str, Any]) -> None:
            nonlocal corrected_overlap_minutes
            event_start, event_end = correction["gap_start"], correction["gap_end"]
            if not shift_start or not shift_end or not event_start or not event_end:
                return
            if rulebook.cap_event_to_schedule:
                event_start, event_end = max(event_start, shift_start), min(event_end, shift_end)
            if event_end <= event_start:
                return
            activity = correction.get("verint_activity") or correction.get("suggested_activity") or correction["detected_issue"]
            final_rule = rulebook.classify_activity(correction.get("verint_activity")) if correction.get("verint_activity") else None
            fallback_category = {
                "No show": "NO_SHOW", "Late": "LATE", "Early leave": "EARLY_LEAVE",
                "Mid-shift logged off": "STATUS_LOGGED_OFF",
                "Unavailable in shift": "STATUS_UNAVAILABLE",
            }.get(correction["detected_issue"], "UNEXPLAINED_ABSENCE")
            flags = _rule_flags(
                final_rule, fallback_category,
                absence=True, unpaid=False, shrinkage=True, working=False, planned=False,
            )
            minutes = int((event_end - event_start).total_seconds() // 60)
            if minutes <= 0:
                return
            event_key = hashlib.sha256(
                f"{base['agent_day_key']}|{correction['correction_id']}|{event_start}|{event_end}".encode("utf-8")
            ).hexdigest()
            event_rows.append({
                "event_key": event_key, "agent_day_key": base["agent_day_key"],
                "business_date": base["business_date"], "agent_id": base["agent_id"],
                "agent_name": base["agent_name"], "team_leader": base["team_leader"],
                "ops_manager": base["ops_manager"], "lob": base["lob"],
                "market": base["market"], "language": base["language"], "location": base["location"],
                "activity": activity, "category": flags["category"], "event_start": event_start,
                "event_end": event_end, "minutes": minutes, "hours": minutes / 60.0,
                "planned": flags["planned"], "working": flags["working"],
                "counts_as_absence": flags["absence"], "counts_as_vacation": flags["vacation"],
                "counts_as_unpaid": flags["unpaid"], "counts_as_shrinkage": flags["shrinkage"],
                "mapped": flags["mapped"], "evidence_type": correction["observed_source"],
                "source_file": correction["source_file"], "rule_version": rulebook.version,
                "rule_sha256": rulebook.sha256,
                "reconciliation_status": correction["verint_reconciliation"],
                "verint_activity": correction.get("verint_activity"),
                "verint_category": correction.get("verint_category"),
                "verint_overlap_minutes": correction.get("verint_overlap_minutes") or 0,
                "verint_source_file": correction.get("verint_source_file"),
            })
            intervals["all"].append((event_start, event_end))
            intervals[flags["category"]].append((event_start, event_end))
            if not flags["working"]:
                intervals["non_working"].append((event_start, event_end))
            for key in ("absence", "vacation", "unpaid", "shrinkage"):
                if flags[key]:
                    intervals[key].append((event_start, event_end))
            if not flags["mapped"]:
                intervals["unmapped"].append((event_start, event_end))
            issue_intervals[correction["detected_issue"]].append((event_start, event_end))
            corrected_overlap_minutes += min(minutes, int(correction.get("verint_overlap_minutes") or 0))

        for correction in corrections_by_key.get(base["agent_day_key"], []):
            add_event(correction)

        def minutes_for(key: str) -> int:
            if not shift_start or not shift_end:
                return 0
            return interval_minutes(shift_start, shift_end, intervals.get(key, []))

        status_exclusive = base.get("_status_exclusive", [])
        lunch = interval_minutes(shift_start, shift_end, [
            (row["interval_start"], row["interval_end"])
            for row in status_exclusive if row["actual_category"] == "Lunch"
        ]) if shift_start and shift_end else 0
        breaks = interval_minutes(shift_start, shift_end, [
            (row["interval_start"], row["interval_end"])
            for row in status_exclusive if row["actual_category"] == "Break"
        ]) if shift_start and shift_end else 0
        # Standard-day hours are already net. Actual lunch/break statuses are
        # descriptive evidence and are not subtracted from the denominator.
        planned_net = 0 if base["assignment_type"] == "Off" else min(standard_day_minutes, scheduled_minutes)
        absence = min(planned_net, minutes_for("absence"))
        vacation = min(planned_net, minutes_for("vacation"))
        unpaid = min(planned_net, minutes_for("unpaid"))
        shrinkage = min(planned_net, minutes_for("shrinkage"))
        production = max(0, planned_net - min(planned_net, minutes_for("all")))
        corrected = min(minutes_for("all"), corrected_overlap_minutes)
        unverified = max(0, minutes_for("all") - corrected)
        values = {
            "absence_hours": absence / 60.0, "vacation_hours": vacation / 60.0,
            "shrinkage_hours": shrinkage / 60.0, "planned_net_hours": planned_net / 60.0,
        }
        day_rows.append({
            **{key: base[key] for key in (
                "agent_day_key", "business_date", "agent_id", "agent_name", "team_leader",
                "ops_manager", "lob", "market", "language", "location",
            )},
            "scheduled_minutes": scheduled_minutes, "break_minutes": breaks,
            "lunch_minutes": lunch, "planned_net_minutes": planned_net,
            "production_minutes": production, "absence_minutes": absence,
            "vacation_minutes": vacation, "unpaid_minutes": unpaid,
            "shrinkage_minutes": shrinkage,
            "late_minutes": interval_minutes(shift_start, shift_end, issue_intervals["Late"]) if shift_start and shift_end else 0,
            "early_leave_minutes": interval_minutes(shift_start, shift_end, issue_intervals["Early leave"]) if shift_start and shift_end else 0,
            "no_show_minutes": interval_minutes(shift_start, shift_end, issue_intervals["No show"]) if shift_start and shift_end else 0,
            "unmapped_minutes": unverified,
            "absence_rate": evaluate_formula(rulebook.formulas["absence_rate"].formula, values),
            "vacation_rate": evaluate_formula(rulebook.formulas["vacation_rate"].formula, values),
            "shrinkage_rate": evaluate_formula(rulebook.formulas["shrinkage_rate"].formula, values),
            "absence_day": absence > 0, "absence_spell": None, "absence_spells": 0,
            "absence_days": 0.0, "bradford_factor": 0.0,
            "rule_version": rulebook.version, "rule_sha256": rulebook.sha256,
            "unverified_minutes": unverified, "corrected_minutes": corrected,
        })

    by_agent: dict[str, list[dict[str, Any]]] = defaultdict(list)
    for row in day_rows:
        by_agent[row["agent_id"]].append(row)
    for agent_id, rows in by_agent.items():
        absent = sorted((row for row in rows if row["absence_day"]), key=lambda item: item["business_date"])
        spell_count = 0
        previous = None
        for row in absent:
            if previous is None or (row["business_date"] - previous).days > rulebook.spell_gap_days:
                spell_count += 1
            row["absence_spell"] = f"{agent_id}-{spell_count:03d}"
            previous = row["business_date"]
        absence_days = sum(row["absence_minutes"] for row in rows) / standard_day_minutes
        bradford = float(spell_count * spell_count * absence_days)
        for row in rows:
            row["absence_spells"] = spell_count
            row["absence_days"] = absence_days
            row["bradford_factor"] = bradford

    _insert_dicts(conn, "mart.absence_event", ABSENCE_EVENT_COLUMNS, event_rows)
    _insert_dicts(conn, "mart.absence_agent_day", ABSENCE_DAY_COLUMNS, day_rows)
    return len(day_rows), len(event_rows)


SERVICE_COLUMNS = [
    "business_date", "interval_start", "hour_start", "source_system", "queue",
    "business_partner", "lob", "language", "offered", "answered", "abandoned",
    "short_abandoned", "answered_within_target", "handled_seconds", "sl_gross",
    "sl_adjusted", "sl_profile", "service_level", "service_availability",
    "abandon_rate", "aht_seconds", "source_file", "rule_version", "rule_sha256",
    "service_scope", "comparison_scope", "designation", "mapping_status", "mapping_sha256",
]


def _build_service(
    conn: DatabaseConnection, rulebook: Rulebook, mapping: QueueMapping, start: date, end: date,
) -> int:
    conn.execute("DELETE FROM mart.service_interval")
    rows = _dicts(conn.execute(
        """
        SELECT business_date, interval_start, hour_start, source_system, queue,
               business_partner, lob, language, offered, answered, abandoned,
               short_calls, abandoned_20s, answered_20s, aht_seconds, source_file
        FROM (
            SELECT r.*, f.file_name AS source_file,
                   row_number() OVER (
                       PARTITION BY source_system, business_date, interval_time, coalesce(queue,''),
                                    coalesce(business_partner,''), coalesce(lob,''), coalesce(language,'')
                       ORDER BY f.modified_at DESC NULLS LAST, f.file_name DESC, source_row DESC
                   ) row_rank
            FROM raw.queue_actual r
            JOIN meta.source_file f ON f.file_id=r.source_file_id AND f.active AND f.status='SUCCESS'
            WHERE business_date BETWEEN ? AND ?
        ) x WHERE row_rank=1
        """, [start, end],
    ))
    output: list[dict[str, Any]] = []
    gross = rulebook.service_profiles["gross_20"]
    adjusted = rulebook.service_profiles["adjusted_20"]
    for row in rows:
        short_abandoned = row["abandoned_20s"] if row["abandoned_20s"] is not None else row["short_calls"]
        handled_seconds = (
            float(row["aht_seconds"]) * float(row["answered"])
            if row["aht_seconds"] is not None and row["answered"] is not None else None
        )
        values = {
            "offered": row["offered"], "answered": row["answered"],
            "abandoned": row["abandoned"], "short_abandoned": short_abandoned,
            "answered_within_target": row["answered_20s"], "handled_seconds": handled_seconds,
        }
        profile = rulebook.service_profile_for(row["source_system"], row["lob"], row["language"])
        mapped = mapping.map_actual(row["source_system"], row["queue"], row["business_partner"], row["lob"])
        output.append({
            **{key: row[key] for key in (
                "business_date", "interval_start", "hour_start", "source_system", "queue",
                "business_partner", "lob", "language", "offered", "answered", "abandoned",
                "source_file",
            )},
            "short_abandoned": short_abandoned, "answered_within_target": row["answered_20s"],
            "handled_seconds": handled_seconds,
            "sl_gross": evaluate_formula(gross.formula, values),
            "sl_adjusted": evaluate_formula(adjusted.formula, values),
            "sl_profile": profile.key, "service_level": evaluate_formula(profile.formula, values),
            "service_availability": evaluate_formula(rulebook.formulas["service_availability"].formula, values),
            "abandon_rate": evaluate_formula(rulebook.formulas["abandon_rate"].formula, values),
            "aht_seconds": evaluate_formula(rulebook.formulas["aht_seconds"].formula, values),
            "rule_version": rulebook.version, "rule_sha256": rulebook.sha256,
            "service_scope": mapped.service_scope, "comparison_scope": mapped.comparison_scope,
            "designation": mapped.designation,
            "mapping_status": mapped.status, "mapping_sha256": mapping.sha256,
        })
    _insert_dicts(conn, "mart.service_interval", SERVICE_COLUMNS, output)
    return len(output)


def _record_rule_application(conn: DatabaseConnection, run_id: str, rulebook: Rulebook) -> None:
    conn.execute(
        """INSERT INTO meta.rule_application(
               run_id, rule_version, rule_sha256, rule_file, effective_from, applied_at
           ) VALUES (?, ?, ?, ?, ?, ?)
           ON CONFLICT(run_id) DO UPDATE SET
               rule_version=excluded.rule_version, rule_sha256=excluded.rule_sha256,
               rule_file=excluded.rule_file, effective_from=excluded.effective_from,
               applied_at=excluded.applied_at""",
        [run_id, rulebook.version, rulebook.sha256, str(rulebook.file), rulebook.effective_from, datetime.now()],
    )


def _record_mapping_application(conn: DatabaseConnection, run_id: str, mapping: QueueMapping) -> None:
    conn.execute(
        """INSERT INTO meta.mapping_application(run_id, mapping_sha256, mapping_file, applied_at)
           VALUES (?, ?, ?, ?)
           ON CONFLICT(run_id) DO UPDATE SET
               mapping_sha256=excluded.mapping_sha256,
               mapping_file=excluded.mapping_file,
               applied_at=excluded.applied_at""",
        [run_id, mapping.sha256, str(mapping.file), datetime.now()],
    )


def _issue_id(*parts: Any) -> str:
    return hashlib.sha256("|".join(str(part or "") for part in parts).encode("utf-8")).hexdigest()


def _build_quality(
    conn: DatabaseConnection,
    config: Config,
    run_id: str,
    start: date,
    end: date,
) -> int:
    conn.execute("DELETE FROM meta.quality_issue")
    now = datetime.now()
    # A single underlying problem can be discovered through more than one
    # metadata row (for example, repeated failed attempts for the same file).
    # Keep one row per deterministic issue_id before touching SQLite so one
    # noisy source can never abort the whole model refresh.
    issues: dict[str, list[Any]] = {}

    def add(family: str | None, source_file: str | None, business_date: date | None, agent_id: str | None, issue: str, severity: str, details: str) -> None:
        issue_id = _issue_id(run_id, family, source_file, business_date, agent_id, issue, details)
        candidate = [issue_id, run_id, now, family, source_file, business_date, agent_id, issue, severity, details]
        existing = issues.get(issue_id)
        severity_rank = {"INFO": 0, "REVIEW": 1, "ERROR": 2}
        if existing is None or severity_rank.get(severity, 0) > severity_rank.get(existing[8], 0):
            issues[issue_id] = candidate

    for family, key in (
        ("fte", "fte_file"), ("schedule", "schedule_folder"), ("lilo", "lilo_folder"),
        ("agent_status", "agent_status_folder"), ("forecast", "forecast_folder"),
        ("apbe", "apbe_folder"), ("apfr", "apfr_folder"), ("apde", "apde_folder"), ("calls", "call_folder"),
    ):
        if family == "agent_status" and not config.modules.get("agent_status", True):
            continue
        if family == "forecast" and not config.modules.get("forecast", True):
            continue
        if family in {"apbe", "apfr", "apde"} and not config.modules.get("intraday", True):
            continue
        if family == "calls" and not config.modules.get("pcs", True):
            continue
        path = config.source_path(key)
        if not path.exists():
            add(family, str(path), None, None, "Missing source", "ERROR", f"Expected path does not exist: {path}")
    for row in _dicts(conn.execute("SELECT * FROM meta.source_file WHERE status='ERROR'")):
        if row["source_family"] == "agent_status" and not config.modules.get("agent_status", False):
            continue
        add(row["source_family"], row["file_name"], None, None, "Source load error", "ERROR", row["error_message"] or "Unknown load error")
    rulebook = load_rulebook(config.home, config.business_rules)
    for row in _dicts(conn.execute(
        """SELECT schedule_date, agent_id_raw, agent_name, parse_ok, f.file_name
           FROM raw.schedule_shift r JOIN meta.source_file f ON f.file_id=r.source_file_id AND f.active
           WHERE schedule_date BETWEEN ? AND ? AND (agent_id IS NULL OR NOT parse_ok)""", [start, end]
    )):
        issue = "Invalid schedule Agent ID" if row["agent_id_raw"] in {None, "", "-", "N/A", "NA", "NULL"} else "Schedule parse error"
        add("schedule", row["file_name"], row["schedule_date"], row["agent_id_raw"], issue, "ERROR", f"Agent={row['agent_name'] or ''}; parse_ok={row['parse_ok']}")
    for business_date, count in conn.execute(
        """SELECT business_date, count(*) FROM mart.attendance_agent_day
           WHERE attendance_result='Data not loaded' GROUP BY business_date ORDER BY business_date"""
    ).fetchall():
        add("attendance", None, business_date, None, "Actual evidence date not loaded", "ERROR", f"{count} scheduled Agent ID rows cannot be judged from LILO or Agent Status")
    for row in _dicts(conn.execute(
        """SELECT business_date, agent_id, attendance_result, schedule_source
           FROM mart.attendance_agent_day
           WHERE attendance_result IN ('Schedule parse error','Missing actual evidence','Incomplete actual evidence','No schedule overlap')"""
    )):
        severity = "ERROR" if row["attendance_result"] in {"Data not loaded", "Schedule parse error"} else "REVIEW"
        add("attendance", row["schedule_source"], row["business_date"], row["agent_id"], row["attendance_result"], severity, "Attendance result requires review before payroll use")
    if config.modules.get("agent_status", True):
        for business_date, low_rows in conn.execute(
            """SELECT business_date, count(*)
               FROM mart.attendance_agent_day
               WHERE scheduled_minutes>0 AND status_source IS NOT NULL
                 AND 1.0*status_covered_minutes/scheduled_minutes < ?
               GROUP BY business_date ORDER BY business_date""",
            [config.rules.minimum_status_coverage],
        ).fetchall():
            add(
                "agent_status", None, business_date, None, "Low Agent Status coverage", "REVIEW",
                f"{low_rows} agent-day rows have less than {config.rules.minimum_status_coverage:.0%} status coverage; LILO boundaries still remain usable.",
            )
    forecast_unmapped = conn.execute("SELECT count(*) FROM mart.forecast_hour WHERE mapping_status='UNMAPPED'").fetchone()[0]
    actual_unmapped = conn.execute("SELECT count(*) FROM mart.intraday_queue_interval WHERE mapping_status='UNMAPPED'").fetchone()[0]
    if forecast_unmapped or actual_unmapped:
        add(
            "intraday", str(config.queue_mapping), None, None, "Unmapped service scope", "REVIEW",
            f"{forecast_unmapped} forecast and {actual_unmapped} actual rows are unmapped. Edit config/queue_mapping.csv and refresh; extracts stay untouched.",
        )
    if config.modules.get("pcs", True):
        call_rows = conn.execute("SELECT count(*) FROM core.clean_call_leg WHERE business_date BETWEEN ? AND ?", [start, end]).fetchone()[0]
        responses = conn.execute("SELECT coalesce(sum(survey_responses),0) FROM mart.agent_pcs_day").fetchone()[0]
        if call_rows and not responses:
            add(
                "calls", None, None, None, "No in-scope PCS responses", "REVIEW",
                f"{call_rows} clean in-scope call legs were available, but no valid configured survey score was found.",
            )
    for row in _dicts(conn.execute(
        """SELECT business_date, agent_id, activity, category, source_file, sum(minutes) AS minutes
           FROM mart.absence_event
           WHERE mapped=false AND business_date BETWEEN ? AND ?
           GROUP BY business_date, agent_id, activity, category, source_file
           ORDER BY business_date, agent_id, activity""", [start, end]
    )):
        severity = "REVIEW"
        add(
            "absence", row["source_file"], row["business_date"], row["agent_id"],
            "Observed gap not verified in Verint", severity,
            f"{row['activity']} -> {row['category']} ({row['minutes']} minutes). Correct it in Verint, then export Activities and refresh.",
        )
    for row in _dicts(conn.execute(
        """SELECT business_date, agent_id, agent_name, activity, category, minutes, source_file
           FROM mart.verint_final_exception WHERE business_date BETWEEN ? AND ?""", [start, end]
    )):
        add(
            "schedule", row["source_file"], row["business_date"], row["agent_id"],
            "Verint final activity without observed gap", "REVIEW",
            f"{row['activity']} / {row['category']} covers {row['minutes']} minutes for {row['agent_name'] or row['agent_id']}; compare LILO and Agent Status.",
        )
    for row in _dicts(conn.execute(
        """SELECT business_date, source_system, queue, offered, answered, source_file
           FROM mart.service_interval
           WHERE business_date BETWEEN ? AND ? AND coalesce(answered,0) > coalesce(offered,0)""", [start, end]
    )):
        add(
            "service", row["source_file"], row["business_date"], None,
            "Answered exceeds offered", "REVIEW",
            f"{row['source_system']} / {row['queue'] or '(no queue)'}: answered={row['answered']}, offered={row['offered']}",
        )
    for row in _dicts(conn.execute(
        """SELECT business_date, agent_id, activity, source_file, sum(minutes) AS minutes
           FROM mart.absence_event
           WHERE category='NO_ACTIVITY' AND business_date BETWEEN ? AND ?
           GROUP BY business_date, agent_id, activity, source_file""", [start, end]
    )):
        add(
            "absence", row["source_file"], row["business_date"], row["agent_id"],
            "Verint No Activity", "REVIEW",
            f"No Activity covers {row['minutes']} scheduled minutes. Confirm whether this is shrinkage or a schedule defect.",
        )
    for row in _dicts(conn.execute(
        """SELECT DISTINCT a.business_date, a.agent_id, a.agent_name,
                          a.category AS category_a, b.category AS category_b,
                          coalesce(a.source_file, b.source_file) AS source_file
           FROM mart.absence_event a
           JOIN mart.absence_event b
             ON b.agent_day_key=a.agent_day_key AND b.event_key>a.event_key
            AND b.event_start<a.event_end AND b.event_end>a.event_start
           WHERE a.business_date BETWEEN ? AND ?
             AND a.counts_as_absence=true AND b.counts_as_absence=true
             AND a.category<>b.category""", [start, end]
    )):
        add(
            "absence", row["source_file"], row["business_date"], row["agent_id"],
            "Conflicting absence evidence", "REVIEW",
            f"Overlapping categories {row['category_a']} and {row['category_b']} for {row['agent_name'] or row['agent_id']}. Daily absence is unioned, but payroll classification needs review.",
        )
    for family, file_name, scoped_out in conn.execute(
        """SELECT source_family, file_name, scoped_out_count
           FROM meta.source_file
           WHERE active=true AND status='SUCCESS' AND source_family IN ('schedule','lilo','agent_status','calls')
             AND row_count=0 AND scoped_out_count>0"""
    ).fetchall():
        add(
            family, file_name, None, None, "Agent scope mismatch", "ERROR",
            f"All {scoped_out} source rows were outside the active FTE roster; no rows were used.",
        )
    issue_rows = list(issues.values())
    if issue_rows:
        row_placeholders = "(" + ", ".join("?" for _ in range(10)) + ")"
        for offset in range(0, len(issue_rows), 500):
            batch = issue_rows[offset : offset + 500]
            conn.execute(
                "INSERT INTO meta.quality_issue VALUES " + ", ".join(row_placeholders for _ in batch),
                [value for row in batch for value in row],
            )
    return len(issue_rows)


def _build_source_health(conn: DatabaseConnection, config: Config) -> None:
    conn.execute("DELETE FROM mart.source_health")
    specs = [
        ("fte", config.source_path("fte_file")), ("schedule", config.source_path("schedule_folder")),
        ("lilo", config.source_path("lilo_folder")), ("agent_status", config.source_path("agent_status_folder")),
        ("forecast", config.source_path("forecast_folder")), ("apbe", config.source_path("apbe_folder")),
        ("apfr", config.source_path("apfr_folder")), ("apde", config.source_path("apde_folder")),
        ("calls", config.source_path("call_folder")),
    ]
    specs = [
        (family, path) for family, path in specs
        if not (family == "agent_status" and not config.modules.get("agent_status", True))
        and not (family == "forecast" and not config.modules.get("forecast", True))
        and not (family in {"apbe", "apfr", "apde"} and not config.modules.get("intraday", True))
        and not (family == "calls" and not config.modules.get("pcs", True))
    ]
    for family, expected in specs:
        latest = conn.execute(
            """SELECT file_name, modified_at, loaded_at, status, error_message
               FROM meta.source_file WHERE source_family=?
               ORDER BY loaded_at DESC NULLS LAST, modified_at DESC NULLS LAST LIMIT 1""", [family]
        ).fetchone()
        rows, rejected, scoped_out = conn.execute(
            """SELECT coalesce(sum(row_count),0), coalesce(sum(rejected_count),0),
                      coalesce(sum(scoped_out_count),0)
               FROM meta.source_file
               WHERE source_family=? AND active=true AND status='SUCCESS'""",
            [family],
        ).fetchone()
        business_date = conn.execute(
            """
            SELECT max(d) FROM (
                SELECT max(extract_date) d FROM raw.lilo r JOIN meta.source_file f ON f.file_id=r.source_file_id AND f.active WHERE ?='lilo'
                UNION ALL SELECT max(extract_date) FROM raw.agent_status r JOIN meta.source_file f ON f.file_id=r.source_file_id AND f.active WHERE ?='agent_status'
                UNION ALL SELECT max(schedule_date) FROM raw.schedule_shift r JOIN meta.source_file f ON f.file_id=r.source_file_id AND f.active WHERE ?='schedule'
                UNION ALL SELECT max(business_date) FROM raw.forecast_interval r JOIN meta.source_file f ON f.file_id=r.source_file_id AND f.active WHERE ?='forecast'
                UNION ALL SELECT max(business_date) FROM raw.queue_actual r JOIN meta.source_file f ON f.file_id=r.source_file_id AND f.active WHERE source_system=upper(?)
                UNION ALL SELECT max(business_date) FROM raw.call_leg r JOIN meta.source_file f ON f.file_id=r.source_file_id AND f.active WHERE ?='calls'
            ) x
            """, [family, family, family, family, family, family]
        ).fetchone()[0]
        if latest:
            file_name, modified, loaded, status, error = latest
            if error:
                details = error
            elif not rows and scoped_out:
                status = "ERROR"
                details = f"All {scoped_out} rows were outside the active FTE roster; no rows were used"
            else:
                notes = []
                if rejected:
                    notes.append(f"{rejected} rejected/flagged")
                if scoped_out:
                    notes.append(f"{scoped_out} outside roster excluded")
                details = "Loaded successfully" + (f"; {', '.join(notes)} rows" if notes else "")
        else:
            file_name = modified = loaded = business_date = None
            status = "MISSING" if not expected.exists() else "EMPTY"
            details = "Path not found" if status == "MISSING" else "No matching files loaded"
        conn.execute(
            "INSERT INTO mart.source_health VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
            [family, str(expected), file_name, business_date, modified, loaded, rows, rejected, status, details, scoped_out],
        )


def refresh_models(
    conn: DatabaseConnection,
    config: Config,
    run_id: str,
    start: date | None = None,
    end: date | None = None,
    use_config_period: bool = True,
    progress: ProgressCallback | None = None,
) -> ModelSummary:
    total_stages = 17

    def stage(completed: int, label: str) -> None:
        if progress is not None:
            progress(completed, total_stages, label)

    conn.execute("SAVEPOINT refresh_models")
    try:
        stage(0, "Validating business rules")
        rulebook = load_rulebook(config.home, config.business_rules)
        mapping = load_queue_mapping(config.queue_mapping)
        stage(1, "Selecting reporting period")
        start, end = resolve_period(conn, config, start, end, use_config_period)
        stage(2, "Loading schedules")
        schedules = _load_schedules(conn, start, end)
        stage(3, "Loading schedule activities")
        events = _load_events(conn, start, end)
        events_by_agent = _events_by_agent(events)
        stage(4, "Loading LILO")
        lilo, loaded_dates, seen_ids = _load_lilo(conn, start, end)
        stage(5, "Loading Agent Status attendance evidence")
        if config.modules.get("agent_status", True):
            statuses, status_loaded_dates = _load_statuses(conn, start, end)
        else:
            statuses, status_loaded_dates = {}, set()
        stage(6, "Building employee dimension")
        agents = _build_agents(conn)
        stage(7, "Building attendance")
        attendance = _build_attendance(
            conn, rulebook, schedules, events_by_agent, lilo, loaded_dates, seen_ids,
            agents, statuses, status_loaded_dates,
        )
        stage(8, "Keeping adherence disabled")
        conn.execute("DELETE FROM mart.conformance_agent_day")
        conformance = []
        stage(9, "Finding observed LILO and status gaps")
        corrections = _build_corrections(conn, rulebook, attendance)
        stage(10, "Keeping legacy RTA disabled")
        conn.execute("DELETE FROM mart.rta_snapshot")
        rta = []
        stage(11, "Building intraday actual and forecast")
        forecast, actual = _build_intraday(conn, start, end, mapping)
        stage(12, "Building Agent PCS")
        pcs = _build_pcs(conn, config, start, end)
        stage(13, "Building absence and shrinkage")
        absence, absence_events = _build_absence(conn, config, rulebook, attendance, corrections)
        stage(14, "Building service performance")
        service = _build_service(conn, rulebook, mapping, start, end)
        _record_rule_application(conn, run_id, rulebook)
        _record_mapping_application(conn, run_id, mapping)
        stage(15, "Checking source health")
        _build_source_health(conn, config)
        stage(16, "Running data-quality checks")
        quality = _build_quality(conn, config, run_id, start, end)
        result = ModelSummary(
            start=start, end=end, attendance_rows=len(attendance), conformance_rows=len(conformance),
            correction_rows=len(corrections), rta_rows=len(rta), forecast_rows=forecast,
            intraday_rows=actual, pcs_rows=pcs, quality_rows=quality,
            absence_rows=absence, absence_event_rows=absence_events, service_rows=service,
        )
        conn.execute("RELEASE SAVEPOINT refresh_models")
        stage(total_stages, "Models ready")
        return result
    except Exception:
        conn.execute("ROLLBACK TO SAVEPOINT refresh_models")
        conn.execute("RELEASE SAVEPOINT refresh_models")
        raise
