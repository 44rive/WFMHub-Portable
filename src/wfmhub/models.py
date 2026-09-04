"""Materialize clean dimensions and WFM report marts."""

from __future__ import annotations

import hashlib
from collections import defaultdict
from dataclasses import dataclass
from datetime import date, datetime, time, timedelta
from pathlib import Path
from typing import Any, Iterable
from zoneinfo import ZoneInfo, ZoneInfoNotFoundError

from .analytics import build_findings, load_analytics_rules, validate_analytics_rules
from .config import Config
from .database import DatabaseConnection, DatabaseCursor
from .mapping import QueueMapping, load_queue_mapping
from .metrics import MetricCatalog, MetricEvaluation, evaluate_metric, load_metric_catalog, validate_metric_catalog
from .progress import ProgressCallback
from .rules import Rulebook, load_rulebook
from .semantic import SOURCE_COMPONENTS, build_metric_values
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
    staffing_rows: int = 0
    timeline_rows: int = 0
    final_absence_rows: int = 0
    final_absence_event_rows: int = 0
    metric_rows: int = 0
    finding_rows: int = 0


def _evaluation_time(timezone_name: str, as_of: datetime | None) -> datetime:
    """Return one timezone-local naive cutoff, even on stripped Python builds."""
    try:
        local_zone = ZoneInfo(timezone_name)
    except ZoneInfoNotFoundError:
        # The portable package ships tzdata, but a damaged/incomplete copy must
        # still run using the Windows-configured local clock instead of aborting.
        if as_of is None:
            return datetime.now()
        return as_of.astimezone().replace(tzinfo=None) if as_of.tzinfo is not None else as_of
    if as_of is None:
        return datetime.now(local_zone).replace(tzinfo=None)
    if as_of.tzinfo is not None:
        return as_of.astimezone(local_zone).replace(tzinfo=None)
    return as_of


def _dicts(cursor: DatabaseCursor) -> list[dict[str, Any]]:
    columns = [item[0] for item in cursor.description]
    return [dict(zip(columns, row)) for row in cursor.fetchall()]


def _metric_evaluation(
    catalog: MetricCatalog,
    metric_id: str,
    business_date: date,
    dimensions: dict[str, Any],
    components: dict[str, Any],
) -> MetricEvaluation:
    method = catalog.method_for(metric_id, business_date, dimensions)
    if method is None:
        raise ValueError(
            f"No effective metric method for {metric_id!r} on {business_date} "
            f"with dimensions {dimensions}"
        )
    return evaluate_metric(method, components)


def _optional_metric_value(
    catalog: MetricCatalog,
    metric_id: str,
    business_date: date,
    dimensions: dict[str, Any],
    components: dict[str, Any],
) -> float | None:
    """Keep non-headline compatibility columns blank on an older user catalog."""
    method = catalog.method_for(metric_id, business_date, dimensions)
    return evaluate_metric(method, components).value if method is not None else None


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
            UNION ALL SELECT min(period || '-01') FROM raw.bonus_import WHERE active=true
            UNION ALL SELECT max(period || '-01') FROM raw.bonus_import WHERE active=true
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
                                       AND f.source_variant='START_END'
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
        WITH chosen_shift AS (
            SELECT r.source_file_id, r.source_row, r.schedule_date, r.agent_id,
                   r.agent_name, r.assignment, r.scheduled_start, r.scheduled_end,
                   r.parse_ok, f.file_name AS source_file,
                   row_number() OVER (
                       PARTITION BY r.schedule_date, r.agent_id
                       ORDER BY f.modified_at DESC NULLS LAST, f.file_name DESC, r.source_row DESC
                   ) AS row_rank
            FROM raw.schedule_shift r
            JOIN meta.source_file f ON f.file_id=r.source_file_id
            WHERE f.active=true AND f.status='SUCCESS'
              AND f.source_variant='ACTIVITIES' AND r.agent_id IS NOT NULL
              AND r.schedule_date BETWEEN ? AND ?
        )
        SELECT e.source_file_id, e.source_row, e.event_index, e.schedule_date,
               e.agent_id, e.agent_name, e.activity, e.activity_type,
               e.event_start, e.event_end, e.parse_ok, s.source_file
        FROM chosen_shift s
        JOIN raw.schedule_event e
          ON e.source_file_id=s.source_file_id AND e.source_row=s.source_row
        WHERE s.row_rank=1 AND e.parse_ok AND e.event_end >= ?
          AND e.event_start < ?
        UNION ALL
        SELECT s.source_file_id, s.source_row, -1 AS event_index,
               s.schedule_date, s.agent_id, s.agent_name,
               s.assignment AS activity, 'FINAL_ASSIGNMENT' AS activity_type,
               s.scheduled_start AS event_start, s.scheduled_end AS event_end,
               s.parse_ok, s.source_file
        FROM chosen_shift s
        WHERE s.row_rank=1 AND s.parse_ok AND s.assignment IS NOT NULL
          AND s.scheduled_end >= ? AND s.scheduled_start < ?
        ORDER BY schedule_date, agent_id, event_start, event_index
        """,
        [
            start - timedelta(days=1), end + timedelta(days=1),
            window_start, window_end, window_start, window_end,
        ],
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
            UNION SELECT agent_id FROM raw.fte_agent r
                  JOIN meta.source_file f ON f.file_id=r.source_file_id AND f.active AND f.status='SUCCESS'
                  WHERE agent_id IS NOT NULL
                    AND (
                        upper(trim(coalesce(employment_status,'')))='ACTIVE'
                        OR (upper(trim(coalesce(employment_status,'')))='LEAVER' AND end_date IS NOT NULL)
                    )
        ), fte_ranked AS (
                SELECT r.agent_id, r.agent_name, r.employment_status, r.team_leader,
                       r.ops_manager, r.lob, r.market, r.language, r.location,
                       r.city, r.fte, r.end_date,
                       row_number() OVER (
                    PARTITION BY r.agent_id
                    ORDER BY CASE WHEN upper(coalesce(employment_status,''))='ACTIVE' THEN 0 ELSE 1 END,
                             end_date DESC NULLS LAST, f.modified_at DESC NULLS LAST, source_row DESC
                ) row_rank
                FROM raw.fte_agent r JOIN meta.source_file f ON f.file_id=r.source_file_id AND f.active AND f.status='SUCCESS'
                WHERE r.agent_id IS NOT NULL
                  AND (
                      upper(trim(coalesce(r.employment_status,'')))='ACTIVE'
                      OR (upper(trim(coalesce(r.employment_status,'')))='LEAVER' AND r.end_date IS NOT NULL)
                  )
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


PLANNED_TIME_OFF_COLUMNS = [
    "segment_key", "agent_day_key", "business_date", "agent_id", "agent_name",
    "team_leader", "ops_manager", "lob", "language", "source_kind",
    "absence_type", "record_status", "segment_start", "segment_end",
    "planned_minutes", "source_file", "source_sheet", "source_row",
]


def _build_planned_time_off(
    conn: DatabaseConnection,
    schedules: list[dict[str, Any]],
    agents: dict[str, dict[str, Any]],
    start: date,
    end: date,
    as_of: datetime,
) -> dict[str, list[dict[str, Any]]]:
    """Clip approved PTO/Away rows to shifts and make overlaps exclusive."""

    source = _dicts(conn.execute(
        """SELECT r.*, f.file_name AS source_file
           FROM raw.fte_time_off r
           JOIN meta.source_file f
             ON f.file_id=r.source_file_id AND f.active AND f.status='SUCCESS'
           WHERE r.start_date<=? AND coalesce(r.end_date, ?) >= ?
             AND (
               (r.source_kind='PTO' AND r.record_status='APPROVED')
               OR (r.source_kind='AWAY' AND r.record_status IN ('ACTIVE','PLANNED','CLOSED'))
             )
           ORDER BY r.agent_id, r.start_date, r.source_kind, r.source_row""",
        [end, end, start],
    ))
    by_agent: dict[str, list[dict[str, Any]]] = defaultdict(list)
    for row in source:
        by_agent[row["agent_id"]].append(row)

    output: list[dict[str, Any]] = []
    grouped: dict[str, list[dict[str, Any]]] = defaultdict(list)
    for shift in schedules:
        shift_start, shift_end = shift["scheduled_start"], shift["scheduled_end"]
        agent_id = shift["agent_id"]
        business_date = shift["schedule_date"]
        if (
            not agent_id or not shift_start or not shift_end or shift_end <= shift_start
            or shift["assignment_type"] == "Off"
        ):
            continue
        candidates: list[dict[str, Any]] = []
        for row in by_agent.get(agent_id, []):
            if business_date < row["start_date"]:
                continue
            if row["end_date"] is not None and business_date > row["end_date"]:
                continue
            if row["source_kind"] == "PTO" and row["day_coverage"] == "PARTIAL_DAY":
                segment_start = datetime.combine(business_date, row["start_time"])
                segment_end = datetime.combine(business_date, row["end_time"])
            else:
                segment_start, segment_end = shift_start, shift_end
            # Planned Away is a capacity planning input only.  It cannot erase
            # attendance evidence which has already happened.
            if row["source_kind"] == "AWAY" and row["record_status"] == "PLANNED":
                if segment_end <= as_of:
                    continue
                segment_start = max(segment_start, as_of)
            segment_start = max(shift_start, segment_start)
            segment_end = min(shift_end, segment_end)
            if segment_end > segment_start:
                candidates.append({**row, "segment_start": segment_start, "segment_end": segment_end})
        if not candidates:
            continue

        # Away outranks PTO when two registers overlap. Within the same kind,
        # the later physical row is the deterministic winner. The boundary
        # sweep ensures downstream minutes are never double counted.
        boundaries = sorted({
            value
            for item in candidates
            for value in (item["segment_start"], item["segment_end"])
        })
        exclusive: list[dict[str, Any]] = []
        for left, right in zip(boundaries, boundaries[1:]):
            active = [
                item for item in candidates
                if item["segment_start"] < right and item["segment_end"] > left
            ]
            if not active:
                continue
            chosen = max(
                active,
                key=lambda item: (
                    item["source_kind"] == "AWAY",
                    item["source_row"],
                ),
            )
            current = {**chosen, "segment_start": left, "segment_end": right}
            identity = (
                chosen["source_kind"], chosen["absence_type"], chosen["record_status"],
                chosen["source_file"], chosen["source_sheet"], chosen["source_row"],
            )
            if exclusive and exclusive[-1]["segment_end"] == left and exclusive[-1]["identity"] == identity:
                exclusive[-1]["segment_end"] = right
            else:
                exclusive.append({**current, "identity": identity})
        agent = agents.get(agent_id, {})
        agent_day_key = f"{business_date:%Y%m%d}-{agent_id}"
        for item in exclusive:
            minutes = int((item["segment_end"] - item["segment_start"]).total_seconds() // 60)
            if minutes <= 0:
                continue
            segment_key = hashlib.sha256(
                f"{agent_day_key}|{item['source_kind']}|{item['absence_type']}|"
                f"{item['segment_start']}|{item['segment_end']}".encode("utf-8")
            ).hexdigest()
            result = {
                "segment_key": segment_key, "agent_day_key": agent_day_key,
                "business_date": business_date, "agent_id": agent_id,
                "agent_name": agent.get("canonical_name") or shift["agent_name"],
                "team_leader": agent.get("team_leader"),
                "ops_manager": agent.get("ops_manager"), "lob": agent.get("lob"),
                "language": agent.get("language"),
                "source_kind": item["source_kind"], "absence_type": item["absence_type"],
                "record_status": item["record_status"],
                "segment_start": item["segment_start"], "segment_end": item["segment_end"],
                "planned_minutes": minutes, "source_file": item["source_file"],
                "source_sheet": item["source_sheet"], "source_row": item["source_row"],
            }
            output.append(result)
            grouped[agent_day_key].append(result)
    output.sort(key=lambda row: (row["business_date"], row["agent_id"], row["segment_start"]))
    conn.execute("DELETE FROM mart.planned_time_off_segment")
    _insert_dicts(conn, "mart.planned_time_off_segment", PLANNED_TIME_OFF_COLUMNS, output)
    return grouped


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
    "shift_state", "call_action", "requires_call", "is_provisional", "evaluation_as_of",
    "planned_work_minutes", "planning_overlay", "planning_overlay_minutes",
    "planning_overlay_source",
]

UNRELIABLE_ATTENDANCE_RESULTS = {
    "Schedule parse error", "Data not loaded", "Missing actual evidence",
    "Incomplete actual evidence", "No schedule overlap",
}


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
    planned_time_off: dict[str, list[dict[str, Any]]],
    as_of: datetime,
    minimum_status_coverage: float,
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
        # Activities are the corrected Verint ledger, never operational attendance
        # evidence. Only the StartEndTimes assignment can make the shift planned
        # absence at this stage.
        agent_day_key = f"{shift['schedule_date']:%Y%m%d}-{agent_id}"
        time_off_rows = planned_time_off.get(agent_day_key, [])
        time_off_intervals = [
            (item["segment_start"], item["segment_end"])
            for item in time_off_rows
        ]
        time_off_minutes = (
            interval_minutes(start, end, time_off_intervals)
            if start and end else 0
        )
        time_off_labels = list(dict.fromkeys(
            f"{item['source_kind']}: {item['absence_type']}" for item in time_off_rows
        ))
        planning_overlay = "; ".join(time_off_labels) or None
        planning_source = "; ".join(sorted({item["source_file"] for item in time_off_rows})) or None
        full_time_off = bool(scheduled_minutes and time_off_minutes >= scheduled_minutes)
        effective_assignment_type = (
            "Planned absence"
            if shift["assignment_type"] == "Planned absence" or full_time_off
            else shift["assignment_type"]
        )
        effective_assignment = planning_overlay if full_time_off and planning_overlay else shift["assignment"]
        planned_absence = (
            scheduled_minutes
            if shift["assignment_type"] == "Planned absence"
            else min(scheduled_minutes, time_off_minutes)
        )
        planned_work_minutes = max(0, scheduled_minutes - planned_absence)
        working_intervals = (
            subtract_intervals(start, end, time_off_intervals)
            if start and end and effective_assignment_type != "Planned absence"
            else []
        )
        first, last, source_loaded, row_present, lilo_source = _lilo_boundaries(shift, lilo, loaded_dates)
        evidence_end = min(end, as_of) if end and start and as_of > start else start
        statuses = (
            _statuses_for_shift(agent_id, start, evidence_end, statuses_by_day)
            if start and evidence_end and evidence_end > start else []
        )
        if start and end and end > start:
            status_window_end = evidence_end if evidence_end and evidence_end > start else start
            status_category, status_covered, status_exclusive = (
                _exclusive_category_minutes(start, status_window_end, statuses)
                if status_window_end > start else ({}, 0, [])
            )
            gross_status_exclusive = status_exclusive
            expected_exclusive: list[dict[str, Any]] = []
            for item in status_exclusive:
                for left, right in subtract_intervals(
                    item["interval_start"], item["interval_end"], time_off_intervals,
                ):
                    expected_exclusive.append({**item, "interval_start": left, "interval_end": right})
            status_exclusive = expected_exclusive
            status_category = defaultdict(int)
            for item in status_exclusive:
                status_category[item["actual_category"]] += int(
                    (item["interval_end"] - item["interval_start"]).total_seconds() // 60
                )
            status_category = dict(status_category)
            status_covered = sum(status_category.values())
            required_status_dates = {
                start.date() + timedelta(days=offset)
                for offset in range(((end - timedelta(microseconds=1)).date() - start.date()).days + 1)
            }
        else:
            status_category, status_covered, status_exclusive = {}, 0, []
            gross_status_exclusive = []
            required_status_dates = {shift["schedule_date"]}
        status_source_loaded = required_status_dates <= status_loaded_dates
        status_sources = "; ".join(sorted({row["source_file"] for row in statuses})) or None
        # Any non-Logged-Off state proves the agent is connected, including an
        # Unavailable state that may still require its own correction. Using
        # the complete interval timeline means a logout followed by a later
        # return is an internal gap, never a false early leave.
        status_presence = [
            row for row in status_exclusive
            if row["actual_category"] != "Logged Off"
        ]
        status_first = min((row["interval_start"] for row in status_presence), default=None)
        status_last = max((row["interval_end"] for row in status_presence), default=None)
        bounded_first = first if first is not None and first <= as_of else None
        bounded_last = min(last, as_of) if last is not None else None
        elapsed_minutes = (
            max(0, int((min(end, as_of) - start).total_seconds() // 60))
            if start and end and as_of > start else 0
        )
        elapsed_work_minutes = sum(
            int((min(right, as_of) - left).total_seconds() // 60)
            for left, right in working_intervals if left < as_of and min(right, as_of) > left
        )
        status_coverage_ratio = status_covered / elapsed_work_minutes if elapsed_work_minutes else 0.0
        status_is_primary = bool(statuses) and status_coverage_ratio >= minimum_status_coverage
        if status_is_primary and status_first is not None:
            actual_first = status_first
        else:
            actual_first = min((value for value in (bounded_first, status_first) if value is not None), default=None)
        if status_is_primary and status_last is not None:
            actual_last = status_last
        else:
            actual_last = max((value for value in (bounded_last, status_last) if value is not None), default=None)
        shift_not_started = bool(start and as_of < start)
        shift_in_progress = bool(start and end and start <= as_of < end)
        shift_complete = bool(end and as_of >= end)
        evidence_parts = []
        if first is not None or last is not None or row_present:
            evidence_parts.append("LILO")
        if statuses:
            evidence_parts.append("AGENT_STATUS")
        actual_evidence = "+".join(evidence_parts) or "NONE"
        raw_late = max(0, int((actual_first - start).total_seconds() // 60)) if actual_first and start else 0
        raw_early = max(0, int((end - actual_last).total_seconds() // 60)) if actual_last and end and shift_complete else 0
        usable_pair = bool(actual_first and actual_last and start and end and actual_last >= actual_first and actual_last > start and actual_first < end)
        late_segments = (
            subtract_intervals(start, min(actual_first, end), time_off_intervals)
            if usable_pair and actual_first > start else []
        )
        early_segments = (
            subtract_intervals(max(actual_last, start), end, time_off_intervals)
            if usable_pair and shift_complete and end > actual_last else []
        )
        late = sum(int((b - a).total_seconds() // 60) for a, b in late_segments)
        early = sum(int((b - a).total_seconds() // 60) for a, b in early_segments)
        late = 0 if late <= tolerance else late
        early = 0 if early <= tolerance else early
        # Two independent sources can prove a completed no-show:
        #   1. LILO contains the scheduled agent/day but both boundaries are blank.
        #   2. Agent Status covers enough of the shift and every observed interval
        #      is explicitly Logged Off.
        # Merely missing from an otherwise loaded extract is not proof of absence;
        # that stays Missing actual evidence and is surfaced as a data exception.
        blank_lilo_row = bool(source_loaded and row_present and first is None and last is None)
        status_proves_disconnected = bool(
            status_exclusive and not status_presence
            and status_coverage_ratio >= minimum_status_coverage
        )
        no_show_segments = working_intervals if (
            (blank_lilo_row or status_proves_disconnected)
            and effective_assignment_type not in {"Off", "Planned absence"}
            and shift_complete
        ) else []
        no_show = sum(int((b - a).total_seconds() // 60) for a, b in no_show_segments)
        worked_span = int((actual_last - actual_first).total_seconds() // 60) if actual_first and actual_last and actual_last >= actual_first else 0
        parse_ok = bool(shift["parse_ok"])
        if not parse_ok and effective_assignment_type != "Off":
            result = "Schedule parse error"
        elif effective_assignment_type == "Off":
            result = "Off"
        elif effective_assignment_type == "Planned absence":
            if full_time_off and time_off_rows:
                kinds = {item["source_kind"] for item in time_off_rows}
                result = "PTO" if kinds == {"PTO"} else "Away" if kinds == {"AWAY"} else "Planned time off"
            else:
                result = "Planned absence"
        elif shift_not_started:
            result = "Not started"
        elif shift_in_progress and late:
            result = "Late - shift in progress"
        elif (
            shift_in_progress and actual_first is None and start
            and elapsed_work_minutes > tolerance
            and (source_loaded or status_source_loaded)
        ):
            result = "Not seen - shift in progress"
        elif shift_in_progress:
            result = "Shift in progress"
        elif not source_loaded and not status_source_loaded:
            result = "Data not loaded"
        elif no_show:
            result = "No show - partial time off" if time_off_minutes else "No show"
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
            result = "Present - partial time off" if time_off_minutes else "Present"
        shift_state = (
            "NOT_STARTED" if shift_not_started else
            "IN_PROGRESS" if shift_in_progress else
            "COMPLETE" if shift_complete else
            "INVALID"
        )
        if effective_assignment_type in {"Off", "Planned absence"}:
            call_action = "NONE"
        elif no_show:
            call_action = "CALL_NO_SHOW"
        elif result == "Not seen - shift in progress":
            call_action = "CALL_NOT_SEEN_NOW"
        elif late:
            call_action = "CALL_LATE"
        else:
            call_action = "NONE"
        agent = agents.get(agent_id, {})
        rows.append({
            "agent_day_key": agent_day_key,
            "business_date": shift["schedule_date"], "agent_id": agent_id,
            "agent_name": agent.get("canonical_name") or shift["agent_name"],
            "team_leader": agent.get("team_leader"), "ops_manager": agent.get("ops_manager"),
            "lob": agent.get("lob"), "market": agent.get("market"), "language": agent.get("language"), "location": agent.get("location"),
            "scheduled_start": start, "scheduled_end": end, "scheduled_minutes": scheduled_minutes,
            "assignment": effective_assignment, "assignment_type": effective_assignment_type,
            "planned_absence_minutes": planned_absence,
            "first_login": first, "last_logout": last, "source_loaded": source_loaded or status_source_loaded, "lilo_row_present": row_present,
            "seen_in_lilo": agent_id in seen_ids, "raw_late_minutes": raw_late, "raw_early_leave_minutes": raw_early,
            "uncoded_late_minutes": late, "uncoded_early_leave_minutes": early, "no_show_minutes": no_show,
            "worked_span_minutes": worked_span, "attendance_result": result, "attendance_percent": None,
            "schedule_source": shift["source_file"], "lilo_source": lilo_source,
            "actual_first_seen": actual_first, "actual_last_seen": actual_last,
            "actual_evidence": actual_evidence, "status_covered_minutes": status_covered,
            "status_source": status_sources,
            "shift_state": shift_state, "call_action": call_action,
            "requires_call": call_action != "NONE",
            "is_provisional": shift_state in {"NOT_STARTED", "IN_PROGRESS"},
            "evaluation_as_of": as_of,
            "planned_work_minutes": planned_work_minutes,
            "planning_overlay": planning_overlay,
            "planning_overlay_minutes": time_off_minutes,
            "planning_overlay_source": planning_source,
            "_events": events, "_late_segments": late_segments,
            "_early_segments": early_segments, "_no_show_segments": no_show_segments,
            "_planned_time_off_segments": time_off_rows,
            "_status_exclusive": status_exclusive, "_status_category": status_category,
            "_status_exclusive_gross": gross_status_exclusive,
            "_status_is_primary": status_is_primary,
            "_evidence_complete": result not in UNRELIABLE_ATTENDANCE_RESULTS,
            "_schedule_source_file_id": shift["source_file_id"],
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
        if not start or not end or end <= start or row["assignment_type"] in {"Off", "Planned absence"}:
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


def _exact_status_gaps(
    intervals: Iterable[tuple[datetime, datetime]],
    minimum_minutes: int,
) -> list[tuple[datetime, datetime]]:
    """Keep physical status boundaries; never bridge a real return to service.

    The configured tolerance suppresses tiny individual gaps. It must not join
    two separate Logged Off/Unavailable spells across even a short active
    interval, because Verint needs one correction per exact continuous spell.
    """

    return [
        (start, end)
        for start, end in merge_intervals(intervals)
        if int((end - start).total_seconds() // 60) > minimum_minutes
    ]


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

CORRECTION_RESIDUAL_COLUMNS = [
    "residual_id", "correction_id", "business_date", "agent_id",
    "residual_start", "residual_end", "residual_minutes",
    "suggested_activity", "observed_source", "source_file",
    "verint_reconciliation",
]


def _correction_id(row: dict[str, Any], issue: str, start: datetime | None, end: datetime | None) -> str:
    clean = "".join(char for char in issue.upper() if char.isalnum())
    start_key = start.strftime("%H%M%S") if start else "DAY"
    end_key = end.strftime("%H%M%S") if end else "DAY"
    return f"{row['business_date']:%Y%m%d}-{row['agent_id']}-{start_key}-{end_key}-{clean}"


def _final_verint_events(base: dict[str, Any], rulebook: Rulebook) -> list[dict[str, Any]]:
    """Return post-day Verint activities that can explain an observed gap."""
    candidates: list[dict[str, Any]] = []
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
        if row["no_show_minutes"] > 0 and start and end:
            for gap_start, gap_end in row.get("_no_show_segments", []):
                minutes = int((gap_end - gap_start).total_seconds() // 60)
                if minutes <= 0:
                    continue
                source = "; ".join(value for value in (row["lilo_source"], row["status_source"]) if value) or row["schedule_source"]
                add(row, "No show", gap_start, gap_end, minutes, 1, "High", "Review absence reason", row["actual_evidence"], source)
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
        exclusive = row.get("_status_exclusive_gross", row.get("_status_exclusive", []))
        for category, issue, priority in (("Logged Off", "Mid-shift logged off", 4), ("Unavailable", "Unavailable in shift", 7)):
            raw = [(item["interval_start"], item["interval_end"]) for item in exclusive if item["actual_category"] == category]
            exact_gaps = _exact_status_gaps(raw, status_tolerance)
            clipped = clip_intervals(actual_first, actual_last, exact_gaps)
            time_off = [
                (item["segment_start"], item["segment_end"])
                for item in row.get("_planned_time_off_segments", [])
            ]
            for raw_start, raw_end in clipped:
                for hit_start, hit_end in subtract_intervals(raw_start, raw_end, time_off):
                    minutes = int((hit_end - hit_start).total_seconds() // 60)
                    if minutes <= status_tolerance:
                        continue
                    confidence = "High" if category == "Logged Off" else "Review"
                    add(row, issue, hit_start, hit_end, minutes, priority, confidence, "General Unavailability", "AGENT_STATUS", row["status_source"])

    by_key = {row["agent_day_key"]: row for row in attendance}
    residual_rows: list[dict[str, Any]] = []
    for item in output:
        base = by_key[item["business_date"].strftime("%Y%m%d") + "-" + item["agent_id"]]
        gap_start, gap_end = item["gap_start"], item["gap_end"]
        matches: list[tuple[int, dict[str, Any]]] = []
        matched_intervals: list[tuple[datetime, datetime]] = []
        if gap_start and gap_end:
            for event in _final_verint_events(base, rulebook):
                overlap = interval_minutes(gap_start, gap_end, [(event["start"], event["end"])])
                if overlap > 0:
                    matches.append((overlap, event))
                    matched_intervals.append((max(gap_start, event["start"]), min(gap_end, event["end"])))
        if matches:
            overlap = interval_minutes(gap_start, gap_end, matched_intervals)
            status = "CORRECTED" if overlap >= max(1, item["gap_minutes"] - match_tolerance) else "PARTIAL"
            activities = "; ".join(sorted({str(event["activity"]) for _, event in matches}))
            categories = "; ".join(sorted({str(event["category"]) for _, event in matches}))
            sources = "; ".join(sorted({str(event["source_file"]) for _, event in matches if event["source_file"]})) or None
            item.update({
                "verint_reconciliation": status, "verint_activity": activities,
                "verint_category": categories, "verint_overlap_minutes": overlap,
                "verint_source_file": sources,
            })
        else:
            item.update({
                "verint_reconciliation": "NOT_APPLICABLE" if not gap_start or not gap_end else "NOT_CORRECTED",
                "verint_activity": None, "verint_category": None, "verint_overlap_minutes": 0,
                "verint_source_file": None,
            })
        if gap_start and gap_end:
            for residual_start, residual_end in subtract_intervals(gap_start, gap_end, matched_intervals):
                residual_minutes = int((residual_end - residual_start).total_seconds() // 60)
                if residual_minutes <= match_tolerance:
                    continue
                residual_id = hashlib.sha256(
                    f"{item['correction_id']}|{residual_start}|{residual_end}".encode("utf-8")
                ).hexdigest()
                residual_rows.append({
                    "residual_id": residual_id, "correction_id": item["correction_id"],
                    "business_date": item["business_date"], "agent_id": item["agent_id"],
                    "residual_start": residual_start, "residual_end": residual_end,
                    "residual_minutes": residual_minutes,
                    "suggested_activity": item["suggested_activity"],
                    "observed_source": item["observed_source"],
                    "source_file": item["source_file"],
                    "verint_reconciliation": item["verint_reconciliation"],
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
    conn.execute("DELETE FROM mart.correction_residual_segment")
    _insert_dicts(conn, "mart.correction_residual_segment", CORRECTION_RESIDUAL_COLUMNS, residual_rows)

    exceptions: list[dict[str, Any]] = []
    gaps_by_key: dict[str, list[dict[str, Any]]] = defaultdict(list)
    for item in output:
        gaps_by_key[f"{item['business_date']:%Y%m%d}-{item['agent_id']}"].append(item)
    for base in attendance:
        for event in _final_verint_events(base, rulebook):
            planned_overlap = interval_minutes(
                event["start"], event["end"],
                [
                    (item["segment_start"], item["segment_end"])
                    for item in base.get("_planned_time_off_segments", [])
                ],
            )
            if planned_overlap > 0:
                continue
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


STAFFING_COLUMNS = [
    "business_date", "interval_start", "interval_end", "lob", "language",
    "scheduled_agents", "observed_agents", "productive_agents", "auxiliary_agents",
    "scheduled_fte", "elapsed_scheduled_fte", "observed_fte", "productive_fte",
    "staffing_variance_fte", "staffing_gap_fte", "staffing_state",
    "evidence_basis", "evaluation_as_of",
    "gross_scheduled_fte", "planned_time_off_fte",
]

TIMELINE_COLUMNS = [
    "segment_key", "agent_day_key", "business_date", "agent_id", "agent_name",
    "team_leader", "ops_manager", "lob", "language", "scheduled_start",
    "scheduled_end", "segment_start", "segment_end", "segment_minutes",
    "planned_state", "actual_status", "actual_category", "mismatch_type",
    "is_gap", "observed_source", "source_file", "evaluation_as_of",
]


def _quarter_intervals(start: datetime, end: datetime) -> Iterable[tuple[datetime, datetime]]:
    cursor = start.replace(minute=(start.minute // 15) * 15, second=0, microsecond=0)
    while cursor < end:
        right = cursor + timedelta(minutes=15)
        yield cursor, right
        cursor = right


def _overlap_seconds(left: datetime, right: datetime, start: datetime, end: datetime) -> float:
    return max(0.0, (min(right, end) - max(left, start)).total_seconds())


def _build_staffing(
    conn: DatabaseConnection,
    attendance: list[dict[str, Any]],
    as_of: datetime,
) -> int:
    buckets: dict[tuple[date, datetime, str, str], dict[str, Any]] = {}

    def bucket_for(row: dict[str, Any], left: datetime, right: datetime) -> dict[str, Any]:
        lob = str(row.get("lob") or "(blank)")
        language = str(row.get("language") or "(blank)")
        key = (row["business_date"], left, lob, language)
        return buckets.setdefault(key, {
            "business_date": row["business_date"], "interval_start": left,
            "interval_end": right, "lob": lob, "language": language,
            "scheduled_ids": set(), "observed_ids": set(), "productive_ids": set(),
            "auxiliary_ids": set(), "scheduled_seconds": 0.0,
            "gross_scheduled_seconds": 0.0, "time_off_seconds": 0.0,
            "source_available_seconds": 0.0,
            "elapsed_scheduled_seconds": 0.0, "observed_seconds": 0.0,
            "productive_seconds": 0.0, "evidence": set(),
        })

    for row in attendance:
        start, end = row["scheduled_start"], row["scheduled_end"]
        if (
            not start or not end or end <= start
            or row["assignment_type"] == "Off"
        ):
            continue
        time_off_intervals = [
            (item["segment_start"], item["segment_end"])
            for item in row.get("_planned_time_off_segments", [])
        ]
        if row["assignment_type"] == "Planned absence" and not time_off_intervals:
            time_off_intervals = [(start, end)]
        working_intervals = subtract_intervals(start, end, time_off_intervals)
        elapsed_end = min(end, as_of)
        for left, right in _quarter_intervals(start, end):
            bucket = bucket_for(row, left, right)
            gross = _overlap_seconds(left, right, start, end)
            scheduled = sum(_overlap_seconds(left, right, a, b) for a, b in working_intervals)
            planned = max(0.0, gross - scheduled)
            elapsed = sum(
                _overlap_seconds(left, right, a, min(b, elapsed_end))
                for a, b in working_intervals if a < elapsed_end
            )
            bucket["gross_scheduled_seconds"] += gross
            bucket["time_off_seconds"] += planned
            if scheduled:
                bucket["scheduled_ids"].add(row["agent_id"])
                bucket["scheduled_seconds"] += scheduled
                if row.get("source_loaded") and row.get("_evidence_complete"):
                    bucket["source_available_seconds"] += scheduled
                bucket["elapsed_scheduled_seconds"] += elapsed

        expected_exclusive = row.get("_status_exclusive", [])
        exclusive = row.get("_status_exclusive_gross", expected_exclusive)
        active_statuses = [
            item for item in exclusive
            if item["actual_category"] in {"Productive", "Auxiliary", "Lunch", "Break"}
        ]
        actual_intervals: list[tuple[datetime, datetime, str, str]] = []
        for item in active_statuses:
            actual_intervals.append((
                item["interval_start"], min(item["interval_end"], as_of),
                item["actual_category"], "AGENT_STATUS",
            ))
        # LILO can fill only portions for which Agent Status has no state at all.
        # It must not turn an explicit Logged Off/Unavailable interval into presence.
        if row["actual_first_seen"] and row["actual_last_seen"]:
            lilo_start = max(start, row["actual_first_seen"])
            lilo_end = min(end, as_of, row["actual_last_seen"])
            covered = merge_intervals(
                (item["interval_start"], min(item["interval_end"], as_of))
                for item in exclusive if min(item["interval_end"], as_of) > item["interval_start"]
            )
            if lilo_end > lilo_start:
                for residual_start, residual_end in subtract_intervals(lilo_start, lilo_end, covered):
                    actual_intervals.append((residual_start, residual_end, "LILO_PRESENT", "LILO"))
        for actual_start, actual_end, category, basis in actual_intervals:
            if actual_end <= actual_start:
                continue
            for left, right in _quarter_intervals(actual_start, actual_end):
                bucket = bucket_for(row, left, right)
                seconds = _overlap_seconds(left, right, actual_start, actual_end)
                if not seconds:
                    continue
                bucket["observed_ids"].add(row["agent_id"])
                bucket["observed_seconds"] += seconds
                bucket["evidence"].add(basis)
                if category == "Productive":
                    bucket["productive_ids"].add(row["agent_id"])
                    bucket["productive_seconds"] += seconds
                elif category == "Auxiliary":
                    bucket["auxiliary_ids"].add(row["agent_id"])

    output: list[dict[str, Any]] = []
    for item in buckets.values():
        scheduled_fte = item["scheduled_seconds"] / 900
        gross_scheduled_fte = item["gross_scheduled_seconds"] / 900
        planned_time_off_fte = item["time_off_seconds"] / 900
        elapsed_scheduled_fte = item["elapsed_scheduled_seconds"] / 900
        observed_fte = item["observed_seconds"] / 900
        productive_fte = item["productive_seconds"] / 900
        variance = observed_fte - elapsed_scheduled_fte
        gap = max(0.0, -variance)
        scheduled_seconds = item["scheduled_seconds"]
        available_seconds = item["source_available_seconds"]
        if item["interval_start"] >= as_of:
            state = "FUTURE"
            variance = None
            gap = None
        elif available_seconds + 0.001 < scheduled_seconds:
            state = "DATA_MISSING" if available_seconds <= 0.001 else "DATA_PARTIAL"
            variance = None
            gap = None
        elif item["interval_start"] < as_of < item["interval_end"]:
            state = "PARTIAL_GAP" if gap is not None and gap > 0.001 else "PARTIAL_OK"
        else:
            state = "GAP" if gap is not None and gap > 0.001 else "OK"
        evidence = "+".join(sorted(item["evidence"])) or "NONE"
        if available_seconds + 0.001 < scheduled_seconds:
            evidence = f"{evidence}+MISSING_SOURCE"
        output.append({
            **{key: item[key] for key in (
                "business_date", "interval_start", "interval_end", "lob", "language",
            )},
            "scheduled_agents": len(item["scheduled_ids"]),
            "observed_agents": len(item["observed_ids"]),
            "productive_agents": len(item["productive_ids"]),
            "auxiliary_agents": len(item["auxiliary_ids"]),
            "scheduled_fte": scheduled_fte,
            "elapsed_scheduled_fte": elapsed_scheduled_fte,
            "observed_fte": observed_fte,
            "productive_fte": productive_fte,
            "staffing_variance_fte": variance,
            "staffing_gap_fte": gap,
            "staffing_state": state,
            "evidence_basis": evidence,
            "evaluation_as_of": as_of,
            "gross_scheduled_fte": gross_scheduled_fte,
            "planned_time_off_fte": planned_time_off_fte,
        })
    output.sort(key=lambda row: (row["business_date"], row["interval_start"], row["lob"], row["language"]))
    conn.execute("DELETE FROM mart.staffing_interval")
    _insert_dicts(conn, "mart.staffing_interval", STAFFING_COLUMNS, output)
    return len(output)


def _build_shift_timeline(
    conn: DatabaseConnection,
    attendance: list[dict[str, Any]],
    as_of: datetime,
) -> int:
    output: list[dict[str, Any]] = []
    for row in attendance:
        start, end = row["scheduled_start"], row["scheduled_end"]
        if (
            not start or not end or end <= start or row["assignment_type"] == "Off"
            or (
                row["assignment_type"] == "Planned absence"
                and not row.get("_planned_time_off_segments")
            )
        ):
            continue
        exclusive = row.get("_status_exclusive_gross", row.get("_status_exclusive", []))
        boundaries = {start, end}
        if start < as_of < end:
            boundaries.add(as_of)
        for item in exclusive:
            boundaries.update((item["interval_start"], item["interval_end"]))
        planned_segments = row.get("_planned_time_off_segments", [])
        for item in planned_segments:
            boundaries.update((item["segment_start"], item["segment_end"]))
        for value in (row["actual_first_seen"], row["actual_last_seen"]):
            if value is not None and start < value < end:
                boundaries.add(value)
        ordered = sorted(boundaries)
        for left, right in zip(ordered, ordered[1:]):
            minutes = int((right - left).total_seconds() // 60)
            if minutes <= 0:
                continue
            chosen = next((
                item for item in exclusive
                if item["interval_start"] <= left and item["interval_end"] >= right
            ), None)
            actual_status = chosen.get("status") if chosen else None
            source_file = chosen.get("source_file") if chosen else None
            planned = next((
                item for item in planned_segments
                if item["segment_start"] <= left and item["segment_end"] >= right
            ), None)
            if planned is not None:
                active_during_time_off = chosen is not None and chosen["actual_category"] != "Logged Off"
                category = chosen["actual_category"] if active_during_time_off else planned["source_kind"]
                mismatch = "WORK_DURING_TIME_OFF" if active_during_time_off else "PLANNED_TIME_OFF"
                is_gap = False
                source = "AGENT_STATUS" if active_during_time_off else planned["source_kind"]
                source_file = chosen.get("source_file") if active_during_time_off else planned["source_file"]
            elif left >= as_of:
                category, mismatch, is_gap, source = "FUTURE", "FUTURE", False, "NONE"
            elif any(a <= left and b >= right for a, b in row.get("_no_show_segments", [])):
                category, mismatch, is_gap, source = "NO_ACTIVITY", "NO_SHOW", True, row["actual_evidence"]
                source_file = row["lilo_source"] or row["status_source"]
            elif chosen is not None:
                category = chosen["actual_category"]
                if category == "Logged Off":
                    if row["actual_first_seen"] and right <= row["actual_first_seen"]:
                        mismatch, is_gap = "LATE", True
                    elif row["actual_last_seen"] and left >= row["actual_last_seen"] and as_of >= end:
                        mismatch, is_gap = "EARLY_LEAVE", True
                    else:
                        mismatch, is_gap = "LOGGED_OFF", True
                elif category == "Unavailable":
                    mismatch, is_gap = "UNAVAILABLE", True
                else:
                    mismatch, is_gap = "MATCH", False
                source = "AGENT_STATUS"
            elif row["actual_first_seen"] and right <= row["actual_first_seen"]:
                category, mismatch, is_gap, source = "NO_ACTIVITY", "LATE", True, row["actual_evidence"]
                source_file = row["lilo_source"] or row["status_source"]
            elif row["actual_last_seen"] and left >= row["actual_last_seen"] and as_of >= end:
                category, mismatch, is_gap, source = "NO_ACTIVITY", "EARLY_LEAVE", True, row["actual_evidence"]
                source_file = row["lilo_source"] or row["status_source"]
            elif not exclusive and row["actual_first_seen"] and row["actual_last_seen"]:
                category, mismatch, is_gap, source = "LILO_PRESENT", "MATCH", False, "LILO"
                source_file = row["lilo_source"]
            else:
                category, mismatch, is_gap, source = "NO_STATUS_EVIDENCE", "MISSING_STATUS_DATA", False, "NONE"
                source_file = row["status_source"] or row["lilo_source"]
            segment_key = hashlib.sha256(
                f"{row['agent_day_key']}|{left}|{right}|{category}|{mismatch}".encode("utf-8")
            ).hexdigest()
            output.append({
                "segment_key": segment_key, "agent_day_key": row["agent_day_key"],
                "business_date": row["business_date"], "agent_id": row["agent_id"],
                "agent_name": row["agent_name"], "team_leader": row["team_leader"],
                "ops_manager": row["ops_manager"], "lob": row["lob"],
                "language": row["language"], "scheduled_start": start,
                "scheduled_end": end, "segment_start": left, "segment_end": right,
                "segment_minutes": minutes,
                "planned_state": (
                    f"{planned['source_kind']}: {planned['absence_type']}"
                    if planned is not None else row["assignment_type"]
                ),
                "actual_status": actual_status, "actual_category": category,
                "mismatch_type": mismatch, "is_gap": is_gap,
                "observed_source": source, "source_file": source_file,
                "evaluation_as_of": as_of,
            })
    output.sort(key=lambda row: (row["business_date"], row["agent_id"], row["segment_start"]))
    conn.execute("DELETE FROM mart.shift_timeline_segment")
    _insert_dicts(conn, "mart.shift_timeline_segment", TIMELINE_COLUMNS, output)
    return len(output)


FINAL_ABSENCE_EVENT_COLUMNS = [
    "event_key", "agent_day_key", "business_date", "agent_id", "agent_name",
    "team_leader", "ops_manager", "lob", "market", "language", "location",
    "activity", "category", "event_start", "event_end", "minutes", "hours",
    "counts_as_absence", "counts_as_vacation", "counts_as_unpaid",
    "counts_as_shrinkage", "mapped", "evidence_type", "source_file",
    "rule_version", "rule_sha256",
]

FINAL_ABSENCE_DAY_COLUMNS = [
    "agent_day_key", "business_date", "agent_id", "agent_name", "team_leader",
    "ops_manager", "lob", "market", "language", "location", "scheduled_minutes",
    "planned_net_minutes",
    "final_absence_minutes", "final_vacation_minutes", "final_unpaid_minutes",
    "final_shrinkage_minutes", "final_unmapped_minutes", "final_absence_hours",
    "final_absence_rate", "final_absence_day", "final_ledger_status",
    "rule_version", "rule_sha256",
]


def _build_verint_final_absence(
    conn: DatabaseConnection,
    rulebook: Rulebook,
    metric_catalog: MetricCatalog,
    start: date,
    end: date,
    as_of: datetime,
) -> tuple[int, int]:
    # StartEndTimes is the authoritative roster of expected agent-days.  The
    # post-day Activities export is left-attached below; this keeps completely
    # empty or missing Verint rows visible instead of silently dropping them.
    shifts = _dicts(conn.execute(
        """
        WITH ranked AS (
            SELECT r.*, f.file_name AS source_file, f.modified_at,
                   d.canonical_name, d.team_leader, d.ops_manager,
                   d.lob AS roster_lob, d.market, d.language, d.location,
                   row_number() OVER (
                       PARTITION BY r.schedule_date, r.agent_id
                       ORDER BY f.modified_at DESC NULLS LAST, f.file_name DESC, r.source_row DESC
                   ) AS row_rank
            FROM raw.schedule_shift r
            JOIN meta.source_file f ON f.file_id=r.source_file_id
            LEFT JOIN core.dim_agent d ON d.agent_id=r.agent_id
            WHERE f.active=true AND f.status='SUCCESS'
              AND f.source_variant='START_END'
              AND r.schedule_date BETWEEN ? AND ? AND r.agent_id IS NOT NULL
        )
        SELECT * FROM ranked WHERE row_rank=1
        """,
        [start, end],
    ))
    activity_shifts = _dicts(conn.execute(
        """
        WITH ranked AS (
            SELECT r.*, f.file_name AS source_file, f.modified_at,
                   row_number() OVER (
                       PARTITION BY r.schedule_date, r.agent_id
                       ORDER BY f.modified_at DESC NULLS LAST, f.file_name DESC, r.source_row DESC
                   ) AS row_rank
            FROM raw.schedule_shift r
            JOIN meta.source_file f ON f.file_id=r.source_file_id
            WHERE f.active=true AND f.status='SUCCESS'
              AND f.source_variant='ACTIVITIES'
              AND r.schedule_date BETWEEN ? AND ? AND r.agent_id IS NOT NULL
        )
        SELECT * FROM ranked WHERE row_rank=1
        """,
        [start, end],
    ))
    activity_by_key = {
        f"{row['schedule_date']:%Y%m%d}-{row['agent_id']}": row
        for row in activity_shifts
    }
    events = _dicts(conn.execute(
        """SELECT r.*, f.file_name AS source_file
           FROM raw.schedule_event r
           JOIN meta.source_file f ON f.file_id=r.source_file_id
           WHERE f.active=true AND f.status='SUCCESS'
             AND f.source_variant='ACTIVITIES'
             AND r.schedule_date BETWEEN ? AND ? AND r.parse_ok=true""",
        [start, end],
    ))
    events_by_row: dict[tuple[str, int], list[dict[str, Any]]] = defaultdict(list)
    for event in events:
        events_by_row[(event["source_file_id"], event["source_row"])].append(event)
    attendance_by_key = {
        row["agent_day_key"]: row
        for row in _dicts(conn.execute(
            """SELECT agent_day_key, attendance_result, actual_evidence,
                      source_loaded, shift_state, planning_overlay_minutes,
                      planning_overlay
               FROM mart.attendance_agent_day WHERE business_date BETWEEN ? AND ?""",
            [start, end],
        ))
    }
    residual_by_key = {
        f"{row['business_date']:%Y%m%d}-{row['agent_id']}": int(row["minutes"] or 0)
        for row in _dicts(conn.execute(
            """SELECT business_date, agent_id, sum(residual_minutes) AS minutes
               FROM mart.correction_residual_segment
               WHERE business_date BETWEEN ? AND ? GROUP BY business_date, agent_id""",
            [start, end],
        ))
    }
    verint_without_observed_gap = {
        str(row[0])
        for row in conn.execute(
            """SELECT DISTINCT agent_day_key FROM mart.verint_final_exception
               WHERE business_date BETWEEN ? AND ?
                 AND exception_type='VERINT_FINAL_WITHOUT_OBSERVED_GAP'""",
            [start, end],
        ).fetchall()
    }

    event_rows: list[dict[str, Any]] = []
    day_rows: list[dict[str, Any]] = []
    for shift in shifts:
        shift_start, shift_end = shift["scheduled_start"], shift["scheduled_end"]
        if not shift["parse_ok"] or not shift_start or not shift_end or shift_end <= shift_start:
            continue
        agent_day_key = f"{shift['schedule_date']:%Y%m%d}-{shift['agent_id']}"
        candidates: list[tuple[str | None, datetime, datetime, str, str]] = []
        activity_shift = activity_by_key.get(agent_day_key)
        if activity_shift is not None:
            candidates.append((
                activity_shift["assignment"], activity_shift["scheduled_start"],
                activity_shift["scheduled_end"], "SHIFT_ASSIGNMENT",
                activity_shift["source_file"],
            ))
            for event in events_by_row.get(
                (activity_shift["source_file_id"], activity_shift["source_row"]), [],
            ):
                candidates.append((
                    event["activity"], event["event_start"], event["event_end"],
                    "SHIFT_EVENT", event["source_file"],
                ))
        flag_intervals: dict[str, list[tuple[datetime, datetime]]] = defaultdict(list)
        seen: set[tuple[Any, ...]] = set()
        unmapped = 0
        for activity, raw_start, raw_end, evidence_type, source_file in candidates:
            if not raw_start or not raw_end:
                continue
            event_start, event_end = max(shift_start, raw_start), min(shift_end, raw_end)
            if event_end <= event_start:
                continue
            rule = rulebook.classify_activity(activity)
            if evidence_type == "SHIFT_ASSIGNMENT" and rule is None:
                continue
            if rule is not None and (
                rule.working or rule.category in {"OFF", "LUNCH", "BREAK"}
                or not (rule.absence or rule.vacation or rule.unpaid or rule.shrinkage)
            ):
                continue
            mapped = rule is not None
            category = rule.category if rule is not None else "UNMAPPED"
            key = (activity, category, event_start, event_end, evidence_type, source_file)
            if key in seen:
                continue
            seen.add(key)
            minutes = int((event_end - event_start).total_seconds() // 60)
            if minutes <= 0:
                continue
            flags = {
                "absence": bool(rule and rule.absence),
                "vacation": bool(rule and rule.vacation),
                "unpaid": bool(rule and rule.unpaid),
                "shrinkage": bool(rule and rule.shrinkage),
            }
            if not mapped:
                unmapped += minutes
                flag_intervals["unmapped"].append((event_start, event_end))
            for flag, enabled in flags.items():
                if enabled:
                    flag_intervals[flag].append((event_start, event_end))
            event_key = hashlib.sha256(
                f"{agent_day_key}|{activity}|{event_start}|{event_end}|{evidence_type}|{source_file}".encode("utf-8")
            ).hexdigest()
            event_rows.append({
                "event_key": event_key, "agent_day_key": agent_day_key,
                "business_date": shift["schedule_date"], "agent_id": shift["agent_id"],
                "agent_name": shift["canonical_name"] or shift["agent_name"],
                "team_leader": shift["team_leader"], "ops_manager": shift["ops_manager"],
                "lob": shift["roster_lob"], "market": shift["market"],
                "language": shift["language"], "location": shift["location"],
                "activity": activity, "category": category,
                "event_start": event_start, "event_end": event_end,
                "minutes": minutes, "hours": minutes / 60,
                "counts_as_absence": flags["absence"],
                "counts_as_vacation": flags["vacation"],
                "counts_as_unpaid": flags["unpaid"],
                "counts_as_shrinkage": flags["shrinkage"],
                "mapped": mapped, "evidence_type": evidence_type,
                "source_file": source_file, "rule_version": rulebook.version,
                "rule_sha256": rulebook.sha256,
            })
        scheduled_minutes = int((shift_end - shift_start).total_seconds() // 60)
        planned_net_minutes = min(scheduled_minutes, int(round(rulebook.standard_day_hours * 60)))
        totals = {
            flag: interval_minutes(shift_start, shift_end, intervals)
            for flag, intervals in flag_intervals.items()
        }
        # The configured standard day is a net payroll denominator.  Final
        # Activities may overlap or contain full-span assignments, so every
        # classified daily numerator is unioned above and capped again here.
        # This guarantees that no final rate can exceed 100% and mirrors the
        # governed observed-absence mart.
        absence_minutes = min(planned_net_minutes, totals.get("absence", 0))
        vacation_minutes = min(planned_net_minutes, totals.get("vacation", 0))
        unpaid_minutes = min(planned_net_minutes, totals.get("unpaid", 0))
        shrinkage_minutes = min(planned_net_minutes, totals.get("shrinkage", 0))
        unmapped_minutes = min(planned_net_minutes, totals.get("unmapped", 0))
        dimensions = {
            "lob": shift["roster_lob"],
            "language": shift["language"],
            "team_leader": shift["team_leader"],
        }
        final_absence = _metric_evaluation(
            metric_catalog,
            "final_absence_rate",
            shift["schedule_date"],
            dimensions,
            {
                "final_absence_minutes": absence_minutes,
                "final_vacation_minutes": vacation_minutes,
                "final_shrinkage_minutes": shrinkage_minutes,
                "planned_net_minutes": planned_net_minutes,
            },
        )
        attendance = attendance_by_key.get(agent_day_key)
        unresolved_observed_minutes = residual_by_key.get(agent_day_key, 0)
        final_code_present = bool(
            absence_minutes or vacation_minutes or unpaid_minutes
            or shrinkage_minutes or unmapped_minutes
        )
        expected_time_off_minutes = int(
            attendance.get("planning_overlay_minutes") or 0
        ) if attendance is not None else 0
        coded_time_off_minutes = max(
            absence_minutes, vacation_minutes, unpaid_minutes, shrinkage_minutes,
        )
        working_shift = shift.get("assignment_type") not in {"Off", "Planned absence"}
        shift_is_complete = bool(
            attendance.get("shift_state") == "COMPLETE"
            if attendance is not None else shift_end <= as_of
        )
        provisional_day = bool(working_shift and not shift_is_complete)
        missing_operational_evidence = bool(
            attendance is None
            or str(attendance.get("attendance_result") or "")
            in ({"No show"} | UNRELIABLE_ATTENDANCE_RESULTS)
        )
        uncoded_empty_shift = bool(
            working_shift and shift_is_complete
            and not final_code_present and missing_operational_evidence
        )
        uncorrected_observed_gap = bool(
            working_shift and shift_is_complete and not final_code_present
            and not uncoded_empty_shift and unresolved_observed_minutes > 0
        )
        partially_corrected_gap = bool(
            working_shift and shift_is_complete
            and final_code_present and unresolved_observed_minutes > 0
        )
        unsupported_verint_code = agent_day_key in verint_without_observed_gap
        day_rows.append({
            "agent_day_key": agent_day_key, "business_date": shift["schedule_date"],
            "agent_id": shift["agent_id"],
            "agent_name": shift["canonical_name"] or shift["agent_name"],
            "team_leader": shift["team_leader"], "ops_manager": shift["ops_manager"],
            "lob": shift["roster_lob"], "market": shift["market"],
            "language": shift["language"], "location": shift["location"],
            "scheduled_minutes": scheduled_minutes,
            "planned_net_minutes": planned_net_minutes,
            "final_absence_minutes": absence_minutes,
            "final_vacation_minutes": vacation_minutes,
            "final_unpaid_minutes": unpaid_minutes,
            "final_shrinkage_minutes": shrinkage_minutes,
            "final_unmapped_minutes": unmapped_minutes,
            "final_absence_hours": absence_minutes / 60,
            "final_absence_rate": final_absence.value,
            "final_absence_day": absence_minutes > 0,
            "final_ledger_status": (
                "UNMAPPED_REVIEW" if unmapped_minutes
                else "PROVISIONAL_DAY" if provisional_day
                else "PLANNED_TIME_OFF_NOT_IN_VERINT"
                if expected_time_off_minutes and not final_code_present
                else "TIME_OFF_PARTIALLY_IN_VERINT"
                if expected_time_off_minutes and coded_time_off_minutes < expected_time_off_minutes
                else "VERINT_WITHOUT_OBSERVED_GAP" if unsupported_verint_code
                else "UNCODED_EMPTY_SHIFT" if uncoded_empty_shift
                else "UNCORRECTED_OBSERVED_GAP" if uncorrected_observed_gap
                else "PARTIAL_CORRECTION_REVIEW" if partially_corrected_gap
                else "ABSENCE_RECORDED" if absence_minutes
                else "CLEAR"
            ),
            "rule_version": rulebook.version, "rule_sha256": rulebook.sha256,
        })
    conn.execute("DELETE FROM mart.verint_final_absence_event")
    conn.execute("DELETE FROM mart.verint_final_absence_agent_day")
    _insert_dicts(conn, "mart.verint_final_absence_event", FINAL_ABSENCE_EVENT_COLUMNS, event_rows)
    _insert_dicts(conn, "mart.verint_final_absence_agent_day", FINAL_ABSENCE_DAY_COLUMNS, day_rows)
    return len(event_rows), len(day_rows)


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
    conn: DatabaseConnection,
    start: date,
    end: date,
    mapping: QueueMapping,
    metric_catalog: MetricCatalog,
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
        dimensions = {
            "source_system": row["source_system"],
            "queue": row["queue"],
            "business_partner": row["business_partner"],
            "lob": row["lob"],
            "language": row["language"],
        }
        components = {
            "offered": offered,
            "answered": answered,
            "abandoned": abandoned,
            "short_abandoned": row["short_calls"],
            "answered_within_target": row["answered_20s"],
            "handled_seconds": (
                float(row["aht_seconds"]) * float(answered)
                if row["aht_seconds"] is not None and answered is not None else None
            ),
        }
        actual_rows.append({
            **row,
            "service_level_20s": _metric_evaluation(
                metric_catalog, "service_level_gross", row["business_date"], dimensions, components,
            ).value,
            "abandon_rate": _metric_evaluation(
                metric_catalog, "abandon_rate", row["business_date"], dimensions, components,
            ).value,
            "service_scope": mapped.service_scope, "comparison_scope": mapped.comparison_scope,
            "designation": mapped.designation,
            "mapping_status": mapped.status, "mapping_sha256": mapping.sha256,
        })
    _insert_dicts(conn, "mart.intraday_queue_interval", INTRADAY_COLUMNS, actual_rows)
    return len(forecast_rows), len(actual_rows)


def _build_pcs(
    conn: DatabaseConnection,
    config: Config,
    metric_catalog: MetricCatalog,
    start: date,
    end: date,
) -> int:
    """Aggregate deduplicated, FTE-scoped call legs to one agent/day."""
    conn.execute("DELETE FROM mart.agent_pcs_day")
    if not config.modules.get("pcs", True):
        return 0
    minimum = config.pcs.minimum_score
    maximum = config.pcs.maximum_score
    primary = config.pcs.primary_score_question
    participation = config.pcs.participation_question
    primary_score = f"question_{primary}_score"
    participation_answer = f"question_{participation}"
    allowed_scores = ", ".join(f"{value:g}" for value in config.pcs.allowed_scores)
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
                   CASE WHEN upper(coalesce(c.call_direction,''))='I' THEN 1 ELSE 0 END AS is_inbound,
                   CASE WHEN upper(coalesce(c.call_direction,''))='I'
                              AND coalesce(c.post_call_survey_mode,'')=? THEN 1 ELSE 0 END AS pcs_eligible,
                   CASE WHEN upper(coalesce(c.call_direction,''))='I'
                              AND coalesce(c.pcs_status,'')=? THEN 1 ELSE 0 END AS pcs_status_call,
                   CASE WHEN upper(coalesce(c.call_direction,''))='I'
                              AND coalesce(trim(c.{participation_answer}),'')<>'' THEN 1 ELSE 0 END AS participation_answered,
                   CASE WHEN upper(coalesce(c.call_direction,''))='I'
                              AND c.{primary_score} IN ({allowed_scores})
                        THEN c.{primary_score} END AS primary_score,
                   CASE WHEN upper(coalesce(c.call_direction,''))='I'
                              AND question_1_score IN ({allowed_scores}) THEN question_1_score END AS valid_q1,
                   CASE WHEN upper(coalesce(c.call_direction,''))='I'
                              AND question_2_score IN ({allowed_scores}) THEN question_2_score END AS valid_q2,
                   CASE WHEN {comment_test} THEN 1 ELSE 0 END AS has_comment
            FROM core.clean_call_leg c
            LEFT JOIN core.dim_agent d ON d.agent_id=c.agent_id
            WHERE c.business_date BETWEEN ? AND ? AND c.agent_id IS NOT NULL
              AND d.match_method='Agent ID'
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
                   sum(CASE WHEN coalesce(transferred,false) THEN 1 ELSE 0 END) AS transferred_legs,
                   sum(coalesce(talk_seconds,0)) AS talk_seconds,
                   sum(coalesce(hold_seconds,0)) AS hold_seconds,
                   sum(coalesce(wrap_seconds,0)) AS wrap_seconds,
                   sum(handle_seconds) AS handle_seconds,
                   sum(pcs_eligible) AS pcs_enabled_calls,
                   sum(CASE WHEN primary_score IS NOT NULL THEN 1 ELSE 0 END) AS survey_responses,
                   sum(pcs_status_call) AS pcs_status_calls,
                   sum(participation_answered) AS pcs_participation_responses,
                   sum(CASE WHEN participation_answered=1 AND primary_score IS NULL THEN 1 ELSE 0 END) AS pcs_invalid_responses,
                   sum(CASE WHEN pcs_status_call=1 AND participation_answered=0 THEN 1 ELSE 0 END) AS pcs_status_blank_responses,
                   sum(CASE WHEN participation_answered=1 AND pcs_status_call=0 THEN 1 ELSE 0 END) AS pcs_response_without_status,
                   sum(CASE WHEN valid_q1 IS NOT NULL THEN 1 ELSE 0 END) AS q1_response_count,
                   sum(coalesce(valid_q1,0)) AS q1_score_sum,
                   sum(CASE WHEN valid_q2 IS NOT NULL THEN 1 ELSE 0 END) AS q2_response_count,
                   sum(coalesce(valid_q2,0)) AS q2_score_sum,
                   sum(CASE WHEN primary_score IS NOT NULL THEN 1 ELSE 0 END) AS pcs_score_count,
                   sum(coalesce(primary_score,0)) AS pcs_score_sum,
                   sum(CASE WHEN primary_score > {config.pcs.negative_score_maximum:g} THEN 1 ELSE 0 END) AS top_box_responses,
                   sum(CASE WHEN primary_score <= {config.pcs.negative_score_maximum:g} THEN 1 ELSE 0 END) AS low_score_responses,
                   sum(CASE WHEN pcs_eligible=1 THEN has_comment ELSE 0 END) AS comments_count
            FROM prepared
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
            , transferred_legs, pcs_status_calls,
            pcs_participation_responses, pcs_participation_rate,
            pcs_invalid_responses, pcs_status_blank_responses,
            pcs_response_without_status
        )
        SELECT replace(business_date,'-','') || '-' || agent_id,
               business_date, agent_id, agent_name, team_leader, ops_manager,
               lob, market, language, location, call_legs, handled_calls,
               inbound_calls, outbound_calls, talk_seconds, hold_seconds,
               wrap_seconds, handle_seconds,
               NULL,
               NULL,
               NULL,
               NULL,
               pcs_enabled_calls, survey_responses,
               NULL,
               q1_response_count, q1_score_sum,
               NULL,
               q2_response_count, q2_score_sum,
               NULL,
               pcs_score_count, pcs_score_sum,
               NULL,
               top_box_responses, low_score_responses,
               NULL,
               NULL,
               comments_count, transferred_legs, pcs_status_calls,
               pcs_participation_responses,
               NULL,
               pcs_invalid_responses, pcs_status_blank_responses,
               pcs_response_without_status
        FROM aggregated
    """
    conn.execute(
        sql,
        [config.pcs.survey_mode, config.pcs.participation_status, start, end],
    )
    rows = _dicts(conn.execute(
        """SELECT agent_day_key, business_date, team_leader, lob, language,
                  talk_seconds, hold_seconds, wrap_seconds, handle_seconds,
                  handled_calls, q1_score_sum, q1_response_count,
                  q2_score_sum, q2_response_count, pcs_score_sum, pcs_score_count,
                  pcs_participation_responses, pcs_status_calls,
                  top_box_responses, low_score_responses, survey_responses
           FROM mart.agent_pcs_day"""
    ))
    for row in rows:
        dimensions = {
            "team_leader": row["team_leader"],
            "lob": row["lob"],
            "language": row["language"],
        }
        components = {
            "talk_seconds": row["talk_seconds"],
            "hold_seconds": row["hold_seconds"],
            "wrap_seconds": row["wrap_seconds"],
            "handle_seconds": row["handle_seconds"],
            "handled_calls": row["handled_calls"],
            "q1_score_sum": row["q1_score_sum"],
            "q1_response_count": row["q1_response_count"],
            "q2_score_sum": row["q2_score_sum"],
            "q2_response_count": row["q2_response_count"],
            "pcs_score_sum": row["pcs_score_sum"],
            "pcs_score_count": row["pcs_score_count"],
            "pcs_participation_responses": row["pcs_participation_responses"],
            "pcs_status_calls": row["pcs_status_calls"],
            "top_box_responses": row["top_box_responses"],
            "low_score_responses": row["low_score_responses"],
            "survey_responses": row["survey_responses"],
        }
        aht = _metric_evaluation(
            metric_catalog, "agent_aht_seconds", row["business_date"], dimensions, components,
        ).value
        average_talk = _optional_metric_value(
            metric_catalog, "agent_talk_seconds", row["business_date"], dimensions, components,
        )
        average_hold = _optional_metric_value(
            metric_catalog, "agent_hold_seconds", row["business_date"], dimensions, components,
        )
        average_wrap = _optional_metric_value(
            metric_catalog, "agent_wrap_seconds", row["business_date"], dimensions, components,
        )
        q1_average = _optional_metric_value(
            metric_catalog, "pcs_q1_average", row["business_date"], dimensions, components,
        )
        q2_average = _optional_metric_value(
            metric_catalog, "pcs_q2_average", row["business_date"], dimensions, components,
        )
        average = _metric_evaluation(
            metric_catalog, "pcs_average", row["business_date"], dimensions, components,
        ).value
        participation = _metric_evaluation(
            metric_catalog, "pcs_participation", row["business_date"], dimensions, components,
        ).value
        positive = _metric_evaluation(
            metric_catalog, "pcs_positive_rate", row["business_date"], dimensions, components,
        ).value
        negative = _metric_evaluation(
            metric_catalog, "pcs_negative_rate", row["business_date"], dimensions, components,
        ).value
        conn.execute(
            """UPDATE mart.agent_pcs_day
               SET average_talk_seconds=?, average_hold_seconds=?, average_wrap_seconds=?,
                   average_handle_seconds=?, response_rate=?, q1_average=?, q2_average=?, pcs_average=?,
                   top_box_percent=?, low_score_percent=?, pcs_participation_rate=?
               WHERE agent_day_key=?""",
            [average_talk, average_hold, average_wrap, aht, participation,
             q1_average, q2_average, average, positive, negative, participation,
             row["agent_day_key"]],
        )
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
    metric_catalog: MetricCatalog,
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
        dimensions = {
            "lob": base["lob"],
            "language": base["language"],
            "team_leader": base["team_leader"],
        }
        components = {
            "absence_minutes": absence,
            "vacation_minutes": vacation,
            "shrinkage_minutes": shrinkage,
            "planned_net_minutes": planned_net,
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
            "absence_rate": _metric_evaluation(
                metric_catalog, "observed_absence_rate", base["business_date"], dimensions, components,
            ).value,
            "vacation_rate": _metric_evaluation(
                metric_catalog, "observed_vacation_rate", base["business_date"], dimensions, components,
            ).value,
            "shrinkage_rate": _metric_evaluation(
                metric_catalog, "observed_shrinkage_rate", base["business_date"], dimensions, components,
            ).value,
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
    "sl_target", "sl_state",
]


def _build_service(
    conn: DatabaseConnection,
    rulebook: Rulebook,
    metric_catalog: MetricCatalog,
    mapping: QueueMapping,
    start: date,
    end: date,
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
    for row in rows:
        short_abandoned = row["abandoned_20s"] if row["abandoned_20s"] is not None else row["short_calls"]
        handled_seconds = (
            float(row["aht_seconds"]) * float(row["answered"])
            if row["aht_seconds"] is not None and row["answered"] is not None else None
        )
        components = {
            "offered": row["offered"], "answered": row["answered"],
            "abandoned": row["abandoned"], "short_abandoned": short_abandoned,
            "answered_within_target": row["answered_20s"], "handled_seconds": handled_seconds,
        }
        mapped = mapping.map_actual(row["source_system"], row["queue"], row["business_partner"], row["lob"])
        dimensions = {
            "source_system": row["source_system"],
            "queue": row["queue"],
            "business_partner": row["business_partner"],
            "lob": row["lob"],
            "language": row["language"],
        }
        service_level = _metric_evaluation(
            metric_catalog, "service_level", row["business_date"], dimensions, components,
        )
        gross = _metric_evaluation(
            metric_catalog, "service_level_gross", row["business_date"], dimensions, components,
        )
        availability = _metric_evaluation(
            metric_catalog, "service_availability", row["business_date"], dimensions, components,
        )
        abandon = _metric_evaluation(
            metric_catalog, "abandon_rate", row["business_date"], dimensions, components,
        )
        aht = _metric_evaluation(
            metric_catalog, "aht_seconds", row["business_date"], dimensions, components,
        )
        output.append({
            **{key: row[key] for key in (
                "business_date", "interval_start", "hour_start", "source_system", "queue",
                "business_partner", "lob", "language", "offered", "answered", "abandoned",
                "source_file",
            )},
            "short_abandoned": short_abandoned, "answered_within_target": row["answered_20s"],
            "handled_seconds": handled_seconds,
            "sl_gross": gross.value,
            "sl_adjusted": service_level.value,
            "sl_profile": service_level.method.method_id,
            "service_level": service_level.value,
            "service_availability": availability.value,
            "abandon_rate": abandon.value,
            "aht_seconds": aht.value,
            "rule_version": rulebook.version, "rule_sha256": rulebook.sha256,
            "service_scope": mapped.service_scope, "comparison_scope": mapped.comparison_scope,
            "designation": mapped.designation,
            "mapping_status": mapped.status, "mapping_sha256": mapping.sha256,
            "sl_target": service_level.method.target,
            "sl_state": service_level.state,
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
    schedule_variants = {
        row[0] for row in conn.execute(
            """SELECT DISTINCT source_variant FROM meta.source_file
               WHERE source_family='schedule' AND active=true AND status='SUCCESS'"""
        ).fetchall() if row[0]
    }
    if "START_END" not in schedule_variants:
        add(
            "schedule", str(config.source_path("schedule_folder")), None, None,
            "Missing StartEndTimes schedule", "ERROR",
            "Operational attendance requires a StartEndTimes export. Activities is the corrected final ledger and cannot replace it.",
        )
    if "ACTIVITIES" not in schedule_variants:
        add(
            "schedule", str(config.source_path("schedule_folder")), None, None,
            "Missing Activities final ledger", "REVIEW",
            "Final Verint absenteeism and correction reconciliation require an Activities export.",
        )
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
               WHERE planned_work_minutes>0 AND status_source IS NOT NULL
                 AND 1.0*status_covered_minutes/planned_work_minutes < ?
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
            """SELECT business_date, agent_id, survey_responses,
                      low_score_responses, top_box_responses,
                      pcs_participation_responses, pcs_invalid_responses,
                      pcs_status_calls, pcs_response_without_status
               FROM mart.agent_pcs_day
               WHERE business_date BETWEEN ? AND ? AND (
                   survey_responses<>low_score_responses+top_box_responses OR
                   pcs_participation_responses<>survey_responses+pcs_invalid_responses OR
                   pcs_participation_responses>pcs_status_calls OR
                   pcs_response_without_status>0
               )""", [start, end]
        )):
            details = (
                f"valid={row['survey_responses']}, <=3={row['low_score_responses']}, "
                f">3={row['top_box_responses']}, raw Q1={row['pcs_participation_responses']}, "
                f"invalid={row['pcs_invalid_responses']}, PCSStatus=1={row['pcs_status_calls']}, "
                f"Q1 without status={row['pcs_response_without_status']}"
            )
            add(
                "calls", None, row["business_date"], row["agent_id"],
                "PCS counter reconciliation", "REVIEW", details,
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
        """SELECT business_date, agent_id, agent_name, final_ledger_status
           FROM mart.verint_final_absence_agent_day
           WHERE business_date BETWEEN ? AND ?
             AND final_ledger_status IN (
               'PLANNED_TIME_OFF_NOT_IN_VERINT','TIME_OFF_PARTIALLY_IN_VERINT'
             )""", [start, end]
    )):
        add(
            "fte", None, row["business_date"], row["agent_id"],
            "Planned time off not final in Verint", "ERROR",
            f"{row['agent_name'] or row['agent_id']}: {row['final_ledger_status']}. "
            "Correct Verint Activities, export again, and refresh before payroll use.",
        )
    for row in _dicts(conn.execute(
        """SELECT r.agent_id, r.source_sheet, r.source_row, r.start_date,
                  r.absence_type, f.file_name
           FROM raw.fte_time_off r
           JOIN meta.source_file f ON f.file_id=r.source_file_id
           LEFT JOIN core.dim_agent d ON d.agent_id=r.agent_id
           WHERE f.active=true AND f.status='SUCCESS'
             AND coalesce(d.match_method,'Unmatched to FTE')<>'Agent ID'"""
    )):
        add(
            "fte", row["file_name"], row["start_date"], row["agent_id"],
            "Time-off Agent ID not in active roster", "ERROR",
            f"{row['source_sheet']} row {row['source_row']} ({row['absence_type']}) "
            "does not match an admitted Agent-sheet ID.",
        )
    for row in _dicts(conn.execute(
        """SELECT business_date, agent_id, activity, source_file, sum(minutes) AS minutes
           FROM mart.verint_final_absence_event
           WHERE mapped=false AND business_date BETWEEN ? AND ?
           GROUP BY business_date, agent_id, activity, source_file""", [start, end]
    )):
        add(
            "schedule", row["source_file"], row["business_date"], row["agent_id"],
            "Unmapped final Verint activity", "REVIEW",
            f"{row['activity'] or '(blank)'} covers {row['minutes']} minutes and needs a rulebook classification.",
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
    as_of: datetime | None = None,
) -> ModelSummary:
    total_stages = 22

    def stage(completed: int, label: str) -> None:
        if progress is not None:
            progress(completed, total_stages, label)

    conn.execute("SAVEPOINT refresh_models")
    try:
        stage(0, "Validating business rules")
        rulebook = load_rulebook(config.home, config.business_rules)
        metric_catalog = load_metric_catalog(config.home, config.metric_catalog)
        validate_metric_catalog(metric_catalog, SOURCE_COMPONENTS)
        analytics_rules = load_analytics_rules(config.home, config.analytics_rules)
        validate_analytics_rules(analytics_rules, metric_catalog)
        mapping = load_queue_mapping(config.queue_mapping)
        evaluation_as_of = _evaluation_time(config.timezone, as_of)
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
        planned_time_off = _build_planned_time_off(
            conn, schedules, agents, start, end, evaluation_as_of,
        )
        stage(7, "Building attendance")
        attendance = _build_attendance(
            conn, rulebook, schedules, events_by_agent, lilo, loaded_dates, seen_ids,
            agents, statuses, status_loaded_dates, planned_time_off, evaluation_as_of,
            config.rules.minimum_status_coverage,
        )
        stage(8, "Keeping adherence disabled")
        conn.execute("DELETE FROM mart.conformance_agent_day")
        conformance = []
        stage(9, "Finding observed LILO and status gaps")
        corrections = _build_corrections(conn, rulebook, attendance)
        stage(10, "Building LOB and language staffing intervals")
        staffing = _build_staffing(conn, attendance, evaluation_as_of)
        stage(11, "Building shift evidence timelines")
        timeline = _build_shift_timeline(conn, attendance, evaluation_as_of)
        stage(12, "Keeping legacy RTA disabled")
        conn.execute("DELETE FROM mart.rta_snapshot")
        rta = []
        stage(13, "Building intraday actual and forecast")
        forecast, actual = _build_intraday(conn, start, end, mapping, metric_catalog)
        stage(14, "Building exact PCS counters")
        # PCS management always needs current MTD and the previous full month,
        # even when the user asks for Today or Current Week. Other operational
        # marts retain the exact selected boundary.
        previous_month_end = end.replace(day=1) - timedelta(days=1)
        pcs_start = min(start, previous_month_end.replace(day=1))
        pcs = _build_pcs(conn, config, metric_catalog, pcs_start, end)
        stage(15, "Building observed absence and shrinkage")
        absence, absence_events = _build_absence(
            conn, config, rulebook, metric_catalog, attendance, corrections,
        )
        stage(16, "Building corrected Verint final absence")
        final_absence_events, final_absence = _build_verint_final_absence(
            conn, rulebook, metric_catalog, start, end, evaluation_as_of,
        )
        stage(17, "Building service performance")
        service = _build_service(conn, rulebook, metric_catalog, mapping, start, end)
        _record_rule_application(conn, run_id, rulebook)
        _record_mapping_application(conn, run_id, mapping)
        stage(18, "Checking source health")
        _build_source_health(conn, config)
        stage(19, "Running data-quality checks")
        quality = _build_quality(conn, config, run_id, start, end)
        stage(20, "Materializing configured KPI values")
        metric_rows = build_metric_values(
            conn, metric_catalog, rulebook, run_id, start, end,
        )
        stage(21, "Generating deterministic Python findings")
        finding_rows = build_findings(
            conn, metric_catalog, analytics_rules, run_id, start, end,
        )
        result = ModelSummary(
            start=start, end=end, attendance_rows=len(attendance), conformance_rows=len(conformance),
            correction_rows=len(corrections), rta_rows=len(rta), forecast_rows=forecast,
            intraday_rows=actual, pcs_rows=pcs, quality_rows=quality,
            absence_rows=absence, absence_event_rows=absence_events, service_rows=service,
            staffing_rows=staffing, timeline_rows=timeline,
            final_absence_rows=final_absence,
            final_absence_event_rows=final_absence_events,
            metric_rows=metric_rows, finding_rows=finding_rows,
        )
        conn.execute("RELEASE SAVEPOINT refresh_models")
        stage(total_stages, "Models ready")
        return result
    except Exception:
        conn.execute("ROLLBACK TO SAVEPOINT refresh_models")
        conn.execute("RELEASE SAVEPOINT refresh_models")
        raise
