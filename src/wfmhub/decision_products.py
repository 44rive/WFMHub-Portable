"""Focused WFM/Operations report products using one workbook design contract."""

from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any, Iterable, Sequence

from openpyxl import load_workbook

from .capacity_mapping import load_capacity_mapping
from .config import Config
from .database import DatabaseConnection
from .excel_layout import (
    ACTION_FIRST_ROW,
    ACTION_HEADER_ROW,
    ACTION_VISIBLE_ROWS,
    V2ChartSpec,
    configure_v2_cell_canvas,
    insert_v2_charts,
    make_v2_formats,
    render_v2_dashboard,
    style_v2_chart,
    write_v2_filters,
    write_v2_header,
    write_v2_kpis,
    write_v2_section,
)
from .mapping import load_queue_mapping
from .metrics import MetricCatalog, evaluate_metric, load_metric_catalog
from .report_packs import publish_report, report_current_path
from .reports import COLORS, _query
from .rules import load_rulebook
from .service_profiles import ServiceProfile, load_service_profiles
from .template_reports import DecisionWorkbook, KpiCard, ModelTable
from .utils import merge_intervals


@dataclass(frozen=True)
class NamedPeriod:
    label: str
    start: date
    end: date


def _month_start(value: date) -> date:
    return value.replace(day=1)


def _previous_month(value: date) -> tuple[date, date]:
    end = value.replace(day=1) - timedelta(days=1)
    return end.replace(day=1), end


def _comparison_periods(start: date, end: date) -> list[NamedPeriod]:
    previous_start, previous_end = _previous_month(end)
    prior_mtd_end = min(previous_end, previous_start + timedelta(days=end.day - 1))
    duration = (end - start).days + 1
    prior_equal_end = start - timedelta(days=1)
    prior_equal_start = prior_equal_end - timedelta(days=duration - 1)
    periods = [
        NamedPeriod("Latest day", end, end),
        NamedPeriod("Selected period", start, end),
        NamedPeriod("Current MTD", _month_start(end), end),
        NamedPeriod("Prior-month same days", previous_start, prior_mtd_end),
        NamedPeriod("Previous full month", previous_start, previous_end),
    ]
    if start != _month_start(end):
        periods.insert(2, NamedPeriod("Previous equal period", prior_equal_start, prior_equal_end))
    return periods


HOUR_BUCKETS = (
    "available", "inbound", "outbound", "bo", "productive_other",
    "management", "support", "training", "it_incident", "other_aux",
    "break", "lunch", "logged_off", "lilo_only", "unknown", "future",
    "pto", "away",
)


def _timeline_hour_bucket(actual_category: Any, actual_status: Any, rulebook: Any) -> str:
    """Put every exclusive shift-timeline segment in exactly one hour bucket."""

    category = str(actual_category or "").strip().casefold()
    if category in {"pto", "away"}:
        return category
    if category == "future":
        return "future"
    if category in {"no_activity", "logged off"}:
        return "logged_off"
    if category in {"no_status_evidence", "unavailable"}:
        return "unknown" if category == "no_status_evidence" else "other_aux"
    if category == "lilo_present":
        return "lilo_only"
    if category in {"break", "lunch"}:
        return category
    rule = rulebook.classify_status(actual_status)
    aux = str(rule.aux_classification if rule else "").strip().casefold()
    if category == "productive":
        return {
            "available": "available", "inbound": "inbound",
            "outbound": "outbound", "bo": "bo",
        }.get(aux, "productive_other")
    if category == "auxiliary":
        return {
            "management": "management", "support": "support",
            "training": "training", "it incident": "it_incident",
        }.get(aux, "other_aux")
    return "unknown"


def _realisations_hours(
    conn: DatabaseConnection,
    profiles: Sequence[ServiceProfile],
    start: date,
    end: date,
    rulebook: Any,
) -> tuple[list[str], list[tuple[Any, ...]], list[tuple[Any, ...]]]:
    """Return reconciled agent and LOB hours from exclusive timeline segments."""

    lob_to_profile = {
        lob: profile.label for profile in profiles for lob in profile.staffing_lobs
    }
    headers = [
        "business_date", "reporting_lob", "workforce_lob", "agent_id",
        "agent_name", "team_leader", "scheduled_hours",
        *[f"{name}_hours" for name in HOUR_BUCKETS],
        "allocated_hours", "reconciliation_delta_hours",
    ]
    if not lob_to_profile:
        return headers, [], []
    marks = ",".join("?" for _ in lob_to_profile)
    cursor = conn.execute(
        f"""SELECT business_date, lob, agent_id, agent_name, team_leader,
                   segment_minutes, actual_status, actual_category
            FROM mart.shift_timeline_segment
            WHERE business_date BETWEEN ? AND ? AND lob IN ({marks})
            ORDER BY business_date, lob, agent_id, segment_start""",
        [start, end, *lob_to_profile],
    )
    agents: dict[tuple[str, str, str], dict[str, Any]] = {}
    for business_date, lob, agent_id, name, leader, minutes, status, category in cursor.fetchall():
        key = (str(business_date)[:10], str(lob), str(agent_id))
        item = agents.setdefault(key, {
            "name": name, "leader": leader, "scheduled": 0,
            "buckets": defaultdict(int),
        })
        duration = max(0, int(minutes or 0))
        item["scheduled"] += duration
        item["buckets"][_timeline_hour_bucket(category, status, rulebook)] += duration
    attendance = conn.execute(
        f"""SELECT business_date, lob, agent_id, agent_name, team_leader,
                   scheduled_minutes, scheduled_start, scheduled_end,
                   attendance_result, assignment_type
            FROM mart.attendance_agent_day
            WHERE business_date BETWEEN ? AND ? AND lob IN ({marks})
            ORDER BY business_date, lob, agent_id""",
        [start, end, *lob_to_profile],
    ).fetchall()
    for business_date, lob, agent_id, name, leader, scheduled_minutes, shift_start, shift_end, result, assignment in attendance:
        key = (str(business_date)[:10], str(lob), str(agent_id))
        if str(assignment or "") == "Off":
            continue
        item = agents.setdefault(key, {
            "name": name, "leader": leader, "scheduled": 0,
            "buckets": defaultdict(int),
        })
        item["name"] = item["name"] or name
        item["leader"] = item["leader"] or leader
        published = int(scheduled_minutes or 0)
        if published <= 0 and shift_start is not None and shift_end is not None:
            left = shift_start if isinstance(shift_start, datetime) else datetime.fromisoformat(str(shift_start))
            right = shift_end if isinstance(shift_end, datetime) else datetime.fromisoformat(str(shift_end))
            published = max(0, int((right - left).total_seconds() // 60))
        if published <= 0:
            continue
        observed = item["scheduled"]
        if published > observed:
            result_text = str(result or "").casefold()
            fallback = "pto" if result_text == "pto" else "away" if result_text == "away" else "unknown"
            item["buckets"][fallback] += published - observed
        item["scheduled"] = published
    agent_rows: list[tuple[Any, ...]] = []
    lob_totals: dict[tuple[str, str, str], dict[str, Any]] = {}
    for (business_date, lob, agent_id), item in sorted(agents.items()):
        allocated = sum(item["buckets"].values())
        profile_label = lob_to_profile[lob]
        values = [
            business_date, profile_label, lob, agent_id,
            item["name"], item["leader"], item["scheduled"] / 60,
            *[item["buckets"].get(bucket, 0) / 60 for bucket in HOUR_BUCKETS],
            allocated / 60, (item["scheduled"] - allocated) / 60,
        ]
        agent_rows.append(tuple(values))
        lob_key = (business_date, profile_label, lob)
        total = lob_totals.setdefault(lob_key, {
            "scheduled": 0, "buckets": defaultdict(int),
        })
        total["scheduled"] += item["scheduled"]
        for bucket, amount in item["buckets"].items():
            total["buckets"][bucket] += amount
    lob_rows = []
    for (business_date, profile_label, lob), item in sorted(lob_totals.items()):
        allocated = sum(item["buckets"].values())
        lob_rows.append((
            business_date, profile_label, lob, None, None, None,
            item["scheduled"] / 60,
            *[item["buckets"].get(bucket, 0) / 60 for bucket in HOUR_BUCKETS],
            allocated / 60, (item["scheduled"] - allocated) / 60,
        ))
    return headers, agent_rows, lob_rows


def _realisations_aux_breakdown(
    conn: DatabaseConnection,
    profiles: Sequence[ServiceProfile],
    start: date,
    end: date,
    rulebook: Any,
) -> tuple[list[str], list[tuple[Any, ...]], list[str], list[tuple[Any, ...]]]:
    """Break down exact Verint Agent Status AUX inside published shifts."""

    lob_to_profile = {
        lob: profile.label for profile in profiles for lob in profile.staffing_lobs
    }
    agent_headers = [
        "business_date", "reporting_lob", "workforce_lob", "team_leader",
        "agent_id", "agent_name", "aux_class", "exact_status", "hours",
    ]
    tl_headers = [
        "business_date", "reporting_lob", "workforce_lob", "team_leader",
        "aux_class", "exact_status", "agents", "hours",
    ]
    if not lob_to_profile:
        return agent_headers, [], tl_headers, []
    marks = ",".join("?" for _ in lob_to_profile)
    rows = conn.execute(
        f"""SELECT business_date, lob, agent_id, agent_name, team_leader,
                   actual_status, actual_category, segment_minutes
            FROM mart.shift_timeline_segment
            WHERE business_date BETWEEN ? AND ? AND lob IN ({marks})
              AND observed_source='AGENT_STATUS'
            ORDER BY business_date, lob, team_leader, agent_name, segment_start""",
        [start, end, *lob_to_profile],
    ).fetchall()
    agent_totals: dict[tuple[Any, ...], int] = defaultdict(int)
    tl_totals: dict[tuple[Any, ...], dict[str, Any]] = {}
    for day, lob, agent_id, name, leader, status, category, minutes in rows:
        rule = rulebook.classify_status(status)
        aux = str(rule.aux_classification if rule else "").strip()
        category_name = str(category or "").strip().casefold()
        if aux.casefold() in {"", "available", "inbound", "outbound"}:
            if category_name not in {"auxiliary", "break", "lunch", "unavailable"}:
                continue
            aux = "Other AUX" if category_name in {"auxiliary", "unavailable"} else category_name.title()
        exact_status = str(status or "Unknown AUX").strip()
        key = (str(day)[:10], lob_to_profile[lob], lob, leader or "Unassigned",
               str(agent_id), name or str(agent_id), aux, exact_status)
        amount = max(0, int(minutes or 0))
        agent_totals[key] += amount
        tl_key = key[:4] + key[6:]
        item = tl_totals.setdefault(tl_key, {"agents": set(), "minutes": 0})
        item["agents"].add(str(agent_id))
        item["minutes"] += amount
    agent_rows = [(*key, amount / 60) for key, amount in sorted(agent_totals.items())]
    tl_rows = [(*key, len(item["agents"]), item["minutes"] / 60)
               for key, item in sorted(tl_totals.items())]
    return agent_headers, agent_rows, tl_headers, tl_rows


def _realisations_final_absence(
    conn: DatabaseConnection,
    profiles: Sequence[ServiceProfile],
    start: date,
    end: date,
) -> tuple[dict[tuple[str, str], tuple[float, float, float, float, int]],
           list[str], list[tuple[Any, ...]], list[str], list[tuple[Any, ...]]]:
    """Final Activities absence; exclude UNPAID_LEAVE but retain No Show."""

    lob_to_profile = {
        lob: profile.label for profile in profiles for lob in profile.staffing_lobs
    }
    agent_headers = [
        "business_date", "reporting_lob", "workforce_lob", "team_leader",
        "agent_id", "agent_name", "planned_net_hours", "absence_hours",
        "absence_rate", "unpaid_leave_hours", "vacation_hours",
        "shrinkage_hours", "ledger_status",
    ]
    lob_headers = [
        "business_date", "reporting_lob", "planned_net_hours",
        "absence_hours", "absence_rate", "unpaid_leave_hours",
        "vacation_hours", "shrinkage_hours", "review_cases",
    ]
    if not lob_to_profile:
        return {}, agent_headers, [], lob_headers, []
    marks = ",".join("?" for _ in lob_to_profile)
    event_rows = conn.execute(
        f"""SELECT agent_day_key, category, counts_as_absence,
                   event_start, event_end
            FROM mart.verint_final_absence_event
            WHERE business_date BETWEEN ? AND ? AND lob IN ({marks})""",
        [start, end, *lob_to_profile],
    ).fetchall()
    intervals: dict[str, list[tuple[datetime, datetime]]] = defaultdict(list)
    unpaid_intervals: dict[str, list[tuple[datetime, datetime]]] = defaultdict(list)
    for key, category, counts, event_start, event_end in event_rows:
        left = event_start if isinstance(event_start, datetime) else datetime.fromisoformat(str(event_start))
        right = event_end if isinstance(event_end, datetime) else datetime.fromisoformat(str(event_end))
        if right > left:
            if str(category or "").upper() == "UNPAID_LEAVE":
                unpaid_intervals[str(key)].append((left, right))
            elif counts:
                intervals[str(key)].append((left, right))
    day_rows = conn.execute(
        f"""SELECT agent_day_key, business_date, lob, team_leader,
                   agent_id, agent_name, planned_net_minutes,
                   final_vacation_minutes, final_shrinkage_minutes,
                   final_ledger_status
            FROM mart.verint_final_absence_agent_day
            WHERE business_date BETWEEN ? AND ? AND lob IN ({marks})
            ORDER BY business_date, lob, team_leader, agent_name""",
        [start, end, *lob_to_profile],
    ).fetchall()
    agent_rows: list[tuple[Any, ...]] = []
    totals: dict[tuple[str, str], dict[str, float | int]] = defaultdict(
        lambda: {"planned": 0, "absence": 0, "unpaid": 0,
                 "vacation": 0, "shrinkage": 0, "review": 0}
    )
    for key, day, lob, leader, agent_id, name, planned, vacation, shrinkage, status in day_rows:
        planned_minutes = max(0, int(planned or 0))
        merged = merge_intervals(intervals.get(str(key), ()))
        absence_minutes = min(planned_minutes, int(sum((right - left).total_seconds()
                                                        for left, right in merged) // 60))
        # Unpaid leave has a separate diagnostic column; it is not deducted
        # from No Show, which may also carry the source's unpaid flag.
        unpaid_minutes = int(sum((right - left).total_seconds()
                                 for left, right in merge_intervals(unpaid_intervals.get(str(key), ()))) // 60)
        reporting_lob = lob_to_profile[lob]
        day_key = (str(day)[:10], reporting_lob)
        review = str(status or "") not in {"CLEAR", "ABSENCE_RECORDED"}
        item = totals[day_key]
        for field, amount in (
            ("planned", planned_minutes), ("absence", absence_minutes),
            ("unpaid", min(planned_minutes, unpaid_minutes)),
            ("vacation", int(vacation or 0)), ("shrinkage", int(shrinkage or 0)),
        ):
            item[field] += amount
        item["review"] += int(review)
        agent_rows.append((
            str(day)[:10], reporting_lob, lob, leader, str(agent_id), name,
            planned_minutes / 60, absence_minutes / 60,
            absence_minutes / planned_minutes if planned_minutes else None,
            min(planned_minutes, unpaid_minutes) / 60,
            float(vacation or 0) / 60, float(shrinkage or 0) / 60, status,
        ))
    lob_rows = [
        (day, lob, item["planned"] / 60, item["absence"] / 60,
         item["absence"] / item["planned"] if item["planned"] else None,
         item["unpaid"] / 60, item["vacation"] / 60,
         item["shrinkage"] / 60, item["review"])
        for (day, lob), item in sorted(totals.items())
    ]
    daily = {
        key: (item["planned"] / 60, item["absence"] / 60,
              item["vacation"] / 60, item["shrinkage"] / 60,
              int(item["review"]))
        for key, item in totals.items()
    }
    return daily, agent_headers, agent_rows, lob_headers, lob_rows


def _ratio(numerator: float | int | None, denominator: float | int | None) -> float | None:
    return float(numerator) / float(denominator) if denominator else None


def _delta(value: float | None, reference: float | None, percent: bool = False) -> str:
    if value is None or reference is None:
        return "Comparison not available"
    difference = value - reference
    return f"{difference:+.1%} vs prior" if percent else f"{difference:+.2f} vs prior"


def _output_path(config: Config, key: str, start: date, end: date, generated: datetime, output: Path | None) -> Path:
    del start, end, generated
    return (output or report_current_path(config, key)).resolve()


def _source_state(conn: DatabaseConnection, families: Sequence[str], through: date, final: bool = False) -> tuple[str, str]:
    if not families:
        return ("FINAL" if final else "LIVE"), "All required datasets are available"
    placeholders = ",".join("?" for _ in families)
    rows = conn.execute(
        f"""SELECT source_family, newest_business_date, status
            FROM mart.source_health WHERE source_family IN ({placeholders})""",
        list(families),
    ).fetchall()
    found = {str(row[0]).lower(): row for row in rows}
    problems = []
    for family in families:
        if family.casefold() in {"start_end", "activities"}:
            variant = "START_END" if family.casefold() == "start_end" else "ACTIVITIES"
            variant_row = conn.execute(
                """SELECT max(r.schedule_date), count(*)
                   FROM raw.schedule_shift r
                   JOIN meta.source_file f ON f.file_id=r.source_file_id
                   WHERE f.active=true AND f.status='SUCCESS' AND f.source_variant=?""",
                [variant],
            ).fetchone()
            variant_date, variant_count = variant_row
            if family.casefold() == "start_end" and not variant_count:
                variant = "ACTIVITIES"
                variant_row = conn.execute(
                    """SELECT max(r.schedule_date), count(*)
                       FROM raw.schedule_shift r
                       JOIN meta.source_file f ON f.file_id=r.source_file_id
                       WHERE f.active=true AND f.status='SUCCESS'
                         AND f.source_variant='ACTIVITIES' AND r.parse_ok=true
                         AND r.scheduled_start IS NOT NULL
                         AND r.scheduled_end IS NOT NULL"""
                ).fetchone()
                variant_date, variant_count = variant_row
            if not variant_count:
                problems.append(f"{family}: no successful {variant} extract")
            elif variant_date is None or str(variant_date)[:10] < through.isoformat():
                problems.append(f"{family}: latest {variant_date or 'unknown'}")
            continue
        row = found.get(family.lower())
        if row is None:
            problems.append(f"{family}: no health record")
        elif str(row[2]).upper() != "SUCCESS":
            problems.append(f"{family}: {row[2]}")
        # FTE is a point-in-time scope roster rather than a dated fact source.
        # Its freshness is represented by load status and source hash, so a
        # NULL business date is expected and must not make every report red.
        elif family.casefold() != "fte" and (row[1] is None or str(row[1])[:10] < through.isoformat()):
            problems.append(f"{family}: latest {row[1] or 'unknown'}")
    if problems:
        return "INCOMPLETE", "; ".join(problems)
    if final:
        return "FINAL", f"Required sources loaded through {through}"
    if through >= date.today():
        return "PROVISIONAL", "Current-day values can still change before shifts and queues close"
    return "LIVE", f"Required sources loaded through {through}"


def _audit_rows(
    conn: DatabaseConnection,
    config: Config,
    report_key: str,
    start: date,
    end: date,
    extra: Iterable[Sequence[Any]] = (),
) -> list[Sequence[Any]]:
    latest = conn.execute(
        "SELECT run_id, finished_at, details FROM meta.refresh_run WHERE status='SUCCESS' ORDER BY finished_at DESC LIMIT 1"
    ).fetchone()
    rows: list[Sequence[Any]] = [
        ("Report", report_key, "WFM report product"),
        ("Selected period", f"{start} to {end}", "Dates included"),
        ("Report generated at", datetime.now(), "Local work-machine time"),
        (
            "Last successful Hub refresh",
            latest[1] if latest else None,
            latest[2] if latest else "No successful refresh metadata",
        ),
        ("Data through", end, "Latest selected business date"),
        ("Refresh run ID", latest[0] if latest else None, "Database lineage"),
        ("Prepared by", "Anass ASSRI", "WFM"),
    ]
    rows.extend(extra)
    return rows


def _atomic_book(
    config: Config,
    key: str,
    title: str,
    start: date,
    end: date,
    output: Path | None,
) -> tuple[DecisionWorkbook, Path, Path]:
    generated = datetime.now()
    target = _output_path(config, key, start, end, generated, output)
    target.parent.mkdir(parents=True, exist_ok=True)
    partial = target.with_name(f"{target.stem}.partial{target.suffix}")
    return DecisionWorkbook(partial, config, key, title, start, end, generated), partial, target


def _finish(book: DecisionWorkbook, partial: Path, target: Path) -> Path:
    try:
        book.close()
        publish_report(book.config, book.report_key, partial, target, book.generated)
    except Exception:
        partial.unlink(missing_ok=True)
        raise
    return target


def _load_persistent_rows(
    path: Path,
    sheet_name: str,
    key_header: str,
) -> dict[str, dict[str, Any]]:
    """Read a prior editable ledger without making report generation depend on it."""

    if not path.is_file():
        return {}
    try:
        workbook = load_workbook(path, read_only=True, data_only=False)
        if sheet_name not in workbook.sheetnames:
            workbook.close()
            return {}
        sheet = workbook[sheet_name]
        headers = [str(cell.value or "") for cell in sheet[4]]
        key_index = headers.index(key_header)
        rows: dict[str, dict[str, Any]] = {}
        for values in sheet.iter_rows(min_row=5, values_only=True):
            if not any(value is not None and str(value).strip() for value in values):
                continue
            key = str(values[key_index] or "").strip()
            if key:
                rows[key] = {
                    header: values[index] if index < len(values) else None
                    for index, header in enumerate(headers)
                }
        workbook.close()
        return rows
    except (OSError, ValueError, KeyError):
        return {}


def _service_rows(
    conn: DatabaseConnection,
    profile: ServiceProfile,
    start: date,
    end: date,
) -> list[dict[str, Any]]:
    filter_values = (
        [queue.upper() for queue in profile.flash_queues]
        if profile.flash_queues else list(profile.service_scopes)
    )
    source_filter = (
        f"upper(queue) IN ({','.join('?' for _ in filter_values)})"
        if profile.flash_queues
        else f"service_scope IN ({','.join('?' for _ in filter_values)})"
    )
    system_marks = ",".join("?" for _ in profile.source_systems)
    cursor = conn.execute(
        f"""SELECT business_date, interval_start, hour_start, source_system, queue,
                   service_scope, comparison_scope, designation, mapping_status,
                   offered, answered, abandoned, short_abandoned,
                   abandoned_within_target,
                   answered_within_target, handled_seconds, source_file
            FROM mart.service_interval
            WHERE business_date BETWEEN ? AND ?
              AND {source_filter}
              AND source_system IN ({system_marks})
            ORDER BY business_date, interval_start, source_system, queue""",
        [start, end, *filter_values, *profile.source_systems],
    )
    headers = [item[0] for item in cursor.description]
    rows = [dict(zip(headers, row)) for row in cursor.fetchall()]
    if profile.flash_total_groups and not profile.flash_queues:
        allowed = set(profile.flash_total_groups)
        rows = [
            row for row in rows
            if profile.group_for(row.get("queue")) in allowed
        ]
    return rows


def _profile_metric(catalog: MetricCatalog, profile: ServiceProfile, metric_id: str, on_date: date):
    methods = {}
    for service_scope in profile.service_scopes:
        for source_system in profile.source_systems:
            method = catalog.method_for(metric_id, on_date, {
                "lob": service_scope,
                "source_system": source_system,
            })
            if method is None:
                raise ValueError(
                    f"Service profile {profile.profile_id!r} selects metric {metric_id!r}, "
                    f"but no method applies to {service_scope}/{source_system} on {on_date}"
                )
            methods[(method.method_id, method.effective_from, method.priority)] = method
    if len(methods) != 1:
        names = ", ".join(key[0] for key in methods)
        raise ValueError(
            f"Service profile {profile.profile_id!r} crosses incompatible {metric_id} methods: {names}. "
            "Split it into separate service profiles."
        )
    return next(iter(methods.values()))


def _service_aggregate(
    rows: Iterable[dict[str, Any]],
    profile: ServiceProfile,
    catalog: MetricCatalog,
    on_date: date,
) -> dict[str, float | str | None]:
    values = list(rows)
    offered = sum(float(row.get("offered") or 0) for row in values)
    answered = sum(float(row.get("answered") or 0) for row in values)
    abandoned = sum(float(row.get("abandoned") or 0) for row in values)
    short = sum(float(row.get("short_abandoned") or 0) for row in values)
    abandoned_in_target = sum(
        float(row.get("abandoned_within_target") or 0) for row in values
    )
    within = sum(float(row.get("answered_within_target") or 0) for row in values)
    handled = sum(float(row.get("handled_seconds") or 0) for row in values)
    components = {
        "offered": offered,
        "answered": answered,
        "abandoned": abandoned,
        "short_abandoned": short,
        "abandoned_within_target": abandoned_in_target,
        "answered_within_target": within,
        "handled_seconds": handled,
    }
    service_level = evaluate_metric(
        _profile_metric(catalog, profile, profile.service_level_metric, on_date), components,
    )
    availability = evaluate_metric(
        _profile_metric(catalog, profile, profile.availability_metric, on_date), components,
    )
    aht = evaluate_metric(
        _profile_metric(catalog, profile, profile.aht_metric, on_date), components,
    )
    return {
        "raw_offered": offered,
        "offered": offered,
        "business_offered": max(0.0, offered - abandoned_in_target),
        "answered": answered,
        "short_abandoned": short,
        "abandoned_within_target": abandoned_in_target,
        "within_target": within,
        "service_level": service_level.value,
        "service_target": service_level.method.target,
        "service_state": service_level.state,
        "service_method": service_level.method.method_id,
        "availability": availability.value,
        "aht_seconds": aht.value,
    }


def build_realisations_workbook(
    conn: DatabaseConnection,
    config: Config,
    start: date,
    end: date,
    output: Path | None = None,
    profile_id: str | None = None,
) -> Path:
    """Build one management view across every configured service LOB.

    Supplying ``profile_id`` keeps the focused single-LOB route available for
    command-line use.  The normal menu intentionally includes all active
    profiles so Operations does not need to build four separate workbooks.
    """

    profiles = load_service_profiles(config.home, config.service_profiles)
    selected_profiles = (
        (profiles.select(profile_id, end),)
        if profile_id
        else tuple(profile for profile in profiles.profiles if profile.active_on(end))
    )
    metrics_catalog = load_metric_catalog(config.home, config.metric_catalog)
    queue_mapping = load_queue_mapping(config.queue_mapping)
    rulebook = load_rulebook(config.home, config.business_rules)
    title = (
        f"REALISATIONS  /  {selected_profiles[0].label.upper()}"
        if len(selected_profiles) == 1
        else "REALISATIONS  /  ALL MAPPED LOBS"
    )
    book, partial, target = _atomic_book(
        config, "realisations", title, start, end, output,
    )
    daily_rows: list[tuple[Any, ...]] = []
    all_source_rows: list[dict[str, Any]] = []
    profile_by_label = {profile.label: profile for profile in selected_profiles}
    (final_absence_by_day, absence_agent_headers, absence_agent_rows,
     absence_lob_headers, absence_lob_rows) = _realisations_final_absence(
        conn, selected_profiles, start, end,
    )

    for profile in selected_profiles:
        source_rows = _service_rows(conn, profile, start, end)
        for row in source_rows:
            all_source_rows.append({**row, "reporting_lob": profile.label})
        forecast_scopes = queue_mapping.comparison_scopes_for(
            profile.service_scopes,
        )
        forecast_marks = ",".join("?" for _ in forecast_scopes)
        forecasts = conn.execute(
            f"""SELECT business_date, sum(volume_forecast), avg(fte_forecast),
                       avg(fte_required),
                       CASE WHEN sum(CASE WHEN sl_forecast IS NOT NULL
                                          THEN volume_forecast ELSE 0 END)>0
                            THEN sum(CASE WHEN sl_forecast IS NOT NULL
                                          THEN sl_forecast*volume_forecast ELSE 0 END)
                                 /sum(CASE WHEN sl_forecast IS NOT NULL
                                           THEN volume_forecast ELSE 0 END)
                            ELSE avg(sl_forecast) END,
                       CASE WHEN sum(CASE WHEN sl_required IS NOT NULL
                                          THEN volume_forecast ELSE 0 END)>0
                            THEN sum(CASE WHEN sl_required IS NOT NULL
                                          THEN sl_required*volume_forecast ELSE 0 END)
                                 /sum(CASE WHEN sl_required IS NOT NULL
                                           THEN volume_forecast ELSE 0 END)
                            ELSE avg(sl_required) END,
                       CASE WHEN sum(CASE WHEN aht_forecast_seconds IS NOT NULL
                                          THEN volume_forecast ELSE 0 END)>0
                            THEN sum(CASE WHEN aht_forecast_seconds IS NOT NULL
                                          THEN aht_forecast_seconds*volume_forecast ELSE 0 END)
                                 /sum(CASE WHEN aht_forecast_seconds IS NOT NULL
                                           THEN volume_forecast ELSE 0 END)
                            ELSE avg(aht_forecast_seconds) END
                FROM mart.forecast_hour
                WHERE business_date BETWEEN ? AND ?
                  AND comparison_scope IN ({forecast_marks})
                GROUP BY business_date ORDER BY business_date""",
            [start, end, *forecast_scopes],
        ).fetchall()
        forecast_by_day = {str(row[0])[:10]: row[1:] for row in forecasts}
        staffing_marks = ",".join("?" for _ in profile.staffing_lobs)
        staffing_rows = conn.execute(
            f"""SELECT business_date, avg(scheduled_fte), avg(observed_fte),
                       avg(productive_fte), max(staffing_gap_fte)
                FROM (
                    SELECT business_date, interval_start,
                           sum(scheduled_fte) AS scheduled_fte,
                           sum(observed_fte) AS observed_fte,
                           sum(productive_fte) AS productive_fte,
                           sum(staffing_gap_fte) AS staffing_gap_fte
                    FROM mart.staffing_interval
                    WHERE business_date BETWEEN ? AND ? AND lob IN ({staffing_marks})
                    GROUP BY business_date, interval_start
                ) x GROUP BY business_date""",
            [start, end, *profile.staffing_lobs],
        ).fetchall()
        staffing_by_day = {str(row[0])[:10]: row[1:] for row in staffing_rows}
        rows_by_day: dict[str, list[dict[str, Any]]] = defaultdict(list)
        for row in source_rows:
            rows_by_day[str(row["business_date"])[:10]].append(row)

        cursor = start
        while cursor <= end:
            key = cursor.isoformat()
            actual = _service_aggregate(
                rows_by_day.get(key, []), profile, metrics_catalog, cursor,
            )
            forecast = forecast_by_day.get(key, (None, None, None, None, None, None))
            staffing = staffing_by_day.get(key, (None, None, None, None))
            absence_values = final_absence_by_day.get((key, profile.label))
            forecast_volume = forecast[0]
            planned_hours, absence_hours, vacation_hours, shrinkage_hours, review_cases = (
                absence_values if absence_values is not None
                else (None, None, None, None, None)
            )
            has_actual = bool(rows_by_day.get(key))
            has_forecast = forecast_volume is not None
            data_state = (
                "Provisional" if cursor >= date.today()
                else "Review" if review_cases
                else "No final absence" if absence_values is None
                else "No actual" if not has_actual
                else "No forecast" if not has_forecast
                else "Final"
            )
            daily_rows.append((
                cursor, cursor.strftime("%Y-%m"), cursor.isocalendar().week,
                f"Q{((cursor.month - 1) // 3) + 1}", profile.label,
                ", ".join(profile.service_scopes), actual["offered"],
                forecast_volume,
                float(actual["offered"] or 0) - float(forecast_volume)
                if forecast_volume is not None else None,
                _ratio(actual["offered"], forecast_volume), actual["answered"],
                actual["within_target"], actual["short_abandoned"],
                actual["abandoned_within_target"],
                actual["service_level"], actual["availability"],
                actual["aht_seconds"],
                _ratio(
                    float(actual["aht_seconds"] or 0) * float(actual["answered"] or 0),
                    3600,
                ),
                staffing[0], staffing[1], staffing[2], staffing[3],
                planned_hours, absence_hours, _ratio(absence_hours, planned_hours),
                vacation_hours, shrinkage_hours, _ratio(shrinkage_hours, planned_hours),
                review_cases, data_state, profile.profile_id,
                ", ".join(profile.staffing_lobs),
            ))
            cursor += timedelta(days=1)
    daily_headers = [
        "business_date", "month", "iso_week", "quarter", "lob",
        "service_scopes", "actual_volume", "forecast_volume",
        "variance_calls", "forecast_attainment", "answered",
        "answered_within_target", "short_abandoned",
        "abandoned_within_target", "service_level",
        "service_availability", "aht_seconds", "processing_hours",
        "scheduled_fte", "observed_fte", "productive_fte", "peak_gap_fte",
        "planned_hours", "absence_hours", "absence_rate", "vacation_hours",
        "shrinkage_hours", "shrinkage_rate", "review_cases", "data_state",
        "profile_id", "staffing_lobs",
    ]

    total_actual = sum(float(row[6] or 0) for row in daily_rows)
    total_forecast = sum(float(row[7] or 0) for row in daily_rows if row[7] is not None)
    forecast_present = any(row[7] is not None for row in daily_rows)
    total_answered = sum(float(row[10] or 0) for row in daily_rows)
    total_within = sum(float(row[11] or 0) for row in daily_rows)
    total_short = sum(float(row[12] or 0) for row in daily_rows)
    total_handled_seconds = sum(
        float(row[16] or 0) * float(row[10] or 0) for row in daily_rows
    )
    total_planned = sum(float(row[22] or 0) for row in daily_rows)
    total_absence = sum(float(row[23] or 0) for row in daily_rows)
    total_shrinkage = sum(float(row[26] or 0) for row in daily_rows)
    availability_value = _ratio(total_answered, total_actual)
    aht_value = _ratio(total_handled_seconds, total_answered)
    profile_summary: list[tuple[Any, ...]] = []
    for profile in selected_profiles:
        rows = [row for row in daily_rows if row[4] == profile.label]
        offered = sum(float(row[6] or 0) for row in rows)
        forecast = sum(float(row[7] or 0) for row in rows if row[7] is not None)
        answered = sum(float(row[10] or 0) for row in rows)
        within = sum(float(row[11] or 0) for row in rows)
        short = sum(float(row[12] or 0) for row in rows)
        abandoned_in_target = sum(float(row[13] or 0) for row in rows)
        handled = sum(float(row[16] or 0) * float(row[10] or 0) for row in rows)
        planned_hours = sum(float(row[22] or 0) for row in rows)
        absence_hours = sum(float(row[23] or 0) for row in rows)
        shrinkage_hours = sum(float(row[26] or 0) for row in rows)
        profile_components = {
            "offered": offered, "answered": answered,
            "short_abandoned": short,
            "abandoned_within_target": abandoned_in_target,
            "answered_within_target": within,
            "handled_seconds": handled,
        }
        sl_result = evaluate_metric(
            _profile_metric(metrics_catalog, profile, profile.service_level_metric, end),
            profile_components,
        )
        availability_result = evaluate_metric(
            _profile_metric(metrics_catalog, profile, profile.availability_metric, end),
            profile_components,
        )
        aht_result = evaluate_metric(
            _profile_metric(metrics_catalog, profile, profile.aht_metric, end),
            profile_components,
        )
        state = (
            "NO ACTUAL" if not any(float(row[6] or 0) for row in rows)
            else "NO FORECAST" if not any(row[7] is not None for row in rows)
            else "REVIEW" if any(row[-4] for row in rows)
            else "READY"
        )
        profile_summary.append((
            profile.label, offered,
            forecast if any(row[7] is not None for row in rows) else None,
            _ratio(offered, forecast), sl_result.value, sl_result.method.target,
            availability_result.value, aht_result.value,
            _ratio(absence_hours, planned_hours),
            _ratio(shrinkage_hours, planned_hours), state,
        ))
    status, status_text = _source_state(conn, ("calls", "forecast", "activities"), end)
    review_days = sum(1 for row in daily_rows if row[-4])
    if review_days:
        status, status_text = "INCOMPLETE", f"{review_days:,} day(s) include absence review cases"
    dashboard = book.report.workbook.add_worksheet("DASHBOARD")
    render_v2_dashboard(
        book.report.workbook, dashboard,
        title="REALISATIONS",
        filters=(
            ("Period", f"{start:%d %b} – {end:%d %b %Y}"),
            ("LOB", "All mapped" if len(selected_profiles) > 1 else selected_profiles[0].label),
            ("Grain", "Daily"),
            ("Data state", "Reviewed" if status != "INCOMPLETE" else "Check data"),
        ),
        kpis=(
            ("Actual volume", total_actual, "integer"),
            ("Forecast attainment", _ratio(total_actual, total_forecast) if forecast_present else None, "percent"),
            ("Routed rate", availability_value, "percent"),
            ("Weighted AHT", aht_value, "decimal"),
        ),
        left_chart=V2ChartSpec(
            "ACTUAL VS FORECAST BY LOB", "column",
            tuple(row[0] for row in profile_summary),
            (
                ("Actual", tuple(row[1] for row in profile_summary), COLORS["teal"]),
                ("Forecast", tuple(row[2] for row in profile_summary), COLORS["muted"]),
            ),
        ),
        right_chart=V2ChartSpec(
            "SERVICE LEVEL VS TARGET", "bar",
            tuple(row[0] for row in profile_summary),
            (
                ("Service level", tuple(row[4] for row in profile_summary), COLORS["teal"]),
                ("Target", tuple(row[5] for row in profile_summary), COLORS["muted"]),
            ),
            "percent", 0, 1,
        ),
        action_title="LOB realisation summary",
        action_headers=("LOB", "ACTUAL", "FORECAST", "ATTAINMENT", "SERVICE LEVEL", "ABSENCE %", "STATE"),
        action_rows=tuple((row[0], row[1], row[2], row[3], row[4], row[8], row[10]) for row in profile_summary),
        action_kinds=("text", "integer", "integer", "percent", "percent", "percent", "alert"),
        status="CHECK DATA" if status == "INCOMPLETE" else "OPERATIONAL",
        status_kind="INCOMPLETE" if status == "INCOMPLETE" else "PROVISIONAL",
        status_note=status_text,
    )
    book.table(
        "LOB_RESULTS", "Daily LOB results",
        "One normalized row per mapped management LOB and day. Use this sheet for pivots, charts and management checks.",
        daily_headers, daily_rows,
    )

    trend_rows: list[tuple[Any, ...]] = []
    for grain, index in (("Month", 1), ("ISO Week", 2), ("Quarter", 3)):
        grouped: dict[tuple[str, Any], list[tuple[Any, ...]]] = defaultdict(list)
        for row in daily_rows:
            grouped[(str(row[4]), row[index])].append(row)
        for (lob_label, label), group in sorted(grouped.items(), key=lambda item: (item[0][0], str(item[0][1]))):
            offered = sum(float(row[6] or 0) for row in group)
            forecast = sum(float(row[7] or 0) for row in group if row[7] is not None)
            answered = sum(float(row[10] or 0) for row in group)
            within = sum(float(row[11] or 0) for row in group)
            short = sum(float(row[12] or 0) for row in group)
            abandoned_in_target = sum(float(row[13] or 0) for row in group)
            handled = sum(float(row[16] or 0) * float(row[10] or 0) for row in group)
            planned_hours = sum(float(row[22] or 0) for row in group)
            absence_hours = sum(float(row[23] or 0) for row in group)
            shrinkage_hours = sum(float(row[26] or 0) for row in group)
            trend_components = {
                "offered": offered, "answered": answered,
                "short_abandoned": short,
                "abandoned_within_target": abandoned_in_target,
                "answered_within_target": within,
                "handled_seconds": handled,
            }
            trend_profile = profile_by_label[lob_label]
            trend_sl = evaluate_metric(
                _profile_metric(metrics_catalog, trend_profile, trend_profile.service_level_metric, end),
                trend_components,
            ).value
            trend_rows.append((
                lob_label, grain, label, min(row[0] for row in group), max(row[0] for row in group),
                offered, forecast if any(row[7] is not None for row in group) else None,
                _ratio(offered, forecast) if forecast else None,
                trend_sl, _ratio(answered, offered),
                _ratio(handled, answered), planned_hours, absence_hours,
                _ratio(absence_hours, planned_hours), shrinkage_hours,
                _ratio(shrinkage_hours, planned_hours),
            ))
    book.table(
        "TREND", "Period trend",
        "Month, ISO week and quarter summaries calculated from additive daily counters.",
        [
            "lob", "grain", "period", "start", "end", "actual_volume",
            "forecast_volume", "forecast_attainment", "service_level",
            "service_availability", "aht_seconds", "planned_hours",
            "absence_hours", "absence_rate", "shrinkage_hours", "shrinkage_rate",
        ],
        trend_rows,
    )
    detail_headers = [
        "reporting_lob", "business_date", "interval_start", "source_system", "queue",
        "service_scope", "designation", "mapping_status", "offered", "answered",
        "abandoned", "short_abandoned", "abandoned_within_target",
        "answered_within_target",
        "handled_seconds", "source_file",
    ]
    detail_rows = [tuple(row.get(header) for header in detail_headers) for row in all_source_rows]
    book.table(
        "DATA", "Mapped service data",
        "Filterable queue and interval evidence behind the LOB results.",
        detail_headers, detail_rows,
    )
    hour_headers, agent_hours, lob_hours = _realisations_hours(
        conn, selected_profiles, start, end, rulebook,
    )
    book.table(
        "HOURS_BY_LOB", "Where scheduled hours went by LOB",
        "Exclusive Agent Status categories inside published shifts; LILO-only and missing evidence stay visible, not guessed as productive.",
        hour_headers, lob_hours,
    )
    book.table(
        "HOURS_BY_AGENT", "Where scheduled hours went by agent",
        "One agent-day per workforce LOB. Every shift minute goes to exactly one category; reconciliation delta should be zero.",
        hour_headers, agent_hours,
    )
    aux_agent_headers, aux_agent_rows, aux_tl_headers, aux_tl_rows = (
        _realisations_aux_breakdown(conn, selected_profiles, start, end, rulebook)
    )
    book.table(
        "AUX_BY_TL", "Agent Status AUX by Team Lead",
        "Exact status and governed AUX class inside published shifts; group by day, LOB and Team Lead.",
        aux_tl_headers, aux_tl_rows,
    )
    book.table(
        "AUX_BY_AGENT", "Agent Status AUX by agent",
        "Exact observed status hours inside published shifts; BO is shown separately from other AUX.",
        aux_agent_headers, aux_agent_rows,
    )
    book.table(
        "ABS_BY_LOB", "Final absenteeism by LOB",
        "Verint Activities only. UNPAID_LEAVE is excluded from absence; No Show remains absence.",
        absence_lob_headers, absence_lob_rows,
    )
    book.table(
        "ABS_BY_AGENT", "Final absenteeism by agent",
        "Final Verint Activities evidence, unioned and capped per agent-day; unpaid leave is separate.",
        absence_agent_headers, absence_agent_rows,
    )
    book.definitions([
        ("Actual / forecast", "Actual offered contacts / forecast contacts", "Demand realisation", "Use summed volumes"),
        ("Service level", "Configured numerator / denominator for each service profile", "Service performance", "Calculated from summed counters; never average LOB percentages"),
        ("Routed Rate", "Answered / total entered", "Ability of the service to route demand", "Not agent availability"),
        ("Weighted AHT", "Handled seconds / answered contacts", "Workload", "Never average daily AHT values"),
        ("Absence rate", "Reviewed absence hours / planned hours", "Capacity impact", "Open attendance gaps remain review cases"),
        ("Hour allocation", "Exclusive Agent Status inside published shifts; LILO-only and missing evidence remain separate", "Where scheduled hours went", "Allocated hours plus reconciliation delta equals scheduled hours"),
        ("Realisations absence", "Unioned final Verint Activities absence except UNPAID_LEAVE / planned net hours", "Final absenteeism", "No Show remains absence even though its source rule also flags unpaid; review ledger exceptions"),
        ("AUX detail", "Exact Agent Status labels classified by the governed status rulebook", "Team Lead and agent coaching context", "BO is listed as its own work bucket; these are hours, not person counts"),
    ])
    book.audit(_audit_rows(conn, config, "realisations", start, end, [
        ("Service profiles", ", ".join(profile.profile_id for profile in selected_profiles), profiles.version),
        ("Included service scopes", "; ".join(
            f"{profile.label}: {', '.join(profile.service_scopes)}"
            for profile in selected_profiles
        ), "Queue Mapping"),
        ("Included staffing LOBs", "; ".join(
            f"{profile.label}: {', '.join(profile.staffing_lobs)}"
            for profile in selected_profiles
        ), "Service Profiles"),
    ]))
    return _finish(book, partial, target)


def build_staffing_coverage_workbook(
    conn: DatabaseConnection,
    config: Config,
    start: date,
    end: date,
    output: Path | None = None,
) -> Path:
    """Build actual staffing control and future required-versus-scheduled plan."""

    book, partial, target = _atomic_book(
        config, "staffing", "STAFFING & CAPACITY PLAN", start, end, output,
    )
    capacity_mapping = load_capacity_mapping(config.capacity_mapping)
    # Verint forecast ``Queue Name`` is a Staff Type identity. It must be
    # governed by capacity_mapping.csv and must never be joined to Storm queue
    # names or service-profile LOBs.
    forecast_by_capacity_interval: dict[
        tuple[str, str, str, str, datetime], list[float | None]
    ] = {}
    forecast_rows = conn.execute(
        """SELECT business_date, interval_start, interval_minutes, queue_name,
                  volume_forecast, fte_forecast, fte_required, source_file
           FROM mart.forecast_interval
           WHERE business_date BETWEEN ? AND ?
           ORDER BY business_date, interval_start, source_file, queue_name""",
        [start, end],
    ).fetchall()
    for (
        business_date, interval_start, interval_minutes, source_staff_type,
        volume, fte_forecast, fte_required, source_file,
    ) in forecast_rows:
        mapped = capacity_mapping.map_forecast(
            str(source_file or ""), str(source_staff_type or ""),
        )
        if mapped.status != "MAPPED":
            continue
        stamp = interval_start
        if isinstance(stamp, str):
            stamp = datetime.fromisoformat(stamp)
        source_minutes = max(15, int(interval_minutes or 60))
        slots = max(1, (source_minutes + 14) // 15)
        volume_per_slot = float(volume) / slots if volume is not None else None
        for offset in range(slots):
            key = (
                mapped.management_lob.casefold(), mapped.planning_group.casefold(),
                mapped.staff_type.casefold(), str(business_date)[:10],
                stamp + timedelta(minutes=15 * offset),
            )
            bucket = forecast_by_capacity_interval.setdefault(
                key, [None, None, None],
            )
            for index, value in enumerate(
                (volume_per_slot, fte_forecast, fte_required),
            ):
                if value is not None:
                    bucket[index] = float(bucket[index] or 0) + float(value)

    raw_headers, raw_rows = _query(
        conn,
        """SELECT business_date, interval_start, max(interval_end) AS interval_end,
                  lob, group_concat(DISTINCT language) AS language,
                  planning_group, staff_type, capacity_mapping_status,
                  sum(scheduled_agents) AS scheduled_agents,
                  sum(observed_agents) AS observed_agents,
                  sum(productive_agents) AS productive_agents,
                  sum(gross_scheduled_fte) AS gross_scheduled_fte,
                  sum(planned_time_off_fte) AS planned_time_off_fte,
                  sum(scheduled_fte) AS scheduled_fte,
                  sum(elapsed_scheduled_fte) AS elapsed_scheduled_fte,
                  sum(observed_fte) AS observed_fte,
                  sum(productive_fte) AS productive_fte,
                  sum(staffing_variance_fte) AS staffing_variance_fte,
                  sum(staffing_gap_fte) AS staffing_gap_fte,
                  max(staffing_state) AS staffing_state,
                  group_concat(DISTINCT evidence_basis) AS evidence_basis,
                  max(evaluation_as_of) AS evaluation_as_of
           FROM mart.staffing_interval
           WHERE business_date BETWEEN ? AND ?
           GROUP BY business_date, interval_start, lob, planning_group,
                    staff_type, capacity_mapping_status
           ORDER BY business_date, interval_start, planning_group, staff_type,
                    lob, language""",
        [start, end],
    )
    source_indexes = {header: index for index, header in enumerate(raw_headers)}
    evaluation_values: list[datetime] = []
    for raw in raw_rows:
        evaluation = raw[source_indexes["evaluation_as_of"]]
        if isinstance(evaluation, str):
            evaluation = datetime.fromisoformat(evaluation)
        if evaluation is not None:
            evaluation_values.append(evaluation)
    evaluation_default = max(evaluation_values, default=datetime.now())
    plan_headers = [
        "business_date", "iso_week", "interval_start", "interval_end",
        "management_lob", "planning_group", "staff_type", "mode",
        "forecast_volume_interval", "fte_forecast", "fte_required",
        "gross_scheduled_fte", "planned_time_off_fte", "net_scheduled_fte",
        "capacity_variance_fte", "capacity_gap_fte", "observed_fte",
        "productive_fte", "actual_gap_fte", "decision_state",
        "evidence_basis", "evaluation_as_of", "capacity_key", "source_lob",
        "language",
    ]
    plan_rows: list[tuple[Any, ...]] = []
    covered_intervals: set[tuple[str, str, str, str, datetime]] = set()
    for raw in raw_rows:
        values = {header: raw[index] for header, index in source_indexes.items()}
        business_date = values["business_date"]
        if isinstance(business_date, str):
            business_date = date.fromisoformat(business_date[:10])
        interval_start = values["interval_start"]
        if isinstance(interval_start, str):
            interval_start = datetime.fromisoformat(interval_start)
        interval_end = values["interval_end"]
        if isinstance(interval_end, str):
            interval_end = datetime.fromisoformat(interval_end)
        evaluation_as_of = values["evaluation_as_of"]
        if isinstance(evaluation_as_of, str):
            evaluation_as_of = datetime.fromisoformat(evaluation_as_of)
        source_lob = str(values["lob"] or "")
        workforce = capacity_mapping.map_schedule(source_lob, None)
        management_lob = workforce.management_lob
        planning_group = str(values["planning_group"] or "")
        staff_type = str(values["staff_type"] or "")
        capacity_key = f"{management_lob}|{planning_group}|{staff_type}|{business_date}|{interval_start:%H:%M}"
        forecast = forecast_by_capacity_interval.get((
            management_lob.casefold(), planning_group.casefold(),
            staff_type.casefold(), business_date.isoformat(), interval_start,
        ))
        volume, fte_forecast, fte_required = forecast or (None, None, None)
        net_scheduled = float(values["scheduled_fte"] or 0)
        capacity_variance = net_scheduled - float(fte_required) if fte_required is not None else None
        capacity_gap = max(0.0, -capacity_variance) if capacity_variance is not None else None
        future = interval_start > evaluation_as_of
        mode = "FUTURE PLAN" if future else "ACTUAL CONTROL"
        actual_gap = (
            max(
                0.0,
                float(values["elapsed_scheduled_fte"] or 0)
                - float(values["observed_fte"] or 0),
            )
            if str(values["staffing_state"] or "").upper()
            not in {"FUTURE", "DATA_MISSING", "DATA_PARTIAL"}
            else None
        )
        if str(values["capacity_mapping_status"] or "").upper() != "MAPPED":
            state = "UNMAPPED CAPACITY"
        elif fte_required is None:
            state = "NO FORECAST"
        elif future:
            state = "FUTURE GAP" if capacity_gap and capacity_gap > 0.001 else "FUTURE OK"
        else:
            state = str(values["staffing_state"] or "DATA MISSING").replace("_", " ")
        plan_rows.append((
            business_date, f"{business_date.isocalendar().year}-W{business_date.isocalendar().week:02d}",
            interval_start, interval_end, management_lob, planning_group,
            staff_type, mode, volume, fte_forecast,
            fte_required, values["gross_scheduled_fte"],
            values["planned_time_off_fte"], net_scheduled, capacity_variance,
            capacity_gap, values["observed_fte"], values["productive_fte"],
            actual_gap, state, values["evidence_basis"], evaluation_as_of,
            capacity_key, source_lob, values["language"],
        ))
        covered_intervals.add((
            management_lob.casefold(), planning_group.casefold(),
            staff_type.casefold(), business_date.isoformat(), interval_start,
        ))

    # A demand interval with zero scheduled agents does not exist in the
    # staffing mart. Add it here so an empty roster can never hide a shortage.
    for (
        management_lob_key, planning_group_key, staff_type_key,
        business_date_text, interval_start,
    ), forecast in forecast_by_capacity_interval.items():
        volume, fte_forecast, fte_required = forecast
        mapped = next(
            result for _prefix, _staff, result in capacity_mapping.forecast_rows
            if result.management_lob.casefold() == management_lob_key
            and result.planning_group.casefold() == planning_group_key
            and result.staff_type.casefold() == staff_type_key
        )
        business_date = date.fromisoformat(business_date_text)
        key = (
            management_lob_key, planning_group_key, staff_type_key,
            business_date_text, interval_start,
        )
        if key in covered_intervals:
            continue
        interval_end = interval_start + timedelta(minutes=15)
        required = float(fte_required or 0)
        future = interval_start > evaluation_default
        mode = "FUTURE PLAN" if future else "ACTUAL CONTROL"
        state = "FUTURE GAP" if future and required > 0 else "NO SCHEDULE"
        plan_rows.append((
            business_date,
            f"{business_date.isocalendar().year}-W{business_date.isocalendar().week:02d}",
            interval_start, interval_end, mapped.management_lob,
            mapped.planning_group, mapped.staff_type,
            mode, volume, fte_forecast, fte_required, 0.0, 0.0, 0.0,
            -required, required, None, None, required if not future else None,
            state, "Forecast demand with no scheduled roster interval",
            evaluation_default,
            f"{mapped.management_lob}|{mapped.planning_group}|{mapped.staff_type}|{business_date}|{interval_start:%H:%M}",
            mapped.workforce_lob, "Unspecified",
        ))

    plan_rows.sort(key=lambda row: (row[0], row[2], str(row[4]), str(row[5]), str(row[6])))

    decision_state = plan_headers.index("decision_state")
    capacity_gap_column = plan_headers.index("capacity_gap_fte")
    actual_gap_column = plan_headers.index("actual_gap_fte")
    action_states = {"FUTURE GAP", "NO FORECAST", "NO SCHEDULE", "UNMAPPED CAPACITY", "GAP", "PARTIAL GAP", "DATA MISSING"}
    actions = [row for row in plan_rows if str(row[decision_state]).upper() in action_states]
    future_rows = [row for row in plan_rows if row[plan_headers.index("mode")] == "FUTURE PLAN"]
    required_hours = sum(float(row[plan_headers.index("fte_required")] or 0) * 0.25 for row in future_rows)
    net_hours = sum(float(row[plan_headers.index("net_scheduled_fte")] or 0) * 0.25 for row in future_rows)
    pto_hours = sum(float(row[plan_headers.index("planned_time_off_fte")] or 0) * 0.25 for row in future_rows)
    gap_hours = sum(float(row[capacity_gap_column] or 0) * 0.25 for row in future_rows)
    future_gap_intervals = sum(1 for row in future_rows if row[decision_state] == "FUTURE GAP")
    no_forecast_intervals = sum(1 for row in future_rows if row[decision_state] == "NO FORECAST")
    actual_action_rows = [row for row in plan_rows if row[plan_headers.index("mode")] == "ACTUAL CONTROL"]
    actual_gap_intervals = sum(
        1 for row in actual_action_rows
        if str(row[decision_state]).upper() in {"GAP", "PARTIAL GAP"}
    )
    peak_gap = max(
        [float(row[capacity_gap_column] or 0) for row in future_rows]
        + [float(row[actual_gap_column] or 0) for row in actual_action_rows]
        + [0.0]
    )
    status, status_text = _source_state(conn, ("fte", "start_end", "forecast"), end)
    if no_forecast_intervals:
        status, status_text = "INCOMPLETE", f"{no_forecast_intervals:,} future interval(s) have no mapped forecast"
    chart_rows = future_rows or plan_rows
    by_lob: dict[str, list[float]] = defaultdict(lambda: [0.0, 0.0])
    by_day: dict[str, float] = defaultdict(float)
    for row in chart_rows:
        lob_label = str(row[4] or "Unmapped")
        by_lob[lob_label][0] += float(row[10] or 0) * 0.25
        by_lob[lob_label][1] += float(row[13] or 0) * 0.25
        gap_value = float((row[15] if row[7] == "FUTURE PLAN" else row[18]) or 0)
        day_label = row[0].strftime("%a %d") if isinstance(row[0], date) else str(row[0])
        by_day[day_label] = max(by_day[day_label], gap_value)
    priority_actions = sorted(
        actions,
        key=lambda row: max(float(row[15] or 0), float(row[18] or 0)),
        reverse=True,
    )[:8]
    dashboard = book.report.workbook.add_worksheet("DASHBOARD")
    render_v2_dashboard(
        book.report.workbook, dashboard,
        title="STAFFING & COVERAGE",
        filters=(
            ("Period", f"{start:%d %b} – {end:%d %b %Y}"),
            ("Mode", "Future plan" if future_rows else "Actual control"),
            ("LOB", "All"),
            ("Staff Type", "All"),
        ),
        kpis=(
            ("Peak gap FTE", peak_gap, "decimal"),
            ("Future gap hours", gap_hours, "decimal"),
            ("Forecast coverage", _ratio(required_hours - gap_hours, required_hours), "percent"),
            ("PTO / Away impact", pto_hours, "decimal"),
        ),
        left_chart=V2ChartSpec(
            "REQUIRED VS NET SCHEDULED HOURS", "column", tuple(by_lob),
            (
                ("Required", tuple(values[0] for values in by_lob.values()), COLORS["teal"]),
                ("Net scheduled", tuple(values[1] for values in by_lob.values()), COLORS["muted"]),
            ),
        ),
        right_chart=V2ChartSpec(
            "PEAK GAP BY DAY", "line", tuple(by_day),
            (("Gap FTE", tuple(by_day.values()), COLORS["red"]),),
        ),
        action_title="Prioritized capacity actions",
        action_headers=("INTERVAL", "PLANNING GROUP / STAFF TYPE", "MODE", "REQUIRED FTE", "NET / OBSERVED", "GAP FTE", "STATE"),
        action_rows=tuple((
            row[2].strftime("%d %b %H:%M") if isinstance(row[2], datetime) else str(row[2]),
            f"{row[4]} / {row[6]}",
            "Future" if row[7] == "FUTURE PLAN" else "Actual",
            row[10], row[13] if row[7] == "FUTURE PLAN" else row[16],
            row[15] if row[7] == "FUTURE PLAN" else row[18], row[19],
        ) for row in priority_actions),
        action_kinds=("text", "text", "text", "decimal", "decimal", "decimal", "alert"),
        status="CHECK DATA" if status == "INCOMPLETE" else "IN DEVELOPMENT",
        status_kind="INCOMPLETE" if status == "INCOMPLETE" else "PROVISIONAL",
        status_note=status_text,
    )

    weekly: dict[tuple[str, str, str, str], list[tuple[Any, ...]]] = defaultdict(list)
    for row in plan_rows:
        weekly[(str(row[1]), str(row[4]), str(row[5]), str(row[6]))].append(row)
    weekly_rows = []
    for (iso_week, management_lob, planning_group, staff_type), rows in sorted(weekly.items()):
        required = sum(float(row[10] or 0) * 0.25 for row in rows)
        gross = sum(float(row[11] or 0) * 0.25 for row in rows)
        time_off = sum(float(row[12] or 0) * 0.25 for row in rows)
        net = sum(float(row[13] or 0) * 0.25 for row in rows)
        gap = sum(float(row[15] or 0) * 0.25 for row in rows)
        weekly_rows.append((
            iso_week, min(row[0] for row in rows), max(row[0] for row in rows),
            management_lob, planning_group, staff_type, required, gross, time_off, net,
            net - required if required else None, gap,
            _ratio(net, required),
            sum(1 for row in rows if row[19] == "FUTURE GAP"),
            sum(1 for row in rows if row[19] == "NO FORECAST"),
        ))
    book.table(
        "WEEKLY_PLAN", "Weekly capacity plan",
        "Required, gross scheduled, PTO/Away and net scheduled FTE-hours by ISO week, LOB and language.",
        [
            "iso_week", "start_date", "end_date", "management_lob", "planning_group",
            "staff_type", "required_fte_hours", "gross_scheduled_fte_hours",
            "planned_time_off_fte_hours", "net_scheduled_fte_hours",
            "capacity_variance_fte_hours", "capacity_gap_fte_hours",
            "forecast_coverage", "gap_intervals", "no_forecast_intervals",
        ],
        weekly_rows,
    )
    intraday = book.table(
        "INTRADAY", "15-minute staffing control and plan",
        "All selected dates at governed Management LOB, Planning Group and Verint Staff Type grain. FUTURE PLAN uses forecast demand; ACTUAL CONTROL uses observed attendance evidence.",
        plan_headers, plan_rows,
    )
    if plan_rows:
        intraday.conditional_format(
            4, decision_state, 3 + len(plan_rows), decision_state,
            {"type": "text", "criteria": "containing", "value": "GAP", "format": book.report.error},
        )
    book.table(
        "CAPACITY GAPS", "Current staffing exceptions",
        "Read-only future shortages, missing forecasts and actual gaps. Capacity Key is the stable handoff identifier.",
        plan_headers, actions,
    )
    prior_actions = _load_persistent_rows(target, "ACTIONS", "capacity_key")
    ledger_headers = [
        "capacity_key", "business_date", "interval_start", "management_lob",
        "planning_group", "staff_type", "gap_fte", "decision_state",
        "action", "planned_fte", "owner", "status", "note", "last_updated",
    ]
    ledger_rows: list[tuple[Any, ...]] = []
    current_keys: set[str] = set()
    for row in actions:
        key = str(row[22])
        current_keys.add(key)
        prior = prior_actions.get(key, {})
        ledger_rows.append((
            key, row[0], row[2], row[4], row[5], row[6],
            max(float(row[15] or 0), float(row[18] or 0)), row[19],
            prior.get("action"), prior.get("planned_fte"), prior.get("owner"),
            prior.get("status") or "OPEN", prior.get("note"),
            prior.get("last_updated"),
        ))
    for key, prior in sorted(prior_actions.items()):
        if key in current_keys:
            continue
        ledger_rows.append(tuple(
            prior.get(header) if header != "decision_state" else "RESOLVED BY DATA"
            for header in ledger_headers
        ))
    book.table(
        "ACTIONS", "Persistent staffing action ledger",
        "One durable row per Capacity Key. Blue columns are retained when this report is rebuilt.",
        ledger_headers, ledger_rows,
        editable_headers={
            "action", "planned_fte", "owner", "status", "note", "last_updated",
        },
    )
    book.definitions([
        ("Required FTE", "Verint required FTE at native 15-minute grain; historical hourly files expand to four quarters", "Demand requirement", "Forecast only; missing stays blank"),
        ("Capacity grain", "Management LOB + Planning Group + Verint Staff Type", "Staff preparation identity", "Never Storm queue or activity label"),
        ("Capacity Key", "Management LOB|Planning Group|Staff Type|date|15-minute interval", "Stable action and reconciliation key", "One requirement bucket"),
        ("Gross scheduled FTE", "Scheduled agent-seconds / 900 before time off", "Roster capacity", "Published assignment mapped by capacity_mapping.csv"),
        ("Net scheduled FTE", "Gross scheduled FTE - approved PTO/effective Away FTE", "Usable planned capacity", "Planned Away affects future only"),
        ("Future capacity gap", "MAX(0, required FTE - net scheduled FTE)", "Hiring, OT or redeployment action", "15-minute interval"),
        ("Observed FTE", "Observed agent-seconds / 900", "Actual presence", "LILO + Agent Status evidence"),
        ("Productive FTE", "Productive-status seconds / 900", "Available handling capacity", "Not adherence"),
        ("Actual staffing gap", "MAX(0, elapsed net scheduled FTE - observed FTE)", "Same-day staffing deficit", "Blank for future/missing evidence"),
    ])
    book.audit(_audit_rows(conn, config, "staffing", start, end, [
        ("Capacity mapping", capacity_mapping.sha256, str(capacity_mapping.file)),
        ("Planning grain", "15 minutes", "Weekly summary uses FTE-hours"),
    ]))
    return _finish(book, partial, target)


def _exclusive_final_components(
    conn: DatabaseConnection,
    rulebook,
    start: date,
    end: date,
    flag: str,
) -> tuple[list[str], list[tuple[Any, ...]]]:
    """Allocate overlapping final Activities once for a component breakdown."""

    columns = {
        "absence": "e.counts_as_absence",
        "shrinkage": "e.counts_as_shrinkage",
    }
    if flag not in columns:
        raise ValueError(f"Unsupported final component flag {flag!r}")
    headers, raw_rows = _query(
        conn,
        f"""SELECT e.agent_day_key, e.business_date, e.agent_id, e.agent_name,
                   e.team_leader, e.lob, e.activity, e.category,
                   e.event_start, e.event_end, d.planned_net_minutes
            FROM mart.verint_final_absence_event e
            JOIN mart.verint_final_absence_agent_day d
              ON d.agent_day_key=e.agent_day_key
            WHERE e.business_date BETWEEN ? AND ? AND {columns[flag]}=true
            ORDER BY e.business_date, e.agent_id, e.event_start, e.event_end""",
        [start, end],
    )
    events = [dict(zip(headers, row)) for row in raw_rows]
    precedence = {rule.name: index for index, rule in enumerate(rulebook.activity_rules)}
    by_day: dict[str, list[dict[str, Any]]] = defaultdict(list)
    for event in events:
        by_day[str(event["agent_day_key"])].append(event)
    totals: dict[tuple[Any, ...], int] = defaultdict(int)
    days: dict[tuple[Any, ...], set[Any]] = defaultdict(set)
    for day_events in by_day.values():
        remaining = int(day_events[0].get("planned_net_minutes") or 0)
        boundaries = sorted({
            stamp for event in day_events
            for stamp in (event["event_start"], event["event_end"])
            if stamp is not None
        })
        for left, right in zip(boundaries, boundaries[1:]):
            if right <= left:
                continue
            active = [
                event for event in day_events
                if event["event_start"] <= left and event["event_end"] >= right
            ]
            if not active:
                continue

            def rank(event: dict[str, Any]) -> tuple[int, str, str]:
                rule = rulebook.classify_activity(event.get("activity"))
                return (
                    precedence.get(rule.name if rule else "", len(precedence) + 1),
                    str(event.get("category") or "Other"),
                    str(event.get("activity") or ""),
                )

            chosen = min(active, key=rank)
            key = (
                chosen.get("category") or "OTHER",
                chosen.get("lob"), chosen.get("team_leader"),
                chosen.get("agent_id"), chosen.get("agent_name"),
            )
            allocated = min(remaining, int((right - left).total_seconds() // 60))
            if allocated <= 0:
                break
            totals[key] += allocated
            remaining -= allocated
            days[key].add(chosen.get("business_date"))
    output = [
        (*key, len(days[key]), minutes, minutes / 60.0)
        for key, minutes in totals.items()
        if minutes > 0
    ]
    output.sort(key=lambda row: (str(row[0]), str(row[1]), str(row[2]), str(row[4])))
    return (
        [
            "component", "lob", "team_leader", "agent_id", "agent_name",
            "agent_days", "minutes", "hours",
        ],
        output,
    )


def _add_absence_team_view(
    book: DecisionWorkbook,
    start: date,
    end: date,
    data_start: date,
    latest: date,
) -> None:
    """Add one selector-driven view for Operations and Payroll reviewers."""

    wb = book.report.workbook
    wb.set_calc_mode("auto")
    ws = wb.add_worksheet("TEAM_VIEW")
    ws.hide_gridlines(2)
    ws.set_tab_color(COLORS["gold"])
    ws.set_zoom(85)
    ws.freeze_panes(16, 0)
    ws.merge_range("A1:AA1", "ABSENTEEISM & SHRINKAGE  /  TEAM REVIEW", book.report.title)
    ws.merge_range(
        "A2:AA2",
        f"Latest final-ledger date {latest:%Y-%m-%d}  |  select a period and team below  |  prepared by Anass ASSRI",
        book.report.subtitle,
    )
    selector_label = wb.add_format({
        "font_name": "Aptos", "font_size": 8, "bold": True,
        "font_color": COLORS["muted"], "bg_color": COLORS["canvas"],
        "align": "left", "valign": "vcenter", "indent": 1,
    })
    selector = wb.add_format({
        "font_name": "Aptos Display", "font_size": 11, "bold": True,
        "font_color": COLORS["dark"], "bg_color": COLORS["white"],
        "border": 1, "border_color": COLORS["teal"], "align": "left",
        "valign": "vcenter", "indent": 1, "num_format": "yyyy-mm-dd",
    })
    for label, label_range, value_range, value in (
        ("PERIOD VIEW", "A4:C4", "A5:C6", "Current MTD"),
        ("CUSTOM FROM", "E4:F4", "E5:F6", start),
        ("CUSTOM TO", "H4:I4", "H5:I6", end),
        ("LOB", "K4:L4", "K5:L6", "All"),
        ("TEAM LEADER", "N4:O4", "N5:O6", "All"),
        ("AGENT", "Q4:R4", "Q5:R6", "All"),
    ):
        ws.merge_range(label_range, label, selector_label)
        ws.merge_range(value_range, value, selector)
    ws.data_validation("A5", {
        "validate": "list",
        "source": [
            "Latest day", "Current week", "Previous week", "Current MTD",
            "Previous-month same days", "Previous full month", "Custom period",
        ],
    })
    ws.data_validation("E5", {
        "validate": "date", "criteria": "between", "minimum": data_start,
        "maximum": latest,
    })
    ws.data_validation("H5", {
        "validate": "date", "criteria": "between", "minimum": data_start,
        "maximum": latest,
    })
    ws.data_validation("K5", {"validate": "list", "source": "=ABS_LOB_LIST"})
    ws.data_validation("N5", {"validate": "list", "source": "=ABS_TL_LIST"})
    ws.data_validation("Q5", {"validate": "list", "source": "=ABS_AGENT_LIST"})
    wb.define_name("ABS_Latest", "=MAX(tblAbsenceData[Date])")
    wb.define_name(
        "ABS_From",
        '=IF(TEAM_VIEW!$A$5="Latest day",ABS_Latest,'
        'IF(TEAM_VIEW!$A$5="Current week",ABS_Latest-WEEKDAY(ABS_Latest,2)+1,'
        'IF(TEAM_VIEW!$A$5="Previous week",ABS_Latest-WEEKDAY(ABS_Latest,2)-6,'
        'IF(TEAM_VIEW!$A$5="Current MTD",EOMONTH(ABS_Latest,-1)+1,'
        'IF(TEAM_VIEW!$A$5="Previous-month same days",EOMONTH(ABS_Latest,-2)+1,'
        'IF(TEAM_VIEW!$A$5="Previous full month",EOMONTH(ABS_Latest,-2)+1,TEAM_VIEW!$E$5))))))',
    )
    wb.define_name(
        "ABS_To",
        '=IF(TEAM_VIEW!$A$5="Latest day",ABS_Latest,'
        'IF(TEAM_VIEW!$A$5="Current week",ABS_Latest,'
        'IF(TEAM_VIEW!$A$5="Previous week",ABS_Latest-WEEKDAY(ABS_Latest,2),'
        'IF(TEAM_VIEW!$A$5="Current MTD",ABS_Latest,'
        'IF(TEAM_VIEW!$A$5="Previous-month same days",EDATE(ABS_Latest,-1),'
        'IF(TEAM_VIEW!$A$5="Previous full month",EOMONTH(ABS_Latest,-1),TEAM_VIEW!$H$5))))))',
    )
    scope = (
        '(tblAbsenceData[Date]>=ABS_From)*(tblAbsenceData[Date]<=ABS_To)*'
        'IF(TEAM_VIEW!$K$5="All",1,--(tblAbsenceData[LOB]=TEAM_VIEW!$K$5))*'
        'IF(TEAM_VIEW!$N$5="All",1,--(tblAbsenceData[Team Leader]=TEAM_VIEW!$N$5))*'
        'IF(TEAM_VIEW!$Q$5="All",1,--(tblAbsenceData[Agent Selector]=TEAM_VIEW!$Q$5))'
    )
    final = '((tblAbsenceData[Final Ledger Status]="CLEAR")+(tblAbsenceData[Final Ledger Status]="ABSENCE_RECORDED"))'
    planned = f"SUMPRODUCT({scope}*{final}*N(tblAbsenceData[Planned Net Hours]))"
    absence = f"SUMPRODUCT({scope}*{final}*N(tblAbsenceData[Absence Hours]))"
    shrinkage = f"SUMPRODUCT({scope}*{final}*N(tblAbsenceData[Shrinkage Hours]))"
    review = (
        f'SUMPRODUCT({scope}*--(tblAbsenceData[Final Ledger Status]<>"CLEAR")*'
        '--(tblAbsenceData[Final Ledger Status]<>"ABSENCE_RECORDED"))'
    )
    ws.merge_range("A8:AA8", "SELECTED TEAM POSITION", book.report.section)
    cards = (
        ("ABSENCE RATE", f'=IFERROR({absence}/{planned},"")', book.card_percent),
        ("SHRINKAGE RATE", f'=IFERROR({shrinkage}/{planned},"")', book.card_percent),
        ("ABSENCE HOURS", f"={absence}", book.card_decimal),
        ("REVIEW CASES", f"={review}", book.card_integer),
    )
    for index, (label, formula, fmt) in enumerate(cards):
        column = index * 4
        ws.merge_range(9, column, 9, column + 2, label, book.report.kpi_label)
        ws.merge_range(10, column, 11, column + 2, "", fmt)
        ws.write_formula(10, column, formula, fmt, "")
    ws.merge_range("A14:L14", "AGENT RESULTS", book.report.section)
    agent_headers = [
        "LOB", "Team Leader", "Agent Selector", "Agent ID", "Planned Hours",
        "Absence Hours", "Absence Rate", "Shrinkage Hours", "Shrinkage Rate",
        "Vacation Hours", "Unpaid Hours", "Review Cases",
    ]
    for column, header in enumerate(agent_headers):
        ws.write(15, column, header, book.report.header)
    agent_formula = (
        '=LET(d,tblAbsenceData,'
        'm,(d[Date]>=ABS_From)*(d[Date]<=ABS_To)*'
        'IF(TEAM_VIEW!$K$5="All",1,--(d[LOB]=TEAM_VIEW!$K$5))*'
        'IF(TEAM_VIEW!$N$5="All",1,--(d[Team Leader]=TEAM_VIEW!$N$5))*'
        'IF(TEAM_VIEW!$Q$5="All",1,--(d[Agent Selector]=TEAM_VIEW!$Q$5)),'
        'f,((d[Final Ledger Status]="CLEAR")+(d[Final Ledger Status]="ABSENCE_RECORDED")),'
        'a,SORT(UNIQUE(FILTER(d[Agent Selector],m,""))),'
        'ph,MAP(a,LAMBDA(x,SUMPRODUCT(m*f*(d[Agent Selector]=x)*N(d[Planned Net Hours])))),'
        'ah,MAP(a,LAMBDA(x,SUMPRODUCT(m*f*(d[Agent Selector]=x)*N(d[Absence Hours])))),'
        'sh,MAP(a,LAMBDA(x,SUMPRODUCT(m*f*(d[Agent Selector]=x)*N(d[Shrinkage Hours])))),'
        'vh,MAP(a,LAMBDA(x,SUMPRODUCT(m*f*(d[Agent Selector]=x)*N(d[Vacation Hours])))),'
        'uh,MAP(a,LAMBDA(x,SUMPRODUCT(m*f*(d[Agent Selector]=x)*N(d[Unpaid Hours])))),'
        'rv,MAP(a,LAMBDA(x,SUMPRODUCT(m*(d[Agent Selector]=x)*--(d[Final Ledger Status]<>"CLEAR")*--(d[Final Ledger Status]<>"ABSENCE_RECORDED")))),'
        'IFERROR(HSTACK('
        'XLOOKUP(a,d[Agent Selector],d[LOB],""),'
        'XLOOKUP(a,d[Agent Selector],d[Team Leader],""),a,'
        'XLOOKUP(a,d[Agent Selector],d[Agent ID],""),ph,ah,IFERROR(ah/ph,""),'
        'sh,IFERROR(sh/ph,""),vh,uh,rv),"No matching agent data"))'
    )
    ws.write_dynamic_array_formula("A17", agent_formula, book.report.body, "Open in desktop Excel")
    ws.merge_range("N14:AA14", "CASES TO REVIEW", book.report.section)
    queue_headers = [
        "Case ID", "Date", "Agent ID", "Agent", "Team Leader", "LOB",
        "Result Status", "Absence Hours", "Shrinkage Hours", "Unmapped Hours",
    ]
    for column, header in enumerate(queue_headers, 13):
        ws.write(15, column, header, book.report.header)
    queue_formula = (
        '=LET(q,tblActionQueue,'
        'm,(q[Date]>=ABS_From)*(q[Date]<=ABS_To)*'
        'IF(TEAM_VIEW!$K$5="All",1,--(q[LOB]=TEAM_VIEW!$K$5))*'
        'IF(TEAM_VIEW!$N$5="All",1,--(q[Team Leader]=TEAM_VIEW!$N$5))*'
        'IF(TEAM_VIEW!$Q$5="All",1,--(q[Agent Selector]=TEAM_VIEW!$Q$5)),'
        'IFERROR(FILTER(CHOOSECOLS(q,1,2,3,4,6,8,10,12,13,14),m),'
        '"No cases in this selection"))'
    )
    ws.write_dynamic_array_formula("N17", queue_formula, book.report.body, "Open in desktop Excel")
    ws.set_column("A:A", 18)
    ws.set_column("B:B", 22)
    ws.set_column("C:C", 30)
    ws.set_column("D:L", 15)
    ws.set_column("M:M", 3)
    ws.set_column("N:AA", 18)
    ws.set_landscape()
    ws.fit_to_pages(1, 0)


def _add_absence_component_view(book: DecisionWorkbook) -> None:
    """Show filtered absence and shrinkage components from exact intervals."""

    wb = book.report.workbook
    ws = wb.add_worksheet("COMPONENT_VIEW")
    ws.hide_gridlines(2)
    ws.set_tab_color(COLORS["gold"])
    ws.set_zoom(90)
    ws.merge_range("A1:N1", "ABSENCE & SHRINKAGE  /  COMPONENT VIEW", book.report.title)
    ws.merge_range(
        "A2:N2", "This view follows TEAM_VIEW period, LOB, Team Leader and Agent selectors.",
        book.report.subtitle,
    )
    ws.merge_range("A4:F4", "ABSENCE COMPONENTS", book.report.section)
    ws.merge_range("H4:N4", "SHRINKAGE COMPONENTS", book.report.section)
    headers = ("Component", "Hours", "Intervals")
    for column, header in enumerate(headers):
        ws.write(5, column, header, book.report.header)
        ws.write(5, column + 7, header, book.report.header)
    base_scope = (
        '(t[Date]>=ABS_From)*(t[Date]<=ABS_To)*'
        'IF(TEAM_VIEW!$K$5="All",1,--(t[LOB]=TEAM_VIEW!$K$5))*'
        'IF(TEAM_VIEW!$N$5="All",1,--(t[Team Leader]=TEAM_VIEW!$N$5))*'
        'IF(TEAM_VIEW!$Q$5="All",1,--(t[Agent Selector]=TEAM_VIEW!$Q$5))'
    )
    for cell, flag, empty_text in (
        ("A7", "Counts As Absence", "No absence components in this selection"),
        ("H7", "Counts As Shrinkage", "No shrinkage components in this selection"),
    ):
        formula = (
            '=LET(t,tblActivityDetail,'
            f'm,{base_scope}*--(t[{flag}]=TRUE),'
            'c,SORT(UNIQUE(FILTER(t[Category],m,""))),'
            'h,MAP(c,LAMBDA(x,SUMPRODUCT(m*(t[Category]=x)*N(t[Hours])))),'
            'n,MAP(c,LAMBDA(x,SUMPRODUCT(m*(t[Category]=x)))),'
            f'IFERROR(HSTACK(c,h,n),"{empty_text}"))'
        )
        ws.write_dynamic_array_formula(cell, formula, book.report.body, "Open in desktop Excel")
    ws.merge_range(
        "A22:N22", "Exact start/end evidence remains on ACTIVITY_DETAIL. Component hours are exclusive inside each KPI view; never add Absence Rate and Shrinkage Rate together.",
        book.report.note,
    )
    ws.write_url(
        "A24", "internal:'ACTIVITY_DETAIL'!A1", book.report.editable,
        string="OPEN EXACT VERINT ACTIVITY DETAIL",
    )
    ws.set_column("A:A", 28)
    ws.set_column("B:C", 15)
    ws.set_column("D:G", 4)
    ws.set_column("H:H", 28)
    ws.set_column("I:J", 15)


def _add_absence_lookups(book: DecisionWorkbook) -> None:
    """Create cascading LOB, Team Leader and Agent lists for Absenteeism."""

    wb = book.report.workbook
    ws = wb.add_worksheet("_LOOKUPS")
    ws.write("A1", "All")
    ws.write_dynamic_array_formula(
        "A2", '=SORT(UNIQUE(FILTER(tblAbsenceData[LOB],tblAbsenceData[LOB]<>"","")))',
    )
    ws.write("B1", "All")
    ws.write_dynamic_array_formula(
        "B2",
        '=SORT(UNIQUE(FILTER(tblAbsenceData[Team Leader],'
        '(tblAbsenceData[Team Leader]<>"")*IF(TEAM_VIEW!$K$5="All",1,'
        'tblAbsenceData[LOB]=TEAM_VIEW!$K$5),"")))',
    )
    ws.write("C1", "All")
    ws.write_dynamic_array_formula(
        "C2",
        '=SORT(UNIQUE(FILTER(tblAbsenceData[Agent Selector],'
        '(tblAbsenceData[Agent Selector]<>"")*IF(TEAM_VIEW!$K$5="All",1,'
        'tblAbsenceData[LOB]=TEAM_VIEW!$K$5)*IF(TEAM_VIEW!$N$5="All",1,'
        'tblAbsenceData[Team Leader]=TEAM_VIEW!$N$5),"")))',
    )
    wb.define_name("ABS_LOB_LIST", "=_LOOKUPS!$A$1:INDEX(_LOOKUPS!$A:$A,COUNTA(_LOOKUPS!$A:$A))")
    wb.define_name("ABS_TL_LIST", "=_LOOKUPS!$B$1:INDEX(_LOOKUPS!$B:$B,COUNTA(_LOOKUPS!$B:$B))")
    wb.define_name("ABS_AGENT_LIST", "=_LOOKUPS!$C$1:INDEX(_LOOKUPS!$C:$C,COUNTA(_LOOKUPS!$C:$C))")
    ws.hide()


def build_final_absence_product_workbook(
    conn: DatabaseConnection,
    config: Config,
    start: date,
    end: date,
    output: Path | None = None,
) -> Path:
    """Build the reviewed Absenteeism and Shrinkage collaboration report."""

    from .shared_feeds import publish_absence_feeds

    publish_absence_feeds(conn, config, start, end)
    rulebook = load_rulebook(config.home, config.business_rules)
    book, partial, target = _atomic_book(
        config, "absence", "ABSENTEEISM & SHRINKAGE", start, end, output,
    )
    available_start, available_end = conn.execute(
        "SELECT min(business_date), max(business_date) FROM mart.verint_final_absence_agent_day"
    ).fetchone()
    data_start = (
        available_start if isinstance(available_start, date)
        else date.fromisoformat(str(available_start)[:10]) if available_start else start
    )
    latest = (
        available_end if isinstance(available_end, date)
        else date.fromisoformat(str(available_end)[:10]) if available_end else end
    )
    totals = conn.execute(
        """SELECT coalesce(sum(planned_net_minutes),0),
                  coalesce(sum(final_absence_minutes),0),
                  coalesce(sum(final_vacation_minutes),0),
                  coalesce(sum(final_unpaid_minutes),0),
                  coalesce(sum(final_shrinkage_minutes),0),
                  coalesce(sum(final_unmapped_minutes),0),
                  coalesce(sum(CASE WHEN final_absence_day THEN 1 ELSE 0 END),0),
                  count(*),
                  coalesce(sum(CASE WHEN final_ledger_status IN ('CLEAR','ABSENCE_RECORDED')
                                    THEN planned_net_minutes ELSE 0 END),0),
                  coalesce(sum(CASE WHEN final_ledger_status NOT IN ('CLEAR','ABSENCE_RECORDED')
                                    THEN 1 ELSE 0 END),0)
           FROM mart.verint_final_absence_agent_day
           WHERE business_date BETWEEN ? AND ?""",
        [start, end],
    ).fetchone()
    (
        all_planned, all_absence, all_vacation, all_unpaid, all_shrinkage,
        unmapped, absence_days, agent_days, finalized_planned, exceptions,
    ) = totals
    finalized = conn.execute(
        """SELECT coalesce(sum(planned_net_minutes),0),
                  coalesce(sum(final_absence_minutes),0),
                  coalesce(sum(final_vacation_minutes),0),
                  coalesce(sum(final_unpaid_minutes),0),
                  coalesce(sum(final_shrinkage_minutes),0)
           FROM mart.verint_final_absence_agent_day
           WHERE business_date BETWEEN ? AND ?
             AND final_ledger_status IN ('CLEAR','ABSENCE_RECORDED')""",
        [start, end],
    ).fetchone()
    planned, absence, vacation, unpaid, shrinkage = finalized
    uncoded_empty = conn.execute(
        """SELECT count(*) FROM mart.verint_final_absence_agent_day
           WHERE business_date BETWEEN ? AND ?
             AND final_ledger_status='UNCODED_EMPTY_SHIFT'""",
        [start, end],
    ).fetchone()[0]
    status, status_text = _source_state(
        conn, ("fte", "start_end", "activities"), end, final=True,
    )
    if unmapped:
        status, status_text = "INCOMPLETE", f"{unmapped / 60:,.2f} Verint Activities hour(s) are unmapped"
    elif uncoded_empty:
        status, status_text = "INCOMPLETE", f"{uncoded_empty:,} scheduled shift(s) have no final Verint code or reliable work evidence"
    elif exceptions:
        status, status_text = "INCOMPLETE", f"{exceptions:,} case(s) still require review"

    period_rows = []
    for period in _comparison_periods(start, end):
        values = conn.execute(
            """SELECT coalesce(sum(planned_net_minutes),0),
                      coalesce(sum(final_absence_minutes),0),
                      coalesce(sum(final_vacation_minutes),0),
                      coalesce(sum(final_shrinkage_minutes),0)
               FROM mart.verint_final_absence_agent_day
               WHERE business_date BETWEEN ? AND ?
                 AND final_ledger_status IN ('CLEAR','ABSENCE_RECORDED')""",
            [period.start, period.end],
        ).fetchone()
        period_rows.append((
            period.label, period.start, period.end, values[0] / 60,
            values[1] / 60, _ratio(values[1], values[0]), values[2] / 60,
            values[3] / 60, _ratio(values[3], values[0]),
        ))
    coverage = _ratio(finalized_planned, all_planned)
    lob_dashboard_rows = conn.execute(
        """SELECT coalesce(lob,'UNMAPPED'),
                  CASE WHEN sum(planned_net_minutes)>0
                       THEN sum(final_absence_minutes)*1.0/sum(planned_net_minutes) END,
                  CASE WHEN sum(planned_net_minutes)>0
                       THEN sum(final_shrinkage_minutes)*1.0/sum(planned_net_minutes) END
           FROM mart.verint_final_absence_agent_day
           WHERE business_date BETWEEN ? AND ?
             AND final_ledger_status IN ('CLEAR','ABSENCE_RECORDED')
           GROUP BY coalesce(lob,'UNMAPPED') ORDER BY coalesce(lob,'UNMAPPED')""",
        [start, end],
    ).fetchall()
    daily_dashboard_rows = conn.execute(
        """SELECT business_date, sum(final_absence_minutes)/60.0,
                  sum(final_shrinkage_minutes)/60.0
           FROM mart.verint_final_absence_agent_day
           WHERE business_date BETWEEN ? AND ?
             AND final_ledger_status IN ('CLEAR','ABSENCE_RECORDED')
           GROUP BY business_date ORDER BY business_date""",
        [start, end],
    ).fetchall()
    dashboard_cases = conn.execute(
        """SELECT business_date,
                  coalesce(agent_name,'Agent') || ' [' || agent_id || ']',
                  coalesce(team_leader,'Unmapped'), coalesce(lob,'UNMAPPED'),
                  final_ledger_status, final_absence_minutes/60.0
           FROM mart.verint_final_absence_agent_day
           WHERE business_date BETWEEN ? AND ?
             AND (final_absence_day=true
                  OR final_ledger_status NOT IN ('CLEAR','ABSENCE_RECORDED'))
           ORDER BY CASE WHEN final_ledger_status IN ('CLEAR','ABSENCE_RECORDED')
                         THEN 1 ELSE 0 END,
                    business_date DESC, lob, team_leader, agent_name LIMIT 8""",
        [start, end],
    ).fetchall()
    dashboard = book.report.workbook.add_worksheet("DASHBOARD")
    render_v2_dashboard(
        book.report.workbook, dashboard,
        title="ABSENTEEISM & SHRINKAGE",
        filters=(
            ("Period", f"{start:%d %b} – {end:%d %b %Y}"),
            ("LOB", "All"), ("Team Leader", "All"), ("Agent", "All"),
        ),
        kpis=(
            ("Absence rate", _ratio(absence, planned), "percent"),
            ("Shrinkage rate", _ratio(shrinkage, planned), "percent"),
            ("Finalized coverage", coverage, "percent"),
            ("Review cases", exceptions, "integer"),
        ),
        left_chart=V2ChartSpec(
            "ABSENCE & SHRINKAGE BY LOB", "bar",
            tuple(str(row[0]) for row in lob_dashboard_rows),
            (
                ("Absence", tuple(row[1] for row in lob_dashboard_rows), COLORS["red"]),
                ("Shrinkage", tuple(row[2] for row in lob_dashboard_rows), COLORS["teal"]),
            ),
            "percent", 0,
        ),
        right_chart=V2ChartSpec(
            "DAILY ABSENCE & SHRINKAGE HOURS", "line",
            tuple(row[0].strftime("%d") if isinstance(row[0], date) else str(row[0]) for row in daily_dashboard_rows),
            (
                ("Absence hours", tuple(row[1] for row in daily_dashboard_rows), COLORS["red"]),
                ("Shrinkage hours", tuple(row[2] for row in daily_dashboard_rows), COLORS["teal"]),
            ),
        ),
        action_title="Prioritized absence review cases",
        action_headers=("DATE", "AGENT", "TEAM LEADER", "LOB", "RESULT STATUS", "ABSENCE H"),
        action_rows=dashboard_cases,
        action_kinds=("text", "text", "text", "text", "text", "decimal"),
        status="CHECK DATA" if status == "INCOMPLETE" else "IN DEVELOPMENT",
        status_kind="INCOMPLETE" if status == "INCOMPLETE" else "PROVISIONAL",
        status_note=status_text,
    )
    _add_absence_team_view(book, start, end, data_start, latest)

    team_headers, team_rows = _query(
        conn,
        """SELECT lob, team_leader, count(DISTINCT agent_id) AS agents,
                  count(*) AS agent_days,
                  sum(planned_net_minutes)/60.0 AS planned_hours,
                  sum(final_absence_minutes)/60.0 AS absence_hours,
                  CASE WHEN sum(planned_net_minutes)>0
                       THEN sum(final_absence_minutes)*1.0/sum(planned_net_minutes) END AS absence_rate,
                  sum(final_shrinkage_minutes)/60.0 AS shrinkage_hours,
                  CASE WHEN sum(planned_net_minutes)>0
                       THEN sum(final_shrinkage_minutes)*1.0/sum(planned_net_minutes) END AS shrinkage_rate,
                  sum(final_vacation_minutes)/60.0 AS vacation_hours,
                  sum(final_unpaid_minutes)/60.0 AS unpaid_hours,
                  sum(CASE WHEN final_ledger_status NOT IN ('CLEAR','ABSENCE_RECORDED')
                           THEN 1 ELSE 0 END) AS review_cases
           FROM mart.verint_final_absence_agent_day
           WHERE business_date BETWEEN ? AND ?
           GROUP BY lob, team_leader
           ORDER BY lob, team_leader""",
        [start, end],
    )
    team_sheet = book.table(
        "TEAM_SUMMARY", "Team absence and shrinkage",
        "Filter LOB or Team Leader. Rates use the summed hours shown in the same row.",
        team_headers, team_rows,
    )
    if team_rows:
        for name in ("absence_rate", "shrinkage_rate"):
            column = team_headers.index(name)
            team_sheet.conditional_format(
                4, column, 3 + len(team_rows), column,
                {"type": "3_color_scale", "min_color": COLORS["green_light"],
                 "mid_color": COLORS["amber_light"], "max_color": COLORS["red_light"]},
            )

    agent_headers, agent_rows = _query(
        conn,
        """SELECT lob, team_leader,
                  coalesce(agent_name,'Agent') || ' [' || agent_id || ']' AS agent_selector,
                  agent_id, agent_name, count(*) AS scheduled_days,
                  sum(planned_net_minutes)/60.0 AS planned_hours,
                  sum(final_absence_minutes)/60.0 AS absence_hours,
                  CASE WHEN sum(planned_net_minutes)>0
                       THEN sum(final_absence_minutes)*1.0/sum(planned_net_minutes) END AS absence_rate,
                  sum(final_shrinkage_minutes)/60.0 AS shrinkage_hours,
                  CASE WHEN sum(planned_net_minutes)>0
                       THEN sum(final_shrinkage_minutes)*1.0/sum(planned_net_minutes) END AS shrinkage_rate,
                  sum(final_vacation_minutes)/60.0 AS vacation_hours,
                  sum(final_unpaid_minutes)/60.0 AS unpaid_hours,
                  sum(CASE WHEN final_absence_day THEN 1 ELSE 0 END) AS absence_days,
                  sum(CASE WHEN final_ledger_status NOT IN ('CLEAR','ABSENCE_RECORDED')
                           THEN 1 ELSE 0 END) AS review_cases
           FROM mart.verint_final_absence_agent_day
           WHERE business_date BETWEEN ? AND ?
           GROUP BY lob, team_leader, agent_id, agent_name
           ORDER BY lob, team_leader, agent_name, agent_id""",
        [start, end],
    )
    book.table(
        "AGENT_RESULTS", "Agent absence and shrinkage",
        "Use a personal Sheet View before filtering LOB, Team Leader or Agent Selector.",
        agent_headers, agent_rows,
    )

    queue_headers, queue_source = _query(
        conn,
        """SELECT agent_day_key AS case_id, business_date,
                  agent_id, agent_name,
                  coalesce(agent_name,'Agent') || ' [' || agent_id || ']' AS agent_selector,
                  team_leader, ops_manager, lob, language,
                  final_ledger_status AS result_status,
                  planned_net_minutes/60.0 AS planned_net_hours,
                  final_absence_minutes/60.0 AS absence_hours,
                  final_shrinkage_minutes/60.0 AS shrinkage_hours,
                  final_unmapped_minutes/60.0 AS unmapped_hours
           FROM mart.verint_final_absence_agent_day
           WHERE business_date BETWEEN ? AND ?
             AND (final_absence_day=true
                  OR final_ledger_status NOT IN ('CLEAR','ABSENCE_RECORDED'))
           ORDER BY business_date DESC, lob, team_leader, agent_name""",
        [data_start, latest],
    )
    queue_rows = queue_source
    queue_sheet = book.table(
        "ACTION_QUEUE", "Absence action queue",
        "Read-only final-ledger cases and recorded absence days. Correct source coding in Verint; no action is entered back into WFMHub.",
        queue_headers, queue_rows or [tuple(None for _ in queue_headers)],
    )

    absence_headers, absence_components = _exclusive_final_components(
        conn, rulebook, start, end, "absence",
    )
    shrinkage_headers, shrinkage_components = _exclusive_final_components(
        conn, rulebook, start, end, "shrinkage",
    )
    book.table(
        "ABSENCE_COMPONENTS", "Absence components",
        "Overlapping final Verint Activities are counted once so component hours reconcile to the selected absence scope.",
        absence_headers, absence_components,
    )
    book.table(
        "SHRINKAGE_COMPONENTS", "Shrinkage components",
        "Overlapping final Verint Activities are counted once inside the shrinkage view. Do not add this table to Absence Components.",
        shrinkage_headers, shrinkage_components,
    )
    _add_absence_component_view(book)

    activity_headers, activity_rows = _query(
        conn,
        """SELECT business_date, lob, team_leader,
                  coalesce(agent_name,'Agent') || ' [' || agent_id || ']' AS agent_selector,
                  agent_id, agent_name, activity, category, event_start, event_end,
                  minutes, hours, counts_as_absence, counts_as_vacation,
                  counts_as_unpaid, counts_as_shrinkage, mapped,
                  evidence_type, event_key
           FROM mart.verint_final_absence_event
           WHERE business_date BETWEEN ? AND ?
           ORDER BY business_date, lob, team_leader, agent_name, event_start""",
        [data_start, latest],
    )
    book.table(
        "ACTIVITY_DETAIL", "Final Verint activity detail",
        "Exact mapped Verint Activities intervals. Use component sheets or ABSENCE_DATA for totals.",
        activity_headers, activity_rows,
    )

    data_headers, data_rows = _query(
        conn,
        """SELECT business_date, lob, team_leader,
                  coalesce(agent_name,'Agent') || ' [' || agent_id || ']' AS agent_selector,
                  agent_id, agent_name, ops_manager, language, location,
                  scheduled_minutes/60.0 AS scheduled_hours,
                  planned_net_minutes/60.0 AS planned_net_hours,
                  final_absence_minutes/60.0 AS absence_hours,
                  final_vacation_minutes/60.0 AS vacation_hours,
                  final_unpaid_minutes/60.0 AS unpaid_hours,
                  final_shrinkage_minutes/60.0 AS shrinkage_hours,
                  final_unmapped_minutes/60.0 AS unmapped_hours,
                  final_absence_rate AS absence_rate, final_absence_day,
                  final_ledger_status, agent_day_key AS case_id
           FROM mart.verint_final_absence_agent_day
           WHERE business_date BETWEEN ? AND ?
           ORDER BY business_date, lob, team_leader, agent_name""",
        [data_start, latest],
    )
    book.table(
        "ABSENCE_DATA", "Absence clean data",
        "One agent per day. Use a personal Sheet View before filtering LOB, Team Leader or Agent Selector.",
        data_headers, data_rows,
    )
    book.table(
        "HELP", "How to use this report", "A short operating guide for the shared workbook.",
        ["Step", "What to do", "Why"],
        [
            (1, "Run WFM Hub refresh and confirm StartEndTimes and Activities are current.", "Rebuilds the separate final post-day ledger without using Activities as presence."),
            (2, "Use TEAM_VIEW for period, LOB, Team Leader and Agent selection.", "Agent results, cases and components follow one selection."),
            (3, "Review Finalized coverage and every non-final ledger status before sharing totals.", "Empty, unmapped, provisional and partial rows must not dilute the rate."),
            (4, "Use COMPONENT_VIEW for totals and ACTIVITY_DETAIL for exact intervals.", "Raw intervals may overlap; KPI components remain separate."),
            (5, "Use Attendance Review separately when operational Agent Status gaps need diagnosis.", "Observed gaps do not overwrite the final Verint coding."),
            (6, "For a permanent shared file, connect ABSENCE_DATA, ACTION_QUEUE and ACTIVITY_DETAIL once to the fixed CSV feeds.", "Data > Refresh All updates read-only facts without replacing the workbook."),
        ],
    )
    book.definitions([
        ("Absence rate", "Final absence minutes / finalized planned net minutes", "Payroll and attendance result", "Incomplete cases are shown separately"),
        ("Shrinkage rate", "Final shrinkage minutes / finalized planned net minutes", "Capacity loss", "A parallel view; do not add to absence rate"),
        ("Finalized coverage", "Finalized planned minutes / all planned minutes", "Confidence in the headline", "Review when below 100%"),
        ("Review status", "Final ledger row with incomplete, unsupported or unmapped Verint coding", "Review completeness", "Never treated as zero absence"),
        ("Component", "One exclusive final Verint activity classification inside its KPI view", "Management breakdown", "Raw overlapping intervals are counted once"),
    ])
    _add_absence_lookups(book)
    book.audit(_audit_rows(
        conn, config, "absence", start, end,
        (
            ("Shared feed", str(config.feed / "Absenteeism"), "Updated with this report"),
            ("Template version", "absence-1.0.0", "Read-only Activities-final collaboration report contract"),
            ("All planned hours", all_planned / 60, f"{agent_days:,} agent-day row(s)"),
            ("All absence hours", all_absence / 60, "Includes review rows"),
            ("All shrinkage hours", all_shrinkage / 60, "Includes review rows"),
            ("All vacation hours", all_vacation / 60, "Includes review rows"),
            ("All unpaid hours", all_unpaid / 60, "Includes review rows"),
        ),
    ))
    return _finish(book, partial, target)


def _break_meal_control_rows(
    conn: DatabaseConnection,
    config: Config,
    start: date,
    completed_through: date,
) -> tuple[list[str], list[tuple[Any, ...]]]:
    """Build one evidence-gated break/meal result per completed agent-day."""

    headers = [
        "business_date", "lob", "team_leader", "agent_id", "agent_name",
        "scheduled_start", "scheduled_end", "status_coverage_percent",
        "break_minutes", "break_allowance_minutes", "break_overrun_minutes",
        "break_spells", "longest_break_minutes", "meal_minutes",
        "meal_allowance_minutes", "meal_overrun_minutes", "meal_spells",
        "longest_meal_minutes", "alert", "evidence",
    ]
    if completed_through < start:
        return headers, []
    attendance_headers, attendance_rows = _query(
        conn,
        """SELECT business_date, lob, team_leader, agent_id, agent_name,
                  scheduled_start, scheduled_end, planned_work_minutes,
                  status_covered_minutes, actual_evidence
           FROM mart.attendance_agent_day
           WHERE business_date BETWEEN ? AND ?
             AND assignment_type NOT IN ('Off','Planned absence')
             AND scheduled_start IS NOT NULL AND scheduled_end IS NOT NULL
             AND planned_work_minutes>0
           ORDER BY business_date, lob, team_leader, agent_name""",
        [start, completed_through],
    )
    timeline_headers, timeline_rows = _query(
        conn,
        """SELECT business_date, agent_id, actual_category,
                  segment_start, segment_end
           FROM mart.shift_timeline_segment
           WHERE business_date BETWEEN ? AND ?
             AND actual_category IN ('Break','Lunch')
             AND mismatch_type<>'WORK_DURING_TIME_OFF'
           ORDER BY business_date, agent_id, actual_category, segment_start""",
        [start, completed_through],
    )
    spells: dict[tuple[Any, str, str], list[tuple[datetime, datetime]]] = defaultdict(list)

    def stamp(value: Any) -> datetime:
        return value if isinstance(value, datetime) else datetime.fromisoformat(str(value))

    for values in timeline_rows:
        item = dict(zip(timeline_headers, values))
        spells[(
            item["business_date"], str(item["agent_id"]),
            str(item["actual_category"]),
        )].append((stamp(item["segment_start"]), stamp(item["segment_end"])))

    def merged(values: Sequence[tuple[datetime, datetime]]) -> list[tuple[datetime, datetime]]:
        output: list[list[datetime]] = []
        for left, right in sorted(values):
            if right <= left:
                continue
            if not output or left > output[-1][1]:
                output.append([left, right])
            else:
                output[-1][1] = max(output[-1][1], right)
        return [(left, right) for left, right in output]

    output: list[tuple[Any, ...]] = []
    alert_order = {
        "BREAK & MEAL EXCEEDED": 0, "BREAK EXCEEDED": 1,
        "MEAL EXCEEDED": 2, "INSUFFICIENT EVIDENCE": 3,
        "WITHIN LIMIT": 4,
    }
    for values in attendance_rows:
        row = dict(zip(attendance_headers, values))
        key = (row["business_date"], str(row["agent_id"]))
        break_spells = merged(spells.get((*key, "Break"), ()))
        meal_spells = merged(spells.get((*key, "Lunch"), ()))
        break_minutes = int(sum(
            (right - left).total_seconds() for left, right in break_spells
        ) // 60)
        meal_minutes = int(sum(
            (right - left).total_seconds() for left, right in meal_spells
        ) // 60)
        longest_break = max((
            int((right - left).total_seconds() // 60)
            for left, right in break_spells
        ), default=0)
        longest_meal = max((
            int((right - left).total_seconds() // 60)
            for left, right in meal_spells
        ), default=0)
        planned = float(row.get("planned_work_minutes") or 0)
        coverage = min(
            1.0, float(row.get("status_covered_minutes") or 0) / planned,
        ) if planned else None
        reliable = bool(
            coverage is not None
            and coverage >= config.rules.minimum_status_coverage
            and "AGENT_STATUS" in str(row.get("actual_evidence") or "").upper()
        )
        break_overrun = (
            max(0, break_minutes - config.rules.break_minutes)
            if reliable else None
        )
        meal_overrun = (
            max(0, meal_minutes - config.rules.lunch_minutes)
            if reliable else None
        )
        if not reliable:
            alert = "INSUFFICIENT EVIDENCE"
        elif break_overrun and meal_overrun:
            alert = "BREAK & MEAL EXCEEDED"
        elif break_overrun:
            alert = "BREAK EXCEEDED"
        elif meal_overrun:
            alert = "MEAL EXCEEDED"
        else:
            alert = "WITHIN LIMIT"
        output.append((
            row.get("business_date"), row.get("lob"), row.get("team_leader"),
            row.get("agent_id"), row.get("agent_name"),
            row.get("scheduled_start"), row.get("scheduled_end"), coverage,
            break_minutes, config.rules.break_minutes, break_overrun,
            len(break_spells), longest_break, meal_minutes,
            config.rules.lunch_minutes, meal_overrun, len(meal_spells),
            longest_meal, alert, row.get("actual_evidence"),
        ))
    output.sort(key=lambda row: (
        alert_order.get(str(row[18]), 9), row[0], str(row[1] or ""),
        str(row[2] or ""), str(row[4] or ""),
    ))
    return headers, output


def _add_attendance_review_control(
    book: DecisionWorkbook,
    status: str,
    status_text: str,
    start: date,
    end: date,
    completed_through: date,
    gap_count: int,
    gap_minutes: int,
    partial_count: int,
    missing: int,
    lob_rows: Sequence[Sequence[Any]],
) -> None:
    """Create the read-only residual-gap first screen."""

    wb = book.report.workbook
    ws = wb.add_worksheet("CONTROL")
    summary_headers = (
        "LOB", "Residual Gaps", "Residual Hours", "Agents",
        "Partial Verint",
    )
    display_rows = list(lob_rows) or [("No residual gaps", 0, 0, 0, 0)]
    lob_labels = tuple(str(row[0]) for row in display_rows)
    reconciliation_categories = ("Not in Verint", "Partial Verint", "Missing evidence")
    reconciliation_values = (gap_count - partial_count, partial_count, missing)
    formats = render_v2_dashboard(
        wb, ws,
        title="ATTENDANCE REVIEW",
        filters=(
            ("Period", f"{start:%d %b} – {end:%d %b %Y}"),
            ("Completed", f"Through {completed_through:%d %b}"),
            ("Evidence", "Status + LILO"),
            ("Workflow", "Correct in Verint"),
        ),
        kpis=(
            ("Review gaps", gap_count, "integer"),
            ("Gap hours", gap_minutes / 60 if gap_minutes else 0, "decimal"),
            ("Partly corrected", partial_count, "integer"),
            ("Missing evidence", missing, "integer"),
        ),
        left_chart=V2ChartSpec(
            "GAP HOURS BY LOB", "bar", lob_labels,
            (("Gap hours", tuple(float(row[2] or 0) for row in display_rows), COLORS["teal"]),),
        ),
        right_chart=V2ChartSpec(
            "RESIDUAL STATUS", "column", reconciliation_categories,
            (("Cases", reconciliation_values, COLORS["teal"]),),
        ),
        action_title="By-LOB review actions",
        action_headers=summary_headers,
        action_rows=display_rows,
        action_kinds=("text", "integer", "decimal", "integer", "alert"),
        status="VERINT REVIEW" if status != "INCOMPLETE" else "ACTION REQUIRED",
        status_kind=status,
        status_note=status_text,
    )
    ws.write_url(
        1, 23, "internal:'REVIEW BOARD'!A1", formats.filter_value,
        string="Open review board",
    )
    table_row = 34
    for column, header in enumerate(summary_headers):
        ws.write(table_row, column, header, book.report.header)
    for row_index, values in enumerate(display_rows, table_row + 1):
        for column, value in enumerate(values):
            fmt = book.report.decimal if column == 2 else (
                book.report.integer if column > 0 else book.report.body
            )
            ws.write(row_index, column, value, fmt)
    ws.add_table(
        table_row, 0, table_row + len(display_rows), len(summary_headers) - 1,
        {
            "name": "tblAttendanceReviewSummary",
            "style": "Table Style Light 9",
            "columns": [
                {"header": header, "header_format": book.report.header}
                for header in summary_headers
            ],
        },
    )
    book.tables.append(ModelTable("CONTROL", summary_headers, display_rows))
    notes_row = table_row + len(display_rows) + 3
    ws.merge_range(notes_row, 0, notes_row, 18, "OPERATING NOTES", book.report.section)
    for offset, note in enumerate((
        "Open REVIEW BOARD: each residual keeps SCHEDULE directly above ACTUAL and shows the exact interval still missing from final Verint Activities.",
        "Agent Status is primary attendance evidence. LILO fills sparse coverage; PTO/Away remains planned time and is never a no-show.",
        "Correct the exact residual in Verint. Export Activities and refresh WFMHub; fully covered gaps disappear automatically.",
    ), 1):
        ws.merge_range(notes_row + offset, 0, notes_row + offset, 18, note, book.report.note)
        ws.set_row(notes_row + offset, 22)
    ws.set_landscape()
    ws.fit_to_pages(1, 1)


def build_attendance_corrections_workbook(
    conn: DatabaseConnection,
    config: Config,
    start: date,
    end: date,
    output: Path | None = None,
) -> Path:
    """Build the read-only exact residual-gap review for Verint correction."""

    from .shift_view import add_review_board

    today = date.today()
    completed_through = min(end, today - timedelta(days=1))
    book, partial, target = _atomic_book(
        config, "corrections", "ATTENDANCE REVIEW", start, end, output,
    )
    # The rulebook is audit authority even when the selected period has no
    # reviewable gaps. Load it before conditional validation setup so an empty
    # Attendance Review can still be generated and audited.
    rulebook = load_rulebook(config.home, config.business_rules)
    gap_count, gap_minutes, agents, partial_count = conn.execute(
        """SELECT count(*), coalesce(sum(gap_minutes),0), count(DISTINCT agent_id),
                  coalesce(sum(CASE WHEN verint_reconciliation='PARTIALLY_IN_VERINT' THEN 1 ELSE 0 END),0)
           FROM mart.correction_candidate
           WHERE business_date BETWEEN ? AND ? AND business_date<?
             AND gap_start IS NOT NULL AND gap_end IS NOT NULL""",
        [start, end, today],
    ).fetchone()
    missing = conn.execute(
        """SELECT count(*) FROM mart.attendance_agent_day
           WHERE business_date BETWEEN ? AND ? AND business_date<?
             AND assignment_type NOT IN ('Off','Planned absence')
             AND attendance_result IN
               ('Schedule parse error','Data not loaded','Missing actual evidence',
                'Incomplete actual evidence','No schedule overlap')""",
        [start, end, today],
    ).fetchone()[0]
    status, status_text = _source_state(
        conn, ("fte", "start_end", "lilo", "agent_status", "activities"), completed_through,
        final=True,
    )
    if gap_count or missing:
        status = "INCOMPLETE"
        status_text = (
            f"{gap_count:,} exact residual gap(s) are not fully covered in Verint; "
            f"{missing:,} scheduled row(s) lack complete evidence"
        )
    lob_rows = conn.execute(
        """SELECT coalesce(lob,'UNMAPPED') AS lob,
                  count(*) AS exact_gaps,
                  coalesce(sum(gap_minutes),0)/60.0 AS gap_hours,
                  count(DISTINCT agent_id) AS agents,
                  coalesce(sum(CASE WHEN verint_reconciliation='PARTIALLY_IN_VERINT' THEN 1 ELSE 0 END),0) AS partial_count
           FROM mart.correction_candidate
           WHERE business_date BETWEEN ? AND ? AND business_date<?
             AND gap_start IS NOT NULL AND gap_end IS NOT NULL
           GROUP BY coalesce(lob,'UNMAPPED')
           ORDER BY gap_hours DESC, lob""",
        [start, end, today],
    ).fetchall()
    _add_attendance_review_control(
        book, status, status_text, start, end, completed_through,
        gap_count, gap_minutes, partial_count, missing, lob_rows,
    )
    headers, rows = _query(
        conn,
        """SELECT c.correction_id AS gap_id, c.business_date, c.agent_id,
                  c.agent_name, c.team_leader, c.lob,
                  c.detected_issue, c.gap_start AS exact_start,
                  c.gap_end AS exact_end, c.gap_minutes AS minutes,
                  c.suggested_activity AS suggested_verint_activity,
                  c.verint_activity AS final_activity_found,
                  c.verint_overlap_minutes AS final_overlap_minutes,
                  c.verint_reconciliation AS residual_status,
                  c.confidence, c.observed_source,
                  a.assignment AS published_assignment
           FROM mart.correction_candidate c
           LEFT JOIN mart.attendance_agent_day a
             ON a.business_date=c.business_date AND a.agent_id=c.agent_id
           WHERE c.business_date BETWEEN ? AND ? AND c.business_date<?
             AND c.gap_start IS NOT NULL AND c.gap_end IS NOT NULL
           ORDER BY c.business_date, c.priority, c.gap_minutes DESC,
                    c.agent_id, c.gap_start""",
        [start, end, today],
    )
    capacity_mapping = load_capacity_mapping(config.capacity_mapping)
    correction_indexes = {header: index for index, header in enumerate(headers)}
    headers.extend([
        "management_lob", "planning_group", "staff_type",
        "capacity_mapping_status",
    ])
    rows = [
        (*row, *(
            lambda mapped: (
                mapped.management_lob, mapped.planning_group,
                mapped.staff_type, mapped.status,
            )
        )(capacity_mapping.map_schedule(
            str(row[correction_indexes["lob"]] or ""),
            str(row[correction_indexes["published_assignment"]] or ""),
        )))
        for row in rows
    ]
    timeline_headers, timeline_rows = _query(
        conn,
        """SELECT t.segment_key, t.business_date, t.agent_id, t.agent_name, t.team_leader,
                  t.ops_manager, t.lob, t.language, t.scheduled_start, t.scheduled_end,
                  t.segment_start, t.segment_end, t.segment_minutes, t.planned_state,
                  t.actual_status, t.actual_category, t.mismatch_type, t.is_gap,
                  t.observed_source, t.source_file, t.evaluation_as_of,
                  a.assignment AS published_assignment
           FROM mart.shift_timeline_segment t
           LEFT JOIN mart.attendance_agent_day a
             ON a.business_date=t.business_date AND a.agent_id=t.agent_id
           WHERE t.business_date BETWEEN ? AND ? AND t.business_date<?
             AND EXISTS (
                 SELECT 1 FROM mart.correction_candidate c
                 WHERE c.business_date=t.business_date AND c.agent_id=t.agent_id
                   AND c.gap_start IS NOT NULL AND c.gap_end IS NOT NULL
             )
           ORDER BY t.business_date, t.agent_id, t.segment_start""",
        [start, end, today],
    )
    timeline_indexes = {
        header: index for index, header in enumerate(timeline_headers)
    }
    timeline_headers.extend([
        "management_lob", "planning_group", "staff_type",
        "capacity_mapping_status",
    ])
    timeline_rows = [
        (*row, *(
            lambda mapped: (
                mapped.management_lob, mapped.planning_group,
                mapped.staff_type, mapped.status,
            )
        )(capacity_mapping.map_schedule(
            str(row[timeline_indexes["lob"]] or ""),
            str(row[timeline_indexes["published_assignment"]] or ""),
        )))
        for row in timeline_rows
    ]
    add_review_board(
        book.report, headers, rows,
        [dict(zip(timeline_headers, row)) for row in timeline_rows],
        start, completed_through,
    )
    book.tables.append(ModelTable("REVIEW BOARD", headers, rows))
    break_meal_headers, break_meal_rows = _break_meal_control_rows(
        conn, config, start, completed_through,
    )
    break_meal_sheet = book.table(
        "BREAK & MEAL", "Daily break and meal control",
        "Completed shifts only. Totals come from Agent Status inside the scheduled shift; incomplete coverage is never reported as zero compliance.",
        break_meal_headers, break_meal_rows,
    )
    if break_meal_rows:
        alert_column = break_meal_headers.index("alert")
        break_meal_sheet.conditional_format(
            4, alert_column, 3 + len(break_meal_rows), alert_column,
            {
                "type": "text", "criteria": "containing", "value": "EXCEEDED",
                "format": book.report.error,
            },
        )
        incomplete_format = book.report.workbook.add_format({
            "font_name": "Aptos", "font_size": 10, "bold": True,
            "font_color": COLORS["amber"], "bg_color": COLORS["amber_light"],
        })
        break_meal_sheet.conditional_format(
            4, alert_column, 3 + len(break_meal_rows), alert_column,
            {
                "type": "text", "criteria": "containing",
                "value": "INSUFFICIENT EVIDENCE", "format": incomplete_format,
            },
        )
    evidence_sheet = book.table(
        "EVIDENCE", "Exact attendance evidence",
        "Exact schedule and observed segments behind every colored REVIEW BOARD cell. The 15-minute visual never changes these boundaries.",
        timeline_headers, timeline_rows,
    )
    evidence_sheet.hide()
    book.definitions([
        ("Exact gap", "One continuous scheduled interval without accepted working evidence", "Human review unit", "Never rounded or merged across a return"),
        ("Gap ID", "Stable date/agent/start/end/issue key", "Residual evidence key", "No value is imported from Excel"),
        ("Residual", "Observed gap portion not covered by final Verint Activities", "Correction backlog", "Uses exact timestamps"),
        ("Partially in Verint", "A final activity covers only part of the observed gap", "Only the remaining fragment is shown", "Overlap remains visible for audit"),
        ("Not in Verint", "No final correction activity overlaps the observed gap", "Full exact interval remains", "Never silently treated as absence or zero"),
        ("Logged", "Observed working, auxiliary, break or lunch evidence according to its separate color", "Visual shift context", "Exact raw state remains in EVIDENCE"),
        ("Schedule band", "Scheduled work and PTO/Away drawn directly above ACTUAL", "Plan-versus-actual comparison", "It is visual context and is never imported"),
        ("Gap", "Scheduled interval without accepted working evidence", "Review candidate", "The exact start/end at the left remain authoritative"),
        ("Current-day tail", "Unfinished part of today's shift", "No review row", "Never classified as early leave"),
        ("Break and meal control", "Daily Agent Status totals compared with configured allowances", "Operational overrun alert", "Completed shifts with insufficient coverage remain unknown"),
    ])
    book.audit(_audit_rows(
        conn, config, "corrections", start, end,
        (
            ("Completed-date cutoff", completed_through, "Today is excluded"),
            ("Residual authority", "Agent Status/LILO minus final Verint Activities", "Read-only automatic reconciliation"),
            ("Classification authority", rulebook.sha256, rulebook.version),
        ),
    ))
    return _finish(book, partial, target)
