"""Configuration-driven daily service flashes reconstructed from Book1."""

from __future__ import annotations

import csv
import math
import re
from collections import defaultdict
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any, Iterable, Sequence

from .config import Config
from .database import DatabaseConnection
from .mapping import QueueMapping, load_queue_mapping
from .metrics import MetricCatalog, evaluate_metric, load_metric_catalog
from .report_packs import publish_report, report_current_path
from .reports import COLORS
from .rules import Rulebook, load_rulebook
from .service_profiles import ServiceProfile, load_service_profiles
from .template_reports import DecisionWorkbook


def _marks(values: Sequence[str]) -> str:
    return ",".join("?" for _ in values)


def _as_datetime(value: Any) -> datetime | None:
    if isinstance(value, datetime):
        return value
    if value is None:
        return None
    try:
        return datetime.fromisoformat(str(value))
    except ValueError:
        return None


def _ratio(numerator: float | int | None, denominator: float | int | None) -> float | None:
    if numerator is None or denominator is None or float(denominator) == 0:
        return None
    return float(numerator) / float(denominator)


def _profile_comparison_scopes(
    profile: ServiceProfile,
    mapping: QueueMapping,
) -> tuple[str, ...]:
    return mapping.comparison_scopes_for(profile.service_scopes)


def _profile_method(
    catalog: MetricCatalog,
    profile: ServiceProfile,
    metric_id: str,
    on_date: date,
):
    methods = {}
    for service_scope in profile.service_scopes:
        for source_system in profile.flash_source_systems:
            method = catalog.method_for(
                metric_id, on_date,
                {"lob": service_scope, "source_system": source_system},
            )
            if method is None:
                raise ValueError(
                    f"Flash profile {profile.profile_id!r} has no {metric_id!r} "
                    f"method for {service_scope}/{source_system} on {on_date}"
                )
            methods[(method.method_id, method.effective_from, method.priority)] = method
    if len(methods) != 1:
        names = ", ".join(key[0] for key in methods)
        raise ValueError(
            f"Flash profile {profile.profile_id!r} crosses incompatible "
            f"{metric_id} methods: {names}"
        )
    return next(iter(methods.values()))


def _aggregate(
    rows: Iterable[dict[str, Any]],
    profile: ServiceProfile,
    catalog: MetricCatalog,
    on_date: date,
) -> dict[str, Any] | None:
    values = list(rows)
    if not values:
        return None
    components = {
        name: sum(float(row.get(name) or 0) for row in values)
        for name in (
            "offered", "answered", "abandoned", "short_abandoned",
            "abandoned_within_target", "answered_within_target",
            "handled_seconds",
        )
    }
    service = evaluate_metric(
        _profile_method(catalog, profile, profile.service_level_metric, on_date),
        components,
    )
    availability = evaluate_metric(
        _profile_method(catalog, profile, profile.availability_metric, on_date),
        components,
    )
    abandon = evaluate_metric(
        _profile_method(catalog, profile, "abandon_rate", on_date), components,
    )
    aht = evaluate_metric(
        _profile_method(catalog, profile, profile.aht_metric, on_date), components,
    )
    return {
        **components,
        "raw_offered": components["offered"],
        # Storm displays every inbound queue entry as Total Entered. Its custom
        # C/(A+B-D) SLA removes lost calls from 5 seconds up to the configured
        # target. Lost calls below 5 seconds remain in the denominator.
        "offered": components["offered"],
        "business_offered": max(
            0.0,
            components["offered"] - components["abandoned_within_target"],
        ),
        "storm_a_lost": components["abandoned"],
        "storm_b_connected": components["answered"],
        "storm_c_connected_within_sla": components["answered_within_target"],
        "storm_d_lost_within_sla": components["abandoned_within_target"],
        "storm_sl_denominator": max(
            0.0,
            components["offered"] - components["abandoned_within_target"],
        ),
        "service_level": service.value,
        "service_target": service.method.target,
        "service_method": service.method.method_id,
        "service_state": service.state,
        "availability": availability.value,
        "abandon_rate": abandon.value,
        "aht_seconds": aht.value,
    }


def _profile_rows(
    conn: DatabaseConnection,
    profile: ServiceProfile,
    start: date,
    end: date,
) -> list[dict[str, Any]]:
    queue_filter = (
        f"upper(queue) IN ({_marks(profile.flash_queues)})"
        if profile.flash_queues
        else f"service_scope IN ({_marks(profile.service_scopes)})"
    )
    queue_values = (
        [queue.upper() for queue in profile.flash_queues]
        if profile.flash_queues else list(profile.service_scopes)
    )
    cursor = conn.execute(
        f"""SELECT business_date, hour_start, source_system, service_scope,
                   comparison_scope, queue, designation, language, offered,
                   answered, abandoned, short_abandoned,
                   abandoned_within_target,
                   answered_within_target, talk_seconds, hold_seconds,
                   wrap_seconds, handled_seconds, service_level,
                   service_availability, abandon_rate, aht_seconds, call_legs,
                   transferred_legs, source_files
            FROM mart.call_service_hour
            WHERE business_date BETWEEN ? AND ?
              AND {queue_filter}
              AND source_system IN ({_marks(profile.flash_source_systems)})
            ORDER BY business_date, hour_start, queue""",
        [start, end, *queue_values, *profile.flash_source_systems],
    )
    headers = [item[0] for item in cursor.description]
    return [dict(zip(headers, row)) for row in cursor.fetchall()]


def _included_in_flash_total(
    profile: ServiceProfile,
    row: dict[str, Any],
) -> bool:
    return profile.includes_flash_queue(row.get("queue"))


def _forecast_by_hour(
    conn: DatabaseConnection,
    profile: ServiceProfile,
    mapping: QueueMapping,
    report_day: date,
) -> dict[int, float]:
    scopes = _profile_comparison_scopes(profile, mapping)
    rows = conn.execute(
        f"""SELECT hour_start, sum(volume_forecast)
            FROM mart.forecast_hour
            WHERE business_date=? AND comparison_scope IN ({_marks(scopes)})
            GROUP BY hour_start ORDER BY hour_start""",
        [report_day, *scopes],
    ).fetchall()
    result: dict[int, float] = {}
    for hour_start, value in rows:
        parsed = _as_datetime(hour_start)
        if parsed is not None and value is not None:
            result[parsed.hour] = float(value)
    return result


def _workforce_by_hour(
    conn: DatabaseConnection,
    profile: ServiceProfile,
    report_day: date,
) -> dict[int, dict[str, float]]:
    """Return rough intraday absent HC from the governed attendance timeline.

    Agent Status is already the primary source behind this timeline and LILO is
    only its fallback/control.  A future segment or missing evidence is never
    turned into absence here.
    """

    lob_marks = _marks(profile.staffing_lobs)
    attendance = conn.execute(
        f"""SELECT agent_id, scheduled_start, scheduled_end, assignment_type,
                   planned_work_minutes
            FROM mart.attendance_agent_day
            WHERE business_date=? AND lob IN ({lob_marks})""",
        [report_day, *profile.staffing_lobs],
    ).fetchall()
    timeline = conn.execute(
        f"""SELECT agent_id, segment_start, segment_end
            FROM mart.shift_timeline_segment
            WHERE business_date=? AND lob IN ({lob_marks}) AND is_gap=true""",
        [report_day, *profile.staffing_lobs],
    ).fetchall()
    output: dict[int, dict[str, float]] = {}
    for hour in range(profile.operating_start_hour, profile.operating_end_hour + 1):
        left = datetime.combine(report_day, datetime.min.time()) + timedelta(hours=hour)
        right = left + timedelta(hours=1)
        planned = {
            str(agent_id)
            for agent_id, raw_start, raw_end, assignment_type, planned_work in attendance
            if str(assignment_type or "").casefold() != "off"
            and float(planned_work or 0) > 0
            and (_as_datetime(raw_start) or right) < right
            and (_as_datetime(raw_end) or left) > left
        }
        absent = {
            str(agent_id)
            for agent_id, raw_start, raw_end in timeline
            if str(agent_id) in planned
            and (_as_datetime(raw_start) or right) < right
            and (_as_datetime(raw_end) or left) > left
        }
        output[hour] = {
            "absence_hc": float(len(absent)),
        }
    return output


def _attendance_pulse(
    conn: DatabaseConnection,
    profile: ServiceProfile,
    report_day: date,
    cutoff: int | None,
) -> tuple[list[dict[str, Any]], dict[str, Any]]:
    """Build one reconcilable same-day attendance pulse for a Flash.

    The pulse uses the same mart as Attendance Callouts.  The full-shift
    timeline only supplies the agent's state at the checkpoint, so a temporary
    mid-shift gap is visible even when the agent logged back in later.
    """

    lob_marks = _marks(profile.staffing_lobs)
    cursor = conn.execute(
        f"""SELECT business_date, agent_id, agent_name, team_leader,
                   ops_manager, lob, language, scheduled_start, scheduled_end,
                   assignment_type, attendance_result, call_action,
                   requires_call, actual_first_seen, actual_last_seen,
                   uncoded_late_minutes, actual_evidence, source_loaded,
                   is_provisional, evaluation_as_of
            FROM mart.attendance_agent_day
            WHERE business_date=? AND lob IN ({lob_marks})
              AND assignment_type NOT IN ('Off','Planned absence')
            ORDER BY scheduled_start, lob, team_leader, agent_name""",
        [report_day, *profile.staffing_lobs],
    )
    headers = [item[0] for item in cursor.description]
    attendance = [dict(zip(headers, row)) for row in cursor.fetchall()]
    evaluation_points = [
        stamp for stamp in (_as_datetime(row.get("evaluation_as_of")) for row in attendance)
        if stamp is not None and stamp.date() == report_day
    ]
    if evaluation_points:
        checkpoint = max(evaluation_points)
    elif cutoff is not None:
        checkpoint = datetime.combine(report_day, datetime.min.time()) + timedelta(
            hours=cutoff, minutes=59, seconds=59,
        )
    else:
        checkpoint = datetime.combine(report_day, datetime.min.time()) + timedelta(
            hours=23, minutes=59, seconds=59,
        )

    timeline_cursor = conn.execute(
        f"""SELECT agent_id, segment_start, segment_end, planned_state,
                   actual_category, mismatch_type, is_gap, observed_source
            FROM mart.shift_timeline_segment
            WHERE business_date=? AND lob IN ({lob_marks})
            ORDER BY agent_id, segment_start""",
        [report_day, *profile.staffing_lobs],
    )
    timeline_headers = [item[0] for item in timeline_cursor.description]
    segments_by_agent: dict[str, list[dict[str, Any]]] = defaultdict(list)
    for values in timeline_cursor.fetchall():
        segment = dict(zip(timeline_headers, values))
        segments_by_agent[str(segment["agent_id"])].append(segment)

    pulse: list[dict[str, Any]] = []
    for row in attendance:
        start = _as_datetime(row.get("scheduled_start"))
        end = _as_datetime(row.get("scheduled_end"))
        scheduled_now = bool(start and end and start <= checkpoint < end)
        state_checkpoint = checkpoint - timedelta(seconds=1)
        current_segment = next((
            item for item in segments_by_agent.get(str(row["agent_id"]), [])
            if (_as_datetime(item.get("segment_start")) or state_checkpoint) <= state_checkpoint
            < (_as_datetime(item.get("segment_end")) or state_checkpoint)
        ), None)
        state = "NOT SCHEDULED NOW"
        if scheduled_now:
            if current_segment is not None and bool(current_segment.get("is_gap")):
                state = "ABSENT NOW"
            elif current_segment is not None and str(
                current_segment.get("actual_category") or ""
            ).upper() in {"PRODUCTIVE", "AUXILIARY", "BREAK", "LUNCH", "LILO_PRESENT"}:
                state = "PRESENT NOW"
            else:
                first_seen = _as_datetime(row.get("actual_first_seen"))
                last_seen = _as_datetime(row.get("actual_last_seen"))
                if first_seen and last_seen and first_seen <= checkpoint <= last_seen:
                    state = "PRESENT NOW"
                elif bool(row.get("source_loaded")):
                    state = "ABSENT NOW"
                else:
                    state = "UNKNOWN"
        late_today = float(row.get("uncoded_late_minutes") or 0) > 0
        absence_now = state == "ABSENT NOW"
        call_action = str(row.get("call_action") or "NONE")
        call_now = bool(row.get("requires_call")) or absence_now
        if absence_now and call_action == "NONE":
            call_action = "CHECK_CURRENT_GAP"
        pulse.append({
            "flash": profile.label, "checkpoint": checkpoint,
            **row, "scheduled_now": scheduled_now,
            "attendance_state": state, "present_now": state == "PRESENT NOW",
            "absence_now": absence_now, "late_today": late_today,
            "call_now": call_now, "pulse_action": call_action,
            "current_evidence": (
                current_segment.get("observed_source")
                if current_segment is not None else row.get("actual_evidence")
            ),
        })
    summary = {
        "checkpoint": checkpoint,
        "scheduled_now": sum(bool(row["scheduled_now"]) for row in pulse),
        "present_now": sum(bool(row["present_now"]) for row in pulse),
        "absence_hc": sum(bool(row["absence_now"]) for row in pulse),
        "call_now": sum(bool(row["call_now"]) for row in pulse),
        "late_today": sum(bool(row["late_today"]) for row in pulse),
        "unknown_now": sum(row["attendance_state"] == "UNKNOWN" for row in pulse),
    }
    return pulse, summary


def _hourly_model(
    conn: DatabaseConnection,
    profile: ServiceProfile,
    mapping: QueueMapping,
    metrics: MetricCatalog,
    report_day: date,
    all_rows: Sequence[dict[str, Any]],
) -> tuple[list[dict[str, Any]], dict[str, Any] | None, dict[str, dict[str, Any] | None], int | None]:
    rows = [
        row for row in all_rows
        if str(row["business_date"])[:10] == report_day.isoformat()
        and _included_in_flash_total(profile, row)
    ]
    by_hour: dict[int, list[dict[str, Any]]] = defaultdict(list)
    for row in rows:
        hour = _as_datetime(row.get("hour_start"))
        if hour is not None and 0 <= hour.hour <= 23:
            by_hour[hour.hour].append(row)
    forecast = _forecast_by_hour(conn, profile, mapping, report_day)
    workforce = _workforce_by_hour(conn, profile, report_day)
    cutoff = max(by_hour) if by_hour else None
    hourly: list[dict[str, Any]] = []
    for hour in range(profile.operating_start_hour, profile.operating_end_hour + 1):
        source = by_hour.get(hour, [])
        aggregate = _aggregate(source, profile, metrics, report_day)
        group_values = {
            group.label: _aggregate(
                [row for row in source if profile.group_for(row.get("queue")) == group.label],
                profile, metrics, report_day,
            )
            for group in profile.groups
        }
        forecast_value = forecast.get(hour)
        actual_value = aggregate["offered"] if aggregate is not None else None
        state = (
            ("FUTURE" if report_day >= date.today() else "AFTER CUTOFF")
            if cutoff is not None and hour > cutoff
            else "NO MAPPED CALLS" if aggregate is None
            else "FORECAST MISSING" if forecast_value is None
            else "READY"
        )
        hourly.append({
            "profile_id": profile.profile_id, "flash": profile.label,
            "business_date": report_day, "hour": hour,
            "hour_label": f"{hour:02d}:00", "forecast": forecast_value,
            "volume_variance": (
                actual_value - forecast_value
                if actual_value is not None and forecast_value is not None
                else None
            ),
            "forecast_attainment": _ratio(actual_value, forecast_value),
            "data_state": state, **(aggregate or {
                "offered": None, "answered": None, "abandoned": None,
                "short_abandoned": None, "abandoned_within_target": None,
                "answered_within_target": None,
                "handled_seconds": None, "service_level": None,
                "service_target": _profile_method(
                    metrics, profile, profile.service_level_metric, report_day,
                ).target,
                "service_method": _profile_method(
                    metrics, profile, profile.service_level_metric, report_day,
                ).method_id,
                "service_state": "NO_SAMPLE", "availability": None,
                "abandon_rate": None, "aht_seconds": None,
            }),
            "groups": group_values,
            **workforce.get(hour, {}),
        })
    total_source = [row for hour in by_hour if cutoff is not None and hour <= cutoff for row in by_hour[hour]]
    total = _aggregate(total_source, profile, metrics, report_day)
    if total is not None:
        forecast_hours = [
            value for hour, value in forecast.items()
            if cutoff is not None and hour <= cutoff
        ]
        total["forecast"] = (
            sum(float(value) for value in forecast_hours)
            if forecast_hours else None
        )
        total["forecast_attainment"] = _ratio(total["offered"], total["forecast"])
        total["volume_variance"] = (
            total["offered"] - total["forecast"]
            if total.get("forecast") is not None else None
        )
        absence_values = [
            row.get("absence_hc") for row in hourly
            if row.get("absence_hc") is not None
        ]
        total["absence_hc"] = max(absence_values) if absence_values else None
    group_totals = {
        group.label: _aggregate(
            [row for row in total_source if profile.group_for(row.get("queue")) == group.label],
            profile, metrics, report_day,
        )
        for group in profile.groups
    }
    return hourly, total, group_totals, cutoff


def _formats(book: DecisionWorkbook) -> dict[str, Any]:
    add = book.report.workbook.add_format
    return {
        "card_label": add({
            "font_name": "Aptos", "font_size": 9, "bold": True,
            "font_color": COLORS["white"], "bg_color": COLORS["teal"],
            "align": "center", "valign": "vcenter", "border": 1,
            "border_color": COLORS["thin"],
        }),
        "card_value": add({
            "font_name": "Aptos Display", "font_size": 18, "bold": True,
            "font_color": COLORS["dark"], "bg_color": COLORS["white"],
            "align": "center", "valign": "vcenter", "border": 1,
            "border_color": COLORS["thin"],
        }),
        "card_integer": add({
            "font_name": "Aptos Display", "font_size": 18, "bold": True,
            "font_color": COLORS["dark"], "bg_color": COLORS["white"],
            "align": "center", "valign": "vcenter", "num_format": "#,##0",
            "border": 1, "border_color": COLORS["thin"],
        }),
        "card_percent": add({
            "font_name": "Aptos Display", "font_size": 18, "bold": True,
            "font_color": COLORS["dark"], "bg_color": COLORS["white"],
            "align": "center", "valign": "vcenter", "num_format": "0.0%",
            "border": 1, "border_color": COLORS["thin"],
        }),
        "card_seconds": add({
            "font_name": "Aptos Display", "font_size": 18, "bold": True,
            "font_color": COLORS["dark"], "bg_color": COLORS["white"],
            "align": "center", "valign": "vcenter", "num_format": '0 "s"',
            "border": 1, "border_color": COLORS["thin"],
        }),
        "card_note": add({
            "font_name": "Aptos", "font_size": 8, "font_color": COLORS["muted"],
            "bg_color": COLORS["canvas"], "align": "center", "valign": "vcenter",
            "text_wrap": True, "border": 1, "border_color": COLORS["thin"],
        }),
        "total": add({
            "font_name": "Aptos", "font_size": 10, "bold": True,
            "font_color": COLORS["white"], "bg_color": COLORS["dark"],
            "border": 1, "border_color": COLORS["thin"],
        }),
        "total_integer": add({
            "font_name": "Aptos", "font_size": 10, "bold": True,
            "font_color": COLORS["white"], "bg_color": COLORS["dark"],
            "num_format": "#,##0", "border": 1, "border_color": COLORS["thin"],
        }),
        "total_percent": add({
            "font_name": "Aptos", "font_size": 10, "bold": True,
            "font_color": COLORS["white"], "bg_color": COLORS["dark"],
            "num_format": "0.0%", "border": 1, "border_color": COLORS["thin"],
        }),
        "total_seconds": add({
            "font_name": "Aptos", "font_size": 10, "bold": True,
            "font_color": COLORS["white"], "bg_color": COLORS["dark"],
            "num_format": '0 "s"', "border": 1, "border_color": COLORS["thin"],
        }),
    }


def _write_card(
    ws,
    formats: dict[str, Any],
    column: int,
    label: str,
    value: Any,
    kind: str,
    note: str,
) -> None:
    ws.merge_range(4, column, 4, column + 1, label.upper(), formats["card_label"])
    value_format = formats.get(f"card_{kind}", formats["card_value"])
    ws.merge_range(5, column, 6, column + 1, value if value is not None else "—", value_format)
    ws.merge_range(7, column, 7, column + 1, note, formats["card_note"])


def _table_value_format(book: DecisionWorkbook, header: str):
    lowered = header.casefold()
    if (
        lowered.startswith("sl ")
        or any(token in lowered for token in (
            "availability", "routed rate", "deviation", "service level", "tsl",
        ))
    ):
        return book.report.percent
    if "aht" in lowered:
        return book.report.workbook.add_format({
            "font_name": "Aptos", "font_size": 10, "font_color": COLORS["dark"],
            "num_format": '0 "s"', "bottom": 1, "bottom_color": COLORS["thin"],
        })
    if any(token in lowered for token in ("forecast", "actual", "handled", "variance", "absence hc", "abs hc")):
        return book.report.integer
    return book.report.body


def _flash_columns(
    profile: ServiceProfile,
    hourly: Sequence[dict[str, Any]],
) -> tuple[list[str], list[list[Any]], int, int, int]:
    if profile.flash_layout == "oem_split":
        headers = [
            "Hour", "Volume Forecasted", "Volume Variance", "Volume Ford",
            "Volume Chery", "Volume Toyota", "SL Ford", "SL Chery", "SL Toyota",
            "Routed Rate Ford", "Routed Rate Chery", "Routed Rate Toyota", "AHT",
            "ABS HC",
        ]
        rows = []
        for row in hourly:
            ford = row["groups"].get("Ford") or {}
            chery = row["groups"].get("Chery") or {}
            toyota = row["groups"].get("Toyota") or {}
            rows.append([
                row["hour_label"], row["forecast"], row.get("volume_variance"), ford.get("offered"),
                chery.get("offered"), toyota.get("offered"),
                ford.get("service_level"), chery.get("service_level"),
                toyota.get("service_level"), ford.get("availability"),
                chery.get("availability"), toyota.get("availability"),
                row["aht_seconds"], row.get("absence_hc"),
            ])
        return headers, rows, 1, 3, 6
    headers = [
        "Hour", "Volume Forecasted", "Volume Actual", "Volume Variance",
        "Volume Handled", "Volume Handled in SL", "Deviation", "Routed Rate",
        "TSL", "AHT", "ABS HC", "Data State",
    ]
    rows = [[
        row["hour_label"], row["forecast"], row["offered"], row.get("volume_variance"),
        row["answered"], row["answered_within_target"], row["forecast_attainment"],
        row["availability"], row["service_level"], row["aht_seconds"],
        row.get("absence_hc"), row["data_state"],
    ] for row in hourly]
    return headers, rows, 1, 2, 8


def _flash_cards(
    profile: ServiceProfile,
    total: dict[str, Any] | None,
    groups: dict[str, dict[str, Any] | None],
    pulse: dict[str, Any],
) -> list[tuple[str, Any, str, str]]:
    value = total or {}
    if profile.flash_layout == "oem_split":
        ford = groups.get("Ford") or {}
        chery = groups.get("Chery") or {}
        toyota = groups.get("Toyota") or {}
        return [
            ("Routed Rate OEM", value.get("availability"), "percent", "Total routed / total entered"),
            ("SLA OEM", value.get("service_level"), "percent", value.get("service_method") or "Configured method"),
            ("SLA Ford", ford.get("service_level"), "percent", "APFR Ford"),
            ("SLA Chery", chery.get("service_level"), "percent", "APFR Chery"),
            ("SLA Toyota", toyota.get("service_level"), "percent", "APFR Toyota and Lexus"),
            ("Deviation", value.get("forecast_attainment"), "percent", "Actual / forecast through cutoff"),
            ("Volume Variance", value.get("volume_variance"), "integer", "Actual - forecast through cutoff"),
            ("AHT", value.get("aht_seconds"), "seconds", "Weighted handled seconds"),
            ("ABS HC", pulse.get("absence_hc"), "integer", "Scheduled but absent at attendance checkpoint"),
        ]
    return [
        ("Forecast", value.get("forecast"), "integer", "Through latest actual hour"),
        ("Actual", value.get("offered"), "integer", "Inbound queue entries"),
        ("Handled", value.get("answered"), "integer", "Routed queue entries"),
        ("Handled in SL", value.get("answered_within_target"), "integer", "Answered inside threshold"),
        ("Deviation", value.get("forecast_attainment"), "percent", "Actual / forecast through cutoff"),
        ("Volume Variance", value.get("volume_variance"), "integer", "Actual - forecast through cutoff"),
        ("Routed Rate", value.get("availability"), "percent", "Routed / entered"),
        ("TSL", value.get("service_level"), "percent", value.get("service_method") or "Configured method"),
        ("AHT", value.get("aht_seconds"), "seconds", "Weighted handled seconds"),
        ("ABS HC", pulse.get("absence_hc"), "integer", "Scheduled but absent at attendance checkpoint"),
    ]


def _add_flash_sheet(
    book: DecisionWorkbook,
    profile: ServiceProfile,
    report_day: date,
    hourly: Sequence[dict[str, Any]],
    total: dict[str, Any] | None,
    group_totals: dict[str, dict[str, Any] | None],
    pulse: dict[str, Any],
    cutoff: int | None,
) -> None:
    ws = book.report.workbook.add_worksheet(profile.flash_sheet)
    ws.hide_gridlines(2)
    ws.set_tab_color(COLORS["gold"])
    ws.set_zoom(85)
    ws.merge_range("A1:T1", f"FLASH  /  {profile.label.upper()}", book.report.title)
    cutoff_text = f"through {cutoff:02d}:59" if cutoff is not None else "no mapped queue entries"
    ws.merge_range(
        "A2:T2",
        f"{report_day:%Y-%m-%d}  |  Call-by-Call actuals {cutoff_text}  |  generated {book.generated:%Y-%m-%d %H:%M}",
        book.report.subtitle,
    )
    ws.set_row(0, 34)
    ws.set_row(1, 21)
    formats = _formats(book)
    for index, card in enumerate(_flash_cards(profile, total, group_totals, pulse)):
        _write_card(ws, formats, index * 2, *card)
    headers, display_rows, forecast_col, actual_col, sl_col = _flash_columns(profile, hourly)
    table_row = 10
    for column, header in enumerate(headers):
        ws.write(table_row, column, header, book.report.header)
    for row_index, values in enumerate(display_rows, table_row + 1):
        for column, value in enumerate(values):
            fmt = _table_value_format(book, headers[column])
            if value is None:
                ws.write_blank(row_index, column, None, fmt)
            else:
                ws.write(row_index, column, value, fmt)
    if display_rows:
        table_name = "tbl" + re.sub(r"[^A-Za-z0-9]", "", profile.flash_sheet)
        ws.add_table(table_row, 0, table_row + len(display_rows), len(headers) - 1, {
            "name": table_name,
            "style": "Table Style Light 9",
            "columns": [{"header": header, "header_format": book.report.header} for header in headers],
        })
    total_row = table_row + 1 + len(display_rows)
    label = f"Day through {cutoff:02d}:59" if cutoff is not None else "Day total unavailable"
    ws.write(total_row, 0, label, formats["total"])
    total_values = total or {}
    if profile.flash_layout == "oem_split":
        ford = group_totals.get("Ford") or {}
        chery = group_totals.get("Chery") or {}
        toyota = group_totals.get("Toyota") or {}
        values = [
            total_values.get("forecast"), total_values.get("volume_variance"),
            ford.get("offered"), chery.get("offered"),
            toyota.get("offered"), ford.get("service_level"),
            chery.get("service_level"), toyota.get("service_level"),
            ford.get("availability"), chery.get("availability"),
            toyota.get("availability"), total_values.get("aht_seconds"),
            total_values.get("absence_hc"),
        ]
    else:
        values = [
            total_values.get("forecast"), total_values.get("offered"),
            total_values.get("volume_variance"), total_values.get("answered"),
            total_values.get("answered_within_target"),
            total_values.get("forecast_attainment"), total_values.get("availability"),
            total_values.get("service_level"), total_values.get("aht_seconds"),
            total_values.get("absence_hc"),
            "READY" if total else "INCOMPLETE",
        ]
    for column, value in enumerate(values, 1):
        header = headers[column]
        lowered = header.casefold()
        fmt = (
            formats["total_percent"] if lowered.startswith("sl ") or any(
                token in lowered
                for token in ("availability", "routed rate", "deviation", "tsl")
            )
            else formats["total_seconds"] if "aht" in lowered
            else formats["total_integer"] if isinstance(value, (int, float)) and value is not None
            else formats["total"]
        )
        if value is None:
            ws.write_blank(total_row, column, None, fmt)
        else:
            ws.write(total_row, column, value, fmt)

    data_last_row = table_row + len(display_rows)
    if display_rows:
        chart = book.report.workbook.add_chart({"type": "column"})
        volume_series = (
            (
                ("Forecast", forecast_col, COLORS["muted"]),
                ("Ford", 3, COLORS["teal"]),
                ("Chery", 4, COLORS["red"]),
                ("Toyota", 5, COLORS["gold"]),
            )
            if profile.flash_layout == "oem_split"
            else (
                ("Forecast", forecast_col, COLORS["muted"]),
                ("Actual", actual_col, COLORS["teal"]),
            )
        )
        for label_text, column, color in volume_series:
            chart.add_series({
                "name": label_text,
                "categories": [profile.flash_sheet, table_row + 1, 0, data_last_row, 0],
                "values": [profile.flash_sheet, table_row + 1, column, data_last_row, column],
                "fill": {"color": color}, "border": {"none": True},
            })
        line = book.report.workbook.add_chart({"type": "line"})
        line.add_series({
            "name": "TSL",
            "categories": [profile.flash_sheet, table_row + 1, 0, data_last_row, 0],
            "values": [profile.flash_sheet, table_row + 1, sl_col, data_last_row, sl_col],
            "y2_axis": True, "line": {"color": COLORS["gold"], "width": 2.25},
        })
        chart.combine(line)
        chart.set_title({"name": "Hourly demand and service"})
        chart.set_legend({"position": "bottom"})
        chart.set_y2_axis({"num_format": "0%", "min": 0, "max": 1})
        chart.set_chartarea({"border": {"none": True}, "fill": {"color": COLORS["white"]}})
        chart.set_plotarea({"border": {"none": True}, "fill": {"color": COLORS["white"]}})
        ws.insert_chart(table_row, len(headers) + 1, chart, {"x_scale": 1.05, "y_scale": 1.0})

    target = total_values.get("service_target")
    if target is not None and display_rows:
        ws.conditional_format(table_row + 1, sl_col, data_last_row, sl_col, {
            "type": "cell", "criteria": "<", "value": target,
            "format": book.report.error,
        })
    for column, header in enumerate(headers):
        width = 13
        if header == "Hour":
            width = 10
        elif header == "Data State":
            width = 20
        elif len(header) > 17:
            width = 18
        ws.set_column(column, column, width)
    ws.set_column(len(headers) + 1, len(headers) + 10, 11)
    ws.freeze_panes(table_row + 1, 1)
    ws.set_landscape()
    ws.fit_to_pages(1, 1)
    ws.repeat_rows(0, table_row)


def _add_control_sheet(
    book: DecisionWorkbook,
    summaries: Sequence[
        tuple[ServiceProfile, dict[str, Any] | None, int | None, dict[str, Any]]
    ],
    report_day: date,
    rulebook: Rulebook,
) -> None:
    ws = book.report.workbook.add_worksheet("CONTROL")
    ws.hide_gridlines(2)
    ws.set_tab_color(COLORS["gold"])
    ws.merge_range("A1:P1", "SERVICE FLASH CONTROL", book.report.title)
    ws.merge_range(
        "A2:P2",
        f"Daily control for {report_day:%Y-%m-%d}  |  actuals: mapped inbound Call-by-Call queue entries  |  forecast: Verint",
        book.report.subtitle,
    )
    headers = [
        "Flash", "Cutoff", "Forecast", "Actual", "Volume Variance", "Handled",
        "Deviation", "Routed Rate", "TSL", "AHT", "Scheduled Now",
        "Present Now", "ABS HC", "Call Now", "Late Today", "Status",
    ]
    for col, header in enumerate(headers):
        ws.write(4, col, header, book.report.header)
    for offset, (profile, total, cutoff, pulse) in enumerate(summaries, 5):
        value = total or {}
        status = (
            "CALL DATA MISSING" if total is None
            else "FORECAST MISSING" if value.get("forecast") is None
            else "READY"
        )
        row = [
            profile.label, f"{cutoff:02d}:59" if cutoff is not None else None,
            value.get("forecast"), value.get("offered"),
            value.get("volume_variance"), value.get("answered"),
            value.get("forecast_attainment"),
            value.get("availability"), value.get("service_level"),
            value.get("aht_seconds"), pulse.get("scheduled_now"),
            pulse.get("present_now"), pulse.get("absence_hc"),
            pulse.get("call_now"), pulse.get("late_today"), status,
        ]
        for col, item in enumerate(row):
            fmt = (
                book.report.percent if col in {6, 7, 8}
                else book.report.decimal if col == 9
                else book.report.integer if col in {2, 3, 4, 5, 10, 11, 12, 13, 14}
                else book.report.body
            )
            if item is None:
                ws.write_blank(offset, col, None, fmt)
            else:
                ws.write(offset, col, item, fmt)
        ws.write_url(offset, 0, f"internal:'{profile.flash_sheet}'!A1", book.report.body, profile.label)
    if summaries:
        ws.add_table(4, 0, 4 + len(summaries), len(headers) - 1, {
            "name": "tblFlashControl", "style": "Table Style Light 9",
            "columns": [{"header": header, "header_format": book.report.header} for header in headers],
        })
        ws.conditional_format(5, 15, 4 + len(summaries), 15, {
            "type": "text", "criteria": "containing", "value": "MISSING",
            "format": book.report.error,
        })
        chart = book.report.workbook.add_chart({"type": "column"})
        for label, column, color in (
                ("Forecast", 2, COLORS["muted"]), ("Actual", 3, COLORS["teal"]),
        ):
            chart.add_series({
                "name": label,
                "categories": ["CONTROL", 5, 0, 4 + len(summaries), 0],
                "values": ["CONTROL", 5, column, 4 + len(summaries), column],
                "fill": {"color": color}, "border": {"none": True},
            })
        chart.set_title({"name": "Forecast versus actual through cutoff"})
        chart.set_legend({"position": "bottom"})
        chart.set_chartarea({"border": {"none": True}})
        ws.insert_chart("A18", chart, {"x_scale": 1.25, "y_scale": 1.1})
    notes = [
        "Open a Flash name to jump to its hourly sheet.",
        "Deviation follows the reference workbook: actual offered / forecast through the latest actual hour.",
        f"Storm SLA = C / (A + B - D): connected within {rulebook.target_seconds}s / "
        f"(lost + connected - lost from {rulebook.short_abandon_seconds}s to {rulebook.target_seconds}s).",
        "Every Flash shows 00:00-23:00; headline totals reset at midnight.",
        "Routed Rate = total routed / total entered; it does not remove any lost calls.",
        "Scheduled Now, Present Now, ABS HC, Call Now and Late Today reconcile to ATTENDANCE PULSE at its stated checkpoint.",
        "No mapped calls and missing forecasts remain blank; the workbook never turns missing evidence into zero.",
    ]
    ws.write("A10", "OPERATING NOTES", book.report.section)
    for index, note in enumerate(notes, 10):
        ws.merge_range(index, 0, index, 15, note, book.report.note)
    ws.set_column("A:A", 23)
    ws.set_column("B:P", 15)
    ws.freeze_panes(5, 0)
    ws.set_landscape()
    ws.fit_to_pages(1, 1)


def _flat_hour_rows(
    profiles: Sequence[ServiceProfile],
    hourly_by_profile: dict[str, Sequence[dict[str, Any]]],
) -> tuple[list[str], list[tuple[Any, ...]]]:
    headers = [
        "profile_id", "flash", "business_date", "hour", "forecast", "offered",
        "volume_variance",
        "business_offered", "answered", "abandoned", "short_abandoned",
        "abandoned_within_target", "answered_within_target",
        "storm_a_lost", "storm_b_connected", "storm_c_connected_within_sla",
        "storm_d_lost_within_sla", "storm_sl_denominator",
        "forecast_attainment", "availability", "service_level", "service_target",
        "service_method", "abandon_rate", "aht_seconds", "absence_hc",
        "ford_offered", "ford_answered",
        "ford_service_level", "chery_offered", "chery_answered",
        "chery_service_level", "toyota_offered", "toyota_answered",
        "toyota_service_level", "data_state",
    ]
    rows: list[tuple[Any, ...]] = []
    for profile in profiles:
        for row in hourly_by_profile[profile.profile_id]:
            ford = row["groups"].get("Ford") or {}
            chery = row["groups"].get("Chery") or {}
            toyota = row["groups"].get("Toyota") or {}
            rows.append(tuple([
                row.get("profile_id"), row.get("flash"), row.get("business_date"),
                row.get("hour"), row.get("forecast"), row.get("offered"),
                row.get("volume_variance"),
                row.get("business_offered"), row.get("answered"),
                row.get("abandoned"), row.get("short_abandoned"),
                row.get("abandoned_within_target"),
                row.get("answered_within_target"), row.get("storm_a_lost"),
                row.get("storm_b_connected"),
                row.get("storm_c_connected_within_sla"),
                row.get("storm_d_lost_within_sla"),
                row.get("storm_sl_denominator"), row.get("forecast_attainment"),
                row.get("availability"), row.get("service_level"),
                row.get("service_target"), row.get("service_method"),
                row.get("abandon_rate"), row.get("aht_seconds"),
                row.get("absence_hc"),
                ford.get("offered"), ford.get("answered"), ford.get("service_level"),
                chery.get("offered"), chery.get("answered"), chery.get("service_level"),
                toyota.get("offered"), toyota.get("answered"), toyota.get("service_level"),
                row.get("data_state"),
            ]))
    return headers, rows


def _queue_diagnosis_rows(
    profiles: Sequence[ServiceProfile],
    source_by_profile: dict[str, Sequence[dict[str, Any]]],
    metrics: MetricCatalog,
    report_day: date,
    cutoff_by_profile: dict[str, int | None],
) -> tuple[list[str], list[tuple[Any, ...]]]:
    """Explain which exact configured queues are helping or hurting SLA."""

    headers = [
        "flash", "queue", "flash_group", "data_state", "total_entered",
        "total_routed", "lost", "connected_in_sla", "lost_5_to_sla",
        "sla_denominator", "tsl", "target", "gap_to_target",
        "late_connected", "lost_in_denominator", "contacts_to_target",
        "routed_rate", "diagnosis", "source_files",
    ]
    output: list[tuple[Any, ...]] = []
    for profile in profiles:
        cutoff = cutoff_by_profile.get(profile.profile_id)
        day_rows = [
            row for row in source_by_profile.get(profile.profile_id, ())
            if str(row.get("business_date"))[:10] == report_day.isoformat()
            and (cutoff is None or (
                _as_datetime(row.get("hour_start")) is not None
                and _as_datetime(row.get("hour_start")).hour <= cutoff
            ))
        ]
        by_queue: dict[str, list[dict[str, Any]]] = defaultdict(list)
        for row in day_rows:
            by_queue[str(row.get("queue") or "").strip().upper()].append(row)
        configured = (
            list(profile.flash_queues)
            if profile.flash_queues else sorted(by_queue)
        )
        target = _profile_method(
            metrics, profile, profile.service_level_metric, report_day,
        ).target
        for queue in configured:
            rows = by_queue.get(str(queue).strip().upper(), [])
            aggregate = _aggregate(rows, profile, metrics, report_day)
            if aggregate is None:
                values: dict[str, Any] = {}
                diagnosis = "NO DATA — confirm zero demand or source coverage"
                state = "NO DATA"
            else:
                values = aggregate
                state = "READY"
                late_connected = max(
                    0.0, values["answered"] - values["answered_within_target"],
                )
                lost_in_denominator = max(
                    0.0, values["abandoned"] - values["abandoned_within_target"],
                )
                if values["service_level"] is None:
                    diagnosis = "NO SLA SAMPLE"
                elif target is not None and values["service_level"] >= target:
                    diagnosis = "AT / ABOVE TARGET"
                elif late_connected > lost_in_denominator * 1.5:
                    diagnosis = "BELOW TARGET — late routed contacts dominate"
                elif lost_in_denominator > late_connected * 1.5:
                    diagnosis = "BELOW TARGET — lost contacts dominate"
                else:
                    diagnosis = "BELOW TARGET — mixed delay and loss"
            late_connected = (
                max(0.0, values["answered"] - values["answered_within_target"])
                if aggregate is not None else None
            )
            lost_in_denominator = (
                max(0.0, values["abandoned"] - values["abandoned_within_target"])
                if aggregate is not None else None
            )
            contacts_to_target = None
            if aggregate is not None and target is not None:
                deficit = (
                    target * values["storm_sl_denominator"]
                    - values["answered_within_target"]
                )
                if deficit <= 0:
                    contacts_to_target = 0
                elif target < 1:
                    # Each additional good contact increases both C and the
                    # denominator: (C+x)/(denominator+x) >= target.
                    contacts_to_target = math.ceil(deficit / (1 - target))
            sources = "; ".join(sorted({
                value.strip()
                for row in rows
                for value in str(row.get("source_files") or "").split(";")
                if value.strip()
            })) or None
            output.append((
                profile.label, queue, profile.group_for(queue), state,
                values.get("offered"), values.get("answered"),
                values.get("abandoned"), values.get("answered_within_target"),
                values.get("abandoned_within_target"),
                values.get("storm_sl_denominator"), values.get("service_level"),
                target, (
                    values["service_level"] - target
                    if values.get("service_level") is not None and target is not None
                    else None
                ), late_connected, lost_in_denominator, contacts_to_target,
                values.get("availability"), diagnosis, sources,
            ))
    return headers, output


def _attendance_pulse_rows(
    profiles: Sequence[ServiceProfile],
    pulse_by_profile: dict[str, Sequence[dict[str, Any]]],
) -> tuple[list[str], list[tuple[Any, ...]]]:
    headers = [
        "flash", "checkpoint", "business_date", "lob", "agent_id",
        "agent_name", "team_leader", "scheduled_start", "scheduled_end",
        "attendance_state", "scheduled_now", "present_now", "abs_hc",
        "call_now", "late_today", "pulse_action", "attendance_result",
        "uncoded_late_minutes", "current_evidence", "source_loaded",
    ]
    output: list[tuple[Any, ...]] = []
    for profile in profiles:
        rows = sorted(
            pulse_by_profile.get(profile.profile_id, ()),
            key=lambda row: (
                not bool(row.get("call_now")),
                not bool(row.get("absence_now")),
                str(row.get("team_leader") or ""),
                str(row.get("agent_name") or ""),
            ),
        )
        for row in rows:
            output.append((
                profile.label, row.get("checkpoint"), row.get("business_date"),
                row.get("lob"), row.get("agent_id"), row.get("agent_name"),
                row.get("team_leader"), row.get("scheduled_start"),
                row.get("scheduled_end"), row.get("attendance_state"),
                "YES" if row.get("scheduled_now") else "NO",
                "YES" if row.get("present_now") else "NO",
                1 if row.get("absence_now") else 0,
                "YES" if row.get("call_now") else "NO",
                "YES" if row.get("late_today") else "NO",
                row.get("pulse_action"), row.get("attendance_result"),
                row.get("uncoded_late_minutes"), row.get("current_evidence"),
                "YES" if row.get("source_loaded") else "NO",
            ))
    return headers, output


def build_service_flashes_workbook(
    conn: DatabaseConnection,
    config: Config,
    start: date,
    end: date,
    output: Path | None = None,
    profile_id: str | None = None,
) -> Path:
    """Build all Storm Flash layouts from mapped Call-by-Call queue entries."""

    catalog = load_service_profiles(config.home, config.service_profiles)
    profiles = [profile for profile in catalog.profiles if profile.active_on(end)]
    profiles.sort(key=lambda profile: (profile.display_order, profile.profile_id))
    if profile_id is not None:
        # Retain CLI validation without changing the stable all-Flash workbook.
        catalog.select(profile_id, end)
    mapping = load_queue_mapping(config.queue_mapping)
    metrics = load_metric_catalog(config.home, config.metric_catalog)
    rulebook = load_rulebook(config.home, config.business_rules)
    generated = datetime.now()
    target = (output or report_current_path(config, "service")).resolve()
    target.parent.mkdir(parents=True, exist_ok=True)
    partial = target.with_name(f"{target.stem}.partial{target.suffix}")
    book = DecisionWorkbook(
        partial, config, "service", "SERVICE FLASHES", start, end, generated,
    )
    hourly_by_profile: dict[str, Sequence[dict[str, Any]]] = {}
    source_by_profile: dict[str, Sequence[dict[str, Any]]] = {}
    pulse_by_profile: dict[str, Sequence[dict[str, Any]]] = {}
    summaries: list[
        tuple[ServiceProfile, dict[str, Any] | None, int | None, dict[str, Any]]
    ] = []
    group_totals_by_profile: dict[str, dict[str, dict[str, Any] | None]] = {}
    try:
        for profile in profiles:
            source_rows = _profile_rows(conn, profile, start, end)
            source_by_profile[profile.profile_id] = source_rows
            hourly, total, group_totals, cutoff = _hourly_model(
                conn, profile, mapping, metrics, end, source_rows,
            )
            pulse_rows, pulse_summary = _attendance_pulse(
                conn, profile, end, cutoff,
            )
            hourly_by_profile[profile.profile_id] = hourly
            pulse_by_profile[profile.profile_id] = pulse_rows
            group_totals_by_profile[profile.profile_id] = group_totals
            summaries.append((profile, total, cutoff, pulse_summary))
        _add_control_sheet(book, summaries, end, rulebook)
        for profile, total, cutoff, pulse_summary in summaries:
            _add_flash_sheet(
                book, profile, end, hourly_by_profile[profile.profile_id], total,
                group_totals_by_profile[profile.profile_id], pulse_summary, cutoff,
            )
        headers, rows = _flat_hour_rows(profiles, hourly_by_profile)
        book.table(
            "FLASH_DATA", "Flash clean hourly data",
            "One profile/hour for the report day. Actual demand counts mapped inbound queue entries like Storm Total Entered.",
            headers, rows,
        )
        pulse_headers, pulse_rows = _attendance_pulse_rows(
            profiles, pulse_by_profile,
        )
        pulse_sheet = book.table(
            "ATTENDANCE PULSE", "Attendance pulse and call list",
            "Same governed attendance rows as Attendance Callouts. Filter Call Now=YES for action; ABS HC is only the reliable scheduled-and-absent checkpoint count.",
            pulse_headers, pulse_rows,
        )
        if pulse_rows:
            state_col = pulse_headers.index("attendance_state")
            pulse_sheet.conditional_format(4, state_col, 3 + len(pulse_rows), state_col, {
                "type": "text", "criteria": "containing", "value": "ABSENT",
                "format": book.report.error,
            })
        cutoff_by_profile = {
            profile.profile_id: cutoff
            for profile, _total, cutoff, _pulse in summaries
        }
        diagnosis_headers, diagnosis_rows = _queue_diagnosis_rows(
            profiles, source_by_profile, metrics, end, cutoff_by_profile,
        )
        diagnosis_sheet = book.table(
            "QUEUE DIAGNOSIS", "Queue-level service diagnosis",
            "Exact configured Flash queues through each Flash cutoff. Sort Contacts to Target or filter Diagnosis to locate the service-level drivers.",
            diagnosis_headers, diagnosis_rows,
        )
        if diagnosis_rows:
            diagnosis_col = diagnosis_headers.index("diagnosis")
            diagnosis_sheet.conditional_format(
                4, diagnosis_col, 3 + len(diagnosis_rows), diagnosis_col,
                {"type": "text", "criteria": "containing", "value": "BELOW TARGET", "format": book.report.error},
            )
        with mapping.file.open("r", encoding="utf-8-sig", newline="") as handle:
            source_mapping = list(csv.DictReader(handle))
        mapping_headers = [
            "mapping_type", "source_system", "source_value", "service_scope",
            "comparison_scope", "designation", "flash_group", "used_by_flash",
        ]
        mapping_rows = []
        for row in source_mapping:
            if str(row.get("mapping_type") or "").strip().lower() != "queue":
                continue
            mapped = mapping.map_actual(
                row.get("source_system"), row.get("source_value"), None, None,
            )
            memberships = [
                profile for profile in profiles
                if profile.includes_flash_queue(row.get("source_value"))
            ]
            flash_group = " | ".join(
                f"{profile.label}: {profile.group_for(row.get('source_value'))}"
                for profile in memberships
            ) or None
            used = " | ".join(profile.label for profile in memberships)
            mapping_rows.append((
                row.get("mapping_type"), row.get("source_system"),
                row.get("source_value"), mapped.service_scope,
                mapped.comparison_scope, mapped.designation, flash_group,
                used or "NO",
            ))
        book.table(
            "QUEUE_MAP", "Queue-to-Flash control",
            "Source mapping plus exact screenshot-derived Flash membership. Edit flash_queues in service_profiles.toml; queue_mapping.csv controls source-to-scope mapping.",
            mapping_headers, mapping_rows,
        )
        exception_headers = [
            "flash", "business_date", "hour", "issue", "actual", "forecast",
            "service_level", "target", "action",
        ]
        exception_rows = []
        for profile in profiles:
            for row in hourly_by_profile[profile.profile_id]:
                issue = None
                action = None
                if row["data_state"] == "FORECAST MISSING":
                    issue, action = "FORECAST MISSING", "Load or map the Verint forecast extract"
                elif row["data_state"] == "NO MAPPED CALLS" and row["hour"] <= (max((item["hour"] for item in hourly_by_profile[profile.profile_id] if item["offered"] is not None), default=-1)):
                    issue, action = "NO MAPPED CALLS", "Confirm zero demand or review queue mapping"
                elif row.get("service_level") is not None and row.get("service_target") is not None and row["service_level"] < row["service_target"]:
                    issue, action = "BELOW TSL TARGET", "Review demand, staffing and long waits"
                if issue:
                    exception_rows.append((
                        profile.label, row["business_date"], row["hour_label"], issue,
                        row["offered"], row["forecast"], row["service_level"],
                        row["service_target"], action,
                    ))
        book.table(
            "EXCEPTIONS", "Flash exceptions",
            "Only missing evidence and below-target service intervals requiring review.",
            exception_headers, exception_rows,
        )
        oem_groups = next(
            (
                profile.flash_total_groups for profile in profiles
                if profile.flash_layout == "oem_split"
            ),
            (),
        )
        book.definitions([
            ("Volume Actual", "Every inbound entry into a mapped Flash queue", "Storm Total Entered", "A transfer entering another mapped queue is another queue entry"),
            ("OEM visible scope", " and ".join(oem_groups) or "Every configured group", "Matches the Storm OEM platform", "Four Ford FR queues plus the Chery and Toyota/Lexus queues shown in Storm"),
            ("Volume Handled", "Inbound queue entry routed to an agent", "Storm Total Routed", "Agent may be outside the FTE roster; the queue is the service boundary"),
            ("Response time", "Total Queue Wait Time + Ringing Duration", "Storm threshold clock", "Reproduced from the Call-by-Call business reference"),
            ("Volume Handled in SL", f"Routed queue entry with response time < {rulebook.target_seconds} seconds", "SLA numerator", "Threshold is editable in wfm_rules.toml"),
            ("Short Abandon", f"Unanswered queue entry with response time < {rulebook.short_abandon_seconds} seconds", "Retained in the Storm SLA denominator", "Storm excludes only lost calls from 5 seconds to the SLA target"),
            ("Abandoned in SL", f"Unanswered queue entry with response time from {rulebook.short_abandon_seconds} to < {rulebook.target_seconds} seconds", "Storm variable D", "Subtracted from lost plus connected calls"),
            ("Deviation", "Total entered / forecast through the latest actual hour", "Demand tracking", "Uses the visible Storm volume"),
            ("Volume Variance", "Total entered minus forecast through the latest actual hour", "Absolute demand gap", "Positive means demand is above forecast"),
            ("Routed Rate", "Total routed / total entered", "Service availability", "Matches the Storm screenshot; not agent availability or adherence"),
            ("TSL", "Connected within target / (lost + connected - lost from 5 seconds to target)", "Storm C/(A+B-D)", "The supplied Storm custom-equation screen is the business authority"),
            ("AHT", "Sum of inbound talk + hold + wrap / routed queue entries", "Workload", "Weighted; never an average of hourly averages"),
            ("ABS HC", "Distinct scheduled agents in a reliable attendance gap at the checkpoint/hour", "Rough same-day capacity signal", "PTO/Away and missing evidence are not counted; exact treatment stays in Attendance Review"),
            ("Attendance Pulse", "The same attendance-agent-day mart used by Attendance Callouts plus checkpoint timeline state", "Call and late action list", "It is operational context, not a final absence decision"),
            ("Queue Diagnosis", "The exact configured queue counters behind each Flash", "Find queues pulling TSL below target", "It does not change the validated Flash queue membership or Storm formula"),
        ])
        clean_calls, unique_interactions = conn.execute(
            """SELECT count(*), count(DISTINCT interaction_key)
               FROM core.clean_call_leg WHERE business_date BETWEEN ? AND ?""",
            [start, end],
        ).fetchone()
        mapped_offered = conn.execute(
            """SELECT coalesce(sum(offered),0) FROM mart.call_service_hour
               WHERE business_date BETWEEN ? AND ?""",
            [start, end],
        ).fetchone()[0]
        book.audit([
            ("Report", "service", "All Book1 Flash profiles"),
            ("Report day", end, "Visible Flash sheets use the selected end date"),
            ("Selected data period", f"{start} to {end}", "Model boundary"),
            ("Clean Call-by-Call legs", clean_calls, "After stable call-leg deduplication"),
            ("Unique clean interactions", unique_interactions, "Before queue mapping"),
            ("Mapped Flash offered", mapped_offered, "Inbound queue-entry count"),
            ("Queue mapping", mapping.file.name, mapping.sha256),
            ("Service profiles", catalog.version, catalog.sha256),
            ("OEM visible groups", " | ".join(oem_groups) or "ALL", "Configured in service_profiles.toml"),
            ("Metric catalog", metrics.version, metrics.sha256),
            ("Rulebook", rulebook.version, rulebook.sha256),
            ("Storm SLA equation", "C / (A + B - D)", "A=lost; B=connected; C=connected within SLA; D=lost from 5 seconds to SLA target"),
            ("Storm SLA target", f"{rulebook.target_seconds} seconds", f"Lost-call exclusion band: {rulebook.short_abandon_seconds} to < {rulebook.target_seconds} seconds"),
            ("Day view", "00:00-23:00", "Actual headline totals run from midnight through the latest mapped call hour"),
            ("Attendance source", "mart.attendance_agent_day + mart.shift_timeline_segment", "Agent Status primary; LILO fallback/control"),
            ("Design reference", "TOLEARN/Book1.xlsx", "Four pasted Flash references reconstructed as native Excel"),
            ("Prepared by", "Anass ASSRI", "WFM"),
        ])
        book.close()
        publish_report(config, "service", partial, target, generated)
    except Exception:
        try:
            book.report.close()
        except Exception:
            pass
        partial.unlink(missing_ok=True)
        raise
    return target
