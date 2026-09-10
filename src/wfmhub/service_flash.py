"""Configuration-driven real-time management control workbook."""

from __future__ import annotations

import re
from collections import defaultdict
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any, Iterable, Sequence

from .config import Config
from .database import DatabaseConnection
from .excel_layout import V2ChartSpec, render_v2_dashboard
from .mapping import QueueMapping, load_queue_mapping
from .metrics import MetricCatalog, evaluate_metric, load_metric_catalog
from .report_packs import (
    archive_superseded_reports,
    publish_report,
    report_current_path,
)
from .reports import COLORS
from .rules import Rulebook, load_rulebook
from .service_profiles import ServiceProfile, load_service_profiles
from .template_reports import DecisionWorkbook, KpiCard


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


def _planned_time_off_kind(value: Any) -> str | None:
    """Return the governed FTE-register kind encoded in a timeline state."""

    normalized = str(value or "").strip().upper()
    for kind in ("PTO", "AWAY"):
        if normalized == kind or normalized.startswith(f"{kind}:"):
            return kind
    return None


def _ratio(numerator: float | int | None, denominator: float | int | None) -> float | None:
    if numerator is None or denominator is None or float(denominator) == 0:
        return None
    return float(numerator) / float(denominator)


def _rtm_sheet(profile: ServiceProfile) -> str:
    """Keep the four operational tabs short without changing profile config."""

    return {
        "rsa_nl": "RSA NL",
        "rsa_be": "RSA BE",
        "ford_nl": "FORD NL",
        "ford_oem_fr": "OEM",
    }.get(profile.profile_id, profile.flash_sheet.replace("Flash ", "")[:31])


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
    profile: ServiceProfile,
    report_day: date,
    attendance: Sequence[dict[str, Any]],
    pulse: dict[str, Any],
) -> dict[int, dict[str, float | None]]:
    """Return proven no-show HC by scheduled hour.

    Late arrivals, internal gaps and early leavers have attended the day and
    therefore never enter this counter. Missing evidence remains unknown.
    """

    checkpoint = _as_datetime(pulse.get("checkpoint"))
    live = pulse.get("mode") == "LIVE"
    output: dict[int, dict[str, float | None]] = {}
    for hour in range(profile.operating_start_hour, profile.operating_end_hour + 1):
        left = datetime.combine(report_day, datetime.min.time()) + timedelta(hours=hour)
        right = left + timedelta(hours=1)
        if live and checkpoint is not None and left >= checkpoint:
            output[hour] = {"no_show_hc": None}
            continue
        no_show = set()
        for row in attendance:
            if not bool(row.get("no_show")):
                continue
            expected_intervals = row.get("expected_work_intervals")
            if expected_intervals is None:
                start = _as_datetime(row.get("scheduled_start"))
                end = _as_datetime(row.get("scheduled_end"))
                expected_intervals = (
                    [(start, end)]
                    if start is not None and end is not None
                    and str(row.get("assignment_type") or "") != "Planned absence"
                    else []
                )
            if any(start < right and end > left for start, end in expected_intervals):
                no_show.add(str(row.get("agent_id")))
        output[hour] = {"no_show_hc": float(len(no_show))}
    return output


def _attendance_pulse(
    conn: DatabaseConnection,
    profile: ServiceProfile,
    report_day: date,
    stale_minutes: int = 30,
) -> tuple[list[dict[str, Any]], dict[str, Any]]:
    """Build one reconcilable same-day attendance pulse for an RTM LOB.

    The pulse uses the governed attendance agent/day mart.  The full-shift
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
                   is_provisional, evaluation_as_of, planning_overlay,
                   planning_overlay_minutes
            FROM mart.attendance_agent_day
            WHERE business_date=? AND lob IN ({lob_marks})
              AND assignment_type<>'Off'
            ORDER BY scheduled_start, lob, team_leader, agent_name""",
        [report_day, *profile.staffing_lobs],
    )
    headers = [item[0] for item in cursor.description]
    attendance = [dict(zip(headers, row)) for row in cursor.fetchall()]
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

    live = report_day == date.today()
    if live:
        evaluation_points = [
            stamp for stamp in (
                _as_datetime(row.get("evaluation_as_of")) for row in attendance
            ) if stamp is not None and stamp.date() == report_day
        ]
        evaluation_ceiling = max(evaluation_points, default=datetime.now())
        status_points = [
            _as_datetime(segment.get("segment_end"))
            for segments in segments_by_agent.values() for segment in segments
            if str(segment.get("observed_source") or "").strip().upper()
            == "AGENT_STATUS"
        ]
        status_points = [
            stamp for stamp in status_points
            if stamp is not None and stamp <= evaluation_ceiling
        ]
        lilo_points = [
            _as_datetime(segment.get("segment_end"))
            for segments in segments_by_agent.values() for segment in segments
            if str(segment.get("observed_source") or "").strip().upper() == "LILO"
        ]
        lilo_points = [
            stamp for stamp in lilo_points
            if stamp is not None and stamp <= evaluation_ceiling
        ]
        # Agent Status is the live attendance clock. Refresh time can be later
        # than the extract cutoff and would otherwise turn valid agents into
        # UNKNOWN. LILO is only a fallback when no status interval exists.
        checkpoint = (
            max(status_points) if status_points
            else max(lilo_points) if lilo_points
            else evaluation_ceiling
        )
    else:
        checkpoint = datetime.combine(report_day, datetime.min.time()) + timedelta(
            hours=23, minutes=59, seconds=59,
        )

    pulse: list[dict[str, Any]] = []
    for row in attendance:
        start = _as_datetime(row.get("scheduled_start"))
        end = _as_datetime(row.get("scheduled_end"))
        state_checkpoint = checkpoint - timedelta(seconds=1)
        agent_segments = segments_by_agent.get(str(row["agent_id"]), [])
        current_segment = next((
            item for item in agent_segments
            if (_as_datetime(item.get("segment_start")) or state_checkpoint) <= state_checkpoint
            < (_as_datetime(item.get("segment_end")) or state_checkpoint)
        ), None)
        full_time_off = str(row.get("assignment_type") or "") == "Planned absence"
        planned_segments = [
            item for item in agent_segments
            if _planned_time_off_kind(item.get("planned_state")) is not None
        ]
        current_time_off_kind = (
            _planned_time_off_kind(current_segment.get("planned_state"))
            if current_segment is not None else None
        )
        result = str(row.get("attendance_result") or "").upper()
        full_time_off_kind = (
            "PTO" if result == "PTO" else
            "AWAY" if result == "AWAY" else
            _planned_time_off_kind(row.get("planning_overlay"))
        )
        time_off_kind = current_time_off_kind or full_time_off_kind
        time_off_today = bool(
            full_time_off_kind or planned_segments
            or float(row.get("planning_overlay_minutes") or 0) > 0
        )
        time_off_now = bool(full_time_off or current_time_off_kind)
        expected_work_intervals = [
            (segment_start, segment_end)
            for item in agent_segments
            if _planned_time_off_kind(item.get("planned_state")) is None
            and (segment_start := _as_datetime(item.get("segment_start"))) is not None
            and (segment_end := _as_datetime(item.get("segment_end"))) is not None
            and segment_end > segment_start
        ]
        if not agent_segments and not full_time_off and start and end:
            expected_work_intervals = [(start, end)]
        expected_work_due = any(
            segment_start < checkpoint
            for segment_start, _segment_end in expected_work_intervals
        )
        scheduled_now = bool(
            start and end and start <= checkpoint < end and not time_off_now
        )
        current_source = str(
            current_segment.get("observed_source") if current_segment else ""
        ).strip().upper()
        recent_segment = (
            current_segment if current_source not in {"", "NONE"} else None
        )
        if recent_segment is None and live:
            freshness_floor = checkpoint - timedelta(minutes=max(0, stale_minutes))
            recent_segment = max((
                item for item in agent_segments
                if str(item.get("observed_source") or "").strip().upper()
                not in {"", "NONE"}
                and freshness_floor <= (
                    _as_datetime(item.get("segment_end")) or freshness_floor
                ) <= checkpoint
            ), key=lambda item: (
                _as_datetime(item.get("segment_end")) or datetime.min
            ), default=None)
        evidence_segment = recent_segment
        segment_source = str(
            evidence_segment.get("observed_source") if evidence_segment else ""
        ).strip().upper()
        row_evidence = str(row.get("actual_evidence") or "NONE").strip().upper()
        explicit_segment = segment_source not in {"", "NONE"}
        first_seen = _as_datetime(row.get("actual_first_seen"))

        late_today = float(row.get("uncoded_late_minutes") or 0) > 0
        early_leave = "EARLY LEAVE" in result
        presence_categories = {
            "PRODUCTIVE", "AUXILIARY", "BREAK", "LUNCH", "LILO_PRESENT",
        }
        has_presence = bool(first_seen and first_seen <= checkpoint) or any(
            str(item.get("actual_category") or "").strip().upper()
            in presence_categories
            and (_as_datetime(item.get("segment_start")) or checkpoint) < checkpoint
            for item in agent_segments
        )
        proven_no_show = bool(
            not full_time_off and expected_work_due
            and not has_presence and row_evidence != "NONE"
            and ("NO SHOW" in result or "NOT SEEN" in result)
        )
        offline_now = bool(
            live and scheduled_now and has_presence and explicit_segment
            and evidence_segment is not None
            and bool(evidence_segment.get("is_gap"))
        )
        if full_time_off:
            state = time_off_kind or "PLANNED TIME OFF"
        elif time_off_now and not expected_work_due:
            state = f"{time_off_kind or 'TIME OFF'} — NOT DUE"
        elif proven_no_show:
            state = "NO SHOW" + (
                f" — {time_off_kind} NOW" if time_off_now and time_off_kind else ""
            )
        elif has_presence:
            qualifiers = []
            if late_today:
                qualifiers.append("LATE")
            if early_leave:
                qualifiers.append("EARLY LEAVE")
            elif offline_now:
                qualifiers.append("OFFLINE NOW")
            if time_off_now and time_off_kind:
                working_during_time_off = bool(
                    current_segment is not None
                    and str(current_segment.get("observed_source") or "").upper()
                    == "AGENT_STATUS"
                )
                qualifiers.append(
                    f"WORKING DURING {time_off_kind}"
                    if working_during_time_off else f"{time_off_kind} NOW"
                )
            state = "PRESENT" + (" — " + " + ".join(qualifiers) if qualifiers else "")
        elif live and start and checkpoint < start:
            state = "UPCOMING"
        elif start and checkpoint >= start:
            # No positive presence and no agent-specific no-show proof.
            state = "UNKNOWN — POSSIBLE NO SHOW"
        else:
            state = "UNKNOWN"
        call_action = str(row.get("call_action") or "NONE")
        if full_time_off or (time_off_now and not expected_work_due):
            call_action = "NONE"
        elif proven_no_show:
            call_action = "CALL_NO_SHOW"
        elif offline_now and call_action == "NONE":
            call_action = "CHECK_OFFLINE_NOW"
        elif state.startswith("UNKNOWN"):
            call_action = "CHECK_DATA_POSSIBLE_NO_SHOW"
        reliable_action = (
            call_action in {"CALL_NO_SHOW", "CALL_LATE"}
            and row_evidence != "NONE"
        )
        call_now = bool(reliable_action)
        current_evidence = (
            segment_source if explicit_segment else row_evidence
        )
        pulse.append({
            "flash": profile.label, "checkpoint": checkpoint,
            **row, "scheduled_now": scheduled_now,
            "attendance_state": state, "present": has_presence,
            "no_show": proven_no_show, "absence_now": proven_no_show,
            "offline_now": offline_now, "late_today": late_today,
            "early_leave": early_leave,
            "call_now": call_now, "pulse_action": call_action,
            "current_evidence": current_evidence,
            "time_off_today": time_off_today, "time_off_now": time_off_now,
            "time_off_kind": time_off_kind,
            "expected_work_due": expected_work_due,
            "expected_work_intervals": expected_work_intervals,
        })
    due_rows = [
        row for row in pulse
        if bool(row.get("expected_work_due"))
    ]
    summary = {
        "checkpoint": checkpoint, "mode": "LIVE" if live else "FINAL DAY",
        "agent_rows": len(pulse),
        "due_hc": len(due_rows),
        "scheduled_now": sum(bool(row["scheduled_now"]) for row in pulse),
        "time_off_hc": sum(bool(row["time_off_today"]) for row in pulse),
        "present_hc": sum(bool(row["present"]) for row in due_rows),
        "no_show_hc": sum(bool(row["no_show"]) for row in due_rows),
        "offline_now": sum(bool(row["offline_now"]) for row in due_rows),
        "call_now": sum(bool(row["call_now"]) for row in due_rows),
        "late_today": sum(bool(row["late_today"]) for row in due_rows),
        "early_leave": sum(bool(row["early_leave"]) for row in due_rows),
        "unknown_hc": sum(
            str(row["attendance_state"]).startswith("UNKNOWN")
            for row in due_rows
        ),
    }
    return pulse, summary


def _hourly_model(
    conn: DatabaseConnection,
    profile: ServiceProfile,
    mapping: QueueMapping,
    metrics: MetricCatalog,
    report_day: date,
    all_rows: Sequence[dict[str, Any]],
    attendance: Sequence[dict[str, Any]] = (),
    pulse: dict[str, Any] | None = None,
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
    workforce = _workforce_by_hour(profile, report_day, attendance, pulse or {})
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
        total["no_show_hc"] = (
            pulse.get("no_show_hc") if pulse is not None else None
        )
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
            "font_color": COLORS["muted"], "bg_color": COLORS["canvas"],
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
    ws.merge_range(4, column, 4, column + 2, label.upper(), formats["card_label"])
    value_format = formats.get(f"card_{kind}", formats["card_value"])
    ws.merge_range(5, column, 6, column + 2, value if value is not None else "—", value_format)
    ws.merge_range(7, column, 7, column + 2, note, formats["card_note"])


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
    if any(token in lowered for token in (
        "forecast", "actual", "handled", "variance", "absence hc", "abs hc",
        "no show hc",
    )):
        return book.report.integer
    return book.report.body


def _flash_columns(
    profile: ServiceProfile,
    hourly: Sequence[dict[str, Any]],
) -> tuple[list[str], list[list[Any]], int, int, int]:
    if profile.flash_layout == "oem_split":
        headers = [
            "Hour", "Forecast", "Actual", "Volume Handled", "Handled in SL",
            "Variance", "Ford Volume", "Chery Volume", "Toyota Volume", "TSL OEM", "TSL Ford",
            "TSL Chery", "TSL Toyota", "Routed Rate", "AHT", "No Show HC",
            "Data State",
        ]
        rows = []
        for row in hourly:
            ford = row["groups"].get("Ford") or {}
            chery = row["groups"].get("Chery") or {}
            toyota = row["groups"].get("Toyota") or {}
            rows.append([
                row["hour_label"], row["forecast"], row["offered"],
                row.get("answered"), row.get("answered_within_target"),
                row.get("volume_variance"), ford.get("offered"),
                chery.get("offered"), toyota.get("offered"), row["service_level"],
                ford.get("service_level"), chery.get("service_level"),
                toyota.get("service_level"), row["availability"],
                row["aht_seconds"], row.get("no_show_hc"), row["data_state"],
            ])
        return headers, rows, 1, 2, 9
    headers = [
        "Hour", "Forecast", "Actual", "Volume Handled", "Handled in SL",
        "Variance", "TSL", "Routed Rate", "AHT", "No Show HC", "Data State",
    ]
    rows = [[
        row["hour_label"], row["forecast"], row["offered"], row.get("answered"),
        row.get("answered_within_target"), row.get("volume_variance"),
        row["service_level"], row["availability"], row["aht_seconds"],
        row.get("no_show_hc"), row["data_state"],
    ] for row in hourly]
    return headers, rows, 1, 2, 6


def _flash_cards(
    profile: ServiceProfile,
    total: dict[str, Any] | None,
    groups: dict[str, dict[str, Any] | None],
    pulse: dict[str, Any],
) -> list[tuple[str, Any, str, str]]:
    value = total or {}
    del profile, groups
    forecast = value.get("forecast")
    variance = value.get("volume_variance")
    actual_note = (
        f"Forecast {forecast:,.0f} | variance {variance:+,.0f}"
        if forecast is not None and variance is not None
        else "Forecast or actual unavailable"
    )
    target = value.get("service_target")
    return [
        ("TSL", value.get("service_level"), "percent", f"Target {target:.0%}" if target is not None else "Target unavailable"),
        ("Offered", value.get("offered"), "integer", actual_note),
        ("Volume Variance", variance, "integer", "Actual minus forecast through cutoff"),
        ("No Show HC", pulse.get("no_show_hc"), "integer", "No presence; evidence proven"),
    ]


def _add_flash_sheet(
    book: DecisionWorkbook,
    profile: ServiceProfile,
    report_day: date,
    hourly: Sequence[dict[str, Any]],
    total: dict[str, Any] | None,
    group_totals: dict[str, dict[str, Any] | None],
    attendance: Sequence[dict[str, Any]],
    pulse: dict[str, Any],
    cutoff: int | None,
) -> None:
    sheet_name = _rtm_sheet(profile)
    ws = book.report.workbook.add_worksheet(sheet_name)
    cutoff_text = f"through {cutoff:02d}:59" if cutoff is not None else "no mapped queue entries"
    ready = total is not None and total.get("forecast") is not None
    headers, display_rows, forecast_col, actual_col, sl_col = _flash_columns(profile, hourly)
    action_labels = {
        "CALL_NO_SHOW": "CALL NOW — NO SHOW",
        "CALL_NOT_SEEN_NOW": "CALL NOW — NOT SEEN",
        "CALL_LATE": "CALL / FOLLOW UP — LATE",
        "CHECK_OFFLINE_NOW": "CHECK — OFFLINE NOW",
        "CHECK_DATA_POSSIBLE_NO_SHOW": "CHECK DATA",
        "NONE": "—",
    }
    sorted_attendance = sorted(
        attendance,
        key=lambda row: (
            0 if row.get("no_show") else
            1 if row.get("call_now") else
            2 if row.get("offline_now") else
            3 if str(row.get("attendance_state")).startswith("UNKNOWN") else
            4 if row.get("late_today") else 5,
            str(row.get("team_leader") or ""),
            str(row.get("agent_name") or ""),
        ),
    )
    action_rows = []
    for row in sorted_attendance[:8]:
        start = _as_datetime(row.get("scheduled_start"))
        end = _as_datetime(row.get("scheduled_end"))
        shift = f"{start:%H:%M}–{end:%H:%M}" if start and end else "—"
        action_rows.append((
            f"{row.get('agent_name') or 'Agent'} [{row.get('agent_id') or '—'}]",
            shift, row.get("attendance_state") or "UNKNOWN",
            int(bool(row.get("no_show"))),
            int(str(row.get("attendance_state") or "").startswith("UNKNOWN")),
            int(bool(row.get("late_today"))),
            action_labels.get(
                str(row.get("pulse_action") or "NONE"),
                str(row.get("pulse_action") or "—"),
            ),
        ))
    hourly_categories = tuple(str(row.get("hour_label") or "") for row in hourly)
    render_v2_dashboard(
        book.report.workbook, ws,
        title=f"{sheet_name.upper()} SERVICE FLASH",
        filters=(
            ("Date", report_day.isoformat()),
            ("Snapshot", cutoff_text.replace("through ", "")),
            ("Queue scope", "Configured"),
            ("Attendance", f"{pulse['checkpoint']:%H:%M}"),
        ),
        kpis=(
            ("TSL", (total or {}).get("service_level"), "percent"),
            ("Offered", (total or {}).get("offered"), "integer"),
            ("Volume variance", (total or {}).get("volume_variance"), "integer"),
            ("No Show HC", pulse.get("no_show_hc"), "integer"),
        ),
        left_chart=V2ChartSpec(
            "HOURLY ACTUAL VS FORECAST", "column", hourly_categories,
            (
                ("Actual", tuple(row.get("offered") for row in hourly), COLORS["teal"]),
                ("Forecast", tuple(row.get("forecast") for row in hourly), COLORS["muted"]),
            ),
        ),
        right_chart=V2ChartSpec(
            "HOURLY SERVICE LEVEL", "line", hourly_categories,
            (
                ("TSL", tuple(row.get("service_level") for row in hourly), COLORS["teal"]),
                ("Target", tuple(row.get("service_target") for row in hourly), COLORS["muted"]),
            ),
            "percent", 0, 1,
        ),
        action_title="Attendance pulse & call list",
        action_headers=("AGENT", "SHIFT", "STATE NOW", "NO SHOW", "POSSIBLE NS", "LATE", "ACTION"),
        action_rows=action_rows,
        action_kinds=("text", "text", "text", "integer", "integer", "integer", "alert"),
        status=f"UPDATED {book.generated:%H:%M}" if ready else "CHECK DATA",
        status_kind="LIVE" if ready else "INCOMPLETE",
        status_note=(
            f"Calls {cutoff_text}; attendance {pulse['mode'].lower()} at "
            f"{pulse['checkpoint']:%H:%M}."
        ),
    )
    formats = _formats(book)
    table_row = 34
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
        table_name = "tblRtm" + re.sub(r"[^A-Za-z0-9]", "", sheet_name)
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
            total_values.get("forecast"), total_values.get("offered"),
            total_values.get("answered"), total_values.get("answered_within_target"),
            total_values.get("volume_variance"), ford.get("offered"),
            chery.get("offered"), toyota.get("offered"),
            total_values.get("service_level"), ford.get("service_level"),
            chery.get("service_level"), toyota.get("service_level"),
            total_values.get("availability"), total_values.get("aht_seconds"),
            total_values.get("no_show_hc"), "READY" if total else "INCOMPLETE",
        ]
    else:
        values = [
            total_values.get("forecast"), total_values.get("offered"),
            total_values.get("answered"), total_values.get("answered_within_target"),
            total_values.get("volume_variance"), total_values.get("service_level"),
            total_values.get("availability"), total_values.get("aht_seconds"),
            total_values.get("no_show_hc"),
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
    target = total_values.get("service_target")
    if target is not None and display_rows:
        ws.conditional_format(table_row + 1, sl_col, data_last_row, sl_col, {
            "type": "cell", "criteria": "<", "value": target,
            "format": book.report.error,
        })

    attendance_section = total_row + 3
    ws.merge_range(
        attendance_section, 0, attendance_section, 14,
        "ATTENDANCE  /  SAME-DAY OPERATIONAL LIST", book.report.section,
    )
    attendance_summary_headers = [
        "Checkpoint", "Due HC", "PTO / Away HC",
        "Present HC", "No Show HC", "Late HC", "Early Leave HC",
        "Offline Now", "Possible No Show HC", "Call Now",
    ]
    attendance_summary_values = [
        pulse.get("checkpoint"), pulse.get("due_hc"), pulse.get("time_off_hc"),
        pulse.get("present_hc"), pulse.get("no_show_hc"),
        pulse.get("late_today"), pulse.get("early_leave"),
        pulse.get("offline_now"), pulse.get("unknown_hc"),
        pulse.get("call_now"),
    ]
    for column, header in enumerate(attendance_summary_headers):
        ws.write(attendance_section + 2, column, header, book.report.header)
        value = attendance_summary_values[column]
        fmt = book.report.datetime if isinstance(value, datetime) else book.report.integer
        if value is None:
            ws.write_blank(attendance_section + 3, column, None, fmt)
        else:
            ws.write(attendance_section + 3, column, value, fmt)

    attendance_headers = [
        "Agent", "Agent ID", "Team Leader", "Shift", "First Seen", "State",
        "Time Off", "Late Min", "Action", "Evidence",
    ]
    attendance_rows = []
    for row in sorted_attendance:
        start = _as_datetime(row.get("scheduled_start"))
        end = _as_datetime(row.get("scheduled_end"))
        shift = (
            f"{start:%H:%M}–{end:%H:%M}" if start is not None and end is not None
            else "—"
        )
        attendance_rows.append((
            row.get("agent_name"), row.get("agent_id"), row.get("team_leader"),
            shift, row.get("actual_first_seen"), row.get("attendance_state"),
            row.get("planning_overlay"), row.get("uncoded_late_minutes"),
            action_labels.get(str(row.get("pulse_action") or "NONE"), str(row.get("pulse_action") or "—")),
            row.get("current_evidence"),
        ))
    attendance_header_row = attendance_section + 6
    for column, header in enumerate(attendance_headers):
        ws.write(attendance_header_row, column, header, book.report.header)
    for row_index, values in enumerate(attendance_rows, attendance_header_row + 1):
        for column, value in enumerate(values):
            fmt = (
                book.report.datetime if isinstance(value, datetime)
                else book.report.integer if column == 7
                else book.report.body
            )
            if value is None:
                ws.write_blank(row_index, column, None, fmt)
            else:
                ws.write(row_index, column, value, fmt)
    if attendance_rows:
        ws.add_table(
            attendance_header_row, 0,
            attendance_header_row + len(attendance_rows), len(attendance_headers) - 1,
            {
                "name": "tblAttendance" + re.sub(r"[^A-Za-z0-9]", "", sheet_name),
                "style": "Table Style Light 9",
                "columns": [
                    {"header": header, "header_format": book.report.header}
                    for header in attendance_headers
                ],
            },
        )
        state_col = attendance_headers.index("State")
        action_col = attendance_headers.index("Action")
        ws.conditional_format(
            attendance_header_row + 1, state_col,
            attendance_header_row + len(attendance_rows), state_col,
            {"type": "text", "criteria": "containing", "value": "NO SHOW", "format": book.report.error},
        )
        ws.conditional_format(
            attendance_header_row + 1, state_col,
            attendance_header_row + len(attendance_rows), state_col,
            {"type": "text", "criteria": "containing", "value": "UNKNOWN", "format": book.report.error},
        )
        ws.conditional_format(
            attendance_header_row + 1, action_col,
            attendance_header_row + len(attendance_rows), action_col,
            {"type": "text", "criteria": "containing", "value": "CALL", "format": book.report.error},
        )
    else:
        ws.write(
            attendance_header_row + 1, 0,
            "No scheduled working agents were found for this LOB/date.",
            book.report.note,
        )
    ws.set_landscape()
    ws.fit_to_pages(1, 0)


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
    ready_count = sum(
        1 for _profile, total, _cutoff, _pulse in summaries
        if total is not None and total.get("forecast") is not None
    )
    on_target = sum(
        1 for _profile, total, _cutoff, _pulse in summaries
        if total is not None
        and total.get("service_level") is not None
        and total.get("service_target") is not None
        and total["service_level"] >= total["service_target"]
    )
    variances = [
        total.get("volume_variance")
        for _profile, total, _cutoff, _pulse in summaries
        if total is not None and total.get("volume_variance") is not None
    ]
    no_show_hc = sum(int(pulse.get("no_show_hc") or 0) for *_rest, pulse in summaries)
    call_now = sum(int(pulse.get("call_now") or 0) for *_rest, pulse in summaries)
    workbook_ready = bool(summaries) and ready_count == len(summaries)
    latest_cutoffs = [cutoff for _profile, _total, cutoff, _pulse in summaries if cutoff is not None]
    checkpoint_values = [
        pulse.get("checkpoint") for _profile, _total, _cutoff, pulse in summaries
        if pulse.get("checkpoint") is not None
    ]
    lob_labels = tuple(profile.label for profile, *_rest in summaries)
    render_v2_dashboard(
        book.report.workbook, ws,
        title="RTM DAILY CONTROL",
        filters=(
            ("Date", report_day.isoformat()),
            ("Snapshot", f"Through {max(latest_cutoffs):02d}:59" if latest_cutoffs else "No calls"),
            ("LOB", f"All {len(summaries)}"),
            ("Attendance", max(checkpoint_values).strftime("%H:%M") if checkpoint_values else "No evidence"),
        ),
        kpis=(
            ("LOBs on target", f"{on_target} / {len(summaries)}", "text"),
            ("Demand variance", sum(variances) if variances else None, "integer"),
            ("No Show HC", no_show_hc, "integer"),
            ("Call now", call_now, "integer"),
        ),
        left_chart=V2ChartSpec(
            "SERVICE LEVEL BY LOB", "bar", lob_labels,
            (
                ("TSL", tuple((total or {}).get("service_level") for _p, total, _c, _a in summaries), COLORS["teal"]),
                ("Target", tuple((total or {}).get("service_target") for _p, total, _c, _a in summaries), COLORS["muted"]),
            ),
            "percent", 0, 1,
        ),
        right_chart=V2ChartSpec(
            "ACTUAL VS FORECAST", "column", lob_labels,
            (
                ("Actual", tuple((total or {}).get("offered") for _p, total, _c, _a in summaries), COLORS["teal"]),
                ("Forecast", tuple((total or {}).get("forecast") for _p, total, _c, _a in summaries), COLORS["muted"]),
            ),
        ),
        action_title="LOB service & attendance actions",
        action_headers=("LOB", "TSL", "HANDLED", "NO SHOW", "POSSIBLE NS", "LATE", "CALL NOW"),
        action_rows=tuple((
            profile.label, (total or {}).get("service_level"),
            (total or {}).get("answered"), pulse.get("no_show_hc"),
            pulse.get("unknown_hc"), pulse.get("late_today"), pulse.get("call_now"),
        ) for profile, total, _cutoff, pulse in summaries),
        action_kinds=("text", "percent", "integer", "integer", "integer", "integer", "alert"),
        status=f"UPDATED {book.generated:%H:%M}" if workbook_ready else "CHECK DATA",
        status_kind="LIVE" if workbook_ready else "INCOMPLETE",
        status_note=f"{ready_count}/{len(summaries)} operational LOBs have call and forecast data.",
    )
    headers = [
        "LOB", "Last Call Hour", "TSL", "Target", "Actual", "Volume Handled",
        "Handled in SL", "Forecast", "Variance", "Routed Rate", "No Show HC",
        "PTO / Away HC", "Offline Now", "Possible No Show HC", "Call Now", "Status",
    ]
    table_row = 34
    for col, header in enumerate(headers):
        ws.write(table_row, col, header, book.report.header)
    for offset, (profile, total, cutoff, pulse) in enumerate(summaries, table_row + 1):
        value = total or {}
        status = (
            "CALL DATA MISSING" if total is None
            else "FORECAST MISSING" if value.get("forecast") is None
            else "READY"
        )
        row = [
            profile.label, f"{cutoff:02d}:59" if cutoff is not None else None,
            value.get("service_level"), value.get("service_target"),
            value.get("offered"), value.get("answered"),
            value.get("answered_within_target"), value.get("forecast"),
            value.get("volume_variance"), value.get("availability"),
            pulse.get("no_show_hc"), pulse.get("time_off_hc"),
            pulse.get("offline_now"), pulse.get("unknown_hc"),
            pulse.get("call_now"), status,
        ]
        for col, item in enumerate(row):
            fmt = _table_value_format(book, headers[col])
            if item is None:
                ws.write_blank(offset, col, None, fmt)
            else:
                ws.write(offset, col, item, fmt)
        ws.write_url(offset, 0, f"internal:'{_rtm_sheet(profile)}'!A1", book.report.body, profile.label)
    if summaries:
        ws.add_table(table_row, 0, table_row + len(summaries), len(headers) - 1, {
            "name": "tblFlashControl", "style": "Table Style Light 9",
            "columns": [{"header": header, "header_format": book.report.header} for header in headers],
        })
        status_column = headers.index("Status")
        ws.conditional_format(table_row + 1, status_column, table_row + len(summaries), status_column, {
            "type": "text", "criteria": "containing", "value": "MISSING",
            "format": book.report.error,
        })
    notes = [
        "Open a LOB name for its full-day service curve and matching attendance call list.",
        f"TSL follows the approved Storm C/(A+B-D) method at {rulebook.target_seconds}s; each LOB keeps its configured target.",
        "No Show requires explicit evidence. PTO/Away is excluded; Unknown remains a data check, never an automatic call.",
    ]
    notes_row = table_row + len(summaries) + 3
    ws.merge_range(notes_row, 0, notes_row, 13, "OPERATING NOTES", book.report.section)
    for offset, note in enumerate(notes):
        note_row = notes_row + 1 + offset * 2
        ws.merge_range(note_row, 0, note_row + 1, 13, note, book.report.note)
        ws.set_row(note_row, 24)
        ws.set_row(note_row + 1, 24)
    ws.set_landscape()
    ws.fit_to_pages(1, 1)


def _issues_and_drivers_rows(
    profiles: Sequence[ServiceProfile],
    source_by_profile: dict[str, Sequence[dict[str, Any]]],
    hourly_by_profile: dict[str, Sequence[dict[str, Any]]],
    metrics: MetricCatalog,
    report_day: date,
    cutoff_by_profile: dict[str, int | None],
    pulse_summary_by_profile: dict[str, dict[str, Any]],
) -> tuple[list[str], list[tuple[Any, ...]]]:
    """Return only actionable source issues and below-target queues with volume."""

    headers = [
        "item_type", "lob", "hour_or_queue", "volume", "tsl", "target",
        "gap_to_target", "late_routed", "lost_contacts",
        "issue_or_driver", "action",
    ]
    output: list[tuple[Any, ...]] = []
    for profile in profiles:
        cutoff = cutoff_by_profile.get(profile.profile_id)
        for row in hourly_by_profile.get(profile.profile_id, ()):
            if row.get("data_state") == "FORECAST MISSING":
                output.append((
                    "DATA ISSUE", profile.label, row.get("hour_label"),
                    row.get("offered"), row.get("service_level"),
                    row.get("service_target"), None, None, None,
                    "FORECAST MISSING", "Load or map the forecast extract",
                ))
            elif (
                row.get("data_state") == "NO MAPPED CALLS"
                and cutoff is not None and int(row.get("hour") or 0) <= cutoff
            ):
                output.append((
                    "DATA ISSUE", profile.label, row.get("hour_label"),
                    None, None, row.get("service_target"), None, None, None,
                    "NO MAPPED CALLS", "Confirm zero demand or review source mapping",
                ))

        pulse = pulse_summary_by_profile.get(profile.profile_id, {})
        if not pulse.get("agent_rows"):
            output.append((
                "DATA ISSUE", profile.label, f"{pulse.get('checkpoint'):%H:%M}",
                None, None, None, None, None, None,
                "NO SCHEDULED AGENT ROWS", "Check schedule and active FTE coverage",
            ))
        elif pulse.get("unknown_hc"):
            output.append((
                "DATA ISSUE", profile.label, f"{pulse.get('checkpoint'):%H:%M}",
                pulse.get("unknown_hc"), None, None, None, None, None,
                "POSSIBLE NO SHOW / DATA UNKNOWN",
                "Validate Agent Status evidence before calling",
            ))

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
            if aggregate is None or not aggregate.get("offered"):
                continue
            service_level = aggregate.get("service_level")
            if service_level is not None and (
                target is None or service_level >= target
            ):
                continue
            late_connected = max(
                0.0, aggregate["answered"] - aggregate["answered_within_target"],
            )
            lost_in_denominator = max(
                0.0, aggregate["abandoned"] - aggregate["abandoned_within_target"],
            )
            if service_level is None:
                diagnosis = "NO SLA SAMPLE"
                action = "Review queue counters and source coverage"
            elif late_connected > lost_in_denominator * 1.5:
                diagnosis = "LATE ROUTED CONTACTS"
                action = "Review answer delay and staffing"
            elif lost_in_denominator > late_connected * 1.5:
                diagnosis = "LOST CONTACTS"
                action = "Review abandonment and coverage"
            else:
                diagnosis = "MIXED DELAY AND LOSS"
                action = "Review demand, delay and coverage"
            output.append((
                "QUEUE DRIVER", profile.label, queue, aggregate.get("offered"),
                service_level, target, (
                    service_level - target
                    if service_level is not None and target is not None
                    else None
                ), late_connected, lost_in_denominator, diagnosis, action,
            ))
    output.sort(key=lambda row: (
        0 if row[0] == "DATA ISSUE" else 1,
        str(row[1] or ""),
        float(row[6]) if isinstance(row[6], (int, float)) else 0.0,
        -float(row[3]) if isinstance(row[3], (int, float)) else 0.0,
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
    """Build the daily RTM service and attendance control."""

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
        partial, config, "service", "RTM DAILY CONTROL", start, end, generated,
    )
    hourly_by_profile: dict[str, Sequence[dict[str, Any]]] = {}
    source_by_profile: dict[str, Sequence[dict[str, Any]]] = {}
    pulse_by_profile: dict[str, Sequence[dict[str, Any]]] = {}
    pulse_summary_by_profile: dict[str, dict[str, Any]] = {}
    summaries: list[
        tuple[ServiceProfile, dict[str, Any] | None, int | None, dict[str, Any]]
    ] = []
    group_totals_by_profile: dict[str, dict[str, dict[str, Any] | None]] = {}
    try:
        for profile in profiles:
            source_rows = _profile_rows(conn, profile, start, end)
            source_by_profile[profile.profile_id] = source_rows
            pulse_rows, pulse_summary = _attendance_pulse(
                conn, profile, end, config.rules.rta_stale_minutes,
            )
            hourly, total, group_totals, cutoff = _hourly_model(
                conn, profile, mapping, metrics, end, source_rows,
                pulse_rows, pulse_summary,
            )
            hourly_by_profile[profile.profile_id] = hourly
            pulse_by_profile[profile.profile_id] = pulse_rows
            pulse_summary_by_profile[profile.profile_id] = pulse_summary
            group_totals_by_profile[profile.profile_id] = group_totals
            summaries.append((profile, total, cutoff, pulse_summary))
        _add_control_sheet(book, summaries, end, rulebook)
        for profile, total, cutoff, pulse_summary in summaries:
            _add_flash_sheet(
                book, profile, end, hourly_by_profile[profile.profile_id], total,
                group_totals_by_profile[profile.profile_id],
                pulse_by_profile[profile.profile_id], pulse_summary, cutoff,
            )
        cutoff_by_profile = {
            profile.profile_id: cutoff
            for profile, _total, cutoff, _pulse in summaries
        }
        issue_headers, issue_rows = _issues_and_drivers_rows(
            profiles, source_by_profile, hourly_by_profile, metrics, end,
            cutoff_by_profile, pulse_summary_by_profile,
        )
        issue_sheet = book.table(
            "ISSUES & DRIVERS", "RTM issues and service drivers",
            "Only actionable missing evidence and below-target queues with real demand. Empty or healthy queues are intentionally omitted.",
            issue_headers, issue_rows,
        )
        if issue_rows:
            issue_column = issue_headers.index("issue_or_driver")
            issue_sheet.conditional_format(
                4, issue_column, 3 + len(issue_rows), issue_column,
                {"type": "no_errors", "format": book.report.error},
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
            ("Volume Variance", "Total entered minus forecast through the latest actual hour", "Absolute demand gap", "Positive means demand is above forecast"),
            ("Routed Rate", "Total routed / total entered", "Service availability", "Matches the Storm screenshot; not agent availability or adherence"),
            ("TSL", "Connected within target / (lost + connected - lost from 5 seconds to target)", "Storm C/(A+B-D)", "The supplied Storm custom-equation screen is the business authority"),
            ("AHT", "Sum of inbound talk + hold + wrap / routed queue entries", "Workload", "Weighted; never an average of hourly averages"),
            ("No Show HC", "Scheduled agents with no presence at all and explicit agent-specific no-show evidence", "Confirmed RTM callout population", "Late, early leave and offline-after-presence remain present"),
            ("Due HC", "Working shifts started by the attendance checkpoint", "Attendance reconciliation", "Due HC = Present HC + No Show HC + Unknown HC"),
            ("Offline Now", "Agents with presence earlier in the day whose latest reliable state is a current gap", "Operational follow-up", "It is not No Show HC; the exact gap stays in Attendance Review"),
            ("Unknown HC", "Shift started but neither presence nor agent-specific no-show proof exists", "Possible no-show or data problem", "Validate Agent Status before calling; never add it to No Show HC"),
            ("LOB attendance list", "Scheduled working agents from the governed attendance mart plus checkpoint timeline state", "No-show, late, offline and data-check list", "It is operational context, not a final absence decision"),
            ("Issues & Drivers", "Missing source evidence plus below-target configured queues with real demand", "One operational exception list", "Healthy and empty queues are omitted"),
        ])
        book.report.workbook.get_worksheet_by_name("DEFINITIONS").hide()
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
            ("Report", "service", "RTM Daily Control"),
            ("Report day", end, "Visible RTM sheets use the selected end date"),
            ("Selected data period", f"{start} to {end}", "Model boundary"),
            ("Clean Call-by-Call legs", clean_calls, "After stable call-leg deduplication"),
            ("Unique clean interactions", unique_interactions, "Before queue mapping"),
            ("Mapped RTM offered", mapped_offered, "Inbound queue-entry count"),
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
        if target == report_current_path(config, "service"):
            archive_superseded_reports(
                config,
                (
                    "Attendance Callout.xlsx",
                    "Service Flashes.xlsx",
                    "OEM Flash.xlsx",
                ),
                generated,
            )
    except Exception:
        try:
            book.report.close()
        except Exception:
            pass
        partial.unlink(missing_ok=True)
        raise
    return target
