"""Human-readable full-shift evidence view used before Verint correction."""

from __future__ import annotations

from collections import defaultdict
from datetime import date, datetime
from typing import Any

from .reports import COLORS, ExcelReport


def _as_datetime(value: Any) -> datetime:
    if isinstance(value, datetime):
        return value
    return datetime.fromisoformat(str(value))


def _state(segment: dict[str, Any]) -> str:
    """Collapse technical timeline states into five operational labels."""

    mismatch = str(segment.get("mismatch_type") or "").upper()
    category = str(segment.get("actual_category") or "").upper()
    planned = str(segment.get("planned_state") or "").upper()
    if "FUTURE" in mismatch:
        return "Future"
    if any(value in planned for value in ("PTO", "AWAY", "PLANNED ABSENCE", "VACATION")):
        return "PTO / Away"
    if bool(segment.get("is_gap")):
        return "Gap"
    if category == "LUNCH":
        return "Lunch"
    if category == "BREAK":
        return "Break"
    if category in {"PRODUCTIVE", "AUXILIARY", "LILO_PRESENT"}:
        return "Logged"
    if str(segment.get("observed_source") or "").upper() == "LILO":
        return "Logged"
    return "Unknown"


def add_shift_view(
    report: ExcelReport,
    segments: list[dict[str, Any]],
    period_start: date,
    period_end: date,
) -> None:
    """Show a compact 15-minute visual; exact times stay in VERINT_INJECTION."""

    ws = report.workbook.add_worksheet("SHIFT_VIEW")
    ws.set_tab_color(COLORS["purple"])
    ws.hide_gridlines(2)
    ws.set_zoom(75)
    ws.merge_range("A1:X1", "ATTENDANCE REVIEW  /  SHIFT VIEW", report.title)
    ws.merge_range(
        "A2:X2",
        f"Completed dates {period_start:%Y-%m-%d} to {period_end:%Y-%m-%d}  |  visual check before Verint entry",
        report.subtitle,
    )
    legend = [
        ("Logged", COLORS["green"], COLORS["green_light"]),
        ("Break", COLORS["amber"], COLORS["amber_light"]),
        ("Lunch", COLORS["gold"], "#FFF7DD"),
        ("Gap", COLORS["red"], COLORS["red_light"]),
        ("PTO / Away", COLORS["blue"], COLORS["blue_light"]),
        ("Unknown", COLORS["muted"], COLORS["future_light"]),
    ]
    formats: dict[str, Any] = {}
    for index, (label, font, fill) in enumerate(legend):
        formats[label] = report.workbook.add_format({
            "font_name": "Aptos", "font_size": 8, "bold": True,
            "font_color": font, "bg_color": fill, "align": "center",
            "valign": "vcenter", "border": 1, "border_color": COLORS["white"],
        })
        column = index * 3
        ws.merge_range(3, column, 3, column + 1, label, formats[label])

    if not segments:
        ws.merge_range(
            "A7:X7", "No correction gaps exist in the selected completed dates.",
            report.note,
        )
        ws.set_column("A:X", 10)
        return

    def as_date(value: Any) -> date:
        if isinstance(value, datetime):
            return value.date()
        if isinstance(value, date):
            return value
        return date.fromisoformat(str(value)[:10])

    def minute_offset(value: Any, business_day: date) -> float:
        stamp = _as_datetime(value)
        midnight = datetime.combine(business_day, datetime.min.time())
        return (stamp - midnight).total_seconds() / 60.0

    by_agent_day: dict[tuple[date, str, str], list[dict[str, Any]]] = defaultdict(list)
    for segment in segments:
        business_day = as_date(segment["business_date"])
        key = (
            business_day, str(segment["agent_id"]),
            str(segment.get("agent_name") or ""),
        )
        by_agent_day[key].append(segment)

    first_minute = min(
        minute_offset(item["scheduled_start"], as_date(item["business_date"]))
        for item in segments
    )
    last_minute = max(
        minute_offset(item["scheduled_end"], as_date(item["business_date"]))
        for item in segments
    )
    first_slot = int(first_minute // 15) * 15
    last_slot = int((last_minute + 14.9999) // 15) * 15
    slots = list(range(first_slot, min(last_slot, first_slot + 30 * 60), 15))

    metadata = [
        "Date", "Agent ID", "Agent", "Team Leader", "LOB", "Language",
        "Scheduled Start", "Scheduled End", "Gap Minutes", "Evidence",
    ]
    header_row = 5
    for column, value in enumerate(metadata):
        ws.write(header_row, column, value, report.header)
    rotated = report.workbook.add_format({
        "font_name": "Aptos", "font_size": 8, "bold": True,
        "font_color": COLORS["white"], "bg_color": COLORS["teal"],
        "rotation": 90, "align": "center", "valign": "vcenter",
    })
    for index, slot in enumerate(slots, len(metadata)):
        days_after = slot // 1440
        minute_of_day = slot % 1440
        label = f"{minute_of_day // 60:02d}:{minute_of_day % 60:02d}"
        if days_after:
            label += f"+{days_after}"
        ws.write(header_row, index, label, rotated)
    ws.set_row(header_row, 58)

    state_priority = {
        "Gap": 6, "PTO / Away": 5, "Lunch": 4, "Break": 3,
        "Logged": 2, "Unknown": 1, "Future": 0,
    }
    for row_index, ((_business_day, _agent_id, _agent_name), agent_segments) in enumerate(
        sorted(by_agent_day.items()), header_row + 1,
    ):
        business_day = as_date(agent_segments[0]["business_date"])
        agent_segments.sort(key=lambda item: _as_datetime(item["segment_start"]))
        base = agent_segments[0]
        gap_minutes = sum(
            int(item.get("segment_minutes") or 0)
            for item in agent_segments if item.get("is_gap")
        )
        evidence = "+".join(sorted({
            str(item.get("observed_source") or "") for item in agent_segments
            if item.get("observed_source")
        }))
        values = [
            business_day, base["agent_id"], base.get("agent_name"),
            base.get("team_leader"), base.get("lob"), base.get("language"),
            base.get("scheduled_start"), base.get("scheduled_end"), gap_minutes,
            evidence,
        ]
        for column, value in enumerate(values):
            fmt = (
                report.datetime if isinstance(value, datetime)
                else report.date if isinstance(value, date)
                else report.integer if column == 8 else report.body
            )
            ws.write(row_index, column, value, fmt)
        scheduled_start = minute_offset(base["scheduled_start"], business_day)
        scheduled_end = minute_offset(base["scheduled_end"], business_day)
        for column, slot in enumerate(slots, len(metadata)):
            right = slot + 15
            if right <= scheduled_start or slot >= scheduled_end:
                continue
            best_state = "Unknown"
            best_rank = (state_priority[best_state], 0.0)
            for segment in agent_segments:
                overlap = max(0.0, min(
                    right, minute_offset(segment["segment_end"], business_day),
                ) - max(
                    slot, minute_offset(segment["segment_start"], business_day),
                ))
                state = _state(segment)
                rank = (state_priority[state], overlap)
                if overlap > 0 and rank > best_rank:
                    best_state, best_rank = state, rank
            if best_state != "Future":
                ws.write(row_index, column, best_state, formats[best_state])
        ws.set_row(row_index, 19)

    ws.set_column(0, 0, 12)
    ws.set_column(1, 1, 13)
    ws.set_column(2, 5, 20)
    ws.set_column(6, 7, 18)
    ws.set_column(8, 8, 12)
    ws.set_column(9, 9, 18)
    if slots:
        ws.set_column(len(metadata), len(metadata) + len(slots) - 1, 7.5)
    ws.freeze_panes(header_row + 1, len(metadata))
    ws.autofilter(
        header_row, 0, header_row + len(by_agent_day), len(metadata) - 1,
    )
