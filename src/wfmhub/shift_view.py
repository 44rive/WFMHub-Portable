"""Human-readable full-shift evidence view for attendance decisions."""

from __future__ import annotations

from collections import defaultdict
from datetime import date, datetime, timedelta
from typing import Any

from .reports import COLORS, ExcelReport, _display_header


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
    """Show a compact 15-minute visual; exact times stay in DECISIONS."""

    ws = report.workbook.add_worksheet("SHIFT_VIEW")
    ws.set_tab_color(COLORS["purple"])
    ws.hide_gridlines(2)
    ws.set_zoom(75)
    ws.merge_range("A1:X1", "ATTENDANCE REVIEW  /  SHIFT VIEW", report.title)
    ws.merge_range(
        "A2:X2",
        f"Completed dates {period_start:%Y-%m-%d} to {period_end:%Y-%m-%d}  |  visual evidence for human review",
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


def add_review_board(
    report: ExcelReport,
    decision_headers: list[str],
    decision_rows: list[tuple[Any, ...]],
    segments: list[dict[str, Any]],
    period_start: date,
    period_end: date,
    category_choices: list[str],
):
    """Place the scheduled band directly above actual evidence for each gap."""

    ws = report.workbook.add_worksheet("REVIEW BOARD")
    ws.set_tab_color(COLORS["purple"])
    ws.hide_gridlines(2)
    ws.set_zoom(70)

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

    by_agent_day: dict[tuple[date, str], list[dict[str, Any]]] = defaultdict(list)
    for segment in segments:
        by_agent_day[(as_date(segment["business_date"]), str(segment["agent_id"]))].append(segment)

    if segments:
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
    else:
        slots = []

    display_headers = [_display_header(header) for header in decision_headers]
    band_column = len(display_headers)
    timeline_start = band_column + 1
    last_column = max(timeline_start + len(slots) - 1, len(display_headers) - 1)
    ws.merge_range(0, 0, 0, last_column, "ATTENDANCE REVIEW  /  VISUAL DECISION BOARD", report.title)
    ws.merge_range(
        1, 0, 1, last_column,
        f"Completed dates {period_start:%Y-%m-%d} to {period_end:%Y-%m-%d}  |  exact Agent Status/LILO evidence  |  decisions return to WFM Hub",
        report.subtitle,
    )
    ws.set_row(0, 34)
    ws.set_row(1, 21)

    legend = [
        ("Scheduled", "Scheduled work", COLORS["teal"], COLORS["teal_light"]),
        ("Logged", "Logged", COLORS["green"], COLORS["green_light"]),
        ("Break", "Break", COLORS["amber"], COLORS["amber_light"]),
        ("Lunch", "Lunch", COLORS["gold"], "#FFF7DD"),
        ("Gap", "Gap: this row", COLORS["white"], COLORS["red"]),
        ("Other gap", "Gap: other row", COLORS["red"], COLORS["red_light"]),
        ("Tolerance", "Within tolerance", COLORS["muted"], COLORS["future_light"]),
        ("PTO / Away", "PTO / Away", COLORS["blue"], COLORS["blue_light"]),
        ("Unknown", "Unknown", COLORS["muted"], COLORS["future_light"]),
    ]
    state_formats: dict[str, Any] = {}
    ws.merge_range(
        2, 0, 2, min(8, last_column),
        "SCHEDULE ABOVE ACTUAL  |  edit blue cells on ACTUAL rows only, save, then import",
        report.note,
    )
    legend_column = min(9, last_column + 1)
    for key, label, font, fill in legend:
        state_formats[key] = report.workbook.add_format({
            "font_name": "Aptos", "font_size": 8, "bold": True,
            "font_color": font, "bg_color": fill, "align": "center",
            "valign": "vcenter", "border": 1, "border_color": COLORS["white"],
        })
        if legend_column <= last_column:
            ws.write(2, legend_column, label, state_formats[key])
            legend_column += 1

    header_row = 3
    for column, header in enumerate(display_headers):
        ws.write(header_row, column, header, report.header)
    ws.write(header_row, band_column, "Band", report.header)
    rotated = report.workbook.add_format({
        "font_name": "Aptos", "font_size": 8, "bold": True,
        "font_color": COLORS["white"], "bg_color": COLORS["teal"],
        "rotation": 90, "align": "center", "valign": "vcenter",
    })
    for index, slot in enumerate(slots, timeline_start):
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
    date_index = display_headers.index("Date")
    agent_index = display_headers.index("Agent ID")
    exact_start_index = display_headers.index("Exact Start")
    exact_end_index = display_headers.index("Exact End")
    review_gaps: dict[tuple[date, str], list[tuple[datetime, datetime]]] = defaultdict(list)
    for values in decision_rows:
        raw_start, raw_end = values[exact_start_index], values[exact_end_index]
        if raw_start is None or raw_end is None:
            continue
        review_gaps[(as_date(values[date_index]), str(values[agent_index]))].append((
            _as_datetime(raw_start), _as_datetime(raw_end),
        ))
    editable = {
        "Decision Category", "Decision Status", "Reviewed By", "Comment",
        "Reviewed Date",
    }
    gap_id_index = display_headers.index("Gap ID")
    schedule_body = report.workbook.add_format({
        "font_name": "Aptos", "font_size": 9, "font_color": COLORS["muted"],
        "bg_color": COLORS["blue_light"], "bottom": 1,
        "bottom_color": COLORS["white"],
    })
    schedule_date = report.workbook.add_format({
        "font_name": "Aptos", "font_size": 9, "font_color": COLORS["muted"],
        "bg_color": COLORS["blue_light"], "num_format": "yyyy-mm-dd",
        "bottom": 1, "bottom_color": COLORS["white"],
    })
    schedule_datetime = report.workbook.add_format({
        "font_name": "Aptos", "font_size": 9, "font_color": COLORS["muted"],
        "bg_color": COLORS["blue_light"], "num_format": "yyyy-mm-dd hh:mm",
        "bottom": 1, "bottom_color": COLORS["white"],
    })
    schedule_integer = report.workbook.add_format({
        "font_name": "Aptos", "font_size": 9, "font_color": COLORS["muted"],
        "bg_color": COLORS["blue_light"], "num_format": "0",
        "bottom": 1, "bottom_color": COLORS["white"],
    })
    schedule_band = report.workbook.add_format({
        "font_name": "Aptos", "font_size": 9, "bold": True,
        "font_color": COLORS["blue"], "bg_color": COLORS["blue_light"],
        "align": "center", "bottom": 1, "bottom_color": COLORS["white"],
    })
    actual_band = report.workbook.add_format({
        "font_name": "Aptos", "font_size": 9, "bold": True,
        "font_color": COLORS["green"], "bg_color": COLORS["green_light"],
        "align": "center", "bottom": 1, "bottom_color": COLORS["thin"],
    })
    for item_index, values in enumerate(decision_rows):
        # Keep every schedule/actual pair visually self-contained.  The blank,
        # short row between cases is deliberately not an editable ledger row;
        # the importer ignores it because it has no Gap ID.
        schedule_row = header_row + 1 + item_index * 3
        row_index = schedule_row + 1
        business_day = as_date(values[date_index])
        agent_id = str(values[agent_index])
        current_start = _as_datetime(values[exact_start_index])
        current_end = _as_datetime(values[exact_end_index])
        agent_segments = sorted(
            by_agent_day.get((business_day, agent_id), []),
            key=lambda item: _as_datetime(item["segment_start"]),
        )
        for column, value in enumerate(values):
            if column == gap_id_index:
                ws.write_blank(schedule_row, column, None, schedule_body)
                continue
            fmt = (
                schedule_datetime if isinstance(value, datetime)
                else schedule_date if isinstance(value, date)
                else schedule_integer if isinstance(value, (int, float)) and not isinstance(value, bool)
                else schedule_body
            )
            if value is None:
                ws.write_blank(schedule_row, column, None, fmt)
            else:
                ws.write(schedule_row, column, value, fmt)
        ws.write(schedule_row, band_column, "SCHEDULE", schedule_band)
        for column, value in enumerate(values):
            header = display_headers[column]
            if header in editable:
                fmt = report.editable_date if header == "Reviewed Date" else report.editable
            elif isinstance(value, datetime):
                fmt = report.datetime
            elif isinstance(value, date):
                fmt = report.date
            elif isinstance(value, (int, float)) and not isinstance(value, bool):
                fmt = report.integer
            else:
                fmt = report.body
            if value is None:
                ws.write_blank(row_index, column, None, fmt)
            else:
                ws.write(row_index, column, value, fmt)
        ws.write(row_index, band_column, "ACTUAL", actual_band)
        if agent_segments:
            scheduled_start = minute_offset(agent_segments[0]["scheduled_start"], business_day)
            scheduled_end = minute_offset(agent_segments[0]["scheduled_end"], business_day)
            for column, slot in enumerate(slots, timeline_start):
                right = slot + 15
                if right <= scheduled_start or slot >= scheduled_end:
                    continue
                slot_start = datetime.combine(business_day, datetime.min.time()) + timedelta(minutes=slot)
                slot_end = slot_start + timedelta(minutes=15)
                planned_time_off = any(
                    segment_start < slot_end and segment_end > slot_start
                    and any(token in str(segment.get("planned_state") or "").upper()
                            for token in ("PTO", "AWAY", "PLANNED ABSENCE", "VACATION"))
                    for segment in agent_segments
                    for segment_start, segment_end in [(
                        _as_datetime(segment["segment_start"]),
                        _as_datetime(segment["segment_end"]),
                    )]
                )
                if planned_time_off:
                    ws.write(schedule_row, column, "PTO / Away", state_formats["PTO / Away"])
                else:
                    ws.write(schedule_row, column, "Work", state_formats["Scheduled"])
                if current_start < slot_end and current_end > slot_start:
                    ws.write(row_index, column, "Gap", state_formats["Gap"])
                    continue
                other_gap = any(
                    gap_start < slot_end and gap_end > slot_start
                    and (gap_start, gap_end) != (current_start, current_end)
                    for gap_start, gap_end in review_gaps.get((business_day, agent_id), [])
                )
                if other_gap:
                    ws.write(row_index, column, "Gap", state_formats["Other gap"])
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
                if best_state == "Gap":
                    ws.write(row_index, column, "Tolerance", state_formats["Tolerance"])
                elif best_state != "Future":
                    ws.write(row_index, column, best_state, state_formats[best_state])
        ws.set_row(schedule_row, 15)
        ws.set_row(row_index, 20)
        if item_index < len(decision_rows) - 1:
            ws.set_row(row_index + 1, 6)

    if not decision_rows:
        ws.write(header_row + 1, 0, "No rows for this period.", report.subtitle)
    else:
        category_col = display_headers.index("Decision Category")
        status_col = display_headers.index("Decision Status")
        date_col = display_headers.index("Reviewed Date")
        data_last_row = header_row + len(decision_rows) * 3 - 1
        ws.data_validation(
            header_row + 1, category_col, data_last_row, category_col,
            {"validate": "list", "source": category_choices},
        )
        ws.data_validation(
            header_row + 1, status_col, data_last_row, status_col,
            {"validate": "list", "source": ["Open", "Approved", "Dismissed"]},
        )
        ws.data_validation(
            header_row + 1, date_col, data_last_row, date_col,
            {"validate": "date", "criteria": "between",
             "minimum": date(2020, 1, 1), "maximum": date(2100, 12, 31)},
        )
        ws.conditional_format(
            header_row + 1, status_col, data_last_row, status_col,
            {"type": "text", "criteria": "containing", "value": "Open", "format": report.error},
        )
        ws.autofilter(
            header_row, 0, data_last_row, last_column,
        )

    widths = {
        "Gap ID": 42, "Date": 12, "Agent ID": 13, "Agent": 22,
        "Team Leader": 20, "LOB": 16, "Detected Issue": 22,
        "Exact Start": 19, "Exact End": 19, "Minutes": 10,
        "Suggested Activity": 21, "Decision Category": 21,
        "Decision Status": 16, "Reviewed By": 18, "Comment": 30,
        "Reviewed Date": 15,
    }
    for column, header in enumerate(display_headers):
        ws.set_column(column, column, widths.get(header, 16))
    ws.set_column(band_column, band_column, 11)
    if slots:
        ws.set_column(timeline_start, timeline_start + len(slots) - 1, 7.5)
    ws.freeze_panes(header_row + 1, 7)
    ws.set_landscape()
    ws.fit_to_pages(1, 0)
    return ws
