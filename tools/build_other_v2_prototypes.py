#!/usr/bin/env python3
"""Build dummy V2 workbooks for visual acceptance before production migration."""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
import re
from zipfile import ZIP_DEFLATED, ZipFile

import xlsxwriter
from wfmhub.design import COLORS
from wfmhub.excel_layout import (
    ACTION_FIRST_ROW,
    ACTION_HEADER_ROW,
    ACTION_VISIBLE_ROWS,
    configure_v2_cell_canvas,
    insert_v2_charts,
    make_v2_formats,
    style_v2_chart,
    write_v2_filters,
    write_v2_header,
    write_v2_kpis,
    write_v2_section,
)


ROOT = Path(__file__).resolve().parents[1]
OUTPUT = ROOT / "dist" / "report-prototypes"
PACK = ROOT / "dist" / "WFMHub-V2-Other-Reports-Prototypes.zip"


@dataclass(frozen=True)
class ChartSpec:
    title: str
    kind: str
    categories: tuple[object, ...]
    series: tuple[tuple[str, tuple[float, ...], str], ...]
    value_kind: str = "number"
    minimum: float | None = None
    maximum: float | None = None


@dataclass(frozen=True)
class DashboardSpec:
    sheet: str
    title: str
    maturity: str
    filters: tuple[tuple[str, object], ...]
    kpis: tuple[tuple[str, object, str], ...]
    left_chart: ChartSpec
    right_chart: ChartSpec
    action_title: str
    action_headers: tuple[str, ...]
    action_rows: tuple[tuple[object, ...], ...]
    action_kinds: tuple[str, ...]


def _properties(workbook: xlsxwriter.Workbook, title: str) -> None:
    workbook.set_properties({
        "title": f"{title} — V2 dummy prototype",
        "subject": "Visual acceptance only",
        "author": "Anass ASSRI",
        "company": "WFM",
        "comments": "Dummy data only. Not a production report.",
    })
    workbook.set_custom_property("WFMHub Prototype", "DUMMY DATA")
    workbook.set_calc_mode("auto")


def _chart(
    workbook: xlsxwriter.Workbook,
    ws,
    spec: ChartSpec,
    start_column: int,
):
    category_column = start_column
    first_row = 1
    last_row = first_row + len(spec.categories) - 1
    ws.write(0, category_column, "Category")
    for row, value in enumerate(spec.categories, first_row):
        ws.write(row, category_column, value)
    chart = workbook.add_chart({"type": spec.kind})
    for offset, (name, values, color) in enumerate(spec.series, 1):
        value_column = start_column + offset
        ws.write(0, value_column, name)
        for row, value in enumerate(values, first_row):
            ws.write_number(row, value_column, float(value))
        series = {
            "name": name,
            "categories": [ws.name, first_row, category_column, last_row, category_column],
            "values": [ws.name, first_row, value_column, last_row, value_column],
        }
        if spec.kind == "line":
            series.update({
                "line": {"color": color, "width": 2.25},
                "marker": {"type": "circle", "size": 4,
                           "border": {"color": color}, "fill": {"color": color}},
            })
        else:
            series.update({
                "fill": {"color": color}, "border": {"none": True},
                "data_labels": {
                    "value": True,
                    "num_format": "0%" if spec.value_kind == "percent" else "0.0",
                },
            })
        chart.add_series(series)
    style_v2_chart(chart, title=spec.title, kind="bar" if spec.kind == "bar" else "line")
    value_axis = {
        "major_gridlines": {"visible": True, "line": {"color": COLORS["line"]}},
    }
    if spec.value_kind == "percent":
        value_axis["num_format"] = "0%"
    if spec.minimum is not None:
        value_axis["min"] = spec.minimum
    if spec.maximum is not None:
        value_axis["max"] = spec.maximum
    if spec.kind == "bar":
        chart.set_x_axis(value_axis)
        chart.set_y_axis({"reverse": True, "major_gridlines": {"visible": False}})
    else:
        chart.set_y_axis(value_axis)
        chart.set_x_axis({"label_position": "low"})
    chart.show_hidden_data()
    return chart, start_column + len(spec.series)


def _value_format(formats, kind: str, value: object):
    if kind == "percent":
        return formats.table_percent
    if kind == "decimal":
        return formats.table_decimal
    if kind == "integer":
        return formats.table_integer
    if kind == "change":
        return formats.positive if isinstance(value, (int, float)) and value >= 0 else formats.negative
    if kind == "alert":
        return formats.due if value not in (0, "", None, "READY", "OK") else formats.clear
    return formats.table_text


def _dashboard(workbook: xlsxwriter.Workbook, spec: DashboardSpec):
    ws = workbook.add_worksheet(spec.sheet)
    formats = make_v2_formats(workbook)
    configure_v2_cell_canvas(ws, formats, zoom=85)
    ws.hide_row_col_headers()
    ws.set_tab_color(COLORS["gold"])
    ws.freeze_panes(2, 0)
    status = "DUMMY DATA" if spec.maturity == "OPERATIONAL" else "IN DEVELOPMENT"
    write_v2_header(
        ws, formats, spec.title, status=status, status_kind="PROVISIONAL",
    )
    write_v2_filters(ws, formats, tuple(
        (label, value, None) for label, value in spec.filters
    ))
    write_v2_kpis(ws, formats, tuple(
        (label, value, kind, None) for label, value, kind in spec.kpis
    ))
    left_chart, left_last = _chart(workbook, ws, spec.left_chart, 29)
    right_chart, right_last = _chart(workbook, ws, spec.right_chart, 35)
    insert_v2_charts(ws, left_chart, right_chart)
    write_v2_section(ws, formats, spec.action_title)
    for index, header in enumerate(spec.action_headers[:7]):
        ws.merge_range(
            ACTION_HEADER_ROW, index * 4, ACTION_HEADER_ROW, index * 4 + 3,
            header, formats.table_header,
        )
    for offset in range(ACTION_VISIBLE_ROWS):
        values = spec.action_rows[offset] if offset < len(spec.action_rows) else tuple(
            "" for _ in spec.action_headers
        )
        for index, value in enumerate(values[:7]):
            fmt = _value_format(formats, spec.action_kinds[index], value)
            ws.merge_range(
                ACTION_FIRST_ROW + offset, index * 4,
                ACTION_FIRST_ROW + offset, index * 4 + 3,
                value, fmt,
            )
    ws.set_column(29, max(left_last, right_last), None, None, {"hidden": True})
    ws.set_footer("&LPrepared by Anass ASSRI | WFM&CPrototype — dummy data&RPage &P of &N")
    return ws


def _table_sheet(
    workbook: xlsxwriter.Workbook,
    name: str,
    title: str,
    subtitle: str,
    headers: tuple[str, ...],
    rows: tuple[tuple[object, ...], ...],
    *,
    editable: tuple[str, ...] = (),
    table_name: str | None = None,
) -> None:
    ws = workbook.add_worksheet(name)
    ws.hide_gridlines(2)
    ws.set_zoom(85)
    ws.freeze_panes(4, 0)
    ws.set_tab_color(COLORS["gold"] if editable else COLORS["teal"])
    title_fmt = workbook.add_format({
        "font_name": "Aptos Display", "font_size": 20, "bold": True,
        "font_color": COLORS["white"], "bg_color": COLORS["navy"],
        "align": "left", "valign": "vcenter", "indent": 1,
    })
    subtitle_fmt = workbook.add_format({
        "font_name": "Aptos", "font_size": 10,
        "font_color": COLORS["white"], "bg_color": COLORS["teal"],
        "align": "left", "valign": "vcenter", "indent": 1,
    })
    edit_fmt = workbook.add_format({
        "font_name": "Aptos", "font_size": 10, "font_color": COLORS["blue"],
        "bg_color": COLORS["blue_light"], "border": 1,
        "border_color": COLORS["line"],
    })
    date_fmt = workbook.add_format({"num_format": "yyyy-mm-dd"})
    datetime_fmt = workbook.add_format({"num_format": "yyyy-mm-dd hh:mm"})
    percent_fmt = workbook.add_format({"num_format": "0.0%"})
    ws.merge_range(0, 0, 0, len(headers) - 1, title, title_fmt)
    ws.merge_range(1, 0, 1, len(headers) - 1, subtitle, subtitle_fmt)
    ws.set_row_pixels(0, 52)
    ws.set_row_pixels(1, 30)
    for row_index, values in enumerate(rows, 4):
        for column, value in enumerate(values):
            fmt = edit_fmt if headers[column] in editable else None
            if fmt is None and isinstance(value, datetime):
                fmt = datetime_fmt
            elif fmt is None and isinstance(value, date):
                fmt = date_fmt
            elif fmt is None and isinstance(value, float) and any(
                word in headers[column].casefold() for word in ("rate", "%", "attainment", "service level", "coverage")
            ):
                fmt = percent_fmt
            ws.write(row_index, column, value, fmt)
    safe_name = table_name or "tblProto" + re.sub(r"[^A-Za-z0-9]", "", name)
    ws.add_table(3, 0, 3 + len(rows), len(headers) - 1, {
        "name": safe_name, "style": "Table Style Medium 2",
        "columns": [{"header": header} for header in headers],
    })
    for column, header in enumerate(headers):
        width = 16
        if any(word in header.casefold() for word in ("agent", "comment", "action", "evidence", "reason")):
            width = 27
        elif any(word in header.casefold() for word in ("date", "time", "interval")):
            width = 19
        ws.set_column(column, column, width)
    ws.set_footer("&LPrepared by Anass ASSRI | WFM&CPrototype — dummy data&RPage &P of &N")


def _help(
    workbook: xlsxwriter.Workbook,
    report: str,
    points: tuple[str, ...],
    *,
    sheet_name: str = "HOW TO TEST",
) -> None:
    ws = workbook.add_worksheet(sheet_name)
    ws.hide_gridlines(2)
    title = workbook.add_format({
        "font_name": "Aptos Display", "font_size": 20, "bold": True,
        "font_color": COLORS["white"], "bg_color": COLORS["navy"],
        "valign": "vcenter", "indent": 1,
    })
    warning = workbook.add_format({
        "font_name": "Aptos", "font_size": 11, "bold": True,
        "font_color": COLORS["navy"], "bg_color": COLORS["amber_light"],
        "valign": "vcenter", "indent": 1,
    })
    body = workbook.add_format({
        "font_name": "Aptos", "font_size": 11, "font_color": COLORS["navy"],
        "text_wrap": True, "valign": "top", "bottom": 1,
        "bottom_color": COLORS["line"],
    })
    ws.merge_range("A1:H1", f"{report} — V2 PROTOTYPE", title)
    ws.merge_range("A3:H3", "DUMMY DATA ONLY · VISUAL AND WORKFLOW VALIDATION", warning)
    for row, point in enumerate(points, 4):
        ws.merge_range(row, 0, row, 7, f"{row - 3}. {point}", body)
        ws.set_row_pixels(row, 52)
    ws.set_column_pixels(0, 7, 150)
    ws.set_tab_color(COLORS["gold"])


def _hidden(workbook: xlsxwriter.Workbook, name: str, note: str) -> None:
    ws = workbook.add_worksheet(name)
    ws.write("A1", note)
    ws.hide()


def _review_board(workbook: xlsxwriter.Workbook) -> None:
    ws = workbook.add_worksheet("REVIEW BOARD")
    ws.hide_gridlines(2)
    ws.set_zoom(80)
    ws.freeze_panes(4, 0)
    title = workbook.add_format({
        "font_name": "Aptos Display", "font_size": 20, "bold": True,
        "font_color": COLORS["white"], "bg_color": COLORS["navy"],
        "valign": "vcenter", "indent": 1,
    })
    subtitle = workbook.add_format({
        "font_name": "Aptos", "font_size": 10,
        "font_color": COLORS["white"], "bg_color": COLORS["teal"],
        "valign": "vcenter", "indent": 1,
    })
    header = workbook.add_format({
        "font_name": "Aptos", "font_size": 9, "bold": True,
        "font_color": COLORS["navy"], "bg_color": "#DCEAF2",
        "align": "center", "valign": "vcenter", "border": 1,
        "border_color": COLORS["line"],
    })
    schedule = workbook.add_format({
        "font_name": "Aptos", "font_size": 10, "font_color": COLORS["navy"],
        "bg_color": COLORS["white"], "border": 1, "border_color": COLORS["line"],
    })
    actual = workbook.add_format({
        "font_name": "Aptos", "font_size": 10, "font_color": COLORS["teal"],
        "bg_color": COLORS["teal_light"], "border": 1, "border_color": COLORS["line"],
    })
    editable = workbook.add_format({
        "font_name": "Aptos", "font_size": 10, "font_color": COLORS["blue"],
        "bg_color": COLORS["blue_light"], "border": 1, "border_color": COLORS["blue"],
    })
    headers = (
        "BAND", "GAP ID", "DATE", "LOB", "AGENT ID", "AGENT",
        "EXACT START", "EXACT END", "MINUTES", "EVIDENCE",
        "DECISION CATEGORY", "DECISION STATUS", "REVIEWED BY",
        "COMMENT", "REVIEWED DATE",
    )
    ws.merge_range(0, 0, 0, len(headers) - 1, "ATTENDANCE REVIEW BOARD", title)
    ws.merge_range(
        1, 0, 1, len(headers) - 1,
        "Schedule is directly above observed activity · edit only blue ACTUAL decision fields",
        subtitle,
    )
    for column, value in enumerate(headers):
        ws.write(3, column, value, header)
    cases = (
        ("GAP-260907-001", "RSA NL", "00124", "Amina El Idrissi", "09:00", "17:00", "10:22", "10:47", 25, "Agent Status", "Break overrun"),
        ("GAP-260907-002", "RSA BE", "00325", "Lucas Peeters", "08:30", "16:30", "08:30", "09:04", 34, "LILO fallback", "Late start"),
        ("GAP-260907-003", "OEM", "00518", "Omar Diallo", "10:00", "18:00", "16:41", "17:06", 25, "Agent Status", "Internal gap"),
    )
    row = 4
    for gap_id, lob, agent_id, agent_name, shift_start, shift_end, gap_start, gap_end, minutes, evidence, reason in cases:
        scheduled_row = ("SCHEDULE", gap_id, date(2026, 9, 7), lob, agent_id, agent_name, shift_start, shift_end, 480, "Verint StartEnd", "", "", "", "", "")
        actual_row = ("ACTUAL", gap_id, date(2026, 9, 7), lob, agent_id, agent_name, gap_start, gap_end, minutes, evidence, reason, "Open", "", "", "")
        for column, value in enumerate(scheduled_row):
            ws.write(row, column, value, schedule)
        for column, value in enumerate(actual_row):
            ws.write(row + 1, column, value, editable if column in {10, 11, 12, 13, 14} else actual)
        ws.set_row_pixels(row + 2, 12)
        row += 3
    widths = (12, 22, 14, 14, 14, 24, 12, 12, 10, 20, 16, 18, 22, 18, 30)
    for column, width in enumerate(widths):
        ws.set_column(column, column, width)
    ws.set_footer("&LPrepared by Anass ASSRI | WFM&CPrototype — dummy data&RPage &P of &N")


def _rtm(path: Path) -> None:
    workbook = xlsxwriter.Workbook(path)
    _properties(workbook, "RTM Daily Control")
    lobs = ("RSA NL", "RSA BE", "FORD NL", "OEM")
    sl = (.899, .858, .973, .949)
    targets = (.80, .80, .80, .80)
    actual = (840, 530, 320, 410)
    forecast = (810, 552, 300, 354)
    _dashboard(workbook, DashboardSpec(
        "CONTROL", "RTM DAILY CONTROL", "OPERATIONAL",
        (("Date", "08 Sep 2026"), ("Snapshot", "Through 17:59"), ("LOB", "All 4"), ("Attendance", "17:55")),
        (("LOBS ON TARGET", "4 / 4", "text"), ("DEMAND VARIANCE", 84, "integer"),
         ("NO SHOW HC", 6, "integer"), ("CALL NOW", 8, "integer")),
        ChartSpec("SERVICE LEVEL BY LOB", "bar", lobs,
                  (("TSL", sl, COLORS["teal"]), ("Target", targets, COLORS["muted"])), "percent", 0, 1),
        ChartSpec("ACTUAL VS FORECAST", "column", lobs,
                  (("Actual", actual, COLORS["teal"]), ("Forecast", forecast, COLORS["muted"]))),
        "LOB SERVICE & ATTENDANCE ACTIONS",
        ("LOB", "TSL", "VAR CALLS", "NO SHOW", "LATE", "EARLY LEAVE", "CALL NOW"),
        tuple((lob, sl[i], actual[i] - forecast[i], (2, 1, 1, 2)[i], (1, 2, 0, 1)[i], (0, 1, 1, 0)[i], (2, 2, 1, 3)[i]) for i, lob in enumerate(lobs)),
        ("text", "percent", "change", "integer", "integer", "integer", "alert"),
    ))
    hours = ("07", "08", "09", "10", "11", "12", "13", "14", "15", "16", "17", "18")
    for index, lob in enumerate(lobs):
        hourly_actual = tuple(35 + ((hour * 7 + index * 11) % 42) for hour in range(len(hours)))
        hourly_forecast = tuple(38 + ((hour * 5 + index * 9) % 37) for hour in range(len(hours)))
        hourly_sl = tuple(min(.99, .78 + ((hour + index * 2) % 8) * .025) for hour in range(len(hours)))
        lob_ws = _dashboard(workbook, DashboardSpec(
            lob, f"{lob} SERVICE FLASH", "OPERATIONAL",
            (("Date", "08 Sep 2026"), ("Snapshot", "17:59"), ("Queue scope", "Configured"), ("Attendance", "17:55")),
            (("TSL", sl[index], "percent"), ("OFFERED", actual[index], "integer"),
             ("VOLUME VARIANCE", actual[index] - forecast[index], "integer"),
             ("NO SHOW HC", (2, 1, 1, 2)[index], "integer")),
            ChartSpec("HOURLY ACTUAL VS FORECAST", "column", hours,
                      (("Actual", hourly_actual, COLORS["teal"]), ("Forecast", hourly_forecast, COLORS["muted"]))),
            ChartSpec("HOURLY SERVICE LEVEL", "line", hours,
                      (("TSL", hourly_sl, COLORS["teal"]), ("Target", tuple(.80 for _ in hours), COLORS["muted"])), "percent", 0, 1),
            "ATTENDANCE PULSE & CALL LIST",
            ("AGENT", "SHIFT", "STATE NOW", "NO SHOW", "LATE", "EARLY LEAVE", "ACTION"),
            (("Amina El Idrissi [00124]", "09:00–17:00", "Logged", 0, 1, 0, "Monitor"),
             ("Youssef Benali [00135]", "08:30–16:30", "No show", 1, 0, 0, "Call now"),
             ("Nora Amrani [00148]", "10:00–18:00", "PTO", 0, 0, 0, "No call")),
            ("text", "text", "text", "integer", "integer", "integer", "alert"),
        ))
        detail_headers = [
            "Hour", "Forecast", "Actual", "TSL", "Target", "No Show HC", "Call Now",
        ]
        if lob == "OEM":
            detail_headers.extend((
                "Ford Volume", "Ford TSL", "Toyota Volume", "Toyota TSL",
                "Chery Volume", "Chery TSL",
            ))
        detail_rows = []
        for hour_index, hour in enumerate(hours):
            values: list[object] = [
                f"{hour}:00", hourly_forecast[hour_index], hourly_actual[hour_index],
                hourly_sl[hour_index], .80, hour_index % 3, hour_index % 2,
            ]
            if lob == "OEM":
                values.extend((
                    round(hourly_actual[hour_index] * .45), min(.99, hourly_sl[hour_index] + .02),
                    round(hourly_actual[hour_index] * .35), max(0, hourly_sl[hour_index] - .01),
                    round(hourly_actual[hour_index] * .20), min(.99, hourly_sl[hour_index] + .01),
                ))
            detail_rows.append(tuple(values))
        detail_header_row = 34
        for column, value in enumerate(detail_headers):
            lob_ws.write(detail_header_row, column, value)
        for row_index, values in enumerate(detail_rows, detail_header_row + 1):
            for column, value in enumerate(values):
                lob_ws.write(row_index, column, value)
        lob_ws.add_table(
            detail_header_row, 0, detail_header_row + len(detail_rows),
            len(detail_headers) - 1,
            {
                "name": "tblProto" + re.sub(r"[^A-Za-z0-9]", "", lob) + "Hourly",
                "style": "Table Style Medium 2",
                "columns": [{"header": value} for value in detail_headers],
            },
        )
    _table_sheet(
        workbook, "ISSUES & DRIVERS", "RTM ISSUES & DRIVERS",
        "Only actionable missing sources and below-target queues · dummy data",
        ("Type", "LOB", "Hour / Queue", "Volume", "TSL", "Target", "Gap", "Driver", "Action"),
        (("Queue driver", "RSA BE", "RSA_BE_FR_02", 84, .74, .80, -.06, "Late routed contacts", "Validate routing"),
         ("Source issue", "OEM", "Forecast 18:00", 0, "", "", "", "No forecast row", "Check extract")),
    )
    _help(workbook, "RTM DAILY CONTROL", (
        "Judge CONTROL density first: service and attendance must tell one story without extra cards.",
        "Open every LOB tab and verify the full-day hourly view, four KPIs and call list are immediately readable.",
        "Confirm No Show means evidence-proven only; PTO/Away is visible but never called.",
        "Confirm Issues & Drivers remains a detail table and does not clutter the first screen.",
    ))
    _hidden(workbook, "DEFINITIONS", "Prototype placeholder for governed RTM definitions.")
    _hidden(workbook, "_AUDIT", "Prototype placeholder for source and rule lineage.")
    workbook.close()


def _attendance(path: Path) -> None:
    workbook = xlsxwriter.Workbook(path)
    _properties(workbook, "Attendance Review")
    lobs = ("RSA NL", "RSA BE", "FORD NL", "OEM")
    gaps = (6, 5, 4, 3)
    hours = (2.8, 2.1, 1.6, .9)
    _dashboard(workbook, DashboardSpec(
        "CONTROL", "ATTENDANCE REVIEW", "OPERATIONAL",
        (("Period", "01–07 Sep 2026"), ("Completed", "Through 07 Sep"), ("Evidence", "Status + LILO"), ("Decision", "Review board")),
        (("REVIEW GAPS", 18, "integer"), ("GAP HOURS", 7.4, "decimal"),
         ("OPEN DECISIONS", 11, "integer"), ("MISSING EVIDENCE", 2, "integer")),
        ChartSpec("GAP HOURS BY LOB", "bar", lobs, (("Gap hours", hours, COLORS["teal"]),)),
        ChartSpec("DECISION STATUS", "column", ("Open", "Approved", "Dismissed", "Missing"),
                  (("Cases", (11, 4, 3, 2), COLORS["teal"]),)),
        "BY-LOB REVIEW ACTIONS",
        ("LOB", "EXACT GAPS", "GAP HOURS", "AGENTS", "OPEN", "APPROVED", "DISMISSED"),
        tuple((lob, gaps[i], hours[i], (4, 4, 3, 2)[i], (3, 4, 2, 2)[i], (2, 1, 1, 1)[i], (1, 0, 1, 0)[i]) for i, lob in enumerate(lobs)),
        ("text", "integer", "decimal", "integer", "alert", "integer", "integer"),
    ))
    _review_board(workbook)
    _table_sheet(
        workbook, "BREAK & MEAL", "BREAK & MEAL CONTROL",
        "Completed shifts only · explicit Meal Aux stays Lunch · dummy data",
        ("Date", "LOB", "Agent ID", "Agent", "Break Min", "Break Limit", "Meal Min", "Meal Limit", "Overrun Min", "Evidence", "Action"),
        ((date(2026, 9, 7), "RSA NL", "00124", "Amina El Idrissi", 31, 30, 60, 60, 1, "Agent Status", "Review break"),
         (date(2026, 9, 7), "OEM", "00518", "Omar Diallo", 27, 30, 74, 60, 14, "Agent Status", "Review meal"),
         (date(2026, 9, 7), "FORD NL", "00417", "Ines Verhoeven", 24, 30, 57, 60, 0, "Agent Status", "Within limits"),
         (date(2026, 9, 7), "RSA BE", "00339", "Mila Laurent", "", 30, "", 60, "", "Coverage incomplete", "Insufficient evidence")),
    )
    _help(workbook, "ATTENDANCE REVIEW", (
        "Validate that CONTROL tells you the size and ownership of the review workload.",
        "In REVIEW BOARD, verify every SCHEDULE row sits directly above ACTUAL and each case has a visible spacer.",
        "Only the five blue ACTUAL fields are intended for human decisions.",
        "Confirm BREAK & MEAL is separate, simple and based on completed-shift evidence.",
    ))
    _hidden(workbook, "DECISION LEDGER", "Prototype placeholder for imported decisions by immutable Gap ID.")
    _hidden(workbook, "EVIDENCE", "Prototype placeholder for exact schedule and observed intervals.")
    _hidden(workbook, "DEFINITIONS", "Prototype placeholder for governed attendance definitions.")
    _hidden(workbook, "_LOOKUPS", "Prototype placeholder for controlled decision lists.")
    _hidden(workbook, "_AUDIT", "Prototype placeholder for source and rule lineage.")
    workbook.close()


def _staffing(path: Path) -> None:
    workbook = xlsxwriter.Workbook(path)
    _properties(workbook, "Staffing & Coverage")
    lobs = ("RSA NL", "RSA BE", "FORD NL", "OEM")
    required = (420, 310, 245, 280)
    net = (401, 298, 251, 260)
    _dashboard(workbook, DashboardSpec(
        "DASHBOARD", "STAFFING & COVERAGE", "IN DEVELOPMENT",
        (("Period", "Current + 4 weeks"), ("Mode", "Future plan"), ("LOB", "All"), ("Language", "All")),
        (("PEAK GAP FTE", 4.5, "decimal"), ("FUTURE GAP HOURS", 36.8, "decimal"),
         ("FORECAST COVERAGE", .943, "percent"), ("PTO / AWAY IMPACT", 61.0, "decimal")),
        ChartSpec("REQUIRED VS NET SCHEDULED HOURS", "column", lobs,
                  (("Required", required, COLORS["teal"]), ("Net scheduled", net, COLORS["muted"]))),
        ChartSpec("PEAK GAP BY DAY", "line", ("Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"),
                  (("Gap FTE", (2.1, 3.4, 4.5, 2.8, 3.9, 1.2, .8), COLORS["red"]),)),
        "PRIORITIZED CAPACITY ACTIONS",
        ("INTERVAL", "LOB / LANGUAGE", "MODE", "REQUIRED FTE", "NET / OBSERVED", "GAP FTE", "STATE"),
        (("09 Sep 10:00", "RSA BE / FR-NL", "Future", 22.5, 18.0, 4.5, "FUTURE GAP"),
         ("09 Sep 11:15", "RSA NL / NL", "Future", 25.0, 21.6, 3.4, "FUTURE GAP"),
         ("08 Sep 16:30", "OEM / FR", "Actual", 19.0, 16.2, 2.8, "PARTIAL GAP"),
         ("10 Sep 14:00", "FORD NL / NL", "Future", 14.0, 15.0, 0.0, "OK")),
        ("text", "text", "text", "decimal", "decimal", "decimal", "alert"),
    ))
    _table_sheet(workbook, "WEEKLY_PLAN", "WEEKLY CAPACITY PLAN", "Required versus net schedule after PTO/Away · dummy data",
                 ("ISO Week", "LOB", "Language", "Required Hours", "Gross Scheduled", "PTO/Away Hours", "Net Scheduled", "Gap Hours", "Coverage Rate", "State"),
                 (("2026-W37", "RSA NL", "NL", 420, 419, 18, 401, 19, .955, "FUTURE GAP"),
                  ("2026-W37", "RSA BE", "FR / NL", 310, 310, 12, 298, 12, .961, "FUTURE GAP"),
                  ("2026-W37", "FORD NL", "NL", 245, 260, 9, 251, 0, 1.0, "FUTURE OK")))
    _table_sheet(workbook, "INTRADAY", "15-MINUTE CAPACITY", "Exact interval planning and actual control · dummy data",
                 ("Date", "Interval", "LOB", "Required FTE", "Net Scheduled FTE", "Observed FTE", "Gap FTE", "Mode", "State"),
                 ((date(2026, 9, 9), "10:00–10:15", "RSA BE", 22.5, 18.0, "", 4.5, "FUTURE PLAN", "FUTURE GAP"),
                  (date(2026, 9, 9), "10:15–10:30", "RSA NL", 25.0, 21.6, "", 3.4, "FUTURE PLAN", "FUTURE GAP")))
    _table_sheet(workbook, "ACTIONS", "STAFFING ACTIONS", "Exact shortages to prepare with Operations · dummy data",
                 ("Week", "Date", "Interval", "LOB", "Gap FTE", "Cause", "Owner", "Action", "Status", "Comment"),
                 (("2026-W37", date(2026, 9, 9), "10:00–10:15", "RSA BE", 4.5, "Net schedule below requirement", "WFM", "Review staffing move", "Open", ""),))
    _help(workbook, "STAFFING & COVERAGE", (
        "Judge whether the dashboard separates future preparation from historical actual control.",
        "Confirm PTO/Away impact is visible before the shortage is calculated.",
        "Validate weekly management view, exact 15-minute evidence and a small owned action list.",
        "This remains IN DEVELOPMENT until forecast-to-roster mappings and operating decisions are validated.",
    ))
    _hidden(workbook, "DEFINITIONS", "Prototype placeholder for governed staffing definitions.")
    _hidden(workbook, "_AUDIT", "Prototype placeholder for source and rule lineage.")
    workbook.close()


def _realisations(path: Path) -> None:
    workbook = xlsxwriter.Workbook(path)
    _properties(workbook, "Realisations")
    lobs = ("RSA NL", "RSA BE", "FORD NL", "OEM")
    actual = (18400, 12150, 8830, 10240)
    forecast = (17950, 12600, 8460, 9980)
    sl = (.899, .858, .973, .949)
    target = (.80, .80, .80, .80)
    _dashboard(workbook, DashboardSpec(
        "DASHBOARD", "REALISATIONS", "IN DEVELOPMENT",
        (("Period", "Current MTD"), ("LOB", "All mapped"), ("Grain", "Daily"), ("Data state", "Reviewed")),
        (("ACTUAL VOLUME", sum(actual), "integer"), ("FORECAST ATTAINMENT", sum(actual) / sum(forecast), "percent"),
         ("ROUTED RATE", .958, "percent"), ("WEIGHTED AHT", 294.0, "decimal")),
        ChartSpec("ACTUAL VS FORECAST BY LOB", "column", lobs,
                  (("Actual", actual, COLORS["teal"]), ("Forecast", forecast, COLORS["muted"]))),
        ChartSpec("SERVICE LEVEL VS TARGET", "bar", lobs,
                  (("Service level", sl, COLORS["teal"]), ("Target", target, COLORS["muted"])), "percent", 0, 1),
        "LOB REALISATION SUMMARY",
        ("LOB", "ACTUAL", "FORECAST", "ATTAINMENT", "SERVICE LEVEL", "ABSENCE %", "STATE"),
        tuple((lob, actual[i], forecast[i], actual[i] / forecast[i], sl[i], (.041, .052, .033, .046)[i], "READY") for i, lob in enumerate(lobs)),
        ("text", "integer", "integer", "percent", "percent", "percent", "alert"),
    ))
    _table_sheet(workbook, "LOB_RESULTS", "DAILY LOB RESULTS", "Additive daily counters behind management ratios · dummy data",
                 ("Date", "LOB", "Actual Volume", "Forecast Volume", "Attainment", "Service Level", "Routed Rate", "AHT Seconds", "Absence Rate", "Shrinkage Rate", "State"),
                 ((date(2026, 9, 7), "RSA NL", 2410, 2350, 1.026, .899, .963, 292, .041, .118, "Final"),
                  (date(2026, 9, 7), "RSA BE", 1650, 1710, .965, .858, .952, 315, .052, .129, "Review")))
    _table_sheet(workbook, "TREND", "PERIOD TREND", "Month, week and quarter summaries from summed counters · dummy data",
                 ("LOB", "Grain", "Period", "Actual", "Forecast", "Attainment", "Service Level", "AHT", "Absence Rate", "Shrinkage Rate"),
                 (("RSA NL", "ISO Week", "2026-W36", 11840, 11590, 1.022, .894, 295, .043, .121),
                  ("RSA NL", "ISO Week", "2026-W37", 6560, 6360, 1.031, .907, 286, .038, .113)))
    _table_sheet(workbook, "DATA", "MAPPED SERVICE DATA", "Queue and interval evidence · dummy data",
                 ("Date", "Interval", "Reporting LOB", "Queue", "Offered", "Answered", "Within Target", "Handled Seconds", "Mapping Status"),
                 ((date(2026, 9, 7), "10:00", "RSA NL", "RSA_NL_01", 64, 61, 57, 17640, "MAPPED"),))
    _help(workbook, "REALISATIONS", (
        "Validate the management question: what was forecast, what happened, and where did performance diverge?",
        "Ratios must remain based on summed counters; the prototype numbers are illustrative only.",
        "Check that daily LOB results, period trend and queue evidence form a clear drill path.",
        "This remains IN DEVELOPMENT until Operations validates the final management scope.",
    ))
    _hidden(workbook, "DEFINITIONS", "Prototype placeholder for governed realisation definitions.")
    _hidden(workbook, "_AUDIT", "Prototype placeholder for source and rule lineage.")
    workbook.close()


def _absence(path: Path) -> None:
    workbook = xlsxwriter.Workbook(path)
    _properties(workbook, "Absenteeism & Shrinkage")
    lobs = ("RSA NL", "RSA BE", "FORD NL", "OEM")
    absence = (.041, .052, .033, .046)
    shrinkage = (.118, .129, .104, .123)
    _dashboard(workbook, DashboardSpec(
        "DASHBOARD", "ABSENTEEISM & SHRINKAGE", "IN DEVELOPMENT",
        (("Period", "Current MTD"), ("LOB", "All"), ("Team Leader", "All"), ("Agent", "All")),
        (("ABSENCE RATE", .043, "percent"), ("SHRINKAGE RATE", .118, "percent"),
         ("FINALIZED COVERAGE", .988, "percent"), ("REVIEW CASES", 7, "integer")),
        ChartSpec("ABSENCE & SHRINKAGE BY LOB", "bar", lobs,
                  (("Absence", absence, COLORS["red"]), ("Shrinkage", shrinkage, COLORS["teal"])), "percent", 0, .18),
        ChartSpec("DAILY ABSENCE & SHRINKAGE HOURS", "line", ("01", "02", "03", "04", "05", "06", "07"),
                  (("Absence hours", (31, 27, 38, 42, 35, 29, 41), COLORS["red"]),
                   ("Shrinkage hours", (72, 68, 81, 77, 74, 69, 83), COLORS["teal"]))),
        "PRIORITIZED ABSENCE REVIEW CASES",
        ("DATE", "AGENT", "TEAM LEADER", "LOB", "RESULT STATUS", "ABSENCE H", "ACTION STATUS"),
        (("07 Sep", "Amina El Idrissi", "Sophie Martin", "RSA NL", "ABSENCE RECORDED", 8.0, "Pending"),
         ("07 Sep", "Lucas Peeters", "Elena Rossi", "RSA BE", "UNCODED EMPTY SHIFT", 0.0, "Needs review"),
         ("06 Sep", "Omar Diallo", "Laura Jensen", "OEM", "UNMAPPED INTERVAL", 2.3, "Needs review")),
        ("text", "text", "text", "text", "text", "decimal", "alert"),
    ))
    _table_sheet(workbook, "TEAM_VIEW", "TEAM ABSENCE VIEW", "LOB and Team Leader ratio-of-sums view · dummy data",
                 ("LOB", "Team Leader", "Agents", "Planned Hours", "Absence Hours", "Absence Rate", "Shrinkage Hours", "Shrinkage Rate", "Review Cases"),
                 (("RSA NL", "Sophie Martin", 12, 960, 38, .040, 112, .117, 1),
                  ("RSA BE", "Elena Rossi", 11, 880, 49, .056, 118, .134, 2)))
    _table_sheet(workbook, "TEAM_SUMMARY", "TEAM ABSENCE & SHRINKAGE", "Team ratio-of-sums summary · dummy data",
                 ("LOB", "Team Leader", "Agents", "Planned Hours", "Absence Hours", "Absence Rate", "Shrinkage Hours", "Shrinkage Rate", "Review Cases"),
                 (("RSA NL", "Sophie Martin", 12, 960, 38, .040, 112, .117, 1),
                  ("RSA BE", "Elena Rossi", 11, 880, 49, .056, 118, .134, 2)))
    _table_sheet(workbook, "AGENT_RESULTS", "AGENT ABSENCE & SHRINKAGE", "Filter by LOB, Team Leader or Agent · dummy data",
                 ("LOB", "Team Leader", "Agent ID", "Agent", "Planned Hours", "Absence Hours", "Absence Rate", "Shrinkage Hours", "Shrinkage Rate", "Review Cases"),
                 (("RSA NL", "Sophie Martin", "00124", "Amina El Idrissi", 176, 8, .045, 21, .119, 0),
                  ("RSA BE", "Elena Rossi", "00325", "Lucas Peeters", 168, 0, 0, 18, .107, 1)))
    _table_sheet(workbook, "ACTIONS", "ABSENCE REVIEW & FOLLOW-UP", "Permanent Case ID keeps collaborative work attached · dummy data",
                 ("Case ID", "Date", "LOB", "Agent ID", "Agent", "Result Status", "Absence Hours", "Shrinkage Hours", "Review Status", "Owner", "Due Date", "Action", "Comment"),
                 (("ABS-20260907-00124", date(2026, 9, 7), "RSA NL", "00124", "Amina El Idrissi", "ABSENCE_RECORDED", 8.0, 8.0, "Pending", "", "", "", ""),
                  ("ABS-20260907-00325", date(2026, 9, 7), "RSA BE", "00325", "Lucas Peeters", "UNCODED_EMPTY_SHIFT", 0.0, 0.0, "Needs review", "", "", "", "")),
                 editable=("Review Status", "Owner", "Due Date", "Action", "Comment"))
    _table_sheet(workbook, "ACTION_QUEUE", "ABSENCE ACTION QUEUE", "Refreshable cases linked to permanent ACTIONS by Case ID · dummy data",
                 ("Case ID", "Date", "LOB", "Agent ID", "Agent", "Result Status", "Planned Hours", "Absence Hours", "Shrinkage Hours", "Action Status"),
                 (("ABS-20260907-00124", date(2026, 9, 7), "RSA NL", "00124", "Amina El Idrissi", "ABSENCE_RECORDED", 8, 8, 8, "Pending"),
                  ("ABS-20260907-00325", date(2026, 9, 7), "RSA BE", "00325", "Lucas Peeters", "UNCODED_EMPTY_SHIFT", 8, 0, 0, "Needs review")))
    _table_sheet(workbook, "ABSENCE_COMPONENTS", "ABSENCE COMPONENTS", "Exclusive absence components reconcile to absence totals · dummy data",
                 ("Component", "Hours", "Percent of Planned", "Evidence", "Mapped"),
                 (("Sickness", 176, .030, "Reviewed attendance decision", "YES"),
                  ("Unpaid", 58, .010, "PTO/Away + reviewed decision", "YES"),
                  ("Vacation", 212, .036, "Approved PTO", "YES")))
    _table_sheet(workbook, "SHRINKAGE_COMPONENTS", "SHRINKAGE COMPONENTS", "Parallel shrinkage view; never add it to absence · dummy data",
                 ("Component", "Hours", "Percent of Planned", "Evidence", "Mapped"),
                 (("Training", 146, .025, "Reviewed activity", "YES"),
                  ("Break / Lunch", 97, .017, "Agent Status", "YES"),
                  ("Absence inside shrinkage", 252, .043, "Reviewed decision", "YES")))
    _table_sheet(workbook, "COMPONENT_VIEW", "COMPONENT ANALYSIS", "Absence and shrinkage stay visibly separate · dummy data",
                 ("Scope", "Component", "Hours", "Share of Scope", "Trend"),
                 (("Absence", "Sickness", 176, .699, "Stable"), ("Absence", "Unpaid", 58, .230, "Up"),
                  ("Shrinkage", "Training", 146, .209, "Down")))
    _table_sheet(workbook, "ACTIVITY_DETAIL", "EXACT ACTIVITY DETAIL", "Exact start/end evidence behind classified components · dummy data",
                 ("Date", "LOB", "Agent ID", "Agent", "Activity", "Category", "Start", "End", "Minutes", "Evidence", "Event Key"),
                 ((date(2026, 9, 7), "RSA NL", "00124", "Amina El Idrissi", "Sickness", "ABSENCE", "09:00", "17:00", 480, "Attendance decision", "EVT-001"),))
    _table_sheet(workbook, "ABSENCE_DATA", "ABSENCE AGENT-DAY DATA", "Reconciliation counters used by all report views · dummy data",
                 ("Date", "LOB", "Team Leader", "Agent ID", "Agent", "Planned Minutes", "Absence Minutes", "Shrinkage Minutes", "Ledger Status", "Case ID"),
                 ((date(2026, 9, 7), "RSA NL", "Sophie Martin", "00124", "Amina El Idrissi", 480, 480, 480, "ABSENCE_RECORDED", "ABS-20260907-00124"),))
    _help(workbook, "ABSENTEEISM & SHRINKAGE", (
        "Validate that absence and shrinkage remain parallel, clearly separated views.",
        "Check that unresolved empty shifts and unmapped intervals remain review cases, never zero.",
        "Confirm the dashboard, team drill-down, components and permanent action log are enough for collaboration.",
        "This remains IN DEVELOPMENT and is not payroll-ready until the final decision workflow is validated.",
    ), sheet_name="HELP")
    _hidden(workbook, "DEFINITIONS", "Prototype placeholder for governed absence and shrinkage definitions.")
    _hidden(workbook, "_LOOKUPS", "Prototype placeholder for controlled filters and lists.")
    _hidden(workbook, "_AUDIT", "Prototype placeholder for source and rule lineage.")
    workbook.close()


def _bonus(path: Path) -> None:
    workbook = xlsxwriter.Workbook(path)
    _properties(workbook, "Bonus Management")
    lobs = ("RSA NL", "RSA BE", "FORD NL", "OEM")
    payouts = (184500, 132400, 97600, 118900)
    dashboard_spec = DashboardSpec(
        "Dashboard", "BONUS MANAGEMENT", "IN DEVELOPMENT",
        (("Period", "2026-07"), ("Population", "All"), ("Team Lead", "All"), ("Result status", "All")),
        (("TOTAL PAYOUT", sum(payouts), "integer"), ("PAID AGENTS", 107, "integer"),
         ("AVERAGE PAID PAYOUT", sum(payouts) / 107, "decimal"), ("REVIEW ITEMS", 9, "integer")),
        ChartSpec("ESTIMATED PAYOUT BY LOB", "column", lobs, (("MAD", payouts, COLORS["teal"]),)),
        ChartSpec(
            "KPI ATTAINMENT OVERVIEW", "column",
            ("AHT", "Productivity", "PCS", "Participation", "QM", "Abs%"),
            (("Achievement", (.91, .97, .88, .84, .93, .86), COLORS["gold"]),),
            "percent", 0, 1.2,
        ),
        "POPULATION BONUS SUMMARY",
        ("POPULATION", "AGENTS", "PAID AGENTS", "PAYOUT RATE", "TOTAL PAYOUT", "AVG ACHIEVEMENT", "REVIEW"),
        (("RSA NL", 38, 36, .947, payouts[0], .886, 2), ("RSA BE", 29, 25, .862, payouts[1], .817, 4),
         ("FORD NL", 21, 21, 1.0, payouts[2], .928, 1), ("OEM", 28, 25, .893, payouts[3], .856, 2)),
        ("text", "integer", "integer", "percent", "integer", "percent", "alert"),
    )
    _table_sheet(workbook, "Policy_Decisions", "POLICY DECISIONS", "Management/HR decisions required before release · dummy data",
                 ("Policy", "Selected Decision", "Allowed Values", "Formula Impact", "Owner", "Status", "Comments"),
                 (("Absence treatment", "KPI only", "KPI only | Eligibility only | Both", "Avoid duplicate absence impact", "HR / Operations", "Validated", ""),
                  ("Rounding", "Centime", "Centime | Whole MAD", "Final payout decimals", "Payroll", "To validate", "")),
                 editable=("Selected Decision", "Owner", "Status", "Comments"))
    _table_sheet(workbook, "Control_Checks", "CONTROL CHECKS", "Resolve REVIEW before management or payroll use · dummy data",
                 ("Area", "Check", "Count", "Expected", "Status", "Owner", "Evidence", "Comment"),
                 (("Monetary inputs", "Missing currency", 0, 0, "OK", "WFM / HR", "Currency required", ""),
                  ("Policy", "Policy decisions not validated", 2, 0, "REVIEW", "WFM / HR", "Owner decision required", "")),
                 editable=("Comment",))
    _table_sheet(workbook, "KPI_Config", "KPI CONFIGURATION", "Blue cells are governed policy inputs · dummy data",
                 ("Population", "KPI", "Direction", "Tier 1 Bonus %", "Tier 1 Target", "Tier 2 Bonus %", "Tier 2 Target"),
                 (("RSA", "PCS Score", "H", .15, 4.2, .08, 4.0), ("RSA", "Abs%", "L", .10, .04, .05, .06)),
                 editable=("Direction", "Tier 1 Bonus %", "Tier 1 Target", "Tier 2 Bonus %", "Tier 2 Target"))
    _table_sheet(workbook, "Raw_Data", "RAW PERFORMANCE DATA", "Paste/validate monthly inputs · dummy data",
                 ("Agent ID", "Agent", "Population", "Period", "AHT", "Productivity", "PCS Score", "PCS Participation", "QM", "Abs%", "VOC Detractors", "Currency", "Salary", "Bonus Rate", "Eligible Days", "Scheduled Days", "Status"),
                 (("00124", "Amina El Idrissi", "RSA", "2026-07", 284, 1.04, 4.35, .29, .94, .031, 1, "MAD", 12000, .12, 22, 22, "VALIDATED"),),
                 editable=("Agent ID", "Agent", "Population", "Period", "AHT", "Productivity", "PCS Score", "PCS Participation", "QM", "Abs%", "VOC Detractors", "Currency", "Salary", "Bonus Rate", "Eligible Days", "Scheduled Days", "Status"))
    _table_sheet(workbook, "Results", "BONUS RESULTS", "Formula-result structure inherited from Bonus Matrix v1.2 · dummy data",
                 ("Agent ID", "Agent", "LOB", "Period", "Gross Achievement", "Malus Impact", "Final Achievement", "Reference Bonus", "Proration", "Final Payout", "Status", "Team Lead"),
                 (("00124", "Amina El Idrissi", "RSA NL", "2026-07", .96, .10, .864, 4500, 1.0, 3888, "Eligible", "Sophie Martin"),
                  ("00325", "Lucas Peeters", "RSA BE", "2026-07", .83, .20, .664, 4300, .92, 2626.66, "Malus applied", "Elena Rossi")))
    _table_sheet(workbook, "KPI_Analysis", "KPI ANALYSIS", "Read-only KPI award and attainment summary · dummy data",
                 ("Population", "KPI", "Agents", "Average Actual", "Tier 1 Earned", "Tier 2 Earned", "No Award"),
                 (("RSA", "PCS Score", 67, 4.18, 31, 22, 14), ("RSA", "Abs%", 67, .046, 29, 21, 17)))
    _table_sheet(workbook, "Team_Lead_Analysis", "TEAM LEAD ANALYSIS", "Read-only team payout and review summary · dummy data",
                 ("Team Lead", "Population", "Agents", "Paid Agents", "Average Achievement", "Total Payout", "Review Items"),
                 (("Sophie Martin", "RSA NL", 12, 12, .904, 59800, 1),
                  ("Elena Rossi", "RSA BE", 11, 9, .812, 43100, 2)))
    dashboard = _dashboard(workbook, dashboard_spec)
    dashboard.activate()
    _help(workbook, "BONUS MANAGEMENT", (
        "Compare the management dashboard with the Bonus Matrix v1.2 workflow, not with an invented KPI policy.",
        "Validate the sequence: Raw Data → KPI Config/Policy Decisions → Control Checks → Results.",
        "Blue means editable policy or input; payout results are calculated and must remain non-editable.",
        "This remains IN DEVELOPMENT until every policy decision and payroll control is validated.",
    ))
    workbook.close()


def build_all() -> tuple[Path, ...]:
    OUTPUT.mkdir(parents=True, exist_ok=True)
    builders = (
        ("RTM-Daily-Control-V2-Prototype-Dummy.xlsx", _rtm),
        ("Attendance-Review-V2-Prototype-Dummy.xlsx", _attendance),
        ("Staffing-Coverage-V2-Prototype-Dummy.xlsx", _staffing),
        ("Realisations-V2-Prototype-Dummy.xlsx", _realisations),
        ("Absence-Shrinkage-V2-Prototype-Dummy.xlsx", _absence),
        ("Bonus-Management-V2-Prototype-Dummy.xlsx", _bonus),
    )
    paths = []
    for filename, builder in builders:
        path = OUTPUT / filename
        builder(path)
        paths.append(path)
    with ZipFile(PACK, "w", ZIP_DEFLATED) as archive:
        for path in paths:
            archive.write(path, path.name)
    return tuple(paths)


if __name__ == "__main__":
    for generated in build_all():
        print(generated)
    print(PACK)
