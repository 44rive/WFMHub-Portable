#!/usr/bin/env python3
"""Build the formula-free PCS Power Query/slicer acceptance prototype.

This workbook is a visual and workflow prototype, not the production tracker.
It deliberately contains no custom selector cells, calculation formulas,
Power Query connection XML, PivotCaches, or slicer XML.  It shows the compact
presentation and lightweight query-output tables that those native Excel
features will drive in production. Raw call-leg data remains in the external
clean CSV and is never loaded to a heavy worksheet.
"""

from __future__ import annotations

from datetime import date, datetime, timedelta
from pathlib import Path
import random
from typing import Any, Iterable, Sequence

import xlsxwriter

from wfmhub.design import BODY_FONT, COLORS, TITLE_FONT
from wfmhub.excel_layout import (
    ACTION_FIRST_ROW,
    ACTION_HEADER_ROW,
    ACTION_SECTION_ROW,
    configure_v2_cell_canvas,
    make_v2_formats,
    style_v2_chart,
    write_v2_header,
    write_v2_kpis,
)


OUTPUT = (
    Path(__file__).resolve().parents[1]
    / "dist"
    / "PCS-PowerQuery-Slicer-Prototype-V2-Dummy.xlsx"
)

LOBS = ("RSA NL", "RSA FR", "RSA VL", "FORD NL", "FORD FR", "OEM FR")
TEAMS = {
    "RSA NL": ("Sophie Martin", "Karim Belkacem"),
    "RSA FR": ("Elena Rossi",),
    "RSA VL": ("Lucas Peeters",),
    "FORD NL": ("Daniel Weber",),
    "FORD FR": ("Laura Jensen",),
    "OEM FR": ("Miguel Santos",),
}
AGENTS = {
    "Sophie Martin": ("Amina El Idrissi", "Youssef Benali", "Nora Amrani"),
    "Karim Belkacem": ("Samir Haddad", "Leila Mansouri", "Mehdi Alaoui"),
    "Elena Rossi": ("Eva De Smet", "Mila Laurent", "Adam Lemaire"),
    "Lucas Peeters": ("Lina Dubois", "Noah Bernard", "Sarah Muller"),
    "Daniel Weber": ("Tom Janssen", "Ines Verhoeven", "Omar Diallo"),
    "Laura Jensen": ("Rayan Costa", "Sofia Martin", "Emma Leroy"),
    "Miguel Santos": ("Nina Bernard", "Yanis Robert", "Liam Petit"),
}
TEAM_INDEX = {
    leader: index
    for index, leader in enumerate(
        leader for lob in LOBS for leader in TEAMS[lob]
    )
}


def _agent_id(team_index: int, agent_index: int) -> str:
    return f"{team_index + 1:02d}{agent_index + 1:03d}"


def _dummy_call_rows() -> list[dict[str, Any]]:
    """Return deterministic, realistic inbound call-leg rows."""

    random.seed(26092)
    rows: list[dict[str, Any]] = []
    base = {
        "RSA NL": 4.25,
        "RSA FR": 4.05,
        "RSA VL": 3.92,
        "FORD NL": 4.31,
        "FORD FR": 4.12,
        "OEM FR": 4.38,
    }
    teams = [(lob, leader) for lob in LOBS for leader in TEAMS[lob]]
    call_number = 920000
    for month_start, day_count in ((date(2026, 8, 1), 31), (date(2026, 9, 1), 8)):
        for day_offset in range(day_count):
            business_date = month_start + timedelta(days=day_offset)
            if business_date.weekday() >= 5:
                continue
            for team_index, (lob, leader) in enumerate(teams):
                for agent_index, agent in enumerate(AGENTS[leader]):
                    for leg_index in range(random.randint(1, 3)):
                        call_number += 1
                        score = max(1, min(5, round(base[lob] + random.uniform(-2.0, 0.9))))
                        pcs_status = 1 if random.random() < 0.76 else 0
                        raw_nonblank = 1 if pcs_status and random.random() < 0.31 else 0
                        valid = 1 if raw_nonblank and random.random() < 0.88 else 0
                        q1 = score if valid else None
                        hour = 8 + ((team_index * 2 + agent_index + leg_index) % 10)
                        minute = (call_number * 7) % 60
                        call_start = datetime.combine(
                            business_date,
                            datetime.min.time(),
                        ).replace(hour=hour, minute=minute)
                        agent_id = _agent_id(TEAM_INDEX[leader], agent_index)
                        call_id = f"CBC-{business_date:%Y%m%d}-{call_number}"
                        coaching_key = f"PCS|{agent_id}|{business_date:%Y%m%d}|{call_id}"
                        rows.append({
                            "Date": business_date,
                            "Month": business_date.strftime("%Y-%m"),
                            "LOB": lob,
                            "Team Leader": leader,
                            "Agent": agent,
                            "Agent ID": agent_id,
                            "Language": "NL" if "NL" in lob or "VL" in lob else "FR",
                            "Call ID": call_id,
                            "Call Start": call_start,
                            "PCS Status 1": pcs_status,
                            "Q1 Nonblank": raw_nonblank,
                            "Valid Q1": valid,
                            "Q1 Score Sum": q1 or 0,
                            "Q1 Score": q1,
                            "Score <= 3": int(q1 is not None and q1 <= 3),
                            "Score > 3": int(q1 is not None and q1 > 3),
                            "Invalid Q1": int(bool(raw_nonblank and not valid)),
                            "Customer Comment": (
                                "Please clarify the next step."
                                if q1 is not None and q1 <= 3 else ""
                            ),
                            "Coaching Key": coaching_key,
                            "Priority": "HIGH" if q1 is not None and q1 <= 2 else "NORMAL",
                        })
    return rows


def _aggregate(
    rows: Iterable[dict[str, Any]],
    *,
    start: date,
    end: date,
    lob: str | None = None,
) -> dict[str, float | int | None]:
    selected = [
        row for row in rows
        if start <= row["Date"] <= end and (lob is None or row["LOB"] == lob)
    ]
    score_sum = sum(int(row["Q1 Score Sum"]) for row in selected)
    valid = sum(int(row["Valid Q1"]) for row in selected)
    nonblank = sum(int(row["Q1 Nonblank"]) for row in selected)
    eligible = sum(int(row["PCS Status 1"]) for row in selected)
    return {
        "score_sum": score_sum,
        "valid": valid,
        "nonblank": nonblank,
        "eligible": eligible,
        "pcs": score_sum / valid if valid else None,
        "participation": nonblank / eligible if eligible else None,
        "low": sum(int(row["Score <= 3"]) for row in selected),
        "calls": len(selected),
    }


def _add_native_table(
    workbook: xlsxwriter.Workbook,
    worksheet,
    *,
    first_row: int,
    first_col: int,
    name: str,
    headers: Sequence[str],
    rows: Sequence[Sequence[Any]],
    editable: set[str] | None = None,
) -> None:
    """Write one compact native Excel table with stable formats."""

    editable = editable or set()
    date_format = workbook.add_format({"num_format": "yyyy-mm-dd"})
    datetime_format = workbook.add_format({"num_format": "yyyy-mm-dd hh:mm"})
    decimal_format = workbook.add_format({"num_format": "0.00"})
    percent_format = workbook.add_format({"num_format": "0.0%"})
    text_format = workbook.add_format({"num_format": "@"})
    editable_text = workbook.add_format({
        "bg_color": COLORS["blue_light"],
        "font_color": COLORS["blue"],
    })
    editable_date = workbook.add_format({
        "bg_color": COLORS["blue_light"],
        "font_color": COLORS["blue"],
        "num_format": "yyyy-mm-dd",
    })
    for row_offset, values in enumerate(rows, 1):
        for column_offset, value in enumerate(values):
            header = headers[column_offset]
            cell_format = None
            if header in editable:
                cell_format = editable_date if "Date" in header else editable_text
            elif header in {"Date", "From", "To", "Call Date"}:
                cell_format = date_format
            elif header in {"Call Start"}:
                cell_format = datetime_format
            elif header in {"PCS", "Prior PCS", "Change", "Q1 Score"}:
                cell_format = decimal_format
            elif header in {"Participation"}:
                cell_format = percent_format
            elif header in {"Agent ID", "Call ID", "Coaching Key"}:
                cell_format = text_format
            worksheet.write(first_row + row_offset, first_col + column_offset, value, cell_format)
    worksheet.add_table(
        first_row,
        first_col,
        first_row + max(1, len(rows)),
        first_col + len(headers) - 1,
        {
            "name": name,
            "style": "Table Style Medium 2",
            "columns": [{"header": header} for header in headers],
        },
    )


def _write_sheet_title(
    workbook: xlsxwriter.Workbook,
    worksheet,
    *,
    title: str,
    subtitle: str,
    last_column: int,
    tab_color: str,
) -> None:
    title_format = workbook.add_format({
        "font_name": TITLE_FONT,
        "font_size": 20,
        "bold": True,
        "font_color": COLORS["white"],
        "bg_color": COLORS["navy"],
        "align": "left",
        "valign": "vcenter",
        "indent": 1,
    })
    subtitle_format = workbook.add_format({
        "font_name": BODY_FONT,
        "font_size": 10,
        "font_color": COLORS["white"],
        "bg_color": COLORS["teal"],
        "align": "left",
        "valign": "vcenter",
        "indent": 1,
    })
    worksheet.merge_range(0, 0, 0, last_column, title, title_format)
    worksheet.merge_range(1, 0, 1, last_column, subtitle, subtitle_format)
    worksheet.set_row_pixels(0, 52)
    worksheet.set_row_pixels(1, 30)
    worksheet.hide_gridlines(2)
    worksheet.set_zoom(85)
    worksheet.set_tab_color(tab_color)
    worksheet.set_footer("&LPrepared by Anass ASSRI | WFM&RPage &P of &N")


def _overview(
    workbook: xlsxwriter.Workbook,
    rows: Sequence[dict[str, Any]],
) -> None:
    worksheet = workbook.add_worksheet("OVERVIEW")
    formats = make_v2_formats(workbook)
    configure_v2_cell_canvas(worksheet, formats, zoom=82)
    worksheet.hide_row_col_headers()
    worksheet.set_tab_color(COLORS["gold"])
    worksheet.freeze_panes(2, 0)
    write_v2_header(
        worksheet,
        formats,
        "PCS PERFORMANCE & COACHING",
        status="DUMMY DATA",
        status_kind="PROVISIONAL",
    )

    info_label = workbook.add_format({
        "font_name": BODY_FONT,
        "font_size": 8,
        "bold": True,
        "font_color": COLORS["muted"],
        "bg_color": COLORS["canvas"],
        "align": "left",
        "valign": "vcenter",
    })
    info_value = workbook.add_format({
        "font_name": BODY_FONT,
        "font_size": 10,
        "bold": True,
        "font_color": COLORS["navy"],
        "bg_color": COLORS["white"],
        "align": "left",
        "valign": "vcenter",
        "indent": 1,
        "border": 1,
        "border_color": COLORS["line"],
    })
    info = (
        ("DATA THROUGH", "08 SEP 2026"),
        ("REFRESH", "POWER QUERY"),
        ("INTERACTION", "NATIVE SLICERS"),
        ("SCOPE", "ACTIVE FTE"),
    )
    for block, (label, value) in enumerate(info):
        first = block * 7
        worksheet.merge_range(1, first, 1, first + 1, label, info_label)
        worksheet.merge_range(1, first + 2, 1, first + 6, value, info_value)
    worksheet.set_row_pixels(1, 52)

    current_start = date(2026, 9, 1)
    current_end = date(2026, 9, 8)
    prior_start = date(2026, 8, 1)
    prior_end = date(2026, 8, 8)
    current = _aggregate(rows, start=current_start, end=current_end)
    prior = _aggregate(rows, start=prior_start, end=prior_end)
    change = (
        None
        if current["pcs"] is None or prior["pcs"] is None
        else float(current["pcs"]) - float(prior["pcs"])
    )
    write_v2_kpis(worksheet, formats, (
        ("CURRENT MTD PCS", current["pcs"], "decimal", current["pcs"]),
        ("PARTICIPATION", current["participation"], "percent", current["participation"]),
        ("PRIOR MTD PCS", prior["pcs"], "decimal", prior["pcs"]),
        ("CHANGE", change, "decimal", change),
    ))

    chart_sheet = workbook.add_worksheet("_CHARTS")
    chart_headers = (
        "LOB", "Current PCS", "Prior PCS", "Current Participation",
        "Date", "Daily PCS", "Daily Participation",
    )
    for column, header in enumerate(chart_headers):
        chart_sheet.write(0, column, header)
    lob_summary: list[tuple[Any, ...]] = []
    for row_index, lob in enumerate(LOBS, 1):
        current_lob = _aggregate(rows, start=current_start, end=current_end, lob=lob)
        prior_lob = _aggregate(rows, start=prior_start, end=prior_end, lob=lob)
        chart_sheet.write(row_index, 0, lob)
        chart_sheet.write(row_index, 1, current_lob["pcs"])
        chart_sheet.write(row_index, 2, prior_lob["pcs"])
        chart_sheet.write(row_index, 3, current_lob["participation"])
        lob_summary.append((
            lob,
            current_lob["valid"],
            current_lob["pcs"],
            prior_lob["pcs"],
            None if current_lob["pcs"] is None or prior_lob["pcs"] is None
            else float(current_lob["pcs"]) - float(prior_lob["pcs"]),
            current_lob["participation"],
            current_lob["low"],
        ))
    for day_offset in range(8):
        business_date = current_start + timedelta(days=day_offset)
        metric = _aggregate(rows, start=business_date, end=business_date)
        chart_sheet.write_datetime(
            day_offset + 1,
            4,
            datetime.combine(business_date, datetime.min.time()),
            workbook.add_format({"num_format": "d-mmm"}),
        )
        chart_sheet.write(day_offset + 1, 5, metric["pcs"])
        chart_sheet.write(day_offset + 1, 6, metric["participation"])
    chart_sheet.hide()

    lob_chart = workbook.add_chart({"type": "bar"})
    lob_chart.add_series({
        "name": "Current MTD",
        "categories": "='_CHARTS'!$A$2:$A$7",
        "values": "='_CHARTS'!$B$2:$B$7",
        "fill": {"color": COLORS["teal"]},
        "border": {"none": True},
        "data_labels": {"value": True, "num_format": "0.00"},
    })
    lob_chart.add_series({
        "name": "Prior comparable",
        "categories": "='_CHARTS'!$A$2:$A$7",
        "values": "='_CHARTS'!$C$2:$C$7",
        "fill": {"color": COLORS["muted"]},
        "border": {"none": True},
    })
    style_v2_chart(lob_chart, title="PCS BY LOB", kind="bar")
    lob_chart.set_x_axis({"min": 1, "max": 5, "major_unit": 1})
    lob_chart.set_y_axis({"reverse": True})

    trend_chart = workbook.add_chart({"type": "line"})
    trend_chart.add_series({
        "name": "Daily PCS",
        "categories": "='_CHARTS'!$E$2:$E$9",
        "values": "='_CHARTS'!$F$2:$F$9",
        "line": {"color": COLORS["teal"], "width": 2.25},
        "marker": {
            "type": "circle",
            "size": 5,
            "fill": {"color": COLORS["teal"]},
            "border": {"color": COLORS["teal"]},
        },
    })
    style_v2_chart(trend_chart, title="CURRENT MONTH PCS TREND")
    trend_chart.set_y_axis({"min": 1, "max": 5, "major_unit": 1})
    trend_chart.set_x_axis({"date_axis": True, "num_format": "d-mmm"})

    worksheet.insert_chart(7, 0, lob_chart, {"x_scale": 1.0, "y_scale": 1.0})
    worksheet.insert_chart(7, 14, trend_chart, {"x_scale": 1.0, "y_scale": 1.0})
    worksheet.merge_range(
        ACTION_SECTION_ROW,
        0,
        ACTION_SECTION_ROW,
        27,
        "LOB PERFORMANCE · FINAL VERSION WILL BE DRIVEN BY NATIVE PIVOT SLICERS",
        formats.section,
    )
    headers = (
        "LOB", "VALID Q1", "CURRENT PCS", "PRIOR PCS",
        "CHANGE", "PARTICIPATION", "COACHING DUE",
    )
    for column, header in enumerate(headers):
        first = column * 4
        worksheet.merge_range(
            ACTION_HEADER_ROW,
            first,
            ACTION_HEADER_ROW,
            first + 3,
            header,
            formats.table_header,
        )
    row_formats = (
        formats.table_text,
        formats.table_integer,
        formats.table_decimal,
        formats.table_decimal,
        formats.table_decimal,
        formats.table_percent,
        formats.table_integer,
    )
    for row_offset, values in enumerate(lob_summary):
        row_index = ACTION_FIRST_ROW + row_offset
        for column, (value, cell_format) in enumerate(zip(values, row_formats)):
            first = column * 4
            worksheet.merge_range(row_index, first, row_index, first + 3, value, cell_format)
        worksheet.set_row_pixels(row_index, 26)

    agent_section_row = ACTION_FIRST_ROW + len(lob_summary) + 2
    agent_header_row = agent_section_row + 1
    worksheet.merge_range(
        agent_section_row,
        0,
        agent_section_row,
        27,
        "AGENT PERFORMANCE · FILTER THE PERFORMANCE TABLE WITH NATIVE SLICERS",
        formats.section,
    )
    agent_headers = (
        ("AGENT", 0, 4),
        ("AGENT ID", 5, 7),
        ("LOB", 8, 10),
        ("TEAM LEADER", 11, 15),
        ("CURRENT PCS", 16, 18),
        ("PRIOR PCS", 19, 21),
        ("CHANGE", 22, 24),
        ("PARTICIPATION", 25, 27),
    )
    for header, first, last in agent_headers:
        worksheet.merge_range(
            agent_header_row,
            first,
            agent_header_row,
            last,
            header,
            formats.table_header,
        )
    agent_summary: list[tuple[Any, ...]] = []
    for lob in LOBS:
        for leader in TEAMS[lob]:
            for agent_index, agent in enumerate(AGENTS[leader]):
                agent_rows = [row for row in rows if row["Agent"] == agent]
                current_agent = _aggregate(
                    agent_rows,
                    start=current_start,
                    end=current_end,
                )
                prior_agent = _aggregate(
                    agent_rows,
                    start=prior_start,
                    end=prior_end,
                )
                change_value = (
                    None
                    if current_agent["pcs"] is None or prior_agent["pcs"] is None
                    else float(current_agent["pcs"]) - float(prior_agent["pcs"])
                )
                agent_summary.append((
                    agent,
                    _agent_id(TEAM_INDEX[leader], agent_index),
                    lob,
                    leader,
                    current_agent["pcs"],
                    prior_agent["pcs"],
                    change_value,
                    current_agent["participation"],
                ))
    agent_spans = tuple((first, last) for _header, first, last in agent_headers)
    agent_formats = (
        formats.table_text,
        formats.table_text,
        formats.table_text,
        formats.table_text,
        formats.table_decimal,
        formats.table_decimal,
        formats.table_decimal,
        formats.table_percent,
    )
    for row_offset, values in enumerate(agent_summary):
        row_index = agent_header_row + 1 + row_offset
        for (first, last), value, cell_format in zip(
            agent_spans,
            values,
            agent_formats,
        ):
            worksheet.merge_range(
                row_index,
                first,
                row_index,
                last,
                value,
                cell_format,
            )
        worksheet.set_row_pixels(row_index, 25)
    worksheet.set_footer("&LPrepared by Anass ASSRI | WFM&CPCS prototype&RPage &P of &N")


def _performance(
    workbook: xlsxwriter.Workbook,
    rows: Sequence[dict[str, Any]],
) -> None:
    worksheet = workbook.add_worksheet("PERFORMANCE")
    headers = (
        "Period", "From", "To", "LOB", "Team Leader", "Agent", "Agent ID",
        "Q1 Score Sum", "Valid Q1", "PCS", "Q1 Nonblank", "PCS Status 1",
        "Participation", "Score <= 3", "Score > 3", "Inbound Legs",
    )
    _write_sheet_title(
        workbook,
        worksheet,
        title="PCS PERFORMANCE",
        subtitle="LIGHTWEIGHT POWER QUERY OUTPUT · one row per period and agent · native filters and slicers only",
        last_column=len(headers) - 1,
        tab_color=COLORS["teal"],
    )
    periods = (
        ("Current MTD", date(2026, 9, 1), date(2026, 9, 8)),
        ("Previous MTD same days", date(2026, 8, 1), date(2026, 8, 8)),
        ("Previous full month", date(2026, 8, 1), date(2026, 8, 31)),
    )
    table_rows: list[tuple[Any, ...]] = []
    for period, start, end in periods:
        for lob in LOBS:
            for leader in TEAMS[lob]:
                for agent_index, agent in enumerate(AGENTS[leader]):
                    agent_rows = [
                        item for item in rows
                        if start <= item["Date"] <= end
                        and item["LOB"] == lob
                        and item["Team Leader"] == leader
                        and item["Agent"] == agent
                    ]
                    metric = _aggregate(agent_rows, start=start, end=end)
                    table_rows.append((
                        period,
                        start,
                        end,
                        lob,
                        leader,
                        agent,
                        _agent_id(TEAM_INDEX[leader], agent_index),
                        metric["score_sum"],
                        metric["valid"],
                        metric["pcs"],
                        metric["nonblank"],
                        metric["eligible"],
                        metric["participation"],
                        metric["low"],
                        int(metric["valid"]) - int(metric["low"]),
                        metric["calls"],
                    ))
    _add_native_table(
        workbook,
        worksheet,
        first_row=3,
        first_col=0,
        name="tblPCSPerformance",
        headers=headers,
        rows=table_rows,
    )
    worksheet.freeze_panes(4, 0)
    widths = (24, 12, 12, 14, 22, 24, 13, 15, 12, 11, 14, 14, 15, 13, 12, 14)
    for column, width in enumerate(widths):
        worksheet.set_column(column, column, width)


def _coaching(
    workbook: xlsxwriter.Workbook,
    rows: Sequence[dict[str, Any]],
) -> None:
    worksheet = workbook.add_worksheet("COACHING")
    queue_headers = (
        "Call Date", "LOB", "Team Leader", "Agent", "Agent ID", "Call ID",
        "Call Start", "Q1 Score", "Priority", "Customer Comment", "Coaching Key",
    )
    _write_sheet_title(
        workbook,
        worksheet,
        title="PCS COACHING",
        subtitle="ONE COMPACT WORKSPACE · filter the queue with native slicers, then record the decision in the blue action log below",
        last_column=len(queue_headers) - 1,
        tab_color=COLORS["gold"],
    )
    low_rows = [
        row for row in rows
        if row["Date"] >= date(2026, 9, 1) and row["Score <= 3"] == 1
    ][:12]
    queue_rows = [
        (
            row["Date"], row["LOB"], row["Team Leader"], row["Agent"],
            row["Agent ID"], row["Call ID"], row["Call Start"], row["Q1 Score"],
            row["Priority"], row["Customer Comment"], row["Coaching Key"],
        )
        for row in low_rows
    ]
    _add_native_table(
        workbook,
        worksheet,
        first_row=3,
        first_col=0,
        name="tblCoachingQueue",
        headers=queue_headers,
        rows=queue_rows,
    )
    section_row = 4 + len(queue_rows) + 2
    section_format = workbook.add_format({
        "font_name": TITLE_FONT,
        "font_size": 12,
        "bold": True,
        "font_color": COLORS["white"],
        "bg_color": COLORS["navy"],
        "align": "left",
        "valign": "vcenter",
        "indent": 1,
    })
    worksheet.merge_range(
        section_row,
        0,
        section_row,
        10,
        "COACHING ACTION LOG · PERMANENT · BLUE CELLS ARE EDITABLE",
        section_format,
    )
    action_headers = (
        "Coaching Key", "Call ID", "Coaching Status", "Coach",
        "Coaching Date", "Due Date", "Coaching Comment",
    )
    action_rows = [
        (
            low_rows[index]["Coaching Key"],
            low_rows[index]["Call ID"],
            status,
            coach,
            coaching_date,
            due_date,
            comment,
        )
        for index, (status, coach, coaching_date, due_date, comment) in enumerate((
            ("Completed", "Sophie Martin", date(2026, 9, 7), date(2026, 9, 8), "Call reviewed with agent"),
            ("Planned", "Karim Belkacem", None, date(2026, 9, 10), "Session booked"),
            ("Pending", "Elena Rossi", None, date(2026, 9, 11), ""),
            ("Pending", "Daniel Weber", None, date(2026, 9, 12), ""),
        ))
        if index < len(low_rows)
    ]
    action_rows.extend(
        tuple(None for _ in action_headers)
        for _ in range(max(0, 8 - len(action_rows)))
    )
    _add_native_table(
        workbook,
        worksheet,
        first_row=section_row + 1,
        first_col=0,
        name="tblCoachingActions",
        headers=action_headers,
        rows=action_rows,
        editable={
            "Coaching Key", "Call ID", "Coaching Status", "Coach",
            "Coaching Date", "Due Date", "Coaching Comment",
        },
    )
    action_first_data = section_row + 3
    worksheet.data_validation(
        action_first_data - 1,
        2,
        action_first_data + 499,
        2,
        {
            "validate": "list",
            "source": ["Pending", "Planned", "Completed", "Not required"],
        },
    )
    worksheet.conditional_format(
        action_first_data - 1,
        0,
        action_first_data + 499,
        0,
        {"type": "duplicate", "format": workbook.add_format({"bg_color": COLORS["red_light"]})},
    )
    worksheet.freeze_panes(4, 0)
    widths = (13, 14, 22, 24, 13, 25, 19, 11, 12, 34, 48)
    for column, width in enumerate(widths):
        worksheet.set_column(column, column, width)


def _help(workbook: xlsxwriter.Workbook) -> None:
    worksheet = workbook.add_worksheet("HOW IT WORKS")
    _write_sheet_title(
        workbook,
        worksheet,
        title="PCS — SIMPLE OPERATING MODEL",
        subtitle="DUMMY FILE FOR VALIDATION · no query, pivot or slicer has been installed in this prototype",
        last_column=7,
        tab_color=COLORS["gold"],
    )
    section = workbook.add_format({
        "font_name": TITLE_FONT,
        "font_size": 12,
        "bold": True,
        "font_color": COLORS["white"],
        "bg_color": COLORS["navy"],
        "align": "left",
        "valign": "vcenter",
        "indent": 1,
    })
    body = workbook.add_format({
        "font_name": BODY_FONT,
        "font_size": 11,
        "font_color": COLORS["navy"],
        "bg_color": COLORS["white"],
        "text_wrap": True,
        "valign": "top",
        "border": 1,
        "border_color": COLORS["line"],
    })
    number = workbook.add_format({
        "font_name": TITLE_FONT,
        "font_size": 18,
        "bold": True,
        "font_color": COLORS["teal"],
        "bg_color": COLORS["white"],
        "align": "center",
        "valign": "vcenter",
        "border": 1,
        "border_color": COLORS["line"],
    })
    worksheet.merge_range(3, 0, 3, 7, "THE FINAL REFRESH PIPELINE", section)
    steps = (
        ("1", "Run PCS Update in WFMHub", "The Hub refreshes SQLite and atomically replaces five lightweight PCS CSV feeds outside the workbook."),
        ("2", "Open the permanent PCS workbook", "The shared workbook stays in place; coaching history is not rebuilt."),
        ("3", "Choose Data > Refresh All", "Power Query reads the CSV directly and refreshes the lightweight performance and coaching outputs."),
        ("4", "Use native Excel slicers", "Period, LOB, Team Leader and Agent slicers control PivotTables and PivotCharts."),
        ("5", "Quality completes coaching", "Coaching decisions stay in the separate blue action table and survive refreshes."),
    )
    for offset, (step, title, detail) in enumerate(steps, 4):
        worksheet.write(offset, 0, step, number)
        worksheet.merge_range(offset, 1, offset, 2, title, body)
        worksheet.merge_range(offset, 3, offset, 7, detail, body)
        worksheet.set_row_pixels(offset, 52)
    worksheet.merge_range(10, 0, 10, 7, "WHAT IS DELIBERATELY REMOVED", section)
    removals = (
        "No fake filter cells or cascading dropdowns",
        "No AGGREGATE / INDEX-generated result lists",
        "No dynamic arrays and no hidden formula ranking",
        "No raw call-leg worksheet inside the permanent workbook",
        "No Hub attempt to overwrite an open shared workbook",
        "No Power Query load into the editable coaching-action table",
    )
    for offset, item in enumerate(removals, 11):
        worksheet.merge_range(offset, 0, offset, 7, f"• {item}", body)
        worksheet.set_row_pixels(offset, 34)
    worksheet.merge_range(17, 0, 17, 7, "PRODUCTION TABLE OWNERSHIP", section)
    ownership = (
        ("Five PCS CSV feeds", "WFMHub", "Governed scorecards and exact coaching calls; no raw call-leg worksheet"),
        ("tblPcsPerformance", "Power Query", "Period/LOB/team/agent results used by native filters and slicers"),
        ("tblCoachingQueue", "Power Query", "Exact low-score calls with Call ID and Coaching Key"),
        ("tblCoachingActions", "People", "Permanent coaching status, owner, dates and comment"),
    )
    for offset, (table, owner, purpose) in enumerate(ownership, 18):
        worksheet.merge_range(offset, 0, offset, 1, table, body)
        worksheet.merge_range(offset, 2, offset, 3, owner, body)
        worksheet.merge_range(offset, 4, offset, 7, purpose, body)
        worksheet.set_row_pixels(offset, 42)
    worksheet.set_column_pixels(0, 0, 60)
    worksheet.set_column_pixels(1, 2, 140)
    worksheet.set_column_pixels(3, 7, 155)


def build(output: Path = OUTPUT) -> Path:
    output.parent.mkdir(parents=True, exist_ok=True)
    rows = _dummy_call_rows()
    workbook = xlsxwriter.Workbook(output)
    workbook.set_properties({
        "title": "PCS Power Query and Slicer Prototype — Dummy Data",
        "subject": "Visual and workflow acceptance prototype",
        "author": "Anass ASSRI",
        "company": "WFM",
        "comments": "Dummy data only. Not a production report.",
    })
    workbook.set_custom_property("WFMHub Prototype", "PCS PQ SLICER DUMMY")
    workbook.set_calc_mode("auto")
    _overview(workbook, rows)
    _performance(workbook, rows)
    _coaching(workbook, rows)
    _help(workbook)
    workbook.close()
    return output


if __name__ == "__main__":
    print(build())
