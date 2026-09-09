#!/usr/bin/env python3
"""Build the interactive PCS dropdown-filter acceptance prototype.

The workbook keeps the approved compact Overview. Dummy results are calculated
in Python for every valid Period/LOB/Team/Agent selection. Excel only performs
exact INDEX/MATCH lookups, so the prototype has no dynamic arrays, aggregate
ranking, PivotCache, slicer XML, raw call-leg worksheet, or external link.
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime, timedelta
from pathlib import Path
import re
from typing import Any, Iterable, Sequence

import xlsxwriter
from xlsxwriter.utility import xl_col_to_name

from build_pcs_pq_slicer_prototype import (
    AGENTS,
    LOBS,
    TEAMS,
    TEAM_INDEX,
    _add_native_table,
    _aggregate,
    _agent_id,
    _dummy_call_rows,
    _write_sheet_title,
)
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
    / "PCS-Overview-Dropdown-Filters-Prototype-Dummy.xlsx"
)
ALL = "All"


@dataclass(frozen=True)
class Period:
    label: str
    start: date
    end: date
    prior_start: date
    prior_end: date


PERIODS = (
    Period(
        "Latest day",
        date(2026, 9, 8), date(2026, 9, 8),
        date(2026, 9, 7), date(2026, 9, 7),
    ),
    Period(
        "Current week",
        date(2026, 9, 7), date(2026, 9, 8),
        date(2026, 8, 31), date(2026, 9, 1),
    ),
    Period(
        "Current MTD",
        date(2026, 9, 1), date(2026, 9, 8),
        date(2026, 8, 1), date(2026, 8, 8),
    ),
    Period(
        "Previous MTD same days",
        date(2026, 8, 1), date(2026, 8, 8),
        date(2026, 7, 1), date(2026, 7, 8),
    ),
    Period(
        "Previous full month",
        date(2026, 8, 1), date(2026, 8, 31),
        date(2026, 7, 1), date(2026, 7, 31),
    ),
)

KPI_HEADERS = ("Selection Key", "PCS", "Participation", "Prior PCS", "Change")
LOB_HEADERS = (
    "View Key", "LOB", "Valid Q1", "PCS", "Prior PCS", "Change",
    "Participation", "Coaching Due",
)
AGENT_HEADERS = (
    "View Key", "Agent", "Agent ID", "LOB", "Team Leader", "PCS",
    "Prior PCS", "Change", "Participation",
)
DAILY_HEADERS = ("View Key", "Date", "PCS", "Participation")
COACH_HEADERS = (
    "View Key", "Date", "LOB", "Team Leader", "Agent", "Agent ID",
    "Q1 Score", "Priority", "Call ID", "Customer Comment", "Coaching Key",
)


def _slug(value: str) -> str:
    cleaned = re.sub(r"[^A-Za-z0-9_]", "_", value.strip())
    return re.sub(r"_+", "_", cleaned).strip("_") or "ALL"


def _excel_serial(value: Any) -> Any:
    if isinstance(value, datetime):
        days = (value.date() - date(1899, 12, 30)).days
        seconds = value.hour * 3600 + value.minute * 60 + value.second
        return days + seconds / 86_400
    if isinstance(value, date):
        return (value - date(1899, 12, 30)).days
    return "" if value is None else value


def _matching(
    rows: Iterable[dict[str, Any]],
    *,
    start: date,
    end: date,
    lob: str = ALL,
    leader: str = ALL,
    agent: str = ALL,
) -> list[dict[str, Any]]:
    return [
        row for row in rows
        if start <= row["Date"] <= end
        and (lob == ALL or row["LOB"] == lob)
        and (leader == ALL or row["Team Leader"] == leader)
        and (agent == ALL or row["Agent"] == agent)
    ]


def _metric(rows: Sequence[dict[str, Any]], start: date, end: date) -> dict[str, Any]:
    return _aggregate(rows, start=start, end=end)


def _change(current: dict[str, Any], prior: dict[str, Any]) -> float | None:
    if current["pcs"] is None or prior["pcs"] is None:
        return None
    return float(current["pcs"]) - float(prior["pcs"])


def _selection_key(period: str, lob: str, leader: str, agent: str) -> str:
    return "|".join((period, lob, leader, agent))


def _selection_options() -> Iterable[tuple[str, str, str]]:
    all_leaders = tuple(leader for lob in LOBS for leader in TEAMS[lob])
    all_agents = tuple(agent for leader in all_leaders for agent in AGENTS[leader])
    for lob in (ALL, *LOBS):
        leaders = all_leaders if lob == ALL else TEAMS[lob]
        for leader in (ALL, *leaders):
            if leader == ALL:
                agents = all_agents if lob == ALL else tuple(
                    agent for item in TEAMS[lob] for agent in AGENTS[item]
                )
            else:
                agents = AGENTS[leader]
            for agent in (ALL, *agents):
                yield lob, leader, agent


def _build_filter_cache(
    rows: Sequence[dict[str, Any]],
) -> tuple[
    list[tuple[Any, ...]],
    list[tuple[Any, ...]],
    list[tuple[Any, ...]],
    list[tuple[Any, ...]],
]:
    kpis: list[tuple[Any, ...]] = []
    lob_view: list[tuple[Any, ...]] = []
    agent_view: list[tuple[Any, ...]] = []
    daily_view: list[tuple[Any, ...]] = []
    for period in PERIODS:
        for lob, leader, agent in _selection_options():
            key = _selection_key(period.label, lob, leader, agent)
            current_rows = _matching(
                rows, start=period.start, end=period.end,
                lob=lob, leader=leader, agent=agent,
            )
            prior_rows = _matching(
                rows, start=period.prior_start, end=period.prior_end,
                lob=lob, leader=leader, agent=agent,
            )
            current = _metric(current_rows, period.start, period.end)
            prior = _metric(prior_rows, period.prior_start, period.prior_end)
            kpis.append((
                key, current["pcs"], current["participation"], prior["pcs"],
                _change(current, prior),
            ))

            visible_lobs = [
                item for item in LOBS
                if any(row["LOB"] == item for row in current_rows)
            ]
            for rank, item in enumerate(visible_lobs, 1):
                current_lob_rows = [row for row in current_rows if row["LOB"] == item]
                prior_lob_rows = [row for row in prior_rows if row["LOB"] == item]
                current_lob = _metric(current_lob_rows, period.start, period.end)
                prior_lob = _metric(prior_lob_rows, period.prior_start, period.prior_end)
                lob_view.append((
                    f"{key}|{rank}", item, current_lob["valid"],
                    current_lob["pcs"], prior_lob["pcs"],
                    _change(current_lob, prior_lob),
                    current_lob["participation"], current_lob["low"],
                ))

            visible_agents = sorted(
                {row["Agent"] for row in current_rows},
                key=str.casefold,
            )
            for rank, item in enumerate(visible_agents, 1):
                source_row = next(row for row in current_rows if row["Agent"] == item)
                current_agent_rows = [row for row in current_rows if row["Agent"] == item]
                prior_agent_rows = [row for row in prior_rows if row["Agent"] == item]
                current_agent = _metric(current_agent_rows, period.start, period.end)
                prior_agent = _metric(
                    prior_agent_rows, period.prior_start, period.prior_end,
                )
                agent_view.append((
                    f"{key}|{rank}", item, source_row["Agent ID"],
                    source_row["LOB"], source_row["Team Leader"],
                    current_agent["pcs"], prior_agent["pcs"],
                    _change(current_agent, prior_agent),
                    current_agent["participation"],
                ))

            for rank, offset in enumerate(
                range((period.end - period.start).days + 1), 1,
            ):
                business_date = period.start + timedelta(days=offset)
                day_rows = [
                    row for row in current_rows if row["Date"] == business_date
                ]
                day = _metric(day_rows, business_date, business_date)
                daily_view.append((
                    f"{key}|{rank}", business_date,
                    day["pcs"], day["participation"],
                ))
    return kpis, lob_view, agent_view, daily_view


def _build_coaching_cache(
    rows: Sequence[dict[str, Any]],
) -> list[tuple[Any, ...]]:
    output: list[tuple[Any, ...]] = []
    for period in PERIODS:
        for lob in (ALL, *LOBS):
            matches = [
                row for row in _matching(
                    rows, start=period.start, end=period.end, lob=lob,
                )
                if row["Score <= 3"] == 1
            ]
            matches.sort(
                key=lambda row: (
                    row["Date"], row["LOB"], row["Team Leader"],
                    row["Agent"], row["Call Start"],
                ),
                reverse=True,
            )
            for rank, row in enumerate(matches, 1):
                output.append((
                    f"{period.label}|{lob}|{rank}",
                    row["Date"], row["LOB"], row["Team Leader"],
                    row["Agent"], row["Agent ID"], row["Q1 Score"],
                    row["Priority"], row["Call ID"],
                    row["Customer Comment"], row["Coaching Key"],
                ))
    return output


def _lookup_formula(
    sheet: str,
    value_column: int,
    key_expression: str,
    last_row: int,
) -> str:
    value_letter = xl_col_to_name(value_column)
    return (
        f'=IFERROR(INDEX(\'{sheet}\'!${value_letter}$2:${value_letter}${last_row},'
        f'MATCH({key_expression},\'{sheet}\'!$A$2:$A${last_row},0)),"")'
    )


def _write_merged_formula(
    worksheet,
    row: int,
    first: int,
    last: int,
    formula: str,
    cell_format,
    cached: Any,
) -> None:
    worksheet.merge_range(row, first, row, last, "", cell_format)
    worksheet.write_formula(
        row, first, formula, cell_format, _excel_serial(cached),
    )


def _cache_map(rows: Sequence[Sequence[Any]]) -> dict[str, tuple[Any, ...]]:
    return {str(row[0]): tuple(row) for row in rows}


def _overview(
    workbook: xlsxwriter.Workbook,
    kpis: Sequence[Sequence[Any]],
    lob_rows: Sequence[Sequence[Any]],
    agent_rows: Sequence[Sequence[Any]],
    daily_rows: Sequence[Sequence[Any]],
) -> None:
    worksheet = workbook.add_worksheet("OVERVIEW")
    formats = make_v2_formats(workbook)
    configure_v2_cell_canvas(worksheet, formats, zoom=82)
    worksheet.hide_row_col_headers()
    worksheet.set_tab_color(COLORS["gold"])
    worksheet.freeze_panes(2, 0)
    write_v2_header(
        worksheet, formats, "PCS PERFORMANCE & COACHING",
        status="FILTER PROTOTYPE", status_kind="PROVISIONAL",
    )

    defaults = ("Current MTD", ALL, ALL, ALL)
    filter_items = (
        ("PERIOD", defaults[0], "=PCS_PERIOD_LIST"),
        ("LOB", defaults[1], "=PCS_LOB_LIST"),
        (
            "TEAM LEADER", defaults[2],
            '=INDIRECT("PCS_TL_"&SUBSTITUTE($J$2," ","_"))',
        ),
        (
            "AGENT", defaults[3],
            '=INDIRECT("PCS_AGENT_"&SUBSTITUTE($J$2," ","_")&"_"&SUBSTITUTE($Q$2," ","_"))',
        ),
    )
    for block, (label, value, source) in enumerate(filter_items):
        first = block * 7
        worksheet.merge_range(1, first, 1, first + 1, label, formats.filter_label)
        worksheet.merge_range(1, first + 2, 1, first + 6, value, formats.filter_value)
        worksheet.data_validation(1, first + 2, 1, first + 2, {
            "validate": "list",
            "source": source,
            "input_title": label,
            "input_message": "Choose from the governed list",
            "error_title": "Invalid filter",
            "error_message": "Choose a value from the dropdown list.",
        })
    worksheet.set_row_pixels(1, 52)

    selected_key = '$C$2&"|"&$J$2&"|"&$Q$2&"|"&$X$2'
    default_key = _selection_key(*defaults)
    kpi_map = _cache_map(kpis)
    default_kpi = kpi_map[default_key]
    kpi_last = len(kpis) + 1
    card_specs = (
        ("SELECTED PCS", 1, "decimal"),
        ("PARTICIPATION", 2, "percent"),
        ("PRIOR COMPARABLE", 3, "decimal"),
        ("CHANGE", 4, "decimal"),
    )
    write_v2_kpis(worksheet, formats, tuple(
        (
            label,
            _lookup_formula("_KPI_CACHE", column, selected_key, kpi_last),
            kind,
            default_kpi[column],
        )
        for label, column, kind in card_specs
    ))

    lob_map = _cache_map(lob_rows)
    daily_map = _cache_map(daily_rows)
    chart_sheet = workbook.add_worksheet("_CHARTS")
    chart_sheet.hide_gridlines(2)
    for column, header in enumerate((
        "LOB", "PCS", "Prior PCS", "Date", "Daily PCS", "Daily Participation",
    )):
        chart_sheet.write(0, column, header)
    lob_last = len(lob_rows) + 1
    daily_last = len(daily_rows) + 1
    for offset in range(len(LOBS)):
        rank = offset + 1
        key_expression = f'{selected_key}&"|{rank}"'
        cached = lob_map.get(f"{default_key}|{rank}", tuple(None for _ in LOB_HEADERS))
        for column, cache_column in enumerate((1, 3, 4)):
            chart_sheet.write_formula(
                offset + 1,
                column,
                _lookup_formula(
                    "_LOB_CACHE", cache_column, key_expression, lob_last,
                ),
                None,
                _excel_serial(cached[cache_column]),
            )
    for offset in range(31):
        rank = offset + 1
        key_expression = f'{selected_key}&"|{rank}"'
        cached = daily_map.get(
            f"{default_key}|{rank}", tuple(None for _ in DAILY_HEADERS),
        )
        for output_column, cache_column in ((3, 1), (4, 2), (5, 3)):
            cell_format = None
            if output_column == 3:
                cell_format = workbook.add_format({"num_format": "d-mmm"})
            chart_sheet.write_formula(
                offset + 1,
                output_column,
                _lookup_formula(
                    "_DAILY_CACHE", cache_column, key_expression, daily_last,
                ),
                cell_format,
                _excel_serial(cached[cache_column]),
            )
    chart_sheet.hide()

    lob_chart = workbook.add_chart({"type": "bar"})
    lob_chart.add_series({
        "name": "Selected",
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
    style_v2_chart(lob_chart, title="PCS BY LOB · SELECTED SCOPE", kind="bar")
    lob_chart.set_x_axis({"min": 1, "max": 5, "major_unit": 1})
    lob_chart.set_y_axis({"reverse": True})

    trend_chart = workbook.add_chart({"type": "line"})
    trend_chart.add_series({
        "name": "Daily PCS",
        "categories": "='_CHARTS'!$D$2:$D$32",
        "values": "='_CHARTS'!$E$2:$E$32",
        "line": {"color": COLORS["teal"], "width": 2.25},
        "marker": {
            "type": "circle", "size": 5,
            "fill": {"color": COLORS["teal"]},
            "border": {"color": COLORS["teal"]},
        },
    })
    style_v2_chart(trend_chart, title="DAILY PCS · SELECTED PERIOD")
    trend_chart.set_y_axis({"min": 1, "max": 5, "major_unit": 1})
    trend_chart.set_x_axis({"date_axis": True, "num_format": "d-mmm"})
    worksheet.insert_chart(7, 0, lob_chart)
    worksheet.insert_chart(7, 14, trend_chart)

    worksheet.merge_range(
        ACTION_SECTION_ROW, 0, ACTION_SECTION_ROW, 27,
        "LOB PERFORMANCE · RESPONDS TO ALL FOUR DROPDOWNS", formats.section,
    )
    lob_headers = (
        "LOB", "VALID Q1", "SELECTED PCS", "PRIOR PCS",
        "CHANGE", "PARTICIPATION", "COACHING DUE",
    )
    for column, header in enumerate(lob_headers):
        first = column * 4
        worksheet.merge_range(
            ACTION_HEADER_ROW, first, ACTION_HEADER_ROW, first + 3,
            header, formats.table_header,
        )
    lob_formats = (
        formats.table_text, formats.table_integer, formats.table_decimal,
        formats.table_decimal, formats.table_decimal, formats.table_percent,
        formats.table_integer,
    )
    for offset in range(len(LOBS)):
        rank = offset + 1
        row_index = ACTION_FIRST_ROW + offset
        key_expression = f'{selected_key}&"|{rank}"'
        cached = lob_map.get(f"{default_key}|{rank}", tuple(None for _ in LOB_HEADERS))
        for column, (cell_format, cache_column) in enumerate(
            zip(lob_formats, range(1, len(LOB_HEADERS))),
        ):
            first = column * 4
            _write_merged_formula(
                worksheet, row_index, first, first + 3,
                _lookup_formula(
                    "_LOB_CACHE", cache_column, key_expression, lob_last,
                ),
                cell_format, cached[cache_column],
            )
        worksheet.set_row_pixels(row_index, 26)

    agent_section_row = ACTION_FIRST_ROW + len(LOBS) + 2
    agent_header_row = agent_section_row + 1
    worksheet.merge_range(
        agent_section_row, 0, agent_section_row, 27,
        "AGENT PERFORMANCE · CASCADING LOB → TEAM LEADER → AGENT",
        formats.section,
    )
    agent_headers = (
        ("AGENT", 0, 4), ("AGENT ID", 5, 7), ("LOB", 8, 10),
        ("TEAM LEADER", 11, 15), ("SELECTED PCS", 16, 18),
        ("PRIOR PCS", 19, 21), ("CHANGE", 22, 24),
        ("PARTICIPATION", 25, 27),
    )
    for header, first, last in agent_headers:
        worksheet.merge_range(
            agent_header_row, first, agent_header_row, last,
            header, formats.table_header,
        )
    agent_formats = (
        formats.table_text, formats.table_text, formats.table_text,
        formats.table_text, formats.table_decimal, formats.table_decimal,
        formats.table_decimal, formats.table_percent,
    )
    agent_map = _cache_map(agent_rows)
    agent_last = len(agent_rows) + 1
    agent_capacity = sum(len(AGENTS[leader]) for lob in LOBS for leader in TEAMS[lob])
    for offset in range(agent_capacity):
        rank = offset + 1
        row_index = agent_header_row + 1 + offset
        key_expression = f'{selected_key}&"|{rank}"'
        cached = agent_map.get(
            f"{default_key}|{rank}", tuple(None for _ in AGENT_HEADERS),
        )
        for (header, first, last), cell_format, cache_column in zip(
            agent_headers, agent_formats, range(1, len(AGENT_HEADERS)),
        ):
            del header
            _write_merged_formula(
                worksheet, row_index, first, last,
                _lookup_formula(
                    "_AGENT_CACHE", cache_column, key_expression, agent_last,
                ),
                cell_format, cached[cache_column],
            )
        worksheet.set_row_pixels(row_index, 25)
    worksheet.set_footer(
        "&LPrepared by Anass ASSRI | WFM&CInteractive PCS prototype&RPage &P of &N",
    )


def _performance(
    workbook: xlsxwriter.Workbook,
    rows: Sequence[dict[str, Any]],
) -> None:
    worksheet = workbook.add_worksheet("PERFORMANCE")
    headers = (
        "Period", "From", "To", "Prior From", "Prior To", "LOB",
        "Team Leader", "Agent", "Agent ID", "PCS", "Participation",
        "Valid Q1", "Score <= 3", "Inbound Legs", "Prior PCS",
        "Prior Participation",
    )
    _write_sheet_title(
        workbook, worksheet,
        title="PCS PERFORMANCE",
        subtitle="LIGHTWEIGHT GOVERNED RESULTS · native table filters remain available for detailed checks",
        last_column=len(headers) - 1,
        tab_color=COLORS["teal"],
    )
    table_rows: list[tuple[Any, ...]] = []
    for period in PERIODS:
        for lob in LOBS:
            for leader in TEAMS[lob]:
                for agent_index, agent in enumerate(AGENTS[leader]):
                    current_rows = _matching(
                        rows, start=period.start, end=period.end,
                        lob=lob, leader=leader, agent=agent,
                    )
                    prior_rows = _matching(
                        rows, start=period.prior_start, end=period.prior_end,
                        lob=lob, leader=leader, agent=agent,
                    )
                    current = _metric(current_rows, period.start, period.end)
                    prior = _metric(
                        prior_rows, period.prior_start, period.prior_end,
                    )
                    table_rows.append((
                        period.label, period.start, period.end,
                        period.prior_start, period.prior_end,
                        lob, leader, agent,
                        _agent_id(TEAM_INDEX[leader], agent_index),
                        current["pcs"], current["participation"],
                        current["valid"], current["low"], current["calls"],
                        prior["pcs"], prior["participation"],
                    ))
    _add_native_table(
        workbook, worksheet,
        first_row=3, first_col=0, name="tblPcsPerformance",
        headers=headers, rows=table_rows,
    )
    worksheet.freeze_panes(4, 0)
    widths = (25, 12, 12, 12, 12, 14, 22, 24, 13, 11, 15, 12, 13, 13, 11, 16)
    for column, width in enumerate(widths):
        worksheet.set_column(column, column, width)


def _coaching(
    workbook: xlsxwriter.Workbook,
    cache_rows: Sequence[Sequence[Any]],
) -> None:
    worksheet = workbook.add_worksheet("COACHING")
    worksheet.hide_gridlines(2)
    worksheet.set_zoom(78)
    worksheet.set_tab_color(COLORS["gold"])
    worksheet.freeze_panes(4, 0)
    title_format = workbook.add_format({
        "font_name": TITLE_FONT, "font_size": 20, "bold": True,
        "font_color": COLORS["white"], "bg_color": COLORS["navy"],
        "align": "left", "valign": "vcenter", "indent": 1,
    })
    filter_label = workbook.add_format({
        "font_name": BODY_FONT, "font_size": 9, "bold": True,
        "font_color": COLORS["navy"], "bg_color": COLORS["canvas"],
        "align": "left", "valign": "vcenter",
    })
    filter_value = workbook.add_format({
        "font_name": BODY_FONT, "font_size": 11,
        "font_color": COLORS["navy"], "bg_color": COLORS["white"],
        "align": "left", "valign": "vcenter", "indent": 1,
        "border": 1, "border_color": COLORS["line"],
    })
    worksheet.merge_range("A1:R1", "PCS COACHING", title_format)
    worksheet.merge_range("A2:B2", "PERIOD", filter_label)
    worksheet.merge_range("C2:E2", "Current MTD", filter_value)
    worksheet.merge_range("F2:G2", "LOB", filter_label)
    worksheet.merge_range("H2:J2", ALL, filter_value)
    worksheet.merge_range(
        "L2:R2", "Blue table = permanent human-owned action log",
        filter_label,
    )
    worksheet.data_validation("C2", {
        "validate": "list", "source": "=PCS_PERIOD_LIST",
    })
    worksheet.data_validation("H2", {
        "validate": "list", "source": "=PCS_LOB_LIST",
    })
    worksheet.set_row_pixels(0, 52)
    worksheet.set_row_pixels(1, 40)

    queue_headers = COACH_HEADERS[1:]
    queue_capacity = 30
    default_key = "Current MTD|All"
    cache_map = _cache_map(cache_rows)
    cache_last = len(cache_rows) + 1
    queue_rows: list[tuple[Any, ...]] = []
    queue_formulas: list[tuple[str, ...]] = []
    for offset in range(queue_capacity):
        rank = offset + 1
        key_expression = f'$C$2&"|"&$H$2&"|{rank}"'
        cached = cache_map.get(
            f"{default_key}|{rank}", tuple(None for _ in COACH_HEADERS),
        )
        queue_rows.append(tuple(cached[index] for index in range(1, len(COACH_HEADERS))))
        queue_formulas.append(tuple(
            _lookup_formula(
                "_COACH_CACHE", index, key_expression, cache_last,
            )
            for index in range(1, len(COACH_HEADERS))
        ))
    date_format = workbook.add_format({"num_format": "yyyy-mm-dd"})
    decimal_format = workbook.add_format({"num_format": "0.00"})
    text_format = workbook.add_format({"num_format": "@"})
    for row_offset, (formulas, cached_values) in enumerate(
        zip(queue_formulas, queue_rows), 1,
    ):
        for column, (formula, cached) in enumerate(zip(formulas, cached_values)):
            header = queue_headers[column]
            cell_format = None
            if header == "Date":
                cell_format = date_format
            elif header == "Q1 Score":
                cell_format = decimal_format
            elif header in {"Agent ID", "Call ID", "Coaching Key"}:
                cell_format = text_format
            worksheet.write_formula(
                3 + row_offset, column, formula, cell_format,
                _excel_serial(cached),
            )
    worksheet.add_table(
        3, 0, 3 + queue_capacity, len(queue_headers) - 1,
        {
            "name": "tblCoachingQueue",
            "style": "Table Style Medium 2",
            "columns": [{"header": header} for header in queue_headers],
        },
    )

    action_headers = (
        "Coaching Key", "Call ID", "Coaching Status", "Coach",
        "Coaching Date", "Due Date", "Coaching Comment",
    )
    action_rows = [
        (
            queue_rows[index][9], queue_rows[index][7], status, coach,
            coaching_date, due_date, comment,
        )
        for index, (status, coach, coaching_date, due_date, comment) in enumerate((
            ("Completed", "Sophie Martin", date(2026, 9, 8), date(2026, 9, 8), "Call reviewed"),
            ("Planned", "Karim Belkacem", None, date(2026, 9, 10), "Session booked"),
            ("Pending", "Elena Rossi", None, date(2026, 9, 11), ""),
        ))
        if index < len(queue_rows) and queue_rows[index][9]
    ]
    action_rows.extend(
        tuple(None for _ in action_headers)
        for _ in range(max(25, 100 - len(action_rows)))
    )
    _add_native_table(
        workbook, worksheet,
        first_row=3, first_col=11,
        name="tblCoachingActions", headers=action_headers, rows=action_rows,
        editable=set(action_headers),
    )
    worksheet.data_validation(4, 13, 5003, 13, {
        "validate": "list",
        "source": ["Pending", "Planned", "Completed", "Not required"],
    })
    worksheet.conditional_format(4, 11, 5003, 11, {
        "type": "duplicate",
        "format": workbook.add_format({"bg_color": COLORS["red_light"]}),
    })
    queue_widths = (12, 13, 20, 24, 12, 10, 11, 24, 34, 40)
    for column, width in enumerate(queue_widths):
        worksheet.set_column(column, column, width)
    worksheet.set_column(10, 10, 3)
    action_widths = (40, 24, 16, 20, 16, 16, 34)
    for offset, width in enumerate(action_widths, 11):
        worksheet.set_column(offset, offset, width)
    worksheet.set_footer(
        "&LPrepared by Anass ASSRI | WFM&CInteractive PCS coaching prototype&RPage &P of &N",
    )


def _write_cache_sheet(
    workbook: xlsxwriter.Workbook,
    name: str,
    headers: Sequence[str],
    rows: Sequence[Sequence[Any]],
) -> None:
    worksheet = workbook.add_worksheet(name)
    for column, header in enumerate(headers):
        worksheet.write(0, column, header)
    date_format = workbook.add_format({"num_format": "yyyy-mm-dd"})
    for row_number, row in enumerate(rows, 1):
        for column, value in enumerate(row):
            worksheet.write(
                row_number, column, value,
                date_format if isinstance(value, date) else None,
            )
    worksheet.hide()


def _add_lists(workbook: xlsxwriter.Workbook) -> None:
    worksheet = workbook.add_worksheet("_LISTS")
    lists: list[tuple[str, tuple[str, ...]]] = [
        ("PCS_PERIOD_LIST", tuple(period.label for period in PERIODS)),
        ("PCS_LOB_LIST", (ALL, *LOBS)),
    ]
    all_leaders = tuple(leader for lob in LOBS for leader in TEAMS[lob])
    all_agents = tuple(agent for leader in all_leaders for agent in AGENTS[leader])
    lists.append(("PCS_TL_ALL", (ALL, *all_leaders)))
    for lob in LOBS:
        lists.append((f"PCS_TL_{_slug(lob)}", (ALL, *TEAMS[lob])))
    for lob in (ALL, *LOBS):
        leaders = all_leaders if lob == ALL else TEAMS[lob]
        lob_agents = all_agents if lob == ALL else tuple(
            agent for leader in leaders for agent in AGENTS[leader]
        )
        lists.append((f"PCS_AGENT_{_slug(lob)}_ALL", (ALL, *lob_agents)))
        for leader in leaders:
            lists.append((
                f"PCS_AGENT_{_slug(lob)}_{_slug(leader)}",
                (ALL, *AGENTS[leader]),
            ))
    for column, (name, values) in enumerate(lists):
        worksheet.write(0, column, name)
        for row, value in enumerate(values, 1):
            worksheet.write(row, column, value)
        letter = xl_col_to_name(column)
        workbook.define_name(
            name,
            f"='_LISTS'!${letter}$2:${letter}${len(values) + 1}",
        )
    worksheet.hide()


def _help(workbook: xlsxwriter.Workbook) -> None:
    worksheet = workbook.add_worksheet("HOW IT WORKS")
    _write_sheet_title(
        workbook, worksheet,
        title="PCS — FILTER PROTOTYPE",
        subtitle="DUMMY DATA · test the dropdown interaction before production migration",
        last_column=7, tab_color=COLORS["gold"],
    )
    section = workbook.add_format({
        "font_name": TITLE_FONT, "font_size": 12, "bold": True,
        "font_color": COLORS["white"], "bg_color": COLORS["navy"],
        "align": "left", "valign": "vcenter", "indent": 1,
    })
    body = workbook.add_format({
        "font_name": BODY_FONT, "font_size": 11,
        "font_color": COLORS["navy"], "bg_color": COLORS["white"],
        "text_wrap": True, "valign": "top",
        "border": 1, "border_color": COLORS["line"],
    })
    worksheet.merge_range(3, 0, 3, 7, "HOW TO TEST", section)
    instructions = (
        ("1", "Open OVERVIEW", "Use the four dropdowns from left to right."),
        ("2", "Select a LOB", "Team Leader and Agent lists become specific to that LOB."),
        ("3", "Select a Team Leader", "The Agent dropdown becomes specific to that leader."),
        ("4", "Review the page", "Cards, charts, LOB rows and agent rows respond together."),
        ("5", "Open COACHING", "Period and LOB update the exact low-score call queue."),
        ("6", "Change a parent filter", "Reset child filters to All, then continue left to right."),
    )
    number = workbook.add_format({
        "font_name": TITLE_FONT, "font_size": 18, "bold": True,
        "font_color": COLORS["teal"], "bg_color": COLORS["white"],
        "align": "center", "valign": "vcenter",
        "border": 1, "border_color": COLORS["line"],
    })
    for offset, (step, action, result) in enumerate(instructions, 4):
        worksheet.write(offset, 0, step, number)
        worksheet.merge_range(offset, 1, offset, 3, action, body)
        worksheet.merge_range(offset, 4, offset, 7, result, body)
        worksheet.set_row_pixels(offset, 46)
    worksheet.merge_range(11, 0, 11, 7, "WHAT THIS PROVES", section)
    notes = (
        "The approved Overview can stay visually unchanged while becoming interactive.",
        "All business results are precomputed; Excel performs exact lookups only.",
        "No raw call-leg worksheet, dynamic array, PivotCache, slicer XML, macro, or external link.",
        "Production will replace the hidden dummy caches with lightweight Power Query feeds.",
    )
    for offset, note in enumerate(notes, 12):
        worksheet.merge_range(offset, 0, offset, 7, f"• {note}", body)
        worksheet.set_row_pixels(offset, 38)
    worksheet.set_column_pixels(0, 0, 60)
    worksheet.set_column_pixels(1, 7, 165)


def build(output: Path = OUTPUT) -> Path:
    output.parent.mkdir(parents=True, exist_ok=True)
    rows = _dummy_call_rows()
    kpis, lob_rows, agent_rows, daily_rows = _build_filter_cache(rows)
    coaching_rows = _build_coaching_cache(rows)
    workbook = xlsxwriter.Workbook(output)
    workbook.set_properties({
        "title": "PCS Overview Dropdown Filters Prototype — Dummy Data",
        "subject": "Interactive visual and workflow acceptance prototype",
        "author": "Anass ASSRI",
        "company": "WFM",
        "comments": "Dummy data only. Not a production report.",
    })
    workbook.set_custom_property(
        "WFMHub Prototype", "PCS DROPDOWN FILTERS DUMMY V1",
    )
    workbook.set_calc_mode("auto")
    _overview(workbook, kpis, lob_rows, agent_rows, daily_rows)
    _performance(workbook, rows)
    _coaching(workbook, coaching_rows)
    _help(workbook)
    _write_cache_sheet(workbook, "_KPI_CACHE", KPI_HEADERS, kpis)
    _write_cache_sheet(workbook, "_LOB_CACHE", LOB_HEADERS, lob_rows)
    _write_cache_sheet(workbook, "_AGENT_CACHE", AGENT_HEADERS, agent_rows)
    _write_cache_sheet(workbook, "_DAILY_CACHE", DAILY_HEADERS, daily_rows)
    _write_cache_sheet(workbook, "_COACH_CACHE", COACH_HEADERS, coaching_rows)
    _add_lists(workbook)
    workbook.close()
    return output


if __name__ == "__main__":
    print(build())
