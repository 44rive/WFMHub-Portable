#!/usr/bin/env python3
"""Build a data-free/dummy PCS V2 workbook for visual user acceptance.

This is intentionally not the production PCS builder. It exercises the exact
V2 pixel grid, real Excel dropdowns, formula recalculation, native charts,
filterable results and editable coaching fields before the Hub is migrated.
"""

from __future__ import annotations

from collections import defaultdict
from datetime import date, datetime, timedelta
from pathlib import Path
import random

import xlsxwriter

from wfmhub.design import COLORS
from wfmhub.excel_layout import (
    ACTION_FIRST_ROW,
    ACTION_HEADER_ROW,
    ACTION_VISIBLE_ROWS,
    V2_GEOMETRY,
    configure_v2_cell_canvas,
    insert_v2_charts,
    make_v2_formats,
    style_v2_chart,
    write_v2_filters,
    write_v2_header,
    write_v2_kpis,
    write_v2_section,
)


OUTPUT = Path(__file__).resolve().parents[1] / "dist" / "PCS-V2-Prototype-Dummy.xlsx"

TEAMS = {
    "RSA NL": ("Sophie Martin", "Karim Belkacem"),
    "RSA BE": ("Elena Rossi",),
    "FORD": ("Daniel Weber",),
    "OEM": ("Laura Jensen", "Miguel Santos"),
}

AGENTS = {
    "Sophie Martin": ("Amina El Idrissi [00124]", "Youssef Benali [00135]", "Nora Amrani [00148]"),
    "Karim Belkacem": ("Samir Haddad [00203]", "Leila Mansouri [00219]", "Mehdi Alaoui [00228]"),
    "Elena Rossi": ("Eva De Smet [00311]", "Lucas Peeters [00325]", "Mila Laurent [00339]"),
    "Daniel Weber": ("Tom Janssen [00402]", "Ines Verhoeven [00417]", "Adam Lemaire [00429]"),
    "Laura Jensen": ("Sarah Muller [00504]", "Omar Diallo [00518]", "Lina Dubois [00531]"),
    "Miguel Santos": ("Rayan Costa [00602]", "Sofia Martin [00616]", "Noah Bernard [00627]"),
}


def _dummy_rows() -> list[dict[str, object]]:
    random.seed(2300)
    rows: list[dict[str, object]] = []
    team_base = {
        "Sophie Martin": 4.28,
        "Karim Belkacem": 3.92,
        "Elena Rossi": 4.10,
        "Daniel Weber": 4.22,
        "Laura Jensen": 4.00,
        "Miguel Santos": 4.18,
    }
    for month_start, days in ((date(2026, 8, 1), 31), (date(2026, 9, 1), 8)):
        for offset in range(days):
            business_date = month_start + timedelta(days=offset)
            if business_date.weekday() >= 5:
                continue
            for lob, leaders in TEAMS.items():
                for leader in leaders:
                    for agent_index, selector in enumerate(AGENTS[leader]):
                        valid = random.randint(2, 7)
                        eligible = random.randint(valid + 2, valid + 10)
                        nonblank = min(eligible, valid + random.randint(0, 3))
                        score = max(2.5, min(5.0, team_base[leader] + random.uniform(-0.45, 0.45)))
                        low = max(0, round(valid * max(0.0, (4.1 - score) / 4.0)))
                        rows.append({
                            "date": business_date,
                            "lob": lob,
                            "team": leader,
                            "agent": selector,
                            "agent_id": selector.rsplit("[", 1)[1].rstrip("]"),
                            "score_sum": score * valid,
                            "valid": valid,
                            "nonblank": nonblank,
                            "eligible": eligible,
                            "low": low,
                            "positive": valid - low,
                            "inbound": random.randint(28, 65),
                            "coaching_done": min(low, random.randint(0, 1)),
                        })
    return rows


def _filter_rows(
    rows: list[dict[str, object]], start: date, end: date,
    lob: str = "All", team: str = "All", agent: str = "All",
) -> list[dict[str, object]]:
    return [
        row for row in rows
        if start <= row["date"] <= end
        and (lob == "All" or row["lob"] == lob)
        and (team == "All" or row["team"] == team)
        and (agent == "All" or row["agent"] == agent)
    ]


def _aggregate(rows: list[dict[str, object]]) -> dict[str, float | int]:
    valid = sum(int(row["valid"]) for row in rows)
    eligible = sum(int(row["eligible"]) for row in rows)
    return {
        "pcs": sum(float(row["score_sum"]) for row in rows) / valid if valid else 0.0,
        "participation": sum(int(row["nonblank"]) for row in rows) / eligible if eligible else 0.0,
        "valid": valid,
        "low": sum(int(row["low"]) for row in rows),
        "inbound": sum(int(row["inbound"]) for row in rows),
        "coaching_done": sum(int(row["coaching_done"]) for row in rows),
    }


def _data_formula_sum(column: str, *, team_cell: str | None = None) -> str:
    criteria = (
        "--(_DATA!$A$5:$A$5000>=PCS_From),"
        "--(_DATA!$A$5:$A$5000<=PCS_To),"
        'IF(OVERVIEW!$J$2="All",1,--(_DATA!$B$5:$B$5000=OVERVIEW!$J$2)),'
        'IF(OVERVIEW!$Q$2="All",1,--(_DATA!$C$5:$C$5000=OVERVIEW!$Q$2)),'
        'IF(OVERVIEW!$X$2="All",1,--(_DATA!$D$5:$D$5000=OVERVIEW!$X$2))'
    )
    if team_cell:
        criteria += f',--(_DATA!$C$5:$C$5000={team_cell})'
    return f"SUMPRODUCT({criteria},N(_DATA!${column}$5:${column}$5000))"


def _prior_formula_sum(column: str, *, lob_cell: str | None = None, team_cell: str | None = None) -> str:
    criteria = (
        "--(_DATA!$A$5:$A$5000>=PCS_Prior_From),"
        "--(_DATA!$A$5:$A$5000<=PCS_Prior_To)"
    )
    if lob_cell:
        criteria += f',--(_DATA!$B$5:$B$5000={lob_cell})'
    else:
        criteria += ',IF(OVERVIEW!$J$2="All",1,--(_DATA!$B$5:$B$5000=OVERVIEW!$J$2))'
    if team_cell:
        criteria += f',--(_DATA!$C$5:$C$5000={team_cell})'
    else:
        criteria += ',IF(OVERVIEW!$Q$2="All",1,--(_DATA!$C$5:$C$5000=OVERVIEW!$Q$2))'
    criteria += ',IF(OVERVIEW!$X$2="All",1,--(_DATA!$D$5:$D$5000=OVERVIEW!$X$2))'
    return f"SUMPRODUCT({criteria},N(_DATA!${column}$5:${column}$5000))"


def _write_merged_value(ws, row: int, start: int, end: int, value, fmt, cached=None) -> None:
    ws.merge_range(row, start, row, end, "", fmt)
    if isinstance(value, str) and value.startswith("="):
        ws.write_formula(row, start, value, fmt, "" if cached is None else cached)
    else:
        ws.write(row, start, value, fmt)


def _overview(workbook, rows: list[dict[str, object]]) -> None:
    ws = workbook.add_worksheet("OVERVIEW")
    formats = make_v2_formats(workbook)
    configure_v2_cell_canvas(ws, formats, zoom=85)
    ws.hide_row_col_headers()
    ws.set_tab_color(COLORS["gold"])
    ws.freeze_panes(2, 0)
    write_v2_header(ws, formats, "PCS OPERATIONS", status="DUMMY DATA", status_kind="PROVISIONAL")
    write_v2_filters(ws, formats, (
        ("Period", "Current MTD", {"validate": "list", "source": "=PeriodList"}),
        ("LOB", "All", {"validate": "list", "source": "=LOBList"}),
        ("Team Leader", "All", {"validate": "list", "source": '=INDIRECT("Teams_"&SUBSTITUTE($J$2," ","_"))'}),
        ("Agent", "All", {"validate": "list", "source": '=INDIRECT("Agents_"&SUBSTITUTE($Q$2," ","_"))'}),
    ))

    current = _aggregate(_filter_rows(rows, date(2026, 9, 1), date(2026, 9, 8)))
    prior = _aggregate(_filter_rows(rows, date(2026, 8, 1), date(2026, 8, 8)))
    score_sum = _data_formula_sum("F")
    valid = _data_formula_sum("G")
    nonblank = _data_formula_sum("H")
    eligible = _data_formula_sum("I")
    prior_score_sum = _prior_formula_sum("F")
    prior_valid = _prior_formula_sum("G")
    current_formula = f'=IFERROR({score_sum}/{valid},"")'
    prior_formula = f'=IFERROR({prior_score_sum}/{prior_valid},"")'
    write_v2_kpis(ws, formats, (
        ("CURRENT MTD PCS", current_formula, "decimal", current["pcs"]),
        ("PARTICIPATION", f'=IFERROR({nonblank}/{eligible},"")', "percent", current["participation"]),
        ("PRIOR MTD PCS", prior_formula, "decimal", prior["pcs"]),
        ("CHANGE", f'=IFERROR(({score_sum}/{valid})-({prior_score_sum}/{prior_valid}),"")', "decimal", current["pcs"] - prior["pcs"]),
    ))
    ws.write_formula("A5", '=UPPER($C$2)&" PCS"', formats.card_label, "CURRENT MTD PCS")
    ws.write_formula("O5", '="PRIOR "&UPPER($C$2)&" PCS"', formats.card_label, "PRIOR MTD PCS")

    lob_chart = workbook.add_chart({"type": "bar"})
    lob_chart.add_series({
        "name": "Current period", "categories": "='_CALC'!$A$2:$A$5",
        "values": "='_CALC'!$B$2:$B$5", "fill": {"color": COLORS["teal"]},
        "border": {"none": True}, "data_labels": {"value": True, "num_format": "0.00"},
    })
    lob_chart.add_series({
        "name": "Prior comparable", "categories": "='_CALC'!$A$2:$A$5",
        "values": "='_CALC'!$C$2:$C$5", "fill": {"color": "#9FB0BC"},
        "border": {"none": True}, "data_labels": {"value": True, "num_format": "0.00"},
    })
    style_v2_chart(lob_chart, title="PCS BY LOB", kind="bar")
    lob_chart.set_x_axis({
        "min": 1, "max": 5, "major_unit": 1,
        "major_gridlines": {"visible": True, "line": {"color": COLORS["line"]}},
    })
    lob_chart.set_y_axis({"reverse": True, "major_gridlines": {"visible": False}})

    trend_chart = workbook.add_chart({"type": "line"})
    trend_chart.add_series({
        "name": "Current month", "categories": "='_CALC'!$E$2:$E$9",
        "values": "='_CALC'!$F$2:$F$9",
        "line": {"color": COLORS["teal"], "width": 2.25},
        "marker": {"type": "circle", "size": 5, "fill": {"color": COLORS["teal"]}},
    })
    trend_chart.add_series({
        "name": "Prior month", "categories": "='_CALC'!$E$2:$E$9",
        "values": "='_CALC'!$G$2:$G$9",
        "line": {"color": "#9FB0BC", "width": 2.25},
        "marker": {"type": "circle", "size": 5, "fill": {"color": "#9FB0BC"}},
    })
    style_v2_chart(trend_chart, title="DAILY PCS TREND")
    trend_chart.set_y_axis({
        "min": 2, "max": 5, "major_unit": .5,
        "major_gridlines": {"visible": True, "line": {"color": COLORS["line"]}},
    })
    trend_chart.set_x_axis({"date_axis": True, "num_format": "d-mmm"})
    insert_v2_charts(ws, lob_chart, trend_chart)

    write_v2_section(ws, formats, "TEAM PERFORMANCE & ACTIONS")
    action_headers = (
        "TEAM LEADER", "PCS AGENTS", "PARTICIPATION", "CURRENT PCS",
        "PRIOR PCS", "CHANGE", "COACHING DUE",
    )
    for index, header in enumerate(action_headers):
        ws.merge_range(ACTION_HEADER_ROW, index * 4, ACTION_HEADER_ROW, index * 4 + 3, header, formats.table_header)
    team_values = list(team for leaders in TEAMS.values() for team in leaders)
    for offset in range(ACTION_VISIBLE_ROWS):
        row_index = ACTION_FIRST_ROW + offset
        team = team_values[offset] if offset < len(team_values) else ""
        team_cell = f"$A${row_index + 1}"
        list_position = offset + 2  # every named team list begins with All
        row_counter = offset + 1
        team_formula = (
            f'=IF(OVERVIEW!$X$2<>"All",IF(ROW(A{row_counter})=1,IFERROR(INDEX(_DATA!$C$5:$C$5000,'
            'MATCH(OVERVIEW!$X$2,_DATA!$D$5:$D$5000,0)),""),""),'
            f'IF(OVERVIEW!$Q$2<>"All",IF(ROW(A{row_counter})=1,OVERVIEW!$Q$2,""),'
            f'IFERROR(INDEX(INDIRECT("Teams_"&SUBSTITUTE(OVERVIEW!$J$2," ","_")),{list_position}),"")))'
        )
        _write_merged_value(ws, row_index, 0, 3, team_formula, formats.table_text, team)
        if not team:
            for start in range(4, 28, 4):
                _write_merged_value(ws, row_index, start, start + 3, "", formats.table_text)
            continue
        lob = next(lob for lob, leaders in TEAMS.items() if team in leaders)
        agent_count = len(AGENTS[team])
        current_team = _aggregate(_filter_rows(rows, date(2026, 9, 1), date(2026, 9, 8), lob=lob, team=team))
        prior_team = _aggregate(_filter_rows(rows, date(2026, 8, 1), date(2026, 8, 8), lob=lob, team=team))
        team_score = _data_formula_sum("F", team_cell=team_cell)
        team_valid = _data_formula_sum("G", team_cell=team_cell)
        team_nonblank = _data_formula_sum("H", team_cell=team_cell)
        team_eligible = _data_formula_sum("I", team_cell=team_cell)
        team_prior_score = _prior_formula_sum("F", team_cell=team_cell)
        team_prior_valid = _prior_formula_sum("G", team_cell=team_cell)
        current_pcs = f'=IFERROR({team_score}/{team_valid},"")'
        prior_pcs = f'=IFERROR({team_prior_score}/{team_prior_valid},"")'
        due = max(0, int(current_team["low"]) - int(current_team["coaching_done"]))
        values = (
            (f'=IF({team_cell}="","",COUNTA(INDIRECT("Agents_"&SUBSTITUTE({team_cell}," ","_")))-1)', formats.table_integer, agent_count),
            (f'=IFERROR({team_nonblank}/{team_eligible},"")', formats.table_percent, current_team["participation"]),
            (current_pcs, formats.table_decimal, current_team["pcs"]),
            (prior_pcs, formats.table_decimal, prior_team["pcs"]),
            (f'=IFERROR(({team_score}/{team_valid})-({team_prior_score}/{team_prior_valid}),"")', formats.positive if current_team["pcs"] >= prior_team["pcs"] else formats.negative, current_team["pcs"] - prior_team["pcs"]),
            (f'=IF({team_cell}="","",MAX(0,{_data_formula_sum("J", team_cell=team_cell)}-{_data_formula_sum("M", team_cell=team_cell)}))', formats.due if due else formats.clear, due),
        )
        for cell_index, (value, fmt, cached) in enumerate(values, 1):
            start = cell_index * 4
            _write_merged_value(ws, row_index, start, start + 3, value, fmt, cached)
    ws.set_footer("&LPrepared by Anass ASSRI | WFM&RPage &P of &N")


def _calc_sheet(workbook, rows: list[dict[str, object]]) -> None:
    ws = workbook.add_worksheet("_CALC")
    header = workbook.add_format({"bold": True, "bg_color": COLORS["navy"], "font_color": "#FFFFFF"})
    num = workbook.add_format({"num_format": "0.00"})
    dt = workbook.add_format({"num_format": "d-mmm"})
    for col, value in enumerate(("LOB", "Current PCS", "Prior PCS", "", "Date", "Current", "Prior")):
        ws.write(0, col, value, header)
    for offset, lob in enumerate(TEAMS, 1):
        current = _aggregate(_filter_rows(rows, date(2026, 9, 1), date(2026, 9, 8), lob=lob))
        prior = _aggregate(_filter_rows(rows, date(2026, 8, 1), date(2026, 8, 8), lob=lob))
        ws.write(offset, 0, lob)
        selector_scope = (
            'IF(OVERVIEW!$Q$2="All",1,--(_DATA!$C$5:$C$5000=OVERVIEW!$Q$2)),'
            'IF(OVERVIEW!$X$2="All",1,--(_DATA!$D$5:$D$5000=OVERVIEW!$X$2))'
        )
        current_sum = (
            f'SUMPRODUCT(--(_DATA!$A$5:$A$5000>=PCS_From),--(_DATA!$A$5:$A$5000<=PCS_To),'
            f'--(_DATA!$B$5:$B$5000=$A{offset + 1}),{selector_scope},N(_DATA!$F$5:$F$5000))'
        )
        current_valid = current_sum.replace("$F$5:$F$5000", "$G$5:$G$5000")
        prior_sum = (
            f'SUMPRODUCT(--(_DATA!$A$5:$A$5000>=PCS_Prior_From),--(_DATA!$A$5:$A$5000<=PCS_Prior_To),'
            f'--(_DATA!$B$5:$B$5000=$A{offset + 1}),{selector_scope},N(_DATA!$F$5:$F$5000))'
        )
        prior_valid = prior_sum.replace("$F$5:$F$5000", "$G$5:$G$5000")
        ws.write_formula(
            offset, 1,
            f'=IF(AND(OVERVIEW!$J$2<>"All",OVERVIEW!$J$2<>$A{offset + 1}),NA(),IFERROR({current_sum}/{current_valid},NA()))',
            num, current["pcs"],
        )
        ws.write_formula(
            offset, 2,
            f'=IF(AND(OVERVIEW!$J$2<>"All",OVERVIEW!$J$2<>$A{offset + 1}),NA(),IFERROR({prior_sum}/{prior_valid},NA()))',
            num, prior["pcs"],
        )
    for offset in range(8):
        current_date = date(2026, 9, 1) + timedelta(days=offset)
        prior_date = date(2026, 8, 1) + timedelta(days=offset)
        current = _aggregate(_filter_rows(rows, current_date, current_date))
        prior = _aggregate(_filter_rows(rows, prior_date, prior_date))
        excel_row = offset + 2
        ws.write_datetime(offset + 1, 4, datetime.combine(current_date, datetime.min.time()), dt)
        current_formula = (
            f'=IFERROR(SUMPRODUCT(--(_DATA!$A$5:$A$5000=$E{excel_row}),'
            'IF(OVERVIEW!$J$2="All",1,--(_DATA!$B$5:$B$5000=OVERVIEW!$J$2)),'
            'IF(OVERVIEW!$Q$2="All",1,--(_DATA!$C$5:$C$5000=OVERVIEW!$Q$2)),'
            'IF(OVERVIEW!$X$2="All",1,--(_DATA!$D$5:$D$5000=OVERVIEW!$X$2)),N(_DATA!$F$5:$F$5000))/'
            f'SUMPRODUCT(--(_DATA!$A$5:$A$5000=$E{excel_row}),'
            'IF(OVERVIEW!$J$2="All",1,--(_DATA!$B$5:$B$5000=OVERVIEW!$J$2)),'
            'IF(OVERVIEW!$Q$2="All",1,--(_DATA!$C$5:$C$5000=OVERVIEW!$Q$2)),'
            'IF(OVERVIEW!$X$2="All",1,--(_DATA!$D$5:$D$5000=OVERVIEW!$X$2)),N(_DATA!$G$5:$G$5000)),NA())'
        )
        prior_formula = current_formula.replace(f"$E{excel_row}", f"EDATE($E{excel_row},-1)")
        ws.write_formula(offset + 1, 5, current_formula, num, current["pcs"] or "#N/A")
        ws.write_formula(offset + 1, 6, prior_formula, num, prior["pcs"] or "#N/A")
    ws.hide()


def _data_sheet(workbook, rows: list[dict[str, object]]) -> None:
    ws = workbook.add_worksheet("_DATA")
    headers = (
        "Date", "LOB", "Team Leader", "Agent Selector", "Agent ID",
        "Q1 Score Sum", "Valid Q1", "Q1 Nonblank", "PCS Status 1",
        "Score <= 3", "Score > 3", "Inbound Call Legs", "Coaching Done",
    )
    date_fmt = workbook.add_format({"num_format": "yyyy-mm-dd"})
    text_fmt = workbook.add_format({"num_format": "@"})
    for row_index, item in enumerate(rows, 4):
        values = (
            item["date"], item["lob"], item["team"], item["agent"], item["agent_id"],
            item["score_sum"], item["valid"], item["nonblank"], item["eligible"],
            item["low"], item["positive"], item["inbound"], item["coaching_done"],
        )
        for col, value in enumerate(values):
            fmt = date_fmt if col == 0 else text_fmt if col == 4 else None
            ws.write(row_index, col, value, fmt)
    ws.add_table(3, 0, 3 + len(rows), len(headers) - 1, {
        "name": "tblDummyPcsData", "style": "Table Style Light 9",
        "columns": [{"header": header} for header in headers],
    })
    ws.hide()


def _lists_sheet(workbook) -> None:
    ws = workbook.add_worksheet("_LISTS")
    periods = ("Latest day", "Current week", "Current MTD", "Previous full month")
    lobs = ("All", *TEAMS)
    lists: dict[str, tuple[str, ...]] = {
        "PeriodList": periods,
        "LOBList": lobs,
        "Teams_All": ("All", *(team for leaders in TEAMS.values() for team in leaders)),
        "Agents_All": ("All", *(agent for values in AGENTS.values() for agent in values)),
    }
    for lob, leaders in TEAMS.items():
        lists[f"Teams_{lob.replace(' ', '_')}"] = ("All", *leaders)
    for leader, agents in AGENTS.items():
        lists[f"Agents_{leader.replace(' ', '_')}"] = ("All", *agents)
    for index, (name, values) in enumerate(lists.items()):
        ws.write(0, index, name)
        for row, value in enumerate(values, 1):
            ws.write(row, index, value)
        col = xlsxwriter.utility.xl_col_to_name(index)
        workbook.define_name(name, f"='_LISTS'!${col}$2:${col}${len(values) + 1}")
    workbook.define_name("PCS_Latest", "=MAX(_DATA!$A$5:$A$5000)")
    workbook.define_name(
        "PCS_From",
        '=IF(OVERVIEW!$C$2="Latest day",PCS_Latest,'
        'IF(OVERVIEW!$C$2="Current week",PCS_Latest-WEEKDAY(PCS_Latest,2)+1,'
        'IF(OVERVIEW!$C$2="Current MTD",EOMONTH(PCS_Latest,-1)+1,EOMONTH(PCS_Latest,-2)+1)))',
    )
    workbook.define_name(
        "PCS_To",
        '=IF(OVERVIEW!$C$2="Previous full month",EOMONTH(PCS_Latest,-1),PCS_Latest)',
    )
    workbook.define_name(
        "PCS_Prior_From",
        '=IF(OVERVIEW!$C$2="Latest day",PCS_From-1,'
        'IF(OVERVIEW!$C$2="Current week",PCS_From-7,EDATE(PCS_From,-1)))',
    )
    workbook.define_name(
        "PCS_Prior_To",
        '=IF(OVERVIEW!$C$2="Latest day",PCS_To-1,'
        'IF(OVERVIEW!$C$2="Current week",PCS_To-7,EDATE(PCS_To,-1)))',
    )
    ws.hide()


def _table_sheet(workbook, name: str, title: str, headers, rows, *, editable=()):
    ws = workbook.add_worksheet(name)
    ws.hide_gridlines(2)
    ws.set_zoom(85)
    ws.set_tab_color(COLORS["teal"] if not editable else COLORS["gold"])
    title_fmt = workbook.add_format({
        "font_name": "Aptos Display", "font_size": 20, "bold": True,
        "font_color": "#FFFFFF", "bg_color": COLORS["navy"],
        "align": "left", "valign": "vcenter", "indent": 1,
    })
    subtitle = workbook.add_format({
        "font_name": "Aptos", "font_size": 10, "font_color": "#FFFFFF",
        "bg_color": COLORS["teal"], "align": "left", "valign": "vcenter", "indent": 1,
    })
    ws.merge_range(0, 0, 0, len(headers) - 1, title, title_fmt)
    ws.merge_range(1, 0, 1, len(headers) - 1, "DUMMY DATA · use native Excel filters; blue columns are editable", subtitle)
    ws.set_row_pixels(0, 52)
    ws.set_row_pixels(1, 28)
    edit_fmt = workbook.add_format({"bg_color": "#E3F0FA", "font_color": COLORS["blue"]})
    for r_index, values in enumerate(rows, 4):
        for col, value in enumerate(values):
            ws.write(r_index, col, value, edit_fmt if headers[col] in editable else None)
    ws.add_table(3, 0, 3 + len(rows), len(headers) - 1, {
        "name": "tblPrototype" + name.replace(" ", "").replace("&", ""),
        "style": "Table Style Medium 2",
        "columns": [{"header": header} for header in headers],
    })
    ws.freeze_panes(4, 0)
    for col, header in enumerate(headers):
        width = 16
        if "Agent" in header or "Comment" in header:
            width = 27
        elif "Date" in header or "Status" in header:
            width = 18
        ws.set_column(col, col, width)
    return ws


def _help_sheet(workbook) -> None:
    ws = workbook.add_worksheet("HOW TO TEST")
    ws.hide_gridlines(2)
    ws.set_zoom(90)
    title = workbook.add_format({
        "font_name": "Aptos Display", "font_size": 20, "bold": True,
        "font_color": "#FFFFFF", "bg_color": COLORS["navy"],
        "align": "left", "valign": "vcenter", "indent": 1,
    })
    section = workbook.add_format({
        "font_name": "Aptos", "font_size": 11, "bold": True,
        "font_color": "#FFFFFF", "bg_color": COLORS["teal"],
        "align": "left", "valign": "vcenter", "indent": 1,
    })
    body = workbook.add_format({
        "font_name": "Aptos", "font_size": 11, "font_color": COLORS["navy"],
        "text_wrap": True, "valign": "top", "bottom": 1,
        "bottom_color": COLORS["line"],
    })
    ws.merge_range("A1:H1", "PCS V2 PROTOTYPE — HOW TO VALIDATE IT", title)
    ws.merge_range("A3:H3", "THIS FILE CONTAINS DUMMY DATA ONLY", section)
    steps = (
        "1. Open OVERVIEW first. Judge the density, equal card sizes, alignment, colours and whether the team action grid is visible without excessive scrolling.",
        "2. Change PERIOD, then LOB, TEAM LEADER and AGENT from left to right. The KPI cards and daily chart recalculate; the downstream dropdown lists follow the selected scope.",
        "3. Check the two charts: both must have exactly the same width and height. No chart should overlap the action section.",
        "4. In RESULTS, use normal Excel filter arrows. This is the detailed reconciliation view, not the management landing page.",
        "5. In AGENT VIEW, confirm the agent-level detail is readable enough for a Team Leader or Quality user.",
        "6. In COACHING, confirm that only the blue Status, Owner, Due Date and Coaching Comment fields look editable.",
        "7. Test at your normal Excel display scaling. Tiny font rendering can vary by computer, but every card, column, chart and section boundary must remain equal.",
        "8. Tell WFMHub what you want changed. Production code will be migrated only after this real workbook is approved.",
    )
    for row, step in enumerate(steps, 4):
        ws.merge_range(row, 0, row, 7, step, body)
        ws.set_row_pixels(row, 54)
    ws.set_column_pixels(0, 7, 150)
    ws.set_row_pixels(0, 52)
    ws.set_tab_color(COLORS["gold"])


def build(output: Path = OUTPUT) -> Path:
    output.parent.mkdir(parents=True, exist_ok=True)
    rows = _dummy_rows()
    workbook = xlsxwriter.Workbook(output)
    workbook.set_properties({
        "title": "PCS V2 Prototype — Dummy Data",
        "subject": "Visual acceptance prototype",
        "author": "Anass ASSRI",
        "company": "WFM",
        "comments": "Dummy data only. Not a production report.",
    })
    workbook.set_custom_property("WFMHub Prototype", "PCS V2 DUMMY")
    workbook.set_calc_mode("auto")
    _overview(workbook, rows)

    results = []
    for item in rows[-80:]:
        results.append((
            item["date"], item["lob"], item["team"], item["agent"], item["agent_id"],
            float(item["score_sum"]) / int(item["valid"]),
            int(item["nonblank"]) / int(item["eligible"]), item["valid"], item["low"], item["inbound"],
        ))
    _table_sheet(
        workbook, "RESULTS", "PCS RESULTS",
        ("Date", "LOB", "Team Leader", "Agent", "Agent ID", "PCS", "Participation", "Valid Q1", "Score <= 3", "Inbound Legs"),
        results,
    )
    _table_sheet(
        workbook, "AGENT VIEW", "AGENT REALISATIONS",
        ("Date", "LOB", "Team Leader", "Agent", "Agent ID", "PCS", "Participation", "Valid Q1", "Low Scores"),
        [row[:-1] for row in results[:45]],
    )
    coaching_rows = [
        (
            f"PCS-{20260902 + index}-{1000 + index}",
            f"CBC-202609-{100001 + index}",
            lob,
            team,
            AGENTS[team][index % 3],
            date(2026, 9, 2 + index % 5),
            score,
            status,
            owner,
            due,
            comment,
        )
        for index, (lob, team, score, status, owner, due, comment) in enumerate((
            ("RSA NL", "Sophie Martin", 2, "Completed", "Sophie", date(2026, 9, 10), "Call reviewed"),
            ("RSA NL", "Karim Belkacem", 3, "Pending", "Karim", date(2026, 9, 12), ""),
            ("RSA BE", "Elena Rossi", 1, "Planned", "Elena", date(2026, 9, 11), "Session booked"),
            ("FORD", "Daniel Weber", 3, "Pending", "Daniel", date(2026, 9, 13), ""),
            ("OEM", "Laura Jensen", 2, "Completed", "Laura", date(2026, 9, 9), "Feedback shared"),
            ("OEM", "Miguel Santos", 3, "Pending", "Miguel", date(2026, 9, 14), ""),
        ))
    ]
    coaching_ws = _table_sheet(
        workbook, "COACHING", "PCS COACHING ACTIONS",
        (
            "Coaching Key", "Call ID", "LOB", "Team Leader", "Agent",
            "Call Date", "Q1 Score", "Status", "Owner", "Due Date",
            "Coaching Comment",
        ),
        coaching_rows,
        editable=("Status", "Owner", "Due Date", "Coaching Comment"),
    )
    coaching_ws.data_validation("H5:H1004", {
        "validate": "list", "source": ["Pending", "Planned", "Completed", "Not required"],
    })
    coaching_ws.data_validation("J5:J1004", {
        "validate": "date", "criteria": "between",
        "minimum": date(2026, 1, 1), "maximum": date(2030, 12, 31),
    })
    _help_sheet(workbook)
    _calc_sheet(workbook, rows)
    _data_sheet(workbook, rows)
    _lists_sheet(workbook)
    workbook.close()
    return output


if __name__ == "__main__":
    print(build())
