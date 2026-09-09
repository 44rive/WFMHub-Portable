"""Permanent one-paste PCS tracker and clean data-pack lifecycle.

The Hub prepares one clean call-leg table.  Users paste that table into the
permanent tracker's DATA sheet; Excel formulas, charts and coaching views then
recalculate locally.  WFMHub creates the tracker once and never overwrites it.
"""

from __future__ import annotations

import os
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any, Sequence

from openpyxl import load_workbook

from .config import Config
from .database import DatabaseConnection
from .design import COLORS, REPORT_DESIGN_ID, REPORT_DESIGN_VERSION
from .excel_layout import (
    ACTION_FIRST_ROW,
    ACTION_HEADER_ROW,
    ACTION_SECTION_ROW,
    configure_v2_cell_canvas,
    make_v2_formats,
    style_v2_chart,
    write_v2_filters,
    write_v2_header,
    write_v2_kpis,
)
from .reports import ExcelReport


PCS_TRACKER_FILENAME = "PCS Live Tracker.xlsx"
PCS_PASTE_PREFIX = "PCS Paste Data - "
PCS_PASTE_GLOB = f"{PCS_PASTE_PREFIX}*.xlsx"
PCS_TRACKER_VERSION = "2026.10.1"

PCS_INPUT_HEADERS = (
    "Date", "LOB", "Team Leader", "Agent", "Agent ID", "Language",
    "Call ID", "Call Start", "Call Reference Number", "PCS Status 1",
    "Q1 Nonblank", "Valid Q1", "Q1 Score Sum", "Score <= 3",
    "Score > 3", "Invalid Q1", "Q1 Score", "Raw Q1",
    "Customer Comment", "Coaching Key", "Priority",
)
COACHING_ACTION_HEADERS = (
    "Coaching Key", "Call ID", "Coaching Status", "Coach",
    "Coaching Date", "Due Date", "Coaching Comment",
)
PERIODS = (
    "Current MTD", "Latest day", "Current week",
    "Previous MTD same days", "Previous full month",
)


def tracker_path(config: Config) -> Path:
    return (config.reports / PCS_TRACKER_FILENAME).resolve()


def latest_pcs_report(config: Config) -> Path | None:
    path = tracker_path(config)
    return path if path.is_file() else None


def latest_pcs_paste(config: Config) -> Path | None:
    candidates = [path for path in config.reports.glob(PCS_PASTE_GLOB) if path.is_file()]
    return max(candidates, key=lambda path: path.stat().st_mtime, default=None)


def open_workbook(path: Path) -> None:
    if not path.is_file():
        raise FileNotFoundError(path)
    if os.name != "nt":
        raise RuntimeError(f"Open this workbook manually: {path}")
    os.startfile(path)  # type: ignore[attr-defined]


def _as_date(value: Any) -> date | None:
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if value in (None, ""):
        return None
    try:
        return date.fromisoformat(str(value)[:10])
    except ValueError:
        return None


def _month_shift(value: date, months: int) -> date:
    index = value.year * 12 + value.month - 1 + months
    year, month_index = divmod(index, 12)
    month = month_index + 1
    next_index = index + 1
    next_year, next_month_index = divmod(next_index, 12)
    next_month = date(next_year, next_month_index + 1, 1)
    month_end = next_month - timedelta(days=1)
    return date(year, month, min(value.day, month_end.day))


def _initial_periods(latest: date) -> tuple[date, date, date, date]:
    current_from = latest.replace(day=1)
    prior_from = _month_shift(current_from, -1)
    prior_to = min(_month_shift(latest, -1), _month_shift(current_from, 0) - timedelta(days=1))
    return current_from, latest, prior_from, prior_to


def _pcs_input_rows(
    conn: DatabaseConnection,
    config: Config,
    start: date,
    end: date,
) -> list[tuple[Any, ...]]:
    """Return one additive, FTE-scoped inbound call-leg table for Excel."""

    primary = config.pcs.primary_score_question
    participation = config.pcs.participation_question
    primary_score = f"question_{primary}_score"
    raw_answer = f"question_{participation}"
    allowed = ", ".join(f"{value:g}" for value in config.pcs.allowed_scores)
    cursor = conn.execute(
        f"""SELECT c.business_date,
                   coalesce(d.lob,c.lob,'Unassigned'),
                   coalesce(d.team_leader,'Unassigned'),
                   coalesce(d.canonical_name,c.agent_name,'Agent') ||
                       ' [' || c.agent_id || ']',
                   c.agent_id, coalesce(d.language,c.language,''),
                   c.call_id, c.call_start, c.call_reference_number,
                   CASE WHEN coalesce(c.pcs_status,'')=? THEN 1 ELSE 0 END,
                   CASE WHEN coalesce(trim(c.{raw_answer}),'')<>'' THEN 1 ELSE 0 END,
                   CASE WHEN c.{primary_score} IN ({allowed}) THEN 1 ELSE 0 END,
                   CASE WHEN c.{primary_score} IN ({allowed})
                        THEN c.{primary_score} ELSE 0 END,
                   CASE WHEN c.{primary_score} IN ({allowed})
                                  AND c.{primary_score}<=? THEN 1 ELSE 0 END,
                   CASE WHEN c.{primary_score} IN ({allowed})
                                  AND c.{primary_score}>? THEN 1 ELSE 0 END,
                   CASE WHEN coalesce(trim(c.{raw_answer}),'')<>''
                                  AND (c.{primary_score} IS NULL
                                       OR c.{primary_score} NOT IN ({allowed}))
                        THEN 1 ELSE 0 END,
                   CASE WHEN c.{primary_score} IN ({allowed})
                        THEN c.{primary_score} END,
                   c.{raw_answer}, c.question_3, c.call_key,
                   CASE WHEN c.{primary_score} IN ({allowed})
                                  AND c.{primary_score}<=2 THEN 'HIGH'
                        WHEN c.{primary_score} IN ({allowed})
                                  AND c.{primary_score}<=? THEN 'NORMAL'
                        ELSE '' END
            FROM core.clean_call_leg c
            JOIN core.dim_agent d ON d.agent_id=c.agent_id
                                  AND d.match_method='Agent ID'
            WHERE c.business_date BETWEEN ? AND ?
              AND upper(coalesce(c.call_direction,''))='I'
              AND c.agent_id IS NOT NULL
            ORDER BY c.business_date DESC,
                     CASE WHEN c.{primary_score}<=2 THEN 0 ELSE 1 END,
                     d.lob, d.team_leader, d.canonical_name, c.call_start""",
        [
            config.pcs.participation_status,
            config.pcs.negative_score_maximum,
            config.pcs.negative_score_maximum,
            config.pcs.negative_score_maximum,
            start, end,
        ],
    )
    return [tuple(row) for row in cursor.fetchall()]


def _read_actions_from(path: Path) -> list[dict[str, Any]]:
    if not path.is_file():
        return []
    try:
        workbook = load_workbook(path, read_only=True, data_only=True, keep_links=False)
    except Exception:
        return []
    try:
        if "COACHING" not in workbook.sheetnames:
            return []
        sheet = workbook["COACHING"]
        headers = {
            str(cell.value or "").strip(): cell.column
            for cell in sheet[4] if cell.value not in (None, "")
        }
        if "Coaching Key" not in headers:
            return []
        output: list[dict[str, Any]] = []
        for values in sheet.iter_rows(min_row=5, values_only=True):
            key_column = headers["Coaching Key"]
            key = values[key_column - 1] if key_column <= len(values) else None
            if not str(key or "").strip():
                continue
            output.append({
                header: values[column - 1] if column <= len(values) else None
                for header, column in headers.items()
                if header in COACHING_ACTION_HEADERS
            })
        return output
    finally:
        workbook.close()


def _legacy_actions(config: Config) -> list[tuple[Any, ...]]:
    """Import existing keyed actions once into the new permanent tracker."""

    candidates = (
        config.reports / "PCS Coaching Log.xlsx",
        config.reports / "PCS Operational Tracker.xlsx",
    )
    selected: dict[str, dict[str, Any]] = {}
    for path in candidates:
        for row in _read_actions_from(path):
            key = str(row.get("Coaching Key") or "").strip()
            if key and key not in selected:
                selected[key] = row
    return [
        tuple(row.get(header) for header in COACHING_ACTION_HEADERS)
        for row in selected.values()
    ]


def _dynamic_name(workbook, name: str, column: str, maximum_row: int) -> None:
    workbook.define_name(
        name,
        f"='_LISTS'!${column}$2:INDEX('_LISTS'!${column}$2:${column}${maximum_row},"
        f"MAX(1,COUNTA('_LISTS'!${column}$2:${column}${maximum_row})))",
    )


def _period_bounds_formula(period_ref: str, latest_ref: str) -> tuple[str, str, str, str]:
    current_month = f"DATE(YEAR({latest_ref}),MONTH({latest_ref}),1)"
    start = (
        f'IF({period_ref}="Latest day",{latest_ref},'
        f'IF({period_ref}="Current week",{latest_ref}-WEEKDAY({latest_ref},2)+1,'
        f'IF({period_ref}="Current MTD",{current_month},'
        f'IF({period_ref}="Previous MTD same days",EOMONTH({latest_ref},-2)+1,'
        f'EOMONTH({latest_ref},-2)+1))))'
    )
    end = (
        f'IF({period_ref}="Latest day",{latest_ref},'
        f'IF({period_ref}="Current week",{latest_ref},'
        f'IF({period_ref}="Current MTD",{latest_ref},'
        f'IF({period_ref}="Previous MTD same days",'
        f'MIN(EOMONTH({latest_ref},-1),EDATE({latest_ref},-1)),EOMONTH({latest_ref},-1)))))'
    )
    prior_start = (
        f'IF({period_ref}="Latest day",{latest_ref}-1,'
        f'IF({period_ref}="Current week",({start})-7,'
        f'IF({period_ref}="Current MTD",EOMONTH({latest_ref},-2)+1,'
        f'IF({period_ref}="Previous MTD same days",EOMONTH({latest_ref},-3)+1,'
        f'EOMONTH({latest_ref},-3)+1))))'
    )
    prior_end = (
        f'IF({period_ref}="Latest day",{latest_ref}-1,'
        f'IF({period_ref}="Current week",({end})-7,'
        f'IF({period_ref}="Current MTD",'
        f'MIN(EOMONTH({latest_ref},-1),EDATE({latest_ref},-1)),'
        f'IF({period_ref}="Previous MTD same days",'
        f'MIN(EOMONTH({latest_ref},-2),EDATE({latest_ref},-2)),EOMONTH({latest_ref},-2)))))'
    )
    return start, end, prior_start, prior_end


def _add_lists(report: ExcelReport) -> None:
    workbook = report.workbook
    sheet = workbook.add_worksheet("_LISTS")
    sheet.write("A2", "All")
    sheet.write_dynamic_array_formula(
        "A3", '=SORT(UNIQUE(FILTER(tblData[LOB],tblData[LOB]<>"","")))',
    )
    sheet.write("B2", "All")
    sheet.write_dynamic_array_formula(
        "B3",
        '=SORT(UNIQUE(FILTER(tblData[Team Leader],(tblData[Team Leader]<>"")*'
        '(tblData[Date]>=PCS_OV_From)*(tblData[Date]<=PCS_OV_To)*'
        'IF(OVERVIEW!$J$2="All",1,--(tblData[LOB]=OVERVIEW!$J$2)),"")))',
    )
    sheet.write("C2", "All")
    sheet.write_dynamic_array_formula(
        "C3",
        '=SORT(UNIQUE(FILTER(tblData[Agent],(tblData[Agent]<>"")*'
        '(tblData[Date]>=PCS_OV_From)*(tblData[Date]<=PCS_OV_To)*'
        'IF(OVERVIEW!$J$2="All",1,--(tblData[LOB]=OVERVIEW!$J$2))*'
        'IF(OVERVIEW!$Q$2="All",1,--(tblData[Team Leader]=OVERVIEW!$Q$2)),"")))',
    )
    sheet.write("D2", "All")
    sheet.write_dynamic_array_formula(
        "D3",
        '=SORT(UNIQUE(FILTER(tblData[Team Leader],(tblData[Team Leader]<>"")*'
        '(tblData[Date]>=PCS_COACH_From)*(tblData[Date]<=PCS_COACH_To)*'
        'IF(COACHING!$J$2="All",1,--(tblData[LOB]=COACHING!$J$2)),"")))',
    )
    sheet.write("E2", "All")
    sheet.write_dynamic_array_formula(
        "E3",
        '=SORT(UNIQUE(FILTER(tblData[Agent],(tblData[Agent]<>"")*'
        '(tblData[Date]>=PCS_COACH_From)*(tblData[Date]<=PCS_COACH_To)*'
        'IF(COACHING!$J$2="All",1,--(tblData[LOB]=COACHING!$J$2))*'
        'IF(COACHING!$Q$2="All",1,--(tblData[Team Leader]=COACHING!$Q$2)),"")))',
    )
    for row, value in enumerate(PERIODS, 1):
        sheet.write(row, 7, value)
    sheet.write_formula("J2", "=IFERROR(MAX(tblData[Date]),TODAY())", report.date)
    overview = _period_bounds_formula("OVERVIEW!$C$2", "_LISTS!$J$2")
    coaching = _period_bounds_formula("COACHING!$C$2", "_LISTS!$J$2")
    for column, formula in enumerate(overview, 9):
        sheet.write_formula(2, column, f"={formula}", report.date)
    for column, formula in enumerate(coaching[:2], 9):
        sheet.write_formula(3, column, f"={formula}", report.date)

    workbook.define_name("PCS_PERIOD_LIST", "='_LISTS'!$H$2:$H$6")
    _dynamic_name(workbook, "PCS_LOB_LIST", "A", 100)
    _dynamic_name(workbook, "PCS_OV_TL_LIST", "B", 500)
    _dynamic_name(workbook, "PCS_OV_AGENT_LIST", "C", 5000)
    _dynamic_name(workbook, "PCS_COACH_TL_LIST", "D", 500)
    _dynamic_name(workbook, "PCS_COACH_AGENT_LIST", "E", 5000)
    workbook.define_name("PCS_Latest", "='_LISTS'!$J$2")
    workbook.define_name("PCS_OV_From", "='_LISTS'!$J$3")
    workbook.define_name("PCS_OV_To", "='_LISTS'!$K$3")
    workbook.define_name("PCS_OV_Prior_From", "='_LISTS'!$L$3")
    workbook.define_name("PCS_OV_Prior_To", "='_LISTS'!$M$3")
    workbook.define_name("PCS_COACH_From", "='_LISTS'!$J$4")
    workbook.define_name("PCS_COACH_To", "='_LISTS'!$K$4")
    sheet.hide()


def _sumifs(
    value_header: str,
    from_name: str,
    to_name: str,
    *,
    lob_ref: str,
    leader_ref: str,
    agent_ref: str,
) -> str:
    return (
        f'SUMIFS(tblData[{value_header}],tblData[Date],">="&{from_name},'
        f'tblData[Date],"<="&{to_name},tblData[LOB],IF({lob_ref}="All","*",{lob_ref}),'
        f'tblData[Team Leader],IF({leader_ref}="All","*",{leader_ref}),'
        f'tblData[Agent],IF({agent_ref}="All","*",{agent_ref}))'
    )


def _metric_cache(
    rows: Sequence[Sequence[Any]],
    start: date,
    end: date,
    *,
    lob: str = "All",
    leader: str = "All",
    agent: str = "All",
) -> dict[str, Any]:
    indexes = {header: index for index, header in enumerate(PCS_INPUT_HEADERS)}
    selected = []
    for row in rows:
        business_date = _as_date(row[indexes["Date"]])
        if business_date is None or not start <= business_date <= end:
            continue
        if lob != "All" and str(row[indexes["LOB"]]) != lob:
            continue
        if leader != "All" and str(row[indexes["Team Leader"]]) != leader:
            continue
        if agent != "All" and str(row[indexes["Agent"]]) != agent:
            continue
        selected.append(row)
    def total(header: str) -> float:
        return sum(float(row[indexes[header]] or 0) for row in selected)
    valid = total("Valid Q1")
    eligible = total("PCS Status 1")
    return {
        "pcs": total("Q1 Score Sum") / valid if valid else None,
        "participation": total("Q1 Nonblank") / eligible if eligible else None,
        "valid": int(valid),
        "low": int(total("Score <= 3")),
    }


def _write_merged_formula(ws, row: int, first: int, last: int, formula: str, fmt, cached: Any = "") -> None:
    ws.merge_range(row, first, row, last, "", fmt)
    ws.write_formula(row, first, formula, fmt, "" if cached is None else cached)


def _add_overview(report: ExcelReport, rows: Sequence[Sequence[Any]]) -> None:
    workbook = report.workbook
    ws = workbook.add_worksheet("OVERVIEW")
    formats = make_v2_formats(workbook)
    configure_v2_cell_canvas(ws, formats, zoom=82)
    ws.hide_row_col_headers()
    ws.set_tab_color(COLORS["gold"])
    ws.freeze_panes(2, 0)
    write_v2_header(ws, formats, "PCS LIVE TRACKER", status="PASTE READY", status_kind="LIVE")
    write_v2_filters(ws, formats, (
        ("Period", "Current MTD", {"validate": "list", "source": "=PCS_PERIOD_LIST"}),
        ("LOB", "All", {"validate": "list", "source": "=PCS_LOB_LIST"}),
        ("Team Leader", "All", {"validate": "list", "source": "=PCS_OV_TL_LIST"}),
        ("Agent", "All", {"validate": "list", "source": "=PCS_OV_AGENT_LIST"}),
    ))
    workbook.define_name("PCS_OV_Period", "='OVERVIEW'!$C$2")
    workbook.define_name("PCS_OV_LOB", "='OVERVIEW'!$J$2")
    workbook.define_name("PCS_OV_TL", "='OVERVIEW'!$Q$2")
    workbook.define_name("PCS_OV_Agent", "='OVERVIEW'!$X$2")

    latest = max((_as_date(row[0]) for row in rows), default=None) or date.today()
    current_from, current_to, prior_from, prior_to = _initial_periods(latest)
    current_cache = _metric_cache(rows, current_from, current_to)
    prior_cache = _metric_cache(rows, prior_from, prior_to)
    current_score = _sumifs("Q1 Score Sum", "PCS_OV_From", "PCS_OV_To", lob_ref="PCS_OV_LOB", leader_ref="PCS_OV_TL", agent_ref="PCS_OV_Agent")
    current_valid = _sumifs("Valid Q1", "PCS_OV_From", "PCS_OV_To", lob_ref="PCS_OV_LOB", leader_ref="PCS_OV_TL", agent_ref="PCS_OV_Agent")
    current_nonblank = _sumifs("Q1 Nonblank", "PCS_OV_From", "PCS_OV_To", lob_ref="PCS_OV_LOB", leader_ref="PCS_OV_TL", agent_ref="PCS_OV_Agent")
    current_eligible = _sumifs("PCS Status 1", "PCS_OV_From", "PCS_OV_To", lob_ref="PCS_OV_LOB", leader_ref="PCS_OV_TL", agent_ref="PCS_OV_Agent")
    prior_score = _sumifs("Q1 Score Sum", "PCS_OV_Prior_From", "PCS_OV_Prior_To", lob_ref="PCS_OV_LOB", leader_ref="PCS_OV_TL", agent_ref="PCS_OV_Agent")
    prior_valid = _sumifs("Valid Q1", "PCS_OV_Prior_From", "PCS_OV_Prior_To", lob_ref="PCS_OV_LOB", leader_ref="PCS_OV_TL", agent_ref="PCS_OV_Agent")
    current_formula = f'=IFERROR({current_score}/{current_valid},"")'
    prior_formula = f'=IFERROR({prior_score}/{prior_valid},"")'
    write_v2_kpis(ws, formats, (
        ("Current PCS", current_formula, "decimal", current_cache["pcs"]),
        ("Participation", f'=IFERROR({current_nonblank}/{current_eligible},"")', "percent", current_cache["participation"]),
        ("Prior PCS", prior_formula, "decimal", prior_cache["pcs"]),
        ("Change", f'=IFERROR(({current_score}/{current_valid})-({prior_score}/{prior_valid}),"")', "decimal", None if current_cache["pcs"] is None or prior_cache["pcs"] is None else current_cache["pcs"] - prior_cache["pcs"]),
    ))

    ws.merge_range(7, 0, 7, 13, "LOB PERFORMANCE", formats.section)
    lob_headers = ("LOB", "CURRENT PCS", "PRIOR PCS", "CHANGE")
    spans = ((0, 3), (4, 6), (7, 9), (10, 13))
    for (first, last), header in zip(spans, lob_headers):
        ws.merge_range(8, first, 8, last, header, formats.table_header)
    lobs = sorted({str(row[1]) for row in rows if str(row[1] or "").strip()}, key=str.casefold)
    for offset in range(12):
        row_index = 9 + offset
        row_number = row_index + 1
        lob_cache = lobs[offset] if offset < len(lobs) else ""
        lob_formula = (
            f'=IF(PCS_OV_LOB<>"All",IF(ROW(A{offset + 1})=1,PCS_OV_LOB,""),'
            f'IFERROR(INDEX(_LISTS!$A$3:$A$100,{offset + 1}),""))'
        )
        _write_merged_formula(ws, row_index, 0, 3, lob_formula, formats.table_text, lob_cache)
        lob_ref = f"$A${row_number}"
        score = _sumifs("Q1 Score Sum", "PCS_OV_From", "PCS_OV_To", lob_ref=lob_ref, leader_ref="PCS_OV_TL", agent_ref="PCS_OV_Agent")
        valid = _sumifs("Valid Q1", "PCS_OV_From", "PCS_OV_To", lob_ref=lob_ref, leader_ref="PCS_OV_TL", agent_ref="PCS_OV_Agent")
        pscore = _sumifs("Q1 Score Sum", "PCS_OV_Prior_From", "PCS_OV_Prior_To", lob_ref=lob_ref, leader_ref="PCS_OV_TL", agent_ref="PCS_OV_Agent")
        pvalid = _sumifs("Valid Q1", "PCS_OV_Prior_From", "PCS_OV_Prior_To", lob_ref=lob_ref, leader_ref="PCS_OV_TL", agent_ref="PCS_OV_Agent")
        cached_current = _metric_cache(rows, current_from, current_to, lob=lob_cache)["pcs"] if lob_cache else None
        cached_prior = _metric_cache(rows, prior_from, prior_to, lob=lob_cache)["pcs"] if lob_cache else None
        _write_merged_formula(ws, row_index, 4, 6, f'=IF({lob_ref}="","",IFERROR({score}/{valid},""))', formats.table_decimal, cached_current)
        _write_merged_formula(ws, row_index, 7, 9, f'=IF({lob_ref}="","",IFERROR({pscore}/{pvalid},""))', formats.table_decimal, cached_prior)
        change_cache = None if cached_current is None or cached_prior is None else cached_current - cached_prior
        _write_merged_formula(ws, row_index, 10, 13, f'=IF({lob_ref}="","",IFERROR(({score}/{valid})-({pscore}/{pvalid}),""))', formats.table_decimal, change_cache)
        ws.set_row_pixels(row_index, 23)

    chart = workbook.add_chart({"type": "bar"})
    chart.add_series({
        "name": "Current period", "categories": "=OVERVIEW!$A$10:$A$21",
        "values": "=OVERVIEW!$E$10:$E$21",
        "fill": {"color": COLORS["teal"]}, "border": {"none": True},
    })
    chart.add_series({
        "name": "Prior comparable", "categories": "=OVERVIEW!$A$10:$A$21",
        "values": "=OVERVIEW!$H$10:$H$21",
        "fill": {"color": COLORS["muted"]}, "border": {"none": True},
    })
    style_v2_chart(chart, title="PCS BY LOB", kind="bar")
    chart.set_x_axis({"min": 1, "max": 5, "major_unit": 1})
    ws.insert_chart(7, 14, chart, {"x_offset": 0, "y_offset": 0})

    ws.merge_range(ACTION_SECTION_ROW, 0, ACTION_SECTION_ROW, 27, "AGENT RESULTS · FILTERS ABOVE CONTROL THIS LIST", formats.section)
    agent_headers = ("AGENT", "AGENT ID", "LOB", "TEAM LEADER", "PCS", "PARTICIPATION", "LOW SCORES")
    for index, header in enumerate(agent_headers):
        ws.merge_range(ACTION_HEADER_ROW, index * 4, ACTION_HEADER_ROW, index * 4 + 3, header, formats.table_header)
    agents = sorted({str(row[3]) for row in rows if str(row[3] or "").strip()}, key=str.casefold)
    index = {header: position for position, header in enumerate(PCS_INPUT_HEADERS)}
    agent_identity = {
        str(row[index["Agent"]]): (str(row[index["Agent ID"]]), str(row[index["LOB"]]), str(row[index["Team Leader"]]))
        for row in rows if str(row[index["Agent"]]).strip()
    }
    for offset in range(250):
        row_index = ACTION_FIRST_ROW + offset
        excel_row = row_index + 1
        agent_cache = agents[offset] if offset < len(agents) else ""
        formula = (
            f'=IF(PCS_OV_Agent<>"All",IF(ROW(A{offset + 1})=1,PCS_OV_Agent,""),'
            f'IFERROR(INDEX(_LISTS!$C$3:$C$5000,{offset + 1}),""))'
        )
        _write_merged_formula(ws, row_index, 0, 3, formula, formats.table_text, agent_cache)
        agent_ref = f"$A${excel_row}"
        identity = agent_identity.get(agent_cache, ("", "", ""))
        for start_column, source_header, cached in (
            (4, "Agent ID", identity[0]), (8, "LOB", identity[1]), (12, "Team Leader", identity[2]),
        ):
            identity_formula = f'=IF({agent_ref}="","",INDEX(tblData[{source_header}],MATCH({agent_ref},tblData[Agent],0)))'
            _write_merged_formula(ws, row_index, start_column, start_column + 3, identity_formula, formats.table_text, cached)
        score = _sumifs("Q1 Score Sum", "PCS_OV_From", "PCS_OV_To", lob_ref='"All"', leader_ref='"All"', agent_ref=agent_ref)
        valid = _sumifs("Valid Q1", "PCS_OV_From", "PCS_OV_To", lob_ref='"All"', leader_ref='"All"', agent_ref=agent_ref)
        nonblank = _sumifs("Q1 Nonblank", "PCS_OV_From", "PCS_OV_To", lob_ref='"All"', leader_ref='"All"', agent_ref=agent_ref)
        eligible = _sumifs("PCS Status 1", "PCS_OV_From", "PCS_OV_To", lob_ref='"All"', leader_ref='"All"', agent_ref=agent_ref)
        low = _sumifs("Score <= 3", "PCS_OV_From", "PCS_OV_To", lob_ref='"All"', leader_ref='"All"', agent_ref=agent_ref)
        cache = _metric_cache(rows, current_from, current_to, agent=agent_cache) if agent_cache else {"pcs": None, "participation": None, "low": 0}
        _write_merged_formula(ws, row_index, 16, 19, f'=IF({agent_ref}="","",IFERROR({score}/{valid},""))', formats.table_decimal, cache["pcs"])
        _write_merged_formula(ws, row_index, 20, 23, f'=IF({agent_ref}="","",IFERROR({nonblank}/{eligible},""))', formats.table_percent, cache["participation"])
        _write_merged_formula(ws, row_index, 24, 27, f'=IF({agent_ref}="","",{low})', formats.table_integer, cache["low"])
        ws.set_row_pixels(row_index, 23)
    ws.set_footer("&LPrepared by Anass ASSRI | WFM&CPCS live tracker&RPage &P of &N")


def _add_coaching(
    report: ExcelReport,
    actions: Sequence[Sequence[Any]],
) -> None:
    workbook = report.workbook
    ws = workbook.add_worksheet("COACHING")
    formats = make_v2_formats(workbook)
    configure_v2_cell_canvas(ws, formats, zoom=80)
    ws.set_tab_color(COLORS["gold"])
    ws.freeze_panes(9, 0)
    write_v2_header(ws, formats, "PCS COACHING", status="LIVE ACTION LOG", status_kind="LIVE")
    write_v2_filters(ws, formats, (
        ("Period", "Current MTD", {"validate": "list", "source": "=PCS_PERIOD_LIST"}),
        ("LOB", "All", {"validate": "list", "source": "=PCS_LOB_LIST"}),
        ("Team Leader", "All", {"validate": "list", "source": "=PCS_COACH_TL_LIST"}),
        ("Agent", "All", {"validate": "list", "source": "=PCS_COACH_AGENT_LIST"}),
    ))
    workbook.define_name("PCS_COACH_Period", "='COACHING'!$C$2")
    workbook.define_name("PCS_COACH_LOB", "='COACHING'!$J$2")
    workbook.define_name("PCS_COACH_TL", "='COACHING'!$Q$2")
    workbook.define_name("PCS_COACH_Agent", "='COACHING'!$X$2")
    opportunities = _sumifs("Score <= 3", "PCS_COACH_From", "PCS_COACH_To", lob_ref="PCS_COACH_LOB", leader_ref="PCS_COACH_TL", agent_ref="PCS_COACH_Agent")
    high = (
        'COUNTIFS(tblData[Date],">="&PCS_COACH_From,tblData[Date],"<="&PCS_COACH_To,'
        'tblData[LOB],IF(PCS_COACH_LOB="All","*",PCS_COACH_LOB),'
        'tblData[Team Leader],IF(PCS_COACH_TL="All","*",PCS_COACH_TL),'
        'tblData[Agent],IF(PCS_COACH_Agent="All","*",PCS_COACH_Agent),tblData[Priority],"HIGH")'
    )
    write_v2_kpis(ws, formats, (
        ("Opportunities", f"={opportunities}", "integer", None),
        ("High priority", f"={high}", "integer", None),
        ("Actions logged", '=COUNTIF(tblCoachingActions[Coaching Key],"<>")', "integer", None),
        ("Completed", '=COUNTIF(tblCoachingActions[Coaching Status],"Completed")', "integer", None),
    ))
    ws.merge_range(7, 0, 7, 9, "FILTERED COACHING QUEUE", formats.section)
    ws.merge_range(7, 11, 7, 17, "EDITABLE ACTION LOG", formats.section)
    queue_headers = ("DATE", "LOB", "TEAM LEADER", "AGENT", "Q1", "PRIORITY", "STATUS", "CALL ID", "CUSTOMER COMMENT", "COACHING KEY")
    for column, header in enumerate(queue_headers):
        ws.write(8, column, header, formats.table_header)
    queue_formula = (
        '=LET(m,(tblData[Date]>=PCS_COACH_From)*(tblData[Date]<=PCS_COACH_To)*'
        '(tblData[Score <= 3]=1)*IF(PCS_COACH_LOB="All",1,--(tblData[LOB]=PCS_COACH_LOB))*'
        'IF(PCS_COACH_TL="All",1,--(tblData[Team Leader]=PCS_COACH_TL))*'
        'IF(PCS_COACH_Agent="All",1,--(tblData[Agent]=PCS_COACH_Agent)),'
        'r,FILTER(CHOOSECOLS(tblData,1,2,3,4,17,21,7,19,20),m,""),'
        'IFERROR(HSTACK(CHOOSECOLS(r,1,2,3,4,5,6),'
        'XLOOKUP(CHOOSECOLS(r,9),tblCoachingActions[Coaching Key],'
        'tblCoachingActions[Coaching Status],"Pending"),CHOOSECOLS(r,7,8,9)),""))'
    )
    ws.write_dynamic_array_formula(9, 0, 9, 0, queue_formula, report.date)
    ws.set_column(0, 0, 12, report.date)
    ws.set_column(1, 1, 14)
    ws.set_column(2, 2, 18)
    ws.set_column(3, 3, 28)
    ws.set_column(4, 4, 12, report.decimal)
    ws.set_column(5, 6, 12)
    ws.set_column(7, 7, 22)
    ws.set_column(8, 8, 36)
    ws.set_column(9, 9, 34)

    padded_actions = [tuple(row) for row in actions]
    padded_actions.extend(
        tuple(None for _ in COACHING_ACTION_HEADERS)
        for _ in range(max(100, 250 - len(padded_actions)))
    )
    for column, header in enumerate(COACHING_ACTION_HEADERS, 11):
        ws.write(8, column, header, report.header)
    for row_number, values in enumerate(padded_actions, 9):
        for offset, value in enumerate(values):
            column = 11 + offset
            header = COACHING_ACTION_HEADERS[offset]
            fmt = report.editable_date if header in {"Coaching Date", "Due Date"} else report.editable
            ws.write(row_number, column, value, fmt)
    ws.add_table(8, 11, 8 + len(padded_actions), 17, {
        "name": "tblCoachingActions",
        "style": "Table Style Light 9",
        "columns": [{"header": header, "header_format": report.header} for header in COACHING_ACTION_HEADERS],
    })
    ws.data_validation(9, 13, 5000, 13, {
        "validate": "list", "source": ["Pending", "Planned", "Completed", "Not required"],
    })
    ws.conditional_format(9, 11, 5000, 11, {
        "type": "duplicate", "format": report.error,
    })
    ws.set_column(11, 11, 34)
    ws.set_column(12, 12, 22)
    ws.set_column(13, 16, 18)
    ws.set_column(17, 17, 36)
    ws.set_footer("&LPrepared by Anass ASSRI | WFM&CPCS coaching&RPage &P of &N")


def _add_help(report: ExcelReport) -> None:
    report.add_table_sheet(
        "HELP", "PCS — FOUR SIMPLE STEPS",
        "The tracker is permanent. WFMHub prepares clean data; it never replaces this workbook.",
        ["Step", "Action", "Where", "Important"],
        [
            (1, "Run Prepare latest PCS data", "WFMHub", "This creates a new PCS Paste Data file"),
            (2, "Open the paste file and copy DATA rows below the header", "PCS Paste Data", "Do not copy the title or header row"),
            (3, "Clear the old tblData body and paste values into cell A5", "This tracker > DATA", "Use Paste Values; keep the table headers unchanged"),
            (4, "Use Period, LOB, Team Leader and Agent from left to right", "OVERVIEW / COACHING", "Cards, chart and lists recalculate automatically"),
            (5, "Paste Coaching Key and Call ID into the blue action log", "COACHING", "Complete status, coach, dates and comment; this remains in the shared file"),
        ],
    )


def _add_audit(report: ExcelReport, generated: datetime) -> None:
    sheet = report.add_table_sheet(
        "_AUDIT", "PCS TRACKER CONTRACT", "Technical support information.",
        ["Field", "Value"],
        [
            ("Tracker version", PCS_TRACKER_VERSION),
            ("Design", f"{REPORT_DESIGN_ID} {REPORT_DESIGN_VERSION}"),
            ("Created", generated),
            ("Input grain", "One FTE-scoped inbound call leg"),
            ("Update method", "Manual Paste Values into tblData body"),
            ("KPI method", "Ratio of additive sums; never average percentages"),
        ],
    )
    sheet.hide()


def ensure_pcs_tracker(
    config: Config,
    initial_rows: Sequence[Sequence[Any]] = (),
) -> Path:
    """Create the collaborative tracker once; never replace user actions."""

    target = tracker_path(config)
    if target.is_file():
        return target
    target.parent.mkdir(parents=True, exist_ok=True)
    partial = target.with_name(f"{target.stem}.partial{target.suffix}")
    generated = datetime.now()
    report = ExcelReport(partial)
    report.workbook.set_calc_mode("auto")
    report.workbook.set_custom_property(
        "WFMHub Report Design", f"{REPORT_DESIGN_ID} {REPORT_DESIGN_VERSION}",
    )
    report.workbook.set_custom_property("WFMHub PCS Live Tracker", PCS_TRACKER_VERSION)
    try:
        data_rows = [tuple(row) for row in initial_rows] or [tuple(None for _ in PCS_INPUT_HEADERS)]
        # Sheet creation order is the user-facing order. Excel formulas can
        # safely refer to tables and defined names that are written later.
        _add_overview(report, data_rows)
        _add_coaching(report, _legacy_actions(config))
        data_sheet = report.add_table_sheet(
            "DATA", "PCS DATA — PASTE HERE",
            "Clear the old table body, then paste VALUES from the newest clean PCS file into A5. Keep headers unchanged.",
            list(PCS_INPUT_HEADERS), data_rows,
            editable_headers=set(PCS_INPUT_HEADERS),
        )
        data_sheet.set_tab_color(COLORS["blue"])
        _add_help(report)
        _add_lists(report)
        _add_audit(report, generated)
        report.close()
        partial.replace(target)
    except Exception:
        partial.unlink(missing_ok=True)
        raise
    return target


def build_pcs_paste_workbook(
    conn: DatabaseConnection,
    config: Config,
    start: date,
    end: date,
) -> tuple[Path, list[tuple[Any, ...]]]:
    rows = _pcs_input_rows(conn, config, start, end)
    if not rows:
        raise RuntimeError("No in-scope inbound PCS call data exists for the selected period.")
    generated = datetime.now()
    target = (config.reports / f"{PCS_PASTE_PREFIX}{generated:%Y-%m-%d %H%M%S}.xlsx").resolve()
    if target.exists():
        target = target.with_name(f"{PCS_PASTE_PREFIX}{generated:%Y-%m-%d %H%M%S_%f}.xlsx")
    partial = target.with_name(f"{target.stem}.partial{target.suffix}")
    report = ExcelReport(partial)
    try:
        sheet = report.add_table_sheet(
            "DATA", "PCS CLEAN PASTE DATA",
            f"Copy rows A5:U{len(rows) + 4} and paste them into {PCS_TRACKER_FILENAME} > DATA!A5.",
            list(PCS_INPUT_HEADERS), rows,
        )
        sheet.set_tab_color(COLORS["blue"])
        report.close()
        partial.replace(target)
    except Exception:
        partial.unlink(missing_ok=True)
        raise
    return target, rows


def build_pcs_live_tracker(
    conn: DatabaseConnection,
    config: Config,
    start: date,
    end: date,
    output: Path | None = None,
) -> Path:
    """Prepare clean paste data and ensure the permanent tracker exists."""

    del output  # The permanent tracker path is an explicit product contract.
    first, last = conn.execute(
        "SELECT min(business_date), max(business_date) FROM mart.agent_pcs_day",
    ).fetchone()
    available_start = _as_date(first) or start
    available_end = _as_date(last) or end
    _paste, rows = build_pcs_paste_workbook(
        conn, config, available_start, available_end,
    )
    tracker = ensure_pcs_tracker(config, rows)
    return tracker
