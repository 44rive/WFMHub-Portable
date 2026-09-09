"""Permanent one-paste PCS tracker and clean data-pack lifecycle.

The Hub prepares one clean call-leg table.  Users paste that table into the
permanent tracker's DATA sheet; Excel formulas, charts and coaching views then
recalculate locally. WFMHub preserves the tracker after any explicit,
action-preserving contract migration.
"""

from __future__ import annotations

import os
import shutil
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any, Sequence
from xml.etree import ElementTree
from zipfile import BadZipFile, ZipFile

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
PCS_TRACKER_VERSION = "2026.10.2"

PCS_INPUT_HEADERS = (
    "Date", "LOB", "Team Leader", "Agent", "Agent ID", "Language",
    "Call ID", "Call Start", "Call Reference Number", "PCS Status 1",
    "Q1 Nonblank", "Valid Q1", "Q1 Score Sum", "Score <= 3",
    "Score > 3", "Invalid Q1", "Q1 Score", "Raw Q1",
    "Customer Comment", "Coaching Key", "Priority",
    "LOB List Flag", "Team Leader List Flag", "Agent List Flag",
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


def _excel_formula_cache(value: Any) -> Any:
    """Return an OOXML-safe cached scalar for an Excel formula cell."""

    if value is None:
        return ""
    if isinstance(value, datetime):
        day = (value.date() - date(1899, 12, 30)).days
        seconds = (
            value.hour * 3600 + value.minute * 60 + value.second
            + value.microsecond / 1_000_000
        )
        return day + seconds / 86_400
    if isinstance(value, date):
        return (value - date(1899, 12, 30)).days
    return value


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
    output: list[tuple[Any, ...]] = []
    seen_lobs: set[str] = set()
    seen_leaders: set[str] = set()
    seen_agents: set[str] = set()
    for raw_row in cursor.fetchall():
        row = tuple(raw_row)
        lob = str(row[1] or "").strip()
        leader = str(row[2] or "").strip()
        agent = str(row[3] or "").strip()
        flags = (
            int(bool(lob) and lob not in seen_lobs),
            int(bool(leader) and leader not in seen_leaders),
            int(bool(agent) and agent not in seen_agents),
        )
        if lob:
            seen_lobs.add(lob)
        if leader:
            seen_leaders.add(leader)
        if agent:
            seen_agents.add(agent)
        output.append((*row, *flags))
    return output


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
        header_row = None
        headers: dict[str, int] = {}
        for row_number in range(1, min(30, sheet.max_row) + 1):
            candidate = {
                str(cell.value or "").strip(): cell.column
                for cell in sheet[row_number] if cell.value not in (None, "")
            }
            if {"Coaching Key", "Coaching Status", "Coach"} <= set(candidate):
                header_row = row_number
                headers = candidate
                break
        if header_row is None:
            return []
        output: list[dict[str, Any]] = []
        for values in sheet.iter_rows(min_row=header_row + 1, values_only=True):
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


def _tracker_contract_version(path: Path) -> str | None:
    """Read the tracker contract without asking Excel or openpyxl to repair it."""

    try:
        with ZipFile(path) as archive:
            root = ElementTree.fromstring(archive.read("docProps/custom.xml"))
    except (BadZipFile, KeyError, OSError, ElementTree.ParseError):
        return None
    for prop in root:
        if prop.attrib.get("name") != "WFMHub PCS Live Tracker":
            continue
        for child in prop:
            return child.text
    return None


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
        f"MAX(1,LOOKUP(2,1/('_LISTS'!${column}$2:${column}${maximum_row}<>\"\"),"
        f"ROW('_LISTS'!${column}$2:${column}${maximum_row}))-ROW('_LISTS'!${column}$2)+1))",
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


def _add_lists(report: ExcelReport, rows: Sequence[Sequence[Any]]) -> None:
    workbook = report.workbook
    sheet = workbook.add_worksheet("_LISTS")
    indexes = {header: index for index, header in enumerate(PCS_INPUT_HEADERS)}
    lobs = list(dict.fromkeys(
        str(row[indexes["LOB"]]).strip() for row in rows
        if str(row[indexes["LOB"]] or "").strip()
    ))
    leaders = list(dict.fromkeys(
        str(row[indexes["Team Leader"]]).strip() for row in rows
        if str(row[indexes["Team Leader"]] or "").strip()
    ))
    agent_dimension = list(dict.fromkeys(
        (
            str(row[indexes["Agent"]]).strip(),
            str(row[indexes["LOB"]] or "").strip(),
            str(row[indexes["Team Leader"]] or "").strip(),
        )
        for row in rows if str(row[indexes["Agent"]] or "").strip()
    ))
    sheet.write_row("A1", ("LOB", "Team Leader", "Agent", "Agent LOB", "Agent Team Leader"))
    sheet.write("A2", "All")
    sheet.write("B2", "All")
    sheet.write("C2", "All")
    capacities = {
        "LOB": max(20, len(lobs) + 10),
        "Team Leader": max(100, len(leaders) + 25),
        "Agent": max(500, len(agent_dimension) + 100),
    }

    def unique_formula(header: str, flag_header: str, rank: int) -> str:
        position = (
            f"ROW(tblData[{header}])-ROW(INDEX(tblData[{header}],1,1))+1"
        )
        return (
            f'=IFERROR(INDEX(tblData[{header}],AGGREGATE(15,6,{position}/'
            f'(tblData[{flag_header}]=1),{rank})),"")'
        )

    for rank in range(1, capacities["LOB"] + 1):
        sheet.write_formula(1 + rank, 0, unique_formula("LOB", "LOB List Flag", rank), None,
                            lobs[rank - 1] if rank <= len(lobs) else "")
    for rank in range(1, capacities["Team Leader"] + 1):
        sheet.write_formula(1 + rank, 1, unique_formula("Team Leader", "Team Leader List Flag", rank), None,
                            leaders[rank - 1] if rank <= len(leaders) else "")
    for rank in range(1, capacities["Agent"] + 1):
        cached = agent_dimension[rank - 1] if rank <= len(agent_dimension) else ("", "", "")
        sheet.write_formula(1 + rank, 2, unique_formula("Agent", "Agent List Flag", rank), None, cached[0])
        excel_row = rank + 2
        sheet.write_formula(
            1 + rank, 3,
            f'=IF($C${excel_row}="","",INDEX(tblData[LOB],MATCH($C${excel_row},tblData[Agent],0)))',
            None, cached[1],
        )
        sheet.write_formula(
            1 + rank, 4,
            f'=IF($C${excel_row}="","",INDEX(tblData[Team Leader],MATCH($C${excel_row},tblData[Agent],0)))',
            None, cached[2],
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
    _dynamic_name(workbook, "PCS_LOB_LIST", "A", capacities["LOB"] + 2)
    _dynamic_name(workbook, "PCS_OV_TL_LIST", "B", capacities["Team Leader"] + 2)
    _dynamic_name(workbook, "PCS_OV_AGENT_LIST", "C", capacities["Agent"] + 2)
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

    available_dates = [value for row in rows if (value := _as_date(row[0])) is not None]
    latest = max(available_dates, default=date.today())
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
            'IFERROR(INDEX(_LISTS!$C$3:$C$5000,'
            'AGGREGATE(15,6,(ROW(_LISTS!$C$3:$C$5000)-ROW(_LISTS!$C$3)+1)/'
            '((_LISTS!$C$3:$C$5000<>"")*'
            'IF(PCS_OV_LOB="All",1,--(_LISTS!$D$3:$D$5000=PCS_OV_LOB))*'
            'IF(PCS_OV_TL="All",1,--(_LISTS!$E$3:$E$5000=PCS_OV_TL)),'
            f'{offset + 1})),""))'
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
    rows: Sequence[Sequence[Any]],
) -> None:
    workbook = report.workbook
    ws = workbook.add_worksheet("COACHING")
    formats = make_v2_formats(workbook)
    configure_v2_cell_canvas(ws, formats, zoom=80)
    ws.set_tab_color(COLORS["gold"])
    ws.freeze_panes(9, 0)
    write_v2_header(ws, formats, "PCS COACHING", status="LIVE ACTION LOG", status_kind="LIVE")
    ws.merge_range(1, 0, 1, 1, "PERIOD", formats.filter_label)
    ws.merge_range(1, 2, 1, 13, "Current MTD", formats.filter_value)
    ws.data_validation(1, 2, 1, 2, {
        "validate": "list", "source": "=PCS_PERIOD_LIST",
    })
    ws.merge_range(1, 14, 1, 15, "LOB", formats.filter_label)
    ws.merge_range(1, 16, 1, 27, "All", formats.filter_value)
    ws.data_validation(1, 16, 1, 16, {
        "validate": "list", "source": "=PCS_LOB_LIST",
    })
    workbook.define_name("PCS_COACH_Period", "='COACHING'!$C$2")
    workbook.define_name("PCS_COACH_LOB", "='COACHING'!$Q$2")
    opportunities = _sumifs(
        "Score <= 3", "PCS_COACH_From", "PCS_COACH_To",
        lob_ref="PCS_COACH_LOB", leader_ref='"All"', agent_ref='"All"',
    )
    high = (
        'COUNTIFS(tblData[Date],">="&PCS_COACH_From,tblData[Date],"<="&PCS_COACH_To,'
        'tblData[LOB],IF(PCS_COACH_LOB="All","*",PCS_COACH_LOB),'
        'tblData[Priority],"HIGH")'
    )
    indexes = {header: index for index, header in enumerate(PCS_INPUT_HEADERS)}
    available_dates = [
        value for row in rows
        if (value := _as_date(row[indexes["Date"]])) is not None
    ]
    latest = max(available_dates, default=date.today())
    current_from, current_to, _prior_from, _prior_to = _initial_periods(latest)
    initial_matches = [
        (source_index, row) for source_index, row in enumerate(rows, 1)
        if (_as_date(row[indexes["Date"]]) is not None
            and current_from <= _as_date(row[indexes["Date"]]) <= current_to
            and int(row[indexes["Score <= 3"]] or 0) == 1)
    ]
    actions_by_key = {
        str(row[0]): tuple(row) for row in actions if str(row[0] or "").strip()
    }
    completed_actions = sum(
        1 for row in actions_by_key.values()
        if str(row[2] or "").strip().casefold() == "completed"
    )
    write_v2_kpis(ws, formats, (
        ("Opportunities", f"={opportunities}", "integer", len(initial_matches)),
        ("High priority", f"={high}", "integer", sum(1 for _, row in initial_matches if row[indexes["Priority"]] == "HIGH")),
        ("All actions logged", '=COUNTIF(tblCoachingActions[Coaching Key],"<>")', "integer", len(actions_by_key)),
        ("All completed", '=COUNTIF(tblCoachingActions[Coaching Status],"Completed")', "integer", completed_actions),
    ))
    ws.merge_range(7, 0, 7, 9, "FILTERED COACHING QUEUE", formats.section)
    ws.merge_range(7, 11, 7, 17, "EDITABLE ACTION LOG", formats.section)
    queue_headers = ("DATE", "LOB", "TEAM LEADER", "AGENT", "Q1", "PRIORITY", "CALL ID", "CUSTOMER COMMENT", "COACHING KEY", "STATUS")
    for column, header in enumerate(queue_headers):
        ws.write(8, column, header, formats.table_header)
    queue_sources = (
        ("Date", report.date), ("LOB", formats.table_text),
        ("Team Leader", formats.table_text), ("Agent", formats.table_text),
        ("Q1 Score", formats.table_decimal), ("Priority", formats.table_text),
        ("Call ID", formats.table_text), ("Customer Comment", formats.table_text),
        ("Coaching Key", formats.table_text),
    )
    for offset in range(250):
        row_index = 9 + offset
        excel_row = row_index + 1
        cached_index, cached_row = initial_matches[offset] if offset < len(initial_matches) else ("", None)
        helper_formula = (
            '=IFERROR(AGGREGATE(15,6,(ROW(tblData[Date])-'
            'ROW(INDEX(tblData[Date],1,1))+1)/'
            '((tblData[Date]>=PCS_COACH_From)*(tblData[Date]<=PCS_COACH_To)*'
            '(tblData[Score <= 3]=1)*'
            'IF(PCS_COACH_LOB="All",1,--(tblData[LOB]=PCS_COACH_LOB))),'
            f'{offset + 1}),"")'
        )
        ws.write_formula(row_index, 18, helper_formula, None, cached_index)
        for column, (source_header, fmt) in enumerate(queue_sources):
            cached = "" if cached_row is None else _excel_formula_cache(
                cached_row[indexes[source_header]],
            )
            ws.write_formula(
                row_index, column,
                f'=IF($S${excel_row}="","",INDEX(tblData[{source_header}],$S${excel_row}))',
                fmt, cached,
            )
        key_cache = "" if cached_row is None else str(cached_row[indexes["Coaching Key"]] or "")
        action = actions_by_key.get(key_cache)
        status_cache = "" if not key_cache else str(action[2] or "Pending") if action else "Pending"
        ws.write_formula(
            row_index, 9,
            f'=IF($I${excel_row}="","",IFERROR(INDEX(tblCoachingActions[Coaching Status],'
            f'MATCH($I${excel_row},tblCoachingActions[Coaching Key],0)),"Pending"))',
            formats.table_text, status_cache,
        )
        ws.set_row_pixels(row_index, 23)
    ws.set_column(18, 18, None, None, {"hidden": True})
    ws.set_column(0, 0, 12, report.date)
    ws.set_column(1, 1, 14)
    ws.set_column(2, 2, 18)
    ws.set_column(3, 3, 28)
    ws.set_column(4, 4, 12, report.decimal)
    ws.set_column(5, 6, 12)
    ws.set_column(7, 7, 36)
    ws.set_column(8, 8, 34)
    ws.set_column(9, 9, 14)

    padded_actions = [tuple(row) for row in actions]
    padded_actions.extend(
        tuple(None for _ in COACHING_ACTION_HEADERS)
        for _ in range(max(25, 100 - len(padded_actions)))
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
        "HELP", "PCS — SIMPLE UPDATE",
        "The tracker is permanent. Only a versioned repair can rebuild it, and saved coaching actions are carried forward.",
        ["Step", "Action", "Where", "Important"],
        [
            (1, "Run Prepare latest PCS data", "WFMHub", "This creates a new PCS Paste Data file"),
            (2, "Open the paste file and copy DATA rows below the header", "PCS Paste Data", "Do not copy the title or header row"),
            (3, "Clear the old tblData body and paste values into cell A5", "This tracker > DATA", "Use Paste Values; keep the table headers unchanged"),
            (4, "Use Period, LOB, Team Leader and Agent", "OVERVIEW", "Cards, chart and agent list recalculate automatically"),
            (5, "Use Period and LOB", "COACHING", "The exact low-score calls load automatically"),
            (6, "Paste Coaching Key and Call ID into the blue action log", "COACHING", "Complete status, coach, dates and comment; this remains in the shared file"),
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
    """Create or explicitly migrate the tracker while preserving user actions."""

    target = tracker_path(config)
    existing_version = _tracker_contract_version(target) if target.is_file() else None
    if existing_version == PCS_TRACKER_VERSION:
        return target
    migrated_records = _read_actions_from(target) if target.is_file() else []
    migrated_actions = [
        tuple(record.get(header) for header in COACHING_ACTION_HEADERS)
        for record in migrated_records
    ]
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
        _add_coaching(report, migrated_actions or _legacy_actions(config), data_rows)
        data_sheet = report.add_table_sheet(
            "DATA", "PCS DATA — PASTE HERE",
            "Clear the old table body, then paste VALUES from the newest clean PCS file into A5. Keep headers unchanged.",
            list(PCS_INPUT_HEADERS), data_rows,
            editable_headers=set(PCS_INPUT_HEADERS),
        )
        data_sheet.set_tab_color(COLORS["blue"])
        _add_help(report)
        _add_lists(report, data_rows)
        _add_audit(report, generated)
        report.close()
        if target.is_file():
            archive_dir = target.parent / "Archive" / generated.strftime("%Y-%m-%d")
            archive_dir.mkdir(parents=True, exist_ok=True)
            archived = archive_dir / (
                f"{target.stem}_pre_{PCS_TRACKER_VERSION.replace('.', '_')}_"
                f"{generated:%Y%m%d_%H%M%S_%f}{target.suffix}"
            )
            shutil.copy2(target, archived)
        partial.replace(target)
    except PermissionError as exc:
        partial.unlink(missing_ok=True)
        raise RuntimeError(
            "Close PCS Live Tracker.xlsx, wait for OneDrive to finish syncing, "
            "then prepare PCS again. This version must repair the tracker once; "
            "your existing workbook was not deleted."
        ) from exc
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
            f"Copy rows A5:X{len(rows) + 4} and paste values into {PCS_TRACKER_FILENAME} > DATA!A5.",
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
