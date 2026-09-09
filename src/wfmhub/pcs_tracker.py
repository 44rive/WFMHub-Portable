"""Permanent, lightweight PCS Power Query tracker lifecycle.

Python and SQLite calculate the governed PCS products and publish fixed CSV
feeds. Desktop Excel uses Power Query only as transport. The permanent workbook
contains lightweight report tables and a permanent human-owned coaching action
log; raw call-leg data is never loaded to a worksheet.
"""

from __future__ import annotations

import os
import shutil
from datetime import date, datetime
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
    write_v2_header,
    write_v2_kpis,
)
from .metrics import load_metric_catalog
from .reports import ExcelReport
from .shared_feeds import (
    PCS_AGENT_SCORECARD_HEADERS,
    PCS_COACHING_HEADERS,
    PCS_DAILY_SCORECARD_HEADERS,
    PCS_LOB_SCORECARD_HEADERS,
    PCS_RESULTS_HEADERS,
    pcs_agent_scorecard_rows,
    pcs_daily_scorecard_rows,
    pcs_lob_scorecard_rows,
    pcs_result_rows,
    publish_pcs_feeds,
)


PCS_TRACKER_FILENAME = "PCS Live Tracker.xlsx"
PCS_TRACKER_VERSION = "2026.11.0"
COACHING_ACTION_HEADERS = (
    "Coaching Key", "Call ID", "Coaching Status", "Coach",
    "Coaching Date", "Due Date", "Coaching Comment",
)


def tracker_path(config: Config) -> Path:
    return (config.reports / PCS_TRACKER_FILENAME).resolve()


def latest_pcs_report(config: Config) -> Path | None:
    path = tracker_path(config)
    return path if path.is_file() else None


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


def _formula_cache(value: Any) -> Any:
    """Return a scalar Excel can safely store as a formula cached value."""

    if value is None:
        return ""
    if isinstance(value, datetime):
        days = (value.date() - date(1899, 12, 30)).days
        seconds = (
            value.hour * 3600 + value.minute * 60 + value.second
            + value.microsecond / 1_000_000
        )
        return days + seconds / 86_400
    if isinstance(value, date):
        return (value - date(1899, 12, 30)).days
    return value


def _read_actions_from(path: Path) -> list[dict[str, Any]]:
    """Read keyed action rows from every supported PCS tracker layout."""

    if not path.is_file():
        return []
    try:
        workbook = load_workbook(
            path, read_only=True, data_only=True, keep_links=False,
        )
    except Exception:
        return []
    try:
        if "COACHING" not in workbook.sheetnames:
            return []
        sheet = workbook["COACHING"]
        header_row = None
        headers: dict[str, int] = {}
        for row_number in range(1, min(40, sheet.max_row) + 1):
            candidate = {
                str(cell.value or "").strip(): cell.column
                for cell in sheet[row_number]
                if cell.value not in (None, "")
            }
            if {"Coaching Key", "Coaching Status", "Coach"} <= set(candidate):
                header_row = row_number
                headers = candidate
                break
        if header_row is None:
            return []
        output: list[dict[str, Any]] = []
        seen: set[str] = set()
        for values in sheet.iter_rows(min_row=header_row + 1, values_only=True):
            key_column = headers["Coaching Key"]
            key = values[key_column - 1] if key_column <= len(values) else None
            key_text = str(key or "").strip()
            if not key_text or key_text in seen:
                continue
            output.append({
                header: values[column - 1] if column <= len(values) else None
                for header, column in headers.items()
                if header in COACHING_ACTION_HEADERS
            })
            seen.add(key_text)
        return output
    finally:
        workbook.close()


def _tracker_contract_version(path: Path) -> str | None:
    """Read the tracker contract without asking Excel to repair the file."""

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
    """Collect unique keyed actions during the one-time architecture migration."""

    candidates = (
        tracker_path(config),
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


def _starter_rows(
    rows: Sequence[Sequence[Any]],
    headers: Sequence[str],
) -> list[tuple[Any, ...]]:
    return [tuple(row) for row in rows] or [tuple(None for _ in headers)]


def _cell_format(report: ExcelReport, header: str, value: Any, *, editable: bool = False):
    if editable:
        return report.editable_date if "Date" in header else report.editable
    if header in {"Date", "As Of Date", "Period Start", "Period End", "Data Through"}:
        return report.date
    if header in {"Call Start", "Feed Refreshed At"}:
        return report.datetime
    if "Participation" in header or header.endswith("Rate"):
        return report.percent
    if any(token in header for token in ("PCS", "Change")) and not any(
        token in header for token in ("Status", "Rule")
    ):
        return report.decimal
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return report.integer
    return report.body


def _write_table(
    report: ExcelReport,
    worksheet,
    *,
    first_row: int,
    first_column: int,
    table_name: str,
    headers: Sequence[str],
    rows: Sequence[Sequence[Any]],
    editable_headers: set[str] | None = None,
) -> None:
    values = _starter_rows(rows, headers)
    editable_headers = editable_headers or set()
    for row_offset, row in enumerate(values, 1):
        for column_offset, value in enumerate(row):
            header = headers[column_offset]
            worksheet.write(
                first_row + row_offset,
                first_column + column_offset,
                value,
                _cell_format(
                    report,
                    header,
                    value,
                    editable=header in editable_headers,
                ),
            )
    worksheet.add_table(
        first_row,
        first_column,
        first_row + len(values),
        first_column + len(headers) - 1,
        {
            "name": table_name,
            "style": "Table Style Light 9",
            "columns": [
                {"header": header, "header_format": report.header}
                for header in headers
            ],
        },
    )


def _query_sheet(
    report: ExcelReport,
    *,
    name: str,
    title: str,
    subtitle: str,
    table_name: str,
    headers: Sequence[str],
    rows: Sequence[Sequence[Any]],
    hidden: bool = False,
):
    worksheet = report.workbook.add_worksheet(name)
    worksheet.hide_gridlines(2)
    worksheet.set_tab_color(COLORS["teal"])
    worksheet.merge_range(0, 0, 0, len(headers) - 1, title, report.title)
    worksheet.merge_range(1, 0, 1, len(headers) - 1, subtitle, report.subtitle)
    worksheet.set_row(0, 30)
    worksheet.set_row(1, 19)
    worksheet.freeze_panes(4, 0)
    _write_table(
        report,
        worksheet,
        first_row=3,
        first_column=0,
        table_name=table_name,
        headers=headers,
        rows=rows,
    )
    for column, header in enumerate(headers):
        width = 15
        if header in {"Agent", "Agent Selector", "Team Leader"}:
            width = 24
        elif header in {"Customer Comment", "Coaching Key"}:
            width = 40
        elif header in {"Call ID"}:
            width = 25
        elif "Date" in header or header.endswith(" At"):
            width = 19
        worksheet.set_column(column, column, width)
    worksheet.set_zoom(85)
    if hidden:
        worksheet.hide()
    return worksheet


def _cached(row: Sequence[Any] | None, header: str, headers: Sequence[str]) -> Any:
    if row is None:
        return None
    return row[headers.index(header)]


def _write_merged_formula(
    worksheet,
    row: int,
    first: int,
    last: int,
    formula: str,
    cell_format,
    cached: Any = "",
) -> None:
    worksheet.merge_range(row, first, row, last, "", cell_format)
    worksheet.write_formula(
        row,
        first,
        formula,
        cell_format,
        _formula_cache(cached),
    )


def _add_overview(
    report: ExcelReport,
    lob_rows: Sequence[Sequence[Any]],
    agent_rows: Sequence[Sequence[Any]],
    daily_rows: Sequence[Sequence[Any]],
) -> None:
    workbook = report.workbook
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
        status="POWER QUERY",
        status_kind="LIVE",
    )

    all_row = next(
        (tuple(row) for row in lob_rows if str(row[0] or "").upper() == "ALL"),
        tuple(lob_rows[0]) if lob_rows else None,
    )
    info_label = workbook.add_format({
        "font_name": "Aptos", "font_size": 8, "bold": True,
        "font_color": COLORS["muted"], "bg_color": COLORS["canvas"],
        "align": "left", "valign": "vcenter",
    })
    info_value = workbook.add_format({
        "font_name": "Aptos", "font_size": 10, "bold": True,
        "font_color": COLORS["navy"], "bg_color": COLORS["white"],
        "align": "left", "valign": "vcenter", "indent": 1,
        "border": 1, "border_color": COLORS["line"],
    })
    info_date = workbook.add_format({
        "font_name": "Aptos", "font_size": 10, "bold": True,
        "font_color": COLORS["navy"], "bg_color": COLORS["white"],
        "align": "left", "valign": "vcenter", "indent": 1,
        "border": 1, "border_color": COLORS["line"],
        "num_format": "yyyy-mm-dd",
    })
    info = (
        ("DATA THROUGH", "=IF('_PCS_LOB'!A5=\"\",\"\",'_PCS_LOB'!R5)", info_date, _cached(all_row, "Data Through", PCS_LOB_SCORECARD_HEADERS)),
        ("REFRESH", "POWER QUERY", info_value, None),
        ("INTERACTION", "NATIVE SLICERS", info_value, None),
        ("SCOPE", "ACTIVE FTE", info_value, None),
    )
    for block, (label, value, value_format, cache) in enumerate(info):
        first = block * 7
        worksheet.merge_range(1, first, 1, first + 1, label, info_label)
        worksheet.merge_range(1, first + 2, 1, first + 6, "", value_format)
        if isinstance(value, str) and value.startswith("="):
            worksheet.write_formula(
                1, first + 2, value, value_format, _formula_cache(cache),
            )
        else:
            worksheet.write(1, first + 2, value, value_format)
    worksheet.set_row_pixels(1, 52)

    write_v2_kpis(worksheet, formats, (
        (
            "Current MTD PCS",
            '=IF(\'_PCS_LOB\'!A5="","",\'_PCS_LOB\'!F5)',
            "decimal",
            _cached(all_row, "Current MTD PCS", PCS_LOB_SCORECARD_HEADERS),
        ),
        (
            "Participation",
            '=IF(\'_PCS_LOB\'!A5="","",\'_PCS_LOB\'!I5)',
            "percent",
            _cached(all_row, "Current MTD Participation", PCS_LOB_SCORECARD_HEADERS),
        ),
        (
            "Prior MTD PCS",
            '=IF(\'_PCS_LOB\'!A5="","",\'_PCS_LOB\'!G5)',
            "decimal",
            _cached(all_row, "Prior MTD PCS", PCS_LOB_SCORECARD_HEADERS),
        ),
        (
            "Change",
            '=IF(\'_PCS_LOB\'!A5="","",\'_PCS_LOB\'!H5)',
            "decimal",
            _cached(all_row, "MTD Change", PCS_LOB_SCORECARD_HEADERS),
        ),
    ))

    lob_chart = workbook.add_chart({"type": "bar"})
    lob_chart.add_series({
        "name": "Current MTD",
        "categories": "='_PCS_LOB'!$A$6:$A$17",
        "values": "='_PCS_LOB'!$F$6:$F$17",
        "fill": {"color": COLORS["teal"]},
        "border": {"none": True},
        "data_labels": {"value": True, "num_format": "0.00"},
    })
    lob_chart.add_series({
        "name": "Prior comparable",
        "categories": "='_PCS_LOB'!$A$6:$A$17",
        "values": "='_PCS_LOB'!$G$6:$G$17",
        "fill": {"color": COLORS["muted"]},
        "border": {"none": True},
    })
    style_v2_chart(lob_chart, title="PCS BY LOB", kind="bar")
    lob_chart.set_x_axis({"min": 1, "max": 5, "major_unit": 1})
    lob_chart.set_y_axis({"reverse": True})

    trend_chart = workbook.add_chart({"type": "line"})
    trend_chart.add_series({
        "name": "Daily PCS",
        "categories": "='_PCS_DAILY'!$A$5:$A$35",
        "values": "='_PCS_DAILY'!$B$5:$B$35",
        "line": {"color": COLORS["teal"], "width": 2.25},
        "marker": {
            "type": "circle", "size": 5,
            "fill": {"color": COLORS["teal"]},
            "border": {"color": COLORS["teal"]},
        },
    })
    style_v2_chart(trend_chart, title="CURRENT MONTH PCS TREND")
    trend_chart.set_y_axis({"min": 1, "max": 5, "major_unit": 1})
    trend_chart.set_x_axis({"date_axis": True, "num_format": "d-mmm"})
    worksheet.insert_chart(7, 0, lob_chart)
    worksheet.insert_chart(7, 14, trend_chart)

    worksheet.merge_range(
        ACTION_SECTION_ROW,
        0,
        ACTION_SECTION_ROW,
        27,
        "LOB PERFORMANCE · CURRENT MTD VS PRIOR COMPARABLE",
        formats.section,
    )
    lob_headers = (
        "LOB", "VALID Q1", "CURRENT PCS", "PRIOR PCS",
        "CHANGE", "PARTICIPATION", "COACHING DUE",
    )
    for column, header in enumerate(lob_headers):
        first = column * 4
        worksheet.merge_range(
            ACTION_HEADER_ROW, first, ACTION_HEADER_ROW, first + 3,
            header, formats.table_header,
        )
    lob_source_columns = ("A", "K", "F", "G", "H", "I", "N")
    lob_formats = (
        formats.table_text, formats.table_integer, formats.table_decimal,
        formats.table_decimal, formats.table_decimal, formats.table_percent,
        formats.table_integer,
    )
    visible_lobs = [
        tuple(row) for row in lob_rows if str(row[0] or "").upper() != "ALL"
    ]
    for offset in range(12):
        row_index = ACTION_FIRST_ROW + offset
        source_row = 6 + offset
        cached_row = visible_lobs[offset] if offset < len(visible_lobs) else None
        for column, (source_column, cell_format) in enumerate(
            zip(lob_source_columns, lob_formats),
        ):
            first = column * 4
            header = (
                "LOB", "Current MTD Valid Q1", "Current MTD PCS",
                "Prior MTD PCS", "MTD Change", "Current MTD Participation",
                "Current MTD Score <= 3",
            )[column]
            formula = (
                f'=IF(\'_PCS_LOB\'!$A${source_row}="","",'
                f'\'_PCS_LOB\'!${source_column}${source_row})'
            )
            _write_merged_formula(
                worksheet, row_index, first, first + 3, formula, cell_format,
                _cached(cached_row, header, PCS_LOB_SCORECARD_HEADERS),
            )
        worksheet.set_row_pixels(row_index, 26)

    agent_section_row = ACTION_FIRST_ROW + 14
    agent_header_row = agent_section_row + 1
    worksheet.merge_range(
        agent_section_row,
        0,
        agent_section_row,
        27,
        "AGENT PERFORMANCE · FULL FILTERABLE VIEW IS ON PERFORMANCE",
        formats.section,
    )
    agent_headers = (
        ("AGENT", 0, 4), ("AGENT ID", 5, 7), ("LOB", 8, 10),
        ("TEAM LEADER", 11, 15), ("CURRENT PCS", 16, 18),
        ("PRIOR PCS", 19, 21), ("CHANGE", 22, 24),
        ("PARTICIPATION", 25, 27),
    )
    for header, first, last in agent_headers:
        worksheet.merge_range(
            agent_header_row, first, agent_header_row, last,
            header, formats.table_header,
        )
    agent_source_columns = ("A", "B", "C", "D", "E", "F", "G", "H")
    agent_source_headers = (
        "Agent", "Agent ID", "LOB", "Team Leader", "Current MTD PCS",
        "Prior MTD PCS", "MTD Change", "Current MTD Participation",
    )
    agent_formats = (
        formats.table_text, formats.table_text, formats.table_text,
        formats.table_text, formats.table_decimal, formats.table_decimal,
        formats.table_decimal, formats.table_percent,
    )
    capacity = max(50, min(1000, len(agent_rows) + 50))
    for offset in range(capacity):
        row_index = agent_header_row + 1 + offset
        source_row = 5 + offset
        cached_row = tuple(agent_rows[offset]) if offset < len(agent_rows) else None
        for (_header, first, last), source_column, source_header, cell_format in zip(
            agent_headers,
            agent_source_columns,
            agent_source_headers,
            agent_formats,
        ):
            formula = (
                f'=IF(\'_PCS_AGENT\'!$A${source_row}="","",'
                f'\'_PCS_AGENT\'!${source_column}${source_row})'
            )
            _write_merged_formula(
                worksheet, row_index, first, last, formula, cell_format,
                _cached(cached_row, source_header, PCS_AGENT_SCORECARD_HEADERS),
            )
        worksheet.set_row_pixels(row_index, 25)
    worksheet.set_footer(
        "&LPrepared by Anass ASSRI | WFM&CPCS operational tracker&RPage &P of &N",
    )


def _add_coaching(
    report: ExcelReport,
    coaching_rows: Sequence[Sequence[Any]],
    actions: Sequence[Sequence[Any]],
) -> None:
    worksheet = report.workbook.add_worksheet("COACHING")
    worksheet.hide_gridlines(2)
    worksheet.set_zoom(78)
    worksheet.set_tab_color(COLORS["gold"])
    worksheet.freeze_panes(4, 0)
    worksheet.merge_range("A1:R1", "PCS COACHING", report.title)
    worksheet.merge_range(
        "A2:R2",
        "Power Query queue on the left · permanent blue action log on the right · use native table filters or add slicers",
        report.subtitle,
    )
    _write_table(
        report,
        worksheet,
        first_row=3,
        first_column=0,
        table_name="tblCoachingQueue",
        headers=PCS_COACHING_HEADERS,
        rows=coaching_rows,
    )
    padded_actions = [tuple(row) for row in actions]
    padded_actions.extend(
        tuple(None for _ in COACHING_ACTION_HEADERS)
        for _ in range(max(25, 100 - len(padded_actions)))
    )
    _write_table(
        report,
        worksheet,
        first_row=3,
        first_column=11,
        table_name="tblCoachingActions",
        headers=COACHING_ACTION_HEADERS,
        rows=padded_actions,
        editable_headers=set(COACHING_ACTION_HEADERS),
    )
    worksheet.data_validation(4, 13, 5003, 13, {
        "validate": "list",
        "source": ["Pending", "Planned", "Completed", "Not required"],
    })
    worksheet.conditional_format(4, 11, 5003, 11, {
        "type": "duplicate",
        "format": report.error,
    })
    queue_widths = (12, 13, 20, 24, 12, 10, 11, 24, 34, 40)
    for column, width in enumerate(queue_widths):
        worksheet.set_column(column, column, width)
    worksheet.set_column(10, 10, 3)
    action_widths = (40, 24, 16, 20, 16, 16, 34)
    for offset, width in enumerate(action_widths, 11):
        worksheet.set_column(offset, offset, width)
    worksheet.set_footer(
        "&LPrepared by Anass ASSRI | WFM&CPCS coaching&RPage &P of &N",
    )


def _add_setup(report: ExcelReport, config: Config) -> None:
    folder = config.feed / "PCS"
    worksheet = report.add_table_sheet(
        "SETUP",
        "PCS CONNECTION SETUP",
        "Install once with WFMHub; normal updates replace CSV feeds, then Excel Data > Refresh All reloads the tables.",
        ["Setting", "Value", "Why it exists"],
        [
            ("Power Query Installed", "NO", "Set automatically after all five query tables refresh"),
            ("Connection Mode", "LOCAL", "Use the locally synced WFMHub feed folder"),
            ("Connection Owner", "Anass ASSRI", "One owner controls connection changes"),
            ("Local Feed Folder", str(folder), "Fixed clean CSV folder outside the workbook"),
            ("SharePoint Site URL", "https://company.sharepoint.com/sites/WFM", "Reserved for a later SharePoint-feed migration"),
            ("SharePoint Feed Folder", "/Shared Documents/WFMHub/Feed/PCS/", "Reserved SharePoint folder fragment"),
            ("LOB Script", str(folder / "POWER_QUERY_PCS_LOB_LOCAL.txt"), "Overview LOB scorecard"),
            ("Agent Script", str(folder / "POWER_QUERY_PCS_AGENT_LOCAL.txt"), "Overview agent scorecard"),
            ("Daily Script", str(folder / "POWER_QUERY_PCS_DAILY_LOCAL.txt"), "Overview daily trend"),
            ("Performance Script", str(folder / "POWER_QUERY_PCS_RESULTS_LOCAL.txt"), "Native-filter and slicer-ready performance table"),
            ("Coaching Script", str(folder / "POWER_QUERY_COACHING_QUEUE_LOCAL.txt"), "Exact low-score call queue"),
            ("Workbook Last Refreshed", "Never", "Written after desktop Excel finishes Refresh All"),
            ("Last Installer Result", "Not run", "Latest desktop Excel connection result"),
            ("Template Version", PCS_TRACKER_VERSION, "Controls the one-time action-preserving migration"),
        ],
        editable_headers={"Value"},
    )
    worksheet.set_column("A:A", 28)
    worksheet.set_column("B:B", 78)
    worksheet.set_column("C:C", 58)


def _add_help(report: ExcelReport) -> None:
    worksheet = report.add_table_sheet(
        "HELP",
        "PCS — SIMPLE OPERATING MODEL",
        "Power Query is transport only. Python/SQLite calculate every governed result before Excel reads the CSVs.",
        ["Step", "What to do", "Where", "Result"],
        [
            (1, "Run Update PCS data", "WFMHub", "SQLite and the fixed clean CSV feeds are refreshed"),
            (2, "Run Install/repair Power Query once after this upgrade", "PCS menu", "The permanent workbook is connected to the CSV feeds"),
            (3, "Open PCS Live Tracker and choose Data > Refresh All", "Excel", "Overview, Performance and Coaching reload"),
            (4, "Use table filter arrows or Insert > Slicer", "PERFORMANCE", "Filter Period View, Scope Level, LOB, Team Leader or Agent"),
            (5, "Filter the exact low-score calls", "COACHING", "Call ID and Coaching Key identify the call to review"),
            (6, "Copy the key and call ID into the blue action table", "COACHING", "Quality records status, coach, dates and comment permanently"),
        ],
    )
    worksheet.set_column("A:A", 10)
    worksheet.set_column("B:B", 62)
    worksheet.set_column("C:C", 24)
    worksheet.set_column("D:D", 64)


def _add_audit(report: ExcelReport, generated: datetime) -> None:
    worksheet = report.add_table_sheet(
        "_AUDIT",
        "PCS TRACKER CONTRACT",
        "Technical support information.",
        ["Field", "Value"],
        [
            ("Tracker version", PCS_TRACKER_VERSION),
            ("Design", f"{REPORT_DESIGN_ID} {REPORT_DESIGN_VERSION}"),
            ("Created", generated),
            ("Workbook grain", "Lightweight governed scorecards and low-score calls"),
            ("Raw data", "External fixed CSV feeds only; no raw-data worksheet"),
            ("Update method", "Power Query direct from CSV"),
            ("KPI method", "Python/SQLite ratio of additive sums"),
            ("Human-owned table", "COACHING!tblCoachingActions"),
        ],
    )
    worksheet.hide()


def ensure_pcs_tracker(
    config: Config,
    *,
    lob_rows: Sequence[Sequence[Any]] = (),
    agent_rows: Sequence[Sequence[Any]] = (),
    daily_rows: Sequence[Sequence[Any]] = (),
    result_rows: Sequence[Sequence[Any]] = (),
    coaching_rows: Sequence[Sequence[Any]] = (),
) -> Path:
    """Create or explicitly migrate the permanent action-preserving tracker."""

    target = tracker_path(config)
    if target.is_file() and _tracker_contract_version(target) == PCS_TRACKER_VERSION:
        return target
    migrated_actions = _legacy_actions(config)
    target.parent.mkdir(parents=True, exist_ok=True)
    partial = target.with_name(f"{target.stem}.partial{target.suffix}")
    generated = datetime.now()
    report = ExcelReport(partial)
    report.workbook.set_calc_mode("auto")
    report.workbook.set_custom_property(
        "WFMHub Report Design", f"{REPORT_DESIGN_ID} {REPORT_DESIGN_VERSION}",
    )
    report.workbook.set_custom_property(
        "WFMHub PCS Live Tracker", PCS_TRACKER_VERSION,
    )
    try:
        _add_overview(report, lob_rows, agent_rows, daily_rows)
        _query_sheet(
            report,
            name="PERFORMANCE",
            title="PCS PERFORMANCE",
            subtitle="Lightweight Power Query output · use native table filters or add slicers · rates are ratio-of-sums results from the Hub",
            table_name="tblPcsPerformance",
            headers=PCS_RESULTS_HEADERS,
            rows=result_rows,
        )
        _add_coaching(report, coaching_rows, migrated_actions)
        _add_setup(report, config)
        _add_help(report)
        _query_sheet(
            report,
            name="_PCS_LOB",
            title="PCS LOB STAGING",
            subtitle="Power Query destination for the lightweight Overview LOB scorecard.",
            table_name="tblPcsLob",
            headers=PCS_LOB_SCORECARD_HEADERS,
            rows=lob_rows,
            hidden=True,
        )
        _query_sheet(
            report,
            name="_PCS_AGENT",
            title="PCS AGENT STAGING",
            subtitle="Power Query destination for the lightweight Overview agent scorecard.",
            table_name="tblPcsAgent",
            headers=PCS_AGENT_SCORECARD_HEADERS,
            rows=agent_rows,
            hidden=True,
        )
        _query_sheet(
            report,
            name="_PCS_DAILY",
            title="PCS DAILY STAGING",
            subtitle="Power Query destination for the lightweight Overview trend.",
            table_name="tblPcsDaily",
            headers=PCS_DAILY_SCORECARD_HEADERS,
            rows=daily_rows,
            hidden=True,
        )
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
            "Close PCS Live Tracker.xlsx and wait for OneDrive to finish syncing. "
            "This upgrade must replace the old tracker once; existing coaching "
            "actions were read before the replacement was attempted."
        ) from exc
    except Exception:
        partial.unlink(missing_ok=True)
        raise
    return target


def build_pcs_live_tracker(
    conn: DatabaseConnection,
    config: Config,
    start: date,
    end: date,
    output: Path | None = None,
) -> Path:
    """Publish fixed CSV feeds and ensure the permanent tracker exists."""

    del output
    first, last = conn.execute(
        "SELECT min(business_date), max(business_date) FROM mart.agent_pcs_day",
    ).fetchone()
    available_start = _as_date(first) or start
    available_end = _as_date(last) or end
    publish_pcs_feeds(conn, config, available_start, available_end)
    target = tracker_path(config)
    if target.is_file() and _tracker_contract_version(target) == PCS_TRACKER_VERSION:
        return target
    metric_catalog = load_metric_catalog(config.home, config.metric_catalog)
    method = metric_catalog.method_for("pcs_average", available_end, {})
    minimum_sample = int(method.minimum_sample) if method is not None else 1
    generated = datetime.now()
    lob_rows = pcs_lob_scorecard_rows(
        conn, available_end, minimum_sample, generated,
    )
    agent_rows = pcs_agent_scorecard_rows(
        conn, available_end, minimum_sample, generated,
    )
    daily_rows = pcs_daily_scorecard_rows(
        conn, available_end.replace(day=1), available_end, generated,
    )
    result_rows = pcs_result_rows(
        conn, available_end, minimum_sample, generated,
    )
    primary_score = f"question_{config.pcs.primary_score_question}_score"
    allowed_scores = ", ".join(
        f"{value:g}" for value in config.pcs.allowed_scores
    )
    coaching_rows = conn.execute(
        f"""SELECT c.business_date, coalesce(d.lob,c.lob), d.team_leader,
                   coalesce(d.canonical_name,c.agent_name), c.agent_id,
                   c.{primary_score},
                   CASE WHEN c.{primary_score}<=2 THEN 'High' ELSE 'Normal' END,
                   c.call_id, c.question_3, c.call_key
            FROM core.clean_call_leg c
            LEFT JOIN core.dim_agent d ON d.agent_id=c.agent_id
            WHERE c.business_date BETWEEN ? AND ?
              AND upper(coalesce(c.call_direction,''))='I'
              AND c.{primary_score} IN ({allowed_scores})
              AND c.{primary_score}<=?
            ORDER BY c.business_date DESC, d.team_leader,
                     coalesce(d.canonical_name,c.agent_name), c.call_start""",
        (
            available_start,
            available_end,
            config.pcs.negative_score_maximum,
        ),
    ).fetchall()
    return ensure_pcs_tracker(
        config,
        lob_rows=lob_rows,
        agent_rows=agent_rows,
        daily_rows=daily_rows,
        result_rows=result_rows,
        coaching_rows=coaching_rows,
    )
