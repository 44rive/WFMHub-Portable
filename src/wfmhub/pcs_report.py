"""Python-only PCS snapshot and permanent coaching-log lifecycle.

PCS calculations belong to SQLite/Python.  The generated report is a disposable
snapshot containing final values and native charts; it never asks desktop Excel
to import, calculate, refresh, or save anything.  Human coaching actions live in
one separate workbook that WFMHub reads but never replaces after creation.
"""

from __future__ import annotations

import os
from calendar import monthrange
from collections import defaultdict
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any, Mapping, Sequence

from openpyxl import load_workbook
from xlsxwriter.utility import xl_col_to_name

from .config import Config
from .database import DatabaseConnection
from .design import COLORS, REPORT_DESIGN_ID, REPORT_DESIGN_VERSION
from .excel_layout import V2ChartSpec, render_v2_dashboard
from .metrics import load_metric_catalog
from .reports import ExcelReport, _query
from .rules import load_rulebook
from .shared_feeds import (
    PCS_AGENT_DAY_HEADERS,
    PCS_COACHING_HEADERS,
    PCS_RESULTS_HEADERS,
    _pcs_scope_aggregates,
    pcs_lob_scorecard_rows,
    pcs_result_rows,
)
from .template_reports import DecisionWorkbook


PCS_COACHING_LOG_FILENAME = "PCS Coaching Log.xlsx"
PCS_REPORT_PREFIX = "PCS Operational Report - "
PCS_REPORT_GLOB = f"{PCS_REPORT_PREFIX}*.xlsx"
PCS_SNAPSHOT_VERSION = "2026.09.27"
PCS_SNAPSHOT_RESULTS_HEADERS = tuple(
    "Report Generated At" if header == "Feed Refreshed At" else header
    for header in PCS_RESULTS_HEADERS
)
PCS_SNAPSHOT_DATA_HEADERS = tuple(
    "Report Generated At" if header == "Feed Refreshed At" else header
    for header in PCS_AGENT_DAY_HEADERS
)

COACHING_LOG_HEADERS = (
    "Coaching Key", "Call ID", "LOB", "Team Leader", "Agent Selector",
    "Agent ID", "Agent", "Date", "Call Start", "Q1 Score",
    "Customer Comment", "Call Reference Number", "Language",
    "Coaching Status", "Coach", "Coaching Date", "Due Date",
    "Coaching Comment",
)
COACHING_QUEUE_HEADERS = (*COACHING_LOG_HEADERS[:13], "Priority", "Action Status",
                          "Coach", "Coaching Date", "Due Date", "Coaching Comment")
LOB_SUMMARY_HEADERS = (
    "LOB", "Current MTD PCS", "Prior MTD PCS", "Change",
    "Current MTD Participation", "Valid Q1", "Score <= 3", "Sample State",
)
DAILY_TREND_HEADERS = (
    "Date", "Current PCS", "Prior Comparable PCS", "Participation",
    "Valid Q1", "Score <= 3",
)
DASHBOARD_HEADERS = (
    "Selection Key", "Period View", "Scope Level", "LOB", "Team Leader",
    "Agent Selector", "Display Scope", "PCS Agents", "PCS Average",
    "Prior PCS", "Change", "Participation", "Prior Participation",
    "Valid Q1", "Score <= 3", "Coaching Due",
)


def coaching_log_path(config: Config) -> Path:
    return (config.reports / PCS_COACHING_LOG_FILENAME).resolve()


def latest_pcs_report(config: Config) -> Path | None:
    candidates = [
        path for path in config.reports.glob(PCS_REPORT_GLOB)
        if path.is_file() and ".partial" not in path.stem.casefold()
    ]
    return max(candidates, key=lambda path: path.stat().st_mtime, default=None)


def open_workbook(path: Path) -> None:
    if not path.is_file():
        raise FileNotFoundError(path)
    if os.name != "nt":
        raise RuntimeError(f"Open this workbook manually: {path}")
    os.startfile(path)  # type: ignore[attr-defined]


def _read_coaching_rows(path: Path) -> tuple[list[dict[str, Any]], set[str]]:
    """Read saved actions by Coaching Key without editing the workbook."""

    if not path.is_file():
        return [], set()
    workbook = load_workbook(path, read_only=True, data_only=True, keep_links=False)
    try:
        if "COACHING" not in workbook.sheetnames:
            return [], set()
        sheet = workbook["COACHING"]
        headers = {
            str(cell.value or "").strip(): cell.column
            for cell in sheet[4]
            if cell.value not in (None, "")
        }
        key_column = headers.get("Coaching Key")
        if key_column is None:
            return [], set()
        rows: list[dict[str, Any]] = []
        seen: set[str] = set()
        duplicates: set[str] = set()
        for values in sheet.iter_rows(min_row=5, values_only=True):
            key = values[key_column - 1] if key_column <= len(values) else None
            key_text = str(key or "").strip()
            if not key_text:
                continue
            if key_text in seen:
                duplicates.add(key_text)
                continue
            rows.append({
                header: values[column - 1] if column <= len(values) else None
                for header, column in headers.items()
                if header in COACHING_LOG_HEADERS
            })
            seen.add(key_text)
        return rows, duplicates
    finally:
        workbook.close()


def _queue_by_key(queue_rows: Sequence[Sequence[Any]]) -> dict[str, dict[str, Any]]:
    output: dict[str, dict[str, Any]] = {}
    for values in queue_rows:
        item = dict(zip(PCS_COACHING_HEADERS, values))
        key = str(item.get("Coaching Key") or "").strip()
        if key:
            output[key] = item
    return output


def _legacy_coaching_rows(config: Config) -> list[dict[str, Any]]:
    legacy = (config.reports / "PCS Operational Tracker.xlsx").resolve()
    try:
        rows, _duplicates = _read_coaching_rows(legacy)
        return rows
    except Exception:
        # The legacy workbook remains untouched. A locked or damaged copy must
        # never prevent creation of the new Python-only report architecture.
        return []


def ensure_coaching_log(
    config: Config,
    queue_rows: Sequence[Sequence[Any]] = (),
) -> Path:
    """Create the permanent coaching workbook once and never replace it."""

    target = coaching_log_path(config)
    if target.is_file():
        return target
    target.parent.mkdir(parents=True, exist_ok=True)
    queue = _queue_by_key(queue_rows)
    migrated = _legacy_coaching_rows(config)
    prepared: list[tuple[Any, ...]] = []
    for saved in migrated:
        key = str(saved.get("Coaching Key") or "").strip()
        source = queue.get(key, {})
        combined = {
            header: saved.get(header) if saved.get(header) not in (None, "")
            else source.get(header)
            for header in COACHING_LOG_HEADERS
        }
        combined["Coaching Key"] = key
        prepared.append(tuple(combined.get(header) for header in COACHING_LOG_HEADERS))

    # Ready-to-paste rows keep the first-use workflow obvious. The Hub reads
    # only keyed rows and never rewrites this file after the atomic creation.
    prepared.extend(
        tuple(None for _ in COACHING_LOG_HEADERS)
        for _ in range(max(100, 250 - len(prepared)))
    )
    partial = target.with_name(f"{target.stem}.partial{target.suffix}")
    report = ExcelReport(partial)
    report.workbook.set_custom_property(
        "WFMHub Report Design", f"{REPORT_DESIGN_ID} {REPORT_DESIGN_VERSION}",
    )
    report.workbook.set_custom_property("WFMHub PCS Coaching Log", PCS_SNAPSHOT_VERSION)
    try:
        sheet = report.add_table_sheet(
            "COACHING", "PCS COACHING LOG",
            "Paste columns A:M from the report's COACHING_QUEUE, then complete the blue action columns N:R. This file is permanent; WFMHub only reads it.",
            list(COACHING_LOG_HEADERS), prepared,
            editable_headers=set(COACHING_LOG_HEADERS),
        )
        sheet.freeze_panes(4, 0)
        sheet.data_validation("N5:N10004", {
            "validate": "list",
            "source": ["Pending", "Planned", "Completed", "Not required"],
        })
        sheet.conditional_format("A5:A10004", {
            "type": "duplicate", "format": report.error,
        })
        sheet.set_column("A:A", 34)
        sheet.set_column("B:B", 24)
        sheet.set_column("C:D", 20)
        sheet.set_column("E:E", 30)
        sheet.set_column("F:J", 18)
        sheet.set_column("K:K", 38)
        sheet.set_column("L:Q", 20)
        sheet.set_column("R:R", 42)
        sheet.set_footer("&LPrepared by Anass ASSRI | WFM&CPCS coaching actions&RPage &P of &N")
        report.close()
        partial.replace(target)
    except Exception:
        partial.unlink(missing_ok=True)
        raise
    return target


def _daily_trend_rows(
    conn: DatabaseConnection,
    latest: date,
) -> list[tuple[Any, ...]]:
    current_start = latest.replace(day=1)
    previous_end = current_start - timedelta(days=1)
    previous_start = previous_end.replace(day=1)
    cursor = conn.execute(
        """SELECT business_date, coalesce(sum(pcs_score_sum),0),
                  coalesce(sum(survey_responses),0),
                  coalesce(sum(pcs_participation_responses),0),
                  coalesce(sum(pcs_status_calls),0),
                  coalesce(sum(low_score_responses),0)
           FROM mart.agent_pcs_day
           WHERE business_date BETWEEN ? AND ?
           GROUP BY business_date ORDER BY business_date""",
        [previous_start, latest],
    )
    daily: dict[date, tuple[float, int, int, int, int]] = {}
    for raw_date, score_sum, valid, nonblank, eligible, low in cursor.fetchall():
        business_date = (
            raw_date.date() if isinstance(raw_date, datetime)
            else raw_date if isinstance(raw_date, date)
            else date.fromisoformat(str(raw_date)[:10])
        )
        daily[business_date] = (
            float(score_sum or 0), int(valid or 0), int(nonblank or 0),
            int(eligible or 0), int(low or 0),
        )
    rows = []
    for offset in range(latest.day):
        current_date = current_start + timedelta(days=offset)
        prior_date = previous_start + timedelta(days=offset)
        current = daily.get(current_date, (0.0, 0, 0, 0, 0))
        prior = daily.get(prior_date, (0.0, 0, 0, 0, 0)) if prior_date <= previous_end else (0.0, 0, 0, 0, 0)
        rows.append((
            current_date,
            current[0] / current[1] if current[1] else None,
            prior[0] / prior[1] if prior[1] else None,
            current[2] / current[3] if current[3] else None,
            current[1], current[4],
        ))
    return rows


def _month_bounds(reference: date, offset: int) -> tuple[date, date]:
    month_index = reference.year * 12 + reference.month - 1 + offset
    year, zero_month = divmod(month_index, 12)
    first = date(year, zero_month + 1, 1)
    return first, date(year, zero_month + 1, monthrange(year, zero_month + 1)[1])


def _dashboard_periods(
    latest: date,
) -> tuple[tuple[str, date, date, date, date], ...]:
    """Return each selectable period and its fair prior comparison."""

    current_month, _current_month_end = _month_bounds(latest, 0)
    previous_month, previous_month_end = _month_bounds(latest, -1)
    two_months_back, two_months_back_end = _month_bounds(latest, -2)
    previous_mtd_end = min(
        previous_month_end,
        previous_month + timedelta(days=latest.day - 1),
    )
    two_months_mtd_end = min(
        two_months_back_end,
        two_months_back + timedelta(days=(previous_mtd_end - previous_month).days),
    )
    current_week = latest - timedelta(days=latest.weekday())
    return (
        ("Latest day", latest, latest, latest - timedelta(days=1), latest - timedelta(days=1)),
        ("Current week", current_week, latest, current_week - timedelta(days=7), latest - timedelta(days=7)),
        ("Current MTD", current_month, latest, previous_month, previous_mtd_end),
        ("Previous MTD same days", previous_month, previous_mtd_end, two_months_back, two_months_mtd_end),
        ("Previous full month", previous_month, previous_month_end, two_months_back, two_months_back_end),
    )


def _scope_identity(level: str, item: Mapping[str, Any]) -> tuple[str, str, str]:
    lob = "All" if level == "ALL" else str(item.get("lob") or "Unassigned")
    leader = "All" if level in {"ALL", "LOB"} else str(
        item.get("team_leader") or "Unassigned",
    )
    agent = "All" if level != "AGENT" else str(
        item.get("agent_selector") or "Unassigned",
    )
    return lob, leader, agent


def _scope_display(level: str, identity: tuple[str, str, str]) -> str:
    if level == "ALL":
        return "All active FTE"
    if level == "LOB":
        return identity[0]
    if level == "TEAM":
        return identity[1]
    return identity[2]


def _dashboard_rows(
    conn: DatabaseConnection,
    latest: date,
    queue_rows: Sequence[Sequence[Any]],
    actions: Mapping[str, Mapping[str, Any]],
) -> list[tuple[Any, ...]]:
    """Precalculate every dashboard selection in Python.

    Excel uses only INDEX/MATCH-style lookup formulas. KPI arithmetic remains
    here, which prevents dropdowns from changing the governed business logic.
    """

    queue_indexes = {header: index for index, header in enumerate(PCS_COACHING_HEADERS)}
    output: list[tuple[Any, ...]] = []
    for label, current_start, current_end, prior_start, prior_end in _dashboard_periods(latest):
        current_by_level: dict[str, list[Mapping[str, Any]]] = {}
        prior_by_level: dict[str, dict[tuple[str, str, str], Mapping[str, Any]]] = {}
        for level in ("ALL", "LOB", "TEAM", "AGENT"):
            current_values = _pcs_scope_aggregates(conn, current_start, current_end, level)
            prior_values = _pcs_scope_aggregates(conn, prior_start, prior_end, level)
            current_by_level[level] = current_values
            prior_by_level[level] = {
                _scope_identity(level, item): item for item in prior_values
            }

        agent_counts: defaultdict[tuple[str, str, str], set[str]] = defaultdict(set)
        for agent in current_by_level["AGENT"]:
            identity = _scope_identity("AGENT", agent)
            agent_id = str(agent.get("agent_id") or identity[2])
            agent_counts[("All", "All", "All")].add(agent_id)
            agent_counts[(identity[0], "All", "All")].add(agent_id)
            agent_counts[(identity[0], identity[1], "All")].add(agent_id)
            agent_counts[identity].add(agent_id)

        due: defaultdict[tuple[str, str, str], set[str]] = defaultdict(set)
        for queue_row in queue_rows:
            raw_date = queue_row[queue_indexes["Date"]]
            call_date = (
                raw_date.date() if isinstance(raw_date, datetime)
                else raw_date if isinstance(raw_date, date)
                else date.fromisoformat(str(raw_date)[:10])
            )
            if not current_start <= call_date <= current_end:
                continue
            coaching_key = str(queue_row[queue_indexes["Coaching Key"]] or "").strip()
            state = str(actions.get(coaching_key, {}).get("Coaching Status") or "Pending").strip().casefold()
            if not coaching_key or state in {"completed", "not required"}:
                continue
            identity = (
                str(queue_row[queue_indexes["LOB"]] or "Unassigned"),
                str(queue_row[queue_indexes["Team Leader"]] or "Unassigned"),
                str(queue_row[queue_indexes["Agent Selector"]] or "Unassigned"),
            )
            due[("All", "All", "All")].add(coaching_key)
            due[(identity[0], "All", "All")].add(coaching_key)
            due[(identity[0], identity[1], "All")].add(coaching_key)
            due[identity].add(coaching_key)

        for level in ("ALL", "LOB", "TEAM", "AGENT"):
            for current in current_by_level[level]:
                identity = _scope_identity(level, current)
                prior = prior_by_level[level].get(identity, {})
                current_pcs = current.get("pcs_average")
                prior_pcs = prior.get("pcs_average")
                change = (
                    float(current_pcs) - float(prior_pcs)
                    if current_pcs is not None and prior_pcs is not None else None
                )
                selection_key = "|".join((label, *identity))
                output.append((
                    selection_key, label, level, *identity,
                    _scope_display(level, identity), len(agent_counts[identity]),
                    current_pcs, prior_pcs, change,
                    current.get("participation_rate"), prior.get("participation_rate"),
                    int(current.get("valid_q1") or 0),
                    int(current.get("low_scores") or 0), len(due[identity]),
                ))
    level_order = {"ALL": 0, "LOB": 1, "TEAM": 2, "AGENT": 3}
    output.sort(key=lambda row: (
        tuple(item[0] for item in _dashboard_periods(latest)).index(str(row[1])),
        level_order.get(str(row[2]), 9),
        -int(row[15] or 0),
        row[8] is None,
        float(row[8] or 0),
        str(row[6]).casefold(),
    ))
    return output


def _coaching_action_map(
    rows: Sequence[Mapping[str, Any]],
) -> dict[str, Mapping[str, Any]]:
    return {
        str(row.get("Coaching Key") or "").strip(): row
        for row in rows
        if str(row.get("Coaching Key") or "").strip()
    }


def _cell_range(column: int, first_row: int, last_row: int) -> str:
    name = xl_col_to_name(column)
    return f"${name}${first_row + 1}:${name}${last_row + 1}"


def _define_range(workbook, name: str, column: int, first_row: int, last_row: int) -> None:
    workbook.define_name(name, f"='OVERVIEW'!{_cell_range(column, first_row, last_row)}")


def _dashboard_choices(
    rows: Sequence[Sequence[Any]],
) -> tuple[list[str], dict[str, list[str]], dict[tuple[str, str], list[str]]]:
    agents = {
        (str(row[3]), str(row[4]), str(row[5]))
        for row in rows if str(row[2]) == "AGENT"
    }
    lobs = sorted({item[0] for item in agents}, key=str.casefold)
    teams: dict[str, list[str]] = {}
    agent_choices: dict[tuple[str, str], list[str]] = {}
    for lob in ["All", *lobs]:
        matching = [item for item in agents if lob == "All" or item[0] == lob]
        lob_teams = sorted({item[1] for item in matching}, key=str.casefold)
        teams[lob] = ["All", *lob_teams]
        for leader in teams[lob]:
            values = sorted({
                item[2] for item in matching
                if leader == "All" or item[1] == leader
            }, key=str.casefold)
            agent_choices[(lob, leader)] = ["All", *values]
    return ["All", *lobs], teams, agent_choices


def _add_dashboard_interactivity(
    book: DecisionWorkbook,
    worksheet,
    formats,
    rows: Sequence[Sequence[Any]],
) -> None:
    """Attach safe dropdowns and lookup-only dashboard formulas."""

    workbook = book.report.workbook
    periods = [item[0] for item in _dashboard_periods(book.end)]
    lobs, teams, agents = _dashboard_choices(rows)

    # All support data stays in hidden columns on OVERVIEW. There is no hidden
    # calculation sheet and no external connection.
    period_column = 40
    lob_column = 41
    team_values_column = 42
    agent_values_column = 43
    team_key_column = 44
    team_name_column = 45
    agent_key_column = 46
    agent_name_column = 47
    dashboard_start_column = 48

    for offset, value in enumerate(periods, 1):
        worksheet.write(offset, period_column, value)
    for offset, value in enumerate(lobs, 1):
        worksheet.write(offset, lob_column, value)
    _define_range(workbook, "PCS_PERIOD_LIST", period_column, 1, len(periods))
    _define_range(workbook, "PCS_LOB_LIST", lob_column, 1, len(lobs))

    team_cursor = 1
    team_mappings: list[tuple[str, str]] = []
    for index, lob in enumerate(lobs):
        values = teams[lob]
        first = team_cursor
        for value in values:
            worksheet.write(team_cursor, team_values_column, value)
            team_cursor += 1
        range_name = f"PCS_TL_LIST_{index:03d}"
        _define_range(workbook, range_name, team_values_column, first, team_cursor - 1)
        team_mappings.append((lob, range_name))
    for offset, (key, range_name) in enumerate(team_mappings, 1):
        worksheet.write(offset, team_key_column, key)
        worksheet.write(offset, team_name_column, range_name)
    _define_range(workbook, "PCS_TL_LOOKUP_KEYS", team_key_column, 1, len(team_mappings))
    _define_range(workbook, "PCS_TL_LOOKUP_NAMES", team_name_column, 1, len(team_mappings))

    agent_cursor = 1
    agent_mappings: list[tuple[str, str]] = []
    for index, ((lob, leader), values) in enumerate(agents.items()):
        first = agent_cursor
        for value in values:
            worksheet.write(agent_cursor, agent_values_column, value)
            agent_cursor += 1
        range_name = f"PCS_AGENT_LIST_{index:04d}"
        _define_range(workbook, range_name, agent_values_column, first, agent_cursor - 1)
        agent_mappings.append((f"{lob}|{leader}", range_name))
    for offset, (key, range_name) in enumerate(agent_mappings, 1):
        worksheet.write(offset, agent_key_column, key)
        worksheet.write(offset, agent_name_column, range_name)
    _define_range(workbook, "PCS_AGENT_LOOKUP_KEYS", agent_key_column, 1, len(agent_mappings))
    _define_range(workbook, "PCS_AGENT_LOOKUP_NAMES", agent_name_column, 1, len(agent_mappings))

    workbook.define_name("PCS_Period", "='OVERVIEW'!$C$2")
    workbook.define_name("PCS_LOB", "='OVERVIEW'!$J$2")
    workbook.define_name("PCS_TL", "='OVERVIEW'!$Q$2")
    workbook.define_name("PCS_Agent", "='OVERVIEW'!$X$2")
    workbook.define_name(
        "PCS_SELECTED_TL_LIST",
        "=INDIRECT(INDEX(PCS_TL_LOOKUP_NAMES,MATCH(PCS_LOB,PCS_TL_LOOKUP_KEYS,0)))",
    )
    workbook.define_name(
        "PCS_SELECTED_AGENT_LIST",
        "=INDIRECT(INDEX(PCS_AGENT_LOOKUP_NAMES,MATCH(PCS_LOB&\"|\"&PCS_TL,PCS_AGENT_LOOKUP_KEYS,0)))",
    )
    for column, source, prompt in (
        (2, "=PCS_PERIOD_LIST", "Choose the reporting period"),
        (9, "=PCS_LOB_LIST", "Choose a LOB, then continue left to right"),
        (16, "=PCS_SELECTED_TL_LIST", "Choose a Team Leader for the selected LOB"),
        (23, "=PCS_SELECTED_AGENT_LIST", "Choose an Agent for the selected team"),
    ):
        worksheet.data_validation(1, column, 1, column, {
            "validate": "list", "source": source,
            "input_title": "PCS filter", "input_message": prompt,
            "error_title": "Invalid filter", "error_message": "Select a value from the list.",
        })

    for column, header in enumerate(DASHBOARD_HEADERS, dashboard_start_column):
        worksheet.write(0, column, header)
    for row_number, values in enumerate(rows, 1):
        for column, value in enumerate(values, dashboard_start_column):
            worksheet.write(row_number, column, value)
    last_dashboard_row = len(rows)
    dashboard_names = {
        "PCS_DASH_KEY": 0, "PCS_DASH_PERIOD": 1, "PCS_DASH_LEVEL": 2,
        "PCS_DASH_LOB": 3, "PCS_DASH_TL": 4, "PCS_DASH_AGENT": 5,
        "PCS_DASH_DISPLAY": 6, "PCS_DASH_AGENT_COUNT": 7,
        "PCS_DASH_PCS": 8, "PCS_DASH_PRIOR_PCS": 9,
        "PCS_DASH_CHANGE": 10, "PCS_DASH_PART": 11,
        "PCS_DASH_PRIOR_PART": 12, "PCS_DASH_VALID": 13,
        "PCS_DASH_LOW": 14, "PCS_DASH_DUE": 15,
    }
    for name, offset in dashboard_names.items():
        _define_range(
            workbook, name, dashboard_start_column + offset, 1, last_dashboard_row,
        )

    selected_index_column = 39
    current_key = "Current MTD|All|All|All"
    selected_index = next(
        (index for index, row in enumerate(rows, 1) if row[0] == current_key), 0,
    )
    worksheet.write_formula(
        1, selected_index_column,
        '=IFERROR(MATCH(PCS_Period&"|"&PCS_LOB&"|"&PCS_TL&"|"&PCS_Agent,PCS_DASH_KEY,0),0)',
        None, selected_index,
    )
    workbook.define_name("PCS_Selected_Row", "='OVERVIEW'!$AN$2")
    selected = rows[selected_index - 1] if selected_index else (None,) * len(DASHBOARD_HEADERS)
    for start, name, cached, fmt in (
        (0, "PCS_DASH_PCS", selected[8], formats.card_decimal),
        (7, "PCS_DASH_PART", selected[11], formats.card_percent),
        (14, "PCS_DASH_PRIOR_PCS", selected[9], formats.card_decimal),
        (21, "PCS_DASH_CHANGE", selected[10], formats.card_decimal),
    ):
        worksheet.write_formula(
            5, start,
            f'=IF(PCS_Selected_Row=0,"",INDEX({name},PCS_Selected_Row))',
            fmt, "" if cached is None else cached,
        )

    desired_level = 'IF(PCS_Agent<>"All","AGENT",IF(PCS_TL<>"All","AGENT",IF(PCS_LOB<>"All","TEAM","LOB")))'
    action_index_column = 38
    initial_indexes = [
        index for index, row in enumerate(rows, 1)
        if row[1] == "Current MTD" and row[2] == "LOB"
    ][:8]
    field_specs = (
        ("PCS_DASH_DISPLAY", 6, "text"),
        ("PCS_DASH_LOB", 3, "text"),
        ("PCS_DASH_TL", 4, "text"),
        ("PCS_DASH_AGENT_COUNT", 7, "integer"),
        ("PCS_DASH_PART", 11, "percent"),
        ("PCS_DASH_PCS", 8, "decimal"),
        ("PCS_DASH_DUE", 15, "alert"),
    )
    for offset in range(8):
        helper_row = 1 + offset
        visible_row = 24 + offset
        cached_index = initial_indexes[offset] if offset < len(initial_indexes) else ""
        rank = offset + 1
        index_formula = (
            '=IFERROR(AGGREGATE(15,6,(ROW(PCS_DASH_PERIOD)-MIN(ROW(PCS_DASH_PERIOD))+1)/'
            f'((PCS_DASH_PERIOD=PCS_Period)*(PCS_DASH_LEVEL={desired_level})*'
            '(IF(PCS_LOB="All",1,--(PCS_DASH_LOB=PCS_LOB)))*'
            '(IF(PCS_TL="All",1,--(PCS_DASH_TL=PCS_TL)))*'
            '(IF(PCS_Agent="All",1,--(PCS_DASH_AGENT=PCS_Agent)))), '
            f'{rank}),"")'
        )
        worksheet.write_formula(
            helper_row, action_index_column, index_formula, None, cached_index,
        )
        helper_ref = f"${xl_col_to_name(action_index_column)}${helper_row + 1}"
        cached_row = rows[cached_index - 1] if isinstance(cached_index, int) else None
        for field, (name, source_index, kind) in enumerate(field_specs):
            cached = cached_row[source_index] if cached_row is not None else ""
            value_format = {
                "integer": formats.table_integer,
                "percent": formats.table_percent,
                "decimal": formats.table_decimal,
                "alert": formats.due if cached not in {0, "", None} else formats.clear,
            }.get(kind, formats.table_text)
            worksheet.write_formula(
                visible_row, field * 4,
                f'=IF({helper_ref}="","",INDEX({name},{helper_ref}))',
                value_format, "" if cached is None else cached,
            )
        # Both charts use the same selected-scope ranking as the action grid.
        for column, name, source_index in (
            (29, "PCS_DASH_DISPLAY", 6),
            (30, "PCS_DASH_PCS", 8),
            (31, "PCS_DASH_PRIOR_PCS", 9),
            (35, "PCS_DASH_DISPLAY", 6),
            (36, "PCS_DASH_PART", 11),
            (37, "PCS_DASH_PRIOR_PART", 12),
        ):
            cached = cached_row[source_index] if cached_row is not None else ""
            worksheet.write_formula(
                helper_row, column,
                f'=IF({helper_ref}="","",INDEX({name},{helper_ref}))',
                None, "" if cached is None else cached,
            )
    worksheet.conditional_format(24, 24, 31, 24, {
        "type": "cell", "criteria": ">", "value": 0, "format": formats.due,
    })
    worksheet.set_column(29, dashboard_start_column + len(DASHBOARD_HEADERS) - 1, None, None, {"hidden": True})


def _add_overview(
    book: DecisionWorkbook,
    status: str,
    status_note: str,
    latest: date,
    dashboard_rows: Sequence[Sequence[Any]],
    target: float | None,
) -> None:
    worksheet = book.report.workbook.add_worksheet("OVERVIEW")
    selected = next(
        (row for row in dashboard_rows if row[0] == "Current MTD|All|All|All"),
        (None,) * len(DASHBOARD_HEADERS),
    )
    initial_rows = [
        row for row in dashboard_rows
        if row[1] == "Current MTD" and row[2] == "LOB"
    ][:8]
    padded = [*initial_rows, *((None,) * len(DASHBOARD_HEADERS) for _ in range(8 - len(initial_rows)))]
    display_status = "DATA FRESH" if status in {"LIVE", "FINAL"} else "CHECK DATA"
    target_text = f" · TARGET {target:.2f}" if target is not None else ""
    formats = render_v2_dashboard(
        book.report.workbook, worksheet,
        title="PCS OPERATIONS",
        filters=(
            ("Period", "Current MTD"),
            ("LOB", "All"),
            ("Team Leader", "All"),
            ("Agent", "All"),
        ),
        kpis=(
            ("Current PCS", selected[8], "decimal"),
            ("Participation", selected[11], "percent"),
            ("Prior PCS", selected[9], "decimal"),
            ("Change", selected[10], "decimal"),
        ),
        left_chart=V2ChartSpec(
            f"PCS CURRENT VS PRIOR{target_text}", "bar",
            [row[6] or "" for row in padded],
            (
                ("Selected period", [row[8] for row in padded], COLORS["teal"]),
                ("Prior comparable", [row[9] for row in padded], COLORS["muted"]),
            ),
            minimum=1, maximum=5,
        ),
        right_chart=V2ChartSpec(
            "PARTICIPATION CURRENT VS PRIOR", "bar",
            [row[6] or "" for row in padded],
            (
                ("Selected period", [row[11] for row in padded], COLORS["teal"]),
                ("Prior comparable", [row[12] for row in padded], COLORS["muted"]),
            ),
            value_kind="percent", minimum=0, maximum=1,
        ),
        action_title="SELECTED SCOPE PERFORMANCE",
        action_headers=(
            "SCOPE", "LOB", "TEAM LEADER", "PCS AGENTS",
            "PARTICIPATION", "PCS", "COACHING DUE",
        ),
        action_rows=[
            (row[6], row[3], row[4], row[7], row[11], row[8], row[15])
            for row in initial_rows
        ],
        action_kinds=("text", "text", "text", "integer", "percent", "decimal", "alert"),
        status=display_status,
        status_kind=status,
        status_note=f"Data through {latest:%Y-%m-%d}. {status_note}",
    )
    _add_dashboard_interactivity(book, worksheet, formats, dashboard_rows)


def _build_queue_rows(
    source_rows: Sequence[Sequence[Any]],
    actions: Mapping[str, Mapping[str, Any]],
    duplicates: set[str],
) -> list[tuple[Any, ...]]:
    output = []
    for source in source_rows:
        source_record = dict(zip(PCS_COACHING_HEADERS, source))
        key = str(source_record.get("Coaching Key") or "").strip()
        action = actions.get(key, {})
        status = "DUPLICATE KEY IN LOG" if key in duplicates else str(
            action.get("Coaching Status") or "Pending",
        )
        output.append((
            *(source_record.get(header) for header in COACHING_LOG_HEADERS[:13]),
            source_record.get("Priority"), status, action.get("Coach"),
            action.get("Coaching Date"), action.get("Due Date"),
            action.get("Coaching Comment"),
        ))
    return output


def build_pcs_snapshot_workbook(
    conn: DatabaseConnection,
    config: Config,
    start: date,
    end: date,
    output: Path | None = None,
) -> Path:
    """Generate one self-contained PCS workbook without opening Excel."""

    from .decision_products import _audit_rows, _pcs_coaching_rows, _source_state

    generated = datetime.now()
    target = (
        output.resolve() if output is not None else
        (config.reports / f"{PCS_REPORT_PREFIX}{generated:%Y-%m-%d %H%M%S}.xlsx").resolve()
    )
    if target.exists() and output is None:
        target = target.with_name(
            f"{PCS_REPORT_PREFIX}{generated:%Y-%m-%d %H%M%S_%f}.xlsx",
        )
    target.parent.mkdir(parents=True, exist_ok=True)
    partial = target.with_name(f"{target.stem}.partial{target.suffix}")

    latest_value = conn.execute("SELECT max(business_date) FROM mart.agent_pcs_day").fetchone()[0]
    if latest_value is None:
        raise RuntimeError("No PCS data exists. Add a valid Call-by-Call extract first.")
    latest = (
        latest_value.date() if isinstance(latest_value, datetime)
        else latest_value if isinstance(latest_value, date)
        else date.fromisoformat(str(latest_value)[:10])
    )
    metric_catalog = load_metric_catalog(config.home, config.metric_catalog)
    pcs_method = metric_catalog.method_for("pcs_average", latest, {})
    minimum_sample = int(pcs_method.minimum_sample) if pcs_method is not None else 1
    rulebook = load_rulebook(config.home, config.business_rules)
    month_index = latest.year * 12 + latest.month - 1
    first_index = month_index - (config.pcs_tracker.history_months - 1)
    data_start = min(start, date(first_index // 12, first_index % 12 + 1, 1))

    lob_rows = pcs_lob_scorecard_rows(conn, latest, minimum_sample, generated)
    result_rows = pcs_result_rows(conn, latest, minimum_sample, generated)
    trend_rows = _daily_trend_rows(conn, latest)
    action_headers, raw_actions = _pcs_coaching_rows(conn, config, data_start, latest)
    indexes = {header: index for index, header in enumerate(action_headers)}
    sources = {
        "LOB": "lob", "Team Leader": "team_leader", "Agent Selector": "agent_selector",
        "Agent": "agent_name", "Agent ID": "agent_id", "Priority": "priority",
        "Date": "business_date", "Call Start": "call_start", "Q1 Score": "q1_score",
        "Customer Comment": "customer_comment", "Call Reference Number": "call_reference_number",
        "Call ID": "call_id", "Language": "language", "Coaching Key": "coaching_key",
    }
    queue_source = [
        tuple(values[indexes[sources[header]]] for header in PCS_COACHING_HEADERS)
        for values in raw_actions
    ]
    log_path = ensure_coaching_log(config, queue_source)
    coaching_rows, duplicates = _read_coaching_rows(log_path)
    coaching = _coaching_action_map(coaching_rows)
    queue_rows = _build_queue_rows(queue_source, coaching, duplicates)
    dashboard_rows = _dashboard_rows(conn, latest, queue_source, coaching)
    status, status_note = _source_state(conn, ("fte", "calls"), latest)

    book = DecisionWorkbook(
        partial, config, "pcs", "PCS OPERATIONAL REPORT", data_start, latest, generated,
    )
    try:
        _add_overview(
            book, status, status_note, latest, dashboard_rows,
            pcs_method.target if pcs_method is not None else None,
        )
        lob_summary = [
            (row[0], row[5], row[6], row[7], row[8], row[10], row[13], row[16])
            for row in lob_rows
        ]
        lob_sheet = book.table(
            "LOB_SUMMARY", "PCS LOB SUMMARY",
            "Current MTD and prior-month same-days results. Values are calculated by Python before the workbook is created.",
            list(LOB_SUMMARY_HEADERS), lob_summary,
        )
        lob_sheet.conditional_format("H5:H1004", {
            "type": "text", "criteria": "containing", "value": "LOW SAMPLE",
            "format": book.report.error,
        })
        book.table(
            "DAILY_TREND", "PCS DAILY TREND",
            "Current-month daily PCS against the matching day of the previous month.",
            list(DAILY_TREND_HEADERS), trend_rows,
        )
        result_sheet = book.table(
            "RESULTS", "PCS TEAM AND AGENT RESULTS",
            "Filter from left to right: Period View, Scope Level, LOB, Team Leader and Agent. Values are final; no recalculation is required.",
            list(PCS_SNAPSHOT_RESULTS_HEADERS), result_rows,
        )
        if pcs_method is not None and pcs_method.target is not None:
            result_sheet.conditional_format("K5:K100004", {
                "type": "cell", "criteria": "<", "value": pcs_method.target,
                "format": book.report.error,
            })
        queue_sheet = book.table(
            "COACHING_QUEUE", "PCS COACHING OPPORTUNITIES",
            f"Filter the queue, copy columns A:M into {PCS_COACHING_LOG_FILENAME}, then complete its blue action fields. Call ID opens the exact call.",
            list(COACHING_QUEUE_HEADERS), queue_rows,
        )
        queue_sheet.conditional_format("N5:N100004", {
            "type": "text", "criteria": "containing", "value": "HIGH",
            "format": book.report.error,
        })
        queue_sheet.conditional_format("O5:O100004", {
            "type": "text", "criteria": "containing", "value": "DUPLICATE",
            "format": book.report.error,
        })
        coaching_values = [
            tuple(row.get(header) for header in COACHING_LOG_HEADERS)
            for row in coaching_rows
        ]
        book.table(
            "COACHING", "PCS COACHING STATUS",
            f"Read-only snapshot from {PCS_COACHING_LOG_FILENAME}. Edit the separate permanent log, then build a fresh PCS report.",
            list(COACHING_LOG_HEADERS), coaching_values,
        )
        _headers, data_rows = _query(
            conn,
            """SELECT lob, team_leader,
                      coalesce(agent_name,'Agent') || ' [' || agent_id || ']',
                      agent_id, agent_name, business_date, ops_manager, language,
                      inbound_calls, pcs_status_calls, pcs_participation_responses,
                      survey_responses, pcs_score_sum, pcs_average,
                      pcs_participation_rate, low_score_responses,
                      top_box_responses, pcs_invalid_responses,
                      CASE WHEN survey_responses<? THEN 'LOW_SAMPLE' ELSE 'OK' END,
                      agent_id || '|' || business_date, ?, ?, ?, ?, ?, ?
               FROM mart.agent_pcs_day WHERE business_date BETWEEN ? AND ?
               ORDER BY business_date, lob, team_leader, agent_name""",
            [
                minimum_sample, latest, generated, rulebook.version, rulebook.sha256,
                metric_catalog.version, metric_catalog.sha256, data_start, latest,
            ],
        )
        book.table(
            "PCS_DATA", "PCS CLEAN AGENT-DAY DATA",
            "One row per in-scope agent and day. Use native filters or build optional pivots from this static table.",
            list(PCS_SNAPSHOT_DATA_HEADERS), data_rows,
        )
        book.table(
            "HELP", "PCS OPERATING GUIDE",
            "The report is generated entirely by Python. Excel is only the viewing and filtering tool.",
            ["Role", "Action", "Where", "Important"],
            [
                ("WFM", "Paste the latest untouched FTE and Call-by-Call extracts, then choose Build latest PCS report", "WFMHub", "No workbook needs to be closed"),
                ("WFM", "Open or send the newest timestamped report", "Reports", "Every report is self-contained"),
                ("Quality / TL", "Filter COACHING_QUEUE and copy columns A:M for the selected call", "Generated report", "Coaching Key and Call ID identify the exact call"),
                ("Quality / TL", "Paste A:M and complete status, coach, dates and comment", PCS_COACHING_LOG_FILENAME, "This is the only permanent editable file"),
                ("WFM", "Build again from the current database after coaching updates", "PCS menu", "Fast path: source extracts are not scanned again"),
            ],
        )
        book.definitions([
            ("PCS Average", "Sum of valid inbound Q1 scores / valid inbound Q1 responses", "Weighted LOB, team and agent result", "Never average agent averages"),
            ("PCS Participation", "Inbound raw Q1 nonblank / inbound PCSStatus=1", "Survey participation opportunity", "Invalid nonblank Q1 remains in the numerator"),
            ("Score <= 3", "Count of valid inbound Q1 responses at or below 3", "Coaching opportunity", "One exact call equals one Coaching Key"),
            ("Current MTD", "First day of the latest data month through the latest data date", "Current operating result", "Driven by the data-through date"),
            ("Previous MTD same days", "Previous month through the comparable day number", "Fair MTD comparison", "Month length is capped safely"),
            ("Python snapshot", "All cards, charts and tables contain final values", "No refresh dependency", "Excel performs no KPI calculation"),
        ])
        audit = _audit_rows(conn, config, "pcs", data_start, latest, (
            ("Workbook engine", "Python-only snapshot", "No Power Query, Excel automation, Data Model or dashboard formulas"),
            ("Snapshot version", PCS_SNAPSHOT_VERSION, "Generated workbook contract"),
            ("Coaching log", str(log_path), "Permanent human-owned workbook; read only by WFMHub"),
            ("Duplicate Coaching Keys", len(duplicates), "Duplicate keys do not increase completed coaching counts"),
        ))
        book.audit(audit)
        book.close()
        partial.replace(target)
    except Exception:
        partial.unlink(missing_ok=True)
        raise
    return target
