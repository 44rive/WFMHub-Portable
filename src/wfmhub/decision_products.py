"""Focused WFM/Operations report products using one workbook design contract."""

from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any, Iterable, Sequence

from openpyxl import load_workbook

from .config import Config
from .database import DatabaseConnection
from .mapping import load_queue_mapping
from .metrics import MetricCatalog, evaluate_metric, load_metric_catalog
from .pcs_excel import PCS_TEMPLATE_VERSION
from .report_packs import archive_superseded_reports, publish_report, report_current_path
from .reports import COLORS, _query
from .rules import load_rulebook
from .service_profiles import ServiceProfile, load_service_profiles
from .template_reports import DecisionWorkbook, KpiCard, ModelTable


@dataclass(frozen=True)
class NamedPeriod:
    label: str
    start: date
    end: date


def _month_start(value: date) -> date:
    return value.replace(day=1)


def _previous_month(value: date) -> tuple[date, date]:
    end = value.replace(day=1) - timedelta(days=1)
    return end.replace(day=1), end


def _comparison_periods(start: date, end: date) -> list[NamedPeriod]:
    previous_start, previous_end = _previous_month(end)
    prior_mtd_end = min(previous_end, previous_start + timedelta(days=end.day - 1))
    duration = (end - start).days + 1
    prior_equal_end = start - timedelta(days=1)
    prior_equal_start = prior_equal_end - timedelta(days=duration - 1)
    periods = [
        NamedPeriod("Latest day", end, end),
        NamedPeriod("Selected period", start, end),
        NamedPeriod("Current MTD", _month_start(end), end),
        NamedPeriod("Prior-month same days", previous_start, prior_mtd_end),
        NamedPeriod("Previous full month", previous_start, previous_end),
    ]
    if start != _month_start(end):
        periods.insert(2, NamedPeriod("Previous equal period", prior_equal_start, prior_equal_end))
    return periods


def _ratio(numerator: float | int | None, denominator: float | int | None) -> float | None:
    return float(numerator) / float(denominator) if denominator else None


def _delta(value: float | None, reference: float | None, percent: bool = False) -> str:
    if value is None or reference is None:
        return "Comparison not available"
    difference = value - reference
    return f"{difference:+.1%} vs prior" if percent else f"{difference:+.2f} vs prior"


def _output_path(config: Config, key: str, start: date, end: date, generated: datetime, output: Path | None) -> Path:
    del start, end, generated
    return (output or report_current_path(config, key)).resolve()


def _source_state(conn: DatabaseConnection, families: Sequence[str], through: date, final: bool = False) -> tuple[str, str]:
    if not families:
        return ("FINAL" if final else "LIVE"), "All required datasets are available"
    placeholders = ",".join("?" for _ in families)
    rows = conn.execute(
        f"""SELECT source_family, newest_business_date, status
            FROM mart.source_health WHERE source_family IN ({placeholders})""",
        list(families),
    ).fetchall()
    found = {str(row[0]).lower(): row for row in rows}
    problems = []
    for family in families:
        if family.casefold() in {"start_end", "activities"}:
            variant = "START_END" if family.casefold() == "start_end" else "ACTIVITIES"
            variant_row = conn.execute(
                """SELECT max(r.schedule_date), count(*)
                   FROM raw.schedule_shift r
                   JOIN meta.source_file f ON f.file_id=r.source_file_id
                   WHERE f.active=true AND f.status='SUCCESS' AND f.source_variant=?""",
                [variant],
            ).fetchone()
            variant_date, variant_count = variant_row
            if family.casefold() == "start_end" and not variant_count:
                variant = "ACTIVITIES"
                variant_row = conn.execute(
                    """SELECT max(r.schedule_date), count(*)
                       FROM raw.schedule_shift r
                       JOIN meta.source_file f ON f.file_id=r.source_file_id
                       WHERE f.active=true AND f.status='SUCCESS'
                         AND f.source_variant='ACTIVITIES' AND r.parse_ok=true
                         AND r.scheduled_start IS NOT NULL
                         AND r.scheduled_end IS NOT NULL"""
                ).fetchone()
                variant_date, variant_count = variant_row
            if not variant_count:
                problems.append(f"{family}: no successful {variant} extract")
            elif variant_date is None or str(variant_date)[:10] < through.isoformat():
                problems.append(f"{family}: latest {variant_date or 'unknown'}")
            continue
        row = found.get(family.lower())
        if row is None:
            problems.append(f"{family}: no health record")
        elif str(row[2]).upper() != "SUCCESS":
            problems.append(f"{family}: {row[2]}")
        # FTE is a point-in-time scope roster rather than a dated fact source.
        # Its freshness is represented by load status and source hash, so a
        # NULL business date is expected and must not make every report red.
        elif family.casefold() != "fte" and (row[1] is None or str(row[1])[:10] < through.isoformat()):
            problems.append(f"{family}: latest {row[1] or 'unknown'}")
    if problems:
        return "INCOMPLETE", "; ".join(problems)
    if final:
        return "FINAL", f"Required sources loaded through {through}"
    if through >= date.today():
        return "PROVISIONAL", "Current-day values can still change before shifts and queues close"
    return "LIVE", f"Required sources loaded through {through}"


def _audit_rows(
    conn: DatabaseConnection,
    config: Config,
    report_key: str,
    start: date,
    end: date,
    extra: Iterable[Sequence[Any]] = (),
) -> list[Sequence[Any]]:
    latest = conn.execute(
        "SELECT run_id, finished_at, details FROM meta.refresh_run WHERE status='SUCCESS' ORDER BY finished_at DESC LIMIT 1"
    ).fetchone()
    rows: list[Sequence[Any]] = [
        ("Report", report_key, "WFM report product"),
        ("Selected period", f"{start} to {end}", "Dates included"),
        ("Last refreshed", datetime.now(), "Local work-machine time"),
        ("Refresh run", latest[0] if latest else None, latest[2] if latest else "No successful refresh metadata"),
        ("Prepared by", "Anass ASSRI", "WFM"),
    ]
    rows.extend(extra)
    return rows


def _atomic_book(
    config: Config,
    key: str,
    title: str,
    start: date,
    end: date,
    output: Path | None,
) -> tuple[DecisionWorkbook, Path, Path]:
    generated = datetime.now()
    target = _output_path(config, key, start, end, generated, output)
    target.parent.mkdir(parents=True, exist_ok=True)
    partial = target.with_name(f"{target.stem}.partial{target.suffix}")
    return DecisionWorkbook(partial, config, key, title, start, end, generated), partial, target


def _finish(book: DecisionWorkbook, partial: Path, target: Path) -> Path:
    try:
        book.close()
        publish_report(book.config, book.report_key, partial, target, book.generated)
    except Exception:
        partial.unlink(missing_ok=True)
        raise
    return target


def _pcs_aggregate(
    conn: DatabaseConnection,
    config: Config,
    period: NamedPeriod,
) -> tuple[Any, ...]:
    row = conn.execute(
        """SELECT coalesce(sum(pcs_score_sum),0), coalesce(sum(survey_responses),0),
                  coalesce(sum(pcs_participation_responses),0), coalesce(sum(pcs_status_calls),0),
                  coalesce(sum(low_score_responses),0), coalesce(sum(top_box_responses),0),
                  coalesce(sum(inbound_calls),0)
           FROM mart.agent_pcs_day WHERE business_date BETWEEN ? AND ?""",
        [period.start, period.end],
    ).fetchone()
    score_sum, valid, participants, eligible, low, high, inbound = row
    return (
        period.label, period.start, period.end, _ratio(score_sum, valid),
        _ratio(participants, eligible), valid, eligible, low, high, inbound,
    )


def _pcs_coaching_rows(
    conn: DatabaseConnection,
    config: Config,
    start: date,
    end: date,
) -> tuple[list[str], list[tuple[Any, ...]]]:
    """Return a self-contained coaching queue; manual fields stay in Excel."""

    primary = config.pcs.primary_score_question
    primary_score = f"question_{primary}_score"
    allowed_scores = ", ".join(f"{value:g}" for value in config.pcs.allowed_scores)
    return _query(
        conn,
        f"""SELECT coalesce(d.lob,c.lob) AS lob,
                   d.team_leader,
                   coalesce(d.canonical_name,c.agent_name) || ' [' || c.agent_id || ']' AS agent_selector,
                   coalesce(d.canonical_name,c.agent_name) AS agent_name,
                   c.agent_id,
                   CASE WHEN c.{primary_score}<=2 THEN 'HIGH' ELSE 'NORMAL' END AS priority,
                   c.business_date, c.call_start,
                   c.{primary_score} AS q1_score,
                   c.question_3 AS customer_comment,
                   c.call_reference_number,
                   coalesce(d.language,c.language) AS language,
                   c.call_key AS coaching_key,
                   'Pending' AS coaching_status,
                   NULL AS coach, NULL AS coaching_date, NULL AS due_date,
                   NULL AS coaching_comment,
                   d.ops_manager
            FROM core.clean_call_leg c
            LEFT JOIN core.dim_agent d ON d.agent_id=c.agent_id
            WHERE c.business_date BETWEEN ? AND ?
              AND upper(coalesce(c.call_direction,''))='I'
              AND c.{primary_score} IN ({allowed_scores})
              AND c.{primary_score} <= ?
            ORDER BY priority, c.business_date DESC, d.team_leader,
                     coalesce(d.canonical_name,c.agent_name), c.call_start""",
        [start, end, config.pcs.negative_score_maximum],
    )


def _pcs_sum_formula(column: str, from_name: str = "PCS_From", to_name: str = "PCS_To") -> str:
    scope = (
        f'(tblPcsData[Date]>={from_name})*(tblPcsData[Date]<={to_name})*'
        'IF(CONTROL!$K$6="All",1,--(tblPcsData[LOB]=CONTROL!$K$6))*'
        'IF(CONTROL!$N$6="All",1,--(tblPcsData[Team Leader]=CONTROL!$N$6))*'
        'IF(CONTROL!$Q$6="All",1,--(tblPcsData[Agent Selector]=CONTROL!$Q$6))'
    )
    return f"SUMPRODUCT({scope}*N(tblPcsData[{column}]))"


def _pcs_completed_formula() -> str:
    return (
        'IFERROR(ROWS(UNIQUE(FILTER(tblCoaching[Coaching Key],'
        '(tblCoaching[Coaching Key]<>"")*'
        '(tblCoaching[Date]>=PCS_From)*(tblCoaching[Date]<=PCS_To)*'
        'IF(CONTROL!$K$6="All",1,--(tblCoaching[LOB]=CONTROL!$K$6))*'
        'IF(CONTROL!$N$6="All",1,--(tblCoaching[Team Leader]=CONTROL!$N$6))*'
        'IF(CONTROL!$Q$6="All",1,--(tblCoaching[Agent Selector]=CONTROL!$Q$6))*'
        '--(tblCoaching[Coaching Status]="Completed")))),0)'
    )


def _add_pcs_dashboard(
    book: DecisionWorkbook,
    status: str,
    status_text: str,
    start: date,
    end: date,
    lobs: Sequence[str],
    team_leaders: Sequence[str],
    agents: Sequence[str],
    data_start: date,
    latest: date,
    trend_count: int,
    minimum_sample: int,
    default_values: dict[str, float | int | None],
) -> None:
    """Create the permanent selector cockpit driven by refreshed Excel tables."""

    wb = book.report.workbook
    wb.set_calc_mode("auto")
    ws = wb.add_worksheet("CONTROL")
    ws.hide_gridlines(2)
    ws.set_tab_color(COLORS["gold"])
    ws.set_zoom(85)
    ws.set_landscape()
    ws.fit_to_pages(1, 1)
    ws.freeze_panes(4, 0)
    ws.merge_range("A1:R1", "PCS  /  OPERATIONAL PERFORMANCE TRACKER", book.report.title)
    ws.merge_range(
        "A2:R2",
        "",
        book.report.subtitle,
    )
    ws.write_formula(
        "A2",
        '="Data through "&TEXT(PCS_Data_Through,"yyyy-mm-dd")&"  |  feed refreshed "&TEXT(PCS_Feed_Refreshed,"yyyy-mm-dd hh:mm")&"  |  prepared by Anass ASSRI"',
        book.report.subtitle,
        f"Data through {latest:%Y-%m-%d}",
    )
    badge = status if status in book.badge_formats else "INCOMPLETE"
    ws.merge_range("A4:R4", "", book.badge_formats[badge])
    ws.write_formula(
        "A4",
        '=IF(UPPER(TRIM(XLOOKUP("Power Query Installed",tblSetup[Setting],tblSetup[Value],"NO")))="YES",'
        '"OPERATIONAL  /  REFRESH ALL UPDATES DATA; COACHING IS PRESERVED",'
        '"SETUP REQUIRED  /  OPEN SETUP AND HELP FOR THE ONE-TIME POWER QUERY LINK")',
        book.badge_formats[badge],
        f"{badge} / {status_text}",
    )

    selector_label = wb.add_format({
        "font_name": "Aptos", "font_size": 8, "bold": True,
        "font_color": COLORS["muted"], "bg_color": COLORS["canvas"],
        "align": "left", "valign": "vcenter", "indent": 1,
    })
    selector = wb.add_format({
        "font_name": "Aptos Display", "font_size": 11, "bold": True,
        "font_color": COLORS["dark"], "bg_color": COLORS["white"],
        "border": 1, "border_color": COLORS["teal"], "align": "left",
        "valign": "vcenter", "indent": 1, "num_format": "yyyy-mm-dd",
    })
    for label, label_range, value_range, value in (
        ("PERIOD VIEW", "A5:C5", "A6:C7", "Current MTD"),
        ("CUSTOM FROM", "E5:F5", "E6:F7", start),
        ("CUSTOM TO", "H5:I5", "H6:I7", end),
        ("LOB", "K5:L5", "K6:L7", "All"),
        ("TEAM LEADER", "N5:O5", "N6:O7", "All"),
        ("AGENT", "Q5:R5", "Q6:R7", "All"),
    ):
        ws.merge_range(label_range, label, selector_label)
        if isinstance(value, date):
            ws.merge_range(value_range, value, selector)
        else:
            ws.merge_range(value_range, value, selector)
    ws.set_row(5, 22)
    ws.set_row(6, 22)
    period_choices = [
        "Latest day", "Current week", "Previous week", "Current MTD",
        "Previous-month same days", "Previous full month", "Custom period",
    ]
    ws.data_validation("A6", {"validate": "list", "source": period_choices})
    ws.data_validation("E6", {"validate": "date", "criteria": "between", "minimum": date(2020, 1, 1), "maximum": date(2100, 12, 31)})
    ws.data_validation("H6", {"validate": "date", "criteria": "between", "minimum": date(2020, 1, 1), "maximum": date(2100, 12, 31)})

    del lobs, team_leaders, agents
    ws.data_validation("K6", {"validate": "list", "source": "=PCS_LOB_LIST"})
    ws.data_validation("N6", {"validate": "list", "source": "=PCS_TL_LIST"})
    ws.data_validation("Q6", {"validate": "list", "source": "=PCS_AGENT_LIST"})

    # Period selectors follow the latest date inside the refreshable PCS table.
    # Replacing PCS_DATA therefore advances Current week/MTD without rebuilding.
    wb.define_name("PCS_Latest", "=MAX(tblPcsData[Date])")
    wb.define_name("PCS_Data_Through", "=MAX(tblPcsData[Data Through])")
    wb.define_name("PCS_Feed_Refreshed", "=MAX(tblPcsData[Feed Refreshed At])")
    wb.define_name(
        "PCS_From",
        '=IF(CONTROL!$A$6="Latest day",PCS_Latest,'
        'IF(CONTROL!$A$6="Current week",PCS_Latest-WEEKDAY(PCS_Latest,2)+1,'
        'IF(CONTROL!$A$6="Previous week",PCS_Latest-WEEKDAY(PCS_Latest,2)-6,'
        'IF(CONTROL!$A$6="Current MTD",EOMONTH(PCS_Latest,-1)+1,'
        'IF(CONTROL!$A$6="Previous-month same days",EOMONTH(PCS_Latest,-2)+1,'
        'IF(CONTROL!$A$6="Previous full month",EOMONTH(PCS_Latest,-2)+1,CONTROL!$E$6))))))',
    )
    wb.define_name(
        "PCS_To",
        '=IF(CONTROL!$A$6="Latest day",PCS_Latest,'
        'IF(CONTROL!$A$6="Current week",PCS_Latest,'
        'IF(CONTROL!$A$6="Previous week",PCS_Latest-WEEKDAY(PCS_Latest,2),'
        'IF(CONTROL!$A$6="Current MTD",PCS_Latest,'
        'IF(CONTROL!$A$6="Previous-month same days",EDATE(PCS_Latest,-1),'
        'IF(CONTROL!$A$6="Previous full month",EOMONTH(PCS_Latest,-1),CONTROL!$H$6))))))',
    )
    wb.define_name("PCS_Prior_From", "=EOMONTH(PCS_Latest,-2)+1")
    wb.define_name("PCS_Prior_To", "=EDATE(PCS_Latest,-1)")

    scope_count = _pcs_sum_formula("Valid Q1")
    ws.merge_range("A8:R8", "", book.report.note)
    ws.write_formula(
        "A8", f'=IF({scope_count}=0,"NO MATCHING PCS DATA - CHECK THE SELECTORS",'
        f'"Showing "&TEXT(PCS_From,"yyyy-mm-dd")&" to "&TEXT(PCS_To,"yyyy-mm-dd")&"  |  choose filters from left to right: LOB, Team Leader, Agent")',
        book.report.note, "Showing current MTD",
    )

    score_sum = _pcs_sum_formula("Q1 Score Sum")
    valid = _pcs_sum_formula("Valid Q1")
    participating = _pcs_sum_formula("Q1 Nonblank")
    eligible = _pcs_sum_formula("PCS Status 1")
    low = _pcs_sum_formula("Score <= 3")
    positive = _pcs_sum_formula("Score > 3")
    inbound = _pcs_sum_formula("Inbound Call Legs")
    completed = _pcs_completed_formula()
    cards = [
        ("PCS AVERAGE", f'=IFERROR({score_sum}/{valid},"")', book.card_decimal, "Weighted score / valid responses", default_values.get("pcs_average")),
        ("PARTICIPATION", f'=IFERROR({participating}/{eligible},"")', book.card_percent, "Q1 nonblank / PCS Status 1", default_values.get("participation")),
        ("VALID RESPONSES", f"={valid}", book.card_integer, f"Low sample below {minimum_sample}", default_values.get("valid")),
        ("INBOUND CALL LEGS", f"={inbound}", book.card_integer, "Inbound legs in selected scope", default_values.get("inbound")),
        ("SCORE <= 3", f"={low}", book.card_integer, "Coaching opportunities", default_values.get("low")),
        ("POSITIVE > 3", f"={positive}", book.card_integer, "Positive valid responses", default_values.get("positive")),
        ("COACHING COMPLETED", f"={completed}", book.card_integer, "Updates as the team fills COACHING", 0),
        ("ACTIONS RATE", f'=IFERROR({completed}/{low},"")', book.card_percent, "Completed / score <= 3", 0),
    ]
    for index, (label, formula, fmt, note, cached) in enumerate(cards):
        row = 9 if index < 4 else 14
        column = (index % 4) * 4
        ws.merge_range(row, column, row, column + 2, label, book.report.kpi_label)
        ws.merge_range(row + 1, column, row + 2, column + 2, "", fmt)
        ws.write_formula(row + 1, column, formula, fmt, cached if cached is not None else "")
        ws.merge_range(row + 3, column, row + 3, column + 2, note, book.card_compare)
        ws.set_row(row + 1, 26)
        ws.set_row(row + 2, 26)

    table_row = 20
    ws.merge_range(table_row, 0, table_row, 9, "PERIOD BENCHMARK", book.report.section)
    compare_headers = ["Period", "Start", "End", "PCS Average", "Participation %", "Valid Responses", "PCS Status 1", "Score <= 3", "Score > 3", "Inbound Legs"]
    for column, header in enumerate(compare_headers):
        ws.write(table_row + 2, column, header, book.report.header)
    current_values = [
        "Selected scope", "=PCS_From", "=PCS_To",
        f'=IFERROR({_pcs_sum_formula("Q1 Score Sum")}/{_pcs_sum_formula("Valid Q1")},"")',
        f'=IFERROR({_pcs_sum_formula("Q1 Nonblank")}/{_pcs_sum_formula("PCS Status 1")},"")',
        f'={_pcs_sum_formula("Valid Q1")}', f'={_pcs_sum_formula("PCS Status 1")}',
        f'={_pcs_sum_formula("Score <= 3")}', f'={_pcs_sum_formula("Score > 3")}',
        f'={_pcs_sum_formula("Inbound Call Legs")}',
    ]
    prior_values = [
        "Previous-month same days", "=PCS_Prior_From", "=PCS_Prior_To",
        f'=IFERROR({_pcs_sum_formula("Q1 Score Sum", "PCS_Prior_From", "PCS_Prior_To")}/{_pcs_sum_formula("Valid Q1", "PCS_Prior_From", "PCS_Prior_To")},"")',
        f'=IFERROR({_pcs_sum_formula("Q1 Nonblank", "PCS_Prior_From", "PCS_Prior_To")}/{_pcs_sum_formula("PCS Status 1", "PCS_Prior_From", "PCS_Prior_To")},"")',
        f'={_pcs_sum_formula("Valid Q1", "PCS_Prior_From", "PCS_Prior_To")}',
        f'={_pcs_sum_formula("PCS Status 1", "PCS_Prior_From", "PCS_Prior_To")}',
        f'={_pcs_sum_formula("Score <= 3", "PCS_Prior_From", "PCS_Prior_To")}',
        f'={_pcs_sum_formula("Score > 3", "PCS_Prior_From", "PCS_Prior_To")}',
        f'={_pcs_sum_formula("Inbound Call Legs", "PCS_Prior_From", "PCS_Prior_To")}',
    ]
    for offset, values in enumerate((current_values, prior_values)):
        for column, value in enumerate(values):
            fmt = book.report.date if column in {1, 2} else book.report.percent if column == 4 else book.report.decimal if column == 3 else book.report.integer if column >= 5 else book.report.body
            if isinstance(value, str) and value.startswith("="):
                ws.write_formula(table_row + 3 + offset, column, value, fmt)
            else:
                ws.write(table_row + 3 + offset, column, value, fmt)

    chart = wb.add_chart({"type": "line"})
    for name, column, color, secondary in (
        ("PCS Average", 5, COLORS["teal"], False),
        ("Participation", 6, COLORS["gold"], True),
    ):
        chart.add_series({
            "name": name,
            "categories": ["_LOOKUPS", 1, 0, trend_count, 0],
            "values": ["_LOOKUPS", 1, column, trend_count, column],
            "line": {"color": color, "width": 2.25},
            "marker": {"type": "circle", "size": 4, "border": {"color": color}, "fill": {"color": COLORS["white"]}},
            "y2_axis": secondary,
        })
    chart.set_title({"name": "Daily PCS & participation"})
    chart.set_legend({"position": "bottom"})
    chart.set_chartarea({"border": {"none": True}, "fill": {"color": COLORS["white"]}})
    chart.set_plotarea({"border": {"none": True}, "fill": {"color": COLORS["white"]}})
    chart.set_y_axis({"min": 1, "max": 5, "major_unit": 1, "major_gridlines": {"visible": False}, "name": "PCS score"})
    chart.set_y2_axis({"min": 0, "max": 1, "major_unit": 0.2, "num_format": "0%", "name": "Participation"})
    ws.insert_chart("L21", chart, {"x_scale": 1.15, "y_scale": 1.05})

    note_row = table_row + 11
    ws.merge_range(note_row, 0, note_row, 17, "HOW TO USE THIS PAGE", book.report.section)
    ws.merge_range(note_row + 1, 0, note_row + 1, 17, "1. Choose a period and your Team Leader name. Every KPI, benchmark and chart on this page follows those selectors.", book.report.note)
    ws.merge_range(note_row + 2, 0, note_row + 2, 17, "2. Open AGENT_RESULTS for the realization list. Quality records completed follow-up in the blue COACHING columns.", book.report.note)
    ws.set_column("A:R", 11)
    ws.set_column("A:A", 13)
    ws.set_column("K:R", 13)


def _add_pcs_overview(book: DecisionWorkbook, minimum_sample: int, trend_count: int) -> None:
    """Add a clean management view driven by the permanent CONTROL selectors."""

    wb = book.report.workbook
    ws = wb.add_worksheet("OVERVIEW")
    ws.hide_gridlines(2)
    ws.set_tab_color(COLORS["gold"])
    ws.set_zoom(85)
    ws.freeze_panes(4, 0)
    ws.set_landscape()
    ws.fit_to_pages(1, 1)
    ws.merge_range("A1:R1", "PCS  /  MANAGEMENT OVERVIEW", book.report.title)
    ws.merge_range("A2:R2", "", book.report.subtitle)
    ws.write_formula(
        "A2",
        '="Selected "&TEXT(PCS_From,"yyyy-mm-dd")&" to "&TEXT(PCS_To,"yyyy-mm-dd")&'
        '"  |  "&CONTROL!$K$6&"  /  "&CONTROL!$N$6&"  /  "&CONTROL!$Q$6',
        book.report.subtitle,
        "Use CONTROL to choose period, LOB, Team Leader and Agent",
    )
    ws.write_url("A4", "internal:'CONTROL'!A1", book.report.editable, string="CHANGE FILTERS")
    ws.write_url("D4", "internal:'TEAM_VIEW'!A1", book.report.editable, string="TEAM REALISATIONS")
    ws.write_url("G4", "internal:'AGENT_RESULTS'!A1", book.report.editable, string="AGENT EXPLORER")
    ws.write_url("J4", "internal:'COACHING_WORKSPACE'!A1", book.report.editable, string="COACHING WORKSPACE")

    score_sum = _pcs_sum_formula("Q1 Score Sum")
    valid = _pcs_sum_formula("Valid Q1")
    participating = _pcs_sum_formula("Q1 Nonblank")
    eligible = _pcs_sum_formula("PCS Status 1")
    low = _pcs_sum_formula("Score <= 3")
    positive = _pcs_sum_formula("Score > 3")
    completed = _pcs_completed_formula()
    prior_score = _pcs_sum_formula("Q1 Score Sum", "PCS_Prior_From", "PCS_Prior_To")
    prior_valid = _pcs_sum_formula("Valid Q1", "PCS_Prior_From", "PCS_Prior_To")
    cards = [
        ("PCS AVERAGE", f'=IFERROR({score_sum}/{valid},"")', book.card_decimal, "Weighted result"),
        ("PARTICIPATION", f'=IFERROR({participating}/{eligible},"")', book.card_percent, "Q1 nonblank / eligible"),
        ("VALID RESPONSES", f"={valid}", book.card_integer, f"Sample warning below {minimum_sample}"),
        ("LOW SCORES", f"={low}", book.card_integer, "Valid Q1 score <= 3"),
        ("VS PRIOR MTD", f'=IFERROR({score_sum}/{valid}-{prior_score}/{prior_valid},"")', book.card_decimal, "Comparable movement"),
        ("POSITIVE SCORES", f"={positive}", book.card_integer, "Valid Q1 score > 3"),
        ("COACHING DONE", f"={completed}", book.card_integer, "Unique completed actions"),
        ("ACTION RATE", f'=IFERROR({completed}/{low},"")', book.card_percent, "Completed / opportunities"),
    ]
    for index, (label, formula, fmt, note) in enumerate(cards):
        row = 6 if index < 4 else 11
        column = (index % 4) * 4
        ws.merge_range(row, column, row, column + 2, label, book.report.kpi_label)
        ws.merge_range(row + 1, column, row + 2, column + 2, "", fmt)
        ws.write_formula(row + 1, column, formula, fmt, "")
        ws.merge_range(row + 3, column, row + 3, column + 2, note, book.card_compare)

    ws.merge_range("A17:I17", "PCS DAILY TREND", book.report.section)
    score_chart = wb.add_chart({"type": "line"})
    score_chart.add_series({
        "name": "PCS Average",
        "categories": ["_LOOKUPS", 1, 0, trend_count, 0],
        "values": ["_LOOKUPS", 1, 5, trend_count, 5],
        "line": {"color": COLORS["teal"], "width": 2.5},
        "marker": {"type": "circle", "size": 4, "fill": {"color": COLORS["white"]}},
    })
    score_chart.set_y_axis({"min": 1, "max": 5, "major_unit": 1, "major_gridlines": {"visible": False}})
    score_chart.set_legend({"none": True})
    score_chart.set_chartarea({"border": {"none": True}, "fill": {"color": COLORS["white"]}})
    score_chart.set_plotarea({"border": {"none": True}, "fill": {"color": COLORS["white"]}})
    ws.insert_chart("A18", score_chart, {"x_scale": 1.2, "y_scale": 1.0})

    ws.merge_range("J17:R17", "PARTICIPATION DAILY TREND", book.report.section)
    participation_chart = wb.add_chart({"type": "line"})
    participation_chart.add_series({
        "name": "Participation",
        "categories": ["_LOOKUPS", 1, 0, trend_count, 0],
        "values": ["_LOOKUPS", 1, 6, trend_count, 6],
        "line": {"color": COLORS["gold"], "width": 2.5},
        "marker": {"type": "circle", "size": 4, "fill": {"color": COLORS["white"]}},
    })
    participation_chart.set_y_axis({"min": 0, "max": 1, "major_unit": 0.2, "num_format": "0%", "major_gridlines": {"visible": False}})
    participation_chart.set_legend({"none": True})
    participation_chart.set_chartarea({"border": {"none": True}, "fill": {"color": COLORS["white"]}})
    participation_chart.set_plotarea({"border": {"none": True}, "fill": {"color": COLORS["white"]}})
    ws.insert_chart("J18", participation_chart, {"x_scale": 1.2, "y_scale": 1.0})

    ws.merge_range("A34:R34", "MANAGEMENT READING", book.report.section)
    ws.merge_range(
        "A35:R35",
        "Read PCS together with participation and valid sample. Use TEAM_VIEW to prioritize teams, then COACHING_WORKSPACE to record follow-up.",
        book.report.note,
    )
    ws.set_column("A:R", 11)
    ws.set_column("A:A", 14)


def _add_pcs_team_view(book: DecisionWorkbook, minimum_sample: int) -> None:
    """Create the novice-facing PCS view driven by the Dashboard selectors."""

    wb = book.report.workbook
    ws = wb.add_worksheet("TEAM_VIEW")
    ws.hide_gridlines(2)
    ws.set_tab_color(COLORS["gold"])
    ws.set_zoom(85)
    ws.freeze_panes(10, 0)
    ws.merge_range("A1:AA1", "PCS  /  TEAM REALISATIONS & COACHING", book.report.title)
    ws.merge_range(
        "A2:AA2",
        "Select period, LOB, Team Leader and Agent on CONTROL. This page follows the same selection automatically.",
        book.report.subtitle,
    )
    selector_label = wb.add_format({
        "font_name": "Aptos", "font_size": 8, "bold": True,
        "font_color": COLORS["muted"], "bg_color": COLORS["canvas"],
        "align": "left", "valign": "vcenter", "indent": 1,
    })
    selector_value = wb.add_format({
        "font_name": "Aptos Display", "font_size": 11, "bold": True,
        "font_color": COLORS["dark"], "bg_color": COLORS["white"],
        "border": 1, "border_color": COLORS["teal"], "align": "left",
        "valign": "vcenter", "indent": 1,
    })
    selectors = (
        ("PERIOD", "A4:C4", "A5:C6", "=CONTROL!$A$6"),
        ("LOB", "E4:G4", "E5:G6", "=CONTROL!$K$6"),
        ("TEAM LEADER", "I4:L4", "I5:L6", "=CONTROL!$N$6"),
        ("AGENT", "N4:R4", "N5:R6", "=CONTROL!$Q$6"),
    )
    for label, label_range, value_range, formula in selectors:
        ws.merge_range(label_range, label, selector_label)
        ws.merge_range(value_range, "", selector_value)
        first = value_range.split(":", 1)[0]
        ws.write_formula(first, formula, selector_value, "All")
    ws.write_url(
        "T5", "internal:'CONTROL'!A1", book.report.editable,
        string="CHANGE FILTERS ON CONTROL",
    )
    ws.merge_range(
        "A8:L8", "AGENT REALISATIONS", book.report.section,
    )
    agent_headers = [
        "LOB", "Team Leader", "Agent Selector", "Agent ID", "Language",
        "PCS Average", "Participation %", "Valid Q1", "PCS Status 1",
        "Score <= 3", "Prior MTD PCS", "Priority",
    ]
    for column, header in enumerate(agent_headers):
        ws.write(9, column, header, book.report.header)
    agent_formula = (
        '=LET(d,tblPcsData,'
        'm,(d[Date]>=PCS_From)*(d[Date]<=PCS_To)*'
        'IF(CONTROL!$K$6="All",1,--(d[LOB]=CONTROL!$K$6))*'
        'IF(CONTROL!$N$6="All",1,--(d[Team Leader]=CONTROL!$N$6))*'
        'IF(CONTROL!$Q$6="All",1,--(d[Agent Selector]=CONTROL!$Q$6)),'
        'pm,(d[Date]>=PCS_Prior_From)*(d[Date]<=PCS_Prior_To)*'
        'IF(CONTROL!$K$6="All",1,--(d[LOB]=CONTROL!$K$6))*'
        'IF(CONTROL!$N$6="All",1,--(d[Team Leader]=CONTROL!$N$6))*'
        'IF(CONTROL!$Q$6="All",1,--(d[Agent Selector]=CONTROL!$Q$6)),'
        'a,SORT(UNIQUE(FILTER(d[Agent Selector],m,""))),'
        'v,MAP(a,LAMBDA(x,SUMPRODUCT(m*(d[Agent Selector]=x)*N(d[Valid Q1])))),'
        's,MAP(a,LAMBDA(x,SUMPRODUCT(m*(d[Agent Selector]=x)*N(d[Q1 Score Sum])))),'
        'e,MAP(a,LAMBDA(x,SUMPRODUCT(m*(d[Agent Selector]=x)*N(d[PCS Status 1])))),'
        'p,MAP(a,LAMBDA(x,SUMPRODUCT(m*(d[Agent Selector]=x)*N(d[Q1 Nonblank])))),'
        'lo,MAP(a,LAMBDA(x,SUMPRODUCT(m*(d[Agent Selector]=x)*N(d[Score <= 3])))),'
        'pv,MAP(a,LAMBDA(x,SUMPRODUCT(pm*(d[Agent Selector]=x)*N(d[Valid Q1])))),'
        'ps,MAP(a,LAMBDA(x,SUMPRODUCT(pm*(d[Agent Selector]=x)*N(d[Q1 Score Sum])))),'
        'IFERROR(HSTACK('
        'XLOOKUP(a,d[Agent Selector],d[LOB],""),'
        'XLOOKUP(a,d[Agent Selector],d[Team Leader],""),a,'
        'XLOOKUP(a,d[Agent Selector],d[Agent ID],""),'
        'XLOOKUP(a,d[Agent Selector],d[Language],""),'
        'IFERROR(s/v,""),IFERROR(p/e,""),v,e,lo,IFERROR(ps/pv,""),'
        f'IF(v=0,"NO RESPONSE",IF(v<{minimum_sample},"LOW SAMPLE",IF(lo>0,"COACH","ON TRACK")))),'
        '"No matching agent data"))'
    )
    ws.write_dynamic_array_formula("A11", agent_formula, book.report.body, "Open in desktop Excel")
    ws.conditional_format(
        "L11:L1048576",
        {"type": "text", "criteria": "containing", "value": "COACH", "format": book.report.error},
    )
    ws.merge_range("N8:AA8", "COACHING OPPORTUNITIES", book.report.section)
    coaching_headers = [
        "LOB", "Team Leader", "Agent Selector", "Agent ID", "Priority", "Date",
        "Call Start", "Q1 Score", "Customer Comment", "Call Reference Number",
        "Coaching Key", "Action Status",
    ]
    for column, header in enumerate(coaching_headers, 13):
        ws.write(9, column, header, book.report.header)
    coaching_formula = (
        '=LET(q,tblCoachingQueue,'
        'm,(q[Date]>=PCS_From)*(q[Date]<=PCS_To)*'
        'IF(CONTROL!$K$6="All",1,--(q[LOB]=CONTROL!$K$6))*'
        'IF(CONTROL!$N$6="All",1,--(q[Team Leader]=CONTROL!$N$6))*'
        'IF(CONTROL!$Q$6="All",1,--(q[Agent Selector]=CONTROL!$Q$6)),'
        'x,FILTER(CHOOSECOLS(q,1,2,3,5,6,7,8,9,10,11,13),m),'
        'k,CHOOSECOLS(x,11),'
        'IFERROR(HSTACK(x,XLOOKUP(k,tblCoaching[Coaching Key],tblCoaching[Coaching Status],"Not started")),'
        '"No coaching opportunities in this selection"))'
    )
    ws.write_dynamic_array_formula("N11", coaching_formula, book.report.body, "Open in desktop Excel")
    ws.write_url(
        "T7", "internal:'COACHING'!A1", book.report.editable,
        string="OPEN PERMANENT COACHING LOG",
    )
    ws.set_column("A:A", 18)
    ws.set_column("B:B", 22)
    ws.set_column("C:C", 30)
    ws.set_column("D:E", 15)
    ws.set_column("F:K", 15)
    ws.set_column("L:L", 16)
    ws.set_column("M:M", 3)
    ws.set_column("N:AA", 18)
    ws.set_column("V:V", 34)
    ws.set_landscape()
    ws.fit_to_pages(1, 0)


def _add_pcs_agent_results(book: DecisionWorkbook, minimum_sample: int) -> None:
    """Add a refresh-aware full realization list for Microsoft 365 users."""

    wb = book.report.workbook
    ws = wb.add_worksheet("AGENT_RESULTS")
    ws.hide_gridlines(2)
    ws.set_tab_color(COLORS["teal"])
    ws.set_zoom(85)
    ws.freeze_panes(4, 0)
    ws.merge_range("A1:P1", "PCS  /  AGENT RESULTS", book.report.title)
    ws.merge_range(
        "A2:P2",
        "This list follows every CONTROL selector and expands automatically when refreshed PCS data contains new agents.",
        book.report.subtitle,
    )
    headers = [
        "LOB", "Team Leader", "Agent Selector", "Agent ID", "Agent", "Language",
        "Priority", "Latest Day Average", "Selected Period Average",
        "Prior MTD Average", "Movement", "Participation Rate", "Valid Q1",
        "PCS Status 1", "Score <= 3", "Next Action",
    ]
    for column, header in enumerate(headers):
        ws.write(3, column, header, book.report.header)
    formula = (
        '=LET(d,tblPcsData,'
        'm,(d[Date]>=PCS_From)*(d[Date]<=PCS_To)*'
        'IF(CONTROL!$K$6="All",1,--(d[LOB]=CONTROL!$K$6))*'
        'IF(CONTROL!$N$6="All",1,--(d[Team Leader]=CONTROL!$N$6))*'
        'IF(CONTROL!$Q$6="All",1,--(d[Agent Selector]=CONTROL!$Q$6)),'
        'pm,(d[Date]>=PCS_Prior_From)*(d[Date]<=PCS_Prior_To)*'
        'IF(CONTROL!$K$6="All",1,--(d[LOB]=CONTROL!$K$6))*'
        'IF(CONTROL!$N$6="All",1,--(d[Team Leader]=CONTROL!$N$6))*'
        'IF(CONTROL!$Q$6="All",1,--(d[Agent Selector]=CONTROL!$Q$6)),'
        'lm,d[Date]=PCS_Latest,'
        'a,SORT(UNIQUE(FILTER(d[Agent Selector],m))),'
        'v,MAP(a,LAMBDA(x,SUMPRODUCT(m*(d[Agent Selector]=x)*N(d[Valid Q1])))),'
        's,MAP(a,LAMBDA(x,SUMPRODUCT(m*(d[Agent Selector]=x)*N(d[Q1 Score Sum])))),'
        'e,MAP(a,LAMBDA(x,SUMPRODUCT(m*(d[Agent Selector]=x)*N(d[PCS Status 1])))),'
        'p,MAP(a,LAMBDA(x,SUMPRODUCT(m*(d[Agent Selector]=x)*N(d[Q1 Nonblank])))),'
        'lo,MAP(a,LAMBDA(x,SUMPRODUCT(m*(d[Agent Selector]=x)*N(d[Score <= 3])))),'
        'pv,MAP(a,LAMBDA(x,SUMPRODUCT(pm*(d[Agent Selector]=x)*N(d[Valid Q1])))),'
        'ps,MAP(a,LAMBDA(x,SUMPRODUCT(pm*(d[Agent Selector]=x)*N(d[Q1 Score Sum])))),'
        'lv,MAP(a,LAMBDA(x,SUMPRODUCT(lm*(d[Agent Selector]=x)*N(d[Valid Q1])))),'
        'ls,MAP(a,LAMBDA(x,SUMPRODUCT(lm*(d[Agent Selector]=x)*N(d[Q1 Score Sum])))),'
        'priority,IF(v=0,"NO RESPONSE",IF(v<' + str(minimum_sample) + ',"LOW SAMPLE",IF(lo>0,"COACH","ON TRACK"))),'
        'IFERROR(HSTACK('
        'XLOOKUP(a,d[Agent Selector],d[LOB],""),XLOOKUP(a,d[Agent Selector],d[Team Leader],""),a,'
        'XLOOKUP(a,d[Agent Selector],d[Agent ID],""),XLOOKUP(a,d[Agent Selector],d[Agent],""),'
        'XLOOKUP(a,d[Agent Selector],d[Language],""),priority,IFERROR(ls/lv,""),IFERROR(s/v,""),'
        'IFERROR(ps/pv,""),IFERROR(s/v-ps/pv,""),IFERROR(p/e,""),v,e,lo,'
        'IF(lo>0,"Open COACHING",IF(v=0,"Check sample","Monitor"))),"No matching agent data"))'
    )
    ws.write_dynamic_array_formula("A5", formula, book.report.body, "Open in desktop Excel")
    ws.conditional_format(
        "G5:G1048576",
        {"type": "text", "criteria": "containing", "value": "COACH", "format": book.report.error},
    )
    ws.set_column("A:B", 20)
    ws.set_column("C:C", 31)
    ws.set_column("D:F", 16)
    ws.set_column("G:P", 18)


def _add_pcs_coaching_actions(
    book: DecisionWorkbook,
    previous: Sequence[dict[str, Any]],
) -> None:
    """Create the permanent action ledger with one easy Coaching Key input."""

    wb = book.report.workbook
    ws = wb.add_worksheet("COACHING")
    ws.hide_gridlines(2)
    ws.set_tab_color(COLORS["gold"])
    ws.freeze_panes(4, 0)
    ws.set_zoom(85)
    headers = [
        "Coaching Key", "LOB", "Team Leader", "Agent Selector", "Agent ID",
        "Agent", "Date", "Call Start", "Q1 Score", "Customer Comment",
        "Call Reference Number", "Language", "Coaching Status", "Coach",
        "Coaching Date", "Due Date", "Coaching Comment",
    ]
    ws.merge_range("A1:Q1", "PCS  /  COACHING ACTIONS", book.report.title)
    ws.merge_range(
        "A2:Q2",
        "Go to the first blank row, select one Coaching Key, then fill only the blue action fields. Agent and call details fill automatically.",
        book.report.subtitle,
    )
    rows = list(previous) or [{}]
    editable = {
        "Coaching Key", "Coaching Status", "Coach", "Coaching Date",
        "Due Date", "Coaching Comment",
    }
    for row_index, item in enumerate(rows, 4):
        for column, header in enumerate(headers):
            key = header.casefold().replace(" ", "_")
            value = item.get(key)
            fmt = book.report.editable if header in editable else book.report.body
            if header in {"Coaching Date", "Due Date"}:
                fmt = book.report.editable_date
            ws.write(row_index, column, value, fmt)
    lookup_headers = {
        "LOB": "LOB", "Team Leader": "Team Leader",
        "Agent Selector": "Agent Selector", "Agent ID": "Agent ID",
        "Agent": "Agent", "Date": "Date", "Call Start": "Call Start",
        "Q1 Score": "Q1 Score", "Customer Comment": "Customer Comment",
        "Call Reference Number": "Call Reference Number", "Language": "Language",
    }
    columns = []
    for header in headers:
        column = {"header": header, "header_format": book.report.header}
        if header in lookup_headers:
            source = lookup_headers[header]
            column["formula"] = (
                '=IF([@[Coaching Key]]="","",'
                f'XLOOKUP([@[Coaching Key]],tblCoachingQueue[Coaching Key],tblCoachingQueue[{source}],"KEY NOT IN CURRENT QUEUE"))'
            )
            if header == "Date":
                column["format"] = book.report.date
            elif header == "Call Start":
                column["format"] = book.report.datetime
        columns.append(column)
    ws.add_table(3, 0, 3 + len(rows), len(headers) - 1, {
        "name": "tblCoaching",
        "style": "Table Style Light 9",
        "columns": columns,
    })
    book.tables.append(ModelTable(
        "COACHING", headers,
        [tuple(item.get(header.casefold().replace(" ", "_")) for header in headers) for item in rows],
    ))
    ws.data_validation("A5:A1004", {
        "validate": "list", "source": "=PCS_COACHING_KEY_LIST",
        "input_title": "Choose a PCS case",
        "input_message": "Select the exact call key. Details fill automatically.",
        "error_title": "Unknown Coaching Key",
        "error_message": "Choose a key from the refreshed coaching queue.",
    })
    ws.data_validation("M5:M1004", {
        "validate": "list",
        "source": ["Pending", "Planned", "Completed", "Not required"],
        "input_title": "Coaching status",
        "input_message": "Choose the current follow-up status.",
    })
    ws.conditional_format("A5:A1004", {
        "type": "duplicate", "format": book.report.error,
    })
    ws.conditional_format("M5:M1004", {
        "type": "text", "criteria": "containing", "value": "Pending",
        "format": book.report.error,
    })
    ws.write_url("A3", "internal:'COACHING_WORKSPACE'!A1", book.report.editable, string="VIEW OPPORTUNITIES")
    ws.set_column("A:A", 34)
    ws.set_column("B:C", 20)
    ws.set_column("D:D", 30)
    ws.set_column("E:F", 18)
    ws.set_column("G:I", 18)
    ws.set_column("J:J", 38)
    ws.set_column("K:L", 21)
    ws.set_column("M:P", 18)
    ws.set_column("Q:Q", 42)


def _add_pcs_coaching_workspace(book: DecisionWorkbook) -> None:
    """Show filtered coaching opportunities joined to the permanent action ledger."""

    wb = book.report.workbook
    ws = wb.add_worksheet("COACHING_WORKSPACE")
    ws.hide_gridlines(2)
    ws.set_tab_color(COLORS["gold"])
    ws.freeze_panes(7, 0)
    ws.set_zoom(85)
    ws.merge_range("A1:Q1", "PCS  /  COACHING WORKSPACE", book.report.title)
    ws.merge_range(
        "A2:Q2",
        "This view follows CONTROL. Open COACHING to select a key and record the action; never type in this calculated view.",
        book.report.subtitle,
    )
    for cell, target, label in (
        ("A4", "CONTROL", "CHANGE FILTERS"),
        ("D4", "COACHING", "RECORD AN ACTION"),
        ("G4", "TEAM_VIEW", "TEAM REALISATIONS"),
    ):
        ws.write_url(cell, f"internal:'{target}'!A1", book.report.editable, string=label)
    headers = [
        "LOB", "Team Leader", "Agent Selector", "Agent ID", "Priority", "Date",
        "Call Start", "Q1 Score", "Customer Comment", "Call Reference Number",
        "Coaching Key", "Action Status", "Coach", "Coaching Date", "Due Date",
        "Action Comment",
    ]
    for column, header in enumerate(headers):
        ws.write(6, column, header, book.report.header)
    formula = (
        '=LET(q,tblCoachingQueue,'
        'm,(q[Date]>=PCS_From)*(q[Date]<=PCS_To)*'
        'IF(CONTROL!$K$6="All",1,--(q[LOB]=CONTROL!$K$6))*'
        'IF(CONTROL!$N$6="All",1,--(q[Team Leader]=CONTROL!$N$6))*'
        'IF(CONTROL!$Q$6="All",1,--(q[Agent Selector]=CONTROL!$Q$6)),'
        'x,FILTER(CHOOSECOLS(q,1,2,3,5,6,7,8,9,10,11,13),m),'
        'k,CHOOSECOLS(x,11),'
        'IFERROR(HSTACK(x,'
        'XLOOKUP(k,tblCoaching[Coaching Key],tblCoaching[Coaching Status],"Not started"),'
        'XLOOKUP(k,tblCoaching[Coaching Key],tblCoaching[Coach],""),'
        'XLOOKUP(k,tblCoaching[Coaching Key],tblCoaching[Coaching Date],""),'
        'XLOOKUP(k,tblCoaching[Coaching Key],tblCoaching[Due Date],""),'
        'XLOOKUP(k,tblCoaching[Coaching Key],tblCoaching[Coaching Comment],"")),"No coaching opportunities in this selection"))'
    )
    ws.write_dynamic_array_formula("A8", formula, book.report.body, "Open in Microsoft 365 desktop Excel")
    ws.conditional_format("L8:L1048576", {
        "type": "text", "criteria": "containing", "value": "Not started",
        "format": book.report.error,
    })
    ws.set_column("A:B", 20)
    ws.set_column("C:C", 30)
    ws.set_column("D:H", 17)
    ws.set_column("I:I", 40)
    ws.set_column("J:O", 20)
    ws.set_column("P:P", 40)


def _add_pcs_filtered_data(book: DecisionWorkbook) -> None:
    """Expose a selector-driven clean export without pretending to filter raw tables."""

    wb = book.report.workbook
    ws = wb.add_worksheet("FILTERED_DATA")
    ws.hide_gridlines(2)
    ws.set_tab_color(COLORS["teal"])
    ws.freeze_panes(4, 0)
    ws.set_zoom(80)
    headers = [
        "LOB", "Team Leader", "Agent Selector", "Agent ID", "Agent", "Date",
        "Ops Manager", "Language", "Inbound Call Legs", "PCS Status 1",
        "Q1 Nonblank", "Valid Q1", "Q1 Score Sum", "PCS Average",
        "Participation Rate", "Score <= 3", "Score > 3", "Invalid Q1",
        "Sample State", "Agent Day Key",
    ]
    ws.merge_range("A1:T1", "PCS  /  FILTERED CLEAN DATA", book.report.title)
    ws.merge_range(
        "A2:T2",
        "Copy or analyze this view. It follows the period, LOB, Team Leader and Agent selected on CONTROL.",
        book.report.subtitle,
    )
    for column, header in enumerate(headers):
        ws.write(3, column, header, book.report.header)
    formula = (
        '=LET(d,tblPcsData,m,(d[Date]>=PCS_From)*(d[Date]<=PCS_To)*'
        'IF(CONTROL!$K$6="All",1,--(d[LOB]=CONTROL!$K$6))*'
        'IF(CONTROL!$N$6="All",1,--(d[Team Leader]=CONTROL!$N$6))*'
        'IF(CONTROL!$Q$6="All",1,--(d[Agent Selector]=CONTROL!$Q$6)),'
        'FILTER(CHOOSECOLS(d,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20),m,"No matching rows"))'
    )
    ws.write_dynamic_array_formula("A5", formula, book.report.body, "Open in Microsoft 365 desktop Excel")
    ws.set_column("A:B", 20)
    ws.set_column("C:C", 30)
    ws.set_column("D:H", 18)
    ws.set_column("I:T", 17)


def _add_pcs_setup(book: DecisionWorkbook, config: Config) -> None:
    folder = config.feed / "PCS"
    ws = book.table(
        "SETUP", "PCS connection setup",
        "One owner performs this once. Normal users only use Data > Refresh All and never replace this workbook.",
        ["Setting", "Value", "Why it exists"],
        [
            ("Power Query Installed", "NO", "Change to YES only after both queries refresh successfully"),
            ("Connection Mode", "LOCAL", "LOCAL: one owner refreshes on the WFM machine; SHAREPOINT: feeds must also exist in SharePoint"),
            ("Connection Owner", "Anass ASSRI", "One named owner avoids competing connection edits"),
            ("SharePoint Site URL", "https://company.sharepoint.com/sites/WFM", "Replace with your SharePoint site URL"),
            ("SharePoint Feed Folder", "/Shared Documents/WFMHub/Feed/PCS/", "Unique folder fragment containing the fixed CSV feeds"),
            ("Local Feed Folder", str(folder), "Use this for the local scripts when SharePoint is not ready"),
            ("SharePoint Data Script", str(folder / "POWER_QUERY_PCS_DATA_SHAREPOINT.txt"), "Use for PCS_DATA only when the CSV is in the SharePoint feed folder"),
            ("SharePoint Coaching Script", str(folder / "POWER_QUERY_COACHING_QUEUE_SHAREPOINT.txt"), "Use for COACHING_QUEUE only when the CSV is in SharePoint"),
            ("Local Data Script", str(folder / "POWER_QUERY_PCS_DATA_LOCAL.txt"), "Simplest option for PCS_DATA on the WFM work machine"),
            ("Local Coaching Script", str(folder / "POWER_QUERY_COACHING_QUEUE_LOCAL.txt"), "Simplest option for COACHING_QUEUE on the WFM work machine"),
            ("Workbook Last Refreshed", "Never", "Written by the Excel refresh helper after both queries finish"),
            ("Last Installer Result", "Not run", "Connection or refresh result from desktop Excel"),
            ("Template Version", PCS_TEMPLATE_VERSION, "WFMHub upgrades older designs after preserving coaching actions"),
        ],
        editable_headers={"Value"},
    )
    ws.set_column("A:A", 28)
    ws.set_column("B:B", 76)
    ws.set_column("C:C", 58)
    ws.data_validation("B5", {
        "validate": "list", "source": ["NO", "YES"],
        "input_title": "Setup state",
        "input_message": "Use YES only after both queries refresh successfully.",
    })
    ws.data_validation("B6", {
        "validate": "list", "source": ["LOCAL", "SHAREPOINT"],
        "input_title": "Connection mode",
        "input_message": "LOCAL is simplest when one owner refreshes on the WFM machine.",
    })


def _add_pcs_stable_overview(
    book: DecisionWorkbook,
    status: str,
    status_text: str,
    latest: date,
    headers: Sequence[str],
    rows: Sequence[Sequence[Any]],
) -> None:
    """Create a repair-resistant management page with a visible per-LOB table."""

    wb = book.report.workbook
    ws = wb.add_worksheet("OVERVIEW")
    ws.hide_gridlines(2)
    ws.set_tab_color(COLORS["gold"])
    ws.set_zoom(82)
    ws.set_landscape()
    ws.fit_to_pages(1, 0)
    ws.freeze_panes(34, 0)
    ws.merge_range("A1:S1", "PCS  /  OPERATIONAL SCORECARD", book.report.title)
    ws.merge_range(
        "A2:S2",
        f"Data through {latest:%Y-%m-%d}  |  latest day and current month versus the comparable previous month  |  prepared by Anass ASSRI",
        book.report.subtitle,
    )
    badge = status if status in book.badge_formats else "INCOMPLETE"
    ws.merge_range("A4:S4", f"{badge}  /  {status_text}", book.badge_formats[badge])

    all_values = list(rows[0]) if rows else [None] * len(headers)
    card_specs = (
        ("LATEST DAY PCS", 2, book.card_decimal, "Latest available business day"),
        ("CURRENT MTD PCS", 5, book.card_decimal, "Weighted valid Q1 score"),
        ("PRIOR MTD PCS", 6, book.card_decimal, "Previous month, same number of days"),
        ("MTD MOVEMENT", 7, book.card_decimal, "Current MTD minus prior MTD"),
        ("MTD PARTICIPATION", 8, book.card_percent, "Q1 nonblank / PCS Status 1"),
        ("VALID RESPONSES", 10, book.card_integer, "Current MTD sample"),
        ("SCORE <= 3", 13, book.card_integer, "Current MTD coaching opportunities"),
        ("INBOUND LEGS", 15, book.card_integer, "Current MTD inbound volume"),
    )
    for index, (label, source_column, fmt, note) in enumerate(card_specs):
        row = 5 if index < 4 else 10
        column = (index % 4) * 4
        excel_column = chr(ord("A") + source_column)
        formula = (
            f'=IFERROR(INDEX(${excel_column}$35:${excel_column}$234,'
            'MATCH("ALL",$A$35:$A$234,0)),"")'
        )
        cached = all_values[source_column] if source_column < len(all_values) else ""
        ws.merge_range(row, column, row, column + 2, label, book.report.kpi_label)
        ws.merge_range(row + 1, column, row + 2, column + 2, "", fmt)
        ws.write_formula(row + 1, column, formula, fmt, cached if cached is not None else "")
        ws.merge_range(row + 3, column, row + 3, column + 2, note, book.card_compare)

    score_chart = wb.add_chart({"type": "column"})
    for name, column, color in (
        ("Current MTD PCS", 5, COLORS["teal"]),
        ("Prior MTD PCS", 6, COLORS["muted"]),
    ):
        score_chart.add_series({
            "name": name,
            "categories": ["OVERVIEW", 35, 0, 64, 0],
            "values": ["OVERVIEW", 35, column, 64, column],
            "fill": {"color": color}, "border": {"none": True},
        })
    score_chart.set_title({"name": "PCS by LOB — current MTD vs prior MTD"})
    score_chart.set_y_axis({"min": 1, "max": 5, "major_unit": 1, "major_gridlines": {"visible": False}})
    score_chart.set_legend({"position": "bottom"})
    score_chart.set_chartarea({"border": {"none": True}, "fill": {"color": COLORS["white"]}})
    score_chart.set_plotarea({"border": {"none": True}, "fill": {"color": COLORS["white"]}})
    ws.insert_chart("A16", score_chart, {"x_scale": 1.2, "y_scale": 1.0})

    participation_chart = wb.add_chart({"type": "column"})
    for name, column, color in (
        ("Current MTD participation", 8, COLORS["gold"]),
        ("Prior MTD participation", 9, COLORS["muted"]),
    ):
        participation_chart.add_series({
            "name": name,
            "categories": ["OVERVIEW", 35, 0, 64, 0],
            "values": ["OVERVIEW", 35, column, 64, column],
            "fill": {"color": color}, "border": {"none": True},
        })
    participation_chart.set_title({"name": "Participation by LOB"})
    participation_chart.set_y_axis({"min": 0, "max": 1, "major_unit": 0.2, "num_format": "0%", "major_gridlines": {"visible": False}})
    participation_chart.set_legend({"position": "bottom"})
    participation_chart.set_chartarea({"border": {"none": True}, "fill": {"color": COLORS["white"]}})
    participation_chart.set_plotarea({"border": {"none": True}, "fill": {"color": COLORS["white"]}})
    ws.insert_chart("J16", participation_chart, {"x_scale": 1.2, "y_scale": 1.0})

    ws.merge_range("A32:S32", "PER-LOB PCS RECONCILIATION", book.report.section)
    display_headers = list(headers)
    table_rows = list(rows) or [tuple(None for _ in headers)]
    for row_index, values in enumerate(table_rows, 34):
        for column, value in enumerate(values):
            header = display_headers[column]
            fmt = book.report.body
            if isinstance(value, datetime):
                fmt = book.report.datetime
            elif isinstance(value, date):
                fmt = book.report.date
            elif "Participation" in header:
                fmt = book.report.percent
            elif "PCS" in header or header == "MTD Change":
                fmt = book.report.decimal
            elif isinstance(value, (int, float)) and not isinstance(value, bool):
                fmt = book.report.integer
            ws.write(row_index, column, value, fmt)
    ws.add_table(33, 0, 33 + len(table_rows), len(headers) - 1, {
        "name": "tblPcsLob", "style": "Table Style Light 9",
        "columns": [{"header": header, "header_format": book.report.header} for header in display_headers],
    })
    book.tables.append(ModelTable("OVERVIEW", display_headers, table_rows))
    ws.conditional_format("H35:H234", {"type": "cell", "criteria": "<", "value": 0, "format": book.report.error})
    ws.set_column("A:A", 18)
    ws.set_column("B:B", 14)
    ws.set_column("C:J", 18)
    ws.set_column("K:P", 19)
    ws.set_column("Q:Q", 16)
    ws.set_column("R:S", 20)


def _add_pcs_stable_coaching(
    book: DecisionWorkbook,
    previous: Sequence[dict[str, Any]],
) -> None:
    """Create a permanent action table using only classic INDEX/MATCH formulas."""

    wb = book.report.workbook
    ws = wb.add_worksheet("COACHING")
    ws.hide_gridlines(2)
    ws.set_tab_color(COLORS["gold"])
    ws.freeze_panes(4, 0)
    ws.set_zoom(85)
    headers = [
        "Coaching Key", "LOB", "Team Leader", "Agent Selector", "Agent ID",
        "Agent", "Date", "Call Start", "Q1 Score", "Customer Comment",
        "Call Reference Number", "Language", "Coaching Status", "Coach",
        "Coaching Date", "Due Date", "Coaching Comment",
    ]
    ws.merge_range("A1:Q1", "PCS  /  COACHING ACTIONS", book.report.title)
    ws.merge_range(
        "A2:Q2",
        "Filter COACHING_QUEUE, copy its Coaching Key, then paste it in the first blank blue cell here. Identity fields fill automatically.",
        book.report.subtitle,
    )
    rows = list(previous) or [{}]
    editable = {
        "Coaching Key", "Coaching Status", "Coach", "Coaching Date",
        "Due Date", "Coaching Comment",
    }
    for row_index, item in enumerate(rows, 4):
        for column, header in enumerate(headers):
            key = header.casefold().replace(" ", "_")
            value = item.get(key)
            fmt = book.report.editable if header in editable else book.report.body
            if header in {"Coaching Date", "Due Date"}:
                fmt = book.report.editable_date
            ws.write(row_index, column, value, fmt)
    queue_columns = {
        "LOB": "A", "Team Leader": "B", "Agent Selector": "C",
        "Agent": "D", "Agent ID": "E", "Date": "G", "Call Start": "H",
        "Q1 Score": "I", "Customer Comment": "J",
        "Call Reference Number": "K", "Language": "L",
    }
    columns = []
    for header in headers:
        column: dict[str, Any] = {"header": header, "header_format": book.report.header}
        if header in queue_columns:
            source = queue_columns[header]
            column["formula"] = (
                '=IF([@[Coaching Key]]="","",IFERROR('
                f'INDEX(COACHING_QUEUE!${source}$5:${source}$100004,'
                'MATCH([@[Coaching Key]],COACHING_QUEUE!$M$5:$M$100004,0)),'
                '"KEY NOT IN CURRENT QUEUE"))'
            )
            if header == "Date":
                column["format"] = book.report.date
            elif header == "Call Start":
                column["format"] = book.report.datetime
        columns.append(column)
    ws.add_table(3, 0, 3 + len(rows), len(headers) - 1, {
        "name": "tblCoaching", "style": "Table Style Light 9", "columns": columns,
    })
    book.tables.append(ModelTable(
        "COACHING", headers,
        [tuple(item.get(header.casefold().replace(" ", "_")) for header in headers) for item in rows],
    ))
    wb.define_name(
        "PCS_COACHING_KEY_LIST",
        "=COACHING_QUEUE!$M$5:INDEX(COACHING_QUEUE!$M:$M,MAX(5,COUNTA(COACHING_QUEUE!$M:$M)+3))",
    )
    ws.data_validation("A5:A1004", {
        "validate": "list", "source": "=PCS_COACHING_KEY_LIST",
        "input_title": "Exact PCS case",
        "input_message": "Paste or select a key from COACHING_QUEUE.",
    })
    ws.data_validation("M5:M1004", {
        "validate": "list", "source": ["Pending", "Planned", "Completed", "Not required"],
    })
    ws.conditional_format("A5:A1004", {"type": "duplicate", "format": book.report.error})
    ws.write_url("A3", "internal:'COACHING_QUEUE'!A1", book.report.editable, string="OPEN FILTERABLE OPPORTUNITY QUEUE")
    ws.set_column("A:A", 34)
    ws.set_column("B:C", 20)
    ws.set_column("D:D", 30)
    ws.set_column("E:I", 18)
    ws.set_column("J:J", 38)
    ws.set_column("K:P", 20)
    ws.set_column("Q:Q", 42)


def _add_pcs_stable_setup(book: DecisionWorkbook, config: Config) -> None:
    folder = config.feed / "PCS"
    ws = book.table(
        "SETUP", "PCS connection status",
        "The WFMHub PCS menu installs or refreshes four ordinary Power Query tables. No Data Model is used.",
        ["Setting", "Value", "Why it exists"],
        [
            ("Power Query Installed", "NO", "Set automatically after all four query tables refresh"),
            ("Connection Mode", "LOCAL", "One WFM owner refreshes the locally synced workbook"),
            ("Connection Owner", "Anass ASSRI", "Prevents competing setup changes"),
            ("Local Feed Folder", str(folder), "Fixed clean CSV folder"),
            ("SharePoint Site URL", "https://company.sharepoint.com/sites/WFM", "Only needed for advanced SharePoint-feed mode"),
            ("SharePoint Feed Folder", "/Shared Documents/WFMHub/Feed/PCS/", "Unique folder fragment containing the fixed feeds"),
            ("LOB Script", str(folder / "POWER_QUERY_PCS_LOB_LOCAL.txt"), "Per-LOB scorecard"),
            ("Results Script", str(folder / "POWER_QUERY_PCS_RESULTS_LOCAL.txt"), "Filterable period/team/agent results"),
            ("Coaching Script", str(folder / "POWER_QUERY_COACHING_QUEUE_LOCAL.txt"), "Filterable opportunity queue"),
            ("Data Script", str(folder / "POWER_QUERY_PCS_DATA_LOCAL.txt"), "Clean agent/day table for pivots"),
            ("Workbook Last Refreshed", "Never", "Written only after all queries complete"),
            ("Last Installer Result", "Not run", "Latest desktop Excel result"),
            ("Template Version", PCS_TEMPLATE_VERSION, "Controls safe workbook upgrades"),
        ],
        editable_headers={"Value"},
    )
    ws.set_column("A:A", 28)
    ws.set_column("B:B", 76)
    ws.set_column("C:C", 58)


def _pcs_workbook_template_version(path: Path) -> str | None:
    """Read the permanent tracker's declared template version without changing it."""

    if not path.is_file():
        return None
    try:
        workbook = load_workbook(path, read_only=True, data_only=False, keep_links=False)
    except Exception:
        return None
    try:
        if "SETUP" not in workbook.sheetnames:
            return None
        for setting, value, *_rest in workbook["SETUP"].iter_rows(min_row=5, values_only=True):
            if str(setting or "").strip() == "Template Version":
                return str(value or "").strip() or None
        return None
    finally:
        workbook.close()


def _previous_coaching_actions(path: Path) -> list[dict[str, Any]]:
    """Preserve old and new coaching ledgers during a versioned tracker upgrade."""

    if not path.is_file():
        return []
    try:
        workbook = load_workbook(path, read_only=True, data_only=False, keep_links=False)
    except Exception:
        return []
    try:
        if "COACHING" not in workbook.sheetnames:
            return []
        sheet = workbook["COACHING"]
        headers = {
            str(cell.value or "").strip(): cell.column
            for cell in sheet[4]
            if cell.value not in (None, "")
        }
        key_column = headers.get("Coaching Key")
        if key_column is None:
            return []
        aliases = {
            "coaching_key": "Coaching Key",
            "coaching_status": "Coaching Status",
            "coach": "Coach",
            "coaching_date": "Coaching Date",
            "due_date": "Due Date",
            "coaching_comment": "Coaching Comment",
        }
        output: list[dict[str, Any]] = []
        seen: set[str] = set()
        for row in sheet.iter_rows(min_row=5, values_only=True):
            key = row[key_column - 1] if key_column <= len(row) else None
            key_text = str(key or "").strip()
            if not key_text or key_text in seen:
                continue
            values: dict[str, Any] = {"coaching_key": key_text}
            for field, header in aliases.items():
                column = headers.get(header)
                if column is not None and column <= len(row):
                    value = row[column - 1]
                    if value not in (None, ""):
                        values[field] = value
            if values.get("coaching_status") in (None, ""):
                values["coaching_status"] = "Pending"
            output.append(values)
            seen.add(key_text)
        return output
    finally:
        workbook.close()


def _previous_table_values(
    path: Path,
    sheet_name: str,
    key_header: str,
    editable_headers: Sequence[str],
) -> dict[str, dict[str, Any]]:
    """Read manual values from a prior report without importing them to the Hub."""

    if not path.exists():
        return {}
    try:
        workbook = load_workbook(path, read_only=True, data_only=False, keep_links=False)
    except Exception:
        return {}
    try:
        if sheet_name not in workbook.sheetnames:
            return {}
        sheet = workbook[sheet_name]
        headers = {
            str(cell.value).strip(): index
            for index, cell in enumerate(next(sheet.iter_rows(min_row=4, max_row=4)), 1)
            if cell.value is not None
        }
        key_column = headers.get(key_header)
        if key_column is None:
            return {}
        output: dict[str, dict[str, Any]] = {}
        for values in sheet.iter_rows(min_row=5, values_only=True):
            key = values[key_column - 1] if key_column <= len(values) else None
            if not key:
                continue
            output[str(key)] = {
                field: values[column - 1] if column <= len(values) else None
                for field in editable_headers
                if (column := headers.get(field)) is not None
            }
        return output
    finally:
        workbook.close()


def _carry_table_values(
    headers: Sequence[str],
    rows: Sequence[Sequence[Any]],
    key_header: str,
    editable_headers: Sequence[str],
    previous: dict[str, dict[str, Any]],
) -> list[tuple[Any, ...]]:
    display = [header.replace("_", " ").title().replace("Id", "ID") for header in headers]
    indexes = {header: index for index, header in enumerate(display)}
    key_index = indexes.get(key_header)
    if key_index is None:
        return [tuple(row) for row in rows]
    output: list[tuple[Any, ...]] = []
    for raw in rows:
        values = list(raw)
        saved = previous.get(str(values[key_index]), {})
        for field in editable_headers:
            if field in indexes and field in saved:
                values[indexes[field]] = saved[field]
        output.append(tuple(values))
    return output


def _add_pcs_lookups(
    book: DecisionWorkbook,
    trend_days: int,
) -> None:
    """Hidden rolling chart calculations and cascading selector spills."""
    ws = book.report.workbook.add_worksheet("_LOOKUPS")
    headers = [
        "Date", "Q1 Score Sum", "Valid Q1", "Q1 Nonblank", "PCS Status 1",
        "PCS Average", "Participation",
    ]
    for column, header in enumerate(headers):
        ws.write(0, column, header)
    for row_index in range(1, trend_days + 1):
        excel_row = row_index + 1
        ws.write_formula(row_index, 0, f"=PCS_Latest-{trend_days - row_index}")
        criteria = (
            f'(tblPcsData[Date]=$A${excel_row})*'
            'IF(CONTROL!$K$6="All",1,--(tblPcsData[LOB]=CONTROL!$K$6))*'
            'IF(CONTROL!$N$6="All",1,--(tblPcsData[Team Leader]=CONTROL!$N$6))*'
            'IF(CONTROL!$Q$6="All",1,--(tblPcsData[Agent Selector]=CONTROL!$Q$6))'
        )
        for column, source in enumerate(
            ("Q1 Score Sum", "Valid Q1", "Q1 Nonblank", "PCS Status 1"), 1,
        ):
            ws.write_formula(
                row_index, column,
                f'=IF(OR($A${excel_row}<PCS_From,$A${excel_row}>PCS_To),NA(),SUMPRODUCT({criteria}*N(tblPcsData[{source}])))',
            )
        ws.write_formula(
            row_index, 5,
            f'=IFERROR($B${excel_row}/$C${excel_row},NA())',
        )
        ws.write_formula(
            row_index, 6,
            f'=IFERROR($D${excel_row}/$E${excel_row},NA())',
        )
    ws.write_dynamic_array_formula(
        "J2", '=VSTACK("All",SORT(UNIQUE(FILTER(tblPcsData[LOB],tblPcsData[LOB]<>""))))',
    )
    ws.write_dynamic_array_formula(
        "K2", '=LET(d,tblPcsData,VSTACK("All",SORT(UNIQUE(FILTER(d[Team Leader],(d[Team Leader]<>"")*IF(CONTROL!$K$6="All",1,--(d[LOB]=CONTROL!$K$6)))))))',
    )
    ws.write_dynamic_array_formula(
        "L2", '=LET(d,tblPcsData,VSTACK("All",SORT(UNIQUE(FILTER(d[Agent Selector],(d[Agent Selector]<>"")*IF(CONTROL!$K$6="All",1,--(d[LOB]=CONTROL!$K$6))*IF(CONTROL!$N$6="All",1,--(d[Team Leader]=CONTROL!$N$6)))))))',
    )
    ws.write_dynamic_array_formula(
        "M2", '=SORT(UNIQUE(FILTER(tblCoachingQueue[Coaching Key],tblCoachingQueue[Coaching Key]<>"","")))',
    )
    wb = book.report.workbook
    for name, column in (
        ("PCS_LOB_LIST", "J"),
        ("PCS_TL_LIST", "K"),
        ("PCS_AGENT_LIST", "L"),
        ("PCS_COACHING_KEY_LIST", "M"),
    ):
        wb.define_name(name, f"=_LOOKUPS!${column}$2#")
    ws.hide()


def build_pcs_performance_workbook(
    conn: DatabaseConnection,
    config: Config,
    start: date,
    end: date,
    output: Path | None = None,
) -> Path:
    """Create or safely upgrade the repair-resistant permanent PCS tracker."""

    from .shared_feeds import (
        PCS_AGENT_DAY_HEADERS,
        PCS_COACHING_HEADERS,
        PCS_LOB_SCORECARD_HEADERS,
        PCS_RESULTS_HEADERS,
        pcs_lob_scorecard_rows,
        pcs_result_rows,
        publish_pcs_feeds,
    )

    publish_pcs_feeds(conn, config, start, end)
    target = _output_path(config, "pcs", start, end, datetime.now(), output)
    current_version = _pcs_workbook_template_version(target)
    if output is None and target.exists() and current_version == PCS_TEMPLATE_VERSION:
        return target
    previous_target = target if target.exists() else (config.reports / "PCS Performance.xlsx").resolve()
    previous_actions = _previous_coaching_actions(previous_target)
    book, partial, target = _atomic_book(
        config, "pcs", "PCS OPERATIONAL TRACKER", start, end, output,
    )
    latest_value = conn.execute("SELECT max(business_date) FROM mart.agent_pcs_day").fetchone()[0]
    latest = latest_value or end
    if isinstance(latest, str):
        latest = date.fromisoformat(latest[:10])
    metric_catalog = load_metric_catalog(config.home, config.metric_catalog)
    pcs_method = metric_catalog.method_for("pcs_average", latest, {})
    minimum_sample = int(pcs_method.minimum_sample) if pcs_method is not None else 1
    status, status_text = _source_state(conn, ("fte", "calls"), latest)
    month_index = latest.year * 12 + latest.month - 1
    first_index = month_index - (config.pcs_tracker.history_months - 1)
    history_start = date(first_index // 12, first_index % 12 + 1, 1)
    data_start = min(start, history_start)

    lob_rows = pcs_lob_scorecard_rows(conn, latest, minimum_sample, book.generated)
    result_rows = pcs_result_rows(conn, latest, minimum_sample, book.generated)
    _add_pcs_stable_overview(
        book, status, status_text, latest,
        list(PCS_LOB_SCORECARD_HEADERS), lob_rows,
    )
    results_sheet = book.table(
        "RESULTS", "PCS team and agent results",
        "Use the table filters: first Period View, then Scope Level, LOB, Team Leader and Agent. Add native slicers here if desired.",
        list(PCS_RESULTS_HEADERS), result_rows or [tuple(None for _ in PCS_RESULTS_HEADERS)],
    )
    if pcs_method is not None and pcs_method.target is not None:
        results_sheet.conditional_format("K5:K100004", {
            "type": "cell", "criteria": "<", "value": pcs_method.target,
            "format": book.report.error,
        })
    results_sheet.conditional_format("S5:S100004", {
        "type": "text", "criteria": "containing", "value": "LOW SAMPLE", "format": book.report.error,
    })

    action_headers, actions = _pcs_coaching_rows(conn, config, data_start, latest)
    indexes = {header: index for index, header in enumerate(action_headers)}
    sources = {
        "LOB": "lob", "Team Leader": "team_leader", "Agent Selector": "agent_selector",
        "Agent": "agent_name", "Agent ID": "agent_id", "Priority": "priority",
        "Date": "business_date", "Call Start": "call_start", "Q1 Score": "q1_score",
        "Customer Comment": "customer_comment", "Call Reference Number": "call_reference_number",
        "Language": "language", "Coaching Key": "coaching_key",
    }
    queue_rows = [
        tuple(values[indexes[sources[header]]] for header in PCS_COACHING_HEADERS)
        for values in actions
    ]
    queue_sheet = book.table(
        "COACHING_QUEUE", "PCS coaching opportunity queue",
        "Filter by date, LOB, Team Leader or Agent. Copy only the exact Coaching Key into the permanent COACHING action table.",
        list(PCS_COACHING_HEADERS), queue_rows or [tuple(None for _ in PCS_COACHING_HEADERS)],
    )
    queue_sheet.conditional_format("F5:F100004", {
        "type": "text", "criteria": "containing", "value": "HIGH", "format": book.report.error,
    })
    _add_pcs_stable_coaching(book, previous_actions)

    rulebook = load_rulebook(config.home, config.business_rules)
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
            minimum_sample, latest, book.generated, rulebook.version, rulebook.sha256,
            metric_catalog.version, metric_catalog.sha256, data_start, latest,
        ],
    )
    book.table(
        "PCS_DATA", "PCS clean agent-day data",
        "Use this ordinary refreshable table for pivots, slicers or a custom date range. Calculated reports use summed counters, never averaged averages.",
        list(PCS_AGENT_DAY_HEADERS), data_rows or [tuple(None for _ in PCS_AGENT_DAY_HEADERS)],
    )
    _add_pcs_stable_setup(book, config)
    book.table(
        "HELP", "PCS operating guide",
        "This design uses ordinary tables and classic formulas to avoid Excel repair prompts.",
        ["Step", "What to do", "Result", "Important"],
        [
            (1, "Close the workbook, then choose PCS > Update PCS now in WFMHub", "Four clean feeds and four Excel query tables refresh", "Keep one permanent workbook"),
            (2, "Open OVERVIEW", "See total and per-LOB latest day, current MTD and prior-MTD PCS and participation", "The ALL row reconciles the headline cards"),
            (3, "Open RESULTS and use its filter arrows", "Choose a period, LOB, team or agent without formulas or technical selectors", "Add slicers to this table if desired"),
            (4, "Quality filters COACHING_QUEUE and copies one Coaching Key", "Paste the key in the first blank blue COACHING row", "Fill only status, owner, dates and comment"),
            (5, "Save the same shared workbook", "Coaching remains under SharePoint version history", "A template upgrade archives and migrates keyed actions"),
        ],
    )
    book.definitions([
        ("PCS Average", "Sum of valid inbound Q1 scores / valid inbound Q1 responses", "Weighted service, LOB, team and agent result", "Never sum Q1 score or average agent averages"),
        ("PCS Participation", "Inbound raw Q1 nonblank / inbound PCSStatus=1", "Survey participation opportunity", "Invalid nonblank Q1 remains in the numerator"),
        ("Score <= 3", "Count of valid inbound Q1 responses at or below 3", "Coaching opportunity volume", "One exact call equals one Coaching Key"),
        ("Current MTD", "First day of the latest data month through the latest data date", "Current-month operating result", "Driven by data, not today's computer date"),
        ("Previous MTD same days", "Previous month from day one through the comparable day", "Fair MTD comparison", "Month length is capped safely"),
        ("Low sample", f"Fewer than {minimum_sample} valid responses in the row period", "Interpretation warning", "Use a broader period before concluding"),
    ])
    book.audit(_audit_rows(conn, config, "pcs", start, end, (("Workbook engine", "Compatibility", "No dynamic-array report formulas"),)))
    result = _finish(book, partial, target)
    if output is None:
        archive_superseded_reports(config, ("PCS Performance.xlsx",), book.generated)
    return result


def _service_rows(
    conn: DatabaseConnection,
    profile: ServiceProfile,
    start: date,
    end: date,
) -> list[dict[str, Any]]:
    filter_values = (
        [queue.upper() for queue in profile.flash_queues]
        if profile.flash_queues else list(profile.service_scopes)
    )
    source_filter = (
        f"upper(queue) IN ({','.join('?' for _ in filter_values)})"
        if profile.flash_queues
        else f"service_scope IN ({','.join('?' for _ in filter_values)})"
    )
    system_marks = ",".join("?" for _ in profile.source_systems)
    cursor = conn.execute(
        f"""SELECT business_date, interval_start, hour_start, source_system, queue,
                   service_scope, comparison_scope, designation, mapping_status,
                   offered, answered, abandoned, short_abandoned,
                   abandoned_within_target,
                   answered_within_target, handled_seconds, source_file
            FROM mart.service_interval
            WHERE business_date BETWEEN ? AND ?
              AND {source_filter}
              AND source_system IN ({system_marks})
            ORDER BY business_date, interval_start, source_system, queue""",
        [start, end, *filter_values, *profile.source_systems],
    )
    headers = [item[0] for item in cursor.description]
    rows = [dict(zip(headers, row)) for row in cursor.fetchall()]
    if profile.flash_total_groups and not profile.flash_queues:
        allowed = set(profile.flash_total_groups)
        rows = [
            row for row in rows
            if profile.group_for(row.get("queue")) in allowed
        ]
    return rows


def _profile_metric(catalog: MetricCatalog, profile: ServiceProfile, metric_id: str, on_date: date):
    methods = {}
    for service_scope in profile.service_scopes:
        for source_system in profile.source_systems:
            method = catalog.method_for(metric_id, on_date, {
                "lob": service_scope,
                "source_system": source_system,
            })
            if method is None:
                raise ValueError(
                    f"Service profile {profile.profile_id!r} selects metric {metric_id!r}, "
                    f"but no method applies to {service_scope}/{source_system} on {on_date}"
                )
            methods[(method.method_id, method.effective_from, method.priority)] = method
    if len(methods) != 1:
        names = ", ".join(key[0] for key in methods)
        raise ValueError(
            f"Service profile {profile.profile_id!r} crosses incompatible {metric_id} methods: {names}. "
            "Split it into separate service profiles."
        )
    return next(iter(methods.values()))


def _service_aggregate(
    rows: Iterable[dict[str, Any]],
    profile: ServiceProfile,
    catalog: MetricCatalog,
    on_date: date,
) -> dict[str, float | str | None]:
    values = list(rows)
    offered = sum(float(row.get("offered") or 0) for row in values)
    answered = sum(float(row.get("answered") or 0) for row in values)
    abandoned = sum(float(row.get("abandoned") or 0) for row in values)
    short = sum(float(row.get("short_abandoned") or 0) for row in values)
    abandoned_in_target = sum(
        float(row.get("abandoned_within_target") or 0) for row in values
    )
    within = sum(float(row.get("answered_within_target") or 0) for row in values)
    handled = sum(float(row.get("handled_seconds") or 0) for row in values)
    components = {
        "offered": offered,
        "answered": answered,
        "abandoned": abandoned,
        "short_abandoned": short,
        "abandoned_within_target": abandoned_in_target,
        "answered_within_target": within,
        "handled_seconds": handled,
    }
    service_level = evaluate_metric(
        _profile_metric(catalog, profile, profile.service_level_metric, on_date), components,
    )
    availability = evaluate_metric(
        _profile_metric(catalog, profile, profile.availability_metric, on_date), components,
    )
    aht = evaluate_metric(
        _profile_metric(catalog, profile, profile.aht_metric, on_date), components,
    )
    return {
        "raw_offered": offered,
        "offered": offered,
        "business_offered": max(0.0, offered - abandoned_in_target),
        "answered": answered,
        "short_abandoned": short,
        "abandoned_within_target": abandoned_in_target,
        "within_target": within,
        "service_level": service_level.value,
        "service_target": service_level.method.target,
        "service_state": service_level.state,
        "service_method": service_level.method.method_id,
        "availability": availability.value,
        "aht_seconds": aht.value,
    }


def build_service_performance_workbook(
    conn: DatabaseConnection,
    config: Config,
    start: date,
    end: date,
    output: Path | None = None,
    profile_id: str | None = None,
) -> Path:
    """Compatibility wrapper for the all-profile Book1 Flash workbook."""

    from .service_flash import build_service_flashes_workbook

    return build_service_flashes_workbook(
        conn, config, start, end, output, profile_id,
    )


def build_realisations_workbook(
    conn: DatabaseConnection,
    config: Config,
    start: date,
    end: date,
    output: Path | None = None,
    profile_id: str | None = None,
) -> Path:
    """Build one management view across every configured service LOB.

    Supplying ``profile_id`` keeps the focused single-LOB route available for
    command-line use.  The normal menu intentionally includes all active
    profiles so Operations does not need to build four separate workbooks.
    """

    profiles = load_service_profiles(config.home, config.service_profiles)
    selected_profiles = (
        (profiles.select(profile_id, end),)
        if profile_id
        else tuple(profile for profile in profiles.profiles if profile.active_on(end))
    )
    metrics_catalog = load_metric_catalog(config.home, config.metric_catalog)
    queue_mapping = load_queue_mapping(config.queue_mapping)
    title = (
        f"REALISATIONS  /  {selected_profiles[0].label.upper()}"
        if len(selected_profiles) == 1
        else "REALISATIONS  /  ALL MAPPED LOBS"
    )
    book, partial, target = _atomic_book(
        config, "realisations", title, start, end, output,
    )
    daily_rows: list[tuple[Any, ...]] = []
    all_source_rows: list[dict[str, Any]] = []
    profile_by_label = {profile.label: profile for profile in selected_profiles}

    for profile in selected_profiles:
        source_rows = _service_rows(conn, profile, start, end)
        for row in source_rows:
            all_source_rows.append({**row, "reporting_lob": profile.label})
        forecast_scopes = queue_mapping.comparison_scopes_for(
            profile.service_scopes,
        )
        forecast_marks = ",".join("?" for _ in forecast_scopes)
        forecasts = conn.execute(
            f"""SELECT business_date, sum(volume_forecast), avg(fte_forecast),
                       avg(fte_required),
                       CASE WHEN sum(CASE WHEN sl_forecast IS NOT NULL
                                          THEN volume_forecast ELSE 0 END)>0
                            THEN sum(CASE WHEN sl_forecast IS NOT NULL
                                          THEN sl_forecast*volume_forecast ELSE 0 END)
                                 /sum(CASE WHEN sl_forecast IS NOT NULL
                                           THEN volume_forecast ELSE 0 END)
                            ELSE avg(sl_forecast) END,
                       CASE WHEN sum(CASE WHEN sl_required IS NOT NULL
                                          THEN volume_forecast ELSE 0 END)>0
                            THEN sum(CASE WHEN sl_required IS NOT NULL
                                          THEN sl_required*volume_forecast ELSE 0 END)
                                 /sum(CASE WHEN sl_required IS NOT NULL
                                           THEN volume_forecast ELSE 0 END)
                            ELSE avg(sl_required) END,
                       CASE WHEN sum(CASE WHEN aht_forecast_seconds IS NOT NULL
                                          THEN volume_forecast ELSE 0 END)>0
                            THEN sum(CASE WHEN aht_forecast_seconds IS NOT NULL
                                          THEN aht_forecast_seconds*volume_forecast ELSE 0 END)
                                 /sum(CASE WHEN aht_forecast_seconds IS NOT NULL
                                           THEN volume_forecast ELSE 0 END)
                            ELSE avg(aht_forecast_seconds) END
                FROM mart.forecast_hour
                WHERE business_date BETWEEN ? AND ?
                  AND comparison_scope IN ({forecast_marks})
                GROUP BY business_date ORDER BY business_date""",
            [start, end, *forecast_scopes],
        ).fetchall()
        forecast_by_day = {str(row[0])[:10]: row[1:] for row in forecasts}
        staffing_marks = ",".join("?" for _ in profile.staffing_lobs)
        staffing_rows = conn.execute(
            f"""SELECT business_date, avg(scheduled_fte), avg(observed_fte),
                       avg(productive_fte), max(staffing_gap_fte)
                FROM (
                    SELECT business_date, interval_start,
                           sum(scheduled_fte) AS scheduled_fte,
                           sum(observed_fte) AS observed_fte,
                           sum(productive_fte) AS productive_fte,
                           sum(staffing_gap_fte) AS staffing_gap_fte
                    FROM mart.staffing_interval
                    WHERE business_date BETWEEN ? AND ? AND lob IN ({staffing_marks})
                    GROUP BY business_date, interval_start
                ) x GROUP BY business_date""",
            [start, end, *profile.staffing_lobs],
        ).fetchall()
        staffing_by_day = {str(row[0])[:10]: row[1:] for row in staffing_rows}
        absence_rows = conn.execute(
            f"""SELECT d.business_date, sum(d.planned_net_minutes)/60.0,
                       sum(d.absence_minutes)/60.0,
                       sum(d.vacation_minutes)/60.0,
                       sum(d.shrinkage_minutes)/60.0,
                       sum(CASE WHEN d.unverified_minutes>0
                                     OR coalesce(a.is_provisional,false)
                                THEN 1 ELSE 0 END)
                FROM mart.absence_agent_day d
                LEFT JOIN mart.attendance_agent_day a
                  ON a.agent_day_key=d.agent_day_key
                WHERE d.business_date BETWEEN ? AND ? AND d.lob IN ({staffing_marks})
                GROUP BY d.business_date""",
            [start, end, *profile.staffing_lobs],
        ).fetchall()
        absence_by_day = {str(row[0])[:10]: row[1:] for row in absence_rows}
        rows_by_day: dict[str, list[dict[str, Any]]] = defaultdict(list)
        for row in source_rows:
            rows_by_day[str(row["business_date"])[:10]].append(row)

        cursor = start
        while cursor <= end:
            key = cursor.isoformat()
            actual = _service_aggregate(
                rows_by_day.get(key, []), profile, metrics_catalog, cursor,
            )
            forecast = forecast_by_day.get(key, (None, None, None, None, None, None))
            staffing = staffing_by_day.get(key, (None, None, None, None))
            absence_values = absence_by_day.get(key, (None, None, None, None, 0))
            forecast_volume = forecast[0]
            planned_hours, absence_hours, vacation_hours, shrinkage_hours, review_cases = absence_values
            has_actual = bool(rows_by_day.get(key))
            has_forecast = forecast_volume is not None
            data_state = (
                "Provisional" if cursor >= date.today()
                else "Review" if review_cases
                else "No actual" if not has_actual
                else "No forecast" if not has_forecast
                else "Final"
            )
            daily_rows.append((
                cursor, cursor.strftime("%Y-%m"), cursor.isocalendar().week,
                f"Q{((cursor.month - 1) // 3) + 1}", profile.label,
                ", ".join(profile.service_scopes), actual["offered"],
                forecast_volume,
                float(actual["offered"] or 0) - float(forecast_volume)
                if forecast_volume is not None else None,
                _ratio(actual["offered"], forecast_volume), actual["answered"],
                actual["within_target"], actual["short_abandoned"],
                actual["abandoned_within_target"],
                actual["service_level"], actual["availability"],
                actual["aht_seconds"],
                _ratio(
                    float(actual["aht_seconds"] or 0) * float(actual["answered"] or 0),
                    3600,
                ),
                staffing[0], staffing[1], staffing[2], staffing[3],
                planned_hours, absence_hours, _ratio(absence_hours, planned_hours),
                vacation_hours, shrinkage_hours, _ratio(shrinkage_hours, planned_hours),
                review_cases, data_state, profile.profile_id,
                ", ".join(profile.staffing_lobs),
            ))
            cursor += timedelta(days=1)
    daily_headers = [
        "business_date", "month", "iso_week", "quarter", "lob",
        "service_scopes", "actual_volume", "forecast_volume",
        "variance_calls", "forecast_attainment", "answered",
        "answered_within_target", "short_abandoned",
        "abandoned_within_target", "service_level",
        "service_availability", "aht_seconds", "processing_hours",
        "scheduled_fte", "observed_fte", "productive_fte", "peak_gap_fte",
        "planned_hours", "absence_hours", "absence_rate", "vacation_hours",
        "shrinkage_hours", "shrinkage_rate", "review_cases", "data_state",
        "profile_id", "staffing_lobs",
    ]

    total_actual = sum(float(row[6] or 0) for row in daily_rows)
    total_forecast = sum(float(row[7] or 0) for row in daily_rows if row[7] is not None)
    forecast_present = any(row[7] is not None for row in daily_rows)
    total_answered = sum(float(row[10] or 0) for row in daily_rows)
    total_within = sum(float(row[11] or 0) for row in daily_rows)
    total_short = sum(float(row[12] or 0) for row in daily_rows)
    total_handled_seconds = sum(
        float(row[16] or 0) * float(row[10] or 0) for row in daily_rows
    )
    total_planned = sum(float(row[22] or 0) for row in daily_rows)
    total_absence = sum(float(row[23] or 0) for row in daily_rows)
    total_shrinkage = sum(float(row[26] or 0) for row in daily_rows)
    availability_value = _ratio(total_answered, total_actual)
    aht_value = _ratio(total_handled_seconds, total_answered)
    profile_summary: list[tuple[Any, ...]] = []
    for profile in selected_profiles:
        rows = [row for row in daily_rows if row[4] == profile.label]
        offered = sum(float(row[6] or 0) for row in rows)
        forecast = sum(float(row[7] or 0) for row in rows if row[7] is not None)
        answered = sum(float(row[10] or 0) for row in rows)
        within = sum(float(row[11] or 0) for row in rows)
        short = sum(float(row[12] or 0) for row in rows)
        abandoned_in_target = sum(float(row[13] or 0) for row in rows)
        handled = sum(float(row[16] or 0) * float(row[10] or 0) for row in rows)
        planned_hours = sum(float(row[22] or 0) for row in rows)
        absence_hours = sum(float(row[23] or 0) for row in rows)
        shrinkage_hours = sum(float(row[26] or 0) for row in rows)
        profile_components = {
            "offered": offered, "answered": answered,
            "short_abandoned": short,
            "abandoned_within_target": abandoned_in_target,
            "answered_within_target": within,
            "handled_seconds": handled,
        }
        sl_result = evaluate_metric(
            _profile_metric(metrics_catalog, profile, profile.service_level_metric, end),
            profile_components,
        )
        availability_result = evaluate_metric(
            _profile_metric(metrics_catalog, profile, profile.availability_metric, end),
            profile_components,
        )
        aht_result = evaluate_metric(
            _profile_metric(metrics_catalog, profile, profile.aht_metric, end),
            profile_components,
        )
        state = (
            "NO ACTUAL" if not any(float(row[6] or 0) for row in rows)
            else "NO FORECAST" if not any(row[7] is not None for row in rows)
            else "REVIEW" if any(row[-4] for row in rows)
            else "READY"
        )
        profile_summary.append((
            profile.label, offered,
            forecast if any(row[7] is not None for row in rows) else None,
            _ratio(offered, forecast), sl_result.value, sl_result.method.target,
            availability_result.value, aht_result.value,
            _ratio(absence_hours, planned_hours),
            _ratio(shrinkage_hours, planned_hours), state,
        ))
    status, status_text = _source_state(conn, ("calls", "forecast"), end)
    review_days = sum(1 for row in daily_rows if row[-4])
    if review_days:
        status, status_text = "INCOMPLETE", f"{review_days:,} day(s) include absence review cases"
    book.dashboard(
        [
            KpiCard("Actual volume", total_actual, "integer"),
            KpiCard("Forecast volume", total_forecast if forecast_present else None, "integer"),
            KpiCard("Forecast attainment", _ratio(total_actual, total_forecast) if forecast_present else None, "percent"),
            KpiCard("Mapped LOBs", len(selected_profiles), "integer"),
            KpiCard("Routed rate", availability_value, "percent"),
            KpiCard("Weighted AHT", aht_value, "decimal"),
            KpiCard("Absence rate", _ratio(total_absence, total_planned), "percent"),
            KpiCard("Shrinkage rate", _ratio(total_shrinkage, total_planned), "percent"),
        ],
        status,
        status_text,
        [
            "LOB", "Actual", "Forecast", "Attainment %", "Service Level %",
            "SL Target %", "Routed Rate %", "AHT Seconds", "Absence Rate %",
            "Shrinkage Rate %", "State",
        ],
        profile_summary,
        [
            "Queue membership is maintained in Queue Mapping; service and roster LOB links are maintained in Service Profiles.",
            "Forecast comes from Verint. Actual volume, service level and AHT come from mapped inbound Call-by-Call queue entries.",
            "Routed Rate means answered / total entered, matching the Storm dashboard; it is not agent availability.",
            "Absence uses reviewed Attendance decisions; open gaps remain visible and cannot silently dilute results.",
            "Adherence is intentionally excluded.",
        ],
        (("Service Level", 4), ("Routed Rate", 6)),
    )
    book.table(
        "LOB_RESULTS", "Daily LOB results",
        "One normalized row per mapped management LOB and day. Use this sheet for pivots, charts and management checks.",
        daily_headers, daily_rows,
    )

    trend_rows: list[tuple[Any, ...]] = []
    for grain, index in (("Month", 1), ("ISO Week", 2), ("Quarter", 3)):
        grouped: dict[tuple[str, Any], list[tuple[Any, ...]]] = defaultdict(list)
        for row in daily_rows:
            grouped[(str(row[4]), row[index])].append(row)
        for (lob_label, label), group in sorted(grouped.items(), key=lambda item: (item[0][0], str(item[0][1]))):
            offered = sum(float(row[6] or 0) for row in group)
            forecast = sum(float(row[7] or 0) for row in group if row[7] is not None)
            answered = sum(float(row[10] or 0) for row in group)
            within = sum(float(row[11] or 0) for row in group)
            short = sum(float(row[12] or 0) for row in group)
            abandoned_in_target = sum(float(row[13] or 0) for row in group)
            handled = sum(float(row[16] or 0) * float(row[10] or 0) for row in group)
            planned_hours = sum(float(row[22] or 0) for row in group)
            absence_hours = sum(float(row[23] or 0) for row in group)
            shrinkage_hours = sum(float(row[26] or 0) for row in group)
            trend_components = {
                "offered": offered, "answered": answered,
                "short_abandoned": short,
                "abandoned_within_target": abandoned_in_target,
                "answered_within_target": within,
                "handled_seconds": handled,
            }
            trend_profile = profile_by_label[lob_label]
            trend_sl = evaluate_metric(
                _profile_metric(metrics_catalog, trend_profile, trend_profile.service_level_metric, end),
                trend_components,
            ).value
            trend_rows.append((
                lob_label, grain, label, min(row[0] for row in group), max(row[0] for row in group),
                offered, forecast if any(row[7] is not None for row in group) else None,
                _ratio(offered, forecast) if forecast else None,
                trend_sl, _ratio(answered, offered),
                _ratio(handled, answered), planned_hours, absence_hours,
                _ratio(absence_hours, planned_hours), shrinkage_hours,
                _ratio(shrinkage_hours, planned_hours),
            ))
    book.table(
        "TREND", "Period trend",
        "Month, ISO week and quarter summaries calculated from additive daily counters.",
        [
            "lob", "grain", "period", "start", "end", "actual_volume",
            "forecast_volume", "forecast_attainment", "service_level",
            "service_availability", "aht_seconds", "planned_hours",
            "absence_hours", "absence_rate", "shrinkage_hours", "shrinkage_rate",
        ],
        trend_rows,
    )
    detail_headers = [
        "reporting_lob", "business_date", "interval_start", "source_system", "queue",
        "service_scope", "designation", "mapping_status", "offered", "answered",
        "abandoned", "short_abandoned", "abandoned_within_target",
        "answered_within_target",
        "handled_seconds", "source_file",
    ]
    detail_rows = [tuple(row.get(header) for header in detail_headers) for row in all_source_rows]
    book.table(
        "DATA", "Mapped service data",
        "Filterable queue and interval evidence behind the LOB results.",
        detail_headers, detail_rows,
    )
    book.definitions([
        ("Actual / forecast", "Actual offered contacts / forecast contacts", "Demand realisation", "Use summed volumes"),
        ("Service level", "Configured numerator / denominator for each service profile", "Service performance", "Calculated from summed counters; never average LOB percentages"),
        ("Routed Rate", "Answered / total entered", "Ability of the service to route demand", "Not agent availability"),
        ("Weighted AHT", "Handled seconds / answered contacts", "Workload", "Never average daily AHT values"),
        ("Absence rate", "Reviewed absence hours / planned hours", "Capacity impact", "Open attendance gaps remain review cases"),
    ])
    book.audit(_audit_rows(conn, config, "realisations", start, end, [
        ("Service profiles", ", ".join(profile.profile_id for profile in selected_profiles), profiles.version),
        ("Included service scopes", "; ".join(
            f"{profile.label}: {', '.join(profile.service_scopes)}"
            for profile in selected_profiles
        ), "Queue Mapping"),
        ("Included staffing LOBs", "; ".join(
            f"{profile.label}: {', '.join(profile.staffing_lobs)}"
            for profile in selected_profiles
        ), "Service Profiles"),
    ]))
    return _finish(book, partial, target)


def build_attendance_today_workbook(
    conn: DatabaseConnection,
    config: Config,
    start: date,
    end: date,
    output: Path | None = None,
) -> Path:
    """Build the attendance callout product for the full selected period."""

    book, partial, target = _atomic_book(
        config, "attendance", "ATTENDANCE CALLOUTS", start, end, output,
    )
    totals = conn.execute(
        """SELECT count(*),
                  coalesce(sum(CASE WHEN requires_call THEN 1 ELSE 0 END),0),
                  coalesce(sum(CASE WHEN call_action='CALL_NO_SHOW' THEN 1 ELSE 0 END),0),
                  coalesce(sum(CASE WHEN call_action='CALL_LATE' THEN 1 ELSE 0 END),0),
                  coalesce(sum(CASE WHEN call_action='CALL_NOT_SEEN_NOW' THEN 1 ELSE 0 END),0),
                  coalesce(sum(CASE WHEN attendance_result IN
                    ('Schedule parse error','Data not loaded','Missing actual evidence',
                     'Incomplete actual evidence','No schedule overlap')
                    THEN 1 ELSE 0 END),0)
           FROM mart.attendance_agent_day
           WHERE business_date BETWEEN ? AND ?
             AND assignment_type NOT IN ('Off','Planned absence')""",
        [start, end],
    ).fetchone()
    scheduled, call_now, no_show, late, not_seen, missing = totals
    status, status_text = _source_state(conn, ("fte", "start_end", "lilo", "agent_status"), end)
    if missing:
        status, status_text = "INCOMPLETE", f"{missing:,} scheduled row(s) do not have complete attendance evidence"
    book.dashboard(
        [
            KpiCard("Scheduled working", scheduled, "integer", "Selected agent-day rows"),
            KpiCard("Call/action cases", call_now, "integer", "Selected-period queue"),
            KpiCard("Confirmed no-show", no_show, "integer", "Only after completed shift"),
            KpiCard("Late", late, "integer", "Beyond configured tolerance"),
            KpiCard("Not seen yet", not_seen, "integer", "Provisional current-day state"),
            KpiCard("Missing evidence", missing, "integer", "Unknown, never no-show"),
        ],
        status,
        status_text,
        ["Result", "Agents"],
        [("Scheduled working", scheduled), ("Call/action cases", call_now), ("Confirmed no-show", no_show), ("Late", late), ("Not seen yet", not_seen), ("Missing evidence", missing)],
        [
            "This workbook is the selected-period callout register; it is not an adherence report.",
            "Choose Today for a live queue or Current Week for every callout case and daily trend in that week.",
            "An unfinished shift can be late or not seen, but can never be marked as early leave.",
            "No-show requires completed-shift evidence: either a loaded blank LILO row or sufficiently covered Agent Status that stays Logged Off.",
            "Use Attendance Review only after the operating day is complete.",
        ],
        (("Agents", 1),),
        "column",
    )
    headers, rows = _query(
        conn,
        """SELECT business_date, agent_id, agent_name, team_leader, ops_manager,
                  lob, language, scheduled_start, scheduled_end, shift_state,
                  call_action, attendance_result, actual_first_seen, actual_last_seen,
                  uncoded_late_minutes, no_show_minutes, actual_evidence,
                  source_loaded, is_provisional, evaluation_as_of
           FROM mart.attendance_agent_day
           WHERE business_date BETWEEN ? AND ? AND requires_call=true
           ORDER BY CASE call_action WHEN 'CALL_NO_SHOW' THEN 1 WHEN 'CALL_LATE' THEN 2 ELSE 3 END,
                    business_date, scheduled_start, lob, team_leader, agent_name""",
        [start, end],
    )
    ws = book.table(
        "ACTIONS", "Attendance callout cases",
        "Every actionable case in the selected period, ordered by severity, date and scheduled start.",
        headers, rows,
    )
    if rows:
        col = headers.index("call_action")
        ws.conditional_format(4, col, 3 + len(rows), col, {"type": "text", "criteria": "containing", "value": "CALL_NO_SHOW", "format": book.report.error})
    trend_headers, trend_rows = _query(
        conn,
        """SELECT business_date, count(*) AS scheduled_working,
                  sum(CASE WHEN requires_call THEN 1 ELSE 0 END) AS call_actions,
                  sum(CASE WHEN call_action='CALL_NO_SHOW' THEN 1 ELSE 0 END) AS no_shows,
                  sum(CASE WHEN call_action='CALL_LATE' THEN 1 ELSE 0 END) AS late,
                  sum(CASE WHEN source_loaded=false THEN 1 ELSE 0 END) AS missing_evidence
           FROM mart.attendance_agent_day
           WHERE business_date BETWEEN ? AND ? AND assignment_type NOT IN ('Off','Planned absence')
           GROUP BY business_date ORDER BY business_date""",
        [start, end],
    )
    book.table("TREND", "Attendance action trend", "Daily counts for the requested range; today remains provisional until shifts close.", trend_headers, trend_rows)
    book.definitions([
        ("Call no-show", "Completed shift + blank LILO row, or sufficient all-Logged-Off Agent Status coverage", "Call and validate absence", "Missing source is never a no-show"),
        ("Call late", "First observed evidence after scheduled start plus tolerance", "Contact / record explanation", "Current shift may still be running"),
        ("Not seen now", "Shift started, no observed evidence yet", "Immediate operational check", "Always provisional"),
        ("Early leave", "Last observed evidence before completed scheduled end", "Historical correction only", "Never evaluated before shift end"),
    ])
    book.audit(_audit_rows(conn, config, "attendance", start, end))
    return _finish(book, partial, target)


def build_staffing_coverage_workbook(
    conn: DatabaseConnection,
    config: Config,
    start: date,
    end: date,
    output: Path | None = None,
) -> Path:
    """Build actual staffing control and future required-versus-scheduled plan."""

    book, partial, target = _atomic_book(
        config, "staffing", "STAFFING & CAPACITY PLAN", start, end, output,
    )
    profiles = load_service_profiles(config.home, config.service_profiles)
    active_profiles = tuple(profile for profile in profiles.profiles if profile.active_on(end))
    staffing_profile: dict[str, ServiceProfile] = {}
    for profile in active_profiles:
        for lob in profile.staffing_lobs:
            staffing_profile.setdefault(lob.casefold(), profile)

    forecast_by_lob_interval: dict[tuple[str, str, datetime], tuple[Any, ...]] = {}
    for profile in active_profiles:
        for service_scope, staffing_lob in profile.staffing_pairs():
            forecast_rows = conn.execute(
                """SELECT business_date, interval_start, max(interval_minutes),
                          sum(volume_forecast), sum(fte_forecast),
                          sum(fte_required), avg(sl_required)
                   FROM mart.forecast_interval
                   WHERE business_date BETWEEN ? AND ? AND service_scope=?
                   GROUP BY business_date, interval_start""",
                [start, end, service_scope],
            ).fetchall()
            for (
                business_date, interval_start, interval_minutes, volume,
                fte_forecast, fte_required, sl_required,
            ) in forecast_rows:
                stamp = interval_start
                if isinstance(stamp, str):
                    stamp = datetime.fromisoformat(stamp)
                source_minutes = max(15, int(interval_minutes or 60))
                slots = max(1, (source_minutes + 14) // 15)
                volume_per_slot = (
                    float(volume) / slots if volume is not None else None
                )
                for offset in range(slots):
                    key = (
                        staffing_lob.casefold(), str(business_date)[:10],
                        stamp + timedelta(minutes=15 * offset),
                    )
                    previous = forecast_by_lob_interval.get(key)
                    slot_volume = volume_per_slot
                    slot_fte_forecast = fte_forecast
                    slot_fte_required = fte_required
                    if previous:
                        slot_volume = (
                            float(previous[1] or 0)
                            + float(volume_per_slot or 0)
                        )
                        slot_fte_forecast = (
                            float(previous[2] or 0)
                            + float(fte_forecast or 0)
                        )
                        slot_fte_required = (
                            float(previous[3] or 0)
                            + float(fte_required or 0)
                        )
                    forecast_by_lob_interval[key] = (
                        profile, slot_volume, slot_fte_forecast,
                        slot_fte_required, sl_required,
                    )

    raw_headers, raw_rows = _query(
        conn,
        """SELECT business_date, interval_start, interval_end, lob, language,
                  scheduled_agents, observed_agents, productive_agents,
                  gross_scheduled_fte, planned_time_off_fte, scheduled_fte,
                  elapsed_scheduled_fte, observed_fte, productive_fte,
                  staffing_variance_fte, staffing_gap_fte, staffing_state,
                  evidence_basis, evaluation_as_of
           FROM mart.staffing_interval
           WHERE business_date BETWEEN ? AND ?
           ORDER BY business_date, interval_start, lob, language""",
        [start, end],
    )
    source_indexes = {header: index for index, header in enumerate(raw_headers)}
    known_languages: dict[str, set[str]] = defaultdict(set)
    evaluation_values: list[datetime] = []
    for raw in raw_rows:
        lob = str(raw[source_indexes["lob"]] or "")
        language = str(raw[source_indexes["language"]] or "Unspecified")
        known_languages[lob.casefold()].add(language)
        evaluation = raw[source_indexes["evaluation_as_of"]]
        if isinstance(evaluation, str):
            evaluation = datetime.fromisoformat(evaluation)
        if evaluation is not None:
            evaluation_values.append(evaluation)
    evaluation_default = max(evaluation_values, default=datetime.now())
    plan_headers = [
        "business_date", "iso_week", "interval_start", "interval_end",
        "reporting_lob", "roster_lob", "language", "mode",
        "forecast_volume_interval", "fte_forecast", "fte_required",
        "gross_scheduled_fte", "planned_time_off_fte", "net_scheduled_fte",
        "capacity_variance_fte", "capacity_gap_fte", "observed_fte",
        "productive_fte", "actual_gap_fte", "decision_state",
        "evidence_basis", "evaluation_as_of",
    ]
    plan_rows: list[tuple[Any, ...]] = []
    covered_intervals: set[tuple[str, str, datetime]] = set()
    for raw in raw_rows:
        values = {header: raw[index] for header, index in source_indexes.items()}
        business_date = values["business_date"]
        if isinstance(business_date, str):
            business_date = date.fromisoformat(business_date[:10])
        interval_start = values["interval_start"]
        if isinstance(interval_start, str):
            interval_start = datetime.fromisoformat(interval_start)
        interval_end = values["interval_end"]
        if isinstance(interval_end, str):
            interval_end = datetime.fromisoformat(interval_end)
        evaluation_as_of = values["evaluation_as_of"]
        if isinstance(evaluation_as_of, str):
            evaluation_as_of = datetime.fromisoformat(evaluation_as_of)
        roster_lob = str(values["lob"] or "")
        profile = staffing_profile.get(roster_lob.casefold())
        forecast = forecast_by_lob_interval.get((
            roster_lob.casefold(), business_date.isoformat(),
            interval_start,
        ))
        if forecast:
            profile = forecast[0]
        _forecast_profile, volume, fte_forecast, fte_required, _sl_required = forecast or (None, None, None, None, None)
        net_scheduled = float(values["scheduled_fte"] or 0)
        capacity_variance = net_scheduled - float(fte_required) if fte_required is not None else None
        capacity_gap = max(0.0, -capacity_variance) if capacity_variance is not None else None
        future = interval_start > evaluation_as_of
        mode = "FUTURE PLAN" if future else "ACTUAL CONTROL"
        actual_gap = values["staffing_gap_fte"]
        if profile is None:
            state = "UNMAPPED LOB"
        elif fte_required is None:
            state = "NO FORECAST"
        elif future:
            state = "FUTURE GAP" if capacity_gap and capacity_gap > 0.001 else "FUTURE OK"
        else:
            state = str(values["staffing_state"] or "DATA MISSING").replace("_", " ")
        plan_rows.append((
            business_date, f"{business_date.isocalendar().year}-W{business_date.isocalendar().week:02d}",
            interval_start, interval_end, profile.label if profile else "Unmapped",
            roster_lob, values["language"], mode, volume, fte_forecast,
            fte_required, values["gross_scheduled_fte"],
            values["planned_time_off_fte"], net_scheduled, capacity_variance,
            capacity_gap, values["observed_fte"], values["productive_fte"],
            actual_gap, state, values["evidence_basis"], evaluation_as_of,
        ))
        covered_intervals.add((roster_lob.casefold(), business_date.isoformat(), interval_start))

    # A demand interval with zero scheduled agents does not exist in the
    # staffing mart. Add it here so an empty roster can never hide a shortage.
    for (
        lob_key, business_date_text, interval_start,
    ), forecast in forecast_by_lob_interval.items():
        profile, volume, fte_forecast, fte_required, _sl_required = forecast
        roster_lob = next(
            (lob for lob in profile.staffing_lobs if lob.casefold() == lob_key), lob_key,
        )
        business_date = date.fromisoformat(business_date_text)
        language = " / ".join(sorted(known_languages.get(lob_key, {"Unspecified"})))
        key = (lob_key, business_date_text, interval_start)
        if key in covered_intervals:
            continue
        interval_end = interval_start + timedelta(minutes=15)
        required = float(fte_required or 0)
        future = interval_start > evaluation_default
        mode = "FUTURE PLAN" if future else "ACTUAL CONTROL"
        state = "FUTURE GAP" if future and required > 0 else "NO SCHEDULE"
        plan_rows.append((
            business_date,
            f"{business_date.isocalendar().year}-W{business_date.isocalendar().week:02d}",
            interval_start, interval_end, profile.label, roster_lob, language,
            mode, volume, fte_forecast, fte_required, 0.0, 0.0, 0.0,
            -required, required, None, None, required if not future else None,
            state, "Forecast demand with no scheduled roster interval",
            evaluation_default,
        ))

    plan_rows.sort(key=lambda row: (row[0], row[2], str(row[5]), str(row[6])))

    decision_state = plan_headers.index("decision_state")
    capacity_gap_column = plan_headers.index("capacity_gap_fte")
    actual_gap_column = plan_headers.index("actual_gap_fte")
    action_states = {"FUTURE GAP", "NO FORECAST", "UNMAPPED LOB", "NO SCHEDULE", "GAP", "PARTIAL GAP", "DATA MISSING"}
    actions = [row for row in plan_rows if str(row[decision_state]).upper() in action_states]
    future_rows = [row for row in plan_rows if row[plan_headers.index("mode")] == "FUTURE PLAN"]
    required_hours = sum(float(row[plan_headers.index("fte_required")] or 0) * 0.25 for row in future_rows)
    net_hours = sum(float(row[plan_headers.index("net_scheduled_fte")] or 0) * 0.25 for row in future_rows)
    pto_hours = sum(float(row[plan_headers.index("planned_time_off_fte")] or 0) * 0.25 for row in future_rows)
    gap_hours = sum(float(row[capacity_gap_column] or 0) * 0.25 for row in future_rows)
    future_gap_intervals = sum(1 for row in future_rows if row[decision_state] == "FUTURE GAP")
    no_forecast_intervals = sum(1 for row in future_rows if row[decision_state] == "NO FORECAST")
    actual_action_rows = [row for row in plan_rows if row[plan_headers.index("mode")] == "ACTUAL CONTROL"]
    actual_gap_intervals = sum(
        1 for row in actual_action_rows
        if str(row[decision_state]).upper() in {"GAP", "PARTIAL GAP"}
    )
    peak_gap = max(
        [float(row[capacity_gap_column] or 0) for row in future_rows]
        + [float(row[actual_gap_column] or 0) for row in actual_action_rows]
        + [0.0]
    )
    status, status_text = _source_state(conn, ("fte", "start_end", "forecast"), end)
    if no_forecast_intervals:
        status, status_text = "INCOMPLETE", f"{no_forecast_intervals:,} future interval(s) have no mapped forecast"
    book.dashboard(
        [
            KpiCard("Peak gap FTE", peak_gap, "decimal", "Largest 15-minute deficit"),
            KpiCard("Future gap intervals", future_gap_intervals, "integer"),
            KpiCard("Future required hours", required_hours, "decimal", "FTE-hours from Verint"),
            KpiCard("Future net scheduled", net_hours, "decimal", "After PTO and Away"),
            KpiCard("Future gap hours", gap_hours, "decimal"),
            KpiCard("PTO / Away impact", pto_hours, "decimal", "Capacity hours removed"),
            KpiCard("Actual gap intervals", actual_gap_intervals, "integer"),
            KpiCard("Forecast coverage", _ratio(required_hours - gap_hours, required_hours), "percent"),
        ],
        status,
        status_text,
        ["Measure", "Value"],
        [
            ("Required future FTE-hours", required_hours),
            ("Net scheduled future FTE-hours", net_hours),
            ("Future shortage FTE-hours", gap_hours),
            ("PTO / Away FTE-hours", pto_hours),
            ("Future gap intervals", future_gap_intervals),
            ("Actual gap intervals", actual_gap_intervals),
        ],
        [
            "Future plan compares Verint required FTE with net scheduled FTE after approved PTO and effective Away.",
            "Service-to-roster LOB links are editable in Service Profiles; no text guess is made in the report.",
            "Observed FTE is calculated from agent-seconds inside each 15-minute interval.",
            "Missing forecast remains NO FORECAST. It is never converted to zero demand.",
        ],
        (("Value", 1),),
        "column",
    )

    weekly: dict[tuple[str, str, str, str], list[tuple[Any, ...]]] = defaultdict(list)
    for row in plan_rows:
        weekly[(str(row[1]), str(row[4]), str(row[5]), str(row[6]))].append(row)
    weekly_rows = []
    for (iso_week, reporting_lob, roster_lob, language), rows in sorted(weekly.items()):
        required = sum(float(row[10] or 0) * 0.25 for row in rows)
        gross = sum(float(row[11] or 0) * 0.25 for row in rows)
        time_off = sum(float(row[12] or 0) * 0.25 for row in rows)
        net = sum(float(row[13] or 0) * 0.25 for row in rows)
        gap = sum(float(row[15] or 0) * 0.25 for row in rows)
        weekly_rows.append((
            iso_week, min(row[0] for row in rows), max(row[0] for row in rows),
            reporting_lob, roster_lob, language, required, gross, time_off, net,
            net - required if required else None, gap,
            _ratio(net, required),
            sum(1 for row in rows if row[19] == "FUTURE GAP"),
            sum(1 for row in rows if row[19] == "NO FORECAST"),
        ))
    book.table(
        "WEEKLY_PLAN", "Weekly capacity plan",
        "Required, gross scheduled, PTO/Away and net scheduled FTE-hours by ISO week, LOB and language.",
        [
            "iso_week", "start_date", "end_date", "reporting_lob", "roster_lob",
            "language", "required_fte_hours", "gross_scheduled_fte_hours",
            "planned_time_off_fte_hours", "net_scheduled_fte_hours",
            "capacity_variance_fte_hours", "capacity_gap_fte_hours",
            "forecast_coverage", "gap_intervals", "no_forecast_intervals",
        ],
        weekly_rows,
    )
    intraday = book.table(
        "INTRADAY", "15-minute staffing control and plan",
        "All selected dates. FUTURE PLAN uses forecast demand; ACTUAL CONTROL uses observed attendance evidence.",
        plan_headers, plan_rows,
    )
    if plan_rows:
        intraday.conditional_format(
            4, decision_state, 3 + len(plan_rows), decision_state,
            {"type": "text", "criteria": "containing", "value": "GAP", "format": book.report.error},
        )
    book.table(
        "ACTIONS", "Staffing exceptions",
        "Future shortages, missing forecasts, unmapped LOBs and actual gaps requiring action.",
        plan_headers, actions,
    )
    book.definitions([
        ("Required FTE", "Verint required FTE at native 15-minute grain; historical hourly files expand to four quarters", "Demand requirement", "Forecast only; missing stays blank"),
        ("Gross scheduled FTE", "Scheduled agent-seconds / 900 before time off", "Roster capacity", "FTE roster LOB/language"),
        ("Net scheduled FTE", "Gross scheduled FTE - approved PTO/effective Away FTE", "Usable planned capacity", "Planned Away affects future only"),
        ("Future capacity gap", "MAX(0, required FTE - net scheduled FTE)", "Hiring, OT or redeployment action", "15-minute interval"),
        ("Observed FTE", "Observed agent-seconds / 900", "Actual presence", "LILO + Agent Status evidence"),
        ("Productive FTE", "Productive-status seconds / 900", "Available handling capacity", "Not adherence"),
        ("Actual staffing gap", "MAX(0, elapsed net scheduled FTE - observed FTE)", "Same-day staffing deficit", "Blank for future/missing evidence"),
    ])
    book.audit(_audit_rows(conn, config, "staffing", start, end, [
        ("Service profile mapping", profiles.version, profiles.sha256),
        ("Planning grain", "15 minutes", "Weekly summary uses FTE-hours"),
    ]))
    return _finish(book, partial, target)


def _exclusive_final_components(
    conn: DatabaseConnection,
    rulebook,
    start: date,
    end: date,
    flag: str,
) -> tuple[list[str], list[tuple[Any, ...]]]:
    """Allocate overlapping final Activities once for a component breakdown."""

    columns = {
        "absence": "e.counts_as_absence",
        "shrinkage": "e.counts_as_shrinkage",
    }
    if flag not in columns:
        raise ValueError(f"Unsupported final component flag {flag!r}")
    headers, raw_rows = _query(
        conn,
        f"""SELECT e.agent_day_key, e.business_date, e.agent_id, e.agent_name,
                   e.team_leader, e.lob, e.activity, e.category,
                   e.event_start, e.event_end, d.planned_net_minutes
            FROM mart.verint_final_absence_event e
            JOIN mart.verint_final_absence_agent_day d
              ON d.agent_day_key=e.agent_day_key
            WHERE e.business_date BETWEEN ? AND ? AND {columns[flag]}=true
            ORDER BY e.business_date, e.agent_id, e.event_start, e.event_end""",
        [start, end],
    )
    events = [dict(zip(headers, row)) for row in raw_rows]
    precedence = {rule.name: index for index, rule in enumerate(rulebook.activity_rules)}
    by_day: dict[str, list[dict[str, Any]]] = defaultdict(list)
    for event in events:
        by_day[str(event["agent_day_key"])].append(event)
    totals: dict[tuple[Any, ...], int] = defaultdict(int)
    days: dict[tuple[Any, ...], set[Any]] = defaultdict(set)
    for day_events in by_day.values():
        remaining = int(day_events[0].get("planned_net_minutes") or 0)
        boundaries = sorted({
            stamp for event in day_events
            for stamp in (event["event_start"], event["event_end"])
            if stamp is not None
        })
        for left, right in zip(boundaries, boundaries[1:]):
            if right <= left:
                continue
            active = [
                event for event in day_events
                if event["event_start"] <= left and event["event_end"] >= right
            ]
            if not active:
                continue

            def rank(event: dict[str, Any]) -> tuple[int, str, str]:
                rule = rulebook.classify_activity(event.get("activity"))
                return (
                    precedence.get(rule.name if rule else "", len(precedence) + 1),
                    str(event.get("category") or "Other"),
                    str(event.get("activity") or ""),
                )

            chosen = min(active, key=rank)
            key = (
                chosen.get("category") or "OTHER",
                chosen.get("lob"), chosen.get("team_leader"),
                chosen.get("agent_id"), chosen.get("agent_name"),
            )
            allocated = min(remaining, int((right - left).total_seconds() // 60))
            if allocated <= 0:
                break
            totals[key] += allocated
            remaining -= allocated
            days[key].add(chosen.get("business_date"))
    output = [
        (*key, len(days[key]), minutes, minutes / 60.0)
        for key, minutes in totals.items()
        if minutes > 0
    ]
    output.sort(key=lambda row: (str(row[0]), str(row[1]), str(row[2]), str(row[4])))
    return (
        [
            "component", "lob", "team_leader", "agent_id", "agent_name",
            "agent_days", "minutes", "hours",
        ],
        output,
    )


def _legacy_build_final_absence_product_workbook(
    conn: DatabaseConnection,
    config: Config,
    start: date,
    end: date,
    output: Path | None = None,
) -> Path:
    """Build the final corrected Verint absence and shrinkage product."""

    book, partial, target = _atomic_book(config, "absence", "FINAL ABSENCE & SHRINKAGE", start, end, output)
    totals = conn.execute(
        """SELECT coalesce(sum(planned_net_minutes),0), coalesce(sum(final_absence_minutes),0),
                  coalesce(sum(final_vacation_minutes),0), coalesce(sum(final_unpaid_minutes),0),
                  coalesce(sum(final_shrinkage_minutes),0), coalesce(sum(final_unmapped_minutes),0),
                  coalesce(sum(CASE WHEN final_absence_day THEN 1 ELSE 0 END),0)
           FROM mart.verint_final_absence_agent_day
           WHERE business_date BETWEEN ? AND ?
             AND final_ledger_status IN ('CLEAR','ABSENCE_RECORDED')""",
        [start, end],
    ).fetchone()
    planned, absence, vacation, unpaid, shrinkage, unmapped, absence_days = totals
    unmapped = conn.execute(
        """SELECT coalesce(sum(final_unmapped_minutes),0)
           FROM mart.verint_final_absence_agent_day
           WHERE business_date BETWEEN ? AND ?""",
        [start, end],
    ).fetchone()[0]
    ledger_exceptions = conn.execute(
        """SELECT count(*) FROM mart.verint_final_absence_agent_day
           WHERE business_date BETWEEN ? AND ?
             AND final_ledger_status NOT IN ('CLEAR','ABSENCE_RECORDED')""",
        [start, end],
    ).fetchone()[0]
    uncoded_empty = conn.execute(
        """SELECT count(*) FROM mart.verint_final_absence_agent_day
           WHERE business_date BETWEEN ? AND ? AND final_ledger_status='UNCODED_EMPTY_SHIFT'""",
        [start, end],
    ).fetchone()[0]
    status, status_text = _source_state(conn, ("fte", "start_end", "activities"), end, final=True)
    if unmapped:
        status, status_text = "INCOMPLETE", f"{unmapped / 60:,.2f} hour(s) of Verint Activities are unmapped"
    if uncoded_empty:
        status, status_text = "INCOMPLETE", f"{uncoded_empty:,} scheduled shift(s) have no final code and no reliable work evidence"
    elif ledger_exceptions:
        status, status_text = "INCOMPLETE", f"{ledger_exceptions:,} final-ledger row(s) still require review"
    period_rows = []
    for period in _comparison_periods(start, end):
        row = conn.execute(
            """SELECT coalesce(sum(planned_net_minutes),0), coalesce(sum(final_absence_minutes),0),
                      coalesce(sum(final_vacation_minutes),0), coalesce(sum(final_shrinkage_minutes),0)
               FROM mart.verint_final_absence_agent_day
               WHERE business_date BETWEEN ? AND ?
                 AND final_ledger_status IN ('CLEAR','ABSENCE_RECORDED')""",
            [period.start, period.end],
        ).fetchone()
        period_rows.append((period.label, period.start, period.end, row[0] / 60, row[1] / 60, _ratio(row[1], row[0]), row[2] / 60, _ratio(row[3], row[0])))
    book.dashboard(
        [
            KpiCard("Finalized planned hours", planned / 60, "decimal"),
            KpiCard("Final absence hours", absence / 60, "decimal"),
            KpiCard("Final absence rate", _ratio(absence, planned), "percent"),
            KpiCard("Absence agent-days", absence_days, "integer"),
            KpiCard("Vacation hours", vacation / 60, "decimal"),
            KpiCard("Unpaid hours", unpaid / 60, "decimal"),
            KpiCard("Shrinkage rate", _ratio(shrinkage, planned), "percent"),
            KpiCard("Ledger exceptions", ledger_exceptions, "integer", f"{uncoded_empty:,} uncoded empty shift(s)"),
        ],
        status,
        status_text,
        ["Period", "Start", "End", "Planned Hours", "Absence Hours", "Absence Rate %", "Vacation Hours", "Shrinkage Rate %"],
        period_rows,
        [
            "This is the final corrected ledger: Verint Activities only, clipped to StartEndTimes schedule boundaries.",
            "LILO and Agent Status detect operational gaps but do not create final payroll categories.",
            "A scheduled working shift with neither a final Verint code nor reliable operational evidence is UNCODED_EMPTY_SHIFT, never a silent zero.",
            "Observed Agent Status/LILO gaps without complete Verint coverage remain ledger exceptions until corrected.",
            "A Verint code with no matching observed gap remains an exception too; the final ledger never assumes it is valid.",
            "Headline rates include finalized CLEAR/ABSENCE_RECORDED rows only; exception and current provisional rows cannot dilute the result.",
            "Overlapping Activities are unioned before totals; never sum event evidence directly.",
        ],
        (("Absence Rate", 5), ("Shrinkage Rate", 7)),
    )
    trend_headers, trend_rows = _query(
        conn,
        """SELECT business_date, sum(planned_net_minutes)/60.0 AS planned_hours,
                  sum(final_absence_minutes)/60.0 AS absence_hours,
                  CASE WHEN sum(planned_net_minutes)>0 THEN sum(final_absence_minutes)*1.0/sum(planned_net_minutes) END AS absence_rate,
                  sum(final_vacation_minutes)/60.0 AS vacation_hours,
                  sum(final_unpaid_minutes)/60.0 AS unpaid_hours,
                  sum(final_shrinkage_minutes)/60.0 AS shrinkage_hours,
                  CASE WHEN sum(planned_net_minutes)>0 THEN sum(final_shrinkage_minutes)*1.0/sum(planned_net_minutes) END AS shrinkage_rate
           FROM mart.verint_final_absence_agent_day WHERE business_date BETWEEN ? AND ?
             AND final_ledger_status IN ('CLEAR','ABSENCE_RECORDED')
           GROUP BY business_date ORDER BY business_date""",
        [start, end],
    )
    book.table("TREND", "Final absence daily trend", "Daily overlap-safe final counters.", trend_headers, trend_rows)
    detail_headers, detail_rows = _query(
        conn,
        """SELECT business_date, agent_id, agent_name, team_leader, ops_manager,
                  lob, language, planned_net_minutes/60.0 AS planned_net_hours,
                  final_absence_minutes/60.0 AS final_absence_hours,
                  final_vacation_minutes/60.0 AS final_vacation_hours,
                  final_unpaid_minutes/60.0 AS final_unpaid_hours,
                  final_shrinkage_minutes/60.0 AS final_shrinkage_hours,
                  final_unmapped_minutes/60.0 AS final_unmapped_hours,
                  final_absence_rate, final_absence_day, final_ledger_status
           FROM mart.verint_final_absence_agent_day WHERE business_date BETWEEN ? AND ?
           ORDER BY business_date, lob, team_leader, agent_name""",
        [start, end],
    )
    book.table("AGENT_DETAIL", "Final absence agent ledger", "Authoritative agent/day grain for payroll and Operations.", detail_headers, detail_rows)
    exceptions = [
        row for row in detail_rows
        if row[detail_headers.index("final_unmapped_hours")]
        or row[detail_headers.index("final_ledger_status")] not in {"CLEAR", "ABSENCE_RECORDED"}
    ]
    book.table("EXCEPTIONS", "Final-ledger exceptions", "Unmapped, unsupported, empty, uncorrected and partially corrected rows must be resolved before payroll use.", detail_headers, exceptions)
    book.definitions([
        ("Final absence rate", "Unioned classified absence minutes / planned net minutes", "Payroll absence", "Activities-only final ledger"),
        ("Final shrinkage rate", "Unioned configured shrinkage minutes / planned net minutes", "Capacity loss", "Includes configured nonproductive categories"),
        ("Planned net", "Scheduled span capped by configured standard day", "Common denominator", "StartEndTimes only"),
        ("Unmapped", "Verint Activity without an approved classification", "Rulebook action", "Blocks final status"),
        ("Uncoded empty shift", "Working shift with no final code and no reliable Agent Status/LILO evidence", "Correction completeness check", "Blocks final status; never treated as zero absence"),
        ("Uncorrected observed gap", "Agent Status/LILO proves a gap but Verint has no final code", "Correction completeness check", "Blocks final status"),
        ("Partial correction review", "Verint has a code but some observed gap minutes remain uncovered", "Correction completeness check", "Blocks final status"),
        ("Verint without observed gap", "A final Verint code has no matching Agent Status/LILO gap", "Correction validity check", "Blocks final status"),
    ])
    book.audit(_audit_rows(conn, config, "absence", start, end))
    return _finish(book, partial, target)


def _add_absence_team_view(
    book: DecisionWorkbook,
    start: date,
    end: date,
    data_start: date,
    latest: date,
) -> None:
    """Add one selector-driven view for Operations and Payroll reviewers."""

    wb = book.report.workbook
    wb.set_calc_mode("auto")
    ws = wb.add_worksheet("TEAM_VIEW")
    ws.hide_gridlines(2)
    ws.set_tab_color(COLORS["gold"])
    ws.set_zoom(85)
    ws.freeze_panes(16, 0)
    ws.merge_range("A1:AA1", "ABSENTEEISM & SHRINKAGE  /  TEAM REVIEW", book.report.title)
    ws.merge_range(
        "A2:AA2",
        f"Latest final-ledger date {latest:%Y-%m-%d}  |  select a period and team below  |  prepared by Anass ASSRI",
        book.report.subtitle,
    )
    selector_label = wb.add_format({
        "font_name": "Aptos", "font_size": 8, "bold": True,
        "font_color": COLORS["muted"], "bg_color": COLORS["canvas"],
        "align": "left", "valign": "vcenter", "indent": 1,
    })
    selector = wb.add_format({
        "font_name": "Aptos Display", "font_size": 11, "bold": True,
        "font_color": COLORS["dark"], "bg_color": COLORS["white"],
        "border": 1, "border_color": COLORS["teal"], "align": "left",
        "valign": "vcenter", "indent": 1, "num_format": "yyyy-mm-dd",
    })
    for label, label_range, value_range, value in (
        ("PERIOD VIEW", "A4:C4", "A5:C6", "Current MTD"),
        ("CUSTOM FROM", "E4:F4", "E5:F6", start),
        ("CUSTOM TO", "H4:I4", "H5:I6", end),
        ("LOB", "K4:L4", "K5:L6", "All"),
        ("TEAM LEADER", "N4:O4", "N5:O6", "All"),
        ("AGENT", "Q4:R4", "Q5:R6", "All"),
    ):
        ws.merge_range(label_range, label, selector_label)
        ws.merge_range(value_range, value, selector)
    ws.data_validation("A5", {
        "validate": "list",
        "source": [
            "Latest day", "Current week", "Previous week", "Current MTD",
            "Previous-month same days", "Previous full month", "Custom period",
        ],
    })
    ws.data_validation("E5", {
        "validate": "date", "criteria": "between", "minimum": data_start,
        "maximum": latest,
    })
    ws.data_validation("H5", {
        "validate": "date", "criteria": "between", "minimum": data_start,
        "maximum": latest,
    })
    ws.data_validation("K5", {"validate": "list", "source": "=ABS_LOB_LIST"})
    ws.data_validation("N5", {"validate": "list", "source": "=ABS_TL_LIST"})
    ws.data_validation("Q5", {"validate": "list", "source": "=ABS_AGENT_LIST"})
    wb.define_name("ABS_Latest", "=MAX(tblAbsenceData[Date])")
    wb.define_name(
        "ABS_From",
        '=IF(TEAM_VIEW!$A$5="Latest day",ABS_Latest,'
        'IF(TEAM_VIEW!$A$5="Current week",ABS_Latest-WEEKDAY(ABS_Latest,2)+1,'
        'IF(TEAM_VIEW!$A$5="Previous week",ABS_Latest-WEEKDAY(ABS_Latest,2)-6,'
        'IF(TEAM_VIEW!$A$5="Current MTD",EOMONTH(ABS_Latest,-1)+1,'
        'IF(TEAM_VIEW!$A$5="Previous-month same days",EOMONTH(ABS_Latest,-2)+1,'
        'IF(TEAM_VIEW!$A$5="Previous full month",EOMONTH(ABS_Latest,-2)+1,TEAM_VIEW!$E$5))))))',
    )
    wb.define_name(
        "ABS_To",
        '=IF(TEAM_VIEW!$A$5="Latest day",ABS_Latest,'
        'IF(TEAM_VIEW!$A$5="Current week",ABS_Latest,'
        'IF(TEAM_VIEW!$A$5="Previous week",ABS_Latest-WEEKDAY(ABS_Latest,2),'
        'IF(TEAM_VIEW!$A$5="Current MTD",ABS_Latest,'
        'IF(TEAM_VIEW!$A$5="Previous-month same days",EDATE(ABS_Latest,-1),'
        'IF(TEAM_VIEW!$A$5="Previous full month",EOMONTH(ABS_Latest,-1),TEAM_VIEW!$H$5))))))',
    )
    scope = (
        '(tblAbsenceData[Date]>=ABS_From)*(tblAbsenceData[Date]<=ABS_To)*'
        'IF(TEAM_VIEW!$K$5="All",1,--(tblAbsenceData[LOB]=TEAM_VIEW!$K$5))*'
        'IF(TEAM_VIEW!$N$5="All",1,--(tblAbsenceData[Team Leader]=TEAM_VIEW!$N$5))*'
        'IF(TEAM_VIEW!$Q$5="All",1,--(tblAbsenceData[Agent Selector]=TEAM_VIEW!$Q$5))'
    )
    final = '((tblAbsenceData[Final Ledger Status]="CLEAR")+(tblAbsenceData[Final Ledger Status]="ABSENCE_RECORDED"))'
    planned = f"SUMPRODUCT({scope}*{final}*N(tblAbsenceData[Planned Net Hours]))"
    absence = f"SUMPRODUCT({scope}*{final}*N(tblAbsenceData[Absence Hours]))"
    shrinkage = f"SUMPRODUCT({scope}*{final}*N(tblAbsenceData[Shrinkage Hours]))"
    review = (
        f'SUMPRODUCT({scope}*--(tblAbsenceData[Final Ledger Status]<>"CLEAR")*'
        '--(tblAbsenceData[Final Ledger Status]<>"ABSENCE_RECORDED"))'
    )
    ws.merge_range("A8:AA8", "SELECTED TEAM POSITION", book.report.section)
    cards = (
        ("ABSENCE RATE", f'=IFERROR({absence}/{planned},"")', book.card_percent),
        ("SHRINKAGE RATE", f'=IFERROR({shrinkage}/{planned},"")', book.card_percent),
        ("ABSENCE HOURS", f"={absence}", book.card_decimal),
        ("REVIEW CASES", f"={review}", book.card_integer),
    )
    for index, (label, formula, fmt) in enumerate(cards):
        column = index * 4
        ws.merge_range(9, column, 9, column + 2, label, book.report.kpi_label)
        ws.merge_range(10, column, 11, column + 2, "", fmt)
        ws.write_formula(10, column, formula, fmt, "")
    ws.merge_range("A14:L14", "AGENT RESULTS", book.report.section)
    agent_headers = [
        "LOB", "Team Leader", "Agent Selector", "Agent ID", "Planned Hours",
        "Absence Hours", "Absence Rate", "Shrinkage Hours", "Shrinkage Rate",
        "Vacation Hours", "Unpaid Hours", "Review Cases",
    ]
    for column, header in enumerate(agent_headers):
        ws.write(15, column, header, book.report.header)
    agent_formula = (
        '=LET(d,tblAbsenceData,'
        'm,(d[Date]>=ABS_From)*(d[Date]<=ABS_To)*'
        'IF(TEAM_VIEW!$K$5="All",1,--(d[LOB]=TEAM_VIEW!$K$5))*'
        'IF(TEAM_VIEW!$N$5="All",1,--(d[Team Leader]=TEAM_VIEW!$N$5))*'
        'IF(TEAM_VIEW!$Q$5="All",1,--(d[Agent Selector]=TEAM_VIEW!$Q$5)),'
        'f,((d[Final Ledger Status]="CLEAR")+(d[Final Ledger Status]="ABSENCE_RECORDED")),'
        'a,SORT(UNIQUE(FILTER(d[Agent Selector],m,""))),'
        'ph,MAP(a,LAMBDA(x,SUMPRODUCT(m*f*(d[Agent Selector]=x)*N(d[Planned Net Hours])))),'
        'ah,MAP(a,LAMBDA(x,SUMPRODUCT(m*f*(d[Agent Selector]=x)*N(d[Absence Hours])))),'
        'sh,MAP(a,LAMBDA(x,SUMPRODUCT(m*f*(d[Agent Selector]=x)*N(d[Shrinkage Hours])))),'
        'vh,MAP(a,LAMBDA(x,SUMPRODUCT(m*f*(d[Agent Selector]=x)*N(d[Vacation Hours])))),'
        'uh,MAP(a,LAMBDA(x,SUMPRODUCT(m*f*(d[Agent Selector]=x)*N(d[Unpaid Hours])))),'
        'rv,MAP(a,LAMBDA(x,SUMPRODUCT(m*(d[Agent Selector]=x)*--(d[Final Ledger Status]<>"CLEAR")*--(d[Final Ledger Status]<>"ABSENCE_RECORDED")))),'
        'IFERROR(HSTACK('
        'XLOOKUP(a,d[Agent Selector],d[LOB],""),'
        'XLOOKUP(a,d[Agent Selector],d[Team Leader],""),a,'
        'XLOOKUP(a,d[Agent Selector],d[Agent ID],""),ph,ah,IFERROR(ah/ph,""),'
        'sh,IFERROR(sh/ph,""),vh,uh,rv),"No matching agent data"))'
    )
    ws.write_dynamic_array_formula("A17", agent_formula, book.report.body, "Open in desktop Excel")
    ws.merge_range("N14:AA14", "CASES TO REVIEW", book.report.section)
    queue_headers = [
        "Case ID", "Date", "Agent ID", "Agent", "Team Leader", "LOB",
        "Result Status", "Absence Hours", "Shrinkage Hours", "Unmapped Hours",
        "Action Status",
    ]
    for column, header in enumerate(queue_headers, 13):
        ws.write(15, column, header, book.report.header)
    queue_formula = (
        '=LET(q,tblActionQueue,'
        'm,(q[Date]>=ABS_From)*(q[Date]<=ABS_To)*'
        'IF(TEAM_VIEW!$K$5="All",1,--(q[LOB]=TEAM_VIEW!$K$5))*'
        'IF(TEAM_VIEW!$N$5="All",1,--(q[Team Leader]=TEAM_VIEW!$N$5))*'
        'IF(TEAM_VIEW!$Q$5="All",1,--(q[Agent Selector]=TEAM_VIEW!$Q$5)),'
        'IFERROR(FILTER(CHOOSECOLS(q,1,2,3,4,6,8,10,12,13,14,15),m),'
        '"No cases in this selection"))'
    )
    ws.write_dynamic_array_formula("N17", queue_formula, book.report.body, "Open in desktop Excel")
    ws.write_url(
        "T10", "internal:'ACTIONS'!A1", book.report.editable,
        string="OPEN PERMANENT ACTION LOG",
    )
    ws.set_column("A:A", 18)
    ws.set_column("B:B", 22)
    ws.set_column("C:C", 30)
    ws.set_column("D:L", 15)
    ws.set_column("M:M", 3)
    ws.set_column("N:AA", 18)
    ws.set_landscape()
    ws.fit_to_pages(1, 0)


def _add_absence_component_view(book: DecisionWorkbook) -> None:
    """Show filtered absence and shrinkage components from exact intervals."""

    wb = book.report.workbook
    ws = wb.add_worksheet("COMPONENT_VIEW")
    ws.hide_gridlines(2)
    ws.set_tab_color(COLORS["gold"])
    ws.set_zoom(90)
    ws.merge_range("A1:N1", "ABSENCE & SHRINKAGE  /  COMPONENT VIEW", book.report.title)
    ws.merge_range(
        "A2:N2", "This view follows TEAM_VIEW period, LOB, Team Leader and Agent selectors.",
        book.report.subtitle,
    )
    ws.merge_range("A4:F4", "ABSENCE COMPONENTS", book.report.section)
    ws.merge_range("H4:N4", "SHRINKAGE COMPONENTS", book.report.section)
    headers = ("Component", "Hours", "Intervals")
    for column, header in enumerate(headers):
        ws.write(5, column, header, book.report.header)
        ws.write(5, column + 7, header, book.report.header)
    base_scope = (
        '(t[Date]>=ABS_From)*(t[Date]<=ABS_To)*'
        'IF(TEAM_VIEW!$K$5="All",1,--(t[LOB]=TEAM_VIEW!$K$5))*'
        'IF(TEAM_VIEW!$N$5="All",1,--(t[Team Leader]=TEAM_VIEW!$N$5))*'
        'IF(TEAM_VIEW!$Q$5="All",1,--(t[Agent Selector]=TEAM_VIEW!$Q$5))'
    )
    for cell, flag, empty_text in (
        ("A7", "Counts As Absence", "No absence components in this selection"),
        ("H7", "Counts As Shrinkage", "No shrinkage components in this selection"),
    ):
        formula = (
            '=LET(t,tblActivityDetail,'
            f'm,{base_scope}*--(t[{flag}]=TRUE),'
            'c,SORT(UNIQUE(FILTER(t[Category],m,""))),'
            'h,MAP(c,LAMBDA(x,SUMPRODUCT(m*(t[Category]=x)*N(t[Hours])))),'
            'n,MAP(c,LAMBDA(x,SUMPRODUCT(m*(t[Category]=x)))),'
            f'IFERROR(HSTACK(c,h,n),"{empty_text}"))'
        )
        ws.write_dynamic_array_formula(cell, formula, book.report.body, "Open in desktop Excel")
    ws.merge_range(
        "A22:N22", "Exact start/end evidence remains on ACTIVITY_DETAIL. Component hours are exclusive inside each KPI view; never add Absence Rate and Shrinkage Rate together.",
        book.report.note,
    )
    ws.write_url(
        "A24", "internal:'ACTIVITY_DETAIL'!A1", book.report.editable,
        string="OPEN EXACT VERINT ACTIVITY DETAIL",
    )
    ws.set_column("A:A", 28)
    ws.set_column("B:C", 15)
    ws.set_column("D:G", 4)
    ws.set_column("H:H", 28)
    ws.set_column("I:J", 15)


def _add_absence_lookups(book: DecisionWorkbook) -> None:
    """Create cascading LOB, Team Leader and Agent lists for Absenteeism."""

    wb = book.report.workbook
    ws = wb.add_worksheet("_LOOKUPS")
    ws.write("A1", "All")
    ws.write_dynamic_array_formula(
        "A2", '=SORT(UNIQUE(FILTER(tblAbsenceData[LOB],tblAbsenceData[LOB]<>"","")))',
    )
    ws.write("B1", "All")
    ws.write_dynamic_array_formula(
        "B2",
        '=SORT(UNIQUE(FILTER(tblAbsenceData[Team Leader],'
        '(tblAbsenceData[Team Leader]<>"")*IF(TEAM_VIEW!$K$5="All",1,'
        'tblAbsenceData[LOB]=TEAM_VIEW!$K$5),"")))',
    )
    ws.write("C1", "All")
    ws.write_dynamic_array_formula(
        "C2",
        '=SORT(UNIQUE(FILTER(tblAbsenceData[Agent Selector],'
        '(tblAbsenceData[Agent Selector]<>"")*IF(TEAM_VIEW!$K$5="All",1,'
        'tblAbsenceData[LOB]=TEAM_VIEW!$K$5)*IF(TEAM_VIEW!$N$5="All",1,'
        'tblAbsenceData[Team Leader]=TEAM_VIEW!$N$5),"")))',
    )
    wb.define_name("ABS_LOB_LIST", "=_LOOKUPS!$A$1:INDEX(_LOOKUPS!$A:$A,COUNTA(_LOOKUPS!$A:$A))")
    wb.define_name("ABS_TL_LIST", "=_LOOKUPS!$B$1:INDEX(_LOOKUPS!$B:$B,COUNTA(_LOOKUPS!$B:$B))")
    wb.define_name("ABS_AGENT_LIST", "=_LOOKUPS!$C$1:INDEX(_LOOKUPS!$C:$C,COUNTA(_LOOKUPS!$C:$C))")
    ws.hide()


def build_final_absence_product_workbook(
    conn: DatabaseConnection,
    config: Config,
    start: date,
    end: date,
    output: Path | None = None,
) -> Path:
    """Build the reviewed Absenteeism and Shrinkage collaboration report."""

    from .shared_feeds import publish_absence_feeds

    publish_absence_feeds(conn, config, start, end)
    rulebook = load_rulebook(config.home, config.business_rules)
    book, partial, target = _atomic_book(
        config, "absence", "ABSENTEEISM & SHRINKAGE", start, end, output,
    )
    available_start, available_end = conn.execute(
        "SELECT min(business_date), max(business_date) FROM mart.verint_final_absence_agent_day"
    ).fetchone()
    data_start = (
        available_start if isinstance(available_start, date)
        else date.fromisoformat(str(available_start)[:10]) if available_start else start
    )
    latest = (
        available_end if isinstance(available_end, date)
        else date.fromisoformat(str(available_end)[:10]) if available_end else end
    )
    totals = conn.execute(
        """SELECT coalesce(sum(planned_net_minutes),0),
                  coalesce(sum(final_absence_minutes),0),
                  coalesce(sum(final_vacation_minutes),0),
                  coalesce(sum(final_unpaid_minutes),0),
                  coalesce(sum(final_shrinkage_minutes),0),
                  coalesce(sum(final_unmapped_minutes),0),
                  coalesce(sum(CASE WHEN final_absence_day THEN 1 ELSE 0 END),0),
                  count(*),
                  coalesce(sum(CASE WHEN final_ledger_status IN ('CLEAR','ABSENCE_RECORDED')
                                    THEN planned_net_minutes ELSE 0 END),0),
                  coalesce(sum(CASE WHEN final_ledger_status NOT IN ('CLEAR','ABSENCE_RECORDED')
                                    THEN 1 ELSE 0 END),0)
           FROM mart.verint_final_absence_agent_day
           WHERE business_date BETWEEN ? AND ?""",
        [start, end],
    ).fetchone()
    (
        all_planned, all_absence, all_vacation, all_unpaid, all_shrinkage,
        unmapped, absence_days, agent_days, finalized_planned, exceptions,
    ) = totals
    finalized = conn.execute(
        """SELECT coalesce(sum(planned_net_minutes),0),
                  coalesce(sum(final_absence_minutes),0),
                  coalesce(sum(final_vacation_minutes),0),
                  coalesce(sum(final_unpaid_minutes),0),
                  coalesce(sum(final_shrinkage_minutes),0)
           FROM mart.verint_final_absence_agent_day
           WHERE business_date BETWEEN ? AND ?
             AND final_ledger_status IN ('CLEAR','ABSENCE_RECORDED')""",
        [start, end],
    ).fetchone()
    planned, absence, vacation, unpaid, shrinkage = finalized
    uncoded_empty = conn.execute(
        """SELECT count(*) FROM mart.verint_final_absence_agent_day
           WHERE business_date BETWEEN ? AND ?
             AND final_ledger_status='UNCODED_EMPTY_SHIFT'""",
        [start, end],
    ).fetchone()[0]
    status, status_text = _source_state(
        conn, ("fte", "start_end", "lilo", "agent_status"), end, final=True,
    )
    if unmapped:
        status, status_text = "INCOMPLETE", f"{unmapped / 60:,.2f} exact gap hour(s) still need a human decision"
    elif uncoded_empty:
        status, status_text = "INCOMPLETE", f"{uncoded_empty:,} scheduled shift(s) lack reliable attendance evidence"
    elif exceptions:
        status, status_text = "INCOMPLETE", f"{exceptions:,} case(s) still require review"

    period_rows = []
    for period in _comparison_periods(start, end):
        values = conn.execute(
            """SELECT coalesce(sum(planned_net_minutes),0),
                      coalesce(sum(final_absence_minutes),0),
                      coalesce(sum(final_vacation_minutes),0),
                      coalesce(sum(final_shrinkage_minutes),0)
               FROM mart.verint_final_absence_agent_day
               WHERE business_date BETWEEN ? AND ?
                 AND final_ledger_status IN ('CLEAR','ABSENCE_RECORDED')""",
            [period.start, period.end],
        ).fetchone()
        period_rows.append((
            period.label, period.start, period.end, values[0] / 60,
            values[1] / 60, _ratio(values[1], values[0]), values[2] / 60,
            values[3] / 60, _ratio(values[3], values[0]),
        ))
    coverage = _ratio(finalized_planned, all_planned)
    book.dashboard(
        [
            KpiCard("Finalized planned hours", planned / 60, "decimal"),
            KpiCard("Absence hours", absence / 60, "decimal"),
            KpiCard("Absence rate", _ratio(absence, planned), "percent"),
            KpiCard("Absence agent-days", absence_days, "integer"),
            KpiCard("Shrinkage hours", shrinkage / 60, "decimal"),
            KpiCard("Shrinkage rate", _ratio(shrinkage, planned), "percent"),
            KpiCard("Vacation / unpaid", (vacation + unpaid) / 60, "decimal"),
            KpiCard("Finalized coverage", coverage, "percent", f"{exceptions:,} review case(s)"),
        ],
        status,
        status_text,
        [
            "Period", "Start", "End", "Planned Hours", "Absence Hours",
            "Absence Rate %", "Vacation Hours", "Shrinkage Hours",
            "Shrinkage Rate %",
        ],
        period_rows,
        [
            "Results use exact Agent Status/LILO gaps inside schedule boundaries and the imported Attendance Review decisions.",
            "Agent Status is primary evidence; LILO fills missing coverage. Neither source assigns a reason without human review.",
            "Absence and shrinkage are parallel views. Do not add their percentages together.",
            "Open ACTIONS for unresolved cases and use ACTIVITY_DETAIL when an exact interval needs investigation.",
        ],
        (("Absence rate", 5), ("Shrinkage rate", 8)),
    )
    _add_absence_team_view(book, start, end, data_start, latest)

    team_headers, team_rows = _query(
        conn,
        """SELECT lob, team_leader, count(DISTINCT agent_id) AS agents,
                  count(*) AS agent_days,
                  sum(planned_net_minutes)/60.0 AS planned_hours,
                  sum(final_absence_minutes)/60.0 AS absence_hours,
                  CASE WHEN sum(planned_net_minutes)>0
                       THEN sum(final_absence_minutes)*1.0/sum(planned_net_minutes) END AS absence_rate,
                  sum(final_shrinkage_minutes)/60.0 AS shrinkage_hours,
                  CASE WHEN sum(planned_net_minutes)>0
                       THEN sum(final_shrinkage_minutes)*1.0/sum(planned_net_minutes) END AS shrinkage_rate,
                  sum(final_vacation_minutes)/60.0 AS vacation_hours,
                  sum(final_unpaid_minutes)/60.0 AS unpaid_hours,
                  sum(CASE WHEN final_ledger_status NOT IN ('CLEAR','ABSENCE_RECORDED')
                           THEN 1 ELSE 0 END) AS review_cases
           FROM mart.verint_final_absence_agent_day
           WHERE business_date BETWEEN ? AND ?
           GROUP BY lob, team_leader
           ORDER BY lob, team_leader""",
        [start, end],
    )
    team_sheet = book.table(
        "TEAM_SUMMARY", "Team absence and shrinkage",
        "Filter LOB or Team Leader. Rates use the summed hours shown in the same row.",
        team_headers, team_rows,
    )
    if team_rows:
        for name in ("absence_rate", "shrinkage_rate"):
            column = team_headers.index(name)
            team_sheet.conditional_format(
                4, column, 3 + len(team_rows), column,
                {"type": "3_color_scale", "min_color": COLORS["green_light"],
                 "mid_color": COLORS["amber_light"], "max_color": COLORS["red_light"]},
            )

    agent_headers, agent_rows = _query(
        conn,
        """SELECT lob, team_leader,
                  coalesce(agent_name,'Agent') || ' [' || agent_id || ']' AS agent_selector,
                  agent_id, agent_name, count(*) AS scheduled_days,
                  sum(planned_net_minutes)/60.0 AS planned_hours,
                  sum(final_absence_minutes)/60.0 AS absence_hours,
                  CASE WHEN sum(planned_net_minutes)>0
                       THEN sum(final_absence_minutes)*1.0/sum(planned_net_minutes) END AS absence_rate,
                  sum(final_shrinkage_minutes)/60.0 AS shrinkage_hours,
                  CASE WHEN sum(planned_net_minutes)>0
                       THEN sum(final_shrinkage_minutes)*1.0/sum(planned_net_minutes) END AS shrinkage_rate,
                  sum(final_vacation_minutes)/60.0 AS vacation_hours,
                  sum(final_unpaid_minutes)/60.0 AS unpaid_hours,
                  sum(CASE WHEN final_absence_day THEN 1 ELSE 0 END) AS absence_days,
                  sum(CASE WHEN final_ledger_status NOT IN ('CLEAR','ABSENCE_RECORDED')
                           THEN 1 ELSE 0 END) AS review_cases
           FROM mart.verint_final_absence_agent_day
           WHERE business_date BETWEEN ? AND ?
           GROUP BY lob, team_leader, agent_id, agent_name
           ORDER BY lob, team_leader, agent_name, agent_id""",
        [start, end],
    )
    book.table(
        "AGENT_RESULTS", "Agent absence and shrinkage",
        "Use a personal Sheet View before filtering LOB, Team Leader or Agent Selector.",
        agent_headers, agent_rows,
    )

    action_headers, action_rows = _query(
        conn,
        """SELECT lob, team_leader,
                  coalesce(agent_name,'Agent') || ' [' || agent_id || ']' AS agent_selector,
                  business_date, agent_id, agent_name, final_ledger_status,
                  planned_net_minutes/60.0 AS planned_net_hours,
                  final_absence_minutes/60.0 AS absence_hours,
                  final_shrinkage_minutes/60.0 AS shrinkage_hours,
                  final_unmapped_minutes/60.0 AS unmapped_hours,
                  CASE WHEN final_ledger_status IN ('CLEAR','ABSENCE_RECORDED')
                       THEN 'Pending' ELSE 'Needs review' END AS review_status,
                  NULL AS owner, NULL AS due_date, NULL AS action,
                  NULL AS comment, agent_day_key AS case_id
           FROM mart.verint_final_absence_agent_day
           WHERE business_date BETWEEN ? AND ?
             AND (final_absence_day=true
                  OR final_ledger_status NOT IN ('CLEAR','ABSENCE_RECORDED'))
           ORDER BY business_date DESC, lob, team_leader, agent_name""",
        [start, end],
    )
    editable = ("Review Status", "Owner", "Due Date", "Action", "Comment")
    action_rows = _carry_table_values(
        action_headers, action_rows, "Case ID", editable,
        _previous_table_values(target, "ACTIONS", "Case ID", editable),
    )
    action_sheet = book.table(
        "ACTIONS", "Absence review and follow-up",
        "Filter your team and complete only the blue columns. Case ID keeps saved work attached to the correct agent-day.",
        action_headers, action_rows,
        editable_headers=set(editable),
    )
    if action_rows:
        review_column = action_headers.index("review_status")
        action_sheet.data_validation(
            4, review_column, 3 + len(action_rows), review_column,
            {"validate": "list", "source": ["Pending", "Needs review", "In progress", "Resolved", "No action"]},
        )

    queue_headers, queue_source = _query(
        conn,
        """SELECT agent_day_key AS case_id, business_date,
                  agent_id, agent_name,
                  coalesce(agent_name,'Agent') || ' [' || agent_id || ']' AS agent_selector,
                  team_leader, ops_manager, lob, language,
                  final_ledger_status AS result_status,
                  planned_net_minutes/60.0 AS planned_net_hours,
                  final_absence_minutes/60.0 AS absence_hours,
                  final_shrinkage_minutes/60.0 AS shrinkage_hours,
                  final_unmapped_minutes/60.0 AS unmapped_hours
           FROM mart.verint_final_absence_agent_day
           WHERE business_date BETWEEN ? AND ?
             AND (final_absence_day=true
                  OR final_ledger_status NOT IN ('CLEAR','ABSENCE_RECORDED'))
           ORDER BY business_date DESC, lob, team_leader, agent_name""",
        [data_start, latest],
    )
    queue_headers.append("action_status")
    queue_rows = [
        (*row, '=IFERROR(XLOOKUP([@[Case ID]],tblActions[Case ID],tblActions[Review Status]),"Not started")')
        for row in queue_source
    ]
    queue_sheet = book.table(
        "ACTION_QUEUE", "Absence action queue",
        "Refreshable cases and recorded absence days. Action Status reads the permanent ACTIONS log by Case ID.",
        queue_headers, queue_rows or [tuple(None for _ in queue_headers)],
    )
    if queue_rows:
        action_status_column = queue_headers.index("action_status")
        queue_sheet.conditional_format(
            4, action_status_column, 3 + len(queue_rows), action_status_column,
            {"type": "text", "criteria": "containing", "value": "Not started", "format": book.report.error},
        )

    absence_headers, absence_components = _exclusive_final_components(
        conn, rulebook, start, end, "absence",
    )
    shrinkage_headers, shrinkage_components = _exclusive_final_components(
        conn, rulebook, start, end, "shrinkage",
    )
    book.table(
        "ABSENCE_COMPONENTS", "Absence components",
        "Overlapping reviewed intervals are counted once so component hours reconcile to the selected absence scope.",
        absence_headers, absence_components,
    )
    book.table(
        "SHRINKAGE_COMPONENTS", "Shrinkage components",
        "Overlapping reviewed intervals are counted once inside the shrinkage view. Do not add this table to Absence Components.",
        shrinkage_headers, shrinkage_components,
    )
    _add_absence_component_view(book)

    activity_headers, activity_rows = _query(
        conn,
        """SELECT business_date, lob, team_leader,
                  coalesce(agent_name,'Agent') || ' [' || agent_id || ']' AS agent_selector,
                  agent_id, agent_name, activity, category, event_start, event_end,
                  minutes, hours, counts_as_absence, counts_as_vacation,
                  counts_as_unpaid, counts_as_shrinkage, mapped,
                  evidence_type, event_key
           FROM mart.verint_final_absence_event
           WHERE business_date BETWEEN ? AND ?
           ORDER BY business_date, lob, team_leader, agent_name, event_start""",
        [data_start, latest],
    )
    book.table(
        "ACTIVITY_DETAIL", "Reviewed attendance component detail",
        "Exact classified decision and PTO/Away intervals. Use component sheets or ABSENCE_DATA for totals.",
        activity_headers, activity_rows,
    )

    data_headers, data_rows = _query(
        conn,
        """SELECT business_date, lob, team_leader,
                  coalesce(agent_name,'Agent') || ' [' || agent_id || ']' AS agent_selector,
                  agent_id, agent_name, ops_manager, language, location,
                  scheduled_minutes/60.0 AS scheduled_hours,
                  planned_net_minutes/60.0 AS planned_net_hours,
                  final_absence_minutes/60.0 AS absence_hours,
                  final_vacation_minutes/60.0 AS vacation_hours,
                  final_unpaid_minutes/60.0 AS unpaid_hours,
                  final_shrinkage_minutes/60.0 AS shrinkage_hours,
                  final_unmapped_minutes/60.0 AS unmapped_hours,
                  final_absence_rate AS absence_rate, final_absence_day,
                  final_ledger_status, agent_day_key AS case_id
           FROM mart.verint_final_absence_agent_day
           WHERE business_date BETWEEN ? AND ?
           ORDER BY business_date, lob, team_leader, agent_name""",
        [data_start, latest],
    )
    book.table(
        "ABSENCE_DATA", "Absence clean data",
        "One agent per day. Use a personal Sheet View before filtering LOB, Team Leader or Agent Selector.",
        data_headers, data_rows,
    )
    book.table(
        "HELP", "How to use this report", "A short operating guide for the shared workbook.",
        ["Step", "What to do", "Why"],
        [
            (1, "Run WFM Hub refresh and confirm the latest source date.", "Updates Agent Status, LILO, schedules and clean feeds."),
            (2, "Build Attendance Review, classify exact gaps and import the saved workbook.", "Turns observed gaps into reviewed components."),
            (3, "Use TEAM_VIEW for period, LOB, Team Leader and Agent selection.", "Agent results, cases and components follow one selection."),
            (4, "For a permanent shared file, connect ABSENCE_DATA, ACTION_QUEUE and ACTIVITY_DETAIL once to the fixed CSV feeds.", "Data > Refresh All updates facts without replacing the workbook."),
            (5, "Review Finalized coverage and unresolved cases before sharing totals.", "Open decisions must not dilute the rate."),
            (6, "Use COMPONENT_VIEW for totals and ACTIVITY_DETAIL for exact intervals.", "Raw intervals may overlap; KPI components remain separate."),
        ],
    )
    book.definitions([
        ("Absence rate", "Final absence minutes / finalized planned net minutes", "Payroll and attendance result", "Incomplete cases are shown separately"),
        ("Shrinkage rate", "Final shrinkage minutes / finalized planned net minutes", "Capacity loss", "A parallel view; do not add to absence rate"),
        ("Finalized coverage", "Finalized planned minutes / all planned minutes", "Confidence in the headline", "Review when below 100%"),
        ("Pending review", "Observed exact gap with no imported Approved or Dismissed decision", "Review completeness", "Never treated as zero absence"),
        ("Component", "One exclusive activity classification inside its KPI view", "Management breakdown", "Raw overlapping intervals are counted once"),
    ])
    _add_absence_lookups(book)
    book.audit(_audit_rows(
        conn, config, "absence", start, end,
        (
            ("Shared feed", str(config.feed / "Absenteeism"), "Updated with this report"),
            ("Template version", "absence-2026.09.2", "Collaboration report contract"),
            ("All planned hours", all_planned / 60, f"{agent_days:,} agent-day row(s)"),
            ("All absence hours", all_absence / 60, "Includes review rows"),
            ("All shrinkage hours", all_shrinkage / 60, "Includes review rows"),
            ("All vacation hours", all_vacation / 60, "Includes review rows"),
            ("All unpaid hours", all_unpaid / 60, "Includes review rows"),
        ),
    ))
    return _finish(book, partial, target)


def _break_meal_control_rows(
    conn: DatabaseConnection,
    config: Config,
    start: date,
    completed_through: date,
) -> tuple[list[str], list[tuple[Any, ...]]]:
    """Build one evidence-gated break/meal result per completed agent-day."""

    headers = [
        "business_date", "lob", "team_leader", "agent_id", "agent_name",
        "scheduled_start", "scheduled_end", "status_coverage_percent",
        "break_minutes", "break_allowance_minutes", "break_overrun_minutes",
        "break_spells", "longest_break_minutes", "meal_minutes",
        "meal_allowance_minutes", "meal_overrun_minutes", "meal_spells",
        "longest_meal_minutes", "alert", "evidence",
    ]
    if completed_through < start:
        return headers, []
    attendance_headers, attendance_rows = _query(
        conn,
        """SELECT business_date, lob, team_leader, agent_id, agent_name,
                  scheduled_start, scheduled_end, planned_work_minutes,
                  status_covered_minutes, actual_evidence
           FROM mart.attendance_agent_day
           WHERE business_date BETWEEN ? AND ?
             AND assignment_type NOT IN ('Off','Planned absence')
             AND scheduled_start IS NOT NULL AND scheduled_end IS NOT NULL
             AND planned_work_minutes>0
           ORDER BY business_date, lob, team_leader, agent_name""",
        [start, completed_through],
    )
    timeline_headers, timeline_rows = _query(
        conn,
        """SELECT business_date, agent_id, actual_category,
                  segment_start, segment_end
           FROM mart.shift_timeline_segment
           WHERE business_date BETWEEN ? AND ?
             AND actual_category IN ('Break','Lunch')
             AND mismatch_type<>'WORK_DURING_TIME_OFF'
           ORDER BY business_date, agent_id, actual_category, segment_start""",
        [start, completed_through],
    )
    spells: dict[tuple[Any, str, str], list[tuple[datetime, datetime]]] = defaultdict(list)

    def stamp(value: Any) -> datetime:
        return value if isinstance(value, datetime) else datetime.fromisoformat(str(value))

    for values in timeline_rows:
        item = dict(zip(timeline_headers, values))
        spells[(
            item["business_date"], str(item["agent_id"]),
            str(item["actual_category"]),
        )].append((stamp(item["segment_start"]), stamp(item["segment_end"])))

    def merged(values: Sequence[tuple[datetime, datetime]]) -> list[tuple[datetime, datetime]]:
        output: list[list[datetime]] = []
        for left, right in sorted(values):
            if right <= left:
                continue
            if not output or left > output[-1][1]:
                output.append([left, right])
            else:
                output[-1][1] = max(output[-1][1], right)
        return [(left, right) for left, right in output]

    output: list[tuple[Any, ...]] = []
    alert_order = {
        "BREAK & MEAL EXCEEDED": 0, "BREAK EXCEEDED": 1,
        "MEAL EXCEEDED": 2, "INSUFFICIENT EVIDENCE": 3,
        "WITHIN LIMIT": 4,
    }
    for values in attendance_rows:
        row = dict(zip(attendance_headers, values))
        key = (row["business_date"], str(row["agent_id"]))
        break_spells = merged(spells.get((*key, "Break"), ()))
        meal_spells = merged(spells.get((*key, "Lunch"), ()))
        break_minutes = int(sum(
            (right - left).total_seconds() for left, right in break_spells
        ) // 60)
        meal_minutes = int(sum(
            (right - left).total_seconds() for left, right in meal_spells
        ) // 60)
        longest_break = max((
            int((right - left).total_seconds() // 60)
            for left, right in break_spells
        ), default=0)
        longest_meal = max((
            int((right - left).total_seconds() // 60)
            for left, right in meal_spells
        ), default=0)
        planned = float(row.get("planned_work_minutes") or 0)
        coverage = min(
            1.0, float(row.get("status_covered_minutes") or 0) / planned,
        ) if planned else None
        reliable = bool(
            coverage is not None
            and coverage >= config.rules.minimum_status_coverage
            and "AGENT_STATUS" in str(row.get("actual_evidence") or "").upper()
        )
        break_overrun = (
            max(0, break_minutes - config.rules.break_minutes)
            if reliable else None
        )
        meal_overrun = (
            max(0, meal_minutes - config.rules.lunch_minutes)
            if reliable else None
        )
        if not reliable:
            alert = "INSUFFICIENT EVIDENCE"
        elif break_overrun and meal_overrun:
            alert = "BREAK & MEAL EXCEEDED"
        elif break_overrun:
            alert = "BREAK EXCEEDED"
        elif meal_overrun:
            alert = "MEAL EXCEEDED"
        else:
            alert = "WITHIN LIMIT"
        output.append((
            row.get("business_date"), row.get("lob"), row.get("team_leader"),
            row.get("agent_id"), row.get("agent_name"),
            row.get("scheduled_start"), row.get("scheduled_end"), coverage,
            break_minutes, config.rules.break_minutes, break_overrun,
            len(break_spells), longest_break, meal_minutes,
            config.rules.lunch_minutes, meal_overrun, len(meal_spells),
            longest_meal, alert, row.get("actual_evidence"),
        ))
    output.sort(key=lambda row: (
        alert_order.get(str(row[18]), 9), row[0], str(row[1] or ""),
        str(row[2] or ""), str(row[4] or ""),
    ))
    return headers, output


def build_attendance_corrections_workbook(
    conn: DatabaseConnection,
    config: Config,
    start: date,
    end: date,
    output: Path | None = None,
) -> Path:
    """Build the editable, exact-gap Attendance Review decision ledger."""

    from .shift_view import add_review_board

    today = date.today()
    completed_through = min(end, today - timedelta(days=1))
    book, partial, target = _atomic_book(
        config, "corrections", "ATTENDANCE REVIEW", start, end, output,
    )
    # The rulebook is audit authority even when the selected period has no
    # reviewable gaps. Load it before conditional validation setup so an empty
    # Attendance Review can still be generated and audited.
    rulebook = load_rulebook(config.home, config.business_rules)
    gap_count, gap_minutes, agents, approved, dismissed, open_count = conn.execute(
        """SELECT count(*), coalesce(sum(gap_minutes),0), count(DISTINCT agent_id),
                  coalesce(sum(CASE WHEN validation_status='Approved' THEN 1 ELSE 0 END),0),
                  coalesce(sum(CASE WHEN validation_status='Dismissed' THEN 1 ELSE 0 END),0),
                  coalesce(sum(CASE WHEN validation_status='Open' THEN 1 ELSE 0 END),0)
           FROM mart.correction_candidate
           WHERE business_date BETWEEN ? AND ? AND business_date<?
             AND gap_start IS NOT NULL AND gap_end IS NOT NULL""",
        [start, end, today],
    ).fetchone()
    missing = conn.execute(
        """SELECT count(*) FROM mart.attendance_agent_day
           WHERE business_date BETWEEN ? AND ? AND business_date<?
             AND assignment_type NOT IN ('Off','Planned absence')
             AND attendance_result IN
               ('Schedule parse error','Data not loaded','Missing actual evidence',
                'Incomplete actual evidence','No schedule overlap')""",
        [start, end, today],
    ).fetchone()[0]
    status, status_text = _source_state(
        conn, ("fte", "start_end", "lilo", "agent_status"), completed_through,
        final=True,
    )
    if open_count or missing:
        status = "INCOMPLETE"
        status_text = (
            f"{open_count:,} exact gap(s) await a decision; "
            f"{missing:,} scheduled row(s) lack complete evidence"
        )
    book.dashboard(
        [
            KpiCard("Exact gaps", gap_count, "integer"),
            KpiCard("Gap hours", gap_minutes / 60 if gap_minutes else 0, "decimal"),
            KpiCard("Agents", agents, "integer"),
            KpiCard("Open decisions", open_count, "integer"),
            KpiCard("Approved", approved, "integer"),
            KpiCard("Dismissed", dismissed, "integer"),
            KpiCard("Missing evidence", missing, "integer"),
        ],
        status,
        status_text,
        ["Measure", "Value"],
        [
            ("Exact gaps", gap_count), ("Gap hours", gap_minutes / 60 if gap_minutes else 0),
            ("Open", open_count), ("Approved", approved), ("Dismissed", dismissed),
        ],
        [
            f"Every completed date from {start:%Y-%m-%d} through {completed_through:%Y-%m-%d} is included; today is excluded.",
            "Agent Status is the primary evidence. LILO fills missing coverage and acts as a control; extracts are never edited.",
            "Each Gap ID owns one exact interval. SCHEDULE sits above ACTUAL; edit only the five blue cells on the ACTUAL row.",
            "Approved uses the selected rulebook category; Dismissed counts as no loss; Open remains unverified.",
            "Import this same workbook from the Attendance Review menu to recalculate absence and shrinkage.",
            "Teal is scheduled work. Dark red is this ACTUAL row's exact gap; light red is another gap; grey is inside tolerance.",
            "BREAK & MEAL totals completed-shift Agent Status spells and judges overruns only when coverage is sufficient.",
        ],
        sheet_name="CONTROL",
    )
    headers, rows = _query(
        conn,
        """SELECT c.correction_id AS gap_id, c.business_date, c.agent_id,
                  c.agent_name, c.team_leader, c.lob,
                  c.detected_issue, c.gap_start AS exact_start,
                  c.gap_end AS exact_end, c.gap_minutes AS minutes,
                  c.suggested_activity,
                  c.confirmed_activity AS decision_category,
                  c.validation_status AS decision_status,
                  c.owner AS reviewed_by, c.comment,
                  c.injected_date AS reviewed_date
           FROM mart.correction_candidate c
           WHERE c.business_date BETWEEN ? AND ? AND c.business_date<?
             AND c.gap_start IS NOT NULL AND c.gap_end IS NOT NULL
           ORDER BY c.business_date, c.priority, c.gap_minutes DESC,
                    c.agent_id, c.gap_start""",
        [start, end, today],
    )
    timeline_headers, timeline_rows = _query(
        conn,
        """SELECT business_date, agent_id, agent_name, team_leader,
                  ops_manager, lob, language, scheduled_start, scheduled_end,
                  segment_start, segment_end, segment_minutes, planned_state,
                  actual_status, actual_category, mismatch_type, is_gap,
                  observed_source, source_file, evaluation_as_of
           FROM mart.shift_timeline_segment t
           WHERE business_date BETWEEN ? AND ? AND business_date<?
             AND EXISTS (
                 SELECT 1 FROM mart.correction_candidate c
                 WHERE c.business_date=t.business_date AND c.agent_id=t.agent_id
                   AND c.gap_start IS NOT NULL AND c.gap_end IS NOT NULL
             )
           ORDER BY business_date, agent_id, segment_start""",
        [start, end, today],
    )
    choices = list(dict.fromkeys(
        rule.patterns[0] for rule in rulebook.activity_rules
        if rule.category not in {"OFF"}
    ))
    add_review_board(
        book.report, headers, rows,
        [dict(zip(timeline_headers, row)) for row in timeline_rows],
        start, completed_through, choices,
    )
    book.tables.append(ModelTable("REVIEW BOARD", headers, rows))
    break_meal_headers, break_meal_rows = _break_meal_control_rows(
        conn, config, start, completed_through,
    )
    break_meal_sheet = book.table(
        "BREAK & MEAL", "Daily break and meal control",
        "Completed shifts only. Totals come from Agent Status inside the scheduled shift; incomplete coverage is never reported as zero compliance.",
        break_meal_headers, break_meal_rows,
    )
    if break_meal_rows:
        alert_column = break_meal_headers.index("alert")
        break_meal_sheet.conditional_format(
            4, alert_column, 3 + len(break_meal_rows), alert_column,
            {
                "type": "text", "criteria": "containing", "value": "EXCEEDED",
                "format": book.report.error,
            },
        )
        incomplete_format = book.report.workbook.add_format({
            "font_name": "Aptos", "font_size": 10, "bold": True,
            "font_color": COLORS["amber"], "bg_color": COLORS["amber_light"],
        })
        break_meal_sheet.conditional_format(
            4, alert_column, 3 + len(break_meal_rows), alert_column,
            {
                "type": "text", "criteria": "containing",
                "value": "INSUFFICIENT EVIDENCE", "format": incomplete_format,
            },
        )
    ledger_headers, ledger_rows = _query(
        conn,
        """SELECT c.correction_id AS gap_id, c.business_date, c.agent_id,
                  c.agent_name, c.detected_issue, c.gap_start AS exact_start,
                  c.gap_end AS exact_end, c.gap_minutes AS minutes,
                  coalesce(a.validation_status, 'Open') AS decision_status,
                  a.confirmed_activity AS decision_category,
                  a.owner AS reviewed_by, a.comment,
                  a.injected_date AS reviewed_date, a.updated_at,
                  a.imported_from
           FROM mart.correction_candidate c
           LEFT JOIN core.correction_action a
             ON a.correction_id=c.correction_id
           WHERE c.business_date BETWEEN ? AND ? AND c.business_date<?
             AND c.gap_start IS NOT NULL AND c.gap_end IS NOT NULL
           ORDER BY c.business_date, c.agent_id, c.gap_start""",
        [start, end, today],
    )
    ledger_sheet = book.table(
        "DECISION LEDGER", "Internal decision ledger snapshot",
        "Read-only decisions already stored in WFM Hub. Make new edits only in REVIEW BOARD, then import that same workbook.",
        ledger_headers, ledger_rows,
    )
    ledger_sheet.hide()
    evidence_sheet = book.table(
        "EVIDENCE", "Exact attendance evidence",
        "Exact schedule and observed segments behind every colored REVIEW BOARD cell. The 15-minute visual never changes these boundaries.",
        timeline_headers, timeline_rows,
    )
    evidence_sheet.hide()
    book.definitions([
        ("Exact gap", "One continuous scheduled interval without accepted working evidence", "Human review unit", "Never rounded or merged across a return"),
        ("Gap ID", "Stable date/agent/start/end/issue key", "Safe import key", "Edited timestamps are never trusted on import"),
        ("Approved", "The reviewer accepts the gap and assigns a rulebook category", "Calculates absence/shrinkage flags", "Category is mandatory"),
        ("Dismissed", "The detected gap should not count as loss", "Removes it from absence/shrinkage", "Comment recommended"),
        ("Open", "No final human decision exists", "Unverified minutes", "Never silently treated as absence or zero"),
        ("Logged", "Observed working, auxiliary, break or lunch evidence according to its separate color", "Visual shift context", "Exact raw state remains in EVIDENCE"),
        ("Schedule band", "Scheduled work and PTO/Away drawn directly above ACTUAL", "Plan-versus-actual comparison", "It is visual context and is never imported"),
        ("Gap", "Scheduled interval without accepted working evidence", "Review candidate", "The exact start/end at the left remain authoritative"),
        ("Current-day tail", "Unfinished part of today's shift", "No review row", "Never classified as early leave"),
        ("Break and meal control", "Daily Agent Status totals compared with configured allowances", "Operational overrun alert", "Completed shifts with insufficient coverage remain unknown"),
    ])
    lookup = book.report.workbook.add_worksheet("_LOOKUPS")
    lookup.write_row(0, 0, ["Decision Status", "Meaning"])
    lookup.write_row(1, 0, ["Open", "Awaiting review"])
    lookup.write_row(2, 0, ["Approved", "Use selected category"])
    lookup.write_row(3, 0, ["Dismissed", "Do not count as loss"])
    lookup.hide()
    book.audit(_audit_rows(
        conn, config, "corrections", start, end,
        (
            ("Completed-date cutoff", completed_through, "Today is excluded"),
            ("Decision authority", "core.correction_action", "Persistent human ledger"),
            ("Classification authority", rulebook.sha256, rulebook.version),
        ),
    ))
    return _finish(book, partial, target)
