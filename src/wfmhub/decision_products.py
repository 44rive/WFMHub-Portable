"""Focused WFM/Operations report products using one workbook design contract."""

from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any, Iterable, Sequence

from .config import Config
from .database import DatabaseConnection
from .metrics import MetricCatalog, evaluate_metric, load_metric_catalog
from .report_packs import publish_report, report_current_path
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
        return ("FINAL" if final else "LIVE"), "All required governed datasets are available"
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
        ("Report product", report_key, "One workbook = one operational decision"),
        ("Selected period", f"{start} to {end}", "Explicit report boundary"),
        ("Generated", datetime.now(), "Local work-machine time"),
        ("Refresh run", latest[0] if latest else None, latest[2] if latest else "No successful refresh metadata"),
        ("Calculation authority", "Python + SQLite", "Excel contains presentation only"),
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
    primary_answer = f"question_{primary}"
    allowed_scores = ", ".join(f"{value:g}" for value in config.pcs.allowed_scores)
    return _query(
        conn,
        f"""SELECT c.call_key AS coaching_key,
                   c.business_date, c.call_start,
                   c.call_reference_number, c.call_id,
                   c.agent_id,
                   coalesce(d.canonical_name,c.agent_name) AS agent_name,
                   d.team_leader, d.ops_manager,
                   coalesce(d.lob,c.lob) AS lob,
                   coalesce(d.language,c.language) AS language,
                   c.queue,
                   c.{primary_answer} AS q1_answer,
                   c.{primary_score} AS q1_score,
                   c.question_2 AS q2_answer,
                   c.question_3 AS customer_comment,
                   c.pcs_status, c.post_call_survey_mode,
                   c.source_file,
                   'Pending' AS coaching_status,
                   NULL AS coaching_date, NULL AS coach, NULL AS coaching_comment
            FROM core.clean_call_leg c
            LEFT JOIN core.dim_agent d ON d.agent_id=c.agent_id
            WHERE c.business_date BETWEEN ? AND ?
              AND upper(coalesce(c.call_direction,''))='I'
              AND c.{primary_score} IN ({allowed_scores})
              AND c.{primary_score} <= ?
            ORDER BY c.business_date DESC, d.team_leader,
                     coalesce(d.canonical_name,c.agent_name), c.call_start""",
        [start, end, config.pcs.negative_score_maximum],
    )


def _pcs_sum_formula(column: str) -> str:
    criteria = (
        'tblPcsData[Date],">="&PCS_From,'
        'tblPcsData[Date],"<="&PCS_To,'
        'tblPcsData[LOB],IF(DASHBOARD!$K$6="All","*",DASHBOARD!$K$6),'
        'tblPcsData[Team Leader],IF(DASHBOARD!$N$6="All","*",DASHBOARD!$N$6),'
        'tblPcsData[Agent],IF(DASHBOARD!$Q$6="All","*",DASHBOARD!$Q$6)'
    )
    return f"SUMIFS(tblPcsData[{column}],{criteria})"


def _pcs_completed_formula() -> str:
    criteria = (
        'tblCoaching[Date],">="&PCS_From,'
        'tblCoaching[Date],"<="&PCS_To,'
        'tblCoaching[LOB],IF(DASHBOARD!$K$6="All","*",DASHBOARD!$K$6),'
        'tblCoaching[Team Leader],IF(DASHBOARD!$N$6="All","*",DASHBOARD!$N$6),'
        'tblCoaching[Agent],IF(DASHBOARD!$Q$6="All","*",DASHBOARD!$Q$6),'
        'tblCoaching[Coaching Status],"Completed"'
    )
    return f"COUNTIFS({criteria})"


def _add_pcs_dashboard(
    book: DecisionWorkbook,
    status: str,
    status_text: str,
    start: date,
    end: date,
    comparison: Sequence[Sequence[Any]],
    lobs: Sequence[str],
    team_leaders: Sequence[str],
    agents: Sequence[str],
    trend_rows: Sequence[Sequence[Any]],
) -> None:
    """Create an Excel-native selector cockpit without Power Query or macros."""

    wb = book.report.workbook
    wb.set_calc_mode("auto")
    ws = wb.add_worksheet("DASHBOARD")
    ws.hide_gridlines(2)
    ws.set_tab_color(COLORS["gold"])
    ws.set_zoom(85)
    ws.set_landscape()
    ws.fit_to_pages(1, 1)
    ws.freeze_panes(4, 0)
    ws.merge_range("A1:R1", "PCS  /  PERFORMANCE & COACHING CONTROL", book.report.title)
    ws.merge_range(
        "A2:R2",
        f"Generated {book.generated:%Y-%m-%d %H:%M}  |  available selector window includes previous month through {end:%Y-%m-%d}",
        book.report.subtitle,
    )
    badge = status if status in book.badge_formats else "INCOMPLETE"
    ws.merge_range("A4:R4", f"{badge}  /  {status_text}", book.badge_formats[badge])

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
        ("PERIOD VIEW", "A5:C5", "A6:C7", "Selected period"),
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
    ws.data_validation("A6", {"validate": "list", "source": ["Selected period", "Latest day", "Current MTD", "Previous month"]})
    ws.data_validation("E6", {"validate": "date", "criteria": "between", "minimum": start, "maximum": end})
    ws.data_validation("H6", {"validate": "date", "criteria": "between", "minimum": start, "maximum": end})

    lists = ((23, ["All", *lobs]), (24, ["All", *team_leaders]), (25, ["All", *agents]))
    for column, values in lists:
        for row, value in enumerate(values):
            ws.write(row, column, value)
    ws.set_column(23, 25, None, None, {"hidden": True})
    ws.data_validation("K6", {"validate": "list", "source": f"=DASHBOARD!$X$1:$X${len(lobs) + 1}"})
    ws.data_validation("N6", {"validate": "list", "source": f"=DASHBOARD!$Y$1:$Y${len(team_leaders) + 1}"})
    ws.data_validation("Q6", {"validate": "list", "source": f"=DASHBOARD!$Z$1:$Z${len(agents) + 1}"})

    month_start = end.replace(day=1)
    previous_end = month_start - timedelta(days=1)
    previous_start = previous_end.replace(day=1)
    wb.define_name(
        "PCS_From",
        "=IF(DASHBOARD!$A$6=\"Latest day\",DATE(%d,%d,%d),"
        "IF(DASHBOARD!$A$6=\"Current MTD\",DATE(%d,%d,1),"
        "IF(DASHBOARD!$A$6=\"Previous month\",DATE(%d,%d,1),DASHBOARD!$E$6)))"
        % (end.year, end.month, end.day, end.year, end.month, previous_start.year, previous_start.month),
    )
    wb.define_name(
        "PCS_To",
        "=IF(DASHBOARD!$A$6=\"Latest day\",DATE(%d,%d,%d),"
        "IF(DASHBOARD!$A$6=\"Current MTD\",DATE(%d,%d,%d),"
        "IF(DASHBOARD!$A$6=\"Previous month\",DATE(%d,%d,%d),DASHBOARD!$H$6)))"
        % (end.year, end.month, end.day, end.year, end.month, end.day,
           previous_end.year, previous_end.month, previous_end.day),
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
        ("PCS AVERAGE", f"=IFERROR({score_sum}/{valid},0)", book.card_decimal, "Weighted score / valid responses"),
        ("PARTICIPATION", f"=IFERROR({participating}/{eligible},0)", book.card_percent, "Q1 nonblank / PCS Status 1"),
        ("VALID RESPONSES", f"={valid}", book.card_integer, "Configured valid Q1 scores"),
        ("INBOUND CALL LEGS", f"={inbound}", book.card_integer, "Inbound legs in selected scope"),
        ("SCORE <= 3", f"={low}", book.card_integer, "Coaching opportunities"),
        ("POSITIVE > 3", f"={positive}", book.card_integer, "Positive valid responses"),
        ("COACHING COMPLETED", f"={completed}", book.card_integer, "Updates as the team fills COACHING"),
        ("ACTIONS RATE", f"=IFERROR({completed}/{low},0)", book.card_percent, "Completed / score <= 3"),
    ]
    for index, (label, formula, fmt, note) in enumerate(cards):
        row = 9 if index < 4 else 14
        column = (index % 4) * 4
        ws.merge_range(row, column, row, column + 2, label, book.report.kpi_label)
        ws.merge_range(row + 1, column, row + 2, column + 2, "", fmt)
        ws.write_formula(row + 1, column, formula, fmt)
        ws.merge_range(row + 3, column, row + 3, column + 2, note, book.card_compare)
        ws.set_row(row + 1, 26)
        ws.set_row(row + 2, 26)

    table_row = 20
    ws.merge_range(table_row, 0, table_row, 9, "PERIOD BENCHMARK", book.report.section)
    compare_headers = ["Period", "Start", "End", "PCS Average", "Participation %", "Valid Responses", "PCS Status 1", "Score <= 3", "Score > 3", "Inbound Legs"]
    for column, header in enumerate(compare_headers):
        ws.write(table_row + 2, column, header, book.report.header)
    for offset, values in enumerate(comparison):
        for column, value in enumerate(values):
            fmt = book.report.date if isinstance(value, date) else book.report.percent if column == 4 else book.report.decimal if column == 3 else book.report.integer if isinstance(value, (int, float)) else book.report.body
            ws.write(table_row + 3 + offset, column, value, fmt)

    chart = wb.add_chart({"type": "line"})
    for name, column, color, secondary in (
        ("PCS Average", 5, COLORS["teal"], False),
        ("Participation", 6, COLORS["gold"], True),
    ):
        chart.add_series({
            "name": name,
            "categories": ["TREND", 4, 0, 3 + len(trend_rows), 0],
            "values": ["TREND", 4, column, 3 + len(trend_rows), column],
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

    note_row = table_row + 4 + max(len(comparison), 7)
    ws.merge_range(note_row, 0, note_row, 17, "HOW TO USE THIS PAGE", book.report.section)
    ws.merge_range(note_row + 1, 0, note_row + 1, 17, "Choose a period, LOB, team leader, or agent in the boxes above. Custom dates apply only when Period View is Selected period.", book.report.note)
    ws.merge_range(note_row + 2, 0, note_row + 2, 17, "For native slicers: click inside the PCS_DATA or COACHING Excel Table, then Table Design > Insert Slicer. No refresh, connection, or Data Model is required.", book.report.note)
    ws.set_column("A:R", 11)
    ws.set_column("A:A", 13)
    ws.set_column("K:R", 13)


def build_pcs_performance_workbook(
    conn: DatabaseConnection,
    config: Config,
    start: date,
    end: date,
    output: Path | None = None,
) -> Path:
    """Build one polished, self-contained PCS workbook with Excel selectors."""

    book, partial, target = _atomic_book(config, "pcs", "PCS PERFORMANCE", start, end, output)
    periods = _comparison_periods(start, end)
    comparison = [_pcs_aggregate(conn, config, period) for period in periods]
    status, status_text = _source_state(conn, ("fte", "calls"), end)
    previous_start, _previous_end = _previous_month(end)
    data_start = min(start, previous_start)
    headers, rows = _query(
        conn,
        """SELECT p.business_date, sum(p.pcs_score_sum) AS pcs_score_sum,
                  sum(survey_responses) AS valid_responses,
                  sum(pcs_participation_responses) AS participating_responses,
                  sum(pcs_status_calls) AS eligible_calls,
                  CASE WHEN sum(survey_responses)>0 THEN sum(pcs_score_sum)*1.0/sum(survey_responses) END AS pcs_average,
                  CASE WHEN sum(pcs_status_calls)>0 THEN sum(pcs_participation_responses)*1.0/sum(pcs_status_calls) END AS participation_rate,
                  sum(low_score_responses) AS score_le_3,
                  sum(top_box_responses) AS score_gt_3
           FROM mart.agent_pcs_day p
           WHERE p.business_date BETWEEN ? AND ?
           GROUP BY p.business_date ORDER BY p.business_date""",
        [data_start, end],
    )
    trend_headers, trend_rows = headers, rows

    selector_rows = conn.execute(
        """SELECT DISTINCT coalesce(lob,''), coalesce(team_leader,''), coalesce(agent_name,'')
           FROM mart.agent_pcs_day WHERE business_date BETWEEN ? AND ?""",
        [data_start, end],
    ).fetchall()
    lobs = sorted({str(row[0]) for row in selector_rows if row[0]})
    team_leaders = sorted({str(row[1]) for row in selector_rows if row[1]})
    agents = sorted({str(row[2]) for row in selector_rows if row[2]})
    _add_pcs_dashboard(book, status, status_text, start, end, comparison, lobs, team_leaders, agents, trend_rows)

    headers, rows = _query(
        conn,
        """SELECT max(lob) AS lob, max(team_leader) AS team_leader,
                  count(DISTINCT agent_id) AS agents,
                  sum(inbound_calls) AS inbound_call_legs,
                  sum(pcs_status_calls) AS pcs_status_1,
                  sum(pcs_participation_responses) AS q1_nonblank,
                  sum(survey_responses) AS valid_q1,
                  sum(pcs_score_sum) AS q1_score_sum,
                  CASE WHEN sum(survey_responses)>0 THEN sum(pcs_score_sum)*1.0/sum(survey_responses) END AS pcs_average,
                  CASE WHEN sum(pcs_status_calls)>0 THEN sum(pcs_participation_responses)*1.0/sum(pcs_status_calls) END AS participation_rate,
                  sum(low_score_responses) AS score_le_3,
                  sum(top_box_responses) AS score_gt_3,
                  CASE WHEN sum(survey_responses)<20 THEN 'LOW_SAMPLE' ELSE 'OK' END AS sample_state
           FROM mart.agent_pcs_day WHERE business_date BETWEEN ? AND ?
           GROUP BY coalesce(lob,''), coalesce(team_leader,'')
           ORDER BY max(lob), max(team_leader)""",
        [start, end],
    )
    team_sheet = book.table("TEAM_VIEW", "PCS team performance", "Selected report period; filter this Excel Table or add native Table slicers.", headers, rows)
    if rows:
        state_col = headers.index("sample_state")
        team_sheet.conditional_format(4, state_col, 3 + len(rows), state_col, {"type": "text", "criteria": "containing", "value": "LOW_SAMPLE", "format": book.report.error})

    headers, rows = _query(
        conn,
        """SELECT agent_id, max(agent_name) AS agent_name, max(team_leader) AS team_leader,
                  max(ops_manager) AS ops_manager, max(lob) AS lob, max(language) AS language,
                  sum(inbound_calls) AS inbound_call_legs,
                  sum(pcs_status_calls) AS pcs_status_1,
                  sum(pcs_participation_responses) AS q1_nonblank,
                  sum(survey_responses) AS valid_q1,
                  sum(pcs_score_sum) AS q1_score_sum,
                  CASE WHEN sum(survey_responses)>0 THEN sum(pcs_score_sum)*1.0/sum(survey_responses) END AS pcs_average,
                  CASE WHEN sum(pcs_status_calls)>0 THEN sum(pcs_participation_responses)*1.0/sum(pcs_status_calls) END AS participation_rate,
                  sum(low_score_responses) AS score_le_3,
                  sum(top_box_responses) AS score_gt_3,
                  sum(pcs_invalid_responses) AS invalid_q1,
                  CASE WHEN sum(survey_responses)<20 THEN 'LOW_SAMPLE' ELSE 'OK' END AS sample_state
           FROM mart.agent_pcs_day WHERE business_date BETWEEN ? AND ?
           GROUP BY agent_id ORDER BY max(lob), max(team_leader), max(agent_name)""",
        [start, end],
    )
    ws = book.table("AGENT_VIEW", "PCS agent performance", "Selected report period; click the table to filter or insert native slicers.", headers, rows)
    if rows:
        state_col = headers.index("sample_state")
        ws.conditional_format(4, state_col, 3 + len(rows), state_col, {"type": "text", "criteria": "containing", "value": "LOW_SAMPLE", "format": book.report.error})
    book.table("TREND", "PCS daily trend", "Daily governed counters covering the selected and previous-month comparison window.", trend_headers, trend_rows)

    action_headers, actions = _pcs_coaching_rows(conn, config, data_start, end)
    action_rows = actions or [tuple(None for _ in action_headers)]
    action_sheet = book.table(
        "COACHING",
        "PCS coaching action plan",
        "Filter to your team and dates. The team fills only the four blue columns; no Hub import or refresh is required.",
        action_headers,
        action_rows,
        editable_headers={"Coaching Status", "Coaching Date", "Coach", "Coaching Comment"},
    )
    if action_rows:
        status_col = action_headers.index("coaching_status")
        action_sheet.data_validation(
            4, status_col, 3 + len(action_rows), status_col,
            {
                "validate": "list",
                "source": ["Pending", "Completed", "Not required"],
                "input_title": "Coaching status",
                "input_message": "Choose one of the three governed statuses.",
                "error_title": "Invalid status",
                "error_message": "Use Pending, Completed or Not required.",
            },
        )
        action_sheet.conditional_format(
            4, status_col, 3 + len(action_rows), status_col,
            {"type": "text", "criteria": "containing", "value": "Pending", "format": book.report.error},
        )
    data_headers, data_rows = _query(
        conn,
        """SELECT business_date, agent_id, agent_name, team_leader, ops_manager,
                  lob, language, inbound_calls AS inbound_call_legs,
                  pcs_status_calls AS pcs_status_1,
                  pcs_participation_responses AS q1_nonblank,
                  survey_responses AS valid_q1,
                  pcs_score_sum AS q1_score_sum,
                  pcs_average, pcs_participation_rate AS participation_rate,
                  low_score_responses AS score_le_3,
                  top_box_responses AS score_gt_3,
                  pcs_invalid_responses AS invalid_q1,
                  CASE WHEN survey_responses<20 THEN 'LOW_SAMPLE' ELSE 'OK' END AS sample_state
           FROM mart.agent_pcs_day WHERE business_date BETWEEN ? AND ?
           ORDER BY business_date, lob, team_leader, agent_name""",
        [data_start, end],
    )
    book.table("PCS_DATA", "PCS clean calculation table", "One agent/day. Use the table filters or Table Design > Insert Slicer; formulas on DASHBOARD read this table directly.", data_headers, data_rows or [tuple(None for _ in data_headers)])
    book.definitions([
        ("PCS Average", "Sum of valid inbound Q1 scores / valid inbound Q1 responses", "Customer experience result", "Only configured discrete Q1 scores are valid"),
        ("PCS Participation", "Inbound raw Q1 nonblank / inbound PCSStatus=1", "Survey participation opportunity", "Invalid nonblank Q1 remains in the numerator"),
        ("Score <= 3", "Count of valid Q1 responses at or below 3", "Follow-up volume", "A count, not a percentage"),
        ("Actions Rate", "Completed rows in COACHING / valid Q1 responses at or below 3", "Coaching completion", "Team entry stays in the workbook; WFMHub does not import it"),
        ("Low sample", "Fewer than 20 valid responses in the selected period", "Interpretation warning", "Does not change the calculated score"),
        ("Selector mechanics", "Excel SUMIFS and COUNTIFS over the visible PCS_DATA and COACHING tables", "Interactive management view", "No Power Query, connection, macro, or Data Model"),
        ("Native slicers", "Click an Excel Table, then Table Design > Insert Slicer", "Optional manual filtering", "Slicers filter the selected table; dashboard selector boxes drive KPI cards"),
    ])
    book.audit(_audit_rows(conn, config, "pcs", start, end))
    return _finish(book, partial, target)


def _service_rows(
    conn: DatabaseConnection,
    profile: ServiceProfile,
    start: date,
    end: date,
) -> list[dict[str, Any]]:
    scope_marks = ",".join("?" for _ in profile.service_scopes)
    system_marks = ",".join("?" for _ in profile.source_systems)
    cursor = conn.execute(
        f"""SELECT business_date, interval_start, hour_start, source_system, queue,
                   service_scope, comparison_scope, designation, mapping_status,
                   offered, answered, abandoned, short_abandoned,
                   answered_within_target, handled_seconds, source_file
            FROM mart.service_interval
            WHERE business_date BETWEEN ? AND ?
              AND service_scope IN ({scope_marks})
              AND source_system IN ({system_marks})
            ORDER BY business_date, interval_start, source_system, queue""",
        [start, end, *profile.service_scopes, *profile.source_systems],
    )
    headers = [item[0] for item in cursor.description]
    return [dict(zip(headers, row)) for row in cursor.fetchall()]


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
    within = sum(float(row.get("answered_within_target") or 0) for row in values)
    handled = sum(float(row.get("handled_seconds") or 0) for row in values)
    components = {
        "offered": offered,
        "answered": answered,
        "abandoned": abandoned,
        "short_abandoned": short,
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
        "offered": offered,
        "answered": answered,
        "short_abandoned": short,
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
    """Build LOB service performance, beginning with the Ford OEM profile."""

    catalog = load_service_profiles(config.home, config.service_profiles)
    profile = catalog.select(profile_id, end)
    metric_catalog = load_metric_catalog(config.home, config.metric_catalog)
    rulebook = load_rulebook(config.home, config.business_rules)
    book, partial, target = _atomic_book(config, "service", f"OEM FLASH  /  {profile.label.upper()}", start, end, output)
    periods = _comparison_periods(start, end)
    all_start, all_end = min(p.start for p in periods), max(p.end for p in periods)
    raw_rows = _service_rows(conn, profile, all_start, all_end)
    period_rows = []
    for period in periods:
        scoped = [row for row in raw_rows if period.start.isoformat() <= str(row["business_date"])[:10] <= period.end.isoformat()]
        metrics = _service_aggregate(scoped, profile, metric_catalog, period.end)
        period_rows.append((period.label, period.start, period.end, metrics["offered"], metrics["answered"], metrics["service_level"], metrics["availability"], metrics["aht_seconds"]))
    by_label = {row[0]: row for row in period_rows}
    latest, mtd, prior = by_label["Latest day"], by_label["Current MTD"], by_label["Prior-month same days"]
    latest_method = _profile_metric(metric_catalog, profile, profile.service_level_metric, end)
    forecast_marks = ",".join("?" for _ in profile.service_scopes)
    forecast_total, forecast_rows_count = conn.execute(
        f"""SELECT sum(volume_forecast), count(*) FROM mart.forecast_hour
            WHERE business_date=? AND service_scope IN ({forecast_marks})""",
        [end, *profile.service_scopes],
    ).fetchone()
    if not forecast_rows_count:
        forecast_total = None
    forecast_method = metric_catalog.method_for(
        "forecast_attainment", end, {"lob": profile.service_scopes[0]},
    )
    forecast_evaluation = evaluate_metric(
        forecast_method, {"actual_volume": latest[3], "forecast_volume": forecast_total},
    ) if forecast_method and forecast_total is not None else None
    families = tuple(system.lower() for system in profile.source_systems) + ("forecast",)
    status, status_text = _source_state(conn, families, end)
    latest_raw = [
        row for row in raw_rows
        if str(row["business_date"])[:10] == end.isoformat()
    ]
    group_metrics: dict[str, dict[str, float | str | None]] = {}
    for group in profile.groups:
        group_metrics[group.label] = _service_aggregate(
            [row for row in latest_raw if profile.group_for(row.get("queue")) == group.label],
            profile,
            metric_catalog,
            end,
        )

    def group_value(label: str, metric: str) -> float | str | None:
        return group_metrics.get(label, {}).get(metric)

    variance_calls = (
        float(latest[3] or 0) - float(forecast_total)
        if forecast_total is not None else None
    )
    book.dashboard(
        [
            KpiCard("OEM availability", latest[6], "percent", "Answered / offered"),
            KpiCard("OEM service level", latest[5], "percent", f"Target {latest_method.target:.0%}" if latest_method.target is not None else "No target"),
            KpiCard("Ford availability", group_value("Ford", "availability"), "percent", "Mapped Ford queues"),
            KpiCard("Ford service level", group_value("Ford", "service_level"), "percent", "Mapped Ford queues"),
            KpiCard("Toyota availability", group_value("Toyota", "availability"), "percent", "Mapped Toyota/Lexus queues"),
            KpiCard("Toyota service level", group_value("Toyota", "service_level"), "percent", "Mapped Toyota/Lexus queues"),
            KpiCard("Forecast attainment", forecast_evaluation.value if forecast_evaluation else None, "percent", "Actual offered / forecast"),
            KpiCard("Variance calls", variance_calls, "integer", "Actual offered - forecast"),
        ],
        status,
        status_text,
        ["Period", "Start", "End", "Offered", "Answered", "Service Level %", "Availability %", "AHT Seconds"],
        period_rows,
        [
            "This report includes only queues mapped into the selected service profile; queue membership is controlled in queue_mapping.csv.",
            f"{profile.label}: the extract's within-{rulebook.target_seconds}s counter is evaluated by {profile.service_level_metric}.{latest_method.method_id} from metric_catalog.toml.",
            "Service availability is answered / offered. It is never agent availability or adherence.",
            "Forecast contributes forecast only. APBE/APFR/APDE contribute actual performance only.",
            "Back-office workload is shown as not configured until a governed source is supplied; WFMHub does not invent those counters.",
        ],
        (("Service Level", 5), ("Availability", 6)),
        sheet_name="FLASH",
    )

    hourly: dict[tuple[str, str, str], list[dict[str, Any]]] = defaultdict(list)
    for row in raw_rows:
        if start.isoformat() <= str(row["business_date"])[:10] <= end.isoformat():
            group = profile.group_for(row.get("queue"))
            hourly[(str(row["business_date"])[:10], str(row["hour_start"]), group)].append(row)
    forecast_rows = conn.execute(
        f"""SELECT hour_start, sum(volume_forecast)
            FROM mart.forecast_hour
            WHERE business_date BETWEEN ? AND ?
              AND service_scope IN ({forecast_marks})
            GROUP BY hour_start""",
        [start, end, *profile.service_scopes],
    ).fetchall()
    forecast_by_hour = {str(row[0])[:13]: float(row[1] or 0) for row in forecast_rows}
    staffing_marks = ",".join("?" for _ in profile.service_scopes)
    staffing_rows = conn.execute(
        f"""SELECT business_date, hour_key,
                   avg(scheduled_fte), avg(observed_fte), avg(productive_fte)
            FROM (
                SELECT business_date, substr(cast(interval_start AS TEXT),1,13) AS hour_key,
                       interval_start, sum(scheduled_fte) AS scheduled_fte,
                       sum(observed_fte) AS observed_fte,
                       sum(productive_fte) AS productive_fte
                FROM mart.staffing_interval
                WHERE business_date BETWEEN ? AND ?
                  AND lob IN ({staffing_marks})
                GROUP BY business_date, interval_start
            ) staffing
            GROUP BY business_date, hour_key""",
        [start, end, *profile.service_scopes],
    ).fetchall()
    staffing_by_hour = {
        f"{str(row[0])[:10]}T{str(row[1])[-2:]}": (
            float(row[2] or 0), float(row[3] or 0), float(row[4] or 0),
        )
        for row in staffing_rows
    }
    grouped_by_hour: dict[tuple[str, str], list[dict[str, Any]]] = defaultdict(list)
    for row in raw_rows:
        if start.isoformat() <= str(row["business_date"])[:10] <= end.isoformat():
            grouped_by_hour[(str(row["business_date"])[:10], str(row["hour_start"]))].append(row)

    hourly_rows = []
    for (business_date, hour_start), rows in sorted(grouped_by_hour.items()):
        row_date = date.fromisoformat(business_date)
        total = _service_aggregate(rows, profile, metric_catalog, row_date)
        hour_key = str(hour_start)[:13]
        forecast = forecast_by_hour.get(hour_key)
        attainment = _ratio(total["offered"], forecast)
        scheduled, observed, productive = staffing_by_hour.get(
            f"{business_date}T{hour_key[-2:]}", (None, None, None),
        )
        hourly_rows.append((
            business_date, hour_start, "OEM TOTAL", "OEM Total",
            total["offered"], forecast,
            float(total["offered"] or 0) - forecast if forecast is not None else None,
            attainment, total["answered"], total["within_target"],
            total["short_abandoned"], total["service_level"], total["availability"],
            total["aht_seconds"], scheduled, observed, productive,
            total["service_method"], total["service_state"],
        ))
    for (business_date, hour_start, group), rows in sorted(hourly.items()):
        row_date = date.fromisoformat(business_date)
        metrics = _service_aggregate(rows, profile, metric_catalog, row_date)
        hourly_rows.append((
            business_date, hour_start, "QUEUE GROUP", group,
            metrics["offered"], None, None, None, metrics["answered"],
            metrics["within_target"], metrics["short_abandoned"],
            metrics["service_level"], metrics["availability"], metrics["aht_seconds"],
            None, None, None, metrics["service_method"], metrics["service_state"],
        ))
    hourly_rows.sort(key=lambda row: (str(row[0]), str(row[1]), 0 if row[2] == "OEM TOTAL" else 1, str(row[3])))
    headers = [
        "business_date", "hour_start", "row_type", "service_group", "offered",
        "forecast", "variance_calls", "forecast_attainment", "answered",
        "answered_within_target", "short_abandoned", "service_level",
        "service_availability", "aht_seconds", "scheduled_fte", "observed_fte",
        "productive_fte", "service_method", "service_state",
    ]
    ws = book.table("HOURLY", f"{profile.label} hourly control", "OEM total rows reconcile forecast and staffing; Ford / Toyota / Chery rows show mapped queue-group performance.", headers, hourly_rows)
    if hourly_rows:
        state_col = headers.index("service_state")
        ws.conditional_format(4, state_col, 3 + len(hourly_rows), state_col, {"type": "text", "criteria": "containing", "value": "BELOW_TARGET", "format": book.report.error})
    detail_headers = ["business_date", "interval_start", "source_system", "queue", "service_scope", "designation", "mapping_status", "offered", "answered", "abandoned", "short_abandoned", "answered_within_target", "handled_seconds", "source_file"]
    detail_rows = [tuple(row.get(header) for header in detail_headers) for row in raw_rows if start.isoformat() <= str(row["business_date"])[:10] <= end.isoformat()]
    book.table("QUEUES", "Mapped queue evidence", "Compact governed queue intervals for reconciliation; no original extract rows are modified.", detail_headers, detail_rows)
    action_rows = [row for row in hourly_rows if row[-1] == "BELOW_TARGET"]
    book.table("EXCEPTIONS", "Service intervals below target", "Prioritize intervals with high offered volume and validate staffing/forecast before escalation.", headers, action_rows)
    source_headers, source_rows = _query(
        conn,
        """SELECT source_family, newest_business_date, row_count, scoped_out_count,
                  status, newest_file, details
           FROM mart.source_health
           WHERE lower(source_family) IN ('apbe','apfr','apde','forecast','fte','schedule','lilo','agent_status')
           ORDER BY source_family""",
    )
    source_rows.append((
        "back_office", None, None, None, "NOT_CONFIGURED", None,
        "The reference workbook contains manually sourced backlog counters; no governed source is configured yet.",
    ))
    book.table("DATA_STATUS", "Flash source status", "The Flash stays explicit about missing or stale inputs.", source_headers, source_rows)
    book.definitions([
        ("Service Level", f"({latest_method.numerator}) / ({latest_method.denominator})", "Contract performance", f"Governed method {latest_method.method_id}; ratio of summed counters"),
        ("Service Availability", "Answered / offered", "Ability of the service to answer demand", "Not an agent metric"),
        ("Forecast Attainment", "Actual offered / forecast", "Demand realization", "Forecast and actual joined only by mapped comparison scope"),
        ("AHT", "Total handled seconds / answered contacts", "Workload driver", "Weighted, never average of interval AHT"),
    ])
    book.audit(_audit_rows(conn, config, "service", start, end, [
        ("Service profile", profile.profile_id, f"{catalog.version} / {catalog.sha256}"),
        ("Included scopes", ", ".join(profile.service_scopes), ", ".join(profile.source_systems)),
        ("Service metric", profile.service_level_metric, f"{metric_catalog.version} / {metric_catalog.sha256}"),
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
    """Build the standalone LOB/language staffing decision product."""

    report_day = end
    book, partial, target = _atomic_book(config, "staffing", "STAFFING & COVERAGE", report_day, report_day, output)
    totals = conn.execute(
        """SELECT max(staffing_gap_fte),
                  coalesce(sum(CASE WHEN staffing_state IN ('GAP','PARTIAL_GAP') THEN 1 ELSE 0 END),0),
                  coalesce(sum(CASE WHEN staffing_state='DATA_MISSING' THEN 1 ELSE 0 END),0),
                  avg(scheduled_fte), avg(observed_fte), avg(productive_fte)
           FROM mart.staffing_interval WHERE business_date=?""",
        [report_day],
    ).fetchone()
    peak_gap, gap_intervals, missing_intervals, avg_scheduled, avg_observed, avg_productive = totals
    status, status_text = _source_state(conn, ("fte", "start_end", "lilo", "agent_status"), report_day)
    if missing_intervals:
        status, status_text = "INCOMPLETE", f"{missing_intervals:,} interval(s) have missing attendance evidence"
    book.dashboard(
        [
            KpiCard("Peak gap FTE", peak_gap, "decimal", "Largest 15-minute deficit"),
            KpiCard("Gap intervals", gap_intervals, "integer", "LOB/language intervals"),
            KpiCard("Average scheduled FTE", avg_scheduled, "decimal"),
            KpiCard("Average observed FTE", avg_observed, "decimal"),
            KpiCard("Average productive FTE", avg_productive, "decimal"),
            KpiCard("Missing intervals", missing_intervals, "integer", "Unknown, not zero"),
        ],
        status,
        status_text,
        ["Measure", "Value"],
        [("Peak gap FTE", peak_gap), ("Gap intervals", gap_intervals), ("Average scheduled FTE", avg_scheduled), ("Average observed FTE", avg_observed), ("Average productive FTE", avg_productive)],
        [
            "Staffing is grouped by roster LOB and language; it is not mixed with service or attendance calling.",
            "Observed FTE is calculated from agent-seconds inside each 15-minute interval.",
            "Future and missing-evidence intervals stay unknown rather than becoming false staffing deficits.",
        ],
    )
    headers, rows = _query(
        conn,
        """SELECT business_date, interval_start, interval_end, lob, language,
                  scheduled_agents, observed_agents, productive_agents,
                  scheduled_fte, elapsed_scheduled_fte, observed_fte,
                  productive_fte, staffing_variance_fte, staffing_gap_fte,
                  staffing_state, evidence_basis, evaluation_as_of
           FROM mart.staffing_interval WHERE business_date=?
           ORDER BY interval_start, lob, language""",
        [report_day],
    )
    book.table("INTRADAY", "15-minute staffing coverage", "Complete daily evidence at LOB/language grain.", headers, rows)
    actions = [row for row in rows if row[headers.index("staffing_state")] in {"GAP", "PARTIAL_GAP", "DATA_MISSING"}]
    book.table("ACTIONS", "Staffing exceptions", "Intervals requiring redeployment, investigation or evidence repair.", headers, actions)
    book.definitions([
        ("Scheduled FTE", "Scheduled agent-seconds / 900", "Expected interval capacity", "Roster LOB/language"),
        ("Observed FTE", "Observed agent-seconds / 900", "Actual presence", "LILO + Agent Status evidence"),
        ("Productive FTE", "Productive-status seconds / 900", "Available handling capacity", "Not adherence"),
        ("Staffing gap", "MAX(0, elapsed scheduled FTE - observed FTE)", "Immediate staffing deficit", "Blank for future/missing evidence"),
    ])
    book.audit(_audit_rows(conn, config, "staffing", report_day, report_day))
    return _finish(book, partial, target)


def build_final_absence_product_workbook(
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


def build_attendance_corrections_workbook(
    conn: DatabaseConnection,
    config: Config,
    start: date,
    end: date,
    output: Path | None = None,
) -> Path:
    """Build the selected period's completed-day correction queue and timeline."""

    from .governed_workbooks import _add_shift_view

    today = date.today()
    completed_through = min(end, today - timedelta(days=1))
    book, partial, target = _atomic_book(
        config, "corrections", "ATTENDANCE REVIEW", start, end, output,
    )
    gap_count, gap_minutes, agents = conn.execute(
        """SELECT count(*), coalesce(sum(residual_minutes),0), count(DISTINCT agent_id)
           FROM mart.correction_residual_segment
           WHERE business_date BETWEEN ? AND ? AND business_date<?""",
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
        conn, ("fte", "start_end", "lilo", "agent_status", "activities"), completed_through, final=True,
    )
    if missing:
        status, status_text = "INCOMPLETE", f"{missing:,} scheduled row(s) lack complete observed evidence"
    book.dashboard(
        [
            KpiCard("Residual segments", gap_count, "integer"),
            KpiCard("Residual gap hours", gap_minutes / 60 if gap_minutes else 0, "decimal"),
            KpiCard("Agents to review", agents, "integer"),
            KpiCard("Missing evidence", missing, "integer"),
        ],
        status,
        status_text,
        ["Measure", "Value"],
        [("Residual segments", gap_count), ("Residual gap hours", gap_minutes / 60 if gap_minutes else 0), ("Agents to review", agents), ("Missing evidence", missing)],
        [
            f"The selected range is {start:%Y-%m-%d} to {end:%Y-%m-%d}; every completed date through {completed_through:%Y-%m-%d} is included, not only yesterday.",
            "Today is excluded from correction review, so an unfinished shift can never become an early-leave correction.",
            "GAPS is the editable action list. Pale-blue columns are the only human decisions accepted for import.",
            "SHIFT_VIEW visualizes the full planned-versus-observed shift; every selected completed date remains in the workbook evidence.",
            "Verint Activities verify whether an observed gap is corrected; they never create the original gap.",
        ],
    )
    headers, rows = _query(
        conn,
        """SELECT r.residual_id, r.correction_id, r.business_date,
                  r.agent_id, c.agent_name, c.team_leader, c.ops_manager,
                  c.lob, d.language, c.scheduled_start, c.scheduled_end,
                  c.detected_issue, r.residual_start AS gap_start,
                  r.residual_end AS gap_end, r.residual_minutes AS gap_minutes,
                  c.confidence, r.suggested_activity, r.observed_source,
                  r.verint_reconciliation, c.verint_activity,
                  c.verint_overlap_minutes, c.validation_status,
                  c.confirmed_activity, c.owner, c.comment, c.injected_date,
                  r.source_file
           FROM mart.correction_residual_segment r
           JOIN mart.correction_candidate c ON c.correction_id=r.correction_id
           LEFT JOIN core.dim_agent d ON d.agent_id=r.agent_id
           WHERE r.business_date BETWEEN ? AND ? AND r.business_date<?
             AND coalesce(c.validation_status,'Open') NOT IN ('Injected','Rejected')
           ORDER BY r.business_date, c.priority, r.residual_minutes DESC, r.agent_id, r.residual_start""",
        [start, end, today],
    )
    ws = book.table(
        "GAPS", "Residual gaps ready for Verint correction",
        "Edit only Confirmed Activity, Validation Status, Owner, Comment and Injected Date; then import this workbook.",
        headers, rows,
        editable_headers={"Confirmed Activity", "Validation Status", "Owner", "Comment", "Injected Date"},
    )
    if rows:
        display = [value.replace("_", " ").title().replace("Id", "ID") for value in headers]
        status_col = display.index("Validation Status")
        ws.data_validation(4, status_col, 3 + len(rows), status_col, {
            "validate": "list", "source": ["Open", "Validated", "Injected", "Rejected"],
        })
        suggestions = sorted({str(row[headers.index("suggested_activity")]) for row in rows if row[headers.index("suggested_activity")]})
        activities = list(dict.fromkeys([*suggestions, "Absent", "Late", "Early Leave", "Unpaid Leave", "Vacation"]))[:20]
        if activities:
            activity_col = display.index("Confirmed Activity")
            ws.data_validation(4, activity_col, 3 + len(rows), activity_col, {"validate": "list", "source": activities})

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
                 SELECT 1 FROM mart.correction_residual_segment r
                 WHERE r.business_date=t.business_date AND r.agent_id=t.agent_id
             )
           ORDER BY business_date, agent_id, segment_start""",
        [start, end, today],
    )
    timeline_dicts = [dict(zip(timeline_headers, row)) for row in timeline_rows]
    _add_shift_view(book.report, timeline_dicts, start, completed_through)
    book.tables.append(ModelTable("TIMELINE", timeline_headers, timeline_rows))
    book.definitions([
        ("Observed gap", "Scheduled time minus unioned LILO/Agent Status evidence", "Correction candidate", "Activities cannot create the gap"),
        ("Residual gap", "Observed gap minus unioned corrected Verint Activities", "Minutes still requiring review", "Overlap-safe"),
        ("Current-day tail", "Future portion of an unfinished shift", "No action", "Never early leave"),
        ("Correction ID", "Stable action key", "Import decisions", "One correction may have several residual segments"),
    ])
    book.audit(_audit_rows(
        conn, config, "corrections", start, end,
        (("Completed-date cutoff", completed_through, "Today is excluded from correction review"),),
    ))
    return _finish(book, partial, target)
