"""Focused Excel workbooks rendered from governed datasets and metric values."""

from __future__ import annotations

from collections import defaultdict
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any, Iterable

from .config import Config
from .database import DatabaseConnection
from .datasets import (
    final_absence_lob_month,
    pcs_agent_month,
    pcs_team_day,
    service_scope_interval,
)
from .metrics import load_metric_catalog
from .report_packs import report_pack, report_pack_folder
from .reporting import (
    add_domain_rules_sheet,
    add_findings_sheet,
    add_methods_sheet,
    add_provenance_sheet,
    report_spec,
    validate_workbook_contract,
)
from .reports import COLORS, ExcelReport, _display_header, _query
from .rules import load_rulebook
from .semantic import aggregate_metric_values


def _output_path(
    config: Config,
    key: str,
    start: date,
    end: date,
    generated: datetime,
    output: Path | None,
) -> Path:
    pack = report_pack(key)
    return (
        output
        or report_pack_folder(config, key)
        / f"{pack.filename_prefix}_{start:%Y-%m-%d}_to_{end:%Y-%m-%d}_{generated:%H%M%S_%f}.xlsx"
    ).resolve()


def _add_landing(
    report: ExcelReport,
    sheet_name: str,
    title: str,
    subtitle: str,
    kpis: list[tuple[str, Any, str]],
    notes: Iterable[str],
    badge: str | None = None,
) -> None:
    ws = report.workbook.add_worksheet(sheet_name)
    ws.set_tab_color(COLORS["gold"])
    ws.hide_gridlines(2)
    ws.set_zoom(95)
    ws.merge_range("A1:L1", title, report.title)
    ws.merge_range("A2:L2", subtitle, report.subtitle)
    ws.set_row(0, 34)
    ws.set_row(1, 20)
    if badge:
        badge_fmt = report.workbook.add_format({
            "font_name": "Aptos", "font_size": 10, "bold": True,
            "font_color": COLORS["white"], "bg_color": COLORS["purple"],
            "align": "center", "valign": "vcenter",
        })
        ws.merge_range("A4:L4", badge, badge_fmt)
        ws.set_row(3, 22)
    start_row = 5
    for index, (label, value, kind) in enumerate(kpis[:4]):
        left = index * 3
        right = left + 2
        ws.merge_range(start_row, left, start_row, right, label.upper(), report.kpi_label)
        value_fmt = report.kpi_value
        if kind == "percent":
            value_fmt = report.workbook.add_format({
                "font_name": "Aptos Display", "font_size": 20, "bold": True,
                "font_color": COLORS["dark"], "bg_color": COLORS["white"],
                "align": "center", "valign": "vcenter", "num_format": "0.0%",
                "bottom": 1, "left": 1, "right": 1,
                "bottom_color": COLORS["thin"], "left_color": COLORS["thin"],
                "right_color": COLORS["thin"],
            })
        elif kind == "decimal":
            value_fmt = report.workbook.add_format({
                "font_name": "Aptos Display", "font_size": 20, "bold": True,
                "font_color": COLORS["dark"], "bg_color": COLORS["white"],
                "align": "center", "valign": "vcenter", "num_format": "#,##0.00",
                "bottom": 1, "left": 1, "right": 1,
                "bottom_color": COLORS["thin"], "left_color": COLORS["thin"],
                "right_color": COLORS["thin"],
            })
        ws.merge_range(start_row + 1, left, start_row + 3, right, value, value_fmt)
    note_row = start_row + 5
    ws.merge_range(note_row, 0, note_row, 11, "HOW TO USE THIS WORKBOOK", report.section)
    for offset, note in enumerate(notes, 1):
        ws.merge_range(note_row + offset, 0, note_row + offset, 11, note, report.note)
        ws.set_row(note_row + offset, 27)
    ws.set_column("A:L", 11)
    ws.freeze_panes(2, 0)


def _semantic_formats(report: ExcelReport) -> dict[str, Any]:
    def fmt(font: str, fill: str):
        return report.workbook.add_format({
            "font_name": "Aptos", "font_size": 10, "bold": True,
            "font_color": font, "bg_color": fill,
        })

    return {
        "good": fmt(COLORS["green"], COLORS["green_light"]),
        "warn": fmt(COLORS["amber"], COLORS["amber_light"]),
        "bad": fmt(COLORS["red"], COLORS["red_light"]),
        "info": fmt(COLORS["purple"], COLORS["purple_light"]),
        "future": fmt(COLORS["future"], COLORS["future_light"]),
    }


def _color_statuses(
    report: ExcelReport,
    ws,
    headers: list[str],
    rows: list[tuple[Any, ...]],
    column: str,
    states: dict[str, str],
) -> None:
    display = [_display_header(header) for header in headers]
    display_column = _display_header(column)
    if not rows or display_column not in display:
        return
    col = display.index(display_column)
    formats = _semantic_formats(report)
    for text, style in states.items():
        ws.conditional_format(4, col, 3 + len(rows), col, {
            "type": "text", "criteria": "containing", "value": text,
            "format": formats[style],
        })


def _quality_sheet(
    report: ExcelReport,
    conn: DatabaseConnection,
    start: date,
    end: date,
    families: tuple[str, ...],
) -> None:
    placeholders = ",".join("?" for _ in families)
    headers, rows = _query(
        conn,
        f"""SELECT detected_at, source_family, source_file, business_date,
                   agent_id, issue_type, severity, details
            FROM meta.quality_issue
            WHERE (business_date IS NULL OR business_date BETWEEN ? AND ?)
              AND source_family IN ({placeholders})
            ORDER BY CASE severity WHEN 'ERROR' THEN 1 WHEN 'REVIEW' THEN 2 ELSE 3 END,
                     business_date, issue_type""",
        [start, end, *families],
    )
    ws = report.add_table_sheet(
        "DATA_QUALITY", "Data quality gates",
        "Resolve ERROR rows before operational use; REVIEW rows need a human check.",
        headers, rows,
    )
    _color_statuses(report, ws, headers, rows, "severity", {
        "ERROR": "bad", "REVIEW": "warn", "INFO": "info",
    })


def _source_health_sheet(
    report: ExcelReport,
    conn: DatabaseConnection,
    families: tuple[str, ...],
) -> None:
    placeholders = ",".join("?" for _ in families)
    headers, rows = _query(
        conn,
        f"""SELECT source_family, expected_path, newest_file,
                   newest_business_date, modified_at, loaded_at, row_count,
                   rejected_count, scoped_out_count, status, details
            FROM mart.source_health
            WHERE source_family IN ({placeholders})
            ORDER BY source_family""",
        list(families),
    )
    ws = report.add_table_sheet(
        "SOURCE_HEALTH", "Source coverage",
        "Newest available source date and latest load result. Raw files remain untouched.",
        headers, rows,
    )
    _color_statuses(report, ws, headers, rows, "status", {
        "SUCCESS": "good", "ERROR": "bad", "MISSING": "bad", "REVIEW": "warn",
    })


def _schedule_variant_sheet(report: ExcelReport, conn: DatabaseConnection) -> None:
    headers, rows = _query(
        conn,
        """SELECT coalesce(source_variant,'UNKNOWN') AS schedule_variant,
                  count(*) AS active_files, max(file_name) AS newest_file,
                  max(modified_at) AS modified_at, max(loaded_at) AS loaded_at,
                  sum(row_count) AS rows_loaded,
                  sum(rejected_count) AS rows_rejected,
                  CASE WHEN sum(CASE WHEN status='SUCCESS' THEN 1 ELSE 0 END)>0
                       THEN 'SUCCESS' ELSE 'ERROR' END AS status
           FROM meta.source_file
           WHERE source_family='schedule' AND active=true
           GROUP BY coalesce(source_variant,'UNKNOWN')
           ORDER BY schedule_variant""",
    )
    ws = report.add_table_sheet(
        "SCHEDULE_SOURCES", "Verint schedule source roles",
        "START_END drives shifts. ACTIVITIES is the corrected final ledger; neither may silently replace the other.",
        headers, rows,
    )
    _color_statuses(report, ws, headers, rows, "status", {"SUCCESS": "good", "ERROR": "bad"})


def build_daily_operations_workbook(
    conn: DatabaseConnection,
    config: Config,
    start: date,
    end: date,
    output: Path | None = None,
) -> Path:
    """Build the single-day call list, staffing gaps and APDE SL workbook."""
    generated = datetime.now()
    report_day = end
    output = _output_path(config, "operations", report_day, report_day, generated, output)
    output.parent.mkdir(parents=True, exist_ok=True)
    partial = output.with_name(f"{output.stem}.partial{output.suffix}")
    report = ExcelReport(partial)
    rulebook = load_rulebook(config.home, config.business_rules)
    metric_catalog = load_metric_catalog(config.home, config.metric_catalog)
    spec = report_spec(config, "operations")
    try:
        service_dataset = service_scope_interval(
            conn, metric_catalog, report_day, "APDE",
            config.report_limits.get("max_service_rows", 100000),
        )
        call_now, no_shows = conn.execute(
            """SELECT count(*), coalesce(sum(CASE WHEN call_action='CALL_NO_SHOW' THEN 1 ELSE 0 END),0)
               FROM mart.attendance_agent_day
               WHERE business_date=? AND requires_call=true""",
            [report_day],
        ).fetchone()
        scheduled_rows, missing_evidence_rows = conn.execute(
            """SELECT count(*),
                      coalesce(sum(CASE WHEN source_loaded=false THEN 1 ELSE 0 END),0)
               FROM mart.attendance_agent_day
               WHERE business_date=? AND assignment_type NOT IN ('Off','Planned absence')""",
            [report_day],
        ).fetchone()
        coverage_badge = (
            f"INCOMPLETE DATA  /  {missing_evidence_rows:,} OF {scheduled_rows:,} WORKING SHIFTS LACK ATTENDANCE EVIDENCE"
            if not scheduled_rows or missing_evidence_rows
            else "ATTENDANCE EVIDENCE COMPLETE FOR ALL SCHEDULED WORKING SHIFTS"
        )
        peak_gap = conn.execute(
            """SELECT max(staffing_gap_fte) FROM mart.staffing_interval
               WHERE business_date=? AND staffing_gap_fte IS NOT NULL""",
            [report_day],
        ).fetchone()[0]
        state_index = service_dataset.headers.index("sl_state")
        below_target = sum(1 for row in service_dataset.rows if row[state_index] == "BELOW_TARGET")
        _add_landing(
            report, "DAILY_SUMMARY", spec.title,
            f"Operational day {report_day:%Y-%m-%d}  |  generated {generated:%Y-%m-%d %H:%M}",
            [
                ("Call now", call_now or 0, "integer"),
                ("Confirmed no-show", no_shows or 0, "integer"),
                ("Peak staffing gap FTE", peak_gap or 0, "decimal"),
                ("APDE intervals below SL", below_target or 0, "integer"),
            ],
            [
                "ATTENDANCE_CALLS is the calling queue. CALL_NOT_SEEN_NOW is provisional; CALL_NO_SHOW is final only after the shift has completed.",
                "STAFFING_GAPS uses roster LOB/language and 15-minute agent-seconds. DATA_MISSING is unknown—not a zero and not a staffing shortage.",
                "SERVICE_LEVEL contains APDE only. Service availability means answered / offered; service level is recalculated from summed counters.",
                "This is a presentation of governed marts. No extract is edited and no raw row is loaded into this workbook.",
            ],
            badge=coverage_badge,
        )
        add_findings_sheet(report, conn, spec, report_day, report_day)

        headers, rows = _query(
            conn,
            """SELECT business_date, agent_id, agent_name, team_leader,
                      ops_manager, lob, language, scheduled_start, scheduled_end,
                      shift_state, attendance_result, call_action, requires_call,
                      is_provisional, actual_first_seen, actual_last_seen,
                      actual_evidence, uncoded_late_minutes, no_show_minutes,
                      source_loaded, schedule_source, lilo_source, status_source,
                      evaluation_as_of
               FROM mart.attendance_agent_day
               WHERE business_date=? AND requires_call=true
               ORDER BY CASE call_action WHEN 'CALL_NO_SHOW' THEN 1
                                         WHEN 'CALL_LATE' THEN 2
                                         WHEN 'CALL_NOT_SEEN_NOW' THEN 3 ELSE 4 END,
                        scheduled_start, lob, language, agent_name
               LIMIT ?""",
            [report_day, config.report_limits.get("max_attendance_rows", 100000)],
        )
        ws = report.add_table_sheet(
            "ATTENDANCE_CALLS", "Daily absence and lateness call list",
            "Only governed rows requiring a call. Provisional current-day states are visibly separated from completed-shift outcomes.",
            headers, rows,
        )
        _color_statuses(report, ws, headers, rows, "call_action", {
            "CALL_NO_SHOW": "bad", "CALL_LATE": "warn", "CALL_NOT_SEEN_NOW": "info",
        })

        headers, rows = _query(
            conn,
            """SELECT business_date, interval_start, interval_end, lob, language,
                      scheduled_agents, observed_agents, productive_agents,
                      auxiliary_agents, scheduled_fte, elapsed_scheduled_fte,
                      observed_fte, productive_fte, staffing_variance_fte,
                      staffing_gap_fte, staffing_state, evidence_basis,
                      evaluation_as_of
               FROM mart.staffing_interval WHERE business_date=?
               ORDER BY interval_start, lob, language
               LIMIT ?""",
            [report_day, config.report_limits.get("max_staffing_rows", 100000)],
        )
        ws = report.add_table_sheet(
            "STAFFING_GAPS", "15-minute staffing gaps",
            "Agent-seconds converted to FTE. Future and missing-evidence intervals keep variance/gap blank by design.",
            headers, rows,
        )
        _color_statuses(report, ws, headers, rows, "staffing_state", {
            "DATA_MISSING": "bad", "DATA_PARTIAL": "warn", "PARTIAL_GAP": "warn",
            "GAP": "bad", "FUTURE": "future", "OK": "good",
        })
        display = [_display_header(item) for item in headers]
        if rows and "Staffing Gap FTE" in display:
            col = display.index("Staffing Gap FTE")
            ws.conditional_format(4, col, 3 + len(rows), col, {
                "type": "data_bar", "bar_color": COLORS["red"],
            })

        headers, rows = service_dataset.headers, service_dataset.rows
        ws = report.add_table_sheet(
            "SERVICE_LEVEL", "APDE service-level state by LOB",
            "Values come from effective-dated metric methods; additive counters remain visible for reconciliation.",
            headers, rows,
        )
        _color_statuses(report, ws, headers, rows, "sl_state", {
            "BELOW_TARGET": "bad", "ON_TARGET": "good",
            "NO_DATA": "future", "NO_TRAFFIC": "future", "LOW_SAMPLE": "warn",
        })
        _quality_sheet(report, conn, report_day, report_day, ("fte", "schedule", "lilo", "agent_status", "attendance", "apde", "service"))
        _source_health_sheet(report, conn, ("fte", "schedule", "lilo", "agent_status", "apde"))
        _schedule_variant_sheet(report, conn)
        add_domain_rules_sheet(report, config, spec)
        add_methods_sheet(report, config, spec)
        add_provenance_sheet(report, conn, config, spec, report_day, report_day)
        validate_workbook_contract(report, spec)
    except Exception:
        report.close()
        partial.unlink(missing_ok=True)
        raise
    else:
        report.close()
        partial.replace(output)
    return output


def build_exact_pcs_workbook(
    conn: DatabaseConnection,
    config: Config,
    start: date,
    end: date,
    output: Path | None = None,
) -> Path:
    """Build the exact Q1 PCS workbook reconciled to the reference O15:R logic."""
    generated = datetime.now()
    output = _output_path(config, "quality_pcs", start, end, generated, output)
    output.parent.mkdir(parents=True, exist_ok=True)
    partial = output.with_name(f"{output.stem}.partial{output.suffix}")
    report = ExcelReport(partial)
    allowed = ", ".join(f"{value:g}" for value in config.pcs.allowed_scores)
    metric_catalog = load_metric_catalog(config.home, config.metric_catalog)
    spec = report_spec(config, "quality_pcs")
    try:
        totals = conn.execute(
            """SELECT coalesce(sum(inbound_calls),0),
                      coalesce(sum(survey_responses),0),
                      coalesce(sum(pcs_score_sum),0),
                      coalesce(sum(low_score_responses),0),
                      coalesce(sum(top_box_responses),0),
                      coalesce(sum(pcs_participation_responses),0),
                      coalesce(sum(pcs_status_calls),0)
               FROM mart.agent_pcs_day WHERE business_date BETWEEN ? AND ?""",
            [start, end],
        ).fetchone()
        inbound, valid_count, score_sum, negative, positive, participants, denominator = totals
        semantic_summary = {
            item.metric_id: item
            for item in aggregate_metric_values(
                conn, metric_catalog, start, end,
                ["pcs_average", "pcs_participation"], (),
            )
        }
        pcs_average = semantic_summary.get("pcs_average").value if semantic_summary.get("pcs_average") else None
        participation = semantic_summary.get("pcs_participation").value if semantic_summary.get("pcs_participation") else None
        _add_landing(
            report, "PCS_SUMMARY", spec.title,
            f"Period {start:%Y-%m-%d} to {end:%Y-%m-%d}  |  generated {generated:%Y-%m-%d %H:%M}",
            [
                ("Valid Q1 average", pcs_average, "decimal"),
                ("Q1 score <= 3", negative, "integer"),
                ("Q1 score > 3", positive, "integer"),
                ("PCS participation", participation, "percent"),
            ],
            [
                f"Exact score: inbound Q1 is valid only when its parsed value is one of {{{allowed}}}; average = valid Q1 score sum / valid Q1 count.",
                f"Negative = valid Q1 <= {config.pcs.negative_score_maximum:g}; positive = valid Q1 > {config.pcs.negative_score_maximum:g}. Blank and invalid Q1 values are excluded from both.",
                f"Participation = inbound raw Q1 nonblank ({participants:,}) / inbound PCSStatus={config.pcs.participation_status} ({denominator:,}). It is not valid scores / Mode 2.",
                f"Inbound call legs in scope: {inbound:,}. PostCallSurveyMode={config.pcs.survey_mode} and Q2 are diagnostics only, never headline PCS inputs.",
            ],
            badge="REFERENCE-COMPATIBLE Q1 LOGIC  /  COUNTERS FIRST, RATIOS SECOND",
        )
        add_findings_sheet(report, conn, spec, start, end)

        agent_day_sql = """SELECT business_date, agent_id, agent_name, team_leader,
                   ops_manager, lob, market, language, location,
                   call_legs, handled_calls, inbound_calls AS inbound_call_legs,
                   outbound_calls AS outbound_call_legs, transferred_legs,
                   average_handle_seconds,
                   pcs_enabled_calls AS pcs_mode_2_inbound_legs,
                   pcs_status_calls AS pcs_status_1_inbound_legs,
                   pcs_participation_responses AS pcs_q1_nonblank_inbound_legs,
                   survey_responses AS pcs_q1_valid_score_count,
                   pcs_score_sum AS pcs_q1_score_sum,
                   low_score_responses AS pcs_score_le_3_count,
                   top_box_responses AS pcs_score_gt_3_count,
                   pcs_invalid_responses AS pcs_q1_invalid_nonblank_count,
                   pcs_status_blank_responses AS pcs_status_1_q1_blank_count,
                   pcs_response_without_status AS pcs_q1_nonblank_without_status_1_count,
                   pcs_average, pcs_participation_rate,
                   CASE WHEN survey_responses<5 THEN 'LOW_SAMPLE' ELSE 'OK' END AS sample_flag
            FROM mart.agent_pcs_day WHERE business_date BETWEEN ? AND ?
            ORDER BY business_date, agent_id LIMIT ?"""
        headers, rows = _query(conn, agent_day_sql, [start, end, config.report_limits.get("max_pcs_agent_rows", 100000)])
        ws = report.add_table_sheet(
            "AGENT_DAY", "Exact PCS by agent and day",
            "Governed counters plus ratios. LOW_SAMPLE is a coaching warning; it does not suppress the exact score.",
            headers, rows,
        )
        _color_statuses(report, ws, headers, rows, "sample_flag", {"LOW_SAMPLE": "warn", "OK": "good"})

        team_dataset = pcs_team_day(conn, metric_catalog, start, end)
        headers, rows = team_dataset.headers, team_dataset.rows
        report.add_table_sheet(
            "TEAM_DAY", "Exact PCS by team and day",
            "Counters are additive; KPI values come from the effective-dated semantic metric engine.",
            headers, rows,
        )

        month_dataset = pcs_agent_month(conn, metric_catalog, start, end)
        headers, rows = month_dataset.headers, month_dataset.rows
        ws = report.add_table_sheet(
            "AGENT_MONTH", "Exact monthly agent PCS",
            "This is the presentation-ready monthly grain. A LOW_SAMPLE flag never changes the calculated score.",
            headers, rows,
        )
        _color_statuses(report, ws, headers, rows, "sample_flag", {"LOW_SAMPLE": "warn", "OK": "good"})

        primary = config.pcs.primary_score_question
        participation_question = config.pcs.participation_question
        score_column = f"question_{primary}_score"
        raw_column = f"question_{participation_question}"
        headers, rows = _query(
            conn,
            f"""SELECT c.business_date, c.call_start, c.call_reference_number,
                       c.call_id, c.agent_id,
                       coalesce(d.canonical_name,c.agent_name) AS agent_name,
                       d.team_leader, coalesce(d.lob,c.lob) AS lob,
                       coalesce(d.language,c.language) AS language,
                       c.queue, c.service, c.call_direction,
                       c.question_1 AS raw_q1, c.question_1_score AS parsed_q1,
                       CASE WHEN c.{score_column} IN ({allowed}) THEN c.{score_column} END AS valid_q1_score,
                       CASE WHEN c.{score_column} IN ({allowed})
                                  AND c.{score_column}<={config.pcs.negative_score_maximum:g} THEN '<=3'
                            WHEN c.{score_column} IN ({allowed})
                                  AND c.{score_column}>{config.pcs.negative_score_maximum:g} THEN '>3'
                            WHEN coalesce(trim(c.{raw_column}),'')='' THEN 'BLANK'
                            ELSE 'INVALID' END AS q1_class,
                       c.pcs_status,
                       CASE WHEN coalesce(c.pcs_status,'')='{config.pcs.participation_status}' THEN 1 ELSE 0 END AS participation_denominator,
                       CASE WHEN coalesce(trim(c.{raw_column}),'')<>'' THEN 1 ELSE 0 END AS participation_numerator,
                       c.post_call_survey_mode, c.question_2 AS q2_diagnostic,
                       c.source_file
                FROM core.clean_call_leg c
                LEFT JOIN core.dim_agent d ON d.agent_id=c.agent_id
                WHERE c.business_date BETWEEN ? AND ?
                  AND upper(coalesce(c.call_direction,''))='I'
                  AND (coalesce(trim(c.{raw_column}),'')<>''
                       OR coalesce(c.pcs_status,'')='{config.pcs.participation_status}'
                       OR coalesce(c.post_call_survey_mode,'')='{config.pcs.survey_mode}')
                ORDER BY c.business_date, c.call_start, c.agent_id
                LIMIT ?""",
            [start, end, config.report_limits.get("max_pcs_exception_rows", 100000)],
        )
        ws = report.add_table_sheet(
            "RESPONSE_DETAIL", "Inbound PCS response evidence",
            "Q2 and Mode 2 are shown for diagnosis only. Every official PCS counter comes from Q1/status logic shown here.",
            headers, rows,
        )
        _color_statuses(report, ws, headers, rows, "q1_class", {
            "INVALID": "bad", "BLANK": "future", "<=3": "warn", ">3": "good",
        })

        _quality_sheet(report, conn, start, end, ("fte", "calls", "pcs"))
        _source_health_sheet(report, conn, ("fte", "calls"))
        add_domain_rules_sheet(report, config, spec)
        add_methods_sheet(report, config, spec)
        add_provenance_sheet(report, conn, config, spec, start, end)
        validate_workbook_contract(report, spec)
    except Exception:
        report.close()
        partial.unlink(missing_ok=True)
        raise
    else:
        report.close()
        partial.replace(output)
    return output


def _completed_timeline_day(conn: DatabaseConnection, end: date) -> date:
    latest_allowed = min(end, date.today() - timedelta(days=1))
    value = conn.execute(
        """SELECT max(business_date) FROM mart.attendance_agent_day
           WHERE business_date<=? AND source_loaded=true
             AND shift_state='COMPLETE'""",
        [latest_allowed],
    ).fetchone()[0]
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    return date.fromisoformat(str(value)) if value else latest_allowed


def _as_datetime(value: Any) -> datetime:
    if isinstance(value, datetime):
        return value
    return datetime.fromisoformat(str(value))


def _floor_quarter(value: datetime) -> datetime:
    return value.replace(minute=(value.minute // 15) * 15, second=0, microsecond=0)


def _ceil_quarter(value: datetime) -> datetime:
    floored = _floor_quarter(value)
    return floored if floored == value else floored + timedelta(minutes=15)


def _segment_code(segment: dict[str, Any]) -> str:
    mismatch = str(segment.get("mismatch_type") or "").upper()
    category = str(segment.get("actual_category") or "").upper()
    source = str(segment.get("observed_source") or "").upper()
    if "FUTURE" in mismatch:
        return "F"
    if mismatch == "NO_SHOW":
        return "N"
    if mismatch == "LATE":
        return "T"
    if mismatch == "EARLY_LEAVE":
        return "E"
    if mismatch == "LOGGED_OFF":
        return "O"
    if mismatch == "UNAVAILABLE":
        return "U"
    if bool(segment.get("is_gap")):
        return "G"
    if "LILO" in category or source == "LILO":
        return "L"
    if category == "PRODUCTIVE":
        return "P"
    if category == "AUXILIARY":
        return "A"
    if category in {"LUNCH", "BREAK"}:
        return "B"
    if category == "LOGGED OFF" or "LOGGED_OFF" in category:
        return "O"
    if category == "UNAVAILABLE":
        return "U"
    return "?"


def _add_shift_view(report: ExcelReport, segments: list[dict[str, Any]], report_day: date) -> None:
    ws = report.workbook.add_worksheet("SHIFT_VIEW")
    ws.set_tab_color(COLORS["purple"])
    ws.hide_gridlines(2)
    ws.set_zoom(75)
    ws.merge_range("A1:X1", "WFM HUB  /  FULL-SHIFT EVIDENCE VIEW", report.title)
    ws.merge_range("A2:X2", f"Latest evidence-complete day {report_day:%Y-%m-%d}; 15-minute view for agents with residual correction gaps.", report.subtitle)
    legend = [
        ("P", "Productive", COLORS["green"], COLORS["green_light"]),
        ("A", "Auxiliary", COLORS["teal"], COLORS["teal_light"]),
        ("B", "Break/Lunch", COLORS["amber"], COLORS["amber_light"]),
        ("L", "LILO fallback", COLORS["blue"], COLORS["blue_light"]),
        ("N", "No-show", COLORS["red"], COLORS["red_light"]),
        ("T", "Late", COLORS["red"], COLORS["red_light"]),
        ("E", "Early leave", COLORS["red"], COLORS["red_light"]),
        ("G", "Other gap", COLORS["red"], COLORS["red_light"]),
        ("O", "Logged off", COLORS["red"], COLORS["red_light"]),
        ("U", "Unavailable", COLORS["purple"], COLORS["purple_light"]),
        ("F", "Future", COLORS["future"], COLORS["future_light"]),
        ("?", "Missing/unknown", COLORS["muted"], COLORS["future_light"]),
    ]
    code_formats: dict[str, Any] = {}
    for index, (code, label, font, fill) in enumerate(legend):
        code_formats[code] = report.workbook.add_format({
            "font_name": "Aptos", "font_size": 9, "bold": True,
            "font_color": font, "bg_color": fill, "align": "center",
            "valign": "vcenter", "border": 1, "border_color": COLORS["white"],
        })
        column = index * 2
        ws.write(3, column, code, code_formats[code])
        ws.write(3, column + 1, label, report.note)

    if not segments:
        ws.merge_range("A7:X7", "No residual correction gaps exist on the latest evidence-complete day.", report.note)
        ws.set_column("A:X", 10)
        return

    by_agent: dict[tuple[str, str], list[dict[str, Any]]] = defaultdict(list)
    for segment in segments:
        by_agent[(str(segment["agent_id"]), str(segment.get("agent_name") or ""))].append(segment)
    first = min(_as_datetime(item["scheduled_start"]) for item in segments)
    last = max(_as_datetime(item["scheduled_end"]) for item in segments)
    cursor = _floor_quarter(first)
    slots: list[datetime] = []
    while cursor < _ceil_quarter(last) and len(slots) < 120:
        slots.append(cursor)
        cursor += timedelta(minutes=15)

    metadata = ["Agent ID", "Agent", "Team Leader", "LOB", "Language", "Scheduled Start", "Scheduled End", "Gap Minutes", "Source"]
    header_row = 5
    for column, value in enumerate(metadata):
        ws.write(header_row, column, value, report.header)
    rotated = report.workbook.add_format({
        "font_name": "Aptos", "font_size": 8, "bold": True,
        "font_color": COLORS["white"], "bg_color": COLORS["teal"],
        "rotation": 90, "align": "center", "valign": "vcenter",
    })
    for index, slot in enumerate(slots, len(metadata)):
        day_suffix = "+1" if slot.date() > report_day else ""
        ws.write(header_row, index, slot.strftime("%H:%M") + day_suffix, rotated)
    ws.set_row(header_row, 58)

    for row_index, ((_agent_id, _agent_name), agent_segments) in enumerate(sorted(by_agent.items()), header_row + 1):
        agent_segments.sort(key=lambda item: _as_datetime(item["segment_start"]))
        base = agent_segments[0]
        gap_minutes = sum(int(item.get("segment_minutes") or 0) for item in agent_segments if item.get("is_gap"))
        source = "+".join(sorted({str(item.get("observed_source") or "") for item in agent_segments if item.get("observed_source")}))
        meta_values = [
            base["agent_id"], base.get("agent_name"), base.get("team_leader"),
            base.get("lob"), base.get("language"), base.get("scheduled_start"),
            base.get("scheduled_end"), gap_minutes, source,
        ]
        for column, value in enumerate(meta_values):
            fmt = report.datetime if isinstance(value, datetime) else report.integer if column == 7 else report.body
            ws.write(row_index, column, value, fmt)
        for slot_index, slot in enumerate(slots, len(metadata)):
            right = slot + timedelta(minutes=15)
            scheduled_start = _as_datetime(base["scheduled_start"])
            scheduled_end = _as_datetime(base["scheduled_end"])
            if right <= scheduled_start or slot >= scheduled_end:
                continue
            best = None
            best_rank = (-1, 0.0)
            for segment in agent_segments:
                left_overlap = max(slot, _as_datetime(segment["segment_start"]))
                right_overlap = min(right, _as_datetime(segment["segment_end"]))
                overlap = max(0.0, (right_overlap - left_overlap).total_seconds())
                rank = (1 if bool(segment.get("is_gap")) else 0, overlap)
                if overlap > 0 and rank > best_rank:
                    best, best_rank = segment, rank
            if best is not None:
                code = _segment_code(best)
                ws.write(row_index, slot_index, code, code_formats[code])
        ws.set_row(row_index, 19)
    ws.set_column(0, 0, 13)
    ws.set_column(1, 4, 20)
    ws.set_column(5, 6, 18)
    ws.set_column(7, 7, 12)
    ws.set_column(8, 8, 18)
    if slots:
        ws.set_column(len(metadata), len(metadata) + len(slots) - 1, 3.4)
    ws.freeze_panes(header_row + 1, len(metadata))
    ws.autofilter(header_row, 0, header_row + len(by_agent), len(metadata) - 1)


def build_corrections_workbook(
    conn: DatabaseConnection,
    config: Config,
    start: date,
    end: date,
    output: Path | None = None,
) -> Path:
    """Build the latest completed-day residual correction workbook."""
    generated = datetime.now()
    report_day = _completed_timeline_day(conn, end)
    output = _output_path(config, "corrections", report_day, report_day, generated, output)
    output.parent.mkdir(parents=True, exist_ok=True)
    partial = output.with_name(f"{output.stem}.partial{output.suffix}")
    report = ExcelReport(partial)
    rulebook = load_rulebook(config.home, config.business_rules)
    spec = report_spec(config, "corrections")
    try:
        gap_count, gap_minutes, agents = conn.execute(
            """SELECT count(*), coalesce(sum(residual_minutes),0), count(DISTINCT agent_id)
               FROM mart.correction_residual_segment WHERE business_date=?""",
            [report_day],
        ).fetchone()
        source_missing = conn.execute(
            """SELECT count(*) FROM mart.attendance_agent_day
               WHERE business_date=? AND assignment_type NOT IN ('Off','Planned absence')
                 AND source_loaded=false""",
            [report_day],
        ).fetchone()[0]
        not_corrected = conn.execute(
            """SELECT count(*) FROM mart.correction_candidate
               WHERE business_date=? AND verint_reconciliation='NOT_CORRECTED'""",
            [report_day],
        ).fetchone()[0]
        _add_landing(
            report, "DAY_SUMMARY", spec.title,
            f"Latest evidence-complete day {report_day:%Y-%m-%d}  |  requested through {end:%Y-%m-%d}",
            [
                ("Residual segments", gap_count or 0, "integer"),
                ("Residual gap hours", (gap_minutes or 0) / 60, "decimal"),
                ("Agents to review", agents or 0, "integer"),
                ("Missing evidence rows", source_missing or 0, "integer"),
            ],
            [
                "GAPS contains only residual observed gaps not covered by the union of corrected Verint Activities. Edit only the pale-blue decision columns.",
                "SHIFT_VIEW is the full-shift 15-minute evidence picture. TIMELINE_DATA retains the exact segments behind every colored cell.",
                f"Original candidates still NOT_CORRECTED: {not_corrected:,}. One correction can have several residual segments; the Correction ID remains the import key.",
                "The workbook chooses the newest completed day with loaded LILO or Agent Status evidence. An unfinished current-day shift is never published here as a final early leave.",
            ],
            badge="OBSERVED LILO + AGENT STATUS GAPS  /  VERIFIED AGAINST FINAL VERINT ACTIVITIES",
        )
        add_findings_sheet(report, conn, spec, report_day, report_day)

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
               WHERE r.business_date=?
                 AND coalesce(c.validation_status,'Open') NOT IN ('Injected','Rejected')
               ORDER BY c.priority, r.residual_minutes DESC, r.agent_id, r.residual_start
               LIMIT ?""",
            [report_day, config.report_limits.get("max_gap_rows", 100000)],
        )
        ws = report.add_table_sheet(
            "GAPS", "Residual gaps ready for Verint correction",
            "Blue cells are human decisions. Save this workbook, then use Import correction decisions in WFMHub.cmd.",
            headers, rows,
            editable_headers={"Confirmed Activity", "Validation Status", "Owner", "Comment", "Injected Date"},
        )
        if rows:
            display = [_display_header(item) for item in headers]
            status_col = display.index("Validation Status")
            ws.data_validation(4, status_col, 3 + len(rows), status_col, {
                "validate": "list", "source": ["Open", "Validated", "Injected", "Rejected"],
            })
            suggestions = sorted({str(row[headers.index("suggested_activity")]) for row in rows if row[headers.index("suggested_activity")]})
            activities = list(dict.fromkeys([*suggestions, "Absent", "Late", "Early Leave", "Unpaid Leave", "Vacation"]))[:20]
            activity_col = display.index("Confirmed Activity")
            if activities:
                ws.data_validation(4, activity_col, 3 + len(rows), activity_col, {
                    "validate": "list", "source": activities,
                })
            date_col = display.index("Injected Date")
            ws.data_validation(4, date_col, 3 + len(rows), date_col, {
                "validate": "date", "criteria": "between",
                "minimum": date(2020, 1, 1), "maximum": date(2100, 12, 31),
            })
        _color_statuses(report, ws, headers, rows, "verint_reconciliation", {
            "NOT_CORRECTED": "bad", "PARTIAL": "warn", "CORRECTED": "good",
        })

        timeline_headers, timeline_rows = _query(
            conn,
            """SELECT business_date, agent_id, agent_name, team_leader,
                      ops_manager, lob, language, scheduled_start, scheduled_end,
                      segment_start, segment_end, segment_minutes, planned_state,
                      actual_status, actual_category, mismatch_type, is_gap,
                      observed_source, source_file, evaluation_as_of
               FROM mart.shift_timeline_segment t WHERE business_date=?
                 AND EXISTS (
                     SELECT 1 FROM mart.correction_residual_segment r
                     WHERE r.business_date=t.business_date AND r.agent_id=t.agent_id
                 )
               ORDER BY agent_id, segment_start
               LIMIT ?""",
            [report_day, config.report_limits.get("max_timeline_rows", 250000)],
        )
        timeline_dicts = [dict(zip(timeline_headers, row)) for row in timeline_rows]
        _add_shift_view(report, timeline_dicts, report_day)
        report.add_table_sheet(
            "TIMELINE_DATA", "Exact full-shift timeline segments",
            "Audit data behind SHIFT_VIEW; planned versus observed segments remain exact, not rounded to the visual grid.",
            timeline_headers, timeline_rows,
        )
        _quality_sheet(report, conn, report_day, report_day, ("fte", "schedule", "lilo", "agent_status", "attendance"))
        _source_health_sheet(report, conn, ("fte", "schedule", "lilo", "agent_status"))
        _schedule_variant_sheet(report, conn)
        add_domain_rules_sheet(report, config, spec)
        add_methods_sheet(report, config, spec)
        add_provenance_sheet(report, conn, config, spec, report_day, report_day)
        validate_workbook_contract(report, spec)
    except Exception:
        report.close()
        partial.unlink(missing_ok=True)
        raise
    else:
        report.close()
        partial.replace(output)
    return output


def build_final_absence_workbook(
    conn: DatabaseConnection,
    config: Config,
    start: date,
    end: date,
    output: Path | None = None,
) -> Path:
    """Build final absenteeism solely from corrected Verint Activities marts."""
    generated = datetime.now()
    output = _output_path(config, "absence", start, end, generated, output)
    output.parent.mkdir(parents=True, exist_ok=True)
    partial = output.with_name(f"{output.stem}.partial{output.suffix}")
    report = ExcelReport(partial)
    rulebook = load_rulebook(config.home, config.business_rules)
    metric_catalog = load_metric_catalog(config.home, config.metric_catalog)
    spec = report_spec(config, "absence")
    try:
        totals = conn.execute(
            """SELECT coalesce(sum(planned_net_minutes),0),
                      coalesce(sum(final_absence_minutes),0),
                      coalesce(sum(final_vacation_minutes),0),
                      coalesce(sum(final_unpaid_minutes),0),
                      coalesce(sum(final_unmapped_minutes),0),
                      coalesce(sum(CASE WHEN final_absence_day THEN 1 ELSE 0 END),0)
               FROM mart.verint_final_absence_agent_day
               WHERE business_date BETWEEN ? AND ?""",
            [start, end],
        ).fetchone()
        planned, absence, vacation, unpaid, unmapped, absence_days = totals
        semantic_summary = {
            item.metric_id: item
            for item in aggregate_metric_values(
                conn, metric_catalog, start, end, ["final_absence_rate"], (),
            )
        }
        absence_rate = semantic_summary.get("final_absence_rate").value if semantic_summary.get("final_absence_rate") else None
        _add_landing(
            report, "FINAL_SUMMARY", spec.title,
            f"Period {start:%Y-%m-%d} to {end:%Y-%m-%d}  |  rule {rulebook.version}",
            [
                ("Planned net hours", planned / 60, "decimal"),
                ("Final absence hours", absence / 60, "decimal"),
                ("Final absence rate", absence_rate, "percent"),
                ("Absence agent-days", absence_days, "integer"),
            ],
            [
                "This workbook reads only mart.verint_final_absence_*: corrected Verint Activities are the final ledger. LILO and Agent Status do not create these metrics.",
                f"Daily planned net and every classified numerator are capped at {rulebook.standard_day_hours:g} hours. Overlapping Activities are unioned before daily totals.",
                f"Vacation: {vacation / 60:,.2f} h  |  unpaid: {unpaid / 60:,.2f} h  |  unmapped review: {unmapped / 60:,.2f} h.",
                "Event rows are audit evidence and may overlap across categories. Use AGENT_DAY or LOB_MONTH for totals; do not sum ACTIVITY_EVENTS into a headline KPI.",
            ],
            badge="FINAL  /  VERINT ACTIVITIES ONLY  /  OVERLAP-SAFE DAILY TOTALS",
        )
        add_findings_sheet(report, conn, spec, start, end)

        headers, rows = _query(
            conn,
            """SELECT business_date, agent_id, agent_name, team_leader,
                      ops_manager, lob, market, language, location,
                      scheduled_minutes/60.0 AS scheduled_hours,
                      planned_net_minutes/60.0 AS planned_net_hours,
                      final_absence_minutes/60.0 AS final_absence_hours,
                      final_vacation_minutes/60.0 AS final_vacation_hours,
                      final_unpaid_minutes/60.0 AS final_unpaid_hours,
                      final_shrinkage_minutes/60.0 AS final_shrinkage_hours,
                      final_unmapped_minutes/60.0 AS final_unmapped_hours,
                      final_absence_rate, final_absence_day,
                      final_ledger_status, rule_version, rule_sha256
               FROM mart.verint_final_absence_agent_day
               WHERE business_date BETWEEN ? AND ?
               ORDER BY business_date, agent_id LIMIT ?""",
            [start, end, config.report_limits.get("max_final_absence_rows", 100000)],
        )
        ws = report.add_table_sheet(
            "AGENT_DAY", "Final Verint absence by agent and day",
            "This is the authoritative aggregation grain. All rates use capped, overlap-safe daily counters.",
            headers, rows,
        )
        _color_statuses(report, ws, headers, rows, "final_ledger_status", {
            "UNMAPPED_REVIEW": "bad", "ABSENCE_RECORDED": "warn", "CLEAR": "good",
        })

        month_dataset = final_absence_lob_month(conn, metric_catalog, start, end)
        headers, rows = month_dataset.headers, month_dataset.rows
        report.add_table_sheet(
            "LOB_MONTH", "Final absence by LOB and month",
            "Presentation grain for monthly reporting; KPI values come from configured metric methods.",
            headers, rows,
        )

        headers, rows = _query(
            conn,
            """SELECT business_date, agent_id, agent_name, team_leader,
                      ops_manager, lob, language, activity, category,
                      event_start, event_end, minutes, hours,
                      counts_as_absence, counts_as_vacation, counts_as_unpaid,
                      counts_as_shrinkage, mapped, evidence_type, source_file,
                      rule_version, rule_sha256
               FROM mart.verint_final_absence_event
               WHERE business_date BETWEEN ? AND ?
               ORDER BY business_date, agent_id, event_start LIMIT ?""",
            [start, end, config.report_limits.get("max_final_absence_event_rows", 100000)],
        )
        report.add_table_sheet(
            "ACTIVITY_EVENTS", "Corrected Verint Activities evidence",
            "Audit ledger only. Events may overlap; use AGENT_DAY for final totals.",
            headers, rows,
        )
        unmapped_headers = headers
        unmapped_rows = [row for row in rows if not bool(row[headers.index("mapped")])]
        report.add_table_sheet(
            "UNMAPPED_REVIEW", "Unmapped final Activities",
            "Every row needs a deliberate activity rule decision before payroll use.",
            unmapped_headers, unmapped_rows,
        )

        rule_headers = ["name", "category", "patterns", "match", "planned", "working", "absence", "vacation", "unpaid", "shrinkage"]
        rule_rows = [
            (
                rule.name, rule.category, " | ".join(rule.patterns), rule.match,
                rule.planned, rule.working, rule.absence, rule.vacation,
                rule.unpaid, rule.shrinkage,
            )
            for rule in rulebook.activity_rules
        ]
        report.add_table_sheet(
            "ACTIVITY_RULES", "Central activity classification rules",
            f"Loaded from {rulebook.file.name}; version {rulebook.version}; SHA-256 {rulebook.sha256[:16]}... Maternity classification is explicit here.",
            rule_headers, rule_rows,
        )
        _quality_sheet(report, conn, start, end, ("fte", "schedule"))
        _source_health_sheet(report, conn, ("fte", "schedule"))
        _schedule_variant_sheet(report, conn)
        add_domain_rules_sheet(report, config, spec)
        add_methods_sheet(report, config, spec)
        add_provenance_sheet(report, conn, config, spec, start, end)
        validate_workbook_contract(report, spec)
    except Exception:
        report.close()
        partial.unlink(missing_ok=True)
        raise
    else:
        report.close()
        partial.replace(output)
    return output
