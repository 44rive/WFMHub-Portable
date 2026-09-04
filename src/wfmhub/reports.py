"""Create the shared WFM Hub Excel visual system."""

from __future__ import annotations

from collections import Counter
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any

import xlsxwriter

from .config import Config
from .database import DatabaseConnection
from .report_packs import report_pack, report_pack_folder


COLORS = {
    "dark": "#0B1F33",
    "muted": "#536474",
    "teal": "#007C83",
    "teal_light": "#DFF3F3",
    "canvas": "#F4F7F9",
    "gold": "#D6A84B",
    "rule": "#9AA6B2",
    "thin": "#D8E0E6",
    "blue": "#0563C1",
    "blue_light": "#E3F0FA",
    "green": "#1F7A53",
    "green_light": "#DDF3E8",
    "amber": "#A65F00",
    "amber_light": "#FFF1CC",
    "red": "#B42318",
    "red_light": "#FDE7E5",
    "purple": "#6E56CF",
    "purple_light": "#EEEAFE",
    "future": "#9AA6B2",
    "future_light": "#EEF1F4",
    "white": "#FFFFFF",
}


def _query(conn: DatabaseConnection, sql: str, params: list[Any] | None = None) -> tuple[list[str], list[tuple[Any, ...]]]:
    cursor = conn.execute(sql, params or [])
    return [item[0] for item in cursor.description], cursor.fetchall()


def _display_header(name: str) -> str:
    custom = {
        "agent_id": "Agent ID", "agent_name": "Agent", "business_date": "Date",
        "lob": "LOB", "rta_result": "RTA Result", "status_coverage_percent": "Status Coverage %",
        "conformance_percent": "Conformance %", "attendance_percent": "Attendance %",
        "sl_forecast": "SL Forecast", "sl_required": "SL Required",
        "aht_seconds": "AHT Seconds", "aht_forecast_seconds": "AHT Forecast Seconds",
        "asa_seconds": "ASA Seconds", "fte_forecast": "FTE Forecast", "fte_required": "FTE Required",
        "response_rate": "Response Rate %", "top_box_percent": "Top Box %",
        "low_score_percent": "Low Score %", "pcs_average": "PCS Average",
        "q1_average": "Q1 Average", "q2_average": "Q2 Average",
        "sl_gross": "Gross SL %", "sl_adjusted": "Adjusted SL %",
        "service_level": "Configured SL %", "service_availability": "Service Availability %",
        "abandon_rate": "Abandon Rate %", "absence_rate": "Absence Rate %",
        "vacation_rate": "Vacation Rate %", "shrinkage_rate": "Shrinkage Rate %",
        "rule_sha256": "Rule SHA-256", "rule_version": "Rule Version",
        "forecast_attainment": "Forecast Attainment %",
        "gross_sl_20s": "Gross SL 20s %", "adjusted_sl_20s": "Adjusted SL 20s %",
        "verint_reconciliation": "Verint Final Check",
        "pcs_status": "PCS Status", "post_call_survey_mode": "Post Call Survey Mode",
        "pcs_status_1": "PCS Status 1", "q1_nonblank": "Q1 Nonblank",
        "valid_q1": "Valid Q1", "q1_score_sum": "Q1 Score Sum",
        "score_le_3": "Score <= 3", "score_gt_3": "Score > 3",
        "inbound_call_legs": "Inbound Call Legs", "invalid_q1": "Invalid Q1",
        "sample_state": "Sample State", "coaching_key": "Coaching Key",
        "agent_key": "Agent Selector", "agent_selector": "Agent Selector",
        "latest_day_average": "Latest Day PCS Average",
        "current_mtd_average": "Current MTD PCS Average",
        "prior_mtd_average": "Prior MTD PCS Average",
        "coaching_status": "Coaching Status", "coaching_date": "Coaching Date",
        "coaching_comment": "Coaching Comment",
        "team_lead": "Team Lead", "ops_manager": "Ops Manager",
        "tier_1_bonus_percent": "Tier 1 Bonus %",
        "tier_2_bonus_percent": "Tier 2 Bonus %",
        "tier_1_target": "Tier 1 Target", "tier_2_target": "Tier 2 Target",
        "pcs_participation": "PCS % Participation",
        "absence_percent": "Abs%", "count_value": "Count / Value",
    }
    if name in custom:
        return custom[name]
    lowered = name.casefold()
    if lowered in custom:
        return custom[lowered]
    title = name.replace("_", " ").title()
    acronyms = {
        "Id": "ID", "Lob": "LOB", "Aht": "AHT", "Pcs": "PCS",
        "Kpi": "KPI", "Qm": "QM", "Voc": "VOC", "Fte": "FTE",
        "Pto": "PTO", "Mtd": "MTD", "Rta": "RTA", "Asa": "ASA",
        "Sl": "SL", "Iso": "ISO",
    }
    return " ".join(acronyms.get(word, word) for word in title.split())


class ExcelReport:
    def __init__(self, path: Path):
        self.path = path
        # Tables are intentional: filters and decision columns must expand and
        # remain easy to use. XlsxWriter's constant-memory mode cannot create
        # real Excel Tables, so report row limits are enforced in configuration.
        self.workbook = xlsxwriter.Workbook(path)
        self.workbook.set_properties({
            "title": "WFM Hub Report",
            "subject": "Workforce Management reporting",
            "author": "Anass ASSRI",
            "company": "WFM",
            "comments": "Prepared by Anass ASSRI",
        })
        self.title = self.workbook.add_format({"font_name": "Aptos Display", "font_size": 18, "bold": True, "font_color": COLORS["white"], "bg_color": COLORS["dark"], "align": "left", "valign": "vcenter", "indent": 1})
        self.subtitle = self.workbook.add_format({"font_name": "Aptos", "font_size": 9, "font_color": COLORS["white"], "bg_color": COLORS["teal"], "align": "left", "valign": "vcenter", "indent": 1})
        self.section = self.workbook.add_format({"font_name": "Aptos Display", "font_size": 11, "bold": True, "font_color": COLORS["teal"], "bottom": 2, "bottom_color": COLORS["gold"]})
        self.body = self.workbook.add_format({"font_name": "Aptos", "font_size": 10, "font_color": COLORS["dark"], "bottom": 1, "bottom_color": COLORS["thin"]})
        self.header = self.workbook.add_format({"font_name": "Aptos", "font_size": 10, "bold": True, "font_color": COLORS["white"], "text_wrap": True, "bg_color": COLORS["teal"], "bottom": 2, "bottom_color": COLORS["gold"], "valign": "vcenter"})
        self.editable = self.workbook.add_format({"font_name": "Aptos", "font_size": 10, "font_color": COLORS["blue"], "bg_color": COLORS["blue_light"], "bottom": 1, "bottom_color": COLORS["thin"]})
        self.editable_date = self.workbook.add_format({"font_name": "Aptos", "font_size": 10, "font_color": COLORS["blue"], "bg_color": COLORS["blue_light"], "num_format": "yyyy-mm-dd", "bottom": 1, "bottom_color": COLORS["thin"]})
        self.error = self.workbook.add_format({"font_name": "Aptos", "font_size": 10, "bold": True, "font_color": COLORS["red"], "bg_color": COLORS["red_light"]})
        self.integer = self.workbook.add_format({"font_name": "Aptos", "font_size": 10, "font_color": COLORS["dark"], "num_format": "#,##0", "bottom": 1, "bottom_color": COLORS["thin"]})
        self.decimal = self.workbook.add_format({"font_name": "Aptos", "font_size": 10, "font_color": COLORS["dark"], "num_format": "#,##0.00", "bottom": 1, "bottom_color": COLORS["thin"]})
        self.percent = self.workbook.add_format({"font_name": "Aptos", "font_size": 10, "font_color": COLORS["dark"], "num_format": "0.0%", "bottom": 1, "bottom_color": COLORS["thin"]})
        self.money = self.workbook.add_format({"font_name": "Aptos", "font_size": 10, "font_color": COLORS["dark"], "num_format": '#,##0.00 "MAD"', "bottom": 1, "bottom_color": COLORS["thin"]})
        self.date = self.workbook.add_format({"font_name": "Aptos", "font_size": 10, "font_color": COLORS["dark"], "num_format": "yyyy-mm-dd", "bottom": 1, "bottom_color": COLORS["thin"]})
        self.datetime = self.workbook.add_format({"font_name": "Aptos", "font_size": 10, "font_color": COLORS["dark"], "num_format": "yyyy-mm-dd hh:mm:ss", "bottom": 1, "bottom_color": COLORS["thin"]})
        self.kpi_label = self.workbook.add_format({"font_name": "Aptos", "font_size": 9, "bold": True, "font_color": COLORS["muted"], "bg_color": COLORS["canvas"], "align": "center", "valign": "vcenter", "top": 1, "left": 1, "right": 1, "top_color": COLORS["thin"], "left_color": COLORS["thin"], "right_color": COLORS["thin"]})
        self.kpi_value = self.workbook.add_format({"font_name": "Aptos Display", "font_size": 20, "bold": True, "font_color": COLORS["dark"], "bg_color": COLORS["white"], "align": "center", "valign": "vcenter", "bottom": 1, "left": 1, "right": 1, "bottom_color": COLORS["thin"], "left_color": COLORS["thin"], "right_color": COLORS["thin"]})
        self.note = self.workbook.add_format({"font_name": "Aptos", "font_size": 9, "font_color": COLORS["muted"], "text_wrap": True, "valign": "top"})

    def add_table_sheet(
        self,
        name: str,
        title: str,
        subtitle: str,
        headers: list[str],
        rows: list[tuple[Any, ...]],
        editable_headers: set[str] | None = None,
        exception_column: str | None = None,
    ):
        editable_headers = editable_headers or set()
        worksheet = self.workbook.add_worksheet(name)
        worksheet.set_tab_color(COLORS["teal"])
        worksheet.hide_gridlines(2)
        worksheet.freeze_panes(4, 0)
        last_col = max(0, len(headers) - 1)
        worksheet.merge_range(0, 0, 0, last_col, title, self.title)
        worksheet.merge_range(1, 0, 1, last_col, subtitle, self.subtitle)
        worksheet.set_row(0, 30)
        worksheet.set_row(1, 19)
        worksheet.set_row(3, 30)
        display = [_display_header(header) for header in headers]
        for column, header in enumerate(display):
            worksheet.write(3, column, header, self.header)
        for row_index, values in enumerate(rows, 4):
            worksheet.set_row(row_index, 20)
            for column, value in enumerate(values):
                header = display[column]
                fmt = self.editable_date if header in {"Injected Date", "Coaching Date", "Due Date"} and header in editable_headers else self.editable if header in editable_headers else self.body
                if isinstance(value, datetime):
                    fmt = self.datetime
                elif isinstance(value, date):
                    fmt = self.editable_date if header in editable_headers else self.date
                elif header.endswith(" %") or any(token in header for token in (" Rate", "Participation", "Availability", "Service Level", "Achievement", "Proration")):
                    fmt = self.percent
                elif any(token in header for token in ("Payout", "Bonus Amount", "Salary", "Deduction")):
                    fmt = self.money
                elif any(token in header for token in ("Average", "Hours", "FTE", "Variance", "Movement")):
                    fmt = self.decimal
                elif isinstance(value, (int, float)) and not isinstance(value, bool):
                    fmt = self.integer
                worksheet.write(row_index, column, value, fmt)
        if rows:
            columns = [{"header": header, "header_format": self.header} for header in display]
            worksheet.add_table(3, 0, 3 + len(rows), last_col, {
                "name": "tbl" + "".join(char for char in name.title() if char.isalnum()),
                "style": "Table Style Light 9",
                "columns": columns,
            })
        else:
            worksheet.autofilter(3, 0, 3, last_col)
            worksheet.write(4, 0, "No rows for this period.", self.subtitle)
        for column, header in enumerate(display):
            width = 16
            if any(token in header for token in ("Comment", "Details", "Source", "Path", "Assignment")):
                width = 32
            elif any(token in header for token in ("Agent", "Activity", "Result", "Issue", "Status")):
                width = 22
            elif any(token in header for token in ("Start", "End", "At", "Modified", "Loaded")):
                width = 19
            worksheet.set_column(column, column, width)
        if exception_column and exception_column in display and rows:
            col = display.index(exception_column)
            worksheet.conditional_format(4, col, 3 + len(rows), col, {
                "type": "text", "criteria": "containing", "value": "ERROR", "format": self.error,
            })
        worksheet.set_zoom(90)
        return worksheet

    def close(self) -> None:
        self.workbook.close()


def _summary_sheet(report: ExcelReport, conn: DatabaseConnection, start: date, end: date) -> None:
    worksheet = report.workbook.add_worksheet("SUMMARY")
    worksheet.hide_gridlines(2)
    worksheet.merge_range("A1:F1", "WFM HUB SUMMARY", report.title)
    worksheet.merge_range("A2:F2", f"Reporting period {start:%Y-%m-%d} through {end:%Y-%m-%d}.", report.subtitle)
    metrics = [
        ("Scheduled agent-days", "SELECT count(*) FROM mart.attendance_agent_day WHERE business_date BETWEEN ? AND ? AND assignment_type <> 'Off'", [start, end]),
        ("Scheduled hours", "SELECT coalesce(sum(scheduled_minutes),0)/60.0 FROM mart.attendance_agent_day WHERE business_date BETWEEN ? AND ? AND assignment_type <> 'Off'", [start, end]),
        ("Detected gap hours", "SELECT coalesce(sum(gap_minutes),0)/60.0 FROM mart.correction_candidate WHERE business_date BETWEEN ? AND ?", [start, end]),
        ("Not corrected in Verint", "SELECT count(*) FROM mart.correction_candidate WHERE business_date BETWEEN ? AND ? AND verint_reconciliation='NOT_CORRECTED'", [start, end]),
        ("Corrected in Verint", "SELECT count(*) FROM mart.correction_candidate WHERE business_date BETWEEN ? AND ? AND verint_reconciliation='CORRECTED'", [start, end]),
        ("No-shows", "SELECT count(*) FROM mart.attendance_agent_day WHERE business_date BETWEEN ? AND ? AND attendance_result='No show'", [start, end]),
        ("Present", "SELECT count(*) FROM mart.attendance_agent_day WHERE business_date BETWEEN ? AND ? AND attendance_result='Present'", [start, end]),
        ("Absence hours", "SELECT coalesce(sum(absence_minutes),0)/60.0 FROM mart.absence_agent_day WHERE business_date BETWEEN ? AND ?", [start, end]),
        ("Quality errors", """SELECT count(*) FROM meta.quality_issue
             WHERE severity='ERROR' AND business_date BETWEEN ? AND ?
               AND (source_family IS NULL OR source_family NOT IN ('calls','pcs','intraday','forecast','apbe','apfr'))""", [start, end]),
    ]
    worksheet.write("A4", "KPI", report.header)
    worksheet.write("B4", "Value", report.header)
    for index, (label, sql, params) in enumerate(metrics, 4):
        value = conn.execute(sql, params).fetchone()[0]
        worksheet.write(index, 0, label, report.body)
        worksheet.write(index, 1, value, report.integer)
    issue_rows = conn.execute(
        """SELECT attendance_result, count(*) FROM mart.attendance_agent_day
           WHERE business_date BETWEEN ? AND ?
           GROUP BY 1 ORDER BY count(*) DESC""",
        [start, end],
    ).fetchall()
    worksheet.write("D4", "Attendance result", report.header)
    worksheet.write("E4", "Rows", report.header)
    for index, (label, count) in enumerate(issue_rows, 4):
        worksheet.write(index, 3, label, report.body)
        worksheet.write(index, 4, count, report.integer)
    worksheet.set_column("A:A", 28)
    worksheet.set_column("B:B", 16)
    worksheet.set_column("C:C", 4)
    worksheet.set_column("D:D", 28)
    worksheet.set_column("E:E", 14)
    worksheet.freeze_panes(4, 0)


def _start_sheet(report: ExcelReport, config: Config, start: date, end: date, generated: datetime) -> None:
    ws = report.workbook.add_worksheet("START_HERE")
    ws.hide_gridlines(2)
    ws.merge_range("A1:H1", "WFM HUB", report.title)
    ws.merge_range("A2:H2", f"Last refreshed {generated:%Y-%m-%d %H:%M}; period {start:%Y-%m-%d} to {end:%Y-%m-%d}", report.subtitle)
    sections = [
        (4, "Daily routine", [
            "1. Put untouched exports in their normal source folders.",
            "2. Run WFMHub.cmd and choose Refresh + build report.",
            "3. Open SOURCE_HEALTH first. If it is red or ERROR, stop.",
            "4. Review ATTENDANCE, GAPS and VERINT_FINAL_CHECK. Adherence is intentionally not calculated.",
        ]),
        (11, "Monthly gaps", [
            "Run a custom period from the first through last day of the month.",
            "LILO and Agent Status are the observed evidence. Verint Activities never create an initial gap.",
            "A no-show requires a loaded LILO row with both times blank and no active Agent Status evidence.",
        ]),
        (17, "Correction decisions", [
            "Blue columns on GAPS may be edited. Verint Final Check is automatic from the latest Activities export.",
            "Save the report, then choose Import correction decisions in WFMHub.cmd.",
            "Only Validated or Injected decisions should be used operationally.",
        ]),
        (23, "Configuration", [
            f"Source root: {config.source_root}",
            f"Database: {config.database}",
            f"Central business rules: {config.business_rules}",
            "Agent scope: FTE roster ID or one unique normalized-name match; operational source IDs are preserved.",
        ]),
    ]
    for row, heading, lines in sections:
        ws.write(row - 1, 0, heading, report.section)
        for offset, line in enumerate(lines, 1):
            ws.write(row - 1 + offset, 0, line, report.body)
    ws.set_column("A:A", 100)
    ws.set_column("B:H", 3)


def build_legacy_report(
    conn: DatabaseConnection,
    config: Config,
    start: date,
    end: date,
    output: Path | None = None,
) -> Path:
    generated = datetime.now()
    pack = report_pack("operations")
    output = (
        output
        or report_pack_folder(config, pack.key)
        / f"{pack.filename_prefix}_{start:%Y-%m-%d}_to_{end:%Y-%m-%d}_{generated:%H%M%S_%f}.xlsx"
    ).resolve()
    output.parent.mkdir(parents=True, exist_ok=True)
    partial = output.with_name(f"{output.stem}.partial{output.suffix}")
    report = ExcelReport(partial)
    try:
        _start_sheet(report, config, start, end, generated)
        _summary_sheet(report, conn, start, end)

        headers, rows = _query(conn, """
            SELECT a.business_date, a.agent_id, a.agent_name, a.team_leader, a.ops_manager, a.lob,
                   a.scheduled_start, a.scheduled_end, a.scheduled_minutes, a.assignment_type,
                   a.first_login, a.last_logout, a.raw_late_minutes, a.raw_early_leave_minutes,
                   a.uncoded_late_minutes, a.uncoded_early_leave_minutes, a.no_show_minutes,
                   a.actual_first_seen, a.actual_last_seen, a.actual_evidence,
                   a.status_covered_minutes, a.attendance_result, a.attendance_percent,
                   a.schedule_source, a.lilo_source, a.status_source
            FROM mart.attendance_agent_day a
            WHERE a.business_date BETWEEN ? AND ?
            ORDER BY a.business_date, CASE a.attendance_result WHEN 'No show' THEN 1 WHEN 'Data not loaded' THEN 2 ELSE 5 END, a.agent_name
            LIMIT ?
        """, [start, end, config.report_limits.get("max_attendance_rows", 100000)])
        report.add_table_sheet("ATTENDANCE", "Attendance detail", "One derived row per scheduled Agent ID and day. Raw extracts are not copied here.", headers, rows, exception_column="Attendance Result")

        headers, rows = _query(conn, """
            SELECT correction_id, business_date, agent_id, agent_name, team_leader, ops_manager, lob,
                   scheduled_start, scheduled_end, first_login, last_logout, priority, detected_issue,
                   gap_start, gap_end, gap_minutes, confidence, observed_source,
                   suggested_activity, source_file, verint_reconciliation,
                   verint_activity, verint_category, verint_overlap_minutes, verint_source_file,
                   confirmed_activity, validation_status, owner, comment, injected_date
            FROM mart.correction_candidate
            WHERE business_date BETWEEN ? AND ?
            ORDER BY CASE validation_status WHEN 'Open' THEN 1 WHEN 'Validated' THEN 2 ELSE 5 END,
                     business_date, priority, gap_minutes DESC
            LIMIT ?
        """, [start, end, config.report_limits.get("max_gap_rows", 100000)])
        report.add_table_sheet(
            "GAPS", "Detected correction gaps",
            "Derived candidates. Edit only the blue decision columns, save, then import the workbook through WFMHub.cmd.",
            headers, rows,
            editable_headers={"Confirmed Activity", "Validation Status", "Owner", "Comment", "Injected Date"},
            exception_column="Confidence",
        )

        headers, rows = _query(conn, """
            SELECT business_date, agent_id, agent_name, activity, category,
                   event_start, event_end, minutes, exception_type, source_file
            FROM mart.verint_final_exception
            WHERE business_date BETWEEN ? AND ?
            ORDER BY business_date, agent_name, event_start
        """, [start, end])
        report.add_table_sheet(
            "VERINT_FINAL_CHECK", "Verint final activities without observed gaps",
            "These final Verint codes were not supported by a LILO/Agent Status gap. Review source coverage or the correction.",
            headers, rows, exception_column="Exception Type",
        )

        headers, rows = _query(
            conn,
            """SELECT detected_at, source_family, source_file, business_date, agent_id,
                      issue_type, severity, details
               FROM meta.quality_issue
               WHERE (business_date IS NULL OR business_date BETWEEN ? AND ?)
                 AND (source_family IS NULL OR source_family NOT IN ('calls','pcs','intraday','forecast','apbe','apfr'))
               ORDER BY CASE severity WHEN 'ERROR' THEN 1 ELSE 2 END, business_date, issue_type""",
            [start, end],
        )
        report.add_table_sheet("DATA_QUALITY", "Data quality", "Resolve ERROR rows before using attendance, correction or payroll results.", headers, rows, exception_column="Severity")

        headers, rows = _query(conn, """SELECT source_family, expected_path, newest_file, newest_business_date,
                                               modified_at, loaded_at, row_count, rejected_count,
                                               scoped_out_count, status, details
                                        FROM mart.source_health
                                        WHERE source_family IN ('fte','schedule','lilo','agent_status')
                                        ORDER BY source_family""")
        report.add_table_sheet("SOURCE_HEALTH", "Source health", "What was found, loaded, rejected or missing for every configured feed.", headers, rows, exception_column="Status")
    except Exception:
        report.close()
        partial.unlink(missing_ok=True)
        raise
    else:
        report.close()
        partial.replace(output)
    return output


def build_report(
    conn: DatabaseConnection,
    config: Config,
    start: date,
    end: date,
    output: Path | None = None,
) -> Path:
    """Compatibility entry point for the governed Daily Operations workbook."""
    from .governed_workbooks import build_daily_operations_workbook

    return build_daily_operations_workbook(conn, config, start, end, output)
