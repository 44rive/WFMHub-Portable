"""Create the shared WFM Hub Excel visual system."""

from __future__ import annotations

from collections import Counter
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any

import xlsxwriter

from .config import Config
from .database import DatabaseConnection
from .design import COLORS
from .report_packs import report_pack, report_pack_folder


def _query(conn: DatabaseConnection, sql: str, params: list[Any] | None = None) -> tuple[list[str], list[tuple[Any, ...]]]:
    cursor = conn.execute(sql, params or [])
    return [item[0] for item in cursor.description], cursor.fetchall()


def _display_header(name: str) -> str:
    custom = {
        "agent_id": "Agent ID", "agent_name": "Agent", "business_date": "Date",
        "lob": "LOB", "rta_result": "RTA Result", "status_coverage_percent": "Status Coverage %",
        "attendance_percent": "Attendance %",
        "sl_forecast": "SL Forecast", "sl_required": "SL Required",
        "aht_seconds": "AHT Seconds", "aht_forecast_seconds": "AHT Forecast Seconds",
        "asa_seconds": "ASA Seconds", "fte_forecast": "FTE Forecast", "fte_required": "FTE Required",
        "response_rate": "Response Rate %", "top_box_percent": "Top Box %",
        "low_score_percent": "Low Score %", "pcs_average": "PCS Average",
        "q1_average": "Q1 Average", "q2_average": "Q2 Average",
        "sl_gross": "Gross SL %", "sl_adjusted": "Adjusted SL %",
        "service_level": "Configured SL %", "service_availability": "Routed Rate %",
        "abandon_rate": "Abandon Rate %", "absence_rate": "Absence Rate %",
        "vacation_rate": "Vacation Rate %", "shrinkage_rate": "Shrinkage Rate %",
        "rule_sha256": "Rule SHA-256", "rule_version": "Rule Version",
        "forecast_attainment": "Forecast Attainment %",
        "gross_sl_20s": "Gross SL 20s %", "adjusted_sl_20s": "Adjusted SL 20s %",
        "verint_reconciliation": "Residual Status",
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
        "abs_hc": "ABS HC", "tsl": "TSL",
        "connected_in_sla": "Connected in SLA",
        "lost_5_to_sla": "Lost 5s to SLA",
        "sla_denominator": "SLA Denominator",
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
        "Sl": "SL", "Tsl": "TSL", "Hc": "HC", "Iso": "ISO",
        "Sha-256": "SHA-256",
    }
    return " ".join(acronyms.get(word, word) for word in title.split())


_DATE_HEADERS = {
    "Date", "As Of Date", "Period Start", "Period End", "Data Through",
    "Coaching Date", "Due Date", "Reviewed Date", "Injected Date",
}
_DATETIME_HEADERS = {
    "Call Start", "Call End", "Report Generated At", "Feed Refreshed At",
}


def _excel_temporal_value(header: str, value: Any) -> tuple[Any, str | None]:
    """Normalize ISO-like report dates so Excel never displays raw serials.

    SQLite adapters may return a date, datetime, ISO text, or an Excel serial.
    The header is the stable report contract, so it is safer than guessing from
    the runtime value alone.
    """

    kind = "datetime" if header in _DATETIME_HEADERS else "date" if header in _DATE_HEADERS else None
    if kind is None:
        return value, None
    if isinstance(value, datetime):
        return (value if kind == "datetime" else value.date()), kind
    if isinstance(value, date) or isinstance(value, (int, float)):
        return value, kind
    if isinstance(value, str):
        text = value.strip()
        if not text:
            return value, kind
        try:
            parsed = datetime.fromisoformat(text.replace("Z", "+00:00")).replace(tzinfo=None)
            return (parsed if kind == "datetime" else parsed.date()), kind
        except ValueError:
            try:
                parsed_date = date.fromisoformat(text[:10])
                return (
                    datetime.combine(parsed_date, datetime.min.time())
                    if kind == "datetime" else parsed_date
                ), kind
            except ValueError:
                return value, kind
    return value, kind


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
                value, temporal_kind = _excel_temporal_value(header, value)
                fmt = self.editable_date if header in {"Injected Date", "Coaching Date", "Due Date"} and header in editable_headers else self.editable if header in editable_headers else self.body
                if temporal_kind == "datetime":
                    fmt = self.datetime
                elif temporal_kind == "date":
                    fmt = self.editable_date if header in editable_headers else self.date
                elif isinstance(value, datetime):
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
