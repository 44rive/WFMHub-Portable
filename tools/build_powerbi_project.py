#!/usr/bin/env python3
"""Generate the source-controlled WFMHub Power BI Project (PBIP).

The project deliberately imports only the governed CSV contract written under
Feed/PowerBI.  WFMHub owns source parsing and KPI components; Power BI owns the
star relationships, explicit DAX measures, filters and presentation.
"""

from __future__ import annotations

import hashlib
import json
import shutil
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable


ROOT = Path(__file__).resolve().parents[1]
PROJECT_ROOT = ROOT / "templates" / "powerbi" / "WFMHub BI"
PROJECT_NAME = "WFMHub BI"
REPORT_SCHEMA = "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/report/3.3.0/schema.json"
PAGE_SCHEMA = "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/page/2.1.0/schema.json"
VISUAL_SCHEMA = "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/visualContainer/2.9.0/schema.json"
THEME_FILE = "WFMHub-Premium-7f43c96de6f44c729af8d6981ae229c4.json"


@dataclass(frozen=True)
class Column:
    name: str
    kind: str = "string"
    hidden: bool = False


@dataclass(frozen=True)
class Measure:
    name: str
    dax: str
    format_string: str
    description: str
    folder: str


@dataclass(frozen=True)
class Table:
    name: str
    file: str
    description: str
    columns: tuple[Column, ...]
    measures: tuple[Measure, ...] = ()


def c(name: str, kind: str = "string", hidden: bool = False) -> Column:
    return Column(name, kind, hidden)


def m(name: str, dax: str, fmt: str, description: str, folder: str) -> Measure:
    return Measure(name, dax, fmt, description, folder)


TABLES = (
    Table(
        "Date", "DimDate.csv", "Contiguous reporting calendar for the complete governed feed period.",
        (
            c("Date", "date"), c("Year", "int"), c("Month Number", "int"),
            c("Month"), c("Year Month"), c("ISO Week", "int"), c("Weekday"),
            c("Weekday Number", "int", True), c("Month Start", "date"),
            c("Is Data Through"),
        ),
    ),
    Table(
        "Time", "DimTime.csv", "Reusable 15-minute time-of-day dimension.",
        (c("Time"), c("Quarter Hour Index", "int", True), c("Hour", "int"), c("Hour Label"), c("Time Label")),
    ),
    Table(
        "Employee", "DimEmployee.csv", "Current governed FTE roster and organisation attributes.",
        (
            c("Agent ID", hidden=True), c("Agent"), c("Employment Status"), c("Team Leader"),
            c("Ops Manager"), c("LOB"), c("Management LOB"), c("Market"), c("Language"),
            c("Location"), c("City"), c("FTE", "decimal"),
        ),
    ),
    Table(
        "Management LOB", "DimManagementLOB.csv", "Explicit conformed management LOB used across operational domains.",
        (c("Management LOB"), c("Sort Order", "int", True)),
    ),
    Table(
        "Queue", "DimQueue.csv", "Reviewed queue reference from the active service model.",
        (c("Queue"), c("Source System"), c("Service Scope"), c("Comparison Scope"), c("Designation"), c("Mapping Status")),
        (
            m("Queue Count", "DISTINCTCOUNT('Queue'[Queue])", "#,##0", "Distinct queues visible in the active governed model.", "Data quality"),
            m("Mapped Queue Count", "CALCULATE(DISTINCTCOUNT('Queue'[Queue]), 'Queue'[Mapping Status] = \"MAPPED\")", "#,##0", "Distinct queues explicitly mapped to a governed scope.", "Data quality"),
            m("Queue Mapping %", "DIVIDE([Mapped Queue Count], [Queue Count])", "0.0%", "Mapped queues divided by all queues in context.", "Data quality"),
        ),
    ),
    Table(
        "Service", "FactServiceHour.csv", "Additive service counters at date, hour and queue grain.",
        (
            c("Date", "date", True), c("Time Slot", "int", True), c("Service Scope"),
            c("Management LOB", hidden=True), c("Queue", hidden=True), c("Offered", "int", True),
            c("Answered", "int", True), c("Abandoned", "int", True),
            c("Answered Within Target", "int", True), c("Handled Seconds", "int", True),
            c("SLA Denominator", "int", True), c("Source Files"),
        ),
        (
            m("Offered Calls", "SUM('Service'[Offered])", "#,##0", "Calls offered to the governed queue scope.", "Service"),
            m("Answered Calls", "SUM('Service'[Answered])", "#,##0", "Answered call legs in the governed scope.", "Service"),
            m("Abandoned Calls", "SUM('Service'[Abandoned])", "#,##0", "Abandoned call legs in the governed scope.", "Service"),
            m("Handled in SL", "SUM('Service'[Answered Within Target])", "#,##0", "Answered calls inside the configured service threshold.", "Service"),
            m("SLA Eligible Calls", "SUM('Service'[SLA Denominator])", "#,##0", "Offered calls after the configured short-abandon treatment.", "Service"),
            m("Service Level %", "DIVIDE([Handled in SL], [SLA Eligible Calls])", "0.0%", "Ratio of summed handled-in-threshold calls to summed eligible demand.", "Service"),
            m("Routing Availability %", "DIVIDE([Answered Calls], [Offered Calls])", "0.0%", "Answered divided by offered; this is service routing availability, not agent availability.", "Service"),
            m("Abandon Rate %", "DIVIDE([Abandoned Calls], [Offered Calls])", "0.0%", "Abandoned divided by offered.", "Service"),
            m("AHT Seconds", "DIVIDE(SUM('Service'[Handled Seconds]), [Answered Calls])", "#,##0", "Weighted average handle time from additive handled seconds and answered calls.", "Service"),
            m("Volume Variance", "[Offered Calls] - [Forecast Volume]", "#,##0;[Red]-#,##0", "Actual offered demand minus the governed forecast volume.", "Forecast comparison"),
            m("Volume Variance %", "DIVIDE([Volume Variance], [Forecast Volume])", "0.0%;[Red]-0.0%", "Volume variance divided by forecast volume.", "Forecast comparison"),
            m(
                "Forecast Accuracy %",
                """VAR Grain =
SUMMARIZECOLUMNS(
    'Date'[Date],
    'Time'[Hour],
    'Management LOB'[Management LOB],
    \"Actual Volume\", [Offered Calls],
    \"Forecast Volume At Grain\", [Forecast Volume]
)
VAR Comparable = FILTER(Grain, NOT ISBLANK([Forecast Volume At Grain]))
VAR ForecastTotal = SUMX(Comparable, [Forecast Volume At Grain])
VAR AbsoluteError = SUMX(Comparable, ABS(COALESCE([Actual Volume], 0) - [Forecast Volume At Grain]))
RETURN IF(ForecastTotal <= 0, BLANK(), MAX(0, 1 - DIVIDE(AbsoluteError, ForecastTotal)))""",
                "0.0%", "One minus WAPE across Date, Hour and Management LOB using only intervals with supplied forecast demand.", "Forecast comparison",
            ),
        ),
    ),
    Table(
        "Forecast", "FactForecastInterval.csv", "Verint forecast at its native 15-minute queue interval.",
        (
            c("Date", "date", True), c("Time Slot", "int", True), c("Queue", hidden=True),
            c("Volume Forecast", "decimal", True), c("FTE Required", "decimal", True),
            c("SL Forecast", "decimal", True), c("SL Required", "decimal", True),
            c("AHT Forecast Seconds", "decimal", True), c("Service Scope"),
            c("Management LOB", hidden=True), c("Mapping Status"), c("Source File"),
        ),
        (
            m("Forecast Volume", "SUM('Forecast'[Volume Forecast])", "#,##0", "Summed Verint forecast demand at native interval grain.", "Forecast"),
            m(
                "Required FTE",
                "AVERAGEX(SUMMARIZE('Forecast', 'Forecast'[Date], 'Forecast'[Time Slot], \"Interval FTE\", SUM('Forecast'[FTE Required])), [Interval FTE])",
                "#,##0.0", "Average summed required FTE across selected date/time intervals.", "Forecast",
            ),
            m("Requirement Intervals", "COUNT('Forecast'[FTE Required])", "#,##0", "Native forecast intervals containing an explicit Verint required-FTE value.", "Forecast"),
        ),
    ),
    Table(
        "Staffing", "FactStaffing15Min.csv", "Scheduled and observed capacity at date, interval, roster LOB and language grain.",
        (
            c("Date", "date", True), c("Time Slot", "int", True), c("LOB"), c("Management LOB", hidden=True),
            c("Language"), c("Scheduled HC", "int", True), c("Observed HC", "int", True),
            c("Productive HC", "int", True), c("Auxiliary HC", "int", True),
            c("Gross Scheduled FTE", "decimal", True), c("Planned Time Off FTE", "decimal", True),
            c("Scheduled FTE", "decimal", True), c("Observed FTE", "decimal", True),
            c("Productive FTE", "decimal", True), c("Staffing Variance FTE", "decimal", True),
            c("Staffing Gap FTE", "decimal", True), c("Staffing State"), c("Evidence Basis"),
        ),
        (
            m("Average Scheduled FTE", "AVERAGEX(SUMMARIZE('Staffing', 'Staffing'[Date], 'Staffing'[Time Slot], \"Interval FTE\", SUM('Staffing'[Scheduled FTE])), [Interval FTE])", "#,##0.0", "Average net scheduled FTE across selected intervals after governed time off.", "Capacity"),
            m("Average Observed FTE", "AVERAGEX(SUMMARIZE('Staffing', 'Staffing'[Date], 'Staffing'[Time Slot], \"Interval FTE\", SUM('Staffing'[Observed FTE])), [Interval FTE])", "#,##0.0", "Average observed FTE across selected completed intervals.", "Capacity"),
            m("Average Productive FTE", "AVERAGEX(SUMMARIZE('Staffing', 'Staffing'[Date], 'Staffing'[Time Slot], \"Interval FTE\", SUM('Staffing'[Productive FTE])), [Interval FTE])", "#,##0.0", "Average productive FTE across selected completed intervals.", "Capacity"),
            m("Average Scheduled HC", "AVERAGEX(SUMMARIZE('Staffing', 'Staffing'[Date], 'Staffing'[Time Slot], \"Interval HC\", SUM('Staffing'[Scheduled HC])), [Interval HC])", "#,##0.0", "Average scheduled headcount across selected native intervals.", "Capacity"),
            m("Average Observed HC", "AVERAGEX(SUMMARIZE('Staffing', 'Staffing'[Date], 'Staffing'[Time Slot], \"Interval HC\", SUM('Staffing'[Observed HC])), [Interval HC])", "#,##0.0", "Average observed headcount across selected completed intervals.", "Capacity"),
            m("Capacity Gap FTE", "IF(ISBLANK([Required FTE]), BLANK(), MAX([Required FTE] - [Average Scheduled FTE], 0))", "#,##0.0", "Required FTE less average scheduled capacity, floored at zero; blank when Verint requirement is not supplied.", "Capacity"),
            m("Net Capacity Gap FTE", "IF(ISBLANK([Required FTE]), BLANK(), [Average Scheduled FTE] - [Required FTE])", "#,##0.0;[Red]-#,##0.0", "Average scheduled capacity less required FTE; blank when Verint requirement is not supplied.", "Capacity"),
            m("Observed Gap FTE", "MAX([Average Scheduled FTE] - [Average Observed FTE], 0)", "#,##0.0", "Average scheduled FTE less average observed FTE, floored at zero.", "Capacity"),
            m("Coverage %", "IF(ISBLANK([Required FTE]), BLANK(), DIVIDE([Average Scheduled FTE], [Required FTE]))", "0.0%", "Average scheduled capacity divided by explicit Verint required FTE; blank when requirement is unavailable.", "Capacity"),
        ),
    ),
    Table(
        "Attendance", "FactAttendanceDay.csv", "One governed attendance result per scheduled agent day.",
        (
            c("Date", "date", True), c("Agent Day Key", hidden=True), c("Agent ID", hidden=True),
            c("Agent"), c("Team Leader"), c("Ops Manager"), c("LOB"), c("Management LOB", hidden=True),
            c("Market"), c("Language"), c("Location"), c("Scheduled Start"), c("Scheduled End"),
            c("Scheduled Minutes", "int", True), c("Planned Work Minutes", "int", True),
            c("Planning Overlay"), c("First Login"), c("Last Logout"), c("Attendance Result"),
            c("Call Action"), c("Requires Call"), c("Shift State"), c("Actual Evidence"),
            c("Is Provisional"), c("Late Minutes", "int", True), c("Early Leave Minutes", "int", True),
            c("No Show Minutes", "int", True), c("Status Covered Minutes", "int", True),
        ),
        (
            m("Due HC", "CALCULATE(DISTINCTCOUNT('Attendance'[Agent Day Key]), 'Attendance'[Planned Work Minutes] > 0)", "#,##0", "Scheduled agent-days with positive governed work minutes.", "Attendance"),
            m("Present HC", "CALCULATE(DISTINCTCOUNT('Attendance'[Agent Day Key]), 'Attendance'[Planned Work Minutes] > 0, 'Attendance'[First Login] <> BLANK(), 'Attendance'[First Login] <> \"\")", "#,##0", "Due agent-days with observed presence; late arrivals remain present.", "Attendance"),
            m("No Show HC", "CALCULATE(DISTINCTCOUNT('Attendance'[Agent Day Key]), 'Attendance'[Attendance Result] = \"No show\")", "#,##0", "Evidence-backed completed no-show agent-days only.", "Attendance"),
            m("Unknown / Possible No Show HC", "MAX([Due HC] - [Present HC] - [No Show HC], 0)", "#,##0", "Due population not yet proven present or no-show; never silently promoted to no-show.", "Attendance"),
            m("Callout HC", "CALCULATE(DISTINCTCOUNT('Attendance'[Agent Day Key]), 'Attendance'[Requires Call] = \"1\")", "#,##0", "Attendance cases requiring an operational call or follow-up.", "Attendance"),
            m("Late HC", "CALCULATE(DISTINCTCOUNT('Attendance'[Agent Day Key]), 'Attendance'[Late Minutes] > 0)", "#,##0", "Agent-days with governed late minutes.", "Attendance"),
            m("Early Leave HC", "CALCULATE(DISTINCTCOUNT('Attendance'[Agent Day Key]), 'Attendance'[Early Leave Minutes] > 0)", "#,##0", "Completed agent-days with governed early-leave minutes.", "Attendance"),
            m("On Time HC", "MAX([Present HC] - [Late HC], 0)", "#,##0", "Present agent-days without governed late minutes.", "Attendance"),
            m("Attendance %", "DIVIDE([Present HC], [Due HC])", "0.0%", "Present agent-days divided by due agent-days.", "Attendance"),
            m("Status Evidence Coverage %", "DIVIDE(SUM('Attendance'[Status Covered Minutes]), SUM('Attendance'[Planned Work Minutes]))", "0.0%", "Agent Status-covered minutes divided by governed planned work minutes.", "Data quality"),
        ),
    ),
    Table(
        "Attendance Gap", "FactAttendanceGap.csv", "Exact start/end attendance gaps prepared for human review.",
        (
            c("Date", "date", True), c("Correction ID", hidden=True), c("Agent ID", hidden=True),
            c("Agent"), c("Team Leader"), c("Ops Manager"), c("LOB"), c("Management LOB", hidden=True),
            c("Scheduled Start"), c("Scheduled End"), c("Gap Start"), c("Gap End"),
            c("Gap Minutes", "int", True), c("Detected Issue"), c("Priority"), c("Confidence"),
            c("Suggested Activity"), c("Observed Source"), c("Reconciliation"), c("Source File"),
        ),
        (m("Total Gap Minutes", "SUM('Attendance Gap'[Gap Minutes])", "#,##0", "Summed exact review-gap minutes.", "Attendance review"),),
    ),
    Table(
        "PCS", "FactPCSAgentDay.csv", "Additive PCS and participation counters per governed agent day.",
        (
            c("Date", "date", True), c("Agent Day Key", hidden=True), c("Agent ID", hidden=True), c("Agent"),
            c("Team Leader"), c("Ops Manager"), c("LOB"), c("Management LOB", hidden=True),
            c("Market"), c("Language"), c("Location"), c("Call Legs", "int", True),
            c("Handled Calls", "int", True), c("Inbound Calls", "int", True), c("Outbound Calls", "int", True),
            c("Handle Seconds", "int", True), c("PCS Score Sum", "decimal", True),
            c("Valid PCS", "int", True), c("PCS Status Calls", "int", True),
            c("PCS Participation Responses", "int", True), c("Low Score Responses", "int", True),
            c("Top Box Responses", "int", True), c("Comments Count", "int", True),
        ),
        (
            m("PCS Average", "DIVIDE(SUM('PCS'[PCS Score Sum]), SUM('PCS'[Valid PCS]))", "0.00", "Ratio of summed PCS scores to summed valid responses.", "PCS"),
            m("PCS Participation %", "DIVIDE(SUM('PCS'[PCS Participation Responses]), SUM('PCS'[PCS Status Calls]))", "0.0%", "Summed participation responses divided by summed participation-eligible calls.", "PCS"),
            m("Valid PCS Responses", "SUM('PCS'[Valid PCS])", "#,##0", "Count of valid PCS responses.", "PCS"),
            m("Low Score Response Count", "SUM('PCS'[Low Score Responses])", "#,##0", "Responses inside the configured low-score band.", "PCS"),
            m("Top Box Response Count", "SUM('PCS'[Top Box Responses])", "#,##0", "Responses inside the configured top-box band.", "PCS"),
            m("PCS Index %", "DIVIDE([PCS Average], 5)", "0.0%", "PCS Average expressed against the five-point scale for cross-domain visual comparison.", "PCS"),
            m("Prior Month PCS Average", "CALCULATE([PCS Average], DATEADD('Date'[Date], -1, MONTH))", "0.00", "PCS average for the equivalent prior-month date context.", "PCS comparison"),
            m("PCS Change", "[PCS Average] - [Prior Month PCS Average]", "0.00;[Red]-0.00", "Current PCS average less the equivalent prior-month result.", "PCS comparison"),
        ),
    ),
    Table(
        "PCS Coaching", "FactPCSCoaching.csv", "Exact low-score call queue with Call ID for coaching follow-up.",
        (
            c("Date", "date", True), c("Call Start"), c("Agent ID", hidden=True), c("Agent"),
            c("Team Leader"), c("LOB"), c("Management LOB", hidden=True), c("Call ID"),
            c("Call Key", hidden=True), c("Q1 Score", "decimal"), c("Customer Comment"), c("PCS Status"),
        ),
        (m("Coaching Cases", "COUNTROWS('PCS Coaching')", "#,##0", "Low-score call-level cases available for coaching follow-up.", "PCS coaching"),),
    ),
    Table(
        "Time Off", "FactTimeOff.csv", "Approved PTO and effective Away segments after roster validation.",
        (
            c("Date", "date", True), c("Segment Key", hidden=True), c("Agent Day Key", hidden=True),
            c("Agent ID", hidden=True), c("Agent"), c("Team Leader"), c("Ops Manager"), c("LOB"),
            c("Management LOB", hidden=True), c("Language"), c("Source Kind"), c("Absence Type"),
            c("Record Status"), c("Segment Start"), c("Segment End"), c("Planned Minutes", "int", True),
        ),
        (m("PTO / Away HC", "DISTINCTCOUNT('Time Off'[Agent Day Key])", "#,##0", "Distinct agent-days covered by effective governed PTO or Away segments.", "Time off"),),
    ),
    Table(
        "Final Absence", "FactFinalAbsenceDay.csv", "Final Verint Activities absence and shrinkage result per agent day.",
        (
            c("Date", "date", True), c("Agent Day Key", hidden=True), c("Agent ID", hidden=True), c("Agent"),
            c("Team Leader"), c("Ops Manager"), c("LOB"), c("Management LOB", hidden=True),
            c("Market"), c("Language"), c("Location"), c("Scheduled Minutes", "int", True),
            c("Planned Net Minutes", "int", True), c("Absence Minutes", "int", True),
            c("Vacation Minutes", "int", True), c("Unpaid Minutes", "int", True),
            c("Shrinkage Minutes", "int", True), c("Unmapped Minutes", "int", True),
            c("Final Ledger Status"), c("Rule Version"),
        ),
        (
            m("Final Planned Minutes", "CALCULATE(SUM('Final Absence'[Planned Net Minutes]), 'Final Absence'[Final Ledger Status] IN {\"CLEAR\", \"ABSENCE_RECORDED\"})", "#,##0", "Planned net minutes from finalized rows only.", "Final absence"),
            m("Final Absence Minutes", "CALCULATE(SUM('Final Absence'[Absence Minutes]), 'Final Absence'[Final Ledger Status] IN {\"CLEAR\", \"ABSENCE_RECORDED\"})", "#,##0", "Mapped final absence minutes from finalized rows.", "Final absence"),
            m("Final Shrinkage Minutes", "CALCULATE(SUM('Final Absence'[Shrinkage Minutes]), 'Final Absence'[Final Ledger Status] IN {\"CLEAR\", \"ABSENCE_RECORDED\"})", "#,##0", "Mapped final shrinkage minutes from finalized rows.", "Final absence"),
            m("Final PTO Minutes", "CALCULATE(SUM('Final Absence'[Vacation Minutes]), 'Final Absence'[Final Ledger Status] IN {\"CLEAR\", \"ABSENCE_RECORDED\"})", "#,##0", "Mapped vacation/PTO minutes from finalized rows.", "Final absence"),
            m("Final Absence %", "DIVIDE([Final Absence Minutes], [Final Planned Minutes])", "0.0%", "Final absence minutes divided by finalized planned net minutes.", "Final absence"),
            m("Final Shrinkage %", "DIVIDE([Final Shrinkage Minutes], [Final Planned Minutes])", "0.0%", "Final shrinkage minutes divided by finalized planned net minutes.", "Final absence"),
            m("Final PTO %", "DIVIDE([Final PTO Minutes], [Final Planned Minutes])", "0.0%", "Final vacation minutes divided by finalized planned net minutes.", "Final absence"),
            m("Absence Review HC", "CALCULATE(DISTINCTCOUNT('Final Absence'[Agent Day Key]), NOT('Final Absence'[Final Ledger Status] IN {\"CLEAR\", \"ABSENCE_RECORDED\"}))", "#,##0", "Agent-days not yet in a finalized ledger state.", "Final absence"),
            m("Prior Month Final Absence %", "CALCULATE([Final Absence %], DATEADD('Date'[Date], -1, MONTH))", "0.0%", "Final absence rate for the equivalent prior-month date context.", "Final absence comparison"),
        ),
    ),
    Table(
        "Absence Component", "FactFinalAbsenceComponent.csv", "Mapped Verint activity evidence underlying final absence and shrinkage.",
        (
            c("Date", "date", True), c("Event Key", hidden=True), c("Agent Day Key", hidden=True),
            c("Agent ID", hidden=True), c("Agent"), c("Team Leader"), c("LOB"), c("Management LOB", hidden=True),
            c("Activity"), c("Category"), c("Event Start"), c("Event End"), c("Minutes", "int", True),
            c("Counts As Absence"), c("Counts As Vacation"), c("Counts As Unpaid"), c("Counts As Shrinkage"),
            c("Mapped"), c("Evidence Type"), c("Source File"), c("Rule Version"),
        ),
        (m("Activity Minutes", "SUM('Absence Component'[Minutes])", "#,##0", "Summed mapped/unmapped activity evidence minutes.", "Absence components"),),
    ),
    Table(
        "Finding", "FactFinding.csv", "Deterministic on-demand analysis findings and their evidence pointers.",
        (
            c("Period Start", "date", True), c("Period End", "date", True), c("Finding ID", hidden=True),
            c("Rank", "int"), c("Finding Type"), c("Severity"), c("Domain"), c("Metric ID"),
            c("Source System"), c("LOB"), c("Management LOB", hidden=True), c("Language"),
            c("Team Leader"), c("Agent ID", hidden=True), c("Title"), c("Summary"),
            c("Current Value", "decimal"), c("Reference Value", "decimal"), c("Target Value", "decimal"),
            c("Delta Value", "decimal"), c("Unit"), c("Evidence Dataset"), c("Evidence Filter"), c("Created At"),
        ),
        (m("Finding Count", "COUNTROWS('Finding')", "#,##0", "Count of deterministic analysis findings in context.", "Analysis"),),
    ),
    Table(
        "Source Health", "FactSourceHealth.csv", "Latest governed source freshness and load status.",
        (
            c("Source Family"), c("Expected Path"), c("Newest File"), c("Newest Date", "date"),
            c("Modified At"), c("Loaded At"), c("Rows", "int"), c("Rejected", "int"), c("Status"), c("Details"),
        ),
        (
            m("Source Count", "COUNTROWS('Source Health')", "#,##0", "Configured source families in the latest health snapshot.", "Data quality"),
            m("Sources Ready", "CALCULATE(COUNTROWS('Source Health'), 'Source Health'[Status] IN {\"OK\", \"READY\", \"LOADED\"})", "#,##0", "Source families reporting a ready status.", "Data quality"),
            m("Sources Ready %", "DIVIDE([Sources Ready], [Source Count])", "0.0%", "Ready source families divided by configured source families.", "Data quality"),
            m("Latest Data Date", "MAX('Source Health'[Newest Date])", "dd mmm yyyy", "Latest business date present across source families.", "Data quality"),
        ),
    ),
    Table(
        "Quality Issue", "FactQualityIssue.csv", "Governed refresh quality issues with source and agent evidence.",
        (
            c("Issue ID", hidden=True), c("Run ID", hidden=True), c("Detected At"), c("Source Family"),
            c("Source File"), c("Date", "date", True), c("Agent ID", hidden=True), c("Issue Type"),
            c("Severity"), c("Details"),
        ),
        (
            m("Quality Issues", "COUNTROWS('Quality Issue')", "#,##0", "Quality issues recorded in the selected context.", "Data quality"),
            m("Critical Quality Issues", "CALCULATE(COUNTROWS('Quality Issue'), 'Quality Issue'[Severity] IN {\"ERROR\", \"CRITICAL\"})", "#,##0", "Quality issues at error or critical severity.", "Data quality"),
        ),
    ),
)


RELATIONSHIPS = (
    ("Service", "Date", "Date", "Date"), ("Service", "Time Slot", "Time", "Quarter Hour Index"),
    ("Service", "Management LOB", "Management LOB", "Management LOB"), ("Service", "Queue", "Queue", "Queue"),
    ("Forecast", "Date", "Date", "Date"), ("Forecast", "Time Slot", "Time", "Quarter Hour Index"),
    ("Forecast", "Management LOB", "Management LOB", "Management LOB"), ("Forecast", "Queue", "Queue", "Queue"),
    ("Staffing", "Date", "Date", "Date"), ("Staffing", "Time Slot", "Time", "Quarter Hour Index"),
    ("Staffing", "Management LOB", "Management LOB", "Management LOB"),
    ("Attendance", "Date", "Date", "Date"), ("Attendance", "Agent ID", "Employee", "Agent ID"),
    ("Attendance", "Management LOB", "Management LOB", "Management LOB"),
    ("Attendance Gap", "Date", "Date", "Date"), ("Attendance Gap", "Agent ID", "Employee", "Agent ID"),
    ("Attendance Gap", "Management LOB", "Management LOB", "Management LOB"),
    ("PCS", "Date", "Date", "Date"), ("PCS", "Agent ID", "Employee", "Agent ID"),
    ("PCS", "Management LOB", "Management LOB", "Management LOB"),
    ("PCS Coaching", "Date", "Date", "Date"), ("PCS Coaching", "Agent ID", "Employee", "Agent ID"),
    ("PCS Coaching", "Management LOB", "Management LOB", "Management LOB"),
    ("Time Off", "Date", "Date", "Date"), ("Time Off", "Agent ID", "Employee", "Agent ID"),
    ("Time Off", "Management LOB", "Management LOB", "Management LOB"),
    ("Final Absence", "Date", "Date", "Date"), ("Final Absence", "Agent ID", "Employee", "Agent ID"),
    ("Final Absence", "Management LOB", "Management LOB", "Management LOB"),
    ("Absence Component", "Date", "Date", "Date"), ("Absence Component", "Agent ID", "Employee", "Agent ID"),
    ("Absence Component", "Management LOB", "Management LOB", "Management LOB"),
    ("Finding", "Period End", "Date", "Date"), ("Finding", "Agent ID", "Employee", "Agent ID"),
    ("Finding", "Management LOB", "Management LOB", "Management LOB"),
    ("Quality Issue", "Date", "Date", "Date"), ("Quality Issue", "Agent ID", "Employee", "Agent ID"),
)


PAGES = (
    {
        "title": "Executive Overview", "nav": "EXECUTIVE", "status": "MANAGEMENT",
        "subtitle": "One management pulse across service, attendance, PCS and priorities",
        "slicers": (("Date", "Date", "DATE", "Between"), ("Management LOB", "Management LOB", "LOB", "Dropdown"), ("Date", "Year Month", "MONTH", "Dropdown"), ("Date", "ISO Week", "ISO WEEK", "Dropdown")),
        "cards": (("Service", "Service Level %", "SERVICE LEVEL", "#007C83"), ("Service", "Forecast Accuracy %", "FORECAST ACCURACY", "#3B82F6"), ("Attendance", "Attendance %", "ATTENDANCE", "#10B981"), ("PCS", "PCS Average", "PCS", "#F59E0B")),
        "charts": (
            {"type": "clusteredColumnChart", "title": "LOB PERFORMANCE COMPARISON", "category": ("Management LOB", "Management LOB"), "values": (("Service", "Service Level %"), ("Attendance", "Attendance %"), ("PCS", "PCS Index %"))},
            {"type": "lineChart", "title": "8-WEEK OPERATING TREND", "category": ("Date", "Date"), "values": (("Service", "Service Level %"), ("Attendance", "Attendance %"), ("PCS", "PCS Index %"))},
        ),
        "tables": (("MANAGEMENT PRIORITIES", (("Finding", "Severity"), ("Finding", "Domain"), ("Finding", "Title"), ("Finding", "Summary"), ("Finding", "Current Value"), ("Finding", "Reference Value")), "full"),),
    },
    {
        "title": "Service & Forecast", "nav": "SERVICE", "status": "OPERATIONAL",
        "subtitle": "Intraday service, demand variance and exact queue drivers",
        "slicers": (("Date", "Date", "DATE", "Between"), ("Management LOB", "Management LOB", "LOB", "Dropdown"), ("Service", "Service Scope", "SERVICE SCOPE", "Dropdown"), ("Queue", "Queue", "QUEUE", "Dropdown")),
        "cards": (("Service", "Service Level %", "SERVICE LEVEL", "#007C83"), ("Service", "Offered Calls", "TOTAL ENTERED", "#3B82F6"), ("Service", "Answered Calls", "VOLUME HANDLED", "#10B981"), ("Service", "Volume Variance", "VOLUME VARIANCE", "#F59E0B")),
        "charts": (
            {"type": "lineStackedColumnComboChart", "title": "INTRADAY ACTUAL VS FORECAST", "category": ("Time", "Time Label"), "values": (("Service", "Offered Calls"), ("Forecast", "Forecast Volume")), "secondary": (("Service", "Service Level %"),)},
            {"type": "clusteredBarChart", "title": "SERVICE LEVEL BY LOB", "category": ("Management LOB", "Management LOB"), "values": (("Service", "Service Level %"),)},
        ),
        "tables": (("QUEUE DIAGNOSIS", (("Queue", "Queue"), ("Queue", "Designation"), ("Service", "Offered Calls"), ("Service", "Answered Calls"), ("Service", "Handled in SL"), ("Service", "Service Level %"), ("Service", "AHT Seconds")), "full"),),
    },
    {
        "title": "Attendance Control", "nav": "ATTENDANCE", "status": "OPERATIONAL",
        "subtitle": "Agent Status-first attendance; unknown evidence stays separate",
        "slicers": (("Date", "Date", "DATE", "Between"), ("Management LOB", "Management LOB", "LOB", "Dropdown"), ("Employee", "Team Leader", "TEAM LEADER", "Dropdown"), ("Attendance", "Attendance Result", "RESULT", "Dropdown")),
        "cards": (("Attendance", "Due HC", "DUE", "#3B82F6"), ("Attendance", "Present HC", "PRESENT", "#10B981"), ("Attendance", "No Show HC", "NO SHOW", "#EF4444"), ("Attendance", "Late HC", "LATE", "#F59E0B")),
        "charts": (
            {"type": "hundredPercentStackedBarChart", "title": "ATTENDANCE MIX BY LOB", "category": ("Management LOB", "Management LOB"), "values": (("Attendance", "On Time HC"), ("Attendance", "Late HC"), ("Attendance", "No Show HC"), ("Attendance", "Unknown / Possible No Show HC"))},
            {"type": "lineChart", "title": "SCHEDULED VS PRESENT — 15 MIN", "category": ("Time", "Time Label"), "values": (("Staffing", "Average Scheduled HC"), ("Staffing", "Average Observed HC"))},
        ),
        "tables": (("CALL & FOLLOW-UP LIST", (("Attendance", "Date"), ("Attendance", "Agent"), ("Attendance", "Team Leader"), ("Attendance", "Scheduled Start"), ("Attendance", "First Login"), ("Attendance", "Attendance Result"), ("Attendance", "Call Action")), "full"),),
    },
    {
        "title": "PCS Performance & Coaching", "nav": "PCS", "status": "OPERATIONAL",
        "subtitle": "Current and prior PCS with agent realization and call-level coaching",
        "slicers": (("Date", "Date", "DATE", "Between"), ("Management LOB", "Management LOB", "LOB", "Dropdown"), ("Employee", "Team Leader", "TEAM LEADER", "Dropdown"), ("Employee", "Agent", "AGENT", "Dropdown")),
        "cards": (("PCS", "PCS Average", "PCS", "#007C83"), ("PCS", "PCS Participation %", "PARTICIPATION", "#3B82F6"), ("PCS", "Low Score Response Count", "LOW-SCORE", "#EF4444"), ("PCS Coaching", "Coaching Cases", "COACHING QUEUE", "#F59E0B")),
        "charts": (
            {"type": "clusteredColumnChart", "title": "CURRENT VS PRIOR PCS BY LOB", "category": ("Management LOB", "Management LOB"), "values": (("PCS", "PCS Average"), ("PCS", "Prior Month PCS Average"))},
            {"type": "lineStackedColumnComboChart", "title": "DAILY PCS & PARTICIPATION", "category": ("Date", "Date"), "values": (("PCS", "PCS Average"),), "secondary": (("PCS", "PCS Participation %"),)},
        ),
        "tables": (
            ("AGENT RESULTS", (("Employee", "Team Leader"), ("Employee", "Agent"), ("PCS", "PCS Average"), ("PCS", "Prior Month PCS Average"), ("PCS", "PCS Participation %"), ("PCS", "Valid PCS Responses"), ("PCS", "Low Score Response Count")), "left"),
            ("COACHING QUEUE", (("PCS Coaching", "Date"), ("PCS Coaching", "Agent"), ("PCS Coaching", "Call ID"), ("PCS Coaching", "Q1 Score"), ("PCS Coaching", "Customer Comment")), "right"),
        ),
    },
    {
        "title": "Staffing & Capacity", "nav": "STAFFING", "status": "PLANNING",
        "subtitle": "Schedule capacity and explicit Verint requirement at 15-minute grain",
        "slicers": (("Date", "Date", "DATE", "Between"), ("Management LOB", "Management LOB", "LOB", "Dropdown"), ("Staffing", "LOB", "ROSTER LOB", "Dropdown"), ("Staffing", "Language", "LANGUAGE", "Dropdown")),
        "cards": (("Forecast", "Required FTE", "REQUIRED FTE", "#3B82F6"), ("Staffing", "Average Scheduled FTE", "SCHEDULED FTE", "#007C83"), ("Staffing", "Net Capacity Gap FTE", "NET GAP", "#EF4444"), ("Staffing", "Coverage %", "COVERAGE", "#10B981")),
        "charts": (
            {"type": "lineChart", "title": "INTRADAY REQUIRED VS SCHEDULED", "category": ("Time", "Time Label"), "values": (("Forecast", "Required FTE"), ("Staffing", "Average Scheduled FTE"))},
            {"type": "matrix", "title": "LOB / DAY COVERAGE", "rows": (("Management LOB", "Management LOB"),), "columns": (("Date", "Weekday"),), "values": (("Staffing", "Coverage %"),)},
        ),
        "tables": (("CRITICAL CAPACITY GAPS", (("Date", "Date"), ("Time", "Time Label"), ("Management LOB", "Management LOB"), ("Forecast", "Required FTE"), ("Staffing", "Average Scheduled FTE"), ("Staffing", "Net Capacity Gap FTE"), ("Staffing", "Coverage %")), "full"),),
    },
    {
        "title": "Absence & Shrinkage", "nav": "ABSENCE", "status": "FINAL",
        "subtitle": "Final Verint Activities result with transparent components",
        "slicers": (("Date", "Date", "DATE", "Between"), ("Management LOB", "Management LOB", "LOB", "Dropdown"), ("Employee", "Team Leader", "TEAM LEADER", "Dropdown"), ("Absence Component", "Category", "CATEGORY", "Dropdown")),
        "cards": (("Final Absence", "Final Absence %", "ABSENTEEISM", "#EF4444"), ("Final Absence", "Final Shrinkage %", "SHRINKAGE", "#F59E0B"), ("Final Absence", "Final PTO %", "PTO", "#3B82F6"), ("Final Absence", "Absence Review HC", "OPEN REVIEW", "#007C83")),
        "charts": (
            {"type": "lineChart", "title": "CURRENT VS PRIOR ABSENTEEISM", "category": ("Date", "Date"), "values": (("Final Absence", "Final Absence %"), ("Final Absence", "Prior Month Final Absence %"))},
            {"type": "columnChart", "title": "ABSENCE & SHRINKAGE COMPONENTS", "category": ("Management LOB", "Management LOB"), "values": (("Absence Component", "Activity Minutes"),), "series": ("Absence Component", "Category")},
        ),
        "tables": (("FINAL ABSENCE REVIEW", (("Final Absence", "Date"), ("Final Absence", "Agent"), ("Final Absence", "Team Leader"), ("Final Absence", "Final Ledger Status"), ("Final Absence", "Absence Minutes"), ("Final Absence", "Vacation Minutes"), ("Final Absence", "Shrinkage Minutes")), "full"),),
    },
    {
        "title": "Data Quality & Governance", "nav": "DATA QUALITY", "status": "GOVERNED",
        "subtitle": "Freshness, mapping, evidence coverage and active blockers",
        "slicers": (("Date", "Date", "DATE", "Between"), ("Quality Issue", "Source Family", "SOURCE", "Dropdown"), ("Quality Issue", "Severity", "SEVERITY", "Dropdown"), ("Finding", "Domain", "DOMAIN", "Dropdown")),
        "cards": (("Source Health", "Latest Data Date", "DATA THROUGH", "#007C83"), ("Queue", "Queue Mapping %", "MAPPED QUEUES", "#3B82F6"), ("Attendance", "Status Evidence Coverage %", "STATUS COVERAGE", "#10B981"), ("Quality Issue", "Critical Quality Issues", "CRITICAL BLOCKERS", "#EF4444")),
        "charts": (
            {"type": "table", "title": "SOURCE FRESHNESS", "fields": (("Source Health", "Source Family"), ("Source Health", "Newest Date"), ("Source Health", "Rows"), ("Source Health", "Rejected"), ("Source Health", "Status"))},
            {"type": "columnChart", "title": "QUALITY ISSUES — 30 DAY TREND", "category": ("Quality Issue", "Date"), "values": (("Quality Issue", "Quality Issues"),), "series": ("Quality Issue", "Severity")},
        ),
        "tables": (
            ("ACTIVE QUALITY ISSUES", (("Quality Issue", "Date"), ("Quality Issue", "Source Family"), ("Quality Issue", "Issue Type"), ("Quality Issue", "Severity"), ("Quality Issue", "Details")), "left"),
            ("GOVERNANCE EVIDENCE", (("Source Health", "Source Family"), ("Source Health", "Newest File"), ("Source Health", "Loaded At"), ("Source Health", "Status"), ("Source Health", "Details")), "right"),
        ),
    },
)


def _stable_hex(value: str, length: int) -> str:
    return hashlib.sha256(value.encode("utf-8")).hexdigest()[:length]


def _stable_guid(value: str) -> str:
    value = hashlib.md5(value.encode("utf-8"), usedforsecurity=False).hexdigest()
    return f"{value[:8]}-{value[8:12]}-{value[12:16]}-{value[16:20]}-{value[20:]}"


def _quoted(value: str) -> str:
    return "'" + value.replace("'", "''") + "'"


def _write_json(path: Path, value: object) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(value, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")


def _m_string(value: str) -> str:
    return '"' + value.replace('"', '""') + '"'


def _table_tmdl(spec: Table) -> str:
    measure_names = {measure.name.casefold() for measure in spec.measures}
    column_names = {column.name.casefold() for column in spec.columns}
    collisions = sorted(measure_names & column_names)
    if collisions:
        raise ValueError(
            f"Power BI table {spec.name!r} has measure/column name collisions: {collisions}"
        )
    kind_map = {"string": "string", "int": "int64", "decimal": "decimal", "date": "dateTime"}
    m_map = {"string": "type text", "int": "Int64.Type", "decimal": "type number", "date": "type date"}
    lines = [f"/// {spec.description}", f"table {_quoted(spec.name)}", ""]
    for measure in spec.measures:
        lines.extend(
            [
                f"\t/// {measure.description}",
                f"\tmeasure {_quoted(measure.name)} = ```",
                *[f"\t\t\t{part}" for part in measure.dax.splitlines()],
                "\t\t\t```",
                f"\t\tformatString: {measure.format_string}",
                f"\t\tdisplayFolder: {_quoted(measure.folder)}",
                "",
            ]
        )
    for column in spec.columns:
        lines.extend(
            [
                f"\tcolumn {_quoted(column.name)}",
                f"\t\tdataType: {kind_map[column.kind]}",
                *( ["\t\tisHidden", "\t\tisAvailableInMdx: false"] if column.hidden else [] ),
                *( ["\t\tsummarizeBy: none"] if column.kind in {"string", "date"} or column.hidden else [] ),
                f"\t\tsourceColumn: {column.name}",
                "",
            ]
        )
    select = ", ".join(_m_string(column.name) for column in spec.columns)
    types = ", ".join("{" + _m_string(column.name) + ", " + m_map[column.kind] + "}" for column in spec.columns)
    nullable = ", ".join(
        _m_string(column.name) for column in spec.columns if column.kind != "string"
    )
    source = [
        "let",
        f"    Source = Csv.Document(File.Contents(HubRoot & \"\\Feed\\PowerBI\\{spec.file}\"), [Delimiter=\",\", Encoding=65001, QuoteStyle=QuoteStyle.Csv]),",
        "    Promoted = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),",
        f"    Selected = Table.SelectColumns(Promoted, {{{select}}}, MissingField.Error),",
        f"    NullBlanks = Table.ReplaceValue(Selected, \"\", null, Replacer.ReplaceValue, {{{nullable}}}),",
        f"    Typed = Table.TransformColumnTypes(NullBlanks, {{{types}}}, \"en-US\")",
        "in",
        "    Typed",
    ]
    lines.extend(
        [
            f"\tpartition {_quoted(spec.name)} = m",
            "\t\tmode: import",
            "\t\tsource =",
            *[f"\t\t\t{part}" for part in source],
            "",
            "\tannotation PBI_ResultType = Table",
        ]
    )
    return "\n".join(lines) + "\n"


def _write_model(root: Path) -> None:
    definition = root / f"{PROJECT_NAME}.SemanticModel" / "definition"
    tables = definition / "tables"
    tables.mkdir(parents=True, exist_ok=True)
    (definition / "database.tmdl").write_text(
        "database WFMHubBI\n\tcompatibilityLevel: 1702\n\tcompatibilityMode: powerBI\n\tlanguage: 1033\n",
        encoding="utf-8",
    )
    model = [
        "model Model", "\tculture: en-US", "\tdefaultPowerBIDataSourceVersion: powerBI_V3",
        "\tdiscourageImplicitMeasures", "\tsourceQueryCulture: en-US", "\tdataAccessOptions",
        "\t\tfastCombine", "\t\tlegacyRedirects", "\t\treturnErrorValuesAsNull", "",
        "annotation __PBI_TimeIntelligenceEnabled = 0", "",
    ]
    model.extend(f"ref table {_quoted(spec.name)}" for spec in TABLES)
    model.extend(["", "ref cultureInfo en-US", ""])
    (definition / "model.tmdl").write_text("\n".join(model), encoding="utf-8")
    (definition / "expressions.tmdl").write_text(
        'expression HubRoot = "C:\\WFMHub" meta [IsParameterQuery=true, Type="Text", IsParameterQueryRequired=true]\n'
        '\tannotation PBI_ResultType = Text\n',
        encoding="utf-8",
    )
    relationship_lines: list[str] = []
    for source_table, source_column, target_table, target_column in RELATIONSHIPS:
        label = f"{source_table}.{source_column}->{target_table}.{target_column}"
        relationship_lines.extend(
            [
                f"relationship {_stable_guid(label)}",
                f"\tfromColumn: {_quoted(source_table)}.{_quoted(source_column)}",
                f"\ttoColumn: {_quoted(target_table)}.{_quoted(target_column)}",
                "",
            ]
        )
    (definition / "relationships.tmdl").write_text("\n".join(relationship_lines), encoding="utf-8")
    for spec in TABLES:
        (tables / f"{spec.name}.tmdl").write_text(_table_tmdl(spec), encoding="utf-8")
    _write_json(
        root / f"{PROJECT_NAME}.SemanticModel" / "definition.pbism",
        {
            "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/semanticModel/definitionProperties/1.0.0/schema.json",
            "version": "4.2", "settings": {"qnaEnabled": False},
        },
    )


def _literal(value: str | int | float | bool) -> dict:
    if isinstance(value, bool):
        encoded = "true" if value else "false"
    elif isinstance(value, int):
        encoded = f"{value}L"
    elif isinstance(value, float):
        encoded = f"{value}D"
    else:
        encoded = "'" + value.replace("'", "''") + "'"
    return {"expr": {"Literal": {"Value": encoded}}}


def _color(value: str) -> dict:
    return {"solid": {"color": _literal(value)}}


def _column_field(table: str, column: str) -> dict:
    return {"Column": {"Expression": {"SourceRef": {"Entity": table}}, "Property": column}}


def _measure_field(table: str, measure: str) -> dict:
    return {"Measure": {"Expression": {"SourceRef": {"Entity": table}}, "Property": measure}}


def _projection(table: str, field: str, measure: bool = False) -> dict:
    return {
        "field": _measure_field(table, field) if measure else _column_field(table, field),
        "queryRef": f"{table}.{field}", "nativeQueryRef": field,
    }


def _aggregation_projection(table: str, field: str, function: int = 0) -> dict:
    label = ("Sum", "Avg", "Count", "Min", "Max")[function] if function < 5 else "Aggregate"
    return {
        "field": {"Aggregation": {"Expression": _column_field(table, field), "Function": function}},
        "queryRef": f"{label}({table}.{field})", "nativeQueryRef": f"{label} of {field}",
    }


def _position(x: int, y: int, width: int, height: int, z: int, tab: int | None = None) -> dict:
    return {"x": x, "y": y, "z": z, "height": height, "width": width, "tabOrder": tab if tab is not None else z}


def _no_chrome() -> dict:
    return {
        "background": [{"properties": {"show": _literal(False)}}],
        "border": [{"properties": {"show": _literal(False)}}],
        "visualHeader": [{"properties": {"show": _literal(False)}}],
        "padding": [{"properties": {key: _literal(0) for key in ("top", "bottom", "left", "right")}}],
    }


def _shape(
    name: str, x: int, y: int, width: int, height: int, color: str, z: int,
    *, rounded: bool = False,
) -> dict:
    return {
        "$schema": VISUAL_SCHEMA, "name": name, "position": _position(x, y, width, height, z),
        "visual": {
            "visualType": "shape",
            "objects": {
                "shape": [{"properties": {
                    "tileShape": _literal("rectangleRoundedByPixel" if rounded else "rectangle"),
                    **({"rectangleRoundedCurve": _literal(8)} if rounded else {}),
                }}],
                "fill": [{"properties": {"fillColor": _color(color), "transparency": _literal(0)}, "selector": {"id": "default"}}],
                "outline": [{"properties": {"show": _literal(False)}, "selector": {"id": "default"}}],
            },
            "visualContainerObjects": _no_chrome(),
        },
    }


def _textbox(name: str, text: str, x: int, y: int, width: int, height: int, size: int, color: str, z: int, bold: bool = False, align: str = "left") -> dict:
    return {
        "$schema": VISUAL_SCHEMA, "name": name, "position": _position(x, y, width, height, z),
        "visual": {
            "visualType": "textbox",
            "objects": {"general": [{"properties": {"paragraphs": [{
                "textRuns": [{"value": text, "textStyle": {
                    "fontFamily": "Segoe UI Semibold" if bold else "Segoe UI",
                    "fontSize": f"{size}px", "fontWeight": "bold" if bold else "normal", "color": color,
                }}], "horizontalTextAlignment": align,
            }]}}]},
            "visualContainerObjects": _no_chrome(),
        },
    }


def _title_vco(title: str) -> dict:
    return {
        "title": [{"properties": {
            "show": _literal(True), "text": _literal(title),
            "fontColor": _color("#17324D"), "fontSize": _literal(11),
            "bold": _literal(True), "alignment": _literal("left"),
        }}],
        "background": [{"properties": {"show": _literal(False)}}],
        "border": [{"properties": {"show": _literal(False)}}],
        "visualHeader": [{"properties": {"show": _literal(False)}}],
        "padding": [{"properties": {
            "top": _literal(12), "bottom": _literal(8),
            "left": _literal(14), "right": _literal(14),
        }}],
    }


def _slicer(
    name: str, table: str, column: str, label: str, mode: str,
    x: int, y: int, width: int, height: int,
) -> dict:
    return {
        "$schema": VISUAL_SCHEMA, "name": name,
        "position": _position(x, y, width, height, 100, 100),
        "visual": {
            "visualType": "slicer",
            "query": {"queryState": {"Values": {"projections": [_projection(table, column)]}}},
            "objects": {
                "data": [{"properties": {"mode": _literal(mode)}}],
                "header": [{"properties": {
                    "show": _literal(True), "text": _literal(label),
                    "fontColor": _color("#5B7083"), "textSize": _literal(9),
                    "background": _color("#FFFFFF"),
                }}],
                "items": [{"properties": {
                    "fontColor": _color("#17324D"), "textSize": _literal(10),
                    "background": _color("#FFFFFF"),
                }}],
            },
            "visualContainerObjects": {
                "visualHeader": [{"properties": {"show": _literal(False)}}],
                "padding": [{"properties": {key: _literal(0) for key in ("top", "bottom", "left", "right")}}],
            },
        },
    }


def _card(
    name: str, table: str, measure: str, label: str,
    x: int, y: int, width: int, height: int,
) -> dict:
    return {
        "$schema": VISUAL_SCHEMA, "name": name,
        "position": _position(x, y, width, height, 200, 200),
        "visual": {
            "visualType": "cardVisual",
            "query": {"queryState": {"Data": {"projections": [_projection(table, measure, True)]}}},
            "objects": {
                "outline": [{"properties": {"show": _literal(False)}, "selector": {"id": "default"}}],
                "fillCustom": [{"properties": {"show": _literal(False)}, "selector": {"id": "default"}}],
                "value": [{"properties": {
                    "fontSize": _literal(28), "fontColor": _color("#17324D"),
                    "bold": _literal(True), "horizontalAlignment": _literal("left"),
                }, "selector": {"id": "default"}}],
                "label": [{"properties": {
                    "show": _literal(True), "text": _literal(label),
                    "fontSize": _literal(10), "fontColor": _color("#61758A"),
                }, "selector": {"id": "default"}}],
            },
            "visualContainerObjects": {
                "visualHeader": [{"properties": {"show": _literal(False)}}],
                "padding": [{"properties": {
                    "top": _literal(18), "bottom": _literal(12),
                    "left": _literal(24), "right": _literal(14),
                }}],
            },
        },
    }


def _chart(
    name: str, chart_type: str, title: str, category: tuple[str, str],
    values: Iterable[tuple[str, str]], x: int, y: int, width: int, height: int,
    *, secondary: Iterable[tuple[str, str]] = (),
    series: tuple[str, str] | None = None,
    aggregation: bool = False,
) -> dict:
    projections = [
        _aggregation_projection(table, field) if aggregation else _projection(table, field, True)
        for table, field in values
    ]
    query_state: dict[str, dict] = {
        "Category": {"projections": [_projection(*category)]},
        "Y": {"projections": projections},
    }
    if secondary:
        query_state["Y2"] = {
            "projections": [_projection(table, field, True) for table, field in secondary]
        }
    if series:
        query_state["Series"] = {"projections": [_projection(*series)]}
    return {
        "$schema": VISUAL_SCHEMA, "name": name,
        "position": _position(x, y, width, height, 300, 300),
        "visual": {
            "visualType": chart_type,
            "query": {"queryState": query_state},
            "objects": {
                "legend": [{"properties": {
                    "show": _literal(True), "position": _literal("Top"),
                    "labelColor": _color("#61758A"), "fontSize": _literal(9),
                }}],
                "labels": [{"properties": {
                    "show": _literal(False), "fontSize": _literal(9),
                    "color": _color("#52667A"),
                }}],
                "categoryAxis": [{"properties": {
                    "show": _literal(True), "labelColor": _color("#61758A"),
                    "fontSize": _literal(9), "showAxisTitle": _literal(False),
                }}],
                "valueAxis": [{"properties": {
                    "show": _literal(True), "labelColor": _color("#61758A"),
                    "fontSize": _literal(9), "showAxisTitle": _literal(False),
                    "gridlineShow": _literal(True), "gridlineColor": _color("#E7EDF2"),
                }}],
            },
            "visualContainerObjects": _title_vco(title),
        },
    }


def _table(
    name: str, title: str, fields: Iterable[tuple[str, str]],
    x: int, y: int, width: int, height: int,
) -> dict:
    measure_names = {(spec.name, measure.name) for spec in TABLES for measure in spec.measures}
    projections = [_projection(table, field, (table, field) in measure_names) for table, field in fields]
    return {
        "$schema": VISUAL_SCHEMA, "name": name,
        "position": _position(x, y, width, height, 400, 400),
        "visual": {
            "visualType": "tableEx",
            "query": {"queryState": {"Values": {"projections": projections}}},
            "objects": {
                "columnHeaders": [{"properties": {
                    "columnAdjustment": _literal("growToFit"),
                    "autoSizeColumnWidth": _literal(True),
                    "fontColor": _color("#FFFFFF"), "backColor": _color("#17324D"),
                }}],
                "values": [{"properties": {
                    "fontColorPrimary": _color("#263B50"),
                    "backColorPrimary": _color("#FFFFFF"),
                    "fontColorSecondary": _color("#263B50"),
                    "backColorSecondary": _color("#F5F8FA"),
                    "fontSize": _literal(9), "bold": _literal(False),
                }}],
                "grid": [{"properties": {
                    "gridVertical": _literal(False), "gridHorizontal": _literal(True),
                    "rowPadding": _literal(4), "outlineColor": _color("#E5EBF0"),
                }}],
            },
            "visualContainerObjects": _title_vco(title),
        },
    }


def _matrix(
    name: str, title: str, rows: Iterable[tuple[str, str]],
    columns: Iterable[tuple[str, str]], values: Iterable[tuple[str, str]],
    x: int, y: int, width: int, height: int,
) -> dict:
    return {
        "$schema": VISUAL_SCHEMA, "name": name,
        "position": _position(x, y, width, height, 350, 350),
        "visual": {
            "visualType": "matrix",
            "query": {"queryState": {
                "Rows": {"projections": [_projection(*field) for field in rows]},
                "Columns": {"projections": [_projection(*field) for field in columns]},
                "Values": {"projections": [_projection(table, field, True) for table, field in values]},
            }},
            "objects": {
                "columnHeaders": [{"properties": {
                    "fontColor": _color("#FFFFFF"), "backColor": _color("#17324D"),
                }}],
                "rowHeaders": [{"properties": {
                    "fontColor": _color("#263B50"),
                }}],
                "values": [{"properties": {
                    "fontColorPrimary": _color("#263B50"),
                    "backColorPrimary": _color("#FFFFFF"),
                }}],
            },
            "visualContainerObjects": _title_vco(title),
        },
    }


def _navigator(name: str) -> dict:
    return {
        "$schema": VISUAL_SCHEMA, "name": name,
        "position": _position(18, 128, 140, 545, 20, 20),
        "visual": {
            "visualType": "pageNavigator",
            "objects": {
                "layout": [{"properties": {"columnCount": _literal(1), "rowCount": _literal(7), "cellPadding": _literal(8)}}],
                "pages": [{"properties": {"showHiddenPages": _literal(False), "showTooltipPages": _literal(False), "showByDefault": _literal(True)}}],
                "shape": [{"properties": {"tileShape": _literal("rectangleRoundedByPixel"), "rectangleRoundedCurve": _literal(6)}}],
                "text": [
                    {"properties": {"show": _literal(True), "fontSize": _literal(10), "fontColor": _color("#DCE9F2"), "leftMargin": _literal(10)}, "selector": {"id": "default"}},
                    {"properties": {"show": _literal(True), "fontSize": _literal(10), "bold": _literal(True), "fontColor": _color("#FFFFFF"), "leftMargin": _literal(10)}, "selector": {"id": "selected"}},
                ],
                "fill": [
                    {"properties": {"show": _literal(True), "fillColor": _color("#0B1F33"), "transparency": _literal(100)}, "selector": {"id": "default"}},
                    {"properties": {"show": _literal(True), "fillColor": _color("#007C83"), "transparency": _literal(0)}, "selector": {"id": "selected"}},
                ],
                "outline": [{"properties": {"show": _literal(False)}, "selector": {"id": "default"}}],
            },
            "visualContainerObjects": _no_chrome(),
        },
    }


def _write_visual(page_dir: Path, page_title: str, label: str, visual: dict) -> None:
    visual_id = _stable_hex(f"{page_title}:{label}", 20)
    visual["name"] = visual_id
    _write_json(page_dir / "visuals" / visual_id / "visual.json", visual)


def _write_report(root: Path) -> None:
    report_root = root / f"{PROJECT_NAME}.Report"
    definition = report_root / "definition"
    page_ids: list[str] = []
    for page in PAGES:
        title = page["title"]
        page_id = "ReportSection" + _stable_hex(f"page:{title}", 24)
        page_ids.append(page_id)
        page_dir = definition / "pages" / page_id
        page_dir.mkdir(parents=True, exist_ok=True)
        _write_json(
            page_dir / "page.json",
            {
                "$schema": PAGE_SCHEMA, "name": page_id, "displayName": title,
                "displayOption": "FitToPage", "height": 945, "width": 1680,
                "objects": {
                    "background": [{"properties": {"color": _color("#F3F6F8"), "transparency": _literal(0)}}],
                    "outspace": [{"properties": {"color": _color("#E8EEF3"), "transparency": _literal(0)}}],
                },
            },
        )
        _write_visual(page_dir, title, "sidebar", _shape("", 0, 0, 176, 945, "#0B1F33", 0))
        _write_visual(page_dir, title, "header", _shape("", 176, 0, 1504, 58, "#0B1F33", 1))
        _write_visual(page_dir, title, "accent", _shape("", 176, 58, 1504, 4, "#00A3A8", 2))
        _write_visual(page_dir, title, "brand", _textbox("", "WFM\nHUB", 26, 16, 124, 64, 23, "#FFFFFF", 10, True, "center"))
        _write_visual(page_dir, title, "nav label", _textbox("", "NAVIGATION", 26, 100, 124, 20, 9, "#7F9AB2", 10, True))
        _write_visual(page_dir, title, "nav", _navigator(""))
        _write_visual(page_dir, title, "owner", _textbox("", "Prepared by Anass ASSRI\nWorkforce Management", 18, 878, 140, 42, 9, "#9FB3C8", 10, False, "center"))
        _write_visual(page_dir, title, "page title", _textbox("", title.upper(), 210, 10, 620, 32, 19, "#FFFFFF", 10, True))
        _write_visual(page_dir, title, "subtitle", _textbox("", page["subtitle"], 820, 16, 680, 24, 10, "#C9D7E3", 10, False, "right"))
        _write_visual(page_dir, title, "status panel", _shape("", 1520, 13, 128, 32, "#007C83", 8, rounded=True))
        _write_visual(page_dir, title, "status", _textbox("", page["status"], 1525, 20, 118, 18, 9, "#FFFFFF", 10, True, "center"))

        # One compact selector strip shared by every page.
        _write_visual(page_dir, title, "selector panel", _shape("", 196, 76, 1456, 80, "#FFFFFF", 3, rounded=True))
        slicer_x = (212, 570, 928, 1286)
        for index, ((table, field, label, mode), x) in enumerate(zip(page["slicers"], slicer_x), 1):
            _write_visual(page_dir, title, f"slicer {index}", _slicer("", table, field, label, mode, x, 78, 340, 76))

        # Four aligned KPI cards with a restrained semantic accent.
        card_x = (196, 562, 928, 1294)
        for index, ((table, measure, label, accent), x) in enumerate(zip(page["cards"], card_x), 1):
            _write_visual(page_dir, title, f"card panel {index}", _shape("", x, 168, 350, 128, "#FFFFFF", 3, rounded=True))
            _write_visual(page_dir, title, f"card accent {index}", _shape("", x, 168, 8, 128, accent, 4, rounded=True))
            _write_visual(page_dir, title, f"card {index}", _card("", table, measure, label, x + 8, 170, 338, 124))

        # The two analytical panels occupy the same visual rhythm on all pages.
        chart_positions = ((196, 316, 716, 304), (928, 316, 724, 304))
        for index, (chart, position) in enumerate(zip(page["charts"], chart_positions), 1):
            x, y, width, height = position
            _write_visual(page_dir, title, f"chart panel {index}", _shape("", x, y, width, height, "#FFFFFF", 3, rounded=True))
            if chart["type"] == "matrix":
                visual = _matrix(
                    "", chart["title"], chart["rows"], chart["columns"],
                    chart["values"], x + 8, y + 6, width - 16, height - 12,
                )
            elif chart["type"] == "table":
                visual = _table(
                    "", chart["title"], chart["fields"],
                    x + 8, y + 6, width - 16, height - 12,
                )
            else:
                visual = _chart(
                    "", chart["type"], chart["title"], chart["category"],
                    chart["values"], x + 8, y + 6, width - 16, height - 12,
                    secondary=chart.get("secondary", ()),
                    series=chart.get("series"),
                    aggregation=chart.get("aggregation", False),
                )
            _write_visual(page_dir, title, f"chart {index}", visual)

        table_layouts = {
            "full": (196, 638, 1456, 258),
            "left": (196, 638, 716, 258),
            "right": (928, 638, 724, 258),
        }
        for index, (table_title, table_fields, layout) in enumerate(page["tables"], 1):
            x, y, width, height = table_layouts[layout]
            _write_visual(page_dir, title, f"table panel {index}", _shape("", x, y, width, height, "#FFFFFF", 3, rounded=True))
            _write_visual(page_dir, title, f"table {index}", _table("", table_title, table_fields, x + 8, y + 6, width - 16, height - 12))

        _write_visual(page_dir, title, "footer", _textbox("", "WFMHub governed feed  •  Feed\\PowerBI\\POWERBI_MANIFEST_CURRENT.csv", 196, 910, 1456, 18, 8, "#64748B", 10, False, "right"))

    _write_json(
        definition / "pages" / "pages.json",
        {"$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/pagesMetadata/1.0.0/schema.json", "pageOrder": page_ids, "activePageName": page_ids[0]},
    )
    _write_json(
        definition / "version.json",
        {"$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/versionMetadata/1.0.0/schema.json", "version": "2.0.0"},
    )
    theme_source = ROOT / "templates" / "powerbi" / "WFMHub-Premium-Theme.json"
    theme = json.loads(theme_source.read_text(encoding="utf-8"))
    theme["name"] = THEME_FILE
    theme["$schema"] = "https://raw.githubusercontent.com/microsoft/powerbi-desktop-samples/main/Report%20Theme%20JSON%20Schema/reportThemeSchema-2.157.json"
    theme.setdefault("visualStyles", {}).setdefault("tableEx", {}).setdefault("*", {})["columnHeaders"] = [{"autoSizeColumnWidth": True, "columnAdjustment": "growToFit"}]
    theme_path = report_root / "StaticResources" / "RegisteredResources" / THEME_FILE
    _write_json(theme_path, theme)
    _write_json(
        definition / "report.json",
        {
            "$schema": REPORT_SCHEMA,
            "themeCollection": {
                "baseTheme": {"name": "CY19SU12", "reportVersionAtImport": {"visual": "1.8.44", "report": "2.0.44", "page": "1.3.44"}, "type": "SharedResources"},
                "customTheme": {"name": THEME_FILE, "reportVersionAtImport": {"visual": "1.8.44", "report": "2.0.44", "page": "1.3.44"}, "type": "RegisteredResources"},
            },
            "resourcePackages": [{"name": "RegisteredResources", "type": "RegisteredResources", "items": [{"name": THEME_FILE, "path": THEME_FILE, "type": "CustomTheme"}]}],
            "settings": {"useStylableVisualContainerHeader": True, "allowChangeFilterTypes": True, "useEnhancedTooltips": True},
            "slowDataSourceSettings": {"isCrossHighlightingDisabled": False, "isSlicerSelectionsButtonEnabled": False, "isFilterSelectionsButtonEnabled": False, "isFieldWellButtonEnabled": False, "isApplyAllButtonEnabled": False},
        },
    )
    _write_json(
        report_root / "definition.pbir",
        {"$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definitionProperties/2.0.0/schema.json", "version": "4.0", "datasetReference": {"byPath": {"path": f"../{PROJECT_NAME}.SemanticModel"}}},
    )
    _write_json(
        report_root / ".platform",
        {
            "$schema": "https://developer.microsoft.com/json-schemas/fabric/gitIntegration/platformProperties/2.0.0/schema.json",
            "metadata": {"type": "Report", "displayName": PROJECT_NAME},
            "config": {"version": "2.0", "logicalId": _stable_guid("WFMHub BI Report")},
        },
    )


def build() -> Path:
    if PROJECT_ROOT.exists():
        shutil.rmtree(PROJECT_ROOT)
    PROJECT_ROOT.mkdir(parents=True)
    _write_model(PROJECT_ROOT)
    _write_report(PROJECT_ROOT)
    _write_json(
        PROJECT_ROOT / f"{PROJECT_NAME}.SemanticModel" / ".platform",
        {
            "$schema": "https://developer.microsoft.com/json-schemas/fabric/gitIntegration/platformProperties/2.0.0/schema.json",
            "metadata": {"type": "SemanticModel", "displayName": PROJECT_NAME},
            "config": {"version": "2.0", "logicalId": _stable_guid("WFMHub BI Semantic Model")},
        },
    )
    _write_json(
        PROJECT_ROOT / f"{PROJECT_NAME}.pbip",
        {"$schema": "https://developer.microsoft.com/json-schemas/fabric/pbip/pbipProperties/1.0.0/schema.json", "version": "1.0", "artifacts": [{"report": {"path": f"{PROJECT_NAME}.Report"}}], "settings": {"enableAutoRecovery": True}},
    )
    (PROJECT_ROOT / "PROJECT_VERSION.txt").write_text("3\n", encoding="utf-8")
    (PROJECT_ROOT / "README.txt").write_text(
        "WFMHUB BI\n=========\n\n"
        "Open WFMHub BI.pbip with Microsoft Power BI Desktop, then choose Home > Refresh.\n"
        "The HubRoot parameter is set automatically when this project is opened through POWERBI.cmd.\n"
        "The model reads only Feed\\PowerBI CSVs and never reads SQLite or raw extracts.\n",
        encoding="utf-8",
    )
    reference = [
        "// AUDIT REFERENCE — generated by tools/build_powerbi_project.py",
        "// The live measures are stored beside their owning facts in WFMHub BI.SemanticModel/definition/tables.",
        "// Ratios always divide summed governed components; Power BI does not reclassify source data.",
        "",
    ]
    for spec in TABLES:
        if not spec.measures:
            continue
        reference.append(f"// {spec.name}")
        reference.extend(f"{measure.name} = {measure.dax}" for measure in spec.measures)
        reference.append("")
    (ROOT / "templates" / "powerbi" / "WFMHub-Measures.dax").write_text(
        "\n".join(reference), encoding="utf-8",
    )
    return PROJECT_ROOT


if __name__ == "__main__":
    print(build())
