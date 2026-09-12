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
        (
            c("Service Key", hidden=True), c("Queue"), c("Source System"),
            c("Service Scope"), c("Management LOB", hidden=True),
            c("Comparison Scope"), c("Designation"), c("Mapping Status"),
        ),
        (
            m("Queue Count", "DISTINCTCOUNT('Queue'[Service Key])", "#,##0", "Distinct service-scope queue members visible in the active governed model.", "Data quality"),
            m("Mapped Queue Count", "CALCULATE(DISTINCTCOUNT('Queue'[Service Key]), 'Queue'[Mapping Status] = \"MAPPED\")", "#,##0", "Distinct service-scope queue members explicitly mapped to a governed scope.", "Data quality"),
            m("Queue Mapping %", "DIVIDE([Mapped Queue Count], [Queue Count])", "0.0%", "Mapped queues divided by all queues in context.", "Data quality"),
        ),
    ),
    Table(
        "Queue Coverage", "FactQueueCoverage.csv", "Inbound queue entries with explicit reviewed mapping status.",
        (
            c("Date", "date", True), c("Service Key", hidden=True), c("Queue", hidden=True), c("Service Scope"),
            c("Management LOB", hidden=True), c("Comparison Scope"), c("Mapping Status"),
            c("Inbound Entries", "int", True), c("Mapped Inbound Entries", "int", True),
        ),
        (
            m("Inbound Queue Entries", "SUM('Queue Coverage'[Inbound Entries])", "#,##0", "Inbound Call-by-Call queue entries evaluated for mapping.", "Data quality"),
            m("Mapped Inbound Queue Entries", "SUM('Queue Coverage'[Mapped Inbound Entries])", "#,##0", "Inbound queue entries with an explicit governed mapping.", "Data quality"),
            m("Queue Mapped %", "DIVIDE([Mapped Inbound Queue Entries], [Inbound Queue Entries])", "0.0%", "Mapped inbound queue entries divided by all inbound queue entries.", "Data quality"),
        ),
    ),
    Table(
        "Driver", "DimDriver.csv", "Disconnected deterministic driver labels used by the pressure signal visual.",
        (c("Driver"), c("Sort Order", "int", True)),
    ),
    Table(
        "Capacity Stage", "DimCapacityStage.csv", "Disconnected capacity bridge stages in approved display order.",
        (c("Capacity Stage"), c("Sort Order", "int", True)),
    ),
    Table(
        "Service", "FactService15Min.csv", "Additive service counters at date, 15-minute interval and queue grain.",
        (
            c("Date", "date", True), c("Interval Start"), c("Interval End"),
            c("Time Slot", "int", True), c("Service Key", hidden=True), c("Service Scope"),
            c("Management LOB", hidden=True), c("Queue", hidden=True), c("Offered", "int", True),
            c("Answered", "int", True), c("Abandoned", "int", True),
            c("Short Abandoned", "int", True), c("Abandoned Within Target", "int", True),
            c("Answered Within Target", "int", True), c("Handled Seconds", "int", True),
            c("SLA Denominator", "int", True), c("SL Target", "decimal", True), c("Source Files"),
        ),
        (
            m("Offered Calls", "SUM('Service'[Offered])", "#,##0", "Calls offered to the governed queue scope.", "Service"),
            m("Answered Calls", "SUM('Service'[Answered])", "#,##0", "Answered call legs in the governed scope.", "Service"),
            m("Abandoned Calls", "SUM('Service'[Abandoned])", "#,##0", "Abandoned call legs in the governed scope.", "Service"),
            m("Handled in SL", "SUM('Service'[Answered Within Target])", "#,##0", "Answered calls inside the configured service threshold.", "Service"),
            m("SLA Eligible Calls", "SUM('Service'[SLA Denominator])", "#,##0", "Offered calls after the configured short-abandon treatment.", "Service"),
            m("Service Level %", "DIVIDE([Handled in SL], [SLA Eligible Calls])", "0.0%", "Ratio of summed handled-in-threshold calls to summed eligible demand.", "Service"),
            m("SL Target %", "MAX('Service'[SL Target])", "0.0%", "Configured service target in the current single-LOB context.", "Service"),
            m("SL Gap Points", "([Service Level %] - [SL Target %]) * 100", "0.0;[Red]-0.0", "Actual service level less the configured target, expressed in percentage points.", "Service"),
            m("LOBs With Service", "COUNTROWS(FILTER(VALUES('Management LOB'[Management LOB]), NOT ISBLANK([Service Level %])))", "#,##0", "Management LOBs with governed service demand in context.", "Service"),
            m("LOBs On Target", "COUNTROWS(FILTER(VALUES('Management LOB'[Management LOB]), NOT ISBLANK([Service Level %]) && [Service Level %] >= [SL Target %]))", "#,##0", "Management LOBs meeting their configured service target.", "Service"),
            m("LOBs On Target Label", "FORMAT([LOBs On Target], \"0\") & \" / \" & FORMAT([LOBs With Service], \"0\")", "", "Compact on-target LOB count for the daily command card.", "Service"),
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
    'Time'[Quarter Hour Index],
    'Management LOB'[Management LOB],
    \"Actual Volume\", [Offered Calls],
    \"Forecast Volume At Grain\", [Forecast Volume]
)
VAR Comparable = FILTER(Grain, NOT ISBLANK([Forecast Volume At Grain]))
VAR ForecastTotal = SUMX(Comparable, [Forecast Volume At Grain])
VAR AbsoluteError = SUMX(Comparable, ABS(COALESCE([Actual Volume], 0) - [Forecast Volume At Grain]))
RETURN IF(ForecastTotal <= 0, BLANK(), MAX(0, 1 - DIVIDE(AbsoluteError, ForecastTotal)))""",
                "0.0%", "One minus WAPE across Date, native 15-minute slot and Management LOB using only intervals with supplied forecast demand.", "Forecast comparison",
            ),
            m("Forecast Bias %", "DIVIDE([Offered Calls] - [Forecast Volume], [Forecast Volume])", "0.0%;[Red]-0.0%", "Signed actual-minus-forecast demand divided by forecast demand.", "Forecast comparison"),
            m("AHT Error Seconds", "[AHT Seconds] - [Forecast AHT Seconds]", "#,##0;[Red]-#,##0", "Actual weighted AHT less forecast weighted AHT.", "Forecast comparison"),
            m(
                "Peak Accuracy %",
                """VAR Grain =
SUMMARIZECOLUMNS(
    'Date'[Date], 'Time'[Quarter Hour Index], 'Management LOB'[Management LOB],
    \"Actual\", [Offered Calls], \"Forecast\", [Forecast Volume]
)
VAR Comparable = FILTER(Grain, NOT ISBLANK([Forecast]) && [Forecast] > 0)
VAR PeakThreshold = PERCENTILEX.INC(Comparable, [Forecast], 0.9)
VAR Peak = FILTER(Comparable, [Forecast] >= PeakThreshold)
VAR ForecastTotal = SUMX(Peak, [Forecast])
VAR AbsoluteError = SUMX(Peak, ABS(COALESCE([Actual], 0) - [Forecast]))
RETURN IF(ForecastTotal <= 0, BLANK(), MAX(0, 1 - DIVIDE(AbsoluteError, ForecastTotal)))""",
                "0.0%", "One minus WAPE for the top forecast-demand decile in the selected context.", "Forecast comparison",
            ),
            m("Demand Pressure %", "ABS([Volume Variance %])", "0.0%", "Absolute demand variance used as an evidence signal, not an SL-point attribution.", "Driver signals"),
            m("Schedule Pressure %", "IF(ISBLANK([Required FTE]), BLANK(), MAX(DIVIDE([Required FTE] - [Average Scheduled FTE], [Required FTE]), 0))", "0.0%", "Relative required-versus-scheduled shortage signal.", "Driver signals"),
            m("Attendance Pressure %", "MAX(1 - [Presence Realisation %], 0)", "0.0%", "Relative elapsed scheduled capacity not observed.", "Driver signals"),
            m("AUX Pressure %", "MAX(1 - [Productive Realisation %], 0)", "0.0%", "Observed time not classified as governed productive work.", "Driver signals"),
            m("AHT Pressure %", "ABS(DIVIDE([AHT Error Seconds], [Forecast AHT Seconds]))", "0.0%", "Absolute relative AHT variance signal.", "Driver signals"),
            m("Driver Pressure %", "SWITCH(SELECTEDVALUE('Driver'[Driver]), \"Demand\", [Demand Pressure %], \"Schedule\", [Schedule Pressure %], \"Attendance\", [Attendance Pressure %], \"AUX\", [AUX Pressure %], \"AHT\", [AHT Pressure %])", "0.0%", "Selected deterministic pressure signal; it is not an SL-point causal contribution.", "Driver signals"),
            m(
                "Primary Driver",
                """VAR Signals =
UNION(
    ROW(\"Driver\", \"DEMAND\", \"Signal\", [Demand Pressure %]),
    ROW(\"Driver\", \"SCHEDULE\", \"Signal\", [Schedule Pressure %]),
    ROW(\"Driver\", \"ATTENDANCE\", \"Signal\", [Attendance Pressure %]),
    ROW(\"Driver\", \"AUX\", \"Signal\", [AUX Pressure %]),
    ROW(\"Driver\", \"AHT\", \"Signal\", [AHT Pressure %])
)
RETURN MAXX(TOPN(1, FILTER(Signals, NOT ISBLANK([Signal])), [Signal], DESC, [Driver], ASC), [Driver])""",
                "", "Largest supported pressure signal in context; retained as a review lead, not a causal score.", "Driver signals",
            ),
        ),
    ),
    Table(
        "Forecast", "FactForecastInterval.csv", "Verint forecast at its native 15-minute queue interval.",
        (
            c("Date", "date", True), c("Time Slot", "int", True), c("Service Key", hidden=True), c("Queue", hidden=True),
            c("Volume Forecast", "decimal", True), c("FTE Required", "decimal", True),
            c("SL Forecast", "decimal", True), c("SL Required", "decimal", True),
            c("AHT Forecast Seconds", "decimal", True), c("Service Scope"),
            c("Management LOB", hidden=True), c("Mapping Status"), c("Source File"),
        ),
        (
            m("Forecast Volume", "SUM('Forecast'[Volume Forecast])", "#,##0", "Summed Verint forecast demand at native interval grain.", "Forecast"),
            m("Forecast AHT Seconds", "DIVIDE(SUMX('Forecast', 'Forecast'[Volume Forecast] * 'Forecast'[AHT Forecast Seconds]), [Forecast Volume])", "#,##0", "Volume-weighted forecast AHT at native interval grain.", "Forecast"),
            m(
                "Required FTE",
                "AVERAGEX(SUMMARIZE('Forecast', 'Forecast'[Date], 'Forecast'[Time Slot], \"Interval FTE\", SUM('Forecast'[FTE Required])), [Interval FTE])",
                "#,##0.0", "Average summed required FTE across selected date/time intervals.", "Forecast",
            ),
            m("Requirement Intervals", "COUNT('Forecast'[FTE Required])", "#,##0", "Native forecast intervals containing an explicit Verint required-FTE value.", "Forecast"),
            m("Peak Required FTE", "MAXX(SUMMARIZE('Forecast', 'Forecast'[Date], 'Forecast'[Time Slot], \"Interval FTE\", SUM('Forecast'[FTE Required])), [Interval FTE])", "#,##0.0", "Peak summed explicit required FTE in the selected horizon.", "Forecast"),
        ),
    ),
    Table(
        "Staffing", "FactStaffing15Min.csv", "Scheduled and observed capacity at date, interval, roster LOB and language grain.",
        (
            c("Date", "date", True), c("Time Slot", "int", True), c("LOB"), c("Management LOB", hidden=True),
            c("Language"), c("Scheduled HC", "int", True), c("Observed HC", "int", True),
            c("Productive HC", "int", True), c("Auxiliary HC", "int", True),
            c("Gross Scheduled FTE", "decimal", True), c("Planned Time Off FTE", "decimal", True),
            c("Scheduled FTE", "decimal", True), c("Elapsed Scheduled FTE", "decimal", True),
            c("Observed FTE", "decimal", True),
            c("Productive FTE", "decimal", True), c("Staffing Variance FTE", "decimal", True),
            c("Staffing Gap FTE", "decimal", True), c("Staffing State"), c("Evidence Basis"),
        ),
        (
            m("Average Scheduled FTE", "AVERAGEX(SUMMARIZE('Staffing', 'Staffing'[Date], 'Staffing'[Time Slot], \"Interval FTE\", SUM('Staffing'[Scheduled FTE])), [Interval FTE])", "#,##0.0", "Average net scheduled FTE across selected intervals after governed time off.", "Capacity"),
            m("Peak Net Scheduled FTE", "MAXX(SUMMARIZE('Staffing', 'Staffing'[Date], 'Staffing'[Time Slot], \"Interval FTE\", SUM('Staffing'[Scheduled FTE])), [Interval FTE])", "#,##0.0", "Peak net scheduled FTE after governed PTO/Away in the selected horizon.", "Capacity"),
            m("Average Elapsed Scheduled FTE", "AVERAGEX(SUMMARIZE('Staffing', 'Staffing'[Date], 'Staffing'[Time Slot], \"Interval FTE\", SUM('Staffing'[Elapsed Scheduled FTE])), [Interval FTE])", "#,##0.0", "Average elapsed scheduled FTE in completed interval portions.", "Capacity"),
            m("PTO / Away FTE", "AVERAGEX(SUMMARIZE('Staffing', 'Staffing'[Date], 'Staffing'[Time Slot], \"Interval FTE\", SUM('Staffing'[Planned Time Off FTE])), [Interval FTE])", "#,##0.0", "Average governed PTO/Away FTE deducted from gross scheduled capacity.", "Capacity"),
            m("Average Observed FTE", "AVERAGEX(SUMMARIZE('Staffing', 'Staffing'[Date], 'Staffing'[Time Slot], \"Interval FTE\", SUM('Staffing'[Observed FTE])), [Interval FTE])", "#,##0.0", "Average observed FTE across selected completed intervals.", "Capacity"),
            m("Average Productive FTE", "AVERAGEX(SUMMARIZE('Staffing', 'Staffing'[Date], 'Staffing'[Time Slot], \"Interval FTE\", SUM('Staffing'[Productive FTE])), [Interval FTE])", "#,##0.0", "Average productive FTE across selected completed intervals.", "Capacity"),
            m("Average Scheduled HC", "AVERAGEX(SUMMARIZE('Staffing', 'Staffing'[Date], 'Staffing'[Time Slot], \"Interval HC\", SUM('Staffing'[Scheduled HC])), [Interval HC])", "#,##0.0", "Average scheduled headcount across selected native intervals.", "Capacity"),
            m("Average Observed HC", "AVERAGEX(SUMMARIZE('Staffing', 'Staffing'[Date], 'Staffing'[Time Slot], \"Interval HC\", SUM('Staffing'[Observed HC])), [Interval HC])", "#,##0.0", "Average observed headcount across selected completed intervals.", "Capacity"),
            m("Capacity Gap FTE", "IF(ISBLANK([Required FTE]), BLANK(), MAX([Required FTE] - [Average Scheduled FTE], 0))", "#,##0.0", "Required FTE less average scheduled capacity, floored at zero; blank when Verint requirement is not supplied.", "Capacity"),
            m("Net Capacity Gap FTE", "IF(ISBLANK([Required FTE]), BLANK(), [Average Scheduled FTE] - [Required FTE])", "#,##0.0;[Red]-#,##0.0", "Average scheduled capacity less required FTE; blank when Verint requirement is not supplied.", "Capacity"),
            m("Observed Gap FTE", "MAX([Average Scheduled FTE] - [Average Observed FTE], 0)", "#,##0.0", "Average scheduled FTE less average observed FTE, floored at zero.", "Capacity"),
            m("Coverage %", "IF(ISBLANK([Required FTE]), BLANK(), DIVIDE([Average Scheduled FTE], [Required FTE]))", "0.0%", "Average scheduled capacity divided by explicit Verint required FTE; blank when requirement is unavailable.", "Capacity"),
            m("Peak Shortage FTE", "MINX(SUMMARIZECOLUMNS('Date'[Date], 'Time'[Quarter Hour Index], \"Gap\", IF(ISBLANK([Required FTE]), BLANK(), [Average Scheduled FTE] - [Required FTE])), [Gap])", "#,##0.0;[Red]-#,##0.0", "Most negative scheduled-minus-required FTE interval; blank requirements remain excluded.", "Capacity"),
            m("Critical Intervals", "COUNTROWS(FILTER(SUMMARIZECOLUMNS('Date'[Date], 'Time'[Quarter Hour Index], \"Gap\", IF(ISBLANK([Required FTE]), BLANK(), [Average Scheduled FTE] - [Required FTE])), NOT ISBLANK([Gap]) && [Gap] < -1))", "#,##0", "Selected intervals with more than one FTE scheduled shortage.", "Capacity"),
            m("Current Productive Gap FTE", "IF(ISBLANK([Required FTE]), BLANK(), [Average Productive FTE] - [Required FTE])", "#,##0.0;[Red]-#,##0.0", "Productive capacity less explicit requirement in the selected operational context.", "Capacity"),
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
            m("Present HC", "CALCULATE(DISTINCTCOUNT('Attendance'[Agent Day Key]), 'Attendance'[Planned Work Minutes] > 0, 'Attendance'[Attendance Result] IN {\"Present\", \"Present - partial time off\", \"Late\", \"Early leave\", \"Late + early leave\", \"Shift in progress\", \"Late - shift in progress\"})", "#,##0", "Due agent-days with Agent Status-first observed presence; late arrivals and early leavers remain present.", "Attendance"),
            m("No Show HC", "CALCULATE(DISTINCTCOUNT('Attendance'[Agent Day Key]), 'Attendance'[Attendance Result] IN {\"No show\", \"No show - partial time off\"})", "#,##0", "Evidence-backed completed no-show agent-days only.", "Attendance"),
            m("Unknown / Possible No Show HC", "MAX([Due HC] - [Present HC] - [No Show HC], 0)", "#,##0", "Due population not yet proven present or no-show; never silently promoted to no-show.", "Attendance"),
            m("Callout HC", "CALCULATE(DISTINCTCOUNT('Attendance'[Agent Day Key]), 'Attendance'[Call Action] <> \"NONE\")", "#,##0", "Attendance cases requiring an operational call or follow-up.", "Attendance"),
            m("No Show Calls Open", "CALCULATE(DISTINCTCOUNT('Attendance'[Agent Day Key]), 'Attendance'[Call Action] = \"CALL_NO_SHOW\")", "#,##0", "Confirmed no-show cases still surfaced for operational callout.", "Attendance"),
            m("Late HC", "CALCULATE(DISTINCTCOUNT('Attendance'[Agent Day Key]), 'Attendance'[Late Minutes] > 0)", "#,##0", "Agent-days with governed late minutes.", "Attendance"),
            m("Early Leave HC", "CALCULATE(DISTINCTCOUNT('Attendance'[Agent Day Key]), 'Attendance'[Early Leave Minutes] > 0)", "#,##0", "Completed agent-days with governed early-leave minutes.", "Attendance"),
            m("On Time HC", "MAX([Present HC] - [Late HC], 0)", "#,##0", "Present agent-days without governed late minutes.", "Attendance"),
            m("Attendance %", "DIVIDE([Present HC], [Due HC])", "0.0%", "Present agent-days divided by due agent-days.", "Attendance"),
            m("Status Evidence Coverage %", "DIVIDE(SUM('Attendance'[Status Covered Minutes]), SUM('Attendance'[Planned Work Minutes]))", "0.0%", "Agent Status-covered minutes divided by governed planned work minutes.", "Data quality"),
        ),
    ),
    Table(
        "Status", "FactStatusInterval.csv", "Agent Status-first governed shift segments with reviewed AUX classification.",
        (
            c("Date", "date", True), c("Segment Key", hidden=True), c("Agent Day Key", hidden=True),
            c("Agent ID", hidden=True), c("Agent"), c("Team Leader"), c("Ops Manager"),
            c("LOB"), c("Management LOB", hidden=True), c("Language"), c("Scheduled Start"),
            c("Scheduled End"), c("Segment Start"), c("Segment End"), c("Time Slot", "int", True),
            c("Minutes", "int", True), c("Planned State"), c("Actual Status"),
            c("Attendance Category"), c("AUX Classification"), c("Qualification 1"),
            c("Qualification 2"), c("Operational Category"), c("Mismatch"), c("Is Gap", "int"),
            c("Observed Source"), c("Shift State"), c("Is Planned Time Off", "int", True),
            c("Is Observed", "int", True), c("Is Productive", "int", True),
            c("Is Unexplained", "int", True), c("Is Status Mapped", "int", True),
            c("Evaluation As Of"),
        ),
        (
            m("Elapsed Scheduled Minutes", "CALCULATE(SUM('Status'[Minutes]), 'Status'[Shift State] = \"COMPLETE\", 'Status'[Is Planned Time Off] = 0)", "#,##0", "Completed scheduled minutes excluding governed PTO/Away.", "Workforce realisation"),
            m("Observed Minutes", "CALCULATE(SUM('Status'[Minutes]), 'Status'[Shift State] = \"COMPLETE\", 'Status'[Is Planned Time Off] = 0, 'Status'[Is Observed] = 1)", "#,##0", "Completed elapsed minutes with observed connected evidence.", "Workforce realisation"),
            m("Productive Minutes", "CALCULATE(SUM('Status'[Minutes]), 'Status'[Shift State] = \"COMPLETE\", 'Status'[Is Planned Time Off] = 0, 'Status'[Is Productive] = 1)", "#,##0", "Completed voice, available, BO and other governed productive minutes.", "Workforce realisation"),
            m("Unexplained Minutes", "CALCULATE(SUM('Status'[Minutes]), 'Status'[Shift State] = \"COMPLETE\", 'Status'[Is Planned Time Off] = 0, 'Status'[Is Unexplained] = 1)", "#,##0", "Completed scheduled minutes with logged-off, no-activity or missing-status evidence.", "Workforce realisation"),
            m("Elapsed Scheduled Hours", "DIVIDE([Elapsed Scheduled Minutes], 60)", "#,##0.0", "Completed elapsed scheduled hours excluding governed PTO/Away.", "Workforce realisation"),
            m("Observed Hours", "DIVIDE([Observed Minutes], 60)", "#,##0.0", "Observed connected hours in completed shifts.", "Workforce realisation"),
            m("Productive Hours", "DIVIDE([Productive Minutes], 60)", "#,##0.0", "Governed productive hours in completed shifts.", "Workforce realisation"),
            m("Unexplained Hours", "DIVIDE([Unexplained Minutes], 60)", "#,##0.0", "Unexplained scheduled hours requiring evidence review.", "Workforce realisation"),
            m("Presence Realisation %", "DIVIDE([Observed Minutes], [Elapsed Scheduled Minutes])", "0.0%", "Observed minutes divided by elapsed scheduled minutes.", "Workforce realisation"),
            m("Productive Realisation %", "DIVIDE([Productive Minutes], [Observed Minutes])", "0.0%", "Governed productive minutes divided by observed minutes.", "Workforce realisation"),
            m("Capacity Delivered %", "DIVIDE([Productive Minutes], [Elapsed Scheduled Minutes])", "0.0%", "Governed productive minutes divided by elapsed scheduled minutes.", "Workforce realisation"),
            m("Status Hours", "DIVIDE(SUM('Status'[Minutes]), 60)", "#,##0.0", "Hours in the selected governed operational category.", "Workforce realisation"),
            m("Status Mapped %", "DIVIDE(CALCULATE(SUM('Status'[Minutes]), 'Status'[Observed Source] = \"AGENT_STATUS\", 'Status'[Is Status Mapped] = 1), CALCULATE(SUM('Status'[Minutes]), 'Status'[Observed Source] = \"AGENT_STATUS\"))", "0.0%", "Mapped Agent Status minutes divided by Agent Status minutes.", "Data quality"),
            m("Capacity Loss Hours", "MAX([Elapsed Scheduled Hours] - [Productive Hours], 0)", "#,##0.0", "Elapsed scheduled hours not delivered as governed productive hours.", "Driver signals"),
            m("Capacity Bridge Hours", "SWITCH(SELECTEDVALUE('Capacity Stage'[Capacity Stage]), \"Elapsed scheduled\", [Elapsed Scheduled Hours], \"Absence / missing\", -MAX([Elapsed Scheduled Hours]-[Observed Hours],0), \"Observed\", [Observed Hours], \"AUX / breaks\", -MAX([Observed Hours]-[Productive Hours],0), \"Productive\", [Productive Hours])", "#,##0.0;[Red]-#,##0.0", "Auditable capacity bridge values by selected stage.", "Workforce realisation"),
        ),
    ),
    Table(
        "Schedule Integrity", "FactScheduleIntegrity.csv", "Completed published-versus-observed shift placement and supported recurrence evidence.",
        (
            c("Date", "date", True), c("Integrity Key", hidden=True), c("Agent Day Key", hidden=True),
            c("Agent ID", hidden=True), c("Agent"), c("Team Leader"), c("Ops Manager"),
            c("LOB"), c("Management LOB", hidden=True), c("Language"), c("Scheduled Start"),
            c("Scheduled End"), c("Observed Start"), c("Observed End"),
            c("Scheduled Minutes", "int", True), c("Observed Span Minutes", "int", True),
            c("Start Delta Minutes", "int"), c("End Delta Minutes", "int"),
            c("Displaced Minutes", "int", True), c("Internal Gap Minutes", "int"),
            c("Internal Gap Count", "int"), c("Classification"), c("Pattern Family"),
            c("Recurrence Count", "int"), c("Eligible Day Count", "int"),
            c("Is Recurring", "int"), c("Requires Review", "int"), c("Evidence Basis"), c("Confidence"),
            c("Evaluation As Of"),
        ),
        (
            m("Displaced Shifts", "CALCULATE(COUNTROWS('Schedule Integrity'), 'Schedule Integrity'[Requires Review] = 1)", "#,##0", "Completed supported shift-placement cases requiring review.", "Schedule integrity"),
            m("Capacity Displaced Hours", "DIVIDE(CALCULATE(SUM('Schedule Integrity'[Displaced Minutes]), 'Schedule Integrity'[Requires Review] = 1), 60)", "#,##0.0", "Supported displaced or lost boundary minutes expressed as hours.", "Schedule integrity"),
            m("Recurring Agents", "CALCULATE(DISTINCTCOUNT('Schedule Integrity'[Agent ID]), 'Schedule Integrity'[Is Recurring] = 1)", "#,##0", "Agents reaching the configured supported recurrence threshold.", "Schedule integrity"),
            m("Early Leaves", "CALCULATE(COUNTROWS('Schedule Integrity'), 'Schedule Integrity'[Classification] IN {\"Early leave\", \"Late start + early leave\"})", "#,##0", "Completed supported early-leave cases.", "Schedule integrity"),
            m("Integrity Cases", "CALCULATE(COUNTROWS('Schedule Integrity'), 'Schedule Integrity'[Requires Review] = 1)", "#,##0", "Supported schedule-integrity rows requiring human review.", "Schedule integrity"),
        ),
    ),
    Table(
        "Shift Placement", "FactShiftPlacement.csv", "Two-row published/observed placement bridge used by the native Gantt-style chart.",
        (
            c("Date", "date", True), c("Placement Key", hidden=True), c("Agent ID", hidden=True),
            c("Agent"), c("Team Leader"), c("Management LOB", hidden=True), c("Placement"),
            c("Placement Label"), c("Start Hour", "decimal", True), c("Duration Hours", "decimal", True),
            c("Published Hours", "decimal", True), c("Observed Hours", "decimal", True),
            c("Classification"), c("Requires Review", "int"),
        ),
        (
            m("Start Hour Value", "SUM('Shift Placement'[Start Hour])", "0.0", "Hours after midnight before the placement bar starts.", "Schedule integrity"),
            m("Published Placement Hours", "SUM('Shift Placement'[Published Hours])", "0.0", "Published shift duration for the placement chart.", "Schedule integrity"),
            m("Observed Placement Hours", "SUM('Shift Placement'[Observed Hours])", "0.0", "Observed presence span for the placement chart.", "Schedule integrity"),
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
            m("Sources Ready", "CALCULATE(COUNTROWS('Source Health'), 'Source Health'[Status] IN {\"SUCCESS\", \"OK\", \"READY\", \"LOADED\"})", "#,##0", "Source families reporting a successful or ready status.", "Data quality"),
            m("Sources Ready %", "DIVIDE([Sources Ready], [Source Count])", "0.0%", "Ready source families divided by configured source families.", "Data quality"),
            m("Latest Data Date", "MAX('Source Health'[Newest Date])", "dd mmm yyyy", "Latest business date present across source families.", "Data quality"),
            m("Rows Accepted %", "DIVIDE(SUM('Source Health'[Rows]), SUM('Source Health'[Rows]) + SUM('Source Health'[Rejected]))", "0.0%", "Accepted source rows divided by accepted plus rejected rows.", "Data quality"),
            m("Models Built", "22", "#,##0", "Imported semantic tables in Power BI project contract 5.", "Data quality"),
            m("Manifest State", "IF([Critical Quality Issues] > 0, \"HOLD\", \"READY\")", "", "HOLD when blocking quality issues exist; READY otherwise.", "Data quality"),
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
    ("Employee", "Management LOB", "Management LOB", "Management LOB"),
    ("Queue", "Management LOB", "Management LOB", "Management LOB"),
    ("Service", "Date", "Date", "Date"), ("Service", "Time Slot", "Time", "Quarter Hour Index"),
    ("Service", "Service Key", "Queue", "Service Key"),
    ("Queue Coverage", "Date", "Date", "Date"),
    ("Queue Coverage", "Service Key", "Queue", "Service Key"),
    ("Forecast", "Date", "Date", "Date"), ("Forecast", "Time Slot", "Time", "Quarter Hour Index"),
    ("Forecast", "Service Key", "Queue", "Service Key"),
    ("Staffing", "Date", "Date", "Date"), ("Staffing", "Time Slot", "Time", "Quarter Hour Index"),
    ("Staffing", "Management LOB", "Management LOB", "Management LOB"),
    ("Attendance", "Date", "Date", "Date"), ("Attendance", "Agent ID", "Employee", "Agent ID"),
    ("Status", "Date", "Date", "Date"), ("Status", "Agent ID", "Employee", "Agent ID"),
    ("Status", "Time Slot", "Time", "Quarter Hour Index"),
    ("Schedule Integrity", "Date", "Date", "Date"),
    ("Schedule Integrity", "Agent ID", "Employee", "Agent ID"),
    ("Shift Placement", "Date", "Date", "Date"),
    ("Shift Placement", "Agent ID", "Employee", "Agent ID"),
    ("Attendance Gap", "Date", "Date", "Date"), ("Attendance Gap", "Agent ID", "Employee", "Agent ID"),
    ("Time Off", "Date", "Date", "Date"), ("Time Off", "Agent ID", "Employee", "Agent ID"),
    ("Final Absence", "Date", "Date", "Date"), ("Final Absence", "Agent ID", "Employee", "Agent ID"),
    ("Absence Component", "Date", "Date", "Date"), ("Absence Component", "Agent ID", "Employee", "Agent ID"),
    ("Finding", "Period End", "Date", "Date"),
    ("Finding", "Management LOB", "Management LOB", "Management LOB"),
    ("Quality Issue", "Date", "Date", "Date"), ("Quality Issue", "Agent ID", "Employee", "Agent ID"),
)


PAGES = (
    {
        "title": "Daily WFM Command", "nav": "Daily Command", "status": "LIVE",
        "subtitle": "Today’s service, attendance and capacity decisions in one operational view",
        "slicers": (("Date", "Date", "DATE", "Between"), ("Management LOB", "Management LOB", "MANAGEMENT LOB", "Dropdown"), ("Employee", "Team Leader", "TEAM LEADER", "Dropdown"), ("Time", "Time Label", "CHECKPOINT", "Dropdown")),
        "cards": (
            ("Service", "LOBs On Target Label", "LOBS ON TARGET", "#159957", "Profile-specific target"),
            ("Service", "Volume Variance %", "DEMAND VS FORECAST", "#D99815", "Actual entered vs forecast"),
            ("Attendance", "No Show HC", "CONFIRMED NO SHOW", "#C91F2A", "Evidence-backed completed cases"),
            ("Staffing", "Current Productive Gap FTE", "CURRENT CAPACITY GAP", "#C91F2A", "Productive FTE vs required"),
        ),
        "charts": (
            {"type": "clusteredBarChart", "title": "SERVICE LEVEL BY MANAGEMENT LOB", "category": ("Management LOB", "Management LOB"), "values": (("Service", "Service Level %"), ("Service", "SL Target %"))},
            {"type": "lineChart", "title": "TODAY'S CAPACITY PULSE", "category": ("Time", "Time Label"), "values": (("Forecast", "Required FTE"), ("Staffing", "Average Observed FTE"), ("Staffing", "Average Productive FTE"))},
        ),
        "tables": (("WFM ACTION QUEUE", (("Finding", "Rank"), ("Finding", "Severity"), ("Finding", "Management LOB"), ("Finding", "Title"), ("Finding", "Summary"), ("Finding", "Evidence Dataset")), "full"),),
    },
    {
        "title": "Service & SL Drivers", "nav": "SL Drivers", "status": "ANALYSIS",
        "subtitle": "15-minute service pressure, review leads and exact queue evidence",
        "slicers": (("Date", "Date", "DATE", "Between"), ("Management LOB", "Management LOB", "MANAGEMENT LOB", "Dropdown"), ("Queue", "Queue", "QUEUE", "Dropdown"), ("Time", "Time Label", "15-MIN INTERVAL", "Dropdown")),
        "cards": (
            ("Service", "Service Level %", "SERVICE LEVEL", "#C91F2A", "Against configured target"),
            ("Service", "Volume Variance %", "DEMAND VARIANCE", "#D99815", "Actual entered vs forecast"),
            ("Status", "Capacity Loss Hours", "CAPACITY LOSS", "#C91F2A", "Elapsed schedule not productive"),
            ("Service", "Primary Driver", "PRIMARY DRIVER", "#244F78", "Largest supported pressure signal"),
        ),
        "charts": (
            {"type": "lineStackedColumnComboChart", "title": "INTRADAY SL AND DEMAND PRESSURE", "category": ("Time", "Time Label"), "values": (("Service", "Offered Calls"), ("Forecast", "Forecast Volume")), "secondary": (("Service", "Service Level %"), ("Service", "SL Target %"))},
            {"type": "clusteredColumnChart", "title": "SUPPORTED DRIVER PRESSURE SIGNALS", "category": ("Driver", "Driver"), "values": (("Service", "Driver Pressure %"),)},
        ),
        "tables": (("INTERVAL & QUEUE DIAGNOSIS", (("Service", "Interval Start"), ("Queue", "Queue"), ("Queue", "Designation"), ("Service", "Offered Calls"), ("Forecast", "Forecast Volume"), ("Service", "Service Level %"), ("Service", "AHT Seconds"), ("Service", "AHT Error Seconds")), "full"),),
    },
    {
        "title": "Staff Preparation", "nav": "Staff Prep", "status": "PLANNING",
        "subtitle": "Future requirement, net published capacity and PTO/Away impact at native grain",
        "slicers": (("Date", "Date", "PLANNING HORIZON", "Between"), ("Management LOB", "Management LOB", "MANAGEMENT LOB", "Dropdown"), ("Date", "Weekday", "DAY OF WEEK", "Dropdown"), ("Time", "Time Label", "OPERATING WINDOW", "Dropdown")),
        "cards": (
            ("Forecast", "Peak Required FTE", "PEAK REQUIRED FTE", "#244F78", "Explicit Verint requirement"),
            ("Staffing", "Peak Net Scheduled FTE", "NET SCHEDULED FTE", "#159957", "After PTO and Away"),
            ("Staffing", "Peak Shortage FTE", "PEAK SHORTAGE", "#C91F2A", "Worst supported interval"),
            ("Staffing", "Critical Intervals", "CRITICAL INTERVALS", "#D99815", "Shortage greater than 1 FTE"),
        ),
        "charts": (
            {"type": "lineChart", "title": "REQUIRED VS NET SCHEDULED CAPACITY", "category": ("Time", "Time Label"), "values": (("Forecast", "Required FTE"), ("Staffing", "Average Scheduled FTE"))},
            {"type": "matrix", "title": "COVERAGE RISK BY LOB & DAY", "rows": (("Management LOB", "Management LOB"),), "columns": (("Date", "Date"),), "values": (("Staffing", "Coverage %"),)},
        ),
        "tables": (("STAFF PREPARATION ACTIONS", (("Date", "Date"), ("Time", "Time Label"), ("Management LOB", "Management LOB"), ("Forecast", "Required FTE"), ("Staffing", "Average Scheduled FTE"), ("Staffing", "Net Capacity Gap FTE"), ("Staffing", "Coverage %"), ("Staffing", "PTO / Away FTE")), "full"),),
    },
    {
        "title": "Workforce Realisation", "nav": "Realisation", "status": "CONTROL",
        "subtitle": "Observed and productive delivery against completed elapsed schedule",
        "slicers": (("Date", "Date", "PERIOD", "Between"), ("Management LOB", "Management LOB", "MANAGEMENT LOB", "Dropdown"), ("Employee", "Team Leader", "TEAM LEADER", "Dropdown"), ("Date", "Weekday", "DAY TYPE", "Dropdown")),
        "cards": (
            ("Status", "Presence Realisation %", "PRESENCE REALISATION", "#159957", "Observed / elapsed scheduled"),
            ("Status", "Productive Realisation %", "PRODUCTIVE REALISATION", "#087E91", "Productive / observed"),
            ("Status", "Capacity Delivered %", "CAPACITY DELIVERED", "#244F78", "Productive / elapsed scheduled"),
            ("Status", "Unexplained Hours", "UNEXPLAINED TIME", "#C91F2A", "Requires evidence review"),
        ),
        "charts": (
            {"type": "hundredPercentStackedBarChart", "title": "OBSERVED TIME COMPOSITION BY LOB", "category": ("Management LOB", "Management LOB"), "values": (("Status", "Status Hours"),), "series": ("Status", "Operational Category")},
            {"type": "clusteredColumnChart", "title": "CAPACITY REALISATION BRIDGE", "category": ("Capacity Stage", "Capacity Stage"), "values": (("Status", "Capacity Bridge Hours"),)},
        ),
        "tables": (("LOB REALISATION DETAIL", (("Management LOB", "Management LOB"), ("Status", "Elapsed Scheduled Hours"), ("Status", "Observed Hours"), ("Status", "Productive Hours"), ("Status", "Presence Realisation %"), ("Status", "Productive Realisation %"), ("Status", "Capacity Delivered %"), ("Status", "Unexplained Hours")), "full"),),
    },
    {
        "title": "Schedule Integrity & Patterns", "nav": "Schedule Integrity", "status": "REVIEW",
        "subtitle": "Published versus observed placement with evidence-backed recurrence only",
        "slicers": (("Date", "Date", "PERIOD", "Between"), ("Management LOB", "Management LOB", "MANAGEMENT LOB", "Dropdown"), ("Employee", "Team Leader", "TEAM LEADER", "Dropdown"), ("Schedule Integrity", "Pattern Family", "PATTERN", "Dropdown")),
        "cards": (
            ("Schedule Integrity", "Displaced Shifts", "DISPLACED SHIFTS", "#C91F2A", "Completed supported cases"),
            ("Schedule Integrity", "Capacity Displaced Hours", "CAPACITY DISPLACED", "#D99815", "Hours moved or lost"),
            ("Schedule Integrity", "Recurring Agents", "RECURRING AGENTS", "#C91F2A", "Configured recurrence threshold"),
            ("Schedule Integrity", "Early Leaves", "EARLY LEAVES", "#D99815", "Completed days only"),
        ),
        "charts": (
            {"type": "clusteredBarChart", "title": "PUBLISHED VS OBSERVED SHIFT HOURS", "category": ("Shift Placement", "Placement Label"), "values": (("Shift Placement", "Published Placement Hours"), ("Shift Placement", "Observed Placement Hours"))},
            {"type": "matrix", "title": "SUPPORTED RECURRING PATTERNS", "rows": (("Schedule Integrity", "Pattern Family"),), "columns": (("Date", "Weekday"),), "values": (("Schedule Integrity", "Integrity Cases"),)},
        ),
        "tables": (("SCHEDULE INTEGRITY CASES", (("Schedule Integrity", "Agent"), ("Schedule Integrity", "Team Leader"), ("Schedule Integrity", "Date"), ("Schedule Integrity", "Scheduled Start"), ("Schedule Integrity", "Scheduled End"), ("Schedule Integrity", "Observed Start"), ("Schedule Integrity", "Observed End"), ("Schedule Integrity", "Classification"), ("Schedule Integrity", "Start Delta Minutes"), ("Schedule Integrity", "End Delta Minutes"), ("Schedule Integrity", "Recurrence Count"), ("Schedule Integrity", "Confidence")), "full"),),
    },
    {
        "title": "Forecast Accuracy", "nav": "Forecast Accuracy", "status": "PLANNING",
        "subtitle": "Comparable 15-minute demand and AHT forecast quality, separate from execution",
        "slicers": (("Date", "Date", "PERIOD", "Between"), ("Management LOB", "Management LOB", "MANAGEMENT LOB", "Dropdown"), ("Queue", "Queue", "QUEUE", "Dropdown"), ("Date", "Weekday", "DAY OF WEEK", "Dropdown")),
        "cards": (
            ("Service", "Forecast Accuracy %", "VOLUME ACCURACY", "#159957", "One minus weighted absolute error"),
            ("Service", "Forecast Bias %", "FORECAST BIAS", "#D99815", "Signed actual vs forecast"),
            ("Service", "AHT Error Seconds", "AHT ERROR", "#D99815", "Actual vs weighted forecast"),
            ("Service", "Peak Accuracy %", "PEAK ACCURACY", "#C91F2A", "Top forecast-demand decile"),
        ),
        "charts": (
            {"type": "lineChart", "title": "ACTUAL VS FORECAST DEMAND", "category": ("Time", "Time Label"), "values": (("Service", "Offered Calls"), ("Forecast", "Forecast Volume"))},
            {"type": "matrix", "title": "ACCURACY BY LOB & DAY OF WEEK", "rows": (("Management LOB", "Management LOB"),), "columns": (("Date", "Weekday"),), "values": (("Service", "Forecast Accuracy %"),)},
        ),
        "tables": (("RECURRING FORECAST MISSES", (("Date", "Weekday"), ("Management LOB", "Management LOB"), ("Time", "Hour Label"), ("Forecast", "Forecast Volume"), ("Service", "Offered Calls"), ("Service", "Forecast Bias %"), ("Service", "AHT Error Seconds"), ("Service", "Forecast Accuracy %")), "full"),),
    },
    {
        "title": "Absence & Shrinkage", "nav": "Absence & Shrinkage", "status": "FINAL CHECK",
        "subtitle": "Final Verint activity ledger, governed component logic and unresolved evidence",
        "slicers": (("Date", "Date", "PERIOD", "Between"), ("Management LOB", "Management LOB", "MANAGEMENT LOB", "Dropdown"), ("Employee", "Team Leader", "TEAM LEADER", "Dropdown"), ("Final Absence", "Final Ledger Status", "LEDGER STATUS", "Dropdown")),
        "cards": (
            ("Final Absence", "Final Absence %", "FINAL ABSENCE", "#C91F2A", "Mapped absence / finalized planned time"),
            ("Final Absence", "Final Shrinkage %", "FINAL SHRINKAGE", "#D99815", "Mapped shrinkage / finalized planned time"),
            ("Final Absence", "Final PTO %", "FINAL PTO", "#244F78", "Mapped vacation / finalized planned time"),
            ("Final Absence", "Absence Review HC", "REVIEW AGENT-DAYS", "#C91F2A", "Residual, empty or unmapped evidence"),
        ),
        "charts": (
            {"type": "clusteredBarChart", "title": "FINAL RATES BY MANAGEMENT LOB", "category": ("Management LOB", "Management LOB"), "values": (("Final Absence", "Final Absence %"), ("Final Absence", "Final Shrinkage %"))},
            {"type": "clusteredColumnChart", "title": "VERINT ACTIVITY COMPONENT MINUTES", "category": ("Absence Component", "Category"), "values": (("Absence Component", "Activity Minutes"),)},
        ),
        "tables": (("FINAL AGENT-DAY DETAIL", (("Final Absence", "Date"), ("Final Absence", "Agent"), ("Final Absence", "Team Leader"), ("Final Absence", "LOB"), ("Final Absence", "Scheduled Minutes"), ("Final Absence", "Absence Minutes"), ("Final Absence", "Vacation Minutes"), ("Final Absence", "Shrinkage Minutes"), ("Final Absence", "Unmapped Minutes"), ("Final Absence", "Final Ledger Status")), "full"),),
    },
    {
        "title": "Data Readiness", "nav": "Data Readiness", "status": "GOVERNED",
        "subtitle": "Freshness, mappings, accepted evidence and active operational blockers",
        "slicers": (("Date", "Date", "PERIOD", "Between"), ("Quality Issue", "Source Family", "SOURCE FAMILY", "Dropdown"), ("Quality Issue", "Severity", "SEVERITY", "Dropdown"), ("Finding", "Domain", "DOMAIN", "Dropdown")),
        "cards": (
            ("Source Health", "Sources Ready %", "SOURCES FRESH", "#D99815", "Latest complete Hub update"),
            ("Queue Coverage", "Queue Mapped %", "QUEUE MAPPED", "#159957", "By inbound queue entries"),
            ("Status", "Status Mapped %", "STATUS MAPPED", "#D99815", "By observed Agent Status minutes"),
            ("Quality Issue", "Critical Quality Issues", "BLOCKING ISSUES", "#C91F2A", "Never converted to zero"),
        ),
        "charts": (
            {"type": "table", "title": "SOURCE FRESHNESS", "fields": (("Source Health", "Source Family"), ("Source Health", "Newest Date"), ("Source Health", "Rows"), ("Source Health", "Rejected"), ("Source Health", "Status"))},
            {"type": "pipeline", "title": "GOVERNED UPDATE PIPELINE", "stages": (("Source Health", "Source Count", "Source contracts", "#087E91"), ("Source Health", "Rows Accepted %", "Rows accepted", "#087E91"), ("Source Health", "Models Built", "Models built", "#087E91"), ("Quality Issue", "Critical Quality Issues", "Blocking checks", "#C91F2A"), ("Source Health", "Manifest State", "Manifest state", "#244F78"))},
        ),
        "tables": (("ACTIVE DATA ISSUES", (("Quality Issue", "Severity"), ("Quality Issue", "Source Family"), ("Quality Issue", "Issue Type"), ("Quality Issue", "Date"), ("Quality Issue", "Agent ID"), ("Quality Issue", "Details")), "full"),),
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
                *([f"\t\tformatString: {measure.format_string}"] if measure.format_string else []),
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
    name: str, table: str, measure: str, label: str, accent: str,
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
                    "fontSize": _literal(30), "fontColor": _color(accent),
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
        "position": _position(18, 128, 140, len(PAGES) * 48, 20, 20),
        "visual": {
            "visualType": "pageNavigator",
            "objects": {
                "layout": [{"properties": {"columnCount": _literal(1), "rowCount": _literal(len(PAGES)), "cellPadding": _literal(6)}}],
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
                "$schema": PAGE_SCHEMA, "name": page_id, "displayName": page["nav"],
                "displayOption": "FitToPage", "height": 945, "width": 1680,
                "objects": {
                    "background": [{"properties": {"color": _color("#F3F6F8"), "transparency": _literal(0)}}],
                    "outspace": [{"properties": {"color": _color("#E8EEF3"), "transparency": _literal(0)}}],
                },
            },
        )
        _write_visual(page_dir, title, "sidebar", _shape("", 0, 0, 178, 945, "#0B1F33", 0))
        _write_visual(page_dir, title, "header", _shape("", 178, 0, 1502, 58, "#0B1F33", 1))
        _write_visual(page_dir, title, "accent", _shape("", 178, 58, 1502, 4, "#00A3A8", 2))
        _write_visual(page_dir, title, "brand", _textbox("", "WFMHub", 20, 17, 140, 48, 22, "#FFFFFF", 10, True, "center"))
        _write_visual(page_dir, title, "nav label", _textbox("", "NAVIGATION", 26, 100, 124, 20, 9, "#7F9AB2", 10, True))
        # Static labels guarantee a readable sidebar even on Desktop builds
        # that fail to render pageNavigator text.  The native navigator stays
        # above them as the interactive layer; Power BI's bottom page tabs are
        # also retained as a second navigation route.
        current_nav = next(
            index for index, candidate in enumerate(PAGES)
            if candidate["title"] == title
        )
        for index, candidate in enumerate(PAGES):
            nav_y = 132 + index * 48
            if index == current_nav:
                _write_visual(
                    page_dir, title, f"nav selected {index}",
                    _shape("", 20, nav_y, 136, 38, "#007C83", 12, rounded=True),
                )
            _write_visual(
                page_dir, title, f"nav text {index}",
                _textbox(
                    "", f"{index + 1:02d}  {candidate['nav']}",
                    28, nav_y + 9, 124, 20, 9,
                    "#FFFFFF" if index == current_nav else "#DCE9F2",
                    14, index == current_nav,
                ),
            )
        _write_visual(page_dir, title, "nav", _navigator(""))
        _write_visual(page_dir, title, "owner", _textbox("", "Prepared by Anass ASSRI\nWorkforce Management", 18, 878, 140, 42, 9, "#9FB3C8", 10, False, "center"))
        _write_visual(page_dir, title, "page title", _textbox("", f"WFM HUB  |  {title.upper()}", 210, 10, 760, 32, 18, "#FFFFFF", 10, True))
        _write_visual(page_dir, title, "subtitle", _textbox("", page["subtitle"], 820, 16, 680, 24, 10, "#C9D7E3", 10, False, "right"))
        _write_visual(page_dir, title, "status panel", _shape("", 1520, 13, 128, 32, "#007C83", 8, rounded=True))
        _write_visual(page_dir, title, "status", _textbox("", page["status"], 1525, 20, 118, 18, 9, "#FFFFFF", 10, True, "center"))

        # One compact selector strip shared by every page.
        _write_visual(page_dir, title, "selector panel", _shape("", 196, 70, 1456, 70, "#FFFFFF", 3, rounded=True))
        slicer_x = (212, 570, 928, 1286)
        for index, ((table, field, label, mode), x) in enumerate(zip(page["slicers"], slicer_x), 1):
            _write_visual(page_dir, title, f"slicer {index}", _slicer("", table, field, label, mode, x, 72, 340, 66))

        # Four aligned KPI cards with a restrained semantic accent.
        card_x = (196, 562, 928, 1294)
        for index, ((table, measure, label, accent, subtext), x) in enumerate(zip(page["cards"], card_x), 1):
            _write_visual(page_dir, title, f"card panel {index}", _shape("", x, 152, 350, 147, "#FFFFFF", 3, rounded=True))
            _write_visual(page_dir, title, f"card accent {index}", _shape("", x, 152, 8, 147, accent, 4, rounded=True))
            _write_visual(page_dir, title, f"card {index}", _card("", table, measure, label, accent, x + 8, 154, 338, 108))
            _write_visual(page_dir, title, f"card note {index}", _textbox("", subtext, x + 28, 268, 300, 20, 9, "#61758A", 10, False))

        # The two analytical panels occupy the same visual rhythm on all pages.
        chart_positions = ((196, 311, 716, 311), (928, 311, 724, 311))
        for index, (chart, position) in enumerate(zip(page["charts"], chart_positions), 1):
            x, y, width, height = position
            _write_visual(page_dir, title, f"chart panel {index}", _shape("", x, y, width, height, "#FFFFFF", 3, rounded=True))
            if chart["type"] == "pipeline":
                _write_visual(
                    page_dir, title, f"pipeline title {index}",
                    _textbox("", chart["title"], x + 22, y + 13, width - 44, 24, 11, "#17324D", 10, True),
                )
                _write_visual(
                    page_dir, title, f"pipeline caption {index}",
                    _textbox("", "The manifest is published only after the governed feed completes", x + 22, y + 38, width - 44, 18, 9, "#61758A", 10),
                )
                node_width = 116
                node_gap = 20
                node_start = x + 22
                for node_index, (table, measure, label, accent) in enumerate(chart["stages"], 1):
                    node_x = node_start + (node_index - 1) * (node_width + node_gap)
                    _write_visual(
                        page_dir, title, f"pipeline node panel {index}-{node_index}",
                        _shape("", node_x, y + 78, node_width, 142, "#FAFCFD", 2, rounded=True),
                    )
                    _write_visual(
                        page_dir, title, f"pipeline node {index}-{node_index}",
                        _card("", table, measure, label, accent, node_x + 4, y + 82, node_width - 8, 132),
                    )
                    if node_index < len(chart["stages"]):
                        _write_visual(
                            page_dir, title, f"pipeline arrow {index}-{node_index}",
                            _textbox("", "→", node_x + node_width, y + 132, node_gap, 28, 16, "#95A8B6", 10, True, "center"),
                        )
                continue
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
            "full": (196, 634, 1456, 250),
            "left": (196, 634, 716, 250),
            "right": (928, 634, 724, 250),
        }
        for index, (table_title, table_fields, layout) in enumerate(page["tables"], 1):
            x, y, width, height = table_layouts[layout]
            _write_visual(page_dir, title, f"table panel {index}", _shape("", x, y, width, height, "#FFFFFF", 3, rounded=True))
            _write_visual(page_dir, title, f"table {index}", _table("", table_title, table_fields, x + 8, y + 6, width - 16, height - 12))

        _write_visual(page_dir, title, "footer", _textbox("", "Prepared by Anass ASSRI | WFM   •   Internal operational use   •   Governed feed: POWERBI_MANIFEST_CURRENT.csv", 196, 902, 1456, 20, 8, "#64748B", 10, False, "right"))

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
    (PROJECT_ROOT / "PROJECT_VERSION.txt").write_text("5\n", encoding="utf-8")
    (PROJECT_ROOT / "README.txt").write_text(
        "WFMHUB BI\n=========\n\n"
        "Open WFMHub BI.pbip with Microsoft Power BI Desktop, then choose Home > Refresh.\n"
        "The HubRoot parameter is set automatically when this project is opened through POWERBI.cmd.\n"
        "The model reads only Feed\\PowerBI CSVs and never reads SQLite or raw extracts.\n"
        "PCS remains a separate collaborative Excel product and is not imported into this WFM-only model.\n",
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
