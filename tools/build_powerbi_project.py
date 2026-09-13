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
import tempfile
import uuid
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
        "Planning Group", "DimPlanningGroup.csv", "Capacity-planning roll-up below Management LOB; RSA BE FR and VL remain separate here.",
        (c("Planning Group"), c("Management LOB", hidden=True), c("Sort Order", "int", True)),
    ),
    Table(
        "Staff Type", "DimStaffType.csv", "Governed Verint Staff Type and published-assignment bridge; never a call queue.",
        (
            c("Staff Type Key", hidden=True), c("Staff Type"), c("Planning Group", hidden=True),
            c("Management LOB", hidden=True), c("Mapping Status"),
        ),
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
        "Forecast", "FactForecastInterval.csv", "Verint Volume and Absolute Required FTE at native 15-minute Staff Type grain.",
        (
            c("Date", "date", True), c("Interval Start"), c("Interval End"),
            c("Interval Minutes", "int", True), c("Time Slot", "int", True),
            c("Staff Type Key", hidden=True), c("Source Staff Type"), c("Staff Type"),
            c("Planning Group"), c("Management LOB", hidden=True), c("Capacity Mapping Status"),
            c("Volume Forecast", "decimal", True), c("Abandons Forecast", "decimal", True),
            c("FTE Forecast", "decimal", True), c("FTE Required", "decimal", True),
            c("Headcount Forecast", "decimal", True), c("Net Staffing Forecast", "decimal", True),
            c("SL Forecast", "decimal", True), c("SL Required", "decimal", True),
            c("AHT Forecast Seconds", "decimal", True), c("Source File"),
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
            m("Expected Requirement Intervals", "COUNTROWS('Forecast')", "#,##0", "Native Staff Type intervals present in the selected forecast extract.", "Forecast"),
            m("Requirement Coverage %", "DIVIDE([Requirement Intervals], [Expected Requirement Intervals])", "0.0%", "Intervals with explicit required FTE divided by supplied Staff Type intervals.", "Forecast"),
            m("Required FTE Hours", "SUMX('Forecast', 'Forecast'[FTE Required] * DIVIDE('Forecast'[Interval Minutes], 60))", "#,##0.0", "Native-interval required FTE converted to additive FTE-hours.", "Forecast"),
            m("Peak Required FTE", "MAXX(SUMMARIZE('Forecast', 'Forecast'[Date], 'Forecast'[Time Slot], \"Interval FTE\", SUM('Forecast'[FTE Required])), [Interval FTE])", "#,##0.0", "Peak summed explicit required FTE in the selected horizon.", "Forecast"),
        ),
    ),
    Table(
        "Staffing", "FactStaffing15Min.csv", "Scheduled and observed capacity at date, interval, Planning Group and Staff Type grain.",
        (
            c("Date", "date", True), c("Interval Start"), c("Interval End"), c("Time Slot", "int", True),
            c("Staff Type Key", hidden=True), c("LOB"), c("Management LOB", hidden=True),
            c("Planning Group"), c("Staff Type"), c("Capacity Mapping Status"),
            c("Language"), c("Scheduled HC", "int", True), c("Observed HC", "int", True),
            c("Productive HC", "int", True), c("Auxiliary HC", "int", True),
            c("Gross Scheduled FTE", "decimal", True), c("Planned Time Off FTE", "decimal", True),
            c("Scheduled FTE", "decimal", True), c("Elapsed Scheduled FTE", "decimal", True),
            c("Observed FTE", "decimal", True),
            c("Productive FTE", "decimal", True), c("Staffing Variance FTE", "decimal", True),
            c("Staffing Gap FTE", "decimal", True), c("Staffing State"), c("Evidence Basis"), c("Evaluation As Of"),
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
            m("Present FTE Gap", "IF(ISBLANK([Required FTE]), BLANK(), [Average Observed FTE] - [Required FTE])", "#,##0.0;[Red]-#,##0.0", "Agent Status-first observed FTE less explicit Staff Type requirement.", "Capacity"),
            m("Gross Scheduled FTE Hours", "SUM('Staffing'[Gross Scheduled FTE]) * 0.25", "#,##0.0", "Gross published schedule capacity converted from 15-minute FTE to hours.", "Capacity"),
            m("Net Scheduled FTE Hours", "SUM('Staffing'[Scheduled FTE]) * 0.25", "#,##0.0", "Published schedule capacity after PTO/Away converted to hours.", "Capacity"),
            m("PTO / Away FTE Hours", "SUM('Staffing'[Planned Time Off FTE]) * 0.25", "#,##0.0", "Governed PTO/Away capacity removed from gross published schedules.", "Capacity"),
            m("Observed FTE Hours", "SUM('Staffing'[Observed FTE]) * 0.25", "#,##0.0", "Agent Status-first observed capacity converted to hours.", "Capacity"),
            m("Productive FTE Hours", "SUM('Staffing'[Productive FTE]) * 0.25", "#,##0.0", "Governed productive observed capacity converted to hours.", "Capacity"),
            m("Scheduled Coverage %", "DIVIDE([Net Scheduled FTE Hours], [Required FTE Hours])", "0.0%", "Net published schedule FTE-hours divided by Verint required FTE-hours.", "Capacity"),
            m("Uncovered FTE Hours", "MAX([Required FTE Hours] - [Net Scheduled FTE Hours], 0)", "#,##0.0", "Selected requirement FTE-hours not covered by net published schedule capacity.", "Capacity"),
        ),
    ),
    Table(
        "Attendance", "FactAttendanceDay.csv", "One governed attendance result per scheduled agent day.",
        (
            c("Date", "date", True), c("Agent Day Key", hidden=True), c("Agent ID", hidden=True),
            c("Agent"), c("Team Leader"), c("Ops Manager"), c("LOB"), c("Management LOB", hidden=True),
            c("Planning Group"), c("Staff Type"), c("Capacity Mapping Status"),
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
        "Shift Placement", "FactShiftPlacement.csv", "Published, observed and exact residual placement bands used by the schedule review timeline.",
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
            m("Placement Duration Hours", "SUM('Shift Placement'[Duration Hours])", "0.0", "Duration of the selected published, observed or residual band.", "Schedule integrity"),
        ),
    ),
    Table(
        "Attendance Gap", "FactAttendanceGap.csv", "Exact start/end attendance gaps prepared for human review.",
        (
            c("Date", "date", True), c("Correction ID", hidden=True), c("Agent ID", hidden=True),
            c("Agent"), c("Team Leader"), c("Ops Manager"), c("LOB"), c("Management LOB", hidden=True),
            c("Scheduled Start"), c("Scheduled End"), c("Gap Start"), c("Gap End"),
            c("Gap Minutes", "int", True), c("Detected Issue"), c("Priority"), c("Confidence"),
            c("Suggested Activity"), c("Observed Source"), c("Reconciliation"),
            c("Verint Activity"), c("Verint Category"), c("Verint Overlap Minutes", "int", True),
            c("Verint Source File"), c("Source File"),
        ),
        (
            m("Total Gap Minutes", "SUM('Attendance Gap'[Gap Minutes])", "#,##0", "Summed exact residual review-gap minutes.", "Attendance review"),
            m("Residual Gap Hours", "DIVIDE([Total Gap Minutes], 60)", "#,##0.0", "Exact residual gap hours still unsupported by final Verint Activities.", "Attendance review"),
        ),
    ),
    Table(
        "Break Meal", "FactBreakMealControl.csv", "Evidence-gated break and meal control using the same rules as Attendance Review.",
        (
            c("Date", "date", True), c("LOB"), c("Management LOB", hidden=True), c("Team Leader"),
            c("Agent ID", hidden=True), c("Agent"), c("Scheduled Start"), c("Scheduled End"),
            c("Status Coverage %", "decimal"), c("Break Minutes", "int", True),
            c("Break Allowance Minutes", "int"), c("Break Overrun Minutes", "int", True),
            c("Break Spells", "int"), c("Longest Break Minutes", "int"),
            c("Meal Minutes", "int", True), c("Meal Allowance Minutes", "int"),
            c("Meal Overrun Minutes", "int", True), c("Meal Spells", "int"),
            c("Longest Meal Minutes", "int"), c("Alert"), c("Evidence"),
        ),
        (
            m("Total Break Overrun Minutes", "SUM('Break Meal'[Break Overrun Minutes])", "#,##0", "Evidence-supported break minutes above the configured allowance.", "Attendance review"),
            m("Total Meal Overrun Minutes", "SUM('Break Meal'[Meal Overrun Minutes])", "#,##0", "Evidence-supported meal minutes above the configured allowance.", "Attendance review"),
            m("Break / Meal Alerts", "CALCULATE(COUNTROWS('Break Meal'), 'Break Meal'[Alert] IN {\"BREAK EXCEEDED\",\"MEAL EXCEEDED\",\"BREAK & MEAL EXCEEDED\"})", "#,##0", "Completed evidence-supported agent-days exceeding a configured allowance.", "Attendance review"),
        ),
    ),
    Table(
        "Operational Action", "FactOperationalAction.csv", "One compact intraday queue combining capacity gaps and attendance call/follow-up facts.",
        (
            c("Date", "date", True), c("Time Slot", "int", True), c("State"), c("Item"),
            c("Agent ID", hidden=True), c("Agent"), c("Team Leader"), c("Management LOB", hidden=True),
            c("Planning Group"), c("Staff Type"), c("Staff Type Key", hidden=True), c("Required FTE", "decimal"),
            c("Resource FTE", "decimal"), c("Variance FTE", "decimal"), c("Detail"),
            c("Evidence"), c("Priority Sort", "int", True),
        ),
        (m("Operational Actions", "COUNTROWS('Operational Action')", "#,##0", "Current capacity and attendance facts requiring operational attention.", "Intraday"),),
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
    ("Planning Group", "Management LOB", "Management LOB", "Management LOB"),
    ("Staff Type", "Planning Group", "Planning Group", "Planning Group"),
    ("Service", "Date", "Date", "Date"), ("Service", "Time Slot", "Time", "Quarter Hour Index"),
    ("Service", "Service Key", "Queue", "Service Key"),
    ("Queue Coverage", "Date", "Date", "Date"),
    ("Queue Coverage", "Service Key", "Queue", "Service Key"),
    ("Forecast", "Date", "Date", "Date"), ("Forecast", "Time Slot", "Time", "Quarter Hour Index"),
    ("Forecast", "Staff Type Key", "Staff Type", "Staff Type Key"),
    ("Staffing", "Date", "Date", "Date"), ("Staffing", "Time Slot", "Time", "Quarter Hour Index"),
    ("Staffing", "Staff Type Key", "Staff Type", "Staff Type Key"),
    ("Attendance", "Date", "Date", "Date"), ("Attendance", "Agent ID", "Employee", "Agent ID"),
    ("Status", "Date", "Date", "Date"), ("Status", "Agent ID", "Employee", "Agent ID"),
    ("Status", "Time Slot", "Time", "Quarter Hour Index"),
    ("Schedule Integrity", "Date", "Date", "Date"),
    ("Schedule Integrity", "Agent ID", "Employee", "Agent ID"),
    ("Shift Placement", "Date", "Date", "Date"),
    ("Shift Placement", "Agent ID", "Employee", "Agent ID"),
    ("Attendance Gap", "Date", "Date", "Date"), ("Attendance Gap", "Agent ID", "Employee", "Agent ID"),
    ("Break Meal", "Date", "Date", "Date"), ("Break Meal", "Agent ID", "Employee", "Agent ID"),
    ("Operational Action", "Date", "Date", "Date"),
    ("Operational Action", "Time Slot", "Time", "Quarter Hour Index"),
    ("Operational Action", "Staff Type Key", "Staff Type", "Staff Type Key"),
    ("Time Off", "Date", "Date", "Date"), ("Time Off", "Agent ID", "Employee", "Agent ID"),
    ("Final Absence", "Date", "Date", "Date"), ("Final Absence", "Agent ID", "Employee", "Agent ID"),
    ("Absence Component", "Date", "Date", "Date"), ("Absence Component", "Agent ID", "Employee", "Agent ID"),
    ("Finding", "Period End", "Date", "Date"),
    ("Finding", "Management LOB", "Management LOB", "Management LOB"),
    ("Quality Issue", "Date", "Date", "Date"), ("Quality Issue", "Agent ID", "Employee", "Agent ID"),
)


PAGES = (
    {
        "title": "Forecast & Requirement", "nav": "Forecast & Requirement", "status": "MONTH PLAN",
        "subtitle": "Verint Staff Type demand and absolute required FTE",
        "rule": "Forecast and requirement are the planning baseline. Call queues are not Staff Types.",
        "scope": "Grain: 15-minute Staff Type\nSource: Verint Volume + Absolute Required FTE",
        "slicers": (("Date", "Date", "FORECAST PERIOD", "Between"), ("Management LOB", "Management LOB", "MANAGEMENT LOB", "Dropdown"), ("Planning Group", "Planning Group", "PLANNING GROUP", "Dropdown"), ("Staff Type", "Staff Type", "STAFF TYPE", "Dropdown")),
        "cards": (
            ("Forecast", "Forecast Volume", "FORECAST VOLUME", "#315F85", "Sum of populated Staff Type volume"),
            ("Forecast", "Required FTE Hours", "REQUIRED FTE-HOURS", "#008B95", "15-minute required FTE converted to hours"),
            ("Forecast", "Peak Required FTE", "PEAK REQUIRED FTE", "#D18A13", "Peak at selected Staff Type grain"),
            ("Forecast", "Requirement Coverage %", "REQUIREMENT COVERAGE", "#26805A", "Intervals with explicit absolute requirement"),
        ),
        "charts": (
            {"type": "lineStackedColumnComboChart", "title": "VOLUME AND REQUIRED FTE PROFILE", "category": ("Time", "Time Label"), "values": (("Forecast", "Forecast Volume"),), "secondary": (("Forecast", "Required FTE"),)},
            {"type": "table", "title": "REQUIREMENT BY STAFF TYPE", "fields": (("Staff Type", "Staff Type"), ("Forecast", "Forecast Volume"), ("Forecast", "Required FTE Hours"), ("Forecast", "Peak Required FTE"))},
        ),
        "tables": (("STAFF TYPE REQUIREMENT DETAIL", (("Date", "Date"), ("Time", "Time Label"), ("Management LOB", "Management LOB"), ("Planning Group", "Planning Group"), ("Staff Type", "Staff Type"), ("Forecast", "Forecast Volume"), ("Forecast", "Required FTE"), ("Forecast", "Capacity Mapping Status")), "full"),),
    },
    {
        "title": "Staff Preparation", "nav": "Staff Preparation", "status": "NEXT 14 DAYS",
        "subtitle": "Required capacity versus published schedules — by planning group and Staff Type",
        "rule": "Forecast and requirement follow Staff Type. Service level remains a separate LOB result.",
        "scope": "Service: combined management LOB result\nStaffing: planning groups remain separate",
        "slicers": (("Date", "Date", "PLANNING HORIZON", "Between"), ("Management LOB", "Management LOB", "MANAGEMENT LOB", "Dropdown"), ("Planning Group", "Planning Group", "PLANNING GROUP", "Dropdown"), ("Staff Type", "Staff Type", "STAFF TYPE", "Dropdown")),
        "cards": (
            ("Forecast", "Peak Required FTE", "PEAK REQUIRED FTE", "#315F85", "Explicit Verint requirement"),
            ("Staffing", "Peak Shortage FTE", "PEAK STAFFING GAP", "#BD2B32", "Worst net schedule position"),
            ("Staffing", "Uncovered FTE Hours", "UNCOVERED FTE-HOURS", "#BD2B32", "Requirement not covered by net schedule"),
            ("Staffing", "PTO / Away FTE Hours", "PTO / AWAY IMPACT", "#D18A13", "Removed from gross schedule capacity"),
        ),
        "charts": (
            {"type": "lineChart", "title": "REQUIRED FTE vs NET SCHEDULED FTE", "category": ("Time", "Time Label"), "values": (("Forecast", "Required FTE"), ("Staffing", "Average Scheduled FTE"))},
            {"type": "table", "title": "CAPACITY BY PLANNING GROUP", "fields": (("Planning Group", "Planning Group"), ("Staff Type", "Staff Type"), ("Forecast", "Required FTE"), ("Staffing", "Average Scheduled FTE"), ("Staffing", "Net Capacity Gap FTE"))},
        ),
        "tables": (("STAFFING GAPS TO TREAT", (("Date", "Date"), ("Time", "Time Label"), ("Management LOB", "Management LOB"), ("Planning Group", "Planning Group"), ("Staff Type", "Staff Type"), ("Forecast", "Required FTE"), ("Staffing", "Average Scheduled FTE"), ("Staffing", "PTO / Away FTE"), ("Staffing", "Net Capacity Gap FTE")), "full"),),
    },
    {
        "title": "Intraday Control", "nav": "Intraday Control", "status": "LIVE",
        "subtitle": "Live service outcome and current staffing position — kept at their correct grains",
        "rule": "Operate service at LOB level and resources at Planning Group / Staff Type level.",
        "scope": "Service: exact queues rolled to one LOB ratio\nResources: planning groups stay visible",
        "equal_charts": True,
        "slicers": (("Date", "Date", "BUSINESS DATE", "Between"), ("Management LOB", "Management LOB", "MANAGEMENT LOB", "Dropdown"), ("Planning Group", "Planning Group", "PLANNING GROUP", "Dropdown"), ("Time", "Time Label", "CHECKPOINT", "Dropdown")),
        "cards": (
            ("Service", "Service Level %", "SERVICE LEVEL", "#D18A13", "Ratio of summed Storm components"),
            ("Service", "Offered Calls", "OFFERED VOLUME", "#315F85", "Exact configured service queue scope"),
            ("Staffing", "Present FTE Gap", "PRESENT FTE GAP", "#BD2B32", "Requirement less Agent Status presence"),
            ("Attendance", "No Show HC", "CONFIRMED NO SHOW", "#BD2B32", "Unknown evidence remains separate"),
        ),
        "charts": (
            {"type": "lineChart", "title": "COMBINED LOB SERVICE LEVEL", "category": ("Time", "Time Label"), "values": (("Service", "Service Level %"), ("Service", "SL Target %"))},
            {"type": "table", "title": "RESOURCE POSITION BY PLANNING GROUP", "fields": (("Planning Group", "Planning Group"), ("Staff Type", "Staff Type"), ("Forecast", "Required FTE"), ("Staffing", "Average Scheduled FTE"), ("Staffing", "Average Observed FTE"), ("Staffing", "Present FTE Gap"))},
        ),
        "tables": (("NEXT INTERVALS AND ATTENDANCE CONTROL", (("Operational Action", "State"), ("Operational Action", "Item"), ("Operational Action", "Planning Group"), ("Operational Action", "Staff Type"), ("Operational Action", "Required FTE"), ("Operational Action", "Resource FTE"), ("Operational Action", "Variance FTE"), ("Operational Action", "Detail"), ("Operational Action", "Evidence")), "full"),),
    },
    {
        "title": "Attendance & Schedule Review", "nav": "Attendance & Schedule Review", "status": "CURRENT WEEK",
        "subtitle": "Completed-shift evidence, exact residual corrections, breaks and meals",
        "rule": "Agent Status owns observed attendance. Final Verint Activities close exact residual gaps.",
        "scope": "Observed: Agent Status first, LILO fallback\nCorrection: Activities subtract exact overlap",
        "slicers": (("Date", "Date", "COMPLETED PERIOD", "Between"), ("Management LOB", "Management LOB", "MANAGEMENT LOB", "Dropdown"), ("Employee", "Team Leader", "TEAM LEADER", "Dropdown"), ("Attendance Gap", "Detected Issue", "EXCEPTION", "Dropdown")),
        "cards": (
            ("Attendance", "No Show HC", "CONFIRMED NO SHOW", "#BD2B32", "Completed supported cases"),
            ("Attendance", "Late HC", "LATE ARRIVALS", "#D18A13", "Exact start variance above tolerance"),
            ("Attendance", "Early Leave HC", "EARLY LEAVES", "#D18A13", "Completed shifts only"),
            ("Attendance Gap", "Residual Gap Hours", "RESIDUAL GAP HOURS", "#BD2B32", "Still unsupported by final Activities"),
        ),
        "charts": (
            {"type": "stackedBarChart", "title": "PUBLISHED SCHEDULE vs OBSERVED PRESENCE", "category": ("Shift Placement", "Placement Label"), "values": (("Shift Placement", "Start Hour Value"), ("Shift Placement", "Placement Duration Hours"))},
            {"type": "table", "title": "BREAK & MEAL CONTROL", "fields": (("Break Meal", "Agent"), ("Break Meal", "Break Minutes"), ("Break Meal", "Meal Minutes"), ("Break Meal", "Break Allowance Minutes"), ("Break Meal", "Meal Allowance Minutes"), ("Break Meal", "Alert"))},
        ),
        "tables": (("RESIDUAL VERINT CORRECTION QUEUE", (("Attendance Gap", "Date"), ("Attendance Gap", "Agent"), ("Attendance Gap", "Team Leader"), ("Attendance Gap", "Detected Issue"), ("Attendance Gap", "Gap Start"), ("Attendance Gap", "Gap End"), ("Attendance Gap", "Gap Minutes"), ("Attendance Gap", "Verint Activity"), ("Attendance Gap", "Verint Overlap Minutes"), ("Attendance Gap", "Suggested Activity")), "full"),),
    },
    {
        "title": "Performance Review", "nav": "Performance Review", "status": "PERIOD REVIEW",
        "subtitle": "Close the WFM cycle with forecast, staffing, service, absence and shrinkage",
        "rule": "Review source variances, then improve the next forecast and staff plan. No synthetic score.",
        "scope": "Cycle: demand → requirement → schedule → delivery\nRates: recalculated from summed components",
        "equal_charts": True,
        "slicers": (("Date", "Date", "REVIEW PERIOD", "Between"), ("Date", "Year Month", "COMPARISON MONTH", "Dropdown"), ("Management LOB", "Management LOB", "MANAGEMENT LOB", "Dropdown"), ("Planning Group", "Planning Group", "PLANNING GROUP", "Dropdown")),
        "cards": (
            ("Service", "Service Level %", "SERVICE LEVEL", "#D18A13", "Against governed LOB target"),
            ("Service", "Volume Variance %", "VOLUME vs FORECAST", "#315F85", "Actual offered versus compatible roll-up"),
            ("Staffing", "Scheduled Coverage %", "SCHEDULED COVERAGE", "#26805A", "Net scheduled FTE-hours / required"),
            ("Final Absence", "Final Absence %", "FINAL ABSENCE", "#BD2B32", "Final Verint Activities only"),
        ),
        "charts": (
            {"type": "lineChart", "title": "WEEKLY FORECAST vs ACTUAL VOLUME", "category": ("Date", "ISO Week"), "values": (("Forecast", "Forecast Volume"), ("Service", "Offered Calls"))},
            {"type": "clusteredColumnChart", "title": "REQUIREMENT TO DELIVERY — FTE-HOURS", "category": ("Capacity Stage", "Capacity Stage"), "values": (("Status", "Capacity Bridge Hours"),)},
        ),
        "tables": (("MONTHLY WFM SCORECARD", (("Management LOB", "Management LOB"), ("Service", "Service Level %"), ("Service", "Offered Calls"), ("Forecast", "Forecast Volume"), ("Service", "Volume Variance %"), ("Forecast", "Required FTE Hours"), ("Staffing", "Net Scheduled FTE Hours"), ("Staffing", "Observed FTE Hours"), ("Final Absence", "Final Absence %"), ("Final Absence", "Final Shrinkage %")), "full"),),
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
        "position": _position(12, 105, 186, len(PAGES) * 68, 20, 20),
        "visual": {
            "visualType": "pageNavigator",
            "objects": {
                "layout": [{"properties": {"columnCount": _literal(1), "rowCount": _literal(len(PAGES)), "cellPadding": _literal(5)}}],
                "pages": [{"properties": {"showHiddenPages": _literal(False), "showTooltipPages": _literal(False), "showByDefault": _literal(True)}}],
                "shape": [{"properties": {"tileShape": _literal("rectangleRoundedByPixel"), "rectangleRoundedCurve": _literal(6)}}],
                "text": [
                    {"properties": {"show": _literal(False)}, "selector": {"id": "default"}},
                    {"properties": {"show": _literal(False)}, "selector": {"id": "selected"}},
                ],
                "fill": [
                    {"properties": {"show": _literal(True), "fillColor": _color("#0B1F33"), "transparency": _literal(100)}, "selector": {"id": "default"}},
                    {"properties": {"show": _literal(True), "fillColor": _color("#007C83"), "transparency": _literal(100)}, "selector": {"id": "selected"}},
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
        _write_visual(page_dir, title, "sidebar", _shape("", 0, 0, 210, 945, "#0B2239", 0))
        _write_visual(page_dir, title, "header", _shape("", 210, 0, 1470, 72, "#FFFFFF", 1))
        _write_visual(page_dir, title, "header line", _shape("", 210, 71, 1470, 1, "#D6E0E6", 2))
        _write_visual(page_dir, title, "brand", _textbox("", "WFMHub", 20, 16, 170, 45, 27, "#FFFFFF", 10, True, "left"))
        _write_visual(page_dir, title, "nav label", _textbox("", "WFM CYCLE", 24, 88, 160, 20, 10, "#A9BFCE", 10, True))
        # Static labels guarantee a readable sidebar even on Desktop builds
        # that fail to render pageNavigator text.  The native navigator stays
        # above them as the interactive layer; Power BI's bottom page tabs are
        # also retained as a second navigation route.
        current_nav = next(
            index for index, candidate in enumerate(PAGES)
            if candidate["title"] == title
        )
        for index, candidate in enumerate(PAGES):
            nav_y = 112 + index * 68
            if index == current_nav:
                _write_visual(
                    page_dir, title, f"nav selected {index}",
                    _shape("", 12, nav_y, 186, 63, "#008B95", 12, rounded=True),
                )
            _write_visual(
                page_dir, title, f"nav text {index}",
                _textbox(
                    "", f"{index + 1}   {candidate['nav']}",
                    26, nav_y + 20, 158, 28, 11,
                    "#FFFFFF" if index == current_nav else "#DCE9F2",
                    14, index == current_nav,
                ),
            )
        _write_visual(page_dir, title, "nav", _navigator(""))
        _write_visual(page_dir, title, "business rule", _textbox("", "BUSINESS RULE\n" + page["rule"], 24, 792, 162, 104, 10, "#AFC2CF", 10, False, "left"))
        _write_visual(page_dir, title, "page title", _textbox("", title.upper(), 235, 11, 720, 30, 25, "#17324D", 10, True))
        _write_visual(page_dir, title, "subtitle", _textbox("", page["subtitle"], 235, 43, 850, 20, 11, "#607587", 10, False))
        _write_visual(page_dir, title, "update", _textbox("", "Last Hub update · refresh the governed feed", 1150, 24, 330, 20, 10, "#607587", 10, False, "right"))
        _write_visual(page_dir, title, "status panel", _shape("", 1494, 19, 156, 34, "#E4F4F4", 8, rounded=True))
        _write_visual(page_dir, title, "status", _textbox("", page["status"], 1500, 27, 144, 18, 9, "#08757D", 10, True, "center"))

        # One compact selector strip shared by every page.
        _write_visual(page_dir, title, "selector panel", _shape("", 228, 85, 1434, 58, "#FFFFFF", 3, rounded=True))
        slicer_layout = ((240, 270), (522, 235), (769, 235), (1016, 270))
        for index, ((table, field, label, mode), (x, width)) in enumerate(zip(page["slicers"], slicer_layout), 1):
            _write_visual(page_dir, title, f"slicer {index}", _slicer("", table, field, label, mode, x, 88, width, 52))
        _write_visual(page_dir, title, "scope panel", _shape("", 1298, 94, 352, 40, "#E9F5F5", 4, rounded=True))
        _write_visual(page_dir, title, "scope accent", _shape("", 1298, 94, 4, 40, "#008B95", 5))
        _write_visual(page_dir, title, "scope", _textbox("", page["scope"], 1310, 98, 328, 32, 9, "#315169", 10, False))

        # Four aligned KPI cards with a restrained semantic accent.
        card_x = (228, 590, 951, 1313)
        for index, ((table, measure, label, accent, subtext), x) in enumerate(zip(page["cards"], card_x), 1):
            _write_visual(page_dir, title, f"card panel {index}", _shape("", x, 153, 349, 119, "#FFFFFF", 3, rounded=True))
            _write_visual(page_dir, title, f"card accent {index}", _shape("", x, 153, 6, 119, accent, 4, rounded=True))
            _write_visual(page_dir, title, f"card {index}", _card("", table, measure, label, accent, x + 8, 155, 337, 84))
            _write_visual(page_dir, title, f"card note {index}", _textbox("", subtext, x + 24, 244, 305, 20, 9, "#607587", 10, False))

        # The two analytical panels occupy the same visual rhythm on all pages.
        chart_positions = (
            ((228, 282, 711, 330), (951, 282, 711, 330))
            if page.get("equal_charts") else
            ((228, 282, 872, 330), (1112, 282, 550, 330))
        )
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
            "full": (228, 622, 1434, 248),
            "left": (228, 622, 711, 248),
            "right": (951, 622, 711, 248),
        }
        for index, (table_title, table_fields, layout) in enumerate(page["tables"], 1):
            x, y, width, height = table_layouts[layout]
            _write_visual(page_dir, title, f"table panel {index}", _shape("", x, y, width, height, "#FFFFFF", 3, rounded=True))
            _write_visual(page_dir, title, f"table {index}", _table("", table_title, table_fields, x + 8, y + 6, width - 16, height - 12))

        _write_visual(page_dir, title, "footer", _textbox("", "Prepared by Anass ASSRI | WFM   •   Governed feed: POWERBI_MANIFEST_CURRENT.csv", 228, 910, 1434, 18, 8, "#687E8E", 10, False, "right"))

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


def _validate_project(root: Path) -> None:
    """Fail before publication when the generated PBIP is incomplete."""
    for path in root.rglob("*.json"):
        json.loads(path.read_text(encoding="utf-8"))
    page_files = list((
        root / f"{PROJECT_NAME}.Report" / "definition" / "pages"
    ).glob("ReportSection*/page.json"))
    if len(page_files) != len(PAGES):
        raise ValueError(f"Generated {len(page_files)} report pages; expected {len(PAGES)}")
    actual = {
        json.loads(path.read_text(encoding="utf-8"))["displayName"]
        for path in page_files
    }
    expected = {page["nav"] for page in PAGES}
    if actual != expected:
        raise ValueError(f"Generated page contract differs: {actual ^ expected}")
    tables = root / f"{PROJECT_NAME}.SemanticModel" / "definition" / "tables"
    missing = [spec.name for spec in TABLES if not (tables / f"{spec.name}.tmdl").exists()]
    if missing:
        raise ValueError(f"Generated semantic tables missing: {', '.join(missing)}")


def build() -> Path:
    PROJECT_ROOT.parent.mkdir(parents=True, exist_ok=True)
    staged = Path(tempfile.mkdtemp(prefix=".WFMHubBI-build-", dir=PROJECT_ROOT.parent))
    backup = PROJECT_ROOT.parent / f".WFMHubBI-backup-{uuid.uuid4().hex}"
    try:
        _write_model(staged)
        _write_report(staged)
        _write_json(
            staged / f"{PROJECT_NAME}.SemanticModel" / ".platform",
            {
                "$schema": "https://developer.microsoft.com/json-schemas/fabric/gitIntegration/platformProperties/2.0.0/schema.json",
                "metadata": {"type": "SemanticModel", "displayName": PROJECT_NAME},
                "config": {"version": "2.0", "logicalId": _stable_guid("WFMHub BI Semantic Model")},
            },
        )
        _write_json(
            staged / f"{PROJECT_NAME}.pbip",
            {"$schema": "https://developer.microsoft.com/json-schemas/fabric/pbip/pbipProperties/1.0.0/schema.json", "version": "1.0", "artifacts": [{"report": {"path": f"{PROJECT_NAME}.Report"}}], "settings": {"enableAutoRecovery": True}},
        )
        (staged / "PROJECT_VERSION.txt").write_text("6\n", encoding="utf-8")
        (staged / "README.txt").write_text(
            "WFMHUB BI\n=========\n\n"
            "Open WFMHub BI.pbip with Microsoft Power BI Desktop, then choose Home > Refresh.\n"
            "The HubRoot parameter is set automatically when this project is opened through POWERBI.cmd.\n"
            "The model reads only Feed\\PowerBI CSVs and never reads SQLite or raw extracts.\n"
            "The five pages follow the WFM cycle: forecast, staff preparation, intraday, attendance review and performance review.\n"
            "PCS remains a separate collaborative Excel product and is not imported into this WFM-only model.\n",
            encoding="utf-8",
        )
        _validate_project(staged)
        if PROJECT_ROOT.exists():
            PROJECT_ROOT.replace(backup)
        try:
            staged.replace(PROJECT_ROOT)
        except Exception:
            if backup.exists() and not PROJECT_ROOT.exists():
                backup.replace(PROJECT_ROOT)
            raise
        if backup.exists():
            shutil.rmtree(backup)
    except Exception:
        if staged.exists():
            shutil.rmtree(staged)
        raise
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
