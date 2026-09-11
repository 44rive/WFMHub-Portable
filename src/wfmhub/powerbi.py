"""Publish a fixed-schema Power BI star-model feed and its setup assets."""

from __future__ import annotations

import shutil
from datetime import date, datetime, time, timedelta
from pathlib import Path
from typing import Any, Sequence

from .config import Config
from .database import DatabaseConnection
from .shared_feeds import SharedFeedResult, _atomic_csv, _manifest


POWERBI_SCHEMA_VERSION = "1"


def _query(
    conn: DatabaseConnection,
    folder: Path,
    filename: str,
    headers: Sequence[str],
    sql: str,
    parameters: Sequence[Any] = (),
) -> tuple[Path, int]:
    path = folder / filename
    rows = conn.execute(sql, list(parameters)).fetchall()
    return path, _atomic_csv(path, headers, rows)


def _date_rows(start: date, end: date):
    current = start
    while current <= end:
        yield (
            current, current.year, current.month, current.strftime("%B"),
            current.strftime("%Y-%m"), current.isocalendar().week,
            current.strftime("%a"), current.weekday() + 1,
            current.replace(day=1), current == end,
        )
        current += timedelta(days=1)


def _time_rows():
    anchor = datetime.combine(date(2000, 1, 1), time())
    for slot in range(96):
        value = anchor + timedelta(minutes=slot * 15)
        yield (
            value.time().strftime("%H:%M:%S"), slot, value.hour,
            value.strftime("%H:00"), value.strftime("%H:%M"),
        )


def _copy_setup_assets(config: Config, folder: Path) -> tuple[Path, ...]:
    source = config.home / "templates" / "powerbi"
    if not source.is_dir():
        source = config.system / "templates" / "powerbi"
    target = folder / "_SETUP"
    if not source.is_dir():
        return ()
    target.mkdir(parents=True, exist_ok=True)
    copied: list[Path] = []
    for item in source.iterdir():
        if not item.is_file():
            continue
        destination = target / item.name
        shutil.copy2(item, destination)
        copied.append(destination)
    return tuple(copied)


def publish_powerbi_feeds(
    conn: DatabaseConnection,
    config: Config,
    start: date,
    end: date,
) -> SharedFeedResult:
    """Publish additive facts; KPI rates remain ratios of summed counters."""

    available = conn.execute(
        """SELECT min(business_date), max(business_date) FROM (
             SELECT business_date FROM mart.service_interval
             UNION ALL SELECT business_date FROM mart.forecast_interval
             UNION ALL SELECT business_date FROM mart.staffing_interval
             UNION ALL SELECT business_date FROM mart.attendance_agent_day
             UNION ALL SELECT business_date FROM mart.agent_pcs_day
             UNION ALL SELECT business_date FROM mart.verint_final_absence_agent_day
             UNION ALL SELECT business_date FROM mart.planned_time_off_segment
           )"""
    ).fetchone()
    def as_date(value: Any, fallback: date) -> date:
        if isinstance(value, date):
            return value
        return date.fromisoformat(str(value)[:10]) if value else fallback

    feed_start = as_date(available[0], start)
    feed_end = as_date(available[1], end)
    folder = config.feed / "PowerBI"
    folder.mkdir(parents=True, exist_ok=True)
    files: list[Path] = []
    counts: list[tuple[str, int]] = []

    dim_date = folder / "DimDate.csv"
    count = _atomic_csv(
        dim_date,
        (
            "Date", "Year", "Month Number", "Month", "Year Month",
            "ISO Week", "Weekday", "Weekday Number", "Month Start",
            "Is Data Through",
        ),
        _date_rows(feed_start, feed_end),
    )
    files.append(dim_date)
    counts.append((dim_date.name, count))

    dim_time = folder / "DimTime.csv"
    count = _atomic_csv(
        dim_time,
        ("Time", "Quarter Hour Index", "Hour", "Hour Label", "Time Label"),
        _time_rows(),
    )
    files.append(dim_time)
    counts.append((dim_time.name, count))

    specs = (
        (
            "DimEmployee.csv",
            ("Agent ID", "Agent", "Employment Status", "Team Leader", "Ops Manager", "LOB", "Market", "Language", "Location", "City", "FTE"),
            """SELECT agent_id, canonical_name, employment_status, team_leader,
                      ops_manager, lob, market, language, location, city, fte
               FROM core.dim_agent ORDER BY lob, team_leader, canonical_name, agent_id""",
        ),
        (
            "DimLOB.csv",
            ("LOB",),
            """SELECT lob FROM (
                 SELECT DISTINCT trim(lob) AS lob FROM core.dim_agent
                 UNION SELECT DISTINCT trim(service_scope) FROM mart.service_interval
                 UNION SELECT DISTINCT trim(lob) FROM mart.attendance_agent_day
                 UNION SELECT DISTINCT trim(lob) FROM mart.agent_pcs_day
               ) WHERE trim(coalesce(lob,''))<>'' ORDER BY lob""",
        ),
        (
            "DimQueue.csv",
            ("Queue", "Source System", "Service Scope", "Comparison Scope", "Designation", "Mapping Status"),
            """SELECT queue, source_system, service_scope, comparison_scope,
                      designation, mapping_status
               FROM (
                 SELECT DISTINCT queue, source_system, service_scope,
                        comparison_scope, designation, mapping_status
                 FROM mart.service_interval
                 WHERE trim(coalesce(queue,''))<>''
               ) q ORDER BY service_scope, queue""",
        ),
        (
            "FactServiceHour.csv",
            (
                "Date", "Hour Start", "Source System", "Service Scope",
                "Comparison Scope", "Queue", "Designation", "Offered",
                "Answered", "Abandoned", "Short Abandoned",
                "Abandoned Within Target", "Answered Within Target",
                "Handled Seconds", "SLA Denominator", "Source Files",
            ),
            """SELECT business_date, hour_start, source_system, service_scope,
                      comparison_scope, queue, designation,
                      sum(coalesce(offered,0)), sum(coalesce(answered,0)),
                      sum(coalesce(abandoned,0)), sum(coalesce(short_abandoned,0)),
                      sum(coalesce(abandoned_within_target,0)),
                      sum(coalesce(answered_within_target,0)),
                      sum(coalesce(handled_seconds,0)),
                      sum(coalesce(offered,0)-coalesce(abandoned_within_target,0)),
                      group_concat(DISTINCT source_file)
               FROM mart.service_interval
               GROUP BY business_date, hour_start, source_system, service_scope,
                        comparison_scope, queue, designation
               ORDER BY business_date, hour_start, service_scope, queue""",
        ),
        (
            "FactForecastInterval.csv",
            (
                "Date", "Interval Start", "Interval End", "Interval Minutes",
                "Queue", "Volume Forecast", "Abandons Forecast", "FTE Forecast",
                "FTE Required", "Headcount Forecast", "Net Staffing Forecast",
                "SL Forecast", "SL Required", "AHT Forecast Seconds",
                "Service Scope", "Comparison Scope", "Mapping Status", "Source File",
            ),
            """SELECT business_date, interval_start, interval_end, interval_minutes,
                      queue_name, volume_forecast, abandons_forecast, fte_forecast,
                      fte_required, headcount_forecast, net_staffing_forecast,
                      sl_forecast, sl_required, aht_forecast_seconds,
                      service_scope, comparison_scope, mapping_status, source_file
               FROM mart.forecast_interval
               ORDER BY business_date, interval_start, service_scope, queue_name""",
        ),
        (
            "FactStaffing15Min.csv",
            (
                "Date", "Interval Start", "Interval End", "LOB", "Language",
                "Gross Scheduled FTE", "Planned Time Off FTE", "Scheduled FTE",
                "Observed FTE", "Productive FTE", "Staffing Variance FTE",
                "Staffing Gap FTE", "Staffing State", "Evidence Basis", "Evaluation As Of",
            ),
            """SELECT business_date, interval_start, interval_end, lob, language,
                      gross_scheduled_fte, planned_time_off_fte, scheduled_fte,
                      observed_fte, productive_fte, staffing_variance_fte,
                      staffing_gap_fte, staffing_state, evidence_basis, evaluation_as_of
               FROM mart.staffing_interval
               ORDER BY business_date, interval_start, lob, language""",
        ),
        (
            "FactAttendanceDay.csv",
            (
                "Date", "Agent Day Key", "Agent ID", "Agent", "Team Leader",
                "Ops Manager", "LOB", "Market", "Language", "Location",
                "Scheduled Start", "Scheduled End", "Scheduled Minutes",
                "Planned Work Minutes", "Planning Overlay", "Planning Overlay Minutes",
                "First Login", "Last Logout", "Attendance Result", "Call Action",
                "Requires Call", "Shift State", "Actual Evidence", "Is Provisional",
                "Late Minutes", "Early Leave Minutes", "No Show Minutes",
                "Status Covered Minutes", "Evaluation As Of",
            ),
            """SELECT business_date, agent_day_key, agent_id, agent_name, team_leader,
                      ops_manager, lob, market, language, location,
                      scheduled_start, scheduled_end, scheduled_minutes,
                      planned_work_minutes, planning_overlay, planning_overlay_minutes,
                      first_login, last_logout, attendance_result, call_action,
                      requires_call, shift_state, actual_evidence, is_provisional,
                      uncoded_late_minutes, uncoded_early_leave_minutes,
                      no_show_minutes, status_covered_minutes, evaluation_as_of
               FROM mart.attendance_agent_day
               ORDER BY business_date, lob, team_leader, agent_name, agent_id""",
        ),
        (
            "FactAttendanceGap.csv",
            (
                "Date", "Correction ID", "Agent ID", "Agent", "Team Leader",
                "Ops Manager", "LOB", "Scheduled Start", "Scheduled End",
                "Gap Start", "Gap End", "Gap Minutes", "Detected Issue",
                "Priority", "Confidence", "Suggested Activity", "Observed Source",
                "Reconciliation", "Source File",
            ),
            """SELECT c.business_date, c.correction_id, c.agent_id, c.agent_name,
                      c.team_leader, c.ops_manager, c.lob, c.scheduled_start,
                      c.scheduled_end, c.gap_start, c.gap_end, c.gap_minutes,
                      c.detected_issue, c.priority, c.confidence,
                      c.suggested_activity, c.observed_source,
                      c.verint_reconciliation, c.source_file
               FROM mart.correction_candidate c
               ORDER BY c.business_date, c.lob, c.agent_name, c.gap_start""",
        ),
        (
            "FactPCSAgentDay.csv",
            (
                "Date", "Agent Day Key", "Agent ID", "Agent", "Team Leader",
                "Ops Manager", "LOB", "Market", "Language", "Location",
                "Call Legs", "Handled Calls", "Inbound Calls", "Outbound Calls",
                "Handle Seconds", "PCS Score Sum", "Valid PCS", "PCS Status Calls",
                "PCS Participation Responses", "Low Score Responses",
                "Top Box Responses", "Comments Count",
            ),
            """SELECT business_date, agent_day_key, agent_id, agent_name, team_leader,
                      ops_manager, lob, market, language, location,
                      call_legs, handled_calls, inbound_calls, outbound_calls,
                      handle_seconds, pcs_score_sum, survey_responses,
                      pcs_status_calls, pcs_participation_responses,
                      low_score_responses, top_box_responses, comments_count
               FROM mart.agent_pcs_day
               ORDER BY business_date, lob, team_leader, agent_name, agent_id""",
        ),
        (
            "FactPCSCoaching.csv",
            (
                "Date", "Call Start", "Agent ID", "Agent", "Team Leader", "LOB",
                "Call ID", "Call Key", "Q1 Score", "Customer Comment", "PCS Status",
            ),
            """SELECT c.business_date, c.call_start, c.agent_id,
                      coalesce(a.canonical_name,c.agent_name), a.team_leader, a.lob,
                      c.call_id, c.call_key, c.question_1_score,
                      coalesce(c.question_2,c.question_3,c.question_4,c.question_5,
                               c.question_6,c.question_7,c.question_8,c.question_9,
                               c.question_10), c.pcs_status
               FROM core.clean_call_leg c
               LEFT JOIN core.dim_agent a ON a.agent_id=c.agent_id
               WHERE c.question_1_score IS NOT NULL AND c.question_1_score<=3
               ORDER BY c.business_date DESC, c.question_1_score, a.lob,
                        a.team_leader, a.canonical_name""",
        ),
        (
            "FactTimeOff.csv",
            (
                "Date", "Segment Key", "Agent Day Key", "Agent ID", "Agent",
                "Team Leader", "Ops Manager", "LOB", "Language", "Source Kind",
                "Absence Type", "Record Status", "Segment Start", "Segment End",
                "Planned Minutes", "Source File", "Source Sheet", "Source Row",
            ),
            """SELECT business_date, segment_key, agent_day_key, agent_id,
                      agent_name, team_leader, ops_manager, lob, language,
                      source_kind, absence_type, record_status, segment_start,
                      segment_end, planned_minutes, source_file, source_sheet,
                      source_row
               FROM mart.planned_time_off_segment
               ORDER BY business_date, lob, team_leader, agent_name, segment_start""",
        ),
        (
            "FactFinalAbsenceDay.csv",
            (
                "Date", "Agent Day Key", "Agent ID", "Agent", "Team Leader",
                "Ops Manager", "LOB", "Market", "Language", "Location",
                "Scheduled Minutes", "Planned Net Minutes", "Absence Minutes",
                "Vacation Minutes", "Unpaid Minutes", "Shrinkage Minutes",
                "Unmapped Minutes", "Final Ledger Status", "Rule Version",
            ),
            """SELECT business_date, agent_day_key, agent_id, agent_name, team_leader,
                      ops_manager, lob, market, language, location,
                      scheduled_minutes, planned_net_minutes, final_absence_minutes,
                      final_vacation_minutes, final_unpaid_minutes,
                      final_shrinkage_minutes, final_unmapped_minutes,
                      final_ledger_status, rule_version
               FROM mart.verint_final_absence_agent_day
               ORDER BY business_date, lob, team_leader, agent_name, agent_id""",
        ),
        (
            "FactFinalAbsenceComponent.csv",
            (
                "Date", "Event Key", "Agent Day Key", "Agent ID", "Agent",
                "Team Leader", "LOB", "Activity", "Category", "Event Start",
                "Event End", "Minutes", "Counts As Absence", "Counts As Vacation",
                "Counts As Unpaid", "Counts As Shrinkage", "Mapped",
                "Evidence Type", "Source File", "Rule Version",
            ),
            """SELECT business_date, event_key, agent_day_key, agent_id, agent_name,
                      team_leader, lob, activity, category, event_start, event_end,
                      minutes, counts_as_absence, counts_as_vacation,
                      counts_as_unpaid, counts_as_shrinkage, mapped,
                      evidence_type, source_file, rule_version
               FROM mart.verint_final_absence_event
               ORDER BY business_date, lob, agent_name, event_start""",
        ),
        (
            "FactBonusMonth.csv",
            (
                "Period", "Agent ID", "Agent", "Team Leader", "Ops Manager",
                "Population", "Core Ready", "Eligibility", "AHT Earned",
                "Productivity Earned", "PCS Earned", "Participation Earned",
                "QM Earned", "Absence Earned", "Extra PCS Earned",
                "Gross Achievement", "VOC Malus", "Final Achievement",
                "Reference Bonus", "Proration", "Scenario Payout",
                "Released Payout", "Release Status", "Data Issue", "Import ID",
            ),
            """SELECT period, agent_id, agent_name, team_leader, ops_manager,
                      population, core_ready, eligibility, aht_earned,
                      productivity_earned, pcs_earned, participation_earned,
                      qm_earned, absence_earned, extra_pcs_earned,
                      gross_achievement, voc_malus, final_achievement,
                      reference_bonus, proration, scenario_payout,
                      released_payout, release_status, data_issue, import_id
               FROM mart.bonus_agent_month
               ORDER BY period, population, team_leader, agent_name, agent_id""",
        ),
        (
            "FactBonusKPI.csv",
            (
                "Period", "Agent ID", "Agent", "Population", "KPI",
                "Actual Value", "Earned Weight", "Direction", "Tier 1 Target",
                "Tier 2 Target", "Import ID",
            ),
            """SELECT period, agent_id, agent_name, population, kpi,
                      actual_value, earned_weight, direction, tier1_target,
                      tier2_target, import_id
               FROM mart.bonus_kpi_result
               ORDER BY period, population, agent_name, agent_id, kpi""",
        ),
        (
            "FactMetricValue.csv",
            (
                "Date", "Interval Start", "Metric Key", "Metric ID", "Method ID",
                "Domain", "Unit", "Aggregation", "Source System", "LOB", "Language",
                "Team Leader", "Agent ID", "Numerator", "Denominator", "Sample Size",
                "Metric Value", "Target Value", "Metric State", "Catalog Version",
            ),
            """SELECT business_date, interval_start, metric_key, metric_id, method_id,
                      domain, unit, aggregation, source_system, lob, language,
                      team_leader, agent_id, numerator, denominator, sample_size,
                      metric_value, target_value, metric_state, catalog_version
               FROM mart.metric_value
               ORDER BY business_date, domain, metric_id, lob, team_leader, agent_id""",
        ),
        (
            "FactFinding.csv",
            (
                "Period Start", "Period End", "Finding ID", "Rank", "Finding Type",
                "Severity", "Domain", "Metric ID", "Source System", "LOB", "Language",
                "Team Leader", "Agent ID", "Title", "Summary", "Current Value",
                "Reference Value", "Target Value", "Delta Value", "Unit",
                "Evidence Dataset", "Evidence Filter", "Created At",
            ),
            """SELECT period_start, period_end, finding_id, finding_rank, finding_type,
                      severity, domain, metric_id, source_system, lob, language,
                      team_leader, agent_id, title, summary, current_value,
                      reference_value, target_value, delta_value, unit,
                      evidence_dataset, evidence_filter, created_at
               FROM mart.analysis_finding
               ORDER BY period_end DESC, finding_rank, finding_id""",
        ),
        (
            "FactSourceHealth.csv",
            (
                "Source Family", "Expected Path", "Newest File", "Newest Date",
                "Modified At", "Loaded At", "Rows", "Rejected", "Status", "Details",
            ),
            """SELECT source_family, expected_path, newest_file, newest_business_date,
                      modified_at, loaded_at, row_count, rejected_count, status, details
               FROM mart.source_health ORDER BY source_family""",
        ),
        (
            "FactQualityIssue.csv",
            (
                "Issue ID", "Run ID", "Detected At", "Source Family", "Source File",
                "Date", "Agent ID", "Issue Type", "Severity", "Details",
            ),
            """SELECT issue_id, run_id, detected_at, source_family, source_file,
                      business_date, agent_id, issue_type, severity, details
               FROM meta.quality_issue ORDER BY detected_at DESC, issue_id""",
        ),
    )
    for filename, headers, sql in specs:
        path, count = _query(conn, folder, filename, headers, sql)
        files.append(path)
        counts.append((path.name, count))

    setup_files = _copy_setup_assets(config, folder)
    files.extend(setup_files)
    manifest = _manifest(
        folder, "POWERBI", feed_start, feed_end, counts,
        schema_version=POWERBI_SCHEMA_VERSION,
        extra=(
            ("KPI authority", "WFMHub SQLite/Python", "Power BI aggregates governed counters"),
            ("Final ABS authority", "Verint Activities", "Attendance remains Agent Status based"),
            ("Forecast grain", "Native source interval", "Current supplied files are 15 minutes"),
        ),
    )
    files.append(manifest)
    return SharedFeedResult("POWERBI", tuple(files), sum(count for _, count in counts))
