"""Publish a fixed-schema Power BI star-model feed and its setup assets."""

from __future__ import annotations

import shutil
from datetime import date, datetime, time, timedelta
from pathlib import Path
from typing import Any, Sequence

from .config import Config
from .database import DatabaseConnection
from .service_profiles import load_service_profiles
from .shared_feeds import SharedFeedResult, _atomic_csv, _manifest


POWERBI_SCHEMA_VERSION = "2"


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


def _prepare_management_lob_maps(
    conn: DatabaseConnection,
    config: Config,
    feed_start: date,
    feed_end: date,
) -> None:
    """Install effective-dated report-only LOB bridges from service profiles.

    Service scopes and roster LOBs use different labels (for example RSA BE
    versus RSA FR/RSA VL).  The Power BI layer needs one explicit management
    member without changing either governed source label.  Temporary maps keep
    that presentation concern out of persistent marts.
    """

    catalog = load_service_profiles(config.home, config.service_profiles)
    rows: dict[
        tuple[str, str, date, date | None],
        tuple[str, str, date, date | None, str, int],
    ] = {}
    for profile in catalog.profiles:
        for kind, values in (
            ("SERVICE", profile.service_scopes),
            ("WORKFORCE", profile.staffing_lobs),
        ):
            for value in values:
                key = (kind, value.casefold(), profile.effective_from, profile.effective_to)
                rows[key] = (
                    kind, value, profile.effective_from, profile.effective_to,
                    profile.management_lob, profile.display_order,
                )
    ordered = sorted(rows.values(), key=lambda row: (row[0], row[1].casefold(), row[2]))
    for index, left in enumerate(ordered):
        for right in ordered[index + 1:]:
            if left[0] != right[0] or left[1].casefold() != right[1].casefold():
                continue
            if left[2] <= (right[3] or date.max) and right[2] <= (left[3] or date.max):
                if left[4] != right[4]:
                    raise ValueError(
                        f"Ambiguous Power BI management LOB map for {left[0]} "
                        f"{left[1]!r}: {left[4]!r} versus {right[4]!r}"
                    )
    conn.execute("DROP TABLE IF EXISTS pbi_management_lob_map")
    conn.execute(
        """CREATE TEMP TABLE pbi_management_lob_map (
               source_kind TEXT NOT NULL,
               source_value TEXT NOT NULL,
               effective_from DATE NOT NULL,
               effective_to DATE,
               management_lob TEXT NOT NULL,
               sort_order INTEGER NOT NULL
           )"""
    )
    conn.executemany(
        "INSERT INTO pbi_management_lob_map VALUES (?, ?, ?, ?, ?, ?)", ordered,
    )
    conn.execute(
        "CREATE INDEX pbi_management_lob_lookup "
        "ON pbi_management_lob_map(source_kind, source_value, effective_from, effective_to)"
    )
    conn.execute("DROP TABLE IF EXISTS pbi_feed_context")
    conn.execute(
        "CREATE TEMP TABLE pbi_feed_context (feed_start DATE, feed_end DATE)"
    )
    conn.execute(
        "INSERT INTO pbi_feed_context VALUES (?, ?)", [feed_start, feed_end],
    )


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
    _prepare_management_lob_maps(conn, config, feed_start, feed_end)
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
            ("Agent ID", "Agent", "Employment Status", "Team Leader", "Ops Manager", "LOB", "Management LOB", "Market", "Language", "Location", "City", "FTE"),
            """SELECT d.agent_id, d.canonical_name, d.employment_status,
                      d.team_leader, d.ops_manager, d.lob,
                      coalesce(m.management_lob,d.lob), d.market, d.language,
                      d.location, d.city, d.fte
               FROM core.dim_agent d
               LEFT JOIN pbi_management_lob_map m
                 ON m.source_kind='WORKFORCE'
                AND lower(trim(m.source_value))=lower(trim(d.lob))
                AND (SELECT feed_end FROM pbi_feed_context) >= m.effective_from
                AND (m.effective_to IS NULL OR (SELECT feed_end FROM pbi_feed_context) <= m.effective_to)
               ORDER BY coalesce(m.management_lob,d.lob), d.team_leader,
                        d.canonical_name, d.agent_id""",
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
            "DimManagementLOB.csv",
            ("Management LOB", "Sort Order"),
            """SELECT management_lob, min(sort_order)
               FROM (
                 SELECT m.management_lob, m.sort_order
                 FROM pbi_management_lob_map m, pbi_feed_context c
                 WHERE m.effective_from<=c.feed_end
                   AND (m.effective_to IS NULL OR m.effective_to>=c.feed_start)
                 UNION ALL
                 SELECT coalesce(m.management_lob,d.lob), coalesce(m.sort_order,100)
                 FROM core.dim_agent d
                 LEFT JOIN pbi_management_lob_map m
                   ON m.source_kind='WORKFORCE'
                  AND lower(trim(m.source_value))=lower(trim(d.lob))
                  AND (SELECT feed_end FROM pbi_feed_context)>=m.effective_from
                  AND (m.effective_to IS NULL OR (SELECT feed_end FROM pbi_feed_context)<=m.effective_to)
                 WHERE trim(coalesce(d.lob,''))<>''
               ) x
               WHERE trim(coalesce(management_lob,''))<>''
               GROUP BY management_lob ORDER BY min(sort_order), management_lob""",
        ),
        (
            "DimQueue.csv",
            ("Queue", "Source System", "Service Scope", "Comparison Scope", "Designation", "Mapping Status"),
            """SELECT queue, min(source_system), min(service_scope),
                      min(comparison_scope), min(designation), min(mapping_status)
               FROM mart.service_interval
               WHERE trim(coalesce(queue,''))<>''
               GROUP BY queue ORDER BY min(service_scope), queue""",
        ),
        (
            "FactServiceHour.csv",
            (
                "Date", "Hour Start", "Time Slot", "Source System", "Service Scope",
                "Management LOB", "Comparison Scope", "Queue", "Designation", "Offered",
                "Answered", "Abandoned", "Short Abandoned",
                "Abandoned Within Target", "Answered Within Target",
                "Handled Seconds", "SLA Denominator", "Source Files",
            ),
            """SELECT s.business_date, s.hour_start,
                      cast(strftime('%H',s.hour_start) AS INTEGER)*4,
                      s.source_system, s.service_scope,
                      coalesce(m.management_lob,s.service_scope),
                      s.comparison_scope, s.queue, s.designation,
                      sum(coalesce(s.offered,0)), sum(coalesce(s.answered,0)),
                      sum(coalesce(s.abandoned,0)), sum(coalesce(s.short_abandoned,0)),
                      sum(coalesce(s.abandoned_within_target,0)),
                      sum(coalesce(s.answered_within_target,0)),
                      sum(coalesce(s.handled_seconds,0)),
                      sum(coalesce(s.offered,0)-coalesce(s.abandoned_within_target,0)),
                      group_concat(DISTINCT s.source_file)
               FROM mart.service_interval s
               LEFT JOIN pbi_management_lob_map m
                 ON m.source_kind='SERVICE'
                AND lower(trim(m.source_value))=lower(trim(s.service_scope))
                AND s.business_date>=m.effective_from
                AND (m.effective_to IS NULL OR s.business_date<=m.effective_to)
               GROUP BY s.business_date, s.hour_start, s.source_system,
                        s.service_scope, coalesce(m.management_lob,s.service_scope),
                        s.comparison_scope, s.queue, s.designation
               ORDER BY s.business_date, s.hour_start, s.service_scope, s.queue""",
        ),
        (
            "FactForecastInterval.csv",
            (
                "Date", "Interval Start", "Interval End", "Interval Minutes",
                "Time Slot", "Queue", "Volume Forecast", "Abandons Forecast", "FTE Forecast",
                "FTE Required", "Headcount Forecast", "Net Staffing Forecast",
                "SL Forecast", "SL Required", "AHT Forecast Seconds",
                "Service Scope", "Management LOB", "Comparison Scope", "Mapping Status", "Source File",
            ),
            """SELECT f.business_date, f.interval_start, f.interval_end,
                      f.interval_minutes,
                      cast(strftime('%H',f.interval_start) AS INTEGER)*4
                        + cast(strftime('%M',f.interval_start) AS INTEGER)/15,
                      f.queue_name, f.volume_forecast, f.abandons_forecast,
                      f.fte_forecast, f.fte_required, f.headcount_forecast,
                      f.net_staffing_forecast, f.sl_forecast, f.sl_required,
                      f.aht_forecast_seconds, f.service_scope,
                      coalesce(m.management_lob,f.service_scope),
                      f.comparison_scope, f.mapping_status, f.source_file
               FROM mart.forecast_interval f
               LEFT JOIN pbi_management_lob_map m
                 ON m.source_kind='SERVICE'
                AND lower(trim(m.source_value))=lower(trim(f.service_scope))
                AND f.business_date>=m.effective_from
                AND (m.effective_to IS NULL OR f.business_date<=m.effective_to)
               ORDER BY f.business_date, f.interval_start, f.service_scope, f.queue_name""",
        ),
        (
            "FactStaffing15Min.csv",
            (
                "Date", "Interval Start", "Interval End", "Time Slot", "LOB",
                "Management LOB", "Language",
                "Gross Scheduled FTE", "Planned Time Off FTE", "Scheduled FTE",
                "Observed FTE", "Productive FTE", "Staffing Variance FTE",
                "Staffing Gap FTE", "Staffing State", "Evidence Basis", "Evaluation As Of",
            ),
            """SELECT s.business_date, s.interval_start, s.interval_end,
                      cast(strftime('%H',s.interval_start) AS INTEGER)*4
                        + cast(strftime('%M',s.interval_start) AS INTEGER)/15,
                      s.lob, coalesce(m.management_lob,s.lob), s.language,
                      s.gross_scheduled_fte, s.planned_time_off_fte,
                      s.scheduled_fte, s.observed_fte, s.productive_fte,
                      s.staffing_variance_fte, s.staffing_gap_fte,
                      s.staffing_state, s.evidence_basis, s.evaluation_as_of
               FROM mart.staffing_interval s
               LEFT JOIN pbi_management_lob_map m
                 ON m.source_kind='WORKFORCE'
                AND lower(trim(m.source_value))=lower(trim(s.lob))
                AND s.business_date>=m.effective_from
                AND (m.effective_to IS NULL OR s.business_date<=m.effective_to)
               ORDER BY s.business_date, s.interval_start, s.lob, s.language""",
        ),
        (
            "FactAttendanceDay.csv",
            (
                "Date", "Agent Day Key", "Agent ID", "Agent", "Team Leader",
                "Ops Manager", "LOB", "Management LOB", "Market", "Language", "Location",
                "Scheduled Start", "Scheduled End", "Scheduled Minutes",
                "Planned Work Minutes", "Planning Overlay", "Planning Overlay Minutes",
                "First Login", "Last Logout", "Attendance Result", "Call Action",
                "Requires Call", "Shift State", "Actual Evidence", "Is Provisional",
                "Late Minutes", "Early Leave Minutes", "No Show Minutes",
                "Status Covered Minutes", "Evaluation As Of",
            ),
            """SELECT a.business_date, a.agent_day_key, a.agent_id, a.agent_name,
                      a.team_leader, a.ops_manager, a.lob,
                      coalesce(m.management_lob,a.lob), a.market, a.language,
                      a.location, a.scheduled_start, a.scheduled_end,
                      a.scheduled_minutes, a.planned_work_minutes,
                      a.planning_overlay, a.planning_overlay_minutes,
                      a.first_login, a.last_logout, a.attendance_result,
                      a.call_action, a.requires_call, a.shift_state,
                      a.actual_evidence, a.is_provisional, a.uncoded_late_minutes,
                      a.uncoded_early_leave_minutes, a.no_show_minutes,
                      a.status_covered_minutes, a.evaluation_as_of
               FROM mart.attendance_agent_day a
               LEFT JOIN pbi_management_lob_map m
                 ON m.source_kind='WORKFORCE'
                AND lower(trim(m.source_value))=lower(trim(a.lob))
                AND a.business_date>=m.effective_from
                AND (m.effective_to IS NULL OR a.business_date<=m.effective_to)
               ORDER BY a.business_date, a.lob, a.team_leader, a.agent_name, a.agent_id""",
        ),
        (
            "FactAttendanceGap.csv",
            (
                "Date", "Correction ID", "Agent ID", "Agent", "Team Leader",
                "Ops Manager", "LOB", "Management LOB", "Scheduled Start", "Scheduled End",
                "Gap Start", "Gap End", "Gap Minutes", "Detected Issue",
                "Priority", "Confidence", "Suggested Activity", "Observed Source",
                "Reconciliation", "Source File",
            ),
            """SELECT c.business_date, c.correction_id, c.agent_id, c.agent_name,
                      c.team_leader, c.ops_manager, c.lob,
                      coalesce(m.management_lob,c.lob), c.scheduled_start,
                      c.scheduled_end, c.gap_start, c.gap_end, c.gap_minutes,
                      c.detected_issue, c.priority, c.confidence,
                      c.suggested_activity, c.observed_source,
                      c.verint_reconciliation, c.source_file
               FROM mart.correction_candidate c
               LEFT JOIN pbi_management_lob_map m
                 ON m.source_kind='WORKFORCE'
                AND lower(trim(m.source_value))=lower(trim(c.lob))
                AND c.business_date>=m.effective_from
                AND (m.effective_to IS NULL OR c.business_date<=m.effective_to)
               ORDER BY c.business_date, c.lob, c.agent_name, c.gap_start""",
        ),
        (
            "FactPCSAgentDay.csv",
            (
                "Date", "Agent Day Key", "Agent ID", "Agent", "Team Leader",
                "Ops Manager", "LOB", "Management LOB", "Market", "Language", "Location",
                "Call Legs", "Handled Calls", "Inbound Calls", "Outbound Calls",
                "Handle Seconds", "PCS Score Sum", "Valid PCS", "PCS Status Calls",
                "PCS Participation Responses", "Low Score Responses",
                "Top Box Responses", "Comments Count",
            ),
            """SELECT p.business_date, p.agent_day_key, p.agent_id, p.agent_name,
                      p.team_leader, p.ops_manager, p.lob,
                      coalesce(m.management_lob,p.lob), p.market, p.language,
                      p.location, p.call_legs, p.handled_calls, p.inbound_calls,
                      p.outbound_calls, p.handle_seconds, p.pcs_score_sum,
                      p.survey_responses, p.pcs_status_calls,
                      p.pcs_participation_responses, p.low_score_responses,
                      p.top_box_responses, p.comments_count
               FROM mart.agent_pcs_day p
               LEFT JOIN pbi_management_lob_map m
                 ON m.source_kind='WORKFORCE'
                AND lower(trim(m.source_value))=lower(trim(p.lob))
                AND p.business_date>=m.effective_from
                AND (m.effective_to IS NULL OR p.business_date<=m.effective_to)
               ORDER BY p.business_date, p.lob, p.team_leader, p.agent_name, p.agent_id""",
        ),
        (
            "FactPCSCoaching.csv",
            (
                "Date", "Call Start", "Agent ID", "Agent", "Team Leader", "LOB", "Management LOB",
                "Call ID", "Call Key", "Q1 Score", "Customer Comment", "PCS Status",
            ),
            """SELECT c.business_date, c.call_start, c.agent_id,
                      coalesce(a.canonical_name,c.agent_name), a.team_leader, a.lob,
                      coalesce(m.management_lob,a.lob), c.call_id, c.call_key,
                      c.question_1_score,
                      coalesce(c.question_2,c.question_3,c.question_4,c.question_5,
                               c.question_6,c.question_7,c.question_8,c.question_9,
                               c.question_10), c.pcs_status
               FROM core.clean_call_leg c
               LEFT JOIN core.dim_agent a ON a.agent_id=c.agent_id
               LEFT JOIN pbi_management_lob_map m
                 ON m.source_kind='WORKFORCE'
                AND lower(trim(m.source_value))=lower(trim(a.lob))
                AND c.business_date>=m.effective_from
                AND (m.effective_to IS NULL OR c.business_date<=m.effective_to)
               WHERE c.question_1_score IS NOT NULL AND c.question_1_score<=3
               ORDER BY c.business_date DESC, c.question_1_score, a.lob,
                        a.team_leader, a.canonical_name""",
        ),
        (
            "FactTimeOff.csv",
            (
                "Date", "Segment Key", "Agent Day Key", "Agent ID", "Agent",
                "Team Leader", "Ops Manager", "LOB", "Management LOB", "Language", "Source Kind",
                "Absence Type", "Record Status", "Segment Start", "Segment End",
                "Planned Minutes", "Source File", "Source Sheet", "Source Row",
            ),
            """SELECT p.business_date, p.segment_key, p.agent_day_key, p.agent_id,
                      p.agent_name, p.team_leader, p.ops_manager, p.lob,
                      coalesce(m.management_lob,p.lob), p.language,
                      p.source_kind, p.absence_type, p.record_status,
                      p.segment_start, p.segment_end, p.planned_minutes,
                      p.source_file, p.source_sheet, p.source_row
               FROM mart.planned_time_off_segment p
               LEFT JOIN pbi_management_lob_map m
                 ON m.source_kind='WORKFORCE'
                AND lower(trim(m.source_value))=lower(trim(p.lob))
                AND p.business_date>=m.effective_from
                AND (m.effective_to IS NULL OR p.business_date<=m.effective_to)
               ORDER BY p.business_date, p.lob, p.team_leader, p.agent_name, p.segment_start""",
        ),
        (
            "FactFinalAbsenceDay.csv",
            (
                "Date", "Agent Day Key", "Agent ID", "Agent", "Team Leader",
                "Ops Manager", "LOB", "Management LOB", "Market", "Language", "Location",
                "Scheduled Minutes", "Planned Net Minutes", "Absence Minutes",
                "Vacation Minutes", "Unpaid Minutes", "Shrinkage Minutes",
                "Unmapped Minutes", "Final Ledger Status", "Rule Version",
            ),
            """SELECT a.business_date, a.agent_day_key, a.agent_id, a.agent_name,
                      a.team_leader, a.ops_manager, a.lob,
                      coalesce(m.management_lob,a.lob), a.market, a.language,
                      a.location, a.scheduled_minutes, a.planned_net_minutes,
                      a.final_absence_minutes, a.final_vacation_minutes,
                      a.final_unpaid_minutes, a.final_shrinkage_minutes,
                      a.final_unmapped_minutes, a.final_ledger_status,
                      a.rule_version
               FROM mart.verint_final_absence_agent_day a
               LEFT JOIN pbi_management_lob_map m
                 ON m.source_kind='WORKFORCE'
                AND lower(trim(m.source_value))=lower(trim(a.lob))
                AND a.business_date>=m.effective_from
                AND (m.effective_to IS NULL OR a.business_date<=m.effective_to)
               ORDER BY a.business_date, a.lob, a.team_leader, a.agent_name, a.agent_id""",
        ),
        (
            "FactFinalAbsenceComponent.csv",
            (
                "Date", "Event Key", "Agent Day Key", "Agent ID", "Agent",
                "Team Leader", "LOB", "Management LOB", "Activity", "Category", "Event Start",
                "Event End", "Minutes", "Counts As Absence", "Counts As Vacation",
                "Counts As Unpaid", "Counts As Shrinkage", "Mapped",
                "Evidence Type", "Source File", "Rule Version",
            ),
            """SELECT e.business_date, e.event_key, e.agent_day_key, e.agent_id,
                      e.agent_name, e.team_leader, e.lob,
                      coalesce(m.management_lob,e.lob), e.activity, e.category,
                      e.event_start, e.event_end, e.minutes, e.counts_as_absence,
                      e.counts_as_vacation, e.counts_as_unpaid,
                      e.counts_as_shrinkage, e.mapped, e.evidence_type,
                      e.source_file, e.rule_version
               FROM mart.verint_final_absence_event e
               LEFT JOIN pbi_management_lob_map m
                 ON m.source_kind='WORKFORCE'
                AND lower(trim(m.source_value))=lower(trim(e.lob))
                AND e.business_date>=m.effective_from
                AND (m.effective_to IS NULL OR e.business_date<=m.effective_to)
               ORDER BY e.business_date, e.lob, e.agent_name, e.event_start""",
        ),
        (
            "FactBonusMonth.csv",
            (
                "Period", "Agent ID", "Agent", "Team Leader", "Ops Manager",
                "Management LOB", "Population", "Core Ready", "Eligibility", "AHT Earned",
                "Productivity Earned", "PCS Earned", "Participation Earned",
                "QM Earned", "Absence Earned", "Extra PCS Earned",
                "Gross Achievement", "VOC Malus", "Final Achievement",
                "Reference Bonus", "Proration", "Scenario Payout",
                "Released Payout", "Release Status", "Data Issue", "Import ID",
            ),
            """SELECT b.period, b.agent_id, b.agent_name, b.team_leader,
                      b.ops_manager, coalesce(m.management_lob,d.lob),
                      b.population, b.core_ready, b.eligibility, b.aht_earned,
                      b.productivity_earned, b.pcs_earned,
                      b.participation_earned, b.qm_earned, b.absence_earned,
                      b.extra_pcs_earned, b.gross_achievement, b.voc_malus,
                      b.final_achievement, b.reference_bonus, b.proration,
                      b.scenario_payout, b.released_payout, b.release_status,
                      b.data_issue, b.import_id
               FROM mart.bonus_agent_month b
               LEFT JOIN core.dim_agent d ON d.agent_id=b.agent_id
               LEFT JOIN pbi_management_lob_map m
                 ON m.source_kind='WORKFORCE'
                AND lower(trim(m.source_value))=lower(trim(d.lob))
                AND date(b.period)>=m.effective_from
                AND (m.effective_to IS NULL OR date(b.period)<=m.effective_to)
               ORDER BY b.period, b.population, b.team_leader, b.agent_name, b.agent_id""",
        ),
        (
            "FactBonusKPI.csv",
            (
                "Period", "Agent ID", "Agent", "Management LOB", "Population", "KPI",
                "Actual Value", "Earned Weight", "Direction", "Tier 1 Target",
                "Tier 2 Target", "Import ID",
            ),
            """SELECT b.period, b.agent_id, b.agent_name,
                      coalesce(m.management_lob,d.lob), b.population, b.kpi,
                      b.actual_value, b.earned_weight, b.direction,
                      b.tier1_target, b.tier2_target, b.import_id
               FROM mart.bonus_kpi_result b
               LEFT JOIN core.dim_agent d ON d.agent_id=b.agent_id
               LEFT JOIN pbi_management_lob_map m
                 ON m.source_kind='WORKFORCE'
                AND lower(trim(m.source_value))=lower(trim(d.lob))
                AND date(b.period)>=m.effective_from
                AND (m.effective_to IS NULL OR date(b.period)<=m.effective_to)
               ORDER BY b.period, b.population, b.agent_name, b.agent_id, b.kpi""",
        ),
        (
            "FactMetricValue.csv",
            (
                "Date", "Interval Start", "Metric Key", "Metric ID", "Method ID",
                "Domain", "Unit", "Aggregation", "Source System", "LOB", "Management LOB", "Language",
                "Team Leader", "Agent ID", "Numerator", "Denominator", "Sample Size",
                "Metric Value", "Target Value", "Metric State", "Catalog Version",
            ),
            """SELECT v.business_date, v.interval_start, v.metric_key,
                      v.metric_id, v.method_id, v.domain, v.unit, v.aggregation,
                      v.source_system, v.lob,
                      coalesce(sm.management_lob,wm.management_lob,v.lob),
                      v.language, v.team_leader, v.agent_id, v.numerator,
                      v.denominator, v.sample_size, v.metric_value,
                      v.target_value, v.metric_state, v.catalog_version
               FROM mart.metric_value v
               LEFT JOIN pbi_management_lob_map sm
                 ON sm.source_kind='SERVICE'
                AND lower(trim(sm.source_value))=lower(trim(v.lob))
                AND v.business_date>=sm.effective_from
                AND (sm.effective_to IS NULL OR v.business_date<=sm.effective_to)
               LEFT JOIN pbi_management_lob_map wm
                 ON wm.source_kind='WORKFORCE'
                AND lower(trim(wm.source_value))=lower(trim(v.lob))
                AND v.business_date>=wm.effective_from
                AND (wm.effective_to IS NULL OR v.business_date<=wm.effective_to)
               ORDER BY v.business_date, v.domain, v.metric_id, v.lob,
                        v.team_leader, v.agent_id""",
        ),
        (
            "FactFinding.csv",
            (
                "Period Start", "Period End", "Finding ID", "Rank", "Finding Type",
                "Severity", "Domain", "Metric ID", "Source System", "LOB", "Management LOB", "Language",
                "Team Leader", "Agent ID", "Title", "Summary", "Current Value",
                "Reference Value", "Target Value", "Delta Value", "Unit",
                "Evidence Dataset", "Evidence Filter", "Created At",
            ),
            """SELECT f.period_start, f.period_end, f.finding_id,
                      f.finding_rank, f.finding_type, f.severity, f.domain,
                      f.metric_id, f.source_system, f.lob,
                      coalesce(sm.management_lob,wm.management_lob,f.lob),
                      f.language, f.team_leader, f.agent_id, f.title, f.summary,
                      f.current_value, f.reference_value, f.target_value,
                      f.delta_value, f.unit, f.evidence_dataset,
                      f.evidence_filter, f.created_at
               FROM mart.analysis_finding f
               LEFT JOIN pbi_management_lob_map sm
                 ON sm.source_kind='SERVICE'
                AND lower(trim(sm.source_value))=lower(trim(f.lob))
                AND f.period_end>=sm.effective_from
                AND (sm.effective_to IS NULL OR f.period_end<=sm.effective_to)
               LEFT JOIN pbi_management_lob_map wm
                 ON wm.source_kind='WORKFORCE'
                AND lower(trim(wm.source_value))=lower(trim(f.lob))
                AND f.period_end>=wm.effective_from
                AND (wm.effective_to IS NULL OR f.period_end<=wm.effective_to)
               ORDER BY f.period_end DESC, f.finding_rank, f.finding_id""",
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
