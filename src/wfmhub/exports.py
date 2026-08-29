"""Stream curated, cleaned hub datasets to CSV or bounded XLSX files."""

from __future__ import annotations

import csv
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Any

import xlsxwriter

from .config import Config
from .database import DatabaseConnection
from .progress import ProgressCallback
from .rules import load_rulebook


@dataclass(frozen=True)
class ExportDataset:
    key: str
    description: str
    sql: str
    dated: bool = True


@dataclass(frozen=True)
class ExportResult:
    dataset: str
    path: Path
    manifest: Path
    rows: int


DATASETS: dict[str, ExportDataset] = {
    "calls": ExportDataset(
        "calls", "Deduplicated, typed and FTE-scoped call legs.",
        "SELECT * FROM core.clean_call_leg WHERE business_date BETWEEN ? AND ? ORDER BY business_date, call_start, agent_id",
    ),
    "pcs_responses": ExportDataset(
        "pcs_responses", "Clean call legs containing at least one PCS answer or comment.",
        """SELECT * FROM core.clean_call_leg
           WHERE business_date BETWEEN ? AND ? AND (
               coalesce(question_1,'')<>'' OR coalesce(question_2,'')<>'' OR
               coalesce(question_3,'')<>'' OR coalesce(question_4,'')<>'' OR
               coalesce(question_5,'')<>'' OR coalesce(question_6,'')<>'' OR
               coalesce(question_7,'')<>'' OR coalesce(question_8,'')<>'' OR
               coalesce(question_9,'')<>'' OR coalesce(question_10,'')<>''
           ) ORDER BY business_date, call_start, agent_id""",
    ),
    "pcs_agent_day": ExportDataset(
        "pcs_agent_day", "Exact PCS workbook counters and call performance per agent/day.",
        """SELECT business_date, agent_id, agent_name, team_leader, ops_manager,
                  lob, market, language, location, call_legs, handled_calls,
                  inbound_calls, outbound_calls, transferred_legs, talk_seconds,
                  hold_seconds, wrap_seconds, handle_seconds, average_handle_seconds,
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
                  substr(business_date,1,7) AS month_key
           FROM mart.agent_pcs_day WHERE business_date BETWEEN ? AND ?
           ORDER BY business_date, agent_id""",
    ),
    "pcs_team_day": ExportDataset(
        "pcs_team_day", "PCS team/day ratios recalculated from summed counters.",
        """SELECT business_date, coalesce(team_leader,'(blank)') AS team_leader,
                  coalesce(lob,'(blank)') AS lob, coalesce(language,'(blank)') AS language,
                  sum(inbound_calls) AS inbound_call_legs,
                  sum(pcs_enabled_calls) AS pcs_mode_2_inbound_legs,
                  sum(pcs_status_calls) AS pcs_status_1_inbound_legs,
                  sum(pcs_participation_responses) AS pcs_q1_nonblank_inbound_legs,
                  sum(survey_responses) AS pcs_q1_valid_score_count,
                  sum(pcs_score_sum) AS pcs_q1_score_sum,
                  sum(low_score_responses) AS pcs_score_le_3_count,
                  sum(top_box_responses) AS pcs_score_gt_3_count,
                  sum(pcs_invalid_responses) AS pcs_q1_invalid_nonblank_count,
                  CASE WHEN sum(survey_responses)>0
                       THEN 1.0*sum(pcs_score_sum)/sum(survey_responses) END AS pcs_average,
                  CASE WHEN sum(pcs_status_calls)>0
                       THEN 1.0*sum(pcs_participation_responses)/sum(pcs_status_calls) END AS pcs_participation_rate
           FROM mart.agent_pcs_day WHERE business_date BETWEEN ? AND ?
           GROUP BY business_date, coalesce(team_leader,'(blank)'),
                    coalesce(lob,'(blank)'), coalesce(language,'(blank)')
           ORDER BY business_date, team_leader, lob, language""",
    ),
    "pcs_agent_month": ExportDataset(
        "pcs_agent_month", "PCS monthly agent totals using the workbook's ratio-of-sums logic.",
        """SELECT substr(business_date,1,7) AS month_key, agent_id,
                  max(agent_name) AS agent_name, max(team_leader) AS team_leader,
                  max(ops_manager) AS ops_manager, max(lob) AS lob,
                  max(language) AS language, sum(inbound_calls) AS inbound_call_legs,
                  sum(pcs_status_calls) AS pcs_status_1_inbound_legs,
                  sum(pcs_participation_responses) AS pcs_q1_nonblank_inbound_legs,
                  sum(survey_responses) AS pcs_q1_valid_score_count,
                  sum(pcs_score_sum) AS pcs_q1_score_sum,
                  sum(low_score_responses) AS pcs_score_le_3_count,
                  sum(top_box_responses) AS pcs_score_gt_3_count,
                  sum(pcs_invalid_responses) AS pcs_q1_invalid_nonblank_count,
                  CASE WHEN sum(survey_responses)>0
                       THEN 1.0*sum(pcs_score_sum)/sum(survey_responses) END AS pcs_average,
                  CASE WHEN sum(pcs_status_calls)>0
                       THEN 1.0*sum(pcs_participation_responses)/sum(pcs_status_calls) END AS pcs_participation_rate
           FROM mart.agent_pcs_day WHERE business_date BETWEEN ? AND ?
           GROUP BY substr(business_date,1,7), agent_id
           ORDER BY month_key, agent_id""",
    ),
    "attendance": ExportDataset(
        "attendance", "Attendance agent/day results without adherence metrics.",
        """SELECT * FROM mart.attendance_agent_day
           WHERE business_date BETWEEN ? AND ? ORDER BY business_date, agent_id""",
    ),
    "daily_attendance_calls": ExportDataset(
        "daily_attendance_calls", "Absent, not-seen and late agents requiring an attendance call.",
        """SELECT business_date, agent_id, agent_name, team_leader, ops_manager,
                  lob, language, scheduled_start, scheduled_end, shift_state,
                  attendance_result, call_action, requires_call, is_provisional,
                  actual_first_seen, actual_last_seen, actual_evidence,
                  uncoded_late_minutes, no_show_minutes, source_loaded,
                  schedule_source, lilo_source, status_source, evaluation_as_of
           FROM mart.attendance_agent_day
           WHERE business_date BETWEEN ? AND ? AND requires_call=true
           ORDER BY business_date, scheduled_start, lob, language, agent_name""",
    ),
    "daily_staffing_gaps": ExportDataset(
        "daily_staffing_gaps", "15-minute scheduled versus observed staffing by roster LOB and language.",
        """SELECT * FROM mart.staffing_interval
           WHERE business_date BETWEEN ? AND ?
           ORDER BY business_date, interval_start, lob, language""",
    ),
    "yesterday_gap_actions": ExportDataset(
        "yesterday_gap_actions", "Uncovered observed gap segments ready for correction review and Verint injection.",
        """SELECT r.business_date, r.agent_id, c.agent_name, c.team_leader,
                  c.ops_manager, c.lob, c.scheduled_start, c.scheduled_end,
                  c.detected_issue, r.residual_start AS gap_start,
                  r.residual_end AS gap_end, r.residual_minutes AS gap_minutes,
                  c.confidence, r.suggested_activity, r.observed_source,
                  r.verint_reconciliation, c.verint_activity,
                  c.verint_overlap_minutes, c.validation_status,
                  c.owner, c.comment, c.injected_date, r.source_file
           FROM mart.correction_residual_segment r
           JOIN mart.correction_candidate c ON c.correction_id=r.correction_id
           WHERE r.business_date BETWEEN ? AND ?
             AND coalesce(c.validation_status,'Open') NOT IN ('Injected','Rejected')
           ORDER BY r.business_date, r.agent_id, r.residual_start""",
    ),
    "shift_evidence_timeline": ExportDataset(
        "shift_evidence_timeline", "Full-shift planned-versus-observed timeline segments for visual review.",
        """SELECT * FROM mart.shift_timeline_segment
           WHERE business_date BETWEEN ? AND ?
           ORDER BY business_date, agent_id, segment_start""",
    ),
    "absence_agent_day": ExportDataset(
        "absence_agent_day", "Rule-versioned payroll absence, vacation and shrinkage per agent/day.",
        "SELECT * FROM mart.absence_agent_day WHERE business_date BETWEEN ? AND ? ORDER BY business_date, agent_id",
    ),
    "absence_events": ExportDataset(
        "absence_events", "Observed LILO/status gaps with Verint-final reconciliation.",
        "SELECT * FROM mart.absence_event WHERE business_date BETWEEN ? AND ? ORDER BY business_date, agent_id, event_start",
    ),
    "gaps": ExportDataset(
        "gaps", "Observed correction candidates, Verint checks and saved decisions.",
        "SELECT * FROM mart.correction_candidate WHERE business_date BETWEEN ? AND ? ORDER BY business_date, agent_id, gap_start",
    ),
    "verint_final_exceptions": ExportDataset(
        "verint_final_exceptions", "Final Verint activities with no supporting observed gap.",
        "SELECT * FROM mart.verint_final_exception WHERE business_date BETWEEN ? AND ? ORDER BY business_date, agent_id, event_start",
    ),
    "verint_final_absence_events": ExportDataset(
        "verint_final_absence_events", "Corrected Verint Activities-only final absence event ledger.",
        """SELECT * FROM mart.verint_final_absence_event
           WHERE business_date BETWEEN ? AND ? ORDER BY business_date, agent_id, event_start""",
    ),
    "verint_final_absence_day": ExportDataset(
        "verint_final_absence_day", "Corrected Verint Activities-only final absence per agent/day.",
        """SELECT * FROM mart.verint_final_absence_agent_day
           WHERE business_date BETWEEN ? AND ? ORDER BY business_date, agent_id""",
    ),
    "schedules": ExportDataset(
        "schedules", "Parsed, FTE-scoped schedule shifts.",
        """SELECT source_file_id, source_row, schedule_date, agent_id_raw,
                  agent_id, agent_name, scheduling_period, shift_assignment,
                  assignment, assignment_type, scheduled_start, scheduled_end,
                  shift_events, parse_ok, source_variant, source_file
           FROM (
               SELECT r.*, f.source_variant, f.file_name AS source_file,
                      row_number() OVER (
                          PARTITION BY r.schedule_date, r.agent_id,
                                       r.scheduled_start, r.scheduled_end,
                                       coalesce(r.assignment,'')
                          ORDER BY f.modified_at DESC, f.loaded_at DESC, r.source_row DESC
                      ) row_rank
               FROM raw.schedule_shift r JOIN meta.source_file f ON f.file_id=r.source_file_id
               WHERE f.active=true AND f.status='SUCCESS'
                 AND f.source_variant='START_END' AND r.schedule_date BETWEEN ? AND ?
           ) x WHERE row_rank=1
           ORDER BY schedule_date, agent_id, scheduled_start""",
    ),
    "events": ExportDataset(
        "events", "Parsed Verint schedule activity intervals.",
        """SELECT source_file_id, source_row, event_index, schedule_date,
                  agent_id, agent_name, activity, activity_type, event_start,
                  event_end, parse_ok, source_variant, source_file
           FROM (
               SELECT r.*, f.source_variant, f.file_name AS source_file,
                      row_number() OVER (
                          PARTITION BY r.schedule_date, r.agent_id,
                                       upper(coalesce(r.activity,'')), r.event_start, r.event_end
                          ORDER BY f.modified_at DESC, f.loaded_at DESC, r.source_row DESC
                      ) row_rank
               FROM raw.schedule_event r JOIN meta.source_file f ON f.file_id=r.source_file_id
               WHERE f.active=true AND f.status='SUCCESS'
                 AND f.source_variant='ACTIVITIES' AND r.schedule_date BETWEEN ? AND ?
           ) x WHERE row_rank=1
           ORDER BY schedule_date, agent_id, event_start""",
    ),
    "lilo": ExportDataset(
        "lilo", "Parsed and FTE-scoped LILO rows using row-level business dates.",
        """SELECT source_file_id, source_row, extract_date, agent_id,
                  agent_name, first_login, raw_last_logout, last_logout,
                  overnight_adjusted, source_file
           FROM (
               SELECT r.*, f.file_name AS source_file,
                      row_number() OVER (
                          PARTITION BY r.extract_date, r.agent_id,
                                       r.first_login, r.raw_last_logout
                          ORDER BY f.modified_at DESC, f.loaded_at DESC, r.source_row DESC
                      ) row_rank
               FROM raw.lilo r JOIN meta.source_file f ON f.file_id=r.source_file_id
               WHERE f.active=true AND f.status='SUCCESS' AND r.extract_date BETWEEN ? AND ?
           ) x WHERE row_rank=1 ORDER BY extract_date, agent_id""",
    ),
    "agent_status": ExportDataset(
        "agent_status", "Parsed and FTE-scoped Agent Status intervals.",
        """SELECT source_file_id, source_row, serial_number, extract_date,
                  agent_id, agent_name, status, actual_category, status_start,
                  status_end, duration_seconds, queue, source_file
           FROM (
               SELECT r.*, f.file_name AS source_file,
                      row_number() OVER (PARTITION BY r.serial_number ORDER BY f.modified_at DESC, f.loaded_at DESC, r.source_row DESC) row_rank
               FROM raw.agent_status r JOIN meta.source_file f ON f.file_id=r.source_file_id
               WHERE f.active=true AND f.status='SUCCESS' AND r.extract_date BETWEEN ? AND ?
           ) x WHERE row_rank=1 ORDER BY extract_date, agent_id, status_start""",
    ),
    "intraday_actual": ExportDataset(
        "intraday_actual", "Legacy clean Storm APBE/APFR/APDE queue intervals.",
        "SELECT * FROM mart.intraday_queue_interval WHERE business_date BETWEEN ? AND ? ORDER BY business_date, interval_start, source_system, queue",
    ),
    "service_actual": ExportDataset(
        "service_actual", "Rule-versioned service performance; availability means answered/offered.",
        "SELECT * FROM mart.service_interval WHERE business_date BETWEEN ? AND ? ORDER BY business_date, interval_start, source_system, queue",
    ),
    "daily_service_lob": ExportDataset(
        "daily_service_lob", "APDE intraday service state recalculated from summed LOB/language counters.",
        """SELECT business_date, interval_start, coalesce(lob,'(blank)') AS lob,
                  coalesce(language,'(blank)') AS language,
                  sum(offered) AS offered, sum(answered) AS answered,
                  sum(abandoned) AS abandoned, sum(short_abandoned) AS short_abandoned,
                  sum(answered_within_target) AS answered_within_target,
                  sum(handled_seconds) AS handled_seconds,
                  CASE WHEN sum(offered)-sum(coalesce(short_abandoned,0))>0
                       THEN 1.0*sum(answered_within_target)/(sum(offered)-sum(coalesce(short_abandoned,0))) END AS service_level,
                  CASE WHEN sum(offered)>0 THEN 1.0*sum(answered)/sum(offered) END AS service_availability,
                  max(sl_target) AS sl_target,
                  CASE WHEN sum(offered)=0 OR sum(offered)-sum(coalesce(short_abandoned,0))<=0 THEN 'NO_TRAFFIC'
                       WHEN 1.0*sum(answered_within_target)/(sum(offered)-sum(coalesce(short_abandoned,0))) >= max(sl_target)
                       THEN 'ON_TARGET' ELSE 'BELOW_TARGET' END AS sl_state,
                  max(mapping_status) AS mapping_status
           FROM mart.service_interval
           WHERE source_system='APDE' AND business_date BETWEEN ? AND ?
           GROUP BY business_date, interval_start, coalesce(lob,'(blank)'), coalesce(language,'(blank)')
           ORDER BY business_date, interval_start, lob, language""",
    ),
    "forecast": ExportDataset(
        "forecast", "Clean Verint forecast and required staffing hours.",
        "SELECT * FROM mart.forecast_hour WHERE business_date BETWEEN ? AND ? ORDER BY business_date, hour_start, queue_name",
    ),
    "source_health": ExportDataset(
        "source_health", "Current source coverage, counts and load status.",
        "SELECT * FROM mart.source_health ORDER BY source_family", dated=False,
    ),
}


def _safe_name(value: str) -> str:
    return "".join(char if char.isalnum() or char in {"-", "_"} else "_" for char in value).strip("_")


def _write_csv(
    cursor,
    path: Path,
    dataset_key: str,
    progress: ProgressCallback | None = None,
) -> int:
    partial = path.with_name(path.name + ".partial")
    headers = [item[0] for item in cursor.description]
    count = 0
    try:
        with partial.open("w", encoding="utf-8-sig", newline="") as handle:
            writer = csv.writer(handle)
            writer.writerow(headers)
            while True:
                rows = cursor.fetchmany(5000)
                if not rows:
                    break
                writer.writerows(rows)
                count += len(rows)
                if progress is not None:
                    progress(count, 0, f"Exporting {dataset_key}: {count:,} rows")
        partial.replace(path)
    except Exception:
        partial.unlink(missing_ok=True)
        raise
    return count


def _write_xlsx(
    cursor,
    path: Path,
    dataset_key: str,
    progress: ProgressCallback | None = None,
) -> int:
    partial = path.with_name(f"{path.stem}.partial{path.suffix}")
    headers = [item[0] for item in cursor.description]
    workbook = xlsxwriter.Workbook(partial, {"constant_memory": True})
    worksheet = workbook.add_worksheet("DATA")
    header = workbook.add_format({"bold": True, "bg_color": "#D9EAF7", "border": 1})
    date_fmt = workbook.add_format({"num_format": "yyyy-mm-dd"})
    datetime_fmt = workbook.add_format({"num_format": "yyyy-mm-dd hh:mm:ss"})
    for column, value in enumerate(headers):
        worksheet.write(0, column, value, header)
        worksheet.set_column(column, column, min(38, max(12, len(value) + 2)))
    count = 0
    try:
        while True:
            rows = cursor.fetchmany(5000)
            if not rows:
                break
            for values in rows:
                if count >= 1_048_575:
                    raise ValueError("XLSX row limit reached; export this dataset as CSV")
                for column, value in enumerate(values):
                    fmt = datetime_fmt if isinstance(value, datetime) else date_fmt if isinstance(value, date) else None
                    worksheet.write(count + 1, column, value, fmt)
                count += 1
            if progress is not None:
                progress(count, 0, f"Exporting {dataset_key}: {count:,} rows")
        worksheet.freeze_panes(1, 0)
        worksheet.autofilter(0, 0, max(0, count), max(0, len(headers) - 1))
        workbook.close()
        partial.replace(path)
    except Exception:
        try:
            workbook.close()
        finally:
            partial.unlink(missing_ok=True)
        raise
    return count


def export_dataset(
    conn: DatabaseConnection,
    config: Config,
    dataset_key: str,
    start: date,
    end: date,
    file_format: str = "csv",
    output: Path | None = None,
    progress: ProgressCallback | None = None,
) -> ExportResult:
    try:
        dataset = DATASETS[dataset_key]
    except KeyError as exc:
        raise ValueError(f"Unknown export dataset {dataset_key!r}. Available: {', '.join(DATASETS)}") from exc
    file_format = file_format.lower()
    if file_format not in {"csv", "xlsx"}:
        raise ValueError("Export format must be csv or xlsx")
    stamp = datetime.now().strftime("%H%M%S_%f")
    period = f"{start:%Y-%m-%d}_to_{end:%Y-%m-%d}" if dataset.dated else "current"
    output = (
        output
        or config.output / "data_exports" / dataset.key
        / f"{_safe_name(dataset.key)}_{period}_{stamp}.{file_format}"
    ).resolve()
    output.parent.mkdir(parents=True, exist_ok=True)
    if progress is not None:
        progress(0, 0, f"Preparing {dataset.key} export")
    cursor = conn.execute(dataset.sql, [start, end] if dataset.dated else [])
    rows = (
        _write_csv(cursor, output, dataset.key, progress)
        if file_format == "csv"
        else _write_xlsx(cursor, output, dataset.key, progress)
    )
    if progress is not None:
        progress(rows, 0, f"Exported {dataset.key}: {rows:,} rows")
    manifest = output.with_name(output.name + ".manifest.txt")
    rulebook = load_rulebook(config.home, config.business_rules)
    manifest.write_text(
        "\n".join([
            "WFMHub clean-data export",
            f"Dataset: {dataset.key}",
            f"Description: {dataset.description}",
            f"Period: {start.isoformat()} to {end.isoformat()}",
            f"Rows: {rows}",
            f"Generated: {datetime.now().isoformat(timespec='seconds')}",
            f"Format: {file_format.upper()}",
            f"Rule version: {rulebook.version}",
            f"Rule SHA-256: {rulebook.sha256}",
            "Source extracts modified: no",
        ]) + "\n",
        encoding="utf-8",
    )
    return ExportResult(dataset.key, output, manifest, rows)
