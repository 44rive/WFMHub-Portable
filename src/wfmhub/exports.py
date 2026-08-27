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
        "pcs_agent_day", "One clean PCS/call-performance row per agent/day.",
        "SELECT * FROM mart.agent_pcs_day WHERE business_date BETWEEN ? AND ? ORDER BY business_date, agent_id",
    ),
    "attendance": ExportDataset(
        "attendance", "Attendance agent/day results without adherence metrics.",
        """SELECT * FROM mart.attendance_agent_day
           WHERE business_date BETWEEN ? AND ? ORDER BY business_date, agent_id""",
    ),
    "absence_agent_day": ExportDataset(
        "absence_agent_day", "Rule-versioned payroll absence, vacation and shrinkage per agent/day.",
        "SELECT * FROM mart.absence_agent_day WHERE business_date BETWEEN ? AND ? ORDER BY business_date, agent_id",
    ),
    "absence_events": ExportDataset(
        "absence_events", "Classified, schedule-clipped Verint and LILO absence evidence.",
        "SELECT * FROM mart.absence_event WHERE business_date BETWEEN ? AND ? ORDER BY business_date, agent_id, event_start",
    ),
    "gaps": ExportDataset(
        "gaps", "Correction candidates and saved decisions.",
        "SELECT * FROM mart.correction_candidate WHERE business_date BETWEEN ? AND ? ORDER BY business_date, agent_id, gap_start",
    ),
    "schedules": ExportDataset(
        "schedules", "Parsed, FTE-scoped schedule shifts.",
        """SELECT source_file_id, source_row, schedule_date, agent_id_raw,
                  agent_id, agent_name, scheduling_period, shift_assignment,
                  assignment, assignment_type, scheduled_start, scheduled_end,
                  shift_events, parse_ok, source_file
           FROM (
               SELECT r.*, f.file_name AS source_file,
                      row_number() OVER (
                          PARTITION BY r.schedule_date, r.agent_id,
                                       r.scheduled_start, r.scheduled_end,
                                       coalesce(r.assignment,'')
                          ORDER BY f.modified_at DESC, f.loaded_at DESC, r.source_row DESC
                      ) row_rank
               FROM raw.schedule_shift r JOIN meta.source_file f ON f.file_id=r.source_file_id
               WHERE f.active=true AND f.status='SUCCESS' AND r.schedule_date BETWEEN ? AND ?
           ) x WHERE row_rank=1
           ORDER BY schedule_date, agent_id, scheduled_start""",
    ),
    "events": ExportDataset(
        "events", "Parsed Verint schedule activity intervals.",
        """SELECT source_file_id, source_row, event_index, schedule_date,
                  agent_id, agent_name, activity, activity_type, event_start,
                  event_end, parse_ok, source_file
           FROM (
               SELECT r.*, f.file_name AS source_file,
                      row_number() OVER (
                          PARTITION BY r.schedule_date, r.agent_id,
                                       upper(coalesce(r.activity,'')), r.event_start, r.event_end
                          ORDER BY f.modified_at DESC, f.loaded_at DESC, r.source_row DESC
                      ) row_rank
               FROM raw.schedule_event r JOIN meta.source_file f ON f.file_id=r.source_file_id
               WHERE f.active=true AND f.status='SUCCESS' AND r.schedule_date BETWEEN ? AND ?
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
