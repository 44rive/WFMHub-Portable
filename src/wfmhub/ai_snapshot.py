"""Build a bounded, read-only SQLite bundle for external analysis tools."""

from __future__ import annotations

import hashlib
import json
import shutil
import sqlite3
import stat
import uuid
from dataclasses import dataclass
from datetime import date, datetime, timezone
from pathlib import Path
from typing import Any

from .config import Config
from .database import DatabaseConnection
from .mapping import load_queue_mapping
from .progress import ProgressCallback
from .rules import load_rulebook


_SNAPSHOT_APPLICATION_ID = 0x57464149  # ASCII "WFAI"
_BUNDLE_FORMAT_VERSION = 1


@dataclass(frozen=True)
class SnapshotColumn:
    name: str
    sqlite_type: str


@dataclass(frozen=True)
class SnapshotDataset:
    key: str
    grain: str
    columns: tuple[SnapshotColumn, ...]
    sql: str
    period_filtered: bool = True
    date_column: str | None = "business_date"


@dataclass(frozen=True)
class AnalysisSnapshotResult:
    bundle_dir: Path
    database: Path
    manifest: Path
    row_counts: dict[str, int]


def _columns(specification: str) -> tuple[SnapshotColumn, ...]:
    """Turn a compact ``name TYPE`` declaration into immutable columns."""
    tokens = specification.split()
    if len(tokens) % 2:
        raise ValueError("Snapshot column specification must contain name/type pairs")
    return tuple(
        SnapshotColumn(tokens[index], tokens[index + 1])
        for index in range(0, len(tokens), 2)
    )


SNAPSHOT_DATASETS: tuple[SnapshotDataset, ...] = (
    SnapshotDataset(
        key="source_health",
        grain="one row per configured source family",
        period_filtered=False,
        date_column="newest_business_date",
        columns=_columns(
            "source_family TEXT newest_file TEXT newest_business_date DATE modified_at TIMESTAMP "
            "loaded_at TIMESTAMP row_count INTEGER rejected_count INTEGER scoped_out_count INTEGER "
            "status TEXT details TEXT"
        ),
        sql="""SELECT source_family, newest_file, newest_business_date, modified_at,
                      loaded_at, row_count, rejected_count, scoped_out_count, status, details
               FROM mart.source_health ORDER BY source_family""",
    ),
    SnapshotDataset(
        key="daily_service_lob",
        grain="business date, interval start, LOB and language",
        columns=_columns(
            "business_date DATE interval_start TIMESTAMP lob TEXT language TEXT offered REAL answered REAL "
            "abandoned REAL short_abandoned REAL answered_within_target REAL handled_seconds REAL "
            "service_level REAL service_availability REAL sl_target REAL sl_state TEXT mapping_status TEXT"
        ),
        sql="""SELECT business_date, interval_start, coalesce(lob,'(blank)') AS lob,
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
    SnapshotDataset(
        key="daily_staffing_gaps",
        grain="business date, interval start, LOB and language",
        columns=_columns(
            "business_date DATE interval_start TIMESTAMP interval_end TIMESTAMP lob TEXT language TEXT "
            "scheduled_agents INTEGER observed_agents INTEGER productive_agents INTEGER auxiliary_agents INTEGER "
            "scheduled_fte REAL elapsed_scheduled_fte REAL observed_fte REAL productive_fte REAL "
            "staffing_variance_fte REAL staffing_gap_fte REAL staffing_state TEXT evidence_basis TEXT "
            "evaluation_as_of TIMESTAMP"
        ),
        sql="""SELECT business_date, interval_start, interval_end, lob, language,
                      scheduled_agents, observed_agents, productive_agents, auxiliary_agents,
                      scheduled_fte, elapsed_scheduled_fte, observed_fte, productive_fte,
                      staffing_variance_fte, staffing_gap_fte, staffing_state, evidence_basis,
                      evaluation_as_of
               FROM mart.staffing_interval
               WHERE business_date BETWEEN ? AND ?
               ORDER BY business_date, interval_start, lob, language""",
    ),
    SnapshotDataset(
        key="pcs_team_day",
        grain="business date, team leader, LOB and language",
        columns=_columns(
            "business_date DATE team_leader TEXT lob TEXT language TEXT inbound_call_legs INTEGER "
            "pcs_mode_2_inbound_legs INTEGER pcs_status_1_inbound_legs INTEGER "
            "pcs_q1_nonblank_inbound_legs INTEGER pcs_q1_valid_score_count INTEGER pcs_q1_score_sum REAL "
            "pcs_score_le_3_count INTEGER pcs_score_gt_3_count INTEGER pcs_q1_invalid_nonblank_count INTEGER "
            "pcs_average REAL pcs_participation_rate REAL"
        ),
        sql="""SELECT business_date, coalesce(team_leader,'(blank)') AS team_leader,
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
)


def _quote_identifier(value: str) -> str:
    return '"' + value.replace('"', '""') + '"'


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def _iso(value: Any) -> str | None:
    if value is None:
        return None
    return value.isoformat() if isinstance(value, date | datetime) else str(value)


def _provenance(conn: DatabaseConnection, config: Config) -> dict[str, Any]:
    applied_rule = conn.execute(
        """SELECT run_id, rule_version, rule_sha256, effective_from, applied_at
           FROM meta.rule_application ORDER BY applied_at DESC LIMIT 1"""
    ).fetchone()
    applied_mapping = conn.execute(
        """SELECT run_id, mapping_sha256, applied_at
           FROM meta.mapping_application ORDER BY applied_at DESC LIMIT 1"""
    ).fetchone()
    refresh = conn.execute(
        """SELECT run_id, started_at, finished_at, requested_start, requested_end
           FROM meta.refresh_run WHERE status='SUCCESS'
           ORDER BY coalesce(finished_at, started_at) DESC LIMIT 1"""
    ).fetchone()
    current_rule = load_rulebook(config.home, config.business_rules)
    current_mapping = load_queue_mapping(config.queue_mapping)
    applied_rule_sha = str(applied_rule[2]) if applied_rule else None
    applied_mapping_sha = str(applied_mapping[1]) if applied_mapping else None
    return {
        "model_run_id": str(applied_rule[0]) if applied_rule else None,
        "latest_successful_refresh": {
            "run_id": str(refresh[0]),
            "started_at": _iso(refresh[1]),
            "finished_at": _iso(refresh[2]),
            "requested_start": _iso(refresh[3]),
            "requested_end": _iso(refresh[4]),
        } if refresh else None,
        "applied_rules": {
            "version": str(applied_rule[1]),
            "sha256": applied_rule_sha,
            "effective_from": _iso(applied_rule[3]),
            "applied_at": _iso(applied_rule[4]),
        } if applied_rule else None,
        "applied_queue_mapping": {
            "run_id": str(applied_mapping[0]),
            "sha256": applied_mapping_sha,
            "applied_at": _iso(applied_mapping[2]),
        } if applied_mapping else None,
        "current_configuration": {
            "rules_version": current_rule.version,
            "rules_sha256": current_rule.sha256,
            "queue_mapping_sha256": current_mapping.sha256,
        },
        "configuration_matches_model": (
            applied_rule_sha == current_rule.sha256
            and applied_mapping_sha == current_mapping.sha256
        ) if applied_rule and applied_mapping else None,
    }


def _copy_dataset(
    source: DatabaseConnection,
    destination: sqlite3.Connection,
    dataset: SnapshotDataset,
    start: date,
    end: date,
    progress: ProgressCallback | None,
) -> int:
    table = _quote_identifier(dataset.key)
    declarations = ", ".join(
        f"{_quote_identifier(column.name)} {column.sqlite_type}"
        for column in dataset.columns
    )
    destination.execute(f"CREATE TABLE {table} ({declarations})")
    parameters = [start, end] if dataset.period_filtered else []
    cursor = source.execute(dataset.sql, parameters)
    returned = tuple(item[0] for item in cursor.description)
    expected = tuple(column.name for column in dataset.columns)
    if returned != expected:
        raise RuntimeError(
            f"Snapshot dataset {dataset.key} returned an unexpected schema: {returned!r}"
        )
    placeholders = ", ".join("?" for _ in expected)
    insert_sql = f"INSERT INTO {table} VALUES ({placeholders})"
    count = 0
    while True:
        rows = cursor.fetchmany(5000)
        if not rows:
            break
        destination.executemany(insert_sql, rows)
        count += len(rows)
        if progress is not None:
            progress(count, 0, f"Snapshotting {dataset.key}: {count:,} rows")
    if dataset.period_filtered and dataset.date_column:
        destination.execute(
            f"CREATE INDEX {_quote_identifier('idx_' + dataset.key + '_date')} "
            f"ON {table} ({_quote_identifier(dataset.date_column)})"
        )
    return count


def _make_read_only(path: Path) -> None:
    path.chmod(path.stat().st_mode & ~(stat.S_IWUSR | stat.S_IWGRP | stat.S_IWOTH))


def create_analysis_snapshot(
    source: DatabaseConnection,
    config: Config,
    start: date,
    end: date,
    output: Path | None = None,
    progress: ProgressCallback | None = None,
) -> AnalysisSnapshotResult:
    """Copy fixed, governed aggregates into an immutable analysis bundle.

    The source connection is queried only with built-in SELECT statements. The
    operational database and original extract files are never attached to or
    exposed by the resulting bundle.
    """
    if start > end:
        raise ValueError("Snapshot start date cannot be after end date")
    if source.execute("PRAGMA query_only").fetchone()[0] != 1:
        raise ValueError("Analysis snapshot source connection must be read-only")
    snapshot_id = uuid.uuid4().hex
    generated_at = datetime.now(timezone.utc).isoformat(timespec="seconds").replace("+00:00", "Z")
    default_name = f"wfmhub_analysis_{start:%Y%m%d}_{end:%Y%m%d}_{datetime.now():%Y%m%d_%H%M%S_%f}"
    target = (output or config.output / "ai_analysis" / default_name).resolve()
    if target.exists():
        raise FileExistsError(f"Analysis snapshot output already exists: {target}")
    target.parent.mkdir(parents=True, exist_ok=True)
    temporary = target.with_name(f".{target.name}.partial-{uuid.uuid4().hex}")
    temporary.mkdir()
    database_path = temporary / "wfmhub_analysis.sqlite3"
    manifest_path = temporary / "manifest.json"
    destination: sqlite3.Connection | None = None
    started_transaction = not source.in_transaction
    try:
        if started_transaction:
            source.execute("BEGIN")
        provenance = _provenance(source, config)
        destination = sqlite3.connect(database_path)
        destination.execute(f"PRAGMA application_id={_SNAPSHOT_APPLICATION_ID}")
        destination.execute(f"PRAGMA user_version={_BUNDLE_FORMAT_VERSION}")
        destination.execute("PRAGMA journal_mode=DELETE")
        destination.execute("PRAGMA synchronous=FULL")
        destination.execute("BEGIN IMMEDIATE")
        row_counts: dict[str, int] = {}
        datasets_manifest: list[dict[str, Any]] = []
        for index, dataset in enumerate(SNAPSHOT_DATASETS):
            if progress is not None:
                progress(index, len(SNAPSHOT_DATASETS), f"Preparing {dataset.key}")
            count = _copy_dataset(
                source, destination, dataset, start, end, progress
            )
            row_counts[dataset.key] = count
            datasets_manifest.append({
                "key": dataset.key,
                "table": dataset.key,
                "grain": dataset.grain,
                "rows": count,
            })
        destination.commit()
        check = destination.execute("PRAGMA quick_check").fetchone()[0]
        if str(check).lower() != "ok":
            raise RuntimeError(f"Analysis snapshot SQLite check failed: {check}")
        destination.close()
        destination = None
        if started_transaction:
            source.execute("ROLLBACK")
            started_transaction = False
        database_sha256 = _sha256(database_path)
        manifest = {
            "snapshot_id": snapshot_id,
            "bundle_format_version": _BUNDLE_FORMAT_VERSION,
            "generated_at": generated_at,
            "purpose": "Read-only governed WFM analysis input",
            "period": {"start": start.isoformat(), "end": end.isoformat()},
            "datasets": datasets_manifest,
            "files": [{
                "name": database_path.name,
                "sha256": database_sha256,
                "bytes": database_path.stat().st_size,
            }],
            "provenance": provenance,
            "safety": {
                "read_only": True,
                "contains_raw_extracts": False,
                "contains_arbitrary_sql": False,
            },
            "source_database_modified": False,
            "source_extracts_modified": False,
        }
        manifest_path.write_text(
            json.dumps(manifest, indent=2, ensure_ascii=False, sort_keys=True) + "\n",
            encoding="utf-8",
        )
        _make_read_only(database_path)
        _make_read_only(manifest_path)
        temporary.rename(target)
        if progress is not None:
            progress(len(SNAPSHOT_DATASETS), len(SNAPSHOT_DATASETS), "Analysis snapshot ready")
        return AnalysisSnapshotResult(
            bundle_dir=target,
            database=target / database_path.name,
            manifest=target / manifest_path.name,
            row_counts=row_counts,
        )
    except Exception:
        if destination is not None:
            destination.close()
        if started_transaction and source.in_transaction:
            source.execute("ROLLBACK")
        shutil.rmtree(temporary, ignore_errors=True)
        raise
