"""Governed component providers and materialized semantic metric values."""

from __future__ import annotations

import hashlib
from dataclasses import dataclass
from datetime import date, datetime
from typing import Any, Iterable, Iterator, Mapping, Sequence

from .database import DatabaseConnection
from .metrics import MetricCatalog, MetricCatalogError, MetricMethod, evaluate_metric
from .rules import Rulebook


@dataclass(frozen=True)
class ComponentRecord:
    source_model: str
    grain: str
    entity_key: str
    business_date: date
    interval_start: datetime | None
    dimensions: Mapping[str, Any]
    components: Mapping[str, Any]


@dataclass(frozen=True)
class MetricAggregate:
    metric_id: str
    method_id: str
    domain: str
    unit: str
    aggregation: str
    dimensions: Mapping[str, Any]
    numerator: float | None
    denominator: float | None
    sample_size: float | None
    value: float | None
    target: float | None
    state: str


SOURCE_COMPONENTS: dict[str, frozenset[str]] = {
    "service_interval": frozenset({
        "offered", "answered", "abandoned", "short_abandoned",
        "answered_within_target", "handled_seconds",
    }),
    "forecast_comparison_hour": frozenset({"forecast_volume", "actual_volume"}),
    "pcs_agent_day": frozenset({
        "handled_calls", "talk_seconds", "hold_seconds", "wrap_seconds",
        "handle_seconds", "pcs_status_calls",
        "pcs_participation_responses", "survey_responses", "pcs_score_count",
        "pcs_score_sum", "q1_response_count", "q1_score_sum",
        "q2_response_count", "q2_score_sum", "top_box_responses",
        "low_score_responses",
    }),
    "observed_absence_agent_day": frozenset({
        "planned_net_minutes", "absence_minutes", "vacation_minutes",
        "unpaid_minutes", "shrinkage_minutes",
    }),
    "final_absence_agent_day": frozenset({
        "planned_net_minutes", "final_absence_minutes", "final_vacation_minutes",
        "final_unpaid_minutes", "final_shrinkage_minutes",
    }),
    "staffing_interval": frozenset({
        "staffing_gap_fte", "staffing_variance_fte", "scheduled_fte",
        "observed_fte", "productive_fte", "evidence_intervals",
    }),
    "attendance_agent_day": frozenset({
        "scheduled_working_count", "no_show_count", "late_count",
        "requires_call_count", "uncoded_late_minutes", "no_show_minutes",
    }),
}


def _as_date(value: Any) -> date:
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    return date.fromisoformat(str(value))


def _as_datetime(value: Any) -> datetime | None:
    if value is None or value == "":
        return None
    if isinstance(value, datetime):
        return value
    return datetime.fromisoformat(str(value))


def _records(conn: DatabaseConnection, source_model: str, start: date, end: date) -> Iterator[ComponentRecord]:
    if source_model == "service_interval":
        cursor = conn.execute(
            """SELECT business_date, interval_start, source_system, queue,
                      business_partner, lob, language, offered, answered,
                      abandoned, short_abandoned, answered_within_target,
                      handled_seconds
               FROM mart.service_interval
               WHERE business_date BETWEEN ? AND ?
               ORDER BY business_date, interval_start, source_system, queue""",
            [start, end],
        )
        for row in cursor:
            day, interval, source, queue, partner, lob, language, *values = row
            yield ComponentRecord(
                source_model, "queue interval",
                f"{day}|{interval}|{source}|{queue or ''}|{partner or ''}|{lob or ''}|{language or ''}",
                _as_date(day), _as_datetime(interval),
                {"source_system": source, "queue": queue, "business_partner": partner,
                 "lob": lob, "language": language},
                dict(zip(("offered", "answered", "abandoned", "short_abandoned",
                          "answered_within_target", "handled_seconds"), values)),
            )
        return

    if source_model == "forecast_comparison_hour":
        cursor = conn.execute(
            """WITH forecast AS (
                   SELECT business_date, hour_start,
                          coalesce(comparison_scope, service_scope, queue_name, '(unmapped)') AS scope,
                          sum(volume_forecast) AS forecast_volume
                   FROM mart.forecast_hour
                   WHERE business_date BETWEEN ? AND ? AND mapping_status='MAPPED'
                   GROUP BY business_date, hour_start,
                            coalesce(comparison_scope, service_scope, queue_name, '(unmapped)')
               ), actual AS (
                   SELECT business_date, hour_start,
                          coalesce(comparison_scope, service_scope, lob, queue, '(unmapped)') AS scope,
                          sum(offered) AS actual_volume
                   FROM mart.service_interval
                   WHERE business_date BETWEEN ? AND ? AND mapping_status='MAPPED'
                   GROUP BY business_date, hour_start,
                            coalesce(comparison_scope, service_scope, lob, queue, '(unmapped)')
               )
               SELECT f.business_date, f.hour_start, f.scope,
                      f.forecast_volume, a.actual_volume
               FROM forecast f
               LEFT JOIN actual a ON a.business_date=f.business_date
                                 AND a.hour_start=f.hour_start AND a.scope=f.scope
               ORDER BY f.business_date, f.hour_start, f.scope""",
            [start, end, start, end],
        )
        for day, hour, scope, forecast, actual in cursor:
            yield ComponentRecord(
                source_model, "comparison scope/hour", f"{day}|{hour}|{scope}",
                _as_date(day), _as_datetime(hour),
                {"source_system": "FORECAST_ACTUAL", "lob": scope, "language": None},
                {"forecast_volume": forecast, "actual_volume": actual},
            )
        return

    if source_model == "pcs_agent_day":
        cursor = conn.execute(
            """SELECT agent_day_key, business_date, agent_id, team_leader, lob, language,
                      handled_calls, talk_seconds, hold_seconds, wrap_seconds,
                      handle_seconds, pcs_status_calls,
                      pcs_participation_responses, survey_responses, pcs_score_count,
                      pcs_score_sum, q1_response_count, q1_score_sum,
                      q2_response_count, q2_score_sum, top_box_responses,
                      low_score_responses
               FROM mart.agent_pcs_day WHERE business_date BETWEEN ? AND ?
               ORDER BY business_date, agent_id""",
            [start, end],
        )
        names = (
            "handled_calls", "talk_seconds", "hold_seconds", "wrap_seconds",
            "handle_seconds", "pcs_status_calls",
            "pcs_participation_responses", "survey_responses", "pcs_score_count",
            "pcs_score_sum", "q1_response_count", "q1_score_sum",
            "q2_response_count", "q2_score_sum", "top_box_responses",
            "low_score_responses",
        )
        for key, day, agent, team, lob, language, *values in cursor:
            yield ComponentRecord(
                source_model, "agent day", str(key), _as_date(day), None,
                {"agent_id": agent, "team_leader": team, "lob": lob, "language": language},
                dict(zip(names, values)),
            )
        return

    if source_model == "observed_absence_agent_day":
        cursor = conn.execute(
            """SELECT agent_day_key, business_date, agent_id, team_leader, lob, language,
                      planned_net_minutes, absence_minutes, vacation_minutes,
                      unpaid_minutes, shrinkage_minutes
               FROM mart.absence_agent_day WHERE business_date BETWEEN ? AND ?
               ORDER BY business_date, agent_id""",
            [start, end],
        )
        names = ("planned_net_minutes", "absence_minutes", "vacation_minutes", "unpaid_minutes", "shrinkage_minutes")
        for key, day, agent, team, lob, language, *values in cursor:
            yield ComponentRecord(
                source_model, "agent day", str(key), _as_date(day), None,
                {"agent_id": agent, "team_leader": team, "lob": lob, "language": language},
                dict(zip(names, values)),
            )
        return

    if source_model == "final_absence_agent_day":
        cursor = conn.execute(
            """SELECT agent_day_key, business_date, agent_id, team_leader, lob, language,
                      planned_net_minutes, final_absence_minutes, final_vacation_minutes,
                      final_unpaid_minutes, final_shrinkage_minutes
               FROM mart.verint_final_absence_agent_day
               WHERE business_date BETWEEN ? AND ?
                 AND final_ledger_status IN ('CLEAR','ABSENCE_RECORDED')
               ORDER BY business_date, agent_id""",
            [start, end],
        )
        names = (
            "planned_net_minutes", "final_absence_minutes", "final_vacation_minutes",
            "final_unpaid_minutes", "final_shrinkage_minutes",
        )
        for key, day, agent, team, lob, language, *values in cursor:
            yield ComponentRecord(
                source_model, "agent day", str(key), _as_date(day), None,
                {"agent_id": agent, "team_leader": team, "lob": lob, "language": language},
                dict(zip(names, values)),
            )
        return

    if source_model == "staffing_interval":
        cursor = conn.execute(
            """SELECT business_date, interval_start, lob, language, staffing_gap_fte,
                      staffing_variance_fte, scheduled_fte, observed_fte, productive_fte
               FROM mart.staffing_interval WHERE business_date BETWEEN ? AND ?
               ORDER BY business_date, interval_start, lob, language""",
            [start, end],
        )
        for day, interval, lob, language, gap, variance, scheduled, observed, productive in cursor:
            yield ComponentRecord(
                source_model, "LOB/language/15 minutes",
                f"{day}|{interval}|{lob or ''}|{language or ''}", _as_date(day), _as_datetime(interval),
                {"lob": lob, "language": language},
                {"staffing_gap_fte": gap, "staffing_variance_fte": variance,
                 "scheduled_fte": scheduled, "observed_fte": observed,
                 "productive_fte": productive, "evidence_intervals": 1 if gap is not None else 0},
            )
        return

    if source_model == "attendance_agent_day":
        cursor = conn.execute(
            """SELECT agent_day_key, business_date, agent_id, team_leader, lob, language,
                      assignment_type, call_action, requires_call,
                      uncoded_late_minutes, no_show_minutes
               FROM mart.attendance_agent_day WHERE business_date BETWEEN ? AND ?
               ORDER BY business_date, agent_id""",
            [start, end],
        )
        for key, day, agent, team, lob, language, assignment, action, requires_call, late_minutes, no_show_minutes in cursor:
            working = assignment not in {"Off", "Planned absence"}
            yield ComponentRecord(
                source_model, "agent day", str(key), _as_date(day), None,
                {"agent_id": agent, "team_leader": team, "lob": lob, "language": language},
                {"scheduled_working_count": 1 if working else 0,
                 "no_show_count": 1 if action == "CALL_NO_SHOW" else 0,
                 "late_count": 1 if action == "CALL_LATE" else 0,
                 "requires_call_count": 1 if requires_call else 0,
                 "uncoded_late_minutes": late_minutes or 0,
                 "no_show_minutes": no_show_minutes or 0},
            )
        return

    raise MetricCatalogError(f"No component provider is registered for {source_model!r}")


_METRIC_COLUMNS = (
    "metric_key", "run_id", "business_date", "interval_start", "source_model",
    "grain", "entity_key", "metric_id", "method_id", "method_effective_from",
    "domain", "unit", "aggregation", "source_system", "lob", "language",
    "team_leader", "agent_id", "numerator", "denominator", "sample_size",
    "metric_value", "target_value", "metric_state", "catalog_version",
    "catalog_sha256", "rule_version", "rule_sha256",
)


def _metric_tuple(
    run_id: str,
    record: ComponentRecord,
    method: MetricMethod,
    catalog: MetricCatalog,
    rulebook: Rulebook,
) -> tuple[Any, ...]:
    result = evaluate_metric(method, record.components)
    dimensions = record.dimensions
    key_material = (
        f"{record.source_model}|{record.entity_key}|{method.metric_id}|"
        f"{method.method_id}|{method.effective_from}"
    )
    metric_key = hashlib.sha256(key_material.encode("utf-8")).hexdigest()
    return (
        metric_key, run_id, record.business_date, record.interval_start,
        record.source_model, record.grain, record.entity_key, method.metric_id,
        method.method_id, method.effective_from, method.domain, method.unit,
        method.aggregation, dimensions.get("source_system"), dimensions.get("lob"),
        dimensions.get("language"), dimensions.get("team_leader"),
        dimensions.get("agent_id"), result.numerator, result.denominator,
        result.sample_size, result.value, method.target, result.state,
        catalog.version, catalog.sha256, rulebook.version, rulebook.sha256,
    )


def build_metric_values(
    conn: DatabaseConnection,
    catalog: MetricCatalog,
    rulebook: Rulebook,
    run_id: str,
    start: date,
    end: date,
) -> int:
    """Materialize base-grain metric observations from registered components."""
    conn.execute("DELETE FROM mart.metric_value")
    placeholders = ", ".join("?" for _ in _METRIC_COLUMNS)
    insert = f"INSERT INTO mart.metric_value ({', '.join(_METRIC_COLUMNS)}) VALUES ({placeholders})"
    count = 0
    batch: list[tuple[Any, ...]] = []
    source_models = tuple(dict.fromkeys(method.source_model for method in catalog.methods))
    for source_model in source_models:
        metric_ids = tuple(dict.fromkeys(
            method.metric_id for method in catalog.methods_for_source(source_model)
        ))
        for record in _records(conn, source_model, start, end):
            for metric_id in metric_ids:
                method = catalog.method_for(metric_id, record.business_date, record.dimensions)
                if method is None:
                    continue
                batch.append(_metric_tuple(run_id, record, method, catalog, rulebook))
                if len(batch) >= 1000:
                    conn.executemany(insert, batch)
                    count += len(batch)
                    batch.clear()
    if batch:
        conn.executemany(insert, batch)
        count += len(batch)
    conn.execute(
        """INSERT INTO meta.metric_application(
               run_id, catalog_version, catalog_sha256, catalog_file, applied_at
           ) VALUES (?, ?, ?, ?, ?)
           ON CONFLICT(run_id) DO UPDATE SET
               catalog_version=excluded.catalog_version,
               catalog_sha256=excluded.catalog_sha256,
               catalog_file=excluded.catalog_file, applied_at=excluded.applied_at""",
        [run_id, catalog.version, catalog.sha256, str(catalog.file), datetime.now()],
    )
    return count


def _state(value: float | None, target: float | None, direction: str, sample: float | None, minimum_sample: float) -> str:
    if value is None:
        return "NO_DATA"
    if minimum_sample and (sample is None or sample < minimum_sample):
        return "LOW_SAMPLE"
    if target is None or direction == "neutral":
        return "UNASSESSED"
    if direction == "higher_is_better":
        return "ON_TARGET" if value >= target else "BELOW_TARGET"
    return "ON_TARGET" if value <= target else "ABOVE_TARGET"


def aggregate_metric_values(
    conn: DatabaseConnection,
    catalog: MetricCatalog,
    start: date,
    end: date,
    metric_ids: Sequence[str] | None = None,
    dimensions: Sequence[str] = ("business_date", "source_system", "lob", "language"),
) -> list[MetricAggregate]:
    dimension_sql = {
        "business_date": "business_date",
        "month_key": "substr(business_date,1,7) AS month_key",
        "interval_start": "interval_start",
        "source_system": "source_system",
        "lob": "lob",
        "language": "language",
        "team_leader": "team_leader",
        "agent_id": "agent_id",
    }
    allowed_dimensions = set(dimension_sql)
    if any(item not in allowed_dimensions for item in dimensions):
        raise ValueError(f"Unsupported metric aggregation dimension: {dimensions}")
    where = "business_date BETWEEN ? AND ?"
    parameters: list[Any] = [start, end]
    if metric_ids:
        where += f" AND metric_id IN ({', '.join('?' for _ in metric_ids)})"
        parameters.extend(metric_ids)
    columns = [
        "metric_id", "method_id", "method_effective_from", "domain", "unit",
        "aggregation", "target_value", *(dimension_sql[item] for item in dimensions),
        "numerator", "denominator", "sample_size", "metric_value",
    ]
    rows = conn.execute(
        f"SELECT {', '.join(columns)} FROM mart.metric_value WHERE {where}", parameters
    ).fetchall()
    grouped: dict[tuple[Any, ...], list[tuple[Any, ...]]] = {}
    prefix = 7 + len(dimensions)
    for row in rows:
        key = tuple(row[:prefix])
        grouped.setdefault(key, []).append(row[prefix:])
    output: list[MetricAggregate] = []
    for key, values in grouped.items():
        metric_id, method_id, effective_from, domain, unit, aggregation, target, *scope_values = key
        method_candidates = [
            item for item in catalog.methods
            if item.metric_id == metric_id and item.method_id == method_id
            and item.effective_from == _as_date(effective_from)
        ]
        if len(method_candidates) != 1:
            raise MetricCatalogError(f"Cannot resolve stored metric method {metric_id}.{method_id}")
        method = method_candidates[0]
        numerators = [float(item[0]) for item in values if item[0] is not None]
        denominators = [float(item[1]) for item in values if item[1] is not None]
        samples = [float(item[2]) for item in values if item[2] is not None]
        raw_values = [float(item[3]) for item in values if item[3] is not None]
        numerator = sum(numerators) if numerators else None
        denominator = sum(denominators) if denominators else None
        sample = sum(samples) if samples else None
        if aggregation in {"ratio_of_sums", "mean"}:
            value = None if numerator is None or denominator in {None, 0} else numerator / denominator
        elif aggregation == "sum":
            value = numerator
        elif aggregation == "maximum":
            value = max(raw_values) if raw_values else None
        else:
            value = min(raw_values) if raw_values else None
        output.append(MetricAggregate(
            metric_id=metric_id, method_id=method_id, domain=domain, unit=unit,
            aggregation=aggregation,
            dimensions=dict(zip(dimensions, scope_values)),
            numerator=numerator, denominator=denominator, sample_size=sample,
            value=value, target=float(target) if target is not None else None,
            state=_state(value, float(target) if target is not None else None,
                         method.direction, sample, method.minimum_sample),
        ))
    return output
