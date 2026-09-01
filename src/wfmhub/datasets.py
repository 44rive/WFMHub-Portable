"""Reusable report dataset contracts built from facts and semantic metrics."""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from typing import Any

from .database import DatabaseConnection
from .metrics import MetricCatalog
from .semantic import MetricAggregate, aggregate_metric_values


@dataclass(frozen=True)
class Dataset:
    key: str
    grain: str
    purpose: str
    headers: list[str]
    rows: list[tuple[Any, ...]]


@dataclass(frozen=True)
class DatasetContract:
    key: str
    grain: str
    purpose: str


DATASET_CONTRACTS = {
    "service_scope_interval": DatasetContract(
        "service_scope_interval", "date/source/LOB/language/interval",
        "Additive service counters with configured ratio-of-sums KPI values.",
    ),
    "pcs_team_day": DatasetContract(
        "pcs_team_day", "date/team leader/LOB/language",
        "Exact PCS counters and configured response-weighted metrics.",
    ),
    "pcs_agent_month": DatasetContract(
        "pcs_agent_month", "month/agent",
        "Monthly agent PCS counters and configured response-weighted metrics.",
    ),
    "final_absence_lob_month": DatasetContract(
        "final_absence_lob_month", "month/LOB/language",
        "Overlap-safe final Verint counters and configured ratios.",
    ),
}


def _metric_map(
    conn: DatabaseConnection,
    catalog: MetricCatalog,
    start: date,
    end: date,
    metric_ids: list[str],
    dimensions: tuple[str, ...],
) -> dict[tuple[Any, ...], dict[str, MetricAggregate]]:
    output: dict[tuple[Any, ...], dict[str, MetricAggregate]] = {}
    for item in aggregate_metric_values(conn, catalog, start, end, metric_ids, dimensions):
        key = tuple(
            "(blank)" if item.dimensions[dimension] in {None, ""} else item.dimensions[dimension]
            for dimension in dimensions
        )
        output.setdefault(key, {})[item.metric_id] = item
    return output


def service_scope_interval(
    conn: DatabaseConnection,
    catalog: MetricCatalog,
    start: date,
    end: date | None = None,
    source_system: str = "APDE",
    limit: int = 100_000,
) -> Dataset:
    end = end or start
    dimensions = ("business_date", "interval_start", "source_system", "lob", "language")
    metrics = _metric_map(
        conn, catalog, start, end,
        ["service_level", "service_availability", "abandon_rate", "aht_seconds"],
        dimensions,
    )
    base = conn.execute(
        """SELECT business_date, interval_start, source_system,
                  coalesce(lob,'(blank)') AS lob,
                  coalesce(language,'(blank)') AS language,
                  sum(offered) AS offered, sum(answered) AS answered,
                  sum(abandoned) AS abandoned,
                  sum(short_abandoned) AS short_abandoned,
                  sum(answered_within_target) AS answered_within_target,
                  max(mapping_status) AS mapping_status
           FROM mart.service_interval
           WHERE source_system=? AND business_date BETWEEN ? AND ?
           GROUP BY business_date, interval_start, source_system,
                    coalesce(lob,'(blank)'), coalesce(language,'(blank)')
           ORDER BY interval_start, lob, language LIMIT ?""",
        [source_system, start, end, limit],
    ).fetchall()
    rows: list[tuple[Any, ...]] = []
    for base_row in base:
        key = tuple(base_row[index] for index in range(5))
        scoped = metrics.get(key, {})
        service = scoped.get("service_level")
        availability = scoped.get("service_availability")
        abandon = scoped.get("abandon_rate")
        aht = scoped.get("aht_seconds")
        rows.append((*base_row[:5], *base_row[5:10],
                     service.value if service else None,
                     availability.value if availability else None,
                     abandon.value if abandon else None,
                     aht.value if aht else None,
                     service.target if service else None,
                     service.state if service else "NO_DATA",
                     base_row[10]))
    return Dataset(
        "service_scope_interval", DATASET_CONTRACTS["service_scope_interval"].grain,
        DATASET_CONTRACTS["service_scope_interval"].purpose,
        ["business_date", "interval_start", "source_system", "lob", "language",
         "offered", "answered", "abandoned", "short_abandoned", "answered_within_target",
         "service_level", "service_availability", "abandon_rate", "aht_seconds",
         "sl_target", "sl_state", "mapping_status"], rows,
    )


def pcs_team_day(
    conn: DatabaseConnection,
    catalog: MetricCatalog,
    start: date,
    end: date,
) -> Dataset:
    dimensions = ("business_date", "team_leader", "lob", "language")
    metrics = _metric_map(
        conn, catalog, start, end, ["pcs_average", "pcs_participation"], dimensions,
    )
    base = conn.execute(
        """SELECT business_date, coalesce(team_leader,'(blank)') AS team_leader,
                  coalesce(lob,'(blank)') AS lob, coalesce(language,'(blank)') AS language,
                  sum(inbound_calls), sum(pcs_enabled_calls), sum(pcs_status_calls),
                  sum(pcs_participation_responses), sum(survey_responses),
                  sum(pcs_score_sum), sum(low_score_responses), sum(top_box_responses),
                  sum(pcs_invalid_responses)
           FROM mart.agent_pcs_day WHERE business_date BETWEEN ? AND ?
           GROUP BY business_date, coalesce(team_leader,'(blank)'),
                    coalesce(lob,'(blank)'), coalesce(language,'(blank)')
           ORDER BY business_date, team_leader, lob, language""",
        [start, end],
    ).fetchall()
    rows = []
    for item in base:
        scoped = metrics.get(tuple(item[:4]), {})
        rows.append((*item,
                     scoped.get("pcs_average").value if scoped.get("pcs_average") else None,
                     scoped.get("pcs_participation").value if scoped.get("pcs_participation") else None))
    return Dataset(
        "pcs_team_day", DATASET_CONTRACTS["pcs_team_day"].grain,
        DATASET_CONTRACTS["pcs_team_day"].purpose,
        ["business_date", "team_leader", "lob", "language", "inbound_call_legs",
         "pcs_mode_2_inbound_legs", "pcs_status_1_inbound_legs",
         "pcs_q1_nonblank_inbound_legs", "pcs_q1_valid_score_count",
         "pcs_q1_score_sum", "pcs_score_le_boundary_count",
         "pcs_score_gt_boundary_count", "pcs_q1_invalid_nonblank_count",
         "pcs_average", "pcs_participation_rate"], rows,
    )


def pcs_agent_month(
    conn: DatabaseConnection,
    catalog: MetricCatalog,
    start: date,
    end: date,
) -> Dataset:
    dimensions = ("month_key", "agent_id")
    metrics = _metric_map(
        conn, catalog, start, end, ["pcs_average", "pcs_participation"], dimensions,
    )
    base = conn.execute(
        """SELECT substr(business_date,1,7) AS month_key, agent_id,
                  max(agent_name), max(team_leader), max(ops_manager), max(lob), max(language),
                  sum(inbound_calls), sum(pcs_status_calls), sum(pcs_participation_responses),
                  sum(survey_responses), sum(pcs_score_sum), sum(low_score_responses),
                  sum(top_box_responses), sum(pcs_invalid_responses)
           FROM mart.agent_pcs_day WHERE business_date BETWEEN ? AND ?
           GROUP BY substr(business_date,1,7), agent_id ORDER BY month_key, agent_id""",
        [start, end],
    ).fetchall()
    rows = []
    for item in base:
        scoped = metrics.get(tuple(item[:2]), {})
        average = scoped.get("pcs_average")
        participation = scoped.get("pcs_participation")
        rows.append((*item, average.value if average else None,
                     participation.value if participation else None,
                     "LOW_SAMPLE" if item[10] < 20 else "OK"))
    return Dataset(
        "pcs_agent_month", DATASET_CONTRACTS["pcs_agent_month"].grain,
        DATASET_CONTRACTS["pcs_agent_month"].purpose,
        ["month_key", "agent_id", "agent_name", "team_leader", "ops_manager", "lob",
         "language", "inbound_call_legs", "pcs_status_1_inbound_legs",
         "pcs_q1_nonblank_inbound_legs", "pcs_q1_valid_score_count", "pcs_q1_score_sum",
         "pcs_score_le_boundary_count", "pcs_score_gt_boundary_count",
         "pcs_q1_invalid_nonblank_count", "pcs_average", "pcs_participation_rate",
         "sample_flag"], rows,
    )


def final_absence_lob_month(
    conn: DatabaseConnection,
    catalog: MetricCatalog,
    start: date,
    end: date,
) -> Dataset:
    dimensions = ("month_key", "lob", "language")
    metrics = _metric_map(
        conn, catalog, start, end,
        ["final_absence_rate", "final_vacation_rate", "final_shrinkage_rate"], dimensions,
    )
    base = conn.execute(
        """SELECT substr(business_date,1,7) AS month_key,
                  coalesce(lob,'(blank)') AS lob, coalesce(language,'(blank)') AS language,
                  count(*), sum(planned_net_minutes)/60.0,
                  sum(final_absence_minutes)/60.0, sum(final_vacation_minutes)/60.0,
                  sum(final_unpaid_minutes)/60.0, sum(final_shrinkage_minutes)/60.0,
                  sum(final_unmapped_minutes)/60.0,
                  sum(CASE WHEN final_absence_day THEN 1 ELSE 0 END)
           FROM mart.verint_final_absence_agent_day
           WHERE business_date BETWEEN ? AND ?
           GROUP BY substr(business_date,1,7), coalesce(lob,'(blank)'),
                    coalesce(language,'(blank)') ORDER BY month_key, lob, language""",
        [start, end],
    ).fetchall()
    rows = []
    for item in base:
        scoped = metrics.get(tuple(item[:3]), {})
        rows.append((*item[:10],
                     scoped.get("final_absence_rate").value if scoped.get("final_absence_rate") else None,
                     scoped.get("final_vacation_rate").value if scoped.get("final_vacation_rate") else None,
                     scoped.get("final_shrinkage_rate").value if scoped.get("final_shrinkage_rate") else None,
                     item[10]))
    return Dataset(
        "final_absence_lob_month", DATASET_CONTRACTS["final_absence_lob_month"].grain,
        DATASET_CONTRACTS["final_absence_lob_month"].purpose,
        ["month_key", "lob", "language", "agent_days", "planned_net_hours",
         "final_absence_hours", "final_vacation_hours", "final_unpaid_hours",
         "final_shrinkage_hours", "final_unmapped_hours", "final_absence_rate",
         "final_vacation_rate", "final_shrinkage_rate", "absence_agent_days"], rows,
    )
