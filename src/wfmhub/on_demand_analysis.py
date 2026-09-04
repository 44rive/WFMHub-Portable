"""First-class deterministic analysis for a selected period and domain."""

from __future__ import annotations

from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any

from .config import Config
from .database import DatabaseConnection
from .decision_products import _audit_rows, _finish, _previous_month, _ratio
from .reports import _query
from .template_reports import DecisionWorkbook, KpiCard


ANALYSIS_DOMAINS = ("pcs", "service", "forecast", "staffing", "attendance", "absence", "bonus")
COMPARISON_MODES = ("previous_equal", "previous_month", "target", "none")


def _reference_period(start: date, end: date, mode: str) -> tuple[date | None, date | None, str]:
    if mode == "previous_equal":
        reference_end = start - timedelta(days=1)
        reference_start = reference_end - timedelta(days=(end - start).days)
        return reference_start, reference_end, "Previous equal period"
    if mode == "previous_month":
        month_start, month_end = _previous_month(end)
        comparable_end = min(month_end, month_start + timedelta(days=end.day - 1))
        return month_start, comparable_end, "Previous-month same days"
    if mode == "target":
        return None, None, "Configured target"
    return None, None, "No comparison"


def _metric_summary(
    conn: DatabaseConnection,
    domain: str,
    start: date,
    end: date,
    reference_start: date | None,
    reference_end: date | None,
    mode: str,
) -> list[tuple[Any, ...]]:
    if domain == "bonus":
        def bonus_values(period_start: date, period_end: date) -> tuple[Any, ...]:
            return conn.execute(
                """SELECT count(*), coalesce(sum(scenario_payout),0),
                          coalesce(sum(released_payout),0),
                          coalesce(sum(CASE WHEN release_status='READY' THEN 1 ELSE 0 END),0),
                          avg(final_achievement), avg(proration)
                   FROM mart.bonus_agent_month WHERE period BETWEEN ? AND ?""",
                [period_start.strftime("%Y-%m"), period_end.strftime("%Y-%m")],
            ).fetchone()

        current = bonus_values(start, end)
        if not current[0]:
            return []
        reference = bonus_values(reference_start, reference_end) if reference_start and reference_end else (0, None, None, 0, None, None)
        definitions = (
            ("scenario_payout_total", "money", current[1], reference[1], current[1], None),
            ("released_payout_total", "money", current[2], reference[2], current[2], None),
            ("release_ready_rate", "percent", _ratio(current[3], current[0]), _ratio(reference[3], reference[0]), current[3], current[0]),
            ("average_achievement", "percent", current[4], reference[4], current[4], current[0]),
            ("average_proration", "percent", current[5], reference[5], current[5], current[0]),
        )
        return [
            (
                metric_id, "bonus_matrix", unit, value, comparison_value,
                value - comparison_value if value is not None and comparison_value is not None else None,
                numerator, denominator, None, "bonus_import_v1",
            )
            for metric_id, unit, value, comparison_value, numerator, denominator in definitions
        ]

    current = conn.execute(
        """SELECT metric_id, method_id, max(unit), sum(numerator), sum(denominator),
                  max(target_value), max(catalog_version)
           FROM mart.metric_value
           WHERE domain=? AND business_date BETWEEN ? AND ?
           GROUP BY metric_id, method_id ORDER BY metric_id, method_id""",
        [domain, start, end],
    ).fetchall()
    rows = []
    for metric_id, method_id, unit, numerator, denominator, target, catalog_version in current:
        current_value = _ratio(numerator, denominator)
        reference_value = None
        if reference_start and reference_end:
            reference = conn.execute(
                """SELECT sum(numerator), sum(denominator) FROM mart.metric_value
                   WHERE domain=? AND metric_id=? AND method_id=?
                     AND business_date BETWEEN ? AND ?""",
                [domain, metric_id, method_id, reference_start, reference_end],
            ).fetchone()
            reference_value = _ratio(reference[0], reference[1])
        elif mode == "target":
            reference_value = target
        rows.append((
            metric_id, method_id, unit, current_value, reference_value,
            current_value - reference_value if current_value is not None and reference_value is not None else None,
            numerator, denominator, target, catalog_version,
        ))
    return rows


def _evidence(conn: DatabaseConnection, domain: str, start: date, end: date) -> tuple[list[str], list[tuple[Any, ...]]]:
    contracts = {
        "pcs": """SELECT business_date, agent_id, agent_name, team_leader, lob, language,
                         inbound_calls, pcs_status_calls, pcs_participation_responses,
                         survey_responses, pcs_score_sum, pcs_average, pcs_participation_rate,
                         low_score_responses, top_box_responses
                  FROM mart.agent_pcs_day WHERE business_date BETWEEN ? AND ?
                  ORDER BY business_date, lob, team_leader, agent_name""",
        "service": """SELECT business_date, interval_start, source_system, queue, service_scope,
                             comparison_scope, offered, answered, short_abandoned,
                             answered_within_target, service_level, service_availability,
                             aht_seconds, mapping_status
                      FROM mart.service_interval WHERE business_date BETWEEN ? AND ?
                      ORDER BY business_date, interval_start, service_scope, queue""",
        "forecast": """WITH forecast AS (
                            SELECT business_date, hour_start,
                                   coalesce(comparison_scope, service_scope, queue_name, '(unmapped)') AS comparison_scope,
                                   sum(volume_forecast) AS forecast_volume
                            FROM mart.forecast_hour
                            WHERE business_date BETWEEN ? AND ? AND mapping_status='MAPPED'
                            GROUP BY business_date, hour_start,
                                     coalesce(comparison_scope, service_scope, queue_name, '(unmapped)')
                        ), actual AS (
                            SELECT business_date, hour_start,
                                   coalesce(comparison_scope, service_scope, lob, queue, '(unmapped)') AS comparison_scope,
                                   sum(offered) AS actual_volume
                            FROM mart.service_interval
                            WHERE business_date BETWEEN ? AND ? AND mapping_status='MAPPED'
                            GROUP BY business_date, hour_start,
                                     coalesce(comparison_scope, service_scope, lob, queue, '(unmapped)')
                        )
                        SELECT f.business_date, f.hour_start, f.comparison_scope,
                               a.actual_volume, f.forecast_volume,
                               CASE WHEN a.actual_volume IS NOT NULL THEN a.actual_volume-f.forecast_volume END AS variance_volume,
                               CASE WHEN a.actual_volume IS NOT NULL THEN abs(a.actual_volume-f.forecast_volume) END AS absolute_error_volume,
                               CASE WHEN f.forecast_volume<>0 THEN a.actual_volume*1.0/f.forecast_volume END AS forecast_attainment,
                               CASE WHEN f.forecast_volume<>0 THEN abs(a.actual_volume-f.forecast_volume)*1.0/f.forecast_volume END AS absolute_error_rate,
                               'MAPPED' AS mapping_status
                        FROM forecast f
                        LEFT JOIN actual a ON a.business_date=f.business_date
                                          AND a.hour_start=f.hour_start
                                          AND a.comparison_scope=f.comparison_scope
                        ORDER BY f.business_date, f.hour_start, f.comparison_scope""",
        "staffing": """SELECT business_date, interval_start, lob, language, scheduled_fte,
                              observed_fte, productive_fte, staffing_gap_fte, staffing_state
                       FROM mart.staffing_interval WHERE business_date BETWEEN ? AND ?
                       ORDER BY business_date, interval_start, lob, language""",
        "attendance": """SELECT business_date, agent_id, agent_name, team_leader, lob, language,
                                scheduled_start, scheduled_end, attendance_result, call_action,
                                uncoded_late_minutes, uncoded_early_leave_minutes, no_show_minutes,
                                source_loaded, is_provisional
                         FROM mart.attendance_agent_day WHERE business_date BETWEEN ? AND ?
                         ORDER BY business_date, lob, team_leader, agent_name""",
        "absence": """SELECT business_date, agent_id, agent_name, team_leader, lob, language,
                             planned_net_minutes, final_absence_minutes, final_vacation_minutes,
                             final_unpaid_minutes, final_shrinkage_minutes, final_unmapped_minutes,
                             final_absence_rate, final_ledger_status
                      FROM mart.verint_final_absence_agent_day WHERE business_date BETWEEN ? AND ?
                      ORDER BY business_date, lob, team_leader, agent_name""",
    }
    if domain == "bonus":
        cursor = conn.execute(
            """SELECT period, agent_id, agent_name, population, core_ready, eligibility,
                      gross_achievement, voc_malus, final_achievement, reference_bonus,
                      proration, scenario_payout, released_payout, release_status, data_issue
               FROM mart.bonus_agent_month
               WHERE period BETWEEN ? AND ? ORDER BY period, population, agent_name""",
            [start.strftime("%Y-%m"), end.strftime("%Y-%m")],
        )
        return [item[0] for item in cursor.description], cursor.fetchall()
    parameters = [start, end, start, end] if domain == "forecast" else [start, end]
    return _query(conn, contracts[domain], parameters)


def build_analysis_workbook(
    conn: DatabaseConnection,
    config: Config,
    domain: str,
    start: date,
    end: date,
    comparison: str = "previous_equal",
    output: Path | None = None,
) -> Path:
    if domain not in ANALYSIS_DOMAINS:
        raise ValueError(f"Unknown analysis domain {domain!r}")
    if comparison not in COMPARISON_MODES:
        raise ValueError(f"Unknown comparison mode {comparison!r}")
    generated = datetime.now()
    target = (
        output
        or config.output / "analysis"
        / f"WFMHub_Analysis_{domain.title()}_{start:%Y-%m-%d}_to_{end:%Y-%m-%d}_{generated:%H%M%S_%f}.xlsx"
    ).resolve()
    target.parent.mkdir(parents=True, exist_ok=True)
    partial = target.with_name(f"{target.stem}.partial{target.suffix}")
    book = DecisionWorkbook(partial, config, f"analysis_{domain}", f"ON-DEMAND ANALYSIS  /  {domain.upper()}", start, end, generated)
    reference_start, reference_end, reference_label = _reference_period(start, end, comparison)
    metrics = _metric_summary(conn, domain, start, end, reference_start, reference_end, comparison)
    finding_params: list[Any] = [domain, start, end]
    finding_where = "domain=? AND period_end BETWEEN ? AND ?"
    if comparison == "target":
        finding_where += " AND target_value IS NOT NULL"
    elif comparison in {"previous_equal", "previous_month"}:
        finding_where += " AND (reference_value IS NOT NULL OR finding_type='PERIOD_CHANGE')"
    findings = conn.execute(
        f"""SELECT finding_rank, severity, finding_type, metric_id, title, summary,
                   current_value, reference_value, target_value, delta_value, unit,
                   lob, language, team_leader, agent_id, evidence_dataset, evidence_filter
            FROM mart.analysis_finding WHERE {finding_where}
            ORDER BY CASE severity WHEN 'CRITICAL' THEN 1 WHEN 'WARNING' THEN 2 ELSE 3 END,
                     finding_rank""",
        finding_params,
    ).fetchall()
    critical = sum(1 for row in findings if row[1] == "CRITICAL")
    warning = sum(1 for row in findings if row[1] == "WARNING")
    information = sum(1 for row in findings if row[1] == "INFO")
    comparable = sum(1 for row in metrics if row[4] is not None)
    status = "INCOMPLETE" if not metrics else "LIVE"
    status_text = f"Compared with {reference_label}; {comparable:,} metric(s) have comparable evidence"
    metric_headers = ["Metric", "Method", "Unit", "Current Value", "Reference Value", "Delta", "Numerator", "Denominator", "Target", "Catalog Version"]
    book.dashboard(
        [
            KpiCard("Critical findings", critical, "integer"),
            KpiCard("Warnings", warning, "integer"),
            KpiCard("Information", information, "integer"),
            KpiCard("Comparable metrics", comparable, "integer", reference_label),
        ],
        status,
        status_text,
        metric_headers,
        metrics,
        [
            "Every statement points to its metric, comparison and evidence dataset.",
            "A comparison is descriptive, not proof of causality. Validate operational drivers against the EVIDENCE sheet.",
            "For deeper writing help, attach only this finished workbook to the approved Copilot account and use the hub prompt file.",
        ],
    )
    finding_headers = ["finding_rank", "severity", "finding_type", "metric_id", "title", "summary", "current_value", "reference_value", "target_value", "delta_value", "unit", "lob", "language", "team_leader", "agent_id", "evidence_dataset", "evidence_filter"]
    book.table("FINDINGS", "Evidence-backed findings", "Ranked exceptions generated from the central metric and analytics catalogs.", finding_headers, findings)
    book.table("METRICS", "Period metric comparison", f"Current selection compared with {reference_label}.", metric_headers, metrics)
    evidence_headers, evidence_rows = _evidence(conn, domain, start, end)
    book.table("EVIDENCE", "Analysis evidence", "Curated domain grain behind the findings; no untouched extract is copied here.", evidence_headers, evidence_rows)
    book.definitions([
        ("Current value", "Ratio of summed numerator and denominator", "Selected period result", "Never average percentages"),
        ("Reference value", reference_label, "Comparison context", "Not available until comparable data exists"),
        ("Delta", "Current minus reference", "Direction and magnitude", "Interpret with KPI direction"),
        ("Finding", "Configured target, sample or trend rule", "Prioritized review", "A comparison is not proof of cause"),
    ])
    book.audit(_audit_rows(conn, config, f"analysis_{domain}", start, end, [
        ("Comparison mode", comparison, reference_label),
        ("Reference period", f"{reference_start or '-'} to {reference_end or '-'}", "Target mode has no reference dates"),
    ]))
    return _finish(book, partial, target)
