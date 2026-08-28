"""Rule-aware absence, executive scorecard, and KPI catalog workbooks."""

from __future__ import annotations

from datetime import date, datetime
from pathlib import Path
from typing import Any

from .config import Config
from .database import DatabaseConnection
from .report_packs import report_pack, report_pack_folder
from .reports import ExcelReport, _query
from .rules import Rulebook, evaluate_formula, load_rulebook


def _catalog_rows(rulebook: Rulebook) -> list[tuple[Any, ...]]:
    rows: list[tuple[Any, ...]] = []
    for profile in rulebook.service_profiles.values():
        rows.append((
            f"service_level.{profile.key}", profile.label, profile.formula, "percent",
            "queue interval; ratio of sums", profile.description, rulebook.version, rulebook.sha256,
        ))
    for formula in rulebook.formulas.values():
        rows.append((
            formula.key, formula.label, formula.formula, formula.unit, formula.grain,
            formula.description, rulebook.version, rulebook.sha256,
        ))
    return rows


def _add_rule_sheets(report: ExcelReport, rulebook: Rulebook) -> None:
    report.add_table_sheet(
        "KPI_CATALOG", "Central KPI catalog",
        f"Generated from {rulebook.file}; version {rulebook.version}; SHA-256 {rulebook.sha256}.",
        ["kpi_key", "kpi_name", "formula", "unit", "grain", "description", "rule_version", "rule_sha256"],
        _catalog_rows(rulebook),
    )
    report.add_table_sheet(
        "ACTIVITY_RULES", "Verint activity rulebook",
        "First matching rule wins. Edit config/wfm_rules.toml, validate it, then refresh the selected period.",
        [
            "order", "rule_name", "category", "patterns", "match", "planned", "working",
            "counts_as_absence", "counts_as_vacation", "counts_as_unpaid", "counts_as_shrinkage",
        ],
        [
            (
                index, rule.name, rule.category, " | ".join(rule.patterns), rule.match,
                rule.planned, rule.working, rule.absence, rule.vacation, rule.unpaid, rule.shrinkage,
            )
            for index, rule in enumerate(rulebook.activity_rules, 1)
        ],
    )
    report.add_table_sheet(
        "QUEUE_SCOPES", "Service KPI scope rules",
        "Scopes select the active service-level formula. Service availability is always answered / offered.",
        ["scope", "sources", "lob_contains", "languages", "sl_profile"],
        [
            (scope.name, " | ".join(scope.sources), " | ".join(scope.lob_contains),
             " | ".join(scope.languages), scope.sl_profile)
            for scope in rulebook.queue_scopes
        ],
    )


def build_kpi_catalog(config: Config, output: Path | None = None) -> Path:
    rulebook = load_rulebook(config.home, config.business_rules)
    output = (output or config.output / "reference" / f"WFMHub_KPI_Catalog_{rulebook.version}.xlsx").resolve()
    output.parent.mkdir(parents=True, exist_ok=True)
    partial = output.with_name(f"{output.stem}.partial{output.suffix}")
    report = ExcelReport(partial)
    try:
        _add_rule_sheets(report, rulebook)
    except Exception:
        report.close()
        partial.unlink(missing_ok=True)
        raise
    else:
        report.close()
        partial.replace(output)
    return output


def _start_sheet(
    report: ExcelReport,
    title: str,
    rulebook: Rulebook,
    config: Config,
    start: date,
    end: date,
    pivot_tables: tuple[str, ...],
) -> None:
    ws = report.workbook.add_worksheet("START_HERE")
    ws.hide_gridlines(2)
    ws.merge_range("A1:H1", title, report.title)
    ws.merge_range(
        "A2:H2", f"{start:%Y-%m-%d} to {end:%Y-%m-%d} | rules {rulebook.version} | generated {datetime.now():%Y-%m-%d %H:%M}",
        report.subtitle,
    )
    lines = [
        "This workbook contains curated model facts only. Raw extracts remain untouched and stay in SQLite.",
        "Attendance and absence are observed from LILO plus Agent Status; Verint Activities only verify the final correction.",
        "No adherence KPI is included. Availability always means service availability: answered / offered.",
        "Create your PivotTable: select a cell in a PIVOT_* sheet, then Insert > PivotTable > From Table/Range.",
        f"Recommended PivotTable sources: {', '.join(pivot_tables)}.",
        "For percentage KPIs, use KPI_DAILY where ratios are already calculated from summed components; do not average percentages.",
        f"Editable rules: {rulebook.file}",
        f"Rule SHA-256: {rulebook.sha256}",
        f"Source root: {config.source_root}",
    ]
    for row, line in enumerate(lines, 4):
        ws.write(row - 1, 0, line, report.body)
    ws.set_column("A:A", 120)
    ws.set_column("B:H", 3)


def _absence_values(rulebook: Rulebook, planned_minutes: float, absence_minutes: float, vacation_minutes: float, shrinkage_minutes: float) -> dict[str, float | None]:
    values = {
        "planned_net_hours": planned_minutes / 60.0,
        "absence_hours": absence_minutes / 60.0,
        "vacation_hours": vacation_minutes / 60.0,
        "shrinkage_hours": shrinkage_minutes / 60.0,
    }
    return {
        **values,
        "absence_rate": evaluate_formula(rulebook.formulas["absence_rate"].formula, values),
        "vacation_rate": evaluate_formula(rulebook.formulas["vacation_rate"].formula, values),
        "shrinkage_rate": evaluate_formula(rulebook.formulas["shrinkage_rate"].formula, values),
    }


def build_absence_report(
    conn: DatabaseConnection,
    config: Config,
    start: date,
    end: date,
    output: Path | None = None,
) -> Path:
    rulebook = load_rulebook(config.home, config.business_rules)
    generated = datetime.now()
    pack = report_pack("absence")
    output = (
        output or report_pack_folder(config, pack.key)
        / f"{pack.filename_prefix}_{start:%Y-%m-%d}_to_{end:%Y-%m-%d}_{generated:%H%M%S_%f}.xlsx"
    ).resolve()
    output.parent.mkdir(parents=True, exist_ok=True)
    partial = output.with_name(f"{output.stem}.partial{output.suffix}")
    report = ExcelReport(partial)
    try:
        _start_sheet(report, "WFMHub Attendance & Absence", rulebook, config, start, end, ("PIVOT_ABSENCE", "ABSENCE_EVENTS", "AGENT_SPELLS"))
        planned, absence, vacation, shrinkage, corrected, unverified, absence_days = conn.execute("""
            SELECT coalesce(sum(planned_net_minutes),0), coalesce(sum(absence_minutes),0),
                   coalesce(sum(vacation_minutes),0), coalesce(sum(shrinkage_minutes),0),
                   coalesce(sum(corrected_minutes),0), coalesce(sum(unverified_minutes),0),
                   count(DISTINCT CASE WHEN absence_day THEN agent_day_key END)
            FROM mart.absence_agent_day WHERE business_date BETWEEN ? AND ?
        """, [start, end]).fetchone()
        calculated = _absence_values(rulebook, planned, absence, vacation, shrinkage)
        headers = [
            "planned_net_hours", "absence_hours", "absence_rate", "vacation_hours",
            "vacation_rate", "shrinkage_hours", "shrinkage_rate", "absence_agent_days",
            "corrected_in_verint_hours", "not_corrected_in_verint_hours",
        ]
        rows = [(
            calculated["planned_net_hours"], calculated["absence_hours"], calculated["absence_rate"],
            calculated["vacation_hours"], calculated["vacation_rate"], calculated["shrinkage_hours"],
            calculated["shrinkage_rate"], absence_days, corrected / 60.0, unverified / 60.0,
        )]
        report.add_table_sheet(
            "SUMMARY", "Absence summary",
            "Payroll absence, vacation and shrinkage are separate. All rates use summed minutes.", headers, rows,
        )
        headers, rows = _query(conn, """
            SELECT business_date, strftime('%Y', business_date) AS calendar_year,
                   strftime('%m', business_date) AS calendar_month,
                   strftime('%W', business_date) AS calendar_week,
                   agent_id, agent_name, team_leader, ops_manager, lob, market, language, location,
                   scheduled_minutes/60.0 AS scheduled_hours,
                   planned_net_minutes/60.0 AS planned_net_hours,
                   production_minutes/60.0 AS production_hours,
                   absence_minutes/60.0 AS absence_hours,
                   vacation_minutes/60.0 AS vacation_hours,
                   unpaid_minutes/60.0 AS unpaid_hours,
                   shrinkage_minutes/60.0 AS shrinkage_hours,
                   late_minutes/60.0 AS late_hours,
                   early_leave_minutes/60.0 AS early_leave_hours,
                   no_show_minutes/60.0 AS no_show_hours,
                   corrected_minutes/60.0 AS corrected_in_verint_hours,
                   unverified_minutes/60.0 AS not_corrected_in_verint_hours,
                   absence_rate, vacation_rate, shrinkage_rate, absence_day,
                   absence_spell, absence_spells, absence_days, bradford_factor,
                   rule_version, rule_sha256
            FROM mart.absence_agent_day
            WHERE business_date BETWEEN ? AND ?
            ORDER BY business_date, agent_name
            LIMIT ?
        """, [start, end, config.report_limits.get("max_absence_rows", 100000)])
        report.add_table_sheet(
            "PIVOT_ABSENCE", "Pivot-ready agent-day absence",
            "One agent/day. Build PivotTables from tblPivotAbsence; use sums for hours and rates from summed hours.", headers, rows,
        )
        headers, rows = _query(conn, """
            SELECT business_date, agent_id, agent_name, team_leader, ops_manager, lob, market,
                   language, location, activity, category, event_start, event_end,
                   hours, planned, working, counts_as_absence, counts_as_vacation,
                   counts_as_unpaid, counts_as_shrinkage, mapped, evidence_type,
                   reconciliation_status, verint_activity, verint_category,
                   verint_overlap_minutes, source_file, verint_source_file, rule_version
            FROM mart.absence_event
            WHERE business_date BETWEEN ? AND ?
            ORDER BY business_date, agent_name, event_start
            LIMIT ?
        """, [start, end, config.report_limits.get("max_absence_rows", 100000)])
        report.add_table_sheet(
            "ABSENCE_EVENTS", "Observed LILO and Agent Status gaps",
            "Verint columns are reconciliation labels only. Daily totals union observed overlaps, so minutes are not double-counted.", headers, rows,
        )
        headers, rows = _query(conn, """
            SELECT agent_id, max(agent_name) AS agent_name, max(team_leader) AS team_leader,
                   max(ops_manager) AS ops_manager, max(lob) AS lob, max(market) AS market,
                   max(language) AS language, max(location) AS location,
                   max(absence_spells) AS absence_spells, max(absence_days) AS absence_days,
                   max(bradford_factor) AS bradford_factor,
                   sum(absence_minutes)/60.0 AS absence_hours,
                   min(CASE WHEN absence_day THEN business_date END) AS first_absence_date,
                   max(CASE WHEN absence_day THEN business_date END) AS last_absence_date
            FROM mart.absence_agent_day WHERE business_date BETWEEN ? AND ?
            GROUP BY agent_id ORDER BY bradford_factor DESC, agent_name
        """, [start, end])
        report.add_table_sheet(
            "AGENT_SPELLS", "Absence frequency and Bradford",
            "Bradford = absence spells squared × standard-day equivalents absent for the selected model period.", headers, rows,
        )
        headers, rows = _query(conn, """
            SELECT business_date, agent_id, agent_name, activity, category, hours,
                   evidence_type, reconciliation_status, source_file
            FROM mart.absence_event
            WHERE reconciliation_status IN ('NOT_CORRECTED','PARTIAL') AND business_date BETWEEN ? AND ?
            ORDER BY business_date, agent_name, activity
        """, [start, end])
        report.add_table_sheet(
            "NOT_CORRECTED", "Observed gaps not fully corrected in Verint",
            "Correct or complete these in Verint, export Activities again, and refresh.", headers, rows,
        )
        headers, rows = _query(conn, """
            SELECT business_date, agent_id, agent_name, activity, category,
                   event_start, event_end, minutes, exception_type, source_file
            FROM mart.verint_final_exception
            WHERE business_date BETWEEN ? AND ?
            ORDER BY business_date, agent_name, event_start
        """, [start, end])
        report.add_table_sheet(
            "VERINT_ONLY", "Verint final activities without observed gaps",
            "Review source coverage and the injected activity. These rows do not create absence by themselves.",
            headers, rows,
        )
        _add_rule_sheets(report, rulebook)
        headers, rows = _query(conn, """
            SELECT source_family, expected_path, newest_file, newest_business_date, modified_at,
                   loaded_at, row_count, rejected_count, scoped_out_count, status, details
            FROM mart.source_health WHERE source_family IN ('fte','schedule','lilo','agent_status')
            ORDER BY source_family
        """)
        report.add_table_sheet("SOURCE_HEALTH", "Absence source health", "Check this before payroll use.", headers, rows, exception_column="Status")
    except Exception:
        report.close()
        partial.unlink(missing_ok=True)
        raise
    else:
        report.close()
        partial.replace(output)
    return output


def _kpi_row(
    business_date: date,
    domain: str,
    source: str,
    lob: str | None,
    language: str | None,
    key: str,
    name: str,
    numerator: float | None,
    denominator: float | None,
    value: float | None,
    unit: str,
    profile: str,
    rulebook: Rulebook,
) -> tuple[Any, ...]:
    return (
        business_date, business_date.isocalendar().year, business_date.isocalendar().week,
        business_date.strftime("%Y-%m"), domain, source, lob, language, key, name,
        numerator, denominator, value, unit, profile, rulebook.version, rulebook.sha256,
    )


def _scorecard_rows(conn: DatabaseConnection, rulebook: Rulebook, start: date, end: date) -> list[tuple[Any, ...]]:
    output: list[tuple[Any, ...]] = []
    service_rows = _query(conn, """
        SELECT business_date, source_system, comparison_scope, language,
               sum(offered), sum(answered), sum(abandoned), sum(short_abandoned),
               sum(answered_within_target), sum(handled_seconds)
        FROM mart.service_interval WHERE business_date BETWEEN ? AND ?
        GROUP BY business_date, source_system, comparison_scope, language
        ORDER BY business_date, source_system, comparison_scope, language
    """, [start, end])[1]
    gross = rulebook.service_profiles["gross_20"]
    adjusted = rulebook.service_profiles["adjusted_20"]
    for day, source, lob, language, offered, answered, abandoned, short_abandoned, answered_target, handled_seconds in service_rows:
        values = {
            "offered": offered, "answered": answered, "abandoned": abandoned,
            "short_abandoned": short_abandoned, "answered_within_target": answered_target,
            "handled_seconds": handled_seconds,
        }
        profile = rulebook.service_profile_for(source, lob, language)
        output.extend([
            _kpi_row(day, "SERVICE", source, lob, language, "service_level", profile.label,
                     answered_target, offered - (short_abandoned or 0) if profile.key == "adjusted_20" and offered is not None else offered,
                     evaluate_formula(profile.formula, values), "percent", profile.key, rulebook),
            _kpi_row(day, "SERVICE", source, lob, language, "sl_gross", gross.label,
                     answered_target, offered, evaluate_formula(gross.formula, values), "percent", gross.key, rulebook),
            _kpi_row(day, "SERVICE", source, lob, language, "sl_adjusted", adjusted.label,
                     answered_target, offered - (short_abandoned or 0) if offered is not None else None,
                     evaluate_formula(adjusted.formula, values), "percent", adjusted.key, rulebook),
            _kpi_row(day, "SERVICE", source, lob, language, "service_availability", "Service availability",
                     answered, offered, evaluate_formula(rulebook.formulas["service_availability"].formula, values),
                     "percent", "answered_over_offered", rulebook),
            _kpi_row(day, "SERVICE", source, lob, language, "abandon_rate", "Abandon rate",
                     abandoned, offered, evaluate_formula(rulebook.formulas["abandon_rate"].formula, values),
                     "percent", "abandoned_over_offered", rulebook),
            _kpi_row(day, "SERVICE", source, lob, language, "aht_seconds", "Average handle time",
                     handled_seconds, answered, evaluate_formula(rulebook.formulas["aht_seconds"].formula, values),
                     "seconds", "weighted", rulebook),
            _kpi_row(day, "SERVICE", source, lob, language, "offered", "Offered contacts", offered, None, offered, "count", "sum", rulebook),
            _kpi_row(day, "SERVICE", source, lob, language, "answered", "Answered contacts", answered, None, answered, "count", "sum", rulebook),
        ])
    absence_rows = _query(conn, """
        SELECT business_date, lob, language, sum(planned_net_minutes)/60.0,
               sum(absence_minutes)/60.0, sum(vacation_minutes)/60.0,
               sum(shrinkage_minutes)/60.0
        FROM mart.absence_agent_day WHERE business_date BETWEEN ? AND ?
        GROUP BY business_date, lob, language ORDER BY business_date, lob, language
    """, [start, end])[1]
    for day, lob, language, planned, absence, vacation, shrinkage in absence_rows:
        values = {
            "planned_net_hours": planned, "absence_hours": absence,
            "vacation_hours": vacation, "shrinkage_hours": shrinkage,
        }
        for key, amount in (("absence_rate", absence), ("vacation_rate", vacation), ("shrinkage_rate", shrinkage)):
            formula = rulebook.formulas[key]
            output.append(_kpi_row(
                day, "ABSENCE", "LILO_STATUS", lob, language, key, formula.label,
                amount, planned, evaluate_formula(formula.formula, values), formula.unit, "rulebook", rulebook,
            ))
        output.extend([
            _kpi_row(day, "ABSENCE", "LILO_STATUS", lob, language, "planned_net_hours", "Planned net hours", planned, None, planned, "hours", "sum", rulebook),
            _kpi_row(day, "ABSENCE", "LILO_STATUS", lob, language, "absence_hours", "Absence hours", absence, None, absence, "hours", "sum", rulebook),
            _kpi_row(day, "ABSENCE", "LILO_STATUS", lob, language, "vacation_hours", "Vacation hours", vacation, None, vacation, "hours", "sum", rulebook),
        ])
    pcs_rows = _query(conn, """
        SELECT business_date, lob, language, sum(pcs_score_sum), sum(pcs_score_count),
               sum(survey_responses), sum(pcs_enabled_calls), sum(handle_seconds), sum(handled_calls)
        FROM mart.agent_pcs_day WHERE business_date BETWEEN ? AND ?
        GROUP BY business_date, lob, language ORDER BY business_date, lob, language
    """, [start, end])[1]
    for day, lob, language, score_sum, score_count, responses, eligible, handle_seconds, handled_calls in pcs_rows:
        output.extend([
            _kpi_row(day, "PCS", "CALL_BY_CALL", lob, language, "pcs_average", "PCS average", score_sum, score_count,
                     evaluate_formula(rulebook.formulas["pcs_average"].formula, {"pcs_score_sum": score_sum, "pcs_score_count": score_count}), "score", "response_weighted", rulebook),
            _kpi_row(day, "PCS", "CALL_BY_CALL", lob, language, "pcs_response_rate", "PCS response rate", responses, eligible,
                     evaluate_formula(rulebook.formulas["pcs_response_rate"].formula, {"survey_responses": responses, "pcs_enabled_calls": eligible}), "percent", "responses_over_enabled", rulebook),
            _kpi_row(day, "PCS", "CALL_BY_CALL", lob, language, "agent_aht_seconds", "Agent AHT", handle_seconds, handled_calls,
                     evaluate_formula(rulebook.formulas["agent_aht_seconds"].formula, {"handle_seconds": handle_seconds, "handled_calls": handled_calls}), "seconds", "weighted", rulebook),
        ])
    forecast_rows = _query(conn, """
        SELECT business_date, comparison_scope, sum(volume_forecast),
               CASE WHEN sum(volume_forecast)>0 THEN sum(volume_forecast*aht_forecast_seconds)/sum(volume_forecast) END
        FROM mart.forecast_hour WHERE business_date BETWEEN ? AND ?
        GROUP BY business_date, comparison_scope ORDER BY business_date, comparison_scope
    """, [start, end])[1]
    for day, queue, volume, aht in forecast_rows:
        output.extend([
            _kpi_row(day, "FORECAST", "VERINT", queue, None, "forecast_volume", "Forecast volume", volume, None, volume, "count", "sum", rulebook),
            _kpi_row(day, "FORECAST", "VERINT", queue, None, "forecast_aht_seconds", "Forecast AHT", None, None, aht, "seconds", "volume_weighted", rulebook),
        ])
    return output


def build_scorecard_report(
    conn: DatabaseConnection,
    config: Config,
    start: date,
    end: date,
    output: Path | None = None,
) -> Path:
    rulebook = load_rulebook(config.home, config.business_rules)
    generated = datetime.now()
    pack = report_pack("scorecard")
    output = (
        output or report_pack_folder(config, pack.key)
        / f"{pack.filename_prefix}_{start:%Y-%m-%d}_to_{end:%Y-%m-%d}_{generated:%H%M%S_%f}.xlsx"
    ).resolve()
    output.parent.mkdir(parents=True, exist_ok=True)
    partial = output.with_name(f"{output.stem}.partial{output.suffix}")
    report = ExcelReport(partial)
    try:
        _start_sheet(report, "WFMHub Executive Scorecard", rulebook, config, start, end, ("KPI_DAILY", "SERVICE_INTERVALS", "ABSENCE_DAILY"))
        kpi_rows = _scorecard_rows(conn, rulebook, start, end)
        headers = [
            "business_date", "iso_year", "iso_week", "calendar_month", "domain", "source",
            "lob", "language", "kpi_key", "kpi_name", "numerator", "denominator",
            "value", "unit", "calculation_profile", "rule_version", "rule_sha256",
        ]
        report.add_table_sheet(
            "KPI_DAILY", "Pivot-ready daily KPI scorecard",
            "One KPI per scope/day with numerator and denominator retained. Filter by Domain and KPI Name before pivoting.",
            headers, kpi_rows,
        )
        headers, rows = _query(conn, """
            SELECT business_date, interval_start, source_system, service_scope,
                   comparison_scope, mapping_status, queue, business_partner, lob,
                   language, offered, answered, abandoned, short_abandoned,
                   answered_within_target, sl_gross, sl_adjusted, sl_profile,
                   service_level, service_availability, abandon_rate, aht_seconds,
                   rule_version, source_file
            FROM mart.service_interval WHERE business_date BETWEEN ? AND ?
            ORDER BY business_date, interval_start, source_system, queue LIMIT ?
        """, [start, end, config.report_limits.get("max_service_rows", 100000)])
        report.add_table_sheet(
            "SERVICE_INTERVALS", "Service performance intervals",
            "Availability here is service availability only. Both gross and adjusted SL are retained.", headers, rows,
        )
        component_rows = _query(conn, """
            SELECT business_date, lob, language, sum(planned_net_minutes),
                   sum(absence_minutes), sum(vacation_minutes), sum(shrinkage_minutes)
            FROM mart.absence_agent_day WHERE business_date BETWEEN ? AND ?
            GROUP BY business_date, lob, language ORDER BY business_date, lob, language
        """, [start, end])[1]
        headers = [
            "business_date", "lob", "language", "planned_net_hours", "absence_hours",
            "vacation_hours", "shrinkage_hours", "absence_rate", "vacation_rate", "shrinkage_rate",
        ]
        rows = []
        for day, lob, language, planned, absence, vacation, shrinkage in component_rows:
            values = _absence_values(rulebook, planned, absence, vacation, shrinkage)
            rows.append((
                day, lob, language, values["planned_net_hours"], values["absence_hours"],
                values["vacation_hours"], values["shrinkage_hours"], values["absence_rate"],
                values["vacation_rate"], values["shrinkage_rate"],
            ))
        report.add_table_sheet("ABSENCE_DAILY", "Daily absence and shrinkage", "Ratio-of-sums daily facts for your PivotTables.", headers, rows)
        _add_rule_sheets(report, rulebook)
        headers, rows = _query(conn, """
            SELECT source_family, newest_file, newest_business_date, row_count,
                   scoped_out_count, status, details FROM mart.source_health ORDER BY source_family
        """)
        report.add_table_sheet("SOURCE_HEALTH", "All source health", "Validate source coverage before using the scorecard.", headers, rows, exception_column="Status")
    except Exception:
        report.close()
        partial.unlink(missing_ok=True)
        raise
    else:
        report.close()
        partial.replace(output)
    return output
