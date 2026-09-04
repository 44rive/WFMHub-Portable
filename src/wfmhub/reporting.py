"""Shared workbook sheets for findings, methods and provenance."""

from __future__ import annotations

from datetime import date
from typing import Any, Sequence

from .analytics import load_analytics_rules
from .config import Config
from .database import DatabaseConnection
from .mapping import load_queue_mapping
from .metrics import load_metric_catalog
from .reports import ExcelReport
from .report_specs import ReportSpec, load_report_catalog
from .rules import load_rulebook


def report_spec(config: Config, key: str) -> ReportSpec:
    return load_report_catalog(config.home, config.report_catalog).pack(key)


def validate_workbook_contract(report: ExcelReport, spec: ReportSpec) -> None:
    actual = tuple(sheet.get_name() for sheet in report.workbook.worksheets())
    if actual != spec.sheets:
        raise ValueError(
        f"Report contract {spec.key!r} does not match the workbook. "
        f"Expected {spec.sheets}; created {actual}."
        )


def add_findings_sheet(
    report: ExcelReport,
    conn: DatabaseConnection,
    spec: ReportSpec,
    start: date,
    end: date,
) -> None:
    if not spec.includes("FINDINGS"):
        return
    domains = tuple(spec.finding_domains)
    where = "period_end BETWEEN ? AND ?"
    parameters: list[Any] = [start, end]
    if domains:
        where += f" AND domain IN ({', '.join('?' for _ in domains)})"
        parameters.extend(domains)
    cursor = conn.execute(
        f"""SELECT finding_rank, severity, finding_type, domain, metric_id,
                   period_start, period_end, source_system, lob, language,
                   team_leader, title, summary, current_value, reference_value,
                   target_value, delta_value, unit, evidence_dataset,
                   evidence_filter, catalog_version, analytics_version
            FROM mart.analysis_finding WHERE {where}
            ORDER BY CASE severity WHEN 'CRITICAL' THEN 1 WHEN 'WARNING' THEN 2 ELSE 3 END,
                     finding_rank""",
        parameters,
    )
    headers = [item[0] for item in cursor.description]
    rows = cursor.fetchall()
    ws = report.add_table_sheet(
        "FINDINGS", "Period findings",
        "Every statement includes the metric, comparison and evidence filter behind it.",
        headers, rows,
        exception_column="Severity",
    )
    if rows:
        display = [header.replace("_", " ").title() for header in headers]
        severity_col = display.index("Severity")
        ws.conditional_format(4, severity_col, 3 + len(rows), severity_col, {
            "type": "text", "criteria": "containing", "value": "CRITICAL",
            "format": report.error,
        })


def add_methods_sheet(
    report: ExcelReport,
    config: Config,
    spec: ReportSpec,
) -> None:
    if not spec.includes("METHODS"):
        return
    catalog = load_metric_catalog(config.home, config.metric_catalog)
    domains = set(spec.finding_domains)
    methods = [method for method in catalog.methods if method.domain in domains]
    headers = [
        "metric_id", "method_id", "label", "domain", "source_model", "grain",
        "aggregation", "numerator", "denominator", "sample", "unit", "target",
        "direction", "minimum_sample", "effective_from", "effective_to", "priority",
        "scope", "catalog_version", "catalog_sha256",
    ]
    rows = [
        (
            method.metric_id, method.method_id, method.label, method.domain,
            method.source_model, method.grain, method.aggregation, method.numerator,
            method.denominator, method.sample, method.unit, method.target,
            method.direction, method.minimum_sample, method.effective_from,
            method.effective_to, method.priority,
            "; ".join(f"{key}={','.join(values)}" for key, values in method.scope.items()) or "all",
            catalog.version, catalog.sha256,
        )
        for method in methods
    ]
    report.add_table_sheet(
        "METHODS", "Effective-dated metric methods",
        f"Loaded from {catalog.file.name}. Edit and validate the catalog; workbooks never own KPI arithmetic.",
        headers, rows,
    )


def add_domain_rules_sheet(
    report: ExcelReport,
    config: Config,
    spec: ReportSpec,
) -> None:
    """Document invariant evidence rules from one shared source."""
    if not spec.includes("DOMAIN_RULES"):
        return
    rulebook = load_rulebook(config.home, config.business_rules)
    headers = ["rule", "configuration", "grain", "guardrail"]
    if spec.key == "operations":
        rows = [
            ("Attendance evidence", "StartEndTimes + LILO + Agent Status", "agent/day",
             "Missing evidence is unknown and never a no-show"),
            ("Late tolerance", f"{rulebook.late_tolerance_minutes} minutes", "agent/day",
             "Running shifts remain provisional"),
            ("Staffing interval", "observed agent-seconds / 900", "LOB/language/15 minutes",
             "Future or missing evidence stays blank"),
            ("No-show", "completed shift + blank LILO row, or sufficient all-Logged-Off Agent Status coverage", "agent/day",
             "All gates must pass"),
        ]
    elif spec.key == "quality_pcs":
        allowed = ", ".join(f"{value:g}" for value in rulebook.pcs_allowed_scores)
        rows = [
            ("Valid primary score", f"Inbound Q{rulebook.pcs_primary_score_question} in {{{allowed}}}", "call leg",
             "Blank and invalid values are excluded from score counters"),
            ("Participation numerator", f"Inbound raw Q{rulebook.pcs_participation_question} nonblank", "call leg",
             "Invalid nonblank answers still participate"),
            ("Participation denominator", f"Inbound PCSStatus={rulebook.pcs_participation_status}", "call leg",
             "PostCallSurveyMode is diagnostic only"),
            ("Positive/negative boundary", f"positive > {rulebook.pcs_negative_score_maximum:g}; negative <= {rulebook.pcs_negative_score_maximum:g}", "valid primary response",
             "Counts are retained before ratios"),
        ]
    elif spec.key == "corrections":
        rows = [
            ("Observed gap", "scheduled time minus unioned LILO/Agent Status evidence", "agent/time segment",
             "Verint Activities never create the initial gap"),
            ("Corrected overlap", "unioned classified Verint Activities intersected with observed gap", "agent/time segment",
             "Overlapping Activities are not double-counted"),
            ("Residual tolerance", f"ignore residuals <= {rulebook.verint_match_tolerance_minutes} minutes", "correction segment",
             "The original correction ID remains the decision key"),
            ("Current-day tail", "future and is_gap=false", "agent/time segment",
             "An unfinished shift cannot become final early leave"),
        ]
    else:
        rows = [
            ("Planned net cap", f"min(shift span, {rulebook.standard_day_hours:g} hours)", "agent/day",
             "The standard day is a net payroll cap"),
            ("Final ledger", "classified Verint Activities clipped to schedule", "activity interval",
             "LILO and Agent Status do not create final payroll absence"),
            ("Overlap handling", "union intervals before category totals", "agent/day",
             "Event evidence rows must never be summed directly"),
            ("Unmapped activity", f"review required={rulebook.unmapped_activity_is_error}", "activity interval",
             "Unmapped labels remain visible"),
        ]
    report.add_table_sheet(
        "DOMAIN_RULES", "Invariant domain and evidence rules",
        f"Current rules from {rulebook.file.name} version {rulebook.version}.",
        headers, rows,
    )


def add_provenance_sheet(
    report: ExcelReport,
    conn: DatabaseConnection,
    config: Config,
    spec: ReportSpec,
    start: date,
    end: date,
) -> None:
    if not spec.includes("PROVENANCE"):
        return
    rules = load_rulebook(config.home, config.business_rules)
    metrics = load_metric_catalog(config.home, config.metric_catalog)
    analytics = load_analytics_rules(config.home, config.analytics_rules)
    mapping = load_queue_mapping(config.queue_mapping)
    latest = conn.execute(
        """SELECT run_id, finished_at, status, details FROM meta.refresh_run
           WHERE status='SUCCESS' ORDER BY finished_at DESC LIMIT 1"""
    ).fetchone()
    rows: Sequence[tuple[Any, ...]] = [
        ("Report contract", spec.key, spec.purpose),
        ("Selected period", f"{start} to {end}", "Explicit report boundary"),
        ("Refresh run", latest[0] if latest else None, latest[3] if latest else "No successful refresh metadata"),
        ("Refresh finished", latest[1] if latest else None, latest[2] if latest else None),
        ("Domain rulebook", rules.version, rules.sha256),
        ("Metric catalog", metrics.version, metrics.sha256),
        ("Analytics rules", analytics.version, analytics.sha256),
        ("Queue mapping", mapping.file.name, mapping.sha256),
        ("Report catalog", load_report_catalog(config.home, config.report_catalog).version,
         load_report_catalog(config.home, config.report_catalog).sha256),
        ("Calculation source", "WFM Hub metric catalog", "Workbook values follow the listed KPI method"),
    ]
    report.add_table_sheet(
        "PROVENANCE", "Calculation and configuration provenance",
        "Technical controls used to reconcile the workbook to its source and configuration.",
        ["item", "value", "evidence"], rows,
    )


def add_governance_sheets(
    report: ExcelReport,
    conn: DatabaseConnection,
    config: Config,
    pack_key: str,
    start: date,
    end: date,
) -> ReportSpec:
    spec = report_spec(config, pack_key)
    add_findings_sheet(report, conn, spec, start, end)
    add_domain_rules_sheet(report, config, spec)
    add_methods_sheet(report, config, spec)
    add_provenance_sheet(report, conn, config, spec, start, end)
    return spec
