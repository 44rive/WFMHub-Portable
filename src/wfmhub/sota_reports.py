"""Governance reference workbook generated from the active catalogs."""

from __future__ import annotations

from pathlib import Path
from typing import Any

from .analytics import load_analytics_rules
from .config import Config
from .metrics import MetricCatalog, load_metric_catalog
from .report_specs import load_report_catalog
from .reports import ExcelReport
from .rules import Rulebook, load_rulebook


def _metric_rows(catalog: MetricCatalog) -> list[tuple[Any, ...]]:
    rows: list[tuple[Any, ...]] = []
    for method in catalog.methods:
        scope = "; ".join(
            f"{key}={','.join(values)}" for key, values in sorted(method.scope.items())
        ) or "ALL"
        rows.append((
            method.metric_id, method.method_id, method.label, method.domain,
            method.source_model, method.grain, method.unit, method.aggregation,
            method.numerator, method.denominator, method.sample, method.target,
            method.direction, method.minimum_sample, method.effective_from,
            method.effective_to, method.priority, scope, method.description,
            catalog.version, catalog.sha256,
        ))
    return rows


def _add_catalog_sheets(
    report: ExcelReport,
    catalog: MetricCatalog,
    rulebook: Rulebook,
    config: Config,
) -> None:
    analytics = load_analytics_rules(config.home, config.analytics_rules)
    reports = load_report_catalog(config.home, config.report_catalog)
    report.add_table_sheet(
        "METRIC_METHODS", "Canonical metric methods",
        "Every active KPI formula, target, scope, aggregation rule and effective date is listed here.",
        [
            "metric_id", "method_id", "label", "domain", "source_model", "grain",
            "unit", "aggregation", "numerator", "denominator", "sample", "target",
            "direction", "minimum_sample", "effective_from", "effective_to",
            "priority", "scope", "description", "catalog_version", "catalog_sha256",
        ],
        _metric_rows(catalog),
    )
    report.add_table_sheet(
        "ACTIVITY_RULES", "Attendance and absence evidence rules",
        "First matching rule wins. These rules classify evidence; they do not calculate KPI rates.",
        [
            "order", "rule_name", "category", "patterns", "match", "planned", "working",
            "counts_as_absence", "counts_as_vacation", "counts_as_unpaid",
            "counts_as_shrinkage", "rule_version", "rule_sha256",
        ],
        [
            (
                index, item.name, item.category, " | ".join(item.patterns), item.match,
                item.planned, item.working, item.absence, item.vacation, item.unpaid,
                item.shrinkage, rulebook.version, rulebook.sha256,
            )
            for index, item in enumerate(rulebook.activity_rules, 1)
        ],
    )
    report.add_table_sheet(
        "ANALYTICS_RULES", "Deterministic finding thresholds",
        "Python findings compare governed values with targets and prior periods. No AI is called.",
        ["scope", "warning_delta", "critical_delta", "trend_delta", "version", "sha256"],
        [
            (f"unit:{unit}", item.warning_delta, item.critical_delta, item.trend_delta,
             analytics.version, analytics.sha256)
            for unit, item in analytics.defaults.items()
        ] + [
            (f"metric:{metric_id}", item.warning_delta, item.critical_delta, item.trend_delta,
             analytics.version, analytics.sha256)
            for metric_id, item in analytics.metric_thresholds.items()
        ],
    )
    report.add_table_sheet(
        "REPORT_CONTRACTS", "Workbook presentation contracts",
        "Report composition is separate from formulas and source parsing.",
        ["pack", "title", "purpose", "finding_domains", "ordered_sheets", "version", "sha256"],
        [
            (key, spec.title, spec.purpose, " | ".join(spec.finding_domains),
             " | ".join(spec.sheets), reports.version, reports.sha256)
            for key, spec in reports.packs.items()
        ],
    )


def build_kpi_catalog(config: Config, output: Path | None = None) -> Path:
    catalog = load_metric_catalog(config.home, config.metric_catalog)
    rulebook = load_rulebook(config.home, config.business_rules)
    output = (
        output or config.output / "reference" / f"WFMHub_Governance_Catalog_{catalog.version}.xlsx"
    ).resolve()
    output.parent.mkdir(parents=True, exist_ok=True)
    partial = output.with_name(f"{output.stem}.partial{output.suffix}")
    report = ExcelReport(partial)
    try:
        _add_catalog_sheets(report, catalog, rulebook, config)
    except Exception:
        report.close()
        partial.unlink(missing_ok=True)
        raise
    else:
        report.close()
        partial.replace(output)
    return output
