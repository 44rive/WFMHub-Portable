"""Deterministic, evidence-backed WFM findings generated only with Python."""

from __future__ import annotations

import hashlib
import json
import shutil
import tomllib
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Any, Mapping

from .database import DatabaseConnection
from .metrics import MetricCatalog, MetricMethod
from .semantic import MetricAggregate, aggregate_metric_values


class AnalyticsRulesError(RuntimeError):
    pass


@dataclass(frozen=True)
class FindingThreshold:
    warning_delta: float
    critical_delta: float
    trend_delta: float


@dataclass(frozen=True)
class AnalyticsRules:
    file: Path
    version: str
    description: str
    sha256: str
    max_findings: int
    include_information: bool
    defaults: Mapping[str, FindingThreshold]
    metric_thresholds: Mapping[str, FindingThreshold]

    def threshold_for(self, metric_id: str, unit: str) -> FindingThreshold:
        return self.metric_thresholds.get(
            metric_id,
            self.defaults.get(unit, self.defaults["number"]),
        )


@dataclass(frozen=True)
class Finding:
    finding_id: str
    finding_type: str
    severity: str
    domain: str
    metric_id: str | None
    method_id: str | None
    period_start: date
    period_end: date
    dimensions: Mapping[str, Any]
    title: str
    summary: str
    current_value: float | None
    reference_value: float | None
    target_value: float | None
    delta_value: float | None
    unit: str | None
    evidence_dataset: str
    evidence_filter: str
    score: float


def _threshold(item: Mapping[str, Any], label: str) -> FindingThreshold:
    try:
        warning = float(item["warning_delta"])
        critical = float(item["critical_delta"])
        trend = float(item["trend_delta"])
    except (KeyError, TypeError, ValueError) as exc:
        raise AnalyticsRulesError(f"{label} requires numeric warning_delta, critical_delta and trend_delta") from exc
    if min(warning, critical, trend) < 0 or critical < warning:
        raise AnalyticsRulesError(f"{label} thresholds must be non-negative and critical >= warning")
    return FindingThreshold(warning, critical, trend)


def ensure_analytics_rules(home: Path, target: Path | None = None) -> Path:
    default = home / "config" / "default_analytics.toml"
    target = target or home / "config" / "analytics_rules.toml"
    if not default.exists():
        raise AnalyticsRulesError(f"Default analytics rules are missing: {default}")
    if not target.exists():
        target.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(default, target)
    return target


def load_analytics_rules(home: Path, file: Path | None = None) -> AnalyticsRules:
    file = ensure_analytics_rules(home, file).resolve()
    try:
        content = file.read_bytes()
        raw = tomllib.loads(content.decode("utf-8"))
    except (OSError, UnicodeDecodeError, tomllib.TOMLDecodeError) as exc:
        raise AnalyticsRulesError(f"Cannot read analytics rules {file}: {exc}") from exc
    meta = raw.get("analytics")
    if not isinstance(meta, dict) or not str(meta.get("version", "")).strip():
        raise AnalyticsRulesError("Missing analytics.version")
    defaults_raw = raw.get("defaults")
    if not isinstance(defaults_raw, dict):
        raise AnalyticsRulesError("Missing [defaults] analytics thresholds")
    defaults = {
        str(unit): _threshold(item, f"defaults.{unit}")
        for unit, item in defaults_raw.items() if isinstance(item, dict)
    }
    if "number" not in defaults:
        raise AnalyticsRulesError("Analytics defaults must include [defaults.number]")
    overrides: dict[str, FindingThreshold] = {}
    for index, item in enumerate(raw.get("metric_thresholds", []), 1):
        if not isinstance(item, dict) or not str(item.get("metric_id", "")).strip():
            raise AnalyticsRulesError(f"metric_thresholds item {index} requires metric_id")
        metric_id = str(item["metric_id"])
        if metric_id in overrides:
            raise AnalyticsRulesError(f"Duplicate metric threshold for {metric_id}")
        overrides[metric_id] = _threshold(item, f"metric_thresholds.{metric_id}")
    maximum = int(meta.get("max_findings", 250))
    if not 1 <= maximum <= 10_000:
        raise AnalyticsRulesError("analytics.max_findings must be between 1 and 10000")
    return AnalyticsRules(
        file=file, version=str(meta["version"]), description=str(meta.get("description", "")),
        sha256=hashlib.sha256(content).hexdigest(), max_findings=maximum,
        include_information=bool(meta.get("include_information", True)),
        defaults=defaults, metric_thresholds=overrides,
    )


def validate_analytics_rules(rules: AnalyticsRules, catalog: MetricCatalog | None = None) -> list[str]:
    if catalog is not None:
        unknown = sorted(set(rules.metric_thresholds) - set(catalog.metric_ids()))
        if unknown:
            raise AnalyticsRulesError(f"Analytics thresholds reference unknown metrics: {', '.join(unknown)}")
    return [
        f"Analytics rules {rules.version} are valid.",
        f"SHA-256: {rules.sha256}",
        f"{len(rules.metric_thresholds)} metric overrides; maximum {rules.max_findings} findings.",
    ]


def _format(value: float | None, unit: str | None) -> str:
    if value is None:
        return "no data"
    if unit == "percent":
        return f"{value:.1%}"
    if unit == "seconds":
        return f"{value:,.1f} seconds"
    if unit == "fte":
        return f"{value:,.2f} FTE"
    if unit == "score":
        return f"{value:,.2f}"
    return f"{value:,.2f}"


def _scope_text(dimensions: Mapping[str, Any]) -> str:
    labels = [
        str(dimensions.get(key)) for key in ("source_system", "lob", "language", "team_leader")
        if dimensions.get(key) not in {None, "", "(blank)"}
    ]
    return " / ".join(labels) if labels else "all scope"


def _identity(*parts: Any) -> str:
    return hashlib.sha256("|".join(str(part or "") for part in parts).encode("utf-8")).hexdigest()


def _method_for(catalog: MetricCatalog, aggregate: MetricAggregate, day: date) -> MetricMethod:
    method = catalog.method_for(aggregate.metric_id, day, aggregate.dimensions)
    if method is None or method.method_id != aggregate.method_id:
        raise AnalyticsRulesError(
            f"Could not resolve {aggregate.metric_id}.{aggregate.method_id} for analytics on {day}"
        )
    return method


def _severity(delta: float, threshold: FindingThreshold) -> str:
    if delta >= threshold.critical_delta:
        return "CRITICAL"
    if delta >= threshold.warning_delta:
        return "WARNING"
    return "INFO"


def _target_findings(
    catalog: MetricCatalog,
    rules: AnalyticsRules,
    aggregate: MetricAggregate,
    day: date,
) -> list[Finding]:
    method = _method_for(catalog, aggregate, day)
    scope = _scope_text(aggregate.dimensions)
    if aggregate.state == "LOW_SAMPLE":
        title = f"Low sample: {method.label} — {scope}"
        summary = (
            f"{method.label} has sample {_format(aggregate.sample_size, 'number')}, below the configured "
            f"minimum {method.minimum_sample:g}. The calculated value {_format(aggregate.value, method.unit)} "
            "is retained but should not be over-interpreted."
        )
        return [Finding(
            _identity("low_sample", method.metric_id, method.method_id, day, scope),
            "LOW_SAMPLE", "INFO", method.domain, method.metric_id, method.method_id,
            day, day, aggregate.dimensions, title, summary, aggregate.value, None,
            method.target, None, method.unit, "metric_value",
            json.dumps({**aggregate.dimensions, "business_date": str(day), "metric_id": method.metric_id}, sort_keys=True),
            0.1,
        )]
    if aggregate.value is None or method.target is None or method.direction == "neutral":
        return []
    delta = (
        method.target - aggregate.value
        if method.direction == "higher_is_better"
        else aggregate.value - method.target
    )
    if delta <= 0:
        return []
    threshold = rules.threshold_for(method.metric_id, method.unit)
    severity = _severity(delta, threshold)
    if severity == "INFO" and not rules.include_information:
        return []
    title = f"{method.label} missed target — {scope}"
    relation = "below" if method.direction == "higher_is_better" else "above"
    summary = (
        f"{method.label} was {_format(aggregate.value, method.unit)}, {_format(delta, method.unit)} "
        f"{relation} the target of {_format(method.target, method.unit)}. "
        f"The result uses {_format(aggregate.numerator, 'number')} numerator and "
        f"{_format(aggregate.denominator, 'number')} denominator components under method {method.method_id}."
    )
    return [Finding(
        _identity("target", method.metric_id, method.method_id, day, scope),
        "TARGET_BREACH", severity, method.domain, method.metric_id, method.method_id,
        day, day, aggregate.dimensions, title, summary, aggregate.value, None,
        method.target, delta, method.unit, "metric_value",
        json.dumps({**aggregate.dimensions, "business_date": str(day), "metric_id": method.metric_id}, sort_keys=True),
        delta / max(threshold.warning_delta, 1e-12),
    )]


def _trend_findings(
    catalog: MetricCatalog,
    rules: AnalyticsRules,
    aggregates: list[MetricAggregate],
) -> list[Finding]:
    by_scope: dict[tuple[Any, ...], list[MetricAggregate]] = {}
    for item in aggregates:
        scope = tuple(sorted((key, value) for key, value in item.dimensions.items() if key != "business_date"))
        by_scope.setdefault((item.metric_id, item.method_id, scope), []).append(item)
    findings: list[Finding] = []
    for (_, _, _), rows in by_scope.items():
        rows.sort(key=lambda item: item.dimensions["business_date"])
        usable = [item for item in rows if item.value is not None and item.state != "LOW_SAMPLE"]
        if len(usable) < 2:
            continue
        previous, current = usable[-2], usable[-1]
        current_day = current.dimensions["business_date"]
        previous_day = previous.dimensions["business_date"]
        method = _method_for(catalog, current, current_day)
        delta = float(current.value) - float(previous.value)
        threshold = rules.threshold_for(method.metric_id, method.unit)
        if abs(delta) < threshold.trend_delta:
            continue
        harmful = (
            delta < 0 if method.direction == "higher_is_better"
            else delta > 0 if method.direction == "lower_is_better"
            else False
        )
        severity = "WARNING" if harmful else "INFO"
        if severity == "INFO" and not rules.include_information:
            continue
        scope = _scope_text(current.dimensions)
        direction = "increased" if delta > 0 else "decreased"
        assessment = "unfavorable" if harmful else "favorable" if method.direction != "neutral" else "material"
        title = f"{method.label} {direction} — {scope}"
        summary = (
            f"{method.label} {direction} from {_format(previous.value, method.unit)} on {previous_day} "
            f"to {_format(current.value, method.unit)} on {current_day}; the change of "
            f"{_format(abs(delta), method.unit)} is {assessment}."
        )
        findings.append(Finding(
            _identity("trend", method.metric_id, method.method_id, current_day, scope),
            "PERIOD_CHANGE", severity, method.domain, method.metric_id, method.method_id,
            previous_day, current_day, current.dimensions, title, summary,
            current.value, previous.value, method.target, delta, method.unit,
            "metric_value",
            json.dumps({key: value for key, value in current.dimensions.items() if key != "business_date"}, sort_keys=True, default=str),
            abs(delta) / max(threshold.trend_delta, 1e-12),
        ))
    return findings


def _source_findings(conn: DatabaseConnection, start: date, end: date) -> list[Finding]:
    output: list[Finding] = []
    for family, last_date, status, details in conn.execute(
        """SELECT source_family, newest_business_date, status, details
           FROM mart.source_health WHERE status IN ('ERROR','MISSING','EMPTY','STALE')
           ORDER BY CASE status WHEN 'ERROR' THEN 1 WHEN 'MISSING' THEN 2 ELSE 3 END,
                    source_family"""
    ):
        severity = "CRITICAL" if status == "ERROR" else "WARNING"
        title = f"Source {family} is {str(status).lower()}"
        summary = (
            f"{family} source health is {status}. Latest available business date: {last_date or 'none'}. "
            f"{details or 'Review the configured path and latest refresh log.'}"
        )
        output.append(Finding(
            _identity("source", family, status, end), "SOURCE_HEALTH", severity,
            "data_quality", None, None, start, end, {"source_system": family},
            title, summary, None, None, None, None, None, "source_health",
            json.dumps({"source_family": family}, sort_keys=True), 10 if severity == "CRITICAL" else 5,
        ))
    return output


_FINDING_COLUMNS = (
    "finding_id", "run_id", "finding_rank", "finding_type", "severity", "domain",
    "metric_id", "method_id", "period_start", "period_end", "source_system",
    "lob", "language", "team_leader", "agent_id", "title", "summary",
    "current_value", "reference_value", "target_value", "delta_value", "unit",
    "evidence_dataset", "evidence_filter", "catalog_version", "catalog_sha256",
    "analytics_version", "analytics_sha256", "created_at",
)


def build_findings(
    conn: DatabaseConnection,
    catalog: MetricCatalog,
    rules: AnalyticsRules,
    run_id: str,
    start: date,
    end: date,
) -> int:
    conn.execute("DELETE FROM mart.analysis_finding")
    findings = _source_findings(conn, start, end)
    for method in catalog.methods:
        dimensions = tuple(dict.fromkeys(("business_date", *method.finding_dimensions)))
        aggregates = aggregate_metric_values(
            conn, catalog, start, end, [method.metric_id], dimensions,
        )
        aggregates = [item for item in aggregates if item.method_id == method.method_id]
        for aggregate in aggregates:
            findings.extend(_target_findings(
                catalog, rules, aggregate, aggregate.dimensions["business_date"],
            ))
        findings.extend(_trend_findings(catalog, rules, aggregates))
    severity_order = {"CRITICAL": 0, "WARNING": 1, "INFO": 2}
    unique: dict[str, Finding] = {}
    for finding in findings:
        existing = unique.get(finding.finding_id)
        if existing is None or finding.score > existing.score:
            unique[finding.finding_id] = finding
    ordered = sorted(
        unique.values(), key=lambda item: (severity_order.get(item.severity, 9), -item.score, item.title)
    )[:rules.max_findings]
    rows: list[tuple[Any, ...]] = []
    now = datetime.now()
    for rank, finding in enumerate(ordered, 1):
        dims = finding.dimensions
        rows.append((
            finding.finding_id, run_id, rank, finding.finding_type, finding.severity,
            finding.domain, finding.metric_id, finding.method_id, finding.period_start,
            finding.period_end, dims.get("source_system"), dims.get("lob"),
            dims.get("language"), dims.get("team_leader"), dims.get("agent_id"),
            finding.title, finding.summary, finding.current_value,
            finding.reference_value, finding.target_value, finding.delta_value,
            finding.unit, finding.evidence_dataset, finding.evidence_filter,
            catalog.version, catalog.sha256, rules.version, rules.sha256, now,
        ))
    if rows:
        placeholders = ", ".join("?" for _ in _FINDING_COLUMNS)
        conn.executemany(
            f"INSERT INTO mart.analysis_finding ({', '.join(_FINDING_COLUMNS)}) VALUES ({placeholders})",
            rows,
        )
    conn.execute(
        """INSERT INTO meta.analytics_application(
               run_id, analytics_version, analytics_sha256, analytics_file, applied_at
           ) VALUES (?, ?, ?, ?, ?)
           ON CONFLICT(run_id) DO UPDATE SET
               analytics_version=excluded.analytics_version,
               analytics_sha256=excluded.analytics_sha256,
               analytics_file=excluded.analytics_file, applied_at=excluded.applied_at""",
        [run_id, rules.version, rules.sha256, str(rules.file), now],
    )
    return len(rows)
