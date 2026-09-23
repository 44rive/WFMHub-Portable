"""Effective-dated, scoped and validated KPI catalog.

Domain models expose trusted additive components.  This module is the only
place that turns those components into configurable KPI values.  It does not
parse extracts or decide evidence states.
"""

from __future__ import annotations

import hashlib
import shutil
import tomllib
from dataclasses import dataclass
from datetime import date
from itertools import combinations
from pathlib import Path
from typing import Any, Iterable, Mapping

from .rules import RulebookError, evaluate_formula, formula_names


class MetricCatalogError(RulebookError):
    pass


_AGGREGATIONS = {"ratio_of_sums", "sum", "mean", "maximum", "minimum"}
_DIRECTIONS = {"higher_is_better", "lower_is_better", "neutral"}
_SCOPE_CONTAINS = {"lob_contains", "queue_contains", "business_partner_contains"}


def _metric_formula_names(expression: str, label: str) -> frozenset[str]:
    try:
        return formula_names(expression, label)
    except RulebookError as exc:
        raise MetricCatalogError(str(exc)) from exc


@dataclass(frozen=True)
class MetricMethod:
    metric_id: str
    method_id: str
    label: str
    description: str
    domain: str
    source_model: str
    grain: str
    unit: str
    aggregation: str
    numerator: str
    denominator: str | None
    sample: str | None
    target: float | None
    direction: str
    minimum_sample: float
    effective_from: date
    effective_to: date | None
    priority: int
    scope: Mapping[str, tuple[str, ...]]
    finding_dimensions: tuple[str, ...]

    @property
    def component_names(self) -> frozenset[str]:
        names = set(_metric_formula_names(self.numerator, f"{self.metric_id}.{self.method_id}.numerator"))
        if self.denominator:
            names.update(_metric_formula_names(self.denominator, f"{self.metric_id}.{self.method_id}.denominator"))
        if self.sample:
            names.update(_metric_formula_names(self.sample, f"{self.metric_id}.{self.method_id}.sample"))
        return frozenset(names)

    def applies_on(self, business_date: date) -> bool:
        return business_date >= self.effective_from and (
            self.effective_to is None or business_date <= self.effective_to
        )

    def matches_scope(self, dimensions: Mapping[str, Any]) -> bool:
        for key, allowed in self.scope.items():
            if not allowed:
                continue
            dimension_key = key.removesuffix("_contains") if key in _SCOPE_CONTAINS else key
            actual = " ".join(str(dimensions.get(dimension_key) or "").upper().split())
            if key in _SCOPE_CONTAINS:
                if not any(token in actual for token in allowed):
                    return False
            elif actual not in allowed:
                return False
        return True


@dataclass(frozen=True)
class MetricEvaluation:
    method: MetricMethod
    numerator: float | None
    denominator: float | None
    sample_size: float | None
    value: float | None
    state: str


@dataclass(frozen=True)
class MetricCatalog:
    file: Path
    version: str
    description: str
    sha256: str
    methods: tuple[MetricMethod, ...]

    def metric_ids(self) -> tuple[str, ...]:
        return tuple(dict.fromkeys(method.metric_id for method in self.methods))

    def methods_for_source(self, source_model: str) -> tuple[MetricMethod, ...]:
        return tuple(method for method in self.methods if method.source_model == source_model)

    def method_for(
        self,
        metric_id: str,
        business_date: date,
        dimensions: Mapping[str, Any],
    ) -> MetricMethod | None:
        matches = [
            method for method in self.methods
            if method.metric_id == metric_id
            and method.applies_on(business_date)
            and method.matches_scope(dimensions)
        ]
        if not matches:
            return None
        matches.sort(key=lambda method: method.priority, reverse=True)
        if len(matches) > 1 and matches[0].priority == matches[1].priority:
            methods = ", ".join(method.method_id for method in matches if method.priority == matches[0].priority)
            raise MetricCatalogError(
                f"Ambiguous metric methods for {metric_id} on {business_date}: {methods}. "
                "Give the more specific method a higher priority."
            )
        return matches[0]

    def explain(self, metric_id: str) -> list[str]:
        methods = [method for method in self.methods if method.metric_id == metric_id]
        if not methods:
            raise MetricCatalogError(
                f"Unknown metric {metric_id!r}. Available: {', '.join(self.metric_ids())}"
            )
        output = [f"Metric: {metric_id}"]
        for method in sorted(methods, key=lambda item: (item.effective_from, -item.priority, item.method_id)):
            scope = ", ".join(f"{key}={list(values)}" for key, values in method.scope.items()) or "all rows"
            period = f"{method.effective_from} to {method.effective_to or 'open'}"
            expression = method.numerator
            if method.denominator:
                expression = f"({method.numerator}) / ({method.denominator})"
            output.extend([
                f"  Method     : {method.method_id}",
                f"  Label      : {method.label}",
                f"  Source     : {method.source_model} at {method.grain}",
                f"  Effective  : {period}",
                f"  Scope      : {scope}",
                f"  Formula    : {expression}",
                f"  Aggregate  : {method.aggregation}",
                f"  Target     : {method.target if method.target is not None else 'none'} ({method.direction})",
            ])
        return output


def _date_value(value: Any, label: str, required: bool = True) -> date | None:
    if value in {None, ""}:
        if required:
            raise MetricCatalogError(f"{label} is required")
        return None
    if isinstance(value, date):
        return value
    try:
        return date.fromisoformat(str(value))
    except ValueError as exc:
        raise MetricCatalogError(f"{label} must be YYYY-MM-DD") from exc


def _scope(raw: Any, label: str) -> Mapping[str, tuple[str, ...]]:
    if raw is None:
        return {}
    if not isinstance(raw, dict):
        raise MetricCatalogError(f"{label} must be an inline TOML table")
    result: dict[str, tuple[str, ...]] = {}
    for key, values in raw.items():
        if not isinstance(values, list):
            raise MetricCatalogError(f"{label}.{key} must be a list")
        result[str(key)] = tuple(
            " ".join(str(value).upper().split()) for value in values if str(value).strip()
        )
    return result


def ensure_metric_catalog(home: Path, target: Path | None = None) -> Path:
    default = home / "config" / "default_metrics.toml"
    target = target or home / "config" / "metric_catalog.toml"
    if not default.exists():
        raise MetricCatalogError(f"Default metric catalog is missing: {default}")
    if not target.exists():
        target.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(default, target)
    return target


def load_metric_catalog(home: Path, file: Path | None = None) -> MetricCatalog:
    file = ensure_metric_catalog(home, file).resolve()
    try:
        content = file.read_bytes()
        raw = tomllib.loads(content.decode("utf-8"))
    except (OSError, UnicodeDecodeError, tomllib.TOMLDecodeError) as exc:
        raise MetricCatalogError(f"Cannot read metric catalog {file}: {exc}") from exc
    meta = raw.get("catalog")
    if not isinstance(meta, dict):
        raise MetricCatalogError("Missing [catalog] table")
    version = str(meta.get("version", "")).strip()
    if not version:
        raise MetricCatalogError("catalog.version cannot be blank")
    methods: list[MetricMethod] = []
    seen: set[tuple[str, str, date]] = set()
    for index, item in enumerate(raw.get("metrics", []), 1):
        if not isinstance(item, dict):
            raise MetricCatalogError(f"metrics item {index} must be a table")
        metric_id = str(item.get("id", "")).strip()
        method_id = str(item.get("method", "")).strip()
        source_model = str(item.get("source_model", "")).strip()
        numerator = str(item.get("numerator", "")).strip()
        aggregation = str(item.get("aggregation", "ratio_of_sums")).strip()
        denominator_text = str(item.get("denominator", "")).strip()
        denominator = denominator_text or None
        sample_text = str(item.get("sample", "")).strip()
        sample = sample_text or None
        effective_from = _date_value(item.get("effective_from"), f"metrics[{index}].effective_from")
        effective_to = _date_value(item.get("effective_to"), f"metrics[{index}].effective_to", required=False)
        if not metric_id or not method_id or not source_model or not numerator:
            raise MetricCatalogError(
                f"metrics item {index} requires id, method, source_model and numerator"
            )
        if aggregation not in _AGGREGATIONS:
            raise MetricCatalogError(f"{metric_id}.{method_id} has unsupported aggregation {aggregation!r}")
        if aggregation in {"ratio_of_sums", "mean"} and denominator is None:
            raise MetricCatalogError(f"{metric_id}.{method_id} requires a denominator")
        direction = str(item.get("direction", "neutral"))
        if direction not in _DIRECTIONS:
            raise MetricCatalogError(f"{metric_id}.{method_id} has unsupported direction {direction!r}")
        _metric_formula_names(numerator, f"{metric_id}.{method_id}.numerator")
        if denominator:
            _metric_formula_names(denominator, f"{metric_id}.{method_id}.denominator")
        if sample:
            _metric_formula_names(sample, f"{metric_id}.{method_id}.sample")
        if effective_to is not None and effective_to < effective_from:
            raise MetricCatalogError(f"{metric_id}.{method_id} effective_to precedes effective_from")
        identity = (metric_id, method_id, effective_from)
        if identity in seen:
            raise MetricCatalogError(f"Duplicate metric method version: {identity}")
        seen.add(identity)
        method = MetricMethod(
            metric_id=metric_id,
            method_id=method_id,
            label=str(item.get("label", metric_id)),
            description=str(item.get("description", "")),
            domain=str(item.get("domain", source_model)),
            source_model=source_model,
            grain=str(item.get("grain", "")),
            unit=str(item.get("unit", "number")),
            aggregation=aggregation,
            numerator=numerator,
            denominator=denominator,
            sample=sample,
            target=float(item["target"]) if "target" in item else None,
            direction=direction,
            minimum_sample=float(item.get("minimum_sample", 0)),
            effective_from=effective_from,
            effective_to=effective_to,
            priority=int(item.get("priority", 0)),
            scope=_scope(item.get("scope"), f"metrics[{index}].scope"),
            finding_dimensions=tuple(str(value) for value in item.get("finding_dimensions", ["lob", "language"])),
        )
        methods.append(method)
    if not methods:
        raise MetricCatalogError("At least one [[metrics]] method is required")
    return MetricCatalog(
        file=file,
        version=version,
        description=str(meta.get("description", "")),
        sha256=hashlib.sha256(content).hexdigest(),
        methods=tuple(methods),
    )


def evaluate_metric(method: MetricMethod, components: Mapping[str, Any]) -> MetricEvaluation:
    missing = sorted(method.component_names - components.keys())
    if missing:
        raise MetricCatalogError(
            f"{method.metric_id}.{method.method_id} requires unavailable components: {', '.join(missing)}"
        )
    numerator = evaluate_formula(method.numerator, components)
    denominator = evaluate_formula(method.denominator, components) if method.denominator else None
    sample = evaluate_formula(method.sample, components) if method.sample else denominator
    if method.aggregation in {"ratio_of_sums", "mean"}:
        value = None if numerator is None or denominator in {None, 0} else numerator / denominator
    else:
        value = numerator
    if value is None:
        state = "NO_DATA"
    elif method.minimum_sample and (sample is None or sample < method.minimum_sample):
        state = "LOW_SAMPLE"
    elif method.target is None or method.direction == "neutral":
        state = "UNASSESSED"
    elif method.direction == "higher_is_better":
        state = "ON_TARGET" if value >= method.target else "BELOW_TARGET"
    else:
        state = "ON_TARGET" if value <= method.target else "ABOVE_TARGET"
    return MetricEvaluation(method, numerator, denominator, sample, value, state)


def validate_metric_catalog(
    catalog: MetricCatalog,
    source_components: Mapping[str, Iterable[str]] | None = None,
) -> list[str]:
    def scopes_may_overlap(left: MetricMethod, right: MetricMethod) -> bool:
        def values(method: MetricMethod, dimension: str, contains: bool) -> tuple[str, ...]:
            key = f"{dimension}_contains" if contains else dimension
            return method.scope.get(key, ())

        dimensions = {
            key.removesuffix("_contains") for key in (*left.scope.keys(), *right.scope.keys())
        }
        for dimension in dimensions:
            left_exact, right_exact = values(left, dimension, False), values(right, dimension, False)
            left_contains, right_contains = values(left, dimension, True), values(right, dimension, True)
            if left_exact and right_exact and not set(left_exact).intersection(right_exact):
                return False
            if left_exact and right_contains and not any(
                token in exact for exact in left_exact for token in right_contains
            ):
                return False
            if right_exact and left_contains and not any(
                token in exact for exact in right_exact for token in left_contains
            ):
                return False
        return True

    for left, right in combinations(catalog.methods, 2):
        if left.metric_id != right.metric_id or left.priority != right.priority:
            continue
        periods_overlap = (
            (left.effective_to is None or right.effective_from <= left.effective_to)
            and (right.effective_to is None or left.effective_from <= right.effective_to)
        )
        if periods_overlap and scopes_may_overlap(left, right):
            raise MetricCatalogError(
                f"Ambiguous equal-priority methods for {left.metric_id}: "
                f"{left.method_id} and {right.method_id} can match the same date/scope. "
                "Close the date ranges, separate the scopes, or raise the specific method priority."
            )
    if source_components is not None:
        for method in catalog.methods:
            if method.source_model not in source_components:
                raise MetricCatalogError(
                    f"{method.metric_id}.{method.method_id} references unknown source_model {method.source_model!r}"
                )
            missing = method.component_names - set(source_components[method.source_model])
            if missing:
                raise MetricCatalogError(
                    f"{method.metric_id}.{method.method_id} references components not exposed by "
                    f"{method.source_model}: {', '.join(sorted(missing))}"
                )
    return [
        f"Metric catalog {catalog.version} is valid.",
        f"SHA-256: {catalog.sha256}",
        f"{len(catalog.metric_ids())} metrics; {len(catalog.methods)} effective-dated methods.",
    ]


def diff_metric_catalogs(before: MetricCatalog, after: MetricCatalog) -> list[str]:
    def signature(method: MetricMethod) -> tuple[Any, ...]:
        return (
            method.label, method.source_model, method.aggregation, method.numerator,
            method.denominator, method.sample, method.target, method.direction,
            method.minimum_sample, method.effective_to, method.priority,
            tuple(sorted(method.scope.items())),
        )

    old = {(item.metric_id, item.method_id, item.effective_from): item for item in before.methods}
    new = {(item.metric_id, item.method_id, item.effective_from): item for item in after.methods}
    lines = [f"Metric catalog diff: {before.version} -> {after.version}"]
    for key in sorted(old.keys() - new.keys()):
        lines.append(f"REMOVED {key[0]}.{key[1]} effective {key[2]}")
    for key in sorted(new.keys() - old.keys()):
        lines.append(f"ADDED   {key[0]}.{key[1]} effective {key[2]}")
    for key in sorted(old.keys() & new.keys()):
        if signature(old[key]) != signature(new[key]):
            lines.append(f"CHANGED {key[0]}.{key[1]} effective {key[2]}")
    if len(lines) == 1:
        lines.append("No method changes.")
    return lines
