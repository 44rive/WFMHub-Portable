"""Validated domain evidence rules and safe metric-expression evaluation."""

from __future__ import annotations

import ast
import hashlib
import math
import shutil
import tomllib
from dataclasses import dataclass
from datetime import date
from functools import lru_cache
from pathlib import Path
from typing import Any, Mapping


class RulebookError(RuntimeError):
    pass


@dataclass(frozen=True)
class ActivityRule:
    name: str
    category: str
    patterns: tuple[str, ...]
    match: str
    planned: bool
    working: bool
    absence: bool
    vacation: bool
    unpaid: bool
    shrinkage: bool


@dataclass(frozen=True)
class Rulebook:
    file: Path
    version: str
    effective_from: date
    description: str
    sha256: str
    standard_day_hours: float
    late_tolerance_minutes: int
    status_gap_tolerance_minutes: int
    verint_match_tolerance_minutes: int
    spell_gap_days: int
    cap_event_to_schedule: bool
    unmapped_activity_is_error: bool
    target_seconds: int
    activity_rules: tuple[ActivityRule, ...]
    pcs_scored_questions: tuple[int, ...]
    pcs_comment_questions: tuple[int, ...]
    pcs_survey_mode: str
    pcs_primary_score_question: int
    pcs_participation_question: int
    pcs_participation_status: str
    pcs_allowed_scores: tuple[float, ...]
    pcs_negative_score_maximum: float
    pcs_minimum_score: float
    pcs_maximum_score: float
    pcs_top_box_minimum: float
    pcs_low_score_maximum: float

    def classify_activity(self, value: str | None) -> ActivityRule | None:
        text = " ".join(str(value or "").upper().split())
        if not text:
            return None
        # Verint prefixes activities with values such as ".AP BEN |".
        suffix = text.split("|", 1)[1].strip() if "|" in text else text
        for rule in self.activity_rules:
            for raw_pattern in rule.patterns:
                pattern = " ".join(raw_pattern.upper().split())
                if rule.match == "exact" and suffix == pattern:
                    return rule
                if rule.match == "contains" and pattern in suffix:
                    return rule
                if rule.match == "exact_or_suffix" and (suffix == pattern or suffix.endswith(" " + pattern)):
                    return rule
        return None

_ALLOWED_BINOPS = (ast.Add, ast.Sub, ast.Mult, ast.Div)
_ALLOWED_UNARYOPS = (ast.UAdd, ast.USub)
_ALLOWED_CMPOPS = (ast.Eq, ast.NotEq, ast.Lt, ast.LtE, ast.Gt, ast.GtE)
_FUNCTIONS = {"min", "max", "coalesce", "nullif", "ifelse", "abs", "round"}


@lru_cache(maxsize=256)
def _validate_expression(expression: str, label: str) -> ast.Expression:
    try:
        tree = ast.parse(expression, mode="eval")
    except SyntaxError as exc:
        raise RulebookError(f"{label} has invalid formula syntax: {exc.msg}") from exc
    for node in ast.walk(tree):
        if isinstance(node, ast.Expression | ast.Load | ast.Constant | ast.Name):
            continue
        if isinstance(node, ast.BinOp) and isinstance(node.op, _ALLOWED_BINOPS):
            continue
        if isinstance(node, _ALLOWED_BINOPS):
            continue
        if isinstance(node, ast.UnaryOp) and isinstance(node.op, _ALLOWED_UNARYOPS):
            continue
        if isinstance(node, _ALLOWED_UNARYOPS):
            continue
        if isinstance(node, ast.Compare) and all(isinstance(operator, _ALLOWED_CMPOPS) for operator in node.ops):
            continue
        if isinstance(node, _ALLOWED_CMPOPS):
            continue
        if isinstance(node, ast.Call) and isinstance(node.func, ast.Name) and node.func.id in _FUNCTIONS and not node.keywords:
            continue
        raise RulebookError(f"{label} uses forbidden formula element: {type(node).__name__}")
    return tree


def _number(value: Any) -> float | None:
    if value is None:
        return None
    if isinstance(value, bool) or not isinstance(value, int | float):
        raise RulebookError(f"Formula received non-numeric value {value!r}")
    value = float(value)
    return value if math.isfinite(value) else None


def evaluate_formula(expression: str, values: Mapping[str, Any]) -> float | None:
    """Evaluate the deliberately tiny formula language without eval()."""
    tree = _validate_expression(expression, "formula")

    def visit(node: ast.AST) -> float | None:
        if isinstance(node, ast.Expression):
            return visit(node.body)
        if isinstance(node, ast.Constant):
            return _number(node.value)
        if isinstance(node, ast.Name):
            if node.id not in values:
                raise RulebookError(f"Formula references unknown value {node.id!r}")
            return _number(values[node.id])
        if isinstance(node, ast.UnaryOp):
            value = visit(node.operand)
            if value is None:
                return None
            return value if isinstance(node.op, ast.UAdd) else -value
        if isinstance(node, ast.BinOp):
            left, right = visit(node.left), visit(node.right)
            if left is None or right is None:
                return None
            if isinstance(node.op, ast.Add):
                return left + right
            if isinstance(node.op, ast.Sub):
                return left - right
            if isinstance(node.op, ast.Mult):
                return left * right
            return None if right == 0 else left / right
        if isinstance(node, ast.Compare):
            left = visit(node.left)
            if left is None:
                return 0.0
            for operator, comparator in zip(node.ops, node.comparators):
                right = visit(comparator)
                if right is None:
                    return 0.0
                matches = (
                    left == right if isinstance(operator, ast.Eq) else
                    left != right if isinstance(operator, ast.NotEq) else
                    left < right if isinstance(operator, ast.Lt) else
                    left <= right if isinstance(operator, ast.LtE) else
                    left > right if isinstance(operator, ast.Gt) else
                    left >= right
                )
                if not matches:
                    return 0.0
                left = right
            return 1.0
        if isinstance(node, ast.Call):
            name = node.func.id
            args = [visit(arg) for arg in node.args]
            if name == "coalesce":
                return next((value for value in args if value is not None), None)
            if name == "nullif":
                if len(args) != 2:
                    raise RulebookError("nullif requires exactly two arguments")
                return None if args[0] == args[1] else args[0]
            if name == "ifelse":
                if len(args) != 3:
                    raise RulebookError("ifelse requires exactly three arguments")
                return args[1] if args[0] else args[2]
            present = [value for value in args if value is not None]
            if not present:
                return None
            if name == "min":
                return min(present)
            if name == "max":
                return max(present)
            if name == "abs":
                if len(present) != 1:
                    raise RulebookError("abs requires exactly one argument")
                return abs(present[0])
            if name == "round":
                if not 1 <= len(present) <= 2:
                    raise RulebookError("round requires one or two arguments")
                return float(round(present[0], int(present[1]) if len(present) == 2 else 0))
        raise RulebookError(f"Unsupported formula element: {type(node).__name__}")

    result = visit(tree)
    return result if result is None or math.isfinite(result) else None


def formula_names(expression: str, label: str = "formula") -> frozenset[str]:
    """Return component names referenced by a validated safe expression."""
    tree = _validate_expression(expression, label)
    function_names = {
        node.func.id
        for node in ast.walk(tree)
        if isinstance(node, ast.Call) and isinstance(node.func, ast.Name)
    }
    return frozenset(
        node.id
        for node in ast.walk(tree)
        if isinstance(node, ast.Name) and node.id not in function_names
    )


def ensure_rulebook(home: Path) -> Path:
    config_dir = home / "config"
    default = config_dir / "default_rules.toml"
    target = config_dir / "wfm_rules.toml"
    if not default.exists():
        raise RulebookError(f"Default business rulebook is missing: {default}")
    if not target.exists():
        config_dir.mkdir(parents=True, exist_ok=True)
        shutil.copy2(default, target)
    return target


def _required_table(raw: dict[str, Any], key: str) -> dict[str, Any]:
    value = raw.get(key)
    if not isinstance(value, dict):
        raise RulebookError(f"Missing [{key}] table")
    return value


def load_rulebook(home: Path, file: Path | None = None) -> Rulebook:
    file = (file or ensure_rulebook(home)).resolve()
    try:
        content = file.read_bytes()
        raw = tomllib.loads(content.decode("utf-8"))
    except (OSError, UnicodeDecodeError, tomllib.TOMLDecodeError) as exc:
        raise RulebookError(f"Cannot read business rulebook {file}: {exc}") from exc
    meta = _required_table(raw, "rulebook")
    absence = _required_table(raw, "absence")
    service = _required_table(raw, "service")
    pcs = _required_table(raw, "pcs")
    try:
        effective_from = date.fromisoformat(str(meta["effective_from"]))
        version = str(meta["version"]).strip()
        standard_day_hours = float(absence["standard_day_hours"])
        target_seconds = int(service["target_seconds"])
    except (KeyError, TypeError, ValueError) as exc:
        raise RulebookError(f"Rulebook metadata contains an invalid or missing value: {exc}") from exc
    if not version:
        raise RulebookError("rulebook.version cannot be blank")
    if not 0 < standard_day_hours <= 24:
        raise RulebookError("absence.standard_day_hours must be greater than 0 and at most 24")
    for key in ("late_tolerance_minutes", "status_gap_tolerance_minutes", "verint_match_tolerance_minutes"):
        if not 0 <= int(absence.get(key, 5)) <= 120:
            raise RulebookError(f"absence.{key} must be between 0 and 120")
    if not 1 <= target_seconds <= 600:
        raise RulebookError("service.target_seconds must be between 1 and 600")
    activity_rules: list[ActivityRule] = []
    seen_categories: set[str] = set()
    valid_matches = {"exact", "contains", "exact_or_suffix"}
    for index, item in enumerate(raw.get("activity_rules", []), 1):
        if not isinstance(item, dict):
            raise RulebookError(f"activity_rules item {index} must be a table")
        name = str(item.get("name", "")).strip()
        category = str(item.get("category", "")).strip().upper()
        patterns = tuple(str(value).strip().upper() for value in item.get("patterns", []) if str(value).strip())
        match = str(item.get("match", "contains"))
        if not name or not category or not patterns:
            raise RulebookError(f"activity_rules item {index} requires name, category and patterns")
        if category in seen_categories:
            raise RulebookError(f"Duplicate activity category {category!r}")
        if match not in valid_matches:
            raise RulebookError(f"Activity rule {name!r} has unsupported match {match!r}")
        seen_categories.add(category)
        activity_rules.append(ActivityRule(
            name=name, category=category, patterns=patterns, match=match,
            planned=bool(item.get("planned", False)), working=bool(item.get("working", False)),
            absence=bool(item.get("absence", False)), vacation=bool(item.get("vacation", False)),
            unpaid=bool(item.get("unpaid", False)), shrinkage=bool(item.get("shrinkage", False)),
        ))
    if not activity_rules:
        raise RulebookError("At least one [[activity_rules]] entry is required")
    required_categories = {"LUNCH", "BREAK", "LATE", "EARLY_LEAVE", "NO_SHOW", "NO_ACTIVITY"}
    missing_categories = sorted(required_categories - seen_categories)
    if missing_categories:
        raise RulebookError(f"Missing engine-required activity categories: {', '.join(missing_categories)}")

    scored_questions = tuple(int(value) for value in pcs.get("scored_questions", []))
    comment_questions = tuple(int(value) for value in pcs.get("comment_questions", []))
    pcs_minimum = float(pcs.get("minimum_score", 1))
    pcs_maximum = float(pcs.get("maximum_score", 5))
    pcs_top = float(pcs.get("top_box_minimum", 4))
    pcs_low = float(pcs.get("low_score_maximum", 2))
    pcs_primary_question = int(pcs.get("primary_score_question", 1))
    pcs_participation_question = int(pcs.get("participation_question", 1))
    pcs_participation_status = str(pcs.get("participation_status", "1")).strip()
    pcs_allowed_scores = tuple(float(value) for value in pcs.get("allowed_scores", [1, 2, 3, 4, 5]))
    pcs_negative_maximum = float(pcs.get("negative_score_maximum", 3))
    if not scored_questions or any(value not in range(1, 11) for value in scored_questions):
        raise RulebookError("pcs.scored_questions must contain question numbers from 1 to 10")
    if any(value not in range(1, 11) for value in comment_questions):
        raise RulebookError("pcs.comment_questions must contain question numbers from 1 to 10")
    if not pcs_minimum < pcs_maximum or not pcs_minimum <= pcs_low <= pcs_top <= pcs_maximum:
        raise RulebookError("PCS score thresholds must satisfy minimum <= low <= top <= maximum")
    if pcs_primary_question not in range(1, 11):
        raise RulebookError("pcs.primary_score_question must be from 1 to 10")
    if pcs_participation_question not in range(1, 11):
        raise RulebookError("pcs.participation_question must be from 1 to 10")
    if not pcs_participation_status:
        raise RulebookError("pcs.participation_status cannot be blank")
    if not pcs_allowed_scores or len(set(pcs_allowed_scores)) != len(pcs_allowed_scores):
        raise RulebookError("pcs.allowed_scores must contain unique numeric values")
    if any(value < pcs_minimum or value > pcs_maximum for value in pcs_allowed_scores):
        raise RulebookError("pcs.allowed_scores must stay inside the configured score range")
    if not pcs_minimum <= pcs_negative_maximum < pcs_maximum:
        raise RulebookError("pcs.negative_score_maximum must be inside the configured score range")

    return Rulebook(
        file=file, version=version, effective_from=effective_from,
        description=str(meta.get("description", "")), sha256=hashlib.sha256(content).hexdigest(),
        standard_day_hours=standard_day_hours,
        late_tolerance_minutes=int(absence.get("late_tolerance_minutes", 5)),
        status_gap_tolerance_minutes=int(absence.get("status_gap_tolerance_minutes", 5)),
        verint_match_tolerance_minutes=int(absence.get("verint_match_tolerance_minutes", 5)),
        spell_gap_days=int(absence.get("spell_gap_days", 1)),
        cap_event_to_schedule=bool(absence.get("cap_event_to_schedule", True)),
        unmapped_activity_is_error=bool(absence.get("unmapped_activity_is_error", True)),
        target_seconds=target_seconds, activity_rules=tuple(activity_rules),
        pcs_scored_questions=scored_questions, pcs_comment_questions=comment_questions,
        pcs_survey_mode=str(pcs.get("survey_mode", "2")),
        pcs_primary_score_question=pcs_primary_question,
        pcs_participation_question=pcs_participation_question,
        pcs_participation_status=pcs_participation_status,
        pcs_allowed_scores=pcs_allowed_scores,
        pcs_negative_score_maximum=pcs_negative_maximum,
        pcs_minimum_score=pcs_minimum,
        pcs_maximum_score=pcs_maximum, pcs_top_box_minimum=pcs_top,
        pcs_low_score_maximum=pcs_low,
    )


def validate_rulebook(rulebook: Rulebook) -> list[str]:
    """Return a concise validation record for non-metric domain rules."""
    return [
        f"Rulebook {rulebook.version} is valid.",
        f"SHA-256: {rulebook.sha256}",
        f"{len(rulebook.activity_rules)} activity rules; KPI arithmetic is in metric_catalog.toml.",
    ]
