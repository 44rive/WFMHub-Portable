"""Configuration loading with portable, root-relative path resolution."""

from __future__ import annotations

import shutil
import tomllib
from dataclasses import dataclass, replace
from datetime import date, datetime
from pathlib import Path
from typing import Any


class ConfigError(RuntimeError):
    pass


@dataclass(frozen=True)
class Rules:
    tolerance_minutes: int
    break_minutes: int
    lunch_minutes: int
    minimum_status_coverage: float
    rta_stale_minutes: int


@dataclass(frozen=True)
class PCSSettings:
    scored_questions: tuple[int, ...]
    comment_questions: tuple[int, ...]
    survey_mode: str
    primary_score_question: int
    participation_question: int
    participation_status: str
    allowed_scores: tuple[float, ...]
    negative_score_maximum: float
    minimum_score: float
    maximum_score: float
    top_box_minimum: float
    low_score_maximum: float


@dataclass(frozen=True)
class Config:
    home: Path
    file: Path
    timezone: str
    source_root: Path
    database: Path
    output: Path
    logs: Path
    backups: Path
    input: Path
    custom: Path
    business_rules: Path
    metric_catalog: Path
    analytics_rules: Path
    report_catalog: Path
    queue_mapping: Path
    service_profiles: Path
    sources: dict[str, str]
    period_start: date | None
    period_end: date | None
    rules: Rules
    pcs: PCSSettings
    modules: dict[str, bool]
    report_limits: dict[str, int]
    report_packs: dict[str, str]

    def source_path(self, key: str) -> Path:
        value = self.sources.get(key)
        if not value:
            raise ConfigError(f"Missing sources.{key} in {self.file}")
        return (self.source_root / value).resolve()


def _portable_path(home: Path, value: str) -> Path:
    path = Path(value).expanduser()
    return path.resolve() if path.is_absolute() else (home / path).resolve()


def _date_or_none(value: Any, label: str) -> date | None:
    text = str(value or "").strip()
    if not text:
        return None
    try:
        return date.fromisoformat(text)
    except ValueError as exc:
        raise ConfigError(f"{label} must be YYYY-MM-DD, received {text!r}") from exc


def ensure_user_config(home: Path) -> Path:
    config_dir = home / "config"
    target = config_dir / "wfmhub.toml"
    source = config_dir / "default.toml"
    if not source.exists():
        raise ConfigError(f"Default configuration is missing: {source}")
    config_dir.mkdir(parents=True, exist_ok=True)
    if not target.exists():
        shutil.copy2(source, target)
    else:
        # v0.1 used DuckDB. Preserve that database file and move the standard
        # configuration to a new SQLite file without changing any source path
        # or user-selected reporting period.
        text = target.read_text(encoding="utf-8")
        legacy = 'database = "database/wfm.duckdb"'
        if legacy in text:
            stamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
            backup = config_dir / f"wfmhub_pre_sqlite_{stamp}.toml"
            shutil.copy2(target, backup)
            target.write_text(
                text.replace(legacy, 'database = "database/wfm.sqlite3"', 1),
                encoding="utf-8",
            )
            print("WFMHub upgraded the standard database setting to database/wfm.sqlite3.")
            print(f"Previous config backup: {backup}")
            print(f"Old DuckDB remains untouched: {home / 'database' / 'wfm.duckdb'}")
    return target


def load_config(home: Path, config_file: Path | None = None) -> Config:
    home = home.resolve()
    file = (config_file or ensure_user_config(home)).resolve()
    try:
        with file.open("rb") as handle:
            raw = tomllib.load(handle)
    except (OSError, tomllib.TOMLDecodeError) as exc:
        raise ConfigError(f"Cannot read {file}: {exc}") from exc

    paths = raw.get("paths", {})
    rules = raw.get("rules", {})
    period = raw.get("period", {})
    cfg = Config(
        home=home,
        file=file,
        timezone=str(raw.get("app", {}).get("timezone", "Europe/Berlin")),
        source_root=_portable_path(home, str(paths.get("source_root", "extracts"))),
        database=_portable_path(home, str(paths.get("database", "database/wfm.sqlite3"))),
        output=_portable_path(home, str(paths.get("output", "output"))),
        logs=_portable_path(home, str(paths.get("logs", "logs"))),
        backups=_portable_path(home, str(paths.get("backups", "backups"))),
        input=_portable_path(home, str(paths.get("input", "input"))),
        custom=_portable_path(home, str(paths.get("custom", "custom"))),
        business_rules=_portable_path(home, str(paths.get("business_rules", "config/wfm_rules.toml"))),
        metric_catalog=_portable_path(home, str(paths.get("metric_catalog", "config/metric_catalog.toml"))),
        analytics_rules=_portable_path(home, str(paths.get("analytics_rules", "config/analytics_rules.toml"))),
        report_catalog=_portable_path(home, str(paths.get("report_catalog", "config/report_catalog.toml"))),
        queue_mapping=_portable_path(home, str(paths.get("queue_mapping", "config/queue_mapping.csv"))),
        service_profiles=_portable_path(home, str(paths.get("service_profiles", "config/service_profiles.toml"))),
        sources={
            "call_folder": "Storm/Call by Call",
            "apde_folder": "Storm/APDE Standard KPIs Inbound Calls",
            **{str(k): str(v) for k, v in raw.get("sources", {}).items()},
        },
        period_start=_date_or_none(period.get("start"), "period.start"),
        period_end=_date_or_none(period.get("end"), "period.end"),
        rules=Rules(
            tolerance_minutes=int(rules.get("tolerance_minutes", 5)),
            break_minutes=int(rules.get("break_minutes", 30)),
            lunch_minutes=int(rules.get("lunch_minutes", 45)),
            minimum_status_coverage=float(rules.get("minimum_status_coverage", 0.80)),
            rta_stale_minutes=int(rules.get("rta_stale_minutes", 30)),
        ),
        pcs=PCSSettings(
            scored_questions=tuple(int(value) for value in raw.get("pcs", {}).get("scored_questions", [1, 2])),
            comment_questions=tuple(int(value) for value in raw.get("pcs", {}).get("comment_questions", [3])),
            survey_mode=str(raw.get("pcs", {}).get("survey_mode", "2")),
            primary_score_question=int(raw.get("pcs", {}).get("primary_score_question", 1)),
            participation_question=int(raw.get("pcs", {}).get("participation_question", 1)),
            participation_status=str(raw.get("pcs", {}).get("participation_status", "1")),
            allowed_scores=tuple(float(value) for value in raw.get("pcs", {}).get("allowed_scores", [1, 2, 3, 4, 5])),
            negative_score_maximum=float(raw.get("pcs", {}).get("negative_score_maximum", 3)),
            minimum_score=float(raw.get("pcs", {}).get("minimum_score", 1)),
            maximum_score=float(raw.get("pcs", {}).get("maximum_score", 5)),
            top_box_minimum=float(raw.get("pcs", {}).get("top_box_minimum", 4)),
            low_score_maximum=float(raw.get("pcs", {}).get("low_score_maximum", 2)),
        ),
        modules={"pcs": True, "absence": True, **{str(k): bool(v) for k, v in raw.get("modules", {}).items()}},
        report_limits={str(k): int(v) for k, v in raw.get("report", {}).items()},
        report_packs={
            "operations": "operations",
            "intraday": "intraday",
            "quality_pcs": "quality_pcs",
            "pcs": "pcs",
            "bonus": "bonus",
            "service": "service",
            "staffing": "staffing",
            "attendance": "attendance",
            "absence": "absence",
            "corrections": "corrections",
            "scorecard": "scorecard",
            **{str(k): str(v) for k, v in raw.get("report_packs", {}).items()},
        },
    )
    if cfg.period_start and cfg.period_end and cfg.period_start > cfg.period_end:
        raise ConfigError("period.start cannot be after period.end")
    if not 0 <= cfg.rules.minimum_status_coverage <= 1:
        raise ConfigError("rules.minimum_status_coverage must be between 0 and 1")
    if not cfg.pcs.scored_questions or any(question not in range(1, 11) for question in cfg.pcs.scored_questions):
        raise ConfigError("pcs.scored_questions must contain question numbers from 1 to 10")
    if any(question not in range(1, 11) for question in cfg.pcs.comment_questions):
        raise ConfigError("pcs.comment_questions must contain question numbers from 1 to 10")
    if cfg.pcs.primary_score_question not in range(1, 11):
        raise ConfigError("pcs.primary_score_question must be from 1 to 10")
    if cfg.pcs.participation_question not in range(1, 11):
        raise ConfigError("pcs.participation_question must be from 1 to 10")
    if not cfg.pcs.participation_status.strip():
        raise ConfigError("pcs.participation_status cannot be blank")
    if not cfg.pcs.allowed_scores or len(set(cfg.pcs.allowed_scores)) != len(cfg.pcs.allowed_scores):
        raise ConfigError("pcs.allowed_scores must contain unique numeric values")
    if not cfg.pcs.minimum_score < cfg.pcs.maximum_score:
        raise ConfigError("pcs.minimum_score must be lower than pcs.maximum_score")
    if not cfg.pcs.minimum_score <= cfg.pcs.low_score_maximum <= cfg.pcs.maximum_score:
        raise ConfigError("pcs.low_score_maximum must be inside the configured score range")
    if not cfg.pcs.minimum_score <= cfg.pcs.top_box_minimum <= cfg.pcs.maximum_score:
        raise ConfigError("pcs.top_box_minimum must be inside the configured score range")
    if any(value < cfg.pcs.minimum_score or value > cfg.pcs.maximum_score for value in cfg.pcs.allowed_scores):
        raise ConfigError("pcs.allowed_scores must stay inside the configured score range")
    if not cfg.pcs.minimum_score <= cfg.pcs.negative_score_maximum < cfg.pcs.maximum_score:
        raise ConfigError("pcs.negative_score_maximum must be inside the configured score range")
    if cfg.database.suffix.lower() == ".duckdb":
        raise ConfigError(
            "This corporate-compatible release uses SQLite. Change paths.database "
            "in config\\wfmhub.toml to database/wfm.sqlite3; the old DuckDB file is preserved."
        )
    from .analytics import ensure_analytics_rules, load_analytics_rules
    from .mapping import ensure_queue_mapping, load_queue_mapping
    from .metrics import ensure_metric_catalog, load_metric_catalog
    from .report_specs import ensure_report_catalog, load_report_catalog
    from .rules import ensure_rulebook, load_rulebook, validate_rulebook
    from .service_profiles import ensure_service_profiles, load_service_profiles, validate_service_profiles

    ensure_rulebook(home)
    ensure_metric_catalog(home, cfg.metric_catalog)
    ensure_analytics_rules(home, cfg.analytics_rules)
    ensure_report_catalog(home, cfg.report_catalog)
    ensure_queue_mapping(home, cfg.queue_mapping)
    ensure_service_profiles(home, cfg.service_profiles)
    load_queue_mapping(cfg.queue_mapping)
    metric_catalog = load_metric_catalog(home, cfg.metric_catalog)
    validate_service_profiles(load_service_profiles(home, cfg.service_profiles), metric_catalog)
    load_analytics_rules(home, cfg.analytics_rules)
    load_report_catalog(home, cfg.report_catalog)
    business = load_rulebook(home, cfg.business_rules)
    validate_rulebook(business)
    cfg = replace(cfg, pcs=PCSSettings(
        scored_questions=business.pcs_scored_questions,
        comment_questions=business.pcs_comment_questions,
        survey_mode=business.pcs_survey_mode,
        primary_score_question=business.pcs_primary_score_question,
        participation_question=business.pcs_participation_question,
        participation_status=business.pcs_participation_status,
        allowed_scores=business.pcs_allowed_scores,
        negative_score_maximum=business.pcs_negative_score_maximum,
        minimum_score=business.pcs_minimum_score,
        maximum_score=business.pcs_maximum_score,
        top_box_minimum=business.pcs_top_box_minimum,
        low_score_maximum=business.pcs_low_score_maximum,
    ))
    for path in (cfg.database.parent, cfg.output, cfg.logs, cfg.backups, cfg.input, cfg.custom):
        path.mkdir(parents=True, exist_ok=True)
    return cfg


def write_source_root(config_file: Path, source_root: Path) -> None:
    """Update only paths.source_root while preserving the friendly TOML file."""
    text = config_file.read_text(encoding="utf-8")
    safe = source_root.resolve().as_posix().replace('"', '\\"')
    lines = text.splitlines()
    in_paths = False
    changed = False
    for index, line in enumerate(lines):
        stripped = line.strip()
        if stripped.startswith("["):
            in_paths = stripped == "[paths]"
        elif in_paths and stripped.startswith("source_root"):
            lines[index] = f'source_root = "{safe}"'
            changed = True
            break
    if not changed:
        raise ConfigError(f"Could not find paths.source_root in {config_file}")
    config_file.write_text("\n".join(lines) + "\n", encoding="utf-8")
