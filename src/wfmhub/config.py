"""Configuration loading with portable, root-relative path resolution."""

from __future__ import annotations

import shutil
import tomllib
from dataclasses import dataclass
from datetime import date
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
    sources: dict[str, str]
    period_start: date | None
    period_end: date | None
    rules: Rules
    modules: dict[str, bool]
    report_limits: dict[str, int]

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
        database=_portable_path(home, str(paths.get("database", "database/wfm.duckdb"))),
        output=_portable_path(home, str(paths.get("output", "output"))),
        logs=_portable_path(home, str(paths.get("logs", "logs"))),
        backups=_portable_path(home, str(paths.get("backups", "backups"))),
        input=_portable_path(home, str(paths.get("input", "input"))),
        sources={str(k): str(v) for k, v in raw.get("sources", {}).items()},
        period_start=_date_or_none(period.get("start"), "period.start"),
        period_end=_date_or_none(period.get("end"), "period.end"),
        rules=Rules(
            tolerance_minutes=int(rules.get("tolerance_minutes", 5)),
            break_minutes=int(rules.get("break_minutes", 30)),
            lunch_minutes=int(rules.get("lunch_minutes", 45)),
            minimum_status_coverage=float(rules.get("minimum_status_coverage", 0.80)),
            rta_stale_minutes=int(rules.get("rta_stale_minutes", 30)),
        ),
        modules={str(k): bool(v) for k, v in raw.get("modules", {}).items()},
        report_limits={str(k): int(v) for k, v in raw.get("report", {}).items()},
    )
    if cfg.period_start and cfg.period_end and cfg.period_start > cfg.period_end:
        raise ConfigError("period.start cannot be after period.end")
    if not 0 <= cfg.rules.minimum_status_coverage <= 1:
        raise ConfigError("rules.minimum_status_coverage must be between 0 and 1")
    for path in (cfg.database.parent, cfg.output, cfg.logs, cfg.backups, cfg.input):
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
