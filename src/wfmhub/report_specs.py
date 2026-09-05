"""Validated presentation contracts, deliberately separate from KPI logic."""

from __future__ import annotations

import hashlib
import shutil
import tomllib
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path


class ReportCatalogError(RuntimeError):
    pass


@dataclass(frozen=True)
class ReportSpec:
    key: str
    title: str
    purpose: str
    finding_domains: tuple[str, ...]
    sheets: tuple[str, ...]

    def includes(self, sheet: str) -> bool:
        return sheet in self.sheets


@dataclass(frozen=True)
class ReportCatalog:
    file: Path
    version: str
    description: str
    sha256: str
    packs: dict[str, ReportSpec]

    def pack(self, key: str) -> ReportSpec:
        try:
            return self.packs[key]
        except KeyError as exc:
            raise ReportCatalogError(
                f"Unknown report contract {key!r}. Available: {', '.join(self.packs)}"
            ) from exc


def ensure_report_catalog(home: Path, target: Path | None = None) -> Path:
    default = home / "config" / "default_reports.toml"
    target = target or home / "config" / "report_catalog.toml"
    if not default.exists():
        raise ReportCatalogError(f"Default report catalog is missing: {default}")
    if not target.exists():
        target.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(default, target)
    else:
        # Presentation contracts are executable workbook schemas. Migrate the
        # known contracts safely when the Service Flashes layout changes;
        # keep a recoverable copy instead of silently accepting a guaranteed
        # runtime mismatch.
        try:
            current_meta = tomllib.loads(target.read_text(encoding="utf-8")).get("catalog", {})
            default_meta = tomllib.loads(default.read_text(encoding="utf-8")).get("catalog", {})
        except (OSError, UnicodeDecodeError, tomllib.TOMLDecodeError):
            current_meta, default_meta = {}, {}
        current_version = str(current_meta.get("version", ""))
        default_version = str(default_meta.get("version", ""))
        if (current_version, default_version) in {
            ("2026.09.3", "2026.09.4"),
            ("2026.09.4", "2026.09.5"),
            ("2026.09.5", "2026.09.6"),
            ("2026.09.3", "2026.09.7"),
            ("2026.09.4", "2026.09.7"),
            ("2026.09.5", "2026.09.7"),
            ("2026.09.6", "2026.09.7"),
            ("2026.09.3", "2026.09.11"),
            ("2026.09.4", "2026.09.11"),
            ("2026.09.5", "2026.09.11"),
            ("2026.09.6", "2026.09.11"),
            ("2026.09.7", "2026.09.11"),
            ("2026.09.8", "2026.09.11"),
            ("2026.09.9", "2026.09.11"),
            ("2026.09.10", "2026.09.11"),
        }:
            stamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
            safe_version = default_version.replace(".", "_")
            shutil.copy2(target, target.with_name(f"{target.stem}_pre_{safe_version}_{stamp}{target.suffix}"))
            shutil.copy2(default, target)
    return target


def load_report_catalog(home: Path, file: Path | None = None) -> ReportCatalog:
    file = ensure_report_catalog(home, file).resolve()
    try:
        content = file.read_bytes()
        raw = tomllib.loads(content.decode("utf-8"))
    except (OSError, UnicodeDecodeError, tomllib.TOMLDecodeError) as exc:
        raise ReportCatalogError(f"Cannot read report catalog {file}: {exc}") from exc
    meta = raw.get("catalog")
    if not isinstance(meta, dict) or not str(meta.get("version", "")).strip():
        raise ReportCatalogError("Missing report catalog.version")
    packs: dict[str, ReportSpec] = {}
    for index, item in enumerate(raw.get("packs", []), 1):
        if not isinstance(item, dict):
            raise ReportCatalogError(f"packs item {index} must be a table")
        key = str(item.get("key", "")).strip()
        title = str(item.get("title", "")).strip()
        sheets = tuple(str(value).strip() for value in item.get("sheets", []) if str(value).strip())
        if not key or not title or not sheets:
            raise ReportCatalogError(f"packs item {index} requires key, title and sheets")
        if key in packs:
            raise ReportCatalogError(f"Duplicate report contract {key!r}")
        if len(set(sheets)) != len(sheets):
            raise ReportCatalogError(f"Report contract {key!r} contains duplicate sheet names")
        packs[key] = ReportSpec(
            key=key, title=title, purpose=str(item.get("purpose", "")),
            finding_domains=tuple(str(value) for value in item.get("finding_domains", [])),
            sheets=sheets,
        )
    if not packs:
        raise ReportCatalogError("At least one [[packs]] report contract is required")
    return ReportCatalog(
        file=file, version=str(meta["version"]), description=str(meta.get("description", "")),
        sha256=hashlib.sha256(content).hexdigest(), packs=packs,
    )


def validate_report_catalog(catalog: ReportCatalog, required: tuple[str, ...] = ()) -> list[str]:
    missing = sorted(set(required) - set(catalog.packs))
    if missing:
        raise ReportCatalogError(f"Missing report contracts: {', '.join(missing)}")
    return [
        f"Report catalog {catalog.version} is valid.",
        f"SHA-256: {catalog.sha256}",
        f"{len(catalog.packs)} focused workbook contracts.",
    ]
