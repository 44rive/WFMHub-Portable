"""Validated presentation contracts, deliberately separate from KPI logic."""

from __future__ import annotations

import hashlib
import shutil
import tomllib
from dataclasses import dataclass
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
