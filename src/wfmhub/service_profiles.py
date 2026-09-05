"""Effective-dated LOB service definitions used by reports, never by extracts."""

from __future__ import annotations

import hashlib
import shutil
import tomllib
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from .metrics import MetricCatalog


class ServiceProfileError(RuntimeError):
    pass


@dataclass(frozen=True)
class ServiceGroup:
    label: str
    queue_contains: tuple[str, ...]

    def matches(self, queue: str | None) -> bool:
        normalized = str(queue or "").upper()
        return any(value.upper() in normalized for value in self.queue_contains)


@dataclass(frozen=True)
class ServiceProfile:
    profile_id: str
    label: str
    service_scopes: tuple[str, ...]
    staffing_lobs: tuple[str, ...]
    source_systems: tuple[str, ...]
    service_level_metric: str
    availability_metric: str
    aht_metric: str
    effective_from: date
    effective_to: date | None
    groups: tuple[ServiceGroup, ...]
    flash_sheet: str
    flash_layout: str
    flash_source_systems: tuple[str, ...]
    operating_start_hour: int
    operating_end_hour: int
    display_order: int

    def active_on(self, value: date) -> bool:
        return self.effective_from <= value and (self.effective_to is None or value <= self.effective_to)

    def group_for(self, queue: str | None) -> str:
        for group in self.groups:
            if group.matches(queue):
                return group.label
        return "Other"

    def staffing_pairs(self) -> tuple[tuple[str, str], ...]:
        """Return explicit service-scope to roster-LOB planning links."""

        if len(self.service_scopes) == len(self.staffing_lobs):
            return tuple(zip(self.service_scopes, self.staffing_lobs))
        if len(self.staffing_lobs) == 1:
            return tuple((scope, self.staffing_lobs[0]) for scope in self.service_scopes)
        raise ServiceProfileError(
            f"Service profile {self.profile_id!r} must define one staffing LOB "
            "or one staffing LOB per service scope"
        )


@dataclass(frozen=True)
class ServiceProfileCatalog:
    file: Path
    version: str
    default_profile: str
    sha256: str
    profiles: tuple[ServiceProfile, ...]

    def select(self, profile_id: str | None, on_date: date) -> ServiceProfile:
        wanted = profile_id or self.default_profile
        matches = [profile for profile in self.profiles if profile.profile_id == wanted and profile.active_on(on_date)]
        if len(matches) != 1:
            raise ServiceProfileError(
                f"Service profile {wanted!r} must have exactly one active definition on {on_date}; found {len(matches)}"
            )
        return matches[0]


def ensure_service_profiles(home: Path, target: Path | None = None) -> Path:
    source = home / "config" / "default_service_profiles.toml"
    if not source.exists():
        packaged = Path(__file__).resolve().parents[2] / "config" / "default_service_profiles.toml"
        if packaged.exists():
            source = packaged
    target = (target or home / "config" / "service_profiles.toml").resolve()
    if not source.exists():
        raise ServiceProfileError(f"Default service profile catalog is missing: {source}")
    if not target.exists():
        target.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(source, target)
    else:
        try:
            current = tomllib.loads(target.read_text(encoding="utf-8")).get("catalog", {})
            default = tomllib.loads(source.read_text(encoding="utf-8")).get("catalog", {})
        except (OSError, UnicodeDecodeError, tomllib.TOMLDecodeError):
            current, default = {}, {}
        if (
            str(current.get("version", "")) == "2026.09.3"
            and str(default.get("version", "")) == "2026.09.4"
        ):
            stamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
            shutil.copy2(
                target,
                target.with_name(f"{target.stem}_pre_flash_layout_{stamp}{target.suffix}"),
            )
            shutil.copy2(source, target)
    return target


def load_service_profiles(home: Path, target: Path | None = None) -> ServiceProfileCatalog:
    file = ensure_service_profiles(home, target)
    content = file.read_bytes()
    try:
        raw = tomllib.loads(content.decode("utf-8"))
    except (UnicodeDecodeError, tomllib.TOMLDecodeError) as exc:
        raise ServiceProfileError(f"Cannot parse {file}: {exc}") from exc
    catalog = raw.get("catalog", {})
    version = str(catalog.get("version", "")).strip()
    default_profile = str(catalog.get("default_profile", "")).strip()
    if not version or not default_profile:
        raise ServiceProfileError("service profile catalog requires version and default_profile")
    profiles: list[ServiceProfile] = []
    for index, item in enumerate(raw.get("profiles", []), 1):
        try:
            profile_id = str(item["id"]).strip()
            effective_from = date.fromisoformat(str(item["effective_from"]))
            effective_to = date.fromisoformat(str(item["effective_to"])) if item.get("effective_to") else None
            scopes = tuple(str(value).strip() for value in item["service_scopes"] if str(value).strip())
            staffing_lobs = tuple(
                str(value).strip()
                for value in item.get("staffing_lobs", item["service_scopes"])
                if str(value).strip()
            )
            systems = tuple(str(value).strip().upper() for value in item["source_systems"] if str(value).strip())
            flash_systems = tuple(
                str(value).strip().upper()
                for value in item.get("flash_source_systems", ["CALL_BY_CALL"])
                if str(value).strip()
            )
            operating_start = int(item.get("operating_start_hour", 0))
            operating_end = int(item.get("operating_end_hour", 23))
            display_order = int(item.get("display_order", 100))
        except (KeyError, TypeError, ValueError) as exc:
            raise ServiceProfileError(f"Invalid service profile item {index}: {exc}") from exc
        if not profile_id or not scopes or not staffing_lobs or not systems or not flash_systems:
            raise ServiceProfileError(
                f"Invalid service profile item {index}: id/scopes/staffing_lobs/systems/flash systems"
            )
        flash_sheet = str(item.get("flash_sheet", item.get("label", profile_id))).strip()
        flash_layout = str(item.get("flash_layout", "standard")).strip().lower()
        if not flash_sheet or len(flash_sheet) > 31:
            raise ServiceProfileError(
                f"Invalid service profile {profile_id!r}: flash_sheet must contain 1-31 characters"
            )
        if flash_layout not in {"standard", "workforce", "oem_split"}:
            raise ServiceProfileError(
                f"Invalid service profile {profile_id!r}: unsupported flash_layout {flash_layout!r}"
            )
        if not 0 <= operating_start <= operating_end <= 23:
            raise ServiceProfileError(
                f"Invalid service profile {profile_id!r}: operating hours must be between 0 and 23"
            )
        groups = tuple(
            ServiceGroup(str(group.get("label", "")).strip(), tuple(str(value) for value in group.get("queue_contains", [])))
            for group in item.get("groups", [])
        )
        if any(not group.label or not group.queue_contains for group in groups):
            raise ServiceProfileError(f"Invalid service group in profile {profile_id!r}")
        profiles.append(ServiceProfile(
            profile_id=profile_id,
            label=str(item.get("label", profile_id)).strip(),
            service_scopes=scopes,
            staffing_lobs=staffing_lobs,
            source_systems=systems,
            service_level_metric=str(item.get("service_level_metric", "service_level")).strip(),
            availability_metric=str(item.get("availability_metric", "service_availability")).strip(),
            aht_metric=str(item.get("aht_metric", "aht_seconds")).strip(),
            effective_from=effective_from,
            effective_to=effective_to,
            groups=groups,
            flash_sheet=flash_sheet,
            flash_layout=flash_layout,
            flash_source_systems=flash_systems,
            operating_start_hour=operating_start,
            operating_end_hour=operating_end,
            display_order=display_order,
        ))
    if not profiles or default_profile not in {profile.profile_id for profile in profiles}:
        raise ServiceProfileError("default_profile must identify at least one profile")
    return ServiceProfileCatalog(file, version, default_profile, hashlib.sha256(content).hexdigest(), tuple(profiles))


def validate_service_profiles(
    catalog: ServiceProfileCatalog,
    metric_catalog: MetricCatalog | None = None,
) -> list[str]:
    flash_sheets: set[str] = set()
    for left_index, left in enumerate(catalog.profiles):
        for right in catalog.profiles[left_index + 1:]:
            if left.profile_id != right.profile_id:
                continue
            if left.effective_from <= (right.effective_to or date.max) and right.effective_from <= (left.effective_to or date.max):
                raise ServiceProfileError(f"Overlapping effective dates for service profile {left.profile_id!r}")
    if metric_catalog is not None:
        known = set(metric_catalog.metric_ids())
        for profile in catalog.profiles:
            selected = {
                profile.service_level_metric,
                profile.availability_metric,
                profile.aht_metric,
            }
            missing = sorted(selected - known)
            if missing:
                raise ServiceProfileError(
                    f"Service profile {profile.profile_id!r} selects unknown metric(s): {', '.join(missing)}"
                )
    for profile in catalog.profiles:
        profile.staffing_pairs()
        key = profile.flash_sheet.casefold()
        if key in flash_sheets:
            raise ServiceProfileError(
                f"Duplicate flash sheet name {profile.flash_sheet!r}"
            )
        flash_sheets.add(key)
    return [
        f"Service profiles {catalog.version} are valid.",
        f"SHA-256: {catalog.sha256}",
        f"{len(catalog.profiles)} effective-dated profile definition(s).",
    ]
