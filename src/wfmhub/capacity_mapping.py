"""Governed mapping between Verint Staff Types and published assignments.

This reference is deliberately separate from the Storm queue map.  A Verint
forecast column named ``Queue Name`` contains a workforce Staff Type, not a
call-routing queue, and the two domains must never be joined by label guessing.
"""

from __future__ import annotations

import csv
import hashlib
import re
import shutil
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path


class CapacityMappingError(RuntimeError):
    pass


def _key(value: str | None) -> str:
    return re.sub(r"[^A-Z0-9]+", "", str(value or "").upper())


@dataclass(frozen=True)
class CapacityResult:
    management_lob: str
    planning_group: str
    staff_type: str
    workforce_lob: str
    status: str


@dataclass(frozen=True)
class CapacityMapping:
    file: Path
    sha256: str
    forecast_rows: tuple[tuple[str, str, CapacityResult], ...]
    schedule_rows: dict[tuple[str, str], CapacityResult]
    workforce_groups: dict[str, CapacityResult]
    forecast_groups: tuple[tuple[str, CapacityResult], ...]

    def map_forecast(self, file_name: str, raw_staff_type: str | None) -> CapacityResult:
        file_key = _key(Path(file_name).stem)
        staff_key = _key(raw_staff_type)
        for prefix, configured_staff, result in self.forecast_rows:
            if (file_key.startswith(prefix) or prefix in file_key) and staff_key == configured_staff:
                return result
        for prefix, result in self.forecast_groups:
            if file_key.startswith(prefix) or prefix in file_key:
                return CapacityResult(
                    result.management_lob,
                    result.planning_group,
                    str(raw_staff_type or "(blank Staff Type)").strip(),
                    result.workforce_lob,
                    "UNMAPPED_STAFF_TYPE",
                )
        return CapacityResult(
            "UNMAPPED", "UNMAPPED", str(raw_staff_type or "(blank Staff Type)").strip(),
            "UNMAPPED", "UNMAPPED_FORECAST_FILE",
        )

    def map_schedule(self, workforce_lob: str | None, assignment: str | None) -> CapacityResult:
        lob_key = _key(workforce_lob)
        assignment_key = _key(assignment)
        configured = self.schedule_rows.get((lob_key, assignment_key))
        if configured is not None:
            return configured
        group = self.workforce_groups.get(lob_key)
        if group is not None:
            return CapacityResult(
                group.management_lob,
                group.planning_group,
                str(assignment or "(blank assignment)").strip(),
                str(workforce_lob or group.workforce_lob).strip(),
                "UNMAPPED_ASSIGNMENT",
            )
        return CapacityResult(
            str(workforce_lob or "UNMAPPED").strip(),
            str(workforce_lob or "UNMAPPED").strip(),
            str(assignment or "(blank assignment)").strip(),
            str(workforce_lob or "UNMAPPED").strip(),
            "UNMAPPED_WORKFORCE_LOB",
        )


def ensure_capacity_mapping(home: Path, target: Path | None = None) -> Path:
    source = home / "config" / "default_capacity_mapping.csv"
    if not source.exists():
        packaged = Path(__file__).resolve().parents[2] / "config" / "default_capacity_mapping.csv"
        if packaged.exists():
            source = packaged
    target = (target or home / "config" / "capacity_mapping.csv").resolve()
    if not source.exists():
        raise CapacityMappingError(f"Default capacity mapping is missing: {source}")
    if not target.exists():
        target.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(source, target)
        return target
    if source.resolve() == target:
        return target

    with source.open("r", encoding="utf-8-sig", newline="") as handle:
        shipped_reader = csv.DictReader(handle)
        shipped = list(shipped_reader)
    with target.open("r", encoding="utf-8-sig", newline="") as handle:
        current_reader = csv.DictReader(handle)
        fieldnames = list(current_reader.fieldnames or [])
        current = list(current_reader)
    required = [
        "management_lob", "planning_group", "workforce_lob",
        "forecast_file_prefix", "forecast_staff_type", "schedule_assignment",
        "staff_type",
    ]
    if not all(name in fieldnames for name in required):
        return target

    def identity(row: dict[str, str]) -> tuple[str, str, str, str]:
        return (
            _key(row.get("workforce_lob")), _key(row.get("forecast_file_prefix")),
            _key(row.get("forecast_staff_type")), _key(row.get("schedule_assignment")),
        )

    known = {identity(row) for row in current}
    additions = [row for row in shipped if identity(row) not in known]
    if additions:
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
        shutil.copy2(
            target,
            target.with_name(f"{target.stem}_pre_default_merge_{stamp}{target.suffix}"),
        )
        with target.open("w", encoding="utf-8-sig", newline="") as handle:
            writer = csv.DictWriter(handle, fieldnames=fieldnames)
            writer.writeheader()
            writer.writerows(current)
            writer.writerows(
                {name: row.get(name, "") for name in fieldnames}
                for row in additions
            )
    return target


def load_capacity_mapping(path: Path) -> CapacityMapping:
    try:
        content = path.read_bytes()
    except OSError as exc:
        raise CapacityMappingError(f"Cannot read capacity mapping {path}: {exc}") from exc
    required = {
        "management_lob", "planning_group", "workforce_lob",
        "forecast_file_prefix", "forecast_staff_type", "schedule_assignment",
        "staff_type",
    }
    forecast_rows: list[tuple[str, str, CapacityResult]] = []
    schedule_rows: dict[tuple[str, str], CapacityResult] = {}
    workforce_groups: dict[str, CapacityResult] = {}
    forecast_groups: dict[str, CapacityResult] = {}
    with path.open("r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle)
        missing = sorted(required - set(reader.fieldnames or []))
        if missing:
            raise CapacityMappingError(
                f"Capacity mapping missing columns: {', '.join(missing)}"
            )
        for line, row in enumerate(reader, 2):
            management_lob = str(row.get("management_lob") or "").strip()
            planning_group = str(row.get("planning_group") or "").strip()
            workforce_lob = str(row.get("workforce_lob") or "").strip()
            prefix = _key(row.get("forecast_file_prefix"))
            forecast_staff = _key(row.get("forecast_staff_type"))
            schedule_assignment = _key(row.get("schedule_assignment"))
            staff_type = str(row.get("staff_type") or "").strip()
            if not management_lob or not planning_group or not workforce_lob or not staff_type:
                raise CapacityMappingError(
                    f"Capacity mapping line {line}: management LOB, planning group, "
                    "workforce LOB and Staff Type are required"
                )
            result = CapacityResult(
                management_lob, planning_group, staff_type, workforce_lob, "MAPPED",
            )
            lob_key = _key(workforce_lob)
            existing_group = workforce_groups.get(lob_key)
            if existing_group is not None and (
                existing_group.management_lob, existing_group.planning_group
            ) != (management_lob, planning_group):
                raise CapacityMappingError(
                    f"Capacity mapping line {line}: ambiguous workforce LOB {workforce_lob!r}"
                )
            workforce_groups[lob_key] = result
            if prefix:
                existing_forecast = forecast_groups.get(prefix)
                if existing_forecast is not None and (
                    existing_forecast.management_lob, existing_forecast.planning_group
                ) != (management_lob, planning_group):
                    raise CapacityMappingError(
                        f"Capacity mapping line {line}: ambiguous forecast prefix "
                        f"{row.get('forecast_file_prefix')!r}"
                    )
                forecast_groups[prefix] = result
            if prefix and forecast_staff:
                if any(
                    current_prefix == prefix and current_staff == forecast_staff
                    for current_prefix, current_staff, _ in forecast_rows
                ):
                    raise CapacityMappingError(
                        f"Capacity mapping line {line}: duplicate forecast Staff Type"
                    )
                forecast_rows.append((prefix, forecast_staff, result))
            if schedule_assignment:
                key = (lob_key, schedule_assignment)
                if key in schedule_rows:
                    raise CapacityMappingError(
                        f"Capacity mapping line {line}: duplicate schedule assignment"
                    )
                schedule_rows[key] = result
    forecast_rows.sort(key=lambda item: len(item[0]), reverse=True)
    ordered_groups = tuple(sorted(forecast_groups.items(), key=lambda item: len(item[0]), reverse=True))
    return CapacityMapping(
        path.resolve(), hashlib.sha256(content).hexdigest(), tuple(forecast_rows),
        schedule_rows, workforce_groups, ordered_groups,
    )
