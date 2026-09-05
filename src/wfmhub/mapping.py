"""Editable, audited mappings between source queues/files and service scopes."""

from __future__ import annotations

import csv
import hashlib
import re
import shutil
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable


class QueueMappingError(RuntimeError):
    pass


@dataclass(frozen=True)
class MappingResult:
    service_scope: str
    comparison_scope: str
    designation: str | None
    status: str


@dataclass(frozen=True)
class QueueMapping:
    file: Path
    sha256: str
    queue_rows: dict[tuple[str, str], MappingResult]
    forecast_rows: tuple[tuple[str, MappingResult], ...]

    def map_actual(
        self,
        source_system: str | None,
        queue: str | None,
        business_partner: str | None,
        lob: str | None,
    ) -> MappingResult:
        for value in (queue, business_partner):
            key = _key(value)
            for system_key in (_key(source_system), "STORM", "ANY"):
                if key and (system_key, key) in self.queue_rows:
                    return self.queue_rows[(system_key, key)]
        if str(lob or "").strip():
            scope = str(lob).strip()
            return MappingResult(scope, scope, None, "FALLBACK_LOB")
        return MappingResult("UNMAPPED", "UNMAPPED", None, "UNMAPPED")

    def map_forecast(self, file_name: str, raw_queue: str | None) -> MappingResult:
        key = _key(Path(file_name).stem)
        for prefix, result in self.forecast_rows:
            # Reviewed exports can be named FORD_FR_... or
            # Forecast_FORD_FR_.... The configured token still identifies the
            # service when it occurs in the normalized filename.
            if key.startswith(prefix) or prefix in key:
                return result
        if raw_queue and _key(raw_queue) not in {"COMBINEDALLMEDIA", "ALL"}:
            scope = str(raw_queue).strip()
            return MappingResult(scope, scope, None, "FALLBACK_QUEUE")
        return MappingResult("UNMAPPED", "UNMAPPED", None, "UNMAPPED")

    def comparison_scopes_for(self, service_scopes: Iterable[str]) -> tuple[str, ...]:
        """Return forecast rollups matching a set of detailed actual scopes."""

        wanted = set(service_scopes)
        values = {
            result.comparison_scope
            for result in self.queue_rows.values()
            if result.service_scope in wanted
        }
        return tuple(sorted(values or wanted))


def _key(value: str | None) -> str:
    return re.sub(r"[^A-Z0-9]+", "", str(value or "").upper())


def ensure_queue_mapping(home: Path, target: Path | None = None) -> Path:
    source = home / "config" / "default_queue_mapping.csv"
    if not source.exists():
        packaged = Path(__file__).resolve().parents[2] / "config" / "default_queue_mapping.csv"
        if packaged.exists():
            source = packaged
    target = (target or home / "config" / "queue_mapping.csv").resolve()
    if not source.exists():
        raise QueueMappingError(f"Default queue mapping is missing: {source}")
    if not target.exists():
        target.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(source, target)
    return target


def load_queue_mapping(path: Path) -> QueueMapping:
    try:
        content = path.read_bytes()
    except OSError as exc:
        raise QueueMappingError(f"Cannot read queue mapping {path}: {exc}") from exc
    queue_pending: list[tuple[int, str, str, str, str | None]] = []
    forecast_pending: list[tuple[int, str, str, str | None]] = []
    rollups: dict[str, str] = {}
    with path.open("r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle)
        required = {"mapping_type", "source_system", "source_value", "service_scope", "designation"}
        missing = sorted(required - set(reader.fieldnames or []))
        if missing:
            raise QueueMappingError(f"Queue mapping missing columns: {', '.join(missing)}")
        for row_number, row in enumerate(reader, 2):
            mapping_type = str(row.get("mapping_type") or "").strip().lower()
            source_system = str(row.get("source_system") or "").strip()
            source_value = str(row.get("source_value") or "").strip()
            service_scope = str(row.get("service_scope") or "").strip()
            designation = str(row.get("designation") or "").strip() or None
            if not source_value and not service_scope and not mapping_type:
                continue
            if mapping_type not in {"queue", "forecast_file", "scope_rollup"}:
                raise QueueMappingError(f"Queue mapping line {row_number}: mapping_type must be queue, forecast_file or scope_rollup")
            if not source_system or not source_value or not service_scope:
                raise QueueMappingError(f"Queue mapping line {row_number}: source_system, source_value and service_scope are required")
            key = _key(source_value)
            if not key:
                raise QueueMappingError(f"Queue mapping line {row_number}: source_value has no usable characters")
            if mapping_type == "queue":
                system_key = _key(source_system)
                if any(existing_system == system_key and existing == key for _, existing_system, existing, _, _ in queue_pending):
                    raise QueueMappingError(f"Queue mapping line {row_number}: duplicate queue {source_value!r}")
                queue_pending.append((row_number, system_key, key, service_scope, designation))
            elif mapping_type == "forecast_file":
                if any(existing == key for _, existing, _, _ in forecast_pending):
                    raise QueueMappingError(f"Queue mapping line {row_number}: duplicate forecast prefix {source_value!r}")
                forecast_pending.append((row_number, key, service_scope, designation))
            else:
                scope_key = _key(source_value)
                if scope_key in rollups:
                    raise QueueMappingError(f"Queue mapping line {row_number}: duplicate scope rollup {source_value!r}")
                rollups[scope_key] = service_scope
    make_result = lambda service_scope, designation: MappingResult(
        service_scope, rollups.get(_key(service_scope), service_scope), designation, "MAPPED"
    )
    queue_rows = {
        (system_key, key): make_result(service_scope, designation)
        for _, system_key, key, service_scope, designation in queue_pending
    }
    forecast_rows = [
        (key, make_result(service_scope, designation))
        for _, key, service_scope, designation in forecast_pending
    ]
    forecast_rows.sort(key=lambda item: len(item[0]), reverse=True)
    return QueueMapping(path.resolve(), hashlib.sha256(content).hexdigest(), queue_rows, tuple(forecast_rows))
