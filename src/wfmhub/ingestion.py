"""Read-only source discovery, parsing and idempotent raw ingestion."""

from __future__ import annotations

import csv
import hashlib
import re
import unicodedata
from dataclasses import dataclass, field
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any, Callable

from openpyxl import load_workbook

from .config import Config
from .database import DatabaseConnection
from .utils import (
    classify_assignment,
    classify_event,
    classify_status,
    duration_seconds,
    file_sha256,
    normalize_header,
    normalize_id,
    parse_date,
    parse_datetime,
    parse_time,
    parse_verint_interval,
)


class SourceSchemaError(RuntimeError):
    pass


@dataclass
class ParseResult:
    tables: dict[str, list[dict[str, Any]]] = field(default_factory=dict)
    rejected: list[str] = field(default_factory=list)
    scoped_out: int = 0

    @property
    def row_count(self) -> int:
        primary = [name for name in self.tables if name != "raw.schedule_event"]
        return sum(len(self.tables[name]) for name in primary)


@dataclass(frozen=True)
class SourceCandidate:
    family: str
    path: Path


@dataclass
class IngestSummary:
    loaded: int = 0
    skipped: int = 0
    failed: int = 0
    rows: int = 0
    scoped_out: int = 0
    errors: list[str] = field(default_factory=list)


def _normalize_agent_name(value: Any) -> str | None:
    """Return a conservative, accent-insensitive key for exact name fallback."""
    text = str(value or "").strip()
    if not text:
        return None
    decomposed = unicodedata.normalize("NFKD", text).casefold()
    words = re.findall(r"[a-z0-9]+", "".join(char for char in decomposed if not unicodedata.combining(char)))
    return " ".join(words) or None


@dataclass(frozen=True)
class AgentScope:
    """Authoritative agent scope derived from the active FTE roster."""

    agent_ids: frozenset[str]
    unique_names: dict[str, str]
    fingerprint: str

    def resolve(self, agent_id: Any, agent_name: Any) -> str | None:
        normalized_id = normalize_id(agent_id)
        if normalized_id in self.agent_ids:
            return normalized_id
        name_key = _normalize_agent_name(agent_name)
        roster_id = self.unique_names.get(name_key) if name_key else None
        if roster_id is None:
            return None
        # Preserve a populated operational source ID (especially Verint Data
        # Source IDs) while using the unique roster name only as the scope gate.
        return normalized_id or roster_id


AGENT_SCOPED_FAMILIES = {"schedule", "lilo", "agent_status"}
AGENT_SCOPE_POLICY_VERSION = "v1-id-or-unique-name-preserve-source-id"


LILO_FILE_RE = re.compile(r"^AP-Historical-Report---Agent-Login (\d{4}-\d{2}-\d{2})\.csv$", re.I)


TABLE_COLUMNS: dict[str, list[str]] = {
    "raw.fte_agent": ["source_file_id", "source_row", "agent_id", "employment_status", "agent_name", "team_leader", "ops_manager", "lob", "market", "language", "location", "city", "fte", "end_date"],
    "raw.schedule_shift": ["source_file_id", "source_row", "schedule_date", "agent_id_raw", "agent_id", "agent_name", "scheduling_period", "shift_assignment", "assignment", "assignment_type", "scheduled_start", "scheduled_end", "shift_events", "parse_ok"],
    "raw.schedule_event": ["source_file_id", "source_row", "event_index", "schedule_date", "agent_id", "agent_name", "activity", "activity_type", "event_start", "event_end", "parse_ok"],
    "raw.lilo": ["source_file_id", "source_row", "extract_date", "agent_id", "agent_name", "first_login", "raw_last_logout", "last_logout", "overnight_adjusted"],
    "raw.agent_status": ["source_file_id", "source_row", "serial_number", "extract_date", "agent_id", "agent_name", "status", "actual_category", "status_start", "status_end", "duration_seconds", "queue"],
    "raw.forecast_interval": ["source_file_id", "source_row", "queue_name", "business_date", "interval_time", "interval_minutes", "interval_start", "volume_forecast", "abandons_forecast", "sl_forecast", "sl_required", "aht_forecast_seconds", "headcount_forecast", "net_staffing_forecast", "fte_forecast", "fte_required"],
    "raw.queue_actual": ["source_file_id", "source_row", "source_system", "business_date", "interval_time", "interval_start", "hour_start", "language", "queue_id", "queue", "business_partner", "lob", "offered", "answered", "abandoned", "short_calls", "answered_15s", "answered_20s", "answered_30s", "asa_seconds", "aht_seconds"],
}


FTE_HEADER_ALIASES: dict[str, tuple[str, ...]] = {
    "agent_id": ("CLIENTID", "AGENTID", "VERINTID", "DATASOURCEID", "DATASOURCEIDS"),
    "agent_name": ("NAME", "AGENTNAME", "AGENT", "EMPLOYEENAME", "FULLNAME"),
    "employment_status": ("STATUS", "EMPLOYMENTSTATUS", "AGENTSTATUS"),
    "team_leader": ("TEAMLEADER", "TEAMLEADERNAME", "TL"),
    "ops_manager": ("OPSMANAGER", "OPERATIONSMANAGER", "OPERATIONSMANAGERNAME"),
    "lob": ("LOB", "LINEOFBUSINESS"),
    "market": ("MARKET",),
    "language": ("LANGUAGE",),
    "location": ("LOCATION", "SITE"),
    "city": ("CITY",),
    "fte": ("FTE", "FTECOUNT"),
    "end_date": ("ENDDATEIFLEAVER", "ENDDATE", "LEAVERENDDATE"),
}

FTE_EXACT_ROSTER_TITLES = {
    "AGENT", "AGENTS", "AGENTLIST", "AGENTROSTER", "FTE", "FTECOUNT",
    "FTEAGENT", "FTEAGENTS", "HEADCOUNT", "ROSTER",
}


def _clean(value: Any) -> str | None:
    text = str(value).strip() if value is not None else ""
    return text or None


def _number(value: Any) -> float | None:
    if value is None or str(value).strip() == "":
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def _extract_date(name: str, pattern: re.Pattern[str], family: str) -> datetime.date:
    match = pattern.match(name)
    if not match:
        raise SourceSchemaError(f"{family} filename must contain its YYYY-MM-DD date: {name}")
    return datetime.strptime(match.group(1), "%Y-%m-%d").date()


def discover_sources(config: Config) -> list[SourceCandidate]:
    candidates: list[SourceCandidate] = []
    single = config.source_path("fte_file")
    if single.is_file() and not single.name.startswith("~$"):
        candidates.append(SourceCandidate("fte", single))
    specs = (
        ("schedule", "schedule_folder", "*.txt"),
        ("lilo", "lilo_folder", "*.csv"),
        ("agent_status", "agent_status_folder", "*.csv"),
        ("forecast", "forecast_folder", "*.txt"),
        ("apbe", "apbe_folder", "*.xlsx"),
        ("apfr", "apfr_folder", "*.xlsx"),
    )
    for family, key, pattern in specs:
        if family in {"agent_status"} and not config.modules.get("agent_status", True):
            continue
        if family in {"forecast"} and not config.modules.get("forecast", True):
            continue
        if family in {"apbe", "apfr"} and not config.modules.get("intraday", True):
            continue
        folder = config.source_path(key)
        if not folder.is_dir():
            continue
        for path in sorted(folder.glob(pattern)):
            if path.is_file() and not path.name.startswith("~$"):
                candidates.append(SourceCandidate(family, path.resolve()))
    return candidates


def _fte_column_indexes(
    headers: tuple[Any, ...] | list[Any],
) -> tuple[dict[str, int], dict[str, list[int]]]:
    normalized: dict[str, list[int]] = {}
    for index, value in enumerate(headers):
        if str(value or "").strip():
            normalized.setdefault(normalize_header(value), []).append(index)
    indexes: dict[str, int] = {}
    duplicates: dict[str, list[int]] = {}
    for field, aliases in FTE_HEADER_ALIASES.items():
        matches = sorted({index for alias in aliases for index in normalized.get(alias, [])})
        if len(matches) == 1:
            indexes[field] = matches[0]
        elif len(matches) > 1:
            duplicates[field] = matches
    return indexes, duplicates


def _find_fte_table(workbook, path: Path) -> tuple[str, int, Any, dict[str, int]]:
    def confidence_tier(sheet_title: str, recognized: set[str]) -> int:
        normalized_title = normalize_header(sheet_title)
        if normalized_title in FTE_EXACT_ROSTER_TITLES:
            return 3
        title_words = set(re.findall(r"[A-Z0-9]+", sheet_title.upper()))
        if title_words & {"ROSTER", "HEADCOUNT"}:
            return 2
        org_fields = {"lob", "market", "language", "location", "city"}
        has_leadership = bool(recognized & {"team_leader", "ops_manager"})
        if "employment_status" in recognized and has_leadership and len(recognized & org_fields) >= 2:
            return 1
        return 0

    candidates: list[tuple[int, Any, int, dict[str, int]]] = []
    ambiguous: list[str] = []
    for sheet in workbook.worksheets:
        rows = sheet.iter_rows(values_only=True)
        for header_row, values in enumerate(rows, 1):
            if header_row > 100:
                break
            indexes, duplicates = _fte_column_indexes(values)
            identity_present = "agent_id" in indexes or "agent_id" in duplicates
            name_present = "agent_name" in indexes or "agent_name" in duplicates
            recognized = set(indexes) | set(duplicates)
            tier = confidence_tier(sheet.title, recognized)
            if identity_present and name_present and duplicates and tier:
                columns = ", ".join(sorted(duplicates))
                ambiguous.append(f"{sheet.title!r} row {header_row} has multiple aliases for: {columns}")
                break
            if {"agent_id", "agent_name"} <= indexes.keys() and tier:
                has_data = any(
                    (
                        indexes["agent_id"] < len(row) and row[indexes["agent_id"]] not in (None, "")
                    )
                    and (
                        indexes["agent_name"] < len(row) and str(row[indexes["agent_name"]] or "").strip()
                    )
                    for row in rows
                )
                if has_data:
                    candidates.append((tier, sheet, header_row, indexes))
                break
    if ambiguous:
        raise SourceSchemaError(
            f"FTE workbook has ambiguous roster headers. {'; '.join(ambiguous)}. "
            "Keep only one ID and one name column in the authoritative roster table."
        )
    if candidates:
        highest_tier = max(candidate[0] for candidate in candidates)
        best = [candidate for candidate in candidates if candidate[0] == highest_tier]
        if len(best) > 1:
            locations = ", ".join(f"{item[1].title!r} row {item[2]}" for item in best)
            raise SourceSchemaError(
                f"FTE workbook has multiple equally likely agent tables: {locations}. "
                "Keep one authoritative roster table or rename its sheet to Agent/Roster/FTE."
            )
        _, sheet, header_row, indexes = best[0]
        rows = sheet.iter_rows(min_row=header_row + 1, values_only=True)
        return sheet.title, header_row, rows, indexes
    searched = ", ".join(workbook.sheetnames) or "(no worksheets)"
    raise SourceSchemaError(
        f"FTE agent table was not found in {path.name}. Looked in sheets: {searched}. "
        "Expected an Agent/FTE/Roster/Headcount sheet with Client ID/Agent ID and Name/Agent Name."
    )


def parse_fte(path: Path, file_id: str) -> ParseResult:
    workbook = load_workbook(path, read_only=True, data_only=True, keep_links=False)
    try:
        sheet_name, header_row, rows, index = _find_fte_table(workbook, path)
        output: list[dict[str, Any]] = []
        for source_row, values in enumerate(rows, header_row + 1):
            get = lambda field: values[index[field]] if field in index and index[field] < len(values) else None
            if get("agent_id") is None and not _clean(get("agent_name")):
                continue
            output.append({
                "source_file_id": file_id,
                "source_row": source_row,
                "agent_id": normalize_id(get("agent_id")),
                "employment_status": _clean(get("employment_status")),
                "agent_name": _clean(get("agent_name")),
                "team_leader": _clean(get("team_leader")),
                "ops_manager": _clean(get("ops_manager")),
                "lob": _clean(get("lob")),
                "market": _clean(get("market")),
                "language": _clean(get("language")),
                "location": _clean(get("location")),
                "city": _clean(get("city")),
                "fte": _number(get("fte")),
                "end_date": parse_date(get("end_date")),
            })
        if not output:
            raise SourceSchemaError(
                f"FTE agent table on sheet {sheet_name!r} has headers but no populated agent rows: {path.name}"
            )
        return ParseResult({"raw.fte_agent": output})
    finally:
        workbook.close()


def parse_schedule(path: Path, file_id: str, scope: AgentScope | None = None) -> ParseResult:
    shifts: list[dict[str, Any]] = []
    events: list[dict[str, Any]] = []
    rejected: list[str] = []
    scoped_out = 0
    marker = None
    with path.open("r", encoding="cp1252", newline="") as handle:
        reader = csv.DictReader(handle, delimiter="\t")
        required = {"Name", "Data Source IDs", "Scheduling Period", "Shift Assignment", "Shift Events"}
        missing = sorted(required - set(reader.fieldnames or []))
        if missing:
            raise SourceSchemaError(f"Schedule missing columns: {', '.join(missing)}")
        for source_row, row in enumerate(reader, 2):
            name = _clean(row.get("Name"))
            raw_id = normalize_id(row.get("Data Source IDs"), reject_placeholders=False)
            agent_id = normalize_id(row.get("Data Source IDs"))
            if name and raw_id is None:
                possible = parse_date(name)
                if possible:
                    marker = possible
                    continue
            if not name and raw_id is None:
                continue
            if marker is None:
                raise SourceSchemaError(f"Agent row before date marker at line {source_row}")
            if scope is not None:
                resolved_id = scope.resolve(agent_id, name)
                if resolved_id is None:
                    scoped_out += 1
                    continue
                agent_id = resolved_id
            raw_assignment = _clean(row.get("Shift Assignment"))
            assignment, start, end = parse_verint_interval(raw_assignment)
            is_off = (raw_assignment or "").strip().upper() == "OFF"
            schedule_date = start.date() if start else marker
            parse_ok = is_off or bool(start and end and end > start)
            shifts.append({
                "source_file_id": file_id, "source_row": source_row,
                "schedule_date": schedule_date, "agent_id_raw": raw_id,
                "agent_id": agent_id, "agent_name": name,
                "scheduling_period": _clean(row.get("Scheduling Period")),
                "shift_assignment": raw_assignment, "assignment": assignment or raw_assignment,
                "assignment_type": classify_assignment(raw_assignment, assignment),
                "scheduled_start": start, "scheduled_end": end,
                "shift_events": _clean(row.get("Shift Events")), "parse_ok": parse_ok,
            })
            if not parse_ok:
                rejected.append(f"line {source_row}: shift interval could not be parsed")
            for event_index, part in enumerate(str(row.get("Shift Events") or "").split(";"), 1):
                if not part.strip():
                    continue
                activity, event_start, event_end = parse_verint_interval(part)
                event_ok = bool(event_start and event_end and event_end > event_start)
                events.append({
                    "source_file_id": file_id, "source_row": source_row, "event_index": event_index,
                    "schedule_date": schedule_date, "agent_id": agent_id, "agent_name": name,
                    "activity": activity, "activity_type": classify_event(activity),
                    "event_start": event_start, "event_end": event_end, "parse_ok": event_ok,
                })
                if not event_ok:
                    rejected.append(f"line {source_row} event {event_index}: interval could not be parsed")
    return ParseResult({"raw.schedule_shift": shifts, "raw.schedule_event": events}, rejected, scoped_out)


def parse_lilo(path: Path, file_id: str, scope: AgentScope | None = None) -> ParseResult:
    extract_date = _extract_date(path.name, LILO_FILE_RE, "LILO")
    output: list[dict[str, Any]] = []
    scoped_out = 0
    with path.open("r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle)
        required = {"[Agent]", "[Agent ID]", "[First Log-on Time]", "[Last Log-off Time]"}
        missing = sorted(required - set(reader.fieldnames or []))
        if missing:
            raise SourceSchemaError(f"LILO missing columns: {', '.join(missing)}")
        for source_row, row in enumerate(reader, 2):
            agent_id = normalize_id(row.get("[Agent ID]"))
            if scope is None and not agent_id:
                continue
            if scope is not None:
                resolved_id = scope.resolve(agent_id, row.get("[Agent]"))
                if resolved_id is None:
                    scoped_out += 1
                    continue
                agent_id = resolved_id
            first = parse_datetime(row.get("[First Log-on Time]"))
            raw_last = parse_datetime(row.get("[Last Log-off Time]"))
            last = raw_last
            adjusted = False
            if first and last and last < first:
                last += timedelta(days=1)
                adjusted = True
            output.append({
                "source_file_id": file_id, "source_row": source_row,
                "extract_date": extract_date, "agent_id": agent_id,
                "agent_name": _clean(row.get("[Agent]")), "first_login": first,
                "raw_last_logout": raw_last, "last_logout": last,
                "overnight_adjusted": adjusted,
            })
    return ParseResult({"raw.lilo": output}, scoped_out=scoped_out)


def _validate_lilo_header(path: Path) -> datetime.date:
    extract_date = _extract_date(path.name, LILO_FILE_RE, "LILO")
    with path.open("r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.reader(handle)
        headers = set(next(reader, []))
    required = {"[Agent]", "[Agent ID]", "[First Log-on Time]", "[Last Log-off Time]"}
    missing = sorted(required - headers)
    if missing:
        raise SourceSchemaError(f"LILO missing columns: {', '.join(missing)}")
    return extract_date


def _insert_lilo_direct(
    conn: DatabaseConnection,
    path: Path,
    file_id: str,
    extract_date,
    scope: AgentScope,
) -> tuple[int, int]:
    """Stream large LILO CSVs into bounded SQLite inserts."""
    count = 0
    scoped_out = 0
    batch: list[dict[str, Any]] = []
    with path.open("r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle)
        for source_row, row in enumerate(reader, 2):
            agent_id = normalize_id(row.get("[Agent ID]"))
            resolved_id = scope.resolve(agent_id, row.get("[Agent]"))
            if resolved_id is None:
                scoped_out += 1
                continue
            agent_id = resolved_id
            first = parse_datetime(row.get("[First Log-on Time]"))
            raw_last = parse_datetime(row.get("[Last Log-off Time]"))
            last = raw_last
            adjusted = False
            if first and last and last < first:
                last += timedelta(days=1)
                adjusted = True
            batch.append({
                "source_file_id": file_id,
                "source_row": source_row,
                "extract_date": extract_date,
                "agent_id": agent_id,
                "agent_name": _clean(row.get("[Agent]")),
                "first_login": first,
                "raw_last_logout": raw_last,
                "last_logout": last,
                "overnight_adjusted": adjusted,
            })
            if len(batch) >= 5000:
                _insert_rows(conn, "raw.lilo", batch)
                count += len(batch)
                batch.clear()
    if batch:
        _insert_rows(conn, "raw.lilo", batch)
        count += len(batch)
    return count, scoped_out


def parse_agent_status(path: Path, file_id: str, scope: AgentScope | None = None) -> ParseResult:
    output: list[dict[str, Any]] = []
    rejected: list[str] = []
    scoped_out = 0
    with path.open("r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle)
        required = {"[Serial Number]", "[Status]", "[Status Start Date and Time]", "[Agent]", "[Agent ID]", "[Status Duration]", "[Queue]"}
        missing = sorted(required - set(reader.fieldnames or []))
        if missing:
            raise SourceSchemaError(f"Agent Status missing columns: {', '.join(missing)}")
        for source_row, row in enumerate(reader, 2):
            agent_id = normalize_id(row.get("[Agent ID]"))
            if scope is not None:
                resolved_id = scope.resolve(agent_id, row.get("[Agent]"))
                if resolved_id is None:
                    scoped_out += 1
                    continue
                agent_id = resolved_id
            start = parse_datetime(row.get("[Status Start Date and Time]"))
            seconds = duration_seconds(row.get("[Status Duration]"))
            end = start + timedelta(seconds=seconds) if start is not None and seconds is not None else None
            serial = _clean(row.get("[Serial Number]"))
            if not serial:
                serial = hashlib.sha256(repr(sorted(row.items())).encode("utf-8")).hexdigest()
            if not start or not end or end <= start:
                rejected.append(f"line {source_row}: invalid status interval")
            output.append({
                "source_file_id": file_id, "source_row": source_row,
                # Agent Status exports can contain one day or a date range. The
                # row timestamp is the authority; the filename is only a label.
                "serial_number": serial, "extract_date": start.date() if start else None,
                "agent_id": agent_id,
                "agent_name": _clean(row.get("[Agent]")), "status": _clean(row.get("[Status]")),
                "actual_category": classify_status(_clean(row.get("[Status]"))),
                "status_start": start, "status_end": end, "duration_seconds": seconds,
                "queue": _clean(row.get("[Queue]")),
            })
    return ParseResult({"raw.agent_status": output}, rejected, scoped_out)


FORECAST_NAMES = {
    "volume_forecast": "Volume (Absolute For)",
    "abandons_forecast": "Abandons (Absolute For)",
    "sl_forecast": "Service Level (Absolute For)",
    "sl_required": "Service Level (Absolute Req)",
    "aht_forecast_seconds": "Activity Handling Time (Absolute For)",
    "headcount_forecast": "Headcount Staffing (Absolute For)",
    "net_staffing_forecast": "Net Staffing (Absolute For)",
    "fte_forecast": "Full Time Equivalents (Absolute For)",
    "fte_required": "Full Time Equivalents (Absolute Req)",
}


def parse_forecast(path: Path, file_id: str) -> ParseResult:
    with path.open("r", encoding="cp1252", newline="") as handle:
        lines = handle.readlines()
    header_index = next((i for i, line in enumerate(lines) if line.lstrip().startswith("Queue Name\tDate\tTime")), None)
    if header_index is None:
        raise SourceSchemaError("Forecast Queue Name / Date / Time header was not found")
    reader = csv.DictReader(lines[header_index:], delimiter="\t")
    required = {"Queue Name", "Date", "Time", "Time Interval", *FORECAST_NAMES.values()}
    missing = sorted(required - set(reader.fieldnames or []))
    if missing:
        raise SourceSchemaError(f"Forecast missing columns: {', '.join(missing)}")
    output: list[dict[str, Any]] = []
    rejected: list[str] = []
    for source_row, row in enumerate(reader, header_index + 2):
        business_date = parse_date(row.get("Date"))
        interval_time = parse_time(row.get("Time"))
        interval_seconds = duration_seconds(row.get("Time Interval"))
        if business_date is None or interval_time is None:
            rejected.append(f"line {source_row}: invalid forecast date/time")
            continue
        result = {
            "source_file_id": file_id, "source_row": source_row,
            "queue_name": _clean(row.get("Queue Name")), "business_date": business_date,
            "interval_time": interval_time, "interval_minutes": int(interval_seconds / 60) if interval_seconds is not None else None,
            "interval_start": datetime.combine(business_date, interval_time),
        }
        for target, source in FORECAST_NAMES.items():
            value = _number(row.get(source))
            result[target] = value / 100 if target in {"sl_forecast", "sl_required"} and value is not None else value
        output.append(result)
    return ParseResult({"raw.forecast_interval": output}, rejected)


def _find_excel_header(workbook, path: Path) -> tuple[Any, list[str], Any]:
    for sheet in workbook.worksheets:
        rows = sheet.iter_rows(values_only=True)
        for source_row, values in enumerate(rows, 1):
            headers = [str(value or "").strip() for value in values]
            normalized = {normalize_header(value) for value in headers}
            if "DATE" in normalized and "15MINUTEPERIODSOFDAY" in normalized:
                return sheet, headers, (source_row, rows)
    raise SourceSchemaError(f"Could not find the queue-actual header row in {path.name}")


def parse_queue_actual(path: Path, file_id: str, source_system: str) -> ParseResult:
    workbook = load_workbook(path, read_only=True, data_only=True, keep_links=False)
    try:
        sheet, headers, state = _find_excel_header(workbook, path)
        header_row, rows = state
        normalized = {normalize_header(name): index for index, name in enumerate(headers) if name}

        def require(*names: str) -> int:
            for name in names:
                key = normalize_header(name)
                if key in normalized:
                    return normalized[key]
            raise SourceSchemaError(f"{source_system} missing required column: {' / '.join(names)}")

        date_i = require("Date")
        time_i = require("15 Minute Periods of Day")
        partner_i = require("BusinessPartnerID")
        lob_i = require("LineOfBusiness")
        if source_system == "APBE":
            offered_i = require("Offered_calls (w/o short calls)")
            answered_i = require("Answered_Calls")
            abandoned_i = require("Abandoned_Calls (w/o short calls)")
            ans15_i = require("Answered_Calls <= 15s")
            ans20_i = require("Answered_Calls <= 20s")
            ans30_i = require("Answered_Calls <= 30s")
        else:
            offered_i = require("APPELS ENTRANTS")
            answered_i = require("APPELS RÉP", "APPELS REP")
            abandoned_i = require("APPELS ABAN")
            ans15_i = require("APPELS RÉP <= 15s", "APPELS REP <= 15s")
            ans20_i = require("APPELS RÉP <= 20s", "APPELS REP <= 20s")
            ans30_i = require("APPELS RÉP <= 30s", "APPELS REP <= 30s")
        short_i = require("Short_calls < 5s")
        asa_i = require("Average_Speed_of_Answer")
        talk_i = require("Average_Talk_Time")
        hold_i = require("Average_Hold_Time")
        wrap_i = require("Average Total Wrap Time")
        queue_i = normalized.get("QUEUE")
        queue_id_i = normalized.get("QUEUEID")
        language_i = normalized.get("LANGUAGE")
        output: list[dict[str, Any]] = []
        rejected: list[str] = []
        for source_row, values in enumerate(rows, header_row + 1):
            get = lambda index: values[index] if index is not None and index < len(values) else None
            business_date = parse_date(get(date_i))
            if business_date is None:
                if str(get(date_i) or "").strip().upper() not in {"", "TOTAL"}:
                    rejected.append(f"line {source_row}: invalid date")
                continue
            interval_time = parse_time(get(time_i))
            if interval_time is None:
                rejected.append(f"line {source_row}: invalid interval")
                continue
            talk = duration_seconds(get(talk_i))
            hold = duration_seconds(get(hold_i))
            wrap = duration_seconds(get(wrap_i))
            parts = [value for value in (talk, hold, wrap) if value is not None]
            partner = _clean(get(partner_i))
            queue = _clean(get(queue_i)) if queue_i is not None else partner
            interval_start = datetime.combine(business_date, interval_time)
            output.append({
                "source_file_id": file_id, "source_row": source_row, "source_system": source_system,
                "business_date": business_date, "interval_time": interval_time,
                "interval_start": interval_start, "hour_start": interval_start.replace(minute=0, second=0, microsecond=0),
                "language": _clean(get(language_i)) if language_i is not None else ("FR" if source_system == "APFR" else None),
                "queue_id": normalize_id(get(queue_id_i), reject_placeholders=False) if queue_id_i is not None else None,
                "queue": queue, "business_partner": partner, "lob": _clean(get(lob_i)),
                "offered": _number(get(offered_i)), "answered": _number(get(answered_i)),
                "abandoned": _number(get(abandoned_i)), "short_calls": _number(get(short_i)),
                "answered_15s": _number(get(ans15_i)), "answered_20s": _number(get(ans20_i)),
                "answered_30s": _number(get(ans30_i)), "asa_seconds": duration_seconds(get(asa_i)),
                "aht_seconds": sum(parts) if parts else None,
            })
        return ParseResult({"raw.queue_actual": output}, rejected)
    finally:
        workbook.close()


PARSERS: dict[str, Callable[[Path, str, AgentScope | None], ParseResult]] = {
    "fte": lambda path, file_id, scope: parse_fte(path, file_id),
    "schedule": parse_schedule,
    "lilo": parse_lilo,
    "agent_status": parse_agent_status,
    "forecast": lambda path, file_id, scope: parse_forecast(path, file_id),
    "apbe": lambda path, file_id, scope: parse_queue_actual(path, file_id, "APBE"),
    "apfr": lambda path, file_id, scope: parse_queue_actual(path, file_id, "APFR"),
}


def _load_agent_scope(conn: DatabaseConnection) -> AgentScope:
    rows = conn.execute(
        """SELECT r.agent_id, r.agent_name
           FROM raw.fte_agent r
           JOIN meta.source_file f ON f.file_id=r.source_file_id
           WHERE f.active=true AND f.status='SUCCESS'"""
    ).fetchall()
    ids = {normalize_id(agent_id) for agent_id, _ in rows}
    ids.discard(None)
    if not ids:
        raise SourceSchemaError(
            "Agent scope is empty. Load a valid FTE roster before agent-level extracts; "
            "no worldwide rows were admitted."
        )
    names: dict[str, set[str]] = {}
    for agent_id, agent_name in rows:
        normalized_id = normalize_id(agent_id)
        name_key = _normalize_agent_name(agent_name)
        if normalized_id and name_key:
            names.setdefault(name_key, set()).add(normalized_id)
    unique_names = {name: next(iter(matches)) for name, matches in names.items() if len(matches) == 1}
    payload = "\n".join([
        f"policy:{AGENT_SCOPE_POLICY_VERSION}",
        *(f"id:{value}" for value in sorted(ids)),
        *(f"name:{key}={value}" for key, value in sorted(unique_names.items())),
    ])
    fingerprint = hashlib.sha256(payload.encode("utf-8")).hexdigest()
    return AgentScope(frozenset(ids), unique_names, fingerprint)


def _insert_rows(conn: DatabaseConnection, table: str, rows: list[dict[str, Any]]) -> None:
    if not rows:
        return
    columns = TABLE_COLUMNS[table]
    row_placeholders = "(" + ", ".join("?" for _ in columns) + ")"
    # Bounded multi-row statements keep SQLite transaction overhead low while
    # staying below its host-parameter limit.
    batch_size = max(1, min(500, conn.max_variable_number // len(columns)))
    for offset in range(0, len(rows), batch_size):
        batch = rows[offset : offset + batch_size]
        sql = f"INSERT INTO {table} ({', '.join(columns)}) VALUES " + ", ".join(row_placeholders for _ in batch)
        params = [row.get(column) for row in batch for column in columns]
        conn.execute(sql, params)


def ingest_all(conn: DatabaseConnection, config: Config) -> IngestSummary:
    summary = IngestSummary()
    for candidate in discover_sources(config):
        path = candidate.path
        sha256 = file_sha256(path)
        path_text = str(path)
        try:
            scope = _load_agent_scope(conn) if candidate.family in AGENT_SCOPED_FAMILIES else None
        except Exception as exc:
            summary.failed += 1
            summary.errors.append(f"{candidate.family}: {path.name}: {exc}")
            continue
        scope_fingerprint = scope.fingerprint if scope is not None else ""
        existing = conn.execute(
            """SELECT file_id, active, row_count, scoped_out_count FROM meta.source_file
               WHERE source_family=? AND source_path=? AND sha256=?
                 AND coalesce(scope_fingerprint, '')=? AND status='SUCCESS'""",
            [candidate.family, path_text, sha256, scope_fingerprint],
        ).fetchone()
        if existing:
            existing_id, active, existing_rows, existing_scoped_out = existing
            if active:
                summary.skipped += 1
                summary.scoped_out += existing_scoped_out or 0
                continue
            # The path changed A -> B -> A. Its original raw rows are still
            # immutable in the hub, so safely reactivate them without parsing
            # or duplicating the same content fingerprint.
            conn.execute("BEGIN TRANSACTION")
            try:
                conn.execute(
                    "UPDATE meta.source_file SET active=false WHERE source_family=? AND source_path=? AND active=true",
                    [candidate.family, path_text],
                )
                conn.execute(
                    "UPDATE meta.source_file SET active=true, loaded_at=? WHERE file_id=?",
                    [datetime.now(), existing_id],
                )
                conn.execute("COMMIT")
            except Exception:
                if conn.in_transaction:
                    conn.execute("ROLLBACK")
                raise
            summary.loaded += 1
            summary.rows += existing_rows or 0
            summary.scoped_out += existing_scoped_out or 0
            continue
        stat = path.stat()
        file_id = hashlib.sha256(
            f"{candidate.family}|{path_text}|{sha256}|{scope_fingerprint}".encode("utf-8")
        ).hexdigest()
        discovered_at = datetime.now()
        modified_at = datetime.fromtimestamp(stat.st_mtime)
        try:
            result = None if candidate.family == "lilo" else PARSERS[candidate.family](path, file_id, scope)
            conn.execute("BEGIN TRANSACTION")
            try:
                conn.execute(
                    "UPDATE meta.source_file SET active=false WHERE source_family=? AND source_path=? AND active=true",
                    [candidate.family, path_text],
                )
                if candidate.family == "lilo":
                    extract_date = _validate_lilo_header(path)
                    row_count, scoped_out_count = _insert_lilo_direct(conn, path, file_id, extract_date, scope)
                    rejected_count = 0
                else:
                    row_count = result.row_count
                    rejected_count = len(result.rejected)
                    scoped_out_count = result.scoped_out
                    for table, rows in result.tables.items():
                        _insert_rows(conn, table, rows)
                conn.execute(
                    """INSERT INTO meta.source_file(
                           file_id, source_family, source_path, file_name, sha256, size_bytes,
                           modified_at, discovered_at, loaded_at, active, status, row_count,
                           rejected_count, error_message, scope_fingerprint, scoped_out_count
                       ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, true, 'SUCCESS', ?, ?, NULL, ?, ?)
                       ON CONFLICT(file_id) DO UPDATE SET
                           source_family=excluded.source_family, source_path=excluded.source_path,
                           file_name=excluded.file_name, sha256=excluded.sha256,
                           size_bytes=excluded.size_bytes, modified_at=excluded.modified_at,
                           discovered_at=excluded.discovered_at, loaded_at=excluded.loaded_at,
                           active=true, status='SUCCESS', row_count=excluded.row_count,
                           rejected_count=excluded.rejected_count, error_message=NULL,
                           scope_fingerprint=excluded.scope_fingerprint,
                           scoped_out_count=excluded.scoped_out_count""",
                    [file_id, candidate.family, path_text, path.name, sha256, stat.st_size, modified_at, discovered_at, datetime.now(), row_count, rejected_count, scope_fingerprint, scoped_out_count],
                )
                conn.execute("COMMIT")
            except Exception:
                if conn.in_transaction:
                    conn.execute("ROLLBACK")
                raise
            summary.loaded += 1
            summary.rows += row_count
            summary.scoped_out += scoped_out_count
        except Exception as exc:
            conn.execute(
                """INSERT INTO meta.source_file(
                       file_id, source_family, source_path, file_name, sha256, size_bytes,
                       modified_at, discovered_at, loaded_at, active, status, row_count,
                       rejected_count, error_message, scope_fingerprint, scoped_out_count
                   ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, false, 'ERROR', 0, 0, ?, ?, 0)
                   ON CONFLICT(file_id) DO UPDATE SET
                       modified_at=excluded.modified_at, discovered_at=excluded.discovered_at,
                       loaded_at=excluded.loaded_at, active=false, status='ERROR',
                       row_count=0, rejected_count=0, error_message=excluded.error_message,
                       scope_fingerprint=excluded.scope_fingerprint, scoped_out_count=0""",
                [file_id, candidate.family, path_text, path.name, sha256, stat.st_size, modified_at, discovered_at, datetime.now(), str(exc)[:4000], scope_fingerprint],
            )
            summary.failed += 1
            summary.errors.append(f"{candidate.family}: {path.name}: {exc}")
    return summary
