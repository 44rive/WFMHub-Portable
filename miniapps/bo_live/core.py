"""Validated immutable captures, local storage, and configured BO calculations."""

from __future__ import annotations

import csv
import hashlib
import json
import re
import sqlite3
import threading
import uuid
from contextlib import contextmanager
from datetime import datetime, timedelta, timezone
from pathlib import Path
from zoneinfo import ZoneInfo, ZoneInfoNotFoundError

UTC = timezone.utc
MAX_BYTES = 16 * 1024 * 1024
DEFAULT_SETTINGS = {
    "source_timezone": None,
    "groups": ["(TOYOTA) Agendas à Traiter"],
    "scope_confirmed": False,
    "assignment_blank_means_unassigned": False,
    "priority_order": ["Very high", "High", "Standard"],
    "upcoming_minutes": 60,
    "poll_minutes": 5,
    "stale_minutes": 10,
}


class ContractError(ValueError):
    """A capture or setting cannot safely support the requested calculation."""


def utcnow():
    return datetime.now(UTC)


def stamp(value):
    return value.astimezone(UTC).isoformat(timespec="seconds")


def text_field(value, label, maximum=300, required=False):
    if not isinstance(value, str) or len(value) > maximum:
        raise ContractError(f"{label} must be text of at most {maximum} characters.")
    if any(ord(c) < 32 for c in value):
        raise ContractError(f"{label} contains unsupported control characters.")
    if required and not value.strip():
        raise ContractError(f"{label} is required.")
    return value


def integer(value, label, minimum, maximum):
    if type(value) is not int or not minimum <= value <= maximum:
        raise ContractError(f"{label} must be an integer from {minimum} to {maximum}.")
    return value


def aware_time(value, label):
    try:
        result = datetime.fromisoformat(value.replace("Z", "+00:00"))
        if result.tzinfo is None or result.utcoffset() is None:
            raise ValueError("missing offset")
        return result.astimezone(UTC)
    except (ValueError, TypeError, AttributeError):
        raise ContractError(f"{label} needs an ISO timestamp with a UTC offset.") from None


def source_time(value, zone_name):
    """Resolve a full source timestamp, refusing ambiguous/nonexistent wall times."""
    if value is None or value == "":
        return None
    text_field(value, "Scheduled for", 80, required=True)
    if not re.search(r"[T ]\d{2}:\d{2}", value):
        raise ContractError("Scheduled for must include a full date and time, not a date alone.")
    zone = ZoneInfo(zone_name)
    try:
        result = datetime.fromisoformat(value.replace("Z", "+00:00"))
    except ValueError:
        result = None
        for pattern in ("%d/%m/%Y %H:%M:%S", "%d/%m/%Y %H:%M"):
            try:
                result = datetime.strptime(value, pattern)
                break
            except ValueError:
                pass
        if result is None:
            raise ContractError("Scheduled for must be a full ISO or DD/MM/YYYY HH:MM timestamp.")
    if result.tzinfo is not None:
        return stamp(result)
    candidates = set()
    for fold in (0, 1):
        localized = result.replace(tzinfo=zone, fold=fold)
        back = localized.astimezone(UTC).astimezone(zone).replace(tzinfo=None)
        if back == result:
            candidates.add(stamp(localized))
    if len(candidates) != 1:
        raise ContractError("Scheduled for has an ambiguous or nonexistent local time; include its UTC offset.")
    return candidates.pop()


def validate_settings(value):
    if not isinstance(value, dict) or set(value) != set(DEFAULT_SETTINGS):
        raise ContractError("Settings must contain exactly the documented setting fields.")
    result = dict(value)
    zone = value["source_timezone"]
    if zone is not None:
        text_field(zone, "Source time zone", 100, required=True)
        try:
            ZoneInfo(zone)
        except (ZoneInfoNotFoundError, ValueError):
            raise ContractError("Unknown source time zone. Use an IANA zone, for example Europe/Brussels.") from None
    groups = value["groups"]
    if not isinstance(groups, list) or not 1 <= len(groups) <= 50:
        raise ContractError("Configure between 1 and 50 exact TOYOTA task groups.")
    for group in groups:
        text_field(group, "Task group", 200, required=True)
        if not group.startswith("(TOYOTA) ") or group != group.strip() or "..." in group or "…" in group:
            raise ContractError("Use full exact TOYOTA group names, without clipped labels or outer whitespace.")
    if len(set(groups)) != len(groups):
        raise ContractError("Task groups must be unique.")
    for field in ("scope_confirmed", "assignment_blank_means_unassigned"):
        if type(value[field]) is not bool:
            raise ContractError(f"{field} must be true or false.")
    priorities = value["priority_order"]
    if not isinstance(priorities, list) or not 1 <= len(priorities) <= 20:
        raise ContractError("Configure a priority order with 1 to 20 labels.")
    for priority in priorities:
        text_field(priority, "Priority label", 80, required=True)
    if len(set(priorities)) != len(priorities):
        raise ContractError("Priority labels must be unique.")
    integer(value["upcoming_minutes"], "Upcoming window", 1, 1440)
    integer(value["poll_minutes"], "Expected capture interval", 1, 120)
    integer(value["stale_minutes"], "Stale threshold", value["poll_minutes"], 1440)
    return result


def validate_capture(payload, settings, now=None):
    if not isinstance(payload, dict):
        raise ContractError("Capture must be a JSON object.")
    if payload.get("schema_version") != 1 or type(payload.get("schema_version")) is not int:
        raise ContractError("Capture schema_version must be 1.")
    capture_id = payload.get("capture_id")
    if not isinstance(capture_id, str) or not re.fullmatch(r"[A-Za-z0-9_.-]{1,100}", capture_id):
        raise ContractError("capture_id must be 1 to 100 letters, digits, dots, underscores, or hyphens.")
    group = payload.get("task_group")
    if group not in settings["groups"]:
        raise ContractError("Task group is outside the configured exact TOYOTA scope.")
    start = aware_time(payload.get("started_at"), "started_at")
    finish = aware_time(payload.get("finished_at"), "finished_at")
    now = now or utcnow()
    if finish < start or finish > now + timedelta(minutes=5):
        raise ContractError("Capture times are reversed or more than five minutes in the future.")
    status = payload.get("status")
    if status not in ("complete", "incomplete", "failed"):
        raise ContractError("Capture status must be complete, incomplete, or failed.")
    error = text_field(payload.get("error", ""), "Capture error", 500)
    rows = payload.get("rows", [])
    if not isinstance(rows, list) or len(rows) > 20000:
        raise ContractError("Capture rows must be a list of at most 20,000 task rows.")
    total = payload.get("source_total")
    normalized = []
    if status == "complete":
        if not settings["source_timezone"]:
            raise ContractError("Set the source time zone in Setup before importing complete captures.")
        if payload.get("source_timezone") != settings["source_timezone"]:
            raise ContractError("Capture time zone must match the configured source time zone.")
        if payload.get("scope_verified") is not True or payload.get("total_verified") is not True:
            raise ContractError("Complete captures require verified scope and source total.")
        if payload.get("coverage") not in ("full_table", "absolute_rows"):
            raise ContractError("Complete captures require full_table or absolute_rows coverage.")
        integer(total, "Source total", 0, 20000)
        if total != len(rows):
            raise ContractError("Extracted row count differs from the verified source total.")
        positions = set()
        for row in rows:
            if not isinstance(row, dict):
                raise ContractError("Each task row must be an object.")
            if "scheduled_for" not in row:
                raise ContractError("Scheduled for must be present; use null only when genuinely unscheduled.")
            position = integer(row.get("row_index"), "Row index", 1, max(1, total))
            if position in positions:
                raise ContractError("Duplicate row positions indicate overlapping or incomplete traversal.")
            positions.add(position)
            normalized.append({
                "row_index": position,
                "case_number": text_field(row.get("case_number"), "Case number", 120, True),
                "task_type": text_field(row.get("task_type"), "Task type", 150, True),
                "priority": text_field(row.get("priority"), "Task priority", 80, True),
                "language": text_field(row.get("language"), "Language", 80),
                "scheduled_for": source_time(row.get("scheduled_for"), settings["source_timezone"]),
                "scheduled_raw": row.get("scheduled_for") or "",
                "assigned_to": text_field(row.get("assigned_to"), "Assigned to", 200),
            })
        if positions != set(range(1, total + 1)):
            raise ContractError("Complete captures require every absolute row position from 1 through the source total.")
    else:
        if not error:
            raise ContractError("Failed or incomplete captures need a short error description.")
        if rows:
            raise ContractError("Failed or incomplete captures contain health metadata only, not partial task rows.")
        if total is not None:
            integer(total, "Source total", 0, 20000)
    return {
        "capture_id": capture_id, "task_group": group, "started_at": stamp(start),
        "finished_at": stamp(finish), "status": status, "source_total": total,
        "error": error, "rows": normalized,
    }


def write_json_atomic(path, value):
    temporary = path.with_name(path.name + "." + uuid.uuid4().hex + ".tmp")
    try:
        temporary.write_text(json.dumps(value, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        temporary.replace(path)
    finally:
        temporary.unlink(missing_ok=True)


class Store:
    def __init__(self, home, demo=False):
        self.home = Path(home).resolve()
        self.demo = demo
        self.inbox_error = ""
        self.lock = threading.RLock()
        self.inbox = self.home / "inbox"
        for folder in (self.home, self.inbox, self.home / "processed", self.home / "rejected"):
            folder.mkdir(parents=True, exist_ok=True)
        self.settings_path = self.home / "settings.json"
        if not self.settings_path.exists():
            write_json_atomic(self.settings_path, DEFAULT_SETTINGS)
        self.settings = validate_settings(json.loads(self.settings_path.read_text(encoding="utf-8")))
        self.database = self.home / "bo-live.sqlite3"
        with self.connect() as db:
            version = db.execute("PRAGMA user_version").fetchone()[0]
            if version not in (0, 1):
                raise ContractError("This BO Live database requires a newer app version.")
            db.executescript(Path(__file__).with_name("schema.sql").read_text(encoding="utf-8"))

    @contextmanager
    def connect(self):
        db = sqlite3.connect(self.database, timeout=30)
        db.row_factory = sqlite3.Row
        db.execute("PRAGMA foreign_keys=ON")
        try:
            with db:
                yield db
        finally:
            db.close()

    def save_settings(self, value):
        with self.lock:
            candidate = validate_settings(value)
            if candidate["source_timezone"] != self.settings["source_timezone"]:
                with self.connect() as db:
                    exists = db.execute("SELECT 1 FROM captures WHERE status='complete' LIMIT 1").fetchone()
                if exists:
                    raise ContractError("Source time zone is fixed after the first complete capture. Use a new data folder to correct it.")
            write_json_atomic(self.settings_path, candidate)
            self.settings = candidate
            return dict(candidate)

    def ingest(self, payload, now=None):
        with self.lock:
            if isinstance(payload, dict) and payload.get("synthetic") and not self.demo:
                raise ContractError("Synthetic demo captures belong in demo mode, not the live workspace.")
            canonical = json.dumps(payload, ensure_ascii=False, sort_keys=True, separators=(",", ":"), allow_nan=False)
            if len(canonical.encode("utf-8")) > MAX_BYTES:
                raise ContractError("Capture exceeds the 16 MB limit.")
            capture = validate_capture(payload, self.settings, now=now)
            digest = hashlib.sha256(canonical.encode("utf-8")).hexdigest()
            with self.connect() as db:
                prior = db.execute("SELECT digest FROM captures WHERE capture_id=?", (capture["capture_id"],)).fetchone()
                if prior:
                    if prior["digest"] != digest:
                        raise ContractError("capture_id already exists with different evidence; publish a new capture ID.")
                    return {"capture_id": capture["capture_id"], "duplicate": True, "status": capture["status"]}
                db.execute("INSERT INTO captures VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)", (
                    capture["capture_id"], digest, capture["task_group"], capture["started_at"], capture["finished_at"],
                    stamp(now or utcnow()), capture["status"], capture["source_total"], capture["error"], canonical,
                ))
                db.executemany("INSERT INTO task_rows VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)", [
                    (capture["capture_id"], row["row_index"], row["case_number"], row["task_type"], row["priority"],
                     row["language"], row["scheduled_for"], row["scheduled_raw"], row["assigned_to"])
                    for row in capture["rows"]
                ])
            return {"capture_id": capture["capture_id"], "duplicate": False, "status": capture["status"]}

    def scan_inbox(self):
        outcomes = []
        with self.lock:
            for source in sorted(self.inbox.glob("*.json"))[:100]:
                target_folder = "processed"
                try:
                    if source.is_symlink() or source.stat().st_size > MAX_BYTES:
                        raise ContractError("Capture must be a regular local file of at most 16 MB.")
                    payload = json.loads(source.read_text(encoding="utf-8-sig"))
                    result = self.ingest(payload)
                    status = "accepted"
                    message = "Capture already imported." if result["duplicate"] else "Capture imported."
                except (ContractError, ValueError, OSError) as exc:
                    target_folder = "rejected"
                    status = "rejected"
                    message = str(exc)[:500] if isinstance(exc, ContractError) else "Could not read a valid UTF-8 JSON capture."
                target = self.home / target_folder / (uuid.uuid4().hex + ".json")
                try:
                    source.replace(target)
                except OSError:
                    message = "Capture handled but could not be archived; check local folder permissions."
                    status = "rejected"
                with self.connect() as db:
                    db.execute("INSERT INTO inbox_events(occurred_at,filename,status,message) VALUES (?,?,?,?)",
                               (stamp(utcnow()), source.name[:200], status, message))
                outcomes.append({"status": status, "message": message})
        return outcomes

    def acknowledge_events(self):
        with self.lock, self.connect() as db:
            db.execute("UPDATE inbox_events SET status='reviewed' WHERE status='rejected'")
        return {"status": "reviewed"}

    def state(self, now=None):
        now = now or utcnow()
        with self.lock, self.connect() as db:
            settings = dict(self.settings)
            groups, tasks = [], []
            order = settings["priority_order"]
            for group in settings["groups"]:
                latest = db.execute("SELECT rowid, * FROM captures WHERE task_group=? ORDER BY finished_at DESC,rowid DESC LIMIT 1", (group,)).fetchone()
                good = db.execute("SELECT * FROM captures WHERE task_group=? AND status='complete' ORDER BY finished_at DESC,rowid DESC LIMIT 1", (group,)).fetchone()
                health = "missing"
                if latest:
                    health = latest["status"] if latest["status"] != "complete" else "fresh"
                if good and health == "fresh" and now - aware_time(good["finished_at"], "capture time") > timedelta(minutes=settings["stale_minutes"]):
                    health = "stale"
                metrics = {"total": None, "past": None, "upcoming": None, "future": None,
                           "unscheduled": None, "urgent": None, "unassigned": None}
                if good:
                    observed = aware_time(good["finished_at"], "capture time")
                    metrics = {key: 0 for key in metrics}
                    if not settings["assignment_blank_means_unassigned"]:
                        metrics["unassigned"] = None
                    for record in db.execute("SELECT * FROM task_rows WHERE capture_id=? ORDER BY row_index", (good["capture_id"],)):
                        row = dict(record)
                        when = aware_time(row["scheduled_for"], "scheduled time") if row["scheduled_for"] else None
                        bucket = ("unscheduled" if when is None else "past" if when <= observed else
                                  "upcoming" if when <= observed + timedelta(minutes=settings["upcoming_minutes"]) else "future")
                        metrics["total"] += 1
                        metrics[bucket] += 1
                        metrics["urgent"] += row["priority"] == order[0]
                        if metrics["unassigned"] is not None:
                            metrics["unassigned"] += not row["assigned_to"].strip()
                        row.update(task_group=group, timing=bucket, health=health, observed_at=good["finished_at"],
                                   priority_rank=order.index(row["priority"]) if row["priority"] in order else len(order))
                        tasks.append(row)
                groups.append({"name": group, "health": health, "metrics": metrics,
                               "observed_at": good["finished_at"] if good else None,
                               "started_at": good["started_at"] if good else None,
                               "latest_attempt": latest["finished_at"] if latest else None,
                               "error": latest["error"] if latest and latest["status"] != "complete" else ""})
            tasks.sort(key=lambda row: (0 if row["timing"] == "past" else 1, row["priority_rank"],
                                       row["scheduled_for"] or "9999", row["task_group"], row["row_index"]))
            totals = {}
            for key in ("total", "past", "upcoming", "future", "unscheduled", "urgent", "unassigned"):
                values = [g["metrics"][key] for g in groups]
                totals[key] = sum(v for v in values if v is not None) if any(v is not None for v in values) else None
            history = [dict(r) for r in db.execute("SELECT capture_id,task_group,started_at,finished_at,status,source_total,error FROM captures ORDER BY received_at DESC,rowid DESC LIMIT 100")]
            events = [dict(r) for r in db.execute("SELECT * FROM inbox_events ORDER BY id DESC LIMIT 30")]
            inbox_attention = db.execute("SELECT COUNT(*) FROM inbox_events WHERE status='rejected'").fetchone()[0]
            starts = [g["started_at"] for g in groups if g["started_at"]]
            ends = [g["observed_at"] for g in groups if g["observed_at"]]
            fresh = sum(g["health"] == "fresh" for g in groups)
            return {"demo": self.demo, "settings": settings, "groups": groups, "tasks": tasks, "totals": totals,
                    "history": history, "events": events, "inbox_attention": inbox_attention,
                    "inbox_error": self.inbox_error,
                    "fresh_groups": fresh, "group_count": len(groups),
                    "all_fresh": fresh == len(groups), "scan_start": min(starts) if starts else None,
                    "scan_end": max(ends) if ends else None, "server_time": stamp(now),
                    "inbox_path": str(self.inbox), "home_path": str(self.home)}


CSV_FIELDS = ("case_number", "task_type", "priority", "language", "scheduled_for", "assigned_to")
CSV_ALIASES = {
    "casenumber": "case_number", "tasktype": "task_type", "taskpriority": "priority", "priority": "priority",
    "language": "language", "scheduledfor": "scheduled_for", "assignedto": "assigned_to",
}


def capture_from_csv(path, group, source_total, started_at, finished_at, settings, coverage="full_table"):
    """Bridge a complete PAD-extracted DataTable CSV into the capture contract."""
    rows = []
    if Path(path).stat().st_size > MAX_BYTES:
        raise ContractError("PAD CSV exceeds the 16 MB limit.")
    with Path(path).open(encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle)
        headers = reader.fieldnames or []
        mapping = {header: CSV_ALIASES.get(re.sub(r"[^a-z0-9]", "", header.lower())) for header in headers}
        selected = [key for key in mapping.values() if key]
        if set(selected) != set(CSV_FIELDS) or len(selected) != len(CSV_FIELDS):
            raise ContractError("CSV needs exactly one column for Case number, Task type, Task priority, Language, Scheduled for, and Assigned to.")
        for index, raw in enumerate(reader, 1):
            if index > 20000:
                raise ContractError("PAD CSV exceeds the 20,000-row capture limit.")
            if None in raw or any(raw[header] is None for header in headers):
                raise ContractError("CSV contains a malformed row.")
            row = {field: raw[header] for header, field in mapping.items() if field}
            row["row_index"] = index
            rows.append(row)
    payload = {"schema_version": 1, "capture_id": uuid.uuid4().hex, "task_group": group,
               "started_at": started_at, "finished_at": finished_at, "source_timezone": settings["source_timezone"],
               "status": "complete", "source_total": source_total, "scope_verified": True, "total_verified": True,
               "coverage": coverage, "rows": rows}
    validate_capture(payload, settings)
    return payload


def seed_demo(store, now=None):
    now = (now or utcnow()).replace(microsecond=0)
    settings = dict(DEFAULT_SETTINGS, source_timezone="UTC", assignment_blank_means_unassigned=True,
                    groups=["(TOYOTA) Agendas à Traiter", "(TOYOTA) REFACTURATION", "(TOYOTA) SUIVI VR", "(TOYOTA) TAKATA"])
    store.save_settings(settings)
    for number, group in enumerate(settings["groups"]):
        rows = []
        for index in range(1, (18, 11, 8, 5)[number] + 1):
            rows.append({"row_index": index, "case_number": f"DEMO-{number + 1:02}-{(index + 1) // 2:04}",
                         "task_type": "Call-back task", "priority": settings["priority_order"][(index + number) % 3],
                         "language": "French" if index % 3 else "Dutch",
                         "scheduled_for": stamp(now + timedelta(minutes=(index % 7 - 3) * 35)) if index % 9 else None,
                         "assigned_to": "" if index % 3 else f"Demo operator {index % 4 + 1}"})
        store.ingest({"schema_version": 1, "synthetic": True, "capture_id": f"demo-{number}-{uuid.uuid4().hex}", "task_group": group,
                      "started_at": stamp(now - timedelta(seconds=35 - number * 5)), "finished_at": stamp(now),
                      "source_timezone": "UTC", "status": "complete", "source_total": len(rows),
                      "scope_verified": True, "total_verified": True, "coverage": "full_table", "rows": rows}, now=now)
