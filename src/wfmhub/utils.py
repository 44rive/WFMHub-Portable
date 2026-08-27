"""Small parsing and interval helpers shared by ingestion and models."""

from __future__ import annotations

import hashlib
import re
import unicodedata
from datetime import date, datetime, time, timedelta
from pathlib import Path
from typing import Any, Iterable


INVALID_IDS = {"", "-", "N/A", "NA", "NULL", "NONE"}
VERINT_INTERVAL_RE = re.compile(
    r"(?P<start>\d{1,2}/\d{1,2}/\d{4}\s+\d{1,2}:\d{2}\s+[AP]M)"
    r"\s*-\s*"
    r"(?P<end>\d{1,2}/\d{1,2}/\d{4}\s+\d{1,2}:\d{2}\s+[AP]M)",
    re.IGNORECASE,
)


def normalize_id(value: Any, reject_placeholders: bool = True) -> str | None:
    if value is None:
        return None
    if isinstance(value, bool):
        text = str(value)
    elif isinstance(value, int):
        text = str(value)
    elif isinstance(value, float) and value.is_integer():
        text = str(int(value))
    else:
        text = str(value).strip()
        if re.fullmatch(r"\d+\.0", text):
            text = text[:-2]
    if reject_placeholders and text.upper() in INVALID_IDS:
        return None
    return text or None


def normalize_name(value: Any) -> str | None:
    if value is None:
        return None
    text = unicodedata.normalize("NFD", str(value).upper().strip())
    text = "".join(char for char in text if unicodedata.category(char) != "Mn")
    tokens = [token for token in re.split(r"[\s,._/()\-]+", text) if token]
    return " ".join(sorted(tokens)) or None


def normalize_header(value: Any) -> str:
    text = unicodedata.normalize("NFKD", str(value or ""))
    text = "".join(char for char in text if not unicodedata.combining(char))
    return re.sub(r"[^A-Z0-9]+", "", text.upper())


def file_sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def parse_datetime(value: Any) -> datetime | None:
    if value is None or str(value).strip() == "":
        return None
    if isinstance(value, datetime):
        return value.replace(tzinfo=None)
    if isinstance(value, date):
        return datetime.combine(value, time())
    text = str(value).strip()
    for fmt in (
        None,
        "%m/%d/%Y %I:%M:%S %p",
        "%m/%d/%Y %I:%M %p",
        "%m/%d/%Y %H:%M:%S",
        "%m/%d/%Y %H:%M",
        "%Y-%m-%d %H:%M:%S",
        "%Y/%m/%d %H:%M:%S",
    ):
        try:
            return datetime.fromisoformat(text) if fmt is None else datetime.strptime(text, fmt)
        except ValueError:
            continue
    return None


def parse_date(value: Any) -> date | None:
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if value is None:
        return None
    text = str(value).strip()
    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%m/%d/%Y", "%d/%m/%Y"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    return None


def parse_time(value: Any) -> time | None:
    if isinstance(value, datetime):
        return value.time().replace(tzinfo=None)
    if isinstance(value, time):
        return value.replace(tzinfo=None)
    if value is None:
        return None
    text = str(value).strip()
    for fmt in ("%H:%M:%S", "%H:%M", "%I:%M %p", "%I:%M:%S %p"):
        try:
            return datetime.strptime(text, fmt).time()
        except ValueError:
            continue
    return None


def duration_seconds(value: Any) -> int | None:
    if value is None or str(value).strip() == "":
        return None
    if isinstance(value, timedelta):
        return round(value.total_seconds())
    if isinstance(value, time):
        return value.hour * 3600 + value.minute * 60 + value.second
    if isinstance(value, (int, float)):
        return round(float(value))
    text = str(value).strip()
    parts = text.split(":")
    try:
        if len(parts) == 3:
            return int(parts[0]) * 3600 + int(parts[1]) * 60 + round(float(parts[2]))
        if len(parts) == 2:
            return int(parts[0]) * 3600 + round(float(parts[1])) * 60
        return round(float(text))
    except ValueError:
        return None


def parse_verint_interval(value: Any) -> tuple[str | None, datetime | None, datetime | None]:
    raw = str(value or "").strip()
    if not raw:
        return None, None, None
    match = VERINT_INTERVAL_RE.search(raw)
    if not match:
        return raw or None, None, None
    activity_text = raw[: match.start()].strip()
    # Verint commonly prefixes activities with ".ORG |". Remove that single
    # prefix, but preserve internal pipes in real labels such as
    # "Product Loss | IT Failure".
    if "|" in activity_text:
        prefix, remainder = activity_text.split("|", 1)
        if prefix.strip().startswith("."):
            activity_text = remainder.strip()
    activity = activity_text or None
    start = datetime.strptime(match.group("start").upper(), "%m/%d/%Y %I:%M %p")
    end = datetime.strptime(match.group("end").upper(), "%m/%d/%Y %I:%M %p")
    return activity, start, end


def classify_assignment(raw: str | None, activity: str | None) -> str:
    shift = (raw or "").strip().upper()
    text = (activity or raw or "").upper()
    if shift == "OFF":
        return "Off"
    if any(token in text for token in ("SICKNESS", "VACATION", "LEAVE - UNPAID", "ANNUAL LEAVE")):
        return "Planned absence"
    if any(token in text for token in ("TRAINING", "TRAINER", "QUALITY MONITORING")):
        return "Non-phone planned"
    return "Work"


def classify_event(activity: str | None) -> str:
    text = (activity or "").upper().strip()
    if "LUNCH" in text:
        return "Lunch"
    if "BREAK" in text:
        return "Break"
    if any(token in text for token in ("SICKNESS", "VACATION", "LEAVE - UNPAID", "ANNUAL LEAVE")):
        return "Planned absence"
    if "GENERAL UNAVAILABILITY" in text or text in {"LATE", "EARLY LEAVING", "EARLY LEAVE"}:
        return "Planned adjustment"
    if any(token in text for token in ("TRAINING", "QUALITY", "KEY-USER", "MEETING")):
        return "Non-phone planned"
    return "Other planned"


def classify_status(status: str | None) -> str:
    text = (status or "").upper().strip()
    if text == "LOGGED OFF":
        return "Logged Off"
    if "LUNCH" in text:
        return "Lunch"
    if text in {"BREAK", "ON BREAK", "PAUSE ECRAN", "PAUSE ÉCRAN"}:
        return "Break"
    if text == "UNAVAILABLE":
        return "Unavailable"
    if any(token in text for token in ("AVAILABLE", "INBOUND", "OUTBOUND", "CALL SETUP", "RINGBACK", "HOLD", "WRAPUP")):
        return "Productive"
    return "Auxiliary"


def merge_intervals(intervals: Iterable[tuple[datetime, datetime]]) -> list[tuple[datetime, datetime]]:
    ordered = sorted((start, end) for start, end in intervals if start and end and end > start)
    merged: list[list[datetime]] = []
    for start, end in ordered:
        if not merged or start > merged[-1][1]:
            merged.append([start, end])
        else:
            merged[-1][1] = max(merged[-1][1], end)
    return [(start, end) for start, end in merged]


def clip_intervals(start: datetime, end: datetime, intervals: Iterable[tuple[datetime, datetime]]) -> list[tuple[datetime, datetime]]:
    return merge_intervals((max(start, a), min(end, b)) for a, b in intervals if b > start and a < end)


def interval_minutes(start: datetime, end: datetime, intervals: Iterable[tuple[datetime, datetime]]) -> int:
    return int(sum((b - a).total_seconds() for a, b in clip_intervals(start, end, intervals)) // 60)


def subtract_intervals(start: datetime, end: datetime, blocked: Iterable[tuple[datetime, datetime]]) -> list[tuple[datetime, datetime]]:
    remaining = [(start, end)]
    for block_start, block_end in clip_intervals(start, end, blocked):
        next_remaining: list[tuple[datetime, datetime]] = []
        for current_start, current_end in remaining:
            if block_end <= current_start or block_start >= current_end:
                next_remaining.append((current_start, current_end))
                continue
            if block_start > current_start:
                next_remaining.append((current_start, block_start))
            if block_end < current_end:
                next_remaining.append((block_end, current_end))
        remaining = next_remaining
    return [(a, b) for a, b in remaining if b > a]
