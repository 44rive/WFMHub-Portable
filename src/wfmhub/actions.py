"""Import audited human classifications from Attendance Review.xlsx."""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path

from openpyxl import load_workbook

from .database import DatabaseConnection
from .rules import Rulebook


DECISION_SHEETS = ("REVIEW BOARD", "DECISIONS")
REQUIRED_HEADERS = {
    "Gap ID", "Decision Category", "Decision Status", "Reviewed By",
    "Comment", "Reviewed Date",
}
ALLOWED_STATUSES = {"Open", "Approved", "Dismissed"}


@dataclass(frozen=True)
class DecisionImportSummary:
    imported: int
    start: date
    end: date


def _as_date(value) -> date | None:
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    text = str(value or "").strip()
    try:
        return date.fromisoformat(text) if text else None
    except ValueError as exc:
        raise ValueError(f"Reviewed Date must be YYYY-MM-DD, received {text!r}") from exc


def _text(value) -> str | None:
    text = str(value or "").strip()
    return text or None


def import_attendance_decisions(
    conn: DatabaseConnection,
    path: Path,
    rulebook: Rulebook,
) -> DecisionImportSummary:
    """Atomically persist decisions; SQLite remains authoritative for gaps."""

    path = path.resolve()
    if not path.is_file():
        raise FileNotFoundError(path)
    workbook = load_workbook(path, read_only=True, data_only=True, keep_links=False)
    try:
        decision_sheet = next(
            (name for name in DECISION_SHEETS if name in workbook.sheetnames),
            None,
        )
        if decision_sheet is None:
            raise ValueError(
                "The workbook has no REVIEW BOARD sheet "
                "(legacy DECISIONS is also accepted)"
            )
        sheet = workbook[decision_sheet]
        headers = [str(cell.value or "").strip() for cell in sheet[4]]
        missing = sorted(REQUIRED_HEADERS - set(headers))
        if missing:
            raise ValueError(
                f"{decision_sheet} is missing decision columns: {', '.join(missing)}"
            )
        index = {header: headers.index(header) for header in headers if header}
        pending: list[tuple[str, str | None, str, str | None, str | None, date | None]] = []
        seen: set[str] = set()
        dates: list[date] = []
        for values in sheet.iter_rows(min_row=5, values_only=True):
            gap_id = _text(values[index["Gap ID"]])
            if not gap_id or gap_id == "No rows for this period.":
                continue
            if gap_id in seen:
                raise ValueError(f"Duplicate Gap ID in {decision_sheet}: {gap_id}")
            seen.add(gap_id)
            authoritative = conn.execute(
                "SELECT business_date FROM mart.correction_candidate WHERE correction_id=?",
                [gap_id],
            ).fetchone()
            if authoritative is None:
                raise ValueError(
                    f"Unknown or stale Gap ID {gap_id}. Rebuild Attendance Review before importing."
                )
            dates.append(authoritative[0])
            raw_status = _text(values[index["Decision Status"]]) or "Open"
            status = raw_status.title()
            if status not in ALLOWED_STATUSES:
                raise ValueError(
                    f"Gap {gap_id} has invalid Decision Status {raw_status!r}; "
                    "use Open, Approved or Dismissed."
                )
            activity = _text(values[index["Decision Category"]])
            if status == "Approved":
                if activity is None:
                    raise ValueError(f"Gap {gap_id} is Approved but has no Decision Category")
                if rulebook.classify_activity(activity) is None:
                    raise ValueError(
                        f"Gap {gap_id} uses unmapped Decision Category {activity!r}. "
                        "Choose a value from the workbook list or update the rulebook."
                    )
            elif status == "Dismissed":
                activity = None
            pending.append((
                gap_id, activity, status,
                _text(values[index["Reviewed By"]]),
                _text(values[index["Comment"]]),
                _as_date(values[index["Reviewed Date"]]),
            ))
        if not pending or not dates:
            raise ValueError("No Attendance Review decision rows were found")

        now = datetime.now()
        conn.execute("SAVEPOINT import_attendance_decisions")
        try:
            for gap_id, activity, status, reviewer, comment, reviewed_date in pending:
                conn.execute(
                    """INSERT INTO core.correction_action (
                           correction_id, confirmed_activity, validation_status,
                           owner, comment, injected_date, updated_at, imported_from
                       ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                       ON CONFLICT(correction_id) DO UPDATE SET
                           confirmed_activity=excluded.confirmed_activity,
                           validation_status=excluded.validation_status,
                           owner=excluded.owner, comment=excluded.comment,
                           injected_date=excluded.injected_date,
                           updated_at=excluded.updated_at,
                           imported_from=excluded.imported_from""",
                    [
                        gap_id, activity, status, reviewer, comment,
                        reviewed_date, now, path.name,
                    ],
                )
            conn.execute("RELEASE SAVEPOINT import_attendance_decisions")
        except Exception:
            conn.execute("ROLLBACK TO SAVEPOINT import_attendance_decisions")
            conn.execute("RELEASE SAVEPOINT import_attendance_decisions")
            raise
        return DecisionImportSummary(len(pending), min(dates), max(dates))
    finally:
        workbook.close()
