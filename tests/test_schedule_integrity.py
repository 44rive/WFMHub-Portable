from __future__ import annotations

import sqlite3
import unittest
from datetime import date, datetime, timedelta
from pathlib import Path

from wfmhub.database import DatabaseConnection, _migration_statements
from wfmhub.models import _build_schedule_integrity
from wfmhub.rules import load_rulebook


REPO = Path(__file__).resolve().parents[1]


class ScheduleIntegrityTests(unittest.TestCase):
    def test_agent_status_detects_supported_recurring_shift_displacement(self):
        raw = sqlite3.connect(
            ":memory:",
            detect_types=sqlite3.PARSE_DECLTYPES | sqlite3.PARSE_COLNAMES,
        )
        conn = DatabaseConnection(raw)
        migration = (
            REPO / "sql" / "migrations" / "016_powerbi_wfm_control_tower.sql"
        ).read_text(encoding="utf-8")
        for statement in _migration_statements(migration):
            conn.execute(statement)

        attendance = []
        statuses = {}
        for offset in range(3):
            day = date(2026, 9, 7) + timedelta(days=offset)
            scheduled_start = datetime.combine(day, datetime.min.time()).replace(hour=9)
            scheduled_end = scheduled_start + timedelta(hours=8)
            attendance.append({
                "agent_day_key": f"{day:%Y%m%d}-100",
                "business_date": day,
                "agent_id": "100",
                "agent_name": "Agent 100",
                "team_leader": "TL A",
                "ops_manager": "Ops A",
                "lob": "RSA NL",
                "language": "NL",
                "scheduled_start": scheduled_start,
                "scheduled_end": scheduled_end,
                "shift_state": "COMPLETE",
                "assignment_type": "Working",
                "planned_work_minutes": 480,
                "planning_overlay_minutes": 0,
                "first_login": None,
                "last_logout": None,
            })
            observed_start = scheduled_start - timedelta(hours=1)
            observed_end = scheduled_end - timedelta(hours=1)
            statuses[(day, "100")] = [{
                "source_file_id": f"file-{offset}",
                "source_row": 2,
                "serial_number": str(offset),
                "agent_id": "100",
                "status": "Disponible",
                "actual_category": "Productive",
                "status_start": observed_start,
                "status_end": observed_end,
                "source_file": "status.csv",
            }]

        count = _build_schedule_integrity(
            conn,
            load_rulebook(REPO, REPO / "config" / "default_rules.toml"),
            attendance,
            statuses,
            datetime(2026, 9, 10, 0, 0),
        )

        self.assertEqual(count, 3)
        rows = conn.execute(
            """SELECT classification, start_delta_minutes, end_delta_minutes,
                      recurrence_count, is_recurring, confidence
               FROM mart.schedule_integrity_agent_day
               ORDER BY business_date"""
        ).fetchall()
        self.assertEqual(rows[0][:3], ("Shifted early", -60, -60))
        self.assertEqual(rows[-1][3:], (3, 1, "High"))
        conn.close()

    def test_time_off_and_incomplete_shifts_are_not_integrity_cases(self):
        raw = sqlite3.connect(
            ":memory:",
            detect_types=sqlite3.PARSE_DECLTYPES | sqlite3.PARSE_COLNAMES,
        )
        conn = DatabaseConnection(raw)
        migration = (
            REPO / "sql" / "migrations" / "016_powerbi_wfm_control_tower.sql"
        ).read_text(encoding="utf-8")
        for statement in _migration_statements(migration):
            conn.execute(statement)
        base = {
            "agent_id": "100", "agent_name": "Agent 100", "team_leader": "TL A",
            "ops_manager": "Ops A", "lob": "RSA NL", "language": "NL",
            "assignment_type": "Working", "planned_work_minutes": 420,
            "first_login": None, "last_logout": None,
        }
        day = date(2026, 9, 7)
        start = datetime(2026, 9, 7, 9)
        rows = [
            {
                **base, "agent_day_key": "20260907-100", "business_date": day,
                "scheduled_start": start, "scheduled_end": start + timedelta(hours=8),
                "shift_state": "COMPLETE", "planning_overlay_minutes": 60,
            },
            {
                **base, "agent_day_key": "20260908-100", "business_date": day + timedelta(days=1),
                "scheduled_start": start + timedelta(days=1),
                "scheduled_end": start + timedelta(days=1, hours=8),
                "shift_state": "IN_PROGRESS", "planning_overlay_minutes": 0,
            },
        ]
        count = _build_schedule_integrity(
            conn,
            load_rulebook(REPO, REPO / "config" / "default_rules.toml"),
            rows, {}, datetime(2026, 9, 8, 12),
        )
        self.assertEqual(count, 0)
        conn.close()


if __name__ == "__main__":
    unittest.main()
