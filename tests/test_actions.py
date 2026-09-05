from __future__ import annotations

import sqlite3
import tempfile
import unittest
from datetime import date
from pathlib import Path

from openpyxl import Workbook

from wfmhub.actions import import_attendance_decisions
from wfmhub.database import DatabaseConnection
from wfmhub.rules import load_rulebook


REPO = Path(__file__).resolve().parents[1]


class AttendanceDecisionImportTests(unittest.TestCase):
    def test_invalid_row_rejects_the_entire_workbook(self):
        raw = sqlite3.connect(
            ":memory:",
            detect_types=sqlite3.PARSE_DECLTYPES | sqlite3.PARSE_COLNAMES,
        )
        conn = DatabaseConnection(raw)
        conn.execute(
            "CREATE TABLE mart.correction_candidate "
            "(correction_id VARCHAR PRIMARY KEY, business_date DATE)"
        )
        conn.execute(
            """CREATE TABLE core.correction_action (
                   correction_id VARCHAR PRIMARY KEY,
                   confirmed_activity VARCHAR, validation_status VARCHAR,
                   owner VARCHAR, comment VARCHAR, injected_date DATE,
                   updated_at TIMESTAMP, imported_from VARCHAR
               )"""
        )
        conn.executemany(
            "INSERT INTO mart.correction_candidate VALUES (?, ?)",
            [("gap-valid", date(2026, 8, 1)), ("gap-invalid", date(2026, 8, 2))],
        )

        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "Attendance Review.xlsx"
            workbook = Workbook()
            sheet = workbook.active
            sheet.title = "DECISIONS"
            sheet.append([])
            sheet.append([])
            sheet.append([])
            sheet.append([
                "Gap ID", "Decision Category", "Decision Status",
                "Reviewed By", "Comment", "Reviewed Date",
            ])
            sheet.append([
                "gap-valid", "Short sickness", "Approved",
                "Reviewer", "valid", date(2026, 8, 3),
            ])
            sheet.append([
                "gap-invalid", "Invented category", "Approved",
                "Reviewer", "invalid", date(2026, 8, 3),
            ])
            workbook.save(path)
            workbook.close()

            with self.assertRaisesRegex(ValueError, "unmapped Decision Category"):
                import_attendance_decisions(
                    conn, path,
                    load_rulebook(REPO, REPO / "config" / "default_rules.toml"),
                )

        self.assertEqual(
            conn.execute("SELECT count(*) FROM core.correction_action").fetchone()[0],
            0,
        )
        raw.close()


if __name__ == "__main__":
    unittest.main()
