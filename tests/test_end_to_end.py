from __future__ import annotations

import csv
import shutil
import tempfile
import unittest
import zipfile
from datetime import date
from pathlib import Path

from openpyxl import Workbook, load_workbook

from wfmhub.actions import import_actions
from wfmhub.config import ensure_user_config, load_config, write_source_root
from wfmhub.database import write_session
from wfmhub.ingestion import ingest_all
from wfmhub.models import refresh_models
from wfmhub.reports import build_report


REPO = Path(__file__).resolve().parents[1]


def make_fte(path: Path):
    path.parent.mkdir(parents=True, exist_ok=True)
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Agent"
    sheet.append(["Client ID", "Status", "Name", "Team leader", "Ops Manager", "LOB", "Market", "Language", "Location", "City", "FTE", "End date if leaver"])
    for agent_id in ("100", "200", "300"):
        sheet.append([agent_id, "Active", f"Agent {agent_id}", "TL 1", "Ops 1", "RSA", "RSA", "EN", "Onsite", "City", 1, None])
    workbook.save(path)


def make_schedule(path: Path):
    path.parent.mkdir(parents=True, exist_ok=True)
    headers = ["Name", "Data Source IDs", "Scheduling Period", "Before Overtime", "After Overtime", "Shift Assignment", "Shift Events"]
    with path.open("w", encoding="cp1252", newline="") as handle:
        writer = csv.writer(handle, delimiter="\t")
        writer.writerow(headers)
        writer.writerow(["08/01/2026", "", "", "", "", "", ""])
        writer.writerow(["Agent 100", "100", "August", "", "", ".ORG | Work 08/01/2026 8:00 AM-08/01/2026 4:00 PM", ".ORG | Late 08/01/2026 8:00 AM-08/01/2026 8:10 AM;"])
        writer.writerow(["Agent 200", "200", "August", "", "", ".ORG | Work 08/01/2026 8:00 AM-08/01/2026 4:00 PM", ""])
        writer.writerow(["Agent 300", "300", "August", "", "", ".ORG | Work 08/01/2026 8:00 AM-08/01/2026 4:00 PM", ""])
        writer.writerow(["08/02/2026", "", "", "", "", "", ""])
        writer.writerow(["Agent 300", "300", "August", "", "", ".ORG | Work 08/02/2026 8:00 AM-08/02/2026 4:00 PM", ""])


def make_lilo(path: Path, rows):
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.writer(handle)
        writer.writerow(["[Agent]", "[Agent ID]", "[First Log-on Time]", "[Last Log-off Time]"])
        writer.writerows(rows)


def make_status(path: Path):
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.writer(handle)
        writer.writerow(["[Serial Number]", "[Status]", "[Status Start Date and Time]", "[Agent]", "[Agent ID]", "[Status Duration]", "[Queue]"])
        writer.writerow(["one", "Available", "8/1/2026 9:00", "Agent 100", "100", "1:00:00", "Queue"])


def make_forecast(path: Path):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(
        "DATE_TIME_FORMAT\nMM/DD/YYYY hh:mm A\n"
        "Queue Name\tDate\tTime\tTime Interval\tVolume (Absolute For)\tAbandons (Absolute For)\tService Level (Absolute For)\tService Level (Absolute Req)\tActivity Handling Time (Absolute For)\tHeadcount Staffing (Absolute For)\tNet Staffing (Absolute For)\tFull Time Equivalents (Absolute For)\tFull Time Equivalents (Absolute Req)\n"
        "All\t08/01/2026\t09:00 AM\t1:00\t10\t1\t79\t80\t200\t3\t-1\t2\t3\n",
        encoding="cp1252",
    )


def make_apbe(path: Path):
    path.parent.mkdir(parents=True, exist_ok=True)
    workbook = Workbook()
    sheet = workbook.active
    sheet.append(["Report Name:", "APBE"])
    sheet.append([])
    sheet.append(["Period:", "test"])
    sheet.append([])
    sheet.append(["Date", "Language", "Queue ID", "Queue", "BusinessPartnerID", "LineOfBusiness", "15 Minute Periods of Day", "Offered_calls (w/o short calls)", "Answered_Calls", "Abandoned_Calls (w/o short calls)", "Short_calls < 5s", "Answered_Calls <= 15s", "Answered_Calls <= 20s", "Answered_Calls <= 30s", "Average_Speed_of_Answer", "Average_Talk_Time", "Average_Hold_Time", "Average Total Wrap Time"])
    sheet.append(["2026/08/01", "EN", 1, "Queue", "Partner", "RSA", "09:00", 10, 8, 2, 0, 7, 8, 8, 10, 100, 20, 10])
    workbook.save(path)


class EndToEndTests(unittest.TestCase):
    def test_refresh_builds_safe_attendance_gaps_and_excel(self):
        with tempfile.TemporaryDirectory() as folder:
            home = Path(folder) / "hub"
            source = Path(folder) / "source"
            shutil.copytree(REPO / "config", home / "config")
            shutil.copytree(REPO / "sql", home / "sql")
            make_fte(source / "FTE/FTE Count.xlsx")
            make_schedule(source / "Verint/Schedules & Activities/schedule.txt")
            make_lilo(source / "Storm/LILO/AP-Historical-Report---Agent-Login 2026-08-01.csv", [
                ["Agent 100", "100", "2026-08-01 08:20:00", "2026-08-01 16:00:00"],
                ["Agent 200", "200", "", ""],
            ])
            make_lilo(source / "Storm/LILO/AP-Historical-Report---Agent-Login 2026-08-02.csv", [
                ["Agent 300", "300", "2026-08-02 08:00:00", "2026-08-02 16:00:00"],
            ])
            make_status(source / "Storm/Agent Status/AP-Historical-Report---Agent-Status 2026-08-01.csv")
            make_forecast(source / "Verint/Forecast/forecast.txt")
            make_apbe(source / "Storm/APBE ALL WFM/apbe.xlsx")
            (source / "Storm/APFR KPI SUIVI JOUR").mkdir(parents=True)

            config_file = ensure_user_config(home)
            write_source_root(config_file, source)
            config = load_config(home)
            with write_session(config) as conn:
                ingest = ingest_all(conn, config)
                self.assertEqual(ingest.failed, 0)
                self.assertEqual(ingest.loaded, 7)
                model = refresh_models(conn, config, "test", date(2026, 8, 1), date(2026, 8, 2))
                attendance = {row[0]: row[1] for row in conn.execute("SELECT agent_day_key, attendance_result FROM mart.attendance_agent_day").fetchall()}
                self.assertEqual(attendance["20260801-200"], "No show")
                self.assertEqual(attendance["20260801-300"], "Missing LILO roster row")
                late = conn.execute("SELECT gap_start, gap_end, gap_minutes FROM mart.correction_candidate WHERE agent_id='100' AND detected_issue='Late'").fetchone()
                self.assertEqual(str(late[0]), "2026-08-01 08:10:00")
                self.assertEqual(str(late[1]), "2026-08-01 08:20:00")
                self.assertEqual(late[2], 10)
                basis = conn.execute("SELECT measurement_basis FROM mart.conformance_agent_day WHERE agent_day_key='20260801-100'").fetchone()[0]
                self.assertEqual(basis, "LILO span")
                self.assertEqual(model.forecast_rows, 1)
                self.assertEqual(model.intraday_rows, 1)
                report = build_report(conn, config, model.start, model.end)

                edited = load_workbook(report)
                gap_sheet = edited["GAPS"]
                gap_headers = {cell.value: cell.column for cell in gap_sheet[4]}
                gap_sheet.cell(5, gap_headers["Validation Status"], "Validated")
                gap_sheet.cell(5, gap_headers["Owner"], "WFM")
                edited.save(report)
                edited.close()
                self.assertEqual(import_actions(conn, report), 1)
                self.assertEqual(conn.execute("SELECT validation_status FROM core.correction_action").fetchone()[0], "Validated")

            self.assertTrue(report.exists())
            workbook = load_workbook(report, read_only=True, data_only=True)
            try:
                self.assertEqual(workbook.sheetnames, ["START_HERE", "SUMMARY", "ATTENDANCE", "GAPS", "RTA", "INTRADAY", "DATA_QUALITY", "SOURCE_HEALTH"])
            finally:
                workbook.close()
            with zipfile.ZipFile(report) as archive:
                self.assertFalse(any("externalLinks" in name for name in archive.namelist()))


if __name__ == "__main__":
    unittest.main()
