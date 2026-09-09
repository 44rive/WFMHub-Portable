from __future__ import annotations

import csv
import shutil
import tempfile
import unittest
import zipfile
from dataclasses import replace
from datetime import date, datetime, timezone
from pathlib import Path
from unittest.mock import patch
from zoneinfo import ZoneInfoNotFoundError

from openpyxl import Workbook, load_workbook

from wfmhub.actions import import_attendance_decisions
from wfmhub.config import ensure_user_config, load_config, write_source_root
from wfmhub.database import write_session
from wfmhub.ingestion import ingest_all
from wfmhub.models import (
    _evaluation_time,
    refresh_models,
    refresh_pcs_models,
    resolve_period,
)
from wfmhub.on_demand_analysis import build_analysis_workbook
from wfmhub.pcs_tracker import PCS_TRACKER_FILENAME
from wfmhub.shared_feeds import (
    PCS_AGENT_SCORECARD_HEADERS,
    PCS_COACHING_HEADERS,
    PCS_DAILY_SCORECARD_HEADERS,
    PCS_LOB_SCORECARD_HEADERS,
    PCS_RESULTS_HEADERS,
)
from wfmhub.exports import export_dataset
from wfmhub.report_packs import build_report_pack
from wfmhub.reports import build_report
from wfmhub.rules import load_rulebook


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
        writer.writerow(["Agent 300", "300", "August", "", "", ".ORG | Work 08/01/2026 8:00 AM-08/01/2026 4:00 PM", ".ORG | Late 08/01/2026 8:00 AM-08/01/2026 8:10 AM;"])
        writer.writerow(["Worldwide Agent", "999", "August", "", "", ".ORG | Work 08/01/2026 8:00 AM-08/01/2026 4:00 PM", ""])
        writer.writerow(["08/02/2026", "", "", "", "", "", ""])
        writer.writerow(["Agent 300", "300", "August", "", "", ".ORG | Work 08/02/2026 8:00 AM-08/02/2026 4:00 PM", ""])


def make_start_end_schedule(path: Path):
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="cp1252", newline="") as handle:
        writer = csv.writer(handle, delimiter="\t")
        writer.writerow(["Name", "Data Source IDs", "08/01/2026", "08/02/2026"])
        writer.writerow(["Agent 100", "100", ".ORG | Work 08/01/2026 8:00 AM-08/01/2026 4:00 PM", ""])
        writer.writerow(["Agent 200", "200", ".ORG | Work 08/01/2026 8:00 AM-08/01/2026 4:00 PM", ""])
        writer.writerow(["Agent 300", "300", ".ORG | Work 08/01/2026 8:00 AM-08/01/2026 4:00 PM", ".ORG | Work 08/02/2026 8:00 AM-08/02/2026 4:00 PM"])


def make_partial_start_end_schedule(path: Path):
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="cp1252", newline="") as handle:
        writer = csv.writer(handle, delimiter="\t")
        writer.writerow(["Name", "Data Source IDs", "08/01/2026"])
        writer.writerow(["Agent 100", "100", ".ORG | Work 08/01/2026 8:00 AM-08/01/2026 4:00 PM"])


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
        writer.writerow(["one", "Available", "8/1/2026 8:20", "Agent 100", "100", "3:40:00", "Queue"])
        writer.writerow(["two", "Logged Off", "8/1/2026 12:00", "Agent 100", "100", "1:00:00", "Queue"])
        writer.writerow(["three", "Available", "8/1/2026 13:00", "Agent 100", "100", "3:00:00", "Queue"])
        writer.writerow(["world", "Available", "8/1/2026 9:00", "Worldwide Agent", "999", "1:00:00", "Queue"])


def make_logged_off_status(path: Path):
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.writer(handle)
        writer.writerow(["[Serial Number]", "[Status]", "[Status Start Date and Time]", "[Agent]", "[Agent ID]", "[Status Duration]", "[Queue]"])
        # Storm serial numbers can restart in another export. This deliberately
        # collides with Agent 100's first row and must not remove either status.
        writer.writerow(["one", "Logged Off", "8/1/2026 8:00", "Agent 200", "200", "8:00:00", "Queue"])


def make_meal_aux_status(path: Path):
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.writer(handle)
        writer.writerow(["[Serial Number]", "[Status]", "[Status Start Date and Time]", "[Agent]", "[Agent ID]", "[Status Duration]", "[Queue]"])
        writer.writerow(["available", "Available", "8/1/2026 8:00", "Agent 100", "100", "4:00:00", "Queue"])
        writer.writerow(["meal", "Meal Aux", "8/1/2026 12:00", "Agent 100", "100", "1:00:00", "Queue"])


def make_forecast(path: Path):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(
        "DATE_TIME_FORMAT\nMM/DD/YYYY hh:mm A\n"
        "Queue Name\tDate\tTime\tTime Interval\tVolume (Absolute For)\tAbandons (Absolute For)\tService Level (Absolute For)\tService Level (Absolute Req)\tActivity Handling Time (Absolute For)\tHeadcount Staffing (Absolute For)\tNet Staffing (Absolute For)\tFull Time Equivalents (Absolute For)\tFull Time Equivalents (Absolute Req)\n"
        "All\t08/01/2026\t09:00 AM\t0:15\t2\t1\t70\t80\t180\t3\t-1\t2\t3\n"
        "All\t08/01/2026\t09:15 AM\t0:15\t3\t0\t80\t80\t200\t3\t-1\t2\t3\n"
        "All\t08/01/2026\t09:30 AM\t0:15\t2\t0\t80\t80\t200\t3\t-1\t2\t3\n"
        "All\t08/01/2026\t09:45 AM\t0:15\t3\t0\t90\t80\t220\t3\t-1\t2\t3\n",
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


def make_apfr(path: Path):
    path.parent.mkdir(parents=True, exist_ok=True)
    workbook = Workbook()
    sheet = workbook.active
    sheet.append(["Report Name:", "APFR"])
    sheet.append([])
    sheet.append(["Date", "BusinessPartnerID", "LineOfBusiness", "15 Minute Periods of Day", "APPELS ENTRANTS", "APPELS RÉP", "APPELS ABAN", "Short_calls < 5s", "APPELS RÉP <= 15s", "APPELS RÉP <= 20s", "APPELS RÉP <= 30s", "Average_Speed_of_Answer", "Average_Talk_Time", "Average_Hold_Time", "Average Total Wrap Time"])
    sheet.append(["2026/08/01", "Partner FR", "RSA", "09:15", 12, 10, 2, 0, 8, 9, 10, 12, 110, 15, 10])
    workbook.save(path)


def make_calls(path: Path):
    path.parent.mkdir(parents=True, exist_ok=True)
    headers = [
        "[Call Date/Time]", "[Call End Date/Time]", "[Call ID]",
        "[Call Reference Number]", "[Agent ID]", "[Agent]",
        "[Talk Time]", "[Hold Time]", "[Total Wrap Time]",
        "[Call Direction]", "[PostCallSurveyMode]", "[PCSStatus]",
        "[Question 1]", "[Question 2]", "[Question 3]", "[Queue]",
    ]
    with path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=headers)
        writer.writeheader()
        writer.writerow({
            "[Call Date/Time]": "8/1/2026 9:00", "[Call End Date/Time]": "8/1/2026 9:05",
            "[Call ID]": "call-100", "[Call Reference Number]": "ref-100",
            "[Agent ID]": "100", "[Agent]": "Agent 100", "[Talk Time]": "0:04:00",
            "[Hold Time]": "0:00:30", "[Total Wrap Time]": "0:00:30",
            "[Call Direction]": "I", "[PostCallSurveyMode]": "2", "[PCSStatus]": "1",
            "[Question 1]": "5", "[Question 2]": "4", "[Question 3]": "Good",
            "[Queue]": "APBN_BRU_MOBILITY_Ford_Assistance_FR",
        })
        writer.writerow({
            "[Call Date/Time]": "8/1/2026 9:00", "[Call End Date/Time]": "8/1/2026 9:02",
            "[Call ID]": "world", "[Call Reference Number]": "world",
            "[Agent ID]": "999", "[Agent]": "Worldwide Agent", "[Talk Time]": "0:02:00",
            "[Hold Time]": "0:00:00", "[Total Wrap Time]": "0:00:00",
            "[Call Direction]": "I",
        })
        writer.writerow({
            "[Call Date/Time]": "8/1/2026 11:00", "[Call End Date/Time]": "8/1/2026 11:04",
            "[Call ID]": "low-200", "[Call Reference Number]": "ref-low-200",
            "[Agent ID]": "200", "[Agent]": "Agent 200", "[Talk Time]": "0:03:00",
            "[Hold Time]": "0:00:30", "[Total Wrap Time]": "0:00:30",
            "[Call Direction]": "I", "[PostCallSurveyMode]": "2", "[PCSStatus]": "1",
            "[Question 1]": "2", "[Question 2]": "3", "[Question 3]": "Needs follow-up",
            "[Queue]": "APBN_BRU_MOBILITY_Ford_Assistance_FR",
        })
        for call_id, direction, status, q1, q2 in (
            ("half-score", "I", "1", "4.5", "5"),
            ("outbound-score", "O", "1", "1", ""),
            ("q2-only", "I", "1", "", "1"),
            ("answer-no-status", "I", "0", "*", ""),
            ("zero-score", "I", "1", "0", ""),
        ):
            writer.writerow({
                "[Call Date/Time]": "8/1/2026 10:00",
                "[Call End Date/Time]": "8/1/2026 10:00",
                "[Call ID]": call_id, "[Call Reference Number]": call_id,
                "[Agent ID]": "100", "[Agent]": "Agent 100",
                "[Call Direction]": direction, "[PostCallSurveyMode]": "2",
                "[PCSStatus]": status, "[Question 1]": q1,
                "[Question 2]": q2,
                "[Queue]": "APBN_BRU_MOBILITY_Ford_Assistance_FR",
            })


class EndToEndTests(unittest.TestCase):
    def test_evaluation_time_uses_named_zone_and_survives_missing_tzdata(self):
        self.assertEqual(
            _evaluation_time("Europe/Berlin", datetime(2026, 8, 1, 12, tzinfo=timezone.utc)),
            datetime(2026, 8, 1, 14),
        )
        supplied = datetime(2026, 8, 1, 14)
        with patch("wfmhub.models.ZoneInfo", side_effect=ZoneInfoNotFoundError("missing")):
            self.assertEqual(_evaluation_time("Europe/Berlin", supplied), supplied)

    def test_lilo_logout_inside_meal_aux_does_not_create_a_meal_gap(self):
        with tempfile.TemporaryDirectory() as folder:
            home = Path(folder) / "hub"
            source = Path(folder) / "source"
            (home / "config").mkdir(parents=True)
            for name in (
                "default.toml", "default_rules.toml", "default_metrics.toml",
                "default_analytics.toml", "default_reports.toml",
            ):
                shutil.copy2(REPO / "config" / name, home / "config" / name)
            shutil.copytree(REPO / "sql", home / "sql")
            make_fte(source / "FTE/FTE Count.xlsx")
            make_start_end_schedule(source / "Verint/Schedules & Activities/StartEndTimes.txt")
            make_lilo(source / "Storm/LILO/LILO 2026-08-01.csv", [
                ["Agent 100", "100", "2026-08-01 08:00:00", "2026-08-01 12:15:00"],
            ])
            make_meal_aux_status(
                source / "Storm/Agent Status/Agent Status 2026-08-01.csv",
            )

            config_file = ensure_user_config(home)
            write_source_root(config_file, source)
            config = load_config(home)
            with write_session(config) as conn:
                self.assertEqual(ingest_all(conn, config).failed, 0)
                refresh_models(
                    conn, config, "meal-aux-boundary",
                    date(2026, 8, 1), date(2026, 8, 1),
                    as_of=datetime(2026, 8, 1, 17, 0),
                )
                attendance = conn.execute(
                    "SELECT actual_last_seen, uncoded_early_leave_minutes "
                    "FROM mart.attendance_agent_day WHERE agent_id='100'",
                ).fetchone()
                self.assertEqual(str(attendance[0]), "2026-08-01 13:00:00")
                self.assertEqual(attendance[1], 180)
                self.assertEqual(
                    conn.execute(
                        "SELECT actual_category, mismatch_type, is_gap "
                        "FROM mart.shift_timeline_segment "
                        "WHERE agent_id='100' AND segment_start=? AND segment_end=?",
                        [datetime(2026, 8, 1, 12), datetime(2026, 8, 1, 13)],
                    ).fetchone(),
                    ("Lunch", "MATCH", 0),
                )
                self.assertEqual(
                    conn.execute(
                        "SELECT count(*) FROM mart.correction_candidate "
                        "WHERE agent_id='100' AND gap_start<? AND gap_end>?",
                        [datetime(2026, 8, 1, 13), datetime(2026, 8, 1, 12)],
                    ).fetchone()[0],
                    0,
                )
                self.assertEqual(
                    str(conn.execute(
                        "SELECT gap_start FROM mart.correction_candidate "
                        "WHERE agent_id='100' AND detected_issue='Early leave'",
                    ).fetchone()[0]),
                    "2026-08-01 13:00:00",
                )

    def test_activities_shift_assignment_is_safe_schedule_fallback(self):
        with tempfile.TemporaryDirectory() as folder:
            home = Path(folder) / "hub"
            source = Path(folder) / "source"
            (home / "config").mkdir(parents=True)
            for name in (
                "default.toml", "default_rules.toml", "default_metrics.toml",
                "default_analytics.toml", "default_reports.toml",
            ):
                shutil.copy2(REPO / "config" / name, home / "config" / name)
            shutil.copytree(REPO / "sql", home / "sql")
            make_fte(source / "FTE/FTE Count.xlsx")
            make_schedule(source / "Verint/Schedules & Activities/Activities.txt")
            make_lilo(source / "Storm/LILO/LILO 2026-08-01.csv", [
                ["Agent 100", "100", "2026-08-01 08:00:00", "2026-08-01 16:00:00"],
                ["Agent 200", "200", "", ""],
                ["Agent 300", "300", "", ""],
            ])

            config_file = ensure_user_config(home)
            write_source_root(config_file, source)
            config = load_config(home)
            with write_session(config) as conn:
                self.assertEqual(ingest_all(conn, config).failed, 0)
                refresh_models(
                    conn, config, "activities-fallback",
                    date(2026, 8, 1), date(2026, 8, 1),
                    as_of=datetime(2026, 8, 1, 17, 0),
                )
                self.assertEqual(
                    conn.execute(
                        "SELECT count(*) FROM meta.source_file "
                        "WHERE source_family='schedule' AND active=true AND source_variant='START_END'"
                    ).fetchone()[0],
                    0,
                )
                self.assertEqual(
                    conn.execute("SELECT count(*) FROM mart.attendance_agent_day").fetchone()[0],
                    3,
                )
                self.assertEqual(
                    conn.execute(
                        "SELECT attendance_result FROM mart.attendance_agent_day WHERE agent_id='100'"
                    ).fetchone()[0],
                    "Present",
                )
                fallback_issue = conn.execute(
                    "SELECT severity, details FROM meta.quality_issue "
                    "WHERE issue_type='Dedicated StartEndTimes schedule not loaded'"
                ).fetchone()
                self.assertEqual(fallback_issue[0], "REVIEW")
                self.assertIn("Activities Shift Assignment", fallback_issue[1])

    def test_activities_fills_start_end_coverage_per_agent_day(self):
        with tempfile.TemporaryDirectory() as folder:
            home = Path(folder) / "hub"
            source = Path(folder) / "source"
            (home / "config").mkdir(parents=True)
            for name in (
                "default.toml", "default_rules.toml", "default_metrics.toml",
                "default_analytics.toml", "default_reports.toml",
            ):
                shutil.copy2(REPO / "config" / name, home / "config" / name)
            shutil.copytree(REPO / "sql", home / "sql")
            make_fte(source / "FTE/FTE Count.xlsx")
            schedule_folder = source / "Verint/Schedules & Activities"
            make_partial_start_end_schedule(schedule_folder / "StartEndTimes.txt")
            make_schedule(schedule_folder / "Activities.txt")
            make_lilo(source / "Storm/LILO/LILO 2026-08-01.csv", [
                ["Agent 100", "100", "2026-08-01 08:00:00", "2026-08-01 16:00:00"],
                ["Agent 200", "200", "2026-08-01 08:00:00", "2026-08-01 16:00:00"],
                ["Agent 300", "300", "2026-08-01 08:00:00", "2026-08-01 16:00:00"],
            ])

            config_file = ensure_user_config(home)
            write_source_root(config_file, source)
            config = load_config(home)
            with write_session(config) as conn:
                self.assertEqual(ingest_all(conn, config).failed, 0)
                refresh_models(
                    conn, config, "partial-start-end",
                    date(2026, 8, 1), date(2026, 8, 1),
                    as_of=datetime(2026, 8, 1, 17, 0),
                )
                self.assertEqual(
                    conn.execute("SELECT count(*) FROM mart.attendance_agent_day").fetchone()[0],
                    3,
                )
                self.assertEqual(
                    conn.execute(
                        "SELECT count(*) FROM meta.quality_issue "
                        "WHERE issue_type='StartEndTimes coverage incomplete' AND severity='REVIEW'"
                    ).fetchone()[0],
                    1,
                )

    def test_full_day_pto_drives_expected_work_and_verint_completeness(self):
        with tempfile.TemporaryDirectory() as folder:
            home = Path(folder) / "hub"
            source = Path(folder) / "source"
            (home / "config").mkdir(parents=True)
            for name in (
                "default.toml", "default_rules.toml", "default_metrics.toml",
                "default_analytics.toml", "default_reports.toml",
            ):
                shutil.copy2(REPO / "config" / name, home / "config" / name)
            shutil.copytree(REPO / "sql", home / "sql")

            fte = source / "FTE/FTE Count.xlsx"
            make_fte(fte)
            workbook = load_workbook(fte)
            pto = workbook.create_sheet("PTO")
            pto.append([
                "Client ID", "Name", "Start date", "End date", "Day coverage",
                "Start time", "End time", "PTO type", "Approval status", "Comment",
            ])
            pto.append(["200", "Agent 200", date(2026, 8, 1), date(2026, 8, 1), "Full day", None, None, "Vacation", "Approved", "Approved request"])
            pto.append(["100", "Agent 100", date(2026, 8, 1), date(2026, 8, 1), "Partial day", "12:00", "13:00", "Personal leave", "Approved", "Appointment"])
            pto.append(["300", "Agent 300", date(2026, 8, 1), date(2026, 8, 1), "Partial day", "10:00", "11:00", "Personal leave", "Approved", "Appointment"])
            pto.append(["300", "Agent 300", date(2026, 8, 1), date(2026, 8, 1), "Full day", None, None, "Vacation", "Pending", "Must not affect attendance yet"])
            away = workbook.create_sheet("Away")
            away.append(["Client ID", "Name", "Start date", "End date", "Away type", "Case status", "Comment"])
            away.append(["300", "Agent 300", date(2026, 8, 1), date(2026, 8, 1), "Long sickness", "Planned", "Future planning only"])
            away.append(["300", "Agent 300", date(2026, 8, 1), date(2026, 8, 1), "Long sickness", "Cancelled", "Must not affect attendance"])
            away.append(["300", "Agent 300", date(2026, 8, 2), None, "Long sickness", "Active", "Open case"])
            workbook.save(fte)
            workbook.close()
            make_start_end_schedule(source / "Verint/Schedules & Activities/StartEndTimes.txt")
            make_lilo(source / "Storm/LILO/LILO 2026-08-01.csv", [
                ["Agent 100", "100", "2026-08-01 08:00:00", "2026-08-01 16:00:00"],
                ["Agent 200", "200", "", ""],
                ["Agent 300", "300", "", ""],
            ])
            status_path = source / "Storm/Agent Status/Status 2026-08-01.csv"
            status_path.parent.mkdir(parents=True, exist_ok=True)
            with status_path.open("w", encoding="utf-8-sig", newline="") as handle:
                writer = csv.writer(handle)
                writer.writerow(["[Serial Number]", "[Status]", "[Status Start Date and Time]", "[Agent]", "[Agent ID]", "[Status Duration]", "[Queue]"])
                writer.writerow(["pto-gap", "Logged Off", "8/1/2026 12:00", "Agent 100", "100", "1:00:00", "Queue"])

            config_file = ensure_user_config(home)
            write_source_root(config_file, source)
            config = load_config(home)
            with write_session(config) as conn:
                self.assertEqual(ingest_all(conn, config).failed, 0)
                refresh_models(
                    conn, config, "pto-test", date(2026, 8, 1), date(2026, 8, 2),
                    as_of=datetime(2026, 8, 2, 17, 0),
                )
                attendance = conn.execute(
                    """SELECT attendance_result, planned_work_minutes,
                              planning_overlay_minutes, requires_call
                       FROM mart.attendance_agent_day
                       WHERE business_date='2026-08-01' AND agent_id='200'"""
                ).fetchone()
                self.assertEqual(attendance, ("PTO", 0, 480, 0))
                self.assertEqual(
                    conn.execute(
                        "SELECT count(*) FROM mart.correction_candidate "
                        "WHERE business_date='2026-08-01' AND agent_id='200'"
                    ).fetchone()[0],
                    0,
                )
                partial = conn.execute(
                    """SELECT planned_work_minutes, planning_overlay_minutes
                       FROM mart.attendance_agent_day
                       WHERE business_date='2026-08-01' AND agent_id='100'"""
                ).fetchone()
                self.assertEqual(partial, (420, 60))
                planned_past = conn.execute(
                    """SELECT attendance_result, planning_overlay_minutes, no_show_minutes
                       FROM mart.attendance_agent_day
                       WHERE business_date='2026-08-01' AND agent_id='300'"""
                ).fetchone()
                self.assertNotEqual(planned_past[0], "Away")
                self.assertEqual(planned_past, ("No show - partial time off", 60, 420))
                split_gaps = conn.execute(
                    """SELECT time(gap_start), time(gap_end), gap_minutes
                       FROM mart.correction_candidate
                       WHERE business_date='2026-08-01' AND agent_id='300'
                       ORDER BY gap_start"""
                ).fetchall()
                self.assertEqual(split_gaps, [
                    ("08:00:00", "10:00:00", 120),
                    ("11:00:00", "16:00:00", 300),
                ])
                self.assertEqual(
                    conn.execute(
                        "SELECT count(*) FROM mart.correction_candidate "
                        "WHERE business_date='2026-08-01' AND agent_id='100'"
                    ).fetchone()[0],
                    0,
                )
                staffing = conn.execute(
                    """SELECT gross_scheduled_fte, planned_time_off_fte, scheduled_fte
                       FROM mart.staffing_interval WHERE time(interval_start)='08:00:00'
                         AND business_date='2026-08-01'
                         AND lob='RSA' AND language='EN'"""
                ).fetchone()
                self.assertEqual(staffing, (3.0, 1.0, 2.0))
                noon_staffing = conn.execute(
                    """SELECT gross_scheduled_fte, planned_time_off_fte, scheduled_fte
                       FROM mart.staffing_interval WHERE time(interval_start)='12:00:00'
                         AND business_date='2026-08-01'
                         AND lob='RSA' AND language='EN'"""
                ).fetchone()
                self.assertEqual(noon_staffing, (3.0, 2.0, 1.0))
                self.assertEqual(
                    conn.execute(
                        "SELECT final_ledger_status FROM mart.verint_final_absence_agent_day "
                        "WHERE business_date='2026-08-01' AND agent_id='200'"
                    ).fetchone()[0],
                    "CLEAR",
                )
                active_away = conn.execute(
                    """SELECT attendance_result, planned_work_minutes,
                              planning_overlay_minutes, requires_call
                       FROM mart.attendance_agent_day
                       WHERE business_date='2026-08-02' AND agent_id='300'"""
                ).fetchone()
                self.assertEqual(active_away, ("Away", 0, 480, 0))
                self.assertEqual(
                    conn.execute(
                        "SELECT count(*) FROM mart.correction_candidate "
                        "WHERE business_date='2026-08-02' AND agent_id='300'"
                    ).fetchone()[0],
                    0,
                )
                away_staffing = conn.execute(
                    """SELECT gross_scheduled_fte, planned_time_off_fte,
                              scheduled_fte
                       FROM mart.staffing_interval
                       WHERE business_date='2026-08-02'
                         AND time(interval_start)='08:00:00'
                         AND lob='RSA' AND language='EN'"""
                ).fetchone()
                self.assertEqual(away_staffing, (1.0, 1.0, 0.0))
                self.assertEqual(
                    conn.execute(
                        """SELECT final_ledger_status, final_absence_minutes,
                                  final_shrinkage_minutes
                           FROM mart.verint_final_absence_agent_day
                           WHERE business_date='2026-08-02' AND agent_id='300'"""
                    ).fetchone(),
                    ("ABSENCE_RECORDED", 480, 480),
                )

    def test_refresh_builds_safe_attendance_gaps_and_excel(self):
        with tempfile.TemporaryDirectory() as folder:
            home = Path(folder) / "hub"
            source = Path(folder) / "source"
            (home / "config").mkdir(parents=True)
            for name in (
                "default.toml", "default_rules.toml", "default_metrics.toml",
                "default_analytics.toml", "default_reports.toml",
            ):
                shutil.copy2(REPO / "config" / name, home / "config" / name)
            shutil.copytree(REPO / "sql", home / "sql")
            make_fte(source / "FTE/FTE Count.xlsx")
            make_start_end_schedule(source / "Verint/Schedules & Activities/StartEndTimes.txt")
            make_schedule(source / "Verint/Schedules & Activities/Activities.txt")
            make_lilo(source / "Storm/LILO/AP-Historical-Report---Agent-Login 2026-08-01.csv", [
                ["Agent 100", "100", "2026-08-01 08:20:00", "2026-08-01 12:00:00"],
                ["Agent 200", "200", "", ""],
                ["Worldwide Agent", "999", "2026-08-01 08:00:00", "2026-08-01 16:00:00"],
            ])
            make_lilo(source / "Storm/LILO/AP-Historical-Report---Agent-Login 2026-08-02.csv", [
                ["Agent 300", "300", "2026-08-02 08:00:00", "2026-08-02 16:00:00"],
            ])
            make_status(source / "Storm/Agent Status/AP-Historical-Report---Agent-Status 2026-08-01.csv")
            make_logged_off_status(source / "Storm/Agent Status/AP-Historical-Report---Agent-Status logged-off 2026-08-01.csv")
            make_forecast(source / "Verint/Forecast/forecast.txt")
            make_apbe(source / "Storm/APBE ALL WFM/apbe.xlsx")
            make_apfr(source / "Storm/APFR KPI SUIVI JOUR/apfr.xlsx")
            make_calls(source / "Storm/Call by Call/AP-Historical-Report---Call-by-Call 2026-08-01 - 2026-08-02.csv")
            make_calls(source / "Storm/Call by Call/AP-Historical-Report---Call-by-Call full history.csv")

            config_file = ensure_user_config(home)
            write_source_root(config_file, source)
            config = load_config(home)
            with write_session(config) as conn:
                ingest_progress = []
                ingest = ingest_all(
                    conn, config,
                    progress=lambda current, total, label: ingest_progress.append((current, total, label)),
                )
                self.assertEqual(ingest.failed, 0)
                self.assertEqual(ingest.loaded, 10)
                self.assertEqual(ingest.scoped_out, 5)
                self.assertEqual(ingest_progress[-1][:2], (10, 10))
                self.assertTrue(any(total == 0 and "Call by Call" in label for _, total, label in ingest_progress))
                for table in ("raw.schedule_shift", "raw.lilo", "raw.agent_status", "raw.call_leg"):
                    self.assertEqual(conn.execute(f"SELECT count(*) FROM {table} WHERE agent_id='999'").fetchone()[0], 0)
                model_progress = []
                model = refresh_models(
                    conn, config, "test", date(2026, 8, 1), date(2026, 8, 2),
                    progress=lambda current, total, label: model_progress.append((current, total, label)),
                )
                self.assertEqual(model_progress[-1], (22, 22, "Models ready"))
                attendance_before_pcs = conn.execute(
                    "SELECT agent_day_key, attendance_result "
                    "FROM mart.attendance_agent_day ORDER BY agent_day_key"
                ).fetchall()
                pcs_progress = []
                pcs_only = refresh_pcs_models(
                    conn, config, "pcs-only-test",
                    date(2026, 8, 1), date(2026, 8, 2),
                    progress=lambda current, total, label: pcs_progress.append(
                        (current, total, label)
                    ),
                )
                self.assertEqual(pcs_only.pcs_rows, 2)
                self.assertEqual(pcs_progress[-1], (5, 5, "PCS data ready"))
                self.assertEqual(
                    conn.execute(
                        "SELECT agent_day_key, attendance_result "
                        "FROM mart.attendance_agent_day ORDER BY agent_day_key"
                    ).fetchall(),
                    attendance_before_pcs,
                )
                self.assertEqual(
                    dict(conn.execute(
                        "SELECT source_variant, count(*) FROM meta.source_file WHERE source_family='schedule' AND active=true GROUP BY source_variant"
                    ).fetchall()),
                    {"ACTIVITIES": 1, "START_END": 1},
                )
                saved_period = replace(config, period_start=date(2026, 8, 1), period_end=date(2026, 8, 1))
                self.assertEqual(resolve_period(conn, saved_period, None, None, True), (date(2026, 8, 1), date(2026, 8, 1)))
                self.assertEqual(resolve_period(conn, saved_period, None, None, False), (date(2026, 8, 1), date(2026, 8, 2)))
                attendance = {row[0]: row[1] for row in conn.execute("SELECT agent_day_key, attendance_result FROM mart.attendance_agent_day").fetchall()}
                self.assertEqual(attendance["20260801-200"], "No show")
                self.assertEqual(attendance["20260801-300"], "Missing actual evidence")
                self.assertEqual(
                    conn.execute(
                        "SELECT status_covered_minutes FROM mart.attendance_agent_day WHERE agent_day_key='20260801-100'"
                    ).fetchone()[0],
                    460,
                )
                late = conn.execute("SELECT gap_start, gap_end, gap_minutes FROM mart.correction_candidate WHERE agent_id='100' AND detected_issue='Late'").fetchone()
                self.assertEqual(str(late[0]), "2026-08-01 08:00:00")
                self.assertEqual(str(late[1]), "2026-08-01 08:20:00")
                self.assertEqual(late[2], 20)
                self.assertEqual(
                    conn.execute("SELECT verint_reconciliation FROM mart.correction_candidate WHERE agent_id='100' AND detected_issue='Late'").fetchone()[0],
                    "PENDING_REVIEW",
                )
                status_gap = conn.execute(
                    "SELECT gap_minutes, observed_source, verint_reconciliation FROM mart.correction_candidate WHERE agent_id='100' AND detected_issue='Mid-shift logged off'"
                ).fetchone()
                self.assertEqual(status_gap, (60, "AGENT_STATUS", "PENDING_REVIEW"))
                self.assertEqual(
                    conn.execute(
                        "SELECT uncoded_early_leave_minutes FROM mart.attendance_agent_day WHERE agent_day_key='20260801-100'"
                    ).fetchone()[0],
                    0,
                )
                self.assertGreater(
                    conn.execute(
                        "SELECT count(*) FROM mart.staffing_interval WHERE staffing_state='DATA_PARTIAL'"
                    ).fetchone()[0],
                    0,
                )
                self.assertEqual(conn.execute("SELECT count(*) FROM mart.conformance_agent_day").fetchone()[0], 0)
                self.assertEqual(model.forecast_rows, 1)
                self.assertEqual(
                    conn.execute("SELECT count(*) FROM mart.forecast_interval").fetchone()[0],
                    4,
                )
                self.assertEqual(
                    conn.execute(
                        """SELECT volume_forecast, fte_required,
                                  source_interval_minutes, source_interval_count
                           FROM mart.forecast_hour"""
                    ).fetchone(),
                    (10.0, 3.0, 15, 4),
                )
                self.assertEqual(model.intraday_rows, 0)
                self.assertEqual(model.pcs_rows, 2)
                self.assertEqual(model.absence_rows, 4)
                self.assertGreater(model.absence_event_rows, 0)
                self.assertEqual(model.service_rows, 3)
                self.assertGreater(model.metric_rows, 0)
                self.assertGreater(model.finding_rows, 0)
                self.assertEqual(
                    conn.execute("SELECT count(*) FROM meta.metric_application WHERE run_id='test'").fetchone()[0],
                    1,
                )
                service_metric = conn.execute(
                    """SELECT sum(numerator), sum(denominator)
                       FROM mart.metric_value WHERE metric_id='service_level'"""
                ).fetchone()
                service_components = conn.execute(
                    """SELECT sum(answered_within_target),
                              sum(offered-abandoned_within_target)
                       FROM mart.service_interval"""
                ).fetchone()
                self.assertAlmostEqual(
                    service_metric[0] / service_metric[1],
                    service_components[0] / service_components[1],
                )
                final_absence_max = conn.execute(
                    "SELECT max(final_absence_rate) FROM mart.verint_final_absence_agent_day"
                ).fetchone()[0]
                self.assertLessEqual(final_absence_max or 0, 1.0)
                absence_100 = conn.execute(
                    "SELECT absence_minutes, absence_rate FROM mart.absence_agent_day WHERE agent_day_key='20260801-100'"
                ).fetchone()
                self.assertEqual(absence_100[0], 0)
                self.assertEqual(absence_100[1], 0)
                self.assertEqual(
                    conn.execute(
                        "SELECT final_ledger_status FROM mart.verint_final_absence_agent_day WHERE agent_day_key='20260801-200'"
                    ).fetchone()[0],
                    "PENDING_REVIEW",
                )
                self.assertEqual(
                    conn.execute(
                        "SELECT final_ledger_status FROM mart.verint_final_absence_agent_day WHERE agent_day_key='20260801-100'"
                    ).fetchone()[0],
                    "PENDING_REVIEW",
                )
                self.assertEqual(
                    conn.execute("SELECT count(*) FROM mart.absence_event WHERE evidence_type IN ('SHIFT_EVENT','SHIFT_ASSIGNMENT')").fetchone()[0],
                    0,
                )
                self.assertEqual(
                    conn.execute("SELECT absence_minutes FROM mart.absence_agent_day WHERE agent_day_key='20260802-300'").fetchone()[0],
                    0,
                )
                self.assertEqual(
                    conn.execute("SELECT count(*) FROM mart.verint_final_exception").fetchone()[0],
                    0,
                )
                self.assertEqual(
                    conn.execute(
                        "SELECT final_ledger_status FROM mart.verint_final_absence_agent_day WHERE agent_day_key='20260801-300'"
                    ).fetchone()[0],
                    "CLEAR",
                )
                service = conn.execute(
                    "SELECT sum(answered), sum(offered), sum(handled_seconds) FROM mart.service_interval"
                ).fetchone()
                self.assertEqual(service[:2], (6, 6))
                self.assertEqual(service[2], 540)
                self.assertEqual(conn.execute("SELECT count(*) FROM raw.call_leg").fetchone()[0], 14)
                self.assertEqual(conn.execute("SELECT count(*) FROM core.clean_call_leg").fetchone()[0], 7)
                pcs = conn.execute(
                    """SELECT survey_responses, pcs_average, average_handle_seconds,
                              pcs_status_calls, pcs_participation_responses,
                              pcs_invalid_responses, pcs_status_blank_responses,
                              pcs_response_without_status, low_score_responses,
                              top_box_responses
                       FROM mart.agent_pcs_day WHERE agent_id='100'"""
                ).fetchone()
                self.assertEqual(pcs[0], 1)
                self.assertEqual(pcs[1], 5.0)
                self.assertEqual(pcs[2], 300)
                self.assertEqual(pcs[3:], (4, 4, 3, 1, 1, 0, 1))
                self.assertEqual(
                    conn.execute(
                        "SELECT metric_value FROM mart.metric_value WHERE metric_id='pcs_average'"
                    ).fetchone()[0],
                    pcs[1],
                )

                refresh_models(
                    conn, config, "pcs-today-only", date(2026, 8, 2), date(2026, 8, 2),
                    as_of=datetime(2026, 8, 2, 17, 0),
                )
                self.assertEqual(
                    str(conn.execute(
                        "SELECT min(business_date) FROM mart.agent_pcs_day"
                    ).fetchone()[0]),
                    "2026-08-01",
                )

                # A refresh during an active shift is provisional: the future
                # part of that shift must not become an early-leave or no-show
                # correction. Once the shift has ended, normal final logic
                # applies again.
                refresh_models(
                    conn, config, "in-progress", date(2026, 8, 1), date(2026, 8, 1),
                    as_of=datetime(2026, 8, 1, 12, 0),
                )
                provisional = dict(conn.execute(
                    "SELECT agent_id, attendance_result FROM mart.attendance_agent_day"
                ).fetchall())
                self.assertEqual(provisional["100"], "Late - shift in progress")
                self.assertEqual(provisional["200"], "Not seen - shift in progress")
                self.assertEqual(
                    conn.execute("SELECT call_action FROM mart.attendance_agent_day WHERE agent_id='200'").fetchone()[0],
                    "CALL_NOT_SEEN_NOW",
                )
                self.assertEqual(
                    conn.execute(
                        "SELECT final_ledger_status FROM mart.verint_final_absence_agent_day WHERE agent_id='200'"
                    ).fetchone()[0],
                    "PROVISIONAL_DAY",
                )
                self.assertEqual(
                    conn.execute(
                        """SELECT count(*) FROM mart.correction_candidate
                           WHERE detected_issue IN ('Early leave','No show')"""
                    ).fetchone()[0],
                    0,
                )
                self.assertIsNone(conn.execute(
                    "SELECT attendance_percent FROM mart.attendance_agent_day WHERE agent_id='100'"
                ).fetchone()[0])
                refresh_models(
                    conn, config, "completed-today", date(2026, 8, 1), date(2026, 8, 2),
                    as_of=datetime(2026, 8, 2, 17, 0),
                )
                self.assertEqual(
                    conn.execute(
                        "SELECT attendance_result FROM mart.attendance_agent_day WHERE agent_day_key='20260801-200'"
                    ).fetchone()[0],
                    "No show",
                )
                before_failure = conn.execute(
                    "SELECT agent_day_key, attendance_result FROM mart.attendance_agent_day ORDER BY agent_day_key"
                ).fetchall()
                with patch("wfmhub.models._build_absence", side_effect=RuntimeError("injected model failure")):
                    with self.assertRaises(RuntimeError):
                        refresh_models(conn, config, "failed", date(2026, 8, 1), date(2026, 8, 2))
                self.assertEqual(
                    conn.execute("SELECT agent_day_key, attendance_result FROM mart.attendance_agent_day ORDER BY agent_day_key").fetchall(),
                    before_failure,
                )
                report = build_report(conn, config, model.start, model.end)
                self.assertEqual(report.parent, home / "_system" / "legacy_reports")
                self.assertEqual(report.name, "Legacy Daily Operations.xlsx")
                corrections_report = build_report_pack("corrections", conn, config, model.start, model.end)
                empty_corrections_report = build_report_pack(
                    "corrections", conn, config, date(2025, 1, 1), date(2025, 1, 1),
                    output=home / "Reports" / "Attendance Review - Empty.xlsx",
                )
                empty_corrections_book = load_workbook(
                    empty_corrections_report, read_only=True, data_only=True,
                )
                try:
                    self.assertIn("REVIEW BOARD", empty_corrections_book.sheetnames)
                    audit_values = {
                        str(cell.value)
                        for row in empty_corrections_book["_AUDIT"].iter_rows()
                        for cell in row if cell.value is not None
                    }
                    self.assertIn("Classification authority", audit_values)
                finally:
                    empty_corrections_book.close()
                decisions_book = load_workbook(corrections_report)
                decisions = decisions_book["REVIEW BOARD"]
                decision_headers = {
                    cell.value: cell.column for cell in decisions[4]
                }
                approved_gap = None
                dismissed_gap = None
                for row_number in range(5, decisions.max_row + 1):
                    if not decisions.cell(
                        row_number, decision_headers["Gap ID"],
                    ).value:
                        continue
                    agent_id = str(
                        decisions.cell(row_number, decision_headers["Agent ID"]).value or ""
                    )
                    if agent_id == "200":
                        approved_gap = decisions.cell(
                            row_number, decision_headers["Gap ID"],
                        ).value
                        decisions.cell(
                            row_number, decision_headers["Decision Category"],
                            "Short sickness",
                        )
                        decisions.cell(
                            row_number, decision_headers["Decision Status"],
                            "Approved",
                        )
                        decisions.cell(
                            row_number, decision_headers["Reviewed By"], "Reviewer",
                        )
                        decisions.cell(
                            row_number, decision_headers["Reviewed Date"],
                            date(2026, 8, 3),
                        )
                    elif agent_id == "100" and dismissed_gap is None:
                        dismissed_gap = decisions.cell(
                            row_number, decision_headers["Gap ID"],
                        ).value
                        decisions.cell(
                            row_number, decision_headers["Decision Status"],
                            "Dismissed",
                        )
                decisions_book.save(corrections_report)
                decisions_book.close()
                imported = import_attendance_decisions(
                    conn, corrections_report,
                    load_rulebook(home, config.business_rules),
                )
                self.assertGreater(imported.imported, 1)
                self.assertIsNotNone(approved_gap)
                self.assertIsNotNone(dismissed_gap)
                refresh_models(
                    conn, config, "reviewed-decisions",
                    date(2026, 8, 1), date(2026, 8, 2),
                    as_of=datetime(2026, 8, 3, 17, 0),
                )
                self.assertEqual(
                    conn.execute(
                        "SELECT absence_minutes, shrinkage_minutes, unverified_minutes "
                        "FROM mart.absence_agent_day WHERE agent_day_key='20260801-200'"
                    ).fetchone(),
                    (480, 480, 0),
                )
                self.assertEqual(
                    conn.execute(
                        "SELECT confirmed_activity, validation_status "
                        "FROM core.correction_action WHERE correction_id=?",
                        [approved_gap],
                    ).fetchone(),
                    ("Short sickness", "Approved"),
                )
                self.assertEqual(
                    conn.execute(
                        "SELECT confirmed_activity, validation_status "
                        "FROM core.correction_action WHERE correction_id=?",
                        [dismissed_gap],
                    ).fetchone(),
                    (None, "Dismissed"),
                )
                self.assertGreater(
                    conn.execute(
                        "SELECT corrected_minutes FROM mart.absence_agent_day "
                        "WHERE agent_day_key='20260801-100'"
                    ).fetchone()[0],
                    0,
                )
                pcs_report = build_report_pack("quality_pcs", conn, config, model.start, model.end)
                focused_pcs_report = build_report_pack("pcs", conn, config, model.start, model.end)
                service_report = build_report_pack(
                    "service", conn, config, model.start, model.end,
                    service_profile="ford_oem_fr",
                )
                realisations_report = build_report_pack(
                    "realisations", conn, config, model.start, model.end,
                    service_profile="ford_oem_fr",
                )
                all_realisations_report = build_report_pack(
                    "realisations", conn, config, model.start, model.end,
                    home / "output" / "all-realisations.xlsx",
                )
                staffing_report = build_report_pack(
                    "staffing", conn, config, model.start, model.end,
                )
                attendance_report = build_report_pack("attendance", conn, config, model.start, model.end)
                absence_report = build_report_pack("absence", conn, config, model.start, model.end)
                analysis_report = build_analysis_workbook(
                    conn, config, "pcs", model.start, model.end,
                )
                export_progress = []
                clean_calls = export_dataset(
                    conn, config, "calls", model.start, model.end,
                    progress=lambda current, total, label: export_progress.append((current, total, label)),
                )
                self.assertEqual(clean_calls.rows, 7)
                self.assertTrue(clean_calls.manifest.exists())
                self.assertEqual(export_progress[-1], (7, 0, "Exported calls: 7 rows"))
                governed_service = export_dataset(
                    conn, config, "daily_service_lob", model.start, model.end,
                )
                governed_pcs = export_dataset(
                    conn, config, "pcs_team_day", model.start, model.end,
                )
                governed_absence = export_dataset(
                    conn, config, "final_absence_lob_month", model.start, model.end,
                )
                self.assertIn(
                    "service_level",
                    governed_service.path.read_text(encoding="utf-8-sig").splitlines()[0],
                )
                self.assertEqual(governed_pcs.rows, 1)
                self.assertGreater(governed_absence.rows, 0)
                self.assertIn(
                    "Metric catalog SHA-256:",
                    governed_service.manifest.read_text(encoding="utf-8"),
                )

                filtered_report = build_report(
                    conn, config, date(2026, 8, 1), date(2026, 8, 1),
                    home / "output" / "filtered.xlsx",
                )
                filtered = load_workbook(filtered_report, read_only=True, data_only=True)
                try:
                    filtered_dates = [row[0] for row in filtered["ATTENDANCE_CALLS"].iter_rows(min_row=5, values_only=True) if row[0] is not None]
                    self.assertGreaterEqual(len(filtered_dates), 1)
                    self.assertTrue(all(value.date() == date(2026, 8, 1) for value in filtered_dates))
                finally:
                    filtered.close()

                # Quality edits the permanent action log inside the live tracker.
                # Updating the external CSV feeds must never replace those edits.
                self.assertEqual(focused_pcs_report.name, PCS_TRACKER_FILENAME)
                tracker_book = load_workbook(focused_pcs_report)
                coaching_sheet = tracker_book["COACHING"]
                queue_headers = {
                    cell.value: cell.column for cell in coaching_sheet[4][:10]
                    if cell.value in PCS_COACHING_HEADERS
                }
                low_row = next(
                    row for row in coaching_sheet.iter_rows(min_row=5, values_only=True)
                    if row[queue_headers["Coaching Key"] - 1]
                )
                coaching_key = low_row[queue_headers["Coaching Key"] - 1]
                call_id = low_row[queue_headers["Call ID"] - 1]
                self.assertTrue(coaching_key)
                coaching_sheet["L5"] = coaching_key
                coaching_sheet["M5"] = call_id
                coaching_sheet["N5"] = "Completed"
                coaching_sheet["O5"] = "TL 1"
                coaching_sheet["P5"] = date(2026, 8, 2)
                coaching_sheet["R5"] = "Reviewed"
                tracker_book.save(focused_pcs_report)
                tracker_book.close()
                saved_tracker = focused_pcs_report.read_bytes()
                self.assertEqual(conn.execute("SELECT count(*) FROM core.pcs_coaching_action").fetchone()[0], 0)
                refreshed_snapshot = build_report_pack(
                    "pcs", conn, config, model.start, model.end,
                )
                self.assertEqual(refreshed_snapshot, focused_pcs_report)
                self.assertEqual(focused_pcs_report.read_bytes(), saved_tracker)
                self.assertFalse(list((home / "Reports").glob("PCS Paste Data - *.xlsx")))
                pcs_feed = home / "Feed" / "PCS"
                for filename, headers in (
                    ("PCS_LOB_SCORECARD_CURRENT.csv", PCS_LOB_SCORECARD_HEADERS),
                    ("PCS_AGENT_SCORECARD_CURRENT.csv", PCS_AGENT_SCORECARD_HEADERS),
                    ("PCS_DAILY_SCORECARD_CURRENT.csv", PCS_DAILY_SCORECARD_HEADERS),
                    ("PCS_RESULTS_CURRENT.csv", PCS_RESULTS_HEADERS),
                    ("PCS_COACHING_OPPORTUNITY_CURRENT.csv", PCS_COACHING_HEADERS),
                ):
                    with (pcs_feed / filename).open(encoding="utf-8-sig", newline="") as handle:
                        self.assertEqual(tuple(next(csv.reader(handle))), headers)

            self.assertTrue(report.exists())
            workbook = load_workbook(report, read_only=True, data_only=True)
            try:
                self.assertEqual(workbook.sheetnames, [
                    "DAILY_SUMMARY", "FINDINGS", "ATTENDANCE_CALLS", "STAFFING_GAPS",
                    "SERVICE_LEVEL", "DATA_QUALITY", "SOURCE_HEALTH", "SCHEDULE_SOURCES",
                    "DOMAIN_RULES", "METHODS", "PROVENANCE",
                ])
            finally:
                workbook.close()
            corrections_book = load_workbook(corrections_report, read_only=False, data_only=True)
            try:
                self.assertEqual(corrections_book.sheetnames, [
                    "CONTROL", "REVIEW BOARD", "BREAK & MEAL",
                    "DECISION LEDGER", "EVIDENCE", "DEFINITIONS", "_LOOKUPS",
                    "_AUDIT",
                ])
                self.assertEqual(corrections_book["CONTROL"]["A2"].value, "PERIOD")
                self.assertIn("01 Aug", corrections_book["CONTROL"]["C2"].value)
                self.assertEqual(
                    [corrections_book["CONTROL"][cell].value for cell in ("A5", "H5", "O5", "V5")],
                    ["REVIEW GAPS", "GAP HOURS", "OPEN DECISIONS", "MISSING EVIDENCE"],
                )
                self.assertIn(
                    "tblAttendanceReviewSummary",
                    corrections_book["CONTROL"].tables,
                )
                self.assertEqual(len(corrections_book["CONTROL"]._charts), 2)
                review = corrections_book["REVIEW BOARD"]
                review_headers = [cell.value for cell in review[4]]
                self.assertIn("Exact Start", review_headers)
                self.assertIn("Exact End", review_headers)
                self.assertIn("Decision Status", review_headers)
                self.assertIn("Band", review_headers)
                self.assertIn("08:00", review_headers)
                start_column = review_headers.index("Exact Start")
                end_column = review_headers.index("Exact End")
                exact_intervals = {
                    (row[start_column], row[end_column])
                    for row in review.iter_rows(min_row=5, values_only=True)
                    if row[start_column] is not None
                }
                self.assertIn(
                    (datetime(2026, 8, 1, 12, 0), datetime(2026, 8, 1, 13, 0)),
                    exact_intervals,
                )
                self.assertEqual(
                    corrections_book["DECISION LEDGER"].sheet_state, "hidden",
                )
                self.assertEqual(corrections_book["EVIDENCE"].sheet_state, "hidden")
                self.assertEqual(corrections_book["BREAK & MEAL"].sheet_state, "visible")
                break_meal = corrections_book["BREAK & MEAL"]
                break_meal_headers = [cell.value for cell in break_meal[4]]
                self.assertIn("Break Minutes", break_meal_headers)
                self.assertIn("Meal Minutes", break_meal_headers)
                alert_column = break_meal_headers.index("Alert")
                alerts = {
                    row[alert_column]
                    for row in break_meal.iter_rows(min_row=5, values_only=True)
                    if row[alert_column] is not None
                }
                self.assertIn("WITHIN LIMIT", alerts)
                self.assertIn("INSUFFICIENT EVIDENCE", alerts)
                self.assertEqual(corrections_book["_AUDIT"].sheet_state, "hidden")
            finally:
                corrections_book.close()
            pcs_book = load_workbook(pcs_report, read_only=True, data_only=True)
            try:
                self.assertIn("AGENT_MONTH", pcs_book.sheetnames)
                self.assertIn("RESPONSE_DETAIL", pcs_book.sheetnames)
                self.assertIn("METHODS", pcs_book.sheetnames)
            finally:
                pcs_book.close()
            attendance_book = load_workbook(attendance_report, read_only=True, data_only=True)
            try:
                self.assertIn("2026-08-01 to 2026-08-02", attendance_book["DASHBOARD"]["A2"].value)
                action_dates = [
                    row[0].date() if isinstance(row[0], datetime) else row[0]
                    for row in attendance_book["ACTIONS"].iter_rows(min_row=5, values_only=True)
                    if row[0] is not None
                ]
                self.assertIn(date(2026, 8, 1), action_dates)
                self.assertEqual(attendance_book["_AUDIT"].sheet_state, "hidden")
            finally:
                attendance_book.close()
            focused_pcs_book = load_workbook(focused_pcs_report, read_only=False, data_only=False)
            try:
                self.assertEqual(focused_pcs_report.name, PCS_TRACKER_FILENAME)
                self.assertEqual(focused_pcs_book.sheetnames, [
                    "OVERVIEW", "PERFORMANCE", "COACHING", "SETUP", "HELP",
                    "_PCS_LOB", "_PCS_AGENT", "_PCS_DAILY", "_AUDIT",
                ])
                self.assertNotIn("DATA", focused_pcs_book.sheetnames)
                self.assertEqual(len(focused_pcs_book["OVERVIEW"]._charts), 2)
                self.assertEqual(
                    [focused_pcs_book["OVERVIEW"][cell].value for cell in ("A5", "H5", "O5", "V5")],
                    ["CURRENT MTD PCS", "PARTICIPATION", "PRIOR MTD PCS", "CHANGE"],
                )
                self.assertTrue(focused_pcs_book["OVERVIEW"]["A6"].value.startswith("=IF("))
                self.assertEqual(focused_pcs_book["OVERVIEW"]["A2"].value, "DATA THROUGH")
                self.assertEqual(focused_pcs_book["OVERVIEW"]["H2"].value, "REFRESH")
                self.assertEqual(focused_pcs_book["OVERVIEW"]["O2"].value, "INTERACTION")
                self.assertEqual(focused_pcs_book["OVERVIEW"]["V2"].value, "SCOPE")
                self.assertEqual(len(focused_pcs_book["OVERVIEW"].data_validations.dataValidation), 0)
                self.assertIn("tblPcsPerformance", focused_pcs_book["PERFORMANCE"].tables)
                self.assertIn("tblCoachingQueue", focused_pcs_book["COACHING"].tables)
                self.assertIn("tblCoachingActions", focused_pcs_book["COACHING"].tables)
                self.assertEqual(
                    [focused_pcs_book["COACHING"][cell].value for cell in ("A4", "B4", "C4", "D4")],
                    ["Date", "LOB", "Team Leader", "Agent"],
                )
                coaching_headers = [
                    focused_pcs_book["COACHING"].cell(4, column).value
                    for column in range(12, 19)
                ]
                self.assertEqual(coaching_headers[:2], ["Coaching Key", "Call ID"])
                coaching_values = {
                    row[0]: row for row in focused_pcs_book["COACHING"].iter_rows(
                        min_row=5, min_col=12, max_col=18, values_only=True,
                    ) if row[0]
                }
                self.assertEqual(coaching_values[coaching_key][2], "Completed")
                self.assertEqual(coaching_values[coaching_key][3], "TL 1")
                self.assertEqual(
                    focused_pcs_book["COACHING"]["A5"].number_format,
                    "yyyy-mm-dd",
                )
                self.assertEqual(focused_pcs_book["_PCS_LOB"].sheet_state, "hidden")
                self.assertEqual(focused_pcs_book["_PCS_AGENT"].sheet_state, "hidden")
                self.assertEqual(focused_pcs_book["_PCS_DAILY"].sheet_state, "hidden")
            finally:
                focused_pcs_book.close()
            with zipfile.ZipFile(focused_pcs_report) as archive:
                self.assertNotIn("xl/connections.xml", archive.namelist())
                self.assertNotIn("xl/metadata.xml", archive.namelist())
                self.assertFalse(any(
                    name.startswith("xl/queryTables/") for name in archive.namelist()
                ))
                worksheet_xml = b"".join(
                    archive.read(name)
                    for name in archive.namelist()
                    if name.startswith("xl/worksheets/sheet") and name.endswith(".xml")
                )
                self.assertNotIn(b"<v>None</v>", worksheet_xml)
                self.assertNotIn(b"#REF!", worksheet_xml)
                self.assertNotIn(b"_xlfn", worksheet_xml)
                self.assertNotIn(b"FILTER(", worksheet_xml)
                self.assertNotIn(b"HSTACK(", worksheet_xml)
                self.assertNotIn(b"AGGREGATE", worksheet_xml)
                self.assertNotIn(b"SUMPRODUCT", worksheet_xml)
                overview_xml = archive.read("xl/worksheets/sheet1.xml")
                self.assertIn(b"<f>", overview_xml)
                self.assertNotIn(b"SUMPRODUCT", overview_xml)
                self.assertNotIn(b"AGGREGATE", overview_xml)
                self.assertNotIn(b"SUMIFS", overview_xml)
            service_book = load_workbook(service_report, read_only=False, data_only=False)
            try:
                self.assertEqual(service_report, home / "Reports" / "RTM Daily Control.xlsx")
                self.assertEqual(service_book.sheetnames, [
                    "CONTROL", "RSA NL", "RSA BE", "FORD NL", "OEM",
                    "ISSUES & DRIVERS", "DEFINITIONS", "_AUDIT",
                ])
                self.assertEqual(service_book["CONTROL"]["A1"].value, "RTM DAILY CONTROL")
                self.assertIn(
                    "PTO / Away HC",
                    [cell.value for cell in service_book["CONTROL"][35]],
                )
                self.assertEqual(
                    [service_book["CONTROL"][cell].value for cell in ("A5", "H5", "O5", "V5")],
                    ["LOBS ON TARGET", "DEMAND VARIANCE", "NO SHOW HC", "CALL NOW"],
                )
                self.assertEqual(
                    [cell.value for cell in service_book["OEM"][35]][:15],
                    [
                        "Hour", "Forecast", "Actual", "Variance", "Ford Volume",
                        "Chery Volume", "Toyota Volume", "TSL OEM", "TSL Ford",
                        "TSL Chery", "TSL Toyota", "Routed Rate", "AHT",
                        "No Show HC", "Data State",
                    ],
                )
                self.assertNotIn(
                    "Short Sickness",
                    [cell.value for cell in service_book["OEM"][35]],
                )
                self.assertIn(
                    "Issue Or Driver",
                    [cell.value for cell in service_book["ISSUES & DRIVERS"][4]],
                )
                self.assertEqual(service_book.properties.creator, "Anass ASSRI")
                self.assertEqual(len(service_book._external_links), 0)
                self.assertEqual(len(service_book["CONTROL"]._charts), 2)
                for sheet_name in ("RSA NL", "RSA BE", "FORD NL", "OEM"):
                    self.assertEqual(len(service_book[sheet_name]._charts), 2)
                    self.assertEqual(
                        [service_book[sheet_name][cell].value for cell in ("A5", "H5", "O5", "V5")],
                        ["TSL", "OFFERED", "VOLUME VARIANCE", "NO SHOW HC"],
                    )
                    table_names = set(service_book[sheet_name].tables)
                    self.assertTrue(any(name.startswith("tblRtm") for name in table_names))
                    visible_values = {
                        cell.value
                        for row in service_book[sheet_name].iter_rows()
                        for cell in row
                        if cell.value is not None
                    }
                    self.assertIn("ATTENDANCE  /  SAME-DAY OPERATIONAL LIST", visible_values)
                    self.assertIn("Agent ID", visible_values)
                    self.assertIn("PTO / Away HC", visible_values)
                    self.assertIn("Time Off", visible_values)
                self.assertEqual(service_book["DEFINITIONS"].sheet_state, "hidden")
                self.assertEqual(service_book["_AUDIT"].sheet_state, "hidden")
                self.assertFalse(any(
                    isinstance(cell.value, str) and cell.value.startswith("=")
                    for sheet in service_book.worksheets
                    for row in sheet.iter_rows()
                    for cell in row
                ))
            finally:
                service_book.close()
            absence_book = load_workbook(absence_report, read_only=False, data_only=False)
            try:
                self.assertEqual(absence_book.sheetnames, [
                    "DASHBOARD", "TEAM_VIEW", "TEAM_SUMMARY", "AGENT_RESULTS",
                    "ACTIONS", "ACTION_QUEUE", "ABSENCE_COMPONENTS",
                    "SHRINKAGE_COMPONENTS", "COMPONENT_VIEW", "ACTIVITY_DETAIL",
                    "ABSENCE_DATA", "HELP", "DEFINITIONS", "_LOOKUPS", "_AUDIT",
                ])
                self.assertEqual(
                    [absence_book["DASHBOARD"][cell].value for cell in ("A5", "H5", "O5", "V5")],
                    ["ABSENCE RATE", "SHRINKAGE RATE", "FINALIZED COVERAGE", "REVIEW CASES"],
                )
                self.assertEqual(len(absence_book["DASHBOARD"]._charts), 2)
                self.assertIn("tblAbsenceData", absence_book["ABSENCE_DATA"].tables)
                self.assertIn("tblActions", absence_book["ACTIONS"].tables)
                self.assertIn("tblActionQueue", absence_book["ACTION_QUEUE"].tables)
                absence_team_formula = absence_book["TEAM_VIEW"]["A17"].value
                self.assertIn("tblAbsenceData", getattr(absence_team_formula, "text", ""))
                component_formula = absence_book["COMPONENT_VIEW"]["A7"].value
                self.assertIn("tblActivityDetail", getattr(component_formula, "text", ""))
                self.assertEqual(absence_book["_LOOKUPS"].sheet_state, "hidden")
                self.assertEqual(absence_book["_AUDIT"].sheet_state, "hidden")
                absence_table_headers = [cell.value for cell in absence_book["ABSENCE_DATA"][4]]
            finally:
                absence_book.close()
            self.assertTrue((home / "Feed").is_dir())
            absence_feed = home / "Feed" / "Absenteeism" / "ABSENCE_AGENT_DAY_CURRENT.csv"
            self.assertTrue((home / "Feed" / "PCS").is_dir())
            self.assertTrue(absence_feed.is_file())
            with absence_feed.open("r", encoding="utf-8-sig", newline="") as handle:
                self.assertEqual(
                    next(csv.reader(handle)),
                    absence_table_headers,
                )
            realisations_book = load_workbook(realisations_report, read_only=False, data_only=False)
            try:
                self.assertEqual(realisations_book.sheetnames, [
                    "DASHBOARD", "LOB_RESULTS", "TREND", "DATA", "DEFINITIONS", "_AUDIT",
                ])
                self.assertIn("Processing Hours", [cell.value for cell in realisations_book["LOB_RESULTS"][4]])
                self.assertEqual(len(realisations_book["DASHBOARD"]._charts), 2)
                self.assertEqual(
                    [realisations_book["DASHBOARD"][cell].value for cell in ("A5", "H5", "O5", "V5")],
                    ["ACTUAL VOLUME", "FORECAST ATTAINMENT", "ROUTED RATE", "WEIGHTED AHT"],
                )
            finally:
                realisations_book.close()
            all_realisations_book = load_workbook(
                all_realisations_report, read_only=True, data_only=True,
            )
            try:
                reporting_lobs = {
                    row[4]
                    for row in all_realisations_book["LOB_RESULTS"].iter_rows(
                        min_row=5, values_only=True,
                    )
                    if row[1]
                }
                self.assertEqual(reporting_lobs, {
                    "Ford OEM France", "Ford Netherlands", "RSA Belgium",
                    "RSA Netherlands",
                })
            finally:
                all_realisations_book.close()
            staffing_book = load_workbook(staffing_report, read_only=False, data_only=True)
            try:
                self.assertIn("WEEKLY_PLAN", staffing_book.sheetnames)
                self.assertEqual(len(staffing_book["DASHBOARD"]._charts), 2)
                self.assertEqual(
                    [staffing_book["DASHBOARD"][cell].value for cell in ("A5", "H5", "O5", "V5")],
                    ["PEAK GAP FTE", "FUTURE GAP HOURS", "FORECAST COVERAGE", "PTO / AWAY IMPACT"],
                )
                staffing_dates = {
                    row[0].date() if isinstance(row[0], datetime) else row[0]
                    for row in staffing_book["INTRADAY"].iter_rows(
                        min_row=5, values_only=True,
                    )
                    if row[0] is not None
                }
                self.assertEqual(staffing_dates, {date(2026, 8, 1), date(2026, 8, 2)})
            finally:
                staffing_book.close()
            self.assertEqual(analysis_report.parent, home / "Reports" / "Analysis")
            self.assertTrue(analysis_report.is_file())
            for generated_report in (
                report, corrections_report, pcs_report, focused_pcs_report,
                service_report, realisations_report, all_realisations_report,
                staffing_report, attendance_report, absence_report, analysis_report,
            ):
                with zipfile.ZipFile(generated_report) as archive:
                    self.assertFalse(any("externalLinks" in name for name in archive.namelist()))


if __name__ == "__main__":
    unittest.main()
