from __future__ import annotations

import csv
import tempfile
import unittest
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook, load_workbook

from wfmhub.ingestion import AgentScope, SourceSchemaError, parse_agent_status, parse_calls, parse_forecast, parse_fte, parse_lilo, parse_schedule


class ParserTests(unittest.TestCase):
    def test_shipped_fte_template_has_stable_contract(self):
        template = Path(__file__).resolve().parents[1] / "templates" / "FTE Count.xlsx"
        workbook = load_workbook(template, read_only=False, data_only=False)
        try:
            self.assertEqual(workbook.sheetnames[:2], ["START_HERE", "Agent"])
            headers = [cell.value for cell in workbook["Agent"][1]]
            self.assertEqual(headers, [
                "Client ID", "Status", "Name", "Team leader", "Ops Manager", "LOB",
                "Market", "Language", "Location", "City", "FTE", "End date if leaver",
            ])
            self.assertIn("tblFTEAgents", workbook["Agent"].tables)
        finally:
            workbook.close()

    def test_fte_finds_renamed_sheet_offset_header_and_aliases(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "FTE Count.xlsx"
            workbook = Workbook()
            sheet = workbook.active
            sheet.title = "August Roster"
            sheet.append(["FTE population report"])
            sheet.append([])
            sheet.append(["Agent ID", "Agent Name", "Employment Status", "TL", "Line of Business", "Site"])
            sheet.append(["00123", "Jane Agent", "Active", "Team Lead", "Quality", "Berlin"])
            workbook.save(path)

            row = parse_fte(path, "file").tables["raw.fte_agent"][0]
            self.assertEqual(row["source_row"], 4)
            self.assertEqual(row["agent_id"], "00123")
            self.assertEqual(row["agent_name"], "Jane Agent")
            self.assertEqual(row["team_leader"], "Team Lead")
            self.assertEqual(row["lob"], "Quality")
            self.assertEqual(row["location"], "Berlin")
            self.assertIsNone(row["ops_manager"])

    def test_fte_error_names_searched_sheets_and_required_headers(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "FTE Count.xlsx"
            workbook = Workbook()
            workbook.active.title = "People"
            workbook.active.append(["Something", "Else"])
            workbook.save(path)
            with self.assertRaisesRegex(SourceSchemaError, "People.*Client ID/Agent ID"):
                parse_fte(path, "file")

    def test_fte_does_not_select_support_sheet_before_agent_roster(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "FTE Count.xlsx"
            workbook = Workbook()
            support = workbook.active
            support.title = "Support"
            support.append(["Client ID", "Name", "Status", "LOB", "Market", "Location", "End Date"])
            support.append(["SUP-1", "Support Person", "Active", "Support", "Market", "Site", None])
            agent = workbook.create_sheet("Agent")
            agent.append(["Client ID", "Name"])
            agent.append(["AGT-1", "Roster Person"])
            workbook.save(path)

            rows = parse_fte(path, "file").tables["raw.fte_agent"]
            self.assertEqual([row["agent_id"] for row in rows], ["AGT-1"])

    def test_fte_rejects_multiple_id_aliases_in_roster_header(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "FTE Count.xlsx"
            workbook = Workbook()
            sheet = workbook.active
            sheet.title = "Agent"
            sheet.append(["Client ID", "Agent ID", "Name"])
            sheet.append(["CLIENT-1", "VERINT-1", "Person"])
            workbook.save(path)
            with self.assertRaisesRegex(SourceSchemaError, "multiple aliases.*agent_id"):
                parse_fte(path, "file")

    def test_fte_rejects_two_authoritative_roster_sheets(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "FTE Count.xlsx"
            workbook = Workbook()
            agent = workbook.active
            agent.title = "Agent"
            agent.append(["Client ID", "Name"])
            agent.append(["A-1", "First Person"])
            roster = workbook.create_sheet("Roster")
            roster.append(["Client ID", "Name"])
            roster.append(["A-2", "Second Person"])
            workbook.save(path)
            with self.assertRaisesRegex(SourceSchemaError, "multiple equally likely agent tables"):
                parse_fte(path, "file")

    def test_schedule_uses_quote_aware_tsv(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "schedule.txt"
            with path.open("w", encoding="cp1252", newline="") as handle:
                writer = csv.writer(handle, delimiter="\t")
                writer.writerow(["Name", "Data Source IDs", "Scheduling Period", "Before Overtime", "After Overtime", "Shift Assignment", "Shift Events"])
                writer.writerow(["08/01/2026", "", "", "", "", "", ""])
                writer.writerow(["Doe,\tJane", "123456", "August", "", "", "Off", ""])
            result = parse_schedule(path, "file")
            self.assertEqual(len(result.tables["raw.schedule_shift"]), 1)
            self.assertEqual(result.tables["raw.schedule_shift"][0]["agent_id"], "123456")
            self.assertEqual(result.tables["raw.schedule_shift"][0]["agent_name"], "Doe,\tJane")

    def test_start_end_times_wide_extract_normalizes_every_date_column(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "VERINT_01082026_03082026_StartEndTimes.txt"
            with path.open("w", encoding="cp1252", newline="") as handle:
                writer = csv.writer(handle, delimiter="\t")
                writer.writerow(["Name", "Data Source IDs", "08/01/2026", "08/02/2026", "08/03/2026"])
                writer.writerow([
                    "Agent One", "123", ".AP | Short Sickness 08/01/2026 8:30 AM-08/01/2026 6:00 PM",
                    "Off", ".AP | Front-office 08/03/2026 9:00 AM-08/03/2026 6:30 PM",
                ])
            result = parse_schedule(path, "file")
            shifts = result.tables["raw.schedule_shift"]
            self.assertEqual(len(shifts), 3)
            self.assertEqual([str(row["schedule_date"]) for row in shifts], ["2026-08-01", "2026-08-02", "2026-08-03"])
            self.assertEqual(shifts[0]["agent_id"], "123")
            self.assertEqual(shifts[0]["assignment_type"], "Planned absence")
            self.assertEqual(shifts[1]["assignment_type"], "Off")
            self.assertEqual(result.tables["raw.schedule_event"], [])

    def test_lilo_preserves_and_adjusts_overnight_boundary(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "AP-Historical-Report---Agent-Login 2026-08-01.csv"
            path.write_text(
                "[Agent],[Agent ID],[First Log-on Time],[Last Log-off Time]\n"
                "Jane,123,2026-08-01 21:55:00,2026-08-01 07:00:00\n",
                encoding="utf-8-sig",
            )
            row = parse_lilo(path, "file").tables["raw.lilo"][0]
            self.assertEqual(row["raw_last_logout"], datetime(2026, 8, 1, 7, 0))
            self.assertEqual(row["last_logout"], datetime(2026, 8, 2, 7, 0))
            self.assertTrue(row["overnight_adjusted"])

    def test_lilo_range_file_uses_row_date_and_preserves_blank_no_show(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "AP-Historical-Report---Agent-Login 2026-08-01 - 2026-08-02.csv"
            path.write_text(
                "[Date],[Agent],[Agent ID],[First Log-on Time],[Last Log-off Time]\n"
                "2026-08-01,Jane,123,,,\n"
                "2026-08-02,Jane,123,2026-08-02 08:00:00,2026-08-02 16:00:00\n",
                encoding="utf-8-sig",
            )
            rows = parse_lilo(path, "file").tables["raw.lilo"]
            self.assertEqual([str(row["extract_date"]) for row in rows], ["2026-08-01", "2026-08-02"])
            self.assertIsNone(rows[0]["first_login"])

    def test_lilo_multiday_blank_row_without_date_is_rejected_not_invented(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "AP-Historical-Report---Agent-Login 2026-08-01 - 2026-08-02.csv"
            path.write_text(
                "[Agent],[Agent ID],[First Log-on Time],[Last Log-off Time]\nJane,123,,\n",
                encoding="utf-8-sig",
            )
            result = parse_lilo(path, "file")
            self.assertEqual(result.tables["raw.lilo"], [])
            self.assertRegex(result.rejected[0], "no row date")

    def test_status_duration_can_exceed_24_hours(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "AP-Historical-Report---Agent-Status 2026-08-24.csv"
            path.write_text(
                "[Serial Number],[Status],[Status Start Date and Time],[Agent],[Agent ID],[Status Duration],[Queue]\n"
                "one,Logged Off,8/24/2026 0:00,Jane,123,37:59:56,\n",
                encoding="utf-8-sig",
            )
            row = parse_agent_status(path, "file").tables["raw.agent_status"][0]
            self.assertEqual(str(row["extract_date"]), "2026-08-24")
            self.assertEqual(row["status_end"], datetime(2026, 8, 25, 13, 59, 56))

    def test_status_range_filename_uses_each_rows_timestamp_date(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "AP-Historical-Report---Agent-Status 2026-08-24 - 2026-08-25.csv"
            path.write_text(
                "[Serial Number],[Status],[Status Start Date and Time],[Agent],[Agent ID],[Status Duration],[Queue]\n"
                "one,Available,8/24/2026 23:55,Jane,123,0:05:00,Main\n"
                "two,Available,8/25/2026 0:05,Jane,123,0:05:00,Main\n",
                encoding="utf-8-sig",
            )
            rows = parse_agent_status(path, "file").tables["raw.agent_status"]
            self.assertEqual([str(row["extract_date"]) for row in rows], ["2026-08-24", "2026-08-25"])

    def test_status_filename_does_not_need_a_date_when_rows_have_timestamps(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "Agent Status full history.csv"
            path.write_text(
                "[Serial Number],[Status],[Status Start Date and Time],[Agent],[Agent ID],[Status Duration],[Queue]\n"
                "one,Available,8/24/2026 12:00,Jane,123,0:05:00,Main\n",
                encoding="utf-8-sig",
            )
            row = parse_agent_status(path, "file").tables["raw.agent_status"][0]
            self.assertEqual(str(row["extract_date"]), "2026-08-24")

    def test_forecast_discovers_header_and_scales_service_level(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "forecast.txt"
            path.write_text(
                "DATE_TIME_FORMAT\nMM/DD/YYYY hh:mm A\n"
                "Queue Name\tDate\tTime\tTime Interval\tVolume (Absolute For)\tAbandons (Absolute For)\tService Level (Absolute For)\tService Level (Absolute Req)\tActivity Handling Time (Absolute For)\tHeadcount Staffing (Absolute For)\tNet Staffing (Absolute For)\tFull Time Equivalents (Absolute For)\tFull Time Equivalents (Absolute Req)\n"
                "All\t08/01/2026\t12:00 AM\t1:00\t10\t1\t79\t80\t200\t3\t-1\t2\t3\n",
                encoding="cp1252",
            )
            row = parse_forecast(path, "file").tables["raw.forecast_interval"][0]
            self.assertEqual(row["interval_minutes"], 60)
            self.assertEqual(row["sl_forecast"], 0.79)
            self.assertEqual(row["net_staffing_forecast"], -1)

    def test_call_by_call_uses_row_dates_scores_and_fte_scope(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "Call by Call 2026-08-01 - 2026-08-31.csv"
            headers = [
                "[Call Date/Time]", "[Call End Date/Time]", "[Call ID]",
                "[Call Reference Number]", "[Agent ID]", "[Agent]",
                "[Talk Time]", "[Hold Time]", "[Total Wrap Time]",
                "[Call Direction]", "[PostCallSurveyMode]", "[Question 1]",
                "[Question 2]", "[Question 3]",
            ]
            with path.open("w", encoding="utf-8-sig", newline="") as handle:
                writer = csv.DictWriter(handle, fieldnames=headers)
                writer.writeheader()
                writer.writerow({
                    "[Call Date/Time]": "8/1/2026 8:00", "[Call End Date/Time]": "8/1/2026 8:05",
                    "[Call ID]": "call-1", "[Call Reference Number]": "ref-1",
                    "[Agent ID]": "123", "[Agent]": "Jane", "[Talk Time]": "0:04:00",
                    "[Hold Time]": "0:00:30", "[Total Wrap Time]": "0:00:30",
                    "[Call Direction]": "I", "[PostCallSurveyMode]": "2",
                    "[Question 1]": "5", "[Question 2]": "4", "[Question 3]": "Helpful",
                })
                writer.writerow({
                    "[Call Date/Time]": "8/2/2026 9:00", "[Call End Date/Time]": "8/2/2026 9:01",
                    "[Call ID]": "world", "[Call Reference Number]": "world",
                    "[Agent ID]": "999", "[Agent]": "Worldwide", "[Talk Time]": "0:01:00",
                    "[Hold Time]": "0:00:00", "[Total Wrap Time]": "0:00:00",
                })
            scope = AgentScope(frozenset({"123"}), {"jane": "123"}, "scope")
            result = parse_calls(path, "file", scope)
            self.assertEqual(result.scoped_out, 1)
            self.assertEqual(len(result.tables["raw.call_leg"]), 1)
            row = result.tables["raw.call_leg"][0]
            self.assertEqual(str(row["business_date"]), "2026-08-01")
            self.assertEqual(row["question_1_score"], 5)
            self.assertEqual(row["question_2_score"], 4)
            self.assertEqual(row["talk_seconds"], 240)


if __name__ == "__main__":
    unittest.main()
