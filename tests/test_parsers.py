from __future__ import annotations

import csv
import tempfile
import unittest
from datetime import datetime
from pathlib import Path

from wfmhub.ingestion import parse_agent_status, parse_forecast, parse_lilo, parse_schedule


class ParserTests(unittest.TestCase):
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

    def test_status_duration_can_exceed_24_hours(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "AP-Historical-Report---Agent-Status 2026-08-24.csv"
            path.write_text(
                "[Serial Number],[Status],[Status Start Date and Time],[Agent],[Agent ID],[Status Duration],[Queue]\n"
                "one,Logged Off,8/24/2026 0:00,Jane,123,37:59:56,\n",
                encoding="utf-8-sig",
            )
            row = parse_agent_status(path, "file").tables["raw.agent_status"][0]
            self.assertEqual(row["status_end"], datetime(2026, 8, 25, 13, 59, 56))

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


if __name__ == "__main__":
    unittest.main()
