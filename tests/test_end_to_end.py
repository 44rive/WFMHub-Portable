"""Fresh-install report smoke test using only synthetic evidence."""

from __future__ import annotations

import shutil
import tempfile
from datetime import date, datetime
from pathlib import Path
from xml.etree import ElementTree
from zipfile import ZipFile

from openpyxl import load_workbook

from wfmhub.config import load_config
from wfmhub.database import connect, migrate
from wfmhub.decision_products import build_realisations_workbook
from wfmhub.service_flash import build_service_flashes_workbook


ROOT = Path(__file__).resolve().parents[1]
DAY = date(2026, 9, 15)


def _home(root: Path) -> Path:
    home = root / "hub"
    (home / "config").mkdir(parents=True)
    for source in (ROOT / "config").glob("default*"):
        if source.is_file():
            shutil.copy2(source, home / "config" / source.name)
    shutil.copytree(ROOT / "sql", home / "sql")
    return home


def _check_xlsx(path: Path) -> None:
    with ZipFile(path) as package:
        assert package.testzip() is None
        for name in package.namelist():
            if name.endswith(".xml"):
                ElementTree.fromstring(package.read(name))


def test_fresh_database_builds_rtm_and_realisations_with_reconciled_hours():
    with tempfile.TemporaryDirectory() as folder:
        home = _home(Path(folder))
        config = load_config(home)
        assert migrate(config) == ["001_schema"]
        conn = connect(config)
        try:
            start = datetime(2026, 9, 15, 8)
            end = datetime(2026, 9, 15, 16)
            conn.execute(
                """INSERT INTO mart.attendance_agent_day (
                    agent_day_key,business_date,agent_id,agent_name,team_leader,
                    lob,scheduled_start,scheduled_end,assignment_type,
                    attendance_result,actual_first_seen,actual_last_seen,
                    actual_evidence,source_loaded,is_provisional,evaluation_as_of,
                    planning_overlay_minutes)
                    VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                ("2026-09-15|001", DAY, "001", "Synthetic Agent", "TL",
                 "RSA NL", start, end, "Working", "Present", start, end,
                 "AGENT_STATUS", True, False, end, 0),
            )
            conn.execute(
                """INSERT INTO mart.shift_timeline_segment (
                    segment_key,agent_day_key,business_date,agent_id,agent_name,
                    team_leader,lob,scheduled_start,scheduled_end,segment_start,
                    segment_end,segment_minutes,planned_state,actual_status,
                    actual_category,mismatch_type,is_gap,observed_source,evaluation_as_of)
                    VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                ("seg001", "2026-09-15|001", DAY, "001", "Synthetic Agent",
                 "TL", "RSA NL", start, end, start, end, 480, "WORK",
                 "Disponible", "Productive", "MATCH", False,
                 "AGENT_STATUS", end),
            )
            rtm = build_service_flashes_workbook(conn, config, DAY, DAY, home / "rtm.xlsx")
            realisations = build_realisations_workbook(conn, config, DAY, DAY, home / "realisations.xlsx")
        finally:
            conn.close()
        _check_xlsx(rtm)
        _check_xlsx(realisations)
        flash = load_workbook(rtm, read_only=True, data_only=False)
        try:
            headers = [cell.value for cell in flash["RSA NL"][35]]
            assert headers.index("Scheduled HC") < headers.index("Logged HC")
            assert headers.index("Available HC") < headers.index("BO HC")
            hour_09 = next(row for row in flash["RSA NL"].iter_rows(min_row=36, values_only=True)
                           if row[0] == "09:00")
            assert hour_09[headers.index("Scheduled HC")] == 1
            assert hour_09[headers.index("Logged HC")] == 1
            assert hour_09[headers.index("Available HC")] == 1
        finally:
            flash.close()
        report = load_workbook(realisations, read_only=True, data_only=False)
        try:
            assert "HOURS_BY_AGENT" in report.sheetnames
            headers = [cell.value for cell in report["HOURS_BY_AGENT"][4]]
            row = dict(zip(headers, (cell.value for cell in report["HOURS_BY_AGENT"][5])))
            assert row["Scheduled Hours"] == 8
            assert row["Available Hours"] == 8
            assert row["Reconciliation Delta Hours"] == 0
        finally:
            report.close()


def test_existing_realisations_catalog_accepts_new_hour_sheets():
    with tempfile.TemporaryDirectory() as folder:
        home = _home(Path(folder))
        config = load_config(home)
        old_contract = (
            'sheets = ["DASHBOARD", "LOB_RESULTS", "TREND", "DATA", '
            '"DEFINITIONS", "_AUDIT"]'
        )
        current_contract = (
            'sheets = ["DASHBOARD", "LOB_RESULTS", "TREND", "DATA", '
            '"HOURS_BY_LOB", "HOURS_BY_AGENT", "AUX_BY_TL", "AUX_BY_AGENT", '
            '"ABS_BY_LOB", "ABS_BY_AGENT", "DEFINITIONS", "_AUDIT"]'
        )
        catalog = config.report_catalog
        content = catalog.read_text(encoding="utf-8")
        assert current_contract in content
        catalog.write_text(content.replace(current_contract, old_contract), encoding="utf-8")
        migrate(config)
        conn = connect(config)
        try:
            path = build_realisations_workbook(conn, config, DAY, DAY, home / "old-catalog.xlsx")
        finally:
            conn.close()
        _check_xlsx(path)
