from __future__ import annotations

import shutil
import tempfile
import unittest
from xml.etree import ElementTree
from datetime import date
from pathlib import Path
from unittest.mock import patch
from zipfile import ZipFile

from openpyxl import Workbook, load_workbook

from wfmhub.bonus import DEFAULT_BONUS_KPIS, import_bonus_matrix, build_bonus_performance_workbook
from wfmhub.config import load_config
from wfmhub.database import connect, migrate
from wfmhub.on_demand_analysis import build_analysis_workbook
from wfmhub.service_profiles import load_service_profiles
from wfmhub.governance_catalog import build_kpi_catalog
from wfmhub.report_packs import report_current_path


REPO = Path(__file__).resolve().parents[1]


def _home(root: Path) -> Path:
    home = root / "hub"
    (home / "config").mkdir(parents=True)
    for source in (REPO / "config").glob("default*"):
        if source.is_file():
            shutil.copy2(source, home / "config" / source.name)
    shutil.copytree(REPO / "sql", home / "sql")
    return home


def _bonus_source(path: Path) -> None:
    workbook = Workbook()
    raw = workbook.active
    raw.title = "Raw_Data"
    raw.append([
        "Agent ID", "Agent Name", "Population", "Period", "AHT", "Productivity",
        "PCS Score", "PCS % Participation", "QM", "Abs%", "VOC Detractor Count",
        "Currency", "Monthly Fixed Salary", "Target Bonus Rate",
        "Reference Bonus Override", "Eligible Days", "Scheduled Days",
        "Employment Status", "Data Status", "Notes", "Team Lead", "Ops Manager",
    ])
    raw.append([
        "007", "Agent Seven", "OEM", "2026-08", 440, 8, 4.5, .55, .95,
        .02, 0, "MAD", None, None, 2000, 31, 31, "Active", "VALIDATED", None,
        "TL Seven", "Ops Seven",
    ])
    rules = workbook.create_sheet("KPI_Config")
    rules.append([
        "Population", "KPI", "Direction", "Tier 1 Bonus %", "Tier 1 Target",
        "Tier 2 Bonus %", "Tier 2 Target",
    ])
    for row in (
        ("OEM", "AHT", "L", .15, 450, 0, 0),
        ("OEM", "Productivity", "H", .15, 7, 0, 0),
        ("OEM", "PCS Score", "H", .20, 4.3, .15, 4.1),
        ("OEM", "PCS % (Participation)", "H", .15, .5, .1, .4),
        ("OEM", "QM", "H", .2, .9, .15, .8),
        ("OEM", "Abs%", "L", .15, .05, 0, 0),
        ("OEM", "Extra Bonus (PCS Score)", "H", .3, 4.5, .2, 4.3),
    ):
        rules.append(row)
    workbook.save(path)
    workbook.close()


class BonusImportTests(unittest.TestCase):
    def test_bonus_formula_upgrade_preserves_entered_raw_and_policy_values(self):
        with tempfile.TemporaryDirectory() as folder:
            home = _home(Path(folder))
            config = load_config(home)
            migrate(config)
            path = report_current_path(config, "bonus")
            conn = connect(config)
            try:
                with patch("wfmhub.bonus.BONUS_TRACKER_VERSION", "1.1.0"):
                    build_bonus_performance_workbook(
                        conn, config, date(2026, 9, 1), date(2026, 9, 30),
                    )
                old = load_workbook(path)
                old["Raw_Data"]["A5"] = "007"
                old["Raw_Data"]["B5"] = "Synthetic Agent"
                old["Raw_Data"]["C5"] = "OEM FR"
                old["Policy_Decisions"]["B5"] = "Percentage points"
                old.save(path)
                old.close()
                build_bonus_performance_workbook(
                    conn, config, date(2026, 9, 1), date(2026, 9, 30),
                )
                current = load_workbook(path, read_only=True)
                try:
                    self.assertEqual(current["Raw_Data"]["A5"].value, "007")
                    self.assertEqual(current["Raw_Data"]["C5"].value, "OEM FR")
                    self.assertEqual(current["Policy_Decisions"]["B5"].value, "Percentage points")
                finally:
                    current.close()
            finally:
                conn.close()

    def test_blank_permanent_bonus_template_is_formula_ready_and_keeps_manual_inputs(self):
        with tempfile.TemporaryDirectory() as folder:
            home = _home(Path(folder))
            config = load_config(home)
            migrate(config)
            path = report_current_path(config, "bonus")
            conn = connect(config)
            try:
                build_bonus_performance_workbook(
                    conn, config, date(2026, 9, 1), date(2026, 9, 30),
                )
                workbook = load_workbook(path, read_only=False, data_only=False)
                try:
                    self.assertEqual(workbook["KPI_Config"].max_row, 46)
                    self.assertEqual(len(DEFAULT_BONUS_KPIS), 42)
                    self.assertEqual(
                        {row[0] for row in DEFAULT_BONUS_KPIS},
                        {"Ford GER", "Ford Dutch", "OEM FR", "RSA FR", "RSA NL", "RSA VL"},
                    )
                    self.assertEqual(workbook["Raw_Data"].max_row, 1004)
                    self.assertIsNotNone(workbook["Results"]["Y1004"].value)
                    self.assertTrue(str(workbook["Dashboard"]["A6"].value).startswith("=SUMIFS"))
                    self.assertEqual(len(workbook["Dashboard"]._charts), 2)
                    workbook["Raw_Data"]["A5"] = "007"
                    workbook["Raw_Data"]["B5"] = "Synthetic Agent"
                    workbook["Raw_Data"]["C5"] = "OEM"
                    workbook["Raw_Data"]["D5"] = "2026-09"
                    workbook.save(path)
                finally:
                    workbook.close()
                before = path.read_bytes()
                build_bonus_performance_workbook(
                    conn, config, date(2026, 9, 1), date(2026, 9, 30),
                )
                self.assertEqual(path.read_bytes(), before)
            finally:
                conn.close()

    def test_bonus_import_is_read_only_idempotent_and_preserves_text_id(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            home = _home(root)
            source = root / "Bonus_Matrix_v1.2.xlsx"
            _bonus_source(source)
            before = source.read_bytes()
            config = load_config(home)
            migrate(config)
            conn = connect(config)
            try:
                first = import_bonus_matrix(conn, source)
                second = import_bonus_matrix(conn, source)
                self.assertFalse(first.unchanged)
                self.assertTrue(second.unchanged)
                self.assertEqual(first.period, "2026-08")
                self.assertEqual(
                    conn.execute("SELECT agent_id FROM mart.bonus_agent_month").fetchone()[0],
                    "007",
                )
                self.assertEqual(
                    conn.execute(
                        "SELECT team_leader, ops_manager FROM mart.bonus_agent_month"
                    ).fetchone(),
                    ("TL Seven", "Ops Seven"),
                )
                self.assertEqual(conn.execute("SELECT count(*) FROM raw.bonus_import").fetchone()[0], 1)
                self.assertEqual(conn.execute("SELECT count(*) FROM mart.bonus_agent_month").fetchone()[0], 1)
            finally:
                conn.close()
            self.assertEqual(source.read_bytes(), before)

    def test_bonus_report_keeps_v12_contract_and_wfm_visuals(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            home = _home(root)
            source = root / "Bonus_Matrix_v1.2.xlsx"
            _bonus_source(source)
            config = load_config(home)
            migrate(config)
            conn = connect(config)
            try:
                import_bonus_matrix(conn, source)
                path = build_bonus_performance_workbook(
                    conn, config, date(2026, 8, 1), date(2026, 8, 31),
                    home / "Bonus Management.xlsx",
                )
            finally:
                conn.close()
            workbook = load_workbook(path, read_only=False, data_only=False, keep_links=False)
            try:
                self.assertEqual(workbook.sheetnames, [
                    "Policy_Decisions", "Control_Checks", "KPI_Config", "Raw_Data",
                    "Results", "KPI_Analysis", "Team_Lead_Analysis", "Dashboard",
                ])
                self.assertEqual(workbook.active.title, "Dashboard")
                self.assertEqual(len(workbook["Dashboard"]._charts), 2)
                self.assertTrue(workbook["Dashboard"]["AJ2"].value.startswith('=KPI_Config!B5&'))
                self.assertEqual(
                    [workbook["Dashboard"][cell].value for cell in ("A5", "H5", "O5", "V5")],
                    ["TOTAL PAYOUT", "PAID AGENTS", "AVG PAID PAYOUT", "REVIEW ITEMS"],
                )
                self.assertEqual(len(workbook._external_links), 0)
                self.assertIn("tblKpiConfig", workbook["KPI_Config"].tables)
                self.assertIn("Tier 1 Bonus %", [cell.value for cell in workbook["KPI_Config"][4]])
                self.assertIn("Team Lead", [cell.value for cell in workbook["Raw_Data"][4]])
                self.assertIn("Final Payout", [cell.value for cell in workbook["Results"][4]])
            finally:
                workbook.close()
            namespace = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
            with ZipFile(path) as archive:
                self.assertIsNone(archive.testzip())
                for filename in archive.namelist():
                    if not filename.startswith("xl/worksheets/sheet") or not filename.endswith(".xml"):
                        continue
                    root = ElementTree.fromstring(archive.read(filename))
                    for formula in root.iter(namespace + "f"):
                        text = formula.text or ""
                        for unsupported in ("XLOOKUP(", "LET(", "SORT(", "UNIQUE(", "FILTER(", " / ="):
                            self.assertNotIn(unsupported, text, (filename, text))


class ExcelTemplateTests(unittest.TestCase):
    def test_governance_workbook_exposes_service_and_mapping_controls(self):
        with tempfile.TemporaryDirectory() as folder:
            home = _home(Path(folder))
            config = load_config(home)
            path = build_kpi_catalog(config)
            workbook = load_workbook(path, read_only=True)
            try:
                self.assertEqual(
                    workbook.sheetnames,
                    [
                        "METRIC_METHODS", "ACTIVITY_RULES", "STATUS_REFERENCE", "ANALYTICS_RULES",
                        "REPORT_CONTRACTS", "SERVICE_PROFILES", "SERVICE_GROUPS",
                        "QUEUE_MAPPING",
                    ],
                )
                self.assertEqual(len(workbook._external_links), 0)
            finally:
                workbook.close()

    def test_default_ford_profile_uses_storm_business_reference(self):
        catalog = load_service_profiles(REPO, REPO / "config" / "default_service_profiles.toml")
        ford = catalog.select("ford_oem_fr", date(2026, 8, 25))
        self.assertEqual(ford.service_level_metric, "service_level")
        self.assertEqual(ford.availability_metric, "service_availability_business")
        self.assertEqual(ford.source_systems, ("CALL_BY_CALL",))
        self.assertEqual(ford.staffing_lobs, ("OEM FR",))
        self.assertEqual(
            [group.label for group in ford.groups],
            ["Ford", "Toyota", "Chery"],
        )
        self.assertEqual(ford.flash_queues, (
            "APFR_PAR_RSA_CSTRUCTR_FORD_ASSISTANCE_FR",
            "APFR_PAR_RSA_CSTRUCTR_TOYOTA-LEXUS_FR",
            "APFR_PAR_RSA_CHERY_ASSISTANCE_FR",
        ))

    def test_service_profile_catalog_preserves_existing_user_file(self):
        with tempfile.TemporaryDirectory() as folder:
            home = Path(folder)
            config = home / "config"
            config.mkdir()
            default_text = (
                REPO / "config" / "default_service_profiles.toml"
            ).read_text(encoding="utf-8")
            (config / "default_service_profiles.toml").write_text(
                default_text, encoding="utf-8",
            )
            previous_text = default_text.replace(
                'description = "Effective-dated service Flash profiles for RSA Netherlands, RSA Belgium, Ford Netherlands, and Ford OEM France."',
                'description = "Locally reviewed service profiles."',
                1,
            )
            target = config / "service_profiles.toml"
            target.write_text(previous_text, encoding="utf-8")

            catalog = load_service_profiles(home, target)

            self.assertEqual(catalog.version, "1.0.0")
            self.assertEqual(
                catalog.select("ford_oem_fr", date(2026, 9, 1)).flash_total_groups,
                ("Ford", "Toyota", "Chery"),
            )
            self.assertEqual(
                len(catalog.select("ford_oem_fr", date(2026, 9, 1)).flash_queues),
                3,
            )
            self.assertIn("Locally reviewed service profiles", target.read_text(encoding="utf-8"))

    def test_forecast_analysis_uses_materialized_forecast_and_service_tables(self):
        with tempfile.TemporaryDirectory() as folder:
            home = _home(Path(folder))
            config = load_config(home)
            migrate(config)
            conn = connect(config)
            try:
                path = build_analysis_workbook(
                    conn, config, "forecast", date(2026, 8, 1), date(2026, 8, 31),
                    "target", home / "output" / "forecast_analysis.xlsx",
                )
            finally:
                conn.close()
            self.assertTrue(path.exists())


if __name__ == "__main__":
    unittest.main()
