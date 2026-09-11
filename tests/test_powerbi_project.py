from __future__ import annotations

import json
import re
import shutil
import tempfile
import unittest
from pathlib import Path

from wfmhub.config import load_config
from wfmhub.powerbi_project import PROJECT_NAME, _parameterize, install_powerbi_project


REPO = Path(__file__).resolve().parents[1]
TEMPLATE = REPO / "templates" / "powerbi" / PROJECT_NAME


class PowerBIProjectTests(unittest.TestCase):
    def test_parameterizer_treats_windows_backslashes_as_literal_text(self):
        with tempfile.TemporaryDirectory() as folder:
            project = Path(folder) / "project"
            shutil.copytree(TEMPLATE, project)
            windows_like_home = Path(folder) / r"C:\Users\JMNKHSP\WFMHub"

            _parameterize(project, windows_like_home)

            expressions = (
                project / f"{PROJECT_NAME}.SemanticModel" /
                "definition" / "expressions.tmdl"
            )
            self.assertIn(
                str(windows_like_home.resolve()),
                expressions.read_text(encoding="utf-8"),
            )

    def test_shipped_pbip_has_seven_valid_pages_and_governed_sources(self):
        project = TEMPLATE / f"{PROJECT_NAME}.pbip"
        self.assertTrue(project.is_file())
        for path in TEMPLATE.rglob("*.json"):
            with self.subTest(path=path.relative_to(TEMPLATE)):
                json.loads(path.read_text(encoding="utf-8"))
        pages_file = (
            TEMPLATE / f"{PROJECT_NAME}.Report" / "definition" / "pages" / "pages.json"
        )
        pages = json.loads(pages_file.read_text(encoding="utf-8"))
        self.assertEqual(len(pages["pageOrder"]), 7)
        names = []
        for page_id in pages["pageOrder"]:
            page_root = pages_file.parent / page_id
            page = json.loads((page_root / "page.json").read_text(encoding="utf-8"))
            names.append(page["displayName"])
            self.assertEqual((page["width"], page["height"]), (1280, 720))
            visuals = list((page_root / "visuals").glob("*/visual.json"))
            self.assertEqual(len(visuals), 20)
            self.assertEqual(len({path.parent.name for path in visuals}), 20)
        self.assertEqual(
            names,
            [
                "Executive Overview", "Service & Forecast", "Attendance Control",
                "PCS Performance & Coaching", "Staffing & Capacity",
                "Absence & Shrinkage", "Data Quality & Governance",
            ],
        )
        expressions = TEMPLATE / f"{PROJECT_NAME}.SemanticModel" / "definition"
        tmdl = "\n".join(path.read_text(encoding="utf-8") for path in expressions.rglob("*.tmdl"))
        self.assertIn('expression HubRoot = "C:\\WFMHub"', tmdl)
        self.assertIn('HubRoot & "\\Feed\\PowerBI\\', tmdl)
        for forbidden in ("sqlite", "raw\\", "source_root", "duckdb"):
            self.assertNotIn(forbidden, tmdl.lower())

    def test_every_visual_binding_exists_in_the_semantic_model(self):
        model_root = TEMPLATE / f"{PROJECT_NAME}.SemanticModel" / "definition" / "tables"
        fields: dict[str, set[str]] = {}
        for path in model_root.glob("*.tmdl"):
            text = path.read_text(encoding="utf-8")
            table_match = re.search(r"^table '([^']+)'", text, re.MULTILINE)
            self.assertIsNotNone(table_match, path)
            fields[table_match.group(1)] = set(
                re.findall(r"^\t(?:column|measure) '([^']+)'", text, re.MULTILINE)
            )
        report_root = TEMPLATE / f"{PROJECT_NAME}.Report" / "definition" / "pages"
        for path in report_root.glob("*/visuals/*/visual.json"):
            visual = json.loads(path.read_text(encoding="utf-8"))
            for binding_type in ("Column", "Measure"):
                for match in self._bindings(visual, binding_type):
                    with self.subTest(visual=path.parent.name, binding=match):
                        self.assertIn(match[0], fields)
                        self.assertIn(match[1], fields[match[0]])

    def test_no_table_has_a_measure_and_column_with_the_same_name(self):
        model_root = TEMPLATE / f"{PROJECT_NAME}.SemanticModel" / "definition" / "tables"
        for path in model_root.glob("*.tmdl"):
            text = path.read_text(encoding="utf-8")
            measures = {
                match.casefold()
                for match in re.findall(r"^\tmeasure '([^']+)' =", text, re.MULTILINE)
            }
            columns = {
                match.casefold()
                for match in re.findall(r"^\tcolumn '([^']+)'$", text, re.MULTILINE)
            }
            with self.subTest(table=path.stem):
                self.assertFalse(measures & columns)

    @staticmethod
    def _bindings(value, binding_type: str):
        if isinstance(value, dict):
            if binding_type in value:
                binding = value[binding_type]
                source = binding.get("Expression", {}).get("SourceRef", {}).get("Entity")
                prop = binding.get("Property")
                if source and prop:
                    yield source, prop
            for child in value.values():
                yield from PowerBIProjectTests._bindings(child, binding_type)
        elif isinstance(value, list):
            for child in value:
                yield from PowerBIProjectTests._bindings(child, binding_type)

    def test_installer_parameterizes_preserves_and_versions_the_project(self):
        with tempfile.TemporaryDirectory() as folder:
            home = Path(folder) / "WFMHub"
            (home / "config").mkdir(parents=True)
            for source in (REPO / "config").glob("default*"):
                if source.is_file():
                    shutil.copy2(source, home / "config" / source.name)
            shutil.copytree(TEMPLATE, home / "templates" / "powerbi" / PROJECT_NAME)
            manifest = home / "Feed" / "PowerBI" / "POWERBI_MANIFEST_CURRENT.csv"
            manifest.parent.mkdir(parents=True)
            manifest.write_text("receipt\n", encoding="utf-8")
            config = load_config(home)
            first = install_powerbi_project(config)
            self.assertTrue(first.installed)
            expressions = (
                first.project.parent / f"{PROJECT_NAME}.SemanticModel" /
                "definition" / "expressions.tmdl"
            )
            self.assertIn(str(home.resolve()), expressions.read_text(encoding="utf-8"))
            marker = first.project.parent / "USER_MARKER.txt"
            marker.write_text("preserve", encoding="utf-8")
            second = install_powerbi_project(config)
            self.assertFalse(second.installed)
            self.assertFalse(second.upgraded)
            self.assertEqual(marker.read_text(encoding="utf-8"), "preserve")

            source_root = home / "templates" / "powerbi" / PROJECT_NAME
            version_file = source_root / "PROJECT_VERSION.txt"
            next_version = int(version_file.read_text(encoding="utf-8").strip()) + 1
            version_file.write_text(f"{next_version}\n", encoding="utf-8")
            third = install_powerbi_project(config)
            self.assertTrue(third.upgraded)
            self.assertIsNotNone(third.archived)
            self.assertEqual((third.archived / "USER_MARKER.txt").read_text(encoding="utf-8"), "preserve")
            self.assertFalse((third.project.parent / "USER_MARKER.txt").exists())


if __name__ == "__main__":
    unittest.main()
