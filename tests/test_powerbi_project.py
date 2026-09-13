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

    def test_shipped_pbip_has_five_valid_wfm_cycle_pages_and_governed_sources(self):
        project = TEMPLATE / f"{PROJECT_NAME}.pbip"
        self.assertTrue(project.is_file())
        for path in TEMPLATE.rglob("*.json"):
            with self.subTest(path=path.relative_to(TEMPLATE)):
                json.loads(path.read_text(encoding="utf-8"))
        pages_file = (
            TEMPLATE / f"{PROJECT_NAME}.Report" / "definition" / "pages" / "pages.json"
        )
        pages = json.loads(pages_file.read_text(encoding="utf-8"))
        self.assertEqual(len(pages["pageOrder"]), 5)
        names = []
        expected_slicers = {
            "Today's Control": 5,
            "Staff Preparation": 4,
            "Intraday Service": 4,
            "Schedule Review": 5,
            "Historical Review": 4,
        }
        for page_id in pages["pageOrder"]:
            page_root = pages_file.parent / page_id
            page = json.loads((page_root / "page.json").read_text(encoding="utf-8"))
            names.append(page["displayName"])
            self.assertEqual((page["width"], page["height"]), (1680, 945))
            visuals = list((page_root / "visuals").glob("*/visual.json"))
            self.assertGreaterEqual(len(visuals), 30)
            self.assertEqual(len({path.parent.name for path in visuals}), len(visuals))
            visual_types = [
                json.loads(path.read_text(encoding="utf-8"))["visual"]["visualType"]
                for path in visuals
            ]
            self.assertEqual(visual_types.count("slicer"), expected_slicers[page["displayName"]])
            self.assertEqual(
                visual_types.count("cardVisual"),
                8 if page["displayName"] == "Historical Review" else 4,
            )
            self.assertIn("pageNavigator", visual_types)
            self.assertGreaterEqual(visual_types.count("shape"), 11)
            self.assertNotIn("stackedBarChart", visual_types)
        self.assertEqual(
            names,
            [
                "Today's Control", "Staff Preparation", "Intraday Service",
                "Schedule Review", "Historical Review",
            ],
        )
        expressions = TEMPLATE / f"{PROJECT_NAME}.SemanticModel" / "definition"
        tmdl = "\n".join(path.read_text(encoding="utf-8") for path in expressions.rglob("*.tmdl"))
        self.assertIn('expression HubRoot = "C:\\WFMHub"', tmdl)
        self.assertIn('HubRoot & "\\Feed\\PowerBI\\', tmdl)
        for forbidden in ("sqlite", "raw\\", "source_root", "duckdb"):
            self.assertNotIn(forbidden, tmdl.lower())
        self.assertNotIn("FactPCS", tmdl)
        self.assertNotIn("table 'PCS'", tmdl)
        self.assertIn("FactService15Min.csv", tmdl)
        self.assertIn("FactScheduleIntegrity.csv", tmdl)
        self.assertIn("FactForecastInterval.csv", tmdl)
        self.assertIn("FactOperationalAction.csv", tmdl)
        self.assertIn("FactBreakMealControl.csv", tmdl)

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

    def test_active_relationships_have_no_ambiguous_filter_paths(self):
        relationships = (
            TEMPLATE / f"{PROJECT_NAME}.SemanticModel" /
            "definition" / "relationships.tmdl"
        ).read_text(encoding="utf-8")
        graph: dict[str, list[str]] = {}
        for block in re.split(r"(?m)^relationship ", relationships)[1:]:
            if re.search(r"(?m)^\tisActive:\s*false\s*$", block):
                continue
            from_match = re.search(r"(?m)^\tfromColumn: '([^']+)'", block)
            to_match = re.search(r"(?m)^\ttoColumn: '([^']+)'", block)
            self.assertIsNotNone(from_match)
            self.assertIsNotNone(to_match)
            # The one-side table filters the many-side table by default.
            graph.setdefault(to_match.group(1), []).append(from_match.group(1))

        ambiguities: list[str] = []
        for source in graph:
            paths: dict[str, list[list[str]]] = {}

            def walk(node: str, path: list[str]) -> None:
                for target in graph.get(node, []):
                    if target in path:
                        continue
                    next_path = [*path, target]
                    paths.setdefault(target, []).append(next_path)
                    walk(target, next_path)

            walk(source, [source])
            for target, candidates in paths.items():
                if len(candidates) > 1:
                    rendered = " | ".join(" -> ".join(path) for path in candidates)
                    ambiguities.append(f"{source} => {target}: {rendered}")
        self.assertFalse(ambiguities, "Ambiguous active filter paths:\n" + "\n".join(ambiguities))

    def test_every_page_keeps_the_approved_cycle_geometry(self):
        pages_root = TEMPLATE / f"{PROJECT_NAME}.Report" / "definition" / "pages"
        required_shapes = {
            (0, 0, 1680, 64),
            (22, 124, 1636, 77),
            (22, 211, 400, 132),
            (434, 211, 400, 132),
            (846, 211, 400, 132),
            (1258, 211, 400, 132),
        }
        for page_dir in pages_root.glob("ReportSection*"):
            by_type: dict[str, set[tuple[int, int, int, int]]] = {}
            for path in page_dir.glob("visuals/*/visual.json"):
                value = json.loads(path.read_text(encoding="utf-8"))
                position = value["position"]
                geometry = tuple(
                    int(position[key]) for key in ("x", "y", "width", "height")
                )
                by_type.setdefault(value["visual"]["visualType"], set()).add(geometry)
            with self.subTest(page=page_dir.name):
                self.assertTrue(required_shapes <= by_type["shape"])
                self.assertTrue(all(item[3] >= 76 for item in by_type["slicer"]))
                self.assertTrue(all(item[3] >= 36 for item in by_type["cardVisual"]))
                self.assertIn((22, 68, 1636, 48), by_type["pageNavigator"])
                self.assertIn((22, 917, 1636, 18), by_type["textbox"])
                self.assertFalse(any(x == 0 and width == 210 for x, _, width, _ in by_type["shape"]))

    def test_navigation_slicer_sync_and_native_timeline_contract(self):
        pages_root = TEMPLATE / f"{PROJECT_NAME}.Report" / "definition" / "pages"
        sync_bindings: dict[str, set[tuple[str, str]]] = {}
        found_timeline = False
        for path in pages_root.glob("*/visuals/*/visual.json"):
            value = json.loads(path.read_text(encoding="utf-8"))
            visual = value["visual"]
            visual_type = visual["visualType"]
            if visual_type == "pageNavigator":
                layout = visual["objects"]["layout"][0]["properties"]
                self.assertEqual(layout["columnCount"]["expr"]["Literal"]["Value"], "5L")
                self.assertEqual(layout["rowCount"]["expr"]["Literal"]["Value"], "1L")
                text_objects = visual["objects"]["text"]
                self.assertTrue(all(
                    item["properties"]["show"]["expr"]["Literal"]["Value"] == "true"
                    for item in text_objects
                ))
            if visual_type == "slicer":
                group = visual.get("syncGroup")
                self.assertIsNotNone(group, path)
                binding = next(self._bindings(visual, "Column"))
                sync_bindings.setdefault(group["groupName"], set()).add(binding)
            if visual_type == "barChart":
                query_state = visual["query"]["queryState"]
                self.assertEqual(set(query_state), {"Category", "Y", "Series"})
                self.assertEqual(next(self._bindings(query_state["Category"], "Column")), ("Shift Placement", "Row Label"))
                self.assertEqual(next(self._bindings(query_state["Series"], "Column")), ("Shift Placement", "Segment Series"))
                found_timeline = True
            if visual_type == "cardVisual":
                title = visual["visualContainerObjects"]["title"][0]["properties"]["show"]
                self.assertEqual(title["expr"]["Literal"]["Value"], "false")
        self.assertTrue(found_timeline)
        self.assertTrue(all(len(bindings) == 1 for bindings in sync_bindings.values()))

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
