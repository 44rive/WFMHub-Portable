from __future__ import annotations

import unittest
from pathlib import Path


REPO = Path(__file__).resolve().parents[1]


class LauncherTests(unittest.TestCase):
    def test_windows_launchers_only_use_embedded_python(self):
        for name in ("SETUP.cmd", "WFMHub.cmd", "UPGRADE.cmd"):
            text = (REPO / name).read_text(encoding="utf-8").lower()
            self.assertIn(r"_system\runtime\python.exe", text)
            self.assertNotIn("py -3", text)
            self.assertNotIn("python -m wfmhub", text)
            self.assertIn("source code zip", text)
            self.assertIn("title wfmhub portable", text)
            self.assertIn("if not defined no_color color 0b", text)

    def test_webapp_launcher_uses_only_the_embedded_runtime(self):
        text = (REPO / "WEBAPP.cmd").read_text(encoding="utf-8").lower()
        self.assertIn(r"_system\runtime\python.exe", text)
        self.assertIn("-m wfmhub", text)
        self.assertIn(" web", text)
        self.assertNotIn("py -3", text)

    def test_pcs_packages_the_in_place_feed_sync(self):
        text = (
            REPO / "packaging" / "windows" / "build_portable.py"
        ).read_text(encoding="utf-8")
        self.assertIn(
            'stage / "_system" / "scripts" / "Sync-PCSWorkbook.ps1"',
            text,
        )
        installer = (
            REPO / "packaging" / "windows" / "Sync-PCSWorkbook.ps1"
        ).read_text(encoding="utf-8")
        self.assertIn('ValidateSet("Sync")', installer)
        self.assertIn('"tblPcsCoachingView"', installer)
        self.assertIn('"tblPcsFilters"', installer)
        self.assertIn("$table.Resize($newRange)", installer)
        self.assertIn("$destination.Value2 = $data", installer)
        self.assertNotIn("Remove-StarterTable", installer)
        self.assertNotIn("Queries.Add", installer)
        self.assertNotIn('"tblCoachingActions"', installer)
        self.assertNotIn("/home/founder", text)
        self.assertNotIn("/home/founder", installer)


if __name__ == "__main__":
    unittest.main()
