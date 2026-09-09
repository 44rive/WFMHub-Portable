from __future__ import annotations

import unittest
from pathlib import Path


REPO = Path(__file__).resolve().parents[1]


class LauncherTests(unittest.TestCase):
    def test_windows_launchers_only_use_embedded_python(self):
        for name in ("SETUP.cmd", "WFMHub.cmd"):
            text = (REPO / name).read_text(encoding="utf-8").lower()
            self.assertIn(r"_system\runtime\python.exe", text)
            self.assertNotIn("py -3", text)
            self.assertNotIn("python -m wfmhub", text)
            self.assertIn("source code zip", text)
            self.assertIn("title wfmhub portable", text)
            self.assertIn("if not defined no_color color 0b", text)

    def test_pcs_packages_the_reviewed_power_query_installer(self):
        text = (
            REPO / "packaging" / "windows" / "build_portable.py"
        ).read_text(encoding="utf-8")
        self.assertIn(
            'stage / "_system" / "scripts" / "Install-PCSWorkbook.ps1"',
            text,
        )
        installer = (
            REPO / "packaging" / "windows" / "Install-PCSWorkbook.ps1"
        ).read_text(encoding="utf-8")
        self.assertIn('ValidateSet("Install", "Refresh")', installer)
        self.assertIn('"tblCoachingQueue"', installer)
        self.assertNotIn('"tblCoachingActions"', installer)
        self.assertNotIn("/home/founder", text)
        self.assertNotIn("/home/founder", installer)


if __name__ == "__main__":
    unittest.main()
