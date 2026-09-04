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


if __name__ == "__main__":
    unittest.main()
