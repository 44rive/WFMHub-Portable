from __future__ import annotations

import io
import unittest

from wfmhub.progress import ProgressBar


class ProgressBarTests(unittest.TestCase):
    def test_determinate_and_indeterminate_output_is_cmd_safe(self):
        stream = io.StringIO()
        bar = ProgressBar("WFMHub", stream=stream, width=12, enabled=True)

        bar.update(0.5, "Building attendance")
        bar.pulse("Call by Call: 2,000 rows scanned")
        bar.finish("Refresh complete")

        output = stream.getvalue()
        self.assertIn(" 50% Building attendance", output)
        self.assertIn("working Call by Call", output)
        self.assertIn("100% Refresh complete", output)
        self.assertNotIn("\x1b", output)
        self.assertTrue(output.endswith("\n"))

    def test_disabled_bar_is_silent(self):
        stream = io.StringIO()
        bar = ProgressBar(stream=stream, enabled=False)
        bar.update(0.5, "Working")
        bar.pulse("Still working")
        bar.finish()
        bar.fail()
        self.assertEqual(stream.getvalue(), "")

    def test_failure_closes_the_line(self):
        stream = io.StringIO()
        bar = ProgressBar(stream=stream, width=12, enabled=True)
        bar.fail("bad input")
        self.assertIn("FAILED bad input", stream.getvalue())
        self.assertTrue(stream.getvalue().endswith("\n"))


if __name__ == "__main__":
    unittest.main()
