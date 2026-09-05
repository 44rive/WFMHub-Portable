from __future__ import annotations

import unittest
from datetime import datetime
from pathlib import Path

from wfmhub.models import _exact_status_gaps, _final_verint_events
from wfmhub.rules import load_rulebook
from wfmhub.utils import classify_status, parse_verint_interval, subtract_intervals


class UtilsTests(unittest.TestCase):
    def test_duplicate_verint_event_across_extracts_is_one_business_event(self):
        rulebook = load_rulebook(
            Path(__file__).resolve().parents[1],
            Path(__file__).resolve().parents[1] / "config" / "default_rules.toml",
        )
        start = datetime(2026, 8, 1, 9, 0)
        end = datetime(2026, 8, 1, 10, 0)
        base = {"_events": [
            {
                "activity": "Short sickness", "event_start": start,
                "event_end": end, "source_file": "Activities daily.txt",
            },
            {
                "activity": "Short sickness", "event_start": start,
                "event_end": end, "source_file": "Activities full month.txt",
            },
        ]}

        events = _final_verint_events(base, rulebook)

        self.assertEqual(len(events), 1)
        self.assertEqual(
            events[0]["source_file"],
            "Activities daily.txt; Activities full month.txt",
        )

    def test_verint_activity_preserves_internal_pipe(self):
        activity, start, end = parse_verint_interval(
            ".AP BEN | Product Loss | IT Failure 08/01/2026 9:00 AM-08/01/2026 10:00 AM"
        )
        self.assertEqual(activity, "Product Loss | IT Failure")
        self.assertEqual(start, datetime(2026, 8, 1, 9, 0))
        self.assertEqual(end, datetime(2026, 8, 1, 10, 0))

    def test_unavailable_is_not_misclassified_as_available(self):
        self.assertEqual(classify_status("Unavailable"), "Unavailable")
        self.assertEqual(classify_status("Pause écran"), "Break")

    def test_subtraction_returns_physical_remaining_interval(self):
        start = datetime(2026, 8, 1, 8, 0)
        end = datetime(2026, 8, 1, 8, 20)
        result = subtract_intervals(start, end, [(start, datetime(2026, 8, 1, 8, 10))])
        self.assertEqual(result, [(datetime(2026, 8, 1, 8, 10), end)])

    def test_exact_status_gaps_never_bridge_a_real_return(self):
        result = _exact_status_gaps([
            (datetime(2026, 8, 1, 12, 0), datetime(2026, 8, 1, 12, 20)),
            # The agent returned for two minutes before going offline again.
            (datetime(2026, 8, 1, 12, 22), datetime(2026, 8, 1, 13, 0)),
        ], minimum_minutes=5)
        self.assertEqual(result, [
            (datetime(2026, 8, 1, 12, 0), datetime(2026, 8, 1, 12, 20)),
            (datetime(2026, 8, 1, 12, 22), datetime(2026, 8, 1, 13, 0)),
        ])


if __name__ == "__main__":
    unittest.main()
