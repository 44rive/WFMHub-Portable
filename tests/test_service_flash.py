from __future__ import annotations

import sqlite3
import tempfile
import unittest
from datetime import date, datetime
from pathlib import Path

from wfmhub.database import DatabaseConnection, _migration_statements
from wfmhub.mapping import load_queue_mapping
from wfmhub.metrics import load_metric_catalog
from wfmhub.models import _build_call_service
from wfmhub.rules import load_rulebook
from wfmhub.service_flash import _ratio


REPO = Path(__file__).resolve().parents[1]


class CallServiceModelTests(unittest.TestCase):
    def test_forecast_only_hour_has_no_attainment_instead_of_crashing(self):
        self.assertIsNone(_ratio(None, 10))
        self.assertIsNone(_ratio(10, None))
        self.assertIsNone(_ratio(10, 0))
        self.assertEqual(_ratio(8, 10), 0.8)

    def test_interactions_not_transfer_legs_drive_flash_volume(self):
        raw = sqlite3.connect(
            ":memory:",
            detect_types=sqlite3.PARSE_DECLTYPES | sqlite3.PARSE_COLNAMES,
        )
        conn = DatabaseConnection(raw)
        conn.execute(
            """CREATE TABLE core.clean_call_leg (
                   business_date DATE, interaction_key VARCHAR, call_key VARCHAR,
                   call_start TIMESTAMP, call_direction VARCHAR, queue VARCHAR,
                   queue_wait_seconds DOUBLE, agent_id VARCHAR, talk_seconds DOUBLE,
                   hold_seconds DOUBLE, wrap_seconds DOUBLE, transferred BOOLEAN,
                   language VARCHAR, lob VARCHAR, source_file VARCHAR
               )"""
        )
        migration = (REPO / "sql" / "migrations" / "013_call_service_flash.sql").read_text(
            encoding="utf-8",
        )
        for statement in _migration_statements(migration):
            conn.execute(statement)

        rows = [
            # Two legs, one customer interaction. Only one offered call.
            (date(2026, 8, 1), "transfer", "leg-1", datetime(2026, 8, 1, 9, 0), "I", "MAPPED_QUEUE", 10, None, 0, 0, 0, False, "NL", None, "calls.csv"),
            (date(2026, 8, 1), "transfer", "leg-2", datetime(2026, 8, 1, 9, 1), "I", "MAPPED_QUEUE", 0, "999", 100, 10, 10, True, "NL", None, "calls.csv"),
            # Two unanswered interactions: one short abandon, one normal abandon.
            (date(2026, 8, 1), "short", "leg-3", datetime(2026, 8, 1, 9, 10), "I", "MAPPED_QUEUE", 3, None, 0, 0, 0, False, "NL", None, "calls.csv"),
            (date(2026, 8, 1), "long", "leg-4", datetime(2026, 8, 1, 9, 20), "I", "MAPPED_QUEUE", 30, None, 0, 0, 0, False, "NL", None, "calls.csv"),
            # Mapped outbound traffic does not enter inbound service demand.
            (date(2026, 8, 1), "outbound", "leg-5", datetime(2026, 8, 1, 9, 30), "O", "MAPPED_QUEUE", 0, "999", 50, 0, 0, False, "NL", None, "calls.csv"),
        ]
        conn.executemany(
            "INSERT INTO core.clean_call_leg VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
            rows,
        )
        with tempfile.TemporaryDirectory() as folder:
            mapping_file = Path(folder) / "queue_mapping.csv"
            mapping_file.write_text(
                "mapping_type,source_system,source_value,service_scope,designation\n"
                "queue,STORM,MAPPED_QUEUE,RSA NL,RSA NL\n",
                encoding="utf-8",
            )
            count = _build_call_service(
                conn,
                load_rulebook(REPO, REPO / "config" / "default_rules.toml"),
                load_metric_catalog(REPO, REPO / "config" / "default_metrics.toml"),
                load_queue_mapping(mapping_file),
                date(2026, 8, 1),
                date(2026, 8, 1),
            )

        self.assertEqual(count, 1)
        result = conn.execute(
            """SELECT offered, answered, abandoned, short_abandoned,
                      answered_within_target, handled_seconds, call_legs,
                      transferred_legs, service_level, service_availability,
                      abandon_rate, aht_seconds
               FROM mart.call_service_hour"""
        ).fetchone()
        self.assertEqual(result[:8], (3, 1, 2, 1, 1, 120.0, 4, 1))
        self.assertAlmostEqual(result[8], 0.5)
        self.assertAlmostEqual(result[9], 1 / 3)
        self.assertAlmostEqual(result[10], 2 / 3)
        self.assertAlmostEqual(result[11], 120.0)
        conn.close()


if __name__ == "__main__":
    unittest.main()
