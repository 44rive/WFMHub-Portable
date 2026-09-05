from __future__ import annotations

import sqlite3
import tempfile
import unittest
from datetime import date, datetime
from pathlib import Path

from wfmhub.database import DatabaseConnection, _migration_statements
from wfmhub.mapping import load_queue_mapping
from wfmhub.metrics import load_metric_catalog
from wfmhub.models import (
    _aggregate_forecast_hour_rows,
    _build_call_service,
    _map_forecast_interval_rows,
)
from wfmhub.rules import load_rulebook
from wfmhub.service_flash import _flash_cards, _flash_columns, _included_in_flash_total, _ratio
from wfmhub.service_profiles import load_service_profiles


REPO = Path(__file__).resolve().parents[1]


class CallServiceModelTests(unittest.TestCase):
    def test_four_fifteen_minute_forecasts_become_one_hour(self):
        mapping = load_queue_mapping(REPO / "config" / "default_queue_mapping.csv")
        rows = []
        for minute, volume, required, service_level, aht in (
            (0, 1, 2, 0.5, 100),
            (15, 2, 4, 0.6, 200),
            (30, 3, 6, 0.7, 300),
            (45, 4, 8, 0.8, 400),
        ):
            rows.append({
                "business_date": date(2026, 9, 1),
                "interval_start": datetime(2026, 9, 1, 8, minute),
                "interval_minutes": 15,
                "queue_name": "Combined - All Media",
                "volume_forecast": volume,
                "fte_forecast": required - 1,
                "fte_required": required,
                "sl_forecast": service_level,
                "sl_required": 0.8,
                "aht_forecast_seconds": aht,
                "source_file": "RSA_NL_09-2026.txt",
            })
        result = _aggregate_forecast_hour_rows(rows, mapping)
        self.assertEqual(len(result), 1)
        hour = result[0]
        self.assertEqual(hour["hour_start"], datetime(2026, 9, 1, 8, 0))
        self.assertEqual(hour["volume_forecast"], 10)
        self.assertEqual(hour["fte_required"], 5)
        self.assertAlmostEqual(hour["sl_forecast"], 0.7)
        self.assertAlmostEqual(hour["aht_forecast_seconds"], 300)
        self.assertEqual(hour["source_interval_minutes"], 15)
        self.assertEqual(hour["source_interval_count"], 4)
        native = _map_forecast_interval_rows(rows, mapping)
        self.assertEqual(len(native), 4)
        self.assertEqual(native[1]["interval_start"], datetime(2026, 9, 1, 8, 15))
        self.assertEqual(native[1]["interval_end"], datetime(2026, 9, 1, 8, 30))
        self.assertEqual(native[1]["fte_required"], 4)

    def test_oem_visible_layout_matches_book1_ford_toyota_contract(self):
        catalog = load_service_profiles(
            REPO, REPO / "config" / "default_service_profiles.toml",
        )
        profile = catalog.select("ford_oem_fr", date(2026, 9, 1))
        self.assertEqual(profile.flash_total_groups, ("Ford", "Toyota"))
        self.assertTrue(_included_in_flash_total(
            profile, {"queue": "APFR_PAR_RSA_CSTRUCTR_FORD_ASSISTANCE_FR"},
        ))
        self.assertTrue(_included_in_flash_total(
            profile, {"queue": "APFR_PAR_RSA_CSTRUCTR_TOYOTA-LEXUS_FR"},
        ))
        self.assertFalse(_included_in_flash_total(
            profile, {"queue": "APFR_PAR_RSA_CHERY_ASSISTANCE_FR"},
        ))
        blank = {
            "hour_label": "08:00", "forecast": 10, "offered": 5,
            "answered": 4, "answered_within_target": 3,
            "forecast_attainment": 0.5, "availability": 0.8,
            "service_level": 0.6, "aht_seconds": 200,
            "data_state": "READY",
            "groups": {
                "Ford": {"offered": 3, "service_level": 2 / 3, "availability": 1},
                "Toyota": {"offered": 2, "service_level": 0.5, "availability": 0.5},
            },
        }
        headers, _, _, _, _ = _flash_columns(profile, [blank])
        self.assertEqual(headers, [
            "Hour", "Volume Forecasted", "Volume Ford", "Volume Toyota",
            "SL Ford", "Availability Ford", "Availability Toyota", "AHT",
        ])
        cards = _flash_cards(
            profile,
            {"availability": 0.8, "service_level": 0.6,
             "service_method": "gross_20", "forecast_attainment": 0.5},
            blank["groups"],
            [blank],
        )
        self.assertEqual([card[0] for card in cards], [
            "Availability OEM", "Availability Ford", "Availability Toyota",
            "Deviation", "TSL OEM", "TSL Ford", "TSL Toyota",
        ])

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
