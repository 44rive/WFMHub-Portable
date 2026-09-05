from __future__ import annotations

import sqlite3
import tempfile
import unittest
from datetime import date, datetime
from pathlib import Path

from wfmhub.database import DatabaseConnection, _migration_statements
from wfmhub.mapping import load_queue_mapping
from wfmhub.metrics import evaluate_metric, load_metric_catalog
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
    def test_storm_screenshot_arithmetic_is_reproduced(self):
        catalog = load_metric_catalog(
            REPO, REPO / "config" / "default_metrics.toml",
        )
        dimensions = {"lob": "Ford FR", "source_system": "CALL_BY_CALL"}
        components = {
            "offered": 278,
            "answered": 266,
            "abandoned": 12,
            "short_abandoned": 2,
            "abandoned_within_target": 0,
            "answered_within_target": 249,
            "handled_seconds": 0,
        }
        service = evaluate_metric(
            catalog.method_for("service_level", date(2026, 9, 5), dimensions),
            components,
        )
        routed = evaluate_metric(
            catalog.method_for(
                "service_availability_business", date(2026, 9, 5), dimensions,
            ),
            components,
        )
        self.assertAlmostEqual(service.value, 0.9021739130434783)
        self.assertAlmostEqual(routed.value, 0.9568345323741007)

        ford_components = dict(components)
        ford_components.update({
            "offered": 74, "answered": 73, "abandoned": 1,
            "short_abandoned": 0, "answered_within_target": 72,
        })
        ford_service = evaluate_metric(
            catalog.method_for("service_level", date(2026, 9, 5), dimensions),
            ford_components,
        )
        self.assertAlmostEqual(ford_service.value, 0.972972972972973)

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

    def test_storm_visible_scopes_and_oem_layout(self):
        catalog = load_service_profiles(
            REPO, REPO / "config" / "default_service_profiles.toml",
        )
        profile = catalog.select("ford_oem_fr", date(2026, 9, 1))
        self.assertEqual(profile.flash_total_groups, ("Ford", "Chery", "Toyota"))
        self.assertTrue(_included_in_flash_total(
            profile, {"queue": "APFR_PAR_RSA_CSTRUCTR_FORD_ASSISTANCE_FR"},
        ))
        self.assertTrue(_included_in_flash_total(
            profile, {"queue": "APFR_PAR_RSA_CSTRUCTR_TOYOTA-LEXUS_FR"},
        ))
        self.assertTrue(_included_in_flash_total(
            profile, {"queue": "APFR_PAR_RSA_CHERY_ASSISTANCE_FR"},
        ))
        self.assertFalse(_included_in_flash_total(
            profile, {"queue": "APBN_BRU_MOBILITY_Ford_Assistance_FR"},
        ))
        ford_nl = catalog.select("ford_nl", date(2026, 9, 1))
        self.assertTrue(_included_in_flash_total(
            ford_nl, {"queue": "APBN_AMS_MOBILITY_Ford_Assistance_NL"},
        ))
        self.assertFalse(_included_in_flash_total(
            ford_nl, {"queue": "APBN_BRU_MOBILITY_Ford_Assistance_VL"},
        ))
        rsa_nl = catalog.select("rsa_nl", date(2026, 9, 1))
        self.assertTrue(_included_in_flash_total(
            rsa_nl, {"queue": "APBN_AMS_MOBILITY_INSURAN_Front_NL"},
        ))
        self.assertFalse(_included_in_flash_total(
            rsa_nl, {"queue": "APBN_AMS_MOBILITY_PROVIDER_Local_NL"},
        ))
        blank = {
            "hour_label": "08:00", "forecast": 10, "offered": 5,
            "answered": 4, "answered_within_target": 3,
            "forecast_attainment": 0.5, "availability": 0.8,
            "service_level": 0.6, "aht_seconds": 200,
            "data_state": "READY",
            "groups": {
                "Ford": {"offered": 3, "service_level": 2 / 3, "availability": 1},
                "Chery": {"offered": 1, "service_level": 1, "availability": 1},
                "Toyota": {"offered": 2, "service_level": 0.5, "availability": 0.5},
            },
        }
        headers, _, _, _, _ = _flash_columns(profile, [blank])
        self.assertEqual(headers, [
            "Hour", "Volume Forecasted", "Volume Ford", "Volume Chery",
            "Volume Toyota", "SL Ford", "SL Chery", "SL Toyota",
            "Routed Rate Ford", "Routed Rate Chery", "Routed Rate Toyota", "AHT",
        ])
        cards = _flash_cards(
            profile,
            {"availability": 0.8, "service_level": 0.6,
             "service_method": "gross_20", "forecast_attainment": 0.5},
            blank["groups"],
            [blank],
        )
        self.assertEqual([card[0] for card in cards], [
            "Routed Rate OEM", "SLA OEM", "SLA Ford", "SLA Chery",
            "SLA Toyota", "Deviation", "AHT",
        ])

    def test_forecast_only_hour_has_no_attainment_instead_of_crashing(self):
        self.assertIsNone(_ratio(None, 10))
        self.assertIsNone(_ratio(10, None))
        self.assertIsNone(_ratio(10, 0))
        self.assertEqual(_ratio(8, 10), 0.8)

    def test_inbound_queue_entries_drive_storm_flash_volume(self):
        raw = sqlite3.connect(
            ":memory:",
            detect_types=sqlite3.PARSE_DECLTYPES | sqlite3.PARSE_COLNAMES,
        )
        conn = DatabaseConnection(raw)
        conn.execute(
            """CREATE TABLE core.clean_call_leg (
                   business_date DATE, interaction_key VARCHAR, call_key VARCHAR,
                   call_start TIMESTAMP, call_direction VARCHAR, queue VARCHAR,
                   queue_wait_seconds DOUBLE, ringing_seconds DOUBLE,
                   agent_id VARCHAR, talk_seconds DOUBLE,
                   hold_seconds DOUBLE, wrap_seconds DOUBLE, transferred BOOLEAN,
                   language VARCHAR, lob VARCHAR, source_file VARCHAR
               )"""
        )
        for migration_name in ("013_call_service_flash.sql",):
            migration = (REPO / "sql" / "migrations" / migration_name).read_text(
                encoding="utf-8",
            )
            for statement in _migration_statements(migration):
                conn.execute(statement)
        conn.execute("CREATE TABLE mart.service_interval (placeholder INTEGER)")
        migration = (REPO / "sql" / "migrations" / "015_storm_service_reference.sql").read_text(
            encoding="utf-8",
        )
        for statement in _migration_statements(migration):
            conn.execute(statement)

        rows = [
            # Two inbound entries in one interaction. Storm counts both queue entries.
            (date(2026, 8, 1), "transfer", "leg-1", datetime(2026, 8, 1, 9, 0), "I", "MAPPED_QUEUE", 10, 0, None, 0, 0, 0, False, "NL", None, "calls.csv"),
            (date(2026, 8, 1), "transfer", "leg-2", datetime(2026, 8, 1, 9, 1), "I", "MAPPED_QUEUE", 0, 0, "999", 100, 10, 10, True, "NL", None, "calls.csv"),
            # Three unanswered interactions: short, in-target non-short and long.
            (date(2026, 8, 1), "short", "leg-3", datetime(2026, 8, 1, 9, 10), "I", "MAPPED_QUEUE", 3, 0, None, 0, 0, 0, False, "NL", None, "calls.csv"),
            (date(2026, 8, 1), "in-target", "leg-4", datetime(2026, 8, 1, 9, 15), "I", "MAPPED_QUEUE", 10, 0, None, 0, 0, 0, False, "NL", None, "calls.csv"),
            (date(2026, 8, 1), "long", "leg-5", datetime(2026, 8, 1, 9, 20), "I", "MAPPED_QUEUE", 30, 0, None, 0, 0, 0, False, "NL", None, "calls.csv"),
            # Wait is below 20, but wait + ringing is not: outside target.
            (date(2026, 8, 1), "ring", "leg-6", datetime(2026, 8, 1, 9, 25), "I", "MAPPED_QUEUE", 18, 3, "999", 50, 0, 0, False, "NL", None, "calls.csv"),
            # Mapped outbound traffic does not enter inbound service demand.
            (date(2026, 8, 1), "outbound", "leg-7", datetime(2026, 8, 1, 9, 30), "O", "MAPPED_QUEUE", 0, 0, "999", 50, 0, 0, False, "NL", None, "calls.csv"),
        ]
        conn.executemany(
            "INSERT INTO core.clean_call_leg VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
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
                      abandoned_within_target, answered_within_target,
                      handled_seconds, call_legs,
                      transferred_legs, service_level, service_availability,
                      abandon_rate, aht_seconds
               FROM mart.call_service_hour"""
        ).fetchone()
        self.assertEqual(result[:9], (6, 2, 4, 1, 2, 1, 170.0, 6, 1))
        self.assertAlmostEqual(result[9], 1 / 5)
        # The Storm screenshot's routed rate uses every entered queue entry.
        self.assertAlmostEqual(result[10], 2 / 6)
        self.assertAlmostEqual(result[11], 4 / 6)
        self.assertAlmostEqual(result[12], 85.0)
        conn.close()


if __name__ == "__main__":
    unittest.main()
