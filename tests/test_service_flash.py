from __future__ import annotations

import sqlite3
import tempfile
import unittest
from dataclasses import replace
from datetime import date, datetime, time, timedelta
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
from wfmhub.service_flash import (
    _attendance_pulse,
    _flash_cards,
    _flash_columns,
    _hourly_model,
    _included_in_flash_total,
    _profile_rows,
    _ratio,
    _workforce_by_hour,
)
from wfmhub.service_profiles import load_service_profiles


REPO = Path(__file__).resolve().parents[1]


class CallServiceModelTests(unittest.TestCase):
    def test_attendance_pulse_counts_only_reliable_current_gaps(self):
        raw = sqlite3.connect(
            ":memory:",
            detect_types=sqlite3.PARSE_DECLTYPES | sqlite3.PARSE_COLNAMES,
        )
        conn = DatabaseConnection(raw)
        conn.execute(
            """CREATE TABLE mart.attendance_agent_day (
                   business_date DATE, agent_id VARCHAR, agent_name VARCHAR,
                   team_leader VARCHAR, ops_manager VARCHAR, lob VARCHAR,
                   language VARCHAR, scheduled_start TIMESTAMP,
                   scheduled_end TIMESTAMP, assignment_type VARCHAR,
                   attendance_result VARCHAR, call_action VARCHAR,
                   requires_call BOOLEAN, actual_first_seen TIMESTAMP,
                   actual_last_seen TIMESTAMP, uncoded_late_minutes INTEGER,
                   actual_evidence VARCHAR, source_loaded BOOLEAN,
                   is_provisional BOOLEAN, evaluation_as_of TIMESTAMP,
                   planning_overlay VARCHAR, planning_overlay_minutes INTEGER
               )"""
        )
        conn.execute(
            """CREATE TABLE mart.shift_timeline_segment (
                   business_date DATE, agent_id VARCHAR, lob VARCHAR,
                   segment_start TIMESTAMP, segment_end TIMESTAMP,
                   planned_state VARCHAR, actual_category VARCHAR,
                   mismatch_type VARCHAR, is_gap BOOLEAN,
                   observed_source VARCHAR
               )"""
        )
        report_day = date.today()
        early_start = datetime.combine(report_day, time(6))
        early_end = datetime.combine(report_day, time(10))
        start = datetime.combine(report_day, time(8))
        end = datetime.combine(report_day, time(16))
        checkpoint = datetime.combine(report_day, time(12))
        present_evidence_end = checkpoint - timedelta(minutes=10)
        refresh_time = checkpoint + timedelta(minutes=20)
        conn.executemany(
            "INSERT INTO mart.attendance_agent_day VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
            [
                (report_day, "present", "Present Agent", "TL", "OM", "RSA NL", "NL", start, end, "Working", "Late - shift in progress", "CALL_LATE", True, start, present_evidence_end, 10, "AGENT_STATUS", True, True, refresh_time, None, 0),
                (report_day, "offline", "Offline Agent", "TL", "OM", "RSA NL", "NL", start, end, "Working", "Shift in progress", "NONE", False, start, checkpoint, 0, "AGENT_STATUS", True, True, refresh_time, None, 0),
                (report_day, "no-show", "No Show Agent", "TL", "OM", "RSA NL", "NL", start, end, "Working", "Not seen - shift in progress", "CALL_NOT_SEEN_NOW", True, None, None, 0, "AGENT_STATUS", True, True, refresh_time, None, 0),
                (report_day, "unknown", "Unknown Agent", "TL", "OM", "RSA NL", "NL", start, end, "Working", "Data not loaded", "NONE", False, None, None, 0, "NONE", True, True, refresh_time, None, 0),
                (report_day, "early", "Early Agent", "TL", "OM", "RSA NL", "NL", early_start, early_end, "Working", "Early leave", "NONE", False, early_start, datetime.combine(report_day, time(9, 30)), 0, "AGENT_STATUS", True, False, refresh_time, None, 0),
                (report_day, "pto", "PTO Agent", "TL", "OM", "RSA NL", "NL", start, end, "Planned absence", "PTO", "NONE", False, None, None, 0, "PTO", True, False, refresh_time, "PTO: Vacation", 480),
                (report_day, "away", "Away Agent", "TL", "OM", "RSA NL", "NL", start, end, "Planned absence", "Away", "NONE", False, None, None, 0, "AWAY", True, False, refresh_time, "AWAY: Long sickness", 480),
                (report_day, "partial-pto", "Partial PTO Agent", "TL", "OM", "RSA NL", "NL", start, end, "Working", "Shift in progress", "NONE", False, None, None, 0, "NONE", True, True, refresh_time, "PTO: Vacation", 240),
            ],
        )
        conn.executemany(
            "INSERT INTO mart.shift_timeline_segment VALUES (?,?,?,?,?,?,?,?,?,?)",
            [
                (report_day, "present", "RSA NL", start, present_evidence_end, "WORK", "Productive", "MATCH", False, "AGENT_STATUS"),
                (report_day, "offline", "RSA NL", datetime.combine(report_day, time(11)), checkpoint, "WORK", "Logged Off", "GAP", True, "AGENT_STATUS"),
                (report_day, "no-show", "RSA NL", start, checkpoint, "WORK", "Logged Off", "GAP", True, "AGENT_STATUS"),
                (report_day, "early", "RSA NL", early_start, datetime.combine(report_day, time(9, 30)), "WORK", "Productive", "MATCH", False, "AGENT_STATUS"),
                (report_day, "pto", "RSA NL", start, end, "PTO: Vacation", "PTO", "PLANNED_TIME_OFF", False, "PTO"),
                (report_day, "away", "RSA NL", start, end, "AWAY: Long sickness", "AWAY", "PLANNED_TIME_OFF", False, "AWAY"),
                (report_day, "partial-pto", "RSA NL", start, checkpoint, "PTO: Vacation", "PTO", "PLANNED_TIME_OFF", False, "PTO"),
                (report_day, "partial-pto", "RSA NL", checkpoint, end, "Working", "FUTURE", "FUTURE", False, "NONE"),
            ],
        )
        profile = load_service_profiles(
            REPO, REPO / "config" / "default_service_profiles.toml",
        ).select("rsa_nl", report_day)
        rows, summary = _attendance_pulse(conn, profile, report_day)
        by_agent = {row["agent_id"]: row for row in rows}
        self.assertEqual(by_agent["pto"]["attendance_state"], "PTO")
        self.assertEqual(by_agent["pto"]["pulse_action"], "NONE")
        self.assertEqual(by_agent["away"]["attendance_state"], "AWAY")
        self.assertEqual(by_agent["away"]["pulse_action"], "NONE")
        self.assertEqual(by_agent["partial-pto"]["attendance_state"], "PTO — NOT DUE")
        self.assertEqual(by_agent["partial-pto"]["pulse_action"], "NONE")
        self.assertEqual(summary["scheduled_now"], 4)
        self.assertEqual(summary["due_hc"], 5)
        self.assertEqual(summary["time_off_hc"], 3)
        self.assertEqual(summary["present_hc"], 3)
        self.assertEqual(summary["no_show_hc"], 1)
        self.assertEqual(summary["offline_now"], 1)
        self.assertEqual(summary["unknown_hc"], 1)
        self.assertEqual(summary["late_today"], 1)
        self.assertEqual(summary["early_leave"], 1)
        self.assertEqual(summary["call_now"], 2)
        self.assertEqual(
            summary["due_hc"],
            summary["present_hc"] + summary["no_show_hc"] + summary["unknown_hc"],
        )
        self.assertEqual(summary["checkpoint"], checkpoint)
        self.assertEqual(by_agent["no-show"]["pulse_action"], "CALL_NO_SHOW")
        self.assertEqual(by_agent["no-show"]["attendance_state"], "NO SHOW")
        self.assertEqual(by_agent["offline"]["pulse_action"], "CHECK_OFFLINE_NOW")
        self.assertEqual(by_agent["offline"]["attendance_state"], "PRESENT — OFFLINE NOW")
        self.assertFalse(by_agent["offline"]["no_show"])
        self.assertEqual(by_agent["present"]["attendance_state"], "PRESENT — LATE")
        self.assertEqual(by_agent["present"]["current_evidence"], "AGENT_STATUS")
        self.assertEqual(by_agent["early"]["attendance_state"], "PRESENT — EARLY LEAVE")
        self.assertFalse(by_agent["early"]["no_show"])
        self.assertEqual(by_agent["unknown"]["attendance_state"], "UNKNOWN — POSSIBLE NO SHOW")
        self.assertEqual(
            by_agent["unknown"]["pulse_action"], "CHECK_DATA_POSSIBLE_NO_SHOW",
        )
        self.assertFalse(by_agent["unknown"]["call_now"])
        hourly = _workforce_by_hour(profile, report_day, rows, summary)
        self.assertEqual(hourly[11]["no_show_hc"], 1)
        self.assertIsNone(hourly[12]["no_show_hc"])
        partial_only = [{
            "agent_id": "partial-absent", "no_show": True,
            "assignment_type": "Working", "scheduled_start": start,
            "scheduled_end": end,
            "expected_work_intervals": [
                (datetime.combine(report_day, time(10)), end),
            ],
        }]
        final_summary = {
            "mode": "FINAL DAY",
            "checkpoint": datetime.combine(report_day, time(23, 59, 59)),
        }
        hourly = _workforce_by_hour(
            profile, report_day, partial_only, final_summary,
        )
        self.assertEqual(hourly[8]["no_show_hc"], 0)
        self.assertEqual(hourly[10]["no_show_hc"], 1)
        conn.close()

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
            "abandoned_within_target": 2,
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
            "short_abandoned": 0, "abandoned_within_target": 0,
            "answered_within_target": 72,
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
        mapping = load_queue_mapping(REPO / "config" / "default_queue_mapping.csv")
        for configured_profile in catalog.profiles:
            for queue in configured_profile.flash_queues:
                self.assertEqual(
                    mapping.map_actual("STORM", queue, None, None).status,
                    "MAPPED",
                    f"{configured_profile.profile_id}: {queue}",
                )
        profile = catalog.select("ford_oem_fr", date(2026, 9, 1))
        self.assertEqual(profile.staffing_lobs, ("OEM FR",))
        self.assertEqual(len(profile.flash_queues), 6)
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
        self.assertTrue(_included_in_flash_total(
            profile, {"queue": "APBN_BRU_MOBILITY_Ford_Assistance_FR"},
        ))
        self.assertFalse(_included_in_flash_total(
            profile, {"queue": "APCH_ZRH_RSA_Ford_Assistance_FR"},
        ))
        ford_nl = catalog.select("ford_nl", date(2026, 9, 1))
        self.assertEqual(len(ford_nl.flash_queues), 6)
        self.assertTrue(_included_in_flash_total(
            ford_nl, {"queue": "APBN_AMS_MOBILITY_Ford_Assistance_NL"},
        ))
        self.assertTrue(_included_in_flash_total(
            ford_nl, {"queue": "APBN_BRU_MOBILITY_Ford_Assistance_VL"},
        ))
        self.assertFalse(_included_in_flash_total(
            ford_nl, {"queue": "APBN_BRU_MOBILITY_Ford_Assistance_FR"},
        ))
        rsa_nl = catalog.select("rsa_nl", date(2026, 9, 1))
        self.assertEqual(len(rsa_nl.flash_queues), 30)
        self.assertTrue(_included_in_flash_total(
            rsa_nl, {"queue": "APBN_AMS_MOBILITY_INSURAN_Front_NL"},
        ))
        self.assertTrue(_included_in_flash_total(
            rsa_nl, {"queue": "APBN_AMS_MOBILITY_PROVIDER_Local_NL"},
        ))
        self.assertTrue(_included_in_flash_total(
            rsa_nl, {"queue": "APBN_AMS_RSA_Ford_Assistance_NL"},
        ))
        self.assertFalse(_included_in_flash_total(
            rsa_nl, {"queue": "APBN_AMS_MOBILITY_VARIOUS_VariousAssist_NL"},
        ))
        rsa_be = catalog.select("rsa_be", date(2026, 9, 1))
        self.assertEqual(rsa_be.staffing_lobs, ("RSA FR", "RSA VL"))
        self.assertEqual(len(rsa_be.flash_queues), 36)
        self.assertTrue(_included_in_flash_total(
            rsa_be, {"queue": "APBN_BRU_MOBILITY_PROVIDER_Interco_EN"},
        ))
        self.assertFalse(_included_in_flash_total(
            rsa_be, {"queue": "APBN_BRU_MOBILITY_POLICE_OBU_EN"},
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
            "Hour", "Forecast", "Actual", "Variance", "Ford Volume",
            "Chery Volume", "Toyota Volume", "TSL OEM", "TSL Ford",
            "TSL Chery", "TSL Toyota", "Routed Rate", "AHT", "No Show HC",
            "Data State",
        ])
        cards = _flash_cards(
            profile,
            {"availability": 0.8, "service_level": 0.6,
             "service_method": "gross_30", "forecast_attainment": 0.5},
            blank["groups"],
            {"no_show_hc": 1, "offline_now": 2},
        )
        self.assertEqual([card[0] for card in cards], [
            "TSL", "Routed Rate", "Actual", "Forecast", "No Show HC",
            "Offline Now", "Call Now",
        ])
        ford_cards = _flash_cards(
            ford_nl,
            {"forecast": 10, "offered": 9, "answered": 8,
             "answered_within_target": 7, "forecast_attainment": 0.9,
             "availability": 8 / 9, "service_level": 7 / 9,
             "service_method": "storm_custom_30", "aht_seconds": 200},
            {},
            {"no_show_hc": 1, "offline_now": 2},
        )
        self.assertEqual([card[0] for card in ford_cards], [
            "TSL", "Routed Rate", "Actual", "Forecast", "No Show HC",
            "Offline Now", "Call Now",
        ])

    def test_forecast_only_hour_has_no_attainment_instead_of_crashing(self):
        self.assertIsNone(_ratio(None, 10))
        self.assertIsNone(_ratio(10, None))
        self.assertIsNone(_ratio(10, 0))
        self.assertEqual(_ratio(8, 10), 0.8)

    def test_flash_headline_and_hourly_view_cover_the_whole_day(self):
        raw = sqlite3.connect(
            ":memory:",
            detect_types=sqlite3.PARSE_DECLTYPES | sqlite3.PARSE_COLNAMES,
        )
        conn = DatabaseConnection(raw)
        conn.execute(
            """CREATE TABLE mart.forecast_hour (
                   business_date DATE, comparison_scope VARCHAR,
                   hour_start TIMESTAMP, volume_forecast DOUBLE
            )"""
        )
        conn.execute(
            """CREATE TABLE mart.attendance_agent_day (
                   business_date DATE, lob VARCHAR, agent_id VARCHAR,
                   scheduled_start TIMESTAMP, scheduled_end TIMESTAMP,
                   assignment_type VARCHAR, planned_work_minutes DOUBLE
               )"""
        )
        conn.execute(
            """CREATE TABLE mart.absence_event (
                   business_date DATE, lob VARCHAR, agent_id VARCHAR,
                   category VARCHAR, event_start TIMESTAMP,
                   event_end TIMESTAMP, counts_as_absence BOOLEAN
            )"""
        )
        conn.execute(
            """CREATE TABLE mart.shift_timeline_segment (
                   business_date DATE, lob VARCHAR, agent_id VARCHAR,
                   segment_start TIMESTAMP, segment_end TIMESTAMP,
                   is_gap BOOLEAN
               )"""
        )
        report_day = date(2026, 9, 5)
        conn.executemany(
            "INSERT INTO mart.forecast_hour VALUES (?,?,?,?)",
            [
                (report_day, "RSA NL", datetime(2026, 9, 5, 1), 10),
                (report_day, "RSA NL", datetime(2026, 9, 5, 7), 20),
            ],
        )
        rows = [
            {
                "business_date": report_day, "hour_start": datetime(2026, 9, 5, 1),
                "queue": "APBN_AMS_RSA_INSURAN_All_NL", "offered": 8,
                "answered": 7, "abandoned": 1, "short_abandoned": 0,
                "abandoned_within_target": 1, "answered_within_target": 6,
                "handled_seconds": 700,
            },
            {
                "business_date": report_day, "hour_start": datetime(2026, 9, 5, 7),
                "queue": "APBN_AMS_RSA_INSURAN_All_NL", "offered": 12,
                "answered": 10, "abandoned": 2, "short_abandoned": 1,
                "abandoned_within_target": 1, "answered_within_target": 9,
                "handled_seconds": 1_000,
            },
        ]
        profiles = load_service_profiles(
            REPO, REPO / "config" / "default_service_profiles.toml",
        )
        profile = profiles.select("rsa_nl", report_day)
        mapping = load_queue_mapping(REPO / "config" / "default_queue_mapping.csv")
        metrics = load_metric_catalog(REPO, REPO / "config" / "default_metrics.toml")
        hourly, total, groups, cutoff = _hourly_model(
            conn, profile, mapping, metrics, report_day, rows,
        )
        self.assertEqual(hourly[0]["hour"], 0)
        self.assertEqual(hourly[-1]["hour"], 23)
        self.assertEqual(len(hourly), 24)
        self.assertEqual(cutoff, 7)
        self.assertEqual(total["offered"], 20)
        self.assertEqual(total["forecast"], 30)
        self.assertAlmostEqual(total["service_level"], 15 / 18)
        self.assertEqual(groups["RSA"]["offered"], 20)
        conn.close()

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
            # Wait plus ringing stays inside the configured 30-second target.
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
        self.assertEqual(result[:9], (6, 2, 4, 1, 2, 2, 170.0, 6, 1))
        self.assertAlmostEqual(result[9], 2 / 4)
        # The Storm screenshot's routed rate uses every entered queue entry.
        self.assertAlmostEqual(result[10], 2 / 6)
        self.assertAlmostEqual(result[11], 4 / 6)
        self.assertAlmostEqual(result[12], 85.0)
        rsa_profile = load_service_profiles(
            REPO, REPO / "config" / "default_service_profiles.toml",
        ).select("rsa_nl", date(2026, 8, 1))
        exact_profile = replace(
            rsa_profile,
            service_scopes=("DIFFERENT MODEL SCOPE",),
            flash_queues=("MAPPED_QUEUE",),
        )
        selected = _profile_rows(
            conn, exact_profile, date(2026, 8, 1), date(2026, 8, 1),
        )
        self.assertEqual(len(selected), 1)
        self.assertEqual(selected[0]["queue"], "MAPPED_QUEUE")
        conn.close()


if __name__ == "__main__":
    unittest.main()
