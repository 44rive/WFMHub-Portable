from __future__ import annotations

import sqlite3
from datetime import date, datetime, time, timedelta
from pathlib import Path

from wfmhub.database import DatabaseConnection
from wfmhub.decision_products import (
    HOUR_BUCKETS, _realisations_aux_breakdown, _realisations_final_absence,
    _realisations_hours,
)
from wfmhub.metrics import evaluate_metric, load_metric_catalog
from wfmhub.rules import load_rulebook
from wfmhub.service_flash import _attendance_pulse, _issues_and_drivers_rows, _workforce_by_hour
from wfmhub.service_profiles import load_service_profiles


ROOT = Path(__file__).resolve().parents[1]
DAY = date(2026, 9, 15)
PROFILE = load_service_profiles(ROOT, ROOT / "config/default_service_profiles.toml").select("rsa_nl", DAY)
RULEBOOK = load_rulebook(ROOT, ROOT / "config/default_rules.toml")


def _segment(start: int, end: int, category: str, status: str, source: str = "AGENT_STATUS") -> dict:
    return {
        "segment_start": datetime(2026, 9, 15, start),
        "segment_end": datetime(2026, 9, 15, end),
        "actual_category": category, "actual_status": status,
        "observed_source": source,
    }


def test_rtm_hourly_headcount_counts_distinct_agents_and_subsets():
    work = (datetime(2026, 9, 15, 8), datetime(2026, 9, 15, 16))
    rows = [
        {"agent_id": "a", "expected_work_intervals": [work],
         "timeline_segments": [_segment(8, 10, "Productive", "Disponible"),
                               _segment(10, 16, "Productive", "Backoffice")],
         "no_show": False},
        {"agent_id": "b", "expected_work_intervals": [work],
         "timeline_segments": [_segment(8, 16, "Auxiliary", "Training")],
         "no_show": False},
        {"agent_id": "c", "expected_work_intervals": [work],
         "timeline_segments": [_segment(8, 16, "Logged Off", "Logged Off")],
         "no_show": True},
        {"agent_id": "d", "expected_work_intervals": [],
         "timeline_segments": [_segment(8, 16, "PTO", "")],
         "no_show": False},
    ]
    output = _workforce_by_hour(PROFILE, DAY, rows, {"mode": "FINAL DAY"}, RULEBOOK)
    assert output[9] == {
        "scheduled_hc": 3, "logged_hc": 2, "productive_hc": 1,
        "available_hc": 1, "unavailable_hc": 1, "bo_hc": 0,
        "no_show_hc": 1,
    }
    assert output[10]["productive_hc"] == 1
    assert output[10]["available_hc"] == 0
    assert output[10]["bo_hc"] == 1


def test_rtm_hour_counts_an_agent_who_handled_then_logged_off():
    work = (datetime(2026, 9, 15, 1), datetime(2026, 9, 15, 2))
    row = {"agent_id": "night", "expected_work_intervals": [work],
           "timeline_segments": [
               {"segment_start": work[0], "segment_end": datetime(2026, 9, 15, 1, 35),
                "actual_category": "Productive", "actual_status": "Disponible",
                "observed_source": "AGENT_STATUS"},
               {"segment_start": datetime(2026, 9, 15, 1, 35), "segment_end": work[1],
                "actual_category": "Logged Off", "actual_status": "Logged Off",
                "observed_source": "AGENT_STATUS"},
           ], "no_show": False}
    output = _workforce_by_hour(PROFILE, DAY, [row], {"mode": "FINAL DAY"}, RULEBOOK)
    assert output[1]["scheduled_hc"] == 1
    assert output[1]["logged_hc"] == 1
    assert output[1]["available_hc"] == 1


def test_realisations_hour_allocation_reconciles_agent_and_lob():
    raw = sqlite3.connect(":memory:", detect_types=sqlite3.PARSE_DECLTYPES)
    conn = DatabaseConnection(raw)
    conn.execute("""CREATE TABLE mart.shift_timeline_segment (
        business_date DATE, lob TEXT, agent_id TEXT, agent_name TEXT,
        team_leader TEXT, segment_start TIMESTAMP, segment_minutes INTEGER,
        actual_status TEXT, actual_category TEXT)""")
    conn.execute("""CREATE TABLE mart.attendance_agent_day (
        business_date DATE, lob TEXT, agent_id TEXT, agent_name TEXT,
        team_leader TEXT, scheduled_minutes INTEGER, scheduled_start TIMESTAMP,
        scheduled_end TIMESTAMP, attendance_result TEXT, assignment_type TEXT)""")
    rows = [
        ("a", "Disponible", "Productive", 60, 8),
        ("a", "Backoffice", "Productive", 60, 9),
        ("a", "Training", "Auxiliary", 60, 10),
        ("a", "", "NO_STATUS_EVIDENCE", 60, 11),
        ("b", "", "PTO", 240, 8),
    ]
    for agent, status, category, minutes, hour in rows:
        conn.execute(
            "INSERT INTO mart.shift_timeline_segment VALUES (?,?,?,?,?,?,?,?,?)",
            (DAY, "RSA NL", agent, agent, "TL", datetime(2026, 9, 15, hour), minutes, status, category),
        )
    for agent, minutes, result in (("a", 240, "Present"), ("b", 240, "PTO"), ("c", 240, "PTO"), ("d", 60, "Missing actual evidence")):
        conn.execute(
            "INSERT INTO mart.attendance_agent_day VALUES (?,?,?,?,?,?,?,?,?,?)",
            (DAY, "RSA NL", agent, agent, "TL", minutes, None, None, result, "Working"),
        )
    headers, agent_rows, lob_rows = _realisations_hours(conn, [PROFILE], DAY, DAY, RULEBOOK)
    agent_a = dict(zip(headers, agent_rows[0]))
    total = dict(zip(headers, lob_rows[0]))
    assert agent_a["scheduled_hours"] == 4
    assert agent_a["available_hours"] == 1
    assert agent_a["bo_hours"] == 1
    assert agent_a["training_hours"] == 1
    assert agent_a["unknown_hours"] == 1
    assert total["scheduled_hours"] == 13
    assert total["pto_hours"] == 8
    assert total["unknown_hours"] == 2
    assert total["allocated_hours"] == 13
    assert total["reconciliation_delta_hours"] == 0
    assert len(HOUR_BUCKETS) == 18
    conn.close()


def test_live_not_seen_is_a_callout_but_not_confirmed_no_show():
    today = date.today()
    raw = sqlite3.connect(":memory:", detect_types=sqlite3.PARSE_DECLTYPES)
    conn = DatabaseConnection(raw)
    conn.execute("""CREATE TABLE mart.attendance_agent_day (
        business_date DATE, agent_id TEXT, agent_name TEXT, team_leader TEXT,
        ops_manager TEXT, lob TEXT, language TEXT, scheduled_start TIMESTAMP,
        scheduled_end TIMESTAMP, assignment_type TEXT, attendance_result TEXT,
        call_action TEXT, requires_call BOOLEAN, actual_first_seen TIMESTAMP,
        actual_last_seen TIMESTAMP, uncoded_late_minutes INTEGER,
        actual_evidence TEXT, source_loaded BOOLEAN, is_provisional BOOLEAN,
        evaluation_as_of TIMESTAMP, planning_overlay TEXT,
        planning_overlay_minutes INTEGER)""")
    conn.execute("""CREATE TABLE mart.shift_timeline_segment (
        business_date DATE, agent_id TEXT, lob TEXT, segment_start TIMESTAMP,
        segment_end TIMESTAMP, planned_state TEXT, actual_status TEXT,
        actual_category TEXT, mismatch_type TEXT, is_gap BOOLEAN,
        observed_source TEXT)""")
    checkpoint = datetime.combine(today, time(12))
    for agent, start_hour, end_hour, result, provisional, action in (
        ("not-seen", 8, 16, "Not seen - shift in progress", True, "CALL_NOT_SEEN_NOW"),
        ("confirmed", 6, 10, "No show", False, "CALL_NO_SHOW"),
    ):
        start = datetime.combine(today, time(start_hour))
        end = datetime.combine(today, time(end_hour))
        conn.execute(
            "INSERT INTO mart.attendance_agent_day VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
            (today, agent, agent, "TL", "OM", "RSA NL", "NL", start, end,
             "Working", result, action, True, None, None, 0,
             "AGENT_STATUS", True, provisional, checkpoint, None, 0),
        )
        conn.execute(
            "INSERT INTO mart.shift_timeline_segment VALUES (?,?,?,?,?,?,?,?,?,?,?)",
            (today, agent, "RSA NL", start, min(checkpoint, end), "WORK",
             "Logged Off", "Logged Off", "LOGGED_OFF", True, "AGENT_STATUS"),
        )
    profile = load_service_profiles(ROOT, ROOT / "config/default_service_profiles.toml").select("rsa_nl", today)
    rows, summary = _attendance_pulse(conn, profile, today)
    states = {row["agent_id"]: row for row in rows}
    assert summary["no_show_hc"] == 1
    assert summary["unknown_hc"] == 1
    assert summary["call_now"] == 2
    assert states["not-seen"]["attendance_state"] == "UNKNOWN — POSSIBLE NO SHOW"
    assert states["not-seen"]["pulse_action"] == "CALL_NOT_SEEN_NOW"
    assert states["confirmed"]["attendance_state"] == "NO SHOW"
    conn.close()


def test_rtm_carries_yesterdays_overnight_shift_into_today():
    raw = sqlite3.connect(":memory:", detect_types=sqlite3.PARSE_DECLTYPES)
    conn = DatabaseConnection(raw)
    conn.execute("""CREATE TABLE mart.attendance_agent_day (
        business_date DATE, agent_id TEXT, agent_name TEXT, team_leader TEXT,
        ops_manager TEXT, lob TEXT, language TEXT, scheduled_start TIMESTAMP,
        scheduled_end TIMESTAMP, assignment_type TEXT, attendance_result TEXT,
        call_action TEXT, requires_call BOOLEAN, actual_first_seen TIMESTAMP,
        actual_last_seen TIMESTAMP, uncoded_late_minutes INTEGER,
        actual_evidence TEXT, source_loaded BOOLEAN, is_provisional BOOLEAN,
        evaluation_as_of TIMESTAMP, planning_overlay TEXT,
        planning_overlay_minutes INTEGER)""")
    conn.execute("""CREATE TABLE mart.shift_timeline_segment (
        business_date DATE, agent_id TEXT, lob TEXT, segment_start TIMESTAMP,
        segment_end TIMESTAMP, planned_state TEXT, actual_status TEXT,
        actual_category TEXT, mismatch_type TEXT, is_gap BOOLEAN,
        observed_source TEXT)""")
    prior = DAY - timedelta(days=1)
    start = datetime.combine(prior, time(22))
    end = datetime.combine(DAY, time(2))
    conn.execute(
        "INSERT INTO mart.attendance_agent_day VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
        (prior, "night", "Night Agent", "TL", "OM", "RSA NL", "NL", start, end,
         "Working", "Present", "NONE", False, start, end, 0,
         "AGENT_STATUS", True, False, end, None, 0),
    )
    conn.execute(
        "INSERT INTO mart.shift_timeline_segment VALUES (?,?,?,?,?,?,?,?,?,?,?)",
        (prior, "night", "RSA NL", start, end, "WORK", "Disponible",
         "Productive", "MATCH", False, "AGENT_STATUS"),
    )
    rows, pulse = _attendance_pulse(conn, PROFILE, DAY)
    assert len(rows) == 1
    assert rows[0]["agent_id"] == "night"
    hourly = _workforce_by_hour(PROFILE, DAY, rows, pulse, RULEBOOK)
    assert hourly[1]["scheduled_hc"] == 1
    assert hourly[1]["logged_hc"] == 1
    assert hourly[1]["available_hc"] == 1
    conn.close()


def test_realisations_aux_and_final_absence_separate_unpaid_leave():
    raw = sqlite3.connect(":memory:", detect_types=sqlite3.PARSE_DECLTYPES)
    conn = DatabaseConnection(raw)
    conn.execute("""CREATE TABLE mart.shift_timeline_segment (
        business_date DATE, lob TEXT, agent_id TEXT, agent_name TEXT,
        team_leader TEXT, actual_status TEXT, actual_category TEXT,
        segment_minutes INTEGER, observed_source TEXT, segment_start TIMESTAMP)""")
    conn.execute("""CREATE TABLE mart.verint_final_absence_event (
        agent_day_key TEXT, business_date DATE, lob TEXT, category TEXT,
        counts_as_absence BOOLEAN, event_start TIMESTAMP, event_end TIMESTAMP)""")
    conn.execute("""CREATE TABLE mart.verint_final_absence_agent_day (
        agent_day_key TEXT, business_date DATE, lob TEXT, team_leader TEXT,
        agent_id TEXT, agent_name TEXT, planned_net_minutes INTEGER,
        final_vacation_minutes INTEGER, final_shrinkage_minutes INTEGER,
        final_ledger_status TEXT)""")
    for agent, status, category, minutes, hour in (
        ("a", "Training", "Auxiliary", 60, 8),
        ("a", "Backoffice", "Productive", 30, 9),
        ("b", "Break", "Break", 15, 8),
    ):
        conn.execute(
            "INSERT INTO mart.shift_timeline_segment VALUES (?,?,?,?,?,?,?,?,?,?)",
            (DAY, "RSA NL", agent, agent, "TL", status, category, minutes,
             "AGENT_STATUS", datetime.combine(DAY, time(hour))),
        )
    for agent in ("a", "b"):
        conn.execute(
            "INSERT INTO mart.verint_final_absence_agent_day VALUES (?,?,?,?,?,?,?,?,?,?)",
            (f"20260915-{agent}", DAY, "RSA NL", "TL", agent, agent,
             480, 0, 0, "ABSENCE_RECORDED"),
        )
    for key, category, start_hour, end_hour in (
        ("a", "UNPAID_LEAVE", 8, 10),
        ("b", "NO_SHOW", 8, 10),
        ("b", "NO_SHOW", 8, 10),
    ):
        conn.execute(
            "INSERT INTO mart.verint_final_absence_event VALUES (?,?,?,?,?,?,?)",
            (f"20260915-{key}", DAY, "RSA NL", category, True,
             datetime.combine(DAY, time(start_hour)),
             datetime.combine(DAY, time(end_hour))),
        )
    agent_headers, aux_agent, tl_headers, aux_tl = _realisations_aux_breakdown(
        conn, [PROFILE], DAY, DAY, RULEBOOK,
    )
    assert len(aux_agent) == 3
    assert sum(dict(zip(agent_headers, row))["hours"] for row in aux_agent) == 1.75
    assert sum(dict(zip(tl_headers, row))["hours"] for row in aux_tl) == 1.75
    daily, absence_headers, absence_agents, _lob_headers, absence_lobs = (
        _realisations_final_absence(conn, [PROFILE], DAY, DAY)
    )
    by_agent = {row[4]: dict(zip(absence_headers, row)) for row in absence_agents}
    assert by_agent["a"]["absence_hours"] == 0
    assert by_agent["a"]["unpaid_leave_hours"] == 2
    assert by_agent["b"]["absence_hours"] == 2
    assert daily[(DAY.isoformat(), PROFILE.label)][1] == 2
    assert len(absence_lobs) == 1
    conn.close()


def test_queue_hour_driver_lists_named_agents_as_context_not_causality():
    queue = PROFILE.flash_queues[0]
    source = [{
        "business_date": DAY, "hour_start": datetime(2026, 9, 15, 9),
        "queue": queue, "offered": 10, "answered": 5,
        "abandoned": 5, "short_abandoned": 0,
        "abandoned_within_target": 0, "answered_within_target": 2,
        "handled_seconds": 500,
    }]
    person = {
        "agent_id": "a", "agent_name": "Synthetic Agent",
        "expected_work_intervals": [(datetime(2026, 9, 15, 8), datetime(2026, 9, 15, 16))],
        "no_show": False,
        "timeline_segments": [{**_segment(9, 10, "Logged Off", "Logged Off"),
                               "is_gap": True, "mismatch_type": "LOGGED_OFF"}],
    }
    training = {
        "agent_id": "b", "agent_name": "Training Agent",
        "expected_work_intervals": [(datetime(2026, 9, 15, 8), datetime(2026, 9, 15, 16))],
        "no_show": False,
        "timeline_segments": [{**_segment(9, 10, "Auxiliary", "Training"),
                               "is_gap": False, "mismatch_type": "MATCH"}],
    }
    metrics = load_metric_catalog(ROOT, ROOT / "config/default_metrics.toml")
    headers, rows = _issues_and_drivers_rows(
        [PROFILE], {PROFILE.profile_id: source}, {PROFILE.profile_id: []},
        metrics, DAY, {PROFILE.profile_id: 9},
        {PROFILE.profile_id: {"agent_rows": 1, "unknown_hc": 0}},
        {PROFILE.profile_id: [person, training]}, RULEBOOK,
    )
    drivers = [dict(zip(headers, row)) for row in rows if row[0] == "QUEUE DRIVER"]
    assert len(drivers) == 1
    assert drivers[0]["impacted_hour"] == datetime(2026, 9, 15, 9)
    assert "Synthetic Agent" in drivers[0]["agent_names"]
    assert "Training Agent" in drivers[0]["agent_names"]


def test_storm_service_level_uses_summed_call_components():
    metrics = load_metric_catalog(ROOT, ROOT / "config/default_metrics.toml")
    method = metrics.method_for(
        "service_level", DAY,
        {"lob": "RSA NL", "source_system": "CALL_BY_CALL"},
    )
    outcome = evaluate_metric(method, {
        "offered": 278, "answered": 266, "abandoned": 12,
        "short_abandoned": 2, "abandoned_within_target": 2,
        "answered_within_target": 249, "handled_seconds": 0,
    })
    assert abs(outcome.value - 249 / (278 - 2)) < 1e-10
