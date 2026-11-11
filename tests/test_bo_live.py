"""Synthetic BO evidence, failure recovery, publisher, and loopback API contracts."""

from __future__ import annotations

import copy
import csv
import importlib.util
import json
import subprocess
import sys
import tempfile
import threading
import unittest
from datetime import datetime, timedelta, timezone
from pathlib import Path
from urllib.error import HTTPError
from urllib.request import Request, urlopen

from miniapps.bo_live.core import (
    ContractError, DEFAULT_SETTINGS, Store, capture_from_csv, seed_demo,
    source_time, stamp, validate_settings, write_json_atomic,
)
from miniapps.bo_live.server import DashboardServer

ROOT = Path(__file__).resolve().parents[1]
GROUP = "(TOYOTA) Agendas à Traiter"
NOW = datetime(2026, 11, 11, 10, 0, tzinfo=timezone.utc)


def task(index=1, minutes=-30, case_number="000012"):
    return {"row_index": index, "case_number": case_number, "task_type": "Call-back task",
            "priority": "Very high", "language": "French", "scheduled_for": stamp(NOW + timedelta(minutes=minutes)),
            "assigned_to": ""}


def capture(identity="capture-1", rows=None, finish=NOW, group=GROUP):
    rows = [task()] if rows is None else rows
    return {"schema_version": 1, "capture_id": identity, "task_group": group,
            "started_at": stamp(finish - timedelta(seconds=30)), "finished_at": stamp(finish),
            "source_timezone": "UTC", "status": "complete", "source_total": len(rows),
            "scope_verified": True, "total_verified": True, "coverage": "full_table", "rows": rows}


class BOLiveTests(unittest.TestCase):
    def setUp(self):
        self.folder = tempfile.TemporaryDirectory()
        self.addCleanup(self.folder.cleanup)
        self.home = Path(self.folder.name)
        self.store = Store(self.home)
        self.settings = dict(DEFAULT_SETTINGS, source_timezone="UTC", assignment_blank_means_unassigned=True)
        self.store.save_settings(self.settings)

    def ingest(self, payload):
        return self.store.ingest(payload, now=NOW + timedelta(minutes=3))

    def test_fresh_schema_reopens_without_resetting_evidence(self):
        self.ingest(capture())
        reopened = Store(self.home)
        self.assertEqual(reopened.state(now=NOW)["totals"]["total"], 1)
        with reopened.connect() as db:
            self.assertEqual(db.execute("PRAGMA user_version").fetchone()[0], 1)
            self.assertEqual(db.execute("PRAGMA integrity_check").fetchone()[0], "ok")

    def test_repeated_cases_and_identical_tasks_are_separate_occurrences(self):
        self.ingest(capture(rows=[task(1), task(2), task(3)]))
        view = self.store.state(now=NOW)
        self.assertEqual(view["totals"]["total"], 3)
        self.assertEqual([r["case_number"] for r in view["tasks"]], ["000012"] * 3)
        self.assertEqual(view["totals"]["unassigned"], 3)

    def test_idempotence_and_conflicting_capture_id(self):
        payload = capture()
        self.assertFalse(self.ingest(payload)["duplicate"])
        self.assertTrue(self.ingest(payload)["duplicate"])
        changed = copy.deepcopy(payload)
        changed["rows"][0]["assigned_to"] = "Synthetic operator"
        with self.assertRaisesRegex(ContractError, "different evidence"):
            self.ingest(changed)
        self.assertEqual(len(self.store.state(now=NOW)["history"]), 1)

    def test_failed_attempt_retains_last_good_rows_and_changes_health(self):
        self.ingest(capture())
        failed = capture("failure", rows=[], finish=NOW + timedelta(minutes=1))
        failed.update(status="failed", error="Session expired", source_total=None)
        self.ingest(failed)
        view = self.store.state(now=NOW + timedelta(minutes=1))
        self.assertEqual(view["totals"]["total"], 1)
        self.assertEqual(view["groups"][0]["health"], "failed")
        self.assertEqual(view["tasks"][0]["observed_at"], stamp(NOW))
        self.assertFalse(view["all_fresh"])

    def test_empty_complete_capture_replaces_good_rows_with_true_zero(self):
        self.ingest(capture())
        self.ingest(capture("empty", rows=[], finish=NOW + timedelta(minutes=1)))
        self.assertEqual(self.store.state(now=NOW + timedelta(minutes=1))["totals"]["total"], 0)

    def test_older_out_of_order_capture_cannot_replace_newer_evidence(self):
        self.ingest(capture("new", rows=[task(1), task(2)]))
        self.ingest(capture("old", rows=[], finish=NOW - timedelta(minutes=20)))
        self.assertEqual(self.store.state(now=NOW)["totals"]["total"], 2)

    def test_timing_boundaries_use_capture_time_not_view_time(self):
        rows = [task(1, 0), task(2, 60), task(3, 61), task(4)]
        rows[3]["scheduled_for"] = None
        self.ingest(capture(rows=rows))
        view = self.store.state(now=NOW + timedelta(minutes=30))
        self.assertEqual([view["totals"][key] for key in ("past", "upcoming", "future", "unscheduled")], [1, 1, 1, 1])
        self.assertEqual(view["groups"][0]["health"], "stale")

    def test_unknown_and_missing_evidence_are_not_zero(self):
        self.store.save_settings(dict(self.settings, groups=[GROUP, "(TOYOTA) TAKATA"],
                                      assignment_blank_means_unassigned=False))
        self.assertIsNone(self.store.state(now=NOW)["totals"]["total"])
        self.ingest(capture())
        view = self.store.state(now=NOW)
        self.assertEqual(view["fresh_groups"], 1)
        self.assertFalse(view["all_fresh"])
        self.assertIsNone(view["groups"][1]["metrics"]["total"])
        self.assertIsNone(view["totals"]["unassigned"])

    def test_invalid_coverage_cannot_replace_good_evidence(self):
        self.ingest(capture())
        variants = []
        wrong_total = capture("bad-total"); wrong_total["source_total"] = 2; variants.append(wrong_total)
        overlap = capture("overlap", rows=[task(1), task(1)]); variants.append(overlap)
        wrong_scope = capture("scope"); wrong_scope["scope_verified"] = False; variants.append(wrong_scope)
        missing_cell = capture("missing-cell"); del missing_cell["rows"][0]["assigned_to"]; variants.append(missing_cell)
        missing_time = capture("missing-time"); del missing_time["rows"][0]["scheduled_for"]; variants.append(missing_time)
        clipped = capture("clipped"); clipped["rows"][0]["scheduled_for"] = "11/11/2026 10:..."; variants.append(clipped)
        other = capture("other", group="(OTHER) Work"); variants.append(other)
        for payload in variants:
            with self.subTest(payload["capture_id"]), self.assertRaises(ContractError):
                self.ingest(payload)
        self.assertEqual(self.store.state(now=NOW)["totals"]["total"], 1)
        self.assertEqual(len(self.store.state(now=NOW)["history"]), 1)

    def test_failed_capture_cannot_carry_partial_rows(self):
        payload = capture(); payload.update(status="incomplete", error="Traversal stopped")
        with self.assertRaisesRegex(ContractError, "health metadata only"):
            self.ingest(payload)

    def test_configuration_scope_and_timezone_are_guarded(self):
        with self.assertRaises(ContractError):
            validate_settings(dict(self.settings, groups=["(TOYOTA) WEBAPP ..."]))
        with self.assertRaises(ContractError):
            validate_settings(dict(self.settings, groups=[GROUP, GROUP]))
        self.ingest(capture())
        with self.assertRaisesRegex(ContractError, "fixed after"):
            self.store.save_settings(dict(self.settings, source_timezone="Europe/Brussels"))

    def test_timezone_is_required_before_complete_capture(self):
        self.store.save_settings(dict(self.settings, source_timezone=None))
        with self.assertRaisesRegex(ContractError, "Set the source time zone"):
            self.ingest(capture())

    def test_local_times_require_full_values_and_resolve_dst(self):
        self.assertEqual(source_time("11/11/2026 10:30", "Europe/Brussels"), "2026-11-11T09:30:00+00:00")
        for value in ("2026-11-11", "2026-03-29T02:30:00", "2026-10-25T02:30:00"):
            with self.subTest(value), self.assertRaises(ContractError):
                source_time(value, "Europe/Brussels")
        self.assertEqual(source_time("2026-10-25T02:30:00+01:00", "Europe/Brussels"), "2026-10-25T01:30:00+00:00")

    def test_inbox_only_reads_published_json_and_retains_rejections(self):
        write_json_atomic(self.store.inbox / "complete.json", capture())
        (self.store.inbox / "unfinished.json.tmp").write_text('{"unfinished":')
        (self.store.inbox / "broken.json").write_text("not JSON")
        results = self.store.scan_inbox()
        self.assertEqual({r["status"] for r in results}, {"accepted", "rejected"})
        self.assertEqual(len(list((self.home / "processed").glob("*.json"))), 1)
        self.assertEqual(len(list((self.home / "rejected").glob("*.json"))), 1)
        self.assertEqual(self.store.state(now=NOW)["inbox_attention"], 1)
        self.store.acknowledge_events()
        self.assertEqual(self.store.state(now=NOW)["inbox_attention"], 0)
        self.assertTrue((self.store.inbox / "unfinished.json.tmp").exists())

    def test_csv_to_inbox_to_store_preserves_case_text_and_duplicates(self):
        path = self.home / "pad.csv"
        with path.open("w", encoding="utf-8", newline="") as handle:
            writer = csv.writer(handle)
            writer.writerow(["Case number", "Task type", "Task priority", "Language", "Scheduled for", "Assigned to"])
            writer.writerow(["000012", "Call-back task", "Very high", "French", "11/11/2026 09:30", ""])
            writer.writerow(["000012", "Call-back task", "Very high", "French", "11/11/2026 09:30", ""])
        payload = capture_from_csv(path, GROUP, 2, stamp(NOW - timedelta(seconds=20)), stamp(NOW), self.settings)
        self.ingest(payload)
        self.assertEqual([r["case_number"] for r in self.store.state(now=NOW)["tasks"]], ["000012", "000012"])
        with self.assertRaisesRegex(ContractError, "row count"):
            capture_from_csv(path, GROUP, 3, stamp(NOW), stamp(NOW), self.settings)

    def test_cli_publisher_writes_atomic_capture_without_writing_task_db(self):
        path = self.home / "empty.csv"
        path.write_text("Case number,Task type,Task priority,Language,Scheduled for,Assigned to\n", encoding="utf-8")
        command = [sys.executable, "-I", str(ROOT / "miniapps/bo_live/app.py"), "publish", "--home", str(self.home),
                   "--csv", str(path), "--group", GROUP, "--source-total", "0", "--started-at", stamp(NOW),
                   "--finished-at", stamp(NOW), "--scope-verified"]
        result = subprocess.run(command, capture_output=True, text=True)
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertEqual(json.loads(result.stdout)["status"], "published")
        self.assertIsNone(self.store.state(now=NOW)["totals"]["total"])
        self.store.scan_inbox()
        self.assertEqual(self.store.state(now=NOW)["totals"]["total"], 0)

    def test_demo_is_separate_and_tagged_captures_do_not_enter_live(self):
        with tempfile.TemporaryDirectory() as demo_home:
            demo = Store(demo_home, demo=True)
            seed_demo(demo, NOW)
            self.assertEqual(demo.state(now=NOW)["totals"]["total"], 42)
            self.assertTrue(demo.state(now=NOW)["demo"])
        payload = capture(); payload["synthetic"] = True
        with self.assertRaisesRegex(ContractError, "demo mode"):
            self.ingest(payload)

    def test_portable_source_allowlist_excludes_user_state(self):
        spec = importlib.util.spec_from_file_location("bo_builder_test", ROOT / "packaging/windows/build_bo_live.py")
        module = importlib.util.module_from_spec(spec); spec.loader.exec_module(module)
        target = self.home / "stage"; target.mkdir()
        module.stage_source(target)
        files = {p.relative_to(target).as_posix() for p in target.rglob("*") if p.is_file()}
        expected = {"bo_live/" + f for f in module.CODE} | set(module.ROOT_FILES) | {"VERSION.txt"}
        self.assertEqual(files, expected)
        self.assertFalse(any("_local" in f or "sqlite" in f for f in files))


class BOApiTests(unittest.TestCase):
    def setUp(self):
        self.folder = tempfile.TemporaryDirectory()
        self.store = Store(self.folder.name)
        self.store.save_settings(dict(DEFAULT_SETTINGS, source_timezone="UTC"))
        self.server = DashboardServer(self.store, port=0)
        self.thread = threading.Thread(target=self.server.serve_forever, daemon=True)
        self.thread.start()
        self.base = f"http://127.0.0.1:{self.server.server_port}"

    def tearDown(self):
        self.server.shutdown(); self.server.server_close(); self.thread.join(timeout=3)
        self.folder.cleanup()

    def request(self, path, payload=None, headers=None):
        headers = dict(headers or {})
        data = None if payload is None else json.dumps(payload).encode()
        if data is not None:
            headers.setdefault("Content-Type", "application/json")
        req = Request(self.base + path, data=data, headers=headers)
        try:
            response = urlopen(req, timeout=5)
        except HTTPError as error:
            response = error
        with response:
            return response.status, response.read(), response.headers

    def test_static_and_api_use_local_headers_and_no_path_traversal(self):
        status, body, headers = self.request("/")
        self.assertEqual(status, 200)
        self.assertIn(b"TOYOTA control room", body)
        self.assertIn("frame-ancestors 'none'", headers["Content-Security-Policy"])
        self.assertEqual(self.request("/../../core.py")[0], 404)
        self.assertEqual(self.request("/api/state", headers={"Host": "foreign.example"})[0], 403)

    def test_mutations_require_csrf_and_same_origin(self):
        self.assertEqual(self.request("/api/import", capture())[0], 403)
        self.assertEqual(self.request("/api/import", capture(), {"X-BO-CSRF": self.server.csrf, "Origin": "https://foreign.example"})[0], 403)
        status, body, _ = self.request("/api/import", capture(), {"X-BO-CSRF": self.server.csrf})
        self.assertEqual(status, 200, body)
        status, body, _ = self.request("/api/state")
        self.assertEqual(status, 200)
        state = json.loads(body)
        self.assertEqual(state["totals"]["total"], 1)
        self.assertEqual(state["csrf"], self.server.csrf)

    def test_invalid_capture_rejects_without_changing_evidence(self):
        invalid = capture(); invalid["source_total"] = 2
        status, body, _ = self.request("/api/import", invalid, {"X-BO-CSRF": self.server.csrf})
        self.assertEqual(status, 400)
        self.assertIn("row count", json.loads(body)["error"])
        self.assertIsNone(self.store.state()["totals"]["total"])

    def test_publisher_inbox_and_dashboard_work_together(self):
        path = Path(self.folder.name) / "pad-table.csv"
        path.write_text("Case number,Task type,Task priority,Language,Scheduled for,Assigned to\n"
                        "000012,Call-back task,Very high,French,11/11/2026 09:30,\n"
                        "000012,Call-back task,High,Dutch,11/11/2026 11:30,Demo operator\n", encoding="utf-8")
        published = subprocess.run([sys.executable, "-I", str(ROOT / "miniapps/bo_live/app.py"), "publish",
                                    "--home", self.folder.name, "--csv", str(path), "--group", GROUP,
                                    "--source-total", "2", "--started-at", stamp(NOW - timedelta(seconds=30)),
                                    "--finished-at", stamp(NOW), "--scope-verified"], capture_output=True, text=True)
        self.assertEqual(published.returncode, 0, published.stderr)
        status, body, _ = self.request("/api/scan", {}, {"X-BO-CSRF": self.server.csrf})
        self.assertEqual(status, 200, body)
        self.assertEqual(json.loads(body)["results"][0]["status"], "accepted")
        status, body, _ = self.request("/api/state")
        view = json.loads(body)
        self.assertEqual(status, 200)
        self.assertEqual(view["totals"]["total"], 2)
        self.assertEqual(view["totals"]["past"], 1)
        self.assertEqual([row["case_number"] for row in view["tasks"]], ["000012", "000012"])


if __name__ == "__main__":
    unittest.main()
