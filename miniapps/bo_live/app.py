#!/usr/bin/env python3
"""Run BO Live or publish a PAD DataTable. No main WFMHub imports required."""

from __future__ import annotations

import argparse
import json
import sys
import tempfile
import uuid
from pathlib import Path

# Explicit bootstrap also works with CPython's isolated Windows embeddable runtime.
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from bo_live.core import ContractError, Store, capture_from_csv, seed_demo, stamp, utcnow, write_json_atomic
from bo_live.server import serve


def main(argv=None):
    parser = argparse.ArgumentParser(description="WFMHub BO Live — standalone local TOYOTA control")
    commands = parser.add_subparsers(dest="command", required=True)
    commands.add_parser("clock", help="Print a UTC capture timestamp for PAD")
    run = commands.add_parser("serve", help="Open the local dashboard and watch the PAD inbox")
    run.add_argument("--demo", action="store_true", help="Use synthetic data in a disposable separate data folder")
    run.add_argument("--port", type=int, default=8787)
    run.add_argument("--no-browser", action="store_true")
    publish = commands.add_parser("publish", help="Validate and atomically publish a complete PAD-extracted CSV")
    publish.add_argument("--csv", required=True, type=Path)
    publish.add_argument("--group", required=True)
    publish.add_argument("--source-total", required=True, type=int)
    publish.add_argument("--started-at", required=True, help="Capture start, ISO timestamp including UTC offset")
    publish.add_argument("--finished-at", help="Capture finish, ISO timestamp including UTC offset; defaults to now")
    publish.add_argument("--scope-verified", action="store_true", required=True,
                         help="Assert that exact group, filters, total meaning, and complete table traversal were verified")
    failed = commands.add_parser("record-failure", help="Publish collection health without replacing the last good rows")
    failed.add_argument("--group", required=True)
    failed.add_argument("--message", required=True)
    for sub in (run, publish, failed):
        sub.add_argument("--home", type=Path, default=Path(__file__).resolve().parent / "_local",
                         help="BO Live data folder; never the main WFMHub database folder")
    args = parser.parse_args(argv)
    if args.command == "clock":
        print(stamp(utcnow()))
        return 0
    temporary = None
    try:
        if args.command == "serve" and args.demo:
            temporary = tempfile.TemporaryDirectory(prefix="wfmhub-bo-demo-")
            store = Store(temporary.name, demo=True)
            seed_demo(store)
        else:
            store = Store(args.home)
        if args.command == "serve":
            if not 0 <= args.port <= 65535:
                raise ContractError("Port must be between 0 and 65535.")
            serve(store, args.port, open_browser=not args.no_browser)
        else:
            now = stamp(utcnow())
            if args.command == "publish":
                payload = capture_from_csv(args.csv, args.group, args.source_total, args.started_at,
                                           args.finished_at or now, store.settings)
            else:
                payload = {"schema_version": 1, "capture_id": uuid.uuid4().hex, "task_group": args.group,
                           "started_at": now, "finished_at": now, "status": "failed", "rows": [], "error": args.message}
                from bo_live.core import validate_capture
                validate_capture(payload, store.settings)
            destination = store.inbox / (payload["capture_id"] + ".json")
            write_json_atomic(destination, payload)
            print(json.dumps({"status": "published", "capture_id": payload["capture_id"], "path": str(destination)}))
        return 0
    except (ContractError, ValueError, OSError) as exc:
        print(f"BO Live: {exc}", file=sys.stderr)
        return 1
    finally:
        if temporary:
            temporary.cleanup()


if __name__ == "__main__":
    raise SystemExit(main())
