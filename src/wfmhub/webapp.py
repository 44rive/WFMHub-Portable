"""Portable localhost server for the WFMHub Manager Workbench."""

from __future__ import annotations

import csv
import io
import json
import mimetypes
import subprocess
import sys
import threading
import uuid
import webbrowser
from collections import deque
from dataclasses import dataclass, field
from datetime import datetime
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Any
from urllib.parse import parse_qs, unquote, urlsplit

from . import __version__
from .config import Config, load_config
from .database import connect
from .on_demand_analysis import build_analysis_workbook
from .report_packs import REPORT_PACKS, build_report_pack
from .web_data import DashboardData, DashboardFilter


WEB_ROOT = Path(__file__).with_name("web")
VIEW_METHODS = {
    "desk": "manager_desk",
    "demand": "demand",
    "capacity": "staffing",
    "scenario": "scenario",
    "service": "service",
    "pulse": "today",
    "integrity": "attendance",
    "realisations": "realisations",
    "absence": "absence",
    "patterns": "patterns",
    "readiness": "readiness",
    "mappings": "mappings",
}
EXPORT_ROWS = {
    "desk": "work_queue",
    "demand": "ledger",
    "capacity": "actions",
    "scenario": "actions",
    "service": "queues",
    "pulse": "attendance_actions",
    "integrity": "gaps",
    "realisations": "by_lob",
    "absence": "days",
    "patterns": "patterns",
    "readiness": "issues",
    "mappings": "capacity_mappings",
}

REPORT_ACTION_PACKS = ("service", "staffing", "corrections", "realisations", "absence", "bonus")
ANALYSIS_ACTION_DOMAINS = ("service", "forecast", "staffing", "attendance", "absence")


def _json_default(value: Any):
    if hasattr(value, "isoformat"):
        return value.isoformat()
    raise TypeError(f"Cannot encode {type(value).__name__}")


@dataclass
class RefreshState:
    lock: threading.Lock = field(default_factory=threading.Lock)
    status: str = "IDLE"
    started_at: str | None = None
    finished_at: str | None = None
    return_code: int | None = None
    lines: deque[str] = field(default_factory=lambda: deque(maxlen=30))

    def snapshot(self) -> dict[str, Any]:
        with self.lock:
            return {
                "status": self.status,
                "started_at": self.started_at,
                "finished_at": self.finished_at,
                "return_code": self.return_code,
                "lines": list(self.lines),
            }

    def start(self, config: Config) -> bool:
        with self.lock:
            if self.status == "RUNNING":
                return False
            self.status = "RUNNING"
            self.started_at = datetime.now().isoformat(timespec="seconds")
            self.finished_at = None
            self.return_code = None
            self.lines.clear()
            self.lines.append("Starting governed WFMHub update…")
        threading.Thread(
            target=self._run, args=(config,), daemon=True, name="wfmhub-refresh",
        ).start()
        return True

    def _run(self, config: Config) -> None:
        config.logs.mkdir(parents=True, exist_ok=True)
        log = config.logs / f"web_update_{datetime.now():%Y%m%d_%H%M%S}.log"
        command = [
            sys.executable, "-m", "wfmhub", "--home", str(config.home),
            "refresh", "--source-group", "all", "--no-report",
        ]
        return_code = 1
        try:
            process = subprocess.Popen(
                command, cwd=config.home, stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT, text=True, encoding="utf-8",
                errors="replace", bufsize=1,
            )
            with log.open("w", encoding="utf-8") as handle:
                assert process.stdout is not None
                for line in process.stdout:
                    clean = line.replace("\r", "").strip()
                    handle.write(line)
                    if clean:
                        with self.lock:
                            self.lines.append(clean[-500:])
            return_code = process.wait()
        except Exception as exc:  # pragma: no cover - operating-system failure
            with self.lock:
                self.lines.append(f"Update failed to start: {exc}")
        with self.lock:
            self.return_code = return_code
            self.status = "SUCCESS" if return_code == 0 else "ERROR"
            self.finished_at = datetime.now().isoformat(timespec="seconds")
            self.lines.append(f"Update {self.status.lower()}. Log: {log}")


@dataclass
class ActionState:
    """One safe background report or analysis job for the local operator."""

    lock: threading.Lock = field(default_factory=threading.Lock)
    status: str = "IDLE"
    job_id: str | None = None
    kind: str | None = None
    label: str | None = None
    started_at: str | None = None
    finished_at: str | None = None
    output: str | None = None
    error: str | None = None
    lines: deque[str] = field(default_factory=lambda: deque(maxlen=30))
    history: deque[dict[str, Any]] = field(default_factory=lambda: deque(maxlen=20))

    def _current(self) -> dict[str, Any]:
        return {
            "status": self.status, "job_id": self.job_id,
            "kind": self.kind, "label": self.label,
            "started_at": self.started_at, "finished_at": self.finished_at,
            "output": self.output, "error": self.error,
            "lines": list(self.lines),
        }

    def snapshot(self) -> dict[str, Any]:
        with self.lock:
            return {**self._current(), "history": list(self.history)}

    def start(
        self,
        config: Config,
        kind: str,
        label: str,
        scope: DashboardFilter,
        options: dict[str, Any],
    ) -> bool:
        with self.lock:
            if self.status == "RUNNING":
                return False
            self.status = "RUNNING"
            self.job_id = uuid.uuid4().hex[:12]
            self.kind = kind
            self.label = label
            self.started_at = datetime.now().isoformat(timespec="seconds")
            self.finished_at = None
            self.output = None
            self.error = None
            self.lines.clear()
            self.lines.append(f"Starting {label}…")
        threading.Thread(
            target=self._run,
            args=(config, kind, scope, dict(options)),
            daemon=True,
            name=f"wfmhub-{kind}",
        ).start()
        return True

    def _run(
        self,
        config: Config,
        kind: str,
        scope: DashboardFilter,
        options: dict[str, Any],
    ) -> None:
        conn = None
        try:
            conn = connect(config, read_only=True)
            if kind == "report":
                pack = str(options["pack"])
                self._line(f"Reading current governed marts for {scope.start} to {scope.end}")
                output_path = build_report_pack(pack, conn, config, scope.start, scope.end)
            elif kind == "analysis":
                domain = str(options["domain"])
                comparison = str(options["comparison"])
                self._line(f"Comparing {domain} evidence for {scope.start} to {scope.end}")
                output_path = build_analysis_workbook(
                    conn, config, domain, scope.start, scope.end, comparison,
                )
            else:  # pragma: no cover - handler validates the action kind
                raise ValueError(f"Unsupported action kind: {kind}")
            relative = output_path.resolve().relative_to(config.reports.resolve()).as_posix()
            with self.lock:
                self.output = relative
                self.status = "SUCCESS"
                self.finished_at = datetime.now().isoformat(timespec="seconds")
                self.lines.append(f"Ready: {relative}")
                self.history.appendleft(self._current())
        except Exception as exc:
            with self.lock:
                self.status = "ERROR"
                self.error = str(exc)
                self.finished_at = datetime.now().isoformat(timespec="seconds")
                self.lines.append(f"Failed: {exc}")
                self.history.appendleft(self._current())
        finally:
            if conn is not None:
                conn.close()

    def _line(self, value: str) -> None:
        with self.lock:
            self.lines.append(value)


class ConsoleServer(ThreadingHTTPServer):
    daemon_threads = True

    def __init__(self, address, config: Config):
        super().__init__(address, ConsoleHandler)
        self.config = config
        self.data = DashboardData(config)
        self.refresh_state = RefreshState()
        self.action_state = ActionState()


class ConsoleHandler(BaseHTTPRequestHandler):
    server: ConsoleServer
    protocol_version = "HTTP/1.1"

    def log_message(self, format: str, *args) -> None:
        return

    def _headers(self, content_type: str, length: int, status: int = 200, **extra) -> None:
        self.send_response(status)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(length))
        self.send_header("X-Content-Type-Options", "nosniff")
        self.send_header("X-Frame-Options", "DENY")
        self.send_header("Referrer-Policy", "no-referrer")
        self.send_header(
            "Content-Security-Policy",
            "default-src 'self'; script-src 'self'; style-src 'self'; "
            "img-src 'self' data:; connect-src 'self'; frame-ancestors 'none'",
        )
        for key, value in extra.items():
            self.send_header(key.replace("_", "-"), str(value))
        self.end_headers()

    def _bytes(self, body: bytes, content_type: str, status: int = 200, **headers) -> None:
        self._headers(content_type, len(body), status, **headers)
        self.wfile.write(body)

    def _json(self, value: Any, status: int = 200) -> None:
        body = json.dumps(value, ensure_ascii=False, default=_json_default).encode("utf-8")
        self._bytes(body, "application/json; charset=utf-8", status, Cache_Control="no-store")

    def _error(self, status: int, message: str) -> None:
        self._json({"error": message}, status)

    def _local_request(self) -> bool:
        return self.client_address[0] in {"127.0.0.1", "::1"}

    def _scope(self, query: dict[str, list[str]]) -> DashboardFilter:
        first = lambda name: query.get(name, [None])[0]
        return DashboardFilter.from_values(
            first("start"), first("end"), self.server.data.latest_date(),
            first("management_lob"), first("planning_group"),
            first("staff_type"), first("team_leader"), first("agent_id"),
            first("scenario_fte"),
        )

    def do_GET(self) -> None:
        if not self._local_request():
            self._error(HTTPStatus.FORBIDDEN, "The WFMHub console is localhost-only")
            return
        parsed = urlsplit(self.path)
        try:
            if parsed.path == "/api/meta":
                payload = self.server.data.meta()
                payload["version"] = __version__
                self._json(payload)
                return
            if parsed.path == "/api/refresh":
                self._json(self.server.refresh_state.snapshot())
                return
            if parsed.path == "/api/actions":
                self._json(self.server.action_state.snapshot())
                return
            if parsed.path == "/api/download":
                self._download(parse_qs(parsed.query, keep_blank_values=True))
                return
            if parsed.path.startswith("/api/view/"):
                view = parsed.path.rsplit("/", 1)[-1]
                if view == "reports":
                    self._json(self._reports_payload())
                    return
                if view == "archive":
                    self._json(self._archive_payload())
                    return
                if view == "jobs":
                    self._json(self._jobs_payload())
                    return
                method = VIEW_METHODS.get(view)
                if method is None:
                    self._error(HTTPStatus.NOT_FOUND, "Unknown dashboard view")
                    return
                payload = getattr(self.server.data, method)(
                    self._scope(parse_qs(parsed.query, keep_blank_values=True)),
                )
                payload["view"] = view
                payload["generated_at"] = datetime.now().isoformat(timespec="seconds")
                self._json(payload)
                return
            if parsed.path.startswith("/api/export/"):
                self._export(
                    parsed.path.rsplit("/", 1)[-1],
                    parse_qs(parsed.query, keep_blank_values=True),
                )
                return
            if parsed.path.startswith("/api/"):
                self._error(HTTPStatus.NOT_FOUND, "Unknown API endpoint")
                return
            self._static(parsed.path)
        except (ValueError, FileNotFoundError, RuntimeError) as exc:
            self._error(HTTPStatus.BAD_REQUEST, str(exc))
        except Exception as exc:  # pragma: no cover - defensive boundary
            self._error(HTTPStatus.INTERNAL_SERVER_ERROR, f"Console error: {exc}")

    def do_POST(self) -> None:
        if not self._local_request():
            self._error(HTTPStatus.FORBIDDEN, "The WFMHub console is localhost-only")
            return
        path = urlsplit(self.path).path
        if path == "/api/actions/report":
            self._start_action("report")
            return
        if path == "/api/actions/analysis":
            self._start_action("analysis")
            return
        if path != "/api/refresh":
            self._error(HTTPStatus.NOT_FOUND, "Unknown action")
            return
        if (
            self.headers.get("Content-Type", "").split(";", 1)[0].strip().lower()
            != "application/json"
            or self.headers.get("X-WFMHub-Action") != "refresh"
        ):
            self._error(HTTPStatus.UNSUPPORTED_MEDIA_TYPE, "Invalid local action request")
            return
        length = int(self.headers.get("Content-Length", "0") or 0)
        if length > 4096:
            self.close_connection = True
            self._error(HTTPStatus.REQUEST_ENTITY_TOO_LARGE, "Action request is too large")
            return
        if length:
            self.rfile.read(min(length, 4096))
        started = self.server.refresh_state.start(self.server.config)
        self._json(
            {**self.server.refresh_state.snapshot(), "started": started},
            HTTPStatus.ACCEPTED if started else HTTPStatus.CONFLICT,
        )

    def _read_action_json(self, expected_action: str) -> dict[str, Any]:
        if (
            self.headers.get("Content-Type", "").split(";", 1)[0].strip().lower()
            != "application/json"
            or self.headers.get("X-WFMHub-Action") != expected_action
        ):
            raise ValueError("Invalid local action request")
        length = int(self.headers.get("Content-Length", "0") or 0)
        if length > 16384:
            self.close_connection = True
            raise ValueError("Action request is too large")
        raw = self.rfile.read(length) if length else b"{}"
        try:
            payload = json.loads(raw.decode("utf-8"))
        except (UnicodeDecodeError, json.JSONDecodeError) as exc:
            raise ValueError("Action body must be valid JSON") from exc
        if not isinstance(payload, dict):
            raise ValueError("Action body must be a JSON object")
        return payload

    def _start_action(self, kind: str) -> None:
        try:
            payload = self._read_action_json(kind)
            scope = DashboardFilter.from_values(
                payload.get("start"), payload.get("end"),
                self.server.data.latest_date(),
                payload.get("management_lob"), payload.get("planning_group"),
                payload.get("staff_type"), payload.get("team_leader"),
                payload.get("agent_id"),
            )
            if kind == "report":
                key = str(payload.get("pack") or "").strip()
                if key not in REPORT_ACTION_PACKS:
                    raise ValueError(
                        f"Report must be one of: {', '.join(REPORT_ACTION_PACKS)}"
                    )
                label = REPORT_PACKS[key].current_filename
                options = {"pack": key}
            else:
                domain = str(payload.get("domain") or "").strip()
                comparison = str(payload.get("comparison") or "previous_equal").strip()
                if domain not in ANALYSIS_ACTION_DOMAINS:
                    raise ValueError(
                        f"Analysis domain must be one of: {', '.join(ANALYSIS_ACTION_DOMAINS)}"
                    )
                if comparison not in {"previous_equal", "previous_month", "target", "none"}:
                    raise ValueError("Unknown analysis comparison")
                label = f"{domain.title()} analysis"
                options = {"domain": domain, "comparison": comparison}
            started = self.server.action_state.start(
                self.server.config, kind, label, scope, options,
            )
            self._json(
                {**self.server.action_state.snapshot(), "started": started},
                HTTPStatus.ACCEPTED if started else HTTPStatus.CONFLICT,
            )
        except ValueError as exc:
            self._error(HTTPStatus.BAD_REQUEST, str(exc))

    def _report_files(self) -> list[dict[str, Any]]:
        root = self.server.config.reports.resolve()
        if not root.exists():
            return []
        output = []
        for path in root.rglob("*"):
            if not path.is_file() or path.suffix.lower() not in {".xlsx", ".csv"}:
                continue
            if ".partial" in path.stem.lower():
                continue
            stat = path.stat()
            relative = path.resolve().relative_to(root).as_posix()
            output.append({
                "name": path.name, "relative": relative,
                "folder": path.parent.resolve().relative_to(root).as_posix() or ".",
                "modified_at": datetime.fromtimestamp(stat.st_mtime).isoformat(timespec="seconds"),
                "size_bytes": stat.st_size,
                "archived": "Archive" in path.relative_to(root).parts,
            })
        output.sort(key=lambda row: row["modified_at"], reverse=True)
        return output

    def _reports_payload(self) -> dict[str, Any]:
        files = self._report_files()
        current = [row for row in files if not row["archived"]]
        last = current[0] if current else None
        reports = [
            {
                "key": key, "name": REPORT_PACKS[key].current_filename,
                "purpose": REPORT_PACKS[key].purpose,
            }
            for key in REPORT_ACTION_PACKS
        ]
        return {
            "view": "reports", "generated_at": datetime.now().isoformat(timespec="seconds"),
            "cards": [
                {"label": "Available products", "value": len(reports), "kind": "integer", "note": "PCS remains outside this console"},
                {"label": "Current outputs", "value": len(current), "kind": "integer", "note": "Reports folder"},
                {"label": "Latest output", "value": last["modified_at"] if last else None, "kind": "datetime", "note": last["name"] if last else "No generated file"},
                {"label": "Background action", "value": self.server.action_state.snapshot()["status"], "kind": "text", "note": "Report or analysis job"},
            ],
            "reports": reports,
            "analysis_domains": list(ANALYSIS_ACTION_DOMAINS),
            "comparison_modes": ["previous_equal", "previous_month", "target", "none"],
            "outputs": current[:50], "empty": False,
        }

    def _archive_payload(self) -> dict[str, Any]:
        files = self._report_files()
        archived = [row for row in files if row["archived"]]
        total_bytes = sum(row["size_bytes"] for row in files)
        return {
            "view": "archive", "generated_at": datetime.now().isoformat(timespec="seconds"),
            "cards": [
                {"label": "All outputs", "value": len(files), "kind": "integer", "note": "Current and archived"},
                {"label": "Archived versions", "value": len(archived), "kind": "integer", "note": "Timestamped prior copies"},
                {"label": "Disk used", "value": total_bytes, "kind": "bytes", "note": "Report files only"},
                {"label": "Latest generated", "value": files[0]["modified_at"] if files else None, "kind": "datetime", "note": files[0]["name"] if files else "No generated file"},
            ],
            "outputs": files[:500], "empty": not files,
        }

    def _jobs_payload(self) -> dict[str, Any]:
        action = self.server.action_state.snapshot()
        refresh = self.server.refresh_state.snapshot()
        logs = []
        if self.server.config.logs.exists():
            for path in sorted(
                self.server.config.logs.glob("*.log"),
                key=lambda item: item.stat().st_mtime, reverse=True,
            )[:30]:
                stat = path.stat()
                logs.append({
                    "name": path.name,
                    "modified_at": datetime.fromtimestamp(stat.st_mtime).isoformat(timespec="seconds"),
                    "size_bytes": stat.st_size,
                })
        jobs = []
        # Completed actions are already copied into history. Only prepend the
        # live record while it is actually running to avoid a duplicate row.
        if action["status"] == "RUNNING":
            jobs.append(action)
        jobs.extend(action["history"])
        if refresh["status"] != "IDLE":
            jobs.append({
                "status": refresh["status"], "job_id": "refresh",
                "kind": "refresh", "label": "Update all data",
                "started_at": refresh["started_at"],
                "finished_at": refresh["finished_at"],
                "output": None, "error": None,
                "lines": refresh["lines"],
            })
        running = sum(1 for row in jobs if row.get("status") == "RUNNING")
        failures = sum(1 for row in jobs if row.get("status") == "ERROR")
        return {
            "view": "jobs", "generated_at": datetime.now().isoformat(timespec="seconds"),
            "cards": [
                {"label": "Running", "value": running, "kind": "integer", "note": "Local background jobs"},
                {"label": "Recent completed", "value": sum(1 for row in jobs if row.get("status") == "SUCCESS"), "kind": "integer", "note": "Current console session"},
                {"label": "Recent failures", "value": failures, "kind": "integer", "note": "Inspect exact message"},
                {"label": "Log files", "value": len(logs), "kind": "integer", "note": "Latest 30 shown"},
            ],
            "jobs": jobs, "logs": logs, "empty": not jobs and not logs,
        }

    def _download(self, query: dict[str, list[str]]) -> None:
        raw = query.get("file", [""])[0]
        relative = unquote(str(raw or "")).replace("\\", "/")
        root = self.server.config.reports.resolve()
        path = (root / relative).resolve()
        if path != root and root not in path.parents:
            self._error(HTTPStatus.NOT_FOUND, "Report file not found")
            return
        if not path.is_file() or path.suffix.lower() not in {".xlsx", ".csv"}:
            self._error(HTTPStatus.NOT_FOUND, "Report file not found")
            return
        mime = mimetypes.guess_type(path.name)[0] or "application/octet-stream"
        self._bytes(
            path.read_bytes(), mime, Cache_Control="no-store",
            Content_Disposition=f'attachment; filename="{path.name.replace(chr(34), "")}"',
        )

    def _export(self, view: str, query: dict[str, list[str]]) -> None:
        method = VIEW_METHODS.get(view)
        field = EXPORT_ROWS.get(view)
        if method is None or field is None:
            self._error(HTTPStatus.NOT_FOUND, "Unknown export view")
            return
        payload = getattr(self.server.data, method)(self._scope(query))
        rows = list(payload.get(field) or [])
        output = io.StringIO(newline="")
        if rows:
            headers = list(rows[0])
            writer = csv.DictWriter(output, fieldnames=headers, extrasaction="ignore")
            writer.writeheader()
            for row in rows:
                writer.writerow({key: _json_default(value) if hasattr(value, "isoformat") else value for key, value in row.items()})
        body = ("\ufeff" + output.getvalue()).encode("utf-8")
        filename = f"WFMHub-{view}-{datetime.now():%Y%m%d-%H%M}.csv"
        self._bytes(
            body, "text/csv; charset=utf-8", Cache_Control="no-store",
            Content_Disposition=f'attachment; filename="{filename}"',
        )

    def _static(self, request_path: str) -> None:
        relative = "index.html" if request_path in {"", "/"} else request_path.lstrip("/")
        if ".." in Path(relative).parts:
            self._error(HTTPStatus.NOT_FOUND, "Not found")
            return
        path = WEB_ROOT / relative
        if not path.is_file():
            path = WEB_ROOT / "index.html"
        body = path.read_bytes()
        mime = mimetypes.guess_type(path.name)[0] or "application/octet-stream"
        cache = "no-cache" if path.name == "index.html" else "public, max-age=3600"
        self._bytes(body, f"{mime}; charset=utf-8" if mime.startswith("text/") else mime, Cache_Control=cache)


def run_console(
    home: Path,
    port: int = 8765,
    *,
    launch: bool = True,
) -> int:
    """Start the localhost console and block until the operator closes it."""

    config = load_config(home)
    if not config.database.is_file():
        raise FileNotFoundError(
            "The WFMHub database does not exist yet. Run SETUP, then UPDATE once."
        )
    server = None
    for candidate in range(port, port + 20):
        try:
            server = ConsoleServer(("127.0.0.1", candidate), config)
            port = candidate
            break
        except OSError:
            continue
    if server is None:
        raise RuntimeError("No free localhost port was found between 8765 and 8784")
    url = f"http://127.0.0.1:{port}/"
    print("\nWFMHUB MANAGER WORKBENCH")
    print(f"Address : {url}")
    print(f"Database: {config.database}")
    print("Scope   : this computer only")
    print("Close this window or press Ctrl+C to stop the console.\n")
    if launch:
        threading.Timer(0.7, lambda: webbrowser.open(url)).start()
    try:
        server.serve_forever(poll_interval=.25)
    except KeyboardInterrupt:
        print("\nConsole stopped.")
    finally:
        server.server_close()
    return 0
