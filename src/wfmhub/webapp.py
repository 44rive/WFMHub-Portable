"""Portable localhost server for the WFMHub operations console."""

from __future__ import annotations

import csv
import io
import json
import mimetypes
import subprocess
import sys
import threading
import webbrowser
from collections import deque
from dataclasses import dataclass, field
from datetime import datetime
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Any
from urllib.parse import parse_qs, urlsplit

from . import __version__
from .config import Config, load_config
from .web_data import DashboardData, DashboardFilter


WEB_ROOT = Path(__file__).with_name("web")
VIEW_METHODS = {
    "today": "today",
    "staffing": "staffing",
    "service": "service",
    "attendance": "attendance",
    "history": "history",
}
EXPORT_ROWS = {
    "today": "attendance_actions",
    "staffing": "actions",
    "service": "queues",
    "attendance": "gaps",
    "history": "patterns",
}


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


class ConsoleServer(ThreadingHTTPServer):
    daemon_threads = True

    def __init__(self, address, config: Config):
        super().__init__(address, ConsoleHandler)
        self.config = config
        self.data = DashboardData(config)
        self.refresh_state = RefreshState()


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
            if parsed.path.startswith("/api/view/"):
                view = parsed.path.rsplit("/", 1)[-1]
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
        if urlsplit(self.path).path != "/api/refresh":
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
    print("\nWFMHUB OPERATIONS CONSOLE")
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
