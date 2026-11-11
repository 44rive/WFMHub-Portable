"""Loopback-only dashboard API. PAD publishes files, not browser mutations."""

from __future__ import annotations

import json
import secrets
import sqlite3
import threading
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import urlsplit

from .core import ContractError, MAX_BYTES

WEB = Path(__file__).with_name("web")
ASSETS = {"/": ("index.html", "text/html; charset=utf-8"),
          "/app.css": ("app.css", "text/css; charset=utf-8"),
          "/app.js": ("app.js", "text/javascript; charset=utf-8")}


class DashboardServer(ThreadingHTTPServer):
    daemon_threads = True
    allow_reuse_address = True

    def __init__(self, store, port=8787):
        self.store = store
        self.csrf = secrets.token_urlsafe(32)
        super().__init__(("127.0.0.1", port), Handler)
        self.origins = {f"http://127.0.0.1:{self.server_port}", f"http://localhost:{self.server_port}"}


class Handler(BaseHTTPRequestHandler):
    server_version = "WFMHubBOLive/0.1"

    def log_message(self, format, *args):
        pass  # No case numbers or request payloads in terminal logs.

    def reply(self, status, content, content_type="application/json; charset=utf-8"):
        if isinstance(content, (dict, list)):
            content = json.dumps(content, ensure_ascii=False, allow_nan=False).encode("utf-8")
        elif isinstance(content, str):
            content = content.encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(content)))
        self.send_header("Cache-Control", "no-store")
        self.send_header("X-Content-Type-Options", "nosniff")
        self.send_header("Referrer-Policy", "no-referrer")
        self.send_header("Content-Security-Policy", "default-src 'self'; script-src 'self'; style-src 'self'; img-src 'self' data:; object-src 'none'; frame-ancestors 'none'; base-uri 'none'")
        self.end_headers()
        self.wfile.write(content)

    def valid_host(self):
        if "http://" + self.headers.get("Host", "") not in self.server.origins:
            self.reply(403, {"error": "Use this app through its local address."})
            return False
        return True

    def do_GET(self):
        if not self.valid_host():
            return
        path = urlsplit(self.path).path
        if path in ASSETS:
            name, content_type = ASSETS[path]
            self.reply(200, (WEB / name).read_bytes(), content_type)
        elif path == "/api/state":
            self.reply(200, dict(self.server.store.state(), csrf=self.server.csrf))
        elif path == "/api/health":
            self.reply(200, {"status": "ok", "demo": self.server.store.demo})
        elif path == "/favicon.ico":
            self.reply(204, b"", "image/x-icon")
        else:
            self.reply(404, {"error": "Not found."})

    def do_POST(self):
        if not self.valid_host():
            return
        origin = self.headers.get("Origin")
        if (origin and origin not in self.server.origins) or not secrets.compare_digest(self.headers.get("X-BO-CSRF", ""), self.server.csrf):
            self.reply(403, {"error": "Reload the local app before making changes."})
            return
        if self.headers.get("Content-Type", "").split(";")[0] != "application/json":
            self.reply(415, {"error": "Use application/json."})
            return
        try:
            length = int(self.headers.get("Content-Length", "0"))
            if not 0 < length <= MAX_BYTES:
                self.reply(413, {"error": "Request must contain JSON of at most 16 MB."})
                return
            value = json.loads(self.rfile.read(length))
            path = urlsplit(self.path).path
            if path == "/api/import":
                result = self.server.store.ingest(value)
            elif path == "/api/settings":
                result = self.server.store.save_settings(value)
            elif path == "/api/scan":
                result = {"results": self.server.store.scan_inbox()}
            elif path == "/api/acknowledge":
                result = self.server.store.acknowledge_events()
            else:
                self.reply(404, {"error": "Not found."})
                return
            self.reply(200, result)
        except ContractError as exc:
            self.reply(400, {"error": str(exc)})
        except (ValueError, UnicodeError):
            self.reply(400, {"error": "Request is not valid JSON."})
        except (OSError, sqlite3.Error):
            self.reply(500, {"error": "Local storage could not be updated. Check folder permissions and disk space."})


def watch_inbox(store, stop):
    while not stop.is_set():
        try:
            store.scan_inbox()
            store.inbox_error = ""
        except (OSError, sqlite3.Error):
            store.inbox_error = "Inbox processing needs attention. Check local folder permissions and disk space."
        stop.wait(3)


def serve(store, port=8787, open_browser=True):
    import webbrowser

    server = DashboardServer(store, port)
    stop = threading.Event()
    watcher = threading.Thread(target=watch_inbox, args=(store, stop), daemon=True, name="bo-inbox")
    watcher.start()
    url = f"http://127.0.0.1:{server.server_port}"
    print(f"WFMHub BO Live {'DEMO (synthetic)' if store.demo else 'LIVE'}: {url}", flush=True)
    print(f"Data folder: {store.home}\nKeep this window open. Press Ctrl+C to stop.", flush=True)
    if open_browser:
        webbrowser.open(url)
    try:
        server.serve_forever(poll_interval=0.3)
    except KeyboardInterrupt:
        pass
    finally:
        stop.set()
        server.server_close()
        watcher.join(timeout=5)
