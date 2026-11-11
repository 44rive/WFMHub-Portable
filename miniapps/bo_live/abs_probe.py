"""Read-only Windows UI Automation diagnostic. No third-party dependencies.

COM slots/property IDs follow Microsoft's UIAutomationClient.h. Only getters
are exposed. A separate, time-bounded worker contains unresponsive providers.
"""
from __future__ import annotations

import argparse
import ctypes as C
import json
import os
from pathlib import Path
import subprocess
import sys
import time
import uuid
from datetime import datetime, timezone


class ProbeError(RuntimeError):
    pass


class GUID(C.Structure):
    _fields_ = [("a", C.c_uint32), ("b", C.c_uint16), ("c", C.c_uint16), ("d", C.c_ubyte * 8)]

    @classmethod
    def parse(cls, value):
        return cls.from_buffer_copy(uuid.UUID(value).bytes_le)


class VariantData(C.Union):
    _fields_ = [("integer", C.c_int32), ("boolean", C.c_int16), ("pointer", C.c_void_p),
                ("record", C.c_void_p * 2)]


class VARIANT(C.Structure):
    _fields_ = [("vt", C.c_uint16), ("reserved", C.c_uint16 * 3), ("data", VariantData)]


HEADERS = {v.casefold(): v for v in ("Task Group", "Task type", "Task priority", "Language",
          "Scheduled for", "Assigned to", "Case number", "Task description", "Contract Name")}
PROPERTIES = {"control_type": 30003, "name": 30005, "automation_id": 30011, "class_name": 30012,
              "offscreen": 30022, "framework": 30024, "grid_item": 30029, "grid": 30030,
              "table": 30038, "text_pattern": 30040, "value_pattern": 30043}


def describe(value):
    """Persist known header labels and string shape, never arbitrary UI text."""
    if not isinstance(value, str):
        return {"present": False, "length": 0}
    label = HEADERS.get(value.strip().casefold())
    return {"present": bool(value), "length": len(value), "header": label,
            "redacted": bool(value) and label is None,
            "ends_in_ellipsis": value.rstrip().endswith(("...", "…"))}


class WindowsUIAutomation:
    def __init__(self):
        if os.name != "nt" or C.sizeof(C.c_void_p) != 8:
            raise ProbeError("This probe requires 64-bit Windows. Use the portable Windows ZIP.")
        self.ole = C.WinDLL("ole32", use_last_error=True)
        self.aut = C.WinDLL("oleaut32", use_last_error=True)
        self.ole.CoInitializeEx.argtypes = [C.c_void_p, C.c_uint32]
        self.ole.CoInitializeEx.restype = C.c_int32
        self.ole.CoCreateInstance.argtypes = [C.POINTER(GUID), C.c_void_p, C.c_uint32, C.POINTER(GUID), C.POINTER(C.c_void_p)]
        self.ole.CoCreateInstance.restype = C.c_int32
        self.ole.CoUninitialize.argtypes = []
        self.ole.CoUninitialize.restype = None
        self.aut.VariantClear.argtypes = [C.POINTER(VARIANT)]
        self.aut.VariantClear.restype = C.c_int32
        self.aut.SysStringLen.argtypes = [C.c_void_p]
        self.aut.SysStringLen.restype = C.c_uint32
        self.client = C.c_void_p()
        self.walker = C.c_void_p()
        self.check(self.ole.CoInitializeEx(None, 0))
        try:
            clsid = GUID.parse("FF48DBA4-60EF-4201-AA87-54103EEF594E")
            iid = GUID.parse("30CBE57D-D9D0-452A-AB13-7AC5AC4825EE")
            self.check(self.ole.CoCreateInstance(C.byref(clsid), None, 1, C.byref(iid), C.byref(self.client)))
            self.call(self.client, 16, [C.POINTER(C.c_void_p)], C.byref(self.walker))
        except BaseException:
            self.close()
            raise

    @staticmethod
    def check(result):
        if result < 0:
            raise ProbeError(f"Windows accessibility returned HRESULT 0x{result & 0xffffffff:08X}.")

    def call(self, pointer, slot, types, *args):
        address = C.cast(pointer, C.POINTER(C.POINTER(C.c_void_p))).contents[slot]
        function = C.WINFUNCTYPE(C.c_int32, C.c_void_p, *types)(address)
        self.check(function(pointer, *args))

    def release(self, pointer):
        if pointer:
            address = C.cast(pointer, C.POINTER(C.POINTER(C.c_void_p))).contents[2]
            C.WINFUNCTYPE(C.c_uint32, C.c_void_p)(address)(pointer)

    def close(self):
        self.release(self.walker)
        self.release(self.client)
        self.walker, self.client = C.c_void_p(), C.c_void_p()
        self.ole.CoUninitialize()

    def element(self, hwnd):
        result = C.c_void_p()
        self.call(self.client, 6, [C.c_void_p, C.POINTER(C.c_void_p)], hwnd, C.byref(result))
        if not result:
            raise ProbeError("The selected window did not expose an accessibility root.")
        return result

    def property(self, element, identity):
        value = VARIANT()
        try:
            self.call(element, 10, [C.c_int32, C.POINTER(VARIANT)], identity, C.byref(value))
            if value.vt == 3:
                return value.data.integer
            if value.vt == 11:
                return value.data.boolean != 0
            if value.vt == 8 and value.data.pointer:
                length = self.aut.SysStringLen(value.data.pointer)
                if length > 100000:
                    raise ProbeError("Accessibility text exceeded the diagnostic size limit.")
                return C.wstring_at(value.data.pointer, length)
            return None  # Includes the provider's NotSupported sentinel.
        finally:
            self.aut.VariantClear(C.byref(value))

    def adjacent(self, element, child=False):
        result = C.c_void_p()
        self.call(self.walker, 4 if child else 6, [C.c_void_p, C.POINTER(C.c_void_p)], element, C.byref(result))
        return result

    def read(self, element):
        if self.property(element, 30019) is True:
            return {"password_control": True}
        record = {}
        for key, identity in PROPERTIES.items():
            value = self.property(element, identity)
            record[key] = describe(value) if key in ("name", "automation_id", "class_name", "framework") else value
        if record["value_pattern"]:
            record["value"] = describe(self.property(element, 30045))
        if record["grid"]:
            record["grid_rows"] = self.property(element, 30062)
            record["grid_columns"] = self.property(element, 30063)
        if record["grid_item"]:
            record["row"] = self.property(element, 30064)
            record["column"] = self.property(element, 30065)
        return record


def windows(title):
    """Enumerate visible Edge windows only. Titles are used locally, not persisted."""
    if os.name != "nt":
        raise ProbeError("This probe runs on Windows. No report of ABS accessibility can be made on this system.")
    user = C.WinDLL("user32", use_last_error=True)
    kernel = C.WinDLL("kernel32", use_last_error=True)
    callback_type = C.WINFUNCTYPE(C.c_int32, C.c_void_p, C.c_ssize_t)
    user.EnumWindows.argtypes = [callback_type, C.c_ssize_t]
    user.EnumWindows.restype = C.c_int32
    user.IsWindowVisible.argtypes = [C.c_void_p]
    user.IsWindowVisible.restype = C.c_int32
    user.GetWindowTextLengthW.argtypes = [C.c_void_p]
    user.GetWindowTextLengthW.restype = C.c_int32
    user.GetWindowTextW.argtypes = [C.c_void_p, C.c_wchar_p, C.c_int32]
    user.GetWindowTextW.restype = C.c_int32
    user.GetWindowThreadProcessId.argtypes = [C.c_void_p, C.POINTER(C.c_uint32)]
    user.GetWindowThreadProcessId.restype = C.c_uint32
    kernel.OpenProcess.argtypes = [C.c_uint32, C.c_int32, C.c_uint32]
    kernel.OpenProcess.restype = C.c_void_p
    kernel.QueryFullProcessImageNameW.argtypes = [C.c_void_p, C.c_uint32, C.c_wchar_p, C.POINTER(C.c_uint32)]
    kernel.QueryFullProcessImageNameW.restype = C.c_int32
    kernel.CloseHandle.argtypes = [C.c_void_p]
    kernel.CloseHandle.restype = C.c_int32
    found = []

    @callback_type
    def visit(hwnd, unused):
        if not user.IsWindowVisible(hwnd):
            return 1
        length = user.GetWindowTextLengthW(hwnd)
        if not 0 < length < 32768:
            return 1
        buffer = C.create_unicode_buffer(length + 1)
        user.GetWindowTextW(hwnd, buffer, length + 1)
        if title.casefold() not in buffer.value.casefold():
            return 1
        pid = C.c_uint32()
        user.GetWindowThreadProcessId(hwnd, C.byref(pid))
        handle = kernel.OpenProcess(0x1000, 0, pid.value)
        if handle:
            try:
                path = C.create_unicode_buffer(32768)
                size = C.c_uint32(len(path))
                if kernel.QueryFullProcessImageNameW(handle, 0, path, C.byref(size)) and Path(path.value).name.casefold() == "msedge.exe":
                    found.append({"handle": hex(hwnd), "process_id": pid.value, "browser": "msedge.exe", "title_length": length})
            finally:
                kernel.CloseHandle(handle)
        return 1

    if not user.EnumWindows(visit, 0):
        raise ProbeError("Windows did not permit window enumeration.")
    return found


def inspect_tree(adapter, root, max_nodes=1200, max_depth=20):
    nodes, pending, limited = [], [(root, None, 0)], False
    try:
        while pending and len(nodes) < max_nodes:
            element, parent, depth = pending.pop()
            index = len(nodes)
            record = {"node": index, "parent": parent, "depth": depth}
            try:
                try:
                    record.update(adapter.read(element))
                except ProbeError as exc:
                    record["read_error"] = str(exc)
                nodes.append(record)
                if depth:
                    sibling = adapter.adjacent(element)
                    if sibling:
                        pending.append((sibling, parent, depth))
                if depth < max_depth and not record.get("password_control"):
                    child = adapter.adjacent(element, child=True)
                    if child:
                        pending.append((child, index, depth + 1))
                elif depth >= max_depth:
                    limited = True
            finally:
                adapter.release(element)
        limited = limited or bool(pending)
    finally:
        for element, _, _ in pending:
            adapter.release(element)
    headers = sorted({n[key]["header"] for n in nodes for key in ("name", "value")
                      if isinstance(n.get(key), dict) and n[key].get("header")})
    return {"nodes": nodes, "summary": {"nodes_scanned": len(nodes), "limit_reached": limited,
            "headers_found": headers, "grid_candidates": [n["node"] for n in nodes if n.get("grid")],
            "positioned_cells": sum(n.get("grid_item") is True for n in nodes),
            "read_errors": sum("read_error" in n for n in nodes),
            "full_backlog_coverage": "not_verified"}}


def worker(title, handle, max_nodes, max_depth):
    candidates = windows(title)
    if handle:
        candidates = [w for w in candidates if int(w["handle"], 16) == int(handle, 0)]
    if len(candidates) != 1:
        return {"status": "no_matching_window" if not candidates else "ambiguous_windows", "candidates": candidates,
                "message": "Open ABS as the active tab in a visible Edge window. Close duplicate ABS windows or select one with --window-handle."}
    automation = WindowsUIAutomation()
    try:
        result = inspect_tree(automation, automation.element(int(candidates[0]["handle"], 16)), max_nodes, max_depth)
        return dict(result, status="probe_completed", target=candidates[0])
    finally:
        automation.close()


def main(argv=None):
    parser = argparse.ArgumentParser(description="Read-only ABS accessibility probe for Windows Edge")
    parser.add_argument("--title", default="ABS@ACM")
    parser.add_argument("--window-handle")
    parser.add_argument("--max-nodes", type=int, default=1200)
    parser.add_argument("--max-depth", type=int, default=20)
    parser.add_argument("--timeout", type=int, default=45)
    parser.add_argument("--output", type=Path, default=Path.cwd() / "_system" / "output" / "abs-probe")
    parser.add_argument("--worker", action="store_true", help=argparse.SUPPRESS)
    args = parser.parse_args(argv)
    if not 1 <= args.max_nodes <= 5000 or not 1 <= args.max_depth <= 40 or not 5 <= args.timeout <= 120 or not args.title.strip():
        parser.error("Use 1–5000 nodes, 1–40 depth, 5–120 seconds, and a nonempty title marker.")
    if args.worker:
        try:
            report = worker(args.title, args.window_handle, args.max_nodes, args.max_depth)
        except (ProbeError, OSError, ValueError) as exc:
            report = {"status": "probe_failed", "message": str(exc)}
        print(json.dumps(report, ensure_ascii=True))
        return 0
    command = [sys.executable, "-I", str(Path(__file__).resolve()), "--worker", "--title", args.title,
               "--max-nodes", str(args.max_nodes), "--max-depth", str(args.max_depth)]
    if args.window_handle:
        command.extend(["--window-handle", args.window_handle])
    start = time.monotonic()
    try:
        result = subprocess.run(command, capture_output=True, text=True, encoding="utf-8", timeout=args.timeout)
        report = json.loads(result.stdout) if result.returncode == 0 else {"status": "probe_failed", "message": "Probe worker exited unexpectedly. ABS was not modified."}
    except subprocess.TimeoutExpired:
        report = {"status": "timeout", "message": "The accessibility provider did not respond within the limit. Only the probe worker was stopped."}
    except (ValueError, OSError):
        report = {"status": "probe_failed", "message": "The local probe worker could not return a diagnostic report."}
    report.update(schema="wfmhub_abs_probe_v1", captured_at=datetime.now(timezone.utc).isoformat(),
                  elapsed_seconds=round(time.monotonic() - start, 2), values_redacted=True,
                  scope="Visible matching Edge window; no clicks, scrolling, login, clipboard access, or task changes.")
    args.output.mkdir(parents=True, exist_ok=True)
    prefix = datetime.now(timezone.utc).strftime("probe-%Y%m%d-%H%M%S-") + uuid.uuid4().hex[:8]
    path = args.output / (prefix + ".json")
    path.write_text(json.dumps(report, ensure_ascii=True, indent=2) + "\n", encoding="utf-8")
    summary = report.get("summary", {})
    lines = ["WFMHub ABS read-only probe", "Status: " + report["status"], report.get("message", ""),
             "Nodes inspected: " + str(summary.get("nodes_scanned", 0)),
             "Headers found: " + ", ".join(summary.get("headers_found", [])),
             "Grid candidates: " + str(len(summary.get("grid_candidates", []))),
             "Cells with row/column pattern: " + str(summary.get("positioned_cells", 0)),
             "Traversal limit reached: " + str(summary.get("limit_reached", False)),
             "Full backlog coverage: NOT VERIFIED. A probe is not a task capture.",
             "Case text, URLs and arbitrary control values are redacted.", "Report: " + str(path.resolve())]
    text = "\n".join(lines) + "\n"
    path.with_suffix(".txt").write_text(text, encoding="utf-8")
    print(text)
    return 0 if report["status"] == "probe_completed" else 1


if __name__ == "__main__":
    raise SystemExit(main())
