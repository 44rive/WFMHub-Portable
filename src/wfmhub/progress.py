"""Small dependency-free progress display for Windows CMD and terminals."""

from __future__ import annotations

import os
import shutil
import sys
from typing import Protocol, TextIO


class ProgressCallback(Protocol):
    """Report completed work; a total of zero means indeterminate work."""

    def __call__(self, current: int, total: int, label: str) -> None: ...


class ProgressBar:
    """Render one in-place ASCII progress line without ANSI escape codes."""

    def __init__(
        self,
        title: str = "WFMHub",
        *,
        stream: TextIO | None = None,
        width: int = 28,
        enabled: bool | None = None,
    ) -> None:
        self.stream = stream or sys.stdout
        setting = os.environ.get("WFMHUB_PROGRESS", "").strip().lower()
        if enabled is None:
            if setting in {"0", "false", "no", "off"}:
                enabled = False
            elif setting in {"1", "true", "yes", "on"}:
                enabled = True
            else:
                try:
                    enabled = bool(self.stream.isatty())
                except (AttributeError, OSError):
                    enabled = False
        self.enabled = enabled
        self.title = title.strip() or "WFMHub"
        self.width = max(12, width)
        self._last_length = 0
        self._pulse_position = 0
        self._closed = False

    def _render(self, body: str) -> None:
        if not self.enabled or self._closed:
            return
        body = str(body).replace("\r", " ").replace("\n", " ")
        terminal_width = shutil.get_terminal_size(fallback=(100, 24)).columns
        prefix = f"{self.title} "
        available = max(12, terminal_width - len(prefix) - 1)
        body = body[:available]
        line = prefix + body
        padding = " " * max(0, self._last_length - len(line))
        self.stream.write("\r" + line + padding)
        self.stream.flush()
        self._last_length = len(line)

    def update(self, fraction: float, label: str) -> None:
        """Show a known 0..1 completion fraction."""
        fraction = min(1.0, max(0.0, float(fraction)))
        filled = int(self.width * fraction)
        bar = "#" * filled + "-" * (self.width - filled)
        self._render(f"[{bar}] {round(fraction * 100):3d}% {label}")

    def pulse(self, label: str) -> None:
        """Show movement when the total amount of work is not known."""
        marker_width = min(5, max(3, self.width // 5))
        travel = max(1, self.width - marker_width + 1)
        position = self._pulse_position % travel
        self._pulse_position += 1
        bar = "-" * position + "#" * marker_width
        bar += "-" * (self.width - len(bar))
        self._render(f"[{bar}] working {label}")

    def finish(self, label: str = "Complete") -> None:
        if not self.enabled or self._closed:
            return
        self.update(1.0, label)
        self.stream.write("\n")
        self.stream.flush()
        self._closed = True

    def fail(self, label: str = "Failed") -> None:
        if not self.enabled or self._closed:
            return
        self._render(f"[{'!' * self.width}] FAILED {label}")
        self.stream.write("\n")
        self.stream.flush()
        self._closed = True
