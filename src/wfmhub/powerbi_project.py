"""Install, parameterize and open the shipped WFMHub Power BI Project."""

from __future__ import annotations

import os
import re
import shutil
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

from .config import Config


PROJECT_NAME = "WFMHub BI"
PROJECT_VERSION_FILE = "PROJECT_VERSION.txt"
_HUB_ROOT_PATTERN = re.compile(
    r'^expression HubRoot = ".*" meta '
    r'\[IsParameterQuery=true, Type="Text", IsParameterQueryRequired=true\]$',
    re.MULTILINE,
)


@dataclass(frozen=True)
class PowerBIProjectResult:
    project: Path
    installed: bool
    upgraded: bool
    archived: Path | None


def _template_root(config: Config) -> Path:
    candidates = (
        config.home / "templates" / "powerbi" / PROJECT_NAME,
        config.system / "templates" / "powerbi" / PROJECT_NAME,
    )
    for candidate in candidates:
        if (candidate / f"{PROJECT_NAME}.pbip").is_file():
            return candidate
    raise FileNotFoundError(
        "The WFMHub Power BI project template is missing. Re-extract the full portable release."
    )


def _version(folder: Path) -> int:
    try:
        return int((folder / PROJECT_VERSION_FILE).read_text(encoding="utf-8").strip())
    except (OSError, ValueError):
        return 0


def _parameterize(project_root: Path, hub_root: Path) -> None:
    expressions = (
        project_root / f"{PROJECT_NAME}.SemanticModel" / "definition" / "expressions.tmdl"
    )
    text = expressions.read_text(encoding="utf-8")
    # M escapes embedded quotes by doubling them; Windows backslashes remain
    # literal and therefore need no extra escaping.
    escaped = str(hub_root.resolve()).replace('"', '""')
    replacement = (
        f'expression HubRoot = "{escaped}" meta '
        '[IsParameterQuery=true, Type="Text", IsParameterQueryRequired=true]'
    )
    # Use a callable replacement so ``re`` never interprets Windows path
    # backslashes (for example ``C:\Users``) as replacement escapes such as
    # ``\U``.  The returned string is inserted literally.
    updated, count = _HUB_ROOT_PATTERN.subn(lambda _match: replacement, text, count=1)
    if count != 1:
        raise RuntimeError(
            f"Cannot set the HubRoot parameter in {expressions}. "
            "Use a fresh WFMHub release or restore the archived Power BI project."
        )
    if updated != text:
        expressions.write_text(updated, encoding="utf-8")


def install_powerbi_project(config: Config) -> PowerBIProjectResult:
    """Install the governed project and safely apply versioned upgrades."""

    source = _template_root(config)
    parent = config.reports / "Power BI"
    target = parent / PROJECT_NAME
    parent.mkdir(parents=True, exist_ok=True)
    installed = False
    upgraded = False
    archived: Path | None = None
    if not target.exists():
        shutil.copytree(source, target)
        installed = True
    elif (
        _version(source) > _version(target)
        or not (target / f"{PROJECT_NAME}.pbip").is_file()
        or not (
            target / f"{PROJECT_NAME}.SemanticModel" / "definition" /
            "expressions.tmdl"
        ).is_file()
    ):
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
        archived = config.reports / "Archive" / "Power BI" / f"{PROJECT_NAME}_{stamp}"
        archived.parent.mkdir(parents=True, exist_ok=True)
        shutil.move(str(target), str(archived))
        shutil.copytree(source, target)
        upgraded = True
    project = target / f"{PROJECT_NAME}.pbip"
    if not project.is_file():
        raise FileNotFoundError(f"Power BI project entry point is missing: {project}")
    _parameterize(target, config.home)
    return PowerBIProjectResult(project, installed, upgraded, archived)


def open_powerbi_project(config: Config, launch: bool = True) -> PowerBIProjectResult:
    """Install and open the project after checking the feed refresh receipt."""

    manifest = config.feed / "PowerBI" / "POWERBI_MANIFEST_CURRENT.csv"
    if not manifest.is_file():
        raise FileNotFoundError(
            "Power BI data is not ready. Run UPDATE > All sources once so "
            "Feed\\PowerBI and its refresh receipt are created."
        )
    result = install_powerbi_project(config)
    if launch:
        if os.name != "nt":
            raise RuntimeError(
                f"Power BI Desktop opening is available on Windows. Project installed at {result.project}"
            )
        os.startfile(str(result.project))  # type: ignore[attr-defined]
    return result
