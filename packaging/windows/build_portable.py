#!/usr/bin/env python3
"""Build a no-admin Windows x64 WFMHub ZIP from Linux, macOS or Windows."""

from __future__ import annotations

import argparse
import hashlib
import shutil
import subprocess
import sys
import urllib.request
import zipfile
from pathlib import Path


ROOT = Path(__file__).resolve().parents[2]
DEFAULT_PYTHON = "3.13.7"
DEFAULT_VERSION = "0.1.1"


def download(url: str, target: Path) -> None:
    target.parent.mkdir(parents=True, exist_ok=True)
    if target.exists():
        return
    print(f"Downloading {url}")
    with urllib.request.urlopen(url) as response, target.open("wb") as handle:
        shutil.copyfileobj(response, handle)


def copy_tree(source: Path, target: Path) -> None:
    if target.exists():
        shutil.rmtree(target)
    shutil.copytree(source, target, ignore=shutil.ignore_patterns("__pycache__", "*.pyc"))


def build(args) -> Path:
    build_root = ROOT / "build" / "portable"
    cache = ROOT / "build" / "downloads"
    stage = build_root / "WFMHub"
    if args.clean and build_root.exists():
        shutil.rmtree(build_root)
    stage.mkdir(parents=True, exist_ok=True)
    runtime = stage / "runtime"
    runtime.mkdir(parents=True, exist_ok=True)

    embed_name = f"python-{args.python_version}-embed-amd64.zip"
    embed_zip = cache / embed_name
    download(f"https://www.python.org/ftp/python/{args.python_version}/{embed_name}", embed_zip)
    with zipfile.ZipFile(embed_zip) as archive:
        archive.extractall(runtime)

    wheelhouse = build_root / "wheelhouse"
    wheelhouse.mkdir(parents=True, exist_ok=True)
    subprocess.run([
        sys.executable, "-m", "pip", "download", "-r", str(ROOT / "packaging" / "windows" / "runtime-requirements.lock"),
        "--dest", str(wheelhouse), "--platform", "win_amd64", "--python-version", "313",
        "--implementation", "cp", "--abi", "cp313", "--only-binary=:all:",
    ], check=True)
    site_packages = runtime / "site-packages"
    site_packages.mkdir(exist_ok=True)
    for wheel in sorted(wheelhouse.glob("*.whl")):
        print(f"Vendoring {wheel.name}")
        with zipfile.ZipFile(wheel) as archive:
            archive.extractall(site_packages)
            # pip normally moves wheel .data/data payloads to sys.prefix. We
            # extract wheels ourselves, so perform that wheel-install step for
            # Microsoft app-local runtime DLLs explicitly.
            for name in archive.namelist():
                if ".data/data/" in name and name.lower().endswith(".dll"):
                    target = runtime / Path(name).name
                    with archive.open(name) as source, target.open("wb") as handle:
                        shutil.copyfileobj(source, handle)

    pth = runtime / "python313._pth"
    lines = [line for line in pth.read_text(encoding="utf-8").splitlines() if line.strip() != "#import site"]
    if "../app" not in lines:
        lines.append("../app")
    if "site-packages" not in lines:
        lines.append("site-packages")
    lines.append("import site")
    pth.write_text("\n".join(lines) + "\n", encoding="utf-8")

    copy_tree(ROOT / "src" / "wfmhub", stage / "app" / "wfmhub")
    copy_tree(ROOT / "sql", stage / "app" / "sql")
    copy_tree(ROOT / "docs", stage / "docs")
    (stage / "config").mkdir(exist_ok=True)
    shutil.copy2(ROOT / "config" / "default.toml", stage / "config" / "default.toml")
    shutil.copy2(ROOT / "WFMHub.cmd", stage / "WFMHub.cmd")
    shutil.copy2(ROOT / "SETUP.cmd", stage / "SETUP.cmd")
    shutil.copy2(ROOT / "README.md", stage / "README.md")
    shutil.copy2(ROOT / "CHANGELOG.md", stage / "CHANGELOG.md")
    (stage / "VERSION.txt").write_text(args.version + "\n", encoding="utf-8")
    for folder in ("database", "backups", "logs", "output", "input", "extracts"):
        (stage / folder).mkdir(exist_ok=True)
    shutil.copy2(ROOT / "input" / "README.md", stage / "input" / "README.md")

    dist = ROOT / "dist"
    dist.mkdir(exist_ok=True)
    target = dist / f"WFMHub-Portable-v{args.version}-win-x64.zip"
    if target.exists():
        target.unlink()
    with zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED, compresslevel=9) as archive:
        for path in sorted(stage.rglob("*")):
            if path.is_file():
                archive.write(path, path.relative_to(build_root))
    digest = hashlib.sha256(target.read_bytes()).hexdigest()
    target.with_suffix(target.suffix + ".sha256").write_text(f"{digest}  {target.name}\n", encoding="utf-8")
    print(f"Built {target}")
    print(f"SHA-256 {digest}")
    return target


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--version", default=DEFAULT_VERSION)
    parser.add_argument("--python-version", default=DEFAULT_PYTHON)
    parser.add_argument("--clean", action="store_true")
    args = parser.parse_args()
    build(args)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
