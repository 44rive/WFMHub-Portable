#!/usr/bin/env python3
"""Build BO Live independently; explicit source allowlist excludes all local state."""

from __future__ import annotations

import argparse
import importlib.util
import shutil
import subprocess
import sys
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
SOURCE = ROOT / "miniapps" / "bo_live"
CODE = ("__init__.py", "app.py", "core.py", "server.py", "schema.sql",
        "web/index.html", "web/app.css", "web/app.js")
ROOT_FILES = ("RUN_PYTHON.cmd", "START_BO_LIVE.cmd", "DEMO_BO_LIVE.cmd", "PUBLISH_CAPTURE.cmd",
              "README.md", "PAD_SETUP.md", "PAD_COLUMNS.csv")


def stage_source(stage):
    for name in CODE:
        target = stage / "bo_live" / name
        target.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(SOURCE / name, target)
    for name in ROOT_FILES:
        shutil.copy2(SOURCE / name, stage / name)
        if name.endswith(".cmd"):
            (stage / name).write_bytes((SOURCE / name).read_text(encoding="utf-8").replace("\r\n", "\n").replace("\n", "\r\n").encode("utf-8"))
    (stage / "VERSION.txt").write_text("0.1.0\n", encoding="utf-8")


def build(source_only=False):
    build_root = ROOT / "build" / "bo-live"
    stage = build_root / "WFMHub-BO-Live"
    if build_root.exists():
        shutil.rmtree(build_root)
    stage.mkdir(parents=True)
    stage_source(stage)
    spec = importlib.util.spec_from_file_location("wfmhub_portable_builder", ROOT / "packaging/windows/build_portable.py")
    helper = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(helper)
    if not source_only:
        cache = ROOT / "build" / "downloads"
        version = helper.DEFAULT_PYTHON
        name = f"python-{version}-embed-amd64.zip"
        embed = cache / name
        helper.download(f"https://www.python.org/ftp/python/{version}/{name}", embed)
        helper.verify_embed_archive(embed, version)
        runtime = stage / "runtime"
        with zipfile.ZipFile(embed) as archive:
            archive.extractall(runtime)
        expected_native = helper.native_manifest(runtime)
        requirement = next(line for line in (ROOT / "packaging/windows/runtime-requirements.lock").read_text().splitlines() if line.startswith("tzdata=="))
        requirements = build_root / "timezone-requirement.txt"
        requirements.write_text(requirement + "\n", encoding="utf-8")
        wheels = runtime / "wheels"
        subprocess.run([sys.executable, "-m", "pip", "download", "-r", str(requirements), "--dest", str(wheels),
                        "--only-binary=:all:", "--require-hashes", "--no-deps"], check=True)
        wheel = next(wheels.glob("tzdata-*.whl"))
        helper.verify_wheel(wheel)
        (runtime / "python313._pth").write_text(f"python313.zip\n.\nwheels/{wheel.name}\n", encoding="utf-8")
        if helper.native_manifest(runtime) != expected_native:
            raise RuntimeError("Native runtime changed after extraction of the reviewed CPython archive.")
        (stage / "RUNTIME_ORIGIN.txt").write_text(
            f"Official CPython {version} Windows embeddable x64 ZIP\n"
            f"Archive SHA-256: {helper.PYTHON_EMBED_SHA256[version]}\n"
            f"Only the hash-pinned pure-Python {wheel.name} is added.\n", encoding="utf-8")
    forbidden = [p for p in stage.rglob("*") if p.is_file() and
                 (p.suffix in (".sqlite3", ".db", ".pyc") or p.name == "settings.json" or "_local" in p.parts)]
    if forbidden:
        raise RuntimeError("BO Live stage contains local data or configuration.")
    manifest = "\n".join(f"{helper.sha256(p)}  {p.relative_to(stage).as_posix()}" for p in sorted(stage.rglob("*")) if p.is_file())
    (stage / "SHA256SUMS.txt").write_text(manifest + "\n", encoding="utf-8")
    dist = ROOT / "dist"
    dist.mkdir(exist_ok=True)
    target = dist / f"WFMHub-BO-Live-v0.1.0-{'source' if source_only else 'win-x64'}.zip"
    with zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED, compresslevel=9) as archive:
        for path in sorted(stage.rglob("*")):
            if path.is_file():
                archive.write(path, path.relative_to(build_root))
    target.with_suffix(".zip.sha256").write_text(f"{helper.sha256(target)}  {target.name}\n", encoding="utf-8")
    print(f"Built {target}", flush=True)
    return target


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--source-only", action="store_true")
    raise SystemExit(0 if build(parser.parse_args().source_only) else 1)
