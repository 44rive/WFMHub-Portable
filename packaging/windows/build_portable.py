#!/usr/bin/env python3
"""Build a no-admin Windows x64 WFMHub ZIP from Linux, macOS or Windows."""

from __future__ import annotations

import argparse
import hashlib
import shutil
import subprocess
import sys
import tomllib
import urllib.request
import zipfile
from pathlib import Path


ROOT = Path(__file__).resolve().parents[2]
DEFAULT_PYTHON = "3.13.7"
DEFAULT_VERSION = "0.13.0"
PYTHON_EMBED_SHA256 = {
    "3.13.7": "f6cca216a359be84797cabb54149ce5e062afb16cc7567eb7fc51cacb2d86b65",
}
NATIVE_SUFFIXES = (".dll", ".exe", ".pyd")


def download(url: str, target: Path) -> None:
    target.parent.mkdir(parents=True, exist_ok=True)
    if target.exists():
        return
    print(f"Downloading {url}")
    with urllib.request.urlopen(url) as response, target.open("wb") as handle:
        shutil.copyfileobj(response, handle)


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def verify_version(version: str) -> None:
    with (ROOT / "pyproject.toml").open("rb") as handle:
        project_version = tomllib.load(handle)["project"]["version"]
    init_text = (ROOT / "src" / "wfmhub" / "__init__.py").read_text(encoding="utf-8")
    init_version = init_text.split('__version__ = "', 1)[1].split('"', 1)[0]
    if len({version, project_version, init_version}) != 1:
        raise RuntimeError(
            f"Version mismatch: requested={version}, pyproject={project_version}, package={init_version}"
        )


def verify_embed_archive(path: Path, python_version: str) -> None:
    expected = PYTHON_EMBED_SHA256.get(python_version)
    if expected is None:
        raise RuntimeError(
            f"No reviewed CPython embeddable hash is registered for {python_version}. "
            "Add the official SHA-256 before building."
        )
    actual = sha256(path)
    if actual != expected:
        raise RuntimeError(f"CPython embeddable SHA-256 mismatch: expected {expected}, got {actual}")


def verify_wheel(wheel: Path) -> None:
    if not wheel.name.lower().endswith("-none-any.whl"):
        raise RuntimeError(f"Only pure-Python wheels are allowed: {wheel.name}")
    with zipfile.ZipFile(wheel) as archive:
        native = [name for name in archive.namelist() if name.lower().endswith((*NATIVE_SUFFIXES, ".so"))]
    if native:
        raise RuntimeError(f"Wheel contains native executable content: {wheel.name}: {native[:5]}")


def native_manifest(runtime: Path) -> dict[str, str]:
    return {
        path.relative_to(runtime).as_posix(): sha256(path)
        for path in sorted(runtime.rglob("*"))
        if path.is_file() and path.suffix.lower() in NATIVE_SUFFIXES
    }


def validate_stage(stage: Path, expected_native: dict[str, str]) -> None:
    banned = [
        path for path in stage.rglob("*")
        if any(token in path.as_posix().lower() for token in ("duckdb", "msvc_runtime"))
    ]
    if banned:
        raise RuntimeError(f"Banned runtime content remains: {[str(path) for path in banned[:10]]}")
    actual_native = native_manifest(stage / "_system" / "runtime")
    if actual_native != expected_native:
        added = sorted(set(actual_native) - set(expected_native))
        missing = sorted(set(expected_native) - set(actual_native))
        changed = sorted(key for key in set(actual_native) & set(expected_native) if actual_native[key] != expected_native[key])
        raise RuntimeError(f"Native runtime differs from reviewed CPython ZIP: added={added}, missing={missing}, changed={changed}")
    forbidden_user_files = [
        stage / "config" / "wfmhub.toml",
        stage / "config" / "wfm_rules.toml",
        stage / "config" / "metric_catalog.toml",
        stage / "config" / "analytics_rules.toml",
        stage / "config" / "report_catalog.toml",
        stage / "config" / "queue_mapping.csv",
        stage / "database" / "wfm.sqlite3",
        stage / "database" / "wfm.duckdb",
        stage / "_system" / "database" / "wfm.sqlite3",
        stage / "_system" / "database" / "wfm.duckdb",
    ]
    if any(path.exists() for path in forbidden_user_files):
        raise RuntimeError("Portable stage contains user configuration or database data")
    local_excel_masters = [
        path for pattern in ("*.xlsx", "*.xlsm")
        for path in (stage / "_system" / "templates" / "reports").glob(pattern)
    ]
    if local_excel_masters:
        raise RuntimeError(f"Portable stage contains local Excel report masters: {local_excel_masters}")
    unexpected_custom = [
        path for path in (stage / "_system" / "custom").rglob("*")
        if path.is_file() and path.name not in {
            "README.txt", "_paste_your_python_here.py", "_paste_your_sql_here.sql"
        }
    ]
    if unexpected_custom:
        raise RuntimeError(f"Portable stage contains runnable/user custom jobs: {unexpected_custom}")
    allowed_root = {
        "WFMHub.cmd", "SETUP.cmd", "README.md", "VERSION.txt",
        "Reports", "Feed", "config", "_system",
    }
    unexpected_root = sorted(path.name for path in stage.iterdir() if path.name not in allowed_root)
    if unexpected_root:
        raise RuntimeError(f"Portable root is cluttered: {unexpected_root}")


def copy_tree(source: Path, target: Path) -> None:
    if target.exists():
        shutil.rmtree(target)
    shutil.copytree(source, target, ignore=shutil.ignore_patterns("__pycache__", "*.pyc"))


def build(args) -> Path:
    verify_version(args.version)
    build_root = ROOT / "build" / "portable"
    cache = ROOT / "build" / "downloads"
    stage = build_root / "WFMHub"
    # Always recreate the entire stage and wheelhouse. This prevents a prior
    # DuckDB/MSVC build from leaking blocked native files into a new archive.
    if build_root.exists():
        shutil.rmtree(build_root)
    stage.mkdir(parents=True, exist_ok=True)
    runtime = stage / "_system" / "runtime"
    runtime.mkdir(parents=True, exist_ok=True)

    embed_name = f"python-{args.python_version}-embed-amd64.zip"
    embed_zip = cache / embed_name
    download(f"https://www.python.org/ftp/python/{args.python_version}/{embed_name}", embed_zip)
    verify_embed_archive(embed_zip, args.python_version)
    with zipfile.ZipFile(embed_zip) as archive:
        archive.extractall(runtime)
    expected_native = native_manifest(runtime)

    wheelhouse = build_root / "wheelhouse"
    wheelhouse.mkdir(parents=True, exist_ok=True)
    subprocess.run([
        sys.executable, "-m", "pip", "download", "-r", str(ROOT / "packaging" / "windows" / "runtime-requirements.lock"),
        "--dest", str(wheelhouse), "--platform", "win_amd64", "--python-version", "313",
        "--implementation", "cp", "--abi", "cp313", "--only-binary=:all:",
        "--require-hashes", "--no-deps",
    ], check=True)
    site_packages = runtime / "site-packages"
    site_packages.mkdir(exist_ok=True)
    for wheel in sorted(wheelhouse.glob("*.whl")):
        verify_wheel(wheel)
        print(f"Vendoring {wheel.name}")
        with zipfile.ZipFile(wheel) as archive:
            archive.extractall(site_packages)

    pth = runtime / "python313._pth"
    lines = [
        line for line in pth.read_text(encoding="utf-8").splitlines()
        if line.strip().lower() not in {"#import site", "import site"}
    ]
    if "../app" not in lines:
        lines.append("../app")
    if "site-packages" not in lines:
        lines.append("site-packages")
    pth.write_text("\n".join(lines) + "\n", encoding="utf-8")

    copy_tree(ROOT / "src" / "wfmhub", stage / "_system" / "app" / "wfmhub")
    copy_tree(ROOT / "sql", stage / "_system" / "app" / "sql")
    copy_tree(ROOT / "docs", stage / "_system" / "docs")
    copy_tree(ROOT / "prompts", stage / "_system" / "prompts")
    copy_tree(ROOT / "templates", stage / "_system" / "templates")
    # Excel-authored masters are local user assets and can contain refreshed
    # operational data. Ship the instructions/query pattern, never the files.
    report_template_dir = stage / "_system" / "templates" / "reports"
    for pattern in ("*.xlsx", "*.xlsm"):
        for master in report_template_dir.glob(pattern):
            master.unlink()
    (stage / "Reports").mkdir(parents=True, exist_ok=True)
    (stage / "Feed").mkdir(parents=True, exist_ok=True)
    shutil.copy2(ROOT / "Reports" / "README.txt", stage / "Reports" / "README.txt")
    shutil.copy2(ROOT / "Feed" / "README.txt", stage / "Feed" / "README.txt")
    # Ship only reviewed underscore templates. Runnable local jobs can contain
    # business logic or data and must never leak into a public portable ZIP.
    (stage / "_system" / "custom" / "jobs").mkdir(parents=True, exist_ok=True)
    (stage / "_system" / "custom" / "sql").mkdir(parents=True, exist_ok=True)
    shutil.copy2(ROOT / "custom" / "README.txt", stage / "_system" / "custom" / "README.txt")
    shutil.copy2(
        ROOT / "custom" / "jobs" / "_paste_your_python_here.py",
        stage / "_system" / "custom" / "jobs" / "_paste_your_python_here.py",
    )
    shutil.copy2(
        ROOT / "custom" / "sql" / "_paste_your_sql_here.sql",
        stage / "_system" / "custom" / "sql" / "_paste_your_sql_here.sql",
    )
    (stage / "config").mkdir(exist_ok=True)
    shutil.copy2(ROOT / "config" / "default.toml", stage / "config" / "default.toml")
    shutil.copy2(ROOT / "config" / "default_rules.toml", stage / "config" / "default_rules.toml")
    shutil.copy2(ROOT / "config" / "default_metrics.toml", stage / "config" / "default_metrics.toml")
    shutil.copy2(ROOT / "config" / "default_analytics.toml", stage / "config" / "default_analytics.toml")
    shutil.copy2(ROOT / "config" / "default_reports.toml", stage / "config" / "default_reports.toml")
    shutil.copy2(ROOT / "config" / "default_queue_mapping.csv", stage / "config" / "default_queue_mapping.csv")
    shutil.copy2(ROOT / "config" / "default_service_profiles.toml", stage / "config" / "default_service_profiles.toml")
    shutil.copy2(ROOT / "WFMHub.cmd", stage / "WFMHub.cmd")
    shutil.copy2(ROOT / "SETUP.cmd", stage / "SETUP.cmd")
    shutil.copy2(ROOT / "README.md", stage / "README.md")
    packaged_readme = stage / "README.md"
    packaged_readme.write_text(
        packaged_readme.read_text(encoding="utf-8").replace(
            "](docs/", "](_system/docs/"
        ),
        encoding="utf-8",
    )
    shutil.copy2(ROOT / "CHANGELOG.md", stage / "_system" / "CHANGELOG.md")
    (stage / "VERSION.txt").write_text(args.version + "\n", encoding="utf-8")
    manifest_lines = [f"{digest}  {name}" for name, digest in sorted(expected_native.items())]
    (stage / "_system" / "RUNTIME_MANIFEST.sha256").write_text("\n".join(manifest_lines) + "\n", encoding="utf-8")
    (stage / "_system" / "RUNTIME_ORIGIN.txt").write_text(
        f"Official CPython {args.python_version} Windows embeddable x64 ZIP\n"
        f"Archive SHA-256: {PYTHON_EMBED_SHA256[args.python_version]}\n"
        "Only pure-Python report libraries are added under runtime/site-packages.\n",
        encoding="utf-8",
    )
    for folder in (
        "Reports/Archive", "Feed", "_system/database", "_system/backups",
        "_system/logs", "_system/output", "_system/input",
    ):
        (stage / folder).mkdir(parents=True, exist_ok=True)
    shutil.copy2(ROOT / "input" / "README.md", stage / "_system" / "input" / "README.md")
    validate_stage(stage, expected_native)

    dist = ROOT / "dist"
    dist.mkdir(exist_ok=True)
    target = dist / f"WFMHub-Portable-v{args.version}-win-x64.zip"
    if target.exists():
        target.unlink()
    with zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED, compresslevel=9) as archive:
        for path in sorted(stage.rglob("*")):
            if path.is_file():
                archive.write(path, path.relative_to(build_root))
    digest = sha256(target)
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
