"""Build the read-only ABS probe with the established WFMHub runtime layout."""
from pathlib import Path
import importlib.util
import shutil
import zipfile

ROOT = Path(__file__).resolve().parents[2]


def build():
    spec = importlib.util.spec_from_file_location("wfmhub_portable_builder", ROOT / "packaging/windows/build_portable.py")
    helper = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(helper)
    source = ROOT / "miniapps/bo_live"
    build_root = ROOT / "build/abs-probe"
    if build_root.exists():
        shutil.rmtree(build_root)
    stage = build_root / "WFMHub-ABS-Probe"
    app = stage / "_system/app"
    app.mkdir(parents=True)
    shutil.copy2(source / "abs_probe.py", app / "abs_probe.py")
    shutil.copy2(source / "ABS_PROBE_README.md", stage / "README.md")
    (stage / "PROBE_ABS.cmd").write_bytes((source / "PROBE_ABS.cmd").read_text().replace("\r\n", "\n").replace("\n", "\r\n").encode())
    version = helper.DEFAULT_PYTHON
    archive = ROOT / "build/downloads" / f"python-{version}-embed-amd64.zip"
    helper.download(f"https://www.python.org/ftp/python/{version}/{archive.name}", archive)
    helper.verify_embed_archive(archive, version)
    with zipfile.ZipFile(archive) as package:
        package.extractall(stage / "_system/runtime")
    (stage / "VERSION.txt").write_text("abs-probe-0.1.0\n")
    (stage / "RUNTIME_ORIGIN.txt").write_text(f"Official CPython {version} Windows x64 embeddable ZIP\nSHA-256: {helper.PYTHON_EMBED_SHA256[version]}\nNo third-party packages added.\n")
    manifest = "\n".join(f"{helper.sha256(p)}  {p.relative_to(stage).as_posix()}" for p in sorted(stage.rglob("*")) if p.is_file())
    (stage / "SHA256SUMS.txt").write_text(manifest + "\n")
    target = ROOT / "dist/WFMHub-ABS-Probe-v0.1.0-win-x64.zip"
    target.parent.mkdir(exist_ok=True)
    with zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED, compresslevel=9) as package:
        for path in sorted(stage.rglob("*")):
            if path.is_file():
                package.write(path, path.relative_to(build_root))
    target.with_suffix(".zip.sha256").write_text(f"{helper.sha256(target)}  {target.name}\n")
    print(target)
    return target


if __name__ == "__main__":
    build()
