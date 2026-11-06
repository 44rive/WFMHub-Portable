"""Build the fully bundled ZIP from the hash-pinned WFMHub runtime archive."""
from __future__ import annotations

import argparse
import hashlib
import shutil
import zipfile
from pathlib import Path

BASE = Path(__file__).resolve().parent
SOURCE_NAME = "WFMHub-Portable-v1.1.2-win-x64.zip"
SOURCE_SHA = "191bfa106d80aa3a72c7584f5333d70e9699111124fbe7bdc679dada6b7f880d"
WHEELS = {
    "openpyxl-3.1.5-py2.py3-none-any.whl": "5282c12b107bffeef825f4617dc029afaf41d0ea60823bbb665ef3079dc79de2",
    "et_xmlfile-2.0.0-py3-none-any.whl": "7a91720bc756843502c3b7504c77b8fe44217c85c537d85037f0f536151b2caa",
}


def build(runtime_archive: Path, output: Path) -> Path:
    if hashlib.sha256(runtime_archive.read_bytes()).hexdigest() != SOURCE_SHA:
        raise ValueError("WFMHub source archive SHA-256 mismatch")
    root = output / "Mr.Propre"
    if root.exists():
        raise ValueError(f"Build folder already exists: {root}; use a fresh output directory")
    runtime = root / "_system" / "runtime"
    runtime.mkdir(parents=True)
    with zipfile.ZipFile(runtime_archive) as source:
        prefix = "WFMHub/_system/runtime/"
        for name in source.namelist():
            if not name.startswith(prefix) or name.endswith("/"):
                continue
            relative = name[len(prefix):]
            path = Path(relative)
            if path.is_absolute() or ".." in path.parts:
                raise ValueError("Invalid archive member")
            if len(path.parts) > 1 and relative not in {"wheels/" + n for n in WHEELS}:
                continue
            data = source.read(name)
            if path.name in WHEELS and hashlib.sha256(data).hexdigest() != WHEELS[path.name]:
                raise ValueError(f"Wheel checksum mismatch: {path.name}")
            destination = runtime / path
            destination.parent.mkdir(parents=True, exist_ok=True)
            destination.write_bytes(data)
    for name in ("python.exe", "python313.dll", "python313.zip", "LICENSE.txt", *("wheels/" + n for n in WHEELS)):
        if not (runtime / name).is_file():
            raise ValueError(f"Missing runtime file: {name}")
    (runtime / "python313._pth").write_text(
        "python313.zip\n.\n../app\n" + "".join(f"wheels/{name}\n" for name in sorted(WHEELS)),
        encoding="ascii",
    )
    app = root / "_system" / "app"
    app.mkdir()
    shutil.copyfile(BASE / "mr_propre.py", app / "mr_propre.py")
    shutil.copyfile(BASE / "README.md", root / "README.md")
    # CMD files use Windows line endings even when built on Linux.
    (root / "Mr.Propre.cmd").write_bytes((BASE / "Mr.Propre.cmd").read_text().replace("\r\n", "\n").replace("\n", "\r\n").encode("utf-8"))
    (root / "VERSION.txt").write_text("Mr.Propre 1.0.1\nMade by Anass Assri\n", encoding="utf-8")
    origin = root / "_system" / "RUNTIME_ORIGIN.txt"
    origin.write_text(
        f"Source: https://github.com/44rive/WFMHub-Portable/releases/download/v1.1.2/{SOURCE_NAME}\n"
        f"SHA256: {SOURCE_SHA}\nCPython 3.13.7 x64; native runtime files unchanged.\n"
        "python313._pth restricted to the standalone app and two pure-Python wheels.\n"
        + "".join(f"{digest}  {name}\n" for name, digest in WHEELS.items()), encoding="utf-8",
    )
    archive = output / "Mr.Propre-Portable-v1.0.1-win-x64.zip"
    with zipfile.ZipFile(archive, "w", zipfile.ZIP_DEFLATED, compresslevel=9) as package:
        for path in sorted(root.rglob("*")):
            if path.is_file():
                info = zipfile.ZipInfo(path.relative_to(output).as_posix(), (2026, 1, 1, 0, 0, 0))
                info.compress_type = zipfile.ZIP_DEFLATED
                info.external_attr = 0o100644 << 16
                package.writestr(info, path.read_bytes(), compresslevel=9)
    digest = hashlib.sha256(archive.read_bytes()).hexdigest()
    archive.with_suffix(".zip.sha256").write_text(f"{digest}  {archive.name}\n", encoding="ascii")
    print(f"Created {archive} ({archive.stat().st_size:,} bytes)\nSHA256: {digest}")
    return archive


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--runtime-archive", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    build(args.runtime_archive, args.output)
