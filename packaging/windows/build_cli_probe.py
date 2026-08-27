#!/usr/bin/env python3
"""Build the signed DuckDB CLI App Control policy probe."""

from __future__ import annotations

import hashlib
import shutil
import struct
import urllib.request
import zipfile
from pathlib import Path


ROOT = Path(__file__).resolve().parents[2]
DUCKDB_VERSION = "1.5.5"
DUCKDB_ASSET = "duckdb_cli-windows-amd64.zip"
DUCKDB_URL = (
    f"https://github.com/duckdb/duckdb/releases/download/v{DUCKDB_VERSION}/"
    f"{DUCKDB_ASSET}"
)
DUCKDB_SHA256 = "e1428b7114a841626b5054723731cbf45c6df91b42ae1a6c355f88fad1f6dc4c"
DUCKDB_EXE_SHA256 = "fde737c7749075f6b54e14772a4e6b33a5fa0201075d03640aca358074ea4554"


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def download_verified(target: Path) -> None:
    target.parent.mkdir(parents=True, exist_ok=True)
    if not target.exists():
        print(f"Downloading {DUCKDB_URL}")
        with urllib.request.urlopen(DUCKDB_URL) as response, target.open("wb") as handle:
            shutil.copyfileobj(response, handle)
    actual = sha256(target)
    if actual != DUCKDB_SHA256:
        raise RuntimeError(
            f"DuckDB asset checksum mismatch: expected {DUCKDB_SHA256}, received {actual}"
        )


def authenticode_table(binary: Path) -> tuple[int, int]:
    """Return the PE certificate-table file offset and size."""
    data = binary.read_bytes()
    if data[:2] != b"MZ":
        raise RuntimeError(f"Not a Windows PE binary: {binary}")
    pe_offset = struct.unpack_from("<I", data, 0x3C)[0]
    if data[pe_offset : pe_offset + 4] != b"PE\0\0":
        raise RuntimeError(f"Invalid PE signature: {binary}")
    machine = struct.unpack_from("<H", data, pe_offset + 4)[0]
    if machine != 0x8664:
        raise RuntimeError(f"Expected x64 PE machine 0x8664, received 0x{machine:04x}")
    optional = pe_offset + 24
    magic = struct.unpack_from("<H", data, optional)[0]
    data_directories = optional + (112 if magic == 0x20B else 96 if magic == 0x10B else 0)
    if data_directories == optional:
        raise RuntimeError(f"Unknown PE optional-header format: 0x{magic:04x}")
    certificate_offset, certificate_size = struct.unpack_from(
        "<II", data, data_directories + (4 * 8)
    )
    if certificate_offset + certificate_size > len(data):
        raise RuntimeError("PE certificate table extends beyond duckdb.exe")
    if certificate_size >= 8:
        _, revision, certificate_type = struct.unpack_from("<IHH", data, certificate_offset)
        if revision != 0x0200 or certificate_type != 0x0002:
            raise RuntimeError(
                "DuckDB certificate table is not an Authenticode PKCS#7 WIN_CERTIFICATE"
            )
    return certificate_offset, certificate_size


def build() -> Path:
    cache = ROOT / "build" / "downloads" / DUCKDB_ASSET
    download_verified(cache)

    build_root = ROOT / "build" / "cli-probe"
    stage = build_root / "WFMHub-DuckDB-CLI-Policy-Probe"
    if build_root.exists():
        shutil.rmtree(build_root)
    stage.mkdir(parents=True)

    with zipfile.ZipFile(cache) as archive:
        member = next((name for name in archive.namelist() if Path(name).name == "duckdb.exe"), None)
        if member is None:
            raise RuntimeError(f"duckdb.exe is missing from {DUCKDB_ASSET}")
        with archive.open(member) as source, (stage / "duckdb.exe").open("wb") as target:
            shutil.copyfileobj(source, target)

    executable_digest = sha256(stage / "duckdb.exe")
    if executable_digest != DUCKDB_EXE_SHA256:
        raise RuntimeError(
            f"duckdb.exe checksum mismatch: expected {DUCKDB_EXE_SHA256}, "
            f"received {executable_digest}"
        )

    certificate_offset, certificate_size = authenticode_table(stage / "duckdb.exe")
    if certificate_offset <= 0 or certificate_size <= 8:
        raise RuntimeError("Official duckdb.exe has no embedded Authenticode certificate table")

    source = ROOT / "tools" / "duckdb_cli_probe"
    shutil.copy2(source / "TEST-DUCKDB.cmd", stage / "TEST-DUCKDB.cmd")
    shutil.copy2(source / "README.txt", stage / "README.txt")
    shutil.copy2(source / "DUCKDB-LICENSE.txt", stage / "DUCKDB-LICENSE.txt")

    dist = ROOT / "dist"
    dist.mkdir(exist_ok=True)
    target = dist / f"WFMHub-DuckDB-CLI-Policy-Probe-v{DUCKDB_VERSION}-win-x64.zip"
    if target.exists():
        target.unlink()
    with zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED, compresslevel=9) as archive:
        for path in sorted(stage.rglob("*")):
            if path.is_file():
                archive.write(path, path.relative_to(build_root))

    digest = sha256(target)
    checksum = target.with_suffix(target.suffix + ".sha256")
    checksum.write_text(f"{digest}  {target.name}\n", encoding="utf-8")
    print(f"Built {target}")
    print(f"DuckDB ZIP SHA-256 {DUCKDB_SHA256}")
    print(f"duckdb.exe SHA-256 {DUCKDB_EXE_SHA256}")
    print(f"duckdb.exe certificate table offset={certificate_offset} size={certificate_size}")
    print(f"Probe SHA-256 {digest}")
    return target


def main() -> int:
    build()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
