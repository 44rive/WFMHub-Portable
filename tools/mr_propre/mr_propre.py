"""Standalone, read-only Client ID membership filter. Made by Anass Assri."""
from __future__ import annotations

import argparse
import codecs
import csv
import re
import sys
from contextlib import contextmanager
from datetime import datetime
from pathlib import Path

VERSION = "1.0.1"
ID_ALIASES = {"AGENTID", "DATASOURCEID", "DATASOURCEIDS", "VERINTID"}
SUPPORTED = {".csv", ".txt"}
BANNER = r"""
 __  __       ____
|  \/  |_ __ |  _ \ _ __ ___  _ __  _ __ ___
| |\/| | '__|| |_) | '__/ _ \| '_ \| '__/ _ \
| |  | | |   |  __/| | | (_) | |_) | | |  __/
|_|  |_|_|   |_|   |_|  \___/| .__/|_|  \___|
                            |_|
Mr.Propre - Client ID Clean Extract
Made by Anass Assri
"""


def norm(value: object) -> str:
    return re.sub(r"[^A-Z0-9]", "", str(value or "").upper())


def as_id(value: object) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def id_column(headers: list, required: bool = True) -> int | None:
    # An explicit Client ID wins over operational IDs in the same extract.
    matches = [i for i, h in enumerate(headers) if norm(h) == "CLIENTID"]
    if not matches:
        matches = [i for i, h in enumerate(headers) if norm(h) in ID_ALIASES]
    if len(matches) > 1:
        raise ValueError("Multiple ID columns; keep one authoritative Client ID column")
    if not matches and required:
        raise ValueError("No Client ID / Agent ID / Data Source IDs / Verint ID column")
    return matches[0] if matches else None


def detect_encoding(path: Path) -> str:
    with path.open("rb") as handle:
        prefix = handle.read(4)
    if prefix.startswith((codecs.BOM_UTF16_LE, codecs.BOM_UTF16_BE)):
        return "utf-16"
    # Validate the entire stream: a late accent must not break a large export.
    for encoding in ("utf-8-sig", "cp1252"):
        try:
            with path.open(encoding=encoding) as handle:
                while handle.read(1024 * 1024):
                    pass
            return encoding
        except UnicodeDecodeError:
            continue
    raise ValueError("Unrecognized encoding; save the source as UTF-8 or Windows-1252")


@contextmanager
def text_rows(path: Path):
    with path.open(encoding=detect_encoding(path), newline="") as handle:
        sample = handle.read(65536)
        handle.seek(0)
        # Sniff the header rather than counting delimiters inside data values.
        header = sample.splitlines()[0] if sample else ""
        try:
            delimiter = csv.Sniffer().sniff(header, delimiters=",;\t|").delimiter
        except csv.Error:
            delimiter = "\t" if path.suffix.lower() == ".txt" else ","
        yield csv.reader(handle, delimiter=delimiter, strict=True)


def load_roster(path: Path, sheet_name: str | None = None) -> set[str]:
    if path.suffix.lower() == ".xlsx":
        from openpyxl import load_workbook
        workbook = load_workbook(path, read_only=True, data_only=True)
        try:
            if sheet_name and sheet_name not in workbook.sheetnames:
                raise ValueError(f"Sheet {sheet_name!r} not found")
            sheets = [workbook[sheet_name]] if sheet_name else workbook.worksheets
            candidates = []
            for sheet in sheets:
                for row_no, values in enumerate(sheet.iter_rows(values_only=True), 1):
                    if row_no > 100:
                        break
                    column = id_column(list(values), required=False)
                    if column is not None:
                        title = norm(sheet.title)
                        rank = 2 if title in {"AGENT", "AGENTS", "ROSTER", "FTE", "FTECOUNT", "HEADCOUNT", "AGENTLIST", "AGENTROSTER"} else 1
                        candidates.append((rank, sheet, row_no, column))
                        break
            if not candidates:
                raise ValueError("No ID header found in the first 100 rows of the workbook")
            best = [c for c in candidates if c[0] == max(c[0] for c in candidates)]
            if len(best) != 1:
                raise ValueError("Multiple roster sheets found; select one using --sheet")
            _, sheet, row_no, column = best[0]
            ids = set()
            for row in sheet.iter_rows(min_row=row_no + 1):
                if column >= len(row):
                    continue
                cell = row[column]
                key = as_id(cell.value)
                # Excel numeric IDs with a pure zero format display leading zeros.
                if key and isinstance(cell.value, (int, float)) and re.fullmatch(r"0+", cell.number_format):
                    key = key.zfill(len(cell.number_format))
                if key:
                    ids.add(key)
        finally:
            workbook.close()
    elif path.suffix.lower() in SUPPORTED:
        if sheet_name:
            raise ValueError("--sheet applies only to XLSX listings")
        with text_rows(path) as rows:
            headers = next(rows, None)
            if not headers:
                raise ValueError("The listing is empty")
            column = id_column(headers)
            ids = set()
            for line, row in enumerate(rows, 2):
                if not row:
                    continue
                if len(row) != len(headers):
                    raise ValueError(f"Malformed listing row {line}: expected {len(headers)} fields")
                key = as_id(row[column])
                if key:
                    ids.add(key)
    else:
        raise ValueError("Listing must be CSV, TXT or XLSX")
    if not ids:
        raise ValueError("The listing contains no Client IDs")
    return ids


def parse_selection(selection: str, count: int) -> list[int]:
    if selection.strip().lower() in {"a", "all"}:
        return list(range(count))
    choices = set()
    for part in selection.split(","):
        match = re.fullmatch(r"\s*(\d+)\s*(?:-\s*(\d+)\s*)?", part)
        if not match:
            raise ValueError("Choose numbers, ranges (1,3-5), or A for all")
        start, end = int(match[1]), int(match[2] or match[1])
        if not 1 <= start <= end <= count:
            raise ValueError(f"Selection must be between 1 and {count}")
        choices.update(range(start - 1, end))
    if not choices:
        raise ValueError("No files selected")
    return sorted(choices)


def choose(paths: list[Path], folder: Path, multiple: bool = True) -> list[Path]:
    for number, path in enumerate(paths, 1):
        print(f" {number:>3}. {path.relative_to(folder)}")
    while True:
        selection = input("Choose files (1,3-5 or A): " if multiple else "Choose listing number: ")
        try:
            indexes = parse_selection(selection, len(paths))
            if not multiple and len(indexes) != 1:
                raise ValueError("Choose exactly one listing")
            return [paths[i] for i in indexes]
        except ValueError as exc:
            print(f"{exc}. Try again; Ctrl+C cancels.")


def filter_file(source: Path, destination: Path, ids: set[str]) -> tuple[int, int]:
    if source.resolve() == destination.resolve():
        raise ValueError("Source and destination must differ")
    if destination.exists():
        raise ValueError(f"Output already exists: {destination}")
    destination.parent.mkdir(parents=True, exist_ok=True)
    # Exclusive creation prevents overwriting existing outputs, even on a rerun.
    partial = destination.with_name(destination.name + ".partial")
    scanned = kept = 0
    with partial.open("x", encoding="utf-8-sig", newline="") as target:
        try:
            with text_rows(source) as rows:
                headers = next(rows, None)
                if not headers:
                    raise ValueError("Source is empty")
                column = id_column(headers)
                writer = csv.writer(target)
                writer.writerow(headers)
                for line, row in enumerate(rows, 2):
                    if not row:
                        continue
                    if len(row) != len(headers):
                        raise ValueError(f"Malformed row {line}: {len(row)} fields instead of {len(headers)}")
                    scanned += 1
                    if as_id(row[column]) in ids:
                        writer.writerow(row)
                        kept += 1
            target.flush()
        except BaseException:
            target.close()
            partial.unlink(missing_ok=True)
            raise
    try:
        # Copy with exclusive creation; rename on POSIX could overwrite a racing file.
        with partial.open("rb") as source_handle, destination.open("xb") as output_handle:
            import shutil
            try:
                shutil.copyfileobj(source_handle, output_handle)
            except BaseException:
                output_handle.close()
                destination.unlink(missing_ok=True)
                raise
    finally:
        partial.unlink(missing_ok=True)
    return scanned, kept


def prompt_path(label: str) -> Path:
    value = input(label).strip().strip('"')
    if not value:
        raise ValueError("A path is required")
    return Path(value)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Mr.Propre: filter extracts by Client ID membership only")
    parser.add_argument("--version", action="version", version=f"Mr.Propre {VERSION}")
    parser.add_argument("--folder", type=Path, help="Folder containing source CSV/TXT extracts")
    parser.add_argument("--roster", type=Path, help="Listing CSV/TXT/XLSX, or a folder to choose it from")
    parser.add_argument("--sheet", help="XLSX roster sheet if automatic detection is ambiguous")
    parser.add_argument("--output", type=Path, help="Separate destination folder")
    parser.add_argument("--select", help="File numbers/ranges, e.g. 1,3-5 (same order as menu)")
    parser.add_argument("--all", action="store_true", help="Process all CSV/TXT sources")
    parser.add_argument("--flat", action="store_true", help="Scan only the source folder, not subfolders")
    args = parser.parse_args(argv)
    try:
        print(BANNER)
        folder = (args.folder or prompt_path("Source folder: ")).expanduser().resolve()
        roster = (args.roster or prompt_path("Listing file or folder (CSV/XLSX): ")).expanduser().resolve()
        if not folder.is_dir():
            raise ValueError(f"Source folder not found: {folder}")
        if roster.is_dir():
            listings = sorted((p for p in roster.iterdir() if p.is_file() and p.suffix.lower() in SUPPORTED | {".xlsx"} and not p.name.startswith("~$")), key=lambda p: p.name.casefold())
            if not listings:
                raise ValueError("No CSV/TXT/XLSX listings in that folder")
            roster = listings[0] if len(listings) == 1 else choose(listings, roster, multiple=False)[0]
        if not roster.is_file():
            raise ValueError(f"Listing not found: {roster}")
        if args.output:
            output = args.output.expanduser().resolve()
        else:
            default = folder.parent / ("Mr.Propre-" + datetime.now().strftime("%Y%m%d-%H%M%S"))
            value = input(f"Output folder [Enter = {default}]: ").strip().strip('"')
            output = Path(value).expanduser().resolve() if value else default
        if output == folder or output.is_relative_to(folder) or folder.is_relative_to(output):
            raise ValueError("Output must be separate from the source folder, outside its folder tree")
        if roster == output or roster.is_relative_to(output):
            raise ValueError("Output must not contain the listing")
        ids = load_roster(roster, args.sheet)
        iterator = folder.iterdir() if args.flat else folder.rglob("*")
        paths = sorted((p for p in iterator if p.is_file() and not p.is_symlink() and p.suffix.lower() in SUPPORTED and p.resolve() != roster), key=lambda p: str(p).casefold())
        if not paths:
            raise ValueError("No CSV/TXT sources found")
        print(f"\nListing: {len(ids):,} unique IDs. Filtering by ID membership only.\n")
        if args.all and args.select:
            raise ValueError("Use either --all or --select")
        selected = paths if args.all else ([paths[i] for i in parse_selection(args.select, len(paths))] if args.select else choose(paths, folder))
        targets = [output / p.relative_to(folder).with_suffix(".csv") for p in selected]
        if len({str(p).casefold() for p in targets}) != len(targets):
            raise ValueError("Selected filenames collide after CSV conversion; select them separately")
        for target in targets:
            if target.exists() or target.with_name(target.name + ".partial").exists():
                raise ValueError(f"Output already exists: {target}; choose a new output folder")
        failures = 0
        for source, target in zip(selected, targets):
            try:
                scanned, kept = filter_file(source, target, ids)
                print(f"OK   {source.relative_to(folder)}: kept {kept:,} / {scanned:,} rows")
            except (OSError, ValueError, UnicodeError, csv.Error) as exc:
                failures += 1
                print(f"FAIL {source.relative_to(folder)}: {exc}", file=sys.stderr)
        print(f"\nFinished: {len(selected) - failures} exported, {failures} failed.\nOutput: {output}")
        return 2 if failures else 0
    except (OSError, ValueError, UnicodeError, csv.Error, ImportError) as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 2
    except (KeyboardInterrupt, EOFError):
        print("\nCancelled.")
        return 130


if __name__ == "__main__":
    raise SystemExit(main())
