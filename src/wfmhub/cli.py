"""Command line and beginner-friendly interactive menu."""

from __future__ import annotations

import argparse
import logging
import os
import sys
import uuid
from datetime import date, datetime
from pathlib import Path

from . import __version__
from .actions import import_actions
from .config import ConfigError, ensure_user_config, load_config, write_source_root
from .database import HubLockedError, backup_database, connect, migrate, write_session
from .doctor import run_doctor
from .ingestion import ingest_all
from .models import refresh_models, resolve_period
from .reports import build_report


def _home(value: str | None) -> Path:
    if value:
        return Path(value).resolve()
    if os.environ.get("WFMHUB_HOME"):
        return Path(os.environ["WFMHUB_HOME"]).resolve()
    current = Path.cwd().resolve()
    if (current / "config" / "default.toml").exists():
        return current
    return Path(__file__).resolve().parents[2]


def _date(value: str) -> date:
    try:
        return date.fromisoformat(value)
    except ValueError as exc:
        raise argparse.ArgumentTypeError("Use YYYY-MM-DD") from exc


def _logging(config) -> Path:
    stamp = datetime.now().strftime("%Y%m%d")
    path = config.logs / f"wfmhub_{stamp}.log"
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s %(levelname)s %(message)s",
        handlers=[logging.FileHandler(path, encoding="utf-8"), logging.StreamHandler(sys.stdout)],
        force=True,
    )
    return path


def setup(home: Path, source_root: Path | None, non_interactive: bool) -> int:
    config_file = ensure_user_config(home)
    if source_root is None and not non_interactive:
        print("Paste the folder that contains FTE, Storm and Verint.")
        entered = input("Source root: ").strip().strip('"')
        source_root = Path(entered) if entered else None
    if source_root:
        write_source_root(config_file, source_root)
    config = load_config(home, config_file)
    log = _logging(config)
    migrations = migrate(config)
    print("\nSetup complete.")
    print(f"Source root : {config.source_root}")
    print(f"Database    : {config.database}")
    print(f"Log         : {log}")
    if migrations:
        print(f"Database migrations applied: {', '.join(migrations)}")
    print("Run WFMHub.cmd and choose Refresh + build report.")
    return 0


def refresh(home: Path, start: date | None, end: date | None, no_report: bool) -> int:
    config = load_config(home)
    _logging(config)
    run_id = uuid.uuid4().hex
    with write_session(config) as conn:
        conn.execute("INSERT INTO meta.refresh_run(run_id, started_at, requested_start, requested_end, status) VALUES (?, ?, ?, ?, 'RUNNING')", [run_id, datetime.now(), start, end])
        try:
            ingested = ingest_all(conn, config)
            model = refresh_models(conn, config, run_id, start, end)
            report_path = None if no_report else build_report(conn, config, model.start, model.end)
            conn.execute(
                """UPDATE meta.refresh_run SET finished_at=?, status='SUCCESS', files_loaded=?, files_skipped=?, files_failed=?, details=? WHERE run_id=?""",
                [datetime.now(), ingested.loaded, ingested.skipped, ingested.failed, f"attendance={model.attendance_rows}; gaps={model.correction_rows}; quality={model.quality_rows}; scoped_out={ingested.scoped_out}", run_id],
            )
        except Exception as exc:
            conn.execute("UPDATE meta.refresh_run SET finished_at=?, status='ERROR', details=? WHERE run_id=?", [datetime.now(), str(exc)[:4000], run_id])
            raise
    print("\nRefresh complete.")
    print(f"Period      : {model.start} to {model.end}")
    print(f"Files       : {ingested.loaded} loaded, {ingested.skipped} unchanged, {ingested.failed} failed")
    print(f"Agent scope : {ingested.scoped_out:,} outside-roster source rows excluded")
    print(f"Attendance  : {model.attendance_rows:,} rows")
    print(f"Gaps        : {model.correction_rows:,} rows")
    print(f"RTA         : {model.rta_rows:,} rows")
    print(f"Intraday    : {model.intraday_rows:,} actual + {model.forecast_rows:,} forecast rows")
    print(f"Quality     : {model.quality_rows:,} issues")
    if ingested.errors:
        print("\nFiles with errors:")
        for error in ingested.errors:
            print(f"- {error}")
    if report_path:
        print(f"Report      : {report_path}")
    return 2 if ingested.failed else 0


def report_only(home: Path, start: date | None, end: date | None, output: Path | None) -> int:
    config = load_config(home)
    _logging(config)
    conn = connect(config, read_only=True)
    try:
        start, end = resolve_period(conn, config, start, end)
        path = build_report(conn, config, start, end, output)
    finally:
        conn.close()
    print(f"Report created: {path}")
    return 0


def import_decisions(home: Path, workbook: Path) -> int:
    config = load_config(home)
    _logging(config)
    with write_session(config) as conn:
        count = import_actions(conn, workbook)
    print(f"Imported {count} correction decision(s).")
    return 0


def show_status(home: Path) -> int:
    config = load_config(home)
    conn = connect(config, read_only=True)
    try:
        rows = conn.execute("SELECT source_family, newest_business_date, row_count, scoped_out_count, status, newest_file FROM mart.source_health ORDER BY source_family").fetchall()
    finally:
        conn.close()
    print("\nSOURCE HEALTH")
    for family, business_date, rows_count, scoped_out, status, file_name in rows:
        date_text = str(business_date) if business_date else "-"
        print(f"{family:14} {status:8} date={date_text:10} kept={rows_count or 0:8,} excluded={scoped_out or 0:8,} file={file_name or '-'}")
    return 0


def create_backup(home: Path) -> int:
    config = load_config(home)
    path = backup_database(config)
    print(f"Backup created: {path}")
    return 0


def menu(home: Path) -> int:
    while True:
        print("\nWFMHUB PORTABLE")
        print("1. Refresh all available data + build report")
        print("2. Refresh current month + build report")
        print("3. Refresh a custom period + build report")
        print("4. Build report only")
        print("5. Import correction decisions from an edited report")
        print("6. Show source health")
        print("7. Create database backup")
        print("8. Change source root")
        print("9. Run system check")
        print("10. Exit")
        choice = input("Choose 1-10: ").strip()
        try:
            if choice == "1":
                refresh(home, None, None, False)
            elif choice == "2":
                today = date.today()
                refresh(home, today.replace(day=1), today, False)
            elif choice == "3":
                start = _date(input("Start date YYYY-MM-DD: ").strip())
                end = _date(input("End date YYYY-MM-DD: ").strip())
                refresh(home, start, end, False)
            elif choice == "4":
                report_only(home, None, None, None)
            elif choice == "5":
                path = Path(input("Paste the edited report path: ").strip().strip('"'))
                import_decisions(home, path)
            elif choice == "6":
                show_status(home)
            elif choice == "7":
                create_backup(home)
            elif choice == "8":
                path = Path(input("Paste the folder containing FTE, Storm and Verint: ").strip().strip('"'))
                setup(home, path, True)
            elif choice == "9":
                run_doctor(home)
            elif choice == "10":
                return 0
            else:
                print("Please choose a number from 1 to 10.")
        except Exception as exc:
            print(f"\nERROR: {exc}")
            print("Nothing was changed in your extract files. Check the latest file in logs.")


def parser() -> argparse.ArgumentParser:
    root = argparse.ArgumentParser(prog="wfmhub", description="Portable SQLite WFM hub")
    root.add_argument("--home", help="WFMHub folder (normally detected automatically)")
    root.add_argument("--version", action="version", version=f"WFMHub {__version__}")
    commands = root.add_subparsers(dest="command")
    setup_p = commands.add_parser("setup", help="Create config and database")
    setup_p.add_argument("--source-root", type=Path)
    setup_p.add_argument("--non-interactive", action="store_true")
    refresh_p = commands.add_parser("refresh", help="Ingest, model and report")
    refresh_p.add_argument("--start", type=_date)
    refresh_p.add_argument("--end", type=_date)
    refresh_p.add_argument("--no-report", action="store_true")
    report_p = commands.add_parser("report", help="Build an Excel report from current marts")
    report_p.add_argument("--start", type=_date)
    report_p.add_argument("--end", type=_date)
    report_p.add_argument("--output", type=Path)
    import_p = commands.add_parser("import-actions", help="Import edited GAPS decision columns")
    import_p.add_argument("workbook", type=Path)
    commands.add_parser("status", help="Show source health")
    commands.add_parser("backup", help="Create a database backup")
    commands.add_parser("doctor", help="Test the corporate runtime, SQLite and Excel libraries")
    commands.add_parser("menu", help="Open the interactive menu")
    return root


def main(argv: list[str] | None = None) -> int:
    args = parser().parse_args(argv)
    home = _home(args.home)
    try:
        if args.command == "setup":
            return setup(home, args.source_root, args.non_interactive)
        if args.command == "refresh":
            return refresh(home, args.start, args.end, args.no_report)
        if args.command == "report":
            return report_only(home, args.start, args.end, args.output)
        if args.command == "import-actions":
            return import_decisions(home, args.workbook)
        if args.command == "status":
            return show_status(home)
        if args.command == "backup":
            return create_backup(home)
        if args.command == "doctor":
            return 0 if run_doctor(home) else 1
        return menu(home)
    except (ConfigError, HubLockedError, FileNotFoundError, ValueError, RuntimeError) as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 1
    except KeyboardInterrupt:
        print("\nCancelled.")
        return 130
