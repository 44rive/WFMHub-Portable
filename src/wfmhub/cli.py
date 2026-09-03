"""Command line and beginner-friendly interactive menu."""

from __future__ import annotations

import argparse
import logging
import os
import sys
import uuid
from datetime import date, datetime, timedelta
from pathlib import Path

from . import __version__
from .actions import import_actions
from .analytics import load_analytics_rules, validate_analytics_rules
from .bonus import import_bonus_matrix
from .coaching import import_pcs_coaching
from .config import ConfigError, ensure_user_config, load_config, write_source_root
from .database import HubLockedError, backup_database, connect, migrate, write_session
from .doctor import run_doctor
from .excel_templates import excel_template, require_new_template
from .exports import DATASETS, export_dataset
from .ingestion import ingest_all
from .models import refresh_models
from .mapping import load_queue_mapping
from .metrics import diff_metric_catalogs, evaluate_metric, load_metric_catalog, validate_metric_catalog
from .on_demand_analysis import ANALYSIS_DOMAINS, COMPARISON_MODES, build_analysis_workbook
from .custom_jobs import list_jobs, run_python_job, run_sql_job
from .progress import ProgressBar, ProgressCallback
from .report_packs import IMPLEMENTED_REPORT_PACK_KEYS, build_report_pack
from .report_specs import load_report_catalog, validate_report_catalog
from .rules import load_rulebook, validate_rulebook
from .semantic import SOURCE_COMPONENTS
from .sota_reports import build_kpi_catalog
from .service_profiles import load_service_profiles, validate_service_profiles
from .ui import clear_screen, render_dashboard


SOURCE_GROUPS = {
    "all": None,
    "operations": {"fte", "schedule", "lilo", "agent_status"},
    "intraday": {"forecast", "apbe", "apfr", "apde"},
    "pcs": {"fte", "calls"},
}


def _phase_progress(bar: ProgressBar, start: float, end: float) -> ProgressCallback:
    """Map a component's progress into its part of the overall bar."""

    def report(current: int, total: int, label: str) -> None:
        if total > 0:
            fraction = min(1.0, max(0.0, current / total))
            bar.update(start + (end - start) * fraction, label)
        else:
            bar.pulse(label)

    return report


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
    bar = ProgressBar()
    try:
        bar.update(0.1, "Reading configuration")
        config = load_config(home, config_file)
        log = _logging(config)
        bar.update(0.4, "Preparing folders and database")
        migrations = migrate(config)
        bar.finish("Setup ready")
    except Exception as exc:
        bar.fail(str(exc))
        raise
    print("\nSetup complete.")
    print(f"Source root : {config.source_root}")
    print(f"Database    : {config.database}")
    print(f"Rules       : {config.business_rules}")
    print(f"Metrics     : {config.metric_catalog}")
    print(f"Analytics   : {config.analytics_rules}")
    print(f"Reports     : {config.report_catalog}")
    print(f"Queue map   : {config.queue_mapping}")
    print(f"Service LOBs: {config.service_profiles}")
    print(f"Log         : {log}")
    if migrations:
        print(f"Database migrations applied: {', '.join(migrations)}")
    print("Run WFMHub.cmd, choose Refresh hub data, then select dates and reports.")
    return 0


def refresh(
    home: Path,
    start: date | None,
    end: date | None,
    packs: tuple[str, ...] = ("attendance",),
    source_group: str = "all",
    use_config_period: bool = True,
    service_profile: str | None = None,
) -> int:
    config = load_config(home)
    _logging(config)
    run_id = uuid.uuid4().hex
    bar = ProgressBar()
    bar.update(0.01, "Opening hub database")
    try:
        with write_session(config) as conn:
            conn.execute("INSERT INTO meta.refresh_run(run_id, started_at, requested_start, requested_end, status) VALUES (?, ?, ?, ?, 'RUNNING')", [run_id, datetime.now(), start, end])
            try:
                ingested = ingest_all(
                    conn, config, SOURCE_GROUPS[source_group],
                    _phase_progress(bar, 0.03, 0.55),
                )
                model = refresh_models(
                    conn, config, run_id, start, end, use_config_period,
                    _phase_progress(bar, 0.55, 0.85),
                )
                report_paths = []
                total_packs = len(packs)
                for index, pack in enumerate(packs):
                    report_start = 0.85 + (0.14 * index / max(1, total_packs))
                    report_end = 0.85 + (0.14 * (index + 1) / max(1, total_packs))
                    bar.update(report_start, f"Writing {pack} report")
                    report_paths.append(build_report_pack(
                        pack, conn, config, model.start, model.end,
                        service_profile=service_profile,
                    ))
                    bar.update(report_end, f"Created {pack} report")
                conn.execute(
                    """UPDATE meta.refresh_run SET finished_at=?, status='SUCCESS', files_loaded=?, files_skipped=?, files_failed=?, details=? WHERE run_id=?""",
                    [datetime.now(), ingested.loaded, ingested.skipped, ingested.failed, f"attendance={model.attendance_rows}; absence={model.absence_rows}; service={model.service_rows}; gaps={model.correction_rows}; metrics={model.metric_rows}; findings={model.finding_rows}; quality={model.quality_rows}; scoped_out={ingested.scoped_out}", run_id],
                )
            except Exception as exc:
                conn.execute("UPDATE meta.refresh_run SET finished_at=?, status='ERROR', details=? WHERE run_id=?", [datetime.now(), str(exc)[:4000], run_id])
                raise
        bar.finish("Refresh complete")
    except Exception as exc:
        bar.fail(str(exc))
        raise
    print("\nRefresh complete.")
    print(f"Period      : {model.start} to {model.end}")
    print(f"Files       : {ingested.loaded} loaded, {ingested.skipped} unchanged, {ingested.failed} failed")
    print(f"Agent scope : {ingested.scoped_out:,} outside-roster source rows excluded")
    print(f"Attendance  : {model.attendance_rows:,} rows")
    print(f"Gaps        : {model.correction_rows:,} rows")
    print(f"Absence     : {model.absence_rows:,} agent-day + {model.absence_event_rows:,} evidence rows")
    print(f"Service     : {model.service_rows:,} actual + {model.forecast_rows:,} forecast rows")
    print(f"Agent PCS   : {model.pcs_rows:,} agent-day rows")
    print(f"Metrics     : {model.metric_rows:,} governed values")
    print(f"Findings    : {model.finding_rows:,} deterministic observations")
    print(f"Quality     : {model.quality_rows:,} issues")
    business = load_rulebook(home, config.business_rules)
    mapping = load_queue_mapping(config.queue_mapping)
    print(f"Rules       : {business.version} ({business.sha256[:12]})")
    print(f"Queue map   : {mapping.sha256[:12]}")
    if ingested.errors:
        print("\nFiles with errors:")
        for error in ingested.errors:
            print(f"- {error}")
    for pack, report_path in zip(packs, report_paths):
        print(f"Report      : {report_path}")
        template = excel_template(config, pack)
        if template.exists:
            print(f"Excel master: {template.path} (open it and choose Refresh All)")
    return 2 if ingested.failed else 0


def report_only(
    home: Path,
    start: date | None,
    end: date | None,
    output: Path | None,
    pack: str = "pcs",
    use_config_period: bool = True,
    service_profile: str | None = None,
) -> int:
    paths = build_reports(home, start, end, (pack,), output, use_config_period, service_profile)
    print(f"Report created: {paths[0]}")
    template = excel_template(load_config(home), pack)
    if template.exists and paths[0].resolve() != template.path:
        print(f"Excel master  : {template.path}")
        print("Next step     : open the master in Excel, choose Refresh All, then Save As.")
    return 0


def initialize_excel_template(
    home: Path,
    pack: str,
    start: date | None,
    end: date | None,
    service_profile: str | None = None,
    force: bool = False,
    use_config_period: bool = True,
) -> int:
    """Create a styled master that the user enriches once in desktop Excel."""

    config = load_config(home)
    existing = excel_template(config, pack)
    if existing.exists and pack == "pcs" and not force:
        print(f"Excel starter : {existing.path}")
        print(f"Current feed  : {existing.feed_folder}")
        print("Status        : existing PCS master kept; WFMHub did not overwrite it.")
        print(f"One-time steps: {config.home / 'docs' / 'EXCEL_TEMPLATE_GUIDE.md'}")
        return 0
    template = require_new_template(config, pack, force)
    path = build_reports(
        home, start, end, (pack,), template.path, use_config_period,
        service_profile,
    )[0]
    print(f"Excel starter : {path}")
    print(f"Current feed  : {template.feed_folder}")
    print("Protected rule: normal WFMHub refreshes never overwrite this master.")
    print(f"One-time steps: {config.home / 'docs' / 'EXCEL_TEMPLATE_GUIDE.md'}")
    return 0


def build_reports(
    home: Path,
    start: date | None,
    end: date | None,
    packs: tuple[str, ...],
    output: Path | None = None,
    use_config_period: bool = True,
    service_profile: str | None = None,
) -> list[Path]:
    config = load_config(home)
    _logging(config)
    if output is not None and len(packs) != 1:
        raise ValueError("An explicit output path can be used with only one report pack")
    bar = ProgressBar()
    bar.update(0.02, "Opening hub database")
    try:
        with write_session(config) as conn:
            model = refresh_models(
                conn, config, f"report-{uuid.uuid4().hex}", start, end,
                use_config_period, _phase_progress(bar, 0.05, 0.70),
            )
            paths = []
            for index, pack in enumerate(packs):
                report_start = 0.70 + (0.29 * index / max(1, len(packs)))
                report_end = 0.70 + (0.29 * (index + 1) / max(1, len(packs)))
                bar.update(report_start, f"Writing {pack} report")
                paths.append(build_report_pack(
                    pack, conn, config, model.start, model.end, output,
                    service_profile=service_profile,
                ))
                bar.update(report_end, f"Created {pack} report")
        bar.finish("Reports complete")
        return paths
    except Exception as exc:
        bar.fail(str(exc))
        raise


def export_clean(
    home: Path,
    dataset: str,
    start: date | None,
    end: date | None,
    file_format: str,
    output: Path | None = None,
    use_config_period: bool = True,
) -> int:
    config = load_config(home)
    _logging(config)
    bar = ProgressBar()
    bar.update(0.02, "Opening hub database")
    try:
        with write_session(config) as conn:
            model = refresh_models(
                conn, config, f"export-{uuid.uuid4().hex}", start, end,
                use_config_period, _phase_progress(bar, 0.05, 0.48),
            )
            bar.update(0.50, f"Preparing {dataset} export")
            result = export_dataset(
                conn, config, dataset, model.start, model.end, file_format, output,
                _phase_progress(bar, 0.50, 0.99),
            )
        bar.finish("Export complete")
    except Exception as exc:
        bar.fail(str(exc))
        raise
    print(f"Clean export : {result.path}")
    print(f"Rows         : {result.rows:,}")
    print(f"Manifest     : {result.manifest}")
    return 0


def run_custom(
    home: Path,
    kind: str,
    job: Path,
    start: date | None,
    end: date | None,
    use_config_period: bool = True,
) -> int:
    config = load_config(home)
    _logging(config)
    bar = ProgressBar()
    bar.update(0.02, "Opening hub database")
    try:
        with write_session(config) as conn:
            model = refresh_models(
                conn, config, f"custom-{uuid.uuid4().hex}", start, end,
                use_config_period, _phase_progress(bar, 0.05, 0.65),
            )
        bar.pulse(f"Running custom {kind}: {job.name}")
        conn = connect(config, read_only=True)
        try:
            result = (
                run_python_job(conn, config, job, model.start, model.end)
                if kind == "python"
                else run_sql_job(conn, config, job, model.start, model.end)
            )
        finally:
            conn.close()
        bar.finish("Custom job complete")
    except Exception as exc:
        bar.fail(str(exc))
        raise
    print(f"Custom job  : {result.job}")
    print(f"Output      : {result.output_dir}")
    if result.result is not None:
        print(f"Result      : {result.result}")
    return 0


def import_decisions(home: Path, workbook: Path) -> int:
    config = load_config(home)
    _logging(config)
    bar = ProgressBar()
    bar.update(0.1, "Reading correction decisions")
    try:
        with write_session(config) as conn:
            count = import_actions(conn, workbook)
        bar.finish("Decisions imported")
    except Exception as exc:
        bar.fail(str(exc))
        raise
    print(f"Imported {count} correction decision(s).")
    return 0


def import_coaching_decisions(home: Path, workbook: Path) -> int:
    config = load_config(home)
    _logging(config)
    bar = ProgressBar()
    bar.update(0.1, "Reading PCS coaching decisions")
    try:
        with write_session(config) as conn:
            count = import_pcs_coaching(conn, config, workbook)
        bar.finish("PCS coaching decisions imported")
    except Exception as exc:
        bar.fail(str(exc))
        raise
    print(f"Imported {count} PCS coaching decision(s).")
    print("Build PCS Performance again to refresh Actions Rate and the stable Excel feeds.")
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


def show_coverage(home: Path) -> int:
    config = load_config(home)
    conn = connect(config, read_only=True)
    queries = [
        ("Schedule", "SELECT min(schedule_date), max(schedule_date), count(*) FROM raw.schedule_shift r JOIN meta.source_file f ON f.file_id=r.source_file_id AND f.active"),
        ("LILO", "SELECT min(extract_date), max(extract_date), count(*) FROM raw.lilo r JOIN meta.source_file f ON f.file_id=r.source_file_id AND f.active"),
        ("Agent Status", "SELECT min(extract_date), max(extract_date), count(*) FROM raw.agent_status r JOIN meta.source_file f ON f.file_id=r.source_file_id AND f.active"),
        ("Calls", "SELECT min(business_date), max(business_date), count(*) FROM core.clean_call_leg"),
        ("Actuals", "SELECT min(business_date), max(business_date), count(*) FROM raw.queue_actual r JOIN meta.source_file f ON f.file_id=r.source_file_id AND f.active"),
        ("Forecast", "SELECT min(business_date), max(business_date), count(*) FROM raw.forecast_interval r JOIN meta.source_file f ON f.file_id=r.source_file_id AND f.active"),
        ("Bonus", "SELECT min(period), max(period), count(*) FROM mart.bonus_agent_month"),
    ]
    try:
        print("\nDATA COVERAGE")
        for label, sql in queries:
            first, last, rows = conn.execute(sql).fetchone()
            print(f"{label:14} {str(first or '-'):10} to {str(last or '-'):10} rows={rows or 0:,}")
    finally:
        conn.close()
    return 0


def create_backup(home: Path) -> int:
    config = load_config(home)
    bar = ProgressBar()
    bar.update(0.1, "Copying and verifying database")
    try:
        path = backup_database(config)
        bar.finish("Backup complete")
    except Exception as exc:
        bar.fail(str(exc))
        raise
    print(f"Backup created: {path}")
    return 0


_METRIC_TEST_COMPONENTS = {
    "service_interval": {
        "offered": 100, "answered": 90, "abandoned": 10, "short_abandoned": 5,
        "answered_within_target": 75, "handled_seconds": 27_000,
    },
    "forecast_comparison_hour": {"forecast_volume": 100, "actual_volume": 90},
    "pcs_agent_day": {
        "handled_calls": 100, "talk_seconds": 20_000, "hold_seconds": 4_000,
        "wrap_seconds": 6_000, "handle_seconds": 30_000, "pcs_status_calls": 100,
        "pcs_participation_responses": 14, "survey_responses": 10,
        "pcs_score_count": 10, "pcs_score_sum": 45,
        "q1_response_count": 10, "q1_score_sum": 45,
        "q2_response_count": 8, "q2_score_sum": 32,
        "top_box_responses": 7, "low_score_responses": 3,
    },
    "observed_absence_agent_day": {
        "planned_net_minutes": 525, "absence_minutes": 60, "vacation_minutes": 0,
        "unpaid_minutes": 0, "shrinkage_minutes": 90,
    },
    "final_absence_agent_day": {
        "planned_net_minutes": 525, "final_absence_minutes": 60,
        "final_vacation_minutes": 0, "final_unpaid_minutes": 0,
        "final_shrinkage_minutes": 90,
    },
    "staffing_interval": {
        "staffing_gap_fte": 1.5, "staffing_variance_fte": -1.5,
        "scheduled_fte": 8, "observed_fte": 6.5, "productive_fte": 6,
        "evidence_intervals": 1,
    },
    "attendance_agent_day": {
        "scheduled_working_count": 1, "no_show_count": 0, "late_count": 1,
        "requires_call_count": 1, "uncoded_late_minutes": 10, "no_show_minutes": 0,
    },
}


def rules_tool(
    home: Path,
    action: str = "validate",
    metric_id: str | None = None,
    against: Path | None = None,
) -> int:
    config = load_config(home)
    rulebook = load_rulebook(home, config.business_rules)
    catalog = load_metric_catalog(home, config.metric_catalog)
    analytics = load_analytics_rules(home, config.analytics_rules)
    reports = load_report_catalog(home, config.report_catalog)
    mapping = load_queue_mapping(config.queue_mapping)
    service_profiles = load_service_profiles(home, config.service_profiles)
    if action == "explain":
        if not metric_id:
            raise ValueError("rules explain requires a metric id")
        print("\n".join(catalog.explain(metric_id)))
        return 0
    if action == "diff":
        if against is None:
            raise ValueError("rules diff requires --against PATH_TO_OLD_METRIC_CATALOG")
        before = load_metric_catalog(home, against.resolve())
        print("\n".join(diff_metric_catalogs(before, catalog)))
        return 0
    for lines in (
        validate_rulebook(rulebook),
        validate_metric_catalog(catalog, SOURCE_COMPONENTS),
        validate_analytics_rules(analytics, catalog),
        validate_report_catalog(reports, IMPLEMENTED_REPORT_PACK_KEYS),
        validate_service_profiles(service_profiles, catalog),
    ):
        for line in lines:
            print(line)
    print(f"Queue mapping is valid: {mapping.file}")
    print(f"Queue mapping SHA-256: {mapping.sha256}")
    if action == "catalog":
        path = build_kpi_catalog(config)
        print(f"Governance catalog: {path}")
    elif action == "test":
        for method in catalog.methods:
            dimensions = {}
            for key, values in method.scope.items():
                dimensions[key.removesuffix("_contains")] = values[0] if values else None
            result = evaluate_metric(
                method, _METRIC_TEST_COMPONENTS[method.source_model],
            )
            print(
                f"PASS {method.metric_id}.{method.method_id} "
                f"effective={method.effective_from} value={result.value} state={result.state}"
            )
    print(f"Domain rules : {rulebook.file}")
    print(f"Metric catalog: {catalog.file}")
    print(f"Analytics    : {analytics.file}")
    print(f"Report specs : {reports.file}")
    print(f"Service LOBs : {service_profiles.file}")
    return 0


def import_bonus_tool(
    home: Path,
    source: Path,
) -> int:
    """Import and calculate one Bonus Matrix period without editing the source."""

    config = load_config(home)
    _logging(config)
    bar = ProgressBar()
    try:
        bar.update(0.1, "Reading Bonus Matrix v1.2")
        with write_session(config) as conn:
            result = import_bonus_matrix(conn, source)
        bar.finish("Bonus period imported")
    except Exception as exc:
        bar.fail(str(exc))
        raise
    print(f"Period       : {result.period}")
    print(f"Agents       : {result.agents:,}")
    print(f"KPI rules    : {result.rules:,}")
    print(f"Policies     : {result.policies:,}")
    print(f"Source hash  : {result.import_id}")
    print(f"Status       : {'unchanged' if result.unchanged else 'new active version'}")
    print("Source file  : unchanged")
    return 0


def analyze_period(
    home: Path,
    domain: str,
    start: date | None,
    end: date | None,
    comparison: str,
    output: Path | None = None,
    use_config_period: bool = True,
) -> int:
    config = load_config(home)
    _logging(config)
    bar = ProgressBar()
    try:
        bar.update(0.02, "Opening hub database")
        with write_session(config) as conn:
            model = refresh_models(
                conn, config, f"analysis-{uuid.uuid4().hex}", start, end,
                use_config_period, _phase_progress(bar, 0.05, 0.72),
            )
            bar.update(0.75, f"Analyzing {domain}")
            path = build_analysis_workbook(
                conn, config, domain, model.start, model.end, comparison, output,
            )
        bar.finish("Analysis complete")
    except Exception as exc:
        bar.fail(str(exc))
        raise
    print(f"Analysis     : {path}")
    print(f"Domain       : {domain}")
    print(f"Period       : {model.start} to {model.end}")
    print(f"Comparison   : {comparison}")
    return 0


def _choose_period() -> tuple[date | None, date | None, bool]:
    print("\nDATE PERIOD")
    print("1. Today")
    print("2. Yesterday")
    print("3. Current week")
    print("4. Current month")
    print("5. Previous month")
    print("6. Custom dates")
    print("7. All available dates")
    print("8. Saved default dates")
    choice = input("Choose 1-8: ").strip()
    today = date.today()
    if choice == "1":
        return today, today, False
    if choice == "2":
        yesterday = today - timedelta(days=1)
        return yesterday, yesterday, False
    if choice == "3":
        return today - timedelta(days=today.weekday()), today, False
    if choice == "4":
        return today.replace(day=1), today, False
    if choice == "5":
        last_day = today.replace(day=1) - timedelta(days=1)
        return last_day.replace(day=1), last_day, False
    if choice == "6":
        start, end = (
            _date(input("Start date YYYY-MM-DD: ").strip()),
            _date(input("End date YYYY-MM-DD: ").strip()),
        )
        return start, end, False
    if choice == "7":
        return None, None, False
    if choice == "8":
        return None, None, True
    raise ValueError("Please choose a date option from 1 to 8")


def _choose_source_group() -> str:
    print("\nDATA TO REFRESH")
    print("1. All sources")
    print("2. Attendance/absence: FTE, Verint schedule, LILO and Agent Status")
    print("3. Service: APBE, APFR, APDE and Verint Forecast")
    print("4. Agent PCS: FTE and Call by Call")
    choice = input("Choose 1-4: ").strip()
    try:
        return {"1": "all", "2": "operations", "3": "intraday", "4": "pcs"}[choice]
    except KeyError as exc:
        raise ValueError("Please choose a data group from 1 to 4") from exc


def _choose_packs(allow_none: bool = True) -> tuple[str, ...]:
    print("\nREPORT PRODUCTS")
    print("1. PCS Performance: daily, MTD and previous month")
    print("2. Bonus Performance: calculation and release controls")
    print("3. Service Performance: mapped LOB queues and forecast")
    print("4. Staffing & Coverage: LOB/language interval gaps")
    print("5. Attendance Callouts: selected dates or live today")
    print("6. Attendance Corrections: completed-day gaps + timeline")
    print("7. Final Absence & Shrinkage: corrected Verint ledger")
    print("8. Build all seven separate workbooks")
    if allow_none:
        print("9. No report")
    choice = input(f"Choose 1-{'9' if allow_none else '8'}: ").strip()
    mapping = {
        "1": ("pcs",),
        "2": ("bonus",),
        "3": ("service",),
        "4": ("staffing",),
        "5": ("attendance",),
        "6": ("corrections",),
        "7": ("absence",),
        "8": IMPLEMENTED_REPORT_PACK_KEYS,
    }
    if allow_none and choice == "9":
        return ()
    try:
        return mapping[choice]
    except KeyError as exc:
        raise ValueError("Please choose a valid report option") from exc


def _choose_template_pack() -> str:
    labels = {
        "pcs": "PCS Performance",
        "bonus": "Bonus Performance",
        "service": "Service Performance",
        "staffing": "Staffing & Coverage",
        "attendance": "Attendance Callouts",
        "corrections": "Attendance Corrections",
        "absence": "Final Absence & Shrinkage",
    }
    print("\nEXCEL PIVOT/SLICER MASTER")
    for index, key in enumerate(IMPLEMENTED_REPORT_PACK_KEYS, 1):
        print(f"{index}. {labels[key]}")
    selected = int(input(f"Choose 1-{len(IMPLEMENTED_REPORT_PACK_KEYS)}: ").strip())
    if selected not in range(1, len(IMPLEMENTED_REPORT_PACK_KEYS) + 1):
        raise ValueError("Please choose a listed report")
    return IMPLEMENTED_REPORT_PACK_KEYS[selected - 1]


def _choose_service_profile(home: Path) -> str:
    config = load_config(home)
    catalog = load_service_profiles(home, config.service_profiles)
    active = [profile for profile in catalog.profiles if profile.active_on(date.today())]
    print("\nSERVICE PROFILE")
    for index, profile in enumerate(active, 1):
        default = " [default]" if profile.profile_id == catalog.default_profile else ""
        print(f"{index}. {profile.label}{default}")
    choice = input(f"Choose 1-{len(active)} [default]: ").strip()
    if not choice:
        return catalog.default_profile
    selected = int(choice)
    if selected not in range(1, len(active) + 1):
        raise ValueError("Please choose a listed service profile")
    return active[selected - 1].profile_id


def _choose_analysis() -> tuple[str, str]:
    print("\nON-DEMAND ANALYSIS")
    for index, domain in enumerate(ANALYSIS_DOMAINS, 1):
        print(f"{index}. {domain.title()}")
    selected = int(input(f"Choose domain 1-{len(ANALYSIS_DOMAINS)}: ").strip())
    if selected not in range(1, len(ANALYSIS_DOMAINS) + 1):
        raise ValueError("Please choose a listed analysis domain")
    print("\nCOMPARISON")
    labels = {
        "previous_equal": "Previous equal-length period",
        "previous_month": "Previous-month same days",
        "target": "Configured target",
        "none": "No comparison",
    }
    for index, mode in enumerate(COMPARISON_MODES, 1):
        print(f"{index}. {labels[mode]}")
    comparison = int(input(f"Choose comparison 1-{len(COMPARISON_MODES)}: ").strip())
    if comparison not in range(1, len(COMPARISON_MODES) + 1):
        raise ValueError("Please choose a listed comparison")
    return ANALYSIS_DOMAINS[selected - 1], COMPARISON_MODES[comparison - 1]


def _choose_dataset() -> str:
    keys = tuple(DATASETS)
    print("\nCLEAN DATASET")
    for index, key in enumerate(keys, 1):
        print(f"{index}. {key} - {DATASETS[key].description}")
    choice = int(input(f"Choose 1-{len(keys)}: ").strip())
    if choice not in range(1, len(keys) + 1):
        raise ValueError("Please choose a listed dataset")
    return keys[choice - 1]


def _choose_custom_job(config) -> tuple[str, Path]:
    print("\nCUSTOM LAB")
    print("1. Python job")
    print("2. Read-only SQL job")
    choice = input("Choose 1-2: ").strip()
    kind = "python" if choice == "1" else "sql" if choice == "2" else None
    if kind is None:
        raise ValueError("Please choose Python or SQL")
    jobs = list_jobs(config, kind)
    if not jobs:
        folder = config.custom / ("jobs" if kind == "python" else "sql")
        raise FileNotFoundError(
            f"No runnable {kind} jobs found in {folder}. Copy the underscore template and rename it first."
        )
    for index, path in enumerate(jobs, 1):
        print(f"{index}. {path.name}")
    selected = int(input(f"Choose 1-{len(jobs)}: ").strip())
    if selected not in range(1, len(jobs) + 1):
        raise ValueError("Please choose a listed custom job")
    return kind, jobs[selected - 1]


def _pause_for_dashboard() -> None:
    try:
        interactive = sys.stdin.isatty()
    except (AttributeError, OSError):
        interactive = False
    if interactive:
        input("\nPress Enter to return to the dashboard...")


def menu(home: Path) -> int:
    while True:
        clear_screen()
        render_dashboard(home)
        print("\n  DAILY WORK")
        print("    [1] Refresh hub data")
        print("    [2] Build reports from existing hub data")
        print("    [3] Analyze a period")
        print("    [4] Export clean data")
        print("    [5] Import Bonus Matrix v1.2")
        print("\n  CONTROL & REVIEW")
        print("    [6] Import attendance correction decisions")
        print("    [7] Import PCS coaching decisions")
        print("    [8] Validate rules and build governance catalog")
        print("    [9] Show source health and date coverage")
        print("\n  HUB TOOLS")
        print("   [10] Create database backup")
        print("   [11] Change source root")
        print("   [12] Run system check")
        print("   [13] Advanced: custom Python or read-only SQL")
        print("   [14] Create an Excel Pivot/slicer master")
        print("   [15] Exit")
        choice = input("\n  Choose 1-15: ").strip()
        try:
            if choice == "1":
                group = _choose_source_group()
                start, end, use_config = _choose_period()
                packs = _choose_packs(True)
                profile = _choose_service_profile(home) if "service" in packs else None
                refresh(home, start, end, packs, group, use_config, profile)
            elif choice == "2":
                start, end, use_config = _choose_period()
                packs = _choose_packs(False)
                profile = _choose_service_profile(home) if "service" in packs else None
                for path in build_reports(home, start, end, packs, use_config_period=use_config, service_profile=profile):
                    print(f"Report created: {path}")
            elif choice == "3":
                domain, comparison = _choose_analysis()
                start, end, use_config = _choose_period()
                analyze_period(home, domain, start, end, comparison, use_config_period=use_config)
            elif choice == "4":
                dataset = _choose_dataset()
                start, end, use_config = _choose_period()
                file_format = input("Format CSV or XLSX [CSV]: ").strip().lower() or "csv"
                export_clean(home, dataset, start, end, file_format, use_config_period=use_config)
            elif choice == "5":
                source = Path(input("Paste the Bonus Matrix v1.2 workbook path: ").strip().strip('"'))
                import_bonus_tool(home, source)
            elif choice == "6":
                path = Path(input("Paste the edited Attendance Corrections workbook path: ").strip().strip('"'))
                import_decisions(home, path)
            elif choice == "7":
                path = Path(input("Paste the edited PCS Performance workbook path: ").strip().strip('"'))
                import_coaching_decisions(home, path)
            elif choice == "8":
                rules_tool(home, "catalog")
            elif choice == "9":
                show_status(home)
                show_coverage(home)
            elif choice == "10":
                create_backup(home)
            elif choice == "11":
                path = Path(input("Paste the folder containing FTE, Storm and Verint: ").strip().strip('"'))
                setup(home, path, True)
            elif choice == "12":
                run_doctor(home)
            elif choice == "13":
                config = load_config(home)
                kind, job = _choose_custom_job(config)
                start, end, use_config = _choose_period()
                run_custom(home, kind, job, start, end, use_config)
            elif choice == "14":
                pack = _choose_template_pack()
                start, end, use_config = _choose_period()
                profile = _choose_service_profile(home) if pack == "service" else None
                initialize_excel_template(home, pack, start, end, profile, use_config_period=use_config)
            elif choice == "15":
                return 0
            else:
                print("Please choose a number from 1 to 15.")
        except Exception as exc:
            print(f"\nERROR: {exc}")
            print("Nothing was changed in your extract files. Check the latest file in logs.")
        _pause_for_dashboard()


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
    refresh_p.add_argument("--pack", action="append", choices=IMPLEMENTED_REPORT_PACK_KEYS)
    refresh_p.add_argument("--all-packs", action="store_true")
    refresh_p.add_argument("--source-group", choices=tuple(SOURCE_GROUPS), default="all")
    refresh_p.add_argument("--service-profile", help="Effective service profile id for the service report")
    report_p = commands.add_parser("report", help="Build an Excel report from current marts")
    report_p.add_argument("--start", type=_date)
    report_p.add_argument("--end", type=_date)
    report_p.add_argument("--output", type=Path)
    report_p.add_argument("--pack", choices=IMPLEMENTED_REPORT_PACK_KEYS, default="pcs")
    report_p.add_argument("--service-profile", help="Effective service profile id for the service report")
    export_p = commands.add_parser("export", help="Export a cleaned hub dataset")
    export_p.add_argument("dataset", choices=tuple(DATASETS))
    export_p.add_argument("--start", type=_date)
    export_p.add_argument("--end", type=_date)
    export_p.add_argument("--format", choices=("csv", "xlsx"), default="csv")
    export_p.add_argument("--output", type=Path)
    custom_p = commands.add_parser("custom", help="Run a trusted Python or read-only SQL job")
    custom_p.add_argument("kind", choices=("python", "sql"))
    custom_p.add_argument("job", type=Path)
    custom_p.add_argument("--start", type=_date)
    custom_p.add_argument("--end", type=_date)
    import_p = commands.add_parser("import-actions", help="Import edited GAPS decision columns")
    import_p.add_argument("workbook", type=Path)
    coaching_p = commands.add_parser("import-pcs-actions", help="Import edited PCS coaching columns")
    coaching_p.add_argument("workbook", type=Path)
    bonus_p = commands.add_parser("import-bonus", help="Import Bonus Matrix v1.2 without changing the source")
    bonus_p.add_argument("workbook", type=Path)
    analysis_p = commands.add_parser("analyze", help="Run deterministic on-demand period analysis")
    analysis_p.add_argument("domain", choices=ANALYSIS_DOMAINS)
    analysis_p.add_argument("--start", type=_date)
    analysis_p.add_argument("--end", type=_date)
    analysis_p.add_argument("--comparison", choices=COMPARISON_MODES, default="previous_equal")
    analysis_p.add_argument("--output", type=Path)
    template_p = commands.add_parser("template-init", help="Create a protected Excel Pivot/slicer starter")
    template_p.add_argument("--pack", choices=IMPLEMENTED_REPORT_PACK_KEYS, default="pcs")
    template_p.add_argument("--start", type=_date)
    template_p.add_argument("--end", type=_date)
    template_p.add_argument("--service-profile", help="Effective service profile id for the service report")
    template_p.add_argument("--force", action="store_true", help="Replace an existing master intentionally")
    commands.add_parser("status", help="Show source health")
    commands.add_parser("coverage", help="Show available dates and row counts")
    commands.add_parser("backup", help="Create a database backup")
    commands.add_parser("doctor", help="Test the corporate runtime, SQLite and Excel libraries")
    rules_p = commands.add_parser("rules", help="Validate, explain, test or compare governed methods")
    rules_p.add_argument(
        "action", choices=("validate", "catalog", "explain", "diff", "test"),
        nargs="?", default="validate",
    )
    rules_p.add_argument("metric", nargs="?", help="Metric id for the explain action")
    rules_p.add_argument("--against", type=Path, help="Earlier metric catalog for the diff action")
    commands.add_parser("menu", help="Open the interactive menu")
    return root


def main(argv: list[str] | None = None) -> int:
    args = parser().parse_args(argv)
    home = _home(args.home)
    try:
        if args.command == "setup":
            return setup(home, args.source_root, args.non_interactive)
        if args.command == "refresh":
            packs = () if args.no_report else IMPLEMENTED_REPORT_PACK_KEYS if args.all_packs else tuple(args.pack or ["attendance"])
            return refresh(home, args.start, args.end, packs, args.source_group, service_profile=args.service_profile)
        if args.command == "report":
            return report_only(home, args.start, args.end, args.output, args.pack, service_profile=args.service_profile)
        if args.command == "import-actions":
            return import_decisions(home, args.workbook)
        if args.command == "import-pcs-actions":
            return import_coaching_decisions(home, args.workbook)
        if args.command == "import-bonus":
            return import_bonus_tool(home, args.workbook)
        if args.command == "analyze":
            return analyze_period(home, args.domain, args.start, args.end, args.comparison, args.output)
        if args.command == "template-init":
            return initialize_excel_template(
                home, args.pack, args.start, args.end, args.service_profile, args.force,
            )
        if args.command == "export":
            return export_clean(home, args.dataset, args.start, args.end, args.format, args.output)
        if args.command == "custom":
            return run_custom(home, args.kind, args.job, args.start, args.end)
        if args.command == "status":
            return show_status(home)
        if args.command == "coverage":
            return show_coverage(home)
        if args.command == "backup":
            return create_backup(home)
        if args.command == "doctor":
            return 0 if run_doctor(home) else 1
        if args.command == "rules":
            return rules_tool(home, args.action, args.metric, args.against)
        return menu(home)
    except (ConfigError, HubLockedError, FileNotFoundError, FileExistsError, ValueError, RuntimeError) as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 1
    except KeyboardInterrupt:
        print("\nCancelled.")
        return 130
