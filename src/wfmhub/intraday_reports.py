"""Build the standalone Intraday actual/forecast workbook."""

from __future__ import annotations

from datetime import date, datetime
from pathlib import Path

from .config import Config
from .database import DatabaseConnection
from .report_packs import report_pack, report_pack_folder
from .reports import ExcelReport, _query
from .rules import evaluate_formula, load_rulebook


def _start(report: ExcelReport, config: Config, start: date, end: date, generated: datetime) -> None:
    ws = report.workbook.add_worksheet("START_HERE")
    ws.hide_gridlines(2)
    ws.merge_range("A1:H1", "WFMHub Intraday", report.title)
    ws.merge_range(
        "A2:H2",
        f"Generated {generated:%Y-%m-%d %H:%M}; period {start:%Y-%m-%d} to {end:%Y-%m-%d}",
        report.subtitle,
    )
    lines = [
        "Actuals come only from Storm APBE/APFR/APDE. Forecast and required staffing come only from Verint Forecast.",
        "Availability means service availability (answered / offered), never agent availability or adherence.",
        "Actual and forecast rows remain separate until a reviewed queue/LOB mapping exists.",
        "Use SOURCE_HEALTH first; do not trust a KPI when its source is ERROR or MISSING.",
        f"Source root: {config.source_root}",
    ]
    for row, line in enumerate(lines, 4):
        ws.write(row - 1, 0, line, report.body)
    ws.set_column("A:A", 115)
    ws.set_column("B:H", 3)


def _summary(report: ExcelReport, conn: DatabaseConnection, config: Config, start: date, end: date) -> None:
    ws = report.workbook.add_worksheet("SUMMARY")
    ws.hide_gridlines(2)
    ws.merge_range("A1:F1", "Intraday summary", report.title)
    ws.merge_range("A2:F2", "Actual and forecast KPIs are shown independently.", report.subtitle)
    rulebook = load_rulebook(config.home, config.business_rules)
    offered, answered, abandoned, short_abandoned, answered_target, handled_seconds = conn.execute(
        """SELECT coalesce(sum(offered),0), coalesce(sum(answered),0),
                  coalesce(sum(abandoned),0), coalesce(sum(short_abandoned),0),
                  coalesce(sum(answered_within_target),0), coalesce(sum(handled_seconds),0)
           FROM mart.service_interval WHERE business_date BETWEEN ? AND ?""", [start, end]
    ).fetchone()
    forecast_volume, forecast_fte, required_fte = conn.execute(
        """SELECT coalesce(sum(volume_forecast),0), coalesce(sum(fte_forecast),0),
                  coalesce(sum(fte_required),0)
           FROM mart.forecast_hour WHERE business_date BETWEEN ? AND ?""", [start, end]
    ).fetchone()
    values = {
        "offered": offered, "answered": answered, "abandoned": abandoned,
        "short_abandoned": short_abandoned, "answered_within_target": answered_target,
        "handled_seconds": handled_seconds,
    }
    metrics = [
        ("Actual offered", offered), ("Actual answered", answered), ("Actual abandoned", abandoned),
        ("Gross SL 20s", evaluate_formula(rulebook.service_profiles["gross_20"].formula, values)),
        ("Adjusted SL 20s", evaluate_formula(rulebook.service_profiles["adjusted_20"].formula, values)),
        ("Service availability", evaluate_formula(rulebook.formulas["service_availability"].formula, values)),
        ("Actual abandon rate", evaluate_formula(rulebook.formulas["abandon_rate"].formula, values)),
        ("Weighted actual AHT", evaluate_formula(rulebook.formulas["aht_seconds"].formula, values)),
        ("Forecast volume", forecast_volume), ("Forecast FTE", forecast_fte), ("Required FTE", required_fte),
    ]
    ws.write("A4", "KPI", report.header)
    ws.write("B4", "Value", report.header)
    for index, (label, value) in enumerate(metrics, 4):
        ws.write(index, 0, label, report.body)
        fmt = report.percent if label in {"Gross SL 20s", "Adjusted SL 20s", "Service availability", "Actual abandon rate"} else report.integer
        ws.write(index, 1, value, fmt)
    ws.set_column("A:A", 28)
    ws.set_column("B:B", 18)


def build_intraday_report(
    conn: DatabaseConnection,
    config: Config,
    start: date,
    end: date,
    output: Path | None = None,
) -> Path:
    generated = datetime.now()
    pack = report_pack("intraday")
    output = (
        output
        or report_pack_folder(config, pack.key)
        / f"{pack.filename_prefix}_{start:%Y-%m-%d}_to_{end:%Y-%m-%d}_{generated:%H%M%S_%f}.xlsx"
    ).resolve()
    output.parent.mkdir(parents=True, exist_ok=True)
    partial = output.with_name(f"{output.stem}.partial{output.suffix}")
    report = ExcelReport(partial)
    try:
        _start(report, config, start, end, generated)
        _summary(report, conn, config, start, end)
        headers, rows = _query(
            conn,
            """SELECT business_date, interval_start, source_system, queue,
                      business_partner, lob, language, offered, answered,
                      abandoned, short_abandoned, answered_within_target,
                      sl_gross, sl_adjusted, sl_profile, service_level,
                      service_availability, abandon_rate, aht_seconds,
                      rule_version, source_file
               FROM mart.service_interval
               WHERE business_date BETWEEN ? AND ?
               ORDER BY business_date, interval_start, source_system, queue
               LIMIT ?""",
            [start, end, config.report_limits.get("max_intraday_rows", 100000)],
        )
        report.add_table_sheet(
            "ACTUALS", "Intraday actuals",
            "Storm APBE/APFR/APDE actual queue intervals. Availability is answered/offered; no Verint actual fields are used.",
            headers, rows,
        )
        headers, rows = _query(
            conn,
            """SELECT business_date, hour_start, queue_name, volume_forecast,
                      fte_forecast, fte_required, sl_forecast, sl_required,
                      aht_forecast_seconds, source_file
               FROM mart.forecast_hour
               WHERE business_date BETWEEN ? AND ?
               ORDER BY business_date, hour_start, queue_name
               LIMIT ?""",
            [start, end, config.report_limits.get("max_intraday_rows", 100000)],
        )
        report.add_table_sheet(
            "FORECAST", "Verint forecast and requirements",
            "Forecast-only feed; keep separate from actuals until scope mapping is approved.",
            headers, rows,
        )
        headers, rows = _query(
            conn,
            """SELECT detected_at, source_family, source_file, business_date,
                      issue_type, severity, details
               FROM meta.quality_issue
               WHERE (business_date IS NULL OR business_date BETWEEN ? AND ?)
                 AND source_family IN ('intraday','service','forecast','apbe','apfr','apde')
               ORDER BY CASE severity WHEN 'ERROR' THEN 1 ELSE 2 END, business_date""",
            [start, end],
        )
        report.add_table_sheet("DATA_QUALITY", "Intraday data quality", "Resolve ERROR before using KPIs.", headers, rows, exception_column="Severity")
        headers, rows = _query(
            conn,
            """SELECT source_family, expected_path, newest_file, newest_business_date,
                      modified_at, loaded_at, row_count, rejected_count,
                      scoped_out_count, status, details
               FROM mart.source_health
               WHERE source_family IN ('forecast','apbe','apfr','apde')
               ORDER BY source_family""",
        )
        report.add_table_sheet("SOURCE_HEALTH", "Intraday source health", "Files and current coverage for this pack.", headers, rows, exception_column="Status")
    except Exception:
        report.close()
        partial.unlink(missing_ok=True)
        raise
    else:
        report.close()
        partial.replace(output)
    return output
