"""Stable report-pack registry for independently growing business outputs."""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from .config import Config
    from .database import DatabaseConnection


@dataclass(frozen=True)
class ReportPack:
    key: str
    default_folder: str
    filename_prefix: str
    purpose: str
    implemented: bool = True


REPORT_PACKS = {
    "operations": ReportPack(
        key="operations",
        default_folder="operations",
        filename_prefix="WFMHub_Daily_Operations",
        purpose="Attendance calls, staffing gaps, and APDE service state without adherence KPIs.",
        implemented=True,
    ),
    "intraday": ReportPack(
        key="intraday",
        default_folder="intraday",
        filename_prefix="WFMHub_Intraday",
        purpose="Legacy combined intraday output retained only for command compatibility.",
        implemented=False,
    ),
    "quality_pcs": ReportPack(
        key="quality_pcs",
        default_folder="quality_pcs",
        filename_prefix="WFMHub_Quality_PCS",
        purpose="Agent call performance and PCS metrics sourced from call-by-call data.",
        implemented=True,
    ),
    "absence": ReportPack(
        key="absence",
        default_folder="absence",
        filename_prefix="WFMHub_Final_Absence_Shrinkage",
        purpose="Corrected Verint final absence and shrinkage ledger.",
    ),
    "corrections": ReportPack(
        key="corrections",
        default_folder="corrections",
        filename_prefix="WFMHub_Yesterday_Corrections",
        purpose="Residual observed gaps, decisions, and a full-shift evidence timeline.",
    ),
    "pcs": ReportPack(
        key="pcs",
        default_folder="pcs",
        filename_prefix="WFMHub_PCS_Performance",
        purpose="Daily, MTD and prior-month PCS performance and participation.",
    ),
    "bonus": ReportPack(
        key="bonus",
        default_folder="bonus",
        filename_prefix="WFMHub_Bonus_Performance",
        purpose="Governed monthly bonus calculation, diagnostics and release gates.",
    ),
    "service": ReportPack(
        key="service",
        default_folder="service",
        filename_prefix="WFMHub_Service_Performance",
        purpose="Mapped-LOB service level, availability, forecast deviation and AHT.",
    ),
    "staffing": ReportPack(
        key="staffing",
        default_folder="staffing",
        filename_prefix="WFMHub_Staffing_Coverage",
        purpose="LOB/language staffing coverage and actionable interval gaps.",
    ),
    "attendance": ReportPack(
        key="attendance",
        default_folder="attendance",
        filename_prefix="WFMHub_Attendance_Callouts",
        purpose="Selected-period attendance callout queue without adherence.",
    ),
    "scorecard": ReportPack(
        key="scorecard",
        default_folder="scorecard",
        filename_prefix="WFMHub_Executive_Scorecard",
        purpose="Legacy combined scorecard retained only for command compatibility.",
        implemented=False,
    ),
}

IMPLEMENTED_REPORT_PACK_KEYS = (
    "pcs", "bonus", "service", "staffing", "attendance", "corrections", "absence",
)


def report_pack(key: str) -> ReportPack:
    try:
        return REPORT_PACKS[key]
    except KeyError as exc:
        raise ValueError(f"Unknown report pack {key!r}. Available: {', '.join(REPORT_PACKS)}") from exc


def report_pack_folder(config: Config, key: str) -> Path:
    pack = report_pack(key)
    configured = config.report_packs.get(key, pack.default_folder)
    folder = Path(configured)
    return folder.resolve() if folder.is_absolute() else (config.output / folder).resolve()


def build_report_pack(
    key: str,
    conn: DatabaseConnection,
    config: Config,
    start: date,
    end: date,
    output: Path | None = None,
    service_profile: str | None = None,
) -> Path:
    pack = report_pack(key)
    if not pack.implemented:
        raise ValueError(
            f"Report pack {key!r} is reserved but not implemented yet. {pack.purpose}"
        )
    if key == "operations":
        from .governed_workbooks import build_daily_operations_workbook

        return build_daily_operations_workbook(conn, config, start, end, output)
    if key == "quality_pcs":
        from .governed_workbooks import build_exact_pcs_workbook

        return build_exact_pcs_workbook(conn, config, start, end, output)
    if key == "pcs":
        from .decision_products import build_pcs_performance_workbook

        return build_pcs_performance_workbook(conn, config, start, end, output)
    if key == "bonus":
        from .bonus import build_bonus_performance_workbook

        return build_bonus_performance_workbook(conn, config, start, end, output)
    if key == "service":
        from .decision_products import build_service_performance_workbook

        return build_service_performance_workbook(conn, config, start, end, output, service_profile)
    if key == "staffing":
        from .decision_products import build_staffing_coverage_workbook

        return build_staffing_coverage_workbook(conn, config, start, end, output)
    if key == "attendance":
        from .decision_products import build_attendance_today_workbook

        return build_attendance_today_workbook(conn, config, start, end, output)
    if key == "absence":
        from .decision_products import build_final_absence_product_workbook

        return build_final_absence_product_workbook(conn, config, start, end, output)
    if key == "corrections":
        from .decision_products import build_attendance_corrections_workbook

        return build_attendance_corrections_workbook(conn, config, start, end, output)
    raise ValueError(f"Report pack {key!r} has no registered builder")
