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
        filename_prefix="WFMHub_Operations",
        purpose="Attendance, corrections, data quality, and source health without adherence KPIs.",
    ),
    "intraday": ReportPack(
        key="intraday",
        default_folder="intraday",
        filename_prefix="WFMHub_Intraday",
        purpose="Storm actual performance and separate Verint forecast/requirements.",
    ),
    "quality_pcs": ReportPack(
        key="quality_pcs",
        default_folder="quality_pcs",
        filename_prefix="WFMHub_Quality_PCS",
        purpose="Agent call performance and PCS metrics sourced from call-by-call data.",
    ),
    "absence": ReportPack(
        key="absence",
        default_folder="absence",
        filename_prefix="WFMHub_Attendance_Absence",
        purpose="Payroll absence, vacation, shrinkage, spells and classified Verint events.",
    ),
    "scorecard": ReportPack(
        key="scorecard",
        default_folder="scorecard",
        filename_prefix="WFMHub_Executive_Scorecard",
        purpose="Rule-versioned service, forecast, absence and PCS KPI facts.",
    ),
}

IMPLEMENTED_REPORT_PACK_KEYS = tuple(
    key for key, pack in REPORT_PACKS.items() if pack.implemented
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
) -> Path:
    pack = report_pack(key)
    if not pack.implemented:
        raise ValueError(
            f"Report pack {key!r} is reserved but not implemented yet. {pack.purpose}"
        )
    if key == "operations":
        # Local import keeps the registry independent from individual builders
        # while preserving the existing reports.build_report() API.
        from .reports import build_report

        return build_report(conn, config, start, end, output)
    if key == "intraday":
        from .intraday_reports import build_intraday_report

        return build_intraday_report(conn, config, start, end, output)
    if key == "quality_pcs":
        from .pcs_reports import build_pcs_report

        return build_pcs_report(conn, config, start, end, output)
    if key == "absence":
        from .sota_reports import build_absence_report

        return build_absence_report(conn, config, start, end, output)
    if key == "scorecard":
        from .sota_reports import build_scorecard_report

        return build_scorecard_report(conn, config, start, end, output)
    raise ValueError(f"Report pack {key!r} has no registered builder")
