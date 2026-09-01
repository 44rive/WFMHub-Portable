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
    ),
    "absence": ReportPack(
        key="absence",
        default_folder="absence",
        filename_prefix="WFMHub_Final_Absenteeism",
        purpose="Activities-only corrected final absenteeism ledger.",
    ),
    "corrections": ReportPack(
        key="corrections",
        default_folder="corrections",
        filename_prefix="WFMHub_Yesterday_Corrections",
        purpose="Residual observed gaps, decisions, and a full-shift evidence timeline.",
    ),
    "scorecard": ReportPack(
        key="scorecard",
        default_folder="scorecard",
        filename_prefix="WFMHub_Executive_Scorecard",
        purpose="Legacy combined scorecard retained only for command compatibility.",
        implemented=False,
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
        from .governed_workbooks import build_daily_operations_workbook

        return build_daily_operations_workbook(conn, config, start, end, output)
    if key == "quality_pcs":
        from .governed_workbooks import build_exact_pcs_workbook

        return build_exact_pcs_workbook(conn, config, start, end, output)
    if key == "absence":
        from .governed_workbooks import build_final_absence_workbook

        return build_final_absence_workbook(conn, config, start, end, output)
    if key == "corrections":
        from .governed_workbooks import build_corrections_workbook

        return build_corrections_workbook(conn, config, start, end, output)
    raise ValueError(f"Report pack {key!r} has no registered builder")
