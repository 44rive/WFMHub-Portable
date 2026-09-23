"""Stable report-pack registry for independently growing business outputs."""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
import shutil
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from .config import Config
    from .database import DatabaseConnection


@dataclass(frozen=True)
class ReportPack:
    key: str
    current_filename: str
    purpose: str


REPORT_PACKS = {
    "absence": ReportPack(
        key="absence",
        current_filename="Final Absenteeism & Shrinkage.xlsx",
        purpose="Final Verint Activities absence and shrinkage totals with exact agent-day components.",
    ),
    "corrections": ReportPack(
        key="corrections",
        current_filename="Attendance Review.xlsx",
        purpose="Read-only exact residual gaps and full-shift evidence after final Verint reconciliation.",
    ),
    "pcs": ReportPack(
        key="pcs",
        current_filename="PCS Live Tracker.xlsx",
        purpose="Permanent direct-CSV Power Query PCS tracker with one combined coaching workspace.",
    ),
    "bonus": ReportPack(
        key="bonus",
        current_filename="Bonus Management.xlsx",
        purpose="Bonus Matrix v1.2 calculation, team analysis and management dashboard.",
    ),
    "service": ReportPack(
        key="service",
        current_filename="RTM Daily Control.xlsx",
        purpose="Daily service, attendance call actions and queue drivers by operational LOB.",
    ),
    "realisations": ReportPack(
        key="realisations",
        current_filename="Realisations.xlsx",
        purpose="Actual versus forecast, service, staffing and final capacity results by mapped LOB.",
    ),
    "staffing": ReportPack(
        key="staffing",
        current_filename="Staffing Preparation.xlsx",
        purpose="15-minute required, gross, PTO/Away, net and gap capacity by governed Planning Group and Staff Type.",
    ),
}

REPORT_PACK_KEYS = tuple(REPORT_PACKS)


class ReportPublishError(RuntimeError):
    """Raised when a finished workbook cannot be published safely."""


def report_pack(key: str) -> ReportPack:
    try:
        return REPORT_PACKS[key]
    except KeyError as exc:
        raise ValueError(f"Unknown report pack {key!r}. Available: {', '.join(REPORT_PACKS)}") from exc


def report_pack_folder(config: Config, key: str) -> Path:
    report_pack(key)
    return config.reports.resolve()


def report_current_path(config: Config, key: str) -> Path:
    """Return the stable, human-facing filename for one report product."""

    pack = report_pack(key)
    return (report_pack_folder(config, key) / pack.current_filename).resolve()


def publish_report(
    config: Config,
    key: str,
    partial: Path,
    target: Path,
    generated: datetime,
) -> Path:
    """Publish a complete workbook and archive the previous current copy.

    Explicit CLI output paths keep their existing replace semantics. Only the
    standard fixed-name product is copied into Reports/Archive.
    """

    try:
        if key not in REPORT_PACKS:
            partial.replace(target)
            return target
        current = report_current_path(config, key)
        if target.resolve() == current and target.exists():
            archive_dir = config.reports / "Archive" / generated.strftime("%Y-%m-%d")
            archive_dir.mkdir(parents=True, exist_ok=True)
            archived = archive_dir / f"{target.stem}_{generated:%Y%m%d_%H%M%S_%f}{target.suffix}"
            shutil.copy2(target, archived)
        partial.replace(target)
    except PermissionError as exc:
        raise ReportPublishError(
            f"Cannot replace '{target.name}'. It is open in Excel or temporarily "
            "locked by OneDrive. Close the workbook, wait for OneDrive to finish "
            "syncing, then retry; the existing workbook was not deleted."
        ) from exc
    return target


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
    if key == "pcs":
        from .pcs_tracker import build_pcs_live_tracker

        return build_pcs_live_tracker(conn, config, start, end, output)
    if key == "bonus":
        from .bonus import build_bonus_performance_workbook

        return build_bonus_performance_workbook(conn, config, start, end, output)
    if key == "service":
        from .service_flash import build_service_flashes_workbook

        return build_service_flashes_workbook(conn, config, start, end, output, service_profile)
    if key == "realisations":
        from .decision_products import build_realisations_workbook

        return build_realisations_workbook(
            conn, config, start, end, output, service_profile,
        )
    if key == "staffing":
        from .decision_products import build_staffing_coverage_workbook

        return build_staffing_coverage_workbook(conn, config, start, end, output)
    if key == "absence":
        from .decision_products import build_final_absence_product_workbook

        return build_final_absence_product_workbook(conn, config, start, end, output)
    if key == "corrections":
        from .decision_products import build_attendance_corrections_workbook

        return build_attendance_corrections_workbook(conn, config, start, end, output)
    raise ValueError(f"Report pack {key!r} has no registered builder")
