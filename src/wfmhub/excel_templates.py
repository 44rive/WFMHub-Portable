"""Safe lifecycle for Excel-authored PivotTable and slicer masters.

WFMHub can create a styled starter, but deliberately never edits it again.
That boundary preserves Excel-only PivotTable, Data Model and slicer parts that
Python workbook libraries cannot safely round-trip.
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from .config import Config


PCS_QUERY_FILES = (
    "PCS_AgentDay.pq",
    "PCS_Calls.pq",
    "PCS_Agents.pq",
    "PCS_Dates.pq",
)
PCS_FEED_PLACEHOLDER = "__PCS_FEED_FOLDER__"


@dataclass(frozen=True)
class ExcelTemplate:
    report_key: str
    path: Path
    model_folder: Path

    @property
    def exists(self) -> bool:
        return self.path.is_file()

    @property
    def feed_folder(self) -> Path:
        return self.model_folder


def excel_template(config: Config, report_key: str) -> ExcelTemplate:
    """Return the stable master and model-data locations for one report."""

    safe_key = "_".join(
        part for part in "".join(
            char.lower() if char.isalnum() else "_" for char in report_key
        ).split("_") if part
    )
    path = (
        config.reports / "PCS Team.xlsx"
        if safe_key == "pcs"
        else config.system / "templates" / f"{safe_key}.xlsx"
    )
    return ExcelTemplate(
        report_key=report_key,
        path=path.resolve(),
        model_folder=(config.system / "feeds" / safe_key / "current").resolve(),
    )


def require_new_template(config: Config, report_key: str, force: bool = False) -> ExcelTemplate:
    """Protect an existing Excel-authored master from accidental replacement."""

    template = excel_template(config, report_key)
    template.path.parent.mkdir(parents=True, exist_ok=True)
    if template.exists and not force:
        raise FileExistsError(
            f"Excel master already exists and was not changed: {template.path}. "
            "Use --force only if you intentionally want to replace its PivotTables and slicers."
        )
    return template


def materialize_pcs_power_queries(config: Config) -> tuple[Path, ...]:
    """Write firewall-safe PCS query scripts for this exact installation.

    A literal file path keeps each query at one data-source boundary. This
    avoids combining Excel.CurrentWorkbook with File.Contents, which can trip
    Power Query's privacy firewall on managed workstations.
    """

    source_dir = config.home / "templates" / "power_query"
    target_dir = config.system / "power_query"
    target_dir.mkdir(parents=True, exist_ok=True)
    feed_folder = str(config.system / "feeds" / "pcs" / "current").replace('"', '""')
    generated: list[Path] = []
    for filename in PCS_QUERY_FILES:
        source = source_dir / filename
        if not source.is_file():
            raise FileNotFoundError(f"PCS Power Query template is missing: {source}")
        text = source.read_text(encoding="utf-8")
        if PCS_FEED_PLACEHOLDER not in text:
            raise ValueError(f"PCS Power Query template has no feed placeholder: {source}")
        rendered = text.replace(PCS_FEED_PLACEHOLDER, feed_folder)
        target = target_dir / filename
        partial = target.with_suffix(target.suffix + ".partial")
        partial.write_text(rendered, encoding="utf-8")
        partial.replace(target)
        generated.append(target)
    return tuple(generated)
