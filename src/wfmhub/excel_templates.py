"""Safe lifecycle for Excel-authored PivotTable and slicer masters.

WFMHub can create a styled starter, but deliberately never edits it again.
That boundary preserves Excel-only PivotTable, Data Model and slicer parts that
Python workbook libraries cannot safely round-trip.
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from .config import Config


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
        if self.report_key.casefold() == "pcs":
            return self.model_folder.parents[1] / "template_feeds" / "pcs" / "current"
        return self.model_folder


def excel_template(config: Config, report_key: str) -> ExcelTemplate:
    """Return the stable master and model-data locations for one report."""

    safe_key = "_".join(
        part for part in "".join(
            char.lower() if char.isalnum() else "_" for char in report_key
        ).split("_") if part
    )
    return ExcelTemplate(
        report_key=report_key,
        path=(config.home / "templates" / "reports" / f"{safe_key}.xlsx").resolve(),
        model_folder=(config.output / "model_data" / safe_key).resolve(),
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
