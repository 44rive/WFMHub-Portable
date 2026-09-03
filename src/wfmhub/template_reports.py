"""Shared decision-workbook shell and compact Excel Data Model exports.

The generated XLSX is immediately usable as a static management snapshot.  In
parallel, every visible dataset is written to a small CSV package.  A reusable
Excel-authored template can connect those CSV files directly to the Data Model
and retain native PivotTables and slicers without Python rewriting the XLSX.
"""

from __future__ import annotations

import csv
import hashlib
import json
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Any, Iterable, Sequence

from .config import Config
from .reports import COLORS, ExcelReport


@dataclass(frozen=True)
class ModelTable:
    key: str
    headers: Sequence[str]
    rows: Sequence[Sequence[Any]]


@dataclass(frozen=True)
class KpiCard:
    label: str
    value: Any
    kind: str = "integer"
    comparison: str | None = None


def _csv_value(value: Any) -> Any:
    if isinstance(value, datetime):
        return value.isoformat(sep=" ", timespec="seconds")
    if isinstance(value, date):
        return value.isoformat()
    if isinstance(value, bool):
        return 1 if value else 0
    return "" if value is None else value


def _safe_key(value: str) -> str:
    cleaned = "".join(char.lower() if char.isalnum() else "_" for char in value)
    return "_".join(part for part in cleaned.split("_") if part)


def write_model_package(
    config: Config,
    report_key: str,
    start: date,
    end: date,
    generated: datetime,
    tables: Iterable[ModelTable],
) -> Path:
    """Atomically replace the compact files consumed by Excel templates."""

    root = config.output / "model_data" / _safe_key(report_key)
    root.mkdir(parents=True, exist_ok=True)
    table_list = list(tables)
    expected = {f"{_safe_key(table.key)}.csv" for table in table_list}
    for existing in root.glob("*.csv"):
        if existing.name not in expected and existing.name != "context.csv":
            existing.unlink()

    files: list[dict[str, Any]] = []
    for table in table_list:
        target = root / f"{_safe_key(table.key)}.csv"
        partial = target.with_suffix(".csv.partial")
        with partial.open("w", encoding="utf-8-sig", newline="") as handle:
            writer = csv.writer(handle, lineterminator="\n")
            writer.writerow(table.headers)
            writer.writerows([_csv_value(value) for value in row] for row in table.rows)
        partial.replace(target)
        files.append({
            "table": table.key,
            "file": target.name,
            "rows": len(table.rows),
            "sha256": hashlib.sha256(target.read_bytes()).hexdigest(),
        })

    context = root / "context.csv"
    context_partial = context.with_suffix(".csv.partial")
    with context_partial.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.writer(handle, lineterminator="\n")
        writer.writerow(["report_key", "period_start", "period_end", "generated_at", "model_folder"])
        writer.writerow([report_key, start.isoformat(), end.isoformat(), generated.isoformat(timespec="seconds"), str(root)])
    context_partial.replace(context)
    files.append({
        "table": "context",
        "file": context.name,
        "rows": 1,
        "sha256": hashlib.sha256(context.read_bytes()).hexdigest(),
    })

    manifest = root / "manifest.json"
    manifest_partial = manifest.with_suffix(".json.partial")
    manifest_partial.write_text(json.dumps({
        "schema_version": 1,
        "report_key": report_key,
        "period_start": start.isoformat(),
        "period_end": end.isoformat(),
        "generated_at": generated.isoformat(timespec="seconds"),
        "calculation_authority": "WFMHub Python + SQLite",
        "excel_load": "Power Query connection-only + Add to Data Model",
        "files": files,
    }, indent=2), encoding="utf-8")
    manifest_partial.replace(manifest)
    return root


class DecisionWorkbook:
    """One visual contract for every WFM/Operations decision product."""

    def __init__(
        self,
        path: Path,
        config: Config,
        report_key: str,
        title: str,
        start: date,
        end: date,
        generated: datetime,
    ):
        self.path = path
        self.config = config
        self.report_key = report_key
        self.title_text = title
        self.start = start
        self.end = end
        self.generated = generated
        self.report = ExcelReport(path)
        self.tables: list[ModelTable] = []
        add = self.report.workbook.add_format
        self.card_percent = add({
            "font_name": "Aptos Display", "font_size": 20, "bold": True,
            "font_color": COLORS["dark"], "bg_color": COLORS["white"],
            "align": "center", "valign": "vcenter", "num_format": "0.0%",
            "border": 1, "border_color": COLORS["thin"],
        })
        self.card_decimal = add({
            "font_name": "Aptos Display", "font_size": 20, "bold": True,
            "font_color": COLORS["dark"], "bg_color": COLORS["white"],
            "align": "center", "valign": "vcenter", "num_format": "#,##0.00",
            "border": 1, "border_color": COLORS["thin"],
        })
        self.card_integer = add({
            "font_name": "Aptos Display", "font_size": 20, "bold": True,
            "font_color": COLORS["dark"], "bg_color": COLORS["white"],
            "align": "center", "valign": "vcenter", "num_format": "#,##0",
            "border": 1, "border_color": COLORS["thin"],
        })
        self.card_money = add({
            "font_name": "Aptos Display", "font_size": 18, "bold": True,
            "font_color": COLORS["dark"], "bg_color": COLORS["white"],
            "align": "center", "valign": "vcenter", "num_format": '#,##0.00 "MAD"',
            "border": 1, "border_color": COLORS["thin"],
        })
        self.card_compare = add({
            "font_name": "Aptos", "font_size": 8, "font_color": COLORS["muted"],
            "bg_color": COLORS["canvas"], "align": "center", "valign": "vcenter",
            "border": 1, "border_color": COLORS["thin"],
        })
        self.badge_formats = {
            "FINAL": add({"font_name": "Aptos", "font_size": 10, "bold": True, "font_color": COLORS["white"], "bg_color": COLORS["green"], "align": "center", "valign": "vcenter"}),
            "LIVE": add({"font_name": "Aptos", "font_size": 10, "bold": True, "font_color": COLORS["white"], "bg_color": COLORS["teal"], "align": "center", "valign": "vcenter"}),
            "PROVISIONAL": add({"font_name": "Aptos", "font_size": 10, "bold": True, "font_color": COLORS["dark"], "bg_color": COLORS["amber_light"], "align": "center", "valign": "vcenter"}),
            "INCOMPLETE": add({"font_name": "Aptos", "font_size": 10, "bold": True, "font_color": COLORS["white"], "bg_color": COLORS["red"], "align": "center", "valign": "vcenter"}),
        }

    def dashboard(
        self,
        cards: Sequence[KpiCard],
        status: str,
        status_text: str,
        comparison_headers: Sequence[str],
        comparison_rows: Sequence[Sequence[Any]],
        notes: Sequence[str],
        chart_series: Sequence[tuple[str, int]] = (),
        chart_type: str = "line",
    ) -> None:
        ws = self.report.workbook.add_worksheet("DASHBOARD")
        ws.hide_gridlines(2)
        ws.set_tab_color(COLORS["gold"])
        ws.set_zoom(90)
        ws.merge_range("A1:N1", self.title_text, self.report.title)
        ws.merge_range(
            "A2:N2",
            f"Period {self.start:%Y-%m-%d} to {self.end:%Y-%m-%d}  |  generated {self.generated:%Y-%m-%d %H:%M}  |  WFM HUB",
            self.report.subtitle,
        )
        ws.set_row(0, 34)
        ws.set_row(1, 20)
        badge = status if status in self.badge_formats else "INCOMPLETE"
        ws.merge_range("A4:N4", f"{badge}  /  {status_text}", self.badge_formats[badge])
        ws.set_row(3, 25)

        for index, card in enumerate(cards[:8]):
            row = 5 if index < 4 else 10
            col = (index % 4) * 3
            ws.merge_range(row, col, row, col + 2, card.label.upper(), self.report.kpi_label)
            fmt = {
                "percent": self.card_percent,
                "decimal": self.card_decimal,
                "money": self.card_money,
            }.get(card.kind, self.card_integer)
            ws.merge_range(row + 1, col, row + 2, col + 2, card.value, fmt)
            ws.merge_range(row + 3, col, row + 3, col + 2, card.comparison or " ", self.card_compare)
            ws.set_row(row + 1, 24)
            ws.set_row(row + 2, 24)

        table_row = 16
        ws.merge_range(table_row, 0, table_row, max(6, len(comparison_headers) - 1), "PERIOD COMPARISON", self.report.section)
        for col, header in enumerate(comparison_headers):
            ws.write(table_row + 2, col, header, self.report.header)
        for offset, row_values in enumerate(comparison_rows):
            unit = None
            if "Unit" in comparison_headers:
                unit_index = comparison_headers.index("Unit")
                if unit_index < len(row_values):
                    unit = str(row_values[unit_index]).casefold()
            for col, value in enumerate(row_values):
                fmt = self.report.body
                if isinstance(value, (int, float)) and not isinstance(value, bool):
                    header = comparison_headers[col].casefold()
                    if unit == "percent" and header in {"current value", "reference value", "delta", "target"}:
                        fmt = self.report.percent
                    elif unit == "money" and header in {"current value", "reference value", "delta", "target"}:
                        fmt = self.card_money
                    else:
                        fmt = self.report.percent if any(token in header for token in ("rate", "%", "participation", "availability", "service level")) else self.report.decimal
                ws.write(table_row + 3 + offset, col, value, fmt)

        if comparison_rows and chart_series:
            chart = self.report.workbook.add_chart({"type": chart_type})
            for series_name, column in chart_series:
                series = {
                    "name": series_name,
                    "categories": ["DASHBOARD", table_row + 3, 0, table_row + 2 + len(comparison_rows), 0],
                    "values": ["DASHBOARD", table_row + 3, column, table_row + 2 + len(comparison_rows), column],
                }
                if chart_type == "line":
                    series.update({"line": {"width": 2.25}, "marker": {"type": "circle", "size": 4}})
                else:
                    series.update({"fill": {"color": COLORS["teal"]}, "border": {"none": True}})
                chart.add_series(series)
            chart.set_legend({"position": "bottom"})
            chart.set_chartarea({"border": {"none": True}, "fill": {"color": COLORS["white"]}})
            chart.set_plotarea({"border": {"none": True}, "fill": {"color": COLORS["white"]}})
            chart.set_y_axis({"major_gridlines": {"visible": False}})
            ws.insert_chart(table_row + 2, 8, chart, {"x_scale": 1.08, "y_scale": 0.92})

        note_row = table_row + 4 + max(len(comparison_rows), 8)
        ws.merge_range(note_row, 0, note_row, 13, "OPERATING NOTES", self.report.section)
        for offset, note in enumerate(notes, 1):
            ws.merge_range(note_row + offset, 0, note_row + offset, 13, note, self.report.note)
            ws.set_row(note_row + offset, 25)
        ws.set_column("A:N", 12)
        ws.set_column("A:A", 18)
        ws.freeze_panes(4, 0)
        ws.set_landscape()
        ws.fit_to_pages(1, 1)

    def table(
        self,
        name: str,
        title: str,
        subtitle: str,
        headers: Sequence[str],
        rows: Sequence[Sequence[Any]],
        *,
        editable_headers: set[str] | None = None,
    ):
        normalized_headers = [str(value) for value in headers]
        normalized_rows = [tuple(row) for row in rows]
        self.tables.append(ModelTable(name, normalized_headers, normalized_rows))
        return self.report.add_table_sheet(
            name, title, subtitle, normalized_headers, normalized_rows,
            editable_headers=editable_headers,
        )

    def definitions(self, rows: Sequence[Sequence[Any]]) -> None:
        self.table(
            "DEFINITIONS",
            "Definitions and operational interpretation",
            "Every KPI is calculated by the governed Python/SQLite engine. Excel only presents the result.",
            ["Metric", "Definition", "Operational use", "Guardrail"],
            rows,
        )

    def audit(self, rows: Sequence[Sequence[Any]]) -> None:
        ws = self.table(
            "_AUDIT",
            "Audit and template connection",
            "Hidden technical lineage. Unhide only for reconciliation or template maintenance.",
            ["Item", "Value", "Evidence"],
            rows,
        )
        for index, row in enumerate(rows, 5):
            if row and str(row[0]) == "Template model folder":
                self.report.workbook.define_name("pModelDataPath", f"='_AUDIT'!$B${index}")
                break
        ws.hide()

    def close(self) -> Path:
        from .report_specs import load_report_catalog

        catalog = load_report_catalog(self.config.home, self.config.report_catalog)
        if self.report_key in catalog.packs:
            expected = catalog.pack(self.report_key).sheets
            actual = tuple(worksheet.get_name() for worksheet in self.report.workbook.worksheets())
            if actual != expected:
                raise ValueError(
                    f"Workbook contract mismatch for {self.report_key}: "
                    f"expected {expected}, created {actual}"
                )
        for ws in self.report.workbook.worksheets():
            ws.set_footer("&LWFM Hub | Anass ASSRI&CPage &P of &N&RConfidential")
        self.report.close()
        write_model_package(
            self.config, self.report_key, self.start, self.end, self.generated,
            (table for table in self.tables if table.key != "_AUDIT"),
        )
        return self.path
