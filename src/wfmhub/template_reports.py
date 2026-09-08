"""Shared visual shell for WFM decision workbooks."""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Any, Sequence

from .config import Config
from .design import REPORT_DESIGN_ID, REPORT_DESIGN_VERSION
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
        self.report.workbook.set_custom_property(
            "WFMHub Report Design", f"{REPORT_DESIGN_ID} {REPORT_DESIGN_VERSION}",
        )
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
        self.card_text = add({
            "font_name": "Aptos Display", "font_size": 18, "bold": True,
            "font_color": COLORS["dark"], "bg_color": COLORS["white"],
            "align": "center", "valign": "vcenter",
            "border": 1, "border_color": COLORS["thin"],
        })
        self.scope_label = add({
            "font_name": "Aptos", "font_size": 8, "bold": True,
            "font_color": COLORS["muted"], "bg_color": COLORS["canvas"],
            "align": "left", "valign": "vcenter", "indent": 1,
            "top": 1, "bottom": 1, "left": 1,
            "top_color": COLORS["thin"], "bottom_color": COLORS["thin"],
            "left_color": COLORS["thin"],
        })
        self.scope_value = add({
            "font_name": "Aptos", "font_size": 10, "bold": True,
            "font_color": COLORS["dark"], "bg_color": COLORS["white"],
            "align": "left", "valign": "vcenter", "indent": 1,
            "top": 1, "bottom": 1, "right": 1,
            "top_color": COLORS["thin"], "bottom_color": COLORS["thin"],
            "right_color": COLORS["thin"],
        })
        self.badge_formats = {
            "FINAL": add({"font_name": "Aptos", "font_size": 10, "bold": True, "font_color": COLORS["white"], "bg_color": COLORS["green"], "align": "center", "valign": "vcenter"}),
            "LIVE": add({"font_name": "Aptos", "font_size": 10, "bold": True, "font_color": COLORS["white"], "bg_color": COLORS["teal"], "align": "center", "valign": "vcenter"}),
            "PROVISIONAL": add({"font_name": "Aptos", "font_size": 10, "bold": True, "font_color": COLORS["dark"], "bg_color": COLORS["amber_light"], "align": "center", "valign": "vcenter"}),
            "INCOMPLETE": add({"font_name": "Aptos", "font_size": 10, "bold": True, "font_color": COLORS["white"], "bg_color": COLORS["red"], "align": "center", "valign": "vcenter"}),
        }

    def compact_header(
        self,
        ws,
        title: str,
        subtitle: str,
        *,
        last_col: int,
        status: str = "LIVE",
        status_label: str = "UPDATED",
    ) -> None:
        """Write the approved compact title, freshness badge, and context band."""

        badge = status if status in self.badge_formats else "INCOMPLETE"
        badge_start = max(1, last_col - 2)
        if badge_start > 0:
            ws.merge_range(0, 0, 0, badge_start - 1, title, self.report.title)
        else:
            ws.write(0, 0, title, self.report.title)
        ws.merge_range(0, badge_start, 0, last_col, status_label, self.badge_formats[badge])
        ws.merge_range(1, 0, 1, last_col, subtitle, self.report.subtitle)
        ws.set_row(0, 34)
        ws.set_row(1, 21)

    def scope_strip(
        self,
        ws,
        entries: Sequence[tuple[str, Any]],
        *,
        row: int = 3,
        last_col: int = 13,
    ) -> None:
        """Show compact, honest report scope fields below the header."""

        if not entries:
            return
        columns = last_col + 1
        base = columns // len(entries)
        remainder = columns % len(entries)
        start = 0
        for index, (label, value) in enumerate(entries):
            width = base + (1 if index < remainder else 0)
            end = min(last_col, start + width - 1)
            ws.write(row, start, label.upper(), self.scope_label)
            if end > start + 1:
                ws.merge_range(row, start + 1, row, end, value, self.scope_value)
            elif end == start + 1:
                ws.write(row, end, value, self.scope_value)
            start = end + 1
        ws.set_row(row, 24)

    def four_card_row(
        self,
        ws,
        cards: Sequence[KpiCard],
        *,
        row: int = 5,
        starts: Sequence[int] = (0, 4, 8, 12),
        width: int = 3,
    ) -> None:
        """Render the first four decision KPIs with one consistent geometry."""

        for start, card in zip(starts, cards[:4]):
            end = start + width - 1
            ws.merge_range(row, start, row, end, card.label.upper(), self.report.kpi_label)
            fmt = {
                "percent": self.card_percent,
                "decimal": self.card_decimal,
                "money": self.card_money,
                "text": self.card_text,
            }.get(card.kind, self.card_integer)
            ws.merge_range(row + 1, start, row + 2, end, card.value, fmt)
            ws.merge_range(
                row + 3, start, row + 3, end,
                card.comparison or " ", self.card_compare,
            )
        ws.set_row(row + 1, 25)
        ws.set_row(row + 2, 25)

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
        sheet_name: str = "DASHBOARD",
    ) -> None:
        ws = self.report.workbook.add_worksheet(sheet_name)
        ws.hide_gridlines(2)
        ws.set_tab_color(COLORS["gold"])
        ws.set_zoom(90)
        ws.merge_range("A1:N1", self.title_text, self.report.title)
        ws.merge_range(
            "A2:N2",
            f"Period {self.start:%Y-%m-%d} to {self.end:%Y-%m-%d}  |  data updated {self.generated:%Y-%m-%d %H:%M}  |  prepared by Anass ASSRI",
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
                    "categories": [sheet_name, table_row + 3, 0, table_row + 2 + len(comparison_rows), 0],
                    "values": [sheet_name, table_row + 3, column, table_row + 2 + len(comparison_rows), column],
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
            "Business definitions",
            "Approved measures and how to use them.",
            ["Metric", "Definition", "Operational use", "Important note"],
            rows,
        )

    def audit(self, rows: Sequence[Sequence[Any]]) -> None:
        audit_rows = list(rows)
        audit_rows.append((
            "Report design",
            f"{REPORT_DESIGN_ID} {REPORT_DESIGN_VERSION}",
            "Canonical tokens and layout are documented in docs/REPORT_DESIGN_SYSTEM.md",
        ))
        ws = self.table(
            "_AUDIT",
            "Report controls",
            "Technical details for reconciliation. This sheet is hidden by default.",
            ["Item", "Value", "Evidence"],
            audit_rows,
        )
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
            ws.set_footer("&LPrepared by Anass ASSRI | WFM&CPage &P of &N&RConfidential")
        self.report.close()
        return self.path
