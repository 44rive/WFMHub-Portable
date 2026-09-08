"""Measured Excel geometry for the WFMHub V2 operational report system.

The V1 workbook helpers shared colours but still let each report choose its own
column widths, merged ranges and chart scaling.  This module is deliberately
different: it defines one pixel grid and a small set of components whose
dimensions are testable.  Business calculations must never depend on it.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Sequence

from .design import BODY_FONT, COLORS, TITLE_FONT


@dataclass(frozen=True)
class PixelRect:
    """One fixed drawing rectangle in worksheet pixels."""

    x: int
    y: int
    width: int
    height: int


@dataclass(frozen=True)
class V2Geometry:
    """Approved 16:9-like first-screen geometry.

    Twenty-eight uniform 52-pixel columns form four equal seven-column modules,
    two equal fourteen-column chart panels and seven equal four-column action
    fields. All report-specific content must fit these rectangles rather than
    resizing the grid.
    """

    content_width: int = 52
    canvas_width: int = 1456
    header_height: int = 56
    filter_height: int = 52
    gap_height: int = 10
    accent_height: int = 7
    card_label_height: int = 36
    card_value_height: int = 96
    chart_height: int = 310
    section_height: int = 34
    table_header_height: int = 30
    table_row_height: int = 26

    @property
    def column_widths(self) -> tuple[int, ...]:
        return (self.content_width,) * 28

    @property
    def blocks(self) -> tuple[tuple[int, int], ...]:
        return ((0, 6), (7, 13), (14, 20), (21, 27))

    @property
    def chart_rects(self) -> tuple[PixelRect, PixelRect]:
        panel_width = self.content_width * 14
        y = (
            self.header_height + self.filter_height + self.gap_height
            + self.accent_height + self.card_label_height
            + self.card_value_height + self.gap_height
        )
        return (
            PixelRect(0, y, panel_width, self.chart_height),
            PixelRect(panel_width, y, panel_width, self.chart_height),
        )


V2_GEOMETRY = V2Geometry()

# Fixed Excel row indexes used by the cell-based PCS overview.
HEADER_ROW = 0
FILTER_ROW = 1
GAP_AFTER_FILTER_ROW = 2
CARD_ACCENT_ROW = 3
CARD_LABEL_ROW = 4
CARD_VALUE_ROW = 5
GAP_AFTER_CARDS_ROW = 6
CHART_FIRST_ROW = 7
CHART_LAST_ROW = 20
GAP_AFTER_CHART_ROW = 21
ACTION_SECTION_ROW = 22
ACTION_HEADER_ROW = 23
ACTION_FIRST_ROW = 24
ACTION_VISIBLE_ROWS = 8


@dataclass(frozen=True)
class V2Formats:
    canvas: Any
    header_title: Any
    header_owner: Any
    status_live: Any
    status_review: Any
    status_error: Any
    filter_label: Any
    filter_value: Any
    card_accent_teal: Any
    card_accent_gold: Any
    card_label: Any
    card_decimal: Any
    card_percent: Any
    card_integer: Any
    card_text: Any
    section: Any
    table_header: Any
    table_text: Any
    table_integer: Any
    table_decimal: Any
    table_percent: Any
    positive: Any
    negative: Any
    due: Any
    clear: Any


def make_v2_formats(workbook) -> V2Formats:
    """Create the exact V2 formats once per workbook."""

    add = workbook.add_format
    border = {
        "border": 1,
        "border_color": COLORS["line"],
    }
    table_border = {
        "bottom": 1,
        "bottom_color": COLORS["line"],
        "right": 1,
        "right_color": COLORS["line"],
    }
    return V2Formats(
        canvas=add({"bg_color": COLORS["canvas"]}),
        header_title=add({
            "font_name": TITLE_FONT, "font_size": 22, "bold": True,
            "font_color": COLORS["white"], "bg_color": COLORS["navy"],
            "align": "left", "valign": "vcenter", "indent": 1,
        }),
        header_owner=add({
            "font_name": BODY_FONT, "font_size": 10, "bold": True,
            "font_color": "#DDE7EE", "bg_color": COLORS["navy"],
            "align": "right", "valign": "vcenter",
        }),
        status_live=add({
            "font_name": BODY_FONT, "font_size": 10, "bold": True,
            "font_color": COLORS["white"], "bg_color": COLORS["green"],
            "align": "center", "valign": "vcenter", **border,
        }),
        status_review=add({
            "font_name": BODY_FONT, "font_size": 10, "bold": True,
            "font_color": COLORS["navy"], "bg_color": COLORS["amber_light"],
            "align": "center", "valign": "vcenter", **border,
        }),
        status_error=add({
            "font_name": BODY_FONT, "font_size": 10, "bold": True,
            "font_color": COLORS["white"], "bg_color": COLORS["red"],
            "align": "center", "valign": "vcenter", **border,
        }),
        filter_label=add({
            "font_name": BODY_FONT, "font_size": 9, "bold": True,
            "font_color": COLORS["navy"], "bg_color": COLORS["canvas"],
            "align": "left", "valign": "vcenter",
        }),
        filter_value=add({
            "font_name": BODY_FONT, "font_size": 11,
            "font_color": COLORS["navy"], "bg_color": COLORS["white"],
            "align": "left", "valign": "vcenter", "indent": 1, **border,
        }),
        card_accent_teal=add({"bg_color": COLORS["teal"]}),
        card_accent_gold=add({"bg_color": COLORS["gold"]}),
        card_label=add({
            "font_name": BODY_FONT, "font_size": 10, "bold": True,
            "font_color": COLORS["navy"], "bg_color": COLORS["white"],
            "align": "center", "valign": "vcenter",
            "left": 1, "right": 1, "left_color": COLORS["line"],
            "right_color": COLORS["line"],
        }),
        card_decimal=add({
            "font_name": TITLE_FONT, "font_size": 34, "bold": True,
            "font_color": COLORS["teal"], "bg_color": COLORS["white"],
            "align": "center", "valign": "vcenter", "num_format": "0.00",
            "left": 1, "right": 1, "bottom": 1,
            "left_color": COLORS["line"], "right_color": COLORS["line"],
            "bottom_color": COLORS["line"],
        }),
        card_percent=add({
            "font_name": TITLE_FONT, "font_size": 34, "bold": True,
            "font_color": COLORS["gold"], "bg_color": COLORS["white"],
            "align": "center", "valign": "vcenter", "num_format": "0%",
            "left": 1, "right": 1, "bottom": 1,
            "left_color": COLORS["line"], "right_color": COLORS["line"],
            "bottom_color": COLORS["line"],
        }),
        card_integer=add({
            "font_name": TITLE_FONT, "font_size": 34, "bold": True,
            "font_color": COLORS["teal"], "bg_color": COLORS["white"],
            "align": "center", "valign": "vcenter", "num_format": "#,##0",
            "left": 1, "right": 1, "bottom": 1,
            "left_color": COLORS["line"], "right_color": COLORS["line"],
            "bottom_color": COLORS["line"],
        }),
        card_text=add({
            "font_name": TITLE_FONT, "font_size": 24, "bold": True,
            "font_color": COLORS["navy"], "bg_color": COLORS["white"],
            "align": "center", "valign": "vcenter",
            "left": 1, "right": 1, "bottom": 1,
            "left_color": COLORS["line"], "right_color": COLORS["line"],
            "bottom_color": COLORS["line"],
        }),
        section=add({
            "font_name": TITLE_FONT, "font_size": 13, "bold": True,
            "font_color": COLORS["white"], "bg_color": COLORS["navy"],
            "align": "left", "valign": "vcenter", "indent": 1,
        }),
        table_header=add({
            "font_name": BODY_FONT, "font_size": 9, "bold": True,
            "font_color": COLORS["navy"], "bg_color": "#DCEAF2",
            "align": "center", "valign": "vcenter", **table_border,
        }),
        table_text=add({
            "font_name": BODY_FONT, "font_size": 10,
            "font_color": COLORS["navy"], "bg_color": COLORS["white"],
            "align": "left", "valign": "vcenter", "indent": 1,
            **table_border,
        }),
        table_integer=add({
            "font_name": BODY_FONT, "font_size": 10,
            "font_color": COLORS["navy"], "bg_color": COLORS["white"],
            "align": "center", "valign": "vcenter", "num_format": "#,##0",
            **table_border,
        }),
        table_decimal=add({
            "font_name": BODY_FONT, "font_size": 10,
            "font_color": COLORS["navy"], "bg_color": COLORS["white"],
            "align": "center", "valign": "vcenter", "num_format": "0.00",
            **table_border,
        }),
        table_percent=add({
            "font_name": BODY_FONT, "font_size": 10,
            "font_color": COLORS["navy"], "bg_color": COLORS["white"],
            "align": "center", "valign": "vcenter", "num_format": "0%",
            **table_border,
        }),
        positive=add({
            "font_name": BODY_FONT, "font_size": 10, "bold": True,
            "font_color": COLORS["green"], "bg_color": COLORS["white"],
            "align": "center", "valign": "vcenter", "num_format": "+0.00;-0.00;0.00",
            **table_border,
        }),
        negative=add({
            "font_name": BODY_FONT, "font_size": 10, "bold": True,
            "font_color": COLORS["red"], "bg_color": COLORS["white"],
            "align": "center", "valign": "vcenter", "num_format": "+0.00;-0.00;0.00",
            **table_border,
        }),
        due=add({
            "font_name": BODY_FONT, "font_size": 10, "bold": True,
            "font_color": COLORS["red"], "bg_color": COLORS["red_light"],
            "align": "center", "valign": "vcenter", **table_border,
        }),
        clear=add({
            "font_name": BODY_FONT, "font_size": 10,
            "font_color": COLORS["navy"], "bg_color": COLORS["green_light"],
            "align": "center", "valign": "vcenter", **table_border,
        }),
    )


def configure_v2_cell_canvas(ws, formats: V2Formats, *, zoom: int = 85) -> None:
    """Apply the measured PCS-style grid to a first-screen worksheet."""

    geometry = V2_GEOMETRY
    ws.hide_gridlines(2)
    ws.set_zoom(zoom)
    ws.set_landscape()
    ws.fit_to_pages(1, 1)
    ws.set_margins(0.15, 0.15, 0.2, 0.2)
    for column, width in enumerate(geometry.column_widths):
        ws.set_column_pixels(column, column, width, formats.canvas)
    ws.set_row_pixels(HEADER_ROW, geometry.header_height)
    ws.set_row_pixels(FILTER_ROW, geometry.filter_height)
    ws.set_row_pixels(GAP_AFTER_FILTER_ROW, geometry.gap_height, formats.canvas)
    ws.set_row_pixels(CARD_ACCENT_ROW, geometry.accent_height)
    ws.set_row_pixels(CARD_LABEL_ROW, geometry.card_label_height)
    ws.set_row_pixels(CARD_VALUE_ROW, geometry.card_value_height)
    ws.set_row_pixels(GAP_AFTER_CARDS_ROW, geometry.gap_height, formats.canvas)
    chart_row_height = geometry.chart_height // (CHART_LAST_ROW - CHART_FIRST_ROW + 1)
    for row in range(CHART_FIRST_ROW, CHART_LAST_ROW + 1):
        ws.set_row_pixels(row, chart_row_height, formats.canvas)
    ws.set_row_pixels(GAP_AFTER_CHART_ROW, geometry.gap_height, formats.canvas)
    ws.set_row_pixels(ACTION_SECTION_ROW, geometry.section_height)
    ws.set_row_pixels(ACTION_HEADER_ROW, geometry.table_header_height)
    for row in range(ACTION_FIRST_ROW, ACTION_FIRST_ROW + ACTION_VISIBLE_ROWS):
        ws.set_row_pixels(row, geometry.table_row_height)


def write_v2_header(
    ws,
    formats: V2Formats,
    title: str,
    *,
    owner: str = "WFMHUB  |  ANASS ASSRI",
    status: str = "DATA FRESH",
    status_kind: str = "LIVE",
) -> None:
    """Write the exact three-part V2 header."""

    ws.merge_range(HEADER_ROW, 0, HEADER_ROW, 16, title, formats.header_title)
    ws.merge_range(HEADER_ROW, 17, HEADER_ROW, 23, owner, formats.header_owner)
    status_format = {
        "LIVE": formats.status_live,
        "FINAL": formats.status_live,
        "PROVISIONAL": formats.status_review,
        "INCOMPLETE": formats.status_error,
    }.get(status_kind, formats.status_error)
    ws.merge_range(HEADER_ROW, 24, HEADER_ROW, 27, status, status_format)


def write_v2_filters(
    ws,
    formats: V2Formats,
    entries: Sequence[tuple[str, Any, dict[str, Any] | None]],
) -> None:
    """Write four equal, real Excel selector cells."""

    for (start, end), (label, value, validation) in zip(
        V2_GEOMETRY.blocks, entries[:4],
    ):
        ws.merge_range(
            FILTER_ROW, start, FILTER_ROW, start + 1,
            label.upper(), formats.filter_label,
        )
        ws.merge_range(
            FILTER_ROW, start + 2, FILTER_ROW, end,
            value, formats.filter_value,
        )
        if validation:
            ws.data_validation(
                FILTER_ROW, start + 2, FILTER_ROW, start + 2,
                validation,
            )


def write_v2_kpis(
    ws,
    formats: V2Formats,
    entries: Sequence[tuple[str, Any, str, Any | None]],
) -> None:
    """Write four cards with provably identical dimensions."""

    value_formats = {
        "decimal": formats.card_decimal,
        "percent": formats.card_percent,
        "integer": formats.card_integer,
        "text": formats.card_text,
    }
    accents = (formats.card_accent_teal, formats.card_accent_gold) * 2
    for index, ((start, end), (label, value, kind, cached)) in enumerate(zip(
        V2_GEOMETRY.blocks, entries[:4],
    )):
        ws.merge_range(
            CARD_ACCENT_ROW, start, CARD_ACCENT_ROW, end, "", accents[index],
        )
        ws.merge_range(
            CARD_LABEL_ROW, start, CARD_LABEL_ROW, end,
            label.upper(), formats.card_label,
        )
        fmt = value_formats.get(kind, formats.card_text)
        ws.merge_range(CARD_VALUE_ROW, start, CARD_VALUE_ROW, end, "", fmt)
        if isinstance(value, str) and value.startswith("="):
            ws.write_formula(
                CARD_VALUE_ROW, start, value, fmt,
                "" if cached is None else cached,
            )
        else:
            ws.write(CARD_VALUE_ROW, start, value, fmt)


def style_v2_chart(chart, *, title: str, kind: str = "line") -> None:
    """Apply the approved native Excel chart treatment and exact size."""

    chart.set_size({
        "width": V2_GEOMETRY.chart_rects[0].width,
        "height": V2_GEOMETRY.chart_rects[0].height,
    })
    chart.set_title({
        "name": title,
        "name_font": {"name": TITLE_FONT, "size": 14, "bold": True, "color": COLORS["navy"]},
        "overlay": False,
    })
    chart.set_legend({
        "position": "bottom",
        "font": {"name": BODY_FONT, "size": 9, "color": COLORS["muted"]},
    })
    chart.set_chartarea({
        "border": {"color": COLORS["line"], "width": 1},
        "fill": {"color": COLORS["white"]},
    })
    chart.set_plotarea({
        "border": {"none": True},
        "fill": {"color": COLORS["white"]},
    })
    if kind == "bar":
        chart.set_y_axis({
            "reverse": True,
            "major_gridlines": {"visible": False},
            "num_font": {"name": BODY_FONT, "size": 9, "color": COLORS["navy"]},
        })
    chart.set_x_axis({
        "major_gridlines": {
            "visible": True,
            "line": {"color": COLORS["line"], "width": 0.75},
        },
        "num_font": {"name": BODY_FONT, "size": 8, "color": COLORS["muted"]},
    })


def insert_v2_charts(ws, left_chart, right_chart) -> None:
    """Place two equal charts without worksheet-dependent scaling."""

    ws.insert_chart(
        CHART_FIRST_ROW, 0, left_chart,
        {"x_offset": 0, "y_offset": 0, "object_position": 3},
    )
    ws.insert_chart(
        CHART_FIRST_ROW, 14, right_chart,
        {"x_offset": 0, "y_offset": 0, "object_position": 3},
    )


def write_v2_section(ws, formats: V2Formats, title: str) -> None:
    ws.merge_range(
        ACTION_SECTION_ROW, 0, ACTION_SECTION_ROW, 14,
        title.upper(), formats.section,
    )
