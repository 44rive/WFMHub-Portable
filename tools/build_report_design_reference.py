"""Generate the approved, data-free WFMHub workbook design reference."""

from __future__ import annotations

from pathlib import Path

import xlsxwriter

from wfmhub.design import COLORS, REPORT_DESIGN_ID, REPORT_DESIGN_VERSION
from wfmhub.excel_layout import (
    V2ChartSpec,
    render_v2_dashboard,
)


ROOT = Path(__file__).resolve().parents[1]
OUTPUT = ROOT / "docs" / "WFMHub Report Design Reference.xlsx"


def formats(workbook: xlsxwriter.Workbook) -> dict[str, xlsxwriter.format.Format]:
    add = workbook.add_format
    border = {"border": 1, "border_color": COLORS["line"]}
    return {
        "title": add({"font_name": "Aptos Display", "font_size": 18, "bold": True,
                      "font_color": COLORS["white"], "bg_color": COLORS["navy"],
                      "valign": "vcenter", "indent": 1}),
        "subtitle": add({"font_name": "Aptos", "font_size": 9,
                         "font_color": COLORS["white"], "bg_color": COLORS["teal"],
                         "valign": "vcenter", "indent": 1}),
        "updated": add({"font_name": "Aptos", "font_size": 10, "bold": True,
                        "font_color": COLORS["white"], "bg_color": COLORS["green"],
                        "align": "center", "valign": "vcenter"}),
        "scope_label": add({"font_name": "Aptos", "font_size": 8, "bold": True,
                            "font_color": COLORS["muted"], "bg_color": COLORS["canvas"],
                            "valign": "vcenter", "indent": 1, **border}),
        "scope_value": add({"font_name": "Aptos", "font_size": 10, "bold": True,
                            "font_color": COLORS["ink"], "bg_color": COLORS["white"],
                            "valign": "vcenter", "indent": 1, **border}),
        "card_label": add({"font_name": "Aptos", "font_size": 9, "bold": True,
                           "font_color": COLORS["muted"], "bg_color": COLORS["canvas"],
                           "align": "center", "valign": "vcenter", **border}),
        "card": add({"font_name": "Aptos Display", "font_size": 20, "bold": True,
                     "font_color": COLORS["navy"], "bg_color": COLORS["white"],
                     "align": "center", "valign": "vcenter", **border}),
        "card_note": add({"font_name": "Aptos", "font_size": 8,
                          "font_color": COLORS["muted"], "bg_color": COLORS["canvas"],
                          "align": "center", "valign": "vcenter", **border}),
        "section": add({"font_name": "Aptos Display", "font_size": 11, "bold": True,
                        "font_color": COLORS["teal"], "bottom": 2,
                        "bottom_color": COLORS["gold"]}),
        "header": add({"font_name": "Aptos", "font_size": 9, "bold": True,
                       "font_color": COLORS["white"], "bg_color": COLORS["teal"],
                       "bottom": 2, "bottom_color": COLORS["gold"]}),
        "body": add({"font_name": "Aptos", "font_size": 9,
                     "font_color": COLORS["ink"], "bottom": 1,
                     "bottom_color": COLORS["line"]}),
        "note": add({"font_name": "Aptos", "font_size": 9,
                     "font_color": COLORS["muted"], "text_wrap": True}),
        "editable": add({"font_name": "Aptos", "font_size": 9,
                         "font_color": COLORS["blue"], "bg_color": COLORS["blue_light"],
                         **border}),
    }


def header(ws, fmt, title: str, subtitle: str, last_col: int = 13) -> None:
    ws.merge_range(0, 0, 0, last_col - 3, title, fmt["title"])
    ws.merge_range(0, last_col - 2, 0, last_col, "UPDATED", fmt["updated"])
    ws.merge_range(1, 0, 1, last_col, subtitle, fmt["subtitle"])
    ws.set_row(0, 34)
    ws.set_row(1, 21)


def scope(ws, fmt, entries: tuple[tuple[str, str], ...], last_col: int = 13) -> None:
    width = (last_col + 1) // len(entries)
    start = 0
    for index, (label, value) in enumerate(entries):
        end = last_col if index == len(entries) - 1 else start + width - 1
        ws.write(3, start, label.upper(), fmt["scope_label"])
        ws.merge_range(3, start + 1, 3, end, value, fmt["scope_value"])
        start = end + 1


def cards(ws, fmt, values: tuple[tuple[str, str, str], ...], last_col: int = 13) -> None:
    starts = (0, 4, 7, 11) if last_col == 13 else (0, 4, 8, 12)
    for start, (label, value, note) in zip(starts, values):
        end = min(last_col, start + 2)
        ws.merge_range(5, start, 5, end, label, fmt["card_label"])
        ws.merge_range(6, start, 7, end, value, fmt["card"])
        ws.merge_range(8, start, 8, end, note, fmt["card_note"])


def table(ws, fmt, row: int, headers: tuple[str, ...], rows: tuple[tuple[object, ...], ...]) -> None:
    for col, value in enumerate(headers):
        ws.write(row, col, value, fmt["header"])
    for offset, values in enumerate(rows, row + 1):
        for col, value in enumerate(values):
            ws.write(offset, col, value, fmt["body"])


def blueprint_sheet(workbook, fmt, kind: str) -> None:
    del fmt
    ws = workbook.add_worksheet(kind)
    lobs = ("RSA NL", "RSA BE", "FORD NL", "OEM")
    definitions = {
        "PCS": (
            "PCS OPERATIONS", (("Period", "Current MTD"), ("LOB", "All"), ("Team Leader", "All"), ("Agent", "All")),
            (("CURRENT PCS", 4.54, "decimal"), ("PARTICIPATION", .268, "percent"), ("PRIOR PCS", 4.42, "decimal"), ("CHANGE", .12, "decimal")),
            V2ChartSpec("PCS BY LOB", "bar", lobs, (("Current", (4.62, 4.38, 4.71, 4.55), COLORS["teal"]), ("Prior", (4.51, 4.31, 4.66, 4.48), COLORS["muted"]))),
            V2ChartSpec("DAILY PCS TREND", "line", ("01", "02", "03", "04", "05"), (("Current", (4.25, 4.32, 4.39, 4.48, 4.54), COLORS["teal"]), ("Prior", (4.18, 4.27, 4.31, 4.38, 4.42), COLORS["muted"]))),
            "TEAM PERFORMANCE & ACTIONS", ("TEAM LEADER", "PCS AGENTS", "PARTICIPATION", "CURRENT PCS", "PRIOR PCS", "CHANGE", "COACHING DUE"),
            (("Sophie Martin", 12, .29, 4.62, 4.51, .11, 2), ("Karim Belkacem", 10, .25, 3.92, 4.08, -.16, 5)),
            ("text", "integer", "percent", "decimal", "decimal", "change", "alert"),
        ),
        "RTM": (
            "RTM DAILY CONTROL", (("Date", "08 Sep 2026"), ("Snapshot", "Through 17:59"), ("LOB", "All 4"), ("Attendance", "17:55")),
            (("LOBS ON TARGET", "4 / 4", "text"), ("DEMAND VARIANCE", 84, "integer"), ("NO SHOW HC", 6, "integer"), ("CALL NOW", 8, "integer")),
            V2ChartSpec("SERVICE LEVEL BY LOB", "bar", lobs, (("TSL", (.899, .858, .973, .949), COLORS["teal"]), ("Target", (.8, .8, .8, .8), COLORS["muted"])), "percent", 0, 1),
            V2ChartSpec("ACTUAL VS FORECAST", "column", lobs, (("Actual", (840, 530, 320, 410), COLORS["teal"]), ("Forecast", (810, 552, 300, 354), COLORS["muted"]))),
            "LOB SERVICE & ATTENDANCE ACTIONS", ("LOB", "TSL", "VAR CALLS", "NO SHOW", "LATE", "EARLY LEAVE", "CALL NOW"),
            (("RSA NL", .899, 30, 2, 1, 0, 2), ("RSA BE", .858, -22, 1, 2, 1, 2)),
            ("text", "percent", "change", "integer", "integer", "integer", "alert"),
        ),
        "ATTENDANCE": (
            "ATTENDANCE REVIEW", (("Period", "01–07 Sep 2026"), ("Completed", "Through 07 Sep"), ("Evidence", "Status + LILO"), ("Decision", "Review board")),
            (("REVIEW GAPS", 18, "integer"), ("GAP HOURS", 7.4, "decimal"), ("OPEN DECISIONS", 11, "integer"), ("MISSING EVIDENCE", 2, "integer")),
            V2ChartSpec("GAP HOURS BY LOB", "bar", lobs, (("Gap hours", (2.8, 2.1, 1.6, .9), COLORS["teal"]),)),
            V2ChartSpec("DECISION STATUS", "column", ("Open", "Approved", "Dismissed", "Missing"), (("Cases", (11, 4, 3, 2), COLORS["teal"]),)),
            "BY-LOB REVIEW ACTIONS", ("LOB", "EXACT GAPS", "GAP HOURS", "AGENTS", "OPEN", "APPROVED", "DISMISSED"),
            (("RSA NL", 6, 2.8, 4, 3, 2, 1), ("RSA BE", 5, 2.1, 4, 4, 1, 0)),
            ("text", "integer", "decimal", "integer", "alert", "integer", "integer"),
        ),
        "STAFFING": (
            "STAFFING & COVERAGE", (("Period", "Current + 4 weeks"), ("Mode", "Future plan"), ("LOB", "All"), ("Language", "All")),
            (("PEAK GAP FTE", 4.5, "decimal"), ("FUTURE GAP HOURS", 36.8, "decimal"), ("FORECAST COVERAGE", .943, "percent"), ("PTO / AWAY IMPACT", 61, "decimal")),
            V2ChartSpec("REQUIRED VS NET SCHEDULED HOURS", "column", lobs, (("Required", (420, 310, 245, 280), COLORS["teal"]), ("Net scheduled", (401, 298, 251, 260), COLORS["muted"]))),
            V2ChartSpec("PEAK GAP BY DAY", "line", ("Mon", "Tue", "Wed", "Thu", "Fri"), (("Gap FTE", (2.1, 3.4, 4.5, 2.8, 3.9), COLORS["red"]),)),
            "PRIORITIZED CAPACITY ACTIONS", ("INTERVAL", "LOB / LANGUAGE", "MODE", "REQUIRED FTE", "NET / OBSERVED", "GAP FTE", "STATE"),
            (("09 Sep 10:00", "RSA BE / FR-NL", "Future", 22.5, 18, 4.5, "FUTURE GAP"),),
            ("text", "text", "text", "decimal", "decimal", "decimal", "alert"),
        ),
        "REALISATIONS": (
            "REALISATIONS", (("Period", "Current MTD"), ("LOB", "All mapped"), ("Grain", "Daily"), ("Data state", "Reviewed")),
            (("ACTUAL VOLUME", 49620, "integer"), ("FORECAST ATTAINMENT", 1.012, "percent"), ("ROUTED RATE", .958, "percent"), ("WEIGHTED AHT", 294, "decimal")),
            V2ChartSpec("ACTUAL VS FORECAST BY LOB", "column", lobs, (("Actual", (18400, 12150, 8830, 10240), COLORS["teal"]), ("Forecast", (17950, 12600, 8460, 9980), COLORS["muted"]))),
            V2ChartSpec("SERVICE LEVEL VS TARGET", "bar", lobs, (("Service level", (.899, .858, .973, .949), COLORS["teal"]), ("Target", (.8, .8, .8, .8), COLORS["muted"])), "percent", 0, 1),
            "LOB REALISATION SUMMARY", ("LOB", "ACTUAL", "FORECAST", "ATTAINMENT", "SERVICE LEVEL", "ABSENCE %", "STATE"),
            (("RSA NL", 18400, 17950, 1.025, .899, .041, "READY"),),
            ("text", "integer", "integer", "percent", "percent", "percent", "alert"),
        ),
        "ABSENCE": (
            "ABSENTEEISM & SHRINKAGE", (("Period", "Current MTD"), ("LOB", "All"), ("Team Leader", "All"), ("Agent", "All")),
            (("ABSENCE RATE", .043, "percent"), ("SHRINKAGE RATE", .118, "percent"), ("FINALIZED COVERAGE", .988, "percent"), ("REVIEW CASES", 7, "integer")),
            V2ChartSpec("ABSENCE & SHRINKAGE BY LOB", "bar", lobs, (("Absence", (.041, .052, .033, .046), COLORS["red"]), ("Shrinkage", (.118, .129, .104, .123), COLORS["teal"])), "percent", 0),
            V2ChartSpec("DAILY ABSENCE & SHRINKAGE HOURS", "line", ("01", "02", "03", "04", "05"), (("Absence", (31, 27, 38, 42, 35), COLORS["red"]), ("Shrinkage", (72, 68, 81, 77, 74), COLORS["teal"]))),
            "PRIORITIZED ABSENCE REVIEW CASES", ("DATE", "AGENT", "TEAM LEADER", "LOB", "RESULT STATUS", "ABSENCE H", "ACTION STATUS"),
            (("07 Sep", "Amina El Idrissi", "Sophie Martin", "RSA NL", "ABSENCE RECORDED", 8, "Pending"),),
            ("text", "text", "text", "text", "text", "decimal", "alert"),
        ),
        "BONUS": (
            "BONUS MANAGEMENT", (("Period", "2026-08"), ("Population", "All"), ("Team Lead", "All"), ("Result status", "All")),
            (("TOTAL PAYOUT", 184520, "money"), ("PAID AGENTS", 84, "integer"), ("AVG PAID PAYOUT", 2196.67, "money"), ("REVIEW ITEMS", 6, "integer")),
            V2ChartSpec("PAYOUT BY POPULATION", "column", ("RSA", "FORD", "OEM"), (("Total payout", (92100, 54120, 38300), COLORS["gold"]),)),
            V2ChartSpec("KPI ATTAINMENT", "bar", ("AHT", "Productivity", "PCS", "QM", "Abs%"), (("Attainment", (.82, .76, .69, .88, .73), COLORS["teal"]),), "percent", 0, 1),
            "POPULATION PAYOUT SUMMARY", ("POPULATION", "AGENTS", "PAID AGENTS", "PAYOUT RATE", "TOTAL PAYOUT", "AVG ACHIEVEMENT", "REVIEW"),
            (("RSA", 48, 44, .917, 92100, 1.06, 2),),
            ("text", "integer", "integer", "percent", "money", "percent", "alert"),
        ),
    }
    title, filters, kpis, left, right, action_title, action_headers, action_rows, action_kinds = definitions[kind]
    render_v2_dashboard(
        workbook, ws, title=title, filters=filters, kpis=kpis,
        left_chart=left, right_chart=right, action_title=action_title,
        action_headers=action_headers, action_rows=action_rows,
        action_kinds=action_kinds,
        status="REFERENCE", status_kind="PROVISIONAL",
        status_note="Illustrative values only; production reports use governed data.",
    )


def main() -> None:
    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    workbook = xlsxwriter.Workbook(OUTPUT)
    workbook.set_properties({
        "title": "WFMHub Report Design Reference",
        "author": "Anass ASSRI",
        "comments": f"{REPORT_DESIGN_ID} {REPORT_DESIGN_VERSION}; illustrative data only",
    })
    fmt = formats(workbook)
    ws = workbook.add_worksheet("DESIGN SYSTEM")
    ws.hide_gridlines(2)
    header(ws, fmt, "WFMHUB REPORT SYSTEM", f"{REPORT_DESIGN_ID} {REPORT_DESIGN_VERSION} · approved visual contract")
    ws.merge_range("A4:N4", "TOKENS", fmt["section"])
    ws.write_row("A6", ("Token", "Hex", "Purpose"), fmt["header"])
    token_rows = (
        ("Navy", COLORS["navy"], "Titles and totals"),
        ("Teal", COLORS["teal"], "Primary series and table headers"),
        ("Gold", COLORS["gold"], "Small separators and primary tabs"),
        ("Blue", COLORS["blue"], "Editable cells only"),
        ("Green / Amber / Red", "Semantic", "Proven state only"),
    )
    for row, values in enumerate(token_rows, 6):
        ws.write_row(row, 0, values, fmt["body"])
    ws.merge_range("A14:N14", "FIRST SCREEN", fmt["section"])
    for offset, text in enumerate((
        "Compact navy title + update badge; one teal context sentence.",
        "Honest scope strip; generated snapshots never display fake dropdowns.",
        "Exactly four headline cards and no more than two purposeful charts.",
        "One filterable action/reconciliation table; blue means editable.",
        "No AI branding, decorative visuals, guessed target, or hidden KPI arithmetic.",
    ), 15):
        ws.merge_range(offset, 0, offset, 13, text, fmt["note"])
    ws.set_column("A:A", 24)
    ws.set_column("B:B", 16)
    ws.set_column("C:N", 14)
    for kind in (
        "PCS", "RTM", "ATTENDANCE", "STAFFING", "REALISATIONS",
        "ABSENCE", "BONUS",
    ):
        blueprint_sheet(workbook, fmt, kind)
    for sheet in workbook.worksheets():
        sheet.set_footer("&LPrepared by Anass ASSRI | WFM&CDesign reference&RConfidential")
    workbook.close()
    print(OUTPUT)


if __name__ == "__main__":
    main()
