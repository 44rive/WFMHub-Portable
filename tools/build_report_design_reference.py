"""Generate the approved, data-free WFMHub workbook design reference."""

from __future__ import annotations

from pathlib import Path

import xlsxwriter

from wfmhub.design import COLORS, REPORT_DESIGN_ID, REPORT_DESIGN_VERSION


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
    ws = workbook.add_worksheet(kind)
    ws.hide_gridlines(2)
    ws.set_zoom(85)
    ws.set_tab_color(COLORS["gold"])
    ws.set_column("A:N", 12)
    if kind == "PCS":
        header(ws, fmt, "PCS PERFORMANCE", "Illustrative layout only · governed data replaces every example")
        scope(ws, fmt, (("Period", "Current MTD"), ("LOB", "All"),
                        ("Team Leader", "RESULTS filter"), ("Agent", "RESULTS filter")))
        cards(ws, fmt, (("CURRENT MTD PCS", "4.54", "Weighted valid Q1"),
                        ("PARTICIPATION", "26.8%", "Q1 nonblank / status 1"),
                        ("PRIOR MTD PCS", "4.42", "Same days"),
                        ("CHANGE", "+0.12", "Current minus prior")))
        ws.merge_range("A11:N11", "PCS BY LOB · DAILY TREND · LOB PERFORMANCE TABLE", fmt["section"])
        table(ws, fmt, 12, ("LOB", "Current MTD", "Prior MTD", "Participation", "Responses", "Sample"),
              (("RSA NL", 4.62, 4.51, "28.1%", 412, "OK"),
               ("RSA BE", 4.38, 4.31, "24.7%", 286, "OK"),
               ("FORD NL", 4.71, 4.66, "28.9%", 398, "OK"),
               ("OEM FR", 4.55, 4.48, "26.1%", 312, "OK")))
    elif kind == "RTM":
        header(ws, fmt, "RTM DAILY CONTROL", "Combined service, attendance pulse and exact call actions")
        scope(ws, fmt, (("Date", "08 Sep 2026"), ("Snapshot", "Through 17:59"),
                        ("LOB", "All operational"), ("Attendance", "17:45")))
        cards(ws, fmt, (("LOBS ON TARGET", "3 / 4", "Configured target"),
                        ("DEMAND VARIANCE", "+84", "Actual minus forecast"),
                        ("NO SHOW HC", "6", "Evidence proven"),
                        ("CALL NOW", "8", "Exact lists by LOB")))
        ws.merge_range("A11:N11", "SERVICE LEVEL BY LOB · LINKED OPERATIONAL TABLE", fmt["section"])
        table(ws, fmt, 12, ("LOB", "TSL", "Target", "Actual", "Forecast", "Variance", "No Show", "Call Now"),
              (("RSA NL", "89.9%", "80.0%", 840, 810, 30, 2, 2),
               ("RSA BE", "85.8%", "80.0%", 530, 552, -22, 1, 2),
               ("FORD NL", "97.3%", "80.0%", 320, 300, 20, 1, 1),
               ("OEM", "94.9%", "80.0%", 410, 354, 56, 2, 3)))
    else:
        header(ws, fmt, "ATTENDANCE REVIEW", "Completed-day schedule-versus-observed decisions")
        scope(ws, fmt, (("Period", "01–07 Sep 2026"), ("Completed", "Through 07 Sep"),
                        ("Evidence", "Agent Status + LILO"), ("Decision", "REVIEW BOARD")))
        cards(ws, fmt, (("REVIEW GAPS", "18", "Exact intervals"),
                        ("GAP HOURS", "7.4", "Exact minutes / 60"),
                        ("OPEN DECISIONS", "11", "Awaiting review"),
                        ("MISSING EVIDENCE", "2", "Never treated as zero")))
        ws.merge_range("A11:N11", "BY-LOB ACTION SUMMARY · SCHEDULE ABOVE ACTUAL", fmt["section"])
        table(ws, fmt, 12, ("LOB", "Exact Gaps", "Gap Hours", "Agents", "Open", "Approved", "Dismissed"),
              (("RSA NL", 6, 2.8, 4, 3, 2, 1), ("RSA BE", 5, 2.1, 4, 4, 1, 0),
               ("FORD NL", 4, 1.6, 3, 2, 1, 1), ("OEM FR", 3, 0.9, 2, 2, 1, 0)))
        ws.write("A20", "Editable decision example", fmt["section"])
        ws.write("A21", "Decision Status", fmt["body"])
        ws.write("B21", "Open", fmt["editable"])
    ws.merge_range("A25:N25", "Examples are visual only. Calculations, targets and LOBs always come from WFMHub authorities.", fmt["note"])


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
    for kind in ("PCS", "RTM", "ATTENDANCE"):
        blueprint_sheet(workbook, fmt, kind)
    for sheet in workbook.worksheets():
        sheet.set_footer("&LPrepared by Anass ASSRI | WFM&CDesign reference&RConfidential")
    workbook.close()
    print(OUTPUT)


if __name__ == "__main__":
    main()
