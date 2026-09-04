"""Generate public, data-free Excel starters for the portable release."""

from __future__ import annotations

from pathlib import Path

import xlsxwriter

from .reports import COLORS


def build_pcs_starter(path: Path) -> Path:
    """Create the shipped PCS layout without operational rows or connections."""

    path.parent.mkdir(parents=True, exist_ok=True)
    workbook = xlsxwriter.Workbook(path)
    workbook.set_properties({
        "title": "WFMHub PCS Management Master",
        "subject": "Data-free PivotTable and slicer starter",
        "author": "Anass ASSRI",
        "company": "WFM",
        "comments": "WFMHub public starter; contains no operational data.",
    })
    add = workbook.add_format
    title = add({
        "font_name": "Aptos Display", "font_size": 20, "bold": True,
        "font_color": COLORS["white"], "bg_color": COLORS["dark"],
        "align": "left", "valign": "vcenter", "indent": 1,
    })
    subtitle = add({
        "font_name": "Aptos", "font_size": 9, "font_color": COLORS["white"],
        "bg_color": COLORS["teal"], "align": "left", "valign": "vcenter", "indent": 1,
    })
    section = add({
        "font_name": "Aptos Display", "font_size": 11, "bold": True,
        "font_color": COLORS["teal"], "bottom": 2, "bottom_color": COLORS["gold"],
    })
    body = add({
        "font_name": "Aptos", "font_size": 10, "font_color": COLORS["dark"],
        "valign": "top", "text_wrap": True,
    })
    note = add({
        "font_name": "Aptos", "font_size": 9, "font_color": COLORS["muted"],
        "bg_color": COLORS["canvas"], "valign": "top", "text_wrap": True,
        "border": 1, "border_color": COLORS["thin"],
    })
    header = add({
        "font_name": "Aptos", "font_size": 10, "bold": True,
        "font_color": COLORS["white"], "bg_color": COLORS["teal"],
        "text_wrap": True, "valign": "vcenter", "bottom": 2,
        "bottom_color": COLORS["gold"],
    })
    input_format = add({
        "font_name": "Aptos", "font_size": 10, "font_color": COLORS["blue"],
        "bg_color": COLORS["blue_light"], "border": 1,
        "border_color": COLORS["thin"], "text_wrap": True,
    })
    card_label = add({
        "font_name": "Aptos", "font_size": 9, "bold": True,
        "font_color": COLORS["muted"], "bg_color": COLORS["canvas"],
        "align": "center", "valign": "vcenter", "border": 1,
        "border_color": COLORS["thin"],
    })
    card_value = add({
        "font_name": "Aptos Display", "font_size": 16, "bold": True,
        "font_color": COLORS["teal"], "bg_color": COLORS["white"],
        "align": "center", "valign": "vcenter", "border": 1,
        "border_color": COLORS["thin"], "text_wrap": True,
    })

    start = workbook.add_worksheet("START_HERE")
    start.hide_gridlines(2)
    start.set_tab_color(COLORS["gold"])
    start.merge_range("A1:H1", "WFMHUB  /  PCS MANAGEMENT MASTER", title)
    start.merge_range("A2:H2", "Data-free public starter  |  designed by Anass ASSRI", subtitle)
    steps = [
        "1. Run WFMHub and build PCS Performance for the dates you need.",
        "2. WFMHub updates four small PCS model files inside _system automatically.",
        "3. Follow docs\\EXCEL_TEMPLATE_GUIDE.md to paste the generated Power Query scripts.",
        "4. Load each query as Only Create Connection + Add to Data Model. Never load it to a sheet.",
        "5. Create the four relationships shown on SETUP, then build PivotTables and slicers.",
        "6. Move or link the finished management visuals to PCS_REPORT.",
        "7. Team coaching stays in COACHING_LOG. Use WFMHub Sync PCS Team, then Refresh All.",
    ]
    start.write("A4", "FIRST-TIME SETUP", section)
    for row, text_value in enumerate(steps, 5):
        start.merge_range(row - 1, 0, row - 1, 7, text_value, body)
        start.set_row(row - 1, 25)
    start.write("A14", "WHAT THIS FILE DOES NOT CONTAIN", section)
    start.merge_range(
        "A15:H17",
        "No source extracts, employee rows, external workbook links, Pivot cache, SQLite connection, "
        "or hidden raw-data sheet is included. The Data Model is created once from stable, typed "
        "CSV feeds generated on your own work machine. This workbook is never regenerated.",
        note,
    )
    start.set_column("A:A", 24)
    start.set_column("B:H", 15)

    setup = workbook.add_worksheet("SETUP")
    setup.hide_gridlines(2)
    setup.merge_range("A1:F1", "PCS MASTER SETUP", title)
    setup.merge_range("A2:F2", "The blue path cell is used by every supplied PCS Power Query.", subtitle)
    setup.write("A5", "Current feed folder", header)
    relative_formula = (
        '=IFERROR(LEFT(CELL("filename",A1),FIND("[",CELL("filename",A1))-1)'
        '&"..\\_system\\feeds\\pcs\\current","")'
    )
    setup.write_formula("B5", relative_formula, input_format, "")
    setup.merge_range("B6:F7", "Keep this master directly in Reports so the relative path remains correct.", note)
    setup.write("A9", "Daily sequence", section)
    setup.merge_range("A10:F11", "Build PCS in WFMHub -> Refresh All in Excel -> verify latest date -> save or send.", note)
    setup.write("A13", "Model relationships", section)
    setup.merge_range(
        "A14:F18",
        "PCS Agents[agent_id] 1 -> * PCS Agent Day[agent_id]\n"
        "PCS Agents[agent_id] 1 -> * PCS Calls[agent_id]\n"
        "PCS Dates[date] 1 -> * PCS Agent Day[business_date]\n"
        "PCS Dates[date] 1 -> * PCS Calls[business_date]\n"
        "Do not accept any other relationship suggested by Excel.",
        note,
    )
    setup.set_column("A:A", 24)
    setup.set_column("B:F", 22)
    cockpit = workbook.add_worksheet("PCS_REPORT")
    cockpit.hide_gridlines(2)
    cockpit.set_tab_color(COLORS["gold"])
    cockpit.merge_range("A1:N1", "PCS MANAGEMENT COCKPIT", title)
    cockpit.merge_range("A2:N2", "Daily / current MTD / prior-month comparison  |  refreshable Data Model layout", subtitle)
    labels = [
        "LATEST-DAY PCS", "LATEST PARTICIPATION", "CURRENT MTD PCS", "MTD VS PRIOR",
        "VALID RESPONSES", "SCORE <= 3", "COACHING COMPLETED", "ACTIONS RATE",
    ]
    for index, label in enumerate(labels):
        row = 4 if index < 4 else 9
        col = (index % 4) * 3
        cockpit.merge_range(row, col, row, col + 2, label, card_label)
        cockpit.merge_range(row + 1, col, row + 2, col + 2, "Link to Pivot / measure", card_value)
    cockpit.merge_range("A14:H14", "DAILY AND MONTHLY PCS TREND", section)
    cockpit.merge_range("A15:H27", "Place or link your PCS Trend PivotChart here.", note)
    cockpit.merge_range("J14:N14", "MANAGEMENT WATCHLIST", section)
    cockpit.merge_range("J15:N27", "Place a filtered low-score / low-sample PivotTable here.", note)
    cockpit.merge_range("A29:N29", "FILTERS", section)
    cockpit.merge_range("A30:N32", "Recommended slicers: Business Date, LOB, Team Leader and Agent. Use Report Connections only for PivotTables built from the same model table.", note)
    cockpit.set_column("A:N", 12)
    cockpit.set_landscape()
    cockpit.fit_to_pages(1, 1)

    coaching = workbook.add_worksheet("COACHING_LOG")
    coaching.hide_gridlines(2)
    coaching.set_tab_color(COLORS["blue"])
    coaching.merge_range("A1:E1", "PCS COACHING LOG", title)
    coaching.merge_range(
        "A2:E2",
        "Persistent team input  |  copy Coaching Key from the PCS Calls view and complete the blue fields",
        subtitle,
    )
    coaching.add_table(
        "A4:E5",
        {
            "name": "tblCoachingLog",
            "style": "Table Style Medium 2",
            "columns": [
                {"header": "Coaching Key"},
                {"header": "Coaching Status"},
                {"header": "Coaching Date"},
                {"header": "Coach"},
                {"header": "Coaching Comment"},
            ],
        },
    )
    coaching.set_column("A:A", 42, input_format)
    coaching.set_column("B:B", 20, input_format)
    coaching.set_column("C:C", 16, input_format)
    coaching.set_column("D:D", 22, input_format)
    coaching.set_column("E:E", 55, input_format)
    coaching.data_validation("B5:B10004", {
        "validate": "list",
        "source": ["Pending", "Completed", "Not required"],
    })
    coaching.freeze_panes(4, 0)

    pivot = workbook.add_worksheet("PIVOT_AREA")
    pivot.hide_gridlines(2)
    pivot.merge_range("A1:N1", "PIVOTTABLE BUILD AREA", title)
    pivot.merge_range("A2:N2", "Build pivots here first; then move charts or link cells into PCS_REPORT.", subtitle)
    pivot.merge_range("A4:G4", "MAIN AGENT / LOB PIVOT", section)
    pivot.merge_range("A5:G20", "Insert PivotTable from Data Model here. Use PCS Agent Day. Rows: LOB > Team Leader > Agent. Values: governed measures from FORMULAS.", note)
    pivot.merge_range("I4:N4", "COACHING PIVOT", section)
    pivot.merge_range("I5:N20", "Insert a separate PivotTable from PCS Calls here. Use Coaching Status, LOB, Team Leader and Business Date.", note)
    pivot.merge_range("A23:N23", "SAFE RELATIONSHIP RULE", section)
    pivot.merge_range("A24:N27", "Create only the four relationships listed on SETUP. Build slicers from PCS Agents and PCS Dates so the same slicer controls agent-day and coaching pivots.", note)
    pivot.set_column("A:N", 13)

    formulas = workbook.add_worksheet("FORMULAS")
    formulas.hide_gridlines(2)
    formulas.merge_range("A1:H1", "GOVERNED PCS MEASURES", title)
    formulas.merge_range("A2:H2", "Copy these exactly in Power Pivot > Measures > New Measure.", subtitle)
    formulas.write_row("A4", ["Measure", "DAX formula", "Format", "Meaning"], header)
    rows = [
        ("PCS Average", "DIVIDE(SUM('PCS Agent Day'[q1_score_sum]), SUM('PCS Agent Day'[valid_q1]))", "0.00", "Original O: response-weighted Q1 average"),
        ("PCS Participation", "DIVIDE(SUM('PCS Agent Day'[q1_nonblank]), SUM('PCS Agent Day'[pcs_status_1]))", "0.0%", "Original R: raw nonblank Q1 / PCSStatus=1"),
        ("Low Score %", "DIVIDE(SUM('PCS Agent Day'[score_le_3]), SUM('PCS Agent Day'[valid_q1]))", "0.0%", "Valid Q1 <=3 / all valid Q1"),
        ("Positive %", "DIVIDE(SUM('PCS Agent Day'[score_gt_3]), SUM('PCS Agent Day'[valid_q1]))", "0.0%", "Valid Q1 >3 / all valid Q1"),
        ("Coaching Opportunities", "COUNTROWS('PCS Calls')", "#,##0", "One row per valid low-score response"),
        ("Coaching Completed", "CALCULATE(COUNTROWS('PCS Calls'), 'PCS Calls'[coaching_status] = \"Completed\")", "#,##0", "Completed coaching rows after Sync PCS Team"),
        ("Actions Rate", "DIVIDE([Coaching Completed], [Coaching Opportunities])", "0.0%", "Original S business meaning without external link"),
    ]
    for row_index, values in enumerate(rows, 4):
        for column, value in enumerate(values):
            formulas.write(row_index, column, value, body)
        formulas.set_row(row_index, 35)
    formulas.set_column("A:A", 25)
    formulas.set_column("B:B", 95)
    formulas.set_column("C:C", 12)
    formulas.set_column("D:D", 45)

    for sheet in workbook.worksheets():
        sheet.set_zoom(90)
        sheet.set_footer("&LWFM Hub | Anass ASSRI&CPage &P of &N&RData-free starter")
    workbook.close()
    return path
