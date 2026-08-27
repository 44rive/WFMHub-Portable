#!/usr/bin/env python3
"""Build the blank, public FTE roster template shipped with WFMHub."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path
from typing import Any, Iterable

from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo


ROOT = Path(__file__).resolve().parents[1]
DEFAULT_TARGET = ROOT / "templates" / "FTE Count.xlsx"
HEADERS = [
    "Client ID",
    "Status",
    "Name",
    "Team leader",
    "Ops Manager",
    "LOB",
    "Market",
    "Language",
    "Location",
    "City",
    "FTE",
    "End date if leaver",
]
ROW_FIELDS = [
    "agent_id", "employment_status", "agent_name", "team_leader", "ops_manager",
    "lob", "market", "language", "location", "city", "fte", "end_date",
]


def build_template(
    target: Path = DEFAULT_TARGET,
    agent_rows: Iterable[dict[str, Any]] = (),
) -> Path:
    target.parent.mkdir(parents=True, exist_ok=True)
    agent_rows = list(agent_rows)
    workbook = Workbook()
    workbook.properties.title = "WFMHub standard FTE roster template"
    workbook.properties.subject = "Authoritative agent scope for WFMHub"
    workbook.properties.creator = "WFMHub Portable"
    workbook.properties.company = "WFM"

    guide = workbook.active
    guide.title = "START_HERE"
    guide.sheet_view.showGridLines = False
    guide.merge_cells("A1:F1")
    guide["A1"] = "WFMHub standard FTE roster"
    guide["A1"].font = Font(name="Calibri", size=16, bold=True, color="1F2933")
    guide["A1"].alignment = Alignment(vertical="center")
    guide.row_dimensions[1].height = 28
    instructions = [
        "1. Open the Agent sheet.",
        "2. Paste or type one row per agent below the headers.",
        "3. Keep Client ID and Name populated. Client ID must be stored as text so leading zeros survive.",
        "4. Keep the sheet name Agent and do not rename the header cells.",
        "5. Save this file as FTE Count.xlsx inside the configured FTE folder.",
        "6. Replace/update rows here; never paste worldwide LILO, status, or schedule data into this workbook.",
        "7. WFMHub uses this roster to admit agent rows by Agent ID or one unique normalized-name match.",
    ]
    guide["A3"] = "How to use it"
    guide["A3"].font = Font(name="Calibri", size=12, bold=True, color="0F3B42")
    for row, instruction in enumerate(instructions, 4):
        guide.cell(row, 1, instruction)
        guide.cell(row, 1).font = Font(name="Calibri", size=11, color="1F2933")
        guide.cell(row, 1).alignment = Alignment(wrap_text=True, vertical="top")
        guide.row_dimensions[row].height = 26
    guide["A13"] = "Required"
    guide["B13"] = "Client ID and Name"
    guide["A14"] = "Recommended"
    guide["B14"] = "Status, Team leader, Ops Manager, LOB, Market, Language, Location, City, and FTE"
    guide["A15"] = "Optional"
    guide["B15"] = "End date if leaver"
    for row in range(13, 16):
        guide.cell(row, 1).font = Font(bold=True, color="0F3B42")
    guide.column_dimensions["A"].width = 105
    guide.column_dimensions["B"].width = 42

    sheet = workbook.create_sheet("Agent")
    sheet.sheet_view.showGridLines = False
    sheet.freeze_panes = "A2"
    header_fill = PatternFill("solid", fgColor="0F3B42")
    header_font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    thin = Side(style="thin", color="D5DADD")
    for column, header in enumerate(HEADERS, 1):
        cell = sheet.cell(1, column, header)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(wrap_text=True, vertical="center")
        cell.border = Border(bottom=thin)
    sheet.row_dimensions[1].height = 32
    if agent_rows:
        for row_number, row in enumerate(agent_rows, 2):
            for column, field in enumerate(ROW_FIELDS, 1):
                sheet.cell(row_number, column, row.get(field))
    else:
        # Keep one blank table row so filters and structured formatting are
        # ready when the user first opens the template.
        for column in range(1, len(HEADERS) + 1):
            sheet.cell(2, column, None)
    last_table_row = max(2, len(agent_rows) + 1)
    table = Table(displayName="tblFTEAgents", ref=f"A1:L{last_table_row}")
    table.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium2",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False,
    )
    sheet.add_table(table)

    widths = [16, 14, 28, 24, 24, 18, 16, 14, 18, 18, 10, 20]
    for column, width in enumerate(widths, 1):
        sheet.column_dimensions[chr(64 + column)].width = width
    for row in range(2, 5001):
        sheet.cell(row, 1).number_format = "@"
        sheet.cell(row, 11).number_format = "0.00"
        sheet.cell(row, 12).number_format = "yyyy-mm-dd"

    status_validation = DataValidation(
        type="list",
        formula1='"Active,Inactive,Leaver"',
        allow_blank=True,
    )
    status_validation.promptTitle = "Employment status"
    status_validation.prompt = "Choose Active, Inactive, or Leaver."
    sheet.add_data_validation(status_validation)
    status_validation.add("B2:B5000")

    fte_validation = DataValidation(
        type="decimal",
        operator="between",
        formula1="0",
        formula2="1",
        allow_blank=True,
    )
    fte_validation.promptTitle = "FTE"
    fte_validation.prompt = "Enter a value from 0 to 1, for example 1 or 0.5."
    sheet.add_data_validation(fte_validation)
    fte_validation.add("K2:K5000")

    missing_fill = PatternFill("solid", fgColor="FCE8E6")
    sheet.conditional_formatting.add(
        "A2:A5000",
        FormulaRule(formula=['AND(COUNTA($A2:$L2)>0,$A2="")'], fill=missing_fill),
    )
    sheet.conditional_formatting.add(
        "C2:C5000",
        FormulaRule(formula=['AND(COUNTA($A2:$L2)>0,$C2="")'], fill=missing_fill),
    )
    sheet["A1"].comment = Comment("Required. Store as text; keep leading zeros.", "WFMHub")
    sheet["C1"].comment = Comment("Required. Use the agent name from operational exports when possible.", "WFMHub")
    sheet["L1"].comment = Comment("Leave blank for active employees.", "WFMHub")

    workbook.active = 0
    workbook.save(target)
    return target


def standardize_source(source: Path, target: Path) -> Path:
    sys.path.insert(0, str(ROOT / "src"))
    from wfmhub.ingestion import parse_fte

    rows = parse_fte(source.resolve(), "standardized-template").tables["raw.fte_agent"]
    return build_template(target.resolve(), rows)


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--source", type=Path, help="Optional populated FTE workbook to standardize")
    parser.add_argument("--output", type=Path, default=DEFAULT_TARGET)
    args = parser.parse_args()
    print(standardize_source(args.source, args.output) if args.source else build_template(args.output))
