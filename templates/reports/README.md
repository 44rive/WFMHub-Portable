# WFMHub Excel templates

This folder holds Excel-authored master files with native PivotTables and
slicers. WFMHub creates the starter once and never rewrites it afterward,
because Python libraries can damage Excel-only slicer and Data Model parts.

Expected optional master names:

- `pcs.xlsx`
- `bonus.xlsx`
- `service.xlsx`
- `staffing.xlsx`
- `attendance.xlsx`
- `corrections.xlsx`
- `absence.xlsx`

Create one from menu option **Create an Excel Pivot/slicer master**, then follow
`docs/EXCEL_TEMPLATE_GUIDE.md`. The master contains the named cell
`pModelDataPath`; compact CSV files are refreshed under
`output/model_data/<report>/`. No raw extract is loaded to a worksheet.

These master files are intentionally ignored by Git so a locally populated
workbook cannot be published by accident.
