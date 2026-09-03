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

The portable release includes a reviewed, data-free `pcs.xlsx` starter so this
folder is no longer empty on first use. For other products, create one from
menu option **Create an Excel Pivot/slicer master**, then follow
`docs/EXCEL_TEMPLATE_GUIDE.md`. The PCS master contains the named cell
`pFeedFolder` for PCS; its stable typed CSV files are refreshed under
`output/template_feeds/pcs/current/`. No raw extract is loaded to a worksheet.

The shipped PCS starter contains no source rows, Pivot caches, connections, or
external links. Use the supplied one-query-per-file scripts under
`templates/power_query` and follow `docs/EXCEL_TEMPLATE_GUIDE.md`.

These master files are intentionally ignored by Git so a locally populated
workbook cannot be published by accident.
