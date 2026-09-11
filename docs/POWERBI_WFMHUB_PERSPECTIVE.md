# Power BI and WFMHub — recommended direction

Power BI can significantly simplify the reporting side, but it should not
replace the WFMHub calculation engine.

## Recommended architecture

```text
Untouched extracts
        ↓
WFMHub: ingest, clean, deduplicate and calculate
        ↓
Stable Power BI feed tables
        ↓
Power BI semantic model and dashboards
```

## What should change

Power BI should gradually replace most complicated generated management
workbooks:

- Service performance and historical Flashes
- PCS monitoring
- Absenteeism and shrinkage analysis
- Staffing and forecast comparison
- Realisations
- Bonus analytics
- Management dashboards

This provides proper slicers, cross-filtering, drill-through, mobile access,
scheduled subscriptions, TL/LOB views, and fewer fragile Excel formulas.

## What should remain in WFMHub

Python and WFMHub remain responsible for:

- Reading and validating extracts
- Active/leaver FTE filtering
- Agent identity
- Queue mappings
- Call-leg rules
- Attendance reconstruction
- Exact gap detection
- PTO/Away treatment
- Additive KPI components
- Data-quality controls

Power BI should not rebuild these rules directly from raw extracts. Maintaining
the same calculations in Python and Power BI would eventually produce
conflicting results.

## Recommended semantic model

- `DimDate`
- `DimEmployee`
- `DimLOB`
- `DimQueue`
- `FactServiceHour`
- `FactForecastInterval`
- `FactAttendanceDay`
- `FactAttendanceGap`
- `FactPCSAgentDay`
- `FactPCSCoaching`
- `FactTimeOff`
- `FactStaffing15Min`
- `FactBonusMonth`

WFMHub exports these as stable clean Power BI feeds. Power BI performs only
safe aggregation measures such as:

```text
TSL = SUM(Handled in SL) / SUM(SL Denominator)
PCS = SUM(Score Sum) / SUM(Valid Score Count)
Absence % = SUM(Absence Hours) / SUM(Scheduled Hours)
```

Percentages must be recalculated from summed components, never averaged.

## Implemented output

Every complete Hub update now atomically refreshes `Feed\PowerBI` with fixed
dimension/fact CSVs and writes `POWERBI_MANIFEST_CURRENT.csv` last. The
`_SETUP` subfolder contains the premium WFMHub theme, governed DAX measures and
the seven-page build map. Power BI Desktop setup is one-time; daily operation is
Hub **UPDATE**, then Power BI **Refresh**. The binary PBIX remains a Power BI
Desktop-owned file and is not synthesized by the portable Python runtime.

## What should remain in Excel

Excel remains useful for workflows requiring human editing:

- Attendance Review decisions and exact gap treatment
- Coaching action entry, until SharePoint Lists or Power Apps are introduced
- Temporary clean-data extracts that someone needs to send
- Manual investigation tables

Power BI is primarily an analysis and presentation layer. Power Apps can add
write-back, but the app and report have separate sharing and refresh concerns.

## Simplest first version

Start with one Power BI report containing:

1. Executive Overview
2. Service and Forecast
3. RTM History
4. PCS
5. Attendance and Absence
6. Staffing
7. Data Quality

WFMHub creates a dedicated fixed-name `PowerBI` feed folder. The user runs Hub
Update and then refreshes the PBIX. This avoids SQLite drivers, ODBC, and another
complicated Excel bridge.

## Local files versus SharePoint

For a pilot, Power BI Desktop can read the local Power BI feed folder. This is
simple, but the Power BI service normally needs an on-premises gateway to
refresh files stored only on a work computer.

For the later shared version, WFMHub's fixed feeds can be synchronized to the
restricted WFM SharePoint library. Power BI connects through the SharePoint
Folder connector using the organisational account. This is the cleaner route
for shared cloud refresh and does not require exposing the original extracts.

## Access and licensing checks

Having access to Power BI does not automatically confirm publishing and sharing
rights. Before building the shared production version, confirm:

- Power BI Free, Pro, PPU, or trial licence
- Permission to create or use a WFM workspace
- Whether the workspace has Fabric/Premium capacity
- Whether report viewers need individual Pro licences
- Permission to publish restricted employee and operational data

## Final recommendation

Do not rebuild WFMHub. Add a governed Power BI output layer. Keep WFMHub as the
deterministic backend and use Power BI for reusable dashboards, distribution,
filtering and management presentation.

Once Power BI reports are validated, retire the equivalent generated management
workbooks gradually. Keep Excel only where users must enter or approve data.

## Microsoft references

- [Microsoft Fabric and Power BI licences](https://learn.microsoft.com/en-us/fabric/enterprise/licenses)
- [SharePoint and OneDrive files in Power Query](https://learn.microsoft.com/en-us/power-query/sharepoint-onedrive-files)
- [Power BI scheduled refresh and gateways](https://learn.microsoft.com/en-us/power-bi/connect-data/refresh-scheduled-refresh)
- [Power Apps visual for Power BI](https://learn.microsoft.com/en-us/power-apps/maker/canvas-apps/powerapps-custom-visual)
