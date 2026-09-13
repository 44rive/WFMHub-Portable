# Power BI — beginner operating guide

The dashboard is already built. You do not create Power Query, relationships,
DAX, charts or pages yourself.

## First installation

1. Install Microsoft Power BI Desktop on the work computer.
2. Install the current WFMHub portable release normally.
3. Add the untouched extracts and run `UPDATE.cmd`.
4. Wait for the Hub update and `Feed\PowerBI\POWERBI_MANIFEST_CURRENT.csv` to
   complete.
5. Open WFMHub and choose **Analyze > Open Power BI dashboard**, or double-click
   `POWERBI.cmd`.
6. WFMHub installs the governed project under
   `Reports\Power BI\WFMHub BI`, writes the correct local `HubRoot` parameter,
   and opens `WFMHub BI.pbip`.
7. In Power BI Desktop choose **Home > Refresh**.

## Normal routine

1. Put the new extracts in their normal source folders without editing them.
2. Close Excel files that the Hub must replace and run `UPDATE.cmd`.
3. When the update finishes, open the same Power BI project with `POWERBI.cmd`.
4. In Power BI Desktop choose **Home > Refresh**.

There are two refreshes because they do different jobs:

- WFMHub validates sources, updates durable SQLite marts and atomically publishes
  already-clean fixed CSV facts to `Feed\PowerBI`.
- Power BI imports those CSV facts into its local semantic model and redraws the
  visuals.

Power BI never reads SQLite or raw extracts. It does not calculate source
classification or silently replace missing data with zero.

## The five pages

1. **Forecast & Requirement** — Verint Volume and Absolute Required FTE at
   native 15-minute Staff Type grain.
2. **Staff Preparation** — required capacity versus gross/net published
   schedules by Planning Group and Staff Type, including PTO/Away impact.
3. **Intraday Control** — combined service result at Management LOB plus the
   separate current resource position by Planning Group/Staff Type.
4. **Attendance & Schedule Review** — published versus observed presence,
   exact residual gaps after final Activities overlap, and break/meal control.
5. **Performance Review** — standard WFM cycle close: demand, forecast,
   requirement, schedule, observed/productive delivery, final absence and
   shrinkage.

RSA BE is intentionally one combined service-level result. RSA BE FR and RSA BE
VL remain separate for forecast, requirements, schedules and staffing. A Verint
forecast field called `Queue Name` contains a workforce Staff Type; it is not a
Storm call queue.

## Filters and drill detail

Use the top strip from left to right: period/date, Management LOB, Planning
Group or Team Leader, then Staff Type/checkpoint/exception depending on the
page. Selecting Management LOB filters its governed children. The detail table
at the bottom is the evidence behind the cards and charts.

If a card is blank, check the source period and the mapping-status fields in its
detail table. Blank means the required evidence was not supplied; it is not a
zero.

## Upgrades and safety

WFMHub project contract `6` has exactly these five pages. A newer portable
release archives the installed older project under `Reports\Archive\Power BI`
before installing the new one. The database, extracts and Excel products are
not rebuilt just because the report layout changes.

The final acceptance step is always to open the upgraded PBIP once in Power BI
Desktop, refresh it, and confirm all five pages render. Source validation can
prove JSON/TMDL and field bindings, but Desktop is the final renderer.

PCS remains the separate permanent collaborative Excel tracker and is not
included in this Power BI model.
