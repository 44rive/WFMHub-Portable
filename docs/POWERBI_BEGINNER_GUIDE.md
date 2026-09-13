# Power BI — beginner operating guide

The dashboard is already built. You do not create Power Query, relationships,
DAX, charts or navigation yourself.

## First installation

1. Install Microsoft Power BI Desktop.
2. Install the current WFMHub portable release.
3. Put the untouched extracts in their normal folders and run `UPDATE.cmd`.
4. Wait until `Feed\PowerBI\POWERBI_MANIFEST_CURRENT.csv` exists.
5. Double-click `POWERBI.cmd`, or choose **Analyze > Open Power BI dashboard**.
6. WFMHub installs the project under `Reports\Power BI\WFMHub BI`, updates only
   the local `HubRoot` parameter and opens `WFMHub BI.pbip`.
7. In Power BI Desktop choose **Home > Refresh**.

## Normal routine

1. Add the new untouched extracts.
2. Run `UPDATE.cmd`. This validates them, updates the durable SQLite database
   and publishes the clean Power BI CSV feed.
3. Open the same installed PBIP with `POWERBI.cmd`.
4. Choose **Home > Refresh** in Power BI Desktop. This imports the new CSV rows
   and redraws the report.

These two refreshes are different and both are required. A new WFMHub release
does not require recreating the database; migrations upgrade it in place.

## The five pages

1. **Today's Control** — latest service, resource and attendance priorities.
2. **Staff Preparation** — required versus gross/PTO/Away/net published capacity
   at 15-minute Planning Group and Staff Type grain.
3. **Intraday Service** — combined service result by Management LOB and exact
   queue diagnosis.
4. **Attendance & Schedule Review** — published schedule above chronological
   Agent Status evidence, with exact residual gaps below.
5. **Historical Review** — demand, forecast, requirement, scheduled delivery,
   final absence and final shrinkage, with a comparison period.

Use the horizontal menu to move between pages. Selectors are connected and
synchronized: choosing a Management LOB limits its Planning Groups, Staff Types,
employees or queues wherever that relationship applies. A blank card means its
required evidence is unavailable in the selected scope; it is not a zero.

## When to open Excel

- open `RTM Daily Control.xlsx` for exact Flash and same-day callout rows;
- open `Staffing Preparation.xlsx` for exact 15-minute shortages and maintain
  only the blue columns in its persistent `ACTIONS` ledger;
- open `Attendance Review.xlsx` for Gap ID, exact start/end and source evidence;
- open `Realisations.xlsx` for detailed historical rows;
- open `Final Absenteeism & Shrinkage.xlsx` for final Verint activity components.

Power BI is for monitoring and prioritization. Excel is for exact portable rows
or durable action fields. Attendance Review is read-only: make the correction in
Verint, export final Activities and refresh; covered residuals then disappear.

## Upgrade check

Project contract `7` has exactly these five pages. On the first open after an
upgrade, choose **Home > Refresh**, then confirm the horizontal navigation,
selector behavior, cards, timeline and tables render correctly. The repository
validator cannot replace this one Power BI Desktop check.

PCS remains its separate permanent collaborative Excel tracker and is not part
of this Power BI model.
