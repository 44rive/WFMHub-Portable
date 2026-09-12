# Power BI dashboard — beginner routine

The Power BI report is already built. You do not create Power Query, DAX,
relationships, charts, cards, pages, or slicers.

## First use

1. Install Microsoft Power BI Desktop on the work computer.
2. Add or replace your normal source extracts without editing them.
3. Open `WFMHub.cmd`.
4. Choose **Refresh source data once**, **All sources**, then the period you
   need. This updates the persistent SQLite database and writes the complete
   governed feed under `Feed\PowerBI`.
5. Choose **Open Power BI dashboard** under **Analyze**, or double-click
   `POWERBI.cmd`.
6. WFMHub copies the report to `Reports\Power BI\WFMHub BI`, puts the real Hub
   folder into the `HubRoot` parameter, and opens `WFMHub BI.pbip`.
7. In Power BI Desktop choose **Home > Refresh**.

## Normal routine after that

1. Put new extracts in their normal source folders.
2. Run one WFMHub update.
3. Open the dashboard with `POWERBI.cmd`.
4. Choose **Refresh** in Power BI Desktop.

WFMHub refresh and Power BI refresh are two different jobs:

- WFMHub reads changed extracts, keeps the same SQLite history, calculates the
  governed models, and atomically replaces `Feed\PowerBI` CSVs.
- Power BI imports those already-clean CSVs into its local semantic model and
  redraws the eight pages.

Power BI never reads the SQLite file or a raw extract. The file
`POWERBI_MANIFEST_CURRENT.csv` is the receipt proving that the complete feed
finished successfully.

## The eight pages

1. **Daily WFM Command** — today’s service, attendance and capacity decisions.
2. **Service & SL Drivers** — 15-minute pressure signals and queue evidence.
3. **Staff Preparation** — future required capacity versus published schedules,
   after approved PTO/Away.
4. **Workforce Realisation** — presence, productive delivery and governed AUX
   composition. These are three separate rates, not one synthetic score.
5. **Schedule Integrity & Patterns** — published versus observed shifts and
   supported recurrence from completed Agent Status evidence.
6. **Forecast Accuracy** — comparable 15-minute volume and AHT accuracy.
7. **Absence & Shrinkage** — final Verint rates, component mix and exact
   completeness cases by LOB and agent-day.
8. **Data Readiness** — source freshness, mapped status time and exact issues.

PCS is intentionally not inside Power BI. Keep using the permanent PCS Excel
tracker for the shared coaching workflow.

Use the four controls at the top of a page to filter. Use the left navigation
to move between pages. A blank KPI means its governed source or denominator is
not available; the report does not invent a replacement value.

## Saving and sharing

- Keep the installed `.pbip` project as the working source.
- Use **File > Save a copy** as `.pbix` when you want one publishable snapshot.
- The Hub does not automatically publish or send anything.
- If a future release upgrades the Power BI project contract, WFMHub archives
  the previous project under `Reports\Archive\Power BI` before installing the
  new one.
- This release advances that contract to version 5. The first open therefore
  archives the older project and installs the eight-page WFM control tower.
