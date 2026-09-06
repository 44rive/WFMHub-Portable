# WFMHub beginner guide

Think of WFMHub as a washing machine for data:

1. You place new extracts in the usual source folders.
2. WFMHub reads them but never changes them.
3. It remembers the useful data in SQLite.
4. It calculates the agreed formulas.
5. It creates the Excel file you asked for.

You do not need to install Python, SQLite, DuckDB, ODBC, or Power BI.

## The two files you click

- `SETUP.cmd`: use it once after extracting WFMHub.
- `WFMHub.cmd`: use it for normal work.

Never move only the `.cmd` file. Keep the full WFMHub folder together.

## Setup once

1. Right-click the downloaded ZIP and choose **Extract All**.
2. Open the extracted WFMHub folder.
3. Double-click `SETUP.cmd`.
4. Paste the folder containing `FTE`, `Storm`, and `Verint`.
5. Wait for **Setup ready**.
6. Double-click `WFMHub.cmd`.

Your source files stay where they are. WFMHub does not edit them.

## Who counts as an agent

The FTE roster is the authority:

- Status `Active`: included.
- Status `Leaver`: included up to and including `End date if leaver`.
- After the leave date: excluded.
- Any other status, or a Leaver without a leave date: excluded.

This check uses the date of each schedule, LILO, Agent Status, or Call by Call
row. Historical work before a person's leave date therefore remains valid.

The standard `FTE Count.xlsx` template also has `PTO` and `Away` registers.
Populate them using the dropdowns and inclusive dates. Approved PTO and
Active/Closed Away are removed from expected-work minutes, attendance calls,
correction gaps, and net staffing. Planned Away changes future staffing only;
it never hides a current or historical no-show. Partial-day PTO uses its exact
start/end time. Reviewed Attendance decisions plus these registers drive the
Hub absence/shrinkage ledger.

## The normal daily routine

1. Add the new extracts to their normal folders.
2. Double-click `WFMHub.cmd`.
3. Read the status panel and latest data date.
4. Choose **Refresh source data once**.
5. Select All, Attendance, Service, or PCS sources.
6. Choose the dates.
7. Wait for **Refresh complete**.
8. Choose the individual report you need.
9. Open it from `Reports`.

You do not need to refresh the sources again for every workbook. After one
refresh, the reports use the same prepared database.

## The Reports folder

Current workbooks always have the same names:

```text
Reports\Staffing Gaps.xlsx
Reports\RTM Daily Control.xlsx
Reports\Realisations.xlsx
Reports\Attendance Review.xlsx
Reports\Final Absenteeism.xlsx
Reports\Bonus Management.xlsx
Reports\PCS Performance.xlsx
Reports\Analysis\...xlsx
```

When a report is replaced, WFMHub first saves its previous version in
`Reports\Archive`. You do not need to rename or move daily reports yourself.

`Feed` is separate from `Reports`: each successful refresh updates the fixed
PCS and Absenteeism CSV feeds there. Any other clean CSV/XLSX export appears
only when you explicitly request it.

## Choosing dates

The menu offers Today, Yesterday, Current Week, Current Month, Previous Month,
custom dates, all available dates, or saved default dates.

The selected dates affect only the calculation and report. They do not affect
the extracts. A multi-day extract is valid because WFMHub reads dates from its
rows instead of assuming the filename contains one day.

## RTM attendance call list

Open `Reports\RTM Daily Control.xlsx`, choose the LOB sheet, and scroll below
the hourly service table. This is the list used to call or follow up agents for
that LOB.

- **Call no-show** requires a completed working shift and positive evidence:
  either a loaded blank LILO row or enough Agent Status coverage showing the
  agent remained Logged Off.
- **Call late** means the first observed evidence is after scheduled start plus
  the configured tolerance.
- **Not seen now** is a provisional same-day warning.
- An unfinished shift is never marked as early leave.
- A missing source is **missing evidence**, never a no-show.

Agent Status is the detailed evidence. A temporary Logged Off or Unavailable
interval becomes an internal gap. If the agent returns later, that return is
kept and the case is not labelled early leave. LILO is used only to complete
the outer login/logout boundaries when Agent Status coverage is too sparse.

RTM is intentionally a same-day control. Use Attendance Review for completed
historical days or a whole current-week review.

## Staffing and Capacity Plan

This report uses every date you select. Completed intervals compare scheduled,
observed, and productive FTE. Future intervals compare Verint required FTE with
net scheduled FTE after approved PTO and Away. `WEEKLY_PLAN` summarizes the
same capacity as FTE-hours by week, LOB, and language. Missing forecast remains
`NO FORECAST`; forecast demand with nobody scheduled remains a visible gap.

## RTM Daily Control

This one workbook contains four daily sheets: RSA Netherlands, RSA Belgium,
Ford Netherlands, and Ford OEM France. Choose the end date you want to send;
that date is the visible Flash day. The selected start date remains in the
audit boundary but does not turn a daily Flash into a multi-day total.

Call-by-Call supplies the actual service figures. The exact Flash queue list
decides which queue entries belong to each Flash. A transfer entering another
displayed queue is a second queue entry, matching Storm Total Entered. An
abandoned call is still demand even though it has no Agent ID. Verint supplies
forecast only.

Open `CONTROL`, then click a LOB name. Each LOB shows hourly forecast, actual,
absolute volume variance, TSL, Routed Rate, weighted AHT, and `ABS HC` through
the latest actual hour. The attendance summary and named call/late list are
directly below that hourly table, so there is no second callout workbook to
reconcile. Open `ISSUES & DRIVERS` only when you need actionable missing data
or the real-volume queues pulling TSL below target.

Service availability is `handled / offered`. It is not agent availability.
The visible `Variance` is `actual offered - forecast`.

From September 2026 onward, export Verint Forecast at 15-minute grain. Put the
files in the same configured `Verint\Forecast` folder; do not edit them. Names
such as `RSA_NL_09-2026.txt`, `RSA_BE_09-2026.txt`, `FORD_NL_09-2026.txt`, and
`FORD_FR_09-2026.txt` are recognized automatically. WFMHub keeps those quarters
for Staffing and builds the hourly Flash forecast itself.

The old Flash also contained manually sourced back-office counters. Until a
reliable source is configured, WFMHub shows `NOT_CONFIGURED` instead of making
up a number.

## Realisations

The normal report includes every configured management LOB in one workbook.
`LOB_RESULTS` has one LOB/day row with actual and forecast volume, the LOB's
configured service level, service availability, weighted AHT, staffing,
absence, and shrinkage. `TREND` summarizes the same counters by month, ISO week,
and quarter. Adherence is not included.

## Attendance Review

Use this after the operating day is complete. Current Week includes every
completed date from Monday through yesterday—not just yesterday. Today is
excluded so an unfinished shift can never become an early-leave correction.
WFMHub compares schedule with Agent Status first and uses LILO as fallback and
control evidence.

1. Open `REVIEW BOARD`.
2. Each row is one exact continuous interval. White evidence cells are not
   imported; edit only the five blue decision cells.
3. Choose `Approved` plus a category, `Dismissed`, or leave `Open`.
4. Read the colored full-shift timeline on that same row: green is Logged,
   amber is Break/Lunch, and red is Gap. Exact support evidence is retained in
   hidden sheets for audit but is not part of the normal workflow.
5. Save and close the workbook.
6. In WFMHub choose **Attendance Review > Import completed decisions**.

## Final Absenteeism

This report uses imported Attendance Review decisions plus PTO/Away registers.
`PENDING_REVIEW` means at least one exact gap is still Open. `PROVISIONAL_DAY`
means the shift is not complete. Both remain incomplete and are never allowed
to dilute the headline rate as silent zero absence.

These rows are never allowed to dilute the headline rate as silent zero absence.

For a long-lived team workbook, use `TEAM_VIEW` for filtered agent results and
cases, `COMPONENT_VIEW` for absence/shrinkage categories, and
`ACTIVITY_DETAIL` for exact reviewed intervals. The blue `ACTIONS` table is
your permanent log. Link the three fixed feeds once with Power Query and use
**Data > Refresh All**; never point Power Query at `ACTIONS`.

## Bonus Management

Choose Bonus Management and then:

- import a new Bonus Matrix v1.2 and build; or
- build from the matrix already imported.

WFMHub hashes and reads the source without changing it. Scenario Payout remains
separate from Released Payout. Absence must have one financial consequence—not
an absence KPI plus a second hidden penalty.

## PCS Performance

Choose **PCS Performance**, select the dates, and open the workbook. The first
build gives you a complete working file. A one-time Power Query link to the
fixed PCS feed is optional when the workbook will stay shared for a long time.

On `DASHBOARD`, choose Latest day, Current/Previous week, Current MTD,
Previous-month same days, Previous full month, or Custom period. Then choose
LOB, Team Leader, or Agent. The KPI cards, benchmark, and chart all recalculate
from the included `PCS_DATA` table.

Open `TEAM_VIEW` for the easiest TL workflow. Filter its normal Excel Table by
LOB, Team Leader, or Agent. Latest day, current week, current MTD, previous MTD,
movement, participation, sample, and coaching priority are already calculated.
It does not use fragile spill formulas. `AGENT_RESULTS` remains the custom-period
formula view.

PCS formulas:

- PCS Average = valid Q1 score sum / valid Q1 response count.
- Participation = inbound nonblank Q1 / inbound `PCSStatus=1`.
- Coaching opportunity = one valid inbound Q1 response `<= 3`.
- Actions Rate = completed coaching / coaching opportunities.

Never average agent PCS percentages or use the raw score sum as the score.

For coaching, use `COACHING_QUEUE` to see all low-score opportunities. Work in
`COACHING` and fill the five blue columns: status, coach, coaching date, due
date, and comment. Save the workbook normally. When the Hub rebuilds the same
closed workbook, those cells carry forward by Coaching Key.

The stable input is `Feed\PCS\PCS_AGENT_DAY_CURRENT.csv`. Agent ID is the real
matching key; Agent Selector is only the friendly `Name [ID]` label. To keep one
shared workbook permanently, connect that CSV once to `tblPcsData` with Power
Query and use **Data > Refresh All** after each Hub refresh. Do not regenerate
the PCS workbook while colleagues are editing it.

Follow [EXCEL_REFRESH_GUIDE.md](EXCEL_REFRESH_GUIDE.md) for every click in the
one-time PCS and Absenteeism setup.

If you prefer native slicers, click inside `PCS_DATA` or `COACHING` and choose
**Table Design > Insert Slicer**.

## Analysis and clean data

**Analyze a period** creates a separate evidence workbook under
`Reports\Analysis` for a selected domain and comparison. Every observation
points to its source metric and evidence.

**Export clean data** creates CSV or XLSX for the selected dataset and dates.
Use CSV for large Call by Call data.

## If something breaks

- **File not found:** keep the complete extracted WFMHub folder together.
- **Another refresh is running:** let it finish; do not delete its lock.
- **No Time zone found:** the portable package was incompletely extracted.
- **Required columns missing:** read the named source file in the error/log.
- **Empty report:** check source health and date coverage in System tools.
- **Red INCOMPLETE badge:** fix the missing or stale source; do not replace the
  blank with zero.

Technical data, logs, and backups are under `_system`. Finished workbooks belong
in `Reports`; requested clean exports belong in `Feed`.
