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
start/end time. Pending or Cancelled records have no calculation effect.
Reviewed Attendance decisions plus these registers drive the Hub
absence/shrinkage ledger.

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
Reports\PCS Operational Report - YYYY-MM-DD HHMMSS.xlsx
Reports\PCS Coaching Log.xlsx
Reports\Analysis\...xlsx
```

When a normal fixed-name report is replaced, WFMHub first saves its previous
version in `Reports\Archive`. PCS reports are timestamped snapshots, so the Hub
never replaces an open file. `PCS Coaching Log.xlsx` is the only permanent
human-edited PCS file and the Hub never overwrites it.

`Feed` is separate from `Reports`: shared Absenteeism feeds live there. PCS no
longer uses feed files or Power Query. Any other clean CSV/XLSX export appears
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
`No Show HC` means the agent never showed presence and Agent Status explicitly
supports a no-show. Late agents, early leavers, and agents who went offline
after working remain present. Offline Now is a separate alert. If the file has
no reliable state for that person, RTM shows `UNKNOWN — POSSIBLE NO SHOW`; check
the data before calling.

The summary balances as `Due HC = Present HC + No Show HC + Unknown HC`.
Offline Now is already included inside Present HC, so never add it again.
`PTO / Away HC` is separate context: full-day leave is shown in the agent list
but not counted as Due, Unknown, No Show, or Call Now. For partial-day PTO, RTM
starts judging attendance only when the next real working interval begins.

The attendance checkpoint is the newest Agent Status evidence time, not the
time the workbook finishes building. A last known state older than
`rules.rta_stale_minutes` is shown as `UNKNOWN`.

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

Open `CONTROL`, then click a LOB name. Each LOB starts with TSL, offered volume,
absolute volume variance and confirmed `No Show HC`. Its hourly table retains
forecast, Routed Rate and weighted AHT through the latest actual hour. The
attendance summary and named call/late list are
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
2. Each case uses two rows: SCHEDULE directly above ACTUAL, followed by a small
   blank separator before the next case. White evidence cells are not imported;
   edit only the five blue cells on the ACTUAL row.
3. Choose `Approved` plus a category, `Dismissed`, or leave `Open`.
4. Compare the paired timeline. Scheduled work is teal, PTO/Away is blue,
   Logged is green, Break/Lunch is amber, dark red is the exact gap owned by the
   ACTUAL row, light red is another counted gap, and grey is inside tolerance.
   Exact support evidence remains in hidden sheets.
   Explicit `Meal Aux` Agent Status appears as Lunch. A LILO logout inside that
   interval does not turn the meal into a gap; any later gap starts only after
   the meal interval ends.
5. Save and close the workbook.
6. In WFMHub choose **Attendance Review > Import completed decisions**.

Open `BREAK & MEAL` to check completed shifts. It totals all Agent Status break
and meal spells inside the shift, compares them with the values in
`config\wfmhub.toml`, and places overruns first. `INSUFFICIENT EVIDENCE` means
the Hub deliberately refused to judge that agent-day.

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

## PCS Report & Coaching

Choose **PCS Report & Coaching**, then **Build latest PCS report**. You do not
need to close Excel. The Hub loads only FTE and Call by Call and creates a new
timestamped report containing final values and charts. Attendance, RTM,
staffing, service, forecast and absence remain unchanged.

Open `OVERVIEW` for the management view. Use the four dropdowns from left to
right: Period, LOB, Team Leader, Agent. Changing LOB narrows the Team Leader and
Agent choices; changing Team Leader narrows Agent choices. Cards, charts and the
eight-row performance panel follow the selection. If you change a parent after
choosing a child, reset the child to `All` and continue left to right.

Open `RESULTS` for the complete result list and use its ordinary filter arrows:
Period View, Scope Level, LOB, Team Leader, then Agent. `LOB` gives one row per
LOB, `TEAM` gives team rows, and `AGENT` gives agent rows. No data refresh is
required in Excel.

PCS formulas:

- PCS Average = valid Q1 score sum / valid Q1 response count.
- Participation = inbound nonblank Q1 / inbound `PCSStatus=1`.
- Coaching opportunity = one valid inbound Q1 response `<= 3`.
- Actions Rate = completed coaching / coaching opportunities.

Never average agent PCS percentages or use the raw score sum as the score.

For coaching, filter `COACHING_QUEUE` by LOB, Team Leader, Agent or date. Copy
columns A:M for the selected call and paste them into the first empty row of
`PCS Coaching Log.xlsx`. Complete Coaching Status, Coach, Coaching Date, Due
Date and Coaching Comment, then save. Coaching Key identifies the exact call;
Call ID helps the coach open it directly.

After Quality updates the log, choose **Build again from current database**.
That fast action does not scan the source folder. It creates a new report with
the latest coaching status and leaves both earlier reports and the permanent log
untouched.

If you prefer native slicers, add them to `RESULTS`, `PCS_DATA`, or
`COACHING_QUEUE`. They filter static Excel Tables; they do not trigger a refresh.

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
- **PCS report does not build:** confirm FTE and Call by Call contain in-scope
  dates and agents, then read the named error in the latest log. No Excel file
  needs to be closed for PCS.
- **Red INCOMPLETE badge:** fix the missing or stale source; do not replace the
  blank with zero.

Technical data, logs, and backups are under `_system`. Finished workbooks belong
in `Reports`; requested clean exports belong in `Feed`.
