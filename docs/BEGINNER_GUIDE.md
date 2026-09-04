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
Populate them now using the dropdowns and inclusive dates, but note that
v0.13.1 does not yet apply those registers to attendance or staffing. Until
that overlay is released, they are controlled planning inputs rather than Hub
calculation evidence.

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
Reports\Attendance Callout.xlsx
Reports\Staffing Gaps.xlsx
Reports\OEM Flash.xlsx
Reports\Attendance Review.xlsx
Reports\Final Absenteeism.xlsx
Reports\Bonus Management.xlsx
Reports\PCS Performance.xlsx
```

When a generated report is replaced, WFMHub first saves its previous version in
`Reports\Archive`. You do not need to rename or move daily reports yourself.

`Feed` is separate from `Reports`: it receives only the clean CSV/XLSX exports
you explicitly request.

## Choosing dates

The menu offers Today, Yesterday, Current Week, Current Month, Previous Month,
custom dates, all available dates, or saved default dates.

The selected dates affect only the calculation and report. They do not affect
the extracts. A multi-day extract is valid because WFMHub reads dates from its
rows instead of assuming the filename contains one day.

## Attendance Callout

This is the list used to call or follow up agents.

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

Choose Current Week to receive every actionable case in the week, not only
yesterday.

## Staffing Gaps

This report compares scheduled, observed, and productive FTE by 15-minute
interval, LOB, and language. Future or missing intervals remain unknown instead
of becoming fake shortages.

## OEM Flash

The Ford OEM pilot combines:

- OEM/Ford/Toyota service level and service availability;
- actual calls, Verint forecast, forecast attainment, and call variance;
- hourly Ford/Toyota/Chery queue-group results;
- scheduled, observed, and productive FTE when matching evidence exists;
- queue mapping and source freshness.

Service availability is `answered / offered`. It is not agent availability.

Forecast attainment is `actual offered / forecast`. Call variance is
`actual offered - forecast`; these two ideas are deliberately shown separately.

The old Flash also contained manually sourced back-office counters. Until a
reliable source is configured, WFMHub shows `NOT_CONFIGURED` instead of making
up a number.

## Attendance Review

Use this after the operating day is complete. Current Week includes every
completed date from Monday through yesterday—not just yesterday. Today is
excluded so an unfinished shift can never become an early-leave correction.
WFMHub compares schedule with LILO and Agent Status, then checks whether Verint
Activities already cover each gap.

1. Open `VERINT_INJECTION`.
2. Each row is one exact continuous interval; use **Start to inject** and
   **End to inject** without rounding or combining rows.
3. Use `SHIFT_VIEW` only when you need to verify the complete shift evidence.
4. Correct the intervals in Verint.
5. Export Activities again and refresh WFMHub. Corrected intervals disappear
   automatically; never import this workbook back into the Hub.

## Final Absenteeism

This is the final payroll/reporting ledger. It uses corrected Verint Activities
inside StartEndTimes boundaries. LILO and Agent Status detect operational gaps;
they do not invent the final payroll category.

Resolve every unmapped activity before using the report as final.
Resolve every final-ledger exception:

- `UNCODED_EMPTY_SHIFT`: completed working shift, no final Verint code, and no
  reliable Agent Status/LILO work evidence;
- `UNCORRECTED_OBSERVED_GAP`: the operational evidence proves a gap but Verint
  has no final code;
- `PARTIAL_CORRECTION_REVIEW`: Verint has a code but it does not cover the whole
  observed gap;
- `VERINT_WITHOUT_OBSERVED_GAP`: Verint contains a final code but the operational
  evidence does not support that interval;
- `PROVISIONAL_DAY`: the shift is not complete and cannot be finalized yet.

These rows are never allowed to dilute the headline rate as silent zero absence.

## Bonus Management

Choose Bonus Management and then:

- import a new Bonus Matrix v1.2 and build; or
- build from the matrix already imported.

WFMHub hashes and reads the source without changing it. Scenario Payout remains
separate from Released Payout. Absence must have one financial consequence—not
an absence KPI plus a second hidden penalty.

## PCS Performance

Choose **PCS Performance**, select the dates, and open the generated workbook.
There is no Power Query, Data Model, or Refresh All step.

On `DASHBOARD`, use the boxes for Period View, custom dates, LOB, Team Leader,
and Agent. The KPI cards recalculate in Excel from the included `PCS_DATA`
table. The report also includes ready-made team, agent, trend, and coaching
views.

PCS formulas:

- PCS Average = valid Q1 score sum / valid Q1 response count.
- Participation = inbound nonblank Q1 / inbound `PCSStatus=1`.
- Coaching opportunity = one valid inbound Q1 response `<= 3`.
- Actions Rate = completed coaching / coaching opportunities.

Never average agent PCS percentages or use the raw score sum as the score.

For coaching, filter `COACHING` to your team and dates, then fill the four blue
columns: status, date, coach, and comment. Save/send the workbook normally.
Nothing must be synced back into WFMHub.

If you prefer native slicers, click inside `PCS_DATA` or `COACHING` and choose
**Table Design > Insert Slicer**.

## Analysis and clean data

**Analyze a period** creates a separate evidence workbook for a selected domain
and comparison. It is deterministic Python/SQLite analysis, not AI.

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
