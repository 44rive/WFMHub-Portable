# Beginner guide — WFMHub without technical words

Think of WFMHub as a washing machine for data:

1. You drop new extract files into the same source folders.
2. WFMHub reads them but never edits them.
3. It keeps the useful rows in its SQLite database.
4. It calculates the agreed rules.
5. It gives you one Excel workbook for one job.

You do not need to install Python, SQLite, DuckDB, ODBC, or Power BI.

## The two files you click

- `SETUP.cmd`: use once when you install WFMHub or move the extract folder.
- `WFMHub.cmd`: use for normal daily work.

Keep the entire WFMHub folder together. Do not move only the `.cmd` files.

## First setup — one time

1. Extract the complete portable ZIP.
2. Double-click `SETUP.cmd`.
3. Paste the path of the folder that contains your `FTE`, `Storm`, and `Verint`
   folders.
4. Wait for **Setup ready**.
5. Double-click `WFMHub.cmd`.
6. Choose **Validate rules and build governance catalog**.

Setup creates local config files and `database\wfm.sqlite3`. Your source files
are not moved or edited.

## Your FTE roster defines “our agents”

WFMHub uses the **Agent** sheet in the FTE workbook. Use
`templates\FTE Count.xlsx` if you need a clean standard file. The important
fields are Agent ID, Agent Name, Team Leader, Ops Manager, LOB, and Language.

An operational row is kept when its Agent ID is on the active roster, or when a
unique roster name can safely identify it. Verint **Data Source IDs** is treated
as the Verint Agent ID. Rows outside the roster are counted and excluded from
agent reports.

## Normal daily routine

1. Put today’s extracts in their normal folders. Existing files can stay.
2. Double-click `WFMHub.cmd`.
3. Read the fixed status panel. It shows the latest loaded source date.
4. Choose **Refresh hub data**.
5. Choose the source group:
   - Attendance/absence for FTE, schedule, LILO, and Agent Status.
   - Service for APBE/APFR/APDE and Verint Forecast.
   - PCS for FTE and Call by Call.
   - All sources when you want everything.
6. Choose a date option.
7. Choose the exact report products you want.
8. Wait until the progress bar says **Refresh complete**.
9. Open the new workbook under `output`.

The progress bar can pause on a large file while rows are still being streamed.
Do not close the window unless an error is shown. The latest details are in
`logs\wfmhub_YYYYMMDD.log`.

## Choose dates

The menu offers today, yesterday, current week, current month, previous month,
a custom range, config dates, or all available dates.

For a monthly exercise choose **Current month** or enter the first and last date
yourself. Date selection never changes an extract. It only controls what the
model and report include.

Multi-day files are okay. WFMHub uses dates inside each row or Verint date
column—not just the filename.

## The seven report products

### PCS Performance

Use this for management PCS updates. It shows the latest day, selected period,
current MTD, comparable previous-month days, and previous full month.

- PCS average = sum of valid inbound Q1 scores / count of valid inbound Q1.
- Participation = inbound nonblank Q1 / inbound `PCSStatus=1`.
- A three-hour rhythm means you generate/send it every three hours. The formula
  and reporting period are still daily and monthly.

Never average agent percentages. WFMHub divides the summed counters.

### Bonus Performance

First choose **Import Bonus Matrix v1.2** and paste the path to the final source
workbook. WFMHub reads and hashes it without changing it.

The report separates:

- Scenario Payout: useful for management discussion.
- Released Payout: usable only after policy, eligibility, and data gates pass.
- Source Workbook Payout: the cached source result used for reconciliation.
- Control Adjustment: the explained difference between governed and source
  scenarios.

Absence must be used once—as a KPI, eligibility gate, or proration—not twice.

### Service Performance

Choose a service profile. Ford OEM France is the first configured example. Its
queues come from `config\queue_mapping.csv`; its scope, governed metric choices,
and display groups come from `config\service_profiles.toml`. KPI formulas and
targets remain in `config\metric_catalog.toml`.

- Ford OEM Service Level matches the original Flash: summed
  answered-within-target / summed offered. Another profile may select another
  governed method without changing Python.
- Service Availability = answered / offered.
- AHT = handled seconds / answered.
- Forecast is used only as forecast. APBE/APFR/APDE are actuals.

Operational forecast files must be placed in the configured
`Verint\Forecast` source folder. Files under `TOLEARN` are references and are
never loaded automatically. Names such as `FORD_FR_08-2026.txt` and
`Forecast_RSA_NL_August.txt` are matched to the configured filename token in
`queue_mapping.csv`.

Service availability is never agent availability and never adherence.

### Staffing & Coverage

Use this to see scheduled, observed, and productive FTE by 15-minute interval,
LOB, and language. Future or missing evidence is left unknown; it is not turned
into a fake staffing gap.

### Attendance Callouts

Choose **Today** for the live calling list. Choose **Current Week** or custom
dates to include every callout case and daily trend in that whole period.

An unfinished shift cannot be marked as early leave. A confirmed no-show needs
a completed scheduled shift, a loaded blank LILO row, and no active Agent
Status evidence. A missing file or missing row is **missing evidence**, not a
no-show.

### Attendance Corrections

Use this after the day is complete. It compares planned schedule boundaries
with LILO and Agent Status evidence, then checks whether Verint Activities
already cover the gap.

1. Open `GAPS`.
2. Edit only the pale-blue columns: Confirmed Activity, Validation Status,
   Owner, Comment, and Injected Date.
3. Use `SHIFT_VIEW` to see planned versus observed time across the shift.
4. Save the workbook.
5. In WFMHub choose **Import attendance correction decisions**.

One invalid row cancels the complete import, so partial bad changes are not
saved.

### Final Absence & Shrinkage

This is the post-correction result for payroll/final reporting. It uses Verint
Activities clipped to StartEndTimes boundaries. LILO and Agent Status detect
the original operational problem, but do not invent the final payroll reason.
Review every unmapped activity before using the result.

## Run an analysis for any period

Choose **Analyze a period**, then select PCS, service, forecast, staffing,
attendance, final absence, or bonus. Choose a comparison:

- previous equal-length period;
- previous-month same calendar days;
- configured target; or
- no comparison.

The workbook contains `FINDINGS`, `METRICS`, and the exact `EVIDENCE` rows. It
uses deterministic Python/SQLite rules, not AI. A comparison describes what
changed; it does not pretend to prove why.

## Export clean data

Choose **Export clean data**, a dataset, dates, and CSV or XLSX. Use CSV for
large call data. WFMHub writes the result to `output\data_exports` and creates a
manifest beside it. The extract is untouched.

## Optional Excel PivotTables and slicers

You do not need this for normal reports. Use it only when you want an
interactive Excel master.

1. Choose **Create an Excel Pivot/slicer master**.
2. Select one report and a useful populated period.
3. Open the created file in `templates\reports`.
4. Follow [EXCEL_TEMPLATE_GUIDE.md](EXCEL_TEMPLATE_GUIDE.md) once.
5. Save the master.

After that, normal report builds refresh only the small CSV tables under
`output\model_data`. Open the master, click **Data > Refresh All**, then use
**Save As** for the file you will email. WFMHub never overwrites the master, so
your PivotTables and slicers stay intact. Raw data is not loaded into sheets.

## Change a KPI safely

Do not edit formulas inside a report workbook. They are presentation files.

- Change KPI formulas/targets in `config\metric_catalog.toml`.
- Change attendance/absence classifications in `config\wfm_rules.toml`.
- Change service queue membership in `config\queue_mapping.csv`.
- Change service report profiles in `config\service_profiles.toml`.
- Change analysis thresholds in `config\analytics_rules.toml`.

Increase the file version, validate rules, refresh, and compare the new report
with the previous one. Read [METRIC_CATALOG_GUIDE.md](METRIC_CATALOG_GUIDE.md)
before changing a formula.

## Optional Copilot writing help

WFMHub does not call any AI or upload data. If your company permits it, attach
only a finished report to your approved Microsoft Copilot and paste
`prompts\COPILOT_WFM_ANALYST.md`. Copilot may help explain the result; it is not
allowed to recalculate the KPI or replace the evidence.

## Custom Python or SQL

Copy an underscore example in `custom\jobs` or `custom\sql`, rename it, and use
the **Advanced** menu. Custom jobs receive selected dates and read-only hub
access. Python code is still executable code: use only code you understand.

## Backups and common errors

- Use **Create database backup** before a major configuration or upgrade.
- “Another refresh is running” means wait for the current job. Do not delete a
  lock while that job is active.
- “No Time zone found” means the portable release is incomplete; keep the whole
  extracted folder and run the system check.
- “Required columns are missing” means the selected file does not match that
  source contract. Read the named file/columns in the log.
- An empty report usually means the selected period has no matching scoped data.
  Choose **Show source health and date coverage**.
- A red **INCOMPLETE** badge means do not silently replace blanks with zero.
