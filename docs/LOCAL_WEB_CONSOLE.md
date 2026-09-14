# WFMHub local Manager Workbench

The Manager Workbench is the default visual surface from WFMHub `0.35.0`.
Double-click `WEBAPP.cmd`; the bundled Python runtime starts a small server on
`127.0.0.1` and opens the normal browser. Close the command window to stop it.

The approved complete design is
[`manager-workbench-v2/all-pages-review.png`](design-prototypes/manager-workbench-v2/all-pages-review.png).
Its numbers are illustrative. The installed application reads the operator's
governed SQLite database.

The application is deliberately local and single-user. Another computer cannot
reach it, it cannot run inside SharePoint, and it sends no WFM data to the
internet. The browser receives governed presentation rows only; it never opens
raw extracts, SQL or the SQLite file.

## The simple operating model

```text
new untouched extracts -> Update data -> SQLite raw/core/marts
                                      -> Manager Workbench pages
                                      -> report / analysis when requested
```

- **Update data** is the only Workbench operation that ingests new extracts and
  rebuilds governed marts. Run it after placing new files in the existing source
  folders.
- Opening a page only reads the current database. It does not rebuild anything.
- **Export detail** downloads the selected page's exact filtered register as CSV.
- **Generate report** opens Deliver, where a workbook or deterministic period
  analysis can be generated from current marts. This does not run Update first.
- Generated files are available from **Generated Archive**. Execution evidence
  is visible in **Jobs & Logs**.

There is no Power BI refresh, Excel query, ODBC connection, Node service, cloud
login or runtime AI in this workflow. PCS remains the permanent collaborative
Excel tracker and is not a Workbench page.

## Pages and when to use them

### Manager Desk

Start here. It combines three decision horizons without manufacturing a blended
KPI: people to contact now, post-day residual review, and exact capacity risks
for the next seven days. Service remains listed separately by Management LOB.

### Plan

- **Demand & Requirement** shows Verint Volume and Absolute Required FTE at
  Staff Type grain. Forecast `Queue Name` is a workforce Staff Type, not a call
  queue.
- **Capacity & Schedule** compares required FTE with gross published schedules,
  approved PTO/Away and net scheduled FTE at 15-minute grain.
- **Scenario Lab** applies one visible temporary FTE adjustment to the selected
  scope. It does not save or rewrite any schedule, extract, mapping or database
  row.

### Operate

- **Live Service** shows one selected Management LOB's service curve, target and
  exact configured queue facts. RSA BE service is combined; no portfolio-wide
  SL is invented.
- **Attendance Pulse** is the latest-day contact list. Present includes late;
  confirmed No Show and evidence-missing possible No Show remain distinct.

### Review

- **Schedule Integrity** overlays the published schedule and chronological
  Agent Status/LILO evidence. It shows exact residual gaps after final Verint
  Activities overlap and a separate break/meal register.
- **Realisations** keeps service, demand, requirement, net schedule, observed
  and productive delivery as separate questions and presents ratios of sums.
- **Absence & Shrinkage** uses finalized Verint Activities components only.
  Empty/unmapped/incomplete evidence remains in the exception register.
- **Workforce Patterns** shows configured recurring schedule-placement evidence.
  It does not infer intent or decide a management action.

### Deliver

- **Reports & Analysis** generates focused current workbooks or an on-demand
  analysis for a selected domain, period and comparison.
- **Generated Archive** finds current and timestamped prior outputs and downloads
  them through the local application.

### Govern

- **Data Readiness** shows source freshness, rejected rows and quality findings.
- **Mappings & Rules** displays effective service profiles, exact queue counts,
  capacity mappings, targets and configuration fingerprints. It is read-only.
- **Jobs & Logs** traces Update, report and analysis execution.

## Filters

Choose filters from left to right: period, Management LOB, Planning Group, Staff
Type, Team Leader and Agent. Later selectors cascade from earlier choices.
Capacity selectors are disabled on service-only pages because service queues and
workforce Staff Types are different domains. Attendance Pulse and Live Service
always return to the latest operational day; Schedule Integrity retains the full
selected range.

Missing evidence renders as blank or Unknown. Rates are calculated in Python
from summed additive components; the browser only formats them.

## If something fails

1. Open **Govern > Data Readiness** to check source dates and quality findings.
2. Open **Govern > Jobs & Logs** for the exact job result.
3. If a generated workbook is open in Excel, close it before replacing the fixed
   current file and run the report again.
4. Do not delete the database and do not edit the source extracts. A normal
   release upgrade preserves the existing SQLite history.
