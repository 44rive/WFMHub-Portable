# WFMHub Portable

WFMHub reads untouched WFM extracts, keeps a durable SQLite history, and
produces focused Excel workbooks. It runs on a locked-down Windows work machine
without admin rights, installed Python, ODBC, DuckDB, or Python in Excel.

The operating rule is simple: **refresh the sources once, then build only the
product you need**. Extract files are never moved or edited.

## What you see

Normal users work from three obvious folders:

```text
WFMHub/
  Reports/     finished Excel reports, including on-demand analysis
  Feed/        fixed shared-report feeds plus clean exports you request
  config/      business rules, KPI methods, queue maps, and source settings
  _system/     runtime, database, logs, backups, code, and documentation
```

Every report has a fixed name. When WFMHub replaces one, it first
copies the previous version into `Reports\Archive\YYYY-MM-DD`.

Technical files live under `_system`. You normally do not open that folder.

## Products

| Product | Operational decision |
|---|---|
| RTM Daily Control | What is the live service state, and who needs attendance follow-up in each LOB? |
| Staffing & Capacity Plan | Where is capacity missing now, and where will forecast demand exceed net schedules in future weeks? |
| Realisations | How did actual volume, service, forecast, staffing, absence, and shrinkage perform across every mapped LOB and period? |
| Attendance Review | Which exact completed-day gaps need an Approved or Dismissed human decision? |
| Final Absenteeism | What has Verint finally coded for absence/shrinkage, and which shifts remain incomplete? |
| Bonus Management | What did Bonus Matrix v1.2 calculate, and is it safe to release? |
| PCS Report & Coaching | How are PCS, participation, low scores, and coaching moving by date, month, LOB, team, and agent? |

The products use one visual identity but not one generic layout. RTM combines
service and its matching LOB attendance list, Attendance Review is an exact-gap
decision board, and Final Absenteeism is a ledger.

Adherence is not calculated. Reported Routed Rate means **total routed / total
entered**. Reported TSL means **connected within 30 seconds / (lost + connected
- lost from 5 to 30 seconds)**. Both follow the supplied Storm configuration; Routed
Rate never means agent availability.

## Source authority

| Source | Used for |
|---|---|
| FTE roster | Date-aware in-scope agents, organisation fields, approved PTO, and Away planning |
| Verint StartEndTimes | Preferred scheduled start/end and assignment boundaries |
| Storm LILO | First/last presence evidence, including loaded blank rows |
| Storm Agent Status | Observed attendance and interval staffing evidence |
| Verint Activities | Final post-day absence/shrinkage codes only |
| Verint Forecast | Forecast only |
| Storm Call by Call | All mapped service actuals, Flash demand/service, agent call performance, and PCS |
| Bonus Matrix v1.2 | Bonus inputs, KPI configuration, and source reconciliation |

Multi-day files are supported. Row dates are authoritative; filename dates are
only fallback hints. Missing evidence remains missing and is never converted to
a false no-show or zero.

If a dedicated StartEndTimes file is unavailable, WFMHub can use a successfully
parsed Shift Assignment boundary from an Activities export and raises a visible
review finding. Activity intervals are never used as attendance or presence;
they feed only the separate final post-day absence/shrinkage ledger.

FTE scope is effective-dated: `Active` rows are admitted; `Leaver` rows are
admitted only through `End date if leaver`; other statuses and undated leavers
are excluded. The same rule is applied to schedules, LILO, Agent Status, and
Call by Call for each row's business date.

RTM workforce ownership follows the roster labels exactly: OEM uses `OEM FR`,
RSA Belgium combines `RSA FR` and `RSA VL`, Ford Netherlands uses `Ford Dutch`,
and RSA Netherlands uses `RSA NL`.

The standard FTE workbook also owns PTO and Away registers. Approved PTO and
effective Away intervals change expected work and net staffing without editing
any extract. Pending/Cancelled entries do not change the calculations. Active
and Closed Away apply inside their effective dates; Planned Away affects future
capacity only and never erases past or current attendance evidence. Exact
attendance gaps are classified through the imported Attendance Review ledger.

## Windows quick start

1. Download the portable release and choose **Extract All**.
2. Double-click `SETUP.cmd` once. For a later release extracted into a different
   folder, use `UPGRADE.cmd` and point it to the previous WFMHub folder.
3. Paste the folder containing `FTE`, `Storm`, and `Verint`.
4. Double-click `WFMHub.cmd`.
5. Choose **Refresh source data once**.
6. Choose a reliable product under **Operational**, or an unfinished product
   under **In Development**.
7. Open it directly from `Reports`.

See the [beginner guide](docs/BEGINNER_GUIDE.md) for the normal routine and the
[Excel refresh guide](docs/EXCEL_REFRESH_GUIDE.md) for the one-time PCS Power
Query installation and the optional shared-Absenteeism setup.
The [Power BI direction](docs/POWERBI_WFMHUB_PERSPECTIVE.md) and
[premium page specification](docs/POWERBI_PREMIUM_DESIGN_V1.md) describe the
governed `Feed\PowerBI` model and seven approved pages.

The approved visual contract is in the
[report design system](docs/REPORT_DESIGN_SYSTEM.md), with a data-free
[Excel design reference](<docs/WFMHub Report Design Reference.xlsx>). Future
developers and assistants must start with [AI_CONTEXT.md](AI_CONTEXT.md).

## RTM Daily Control

`Reports\RTM Daily Control.xlsx` is the single same-day operating file. It
implements the approved four-LOB Flash layout as native Excel
sheets for RSA NL, RSA BE, Ford NL, and Ford OEM France. `CONTROL` gives the
cross-LOB state. Each LOB sheet contains its hourly service table and its own
reconciling attendance/call list. `ISSUES & DRIVERS` contains only actionable
source issues and below-target queues with real demand. Definitions and audit
evidence remain in hidden support sheets.

Actual demand comes from exact queues in `queue_mapping.csv`. WFMHub counts each
mapped inbound queue entry, matching Storm `Total Entered`; outbound companion
legs are ignored. A transfer entering another displayed queue is therefore a
new queue entry. Each Flash uses an exact reviewed queue allowlist: RSA NL has
the 23 queues marked `SL Related = Y` in the supplied reference, RSA BE has 43
queues, Ford NL has six queues, and Ford FR/OEM has the exact Ford Assistance,
Toyota/Lexus and Chery Assistance queues. PCS stays limited to the effective-
dated FTE roster.

Verint Forecast supplies forecast only. New 15-minute exports stay at their
native grain for Staffing and are rolled into hours for RTM. Hourly
volume is the sum of the four quarters; FTE is an average level; forecast SL
and AHT are weighted by volume. Routed Rate is
`total routed / total entered`; TSL is `connected within 30 seconds / (total
entered - lost calls from 5 to 30 seconds)`; AHT is weighted handled seconds per
routed queue entry. The threshold clock is total queue wait plus ringing.

Every LOB displays the full 00:00-23:00 day. Headline controls show TSL, offered
volume, volume variance, and `No Show HC`; the hourly and LOB tables also retain
Forecast, Volume Handled, Handled in SL, Routed Rate and AHT, while the
attendance section retains Offline Now and Call Now. `No Show HC`
means the agent has no observed presence at all and agent-specific evidence
proves the no-show. Late agents, early leavers, and agents who went offline
after attending remain present. `Possible No Show HC` is an unknown/no-evidence
case and is never added to No Show HC or the automatic call count.
The attendance strip reconciles `Due HC = Present HC + No Show HC + Unknown HC`;
Offline Now is a subset of Present HC. `PTO / Away HC` shows scheduled people
with registered leave that day. Full-day leave is visible but excluded from Due
HC; partial-day leave removes only its exact interval, so the remaining working
time is still controlled normally.

The latest call hour and attendance checkpoint are independent. Live attendance
uses the latest Agent Status evidence time, not the later workbook refresh time.
Presence and no-show require evidence for that specific agent at the checkpoint;
a file merely existing for the date is not enough. A last known state is
accepted only inside `rules.rta_stale_minutes`. Missing or stale per-agent
evidence is `UNKNOWN — POSSIBLE NO SHOW` and requires a data check.

Queue membership lives in `config\queue_mapping.csv`; profile scope lives in
`config\service_profiles.toml`; formulas and targets live in
`config\metric_catalog.toml`.

Travel-labelled active roster LOBs already flow into Attendance, Staffing,
Absence, PCS and Bonus whenever those sources contain matching rows. WFMHub does
not fabricate a Travel service Flash or SL from the LOB name: that requires an
exact reviewed Travel queue list and forecast-file mapping.

## PCS Report & Coaching

PCS is one permanent collaborative workbook backed by six fixed-name
CSV feeds. It contains no raw Call-by-Call worksheet. Python/SQLite calculate
the governed scorecards first; Power Query only transports those finished
results into Excel.

Run **PCS Report & Coaching > Install/repair Power Query** once after this
upgrade. Normal use is then:

1. Choose **Update latest PCS data**. The Hub loads only FTE and Call by Call,
   updates the PCS mart, and atomically replaces the CSV feeds under
   `Feed\PCS`.
2. Open `Reports\PCS Live Tracker.xlsx` and choose **Data > Refresh All**.
3. On `OVERVIEW`, choose Period, LOB, Team Leader and Agent from left to right.
   If a parent changes, reset its child selections to `All`. Cards, both charts,
   the LOB table and agent table respond together.
4. On `COACHING`, choose Period and LOB. Use the filter arrows on `PERFORMANCE`
   for detailed checks or add standard table slicers there if useful.

`OVERVIEW` keeps four selection-aware cards, the current/prior LOB comparison,
the selected-period daily trend, and a compact agent scorecard below it. `PERFORMANCE` is the full
native-filter/slicer-ready result table across standard periods and grains.
`COACHING` places the exact low-score call queue beside one permanent editable
action table. Copy Coaching Key and Call ID across before recording the action.

The Hub never needs to open Excel during a normal PCS update and never rewrites
the same-version tracker. A versioned design migration archives the old copy
and carries keyed coaching actions forward. There is no Data Model, Power
Pivot, macro, raw-data sheet, or dynamic-array formula dependency.

## Attendance decisions and shared absenteeism

`Reports\Attendance Review.xlsx` is the auditable decision input. Build it for
the required completed dates, edit only the five blue columns on `REVIEW BOARD`,
save it, then choose **Attendance Review > Import completed decisions**. Gap ID
anchors the exact immutable start/end interval in SQLite; edited evidence is
never trusted. Approved rows use the selected rulebook category, Dismissed rows
count as no loss, and Open rows stay unverified. Every case has a SCHEDULE band
directly above its ACTUAL band, so shift boundaries and PTO/Away can be compared
with Logged, Break, Lunch, Gap and Unknown evidence. Decisions are edited only
on the ACTUAL row. The stored decision ledger and exact source evidence remain
hidden by default because normal reviewers do not need to operate those sheets.
`BREAK & MEAL` totals completed-day Agent Status intervals per agent, compares
them with the configurable break and meal allowances, and raises an overrun
only when source coverage is sufficient.

`Reports\Final Absenteeism.xlsx` follows the same long-lived-file principle and
uses final mapped Verint Activities, clipped to StartEndTimes shifts.
`TEAM_VIEW` filters agent results and review cases; `COMPONENT_VIEW` explains
absence and shrinkage by reviewed category; `ACTIVITY_DETAIL` holds exact
start/end evidence. The blue `ACTIONS` table is the permanent team-owned log
and is never a Power Query target. Link the three fixed Absenteeism feeds once,
then use **Data > Refresh All** without regenerating the shared workbook.

## Staffing and Realisations

Staffing covers the whole selected range, not only its last day. Past intervals
use observed attendance evidence; future intervals compare Verint required FTE
with gross schedules minus approved PTO and effective Away. Forecast demand
with no scheduled roster row is still shown as a gap.

The normal Realisations command produces one workbook for every active service
profile. Queue/service scopes and their matching roster LOBs are explicit
configuration, so service and staffing are never joined by a guessed name.

## Configurable logic

| File | Owns |
|---|---|
| `config\wfm_rules.toml` | Attendance and absence evidence classification |
| `config\metric_catalog.toml` | Effective-dated KPI formulas, targets, units, and aggregation |
| `config\analytics_rules.toml` | Period-analysis thresholds |
| `config\queue_mapping.csv` | Queue-to-LOB and forecast comparison mapping |
| `config\service_profiles.toml` | Flash/service scope, roster-LOB links, and queue groups |

Percentages are stored as decimals: `0.80` means 80%. Aggregated percentages
use ratios of summed components; WFMHub never averages agent percentages.

## Analysis and clean exports

**Analyze a period** creates a visible workbook under `Reports\Analysis` for PCS, service,
forecast, staffing, attendance, final absence, or bonus. Every finding includes
its metric, comparison, and evidence filter.

Attendance/absence refreshes update the fixed Absenteeism CSV feeds under
`Feed`. PCS updates its own six governed fixed CSV feeds under `Feed\PCS`.
Every complete update also replaces the fixed star-model feeds under
`Feed\PowerBI`; Power BI only refreshes those governed outputs.
**Export clean data** produces any additional CSV or
XLSX dataset you request for a selected period. Large call datasets should use
CSV. The original extract is unchanged.

The optional prompt under `prompts` can be used manually with an approved
Microsoft Copilot account after attaching only an approved finished report.

## Developer commands

```bash
python3 -m venv .venv
. .venv/bin/activate
python3 -m pip install -e .
python3 -m wfmhub --home . setup --source-root /path/to/extracts --non-interactive
python3 -m wfmhub --home . refresh --start 2026-08-01 --end 2026-08-31 --no-report
python3 -m wfmhub --home . report --pack service --start 2026-08-31 --end 2026-08-31
python3 -m wfmhub --home . report --pack pcs --start 2026-08-01 --end 2026-08-31
python3 -m wfmhub --home . import-attendance-decisions "Reports/Attendance Review.xlsx"
python3 -m wfmhub --home . analyze pcs --start 2026-08-01 --end 2026-08-31 --comparison previous_month
python3 -m unittest discover -s tests -v
```

For implementation details in the portable package, see
`_system\docs\ARCHITECTURE.md`,
`_system\docs\AI_CONTEXT.md`,
`_system\docs\REPORT_DESIGN_SYSTEM.md`,
`_system\docs\ATTENDANCE_DECISION_LEDGER.md`,
`_system\docs\SERVICE_KPI_REFERENCE.md`, and
`_system\docs\METRIC_CATALOG_GUIDE.md`.
