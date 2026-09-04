# WFMHub Portable

WFMHub reads untouched WFM extracts, keeps a durable SQLite history, and
produces focused Excel workbooks. It runs on a locked-down Windows work machine
without admin rights, installed Python, ODBC, DuckDB, or Python in Excel.

The operating rule is simple: **refresh the sources once, then build only the
product you need**. Extract files are never moved or edited.

## What you see

Normal users work from one folder:

```text
Reports/
  Attendance Callout.xlsx
  Staffing Gaps.xlsx
  OEM Flash.xlsx
  Yesterday Corrections.xlsx
  Final Absenteeism.xlsx
  Bonus Management.xlsx
  PCS Performance.xlsx
  PCS Team.xlsx
  Archive/
```

Every generated product has a fixed name. When WFMHub replaces one, it first
copies the previous version into `Reports\Archive\YYYY-MM-DD`.

`PCS Team.xlsx` is different only in behaviour: it is the persistent team
workspace. WFMHub never overwrites it during a normal refresh, so its
PivotTables, slicers, layout, and `COACHING_LOG` survive.

Technical files live under `_system`. You normally do not open that folder.

## Products

| Product | Operational decision |
|---|---|
| Attendance Callout | Who must be contacted for absence, lateness, or not-seen status? |
| Staffing Gaps | Where is scheduled capacity missing by LOB, language, and interval? |
| OEM Flash | What is Ford OEM service, demand, forecast, and staffing state? |
| Yesterday Corrections | Which completed-day gaps still need a Verint correction? |
| Final Absenteeism | What does corrected Verint contain for final absence and shrinkage? |
| Bonus Management | What did Bonus Matrix v1.2 calculate, and is it safe to release? |
| PCS Team | How are PCS, participation, low scores, and coaching moving daily and monthly? |

The products use one visual identity but not one generic layout. The Flash is an
intraday control page, Attendance is a call list, Corrections is a shift
timeline, and Final Absenteeism is a ledger.

Adherence is not calculated. Service availability means **answered / offered**,
never agent availability.

## Source authority

| Source | Used for |
|---|---|
| FTE roster | Date-aware in-scope agents and organisation fields |
| Verint StartEndTimes | Scheduled start/end and assignment boundaries |
| Storm LILO | First/last presence evidence, including loaded blank rows |
| Storm Agent Status | Observed attendance and interval staffing evidence |
| Verint Activities | Post-correction final absence and shrinkage |
| Verint Forecast | Forecast only |
| Storm APBE/APFR/APDE | Actual service performance only |
| Storm Call by Call | Agent call performance and PCS |
| Bonus Matrix v1.2 | Bonus inputs, KPI configuration, and source reconciliation |

Multi-day files are supported. Row dates are authoritative; filename dates are
only fallback hints. Missing evidence remains missing and is never converted to
a false no-show or zero.

FTE scope is effective-dated: `Active` rows are admitted; `Leaver` rows are
admitted only through `End date if leaver`; other statuses and undated leavers
are excluded. The same rule is applied to schedules, LILO, Agent Status, and
Call by Call for each row's business date.

## Windows quick start

1. Download the portable release and choose **Extract All**.
2. Double-click `SETUP.cmd` once.
3. Paste the folder containing `FTE`, `Storm`, and `Verint`.
4. Double-click `WFMHub.cmd`.
5. Choose **Refresh source data once**.
6. Choose the report you need from Today, Month, PCS Team, or Analyse.
7. Open it directly from `Reports`.

See the [beginner guide](docs/BEGINNER_GUIDE.md) for the exact clicks.

## OEM Flash pilot

The first Flash profile is Ford OEM France. It follows the useful operating
logic in `TOLEARN\Flash FORD OEM.xlsx` without copying its inflated hidden raw
sheets or fragile formulas.

The Flash contains:

- OEM, Ford, and Toyota service level and service availability;
- actual offered calls, Verint forecast, forecast attainment, and call variance;
- hourly mapped queue-group performance;
- scheduled, observed, and productive FTE when matching staffing data exists;
- queue evidence and explicit source freshness.

`Forecast attainment = actual offered / forecast`.
`Variance calls = actual offered - forecast`.

The reference workbook's back-office counters do not yet have a governed raw
source. WFMHub labels that section `NOT_CONFIGURED` rather than fabricating it.

Queue membership lives in `config\queue_mapping.csv`; profile scope lives in
`config\service_profiles.toml`; formulas and targets live in
`config\metric_catalog.toml`.

## PCS Team and Excel Data Model

PCS is the only product that uses Power Query and the Excel Data Model. The
master is `Reports\PCS Team.xlsx`; its four small input files are hidden under
`_system\feeds\pcs\current`:

- `PCS_AgentDay.csv` — fact table, one agent/day;
- `PCS_Calls.csv` — fact table, one low-score coaching opportunity;
- `PCS_Agents.csv` — agent/team/LOB/language dimension;
- `PCS_Dates.csv` — small generated calendar dimension.

The queries load as **Only Create Connection + Add to Data Model**. They do not
load rows to a worksheet. Follow the one-time
[PCS Excel setup guide](docs/EXCEL_TEMPLATE_GUIDE.md).

Setup writes the four ready-to-paste query scripts under
`_system\power_query`. Their local path is literal, so the queries do not
reference one another or combine the workbook with a file source.

Normal PCS routine:

1. Choose **Sync and refresh PCS Team** in WFMHub.
2. Open `Reports\PCS Team.xlsx`.
3. Choose **Data > Refresh All**.
4. Use the PivotTables and slicers.
5. Enter completed coaching in `COACHING_LOG`.
6. On the next sync, WFMHub imports those decisions before refreshing the feed.

The generated `PCS Performance.xlsx` is a static calculation check. The team
master remains the collaborative workbook.

## Configurable logic

| File | Owns |
|---|---|
| `config\wfm_rules.toml` | Attendance and absence evidence classification |
| `config\metric_catalog.toml` | Effective-dated KPI formulas, targets, units, and aggregation |
| `config\analytics_rules.toml` | Deterministic analysis thresholds |
| `config\queue_mapping.csv` | Queue-to-LOB and forecast comparison mapping |
| `config\service_profiles.toml` | Flash/service scope and queue groups |

Percentages are stored as decimals: `0.80` means 80%. Aggregated percentages
use ratios of summed components; WFMHub never averages agent percentages.

## Analysis and clean exports

**Analyze a period** creates an evidence-backed workbook for PCS, service,
forecast, staffing, attendance, final absence, or bonus. It uses deterministic
Python/SQLite logic, not AI.

**Export clean data** produces explicit CSV or XLSX data for a selected period.
Large call datasets should use CSV. The original extract is unchanged.

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
python3 -m wfmhub --home . pcs-team --start 2026-08-01 --end 2026-08-31
python3 -m wfmhub --home . analyze pcs --start 2026-08-01 --end 2026-08-31 --comparison previous_month
python3 -m unittest discover -s tests -v
```

For implementation details, see [Architecture](docs/ARCHITECTURE.md),
[PCS logic](docs/PCS_LOGIC.md), and the
[metric catalog guide](docs/METRIC_CATALOG_GUIDE.md).
