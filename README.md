# WFMHub Portable

WFMHub turns untouched WFM extracts into a durable SQLite database, focused
Excel decision products, compact Excel Data Model inputs, clean exports, and
repeatable on-demand analysis. The Windows package runs without admin rights,
installed Python, ODBC, DuckDB, Power Query, or Python in Excel.

The extracts are read-only inputs. Normal reports contain curated business
grains—not copied raw files—and SQLite remains the calculation authority.

## Current operating model

The hub now produces one workbook for one decision:

| Product | Management question |
|---|---|
| PCS Performance | How are daily and current-month PCS and participation moving versus the comparable prior month? |
| Bonus Performance | What does Bonus Matrix v1.2 calculate, what changed, and is it safe to release? |
| Service Performance | Is the selected mapped LOB meeting SL, service availability, forecast, and AHT expectations? |
| Staffing & Coverage | Where is scheduled capacity not present by LOB/language and interval? |
| Attendance Callouts | Which agents require contact or follow-up on every selected date? |
| Attendance Corrections | Which completed-day observed gaps remain to be corrected in Verint? |
| Final Absence & Shrinkage | What does the corrected Verint ledger contain for payroll and final reporting? |

Every workbook has the same navy/teal/gold shell, a status badge, KPI cards,
action-oriented detail, definitions, and hidden audit lineage. Adherence is not
calculated. Service availability always means **answered / offered**.
The Ford OEM profile follows the original Flash's gross SL and actual/forecast
attainment definitions while labelling them unambiguously.

## Source boundaries

| Source | Used for |
|---|---|
| FTE Agent sheet | Authoritative in-scope roster and organisation fields |
| Verint StartEndTimes | Planned start/end and assignment boundaries |
| Storm LILO | First/last daily presence evidence, including loaded blank rows |
| Storm Agent Status | Within-shift observed presence and staffing evidence |
| Verint Activities | Post-correction final absence/shrinkage ledger and correction reconciliation |
| Verint Forecast | Forecast and schedule requirements only |
| Storm APBE/APFR/APDE | Service actuals only |
| Storm Call by Call | Agent call performance and PCS |
| Bonus Matrix v1.2 | Monthly bonus inputs, configuration, policy, and source reconciliation |

Filename dates are hints. Row dates and Verint date columns define the business
date, so multi-day extracts are supported. Missing evidence is shown as missing;
it is never invented as a no-show or zero.

## Windows quick start

1. Download the portable Windows release and choose **Extract All**.
2. Double-click `SETUP.cmd` once and select the folder containing `FTE`,
   `Storm`, and `Verint`.
3. Double-click `WFMHub.cmd`.
4. Choose **Validate rules and build governance catalog** after setup and after
   every configuration change.
5. Choose **Refresh hub data**, a source group, a period, and the report products
   you need.

Use `templates\FTE Count.xlsx` if you need the standard roster. Start with the
[beginner guide](docs/BEGINNER_GUIDE.md).

## Output folders

| Folder | Contents |
|---|---|
| `output\pcs` | Daily/MTD/prior-month PCS and participation |
| `output\bonus` | Governed Bonus Matrix result and release controls |
| `output\service` | Mapped LOB service performance; Ford OEM France is the starter profile |
| `output\staffing` | LOB/language interval coverage and exceptions |
| `output\attendance` | Live same-day contact queue |
| `output\corrections` | Completed-day residual gaps and shift visualization |
| `output\absence` | Final Activities-only absence and shrinkage ledger |
| `output\analysis` | On-demand domain analysis for any selected period |
| `output\model_data\<product>` | Small CSV tables used by optional Excel Data Model masters |
| `output\data_exports` | Explicit clean CSV/XLSX exports |
| `output\custom` | Trusted custom Python/read-only SQL results |

The former “management shared report” route is retired. Each official report
product is presentation-ready and follows the same contract.

## Excel PivotTable and slicer route

Normal generated workbooks work immediately and do not need Power Query. If you
want native PivotTables and slicers:

1. In `WFMHub.cmd`, choose **Create an Excel Pivot/slicer master** once.
2. Follow [Excel template setup](docs/EXCEL_TEMPLATE_GUIDE.md) to create direct
   CSV queries as **Connection Only + Add to Data Model**.
3. On normal report days, WFMHub refreshes
   `output\model_data\<product>` automatically.
4. Open the protected master in `templates\reports`, choose **Refresh All**, and
   then **Save As** the copy you will send.

WFMHub never edits a configured master again, preserving its Data Model,
PivotTables, slicers, and your layout. Raw data is never loaded to a worksheet.

## Configurable logic

Calculations are centralized and separated by responsibility:

| File | Owns |
|---|---|
| `config\wfm_rules.toml` | Attendance/absence evidence classification and source parsing rules |
| `config\metric_catalog.toml` | Effective-dated KPI formulas, targets, units, scope, priority, and aggregation |
| `config\analytics_rules.toml` | Deterministic finding thresholds and limits |
| `config\report_catalog.toml` | Workbook names and ordered sheet contracts |
| `config\queue_mapping.csv` | Source queues, report LOBs, and forecast comparison scopes |
| `config\service_profiles.toml` | Effective-dated service products, included scopes, selected governed metrics, and queue groups |

Percentages are decimals (`0.80` means 80%). Higher-level KPIs always use a
ratio of summed counters; percentages are never averaged. Run the validation
menu after an edit so invalid formulas, dates, priorities, mappings, and report
contracts fail before a refresh.

## Bonus Matrix and analysis

Import the final `Bonus_Matrix_v1.2.xlsx` from the menu. WFMHub hashes the source,
imports it without modifying it, recalculates the governed scenario, and keeps
Released Payout blocked until the configured input, policy, and eligibility
controls pass. Absence must have one financial consequence—KPI, eligibility, or
proration—not a duplicate penalty.

Choose **Analyze a period** to create an evidence-backed workbook for PCS,
service, forecast, staffing, attendance, final absence, or bonus. This is
deterministic Python/SQLite analysis. The optional prompt under `prompts` can be
used manually with your approved Copilot account after attaching only the
finished workbook.

## Developer commands

```bash
python3 -m venv .venv
. .venv/bin/activate
python3 -m pip install -e .
python3 -m wfmhub --home . setup --source-root /path/to/extracts --non-interactive
python3 -m wfmhub --home . rules catalog
python3 -m wfmhub --home . refresh --start 2026-08-01 --end 2026-08-31 --all-packs
python3 -m wfmhub --home . import-bonus "TOLEARN/Bonus_Matrix_v1.2 (1).xlsx"
python3 -m wfmhub --home . analyze pcs --start 2026-08-01 --end 2026-08-31 --comparison previous_month
python3 -m wfmhub --home . template-init --pack pcs --start 2026-08-01 --end 2026-08-31
python3 -m unittest discover -s tests -v
```

See [Architecture](docs/ARCHITECTURE.md), [PCS logic](docs/PCS_LOGIC.md),
[metric catalog guide](docs/METRIC_CATALOG_GUIDE.md), and
[clean-data contract](docs/CLEAN_DATA_CONTRACT.md) for the technical contracts.
