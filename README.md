# WFMHub Portable

WFMHub Portable turns untouched WFM extracts into a durable SQLite database,
finished Excel reports, pivot-ready model tables, and clean exports. It runs on
Windows x64 without administrator rights, installed Python, Power Query, Power
Pivot, ODBC, DuckDB, or Python in Excel.

Raw extracts are never edited and are never copied into normal report sheets.
The SQLite hub remains the source of truth.

## What v0.8 adds

- A management-ready Bonus proposal rebuilt in the WFMHub navy/teal design,
  with a single period, centralized thresholds, formula documentation, an
  executive view, and a hard payroll-release gate.
- A paste-ready three-hour PCS management template based only on the original
  `PCS Report.xlsx` O/P/Q/R logic—not the discarded generated workbook.
- Separate daily cumulative and timestamp-bounded three-hour PCS views by agent
  and LOB, with exact participation and valid-response rate shown separately.
- A persistent `shared_reports` folder for files you send manually. Generated
  workbooks stay out of the public Git repository and source files stay intact.
- Removal of the original PCS workbook's broken external Actions Rate link,
  million-row roster table, full-column calculation ranges, and worldwide raw
  call cache from the management attachment.

## What v0.7 added

- Four finished, independent Excel workbooks: Daily Operations, Yesterday
  Corrections, exact Agent PCS, and Final Absenteeism.
- A single-day operational control view combining the absent/late call list,
  15-minute staffing gaps by roster LOB/language, and APDE service state.
- A Verint-style correction workbook for the latest evidence-complete day,
  with residual gaps, editable decisions, and a full-shift timeline.
- Exact reference-workbook PCS formulas: inbound discrete Q1 scores, counts
  `<=3` and `>3`, and raw-Q1 participation over inbound `PCSStatus=1`.
- An Activities-only final absenteeism ledger with capped daily numerators,
  activity rules, unmapped review, and no mixing with observed LILO/status gaps.
- Visible source health, calculation contracts, governed Excel Tables, semantic
  exception colors, and explicit incomplete-data states in every relevant pack.
- A cleaner fixed dashboard wordmark while retaining the small
  `made by Anass ASSRI` credit.

## What v0.5 added

- One editable, validated business rulebook: `config\wfm_rules.toml`.
- Attendance and absence detected from LILO boundaries plus Agent Status
  intervals, without calculating adherence.
- Automatic `CORRECTED`, `PARTIAL`, and `NOT_CORRECTED` reconciliation against
  the post-day Verint Activities/final schedule export.
- Rule-versioned absence, vacation, unpaid leave, shrinkage, late, early-leave,
  no-show, spell, and Bradford calculations.
- Automatic support for both Verint Activities and wide StartEndTimes extracts.
- APBE/APFR/APDE XLSX or CSV service actuals plus volume-only or full Verint
  forecast extracts.
- Editable `config\queue_mapping.csv` with detailed and comparison scopes.
- Two named SL definitions: gross and short-abandon-adjusted.
- Service availability defined only as `answered / offered`; it is never agent
  availability or adherence.
- Attendance & Absence and Executive Scorecard report packs.
- Pivot-ready Excel Tables without automatically creating PivotTables.
- A generated KPI catalog showing every formula, grain, unit, active version,
  and rulebook SHA-256.
- A safer Rules tool that validates formulas before any refresh.
- A redesigned CMD dashboard showing the active rule version.
- Streamed Agent Status enabled as attendance evidence. Legacy conformance/RTA
  tables remain empty; reports contain no adherence KPI.

## Existing capabilities

- FTE-authoritative agent scope. Exact Agent ID or one unique normalized name
  admits a row; populated Verint `Data Source IDs` remain the operational ID.
- Immutable, fingerprinted ingestion with safe reprocessing after roster changes.
- Multi-day and overlapping-file support using row dates rather than filenames.
- Attendance and correction gaps from Verint schedule boundaries, Storm LILO,
  and Storm Agent Status.
- Forecast and staffing requirements from Verint only.
- Agent call performance and PCS from FTE-scoped Call-by-Call extracts.
- Clean CSV/XLSX exports and trusted custom portable-Python/read-only SQL jobs.
- Read-only AI analysis snapshots containing fixed service, staffing, PCS and
  source-health aggregates; no raw extracts or arbitrary SQL.
- Progress and fixed source-health/latest-date status displays.

## Windows quick start

1. Download `WFMHub-Portable-v0.8.0-win-x64.zip` from GitHub Releases. Do not
   use GitHub's automatic Source code ZIP.
2. Choose **Extract All** and keep the complete `WFMHub` folder together.
3. Double-click `SETUP.cmd` once and select the folder containing `FTE`,
   `Storm`, and `Verint`.
4. Double-click `WFMHub.cmd` for daily work.
5. Choose **Validate rules and build KPI catalog** once before the first refresh.
6. Refresh the required source group and date period.

Use the blank `templates\FTE Count.xlsx` roster if needed. Read the
[beginner guide](docs/BEGINNER_GUIDE.md), [rulebook guide](docs/RULEBOOK_GUIDE.md),
[clean-data contract](docs/CLEAN_DATA_CONTRACT.md), [PCS logic](docs/PCS_LOGIC.md),
[PivotTable guide](docs/PIVOT_GUIDE.md), and
[AI snapshot guide](docs/AI_ANALYSIS_SNAPSHOT.md).

## Output folders

| Folder | Contents |
|---|---|
| `output\operations` | Daily absent/late calls, staffing gaps, and APDE service state |
| `output\corrections` | Latest completed-day residual gaps and full-shift evidence timeline |
| `output\quality_pcs` | Exact Agent PCS summaries, participation, and response detail |
| `output\absence` | Final Activities-only absenteeism ledger and audit evidence |
| `output\reference` | Generated KPI catalog |
| `output\data_exports` | Explicit clean CSV/XLSX exports |
| `output\ai_analysis` | Read-only governed SQLite bundles for external analysis |
| `output\custom` | Custom Lab results |
| `shared_reports` | Bonus and three-hour PCS workbooks shared manually with management |

The detailed sheets such as `ATTENDANCE_CALLS`, `STAFFING_GAPS`,
`SERVICE_LEVEL`, `AGENT_DAY`, `AGENT_MONTH`, `GAPS`, and `ACTIVITY_EVENTS` are
curated Excel Tables suitable for your own PivotTables. They are not raw extracts.

## Source boundaries

| Source | Used for |
|---|---|
| FTE Agent sheet | Agent scope and organisation context |
| Verint StartEndTimes | Operational planned start/end boundary and assignment |
| Storm LILO | Observed daily presence, first login, and last logout |
| Storm Agent Status | Observed within-shift working/non-working intervals; no adherence |
| Verint Activities | Post-day final correction ledger used only for reconciliation |
| Verint Forecast | Forecast and staffing requirements only |
| Storm APBE/APFR/APDE | Service actuals only |
| Storm Call by Call | Agent performance and PCS |

Filename dates are hints. Each dated row or Verint date column determines its
business date. Missing files and missing LILO roster rows are never invented as
no-shows.

## Central rulebook

Edit `config\wfm_rules.toml`, increase `rulebook.version`, then run:

```text
WFMHub.cmd > Validate rules and build KPI catalog
```

The file controls activity taxonomy, absence/payroll/shrinkage flags, standard
day hours, SL formulas, service availability, AHT, forecast deviation, queue
scopes, and other business definitions. Unsafe Python and unknown formula
elements are rejected.

Every modeled row stores the rule version and SHA-256 so a result can be traced
back to its exact definitions.

Queue/file mappings live separately in `config\queue_mapping.csv`. Change the
mapping there and refresh: the hub remaps existing raw data without changing or
reloading the extracts.

## Developer start

```bash
python3 -m venv .venv
. .venv/bin/activate
python3 -m pip install -e .
python3 -m wfmhub --home . doctor
python3 -m wfmhub --home . setup --source-root /path/to/untouched-extracts --non-interactive
python3 -m wfmhub --home . rules catalog
python3 -m wfmhub --home . refresh --start 2026-08-01 --end 2026-08-31 --all-packs
python3 -m wfmhub --home . shared-report pcs "TOLEARN/PCS Report.xlsx" --date 2026-08-31
python3 -m unittest discover -s tests -v
```

See [Architecture](docs/ARCHITECTURE.md) for grains, formulas, audit behavior,
and the extension pattern.
