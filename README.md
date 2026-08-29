# WFMHub Portable

WFMHub Portable turns untouched WFM extracts into a durable SQLite database,
finished Excel reports, pivot-ready model tables, and clean exports. It runs on
Windows x64 without administrator rights, installed Python, Power Query, Power
Pivot, ODBC, DuckDB, or Python in Excel.

Raw extracts are never edited and are never copied into normal report sheets.
The SQLite hub remains the source of truth.

## What v0.6 adds

- A data-first validation release: the new business workbooks are intentionally
  parked until their governed clean datasets are accepted.
- Strict header-based separation of Verint StartEndTimes (operational plan) and
  Activities (corrected final ledger).
- Explicit daily attendance call actions with provisional current-day gating;
  unfinished shifts never become final early leaves.
- 15-minute staffing gaps by roster LOB/language and exact full-shift evidence
  segments for a later Verint-like visual.
- Overlap-safe residual correction intervals after matching the union of final
  Activities, plus an independent Activities-only final absenteeism ledger.
- Exact reference-workbook PCS logic: inbound discrete Q1 scores, counts `<=3`
  and `>3`, and raw-Q1 participation over inbound `PCSStatus=1`.
- Governed daily/team/month PCS, attendance, staffing, service, correction,
  timeline, and final-absence exports with manifests.

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
- Progress and fixed source-health/latest-date status displays.

## Windows quick start

1. Download `WFMHub-Portable-v0.6.0-win-x64.zip` from GitHub Releases. Do not
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
and [PivotTable guide](docs/PIVOT_GUIDE.md).

## Output folders

| Folder | Contents |
|---|---|
| `output\operations` | Attendance and correction candidates; no adherence KPI |
| `output\absence` | Agent-day absence, classified events, spells, Bradford, and rules |
| `output\intraday` | APBE/APFR/APDE service actuals and separate Verint forecast |
| `output\scorecard` | Daily service, forecast, absence, and PCS KPI facts |
| `output\quality_pcs` | Agent call performance, PCS, and survey responses |
| `output\reference` | Generated KPI catalog |
| `output\data_exports` | Explicit clean CSV/XLSX exports |
| `output\custom` | Custom Lab results |

The `PIVOT_*`, `KPI_DAILY`, `SERVICE_INTERVALS`, and `ABSENCE_DAILY` sheets are
curated Excel Tables designed as PivotTable sources. They are not raw extracts.

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
python3 -m unittest discover -s tests -v
```

See [Architecture](docs/ARCHITECTURE.md) for grains, formulas, audit behavior,
and the extension pattern.
