# WFMHub canonical context for AI and developers

Context version: `1.1.0`
Applies to: WFMHub `0.23.1` and later
Last reviewed: `2026-09-08`

Read this file before proposing or changing WFMHub. When details are needed,
follow the authoritative files listed below. Do not reconstruct decisions from
old conversations, filenames, archived workbooks, or code that is not dispatched.

## Mission

WFMHub is a portable, deterministic Workforce Management hub for a restricted
Windows work machine. Python and SQLite ingest untouched Excel/CSV/TXT extracts,
scope them to the governed FTE roster, calculate auditable marts, and create
focused Excel decision products. Excel Power Query is transport for permanent
collaboration workbooks, not the KPI calculation engine.

Non-goals:

- no runtime AI or hidden statistical judgement;
- no adherence KPI;
- no DuckDB, ODBC driver, external database server, installer, or `pip` at work;
- no APBE/APFR/APDE data sources;
- no editing, moving, renaming, or cleaning source extracts in place;
- no guessed queue membership, activity mapping, absence, or missing values;
- no required Excel Data Model or Power Pivot.

`prompts/COPILOT_WFM_ANALYST.md` is an optional manual prompt used only after a
human attaches a generated workbook to Microsoft Copilot.

## Authority hierarchy

Use the first applicable source:

1. an explicit current user decision;
2. current dispatched code plus passing tests for actual behavior;
3. installed effective `config/*.toml` and mapping files on the work machine;
4. shipped `config/default_*` files for release defaults;
5. approved documents, especially this file and the report design contract;
6. a specifically cited TOLEARN source as historical business evidence only.

`src/wfmhub/report_packs.py` is the report-dispatch authority. Unreachable legacy
builders, `_pre_*` files, archives and old filenames are never current authority.
Repo-local non-default config files are runtime state and are not release
defaults. Change `default_*` when changing shipped behavior.

## Product maturity and menu

Operational:

- **RTM Daily Control**: combined Service Flash, per-LOB attendance pulse and
  same-day call list.
- **Attendance Review**: completed-day exact gaps, schedule-versus-observed
  review, break/meal control, and human decisions.
- **PCS Operational Tracker**: permanent shared performance and coaching file
  refreshed from five governed CSV feeds through Power Query.

In development:

- Staffing & Coverage
- Realisations
- Final Absenteeism / Shrinkage
- Bonus Management

Utilities include on-demand period analysis, clean exports, health/coverage,
backup and validated custom read-only analysis. The PTO/Away Power Apps project
is recorded in `docs/PTO_AWAY_APP_FEASIBILITY.md`; it is not implemented here.

## Source roles

| Source | Authoritative use |
|---|---|
| FTE Count Agent | identity, organisation, FTE, Active/Leaver scope |
| FTE Count PTO/Away | planned time-off precedence and future capacity |
| Verint StartEndTimes | scheduled shift start/end |
| Verint Activities | Shift Assignment boundary fallback only; not attendance evidence |
| Storm Agent Status | primary observed attendance, gaps, breaks and meals |
| Storm LILO | fallback presence/control when Agent Status coverage is missing |
| Storm Call by Call | Service Flash, call workload and PCS call legs |
| Verint Forecast | forecast volume/FTE; source grain is 15 minutes |
| Queue mapping | exact source queue to service/comparison scope |
| Service profiles | exact Flash queue allowlists and staffing LOBs |

Agent identity uses Client ID/operational Agent ID as text. FTE Status `Active`
is eligible. `Leaver` is eligible only through its populated leave date. A
unique normalized-name fallback may attach organisation fields, but ambiguous
identity is excluded. Call demand is also admitted by an exact reviewed queue
mapping so abandoned demand is not lost.

## Non-negotiable business rules

- Row dates win over filename dates. Multi-day files are valid.
- Missing evidence is unknown, never zero and never automatically No Show.
- Agent Status is attendance authority; LILO is fallback/control.
- Verint Activities do not create attendance, absence, or correction truth.
- Today’s unfinished shift tail is never Early Leave.
- A person who leaves and returns has an internal exact gap, not one continuous
  early leave.
- Explicit Meal Aux is Lunch and is not a gap.
- PTO/Away overlapping scheduled work is planned time, not No Show or an
  attendance gap. It also changes net staffing.
- RTM No Show HC requires a started working shift, no presence, and explicit
  agent-specific no-show evidence. Unknown is a data check.
- Availability means routed/offered service availability, never agent
  availability and never adherence.
- PCS average is `SUM(valid inbound Q1 score) / SUM(valid inbound Q1 count)`.
- PCS participation is `SUM(inbound raw Q1 nonblank) / SUM(inbound PCSStatus=1)`.
- Higher grains always sum additive counters before division. Never average
  agent/hour/day percentages.
- A low-score coaching case is one exact call identified by Coaching Key.
- Service Level follows the documented Storm C/(A+B-D) contract and exact
  screenshot-reviewed queue allowlists. See `docs/SERVICE_KPI_REFERENCE.md`.
- Never infer a queue from its name, language suffix, country, or apparent LOB.

## Data architecture

Pipeline:

```text
untouched extract -> parser/scope gate -> immutable raw/core -> additive marts
                  -> configured metrics/findings -> focused Excel product
```

SQLite uses transactions, WAL, integrity checks, a single-writer lock and
versioned migrations. Logical table names are translated by the database facade.
The detailed grains and incremental/atomic refresh rules are in
`docs/ARCHITECTURE.md` and `docs/CLEAN_DATA_CONTRACT.md`.

## Report lifecycles

RTM Daily Control is regenerated under one fixed name and archives the prior
copy. It contains CONTROL plus the four LOB sheets; Service Flash and attendance
callout are not separate workbooks.

Attendance Review is regenerated for the selected completed period. REVIEW
BOARD keeps its exact header on Excel row 4. Users edit five blue ACTUAL fields,
save, then import the same workbook. Decisions persist by immutable Gap ID.

PCS is one permanent shared workbook. Python/SQLite atomically publish:

- `PCS_LOB_SCORECARD_CURRENT.csv`
- `PCS_RESULTS_CURRENT.csv`
- `PCS_COACHING_OPPORTUNITY_CURRENT.csv`
- `PCS_AGENT_DAY_CURRENT.csv`
- `PCS_SCOPE_CURRENT.csv`

Power Query replaces only `tblPcsLob`, `tblResults`, `tblCoachingQueue`,
`tblPcsData` and `tblPcsScope`. It never targets `tblCoaching`. Normal PCS update never rebuilds
or replaces the workbook. Repair/rebuild is explicit, versioned, archives the
old copy and migrates readable keyed coaching actions.

PCS `OVERVIEW` uses the measured grid in `src/wfmhub/excel_layout.py`: 28 equal
52-pixel columns, four equal KPI cards, two equal 728 × 310 charts and a compact
team action grid. Query tables live in detail or hidden staging sheets, never on
the dashboard. Period → LOB → Team Leader → Agent selectors come from
`tblPcsScope`. `COACHING` shows the exact source Call ID beside the permanent
Coaching Key; Call ID is calculated and is not typed by Quality.

Generated reports use `WFMHUB-DESIGN`; see `docs/REPORT_DESIGN_SYSTEM.md`.
They contain no AI branding or generated-by-AI language.

## Authoritative file index

Configuration:

- `config/default.toml`: paths, source discovery and runtime defaults
- `config/default_rules.toml`: attendance evidence and classification settings
- `config/default_metrics.toml`: KPI arithmetic, methods and targets
- `config/default_queue_mapping.csv`: exact queue mapping
- `config/default_service_profiles.toml`: exact Flash allowlists/LOBs
- `config/default_reports.toml`: ordered workbook sheet contracts
- `config/default_analytics.toml`: deterministic analysis thresholds

Implementation:

- `ingestion.py`: parsers, row dates, discovery and roster scope
- `models.py`: dimensions, timelines and marts
- `metrics.py`: versioned KPI methods
- `report_packs.py`: current report lifecycle and dispatch
- `service_flash.py`: RTM and LOB Service Flash
- `decision_products.py`: PCS, Attendance Review and development products
- `shared_feeds.py`: atomic collaboration feeds and Power Query text
- `pcs_excel.py` plus `Install-PCSWorkbook.ps1`: permanent Excel connection
- `actions.py`: Attendance Review decision import
- `design.py`, `reports.py`, `template_reports.py`: visual contract
- `bonus.py`: governed Bonus Matrix import/report

Documentation:

- `docs/ARCHITECTURE.md`
- `docs/CLEAN_DATA_CONTRACT.md`
- `docs/SERVICE_KPI_REFERENCE.md`
- `docs/PCS_LOGIC.md`
- `docs/ATTENDANCE_DECISION_LEDGER.md`
- `docs/REPORT_DESIGN_SYSTEM.md`

## Safe change protocol

1. Identify the current dispatched builder and effective configuration.
2. State the grain, source authority, key and ratio components before coding.
3. Preserve untouched extracts, stable identities, table names and fixed sheet
   anchors.
4. Add or update focused tests and the end-to-end fixture.
5. Reopen every generated XLSX, inspect ZIP/XML for `#REF!` and unsupported
   formula metadata, and run the full suite.
6. For PCS, run Windows Excel install -> refresh -> calculate -> save -> reopen.
7. Bump the workbook template version only for an explicit design migration.
8. Update this context and the design specification when architecture changes.

## Anti-hallucination checklist

- Never guess or broaden queue membership.
- Never revive APBE, APFR or APDE.
- Never treat Verint Activities as observed attendance or final absence truth.
- Never call missing evidence a No Show.
- Never mark today’s unfinished shift as Early Leave.
- Never average rates or invent a target.
- Never rename stable sheets, tables, Agent ID, Gap ID or Coaching Key.
- Never overwrite PCS during normal refresh.
- Never load Power Query into COACHING or a permanent action table.
- Never edit or relocate source extracts.
- Never describe an in-development report as payroll-ready.
- If evidence is insufficient, say what is unknown and inspect the source or
  authority file instead of assuming.
