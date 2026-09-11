# WFMHub canonical context for AI and developers

Context version: `1.6.0`
Applies to: WFMHub `0.28.0` and later
Last reviewed: `2026-09-11`

Read this file before proposing or changing WFMHub. When details are needed,
follow the authoritative files listed below. Do not reconstruct decisions from
old conversations, filenames, archived workbooks, or code that is not dispatched.

## Mission

WFMHub is a portable, deterministic Workforce Management hub for a restricted
Windows work machine. Python and SQLite ingest untouched Excel/CSV/TXT extracts,
scope them to the governed FTE roster, calculate auditable marts, and create
focused Excel decision products. PCS uses one permanent lightweight tracker
fed by fixed CSVs through Power Query. Power Query is transport only; Python,
SQLite, the effective rules, and the metric catalog remain the calculation
authority.

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
- **PCS Report & Coaching**: one permanent direct-CSV Power Query performance
  and coaching tracker.

In development:

- Staffing & Coverage
- Realisations
- Final Absenteeism / Shrinkage
- Bonus Management

Utilities include on-demand period analysis, clean exports, health/coverage,
backup and validated custom read-only analysis. The accepted PTO/Away Power Apps
architecture is governed by `docs/PTO_AWAY_APP_IMPLEMENTATION.md`; its repo-side
implementation kit is under `power-platform/wfm-absence-app`, and Microsoft
tenant construction and deployment are pending.

## Source roles

| Source | Authoritative use |
|---|---|
| FTE Count Agent | identity, organisation, FTE, Active/Leaver scope |
| FTE Count PTO/Away | planned time-off precedence and future capacity |
| Verint StartEndTimes | scheduled shift start/end |
| Verint Activities | Final post-day absence/shrinkage codes; Shift Assignment boundary fallback |
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
- Verint Activities never create observed presence, same-day attendance, or gaps.
- Final Absenteeism uses mapped Verint Activities after the day; empty or
  incompletely coded shifts remain explicit review cases.
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
- Ford FR/OEM contains exactly the supplied APFR Ford Assistance, Toyota/Lexus,
  and Chery Assistance queues. Its Flash shows entered, handled, handled in SL,
  and TSL for each sub-LOB and for OEM combined.
- RSA BE contains exactly 43 supplied queues. Four also occur in the exact Ford
  NL Flash allowlist; their dual-Flash membership is intentional, while their
  primary data-model scope remains Ford NL.
- Agent Status and AUX labels are exact configurable business references in
  `default_rules.toml`; unlisted labels use the conservative legacy fallback.
- Active roster LOBs, including Travel-labelled LOBs, flow automatically into
  roster-driven Attendance, Staffing, Absence, PCS and Bonus outputs. A Travel
  service Flash/SL must not be invented without exact queue and forecast maps.

## Data architecture

Pipeline:

```text
untouched extract -> parser/scope gate -> immutable raw/core -> additive marts
                  -> configured metrics/findings -> Excel + fixed Power BI feeds
```

SQLite uses transactions, WAL, integrity checks, a single-writer lock and
versioned migrations. Releases never ship or recreate user data. `SETUP.cmd`
upgrades the same database in place; `UPGRADE.cmd` safely adopts it when a new
release was extracted into another folder. Logical table names are translated by the database facade.
The detailed grains and incremental/atomic refresh rules are in
`docs/ARCHITECTURE.md` and `docs/CLEAN_DATA_CONTRACT.md`.

## Report lifecycles

RTM Daily Control is regenerated under one fixed name and archives the prior
copy. It contains CONTROL plus the four LOB sheets; Service Flash and attendance
callout are not separate workbooks.

Attendance Review is regenerated for the selected completed period. REVIEW
BOARD keeps its exact header on Excel row 4. Users edit five blue ACTUAL fields,
save, then import the same workbook. Decisions persist by immutable Gap ID.

PCS has one permanent collaboration workbook and six fixed-name CSV feeds:
filter lists, LOB cache, agent cache, daily cache, full filter-ready results,
and low-score coaching opportunities. No raw call-leg table is loaded into a
worksheet. The targeted Hub update ingests FTE/Call by Call, refreshes the PCS
mart, and atomically replaces the feeds. Desktop Excel then uses **Data >
Refresh All**. `OVERVIEW` reads four hidden staging tables through classic
exact lookups; its four dropdowns cascade Period → LOB → Team Leader → Agent.
Selector labels live on row 2 and their dedicated values on row 3 at A3, H3,
O3 and V3 so adjacent merged cells cannot display another selector's value.
Period contains the five operational presets, `All available`, and one
`Month YYYY-MM` choice for every calendar month present in the PCS mart.
`COACHING` reads a hidden cache through Period/LOB controls. `PERFORMANCE` is
the only visible query table and may use native table slicers; its `AGENT DAY`
rows expose every available business date without loading raw call legs. The action log
`tblCoachingActions` on the right of `COACHING` is human-owned and is never a query target. Power Query is
installed or repaired once by desktop Excel automation. There is no Data
Model, Power Pivot, macro, raw-data sheet, spill formula or dynamic-array
metadata. A versioned contract migration archives the prior tracker and carries
keyed actions forward; same-version Hub updates preserve tracker bytes.
Presentation formulas use the stable sheet-backed names `PCS_LOB_DATA`,
`PCS_AGENT_DATA`, `PCS_DAILY_DATA`, and `PCS_COACH_DATA`. They must never
directly reference a `tblPcs*` query destination because setup replaces those
ListObjects and Excel rewrites direct references to `#REF!`.

Complete updates atomically publish the governed Power BI star feed under
`Feed\PowerBI`; its manifest is written last. Power BI never reads SQLite or
raw extracts and never recalculates source classification. It relates stable
dimensions and derives ratios only from summed additive counters. The PBIX is a
Power BI Desktop-owned artifact; the portable runtime ships the theme, DAX and
page build contract rather than fabricating a binary PBIX.

Every current report first screen uses the measured grid in
`src/wfmhub/excel_layout.py`: 28 equal 52-pixel columns, four equal KPI cards,
two equal 728 × 310 charts and a compact seven-field action grid. RTM and
Attendance remain Operational; visual consistency does not promote Staffing,
Realisations, Absenteeism/Shrinkage, or Bonus from In Development. PCS keeps the
LOB comparison and filtered agent list together on `OVERVIEW`. `COACHING` shows
the exact source Call ID beside the Coaching Key and contains the permanent
editable action log on the same sheet.

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
- `pcs_tracker.py`: authoritative lightweight PCS Power Query tracker lifecycle
- `pcs_excel.py`: tracker inspection and Windows Excel install/refresh bridge
- `decision_products.py`: Attendance Review and development products
- `shared_feeds.py`: PCS/Absenteeism fixed feeds and Power Query definitions
- `actions.py`: Attendance Review decision import
- `design.py`, `reports.py`, `template_reports.py`: visual contract
- `bonus.py`: governed Bonus Matrix import/report

Documentation:

- `docs/ARCHITECTURE.md`
- `docs/CLEAN_DATA_CONTRACT.md`
- `docs/SERVICE_KPI_REFERENCE.md`
- `docs/QUEUE_REFERENCE_CHANGE_2026-09-10.md`
- `docs/REFERENCE_UPDATE_2026-09-11.md`
- `docs/PCS_LOGIC.md`
- `docs/ATTENDANCE_DECISION_LEDGER.md`
- `docs/REPORT_DESIGN_SYSTEM.md`
- `docs/POWERBI_WFMHUB_PERSPECTIVE.md`
- `docs/POWERBI_PREMIUM_DESIGN_V1.md`

## Safe change protocol

1. Identify the current dispatched builder and effective configuration.
2. State the grain, source authority, key and ratio components before coding.
3. Preserve untouched extracts, stable identities, table names and fixed sheet
   anchors.
4. Add or update focused tests and the end-to-end fixture.
5. Reopen every generated XLSX, inspect ZIP/XML for `#REF!` and unsupported
   formula metadata, and run the full suite.
6. For PCS, verify populated initial caches, valid OOXML, no raw `DATA` sheet,
   all six fixed feed schemas, one-time Power Query installation, governed dropdown
   destinations, migration/action preservation, and byte preservation during
   later same-version Hub updates.
7. Bump the snapshot/report contract only for an explicit design migration.
8. Update this context and the design specification when architecture changes.

## Anti-hallucination checklist

- Never guess or broaden queue membership.
- Never revive APBE, APFR or APDE.
- Never treat Verint Activities as observed attendance; use them only for the
  separately labelled final post-day absence/shrinkage ledger.
- Never call missing evidence a No Show.
- Never mark today’s unfinished shift as Early Leave.
- Never average rates or invent a target.
- Never rename stable sheets, tables, Agent ID, Gap ID or Coaching Key.
- Never overwrite `PCS Live Tracker.xlsx` except for an explicit versioned
  repair that archives the prior file and preserves keyed actions.
- Never point a PCS query at `tblCoachingActions`, and never turn Power Query
  into a KPI calculation layer.
- PCS Overview selectors are cascading Period → LOB → Team Leader → Agent;
  change parent filters left-to-right and reset invalid children to `All`.
- Never add macros, a Data Model, Power Pivot, or a raw call-leg worksheet to PCS.
- Never edit or relocate source extracts.
- Never describe an in-development report as payroll-ready.
- If evidence is insufficient, say what is unknown and inspect the source or
  authority file instead of assuming.
