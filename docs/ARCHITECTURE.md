# Architecture

## Stable extension seam

Every feature follows the same layers:

```text
untouched source -> parser/scope gate -> raw/core -> additive marts
                 -> effective metric catalog -> semantic values
                 -> deterministic findings + report datasets -> Excel
```

A new forecast KPI, queue feed, or attendance rule adds an adapter, a numbered
migration, a focused model, tests, and curated report output. Existing extract
files are never edited.

## Portable installation

```text
WFMHub/
├── WFMHub.cmd                  daily menu
├── SETUP.cmd                   system check and first setup
├── Reports/                    fixed-name reports + dated archive
├── Feed/                       explicit clean CSV/XLSX exports
├── config/default.toml         shipped defaults
├── config/wfmhub.toml          user configuration
├── config/default_rules.toml   shipped evidence/domain defaults
├── config/wfm_rules.toml       editable/versioned evidence rulebook
├── config/default_metrics.toml shipped metric methods
├── config/metric_catalog.toml  editable formulas/targets/effective dates
├── config/default_analytics.toml shipped finding thresholds
├── config/analytics_rules.toml editable deterministic-analysis settings
├── config/default_reports.toml shipped workbook contracts
├── config/report_catalog.toml  editable/validated report contract
├── config/default_queue_mapping.csv shipped mapping defaults
├── config/queue_mapping.csv    editable queue/file/scope mapping
├── config/default_service_profiles.toml shipped service-product defaults
├── config/service_profiles.toml editable effective-dated service profiles
└── _system/
    ├── runtime/                official embedded CPython + pure-Python packages
    ├── app/wfmhub/             application code
    ├── app/sql/migrations/     versioned SQL schema
    ├── templates/              blank FTE and technical templates
    ├── docs/                   user/developer documentation
    ├── prompts/                optional manual Copilot handoff prompt
    └── database, logs, backups, output, input and custom tools
```

The work computer runs no installer or `pip`. SQLite is Python's standard
`sqlite3` module. Only pure-Python `openpyxl`, `XlsxWriter`, and `et_xmlfile`
packages are added. The release builder rejects any unexpected `.dll`, `.pyd`,
or `.exe` and any path containing DuckDB or `msvc_runtime`.

SQLite uses WAL, `synchronous=FULL`, integrity checks, a 30-second busy timeout,
and online backups. WFMHub must stay on a local writable disk, not a network or
sync-managed folder. A process lock permits one writer; report-only readers are
opened read-only.

Logical names such as `raw.lilo` are translated by the database facade into
SQLite tables such as `raw_lilo`. Business code stays readable and backend
details remain in one module.

## Tables and grains

| Logical table | Grain |
|---|---|
| `meta.source_file` | One immutable file-content + roster-scope version |
| `meta.refresh_run` | One refresh attempt |
| `meta.quality_issue` | One issue detected by the current model run |
| `meta.rule_application` | Exact rule version/hash applied by one model run |
| `meta.mapping_application` | Exact queue-mapping hash applied by one model run |
| `meta.metric_application` | Exact metric-catalog version/hash applied by one model run |
| `meta.analytics_application` | Exact analytics version/hash applied by one model run |
| `raw.fte_agent` | One FTE Agent-sheet row |
| `raw.schedule_shift` | One admitted Verint schedule row |
| `raw.schedule_event` | One parsed event interval belonging to an admitted shift |
| `raw.lilo` | One admitted Storm LILO row |
| `raw.agent_status` | One admitted status interval |
| `raw.forecast_interval` | One queue/forecast interval; not agent-scoped |
| `raw.queue_actual` | One queue/15-minute actual interval; not agent-scoped |
| `raw.call_leg` | One admitted, typed Call-by-Call source leg |
| `raw.bonus_import` | One immutable Bonus Matrix content hash/version |
| `raw.bonus_agent_month` | One imported agent/month input row |
| `raw.bonus_kpi_rule` | One imported population/KPI threshold rule |
| `raw.bonus_policy` | One imported policy decision row |
| `core.clean_call_leg` | Deduplicated active call-leg view by stable Call Key |
| `core.dim_agent` | One operational Agent ID |
| `core.correction_action` | Legacy compatibility table; correction workbooks are no longer imported |
| `core.pcs_coaching_action` | Legacy compatibility table; generated PCS coaching stays in Excel |
| `raw.fte_time_off` | One governed PTO/Away register row from the standard FTE workbook |
| `mart.attendance_agent_day` | One scheduled Agent ID/day |
| `mart.conformance_agent_day` | Legacy compatibility table; empty in v0.5 |
| `mart.correction_candidate` | One observed LILO/status gap plus Verint-final check |
| `mart.correction_residual_segment` | One still-uncovered interval requiring Verint review/injection |
| `mart.staffing_interval` | One 15-minute roster LOB/language staffing interval |
| `mart.shift_timeline_segment` | One exact planned-versus-observed timeline segment |
| `mart.planned_time_off_segment` | One schedule-clipped, non-overlapping PTO/Away interval |
| `mart.rta_snapshot` | Legacy compatibility table; empty in v0.5 |
| `mart.verint_final_exception` | One final Verint interval with no observed supporting gap |
| `mart.forecast_hour` | One raw forecast queue/hour plus mapped scopes |
| `mart.intraday_queue_interval` | One actual queue/15-minute interval |
| `mart.agent_pcs_day` | One admitted Agent ID/day with call and PCS measures |
| `mart.verint_final_absence_event` | One classified Activities-only final-ledger evidence interval |
| `mart.verint_final_absence_agent_day` | One Activities-only final absence result per Agent ID/day |
| `mart.absence_event` | One observed, schedule-clipped LILO/status gap with final label |
| `mart.absence_agent_day` | One payroll absence/vacation/shrinkage result per Agent ID/day |
| `mart.service_interval` | One rule-versioned APBE/APFR/APDE service interval |
| `mart.call_service_hour` | One mapped Call-by-Call queue/hour with interaction-deduplicated service counters |
| `mart.metric_value` | One configured KPI observation per source entity/method |
| `mart.analysis_finding` | One ranked deterministic finding with evidence filter |
| `mart.bonus_agent_month` | One governed monthly bonus result per agent |
| `mart.bonus_kpi_result` | One governed agent/KPI monthly result |
| `mart.source_health` | One configured source family |

## Report packs

The shared SQLite hub can serve multiple workbooks without mixing their grains:

| Pack | Current file | Scope |
|---|---|---|
| `pcs` | `Reports/PCS Performance.xlsx` | Selector-driven PCS performance and coaching workbook |
| `bonus` | `Reports/Bonus Management.xlsx` | Imported Bonus Matrix result and release controls |
| `service` | `Reports/Service Flashes.xlsx` | RSA NL/BE and Ford NL/OEM daily mapped-call service control |
| `realisations` | `Reports/Realisations.xlsx` | All mapped LOB actual/forecast, service, staffing, absence and shrinkage results |
| `staffing` | `Reports/Staffing Gaps.xlsx` | Full-period actual staffing control and future capacity planning |
| `attendance` | `Reports/Attendance Callout.xlsx` | No-show/late/not-seen contact queue |
| `corrections` | `Reports/Attendance Review.xlsx` | Selected-period completed-day residual gaps and shift visualization |
| `absence` | `Reports/Final Absenteeism.xlsx` | Activities-only final absence/shrinkage ledger |

Products share the same visual identity but use purpose-specific layouts. The
Service Flashes begins with `CONTROL` and four purpose-built Flash sheets;
Corrections uses `VERINT_INJECTION` and `SHIFT_VIEW`; Attendance remains an action list. Legacy `operations` and
`quality_pcs` builders remain callable under `_system/legacy_reports` but are
absent from the menu.

PCS is generated directly from SQLite. `PCS_DATA`, `COACHING_QUEUE` and
`COACHING` are ordinary visible Excel Tables. Dashboard selectors use Excel
`SUMPRODUCT`; `TEAM_VIEW` is a conventional, filterable Excel Table with
Python-calculated latest-day, current-week, current-MTD and previous-MTD
results. It deliberately uses no spill arrays or `_xlfn` functions. Users may
still add Table slicers manually. A first build needs no Excel
connection. For one long-lived shared file, the user may link the two
replaceable PCS tables to fixed CSV feeds with Power Query. `COACHING` remains
the permanent editable log. No Data Model or ODBC driver is required.

Final Absenteeism uses the same collaboration boundary. Power Query may replace
`tblAbsenceData`, `tblActionQueue`, and `tblActivityDetail` from stable CSVs;
it must never load into permanent `tblActions`. `TEAM_VIEW` and
`COMPONENT_VIEW` read the refreshed tables directly.

## Agent scope and identity

FTE is the authority for “our agents.” Scope is evaluated on each source row's
business date. Before an agent-level row enters the
active raw layer:

1. Admit FTE Status `Active` for every date.
2. Admit Status `Leaver` only through its populated `End date if leaver`.
3. Exclude other statuses and undated Leavers.
4. Normalize the source Agent ID and apply that effective-dated eligibility.
5. Otherwise normalize accents, case, punctuation, and whitespace in the name.
6. Keep it only when that name maps to exactly one eligible FTE row. If the FTE
   Client ID is blank, retain the real operational Agent ID and attach the
   unique matching FTE organisation fields; ambiguous names stay excluded.
7. Preserve a populated operational source ID. In particular, Verint `Data
   Source IDs` remains the schedule Agent ID.
8. Apply the same gate to schedules, LILO, and Agent Status. For Call-by-Call,
   also admit a row when its queue exactly matches the reviewed queue map; this
   preserves mapped abandoned demand and cross-operation handling.
9. Exclude everything else and count it as outside roster. Agent-level Call-by-
   Call marts still require a join to the governed FTE dimension.

A populated unmatched ID can therefore be admitted by a unique roster name,
but an ambiguous or missing name cannot. This is a scope decision, not fuzzy
matching.

The scope has a deterministic fingerprint. If FTE changes, the same untouched
schedule/LILO/status/call file is reprocessed against the new roster. This prevents
both stale worldwide rows and the “new agent missing from an unchanged file”
problem. Forecast and APBE/APFR/APDE are queue data, so they bypass the agent gate.
Changing the queue map also changes the Call-by-Call scope fingerprint, causing
unchanged call extracts to be safely reprocessed.

## Incremental and atomic refresh

1. Discover configured files read-only.
2. Calculate SHA-256, size, modified time, and agent-scope fingerprint.
3. Skip an already-active successful match.
4. Parse a new version; stream large LILO, Agent Status and Call-by-Call CSVs in bounded batches.
5. In one file transaction, append immutable raw rows, deactivate the previous
   path version, and activate the new version.
6. If parsing fails, roll back its raw rows and leave the previous good version
   active. Retrying the same fingerprint is supported.
7. Rebuild all selected-period models inside one savepoint. Any failure restores
   every previous mart.

A same-path A→B→A change reactivates A's immutable rows rather than duplicating
them. Deleting a physical extract does not silently erase loaded history.

## Row dates and multi-day extracts

Filename dates are hints, never the primary business date. Schedule, Agent
Status, Call-by-Call, Forecast and APBE/APFR/APDE use row fields. LILO prefers a
row-level Date field, then the first/last boundary. A single filename date may
be used only for a boundary-blank daily LILO row. In a multi-day LILO file, a
row with both boundaries blank and no Date is rejected because its day cannot
be proven.

## Call and PCS model

Call CSVs are FTE-or-exact-mapped-queue scoped before storage. A stable Call Key combines call
references, direction, agent and timestamps; `core.clean_call_leg` selects the
newest active version across overlapping history extracts. The official PCS
contract reproduces `TOLEARN/PCS Report.xlsx`: keep inbound legs with an Agent
ID, accept Q1 only when its numeric value is one of the configured discrete
scores (default `1,2,3,4,5`), and calculate `sum(valid Q1) / count(valid Q1)`.
The Flash mart separately groups inbound legs by Interaction Key and mapped
comparison scope, so a transfer remains one offered interaction while handled
leg seconds remain available for weighted AHT.
Counts `<=3` and `>3` remain counts. Participation is `inbound raw-Q1 nonblank /
inbound PCSStatus=1`; invalid raw answers stay in that numerator and are
separately counted. Q2 and Mode 2 are diagnostics only. Higher grains always
sum counters before dividing. A valid inbound Q1 `<=3` is one coaching
opportunity. The generated `COACHING` sheet carries those calls plus five blue
manual columns. Actions Rate is completed workbook rows divided by all
low-score opportunities. Those five fields carry forward from the prior
current workbook by Coaching Key and never enter SQLite. The PCS mart rebuild window expands to the first
day of the previous month through the selected end date so Today/Current Week
reports retain MTD and comparison context. See [PCS logic](PCS_LOGIC.md).

## Clean exports and Custom Lab

Clean exports stream selected-period results to UTF-8 CSV or bounded XLSX with
a manifest. Custom SQL is restricted to one SELECT/WITH statement on a
query-only connection. Custom Python receives the same read-only query context,
but is trusted executable local code and is not an operating-system sandbox.

## Deterministic analysis boundary

`analytics.py` reads only governed semantic metrics and source-health state. It
uses configured targets, sample minimums, threshold deltas, and period changes
to write ranked findings. A finding stores its metric/method, scope, selected
period, values, evidence dataset/filter, and catalog/analytics hashes. There is
no model server, paid API, GPU, prompt execution, or database upload path.

`on_demand_analysis.py` exposes that same boundary for a user-selected period,
domain and comparison. Its workbook contains `FINDINGS`, `METRICS`, and curated
`EVIDENCE`; a period change is descriptive and never presented as proof of
causality. Default outputs are visible under `Reports/Analysis`.

`_system/prompts/COPILOT_WFM_ANALYST.md` is a static manual aid. A user may attach a
chosen finished workbook to an approved Copilot account. The runtime never
connects Copilot to SQLite or raw extracts, and Copilot is never a calculation
authority.

## Terminal dashboard

The daily menu reads a small read-only dashboard snapshot from existing marts.
It never rebuilds a model merely to draw the screen. The fixed-width ASCII
panel reports the last refresh, selected period, database size, agent count,
source-health counts, quality counts, and the maximum loaded source business
date with its family. Identifying the family prevents a future Forecast date
from being presented as the freshness date for operational actuals.

The dashboard is failure-tolerant: a missing, unconfigured, or incompatible
database produces a setup/check state rather than crashing the menu renderer.
The launcher uses native CMD title/color commands; the application itself adds
no terminal package and redirected output is never cleared.

## Attendance, absence and correction gates

The order is deliberate:

1. Select the Verint StartEndTimes schedule boundary for Agent ID/day.
2. Build an exclusive Agent Status timeline clipped to that boundary.
3. Use Agent Status as the primary boundary when its elapsed-shift coverage is
   sufficient; use LILO only as a sparse-status outer-boundary fallback.
4. Use exclusive Agent Status `Logged Off`/`Unavailable` intervals between those
   boundaries for mid-shift gap detection.
5. Build gaps before reading any final Verint activity.
6. Match each observed gap against the Activities-only final ledger.
7. Label it `CORRECTED`, `PARTIAL`, or `NOT_CORRECTED` without changing the
   observed interval.

“No show” requires a completed scheduled working shift plus positive evidence:
either a loaded daily LILO row with both boundaries blank, or sufficient Agent
Status coverage in which every observed interval is Logged Off. Missing files,
missing rows, and incomplete evidence are never a no-show.

A logout followed by a later active status is an internal gap. The later return
extends the actual presence boundary, so the earlier logout cannot become an
early leave. Leading logout intervals are late; trailing logout intervals are
early leave only after the scheduled end; multiple reconnect cycles remain
separate gaps unless they are within the configured merge tolerance.

Final Verint activities are deliberately not subtracted from late, early or
status gaps. Doing so would hide the original problem immediately after it was
corrected, destroying the audit trail.

## Configuration boundaries and calculation audit

`config/wfm_rules.toml` classifies evidence and activities. It owns standard-day
and tolerance policy plus PCS source parsing, but contains no KPI arithmetic.
`config/metric_catalog.toml` is the sole KPI source: formula components,
denominator, sample, aggregation, target, direction, scope, priority, and
effective dates. `analytics_rules.toml` owns finding sensitivity;
`report_catalog.toml` validates the presentation contract; `queue_mapping.csv`
owns queue/service/forecast naming; and `service_profiles.toml` owns
effective-dated report scope, governed metric selection and display groups.

The safe expression engine supports numeric components, arithmetic, comparisons,
and a small function allowlist. It never uses Python `eval`. Scoped methods are
selected by date and highest priority; equal-priority ambiguity aborts refresh.
Higher-grain ratios always sum stored numerators and denominators before
division.

Every semantic value stores catalog, rule and method identity.
`meta.metric_application`, `meta.rule_application`, and workbook `_AUDIT`
record the same lineage by run. The governance workbook is generated directly
from the governed catalogs, preventing documentation/calculation drift.

Clipping, overlap unions, overnight handling, deduplication, identity, and
spell grouping remain tested engine primitives rather than editable formulas.

## Absence engine

Activities and wide StartEndTimes both normalize into `raw.schedule_shift`, but
`meta.source_file.source_variant` keeps them strictly separated. Activities
also produces `raw.schedule_event`. Data Source IDs is the primary operational
Agent ID. StartEndTimes is the preferred plan boundary. LILO and Agent Status
are the actual evidence. Activities is the final correction ledger. When the
dedicated StartEndTimes export is absent for an agent/day, a successfully
parsed Activities Shift Assignment is the explicit plan-boundary fallback and
raises a visible review finding.

The absence engine:

1. selects the active schedule version per Agent ID/day;
2. derives no-show/late/early from LILO plus active status evidence;
3. derives mid-shift logged-off/unavailable gaps from exclusive status states;
4. clips and unions only those observed gaps within the schedule;
5. uses the matching final Verint activity only to classify/reconcile the gap;
6. unions intervals separately for absence, vacation, unpaid and shrinkage;
7. caps planned net minutes at the configured standard day;
8. groups consecutive absence days into spells and calculates Bradford;
9. surfaces uncorrected gaps and Verint-only activities for review.

The final Verint ledger also cross-checks operational attendance completeness.
A completed working shift with neither a final non-working code nor reliable
Agent Status/LILO evidence is `UNCODED_EMPTY_SHIFT`. An evidence-backed gap
with no code is `UNCORRECTED_OBSERVED_GAP`; incomplete Verint coverage is
`PARTIAL_CORRECTION_REVIEW`; a code without a matching observed gap is
`VERINT_WITHOUT_OBSERVED_GAP`; an unfinished shift is `PROVISIONAL_DAY`.
All of these block final-ready status but do not invent a payroll absence type.
Headline final-absence ratios use only `CLEAR` and `ABSENCE_RECORDED` rows, so
exceptions and current-day provisional shifts cannot dilute the percentage.

An uncorrected observed gap is never silently assigned a sickness/vacation
reason. A corrected Verint activity can supply that final business category,
but it cannot create the underlying attendance gap.

## Service model

`mart.service_interval` contains both raw additive components and calculated
KPIs. Service availability has one unambiguous meaning:

```text
answered / offered
```

The model also retains gross SL and short-abandon-adjusted SL. Queue scopes
choose the reported profile, while both variants remain available for audit.
AHT is weighted from handled seconds divided by answered contacts. Higher-grain
reports use ratios of summed components, never averages of interval percentages.

## Agent Status without adherence

Agent Status is enabled by default because it is observed attendance evidence.
It is streamed in bounded batches and indexed by agent/date so a monthly file
does not require repeated full scans. Overlapping states are clipped to the
shift; the newest starting state wins for each exclusive segment.

No conformance percentage, out-of-adherence result, or RTA result is built.
Those legacy tables stay empty for database compatibility.

Agent Status filenames may describe one date, a date range, or full history.
Every row's `Status Start Date and Time` determines its business date.

## Forecast boundary

Verint Forecast contributes only forecast/required values. Exported actual
fields are discarded. APBE/APFR/APDE contributes actual performance. Both remain
joined only through the reviewed `config/queue_mapping.csv`. The raw source
queue/LOB and the mapped detailed/comparison scopes are all retained. Volume-only
forecast exports are valid; absent forecast measures remain NULL.

Forecast has two governed presentation grains. `mart.forecast_interval` keeps
the native Verint interval (15 minutes for new exports; historical hourly files
remain valid) for capacity planning and clean export. `mart.forecast_hour` is a
derived Flash/reporting rollup. Within an hour, volume is additive, FTE and
headcount are averaged because they are staffing levels, and SL/AHT are weighted
by forecast volume (or averaged when all interval volumes are zero). The source
grain and number of rolled intervals remain on the hourly row for audit.

## Capacity planning

Service profiles own explicit parallel mappings from `service_scopes` to
`staffing_lobs`. Staffing expands the complete selected period at 15-minute
grain. A native 15-minute forecast is used directly; an older hourly forecast
is safely spread across its four quarters, with required FTE repeated as a
level and volume divided across the quarters. Completed intervals preserve
observed and productive FTE; future intervals compare Verint required FTE with
net scheduled FTE after approved PTO and effective Away. If forecast demand
exists but no staffing row exists, a zero-schedule quarter-hour is created so
an empty roster cannot hide a shortage. `WEEKLY_PLAN` rolls additive FTE-hours
up by ISO week, management LOB, roster LOB, and language.

Realisations iterates every active service profile by default. Each profile
uses its configured SLA method and target. The dashboard never blends different
LOB service-level contracts into one synthetic overall SL.

## Upgrades

Current releases use SQLite and do not convert or open v0.1 DuckDB data.
Install the portable release in a new folder, point it at the same untouched
source root, and let it rebuild SQLite. Preserve the entire old folder. Saved
Attendance Review is a one-way injection queue. Re-exported Verint Activities
are the automatic reconciliation signal; the workbook is never imported back.

Within the SQLite generation, migrations are additive and never edited after
release. Config upgrades create a timestamped TOML backup. Database upgrades
use an online pre-migration backup and integrity checks.
