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
├── UPGRADE.cmd                 adopt an older portable database into a new folder
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

Program releases never ship a database or user configuration. An in-place
release keeps `_system/database/wfm.sqlite3`; `SETUP.cmd` backs it up when a
numbered migration is pending and applies only the missing SQL files. When a
release is extracted into a new folder, `UPGRADE.cmd` uses the SQLite backup API
to adopt the previous database, copies user-owned config/reports/custom jobs
without overwriting a destination file, runs `quick_check`, and then applies
only missing migrations. Historical extracts are not re-ingested merely because
application code or a report template changed.

The release ships a versioned PBIP source project under
`_system/templates/powerbi/WFMHub BI`. `POWERBI.cmd` installs it under
`Reports/Power BI`, rewrites only the `HubRoot` M parameter, and opens it in
Power BI Desktop. A newer project contract archives the installed project
before replacement. The PBIP model imports only `Feed/PowerBI`; it never opens
SQLite or raw extracts.

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
| `raw.queue_actual` | Retired AP compatibility storage; never discovered or refreshed |
| `raw.call_leg` | One admitted, typed Call-by-Call source leg |
| `raw.bonus_import` | One immutable Bonus Matrix content hash/version |
| `raw.bonus_agent_month` | One imported agent/month input row |
| `raw.bonus_kpi_rule` | One imported population/KPI threshold rule |
| `raw.bonus_policy` | One imported policy decision row |
| `core.clean_call_leg` | Deduplicated active call-leg view by stable Call Key |
| `core.dim_agent` | One operational Agent ID |
| `core.correction_action` | Legacy compatibility storage; not written by the current menu/CLI |
| `core.pcs_coaching_action` | Legacy compatibility table; permanent PCS coaching stays in the tracker |
| `raw.fte_time_off` | One governed PTO/Away register row from the standard FTE workbook |
| `mart.attendance_agent_day` | One scheduled Agent ID/day |
| `mart.conformance_agent_day` | Legacy compatibility table; empty in v0.5 |
| `mart.correction_candidate` | One exact residual Agent Status/LILO gap after final-Verint subtraction |
| `mart.correction_residual_segment` | One exact residual interval with final-Verint reconciliation state |
| `mart.staffing_interval` | One 15-minute roster LOB/language staffing interval |
| `mart.shift_timeline_segment` | One exact planned-versus-observed timeline segment |
| `mart.planned_time_off_segment` | One schedule-clipped, non-overlapping PTO/Away interval |
| `mart.rta_snapshot` | Legacy compatibility table; empty in v0.5 |
| `mart.verint_final_exception` | One deterministic final-Verint completeness exception |
| `mart.forecast_hour` | One raw forecast queue/hour plus mapped scopes |
| `mart.intraday_queue_interval` | Retired AP compatibility table; cleared during refresh |
| `mart.agent_pcs_day` | One admitted Agent ID/day with call and PCS measures |
| `mart.verint_final_absence_event` | One mapped Verint Activities final component interval |
| `mart.verint_final_absence_agent_day` | One Activities-final absence/shrinkage result per Agent ID/day |
| `mart.absence_event` | One reviewed gap or PTO/Away exact component interval |
| `mart.absence_agent_day` | One reviewed absence/vacation/shrinkage result per Agent ID/day |
| `mart.service_interval` | Stable semantic projection of Call-by-Call queue/hour counters |
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
| `pcs` | `Reports/PCS Live Tracker.xlsx` | Permanent direct-CSV Power Query performance and coaching tracker |
| `bonus` | `Reports/Bonus Management.xlsx` | Imported Bonus Matrix result and release controls |
| `service` | `Reports/RTM Daily Control.xlsx` | Same-day service, attendance call actions, and queue drivers for RSA NL/BE and Ford NL/OEM |
| `realisations` | `Reports/Realisations.xlsx` | All mapped LOB actual/forecast, service, staffing, absence and shrinkage results |
| `staffing` | `Reports/Staffing Gaps.xlsx` | Full-period actual staffing control and future capacity planning |
| `attendance` | `_system/legacy_reports/Legacy Attendance Callout.xlsx` | Compatibility-only callout builder; absent from the normal menu |
| `corrections` | `Reports/Attendance Review.xlsx` | Selected-period residual gaps, schedule/actual evidence, and break/meal control |
| `absence` | `Reports/Final Absenteeism.xlsx` | Verint Activities final absence/shrinkage ledger and completeness review |

Products share the same visual identity but use purpose-specific layouts.
The approved identity is versioned as `WFMHUB-DESIGN` in
`docs/REPORT_DESIGN_SYSTEM.md`; its code tokens live only in
`src/wfmhub/design.py`. Operational first screens use a compact title/status
header, an honest scope strip, four headline cards, at most two decision charts,
and one filterable action or reconciliation table. Generated snapshots never
display fake selectors.
The PCS operational update is deliberately domain-scoped: FTE and Call-by-Call
are ingested, then only the employee dimension and PCS mart are rebuilt. Python
writes six fixed, governed CSV feeds and creates the permanent tracker only
when it is missing or its versioned contract changes. It never opens Excel
during a normal update and never replaces a same-version tracker.

RTM Daily Control begins with `CONTROL`, then provides four purpose-built LOB
sheets. Each LOB combines the validated hourly service view with its own
reconciling attendance/call list. `ISSUES & DRIVERS` replaces the separate
exceptions and queue-diagnosis surfaces. Attendance Review pairs a scheduled
band directly above the actual evidence band for every residual gap. Exact
`EVIDENCE` is hidden audit support. Each paired
schedule/actual case has a short visual separator. Explicit Agent Status
`Meal Aux` remains Lunch even when the LILO logout boundary falls inside it.
`BREAK & MEAL`
aggregates completed-day Agent Status
spells and never revives adherence. Standalone Attendance Callout, legacy
`operations`, and legacy `quality_pcs` remain callable under
`_system/legacy_reports` but are absent from the menu.

PCS has a deliberate lightweight-feed lifecycle. SQLite and Python calculate
all ratios from additive call-leg counters, then atomically replace LOB, agent,
daily, period/scope, and coaching-opportunity CSV products. Power Query only
validates columns/types and transports those products to native Excel tables.
`OVERVIEW` contains fixed all-scope current-MTD cards, LOB comparison, daily
trend and agent scorecard. `PERFORMANCE` is the full period/scope table for
native filters or slicers. `COACHING` contains the queried exact-call queue and
the separate permanent human-owned action table. There is no raw `DATA`
worksheet, Data Model, Power Pivot, ODBC driver, macro, or dynamic-array
metadata. Desktop Excel automation is used only for the explicit one-time
query installation/repair command.

Final Absenteeism uses the same collaboration boundary. Power Query may replace
`tblAbsenceData`, `tblActionQueue`, and `tblActivityDetail` from stable CSVs;
it must never load into permanent `tblActions`. `TEAM_VIEW` and
`COMPONENT_VIEW` read the refreshed tables directly.

Power BI receives fixed additive facts under `Feed\PowerBI`. Python/SQLite owns
classification and counters; Power BI owns relationships, measures, slicers and
visuals. `POWERBI_MANIFEST_CURRENT.csv` is written last as the refresh receipt.

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
problem. Forecast is queue data and bypasses the agent gate. Call-by-Call also
admits exact reviewed queue-map matches so abandoned demand is preserved.
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
Status, Call-by-Call and Forecast use row fields. LILO prefers a
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
opportunity. Exact calls appear on `COACHING`; reviewers copy Coaching Key and
Call ID into the editable action table on that same sheet and complete its five
action fields. Actions Rate is unique completed Coaching Keys divided by all
low-score opportunities. Coaching decisions remain only in the permanent
tracker and never enter SQLite. The PCS mart rebuild window expands to retain configured history so
daily, MTD and prior-period comparisons remain available. See
[PCS logic](PCS_LOGIC.md).

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
5. Build one stable Gap ID from date, agent, issue and exact boundaries.
6. Clip mapped final Verint Activities to the same shift and subtract their
   exact overlap from each raw gap.
7. Publish only remaining fragments; a fully covered gap disappears.

“No show” requires a completed scheduled working shift plus positive evidence:
either a loaded daily LILO row with both boundaries blank, or sufficient Agent
Status coverage in which every observed interval is Logged Off. Missing files,
missing rows, and incomplete evidence are never a no-show.

A logout followed by a later active status is an internal gap. The later return
extends the actual presence boundary, so the earlier logout cannot become an
early leave. Leading logout intervals are late; trailing logout intervals are
early leave only after the scheduled end; multiple reconnect cycles remain
separate gaps unless they are within the configured merge tolerance.

Verint activity intervals never create attendance evidence. They are read only
after gap detection to prove final coding overlap. The raw gap remains
traceable while the user-facing backlog contains only unresolved fragments.

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

Activities and wide StartEndTimes can both normalize into `raw.schedule_shift`,
but `meta.source_file.source_variant` keeps them separated. Data Source IDs is
the primary operational Agent ID. StartEndTimes is the preferred plan boundary.
LILO and Agent Status are actual evidence. Activities intervals never become
attendance evidence. They feed only the final post-day absence/shrinkage ledger.
When StartEndTimes is absent for an agent/day, a parsed Activities Shift
Assignment is also an explicit boundary-only fallback and raises a review finding.

The absence engine:

1. selects the active schedule version per Agent ID/day;
2. derives no-show/late/early from LILO plus active status evidence;
3. derives mid-shift logged-off/unavailable gaps from exclusive status states;
4. clips and unions only those observed gaps within the schedule;
5. subtracts exact overlap already coded in final Verint Activities;
6. keeps any remaining observed gap provisional rather than inventing a reason;
7. unions governed PTO/Away separately for absence, vacation and unpaid totals;
8. caps planned net minutes at the configured standard day;
9. groups consecutive absence days into spells and calculates Bradford;
10. surfaces residual gaps and missing evidence for review.

Separately, the final-absence builder clips mapped Verint Activities to the
planned shift, unions overlapping components, caps ratios to planned net time,
and flags unmapped, empty, provisional, missing-planned-time-off and partial
correction cases. This separation prevents a final coding export from proving
same-day presence.

An unresolved observed gap is `PENDING_VERINT`; an unfinished shift is
`PROVISIONAL_DAY`. Both block final-ready status without inventing a reason. Headline ratios use
only `CLEAR` and `ABSENCE_RECORDED` rows, so open/provisional rows cannot dilute
the percentage.

An observed gap is never silently assigned a sickness/vacation reason. Final
business categories come from mapped Verint Activities or governed PTO/Away.

## Service model

`mart.service_interval` is the stable semantic projection of
`mart.call_service_hour`. Reports use the Storm business definition:

```text
answered_within_target /
(offered - abandoned_within_target)
```

The threshold clock is Call-by-Call total queue wait plus ringing duration.
The default target is 30 seconds. `abandoned_within_target` means lost from 5
seconds up to that target; losses below 5 seconds remain in the denominator.
Daily headline counters reset at midnight, independently of visible operating
hours.
Storm Routed Rate is answered divided by every entered queue entry. Queue
profiles use exact reviewed queue allowlists for the four Flashes.
Provider, language and regional labels never add or remove a queue implicitly.
All Flash tables display 00:00-23:00. Staffing and absence columns use the
profile's configured staffing LOB rather than inventing queue-level ownership.
AHT is weighted from handled seconds divided by answered contacts. Higher-grain
reports use ratios of summed components, never averages of interval percentages.
The RTM tables expose the additive answered and answered-within-target counters
as `Volume Handled` and `Handled in SL` so their relationship to TSL is visible.

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
fields are discarded. Call-by-Call contributes all actual performance. The
sources are joined only through reviewed `config/queue_mapping.csv`. The raw source
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

Current releases use SQLite and do not convert or open v0.1 DuckDB data. For an
in-place release, `SETUP.cmd` preserves the existing SQLite file and applies
only missing migrations. For a release extracted into a new folder,
`UPGRADE.cmd` adopts the prior SQLite database through its backup API, verifies
it, copies user-owned configuration/reports/custom jobs without overwriting,
and applies only missing migrations. A code or report-template release does
not require historical extracts to be rebuilt.
Attendance Review is a read-only residual reconciliation workbook. SQLite keeps
authority over exact evidence; corrections are made in Verint and arrive in a
later Activities extract.

Within the SQLite generation, migrations are additive and never edited after
release. Config upgrades create a timestamped TOML backup. Database upgrades
use an online pre-migration backup and integrity checks.
