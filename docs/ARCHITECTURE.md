# Architecture

## Stable extension seam

Every feature follows the same layers:

```text
untouched source -> parser/scope gate -> raw table -> rule engine -> mart -> Excel
```

A new forecast KPI, queue feed, or attendance rule adds an adapter, a numbered
migration, a focused model, tests, and curated report output. Existing extract
files are never edited.

## Portable installation

```text
WFMHub/
├── WFMHub.cmd                  daily menu
├── SETUP.cmd                   system check and first setup
├── runtime/                    official embedded CPython + pure-Python packages
├── RUNTIME_MANIFEST.sha256     reviewed native-file hashes
├── app/wfmhub/                 application code
├── app/sql/migrations/         versioned SQL schema
├── templates/FTE Count.xlsx    blank standard roster workbook
├── config/default.toml         shipped defaults
├── config/wfmhub.toml          user configuration
├── config/default_rules.toml   shipped calculation defaults
├── config/wfm_rules.toml       editable/versioned business rulebook
├── config/default_queue_mapping.csv shipped mapping defaults
├── config/queue_mapping.csv    editable queue/file/scope mapping
├── database/wfm.sqlite3        durable SQLite hub
├── input/                      persistent human inputs
├── custom/                     Python and read-only SQL job templates
├── output/                     finished reports
├── logs/                       daily logs
└── backups/                    SQLite online backups
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
| `raw.fte_agent` | One FTE Agent-sheet row |
| `raw.schedule_shift` | One admitted Verint schedule row |
| `raw.schedule_event` | One parsed event interval belonging to an admitted shift |
| `raw.lilo` | One admitted Storm LILO row |
| `raw.agent_status` | One admitted status interval |
| `raw.forecast_interval` | One queue/forecast interval; not agent-scoped |
| `raw.queue_actual` | One queue/15-minute actual interval; not agent-scoped |
| `raw.call_leg` | One admitted, typed Call-by-Call source leg |
| `core.clean_call_leg` | Deduplicated active call-leg view by stable Call Key |
| `core.dim_agent` | One operational Agent ID |
| `core.correction_action` | One human decision per stable Correction ID |
| `mart.attendance_agent_day` | One scheduled Agent ID/day |
| `mart.conformance_agent_day` | Legacy compatibility table; empty in v0.5 |
| `mart.correction_candidate` | One observed LILO/status gap plus Verint-final check |
| `mart.correction_residual_segment` | One still-uncovered interval requiring Verint review/injection |
| `mart.staffing_interval` | One 15-minute roster LOB/language staffing interval |
| `mart.shift_timeline_segment` | One exact planned-versus-observed timeline segment |
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
| `mart.source_health` | One configured source family |

## Report packs

The shared SQLite hub can serve multiple workbooks without mixing their grains:

| Pack | Folder | Scope |
|---|---|---|
| `operations` | `output/operations` | Selected-day attendance calls, staffing gaps, and APDE service state |
| `corrections` | `output/corrections` | Latest evidence-complete-day residual gaps and shift timeline |
| `quality_pcs` | `output/quality_pcs` | Exact Q1 PCS, participation, day/team/month summaries and responses |
| `absence` | `output/absence` | Activities-only final absence ledger, evidence and activity rules |

The Yesterday Corrections workbook is the only workbook accepted for
correction-action import. Legacy Intraday and Executive Scorecard builders
remain internal compatibility code but are not registered or offered.
Raw call-by-call rows stay in SQLite; PCS reports receive summaries, trends and
bounded responses. Full clean details are explicit CSV/XLSX exports.

## Agent scope and identity

FTE is the authority for “our agents.” Before an agent-level row enters the
active raw layer:

1. Normalize the source Agent ID.
2. Keep it when the ID exists in active FTE `Client ID` values.
3. Otherwise normalize accents, case, punctuation, and whitespace in the name.
4. Keep it only when that name maps to exactly one FTE Agent ID.
5. Preserve a populated operational source ID. In particular, Verint `Data
   Source IDs` remains the schedule Agent ID.
6. Apply the same gate to Call-by-Call agent legs.
7. Exclude everything else and count it as outside roster.

A populated unmatched ID can therefore be admitted by a unique roster name,
but an ambiguous or missing name cannot. This is a scope decision, not fuzzy
matching.

The scope has a deterministic fingerprint. If FTE changes, the same untouched
schedule/LILO/status/call file is reprocessed against the new roster. This prevents
both stale worldwide rows and the “new agent missing from an unchanged file”
problem. Forecast and APBE/APFR/APDE are queue data, so they bypass the agent gate.

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

Call CSVs are FTE-scoped before storage. A stable Call Key combines call
references, direction, agent and timestamps; `core.clean_call_leg` selects the
newest active version across overlapping history extracts. The official PCS
contract reproduces `TOLEARN/PCS Report.xlsx`: keep inbound legs with an Agent
ID, accept Q1 only when its numeric value is one of the configured discrete
scores (default `1,2,3,4,5`), and calculate `sum(valid Q1) / count(valid Q1)`.
Counts `<=3` and `>3` remain counts. Participation is `inbound raw-Q1 nonblank /
inbound PCSStatus=1`; invalid raw answers stay in that numerator and are
separately counted. Q2 and Mode 2 are diagnostics only. Higher grains always
sum counters before dividing. See [PCS logic](PCS_LOGIC.md).

## Clean exports and Custom Lab

Clean exports stream selected-period results to UTF-8 CSV or bounded XLSX with
a manifest. Custom SQL is restricted to one SELECT/WITH statement on a
query-only connection. Custom Python receives the same read-only query context,
but is trusted executable local code and is not an operating-system sandbox.

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
2. Collect LILO and Agent Status evidence clipped to that boundary.
3. Use the earliest/last active evidence for late and early-leave detection.
4. Use exclusive Agent Status `Logged Off`/`Unavailable` intervals between those
   boundaries for mid-shift gap detection.
5. Build gaps before reading any final Verint activity.
6. Match each observed gap against the Activities-only final ledger.
7. Label it `CORRECTED`, `PARTIAL`, or `NOT_CORRECTED` without changing the
   observed interval.

“No show” requires a loaded daily LILO row with both boundaries blank, a
scheduled non-Off shift, and no active Agent Status evidence. Missing files,
missing rows, and incomplete evidence are never a no-show.

Final Verint activities are deliberately not subtracted from late, early or
status gaps. Doing so would hide the original problem immediately after it was
corrected, destroying the audit trail.

## Central rulebook and calculation audit

`config/wfm_rules.toml` is the canonical business definition. It contains
named, validated KPI expressions, activity categories/flags, standard-day
hours, service-level profiles and queue scopes. The expression engine supports
only numeric variables, arithmetic, and a small function allowlist. It never
uses Python `eval`.

Every absence and service row stores the active `rule_version` and
`rule_sha256`. `meta.rule_application` records the same identity by refresh run.
The KPI catalog workbook is generated from this rulebook; documentation and
calculation code therefore cannot silently drift.

Clipping, overlap unions, overnight handling, deduplication, identity, and
spell grouping remain tested engine primitives rather than editable formulas.

## Absence engine

Activities and wide StartEndTimes both normalize into `raw.schedule_shift`, but
`meta.source_file.source_variant` keeps them strictly separated. Activities
also produces `raw.schedule_event`. Data Source IDs is the primary operational
Agent ID. StartEndTimes is the only plan boundary. LILO and Agent Status are the
actual evidence. Activities is the final correction ledger.

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

## Upgrades

Current releases use SQLite and do not convert or open v0.1 DuckDB data.
Install the portable release in a new folder, point it at the same untouched
source root, and let it rebuild SQLite. Preserve the entire old folder. Saved
Yesterday Corrections workbooks can re-import correction decisions.

Within the SQLite generation, migrations are additive and never edited after
release. Config upgrades create a timestamped TOML backup. Database upgrades
use an online pre-migration backup and integrity checks.
