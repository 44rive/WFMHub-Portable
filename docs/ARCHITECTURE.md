# Architecture

## The stable extension seam

Every feature follows the same four layers:

```text
source adapter  ->  raw table  ->  core model  ->  report mart / Excel sheet
```

Adding a forecast KPI, a new queue feed, or another attendance rule does not
require rebuilding the whole hub. Add or extend one adapter, write a versioned
migration, materialize the relevant mart and add its report columns. Existing
extracts, decisions and database history remain intact.

## Portable installation

```text
WFMHub/
├── WFMHub.cmd                 interactive daily menu
├── SETUP.cmd                  first-run setup
├── runtime/                   embedded CPython and vendored packages
├── app/wfmhub/                application code
├── app/sql/migrations/        versioned DuckDB schema
├── config/default.toml        shipped defaults
├── config/wfmhub.toml         user configuration; preserved on upgrade
├── database/wfm.duckdb        durable local database; preserved on upgrade
├── input/                     persistent human inputs
├── output/                    finished reports
├── logs/                      daily logs
└── backups/                   recoverable database copies
```

The release builder packages CPython 3.13 x64, DuckDB 1.4.5 LTS, all Python
dependencies and the required Microsoft C++ runtime DLLs app-local beside
Python. The work machine does not run `pip`, does not require an installed
Python and does not need an administrator install. WFMHub is a single-writer
application: only one refresh may write to the DuckDB file at a time. A lock
file prevents accidental concurrent writes.

## Schemas and grains

| Schema/table | Grain |
|---|---|
| `meta.source_file` | One immutable content fingerprint per discovered file version |
| `meta.refresh_run` | One refresh attempt |
| `meta.quality_issue` | One current detected issue |
| `raw.fte_agent` | One FTE Agent-sheet row |
| `raw.schedule_shift` | One Verint schedule row/interval |
| `raw.schedule_event` | One parsed Verint event interval |
| `raw.lilo` | One Storm LILO roster row |
| `raw.agent_status` | One Storm status interval |
| `raw.forecast_interval` | One forecast queue and interval |
| `raw.queue_actual` | One Storm queue and 15-minute interval |
| `core.dim_agent` | One operational Agent ID |
| `core.correction_action` | One persistent human decision per Correction ID |
| `mart.attendance_agent_day` | One scheduled Agent ID and day |
| `mart.conformance_agent_day` | One worked-time result per scheduled Agent ID and day |
| `mart.correction_candidate` | One stable, non-overlapping detected interval or day review |
| `mart.rta_snapshot` | One scheduled agent at the newest status snapshot |
| `mart.forecast_hour` | One Verint forecast queue/hour |
| `mart.intraday_queue_interval` | One Storm actual queue/15-minute interval |

## File registry and incremental refresh

1. Discover configured files without modifying them.
2. Calculate SHA-256, file size and modified time.
3. If the same path and hash already loaded successfully, skip it.
4. Parse and validate a changed/new file completely.
5. In one transaction, deactivate the older version of that path, register the
   new version, and append its raw rows.
6. If parsing fails, record the failed version and leave the prior successful
   version active.
7. Rebuild the selected-period marts from active source versions.

Deleting an extract from its folder does not silently delete its already-loaded
database rows. Replacing a file at the same path creates an audited new version.
Overlapping schedule, forecast and queue snapshots choose the newest file at
their business grain.

## Identity

Verint `Data Source IDs` and Storm `Agent ID` are the operational Agent ID.
Blank, dash and placeholder IDs are quarantined as quality issues. FTE enriches
an Agent ID with organisation fields when its `Client ID` matches directly.
Names are never used as a silent production join.

## Attendance gates

The order is deliberate:

1. Required LILO source date(s) loaded?
2. Schedule parsed and Agent ID valid?
3. Off, planned absence or non-phone plan?
4. Agent ID ever present in LILO?
5. Daily LILO roster row present?
6. Both boundaries blank, one boundary blank, or usable boundaries?
7. Compare the usable first login/last logout with the scheduled interval.

A no-show therefore means: the required file exists, the Agent ID row exists,
both timestamps are blank, the assignment is work, and a planned absence does
not cover it. A missing file, missing roster row, identity issue or incomplete
punch is never promoted into a no-show.

Overnight shifts require every calendar-date LILO file touched by the shift.
Candidate first/last boundaries use a four-hour window around the shift. Planned
absence and planned adjustment intervals are physically subtracted from late and
early corrections, so displayed correction endpoints equal their gap minutes.

## Agent Status and conformance

Status intervals are clipped to the scheduled shift. Overlapping minute-rounded
transitions form one exclusive timeline: the newest state wins the overlap, so
category totals cannot exceed covered time. Agent Status becomes the conformance
basis only when the union covers the configured percentage of the shift (80% by
default). Otherwise the hub uses the explicitly labelled LILO span or `None`.

Mid-shift Logged Off and Unavailable candidates are produced only with passing
status coverage and only inside first-login to last-logout. Planned absence,
adjustment, lunch and break intervals are subtracted before a candidate is kept.

## Forecast and intraday boundary

Verint Forecast contributes only `For` and `Req` fields. Its exported `Act`
fields are intentionally discarded. Storm APBE/APFR contributes actual queue
performance. The two are shown separately until an approved queue/LOB scope
mapping is configured; the hub never invents that mapping.

## Upgrades

Release upgrades replace `runtime/`, `app/`, launchers and documentation. They
preserve `config/`, `database/`, `input/`, `output/`, `logs/` and `backups/`.
Schema changes are additive versioned SQL migrations. Create a database backup
before applying a release with new migrations.
