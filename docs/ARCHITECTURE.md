# Architecture

## Stable extension seam

Every feature follows the same layers:

```text
untouched source -> parser/scope gate -> raw table -> core model -> mart -> Excel
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
├── config/default.toml         shipped defaults
├── config/wfmhub.toml          user configuration
├── database/wfm.sqlite3        durable SQLite hub
├── input/                      persistent human inputs
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
| `raw.fte_agent` | One FTE Agent-sheet row |
| `raw.schedule_shift` | One admitted Verint schedule row |
| `raw.schedule_event` | One parsed event interval belonging to an admitted shift |
| `raw.lilo` | One admitted Storm LILO row |
| `raw.agent_status` | One admitted status interval |
| `raw.forecast_interval` | One queue/forecast interval; not agent-scoped |
| `raw.queue_actual` | One queue/15-minute actual interval; not agent-scoped |
| `core.dim_agent` | One operational Agent ID |
| `core.correction_action` | One human decision per stable Correction ID |
| `mart.attendance_agent_day` | One scheduled Agent ID/day |
| `mart.conformance_agent_day` | One conformance result per scheduled Agent ID/day |
| `mart.correction_candidate` | One non-overlapping candidate interval/day review |
| `mart.rta_snapshot` | One scheduled agent at the newest admitted status timestamp |
| `mart.forecast_hour` | One forecast queue/hour |
| `mart.intraday_queue_interval` | One actual queue/15-minute interval |
| `mart.source_health` | One configured source family |

## Agent scope and identity

FTE is the authority for “our agents.” Before an agent-level row enters the
active raw layer:

1. Normalize the source Agent ID.
2. Keep it when the ID exists in active FTE `Client ID` values.
3. Otherwise normalize accents, case, punctuation, and whitespace in the name.
4. Keep it only when that name maps to exactly one FTE Agent ID.
5. Preserve a populated operational source ID. In particular, Verint `Data
   Source IDs` remains the schedule Agent ID.
6. Exclude everything else and count it as outside roster.

A populated unmatched ID can therefore be admitted by a unique roster name,
but an ambiguous or missing name cannot. This is a scope decision, not fuzzy
matching.

The scope has a deterministic fingerprint. If FTE changes, the same untouched
schedule/LILO/status file is reprocessed against the new roster. This prevents
both stale worldwide rows and the “new agent missing from an unchanged file”
problem. Forecast and APBE/APFR are queue data, so they bypass the agent gate.

## Incremental and atomic refresh

1. Discover configured files read-only.
2. Calculate SHA-256, size, modified time, and agent-scope fingerprint.
3. Skip an already-active successful match.
4. Parse a new version; stream large LILO CSVs in bounded batches.
5. In one file transaction, append immutable raw rows, deactivate the previous
   path version, and activate the new version.
6. If parsing fails, roll back its raw rows and leave the previous good version
   active. Retrying the same fingerprint is supported.
7. Rebuild all selected-period models inside one savepoint. Any failure restores
   every previous mart.

A same-path A→B→A change reactivates A's immutable rows rather than duplicating
them. Deleting a physical extract does not silently erase loaded history.

## Attendance and correction gates

The order is deliberate:

1. Required LILO date(s) loaded?
2. Schedule and Agent ID valid?
3. Off, planned absence, or non-phone plan?
4. Agent identity ever seen in admitted LILO?
5. Daily LILO roster row present?
6. Both boundaries blank, one blank, or usable pair?
7. Compare the usable first/last boundary with schedule.

“No show” requires a loaded daily LILO file, the admitted agent row, both times
blank, a Work assignment, and no covering planned absence. Missing file,
missing roster row, identity mismatch, or incomplete LILO is never a no-show.

Overnight shifts require every touched LILO date. Planned absence/adjustment is
physically subtracted from late and early intervals, so displayed endpoints and
gap minutes agree.

## Status, conformance, and RTA

Status intervals are clipped to the scheduled shift. When overlaps exist, the
newest state wins; category totals cannot exceed covered time. Agent Status is
the conformance basis only above configured coverage (80% by default), else the
explicitly labelled LILO span or `None` is used.

RTA uses the newest admitted status timestamp. If no status row passes agent
scope, RTA is empty and source health is `ERROR`; current time is never used as
a fake snapshot.

## Forecast boundary

Verint Forecast contributes only forecast/required values. Exported actual
fields are discarded. APBE/APFR contributes actual performance. Both remain
separate until a reviewed queue/LOB mapping is configured.

## Upgrades

v0.2 uses SQLite and does not convert or open v0.1 DuckDB data. Install v0.2 in
a new folder, point it at the same untouched source root, and let it rebuild
SQLite. Preserve the entire v0.1 folder. Saved Excel reports can re-import
correction decisions.

Within the SQLite generation, migrations are additive and never edited after
release. Config upgrades create a timestamped TOML backup. Database upgrades
use an online pre-migration backup and integrity checks.
