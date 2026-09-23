# WFMHub product contract

This file is the authoritative description of WFMHub. Read it before changing
code, formulas, mappings, source contracts, reports, or the local web app.

## Product

WFMHub is Anass ASSRI's portable Workforce Management control system. It keeps
source extracts unchanged, admits only the governed employee population,
stores typed evidence in local SQLite, calculates auditable WFM models, and
publishes operational workbooks, clean CSV exports, and a localhost manager
workbench.

The product serves the WFM cycle:

1. demand and service monitoring;
2. intraday attendance control;
3. schedule-versus-observed review;
4. forward staffing preparation;
5. realization and absence review;
6. governed analysis and export.

The SQLite database is durable user state. Reports are replaceable outputs.
Source extracts are evidence and are never edited.

## Runtime and launchers

The Windows release is a no-admin ZIP containing official embeddable CPython,
SQLite from that runtime, and pure-Python Excel libraries.

- `SETUP.cmd` validates the runtime, records the extract root, and creates the
  database.
- `WFMHub.cmd` opens the terminal menu.
- `WEBAPP.cmd` opens the local manager workbench at `127.0.0.1`.
- `UPGRADE.cmd` copies durable state from a previous extracted release into a
  newly extracted release using SQLite's backup API.

`WFMHub.cmd` remains the operational controller. The web app reads the same
database and can launch bounded refresh, analysis, export, and report jobs.

## Authoritative sources

All locations are relative to `paths.source_root` in `config/wfmhub.toml`.

| Source | Default location | Authority |
|---|---|---|
| FTE Count | `FTE/FTE Count.xlsx` | employee master, Active/Leaver eligibility, organization, FTE, PTO and Away |
| Verint StartEndTimes | `Verint/Schedules & Activities/*.txt` | published shift start/end and assignment |
| Verint Activities | same folder | finalized post-day absence and shrinkage activities only |
| Agent Status | `Storm/Agent Status/*.csv` | primary observed attendance and status intervals |
| LILO | `Storm/LILO/*.csv` | fallback login/logout boundary evidence |
| Call by Call | `Storm/Call by Call/*.csv` | service demand, handling counters and PCS call legs |
| Verint forecast/requirement | `Verint/Forecast/*.txt` | 15-minute Volume and required/forecast FTE by Verint Staff Type |
| Bonus Matrix | selected manually | governed monthly bonus source imported read-only |

Files may cover one or many dates. Business dates are read from rows, never
inferred as the only truth from filenames.

### Employee scope

Client ID is the employee key and is always text. FTE rows are eligible when:

- status is `Active`; or
- status is `Leaver`, an end date exists, and the evidence date is on or before
  that end date.

Agent-scoped evidence is accepted by exact Client ID. A unique normalized name
may gate a record into scope when needed, but a populated operational source ID
is preserved. Missing or ambiguous employee matches stay visible as quality
issues; they are not silently assigned.

## Business boundaries

### Service and forecast are different grains

Service level is calculated from Call-by-Call queue evidence using the exact
effective Flash allowlist in `config/service_profiles.toml`. Forecast Volume
and FTE Requirement are Staff Type series; they are not queue observations.

RSA Belgium has one combined service-level headline, while capacity remains
split into planning groups `RSA BE FR` and `RSA BE VL`. Shared Ford VL/DE call
queues are included in the business-approved Ford Netherlands and RSA Belgium
Flash scopes where configured; staffing ownership is still determined from
employee/schedule mappings rather than inferred from the queue name.

Every Flash shows 00:00–23:00. Current-day actuals stop at the latest loaded
Call-by-Call interval; future cells stay blank.

### Attendance

Agent Status is primary evidence. LILO is a fallback when Agent Status cannot
establish reliable boundaries. PTO and Away are overlays from FTE Count and
must prevent a planned absence from becoming a no-show.

Same-day states are operational: due, present, proven no-show, possible
no-show/unknown, late, offline now, and early leave only after the shift is
complete. A person who returns after an early disconnect remains present while
the exact internal gap remains reviewable.

RTM No Show HC is restricted to completed shifts with explicit agent-specific
absence evidence. An in-progress agent who has not appeared is a named
not-seen callout / Possible No Show, not a confirmed No Show. Its hourly
staffing columns are distinct-agent snapshots: scheduled, logged, productive,
available, unavailable/other AUX, and BO. Available and BO are subsets of
productive, not additive headcounts. Below-target queue-hour rows list named
same-LOB agents with offline, AUX, BO, or missing-evidence states as context,
never as proven individual causes of service loss.

Attendance Review is read-only. It shows schedule-over-actual timelines and
only residual gap fragments not already explained by finalized Verint
Activities. Once corrected Activities are loaded, matched fragments disappear.
It does not store manual correction decisions.

### Absence and shrinkage

Observed attendance gaps are provisional. Final absence and shrinkage use
Verint Activities, effective activity rules, the published schedule, and PTO /
Away overlays. An empty scheduled shift without sufficient final evidence is a
review exception, never an automatic zero.

### Staffing

Capacity is calculated at 15-minute grain by Management LOB, Planning Group,
and Staff Type. Requirement comes from the Staff Type forecast extract.
Scheduled capacity comes from published schedules and active FTE. PTO/Away is
deducted explicitly. Shortage is `max(required FTE - net scheduled FTE, 0)`.

### PCS

PCS is calculated from deduplicated, in-scope Call-by-Call legs. Numerators and
denominators are aggregated separately; percentages and averages are never
averaged from row-level percentages. The permanent `PCS Live Tracker.xlsx`
contains the collaborative coaching table. WFMHub updates fixed CSV feeds;
Excel Power Query refreshes presentation tables. Rebuilding the tracker first
reads and preserves keyed coaching rows from that same tracker.

### Bonus and realisations

`Bonus Management.xlsx` is a permanent, formula-driven monthly working file.
Its editable Policy Decisions, KPI Config, and Raw Data sheets follow the
supplied Bonus Matrix v1.2 structure; results and dashboards recalculate in
Excel. WFMHub does not overwrite a current canonical workbook on each refresh.

`Realisations.xlsx` includes actual-versus-forecast service and staffing plus
`HOURS_BY_LOB` and `HOURS_BY_AGENT` allocation of every published-shift minute
from the exclusive Agent Status timeline. LILO-only, unknown evidence, future,
and PTO/Away remain distinct categories rather than guessed productivity.

## Configuration ownership

Defaults ship as `config/default_*`. First setup creates editable user copies
without modifying the defaults.

| File | Owns |
|---|---|
| `wfmhub.toml` | paths, sources, period, runtime thresholds, enabled domains |
| `wfm_rules.toml` | activity/status classification and attendance evidence rules |
| `metric_catalog.toml` | KPI formulas, targets, units, scopes and effective dates |
| `analytics_rules.toml` | deterministic finding thresholds |
| `report_catalog.toml` | report composition only |
| `queue_mapping.csv` | source queue and forecast-prefix mapping |
| `service_profiles.toml` | exact service Flash scopes and queue allowlists |
| `capacity_mapping.csv` | workforce LOB, assignment, Staff Type and planning-group mapping |

No workbook, browser page, or Python function should invent a competing KPI
formula. Change calculations in the effective-dated catalogs, validate them,
and keep raw components available for reconciliation.

## Current outputs

Operational products:

- `RTM Daily Control.xlsx`: service flashes plus attendance call/pulse evidence;
- `Attendance Review.xlsx`: completed-day schedule integrity, exact residual
  gaps, shift timelines, and break/meal overrun controls;
- `PCS Live Tracker.xlsx`: permanent PCS performance and coaching workspace;
- local Manager Workbench: manager desk, service, attendance, integrity,
  capacity, realizations, absence, readiness, mappings, reports, and jobs.

Controlled products under continued business validation:

- `Staffing Preparation.xlsx`;
- `Realisations.xlsx`;
- `Final Absenteeism & Shrinkage.xlsx`;
- `Bonus Management.xlsx`.

Clean exports in `Feed/` are generated outputs and can be rebuilt. `Reports/`
contains current workbooks and timestamped prior copies under `Archive/`.

## Code map

- `src/wfmhub/ingestion.py`: source discovery, parsers, immutable raw loading;
- `src/wfmhub/models.py`: employee, attendance, staffing, forecast, service,
  PCS, absence, quality, and analysis marts;
- `src/wfmhub/metrics.py`, `rules.py`, `mapping.py`, `capacity_mapping.py`,
  `service_profiles.py`: governed configuration readers;
- `src/wfmhub/service_flash.py`, `decision_products.py`, `pcs_tracker.py`,
  `bonus.py`: workbook products;
- `src/wfmhub/web_data.py`, `webapp.py`, `web/`: localhost workbench;
- `src/wfmhub/exports.py`, `shared_feeds.py`: clean data contracts;
- `sql/migrations/001_schema.sql`: complete current SQLite schema;
- `tests/`: synthetic contract, calculation, workbook, and integration tests;
- `roadmap/`: isolated future projects, never active runtime dependencies.

## Invariants

1. Never modify an extract.
2. Never commit real employee, schedule, absence, call, or customer data.
3. Never guess queue membership or employee identity.
4. Preserve Client IDs as text.
5. Keep raw counts beside calculated KPIs.
6. Treat missing evidence as a visible state, not zero.
7. Do not load large raw datasets into visible Excel sheets.
8. Keep the database on a local writable disk; SharePoint/OneDrive may hold
   shared output workbooks, not the live SQLite file.
9. A release archive contains code and templates, never user config, data,
   database files, or generated reports.
10. A change is complete only after a fresh-schema test, unit tests, launcher
    checks, and workbook ZIP/XML validation pass.

## Documentation map

- [Operations](docs/OPERATIONS.md)
- [Architecture](docs/ARCHITECTURE.md)
- [Business rules](docs/BUSINESS_RULES.md)
- [Data contracts](docs/DATA_CONTRACTS.md)
- [Configuration](docs/CONFIGURATION.md)
- [PCS workflow](docs/PCS.md)
- [Bonus Management](docs/BONUS.md)
- [Design system](docs/DESIGN_SYSTEM.md)
- [Development](docs/DEVELOPMENT.md)
- [Roadmap](docs/ROADMAP.md)
