# WFMHub master product and business contract

Contract version: `2.3.0`
Applies to: WFMHub `0.35.0` and later
Last reviewed: `2026-09-14`

This is the single starting point for humans and coding assistants. Read it,
`AGENTS.md`, and the effective configuration before changing the Hub. Current
user instructions and dispatched code with passing tests override prose. Old
conversation summaries, archived reports, TOLEARN examples, prototype folders,
and unreachable builders are not current authority.

## Mission and operating model

WFMHub is a deterministic, portable WFM system for a restricted Windows work
machine. It leaves source extracts untouched, scopes rows to the effective FTE
roster, persists history in SQLite, calculates governed Python/SQL marts, and
serves a localhost Manager Workbench plus focused Excel reports. There is no
runtime AI, DuckDB, ODBC dependency, server database, Excel Data Model
requirement, or adherence KPI. `prompts/COPILOT_WFM_ANALYST.md` is only an
optional manual aid.

The normal pipeline is:

```text
untouched extracts -> validated raw/core SQLite -> governed marts
                   -> localhost Manager Workbench
                   -> focused Excel decision products + fixed collaboration feeds
```

The Update action refreshes the durable database and collaboration feeds. The
local console and report commands read the database; they do not parse raw files
independently. PCS remains its own permanent collaborative Excel tracker and is
deliberately outside the single-user Manager Workbench.

## Source authority

| Source | Governed purpose | Explicit exclusion |
|---|---|---|
| FTE Count `Agent` | Client ID, organisation, FTE, Active/Leaver scope | not attendance evidence |
| FTE Count `PTO` / `Away` | expected-work overlay and future net capacity | does not change published assignment grain |
| Verint StartEndTimes | published shift start/end and assignment | not observed presence |
| Storm Agent Status | primary observed attendance, interval staffing, AUX, break and meal evidence | not final payroll absence coding |
| Storm LILO | fallback/control when Agent Status evidence is insufficient | never overrides explicit Agent Status states |
| Verint Activities | final post-day absence/shrinkage ledger and exact correction overlap | never creates observed attendance |
| Verint Forecast / FTE Requirement | 15-minute Staff Type Volume and Full Time Equivalents Absolute Req | `Queue Name` is not a call queue |
| Storm Call by Call | service counters, Flash, call workload and PCS call legs | not workforce requirement |
| Bonus Matrix v1.2 | exact management bonus source/calculation contract | not a WFM attendance source |

Row dates are authoritative. Filename dates are fallback hints only, so all
source families support multi-day files. Client ID / Agent ID is text and is the
primary identity. Ambiguous IDs are rejected; a unique normalized-name match may
enrich organisation fields but must never silently pick one duplicate.

FTE eligibility is date-aware: `Active` is included; `Leaver` is included only
through a populated leave date; other statuses and undated leavers are excluded.
The rule applies to schedules, status, LILO and agent-scoped Call-by-Call/PCS
facts. Queue-scoped service demand retains every contact entering an exact
configured queue, including abandons and transfers handled outside the roster.

## Two grains that must never be mixed

Service and capacity have separate governed maps.

- Service grain: exact Call-by-Call queue allowlists from
  `config/service_profiles.toml`. Counters roll to Management LOB. RSA BE service
  level is one combined result across its exact FR and VL queues. A queue may
  intentionally belong to more than one management view when the reviewed
  service profile says so.
- Capacity grain: forecast filename + Verint Staff Type and roster LOB +
  published schedule assignment from `config/capacity_mapping.csv`. Capacity
  rolls through Staff Type -> Planning Group -> Management LOB.

Current planning roll-ups:

| Management LOB | Planning group / workforce identity | Examples of Staff Type |
|---|---|---|
| RSA NL | RSA NL / `RSA NL` | NL RSA FO, NL RSA BO Dispatch |
| RSA BE | RSA BE FR / `RSA FR` | BE RSA FO FR, BE RSA Dispatch FR |
| RSA BE | RSA BE VL / `RSA VL` | BE RSA FO VL, BE RSA Dispatch VL |
| FORD NL | FORD NL / `Ford Dutch` | NL RSA Ford Level 1/2 |
| OEM | Ford FR / `OEM FR` | FR RSA Ford Level 1/2, FR RSA ACM |

PTO/Away reduces gross published capacity to net published capacity while the
underlying published assignment remains the Staff Type key. Unknown Staff Types
or assignments remain visible with an `UNMAPPED_*` state; they are never guessed.

## Core calculations

Service Flash follows the supplied Storm business reference. Components are
summed first and ratios are calculated afterward:

```text
Service Level = answered within target
                / (offered - abandoned within target)

Routed Rate = total routed / total entered

AHT = summed handled seconds / answered calls
```

The exact service thresholds, targets, sources and queue allowlists are in
`config/metric_catalog.toml`, `config/wfm_rules.toml`, and
`config/service_profiles.toml`. Never average row-level SL or AHT percentages.

Attendance is Agent Status first with LILO fallback. Missing source evidence is
`Unknown`, not zero and not No Show. A confirmed No Show needs a completed shift
and positive disconnection evidence. Late is an exact missing boundary after
scheduled start; early leave is evaluated only after shift completion; a later
return turns the middle period into an internal gap. Meal-classified AUX is
shown as meal and does not become a correction gap. PTO/Away changes expected
work before attendance classification.

Attendance Review is read-only. Schedule plus Agent Status/LILO detects exact
gap fragments. Final Verint Activities subtract exact temporal overlap. Fully
covered fragments disappear on refresh; partial overlap leaves only the exact
residual. No decision-import or manual state is required in the Hub.

Final absence and shrinkage use Verint Activities only. Unsupported, empty-shift
or unmapped evidence is a review state, never a zero. Absence and shrinkage are
parallel rates against finalized planned net minutes and are not added together.

Capacity arithmetic:

```text
net scheduled FTE = gross published FTE - approved PTO/effective Away FTE
scheduled coverage = net scheduled FTE-hours / required FTE-hours
future gap = net scheduled FTE - required FTE
intraday present gap = Agent Status-first observed FTE - required FTE
required FTE-hours = sum(required FTE * source interval minutes / 60)
```

Forecast Volume and Absolute Required FTE are independent measures. A blank
Volume is not converted into zero requirement and a supplied requirement remains
valid even when Volume is blank.

PCS is ratio-of-sums: valid inbound Q1 score sum / valid inbound Q1 response
count. Participation is nonblank inbound Q1 responses / inbound PCSStatus=1.
Call legs are source routing records, not necessarily unique customer contacts.
The permanent tracker is refreshed from fixed CSVs; users maintain coaching in
the same shared workbook using the stable Coaching Key and Call ID.

## Product contracts

Operational Excel products:

- RTM Daily Control: validated Service Flash plus per-LOB attendance pulse and
  call/follow-up lists. Present includes late; confirmed No Show and Unknown are
  distinct.
- Attendance Review: completed schedule/observed timeline, exact residual
  correction intervals, and evidence-gated break/meal control.
- PCS Operational Tracker: permanent collaborative Power Query workbook. Do not
  regenerate it merely to update data; refresh its fixed feed and preserve keyed
  coaching actions.

In development: Staffing & Capacity Plan, Realisations, Final Absenteeism,
Bonus Management. They may be used for validation but do not get silently
promoted to operational status.

The PTO/Away submission app is a separate future Microsoft Power Platform
project. Its accepted contract is `docs/PTO_AWAY_APP_IMPLEMENTATION.md`; tenant
construction/deployment still requires the user's Microsoft environment.

## Local Manager Workbench contract

`WEBAPP.cmd` starts a standard-library Python HTTP server on `127.0.0.1` and
opens the default browser. It is reachable only from the work machine, reads
governed SQLite marts through read-only connections, serves no CDN assets and
never exposes raw extracts. Its Update button launches the same governed Hub
refresh used by the command-line menu. The approved interface is the 15-page
Manager Workbench, grouped around the actual WFM cycle:

| Workspace | Pages | Decision purpose |
|---|---|---|
| Manager Desk | Manager Desk | prioritize now, post-day reconciliation and the next seven days |
| Plan | Demand & Requirement; Capacity & Schedule; Scenario Lab | understand Staff Type demand, protect published net coverage and test a temporary FTE assumption |
| Operate | Live Service; Attendance Pulse | recover service by exact Management LOB and contact people using supported attendance evidence |
| Review | Schedule Integrity; Realisations; Absence & Shrinkage; Workforce Patterns | reconcile schedules, review plan-to-delivery, finalize Verint outcomes and surface recurrence evidence |
| Deliver | Reports & Analysis; Generated Archive | generate focused workbooks or deterministic analyses from the current database and retrieve prior outputs |
| Govern | Data Readiness; Mappings & Rules; Jobs & Logs | verify freshness, inspect effective logic and trace execution |

All pages share the compact navy navigator, a single cascading filter strip,
four evidence cards where appropriate, restrained native SVG visuals and exact
tables. Pages differ when their decision requires it: Attendance uses a
schedule-over-status timeline, Mappings uses a configuration register, and
Deliver uses explicit forms. The checked-in acceptance reference is
`docs/design-prototypes/manager-workbench-v2/all-pages-review.png`.

Page grain rules remain strict:

- Live Service and service portions of Manager Desk/Realisations operate at
  configured Management LOB plus exact call-queue membership. There is no
  invented portfolio SL; the Desk may show the lowest individual LOB as a
  prioritization signal.
- Demand, Capacity and Scenario operate at Verint Staff Type -> Planning Group
  -> Management LOB. Scenario adjustments exist only in the request and never
  rewrite source, schedule, database or configuration.
- Attendance Pulse is latest-day contact control. Schedule Integrity supports
  a selected date range and places the published schedule above chronological
  Agent Status/LILO evidence and exact residual gaps.
- Absence & Shrinkage uses only finalized Verint Activities outcomes and keeps
  incomplete evidence in a visible review state.
- Patterns reports configured recurrence evidence; it never infers intent,
  misconduct or a management decision.
- Reports and analyses read the current governed database. Generating one does
  not re-ingest sources. Only **Update data** runs ingestion and modelling.
- PCS is not imported and remains its permanent collaborative Excel workflow.

The matching Excel handoff is narrow: RTM Daily Control for today/service,
Staffing Preparation for 15-minute capacity and its persistent action ledger,
Attendance Review for exact residual evidence, Realisations for detailed cycle
results, and Final Absenteeism & Shrinkage for final Verint components. Bonus
keeps its dedicated management workbook and can be generated from Deliver. PCS
retains its separate permanent collaborative workflow outside the Workbench.

## Repository map and change discipline

- `src/wfmhub/ingestion.py`: untouched-source parsers and idempotent ingestion
- `src/wfmhub/models.py`: attendance, staffing, service, PCS and final ledgers
- `src/wfmhub/capacity_mapping.py`: capacity mapping validation and safe default merge
- `src/wfmhub/web_data.py`: governed read-only console projections
- `src/wfmhub/webapp.py`: localhost HTTP/API, refresh and export boundary
- `src/wfmhub/web/`: offline HTML/CSS/JavaScript presentation
- `src/wfmhub/report_packs.py`: current Excel product dispatch authority
- `src/wfmhub/decision_products.py`: active report builders
- `config/default_*`: shipped defaults; user files are durable runtime state
- `sql/migrations`: append-only SQLite schema upgrades
- `tests`: release gate

Never delete or overwrite user configuration, database, backups, extracts,
reports, Feed, attachments or coaching history during an upgrade. Default-map
changes merge missing identities into user copies and create a backup first.
Database migrations are append-only and existing history survives a release.
The retired PBIP implementation is historical source only. Do not regenerate,
package, publish feeds for, or repair Power BI unless the user explicitly revives
that route in a later project.

Release gate: compile Python and JavaScript, run the complete unit suite, validate
all XLSX archives, exercise every console API/view, build the portable package,
smoke-test the staged package, inspect the diff, then commit, push `main`, tag the
version and publish the GitHub release artifact.

## Known boundaries

- The console is a single-user localhost application. Collaboration remains in
  the governed Excel products; hosting for other users is deliberately absent.
- Capacity-map `UNMAPPED_*` states require a reviewed config row; do not infer.
- Schedule Integrity recurrence is supported evidence, not proof of intent.
- Pressure signals are investigation leads, not mathematical SL causality.
- PTO/Away Power Apps deployment and SharePoint permissions remain tenant work.

Supporting operator documentation begins at `docs/BEGINNER_GUIDE.md` and
`docs/LOCAL_WEB_CONSOLE.md`. Formula governance is documented in
`docs/METRIC_CATALOG_GUIDE.md`; exact service logic is in
`docs/SERVICE_KPI_REFERENCE.md`.
