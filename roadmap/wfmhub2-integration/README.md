# WFMHub technology adoption and RTM vertical slice

Status: revised adoption proposal recorded 2026-09-23. WFMHub-2 grew from the
portable Hub; it is a technology and WFM-architecture reference for upgrading
the **same operational product**, not a competing business product. **This
note does not implement or authorize a runtime/database migration.** Recheck
both repository heads and the managed-workstation result before implementation.

Compared:

- [WFMHub-Portable](https://github.com/44rive/WFMHub-Portable) at `bcb6196`
  (`v1.1.2`): current operational baseline and authoritative business contract
  in this repo's `PROJECT.md` and effective user configuration.
- [WFMHub-2](https://github.com/44rive/WFMHub-2) at `1e21f1e`
  (Phase 1 Source Preview `.7`, 2026-09-22): experimental next-generation
  host, refresh store, and React shell. Its `PROJECT_LEDGER.md` is its own
  authoritative status record.

## What the review actually found

| Area | Current portable product | WFMHub-2 preview `.7` |
| --- | --- | --- |
| Daily WFM output | Service/RTM, attendance review, staffing, realization, final absence, governed analysis, bonus, and permanent PCS workflow exist; some RTM measures still need the explicit reconciliation in `roadmap/rtm-flash-reassessment/`. | Source health and refresh are available. Ingestion covers FTE, published StartEndTimes, Agent Status, LILO, Call-by-Call; it stores additive 15-minute service components and attendance agent-day/gap evidence. Its own ledger says service headline, staffing requirement, agent action screen and Excel product parity are not delivered. |
| Workstation runtime | Proven embedded CPython 3.13, SQLite, pure-Python Excel, CLI and local web workbench; v1.1.2 ZIP/CI pass. | Exact Phase 0.4 host/Edge/WASM compatibility passed on managed workstation. Preview `.7` first refresh reportedly succeeded and unchanged refresh reused its active generation, but no production parity or changed-source duration/counts are recorded yet. |
| Refresh/state | One mature but relatively monolithic operational SQLite/report pipeline with extract fingerprints and migrations. | Better generation-keyed staging, source-version reuse, progress/diagnostics and whole-cut rollback. A changed source still rebuilds derived attendance/service for the whole cut; the affected-date performance goal is not yet achieved. |
| UI | Broad but plain functional local workbench and `.cmd` workflows. | More coherent React/TypeScript shell, but `web/src/app/navigation.ts` enables only Product foundation and Compatibility Doctor; the polished design PNGs are prototypes, not delivered RTA pages. |
| Rules/mappings | Effective user-editable metric catalog, service profiles, capacity mapping, status and source rules; verified per-LOB Flash queues and Staff Type separation must be preserved. | Queue mapping is byte-identical to the old **default** mapping, but the active host currently uses a small immutable packaged policy set and has not ported the full effective metric/service/capacity contracts. A user's effective mapping may differ from either default. |
| Excel/collaboration | PCS remains one SharePoint workbook with six manually copied CSVs and desktop Excel Refresh All; coaching must survive. Bonus is permanent; other workbooks are generated products. | No equivalent PCS/bonus/realisations/RTM workbook pipeline in the portable compatibility host. The Excel dependency currently serves doctor tests. |

The new repo also contains two Python directions: `src/wfmhub2` with a native
analytics/FastAPI graph and `src/wfmhub2_compat` with the actual policy-compatible
portable host. The hybrid ZIP packages the latter. This separation was useful
for qualification but is a maintenance/ownership cost if both remain active.
Its current parity checklist still names old portable `v0.36.0`; the operational
baseline reviewed here is `v1.1.2`, so that checklist must be rebased before
any replacement decision. The Preview `.7` ZIP is about 73 MB versus roughly
12 MB for v1.1.2. That is not a correctness failure, but optional browser
analytics must earn their packaging and support cost with a real workflow.
The two products have **different SQLite paths and schemas** (`_system/database/
wfm.sqlite3` versus `data/control.sqlite`); never point them at one another's
database or merge tables by filename/copy. Both can read the same source-folder
evidence read-only during a controlled comparison.

## Revised judgment

The current portable Hub is the functional product and should remain the
canonical repository, release, CLI, database, and WFM calculation authority.
WFMHub-2 is a derivative whose **portable-compatible** stack and boundaries can
modernize that product. A Git merge or runtime swap would be the wrong unit of
change: its `control.sqlite` schema differs, its native development graph is
blocked on the work PC, and the React preview has not implemented the daily
products. Conversely, leaving the old workbench unchanged wastes a tested
browser stack and a better way to structure WFM services.

This is an **in-place modernization**, not a parallel Hub. Keep the existing
operational workbooks and PCS collaboration path running throughout. Use the
RTM redesign as the first complete vertical slice so that the new architecture
is tested on a real RTA job, not on a decorative dashboard. WFMHub-2 remains a
reference until its useful pieces have been integrated and verified.

## Target technology decision

| Layer | Adopt in the current Hub | Deliberately do not adopt now |
| --- | --- | --- |
| Portable host | Retain the already-proven official embedded CPython, stdlib SQLite, OpenPyXL/XlsxWriter, `.cmd` launchers, and local-only operation. Adapt Hub2's per-launch token, Host checks, static-asset security, and diagnostics to the existing server. | A second Python host, native FastAPI/Pydantic/DuckDB/Polars graph, installer, or cloud dependency. |
| Browser | Build React/TypeScript into static local assets with Vite at release time; use Hub2's shell/design tokens, dense table/virtualization and charts where they solve a WFM task. Node is a **build-time** tool, never required on the work PC. | Copying prototype pages as if they were functional; adding every front-end package before a page needs it. |
| Durable data | Keep the existing `_system/database/wfm.sqlite3` and additive migrations. One Python writer; one effective config and metric catalog; same governed marts feed CLI, Excel, CSV, and browser. | Copying `data/control.sqlite`, maintaining two authorities, or recalculating KPI formulas in JavaScript/Excel. |
| Refresh | Adapt generation/no-op/source-reuse concepts behind the existing refresh contract, measuring actual time and rollback first. | Whole-schema replacement or claims of affected-date incremental refresh before it is implemented and benchmarked. |
| Browser compute | None is needed for the first RTM release. | DuckDB-Wasm, Pyodide, HiGHS-Wasm, OPFS state, and forecasting/optimization until a measured WFM use case justifies cost. |

The meaningful stack upgrade is therefore **secure local host + compiled
browser UI + modular domain services + safer refresh**. Python, SQLite and
Excel are retained because they already work on the managed PC. Hub2's Bronze /
Silver / Gold terminology can describe the existing `raw_*` / `core_*` /
`mart_*` responsibilities without building three databases or renaming tables.

## WFM architecture to adopt

```text
Read-only source adapters (FTE, schedules, Agent Status, LILO, CBC, forecast,
                           finalized Verint Activities)
    -> canonical facts at their natural grain, with source/quality evidence
    -> governed domain services (service, attendance, capacity, absence)
    -> one RTM read model / other WFM products
    -> local API + CLI + workbooks + CSV feeds
```

Extract parsers must not decide KPIs. Domain services must not import browser,
HTTP, or workbook code. The UI may format, filter and visualize server-supplied
components, but may not invent its own SL, no-show, availability, or staffing
arithmetic. Keep the effective mappings and catalogs as the sole rule authority.
Do this **slice by slice**: extract a tested service from `models.py`, retain a
compatibility facade for existing callers, then move the next service. Do not
copy Hub2's empty domain folders or create two live calculation engines.

The UI follows the actual WFM cycle: **Operate** (RTM today and attendance
actions), **Plan** (forecast, requirement, schedule, staffing), **Review**
(realisations, schedule integrity, final absence), **Govern** (data health and
rule versions), and **Deliver** (workbooks/exports). Only working pages appear.
Operational Excel remains a printable/shareable snapshot; the local app is the
interactive control surface. PCS stays the permanent shared workbook and is
outside the first migration slice.

## Adoption sequence: technology and RTM together (not yet started)

1. **Freeze the contracts and baseline.** Put the current effective source roles, Client ID,
   Active/dated-Leaver, exact Flash queue allowlists, metric methods, Staff Type
   requirement, RSA BE service-versus-capacity split, PTO/Away, Agent Status,
   final Verint Activities, PCS and coaching ownership into versioned parity
   fixtures. Defaults alone are insufficient. Capture the current RTM output,
   named-agent evidence and refresh timings on representative real days locally;
   use synthetic fixtures in Git. Do not publish confidential extracts.
2. **Build one host and one RTM service contract.** Keep the current database and
   CLI. Add a narrow, token-protected API to the existing local server. Define
   typed RTM response shapes and source cutoff/freshness states. Extract RTM
   arithmetic behind a Python domain-service boundary while its old callers
   continue to work; the browser and workbook must consume the same governed
   read model. Do not create a second copy of facts just for React.
3. **Ship one usable RTM vertical slice.** Compile a React/TypeScript shell into
   the portable ZIP and connect an **Operate / RTM** page to the real service:
   combined service by approved LOB queues; 15-minute and hourly views; Staff
   Type requirement and net schedule; observed capacity; named agents to call;
   source confidence and issues/drivers. Keep RSA BE combined for SL but FR/VL
   split for staffing. Show 24 hours, exact cutoff, and explicit unknowns. The
   corresponding per-LOB Flash workbook is generated from the same read model,
   not from a parallel formula implementation. `roadmap/rtm-flash-reassessment/`
   defines the detailed RTM behavior and acceptance cases.
4. **Validate without interrupting operations.** Run the new RTM view/workbook
   beside the current operational Flash on the same source cut. Reconcile Storm
   SL components, the OEM BO anomaly, cross-LOB handlers, night shifts,
   unscheduled logins, PTO/Away, no-show versus not-yet-logged, and requirement
   gaps by Client ID and interval. Keep old launchers/reports available and do
   not promote RTM until cases pass on the work PC. Other operational products
   remain on their current code path during this slice.
5. **Improve refresh behind that product.** Adopt generation staging and
   source-version reuse only after measuring old-versus-new counts, error
   behavior, refresh duration, and database size on the same read-only source
   set. The first win should be fast unchanged refresh plus atomic rollback;
   affected-date rebuild comes later. Schema changes need normal additive
   migrations, verified backup, and upgrade tests. Do not transplant the new
   `control.sqlite` wholesale.
6. **Move the rest by WFM cycle, then rationalize.** Port attendance review,
   staffing preparation, realisations, final absence, reports and governance
   only as each view is useful and parity-tested. Keep PCS's SharePoint CSV /
   Power Query / coaching workflow and the permanent Bonus workbook intact.
   Retire superseded old web screens only after their replacement is accepted;
   preserve both Git histories and the upgrade/restore path.

### Acceptance gate before the new RTM is promoted

One extracted ZIP must work offline on the managed PC with **the existing
database and CLI**. The new React page and Flash workbook must agree with one
another and reconcile same-source service/capacity components for RSA NL, RSA
BE combined service plus FR/VL capacity, Ford NL, and OEM. Named attendance,
night shifts, unscheduled logins, PTO/Away, missing/stale evidence and the OEM
BO case must pass the detailed RTM checklist. A failed RTM/API request must
not affect stored facts or existing reports. Keep the old Flash available until
the new one passes; PCS coaching, Bonus and all other operational products must
remain usable, without a forced workbook or database replacement.

Before later refresh-engine adoption, test normal/failed activation and record
elapsed time and memory for bootstrap, unchanged, added-day and corrected-day
refresh. The last valid cut must survive failure. Do not retire any other
workbook or page until its own business-parity gate has passed.

## Current handoff boundaries

- PCS v1.1.2: Hub updates six local CSVs; WFM copies all six to fixed
  SharePoint location; opens the **same** shared workbook in desktop Excel,
  chooses Data > Refresh All, waits and saves. The one-time manual setup is
  `docs/PCS_POWER_QUERY_SETUP.md`. Do not run legacy Hub Sync on a query workbook.
- RTM: the specific Flash/Issues & Drivers/no-show-capacity proposal is in
  `roadmap/rtm-flash-reassessment/README.md`; it is not implemented or promoted.
- WFMHub-2 Preview `.7` is not a substitute for these workflows today. Its
  [ledger](https://github.com/44rive/WFMHub-2/blob/main/PROJECT_LEDGER.md),
  [architecture](https://github.com/44rive/WFMHub-2/blob/main/ARCHITECTURE.md),
  [navigation flags](https://github.com/44rive/WFMHub-2/blob/main/web/src/app/navigation.ts),
  and [release](https://github.com/44rive/WFMHub-2/releases/tag/v0.2.0-phase1-source-preview.7)
  are the evidence for this assessment.
