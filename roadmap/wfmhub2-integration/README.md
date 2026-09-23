# WFMHub-2 assessment and selective integration proposal

Status: assessment/proposal recorded 2026-09-23; **no code or database merge is
authorized by this note**. Reassess at implementation time against current
GitHub heads and the managed-workstation result.

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

## Judgment

Do **not** run a Git branch merge or replace the old app with Preview `.7`.
That would trade useful daily WFM products for a promising but incomplete
foundation. Also do not discard WFMHub-2: its portable security boundary,
atomic generation model, no-op source reuse, diagnostics, and UI composition
solve real weaknesses in the older product.

The recommendation is **one production product and one development track**:
keep WFMHub-Portable as the operational owner for now; selectively port proven
WFMHub-2 components into it behind acceptance tests. `WFMHub-2` remains a
reference/prototype until its distinct capabilities are integrated or a future
full-replacement decision is made. This is a recommendation, not an approved
product migration decision.

## Proposed merge sequence (not yet started)

1. **Freeze the contracts.** Put the current effective source roles, Client ID,
   Active/dated-Leaver, exact Flash queue allowlists, metric methods, Staff Type
   requirement, RSA BE service-versus-capacity split, PTO/Away, Agent Status,
   final Verint Activities, PCS and coaching ownership into versioned parity
   fixtures. Defaults alone are insufficient. Use synthetic fixtures in Git;
   compare confidential operational results locally without publishing them.
2. **Keep one authoritative SQLite.** Start with the current operational DB and
   expose existing governed marts through a narrow read API. Never create a
   second copy of facts solely to make the new UI work. Improve the old local
   HTTP boundary with WFMHub-2's per-launch token/Host validation before
   exposing more actions, while preserving local-only/offline operation.
3. **Port UI as presentation, not arithmetic.** Bring the useful React shell and
   dense RTA visual patterns into the current product, then connect one page at
   a time to existing calculations. First page: one LOB's live service plus
   named attendance and 15-minute staffing ladder. Blank/unknown/stale data
   stays explicit. No fake cards from mockups. Keep `.cmd`, Excel, and CSV as
   working alternatives.
4. **Port refresh mechanics separately.** Adopt generation staging and
   source-version reuse only after measuring old-versus-new counts, error
   behavior, refresh duration, and database size on the same read-only source
   set. The first win should be fast unchanged refresh plus atomic rollback;
   affected-date rebuild comes later. Schema changes need normal additive
   migrations, verified backup, and upgrade tests. Do not transplant the new
   `control.sqlite` wholesale.
5. **Close business parity by WFM cycle.** Reconcile Storm SL numerator and
   denominator by exact profile; night shifts, unscheduled logged agents, PTO/
   Away, and unknown/no-show; Verint Staff Type requirement and RSA BE FR/VL
   capacity; final Activities absence; generated reports and PCS coaching.
   Advance a page only after it helps a real RTA/manager decision and agrees
   with the approved business reference. Keep `Other Tasks` as BO.
6. **Rationalize only after acceptance.** Remove duplicate adapters/servers and
   unused native-stack scaffolding only when the chosen implementation covers
   the validated workflow. Preserve both Git histories and backup/restore paths.
   Feature-gate DuckDB-Wasm, Pyodide and HiGHS until they beat a simpler
   deterministic workflow on measured business value and workstation cost.

### Acceptance gate before any replacement

One extracted ZIP must work offline on the managed PC; normal and failed
refresh must preserve the last valid cut; same-source outputs must reconcile
for RSA NL, RSA BE combined service plus FR/VL capacity, Ford NL and OEM;
RTM/attendance/staffing/absence/PCS/report actions must match or explicitly
supersede the old product; the shared PCS workbook and coaching must survive
updates; source extracts and local SQLite must remain untouched by upgrades.
Record measured elapsed time and memory for bootstrap, unchanged, added-day
and corrected-day refresh—not just a green synthetic CI run.

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
