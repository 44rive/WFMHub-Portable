# RTM flash reassessment — proposal, not active runtime

Status: proposed on 2026-09-23. No report, calculation, menu, or workbook
contract changes are authorized by this note. Confirm the design and reconcile
real operational examples before implementation.

## Decisions to retain

- Keep `Other Tasks` classified as BO in the effective status rulebook. Do not
  recategorize it merely to make an OEM headcount look smaller.
- PCS remains one collaborative workbook. WFM will set up its own Power Query
  connections to the six stable `Feed/PCS/*_CURRENT.csv` presentation feeds.
  Provide a beginner-friendly, exact click-by-click setup and validation guide
  before changing the PCS workbook contract. Raw Call-by-Call rows must not be
  loaded into worksheets. The current six hidden feed tables and overview
  formulas need an explicit one-time transition if Power Query replaces the
  desktop Excel sync; connection-only queries cannot feed those formulas by
  themselves. WFM refreshes and saves the shared workbook; coaching stays
  human-owned.
- RTM must distinguish a due scheduled agent who is not logged in from a
  confirmed completed-shift absence. The live action list should use plain
  language, such as `Not logged — call now`, after a configurable grace period,
  subject to PTO/Away and source-freshness checks. Keep final No Show evidence
  separate; never turn a missing extract into proven absence.
- Include active, eligible agents who are logged but have no published shift in
  an explicit unscheduled/extra coverage bucket. Show cross-LOB call handlers
  separately; queue handling does not automatically reassign home LOB.
- Audit every RTM headcount against named Client IDs, FTE LOB, assignment,
  status label, exact overlap minutes, evidence source, and checkpoint.

## Current code findings that require reconciliation

- The OEM service Flash uses three APFR Ford/Toyota/Chery call queues. Its
  attendance population uses FTE LOB `OEM FR`, not a Ford-only Staff Type.
  The default capacity mapping contains an ACM Staff Type within that broader
  OEM/Ford FR planning group. The reported `17 BO HC` cannot be explained or
  dismissed without its actual workbook and source evidence.
- `Other Tasks` is BO by rule; any overlap with an hour currently marks the
  agent in that hour's BO HC. This is a distinct-agent touched-hour measure,
  not average concurrent capacity. The day total uses the last call hour's
  workforce value; service and workforce may have different data cutoffs.
- Service uses queue call evidence, including handlers who may be outside the
  published scheduled LOB population. Logged HC currently derives from
  scheduled attendance/timeline rows only. Thus handled calls alongside zero
  logged HC need an explicit reconciliation result, not a silent zero.
- A source-data sample is required to identify the exact people behind the
  OEM `17 BO HC`; no operational extracts or RTM output are kept in Git.

## Proposed product shape — awaiting approval

Generate all flashes from one refresh and one evidence cutoff. Keep one
internal RTM Control index, plus one self-contained workbook per operational
LOB for distribution: RSA NL, RSA BE, Ford NL, OEM. Avoid four independent
refreshes that can disagree about the same day.

Each LOB workbook would have four purposeful layers:

1. **FLASH / NOW:** one compact page with SL and demand components, a clearly
   labelled service cutoff and status cutoff, current scheduled/logged/available
   capacity, unaccounted due agents, and prioritized names to call. The live
   no-show alert is operational, not the finalized absence KPI.
2. **HOURLY:** 00:00–23:00 service components and target, queue contribution,
   with hourly workforce expressed as average occupied FTE/agent-minutes or a
   clearly timestamped HC snapshot. Do not label any-overlap HC as concurrent
   staffing. Show no-sample and stale-evidence states explicitly.
3. **15 MIN / CAPACITY:** four intervals per hour from the existing 15-minute
   service and forecast marts. Service queue metrics, Staff Type forecast and
   required FTE, scheduled/net/observed capacity, and source freshness stay in
   separately labelled sections. Do not compare forecast Staff Type volume to
   queue-entered volume as if they were the same grain. RSA BE service remains
   combined; FR/VL capacity remains separate.
4. **AGENTS / AUX / DAY CLOSE:** named scheduled and unscheduled presence,
   source status labels and exact minutes (including `Other Tasks` as BO),
   callout evidence, and a provisional Agent Status-based realization of
   scheduled hours. After the day completes, link/reconcile to the governed
   Realisations and finalized Verint absence reports instead of duplicating
   their business definitions in a Flash.

The internal index should flag `handled calls > 0 AND logged HC = 0`, BO HC
above plausible roster capacity, stale Agent Status, missing schedules,
unmapped statuses, overlapping shifts, and missing Staff Type mappings. These
are review alerts, not invented explanations of service-level loss.

## Acceptance gate

Before release, reconcile at least one real day per LOB, including OEM's
`17 BO HC`, against named Agent Status intervals and FTE/schedule rows. Test
night-shift carry-over, pre-shift/unscheduled logins, cross-LOB handlers,
partial PTO/Away, no login, late login, logout-and-return, missing/stale
extracts, and 15-minute-to-hour aggregation. Verify SL numerator/denominator
against Storm separately from workforce capacity. No live report is promoted
until these checks pass.
