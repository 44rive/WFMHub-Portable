# TOYOTA BO Live Control: design record

Status: the standalone companion is implemented in `miniapps/bo_live/`.
Its dashboard, capture validation, publisher, independent SQLite store, demo,
and Windows packager are available. Actual ABS extraction/selectors remain a
Windows PAD setup task. See [the app guide](../../miniapps/bo_live/README.md).

## Goal and confirmed context

Give a TOYOTA back-office team leader a current view of work needing attention:
which listed tasks are unassigned, high priority, or past their action time.

The source is ABS@ACM B2B2C in Edge, displaying individual tasks in a work-list
grid. It has no CSV or Excel export. Power Automate Desktop (PAD) is the
proposed collector on a Windows machine. WFMHub already supplies local SQLite,
governed calculations, source health, and a localhost manager workbench.

Scope is TOYOTA only. The supplied screens show the home page, Actions menu,
Task Groups dialog, TOYOTA group selection, and Work list file / Task list.
The attachment order does not establish the navigation order. Confirm the
transition from group selection to the work list in the live session.

The user confirmed that `Scheduled for` is the time a task should be actioned
and that ABS exposes a case number only, rather than a unique task ID. Case
numbers therefore identify cases, not individual tasks.

Start extraction validation with one TOYOTA task group, then cover the exact
TOYOTA group list. Target a five-minute full scan if measured collection time
allows it. This is near-real-time polling.

The [PAD build specification](TOYOTA_PAD_BUILD_SPEC.md) maps the observed
navigation and fields into the proposed collector. Screenshots establish UI
labels and layout, not working selectors or extraction capability.

## First version

One TOYOTA BO page in a standalone companion, with a summary by task group and
a prioritized task list linked by case number. Every view shows source scope,
last successful capture, and health. Retain repeated case numbers.

| Element | Proposed behavior | Required source evidence |
|---|---|---|
| Listed backlog | Count validated task rows for each TOYOTA group | One-task-per-row contract, complete list scope |
| Unassigned work | Tasks with a confirmed empty assignment | Full `Assigned to` value and known blank semantics |
| Priority | Counts by exact source priority | Full `Task priority` value |
| Action timing | Past scheduled time, upcoming, and future work | Full `Scheduled for` timestamp and source time zone |
| Action list | Tasks ordered by passed action time, priority, then scheduled time | Agreed ordering and complete row fields |
| Language / work mix | Breakdown by language and task type | Full source values, retained without inferred mappings |
| Capture health | Fresh, stale, failed, or incomplete collection | Run metadata, scope, counts, timestamps |

Full scheduled timestamps must be extracted rather than the truncated display
text. An unscheduled or unreadable value is a visible separate category. The
upcoming window and exact equality boundary belong in governed configuration.
Measure time past schedule at the recorded group capture time, with that time
visible, so a stale snapshot does not look current. This measures action timing;
it does not establish SLA breach, case age, or completion performance.

Case links are available only if the application exposes a stable link. The
first version does not need task descriptions or customer details. Contract
name can be retained later if a confirmed operational requirement needs it;
TOYOTA membership comes from exact task-group scope.

Example action wording: "Review unassigned high-priority tasks whose scheduled
action time has passed." Rules and ordering belong in governed configuration.

## Proposed data flow

```text
ABS: exact TOYOTA task groups
        |
        v
PAD: select group, refresh, read grid, validate full coverage
        |
        v
Immutable local capture + run manifest
        |
        v
WFMHub: validate, ingest, calculate BO measures in local SQLite
        |
        v
Standalone BO Live workbench: summary, action list, source health
```

PAD collects evidence. The BO Live companion is its own database writer and
calculation owner. Its inbox watcher imports validated captures; the main
WFMHub database and runtime remain independent.

### PAD collection flow

1. Open or attach to the authorized browser session and navigate to the agreed
   queue view. Authentication follows the source tool's normal sign-in flow.
2. Confirm expected page, queue, filters, sort order, and required headers.
3. Test browser element/table extraction on the application-style grid. If
   necessary, assess desktop UI automation/accessibility in the actual session.
   Discover stable selectors and handle both grid scroll directions, virtual
   rows, pagination if present, and lazy loading.
4. Collect all rows in the agreed scope, retaining IDs as text. Record capture
   start/end, source time zone, page totals when available, filters, row count,
   extraction version, and success/failure details.
5. Validate required fields, row-position coverage, traversal completion, and
   displayed totals when their meaning is confirmed. Repeated case numbers
   remain separate task rows. Count drift triggers a retry or an
   incomplete flag. A multi-page capture has a capture window; it cannot be
   assumed to be a transactionally consistent instant.
6. Publish a uniquely named capture only after validation, using a temporary
   file followed by an atomic rename. Failed runs produce health evidence and
   do not replace the last good complete capture. Keep captures in the
   configured local data area, outside Git and separate from generated feeds.
7. Let the running companion ingest published captures from its inbox. Prevent
   overlapping PAD runs and use bounded retry/backoff on page/session failures.

If browser extraction cannot read the actual UI, assess the accessibility
tree or another supported source interface. Image/OCR extraction requires a
separate reliability decision before case references or scheduled times become
operational evidence. Do not assume that an API is available.

## Case evidence and calculation rules

- A case number is not a task key. Preserve it as text. One case may have
  several tasks within or across task groups; do not deduplicate them by case
  number or by a guessed combination of visible fields.
- A validated group snapshot may use a row occurrence/index within that capture
  to retain every task. This index is not a persistent identity across scans.
  Repeated scroll reads require verified absolute row positions or a complete
  table extraction; identical visible values alone cannot identify overlap.
- Candidate fields: case number, exact task group, task type, task priority,
  language, scheduled action time, and assignment. Case number requires the
  user-confirmed source location to be identified during setup. Exact labels
  and remaining columns still need a live inspection.
- Missing fields remain unavailable. Missing or failed collection is not zero
  backlog. Retain the last good view with a visible stale/incomplete state.
- Proposed stale threshold: two missed scheduled collection cycles, configurable
  once measured. A failed or incomplete run is visible immediately.
- Confirm which tasks the work-list scope includes, assignment semantics,
  priority ordering, source time zone, and the upcoming-action window before
  calculating indicators. Creation and completion times are not established
  by the screenshots. True ageing and SLA measures remain deferred.
- Count each complete snapshot independently; do not add snapshots together.
  Re-importing a capture must be idempotent and preserve its lineage.
- A row missing from the next snapshot may have moved group, been filtered,
  disappeared during traversal, or completed. Task lifecycle, completions,
  arrivals, and throughput require reliable task identity plus explicit event
  evidence. Snapshot count changes describe net backlog movement only.
- Preserve BO case evidence separately from Agent Status attendance evidence.
  An owner field does not establish presence or productive time. Any later
  employee join requires an exact governed ID and dated FTE eligibility;
  unmapped owners remain visible.

## Pilot and next increment

Pilot scope: TOYOTA, one Windows runner, one browser, five-minute full-scan
target, one local BO workbench page. Validate one task group first, then the
full exact TOYOTA group list. First validate extraction and capture quality,
then implement ingestion, configured metrics, and the page using synthetic
fixtures. Confirm Windows session requirements and the PAD licensing needed
for the chosen attended or unattended execution mode during runner setup.

Pilot acceptance:

- Reconcile a known complete list scope, including vertical/horizontal scroll,
  row positions, case references, and counters where their meaning is confirmed.
  Validate metrics against the same captured rows, including repeated cases.
- Check priority, assignment, and scheduled-action indicators against agreed
  source fields and rules; leave unsupported measures visibly unavailable.
- Demonstrate session expiry, missing columns, partial traversal, valid empty
  results, repeated cases, retries, and runner interruption without false zeros.
- Measure capture duration and freshness across a representative work period.
- Validate re-import idempotence and keep failed captures from replacing good
  evidence. Keep all real operational data outside the public repository.

After the pilot, consider snapshot count trends, explicit completion throughput
once task identity and completion evidence are available, and BO
capacity context from existing WFMHub attendance/schedule sources. Show each
source's own observation time; hourly touched-state BO headcount is not a live
simultaneous headcount. Staffing recommendations also need skill mappings and
agreed handling-time assumptions.

## Remaining discovery

- Full exact names of the two clipped TOYOTA group entries and any additional
  TOYOTA entries outside the visible dialog area.
- Successful PAD extraction of full cell values, case number location, source
  time zone, assignment semantics, and verified grid counter meaning.
- Confirmed work-list filters and the transition after selecting a task group.
- Whether the application exposes a task identifier beyond the currently known
  case number; the first version must work without assuming one exists.

Use synthetic fixtures for repository examples. Real screenshots, case rows,
employee names, private source URLs, and captured outputs stay outside Git.

Implementation proceeds as a standalone mini app. Its source contract,
local governed settings, tests, packaging, and product documentation live
outside the roadmap folder. Integration with the main WFMHub runtime remains
a separate future decision.
