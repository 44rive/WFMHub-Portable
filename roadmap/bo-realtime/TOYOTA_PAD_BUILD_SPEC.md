# TOYOTA ABS collector: PAD build specification

Status: flow design based on supplied screenshots and user answers. The
standalone app and CSV publisher are implemented; use the executable commands
in [PAD setup](../../miniapps/bo_live/PAD_SETUP.md). Actual ABS selectors and
grid extraction have not been recorded or tested in the Windows PAD session.

## Confirmed scope and meaning

- Application: ABS@ACM B2B2C in Edge; no CSV/Excel export.
- TOYOTA only. Scan exact configured task groups, starting with one group.
- Work-list rows represent tasks. `Scheduled for` is the time a task should
  be actioned. A true SLA deadline or creation timestamp is not established.
- The user reports a case number only. A case may have several tasks; case
  number cannot be used as a unique task key.

This document contains navigation labels and source-field metadata only. Keep
actual URLs, screenshots, case records, and employee details out of Git.

## Navigation observed across the screens

```text
ABS home
  -> Actions
  -> Task Groups ...
  -> select exact TOYOTA task group
  -> OK
  -> reach/verify Work list file -> Task list -> List
  -> Refresh and read the agreed scope
```

The screenshots show these states but do not prove whether OK opens the list
automatically. During setup, record whether Work List must be opened afterward.
Use UI state checks rather than assuming the dialog sequence or fixed delays.

Choose groups from the main list of available groups. The Most recently used
list is navigation history and may contain other products. Confirm the full
exact selected group and resulting list scope on every iteration.

### TOYOTA group inventory

Four full labels are readable in the provided group dialog:

- `(TOYOTA) Agendas à Traiter`
- `(TOYOTA) REFACTURATION`
- `(TOYOTA) SUIVI VR`
- `(TOYOTA) TAKATA`

Two further TOYOTA labels begin with `(TOYOTA) EXTRANET ENTRAN...` and
`(TOYOTA) WEBAPP NV DOSSI...`. Those fragments are discovery notes, not selector
values. Read their full text and inspect the whole list for additional TOYOTA
groups during setup. Store full exact names in local collector configuration;
prefix discovery must not silently become production group membership.

## Field contract

| Source field | Capture rule | First-version use |
|---|---|---|
| Task Group | Full exact value; verify against selected group | Scope and grouping |
| Task type | Full value | Work mix |
| Task priority | Exact label; observed labels include Very high, High, Standard | Priority counts and configured ordering |
| Language | Full value | Language breakdown |
| Scheduled for | Full timestamp, raw value, parsing outcome, source time zone | Action-time buckets |
| Assigned to | Full value plus whether extraction succeeded | Assignment breakdown once blank semantics are confirmed |
| Case number | Locate during setup; retain text | Case reference, with repeated values allowed |

`Task description` and `Contract Name` are visible but not needed for the
initial indicators. Several other headers are clipped. Do not invent their
names or meaning. A blank or clipped UI cell is not proof of an empty value.
The header strip above the grid may describe the selected row or grouping;
confirm its behavior before interpreting it as a list filter.

## Build the flow in this order

### 1. Prove one-row extraction in the actual Windows session

Attach PAD to the authorized Edge session. Use its web recorder / element
picker to test whether a header and a cell expose structured browser values.
The application-style windows do not establish whether the grid uses HTML,
canvas, or another rendering mechanism.

If structured browser extraction works, test full table extraction. Otherwise
test desktop UI automation/accessibility and full-value retrieval. Selectors
must come from the real session. Record a one-row experiment before building
the navigation loop, and compare extracted values with the source display.

Get full scheduled time and case number without changing a task. Horizontal
scroll or a read-only Details view may help; record where case number actually
appears. Verify whether structured values remain available for clipped cells.
If no reliable text interface exists, document the result before choosing a
different collection method. OCR is not yet a validated task-data source.

### 2. Build guarded navigation

Suggested flow variables: `BrowserInstance`, `ExactToyotaGroupNames`,
`SelectedGroup`, `RunId`, `BoCaptureRoot`, `MaxScanDurationSeconds`, and
`MaxTraversalAttempts`. Store the source URL in local settings.

Use PAD actions to attach to Edge, wait for the ABS home or known work-list
state, open Actions / Task Groups, select the exact group, and confirm OK.
Wait for the expected work-list state. Reopen Work List only if live setup
shows it is required. Refresh once, wait for loading to finish, and verify
scope and filters before reading.

The collector navigates, refreshes, and reads. Assignment and completion remain
operator actions in ABS. Record selectors for the read path; action controls
such as Assign, Allocate, and Complete selected are outside that path.

### 3. Prove complete grid traversal

- Inspect the horizontal scrollbar to locate all required fields. Capture full
  values for each row without combining columns from different row positions.
- Inspect vertical scroll, virtual row loading, any pagination, and the list
  counter. Confirm whether the counter means selected row / total rows and
  whether its total is filtered or capped; the screenshot alone cannot prove it.
- Prefer a complete structured table read. If traversal requires overlapping
  viewports, use verified absolute row positions to distinguish repeated reads
  from genuine tasks having identical field values.
- Never collapse rows using case number or a hash of visible fields. Two tasks
  on one case, including two visually identical tasks, must both survive.
- If complete row coverage cannot be demonstrated, mark the detailed capture
  incomplete. A source-total-only capture may be a separately labeled fallback
  after the counter semantics are verified; it cannot supply row breakdowns.
- Recheck totals and scope at the end. Detectable list changes trigger bounded
  retry or an incomplete result. Record capture start/end because an unchanged
  total does not guarantee an atomic or unchanged multi-page list.

Use bounded traversal and retry limits. Expired sessions, unreadable cells,
wrong groups, and stalled scrolls end in explicit health states.

### 4. Publish per-group captures and full-scan status

Each group capture records:

- Schema/extraction version, run ID, exact group, filters, and sort state.
- Capture start/end, observation time, source time zone, and completion status.
- Source total if readable and verified, extracted row count, coverage evidence,
  and any quality failures. Missing metadata is distinct from zero.
- Rows with case reference and a capture-local occurrence/row index. The index
  is an extraction artifact and must not become a persistent task identifier.

Write UTF-8 structured rows plus a manifest into a temporary local capture
directory. Close files before atomically publishing the group directory for
WFMHub ingestion. Keep real captures in configured data storage outside Git.
Failed captures retain diagnostic metadata and cannot replace the last good
group evidence.

After all configured TOYOTA groups have been attempted, publish a scan manifest
listing each group's result and observation time. A failed group marks the scan
partial. Last-good group data can remain visible with its own timestamp, but a
mixed-age total must not be presented as a fresh complete TOYOTA snapshot.
Cross-group task moves during scanning also prevent an exact simultaneous
total; show the scan window and the sum of captured group counts as such.

WFMHub remains the sole SQLite writer. A future bounded import job validates
the captures and calculates the measures. PAD does not write SQLite directly.
Prevent overlapping scans; measure all-group duration before setting the
five-minute cadence and use bounded backoff after failures.

## Proposed workbench

Top cards: listed tasks, very-high-priority tasks, unassigned tasks, and tasks
past scheduled action time. Each card requires a complete supporting capture.

Group table: task group, listed total, priority mix, assignment mix, past action
time, upcoming window, future/unscheduled tasks, capture time, and health.

Action list: case number, task group, task type, priority, language, scheduled
action time, and assignment. Preserve each task-row occurrence. Display
source-local scheduled time with its time zone; use the recorded observation
time to calculate timing categories for the captured evidence.

The first version measures task rows and net backlog movement between complete
comparable scans. It cannot identify which task arrived or completed, infer a
completion from a vanished row, or derive case age from scheduled action time.

## Acceptance checks before runtime implementation

1. Verify all exact TOYOTA groups and the work-list transition after OK.
2. Extract full fields from an initial, middle, and final row, including fields
   reached with horizontal scroll; reconcile coverage with the verified total.
3. Preserve multiple tasks on one case and two visually identical tasks. Repeated
   reads of one absolute row do not inflate a snapshot count.
4. Verify task counts, priority, assignment, and scheduled-action categories
   against the same capture; leave ambiguous fields visibly unresolved.
5. Exercise a genuine empty group, partial traversal, count drift, wrong scope,
   session expiry, unreadable timestamps, retries, and interrupted publication.
6. Exercise one failed group in an otherwise successful TOYOTA scan; verify
   partial coverage, separate last-good timestamps, and no false current total.
7. Measure a representative full scan and pick a non-overlapping collection
   cadence. Confirm the execution mode and Windows session requirements.

All repository fixtures are synthetic. Keep PAD validation evidence containing
real case or employee information in local operational storage.
