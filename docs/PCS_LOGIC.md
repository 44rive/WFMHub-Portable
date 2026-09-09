# PCS calculation logic

The official calculation is reproduced from `TOLEARN/PCS Report.xlsx`,
especially `OverView!O15:R15` and the embedded `RDATA` query.

## What a call leg is

A call may pass through a queue, transfer, or agent more than once. Each such
record is a call leg. WFMHub keeps the source leg grain and deduplicates
overlapping extracts by a stable call key. `transferred_legs` is descriptive;
it does not change the PCS formula.

## Exact formula

Only inbound legs with an in-scope Agent ID enter the official PCS counters.

- A valid score is raw Q1 parsed to exactly one configured value. The default
  set is `1, 2, 3, 4, 5`; `4.5`, `0`, `555`, `*`, `No_Response`, and similar
  values are invalid.
- `PCS average = sum(valid Q1) / count(valid Q1)`.
- Negative/count column = valid Q1 `<= 3`.
- Positive/count column = valid Q1 `> 3`.
- Participation numerator = inbound legs where raw Q1 is nonblank, including
  invalid markers.
- Participation denominator = inbound legs where `PCSStatus = 1`.
- `PCS participation % = raw-Q1 nonblank / PCSStatus=1`.

Q2 does not affect the official score. `PostCallSurveyMode=2` is retained as a
diagnostic count but is not the participation denominator.

## Coaching and Actions Rate

The original `OverView!S15` formula counted completed briefing rows from an
external personal workbook and divided them by the agent/date count of valid
Q1 scores `<=3`. WFMHub keeps that business meaning but removes the broken
external link:

- every valid inbound Q1 `<=3` creates one coaching opportunity;
- each opportunity is identified by the stable deduplicated call-leg key;
- `Actions Rate = unique completed Coaching Keys / all coaching opportunities`;
- `PCS Live Tracker.xlsx > COACHING` exposes the exact filtered calls;
- a reviewer copies Coaching Key and Call ID into the blue action table on that
  same sheet, then fills Status, Coach, Coaching Date, Due Date and Comment;
- normal prepares never replace the tracker; an explicit contract repair
  archives the prior file and carries keyed actions into its replacement;
- later DATA pastes update the queue while saved keyed actions remain in place;
- coaching decisions are never imported into SQLite.

`Not required` remains in the denominator and is not counted as completed.
Duplicate Coaching Keys are highlighted and cannot increase the completed count.
Low sample is an interpretation warning, not a coaching opportunity by itself.

At team and month level, counters are summed first and the ratios are then
recalculated. Agent averages and percentages are never averaged together.

Python/SQLite prepare one clean, additive call-leg table. Each prepare writes a
timestamped `PCS Paste Data` workbook. The user replaces the body of
`PCS Live Tracker.xlsx > DATA!tblData`; Excel then calculates cards, LOB and
agent results, and the coaching queue from sums of the additive counters. The
tracker has no Power Query, Data Model, macro, external connection or Hub-driven
refresh. Fixed classic formulas use the Hub-prepared list flags to populate the
selectors and filtered views. There are no spill formulas or dynamic-array
metadata. The permanent coaching-action table is never a formula target or
replaced by the paste.

## Reference reconciliation

The supplied full reference workbook contains 39,982 inbound legs, 937 valid
Q1 scores totaling 4,121, an average of 4.398078975, 143 scores `<=3`, and 794
scores `>3`. Its participation is 1,351 nonblank raw-Q1 legs divided by 9,661
`PCSStatus=1` legs, or 13.9840596%. The 414 invalid/non-score raw Q1 values are
kept in participation but excluded from the average.

The editable settings live in `config\wfm_rules.toml` under `[pcs]`. Changing
them requires a new `rulebook.version`, validation, and a refresh.

The low-sample threshold and any future PCS/participation targets live in
`config\metric_catalog.toml`. The workbook never invents a target when none is
configured.
