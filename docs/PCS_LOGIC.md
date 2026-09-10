# PCS calculation and workbook logic

The official calculation is reproduced from `TOLEARN/PCS Report.xlsx`,
especially `OverView!O15:R15` and the embedded `RDATA` query. The workbook is a
presentation and collaboration surface; Python/SQLite own the arithmetic.

## What a call leg is

A call may pass through a queue, transfer, or agent more than once. Each such
record is a call leg. WFMHub keeps that source grain and deduplicates overlapping
extracts by a stable call key. `transferred_legs` is descriptive and does not
change the PCS formula.

## Exact formula

Only inbound legs with an effective-dated in-scope Agent ID enter official PCS
agent counters.

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
diagnostic count but is not the participation denominator. At team, LOB, day,
week, and month level, additive counters are summed first and the ratio is
calculated once. Agent averages and percentages are never averaged together.

## Coaching and Actions Rate

The original `OverView!S15` formula counted completed briefing rows from an
external personal workbook and divided them by the agent/date count of valid Q1
scores `<=3`. WFMHub preserves the business meaning without the broken link:

- every valid inbound Q1 `<=3` creates one coaching opportunity;
- each opportunity is identified by the stable Coaching Key and exact Call ID;
- `Actions Rate = unique completed Coaching Keys / all coaching opportunities`;
- the filtered lookup table on `COACHING` exposes the low-score calls;
- the reviewer copies Coaching Key and Call ID to `tblCoachingActions`, then
  records status, coach, dates, and comment;
- `tblCoachingActions` is permanent and never a Power Query destination;
- `Not required` remains in the denominator and is not Completed;
- duplicate Coaching Keys are highlighted and cannot increase a completed
  count.

Low sample is an interpretation warning, not a coaching opportunity by itself.

## PCS data pipeline

```text
untouched FTE + Call by Call
          -> targeted PCS ingestion and effective-dated roster scope
          -> deduplicated clean call legs
          -> additive agent/day PCS mart
          -> six governed CSV feeds
          -> Power Query transport
          -> permanent PCS Live Tracker.xlsx
```

The fixed files under `Feed\PCS` are:

| File | Workbook destination | Purpose |
|---|---|---|
| `PCS_FILTER_LIST_CURRENT.csv` | hidden `_PCS_FILTERS` | governed cascading dropdown values |
| `PCS_LOB_SCORECARD_CURRENT.csv` | hidden `_PCS_LOB` | selection-aware cards and LOB comparison cache |
| `PCS_AGENT_SCORECARD_CURRENT.csv` | hidden `_PCS_AGENT` | selection-aware agent cache |
| `PCS_DAILY_SCORECARD_CURRENT.csv` | hidden `_PCS_DAILY` | selected-period daily trend cache |
| `PCS_RESULTS_CURRENT.csv` | `PERFORMANCE` | period/scope table for native filters and slicers |
| `PCS_COACHING_OPPORTUNITY_CURRENT.csv` | hidden `_PCS_COACH` | period/LOB coaching cache with exact calls |

CSV replacement is atomic. A failed Hub calculation leaves the previous
complete feed in place. Power Query performs no KPI arithmetic; it checks the
required schema, applies types, and loads the corresponding Excel table.

## Workbook lifecycle

`PCS Live Tracker.xlsx` is created or migrated only when its versioned contract
changes. That migration archives the prior workbook and carries keyed coaching
actions forward. Once the current version exists, a normal Hub PCS update does
not open, replace, or modify it. The user opens the workbook and chooses **Data
> Refresh All** after the CSV feeds are updated.

There is no raw `DATA` worksheet, Data Model, Power Pivot, macro, ODBC driver,
spill formula, or dynamic-array metadata. `OVERVIEW` uses four dropdowns and
classic exact `INDEX/MATCH` lookups. `COACHING` uses Period and LOB dropdowns.
`PERFORMANCE` remains a native Excel table for detailed filtering or optional
standard slicers.

The query installer replaces each starter Excel table with a real Power Query
destination. Excel rewrites formulas that directly reference a deleted table to
`#REF!`, so every presentation lookup reads through four stable sheet-backed
workbook names: `PCS_LOB_DATA`, `PCS_AGENT_DATA`, `PCS_DAILY_DATA`, and
`PCS_COACH_DATA`. The installer validates those names and all formulas on
`OVERVIEW`, `COACHING`, and `_PCS_CALC` before saving. If integrity fails, it
closes without saving the damaged state.

Selectable comparisons are fixed and explicit: Latest day versus the previous
available data day; Current week versus the same weekdays one week earlier;
Current MTD versus the same day span in the prior month; Previous MTD same days
versus the same span two months earlier; and Previous full month versus the
month before it. Missing comparison data stays blank and is never converted to
zero.

## Reference reconciliation

The supplied full reference workbook contains 39,982 inbound legs, 937 valid Q1
scores totaling 4,121, an average of 4.398078975, 143 scores `<=3`, and 794
scores `>3`. Its participation is 1,351 nonblank raw-Q1 legs divided by 9,661
`PCSStatus=1` legs, or 13.9840596%. The 414 invalid/non-score raw Q1 values are
kept in participation but excluded from the average.

Editable parsing settings live in `config\wfm_rules.toml` under `[pcs]`.
Changing them requires a new `rulebook.version`, validation, and a refresh. The
low-sample threshold and any future PCS/participation targets live in
`config\metric_catalog.toml`. The workbook never invents a target when none is
configured.
