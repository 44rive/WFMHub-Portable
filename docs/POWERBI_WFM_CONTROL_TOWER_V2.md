# WFMHub Power BI — WFM Control Tower V2

| Field | Value |
|---|---|
| Report contract | `WFMHUB-PBI-5` |
| Feed schema | `5` |
| Canvas | `1680 × 945` |
| Owner | Anass ASSRI / WFM |
| Status | Implemented; final Power BI Desktop rendering acceptance required |

This is the implementation reference for the source-controlled project at
`templates/powerbi/WFMHub BI`. The approved full-size visual references remain
under `docs/design-prototypes/powerbi-wfm-control-v2`.

## Approved shell

Every page has the same measured layout:

- fixed 178-pixel navy navigation;
- 58-pixel title bar and teal separator;
- one compact strip with four slicers;
- four equal decision cards with a semantic accent and plain-language note;
- two equal analytical panels;
- one full-width evidence/action table;
- `Prepared by Anass ASSRI | WFM` footer.

The left rail contains the native page navigator and an always-visible text
fallback. Standard Power BI page tabs remain enabled. This makes navigation
usable even when a Desktop build delays or fails to paint navigator captions.

No custom visual is required. The report uses native Power BI cards, slicers,
tables, matrices, bar/column, line and combo charts so it remains portable on a
restricted work computer.

## Eight pages and exact intent

1. **Daily WFM Command** — LOBs on target, demand variance, confirmed no-show,
   productive capacity gap, SL by LOB, capacity pulse and ordered findings.
2. **Service & SL Drivers** — 15-minute SL/demand, supported demand/schedule/
   attendance/AUX/AHT pressure signals and exact interval/queue evidence.
3. **Staff Preparation** — explicit required FTE versus net scheduled FTE,
   approved PTO/Away impact, daily coverage matrix and future gap table.
4. **Workforce Realisation** — observed-time composition and three independent
   ratios: observed/elapsed, productive/observed and productive/elapsed.
5. **Schedule Integrity & Patterns** — completed published versus observed
   placement, supported recurrence and exact evidence. It is not adherence and
   not a disciplinary score.
6. **Forecast Accuracy** — ratio-of-sums 15-minute WAPE accuracy, signed bias,
   weighted AHT error, peak-demand accuracy and comparable-grain misses.
7. **Absence & Shrinkage** — final Verint absence/shrinkage rates, component
   mix, LOB comparison and exact agent/day completeness evidence.
8. **Data Readiness** — source freshness, queue reference coverage, mapped
   Agent Status minutes and blocking evidence.

## Data contract

Power BI imports only fixed CSVs from `Feed\PowerBI`. Important V2 additions:

- `FactService15Min.csv` — additive Call-by-Call counters at native quarter hour;
- `FactQueueCoverage.csv` — mapped inbound entries divided by all inbound entries;
- `FactStatusInterval.csv` — schedule-clipped Agent Status/AUX segments;
- `FactScheduleIntegrity.csv` — one completed integrity result per agent day;
- `FactShiftPlacement.csv` — published and observed rows for the placement chart;
- `DimDriver.csv` and `DimCapacityStage.csv` — stable display bridges.

`DimQueue`, Service, Forecast and Queue Coverage also carry `Service Key`, a
stable `management LOB|service scope|queue` composite. Relationships use that
key so repeated display queue names cannot contaminate another LOB. Management
LOB filters Employee and Queue, which then filter their downstream facts. This
is the required cross-page filter path for LOB -> Team Leader -> Agent.

The existing hourly service mart still powers validated Flash reports. PCS CSVs
are retired from the Power BI contract only; the PCS Excel feed and permanent
coaching tracker are unchanged.

## Business boundaries

- Power BI never reads extracts or SQLite.
- The Hub owns identity, active/leaver scope, queue mapping, PTO/Away, status
  classification and additive counters.
- Missing required FTE stays blank; it never becomes zero.
- Staff Preparation filters both demand and capacity by planning horizon,
  management LOB, weekday and operating interval. It does not pretend a
  roster-language filter changes required FTE because no approved
  queue-to-language eligibility map has been supplied.
- Realisation percentages divide summed governed minutes.
- Schedule integrity uses Agent Status first and LILO only as fallback. It
  excludes unfinished shifts and any day with effective PTO/Away.
- Driver values are pressure signals for review. They are not exact causal
  contributions to service level.

## Operator workflow

```text
new extracts -> WFMHub UPDATE -> SQLite marts -> Feed\PowerBI CSVs
             -> POWERBI.cmd -> Power BI Desktop Home > Refresh
```

The first `POWERBI.cmd` run after this contract upgrade archives the prior
project and installs V5. Later opens preserve the installed V5 project. Save a
`.pbix` copy from Power BI Desktop only after the eight pages refresh without an
error.

## Validation boundary

Repository checks validate migrations, feed headers, table/measure collisions,
all PBIR visual bindings, page order/geometry and deterministic generation.
Linux cannot render Power BI Desktop. The final release gate on the work PC is:

1. run one complete Hub update;
2. open with `POWERBI.cmd`;
3. choose **Home > Refresh**;
4. inspect all eight pages at Fit to page;
5. confirm slicers and cross-highlighting;
6. save the working PBIX snapshot.
