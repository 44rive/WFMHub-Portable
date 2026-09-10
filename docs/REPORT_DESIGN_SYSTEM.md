# WFMHub report design system

| Field | Value |
|---|---|
| Contract | `WFMHUB-DESIGN` |
| Version | `3.2.0` |
| Status | Approved |
| Owner | Anass ASSRI / WFM |
| First compatible Hub version | `0.27.1` |

This is the visual contract for every WFMHub workbook. It changes presentation,
never business calculations, source scope, table keys, or report maturity.

## Design objective

The first screen must answer three questions without technical knowledge:

1. What period and operational scope am I seeing?
2. Which four numbers require attention?
3. Where is the exact team, queue, or agent action list?

Every workbook uses a compact title, an honest scope strip, no more than four
headline cards, no more than two decision charts in the first viewport, and one
filterable action or reconciliation table. Decorative visuals are prohibited.

## Code authority

`src/wfmhub/design.py` is the single code authority for tokens. Report builders
must import the palette; they must not create a competing palette.

| Token | Value | Use |
|---|---|---|
| Navy | `#0B1F33` | Titles and total rows |
| Teal | `#007C83` | Primary series and table headers |
| Gold | `#D6A84B` | Small separators and primary tabs |
| Canvas | `#F4F7F9` | Card labels and quiet background |
| Ink | `#1F2933` | Body text |
| Muted | `#536474` | Notes and comparison series |
| Line | `#D8E0E6` | Borders and gridlines |
| Blue | `#0563C1` / `#E3F0FA` | User-editable cells only |
| Green | `#1F7A53` / `#DDF3E8` | Proven healthy/completed state |
| Amber | `#A65F00` / `#FFF1CC` | Review or provisional state |
| Red | `#B42318` / `#FDE7E5` | Proven exception or missing requirement |
| White | `#FFFFFF` | Cards and title text |

Titles and KPI values use Aptos Display. Body, table and note text use Aptos.

## First-screen grid

- The canonical canvas is 1,456 pixels wide: 28 equal 52-pixel columns.
- Row 1 is a 56-pixel navy title/owner/status bar.
- Row 2 is a 52-pixel selector strip with four equal seven-column controls.
- Rows 4–6 contain four equal seven-column KPI cards.
- Rows 8–21 contain two equal 728 × 310 pixel native Excel charts.
- Rows 23–32 contain one compact seven-field action grid.
- Gridlines hidden; decision surface frozen; landscape print layout.
- Footer: `Prepared by Anass ASSRI | WFM` plus page count and confidentiality.

The approved visual prototypes are available from the v0.23.1 release assets.
Their numbers are illustrative. Real reports use governed data.
A data-free workbook reference for PCS and all six report families is shipped as
`docs/WFMHub Report Design Reference.xlsx`; it is generated from the same
production renderer so later visual work has one verifiable baseline.
A target line or pass/fail state is shown only when the effective metric catalog
defines a target; WFMHub never invents the pictured PCS target.

## Report blueprints

### PCS Report & Coaching

- `PCS Live Tracker.xlsx` is the one permanent shared workbook. Normal Hub
  updates replace external CSV feeds, never this file; a versioned migration
  archives the old tracker and preserves keyed actions.
- `OVERVIEW`: cascading Period, LOB, Team Leader and Agent dropdowns; selected
  PCS, participation, prior comparable and change cards; responsive LOB chart,
  daily trend, compact LOB table, then agent performance below.
- `PERFORMANCE`: lightweight standard-period results at LOB, team and agent
  grain. Native table filters are built in and native slicers may be added.
- `COACHING`: Period/LOB-filtered exact low-score table including Call ID;
  permanent compact blue `tblCoachingActions` on the right.
- `_PCS_FILTERS`, `_PCS_LOB`, `_PCS_AGENT`, `_PCS_DAILY`, and `_PCS_COACH` are
  hidden lightweight Power Query destinations. `_PCS_CALC` is the hidden chart
  bridge. There is no raw `DATA` worksheet.
- `SETUP` records the fixed local feed and six query definitions; `HELP`
  explains the two-step update; `_AUDIT` records the contract.
- Python/SQLite own all KPI arithmetic. Power Query is transport only. There is
  no Data Model, Power Pivot, macro, ODBC dependency, spill formula, or
  dynamic-array metadata.

### RTM Daily Control

- `CONTROL`: LOBs on target, demand variance, confirmed No Show HC, Call Now;
  service level by LOB; linked LOB action table.
- Each LOB tab: TSL, offered volume, forecast variance, confirmed No Show HC;
  full-day hourly demand/service; same-LOB attendance pulse and exact call list.
- Service Flash and attendance callout stay fused in this workbook. There is no
  separate Attendance Callout product.

### Attendance Review

- `CONTROL`: review gaps, gap hours, open decisions, missing evidence; by-LOB
  action summary and gap-hours chart.
- `REVIEW BOARD`: SCHEDULE directly above ACTUAL, followed by a visual spacer.
  Exact times and Gap ID remain authoritative; only five blue ACTUAL fields are
  editable.
- `BREAK & MEAL`: exact completed-shift control with evidence gating.
- Decision ledger and evidence remain hidden audit support.

### In-development workbooks

Staffing, Realisations, Final Absenteeism and Bonus use the same measured first
screen: four scope fields, four decision KPIs, two equal charts and a compact
prioritized action grid. Visual alignment does not promote them to Operational;
their calculations and decision workflows still require production review.

## Chart and table rules

- Bar charts compare LOBs or teams; line charts show time.
- Use ratio-of-sums values produced by the Hub, never averages of visible rates.
- No 3-D, gauges, donuts, gradients, decorative icons, or secondary decoration.
- Maximum two series unless the business question requires more.
- Targets come from the effective metric catalog and are labelled explicitly.
- Teal is current/primary; muted grey is comparison; semantic colors only mark
  states whose evidence is proven.
- Tables keep stable names and headers. Native filters are preferred to fake
  controls. Empty and missing are not converted to zero.

## Interaction and persistence

- Generated RTM and Attendance workbooks are dated decision snapshots.
- `PCS Live Tracker.xlsx` is permanent. Only a versioned, action-preserving
  migration may rebuild it; normal Hub updates never do.
- Each PCS update atomically replaces six fixed governed CSV feeds.
- Desktop Excel **Data > Refresh All** reloads query tables from those feeds.
- PCS dashboard and chart formulas read the query sheets through stable
  sheet-backed workbook names; direct formula references to replaceable
  `tblPcs*` query destinations are prohibited.
- Coaching actions remain in `tblCoachingActions` inside the permanent tracker.
- Blue cells are editable. White, grey and calculated cells are not.
- Sheet names, table names, Agent ID, Gap ID and Coaching Key are contracts.

## Integrity gate

Every design change requires:

1. sheet order and fixed table-anchor tests;
2. formula inspection for `#REF!`, `_xlfn`, dynamic-array metadata, and direct
   `tblPcs*` references from presentation sheets;
3. ZIP integrity and reopen checks with `openpyxl`;
4. verification that the pre-install PCS template contains no malformed query
   parts, Data Model, macros or external links and reopens without repair;
5. verification of all six fixed CSV schemas, Power Query destination tables,
   same-version tracker byte preservation, and migration action preservation;
6. verification that attendance decisions still import from Excel row 4;
7. a version bump when the report contract changes.

## Change log

- `3.2.0`: adds cascading Overview selectors and a filtered Coaching view using
  precomputed caches and classic exact lookups.
- `3.1.0`: replaces manual paste and the raw `DATA` sheet with five lightweight
  direct-CSV Power Query tables; adds agent performance below LOB performance,
  native filtering/slicer readiness and a compact side-by-side coaching log.
- `3.0.1`: removes spill formulas and dynamic-array metadata, makes Coaching a
  compact Period/LOB view, and adds an action-preserving tracker repair.
- `3.0.0`: replaces generated PCS snapshots and the separate coaching log with
  one permanent manual-paste tracker containing Overview, Coaching and DATA.
- `2.3.1`: replaces the PCS dashboard's array ranking formula with Python-
  pre-ranked view rows and direct `MATCH`/`INDEX` lookups so filtered charts and
  tables populate reliably in desktop Excel.
- `2.3.0`: restores cascading PCS Period/LOB/Team/Agent controls using only
  classic lookups over Python-precalculated results; standardizes visible date
  and datetime formats.
- `2.2.0`: PCS becomes a Python-only timestamped report plus a separate
  permanent coaching log; all presentation values and chart series are embedded.
- `2.1.0`: the measured V2 dashboard renderer is shared by RTM, every LOB
  Flash, Attendance Review, Staffing, Realisations, Absenteeism/Shrinkage and
  Bonus while their calculations, table names and editable ledgers remain
  report-owned.
- `2.0.0`: measured 28-column operational grid, equal cards/charts, real PCS
  selectors, hidden query staging and coaching Call ID.
- `1.0.0`: approved compact operational system for PCS, RTM/LOB Service Flash,
  and Attendance Review; canonical palette moved to `src/wfmhub/design.py`.
