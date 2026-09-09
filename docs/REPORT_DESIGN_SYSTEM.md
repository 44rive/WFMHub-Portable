# WFMHub report design system

| Field | Value |
|---|---|
| Contract | `WFMHUB-DESIGN` |
| Version | `3.0.1` |
| Status | Approved |
| Owner | Anass ASSRI / WFM |
| First compatible Hub version | `0.26.1` |

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

- `PCS Live Tracker.xlsx` is the one permanent shared workbook. Normal prepares
  never replace it; a versioned repair archives it and preserves keyed actions.
- `OVERVIEW`: Period, LOB, Team Leader and Agent dropdowns; PCS, participation,
  prior comparable and change cards; compact LOB comparison/chart; filtered
  agent results immediately below.
- `COACHING`: only Period and LOB selectors; automatically loaded exact low-score
  calls including Call ID; permanent blue action table on the same sheet.
- `DATA`: one plain Excel Table at inbound call-leg grain. Users replace only
  its body with rows from the newest `PCS Paste Data` workbook.
- `HELP`: the normal update workflow. `_LISTS` and `_AUDIT` are hidden support.
- No Power Query, Data Model, macro, Excel automation or external connection.
  Excel formulas divide additive counters after the manual paste; rates are
  never averaged.

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
  repair may rebuild it; normal prepares never do.
- Each PCS prepare creates a disposable timestamped clean paste-data workbook.
- Coaching actions remain in `tblCoachingActions` inside the permanent tracker.
- Blue cells are editable. White, grey and calculated cells are not.
- Sheet names, table names, Agent ID, Gap ID and Coaching Key are contracts.

## Integrity gate

Every design change requires:

1. sheet order and fixed table-anchor tests;
2. formula inspection for `#REF!`, `_xlfn` and dynamic-array metadata;
3. ZIP integrity and reopen checks with `openpyxl`;
4. verification that PCS contains no query connections, query tables, Data
   Model, macros or external links and reopens without repair;
5. verification that each PCS prepare creates a new clean paste file, same-
   version prepares do not change the tracker, and migrations preserve actions;
6. verification that attendance decisions still import from Excel row 4;
7. a version bump when the report contract changes.

## Change log

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
