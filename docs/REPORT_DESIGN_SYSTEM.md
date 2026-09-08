# WFMHub report design system

| Field | Value |
|---|---|
| Contract | `WFMHUB-DESIGN` |
| Version | `1.0.0` |
| Status | Approved |
| Owner | Anass ASSRI / WFM |
| First compatible Hub version | `0.23.0` |

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

- Row 1: navy title; small freshness/status badge at the right.
- Row 2: one teal context sentence with dates and evidence state.
- Row 4: scope strip. A generated snapshot shows scope values, not fake filters.
- Rows 6–9: exactly four headline KPI cards.
- Main area: at most two charts, aligned to the same grid.
- Lower area: one ordinary Excel Table with native filters.
- Gridlines hidden; decision surface frozen; landscape print layout.
- Footer: `Prepared by Anass ASSRI | WFM` plus page count and confidentiality.

The approved visual preview is available from the v0.22.3 release assets. Its
numbers and generic LOB names are illustrative. Real reports use governed data.
A target line or pass/fail state is shown only when the effective metric catalog
defines a target; WFMHub never invents the pictured PCS target.

## Report blueprints

### PCS Operational Tracker

- `OVERVIEW`: Current MTD PCS, participation, prior comparable MTD PCS, change;
  PCS by LOB; daily current-versus-prior trend; visible per-LOB table.
- `RESULTS`: the actual linked Period View, Scope Level, LOB, Team Leader and
  Agent filters. Filter the table from left to right.
- `COACHING_QUEUE`: current exact low-score opportunities.
- `COACHING`: permanent blue action fields; never a Power Query target.
- `PCS_DATA`: refreshable agent-day counters for pivots or reconciliation.
- No dynamic-array report formulas. Power Query is transport only.

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

Staffing, Realisations, Final Absenteeism and Bonus may reuse these components,
but visual alignment does not promote them to Operational. Their calculations
and decision workflow must pass their own production review first.

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
- PCS is one permanent shared workbook. A normal update replaces only the four
  query tables and preserves `tblCoaching` byte-for-byte.
- Power Query never loads into `tblCoaching` or a permanent action ledger.
- Blue cells are editable. White, grey and calculated cells are not.
- Sheet names, table names, Agent ID, Gap ID and Coaching Key are contracts.

## Integrity gate

Every design change requires:

1. sheet order and fixed table-anchor tests;
2. formula inspection for `#REF!` and unsupported dynamic-array metadata;
3. ZIP integrity and reopen checks with `openpyxl`;
4. a Windows desktop Excel install/refresh/save/reopen smoke test for PCS;
5. verification that normal PCS update never rebuilds the permanent workbook;
6. verification that attendance decisions still import from Excel row 4;
7. a version bump when a permanent workbook must be rebuilt.

## Change log

- `1.0.0`: approved compact operational system for PCS, RTM/LOB Service Flash,
  and Attendance Review; canonical palette moved to `src/wfmhub/design.py`.
