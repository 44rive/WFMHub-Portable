# WFMHub Power BI — WFM cycle contract

Status: implemented source project; Power BI Desktop render acceptance required

Project contract: `7`

Feed schema: `7`
Design reference: `design-prototypes/powerbi-wfm-cycle-reset-v1/`

## Purpose

Power BI is the personal WFM control and prioritization layer. It follows the
normal WFM cycle; it is not a second source engine, an AI scorecard or a place
to enter corrections. WFMHub owns extraction, classification and additive
counters. Excel retains exact rows and the few durable human-action workflows.

## Five pages

| Page | Primary decision | Evidence |
|---|---|---|
| Today's Control | What needs attention now? | current service, forecast variance, observed capacity and attendance/capacity action rows |
| Staff Preparation | Where is published capacity below requirement? | 15-minute required, gross schedule, PTO/Away, net schedule and gap by Planning Group and Staff Type |
| Intraday Service | Which configured service or queue is moving the result? | combined Management LOB SL and exact selected-LOB queue components |
| Attendance & Schedule Review | Which published shift fragments remain unexplained? | native schedule-over-actual timeline, exact residual Gap ID/start/end and break/meal evidence |
| Historical Review | How did forecast, capacity and delivery close? | separate demand, SL, requirement, scheduled coverage, final absence and final shrinkage measures with comparison period |

Only native Power BI visuals are allowed. The attendance timeline is a native
stacked horizontal `barChart`; an invisible clock-offset series positions each
schedule or actual segment. No marketplace visual is required.

## Filter model

The navigation is a native horizontal page navigator. Date, Management LOB,
Planning Group, Staff Type, Team Leader, Agent and Queue come from conformed
dimensions. The same dimension uses the same slicer synchronization group on
every applicable page.

Service and capacity remain separate:

- service: exact queue -> combined Management LOB;
- capacity: Verint Staff Type -> Planning Group -> Management LOB;
- RSA BE service is combined, while RSA BE FR and RSA BE VL capacity remains
  separate;
- a forecast field labelled `Queue Name` is a Staff Type identity, not a Storm
  routing queue;
- only `MAPPED` capacity identities enter Staff Type/Planning Group visuals.

The Historical Review comparison period is intentionally disconnected. Its DAX
applies the selected comparison month to the same conformed measures without
polluting organisational relationships.

## Data ownership

WFMHub Python/SQLite owns parsing, date scope, identity, exact service
components, PTO/Away overlays, attendance, final-Activities overlap and CSV
publication. Power BI owns relationships, explicit ratio-of-sums measures,
filter context and presentation.

The feed is published atomically under `Feed\PowerBI`; the manifest is written
last. The PBIP imports those fixed CSVs through one `HubRoot` parameter. It has
no SQLite, DuckDB, ODBC, raw-extract or PCS connection.

## Excel handoff

| Workbook | Power BI handoff | Editable content |
|---|---|---|
| `RTM Daily Control.xlsx` | Today's Control and Intraday Service exact Flash/callout rows | none |
| `Staffing Preparation.xlsx` | Staff Preparation 15-minute capacity detail | persistent `ACTIONS` ledger keyed by Capacity Key |
| `Attendance Review.xlsx` | exact schedule, Agent Status and residual gap evidence | none; correct in Verint and refresh |
| `Realisations.xlsx` | detailed historical service/capacity results | none |
| `Final Absenteeism & Shrinkage.xlsx` | final Verint components and exceptions | none |

PCS and Bonus remain outside this Power BI model.

## Layout and validation

Canvas: 1680x945. Header: y0/h64. Native horizontal navigation: y68/h48.
Selector strip: y124/h77. Four compact cards start at y211. Analysis begins at
y355; evidence/action tables start at y684 or y718. Geometry, copy and colors
follow the five approved PNGs.

Automated validation proves page names, JSON/TMDL structure, native visual
types, slicer groups, field bindings, feed headers and measure/column names.
Power BI Desktop remains the final renderer, so every new release needs one
Windows open, Refresh and visual inspection before operational use.
