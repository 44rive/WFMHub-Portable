# WFMHub Power BI — WFM cycle contract

Status: implemented source project, Desktop render acceptance required
Project contract: `6`
Feed schema: `6`
Design reference: `design-prototypes/wfm-manager-cycle-v1/`

## Purpose

This is a personal WFM management cockpit, not a second report factory and not
an invented KPI scorecard. It follows the WFM cycle from planning through
delivery and review. Excel remains available when exact operational rows must be
shared; PCS keeps its permanent collaborative workbook.

## Pages and bindings

| Page | Cards | Analysis panels | Bottom evidence |
|---|---|---|---|
| Forecast & Requirement | Forecast Volume, Required FTE-hours, Peak Required FTE, Requirement Coverage | native interval Volume + requirement; Staff Type requirement summary | source Staff Type rows |
| Staff Preparation | Peak Required, Peak Gap, Uncovered FTE-hours, PTO/Away impact | required vs net scheduled; Planning Group/Staff Type position | staffing gaps to treat |
| Intraday Control | SL, Offered, Present FTE Gap, confirmed No Show | combined LOB SL; split resource position | capacity and attendance actions |
| Attendance & Schedule Review | No Show, Late, Early Leave, residual hours | published/observed/residual placement; break/meal evidence | exact residual correction queue |
| Performance Review | SL, Volume variance, scheduled coverage, final absence | weekly forecast vs actual; requirement-to-delivery bridge | monthly WFM scorecard |

## Semantic grain

Service relationships are Date -> Time -> exact Service Key -> Management LOB.
They come only from reviewed service-profile queue allowlists.

Capacity relationships are Date -> Time -> Staff Type -> Planning Group ->
Management LOB. Forecast identity is filename plus source Staff Type. Scheduled
identity is workforce LOB plus published assignment. The domains never join on
similar-looking text.

Employee facts use Agent ID -> Employee -> Management LOB. Attendance detail may
show Planning Group/Staff Type resolved at build time without creating ambiguous
relationship paths.

## Data ownership

WFMHub Python/SQLite owns parsing, identity, date scope, source classification,
exact SL components, attendance, PTO/Away overlays, final-Activities overlap and
published CSV rows. Power BI owns relationships, explicit ratio-of-sums DAX,
filter context and presentation.

The feed is published under `Feed\PowerBI` and the manifest is written last.
The PBIP imports the CSVs through the single `HubRoot` parameter. It contains no
SQLite, DuckDB, ODBC or raw-source connection.

## Layout contract

Canvas: 1680x945. Rail: x0–210. Header: x210–1680, height 72. Content starts at
x228. Filter strip y85/h58. KPI cards y153/h119. Analysis y282/h330. Detail
y622/h248. Footer y910. Colors, spacing, page copy and panel proportions follow
the five approved PNGs and `cycle-pages.html`.

`tools/build_powerbi_project.py` builds to a temporary sibling directory,
validates every JSON file, the five-page set and all semantic tables, and only
then replaces the checked-in PBIP. This prevents a failed generation from
deleting the last valid project.

## Validation gate

- generated page count and exact names;
- JSON parse of every artifact;
- every visual binding exists in the semantic model;
- no measure/column name collision;
- every referenced feed CSV exists and headers match;
- no forbidden raw/SQLite/PCS data source in TMDL;
- complete automated suite and portable smoke test;
- one manual Power BI Desktop open/refresh/render check before operational use.
