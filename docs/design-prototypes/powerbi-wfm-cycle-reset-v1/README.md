# WFM Power BI cycle reset — validation prototype

Status: approved implementation reference for Power BI project contract 8.

## Page contract

All five pages use one native horizontal page navigator and one conformed
filter model. Date, Management LOB, Planning Group, Staff Type, Team Leader and
Agent come from shared connected dimensions; slicers are synchronized wherever
the dimension is relevant. A selection never migrates into an unrelated field.

1. **Today's Control** answers what needs attention now: current service,
   forecast variance, present capacity and the attendance/staffing action list.
2. **Staff Preparation** compares 15-minute required capacity with published
   schedules and PTO/Away impact at Planning Group and genuine Verint Staff Type
   grain.
3. **Intraday Service** owns the combined operational LOB service result and
   exact selected-LOB queue diagnosis. It does not contain capacity Staff Types.
4. **Attendance & Schedule Review** is the operational centerpiece. A native
   stacked bar uses an invisible clock-offset series to show the published
   schedule above Agent Status segments without installing a custom visual.
   Exact residual intervals remain in the synchronized table below.
5. **Historical Review** closes the WFM cycle using separate forecast, required,
   scheduled, observed, service, final absence and final shrinkage measures. It
   does not manufacture a composite score.

Only native Power BI visuals are allowed. The implementation must not depend on
marketplace visuals, custom icons, bookmarks for basic navigation, disconnected
slicer tables or activity labels pretending to be Staff Types.

## Excel companion contract

Power BI is the monitoring and prioritization layer. Excel is retained only
where exact rows, portable evidence, Verint treatment or persistent human input
are operationally useful.

| Workbook | Power BI handoff | Required detailed content | Editable content |
|---|---|---|---|
| `RTM Daily Control.xlsx` | Today's Control and Intraday Service | LOB flash, queue detail, current attendance call list and evidence | None; generated snapshot |
| `Staffing Preparation.xlsx` | Staff Preparation | 15-minute requirement, gross schedule, PTO/Away, net schedule and exact capacity gaps by Planning Group and Staff Type | One persistent staffing action table: gap key, action, planned FTE, owner, status and note |
| `Attendance Review.xlsx` | Attendance & Schedule Review | Published schedule, Agent Status timeline, exact residual Gap ID/start/end/minutes, final Verint overlap, break and meal totals | None; correction is completed in Verint and disappears after Activities refresh |
| `Realisations.xlsx` | Historical Review | Forecast versus actual volume and required/scheduled/observed/productive FTE-hours by period and LOB | None; generated evidence |
| `Final Absenteeism & Shrinkage.xlsx` | Historical Review | Final Verint absence/shrinkage totals and exact components by agent/day | None; generated evidence |
| `PCS Live Tracker.xlsx` | Outside initial Power BI scope | PCS results, Call ID detail and coaching history | Existing persistent coaching table only |
| `Bonus Management.xlsx` | Outside initial Power BI scope | Governed monthly bonus result and component detail | Existing management workflow only |

Every detail workbook must expose the same handoff keys used in Power BI:
Business Date, Management LOB, Planning Group, Staff Type where applicable,
Agent ID, interval start/end, and the stable Gap/Capacity/Call key. It should
open on a compact control sheet showing the selected scope, last Hub refresh and
data-through date. Power Query is not required for these generated snapshot
workbooks; PCS keeps its existing permanent collaborative workflow.

## Prototype images

- `01-Todays-Control.png`
- `02-Staff-Preparation.png`
- `03-Intraday-Service.png`
- `04-Attendance-Schedule-Review.png`
- `05-Historical-Review.png`

Values and employee names in the images are illustrative. Business grains,
labels and required visual types are the proposed implementation contract.
