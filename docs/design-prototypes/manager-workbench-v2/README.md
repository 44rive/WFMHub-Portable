# WFMHub Manager Workbench V2

Clickable design prototype for the proposed local WFM manager application. The
values are illustrative. This design does not change production application
code, calculations, the database, or source extracts.

Open [the clickable prototype](index.html), then use the top navigation and
workspace tabs. The static [all-pages review board](all-pages-review.png) is the
fastest way to inspect the complete visual system.

For local use, download [WFMHub-Manager-Workbench-V2.zip](WFMHub-Manager-Workbench-V2.zip),
extract it, and open `index.html` in a browser. No installation or server is
needed for this design prototype.

## Pages

- Manager Desk: one prioritized view across now, end of day, next week, and MTD.
- Plan: Demand & Requirement, Capacity & Schedule, and Scenario Lab.
- Operate: Live Service and Attendance Pulse.
- Review: Schedule Integrity, Realisations, Absence & Shrinkage, and Workforce Patterns.
- Deliver: Reports & Analysis and Generated Archive.
- Govern: Data Readiness, Mappings & Rules, and Jobs & Logs.

## Business boundaries shown in the prototype

- Service and capacity are separate grains.
- RSA Belgium service is combined; FR and VL capacity stay separate.
- Agent Status is the primary attendance evidence.
- Verint Activities finalize absence/shrinkage and remove corrected overlaps.
- Active FTE, date-aware leavers, PTO, and Away affect people/capacity views.
- Missing evidence remains Unknown; the interface does not invent a result.
- Service level uses exact configured queues and ratios of sums.
- Source extracts are read-only.
