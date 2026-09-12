# WFMHub Power BI — premium design proposal V1

Status: **superseded by `POWERBI_WFM_CONTROL_TOWER_V2.md` and retained as a
historical design reference**

The existing Excel reports remain part of WFMHub. Power BI is a separate
analytical and management layer. Prototype values shown below are illustrative;
project contract 3 implements the same 1680x945 page geometry and binds it to
governed `Feed\PowerBI` facts and explicit measures.

## Shared experience

All seven pages use one persistent shell:

- WFMHub navy navigation and header;
- active page highlighted in teal;
- one compact global selector strip with four page-appropriate slicers;
- exactly four decision KPI cards;
- at most two analytical visuals before the action area;
- compact, filterable decision tables, including split agent/coaching and
  issue/governance work areas where the approved prototype requires them;
- green, amber, and red only for evidence-backed states;
- footer ownership: `Prepared by Anass ASSRI | WFM`.

## 1. Executive Overview

One management screen combining Service, Forecast, Attendance, PCS, and the
highest-priority cross-domain actions.

![Executive Overview](design-prototypes/powerbi-premium-v1/01-executive-overview.png)

[Open full-size Executive Overview](design-prototypes/powerbi-premium-v1/01-executive-overview.png)

## 2. Service & Forecast

Intraday 15-minute demand versus forecast, LOB TSL comparison, and the queues
that explain performance.

![Service and Forecast](design-prototypes/powerbi-premium-v1/02-service-forecast.png)

[Open full-size Service & Forecast](design-prototypes/powerbi-premium-v1/02-service-forecast.png)

## 3. Attendance Control

Same-day attendance pulse, scheduled-versus-present staffing, and exact call or
follow-up population. Unknown evidence stays separate from confirmed no-show.

![Attendance Control](design-prototypes/powerbi-premium-v1/03-attendance-control.png)

[Open full-size Attendance Control](design-prototypes/powerbi-premium-v1/03-attendance-control.png)

## 4. PCS Performance & Coaching

Current versus prior PCS, participation, agent results, and call-level coaching
queue with visible Call ID.

![PCS Performance and Coaching](design-prototypes/powerbi-premium-v1/04-pcs-performance-coaching.png)

[Open full-size PCS Performance & Coaching](design-prototypes/powerbi-premium-v1/04-pcs-performance-coaching.png)

## 5. Staffing & Capacity

Required versus scheduled FTE at 15-minute grain, LOB/day coverage heatmap, and
the critical future gaps requiring an operational decision.

![Staffing and Capacity](design-prototypes/powerbi-premium-v1/05-staffing-capacity.png)

[Open full-size Staffing & Capacity](design-prototypes/powerbi-premium-v1/05-staffing-capacity.png)

## 6. Absence & Shrinkage

Absenteeism trend, transparent shrinkage composition, and the exact cases still
requiring review.

![Absence and Shrinkage](design-prototypes/powerbi-premium-v1/06-absence-shrinkage.png)

[Open full-size Absence & Shrinkage](design-prototypes/powerbi-premium-v1/06-absence-shrinkage.png)

## 7. Data Quality & Governance

Source freshness, mapping and status coverage, issue history, active blockers,
and the exact configuration versions behind the dashboard.

![Data Quality and Governance](design-prototypes/powerbi-premium-v1/07-data-quality-governance.png)

[Open full-size Data Quality & Governance](design-prototypes/powerbi-premium-v1/07-data-quality-governance.png)

## Validation boundary

The implemented project does not approve illustrative prototype numbers,
invent KPI targets, change current Hub calculations, retire Excel reports, or
authorize publication. Power BI Desktop rendering remains the final visual
acceptance boundary; the source PBIR passes Microsoft's report-authoring
validator before release.

Staffing has one additional truthfulness rule: **Scheduled FTE** is capacity
derived from StartEndTimes and reduced by effective PTO/Away. **Required FTE**
comes only from an explicit Verint requirement column. When that column is not
supplied, Required FTE, Net Gap and Coverage remain blank.
