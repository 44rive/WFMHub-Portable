# WFMHub Power BI — WFM Control Tower V2 design review

Status: **proposal awaiting business validation**

These are deterministic 1680 × 945 visual prototypes with illustrative data.
They define the proposed Power BI page geometry, navigation, density and
decision surfaces. They do not change calculations or implement the Power BI
model.

## Proposed pages

1. [Daily WFM Command](01-daily-wfm-command.png)
2. [Service & SL Drivers](02-service-sl-drivers.png)
3. [Staff Preparation](03-staff-preparation.png)
4. [Workforce Realisation](04-workforce-realisation.png)
5. [Schedule Integrity & Patterns](05-schedule-integrity-patterns.png)
6. [Forecast Accuracy](06-forecast-accuracy.png)
7. [Data Readiness](07-data-readiness.png)

[Open the complete review board](WFMHub-PowerBI-WFM-Control-Tower-V2.png)

## Shared page contract

- WFM-only; PCS is intentionally excluded.
- Fixed left navigation and compact top status bar.
- Four page-appropriate slicers, four decision KPIs, two analytical surfaces
  and one evidence/action table.
- Management LOB is the primary service and planning scope.
- Native 15-minute facts remain available for interval diagnosis.
- Green, amber and red are used only for supported states.
- Percentages are ratios of summed components, never averages of displayed
  percentages.
- Power BI presents and filters; WFMHub Python/SQLite owns business logic,
  classification, recurrence detection and additive components.

## Truthfulness boundaries

- SL diagnosis reports a primary evidence-backed driver, contributing factors
  and residual. It does not claim unsupported causality.
- Schedule-integrity flags are review signals, not disciplinary scores.
- PTO/Away, incomplete current-day shifts and insufficient Agent Status evidence
  are excluded from unsupported schedule judgements.
- Staff Preparation never recommends moving an employee unless the configured
  LOB/language eligibility supports it.
- Missing required FTE remains blank and blocks coverage; it is never replaced
  with zero or an estimate.

## Validation requested

Validate the page names, navigation order, four headline KPIs per page, charts,
action tables, terminology and level of detail. Dummy values are not business
acceptance data.
