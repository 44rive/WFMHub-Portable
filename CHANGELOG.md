# Release notes

## 1.1.1

- Added an independent permanent PCS workbook path for a locally synced WFM
  SharePoint library. First use copies the old tracker once without removing
  its coaching or changing other report/database locations; later updates
  sync the same shared file and keep its link stable.

## 1.1.0

- Fixed RTM overnight carry-over: prior-day night shifts now contribute their
  after-midnight Agent Status to today's Logged, Productive, Available and
  staffing intervals without adding yesterday's attendance to today's KPIs.
- Replaced the fragile PCS Power Query installer with an in-place Excel CSV
  sync. The permanent tracker keeps its ordinary tables, dropdowns, charts and
  human-owned coaching actions; old trackers migrate once with coaching backup.
- Fixed the shared dashboard charts to plot hidden helper data, corrected
  Bonus population payout-rate references, and made KPI attainment readable by
  population and tenure.
- Added Tenure Group to Bonus Raw_Data and separate Untenured KPI configuration.
  The original matrix remains the Tenured baseline; missing Untenured targets
  block final payout until approved rather than inventing a third tier.

## 1.0.0

- Established one clean WFMHub product baseline and current SQLite schema.
- Centralized source roles, employee scope, service, staffing, attendance,
  absence, PCS, configuration, design, and release contracts in `PROJECT.md`.
- Shipped RTM Daily Control, Attendance Review, PCS, staffing, realizations,
  final absence/shrinkage, bonus, clean exports, deterministic analysis, and
  the localhost manager workbench as one supported product surface.
- Removed development artifacts and operational source data from the public
  repository.
- Made the bonus matrix a permanent formula-driven workbook with the uploaded
  v1.2 six-population KPI tiers, and separated PCS filter staging columns to
  prevent dropdown-list cross-contamination.
- Added RTM hourly staffing headcounts, strict completed-shift no-show logic,
  and named agent-state context for below-target queue-hours.
- Added reconciled LOB and agent scheduled-hour allocation to Realisations.
