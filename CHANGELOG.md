# Release notes

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
