# PivotTable field guide

Use [EXCEL_TEMPLATE_GUIDE.md](EXCEL_TEMPLATE_GUIDE.md) for the one-time button
clicks. This page tells you which fields belong in each PivotTable.

The queries must read the compact CSVs in `output\model_data\<report>` and load
as **Connection Only + Add to Data Model**. Do not load raw extracts or model
tables into worksheets.

## PCS Performance

Use `agent_detail.csv`.

- Rows: `agent_name`.
- Slicers: `lob`, `team_leader`, `agent_name`, `language`, `sample_state`.
- Values to sum: `q1_score_sum`, `valid_q1`, `q1_nonblank`, `pcs_status_1`,
  `score_le_3`, and `score_gt_3`.

Create Data Model measures that divide summed counters:

```text
PCS Average = SUM(q1_score_sum) / SUM(valid_q1)
PCS Participation = SUM(q1_nonblank) / SUM(pcs_status_1)
```

Never average `pcs_average` or `participation_rate`.

## Service Performance

Use `intraday.csv` for management and `queue_detail.csv` for audit.

- Rows: `hour_start`.
- Columns: `service_group`.
- Slicers: `business_date`, `service_group`.
- Values to sum: `offered`, `answered`, `answered_within_target`,
  `short_abandoned`, and `handled_seconds` when present.

Measures:

```text
Ford OEM Gross SL = SUM(answered_within_target) / SUM(offered)
Service Availability = SUM(answered) / SUM(offered)
AHT = SUM(handled_seconds) / SUM(answered)
```

The chosen service profile controls which mapped scopes and governed SL method
enter the files. If a different profile selects adjusted SL, its measure is
`SUM(answered_within_target) / (SUM(offered) - SUM(short_abandoned))`. Check
`service_method` and the report `DEFINITIONS` sheet instead of guessing.

## Staffing & Coverage

Use `intraday.csv`.

- Rows: `interval_start`.
- Columns: `lob`, then `language`.
- Slicers: `business_date`, `lob`, `language`, `staffing_state`.
- Values: sum `scheduled_fte`, `observed_fte`, `productive_fte`, and
  `staffing_gap_fte`.

Treat `DATA_MISSING` and future intervals as unknown. Do not replace blank gaps
with zero.

## Attendance Today

Use `actions.csv`.

- Rows: `agent_name`.
- Slicers: `lob`, `team_leader`, `call_action`, `attendance_result`.
- Useful columns: scheduled start/end, first/last evidence, late minutes,
  no-show minutes, source-loaded flag, and provisional flag.

This is a contact queue, not an adherence table.

## Attendance Corrections

Use `gaps.csv` for actions and `timeline.csv` for a planned-versus-observed
visual.

- Rows: `agent_name`, then gap start.
- Slicers: `lob`, `team_leader`, `detected_issue`, `validation_status`.
- Values: sum gap minutes.

Edit and import decisions only in the generated workbook `GAPS` sheet. The
Data Model copy is for analysis and should stay read-only.

## Final Absence & Shrinkage

Use `agent_detail.csv`.

- Rows: `agent_name`.
- Columns: `business_date`; group dates by month if needed.
- Slicers: `lob`, `team_leader`, `language`, `final_ledger_status`.
- Values: sum planned net, final absence, vacation, unpaid, shrinkage, and
  unmapped hours.

Measure:

```text
Final Absence Rate = SUM(final_absence_hours) / SUM(planned_net_hours)
```

Never average `final_absence_rate`.

## Bonus Performance

Use `agent_detail.csv` and `kpi_analysis.csv`.

- Rows: `population`, then `agent_name`.
- Slicers: `period`, `population`, `release_status`, `eligibility`.
- Values: scenario payout, released payout, and achievement components.

Do not present Scenario Payout as payroll output while release rows are blocked.

## Common mistakes

- **Count of** instead of **Sum of**: open Value Field Settings and select Sum.
- Percentage above 100%: divide summed numerators by summed denominators.
- Missing agents: check the FTE roster and Source Health.
- Missing dates: check date coverage and the period used to build the report.
- Refresh error: unhide `_AUDIT` and verify the Template model folder path.
- Empty Pivot: inspect `manifest.json`; it states the exact row count per CSV.
