# PCS PivotTable field guide

PCS is the only WFMHub product that uses Power Query and the Excel Data Model.
All other workbooks are complete generated snapshots with layouts suited to
their operational purpose.

Use [EXCEL_TEMPLATE_GUIDE.md](EXCEL_TEMPLATE_GUIDE.md) for the one-time query,
relationship, and measure setup.

## Main performance PivotTable

- Rows from `PCS Agents`: `lob`, `team_leader`, `agent_name`.
- Slicers from `PCS Agents`: `lob`, `team_leader`, `agent_name`, `language`.
- Timeline from `PCS Dates`: `date`.
- Values from `PCS Agent Day`: governed measures plus `valid_q1`,
  `score_le_3`, and `score_gt_3`.

Required measures:

```text
PCS Average = SUM(q1_score_sum) / SUM(valid_q1)
PCS Participation = SUM(q1_nonblank) / SUM(pcs_status_1)
```

Never average `pcs_average` or `participation_rate` columns.

## Coaching PivotTable

- Rows: Agent, then call date or Coaching Key.
- Filters: coaching status, date, LOB, and Team Leader.
- Values: Coaching Opportunities, Coaching Completed, and Actions Rate.

The editable decisions belong in `PCS Team.xlsx > COACHING_LOG`. The PCS Calls
model table is refreshed data and remains read-only.

## Common mistakes

- **Count of instead of Sum of:** use a supplied measure rather than dropping
  the score-sum field directly into Values.
- **Slicer controls one Pivot only:** build the slicer from `PCS Agents` or
  `PCS Dates`, then use Report Connections.
- **Relationship error:** Agent ID must be text in both facts; dates must be real
  dates, not Month labels.
- **Slow refresh:** all four queries must be Connection Only + Add to Data
  Model. None should load to a worksheet.
- **No new date:** run WFMHub **Sync and refresh PCS Team** before Excel
  **Refresh All**.
