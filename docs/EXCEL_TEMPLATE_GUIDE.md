# Excel PivotTable and slicer setup

This is a one-time setup for each report. WFMHub keeps every calculation in
Python/SQLite and gives Excel only small, report-ready CSV tables.

## What happens every day

1. Run **Build reports** in `WFMHub.cmd`.
2. WFMHub creates the normal static workbook and refreshes the compact files in
   `output/model_data/<report>/`.
3. Open your saved Excel master, choose **Data > Refresh All**, and save a copy
   for management.

The three-hour PCS rhythm is only when you repeat these steps. It does not
change the daily or monthly PCS formulas.

## One-time setup for PCS

1. In `WFMHub.cmd`, choose **Create an Excel Pivot/slicer master**.
2. Choose **PCS Performance** and a populated period.
3. Open `templates/reports/pcs.xlsx`. WFMHub protects this master from normal
   refreshes and refuses to replace it accidentally.
4. Open **Data > Get Data > Launch Power Query Editor**.
5. Choose **New Source > Other Sources > Blank Query**.
6. Open **Advanced Editor**.
7. Copy `templates/power_query/WFMHubCsv.pq` into the editor.
8. Replace `TABLE_FILE.csv` with `agent_detail.csv`.
9. Rename the query `PCS Agent Detail`.
10. Choose **Close & Load To...**.
11. Select **Only Create Connection** and tick **Add this data to the Data
    Model**. Do not load it to a worksheet.
12. Repeat for `trend.csv` and `actions.csv` if you need those PivotTables.
13. Choose **Insert > PivotTable > From Data Model**.
14. Put the PivotTable on `DASHBOARD` or a new `PIVOT_VIEW` sheet.
15. Choose **PivotTable Analyze > Insert Slicer** and add:
    - LOB
    - Team Leader
    - Agent
16. Save the master.

The command-line equivalent for step 1 is:

```text
runtime\python.exe -m wfmhub --home . template-init --pack pcs --start 2026-08-01 --end 2026-08-31
```

Do not use `--force` after you add PivotTables or slicers; it intentionally
replaces the master with a new starter.

## Recommended PivotTable fields

### PCS

- Rows: `agent_name`
- Filters/slicers: `lob`, `team_leader`, `agent_name`
- Values: sum of `q1_score_sum`, sum of `valid_q1`, sum of `q1_nonblank`, sum
  of `pcs_status_1`
- Measures must divide the sums. Never average `pcs_average` or
  `participation_rate`.

### Service

- Rows: `hour_start`
- Columns: `service_group`
- Filters/slicers: `business_date`, `service_group`
- Values: summed additive counters. Calculate SL and availability from sums;
  Ford OEM uses gross SL exactly like the original Flash. Check
  `service_method` before creating the measure for another profile.

### Attendance and staffing

- Attendance slicers: `lob`, `team_leader`, `call_action`
- Staffing slicers: `lob`, `language`, `staffing_state`

### Bonus

- Slicers: `period`, `population`, `release_status`
- Never use Scenario Payout as payroll while any release control is blocked.

## If Refresh All cannot find the files

1. Unhide `_AUDIT`.
2. Find **Template model folder**.
3. Confirm the path points to the matching folder under
   `output/model_data/`.
4. Save the workbook and refresh again.

## What to send

After **Refresh All** finishes, use **File > Save As** and put the dated copy in
the normal product folder under `output`. Send that copy, not the master. The
generated static workbook is the immutable audit snapshot; the master is the
interactive current view.

The named Excel cell `pModelDataPath` points to this value. Each query reads one
CSV directly, so there are no query-to-query privacy/firewall dependencies and
no slow Power Query `DimDate` generation.
