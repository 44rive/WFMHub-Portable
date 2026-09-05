# Excel refresh guide for shared PCS and Absenteeism files

This setup is optional. Use it when one workbook must stay in Teams, SharePoint,
or OneDrive while people add coaching or review notes.

The idea is simple:

1. WFMHub updates a CSV whose name never changes.
2. Excel reads that CSV with Power Query.
3. **Refresh All** replaces report data, not the team’s action table.

Do this once per shared workbook. Make a backup copy before starting.

## PCS: connect the clean data

1. Run **WFMHub > Refresh source data once > Agent PCS**.
2. Confirm this file exists:
   `Feed\PCS\PCS_AGENT_DAY_CURRENT.csv`.
3. Open `Reports\PCS Performance.xlsx`.
4. Open `PCS_DATA`, click inside the table, and note its name in **Table
   Design > Table Name**: `tblPcsData`.
5. Choose **Table Design > Convert to Range**, then confirm **Yes**.
6. Delete the old data area from row 4 downward on `PCS_DATA`. Do not delete
   the sheet.
7. Choose **Data > Get Data > From File > From Text/CSV**.
8. Select `PCS_AGENT_DAY_CURRENT.csv`, then choose **Transform Data**.
9. In Power Query, click the `Date` column. Choose **Data Type > Date**.
10. Choose **Home > Close & Load > Close & Load To...**.
11. Choose **Table**, then **Existing worksheet** and select
    `PCS_DATA!$A$4`.
12. Click in the new table. In **Table Design > Table Name**, rename it exactly
    to `tblPcsData`.
13. Choose **Data > Queries & Connections**, right-click the query, choose
    **Properties**, and clear **Enable background refresh**. This makes the
    dashboard wait until the full file is loaded.
14. Save the workbook.

The Dashboard now follows the newest `Date` in `tblPcsData`. `Current week`
means Monday through that newest date. LOB narrows the Team Leader list; LOB and
Team Leader narrow the Agent list. If an old selection is no longer valid, set
the downstream box back to `All`.

## PCS: connect new coaching opportunities

1. Confirm this file exists:
   `Feed\PCS\PCS_COACHING_OPPORTUNITY_CURRENT.csv`.
2. Repeat the Power Query steps above on `COACHING_QUEUE`, loading at
   `COACHING_QUEUE!$A$4`.
3. Rename the query table exactly to `tblCoachingQueue`.
4. Add one table column at the right named `Action Status`.
5. In its first data row, enter:

   ```excel
   =IFERROR(XLOOKUP([@[Coaching Key]],tblCoaching[Coaching Key],tblCoaching[Coaching Status]),"Not started")
   ```

   Excel fills the formula down the column.
6. For a new case, copy columns A:M from `COACHING_QUEUE` into the next empty
   row of `COACHING`. Fill only the five blue columns there.

`COACHING` is the team’s permanent action log. Power Query must never load into
that sheet. Agent ID identifies the employee; Coaching Key identifies the exact
survey call.

## Absenteeism: connect the clean ledger

1. Run **WFMHub > Refresh source data once > Attendance/absence**.
2. Confirm this file exists:
   `Feed\Absenteeism\ABSENCE_AGENT_DAY_CURRENT.csv`.
3. Open `Reports\Final Absenteeism.xlsx`.
4. Repeat the PCS clean-data steps on `ABSENCE_DATA`, loading at
   `ABSENCE_DATA!$A$4`.
5. Rename the new table exactly to `tblAbsenceData`.
6. Set the `Date` column to **Date** in Power Query.
7. Save the workbook.

The `ACTIONS` sheet is the permanent review log. Do not load a query into it.
Case ID keeps each comment attached to the correct agent and day.

## Absenteeism: connect the review queue

1. Confirm `Feed\Absenteeism\ABSENCE_REVIEW_CASE_CURRENT.csv` exists.
2. Repeat the Power Query steps on `ACTION_QUEUE`, loading at
   `ACTION_QUEUE!$A$4`.
3. Rename the table exactly to `tblActionQueue`.
4. Add one table column at the right named `Action Status`.
5. In its first data row, enter:

   ```excel
   =IFERROR(XLOOKUP([@[Case ID]],tblActions[Case ID],tblActions[Review Status]),"Not started")
   ```

   Excel fills the formula down.
6. Copy a new case into `ACTIONS` only when someone must own and comment on it.

## Absenteeism: connect exact activity detail

1. Confirm `Feed\Absenteeism\ABSENCE_COMPONENT_CURRENT.csv` exists.
2. Repeat the Power Query steps on `ACTIVITY_DETAIL`, loading at
   `ACTIVITY_DETAIL!$A$4`.
3. In Power Query set `Date` to **Date** and `Start`/`End` to **Date/Time**.
4. Rename the table exactly to `tblActivityDetail`.

`TEAM_VIEW` now follows the newest ledger date and shows filtered results and
review cases. `COMPONENT_VIEW` follows the same selectors and explains absence
and shrinkage by category. The `ACTIONS` sheet is the permanent review log;
Power Query must never load into it. Case ID keeps each comment attached to the
correct agent and day.

## Normal refresh after setup

1. Put new untouched exports in the normal source folders.
2. Run **Refresh source data once** in WFMHub.
3. Wait for **Refresh complete**.
4. For PCS, open the shared workbook and choose **Data > Refresh All**.
5. Wait until **Queries & Connections** shows no query still refreshing, then
   save.
6. For Absenteeism, the same **Refresh All** updates `TEAM_VIEW`,
   `COMPONENT_VIEW`, and the review queue. Do not rebuild the shared workbook
   unless WFMHub ships a new workbook design.

## If Excel shows an error

- **The field was not found:** the wrong CSV was selected or a table header was
  manually renamed. Select the fixed `..._CURRENT.csv` file again.
- **Table name already exists:** convert or rename the old table before naming
  the query table.
- **The selector is blank:** set LOB, Team Leader, and Agent back to `All`, then
  choose them again from left to right.
- **Someone is editing:** do not replace the workbook. Refresh only the query,
  and use a personal Sheet View before applying table filters.
- **A new agent is missing:** confirm the correct fixed CSV was refreshed, then
  set LOB, Team Leader, and Agent back to `All`. `TEAM_VIEW` reads new agents
  directly from the refreshed table; no report rebuild is required.
