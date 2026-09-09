# Excel refresh guide for the shared Absenteeism file

Use this setup when one workbook must stay in Teams, SharePoint, or OneDrive
while people add absence review notes. PCS does not use this workflow: its
permanent tracker is updated by replacing one plain DATA table from the newest
Hub-generated paste workbook.

The idea is simple:

1. WFMHub updates a CSV whose name never changes.
2. Excel reads that CSV with Power Query.
3. **Refresh All** replaces report data, not the team’s action table.

Do this once for the shared Absenteeism workbook. Make a backup copy before
starting.

## Absenteeism: connect the clean ledger

1. Run **WFMHub > Refresh source data once > Attendance/absence**.
2. Confirm this file exists:
   `Feed\Absenteeism\ABSENCE_AGENT_DAY_CURRENT.csv`.
3. Open `Reports\Final Absenteeism.xlsx`.
4. Create a Blank Query in desktop Excel from the generated Absenteeism M
   script and load it as a table at `ABSENCE_DATA!$A$4`.
5. Rename the new table exactly to `tblAbsenceData`.
6. Set the `Date` column to **Date** in Power Query.
7. Save the workbook.

The `ACTIONS` sheet is the permanent review log. Do not load a query into it.
Case ID keeps each comment attached to the correct agent and day.

## Absenteeism: connect the review queue

1. Confirm `Feed\Absenteeism\ABSENCE_REVIEW_CASE_CURRENT.csv` exists.
2. Create a Blank Query from the generated review-queue M script, loading at
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
2. Create a Blank Query from the generated component-detail M script, loading at
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
2. For Absenteeism, **Refresh All** updates `TEAM_VIEW`,
   `COMPONENT_VIEW`, and the review queue. Do not rebuild the shared workbook
   unless WFMHub ships a new workbook design.

## If Excel shows an error

- **The field was not found:** the wrong CSV was selected or a table header was
  manually renamed. Select the fixed `..._CURRENT.csv` file again.
- **Table name already exists:** convert or rename the old table before naming
  the query table.
- **A filter appears blank:** clear the current table filters and try again.
- **Someone is editing:** do not replace the workbook. Refresh only the query,
  and use a personal Sheet View before applying table filters.
- **Access denied / workbook read-only:** close the Absenteeism workbook, wait
  for OneDrive sync, then refresh again. Do not reload the extracts.
