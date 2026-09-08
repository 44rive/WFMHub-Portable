# Excel refresh guide for shared PCS and Absenteeism files

Use this setup when one workbook must stay in Teams, SharePoint, or OneDrive
while people add coaching or review notes.

The idea is simple:

1. WFMHub updates a CSV whose name never changes.
2. Excel reads that CSV with Power Query.
3. **Refresh All** replaces report data, not the team’s action table.

Do this once per shared workbook. Make a backup copy before starting.

## PCS: the simple owner workflow

Use one named owner on the WFM work machine:

1. Close `PCS Operational Tracker.xlsx` if it is open.
2. In WFMHub choose **PCS Operational Tracker**.
3. Choose **Update PCS now**.
4. Wait while WFMHub loads FTE and Call by Call, publishes the fixed CSV feeds,
   installs or refreshes four Power Queries through desktop Excel, and saves.
5. The same permanent tracker opens when the update succeeds.

The first run automatically replaces four starter tables with refreshable query
tables: `tblPcsLob` on `OVERVIEW`, `tblResults` on `RESULTS`, `tblPcsData` on
`PCS_DATA`, and `tblCoachingQueue` on `COACHING_QUEUE`. It does not add any table
to the Data Model. Later runs refresh the existing connections without rebuilding
the workbook. If a future WFMHub release changes the template, close Excel and
choose **Repair/rebuild tracker and connection**. The Hub archives the old
workbook and carries its readable coaching action ledger into the new design
before reinstalling the queries.

Keep the permanent workbook in a locally synced SharePoint/Teams folder if the
team must collaborate in it. Team leaders and Quality do not need WFMHub,
Python, SQLite, or the source extracts. Only the owner updates the feeds.

Use **Show tracker status** to compare the latest feed timestamp with the feed
currently loaded in Excel. Use **Repair/rebuild tracker and connection** if the
queries are missing or the design is old. `LOCAL` mode is the supported simple workflow. The
SharePoint query scripts remain available for an advanced setup where the CSV
feeds themselves are also synchronized to SharePoint.

Use `OVERVIEW` for management totals and the per-LOB comparison. In `RESULTS`,
filter `Period View`, then `Scope Level`, LOB, Team Leader, or Agent. For a
coaching action, filter `COACHING_QUEUE`, copy its Coaching Key into the first
blank blue cell in `COACHING`, and fill only the blue action fields. The agent
and call identity fills with classic `INDEX/MATCH`. Power Query never loads into
`COACHING`.

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
2. For PCS, use **PCS Operational Tracker > Update PCS now**. It refreshes the
   targeted PCS data, all four Excel queries, and the permanent workbook in one
   action. It does not rebuild unrelated WFM models.
3. Alternatively, after the owner updates the fixed feeds, open the PCS workbook
   and use **Data > Refresh All**, wait for all queries, then save.
4. For Absenteeism, **Refresh All** updates `TEAM_VIEW`,
   `COMPONENT_VIEW`, and the review queue. Do not rebuild the shared workbook
   unless WFMHub ships a new workbook design.

## If Excel shows an error

- **The field was not found:** the wrong CSV was selected or a table header was
  manually renamed. Select the fixed `..._CURRENT.csv` file again.
- **Table name already exists:** convert or rename the old table before naming
  the query table.
- **A filter appears blank:** clear the preceding table filters, then apply
  Period View, Scope Level, LOB, Team Leader, and Agent in that order.
- **Someone is editing:** do not replace the workbook. Refresh only the query,
  and use a personal Sheet View before applying table filters.
- **Access denied / workbook read-only:** the database and fixed feeds may
  already be current. If WFMHub says they succeeded, close Excel, wait for
  OneDrive sync, and choose **Refresh Excel only**. Do not reload the extracts.
- **A new agent is missing from PCS RESULTS:** clear the filters and choose
  **Data > Refresh All**. Do not rebuild the tracker.
- **Excel reports repaired records:** keep the recovered copy closed, use the
  Hub's **Repair/rebuild tracker and connection** action, and retain the log.
  The current template contains no dynamic-array report formulas.
