# Excel refresh guide for shared PCS and Absenteeism files

Use this setup when one workbook must stay in Teams, SharePoint, or OneDrive
while people add coaching or review notes.

The idea is simple:

1. WFMHub updates a CSV whose name never changes.
2. Excel reads that CSV with Power Query.
3. **Refresh All** replaces report data, not the team’s action table.

Do this once per shared workbook. Make a backup copy before starting.

## PCS: decide who refreshes

The simplest production method is one named owner:

1. Keep `PCS Operational Tracker.xlsx` in SharePoint or Teams for everybody.
2. The owner opens it in desktop Excel on the WFM work machine.
3. The owner uses the `LOCAL` scripts, clicks **Refresh All**, saves, and closes.
4. Team leaders and Quality open the same shared file and update `COACHING`.

They do not need WFMHub, Python, a database driver, or the source extracts.

Use `SHAREPOINT` mode only when the two fixed CSV feeds are also copied or
synced into one SharePoint folder after every Hub refresh. In that case, fill in
the SharePoint URL and unique folder fragment on `SETUP`.

## PCS: connect the clean data once

1. Run **WFMHub > Refresh source data once > Agent PCS**.
2. Open `Reports\PCS Operational Tracker.xlsx` in desktop Excel.
3. Open `SETUP`. Leave **Connection Mode** as `LOCAL` for the simple owner
   workflow. Confirm **Local Feed Folder** points to `Feed\PCS`.
4. Open the file shown in **Local Data Script** using Notepad. Select all and
   copy it.
5. In Excel choose **Data > Get Data > From Other Sources > Blank Query**.
6. In Power Query choose **Home > Advanced Editor**. Delete everything, paste
   the script, and choose **Done**.
7. Rename the query exactly `PCS_DATA`.
8. Back in Excel open `PCS_DATA`, click inside the starter table, choose **Table
   Design > Convert to Range**, and confirm **Yes**.
9. Clear the old area from `A4` through `Z` downward. Keep the sheet itself.
10. In **Queries & Connections**, right-click `PCS_DATA`, choose **Load To...**,
    select **Table** and **Existing worksheet**, then choose `PCS_DATA!$A$4`.
11. Click inside the new query table. Under **Table Design > Table Name** rename
    it exactly `tblPcsData`.
12. Right-click the query, choose **Properties**, and clear **Enable background
    refresh**. Do not select **Add this data to the Data Model**.

The Dashboard now follows the newest `Date` in `tblPcsData`. `Current week`
means Monday through that newest date. LOB, Team Leader and Agent are combined
as filters. If a combination has no data, set one or more boxes back to `All`.

## PCS: connect new coaching opportunities once

1. Open the file shown in `SETUP` under **Local Coaching Script**.
2. Repeat the Blank Query and Advanced Editor steps above.
3. Rename this query exactly `COACHING_QUEUE`.
4. On the worksheet `COACHING_QUEUE`, convert the starter table to a range and
   clear `A4:M` downward.
5. Load the query as a Table to existing cell `COACHING_QUEUE!$A$4`.
6. Rename the new query table exactly `tblCoachingQueue`.
7. Clear **Enable background refresh** and do not add it to the Data Model.
8. Click **Refresh All**. If both queries succeed, set **Power Query Installed**
   to `YES` on `SETUP`, save, and close the workbook.
9. For a new case, copy columns A:M from `COACHING_QUEUE` into the next empty
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
- **A new agent is missing from PCS TEAM_VIEW:** reset LOB, Team Leader, and
  Agent to `All`, then choose **Data > Refresh All**. Do not rebuild the tracker.
- **Formula shows `_xlfn` or `#NAME?`:** open the file in current Microsoft 365
  desktop Excel; the selector views use LET, FILTER, MAP, and dynamic arrays.
