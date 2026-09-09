# Excel refresh guide for PCS and shared Absenteeism

Use these setups when one permanent workbook must refresh calculated data while
people keep team-owned action notes.

The idea is simple:

1. WFMHub updates a CSV whose name never changes.
2. Excel reads that CSV with Power Query.
3. **Refresh All** replaces report data, not the team’s action table.

## PCS: one-time connection setup

1. Install or upgrade WFMHub and run **Update latest PCS data** once.
2. Close `Reports\PCS Live Tracker.xlsx` and wait for OneDrive sync.
3. Choose **PCS Report & Coaching > Install/repair Power Query**.
4. WFMHub asks desktop Excel to connect five query tables to the fixed CSVs in
   `Feed\PCS`, refresh them, save, and close the workbook.
5. Open the tracker. `SETUP` should show `Power Query Installed = YES`.

The query destinations are `_PCS_LOB`, `_PCS_AGENT`, `_PCS_DAILY`,
`PERFORMANCE!tblPcsPerformance`, and `COACHING!tblCoachingQueue`. Never connect
Power Query to the blue `tblCoachingActions` table.

## PCS: normal refresh

1. Put new untouched FTE/Call-by-Call extracts in their normal folders.
2. In WFMHub choose **Update latest PCS data**. The tracker can remain closed or
   open because the Hub changes only external CSV feeds.
3. In Excel choose **Data > Refresh All**.
4. Use native table filters. For slicers, click inside `PERFORMANCE` or the
   coaching queue and choose **Table Design > Insert Slicer**.

Do not rebuild or replace the tracker for a normal update. Power Query is only
transport; it does not calculate PCS.

## Absenteeism: one-time connection setup

Make a backup copy before starting.

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
- **PCS shows old values:** run the Hub PCS update first, then Excel **Data >
  Refresh All**. Check `Feed\PCS\PCS_MANIFEST_CURRENT.csv` for the data-through
  and refresh timestamps.
