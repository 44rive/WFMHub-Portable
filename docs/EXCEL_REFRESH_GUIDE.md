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
4. WFMHub asks desktop Excel to connect six query tables to the fixed CSVs in
   `Feed\PCS`, refresh them, save, and close the workbook.
5. Open the tracker. `SETUP` should show `Power Query Installed = YES`.

For tracker contract `2026.11.3` and later, Install/repair first rebuilds an
outdated or damaged tracker, archives its prior copy, and carries keyed coaching
actions forward. Dashboard formulas read stable sheet-backed workbook names, so
replacing the six starter tables cannot turn their references into `#REF!`.
The installer validates the presentation before saving and closes without
saving if any governed formula or workbook name is broken.

The query destinations are `_PCS_FILTERS`, `_PCS_LOB`, `_PCS_AGENT`,
`_PCS_DAILY`, `_PCS_COACH`, and `PERFORMANCE!tblPcsPerformance`. The visible
coaching queue is a lightweight lookup view over `_PCS_COACH`. Never connect
Power Query to the visible `tblCoachingQueue` or blue `tblCoachingActions`.

## PCS: normal refresh

1. Put new untouched FTE/Call-by-Call extracts in their normal folders.
2. In WFMHub choose **Update latest PCS data**. The tracker can remain closed or
   open because the Hub changes only external CSV feeds.
3. In Excel choose **Data > Refresh All**.
4. Use the four `OVERVIEW` dropdowns from left to right. Reset Team Leader and
   Agent to `All` after changing LOB; reset Agent after changing Team Leader.
5. Use Period and LOB on `COACHING`. Use native filters or optional slicers only
   on `PERFORMANCE`.

Do not rebuild or replace the tracker for a normal update. Power Query is only
transport; it does not calculate PCS.

## Final Absenteeism

1. Put the latest final Verint Activities export in its normal source folder.
2. Run **WFMHub > Refresh source data once > Attendance/absence**.
3. Build **Final Absenteeism / Shrinkage**.
4. Open `Reports\Final Absenteeism.xlsx`.
5. Use `TEAM_VIEW` for filtered agent results, `COMPONENT_VIEW` for category
   totals and `ACTIVITY_DETAIL` for exact source intervals.
6. Use `ACTION_QUEUE` as a read-only list of incomplete or inconsistent coding.
   Correct the source in Verint, export Activities again and refresh.

The generated workbook does not accept attendance decisions and does not write
anything back to WFMHub.

## Normal refresh after setup

1. Put new untouched exports in the normal source folders.
2. Rebuild the focused workbook you need. The Hub archives the prior copy before
   replacing it.

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
