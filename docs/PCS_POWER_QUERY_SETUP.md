# PCS Power Query setup — one shared workbook

This is a **one-time setup** in desktop Excel. Afterward, the normal update is:

```text
WFMHub > PCS > Update PCS data
    → copy six current CSVs to the same SharePoint folder
    → open the same shared PCS workbook in desktop Excel
    → Data > Refresh All > wait > Save
```

WFMHub calculates PCS and writes six small, precomputed CSVs. Excel reads those
CSVs. Team Leaders and Quality edit coaching in the **same workbook**; they do
not need to run WFMHub or refresh Power Query. Do not create a new PCS workbook
on each update. Replacing the shared workbook later can erase their coaching.

## Before you touch Excel

1. Run **PCS > Update PCS data** in WFMHub. It creates or updates the six files
   in the `PCS CSV feeds` path shown by the Hub (normally `Feed/PCS` inside the
   current Hub installation). This Hub folder may change when you install a new
   portable release; it is **not** the folder Power Query should point to.
2. In your restricted WFM SharePoint library, make one folder such as
   `PCS/Source CSV`. Sync the library to your PC with OneDrive so it appears in
   File Explorer. Put the six `*_CURRENT.csv` files below in that folder.
   Keep the names unchanged. Do not connect Power Query to the versioned Hub
   `Feed/PCS` path.
3. Copy the **existing PCS Live Tracker.xlsx with the real coaching** to the
   SharePoint `PCS` folder **once**. If you already have the canonical shared
   workbook there, use that one. Do not overwrite it with a newly generated
   tracker. Make a separate backup copy before the one-time setup.
4. Open the SharePoint workbook from its **locally synced File Explorer path**
   in desktop Excel. A browser `https://` URL is not a `From Text/CSV` file
   path. Keep the backup closed and untouched while you set up the working copy.

The SharePoint folder needs these six files, all from the **same Hub update**:

| CSV filename | Excel sheet | Excel table name after setup |
| --- | --- | --- |
| `PCS_FILTER_LIST_CURRENT.csv` | `_PCS_FILTERS` | `tblPcsFilters` |
| `PCS_LOB_SCORECARD_CURRENT.csv` | `_PCS_LOB` | `tblPcsLob` |
| `PCS_AGENT_SCORECARD_CURRENT.csv` | `_PCS_AGENT` | `tblPcsAgent` |
| `PCS_DAILY_SCORECARD_CURRENT.csv` | `_PCS_DAILY` | `tblPcsDaily` |
| `PCS_RESULTS_CURRENT.csv` | `PERFORMANCE` | `tblPcsPerformance` |
| `PCS_COACHING_OPPORTUNITY_CURRENT.csv` | `_PCS_COACH` | `tblPcsCoachingView` |

Do **not** touch `COACHING!tblCoachingActions`. That is the editable, human-owned
log. `COACHING!tblCoachingQueue` is a formula view, not a CSV destination.

## Do the six queries, one at a time

The existing six sheets already contain normal Excel feed tables. A Power
Query table cannot be placed on top of one of them. Work in the **backup-backed
SharePoint copy** and repeat steps A–D for each row of the map above.

### A. Make room for one query

1. If the target sheet starts with `_PCS_`, right-click any sheet tab, choose
   **Unhide**, and select that target sheet. `PERFORMANCE` is already visible.
2. Click a cell inside its old feed table, open **Table Design**, choose
   **Convert to Range**, and confirm **Yes**. This removes only the old table
   object. It does not delete the sheet or the coaching log.
3. Click `A4` on that sheet. Press `Ctrl+Shift+End` to select the old headers and
   rows, then press `Delete` to **clear contents**. Do not choose **Delete sheet**
   or **Delete rows/columns**; the report expects headers in row 4 and data from
   row 5. Leave the title and subtitle in rows 1–2 alone.
4. Save. If anything looks wrong, close without further edits and restore the
   backup. Do not try to repair formulas by deleting named ranges.

### B. Import its CSV

1. In Excel choose **Data > Get Data > From File > From Text/CSV** (some Excel
   versions show **Data > From Text/CSV**). Pick the matching CSV from the
   **locally synced SharePoint `Source CSV` folder**, not from WFMHub's `Feed`.
2. In the preview, check **File Origin = UTF-8** and **Delimiter = Comma**.
   Click **Transform Data**, not **Load**.
3. In Power Query, check the first row is the exact column headers. If it says
   `Column1`, `Column2`, use **Home > Use First Row as Headers**.
4. In **Applied Steps**, delete the automatic **Changed Type** step if Excel
   added one. Excel must not turn an `Agent ID`, `Call ID`, or filter key into
   a number; IDs may contain leading zeros or letters.
5. Assign column data types using the type map below. All unlisted columns
   are **Text**. Do not remove or reorder columns. In the editor, select one or
   more column headers, then use the small data-type icon at the left of the
   header (or **Transform > Data Type**).
6. In the Query Settings pane, give the query a clear name, for example
   `qPCS_Filters`, `qPCS_Lob`, `qPCS_Agent`, `qPCS_Daily`, `qPCS_Results`, or
   `qPCS_Coaching`. Query names are for your convenience; column names must
   remain exact.

| CSV | Whole number columns | Decimal number columns | Date columns | Date/time column |
| --- | --- | --- | --- | --- |
| Filter list | — | — | — | — |
| LOB scorecard | `Rank`, `Valid Q1`, `Coaching Due` | `PCS`, `Prior PCS`, `Change`, `Participation` | `Data Through` | `Feed Refreshed At` |
| Agent scorecard | `Rank`, `Valid Q1`, `Coaching Due` | `PCS`, `Prior PCS`, `Change`, `Participation` | `Data Through` | `Feed Refreshed At` |
| Daily scorecard | `Rank`, `Valid Q1`, `Coaching Due` | `PCS`, `Participation` | `Date`, `Data Through` | `Feed Refreshed At` |
| Results | `Valid Q1`, `PCS Status 1`, `Q1 Nonblank`, `Score <= 3`, `Score > 3`, `Inbound Call Legs` | `PCS Average`, `Participation Rate` | `Period Start`, `Period End`, `Data Through` | `Feed Refreshed At` |
| Coaching opportunity | `Rank` | `Q1 Score` | `Date`, `Data Through` | `Feed Refreshed At` |

**Important:** All six filter-list columns are Text. `Agent ID`, `Call ID`,
`View Key`, `Coaching Key`, `Team Key`, and `Agent Key` stay Text everywhere.
If a date or decimal shows **Error** in Power Query, stop and check its type
and locale before loading. Do not replace the source CSV with manually edited
data.

### C. Load it at the exact Excel cell

1. Choose **Home > Close & Load > Close & Load To…**. Select **Table** and
   **Existing worksheet**, then choose `A4` on the target sheet in the map.
   Leave **Add this data to the Data Model** unchecked. Connection-only or
   Data Model-only loads do **not** fill this report's cards and charts.
2. If **Close & Load To…** is unavailable, choose **Close & Load** first;
   Excel may put the query on a new sheet. Then use **Data > Queries &
   Connections**, right-click that query, choose **Load To…**, and change it
   to **Table > Existing worksheet > target sheet A4**. Remove any temporary
   new sheet only after the query is correctly loaded at the target.
3. Click inside the newly loaded table. Under **Table Design > Table Name**,
   set its name to the one in the map. If Excel says the name already exists,
   the old table was not converted to a range; stop and check step A.
4. Confirm the row-4 headers and row-5 data line up under the correct columns.
   For the filter list, `Period` must be in column A and `LOB` in column B.
   A month/date showing in the LOB dropdown means this alignment is wrong.
5. Save. Repeat for the next CSV. Hide the five `_PCS_` staging sheets again
   after all six queries work. Leave `PERFORMANCE` visible.

The Hub's report formulas use the **sheet positions** `A4`/`A5`, not a Data
Model relationship. Loading a query elsewhere will leave the overview blank.

### D. Make refresh safe for the team

In **Data > Queries & Connections**, open each query's **Properties**. Keep
**Refresh this connection on Refresh All** enabled, but turn **off**
**Refresh data when opening the file**, **Refresh every n minutes**, and
**Enable background refresh**. Do not select an option that removes external
data before saving. This lets Team Leaders open the saved report without their
PC needing your local SharePoint sync path. Only WFM should press Refresh All.

## Prove it works before sharing the link

1. In desktop Excel set **Formulas > Calculation Options > Automatic**.
2. Choose **Data > Refresh All**. In **Queries & Connections**, wait until all
   six queries finish without errors. Then choose **Formulas > Calculate Now**
   if cards or charts have not updated yet.
3. Check `OVERVIEW`: Period, LOB, Team Leader, and Agent dropdowns contain the
   right kind of values; cards, LOB table, agent table and charts show data.
   Check `PERFORMANCE` has dated rows. Check `COACHING` still has its existing
   action statuses, owners, comments and Call IDs.
4. Save the **same SharePoint workbook** and wait for OneDrive's sync icon to
   finish. Then share that file's link. Keep the backup until the team confirms
   the saved report and coaching.
   In `SETUP`, mark `Feed Sync` as `PQ READY` after this validation; the field is
   informational and does not refresh automatically.
5. In WFMHub, set **PCS > Set permanent PCS workbook path** to this existing,
   locally synced SharePoint workbook. Do this **after** moving it; WFMHub
   should never make a second collaborative copy.

Desktop Excel is required for this one-time setup and WFM's refresh. The Hub
cannot run this Excel GUI step on the server, so verify it on the work PC before
calling the change live.

## The normal update, every time

1. Add the newest FTE/Call-by-Call extracts and choose **PCS > Update PCS
   data**. WFMHub updates its six local CSVs; it does not rewrite the existing
   shared workbook or touch coaching.
2. Copy **all six** `*_CURRENT.csv` files from the Hub's displayed PCS CSV
   folder to the same SharePoint `Source CSV` folder, choosing **Replace** for
   the matching filenames. Do not rename them or refresh halfway through the
   copy. Wait until OneDrive finishes syncing all six.
3. Ask coaching owners to pause edits briefly. Open the **same** shared tracker
   in desktop Excel, choose **Data > Refresh All**, wait for all six queries,
   check a date and one LOB, then save. Wait for SharePoint sync before telling
   the team the report is current. Coaching edits remain in that workbook.

Never use the old **Hub Sync PCS workbook** command after adding Power Query.
It is a legacy Excel-automation route and is deliberately blocked for a
query-enabled workbook. Do not generate and upload a fresh workbook each day.

## If something looks wrong

| Symptom | Check first |
| --- | --- |
| Cards/charts empty | Six queries loaded as tables at the mapped `A4` cells, not Data Model/connection-only; Calculation = Automatic. |
| Month appears in LOB dropdown | `_PCS_FILTERS` headers `Period`/`LOB` are in columns A/B; no leftover old table or shifted load. |
| `Agent ID` loses a leading zero | Remove Power Query's automatic `Changed Type`; set ID/key columns to Text before loading. |
| Refresh says file not found | Source step points to your **synced SharePoint CSV folder**, not an old versioned Hub folder; six files exist with exact names. |
| Team member gets a refresh error | They do not need to refresh; WFM refreshes and saves the shared workbook. Disable Refresh on Open. |
| Excel offers workbook repair | Stop; keep the backup, record the repair XML and which query was added last. Do not overwrite the working shared file. |

Microsoft references: [convert an Excel table to a range](https://support.microsoft.com/en-us/excel/convert-an-excel-table-to-a-range-of-data),
[load a query into an existing worksheet](https://support.microsoft.com/en-us/excel/create-load-or-edit-a-query-in-excel-power-query),
[query connection refresh settings](https://support.microsoft.com/en-us/excel/connection-properties),
and [manage queries](https://support.microsoft.com/en-us/excel/manage-queries-power-query).
