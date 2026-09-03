# Excel PCS master — beginner setup

The normal PCS workbook is already ready to send. This guide is only for the
optional Excel master where you want your own PivotTables and slicers.

## What updates what

Think of it as four boxes:

1. You add untouched Call by Call and FTE extracts.
2. WFMHub reads them into SQLite and recalculates PCS.
3. Building **PCS Performance** replaces four small files in
   `output\template_feeds\pcs\current` and keeps a dated copy in `archive`.
4. Excel **Refresh All** rereads those four current files into its Data Model.

Adding an extract alone does not change Excel. First run WFMHub Refresh or
Build Reports, then open the master and choose **Data > Refresh All**.

Excel does not connect directly to SQLite. That route needs an ODBC driver,
admin installation and a trusted native DLL, which is exactly the type of file
the work-machine application policy already blocked. The CSV feed is faster to
support, easy to inspect, and still keeps data out of worksheets.

## Files Excel reads

| File | Grain | Use |
|---|---|---|
| `PCS_AgentDay.csv` | one agent/day | main PivotTables, daily and monthly score, slicers |
| `PCS_Summary.csv` | one named comparison period | latest day, selected period, current MTD and prior month |
| `PCS_Actions.csv` | one valid Q1 <=3 response | coaching queue and Actions Rate |
| `PCS_Trend.csv` | one date | simple daily trend chart |

The filenames never change. The current files are replaced atomically, so an
Excel query never needs to search dated report folders. `manifest.json` shows
the selected dates, refresh time, row counts, column types and file hashes.

## Open the protected master

Open `templates\reports\pcs.xlsx`. The portable ZIP includes this data-free
starter, so the folder is not empty. WFMHub will not overwrite it during normal
refreshes. If the file was deleted, use **Create an Excel Pivot/slicer master**
in `WFMHub.cmd` to recreate a starter from your current PCS report.

## Add the main PCS query

Do these clicks exactly once:

1. In Excel choose **Data > Get Data > From Other Sources > Blank Query**.
2. In Power Query choose **Home > Advanced Editor**.
3. Delete the text already there.
4. Open `templates\power_query\PCS_AgentDay.pq` in Notepad.
5. Copy everything, paste it into Advanced Editor, then choose **Done**.
6. Rename the query `PCS Agent Day`.
7. Choose **Home > Close & Load > Close & Load To...**.
8. Select **Only Create Connection**.
9. Tick **Add this data to the Data Model**.
10. Choose **OK**.

No data is loaded to a worksheet. Excel stores a compressed copy inside the
Data Model.

Repeat those steps only for what you need:

- `PCS_Summary.pq` → query name `PCS Summary`
- `PCS_Actions.pq` → query name `PCS Actions`
- `PCS_Trend.pq` → query name `PCS Trend`

Each query reads its CSV directly. There are no query-to-query references, so
the Power Query privacy/firewall error cannot be created by this setup.

## Create the PCS PivotTable

1. Choose **Insert > PivotTable > From Data Model**.
2. Put it on a new sheet named `PIVOT_VIEW`.
3. Drag these fields:
   - Rows: `agent_name`
   - Optional rows above Agent: `lob`, then `team_leader`
   - Values: `q1_score_sum`, `valid_q1`, `q1_nonblank`, `pcs_status_1`,
     `score_le_3`, `score_gt_3`
4. Choose **PivotTable Analyze > Insert Slicer** and select `lob`,
   `team_leader`, `agent_name`.
5. Choose **Insert Timeline** and select `business_date`.

Do not create relationships yet. The main PivotTable and all its slicers come
from `PCS Agent Day`, so none are required. Separate Trend, Summary or Actions
PivotTables can use their own filters. This avoids a slow DimDate and ambiguous
relationships while you are learning.

## Create the measures

In Excel choose **Power Pivot > Measures > New Measure**. Set **Table name** to
`PCS Agent Day` and create these one by one:

```text
PCS Average := DIVIDE(SUM('PCS Agent Day'[q1_score_sum]), SUM('PCS Agent Day'[valid_q1]))

PCS Participation := DIVIDE(SUM('PCS Agent Day'[q1_nonblank]), SUM('PCS Agent Day'[pcs_status_1]))

Low Score % := DIVIDE(SUM('PCS Agent Day'[score_le_3]), SUM('PCS Agent Day'[valid_q1]))

Positive % := DIVIDE(SUM('PCS Agent Day'[score_gt_3]), SUM('PCS Agent Day'[valid_q1]))
```

Format PCS Average as `0.00`; format the other three as percentages. Put these
measures in Values. Do not average the `pcs_average` column and do not display
`q1_score_sum` as if it were the score. The correct score is score sum divided
by valid-response count.

For coaching, set **Table name** to `PCS Actions` and create:

```text
Coaching Opportunities := COUNTROWS('PCS Actions')

Coaching Completed := CALCULATE(COUNTROWS('PCS Actions'), 'PCS Actions'[coaching_status] = "Completed")

Actions Rate := DIVIDE([Coaching Completed], [Coaching Opportunities])
```

Format Actions Rate as a percentage.

## How coaching works

The original PCS workbook used this idea:

`Actions Rate = completed briefings / valid Q1 responses <= 3`

WFMHub makes it auditable. Every valid inbound Q1 at or below the configured
threshold creates one row on the generated PCS workbook's `ACTIONS` sheet.

1. Open the newest generated workbook under `output\pcs`.
2. In `ACTIONS`, edit only the blue columns: Coaching Status, Coaching Date,
   Coach and Coaching Comment.
3. Use `Completed`, `Pending`, or `Not required`. Old values `OK`, `Yes`, `Y`
   and `Done` are accepted and normalized to Completed.
4. Save the workbook.
5. In WFMHub choose **Import PCS coaching decisions** and paste its path.
6. Build PCS Performance again.
7. Open the Excel master and choose **Refresh All**.

SQLite remembers the decisions. A later report does not erase them.

## Daily use after setup

1. Add new untouched extracts.
2. Run WFMHub Refresh or Build Reports for PCS.
3. Check the generated static workbook before sending anything.
4. Open `templates\reports\pcs.xlsx`.
5. Choose **Data > Refresh All**.
6. Wait until **Queries & Connections** stops refreshing.
7. Check the newest date in the timeline or Trend PivotTable.
8. Choose **File > Save As** and save the management copy you will email.

The “every three hours” requirement is only the send/refresh rhythm. It never
changes the daily and monthly formulas.

## If something breaks

- **SUM gives an error:** replace the query with the supplied `.pq` file. It
  explicitly types counters as numbers.
- **File not found:** unhide `_AUDIT` and check `Template current feed`, or open
  `output\template_feeds\pcs\current\manifest.json`.
- **No new dates:** WFMHub must rebuild PCS before Excel Refresh All.
- **Refresh is slow:** load only `PCS Agent Day`; add the other three queries
  only when needed. Do not load any query to a worksheet.
- **A slicer does not control another PivotTable:** both PivotTables must use
  the same Data Model table, or you must connect that slicer under **Report
  Connections**. Do not accept automatic relationships just because Excel
  proposes them.
