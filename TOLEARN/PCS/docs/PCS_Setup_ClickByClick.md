# PCS Daily Report — Setup, click by click

Two parts:

- **Part A** — one-off build, roughly 45 minutes. You do this once.
- **Part B** — the daily/3-hourly run. Two clicks.

Excel 365 on Windows assumed. Where the ribbon differs by version it is noted.

---

# PART A — one-off build

## Step 1 — Point the refresh job at your hub

1. Open `PCS_Refresh_3h.cmd` in Notepad (right-click → **Edit**, not double-click).
2. Find this line near the top:

   ```
   set "WFMHUB_HOME=%USERPROFILE%\WFMHub-Portable"
   ```

3. If your hub lives elsewhere, change the path to the folder that contains **`WFMHub.cmd`**
   and the **`runtime`** folder. No trailing backslash.
4. Save and close.

## Step 2 — First run

1. Double-click **`PCS_Refresh_3h.cmd`**.
2. A window opens showing three steps. First run takes longer — it ingests the whole window.
3. It should finish with:

   ```
   DONE  2026-08-28_1500
   Export : ...\data\current\pcs_agent_day.csv
   ```

**If it fails**, the log is in `logs\`. Most common causes:

| Message | Cause |
|---|---|
| `embedded Python was not found` | `WFMHUB_HOME` in step 1 is wrong |
| `refresh returned an error` | Source root not configured, or Call-by-Call extract missing. Open `WFMHub.cmd` → `[6] Show source health` |
| Export succeeds, file has only headers | No call data in the window. Check `[6] Show source health` for the `calls` family |

## Step 3 — New workbook

1. Open Excel → **Blank workbook**.
2. **File → Save As** → `WFM-PCS-Daily\PCS_Daily_Report.xlsx`.
   Save it **outside** `data\current\` — that folder must contain only the export.

## Step 4 — Connect Power Query to the folder

1. **Data** tab → **Get Data** → **From File** → **From Folder**.
2. Browse to `WFM-PCS-Daily\data\current` → **Open**.
3. A preview lists the files. Click the **Combine** dropdown (bottom right) →
   **Combine & Transform Data**.

   ⚠️ **Not** plain "Combine & Load". You need the editor.

4. In the sample dialog: **Delimiter** = Comma, **File origin** = `65001: Unicode (UTF-8)`.
   Click **OK**.

The Power Query Editor opens.

## Step 5 — Shape the query

Do these in order. Each is one click unless stated.

### 5a — Rename it
Right-hand **Query Settings** pane → **Name** → type `qryPcsAgentDay` → Enter.

### 5b — Keep only CSVs
1. Click the **Source** step in Applied Steps.
2. On the `Extension` column dropdown → **Text Filters → Equals** → type `.csv` → **OK**.
3. Click the **last** step in Applied Steps again to return to the combined view.

This keeps the `.manifest.txt` and `last_refresh.txt` files out.

### 5c — Sort newest first
On `Date modified` → dropdown → **Sort Descending**.

*(If `Date modified` is not visible, it was dropped during combine. Skip to 5d — step 5e still
protects you.)*

### 5d — Set the types
1. Select `business_date` → **Transform** tab → **Data Type** → **Date**.
2. Select `pcs_enabled_calls`, `survey_responses`, `pcs_score_count`, `q1_response_count`,
   `q2_response_count`, `top_box_responses`, `low_score_responses`, `call_legs`,
   `handled_calls`, `inbound_calls`, `outbound_calls`, `comments_count`,
   `talk_seconds`, `hold_seconds`, `wrap_seconds`, `handle_seconds`
   (Ctrl-click each header) → **Data Type** → **Whole Number**.
3. Select `pcs_score_sum`, `q1_score_sum`, `q2_score_sum` → **Data Type** → **Decimal Number**.

### 5e — ⚠️ Remove the ten ratio columns

Ctrl-click all ten of these headers:

```
average_talk_seconds     average_hold_seconds    average_wrap_seconds
average_handle_seconds   response_rate           q1_average
q2_average               pcs_average             top_box_percent
low_score_percent
```

Right-click any selected header → **Remove Columns**.

**Why:** these are already-divided per-agent-day ratios. Dropped into a pivot they produce an
average of averages, which on a realistic team shape overstates a problem by more than a full
point. Removing them makes the mistake impossible rather than merely discouraged. Every one is
rebuilt correctly as a measure in step 7.

### 5f — ⚠️ Remove duplicate agent-days

1. Click the `agent_day_key` column header to select it.
2. **Home** tab → **Remove Rows** → **Remove Duplicates**.

**Why this matters even though the folder normally holds one file:** if a second CSV ever lands
in `data\current\`, the folder query combines both and **every count doubles silently**. Response
rate would still look plausible (both halves double), so nothing on the report face would look
wrong. Sorted newest-first, this step keeps one row per agent-day. It is the same newest-wins rule
the hub applies to its own call legs.

### 5g — Add the time columns
**Add Column** tab → **Custom Column** four times:

| New column name | Formula |
|---|---|
| `ISO Week` | `Date.WeekOfYear([business_date], Day.Monday)` |
| `Month` | `Date.ToText([business_date], "yyyy-MM")` |
| `Day Name` | `Date.DayOfWeekName([business_date])` |
| `Date Sort` | `[business_date]` |

### 5h — Add the period flags
**Add Column → Custom Column**, four more:

| Name | Formula |
|---|---|
| `Is Today` | `[business_date] = Date.From(DateTime.LocalNow())` |
| `Is D-1` | `[business_date] = Date.AddDays(Date.From(DateTime.LocalNow()), -1)` |
| `In Rolling 7` | `[business_date] >= Date.AddDays(Date.From(DateTime.LocalNow()), -7) and [business_date] < Date.From(DateTime.LocalNow())` |
| `In MTD` | `[business_date] >= Date.StartOfMonth(Date.From(DateTime.LocalNow()))` |

Set all four to **Data Type → True/False**.

These recompute on every refresh, so the report rolls forward with no maintenance.
**Rolling 7 excludes today on purpose** — today is partial and would drag the average down.

### 5i — Load to the model, not to a sheet
1. **Home** tab → **Close & Load** dropdown → **Close & Load To…**
2. Choose **Only Create Connection**.
3. ✅ Tick **Add this data to the Data Model**.
4. **OK**.

**Why not to a worksheet:** it keeps agent-level rows out of the workbook grid entirely, and
sidesteps the million-row sheet limit. Nothing is lost — the pivots read the model directly.

## Step 6 — Open Power Pivot

**Data** tab → **Manage Data Model**.

*If you cannot see it:* **File → Options → Add-ins**, set **Manage** to **COM Add-ins** → **Go**,
tick **Microsoft Power Pivot for Excel** → **OK**.

*If Power Pivot is not available in your Office licence at all*, stop here and use
`build_pcs_workbook.py` instead — see Part C.

## Step 7 — Create the measures

In the Power Pivot window, the grid at the bottom is the **calculation area**. Click any empty
cell there, type the measure, press Enter. One cell per measure.

Paste these one at a time:

```dax
Call Legs         := SUM ( pcs_agent_day[call_legs] )
Handled Calls     := SUM ( pcs_agent_day[handled_calls] )
Agents            := DISTINCTCOUNT ( pcs_agent_day[agent_id] )

PCS Enabled Calls := SUM ( pcs_agent_day[pcs_enabled_calls] )
Survey Responses  := SUM ( pcs_agent_day[survey_responses] )
Response Rate     := DIVIDE ( [Survey Responses], [PCS Enabled Calls] )

PCS Average       := DIVIDE ( SUM ( pcs_agent_day[pcs_score_sum] ), SUM ( pcs_agent_day[pcs_score_count] ) )
Q1 Average        := DIVIDE ( SUM ( pcs_agent_day[q1_score_sum] ), SUM ( pcs_agent_day[q1_response_count] ) )
Q2 Average        := DIVIDE ( SUM ( pcs_agent_day[q2_score_sum] ), SUM ( pcs_agent_day[q2_response_count] ) )

Top Box %         := DIVIDE ( SUM ( pcs_agent_day[top_box_responses] ), [Survey Responses] )
Low Score %       := DIVIDE ( SUM ( pcs_agent_day[low_score_responses] ), [Survey Responses] )
Comments          := SUM ( pcs_agent_day[comments_count] )

Handle Seconds    := SUM ( pcs_agent_day[handle_seconds] )
AHT Seconds       := DIVIDE ( [Handle Seconds], [Handled Calls] )

Min Responses     := 5
Ranked PCS Average := IF ( [Survey Responses] >= [Min Responses], [PCS Average] )
Data Sufficiency  := IF ( ISBLANK ( [Survey Responses] ) || [Survey Responses] = 0, "No responses",
                        IF ( [Survey Responses] < [Min Responses], "Insufficient", "Rankable" ) )
```

⚠️ **The table name.** If your table is not called `pcs_agent_day` in the model, use whatever
name shows on the tab at the bottom of the Power Pivot window. Easiest: type `SUM(` and let
IntelliSense offer the real name.

### Format each measure
Click a measure cell → in the **Formatting** area of the ribbon:

| Measures | Format |
|---|---|
| PCS / Q1 / Q2 / Ranked PCS Average | **Decimal Number**, 2 decimals |
| Response Rate, Top Box %, Low Score % | **Percentage**, 1 decimal |
| AHT Seconds | **Decimal Number**, 0 decimals |
| All counts | **Whole Number**, ✅ use thousands separator |

Setting it here means the format follows the measure into every pivot. Do it once.

## Step 8 — Build the pivots

For each of the five sheets: **Insert → PivotTable → From Data Model** →
**Existing Worksheet** → pick the cell → **OK**.

Drag from the field list. Measures appear at the top of the list with an **fx** icon.

| Sheet | Rows | Values | Slicers |
|---|---|---|---|
| **ROLLUP** | `ops_manager`, `lob` | PCS Average, Response Rate, Top Box %, Low Score %, Survey Responses, PCS Enabled Calls, Handled Calls, AHT Seconds | `language`, `market` |
| **TL_SCORECARD** | `team_leader`, `agent_name` | Ranked PCS Average, **Survey Responses**, Response Rate, Top Box %, Low Score %, PCS Enabled Calls, AHT Seconds, Comments | `ops_manager`, `lob`, Data Sufficiency |
| **BY_AGENT** | `agent_name`, `team_leader`, `lob` | Ranked PCS Average, Survey Responses, Top Box %, Low Score %, AHT Seconds | Data Sufficiency |
| **DAILY_TREND** | `business_date` | PCS Average, Response Rate, Survey Responses, PCS Enabled Calls | `ops_manager`, `team_leader`, `lob` |
| **EXCEPTIONS** | `agent_name`, `team_leader` | PCS Enabled Calls, Survey Responses, Low Score %, Response Rate | — |

For every pivot: **Design → Report Layout → Show in Tabular Form**, and
**Design → Report Layout → Do Not Repeat Item Labels**.

⚠️ **On TL_SCORECARD, keep `Survey Responses` immediately beside the average.** An average
without its response count invites bad coaching decisions.

### Period slicer
1. Click any pivot → **PivotTable Analyze → Insert Slicer**.
2. Tick `In MTD`, `In Rolling 7`, `Is D-1`, `Is Today`.
3. Select all four slicers → **Slicer → Report Connections** → tick **every** pivot.

That gives one period control driving all five sheets.

### Chart on DAILY_TREND
1. Click the DAILY_TREND pivot → **PivotTable Analyze → PivotChart** → **Clustered Column** → OK.
2. Right-click the *PCS Average* series → **Change Series Chart Type**.
3. Set *PCS Average* to **Line** and tick its **Secondary Axis**. Leave *Survey Responses* as
   columns. **OK**.
4. Right-click the secondary axis → **Format Axis** → Minimum `1`, Maximum `5`.

Volume behind the line, so a swing on three responses reads as a swing on three responses.

## Step 9 — Conditional formatting

On each measure column: **Home → Conditional Formatting → New Rule**.

| Measure | Rule |
|---|---|
| PCS Average | 3-Color Scale. **Type = Number** (not Percentile): Min `1` red, Mid `3.5` amber, Max `5` green |
| Low Score % | Highlight Cells → Greater Than `0.20` → red fill |
| Survey Responses | Data Bars, solid, light blue |

⚠️ **Type = Number, never Percentile.** Percentile rescales on every refresh, so last week's
colours stop meaning the same thing as this week's.

⚠️ **Do not format blank cells.** No fill, no "0". Blank must read as *absent*.

When the rule dialog offers it, choose **All cells showing "<measure>" values for …** so the
formatting survives refresh and layout changes.

## Step 10 — Freshness stamp

1. **Data → Get Data → From File → From Text/CSV** →
   select `data\current\last_refresh.txt` → **Load To… → Only Create Connection**.
2. Or, simpler: put this in a cell on SUMMARY and repeat it top-right on each sheet:

   ```excel
   ="Data as at "&TEXT(NOW(),"yyyy-mm-dd hh:mm")&"  (refreshed on open)"
   ```

   The `last_refresh.txt` file carries the authoritative stamp; the formula above tells you when
   *the workbook* last refreshed, which is what the reader needs.

3. Add this line beneath it, in red:

   > **Today is partial. Responses arrive after the call ends.**

## Step 11 — Refresh on open

**Data → Queries & Connections** → right-click `qryPcsAgentDay` → **Properties** →
✅ **Refresh data when opening the file** → **OK**.

## Step 12 — ⚠️ Reconcile before the first send

Do not skip this. Run the hub's own PCS workbook for the same period
(`WFMHub.cmd` → `[2] Build reports` → `quality_pcs`) and compare its **SUMMARY** sheet
against yours with the period slicer set to **In MTD**:

| # | Check | Expected |
|---|---|---|
| 1 | PCS-enabled calls | identical |
| 2 | Survey responses | identical |
| 3 | PCS average | identical |
| 4 | Response rate | identical |
| 5 | Model row count vs export row count | identical — proves no duplicate-file doubling |
| 6 | An agent with zero responses | **blank**, not `0.00` |

Same source, same grain — they must tie exactly. If they do not, the cause is almost always a
second file in `data\current\`.

---

# PART B — the routine

## Automatic, every 3 hours

1. Press **Win+R**, type `taskschd.msc`, Enter.
2. **Create Basic Task** → Name: `WFM PCS refresh` → **Next**.
3. Trigger: **Daily** → **Next** → start time `07:00` → **Next**.
4. Action: **Start a program** → **Browse** → select `PCS_Refresh_3h.cmd`.
5. **Add arguments:** `/quiet`
6. **Next** → ✅ **Open the Properties dialog** → **Finish**.
7. **Triggers** tab → double-click the trigger → ✅ **Repeat task every** → type `3 hours` →
   *for a duration of* **1 day** → **OK**.
8. **Settings** tab → ✅ **Run task as soon as possible after a scheduled start is missed**.
9. **OK**.

The `/quiet` argument stops it pausing for a keypress when nobody is watching.

## Manual

Double-click `PCS_Refresh_3h.cmd`, wait for **DONE**.

## Then, to send the report

1. Open `PCS_Daily_Report.xlsx`.
2. **Data → Refresh All** (or just open it, if you did step 11).
3. Check the freshness stamp.
4. Send.

---

# PART C — no Power Pivot?

If Power Pivot is not in your Office licence, use the builder script. It produces a complete
static workbook with the same measures and the same guarantees, computed in Python.

```
cd %USERPROFILE%\WFM-PCS-Daily
"%USERPROFILE%\WFMHub-Portable\runtime\python.exe" build_pcs_workbook.py data\current\pcs_agent_day.csv
```

Output lands in `output\PCS_Daily_<date>_<time>.xlsx` with SUMMARY, ROLLUP, TL_SCORECARD,
BY_AGENT, DAILY_TREND, EXCEPTIONS and DEFINITIONS sheets.

Change the ranking threshold with `--min-responses 3`.

You can append this as a fourth step in `PCS_Refresh_3h.cmd` to have the workbook rebuilt on
every 3-hourly run.

---

# Troubleshooting

| Symptom | Cause | Fix |
|---|---|---|
| Every count is exactly double | Two CSVs in `data\current\` | Move extras to `data\archive\`. Step 5f should catch it; confirm the step is present |
| PCS average looks wrong at TL level | A ratio column got into Values | Check step 5e removed all ten. Use the measures |
| Agent shows `0.00` with no responses | Formatting is filling blanks | Remove that rule. `DIVIDE` already returns blank |
| Report only covers a few days | Refresh window too narrow | The marts hold **only** the last refreshed period. Confirm the window in the refresh log |
| Response rate suddenly collapses to blank | `PostCallSurveyMode` values changed in the extract | Open the hub's PCS workbook → DATA_QUALITY sheet. Look for "No in-scope PCS responses" |
| Slicer shows dates that no longer exist | Stale slicer cache | Right-click slicer → **Slicer Settings** → untick *Show items deleted from the data source* |
| Refresh fails at 07:00 but works manually | Task Scheduler running as another user | Task Properties → **Run whether user is logged on or not**, and check the account has access to the extract folder |
