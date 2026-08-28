# PCS Daily Report — Pivot Model Specification

Technical build sheet. Pair with `PCS_Setup_ClickByClick.md` for the ribbon steps.

---

## 1. Source dataset

**WFMHub export:** `pcs_agent_day`
**Hub table:** `mart.agent_pcs_day`
**Grain:** one row per agent per business date
**Volume:** roughly headcount x days in window

The export is `SELECT *`, so column order below is exactly the file's column order.

### The 39 columns

| # | Column | Type | Role in the model |
|---|---|---|---|
| 1 | `agent_day_key` | text | **Natural key** (`YYYYMMDD-agentid`). Dedup key. |
| 2 | `business_date` | date | Date axis |
| 3 | `agent_id` | text | Agent key |
| 4 | `agent_name` | text | Agent label |
| 5 | `team_leader` | text | **Slicer / row field** |
| 6 | `ops_manager` | text | **Slicer / row field** |
| 7 | `lob` | text | **Slicer / row field** |
| 8 | `market` | text | Slicer |
| 9 | `language` | text | Slicer |
| 10 | `location` | text | Slicer |
| 11 | `call_legs` | int | Sum |
| 12 | `handled_calls` | int | **Sum — AHT denominator** |
| 13 | `inbound_calls` | int | Sum |
| 14 | `outbound_calls` | int | Sum |
| 15 | `talk_seconds` | int | **Sum — numerator** |
| 16 | `hold_seconds` | int | **Sum — numerator** |
| 17 | `wrap_seconds` | int | **Sum — numerator** |
| 18 | `handle_seconds` | int | **Sum — numerator** |
| 19 | `average_talk_seconds` | float | ⛔ **Never aggregate** |
| 20 | `average_hold_seconds` | float | ⛔ **Never aggregate** |
| 21 | `average_wrap_seconds` | float | ⛔ **Never aggregate** |
| 22 | `average_handle_seconds` | float | ⛔ **Never aggregate** |
| 23 | `pcs_enabled_calls` | int | **Sum — response-rate denominator** |
| 24 | `survey_responses` | int | **Sum — numerator + denominator** |
| 25 | `response_rate` | float | ⛔ **Never aggregate** |
| 26 | `q1_response_count` | int | **Sum — Q1 denominator** |
| 27 | `q1_score_sum` | float | **Sum — Q1 numerator** |
| 28 | `q1_average` | float | ⛔ **Never aggregate** |
| 29 | `q2_response_count` | int | **Sum — Q2 denominator** |
| 30 | `q2_score_sum` | float | **Sum — Q2 numerator** |
| 31 | `q2_average` | float | ⛔ **Never aggregate** |
| 32 | `pcs_score_count` | int | **Sum — PCS denominator** |
| 33 | `pcs_score_sum` | float | **Sum — PCS numerator** |
| 34 | `pcs_average` | float | ⛔ **Never aggregate** |
| 35 | `top_box_responses` | int | **Sum — numerator** |
| 36 | `low_score_responses` | int | **Sum — numerator** |
| 37 | `top_box_percent` | float | ⛔ **Never aggregate** |
| 38 | `low_score_percent` | float | ⛔ **Never aggregate** |
| 39 | `comments_count` | int | Sum |

### ⛔ The ten forbidden columns

Columns 19, 20, 21, 22, 25, 28, 31, 34, 37 and 38 are **pre-computed agent/day ratios**. They exist so a
single agent-day row is readable on its own. **They must never be dropped into a pivot's Values
area.**

Dragging `pcs_average` into Values gives you an *average of agent-day averages* — an agent with
1 response counts as much as an agent with 40. That is the single most common way this report
gets built wrong.

Every one of them is reproduced correctly as a measure in section 3. **Use the measures.**

To make the mistake impossible, the Power Query step in section 2 **removes all nine columns**
before loading. They cannot be misused if they are not in the model.

---

## 2. Power Query — `qryPcsAgentDay`

**Source:** folder `…\WFM-PCS-Daily\data\current\`
**Load target:** Data Model only — *Only Create Connection*, **not** to a worksheet.

Loading to the model rather than a sheet keeps agent-level data out of the workbook grid
entirely, and removes the 1,048,576-row ceiling.

### Steps in order

| # | Step | Purpose |
|---|---|---|
| 1 | `Folder.Files` on `data\current\` | Folder source as agreed |
| 2 | Filter `Extension` = `.csv` | Ignore manifests and stray files |
| 3 | Sort `Date modified` **descending** | Newest first |
| 4 | Combine files, promote headers | One table |
| 5 | Set types explicitly | Dates as Date, counters as Whole Number, sums as Decimal |
| 6 | **Remove the ten forbidden columns** | Make the classic mistake impossible |
| 7 | **Remove duplicates on `agent_day_key`** | ⚠️ **Safety net — see below** |
| 8 | Add `ISO Week`, `Month`, `Day Name` | Time grouping |
| 9 | Add period flags | Drives the period slicer |
| 10 | Close & Load To → **Only Create Connection** + **Add to Data Model** | |

### ⚠️ Step 7 is not optional

`data\current\` normally holds exactly one file, because the refresh job overwrites it in place.
But if a second export ever lands there — someone exports manually, or an archive copy gets
copied back — the folder query combines both files and **every single count doubles silently**.
Response rate would survive (both halves double), but PCS-enabled calls, responses, comments and
handled calls would all be wrong, and nothing on the face of the report would look odd.

Sorting newest-first (step 3) then removing duplicates on `agent_day_key` (step 7) keeps the
newest row per agent-day. This is the same newest-file-wins rule the hub itself applies when it
deduplicates call legs, so the behaviour is consistent with the system it draws from.

### Step 9 — period flags

```m
// Today (partial)
Is Today       = [business_date] = Date.From(DateTime.LocalNow())

// Yesterday — complete day
Is D-1         = [business_date] = Date.AddDays(Date.From(DateTime.LocalNow()), -1)

// Rolling 7 complete days, excludes today
In Rolling 7   = [business_date] >= Date.AddDays(Date.From(DateTime.LocalNow()), -7)
                 and [business_date] < Date.From(DateTime.LocalNow())

// Month to date, includes today
In MTD         = [business_date] >= Date.StartOfMonth(Date.From(DateTime.LocalNow()))
```

These recalculate on every refresh, so the report rolls forward on its own with no maintenance.

**Rolling 7 deliberately excludes today**, because today is partial and would drag the average
toward whatever has been collected so far this morning.

---

## 3. DAX measures

Create in the Data Model. Every ratio divides a **sum by a sum** — never averages a ratio.

`DIVIDE` returns **blank** on a zero or blank denominator. That is exactly the hub's own
behaviour, and it is what keeps "no responses" from rendering as `0.00`.

**Note the asymmetry, it is intentional.** An agent with 40 PCS-enabled calls and zero responses
gets `Response Rate` = **0.0%** (the denominator is real, so 0% is a true fact worth acting on)
but `PCS Average` = **blank** (nothing to average). An agent with no enabled calls at all gets
blank for both. A ratio is shown whenever its denominator exists.

### Volume

```dax
Call Legs         := SUM ( pcs_agent_day[call_legs] )
Handled Calls     := SUM ( pcs_agent_day[handled_calls] )
Inbound Calls     := SUM ( pcs_agent_day[inbound_calls] )
Outbound Calls    := SUM ( pcs_agent_day[outbound_calls] )
Agents            := DISTINCTCOUNT ( pcs_agent_day[agent_id] )
Agent Days        := COUNTROWS ( pcs_agent_day )
```

### PCS core

```dax
PCS Enabled Calls := SUM ( pcs_agent_day[pcs_enabled_calls] )
Survey Responses  := SUM ( pcs_agent_day[survey_responses] )

Response Rate     := DIVIDE ( [Survey Responses], [PCS Enabled Calls] )

PCS Average       := DIVIDE ( SUM ( pcs_agent_day[pcs_score_sum] ),
                              SUM ( pcs_agent_day[pcs_score_count] ) )

Q1 Average        := DIVIDE ( SUM ( pcs_agent_day[q1_score_sum] ),
                              SUM ( pcs_agent_day[q1_response_count] ) )

Q2 Average        := DIVIDE ( SUM ( pcs_agent_day[q2_score_sum] ),
                              SUM ( pcs_agent_day[q2_response_count] ) )

Top Box %         := DIVIDE ( SUM ( pcs_agent_day[top_box_responses] ),
                              [Survey Responses] )

Low Score %       := DIVIDE ( SUM ( pcs_agent_day[low_score_responses] ),
                              [Survey Responses] )

Comments          := SUM ( pcs_agent_day[comments_count] )
```

### Handle time

```dax
Talk Seconds      := SUM ( pcs_agent_day[talk_seconds] )
Hold Seconds      := SUM ( pcs_agent_day[hold_seconds] )
Wrap Seconds      := SUM ( pcs_agent_day[wrap_seconds] )
Handle Seconds    := SUM ( pcs_agent_day[handle_seconds] )

AHT Seconds       := DIVIDE ( [Handle Seconds],   [Handled Calls] )
Avg Talk Seconds  := DIVIDE ( [Talk Seconds],     [Handled Calls] )
Avg Hold Seconds  := DIVIDE ( [Hold Seconds],     [Handled Calls] )
Avg Wrap Seconds  := DIVIDE ( [Wrap Seconds],     [Handled Calls] )
```

### Reliability gate

```dax
Min Responses     := 5

Ranked PCS Average :=
    IF ( [Survey Responses] >= [Min Responses], [PCS Average] )

Data Sufficiency :=
    IF (
        ISBLANK ( [Survey Responses] ) || [Survey Responses] = 0,
        "No responses",
        IF ( [Survey Responses] < [Min Responses], "Insufficient", "Rankable" )
    )
```

`Ranked PCS Average` is blank below the threshold, so agents with thin data fall out of the
ranking without being hidden — their response count is still on the row.

**To change the threshold, edit `Min Responses` in one place.** Nothing else moves.

### Exception measures

```dax
Enabled No Response :=
    CALCULATE ( [Agents],
        FILTER ( VALUES ( pcs_agent_day[agent_id] ),
                 [PCS Enabled Calls] > 0 && [Survey Responses] = 0 ) )

Low Score Flag :=
    IF ( [Survey Responses] >= [Min Responses] && [Low Score %] > 0.20, "REVIEW" )

Response Rate Flag :=
    IF ( [PCS Enabled Calls] >= 20 && [Response Rate] < 0.02, "LOW COLLECTION" )
```

### Number formats

| Measure | Format |
|---|---|
| PCS / Q1 / Q2 / Ranked average | `0.00` |
| Response Rate, Top Box %, Low Score % | `0.0%` |
| AHT and the Avg …Seconds measures | `0` (seconds) or `[m]:ss` |
| All counts | `#,##0` |

Set these **on the measure**, not on the pivot. They then follow the measure everywhere.

---

## 4. Pivot layouts

Five pivots, all from the one Data Model. Sheet names in bold.

### **ROLLUP** — Ops Manager view

```
Rows    : ops_manager > lob
Values  : PCS Average | Response Rate | Top Box % | Low Score %
          Survey Responses | PCS Enabled Calls | Handled Calls | AHT Seconds
Slicers : Period | language | market
Layout  : Tabular, repeat item labels off, grand total on
```
No agent names. This is the page that gets forwarded.

### **TL_SCORECARD** — Team Leader view

```
Rows    : team_leader > agent_name
Values  : Ranked PCS Average | Survey Responses | Response Rate
          Top Box % | Low Score % | PCS Enabled Calls | AHT Seconds | Comments
Slicers : Period | ops_manager | lob | Data Sufficiency
Sort    : Ranked PCS Average ascending — worst first, that is where coaching goes
Layout  : Tabular, subtotals at top
```

⚠️ **`Survey Responses` sits immediately beside the average, deliberately.** An average without
its response count invites bad decisions.

### **AGENT_RANKING** — flat league table

```
Rows    : agent_name > team_leader > lob
Values  : Ranked PCS Average | Survey Responses | Top Box % | Low Score % | AHT Seconds
Filter  : Data Sufficiency = "Rankable"
Sort    : Ranked PCS Average descending
```
Pair with a second, visually separated pivot filtered to
`Data Sufficiency = "Insufficient"` or `"No responses"`, headed
**"Shown for completeness — not ranked"**.

### **DAILY_TREND** — direction

```
Rows    : business_date
Values  : PCS Average | Response Rate | Survey Responses | PCS Enabled Calls
Slicers : ops_manager | team_leader | lob | language
Chart   : Combo — PCS Average as line on the secondary axis,
          Survey Responses as columns on the primary axis
```
Volume behind the line, so a swing on 3 responses reads as a swing on 3 responses.
**Today's bar is partial** — annotate it.

### **EXCEPTIONS** — the action list

Four blocks on one sheet:

1. **Enabled but silent** — `PCS Enabled Calls > 0` and `Survey Responses = 0`, sorted by
   enabled calls descending
2. **Low-score concentration** — `Low Score Flag = "REVIEW"`
3. **Collection failure** — `Response Rate Flag = "LOW COLLECTION"`
4. **Movers** — largest fall in PCS Average, rolling 7 vs prior 7, rankable agents only

Keep it to **one screen**. An exception list that needs scrolling gets ignored.

---

## 5. Conditional formatting

Apply to measure cells only. Use *Format all cells showing "…" values* so it survives refresh.

| Measure | Rule |
|---|---|
| PCS Average | 3-colour scale, red 1.0 → amber 3.5 → green 5.0. **Fixed endpoints, not percentiles** — percentiles rescale on every refresh and last week stops matching this week. |
| Response Rate | Red below 2%, amber 2–5%, green above 5%. *Placeholder — set from real baseline after two weeks.* |
| Low Score % | Red above 20% |
| Survey Responses | Data bar, so thin data is visible at a glance |
| Blank cells | **Leave unformatted.** No fill, no "0". Blank must read as absent. |

---

## 6. Freshness stamp

Top-right of every sheet, driven from `data\current\last_refresh.txt`:

```
Data as at : 2026-08-28 15:00
Window     : 2026-07-01 to 2026-08-28
Rule ver   : 2026.08.1
Rows       : 12,480
```

Non-negotiable at a 3-hourly cadence. Without it, a 09:00 figure gets quoted at 16:00.
**Today's figures are partial and rise through the day.**

---

## 7. Verification before first send

Run these once against a known period and reconcile to the hub's own PCS workbook
(`output/quality_pcs/`). They must tie **exactly** — same source, same grain.

| # | Check | Expected |
|---|---|---|
| 1 | Grand total `PCS Enabled Calls` | = hub SUMMARY "PCS-enabled calls" |
| 2 | Grand total `Survey Responses` | = hub SUMMARY "Survey responses" |
| 3 | Grand total `PCS Average` | = hub SUMMARY "PCS average" |
| 4 | Grand total `Response Rate` | = hub SUMMARY "Response rate" |
| 5 | `COUNTROWS` vs export row count | Equal — proves no duplicate-file doubling |
| 6 | Agent with zero responses | **Blank**, not `0.00` |
| 7 | Sum of TL-level averages vs grand total | Grand total is **not** the mean of the TL means |

**Check 5 is the one to repeat monthly.** It is the only cheap detector of a stray file in
`data\current\`.
