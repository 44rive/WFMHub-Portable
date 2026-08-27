# Beginner guide — use WFMHub like a simple tool

## The two files you click

Think of WFMHub as a suitcase. Everything it needs is already inside.

- `SETUP.cmd` checks and prepares the suitcase. Run it once.
- `WFMHub.cmd` opens the daily menu.

You do not install Python, SQLite, Power Query, Power Pivot, or ODBC. The
database stays outside Excel. Normal report workbooks contain small curated
sheets; detailed data is created only when you explicitly choose an export.

## First setup — one time

1. Download `WFMHub-Portable-v0.3.2-win-x64.zip` from **Releases**.
2. Do not download GitHub's “Source code” ZIP.
3. Right-click the downloaded ZIP and choose **Extract All**.
4. Put the entire extracted `WFMHub` folder in a place where you can write, for
   example:

   ```text
   C:\Users\YourName\Documents\WFMHub
   ```

5. Do not use `Program Files`, a network drive, OneDrive, or another synced
   folder for the database.
6. Open the folder and confirm `runtime\python.exe` exists.
7. Double-click `SETUP.cmd`.
8. First, the black window runs **WFMHUB SYSTEM CHECK**. Continue only when it
   says **SYSTEM CHECK PASSED**.
9. Paste your source-root folder. It is the folder containing:

   ```text
   FTE
   Storm
   Verint
   ```

10. Press Enter and wait for **Setup complete**.

Setup creates `config\wfmhub.toml` and `database\wfm.sqlite3`. It does not edit,
move, rename, or delete an extract.

## Standard FTE file

The release contains `templates\FTE Count.xlsx`.

1. Open that template and read `START_HERE`.
2. On `Agent`, paste/update one row per agent.
3. Keep `Client ID` and `Name` populated. Store Client ID as text so leading
   zeros survive.
4. Keep the supplied headers.
5. Save a working copy as `FTE\FTE Count.xlsx` under your source root.

The other columns—Status, Team leader, Ops Manager, LOB, Market, Language,
Location, City, FTE, and leaver end date—are recommended enrichment. WFMHub can
find a safe renamed/offset roster, but if two worksheets look authoritative or
two ID/name aliases exist, it stops and tells you exactly where instead of
guessing.

## What “our agents” means

The FTE `Agent` sheet is the list of people allowed into agent-level data.

For every schedule, LILO, Agent Status, and Call-by-Call row, WFMHub asks:

1. Does its Agent ID exist in FTE? Keep it.
2. If not, does its cleaned name match exactly one FTE person? Keep it.
3. Otherwise, exclude it as outside roster.

“Cleaned name” only ignores case, accents, punctuation, and extra spaces. It is
not a fuzzy guess. A populated Verint `Data Source IDs` value stays the
operational Agent ID.

The source file itself remains untouched. `SOURCE_HEALTH` shows:

- `Row Count`: rows kept for the hub.
- `Scoped Out Count`: worldwide/outside-roster rows excluded.
- `Rejected Count`: malformed or flagged rows.

If FTE changes, WFMHub automatically rechecks the same unchanged extracts. If
every Agent Status row is outside roster, RTA stays empty and source health says
`ERROR`. Do not use those status rows.

## Read the dashboard status bar

`WFMHub.cmd` opens with the WFMHUB logo and a status panel. It is your quick
health check before choosing a menu action:

- `READY` means the last refresh succeeded and no current error requires
  attention.
- `REVIEW` means the hub worked, but one or more quality checks need a human
  look.
- `CHECK DATA` means a source or data-quality error needs attention.
- `SETUP REQUIRED` or `READY TO REFRESH` tells you the next safe action.
- `DB` is the current SQLite database size, including its live WAL file.
- `Latest data` is the newest business date actually loaded in source health.
  The family in parentheses matters: a future date marked `(forecast)` is a
  forecast horizon, not proof that LILO, schedules, calls or actuals are current.

The status panel also shows the last refresh time, selected period, active-agent
count, healthy source count, and current error/review totals. It redraws when
you press Enter after an action. Menu numbers remain unchanged.

## Daily routine

1. Put new untouched exports in their normal folders.
2. Keep their original filenames.
3. Double-click `WFMHub.cmd`.
4. Choose **1. Refresh hub data**.
5. Choose the source group: All, Operations, Intraday, or Agent PCS.
6. Choose the date period.
7. Choose one report, all reports, or no report.
8. Wait for **Refresh complete**.
9. Open the report's `SOURCE_HEALTH` first.
10. Stop if a required source says `ERROR` or `MISSING`.

Unchanged files are fingerprinted and skipped. Their already-active kept and
outside-roster counts still appear in the report.

### Understand the progress bar

While WFMHub works, the same line in the black window changes, for example:

```text
WFMHub [############----------------]  43% Building attendance
```

The words tell you exactly what it is doing. The percentage covers the whole
job, not one source file. Very large LILO and Call-by-Call files temporarily
show `working` plus the number of rows scanned. Large clean exports show the
number of rows written. Do not close the window while the line is moving or its
row count is increasing. A completed job reaches `100%`; an error changes the
line to `FAILED` and then prints the normal error explanation.

The bar uses plain Windows CMD characters and needs no installation. It is
normally hidden when output is sent to a log or automation file. For support,
run `set WFMHUB_PROGRESS=1` before the command to force the bar, or
`set WFMHUB_PROGRESS=0` to hide it.

## Run the current month or another date range

Every report, export and custom job offers these date choices:

```text
Today
Yesterday
Current week
Current month
Previous month
Custom dates
All available dates
Saved default dates
```

For custom dates, type:

```text
Start date YYYY-MM-DD: 2026-08-01
End date YYYY-MM-DD:   2026-08-31
```

You can set default dates in `config\wfmhub.toml`:

```toml
[period]
start = "2026-08-01"
end = "2026-08-31"
```

Use straight quotation marks. Leave both blank to use all available dates:

```toml
start = ""
end = ""
```

Menu dates temporarily override setup dates. WFMHub rebuilds the selected-period
marts before a report, clean export, or custom job, so the dates are real
filters—not labels only. Refreshing a new/changed source stores every row date
found in that file so older periods remain available later.

## Files containing several days

The filename is not the business-date authority. WFMHub reads row dates from:

- Schedule shift/event timestamps or the schedule date marker.
- LILO `Date`/`Business Date`/`Extract Date`/`Report Date`, then login/logout.
- Agent Status start timestamp.
- Call start timestamp.
- Forecast and APBE/APFR Date columns.

A filename can contain one date, `start - end`, or no date. One special case is
a multi-day LILO row with both login and logout blank: it must have a row-level
Date column. Without that, no date can be known, so WFMHub rejects and counts
the row instead of inventing a no-show day.

## No-show versus missing data

- **No show:** LILO file loaded, agent row present, both login/logout blank, and
  the schedule says Work.
- **Missing LILO roster row:** LILO file loaded, but that agent row is absent.
- **Data not loaded:** the needed daily LILO file is missing.
- **Identity not in LILO:** the scheduled Agent ID was never found in admitted
  LILO history.
- **Incomplete LILO:** only one boundary is present.

Only the first result is a no-show.

## Monthly Verint correction exercise

1. Refresh the first through last day of the month.
2. Confirm all expected LILO dates and the authoritative Verint schedule in
   `SOURCE_HEALTH`.
3. Check `DATA_QUALITY`; resolve `ERROR` before payroll or injection work.
4. Open `GAPS`.
5. Review the detected interval, reason, confidence, and suggested activity.
6. Edit only the blue columns:

   - Confirmed Activity
   - Validation Status
   - Owner
   - Comment
   - Injected Date

7. Save the workbook.
8. In `WFMHub.cmd`, choose **6. Import correction decisions** and paste the saved
   workbook path.

The whole import is atomic: one invalid later row cancels the entire import.
WFMHub stores accepted decisions in SQLite by stable `Correction ID`; a later
refresh does not erase them.

## Separate report packs

Keep different work grains in different workbooks:

- `output\operations`: attendance, gaps, RTA and operational data quality.
- `output\intraday`: APBE/APFR actuals and separate Verint forecast.
- `output\quality_pcs`: call performance, Agent PCS, trends, survey responses,
  PCS data quality and copy/paste Python-in-Excel recipes.

Call-by-call rows can be enormous. They will stay in SQLite, deduplicated by a
stable call-leg key. The PCS workbook contains agent summaries, trends and
bounded survey responses—not millions of raw calls. Correction imports remain
exclusive to the Operations `GAPS` sheet.

PCS uses valid configured numeric answers. It averages configured questions
within each response, then gives each response equal weight in the agent PCS
average. A missing denominator displays blank, never zero.

## Export clean data

Choose **3. Export clean data**, then select a dataset, dates and CSV/XLSX.
Available exports include calls, PCS responses, PCS agent/day, attendance,
gaps, schedules, events, LILO, Agent Status, actuals, forecast and source
health.

Use CSV for large call data. XLSX stops before Excel's row limit. Every export
gets a `.manifest.txt` beside it containing dates, row count and generation
time. Outputs go to `output\data_exports`; extracts are not changed.

## Custom Lab

1. Open `custom\jobs` for Python or `custom\sql` for SQL.
2. Copy the underscore template.
3. Rename the copy without the leading underscore.
4. Paste/edit your code in the copy.
5. Choose **4. Run custom Python or SQL analysis**.
6. Select the job and dates.

The hub API passed to Python jobs has a read-only database connection. Python
jobs are still trusted executable code, so run only code you understand. No
`pip` runs on the work machine; use the standard library and shipped Excel
libraries. Custom output goes to `output\custom`.

## Forecast and future features

Verint Forecast is used only for forecast and required staffing. APBE/APFR is
used only for actual queue results. WFMHub does not guess how a forecast queue
maps to actual queue/LOB data.

New KPIs or reports can be added later without replacing the source extracts or
attendance rules because forecast, actuals, agent facts, and report marts are
separate layers.

## Backup and restore

1. Open `WFMHub.cmd`.
2. Choose **7. Create database backup**.
3. Find the dated `.sqlite3` copy in `backups`.

SQLite uses companion WAL files while running, so do not make a database backup
with ordinary copy/paste during refresh. Use menu option 7.

To restore, close WFMHub, preserve/rename the current database, copy the chosen
backup to `database\wfm.sqlite3`, and rerun **9. Run system check**. Never delete
the only copy first.

## Upgrade from v0.1

v0.1 used a database extension blocked by the work computer. v0.2 and v0.3 use
SQLite. v0.3 upgrades a v0.2 SQLite database with an automatic pre-migration
backup and additive migration.

1. Extract v0.3 into a new folder. Do not paste it over v0.1.
2. Run v0.3 `SETUP.cmd` and select the same untouched source root.
3. Run a full refresh so SQLite rebuilds from the extracts.
4. Keep the complete v0.1 folder and its `wfm.duckdb`; v0.2 never opens or
   deletes it.
5. If correction decisions exist only in an old Excel report, import that saved
   report into v0.2.

## Troubleshooting

| Problem | What to do |
|---|---|
| System check fails | Copy the exact FAIL line; setup stopped before database creation |
| `py` is not recognized | Ignore `py`; the release uses `runtime\python.exe` only |
| A DLL/application-control error appears | Re-extract the complete v0.3 ZIP; do not overlay v0.1 or move runtime files |
| Source is `MISSING` | Correct `source_root` in `config\wfmhub.toml` |
| Agent Status is `ERROR`, 0 kept | The export does not match FTE IDs/unique names; obtain the correct scoped export |
| RTA is empty | Fix Agent Status scope/freshness first; WFMHub will not invent RTA |
| PCS average is blank | No valid configured survey response exists for the FTE-scoped calls and dates |
| Multi-day LILO rejects blank rows | Add a row-level Date column; filename range cannot identify a blank row's day |
| Report is empty | Check selected dates, `SOURCE_HEALTH`, and FTE scope |
| Database is locked | Close the other WFMHub refresh and retry |
| Decision disappeared | Import the edited report; editing Excel alone does not update SQLite |

Golden rule: if Source Health or Data Quality is red, stop and understand it.
