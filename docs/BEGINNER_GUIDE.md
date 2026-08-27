# Beginner guide — use WFMHub like a simple tool

## The two files you click

Think of WFMHub as a suitcase. Everything it needs is already inside.

- `SETUP.cmd` checks and prepares the suitcase. Run it once.
- `WFMHub.cmd` opens the daily menu.

You do not install Python, SQLite, Power Query, Power Pivot, ODBC, or Python in
Excel. The database stays outside Excel. Excel receives only the small finished
report sheets.

## First setup — one time

1. Download `WFMHub-Portable-v0.2.0-win-x64.zip` from **Releases**.
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

## What “our agents” means

The FTE `Agent` sheet is the list of people allowed into agent-level data.

For every schedule, LILO, and Agent Status row, WFMHub asks:

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

## Daily routine

1. Put new untouched exports in their normal folders.
2. Keep their original filenames.
3. Double-click `WFMHub.cmd`.
4. Choose **1. Refresh all available data + build report**.
5. Wait for **Refresh complete**.
6. Open the newest workbook in `output`.
7. Open `SOURCE_HEALTH` first.
8. Stop if a required source says `ERROR` or `MISSING`.
9. Then review `ATTENDANCE`, `GAPS`, `RTA`, and `INTRADAY`.

Unchanged files are fingerprinted and skipped. Their already-active kept and
outside-roster counts still appear in the report.

## Run the current month or another date range

For the current month, choose **2. Refresh current month + build report**.

For any dates, choose **3. Refresh a custom period + build report** and type:

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

Menu dates temporarily override setup dates. Report-only dates truly filter
every dated KPI and sheet; they are not labels only.

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
8. In `WFMHub.cmd`, choose **5. Import correction decisions** and paste the saved
   workbook path.

The whole import is atomic: one invalid later row cancels the entire import.
WFMHub stores accepted decisions in SQLite by stable `Correction ID`; a later
refresh does not erase them.

## Report sheets

| Sheet | Meaning |
|---|---|
| `START_HERE` | Period and daily instructions |
| `SUMMARY` | Main period-filtered KPIs |
| `ATTENDANCE` | One scheduled admitted Agent ID/day |
| `GAPS` | Reviewable Verint correction candidates |
| `RTA` | Adherence at newest admitted status timestamp |
| `INTRADAY` | Storm actuals and Verint forecast, still separate |
| `DATA_QUALITY` | Unsafe, incomplete, or review conditions |
| `SOURCE_HEALTH` | Files, kept rows, excluded rows, and failures |

The workbook contains no raw extract sheet, Power Query connection, or embedded
Data Model.

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

v0.1 used a database extension blocked by the work computer. v0.2 starts a new
SQLite database.

1. Extract v0.2 into a new folder. Do not paste it over v0.1.
2. Run v0.2 `SETUP.cmd` and select the same untouched source root.
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
| A DLL/application-control error appears | Re-extract the complete v0.2 ZIP; do not overlay v0.1 or move runtime files |
| Source is `MISSING` | Correct `source_root` in `config\wfmhub.toml` |
| Agent Status is `ERROR`, 0 kept | The export does not match FTE IDs/unique names; obtain the correct scoped export |
| RTA is empty | Fix Agent Status scope/freshness first; WFMHub will not invent RTA |
| Report is empty | Check selected dates, `SOURCE_HEALTH`, and FTE scope |
| Database is locked | Close the other WFMHub refresh and retry |
| Decision disappeared | Import the edited report; editing Excel alone does not update SQLite |

Golden rule: if Source Health or Data Quality is red, stop and understand it.
