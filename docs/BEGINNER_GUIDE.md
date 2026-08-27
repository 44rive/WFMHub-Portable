# Beginner guide — use WFMHub like a simple tool

## The two buttons

Think of WFMHub as a suitcase. It already contains Python, DuckDB and everything
else it needs.

- `SETUP.cmd` prepares the suitcase. Run it once.
- `WFMHub.cmd` opens your work menu. Run it every day.

You do not need administrator rights, installed Python, Power Query, Power
Pivot, ODBC or Python in Excel. Excel only opens the finished reports.

## First setup — one time

1. Download the Windows ZIP from the repository Releases page.
2. Right-click the ZIP and choose **Extract All**.
3. Put the entire extracted `WFMHub` folder somewhere you can write, for example:

   ```text
   C:\Users\YourName\Documents\WFMHub
   ```

4. Do not put it in `Program Files`.
5. Check that `runtime\python.exe` exists inside the extracted `WFMHub` folder.
6. Double-click `SETUP.cmd`.
7. A black window asks for your **source root**. This is the folder containing:

   ```text
   FTE
   Storm
   Verint
   ```

8. Paste that folder path and press Enter.
9. Wait for **Setup complete**. Press a key to close the window.

Setup creates `config\wfmhub.toml` and `database\wfm.duckdb`. It never edits an
extract.

## Your daily routine

1. Put each new untouched export in its normal folder.
2. Keep the original filename.
3. Double-click `WFMHub.cmd`.
4. Choose **1. Refresh all available data + build report**.
5. Wait for **Refresh complete**.
6. Copy the displayed report path or open the newest file in `output`.
7. Open `SOURCE_HEALTH` first.
8. If any required source says `ERROR` or `MISSING`, stop and fix it before using
   attendance, Verint corrections or payroll numbers.
9. Review `ATTENDANCE`, `GAPS`, `RTA` and `INTRADAY`.

Unchanged files are skipped. Adding one daily file does not reload every raw
file into the database.

## Run the whole current month

From `WFMHub.cmd` choose **3. Refresh a custom period + build report**.

Type dates exactly like this:

```text
Start date YYYY-MM-DD: 2026-08-01
End date YYYY-MM-DD:   2026-08-31
```

Dates may also be fixed in `config\wfmhub.toml` under `[period]`. Leave them
blank to let the hub use every available database date. A custom menu date
temporarily overrides the file; it does not change the extracts or config.

Before trusting the monthly gaps, check:

- Every daily LILO date is present.
- The authoritative Verint schedule is loaded.
- Invalid Agent IDs and missing roster rows are understood.
- Agent Status coverage passes before using mid-shift gaps.

Remember the difference:

- **No show:** the daily LILO file exists, the agent row exists, and both times
  are blank on a work schedule.
- **Missing LILO roster row:** the file exists but that Agent ID row does not.
- **Data not loaded:** the required daily file is missing.

Only the first one is a no-show.

## Save correction decisions

1. Open the finished report.
2. Go to `GAPS`.
3. Edit only the blue columns:

   - Confirmed Activity
   - Validation Status
   - Owner
   - Comment
   - Injected Date

4. Save the workbook.
5. Open `WFMHub.cmd`.
6. Choose **5. Import correction decisions**.
7. Paste the saved report path.
8. The hub stores the decisions in DuckDB by stable `Correction ID`.
9. Build another report to confirm they survived.

Do not type decisions into raw extracts. Do not use an `Open` or unreviewed
candidate directly for payroll.

## What each report sheet means

| Sheet | Question it answers |
|---|---|
| `START_HERE` | What should I do and which period/config was used? |
| `SUMMARY` | What are the main numbers? |
| `ATTENDANCE` | Did each scheduled agent arrive and leave around the plan? |
| `GAPS` | What should a human review or inject in Verint? |
| `RTA` | At the newest status snapshot, who matched the plan? |
| `INTRADAY` | What do Storm actuals and Verint forecast say? |
| `DATA_QUALITY` | What is unsafe or incomplete? |
| `SOURCE_HEALTH` | Which files were found and loaded? |

The workbook contains curated results, not raw extracts and not the complete raw
database.

## Forecast later

Forecast support is already present but isolated. Add Verint forecast files to
`Verint\Forecast` and APBE/APFR actual files to their Storm folders. The report
shows them separately. Before comparing variance, define which Storm queues,
LOBs and partners equal the Verint forecast queue. WFMHub will not guess.

Future forecast reports or KPIs are easy to add because Forecast has its own raw
table and mart. It cannot change Attendance, absence or corrections.

## Backup

1. Close any running refresh.
2. Open `WFMHub.cmd`.
3. Choose **7. Create database backup**.
4. The dated copy appears in `backups`.

Back up before a software upgrade or important month-end run. Never delete the
original database first when restoring; rename the damaged file and copy the
known-good backup into `database\wfm.duckdb`.

## Quick troubleshooting

| Problem | Check |
|---|---|
| The black window closes | Run `SETUP.cmd`, then read the newest file in `logs` |
| Embedded Python is missing | Download the portable release ZIP, choose **Extract All**, and verify `runtime\python.exe`; do not use GitHub's Source code ZIP |
| Source says MISSING | Open `config\wfmhub.toml` and correct `source_root` |
| Report is empty | Check the selected dates and `SOURCE_HEALTH` |
| RTA is old | RTA is only as fresh as the newest Agent Status export |
| Employee is unmatched | Verify the Agent ID; do not force a name match |
| Database is locked | Close the other WFMHub refresh and try again |
| A decision disappeared | Import the edited report; editing an output alone does not update DuckDB |
| A runtime DLL cannot load | Re-extract the complete ZIP; do not move DLLs out of `runtime` |

Golden rule: if Source Health or Data Quality is red, stop and understand it.
