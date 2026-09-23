# Operations runbook

## Setup

1. Extract the release ZIP to a local folder outside OneDrive.
2. Run `SETUP.cmd`.
3. Paste the root containing `FTE`, `Storm`, and `Verint`.
4. Confirm **SYSTEM CHECK PASSED** and **Setup complete**.
5. Run `WFMHub.cmd` and refresh all sources.

If setup says embedded Python is missing, the full release was not extracted.
Download the release asset, choose **Extract All**, and keep every `.cmd` beside
the `_system` folder.

## Daily cycle

1. Add the newest extracts to their established source folders.
2. Open `WFMHub.cmd` and refresh the relevant group.
3. Check source health, latest business date, rejected files, and quality issues.
4. Use `RTM Daily Control.xlsx` or the local workbench during the day.
5. Use `Attendance Review.xlsx` after completed shifts; correct Verint outside
   the Hub and load the next Activities extract.
6. Generate or refresh the product needed for delivery.

Do not open a workbook while WFMHub is replacing that same file. If Excel or a
sync client holds it, WFMHub preserves the existing workbook and explains the
lock. Close it, wait for sync to finish, and retry.

## PCS cycle

Use **PCS Report & Coaching > Update PCS data** to process new FTE and calls.
Then close Excel and choose **Sync PCS workbook** in the PCS menu. The Hub
updates the six data tables in place and leaves coaching actions untouched. Coaching
owners edit the coaching table only. See [PCS.md](PCS.md).

## Bonus Management

Build the permanent workbook once, then paste monthly rows into `Raw_Data`.
Validate KPI and policy inputs before sharing a copy. The workbook formulas
update its Results and Dashboard in Excel. See [BONUS.md](BONUS.md).

## Realisations

Open `Realisations.xlsx` for daily actual-versus-forecast and final Verint
Activities absenteeism. Use `HOURS_BY_AGENT`/`HOURS_BY_LOB` for exclusive
scheduled-hour allocation, `AUX_BY_AGENT`/`AUX_BY_TL` for exact Agent Status
AUX hours, and `ABS_BY_AGENT`/`ABS_BY_LOB` for final absence. Unpaid leave is
shown separately and excluded from the Realisations absence numerator; No Show
still counts. Review incomplete final-activity days before sharing.

## Local web app

Run `WEBAPP.cmd`. It binds to `127.0.0.1`, so only the current computer can
access it. Filters are hierarchical and multi-select where the page supports
them. Calculations come from SQLite; the browser is presentation and job
control, not a second calculation engine.

## Date selection

Menu date choices apply to model/report output. Source ingestion still reads
changed files and their row-level dates. `All available dates` ignores saved
period bounds. `Saved default dates` uses `[period]` in `config/wfmhub.toml`.

## Upgrade

1. Extract the new release into a new local folder.
2. Close WFMHub and Excel.
3. Run `UPGRADE.cmd` in the new folder.
4. Paste the previous WFMHub folder path.
5. Run `SETUP.cmd` once in the new folder, then refresh.

The previous folder is not modified. Keep it until the new release has passed
your own service, attendance, PCS, and staffing checks.

## Troubleshooting

- Read the newest `_system/logs/wfmhub_YYYYMMDD.log` from the bottom upward.
- Run **System and advanced tools > Run system check**.
- Run **Show source health and date coverage**.
- A long first refresh is normal when large Call-by-Call/Agent Status files are
  parsed. Repeated unchanged refreshes should be materially faster.
- Never delete the SQLite database to solve a calculation issue. Back it up,
  fix the source/config/code cause, and refresh the affected models.
