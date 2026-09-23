# WFMHub Portable

WFMHub turns unchanged Verint, Storm, FTE, and Call-by-Call extracts into a
durable local WFM database, operational Excel products, clean CSV feeds, and a
localhost manager workbench.

## First use on Windows

1. Download the Windows ZIP from GitHub Releases—not the GitHub source ZIP.
2. Choose **Extract All** to a normal local folder outside OneDrive.
3. Run `SETUP.cmd`.
4. Paste the folder that contains your `FTE`, `Storm`, and `Verint` folders.
5. Run `WFMHub.cmd` and choose **Refresh source data once**.
6. Use `WEBAPP.cmd` for the local manager workbench or generate the workbook
   needed from the menu.

The live database must stay on the local disk. Finished reports may be copied
to SharePoint or OneDrive after generation.

## Daily source layout

```text
WFM Database/
├── FTE/FTE Count.xlsx
├── Storm/Agent Status/*.csv
├── Storm/LILO/*.csv
├── Storm/Call by Call/*.csv
├── Verint/Schedules & Activities/*.txt
└── Verint/Forecast/*.txt
```

Add new extracts; do not alter or delete older evidence merely to refresh the
Hub. It fingerprints files, skips unchanged inputs, selects current overlapping
versions deterministically, and reads row-level dates from multi-day files.

## PCS in plain language

`PCS Live Tracker.xlsx` is one permanent shared file. WFMHub processes new FTE
and Call-by-Call evidence and replaces the fixed CSV files under `Feed/PCS`.
Close Excel and choose **Sync PCS workbook** in WFMHub to update the six data
tables in place. Team leaders
and Quality update only the coaching table. See [PCS workflow](docs/PCS.md).

## Updates

Extract each release into a new local folder, run `UPGRADE.cmd` there, and point
it to the previous WFMHub folder. The previous installation is read-only; the
database is copied with SQLite's backup API and validated before migrations.

## Help and product truth

Start with [PROJECT.md](PROJECT.md). It states every source role, business
boundary, current output, configuration owner, and code location. The beginner
runbook is [docs/OPERATIONS.md](docs/OPERATIONS.md).

WFMHub is maintained by Anass ASSRI.
