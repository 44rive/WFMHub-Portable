# WFMHub Portable

WFMHub Portable turns untouched WFM exports into a durable SQLite database and
a finished Excel report. It runs on Windows x64 without administrator rights,
installed Python, Power Query, Power Pivot, ODBC, or Python in Excel.

Excel is the presentation layer only. Raw data is never loaded into worksheets
or embedded in an Excel Data Model.

## What v0.2.0 does

- Ingests FTE, Verint Schedules & Activities, Storm LILO, Storm Agent Status,
  Verint Forecast, and APBE/APFR queue actuals.
- Uses the FTE Agent sheet as the “our agents” gate. Agent-level extracts are
  kept on exact Agent ID or a unique normalized name match; populated Verint
  `Data Source IDs` remain the operational Agent ID.
- Shows kept and outside-roster row counts in `SOURCE_HEALTH` without changing
  a source file.
- Fingerprints every file and skips unchanged content. A roster change safely
  re-evaluates unchanged agent extracts.
- Stores immutable raw versions, clean models, report marts, source health, and
  persistent correction decisions in `database\wfm.sqlite3`.
- Produces `START_HERE`, `SUMMARY`, `ATTENDANCE`, `GAPS`, `RTA`, `INTRADAY`,
  `DATA_QUALITY`, and `SOURCE_HEALTH` sheets.
- Keeps Verint Forecast isolated from attendance, absence, corrections, and
  payroll logic.

## Windows quick start

Download `WFMHub-Portable-v0.2.0-win-x64.zip` from GitHub Releases—not GitHub's
automatic “Source code” ZIP. Choose **Extract All**, keep the complete `WFMHub`
folder together, double-click `SETUP.cmd` once, then use `WFMHub.cmd` each day.

`SETUP.cmd` first tests the bundled runtime, SQLite transactions/backups, and
Excel write/read support. The portable package uses SQLite built into the
official CPython runtime and adds no third-party database extension or app-local
MSVC package.

Read [`docs/BEGINNER_GUIDE.md`](docs/BEGINNER_GUIDE.md) for click-by-click help.

## Source boundaries

| Source | Used for |
|---|---|
| FTE Agent sheet | Agent scope and organisation context |
| Verint Schedules & Activities | Planned shifts, activities, absence, start/end time |
| Storm LILO | Actual first login and last logout |
| Storm Agent Status | Actual status timeline, conformance, and RTA |
| Verint Forecast | Forecast and staffing requirements only |
| Storm APBE/APFR | Queue actuals only |

If an Agent Status file has zero rows matching the active FTE roster, the hub
creates no RTA result and marks `SOURCE_HEALTH` as `ERROR`. It never substitutes
worldwide rows or invents a current snapshot.

## Developer start

```bash
python -m venv .venv
. .venv/bin/activate
python -m pip install -e .
python -m wfmhub --home . doctor
python -m wfmhub --home . setup --source-root /path/to/untouched-extracts --non-interactive
python -m wfmhub --home . refresh --start 2026-08-01 --end 2026-08-31
python -m unittest discover -s tests -v
```

See [`docs/ARCHITECTURE.md`](docs/ARCHITECTURE.md) for storage grains, identity
scope, business rules, and the extension pattern for future modules.
