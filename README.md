# WFMHub Portable

WFMHub Portable turns untouched WFM exports into a durable SQLite database and
a set of finished Excel reports and clean-data exports. It runs on Windows x64 without administrator rights,
installed Python, Power Query, Power Pivot, ODBC, or Python in Excel.

Normal report workbooks contain curated summaries only and no Excel Data Model.
Detailed clean rows appear in Excel only when the user explicitly chooses an
XLSX clean-data export; CSV remains the standard for large call datasets.

## What v0.3.0 does

- Ingests FTE, Verint Schedules & Activities, Storm LILO, Storm Agent Status,
  Verint Forecast, APBE/APFR queue actuals, and Storm Call by Call.
- Uses the FTE Agent sheet as the “our agents” gate. Agent-level extracts are
  kept on exact Agent ID or a unique normalized name match; populated Verint
  `Data Source IDs` remain the operational Agent ID.
- Shows kept and outside-roster row counts in `SOURCE_HEALTH` without changing
  a source file.
- Fingerprints every file and skips unchanged content. A roster change safely
  re-evaluates unchanged agent extracts.
- Stores immutable raw versions, clean models, report marts, source health, and
  persistent correction decisions in `database\wfm.sqlite3`.
- Produces separate Operations, Intraday and Agent PCS workbooks.
- Calculates agent call counts, talk/hold/wrap averages, AHT, PCS response rate,
  Q1/Q2 averages, equal-response-weighted PCS average, top box and low scores.
- Exports clean detailed CSV/XLSX datasets without changing an extract.
- Runs trusted custom portable-Python jobs and read-only custom SQL jobs.
- Lets the user choose the source group, report pack and date period from the
  daily menu.
- Keeps Verint Forecast isolated from attendance, absence, corrections, and
  payroll logic.

## Windows quick start

Download `WFMHub-Portable-v0.3.0-win-x64.zip` from GitHub Releases—not GitHub's
automatic “Source code” ZIP. Choose **Extract All**, keep the complete `WFMHub`
folder together, double-click `SETUP.cmd` once, then use `WFMHub.cmd` each day.

`SETUP.cmd` first tests the bundled runtime, SQLite transactions/backups, and
Excel write/read support. The portable package uses SQLite built into the
official CPython runtime and adds no third-party database extension or app-local
MSVC package.

Read [`docs/BEGINNER_GUIDE.md`](docs/BEGINNER_GUIDE.md) for click-by-click help.

The release includes `templates\FTE Count.xlsx`, the standard blank roster.
Copy it to the configured `FTE` folder and maintain agent rows on its `Agent`
sheet. WFMHub also recognizes safe renamed/offset roster tables and refuses
ambiguous workbooks instead of guessing.

Generated workbooks are independent report packs:

- `output\operations`: Attendance, GAPS and RTA.
- `output\intraday`: APBE/APFR actuals and separate Verint forecast.
- `output\quality_pcs`: agent call performance, PCS and survey comments.
- `output\data_exports`: clean detailed CSV/XLSX data selected from the menu.
- `output\custom`: results from Custom Lab jobs.

## Source boundaries

| Source | Used for |
|---|---|
| FTE Agent sheet | Agent scope and organisation context |
| Verint Schedules & Activities | Planned shifts, activities, absence, start/end time |
| Storm LILO | Actual first login and last logout |
| Storm Agent Status | Actual status timeline, conformance, and RTA |
| Verint Forecast | Forecast and staffing requirements only |
| Storm APBE/APFR | Queue actuals only |
| Storm Call by Call | Agent call performance and PCS survey results |

Every dated source uses the date contained in each row. A filename may contain
one date, a start/end range or no date. For a blank LILO row in a multi-day
file, include a `Date`, `Business Date`, `Extract Date` or `Report Date` column;
without a row date the hub rejects that row instead of inventing a no-show day.

## Custom Lab

Copy `custom\jobs\_paste_your_python_here.py`, rename the copy without the
leading underscore and paste custom code inside `run(ctx)`. Custom jobs receive
the selected start/end dates and a read-only hub query API. Read-only SQL jobs
work the same way under `custom\sql`. Run only Python you trust; a Python script
is executable code. The work computer never runs `pip`.

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
