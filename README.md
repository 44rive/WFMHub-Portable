# WFMHub Portable

WFMHub Portable turns untouched WFM exports into a durable DuckDB database and
finished Excel reports. It runs on Windows x64 without administrator rights,
without installed Python, and without Power Query or Power Pivot.

## What version 1 does

- Ingests FTE, Verint Schedules & Activities, Storm LILO, Storm Agent Status,
  Verint Forecast, APBE and APFR queue actuals.
- Fingerprints every file and skips unchanged files on later refreshes.
- Keeps raw rows, clean models, report marts, source health and quality issues in
  one local DuckDB file.
- Produces one curated Excel workbook with `START_HERE`, `SUMMARY`,
  `ATTENDANCE`, `GAPS`, `RTA`, `INTRADAY`, `DATA_QUALITY` and `SOURCE_HEALTH`.
- Never writes to or changes an extract.
- Keeps Forecast isolated from Attendance and payroll corrections.

## Fast start for developers

```bash
python -m venv .venv
. .venv/bin/activate
python -m pip install -e .
python -m wfmhub --home . setup --source-root /path/to/untouched-extracts --non-interactive
python -m wfmhub --home . refresh --start 2026-08-01 --end 2026-08-31
python -m wfmhub --home . report --start 2026-08-01 --end 2026-08-31
python -m unittest discover -s tests -v
```

## Windows user

Download the `WFMHub-Portable-...-win-x64.zip` asset from Releases--not the
automatically generated GitHub "Source code" ZIP. Choose **Extract All**, keep
the whole `WFMHub` folder together, run `SETUP.cmd` once, then run `WFMHub.cmd`
every day. Read
[`docs/BEGINNER_GUIDE.md`](docs/BEGINNER_GUIDE.md) for click-by-click help.

If company Application Control blocks the Python DuckDB module, run the small
signed-CLI compatibility probe from the `duckdb-cli-probe-v1.5.5` release
before selecting a replacement backend. The probe never reads WFM data.

## Data boundaries

| Source | Used for |
|---|---|
| Verint Schedules & Activities | Planned shifts, planned events and absence |
| Storm LILO | Actual first login and last logout |
| Storm Agent Status | Actual timeline, conformance and RTA |
| FTE Agent sheet | Optional organisation context |
| Verint Forecast | Forecast and staffing requirements only |
| Storm APBE/APFR | Queue actuals only |

Verint `Data Source IDs` and Storm `Agent ID` are the operational Agent ID.
The hub does not silently join agents by name.

## Architecture

```text
Untouched extracts
      |
      v
File registry + SHA-256  -->  source health / audit trail
      |
      v
raw schema  -->  core clean models  -->  mart report tables
                                         |
                                         v
                                 Curated Excel output
```

See [`docs/ARCHITECTURE.md`](docs/ARCHITECTURE.md) for grains, rules and the
extension pattern for future modules.
