# WFMHub BO Live — TOYOTA

An independent local mini app for ABS task snapshots. It has its own SQLite
database and does not import or modify the main WFMHub database, reports,
employee records, or calculations. Version 0.1.0.

## Start on Windows

1. Extract the complete **WFMHub-BO-Live-v0.1.0-win-x64.zip** to a normal local
   writable folder outside OneDrive/SharePoint. No Python installation is needed.
2. Open `DEMO_BO_LIVE.cmd` to explore invented tasks. Demo data uses a separate
   temporary folder, is labeled throughout, and disappears when the app stops.
3. Open `START_BO_LIVE.cmd` for the live workspace at `http://127.0.0.1:8787`.
4. In **Setup & connection**, set the ABS time zone, exact TOYOTA group names,
   assignment interpretation, and capture thresholds. The initial scope is
   only `(TOYOTA) Agendas à Traiter`; add other exact groups after verifying
   their full names. Source time zone is fixed after the first complete capture.
5. Follow [PAD_SETUP.md](PAD_SETUP.md) to connect the ABS read path. Keep the
   terminal window open: it runs the dashboard and watches the inbox.

The source checkout can also use installed Python 3.10+ with available IANA
time-zone data. Windows Python may require `tzdata`; the portable ZIP includes
the pinned package. From an existing full WFMHub folder, this mini app can reuse
`../../_system/runtime/python.exe` without sharing the WFMHub database.

```text
python miniapps/bo_live/app.py serve --demo
python miniapps/bo_live/app.py serve --no-browser
python miniapps/bo_live/app.py serve --home /path/to/bo-live-data
```

## What it does

- Shows task counts, priority, scheduled-action timing, assignment, and language.
- Filters/searches task rows and preserves repeated case references.
- Records complete captures and failed collection attempts independently.
- Keeps last-good evidence with stale/failed/missing states and capture times.
- Flags partial configured scope; sums describe separate capture observations,
  not an exact simultaneous count across task groups.
- Watches an inbox and supports manual JSON import for diagnosis.
- Runs only on loopback with guarded browser mutations and no external assets.

`Scheduled for` is the time to action a task. Timing is evaluated at that
group's capture finish: at/before finish is **past action time**, then the
configured upcoming window, later work, and unscheduled work. This is not an
SLA, case age, employee productivity, completion, or throughput measure. A
case number alone cannot identify an individual task across snapshots.

An unassigned KPI requires confirmation of blank assignment semantics in
Setup. Unknown source priorities remain visible after configured priorities.
Genuinely empty complete captures produce zero; failures never produce zero
by replacing the last good rows.

## Data and backup

Default source-checkout data lives under `miniapps/bo_live/_local/`. In the
Windows ZIP it lives under `bo_live/_local/` beside the mini app's code. The
actual folder and inbox path appear in Setup and the terminal.

```text
_local/
  settings.json          # Local scope and calculation configuration
  bo-live.sqlite3        # Independent durable BO evidence
  inbox/                 # Atomically published captures from PAD
  processed/             # Captures already handled
  rejected/              # Invalid captures retained for review
```

Keep this entire folder on local disk. To back up or upgrade, stop the app and
PAD, copy the complete data folder, extract the next app to a new folder, and
start it with `--home` pointing to the existing data folder. Do not replace the
database with files from an app ZIP. Capture history is retained until an
explicit retention feature is added; monitor disk usage.

The app must remain open for live updates. PAD needs its configured Windows
session and execution licensing. Private ABS navigation/selectors must be
recorded on the actual workstation; the included publisher is executable, but
this distribution does not claim to contain a tested ABS extraction flow.

## Capture contract v1

Each JSON file is one group observation. Metadata fields:

| Field | Contract |
|---|---|
| `schema_version` | Integer `1` |
| `capture_id` | Unique immutable ID: letters/digits/dot/underscore/hyphen, maximum 100 characters |
| `task_group` | Exact configured TOYOTA group name |
| `started_at`, `finished_at` | ISO date/time with UTC offsets; finish is not before start |
| `status` | `complete`, `incomplete`, or `failed` |
| `source_timezone` | For complete captures, exact configured IANA zone |
| `source_total` | For complete captures, verified source total equal to extracted row count |
| `scope_verified`, `total_verified` | Both `true` for complete captures |
| `coverage` | `full_table` or `absolute_rows` for complete captures |
| `rows` | Complete list of task rows; failed/incomplete captures use an empty list |
| `error` | Required short diagnostic for failed/incomplete captures |

Task fields: `row_index` (every integer 1 through total), `case_number` (text),
`task_type`, `priority`, `language`, `scheduled_for` (full ISO timestamp or
`DD/MM/YYYY HH:MM[:SS]`, or null/blank when genuinely unscheduled), and
`assigned_to` (text, blank allowed after successful extraction).

Row indexes identify occurrences inside one capture only. Complete structured
table reads can assign sequential indexes. Overlapping viewport reads require
verified absolute positions before export. Ambiguous/nonexistent local DST
times require an explicit UTC offset. Partially read values and truncated dates
are invalid; missing assignment/language fields are not silently filled in.

Re-importing identical evidence with the same capture ID is idempotent.
Reusing an ID for different evidence is rejected. A later failed attempt does
not replace good rows; an older successful observation cannot supersede a
newer observation. The app can validate the declared scope and counts, but only
the actual PAD extraction can establish that the declared traversal is true.

## Developer checks and packaging

```text
python -m unittest discover -s tests -p test_bo_live.py -v
node --check miniapps/bo_live/web/app.js
python packaging/windows/build_bo_live.py
```

The standalone builder uses the same SHA-256-reviewed official CPython archive
as WFMHub and includes only the hash-pinned pure-Python timezone wheel. An
explicit source allowlist excludes local data and settings. The source-only
variant is `python packaging/windows/build_bo_live.py --source-only`.
