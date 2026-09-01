# Governed AI analysis snapshot

The AI analysis snapshot is a small, read-only SQLite bundle for an external
analysis service. It is not a copy of the hub database. It contains only fixed,
curated aggregate contracts and a JSON manifest.

WFMHub remains the calculation authority. An AI tool may explain, compare and
visualize these results, but it must not recalculate official WFM measures or
write an answer back into attendance, Verint, payroll or the operational hub.

## Create a bundle

Run this after the normal WFMHub refresh has completed:

```text
WFMHub.cmd analysis-snapshot --start 2026-08-01 --end 2026-08-31
```

To select a new destination folder:

```text
WFMHub.cmd analysis-snapshot --start 2026-08-01 --end 2026-08-31 --output D:\WFM-AI\August
```

The start and end dates are mandatory. The output folder must not already
exist, which prevents a previous analysis bundle from being overwritten. When
`--output` is omitted, WFMHub creates a timestamped folder under
`output\ai_analysis`.

This command opens the operational SQLite database in read-only/query-only
mode. It does not ingest files, rebuild models, edit source extracts or accept
custom SQL.

## Bundle contents

```text
wfmhub_analysis_...\
├── wfmhub_analysis.sqlite3
└── manifest.json
```

The SQLite file contains only:

| Table | Grain |
|---|---|
| `source_health` | One configured source family |
| `daily_service_lob` | Date/interval/LOB/language |
| `daily_staffing_gaps` | Date/15 minutes/LOB/language |
| `pcs_team_day` | Date/team leader/LOB/language |

Raw extracts, agent-status intervals, calls, schedules, local source paths and
the operational database are not included. `source_health` is current health;
the other three datasets are filtered to the requested dates.

## Provenance and validation

`manifest.json` is the sole metadata contract and records:

- requested start and end dates;
- UTC generation time;
- row count and grain for every dataset;
- latest successful refresh ID, when available;
- latest model-run ID, applied rule version/hash and queue-mapping hash;
- current configuration hashes and whether they match the latest model run;
- SHA-256 of the SQLite snapshot;
- explicit confirmation that raw extracts and arbitrary SQL are absent.

WFMHub completes a SQLite integrity check before publishing the bundle and
marks the database and manifest files read-only. The SQLite table allowlist is
exactly the four tables above, so an analysis publisher can reject anything
unexpected.

If `configuration_matches_model` is `false`, refresh WFMHub before analysis.
This means a rules or queue-mapping file changed after the current marts were
built.

## Safe handoff

Give the analysis service only the completed bundle folder. Do not mount the
WFMHub database, `extracts`, `input` or `config` folders into the AI service.
Open the snapshot using SQLite read-only mode. For example, a Python service
should use a URI ending in `?mode=ro`.

The snapshot is an analysis input, not a payroll or Verint correction output.
Any AI-generated conclusion remains advisory until a person validates it
against the governed datasets and provenance.
