# Architecture

## Data flow

```text
unchanged extracts
      │
      ▼
source discovery + parser contracts
      │  fingerprints, source version, rejected-row evidence
      ▼
raw_* SQLite tables
      │
      ▼
clean employee/call layer + governed mappings
      │
      ▼
mart_* attendance / staffing / service / PCS / absence / quality models
      │
      ├── Reports/*.xlsx
      ├── Feed/**/*.csv
      └── localhost Manager Workbench
```

The Python process is the only writer. SQLite uses WAL, a process lock, full
synchronous commits, integrity checks, and verified backups. The browser opens
only on loopback and reads governed marts through `DashboardData`.

## Layers

- `meta_*`: source, refresh, configuration-application, and quality evidence.
- `raw_*`: typed records tied to immutable source-file fingerprints.
- `core_*`: reusable clean dimensions/views.
- `mart_*`: business-ready facts used by reports and the web app.

Application SQL uses logical names such as `mart.attendance_agent_day`.
`database.py` maps them to portable SQLite names such as
`mart_attendance_agent_day`.

## Idempotence

Source files are identified by family, path, metadata, and SHA-256. Unchanged
files are skipped. Replacements deactivate the prior active version only after
the new version parses successfully. Derived marts are rebuilt inside one
savepoint, so a failed model refresh cannot leave a half-refreshed layer.

## Release state

The release archive contains program code, default configuration, templates,
documentation, and the embedded runtime. Durable user state lives under
`_system/` and in editable config files and is copied by `UPGRADE.cmd`.
`Reports/` and `Feed/` are outputs, not model authority.

## Performance

Full refresh work consists of file discovery, changed-file hashing/parsing,
employee scoping, mart rebuilds, feed publication, and optional workbook
generation. Focused PCS refresh scans only FTE and Call-by-Call and rebuilds
only PCS data. Unchanged large files are metadata-skipped when safe.

Use the log's per-stage timings to diagnose slowness. Keep the live database
and program outside sync folders; copy completed workbooks to collaboration
storage afterward.
