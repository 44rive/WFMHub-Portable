# Development

## Local setup

```bash
python3.13 -m venv .venv
. .venv/bin/activate
python -m pip install -e .
python -m unittest discover -s tests -v
```

The supported release runtime is Python 3.13. SQLite, openpyxl, XlsxWriter,
et_xmlfile, and tzdata versions are pinned.

## Definition of done

- active code has one caller or a documented public contract;
- source/config/model/report changes are consistent;
- tests use synthetic data;
- a new empty database migrates from `001_schema.sql`;
- an existing current database accepts the baseline migration safely;
- workbook output opens as a valid ZIP/XML package without repair records;
- launchers work with the embedded runtime;
- portable archive contains no database, user config, extracts, feeds, reports,
  or custom jobs;
- `PROJECT.md` and the relevant focused guide are current.

## Useful commands

```bash
python -m compileall -q src tests packaging
python -m unittest discover -s tests -v
python -m wfmhub --home /path/to/test-home rules validate
python packaging/windows/build_portable.py --version 1.1.2
```

## Database changes

`001_schema.sql` is the clean baseline. After 1.0.0, append small numbered,
transaction-safe migrations; never rewrite a released migration. A new table or
column needs model code, export/report consumers, tests, and documentation.

The database must remain backward-safe across program releases. Use SQLite's
online backup API before pending migrations and keep DDL idempotent where
possible.

## Workbook validation

Generated workbooks must be closed before replacement. Validate the ZIP central
directory, parse every XML part, reject external links, and open with openpyxl
in formula and data-only mode.

## Security

The web server binds only to loopback. Custom Python jobs are trusted local code;
custom SQL is read-only. Never weaken path containment for downloads or allow a
browser request to execute arbitrary commands.
