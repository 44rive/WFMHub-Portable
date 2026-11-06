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

## Standalone Mr.Propre build

Download `WFMHub-Portable-v1.1.2-win-x64.zip` from release `v1.1.2`, then run:

```bash
python tools/mr_propre/build_portable.py --runtime-archive /path/to/WFMHub-Portable-v1.1.2-win-x64.zip --output build/mr-propre
python -m unittest discover -s tools/mr_propre/tests -v
```

The builder verifies the source ZIP and Excel wheel hashes, packages the embedded
runtime with the standalone CLI, and writes a deterministic ZIP and SHA-256 file.
The Mr.Propre Windows workflow extracts that ZIP and runs synthetic tests with
its embedded interpreter, including XLSX loading and the CMD launcher in paths
containing spaces. See `tools/mr_propre/README.md` for its separate ID-membership
contract. This utility has no database or report dependency.

To publish the Windows-tested ZIP, dispatch the Mr.Propre workflow on branch
`mr-propre-v1.0.1` with `publish=true`. Publication requires the Windows job to
pass, verifies the artifact checksum, uploads the portable ZIP and source to a
draft release, then publishes it without changing WFMHub's latest release.

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
