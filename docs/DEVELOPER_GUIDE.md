# Developer guide

## Local checks

```bash
python -m pip install -e .
python -m wfmhub --home . doctor
python -m unittest discover -s tests -v
python -m compileall -q src tests packaging
```

Use an isolated `--home` for real-source acceptance so local user configuration
and the preserved v0.1 database are not modified.

## Add a source or feature

1. Add its default source path/module flag when needed.
2. Add a strict parser in `ingestion.py`.
3. Decide whether the feed is agent-scoped or queue-scoped.
4. Add raw/core/mart tables through a new numbered migration. Never edit a
   released migration.
5. Carry immutable file lineage and define the exact business grain.
6. Materialize only the relevant model in `models.py`.
7. Add curated output through `report_packs.py`; add explicit clean export
   datasets separately when detailed rows are needed.
8. Test schema drift, scope, idempotency, rollback, period filtering, and the
   business rule.
9. Update architecture and beginner documentation.

Forecast fields must remain forecast-only. Do not join forecast into employee
attendance, absence, correction, or payroll models.

## SQLite rules

- Use logical `meta.`, `raw.`, `core.`, and `mart.` names through
  `DatabaseConnection`; the facade maps them to SQLite prefixes.
- Use `?` parameters and standard SQL supported by the shipped SQLite.
- Keep source-file changes transactional and model refreshes inside one
  savepoint.
- Use the online backup API; never copy a live WAL database with `shutil`.
- Preserve leading-zero Agent IDs as text.
- Size multi-value inserts from SQLite's runtime host-variable limit.
- Add indexes only for measured query paths; marts are rebuilt.

## Windows portable build

```bash
python packaging/windows/build_portable.py --version 0.3.2 --python-version 3.13.7
```

The builder always deletes and recreates stage and wheelhouse. It:

- verifies the reviewed official CPython embeddable SHA-256;
- downloads all explicit dependencies with `--require-hashes --no-deps`;
- accepts only `none-any` pure-Python wheels;
- rejects native content in wheels;
- verifies that the final native inventory exactly equals the CPython ZIP;
- rejects DuckDB/MSVC-runtime paths and user config/database files;
- writes `RUNTIME_MANIFEST.sha256`, the ZIP, and the ZIP SHA-256.

An unknown Python version intentionally fails until its official archive hash
is reviewed and registered. Do not upgrade the runtime casually on a machine
with strict Application Control.

## Acceptance gates

At minimum, verify:

- unit/end-to-end suite passes;
- all discovered real files load or produce an understood source error;
- source hashes/size/mtime are unchanged;
- `PRAGMA quick_check` and `foreign_key_check` pass;
- one active version exists per source path;
- agent-day, PCS agent-day, call-key, forecast-hour, and actual-interval grains
  are unique;
- multi-day sources use row dates and ambiguous blank LILO rows are rejected;
- every no-show passes the full source/roster/boundary gate;
- correction minutes equal displayed interval endpoints;
- failed model/action operations restore prior state;
- unchanged refresh creates no raw/source-registry growth;
- ZIP checksum and native inventory match the builder result.

The historical DuckDB CLI probe remains in the repository only as evidence of
the rejected compatibility path. It is not part of v0.3.0 or its runtime.

## FTE template and report packs

Rebuild the blank public roster template deterministically:

```bash
python tools/build_fte_template.py
```

Create a local populated standard copy without changing the source workbook:

```bash
python tools/build_fte_template.py --source /path/to/current/FTE.xlsx \
  --output output/setup/FTE\ Count.xlsx
```

Report-pack keys and destinations live in `report_packs.py` and
`[report_packs]`, separate from integer row limits in `[report]`.
`reports.build_report()` remains the Operations compatibility API. Register
builders through `build_report_pack`; do not import them directly in the CLI.

Call-by-call uses a stable deterministic leg key. Full-history extracts may
overlap, so `core.clean_call_leg` chooses the newest active row at that key.
PCS numerator, denominator, eligible survey mode, scored questions, thresholds
and zero-denominator behavior live in configuration/model code and tests.

`exports.py` streams detailed clean data to CSV/XLSX only on explicit request.
`custom_jobs.py` exposes a query-only context, but Python jobs remain trusted
code rather than a sandbox. The portable build copies underscore templates from
`custom`; runnable user copies must never be added to a public release.
