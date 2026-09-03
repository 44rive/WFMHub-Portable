# Developer guide

## Local checks

```bash
python3 -m pip install -e .
python3 -m wfmhub --home . doctor
python3 -m unittest discover -s tests -v
python3 -m compileall -q src tests packaging
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

Business KPI methods belong only in `config/default_metrics.toml` and are
evaluated by `metrics.py` from additive components registered in `semantic.py`.
Activity/evidence classification belongs in `config/default_rules.toml`. Do
not add a second hidden ratio in model SQL or Excel. Every result must carry the
catalog and rule hashes when applicable.

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
python3 packaging/windows/build_portable.py --version 0.10.1 --python-version 3.13.7
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
- agent-day, absence-event, absence agent-day, PCS agent-day, call-key,
  forecast-hour, and service-interval grains are correct;
- multi-day sources use row dates and ambiguous blank LILO rows are rejected;
- every no-show passes the full source/roster/boundary gate;
- all Verint activity labels are mapped or appear in an explicit review output;
- service availability equals answered/offered and is never an agent metric;
- aggregate rates use summed numerators and denominators;
- the applied rule version/hash matches generated reports and exports;
- correction minutes equal displayed interval endpoints;
- failed model/action operations restore prior state;
- unchanged refresh creates no raw/source-registry growth;
- ZIP checksum and native inventory match the builder result.

The historical DuckDB CLI probe remains in the repository only as evidence of
the rejected compatibility path. It is not part of v0.3.0 or its runtime.

## FTE template and report packs

Rebuild the blank public roster template deterministically:

```bash
python3 tools/build_fte_template.py
```

Create a local populated standard copy without changing the source workbook:

```bash
python3 tools/build_fte_template.py --source /path/to/current/FTE.xlsx \
  --output output/setup/FTE\ Count.xlsx
```

Report-pack keys and destinations live in `report_packs.py` and
`[report_packs]`, separate from integer row limits in `[report]`.
`reports.build_report()` remains the Daily Operations compatibility API. Register
builders through `build_report_pack`; do not import them directly in the CLI.
The menu exposes only `pcs`, `bonus`, `service`, `staffing`, `attendance`,
`corrections`, and `absence`. `operations` and `quality_pcs` remain compatibility
API keys, not current product contracts.

All current products use `DecisionWorkbook`. A builder must create the exact
ordered sheets declared in `default_reports.toml`, keep `_AUDIT` hidden, and
write only report-ready grains to `output/model_data/<pack>`. Do not add raw
extract tables to workbooks or model packages.

`template-init` may create a styled `.xlsx` starter. After a user adds Excel
Data Model, PivotTable, or slicer parts, Python must never open or rewrite that
master. Daily jobs update only the stable compact CSV package; Excel performs
Refresh All. Local masters are Git-ignored.

Call-by-call uses a stable deterministic leg key. Full-history extracts may
overlap, so `core.clean_call_leg` chooses the newest active row at that key.
PCS primary/participation questions, participation status, discrete allowed
scores, diagnostic survey mode, thresholds, and zero-denominator behavior live
in the central rulebook and tests. Official score and participation counters
are inbound-only and higher grains divide summed counters.

`exports.py` streams detailed clean data to CSV/XLSX only on explicit request.
`custom_jobs.py` exposes a query-only context, but Python jobs remain trusted
code rather than a sandbox. The portable build copies underscore templates from
`custom`; runnable user copies must never be added to a public release.

`analytics.py` is deterministic and must remain evidence-backed. Add a finding
type only when it can name the governed metric/dataset, scope, period, method,
and evidence filter. Do not add a model/API client or external database upload
path to the runtime. The Copilot file under `prompts` is a manual handoff only.

Bonus Matrix imports are content-hashed and replace the active version for one
period inside a savepoint. Keep the source-cached result as a reconciliation
control, and keep Scenario Payout distinct from Released Payout. Service-report
scope and queue grouping belong in `service_profiles.toml`; queue membership
continues to belong in `queue_mapping.csv`.
