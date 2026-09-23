# Configuration

Setup copies each `config/default_*` file to its editable operational name.
Edit the operational copy. Defaults are release reference and may be updated by
future releases.

## `wfmhub.toml`

- `[paths]`: source root, SQLite, reports, optional dedicated PCS workbook,
  feeds, logs, backups and catalogs.
- `[sources]`: folders/files under the source root.
- `[period]`: optional saved start/end in `YYYY-MM-DD`.
- `[rules]`: runtime tolerances and status coverage thresholds.
- `[modules]`: optional source/model domains.
- `[pcs_tracker]`: retained PCS history months.
- `[report]`: safety row limits.

The database path should remain `_system/database/wfm.sqlite3` and must be on a
local writable disk.

`paths.pcs_workbook = ""` keeps PCS under `Reports`. Set it through the PCS
menu to the local `.xlsx` path in a synced WFM SharePoint library. This does
not relocate the database, feeds, or other reports. A SharePoint web URL is
not a local Excel file path. For PCS Power Query, put the six current CSVs in
a fixed SharePoint folder and connect using its locally synced File Explorer
path; see [the PCS setup guide](PCS_POWER_QUERY_SETUP.md).

## Mappings

`queue_mapping.csv` has five columns:

```text
mapping_type,source_system,source_value,service_scope,designation
```

`forecast_file` maps the export filename prefix. `scope_rollup` defines a
management rollup. `queue` maps an exact source queue. Do not add fuzzy queue
rules when an exact queue is known.

`service_profiles.toml` owns each Flash label, Management LOB, source scope,
workforce LOB, exact queue allowlist, subgroups, hours, KPI method IDs, and
effective dates.

`capacity_mapping.csv` maps schedule and forecast dimensions:

```text
management_lob,planning_group,workforce_lob,forecast_file_prefix,
forecast_staff_type,schedule_assignment,staff_type
```

Multiple rows may belong to one Management LOB. RSA BE intentionally separates
FR and VL planning groups even though its service headline is combined.

## Formula catalogs

`metric_catalog.toml` owns calculation expressions, component names, target,
direction, aggregation, source model, scope, effective dates, priority, and
minimum sample. `wfm_rules.toml` owns classification/evidence policy.
`analytics_rules.toml` owns deterministic finding thresholds.

After a change, run:

```text
wfmhub rules validate
wfmhub rules test
wfmhub rules catalog
```

The generated governance workbook records formulas and hashes for review.

## Safe change process

1. Copy the current operational config outside the Hub as a review backup.
2. Change one business decision at a time.
3. Add/effect-date rather than silently rewriting historical meaning.
4. Validate configuration.
5. Refresh a small known period.
6. reconcile raw numerator/denominator components against the business source.
7. only then use the change for current production output.
