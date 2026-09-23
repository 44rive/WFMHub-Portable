# Repository instructions

Before acting, read `PROJECT.md` and the task-relevant document under `docs/`.
They are the product contract. Configuration files are authoritative for exact
effective mappings and formulas.

## Required behavior

- Preserve source extracts. All ingestion is read-only.
- Never add real employee, schedule, absence, call, customer, or operational
  extract data to Git. Tests and examples must be synthetic.
- Do not infer queue membership, employee identity, or missing evidence.
- Client ID is text. FTE Active/dated-Leaver scope applies everywhere.
- Agent Status is primary attendance evidence; LILO is fallback; Verint
  Activities are the finalized post-day absence/shrinkage source.
- Service is Call-by-Call queue evidence. Forecast Volume/FTE Requirement is
  Staff Type evidence. Do not merge those grains.
- RSA BE service is combined; RSA BE FR/VL capacity remains separate.
- KPI arithmetic belongs in the effective metric/rule catalogs, not workbooks
  or web pages.
- Keep SQLite local and preserve it across releases.
- Keep generated reports and feeds out of Git.

## Change checklist

1. Trace the active caller and data contract before editing.
2. Update code, configuration, documentation, and tests together.
3. Run `python -m compileall -q src tests packaging`.
4. Run `python -m unittest discover -s tests -v`.
5. Test a database created only from `sql/migrations/001_schema.sql`.
6. Validate generated XLSX files as ZIP/XML packages.
7. Build the portable release and inspect its root; no user data or local
   configuration may be present.
