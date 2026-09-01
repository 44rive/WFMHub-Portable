# Editing attendance and absence evidence rules safely

This guide is for classification policy. KPI formulas are in
[METRIC_CATALOG_GUIDE.md](METRIC_CATALOG_GUIDE.md).

## Which file to edit

Edit `config\wfm_rules.toml`. Do not edit `default_rules.toml`; it is the
factory copy for a new installation. Before changing anything, make a dated
copy of your user file.

This rulebook controls:

- how Verint activity labels are classified;
- whether a category is planned, working, absence, vacation, unpaid, or
  shrinkage evidence;
- standard-day and gap-tolerance policy;
- PCS source parsing, valid discrete scores, the primary question, and the
  participation status.

It does not contain service, PCS, attendance, forecast, staffing, or absence
rate formulas. Those live only in `metric_catalog.toml`.

## Safe routine

1. Close generated reports.
2. Copy `wfm_rules.toml` to a dated backup.
3. Change one rule.
4. Increase `rulebook.version`.
5. Run **Validate rules and build governance catalog**.
6. Refresh one small known period and compare it with the old report.
7. Refresh the full month only after the evidence rows reconcile.

Validation happens before model rebuilding. Extracts are never changed.

## Activity rule example

Rules are checked from top to bottom; the first match wins.

```toml
[[activity_rules]]
name = "short_sickness"
category = "SICKNESS_SHORT"
patterns = ["SHORT SICKNESS", "SHORT-SICKNESS"]
match = "contains"
planned = false
working = false
absence = true
vacation = false
unpaid = false
shrinkage = true
```

The flags answer separate questions. A category may count as both absence and
shrinkage. WFMHub unions overlapping intervals before daily totals, so the same
minute is not counted twice inside one category metric.

## Add an unmapped Verint activity

1. Open Final Absenteeism and read `UNMAPPED_REVIEW`.
2. Copy the exact activity wording.
3. Add a specific rule above broad Production rules.
4. Choose the flags with the payroll/process owner.
5. Increase the version, validate, and refresh the same dates.
6. Confirm the row moved to its intended category.

Activity rules classify final Verint evidence. They never invent the original
attendance gap; observed gaps come from schedule + LILO + Agent Status.

## Engine rules that are not text configuration

Clipping to a shift, merging overlaps, overnight handling, file deduplication,
agent identity scope, same-day provisional protection, and spell grouping are
tested engine primitives. Changing those requires code and tests, not a TOML
formula.
