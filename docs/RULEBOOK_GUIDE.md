# Editing WFMHub calculations safely

This guide is intentionally step-by-step. You do not need to know Python or SQL.

## Which file to edit

Edit:

```text
config\wfm_rules.toml
```

Do not edit `config\default_rules.toml`. That file is the factory copy used for
new installations.

Before editing, copy `wfm_rules.toml` somewhere safe. WFMHub also keeps the raw
extracts untouched, so changing a rule never changes a Verint or Storm file.

## The safe routine

1. Close WFMHub and any generated reports.
2. Open `config\wfm_rules.toml` in Notepad.
3. Change one small thing.
4. Increase `rulebook.version`. Example: `2026.08.1` becomes `2026.08.2`.
5. Save the file.
6. Open `WFMHub.cmd`.
7. Choose **Validate rules and build KPI catalog**.
8. If validation is green, refresh a short test period first.
9. Compare the new report with the previous report.
10. Refresh the full month only after the comparison looks correct.

If validation fails, WFMHub does not refresh the marts. The error names the
section or formula that needs correction.

## Formula examples

Service availability is defined as:

```toml
formula = "answered / nullif(offered, 0)"
```

This is availability of the service. It is not time that an agent was Available,
logged in, conformant, or adherent.

The two service-level profiles are:

```toml
# Strict/gross
formula = "answered_within_target / nullif(offered, 0)"

# Excluding short abandons
formula = "answered_within_target / nullif(offered - short_abandoned, 0)"
```

`nullif(offered, 0)` prevents division by zero. A zero denominator produces a
blank KPI instead of a fake zero percent.

Allowed formula tools are ordinary `+`, `-`, `*`, `/`, comparisons such as
`>`, `>=`, `<`, `<=`, `==`, and `!=`, and the small safe list:

- `min(...)`
- `max(...)`
- `coalesce(...)`
- `nullif(a, b)`
- `ifelse(condition, value_if_true, value_if_false)`
- `abs(...)`
- `round(...)`

Python imports, file access, attributes, loops, and arbitrary functions are not
allowed.

## Changing an activity rule

Rules are checked from top to bottom. The first match wins.

Example:

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

The flags mean:

- `absence`: include clipped minutes in payroll absence.
- `vacation`: keep the minutes separately as vacation.
- `unpaid`: keep the minutes separately as unpaid.
- `shrinkage`: include the minutes in shrinkage.
- `working`: treat the interval as productive planned work.
- `planned`: the activity was planned in Verint.

An interval can be both absence and unpaid, or both absence and shrinkage. The
daily model unions overlapping intervals before totaling, so the same minute is
not counted twice within one KPI.

## Adding a new Verint activity

1. Open the Attendance & Absence report.
2. Open `UNMAPPED`.
3. Copy the activity wording exactly.
4. Add a new `[[activity_rules]]` block above broad rules such as Production.
5. Choose the flags carefully.
6. Increase the rule version, validate, and rebuild the same dates.
7. Confirm that `UNMAPPED` is empty or contains only items still awaiting a
   business decision.

`NO_ACTIVITY` is mapped but remains a review warning. Decide whether it is true
shrinkage or a schedule defect before payroll use.

## What cannot be edited as a formula

Some operations remain tested engine primitives because a text formula is not a
safe description of them:

- clipping an activity to a scheduled shift;
- merging overlapping intervals;
- overnight shift handling;
- file and row deduplication;
- agent identity scope;
- consecutive-day spell grouping.

Their outputs and assumptions are visible in the reports and architecture guide.
Business categories and KPI equations remain centralized in the rulebook.
