# Metric catalog guide — step by step

You do not need to edit Python, SQL, Excel formulas, or extracts to change a
supported KPI method.

## The simple picture

Think of the hub as a kitchen:

1. Extract parsers wash and label the ingredients.
2. Domain models produce countable ingredients such as `offered`, `answered`,
   `pcs_score_sum`, or `planned_net_minutes`.
3. `metric_catalog.toml` is the recipe card.
4. Python follows the card and stores the result in SQLite.
5. Reports display that stored result. Excel does not cook it again.

## Files and responsibilities

| User file | Change it when |
|---|---|
| `config\metric_catalog.toml` | Formula, target, unit, sample, scope, aggregation or effective date changes |
| `config\wfm_rules.toml` | Evidence/activity classification or PCS input-parsing policy changes |
| `config\analytics_rules.toml` | A warning should be more or less sensitive |
| `config\report_catalog.toml` | A title/finding domain changes, or code adds a reviewed sheet contract |
| `config\queue_mapping.csv` | A queue/file name or roll-up changes |

Never edit a file beginning with `default_`. On first setup, WFMHub copies each
factory default to the user file you own.

## Read one metric method

```toml
[[metrics]]
id = "service_level"
method = "adjusted_20"
label = "Adjusted service level 20s"
domain = "service"
source_model = "service_interval"
grain = "queue interval"
unit = "percent"
aggregation = "ratio_of_sums"
numerator = "answered_within_target"
denominator = "offered - short_abandoned"
sample = "offered - short_abandoned"
target = 0.80
direction = "higher_is_better"
minimum_sample = 1
effective_from = 2026-01-01
priority = 0
finding_dimensions = ["source_system", "lob", "language"]
```

- `id`: stable business name used by reports.
- `method`: name of this exact recipe.
- `source_model`: trusted component provider. Validation rejects unknown
  components.
- `numerator` and `denominator`: formula parts stored separately for audit and
  correct roll-up.
- `aggregation = "ratio_of_sums"`: at team/month level, sum numerators, sum
  denominators, then divide. Never average percentages.
- `sample`: denominator used for low-sample checks.
- `target` and `direction`: decide `ON_TARGET`, `BELOW_TARGET`, or
  `ABOVE_TARGET`.
- `effective_from` / optional `effective_to`: dates on which the method is valid.
- `priority`: chooses a specific scoped method over a broad default.

Percentages are decimals: `0.80` means 80%.

## Change a target only

1. Copy `metric_catalog.toml` to a dated backup.
2. Increase `[catalog].version`.
3. Change `target = 0.80` to the approved value, for example `0.85`.
4. Save.
5. Run `WFMHub.cmd rules validate`.
6. Run `WFMHub.cmd rules test`.
7. Refresh a known short period and compare `METHODS`, `FINDINGS`, and
   `PROVENANCE` with the previous workbook.

## Change a method from a future date

Do not rewrite historical policy silently. Close the old method and add a new
one:

```toml
# Old method
effective_from = 2026-01-01
effective_to = 2026-09-30

# New [[metrics]] block with the same id
method = "adjusted_20_v2"
effective_from = 2026-10-01
```

Old dates keep the old recipe; new dates use the new recipe. Avoid overlapping
same-priority methods.

## Add a scoped override

Keep a broad default at priority `0`, then add a more specific block:

```toml
scope = { lob_contains = ["RSA"] }
priority = 10
```

Allowed scope fields are ordinary dimensions exposed by the source model,
including `source_system`, `lob`, `language`, `queue`, `business_partner`, and
their supported `_contains` variants. If two matching methods have the same
highest priority, validation/model refresh stops rather than guessing.

## Safe formula language

Allowed operators are `+`, `-`, `*`, `/` and numeric comparisons. Allowed
functions are `min`, `max`, `coalesce`, `nullif`, `ifelse`, `abs`, and `round`.
Python imports, file access, attributes, loops, and arbitrary function calls are
rejected.

A zero denominator produces blank/`NO_DATA`, not a false zero percent.

## Governance commands

From a command prompt in the WFMHub folder:

```text
WFMHub.cmd rules validate
WFMHub.cmd rules test
WFMHub.cmd rules explain service_level
WFMHub.cmd rules diff --against C:\path\metric_catalog_before.toml
WFMHub.cmd rules catalog
```

- `validate` checks every catalog, source component, date, mapping, and report
  contract.
- `test` evaluates every configured method with representative components.
- `explain` prints every effective method for one metric.
- `diff` lists added, removed, and changed methods compared with a backup.
- `catalog` creates a management-readable governance workbook.

## Add a genuinely new KPI

A new formula can be configuration-only when all required additive components
already exist in the named `source_model`. Add the method, validate, and then
add the metric to a dataset/report contract.

If a required component does not exist, code must first expose that trusted
counter in `semantic.py` and a migration/test may be required. Do not hide a
new calculation in report SQL or an Excel cell.

## Final release check

Before sending a management workbook:

1. `SOURCE_HEALTH` has no unexplained missing/stale source.
2. `DATA_QUALITY` has no unexplained critical issue.
3. `METHODS` contains the expected method and effective period.
4. `PROVENANCE` hashes match the refresh you intended.
5. Low samples are disclosed.
6. Historical totals reconcile from summed numerator and denominator counters.
