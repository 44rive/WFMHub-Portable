# Queue and status reference change — 2026-09-10

This record captures the Operations-supplied Queue, AUX, and Agent Status
references without inferring meaning from queue names.

## Effective Flash sets

| Flash | Previous shipped set | New exact set | Change |
|---|---:|---:|---|
| RSA NL | 30 | 30 | no membership change |
| RSA BE | 42 | 43 | add 4, remove 3 |
| Ford NL | 6 | 6 | no membership change |
| OEM | 1 | 3 | add Toyota/Lexus and Chery |

RSA BE additions:

- `APBN_BRU_MOBILITY_Ford_Assistance_DE`
- `APBN_BRU_MOBILITY_Ford_Assistance_VL`
- `APBN_BRU_MOBILITY_Ford_Dealers_VL`
- `APBN_LUX_MOBILITY_Ford_Assistance_DE`

RSA BE removals:

- `APCH_ZRH_RSA_Ford_Assistance_FR`
- `APFR_PAR_RSA_CSTRUCTR_TOYOTA-LEXUS_FR`
- `APFR_PAR_RSA_CHERY_ASSISTANCE_FR`

The two APFR removals are now OEM queues. The four RSA BE additions are also in
Ford NL. That overlap is intentional: their primary model scope is `Ford NL`,
and both exact Flash allowlists count them.

Nine obsolete shipped queue-map rows not present in the supplied authority were
retired: the two prefix-less BRU Bike aliases, five Police OBU queues, Various
Assist NL, and APCH Ford Assistance FR. A local row with a custom scope is not
removed by the guarded migration.

## Recalculation on the supplied clean Call-by-Call example

Using `TOLEARN/calls_2026-09-05_to_2026-09-05_015727_309989.csv`, inbound queue
entries only, the same configured 5/30-second Storm equation, and no change to
the arithmetic:

| Flash | Previous Entered | New Entered | Previous Handled | New Handled | Previous Handled in SL | New Handled in SL | Previous TSL | New TSL |
|---|---:|---:|---:|---:|---:|---:|---:|---:|
| RSA NL | 594 | 594 | 583 | 583 | 524 | 524 | 88.22% | 88.22% |
| RSA BE | 353 | 257 | 337 | 238 | 308 | 217 | 87.50% | 84.77% |
| Ford NL | 125 | 125 | 121 | 121 | 118 | 118 | 94.40% | 94.40% |
| OEM | 47 | 205 | 46 | 204 | 42 | 190 | 89.36% | 92.68% |

New OEM detail on that same file:

| OEM sub-LOB | Entered | Handled | Handled in SL | TSL |
|---|---:|---:|---:|---:|
| Ford | 47 | 46 | 42 | 89.36% |
| Toyota/Lexus | 154 | 154 | 144 | 93.51% |
| Chery | 4 | 4 | 4 | 100.00% |

These figures validate the effect of membership only. A live Storm comparison
must use the same source snapshot and cutoff time.

## Status and AUX reference

The exact supplied labels, AUX classifications, and Status qualifications are
stored as `[[status_rules]]` in `config/default_rules.toml`. The clean Agent
Status category uses the configured exact label first and the conservative
legacy fallback only for an unlisted label. Changing this rulebook fingerprint
causes unchanged Agent Status extracts to be reclassified on the next update.

