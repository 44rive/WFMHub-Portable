# Reference and Power BI audit — 2026-09-12

This note records the release audit against the files in `attachments`. The
canonical executable defaults remain the files under `config/default_*`.

## REF.xlsx reconciliation

- `AUX REF`: 24 populated labels; all 24 exist in `default_rules.toml` with the
  same AUX classification.
- `Status REF`: 27 populated labels; all 27 exist with the same Qualification 1
  and Qualification 2.
- `Queues Ref`: 31 populated queues; all 31 exist in
  `default_queue_mapping.csv`.
- 23 rows marked `SL Related = Y`; all 23 are in an active Flash allowlist.
- No row marked non-SL is present in an active Flash allowlist.

The complete active Flash allowlists remain exact and deliberately separate:

| Flash | Exact queues |
|---|---:|
| RSA NL | 23 |
| RSA BE | 43 |
| Ford NL | 6 |
| OEM | 3 |

Do not infer or move queues from their names. The four queues intentionally
shared by RSA BE and Ford NL retain dual-Flash membership while their primary
semantic mapping remains configured once.

## New source files

`FORD_FR_09-2026.txt`, `FORD_NL_09-2026.txt`, `RSA_BE_09-2026.txt` and
`RSA_NL_09-2026.txt` are Verint forecasts at 15-minute grain. The supplied
columns contain volume forecast only. They do not contain
`Full Time Equivalents (Absolute Req)`, so Power BI Required FTE, Net Gap and
Coverage must remain blank for those intervals. WFMHub does not estimate a
requirement from volume.

`VERINT_ALL_09-2026_01092026_30092026_StartEndTimes.txt` is the preferred
schedule boundary source. The matching `Activities.txt` is final post-day
absence/shrinkage authority and the reconciliation source for residual
Attendance Review gaps. It never creates observed presence.

## Governed Power BI calculations

- Service Level: sum of answered within target divided by sum of offered minus
  abandoned within target, using only each profile's exact queue allowlist.
- Service availability: answered divided by offered.
- Forecast accuracy: ratio-of-sums/WAPE at comparable 15-minute grain.
- Staffing: StartEndTimes scheduled capacity minus governed PTO/Away; actual
  presence comes from Agent Status with LILO fallback.
- Required staffing: only supplied Verint absolute required FTE.
- Final absence/shrinkage: mapped Verint Activities clipped to the schedule.
- Schedule integrity: completed published shifts versus sustained Agent Status
  presence, with LILO fallback, PTO/Away exclusion and configured tolerances;
  it is not adherence.
- Active population: Active FTE plus Leavers only through their populated leave
  date. Client/Agent ID remains text.

Power BI receives additive governed CSVs and divides summed components. It does
not reclassify raw source rows or average stored percentages.
