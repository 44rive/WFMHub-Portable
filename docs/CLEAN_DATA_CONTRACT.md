# Governed clean-data contract

The workbook layer formats and filters governed datasets; it does not invent a
second calculation. Source extracts are opened read-only. FTE defines the
effective-dated agent population: Active rows remain eligible and Leavers remain
eligible through their populated leave date.

## Source boundaries

| Source | Allowed purpose |
|---|---|
| FTE Agent/PTO/Away | In-scope roster, organisation, planned leave and long absence |
| Verint StartEndTimes | Preferred schedule start/end and assignment boundary |
| Storm Agent Status | Primary within-shift attendance evidence |
| Storm LILO | Missing-coverage fallback and first/last/blank-row control |
| Attendance Review decisions | Human category for one exact observed gap |
| Verint Forecast | Forecast and required staffing only; native 15/60-minute grain retained |
| Storm Call by Call | All mapped service actuals, Flashes, agent call performance and PCS |

APBE, APFR and APDE are retired: no directory is discovered, required, loaded or
calculated. Historical raw tables remain physically readable so upgrading is
non-destructive. Activities intervals are ignored. A parsed Activities Shift
Assignment may be used only as a visibly flagged emergency schedule boundary
when StartEndTimes is missing.

## Key datasets

| Dataset | Grain | Business use |
|---|---|---|
| `mart.attendance_agent_day` | Agent/day | Attendance callout and evidence result |
| `mart.correction_candidate` | Exact continuous gap | Attendance Review decision row |
| `mart.shift_timeline_segment` | Exact shift segment | Readable shift evidence |
| `mart.planned_time_off_segment` | Schedule-clipped interval | PTO/Away planning overlay |
| `mart.absence_event` | Reviewed or planned exact interval | Absence/shrinkage component audit |
| `mart.absence_agent_day` | Agent/day | Reviewed absence and shrinkage result |
| `mart.call_service_hour` | Date/hour/mapped queue | Interaction-deduplicated service actual |
| `mart.service_interval` | Date/hour/mapped queue | Stable semantic projection of Call-by-Call |
| `mart.forecast_interval` | Native Verint interval | Staffing and clean sharing |
| `mart.forecast_hour` | Date/hour/mapped scope | Flash and Realisations comparison |
| `mart.agent_pcs_day` | Agent/day | PCS result and participation |

Legacy-named exports remain callable so existing jobs do not break.
`yesterday_gap_actions` covers the entire selected completed period, not only
yesterday. `mart.verint_final_absence_*` is currently a compatibility projection
of `mart.absence_*`; Verint Activities do not supply its values.

## Attendance and decision semantics

Agent Status has precedence. LILO fills only periods without reliable Status
coverage and never overwrites explicit Logged Off or Unavailable states. A
logout followed by a return stays an internal gap. Today never produces an
early-leave decision.

The Excel importer reads only Gap ID and the five editable decision columns.
SQLite supplies the authoritative date, agent and exact start/end. Approved
decisions use `config\wfm_rules.toml`; Dismissed counts as no loss; Open remains
unverified. The import is atomic. PTO/Away intervals enter the same classified
ledger without creating fake gaps.

Overlapping intervals are unioned before totals. Daily numerators are capped to
planned net minutes, so rates cannot exceed 100%. Summary rates include only
`CLEAR` and `ABSENCE_RECORDED` agent-days; `PENDING_REVIEW` and
`PROVISIONAL_DAY` cannot silently act as zero absence.

## Service semantics

One interaction may have several call legs. The service mart counts every
mapped inbound queue entry, matching Storm `Total Entered`, and ignores outbound
companion legs. Reports use the Storm business reference:

- total entered = mapped inbound queue entries;
- Storm A = lost calls; B = connected calls; C = connected within 30 seconds;
- Storm D = lost calls from 5 seconds up to, but not including, 30 seconds;
- SLA denominator = A + B - D = total entered - lost within SLA;
- response time = total queue wait + ringing duration;
- TSL = C / (A + B - D);
- routed rate = answered / total entered;
- deviation = total entered / forecast;
- AHT = total talk + hold + wrap seconds / answered.

Technical gross methods remain in the metric catalog. Higher-grain percentages
are always ratios of summed components, never averages of percentages.

Run **Export clean data** to write a selected CSV/XLSX under `Feed`; every export
has a manifest with its period, row count and rule provenance. Original extracts
remain unchanged.
