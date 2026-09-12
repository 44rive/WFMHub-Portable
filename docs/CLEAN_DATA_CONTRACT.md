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
| Verint Activities | Final post-day absence/shrinkage codes and residual-gap reconciliation; never observed presence |
| Verint Forecast | Forecast and required staffing only; native 15/60-minute grain retained |
| Storm Call by Call | All mapped service actuals, Flashes, agent call performance and PCS |

APBE, APFR and APDE are retired: no directory is discovered, required, loaded or
calculated. Historical raw tables remain physically readable so upgrading is
non-destructive. Activities event intervals are excluded from Attendance/RTM
and are used only by the separate final post-day absence ledger. A parsed
Activities Shift Assignment may also be used as a visibly flagged emergency
schedule boundary when StartEndTimes is missing.

## Key datasets

| Dataset | Grain | Business use |
|---|---|---|
| `mart.attendance_agent_day` | Agent/day | Attendance callout and evidence result |
| `mart.correction_candidate` | Exact residual continuous gap | Attendance Review correction backlog |
| `mart.shift_timeline_segment` | Exact shift segment | Readable shift evidence |
| `mart.planned_time_off_segment` | Schedule-clipped interval | PTO/Away planning overlay |
| `mart.absence_event` | Reviewed or planned exact interval | Absence/shrinkage component audit |
| `mart.absence_agent_day` | Agent/day | Reviewed absence and shrinkage result |
| `mart.call_service_hour` | Date/hour/mapped queue | Interaction-deduplicated service actual |
| `mart.service_interval` | Date/hour/mapped queue | Stable semantic projection of Call-by-Call |
| `mart.forecast_interval` | Native Verint interval | Staffing and clean sharing |
| `mart.forecast_hour` | Date/hour/mapped scope | Flash and Realisations comparison |
| `mart.agent_pcs_day` | Agent/day | PCS result and participation |

The PCS collaboration boundary is six fixed UTF-8 CSV products under
`Feed\PCS`: governed filter lists, LOB cache, agent cache, daily cache,
standard-period results, and exact low-score coaching opportunities. These feeds contain only
governed result grains required by the workbook; the raw/deduplicated call-leg
table remains in SQLite unless explicitly exported. Power Query transports the
fixed schemas to `_PCS_FILTERS`, `_PCS_LOB`, `_PCS_AGENT`, `_PCS_DAILY`,
`_PCS_COACH`, and `PERFORMANCE`. The visible coaching queue and permanent
`tblCoachingActions` table are not feed destinations.

Legacy-named exports remain callable so existing jobs do not break.
`yesterday_gap_actions` covers the entire selected completed period, not only
yesterday. `mart.verint_final_absence_*` is sourced from mapped Verint Activities
inside the StartEndTimes shift boundary. It stays separate from the provisional
observed-gap `mart.absence_*` control tables.

## Attendance and reconciliation semantics

Agent Status has precedence. LILO fills only periods without reliable Status
coverage and never overwrites explicit Logged Off or Unavailable states. A
logout followed by a return stays an internal gap. Today never produces an
early-leave decision.

PTO/Away precedence is interval-based and schedule-clipped. Approved PTO and
effective Active/Closed Away remove their exact intervals before late,
early-leave, no-show, correction-gap, and net-staffing calculations. Pending or
Cancelled rows do nothing. Planned Away is future-capacity information only and
cannot erase elapsed attendance evidence. RTM exposes registered time off while
keeping it outside Due HC and all call/no-show counters.

SQLite supplies the authoritative date, agent and exact start/end. After the raw
gap is detected, mapped final Verint Activities are clipped to the same shift
and their exact overlap is subtracted. Only residual fragments are exported to
Attendance Review. The workbook is read-only; the operator corrects Verint,
exports Activities again and refreshes. PTO/Away intervals remain governed
planned evidence and never create fake gaps.

Overlapping intervals are unioned before totals. Daily numerators are capped to
planned net minutes, so rates cannot exceed 100%. Summary rates include only
`CLEAR` and `ABSENCE_RECORDED` agent-days; `PENDING_VERINT` and
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
