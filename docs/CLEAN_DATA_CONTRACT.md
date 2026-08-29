# Governed clean-data contract

This release deliberately stops at clean data. It does not create the new
business workbooks yet. Validate these datasets first; each future workbook is
then only a presentation layer over an accepted contract.

Source extracts are opened read-only. FTE defines the admitted agents. Dates
come from row content, not filenames.

## Source boundaries

| Source | Allowed purpose |
|---|---|
| Verint StartEndTimes | Operational scheduled start/end and assignment |
| Storm LILO | Daily presence boundaries |
| Storm Agent Status | Actual within-shift states and attendance evidence |
| Verint Activities | Corrected post-day ledger only |
| APDE | Intraday service actuals |
| Verint Forecast | Forecast and required staffing only |
| Call by Call | Call-leg performance and PCS |

The parser identifies StartEndTimes and Activities from their headers and
stores `source_variant`. A missing StartEndTimes file is an error; Activities
cannot silently substitute for the operational schedule.

## Datasets to validate

| Export key | Grain | Future report use |
|---|---|---|
| `daily_attendance_calls` | Agent/day requiring a call | Daily absent/late call list |
| `daily_staffing_gaps` | Date/15 minutes/roster LOB/language | Daily staffing gap sheet |
| `daily_service_lob` | Date/interval/service LOB/language | Intraday SL state |
| `pcs_agent_day` | Agent/day | PCS detail |
| `pcs_team_day` | Team/LOB/language/day | PCS team summary |
| `pcs_agent_month` | Agent/month | PCS monthly view |
| `yesterday_gap_actions` | One uncovered correction interval | Yesterday injection list |
| `shift_evidence_timeline` | Exact shift segment | Verint-like shift visual |
| `verint_final_absence_events` | Corrected Activities evidence interval | Final audit detail |
| `verint_final_absence_day` | Agent/day | Final absenteeism |

`daily_attendance_calls.call_action` is explicit: `CALL_NO_SHOW`, `CALL_LATE`,
or `CALL_NOT_SEEN_NOW`. Current/future rows are marked `is_provisional`; an
unfinished current-day shift can never be finalized as early leave.

Staffing uses agent-seconds divided by 900, not averages of headcounts. Agent
Status has precedence; LILO fills only intervals where Agent Status has no
state. Explicit Logged Off or Unavailable time remains a gap. Future interval
gap and variance values are NULL.

`yesterday_gap_actions` contains only the residual pieces not already covered
by the union of corrected Activities intervals. A partially corrected original
gap can therefore produce one or more exact residual rows.

Final absenteeism is independent of LILO and Agent Status. It is built only
from the selected Activities snapshot, clips evidence to its Activities shift,
unions overlaps, and caps the rate denominator at the configured standard day
(default 8.75 hours).

Run an export from `WFMHub.cmd > Export clean data`. Each CSV/XLSX is written
under `output\data_exports\<dataset>` with a manifest containing its period,
row count, rule version, and rule hash.
