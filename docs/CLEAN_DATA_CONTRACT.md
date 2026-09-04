# Governed clean-data contract

These datasets are the governed contracts behind the focused business workbooks.
The workbook layer formats and filters them; it does not invent a second KPI
definition. The same contracts remain directly exportable for audit, sending,
custom Python, or a future report.

Source extracts are opened read-only. FTE defines the admitted agents. Dates
come from row content, not filenames.

FTE scope is evaluated per business date: Active rows remain eligible, Leavers
remain eligible through their populated leave date, and every other status is
excluded. Historical rows before a leave date remain valid.

## Source boundaries

| Source | Allowed purpose |
|---|---|
| Verint StartEndTimes | Operational scheduled start/end and assignment |
| Storm LILO | Daily presence boundaries |
| Storm Agent Status | Actual within-shift states and attendance evidence |
| Verint Activities | Corrected post-day ledger only |
| APBE/APFR/APDE | Service actuals only |
| Verint Forecast | Forecast and required staffing only |
| Call by Call | Call-leg performance and PCS |

The parser identifies StartEndTimes and Activities from their headers and
stores `source_variant`. A missing StartEndTimes file is an error; Activities
cannot silently substitute for the operational schedule.

## Datasets to validate

| Export key | Grain | Workbook use |
|---|---|---|
| `daily_attendance_calls` | Agent/day requiring a call | Daily absent/late call list |
| `daily_staffing_gaps` | Date/15 minutes/roster LOB/language | Daily staffing gap sheet |
| `planned_time_off` | Exact schedule-clipped agent interval | Approved PTO and effective Away planning overlay |
| `daily_service_lob` | Date/interval/service LOB/language | Intraday SL state |
| `pcs_agent_day` | Agent/day | PCS detail |
| `pcs_team_day` | Team/LOB/language/day | PCS team summary |
| `pcs_agent_month` | Agent/month | PCS monthly view |
| `yesterday_gap_actions` | One uncovered correction interval | Selected-period completed-day review |
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

`yesterday_gap_actions` is a legacy export key. It contains only the residual pieces not already covered
by the union of corrected Activities intervals. A partially corrected original
gap can therefore produce one or more exact residual rows.

Final absenteeism amounts and categories are built only from the selected
Activities snapshot. Evidence is clipped to the Activities shift, overlaps are
unioned, planned net minutes are capped at the configured standard day (default
8.75 hours), and each daily classified numerator is capped to that planned net
value. A final rate therefore cannot exceed 100%.

LILO and Agent Status do not create a payroll category, but they are used as a
completeness control. `UNCODED_EMPTY_SHIFT`, `UNCORRECTED_OBSERVED_GAP`,
`PARTIAL_CORRECTION_REVIEW`, `VERINT_WITHOUT_OBSERVED_GAP`, and
`PROVISIONAL_DAY` remain exceptions. Finalized
summary rates and the LOB/month export include only `CLEAR` and
`ABSENCE_RECORDED` agent-days, preventing incomplete rows from acting like zero
absence.

Run an export from `WFMHub.cmd > Export clean data`. Each CSV/XLSX is written
under the visible `Feed` folder with a manifest containing its period,
row count, rule version, and rule hash.

PCS does not use a separate feed contract. Its generated workbook contains a
visible `PCS_DATA` Excel Table and a visible `COACHING` Excel Table. Dashboard
formulas read those tables directly, so there is no connection or refresh step.
Editable coaching cells carry forward from the previous current workbook by
Coaching Key and never enter SQLite.
