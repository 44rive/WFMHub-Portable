# Governed clean-data contract

These datasets are the governed contracts behind the focused business workbooks.
The workbook layer formats and filters them; it does not invent a second KPI
definition. The same contracts remain directly exportable for audit, sending,
custom Python, or a future report.

Source extracts are opened read-only. FTE defines the admitted agents. Dates
come from row content, not filenames.

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
unions overlaps, caps planned net minutes at the configured standard day
(default 8.75 hours), and caps each daily classified numerator to that planned
net value. A final rate therefore cannot exceed 100%.

Run an export from `WFMHub.cmd > Export clean data`. Each CSV/XLSX is written
under `output\data_exports\<dataset>` with a manifest containing its period,
row count, rule version, and rule hash.

Each current report also writes smaller presentation-grain CSV tables under
`output\model_data\<report>`. Those files are for the optional Excel Data Model
route and are not a second calculation layer. `manifest.json` records their
period, row counts, hashes, and generation time.

PCS has the stronger end-user template contract under
`output\template_feeds\pcs\current`:

- `PCS_AgentDay.csv`
- `PCS_Summary.csv`
- `PCS_Actions.csv`
- `PCS_Trend.csv`
- `manifest.json`

Every file has a stable name and UTF-8 header. Numeric/date types are declared
again in the supplied Power Query scripts. A complete copy of every published
set is kept under `output\template_feeds\pcs\archive`.
