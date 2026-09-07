# WFM PTO & Away Submission App — Recorded Feasibility Plan

Status: recorded for a future project; no implementation has started.

## Decision

Proceed with a controlled pilot only after the business and integration corrections
below are accepted. The Canvas App and bilingual responsive interface are highly
feasible. Direct Excel use is the weak part of the design and is suitable only for
modest, serialized submission volume.

The existing `templates/FTE Count.xlsx` contract remains unchanged:

- `Agent` / `tblFTEAgents`
- `PTO` / `tblFTEPTO`
- `Away` / `tblFTEAway`

No sheet, table, or header is to be renamed. Client ID remains text.

## Required corrections before implementation

1. Operations will not select a PTO type, but the flow will write the fixed value
   `PTO` into the existing `PTO type` column. A blank value is rejected by the
   current WFMHub ingestion contract.
2. A new PTO request is written as `Pending`. It does not affect WFMHub attendance
   or staffing until WFM changes it to `Approved`. The WFM approval responsibility
   must be explicit in the operating procedure.
3. Away status is date-aware:
   - future Start date: `Planned`;
   - started and not expired, including no End date: `Active`;
   - End date in the past: `Closed`.
4. The app sends `ReliquatConfirmed` and `AccuracyConfirmed` to the flow so both
   controls can be validated again server-side.
5. Missing and duplicated Client IDs are a production-launch blocker. Zero matches
   returns Invalid ID; more than one exact match returns Ambiguous ID. The known
   duplicates `204046`, `96457`, and `224100` must be corrected in the populated
   roster before launch.
6. Excel operations are serialized. Operations does not edit the workbook directly.
7. Add a SharePoint request ledger for audit, idempotency, and recoverable errors
   without changing the FTE workbook structure.

## Recommended architecture

```text
Power Apps Canvas App
        |
        v
WFM_Submit_Absence
        |-- exact employee lookup and ambiguity check
        |-- server-side business validation
        |-- duplicate/idempotency protection
        |-- SharePoint audit ledger
        `-- serialized write to the established Excel table
                         |
                         v
                   FTE Count.xlsx
                         |
                         v
                       WFMHub
```

Use one responsive Canvas App. French is the default, with an English toggle. Use
auto-layout containers and one compact Operations submission experience. The app
does not connect directly to Excel; the flow owns workbook access.

Keep one principal flow named `WFM_Submit_Absence`. Add an `Operation` input:

- `LOOKUP`: validate Client ID and return the official employee record;
- `SUBMIT`: repeat lookup and all validation, create the audit record, write Excel,
  and return a durable Request ID or a clear failure.

The flow connection must be owned by a maintained functional owner, packaged in a
Power Platform Solution with connection references and environment variables. Set
trigger concurrency to one for the Excel section.

## SharePoint audit ledger

Use a small list such as `WFM_AbsenceRequests` with at least:

- Request ID and deterministic request fingerprint;
- requester identity and submission timestamp;
- selected language and request type;
- Client ID and official employee name;
- dates, coverage, times, Away type, and comment;
- reliquat and final-accuracy confirmations;
- validation result, workbook-write status, and error detail.

Excel remains the unchanged WFMHub interchange workbook. The SharePoint list is the
submission and technical audit ledger, not a replacement for the workbook.

## Validation contract

General:

- Client ID is trimmed text and is never converted with `Value()`;
- exactly one matching roster row is required;
- Start date is required and End date cannot precede it;
- final accuracy confirmation is required;
- requests beyond the employee's effective leaving date are blocked;
- exact duplicates and overlapping PTO/Away periods are detected;
- data values use locale-neutral `yyyy-MM-dd` dates and `HH:mm:ss` times.

PTO:

- End date and reliquat confirmation are required;
- partial-day PTO must start and end on one date and requires increasing times;
- full-day PTO writes blank time fields;
- `PTO type` is written automatically as `PTO`;
- Approval status is written as `Pending`.

Away:

- Away type is required;
- open cases may have a blank End date;
- only one overlapping/open Away case is allowed per employee;
- Case status follows the date-aware Planned/Active/Closed rule above.

## Acceptance tests

- valid numeric, leading-zero, and alphanumeric Client IDs;
- unknown and duplicate Client IDs;
- full-day and partial-day PTO;
- invalid time and date ranges;
- future, current open, current fixed-end, and historical Away cases;
- duplicate and overlapping submissions;
- active/leaver boundary cases;
- two near-simultaneous submissions;
- workbook-open/locked failure and safe retry;
- bilingual labels, validation, success, and error messages;
- successful WFMHub import after the resulting workbook syncs locally.

## Go/no-go gate

Go when the fixed PTO type, date-aware Away logic, approval ownership, confirmation
payload, audit ledger, roster cleanup, and serialized Excel-writer rules are accepted;
the tenant permits the required standard connectors; and a maintained flow owner and
SharePoint location exist. Otherwise the design is demonstration-only.
