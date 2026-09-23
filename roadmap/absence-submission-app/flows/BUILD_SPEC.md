# Power Automate build specification

Build every flow inside the `WFMAbsenceManagement` Solution. Use connection
references for SharePoint, Excel Online (Business), Office 365 Users and
Microsoft 365 Groups. English is the default response language; every user-facing
error selects English or French from the validated request payload.

All date payloads use `yyyy-MM-dd`. All time payloads use `HH:mm:ss`. Client ID
is always a trimmed string.

## Shared conventions

Use named Scopes:

```text
TRY
CATCH
FINALLY
```

Configure CATCH to run after TRY has failed, timed out or been skipped. Never
return a success before the SharePoint ledger commit completes.

For Excel actions:

- concurrency: one;
- exponential retry policy;
- maximum four retries;
- no parallel Apply to each;
- never delete workbook rows;
- Office Script response `ok = false` is a handled synchronization failure.

### Submission payload schema

```json
{
  "type": "object",
  "required": [
    "contractVersion", "requestId", "requestType", "clientId", "startDate",
    "reliquatConfirmed", "accuracyConfirmed", "selectedLanguage",
    "requesterUpn", "requesterName"
  ],
  "properties": {
    "contractVersion": {"type": "string"},
    "requestId": {"type": "string"},
    "requestType": {"type": "string"},
    "clientId": {"type": "string"},
    "startDate": {"type": "string"},
    "endDate": {"type": "string"},
    "dayCoverage": {"type": "string"},
    "startTime": {"type": "string"},
    "endTime": {"type": "string"},
    "awayType": {"type": "string"},
    "comment": {"type": "string"},
    "reliquatConfirmed": {"type": "boolean"},
    "accuracyConfirmed": {"type": "boolean"},
    "selectedLanguage": {"type": "string"},
    "requesterUpn": {"type": "string"},
    "requesterName": {"type": "string"}
  }
}
```

## Flow 1 — WFM_Sync_Agents

Triggers:

- Recurrence using the agreed roster refresh schedule;
- manual trigger for WFM administrators.

Settings: trigger concurrency one.

Actions:

1. Compose `BatchID = guid()`.
2. Excel **List rows present in a table** from `tblFTEAgents`. Enable pagination
   above the maximum expected roster size.
3. Initialize counters: Total, BlankClientID, Duplicates and Rejected.
4. For each Excel row, validate Name and Status. Create a new inactive
   `WFM_Agents` item with:
   - `IsCurrent = false`;
   - `SyncBatchID = BatchID`;
   - Client ID written as text;
   - every available organisation field copied without inference.
5. If any create failed, terminate Failed. The previous current batch remains
   untouched and the partial new batch remains invisible.
6. Query the new batch, group exact trimmed Client IDs and count blanks and
   duplicates. Duplicates are retained; do not choose a winner.
7. Mark every new batch row `IsCurrent = true`.
8. Only after step 7 succeeds, mark prior current batches `IsCurrent = false`.
9. Remove inactive batches older than the configured retention period.
10. Send the WFM owner a concise data-quality result when blank or duplicated
    Client IDs exist.

Do not update the old batch row-by-row before the complete new batch exists.
This prevents a partial sync from taking employee lookup offline.

## Flow 2 — WFM_Submit_Absence

Trigger: **Power Apps (V2)** with one required Text input named `PayloadJson`.

Response outputs:

- `ok` Boolean;
- `requestId` Text;
- `code` Text;
- `message` Text.

Actions inside TRY:

1. Parse JSON using the submission schema.
2. Normalize:
   - Request ID and Contract Version: trim;
   - Request Type: uppercase;
   - Client ID: trim only, never `int()` or `float()`;
   - language: `FR` only when exactly FR, otherwise EN;
   - times and dates: retain locale-neutral strings.
3. Validate Contract Version `1.0.0`, a GUID Request ID, Request Type, Start date
   and Accuracy confirmation.
4. Query `WFM_AbsenceRequests` by Request ID.
   - one matching identical request returns the stored result as an idempotent
     success;
   - one conflicting request returns `REQUEST_ID_CONFLICT`;
   - more than one is a ledger integrity error.
5. Query current `WFM_Agents` rows by exact escaped Client ID and
   `IsCurrent = true`.
   - zero: `INVALID_CLIENT_ID`;
   - more than one: `AMBIGUOUS_CLIENT_ID`;
   - exactly one: continue and use its official employee fields.
6. Validate employee eligibility:
   - Active: allowed;
   - Leaver: populated End date required and the requested period must not extend
     beyond it;
   - an open-ended Away request is not allowed for a Leaver;
   - anything else: `EMPLOYEE_NOT_ELIGIBLE`.
7. Validate PTO:
   - End date is required;
   - reliquat and accuracy confirmations are true;
   - coverage is Full day or Partial day;
   - Partial day uses one date and increasing `HH:mm:ss` times;
   - Full day ignores and stores blank times.
8. Validate Away:
   - Away type is one of the four governed values;
   - blank End date is allowed;
   - populated End date is not earlier than Start date.
9. Build a deterministic Fingerprint from normalized business fields:

   ```text
   CLIENTID|TYPE|START|END|COVERAGE|STARTTIME|ENDTIME|AWAYTYPE
   ```

10. Query existing Submitted or Approved requests for the Client ID.
11. Reject an identical active fingerprint as `DUPLICATE_REQUEST`.
12. Reject date overlap as `OVERLAPPING_REQUEST`. Blank Away End date is treated
    as open-ended. This cross-check includes PTO-versus-Away overlap.
13. Create the SharePoint item. The flow supplies official name and organisation
    fields; it never uses a user-entered name.
14. Set:
    - Title = `WFM-` plus the first eight Request ID characters;
    - Request Status = Submitted;
    - Sync Status = Not required;
    - Submitted UTC = `utcNow()`;
    - Retry Count = 0;
    - Contract Version = 1.0.0.
15. Respond success only after the Create item action succeeds.

CATCH responds with a short localized error. Do not expose raw connector paths,
tokens, list names or stack traces to Operations.

## Flow 3 — WFM_Review_Absence

Trigger: **Power Apps (V2)** with required `PayloadJson` Text.

Accepted decisions: `APPROVE`, `REJECT`, `CANCEL`, `RETRY`.

Actions:

1. Parse and validate Contract Version and Request ID.
2. Resolve the actual invoking identity from the Power Apps trigger caller
   metadata. Do not authorize from the passed `reviewerUpn` alone.
3. Query the configured Microsoft 365 reviewer group and require exact caller
   membership. Otherwise terminate `UNAUTHORIZED_REVIEWER`.
4. Retrieve exactly one request by Request ID.
5. For APPROVE:
   - current Request Status must be Submitted;
   - repeat employee, date, field, duplicate and overlap validation;
   - set Request Status Approved;
   - set Sync Status Pending;
   - record reviewer, UTC time and comment.
6. For REJECT:
   - current status must be Submitted;
   - require a review comment;
   - set Request Status Rejected and Sync Status Not required.
7. For CANCEL:
   - Submitted becomes Cancelled / Not required;
   - Approved becomes Cancelled / Pending so the workbook row is marked
     Cancelled;
   - Rejected and already Cancelled requests cannot be cancelled again.
8. For RETRY:
   - require Sync Status Failed;
   - keep Request Status unchanged and set Sync Status Pending.
9. Respond with the durable Request ID and new states.

## Flow 4 — WFM_Sync_Absence_To_FTE

Trigger: Recurrence every two minutes during the operating window. Trigger
concurrency one.

Actions:

1. Get up to 20 items where `SyncStatus = Pending`, oldest Reviewed UTC first.
2. Apply to each sequentially.
3. Update the item to Syncing and increment Retry Count.
4. Compose `TodayLocal` with the Power Automate Windows time-zone identifier
   configured for the WFM site, formatted `yyyy-MM-dd`.
5. Compose the Office Script payload:

   ```json
   {
     "contractVersion": "1.0.0",
     "requestId": "<RequestID>",
     "operation": "<CANCEL when RequestStatus is Cancelled, otherwise UPSERT>",
     "requestType": "<PTO or AWAY uppercased>",
     "clientId": "<ClientID>",
     "employeeName": "<EmployeeName>",
     "startDate": "<yyyy-MM-dd>",
     "endDate": "<yyyy-MM-dd or blank>",
     "dayCoverage": "<Full day, Partial day or blank>",
     "startTime": "<HH:mm:ss or blank>",
     "endTime": "<HH:mm:ss or blank>",
     "awayType": "<governed Away type or blank>",
     "comment": "<Comment or blank>",
     "today": "<TodayLocal>"
   }
   ```

6. Run `WFM_Write_Absence` against the configured `FTE Count.xlsx`.
7. If result `ok = true`, update:
   - Sync Status = Synced;
   - Last Sync UTC = `utcNow()`;
   - Last Sync Error = blank;
   - Workbook Action = script action;
   - Workbook Status = script workbook status.
8. If result `ok = false` or the connector fails, update:
   - Sync Status = Failed;
   - Last Sync UTC = `utcNow()`;
   - Last Sync Error = bounded technical message for WFM;
   - never change the approval decision.

The Office Script upsert makes a retry safe when Excel succeeded but the final
SharePoint update failed.

## Flow 5 — WFM_Maintain_Away_Status

Trigger: daily after local midnight and before WFMHub's first planned refresh.
Concurrency one.

1. Compute TodayLocal.
2. Query Approved, Synced Away requests.
3. Set Sync Status Pending when:
   - Workbook Status is Planned and Start Date is today or earlier;
   - Workbook Status is Active, End Date is populated and before today.
4. The normal serialized sync flow derives and writes Active or Closed.
5. Open-ended Away stays Active. A future open-ended Away is also Active because
   the current WFMHub ingestion contract permits blank End date only for Active;
   its Start date still prevents early application.

## Operational failure rules

- A SharePoint submission success never depends on Excel being available.
- A workbook lock creates Failed sync, not a lost request.
- WFM can retry a failed sync from the app.
- Never ask Operations to resubmit a request solely because Excel was locked.
- Do not run desktop Excel edits, Office Scripts and WFMHub workbook replacement
  against the same file simultaneously.
