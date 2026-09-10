# WFM PTO & Away Submission App — Implementation Plan

Status: implementation kit in progress; Microsoft tenant deployment pending.
Owner: Anass ASSRI / WFM
Approved direction: SharePoint-first request workflow, WFM approval, then
serialized export to the unchanged FTE workbook.

## Outcome

Operations gets one compact bilingual app for PTO and Away requests. SharePoint
stores every request and its workflow history. Only WFM-approved operational
truth enters `FTE Count.xlsx`, so the workbook remains a clean WFMHub input and
does not become an unreliable approval queue.

The existing workbook contract does not change:

- `Agent` / `tblFTEAgents`
- `PTO` / `tblFTEPTO`
- `Away` / `tblFTEAway`

Client ID remains text. Sheets, tables and headers remain unchanged.

## Architecture

```text
FTE Count.xlsx / tblFTEAgents
             |
             v
       WFM_Sync_Agents
             |
             v
     WFM_Agents (SharePoint)
             |
             v
        Power Apps
  Submit | My requests | WFM review
             |
             v
 WFM_AbsenceRequests (SharePoint ledger)
             |
        WFM approval
             |
             v
 WFM_Sync_Absence_To_FTE (concurrency 1)
             |
             v
      Office Script writer
             |
             v
 tblFTEPTO / tblFTEAway -> WFMHub
```

SharePoint is the durable request and workflow authority. Excel is the
downstream WFMHub interchange contract. The app never writes to Excel directly.

## Solution components

Create one unmanaged development solution named `WFMAbsenceManagement` with
publisher prefix `wfm`.

### Canvas App

`WFM PTO & Away`

- responsive tablet/web layout using auto-layout containers;
- French default and English toggle;
- Operations: Submit and My requests;
- WFM reviewers: Review queue and request detail;
- one dynamic request form for PTO or Away;
- exact Client ID lookup with official read-only employee confirmation;
- no technical list, table or flow names exposed to Operations.

### SharePoint lists

`WFM_Agents` is a read-only app mirror of `tblFTEAgents`. Duplicate Client IDs
remain visible so lookup can explicitly return Ambiguous rather than silently
choosing one row.

`WFM_AbsenceRequests` is the durable workflow ledger. It owns Request ID,
fingerprint, submitter, reviewer, validation, approval, sync status, retries and
error details in addition to the absence fields.

### Flows

1. `WFM_Sync_Agents`
   - manual and scheduled;
   - reads `tblFTEAgents`;
   - refreshes the SharePoint lookup mirror;
   - reports blank and duplicate Client IDs;
   - never invents an employee identity.
2. `WFM_Submit_Absence`
   - Power Apps (V2) trigger;
   - repeats every important validation server-side;
   - creates one `Submitted` request and returns its Request ID;
   - never writes an unapproved request to Excel.
3. `WFM_Review_Absence`
   - callable only by a WFM reviewer;
   - revalidates the request and records Approve or Reject;
   - an approval sets workbook sync state to `Pending`.
4. `WFM_Sync_Absence_To_FTE`
   - automated and manually retryable;
   - concurrency is one;
   - invokes the governed Office Script;
   - writes approved PTO as `PTO` / `Approved`;
   - writes Away with date-aware `Planned`, `Active` or `Closed` state;
   - records `Synced` or a recoverable `Failed` state in SharePoint.
5. `WFM_Maintain_Away_Status`
   - daily;
   - promotes Planned to Active at Start date;
   - closes fixed-end Active records after End date;
   - uses the same serialized Office Script path.

The first production cut may combine review and immediate sync, but the durable
SharePoint record must be committed before Excel is touched.

## Workflow states

Request status:

```text
Submitted -> Approved -> Cancelled
          -> Rejected
```

Workbook sync status:

```text
Not required -> Pending -> Syncing -> Synced
                                  -> Failed -> Pending (retry)
```

Rejected requests never enter Excel. An approved cancellation changes the exact
matching workbook row to `Cancelled`; it does not delete audit history.

## Business validation

General:

- trim Client ID as text and require exactly one roster match;
- allow Active employees and Leavers only through their populated End date;
- require Start date, final accuracy confirmation and valid date order;
- protect against an identical active request and overlapping effective periods;
- obtain Name and organisational fields from the official roster mirror;
- store locale-neutral ISO dates and times.

PTO:

- require End date and reliquat confirmation;
- Full day writes blank Start time and End time;
- Partial day is one date and requires Start time before End time;
- do not show PTO type to Operations;
- write `PTO type = PTO` and `Approval status = Approved` only after WFM approval.

Away:

- require Away type;
- blank End date means an open case;
- block overlapping Away cases;
- WFM approval is mandatory because effective Away changes No Show, attendance
  and staffing interpretation;
- future Start date is Planned, current open/fixed case is Active, and an ended
  historical case is Closed.

The reliquat checkbox is an attestation, not a live balance check. Record who
confirmed it and when. A future leave-balance integration is a separate project.

## Security and ownership

- Operations receives app run access, not workbook edit access.
- WFM reviewers are controlled by a maintained Microsoft 365 or security group.
- Connections use a maintained functional owner, not a disposable personal owner.
- The SharePoint lists and workbook stay inside the restricted WFM site.
- Solution connection references and environment variables hold tenant-specific
  URLs, library names, workbook paths and reviewer group identity.
- The tenant owner must confirm DLP policy, connector availability and Office
  Scripts availability before production.

## Environment variables

| Name | Purpose |
|---|---|
| `wfm_SiteUrl` | Restricted WFM SharePoint site URL |
| `wfm_DocumentLibrary` | Library containing the FTE workbook |
| `wfm_FTEWorkbookPath` | Server-relative path to `FTE Count.xlsx` |
| `wfm_ReviewerGroup` | WFM reviewer group email or object ID |
| `wfm_DefaultLanguage` | `fr` |
| `wfm_AgentSyncSchedule` | Roster mirror refresh schedule |

## Delivery phases

1. Contract and production roster preflight.
2. SharePoint lists, indexes, views and permissions.
3. Office Script workbook writer and isolated workbook tests.
4. Solution-aware flows and connection references.
5. Responsive Canvas App and bilingual copy.
6. End-to-end test through WFMHub attendance and staffing outputs.
7. UAT, production connection binding, publication, sharing and owner handover.

## Responsibility boundary

Codex can create and test the repo-owned contracts, Office Script, Power Fx,
flow specifications, bilingual copy, validation matrix and WFMHub integration.
Creating connections, granting SharePoint permissions, opening the generated app
in Power Apps Studio, publishing it and sharing it require an authenticated user
in the Allianz Microsoft tenant. These are deployment steps, not missing design.

The current Power Apps source-code format is generated by Power Apps Studio and
is not a dependable supported route for synthesising a production Canvas App
from an unauthenticated Linux VPS. The tenant build will therefore follow the
repo-owned specifications and then export the completed Solution back to this
repository for backup and inspection.

## Production acceptance gate

- exact known-good roster lookup, including leading-zero and alphanumeric IDs;
- blank and duplicated IDs rejected visibly;
- PTO and Away validation matrix passes in both languages;
- unauthorized reviewer action fails;
- duplicate clicks create only one request;
- two concurrent approvals do not corrupt or duplicate Excel rows;
- workbook-open/locked errors remain retryable without losing the request;
- Away status transitions work across Start and End dates;
- FTE workbook reopens without repair and its three table contracts are intact;
- WFMHub ingests the resulting workbook and applies PTO/Away precedence correctly;
- owner, backup, retry and incident procedures are documented.
