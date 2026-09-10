# SharePoint setup

Create both lists in the restricted WFM SharePoint site. Use the internal names
from `sharepoint-lists.json` when each column is first created. The display name
may contain spaces after creation, but flows and formulas use the stable internal
name.

## 1. WFM_Agents

Create a blank list named `WFM_Agents`.

1. Rename the default Title display label to `Roster row` and keep it required.
2. Add every column under the `WFM_Agents` entry in
   `sharepoint-lists.json`.
3. Add indexes to `ClientID`, `EmploymentStatus`, `LOB`, `SyncBatchID` and
   `IsCurrent`.
4. Do not enforce uniqueness on Client ID. More than one row must remain visible
   so the app can return Ambiguous ID.
5. Give Operations read access only. The sync flow owner receives edit access.

The sync flow writes the employee name into both Title and EmployeeName.

## 2. WFM_AbsenceRequests

Create a blank list named `WFM_AbsenceRequests`.

1. Rename the default Title display label to `Request reference`.
2. Add every column under the `WFM_AbsenceRequests` entry in
   `sharepoint-lists.json`.
3. Enforce unique values on `RequestID`.
4. Add indexes to `RequestID`, `Fingerprint`, `ClientID`, `RequestType`,
   `LOB`, `RequesterUPN`, `StartDate`, `RequestStatus` and `SyncStatus`.
5. Create the three specified WFM views. Do not expose technical error columns
   in the Operations-facing view.

Operations needs list access only if the Canvas App reads My Requests directly.
If site governance does not permit row-level list access, replace that gallery
with a read flow before production rather than granting broad list visibility.

## 3. Permissions

Recommended SharePoint groups:

| Group | WFM_Agents | WFM_AbsenceRequests | FTE workbook |
|---|---|---|---|
| WFM App Operations | Read | Read through the app | No access |
| WFM App Reviewers | Read | Edit through the app | No direct edit during automation |
| WFM App Flow Owner | Edit | Edit | Edit |
| WFM Site Owners | Full control | Full control | Full control |

The app hides review screens for non-reviewers, but the review flow must perform
the real Microsoft 365 group check.

## 4. Views

`Current roster`:

- filter `IsCurrent = Yes`;
- show Client ID, Employee Name, Status, Team Leader, LOB and End date.

`WFM review queue`:

- filter `RequestStatus = Submitted`;
- sort Submitted UTC ascending;
- show Request reference, Request type, Client ID, Employee, period, requester
  and submitted time.

`Synchronization exceptions`:

- filter `SyncStatus = Failed`;
- show Request reference, employee, request status, retry count, last attempt and
  a shortened friendly error.

## 5. Do not do these things

- Do not import employee data into a public or Operations-wide site.
- Do not make Client ID a Number column.
- Do not enforce Client ID uniqueness before roster duplicates are resolved.
- Do not allow Operations to edit Approval, Case or Sync status.
- Do not connect the Canvas App directly to the Excel tables.
- Do not add Request ID or helper columns to the established Excel tables.
