# PTO & Away app acceptance tests

Record evidence for every test in the development Solution before production.
Use synthetic employees in a copy of the workbook where possible.

## A. Access and language

| ID | Test | Expected |
|---|---|---|
| A01 | Open app for the first time | English interface is selected |
| A02 | Select FR | Every visible form label and message changes to French |
| A03 | Return to EN | English returns without losing entered values |
| A04 | Operations user opens app | Submit and My Requests visible; WFM Review hidden |
| A05 | Operations invokes review flow outside UI | Flow returns Unauthorized and changes nothing |
| A06 | Reviewer opens app | Review queue is visible |

## B. Employee lookup

| ID | Test | Expected |
|---|---|---|
| B01 | Active numeric-looking text ID `00123` | Leading zeros preserved and exact employee shown |
| B02 | Alphanumeric ID | Exact employee shown |
| B03 | Unknown ID | Invalid ID; submit disabled |
| B04 | Blank ID | Required message; no query submitted |
| B05 | Duplicated ID | Ambiguous ID; no employee selected |
| B06 | Modify Client ID after lookup | Previous employee confirmation is cleared |
| B07 | Leaver request ending on leave date | Allowed |
| B08 | Leaver request after leave date | Blocked |
| B09 | Leaver with no End date | Blocked as roster data-quality issue |

## C. PTO

| ID | Test | Expected |
|---|---|---|
| C01 | Valid full-day single date | Submitted; blank times stored |
| C02 | Valid full-day multi-date | Submitted |
| C03 | Missing End date | Blocked in app and flow |
| C04 | End before Start | Blocked in app and flow |
| C05 | Reliquat unchecked | Blocked in app and flow |
| C06 | Accuracy unchecked | Blocked in app and flow |
| C07 | Partial day with same date and increasing times | Submitted |
| C08 | Partial day across dates | Blocked |
| C09 | Partial day missing a time | Blocked |
| C10 | Partial day End time not after Start | Blocked |
| C11 | Approve valid PTO | One Excel row; PTO type PTO; status Approved |
| C12 | Reject PTO | No Excel row |
| C13 | Cancel synced PTO | Exact Excel row becomes Cancelled |

## D. Away

| ID | Test | Expected |
|---|---|---|
| D01 | Future fixed-end Away | Excel status Planned |
| D02 | Current fixed-end Away | Excel status Active |
| D03 | Historical fixed-end Away | Excel status Closed |
| D04 | Open-ended Away | Excel status Active and blank End date |
| D05 | Missing Away type | Blocked in app and flow |
| D06 | End before Start | Blocked in app and flow |
| D07 | Approve Away | One exact Excel row |
| D08 | Reject Away | No Excel row |
| D09 | Planned case reaches Start date | Daily maintenance updates exact row to Active |
| D10 | Active fixed case passes End date | Daily maintenance updates exact row to Closed |
| D11 | Cancel synced Away | Exact row becomes Cancelled |

## E. Duplicate, overlap and resilience

| ID | Test | Expected |
|---|---|---|
| E01 | Double-click Submit | One SharePoint request only |
| E02 | Retry same Request ID after timeout | Original request returned; no duplicate |
| E03 | Same active business fingerprint | Duplicate blocked |
| E04 | PTO overlaps active PTO | Blocked |
| E05 | Away overlaps active Away | Blocked |
| E06 | PTO overlaps active Away | Blocked |
| E07 | Two simultaneous approvals | Serialized; workbook remains valid |
| E08 | Workbook open/locked | Request remains Approved; Sync Failed and retryable |
| E09 | Retry after workbook unlock | One exact row; Sync becomes Synced |
| E10 | Excel succeeds but final SharePoint update fails | Retry returns NO_CHANGE and recovers ledger |
| E11 | Two exact workbook rows already exist | Script fails Ambiguous; neither row changed |

## F. WFMHub integration

| ID | Test | Expected |
|---|---|---|
| F01 | Run local contract checker | PASS; exact sheets/tables/headers |
| F02 | Ingest approved full-day PTO | Parsed as PTO / APPROVED |
| F03 | Ingest approved partial-day PTO | Exact times parsed |
| F04 | Ingest Active open Away | Parsed with blank End date |
| F05 | Refresh affected day | PTO/Away removes only overlapping scheduled time |
| F06 | Refresh RTM for registered PTO/Away | Registered time off excluded from Due and No Show |
| F07 | Refresh staffing future period | Planned time off reduces net capacity correctly |
| F08 | Open workbook after all flow actions | No Excel repair prompt; all three tables intact |

## Sign-off

Production requires WFM owner, Operations representative and technical owner
sign-off. Record Solution version, SharePoint site, workbook path, test date and
the person who completed each result. Do not include production employee data in
the repository evidence.
