# Canvas App build contract

App name: `WFM PTO & Away`
Default language: English (`en`)
Form factor: responsive tablet/web Canvas App
Design reference: `docs/prototypes/WFM_PTO_Away_App_Form.png`

Use auto-layout containers. Disable **Scale to fit**, **Lock aspect ratio** and
**Lock orientation**. Keep the main content at a maximum width of 1280 px and
use the WFMHub tokens below.

## Theme tokens

Create these App formulas or global variables:

```powerfx
clrNavy = ColorValue("#0B1F33");
clrTeal = ColorValue("#007C83");
clrGold = ColorValue("#D6A84B");
clrCanvas = ColorValue("#F4F7F9");
clrInk = ColorValue("#1F2933");
clrMuted = ColorValue("#536474");
clrLine = ColorValue("#D8E0E6");
clrBlue = ColorValue("#0563C1");
clrGreen = ColorValue("#1F7A53");
clrAmber = ColorValue("#A65F00");
clrRed = ColorValue("#B42318");
```

Use Aptos Display for titles and Aptos for body text where the tenant exposes
them. Do not add gradients, oversized cards, decorative icons or technical
connector names.

## Data connections

- SharePoint list `WFM_Agents`
- SharePoint list `WFM_AbsenceRequests`
- flow `WFM_Submit_Absence`
- flow `WFM_Review_Absence`

The app reads SharePoint but never reads or writes `FTE Count.xlsx` directly.

## Screens

| Screen | Access | Purpose |
|---|---|---|
| `scrSubmit` | Operations and WFM | Employee lookup and one PTO/Away form |
| `scrMyRequests` | Operations and WFM | Requests submitted by the signed-in user |
| `scrReview` | WFM only | Submitted queue and failed synchronization queue |
| `scrReviewDetail` | WFM only | Validate, approve, reject, cancel or retry one request |

The Review navigation item and screens are hidden from non-reviewers. The flow
must repeat reviewer authorization; screen visibility is not security.

## App initialization

The formulas below use English authoring separators. Power Apps automatically
localizes separators when pasted into a differently configured authoring locale.

`App.OnStart`:

```powerfx
Set(gblLanguage, "en");
Set(gblUserEmail, Lower(User().Email));
Set(gblUserName, User().FullName);
Set(gblRequestType, "PTO");
Set(gblCoverage, "Full day");
Set(gblAgent, Blank());
Set(gblLookupState, "EMPTY");
Set(gblSubmitting, false);
Set(gblRequestId, Blank());
Set(gblIsReviewer, false)
```

Populate `colText` from `locales/labels.csv` as a static app collection. Label
text uses this pattern:

```powerfx
LookUp(colText, key = "app_title", If(gblLanguage = "fr", fr, en))
```

English button `OnSelect`:

```powerfx
Set(gblLanguage, "en")
```

French button `OnSelect`:

```powerfx
Set(gblLanguage, "fr")
```

## Submit screen control map

```text
scrSubmit
`-- conPage (vertical)
    |-- conHeader (horizontal)
    |   |-- grpBrand
    |   |-- grpTitle
    |   |-- btnLanguageEN
    |   `-- btnLanguageFR
    `-- conMainWidth (vertical, max 1280)
        |-- conIntro
        `-- conColumns (horizontal; vertical below 900 px)
            |-- conForm (vertical, FillPortions 2)
            |   |-- crdRequestType
            |   |-- crdEmployee
            |   `-- crdAbsence
            `-- crdSummary (FillPortions 1)
```

Use these stable control names because the formulas below depend on them:

```text
txtClientID              btnSearchEmployee
btnTypePTO               btnTypeAway
dpStartDate              dpEndDate
chkNoEndDate             btnCoverageFull             btnCoveragePartial
drpStartHour             drpStartMinute
drpEndHour               drpEndMinute
drpAwayType              txtComment
chkReliquat              chkAccuracy
btnSubmitRequest         btnClearRequest
galMyRequests            galReviewQueue
txtReviewComment         btnApprove                   btnReject
```

## Employee lookup

`btnSearchEmployee.OnSelect`:

```powerfx
Set(gblClientID, Trim(txtClientID.Text));
Set(gblAgent, Blank());
If(
    IsBlank(gblClientID),
    Set(gblLookupState, "EMPTY"),
    ClearCollect(
        colAgentMatches,
        Filter(WFM_Agents, ClientID = gblClientID && IsCurrent = true)
    );
    Switch(
        CountRows(colAgentMatches),
        0, Set(gblLookupState, "NOT_FOUND"),
        1, Set(gblAgent, First(colAgentMatches)); Set(gblLookupState, "FOUND"),
        Set(gblLookupState, "AMBIGUOUS")
    )
)
```

`txtClientID.OnChange` resets stale confirmation:

```powerfx
Set(gblAgent, Blank());
Set(gblLookupState, "EMPTY");
Set(gblRequestId, Blank())
```

Never use `Value(txtClientID.Text)`. Client ID must remain text so leading zeros
and letters survive.

## Dynamic request fields

`btnTypePTO.OnSelect`:

```powerfx
Set(gblRequestType, "PTO");
Set(gblRequestId, Blank())
```

`btnTypeAway.OnSelect`:

```powerfx
Set(gblRequestType, "AWAY");
Set(gblRequestId, Blank())
```

PTO coverage, time and reliquat controls:

```powerfx
gblRequestType = "PTO"
```

Partial-day time controls:

```powerfx
gblRequestType = "PTO" && gblCoverage = "Partial day"
```

Canvas App has no dependable cross-tenant time picker requirement in this
design. Use four compact dropdowns. Hour dropdown items:

```powerfx
ForAll(Sequence(24), Text(Value - 1, "00"))
```

Minute dropdown items:

```powerfx
["00", "15", "30", "45"]
```

Away type and no-end-date controls:

```powerfx
gblRequestType = "AWAY"
```

End date control:

```powerfx
!(gblRequestType = "AWAY" && chkNoEndDate.Value)
```

## Client-side validation

Create `gblFormValid` whenever a relevant field changes, or place this formula
directly in `btnSubmitRequest.DisplayMode`:

```powerfx
If(
    gblLookupState <> "FOUND"
        || IsBlank(dpStartDate.SelectedDate)
        || !chkAccuracy.Value
        || (
            gblRequestType = "PTO"
            && (
                IsBlank(dpEndDate.SelectedDate)
                || dpEndDate.SelectedDate < dpStartDate.SelectedDate
                || !chkReliquat.Value
                || (
                    gblCoverage = "Partial day"
                    && (
                        dpEndDate.SelectedDate <> dpStartDate.SelectedDate
                        || IsBlank(drpStartHour.Selected.Value)
                        || IsBlank(drpStartMinute.Selected.Value)
                        || IsBlank(drpEndHour.Selected.Value)
                        || IsBlank(drpEndMinute.Selected.Value)
                        || (
                            Value(drpEndHour.Selected.Value) * 60
                                + Value(drpEndMinute.Selected.Value)
                            <= Value(drpStartHour.Selected.Value) * 60
                                + Value(drpStartMinute.Selected.Value)
                        )
                    )
                )
            )
        )
        || (
            gblRequestType = "AWAY"
            && (
                IsBlank(drpAwayType.Selected.Value)
                || (!chkNoEndDate.Value && IsBlank(dpEndDate.SelectedDate))
                || (!chkNoEndDate.Value && dpEndDate.SelectedDate < dpStartDate.SelectedDate)
            )
        )
        || gblSubmitting,
    DisplayMode.Disabled,
    DisplayMode.Edit
)
```

## Submission payload

`btnSubmitRequest.OnSelect`:

```powerfx
Set(gblSubmitting, true);
If(IsBlank(gblRequestId), Set(gblRequestId, Text(GUID())));
Set(
    gblSubmitResult,
    WFM_Submit_Absence.Run(
        JSON(
            {
                contractVersion: "1.0.0",
                requestId: gblRequestId,
                requestType: gblRequestType,
                clientId: Trim(txtClientID.Text),
                startDate: Text(dpStartDate.SelectedDate, "[$-en-US]yyyy-mm-dd"),
                endDate: If(
                    gblRequestType = "AWAY" && chkNoEndDate.Value,
                    "",
                    Text(dpEndDate.SelectedDate, "[$-en-US]yyyy-mm-dd")
                ),
                dayCoverage: If(gblRequestType = "PTO", gblCoverage, ""),
                startTime: If(
                    gblRequestType = "PTO" && gblCoverage = "Partial day",
                    drpStartHour.Selected.Value & ":" & drpStartMinute.Selected.Value & ":00",
                    ""
                ),
                endTime: If(
                    gblRequestType = "PTO" && gblCoverage = "Partial day",
                    drpEndHour.Selected.Value & ":" & drpEndMinute.Selected.Value & ":00",
                    ""
                ),
                awayType: If(gblRequestType = "AWAY", drpAwayType.Selected.Value, ""),
                comment: Trim(txtComment.Text),
                reliquatConfirmed: gblRequestType = "PTO" && chkReliquat.Value,
                accuracyConfirmed: chkAccuracy.Value,
                selectedLanguage: Upper(gblLanguage),
                requesterUpn: gblUserEmail,
                requesterName: gblUserName
            },
            JSONFormat.Compact
        )
    )
);
Set(gblSubmitting, false);
If(
    gblSubmitResult.ok,
    Notify(gblSubmitResult.message, NotificationType.Success);
    Navigate(scrMyRequests, ScreenTransition.Fade);
    Set(gblRequestId, Blank()),
    Notify(gblSubmitResult.message, NotificationType.Error)
)
```

Keep the generated Request ID after a failed call. Retrying then remains
idempotent. Generate a new ID only after success or an explicit form reset.

## My requests

`galMyRequests.Items`:

```powerfx
SortByColumns(
    Filter(WFM_AbsenceRequests, Lower(RequesterUPN) = gblUserEmail),
    "SubmittedUTC",
    SortOrder.Descending
)
```

Show Request Type, employee, period, Request Status and Sync Status. Operations
does not see raw error traces; show a friendly retry/contact-WFM message.

## Reviewer actions

The reviewer screen filters:

```powerfx
SortByColumns(
    Filter(
        WFM_AbsenceRequests,
        RequestStatus.Value = "Submitted" || SyncStatus.Value = "Failed"
    ),
    "SubmittedUTC",
    SortOrder.Ascending
)
```

`btnApprove.OnSelect` calls `WFM_Review_Absence` with a compact JSON payload:

```powerfx
WFM_Review_Absence.Run(
    JSON(
        {
            contractVersion: "1.0.0",
            requestId: galReviewQueue.Selected.RequestID,
            decision: "APPROVE",
            reviewComment: Trim(txtReviewComment.Text),
            reviewerUpn: gblUserEmail
        },
        JSONFormat.Compact
    )
)
```

Reject uses `decision: "REJECT"`. The flow verifies reviewer membership and
never trusts screen visibility or the passed email alone.

## Responsive rules

- `conMainWidth.Width = Min(App.Width - 48, 1280)`
- `conColumns.LayoutDirection` becomes vertical below 900 px.
- form column `FillPortions = 2`; summary `FillPortions = 1`.
- all controls have a minimum 40 px touch height.
- the summary follows the form on narrow screens.
- keep request state visible as text, not color alone.
