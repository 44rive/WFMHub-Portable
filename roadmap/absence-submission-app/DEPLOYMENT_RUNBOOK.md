# Deployment runbook — beginner route

This is the order to follow in the Microsoft tenant. Do not start with the app
screen. Establish the data and retry path first.

## Before starting

Confirm that you can:

1. open the restricted WFM SharePoint site;
2. create SharePoint lists there;
3. create a Solution in the selected Power Platform environment;
4. use SharePoint, Excel Online (Business), Office 365 Users and Microsoft 365
   Groups connectors;
5. see the **Automate** tab in Excel on the web;
6. identify a maintained WFM account that will own the connections.

If one item is unavailable, record it before building. Do not replace a blocked
standard connector with a personal account or an unapproved external service.

## Step 1 — prepare the workbook

1. Copy the current `templates/FTE Count.xlsx` into the restricted WFM document
   library.
2. Keep the filename `FTE Count.xlsx`.
3. Do not rename its sheets, tables or headers.
4. Run locally:

   ```bat
   python tools\check_power_platform_absence_contract.py "path\to\FTE Count.xlsx"
   ```

5. Resolve duplicate Client IDs and Leavers without End dates before the pilot.

## Step 2 — create SharePoint lists

Follow `contracts/SHAREPOINT_SETUP.md` exactly. First create the internal column
name, then change only its display label if needed. Index Client ID and workflow
status columns.

## Step 3 — install the Office Script

Follow `office-scripts/README.md`. Test INSERTED, NO_CHANGE, UPDATED and ERROR
against a workbook copy.

## Step 4 — create the Solution

1. Open `make.powerapps.com` and select the approved environment.
2. Open **Solutions > New solution**.
3. Display name: `WFM Absence Management`.
4. Unique name: `WFMAbsenceManagement`.
5. Publisher prefix: `wfm`.
6. Create connection references for the four approved connectors.
7. Create the environment variables from `contracts/sharepoint-lists.json`.

## Step 5 — build and test flows

Build the flows in `flows/BUILD_SPEC.md` order. Test each independently:

1. sync a synthetic roster batch;
2. submit without touching Excel;
3. approve and create Pending sync;
4. sync the exact workbook row;
5. advance Away status using test dates.

Do not connect the production workbook until the isolated copy passes.

## Step 6 — build the Canvas App

1. Add a responsive tablet Canvas App to the Solution.
2. Set English as the default.
3. Build the auto-layout container tree from `powerfx/APP_BUILD.md`.
4. Use the exact stable control names.
5. Connect the two lists and two interactive flows.
6. Apply the WFMHub palette and compare with the approved PNG at 100% zoom.
7. Test desktop and tablet widths before adding more decoration.

## Step 7 — UAT

Execute `tests/ACCEPTANCE.md`. A failed workbook lock test, ambiguous-ID test,
unauthorized-reviewer test or WFMHub integration test blocks production.

## Step 8 — publish and hand over

1. Publish the app.
2. Share run access with the Operations group.
3. Share review access only with the WFM reviewer group.
4. Confirm the functional owner and a backup owner.
5. Export an unmanaged development backup and a managed production package if
   the environment policy uses managed deployment.
6. Record connection ownership, workbook path, list URLs and the retry procedure
   in the restricted WFM operating guide.

## Daily operating model

Operations searches a Client ID and submits. WFM reviews. Approved requests sync
to Excel automatically. WFMHub keeps ingesting the same established workbook.
If Excel is locked, WFM retries the failed sync; Operations does not resubmit.
