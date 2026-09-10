# Install and test the Excel writer

The production flow calls `WFM_Write_Absence.ts`. This is the only component
allowed to write the PTO and Away tables. It never writes `tblFTEAgents`.

## Install

1. Put the governed `FTE Count.xlsx` in the restricted WFM SharePoint library.
2. Open it in Excel on the web.
3. Select **Automate > New Script**.
4. Replace the example code with the complete contents of
   `WFM_Write_Absence.ts`.
5. Save it as `WFM_Write_Absence`.
6. Do not add a workbook button. The serialized flow is the production caller.

If the Automate tab or Run script action is unavailable, stop. Office Scripts
must be enabled by the tenant administrator before the project can use this
unchanged-workbook architecture.

## Test payloads

Create a copy of the blank FTE template for isolated testing. Never test against
the production roster first.

Approved full-day PTO:

```json
{
  "contractVersion": "1.0.0",
  "requestId": "728c0aa8-bb3d-4a50-a014-6578644152f0",
  "operation": "UPSERT",
  "requestType": "PTO",
  "clientId": "00123",
  "employeeName": "Jane Agent",
  "startDate": "2026-09-15",
  "endDate": "2026-09-16",
  "dayCoverage": "Full day",
  "startTime": "",
  "endTime": "",
  "awayType": "",
  "comment": "Approved by WFM",
  "today": "2026-09-10"
}
```

Open-ended Away:

```json
{
  "contractVersion": "1.0.0",
  "requestId": "03be47e2-54b5-4bdb-b72a-b580c20e470f",
  "operation": "UPSERT",
  "requestType": "AWAY",
  "clientId": "00123",
  "employeeName": "Jane Agent",
  "startDate": "2026-09-20",
  "endDate": "",
  "dayCoverage": "",
  "startTime": "",
  "endTime": "",
  "awayType": "Long sickness",
  "comment": "Open case",
  "today": "2026-09-10"
}
```

Pass the whole JSON object as one string parameter named `payloadJson`. Power
Automate displays the returned object as the dynamic output `result`.

## Expected behavior

- the first approved request inserts one row;
- the exact same request returns `NO_CHANGE`;
- a status transition updates the one exact row;
- more than one matching workbook row returns an ambiguity error;
- cancellation marks the row Cancelled and never deletes it;
- invalid dates, times, types, table names or headers return `ok = false`;
- Client ID is written as text;
- full-day PTO has blank time fields;
- PTO type is always PTO;
- future fixed-end Away is Planned, effective Away is Active and ended Away is
  Closed;
- open-ended Away is Active because that is required by current WFMHub ingestion.

## Local static check

From the repository:

```bash
tsc -p power-platform/wfm-absence-app/office-scripts/tsconfig.json
```

This catches TypeScript contract errors. Final execution still must be tested in
Excel on the web because the ExcelScript runtime is tenant-owned.

Run the mock workbook behavior tests:

```bash
tsc -p power-platform/wfm-absence-app/office-scripts/tsconfig.test.json
node power-platform/wfm-absence-app/office-scripts/run-tests.cjs
```
