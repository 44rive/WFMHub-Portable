# WFM PTO & Away Power Platform implementation kit

This directory contains the repo-owned implementation assets for the internal
`WFM PTO & Away` Canvas App. English is the default display language. French is
available from the header toggle.

The architecture and business authority are documented in
`docs/PTO_AWAY_APP_IMPLEMENTATION.md`. Do not build directly from the older
feasibility document.

## What is in this kit

| Folder | Purpose |
|---|---|
| `contracts` | Exact SharePoint columns and environment variables |
| `office-scripts` | Governed, idempotent writer for the unchanged FTE workbook |
| `powerfx` | Canvas App control map and copy/paste formulas |
| `flows` | Action-by-action Power Automate build specifications |
| `locales` | English and French interface copy |
| `tests` | UAT and integration acceptance cases |

## Production boundary

The repository owns the design, validation contract, formulas, script and test
evidence. The completed app and flows must be created inside the approved
Allianz Power Platform environment because connections, SharePoint permissions,
Canvas App metadata and publishing are tenant-owned.

After tenant construction, export the unmanaged development Solution and store
the backup outside the runtime WFMHub folders. Do not commit credentials,
connection tokens, populated employee lists or request data.

## Required tenant components

- restricted WFM SharePoint site;
- `WFM_Agents` list;
- `WFM_AbsenceRequests` list;
- `FTE Count.xlsx` in a restricted document library;
- Excel Online (Business), SharePoint and Office 365 Users connections;
- Office Scripts enabled by the tenant administrator;
- a maintained WFM functional owner and reviewer group.

## Build order

1. Run the local FTE workbook preflight.
2. Create the two SharePoint lists exactly as specified.
3. Install `WFM_Write_Absence.ts` in Excel on the web.
4. Create the solution, connection references and environment variables.
5. Build the flows in the order listed in `flows/BUILD_SPEC.md`.
6. Build the Canvas App from `powerfx/APP_BUILD.md`.
7. Execute every case in `tests/ACCEPTANCE.md`.
8. Export a Solution backup, publish, share and record the functional owner.

No app request is allowed to edit `tblFTEAgents`. Operations never writes to
the workbook directly.
