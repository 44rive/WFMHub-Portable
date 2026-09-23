# Roadmap

## Absence submission app

`roadmap/absence-submission-app/` contains the approved future Power Apps /
Power Automate design for Operations to submit PTO and Away requests. It is not
part of the current WFMHub runtime.

The planned app uses English by default with French toggle, exact Client ID
lookup, duplicate-ID blocking, PTO/Away conditional validation, SharePoint
staging lists, one reviewed flow, and a governed Office Script that writes the
unchanged `tblFTEPTO` or `tblFTEAway` table contract in `FTE Count.xlsx`.

Before production, clean blank and duplicate Client IDs, create a restricted
WFM SharePoint library, test concurrency/retry behavior, publish through a Power
Platform Solution, and complete the acceptance checklist in that folder.

No roadmap artifact may become a runtime dependency until it has an explicit
implementation decision, security review, ownership, tests, and updated product
contract.
