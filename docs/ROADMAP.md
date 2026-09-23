# Roadmap

## WFMHub-2 integration assessment

`roadmap/wfmhub2-integration/README.md` compares the current operational Hub
with the separate WFMHub-2 Phase 1 preview and proposes selective integration,
not a direct Git/database merge. No implementation or product migration has
been approved from that assessment.

## RTM Flash reassessment

`roadmap/rtm-flash-reassessment/README.md` records the proposed per-LOB Flash,
hourly/15-minute detail, Agent Status workforce reconciliation, unscheduled
coverage, live not-logged action wording, Issues & Drivers, and how no-shows
affect observed capacity without changing Verint requirements. It preserves
`Other Tasks` as BO. This is a design note, not active behavior; the proposed
report split awaits validation. The PCS Power Query setup is now documented
separately in [PCS_POWER_QUERY_SETUP.md](PCS_POWER_QUERY_SETUP.md).

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
