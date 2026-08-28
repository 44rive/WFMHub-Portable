# WFM PCS Daily Report

> **Published copy.** Layout in this folder:
> `docs/` proposal, pivot spec, setup guide, data findings ·
> `tools/` the refresh job and the two builder scripts ·
> `report/` the built workbook ·
> `data/` the pcs_agent_day export it was built from.
>
> The working copy lives in `~/WFM-PCS-Daily/`. Note `report/` is named that
> deliberately: the repository's root `.gitignore` excludes any folder called
> `output/`, so a workbook placed there would never be committed.
>
> Built from `TOLEARN/PCS Report.xlsx` (39,982 legs, 12-14 July 2026).
> Every headline measure reconciles exactly to that extract.
> **Two blockers before go-live - see `docs/DATA_FINDINGS.md`.**

A standalone daily/3-hourly Post-Call Survey report for Operations, built on the governed
`pcs_agent_day` export from WFMHub Portable.

**Nothing here modifies WFMHub-Portable or any source extract.** The refresh job calls only the
hub's `refresh` and `export` commands, both of which are read-only against your extract files.

---

## Files

| File | What it is |
|---|---|
| **`PCS_Ops_Proposal.md`** | The document you present. Definitions, caveats, cadence, what needs agreeing. Start here. |
| **`PCS_Pivot_Model_Spec.md`** | Technical build sheet — columns, DAX measures, pivot layouts, verification checks. |
| **`PCS_Setup_ClickByClick.md`** | Every click of the one-off build, then the daily routine. |
| **`PCS_Refresh_3h.cmd`** | The unattended refresh + export job. Edit one line, schedule it. |
| **`build_pcs_workbook.py`** | Fallback builder — produces a complete static workbook without Power Pivot. |
| `data/current/` | Holds **exactly one** CSV. The workbook reads this folder. |
| `data/archive/` | Timestamped previous exports, pruned after 60 days. |
| `logs/` | One log per refresh run, pruned after 60 days. |
| `report/` | The built workbook, ready to send. |

## Quick start

```
1. Edit WFMHUB_HOME at the top of PCS_Refresh_3h.cmd
2. Double-click PCS_Refresh_3h.cmd
3. Follow PCS_Setup_ClickByClick.md, Part A
```

---

## The four things that matter

**1. PCS-enabled = inbound AND `PostCallSurveyMode = 2`.**
It means *in scope to be surveyed*, not *was surveyed*. There is no handle-time condition.

**2. It counts legs, not interactions.**
A transferred call is two PCS-enabled legs but only one possible survey. Response rate is a
floor. This is disclosed in the proposal rather than hidden.

**3. Every ratio is a total over a total.**
Never an average of averages. On a realistic team shape that difference is worth **1.26 points** —
see section 4b of the proposal. The ten pre-divided ratio columns are removed from the data model
so nobody can drag one into a pivot by accident.

**4. Blank is never zero.**
No responses shows blank. A zero would read as poor performance and would wrongly pull the team
average down.

---

## The two operational traps

### ⚠️ `data/current/` must hold exactly one CSV

The Power Query folder source combines everything it finds. A second file doubles every count
**silently** — response rate still looks plausible because both halves double. The refresh job
warns when it sees more than one, and the query deduplicates on `agent_day_key` as a second line
of defence. Verification check 5 in the spec is the routine detector.

### ⚠️ The hub rebuilds marts in full

`mart.agent_pcs_day` is emptied and rebuilt on every refresh, so it holds **only the period last
refreshed**. The job therefore refreshes **1st of last month → today** on every run — narrow that
window and you lose MTD and rolling-7 from the report.

---

## Open question before go-live

The model gates on `PostCallSurveyMode`. The extract also carries a separate **`PCSStatus`**
field which the hub ingests but never uses. If `PCSStatus` is what records whether a survey was
actually *sent*, the response-rate denominator should be reviewed with whoever owns the Storm
configuration. Everything else can be agreed without it.
