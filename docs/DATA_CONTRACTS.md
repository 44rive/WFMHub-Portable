# Data contracts

## Source rules

All sources are read-only. Headers are normalized for harmless punctuation and
case differences, but required business fields must be present. Failed files
remain visible in source health and do not replace the last successful version.

### FTE Count

The workbook template is `templates/FTE Count.xlsx` and contains:

- `tblFTEAgents`: Client ID, Status, Name, Team leader, Ops Manager, LOB,
  Market, Language, Location, City, FTE, End date if leaver;
- `tblFTEPTO`: approved/pending PTO intervals;
- `tblFTEAway`: long-absence intervals.

Client ID is text. Duplicate or blank IDs are quality errors. Only approved PTO
enters operational overlays; Away status/date logic is read from its record.

### Schedule and Activities

StartEndTimes supplies schedule date, Data Source ID/Client ID, agent name,
assignment, start, and end. Activities supplies dated activity intervals. The
parser distinguishes the variants from their headers/content, not merely the
filename.

### Agent Status and LILO

Agent Status requires agent identity, state, start/end or duration, and a
business date. Its exact state classification is governed by status rules.
LILO requires agent identity, business date, first login, and last logout.
Overnight logout values are adjusted only when their row evidence requires it.

### Call by Call

Each row is a call leg. Call ID/reference, treatment, direction, queue, agent,
timings, survey mode/status, and question fields are retained. A customer
interaction may have several legs; queue service and agent performance must use
the correct documented grain.

### Verint forecast and FTE requirement

Rows contain date, 15-minute period, Staff Type, Volume and available forecast /
requirement measures. Volume is additive. FTE/headcount values are levels.
Overlapping exports select the newest source row at the same Staff Type and
interval.

## Clean export rules

Every export has a data file and manifest containing period, row count,
generation timestamp, rule/catalog versions, and source lineage. CSV is the
preferred large-data format. XLSX exports are bounded convenience files, not a
database replacement.

Run `wfmhub export --help` or use **Export clean data** in the menu to see the
current dataset registry. The registry in `src/wfmhub/exports.py` is the code
contract; adding an export requires a description, deterministic ordering, and
a synthetic test.

## Sensitive data

Operational extracts and generated reports can contain personal and customer
data. They belong only in the configured local/source/output locations covered
by `.gitignore`. The public repository contains no real extracts. Synthetic
fixtures use invented names and IDs.
