# PCS workflow and logic

## What a call leg is

A customer contact can pass through queues, treatments, transfers, or agents.
Each handled piece is a call leg. The Call-by-Call extract can therefore contain
several rows for one customer interaction. WFMHub keeps both a stable call-leg
key and an interaction key so a report can use the correct grain.

PCS agent performance is based on deduplicated, eligible inbound legs according
to the effective PCS rules. It does not assume every interaction produced a
valid survey response.

## Stored components

Per agent and day, the Hub stores:

- total/handled/inbound/outbound/transferred legs;
- talk, hold, wrap, and handle seconds;
- survey-mode eligible inbound legs;
- configured PCS-status legs;
- nonblank participation answers;
- valid score count and score sum;
- low-score and top-box counts;
- invalid nonblank scores, status-with-blank-answer, and answer-without-status
  exceptions.

PCS average and participation are ratios of those sums. Team, LOB, period, and
month views sum components first and divide once.

## Permanent tracker pipeline

```text
new FTE + Call-by-Call extracts
          │
          ▼
WFMHub PCS refresh → SQLite PCS mart
          │
          ▼
fixed Feed/PCS CSV files replaced atomically
          │
          ▼
Close Excel; WFMHub > PCS > Sync PCS workbook
          │
          ▼
OVERVIEW / PERFORMANCE update; COACHING remains human-owned
```

WFMHub does not need Excel open to calculate and publish feeds. To put the new
data into the permanent tracker, close Excel and choose **Sync PCS workbook**.
The sync opens Excel briefly, copies the six CSV feeds into ordinary tables,
recalculates, and saves. It never edits the human-owned coaching action table.

The selector feed keeps Period, LOB, Team Leader and Agent in separate columns.
Sync updates those columns in place, so the previous Power Query table
delete/recreate cycle cannot move a month into the LOB dropdown. The first
update with this release migrates the old tracker once and copies keyed coaching
actions forward. Keep the archived pre-upgrade copy until you verify them.

## Human roles

- WFM: load sources, run PCS refresh, check feed manifest and sync the
  workbook before publishing.
- Team leaders / Quality: filter performance and maintain coaching status,
  owner, date, action, notes, and Call ID in the coaching table.
- Management: use overview LOB/agent tables and period comparisons; do not edit
  hidden feed/calculation sheets.

Rebuilding the permanent tracker is exceptional. The builder reads coaching
rows from the current tracker before replacement and restores them by Coaching
Key. Keep one shared canonical tracker rather than generating a new file for
each delivery interval.

## Troubleshooting

- Data feeds updated but cards/charts old: close Excel and run **Sync PCS workbook**.
- Sync says the old workbook has queries: run **Update PCS data** with this
  release once to migrate it, then sync.
- WFMHub cannot replace tracker: close Excel and wait for sync. Coaching is
  preserved from the existing tracker before replacement.
- Dates look numeric: use the table's date column format; source CSV contains
  ISO `YYYY-MM-DD` values.
