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
Excel Data > Refresh All
          │
          ▼
OVERVIEW / PERFORMANCE update; COACHING remains human-owned
```

WFMHub does not need Excel open to refresh data. Excel does not query SQLite.
Power Query reads the fixed CSV paths so the collaborative workbook can remain
the same file. The setup script installs query definitions once.

The selector feed has separate physical columns for Period, LOB, Team Key,
Team Leader, Agent Key, and Agent. The installer binds each dropdown to its
own column and validates that the resolved Team/Agent list stays inside its
selected LOB/team group. This prevents a month label from appearing in a LOB
list after Power Query refresh. On a newly generated tracker, run **Install /
repair Power Query** with Excel closed, then open it and use **Data > Refresh
All**. Do not regenerate the shared tracker for normal PCS updates.

## Human roles

- WFM: load sources, run PCS refresh, check feed manifest and refresh the
  workbook before publishing.
- Team leaders / Quality: filter performance and maintain coaching status,
  owner, date, action, notes, and Call ID in the coaching table.
- Management: use overview LOB/agent tables and period comparisons; do not edit
  hidden query/calculation sheets.

Rebuilding the permanent tracker is exceptional. The builder reads coaching
rows from the current tracker before replacement and restores them by Coaching
Key. Keep one shared canonical tracker rather than generating a new file for
each delivery interval.

## Troubleshooting

- Data feeds updated but cards/charts empty: open Excel, **Data > Refresh All**,
  wait for all queries, then recalculate.
- Query tables show a path error: run **Install / repair Power Query** once with
  the tracker closed.
- WFMHub cannot replace tracker: close Excel and wait for sync. Coaching is
  preserved from the existing tracker before replacement.
- Dates look numeric: use the table's date column format; source CSV contains
  ISO `YYYY-MM-DD` values.
