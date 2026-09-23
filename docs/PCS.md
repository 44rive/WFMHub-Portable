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

## One shared SharePoint file

1. On the WFM work PC, sync the restricted WFM SharePoint document library with
   OneDrive. Open its local folder in File Explorer. Do **not** move the Hub's
   SQLite database into SharePoint.
2. In `WFMHub.cmd`, choose **PCS Report & Coaching > Set permanent PCS workbook
   path**. Paste the local synced folder path or the full path ending in
   `PCS Live Tracker.xlsx`. A browser `https://` link is not a writable path.
3. Close the old local tracker, run **Update PCS data** once, then close the
   shared tracker in desktop Excel and
   choose **Sync PCS workbook**. If the target file was absent, the Hub first
   copies the existing local tracker, including coaching; the local original
   remains as a backup. If a file already exists at the target, the Hub uses it
   and does not replace it merely because the path changed.
4. Check that OneDrive reports the file is synced, then share **that one
   SharePoint file link** with Team Leaders and Quality. Keep the library's
   permissions restricted to the intended WFM collaborators.
5. For future updates, add extracts, run **Update PCS data**, close the shared
   workbook, run **Sync PCS workbook**, and wait for OneDrive to sync. The team
   keeps the same link and coaching rows; no new workbook is sent.

This path setting is independent of `paths.reports`, so RTM, attendance,
realisations, and other generated workbooks remain local. OneDrive/Excel may
refuse the write if the shared file is locked or a sync conflict is active;
close it and retry rather than creating a second report. Avoid simultaneous
coaching edits during the brief sync window.

## Troubleshooting

- Data feeds updated but cards/charts old: close Excel and run **Sync PCS workbook**.
- Sync says the old workbook has queries: run **Update PCS data** with this
  release once to migrate it, then sync.
- WFMHub cannot replace tracker: close Excel and wait for sync. Coaching is
  preserved from the existing tracker before replacement.
- Dates look numeric: use the table's date column format; source CSV contains
  ISO `YYYY-MM-DD` values.
