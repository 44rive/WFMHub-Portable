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
WFM manually replaces the six files in fixed SharePoint Source CSV folder
          │
          ▼
Desktop Excel > Data > Refresh All > Save the SAME shared workbook
          │
          ▼
OVERVIEW / PERFORMANCE update; COACHING remains human-owned
```

WFMHub does not need Excel open to calculate and publish feeds. It does **not**
refresh or replace a query-enabled shared workbook. WFM does a one-time Power
Query setup in desktop Excel, then moves the six fresh CSVs to the fixed
SharePoint source folder and presses **Data > Refresh All** in the same shared
workbook. See the [beginner setup guide](PCS_POWER_QUERY_SETUP.md) before adding
queries. The old Hub Sync command must not be used after that setup.

The selector feed keeps Period, LOB, Team Leader and Agent in separate columns.
Its query must load to `_PCS_FILTERS!A4` without reordering headers. The four
other Overview/coaching caches and PERFORMANCE likewise load to their existing
sheet positions. A connection-only or Data Model-only load will leave formulas,
cards and charts unchanged.

## Human roles

- WFM: load sources, run PCS refresh, copy all six CSVs to SharePoint, refresh
  the one shared workbook in desktop Excel, check results and save.
- Team leaders / Quality: filter performance and maintain coaching status,
  owner, date, action, notes, and Call ID in the coaching table.
- Management: use overview LOB/agent tables and period comparisons; do not edit
  hidden feed/calculation sheets.

Rebuilding the permanent tracker is exceptional. A workbook containing Excel
queries/connections is protected from automatic template rebuilds, even if a
future Hub version has a newer tracker template. Keep one shared canonical
tracker rather than generating a new file for each delivery interval.

## One shared SharePoint file

1. On the WFM work PC, sync the restricted WFM SharePoint document library with
   OneDrive. Open its local folder in File Explorer. Do **not** move the Hub's
   SQLite database into SharePoint.
2. In `WFMHub.cmd`, choose **PCS Report & Coaching > Set permanent PCS workbook
   path**. Paste the local synced folder path or the full path ending in
   `PCS Live Tracker.xlsx`. A browser `https://` link is not a writable path.
3. Run **Update PCS data** once and copy the six current CSVs from the Hub's
   displayed `Feed/PCS` folder into a fixed `Source CSV` folder in that synced
   SharePoint library. Copy the existing tracker with its real coaching to the
   library once. Follow [PCS_POWER_QUERY_SETUP.md](PCS_POWER_QUERY_SETUP.md)
   to connect the six CSVs to the six existing presentation sheets. Keep a
   backup during setup. Set the dedicated PCS path to the **existing** shared
   workbook only after it has been moved there.
4. Check that OneDrive reports the file is synced, then share **that one
   SharePoint file link** with Team Leaders and Quality. Keep the library's
   permissions restricted to the intended WFM collaborators.
5. For future updates, add extracts, run **Update PCS data**, replace all six
   SharePoint CSVs, wait for sync, then open the same shared workbook in desktop
   Excel, **Refresh All**, and save. The team keeps the same link and coaching
   rows; no new workbook is sent.

This path setting is independent of `paths.reports`, so RTM, attendance,
realisations, and other generated workbooks remain local. The Hub's `Feed/PCS`
folder can move between portable releases; Power Query must point to the fixed
SharePoint `Source CSV` folder instead. Avoid simultaneous coaching edits
during the brief Excel refresh/save window.

## Troubleshooting

- Data feeds updated but cards/charts old: replace all six SharePoint CSVs,
  wait for OneDrive sync, then use **Data > Refresh All** in desktop Excel.
- Hub Sync says workbook has queries: this is expected protection; use Excel
  Refresh All instead.
- Power Query says source is missing: check the synced SharePoint CSV path, not
  the versioned Hub `Feed/PCS` path.
- Dates look numeric: use the table's date column format; source CSV contains
  ISO `YYYY-MM-DD` values.
