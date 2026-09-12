WFMHUB POWER BI PROJECT
=======================

WHAT WFMHUB DOES
----------------
Run UPDATE in WFMHub. The Hub ingests only changed extracts, upgrades the same
SQLite database when needed, calculates governed marts, and replaces the CSVs
in Feed\PowerBI. POWERBI_MANIFEST_CURRENT.csv is written last, so it is the
refresh receipt.

WHAT YOU DO
-----------
1. Run UPDATE > All sources once in WFMHub.
2. Double-click POWERBI.cmd, or choose ANALYZE > Open Power BI dashboard.
3. Power BI Desktop opens the complete seven-page WFMHub BI.pbip project.
4. Choose Home > Refresh. The HubRoot parameter already points to this WFMHub.
5. Save a PBIX copy only when you want to publish or share a locked snapshot.

You do not create the queries, relationships, measures, slicers or pages. They
are already inside the project. Normal Hub releases preserve the installed
project; a newer project contract archives the old project before installing
the new one.

SOURCE AUTHORITY
----------------
- Agent Status: actual presence, RTM, attendance and gaps.
- StartEndTimes: planned shift boundaries.
- Verint Activities: final post-day absence/shrinkage codes only. It is never
  used to prove live presence.
- Call by Call: 15-minute service analysis; the separate PCS Excel workflow
  uses its own fixed feeds.
- Verint forecast: native interval volume; current supplied files are 15-minute.
  FTE Required remains blank when the source does not contain it.
- FTE Count: active employee scope, hierarchy and PTO/Away overlays.

IMPORTANT
---------
Power BI is the presentation layer. Do not recreate business formulas in Power
Query. Ratios must use the supplied additive counters and measures. Do not
average per-row percentages.

POWER BI PAGES
--------------
Daily Command, SL Drivers, Staff Prep, Realisation, Schedule Integrity,
Forecast Accuracy and Data Readiness. PCS is intentionally not imported.
