WFMHUB POWER BI STARTER KIT
===========================

WHAT WFMHUB DOES
----------------
Run UPDATE in WFMHub. The Hub ingests only changed extracts, upgrades the same
SQLite database when needed, calculates governed marts, and replaces the CSVs
in Feed\PowerBI. POWERBI_MANIFEST_CURRENT.csv is written last, so it is the
refresh receipt.

WHAT YOU DO ONCE IN POWER BI DESKTOP
------------------------------------
1. Get data > Text/CSV and import the CSV files in Feed\PowerBI. Do not import
   POWERBI_MANIFEST_CURRENT.csv as a business fact.
2. Transform data. Confirm Date fields are Date; start/end fields are Date/Time;
   counters and minutes are Whole/Decimal Number; flags are True/False.
3. Create one-to-many, single-direction relationships:
     DimDate[Date] -> every fact Date column
     DimEmployee[Agent ID] -> attendance, gap, PCS, coaching, final ABS Agent ID
     DimLOB[LOB] -> facts that expose LOB
     DimQueue[Queue] -> service Queue
   Forecast Queue can repeat across forecast files/scopes, so use its Service
   Scope/Date for shared slicing and its own Queue field for forecast detail.
   Keep fact-to-fact relationships disabled. Use dimensions for slicers.
4. Import WFMHub-Premium-Theme.json from View > Themes > Browse for themes.
5. Create a blank _Measures table and add the measures in WFMHub-Measures.dax.
6. Build the seven pages from docs\POWERBI_PREMIUM_DESIGN_V1.md and the supplied
   PNG prototypes. Use a 16:9 canvas and Sync slicers for Date, LOB and Team.
7. Save the PBIX. From now on: WFMHub UPDATE, then Refresh in Power BI.

SOURCE AUTHORITY
----------------
- Agent Status: actual presence, RTM, attendance and gaps.
- StartEndTimes: planned shift boundaries.
- Verint Activities: final post-day absence/shrinkage codes only. It is never
  used to prove live presence.
- Call by Call: service and PCS.
- Verint forecast: native interval volume; current supplied files are 15-minute.
  FTE Required remains blank when the source does not contain it.
- FTE Count: active employee scope, hierarchy and PTO/Away overlays.

IMPORTANT
---------
Power BI is the presentation layer. Do not recreate business formulas in Power
Query. Ratios must use the supplied additive counters and measures. Do not
average per-row percentages.
