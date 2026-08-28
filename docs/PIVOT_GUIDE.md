# Creating PivotTables from WFMHub reports

WFMHub creates clean Excel Tables. You create the PivotTable yourself, which
keeps the portable hub simple and lets you arrange the report your way.

## Example 1: absence hours by agent and month

1. Open the newest workbook under `output\absence`.
2. Open the `PIVOT_ABSENCE` sheet.
3. Click any cell inside the table.
4. On Excel's ribbon, click **Insert**.
5. Click **PivotTable**.
6. Excel should show the table name `tblPivotAbsence`.
7. Select **New Worksheet**, then click **OK**.
8. In PivotTable Fields, drag **Calendar Month** to **Rows**.
9. Drag **Agent** below Calendar Month in **Rows**.
10. Drag **Absence Hours** to **Values**.
11. If Excel says “Count of Absence Hours,” click it, choose **Value Field
    Settings**, select **Sum**, and click **OK**.
12. Drag **Team Leader**, **LOB**, **Language**, or **Location** to **Filters**.

You now have a monthly absence-hours PivotTable.

## Example 2: absence percentage

Do not average the `Absence Rate %` column across days. A correct month is:

```text
Sum of Absence Hours / Sum of Planned Net Hours
```

The easiest beginner method is:

1. Put **Sum of Absence Hours** in Values.
2. Put **Sum of Planned Net Hours** in Values.
3. Next to the PivotTable, divide the first result by the second.
4. Format the result as Percentage.

If you know PivotTable calculated fields, create one using:

```text
='Absence Hours' / 'Planned Net Hours'
```

Always verify that Excel is using sums, not counts or averages.

## Example 3: service performance

1. Open the newest workbook under `output\scorecard`.
2. Open `SERVICE_INTERVALS`.
3. Click inside `tblServiceIntervals` and insert a PivotTable.
4. Put **Date**, **Source System**, **LOB**, or **Language** in Rows/Filters.
5. Use sums of these components:
   - Offered
   - Answered
   - Answered Within Target
   - Short Abandoned
   - Handled Seconds
6. Calculate outside the PivotTable:

```text
Gross SL              = Answered Within Target / Offered
Adjusted SL           = Answered Within Target / (Offered - Short Abandoned)
Service Availability  = Answered / Offered
AHT                    = Handled Seconds / Answered
```

Do not average interval SL, availability, or AHT percentages.

## Example 4: one long KPI PivotTable

The `KPI_DAILY` sheet in the Executive Scorecard is useful when you want one
PivotTable for Service, Absence, Forecast, and PCS.

1. Insert a PivotTable from `tblKpiDaily`.
2. Put **Domain** and **KPI Name** in Rows.
3. Put **Calendar Month** or **ISO Week** in Columns.
4. Put **Source**, **LOB**, and **Language** in Filters.
5. For additive KPIs such as Offered or Absence Hours, use **Sum of Value**.
6. For ratio KPIs, use summed **Numerator** divided by summed **Denominator**.

## Add slicers

1. Click the PivotTable.
2. Click **PivotTable Analyze**.
3. Click **Insert Slicer**.
4. Choose fields such as LOB, Language, Team Leader, Location, Source, or KPI
   Name.
5. Click **OK** and place the slicers beside the PivotTable.

## If the PivotTable looks wrong

- “Count of …” instead of “Sum of …”: change Value Field Settings to Sum.
- Percent above 100%: check the numerator/denominator and active SL profile.
- Monthly percentage differs from the daily average: the monthly ratio-of-sums
  is the correct one.
- Missing agents: check FTE scope and `SOURCE_HEALTH`.
- Missing dates: check the latest source date and selected report period.
- Unexplained absence: check `ABSENCE_EVENTS`, `NOT_CORRECTED`, `VERINT_ONLY`, and the rule version.
