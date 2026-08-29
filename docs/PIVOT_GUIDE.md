# Creating PivotTables from WFMHub reports

WFMHub creates governed Excel Tables; it does not create PivotTables for you.
You can arrange those tables without touching the extracts or the SQLite hub.

## The basic clicks

1. Open one of the four generated workbooks.
2. Open the detailed sheet named below.
3. Click any populated cell inside the striped table.
4. On Excel's ribbon, click **Insert**, then **PivotTable**.
5. Choose **New Worksheet**, then click **OK**.
6. Drag fields into **Rows**, **Columns**, **Values**, and **Filters**.
7. If Excel says **Count of** a numeric field, open **Value Field Settings**
   and change it to **Sum** or **Average**, as described below.

If a sheet says “No rows for this period,” there is no table to pivot. Check
`SOURCE_HEALTH`, the selected dates, and `DATA_QUALITY`, then refresh.

## Daily staffing gaps by LOB and language

Use the newest workbook in `output\operations`, sheet `STAFFING_GAPS`.

- Rows: **LOB**, then **Language**.
- Columns: **Interval Start** if you want the intraday curve.
- Values: **Sum of Required FTE**, **Sum of Available FTE**, and **Sum of
  Staffing Gap FTE**.
- Filters: **Business Date**, **Staffing State**, **Evidence Basis**.

Treat `DATA_MISSING` and `DATA_PARTIAL` as unknown evidence. Their gap and
variance are deliberately blank; do not replace them with zero.

## APDE service state by LOB

Use `output\operations`, sheet `SERVICE_LEVEL`.

- Rows: **LOB**, then **Language**.
- Columns: **Interval Start** or **Interval Label**.
- Values: **Sum of Offered**, **Sum of Answered**, **Sum of Answered Within
  Target**, **Sum of Short Abandoned**, and **Sum of Handled Seconds**.

Do not average interval percentages. Calculate the total after summing the
components:

```text
Gross SL              = Answered Within Target / Offered
Adjusted SL           = Answered Within Target / (Offered - Short Abandoned)
Service Availability  = Answered / Offered
AHT                    = Handled Seconds / Answered
```

Service availability is the availability of the service, never an agent
availability or adherence measure.

## Exact PCS by agent and month

Use `output\quality_pcs`, sheet `AGENT_MONTH` for a monthly PivotTable or
`AGENT_DAY` for daily detail.

- Rows: **Agent Name**.
- Columns: **Month Key** in `AGENT_MONTH`, or **Business Date** in `AGENT_DAY`.
- Values: sum the score and participation counters.
- Filters: **Team Leader**, **LOB**, **Language**, **Sample Flag**.

For a higher-grain result, divide summed counters:

```text
PCS average      = Sum of Score Sum / Sum of Valid Score Responses
PCS participation = Sum of Participation Responses / Sum of PCS Status Calls
```

Never average the already-calculated agent/day percentages. `PCS_LOGIC`
documents the exact Q1 eligibility, `<=3`/`>3` counts, and participation rule.

## Final absenteeism by agent and month

Use `output\absence`, sheet `AGENT_DAY`.

- Rows: **Agent Name**.
- Columns: **Business Date**. In Excel, right-click one date and choose
  **Group**, then **Months**, if you want a monthly agent view.
- Values: **Sum of Final Absence Hours** and **Sum of Planned Net Hours**.
- Filters: **Team Leader**, **LOB**, **Language**, **Final Ledger Status**.

Calculate the correct aggregate rate beside the PivotTable:

```text
Final absence rate = Sum of Final Absence Hours / Sum of Planned Net Hours
```

Do not average `Final Absence Rate`. Use `ACTIVITY_EVENTS` only for audit and
activity detail; it can contain overlapping evidence and is not the headline
numerator. Review `UNMAPPED_REVIEW` before using final payroll results.

## Add slicers

1. Click the PivotTable.
2. Click **PivotTable Analyze**, then **Insert Slicer**.
3. Choose fields such as LOB, Language, Team Leader, Status, or Source.
4. Click **OK** and place the slicers beside the PivotTable.

## If the PivotTable looks wrong

- **Count of …** instead of **Sum of …**: change Value Field Settings.
- Percentage above 100%: use summed numerator divided by summed denominator.
- Missing agents: check the FTE scope and `SOURCE_HEALTH`.
- Missing dates: check the selected period and newest source dates.
- Empty Daily Operations: StartEndTimes is required for the operational plan;
  Activities cannot replace it.
- Unexplained final absence: check `ACTIVITY_EVENTS`, `UNMAPPED_REVIEW`,
  `ACTIVITY_RULES`, and the rule version/hash.
