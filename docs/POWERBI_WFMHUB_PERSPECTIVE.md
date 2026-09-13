# Power BI and WFMHub — current direction

Power BI is the personal WFM management cockpit; it does not replace WFMHub's
calculation engine or the operational Excel products.

- WFMHub ingests untouched extracts, applies roster scope and business rules,
  persists SQLite history, calculates facts and publishes fixed CSVs.
- Power BI imports `Feed\PowerBI`, relates dimensions, recalculates approved
  ratios from additive components, and presents the five-page WFM cycle.
- Excel remains the exact-row sharing channel. PCS remains one permanent
  collaborative Excel tracker and is outside Power BI.
- There is no direct SQLite/ODBC/raw connection in Power BI and no runtime AI.

The implemented design, semantic grains, exact pages and validation gate are in
`POWERBI_WFM_CYCLE.md`. The complete business authority is `../WFM_MASTER.md`.
