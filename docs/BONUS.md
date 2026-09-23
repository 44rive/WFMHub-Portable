# Bonus Management workbook

This is one permanent monthly working file, not a newly generated attachment
for every refresh. Its editable sheets follow the supplied Bonus Matrix v1.2:
`Policy_Decisions`, `KPI_Config`, and `Raw_Data`. Results, analyses, and the
dashboard recalculate inside Excel. The WFMHub design is presentation only;
the KPI tiers, policy choices, VOC malus, and payout formula remain visible in
the workbook.

The KPI configuration covers all six populations in the uploaded v1.2 matrix:
Ford GER, Ford Dutch, OEM FR, RSA FR, RSA NL, and RSA VL. The 42 original
rules are the Tenured baseline. Another 42 Untenured rows are deliberately
blank: HR/Compensation must approve their targets and awards. Tenure is not a
third payout tier. `Raw_Data` has a Tenure Group dropdown for each agent; rows
without a valid tenure or approved targets show a review status and no final
payout. The Dashboard charts show payout and KPI attainment by population.

1. Build Bonus Management once from the Hub. It seeds the v1.2 KPI
   rules when no imported matrix exists. Existing workbook inputs are kept.
2. In `Policy_Decisions`, select and validate the monthly governance choices.
   A payout should not be released while any policy is unvalidated.
3. Paste monthly agent values into `Raw_Data`, keeping Client ID as text.
   Set Period to `YYYY-MM`; review Population, Tenure Group, Team Lead, and Ops Manager.
   Mark Data Status `VALIDATED` only after checking the row.
4. Check `Control_Checks`, `Results`, `KPI_Analysis`, and
   `Team_Lead_Analysis`. Set the month in the Dashboard Period cell.
5. Send a copy of the checked workbook to management. Keep the canonical
   working file; do not overwrite it with a new generated report.

The template has 1,000 prepared input/result rows. If the workforce exceeds
that, increase `BONUS_INPUT_ROWS` in `bonus.py` and migrate the workbook while
preserving the editable sheets. Excel must recalculate formulas after a paste;
Python does not calculate Excel formulas or certify payroll payouts.

To migrate from an earlier workbook, close `Bonus Management.xlsx` and build
Bonus Management once. WFMHub archives the previous copy and preserves populated
Policy, KPI, and Raw Data rows. Fill tenure and approve any Untenured KPI rows
before using the recalculated payout.
