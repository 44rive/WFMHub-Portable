# Bonus Management workbook

This is one permanent monthly working file, not a newly generated attachment
for every refresh. Its editable sheets follow the supplied Bonus Matrix v1.2:
`Policy_Decisions`, `KPI_Config`, and `Raw_Data`. Results, analyses, and the
dashboard recalculate inside Excel. The WFMHub design is presentation only;
the KPI tiers, policy choices, VOC malus, and payout formula remain visible in
the workbook.

The initial KPI configuration reproduces all six populations in the uploaded
v1.2 matrix: Ford GER, Ford Dutch, OEM FR, RSA FR, RSA NL, and RSA VL (42 KPI
rows). It does not collapse them into combined reporting LOBs. AHT and PCS
participation thresholds differ by population; edit `KPI_Config` when the
approved compensation policy changes.

1. Build Bonus Management once from the Hub. It seeds the v1.2 KPI
   rules when no imported matrix exists. Existing workbook inputs are kept.
2. In `Policy_Decisions`, select and validate the monthly governance choices.
   A payout should not be released while any policy is unvalidated.
3. Paste monthly agent values into `Raw_Data`, keeping Client ID as text.
   Set Period to `YYYY-MM`; review Population, Team Lead, and Ops Manager.
   Mark Data Status `VALIDATED` only after checking the row.
4. Check `Control_Checks`, `Results`, `KPI_Analysis`, and
   `Team_Lead_Analysis`. Set the month in the Dashboard Period cell.
5. Send a copy of the checked workbook to management. Keep the canonical
   working file; do not overwrite it with a new generated report.

The template has 1,000 prepared input/result rows. If the workforce exceeds
that, increase `BONUS_INPUT_ROWS` in `bonus.py` and migrate the workbook while
preserving the editable sheets. Excel must recalculate formulas after a paste;
Python does not calculate Excel formulas or certify payroll payouts.
