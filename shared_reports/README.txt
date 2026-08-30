WFMHUB SHARED REPORTS
=====================

This folder is for workbooks you create and share manually with management.
Generated XLSX/PDF files are intentionally not committed to the public repo.

Included builders
-----------------

1. Bonus management proposal
   - Rebuilds the existing Bonus workbook in the WFMHub design.
   - Shows a management scenario while technically blocking payroll release
     until policies, eligibility and input data are validated.
   - Keeps all thresholds, formulas and controls visible.

2. PCS management 3-hour template
   - Uses the exact logic from the original TOLEARN/PCS Report.xlsx:
       PCS Average      = average of valid Question 1 scores 1 to 5
       <=3 / >3         = counts of valid Question 1 scores
       Participation    = nonblank Question 1 / PCSStatus=1
   - Keeps Valid Response Rate separate from Participation.
   - Adds a real timestamp-bounded three-hour interval.
   - Removes the broken external Actions Rate link.

Build from the portable command line
------------------------------------

  runtime\python.exe -m wfmhub --home . shared-report bonus "C:\path\Bonus.xlsx"

  runtime\python.exe -m wfmhub --home . shared-report pcs "C:\path\PCS Report.xlsx" --date 2026-08-31

The source workbook is read-only. WFMHub writes a fresh workbook here.
