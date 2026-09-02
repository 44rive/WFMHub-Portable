WFMHUB SHARED REPORTS
=====================

This folder is for workbooks you create and share manually with management.
The two official starting workbooks are included in the public release:

  Bonus_Management_Proposal.xlsx
  PCS_Management_3H_Template.xlsx

Other generated XLSX/PDF copies remain local and are not committed.

Included builders
-----------------

1. Bonus management proposal
   - Rebuilds the existing Bonus workbook in the WFMHub design.
   - Shows a management scenario while technically blocking payroll release
     until policies, eligibility and input data are validated.
   - Keeps all thresholds, formulas and controls visible.

2. Bonus KPI change case
   - Reconciles the source workbook's current modeled payout.
   - Opens on four simple presentation tabs with a 60-second talk track.
   - Explains the single-treatment rule for absence KPI versus proration.
   - Separates measured coverage from KPI attainment and compares LOB behavior.
   - Isolates participation, Extra PCS and completeness-gate sensitivities.
   - Keeps the agent trace and audit detail hidden but available for questions.
   - Includes a copy-ready management email.
   - Treats every scenario as a policy test, never a payroll instruction.

3. PCS management 3-hour template
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

  runtime\python.exe -m wfmhub --home . shared-report bonus-analysis "C:\path\Bonus.xlsx"

  runtime\python.exe -m wfmhub --home . shared-report pcs "C:\path\PCS Report.xlsx" --date 2026-08-31

The source workbook is read-only. WFMHub writes a fresh workbook here.
