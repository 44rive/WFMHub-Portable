WFMHUB CLEAN DATA FEED

WFMHub writes clean exports here when you choose "Export clean data".
Each requested dataset gets its own folder and a manifest.

Every source refresh also updates fixed-name collaboration feeds:
  Feed\PCS\PCS_AGENT_DAY_CURRENT.csv
  Feed\PCS\PCS_COACHING_OPPORTUNITY_CURRENT.csv
  Feed\PCS\PCS_SCOPE_CURRENT.csv
  Feed\Absenteeism\ABSENCE_AGENT_DAY_CURRENT.csv
  Feed\Absenteeism\ABSENCE_COMPONENT_CURRENT.csv
  Feed\Absenteeism\ABSENCE_REVIEW_CASE_CURRENT.csv

These names stay the same so a shared Excel workbook can use them as a Power
Query source. Each CSV is written completely before the previous copy is
replaced, so Excel never sees a half-written feed.

The original FTE, Storm, and Verint extracts are never moved or changed.
Agent ID and Case/Coaching Key are the matching fields. Display names are not
used as keys.
