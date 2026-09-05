# Service KPI business reference

Service actuals come only from mapped Storm Call-by-Call queue entries. Forecast
comes only from Verint Forecast. APBE, APFR and APDE are not discovered,
refreshed, required, or calculated.

## Queue-entry counters

One customer interaction can contain several call legs. Storm `Total Entered`
counts entries into its displayed queues, so WFMHub counts every mapped inbound
queue leg. An outbound companion leg does not count. If a transfer enters a
second mapped queue, Storm and WFMHub both count another queue entry.

- `total entered` = mapped inbound queue entries
- `response time` = Total Queue Wait Time + Ringing Duration
- `short abandon` = unanswered queue entry with response time below 5 seconds
- `SLA offered` = total entered - short abandons
- `abandoned within target` = non-short unanswered queue entry with response
  time below 20 seconds
- `answered within target` = routed queue entry with response time below 20
  seconds
- `handled seconds` = inbound talk + hold + wrap seconds on routed queue entries

The 5-second and 20-second boundaries are configured in
`config\wfm_rules.toml`.

## Reported formulas

The reports use the Storm dashboard business reference:

- TSL = answered within target / SLA offered
- Routed Rate = answered / total entered
- Deviation = total entered / forecast
- AHT = handled seconds / answered

All higher-grain percentages are ratios of summed counters. Hourly percentages
are never averaged. The separate `abandoned_within_target` counter remains in
the clean model for diagnosis, but the screenshots prove it is not removed
from the SLA denominator.

The OEM screenshot is a reproducible example: Ford contributes 85 in SLA from
91 SLA-offered calls, Chery 5/5, and Toyota/Lexus 159/180. The aggregate is
`249/276 = 90.22%`, exactly the displayed Storm value. The explicit
Call-by-Call reference in `Conso CBC FLASH - VA.xlsx` confirms that ringing is
part of the threshold clock.

The defensive `MAX(answered, total entered)` pattern seen in the old OEM workbook is
not required because the governed mart enforces the business invariant
`answered <= total entered`. Any violation is raised as a
data-quality issue instead of being hidden by a spreadsheet formula.

## Flash queue scope

- RSA NL and RSA BE include RSA/Allianz queues and exclude Provider queues.
- Ford NL includes only `APBN_AMS_MOBILITY_Ford_Assistance_NL` and
  `APBN_AMS_MOBILITY_Ford_Dealers_NL`.
- OEM includes only the APFR Ford, Chery, and Toyota/Lexus platform queues.
- Other mapped Provider and regional Ford queues remain available in clean
  data and queue controls but are marked `Used By Flash = NO`.
