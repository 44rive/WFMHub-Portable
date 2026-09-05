# Service KPI business reference

Service actuals come only from mapped Storm Call-by-Call interactions. Forecast
comes only from Verint Forecast. APBE, APFR and APDE are not discovered,
refreshed, required, or calculated.

## Interaction counters

One customer interaction can contain several call legs, for example queue,
agent and transfer legs. WFMHub deduplicates at interaction + mapped comparison
scope so a transfer does not become a second offered contact. A mapped inbound
interaction with a handled agent leg counts as answered.

- `raw offered` = unique mapped inbound interactions
- `response time` = Total Queue Wait Time + Ringing Duration
- `short abandon` = unanswered interaction with response time below 5 seconds
- `business offered` = raw offered - short abandons
- `abandoned within target` = non-short unanswered interaction with response
  time below 20 seconds
- `answered within target` = answered interaction with response time below 20
  seconds
- `handled seconds` = inbound talk + hold + wrap seconds on handled legs

The 5-second and 20-second boundaries are configured in
`config\wfm_rules.toml`.

## Reported formulas

The reports use the Storm dashboard business reference:

- TSL = answered within target / (business offered - abandoned within target)
- Service availability = answered / business offered
- Deviation = business offered / forecast
- AHT = handled seconds / answered

All higher-grain percentages are ratios of summed counters. Hourly percentages
are never averaged. The technical gross methods—answered within target / raw
offered and answered / raw offered—remain available in
`config\metric_catalog.toml` for audit or a future effective-dated policy
change, but they are not the displayed Flash methods.

This reproduces the Storm export fields `Offered_calls (w/o short calls)`,
`Answered_Calls <= 20s`, and `Abandoned_Calls (w/o s.c.) <= 20s`. The explicit
Call-by-Call reference in `Conso CBC FLASH - VA.xlsx` confirms that ringing is
part of the threshold clock.

The defensive `MAX(answered, business offered)` seen in the old OEM workbook is
not required after interaction deduplication because the governed mart enforces
the business invariant `answered <= raw offered`. Any violation is raised as a
data-quality issue instead of being hidden by a spreadsheet formula.
