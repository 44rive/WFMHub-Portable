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
- `lost within SLA` = unanswered queue entry with response time from 5 seconds
  up to, but not including, 30 seconds (Storm variable D)
- `SLA denominator` = total entered - lost within SLA
- `answered within target` = routed queue entry with response time below 30
  seconds (Storm variable C)
- `handled seconds` = inbound talk + hold + wrap seconds on routed queue entries

The 5-second and 30-second boundaries are configured in
`config\wfm_rules.toml`.

## Reported formulas

The reports use the Storm dashboard business reference:

- TSL = C / (A + B - D), where A is lost, B is connected, C is connected
  within SLA, and D is lost from 5 seconds to the SLA target
- Routed Rate = answered / total entered
- Deviation = total entered / forecast
- AHT = handled seconds / answered

Because `total entered = A + B`, the implemented TSL denominator is
`total entered - lost within SLA`. Lost calls below 5 seconds remain included.
All higher-grain percentages are ratios of summed counters; hourly percentages
are never averaged. Daily headline counters reset at midnight even when a
Flash hides pre-opening hourly rows.

The supplied Storm equation screen is the formula authority. The explicit
Call-by-Call reference in `Conso CBC FLASH - VA.xlsx` confirms that ringing is
part of the threshold clock.

The defensive `MAX(answered, total entered)` pattern seen in the old OEM workbook is
not required because the governed mart enforces the business invariant
`answered <= total entered`. Any violation is raised as a
data-quality issue instead of being hidden by a spreadsheet formula.

## Flash queue scope

The Storm screenshots are the sole authority for Flash membership. Exact queue
allowlists live under `flash_queues` in `config\service_profiles.toml`:

- RSA NL includes all 30 queues displayed in `TOLEARN\RSA NL.png`, including
  its Provider and RSA Ford-labelled rows.
- RSA BE includes all 36 queues displayed in `TOLEARN\RSA BE.png`, across FR,
  VL and EN rows, including Provider.
- Ford NL includes the six NL/VL/DE queues displayed for that Flash.
- OEM includes the four Ford FR activity queues plus the displayed Chery and
  Toyota/Lexus queues: six physical queues in total.

No queue is admitted by substring, suffix, designation or inferred LOB. A
mapped Call-by-Call queue that is absent from the exact profile allowlist stays
out of that Flash.
