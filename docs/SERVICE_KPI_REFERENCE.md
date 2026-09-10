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
- Volume Variance = total entered - forecast
- AHT = handled seconds / answered

Because `total entered = A + B`, the implemented TSL denominator is
`total entered - lost within SLA`. Lost calls below 5 seconds remain included.
All higher-grain percentages are ratios of summed counters; hourly percentages
are never averaged. Every Flash displays 00:00 through 23:00 and daily headline
counters reset at midnight.

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
- RSA BE includes exactly 43 supplied queues. Four of them are also explicitly
  present in the Ford NL Flash allowlist; this dual-Flash membership is
  intentional and is represented by the two exact allowlists.
- Ford NL includes the six NL/VL/DE queues displayed for that Flash.
- Ford FR/OEM includes exactly three APFR queues: Ford Assistance,
  Toyota/Lexus, and Chery Assistance.

Queue mapping has one primary data-model scope per queue, while a Flash
allowlist may intentionally overlap another Flash. The four RSA BE/Ford NL
overlap queues retain `Ford NL` as their primary scope and are counted in both
specified Flash views. The APFR Toyota/Lexus and Chery queues use the internal
`Ford FR` service scope and the user-facing Ford/Toyota/Chery designations.

No queue is admitted by substring, suffix, designation or inferred LOB. A
mapped Call-by-Call queue that is absent from the exact profile allowlist stays
out of that Flash.

Every RTM LOB sheet uses `No Show HC` as its absence callout. It counts only a
scheduled agent who has no observed presence and whose agent-specific evidence
proves a no-show. An agent who arrived late, left early, or is currently offline
after attending remains present. `Offline Now` is a separate operational alert.
`Possible No Show HC` means unknown or missing/stale evidence and is never added
to confirmed No Show HC or the automatic call count. Exact gap treatment remains
in Attendance Review. `Due HC = Present HC + No Show HC + Unknown HC`, while
Offline Now is a subset of Present HC.

The hourly and LOB summary tables expose both handled volume and handled-in-SL
volume. OEM additionally exposes Entered, Handled, Handled in SL, and TSL for
the combined OEM total and separately for Ford, Toyota/Lexus, and Chery.

The roster side of those joins is explicit: OEM = `OEM FR`, RSA Belgium =
`RSA FR` + `RSA VL`, Ford Netherlands = `Ford Dutch`, and RSA Netherlands =
`RSA NL`. The call-data cutoff and attendance checkpoint are displayed
separately. The attendance checkpoint follows the latest Agent Status evidence,
not workbook refresh time, and the configured RTA staleness limit applies to a
last known state. No loaded file or LOB aggregate can substitute for
agent-specific presence/absence evidence.

`ISSUES & DRIVERS` contains only actionable source failures and configured
queues with real demand below target. It identifies whether late-routed
contacts, lost contacts, or both are driving the miss. It cannot change queue
membership or the validated Storm formula.
