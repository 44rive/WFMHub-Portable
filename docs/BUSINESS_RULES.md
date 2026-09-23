# Business rules

This guide explains the model. Exact effective values live in the configuration
files named below.

## Service level

Call-by-Call legs are deduplicated first. For each configured Flash scope and
interval, WFMHub retains raw offered, answered, abandoned, short-abandoned,
answered-in-target, abandoned-in-target, and handled-seconds components.

The headline service-level method is selected from `metric_catalog.toml`.
Reports aggregate the components and then evaluate the ratio. They do not
average queue percentages. Exact queue inclusion and daily operating hours are
in `service_profiles.toml`; source queue mapping is in `queue_mapping.csv`.

Service Volume is observed queue demand. Verint forecast Volume and FTE
Requirement are separately mapped Staff Type series. Forecast levels such as
FTE are averaged across source 15-minute intervals when an hourly presentation
is required; Volume is summed.

## Attendance states

For each eligible scheduled agent-day:

1. clip approved PTO/Away to the shift;
2. build non-overlapping Agent Status intervals;
3. derive first/last positive presence and exact Logged Off/Unavailable gaps;
4. use LILO boundaries only when status evidence is insufficient;
5. evaluate only elapsed work time for a live day;
6. classify due/present/late/no-show/possible-no-show/early-leave;
7. preserve exact gap fragments for post-day reconciliation.

RTM calls an in-progress agent with explicit disconnected/not-seen evidence,
but does not count that person as a confirmed No Show. Confirmed No Show HC
requires a completed scheduled shift, no presence, and agent-specific LILO or
Agent Status evidence. Missing extracts remain Unknown / Possible No Show.

The RTM hourly headcounts are distinct-agent snapshots at the end of each
completed hour, or at the latest Agent Status checkpoint for the current
hour. Scheduled HC excludes PTO/Away. Logged HC requires observed Agent Status
or LILO presence. Productive HC includes Available and BO; those two HC
columns are subsets, not extra people to add to Productive HC. Unavailable HC
covers other AUX, Break, Lunch, and Unavailable. Future hours remain blank.
The Issues & Drivers sheet shows below-target queue-hours and named same-LOB
agents with logged-off, AUX, break/lunch, BO, or missing-evidence states
overlapping those hours. It is staffing context, not proof that a particular
person caused a queue result.

A late agent who appears is present and late, not absent. Early leave requires a
completed shift. Returning later creates an internal gap, not a permanent early
leave. Meal/Lunch and Break categories remain visible and are not attendance
gaps. Full-day PTO/Away suppresses callout; partial-day time off removes only
its covered interval.

## Attendance Review

The shift timeline overlays the published schedule and observed status. Gap
fragments are subtracted against finalized Verint Activities using the
configured match tolerance. Only unmatched residual fragments are reported.
There is no manual decision import: after Operations corrects Verint and a new
Activities extract is loaded, matched fragments disappear naturally.

Break and meal controls sum Agent Status timeline minutes on completed shifts
and compare them with configured allowances. They are operational alerts, not
an employee adherence score.

## Final absence and shrinkage

`wfm_rules.toml` classifies every Verint activity with flags for working,
planned, absence, vacation, unpaid, and shrinkage. Final measures intersect
activity intervals with scheduled work and respect PTO/Away overlays. The model
publishes both components and rates so every result can be reconciled.

Rows remain reviewable when an activity is unmapped, a scheduled shift has no
final evidence, or a planned time-off record is missing from final Activities.

## Realisations hour allocation

`HOURS_BY_AGENT` and `HOURS_BY_LOB` break every published-shift minute into
one exclusive observed category: Available, Inbound, Outbound, BO, other
productive, Management, Support, Training, IT incident, other AUX, Break,
Lunch, Logged Off, LILO-only, unknown, future, PTO, or Away. The effective
status rulebook determines AUX buckets. LILO-only time is never guessed into
a productive subtype. `allocated_hours + reconciliation_delta_hours` must
equal `scheduled_hours`; a nonzero delta is a calculation defect to review.
If a published shift has no timeline segment, its hours remain visible as
PTO/Away only when the attendance register proves that state, otherwise as
unknown. No scheduled shift disappears just because status evidence is absent.
These sheets explain where scheduled hours went, not payroll hours or a new
agent-adherence KPI.

## Staffing

The native grain is 15 minutes and Staff Type. FTE Requirement is the demand
requirement. Gross scheduled FTE is derived from overlapping published shifts
and employee FTE. PTO/Away FTE is the approved time-off overlay. Net scheduled
FTE equals gross less time off. Coverage equals net divided by required when
required is positive. Shortage is positive required less net.

Management LOB, Planning Group, workforce LOB, file prefix, forecast Staff Type,
schedule assignment, and output Staff Type are governed in
`capacity_mapping.csv`.

## PCS

Only inbound legs with the configured survey mode/status/question rules enter
each PCS counter. WFMHub stores eligible legs, valid responses, score sum,
participation responses, blank/invalid exceptions, top-box and low-score counts.

- PCS average = valid score sum / valid score count.
- participation = configured nonblank participation responses / configured
  eligible population.
- team/month results = sums of components followed by the ratio.

Full detail is in [PCS.md](PCS.md).
