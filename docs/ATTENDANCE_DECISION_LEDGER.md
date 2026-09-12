# Attendance residual reconciliation

The filename is retained for compatibility. Attendance Review is no longer a
decision ledger or a write-back workbook.

## Business flow

WFMHub never edits an extract. Verint schedule boundaries say when an agent was
expected. Agent Status is the primary proof of what happened inside that shift;
LILO is fallback/control evidence when Status coverage is insufficient.

The Hub detects every continuous gap at its exact timestamps. It then compares
the gap with final mapped Verint Activities for the same agent and day:

```text
Schedule + Agent Status/LILO -> exact raw gap
exact raw gap - exact final Verint overlap -> residual correction interval
```

- A fully covered gap disappears from Attendance Review.
- A partly covered gap is split and only the remaining fragment is shown.
- A gap with no matching final activity remains in full.
- Activities never prove attendance and never create a gap.
- Today is excluded, so an unfinished shift cannot become Early Leave.

## How to operate it

1. Put the latest StartEndTimes, Agent Status and Verint Activities extracts in
   their normal source folders without editing them.
2. Run the Hub update for the required completed period.
3. Build **Operational > Attendance Review**.
4. Open `Reports\Attendance Review.xlsx`.
5. Use `REVIEW BOARD` as the exact residual backlog. Read Exact Start, Exact End,
   Issue, Suggested Verint Activity and Residual Status.
6. Correct the remaining intervals in Verint.
7. Export the updated Verint Activities file and refresh WFMHub again.
8. Rebuild Attendance Review. Correctly covered intervals disappear.

The workbook is read-only evidence. Nothing is typed back into WFMHub and there
is no import command.

## Visual evidence

Each case uses a SCHEDULE band directly above an ACTUAL band, followed by a
short separator. Scheduled work and PTO/Away are visible above Logged, Break,
Lunch, residual Gap and Unknown evidence. Dark red is the exact residual owned
by the case; lighter red is another residual for the same agent-day. The
15-minute cells are a reading aid; Exact Start and Exact End are authoritative.
Hidden `EVIDENCE` keeps the exact source segments for audit.

Explicit `Meal Aux` is Lunch. If a LILO logout timestamp falls inside that
interval, Agent Status wins: the meal is not a gap and an actual early-leave
fragment can begin only after the meal ends.

`BREAK & MEAL` is a separate completed-day control, not adherence. It totals
Agent Status break and meal intervals inside each scheduled shift. Default
allowances are configured in `config\wfmhub.toml`. A result is judged only when
Status coverage reaches the configured minimum; otherwise it stays
`INSUFFICIENT EVIDENCE`.

## Audit and final absence

Gap ID is derived from the source date, agent, issue and original exact
boundaries. Residual segments retain that lineage plus final-activity overlap
and reconciliation status. Final Absenteeism is separate: it uses mapped Verint
Activities clipped to the planned shift and explicitly flags empty, partial,
unmapped or missing-planned-time-off coding.
