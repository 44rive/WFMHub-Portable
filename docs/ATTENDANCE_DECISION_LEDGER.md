# Attendance decision ledger

## The business flow

WFMHub never edits an extract. Verint schedule boundaries say when an agent was
expected. Agent Status is the primary proof of what happened inside that shift.
LILO fills missing Status coverage and remains a control for first login, last
logout, blank daily rows, and incomplete Status exports.

The Hub detects continuous gaps at their exact timestamps. It does not round a
gap to 15 minutes and does not bridge two gaps when an agent returns between
them. Today is excluded from Attendance Review, so an unfinished shift cannot
become an early leave.

## How to operate it

1. Refresh Attendance sources.
2. Open **Operational > Attendance Review > Build or rebuild**.
3. Open `Reports\Attendance Review.xlsx`.
4. On `REVIEW BOARD`, edit only Decision Category, Decision Status, Reviewed By,
   Comment, and Reviewed Date.
5. Use `Approved` when the interval and category are correct.
6. Use `Dismissed` when the detected interval must count as no loss.
7. Leave it `Open` when it is not decided.
8. Save and close Excel.
9. Choose **Attendance Review > Import completed decisions**.

Each case uses two aligned rows followed by a short blank separator. `SCHEDULE`
is directly above `ACTUAL` and shows scheduled work plus PTO/Away; `ACTUAL`
shows Logged, Break, Lunch, gaps and unknown evidence. Edit blue decision cells
only on the ACTUAL row. Dark red is
the exact gap owned by that row, light red is another counted gap for the same
agent-day, and grey is inside the configured tolerance. The cells are a
15-minute reading aid; Exact Start and Exact End remain authoritative. Hidden
`EVIDENCE` keeps every exact source segment, while hidden `DECISION LEDGER` is
the read-only snapshot already persisted in WFM Hub. They are audit and
troubleshooting sheets, not places to type decisions.

`Meal Aux` is explicit Agent Status presence and is displayed as Lunch. If a
LILO logout timestamp falls inside that interval, Agent Status wins: the meal
is not counted as a gap, and a genuine early-leave interval can begin only
after the Meal Aux interval ends. Missing evidence is never invented as meal.

`BREAK & MEAL` is a separate completed-day control, not adherence. It totals
Agent Status break and meal intervals inside each scheduled shift. The default
allowances are 30 break minutes and 45 meal minutes and remain editable in
`config\wfmhub.toml`. A result is judged only when status coverage reaches the
configured minimum; otherwise it stays `INSUFFICIENT EVIDENCE`.

The import is atomic: one invalid or stale row rejects the whole file. Gap ID
retrieves the authoritative date, agent, start and end from SQLite. Changing a
white evidence cell in Excel cannot alter the stored gap.

## Calculation result

- Approved decisions use the flags in `config\wfm_rules.toml`.
- Dismissed gaps count as neither absence nor shrinkage.
- Open gaps remain unverified and make the affected result incomplete.
- Approved PTO and effective Away intervals come directly from the FTE workbook
  registers. They do not create fake no-show gaps.
- Overlapping intervals are unioned before totals; they are never added twice.
- Absence and shrinkage are parallel measures and must not be added together.

The authoritative outputs are `mart.absence_event` at exact interval grain and
`mart.absence_agent_day` at agent/day grain. Tables whose names begin
`mart.verint_final_absence_` are temporary compatibility projections of those
reviewed marts for existing Excel feeds; Verint Activities no longer supply
their values.

## Audit controls

Every generated event records the rulebook version and SHA-256. Every decision
records its Gap ID, reviewer fields, import filename, and update timestamp in
`core.correction_action`. Rebuilding the model reattaches the stored decision
when the exact Gap ID still exists. A changed physical gap produces a different
ID rather than silently inheriting an old decision.
