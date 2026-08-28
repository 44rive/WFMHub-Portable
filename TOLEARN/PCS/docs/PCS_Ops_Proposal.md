# Post-Call Survey (PCS) — Daily Operations Report
## Proposal

**Prepared by:** WFM
**Data source:** WFMHub Portable — `mart.agent_pcs_day`, built from the Storm Call-by-Call extract
**Refresh cadence:** every 3 hours on working days
**Status:** for validation

---

## 1. Why this report

PCS scores currently sit inside the Call-by-Call extract, which nobody outside WFM opens. The
scores exist, they are already scoped to our roster, and they are already deduplicated — they
are just not reaching the people who can act on them.

This report puts four things in front of Operations, refreshed every 3 hours:

1. **PCS average** by Ops Manager, LOB, Team Leader and agent
2. **Response rate** — are we even collecting enough feedback to judge anyone
3. **Top-box and low-score share** — the shape of the distribution, not just the mean
4. **An exception list** — the handful of rows that need action today

Everything is derived from one governed source. There is no second version of the truth and
no manual re-keying.

---

## 2. What a "PCS-enabled call" is

This is the most important definition in the report, because it is the denominator of the
response rate, and it is the one that will be challenged first.

> **A call leg is PCS-enabled when it is inbound AND the Storm `PostCallSurveyMode` field
> reads `2`.**

That is the entire test. Precisely:

```
inbound  AND  PostCallSurveyMode = "2"   ->  PCS-enabled
```

### What it does not mean

- It does **not** mean a survey was delivered to the customer.
- It does **not** mean the customer stayed on the line.
- It does **not** mean the call was answered or handled — no handle-time condition is applied.

It means the call was **flagged for the survey mode we treat as our PCS programme**. Read it as
*"calls that were in scope to be surveyed"*, not *"calls that were surveyed"*.

### What must be true before a call reaches that test

Four gates, all applied upstream by the hub:

| Gate | Rule |
|---|---|
| **Roster scope** | The agent must be in the active FTE roster — matched by Agent ID, or by a name that maps to exactly one FTE agent. Anything else is discarded before storage. |
| **Valid timestamp** | The call start must parse. |
| **Deduplication** | One row per Call Key. Where the same leg appears in overlapping extracts, the newest active file wins. |
| **Identified agent** | The agent ID must be present. |

The Call Key deliberately excludes the end time and the survey answers, so when a later extract
arrives carrying the customer's answers, it **updates the same leg** instead of creating a
duplicate. Late-arriving survey responses are handled correctly and do not inflate volumes.

### The one caveat we are disclosing up front

**PCS-enabled calls counts call legs, not customer interactions.**

If a call is transferred from Agent A to Agent B, that is two legs. Both are inbound, both carry
the same survey mode, so **both count as PCS-enabled** — but the customer can only return one
survey. On transfer-heavy queues this structurally deflates the reported response rate.

We are choosing to report at leg level because:

- it reconciles exactly, row for row, with the hub's own PCS workbook;
- agent-level accountability is a per-leg concept — each agent owns their own leg;
- introducing a second, interaction-level response rate would put two different numbers into
  circulation for the same question.

**Consequence to communicate:** response rate is a *floor*, not a precise collection rate. Use it
to compare teams and to track direction over time. Do not read it as "X% of our customers
responded".

If Operations wants the true interaction-level rate, the underlying field to group on already
exists in the data and we can quantify the gap as a one-off analysis.

---

## 3. Measure definitions

Every measure below is a ratio of two stored counters. Nothing is an average of an average.

| Measure | Definition | Denominator |
|---|---|---|
| **PCS-enabled calls** | Inbound legs flagged with our survey mode | — |
| **Survey responses** | PCS-enabled legs where at least one scored question came back with a valid score | — |
| **Response rate** | Survey responses / PCS-enabled calls | PCS-enabled calls |
| **PCS average** | Mean of the per-call scores | Survey responses |
| **Q1 / Q2 average** | Mean of that question's valid scores | Responses to that question |
| **Top-box %** | Share of responses scoring **4.0 or higher** | Survey responses |
| **Low-score %** | Share of responses scoring **2.0 or lower** | Survey responses |
| **Comments** | PCS-enabled legs with any free-text answer | — |
| **Handled calls** | Legs with any talk, hold or wrap time greater than zero | — |
| **AHT** | (Talk + hold + wrap) / handled calls | Handled calls |

### How a single call is scored

1. Each configured scored question (currently **Q1 and Q2**) is checked against the valid scale
   **1 to 5**. Anything outside that range, blank, or non-numeric is ignored.
2. The call's score is the **mean of its own valid answers**.
3. That single call score is what enters top-box, low-score, and the PCS average.

### Three definitional points that will be questioned

- **Top box is applied to the call's average, not to a single question.** A call answering Q1=5
  and Q2=3 averages 4.0 and counts as top box.
- **A partial response counts as a full response.** A call answering only Q1=5 scores 5.0 and
  carries the same weight in the PCS average as a call that answered everything.
- **Every response weighs equally in the PCS average**, regardless of how many questions it
  answered.

These are consequences of the standard "average the response, then average the responses"
approach. They are stated here so they are agreed in advance rather than discovered later.

---

## 4. Blank is not zero

**An agent with no valid responses shows blank. Never `0.00`.**

This is enforced end to end — in the hub, in the data model, and in the report. It matters
because a zero would be read as "this agent scored terribly" when the truth is "we have no
feedback on this agent". A zero would also drag their team's average down, and since a team
average is response-weighted, an agent with no responses must contribute nothing at all.

**If a cell is blank, we have no data. It is not a bad score.**

### But a 0% response rate is a real zero

There is one deliberate exception, and it is worth understanding because it looks inconsistent
until you see why.

| Agent had… | PCS average | Response rate |
|---|---|---|
| 40 PCS-enabled calls, 0 responses | **blank** | **0.0%** |
| 0 PCS-enabled calls at all | **blank** | **blank** |

If an agent took 40 surveyable calls and not one customer replied, then **0% is a true and
actionable statement** — we are collecting nothing from this agent. The PCS average stays blank
because there is genuinely nothing to average.

The rule is simply: *a ratio is shown whenever its denominator is real, and blank when the
denominator does not exist.* Blank means "cannot be calculated", never "bad".

The first row of that table is exactly what the **Enabled but silent** block of the exception
list is for.

---

## 4b. Why the arithmetic is done the way it is

Every ratio in this report divides a **total by a total** — total score points over total
responses. It never averages the per-agent averages.

That sounds like a technicality. It is worth 1.26 points on a real team shape:

| Agent | Responses | Their average |
|---|---|---|
| A | 42 | 4.6 |
| B | 28 | 4.4 |
| C | 19 | 4.5 |
| D | 6 | 3.0 |
| E | 3 | 2.0 |
| F | 2 | 1.5 |
| G | 1 | 1.0 |

| Method | Team PCS average |
|---|---|
| **Total points / total responses — what this report does** | **4.26** |
| Mean of the seven agent averages — the intuitive but wrong way | 3.00 |

The three agents holding 6 of the team's 101 responses carry **43% of the weight** in the naive
figure but represent **4% of the actual customer feedback**. A Team Leader reading 3.00 is being
shown a crisis that the customers did not report.

This is why the underlying export stores each ratio's numerator *and* denominator separately, and
why the report is built from those rather than from the ready-made percentage columns. It is also
why the pre-computed ratio columns are deliberately **removed** from the data model before anyone
can drag them into a pivot.

---

## 5. Minimum responses before ranking — proposed

On a 1-to-5 scale, a single low score among four responses moves the mean by roughly 0.75. Agent
league tables built on two or three responses are noise, and coaching conversations built on them
damage trust in the whole report.

**Proposal: agents are ranked only once they have 5 or more responses in the selected period.**

Agents below the threshold are **still shown**, in a clearly separated "insufficient data" block
with their response count visible. They are not hidden and not silently excluded — they are
simply not ranked. The threshold is a single setting and can be tuned once we see real volumes.

---

## 6. What Operations receives

| View | Audience | Content |
|---|---|---|
| **Rollup** | Ops Managers | PCS average, response rate, top-box %, AHT by Ops Manager and LOB. No agent names. |
| **TL scorecard** | Team Leaders | Their agents ranked, response count beside every average, insufficient-data block separated. |
| **Daily trend** | Both | PCS average and response rate by date — direction, not a single day's noise. |
| **Exceptions** | WFM + TLs | Agents with PCS-enabled calls but zero responses; low-score outliers; day-over-day drops. |

Period selector on every view: **Today (partial) / Yesterday / Rolling 7 days / Month to date**.

Per-agent daily response counts are usually too thin to act on alone, which is why rolling 7 is
the recommended default for coaching conversations. Today and Yesterday are for the pulse.

---

## 7. Refresh cadence and freshness

The report refreshes **every 3 hours**. Each run:

1. Re-ingests the FTE roster and the Call-by-Call extract only — not the full estate.
2. Rebuilds the PCS model for the window **1st of last month through today**.
3. Exports one governed file and stamps it.

**Every sheet carries a freshness stamp showing the exact time of the last refresh.** At a
3-hourly cadence this is not decoration — without it, someone quotes a 09:00 number at 16:00.

**The report can only be as fresh as the Storm Call-by-Call extract in the source folder.** If
that extract is dropped once a day, a 3-hourly refresh re-reads the same file three times. The
cadence delivers intraday value only if the extract lands intraday. This needs confirming with
whoever owns the Storm export schedule.

### Today's figures are partial by definition

The current day is always incomplete, and surveys arrive after the call ends, so today's response
rate is always understated and rises through the day. Today is labelled **partial** on every view.
**Do not compare a partial day against a complete one.**

---

## 8. Governance

- **Source extracts are never modified.** The refresh job runs read-only against them.
- Every export is accompanied by a manifest recording the period, the row count, the rule
  version and the **SHA-256 of the rulebook** that produced it.
- All KPI formulas live in the hub's central rulebook as text, not in code, and every stored row
  carries the rule version and hash that produced it. Any number in this report is traceable to
  the exact ruleset that generated it.
- Changing a threshold — the scale, top-box cut-off, scored questions — is a rulebook edit
  followed by a validation step. It is deliberately not something anyone does inside Excel.

---

## 9. What is needed to go live

1. **Agreement on the definitions in sections 2 and 3**, particularly the leg-level denominator
   and the partial-response treatment.
2. **Agreement on the minimum-response threshold** in section 5.
3. **Confirmation of the Storm Call-by-Call extract schedule** — see section 7.
4. **A decision on the response-rate denominator.** The July extract has been analysed and
   `PCSStatus = '1'` is a strict subset of `PostCallSurveyMode = '2'` (9,661 of 10,792 legs).
   **All 1,017 responses carry `PCSStatus = '1'`** — the 1,131 legs without it never returned a
   single response. Moving the denominator to `PCSStatus` raises the reported response rate from
   **9.42% to 10.53%**. See `DATA_FINDINGS.md` §1. **Decide before the first send.**
5. **An FTE roster that covers the agents in the Call-by-Call extract.** The current roster
   matches 1 of its 130 agents against the extract, so the Team Leader and Ops Manager views
   cannot be built yet. See `DATA_FINDINGS.md` §2. **This is the largest gap.**
6. A distribution list and a named owner for the exception list.
