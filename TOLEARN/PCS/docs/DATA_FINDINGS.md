# Findings from the real extract

Source: `WFMHub-Portable/TOLEARN/PCS Report.xlsx`, sheet `RDATA`
39,982 call legs, 12–14 July 2026. Produced while building `PCS_Daily_Report.xlsx`.

Five findings. **Two are blockers**, one is a live risk to the hub, two are quality issues.

---

## 1. ⛔ BLOCKER — `PCSStatus` answers the open question, and it changes the number

The proposal flagged one unresolved technical question: is `PostCallSurveyMode` the right gate,
or is it `PCSStatus`? The real data settles the relationship.

| Population | Legs |
|---|---|
| `PostCallSurveyMode = '2'` | **10,792** ← current denominator |
| `PCSStatus = '1'` | **9,661** |
| Both | 9,661 |
| `Mode = '2'` but `PCSStatus` blank | **1,131** |

`PCSStatus = '1'` is a **strict subset** of `Mode = '2'`. And decisively:

> **All 1,017 survey responses have `PCSStatus = '1'`. Not one response ever came from a leg
> where `PCSStatus` was blank.**

So the 1,131 legs flagged for the survey but with no `PCSStatus` **never had any chance of
returning a response**. They are dead weight in the denominator.

| Denominator | Response rate |
|---|---|
| `Mode = '2'` (today) | **9.42%** |
| `PCSStatus = '1'` | **10.53%** |

**The current definition understates the response rate by 11.7% relative.**

**Recommendation:** ask whoever owns the Storm configuration what `PCSStatus` records. If it
means "survey actually presented", the denominator should move to `PCSStatus = '1'` and
`PostCallSurveyMode = '2'` becomes the programme-scope filter, not the collection base.

This is a one-line rulebook change in the hub, but it must be agreed **before** the first send —
changing a headline KPI after publication costs credibility.

---

## 2. ⛔ BLOCKER — no Team Leader or Ops Manager hierarchy for this population

The `AGENT LIST` sheet holds 154 agents (RSA FR, Morocco: TL *Farouk BOUAZZAOUI*, Ops *Anas
OUAKKAD*). `RDATA` holds 1,378 distinct agents across the international estate.

| Match test | Result |
|---|---|
| Roster IDs appearing in the call data | **1 of 130** |
| Roster names appearing in the call data | 3 of 154 |
| `RDATA` rows with `Team leader` filled | **8 of 39,982** |
| `RDATA` rows with `LOB` filled | **8 of 39,982** |

Running the hub's real FTE scope against this extract admits **211 of 39,982 legs (0.5%)** and
yields 8 agent-days with 5 responses. Correct behaviour, useless report.

**Consequence:** the TL scorecard and the Ops Manager rollup **cannot be built** from this
extract. The delivered workbook substitutes the dimensions that do exist.

**What is needed:** an FTE roster covering the agents in the Call-by-Call extract, or a mapping
from the Storm `Agent ID` to the roster `Client ID`. This is the single largest gap between the
report as proposed and the report as deliverable.

---

## 3. ⚠️ RISK TO THE HUB — durations are Excel day fractions, and would import as zero

In `RDATA`, `Talk Time` is stored as a **fraction of a day**: `0.0010416` is 90 seconds.

WFMHub's `duration_seconds()` (`utils.py:119`) handles `timedelta`, `time`, and `"HH:MM:SS"`
text, but for a bare number it does:

```python
if isinstance(value, (int, float)):
    return round(float(value))     # round(0.0010416) -> 0
```

**If the hub ever ingests this file's numbers directly, every talk, hold and wrap value becomes
0**, so `handled_calls` = 0, `handle_seconds` = 0 and AHT is undefined for every agent — silently,
with no error and no data-quality flag.

This is **not currently a live bug**: `parse_calls()` reads CSV, where Storm writes durations as
text. It becomes live the moment anyone points the hub at an XLSX-derived extract, or if Storm
changes its CSV to emit serial numbers.

The first build of this report hit exactly that: `handled_calls = 0` across the board.
`prepare_from_pcs_report.py` now converts day fractions before the hub's parser sees them —
verified AHT of 388.0 seconds, reconciled to the raw extract.

**Recommendation:** treat a numeric duration below 1 as a day fraction in `duration_seconds()`.
A real call is never 0.4 seconds, and it is never a whole day, so the rule is unambiguous. Worth
raising regardless of this report.

---

## 4. ⚠️ QUALITY — 23 agent names are corrupted in the source

Names such as `Anna KwieciÅ„ska-Zibuschka` and `Demir GÃ¼zel` are stored damaged **in the source
workbook**. UTF-8 bytes were decoded as Windows-1252 somewhere upstream in the Storm export.

The damage is exactly reversible, so `prepare_from_pcs_report.py` repairs it (cp1252 first, then
latin-1, applying the fix only when the round trip is clean) and **reports the count rather than
hiding it**. 23 names repaired, 0 remaining. Run with `--keep-mojibake` to see the raw damage.

Polish names only recover through cp1252, not latin-1 — `‚` is cp1252 `0x82` and undefined in
latin-1. A latin-1-only repair silently misses them.

**Recommendation:** fix the export encoding at source. Repairing downstream works, but these
names go in front of Team Leaders and belong to real people.

---

## 5. ℹ️ NOTED — the 1–5 range check is doing real work

`Question 1` contains far more than scores:

| Value | Count |
|---|---|
| `5` | 689 |
| `*` | 343 |
| `4` | 105 |
| `1` | 74 |
| `Invalid_Response` | 53 |
| `3` | 44 |
| `2` | 25 |
| `No_Response` | 16 |
| **`555`** | **1** |

1,350 answers, **937 valid scores**. The 413 discarded are `*`, `Invalid_Response`, `No_Response`
and one `555`.

That `555` matters. Without the `BETWEEN 1 AND 5` guard it would enter an average built on ~937
values and drag a team mean up by roughly half a point on its own. **The range check is not
defensive decoration — it is load-bearing.**

`Question 3` is confirmed as free text (`NO AUDIO` ×261, `todo correcto`, `todo bien`), matching
its rulebook configuration as a comment question rather than a scored one.

---

## What the delivered workbook actually reports

Because of finding 2, the workbook uses the dimensions that are genuinely populated —
**100% coverage across all 1,017 responses**:

| Dimension | Source | Distinct |
|---|---|---|
| **Site** | `AP<country>_<city>_` prefix of `Service` | 14 |
| **LOB** | `LineOfBusiness` | 217 |
| **Language** | `Language` | 18 |
| **Agent** | `Agent` / `Agent ID` | 1,378 |

Their `Q maping.txt` (Queue → Designation → BE/NL/FR grouping) covers 60 queues and matches only
**0.7%** of these responses — it is scoped to the Benelux operation, not this international
extract. Not used.

### Headline numbers, reconciled to the raw extract

| Measure | Value |
|---|---|
| Handled calls | 36,658 |
| PCS-enabled calls | 10,792 |
| Survey responses | 1,017 |
| Response rate | 9.42% |
| **PCS average** | **4.42** |
| Top box % | 85.1% |
| Low score % | 9.4% |
| AHT | 388 seconds |

### The insight the report surfaces immediately

**`APES/MAD — RSA_Insurance-Autos` is the outlier worth a conversation:** 1,118 PCS-enabled
calls, 172 responses, and a PCS average of **3.93** against a group average of 4.42 — the largest
volume of any group *and* the weakest score. By contrast `APUK/CRO — Travel_NWG` runs **4.78** on
164 responses.

That is the report doing its job on day one.
