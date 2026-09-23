# Design system

WFMHub outputs must feel like one operational suite: compact, calm, precise,
and easy to scan during a live shift.

## Visual language

- Navy: structure, titles, strong text.
- Teal: operational action and active selection.
- Gold: controlled emphasis and section accents.
- Red: verified risk or breach—not decoration.
- Amber: review or incomplete evidence.
- Green: healthy/ready state.
- White and cool gray: dense working canvas.

Exact color tokens and report version are in `src/wfmhub/design.py`. The Excel
reference workbook is `docs/WFMHub Report Design Reference.xlsx`.

## Report composition

1. compact identity band with product, period, refresh time, and evidence state;
2. no more than four to six decision cards;
3. one main operational table or chart;
4. one prioritized exception/action area;
5. detailed evidence tables on separate sheets;
6. definitions and audit lineage at the end.

Cards must show a value, label, and decision meaning. Tables use consistent row
height, date/time formats, filters, frozen panes, and widths. Empty data must
say why; never render a dead unlabeled chart.

## Local workbench

Desktop is the primary target, with tablet fallback. Navigation identifies the
WFM cycle rather than technologies. Global filters remain connected; page-only
controls are visibly scoped. Multi-select filters show their active values and
offer one clear reset.

The browser renders governed data and actions. It does not contain KPI formulas
or editable business mappings. Every exception table exposes evidence and the
next WFM decision.

## Authorship

User-facing files use normal product and business language and credit Anass
ASSRI where authorship is appropriate. Do not add generated-by-tool text,
assistant language, placeholder insights, invented recommendations, or novelty
KPIs.
