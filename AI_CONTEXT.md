# WFMHub mandatory context router

Context version: `2.4.0`
Applies to: WFMHub `0.36.0` and later
Last reviewed: `2026-09-14`

Before proposing or editing anything, read in this order:

1. `AGENTS.md`
2. `WFM_MASTER.md` — single product/business truth
3. the effective config and active source/build files relevant to the task
4. the tests that prove the behavior

Do not reconstruct the product from old chat, archived reports, TOLEARN files,
prototype leftovers or unreachable builders. Current user instructions and
dispatched code with passing tests outrank prose.

Hard stops:

- Never edit, move or clean the source extracts.
- Never guess Client IDs, queue membership, Staff Type mappings, statuses,
  absence reasons or missing values.
- Service queues and workforce Staff Types are separate domains.
- RSA BE service level is combined; RSA BE FR and VL capacity remain separate.
- Agent Status is primary observed attendance; LILO is fallback.
- Verint Activities is final post-day absence/shrinkage and correction-overlap
  evidence only; it never creates observed presence.
- PTO/Away changes expected work and net capacity but not the underlying
  published schedule assignment.
- No adherence KPI, runtime AI, DuckDB, ODBC, hidden database server, or required
  Excel Data Model.
- PCS is a permanent collaborative Excel tracker and is outside the local console.
- The default presentation is the 15-page localhost Manager Workbench. It reads
  governed SQLite marts directly and binds only to `127.0.0.1`.
- Manager Workbench pages follow Desk, Plan, Operate, Review, Deliver and Govern.
  Report/analysis actions read current marts; only Update rebuilds data.
- Power BI implementation files are removed; migration 016 remains only for
  additive database compatibility. Do not regenerate or package it unless
  the user explicitly starts a future Power BI project; normal updates never
  publish its duplicated CSV feed.
- Missing evidence remains unknown; percentages are never averaged when their
  additive numerator and denominator exist.

Primary authorities:

- product and formulas: `WFM_MASTER.md`
- current report dispatch: `src/wfmhub/report_packs.py`
- service scope: `config/service_profiles.toml`
- capacity scope: `config/capacity_mapping.csv` or shipped default
- KPI arithmetic: `config/metric_catalog.toml`
- attendance/activity classification: `config/wfm_rules.toml`
- local console projections: `src/wfmhub/web_data.py`
- local console server/security: `src/wfmhub/webapp.py`
- local console pixels/interactions: `src/wfmhub/web/`
- release behavior: tests plus `packaging/windows/build_portable.py`

If documentation and active behavior differ, stop and reconcile the master,
code and tests in the same change. Never silently preserve a contradiction.
