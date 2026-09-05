# Changelog

## 0.16.1 — 2026-09-05

- Fixed a `mart_verint_final_exception.exception_key` unique-constraint failure
  when overlapping daily or full-history Verint Activities extracts repeat the
  same agent activity and exact interval under different filenames. WFMHub now
  keeps one business event, merges its source-file provenance, and applies a
  defensive primary-key deduplication before materializing final exceptions.

## 0.16.0 — 2026-09-05

- Added a novice-facing PCS `TEAM_VIEW` driven by the existing period, LOB,
  Team Leader, and Agent selectors. It shows agent realizations beside the
  matching coaching queue and reads new agents directly from the refreshable
  Excel Tables; the permanent `COACHING` log remains separate.
- Turned Final Absenteeism into a long-lived collaboration workbook with a
  selector-driven team view, refreshable review queue, filtered absence and
  shrinkage component view, exact Activities detail, and a permanent editable
  `ACTIONS` log keyed by Case ID.
- Rebuilt Staffing as a full-period actual-control and future-capacity product.
  It compares forecast required FTE with net schedules after PTO/Away, exposes
  weekly FTE-hours, and surfaces demand even when no roster interval exists.
- Expanded Realisations from the Ford pilot to every active mapped service
  profile in one management workbook. Service scopes and roster LOBs are now
  explicitly paired, while each LOB retains its own configured SLA method and
  target.
- Moved default on-demand analysis outputs into visible `Reports\Analysis`,
  improved command logging so a run always leaves usable evidence, and removed
  the public-workbook Copilot suggestion while retaining the optional static
  prompt file requested for manual use.
- Added regression coverage for dynamic shared views, all-LOB Realisations,
  full-period Staffing, visible analysis output, workbook contracts, and
  external-link safety.

## 0.15.0 — 2026-09-05

- Rebuilt the public report suite around one WFM Hub visual identity and fixed
  business-specific products: Attendance Callout, Staffing Gaps, OEM Flash,
  Realisations, Attendance Review, Final Absenteeism, Bonus Management and PCS
  Performance. Workbook metadata and footers identify Anass ASSRI as author.
- Added fixed-name PCS and Absenteeism CSV feeds for long-lived shared Excel
  files. PCS now has loaded-date periods, cascading LOB / Team Leader / Agent
  selectors, filterable agent results, a permanent coaching log and a separate
  refreshable opportunity queue keyed by Agent ID and Coaching Key.
- Made Agent Status the primary detailed attendance evidence, retained LILO for
  outer boundaries and sparse-data fallback, protected unfinished shifts from
  early-leave logic, and kept exact separate gaps when an agent returns later.
- Accepted parsed Activities Shift Assignment boundaries when a dedicated
  StartEndTimes extract is unavailable, while raising a visible review finding
  and continuing to use Activities as the post-correction final ledger.
- Extended roster scope for operational extracts that contain a real Agent ID
  when the unique matching FTE name has a blank Client ID. Ambiguous names stay
  excluded; Active and effective-dated Leaver rules still apply.
- Simplified Attendance Review to Logged, Break, Lunch and Gap evidence colors,
  with exact unrounded Start/End timestamps kept in `VERINT_INJECTION` for each
  continuous residual interval across every completed date selected.
- Expanded Final Absenteeism with absence and shrinkage components, exact
  activity detail, permanent review actions, uncoded-empty-shift controls and
  finalized-coverage reporting so incomplete cases cannot dilute KPI rates.
- Recreated Bonus Matrix v1.2 as the eight-sheet management workbook, including
  the original policy/configuration structure, formula mechanics, management
  controls, team analysis and WFM Hub charts without changing the source file.
- Added the Flash OEM Realisations pilot with LOB/day facts plus monthly, ISO
  week and quarterly trends for actual, forecast, service level, service
  availability, AHT, staffing, final absence and shrinkage. Adherence remains
  intentionally excluded.
- Added regression coverage for shared-feed schemas, Activities schedule
  fallback, blank-ID roster matching and report contracts. The release passes
  72 automated tests and workbook external-link checks.

## 0.14.0 — 2026-09-04

- Activated the governed `PTO` and `Away` sheets in the standard FTE workbook.
  Approved PTO and effective Away are clipped to exact schedule intervals and
  applied consistently to attendance, correction gaps, timelines and staffing;
  Planned Away affects future capacity only.
- Split interval staffing into gross scheduled, planned-time-off and net
  scheduled FTE. Verint Activities remain the final payroll authority, while
  missing or partial post-day coding for expected time off is now explicit.
- Rebuilt PCS as a simpler TL cockpit: reliable loaded-date period selectors,
  one agent realization list, selector-driven dashboard/trend, governed sample
  threshold, Agent-ID-safe filtering, and a focused coaching action table.
- Kept PCS coaching outside SQLite while carrying the five editable Excel
  fields forward between report generations using the stable Coaching Key.

## 0.13.1 — 2026-09-04

- Replaced the editable correction-decision sheet with `VERINT_INJECTION`: one
  row per exact continuous residual interval, with second-level start/end
  timestamps, injection suggestion, evidence, schedule context and stable
  lineage keys.
- Stopped bridging separate Agent Status gaps across a real return to service.
  The tolerance now suppresses only tiny individual gaps; it never manufactures
  one longer interval that would be unsafe to inject in Verint.
- Removed correction-decision import from the user-facing CLI. The next Verint
  Activities export is the only reconciliation signal and automatically clears
  covered residual intervals.
- Expanded the blank FTE template with governed `PTO` and `Away` Excel Tables,
  leading-zero-safe Agent IDs, date/time formats, dropdowns, required-field
  highlighting and plain-language instructions. The registers were staged for
  the calculation overlay delivered in v0.14.0.

## 0.13.0 — 2026-09-04

- Replaced the PCS Power Query/Data Model master with one self-contained
  `PCS Performance.xlsx` generated by Python/SQLite. It includes formula-driven
  period, LOB, team-leader and agent selectors, visible Excel Tables, daily and
  monthly comparisons, and an editable coaching plan that needs no Hub sync.
- Made Agent Status the primary attendance boundary when coverage is sufficient;
  LILO remains the sparse-data fallback. A logout followed by a later return is
  an internal gap, while a final logout becomes early leave only after the
  completed shift. Full-shift Logged Off evidence is a no-show.
- Changed `Attendance Review.xlsx` to honor every completed day in the selected
  range, including Current Week, while always excluding today's unfinished tail.
- Added final-ledger completeness states for uncoded empty shifts, observed but
  uncorrected gaps, partial corrections, and provisional days. Exception rows
  cannot dilute final absence rates as silent zeroes.
- Simplified the portable surface to `Reports`, `Feed`, `config`, and `_system`.
  Clean exports now go to the visible `Feed` folder; runtime, SQLite, logs,
  backups, technical output, code and documentation stay under `_system`.
- Removed the PCS setup workbook, Power Query scripts, Hub-side coaching import,
  and workbook-refresh workflow. In-place upgrades preserve old state and move
  superseded PCS/attendance files into a timestamped legacy archive.

## 0.12.0 — 2026-09-04

- Simplified daily work to one flat `Reports` folder with fixed workbook names;
  each replaceable report archives its previous copy automatically.
- Made `Reports/PCS Team.xlsx` the only persistent team workbook. Its
  PivotTables, slicers, and coaching log survive while WFMHub refreshes a small
  four-table star-schema feed under `_system`.
- Rebuilt the Ford OEM pilot as `OEM Flash.xlsx` with OEM/Ford/Toyota service,
  actual-versus-forecast, hourly queue groups, staffing evidence, freshness,
  and an explicit `NOT_CONFIGURED` state for unsourced back-office measures.
- Reorganized the menu around Update, Today, Month, PCS Team, Analyse, and
  Settings; report generation is independent from source ingestion.
- Enforced effective-dated FTE scope everywhere: Active agents are included;
  Leavers are included through their leave date; undated Leavers and all other
  statuses are excluded from schedules, LILO, Agent Status, and Call by Call.
- Updated the portable layout, beginner documentation, PCS Power Query scripts,
  model relationships, governed measures, and packaging checks for the new
  structure.

## 0.11.0 — 2026-09-03

- Rebuilt PCS follow-up at call-response grain: every valid inbound Q1 `<=3`
  is one traceable coaching opportunity instead of an agent-level low-score or
  low-sample warning.
- Restored the original Actions Rate business meaning without its broken
  external workbook: completed coaching divided by low-score opportunities.
- Added persistent PCS coaching decisions in SQLite plus safe workbook import;
  `OK`, `Yes`, `Y`, `Done`, and `Completed` normalize to completed.
- Added stable typed PCS template feeds under
  `output/template_feeds/pcs/current` and immutable refresh archives with
  schema, row counts, and hashes in `manifest.json`.
- Added one-file-per-query Power Query scripts, numeric/date typing, protected
  Excel-master setup, exact Power Pivot measures, and a click-by-click beginner
  guide. No data is loaded to a worksheet and no SQLite ODBC driver is needed.
- Shipped a reviewed, data-free `templates/reports/pcs.xlsx` starter so the
  release no longer opens to an empty template folder; packaging rejects
  external links, connections, and Pivot caches in this public workbook.
- Upgraded the PCS dashboard with coaching-completed and Actions Rate cards,
  and added daily/current-MTD/prior-month coaching comparisons.
- Made the PCS mart always retain the previous full month through the selected
  end date, so choosing Today or Current Week still produces real MTD and
  prior-month comparisons instead of silently showing only the selected days.

## 0.10.1 — 2026-09-03

- Made Attendance Callouts honor the complete selected date range; choosing
  Current Week now includes every available action row and daily trend in that
  week instead of collapsing the product to the end date.
- Added explicit Power Query date, datetime and numeric types so Excel's Data
  Model can sum PCS counters instead of treating CSV values as text.
- Added a model-only PCS agent/day table for date slicers and documented the
  static-report, compact-CSV, Data Model, PivotTable and slicer lifecycle.

## 0.10.0 — 2026-09-03

- Replaced the combined/shared-report workflow with seven focused decision
  products—PCS, Bonus, Service, Staffing, Attendance Callouts, Attendance
  Corrections, and Final Absence—using one management design contract.
- Added daily/MTD/prior-month PCS comparisons, preserved exact original Q1 and
  participation logic, and made the three-hour requirement a sending cadence
  rather than a different KPI.
- Added read-only, content-hashed Bonus Matrix v1.2 ingestion with governed
  agent/KPI marts, source-result reconciliation, scenario/released payout
  separation, and atomic idempotent imports.
- Added effective-dated service profiles and a Ford OEM France product matching
  the original Flash's gross SL, service-availability, forecast-attainment, AHT,
  and Ford/Toyota/Chery presentation logic.
- Added user-selected on-demand analysis for PCS, service, forecast, staffing,
  attendance, final absence, and bonus with explicit comparison modes and
  evidence rows.
- Added the protected Excel-master route: stable compact CSV model packages,
  direct Power Query guidance, connection-only Data Model loading, and a
  `template-init` command that never overwrites a master during daily refresh.
- Split attendance work into a live contact queue, staffing coverage,
  completed-day corrections, and final corrected Verint absence; current-day
  unfinished shifts remain protected from early-leave classification.
## 0.9.0 — 2026-09-01

- Added an effective-dated, scope-aware metric catalog as the single source for
  KPI formulas, targets, samples, units, priorities and aggregation behavior.
- Added a semantic value mart retaining numerator, denominator, method and
  catalog/rule provenance; higher-grain reports use ratio-of-sums.
- Added deterministic Python findings for target breaches, material changes,
  low samples and source-health problems, with evidence filters and hashes.
- Split domain rules, metric methods, finding thresholds, report contracts and
  queue mapping into separately validated editable files.
- Migrated the four operational workbooks to governed datasets and added
  `FINDINGS`, `METHODS`, `DOMAIN_RULES` and `PROVENANCE` sheets.
- Added rule commands to validate, exercise, explain and compare methods and to
  generate one management-readable governance workbook.
- Removed the external AI snapshot/runtime path. Added only a static manual
  Copilot prompt for optional explanation of a user-selected workbook.

## 0.8.0 — 2026-08-31

- Added a management-ready Bonus proposal rebuilt from the audited source with
  centralized rules, explicit policy ownership, input/completeness controls,
  formula documentation, and a technical payroll-release gate.
- Reworked the landing sheet around the uploaded cockpit's executive hierarchy:
  large payout cards, LOB decomposition, input-coverage comparison, payout
  distribution, and a compact governance footer.
- Added a lightweight three-hour PCS template based exclusively on the original
  `PCS Report.xlsx` O/P/Q/R definitions: discrete Q1 average, `<=3`, `>3`, and
  nonblank-Q1 participation over `PCSStatus=1`.
- Retained full call timestamps for true interval reporting and added separate
  valid-response rate, LOB ratio-of-sums, low-sample warnings, and data-quality
  controls.
- Removed the original PCS external Actions Rate link and avoided copying its
  worldwide 39,982-row cache or million-row roster table into the management
  workbook.
- Added the persistent, public-repo-safe `shared_reports` workflow and menu/CLI
  action; generated management workbooks remain local and source files are
  never modified.

## 0.7.0 — 2026-08-29

- Shipped four governed business workbooks: Daily Operations, Yesterday
  Corrections, exact Agent PCS, and Activities-only Final Absenteeism.
- Added an incomplete-data banner and `DATA_MISSING`/`DATA_PARTIAL` staffing
  states; absent attendance evidence now produces blank gap/variance values
  instead of a false staffing shortage.
- Added the latest evidence-complete-day selector, editable residual-gap sheet,
  exact full-shift timeline, and correction calculation contract.
- Added exact reference PCS day/team/month summaries, diagnostic response
  detail, participation logic, and low-sample visibility.
- Added a separately branded final ledger that never uses observed LILO or
  Agent Status for its final absence totals and caps daily classified minutes
  at planned net minutes.
- Included formula, source-health, data-quality, and schedule-role sheets in
  the relevant workbooks; generated files have no external Excel links.
- Replaced the dashboard banner with a compact block-letter WFM HUB wordmark.

## 0.6.1 — 2026-08-29

- Bundled the IANA timezone database required by Windows embedded Python, so
  the default `Europe/Berlin` evaluation clock loads correctly.
- Added a safe fallback to the Windows local clock if a portable installation
  is incomplete, preventing timezone data from aborting attendance refresh.

## 0.6.0 — 2026-08-29

- Parked the next report redesign behind a governed clean-data validation
  layer, with explicit exports for attendance calls, staffing, APDE service,
  PCS, residual gaps, shift timelines and final absenteeism.
- Strictly separated Verint StartEndTimes operational schedules from the
  Activities-only corrected ledger using a header-derived source variant.
- Added one timezone-local evaluation cutoff per refresh, provisional
  current-day attendance states, and protection against unfinished early leave.
- Added agent-second 15-minute staffing, exact shift evidence segments and
  overlap-safe residual correction intervals.
- Added an Activities-only final absence event/day ledger with unioned overlaps
  and the configured 8.75-hour planned-day cap.
- Reproduced the reference PCS workbook exactly: inbound valid discrete Q1,
  `<=3`/`>3` counts and raw-Q1 participation over inbound `PCSStatus=1`.
- Added exact PCS and clean-data documentation plus real-extract acceptance
  coverage for 2,349 operational agent-days and 56,274 APDE intervals.

## 0.5.1 — 2026-08-28

- Prevented repeated source-load findings from generating the same deterministic
  quality-issue key and aborting a refresh with a SQLite unique-constraint error.
- Preserved the highest severity when identical quality findings are combined;
  extract files and existing databases require no changes.

## 0.5.0 — 2026-08-28

- Changed the attendance/absence source of truth to StartEndTimes schedule
  boundaries plus observed LILO and Agent Status evidence.
- Made Verint Activities a post-day reconciliation ledger only; every observed
  gap is labelled `CORRECTED`, `PARTIAL`, or `NOT_CORRECTED`, and Verint-only
  intervals are surfaced separately.
- Enabled Agent Status by default for gap evidence while keeping adherence,
  conformance, and RTA disabled; added bounded streaming and date indexing for
  large multi-day status files.
- Added the editable, hashed `config/queue_mapping.csv`, detailed/comparison
  scopes, and hourly mapped actual-versus-forecast output.
- Added support for the supplied volume-only Verint forecasts and bracketed
  APDE CSV export while preserving missing measures as NULL.
- Corrected the supplied Ford NL queue mappings so they roll up to Ford NL.

## 0.4.0 — 2026-08-28

- Added `config/wfm_rules.toml`, a safe, editable, versioned source of truth for
  activity classification, payroll absence, shrinkage, SL, service
  availability, AHT, forecast deviation, and scope rules.
- Added rule validation, exact SHA-256 audit identity, and generated KPI,
  activity, and queue-scope catalog sheets.
- Added automatic parsing of Verint's wide StartEndTimes extract alongside the
  existing Activities format; Data Source IDs remains the operational Agent ID.
- Added schedule-clipped, overlap-safe absence events and agent-day absence,
  vacation, unpaid, shrinkage, spell, and Bradford marts.
- Added APDE actual ingestion and a service mart retaining gross/adjusted SL,
  weighted AHT, and service availability (`answered / offered`).
- Added Attendance & Absence and Executive Scorecard workbooks with bounded
  PivotTable-ready Excel Tables. WFMHub does not create the PivotTables.
- Added beginner Rulebook and PivotTable guides.
- Removed adherence metrics and RTA from generated Operations reports. Agent
  Status ingestion and legacy marts are disabled by default but preserved for
  compatibility.
- Redesigned the CMD logo and added active rule version/hash to the fixed status
  panel.

## 0.3.2 — 2026-08-27

- Added a compact WFMHUB ASCII logo and `made by Anass ASSRI` credit to the
  daily dashboard.
- Added a fixed-width status panel showing hub readiness, version, database
  size, roster-agent count, last refresh, selected period, source health and
  current data-quality counts.
- Added the latest loaded business date and its source family to the dashboard,
  so a future Forecast horizon is clearly identified instead of looking like
  the latest operational actual date.
- Grouped the unchanged menu options into Daily Work, Control & Review and Hub
  Tools, with a clean screen redraw after each action.
- Added a native Windows CMD title and optional cyan-on-black theme. Set the
  standard `NO_COLOR` environment variable to keep the console's existing
  colors.

## 0.3.1 — 2026-08-27

- Added a Windows CMD-safe progress bar to setup, refresh, model building,
  report generation, clean exports, Custom Lab jobs, correction imports and
  database backups.
- Added named progress phases for schedules, activities, LILO, Agent Status,
  attendance, conformance, correction gaps, RTA, intraday, PCS, source health
  and data quality.
- Added live scanned-row counts for streaming LILO and Call-by-Call ingestion
  and live written-row counts for CSV/XLSX clean-data exports.
- Kept redirected logs clean by showing progress only in an interactive
  terminal; `WFMHUB_PROGRESS=1` forces it and `WFMHUB_PROGRESS=0` disables it.

## 0.3.0 — 2026-08-27

- Added streaming, incremental Call-by-Call ingestion with FTE scope, typed
  durations/survey fields, stable call-leg keys and cross-file deduplication.
- Added daily and period Agent PCS measures: calls, AHT components, survey
  response rate, Q1/Q2 averages, response-weighted PCS average, top box, low
  scores and comments.
- Split generated workbooks into Operations, Intraday and Agent PCS report
  packs with independent output folders and source-quality views.
- Added selectable date presets and source/report choices to the daily menu.
- Added clean CSV/XLSX exports for calls, surveys, PCS, attendance, gaps,
  schedules, events, LILO, Agent Status, actuals, forecast and source health.
- Added Custom Lab templates for trusted portable-Python jobs and one-statement
  read-only SQL jobs, plus Python-in-Excel recipes in the PCS workbook.
- Generalized multi-day handling across source families. Row dates are
  authoritative for schedules, LILO, Agent Status, calls, forecast and queue
  actuals; filenames are only fallback hints for daily blank LILO rows.
- Multi-day LILO rows with blank boundaries must contain a row-level Date field;
  otherwise they are rejected and surfaced in source health rather than dated
  by guesswork.

## 0.2.1 — 2026-08-27

- Added a standard blank `templates\FTE Count.xlsx` workbook with instructions,
  canonical `Agent` headers, dropdowns, validation, and required-field cues.
- Replaced the literal `Agent`-tab dependency with safe header-driven roster
  discovery across the first 100 rows of every worksheet.
- Added confidence tiers so Support/lookup sheets cannot silently become the
  authoritative roster; ambiguous sheets or duplicate ID/name aliases fail with
  actionable diagnostics.
- Made organisation columns optional while retaining Client/Agent ID and Name
  as the identity contract.
- Added report-pack configuration and routing. The current workbook is the
  Operations pack under `output\operations`; Agent PCS/call-by-call quality is
  reserved as a separate future `quality_pcs` pack.
- Added a reproducible template generator that can standardize an existing FTE
  workbook without modifying the source.

## 0.2.0 — 2026-08-27

- Replaced the blocked DuckDB extension with Python's standard SQLite runtime.
- Added WAL, full synchronous durability, application/storage markers,
  integrity checks, read-only connections, process locking, and online backups.
- Added atomic mart rebuilds and correction-action imports.
- Added an FTE-authoritative agent scope gate using exact Agent ID or unique
  normalized name, while preserving populated operational source IDs.
- Added roster fingerprints so unchanged agent extracts are re-evaluated when
  FTE changes; exposed kept and outside-roster counts in source health.
- Prevented empty/out-of-scope Agent Status from creating synthetic RTA rows.
- Fixed failed-file retry, same-path A→B→A reactivation, and custom report date
  filtering.
- Added a pre-setup system doctor and a deterministic corporate-runtime
  manifest.
- Removed DuckDB and app-local MSVC packages from the release. Only official
  CPython native files and hash-pinned pure-Python Excel libraries are shipped.
- Preserved v0.1 databases/config through side-by-side upgrade guidance; no
  automatic DuckDB conversion is attempted.

## 0.1.1 — 2026-08-27

- Windows launchers now require and invoke the bundled `runtime\python.exe`.
- Removed the fallback to a separately installed `py` command.
- Added a clear recovery message for incomplete extraction, the GitHub source
  archive, or a launcher copied away from the portable folder.
- Normalized the application home path before passing it to Python.

## 0.1.0 — 2026-08-26

- First portable Windows x64 release.
- Bundled CPython 3.13, DuckDB 1.4.5 LTS and app-local Microsoft C++ runtime.
- Incremental, fingerprinted ingestion for FTE, Verint schedule/events, LILO,
  Agent Status, Forecast and Storm APBE/APFR actuals.
- Attendance, conformance, correction, RTA, forecast and intraday marts.
- Overnight-aware LILO stitching and strict no-show/source-coverage gates.
- Persistent Correction ID decisions imported from blue Excel report columns.
- Curated Excel output with no raw extract sheets, Power Query connections or
  embedded Data Model.
- Source health, schema failures and data-quality output.
- Windows launchers, beginner guide, versioned migrations, backups and portable
  release builder.
