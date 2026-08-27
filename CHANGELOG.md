# Changelog

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
