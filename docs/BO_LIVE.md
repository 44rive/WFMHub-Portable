# TOYOTA BO Live

## Current diagnostic: portable C# and installed .NET Framework

`miniapps/abs_csharp_probe/` provides CHECK_LAUNCH.cmd followed by PROBE_ABS.cmd.
It targets .NET Framework 4.8, uses the installed Microsoft UI Automation
assemblies, and bundles no Python/native runtime DLLs. It is unsigned and still
requires permission to run on the workstation. No application policy is changed.

The first command reports `launch_ok` without accessing ABS. The second reads
redacted accessibility metadata in a bounded worker; it never clicks, scrolls,
changes focus, reads the clipboard, changes tasks, or publishes a task capture.
It reports missing/ambiguous windows, provider failures, limits and timeouts.
Native UI Automation tests use a synthetic owned window in Windows CI; only
the actual PC can establish compatibility and full ABS value availability.
The package follows `_system/collector`, `_system/source`, `_system/output`.
See its included README for the current run instructions. The prior Python
probe below is historical and was blocked on the workstation's ctypes import.

## Current collector milestone: PAD removed from the new design

The read-only Windows probe is implemented in `miniapps/bo_live/abs_probe.py`.
Build it with `python packaging/windows/build_abs_probe.py` and run its
`PROBE_ABS.cmd` after manually opening the TOYOTA backlog in Edge. The ZIP uses
official embedded Python and built-in Windows UI Automation through ctypes/COM.
It contains no extension, browser driver, or third-party automation package.

Local TXT/JSON reports show candidate grid properties, row/column positions,
known headers, read failures, and traversal limits. Arbitrary control text is
redacted and passwords are skipped. A separate worker has a timeout. The probe
does not click, scroll, change focus, use the clipboard, or publish a snapshot.
Full backlog coverage remains explicitly unverified, even on a completed scan.

A workstation test reported Application Control blocking the `_ctypes` import
before collection starts. The denied file may be `_ctypes.pyd` or a DLL
dependency; the workstation policy event identifies it. Portable packaging
does not grant execution permission. This collection path needs IT approval
of the runtime components. Supported ABS feeds or normal user-initiated copy/
paste are alternatives to evaluate; copying must be tested and is manual.

Once the actual ABS report proves the necessary data is exposed, the next
milestone is a portable collector for one TOYOTA group, with full traversal and
source-total validation. The existing BO database and dashboard remain reusable.

The rest of this guide describes the previously published v0.1.0 package.

WFMHub BO Live is a standalone local companion for the ABS back-office task
list. The implementation lives in `miniapps/bo_live/` and has no imports from
the main WFMHub runtime. It keeps its own database and local settings; the
existing WFMHub database, reports, launchers, and formulas are unchanged.

## Run it

Use the separate `WFMHub-BO-Live-v0.1.0-win-x64.zip`, extract it to a normal
local folder, and run `DEMO_BO_LIVE.cmd` or `START_BO_LIVE.cmd`. The ZIP includes
an official, hash-verified Windows Python runtime and pinned timezone data.
The main WFMHub Windows ZIP does not include the companion.

From a source checkout:

```text
python miniapps/bo_live/app.py serve --demo
python miniapps/bo_live/app.py serve
```

Demo uses labeled synthetic tasks in a disposable separate data folder.
Live mode starts with no evidence. Set the actual ABS source time zone, exact
TOYOTA task-group names, assignment interpretation, and capture thresholds in
Setup. Initial scope is `(TOYOTA) Agendas à Traiter`; full TOYOTA coverage stays
unconfirmed until configured. Do not guess clipped task-group names.

## Connect PAD

The app distribution includes `PAD_SETUP.md`, `PUBLISH_CAPTURE.cmd`, and a
headers-only `PAD_COLUMNS.csv`. PAD must be recorded on the actual workstation
to navigate ABS, read full fields, and verify every task in the agreed scope.

The publisher transforms a complete PAD-extracted CSV into a validated JSON
capture, atomically places it in the local inbox, and leaves the input intact.
The running app watches the inbox, stores immutable evidence, and updates the
dashboard. Failures have their own capture health and preserve last-good rows.
Manual JSON import and rejection review are available in the UI.

`Scheduled for` is the time to action a task. Metrics describe listed task
rows, priority, assignment, and timing at each group's capture. Repeated case
numbers remain separate because ABS currently exposes no unique task ID.
There are no inferred completions, throughput, SLA, case-age, attendance, or
employee-productivity measures. Capture times and partial scope remain visible.

## State and upgrades

The app's independent `_local/` folder holds settings, `bo-live.sqlite3`, and
inbox/processed/rejected captures. The actual path appears in Setup. Keep it on
local disk and outside sync folders. Stop the app and PAD before copying the
whole folder for backup or an upgrade. Start the next version with `--home`
pointing to the preserved folder. The app runs only on loopback.

The existing WFMHub `UPGRADE.cmd` does not migrate companion state. The separate
companion package contains code and defaults only. Source time zone is fixed
after the first complete capture to preserve how evidence was interpreted.

## Build and validate

```text
python packaging/windows/build_bo_live.py
python -m unittest discover -s tests -p test_bo_live.py -v
```

Source-only packaging is available with `--source-only`. The normal builder
reuses the existing reviewed CPython archive hash and timezone wheel pin. The
automated tests cover fresh-schema initialization, case repetition, capture
validation/idempotence, stale/failed/empty evidence, DST, CSV publishing, inbox
processing, and the loopback API. Actual ABS extraction and Windows CMD/PAD
execution require workstation validation.
