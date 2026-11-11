# WFMHub ABS Probe — first collector milestone

This independent portable diagnostic replaces PAD as the first extraction
experiment. It uses official embedded Python and Windows UI Automation through
ctypes/COM. It requires 64-bit Windows and an existing signed-in Edge session.
No extension, browser driver, PowerShell script, or Python installation is used.

1. Extract the whole Windows ZIP to a local writable folder.
2. Open ABS in Edge and navigate manually to TOYOTA / Agendas à Traiter / backlog.
   Keep ABS as the active tab in a visible, unlocked Windows session.
3. Run `PROBE_ABS.cmd` and press a key when ready.
4. Read the new TXT summary under `_system/output/abs-probe/`. Its companion JSON
   contains the bounded accessibility tree and redacted property metadata.

The normal title marker is `ABS@ACM`. If your tab title differs, run
`PROBE_ABS.cmd --title "Your ABS title marker"`. Multiple matching windows stop
the probe; close duplicates or pass `--window-handle 0x...` from the JSON report.

The probe only reads. It does not navigate, click, scroll, focus controls, copy
clipboard contents, invoke task actions, or publish backlog snapshots. It reports
grid candidates, exposed row/column positions, text lengths, and known headers.
Passwords are skipped. Case text, URLs, automation IDs, and arbitrary UI strings
are not persisted. Header matches alone do not establish reliable collection.

Default bounds are 1,200 nodes, depth 20 and 45 seconds. An unresponsive provider
is contained in a worker process; timeout stops that worker, not ABS or Edge.
Reports distinguish no window, ambiguous windows, failure, timeout, limits, and
a completed diagnostic. A completed probe never means full backlog coverage.

Next: inspect the diagnostic from the actual ABS grid. If the required full
values and positions are exposed, implement collection for one TOYOTA group,
including scrolling and source-total reconciliation. Automatic collection is
not implemented by this milestone. The BO Live dashboard/database remain reusable.

```text
PROBE_ABS.cmd
_system/runtime/           Official hash-verified Windows CPython
_system/app/abs_probe.py   Standard-library diagnostic
_system/output/abs-probe/  Created on the workstation; excluded from release ZIPs
```

Keep reports local and outside Git. This ZIP contains code only; it does not
bundle browser profiles, credentials, live reports, or customer/employee data.
