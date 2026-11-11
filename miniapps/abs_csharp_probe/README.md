# WFMHub C# ABS diagnostic

This is a portable .NET Framework 4.8 application. It uses the Framework and
Windows UI Automation assemblies already installed on the PC. It includes no
Python, ctypes, PAD, browser extension, browser driver, or private runtime DLLs.
The executable is unsigned. Application Control can still block it or a component;
successful operation on a Windows build runner does not establish permission on
your workstation. Signing alone would not establish that permission either.

1. Extract the complete ZIP to a normal local writable folder.
2. Run `CHECK_LAUNCH.cmd` first. A `launch_ok` report proves that this executable
   and report writer started; it says nothing yet about ABS accessibility.
3. If that works, open ABS in Edge at TOYOTA / Agendas à Traiter / backlog.
   Leave that tab active in one visible, unlocked Edge window.
4. Run `PROBE_ABS.cmd`. TXT/JSON reports are under
   `_system/output/abs-csharp-probe/`. Read the status and summary.

If Windows blocks execution or the required .NET Framework is missing, the
launcher cannot create a successful report. Keep the exact error for IT to
identify the permitted deployment. Do not change application policy, rename
binaries, or alter browser settings to work around a block.

The launch check and the accessibility check are separate tests. The second
reports exposed headers, grid candidates, pattern availability, and row/column
properties. It does not establish full table coverage or collect backlog tasks.
Arbitrary names, values, URLs, IDs and class strings are redacted. Known header
labels and string lengths are retained. Password controls are skipped.

The reader has no methods to click, invoke, set values, scroll, change focus,
copy the clipboard, or close source windows. A separate worker bounds provider
hangs (45 seconds by default); timeout terminates only that child process.
No reports or task captures are sent anywhere and no database is opened.

```text
CHECK_LAUNCH.cmd
PROBE_ABS.cmd
_system/collector/ABSProbe.exe
_system/collector/ABSProbe.exe.config
_system/source/ABSProbe.cs
_system/output/abs-csharp-probe/     Created on the workstation only
```

To change the title marker: `PROBE_ABS.cmd --title "Your ABS title marker"`.
Other options are `--max-nodes` (1-5000), `--max-depth` (1-40), and `--timeout`
(5-120 seconds). Multiple matching Edge windows are reported as ambiguous.

Next milestone: inspect the actual ABS result, then implement one-group data
collection with complete traversal and source-total reconciliation if the
required values are available. Existing BO Live captures/dashboard remain reusable.

Build on Windows with `packaging/windows/build_abs_csharp_probe.ps1`. Build tools
are used in CI, not installed by this ZIP. The source has no package dependencies.
