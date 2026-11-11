# Connect the TOYOTA ABS grid with Power Automate Desktop

The mini app, CSV publisher, inbox watcher, and dashboard are implemented.
These instructions finish the workstation-specific ABS extraction. Working
selectors cannot be inferred from screenshots: record them in the authorized
Edge session using PAD's element picker. Start with one TOYOTA group.

## 1. Configure and test the app

Open `START_BO_LIVE.cmd`, then Setup. Confirm the actual ABS time zone and
exact task-group names. Confirm assignment semantics only after checking ABS.
Use one source time zone consistently. The app initially monitors only
`(TOYOTA) Agendas à Traiter`. Keep its terminal open.

Define PAD variables for `AppFolder`, `ExactGroup`, `ExtractedCsvPath`, and
`VerifiedTotal`. Keep the working CSV on local disk outside Git. In a
complete standalone ZIP, `AppFolder` is the folder containing the CMD files.

## 2. Record the navigation

1. Attach to the authorized ABS Edge tab. Use the normal source sign-in.
2. Open **Actions → Task Groups ...**.
3. Select the full exact TOYOTA group from available groups and confirm **OK**.
   The recently-used list may contain other products; do not rely on position.
4. Confirm the resulting work-list group. Open **Work List** if the actual
   session requires it. Confirm **Task list → List**, filters, and sort state.
5. Record the **Refresh** read path. Collection navigates, refreshes, and reads;
   assigning, allocating, and completing tasks remain operator actions.

Use waits for actual UI states. Do not assume fixed screen coordinates or
select a row merely because its text contains TOYOTA.

## 3. Prove complete extraction

Try PAD browser table extraction first. If the grid does not expose structured
web elements, inspect desktop UI automation/accessibility. Screenshots do not
establish whether it is HTML, canvas, or another renderer.

Extract full values for these columns, including horizontal scrolling or a
read-only Details view as necessary:

```text
Case number,Task type,Task priority,Language,Scheduled for,Assigned to
```

The included `PAD_COLUMNS.csv` contains these headers only. Rename the
extracted DataTable columns to these names before writing CSV. The publisher
also accepts their underscore/camel-case variants. Extra columns are ignored.
Missing required columns or malformed rows fail validation.
Keep case numbers as text throughout extraction and CSV writing; numeric
conversion can remove leading zeros that the publisher cannot reconstruct.

Read and verify the source total from ABS. Confirm the counter's meaning and
whether filtering, caps, or virtual rows affect it. Traverse every task, not
only the visible viewport. A complete structured table can be exported directly.
Overlapping viewport reads need verified absolute row positions before building
the complete DataTable; do not deduplicate using case number or visible values.
Two identical tasks on one case must both remain.

Confirm that a blank assignment was actually read and a blank scheduled time
really is unscheduled. Clipped dates such as `...` must be replaced by the full
source value. Verify the first, middle, and final rows and the source total
after collection. A changing or unreadable scope/total ends as a failure.

## 4. Timestamp, write CSV, and publish

Use **Run DOS command** in PAD to obtain a UTC timestamp just before refresh
and extraction:

```bat
call "%AppFolder%\RUN_PYTHON.cmd" clock
```

Trim the command's standard output and save it as `CaptureStartedIso`. Run the
same clock command immediately after the final verified extraction and save
`CaptureFinishedIso`. The command produces an ISO timestamp with UTC offset;
do not append `Z` to local wall-clock values.

Write the complete extracted DataTable to `ExtractedCsvPath` as comma-separated
UTF-8 CSV with column names, using normal CSV quoting. This is a collector
output; ABS itself still has no export. Do not set the source total from CSV
length: `VerifiedTotal` must come from the independently verified ABS total.

Then run this DOS command, substituting the actual PAD variables:

```bat
call "%AppFolder%\PUBLISH_CAPTURE.cmd" --csv "%ExtractedCsvPath%" --group "%ExactGroup%" --source-total %VerifiedTotal% --started-at "%CaptureStartedIso%" --finished-at "%CaptureFinishedIso%" --scope-verified
```

If the app uses a custom `--home`, append the same `--home "C:\Your\BOData"`
to this command and the failure command below. Check its exit code: `0` means
validated JSON was atomically published to the inbox. A nonzero code means the
capture was not published; inspect stderr and record the collection failure.
The publisher does not edit the input CSV and never writes the main WFMHub DB.

The app watches the inbox every three seconds and refreshes its displayed state
every ten seconds. Confirm the imported group time and totals in Capture history.
Publishing the CSV alone does not update the screen unless the app is running.

## 5. Publish failure health

On expired sessions, unreadable fields, incomplete traversal, or mismatched
totals, use a short diagnostic without case/customer content:

```bat
call "%AppFolder%\RUN_PYTHON.cmd" record-failure --group "%ExactGroup%" --message "ABS table extraction incomplete"
```

Check that command's exit code too. Good rows remain visible and the group's
health becomes failed after the event is imported. Wrong group names are
rejected, so keep `ExactGroup` synchronized with app Setup.

## 6. Repeat for configured TOYOTA groups

Record full EXTRANET and WEBAPP group names from ABS; the supplied screenshots
clip them. Add all intended groups to app Setup and the PAD loop. Mark complete
scope only when those exact names cover the intended TOYOTA population.

Process one group at a time, publish each result, and measure the full loop.
Target one full loop every five minutes if timing allows it. Wait only for the
remaining interval; prevent overlapping runs and use bounded retry/backoff.
Each group has its own observation time. The overview clearly labels partial
coverage and a combined capture window rather than an instantaneous total.

Validate the chosen attended/unattended PAD licensing and Windows session
requirements on the workstation. Begin with a supervised run: one group, then
all configured groups, then a representative work period with failures tested.

## Acceptance on the actual workstation

- Full cells and all rows match ABS, including repeated case numbers/tasks.
- Source total means all rows in the agreed scope and matches the DataTable.
- Source time zone, action-time buckets, and blank assignment meaning agree.
- A valid empty group shows zero. A failure shows failed/old evidence.
- Session expiry, grid changes, scroll overlap, and interruption never publish
  a falsely complete capture.
- A full scan finishes within the selected cadence without overlapping.

The Python test suite exercises the publisher-to-inbox-to-dashboard path with
synthetic records. It cannot verify private ABS selectors or the Windows PAD
session; those checks remain part of this workstation setup.
