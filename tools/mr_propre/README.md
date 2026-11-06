# Mr.Propre — Portable Windows x64

**Made by Anass Assri** · Version 1.0.1

Extract the entire ZIP, then double-click **Mr.Propre.cmd**. Python and the Excel
reader are included. No installation, administrator rights, internet connection,
WFMHub installation, or first-run runtime preparation is needed.

1. Paste the source folder containing your CSV or Verint TXT extracts.
2. Paste your employee listing file (`FTE Count.xlsx`, CSV or TXT). You can also
   paste its folder and choose the listing by number.
3. Choose a separate output folder, or press Enter for a new timestamped folder
   beside the source folder.
4. Select individual files (`1,3`), ranges (`2-5`), or `A` for all.

Subfolders are scanned, and their structure is preserved in the output. Each
chosen file produces a CSV with its original header names, column order, field
values, duplicate rows, and row order; only unmatched rows are removed. Output is
comma-delimited UTF-8 with a BOM for Excel, so encoding, quoting, and delimiters
can differ from the source. TXT is converted to CSV. Sources are never edited and
existing output files are never overwritten.

## Exactly what “clean” means here

**Only Client ID membership in your listing matters.** Status, leaver dates,
LOB, team leader, names, and evidence dates are ignored. A one-column Client ID
listing is sufficient. Duplicate listing IDs are accepted once; blank listing
IDs are ignored. This is a standalone extract filter, separate from WFMHub's
database, calculated feeds, Active/Leaver eligibility rules, and report refresh.

IDs are compared as text, case-sensitively, with surrounding whitespace ignored
for comparison only. `00123` and `123` remain different. XLSX numeric IDs with a
plain zero format (e.g. `00000`) retain their displayed leading zeros. Store IDs
as text in Excel when they exceed Excel's numeric precision. Original source
cells are always written unchanged. No name fallback or fuzzy matching is used.

Headers accept punctuation/case variations of `Client ID`, `Agent ID`,
`Data Source ID`, `Data Source IDs`, and `Verint ID`. An explicit `Client ID`
column takes precedence if other operational ID columns also exist. Multiple
equivalent ID columns are rejected rather than guessed.

For XLSX, the first 100 rows of each sheet are searched for the ID header.
Agent/Roster/FTE/Headcount sheets take precedence; if several sheets are equally
likely, pass `--sheet "Agent"`. Other workbook sheets are not used as ID lists.
The workbook is read without executing macros. Formula IDs require cached values
from Excel; prefer literal text IDs.

CSV/TXT supports UTF-8 (with/without BOM), Windows-1252, and BOM-marked UTF-16,
with comma, semicolon, tab, or pipe separators. The ID header must be on the
first line of a text file. Files without an ID header (e.g. staff-type forecasts)
are reported as failed and are not copied unfiltered. Malformed rows fail that
file; no incomplete CSV is left. Other selected files continue processing.

## Command line

```bat
Mr.Propre.cmd --folder "C:\Extracts" --roster "C:\FTE\FTE Count.xlsx" --output "C:\Filtered" --all
Mr.Propre.cmd --folder "C:\Extracts" --roster "C:\FTE\FTE Count.xlsx" --sheet Agent --output "C:\Filtered" --select "1,3-5"
```

Use `--flat` to scan only the immediate source folder. Exit status: 0 success,
2 invalid input or any failed files, 130 cancelled. Ctrl+C cancels a selection.
The output folder must be outside the source tree and must not contain the
listing. If `name.csv` and `name.txt` would produce the same output name, process
them separately into different output folders.

## Runtime and source

Includes CPython 3.13.7 x64, openpyxl 3.1.5, and et_xmlfile 2.0.0, from the
SHA-256-verified WFMHub Portable v1.1.2 release. CPython's license is in
`_system/runtime/LICENSE.txt`; library licenses are inside their wheels.
`_system/RUNTIME_ORIGIN.txt` records the source archive and hashes.

Source, synthetic regression tests, and reproducible packaging script are under
`tools/mr_propre` in `44rive/WFMHub-Portable`. No employee data, database, or local
configuration is included in this release.
