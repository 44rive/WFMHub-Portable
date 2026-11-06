Mr.Propre 1.0.1 — Portable Windows x64

**Made by Anass Assri**

Download **Mr.Propre-Portable-v1.0.1-win-x64.zip**, extract the entire ZIP, then
double-click **Mr.Propre.cmd**. CPython 3.13.7 and the Excel reader are included.
No Python installation, WFMHub installation, runtime preparation, administrator
rights, or internet connection is needed.

- Select source CSV/Verint TXT files and an XLSX/CSV/TXT employee listing.
- Keep only rows whose Client ID is in the listing. Status, dates, LOB, and team
  leader are ignored.
- Preserve original column names, column order, field values, duplicate rows,
  and row order.
- Write filtered UTF-8 CSV copies into a separate output folder without changing
  sources or overwriting existing outputs.
- Fix UTF-8 TXT corruption, rejection of multiple date columns, and the mandatory
  Status-column requirement.

Includes a SHA-256 checksum and source/tests archive. The runtime comes from the
SHA-256-verified WFMHub Portable v1.1.2 release. Source: `tools/mr_propre` at this
release tag. WFMHub's existing database, reports, and eligibility rules are unchanged.

Validation: 135 tests passed locally. All 18 Mr.Propre tests passed on Windows
using the bundled interpreter, including XLSX reading and the CMD launcher from
an extracted release in paths containing spaces. The downloadable ZIP is the
Windows-tested artifact; its checksum is verified before publication.
