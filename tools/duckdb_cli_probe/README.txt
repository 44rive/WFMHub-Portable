WFMHub DuckDB CLI policy probe
==============================

Purpose
-------
This tiny package checks whether the work machine permits the official,
Authenticode-signed DuckDB command-line executable. It does not contain the
WFMHub application and does not read or modify any WFM extract.

How to test
-----------
1. Right-click the downloaded ZIP and choose Extract All.
2. Keep TEST-DUCKDB.cmd and duckdb.exe together in the extracted folder.
3. Double-click TEST-DUCKDB.cmd.
4. A successful result displays:

       TEST PASSED: DUCKDB_CLI_OK

5. Report either "CLI works" or send the complete failure message/screenshot.

What the test does
------------------
The test starts duckdb.exe and runs one harmless query in a transient in-memory
database. It creates no database file and never opens the WFMHub database or
source extracts.

Binary provenance
-----------------
DuckDB version: 1.5.5, Windows amd64 CLI
Official source:
https://github.com/duckdb/duckdb/releases/download/v1.5.5/duckdb_cli-windows-amd64.zip
Official ZIP SHA-256:
e1428b7114a841626b5054723731cbf45c6df91b42ae1a6c355f88fad1f6dc4c
duckdb.exe SHA-256:
fde737c7749075f6b54e14772a4e6b33a5fa0201075d03640aca358074ea4554

The embedded duckdb.exe is signed by Stichting DuckDB Foundation through the
Microsoft Identity Verified code-signing certificate chain. You can inspect
this under duckdb.exe Properties > Digital Signatures.

A valid signature does not guarantee that every corporate policy allows this
publisher; that is exactly what TEST-DUCKDB.cmd checks on the target machine.

DuckDB is redistributed under its MIT license in DUCKDB-LICENSE.txt.
