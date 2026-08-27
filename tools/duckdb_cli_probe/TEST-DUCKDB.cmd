@echo off
setlocal EnableExtensions
for %%I in ("%~dp0.") do set "WFM_PROBE_HOME=%%~fI"
set "WFM_DUCKDB=%WFM_PROBE_HOME%\duckdb.exe"

echo.
echo WFMHub DuckDB CLI policy test
echo ==============================
echo This test does not read or change any WFM extract.
echo.

if not exist "%WFM_DUCKDB%" goto :missing_cli

echo Starting the officially signed DuckDB CLI...
"%WFM_DUCKDB%" -batch -noheader -list -c "SELECT 'DUCKDB_CLI_OK' AS policy_test, version() AS duckdb_version;"
set "WFM_EXIT_CODE=%ERRORLEVEL%"

if not "%WFM_EXIT_CODE%"=="0" goto :blocked

echo.
echo TEST PASSED: DUCKDB_CLI_OK
echo Send this exact message: CLI works
echo.
pause
exit /b 0

:missing_cli
echo ERROR: duckdb.exe is missing from this folder.
echo Extract the complete policy-probe ZIP and run this file again.
echo.
pause
exit /b 2

:blocked
echo.
echo TEST FAILED with exit code %WFM_EXIT_CODE%.
echo Do not disable or bypass company security controls.
echo Send the complete error or a screenshot so WFMHub can use another backend.
echo.
pause
exit /b %WFM_EXIT_CODE%
