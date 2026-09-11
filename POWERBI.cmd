@echo off
setlocal EnableExtensions
title WFMHub Portable - Power BI
if not defined NO_COLOR color 0B
for %%I in ("%~dp0.") do set "WFMHUB_HOME=%%~fI"
set "WFMHUB_PYTHON=%WFMHUB_HOME%\_system\runtime\python.exe"

if not exist "%WFMHUB_PYTHON%" goto :missing_runtime

"%WFMHUB_PYTHON%" -m wfmhub --home "%WFMHUB_HOME%" powerbi open
set "EXIT_CODE=%ERRORLEVEL%"
if not "%EXIT_CODE%"=="0" pause
exit /b %EXIT_CODE%

:missing_runtime
echo.
echo ERROR: WFMHub's embedded Python is missing:
echo   "%WFMHUB_PYTHON%"
echo.
echo Download the complete WFMHub Portable ZIP from GitHub Releases and choose Extract All.
echo Do not copy POWERBI.cmd by itself.
echo.
pause
exit /b 9009
