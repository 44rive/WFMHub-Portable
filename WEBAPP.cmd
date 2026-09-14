@echo off
setlocal EnableExtensions
title WFMHub Manager Workbench
if not defined NO_COLOR color 0B
for %%I in ("%~dp0.") do set "WFMHUB_HOME=%%~fI"
set "WFMHUB_PYTHON=%WFMHUB_HOME%\_system\runtime\python.exe"
set "PYTHONHOME="
set "PYTHONPATH="

if not exist "%WFMHUB_PYTHON%" goto :missing_runtime

"%WFMHUB_PYTHON%" -I -m wfmhub --home "%WFMHUB_HOME%" web
set "EXIT_CODE=%ERRORLEVEL%"
if not "%EXIT_CODE%"=="0" pause
exit /b %EXIT_CODE%

:missing_runtime
echo.
echo ERROR: WFMHub's embedded Python is missing:
echo   "%WFMHUB_PYTHON%"
echo.
echo Extract the full WFMHub Portable Windows ZIP before using WEBAPP.cmd.
echo.
pause
exit /b 9009
