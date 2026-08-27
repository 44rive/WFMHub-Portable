@echo off
setlocal EnableExtensions
title WFMHub Portable
if not defined NO_COLOR color 0B
for %%I in ("%~dp0.") do set "WFMHUB_HOME=%%~fI"
set "WFMHUB_PYTHON=%WFMHUB_HOME%\runtime\python.exe"

if not exist "%WFMHUB_PYTHON%" goto :missing_runtime

"%WFMHUB_PYTHON%" -m wfmhub --home "%WFMHUB_HOME%" menu
set "EXIT_CODE=%ERRORLEVEL%"
if not "%EXIT_CODE%"=="0" pause
exit /b %EXIT_CODE%

:missing_runtime
echo.
echo ERROR: WFMHub's embedded Python is missing:
echo   "%WFMHUB_PYTHON%"
echo.
echo Download the WFMHub-Portable Windows ZIP from GitHub Releases,
echo choose Extract All, and keep WFMHub.cmd beside the runtime folder.
echo Do not use GitHub's Source code ZIP and do not copy WFMHub.cmd by itself.
echo.
pause
exit /b 9009
