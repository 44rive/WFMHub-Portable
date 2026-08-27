@echo off
setlocal EnableExtensions
title WFMHub Portable - Setup
if not defined NO_COLOR color 0B
for %%I in ("%~dp0.") do set "WFMHUB_HOME=%%~fI"
set "WFMHUB_PYTHON=%WFMHUB_HOME%\runtime\python.exe"

if not exist "%WFMHUB_PYTHON%" goto :missing_runtime

"%WFMHUB_PYTHON%" -m wfmhub --home "%WFMHUB_HOME%" doctor
if errorlevel 1 goto :doctor_failed

"%WFMHUB_PYTHON%" -m wfmhub --home "%WFMHUB_HOME%" setup
set "EXIT_CODE=%ERRORLEVEL%"
pause
exit /b %EXIT_CODE%

:missing_runtime
echo.
echo ERROR: WFMHub's embedded Python is missing:
echo   "%WFMHUB_PYTHON%"
echo.
echo Download the WFMHub-Portable Windows ZIP from GitHub Releases,
echo choose Extract All, and keep SETUP.cmd beside the runtime folder.
echo Do not use GitHub's Source code ZIP and do not copy SETUP.cmd by itself.
echo.
pause
exit /b 9009

:doctor_failed
echo.
echo SETUP STOPPED: the system check failed before creating the WFMHub database.
echo Read the FAIL line above or send it to your WFM support contact.
echo Your extracts and old database were not changed.
echo.
pause
exit /b 1
