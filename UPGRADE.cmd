@echo off
setlocal EnableExtensions
title WFMHub Portable - Upgrade Existing Data
if not defined NO_COLOR color 0B
for %%I in ("%~dp0.") do set "WFMHUB_HOME=%%~fI"
set "WFMHUB_PYTHON=%WFMHUB_HOME%\_system\runtime\python.exe"

if not exist "%WFMHUB_PYTHON%" goto :missing_runtime

echo.
echo Enter the full folder path of your PREVIOUS WFMHub installation.
echo The old folder will only be read. Nothing there will be deleted.
set /p "WFMHUB_OLD=Previous WFMHub folder: "
set "WFMHUB_OLD=%WFMHUB_OLD:"=%"
if not defined WFMHUB_OLD goto :missing_path

"%WFMHUB_PYTHON%" -m wfmhub --home "%WFMHUB_HOME%" upgrade-install --from "%WFMHUB_OLD%"
set "EXIT_CODE=%ERRORLEVEL%"
pause
exit /b %EXIT_CODE%

:missing_path
echo ERROR: No previous WFMHub folder was entered.
pause
exit /b 1

:missing_runtime
echo.
echo ERROR: WFMHub's embedded Python is missing:
echo   "%WFMHUB_PYTHON%"
echo.
echo Download the WFMHub-Portable Windows ZIP from GitHub Releases,
echo choose Extract All, and keep UPGRADE.cmd beside the _system folder.
echo Do not use GitHub's Source code ZIP and do not copy UPGRADE.cmd by itself.
echo.
pause
exit /b 9009
