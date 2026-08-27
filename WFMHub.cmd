@echo off
setlocal
set "WFMHUB_HOME=%~dp0"
set "PYTHONPATH=%~dp0app;%~dp0src"
if exist "%~dp0runtime\python.exe" (
  "%~dp0runtime\python.exe" -m wfmhub --home "%~dp0" menu
) else (
  py -3.13 -m wfmhub --home "%~dp0" menu
)
set "EXIT_CODE=%ERRORLEVEL%"
if not "%EXIT_CODE%"=="0" pause
exit /b %EXIT_CODE%
