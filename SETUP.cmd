@echo off
setlocal
set "WFMHUB_HOME=%~dp0"
set "PYTHONPATH=%~dp0app;%~dp0src"
if exist "%~dp0runtime\python.exe" (
  "%~dp0runtime\python.exe" -m wfmhub --home "%~dp0" setup
) else (
  py -3.13 -m wfmhub --home "%~dp0" setup
)
set "EXIT_CODE=%ERRORLEVEL%"
pause
exit /b %EXIT_CODE%
