@echo off
setlocal EnableExtensions
title WFMHub ABS read-only probe
cd /d "%~dp0"
echo Open ABS in Edge and navigate to the TOYOTA backlog first.
echo Keep that Edge window visible with ABS as its active tab.
echo This probe reads accessibility information only. It does not click or scroll.
echo.
pause
"%~dp0_system\runtime\python.exe" -I "%~dp0_system\app\abs_probe.py" --output "%~dp0_system\output\abs-probe" %*
set "BO_PROBE_EXIT=%ERRORLEVEL%"
echo.
echo Reports are in _system\output\abs-probe.
pause
exit /b %BO_PROBE_EXIT%
