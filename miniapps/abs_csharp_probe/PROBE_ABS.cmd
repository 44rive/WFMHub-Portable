@echo off
setlocal EnableExtensions
cd /d "%~dp0"
title WFMHub C# ABS read-only diagnostic
echo Open the TOYOTA backlog in Edge first. Keep ABS as the active tab.
echo The diagnostic reads accessibility metadata without clicking or scrolling.
pause
"%~dp0_system\collector\ABSProbe.exe" --output "%~dp0_system\output\abs-csharp-probe" %*
set "BO_CHECK_EXIT=%ERRORLEVEL%"
echo.
if not "%BO_CHECK_EXIT%"=="0" echo Read the error above. A report exists only if its new path was printed above.
echo Probe exit code: %BO_CHECK_EXIT%
pause
exit /b %BO_CHECK_EXIT%
