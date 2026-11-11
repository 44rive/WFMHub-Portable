@echo off
setlocal EnableExtensions
cd /d "%~dp0"
title WFMHub C# launch check
"%~dp0_system\collector\ABSProbe.exe" --check-launch --output "%~dp0_system\output\abs-csharp-probe"
set "BO_CHECK_EXIT=%ERRORLEVEL%"
echo.
if not "%BO_CHECK_EXIT%"=="0" echo The executable could not complete. A report exists only if its new path was printed above.
echo Launch check exit code: %BO_CHECK_EXIT%
pause
exit /b %BO_CHECK_EXIT%
