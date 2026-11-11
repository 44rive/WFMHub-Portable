@echo off
setlocal EnableExtensions
set "BO_SCRIPT=%~dp0app.py"
if exist "%~dp0bo_live\app.py" set "BO_SCRIPT=%~dp0bo_live\app.py"
if exist "%~dp0runtime\python.exe" goto :standalone
if exist "%~dp0..\..\_system\runtime\python.exe" goto :existing_hub
where py >nul 2>nul
if not errorlevel 1 goto :python_launcher
where python >nul 2>nul
if not errorlevel 1 goto :system_python
echo BO Live requires the complete Windows ZIP or Python 3.10 or later.
exit /b 9009

:standalone
"%~dp0runtime\python.exe" -I "%BO_SCRIPT%" %*
exit /b %ERRORLEVEL%
:existing_hub
"%~dp0..\..\_system\runtime\python.exe" -I "%BO_SCRIPT%" %*
exit /b %ERRORLEVEL%
:python_launcher
py -3 -I "%BO_SCRIPT%" %*
exit /b %ERRORLEVEL%
:system_python
python -I "%BO_SCRIPT%" %*
exit /b %ERRORLEVEL%
