@echo off
setlocal EnableExtensions DisableDelayedExpansion
chcp 65001 >nul
title Mr.Propre - Made by Anass Assri
set "PYTHONHOME="
set "PYTHONPATH="
set "MR_PROPRE_PYTHON=%~dp0_system\runtime\python.exe"
if not exist "%MR_PROPRE_PYTHON%" (
    echo ERROR: Embedded Python is missing. Extract the entire Mr.Propre release ZIP first.
    if "%~1"=="" pause
    exit /b 1
)
"%MR_PROPRE_PYTHON%" -I "%~dp0_system\app\mr_propre.py" %*
set "MR_PROPRE_RESULT=%ERRORLEVEL%"
if "%~1"=="" (
    echo.
    pause
)
exit /b %MR_PROPRE_RESULT%
