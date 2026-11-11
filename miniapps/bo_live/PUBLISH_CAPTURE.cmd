@echo off
call "%~dp0RUN_PYTHON.cmd" publish %*
exit /b %ERRORLEVEL%
