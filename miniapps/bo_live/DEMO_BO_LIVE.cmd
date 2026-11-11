@echo off
title WFMHub BO Live - SYNTHETIC DEMO
call "%~dp0RUN_PYTHON.cmd" serve --demo --port 8788 %*
if errorlevel 1 pause
