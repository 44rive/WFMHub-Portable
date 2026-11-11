@echo off
title WFMHub BO Live - TOYOTA
call "%~dp0RUN_PYTHON.cmd" serve %*
if errorlevel 1 pause
