@echo off
setlocal EnableExtensions EnableDelayedExpansion
title WFM PCS - 3-hourly refresh

REM ============================================================================
REM  WFM PCS Daily Report - unattended refresh
REM
REM  Runs every 3 hours. Re-ingests FTE + Call by Call only, rebuilds the PCS
REM  model for the reporting window, and exports one governed file that the
REM  Excel workbook reads.
REM
REM  READ-ONLY against the source extracts. Never writes to them.
REM  Does not modify anything inside the WFMHub-Portable folder except its own
REM  database and log, which is what a refresh is supposed to do.
REM ============================================================================

REM ---- 1. Configure -----------------------------------------------------------
REM Point WFMHUB_HOME at the folder that contains WFMHub.cmd and runtime\.
REM Edit this one line if your hub lives somewhere else.

set "WFMHUB_HOME=%USERPROFILE%\WFMHub-Portable"

set "PCS_HOME=%~dp0"
if "%PCS_HOME:~-1%"=="\" set "PCS_HOME=%PCS_HOME:~0,-1%"

set "WFMHUB_PYTHON=%WFMHUB_HOME%\runtime\python.exe"
set "CURRENT_DIR=%PCS_HOME%\data\current"
set "ARCHIVE_DIR=%PCS_HOME%\data\archive"
set "LOG_DIR=%PCS_HOME%\logs"
set "TARGET=%CURRENT_DIR%\pcs_agent_day.csv"

if not exist "%CURRENT_DIR%" mkdir "%CURRENT_DIR%"
if not exist "%ARCHIVE_DIR%" mkdir "%ARCHIVE_DIR%"
if not exist "%LOG_DIR%"     mkdir "%LOG_DIR%"

REM ---- 2. Preflight -----------------------------------------------------------
if not exist "%WFMHUB_PYTHON%" (
    echo.
    echo ERROR: WFMHub's embedded Python was not found at:
    echo   "%WFMHUB_PYTHON%"
    echo.
    echo Edit WFMHUB_HOME at the top of this file so it points at the folder
    echo containing WFMHub.cmd and the runtime folder.
    echo.
    if not "%1"=="/quiet" pause
    exit /b 9009
)

REM ---- 3. Reporting window ----------------------------------------------------
REM The hub rebuilds marts in full: DELETE FROM mart.agent_pcs_day has no WHERE
REM clause, so the mart ends up holding EXACTLY the period just refreshed.
REM
REM The window must therefore span everything the report needs, every run.
REM 1st of LAST month through today keeps MTD, rolling-7 across the month
REM boundary, and a prior-month comparison all available.

for /f "usebackq tokens=1,2" %%A in (`
    "%WFMHUB_PYTHON%" -c "import datetime as d;t=d.date.today();s=(t.replace(day=1)-d.timedelta(days=1)).replace(day=1);print(s.isoformat(),t.isoformat())"
`) do (
    set "WINDOW_START=%%A"
    set "WINDOW_END=%%B"
)

for /f "usebackq tokens=1,2" %%A in (`
    "%WFMHUB_PYTHON%" -c "import datetime as d;n=d.datetime.now();print(n.strftime('%%Y%%m%%d_%%H%%M%%S'),n.strftime('%%Y-%%m-%%d_%%H%%M'))"
`) do (
    set "STAMP=%%A"
    set "STAMP_HUMAN=%%B"
)

set "LOGFILE=%LOG_DIR%\pcs_refresh_%STAMP%.log"

echo ============================================================ >>"%LOGFILE%"
echo  WFM PCS refresh   %STAMP_HUMAN%                             >>"%LOGFILE%"
echo  Window: %WINDOW_START% to %WINDOW_END%                      >>"%LOGFILE%"
echo ============================================================ >>"%LOGFILE%"

echo.
echo   WFM PCS refresh
echo   Window : %WINDOW_START%  to  %WINDOW_END%
echo   Log    : %LOGFILE%
echo.

REM ---- 4. Refresh the PCS model ----------------------------------------------
REM --source-group pcs  = FTE roster + Call by Call ONLY. Skips Verint
REM                       schedules, LILO and the three Storm KPI feeds.
REM --no-report         = build the model, do not write hub workbooks.

echo   [1/3] Refreshing PCS model...
"%WFMHUB_PYTHON%" -m wfmhub --home "%WFMHUB_HOME%" refresh ^
    --source-group pcs --no-report ^
    --start %WINDOW_START% --end %WINDOW_END% >>"%LOGFILE%" 2>&1

if errorlevel 1 (
    echo   FAILED - refresh returned an error. See the log.
    echo   REFRESH FAILED >>"%LOGFILE%"
    echo   Nothing was changed in your extract files.
    if not "%1"=="/quiet" pause
    exit /b 1
)

REM ---- 5. Archive the previous export -----------------------------------------
REM Keeps an audit trail without leaving a second file in data\current\,
REM where the Power Query folder source would combine it and double every count.

if exist "%TARGET%" (
    echo   [2/3] Archiving previous export...
    move /y "%TARGET%" "%ARCHIVE_DIR%\pcs_agent_day_%STAMP%.csv" >>"%LOGFILE%" 2>&1
    if exist "%TARGET%.manifest.txt" (
        move /y "%TARGET%.manifest.txt" "%ARCHIVE_DIR%\pcs_agent_day_%STAMP%.csv.manifest.txt" >>"%LOGFILE%" 2>&1
    )
) else (
    echo   [2/3] No previous export to archive.
)

REM ---- 6. Export ---------------------------------------------------------------
echo   [3/3] Exporting pcs_agent_day...
"%WFMHUB_PYTHON%" -m wfmhub --home "%WFMHUB_HOME%" export pcs_agent_day ^
    --start %WINDOW_START% --end %WINDOW_END% ^
    --format csv --output "%TARGET%" >>"%LOGFILE%" 2>&1

if errorlevel 1 (
    echo   FAILED - export returned an error. See the log.
    echo   EXPORT FAILED >>"%LOGFILE%"
    if not "%1"=="/quiet" pause
    exit /b 2
)

if not exist "%TARGET%" (
    echo   FAILED - export reported success but produced no file.
    if not "%1"=="/quiet" pause
    exit /b 3
)

REM ---- 7. Guard: data\current\ must hold exactly ONE csv -----------------------
REM If a second csv is ever dropped here, the Power Query folder source combines
REM both files and every count doubles silently. Fail loudly instead.

set /a CSVCOUNT=0
for %%F in ("%CURRENT_DIR%\*.csv") do set /a CSVCOUNT+=1

if not "%CSVCOUNT%"=="1" (
    echo.
    echo   WARNING: data\current\ contains %CSVCOUNT% csv files. It must contain
    echo   exactly one. The workbook will double-count until this is fixed.
    echo   Move the extras into data\archive\.
    echo   WARNING: %CSVCOUNT% csv files in data\current\ >>"%LOGFILE%"
)

REM ---- 8. Freshness stamp ------------------------------------------------------
> "%CURRENT_DIR%\last_refresh.txt" (
    echo Data as at : %STAMP_HUMAN%
    echo Window     : %WINDOW_START% to %WINDOW_END%
    echo Source     : WFMHub mart.agent_pcs_day via export pcs_agent_day
    echo Note       : the current day is PARTIAL - responses arrive after the call
)

REM ---- 9. Prune archives older than 60 days -----------------------------------
forfiles /p "%ARCHIVE_DIR%" /m *.csv /d -60 /c "cmd /c del @path" >nul 2>&1
forfiles /p "%LOG_DIR%" /m *.log /d -60 /c "cmd /c del @path" >nul 2>&1

echo.
echo   DONE  %STAMP_HUMAN%
echo   Export : %TARGET%
echo.
echo   Open the workbook and press Refresh All.
echo.

echo   OK %STAMP_HUMAN% >>"%LOGFILE%"

if not "%1"=="/quiet" (
    timeout /t 8 >nul
)
exit /b 0
