@echo off
REM Daily runner for Windows Task Scheduler.
REM Runs from this script's own folder, using the project's virtualenv.
set SIFMA=US-Corporate-Bonds-Statistics-SIFMA.xlsx

cd /d "%~dp0"
".venv\Scripts\python.exe" credit_warning_monitor.py --sifma-file "%SIFMA%" --html report.html --json latest.json > last_run.log 2>&1
