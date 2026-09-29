@echo off
rem One-click entry for the receipt automation pipeline.
rem run.py picks the Python interpreter that has the dependencies installed.
rem "report" mode = run_report.py wrapper: batch archive to 40_ + run log to 70_ + stage 30_.
cd /d "%~dp0"
python run.py report
pause
