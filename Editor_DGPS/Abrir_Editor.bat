@echo off
cd /d "%~dp0"
py -3 Editor_DGPS.py
if errorlevel 1 python Editor_DGPS.py
