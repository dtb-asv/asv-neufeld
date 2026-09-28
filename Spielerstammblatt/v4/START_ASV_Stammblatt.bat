@echo off
cd /d "%~dp0"
py asv_stammblatt_erfassung.py
if errorlevel 1 pause
