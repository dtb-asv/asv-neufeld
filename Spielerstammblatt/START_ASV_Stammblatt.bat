@echo off
cd /d "%~dp0"
py asv_stammblatt_importer.py
if errorlevel 1 python asv_stammblatt_importer.py
pause
