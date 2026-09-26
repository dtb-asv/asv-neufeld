@echo off
cd /d "%~dp0"
python asv_stammblatt_importer.py
if errorlevel 1 pause
