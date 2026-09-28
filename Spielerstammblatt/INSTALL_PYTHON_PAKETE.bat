@echo off
py -m pip install --upgrade pymupdf pytesseract pillow openpyxl
if errorlevel 1 python -m pip install --upgrade pymupdf pytesseract pillow openpyxl
pause
