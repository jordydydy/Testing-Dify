@echo off
cd /d "%~dp0"
python -c "import requests, openpyxl" 2>nul || (echo Installing requests and openpyxl... & python -m pip install -r requirements.txt)
python procurement_tester.py
pause
