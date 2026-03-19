@echo off
:: This line tells the CMD to look at the folder where this .bat file is saved
cd /d "%~dp0"

:: This runs your specific command
python -m streamlit run .\RIs_v2.py

pause