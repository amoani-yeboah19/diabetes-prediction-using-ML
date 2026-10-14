@echo off
setlocal
cd /d "%~dp0"
py -3.12 -m venv .venv
if errorlevel 1 exit /b 1
call .venv\Scripts\activate.bat
python -m pip install -r requirements-dev.txt
if errorlevel 1 exit /b 1
python train.py
if errorlevel 1 exit /b 1
python -m flask --app app.app run
