@echo off
cd /d "%~dp0"
if not exist ".venv\Scripts\pythonw.exe" (echo Ejecuta primero probar_api.bat para instalar. & pause & exit /b 1)
".venv\Scripts\python.exe" -m pip install -q -r requirements.txt
start "" ".venv\Scripts\pythonw.exe" main.py
