@echo off
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" (echo Ejecuta primero probar_api.bat & pause & exit /b 1)
".venv\Scripts\python.exe" -m pip install -q -r requirements.txt
set /p AMB=Ambiente SAP (QAS/PRD) [PRD]: 
if "%AMB%"=="" set AMB=PRD
".venv\Scripts\python.exe" -m scripts.test_gr55 %AMB%
echo.
pause
