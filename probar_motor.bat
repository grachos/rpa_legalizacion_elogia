@echo off
setlocal
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" (echo Ejecuta primero probar_api.bat & pause & exit /b 1)
set "CSV="
for /f "delims=" %%F in ('dir /b /o:d "output\viaje_sap_*.csv" 2^>nul') do set "CSV=output\%%F"
if not defined CSV (echo No hay CSV en output. Ejecuta primero probar_api.bat & pause & exit /b 1)
echo Usando %CSV%
".venv\Scripts\python.exe" -m scripts.test_engine "%CSV%" "Gastos varios y peajes.xlsx"
echo.
pause
