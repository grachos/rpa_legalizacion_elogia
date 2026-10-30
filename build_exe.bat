@echo off
cd /d "%~dp0"
set PY=python
if exist ".venv\Scripts\python.exe" set PY=.venv\Scripts\python.exe
%PY% -m pip install -q pyinstaller
%PY% -m PyInstaller --noconfirm --clean --onefile --windowed ^
  --name RPA_Legalizacion_Elogia ^
  --collect-data customtkinter ^
  --collect-submodules pywinauto ^
  --collect-submodules comtypes ^
  --hidden-import win32timezone ^
  --add-data ".env.example;." ^
  main.py || exit /b 1
echo.
echo Listo: dist\RPA_Legalizacion_Elogia.exe
