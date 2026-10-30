@echo off
setlocal
cd /d "%~dp0"

where python >nul 2>&1 || (echo [ERROR] Python no esta instalado o no esta en el PATH. & pause & exit /b 1)

if not exist ".venv\Scripts\python.exe" (
    echo Creando entorno virtual...
    python -m venv .venv || (echo [ERROR] No se pudo crear el entorno virtual. & pause & exit /b 1)
)

echo Instalando dependencias...
".venv\Scripts\python.exe" -m pip install -q --upgrade pip
".venv\Scripts\python.exe" -m pip install -q -r requirements.txt || (echo [ERROR] Fallo la instalacion de dependencias. & pause & exit /b 1)

if not exist ".env" (
    copy ".env.example" ".env" >nul
    echo.
    echo Se creo el archivo .env. Pon tu API_KEY, guarda y cierra el Bloc de notas.
    notepad ".env"
)

echo.
".venv\Scripts\python.exe" -m scripts.test_api

echo.
pause
