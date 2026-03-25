@echo off
title Sync Crocoveen

echo.
echo  ==========================================
echo    CROCOVEEN - Sync semanal
echo  ==========================================
echo.

cd /d "%~dp0"

python --version > nul 2>&1
if errorlevel 1 (
    echo  [ERROR] Python no esta instalado.
    echo.
    echo  Descargalo en: https://www.python.org/downloads/
    echo  Marca la opcion "Add Python to PATH" al instalar.
    echo.
    pause
    exit /b 1
)

echo  Verificando dependencias...
pip install openpyxl --quiet 2>nul

echo  Ejecutando sync...
echo.
python sync_crocoveen.py

echo.
echo  ==========================================
echo    Listo.
echo  ==========================================
echo.
pause
