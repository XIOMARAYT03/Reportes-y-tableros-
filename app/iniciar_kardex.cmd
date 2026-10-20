@echo off
setlocal
cd /d "%~dp0"
title Kardex Fimas

echo.
echo   KARDEX FIMAS - Iniciando servidor local...
echo   ---------------------------------------------------------
echo.

where python >nul 2>nul
if errorlevel 1 (
  echo   [ERROR] No se encontro Python en el PATH.
  echo   Instale Python 3.10 o superior y vuelva a ejecutar este archivo.
  echo.
  pause
  exit /b 1
)

start "" "http://localhost:8765/"
python kardex_server.py

echo.
echo   El servidor se detuvo. Para volver a iniciar, ejecute este archivo de nuevo.
echo.
pause