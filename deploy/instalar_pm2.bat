@echo off
setlocal EnableDelayedExpansion
chcp 65001 >nul
title Kardex Fimas - Instalador de PM2

REM ============================================================
REM  KARDEX FIMAS - Instalador de PM2 (Windows Server)
REM
REM  Instala Node.js (si falta), PM2 y el arranque automatico de
REM  PM2, y registra la aplicacion Kardex bajo PM2.
REM
REM  Ejecutar como Administrador (clic derecho > Ejecutar como admin).
REM ============================================================

REM ---------------------- Configuracion ------------------------
set "APP_DIR=C:\Kardex-ERP"
set "APP_NOMBRE=kardex"
set "PUERTO=8765"
set "ENV_PATH=%APP_DIR%\.env"
REM -------------------------------------------------------------

REM Verificar privilegios de administrador
net session >nul 2>&1
if errorlevel 1 (
  echo.
  echo   [ERROR] Ejecute este archivo como Administrador.
  echo           Clic derecho ^> "Ejecutar como administrador".
  echo.
  pause
  exit /b 1
)

echo.
echo   ===========================================================
echo        KARDEX FIMAS - Instalacion de PM2
echo   ===========================================================
echo.

REM ---------------------- Node.js ------------------------------
where node >nul 2>&1
if errorlevel 1 goto INSTALAR_NODE
for /f "delims=" %%v in ('node --version 2^>nul') do set "NODEVER=%%v"
echo   [OK] Node.js detectado: !NODEVER!
goto NODE_OK

:INSTALAR_NODE
echo   [!] Node.js no esta instalado. Intentando instalar Node.js LTS...
where winget >nul 2>&1
if errorlevel 1 goto SIN_WINGET
winget install -e --id OpenJS.NodeJS.LTS --accept-source-agreements --accept-package-agreements
if errorlevel 1 goto SIN_WINGET
set "PATH=%ProgramFiles%\nodejs;%PATH%"
where node >nul 2>&1
if errorlevel 1 (
  echo.
  echo   [ERROR] Node.js se instalo pero no aparece en el PATH de esta sesion.
  echo           Cierre esta ventana, abra una NUEVA como Administrador y reintente.
  echo.
  pause
  exit /b 1
)
echo   [OK] Node.js instalado.
goto NODE_OK

:SIN_WINGET
echo.
echo   [ERROR] No se encontro Node.js ni winget.
echo           Instale Node.js LTS manualmente desde:
echo               https://nodejs.org/en/download
echo           y vuelva a ejecutar este archivo.
echo.
pause
exit /b 1

:NODE_OK
REM ---------------------- PM2 ----------------------------------
echo.
echo   Instalando PM2...
call npm install -g pm2 pm2-windows-startup
if errorlevel 1 (
  echo.
  echo   [ERROR] No se pudo instalar PM2 con npm. Revise su conexion a Internet.
  echo.
  pause
  exit /b 1
)
echo   [OK] PM2 instalado.

REM Arranque automatico de PM2 al iniciar el servidor
echo.
echo   Configurando el arranque automatico de PM2...
call pm2-startup install
if errorlevel 1 (
  echo   [!] No se pudo configurar el arranque automatico ^(pm2-startup^).
  echo       Podra iniciar PM2 manualmente con: pm2 resurrect
) else (
  echo   [OK] PM2 arrancara automaticamente al iniciar el servidor.
)

REM ---------------------- Aplicacion Kardex --------------------
if not exist "%APP_DIR%\kardex_server.py" goto SIN_APP

echo.
echo   Registrando la aplicacion Kardex en PM2...
set "KARDEX_HOST=0.0.0.0"
set "KARDEX_PORT=%PUERTO%"
set "KARDEX_ENV_PATH=%ENV_PATH%"
call pm2 delete %APP_NOMBRE% >nul 2>&1
call pm2 start "%APP_DIR%\kardex_server.py" --name %APP_NOMBRE% --interpreter python --cwd "%APP_DIR%"
if errorlevel 1 (
  echo   [!] No se pudo iniciar la aplicacion en PM2. Revise "pm2 logs".
) else (
  call pm2 save >nul 2>&1
  echo   [OK] Aplicacion registrada en PM2 ^(nombre: %APP_NOMBRE%^).
)
goto RESUMEN

:SIN_APP
echo.
echo   [!] No se encontro "%APP_DIR%\kardex_server.py".
echo       Solo se instalo PM2. Ajuste APP_DIR al inicio de este archivo.

:RESUMEN
echo.
echo   ===========================================================
echo    INSTALACION COMPLETADA
echo   ===========================================================
echo.
echo    Estado .......... pm2 status
echo    Logs ............ pm2 logs %APP_NOMBRE%
echo    Reiniciar ....... pm2 restart %APP_NOMBRE%
echo    Guardar ......... pm2 save
echo.
echo    IMPORTANTE: si la app ya corre como tarea programada de
echo    Windows ^(KardexERP^), detengala para no duplicar el puerto %PUERTO%:
echo        schtasks /End /TN KardexERP
echo.
pause
endlocal
