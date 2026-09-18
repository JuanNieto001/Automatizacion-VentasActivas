@echo off
REM ===================================================================
REM  Verificacion de ventas en AC - lanzador
REM  Doble clic para procesar el archivo mas reciente de la carpeta
REM  "entrada" usando la columna E.
REM ===================================================================
cd /d "%~dp0"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Automatizar-Ventas.ps1" %*
echo.
pause
