@echo off
rem Instalador de doble clic. Acepta los mismos parametros que install.ps1, por ejemplo:
rem   instalar.cmd -Workspace "C:\Users\ella\Documents\Claude-Ella"
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0install.ps1" %*
echo.
pause
