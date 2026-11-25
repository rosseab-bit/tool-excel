@echo off
rem Doble clic o linea de comandos: llama a dev.ps1 sin que PowerShell bloquee el script.
setlocal EnableDelayedExpansion
set "DOBLECLIC=0"
rem Doble clic => cmd se abre con "/c". Se compara con sustitucion de cadenas (sin echo ni pipes),
rem asi no rompen las comillas ni los "^" que puede traer CMDCMDLINE (AutoRun, doskey, etc.).
set "LINEA=!CMDCMDLINE!"
set "SINC=!LINEA: /c =!"
if not "!SINC!"=="!LINEA!" set "DOBLECLIC=1"

powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0dev.ps1" %*
set "CODIGO=!ERRORLEVEL!"

if not "!CODIGO!"=="0" (
    echo.
    echo Algo salio mal, codigo !CODIGO!. Lee el mensaje de arriba.
    if "!DOBLECLIC!"=="1" pause
)
exit /b !CODIGO!
