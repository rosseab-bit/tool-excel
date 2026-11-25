@echo off
rem Doble clic o linea de comandos: llama a build_windows.ps1 sin que PowerShell bloquee el script.
setlocal
set "DOBLECLIC=0"
echo %cmdcmdline% | find /i "/c" >nul && set "DOBLECLIC=1"

powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0build_windows.ps1" %*
set "CODIGO=%ERRORLEVEL%"

if not "%CODIGO%"=="0" (
    echo.
    echo Algo salio mal ^(codigo %CODIGO%^). Lee el mensaje de arriba.
    if "%DOBLECLIC%"=="1" pause
) else (
    if "%DOBLECLIC%"=="1" pause
)
exit /b %CODIGO%
