@echo off
set "drive=%~d0"
if not exist "%drive%\autorun.bat" goto end
PowerShell.exe -ExecutionPolicy Bypass -File "%~dp0\scriptv2.ps1"
pause
:end
exit
