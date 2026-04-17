@echo off
title WinSysMon Manager
cd /d "%~dp0"
echo.
echo === Windows System Monitor ===
echo.
echo  1) Instalar servico
echo  2) Remover servico
echo  3) Ver status
echo  4) Sair
echo.
set /p opt=Opcao: 
if "%opt%"=="1" powershell.exe -ExecutionPolicy Bypass -File "%~dp0winsysmon.ps1" -Install
if "%opt%"=="2" powershell.exe -ExecutionPolicy Bypass -File "%~dp0winsysmon.ps1" -Uninstall
if "%opt%"=="3" powershell.exe -ExecutionPolicy Bypass -File "%~dp0winsysmon.ps1" -Status
if "%opt%"=="4" exit /b
pause
