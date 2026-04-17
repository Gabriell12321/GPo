@echo off
title IPPEL Agent
echo ========================================
echo   IPPEL Agent - Gerenciador
echo ========================================
echo.
echo  [1] Instalar servico
echo  [2] Desinstalar servico
echo  [3] Ver status
echo  [4] Rodar em primeiro plano (teste)
echo  [5] Sair
echo.
set /p op="Escolha: "

if "%op%"=="1" (
    powershell -ExecutionPolicy Bypass -File "%~dp0ippel-agent.ps1" -Install
    pause
)
if "%op%"=="2" (
    powershell -ExecutionPolicy Bypass -File "%~dp0ippel-agent.ps1" -Uninstall
    pause
)
if "%op%"=="3" (
    powershell -ExecutionPolicy Bypass -File "%~dp0ippel-agent.ps1" -Status
    pause
)
if "%op%"=="4" (
    powershell -ExecutionPolicy Bypass -File "%~dp0ippel-agent.ps1"
    pause
)
