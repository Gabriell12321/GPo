@echo off
REM ============================================================
REM  INSTALAR.BAT - Instala WinSysMon como SERVICO WINDOWS REAL
REM  Aparece em services.msc (servicos.msc)
REM  Duplo-clique - auto-eleva para admin
REM ============================================================
title Instalacao WinSysMon - Servico Windows

net session >nul 2>&1
if %errorLevel% NEQ 0 (
    echo Solicitando privilegios de Administrador...
    powershell -NoProfile -Command "Start-Process -FilePath '%~f0' -Verb RunAs"
    exit /b
)

set "SHARE=\\srv-105\Sistema de monitoramento\gpo\aaa\service"
set "INSTALLER=%SHARE%\install-service.ps1"
set "JSON=%SHARE%\blocked-apps.json"

echo.
echo ==============================================
echo   WinSysMon - Instalador (Servico Windows)
echo ==============================================
echo.
echo  Maquina: %COMPUTERNAME%
echo  Share:   %SHARE%
echo.

if not exist "%INSTALLER%" (
    echo [ERRO] Nao consigo ler "%INSTALLER%"
    echo        Verifique permissoes do share/NTFS para Domain Computers
    echo.
    pause
    exit /b 1
)

powershell -NoProfile -ExecutionPolicy Bypass -File "%INSTALLER%" -SharePath "%JSON%"

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [ERRO] Instalacao falhou
    pause
    exit /b 1
)

echo.
echo ==============================================
echo   Abra services.msc e procure "Windows System Monitor"
echo ==============================================
echo.
pause
exit /b 0
