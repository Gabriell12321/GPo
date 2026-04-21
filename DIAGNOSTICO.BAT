@echo off
REM ============================================================
REM  DIAGNOSTICO.BAT - Roda o diag-winsysmon.ps1 como admin
REM ============================================================
title WinSysMon - Diagnostico

net session >nul 2>&1
if %errorLevel% NEQ 0 (
    echo Solicitando privilegios de Administrador...
    powershell -NoProfile -Command "Start-Process -FilePath '%~f0' -Verb RunAs"
    exit /b
)

set "SHARE=\\srv-105\aaa$\service"
set "SHARE_LEGACY=\\srv-105\Sistema de monitoramento\gpo\aaa\service"
if not exist "%SHARE%\diag-winsysmon.ps1" if exist "%SHARE_LEGACY%\diag-winsysmon.ps1" set "SHARE=%SHARE_LEGACY%"
set "DIAG=%SHARE%\diag-winsysmon.ps1"

if not exist "%DIAG%" (
    echo ERRO: diag-winsysmon.ps1 nao encontrado em nenhum share.
    pause
    exit /b 1
)

echo Rodando diagnostico...
echo Saida sera gravada em %TEMP%\winsysmon-diag.log
echo.
powershell -NoProfile -ExecutionPolicy Bypass -Command "& '%DIAG%' *>&1 | Tee-Object -FilePath '%TEMP%\winsysmon-diag.log'"

echo.
echo ============================================================
echo  Log salvo em: %TEMP%\winsysmon-diag.log
echo ============================================================
pause
