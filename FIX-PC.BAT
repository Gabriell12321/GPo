@echo off
REM ============================================================
REM  FIX-PC.BAT - Corrige WinSysMon com Access Denied
REM  Duplo-clique: auto-eleva via UAC e executa fix-acl
REM ============================================================
setlocal

REM Checa se ja esta elevado
net session >nul 2>&1
if %errorlevel% neq 0 (
    echo.
    echo  ====================================================
    echo   FIX-PC - Corretor de ACL do WinSysMon
    echo  ====================================================
    echo.
    echo  Solicitando elevacao UAC...
    echo.
    REM Auto-eleva chamando a si mesmo via PowerShell Start-Process -Verb RunAs
    powershell -NoProfile -Command "Start-Process -FilePath '%~f0' -Verb RunAs"
    exit /b
)

REM Aqui ja esta elevado
title FIX-PC - WinSysMon ACL Repair

echo.
echo  ====================================================
echo   FIX-PC - Rodando com privilegios de Administrador
echo  ====================================================
echo.
echo  Computador: %COMPUTERNAME%
echo  Usuario:    %USERDOMAIN%\%USERNAME%
echo.

REM Caminho do script no share
set "SHARE_SCRIPT=\\srv-105\Sistema de monitoramento\gpo\aaa\service\fix-acl-pc200.ps1"
set "LOCAL_SCRIPT=C:\ProgramData\Microsoft\WinSysMon\fix-acl-pc200.ps1"

REM Tenta o share primeiro, fallback local
if exist "%SHARE_SCRIPT%" (
    echo  Usando script do share: %SHARE_SCRIPT%
    set "SCRIPT=%SHARE_SCRIPT%"
) else if exist "%LOCAL_SCRIPT%" (
    echo  Share inacessivel - usando copia local: %LOCAL_SCRIPT%
    set "SCRIPT=%LOCAL_SCRIPT%"
) else (
    echo.
    echo  [ERRO] Script fix-acl-pc200.ps1 nao encontrado nem no share nem local.
    echo         Share:  %SHARE_SCRIPT%
    echo         Local:  %LOCAL_SCRIPT%
    echo.
    pause
    exit /b 1
)

echo.
echo  ====================================================
echo.

powershell -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT%"

echo.
echo  ====================================================
echo   Fim da execucao
echo  ====================================================
pause
endlocal
