@echo off
REM ============================================================
REM  DEPLOY-WINSYSMON.BAT - Setup completo em qualquer PC
REM  1. Auto-eleva via UAC
REM  2. Se WinSysMon nao instalado: chama INSTALAR.BAT
REM  3. Se instalado mas servico parado: roda fix-acl
REM  4. Se rodando: confirma status e sai
REM ============================================================
setlocal EnableDelayedExpansion

REM Auto-elevacao UAC
net session >nul 2>&1
if %errorlevel% neq 0 (
    echo Solicitando elevacao UAC...
    powershell -NoProfile -Command "Start-Process -FilePath '%~f0' -Verb RunAs"
    exit /b
)

title DEPLOY WinSysMon - %COMPUTERNAME%

echo.
echo  ====================================================
echo   DEPLOY WinSysMon - %COMPUTERNAME%
echo  ====================================================
echo.

set "SHARE=\\srv-105\Sistema de monitoramento\gpo\aaa"
set "INSTALL_DIR=C:\ProgramData\Microsoft\WinSysMon"

REM ---- 1. Verifica se esta instalado ----
sc query WinSysMon >nul 2>&1
if %errorlevel% equ 0 (
    echo  [1/3] Servico WinSysMon EXISTE no SCM
    set "INSTALLED=1"
) else (
    echo  [1/3] Servico WinSysMon NAO existe - precisa instalar
    set "INSTALLED=0"
)

REM ---- 2. Se nao instalado, roda INSTALAR.BAT do share ----
if "!INSTALLED!"=="0" (
    echo.
    echo  [2/3] Instalando WinSysMon do share...
    if not exist "%SHARE%\INSTALAR.BAT" (
        echo  [ERRO] INSTALAR.BAT nao encontrado em %SHARE%
        pause
        exit /b 1
    )
    call "%SHARE%\INSTALAR.BAT"
    echo.
    echo  Aguardando servico estabilizar...
    timeout /t 5 /nobreak >nul
)

REM ---- 3. Verifica status atual ----
echo.
echo  [3/3] Verificando status final...
for /f "tokens=3 delims=: " %%s in ('sc query WinSysMon ^| findstr "STATE"') do set "STATE=%%s"

if "!STATE!"=="RUNNING" (
    echo  [OK] Servico RUNNING - WinSysMon operacional
    goto :END
)

REM Nao esta rodando - tenta fix-acl
echo  [WARN] Servico nao esta RUNNING (estado: !STATE!)
echo  Aplicando fix-acl...
echo.

set "FIX_SCRIPT=%SHARE%\service\fix-acl-pc200.ps1"
if not exist "%FIX_SCRIPT%" (
    echo  [ERRO] fix-acl-pc200.ps1 nao encontrado em %FIX_SCRIPT%
    pause
    exit /b 1
)

powershell -NoProfile -ExecutionPolicy Bypass -File "%FIX_SCRIPT%"

REM Recheck
for /f "tokens=3 delims=: " %%s in ('sc query WinSysMon ^| findstr "STATE"') do set "STATE=%%s"
if "!STATE!"=="RUNNING" (
    echo.
    echo  [OK] Fix aplicado - servico agora RUNNING
) else (
    echo.
    echo  [FAIL] Servico ainda nao subiu. Verifique manualmente:
    echo    - AV quarentena ^(WithSecure/Defender^)
    echo    - Log: %INSTALL_DIR%\sysmon.log
    echo    - Event Viewer: System / WinSysMon
)

:END
echo.
echo  ====================================================
echo   TESTE: tente abrir a calculadora
echo   Se nao abrir = WinSysMon funcionando
echo  ====================================================
echo.
pause
endlocal
