@echo off
REM ============================================================
REM  PARAGPOAA.BAT - Deploy WinSysMon via GPO (Computer Startup)
REM  Executa como SYSTEM. Idempotente com versionamento.
REM ============================================================
setlocal enableextensions

set "SHARE=\\srv-105\Sistema de monitoramento\gpo\aaa\service"
set "INSTALLER=%SHARE%\install-service.ps1"
set "SRC_SCRIPT=%SHARE%\winsysmon.ps1"
set "JSON=%SHARE%\blocked-apps.json"
set "HOSTSJSON=%SHARE%\blocked-hosts.json"
set "POLJSON=%SHARE%\blocked-policies.json"
set "DEST=%ProgramData%\Microsoft\WinSysMon"
set "DEST_SCRIPT=%DEST%\winsysmon.ps1"
set "LOG=%DEST%\install.log"
set "FLAG=%DEST%\installed.flag"

if not exist "%DEST%" mkdir "%DEST%" 2>nul

echo [%DATE% %TIME%] === GPO startup em %COMPUTERNAME% === >> "%LOG%"

REM ---- Aguardar share ficar acessivel (rede pode demorar) ----
set "TRIES=0"
:WAIT_SHARE
if exist "%INSTALLER%" goto SHARE_OK
set /a TRIES+=1
if %TRIES% GEQ 12 (
    echo [%DATE% %TIME%] [ERRO] Share inacessivel apos 60s: %SHARE% >> "%LOG%"
    exit /b 1
)
REM Dorme 5s via ping (nao depende de Start-Sleep)
ping -n 6 127.0.0.1 >nul 2>&1
goto WAIT_SHARE

:SHARE_OK
echo [%DATE% %TIME%] Share acessivel (tentativas=%TRIES%) >> "%LOG%"

REM ---- Comparar versao do script no share vs instalado ----
set "NEEDS_UPDATE=0"

REM Se servico nao existe, precisa instalar
sc.exe query WinSysMon >nul 2>&1
if errorlevel 1 (
    set "NEEDS_UPDATE=1"
    echo [%DATE% %TIME%] Servico nao instalado >> "%LOG%"
    goto DO_INSTALL
)

REM Se script local nao existe, precisa instalar
if not exist "%DEST_SCRIPT%" (
    set "NEEDS_UPDATE=1"
    echo [%DATE% %TIME%] Script local ausente >> "%LOG%"
    goto DO_INSTALL
)

REM Comparar por tamanho + data (fcompare via xcopy /D /L)
xcopy "%SRC_SCRIPT%" "%DEST_SCRIPT%" /D /Y /L 2>nul | findstr /C:"File(s)" | findstr /C:"0 " >nul
if errorlevel 1 (
    set "NEEDS_UPDATE=1"
    echo [%DATE% %TIME%] Script no share e mais novo - atualizando >> "%LOG%"
    goto DO_INSTALL
)

REM Servico existe e script esta up-to-date
sc.exe query WinSysMon 2>nul | findstr /I "RUNNING" >nul
if %ERRORLEVEL%==0 (
    echo [%DATE% %TIME%] Servico rodando e atualizado. Nada a fazer. >> "%LOG%"
    exit /b 0
)

REM Servico parado mas script ok: tentar iniciar
echo [%DATE% %TIME%] Servico parado - tentando iniciar... >> "%LOG%"
sc.exe start WinSysMon >> "%LOG%" 2>&1
ping -n 3 127.0.0.1 >nul 2>&1
sc.exe query WinSysMon 2>nul | findstr /I "RUNNING" >nul
if %ERRORLEVEL%==0 (
    echo [%DATE% %TIME%] Servico reiniciado com sucesso >> "%LOG%"
    exit /b 0
)
REM Falhou ao iniciar - reinstalar
set "NEEDS_UPDATE=1"

:DO_INSTALL
echo [%DATE% %TIME%] Rodando installer... >> "%LOG%"
powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -File "%INSTALLER%" -SharePath "%JSON%" -HostsSharePath "%HOSTSJSON%" -PoliciesSharePath "%POLJSON%" >> "%LOG%" 2>&1
set "RC=%ERRORLEVEL%"

if "%RC%"=="0" (
    echo Instalado em %DATE% %TIME% > "%FLAG%"
    echo [%DATE% %TIME%] Install OK >> "%LOG%"
) else (
    echo [%DATE% %TIME%] Install retornou %RC% >> "%LOG%"
)

exit /b %RC%
