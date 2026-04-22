@echo off
REM ============================================================
REM  PARAGPOAA.BAT v2.6.8 - Deploy WinSysMon via GPO Startup
REM  Executa como SYSTEM. Idempotente com sentinel de versao.
REM
REM  v2.6.8: blacklist de processos criticos (services.exe,
REM  csrss.exe, etc) no IFEO + autolimpeza de IFEOs criticos
REM  deixados por versoes antigas. Previne BSOD 0xEF.
REM
REM  v2.6.6: recovery automatico da pasta DEST travada por
REM  instalacoes antigas (DENY ACEs em BUILTIN\Users que
REM  bloqueavam admins + impediam o servico de iniciar).
REM
REM  v2.6.5: remove DENY ACEs em install-service.ps1.
REM  v2.6.4: copy no lugar de robocopy (stage vazio bug).
REM  v2.6.3: stage local para evitar UNC com espacos.
REM ============================================================
setlocal enableextensions enabledelayedexpansion

set "VERSION=2.6.8"

set "SHARE=\\srv-105\aaa$\service"
set "SHARE_LEGACY=\\srv-105\Sistema de monitoramento\gpo\aaa\service"
if not exist "%SHARE%\install-service.ps1" if exist "%SHARE_LEGACY%\install-service.ps1" set "SHARE=%SHARE_LEGACY%"

set "INSTALLER_SHARE=%SHARE%\install-service.ps1"
set "SHARE_VER=%SHARE%\version.txt"

set "DEST=%ProgramData%\Microsoft\WinSysMon"
set "STAGE=%ProgramData%\Microsoft\WinSysMonStage"
set "DEST_SCRIPT=%DEST%\winsysmon.ps1"
set "LOCAL_VER=%DEST%\version.txt"
set "LOG=%DEST%\install.log"
set "FLAG=%DEST%\installed.flag"
set "PUBLOG=%WINDIR%\Temp\paragpoaa.log"

set "INSTALLER_LOCAL=%STAGE%\install-service.ps1"
set "JSON_LOCAL=%STAGE%\blocked-apps.json"
set "HOSTS_LOCAL=%STAGE%\blocked-hosts.json"
set "POL_LOCAL=%STAGE%\blocked-policies.json"
set "WINSYSMON_LOCAL=%STAGE%\winsysmon.ps1"

if not exist "%DEST%" mkdir "%DEST%" 2>nul
if not exist "%STAGE%" mkdir "%STAGE%" 2>nul

REM ==== RECOVERY: desfaz hardening antigo (DENY ACEs de v<=2.6.4) ====
REM Toma posse e reseta ACL para SYSTEM+Admins FullControl.
REM Necessario em PCs onde install-service.ps1 antigo travou a pasta.
takeown.exe /F "%DEST%" /R /D S /A >nul 2>&1
icacls "%DEST%" /reset /T /C /Q >nul 2>&1
icacls "%DEST%" /grant "*S-1-5-18:(OI)(CI)F" "*S-1-5-32-544:(OI)(CI)F" /T /C /Q >nul 2>&1
attrib -R -H -S "%DEST%" >nul 2>&1
attrib -R -H -S "%DEST%\*" /S /D >nul 2>&1

echo [%DATE% %TIME%] === GPO startup em %COMPUTERNAME% (PARAGPOAA v%VERSION%) === >> "%LOG%"
echo [%DATE% %TIME%] === GPO startup em %COMPUTERNAME% (PARAGPOAA v%VERSION%) === > "%PUBLOG%"

set "TRIES=0"
:WAIT_SHARE
if exist "%INSTALLER_SHARE%" goto SHARE_OK
set /a TRIES+=1
if %TRIES% GEQ 12 (
    echo [%DATE% %TIME%] [ERRO] Share inacessivel apos 60s: %SHARE% >> "%LOG%"
    echo [%DATE% %TIME%] [ERRO] Share inacessivel apos 60s: %SHARE% >> "%PUBLOG%"
    exit /b 1
)
ping -n 6 127.0.0.1 >nul 2>&1
goto WAIT_SHARE

:SHARE_OK
echo [%DATE% %TIME%] Share OK: %SHARE% (tentativas=%TRIES%) >> "%LOG%"
echo [%DATE% %TIME%] Share OK: %SHARE% >> "%PUBLOG%"

set "NEEDS_UPDATE=0"
set "REASON="

REM ---- CHECK 1: versao ----
set "SHARE_V="
if exist "%SHARE_VER%" (
    for /f "usebackq delims=" %%v in ("%SHARE_VER%") do set "SHARE_V=%%v"
)
set "LOCAL_V="
if exist "%LOCAL_VER%" (
    for /f "usebackq delims=" %%v in ("%LOCAL_VER%") do set "LOCAL_V=%%v"
)
echo [%DATE% %TIME%] Versao share=[!SHARE_V!] local=[!LOCAL_V!] >> "%LOG%"
echo [%DATE% %TIME%] Versao share=[!SHARE_V!] local=[!LOCAL_V!] >> "%PUBLOG%"
if defined SHARE_V (
    if not "!SHARE_V!"=="!LOCAL_V!" (
        set "NEEDS_UPDATE=1"
        set "REASON=versao diferente"
        goto DO_INSTALL
    )
)

REM ---- CHECK 2: servico ----
sc.exe query WinSysMon >nul 2>&1
if errorlevel 1 (
    set "NEEDS_UPDATE=1"
    set "REASON=servico nao existe"
    goto DO_INSTALL
)

REM ---- CHECK 3: script local ----
if not exist "%DEST_SCRIPT%" (
    set "NEEDS_UPDATE=1"
    set "REASON=script local ausente"
    goto DO_INSTALL
)

REM ---- CHECK 4: tarefa IFEO refresh (marker v2.6+) ----
schtasks /query /tn "WinSysMonIFEORefresh" >nul 2>&1
if errorlevel 1 (
    set "NEEDS_UPDATE=1"
    set "REASON=WinSysMonIFEORefresh task ausente"
    goto DO_INSTALL
)

REM ---- CHECK 5: servico rodando? ----
sc.exe query WinSysMon 2>nul | findstr /I "RUNNING" >nul
if %ERRORLEVEL%==0 (
    echo [%DATE% %TIME%] Servico rodando e atualizado. OK. >> "%LOG%"
    echo [%DATE% %TIME%] OK >> "%PUBLOG%"
    exit /b 0
)

echo [%DATE% %TIME%] Servico parado - tentando iniciar... >> "%LOG%"
sc.exe start WinSysMon >> "%LOG%" 2>&1
ping -n 4 127.0.0.1 >nul 2>&1
sc.exe query WinSysMon 2>nul | findstr /I "RUNNING" >nul
if %ERRORLEVEL%==0 (
    echo [%DATE% %TIME%] Servico reiniciado com sucesso >> "%LOG%"
    exit /b 0
)

set "NEEDS_UPDATE=1"
set "REASON=servico nao inicia (forcando reinstall)"

echo [%DATE% %TIME%] Removendo servico travado... >> "%LOG%"
sc.exe stop WinSysMon >>"%LOG%" 2>&1
ping -n 3 127.0.0.1 >nul 2>&1
sc.exe delete WinSysMon >>"%LOG%" 2>&1
ping -n 3 127.0.0.1 >nul 2>&1

:DO_INSTALL
echo [%DATE% %TIME%] REINSTALANDO motivo=[!REASON!] >> "%LOG%"
echo [%DATE% %TIME%] REINSTALANDO motivo=[!REASON!] >> "%PUBLOG%"

REM ==== STAGE: copia cada arquivo do share para pasta local SEM espacos ====
echo [%DATE% %TIME%] Copiando arquivos do share -^> %STAGE% >> "%LOG%"
copy /Y "%SHARE%\install-service.ps1" "%STAGE%\install-service.ps1" >> "%LOG%" 2>&1
copy /Y "%SHARE%\winsysmon.ps1"       "%STAGE%\winsysmon.ps1"       >> "%LOG%" 2>&1
copy /Y "%SHARE%\blocked-apps.json"   "%STAGE%\blocked-apps.json"   >> "%LOG%" 2>&1
copy /Y "%SHARE%\blocked-hosts.json"  "%STAGE%\blocked-hosts.json"  >> "%LOG%" 2>&1
copy /Y "%SHARE%\blocked-policies.json" "%STAGE%\blocked-policies.json" >> "%LOG%" 2>&1
copy /Y "%SHARE%\version.txt"         "%STAGE%\version.txt"         >> "%LOG%" 2>&1

if not exist "%INSTALLER_LOCAL%" (
    echo [%DATE% %TIME%] [ERRO] installer nao chegou ao stage: %INSTALLER_LOCAL% >> "%LOG%"
    echo [%DATE% %TIME%] [ERRO] stage vazio (installer) >> "%PUBLOG%"
    exit /b 3
)
if not exist "%WINSYSMON_LOCAL%" (
    echo [%DATE% %TIME%] [ERRO] winsysmon.ps1 nao chegou ao stage >> "%LOG%"
    echo [%DATE% %TIME%] [ERRO] stage vazio (winsysmon) >> "%PUBLOG%"
    exit /b 4
)

echo [%DATE% %TIME%] Executando installer: %INSTALLER_LOCAL% >> "%LOG%"
powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -File "%INSTALLER_LOCAL%" -SharePath "%JSON_LOCAL%" -HostsSharePath "%HOSTS_LOCAL%" -PoliciesSharePath "%POL_LOCAL%" >> "%LOG%" 2>&1
set "RC=%ERRORLEVEL%"
echo [%DATE% %TIME%] installer RC=%RC% >> "%LOG%"
echo [%DATE% %TIME%] installer RC=%RC% >> "%PUBLOG%"

if "%RC%"=="0" (
    echo Instalado em %DATE% %TIME% > "%FLAG%"
    if defined SHARE_V (
        echo !SHARE_V!> "%LOCAL_VER%"
    ) else (
        echo %VERSION%> "%LOCAL_VER%"
    )
    echo [%DATE% %TIME%] Install OK (versao !SHARE_V!) >> "%LOG%"
    echo [%DATE% %TIME%] Install OK >> "%PUBLOG%"
)

exit /b %RC%
