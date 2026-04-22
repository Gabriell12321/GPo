@echo off
REM ============================================================
REM  PARAGPOAA.BAT v2.6.5 - Deploy WinSysMon via GPO Startup
REM  Executa como SYSTEM. Idempotente com sentinel de versao.
REM
REM  v2.6.4: substitui robocopy por copy direto (mais robusto
REM  com shares remotos + SYSTEM). Stage local elimina bug do
REM  powershell.exe -File com UNC contendo espacos.
REM ============================================================
setlocal enableextensions enabledelayedexpansion

set "VERSION=2.6.5"

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
@echo off
REM ============================================================
REM  PARAGPOAA.BAT v2.6.3 - Deploy WinSysMon via GPO Startup
REM  Executa como SYSTEM. Idempotente com sentinel de versao.
REM
REM  MUDANCA v2.6.3 (CRITICA): copia installer+JSONs para pasta
REM  LOCAL (sem espacos) antes de executar. Isso elimina bug do
REM  powershell.exe -File "\\srv-105\pasta com espaco\..." que
REM  falhava silenciosamente em TODOS os PCs.
REM ============================================================
setlocal enableextensions enabledelayedexpansion

set "VERSION=2.6.3"

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

if not exist "%DEST%" mkdir "%DEST%" 2>nul
if not exist "%STAGE%" mkdir "%STAGE%" 2>nul

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

REM ==== STAGE: copia TUDO do share para pasta local SEM espacos ====
echo [%DATE% %TIME%] Copiando share -^> %STAGE% >> "%LOG%"
robocopy "%SHARE%" "%STAGE%" /MIR /R:2 /W:2 /NP /NFL /NDL /NJH /NJS >> "%LOG%" 2>&1
set "RCR=%ERRORLEVEL%"
if %RCR% GEQ 8 (
    echo [%DATE% %TIME%] [ERRO] robocopy falhou rc=%RCR% >> "%LOG%"
    echo [%DATE% %TIME%] [ERRO] robocopy falhou rc=%RCR% >> "%PUBLOG%"
    exit /b 2
)

if not exist "%INSTALLER_LOCAL%" (
    echo [%DATE% %TIME%] [ERRO] installer nao chegou ao stage: %INSTALLER_LOCAL% >> "%LOG%"
    echo [%DATE% %TIME%] [ERRO] stage vazio >> "%PUBLOG%"
    exit /b 3
)

echo [%DATE% %TIME%] Executando installer de %INSTALLER_LOCAL% >> "%LOG%"
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
@echo off
REM ============================================================
REM  PARAGPOAA.BAT v2.6.2 - Deploy WinSysMon via GPO Startup
REM  Executa como SYSTEM. Idempotente com sentinel de versao.
REM
REM  LOGICA v2.6.2:
REM   1) Resolve share (aaa$ preferido, legacy fallback)
REM   2) Le version.txt do share
REM   3) Compara com version.txt local
REM   4) Se diferente OU servico nao existe OU IFEO refresh task ausente
REM      OU servico nao inicia -> reinstala (remove antes se travado)
REM ============================================================
setlocal enableextensions enabledelayedexpansion

set "VERSION=2.6.2"

set "SHARE=\\srv-105\aaa$\service"
set "SHARE_LEGACY=\\srv-105\Sistema de monitoramento\gpo\aaa\service"
if not exist "%SHARE%\install-service.ps1" if exist "%SHARE_LEGACY%\install-service.ps1" set "SHARE=%SHARE_LEGACY%"

set "INSTALLER=%SHARE%\install-service.ps1"
set "SHARE_VER=%SHARE%\version.txt"
set "JSON=%SHARE%\blocked-apps.json"
set "HOSTSJSON=%SHARE%\blocked-hosts.json"
set "POLJSON=%SHARE%\blocked-policies.json"

set "DEST=%ProgramData%\Microsoft\WinSysMon"
set "DEST_SCRIPT=%DEST%\winsysmon.ps1"
set "LOCAL_VER=%DEST%\version.txt"
set "LOG=%DEST%\install.log"
set "FLAG=%DEST%\installed.flag"

if not exist "%DEST%" mkdir "%DEST%" 2>nul

echo [%DATE% %TIME%] === GPO startup em %COMPUTERNAME% (PARAGPOAA v%VERSION%) === >> "%LOG%"

set "TRIES=0"
:WAIT_SHARE
if exist "%INSTALLER%" goto SHARE_OK
set /a TRIES+=1
if %TRIES% GEQ 12 (
    echo [%DATE% %TIME%] [ERRO] Share inacessivel apos 60s: %SHARE% >> "%LOG%"
    exit /b 1
)
ping -n 6 127.0.0.1 >nul 2>&1
goto WAIT_SHARE

:SHARE_OK
echo [%DATE% %TIME%] Share acessivel em %SHARE% (tentativas=%TRIES%) >> "%LOG%"

set "NEEDS_UPDATE=0"
set "REASON="

REM ---- CHECK 1: versao diferente do share ----
set "SHARE_V="
if exist "%SHARE_VER%" (
    for /f "usebackq delims=" %%v in ("%SHARE_VER%") do set "SHARE_V=%%v"
)
set "LOCAL_V="
if exist "%LOCAL_VER%" (
    for /f "usebackq delims=" %%v in ("%LOCAL_VER%") do set "LOCAL_V=%%v"
)
echo [%DATE% %TIME%] Versao share=[!SHARE_V!] local=[!LOCAL_V!] >> "%LOG%"
if defined SHARE_V (
    if not "!SHARE_V!"=="!LOCAL_V!" (
        set "NEEDS_UPDATE=1"
        set "REASON=versao diferente ^(share=!SHARE_V! local=!LOCAL_V!^)"
        goto DO_INSTALL
    )
)

REM ---- CHECK 2: servico nao existe ----
sc.exe query WinSysMon >nul 2>&1
if errorlevel 1 (
    set "NEEDS_UPDATE=1"
    set "REASON=servico nao existe"
    goto DO_INSTALL
)

REM ---- CHECK 3: script local ausente ----
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
set "REASON=servico nao inicia (reinstalacao forcada)"

echo [%DATE% %TIME%] Removendo servico travado antes de reinstalar... >> "%LOG%"
sc.exe stop WinSysMon >>"%LOG%" 2>&1
ping -n 3 127.0.0.1 >nul 2>&1
sc.exe delete WinSysMon >>"%LOG%" 2>&1
ping -n 3 127.0.0.1 >nul 2>&1

:DO_INSTALL
echo [%DATE% %TIME%] REINSTALANDO motivo=[!REASON!] >> "%LOG%"
powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -File "%INSTALLER%" -SharePath "%JSON%" -HostsSharePath "%HOSTSJSON%" -PoliciesSharePath "%POLJSON%" >> "%LOG%" 2>&1
set "RC=%ERRORLEVEL%"

if "%RC%"=="0" (
    echo Instalado em %DATE% %TIME% > "%FLAG%"
    if defined SHARE_V (
        echo !SHARE_V!> "%LOCAL_VER%"
    ) else (
        echo %VERSION%> "%LOCAL_VER%"
    )
    echo [%DATE% %TIME%] Install OK (versao !SHARE_V!) >> "%LOG%"
) else (
    echo [%DATE% %TIME%] Install retornou %RC% >> "%LOG%"
)

exit /b %RC%
@echo off
REM ============================================================
REM  PARAGPOAA.BAT - Deploy WinSysMon via GPO (Computer Startup)
REM  Executa como SYSTEM. Idempotente com versionamento.
REM ============================================================
setlocal enableextensions

set "SHARE=\\srv-105\aaa$\service"
set "SHARE_LEGACY=\\srv-105\Sistema de monitoramento\gpo\aaa\service"
if not exist "%SHARE%\install-service.ps1" if exist "%SHARE_LEGACY%\install-service.ps1" set "SHARE=%SHARE_LEGACY%"
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
