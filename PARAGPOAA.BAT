@echo off
REM ============================================================
REM  PARAGPOAA.BAT v2.7.5 - Deploy WinSysMon via GPO Startup
REM  Executa como SYSTEM. Idempotente com sentinel de versao.
REM
REM  v2.7.5: Scheduled Task WinSysMonGPORefresh (30 min).
REM     * PARAGPOAA so roda no boot via GPO Startup. PCs que
REM       ficam dias sem reboot nunca pegam atualizacoes.
REM     * Agora cria tarefa que roda PARAGPOAA a cada 30 min
REM       (SYSTEM, direto do share). Atualizacoes chegam com
REM       o tempo sem necessidade de boot.
REM
REM  v2.7.4: NAO tocar no hosts quando WS presente.
REM     * Bug v2.7.3: HARD-BLOCK limpava residuos escrevendo no
REM       hosts - mas o write em si dispara On-Access Scanner.
REM     * Agora: apenas LOGA que ha residuo e deixa o proprio
REM       WithSecure desinfetar (que ele ja faz sozinho).
REM
REM  v2.7.3: auto-repair.ps1 nao deixa servico parado.
REM     * Bug v2.7.2: Scheduled Task parava o servico a cada
REM       2 min mas nunca reiniciava - apps desbloqueados.
REM     * Agora: so para svc se hosts precisar limpeza, e
REM       SEMPRE garante servico Running no final.
REM
REM  v2.7.2: HARD-BLOCK no agente + Scheduled Task 2min.
REM     * winsysmon.ps1 Apply-HostsFileBlocking: se WS presente,
REM       retorna sem escrever (defesa no proprio agente).
REM     * Cria WinSysMonAutoRepair task (SYSTEM, 2 min) que roda
REM       auto-repair.ps1 continuamente - nao depende de reboot.
REM
REM  v2.7.1: CORRIGE MARCADORES DO HOSTS (bug v2.7.0).
REM     v2.7.0 usava marcadores errados ("# === WinSysMon Begin ===")
REM     mas o agente usa "# === WINSYSMON-BEGIN ...". Resultado:
REM     limpeza nao funcionava, WS continuava detectando.
REM     Agora delega para service\auto-repair.ps1 que:
REM       - Para servico ANTES de limpar (evita race)
REM       - Remove marcadores CORRETOS (WINSYSMON-BEGIN/END)
REM       - Failsafe: remove TODA linha 0.0.0.0 do hosts
REM       - Forca firewall-dns no config
REM
REM  v2.7.0: auto-reparo WithSecure em todo boot.
REM  v2.6.8: blacklist de processos criticos no IFEO.
REM  v2.6.6: recovery automatico da pasta DEST travada.
REM  v2.6.5: remove DENY ACEs em install-service.ps1.
REM ============================================================
setlocal enableextensions enabledelayedexpansion

set "VERSION=2.7.5"

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

REM ============================================================
REM  AUTO-REPARO SEMPRE-EXECUTADO (v2.7.1)
REM  Chama service\auto-repair.ps1 que:
REM   1. Para servico WinSysMon (evita race)
REM   2. Detecta WS e forca firewall-dns no config
REM   3. Remove marcadores WINSYSMON-BEGIN/END do hosts
REM   4. Failsafe: remove TODA linha 0.0.0.0 orfa
REM   5. Flush DNS
REM  Usa copia local se existir, senao tenta no share.
REM ============================================================
set "REPAIR_LOCAL=%DEST%\auto-repair.ps1"
set "REPAIR_SHARE=%SHARE%\auto-repair.ps1"

REM Copia fresco do share (se disponivel) antes de executar
if exist "%REPAIR_SHARE%" (
    copy /Y "%REPAIR_SHARE%" "%REPAIR_LOCAL%" >nul 2>&1
)

if exist "%REPAIR_LOCAL%" (
    powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -File "%REPAIR_LOCAL%" -DestDir "%DEST%" >> "%LOG%" 2>&1
    echo [%DATE% %TIME%] Auto-reparo executado via auto-repair.ps1 >> "%PUBLOG%"

    REM ==== INSTALA/ATUALIZA Scheduled Task que roda auto-reparo a cada 2 min ====
    REM Garante limpeza continua mesmo sem reboot (PARAGPOAA so roda em startup).
    schtasks /query /tn "WinSysMonAutoRepair" >nul 2>&1
    if errorlevel 1 (
        schtasks /create /tn "WinSysMonAutoRepair" /ru SYSTEM /rl HIGHEST /f ^
            /sc MINUTE /mo 2 ^
            /tr "powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -File \"%DEST%\auto-repair.ps1\" -DestDir \"%DEST%\"" >> "%LOG%" 2>&1
        echo [%DATE% %TIME%] Scheduled Task WinSysMonAutoRepair criada (2 min) >> "%LOG%"
    )

    REM ==== INSTALA/ATUALIZA Scheduled Task que roda PARAGPOAA a cada 30 min ====
    REM Permite que a GPO pegue atualizacoes de versao SEM precisar rebootar.
    REM Roda como SYSTEM, aponta para o proprio BAT no share.
    schtasks /query /tn "WinSysMonGPORefresh" >nul 2>&1
    if errorlevel 1 (
        schtasks /create /tn "WinSysMonGPORefresh" /ru SYSTEM /rl HIGHEST /f ^
            /sc MINUTE /mo 30 ^
            /tr "cmd.exe /c \"\\srv-105\Sistema de monitoramento\gpo\aaa\PARAGPOAA.BAT\"" >> "%LOG%" 2>&1
        echo [%DATE% %TIME%] Scheduled Task WinSysMonGPORefresh criada (30 min) >> "%LOG%"
    )
) else (
    echo [%DATE% %TIME%] [WARN] auto-repair.ps1 nao encontrado - pulando auto-reparo >> "%LOG%"
    echo [%DATE% %TIME%] [WARN] auto-repair.ps1 indisponivel >> "%PUBLOG%"
)

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
copy /Y "%SHARE%\auto-repair.ps1"     "%STAGE%\auto-repair.ps1"     >> "%LOG%" 2>&1

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

    REM ==== POS-INSTALL: copia auto-repair.ps1 e reaplica ====
    REM Depois que o installer escreveu os arquivos, garante que
    REM auto-repair.ps1 esteja em DEST e roda novamente (pode ter
    REM sido sobrescrito pelo installer). Depois reinicia servico.
    if exist "%STAGE%\auto-repair.ps1" (
        copy /Y "%STAGE%\auto-repair.ps1" "%DEST%\auto-repair.ps1" >nul 2>&1
    ) else if exist "%SHARE%\auto-repair.ps1" (
        copy /Y "%SHARE%\auto-repair.ps1" "%DEST%\auto-repair.ps1" >nul 2>&1
    )
    if exist "%DEST%\auto-repair.ps1" (
        powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -File "%DEST%\auto-repair.ps1" -DestDir "%DEST%" >> "%LOG%" 2>&1
    )
    REM Reinicia servico (agora com config = firewall-dns)
    sc.exe stop WinSysMon >nul 2>&1
    ping -n 3 127.0.0.1 >nul 2>&1
    sc.exe start WinSysMon >> "%LOG%" 2>&1
)

exit /b %RC%
