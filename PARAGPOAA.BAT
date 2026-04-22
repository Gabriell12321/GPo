@echo off
REM ============================================================
REM  PARAGPOAA.BAT v2.7.0 - Deploy WinSysMon via GPO Startup
REM  Executa como SYSTEM. Idempotente com sentinel de versao.
REM
REM  v2.7.0: AUTO-REPARO SEMPRE-EXECUTADO
REM     * Detecta WithSecure/F-Secure (servicos + pastas)
REM     * Se WS ativo: forca HostBlockingMethod=firewall-dns em
REM       sysmon-config.json local ANTES do servico iniciar
REM     * Limpa marcadores WinSysMon no hosts em toda execucao
REM       (evita "Redirected hosts file" mesmo sem reinstall)
REM     * Repara ACL da pasta DEST se SYSTEM nao tiver acesso
REM     * Roda em TODO boot (nao so em reinstall)
REM
REM  v2.6.8: blacklist de processos criticos no IFEO + autolimpeza.
REM  v2.6.6: recovery automatico da pasta DEST travada.
REM  v2.6.5: remove DENY ACEs em install-service.ps1.
REM  v2.6.4: copy no lugar de robocopy (stage vazio bug).
REM  v2.6.3: stage local para evitar UNC com espacos.
REM ============================================================
setlocal enableextensions enabledelayedexpansion

set "VERSION=2.7.0"

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
REM  AUTO-REPARO SEMPRE-EXECUTADO (v2.7.0)
REM  Roda em TODO boot, independente de reinstall.
REM  1. Detecta WithSecure/F-Secure
REM  2. Se detectado: forca firewall-dns + limpa hosts
REM  3. Agente ativo vai ler config atualizada e parar de
REM     tocar no hosts imediatamente.
REM ============================================================
powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -Command "& { try { $dest = '%DEST%'; $hasWS = $false; $svcPatterns = @('FSMA','FSDFWD','FSecure*','F-Secure*','WithSecure*','FSHoster','FSGK*','FSAUA*','FSORSPClient'); foreach ($p in $svcPatterns) { if (Get-Service -Name $p -ErrorAction SilentlyContinue) { $hasWS = $true; break } }; if (-not $hasWS) { foreach ($f in @(\"$env:ProgramFiles\F-Secure\",\"$env:ProgramFiles\WithSecure\",\"${env:ProgramFiles(x86)}\F-Secure\",\"${env:ProgramFiles(x86)}\WithSecure\")) { if (Test-Path $f) { $hasWS = $true; break } } }; if ($hasWS) { Write-Output 'WithSecure detectado - aplicando firewall-dns'; $cfgPath = Join-Path $dest 'sysmon-config.json'; if (Test-Path $cfgPath) { try { $cfg = Get-Content $cfgPath -Raw | ConvertFrom-Json } catch { $cfg = [PSCustomObject]@{} }; if ($cfg.HostBlockingMethod -ne 'firewall-dns') { $cfg | Add-Member -NotePropertyName HostBlockingMethod -NotePropertyValue 'firewall-dns' -Force; $cfg | ConvertTo-Json -Depth 10 | Set-Content -Path $cfgPath -Encoding ASCII -Force; Write-Output 'sysmon-config.json atualizado (firewall-dns)' } } else { '{\"HostBlockingMethod\":\"firewall-dns\"}' | Set-Content -Path $cfgPath -Encoding ASCII -Force; Write-Output 'sysmon-config.json criado (firewall-dns)' }; $hostsFile = \"$env:WINDIR\System32\drivers\etc\hosts\"; if (Test-Path $hostsFile) { try { $content = Get-Content $hostsFile -Raw -ErrorAction Stop; $pattern = '(?ms)# === WinSysMon Begin ===.*?# === WinSysMon End ===\r?\n?'; if ($content -match $pattern) { $clean = [regex]::Replace($content, $pattern, ''); $clean = $clean.TrimEnd() + \"`r`n\"; try { (Get-Item $hostsFile).IsReadOnly = $false } catch {}; [System.IO.File]::WriteAllText($hostsFile, $clean, (New-Object System.Text.ASCIIEncoding)); & ipconfig.exe /flushdns | Out-Null; Write-Output 'hosts limpo (marcadores WinSysMon removidos)' } } catch { Write-Output \"Erro limpando hosts: $_\" } } } else { Write-Output 'WithSecure nao detectado - sem acao' } } catch { Write-Output \"Auto-reparo falhou: $_\" } }" >> "%LOG%" 2>&1
echo [%DATE% %TIME%] Auto-reparo WithSecure concluido >> "%PUBLOG%"

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

    REM ==== POS-INSTALL: reaplica auto-reparo WithSecure ====
    REM Garante que sysmon-config.json gerado pelo servico
    REM seja corrigido para firewall-dns se WithSecure existir.
    powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -Command "& { try { $hasWS = $false; foreach ($p in @('FSMA','FSDFWD','FSecure*','F-Secure*','WithSecure*','FSHoster','FSGK*','FSAUA*','FSORSPClient')) { if (Get-Service -Name $p -ErrorAction SilentlyContinue) { $hasWS = $true; break } }; if (-not $hasWS) { foreach ($f in @(\"$env:ProgramFiles\F-Secure\",\"$env:ProgramFiles\WithSecure\",\"${env:ProgramFiles(x86)}\F-Secure\",\"${env:ProgramFiles(x86)}\WithSecure\")) { if (Test-Path $f) { $hasWS = $true; break } } }; if ($hasWS) { $cfgPath = '%DEST%\sysmon-config.json'; $cfg = $null; if (Test-Path $cfgPath) { try { $cfg = Get-Content $cfgPath -Raw | ConvertFrom-Json } catch {} }; if (-not $cfg) { $cfg = [PSCustomObject]@{} }; $cfg | Add-Member -NotePropertyName HostBlockingMethod -NotePropertyValue 'firewall-dns' -Force; $cfg | ConvertTo-Json -Depth 10 | Set-Content -Path $cfgPath -Encoding ASCII -Force; sc.exe stop WinSysMon | Out-Null; Start-Sleep -Seconds 2; sc.exe start WinSysMon | Out-Null; Write-Output 'Pos-install: firewall-dns forcado + servico reiniciado' } } catch { Write-Output \"Pos-install falhou: $_\" } }" >> "%LOG%" 2>&1
)

exit /b %RC%
