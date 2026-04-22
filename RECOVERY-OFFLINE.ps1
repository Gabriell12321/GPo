# =====================================================================
#  RECOVERY-OFFLINE.ps1
#  Recupera PC travado por IFEO/ACL do WinSysMon com HD montado em outro
#  computador como disco secundario.
#
#  NAO APAGA NADA. NAO FORMATA. So RENOMEIA pastas (.bak) e faz BACKUP
#  dos hives do registro antes de mexer. Tudo reversivel.
#
#  USO:
#    1. Conecte o HD do PC ruim no seu notebook (ja esta como D:)
#    2. Abra PowerShell COMO ADMINISTRADOR
#    3. cd c:\gpo
#    4. .\RECOVERY-OFFLINE.ps1 -OfflineDrive D:
# =====================================================================

[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)]
    [string]$OfflineDrive   # ex: "D:"
)

$ErrorActionPreference = 'Stop'

function Write-Step { param($msg) Write-Host "`n=== $msg ===" -ForegroundColor Cyan }
function Write-OK   { param($msg) Write-Host "  [OK] $msg"   -ForegroundColor Green }
function Write-Warn { param($msg) Write-Host "  [AVISO] $msg" -ForegroundColor Yellow }
function Write-Err  { param($msg) Write-Host "  [ERRO] $msg"  -ForegroundColor Red }

# ---- Validacao inicial --------------------------------------------------
$D = $OfflineDrive.TrimEnd('\',':') + ':'

$adminSid = New-Object System.Security.Principal.SecurityIdentifier('S-1-5-32-544')
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole($adminSid)) {
    Write-Err "Rode como Administrador."
    exit 1
}

# Drive existe?
if (-not (Test-Path "$D\" -ErrorAction SilentlyContinue)) {
    Write-Err "Drive $D nao encontrado."
    exit 1
}

# Tomar posse apenas das hives SOFTWARE e SYSTEM (suficiente, evita recursao em pastas problematicas)
Write-Step "0) Tomando posse dos hives SOFTWARE e SYSTEM em $D\Windows\System32\config"
$cfg = "$D\Windows\System32\config"
$ErrorActionPreference = 'Continue'
foreach ($hive in @('SOFTWARE','SYSTEM')) {
    $hivePath = Join-Path $cfg $hive
    & takeown.exe /F $hivePath /A 2>&1 | Out-Null
    if ($LASTEXITCODE -ne 0) { & takeown.exe /F $hivePath 2>&1 | Out-Null }
    & icacls.exe $hivePath /grant "*S-1-5-32-544:F" /C /L /Q 2>&1 | Out-Null
}
# Pasta config em si (para poder ler/escrever arquivos .LOG auxiliares se necessario)
& takeown.exe /F $cfg /A 2>&1 | Out-Null
if ($LASTEXITCODE -ne 0) { & takeown.exe /F $cfg 2>&1 | Out-Null }
& icacls.exe $cfg /grant "*S-1-5-32-544:(OI)(CI)F" /C /L /Q 2>&1 | Out-Null
$ErrorActionPreference = 'Stop'
Write-OK "Posse tomada em SOFTWARE e SYSTEM"

if (-not (Test-Path "$D\Windows\System32\config\SOFTWARE")) {
    Write-Err "Nao encontrei $D\Windows\System32\config\SOFTWARE. Drive errado?"
    exit 1
}

$stamp   = Get-Date -Format 'yyyyMMdd-HHmmss'
$bakRoot = "$D\RECOVERY-BACKUP-$stamp"
New-Item -ItemType Directory -Path $bakRoot -Force | Out-Null
Write-Step "Backup sera salvo em: $bakRoot"

# ---- 1) Backup dos hives do registro ------------------------------------
Write-Step "1) Copiando hives SOFTWARE e SYSTEM para backup"
Copy-Item "$D\Windows\System32\config\SOFTWARE" "$bakRoot\SOFTWARE.hive" -Force
Copy-Item "$D\Windows\System32\config\SYSTEM"   "$bakRoot\SYSTEM.hive"   -Force
Write-OK "Hives copiados ($bakRoot\SOFTWARE.hive / SYSTEM.hive)"

# ---- 2) Carregar hives offline ------------------------------------------
Write-Step "2) Carregando hives offline (HKLM\OFFSW e HKLM\OFFSYS)"
# Descarrega se ja estiver (limpeza de execucao anterior) - ignora erros
$ErrorActionPreference = 'Continue'
cmd /c "reg unload HKLM\OFFSW  >nul 2>&1"
cmd /c "reg unload HKLM\OFFSYS >nul 2>&1"
$ErrorActionPreference = 'Stop'

& reg.exe load HKLM\OFFSW  "$D\Windows\System32\config\SOFTWARE" | Out-Null
if ($LASTEXITCODE -ne 0) { Write-Err "Falha ao carregar SOFTWARE"; exit 2 }
& reg.exe load HKLM\OFFSYS "$D\Windows\System32\config\SYSTEM"   | Out-Null
if ($LASTEXITCODE -ne 0) { & reg.exe unload HKLM\OFFSW | Out-Null; Write-Err "Falha ao carregar SYSTEM"; exit 2 }
Write-OK "Hives carregados"

try {
    # ---- 3) Auditar IFEO --------------------------------------------------
    Write-Step "3) Auditando IFEOs (somente leitura)"
    $ifeoPath = 'HKLM:\OFFSW\Microsoft\Windows NT\CurrentVersion\Image File Execution Options'
    $critical = @(
        'winlogon.exe','csrss.exe','services.exe','lsass.exe','wininit.exe',
        'smss.exe','dwm.exe','svchost.exe','explorer.exe','userinit.exe',
        'sihost.exe','fontdrvhost.exe','ctfmon.exe','dllhost.exe','conhost.exe',
        'logonui.exe','taskhost.exe','taskhostw.exe','runtimebroker.exe'
    )

    $toRemove = @()
    $auditLog = @()

    if (Test-Path $ifeoPath) {
        Get-ChildItem $ifeoPath -ErrorAction SilentlyContinue | ForEach-Object {
            $name      = $_.PSChildName
            $props     = Get-ItemProperty $_.PSPath -ErrorAction SilentlyContinue
            $dbg       = $props.Debugger
            $marker    = $props.WinSysMonBlock
            $isCrit    = $critical -contains $name.ToLower()
            $hasSystray= $dbg -match '(?i)systray\.exe'

            $reason = @()
            if ($marker)   { $reason += 'WinSysMonBlock=1' }
            if ($hasSystray){ $reason += "Debugger=$dbg" }

            if ($reason.Count -gt 0) {
                $line = "{0,-35} CRITICO={1,-5} {2}" -f $name, $isCrit, ($reason -join ' | ')
                $auditLog += $line
                $toRemove  += [pscustomobject]@{
                    Name     = $name
                    Path     = $_.PSPath
                    Critical = $isCrit
                    Reason   = ($reason -join ' | ')
                    Debugger = $dbg
                }
            }
        }
    }

    if ($auditLog.Count -eq 0) {
        Write-OK "Nenhum IFEO suspeito encontrado."
    } else {
        Write-Warn "Encontrados $($auditLog.Count) IFEOs para remocao:"
        $auditLog | ForEach-Object { Write-Host "  $_" -ForegroundColor Gray }
        $auditLog | Out-File "$bakRoot\ifeo-removidos.txt" -Encoding UTF8
    }

    # ---- 4) Exportar chaves antes de remover -----------------------------
    Write-Step "4) Exportando IFEOs suspeitos para .reg (backup reversivel)"
    if ($toRemove.Count -gt 0) {
        $regBakDir = Join-Path $bakRoot 'ifeo-backup'
        New-Item -ItemType Directory -Path $regBakDir -Force | Out-Null
        foreach ($entry in $toRemove) {
            $safeName = $entry.Name -replace '[^\w\.\-]','_'
            $regFile  = Join-Path $regBakDir "$safeName.reg"
            $regKey   = "HKLM\OFFSW\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\$($entry.Name)"
            & reg.exe export $regKey $regFile /y 2>$null | Out-Null
        }
        Write-OK "Backup .reg salvo em $regBakDir"
    }

    # ---- 5) Remover IFEOs suspeitos --------------------------------------
    Write-Step "5) Removendo IFEOs suspeitos"
    foreach ($entry in $toRemove) {
        try {
            Remove-Item $entry.Path -Recurse -Force
            Write-OK "Removido: $($entry.Name)"
        } catch {
            Write-Err "Falhou remover $($entry.Name): $_"
        }
    }

    # ---- 6) Backup do servico WinSysMon antes de remover -----------------
    Write-Step "6) Fazendo backup e removendo servico WinSysMon do registro"
    foreach ($cs in @('ControlSet001','ControlSet002')) {
        $svcKey = "HKLM\OFFSYS\$cs\Services\WinSysMon"
        $svcPs  = "HKLM:\OFFSYS\$cs\Services\WinSysMon"
        if (Test-Path $svcPs) {
            $regFile = Join-Path $bakRoot "service-$cs.reg"
            & reg.exe export $svcKey $regFile /y 2>$null | Out-Null
            Remove-Item $svcPs -Recurse -Force
            Write-OK "Servico removido de $cs (backup: $regFile)"
        } else {
            Write-Warn "Servico WinSysMon nao existe em $cs (ok)"
        }
    }

    # ---- 7) Verificar processos criticos ---------------------------------
    Write-Step "7) Verificacao final: processos criticos nao devem ter Debugger"
    $stillBad = @()
    foreach ($proc in $critical) {
        $k = Join-Path $ifeoPath $proc
        if (Test-Path $k) {
            $d = (Get-ItemProperty $k -ErrorAction SilentlyContinue).Debugger
            if ($d) {
                $stillBad += "$proc -> Debugger=$d"
            }
        }
    }
    if ($stillBad.Count -gt 0) {
        Write-Err "AINDA HA Debugger em processos criticos:"
        $stillBad | ForEach-Object { Write-Host "    $_" -ForegroundColor Red }
    } else {
        Write-OK "Nenhum processo critico tem Debugger. BSOD deve estar resolvido."
    }

} finally {
    # ---- 8) Descarregar hives (SEMPRE) -----------------------------------
    Write-Step "8) Descarregando hives"
    [GC]::Collect(); [GC]::WaitForPendingFinalizers(); [GC]::Collect()
    Start-Sleep -Seconds 1
    & reg.exe unload HKLM\OFFSW  2>&1 | Out-Null
    & reg.exe unload HKLM\OFFSYS 2>&1 | Out-Null
    Write-OK "Hives descarregados"
}

# ---- 9) Renomear pastas WinSysMon (NAO apaga) --------------------------
Write-Step "9) Renomeando pastas WinSysMon (nao apaga)"
$winsysmon      = "$D\ProgramData\Microsoft\WinSysMon"
$winsysmonStage = "$D\ProgramData\Microsoft\WinSysMonStage"

foreach ($p in @($winsysmon, $winsysmonStage)) {
    if (Test-Path $p) {
        $newName = "$(Split-Path $p -Leaf).bak-$stamp"
        $newPath = Join-Path (Split-Path $p -Parent) $newName
        try {
            # Toma posse so para conseguir renomear (ainda seu disco, funciona)
            & takeown.exe /F $p /A /R /D Y 2>&1 | Out-Null
            if ($LASTEXITCODE -ne 0) { & takeown.exe /F $p /A /R /D S 2>&1 | Out-Null }
            & icacls.exe $p /grant "*S-1-5-32-544:(OI)(CI)F" /T /C /L /Q 2>&1 | Out-Null
            Rename-Item $p $newName -Force
            Write-OK "Renomeado: $p -> $newPath"
        } catch {
            Write-Err "Falhou renomear $p : $_"
        }
    } else {
        Write-Warn "Nao existe: $p (ok)"
    }
}

# ---- 10) Relatorio final ------------------------------------------------
Write-Step "CONCLUIDO"
Write-Host ""
Write-Host "  Backup completo em: $bakRoot" -ForegroundColor Cyan
Write-Host "    - SOFTWARE.hive / SYSTEM.hive  (hives originais)" -ForegroundColor Gray
Write-Host "    - ifeo-backup\*.reg            (IFEOs removidos)"  -ForegroundColor Gray
Write-Host "    - service-ControlSet*.reg      (servico WinSysMon)" -ForegroundColor Gray
Write-Host "    - ifeo-removidos.txt           (lista auditoria)"   -ForegroundColor Gray
Write-Host ""
Write-Host "  Pastas renomeadas (NAO apagadas):" -ForegroundColor Cyan
Get-ChildItem "$D\ProgramData\Microsoft" -Directory -Filter 'WinSysMon*.bak-*' -ErrorAction SilentlyContinue |
    ForEach-Object { Write-Host "    $($_.FullName)" -ForegroundColor Gray }
Write-Host ""
Write-Host "  PROXIMO PASSO: desligar PC, devolver HD, ligar." -ForegroundColor Green
Write-Host "  Se algo der errado, tudo em $bakRoot pode restaurar." -ForegroundColor Green
Write-Host ""
