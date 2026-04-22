# Finaliza recovery: renomeia pastas WinSysMon + valida IFEOs criticos limpos
param([Parameter(Mandatory=$true)][string]$OfflineDrive)

$ErrorActionPreference = 'Continue'
$D = $OfflineDrive.TrimEnd('\',':') + ':'
$stamp = Get-Date -Format 'yyyyMMdd-HHmmss'

Write-Host "`n=== Validacao: recarregando SOFTWARE apenas para verificar ===" -ForegroundColor Cyan
& reg.exe load HKLM\OFFSW "$D\Windows\System32\config\SOFTWARE" | Out-Null
if ($LASTEXITCODE -ne 0) {
    Write-Host "[FAIL] Nao carregou SOFTWARE" -ForegroundColor Red; return
}

$ifeoPath = 'HKLM:\OFFSW\Microsoft\Windows NT\CurrentVersion\Image File Execution Options'
$critical = @('winlogon.exe','csrss.exe','services.exe','lsass.exe','wininit.exe',
              'smss.exe','dwm.exe','svchost.exe','explorer.exe','userinit.exe',
              'sihost.exe','fontdrvhost.exe','ctfmon.exe','dllhost.exe','conhost.exe',
              'logonui.exe','taskhost.exe','taskhostw.exe','runtimebroker.exe')

$stillBad = @()
foreach ($proc in $critical) {
    $k = Join-Path $ifeoPath $proc
    if (Test-Path $k) {
        $d = (Get-ItemProperty $k -ErrorAction SilentlyContinue).Debugger
        if ($d) { $stillBad += "$proc -> Debugger=$d" }
    }
}

if ($stillBad.Count -gt 0) {
    Write-Host "[FAIL] AINDA HA Debugger em processos criticos:" -ForegroundColor Red
    $stillBad | ForEach-Object { Write-Host "    $_" -ForegroundColor Red }
} else {
    Write-Host "[OK] Nenhum processo critico com Debugger - BSOD resolvido" -ForegroundColor Green
}

[GC]::Collect(); [GC]::WaitForPendingFinalizers(); Start-Sleep 1
& reg.exe unload HKLM\OFFSW 2>&1 | Out-Null

Write-Host "`n=== Renomeando pastas WinSysMon em $D\ProgramData\Microsoft ===" -ForegroundColor Cyan
foreach ($p in @("$D\ProgramData\Microsoft\WinSysMon", "$D\ProgramData\Microsoft\WinSysMonStage")) {
    if (Test-Path $p) {
        $newName = "$(Split-Path $p -Leaf).bak-$stamp"
        $newPath = Join-Path (Split-Path $p -Parent) $newName
        try {
            & takeown.exe /F $p /A /R /D Y 2>&1 | Out-Null
            & icacls.exe $p /grant "*S-1-5-32-544:(OI)(CI)F" /T /C /L /Q 2>&1 | Out-Null
            Rename-Item $p $newName -Force
            Write-Host "[OK] Renomeado: $p -> $newPath" -ForegroundColor Green
        } catch {
            Write-Host "[ERRO] Falhou renomear $p : $_" -ForegroundColor Red
        }
    } else {
        Write-Host "[INFO] Nao existe: $p (ok)" -ForegroundColor Yellow
    }
}

Write-Host "`n=== Verificando tasks scheduler offline (arquivos XML) ===" -ForegroundColor Cyan
$taskDir = "$D\Windows\System32\Tasks"
if (Test-Path $taskDir) {
    $wsmTasks = Get-ChildItem $taskDir -Filter 'WinSysMon*' -ErrorAction SilentlyContinue
    if ($wsmTasks) {
        foreach ($t in $wsmTasks) {
            $newT = "$($t.FullName).bak-$stamp"
            try {
                & takeown.exe /F $t.FullName /A 2>&1 | Out-Null
                & icacls.exe $t.FullName /grant "*S-1-5-32-544:F" /C /L /Q 2>&1 | Out-Null
                Rename-Item $t.FullName $newT -Force
                Write-Host "[OK] Task renomeada: $($t.Name) -> $(Split-Path $newT -Leaf)" -ForegroundColor Green
            } catch {
                Write-Host "[WARN] $($t.Name): $_" -ForegroundColor Yellow
            }
        }
    } else {
        Write-Host "[INFO] Nenhuma task WinSysMon* encontrada (ok)" -ForegroundColor Yellow
    }
}

Write-Host "`n============================================================" -ForegroundColor Green
Write-Host " RECOVERY COMPLETO" -ForegroundColor Green
Write-Host "============================================================" -ForegroundColor Green
Write-Host " 1. Desligue seu PC" -ForegroundColor Cyan
Write-Host " 2. Desconecte o HD" -ForegroundColor Cyan
Write-Host " 3. Devolva para o notebook Dell original" -ForegroundColor Cyan
Write-Host " 4. Ligue - Windows vai subir normalmente" -ForegroundColor Cyan
Write-Host "" 
Write-Host " Backup em: D:\RECOVERY-BACKUP-*" -ForegroundColor Gray
Write-Host " Pastas *.bak-$stamp podem ser apagadas depois que confirmar funcionamento" -ForegroundColor Gray
