# ============================================================
#  FIX ACL + Start WinSysMon
#  Execute com PowerShell ELEVADO como Administrador
#  no PC onde o servico esta com "Acesso negado" (Event 7000)
# ============================================================

$ErrorActionPreference = 'Continue'
$folder = "C:\ProgramData\Microsoft\WinSysMon"

$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Write-Host "ABORTANDO - nao esta elevado. Abra PowerShell como Administrador." -ForegroundColor Red
    Read-Host "ENTER para sair"
    return
}

Write-Host "==> 1. Parando servico e tasks (se estiverem vivos)..." -ForegroundColor Cyan
Stop-Service WinSysMon -Force -ErrorAction SilentlyContinue
Stop-ScheduledTask -TaskName WinSysMonWatchdog -ErrorAction SilentlyContinue
Stop-ScheduledTask -TaskName WinSysMonGuard -ErrorAction SilentlyContinue
Stop-ScheduledTask -TaskName WinSysMonIFEORefresh -ErrorAction SilentlyContinue

Write-Host "`n==> 2. ACL ANTES do fix:" -ForegroundColor Cyan
try {
    $aclBefore = Get-Acl $folder -ErrorAction Stop
    Write-Host "  Owner: $($aclBefore.Owner)"
    Write-Host "  ACEs:"
    $aclBefore.Access | ForEach-Object { Write-Host "    $($_.IdentityReference) - $($_.FileSystemRights) - $($_.AccessControlType)" }
} catch {
    Write-Host "  Nem consegue ler ACL: $($_.Exception.Message)" -ForegroundColor Yellow
}

Write-Host "`n==> 3. Tomando ownership recursivo (takeown)..." -ForegroundColor Cyan
takeown.exe /F $folder /R /D S /A 2>&1 | Select-Object -Last 3

Write-Host "`n==> 4. Concedendo FullControl para Administradores (icacls)..." -ForegroundColor Cyan
icacls.exe $folder /grant "*S-1-5-32-544:(OI)(CI)F" /T /C 2>&1 | Select-Object -Last 3

Write-Host "`n==> 5. Resetando ACL para heranca + permissoes corretas..." -ForegroundColor Cyan
# Remove todas as ACEs explicitas e herda do pai primeiro
icacls.exe $folder /reset /T /C 2>&1 | Select-Object -Last 3

# Agora aplica a ACL correta: SYSTEM FC, Administrators FC, Users R
Write-Host "`n==> 6. Aplicando ACL correta (SYSTEM=FC, Admins=FC, Users=R)..." -ForegroundColor Cyan
icacls.exe $folder /inheritance:r /T /C 2>&1 | Out-Null
icacls.exe $folder /grant "*S-1-5-18:(OI)(CI)F" /T /C 2>&1 | Out-Null      # SYSTEM
icacls.exe $folder /grant "*S-1-5-32-544:(OI)(CI)F" /T /C 2>&1 | Out-Null  # Administrators
icacls.exe $folder /grant "*S-1-5-32-545:(OI)(CI)RX" /T /C 2>&1 | Out-Null # Users (read-only)

Write-Host "`n==> 7. ACL DEPOIS do fix:" -ForegroundColor Cyan
$aclAfter = Get-Acl $folder
Write-Host "  Owner: $($aclAfter.Owner)"
Write-Host "  ACEs:"
$aclAfter.Access | ForEach-Object { Write-Host "    $($_.IdentityReference) - $($_.FileSystemRights) - $($_.AccessControlType)" }

Write-Host "`n==> 8. Validando leitura dos arquivos criticos..." -ForegroundColor Cyan
foreach ($f in @("WinSysMonSvc.exe", "winsysmon.ps1", "sysmon-config.json")) {
    $p = Join-Path $folder $f
    if (Test-Path $p) {
        try {
            [System.IO.File]::OpenRead($p).Close()
            Write-Host "  [OK] $f lido com sucesso" -ForegroundColor Green
        } catch {
            Write-Host "  [FAIL] $f : $($_.Exception.Message)" -ForegroundColor Red
        }
    } else {
        Write-Host "  [AUSENTE] $f" -ForegroundColor Yellow
    }
}

Write-Host "`n==> 9. Start-Service WinSysMon..." -ForegroundColor Cyan
try {
    Start-Service WinSysMon -ErrorAction Stop
    Start-Sleep 3
    $s = Get-Service WinSysMon
    if ($s.Status -eq 'Running') {
        Write-Host "  [OK] Servico RODANDO" -ForegroundColor Green
    } else {
        Write-Host "  [WARN] Servico nao subiu - Status: $($s.Status)" -ForegroundColor Yellow
    }
} catch {
    Write-Host "  [FAIL] $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host "`n==> 10. Ultimas 15 linhas do sysmon.log:" -ForegroundColor Cyan
$log = Join-Path $folder "sysmon.log"
if (Test-Path $log) {
    Get-Content $log -Tail 15 -ErrorAction SilentlyContinue
}

Write-Host "`n==> 11. IFEO atual:" -ForegroundColor Cyan
Get-ChildItem "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\" -ErrorAction SilentlyContinue |
    ForEach-Object {
        $props = Get-ItemProperty $_.PSPath -ErrorAction SilentlyContinue
        if ($props.WinSysMonBlock) {
            [PSCustomObject]@{ Exe=$_.PSChildName; Debugger=$props.Debugger }
        }
    } | Format-Table -AutoSize

Write-Host "`n============================================================" -ForegroundColor Green
Write-Host " FIX aplicado. Se servico subiu, WinSysMon voltou a operar." -ForegroundColor Green
Write-Host " Teste: abra a calculadora - NAO deve abrir (IFEO ativo)." -ForegroundColor Green
Write-Host "============================================================" -ForegroundColor Green
Read-Host "ENTER para sair"
