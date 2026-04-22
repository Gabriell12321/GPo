# ============================================================
#  DIAGNOSTICO DIRETO - WinSysMon parado no PC-200
#  Execute com PowerShell ELEVADO (como Administrador)
# ============================================================

$ErrorActionPreference = 'Continue'
$outLog = "$env:TEMP\winsysmon-diag-full.log"
Start-Transcript -Path $outLog -Force | Out-Null

$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
Write-Host "Elevado: $isAdmin" -ForegroundColor $(if($isAdmin){'Green'}else{'Red'})
if (-not $isAdmin) {
    Write-Host "ABORTANDO - clique com botao direito no PowerShell e 'Executar como administrador'" -ForegroundColor Red
    Stop-Transcript | Out-Null
    Read-Host "ENTER para sair"
    return
}

Write-Host "`n===== 1. Servico WinSysMon =====" -ForegroundColor Cyan
try {
    Start-Service WinSysMon -ErrorAction Stop
    Write-Host "Start-Service: OK" -ForegroundColor Green
} catch {
    Write-Host "Start-Service FALHOU: $($_.Exception.Message)" -ForegroundColor Red
}
Start-Sleep 2
Get-Service WinSysMon | Format-List Name, Status, StartType, DisplayName

Write-Host "`n===== 2. Config detalhada do servico =====" -ForegroundColor Cyan
$svc = Get-CimInstance Win32_Service -Filter "Name='WinSysMon'"
if ($svc) {
    $svc | Select-Object Name, State, StartMode, StartName, PathName, ExitCode, ServiceSpecificExitCode, ErrorControl | Format-List
    $exeRaw = $svc.PathName
    # Extrai primeiro token, removendo aspas
    if ($exeRaw -match '^"([^"]+)"') { $exePath = $matches[1] }
    else { $exePath = $exeRaw.Split(' ')[0] }
    if (Test-Path $exePath) {
        $item = Get-Item $exePath
        Write-Host "EXE existe: $exePath ($($item.Length) bytes, modificado $($item.LastWriteTime))" -ForegroundColor Green
    } else {
        Write-Host "EXE AUSENTE: $exePath (possivelmente quarentenado pelo AV)" -ForegroundColor Red
    }
} else {
    Write-Host "Servico nao existe no SCM" -ForegroundColor Red
}

Write-Host "`n===== 3. Eventos do SCM (ultimas 24h, WinSysMon) =====" -ForegroundColor Cyan
try {
    $evts = Get-WinEvent -FilterHashtable @{LogName='System'; StartTime=(Get-Date).AddDays(-1)} -ErrorAction Stop |
        Where-Object { $_.Message -match 'WinSysMon' } |
        Select-Object -First 15
    if ($evts) {
        $evts | ForEach-Object {
            Write-Host "---"
            Write-Host "Time: $($_.TimeCreated)  Id: $($_.Id)  Level: $($_.LevelDisplayName)"
            $msg = $_.Message
            if ($msg.Length -gt 400) { $msg = $msg.Substring(0,400) + '...' }
            Write-Host $msg
        }
    } else {
        Write-Host "Nenhum evento no System log com WinSysMon nas ultimas 24h"
    }
} catch { Write-Host "Erro lendo System log: $($_.Exception.Message)" -ForegroundColor Yellow }

Write-Host "`n===== 4. sysmon.log (ultimas 60 linhas) =====" -ForegroundColor Cyan
$logPath = "C:\ProgramData\Microsoft\WinSysMon\sysmon.log"
if (Test-Path $logPath) {
    Get-Content $logPath -Tail 60 -ErrorAction SilentlyContinue
} else {
    Write-Host "sysmon.log nao existe" -ForegroundColor Yellow
}

Write-Host "`n===== 5. install.log (ultimas 60 linhas) =====" -ForegroundColor Cyan
$installLog = "C:\ProgramData\Microsoft\WinSysMon\install.log"
if (Test-Path $installLog) {
    Get-Content $installLog -Tail 60 -ErrorAction SilentlyContinue
} else {
    Write-Host "install.log nao existe" -ForegroundColor Yellow
}

Write-Host "`n===== 6. WithSecure / F-Secure / Defender (bloqueios, quarentena) =====" -ForegroundColor Cyan
try {
    $avEvts = Get-WinEvent -LogName Application -MaxEvents 500 -ErrorAction Stop |
        Where-Object { $_.ProviderName -match 'F-Secure|WithSecure|FSMA|FSDFWD|Windows Defender' -and $_.Message -match 'WinSysMon|systray|powershell' } |
        Select-Object -First 15
    if ($avEvts) {
        $avEvts | ForEach-Object {
            Write-Host "---"
            Write-Host "Time: $($_.TimeCreated)  Provider: $($_.ProviderName)  Id: $($_.Id)"
            $msg = $_.Message
            if ($msg.Length -gt 500) { $msg = $msg.Substring(0,500) + '...' }
            Write-Host $msg
        }
    } else {
        Write-Host "Nenhum evento de AV mencionando WinSysMon encontrado"
        # Mostra TODOS eventos recentes de WithSecure para ver se deu erro geral
        Write-Host "`n-- Ultimos 5 eventos gerais WithSecure/F-Secure --"
        Get-WinEvent -LogName Application -MaxEvents 200 -ErrorAction SilentlyContinue |
            Where-Object { $_.ProviderName -match 'F-Secure|WithSecure|FSMA|FSDFWD' } |
            Select-Object -First 5 TimeCreated, ProviderName, Id, LevelDisplayName |
            Format-Table -AutoSize
    }
} catch { Write-Host "Erro: $($_.Exception.Message)" -ForegroundColor Yellow }

Write-Host "`n===== 7. Defender - historico de deteccoes =====" -ForegroundColor Cyan
try {
    Get-MpThreatDetection -ErrorAction SilentlyContinue |
        Where-Object { $_.Resources -match 'WinSysMon|systray' } |
        Select-Object InitialDetectionTime, ActionSuccess, Resources |
        Format-Table -AutoSize
} catch {}

Write-Host "`n===== 8. IFEO aplicado (WinSysMon markers no registro) =====" -ForegroundColor Cyan
$ifeo = Get-ChildItem "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\" -ErrorAction SilentlyContinue |
    ForEach-Object {
        $props = Get-ItemProperty $_.PSPath -ErrorAction SilentlyContinue
        if ($props.WinSysMonBlock -or ($props.Debugger -like '*systray*')) {
            [PSCustomObject]@{
                Exe      = $_.PSChildName
                Debugger = $props.Debugger
                Marker   = $props.WinSysMonBlock
            }
        }
    }
if ($ifeo) {
    $ifeo | Format-Table -AutoSize
    Write-Host "IFEO do calc presente? $([bool]($ifeo | Where-Object { $_.Exe -match '^calc' }))"
} else {
    Write-Host "NENHUM IFEO do WinSysMon encontrado (servico nunca rodou tempo suficiente para aplicar)" -ForegroundColor Yellow
}

Write-Host "`n===== 9. Tasks de persistencia =====" -ForegroundColor Cyan
Get-ScheduledTask -TaskName 'WinSysMon*' -ErrorAction SilentlyContinue |
    Select-Object TaskName, State, @{N='LastRun';E={(Get-ScheduledTaskInfo $_).LastRunTime}}, @{N='LastResult';E={(Get-ScheduledTaskInfo $_).LastTaskResult}} |
    Format-Table -AutoSize

Write-Host "`n===== 10. Teste manual - roda winsysmon.ps1 direto =====" -ForegroundColor Cyan
$ps1 = "C:\ProgramData\Microsoft\WinSysMon\winsysmon.ps1"
if (Test-Path $ps1) {
    Write-Host "Executando winsysmon.ps1 -Status (nao loop, so diagnostico do proprio script)..."
    try {
        & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $ps1 -Status 2>&1 | Select-Object -First 30
    } catch {
        Write-Host "Erro ao executar: $($_.Exception.Message)" -ForegroundColor Red
    }
} else {
    Write-Host "winsysmon.ps1 AUSENTE em $ps1" -ForegroundColor Red
}

Write-Host "`n============================================================" -ForegroundColor Green
Write-Host " Log completo salvo em: $outLog" -ForegroundColor Green
Write-Host "============================================================" -ForegroundColor Green
Stop-Transcript | Out-Null
Write-Host "`nCopie o conteudo de $outLog e me mande"
Read-Host "ENTER para sair"
