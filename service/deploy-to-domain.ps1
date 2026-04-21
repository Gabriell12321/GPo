# ══════════════════════════════════════════════════════════════════════
#  deploy-to-domain.ps1
#  Roda no srv-105 como Scheduled Task diario.
#  Enumera todos os computadores ativos do AD e instala o WinSysMon
#  neles via WinRM (Invoke-Command) — independente de GPO.
#
#  Tambem se registra como scheduled task permanente se chamado com
#  -SetupTask (uma vez so, como admin do dominio).
#
#  Uso:
#    .\deploy-to-domain.ps1 -SetupTask     # cria a task recorrente (1x)
#    .\deploy-to-domain.ps1                # executa um ciclo agora
#    .\deploy-to-domain.ps1 -Only PC-12    # so em 1 PC (teste)
# ══════════════════════════════════════════════════════════════════════
param(
    [switch]$SetupTask,
    [string]$Only,
    [string]$ShareInstall = "\\srv-105\Sistema de monitoramento\gpo\aaa\service\install-service.ps1",
    [int]$ParallelLimit = 10,
    [int]$TimeoutSeconds = 120
)

$ErrorActionPreference = 'Continue'
$logFile = Join-Path $PSScriptRoot "deploy-to-domain.log"

function Write-L {
    param($m,$lvl="INFO")
    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $line = "[$ts] [$lvl] $m"
    Write-Host $line
    try { Add-Content -Path $logFile -Value $line -ErrorAction SilentlyContinue } catch {}
}

# ──────────────────────────────────────────────────────────────────────
#  SETUP: registra como scheduled task diaria
# ──────────────────────────────────────────────────────────────────────
if ($SetupTask) {
    $taskName = "WinSysMonDeploy"
    $scriptPath = $MyInvocation.MyCommand.Path
    Write-L "Registrando task $taskName apontando para $scriptPath"
    try { Unregister-ScheduledTask -TaskName $taskName -Confirm:$false -ErrorAction SilentlyContinue } catch {}
    $action    = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`""
    # Dispara no boot do servidor + a cada 6 horas
    $trigger1  = New-ScheduledTaskTrigger -AtStartup
    $trigger2  = New-ScheduledTaskTrigger -Once -At (Get-Date).AddMinutes(10) -RepetitionInterval (New-TimeSpan -Hours 6) -RepetitionDuration ([TimeSpan]::FromDays(3650))
    $principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest
    $settings  = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable -MultipleInstances IgnoreNew -ExecutionTimeLimit (New-TimeSpan -Hours 2)
    Register-ScheduledTask -TaskName $taskName -Action $action -Trigger @($trigger1,$trigger2) -Principal $principal -Settings $settings -Force | Out-Null
    Write-L "Task $taskName registrada (AtStartup + cada 6h). Primeiro disparo em 10 min." "OK"
    exit 0
}

# ──────────────────────────────────────────────────────────────────────
#  1) Listar computadores do AD
# ──────────────────────────────────────────────────────────────────────
Write-L "=== Iniciando ciclo de deploy ==="
if (-not (Test-Path $ShareInstall)) {
    Write-L "install-service.ps1 nao existe no share: $ShareInstall" "ERROR"
    exit 1
}

try { Import-Module ActiveDirectory -ErrorAction Stop } catch {
    Write-L "Modulo ActiveDirectory indisponivel: $($_.Exception.Message)" "ERROR"
    Write-L "Instale RSAT-AD-PowerShell no servidor." "ERROR"
    exit 1
}

try {
    if ($Only) {
        $computers = @(Get-ADComputer -Identity $Only)
    } else {
        # Pega so computadores habilitados e que logaram nos ultimos 30 dias
        $cutoff = (Get-Date).AddDays(-30)
        $computers = Get-ADComputer -Filter { Enabled -eq $true -and LastLogonDate -gt $cutoff } -Properties LastLogonDate, OperatingSystem |
                     Where-Object { $_.OperatingSystem -like "*Windows*" -and $_.OperatingSystem -notlike "*Server*" }
    }
    Write-L "Alvos: $($computers.Count) computadores"
} catch {
    Write-L "Falha ao listar AD: $($_.Exception.Message)" "ERROR"
    exit 1
}

# ──────────────────────────────────────────────────────────────────────
#  2) Scriptblock executado em cada PC
# ──────────────────────────────────────────────────────────────────────
$remote = {
    param($shareInstall)
    $ErrorActionPreference = 'SilentlyContinue'
    $dir = "$env:ProgramData\Microsoft\WinSysMon"
    $scr = Join-Path $dir "winsysmon.ps1"
    $svc = Get-Service WinSysMon -ErrorAction SilentlyContinue

    $reason = $null
    if (-not $svc)                  { $reason = "servico-ausente" }
    elseif ($svc.Status -ne 'Running') { $reason = "servico-parado:$($svc.Status)" }
    elseif (-not (Test-Path $scr))  { $reason = "script-ausente" }

    if (-not $reason) {
        return [pscustomobject]@{ Host=$env:COMPUTERNAME; Status="OK"; Detail="ja-instalado-rodando" }
    }

    if (-not (Test-Path $shareInstall)) {
        return [pscustomobject]@{ Host=$env:COMPUTERNAME; Status="FAIL"; Detail="share-inacessivel:$shareInstall" }
    }

    try {
        & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $shareInstall 2>&1 | Out-Null
        Start-Sleep 3
        $svc = Get-Service WinSysMon -ErrorAction SilentlyContinue
        if ($svc -and $svc.Status -eq 'Running') {
            return [pscustomobject]@{ Host=$env:COMPUTERNAME; Status="INSTALLED"; Detail=$reason }
        } else {
            return [pscustomobject]@{ Host=$env:COMPUTERNAME; Status="PARTIAL"; Detail="install-rodou-mas-servico=$($svc.Status)" }
        }
    } catch {
        return [pscustomobject]@{ Host=$env:COMPUTERNAME; Status="FAIL"; Detail=$_.Exception.Message }
    }
}

# ──────────────────────────────────────────────────────────────────────
#  3) Disparar em paralelo (PS 5.1: usa Start-Job)
# ──────────────────────────────────────────────────────────────────────
$results = @()
$queue   = [System.Collections.Queue]::new()
foreach ($c in $computers) { $queue.Enqueue($c.Name) }

$jobs = @{}
function Drain-Finished {
    foreach ($name in @($jobs.Keys)) {
        $j = $jobs[$name]
        if ($j.State -in 'Completed','Failed','Stopped') {
            try {
                $r = Receive-Job -Job $j -ErrorAction SilentlyContinue
                if ($r) { $script:results += $r; Write-L "  $($r.Host): $($r.Status) — $($r.Detail)" }
                else    { $script:results += [pscustomobject]@{Host=$name;Status="FAIL";Detail="sem-retorno"}; Write-L "  $name : FAIL — sem-retorno" "WARN" }
            } catch { Write-L "  $name : erro-receive: $($_.Exception.Message)" "WARN" }
            Remove-Job -Job $j -Force -ErrorAction SilentlyContinue
            $jobs.Remove($name)
        }
    }
}

Write-L "Disparando ate $ParallelLimit em paralelo..."
while ($queue.Count -gt 0 -or $jobs.Count -gt 0) {
    while ($jobs.Count -lt $ParallelLimit -and $queue.Count -gt 0) {
        $hostName = $queue.Dequeue()
        $j = Start-Job -ScriptBlock {
            param($h,$sb,$sh,$timeout)
            try {
                $res = Invoke-Command -ComputerName $h -ScriptBlock ([scriptblock]::Create($sb)) -ArgumentList $sh -ErrorAction Stop
                return $res
            } catch {
                return [pscustomobject]@{ Host=$h; Status="UNREACHABLE"; Detail=$_.Exception.Message }
            }
        } -ArgumentList $hostName, $remote.ToString(), $ShareInstall, $TimeoutSeconds
        $jobs[$hostName] = $j
    }
    Start-Sleep -Milliseconds 500
    Drain-Finished
    # Kill jobs que excederam timeout
    foreach ($name in @($jobs.Keys)) {
        $j = $jobs[$name]
        if (((Get-Date) - $j.PSBeginTime).TotalSeconds -gt $TimeoutSeconds) {
            Write-L "  $name : TIMEOUT (${TimeoutSeconds}s) — matando job" "WARN"
            Stop-Job -Job $j -ErrorAction SilentlyContinue
            $results += [pscustomobject]@{Host=$name;Status="TIMEOUT";Detail="${TimeoutSeconds}s"}
            Remove-Job -Job $j -Force -ErrorAction SilentlyContinue
            $jobs.Remove($name)
        }
    }
}

# ──────────────────────────────────────────────────────────────────────
#  4) Resumo
# ──────────────────────────────────────────────────────────────────────
$grp = $results | Group-Object Status | Sort-Object Name
Write-L "=== Resumo ==="
foreach ($g in $grp) { Write-L "  $($g.Name): $($g.Count)" }
Write-L "=== Fim ==="
