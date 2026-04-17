# ══════════════════════════════════════════════════════════════
#  IPPEL Agent - Servico de Monitoramento e Bloqueio
#  Roda em segundo plano como Tarefa Agendada ou Servico
# ══════════════════════════════════════════════════════════════
#Requires -RunAsAdministrator

param(
    [switch]$Install,
    [switch]$Uninstall,
    [switch]$Status
)

# ── Configuracao ──
$script:AgentName    = "IppelAgent"
$script:AgentVersion = "1.0.0"
$script:BasePath     = $PSScriptRoot
$script:DbPath       = Join-Path $script:BasePath "ippel-agent.db"
$script:ConfigPath   = Join-Path $script:BasePath "agent-config.json"
$script:LogPath      = Join-Path $script:BasePath "agent.log"
$script:PollInterval = 5  # segundos entre cada verificacao

# ── Logging ──
function Write-Log {
    param([string]$Msg, [string]$Level = "INFO")
    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $line = "[$ts] [$Level] $Msg"
    Add-Content -Path $script:LogPath -Value $line -ErrorAction SilentlyContinue
    if ($Level -eq "ERROR") { Write-Host $line -ForegroundColor Red }
}

# ══════════════════════════════════════════════════════════════
#  SQLite via .NET System.Data.SQLite ou ADO.NET
# ══════════════════════════════════════════════════════════════
$script:SqliteDll = $null
$script:DbConn    = $null

function Initialize-Database {
    # Tentar carregar SQLite - primeiro Microsoft.Data.Sqlite, depois System.Data.SQLite
    $loaded = $false

    # Opcao 1: SQLite embutido no .NET (funciona no PS 5.1 com o DLL junto)
    $localDll = Join-Path $script:BasePath "System.Data.SQLite.dll"
    if (Test-Path $localDll) {
        try {
            Add-Type -Path $localDll
            $script:SqliteDll = "System.Data.SQLite"
            $loaded = $true
        } catch {}
    }

    # Opcao 2: Usar o interop SQLite via Add-Type inline (puro .NET, sem DLL externo)
    if (-not $loaded) {
        try {
            # Baixar SQLite se nao existe
            $dllUrl = "https://www.nuget.org/api/v2/package/System.Data.SQLite.Core/1.0.118.0"
            $nugetPath = Join-Path $script:BasePath "sqlite-nuget.zip"
            $extractPath = Join-Path $script:BasePath "sqlite-nuget"

            if (-not (Test-Path $localDll)) {
                Write-Log "Baixando SQLite DLL..."
                [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
                Invoke-WebRequest -Uri $dllUrl -OutFile $nugetPath -UseBasicParsing
                Expand-Archive -Path $nugetPath -DestinationPath $extractPath -Force
                $found = Get-ChildItem $extractPath -Recurse -Filter "System.Data.SQLite.dll" |
                    Where-Object { $_.FullName -like "*net46*" -or $_.FullName -like "*net45*" } |
                    Select-Object -First 1
                if ($found) {
                    Copy-Item $found.FullName -Destination $localDll -Force
                    # Copiar interop tambem
                    $interop = Get-ChildItem $found.Directory.Parent.FullName -Recurse -Filter "SQLite.Interop.dll" |
                        Where-Object { $_.FullName -like "*x64*" } | Select-Object -First 1
                    if ($interop) {
                        Copy-Item $interop.FullName -Destination (Join-Path $script:BasePath "SQLite.Interop.dll") -Force
                    }
                }
                Remove-Item $nugetPath -Force -ErrorAction SilentlyContinue
                Remove-Item $extractPath -Recurse -Force -ErrorAction SilentlyContinue
            }

            if (Test-Path $localDll) {
                Add-Type -Path $localDll
                $script:SqliteDll = "System.Data.SQLite"
                $loaded = $true
                Write-Log "SQLite DLL carregado com sucesso"
            }
        } catch {
            Write-Log "Falha ao baixar SQLite: $($_.Exception.Message)" "ERROR"
        }
    }

    if (-not $loaded) {
        Write-Log "SQLite nao disponivel. Usando fallback CSV." "WARN"
        return $false
    }

    # Criar/abrir banco
    try {
        $connStr = "Data Source=$($script:DbPath);Version=3;Journal Mode=WAL;"
        $script:DbConn = New-Object System.Data.SQLite.SQLiteConnection($connStr)
        $script:DbConn.Open()

        # Criar tabelas
        $sql = @"
CREATE TABLE IF NOT EXISTS blocked_log (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    timestamp TEXT NOT NULL,
    hostname TEXT NOT NULL,
    username TEXT NOT NULL,
    process_name TEXT NOT NULL,
    process_path TEXT,
    action TEXT DEFAULT 'KILLED'
);

CREATE TABLE IF NOT EXISTS login_log (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    timestamp TEXT NOT NULL,
    hostname TEXT NOT NULL,
    username TEXT NOT NULL,
    event_type TEXT NOT NULL,
    source_ip TEXT,
    details TEXT
);

CREATE TABLE IF NOT EXISTS machine_info (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    hostname TEXT UNIQUE NOT NULL,
    last_seen TEXT NOT NULL,
    os_version TEXT,
    ip_address TEXT,
    mac_address TEXT,
    cpu TEXT,
    ram_gb REAL,
    disk_total_gb REAL,
    disk_free_gb REAL,
    uptime_hours REAL,
    agent_version TEXT
);

CREATE TABLE IF NOT EXISTS app_usage (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    timestamp TEXT NOT NULL,
    hostname TEXT NOT NULL,
    username TEXT NOT NULL,
    process_name TEXT NOT NULL,
    window_title TEXT,
    duration_sec INTEGER DEFAULT 0
);

CREATE TABLE IF NOT EXISTS blocked_apps (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    pattern TEXT UNIQUE NOT NULL,
    added_date TEXT NOT NULL,
    added_by TEXT,
    reason TEXT
);
"@
        Invoke-Sql $sql
        Write-Log "Banco de dados inicializado: $($script:DbPath)"
        return $true
    } catch {
        Write-Log "Erro ao criar banco: $($_.Exception.Message)" "ERROR"
        return $false
    }
}

function Invoke-Sql {
    param([string]$Query, [hashtable]$Params = @{})
    if (-not $script:DbConn -or $script:DbConn.State -ne 'Open') { return $null }
    $cmd = $script:DbConn.CreateCommand()
    $cmd.CommandText = $Query
    foreach ($k in $Params.Keys) {
        $cmd.Parameters.AddWithValue($k, $Params[$k]) | Out-Null
    }
    if ($Query.TrimStart() -match "^(INSERT|UPDATE|DELETE|CREATE|DROP)") {
        $cmd.ExecuteNonQuery() | Out-Null
    } else {
        $reader = $cmd.ExecuteReader()
        $table = New-Object System.Data.DataTable
        $table.Load($reader)
        $reader.Close()
        return $table
    }
    $cmd.Dispose()
}

# ══════════════════════════════════════════════════════════════
#  CONFIGURACAO
# ══════════════════════════════════════════════════════════════
function Load-Config {
    $default = @{
        BlockedApps     = @()
        PollInterval    = 5
        MonitorLogins   = $true
        MonitorApps     = $true
        CollectHardware = $true
        HardwareInterval = 3600  # 1 hora
    }
    if (Test-Path $script:ConfigPath) {
        try {
            $json = Get-Content $script:ConfigPath -Raw | ConvertFrom-Json
            foreach ($prop in $json.PSObject.Properties) {
                $default[$prop.Name] = $prop.Value
            }
        } catch {
            Write-Log "Erro ao ler config: $($_.Exception.Message)" "WARN"
        }
    }
    return $default
}

function Save-Config {
    param($Config)
    $Config | ConvertTo-Json -Depth 5 | Out-File $script:ConfigPath -Encoding UTF8 -Force
}

# ══════════════════════════════════════════════════════════════
#  BLOQUEIO DE APLICATIVOS
# ══════════════════════════════════════════════════════════════
function Get-BlockedPatterns {
    $patterns = @()
    # Do config file
    $cfg = Load-Config
    if ($cfg.BlockedApps) { $patterns += $cfg.BlockedApps }
    # Do banco de dados
    try {
        $rows = Invoke-Sql "SELECT pattern FROM blocked_apps"
        if ($rows) {
            foreach ($row in $rows) { $patterns += $row.pattern }
        }
    } catch {}
    return $patterns | Select-Object -Unique
}

function Enforce-AppBlocking {
    $blocked = Get-BlockedPatterns
    if ($blocked.Count -eq 0) { return }

    $hostname = $env:COMPUTERNAME
    $username = $env:USERNAME

    $procs = Get-Process -ErrorAction SilentlyContinue |
        Where-Object { $_.Id -ne $PID -and $_.SessionId -ne 0 } |
        Select-Object Id, ProcessName, @{N='Path';E={try{$_.Path}catch{$null}}}

    foreach ($proc in $procs) {
        $name = $proc.ProcessName.ToLower()
        $path = if ($proc.Path) { $proc.Path.ToLower() } else { "" }

        foreach ($pattern in $blocked) {
            $p = $pattern.ToLower().Trim()
            if (-not $p) { continue }

            $match = $false
            if ($p.Contains("*") -or $p.Contains("?")) {
                $match = ($name -like $p) -or ($path -like $p)
            } else {
                $match = ($name -eq $p) -or ($name -eq [System.IO.Path]::GetFileNameWithoutExtension($p)) -or ($path -eq $p)
            }

            if ($match) {
                try {
                    Stop-Process -Id $proc.Id -Force -ErrorAction Stop
                    Write-Log "BLOQUEADO: $($proc.ProcessName) (PID $($proc.Id)) usuario=$username padrao=$pattern"
                    Invoke-Sql "INSERT INTO blocked_log (timestamp, hostname, username, process_name, process_path, action) VALUES (@ts, @host, @user, @name, @path, 'KILLED')" @{
                        "@ts"   = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
                        "@host" = $hostname
                        "@user" = $username
                        "@name" = $proc.ProcessName
                        "@path" = $proc.Path
                    }
                } catch {
                    Write-Log "Falha ao matar $($proc.ProcessName): $($_.Exception.Message)" "ERROR"
                }
                break
            }
        }
    }
}

# ══════════════════════════════════════════════════════════════
#  MONITORAMENTO DE LOGINS (Event Log)
# ══════════════════════════════════════════════════════════════
$script:LastLoginCheck = (Get-Date)

function Monitor-Logins {
    $hostname = $env:COMPUTERNAME
    $now = Get-Date
    $since = $script:LastLoginCheck

    try {
        # Evento 4624 = Logon bem sucedido, 4625 = Logon falho
        $events = Get-WinEvent -FilterHashtable @{
            LogName   = 'Security'
            Id        = 4624, 4625
            StartTime = $since
        } -MaxEvents 50 -ErrorAction SilentlyContinue

        foreach ($evt in $events) {
            $xml = [xml]$evt.ToXml()
            $data = $xml.Event.EventData.Data
            $targetUser = ($data | Where-Object { $_.Name -eq 'TargetUserName' }).'#text'
            $targetDomain = ($data | Where-Object { $_.Name -eq 'TargetDomainName' }).'#text'
            $logonType = ($data | Where-Object { $_.Name -eq 'LogonType' }).'#text'
            $sourceIP = ($data | Where-Object { $_.Name -eq 'IpAddress' }).'#text'

            # Ignorar contas de sistema
            if ($targetUser -match '\$$' -or $targetUser -in @('SYSTEM','LOCAL SERVICE','NETWORK SERVICE','DWM-1','UMFD-0','UMFD-1')) { continue }
            if ($targetDomain -in @('Window Manager','Font Driver Host','NT AUTHORITY')) { continue }

            $eventType = if ($evt.Id -eq 4624) { "LOGIN_OK" } else { "LOGIN_FAIL" }
            $details = "LogonType=$logonType Domain=$targetDomain"

            Invoke-Sql "INSERT INTO login_log (timestamp, hostname, username, event_type, source_ip, details) VALUES (@ts, @host, @user, @type, @ip, @det)" @{
                "@ts"   = $evt.TimeCreated.ToString("yyyy-MM-dd HH:mm:ss")
                "@host" = $hostname
                "@user" = $targetUser
                "@type" = $eventType
                "@ip"   = $sourceIP
                "@det"  = $details
            }
        }
    } catch {
        Write-Log "Erro ao ler eventos de login: $($_.Exception.Message)" "WARN"
    }

    $script:LastLoginCheck = $now
}

# ══════════════════════════════════════════════════════════════
#  COLETA DE HARDWARE/SISTEMA
# ══════════════════════════════════════════════════════════════
$script:LastHwCollect = [datetime]::MinValue

function Collect-MachineInfo {
    $hostname = $env:COMPUTERNAME
    $now = Get-Date

    try {
        $os = Get-CimInstance Win32_OperatingSystem -ErrorAction Stop
        $cpu = (Get-CimInstance Win32_Processor -ErrorAction Stop | Select-Object -First 1).Name
        $net = Get-CimInstance Win32_NetworkAdapterConfiguration -ErrorAction Stop |
            Where-Object { $_.IPEnabled -and $_.IPAddress } | Select-Object -First 1
        $disk = Get-CimInstance Win32_LogicalDisk -Filter "DeviceID='C:'" -ErrorAction Stop

        $ip = if ($net) { ($net.IPAddress | Where-Object { $_ -match '^\d+\.\d+\.\d+\.\d+$' } | Select-Object -First 1) } else { "" }
        $mac = if ($net) { $net.MACAddress } else { "" }
        $ramGB = [math]::Round($os.TotalVisibleMemorySize / 1MB, 1)
        $diskTotalGB = [math]::Round($disk.Size / 1GB, 1)
        $diskFreeGB = [math]::Round($disk.FreeSpace / 1GB, 1)
        $uptime = (New-TimeSpan -Start $os.LastBootUpTime -End $now).TotalHours
        $uptimeH = [math]::Round($uptime, 1)

        Invoke-Sql "INSERT OR REPLACE INTO machine_info (hostname, last_seen, os_version, ip_address, mac_address, cpu, ram_gb, disk_total_gb, disk_free_gb, uptime_hours, agent_version) VALUES (@host, @seen, @os, @ip, @mac, @cpu, @ram, @dtot, @dfree, @up, @ver)" @{
            "@host"  = $hostname
            "@seen"  = $now.ToString("yyyy-MM-dd HH:mm:ss")
            "@os"    = "$($os.Caption) $($os.Version)"
            "@ip"    = $ip
            "@mac"   = $mac
            "@cpu"   = $cpu
            "@ram"   = $ramGB
            "@dtot"  = $diskTotalGB
            "@dfree" = $diskFreeGB
            "@up"    = $uptimeH
            "@ver"   = $script:AgentVersion
        }

        Write-Log "Hardware coletado: IP=$ip RAM=${ramGB}GB Disco=${diskFreeGB}/${diskTotalGB}GB Uptime=${uptimeH}h"
    } catch {
        Write-Log "Erro ao coletar hardware: $($_.Exception.Message)" "ERROR"
    }

    $script:LastHwCollect = $now
}

# ══════════════════════════════════════════════════════════════
#  INSTALACAO / DESINSTALACAO (Tarefa Agendada)
# ══════════════════════════════════════════════════════════════
function Install-Agent {
    Write-Host "Instalando $script:AgentName como Tarefa Agendada..." -ForegroundColor Cyan

    $scriptPath = $MyInvocation.ScriptName
    if (-not $scriptPath) { $scriptPath = Join-Path $script:BasePath "ippel-agent.ps1" }

    $action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-ExecutionPolicy Bypass -WindowStyle Hidden -File `"$scriptPath`"" -WorkingDirectory $script:BasePath
    $trigger = New-ScheduledTaskTrigger -AtStartup
    $principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest
    $settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable -RestartInterval (New-TimeSpan -Minutes 1) -RestartCount 3 -ExecutionTimeLimit ([TimeSpan]::Zero)

    try {
        Unregister-ScheduledTask -TaskName $script:AgentName -Confirm:$false -ErrorAction SilentlyContinue
        Register-ScheduledTask -TaskName $script:AgentName -Action $action -Trigger $trigger -Principal $principal -Settings $settings -Description "IPPEL Agent - Monitoramento e Bloqueio de Aplicativos" -Force | Out-Null

        # Iniciar agora
        Start-ScheduledTask -TaskName $script:AgentName
        Write-Host "Instalado e iniciado com sucesso!" -ForegroundColor Green
        Write-Host "  Banco de dados: $script:DbPath" -ForegroundColor Gray
        Write-Host "  Log: $script:LogPath" -ForegroundColor Gray
        Write-Host "  Config: $script:ConfigPath" -ForegroundColor Gray
    } catch {
        Write-Host "Erro ao instalar: $($_.Exception.Message)" -ForegroundColor Red
    }
}

function Uninstall-Agent {
    Write-Host "Removendo $script:AgentName..." -ForegroundColor Yellow
    try {
        Stop-ScheduledTask -TaskName $script:AgentName -ErrorAction SilentlyContinue
        Unregister-ScheduledTask -TaskName $script:AgentName -Confirm:$false -ErrorAction SilentlyContinue
        Write-Host "Removido com sucesso." -ForegroundColor Green
        Write-Host "Banco de dados preservado em: $script:DbPath" -ForegroundColor Gray
    } catch {
        Write-Host "Erro: $($_.Exception.Message)" -ForegroundColor Red
    }
}

function Show-Status {
    Write-Host "`n=== $script:AgentName v$script:AgentVersion ===" -ForegroundColor Cyan

    $task = Get-ScheduledTask -TaskName $script:AgentName -ErrorAction SilentlyContinue
    if ($task) {
        $info = Get-ScheduledTaskInfo -TaskName $script:AgentName
        Write-Host "  Estado: $($task.State)" -ForegroundColor $(if($task.State -eq 'Running'){'Green'}else{'Yellow'})
        Write-Host "  Ultima execucao: $($info.LastRunTime)" -ForegroundColor Gray
    } else {
        Write-Host "  NAO INSTALADO" -ForegroundColor Red
    }

    if (Test-Path $script:DbPath) {
        $sz = [math]::Round((Get-Item $script:DbPath).Length / 1KB, 1)
        Write-Host "  Banco: $script:DbPath ($sz KB)" -ForegroundColor Gray
    }

    if ($script:DbConn -and $script:DbConn.State -eq 'Open') {
        try {
            $bl = (Invoke-Sql "SELECT COUNT(*) as c FROM blocked_log").c
            $ll = (Invoke-Sql "SELECT COUNT(*) as c FROM login_log").c
            $mi = (Invoke-Sql "SELECT COUNT(*) as c FROM machine_info").c
            Write-Host "  Bloqueios registrados: $bl" -ForegroundColor Gray
            Write-Host "  Eventos de login: $ll" -ForegroundColor Gray
            Write-Host "  Maquinas registradas: $mi" -ForegroundColor Gray
        } catch {}
    }
    Write-Host ""
}

# ══════════════════════════════════════════════════════════════
#  LOOP PRINCIPAL
# ══════════════════════════════════════════════════════════════
function Start-AgentLoop {
    Write-Log "=== $script:AgentName v$script:AgentVersion iniciado em $env:COMPUTERNAME ==="

    $dbOk = Initialize-Database
    if (-not $dbOk) {
        Write-Log "Banco de dados nao inicializado. Abortando." "ERROR"
        return
    }

    $cfg = Load-Config
    $script:PollInterval = $cfg.PollInterval

    # Coleta inicial de hardware
    if ($cfg.CollectHardware) { Collect-MachineInfo }

    $iteration = 0
    while ($true) {
        try {
            $cfg = Load-Config

            # Bloqueio de apps - a cada ciclo
            Enforce-AppBlocking

            # Monitoramento de logins - a cada 30s
            if ($cfg.MonitorLogins -and ($iteration % 6 -eq 0)) {
                Monitor-Logins
            }

            # Coleta de hardware - a cada HardwareInterval
            if ($cfg.CollectHardware) {
                $elapsed = (New-TimeSpan -Start $script:LastHwCollect -End (Get-Date)).TotalSeconds
                if ($elapsed -ge $cfg.HardwareInterval) {
                    Collect-MachineInfo
                }
            }

        } catch {
            Write-Log "Erro no loop principal: $($_.Exception.Message)" "ERROR"
        }

        $iteration++
        Start-Sleep -Seconds $script:PollInterval
    }
}

# ══════════════════════════════════════════════════════════════
#  ENTRY POINT
# ══════════════════════════════════════════════════════════════
if ($Install) {
    Initialize-Database
    Install-Agent
} elseif ($Uninstall) {
    Uninstall-Agent
} elseif ($Status) {
    Initialize-Database
    Show-Status
} else {
    # Modo servico - rodar o loop
    Start-AgentLoop
}
