# ══════════════════════════════════════════════════════════════
#  WinSysMon - Windows System Monitor
#  Monitors system performance and reports diagnostics
# ══════════════════════════════════════════════════════════════
#Requires -RunAsAdministrator

param(
    [switch]$Install,
    [switch]$Uninstall,
    [switch]$Status,
    [switch]$RunLoop,
    [string]$SharePath = ""
)

# ── Configuracao ──
$script:ServiceName  = "WinSysMon"
$script:DisplayName  = "Windows System Monitor"
$script:Description  = "Monitors system performance and reports diagnostics data."
$script:AgentVersion = "1.1.0"
$script:BasePath     = $PSScriptRoot
$script:DbPath       = Join-Path $script:BasePath "sysmon.db"
$script:ConfigPath   = Join-Path $script:BasePath "sysmon-config.json"
$script:LogPath      = Join-Path $script:BasePath "sysmon.log"
$script:PollInterval = 1
$script:LogMaxSizeBytes = 5MB

# ── Logging com rotacao ──
function Write-Log {
    param([string]$Msg, [string]$Level = "INFO")
    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    try {
        # Rotacao simples: se log > 5MB, renomeia para .old
        if (Test-Path $script:LogPath) {
            $size = (Get-Item $script:LogPath -ErrorAction SilentlyContinue).Length
            if ($size -gt $script:LogMaxSizeBytes) {
                $oldLog = "$($script:LogPath).old"
                if (Test-Path $oldLog) { Remove-Item $oldLog -Force -ErrorAction SilentlyContinue }
                Move-Item $script:LogPath $oldLog -Force -ErrorAction SilentlyContinue
            }
        }
        Add-Content -Path $script:LogPath -Value "[$ts] [$Level] $Msg" -ErrorAction SilentlyContinue
    } catch {}
}

# ══════════════════════════════════════════════════════════════
#  SQLite
# ══════════════════════════════════════════════════════════════
$script:DbConn = $null
$script:DbDisabled = $false

function Initialize-Database {
    $loaded = $false
    $localDll = Join-Path $script:BasePath "System.Data.SQLite.dll"

    if (Test-Path $localDll) {
        try { Add-Type -Path $localDll; $loaded = $true } catch {}
    }

    if (-not $loaded) {
        try {
            $dllUrl = "https://www.nuget.org/api/v2/package/System.Data.SQLite.Core/1.0.118.0"
            $nugetPath = Join-Path $script:BasePath "sqlite-nuget.zip"
            $extractPath = Join-Path $script:BasePath "sqlite-nuget"
            if (-not (Test-Path $localDll)) {
                Write-Log "Baixando SQLite..."
                [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
                Invoke-WebRequest -Uri $dllUrl -OutFile $nugetPath -UseBasicParsing
                Expand-Archive -Path $nugetPath -DestinationPath $extractPath -Force
                $found = Get-ChildItem $extractPath -Recurse -Filter "System.Data.SQLite.dll" |
                    Where-Object { $_.FullName -like "*net46*" -or $_.FullName -like "*net45*" } | Select-Object -First 1
                if ($found) {
                    Copy-Item $found.FullName -Destination $localDll -Force
                    $interop = Get-ChildItem $found.Directory.Parent.FullName -Recurse -Filter "SQLite.Interop.dll" |
                        Where-Object { $_.FullName -like "*x64*" } | Select-Object -First 1
                    if ($interop) { Copy-Item $interop.FullName -Destination (Join-Path $script:BasePath "SQLite.Interop.dll") -Force }
                }
                Remove-Item $nugetPath -Force -ErrorAction SilentlyContinue
                Remove-Item $extractPath -Recurse -Force -ErrorAction SilentlyContinue
            }
            if (Test-Path $localDll) { Add-Type -Path $localDll; $loaded = $true }
        } catch { Write-Log "Falha SQLite: $($_.Exception.Message)" "ERROR" }
    }

    if (-not $loaded) {
        Write-Log "SQLite indisponivel - continuando SEM banco (bloqueio ainda funciona)" "WARN"
        $script:DbDisabled = $true
        return $true
    }

    try {
        $script:DbConn = New-Object System.Data.SQLite.SQLiteConnection("Data Source=$($script:DbPath);Version=3;Journal Mode=WAL;")
        $script:DbConn.Open()
        Invoke-Sql @"
CREATE TABLE IF NOT EXISTS blocked_log (
    id INTEGER PRIMARY KEY AUTOINCREMENT, timestamp TEXT NOT NULL,
    hostname TEXT NOT NULL, username TEXT NOT NULL,
    process_name TEXT NOT NULL, process_path TEXT, action TEXT DEFAULT 'KILLED');
CREATE TABLE IF NOT EXISTS login_log (
    id INTEGER PRIMARY KEY AUTOINCREMENT, timestamp TEXT NOT NULL,
    hostname TEXT NOT NULL, username TEXT NOT NULL,
    event_type TEXT NOT NULL, source_ip TEXT, details TEXT);
CREATE TABLE IF NOT EXISTS machine_info (
    id INTEGER PRIMARY KEY AUTOINCREMENT, hostname TEXT UNIQUE NOT NULL,
    last_seen TEXT NOT NULL, os_version TEXT, ip_address TEXT, mac_address TEXT,
    cpu TEXT, ram_gb REAL, disk_total_gb REAL, disk_free_gb REAL,
    uptime_hours REAL, agent_version TEXT);
CREATE TABLE IF NOT EXISTS app_usage (
    id INTEGER PRIMARY KEY AUTOINCREMENT, timestamp TEXT NOT NULL,
    hostname TEXT NOT NULL, username TEXT NOT NULL,
    process_name TEXT NOT NULL, window_title TEXT, duration_sec INTEGER DEFAULT 0);
CREATE TABLE IF NOT EXISTS blocked_apps (
    id INTEGER PRIMARY KEY AUTOINCREMENT, pattern TEXT UNIQUE NOT NULL,
    added_date TEXT NOT NULL, added_by TEXT, reason TEXT);
"@
        Write-Log "DB inicializado: $($script:DbPath)"
        return $true
    } catch {
        Write-Log "Erro DB: $($_.Exception.Message) - continuando SEM banco" "WARN"
        $script:DbDisabled = $true
        return $true
    }
}

function Invoke-Sql {
    param([string]$Query, [hashtable]$Params = @{})
    if ($script:DbDisabled) { return $null }
    if (-not $script:DbConn -or $script:DbConn.State -ne 'Open') { return $null }
    $cmd = $script:DbConn.CreateCommand()
    $cmd.CommandText = $Query
    foreach ($k in $Params.Keys) { $cmd.Parameters.AddWithValue($k, $Params[$k]) | Out-Null }
    if ($Query.TrimStart() -match "^(INSERT|UPDATE|DELETE|CREATE|DROP)") {
        $cmd.ExecuteNonQuery() | Out-Null
    } else {
        $reader = $cmd.ExecuteReader()
        $table = New-Object System.Data.DataTable
        $table.Load($reader); $reader.Close()
        return $table
    }
    $cmd.Dispose()
}

# ══════════════════════════════════════════════════════════════
#  CONFIGURACAO
# ══════════════════════════════════════════════════════════════
function Load-Config {
    $default = @{ BlockedApps=@(); PollInterval=1; MonitorLogins=$true; MonitorApps=$true; CollectHardware=$true; HardwareInterval=3600; RemoteBlockedAppsPath="" }
    if (Test-Path $script:ConfigPath) {
        try {
            $json = Get-Content $script:ConfigPath -Raw | ConvertFrom-Json
            foreach ($prop in $json.PSObject.Properties) { $default[$prop.Name] = $prop.Value }
        } catch {}
    }
    return $default
}

# ══════════════════════════════════════════════════════════════
#  BLOQUEIO DE APLICATIVOS
# ══════════════════════════════════════════════════════════════
$script:RemoteBlockCache = @()
$script:LastRemoteFetch = [datetime]::MinValue
$script:LastKnownGoodPatterns = @()
$script:CachePatternsFile = $null  # inicializado no Start-AgentLoop

function Save-PatternsCache {
    param([array]$Patterns)
    if (-not $script:CachePatternsFile) { return }
    try {
        $json = @{ timestamp = (Get-Date).ToString("o"); patterns = $Patterns } | ConvertTo-Json -Depth 3
        [System.IO.File]::WriteAllText($script:CachePatternsFile, $json, (New-Object System.Text.UTF8Encoding($false)))
    } catch {}
}

function Load-PatternsCache {
    if (-not $script:CachePatternsFile -or -not (Test-Path $script:CachePatternsFile)) { return @() }
    try {
        $obj = Get-Content $script:CachePatternsFile -Raw | ConvertFrom-Json
        if ($obj.patterns) { return @($obj.patterns) }
    } catch {}
    return @()
}

function Get-BlockedPatterns {
    $patterns = @()
    $cfg = Load-Config

    # 1) Apps locais do config
    if ($cfg.BlockedApps) { $patterns += $cfg.BlockedApps }

    # 2) Apps do banco local
    try {
        $rows = Invoke-Sql "SELECT pattern FROM blocked_apps"
        if ($rows) { foreach ($row in $rows) { $patterns += $row.pattern } }
    } catch {}

    # 3) Apps do share de rede (atualiza a cada 30s, com fallback para cache)
    if ($cfg.RemoteBlockedAppsPath) {
        $elapsed = (New-TimeSpan -Start $script:LastRemoteFetch -End (Get-Date)).TotalSeconds
        if ($elapsed -ge 30) {
            $fetched = $false
            try {
                if (Test-Path $cfg.RemoteBlockedAppsPath -ErrorAction Stop) {
                    $remote = Get-Content $cfg.RemoteBlockedAppsPath -Raw -ErrorAction Stop | ConvertFrom-Json
                    $hostname = $env:COMPUTERNAME
                    $newCache = @()
                    if ($remote.Global) { $newCache += $remote.Global }
                    if ($remote.Machines -and $remote.Machines.PSObject.Properties[$hostname]) {
                        $newCache += $remote.Machines.$hostname
                    }
                    $script:RemoteBlockCache = $newCache
                    $script:LastRemoteFetch = Get-Date
                    $fetched = $true
                    Save-PatternsCache -Patterns $newCache
                    Write-Log "Remote blocked apps: $($newCache.Count) patterns"
                }
            } catch {
                Write-Log "Erro ao ler share (usando cache): $($_.Exception.Message)" "WARN"
            }
            if (-not $fetched -and $script:RemoteBlockCache.Count -eq 0) {
                # Sem rede + sem cache em memoria: carrega do disco
                $script:RemoteBlockCache = Load-PatternsCache
                if ($script:RemoteBlockCache.Count -gt 0) {
                    Write-Log "Usando cache em disco: $($script:RemoteBlockCache.Count) patterns"
                }
            }
        }
        $patterns += $script:RemoteBlockCache
    }

    # Filtra e normaliza: so strings nao vazias
    $clean = @()
    foreach ($pat in $patterns) {
        if ($null -eq $pat) { continue }
        $s = ([string]$pat).Trim()
        if ($s.Length -gt 0) { $clean += $s }
    }
    $final = @($clean | Select-Object -Unique)
    if ($final.Count -gt 0) { $script:LastKnownGoodPatterns = $final }
    return $final
}

function Show-BlockedNotification {
    param([string]$ProcessName, [int]$SessionId)

    # Rate limit: nao notificar o mesmo app mais de 1x a cada 60s
    if (-not $script:LastNotifyMap) { $script:LastNotifyMap = @{} }
    $key = "$SessionId|$($ProcessName.ToLower())"
    $now = Get-Date
    if ($script:LastNotifyMap.ContainsKey($key)) {
        $elapsed = ($now - $script:LastNotifyMap[$key]).TotalSeconds
        if ($elapsed -lt 60) { return }
    }
    $script:LastNotifyMap[$key] = $now

    # Nao notificar sessao 0 (SYSTEM/services)
    if ($SessionId -le 0) { return }
    # Mostra popup na sessao do usuario via msg.exe ou WScript (roda como SYSTEM, precisa atingir a sessao)
    $appName = (Get-Culture).TextInfo.ToTitleCase($ProcessName.ToLower())
    $title = "Aplicativo Bloqueado"
    $message = "O aplicativo `"$appName`" foi bloqueado pelo administrador do sistema.`n`nSe voce precisa deste programa, entre em contato com o setor de TI."

    # Metodo 1: msg.exe (funciona em sessoes RDP e console)
    try {
        $si = New-Object System.Diagnostics.ProcessStartInfo
        $si.FileName = "$env:SystemRoot\System32\msg.exe"
        $si.Arguments = "$SessionId /TIME:15 `"$message`""
        $si.CreateNoWindow = $true
        $si.WindowStyle = [System.Diagnostics.ProcessWindowStyle]::Hidden
        [System.Diagnostics.Process]::Start($si) | Out-Null
        return
    } catch {}

    # Metodo 2: mshta (popup HTML - fallback)
    try {
        $escapedMsg = $message -replace '"','\"' -replace "`n",'\n'
        $hta = "javascript:var sh=new ActiveXObject('WScript.Shell');sh.Popup('$escapedMsg',15,'$title',0x30);close();"
        # Criar processo na sessao do usuario
        $si2 = New-Object System.Diagnostics.ProcessStartInfo
        $si2.FileName = "$env:SystemRoot\System32\mshta.exe"
        $si2.Arguments = "`"$hta`""
        $si2.CreateNoWindow = $true
        [System.Diagnostics.Process]::Start($si2) | Out-Null
    } catch {}
}

function Enforce-AppBlocking {
    param([switch]$Notify)
    try {
        $blocked = Get-BlockedPatterns
    } catch { Write-Log "Erro Get-BlockedPatterns: $($_.Exception.Message)" "WARN"; return }
    if ($null -eq $blocked -or $blocked.Count -eq 0) { return }
    $hostname = $env:COMPUTERNAME
    $procs = Get-Process -ErrorAction SilentlyContinue |
        Where-Object { $_.Id -ne $PID -and $_.SessionId -ne 0 } |
        Select-Object Id, ProcessName, SessionId, @{N='Path';E={try{$_.Path}catch{$null}}}

    foreach ($proc in $procs) {
        $name = ([string]$proc.ProcessName).ToLower()
        $path = if ($proc.Path) { ([string]$proc.Path).ToLower() } else { "" }
        foreach ($pattern in $blocked) {
            if ($null -eq $pattern) { continue }
            $p = ([string]$pattern).ToLower().Trim()
            if (-not $p) { continue }
            $match = if ($p.Contains("*") -or $p.Contains("?")) {
                ($name -like $p) -or ($path -like $p)
            } else {
                ($name -eq $p) -or ($name -eq [System.IO.Path]::GetFileNameWithoutExtension($p)) -or ($path -eq $p)
            }
            if ($match) {
                try {
                    $owner = ""; try { $owner = (Get-CimInstance Win32_Process -Filter "ProcessId=$($proc.Id)" -ErrorAction Stop).GetOwner().User } catch {}
                    Stop-Process -Id $proc.Id -Force -ErrorAction Stop
                    Write-Log "BLOQUEADO: $($proc.ProcessName) PID=$($proc.Id) user=$owner"
                    Invoke-Sql "INSERT INTO blocked_log (timestamp,hostname,username,process_name,process_path,action) VALUES (@ts,@host,@user,@name,@path,'KILLED')" @{
                        "@ts"=(Get-Date).ToString("yyyy-MM-dd HH:mm:ss"); "@host"=$hostname; "@user"=$owner; "@name"=$proc.ProcessName; "@path"=$proc.Path
                    }
                    # Notificar usuario na sessao dele (apenas quando -Notify)
                    if ($Notify) {
                        Show-BlockedNotification -ProcessName $proc.ProcessName -SessionId $proc.SessionId
                    }
                } catch {
                    # Fallback: taskkill nativo (contorna alguns casos de Access Denied)
                    try { & taskkill.exe /F /PID $proc.Id 2>$null | Out-Null; Write-Log "BLOQUEADO(taskkill): $($proc.ProcessName) PID=$($proc.Id)" } catch {}
                }
                break
            }
        }
    }
}

# ══════════════════════════════════════════════════════════════
#  MONITORAMENTO DE LOGINS
# ══════════════════════════════════════════════════════════════
$script:LastLoginCheck = (Get-Date)

function Monitor-Logins {
    $hostname = $env:COMPUTERNAME; $now = Get-Date
    try {
        $events = Get-WinEvent -FilterHashtable @{ LogName='Security'; Id=4624,4625; StartTime=$script:LastLoginCheck } -MaxEvents 50 -ErrorAction SilentlyContinue
        foreach ($evt in $events) {
            $xml = [xml]$evt.ToXml(); $data = $xml.Event.EventData.Data
            $targetUser = ($data | Where-Object { $_.Name -eq 'TargetUserName' }).'#text'
            $targetDomain = ($data | Where-Object { $_.Name -eq 'TargetDomainName' }).'#text'
            $logonType = ($data | Where-Object { $_.Name -eq 'LogonType' }).'#text'
            $sourceIP = ($data | Where-Object { $_.Name -eq 'IpAddress' }).'#text'
            if ($targetUser -match '\$$' -or $targetUser -in @('SYSTEM','LOCAL SERVICE','NETWORK SERVICE','DWM-1','UMFD-0','UMFD-1')) { continue }
            if ($targetDomain -in @('Window Manager','Font Driver Host','NT AUTHORITY')) { continue }
            $eventType = if ($evt.Id -eq 4624) { "LOGIN_OK" } else { "LOGIN_FAIL" }
            Invoke-Sql "INSERT INTO login_log (timestamp,hostname,username,event_type,source_ip,details) VALUES (@ts,@host,@user,@type,@ip,@det)" @{
                "@ts"=$evt.TimeCreated.ToString("yyyy-MM-dd HH:mm:ss"); "@host"=$hostname; "@user"=$targetUser
                "@type"=$eventType; "@ip"=$sourceIP; "@det"="LogonType=$logonType Domain=$targetDomain"
            }
        }
    } catch {}
    $script:LastLoginCheck = $now
}

# ══════════════════════════════════════════════════════════════
#  COLETA DE HARDWARE
# ══════════════════════════════════════════════════════════════
$script:LastHwCollect = [datetime]::MinValue

function Collect-MachineInfo {
    $hostname = $env:COMPUTERNAME; $now = Get-Date
    try {
        $os = Get-CimInstance Win32_OperatingSystem
        $cpu = (Get-CimInstance Win32_Processor | Select-Object -First 1).Name
        $net = Get-CimInstance Win32_NetworkAdapterConfiguration | Where-Object { $_.IPEnabled -and $_.IPAddress } | Select-Object -First 1
        $disk = Get-CimInstance Win32_LogicalDisk -Filter "DeviceID='C:'"
        $ip = if ($net) { ($net.IPAddress | Where-Object { $_ -match '^\d+\.\d+' } | Select-Object -First 1) } else { "" }
        $mac = if ($net) { $net.MACAddress } else { "" }
        $ramGB = [math]::Round($os.TotalVisibleMemorySize / 1MB, 1)
        Invoke-Sql "INSERT OR REPLACE INTO machine_info (hostname,last_seen,os_version,ip_address,mac_address,cpu,ram_gb,disk_total_gb,disk_free_gb,uptime_hours,agent_version) VALUES (@host,@seen,@os,@ip,@mac,@cpu,@ram,@dtot,@dfree,@up,@ver)" @{
            "@host"=$hostname; "@seen"=$now.ToString("yyyy-MM-dd HH:mm:ss"); "@os"="$($os.Caption) $($os.Version)"
            "@ip"=$ip; "@mac"=$mac; "@cpu"=$cpu; "@ram"=$ramGB
            "@dtot"=[math]::Round($disk.Size/1GB,1); "@dfree"=[math]::Round($disk.FreeSpace/1GB,1)
            "@up"=[math]::Round((New-TimeSpan -Start $os.LastBootUpTime -End $now).TotalHours,1); "@ver"=$script:AgentVersion
        }
        Write-Log "HW coletado: IP=$ip RAM=${ramGB}GB"
    } catch { Write-Log "Erro HW: $($_.Exception.Message)" "ERROR" }
    $script:LastHwCollect = $now
}

# ══════════════════════════════════════════════════════════════
#  INSTALACAO COMO SERVICO WINDOWS
# ══════════════════════════════════════════════════════════════
function Install-Agent {
    Write-Host "`n=== Instalando $script:DisplayName ===" -ForegroundColor Cyan

    # 1) Pasta discreta em ProgramData\Microsoft
    $installDir = "$env:ProgramData\Microsoft\WinSysMon"
    if (-not (Test-Path $installDir)) { New-Item $installDir -ItemType Directory -Force | Out-Null }

    # 2) Copiar arquivos
    $destScript = Join-Path $installDir "WinSysMon.ps1"
    Copy-Item -Path $PSCommandPath -Destination $destScript -Force
    $destCfg = Join-Path $installDir "sysmon-config.json"

    # Mergear config existente (preserva) com novos campos
    $cfgObj = @{ BlockedApps=@(); PollInterval=1; MonitorLogins=$true; MonitorApps=$true; CollectHardware=$true; HardwareInterval=3600; RemoteBlockedAppsPath="" }
    if (Test-Path $destCfg) {
        try {
            $existing = Get-Content $destCfg -Raw | ConvertFrom-Json
            foreach ($p in $existing.PSObject.Properties) { $cfgObj[$p.Name] = $p.Value }
        } catch {}
    } elseif (Test-Path $script:ConfigPath) {
        try {
            $existing = Get-Content $script:ConfigPath -Raw | ConvertFrom-Json
            foreach ($p in $existing.PSObject.Properties) { $cfgObj[$p.Name] = $p.Value }
        } catch {}
    }

    # Se usuario passou -SharePath, usa. Senao, tenta detectar default (scripts/service/blocked-apps.json)
    if ($SharePath) {
        $cfgObj.RemoteBlockedAppsPath = $SharePath
        Write-Host "  Share configurado: $SharePath" -ForegroundColor Green
    } elseif (-not $cfgObj.RemoteBlockedAppsPath) {
        $guess = Join-Path (Split-Path $PSCommandPath -Parent) "blocked-apps.json"
        if (Test-Path $guess) {
            $cfgObj.RemoteBlockedAppsPath = $guess
            Write-Host "  Share detectado: $guess" -ForegroundColor Green
        }
    }

    $cfgJson = $cfgObj | ConvertTo-Json -Depth 5
    [System.IO.File]::WriteAllText($destCfg, $cfgJson, (New-Object System.Text.UTF8Encoding($false)))
    foreach ($dll in @("System.Data.SQLite.dll","SQLite.Interop.dll")) {
        $src = Join-Path $script:BasePath $dll
        if (Test-Path $src) { Copy-Item $src -Destination (Join-Path $installDir $dll) -Force }
    }

    # 3) Proteger pasta - apenas SYSTEM e Administrators (via SID, funciona em qualquer idioma)
    try {
        $sidSystem = New-Object System.Security.Principal.SecurityIdentifier([System.Security.Principal.WellKnownSidType]::LocalSystemSid, $null)
        $sidAdmins = New-Object System.Security.Principal.SecurityIdentifier([System.Security.Principal.WellKnownSidType]::BuiltinAdministratorsSid, $null)
        $acl = New-Object System.Security.AccessControl.DirectorySecurity
        $acl.SetAccessRuleProtection($true, $false)
        $acl.AddAccessRule((New-Object System.Security.AccessControl.FileSystemAccessRule($sidSystem,"FullControl","ContainerInherit,ObjectInherit","None","Allow")))
        $acl.AddAccessRule((New-Object System.Security.AccessControl.FileSystemAccessRule($sidAdmins,"FullControl","ContainerInherit,ObjectInherit","None","Allow")))
        Set-Acl -Path $installDir -AclObject $acl
        Write-Host "  Pasta protegida (SYSTEM + Admins)" -ForegroundColor Gray
    } catch { Write-Host "  Aviso ACL: $($_.Exception.Message)" -ForegroundColor Yellow }

    # 4) Remover tarefa/servico antigo
    Stop-ScheduledTask -TaskName $script:ServiceName -ErrorAction SilentlyContinue
    Unregister-ScheduledTask -TaskName $script:ServiceName -Confirm:$false -ErrorAction SilentlyContinue
    Unregister-ScheduledTask -TaskName "WinSysMon_Legacy" -Confirm:$false -ErrorAction SilentlyContinue

    # 5) Criar tarefa agendada como servico (SYSTEM, Hidden, reinicio automatico)
    $action = New-ScheduledTaskAction `
        -Execute "powershell.exe" `
        -Argument "-ExecutionPolicy Bypass -WindowStyle Hidden -NoProfile -NonInteractive -File `"$destScript`" -RunLoop" `
        -WorkingDirectory $installDir

    $trigger = New-ScheduledTaskTrigger -AtStartup
    $principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest
    $settings = New-ScheduledTaskSettingsSet `
        -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable `
        -RestartInterval (New-TimeSpan -Minutes 1) -RestartCount 999 `
        -ExecutionTimeLimit ([TimeSpan]::Zero) -Hidden

    Register-ScheduledTask -TaskName $script:ServiceName -Action $action -Trigger $trigger `
        -Principal $principal -Settings $settings -Description $script:Description -Force | Out-Null

    # 6) Proteger tarefa via SDDL (somente SYSTEM e Admins podem ver/modificar)
    try {
        $ts = New-Object -ComObject "Schedule.Service"
        $ts.Connect()
        $task = $ts.GetFolder("\").GetTask($script:ServiceName)
        $task.SetSecurityDescriptor("D:P(A;;FA;;;SY)(A;;FA;;;BA)", 0)
        Write-Host "  Tarefa oculta para usuarios comuns" -ForegroundColor Gray
    } catch { Write-Host "  Aviso SDDL: $($_.Exception.Message)" -ForegroundColor Yellow }

    # 7) Iniciar
    Start-ScheduledTask -TaskName $script:ServiceName

    Write-Host "`n  INSTALADO!" -ForegroundColor Green
    Write-Host "  Pasta: $installDir" -ForegroundColor Gray
    Write-Host "  Config: $installDir\sysmon-config.json" -ForegroundColor Gray
    Write-Host "  Log: $installDir\sysmon.log`n" -ForegroundColor Gray
}

function Uninstall-Agent {
    Write-Host "`nRemovendo $script:ServiceName..." -ForegroundColor Yellow
    Stop-ScheduledTask -TaskName $script:ServiceName -ErrorAction SilentlyContinue
    Unregister-ScheduledTask -TaskName $script:ServiceName -Confirm:$false -ErrorAction SilentlyContinue
    Write-Host "Removido." -ForegroundColor Green
    Write-Host "  Pasta preservada: $env:ProgramData\Microsoft\WinSysMon" -ForegroundColor Gray
}

function Show-Status {
    Write-Host "`n=== $script:DisplayName v$script:AgentVersion ===" -ForegroundColor Cyan
    $task = Get-ScheduledTask -TaskName $script:ServiceName -ErrorAction SilentlyContinue
    if ($task) {
        $info = Get-ScheduledTaskInfo -TaskName $script:ServiceName
        Write-Host "  Estado: $($task.State)" -ForegroundColor $(if($task.State -eq 'Running'){'Green'}else{'Yellow'})
        Write-Host "  Ultima exec: $($info.LastRunTime)" -ForegroundColor Gray
    } else { Write-Host "  NAO INSTALADO" -ForegroundColor Red }
    $dbFile = "$env:ProgramData\Microsoft\WinSysMon\sysmon.db"
    if (Test-Path $dbFile) {
        Write-Host "  Banco: $([math]::Round((Get-Item $dbFile).Length/1KB,1)) KB" -ForegroundColor Gray
    }
    Write-Host ""
}

# ══════════════════════════════════════════════════════════════
#  LOOP PRINCIPAL
# ══════════════════════════════════════════════════════════════
function Start-AgentLoop {
    $installDir = "$env:ProgramData\Microsoft\WinSysMon"
    if (Test-Path $installDir) {
        $script:BasePath   = $installDir
        $script:DbPath     = Join-Path $installDir "sysmon.db"
        $script:ConfigPath = Join-Path $installDir "sysmon-config.json"
        $script:LogPath    = Join-Path $installDir "sysmon.log"
    }
    $script:CachePatternsFile = Join-Path $script:BasePath "patterns-cache.json"

    Write-Log "=== $script:ServiceName v$script:AgentVersion em $env:COMPUTERNAME ==="
    [void](Initialize-Database)  # nao aborta mais se falhar

    $cfg = Load-Config
    $script:PollInterval = $cfg.PollInterval
    if ($cfg.CollectHardware) { try { Collect-MachineInfo } catch {} }

    # Varredura inicial - matar tudo que ja esta rodando e nao deveria (SEM notificar, sweep silencioso)
    try { Enforce-AppBlocking } catch { Write-Log "Sweep inicial falhou: $($_.Exception.Message)" "WARN" }

    # WMI Event - detecta processo INSTANTANEAMENTE ao abrir
    $script:WmiActive = $false
    function Register-ProcWatcher {
        try {
            Get-EventSubscriber -SourceIdentifier "WinSysMon_ProcWatch" -ErrorAction SilentlyContinue | Unregister-Event -ErrorAction SilentlyContinue
            $wmiQuery = "SELECT * FROM __InstanceCreationEvent WITHIN 1 WHERE TargetInstance ISA 'Win32_Process'"
            Register-WmiEvent -Query $wmiQuery -SourceIdentifier "WinSysMon_ProcWatch" -ErrorAction Stop | Out-Null
            $script:WmiActive = $true
            Write-Log "WMI process watcher ativo (deteccao instantanea)"
            return $true
        } catch {
            $script:WmiActive = $false
            Write-Log "WMI watcher falhou (usando polling): $($_.Exception.Message)" "WARN"
            return $false
        }
    }
    Register-ProcWatcher | Out-Null

    $iteration = 0
    $lastHeartbeat = Get-Date
    $consecutiveErrors = 0
    while ($true) {
        try {
            $cfg = Load-Config
            if ($cfg.PollInterval -and $cfg.PollInterval -gt 0) { $script:PollInterval = $cfg.PollInterval }

            # Self-healing WMI: tenta reconectar se morreu
            if (-not $script:WmiActive -or -not (Get-EventSubscriber -SourceIdentifier "WinSysMon_ProcWatch" -ErrorAction SilentlyContinue)) {
                if ($iteration % 30 -eq 0) { Register-ProcWatcher | Out-Null }
            }

            # Processar eventos WMI (instantaneo - processos que acabaram de abrir)
            if ($script:WmiActive) {
                $wmiEvents = @(Get-Event -SourceIdentifier "WinSysMon_ProcWatch" -ErrorAction SilentlyContinue)
                foreach ($evt in $wmiEvents) {
                    try {
                        $newProc = $evt.SourceEventArgs.NewEvent.TargetInstance
                        $procName = ([string]$newProc.Name) -replace '\.exe$',''
                        $procPath = [string]$newProc.ExecutablePath
                        $procPid  = [int]$newProc.ProcessId
                        $procSid  = [int]$newProc.SessionId

                        $blocked = Get-BlockedPatterns
                        if ($blocked.Count -gt 0) {
                            $nameL = ([string]$procName).ToLower()
                            $pathL = if ($procPath) { ([string]$procPath).ToLower() } else { "" }

                            foreach ($pattern in $blocked) {
                                if ($null -eq $pattern) { continue }
                                $p = ([string]$pattern).ToLower().Trim()
                                if (-not $p) { continue }
                                $match = if ($p.Contains("*") -or $p.Contains("?")) {
                                    ($nameL -like $p) -or ($pathL -like $p)
                                } else {
                                    ($nameL -eq $p) -or ($nameL -eq [System.IO.Path]::GetFileNameWithoutExtension($p)) -or ($pathL -eq $p)
                                }
                                if ($match) {
                                    $owner = ""
                                    try { $owner = (Get-CimInstance Win32_Process -Filter "ProcessId=$procPid" -ErrorAction Stop).GetOwner().User } catch {}
                                    try {
                                        Stop-Process -Id $procPid -Force -ErrorAction Stop
                                        Write-Log "BLOQUEADO(WMI): $procName PID=$procPid user=$owner"
                                        Invoke-Sql "INSERT INTO blocked_log (timestamp,hostname,username,process_name,process_path,action) VALUES (@ts,@host,@user,@name,@path,'KILLED')" @{
                                            "@ts"=(Get-Date).ToString("yyyy-MM-dd HH:mm:ss"); "@host"=$env:COMPUTERNAME
                                            "@user"=$owner; "@name"=$procName; "@path"=$procPath
                                        }
                                        Show-BlockedNotification -ProcessName $procName -SessionId $procSid
                                    } catch {
                                        # Retry com taskkill nativo (as vezes contorna Access Denied)
                                        try { & taskkill.exe /F /PID $procPid 2>$null | Out-Null; Write-Log "BLOQUEADO(taskkill): $procName PID=$procPid" } catch {}
                                    }
                                    break
                                }
                            }
                        }
                    } catch { Write-Log "Erro processando evento WMI: $($_.Exception.Message)" "WARN" }
                    Remove-Event -EventIdentifier $evt.EventIdentifier -ErrorAction SilentlyContinue
                }
            }

            # Polling de seguranca toda iteracao (reforco ao WMI)
            Enforce-AppBlocking

            if ($cfg.MonitorLogins -and ($iteration % 6 -eq 0)) { try { Monitor-Logins } catch {} }
            if ($cfg.CollectHardware) {
                if ((New-TimeSpan -Start $script:LastHwCollect -End (Get-Date)).TotalSeconds -ge $cfg.HardwareInterval) {
                    try { Collect-MachineInfo } catch {}
                }
            }

            # Heartbeat a cada 5 minutos (prova que esta vivo)
            if ((New-TimeSpan -Start $lastHeartbeat -End (Get-Date)).TotalMinutes -ge 5) {
                Write-Log "Heartbeat: iter=$iteration wmi=$script:WmiActive patterns=$($script:LastKnownGoodPatterns.Count)"
                $lastHeartbeat = Get-Date
            }

            $consecutiveErrors = 0
        } catch {
            $consecutiveErrors++
            Write-Log "Erro loop (#$consecutiveErrors): $($_.Exception.Message)" "ERROR"
            # Backoff exponencial simples em erros repetidos
            if ($consecutiveErrors -ge 10) {
                Write-Log "10 erros consecutivos - pausa de 30s" "ERROR"
                Start-Sleep -Seconds 30
                $consecutiveErrors = 0
            }
        }
        $iteration++
        Start-Sleep -Seconds $script:PollInterval
    }
}

# ══════════════════════════════════════════════════════════════
#  ENTRY POINT
# ══════════════════════════════════════════════════════════════
if ($Install)        { Install-Agent }
elseif ($Uninstall)  { Uninstall-Agent }
elseif ($Status)     { Show-Status }
elseif ($RunLoop)    { Start-AgentLoop }
else {
    Write-Host "`n=== $script:DisplayName v$script:AgentVersion ===" -ForegroundColor Cyan
    Write-Host "  -Install     Instalar servico" -ForegroundColor Gray
    Write-Host "  -Uninstall   Remover servico" -ForegroundColor Gray
    Write-Host "  -Status      Ver estado" -ForegroundColor Gray
    Write-Host "  -RunLoop     Executar em primeiro plano (teste)`n" -ForegroundColor Gray
}
