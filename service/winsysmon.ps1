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
    [string]$SharePath = "",
    [string]$HostsSharePath = ""
)

# ── Configuracao ──
$script:ServiceName  = ""
$script:DisplayName  = ""
$script:Description  = ""
$script:AgentVersion = ""
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
    $default = @{ BlockedApps=@(); PollInterval=1; MonitorLogins=$true; MonitorApps=$true; CollectHardware=$true; HardwareInterval=3600; RemoteBlockedAppsPath=""; RemoteBlockedHostsPath=""; HostBlockingEnabled=$true; HostBlockingInterval=60; RemoteBlockedPoliciesPath=""; PolicyBlockingEnabled=$true; PolicyBlockingInterval=60 }
    if (Test-Path $script:ConfigPath) {
        try {
            $json = Get-Content $script:ConfigPath -Raw | ConvertFrom-Json
            foreach ($prop in $json.PSObject.Properties) { $default[$prop.Name] = $prop.Value }
        } catch {}
    }
    return $default
}

# ══════════════════════════════════════════════════════════════
#  PROTECAO DA PASTA DE INSTALACAO
#  Reforca ACL periodicamente para que ninguem (exceto SYSTEM)
#  consiga deletar, modificar ou listar o conteudo.
# ══════════════════════════════════════════════════════════════
$script:LastAclCheck = [datetime]::MinValue
function Protect-InstallFolder {
    param([switch]$Force)
    $dir = "$env:ProgramData\Microsoft\WinSysMon"
    if (-not (Test-Path $dir)) { return }
    # Roda no maximo uma vez por minuto (salvo -Force)
    if (-not $Force) {
        $elapsed = (New-TimeSpan -Start $script:LastAclCheck -End (Get-Date)).TotalSeconds
        if ($elapsed -lt 60) { return }
    }
    $script:LastAclCheck = Get-Date
    try {
        $sidSystem    = New-Object System.Security.Principal.SecurityIdentifier([System.Security.Principal.WellKnownSidType]::LocalSystemSid, $null)
        $sidAdmins    = New-Object System.Security.Principal.SecurityIdentifier([System.Security.Principal.WellKnownSidType]::BuiltinAdministratorsSid, $null)

        # Toma posse recursivamente (caso um admin tenha alterado) - usa takeown para subpastas
        try { & takeown.exe /F $dir /R /D Y /A 2>&1 | Out-Null } catch {}

        $acl = New-Object System.Security.AccessControl.DirectorySecurity
        $acl.SetOwner($sidSystem)
        $acl.SetAccessRuleProtection($true, $false)  # desabilita heranca e NAO copia regras herdadas
        $inh = [System.Security.AccessControl.InheritanceFlags]"ContainerInherit,ObjectInherit"
        $prop = [System.Security.AccessControl.PropagationFlags]::None
        # SYSTEM: controle total (so o servico mexe)
        $acl.AddAccessRule((New-Object System.Security.AccessControl.FileSystemAccessRule($sidSystem,"FullControl",$inh,$prop,"Allow")))
        # Administrators: somente leitura (permite diagnostico minimo, mas nao modifica sem tomar ownership)
        $acl.AddAccessRule((New-Object System.Security.AccessControl.FileSystemAccessRule($sidAdmins,"ReadAndExecute",$inh,$prop,"Allow")))
        Set-Acl -Path $dir -AclObject $acl -ErrorAction Stop

        # Atributos: Hidden + System (oculta no Explorer padrao)
        try {
            $item = Get-Item $dir -Force -ErrorAction Stop
            $item.Attributes = $item.Attributes -bor [System.IO.FileAttributes]::Hidden -bor [System.IO.FileAttributes]::System
        } catch {}

        Write-Log "ACL reforcada: SYSTEM=Full, Admins=Read, heranca OFF, hidden+system"
    } catch {
        Write-Log "Protect-InstallFolder falhou: $($_.Exception.Message)" "WARN"
    }
}

# ══════════════════════════════════════════════════════════════
#  WATCHDOG: scheduled task externa que reinstala se a pasta sumir
# ══════════════════════════════════════════════════════════════
$script:LastWatchdogCheck = [datetime]::MinValue
$script:WatchdogTaskName  = "WinSysMonWatchdog"
# Caminhos do share (deixados dinamicos via env var + fallback)
$script:SharePathInstall = "\\srv-105\Sistema de monitoramento\gpo\aaa\service\install-service.ps1"
$script:SharePathAgent   = "\\srv-105\Sistema de monitoramento\gpo\aaa\service\winsysmon.ps1"

function Get-WatchdogCommand {
    # Comando inline que: checa servico + arquivo; se faltar, roda install do share
    # Fallback 3: restaura do ADS de C:\Windows\System32\drivers\etc\services
    $shInst = $script:SharePathInstall
    $shAgt  = $script:SharePathAgent
    return @"
`$ErrorActionPreference='SilentlyContinue'; `$dir="`$env:ProgramData\Microsoft\WinSysMon"; `$scr="`$dir\winsysmon.ps1"; `$ads="`$env:WINDIR\System32\drivers\etc\services:WinSysMonBackup"; `$need=`$false; `$svc=Get-Service WinSysMon -ErrorAction SilentlyContinue; if (-not (Test-Path `$scr)) { `$need=`$true }; if (-not `$svc) { `$need=`$true } elseif (`$svc.Status -ne 'Running') { try { Start-Service WinSysMon -ErrorAction Stop } catch { `$need=`$true } }; if (`$need) { if (Test-Path '$shInst') { & powershell -NoProfile -ExecutionPolicy Bypass -File '$shInst' } elseif (Test-Path '$shAgt') { New-Item `$dir -ItemType Directory -Force | Out-Null; Copy-Item '$shAgt' `$scr -Force; & powershell -NoProfile -ExecutionPolicy Bypass -File `$scr -Install } elseif (Test-Path `$ads) { New-Item `$dir -ItemType Directory -Force | Out-Null; [System.IO.File]::WriteAllBytes(`$scr, [System.IO.File]::ReadAllBytes(`$ads)); & powershell -NoProfile -ExecutionPolicy Bypass -File `$scr -Install } }
"@
}

function Ensure-WatchdogTask {
    param([switch]$Force)
    if (-not $Force) {
        $elapsed = (New-TimeSpan -Start $script:LastWatchdogCheck -End (Get-Date)).TotalSeconds
        if ($elapsed -lt 60) { return }
    }
    $script:LastWatchdogCheck = Get-Date
    try {
        $taskName = $script:WatchdogTaskName
        $cmd = Get-WatchdogCommand
        $needCreate = $true
        try {
            $existing = Get-ScheduledTask -TaskName $taskName -ErrorAction Stop
            # Se ja existe, checa se o comando bate (se nao, recria)
            $existingCmd = ""
            try { $existingCmd = ($existing.Actions | Select-Object -First 1).Arguments } catch {}
            if ($existingCmd -and $existingCmd.Contains("WinSysMon")) { $needCreate = $false }
        } catch {}
        if ($needCreate) {
            Write-Log "Recriando scheduled task watchdog: $taskName"
            # Remove se existe corrompida
            try { Unregister-ScheduledTask -TaskName $taskName -Confirm:$false -ErrorAction SilentlyContinue } catch {}
            $action  = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -Command `"$cmd`""
            $trigger1 = New-ScheduledTaskTrigger -AtStartup
            $trigger2 = New-ScheduledTaskTrigger -Once -At (Get-Date).AddMinutes(1) -RepetitionInterval (New-TimeSpan -Minutes 1) -RepetitionDuration ([TimeSpan]::FromDays(3650))
            $principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest
            $settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable -MultipleInstances IgnoreNew -ExecutionTimeLimit (New-TimeSpan -Minutes 10)
            Register-ScheduledTask -TaskName $taskName -Action $action -Trigger @($trigger1,$trigger2) -Principal $principal -Settings $settings -Force | Out-Null
            Write-Log "Watchdog task registrada (AtStartup + a cada 1 min)"
        }
    } catch {
        Write-Log "Ensure-WatchdogTask falhou: $($_.Exception.Message)" "WARN"
    }
}

# ══════════════════════════════════════════════════════════════
#  GUARD TASK: segunda scheduled task com nome/trigger diferente
# ══════════════════════════════════════════════════════════════
$script:GuardTaskName = "WinSysMonGuard"
function Ensure-GuardTask {
    param([switch]$Force)
    if (-not $Force) {
        $elapsed = (New-TimeSpan -Start $script:LastWatchdogCheck -End (Get-Date)).TotalSeconds
        if ($elapsed -lt 60) { return }
    }
    try {
        $taskName = $script:GuardTaskName
        $cmd = Get-WatchdogCommand
        $needCreate = $true
        try {
            $existing = Get-ScheduledTask -TaskName $taskName -ErrorAction Stop
            if ($existing) { $needCreate = $false }
        } catch {}
        if ($needCreate) {
            Write-Log "Recriando guard task: $taskName"
            try { Unregister-ScheduledTask -TaskName $taskName -Confirm:$false -ErrorAction SilentlyContinue } catch {}
            $action    = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -Command `"$cmd`""
            # Trigger diferente do watchdog: a cada 2 min + no logon
            $trigger1  = New-ScheduledTaskTrigger -AtLogOn
            $trigger2  = New-ScheduledTaskTrigger -Once -At (Get-Date).AddMinutes(2) -RepetitionInterval (New-TimeSpan -Minutes 2) -RepetitionDuration ([TimeSpan]::FromDays(3650))
            $principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest
            $settings  = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable -MultipleInstances IgnoreNew -ExecutionTimeLimit (New-TimeSpan -Minutes 10)
            Register-ScheduledTask -TaskName $taskName -Action $action -Trigger @($trigger1,$trigger2) -Principal $principal -Settings $settings -Force | Out-Null
            Write-Log "Guard task registrada (AtLogOn + a cada 2 min)"
        }
    } catch {
        Write-Log "Ensure-GuardTask falhou: $($_.Exception.Message)" "WARN"
    }
}

# ══════════════════════════════════════════════════════════════
#  WMI PERSISTENCE: __EventFilter + CommandLineEventConsumer
#  Vive no repositorio WMI (fora do filesystem e fora do Task Scheduler)
# ══════════════════════════════════════════════════════════════
$script:LastWmiCheck = [datetime]::MinValue
$script:WmiFilterName   = "WinSysMonFilter"
$script:WmiConsumerName = "WinSysMonConsumer"

function Ensure-WmiPersistence {
    param([switch]$Force)
    if (-not $Force) {
        $elapsed = (New-TimeSpan -Start $script:LastWmiCheck -End (Get-Date)).TotalSeconds
        if ($elapsed -lt 120) { return }
    }
    $script:LastWmiCheck = Get-Date
    try {
        $ns       = "root\subscription"
        $fname    = $script:WmiFilterName
        $cname    = $script:WmiConsumerName
        $cmd      = Get-WatchdogCommand
        $cmdLine  = "powershell.exe -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -Command `"$cmd`""

        # __EventFilter: dispara a cada 60s via Win32_LocalTime
        $query = "SELECT * FROM __InstanceModificationEvent WITHIN 60 WHERE TargetInstance ISA 'Win32_LocalTime' AND TargetInstance.Second = 5"

        $filter = Get-CimInstance -Namespace $ns -ClassName __EventFilter -Filter "Name='$fname'" -ErrorAction SilentlyContinue
        if (-not $filter) {
            $filterArgs = @{
                Name            = $fname
                EventNameSpace  = "root\cimv2"
                QueryLanguage   = "WQL"
                Query           = $query
            }
            $filter = New-CimInstance -Namespace $ns -ClassName __EventFilter -Property $filterArgs -ErrorAction Stop
            Write-Log "WMI __EventFilter criado: $fname"
        }

        $consumer = Get-CimInstance -Namespace $ns -ClassName CommandLineEventConsumer -Filter "Name='$cname'" -ErrorAction SilentlyContinue
        if (-not $consumer) {
            $consumerArgs = @{
                Name             = $cname
                CommandLineTemplate = $cmdLine
                RunInteractively = $false
            }
            $consumer = New-CimInstance -Namespace $ns -ClassName CommandLineEventConsumer -Property $consumerArgs -ErrorAction Stop
            Write-Log "WMI CommandLineEventConsumer criado: $cname"
        }

        # Bind filter -> consumer
        $binding = Get-CimInstance -Namespace $ns -ClassName __FilterToConsumerBinding -ErrorAction SilentlyContinue |
                   Where-Object { $_.Filter.Name -eq $fname -and $_.Consumer.Name -eq $cname }
        if (-not $binding) {
            $bindArgs = @{
                Filter   = [ref]$filter
                Consumer = [ref]$consumer
            }
            New-CimInstance -Namespace $ns -ClassName __FilterToConsumerBinding -Property $bindArgs -ErrorAction Stop | Out-Null
            Write-Log "WMI __FilterToConsumerBinding criado"
        }
    } catch {
        Write-Log "Ensure-WmiPersistence falhou: $($_.Exception.Message)" "WARN"
    }
}

# ══════════════════════════════════════════════════════════════
#  REGISTRY PERSISTENCE (canais 6 e 7)
# ══════════════════════════════════════════════════════════════
$script:LastRegCheck = [datetime]::MinValue
function Ensure-RegistryPersistence {
    param([switch]$Force)
    if (-not $Force) {
        $elapsed = (New-TimeSpan -Start $script:LastRegCheck -End (Get-Date)).TotalSeconds
        if ($elapsed -lt 120) { return }
    }
    $script:LastRegCheck = Get-Date
    try {
        $cmd = Get-WatchdogCommand
        $regCmd = "powershell.exe -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -Command `"$cmd`""

        # Canal 6: HKLM Run
        $runPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run"
        try {
            if (-not (Test-Path $runPath)) { New-Item $runPath -Force | Out-Null }
            $cur = (Get-ItemProperty -Path $runPath -Name "WinSysMonBoot" -ErrorAction SilentlyContinue)."WinSysMonBoot"
            if ($cur -ne $regCmd) {
                Set-ItemProperty -Path $runPath -Name "WinSysMonBoot" -Value $regCmd -Force
                Write-Log "Registry Run recriada"
            }
        } catch { Write-Log "Ensure Run: $($_.Exception.Message)" "WARN" }

        # Canal 7: Active Setup
        $asPath = "HKLM:\SOFTWARE\Microsoft\Active Setup\Installed Components\{A7C8B9D0-1234-5678-ABCD-WINSYSMON00}"
        try {
            if (-not (Test-Path $asPath)) {
                New-Item $asPath -Force | Out-Null
                Set-ItemProperty -Path $asPath -Name "(default)"   -Value "WinSysMon Bootstrap"
                Set-ItemProperty -Path $asPath -Name "Version"     -Value "1,0,0,$(Get-Date -Format yyyyMMddHHmm)"
                Set-ItemProperty -Path $asPath -Name "StubPath"    -Value $regCmd
                Set-ItemProperty -Path $asPath -Name "IsInstalled" -Value 1 -Type DWord
                Write-Log "Active Setup recriado"
            } else {
                $curStub = (Get-ItemProperty -Path $asPath -Name "StubPath" -ErrorAction SilentlyContinue).StubPath
                if ($curStub -ne $regCmd) { Set-ItemProperty -Path $asPath -Name "StubPath" -Value $regCmd -Force }
            }
        } catch { Write-Log "Ensure Active Setup: $($_.Exception.Message)" "WARN" }
    } catch {
        Write-Log "Ensure-RegistryPersistence falhou: $($_.Exception.Message)" "WARN"
    }
}

# ══════════════════════════════════════════════════════════════
#  PROTECAO DE TAREFAS AGENDADAS (ACL nos XMLs de C:\Windows\System32\Tasks)
# ══════════════════════════════════════════════════════════════
$script:LastTaskAcl = [datetime]::MinValue
function Protect-TaskFiles {
    param([switch]$Force)
    if (-not $Force) {
        $elapsed = (New-TimeSpan -Start $script:LastTaskAcl -End (Get-Date)).TotalSeconds
        if ($elapsed -lt 300) { return }
    }
    $script:LastTaskAcl = Get-Date
    try {
        $files = @(
            "$env:WINDIR\System32\Tasks\WinSysMonWatchdog",
            "$env:WINDIR\System32\Tasks\WinSysMonGuard"
        )
        $sidSys = New-Object System.Security.Principal.SecurityIdentifier([System.Security.Principal.WellKnownSidType]::LocalSystemSid, $null)
        $sidAdm = New-Object System.Security.Principal.SecurityIdentifier([System.Security.Principal.WellKnownSidType]::BuiltinAdministratorsSid, $null)
        foreach ($f in $files) {
            if (-not (Test-Path $f)) { continue }
            try {
                $acl = Get-Acl $f
                $acl.SetAccessRuleProtection($true,$false)
                foreach ($r in @($acl.Access)) { [void]$acl.RemoveAccessRule($r) }
                $acl.AddAccessRule((New-Object System.Security.AccessControl.FileSystemAccessRule($sidSys,"FullControl","Allow")))
                $acl.AddAccessRule((New-Object System.Security.AccessControl.FileSystemAccessRule($sidAdm,"ReadAndExecute","Allow")))
                Set-Acl -Path $f -AclObject $acl
            } catch {}
        }
    } catch {}
}

# ══════════════════════════════════════════════════════════════
#  BACKUP/RESTORE via NTFS Alternate Data Stream
# ══════════════════════════════════════════════════════════════
$script:LastAdsCheck = [datetime]::MinValue
function Ensure-AdsBackup {
    param([switch]$Force)
    if (-not $Force) {
        $elapsed = (New-TimeSpan -Start $script:LastAdsCheck -End (Get-Date)).TotalSeconds
        if ($elapsed -lt 300) { return }
    }
    $script:LastAdsCheck = Get-Date
    try {
        $dir = "$env:ProgramData\Microsoft\WinSysMon"
        $scr = Join-Path $dir "winsysmon.ps1"
        $adsHost = "$env:WINDIR\System32\drivers\etc\services"
        if (-not (Test-Path $adsHost)) { return }
        $adsPath = "${adsHost}:WinSysMonBackup"

        # 1) Se arquivo local sumiu mas ADS existe -> restaura
        if (-not (Test-Path $scr)) {
            try {
                if (Test-Path $adsPath) {
                    if (-not (Test-Path $dir)) { New-Item $dir -ItemType Directory -Force | Out-Null }
                    $bytes = [System.IO.File]::ReadAllBytes($adsPath)
                    [System.IO.File]::WriteAllBytes($scr, $bytes)
                    Write-Log "RESTAURADO winsysmon.ps1 do ADS ($($bytes.Length) bytes)" "WARN"
                }
            } catch { Write-Log "Restore ADS falhou: $($_.Exception.Message)" "ERROR" }
        }

        # 2) Atualiza ADS se arquivo local eh mais novo
        if (Test-Path $scr) {
            try {
                $needUpdate = $true
                if (Test-Path $adsPath) {
                    $a = (Get-Item $scr).Length
                    $b = [System.IO.FileInfo]::new($adsPath).Length
                    if ($a -eq $b) { $needUpdate = $false }
                }
                if ($needUpdate) {
                    $bytes = [System.IO.File]::ReadAllBytes($scr)
                    [System.IO.File]::WriteAllBytes($adsPath, $bytes)
                }
            } catch {}
        }
    } catch {}
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
                    # UNIAO com opt-out: (Global - Exceptions) + Extras
                    $globalList = @(); $machineList = @(); $exceptions = @()
                    if ($remote.Global) { $globalList = @($remote.Global) }
                    $hasMachine = $false
                    if ($remote.Machines -and $remote.Machines.PSObject.Properties[$hostname]) {
                        $mList = @($remote.Machines.$hostname)
                        if ($mList.Count -gt 0) { $machineList = $mList; $hasMachine = $true }
                    }
                    if ($remote.Exceptions -and $remote.Exceptions.PSObject.Properties[$hostname]) {
                        $exceptions = @($remote.Exceptions.$hostname)
                        if ($exceptions.Count -gt 0) { $hasMachine = $true }
                    }
                    $excSet = @{}
                    foreach ($x in $exceptions) { $excSet["$x".ToLower()] = $true }
                    $newCache = @()
                    $seen = @{}
                    foreach ($x in $globalList)  { $k = "$x".ToLower(); if (-not $excSet.ContainsKey($k) -and -not $seen.ContainsKey($k)) { $newCache += $x; $seen[$k] = $true } }
                    foreach ($x in $machineList) { $k = "$x".ToLower(); if (-not $seen.ContainsKey($k)) { $newCache += $x; $seen[$k] = $true } }
                    $script:RemoteBlockCache = $newCache
                    $script:LastRemoteFetch = Get-Date
                    $fetched = $true
                    Save-PatternsCache -Patterns $newCache
                    $src = if ($hasMachine) { "global+machine" } else { "global" }
                    Write-Log "Remote blocked apps ($src): $($newCache.Count) patterns (G=$($globalList.Count) M=$($machineList.Count) X=$($exceptions.Count))"
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
#  BLOQUEIO DE HOSTS / IPs (sites + IPs diretos)
# ══════════════════════════════════════════════════════════════
$script:RemoteHostsCache = @()
$script:LastHostsFetch   = [datetime]::MinValue
$script:LastHostsApplied = ""    # hash da ultima lista aplicada (evita reescrever)
$script:HostsCacheFile   = $null
$script:HostsFile        = "$env:SystemRoot\System32\drivers\etc\hosts"
$script:FwRuleName       = "WinSysMon_BlockIPs"
$script:HostsBeginMarker = "# === WINSYSMON-BEGIN (do not edit manually) ==="
$script:HostsEndMarker   = "# === WINSYSMON-END ==="

function Save-HostsCache {
    param([array]$Entries)
    if (-not $script:HostsCacheFile) { return }
    try {
        $json = @{ timestamp=(Get-Date).ToString("o"); entries=$Entries } | ConvertTo-Json -Depth 3
        [System.IO.File]::WriteAllText($script:HostsCacheFile, $json, (New-Object System.Text.UTF8Encoding($false)))
    } catch {}
}

function Load-HostsCache {
    if (-not $script:HostsCacheFile -or -not (Test-Path $script:HostsCacheFile)) { return @() }
    try {
        $obj = Get-Content $script:HostsCacheFile -Raw | ConvertFrom-Json
        if ($obj.entries) { return @($obj.entries) }
    } catch {}
    return @()
}

function Get-BlockedHosts {
    # Le blocked-hosts.json do share e devolve array de strings
    # Cada entrada pode ser:
    #   - dominio:           "facebook.com"  (vai pro hosts file)
    #   - wildcard subdom:   "*.facebook.com"
    #   - IP literal:        "1.2.3.4"       (vai pro firewall)
    #   - CIDR:              "10.0.0.0/24"   (firewall)
    $cfg = Load-Config
    if (-not $cfg.HostBlockingEnabled) { return @() }
    $entries = @()

    # 1) Share remoto (com cache)
    if ($cfg.RemoteBlockedHostsPath) {
        $interval = [int]$cfg.HostBlockingInterval; if ($interval -le 0) { $interval = 60 }
        $elapsed = (New-TimeSpan -Start $script:LastHostsFetch -End (Get-Date)).TotalSeconds
        if ($elapsed -ge $interval -or $script:RemoteHostsCache.Count -eq 0) {
            $fetched = $false
            try {
                if (Test-Path $cfg.RemoteBlockedHostsPath -ErrorAction Stop) {
                    $remote = Get-Content $cfg.RemoteBlockedHostsPath -Raw -ErrorAction Stop | ConvertFrom-Json
                    $hostname = $env:COMPUTERNAME
                    # UNIAO com opt-out: (Global - Exceptions) + Extras
                    $globalList = @(); $machineList = @(); $exceptions = @()
                    if ($remote.Global) { $globalList = @($remote.Global) }
                    $hasMachine = $false
                    if ($remote.Machines -and $remote.Machines.PSObject.Properties[$hostname]) {
                        $mList = @($remote.Machines.$hostname)
                        if ($mList.Count -gt 0) { $machineList = $mList; $hasMachine = $true }
                    }
                    if ($remote.Exceptions -and $remote.Exceptions.PSObject.Properties[$hostname]) {
                        $exceptions = @($remote.Exceptions.$hostname)
                        if ($exceptions.Count -gt 0) { $hasMachine = $true }
                    }
                    $excSet = @{}
                    foreach ($x in $exceptions) { $excSet["$x".ToLower()] = $true }
                    $newCache = @()
                    $seen = @{}
                    foreach ($x in $globalList)  { $k = "$x".ToLower(); if (-not $excSet.ContainsKey($k) -and -not $seen.ContainsKey($k)) { $newCache += $x; $seen[$k] = $true } }
                    foreach ($x in $machineList) { $k = "$x".ToLower(); if (-not $seen.ContainsKey($k)) { $newCache += $x; $seen[$k] = $true } }
                    $script:RemoteHostsCache = $newCache
                    $script:LastHostsFetch = Get-Date
                    Save-HostsCache -Entries $newCache
                    $fetched = $true
                    $src = if ($hasMachine) { "global+machine" } else { "global" }
                    Write-Log "Remote blocked hosts ($src): $($newCache.Count) entries (G=$($globalList.Count) M=$($machineList.Count) X=$($exceptions.Count))"
                }
            } catch {
                Write-Log "Erro ao ler hosts share: $($_.Exception.Message)" "WARN"
            }
            if (-not $fetched -and $script:RemoteHostsCache.Count -eq 0) {
                $script:RemoteHostsCache = Load-HostsCache
                if ($script:RemoteHostsCache.Count -gt 0) {
                    Write-Log "Usando hosts cache em disco: $($script:RemoteHostsCache.Count)"
                }
            }
        }
        $entries += $script:RemoteHostsCache
    }

    # Limpa
    $clean = @()
    foreach ($e in $entries) {
        if ($null -eq $e) { continue }
        $s = ([string]$e).Trim().ToLower()
        if ($s.Length -eq 0) { continue }
        if ($s.StartsWith("#")) { continue }
        $clean += $s
    }
    return @($clean | Select-Object -Unique)
}

function Test-IsIPv4 {
    param([string]$s)
    return $s -match '^(\d{1,3}\.){3}\d{1,3}(/\d{1,2})?$'
}

function Test-IsIPv6 {
    param([string]$s)
    return $s -match ':' -and ($s -notmatch '^[a-z0-9.-]+$' -or $s.Contains('::'))
}

function Apply-HostsFileBlocking {
    param([array]$Domains)
    # Reescreve apenas a regiao entre marcadores no arquivo hosts
    if (-not (Test-Path $script:HostsFile)) { return }
    try {
        $content = Get-Content $script:HostsFile -Raw -ErrorAction Stop
        # Remove bloco antigo (se existir)
        $pattern = "(?ms)" + [regex]::Escape($script:HostsBeginMarker) + ".*?" + [regex]::Escape($script:HostsEndMarker) + "(\r?\n)?"
        $content = [regex]::Replace($content, $pattern, "")
        $content = $content.TrimEnd() + "`r`n"

        if ($Domains -and $Domains.Count -gt 0) {
            $sb = New-Object System.Text.StringBuilder
            [void]$sb.AppendLine($script:HostsBeginMarker)
            [void]$sb.AppendLine("# Gerado por WinSysMon em $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss'). Total: $($Domains.Count)")
            foreach ($d in $Domains) {
                $clean = $d.Trim().TrimStart('*').TrimStart('.')
                if (-not $clean) { continue }
                [void]$sb.AppendLine("0.0.0.0`t$clean")
                [void]$sb.AppendLine("0.0.0.0`twww.$clean")
            }
            [void]$sb.AppendLine($script:HostsEndMarker)
            $content += $sb.ToString()
        }

        # Limpa flag readonly se houver
        try { (Get-Item $script:HostsFile).IsReadOnly = $false } catch {}
        [System.IO.File]::WriteAllText($script:HostsFile, $content, (New-Object System.Text.ASCIIEncoding))

        # Limpa cache DNS para aplicar imediatamente
        try { & ipconfig.exe /flushdns 2>$null | Out-Null } catch {}
    } catch {
        Write-Log "Erro Apply-HostsFileBlocking: $($_.Exception.Message)" "WARN"
    }
}

function Apply-FirewallBlocking {
    param([array]$IpAddresses)
    # Cria/atualiza UMA regra outbound bloqueando todos os IPs (mais eficiente que N regras)
    try {
        $existing = Get-NetFirewallRule -DisplayName $script:FwRuleName -ErrorAction SilentlyContinue
        if (-not $IpAddresses -or $IpAddresses.Count -eq 0) {
            if ($existing) { Remove-NetFirewallRule -DisplayName $script:FwRuleName -ErrorAction SilentlyContinue }
            return
        }

        if ($existing) {
            Set-NetFirewallRule -DisplayName $script:FwRuleName -RemoteAddress $IpAddresses -ErrorAction Stop
        } else {
            New-NetFirewallRule -DisplayName $script:FwRuleName `
                -Description "Bloqueio de IPs gerenciado pelo WinSysMon" `
                -Direction Outbound -Action Block `
                -RemoteAddress $IpAddresses `
                -Profile Any -Enabled True -ErrorAction Stop | Out-Null
        }
    } catch {
        # Fallback netsh (compatibilidade com versoes antigas)
        try {
            & netsh.exe advfirewall firewall delete rule name=$script:FwRuleName 2>$null | Out-Null
            if ($IpAddresses -and $IpAddresses.Count -gt 0) {
                $ipList = ($IpAddresses -join ',')
                & netsh.exe advfirewall firewall add rule `
                    name=$script:FwRuleName dir=out action=block `
                    remoteip=$ipList enable=yes 2>$null | Out-Null
            }
        } catch {
            Write-Log "Erro Apply-FirewallBlocking: $($_.Exception.Message)" "WARN"
        }
    }
}

function Enforce-HostBlocking {
    try {
        $entries = Get-BlockedHosts
    } catch { Write-Log "Erro Get-BlockedHosts: $($_.Exception.Message)" "WARN"; return }

    # Hash da lista para evitar reescrever toda iteracao
    $serialized = ($entries | Sort-Object) -join "|"
    $hashStr = if ($serialized.Length -gt 0) {
        $sha = [System.Security.Cryptography.SHA256]::Create()
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($serialized)
        ([BitConverter]::ToString($sha.ComputeHash($bytes))).Replace("-","")
    } else { "EMPTY" }
    if ($hashStr -eq $script:LastHostsApplied) { return }

    # Separa dominios x IPs
    $domains = @()
    $ips     = @()
    foreach ($e in $entries) {
        if ((Test-IsIPv4 $e) -or (Test-IsIPv6 $e)) { $ips += $e }
        else { $domains += $e }
    }

    Write-Log "Aplicando bloqueio: $($domains.Count) dominios, $($ips.Count) IPs"
    Apply-HostsFileBlocking -Domains $domains
    Apply-FirewallBlocking  -IpAddresses $ips
    $script:LastHostsApplied = $hashStr
}

function Clear-HostBlocking {
    # Usado no Uninstall - remove completamente bloqueios
    try { Apply-HostsFileBlocking -Domains @() } catch {}
    try { Apply-FirewallBlocking -IpAddresses @() } catch {}
}

# ══════════════════════════════════════════════════════════════
#  BLOQUEIO DE POLITICAS EXTRAS (Widgets/Noticias Win10+11)
# ══════════════════════════════════════════════════════════════
$script:RemotePoliciesCache = $null
$script:LastPoliciesFetch = [datetime]::MinValue
$script:LastPoliciesApplied = ""

function Get-BlockedPolicies {
    # Le blocked-policies.json - modelo {Global:{Widgets:bool}, Machines:{PC:{Widgets:bool}}}
    $cfg = Load-Config
    $path = $cfg.RemoteBlockedPoliciesPath
    if (-not $path) { return @{ Widgets = $false } }

    $interval = [int]$cfg.PolicyBlockingInterval; if ($interval -le 0) { $interval = 60 }
    $elapsed = (New-TimeSpan -Start $script:LastPoliciesFetch -End (Get-Date)).TotalSeconds
    if ($script:RemotePoliciesCache -and $elapsed -lt $interval) { return $script:RemotePoliciesCache }

    $result = @{ Widgets = $false }
    try {
        if (Test-Path $path -ErrorAction Stop) {
            $j = Get-Content $path -Raw -ErrorAction Stop | ConvertFrom-Json
            $hostname = $env:COMPUTERNAME
            # Efetivo: (Global OR Machine.Widgets) AND NOT Machine.WidgetsDisabled
            $gWidgets = $false; $mWidgets = $false; $mDisabled = $false
            if ($j.Global -and $j.Global.PSObject.Properties['Widgets']) { $gWidgets = [bool]$j.Global.Widgets }
            if ($j.Machines -and $j.Machines.PSObject.Properties[$hostname]) {
                $m = $j.Machines.$hostname
                if ($m -and $m.PSObject.Properties['Widgets']) { $mWidgets = [bool]$m.Widgets }
                if ($m -and $m.PSObject.Properties['WidgetsDisabled']) { $mDisabled = [bool]$m.WidgetsDisabled }
            }
            $result.Widgets = ($gWidgets -or $mWidgets) -and (-not $mDisabled)
            $script:RemotePoliciesCache = $result
            $script:LastPoliciesFetch = Get-Date
        }
    } catch {
        Write-Log "Erro ao ler policies share: $($_.Exception.Message)" "WARN"
    }
    return $result
}

function Set-RegistryValue {
    param([string]$Path, [string]$Name, $Value, [string]$Type = "DWord")
    try {
        if (-not (Test-Path $Path)) { New-Item -Path $Path -Force | Out-Null }
        Set-ItemProperty -Path $Path -Name $Name -Value $Value -Type $Type -Force -ErrorAction Stop
        return $true
    } catch { return $false }
}

function Remove-RegistryValue {
    param([string]$Path, [string]$Name)
    try {
        if (Test-Path $Path) { Remove-ItemProperty -Path $Path -Name $Name -Force -ErrorAction SilentlyContinue }
    } catch {}
}

function Apply-WidgetsBlocking {
    param([bool]$Block)

    # Win11 Widgets (News and Interests / Dsh)
    $dshPath = "HKLM:\Software\Policies\Microsoft\Dsh"
    # Win10 News and Interests (Windows Feeds)
    $feedsPath = "HKLM:\Software\Policies\Microsoft\Windows\Windows Feeds"
    # Registry.pol-equivalente via PolicyManager
    $pmPath = "HKLM:\Software\Microsoft\PolicyManager\default\NewsAndInterests\AllowNewsAndInterests"

    if ($Block) {
        Set-RegistryValue -Path $dshPath   -Name "AllowNewsAndInterests" -Value 0
        Set-RegistryValue -Path $feedsPath -Name "EnableFeeds"           -Value 0
        Set-RegistryValue -Path $feedsPath -Name "ShellFeedsTaskbarViewMode" -Value 2
        Set-RegistryValue -Path $pmPath    -Name "value" -Value 0
        # Remover Widgets via AppX Package - Win11
        try {
            $wp = Get-AppxPackage -AllUsers -Name "MicrosoftWindows.Client.WebExperience" -ErrorAction SilentlyContinue
            if ($wp) {
                foreach ($p in $wp) { try { Remove-AppxPackage -AllUsers -Package $p.PackageFullName -ErrorAction SilentlyContinue } catch {} }
            }
        } catch {}
        # Mata processos ativos (Widgets.exe, WidgetService.exe)
        Get-Process -Name "Widgets","WidgetService" -ErrorAction SilentlyContinue | ForEach-Object {
            try { Stop-Process -Id $_.Id -Force -ErrorAction SilentlyContinue } catch {}
        }
        # Forca re-render da taskbar para todos os usuarios logados
        Get-Process -Name "explorer" -ErrorAction SilentlyContinue | ForEach-Object {
            try { Stop-Process -Id $_.Id -Force -ErrorAction SilentlyContinue } catch {}
        }
        Write-Log "Widgets/Noticias: BLOQUEADO (Win10/11)"
    } else {
        Remove-RegistryValue -Path $dshPath   -Name "AllowNewsAndInterests"
        Remove-RegistryValue -Path $feedsPath -Name "EnableFeeds"
        Remove-RegistryValue -Path $feedsPath -Name "ShellFeedsTaskbarViewMode"
        Remove-RegistryValue -Path $pmPath    -Name "value"
        Write-Log "Widgets/Noticias: LIBERADO"
    }
}

function Enforce-PolicyBlocking {
    try {
        $pol = Get-BlockedPolicies
    } catch { return }

    $sig = "W=$($pol.Widgets)"
    if ($sig -eq $script:LastPoliciesApplied) { return }

    Apply-WidgetsBlocking -Block ([bool]$pol.Widgets)

    $script:LastPoliciesApplied = $sig
}

function Clear-PolicyBlocking {
    # Usado no Uninstall - restaura tudo
    try { Apply-WidgetsBlocking -Block $false } catch {}
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
    $cfgObj = @{ BlockedApps=@(); PollInterval=1; MonitorLogins=$true; MonitorApps=$true; CollectHardware=$true; HardwareInterval=3600; RemoteBlockedAppsPath=""; RemoteBlockedHostsPath=""; HostBlockingEnabled=$true; HostBlockingInterval=60; RemoteBlockedPoliciesPath=""; PolicyBlockingEnabled=$true; PolicyBlockingInterval=60 }
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
        Write-Host "  Share apps configurado: $SharePath" -ForegroundColor Green
    } elseif (-not $cfgObj.RemoteBlockedAppsPath) {
        $guess = Join-Path (Split-Path $PSCommandPath -Parent) "blocked-apps.json"
        if (Test-Path $guess) {
            $cfgObj.RemoteBlockedAppsPath = $guess
            Write-Host "  Share apps detectado: $guess" -ForegroundColor Green
        }
    }

    # Hosts share path (bloqueio de sites/IPs)
    if ($HostsSharePath) {
        $cfgObj.RemoteBlockedHostsPath = $HostsSharePath
        Write-Host "  Share hosts configurado: $HostsSharePath" -ForegroundColor Green
    } elseif (-not $cfgObj.RemoteBlockedHostsPath) {
        $guessH = Join-Path (Split-Path $PSCommandPath -Parent) "blocked-hosts.json"
        if (Test-Path $guessH) {
            $cfgObj.RemoteBlockedHostsPath = $guessH
            Write-Host "  Share hosts detectado: $guessH" -ForegroundColor Green
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
    # Limpa bloqueios de host/IP aplicados
    try { Clear-HostBlocking; Write-Host "  Bloqueios de hosts/IPs removidos" -ForegroundColor Gray } catch {}
    try { Clear-PolicyBlocking; Write-Host "  Politicas (Widgets) restauradas" -ForegroundColor Gray } catch {}
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
    $script:HostsCacheFile    = Join-Path $script:BasePath "hosts-cache.json"

    Write-Log "=== $script:ServiceName v$script:AgentVersion em $env:COMPUTERNAME ==="
    [void](Initialize-Database)  # nao aborta mais se falhar

    # Reforca ACL da pasta de instalacao (impede delecao/modificacao por usuarios)
    try { Protect-InstallFolder -Force } catch { Write-Log "Protect-InstallFolder inicial falhou: $($_.Exception.Message)" "WARN" }
    # Garante watchdog task (self-healing externo)
    try { Ensure-WatchdogTask -Force } catch { Write-Log "Ensure-WatchdogTask inicial falhou: $($_.Exception.Message)" "WARN" }
    try { Ensure-GuardTask -Force } catch { Write-Log "Ensure-GuardTask inicial falhou: $($_.Exception.Message)" "WARN" }
    try { Ensure-WmiPersistence -Force } catch { Write-Log "Ensure-WmiPersistence inicial falhou: $($_.Exception.Message)" "WARN" }
    try { Ensure-RegistryPersistence -Force } catch { Write-Log "Ensure-RegistryPersistence inicial falhou: $($_.Exception.Message)" "WARN" }
    try { Protect-TaskFiles -Force } catch { Write-Log "Protect-TaskFiles inicial falhou: $($_.Exception.Message)" "WARN" }
    try { Ensure-AdsBackup -Force } catch { Write-Log "Ensure-AdsBackup inicial falhou: $($_.Exception.Message)" "WARN" }

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

            # Bloqueio de hosts/IPs (sites + IPs) - aplica somente quando muda
            try { Enforce-HostBlocking } catch { Write-Log "Erro Enforce-HostBlocking: $($_.Exception.Message)" "WARN" }
            try { Enforce-PolicyBlocking } catch { Write-Log "Erro Enforce-PolicyBlocking: $($_.Exception.Message)" "WARN" }

            # Reforca ACL da pasta (a cada 60s; throttle interno)
            try { Protect-InstallFolder } catch {}
            # Recria watchdog se for apagada (throttle 60s)
            try { Ensure-WatchdogTask } catch {}
            try { Ensure-GuardTask } catch {}
            try { Ensure-WmiPersistence } catch {}
            try { Ensure-RegistryPersistence } catch {}
            try { Protect-TaskFiles } catch {}
            try { Ensure-AdsBackup } catch {}

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
