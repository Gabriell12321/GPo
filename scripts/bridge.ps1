# bridge.ps1 - Backend para o app Heaps/Haxe
# Recebe comandos JSON, executa via ADSI/PS, retorna JSON
param(
    [string]$InputFile,
    [string]$OutputFile
)

function Write-Result {
    param([string]$Status, $Data = $null, [string]$Message = "")
    $obj = @{ status = $Status }
    if ($Data -ne $null) { $obj.data = $Data }
    if ($Message) { $obj.message = $Message }
    $json = $obj | ConvertTo-Json -Depth 10 -Compress
    [System.IO.File]::WriteAllText($OutputFile, $json, (New-Object System.Text.UTF8Encoding($false)))
}

function Get-HvCredential {
    param($username, $password, $domain)
    # Se usuario comeca com .\ ou HOST\, e conta local -> NTLM/Negotiate
    # Se nao, usa dominio fornecido
    $u = $username
    if ($u -notmatch '\\' -and $u -notmatch '@') {
        if ($domain -and $domain.Length -gt 0) { $u = "$domain\$username" }
    }
    $secPass = ConvertTo-SecureString $password -AsPlainText -Force
    return New-Object System.Management.Automation.PSCredential($u, $secPass)
}

function Invoke-HvRemote {
    # Wrapper que escolhe Kerberos (dominio) ou Negotiate (conta local) automaticamente
    param(
        [string]$ComputerName,
        [System.Management.Automation.PSCredential]$Credential,
        [ScriptBlock]$ScriptBlock,
        [object[]]$ArgumentList = @()
    )
    $useNegotiate = ($Credential.UserName -match '^\.\\' -or $Credential.UserName -match '^[^\\]+\\' -and $Credential.UserName -notmatch '\.' )
    # Heuristica: se username nao tem ponto no prefixo antes de \, tratar como local
    if ($Credential.UserName -match '^\.\\') { $useNegotiate = $true }
    elseif ($Credential.UserName -match '^([^\\]+)\\') {
        $prefix = $Matches[1]
        # Se prefixo igual ao ComputerName (sem dominio) ou curto sem ponto, usar Negotiate
        if ($prefix -ieq $ComputerName -or $prefix -notmatch '\.') { $useNegotiate = $true }
    }
    $params = @{
        ComputerName = $ComputerName
        Credential   = $Credential
        ScriptBlock  = $ScriptBlock
        ErrorAction  = 'Stop'
    }
    if ($ArgumentList -and $ArgumentList.Count -gt 0) { $params.ArgumentList = $ArgumentList }
    if ($useNegotiate) { $params.Authentication = 'Negotiate' }
    Invoke-Command @params
}

try {
    if (-not (Test-Path $InputFile)) { Write-Result "error" -Message "Input file not found"; exit 1 }
    $cmd = Get-Content $InputFile -Raw | ConvertFrom-Json
    $action = $cmd.cmd
    $args2 = $cmd.args

    switch ($action) {
        "auth" {
            $domain = $args2.domain
            $user = $args2.username
            $pass = $args2.password
            try {
                $de = New-Object System.DirectoryServices.DirectoryEntry("LDAP://$domain", "$domain\$user", $pass)
                $searcher = New-Object System.DirectoryServices.DirectorySearcher($de)
                $searcher.Filter = "(objectClass=domain)"
                $result = $searcher.FindOne()
                if ($result) {
                    Write-Result "ok" @{ authenticated = $true }
                } else {
                    Write-Result "ok" @{ authenticated = $false }
                }
            } catch {
                Write-Result "ok" @{ authenticated = $false }
            }
        }

        "list-gpos" {
            $domain = $args2.domain
            $domainDN = ($domain -split '\.' | ForEach-Object { "DC=$_" }) -join ','
            try {
                $de = New-Object System.DirectoryServices.DirectoryEntry("LDAP://CN=Policies,CN=System,$domainDN")
                $searcher = New-Object System.DirectoryServices.DirectorySearcher($de)
                $searcher.Filter = "(objectClass=groupPolicyContainer)"
                $searcher.PropertiesToLoad.AddRange(@("displayName","flags","whenCreated","whenChanged","gPCFileSysPath"))
                $results = $searcher.FindAll()
                $gpos = @()
                foreach ($r in $results) {
                    $flags = [int]$r.Properties["flags"][0]
                    $status = switch ($flags) { 0 {"AllSettingsEnabled"} 1 {"UserSettingsDisabled"} 2 {"ComputerSettingsDisabled"} 3 {"AllSettingsDisabled"} default {"Unknown"} }
                    $gpos += @{
                        name = [string]$r.Properties["displayname"][0]
                        status = $status
                        created = ([datetime]$r.Properties["whencreated"][0]).ToString("dd/MM/yyyy")
                        modified = ([datetime]$r.Properties["whenchanged"][0]).ToString("dd/MM/yyyy")
                        path = [string]$r.Properties["gpcfilesyspath"][0]
                    }
                }
                $gpos = $gpos | Sort-Object { $_.name }
                Write-Result "ok" $gpos
            } catch {
                Write-Result "error" -Message "Erro LDAP: $($_.Exception.Message)"
            }
        }

        "list-computers" {
            $domain = $args2.domain
            $domainDN = ($domain -split '\.' | ForEach-Object { "DC=$_" }) -join ','
            try {
                $de = New-Object System.DirectoryServices.DirectoryEntry("LDAP://$domainDN")
                $searcher = New-Object System.DirectoryServices.DirectorySearcher($de)
                $searcher.Filter = "(&(objectClass=computer)(!(userAccountControl:1.2.840.113556.1.4.803:=2)))"
                $searcher.PropertiesToLoad.AddRange(@("cn","operatingSystem","lastLogonTimestamp","distinguishedName","description"))
                $searcher.PageSize = 1000
                $results = $searcher.FindAll()
                $pcs = @()
                foreach ($r in $results) {
                    $cn = [string]$r.Properties["cn"][0]
                    $os = if ($r.Properties["operatingsystem"].Count -gt 0) { [string]$r.Properties["operatingsystem"][0] } else { "" }
                    $desc = if ($r.Properties["description"].Count -gt 0) { [string]$r.Properties["description"][0] } else { "" }
                    $lastLogon = ""
                    if ($r.Properties["lastlogontimestamp"].Count -gt 0) {
                        try { $lastLogon = [datetime]::FromFileTime([long]$r.Properties["lastlogontimestamp"][0]).ToString("dd/MM/yyyy HH:mm") } catch {}
                    }
                    $dn = [string]$r.Properties["distinguishedname"][0]
                    $ou = ""
                    if ($dn -match "OU=(.+?)," ) { $ou = $Matches[1] }
                    $pcs += @{
                        name = $cn
                        os = $os
                        description = $desc
                        lastLogon = $lastLogon
                        ou = $ou
                        dn = $dn
                    }
                }
                $pcs = $pcs | Sort-Object { $_.name }
                Write-Result "ok" $pcs
            } catch {
                Write-Result "error" -Message "Erro LDAP: $($_.Exception.Message)"
            }
        }

        "get-blocked-apps" {
            # Ler blocked-apps.json do share configurado
            $sharePath = $args2.sharePath
            $hostname = $args2.hostname
            try {
                if (-not (Test-Path $sharePath)) {
                    Write-Result "ok" @{ global = @(); machine = @(); allApps = @() }
                    return
                }
                $json = Get-Content $sharePath -Raw | ConvertFrom-Json
                $global = @()
                $machine = @()
                if ($json.Global) { $global = @($json.Global) }
                if ($hostname -and $json.Machines -and $json.Machines.PSObject.Properties[$hostname]) {
                    $machine = @($json.Machines.$hostname)
                }
                Write-Result "ok" @{ global = $global; machine = $machine; allApps = ($global + $machine) }
            } catch {
                Write-Result "error" -Message "Erro ao ler: $($_.Exception.Message)"
            }
        }

        "save-blocked-apps" {
            $sharePath = $args2.sharePath
            $hostname = $args2.hostname
            $apps = $args2.apps  # array de nomes de processo
            $scope = $args2.scope  # "global" ou "machine"
            try {
                $json = @{ Global = @(); Machines = @{} }
                if (Test-Path $sharePath) {
                    $existing = Get-Content $sharePath -Raw | ConvertFrom-Json
                    if ($existing.Global) { $json.Global = @($existing.Global) }
                    if ($existing.Machines) {
                        foreach ($prop in $existing.Machines.PSObject.Properties) {
                            $json.Machines[$prop.Name] = @($prop.Value)
                        }
                    }
                }
                if ($scope -eq "global") {
                    $json.Global = @($apps)
                } elseif ($scope -eq "machine" -and $hostname) {
                    $json.Machines[$hostname] = @($apps)
                }
                $outJson = $json | ConvertTo-Json -Depth 5
                [System.IO.File]::WriteAllText($sharePath, $outJson, (New-Object System.Text.UTF8Encoding($false)))
                Write-Result "ok" @{ saved = $true }
            } catch {
                Write-Result "error" -Message "Erro ao salvar: $($_.Exception.Message)"
            }
        }

        "install-remote-agent" {
            # Instala winsysmon.ps1 em um PC remoto via WinRM (Invoke-Command)
            $hostname = $args2.hostname
            $sharePath = $args2.sharePath  # UNC ex: \\server\share\blocked-apps.json
            $user = $args2.username
            $pass = $args2.password
            $domain = $args2.domain
            $localScript = Join-Path (Split-Path $PSCommandPath -Parent) "..\service\winsysmon.ps1"
            $localScript = [System.IO.Path]::GetFullPath($localScript)
            try {
                if (-not (Test-Path $localScript)) {
                    Write-Result "error" -Message "winsysmon.ps1 nao encontrado: $localScript"
                    return
                }
                $scriptContent = Get-Content $localScript -Raw
                $secPass = ConvertTo-SecureString $pass -AsPlainText -Force
                $cred = New-Object System.Management.Automation.PSCredential("$domain\$user", $secPass)

                # Testa conectividade WinRM
                try {
                    Test-WSMan -ComputerName $hostname -ErrorAction Stop | Out-Null
                } catch {
                    Write-Result "error" -Message "WinRM nao responde em $hostname. Habilite WinRM via GPO ou execute: Enable-PSRemoting -Force no PC alvo."
                    return
                }

                $result = Invoke-Command -ComputerName $hostname -Credential $cred -ScriptBlock {
                    param($content, $shareArg)
                    $remoteDir = "$env:ProgramData\Microsoft\WinSysMon"
                    if (-not (Test-Path $remoteDir)) { New-Item $remoteDir -ItemType Directory -Force | Out-Null }
                    $scriptFile = Join-Path $remoteDir "winsysmon.ps1"
                    [System.IO.File]::WriteAllText($scriptFile, $content, (New-Object System.Text.UTF8Encoding($false)))
                    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $scriptFile -Install -SharePath $shareArg 2>&1
                    return "OK em $env:COMPUTERNAME"
                } -ArgumentList $scriptContent, $sharePath -ErrorAction Stop

                Write-Result "ok" @{ installed = $true; output = ($result -join "`n") }
            } catch {
                Write-Result "error" -Message "Erro remoto: $($_.Exception.Message)"
            }
        }

        "uninstall-remote-agent" {
            $hostname = $args2.hostname
            $user = $args2.username
            $pass = $args2.password
            $domain = $args2.domain
            try {
                $secPass = ConvertTo-SecureString $pass -AsPlainText -Force
                $cred = New-Object System.Management.Automation.PSCredential("$domain\$user", $secPass)
                $result = Invoke-Command -ComputerName $hostname -Credential $cred -ScriptBlock {
                    $scriptFile = "$env:ProgramData\Microsoft\WinSysMon\winsysmon.ps1"
                    if (Test-Path $scriptFile) {
                        & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $scriptFile -Uninstall 2>&1
                    } else { "Agente nao encontrado" }
                } -ErrorAction Stop
                Write-Result "ok" @{ output = ($result -join "`n") }
            } catch {
                Write-Result "error" -Message "Erro remoto: $($_.Exception.Message)"
            }
        }

        "check-remote-agent" {
            # Verifica se agente esta instalado + rodando no PC
            $hostname = $args2.hostname
            $user = $args2.username
            $pass = $args2.password
            $domain = $args2.domain
            try {
                $secPass = ConvertTo-SecureString $pass -AsPlainText -Force
                $cred = New-Object System.Management.Automation.PSCredential("$domain\$user", $secPass)
                $result = Invoke-Command -ComputerName $hostname -Credential $cred -ScriptBlock {
                    $t = Get-ScheduledTask -TaskName "WinSysMon" -ErrorAction SilentlyContinue
                    if ($t) {
                        $i = Get-ScheduledTaskInfo -TaskName "WinSysMon" -ErrorAction SilentlyContinue
                        return @{ installed = $true; state = [string]$t.State; lastRun = if($i){[string]$i.LastRunTime}else{""} }
                    }
                    return @{ installed = $false }
                } -ErrorAction Stop
                Write-Result "ok" $result
            } catch {
                Write-Result "error" -Message "Erro: $($_.Exception.Message)"
            }
        }

        "gen-gpo-bat" {
            # Gera .bat que a GPO pode executar como Startup/Logon script
            # Distribui o winsysmon.ps1 para um share e instala via GPO
            $sharePath = $args2.sharePath   # ex: \\server\share\blocked-apps.json
            $scriptShare = $args2.scriptShare  # ex: \\server\NETLOGON\winsysmon.ps1 (script em share)
            $outBat = $args2.outBat
            try {
                $bat = @"
@echo off
REM ============================================================
REM  WinSysMon - Startup script (aplicar via GPO Computer Startup)
REM  Gerado automaticamente em $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
REM ============================================================
setlocal
set "SCRIPT_SOURCE=$scriptShare"
set "SHARE_PATH=$sharePath"
set "LOCAL_DIR=%ProgramData%\Microsoft\WinSysMon"
set "LOCAL_SCRIPT=%LOCAL_DIR%\winsysmon.ps1"
set "FLAG=%LOCAL_DIR%\installed.flag"

if not exist "%LOCAL_DIR%" mkdir "%LOCAL_DIR%"

REM Copiar script (sempre, para atualizar)
copy /Y "%SCRIPT_SOURCE%" "%LOCAL_SCRIPT%" >nul 2>&1
if not exist "%LOCAL_SCRIPT%" (
    echo [ERRO] Nao foi possivel copiar %SCRIPT_SOURCE% >> "%LOCAL_DIR%\install.log"
    exit /b 1
)

REM Instalar apenas se nao tiver flag
if exist "%FLAG%" (
    echo [%DATE% %TIME%] Ja instalado. Apenas garantindo tarefa. >> "%LOCAL_DIR%\install.log"
    schtasks /query /TN "WinSysMon" >nul 2>&1
    if not errorlevel 1 exit /b 0
)

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%LOCAL_SCRIPT%" -Install -SharePath "%SHARE_PATH%" >> "%LOCAL_DIR%\install.log" 2>&1
if not errorlevel 1 (
    echo Instalado em %DATE% %TIME% > "%FLAG%"
)
exit /b 0
"@
                [System.IO.File]::WriteAllText($outBat, $bat, (New-Object System.Text.UTF8Encoding($false)))
                Write-Result "ok" @{ path = $outBat }
            } catch {
                Write-Result "error" -Message "Erro: $($_.Exception.Message)"
            }
        }

        "hv-list-hosts" {
            # Procura computadores no AD com Hyper-V provavel (servers)
            $domain = $args2.domain
            $domainDN = ($domain -split '\.' | ForEach-Object { "DC=$_" }) -join ','
            try {
                $de = New-Object System.DirectoryServices.DirectoryEntry("LDAP://$domainDN")
                $searcher = New-Object System.DirectoryServices.DirectorySearcher($de)
                $searcher.Filter = "(&(objectClass=computer)(operatingSystem=*Server*)(!(userAccountControl:1.2.840.113556.1.4.803:=2)))"
                $searcher.PropertiesToLoad.AddRange(@("cn","operatingSystem","dNSHostName"))
                $searcher.PageSize = 500
                $results = $searcher.FindAll()
                $hosts = @()
                foreach ($r in $results) {
                    $cn = [string]$r.Properties["cn"][0]
                    $os = if ($r.Properties["operatingsystem"].Count -gt 0) { [string]$r.Properties["operatingsystem"][0] } else { "" }
                    $dns = if ($r.Properties["dnshostname"].Count -gt 0) { [string]$r.Properties["dnshostname"][0] } else { $cn }
                    $hosts += @{ name = $cn; os = $os; dns = $dns }
                }
                $hosts = $hosts | Sort-Object { $_.name }
                Write-Result "ok" $hosts
            } catch {
                Write-Result "error" -Message "Erro LDAP: $($_.Exception.Message)"
            }
        }

        "hv-list-vms" {
            $hvHost = $args2.hvHost
            $user = $args2.username; $pass = $args2.password; $domain = $args2.domain
            try {
                $cred = Get-HvCredential -username $user -password $pass -domain $domain
                $result = Invoke-HvRemote -ComputerName $hvHost -Credential $cred -ScriptBlock {
                    if (-not (Get-Module -ListAvailable -Name Hyper-V)) { return @{ hvError = "Modulo Hyper-V nao instalado em $env:COMPUTERNAME" } }
                    Import-Module Hyper-V -ErrorAction Stop
                    $vms = Get-VM | ForEach-Object {
                        [pscustomobject]@{
                            name         = $_.Name
                            state        = [string]$_.State
                            status       = [string]$_.Status
                            cpuUsage     = [int]$_.CPUUsage
                            memAssigned  = [int]($_.MemoryAssigned / 1MB)
                            memDemand    = [int]($_.MemoryDemand / 1MB)
                            memStartup   = [int]($_.MemoryStartup / 1MB)
                            processors   = [int]$_.ProcessorCount
                            uptime       = if ($_.Uptime) { $_.Uptime.ToString("d\.hh\:mm\:ss") } else { "" }
                            generation   = [int]$_.Generation
                            version      = [string]$_.Version
                            notes        = [string]$_.Notes
                        }
                    }
                    return @{ ok = $true; vms = @($vms) }
                }
                if ($result.hvError) { Write-Result "error" -Message $result.hvError; return }
                Write-Result "ok" @{ vms = @($result.vms) }
            } catch {
                Write-Result "error" -Message "Erro Hyper-V: $($_.Exception.Message)"
            }
        }

        "hv-vm-action" {
            $hvHost = $args2.hvHost; $vmName = $args2.vmName; $act = $args2.action
            $user = $args2.username; $pass = $args2.password; $domain = $args2.domain
            try {
                $cred = Get-HvCredential -username $user -password $pass -domain $domain
                $result = Invoke-HvRemote -ComputerName $hvHost -Credential $cred -ArgumentList $vmName, $act -ScriptBlock {
                    param($name, $action)
                    Import-Module Hyper-V -ErrorAction Stop
                    switch ($action) {
                        "start"    { Start-VM -Name $name -ErrorAction Stop }
                        "stop"     { Stop-VM -Name $name -TurnOff -Force -ErrorAction Stop }
                        "shutdown" { Stop-VM -Name $name -Force -ErrorAction Stop }
                        "restart"  { Restart-VM -Name $name -Force -ErrorAction Stop }
                        "save"     { Save-VM -Name $name -ErrorAction Stop }
                        "pause"    { Suspend-VM -Name $name -ErrorAction Stop }
                        "resume"   { Resume-VM -Name $name -ErrorAction Stop }
                        default    { throw "Acao invalida: $action" }
                    }
                    $v = Get-VM -Name $name
                    return @{ state = [string]$v.State; status = [string]$v.Status }
                }
                Write-Result "ok" $result
            } catch {
                Write-Result "error" -Message "Erro: $($_.Exception.Message)"
            }
        }

        "hv-list-snapshots" {
            $hvHost = $args2.hvHost; $vmName = $args2.vmName
            $user = $args2.username; $pass = $args2.password; $domain = $args2.domain
            try {
                $cred = Get-HvCredential -username $user -password $pass -domain $domain
                $result = Invoke-HvRemote -ComputerName $hvHost -Credential $cred -ArgumentList $vmName -ScriptBlock {
                    param($name)
                    Import-Module Hyper-V -ErrorAction Stop
                    $snaps = Get-VMSnapshot -VMName $name -ErrorAction SilentlyContinue | ForEach-Object {
                        [pscustomobject]@{
                            name       = $_.Name
                            created    = $_.CreationTime.ToString("dd/MM/yyyy HH:mm")
                            parent     = [string]$_.ParentSnapshotName
                            type       = [string]$_.SnapshotType
                        }
                    }
                    return @{ snapshots = @($snaps) }
                }
                Write-Result "ok" $result
            } catch {
                Write-Result "error" -Message "Erro: $($_.Exception.Message)"
            }
        }

        "hv-snapshot-action" {
            $hvHost = $args2.hvHost; $vmName = $args2.vmName; $snapName = $args2.snapshotName; $act = $args2.action; $newName = $args2.newName
            $user = $args2.username; $pass = $args2.password; $domain = $args2.domain
            try {
                $cred = Get-HvCredential -username $user -password $pass -domain $domain
                $result = Invoke-HvRemote -ComputerName $hvHost -Credential $cred -ArgumentList $vmName, $snapName, $act, $newName -ScriptBlock {
                    param($vm, $snap, $action, $newName)
                    Import-Module Hyper-V -ErrorAction Stop
                    switch ($action) {
                        "create"  {
                            $sn = if ([string]::IsNullOrWhiteSpace($newName)) { "Snap_$(Get-Date -Format 'yyyyMMdd_HHmmss')" } else { $newName }
                            Checkpoint-VM -Name $vm -SnapshotName $sn -ErrorAction Stop
                            return @{ created = $sn }
                        }
                        "restore" {
                            Restore-VMSnapshot -VMName $vm -Name $snap -Confirm:$false -ErrorAction Stop
                            return @{ restored = $snap }
                        }
                        "delete"  {
                            Remove-VMSnapshot -VMName $vm -Name $snap -Confirm:$false -ErrorAction Stop
                            return @{ deleted = $snap }
                        }
                        "rename"  {
                            Rename-VMSnapshot -VMName $vm -Name $snap -NewName $newName -ErrorAction Stop
                            return @{ renamed = $newName }
                        }
                        default   { throw "Acao invalida: $action" }
                    }
                }
                Write-Result "ok" $result
            } catch {
                Write-Result "error" -Message "Erro: $($_.Exception.Message)"
            }
        }

        "hv-vm-stats" {
            $hvHost = $args2.hvHost; $vmName = $args2.vmName
            $user = $args2.username; $pass = $args2.password; $domain = $args2.domain
            try {
                $cred = Get-HvCredential -username $user -password $pass -domain $domain
                $result = Invoke-HvRemote -ComputerName $hvHost -Credential $cred -ArgumentList $vmName -ScriptBlock {
                    param($name)
                    Import-Module Hyper-V -ErrorAction Stop
                    $v = Get-VM -Name $name -ErrorAction Stop
                    $vhd = Get-VHD -VMId $v.VMId -ErrorAction SilentlyContinue | Select-Object -First 1
                    $net = Get-VMNetworkAdapter -VMName $name -ErrorAction SilentlyContinue | Select-Object -First 1
                    return @{
                        name         = $v.Name
                        state        = [string]$v.State
                        status       = [string]$v.Status
                        cpuUsage     = [int]$v.CPUUsage
                        memAssigned  = [int]($v.MemoryAssigned / 1MB)
                        memDemand    = [int]($v.MemoryDemand / 1MB)
                        memStartup   = [int]($v.MemoryStartup / 1MB)
                        memMinimum   = [int]($v.MemoryMinimum / 1MB)
                        memMaximum   = [int]($v.MemoryMaximum / 1MB)
                        processors   = [int]$v.ProcessorCount
                        uptime       = if ($v.Uptime) { $v.Uptime.ToString("d\.hh\:mm\:ss") } else { "" }
                        generation   = [int]$v.Generation
                        vhdPath      = if ($vhd) { [string]$vhd.Path } else { "" }
                        vhdSizeGB    = if ($vhd) { [math]::Round($vhd.Size / 1GB, 2) } else { 0 }
                        vhdUsedGB    = if ($vhd) { [math]::Round($vhd.FileSize / 1GB, 2) } else { 0 }
                        switchName   = if ($net) { [string]$net.SwitchName } else { "" }
                        ipAddresses  = if ($net) { @($net.IPAddresses) -join "," } else { "" }
                        macAddress   = if ($net) { [string]$net.MacAddress } else { "" }
                    }
                }
                Write-Result "ok" $result
            } catch {
                Write-Result "error" -Message "Erro: $($_.Exception.Message)"
            }
        }

        "hv-list-switches" {
            $hvHost = $args2.hvHost
            $user = $args2.username; $pass = $args2.password; $domain = $args2.domain
            try {
                $cred = Get-HvCredential -username $user -password $pass -domain $domain
                $result = Invoke-HvRemote -ComputerName $hvHost -Credential $cred -ScriptBlock {
                    Import-Module Hyper-V -ErrorAction Stop
                    $sw = Get-VMSwitch | ForEach-Object {
                        [pscustomobject]@{ name = $_.Name; type = [string]$_.SwitchType }
                    }
                    return @{ switches = @($sw) }
                }
                Write-Result "ok" $result
            } catch {
                Write-Result "error" -Message "Erro: $($_.Exception.Message)"
            }
        }

        "hv-create-vm" {
            $hvHost = $args2.hvHost; $vmName = $args2.vmName
            $memoryMB = [int]$args2.memoryMB
            $vhdSizeGB = [int]$args2.vhdSizeGB
            $switchName = $args2.switchName
            $generation = if ($args2.generation) { [int]$args2.generation } else { 2 }
            $processors = if ($args2.processors) { [int]$args2.processors } else { 2 }
            $user = $args2.username; $pass = $args2.password; $domain = $args2.domain
            try {
                $cred = Get-HvCredential -username $user -password $pass -domain $domain
                $result = Invoke-HvRemote -ComputerName $hvHost -Credential $cred -ArgumentList $vmName,$memoryMB,$vhdSizeGB,$switchName,$generation,$processors -ScriptBlock {
                    param($name, $memMB, $sizeGB, $sw, $gen, $procs)
                    Import-Module Hyper-V -ErrorAction Stop
                    if (Get-VM -Name $name -ErrorAction SilentlyContinue) { throw "VM '$name' ja existe" }

                    $hvHostCfg = Get-VMHost
                    $vhdDir = $hvHostCfg.VirtualHardDiskPath
                    if (-not (Test-Path $vhdDir)) { New-Item $vhdDir -ItemType Directory -Force | Out-Null }
                    $vhdPath = Join-Path $vhdDir ("$name.vhdx")

                    New-VHD -Path $vhdPath -SizeBytes ([int64]$sizeGB * 1GB) -Dynamic -ErrorAction Stop | Out-Null

                    $vmParams = @{
                        Name               = $name
                        MemoryStartupBytes = ([int64]$memMB * 1MB)
                        Generation         = $gen
                        VHDPath            = $vhdPath
                    }
                    if ($sw) { $vmParams.SwitchName = $sw }
                    New-VM @vmParams -ErrorAction Stop | Out-Null
                    Set-VM -Name $name -ProcessorCount $procs -ErrorAction SilentlyContinue
                    Set-VMMemory -VMName $name -DynamicMemoryEnabled $true `
                        -MinimumBytes ([int64]([math]::Max(512, $memMB / 2)) * 1MB) `
                        -MaximumBytes ([int64]($memMB * 2) * 1MB) `
                        -StartupBytes ([int64]$memMB * 1MB) -ErrorAction SilentlyContinue
                    return @{ created = $name; vhd = $vhdPath }
                }
                Write-Result "ok" $result
            } catch {
                Write-Result "error" -Message "Erro ao criar VM: $($_.Exception.Message)"
            }
        }

        "hv-delete-vm" {
            $hvHost = $args2.hvHost; $vmName = $args2.vmName
            $user = $args2.username; $pass = $args2.password; $domain = $args2.domain
            $deleteVhd = if ($args2.deleteVhd) { [bool]$args2.deleteVhd } else { $false }
            try {
                $cred = Get-HvCredential -username $user -password $pass -domain $domain
                $result = Invoke-HvRemote -ComputerName $hvHost -Credential $cred -ArgumentList $vmName, $deleteVhd -ScriptBlock {
                    param($name, $delVhd)
                    Import-Module Hyper-V -ErrorAction Stop
                    $v = Get-VM -Name $name -ErrorAction Stop
                    $vhdPaths = @()
                    if ($delVhd) {
                        $vhdPaths = @(Get-VHD -VMId $v.VMId -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Path)
                    }
                    if ($v.State -ne 'Off') { Stop-VM -Name $name -TurnOff -Force -ErrorAction SilentlyContinue }
                    Remove-VM -Name $name -Force -ErrorAction Stop
                    if ($delVhd) {
                        foreach ($p in $vhdPaths) { if ($p -and (Test-Path $p)) { Remove-Item $p -Force -ErrorAction SilentlyContinue } }
                    }
                    return @{ deleted = $name; vhdsRemoved = $vhdPaths.Count }
                }
                Write-Result "ok" $result
            } catch {
                Write-Result "error" -Message "Erro ao remover VM: $($_.Exception.Message)"
            }
        }

        "hv-test-connection" {
            # Testa conectividade e auth para um host Hyper-V.
            # Retorna diagnostico detalhado sem explodir a UI.
            $hvHost = $args2.hvHost
            $user = $args2.username; $pass = $args2.password; $domain = $args2.domain
            $diag = @{
                host       = $hvHost
                port5985   = "unknown"
                port5986   = "unknown"
                authOk     = $false
                authMethod = ""
                hvOk       = $false
                error      = ""
            }
            function Test-TcpQuick($h, $p) {
                try {
                    $c = New-Object System.Net.Sockets.TcpClient
                    $ar = $c.BeginConnect($h, $p, $null, $null)
                    $ok = $ar.AsyncWaitHandle.WaitOne(2500, $false)
                    if ($ok -and $c.Connected) { $c.Close(); return "OPEN" }
                    $c.Close(); return "TIMEOUT"
                } catch { return "ERR" }
            }
            $diag.port5985 = Test-TcpQuick $hvHost 5985
            $diag.port5986 = Test-TcpQuick $hvHost 5986
            if ($diag.port5985 -ne "OPEN" -and $diag.port5986 -ne "OPEN") {
                $diag.error = "Portas WinRM (5985/5986) fechadas. Verifique firewall e rede."
                Write-Result "ok" $diag
                return
            }
            try {
                $cred = Get-HvCredential -username $user -password $pass -domain $domain
                $r = Invoke-HvRemote -ComputerName $hvHost -Credential $cred -ScriptBlock {
                    @{ name = $env:COMPUTERNAME; hasHv = ($null -ne (Get-Module -ListAvailable -Name Hyper-V)) }
                }
                $diag.authOk = $true
                $diag.authMethod = if ($cred.UserName -match '^\.\\' -or ($cred.UserName -match '^([^\\]+)\\' -and $Matches[1] -notmatch '\.')) { "Negotiate/NTLM" } else { "Kerberos" }
                $diag.hvOk = [bool]$r.hasHv
                if (-not $diag.hvOk) { $diag.error = "Modulo Hyper-V nao instalado em $hvHost" }
            } catch {
                $diag.error = $_.Exception.Message
            }
            Write-Result "ok" $diag
        }

        "hv-configure-trusted-hosts" {
            # Adiciona hosts a TrustedHosts local. REQUER ADMIN.
            # Se nao for admin, relanca elevado via Start-Process e retorna status.
            $hostsToAdd = $args2.hosts
            if (-not $hostsToAdd -or $hostsToAdd.Count -eq 0) {
                Write-Result "error" -Message "Nenhum host informado"
                return
            }
            $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
            $hostsList = ($hostsToAdd | Where-Object { $_ }) -join ","

            if (-not $isAdmin) {
                # Tenta elevar via Start-Process -Verb RunAs executando script temporario
                $tmp = [System.IO.Path]::Combine($env:TEMP, "set_trustedhosts_$([guid]::NewGuid().ToString('N')).ps1")
                $psBody = @"
try {
    Start-Service WinRM -ErrorAction SilentlyContinue
    `$cur = ''
    try { `$cur = (Get-Item WSMan:\localhost\Client\TrustedHosts -ErrorAction Stop).Value } catch {}
    `$new = '$hostsList'
    if (`$cur -and `$cur.Trim().Length -gt 0) {
        `$parts = @(`$cur -split ',' | ForEach-Object { `$_.Trim() } | Where-Object { `$_ })
        foreach (`$h in (`$new -split ',')) { if (`$parts -notcontains `$h) { `$parts += `$h } }
        `$new = (`$parts -join ',')
    }
    Set-Item WSMan:\localhost\Client\TrustedHosts -Value `$new -Force
    '$tmp.ok' | Out-File "$tmp.result" -Encoding utf8
} catch {
    "ERR: `$(`$_.Exception.Message)" | Out-File "$tmp.result" -Encoding utf8
}
"@
                Set-Content -Path $tmp -Value $psBody -Encoding UTF8
                try {
                    $p = Start-Process powershell -ArgumentList "-NoProfile","-ExecutionPolicy","Bypass","-File",$tmp -Verb RunAs -WindowStyle Hidden -PassThru -ErrorAction Stop
                    $p.WaitForExit(30000) | Out-Null
                    $resFile = "$tmp.result"
                    $msg = if (Test-Path $resFile) { (Get-Content $resFile -Raw).Trim() } else { "sem resposta" }
                    Remove-Item $tmp, $resFile -ErrorAction SilentlyContinue
                    if ($msg -like "ERR:*") {
                        Write-Result "error" -Message $msg
                    } else {
                        Write-Result "ok" @{ configured = $hostsList; method = "elevated" }
                    }
                } catch {
                    Remove-Item $tmp -ErrorAction SilentlyContinue
                    Write-Result "error" -Message "Falha ao elevar: $($_.Exception.Message). Execute o app como Administrador."
                }
                return
            }

            # Ja esta como admin
            try {
                Start-Service WinRM -ErrorAction SilentlyContinue
                $cur = ""
                try { $cur = (Get-Item WSMan:\localhost\Client\TrustedHosts -ErrorAction Stop).Value } catch {}
                $new = $hostsList
                if ($cur -and $cur.Trim().Length -gt 0) {
                    $parts = @($cur -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
                    foreach ($h in ($new -split ',')) { if ($parts -notcontains $h) { $parts += $h } }
                    $new = ($parts -join ',')
                }
                Set-Item WSMan:\localhost\Client\TrustedHosts -Value $new -Force
                Write-Result "ok" @{ configured = $new; method = "direct" }
            } catch {
                Write-Result "error" -Message "Erro: $($_.Exception.Message)"
            }
        }

        default {
            Write-Result "error" -Message "Comando desconhecido: $action"
        }
    }
} catch {
    Write-Result "error" -Message "Excecao: $($_.Exception.Message)"
}