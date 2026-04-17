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

        default {
            Write-Result "error" -Message "Comando desconhecido: $action"
        }
    }
} catch {
    Write-Result "error" -Message "Excecao: $($_.Exception.Message)"
}
