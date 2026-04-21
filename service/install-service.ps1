# ============================================================
#  install-service.ps1 - Instala WinSysMon como SERVICO WINDOWS
#  Aparece em services.msc
#  Compila wrapper C# nativo que chama o winsysmon.ps1 -RunLoop
# ============================================================
#Requires -RunAsAdministrator

param(
    [switch]$Uninstall,
    [string]$SharePath = "\\srv-105\Sistema de monitoramento\gpo\aaa\service\blocked-apps.json",
    [string]$HostsSharePath = "\\srv-105\Sistema de monitoramento\gpo\aaa\service\blocked-hosts.json",
    [string]$PoliciesSharePath = "\\srv-105\Sistema de monitoramento\gpo\aaa\service\blocked-policies.json"
)

$ErrorActionPreference = "Stop"
$serviceName = "WinSysMon"
$displayName = "Windows System Monitor"
$description = "Monitors system performance and enforces application policies."
$installDir  = "$env:ProgramData\Microsoft\WinSysMon"
$exePath     = Join-Path $installDir "WinSysMonSvc.exe"
$psScript    = Join-Path $installDir "winsysmon.ps1"
$configPath  = Join-Path $installDir "sysmon-config.json"
$installLog  = Join-Path $installDir "install.log"

# ---- Garantir diretorio (para log) ----
if (-not (Test-Path $installDir)) { New-Item $installDir -ItemType Directory -Force | Out-Null }

function Write-InstallLog {
    param([string]$Msg, [string]$Level = "INFO")
    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $line = "[$ts] [$Level] $Msg"
    try { Add-Content -Path $installLog -Value $line -ErrorAction SilentlyContinue } catch {}
    $color = switch ($Level) { "ERROR" {"Red"} "WARN" {"Yellow"} "OK" {"Green"} default {"Gray"} }
    Write-Host $line -ForegroundColor $color
}

function Test-Admin {
    $id = [System.Security.Principal.WindowsIdentity]::GetCurrent()
    $p = New-Object System.Security.Principal.WindowsPrincipal($id)
    return $p.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)
}

if (-not (Test-Admin)) {
    Write-InstallLog "Este script precisa rodar como ADMINISTRADOR" "ERROR"
    exit 1
}

Write-InstallLog "=== Instalador $serviceName em $env:COMPUTERNAME ===" "OK"

# ---- Uninstall ----
if ($Uninstall) {
    Write-InstallLog "Modo desinstalar"
    try { Stop-Service $serviceName -Force -ErrorAction SilentlyContinue } catch {}
    & sc.exe delete $serviceName | Out-Null
    Start-Sleep 1
    Write-InstallLog "Servico removido" "OK"
    exit 0
}

# ---- Resolver fonte do winsysmon.ps1 ----
$sourceScript = $null
$candidates = @(1
    (Join-Path $PSScriptRoot "winsysmon.ps1"),
    "\\srv-105\Sistema de monitoramento\gpo\aaa\service\winsysmon.ps1",
    "c:\gpo\service\winsysmon.ps1"
)
foreach ($c in $candidates) {
    try { if ($c -and (Test-Path $c -ErrorAction Stop)) { $sourceScript = $c; break } } catch {}
}
if (-not $sourceScript) {
    Write-InstallLog "winsysmon.ps1 nao encontrado em nenhum candidato" "ERROR"
    exit 1
}
Write-InstallLog "Fonte do script: $sourceScript"

# Validar sintaxe antes de copiar
try {
    $parseErrs = $null
    $null = [System.Management.Automation.Language.Parser]::ParseFile($sourceScript, [ref]$null, [ref]$parseErrs)
    if ($parseErrs -and $parseErrs.Count -gt 0) {
        Write-InstallLog "Script de origem tem $($parseErrs.Count) erros de sintaxe - ABORTANDO" "ERROR"
        $parseErrs | Select-Object -First 3 | ForEach-Object { Write-InstallLog ("  " + $_.Message) "ERROR" }
        exit 1
    }
} catch {
    Write-InstallLog "Falha ao validar sintaxe: $($_.Exception.Message)" "WARN"
}

# ---- Copiar winsysmon.ps1 ----
try {
    Copy-Item $sourceScript $psScript -Force -ErrorAction Stop
    Write-InstallLog "Script copiado: $psScript" "OK"
} catch {
    Write-InstallLog "Falha ao copiar script: $($_.Exception.Message)" "ERROR"
    exit 1
}

# ---- Config (preserva existente, merge com SharePath atualizado) ----
$defaultCfg = @{
    BlockedApps             = @()
    PollInterval            = 1
    MonitorLogins           = $true
    MonitorApps             = $true
    CollectHardware         = $true
    HardwareInterval        = 3600
    RemoteBlockedAppsPath   = $SharePath
    RemoteBlockedHostsPath  = $HostsSharePath
    RemoteBlockedPoliciesPath = $PoliciesSharePath
    HostBlockingEnabled     = $true
    HostBlockingInterval    = 60
    PolicyBlockingEnabled   = $true
    PolicyBlockingInterval  = 60
}
$finalCfg = $defaultCfg.Clone()
if (Test-Path $configPath) {
    try {
        $existing = Get-Content $configPath -Raw -ErrorAction Stop | ConvertFrom-Json
        foreach ($p in $existing.PSObject.Properties) { $finalCfg[$p.Name] = $p.Value }
        Write-InstallLog "Config existente preservada"
    } catch {
        Write-InstallLog "Config corrompida - recriando com defaults" "WARN"
    }
}
# Campos criticos sempre force defaults (evita configs antigas ruins)
$finalCfg.PollInterval = 1
if ($SharePath)      { $finalCfg.RemoteBlockedAppsPath  = $SharePath }
if ($HostsSharePath) { $finalCfg.RemoteBlockedHostsPath = $HostsSharePath }
if ($PoliciesSharePath) { $finalCfg.RemoteBlockedPoliciesPath = $PoliciesSharePath }
try {
    $json = $finalCfg | ConvertTo-Json -Depth 5
    [System.IO.File]::WriteAllText($configPath, $json, (New-Object System.Text.UTF8Encoding($false)))
    Write-InstallLog "Config salva: RemoteBlockedAppsPath=$($finalCfg.RemoteBlockedAppsPath) RemoteBlockedHostsPath=$($finalCfg.RemoteBlockedHostsPath) PollInterval=$($finalCfg.PollInterval)"
} catch {
    Write-InstallLog "Falha ao salvar config: $($_.Exception.Message)" "ERROR"
    exit 1
}

# ---- Parar e remover servico anterior ----
if (Get-Service $serviceName -ErrorAction SilentlyContinue) {
    Write-InstallLog "Parando servico existente..."
    try { Stop-Service $serviceName -Force -ErrorAction SilentlyContinue } catch {}
    # Aguarda parar (ate 10s)
    for ($i = 0; $i -lt 10; $i++) {
        $s = Get-Service $serviceName -ErrorAction SilentlyContinue
        if (-not $s -or $s.Status -eq 'Stopped') { break }
        Start-Sleep 1
    }
    & sc.exe delete $serviceName | Out-Null
    Start-Sleep 2
}

# ---- Remover tarefa agendada legada ----
try {
    if (Get-ScheduledTask -TaskName $serviceName -ErrorAction SilentlyContinue) {
        Unregister-ScheduledTask -TaskName $serviceName -Confirm:$false -ErrorAction SilentlyContinue
        Write-InstallLog "Tarefa agendada legada removida"
    }
} catch {}

# ---- Compilar wrapper C# (apenas se ainda nao existe ou se forcar) ----
$needsCompile = -not (Test-Path $exePath)
if ($needsCompile) {
    Write-InstallLog "Compilando wrapper de servico..."

    $source = @'
using System;
using System.Diagnostics;
using System.IO;
using System.ServiceProcess;
using System.Threading;

namespace WinSysMonSvc
{
    public class Service : ServiceBase
    {
        private Process _child;
        private Thread _watchdog;
        private volatile bool _stopping;
        private string _installDir;
        private string _psScript;
        private string _logFile;
        private int _restartBackoff = 5000;

        public Service()
        {
            this.ServiceName = "WinSysMon";
            this.CanStop = true;
            this.CanShutdown = true;
        }

        private void Log(string msg)
        {
            try { File.AppendAllText(_logFile, "[" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "] [SVC] " + msg + Environment.NewLine); } catch { }
        }

        protected override void OnStart(string[] args)
        {
            _installDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            _psScript   = Path.Combine(_installDir, "winsysmon.ps1");
            _logFile    = Path.Combine(_installDir, "sysmon.log");
            _stopping   = false;

            Log("Servico iniciando (PID=" + Process.GetCurrentProcess().Id + ")");
            _watchdog = new Thread(RunWatchdog);
            _watchdog.IsBackground = true;
            _watchdog.Start();
        }

        private void RunWatchdog()
        {
            int consecutiveFastExits = 0;
            while (!_stopping)
            {
                DateTime startedAt = DateTime.Now;
                try
                {
                    if (!File.Exists(_psScript))
                    {
                        Log("Script nao encontrado: " + _psScript + " - aguardando 30s");
                        Thread.Sleep(30000);
                        continue;
                    }
                    var psi = new ProcessStartInfo();
                    psi.FileName = "powershell.exe";
                    psi.Arguments = "-NoProfile -NonInteractive -ExecutionPolicy Bypass -File \"" + _psScript + "\" -RunLoop";
                    psi.UseShellExecute = false;
                    psi.CreateNoWindow = true;
                    psi.WorkingDirectory = _installDir;
                    psi.RedirectStandardOutput = true;
                    psi.RedirectStandardError = true;

                    _child = Process.Start(psi);
                    Log("PowerShell iniciado PID=" + _child.Id);
                    _child.WaitForExit();
                    Log("PowerShell terminou exitCode=" + _child.ExitCode);
                }
                catch (Exception ex)
                {
                    Log("Erro watchdog: " + ex.Message);
                }

                if (_stopping) break;

                // Backoff exponencial se ficar caindo rapido (max 5min)
                TimeSpan uptime = DateTime.Now - startedAt;
                if (uptime.TotalSeconds < 30) {
                    consecutiveFastExits++;
                    _restartBackoff = Math.Min(_restartBackoff * 2, 300000);
                    Log("Script caiu rapido (" + uptime.TotalSeconds + "s). Backoff=" + _restartBackoff + "ms (falhas seguidas=" + consecutiveFastExits + ")");
                } else {
                    consecutiveFastExits = 0;
                    _restartBackoff = 5000;
                }
                Thread.Sleep(_restartBackoff);
            }
        }

        protected override void OnStop()
        {
            Log("Servico parando...");
            _stopping = true;
            try
            {
                if (_child != null && !_child.HasExited)
                {
                    _child.Kill();
                    _child.WaitForExit(5000);
                }
            }
            catch { }
        }

        protected override void OnShutdown() { OnStop(); }

        public static void Main() { ServiceBase.Run(new Service()); }
    }
}
'@

    $srcFile = Join-Path $installDir "WinSysMonSvc.cs"
    [System.IO.File]::WriteAllText($srcFile, $source, (New-Object System.Text.UTF8Encoding($false)))

    $cscPaths = @(
        "$env:WINDIR\Microsoft.NET\Framework64\v4.0.30319\csc.exe",
        "$env:WINDIR\Microsoft.NET\Framework\v4.0.30319\csc.exe"
    )
    $csc = $cscPaths | Where-Object { Test-Path $_ } | Select-Object -First 1
    if (-not $csc) {
        Write-InstallLog "csc.exe (.NET 4) nao encontrado. Instale .NET Framework 4.x" "ERROR"
        exit 1
    }

    if (Test-Path $exePath) {
        try { Remove-Item $exePath -Force -ErrorAction Stop } catch {
            Write-InstallLog "Nao foi possivel substituir $exePath (em uso?): $($_.Exception.Message)" "ERROR"
            exit 1
        }
    }
    $cscArgs = @("/nologo", "/target:exe", "/out:$exePath", "/r:System.ServiceProcess.dll", $srcFile)
    $p = Start-Process -FilePath $csc -ArgumentList $cscArgs -Wait -PassThru -NoNewWindow -RedirectStandardOutput "$installDir\csc.out.log" -RedirectStandardError "$installDir\csc.err.log"
    if ($p.ExitCode -ne 0 -or -not (Test-Path $exePath)) {
        Write-InstallLog "Falha na compilacao (exit=$($p.ExitCode))" "ERROR"
        if (Test-Path "$installDir\csc.err.log") { Get-Content "$installDir\csc.err.log" | ForEach-Object { Write-InstallLog "  $_" "ERROR" } }
        exit 1
    }
    Remove-Item $srcFile -Force -ErrorAction SilentlyContinue
    Remove-Item "$installDir\csc.out.log" -Force -ErrorAction SilentlyContinue
    Remove-Item "$installDir\csc.err.log" -Force -ErrorAction SilentlyContinue
    Write-InstallLog "Compilado: $exePath" "OK"
} else {
    Write-InstallLog "Wrapper ja existe - reutilizando $exePath"
}

# ---- Criar servico ----
Write-InstallLog "Registrando servico..."
$scOut = & sc.exe create $serviceName binPath= "`"$exePath`"" start= auto DisplayName= "$displayName" 2>&1
if ($LASTEXITCODE -ne 0) {
    Write-InstallLog "sc.exe create falhou: $scOut" "ERROR"
    exit 1
}
& sc.exe description $serviceName "$description" | Out-Null
& sc.exe failure $serviceName reset= 86400 actions= restart/5000/restart/10000/restart/30000 | Out-Null
& sc.exe sdset $serviceName "D:(A;;CCLCSWRPWPDTLOCRRC;;;SY)(A;;CCDCLCSWRPWPDTLOCRSDRCWDWO;;;BA)(A;;CCLCSWLOCRRC;;;IU)(A;;CCLCSWLOCRRC;;;SU)" | Out-Null

# ---- ACL na pasta (usando SIDs bem conhecidos para funcionar em qualquer idioma) ----
try {
    $sidSystem = New-Object System.Security.Principal.SecurityIdentifier([System.Security.Principal.WellKnownSidType]::LocalSystemSid, $null)
    $sidAdmins = New-Object System.Security.Principal.SecurityIdentifier([System.Security.Principal.WellKnownSidType]::BuiltinAdministratorsSid, $null)

    # Toma ownership recursivo (caso alguem tenha mexido antes)
    try { & takeown.exe /F $installDir /R /D Y /A 2>&1 | Out-Null } catch {}

    $acl = New-Object System.Security.AccessControl.DirectorySecurity
    $acl.SetOwner($sidSystem)
    $acl.SetAccessRuleProtection($true, $false)
    $inh  = [System.Security.AccessControl.InheritanceFlags]"ContainerInherit,ObjectInherit"
    $prop = [System.Security.AccessControl.PropagationFlags]::None
    # SYSTEM: controle total (so o servico)
    $acl.AddAccessRule((New-Object System.Security.AccessControl.FileSystemAccessRule($sidSystem,"FullControl",$inh,$prop,"Allow")))
    # Administrators: somente leitura (diagnostico minimo; modificar exige tomar ownership)
    $acl.AddAccessRule((New-Object System.Security.AccessControl.FileSystemAccessRule($sidAdmins,"ReadAndExecute",$inh,$prop,"Allow")))
    Set-Acl -Path $installDir -AclObject $acl

    # Esconde a pasta (hidden + system)
    try {
        $item = Get-Item $installDir -Force
        $item.Attributes = $item.Attributes -bor [System.IO.FileAttributes]::Hidden -bor [System.IO.FileAttributes]::System
    } catch {}

    Write-InstallLog "ACL aplicada: SYSTEM=Full, Admins=ReadOnly, heranca OFF, Hidden+System" "OK"
} catch { Write-InstallLog "ACL nao aplicada: $($_.Exception.Message)" "WARN" }

# ---- Iniciar servico com retry ----
Write-InstallLog "Iniciando servico..."
$started = $false
for ($attempt = 1; $attempt -le 3; $attempt++) {
    try {
        Start-Service $serviceName -ErrorAction Stop
        Start-Sleep 2
        $s = Get-Service $serviceName
        if ($s.Status -eq 'Running') { $started = $true; break }
        Write-InstallLog "Tentativa $attempt : estado=$($s.Status)" "WARN"
        Start-Sleep 3
    } catch {
        Write-InstallLog "Tentativa $attempt falhou: $($_.Exception.Message)" "WARN"
        Start-Sleep 3
    }
}

$final = Get-Service $serviceName -ErrorAction SilentlyContinue
Write-Host ""
Write-Host "==============================================" -ForegroundColor Green
if ($started -and $final.Status -eq 'Running') {
    Write-InstallLog "Servico RUNNING (tentativas=$attempt)" "OK"
    Write-Host "  Status:  OK - Servico rodando" -ForegroundColor Green
} else {
    Write-InstallLog "Servico NAO iniciou (Status=$($final.Status))" "ERROR"
    Write-Host "  Status:  FALHOU - ver sysmon.log e install.log" -ForegroundColor Red
}
Write-Host "  Servico: $serviceName" -ForegroundColor Gray
Write-Host "  Exe:     $exePath" -ForegroundColor Gray
Write-Host "  Log:     $installDir\sysmon.log" -ForegroundColor Gray
Write-Host "  Install: $installLog" -ForegroundColor Gray
Write-Host "==============================================" -ForegroundColor Green
Write-Host ""
Write-Host "Visualizar em: services.msc  (procure por '$displayName')" -ForegroundColor Cyan

if (-not $started) { exit 1 }
exit 0
