# ==============================================================
#  diag-winsysmon.ps1 - Diagnostico completo do WinSysMon
#  Roda na estacao problemática como ADMIN.
#  Mostra estado de todos os 9 canais + share + AV + lista real.
# ==============================================================
$ErrorActionPreference = 'Continue'
$VerbosePreference = 'SilentlyContinue'
$installDir = "$env:ProgramData\Microsoft\WinSysMon"
$cfgPath    = Join-Path $installDir "sysmon-config.json"
$logPath    = Join-Path $installDir "sysmon.log"
$installLog = Join-Path $installDir "install.log"
$hostname   = $env:COMPUTERNAME

function Section { param($t) Write-Host "`n===== $t =====" -ForegroundColor Cyan }
function OK    { param($t) Write-Host "  [OK]   $t" -ForegroundColor Green }
function WARN  { param($t) Write-Host "  [WARN] $t" -ForegroundColor Yellow }
function FAIL  { param($t) Write-Host "  [FAIL] $t" -ForegroundColor Red }
function INFO  { param($t) Write-Host "  [..]   $t" -ForegroundColor Gray }

Write-Host "========================================================" -ForegroundColor Cyan
Write-Host "|  WinSysMon - Diagnostico em $hostname".PadRight(55) + "|" -ForegroundColor Cyan
Write-Host "|  $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')".PadRight(55) + "|" -ForegroundColor Cyan
Write-Host "========================================================" -ForegroundColor Cyan

# === 1. Servico =====
Section "1. Servico Windows"
$svc = Get-Service WinSysMon -ErrorAction SilentlyContinue
if ($svc) {
    OK "Servico existe - Status: $($svc.Status), StartType: $($svc.StartType)"
    $scq = & sc.exe qc WinSysMon 2>$null | Out-String
    if ($scq -match "BINARY_PATH_NAME\s*:\s*(.+)") { INFO "Binario: $($matches[1].Trim())" }
    $scf = & sc.exe qfailure WinSysMon 2>$null | Out-String
    if ($scf -match "RESET_PERIOD.+?:\s*(\d+)") { INFO "Recovery reset period: $($matches[1])" }
    if ($svc.Status -ne 'Running') { FAIL "Servico NAO esta rodando. Rode: Start-Service WinSysMon" }
} else {
    FAIL "Servico NAO existe neste PC - instalacao nunca rodou ou foi removida"
}

# === 2. Pasta install =====
Section "2. Pasta de instalacao"
if (Test-Path $installDir) {
    OK "Existe: $installDir"
    $itm = Get-Item $installDir -Force
    INFO "Atributos: $($itm.Attributes)"
    try {
        $acl = Get-Acl $installDir
        INFO "Owner: $($acl.Owner)"
        foreach ($a in $acl.Access) { INFO "  ACE: $($a.IdentityReference) $($a.AccessControlType) $($a.FileSystemRights)" }
    } catch { WARN "Sem acesso a ACL: $_" }

    foreach ($f in 'winsysmon.ps1','WinSysMonSvc.exe','sysmon-config.json','sysmon.log') {
        $p = Join-Path $installDir $f
        if (Test-Path $p) { $sz = (Get-Item $p -Force).Length; OK "$f existe ($sz bytes)" }
        else { FAIL "$f AUSENTE" }
    }
} else {
    FAIL "Pasta $installDir NAO existe"
}

# === 3. Config =====
Section "3. sysmon-config.json"
if (Test-Path $cfgPath) {
    try {
        $cfg = Get-Content $cfgPath -Raw | ConvertFrom-Json
        OK "Config lido"
        INFO "  RemoteBlockedAppsPath     = $($cfg.RemoteBlockedAppsPath)"
        INFO "  RemoteBlockedHostsPath    = $($cfg.RemoteBlockedHostsPath)"
        INFO "  RemoteBlockedPoliciesPath = $($cfg.RemoteBlockedPoliciesPath)"
        INFO "  PollInterval              = $($cfg.PollInterval)"
        # Testa acesso ao share
        foreach ($field in 'RemoteBlockedAppsPath','RemoteBlockedHostsPath','RemoteBlockedPoliciesPath') {
            $p = $cfg.$field
            if ($p) {
                if (Test-Path $p -ErrorAction SilentlyContinue) { OK "Share acessivel: $p" }
                else { FAIL "Share INACESSIVEL: $p" }
            }
        }
    } catch { FAIL "Erro parseando config: $_" }
} else { FAIL "sysmon-config.json AUSENTE" }

# === 4. Share roots =====
Section "4. Share roots (hidden + legacy)"
$roots = @("\\srv-105\aaa$","\\srv-105\Sistema de monitoramento\gpo\aaa")
foreach ($r in $roots) {
    $probe = Join-Path $r "service\install-service.ps1"
    if (Test-Path $probe -ErrorAction SilentlyContinue) { OK "Acessivel: $r" }
    else { WARN "Inacessivel: $r" }
}

# === 5. Lista de apps que DEVERIA bloquear para esta maquina =====
Section "5. Lista efetiva de apps bloqueados para $hostname"
if ($cfg -and $cfg.RemoteBlockedAppsPath -and (Test-Path $cfg.RemoteBlockedAppsPath -ErrorAction SilentlyContinue)) {
    try {
        $j = Get-Content $cfg.RemoteBlockedAppsPath -Raw | ConvertFrom-Json
        $global = @(); $machine = @(); $exceptions = @()
        if ($j.Global) { $global = @($j.Global) }
        $hasMachine = $false
        if ($j.Machines -and $j.Machines.PSObject.Properties[$hostname]) {
            $m = @($j.Machines.$hostname)
            if ($m.Count -gt 0) { $machine = $m; $hasMachine = $true }
        }
        if ($j.Exceptions -and $j.Exceptions.PSObject.Properties[$hostname]) {
            $exceptions = @($j.Exceptions.$hostname)
            if ($exceptions.Count -gt 0) { $hasMachine = $true }
        }
        $excSet = @{}; foreach ($x in $exceptions) { $excSet["$x".ToLower()] = $true }
        $effective = @(); $seen = @{}
        foreach ($x in $global)  { $k="$x".ToLower(); if (-not $excSet.ContainsKey($k) -and -not $seen.ContainsKey($k)) { $effective += $x; $seen[$k]=$true } }
        foreach ($x in $machine) { $k="$x".ToLower(); if (-not $seen.ContainsKey($k)) { $effective += $x; $seen[$k]=$true } }
        $src = if ($hasMachine) { "global+machine" } else { "global" }
        OK "Origem: $src | Total efetivo: $($effective.Count) patterns"
        INFO "  Global=$($global.Count) Machines=$($machine.Count) Exceptions=$($exceptions.Count)"
        INFO "  Lista: $($effective -join ', ')"
        if ($effective -contains 'cmd') { OK "'cmd' ESTA na lista efetiva" }
        else { WARN "'cmd' NAO esta na lista efetiva (pode estar em Exceptions)" }
    } catch { FAIL "Erro lendo lista: $_" }
} else { WARN "Nao posso ler lista (share inacessivel)" }

# === 6. Cache local =====
Section "6. Cache local (patterns-cache.json)"
$cache = Join-Path $installDir "patterns-cache.json"
if (Test-Path $cache) {
    try {
        $jc = Get-Content $cache -Raw | ConvertFrom-Json
        OK "Cache existe - timestamp: $($jc.timestamp)"
        INFO "  Patterns: $($jc.patterns -join ', ')"
    } catch { WARN "Erro lendo cache: $_" }
} else { WARN "Cache local AUSENTE (nunca leu do share)" }

# === 7. WMI subscription =====
Section "7. WMI event subscription"
try {
    $ev = Get-EventSubscriber -SourceIdentifier "WinSysMon_ProcWatch" -ErrorAction SilentlyContinue
    if ($ev) { OK "WinSysMon_ProcWatch ativo (Id=$($ev.SubscriptionId))" }
    else { WARN "WinSysMon_ProcWatch INATIVO (mas o servico usa polling de backup)" }
} catch { WARN "$_" }

# === 8. Scheduled Tasks =====
Section "8. Tasks (Watchdog + Guard)"
foreach ($tn in 'WinSysMonWatchdog','WinSysMonGuard') {
    $t = Get-ScheduledTask -TaskName $tn -ErrorAction SilentlyContinue
    if ($t) { OK "$tn existe - State: $($t.State)" }
    else    { FAIL "$tn AUSENTE" }
}

# === 9. WMI persistence (root\subscription) =====
Section "9. WMI persistence (root\subscription)"
try {
    $f = Get-WmiObject -Namespace root\subscription -Class __EventFilter -Filter "Name LIKE 'WinSysMon%'" -ErrorAction SilentlyContinue
    if ($f) { OK "EventFilter: $($f.Name)" } else { WARN "EventFilter WinSysMon ausente" }
    $c = Get-WmiObject -Namespace root\subscription -Class CommandLineEventConsumer -Filter "Name LIKE 'WinSysMon%'" -ErrorAction SilentlyContinue
    if ($c) { OK "Consumer: $($c.Name)" } else { WARN "Consumer WinSysMon ausente" }
} catch { WARN "$_" }

# === 10. Registry Run + Active Setup =====
Section "10. Run + Active Setup"
$runK = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run"
$runV = Get-ItemProperty -Path $runK -Name "WinSysMonHealth" -ErrorAction SilentlyContinue
if ($runV) { OK "Run\WinSysMonHealth = $($runV.WinSysMonHealth)" } else { WARN "Run\WinSysMonHealth ausente" }
$asK  = "HKLM:\SOFTWARE\Microsoft\Active Setup\Installed Components"
$asSub = Get-ChildItem $asK -ErrorAction SilentlyContinue | Where-Object { $_.PSChildName -match "WinSysMon|{[A-F0-9\-]+}" } | Where-Object { (Get-ItemProperty $_.PSPath -ErrorAction SilentlyContinue).'(default)' -like "*WinSysMon*" -or (Get-ItemProperty $_.PSPath -ErrorAction SilentlyContinue).StubPath -like "*WinSysMon*" }
if ($asSub) { OK "Active Setup encontrado" } else { WARN "Active Setup WinSysMon ausente" }

# === 11. ADS backup =====
Section "11. Backup em NTFS ADS"
$ads = "$env:WINDIR\System32\drivers\etc\services:WinSysMonBackup"
try {
    $ab = [System.IO.File]::Exists($ads.Substring(0, $ads.IndexOf(':',3)))
    # Cria um handle e le tamanho
    $fs = [System.IO.File]::Open($ads, 'Open', 'Read', 'Read')
    $len = $fs.Length
    $fs.Close()
    OK "ADS backup existe ($len bytes)"
} catch { WARN "ADS backup inacessivel ou ausente" }

# === 12. Processos que DEVERIAM estar mortos agora =====
Section "12. Processos ativos que deveriam estar bloqueados"
if ($effective -and $effective.Count -gt 0) {
    $killable = @()
    foreach ($pat in $effective) {
        $p = $pat.ToLower().Trim()
        if ($p.Contains("*") -or $p.Contains("?")) {
            $procs = Get-Process -ErrorAction SilentlyContinue | Where-Object { $_.Name.ToLower() -like $p }
        } else {
            $procs = Get-Process -Name $p -ErrorAction SilentlyContinue
            if (-not $procs) { $procs = Get-Process -Name ([System.IO.Path]::GetFileNameWithoutExtension($p)) -ErrorAction SilentlyContinue }
        }
        foreach ($pr in $procs) { $killable += [pscustomobject]@{ Pattern=$pat; Name=$pr.Name; Id=$pr.Id; Session=$pr.SessionId } }
    }
    if ($killable.Count -gt 0) {
        FAIL "$($killable.Count) processo(s) RODANDO que deveriam estar bloqueados:"
        $killable | Format-Table -AutoSize | Out-String | Write-Host
        Write-Host "  (Se serviço está Running, ele deveria ter matado. Verifique sysmon.log)" -ForegroundColor Yellow
    } else { OK "Nenhum processo bloqueado rodando agora" }
}

# === 13. Antivirus detectado =====
Section "13. Antivirus instalado"
try {
    $avs = Get-CimInstance -Namespace root\SecurityCenter2 -ClassName AntiVirusProduct -ErrorAction Stop
    foreach ($av in $avs) { INFO "$($av.displayName) - state=$($av.productState)" }
} catch { WARN "Nao conseguiu ler AntiVirusProduct: $_" }
$wsec = Get-Service -Name "FSMA","F-Secure*","WithSecure*" -ErrorAction SilentlyContinue
if ($wsec) { foreach ($s in $wsec) { INFO "$($s.Name) - $($s.Status)" } }

# === 14. Ultimas 30 linhas do log =====
Section "14. Ultimas 30 linhas do sysmon.log"
if (Test-Path $logPath) {
    Get-Content $logPath -Tail 30 | ForEach-Object { Write-Host "    $_" -ForegroundColor DarkGray }
} else { FAIL "sysmon.log ausente" }

Section "15. Ultimas 15 linhas do install.log"
if (Test-Path $installLog) {
    Get-Content $installLog -Tail 15 | ForEach-Object { Write-Host "    $_" -ForegroundColor DarkGray }
} else { WARN "install.log ausente" }

Write-Host ""
Write-Host "===== RESUMO =====" -ForegroundColor Cyan
Write-Host "1. Se servico nao existe / parado - rode INSTALAR.BAT como admin"
Write-Host "2. Se share inacessivel - verifique DNS/SMB para \\srv-105"
Write-Host "3. Se processos listados em [12] - reinstalar (serviço travou)"
Write-Host "4. Se AV WithSecure alertando hostsfile - precisa allow-list no console central"
Write-Host ""
