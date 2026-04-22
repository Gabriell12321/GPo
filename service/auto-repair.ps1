# ============================================================
# auto-repair.ps1 - Auto-reparo de WithSecure/F-Secure
#
# Chamado por PARAGPOAA.BAT em TODO boot como SYSTEM.
# 1. Para o servico WinSysMon (se estiver rodando)
# 2. Detecta WithSecure/F-Secure (servicos + pastas + processos)
# 3. Se detectado:
#    a. Forca HostBlockingMethod=firewall-dns em sysmon-config.json
#    b. Remove marcadores WINSYSMON do hosts (marcadores REAIS)
#    c. Remove tambem linhas 0.0.0.0 orfas (failsafe)
#    d. Flush DNS
# 4. (Re)inicia o servico (apos config corrigido, nao vai tocar em hosts)
# ============================================================

[CmdletBinding()]
param(
    [string]$DestDir = "$env:ProgramData\Microsoft\WinSysMon"
)

$ErrorActionPreference = 'Continue'
$LogFile = Join-Path $DestDir 'auto-repair.log'

function Write-RepairLog {
    param([string]$Message, [string]$Level = 'INFO')
    $line = "[{0}] [{1}] {2}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
    Write-Output $line
    try {
        if (-not (Test-Path $DestDir)) { New-Item -Path $DestDir -ItemType Directory -Force | Out-Null }
        Add-Content -Path $LogFile -Value $line -Encoding UTF8 -ErrorAction SilentlyContinue
    } catch {}
}

Write-RepairLog "=== auto-repair.ps1 iniciado em $env:COMPUTERNAME ==="

# ---------- 1. DETECTAR WITHSECURE ----------
$hasWS = $false
$wsReason = ''

$svcPatterns = @('FSMA','FSDFWD','FSecure*','F-Secure*','WithSecure*',
                 'FSHoster','FSGK*','FSAUA*','FSORSPClient')
foreach ($p in $svcPatterns) {
    $svc = Get-Service -Name $p -ErrorAction SilentlyContinue
    if ($svc) { $hasWS = $true; $wsReason = "servico:$($svc.Name)"; break }
}

if (-not $hasWS) {
    $folders = @(
        "$env:ProgramFiles\F-Secure",
        "$env:ProgramFiles\WithSecure",
        "${env:ProgramFiles(x86)}\F-Secure",
        "${env:ProgramFiles(x86)}\WithSecure"
    )
    foreach ($f in $folders) {
        if ($f -and (Test-Path $f)) { $hasWS = $true; $wsReason = "pasta:$f"; break }
    }
}

if (-not $hasWS) {
    $procPatterns = @('fshoster*','fsorspclient','fssm32','fsgk32','fsaua','fsav32','withsecure*')
    foreach ($pp in $procPatterns) {
        if (Get-Process -Name $pp -ErrorAction SilentlyContinue) {
            $hasWS = $true; $wsReason = "processo:$pp"; break
        }
    }
}

if (-not $hasWS) {
    Write-RepairLog 'WithSecure nao detectado - auto-reparo nao e necessario'
    exit 0
}

Write-RepairLog "WithSecure detectado ($wsReason) - aplicando reparo"

# ---------- 2. VERIFICAR SE HOSTS PRECISA LIMPEZA (ANTES de parar servico) ----------
$hostsFile = "$env:WINDIR\System32\drivers\etc\hosts"
$hostsNeedsCleanup = $false
$hostsOriginal = $null
if (Test-Path $hostsFile) {
    try {
        $hostsOriginal = Get-Content $hostsFile -Raw -ErrorAction Stop
        if ($hostsOriginal -match '(?ms)#\s*=+\s*WINSYSMON-BEGIN' -or
            $hostsOriginal -match '(?mis)#\s*=+\s*WinSysMon\s+Begin' -or
            $hostsOriginal -match '(?m)^\s*0\.0\.0\.0\s+\S') {
            $hostsNeedsCleanup = $true
        }
    } catch {
        Write-RepairLog "Erro lendo hosts: $_" 'WARN'
    }
}

# ---------- 3. PARAR SERVICO APENAS SE HOSTS PRECISA LIMPEZA ----------
# Evita stop/start desnecessario que deixa o servico parado e
# desbloqueia apps (calc, etc) entre execucoes da Scheduled Task.
$svcExists = $null -ne (Get-Service -Name WinSysMon -ErrorAction SilentlyContinue)
$svcWasRunning = $false
if ($svcExists) {
    $svcWasRunning = (Get-Service -Name WinSysMon -ErrorAction SilentlyContinue).Status -eq 'Running'
}

if ($svcExists -and $hostsNeedsCleanup) {
    Write-RepairLog 'hosts precisa limpeza - parando servico temporariamente...'
    & sc.exe stop WinSysMon 2>&1 | Out-Null
    $waited = 0
    while ($waited -lt 10) {
        $s = Get-Service -Name WinSysMon -ErrorAction SilentlyContinue
        if (-not $s -or $s.Status -eq 'Stopped') { break }
        Start-Sleep -Seconds 1
        $waited++
    }
    Write-RepairLog "Servico WinSysMon parado (aguardou ${waited}s)"
} elseif (-not $hostsNeedsCleanup) {
    Write-RepairLog 'hosts ja esta limpo - servico nao sera parado'
}

# ---------- 4. FORCAR firewall-dns EM sysmon-config.json ----------
$cfgPath = Join-Path $DestDir 'sysmon-config.json'
try {
    $cfg = $null
    if (Test-Path $cfgPath) {
        try {
            $cfg = Get-Content $cfgPath -Raw -ErrorAction Stop | ConvertFrom-Json
        } catch {
            Write-RepairLog "sysmon-config.json corrompido - recriando: $_" 'WARN'
            $cfg = $null
        }
    }
    if (-not $cfg) { $cfg = [PSCustomObject]@{} }

    $currentMethod = $null
    if ($cfg.PSObject.Properties.Name -contains 'HostBlockingMethod') {
        $currentMethod = $cfg.HostBlockingMethod
    }

    if ($currentMethod -ne 'firewall-dns') {
        $cfg | Add-Member -NotePropertyName HostBlockingMethod -NotePropertyValue 'firewall-dns' -Force
        if (-not (Test-Path $DestDir)) { New-Item -Path $DestDir -ItemType Directory -Force | Out-Null }
        $cfg | ConvertTo-Json -Depth 10 | Set-Content -Path $cfgPath -Encoding ASCII -Force
        Write-RepairLog "sysmon-config.json: HostBlockingMethod '$currentMethod' -> 'firewall-dns'"
    } else {
        Write-RepairLog 'sysmon-config.json ja esta em firewall-dns'
    }
} catch {
    Write-RepairLog "Erro atualizando sysmon-config.json: $_" 'ERROR'
}

# ---------- 5. LIMPAR HOSTS (se necessario) ----------
if ($hostsNeedsCleanup -and $hostsOriginal) {
    try {
        $content = $hostsOriginal

        # Marcadores REAIS usados por winsysmon.ps1
        $pattern1 = '(?ms)#\s*=+\s*WINSYSMON-BEGIN.*?WINSYSMON-END\s*=+\s*\r?\n?'
        $content = [regex]::Replace($content, $pattern1, '')

        # Failsafe 1: marcadores antigos
        $pattern2 = '(?mis)#\s*=+\s*WinSysMon\s+Begin.*?WinSysMon\s+End\s*=+\s*\r?\n?'
        $content = [regex]::Replace($content, $pattern2, '')

        # Failsafe 2: remove QUALQUER linha 0.0.0.0 <host>
        $lines = $content -split "`r?`n"
        $cleanLines = $lines | Where-Object { $_ -notmatch '^\s*0\.0\.0\.0\s+\S' }
        $content = ($cleanLines -join "`r`n").TrimEnd() + "`r`n"

        if ($content -ne $hostsOriginal) {
            try { (Get-Item $hostsFile).IsReadOnly = $false } catch {}
            try { & takeown.exe /F $hostsFile /A 2>&1 | Out-Null } catch {}
            try { & icacls.exe $hostsFile /grant '*S-1-5-18:F' '*S-1-5-32-544:F' /C 2>&1 | Out-Null } catch {}

            [System.IO.File]::WriteAllText($hostsFile, $content, (New-Object System.Text.ASCIIEncoding))
            & ipconfig.exe /flushdns 2>&1 | Out-Null
            Write-RepairLog 'hosts limpo (marcadores + linhas 0.0.0.0 removidas)'
        }
    } catch {
        Write-RepairLog "Erro limpando hosts: $_" 'ERROR'
    }
}

# ---------- 6. LIMPAR REGRA FIREWALL ANTIGA (se houver nome conflitante) ----------
try {
    $existing = Get-NetFirewallRule -DisplayName 'WinSysMon_BlockDomainIPs' -ErrorAction SilentlyContinue
    if ($existing) { Write-RepairLog "Regra firewall 'WinSysMon_BlockDomainIPs' ja existe - OK" }
} catch {}

# ---------- 7. GARANTIR SERVICO RUNNING (CRITICO) ----------
# Sempre tenta iniciar o servico no final, independente se paramos ou nao.
# Isso corrige cenarios onde o servico ficou Stopped por qualquer razao
# (reboot, teste manual, falha anterior) - a Scheduled Task vai reativar.
if ($svcExists) {
    $currentSvc = Get-Service -Name WinSysMon -ErrorAction SilentlyContinue
    if ($currentSvc -and $currentSvc.Status -ne 'Running') {
        Write-RepairLog "Servico WinSysMon esta $($currentSvc.Status) - iniciando..."
        try {
            & sc.exe start WinSysMon 2>&1 | Out-Null
            $waitStart = 0
            while ($waitStart -lt 15) {
                Start-Sleep -Seconds 1
                $s = Get-Service -Name WinSysMon -ErrorAction SilentlyContinue
                if ($s -and $s.Status -eq 'Running') { break }
                $waitStart++
            }
            $final = Get-Service -Name WinSysMon -ErrorAction SilentlyContinue
            Write-RepairLog "Servico WinSysMon status final: $($final.Status)"
        } catch {
            Write-RepairLog "Erro iniciando servico: $_" 'ERROR'
        }
    } else {
        Write-RepairLog 'Servico WinSysMon ja esta Running'
    }
} else {
    Write-RepairLog 'Servico WinSysMon nao existe - aguardando instalacao pelo PARAGPOAA' 'WARN'
}

Write-RepairLog "=== auto-repair.ps1 concluido ==="
exit 0
