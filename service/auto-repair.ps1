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

# ---------- 2. PARAR SERVICO WINSYSMON (evita race com reescrita) ----------
$svcExists = $null -ne (Get-Service -Name WinSysMon -ErrorAction SilentlyContinue)
if ($svcExists) {
    Write-RepairLog 'Parando servico WinSysMon para evitar reescrita do hosts...'
    & sc.exe stop WinSysMon 2>&1 | Out-Null
    # Aguarda ate 10s para o servico parar completamente
    $waited = 0
    while ($waited -lt 10) {
        $s = Get-Service -Name WinSysMon -ErrorAction SilentlyContinue
        if (-not $s -or $s.Status -eq 'Stopped') { break }
        Start-Sleep -Seconds 1
        $waited++
    }
    Write-RepairLog "Servico WinSysMon parado (aguardou ${waited}s)"
}

# ---------- 3. FORCAR firewall-dns EM sysmon-config.json ----------
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

# ---------- 4. LIMPAR HOSTS (marcadores REAIS + failsafe de 0.0.0.0) ----------
$hostsFile = "$env:WINDIR\System32\drivers\etc\hosts"
if (Test-Path $hostsFile) {
    try {
        $originalContent = Get-Content $hostsFile -Raw -ErrorAction Stop
        $content = $originalContent

        # Marcadores REAIS usados por winsysmon.ps1 (consulte $script:HostsBeginMarker):
        #   # === WINSYSMON-BEGIN (do not edit manually) ===
        #   # === WINSYSMON-END ===
        $pattern1 = '(?ms)#\s*=+\s*WINSYSMON-BEGIN.*?WINSYSMON-END\s*=+\s*\r?\n?'
        $content = [regex]::Replace($content, $pattern1, '')

        # Failsafe 1: marcadores antigos em outros formatos (case-insensitive)
        $pattern2 = '(?mis)#\s*=+\s*WinSysMon\s+Begin.*?WinSysMon\s+End\s*=+\s*\r?\n?'
        $content = [regex]::Replace($content, $pattern2, '')

        # Failsafe 2: remove QUALQUER linha comecando com 0.0.0.0 (nosso IP de bloqueio).
        # Isso garante que mesmo sem marcadores, o arquivo fica limpo do ponto de vista
        # do WithSecure (que detecta '0.0.0.0 <dominio>' como "Redirected hosts file").
        $lines = $content -split "`r?`n"
        $cleanLines = $lines | Where-Object { $_ -notmatch '^\s*0\.0\.0\.0\s+\S' }
        $content = ($cleanLines -join "`r`n").TrimEnd() + "`r`n"

        if ($content -ne $originalContent) {
            try { (Get-Item $hostsFile).IsReadOnly = $false } catch {}
            # takeown + icacls como failsafe se ACL estiver travada
            try { & takeown.exe /F $hostsFile /A 2>&1 | Out-Null } catch {}
            try { & icacls.exe $hostsFile /grant '*S-1-5-18:F' '*S-1-5-32-544:F' /C 2>&1 | Out-Null } catch {}

            [System.IO.File]::WriteAllText($hostsFile, $content, (New-Object System.Text.ASCIIEncoding))
            & ipconfig.exe /flushdns 2>&1 | Out-Null
            Write-RepairLog 'hosts limpo (marcadores + linhas 0.0.0.0 removidas)'
        } else {
            Write-RepairLog 'hosts ja esta limpo - nenhuma alteracao necessaria'
        }
    } catch {
        Write-RepairLog "Erro limpando hosts: $_" 'ERROR'
    }
}

# ---------- 5. LIMPAR REGRA FIREWALL ANTIGA (se houver nome conflitante) ----------
# O servico recriara a regra no proximo ciclo; esta apenas garante estado limpo.
try {
    $existing = Get-NetFirewallRule -DisplayName 'WinSysMon_BlockDomainIPs' -ErrorAction SilentlyContinue
    if ($existing) { Write-RepairLog "Regra firewall 'WinSysMon_BlockDomainIPs' ja existe - OK" }
} catch {}

# ---------- 6. REINICIAR SERVICO (apos config corrigido) ----------
# NOTA: Nao reiniciamos aqui. Deixamos o PARAGPOAA continuar o fluxo
# (check de versao, reinstall se necessario, start no final).
# Se o servico existe e nao houve reinstall, PARAGPOAA faz sc.exe start.

Write-RepairLog "=== auto-repair.ps1 concluido ==="
exit 0
