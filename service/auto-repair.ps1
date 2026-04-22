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

# ---------- 2. VERIFICAR SE HOSTS PRECISA LIMPEZA (apenas para log / decidir stop-svc) ----------
# NOTA v2.7.4: NAO escrevemos no hosts quando WS esta presente. Qualquer write
# dispara On-Access Scanner do WithSecure, causando alerta "Redirected hosts file".
# O proprio WithSecure faz a "desinfeccao" automaticamente (remove 0.0.0.0).
# Nosso trabalho aqui e apenas GARANTIR que nao escreveremos novos 0.0.0.0
# (via config=firewall-dns + HARD-BLOCK no agente).
$hostsFile = "$env:WINDIR\System32\drivers\etc\hosts"
$hostsHasResidue = $false
if (Test-Path $hostsFile) {
    try {
        $hostsCheck = Get-Content $hostsFile -Raw -ErrorAction Stop
        if ($hostsCheck -match '(?ms)#\s*=+\s*WINSYSMON-BEGIN' -or
            $hostsCheck -match '(?m)^\s*0\.0\.0\.0\s+\S') {
            $hostsHasResidue = $true
        }
    } catch {
        Write-RepairLog "Erro lendo hosts: $_" 'WARN'
    }
}

if ($hostsHasResidue) {
    Write-RepairLog 'hosts tem residuo 0.0.0.0 - WithSecure ira desinfetar (nao escrevemos)' 'WARN'
}

# ---------- 3. NAO PARAR SERVICO (v2.7.4) ----------
# v2.7.3 parava o servico pra limpar hosts. v2.7.4 nunca limpa hosts,
# entao o servico continua Running. Mantem calc bloqueada, etc.
$svcExists = $null -ne (Get-Service -Name WinSysMon -ErrorAction SilentlyContinue)

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

# ---------- 5. NAO TOCAR NO HOSTS (v2.7.4) ----------
# O proprio WithSecure cuida disso via "desinfeccao" automatica.
# Se escrevermos aqui (mesmo para limpar), On-Access Scanner alerta.

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
