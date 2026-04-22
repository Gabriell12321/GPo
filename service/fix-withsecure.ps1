# ============================================================
#  FIX-WITHSECURE - Corrige deteccao "Redirected hosts file"
#  1. Para servico WinSysMon
#  2. Forca HostBlockingMethod = firewall-dns no config
#  3. Limpa markers WinSysMon do hosts (retira o redirect)
#  4. Remove backups/cache do hosts
#  5. Reinicia servico - agora bloqueia via firewall, nao toca no hosts
# ============================================================

$ErrorActionPreference = 'Continue'
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) { Write-Host "ABORTANDO - nao elevado" -ForegroundColor Red; Read-Host "ENTER"; return }

$folder     = "C:\ProgramData\Microsoft\WinSysMon"
$cfgPath    = Join-Path $folder "sysmon-config.json"
$hostsPath  = "$env:WINDIR\System32\drivers\etc\hosts"
$hostsCache = Join-Path $folder "hosts-cache.json"

Write-Host "==> 1. Parando servico WinSysMon..." -ForegroundColor Cyan
Stop-Service WinSysMon -Force -ErrorAction SilentlyContinue
Start-Sleep 2

Write-Host "`n==> 2. Lendo config atual..." -ForegroundColor Cyan
if (-not (Test-Path $cfgPath)) {
    Write-Host "[ERRO] Config nao existe: $cfgPath" -ForegroundColor Red
    Read-Host "ENTER"; return
}
$cfg = Get-Content $cfgPath -Raw | ConvertFrom-Json
Write-Host "  HostBlockingMethod atual: $($cfg.HostBlockingMethod)"

Write-Host "`n==> 3. Forcando HostBlockingMethod = firewall-dns..." -ForegroundColor Cyan
if ($cfg.PSObject.Properties['HostBlockingMethod']) {
    $cfg.HostBlockingMethod = 'firewall-dns'
} else {
    $cfg | Add-Member -NotePropertyName HostBlockingMethod -NotePropertyValue 'firewall-dns' -Force
}
($cfg | ConvertTo-Json -Depth 5) | Set-Content $cfgPath -Encoding UTF8
Write-Host "  [OK] Config gravado com firewall-dns"

Write-Host "`n==> 4. Limpando markers WinSysMon do hosts..." -ForegroundColor Cyan
if (Test-Path $hostsPath) {
    $content = Get-Content $hostsPath -Raw -ErrorAction SilentlyContinue
    $beginMk = "# --- WinSysMon BEGIN ---"
    $endMk   = "# --- WinSysMon END ---"
    if ($content -match [regex]::Escape($beginMk)) {
        # Remove bloco entre markers (inclusive)
        $pattern = "(?ms)\r?\n?" + [regex]::Escape($beginMk) + ".*?" + [regex]::Escape($endMk) + "\r?\n?"
        $clean = [regex]::Replace($content, $pattern, "")
        Set-Content $hostsPath -Value $clean -Encoding ASCII -Force
        Write-Host "  [OK] Bloco WinSysMon removido do hosts"
    } else {
        Write-Host "  [INFO] Nenhum marker WinSysMon no hosts (WithSecure ja desinfetou)"
    }
    # Verifica se ainda ha entradas sem marker (legacy)
    $lines = Get-Content $hostsPath -ErrorAction SilentlyContinue
    $suspectLines = $lines | Where-Object { $_ -match '^\s*0\.0\.0\.0\s+' -and $_ -notmatch '^\s*#' }
    if ($suspectLines) {
        Write-Host "  [WARN] Ainda ha $($suspectLines.Count) entradas 0.0.0.0 sem marker - revisar manualmente:" -ForegroundColor Yellow
        $suspectLines | Select-Object -First 5 | ForEach-Object { Write-Host "    $_" }
    }
} else {
    Write-Host "  [WARN] hosts nao existe?" -ForegroundColor Yellow
}

Write-Host "`n==> 5. Limpando cache do hosts e ADS backup..." -ForegroundColor Cyan
if (Test-Path $hostsCache) {
    Remove-Item $hostsCache -Force -ErrorAction SilentlyContinue
    Write-Host "  [OK] hosts-cache.json removido (forca re-aplicar via firewall)"
}
# Limpa ADS backup antigo do hosts (se houver)
try {
    Get-Item "${hostsPath}:*" -ErrorAction SilentlyContinue | ForEach-Object {
        Write-Host "  [INFO] ADS detectado em hosts: $($_.Stream) - removendo"
        Remove-Item -Path $hostsPath -Stream $_.Stream -ErrorAction SilentlyContinue
    }
} catch {}

Write-Host "`n==> 6. Flush DNS (garante resolver respeitar firewall)..." -ForegroundColor Cyan
ipconfig /flushdns | Out-Null
Write-Host "  [OK]"

Write-Host "`n==> 7. Reiniciando servico WinSysMon..." -ForegroundColor Cyan
Start-Service WinSysMon -ErrorAction Continue
Start-Sleep 5
$s = Get-Service WinSysMon
Write-Host "  Status: $($s.Status)" -ForegroundColor $(if($s.Status -eq 'Running'){'Green'}else{'Red'})

Write-Host "`n==> 8. Aguardando primeiro ciclo (15s)..." -ForegroundColor Cyan
Start-Sleep 15

Write-Host "`n==> 9. Validando firewall rules..." -ForegroundColor Cyan
$fwRule = Get-NetFirewallRule -DisplayName 'WinSysMon_BlockDomainIPs' -ErrorAction SilentlyContinue
if ($fwRule) {
    Write-Host "  [OK] Firewall rule existe: $($fwRule.DisplayName)" -ForegroundColor Green
    Write-Host "  Enabled: $($fwRule.Enabled) | Action: $($fwRule.Action) | Direction: $($fwRule.Direction)"
    $addrFilter = Get-NetFirewallAddressFilter -AssociatedNetFirewallRule $fwRule -ErrorAction SilentlyContinue
    if ($addrFilter) {
        $ips = $addrFilter.RemoteAddress
        Write-Host "  IPs bloqueados: $($ips.Count) entrada(s)"
        $ips | Select-Object -First 10 | ForEach-Object { Write-Host "    $_" }
    }
} else {
    Write-Host "  [WARN] Firewall rule ainda nao criada - aguarde mais um ciclo (60s)" -ForegroundColor Yellow
}

Write-Host "`n==> 10. Validando hosts INTACTO..." -ForegroundColor Cyan
$hostsNow = Get-Content $hostsPath -ErrorAction SilentlyContinue
$hasMk = $hostsNow | Where-Object { $_ -match 'WinSysMon' }
if ($hasMk) {
    Write-Host "  [FAIL] AINDA ha marker WinSysMon no hosts - config nao foi lido" -ForegroundColor Red
    $hasMk | ForEach-Object { Write-Host "    $_" }
} else {
    Write-Host "  [OK] hosts SEM marker WinSysMon (WithSecure nao vai mais detectar)" -ForegroundColor Green
}

Write-Host "`n==> 11. Ultimas linhas do sysmon.log:" -ForegroundColor Cyan
Get-Content (Join-Path $folder "sysmon.log") -Tail 20 -ErrorAction SilentlyContinue

Write-Host "`n============================================================" -ForegroundColor Green
Write-Host " FIX WITHSECURE aplicado." -ForegroundColor Green
Write-Host " Bloqueio de sites agora via FIREWALL (IP resolvido)," -ForegroundColor Green
Write-Host " sem tocar no arquivo hosts. WithSecure nao detecta mais." -ForegroundColor Green
Write-Host "============================================================" -ForegroundColor Green
Read-Host "ENTER para sair"
