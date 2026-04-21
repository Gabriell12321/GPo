# discover-mikrotik.ps1
# Descobre IP de MikroTik usando 8 metodos diferentes.
# Uso: powershell -ExecutionPolicy Bypass -File .\discover-mikrotik.ps1 [-Quick] [-Deep]
#   -Quick : pula scan de subnet completo (metodos 5/6)
#   -Deep  : faz scan em TODAS as interfaces (metodo 6)

[CmdletBinding()]
param(
    [switch]$Quick,
    [switch]$Deep,
    [int]$ThreadCount = 50,
    [int]$PortTimeoutMs = 300
)

$ErrorActionPreference = 'SilentlyContinue'
$found = [System.Collections.Generic.List[object]]::new()

function Add-Candidate {
    param([string]$Ip, [string]$Method, [string]$Evidence = "")
    if (-not $Ip) { return }
    if ($script:found | Where-Object { $_.Ip -eq $Ip -and $_.Method -eq $Method }) { return }
    $script:found.Add([pscustomobject]@{
        Ip       = $Ip
        Method   = $Method
        Evidence = $Evidence
        Time     = (Get-Date).ToString('HH:mm:ss')
    })
    Write-Host "  [+] $Ip via $Method $(if($Evidence){"- $Evidence"})" -ForegroundColor Green
}

function Test-TcpPort {
    param([string]$Ip, [int]$Port, [int]$TimeoutMs = 300)
    $client = New-Object System.Net.Sockets.TcpClient
    try {
        $iar = $client.BeginConnect($Ip, $Port, $null, $null)
        $ok = $iar.AsyncWaitHandle.WaitOne($TimeoutMs, $false)
        if ($ok -and $client.Connected) { $client.EndConnect($iar); return $true }
        return $false
    } catch { return $false } finally { $client.Close() }
}

# ============================================================
# Metodo 1 - MNDP (MikroTik Neighbor Discovery Protocol)
# ============================================================
function Discover-MNDP {
    Write-Host "[1] MNDP (UDP 5678) - aguardando broadcast 15s..." -ForegroundColor Cyan
    $udp = New-Object System.Net.Sockets.UdpClient
    try {
        $udp.Client.ReceiveTimeout = 15000
        $udp.ExclusiveAddressUse = $false
        $udp.Client.SetSocketOption([System.Net.Sockets.SocketOptionLevel]::Socket, [System.Net.Sockets.SocketOptionName]::ReuseAddress, $true)
        $ep = New-Object System.Net.IPEndPoint ([System.Net.IPAddress]::Any), 5678
        $udp.Client.Bind($ep)
        $remote = New-Object System.Net.IPEndPoint ([System.Net.IPAddress]::Any), 0
        $deadline = (Get-Date).AddSeconds(15)
        while ((Get-Date) -lt $deadline) {
            try {
                $data = $udp.Receive([ref]$remote)
                if ($data.Length -gt 0) {
                    Add-Candidate $remote.Address.ToString() "MNDP" "bytes=$($data.Length)"
                }
            } catch { break }
        }
    } catch {
        Write-Host "    (falha bind UDP 5678 - porta ocupada ou firewall)" -ForegroundColor DarkYellow
    } finally { $udp.Close() }
}

# ============================================================
# Metodo 2 - Gateway padrao + porta 8291
# ============================================================
function Discover-Gateway {
    Write-Host "[2] Gateway padrao..." -ForegroundColor Cyan
    $gws = Get-NetRoute -DestinationPrefix '0.0.0.0/0' | Where-Object { $_.NextHop -ne '0.0.0.0' } | Select-Object -ExpandProperty NextHop -Unique
    foreach ($gw in $gws) {
        $isMk = Test-TcpPort $gw 8291 500
        if ($isMk) { Add-Candidate $gw "Gateway+8291" "WinBox aberto" }
        else { Write-Host "    GW $gw sem 8291 aberto" -ForegroundColor DarkGray }
    }
}

# ============================================================
# Metodo 3 - Traceroute hop-a-hop
# ============================================================
function Discover-Traceroute {
    Write-Host "[3] Traceroute 8.8.8.8 (ate 15 hops)..." -ForegroundColor Cyan
    try {
        $hops = Test-NetConnection -TraceRoute -ComputerName 8.8.8.8 -Hops 15 -WarningAction SilentlyContinue
        foreach ($h in $hops.TraceRoute) {
            if (-not $h -or $h -eq '0.0.0.0') { continue }
            if (Test-TcpPort $h 8291 400) { Add-Candidate $h "Traceroute+8291" "hop" }
        }
    } catch { Write-Host "    traceroute falhou: $($_.Exception.Message)" -ForegroundColor DarkYellow }
}

# ============================================================
# Metodo 4 - MAC OUI MikroTik na tabela ARP
# ============================================================
function Discover-ArpOUI {
    Write-Host "[4] Tabela ARP (OUI MikroTik)..." -ForegroundColor Cyan
    $ouis = @('00-0C-42','2C-C8-1B','4C-5E-0C','6C-3B-6B','B8-69-F4','CC-2D-E0','D4-CA-6D','DC-2C-6E','48-8F-5A','E4-8D-8C','74-4D-28','18-FD-74','64-D1-54','08-55-31','78-9A-18','C4-AD-34')
    $arp = Get-NetNeighbor -AddressFamily IPv4 -ErrorAction SilentlyContinue | Where-Object { $_.State -ne 'Unreachable' -and $_.LinkLayerAddress }
    foreach ($n in $arp) {
        $mac = $n.LinkLayerAddress.ToUpper()
        $prefix = ($mac -replace ':','-').Substring(0,8)
        if ($ouis -contains $prefix) {
            Add-Candidate $n.IPAddress "ARP-OUI" "MAC=$mac"
        }
    }
}

# ============================================================
# Metodo 5/6 - Scan porta 8291 em subnet(s)
# ============================================================
function Get-SubnetIps {
    param([string]$Ip, [int]$Prefix)
    if ($Prefix -lt 20 -or $Prefix -gt 30) { return @() } # evita /16 gigante
    $ipBytes = ([System.Net.IPAddress]::Parse($Ip)).GetAddressBytes()
    [Array]::Reverse($ipBytes)
    $ipInt = [BitConverter]::ToUInt32($ipBytes,0)
    $mask = [uint32]([math]::Pow(2,32) - [math]::Pow(2,32-$Prefix))
    $net = $ipInt -band $mask
    $bcast = $net -bor (-bnot $mask -band 0xFFFFFFFF)
    $list = @()
    for ($i = $net + 1; $i -lt $bcast; $i++) {
        $b = [BitConverter]::GetBytes([uint32]$i)
        [Array]::Reverse($b)
        $list += ([System.Net.IPAddress]::new($b)).ToString()
    }
    return $list
}

function Scan-Subnet {
    param([string[]]$Ips, [int]$Port = 8291, [string]$Label = "Scan")
    if (-not $Ips -or $Ips.Count -eq 0) { return }
    Write-Host "    escaneando $($Ips.Count) IPs porta $Port..." -ForegroundColor DarkGray
    $jobs = @()
    $pool = [runspacefactory]::CreateRunspacePool(1, $ThreadCount)
    $pool.Open()
    foreach ($ip in $Ips) {
        $ps = [powershell]::Create().AddScript({
            param($ip,$port,$to)
            $c = New-Object System.Net.Sockets.TcpClient
            try {
                $iar = $c.BeginConnect($ip,$port,$null,$null)
                if ($iar.AsyncWaitHandle.WaitOne($to,$false) -and $c.Connected) {
                    $c.EndConnect($iar); return $ip
                }
            } catch {} finally { $c.Close() }
            return $null
        }).AddArgument($ip).AddArgument($Port).AddArgument($PortTimeoutMs)
        $ps.RunspacePool = $pool
        $jobs += [pscustomobject]@{ PS=$ps; Handle=$ps.BeginInvoke(); Ip=$ip }
    }
    foreach ($j in $jobs) {
        $r = $j.PS.EndInvoke($j.Handle)
        $j.PS.Dispose()
        if ($r) { Add-Candidate $r "$Label+$Port" "tcp aberto" }
    }
    $pool.Close(); $pool.Dispose()
}

function Discover-LocalSubnet {
    Write-Host "[5] Scan subnet do adaptador principal (porta 8291)..." -ForegroundColor Cyan
    $gw = Get-NetRoute -DestinationPrefix '0.0.0.0/0' | Where-Object { $_.NextHop -ne '0.0.0.0' } | Select-Object -First 1
    if (-not $gw) { Write-Host "    sem gateway"; return }
    $ifIdx = $gw.InterfaceIndex
    $ip = Get-NetIPAddress -InterfaceIndex $ifIdx -AddressFamily IPv4 | Select-Object -First 1
    if (-not $ip) { return }
    $ips = Get-SubnetIps -Ip $ip.IPAddress -Prefix $ip.PrefixLength
    Scan-Subnet -Ips $ips -Port 8291 -Label "Subnet"
}

function Discover-AllSubnets {
    Write-Host "[6] Scan em TODAS interfaces ativas..." -ForegroundColor Cyan
    $addrs = Get-NetIPAddress -AddressFamily IPv4 | Where-Object {
        $_.PrefixOrigin -ne 'WellKnown' -and $_.IPAddress -notmatch '^(127\.|169\.254\.)'
    }
    foreach ($a in $addrs) {
        if ($a.PrefixLength -lt 22) { Write-Host "    pulando $($a.IPAddress)/$($a.PrefixLength) (rede muito grande)" -ForegroundColor DarkYellow; continue }
        Write-Host "    interface $($a.IPAddress)/$($a.PrefixLength)" -ForegroundColor DarkCyan
        $ips = Get-SubnetIps -Ip $a.IPAddress -Prefix $a.PrefixLength
        Scan-Subnet -Ips $ips -Port 8291 -Label "IfScan"
    }
}

# ============================================================
# Metodo 7 - Banner SSH
# ============================================================
function Test-SshBanner {
    param([string]$Ip)
    $c = New-Object System.Net.Sockets.TcpClient
    try {
        $iar = $c.BeginConnect($Ip,22,$null,$null)
        if (-not ($iar.AsyncWaitHandle.WaitOne(500,$false) -and $c.Connected)) { return $null }
        $c.EndConnect($iar)
        $s = $c.GetStream(); $s.ReadTimeout = 1500
        $buf = New-Object byte[] 128
        Start-Sleep -Milliseconds 200
        $n = $s.Read($buf,0,$buf.Length)
        if ($n -gt 0) { return [Text.Encoding]::ASCII.GetString($buf,0,$n) }
    } catch {} finally { $c.Close() }
    return $null
}

function Discover-SshBanner {
    Write-Host "[7] Banner SSH nos candidatos..." -ForegroundColor Cyan
    $targets = @($script:found | Select-Object -ExpandProperty Ip -Unique)
    foreach ($ip in $targets) {
        $b = Test-SshBanner $ip
        if ($b -and $b -match 'ROSSSH|Mikrotik|RouterOS') {
            Add-Candidate $ip "SSH-Banner" ($b.Trim())
        }
    }
}

# ============================================================
# Metodo 8 - SNMP sysDescr
# ============================================================
function Discover-Snmp {
    Write-Host "[8] SNMP sysDescr (community=public)..." -ForegroundColor Cyan
    try { $snmp = New-Object -ComObject olePrn.OleSNMP } catch {
        Write-Host "    olePrn.OleSNMP indisponivel - pulando" -ForegroundColor DarkYellow; return
    }
    $targets = @($script:found | Select-Object -ExpandProperty Ip -Unique)
    foreach ($ip in $targets) {
        try {
            $snmp.Open($ip,'public',2,1000)
            $v = $snmp.Get('1.3.6.1.2.1.1.1.0')
            if ($v -match 'RouterOS|Mikrotik') { Add-Candidate $ip "SNMP" $v }
            $snmp.Close()
        } catch {}
    }
}

# ============================================================
# Main
# ============================================================
Write-Host "`n==== MikroTik Discovery ====" -ForegroundColor Yellow
Discover-MNDP
Discover-Gateway
Discover-Traceroute
Discover-ArpOUI
if (-not $Quick) { Discover-LocalSubnet }
if ($Deep)       { Discover-AllSubnets }
Discover-SshBanner
Discover-Snmp

Write-Host "`n==== Resultado ====" -ForegroundColor Yellow
if ($found.Count -eq 0) {
    Write-Host "Nenhum MikroTik encontrado." -ForegroundColor Red
    Write-Host "Tente novamente com -Deep para varrer todas as subnets." -ForegroundColor DarkYellow
} else {
    $grouped = $found | Group-Object Ip | ForEach-Object {
        [pscustomobject]@{
            Ip      = $_.Name
            Hits    = $_.Count
            Metodos = ($_.Group.Method -join ', ')
        }
    } | Sort-Object Hits -Descending
    $grouped | Format-Table -AutoSize
    $best = $grouped | Select-Object -First 1
    Write-Host "`n>>> Melhor candidato: $($best.Ip) (detectado por $($best.Hits) metodo[s])" -ForegroundColor Green
}
