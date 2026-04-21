param(
    [string]$Share = '\\srv-105\Sistema de monitoramento\gpo\aaa\service\blocked-apps.json',
    [string]$Out   = 'c:\gpo\scripts\fix-pc12.out.txt'
)
$ErrorActionPreference = 'Continue'
"[$(Get-Date -Format o)] start" | Out-File $Out -Encoding utf8
"Share=$Share exists=$(Test-Path $Share)" | Out-File $Out -Append

if (-not (Test-Path $Share)) { "ABORT: share nao encontrado" | Out-File $Out -Append; exit 1 }

$raw = Get-Content $Share -Raw
"RAW:" | Out-File $Out -Append
$raw | Out-File $Out -Append

$j = $raw | ConvertFrom-Json
$global = @()
if ($j.Global) { $global = @($j.Global) }
"GLOBAL count=$($global.Count)" | Out-File $Out -Append
$global | ForEach-Object { "  G: $_" } | Out-File $Out -Append

$pcs = @()
if ($j.Machines) { $pcs = @($j.Machines.PSObject.Properties.Name) }
"MACHINES=$($pcs -join ', ')" | Out-File $Out -Append

$excPcs = @()
if ($j.Exceptions) { $excPcs = @($j.Exceptions.PSObject.Properties.Name) }
"EXCEPTIONS=$($excPcs -join ', ')" | Out-File $Out -Append

# Encontrar PC com "12" no nome
$target = $pcs | Where-Object { $_ -match '12' } | Select-Object -First 1
if (-not $target) {
    # tenta por LDAP (AD)
    try {
        $ad = New-Object DirectoryServices.DirectorySearcher
        $ad.Filter = "(&(objectCategory=computer)(cn=*12*))"
        $r = $ad.FindAll()
        $target = $r | ForEach-Object { $_.Properties['cn'][0] } | Select-Object -First 1
        "LDAP-found=$target" | Out-File $Out -Append
    } catch { "LDAP-err=$($_.Exception.Message)" | Out-File $Out -Append }
}
"TARGET=$target" | Out-File $Out -Append
