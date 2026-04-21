$p = '\\srv-105\Sistema de monitoramento\gpo\aaa\service\blocked-apps.json'
$out = 'c:\gpo\scripts\fix-pc12.txt'
$lines = @()
$lines += "exists=$(Test-Path $p)"
if (Test-Path $p) {
    $raw = Get-Content $p -Raw -Encoding UTF8
    $lines += "---RAW---"
    $lines += $raw
    try {
        $j = $raw | ConvertFrom-Json
        $g = @(); if ($j.Global) { $g = @($j.Global) }
        $lines += "Global=$($g -join ',')"
        $m = @(); if ($j.Machines) { $m = @($j.Machines.PSObject.Properties.Name) }
        $lines += "Machines=$($m -join ',')"
        $e = @(); if ($j.Exceptions) { $e = @($j.Exceptions.PSObject.Properties.Name) }
        $lines += "Exceptions=$($e -join ',')"
    } catch { $lines += "parse-err=$($_.Exception.Message)" }
}
[System.IO.File]::WriteAllLines($out, $lines, [System.Text.UTF8Encoding]::new($false))
