@echo off
REM =====================================================================
REM  REPARAR-PC.BAT - Reinstala WinSysMon em estacao remota
REM  Uso: REPARAR-PC.BAT PC-200
REM  Requer: credencial admin de dominio (sera solicitada)
REM =====================================================================
setlocal
if "%~1"=="" (
    echo.
    echo Uso: REPARAR-PC.BAT ^<HOSTNAME^>
    echo Ex : REPARAR-PC.BAT PC-200
    echo.
    exit /b 1
)
set "PC=%~1"

powershell -NoProfile -ExecutionPolicy Bypass -Command "& { $pc='%PC%'; $cred=Get-Credential -Message \"Admin de dominio para consertar $pc\"; $sb = { param($shareRoot) $ErrorActionPreference='SilentlyContinue'; $svc = Get-Service WinSysMon -ErrorAction SilentlyContinue; if ($svc) { Stop-Service WinSysMon -Force -ErrorAction SilentlyContinue; Start-Sleep 2; sc.exe delete WinSysMon | Out-Null; Start-Sleep 2 }; $roots = @('\\srv-105\aaa$','\\srv-105\Sistema de monitoramento\gpo\aaa'); $root = $null; foreach ($r in $roots) { if (Test-Path (Join-Path $r 'service\install-service.ps1')) { $root=$r; break } }; if (-not $root) { return 'SHARE INACESSIVEL' }; $inst = Join-Path $root 'service\install-service.ps1'; $json = Join-Path $root 'service\blocked-apps.json'; $hj   = Join-Path $root 'service\blocked-hosts.json'; $pj   = Join-Path $root 'service\blocked-policies.json'; & powershell -NoProfile -ExecutionPolicy Bypass -File $inst -SharePath $json -HostsSharePath $hj -PoliciesSharePath $pj 2>&1; Start-Sleep 3; $svc2 = Get-Service WinSysMon -ErrorAction SilentlyContinue; if ($svc2) { 'SERVICO: ' + $svc2.Status } else { 'SERVICO: AUSENTE' } }; Write-Host ''; Write-Host '=== Executando em' $pc '===' -ForegroundColor Cyan; try { $r = Invoke-Command -ComputerName $pc -Credential $cred -ScriptBlock $sb -ErrorAction Stop; $r | ForEach-Object { Write-Host $_ } } catch { Write-Host ('ERRO: ' + $_.Exception.Message) -ForegroundColor Red; Write-Host 'PSRemoting pode estar desabilitado. Tentando via WMI...' -ForegroundColor Yellow; try { $share = '\\srv-105\Sistema de monitoramento\gpo\aaa\service\install-service.ps1'; $cmd = 'powershell.exe -NoProfile -ExecutionPolicy Bypass -File \"' + $share + '\" -SharePath \"\\srv-105\Sistema de monitoramento\gpo\aaa\service\blocked-apps.json\" -HostsSharePath \"\\srv-105\Sistema de monitoramento\gpo\aaa\service\blocked-hosts.json\" -PoliciesSharePath \"\\srv-105\Sistema de monitoramento\gpo\aaa\service\blocked-policies.json\"'; $proc = Invoke-WmiMethod -ComputerName $pc -Credential $cred -Class Win32_Process -Name Create -ArgumentList $cmd; Write-Host ('WMI ReturnValue: ' + $proc.ReturnValue + ' PID: ' + $proc.ProcessId); Start-Sleep 10; $s = Get-Service -ComputerName $pc -Name WinSysMon -ErrorAction SilentlyContinue; if ($s) { Write-Host ('SERVICO FINAL: ' + $s.Status) -ForegroundColor Green } else { Write-Host 'SERVICO AINDA AUSENTE' -ForegroundColor Red } } catch { Write-Host ('ERRO WMI: ' + $_.Exception.Message) -ForegroundColor Red } } }"
pause
