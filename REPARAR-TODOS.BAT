@echo off
REM =====================================================================
REM  REPARAR-TODOS.BAT - Reinstala WinSysMon em uma lista de PCs
REM
REM  Uso:
REM    REPARAR-TODOS.BAT                  -> usa lista padrao abaixo
REM    REPARAR-TODOS.BAT PC-200 PC-201    -> usa lista passada
REM
REM  Requer: credencial admin de dominio (sera solicitada)
REM =====================================================================
setlocal enabledelayedexpansion

if "%~1"=="" (
    set "PCS=PC-200 PC-201 PC-202 PC-203 PC-204 PC-205"
) else (
    set "PCS=%*"
)

echo.
echo ======================================================================
echo  WinSysMon - REPARAR EM LOTE
echo ======================================================================
echo  PCs alvo: %PCS%
echo.

powershell -NoProfile -ExecutionPolicy Bypass -Command "& { $pcs = '%PCS%' -split ' ' | Where-Object { $_ }; $cred = Get-Credential -Message 'Admin de dominio para reparar PCs em lote'; $share = '\\srv-105\Sistema de monitoramento\gpo\aaa\service'; $cmd = 'cmd.exe /c powershell.exe -NoProfile -ExecutionPolicy Bypass -File \"' + $share + '\install-service.ps1\" -SharePath \"' + $share + '\blocked-apps.json\" -HostsSharePath \"' + $share + '\blocked-hosts.json\" -PoliciesSharePath \"' + $share + '\blocked-policies.json\" > C:\Windows\Temp\wsm-reinstall.log 2>&1'; $results = @(); foreach ($pc in $pcs) { Write-Host ''; Write-Host ('=== ' + $pc + ' ===') -ForegroundColor Cyan; $ping = Test-Connection $pc -Count 1 -Quiet -ErrorAction SilentlyContinue; if (-not $ping) { Write-Host ('  [OFFLINE] nao responde ao ping') -ForegroundColor Red; $results += [pscustomobject]@{PC=$pc;Status='OFFLINE';Service=''}; continue }; try { $r = Invoke-WmiMethod -ComputerName $pc -Credential $cred -Class Win32_Process -Name Create -ArgumentList $cmd -ErrorAction Stop; if ($r.ReturnValue -eq 0) { Write-Host ('  [LANCADO] PID=' + $r.ProcessId + ' - instalando...') -ForegroundColor Green; $results += [pscustomobject]@{PC=$pc;Status='LANCADO';Service='?';PID=$r.ProcessId} } else { Write-Host ('  [FALHOU] WMI ReturnValue=' + $r.ReturnValue) -ForegroundColor Red; $results += [pscustomobject]@{PC=$pc;Status='ERRO-WMI';Service='';PID=''} } } catch { Write-Host ('  [ERRO] ' + $_.Exception.Message) -ForegroundColor Red; $results += [pscustomobject]@{PC=$pc;Status='ERRO';Service='';PID=''} } }; Write-Host ''; Write-Host '==============================================' -ForegroundColor Yellow; Write-Host ' Aguardando 90s para instalacoes terminarem...' -ForegroundColor Yellow; Start-Sleep 90; Write-Host ''; Write-Host '==============================================' -ForegroundColor Cyan; Write-Host ' VERIFICACAO FINAL' -ForegroundColor Cyan; Write-Host '==============================================' -ForegroundColor Cyan; foreach ($row in $results) { if ($row.Status -eq 'LANCADO') { try { $s = Get-WmiObject -ComputerName $row.PC -Credential $cred -Class Win32_Service -Filter ('Name=''WinSysMon''') -ErrorAction Stop; if ($s) { $row.Service = $s.State } else { $row.Service = 'AUSENTE' } } catch { $row.Service = 'ERRO-CHECK' } } }; $results | Format-Table -AutoSize; $ok = @($results | Where-Object { $_.Service -eq 'Running' }).Count; $total = $results.Count; Write-Host ''; Write-Host ('Resultado: ' + $ok + '/' + $total + ' PCs com servico RUNNING') -ForegroundColor $(if ($ok -eq $total) { 'Green' } else { 'Yellow' }); Write-Host ''; Write-Host 'OBS: mesmo onde servico nao subir, o bloqueio IFEO ja foi aplicado' -ForegroundColor Yellow; Write-Host '     via registro durante a instalacao. Apps bloqueados nao abrem.' -ForegroundColor Yellow }"

pause
