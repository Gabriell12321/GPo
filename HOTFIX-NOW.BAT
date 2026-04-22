@echo off
REM ============================================================
REM  HOTFIX-NOW.BAT v1.0 - Aplica fix v2.7.2 SEM reboot
REM
REM  Uso no PC-200 (duplo-clique, aceita UAC):
REM    \\srv-105\Sistema de monitoramento\gpo\aaa\HOTFIX-NOW.BAT
REM
REM  Faz:
REM  1. Copia winsysmon.ps1 v2.7.2 + auto-repair.ps1 do share
REM  2. Roda auto-repair.ps1 (para svc, limpa hosts, forca firewall-dns)
REM  3. Cria Scheduled Task WinSysMonAutoRepair (2 min)
REM  4. Reinicia WinSysMon
REM ============================================================
setlocal EnableDelayedExpansion

REM Auto-elevacao UAC
net session >nul 2>&1
if %errorlevel% neq 0 (
    echo Solicitando privilegios de administrador...
    powershell -NoProfile -Command "Start-Process -FilePath '%~f0' -Verb RunAs"
    exit /b
)

title HOTFIX-NOW v2.7.2 - %COMPUTERNAME%
echo.
echo  ====================================================
echo   HOTFIX-NOW v2.7.2 em %COMPUTERNAME%
echo  ====================================================
echo.

set "SHARE=\\srv-105\Sistema de monitoramento\gpo\aaa"
set "DEST=C:\ProgramData\Microsoft\WinSysMon"

if not exist "%DEST%" mkdir "%DEST%" 2>nul

echo  [1/5] Copiando winsysmon.ps1 v2.7.2 do share...
if not exist "%SHARE%\service\winsysmon.ps1" (
    echo  [ERRO] Share inacessivel: %SHARE%
    pause
    exit /b 1
)
takeown /F "%DEST%\winsysmon.ps1" /A >nul 2>&1
icacls "%DEST%\winsysmon.ps1" /grant "*S-1-5-32-544:F" /C >nul 2>&1
copy /Y "%SHARE%\service\winsysmon.ps1" "%DEST%\winsysmon.ps1" >nul
copy /Y "%SHARE%\service\auto-repair.ps1" "%DEST%\auto-repair.ps1" >nul
copy /Y "%SHARE%\service\version.txt" "%DEST%\version.txt" >nul

echo  [2/5] Executando auto-repair.ps1 (para svc + limpa hosts + firewall-dns)...
powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -File "%DEST%\auto-repair.ps1" -DestDir "%DEST%"

echo.
echo  [3/5] Criando Scheduled Task WinSysMonAutoRepair (2 min)...
schtasks /delete /tn "WinSysMonAutoRepair" /f >nul 2>&1
schtasks /create /tn "WinSysMonAutoRepair" /ru SYSTEM /rl HIGHEST /f ^
    /sc MINUTE /mo 2 ^
    /tr "powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -File \"%DEST%\auto-repair.ps1\" -DestDir \"%DEST%\""
if %errorlevel%==0 (
    echo  [OK] Tarefa criada - rodara a cada 2 minutos
) else (
    echo  [ERRO] Falha criando tarefa
)

echo.
echo  [4/5] Reiniciando servico WinSysMon...
sc.exe stop WinSysMon >nul 2>&1
timeout /t 3 /nobreak >nul
sc.exe start WinSysMon
timeout /t 5 /nobreak >nul

REM Confirma que servico esta rodando (critico - se nao, bloqueios nao funcionam)
sc.exe query WinSysMon | findstr /I "RUNNING" >nul
if errorlevel 1 (
    echo  [WARN] Servico nao iniciou na primeira tentativa - tentando novamente...
    timeout /t 3 /nobreak >nul
    sc.exe start WinSysMon
    timeout /t 5 /nobreak >nul
)

echo.
echo  [5/5] Verificando estado final...
sc.exe query WinSysMon | findstr /I "STATE"
echo.
echo  --- hosts content (ultimas 20 linhas) ---
powershell -NoProfile -Command "Get-Content C:\Windows\System32\drivers\etc\hosts | Select-Object -Last 20"

echo.
echo  ====================================================
echo   HOTFIX aplicado. WithSecure NAO deve mais alertar.
echo   Tarefa WinSysMonAutoRepair limpara a cada 2 min.
echo  ====================================================
pause
endlocal
