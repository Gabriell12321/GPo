@echo off
REM ============================================================
REM  HOTFIX-WITHSECURE.BAT - Corrige PC que ainda gera alerta
REM  "Redirected hosts file" do WithSecure.
REM
REM  Duplo-clique: auto-eleva, atualiza WinSysMon para v2.7.0+
REM  e aplica fix-withsecure.
REM ============================================================
setlocal EnableDelayedExpansion

REM Auto-elevacao UAC
net session >nul 2>&1
if %errorlevel% neq 0 (
    powershell -NoProfile -Command "Start-Process -FilePath '%~f0' -Verb RunAs"
    exit /b
)

title HOTFIX WithSecure - %COMPUTERNAME%

echo.
echo  ====================================================
echo   HOTFIX WITHSECURE - %COMPUTERNAME%
echo  ====================================================
echo.

set "SHARE=\\srv-105\Sistema de monitoramento\gpo\aaa"
set "DEST=C:\ProgramData\Microsoft\WinSysMon"

echo  [1/4] Parando servico WinSysMon...
sc stop WinSysMon >nul 2>&1
timeout /t 3 /nobreak >nul

echo  [2/4] Atualizando winsysmon.ps1 para v2.7.0 do share...
if exist "%SHARE%\service\winsysmon.ps1" (
    takeown /F "%DEST%\winsysmon.ps1" /A >nul 2>&1
    icacls "%DEST%\winsysmon.ps1" /grant "*S-1-5-32-544:F" /C >nul 2>&1
    copy /Y "%SHARE%\service\winsysmon.ps1" "%DEST%\winsysmon.ps1" >nul
    echo  [OK] winsysmon.ps1 atualizado
) else (
    echo  [ERRO] Share inacessivel: %SHARE%
    pause
    exit /b 1
)

copy /Y "%SHARE%\service\version.txt" "%DEST%\version.txt" >nul 2>&1

echo  [3/4] Aplicando fix-withsecure...
if exist "%SHARE%\service\fix-withsecure.ps1" (
    powershell -NoProfile -ExecutionPolicy Bypass -File "%SHARE%\service\fix-withsecure.ps1"
) else (
    echo  [WARN] fix-withsecure.ps1 nao encontrado
)

echo.
echo  [4/4] Verificando status final...
timeout /t 5 /nobreak >nul
for /f "tokens=3 delims=: " %%s in ('sc query WinSysMon ^| findstr "STATE"') do set "STATE=%%s"
echo  Estado do servico: !STATE!

echo.
echo  ====================================================
echo   Aguarde 2-3 minutos. WithSecure deve parar de alertar.
echo  ====================================================
pause
endlocal
