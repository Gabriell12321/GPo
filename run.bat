@echo off
cd /d "%~dp0"
if not exist "bin\gpo.hl" (
    echo Compilando...
    haxe build.hxml
)
hl bin\gpo.hl
