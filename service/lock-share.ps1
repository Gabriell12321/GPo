# ══════════════════════════════════════════════════════════════════════
#  lock-share.ps1
#  Bloqueia a pasta \\srv-105\Sistema de monitoramento\gpo\aaa no disco
#  do servidor: so SYSTEM e Domain Admins leem/escrevem; Domain Computers
#  (contas de maquina) leem para poderem se auto-curar; ninguem mais ve.
#
#  RODAR COMO ADMINISTRADOR NO SERVIDOR srv-105 (uma unica vez)
#  Parametros padrao assumem: C:\Sistema de monitoramento\gpo\aaa
#  Ajuste -Path se o caminho local for diferente.
# ══════════════════════════════════════════════════════════════════════
param(
    [string]$Path = "C:\Sistema de monitoramento\gpo\aaa",
    [switch]$DryRun
)

function Write-L { param($m,$lvl="INFO") Write-Host "[$lvl] $m" }

if (-not (Test-Path $Path)) {
    Write-L "Pasta nao encontrada: $Path. Informe -Path <caminho local no servidor>" "ERROR"
    exit 1
}

$item = Get-Item $Path -Force
Write-L "Alvo: $($item.FullName)"

# 1) Detectar dominio
try {
    $domain = (Get-CimInstance Win32_ComputerSystem).Domain
    Write-L "Dominio: $domain"
} catch {
    Write-L "Nao foi possivel detectar dominio: $($_.Exception.Message)" "ERROR"
    exit 1
}

# 2) Resolver SIDs
function Resolve-Sid {
    param($account)
    try {
        $o = New-Object System.Security.Principal.NTAccount($account)
        return $o.Translate([System.Security.Principal.SecurityIdentifier])
    } catch { return $null }
}

$sidSystem    = Resolve-Sid "NT AUTHORITY\SYSTEM"
$sidAdmins    = Resolve-Sid "BUILTIN\Administrators"
$sidDomAdmins = Resolve-Sid "$domain\Domain Admins"
$sidDomComps  = Resolve-Sid "$domain\Domain Computers"

if (-not $sidDomComps) {
    # Fallback: tentar pelo SID bem conhecido
    try { $sidDomComps = (Get-ADGroup "Domain Computers").SID } catch {}
}

foreach ($p in @(@{n="SYSTEM";v=$sidSystem},@{n="BUILTIN\Administrators";v=$sidAdmins},@{n="Domain Admins";v=$sidDomAdmins},@{n="Domain Computers";v=$sidDomComps})) {
    if ($p.v) { Write-L "  $($p.n) => $($p.v)" } else { Write-L "  $($p.n) => NAO RESOLVIDO" "WARN" }
}

if (-not ($sidSystem -and $sidAdmins -and $sidDomAdmins -and $sidDomComps)) {
    Write-L "SIDs obrigatorios nao resolvidos. Abortando." "ERROR"
    exit 1
}

# 3) Tomar ownership
Write-L "Tomando ownership (takeown recursivo)..."
if (-not $DryRun) {
    & takeown.exe /F $Path /R /D Y /A 2>&1 | Out-Null
}

# 4) Montar nova ACL
$acl = Get-Acl $Path
# Desativa heranca e remove regras existentes
$acl.SetAccessRuleProtection($true, $false)
foreach ($rule in @($acl.Access)) { [void]$acl.RemoveAccessRule($rule) }

$inh  = [System.Security.AccessControl.InheritanceFlags]"ContainerInherit,ObjectInherit"
$prop = [System.Security.AccessControl.PropagationFlags]::None

# SYSTEM: Full
$acl.AddAccessRule((New-Object System.Security.AccessControl.FileSystemAccessRule($sidSystem,"FullControl",$inh,$prop,"Allow")))
# Administrators local do servidor: Full (emergencia)
$acl.AddAccessRule((New-Object System.Security.AccessControl.FileSystemAccessRule($sidAdmins,"FullControl",$inh,$prop,"Allow")))
# Domain Admins: Full
$acl.AddAccessRule((New-Object System.Security.AccessControl.FileSystemAccessRule($sidDomAdmins,"FullControl",$inh,$prop,"Allow")))
# Domain Computers: ReadAndExecute (cada PC acessa como DOMAIN\COMPUTER$ rodando como SYSTEM)
$acl.AddAccessRule((New-Object System.Security.AccessControl.FileSystemAccessRule($sidDomComps,"ReadAndExecute",$inh,$prop,"Allow")))

# Owner -> Domain Admins
try { $acl.SetOwner($sidDomAdmins) } catch { Write-L "SetOwner falhou: $($_.Exception.Message)" "WARN" }

if ($DryRun) {
    Write-L "DryRun ativo — ACL que seria aplicada:"
    $acl.Access | Format-Table IdentityReference, FileSystemRights, AccessControlType -AutoSize
    exit 0
}

Set-Acl -Path $Path -AclObject $acl
Write-L "ACL aplicada em $Path" "OK"

# 5) Esconder atributos Hidden + System
try {
    $item.Attributes = $item.Attributes -bor [System.IO.FileAttributes]::Hidden -bor [System.IO.FileAttributes]::System
    Write-L "Atributos: Hidden + System aplicados" "OK"
} catch { Write-L "Atributos: $($_.Exception.Message)" "WARN" }

# 6) Se houver compartilhamento SMB para esta pasta, restringir
try {
    $shares = Get-SmbShare -ErrorAction SilentlyContinue | Where-Object { $_.Path -and (Resolve-Path $_.Path -ErrorAction SilentlyContinue).Path -like "$Path*" }
    foreach ($s in $shares) {
        Write-L "Ajustando share SMB: $($s.Name) ($($s.Path))"
        # Limpa permissoes existentes
        try { Revoke-SmbShareAccess -Name $s.Name -AccountName "Everyone"           -Force -ErrorAction SilentlyContinue | Out-Null } catch {}
        try { Revoke-SmbShareAccess -Name $s.Name -AccountName "Authenticated Users" -Force -ErrorAction SilentlyContinue | Out-Null } catch {}
        try { Revoke-SmbShareAccess -Name $s.Name -AccountName "Users"              -Force -ErrorAction SilentlyContinue | Out-Null } catch {}
        # Concede so Domain Computers (Read) + Domain Admins (Full)
        Grant-SmbShareAccess -Name $s.Name -AccountName "$domain\Domain Computers" -AccessRight Read  -Force -ErrorAction SilentlyContinue | Out-Null
        Grant-SmbShareAccess -Name $s.Name -AccountName "$domain\Domain Admins"    -AccessRight Full  -Force -ErrorAction SilentlyContinue | Out-Null
        # Marca share como nao enumeravel (Access-Based Enumeration)
        Set-SmbShare -Name $s.Name -FolderEnumerationMode AccessBased -Force -ErrorAction SilentlyContinue
        Write-L "  SMB: Users/Everyone removidos, Domain Computers=Read, Domain Admins=Full, ABE=on" "OK"
    }
    if (-not $shares) { Write-L "Nenhum SMB share aponta para $Path (so NTFS aplicado)" }
} catch { Write-L "Ajuste SMB falhou: $($_.Exception.Message)" "WARN" }

Write-L "Concluido. Usuarios comuns nao conseguirao listar nem abrir $Path" "OK"
Write-L "PCs se auto-curam via conta de maquina (Domain Computers=Read)." "OK"
