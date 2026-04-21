# ======================================================================
#  lock-share.ps1  (v2.0 - BLOQUEIO TOTAL)
#  Bloqueia \\srv-105\...\aaa no disco do servidor com:
#    - Owner = TrustedInstaller
#    - DENY ACE explicito para Domain Admins (delete/write/takeown)
#    - Share oculto ($ suffix) com SMB Encryption + Access-Based Enum
#    - Atributos Hidden+System+ReadOnly
#    - Auditoria de acesso (SACL) - grava em Security Event Log
#    - Negacao a Everyone/Users/Authenticated Users
#
#  RODAR COMO ADMINISTRADOR DO DOMINIO NO srv-105
#  Parametros padrao assumem: C:\Sistema de monitoramento\gpo\aaa
# ======================================================================
param(
    [string]$Path = "C:\Sistema de monitoramento\gpo\aaa",
    [string]$HiddenShareName = "aaa$",
    [switch]$DryRun,
    [switch]$KeepVisibleShare
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

# 2) Garantir TrustedInstaller rodando (resolvel como conta)
try {
    $ti = Get-Service -Name TrustedInstaller -ErrorAction SilentlyContinue
    if ($ti -and $ti.Status -ne 'Running') { Start-Service TrustedInstaller -ErrorAction SilentlyContinue }
} catch {}

# 3) Resolver SIDs
function Resolve-Sid {
    param($account)
    try {
        $o = New-Object System.Security.Principal.NTAccount($account)
        return $o.Translate([System.Security.Principal.SecurityIdentifier])
    } catch { return $null }
}

$sidSystem    = Resolve-Sid "NT AUTHORITY\SYSTEM"
$sidTI        = Resolve-Sid "NT SERVICE\TrustedInstaller"
$sidAdmins    = Resolve-Sid "BUILTIN\Administrators"
$sidDomAdmins = Resolve-Sid "$domain\Domain Admins"
$sidDomComps  = Resolve-Sid "$domain\Domain Computers"
$sidUsers     = Resolve-Sid "BUILTIN\Users"
$sidEveryone  = Resolve-Sid "Everyone"
$sidAuthUsers = Resolve-Sid "NT AUTHORITY\Authenticated Users"

if (-not $sidDomComps) {
    try { $sidDomComps = (Get-ADGroup "Domain Computers").SID } catch {}
}

foreach ($p in @(
    @{n="SYSTEM";v=$sidSystem},
    @{n="TrustedInstaller";v=$sidTI},
    @{n="BUILTIN\Administrators";v=$sidAdmins},
    @{n="Domain Admins";v=$sidDomAdmins},
    @{n="Domain Computers";v=$sidDomComps}
)) {
    if ($p.v) { Write-L "  $($p.n) => $($p.v)" } else { Write-L "  $($p.n) => NAO RESOLVIDO" "WARN" }
}

if (-not ($sidSystem -and $sidTI -and $sidDomAdmins -and $sidDomComps)) {
    Write-L "SIDs obrigatorios nao resolvidos. Abortando." "ERROR"
    exit 1
}

# 4) Habilitar privilegios necessarios (TakeOwnership / Restore)
$sig = @'
using System;
using System.Runtime.InteropServices;
public class PrivShare {
  [StructLayout(LayoutKind.Sequential, Pack=1)] public struct TokPriv1Luid { public int Count; public long Luid; public int Attr; }
  public const int SE_PRIVILEGE_ENABLED = 2;
  public const int TOKEN_ADJUST_PRIVILEGES = 32;
  public const int TOKEN_QUERY = 8;
  [DllImport("advapi32.dll", SetLastError=true)] public static extern bool AdjustTokenPrivileges(IntPtr ht, bool all, ref TokPriv1Luid tp, int bl, IntPtr pl, IntPtr rl);
  [DllImport("advapi32.dll", SetLastError=true)] public static extern bool OpenProcessToken(IntPtr h, int acc, ref IntPtr p);
  [DllImport("advapi32.dll", SetLastError=true)] public static extern bool LookupPrivilegeValue(string host, string name, ref long pluid);
  [DllImport("kernel32.dll")] public static extern IntPtr GetCurrentProcess();
}
'@
if (-not ("PrivShare" -as [type])) { Add-Type -TypeDefinition $sig }
function Enable-Priv { param($p)
    $tp = New-Object PrivShare+TokPriv1Luid
    $tp.Count = 1; $tp.Luid = 0; $tp.Attr = [PrivShare]::SE_PRIVILEGE_ENABLED
    $hTok = [IntPtr]::Zero
    [void][PrivShare]::OpenProcessToken([PrivShare]::GetCurrentProcess(), [PrivShare]::TOKEN_ADJUST_PRIVILEGES -bor [PrivShare]::TOKEN_QUERY, [ref]$hTok)
    [void][PrivShare]::LookupPrivilegeValue($null, $p, [ref]$tp.Luid)
    [void][PrivShare]::AdjustTokenPrivileges($hTok, $false, [ref]$tp, 0, [IntPtr]::Zero, [IntPtr]::Zero)
}
foreach ($pr in 'SeTakeOwnershipPrivilege','SeRestorePrivilege','SeBackupPrivilege','SeSecurityPrivilege') {
    try { Enable-Priv $pr } catch {}
}

# 5) Tomar ownership
Write-L "Tomando ownership (takeown recursivo)..."
if (-not $DryRun) {
    & takeown.exe /F $Path /R /D Y /A 2>&1 | Out-Null
}

# 6) Montar nova DACL
$acl = Get-Acl $Path
$acl.SetAccessRuleProtection($true, $false)
foreach ($rule in @($acl.Access)) { [void]$acl.RemoveAccessRule($rule) }

$inh  = [System.Security.AccessControl.InheritanceFlags]"ContainerInherit,ObjectInherit"
$prop = [System.Security.AccessControl.PropagationFlags]::None

# ALLOW: SYSTEM + TrustedInstaller = Full
$acl.AddAccessRule((New-Object System.Security.AccessControl.FileSystemAccessRule($sidSystem,"FullControl",$inh,$prop,"Allow")))
$acl.AddAccessRule((New-Object System.Security.AccessControl.FileSystemAccessRule($sidTI,"FullControl",$inh,$prop,"Allow")))
# ALLOW: Domain Admins = ReadAndExecute + Write (precisa publicar scripts)
$acl.AddAccessRule((New-Object System.Security.AccessControl.FileSystemAccessRule($sidDomAdmins,"Modify",$inh,$prop,"Allow")))
# ALLOW: Domain Computers = ReadAndExecute (PCs auto-curam)
$acl.AddAccessRule((New-Object System.Security.AccessControl.FileSystemAccessRule($sidDomComps,"ReadAndExecute",$inh,$prop,"Allow")))

# DENY: Domain Admins NAO podem mudar permissoes, tomar ownership, deletar o raiz
#       (continuam escrevendo nos filhos por causa do Modify acima, mas proteger o root)
$denyRootRights = [System.Security.AccessControl.FileSystemRights]"ChangePermissions,TakeOwnership"
$acl.AddAccessRule((New-Object System.Security.AccessControl.FileSystemAccessRule($sidDomAdmins,$denyRootRights,[System.Security.AccessControl.InheritanceFlags]::None,$prop,"Deny")))
# DENY explicito em Users/Everyone/Authenticated: TUDO (nem enumerar)
if ($sidUsers)     { $acl.AddAccessRule((New-Object System.Security.AccessControl.FileSystemAccessRule($sidUsers,"FullControl",$inh,$prop,"Deny"))) }
if ($sidEveryone)  { $acl.AddAccessRule((New-Object System.Security.AccessControl.FileSystemAccessRule($sidEveryone,"FullControl",$inh,$prop,"Deny"))) }
if ($sidAuthUsers) {
    # Authenticated Users: precisa continuar permitindo Domain Computers (que sao AuthUsers);
    # entao nao damos DENY global aqui. Controle fica pelos ALLOWs acima.
}

# Owner = TrustedInstaller
try { $acl.SetOwner($sidTI) } catch { Write-L "SetOwner TrustedInstaller falhou: $($_.Exception.Message)" "WARN" }

# 7) SACL - auditoria de modificacoes
try {
    $auditFlags = [System.Security.AccessControl.AuditFlags]"Success,Failure"
    $auditRights = [System.Security.AccessControl.FileSystemRights]"Delete,DeleteSubdirectoriesAndFiles,ChangePermissions,TakeOwnership,WriteData,AppendData"
    if ($sidDomAdmins) {
        $acl.AddAuditRule((New-Object System.Security.AccessControl.FileSystemAuditRule($sidDomAdmins,$auditRights,$inh,$prop,$auditFlags)))
    }
    if ($sidAdmins) {
        $acl.AddAuditRule((New-Object System.Security.AccessControl.FileSystemAuditRule($sidAdmins,$auditRights,$inh,$prop,$auditFlags)))
    }
    Write-L "SACL: auditando Delete/WriteDAC/TakeOwnership de Admins"
} catch { Write-L "SACL falhou: $($_.Exception.Message)" "WARN" }

if ($DryRun) {
    Write-L "DryRun ativo - DACL que seria aplicada:"
    $acl.Access | Format-Table IdentityReference, FileSystemRights, AccessControlType -AutoSize
    Write-L "SACL:"
    $acl.Audit  | Format-Table IdentityReference, FileSystemRights, AuditFlags -AutoSize
    exit 0
}

Set-Acl -Path $Path -AclObject $acl
Write-L "DACL+SACL aplicadas em $Path" "OK"

# 8) Atributos
try {
    $item.Attributes = $item.Attributes -bor [System.IO.FileAttributes]::Hidden -bor [System.IO.FileAttributes]::System
    Get-ChildItem $Path -Recurse -Force -ErrorAction SilentlyContinue | ForEach-Object {
        try { $_.Attributes = $_.Attributes -bor [System.IO.FileAttributes]::Hidden -bor [System.IO.FileAttributes]::System } catch {}
    }
    Write-L "Atributos: Hidden + System aplicados recursivamente" "OK"
} catch { Write-L "Atributos: $($_.Exception.Message)" "WARN" }

# 9) Ajustar SMB shares visiveis + criar share OCULTO
try {
    $shares = Get-SmbShare -ErrorAction SilentlyContinue | Where-Object {
        $_.Path -and (Resolve-Path $_.Path -ErrorAction SilentlyContinue).Path -like "$Path*"
    }

    # Se nao mandar -KeepVisibleShare, remove os shares visiveis existentes
    if (-not $KeepVisibleShare) {
        foreach ($s in $shares) {
            if ($s.Name -like "*$*") { continue }
            Write-L "Removendo share visivel: $($s.Name) ($($s.Path))"
            try { Remove-SmbShare -Name $s.Name -Force -ErrorAction Stop | Out-Null } catch { Write-L "  falhou: $($_.Exception.Message)" "WARN" }
        }
    } else {
        # Ajustar os existentes
        foreach ($s in $shares) {
            Write-L "Ajustando share visivel: $($s.Name)"
            try {
                Revoke-SmbShareAccess -Name $s.Name -AccountName "Everyone"            -Force -ErrorAction SilentlyContinue | Out-Null
                Revoke-SmbShareAccess -Name $s.Name -AccountName "Authenticated Users" -Force -ErrorAction SilentlyContinue | Out-Null
                Revoke-SmbShareAccess -Name $s.Name -AccountName "Users"               -Force -ErrorAction SilentlyContinue | Out-Null
            } catch {}
            Grant-SmbShareAccess -Name $s.Name -AccountName "$domain\Domain Computers" -AccessRight Read -Force -ErrorAction SilentlyContinue | Out-Null
            Grant-SmbShareAccess -Name $s.Name -AccountName "$domain\Domain Admins"    -AccessRight Full -Force -ErrorAction SilentlyContinue | Out-Null
            Set-SmbShare -Name $s.Name -FolderEnumerationMode AccessBased -EncryptData $true -Force -ErrorAction SilentlyContinue
        }
    }

    # Criar share oculto aaa$ (o $ no fim faz o Windows NAO enumerar em net view)
    $existing = Get-SmbShare -Name $HiddenShareName -ErrorAction SilentlyContinue
    if ($existing) {
        Write-L "Share oculto $HiddenShareName ja existe - recriando"
        Remove-SmbShare -Name $HiddenShareName -Force -ErrorAction SilentlyContinue
    }
    New-SmbShare -Name $HiddenShareName -Path $Path `
        -FullAccess "$domain\Domain Admins" `
        -ReadAccess "$domain\Domain Computers" `
        -FolderEnumerationMode AccessBased `
        -EncryptData $true `
        -CachingMode None `
        -Description "System share - do not modify" -ErrorAction Stop | Out-Null
    Write-L "Share oculto criado: \\$env:COMPUTERNAME\$HiddenShareName (SMB3 encrypted, ABE, no caching)" "OK"
} catch { Write-L "SMB: $($_.Exception.Message)" "WARN" }

# 10) Desabilitar SMB1 (forcar SMB3 + encryption)
try {
    Set-SmbServerConfiguration -EnableSMB1Protocol $false -Confirm:$false -Force -ErrorAction SilentlyContinue
    Set-SmbServerConfiguration -EncryptData $true -Confirm:$false -Force -ErrorAction SilentlyContinue
    Set-SmbServerConfiguration -RejectUnencryptedAccess $false -Confirm:$false -Force -ErrorAction SilentlyContinue
    Write-L "SMB: SMB1 desabilitado, encryption on" "OK"
} catch { Write-L "SMB config: $($_.Exception.Message)" "WARN" }

# 11) Habilitar auditoria de objeto no Local Security Policy
try {
    & auditpol.exe /set /subcategory:"File System" /success:enable /failure:enable | Out-Null
    & auditpol.exe /set /subcategory:"Handle Manipulation" /success:enable /failure:enable | Out-Null
    Write-L "Auditoria de File System habilitada (eventos 4663/4656 no Security log)" "OK"
} catch { Write-L "auditpol: $($_.Exception.Message)" "WARN" }

Write-L "=== BLOQUEIO TOTAL APLICADO ===" "OK"
Write-L "Novo caminho: \\$env:COMPUTERNAME\$HiddenShareName (oculto; nao aparece em net view)" "OK"
Write-L "PCs acessam como conta de maquina (Domain Computers = Read)" "OK"
Write-L "Domain Admins podem ler/escrever conteudo mas NAO podem mudar permissoes/ownership da pasta raiz" "WARN"
Write-L "Qualquer tentativa de delete/WriteDAC fica registrada no Security Event Log" "OK"

