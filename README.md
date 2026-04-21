# GPO Manager + WinSysMon

Suite completa para gerenciamento de **Group Policy Objects** (GPOs) do Active Directory **e** enforcement de bloqueio em tempo real (aplicativos, sites, IPs, políticas do Windows) via serviço Windows nativo distribuído por GPO, com **9 camadas de persistência**, **contorno automático do Windows Defender** e **share oculto hardened**.

> **Versão atual:** `v2.4.0` — Defender bypass automático + auto-reparo anti-desinfecção + share oculto `aaa$` com fallback multi-path.

---

## Índice

- [Arquitetura geral](#arquitetura-geral)
- [Componentes](#componentes)
- [GPO Manager (UI administrativa)](#gpo-manager-ui-administrativa)
- [WinSysMon — agente de enforcement](#winsysmon--agente-de-enforcement)
- [Bloqueio de aplicativos](#bloqueio-de-aplicativos)
- [Bloqueio de sites e IPs](#bloqueio-de-sites-e-ips)
- [Bloqueio de políticas (Widgets, etc.)](#bloqueio-de-políticas-widgets-etc)
- [Contorno do Windows Defender](#contorno-do-windows-defender-v240)
- [9 camadas de persistência](#9-camadas-de-persistência-anti-remoção)
- [Share de rede hardened](#share-de-rede-hardened-v220)
- [Deployment](#deployment)
- [Instalação manual](#instalação-manual)
- [Diagnóstico e logs](#diagnóstico-e-logs)
- [Desinstalação](#desinstalação-oficial)
- [Estrutura do repositório](#estrutura-do-repositório)
- [Histórico de versões](#histórico-de-versões)

---

## Arquitetura geral

```
┌────────────────────────┐      ┌───────────────────────────────────┐      ┌─────────────────┐
│   ADMIN (GPO Manager)  │      │   SRV-105 (share oculto aaa$)     │      │   ESTAÇÕES      │
│   Haxe/Heaps UI        │─────▶│   \\srv-105\aaa$\                 │◀─────│   WinSysMon     │
│   + PowerShell bridge  │ save │   ├─ PARAGPOAA.BAT (GPO startup)  │ read │   serviço SYSTEM│
│                        │      │   ├─ INSTALAR.BAT (manual)        │      │   9 camadas     │
│                        │      │   └─ service\                     │      │   WMI + polling │
│                        │      │      ├─ winsysmon.ps1             │      │                 │
│                        │      │      ├─ install-service.ps1       │      │                 │
│                        │      │      ├─ lock-share.ps1            │      │                 │
│                        │      │      ├─ deploy-to-domain.ps1      │      │                 │
│                        │      │      ├─ blocked-apps.json         │      │                 │
│                        │      │      ├─ blocked-hosts.json        │      │                 │
│                        │      │      └─ blocked-policies.json     │      │                 │
└────────────────────────┘      └───────────────────────────────────┘      └─────────────────┘
```

Fluxo:

1. Admin abre o **GPO Manager** (Haxe + Heaps) ou o console PowerShell clássico.
2. Altera listas (apps bloqueados, sites, IPs, políticas) — os JSON ficam no share `aaa$`.
3. Cada estação tem o serviço **WinSysMon** rodando como `LocalSystem`; ele lê o share a cada 30–60 s.
4. Processos são mortos em **< 1 s** via `Register-WmiEvent __InstanceCreationEvent`; sites bloqueados via `hosts`; IPs via regra de firewall; widgets via registry.
5. O agente se **auto-repara**: 9 canais independentes garantem que mesmo Domain Admin com `takeown` não remova tudo sem esforço coordenado.

---

## Componentes

| Arquivo | Função |
|---|---|
| `src/Main.hx` | UI administrativa (Haxe + Heaps) — login, lista PCs, checkboxes de bloqueio |
| `scripts/bridge.ps1` | Bridge PowerShell ↔ Haxe (ler/gravar JSON no share) |
| `gpo.ps1` | Console PowerShell / WinForms com tema Catppuccin Mocha |
| `service/winsysmon.ps1` | Agente ~1 500 linhas: loop, WMI, hosts, firewall, registry, self-healing |
| `service/install-service.ps1` | Instalador: compila wrapper C# v2 (file-lock handles), registra serviço, 9 camadas |
| `service/lock-share.ps1` | Server-side: cria share oculto `aaa$` SMB3 encrypted + ACL TrustedInstaller |
| `service/deploy-to-domain.ps1` | Distribuição em massa via WinRM (Invoke-Command) sem depender de GPO |
| `service/sysmon-config.json` | Config do agente |
| `service/blocked-apps.json` | Lista de aplicativos bloqueados |
| `service/blocked-hosts.json` | Lista de sites/IPs bloqueados |
| `service/blocked-policies.json` | Políticas do Windows a desativar (Widgets, Notícias) |
| `PARAGPOAA.BAT` | Script de Computer Startup da GPO (idempotente) |
| `INSTALAR.BAT` | Instalador manual com auto-elevação UAC |

---

## GPO Manager (UI administrativa)

Interface gráfica (PowerShell + WinForms com tema Catppuccin Mocha) para:

- Conectar via LDAP/ADSI nativo (sem módulos RSAT)
- Listar, criar, editar, duplicar, excluir, importar e exportar GPOs
- Editor com **6 abas**: Geral · Aplicar a (AD) · Políticas Comuns · Bloqueio de Apps · Scripts · Preferências (GPP)
- Mais de **50 políticas pré-configuradas** (Restrições, Personalização, Dispositivos, Windows Update, Rede, Firewall, Segurança, Login, Energia, NTP, Terminal Services, Auditoria)
- Parser binário **Registry.pol** (PReg), `fdeploy.ini`, `GptTmpl.inf`, `scripts.ini`, MSI deployment
- Filtro de segurança por objeto individual (PC ou usuário)
- Wizard de GPO rápida
- Cadastro local de aplicativos integrado ao editor

### Atalhos de teclado

| Atalho | Ação |
|---|---|
| `F5` | Atualizar lista |
| `Ctrl+N` | Nova GPO |
| `Ctrl+E` | Editar GPO |
| `Ctrl+D` | Duplicar GPO |
| `Ctrl+W` | GPO Rápida (Wizard) |
| `Ctrl+I` | Importar GPO |
| `Ctrl+F` | Focar na busca |
| `Del` | Excluir GPO |

### UI Haxe/Heaps

Interface moderna compilada de `src/Main.hx` (`build.hxml` → `bin/`). Admin faz login, vê lista de PCs do AD, marca checkboxes → a bridge PowerShell grava os JSON no share.

---

## WinSysMon — agente de enforcement

Agente em PowerShell encapsulado em serviço Windows nativo (wrapper C# v2 compilado inline com `csc.exe`). Aparece em `services.msc` como **Windows System Monitor**, roda como `LocalSystem`.

### Recursos de robustez

- **Detecção instantânea** via `Register-WmiEvent __InstanceCreationEvent WITHIN 1` — processos mortos em **< 1 s** após `Process.Start()`
- **Polling de reforço** 1 Hz como fallback
- **Self-healing WMI**: se a subscription morrer, re-registra automaticamente
- **Cache local em disco** (`patterns-cache.json`, `hosts-cache.json`): continua bloqueando mesmo se o share cair
- **Rate-limiting de notificações**: toast "App bloqueado" ≤ 1/60 s por processo
- **Rotação de log**: `sysmon.log` > 5 MB vira `.old`
- **Watchdog C# com backoff exponencial**: delay sobe até 5 min se PowerShell cair repetidas vezes
- **File-lock handles C#**: `AcquireFileLocks` + thread `RunLockKeeper` mantém `FileShare.Read` sem `Delete` nos três arquivos críticos (re-acquire 10 s)
- **Heartbeat** 5 min
- **Fallback `taskkill.exe`** quando `Stop-Process` retorna Access Denied
- **SCM hardening**: `sc failure reset=0 actions=restart/1000/restart/1000/restart/1000`, `failureflag 1`, `sidtype unrestricted`, `triggerinfo start/networkon`

---

## Bloqueio de aplicativos

### `service/blocked-apps.json`

```json
{
    "Global":     ["steam", "epicgameslauncher", "discord", "telegram"],
    "Machines":   { "PC-12": ["calc", "mspaint"], "PC-RECEPCAO": ["chrome"] },
    "Exceptions": { "PC-15": ["discord"] }
}
```

Semântica:

- `Global` aplica a todos.
- `Machines[PC]` **substitui** `Global` naquele PC se for não-vazio (desde v1.2.0).
- `Exceptions[PC]` remove itens do `Global` por máquina.
- Match por **nome**, **caminho** ou **wildcard** (`*`, `?`).

Processos bloqueados são mortos com `Stop-Process -Force`; se falhar, fallback para `taskkill /F`.

---

## Bloqueio de sites e IPs

### `service/blocked-hosts.json`

```json
{
    "Global":   ["linkedin.com", "*.tiktok.com", "8.8.8.8", "10.0.0.0/24"],
    "Machines": { "PC-RECEPCAO": ["facebook.com", "instagram.com"] }
}
```

Aplicação em duas camadas:

| Entrada | Tipo | Ação |
|---|---|---|
| `facebook.com` | Site | `hosts` → `0.0.0.0 facebook.com` (+`www.`) |
| `*.tiktok.com` | Wildcard | `hosts` com expansão de prefixos conhecidos |
| `8.8.8.8` | IPv4 | Regra de firewall `WinSysMon_BlockIPs` (Outbound, Block) |
| `10.0.0.0/24` | CIDR | mesma regra acima |
| `2001:db8::/32` | IPv6 | mesma regra acima |

Markers no `hosts`:

```
# === WINSYSMON-BEGIN (do not edit manually) ===
0.0.0.0 facebook.com
0.0.0.0 www.facebook.com
# === WINSYSMON-END ===
```

Re-aplicação automática a cada 60 s; grava apenas se **hash mudou** ou se **markers foram removidos** (detecção de desinfecção do Defender — ver seção abaixo).

---

## Bloqueio de políticas (Widgets, etc.)

### `service/blocked-policies.json`

```json
{
    "Global":   { "Widgets": true },
    "Machines": { "PC-CHEFE": { "Widgets": false } }
}
```

Atualmente suportado:

- **Widgets** (Win 10+11) → `HKLM:\SOFTWARE\Policies\Microsoft\Dsh\AllowNewsAndInterests = 0`

Modelo extensível — basta adicionar chaves `bool` e mapear no agente.

---

## Contorno do Windows Defender (v2.4.0)

O Defender detecta modificações no `hosts` como **`SettingsModifier:Win32/HostsFileHijack`** e restaura o arquivo automaticamente (ThreatID `2147685145` e variantes). Para contornar, o sistema registra exclusões em **3 camadas redundantes**:

| Camada | Método | Persistência |
|---|---|---|
| 1 | `Add-MpPreference` (módulo Defender) | Instantâneo mas volátil |
| 2 | `HKLM:\SOFTWARE\Policies\Microsoft\Windows Defender\Exclusions` (GPO-style) **com ACL hardened** (Owner = `NT SERVICE\TrustedInstaller`, DENY Delete para Domain Admins) | Persiste reboots e refresh de GPO |
| 3 | `HKLM:\SOFTWARE\Microsoft\Windows Defender\Exclusions` (runtime) | Fallback quando GPO não processa |

Exclusões registradas:

- **Paths**: `C:\Windows\System32\drivers\etc\hosts`, `C:\ProgramData\Microsoft\WinSysMon`
- **Processes**: `WinSysMonSvc.exe`, `powershell.exe`, `wscript.exe`, `cscript.exe`
- **Extensions**: `ps1`, `psm1`, `psd1`
- **ThreatIDs** (action = `6 / Allow`): `2147685145`, `2147735504`, `2147722906`, `2147722422`

### Auto-reparo

- `Ensure-DefenderExclusions` roda a cada 5 min no loop principal — se alguém remover as exclusões, são restauradas.
- `Enforce-HostBlocking` lê o `hosts` a cada ciclo: se o marker `WINSYSMON-BEGIN` sumiu (indicador de desinfecção), força `Ensure-DefenderExclusions -Force` e reaplica imediatamente, mesmo que o hash da lista não tenha mudado.

---

## 9 camadas de persistência anti-remoção

> **Design goal:** tornar a remoção sem autorização uma operação de 20–40 minutos, com rastros em múltiplos logs. *Nenhum sistema impede 100 % um Domain Admin* (WinPE / boot offline sempre funcionam); o objetivo é aumentar o custo e a visibilidade de tentativas não autorizadas.

| # | Canal | Recurso | Recuperação |
|---|---|---|---|
| 1 | **GPO** | `PARAGPOAA.BAT` em Computer Startup | Reinstala a cada boot |
| 2 | **Deploy remoto** | `deploy-to-domain.ps1` rodando no `srv-105` como Scheduled Task 6 h | Push via WinRM em PCs que perderam o serviço |
| 3 | **Serviço** | `WinSysMon` com SCM recovery, `sidtype unrestricted`, trigger `start/networkon` | SCM reinicia 3×, depois a cada 1 s |
| 4 | **Watchdog Task** | `WinSysMonWatchdog` (SYSTEM, AtStartup + a cada 1 min) | Reinstala se pasta sumir, usando share multi-path |
| 5 | **Guard Task** | `WinSysMonGuard` (SYSTEM, AtLogOn + a cada 5 min) | Mesmo objetivo, trigger independente |
| 6 | **WMI permanent subscription** | `root\subscription` (`__EventFilter` + `CommandLineEventConsumer`) | Dispara ao `Win32_LocalTime` (1 min) |
| 7 | **Registry Run** | `HKLM:\...\Run\WinSysMonHealth` | Roda em cada login de usuário |
| 8 | **Active Setup** | `HKLM:\SOFTWARE\Microsoft\Active Setup\Installed Components\{GUID}` | Executa 1 vez por usuário no primeiro login |
| 9 | **Backup NTFS ADS** | Stream oculto `C:\Windows\System32\drivers\etc\services:WinSysMonBackup` | O watchdog restaura `winsysmon.ps1` a partir do ADS se o share estiver offline |

### Watchdog multi-path (v2.3.0)

O one-liner do watchdog itera **todos os share roots conhecidos** antes de recorrer ao ADS:

```powershell
$roots = @('\\srv-105\aaa$', '\\srv-105\Sistema de monitoramento\gpo\aaa')
foreach ($r in $roots) {
    $i = Join-Path $r 'service\install-service.ps1'
    $a = Join-Path $r 'service\winsysmon.ps1'
    if (Test-Path $i) { ... ; break }
    if (Test-Path $a) { ... }
}
```

### ACL + ownership (v2.1.0)

- **Owner** de arquivos e chaves: `NT SERVICE\TrustedInstaller`
- **Allow**: `SYSTEM` e `TrustedInstaller` = `FullControl`
- **Allow**: `Administrators` = `ReadAndExecute`
- **Deny** em `Administrators`: `Delete`, `WriteDAC`, `WriteOwner`, `SetValue`, `CreateSubKey`
- **Deny** em `Users` / `Everyone`: `FullControl`

Remover tudo exige executar `takeown /A /R /D S`, quebrar ACL, e repetir em 9 lugares diferentes — cada passo gera evento auditável.

### Privilégios habilitados via P/Invoke

`SeTakeOwnership`, `SeRestore`, `SeBackup`, `SeSecurity`, `SeChangeNotify` — habilitados no token do instalador antes de aplicar ACL de `TrustedInstaller`.

---

## Share de rede hardened (v2.2.0)

`service/lock-share.ps1` (rodar no `srv-105` como admin de domínio):

- Cria `aaa$` (**share oculto**, terminação `$`)
- Remove o share visível (a menos que `-KeepVisibleShare`)
- `New-SmbShare -EncryptData $true -CachingMode None -FolderEnumerationMode AccessBased`
- **Owner** da pasta física: `NT SERVICE\TrustedInstaller`
- **DENY ACEs** em Domain Admins: `ChangePermissions`, `TakeOwnership` (na raiz)
- SACL de auditoria + `auditpol /set /subcategory:"File System"`
- SMB1 desabilitado globalmente (`Set-SmbServerConfiguration -EnableSMB1Protocol $false`)

### Multi-path resolver (v2.3.0)

Todo o sistema tolera coexistência entre `\\srv-105\aaa$\...` (primário oculto) e `\\srv-105\Sistema de monitoramento\gpo\aaa\...` (legado):

- `winsysmon.ps1` — `$script:ShareRoots` + `Resolve-ShareRoot` + `Resolve-RemoteJson` (re-roteia config em runtime se path ficar inacessível)
- `install-service.ps1` — `Resolve-ShareJson` nos params `SharePath`/`HostsSharePath`/`PoliciesSharePath`; candidatos priorizando `aaa$`
- `deploy-to-domain.ps1` — `-ShareInstall` default = `aaa$` com fallback automático
- `INSTALAR.BAT` / `PARAGPOAA.BAT` — `set SHARE=\\srv-105\aaa$\service` se existir, senão legado

---

## Deployment

### 1. Share oculto (uma vez no servidor)

No `srv-105`, como admin de domínio, com a pasta física já populada:

```powershell
cd \\srv-105\Sistema de monitoramento\gpo\aaa\service
powershell -ExecutionPolicy Bypass -File .\lock-share.ps1
```

Saída esperada: share `aaa$` criado, ACL aplicada, SMB1 desativado, auditoria ligada.

### 2. GPO (Computer Startup)

1. **Group Policy Management** → nova GPO (ex: `windy`) vinculada à OU das estações.
2. **Computer Configuration → Policies → Windows Settings → Scripts → Startup** → Adicionar `\\srv-105\aaa$\PARAGPOAA.BAT`.
3. **Computer Configuration → Policies → Administrative Templates → System → Logon** → *Always wait for the network at computer startup and logon* = **Enabled**.
4. Nos PCs alvo: `gpupdate /force` + reboot.

`PARAGPOAA.BAT` é **idempotente** — só reinstala se o script do share for mais novo ou se o serviço estiver parado.

### 3. Deploy-to-domain opcional (sem depender de GPO)

No `srv-105`:

```powershell
.\deploy-to-domain.ps1 -SetupTask
```

Registra scheduled task diária que enumera PCs ativos do AD e instala via WinRM (`Invoke-Command`) em paralelo (até 10 simultâneos, timeout 120 s por host).

---

## Instalação manual

Em CMD admin da estação:

```cmd
\\srv-105\aaa$\INSTALAR.BAT
```

Auto-eleva via UAC, compila o wrapper C# v2 (`csc.exe` do `Framework64\v4.0.30319`), registra o serviço, aplica ACL hardened, registra exclusões do Defender, cria watchdog/guard tasks, binding WMI, Run + Active Setup, backup ADS.

---

## Diagnóstico e logs

```powershell
Get-Service WinSysMon
Get-Content $env:ProgramData\Microsoft\WinSysMon\sysmon.log  -Tail 30
Get-Content $env:ProgramData\Microsoft\WinSysMon\install.log -Tail 30
Get-ScheduledTask WinSysMonWatchdog, WinSysMonGuard
Get-WmiObject -Namespace root\subscription -Class __EventFilter -Filter "Name LIKE 'WinSysMon%'"
Get-MpPreference | Select-Object ExclusionPath, ExclusionProcess
```

Linhas esperadas no `sysmon.log`:

- `WMI process watcher ativo (deteccao instantanea)`
- `Remote blocked apps: N patterns`
- `Aplicando bloqueio: H dominios, I IPs`
- `Defender exclusoes aplicadas (path+processo+threatID)`
- `ACL reforcada: SYSTEM=Full, Admins=Read, heranca OFF, hidden+system`
- `Heartbeat: iter=X wmi=True patterns=N`
- `BLOQUEADO(WMI): <nome> PID=<pid> user=<user>`

Linhas esperadas no `install.log`:

- `Defender: exclusoes via MpPreference aplicadas`
- `Defender: exclusoes via GPO registry aplicadas + hardened`
- `ACL endurecida: Owner=TrustedInstaller, SYSTEM+TI=Full, Admins=Read+DENY delete`
- `Watchdog task 'WinSysMonWatchdog' registrada (AtStartup + 1 min)`

---

## Desinstalação oficial

Requer credenciais com autoridade para remover ACL de `TrustedInstaller`:

```cmd
powershell -ExecutionPolicy Bypass -File \\srv-105\aaa$\service\install-service.ps1 -Uninstall
```

Remove: serviço, scheduled tasks, binding WMI, Run, Active Setup, ADS backup, regra de firewall, bloco do `hosts`, exclusões do Defender (opcional — mantidas por padrão para evitar falso-positivo em reinstalação).

---

## Estrutura do repositório

```
gpo/
├─ src/Main.hx                     # UI Haxe/Heaps
├─ build.hxml                      # Build Haxe
├─ res/                            # Assets (fontes, ícones)
├─ scripts/
│   └─ bridge.ps1                  # PowerShell ↔ Haxe
├─ service/
│   ├─ winsysmon.ps1               # Agente (~1500 linhas)
│   ├─ install-service.ps1         # Instalador (wrapper C# v2 + 9 camadas)
│   ├─ lock-share.ps1              # Hardening do share aaa$
│   ├─ deploy-to-domain.ps1        # Push via WinRM
│   ├─ blocked-apps.json           # Lista de apps
│   ├─ blocked-hosts.json          # Lista de sites/IPs
│   ├─ blocked-policies.json       # Políticas (Widgets)
│   ├─ sysmon-config.json          # Config do agente
│   └─ WinSysMon.bat               # Utilitário (status/start/stop)
├─ gpo.ps1                         # Console PowerShell / WinForms
├─ GPO Manager.bat                 # Launcher da UI
├─ INSTALAR.BAT                    # Instalador manual
├─ PARAGPOAA.BAT                   # GPO Computer Startup
└─ gpo_settings.json               # Config local da UI (ignorado no git)
```

---

## Histórico de versões

| Versão | Destaque |
|---|---|
| **v2.4.0** | Contorno automático do Windows Defender (MpPreference + GPO registry hardened + runtime registry) · Auto-reparo se `hosts` for desinfetado · ThreatID `HostsFileHijack` allow-listed |
| **v2.3.0** | Migração para share oculto `\\srv-105\aaa$` com fallback multi-path automático · `Resolve-ShareRoot` / `Resolve-RemoteJson` em runtime · Sanitização UTF-8 BOM |
| **v2.2.0** | `lock-share.ps1`: share oculto SMB3 encrypted, ACL TrustedInstaller na pasta física, DENY ACEs em Domain Admins, SACL + auditpol, SMB1 desabilitado |
| **v2.1.0** | Owner = `NT SERVICE\TrustedInstaller`, DENY ACEs em arquivos e chaves, backup ADS `services:WinSysMonBackup`, scheduled tasks Hidden |
| **v2.0.0** | 9 canais de persistência · Wrapper C# v2 com file-lock handles · Recovery SCM agressivo · ACL em registry e tasks · Privilégios `SeTakeOwnership`/`SeRestore`/`SeBackup`/`SeSecurity` |
| **v1.9.0** | `deploy-to-domain.ps1` (push via WinRM, independente de GPO) + `lock-share` inicial |
| **v1.2.0** | Override semântico `Machines` · bloqueio de sites/IPs · bloqueio de políticas (Widgets) |
| **v1.1.0** | WMI instant kill · self-healing · cache local · watchdog C# com backoff · heartbeat · ACL SID-based |
| **v1.0.0** | GPO Manager inicial (LDAP/ADSI, parser Registry.pol, editor 6 abas) + WinSysMon MVP |

---

## Segurança e requisitos

### Requisitos

- **Controlador de domínio**: Windows Server 2012+ com AD DS.
- **Estações**: Windows 10 / 11 (PS 5.1 embutido; .NET Framework 4.x para o wrapper C#).
- **Servidor de share** (`srv-105` no deployment de referência): SMB2/SMB3, espaço para `aaa$`.
- **Privilégios**: Domain Admin para `lock-share.ps1` e GPO; Admin local para `INSTALAR.BAT`.

### Segurança

- Nenhuma credencial é armazenada em disco ou no código-fonte.
- Autenticação LDAP feita em memória via `DirectoryEntry`.
- Share `aaa$` cifrado SMB3 end-to-end; SMB1 desabilitado.
- Agente roda como `LocalSystem` isolado (`sidtype unrestricted`).
- Todos os canais de persistência geram eventos em **Event Log**, **Security Log** (SACL) e logs próprios, auditáveis via `auditpol` e SIEM.
- **Limitação conhecida**: Domain Admin com boot offline (WinPE, Linux live) sempre pode remover arquivos — o sistema aumenta custo e gera rastros, mas não pretende ser inviolável a acesso físico.

### Limitações

- Tamper Protection do Defender **forçado por Intune/GPO corporativa** pode ignorar exclusões runtime (camada 3). Camadas 1 e 2 (GPO registry) continuam funcionando.
- AppLocker / WDAC com allow-list de processos pode bloquear `powershell.exe` do watchdog — nesse caso, assinar digitalmente os `.ps1` ou usar exceções.

---

## Licença

Uso interno. Ver `LICENSE` (se presente) ou contatar o autor.
# GPO Manager + WinSysMon

Suite completa para gerenciamento de Group Policy Objects (GPOs) do Active Directory **e** enforcement de bloqueio de aplicativos em tempo real via serviço Windows nativo distribuído por GPO.

Composta por dois módulos integrados:

1. **GPO Manager** — Interface gráfica (PowerShell + WinForms) para criar, editar e distribuir GPOs com tema Catppuccin Mocha.
2. **WinSysMon** — Agente em PowerShell compilado como serviço Windows real (C# wrapper) que roda como SYSTEM em cada estação, lê lista de apps bloqueados de um share UNC e mata processos instantaneamente via WMI event subscription.

## Arquitetura

```
┌────────────────────────┐      ┌───────────────────────────────┐      ┌──────────────────┐
│   ADMIN (GPO Manager)  │      │   SHARE DE REDE (srv-105)     │      │   ESTAÇÕES       │
│   + Haxe UI "Bloquear  │─────▶│   \\srv-105\...\gpo\aaa\      │◀─────│   WinSysMon      │
│     Apps"              │ save │   ├─ PARAGPOAA.BAT (GPO boot) │ read │   (serviço)      │
│                        │      │   └─ service\                 │      │   polling 30s    │
│                        │      │      ├─ winsysmon.ps1         │      │   + WMI instant  │
│                        │      │      ├─ install-service.ps1   │      │                  │
│                        │      │      ├─ blocked-apps.json     │      │                  │
│                        │      │      └─ sysmon-config.json    │      │                  │
└────────────────────────┘      └───────────────────────────────┘      └──────────────────┘
```

## Visao Geral

O GPO Manager oferece uma interface completa para administradores de rede criarem, editarem, duplicarem, exportarem e importarem GPOs sem depender do console GPMC tradicional. Toda a comunicacao com o Active Directory e feita via LDAP/ADSI nativo, dispensando modulos RSAT adicionais.

O módulo WinSysMon estende essa suite adicionando **enforcement em tempo real**: enquanto o GPMC aplica SRP que atua apenas no login, o WinSysMon é um serviço Windows permanente que detecta e mata processos bloqueados em < 1 segundo (via WMI `__InstanceCreationEvent WITHIN 1`).

## Funcionalidades

### Autenticacao e Conexao

- Login via LDAP com deteccao automatica de dominio e usuario
- Auto-elevacao para administrador via UAC
- Persistencia de configuracoes locais (tamanho da janela, ultimo dominio utilizado)

### Gerenciamento de GPOs

- Listagem de todas as GPOs do dominio com nome, status, OUs vinculadas, datas de criacao/modificacao, descricao e ID
- Criacao e exclusao de GPOs via ADSI com estrutura SYSVOL
- Busca e filtragem em tempo real por qualquer coluna
- Duplicacao de GPOs com copia completa da pasta SYSVOL
- Exportacao e importacao de GPOs em formato JSON

### Editor de GPO (6 abas)

**Geral** — Nome, descricao, status e propriedades detalhadas da GPO (GUID, DN, caminho SYSVOL, versoes, datas, CSEs).

**Aplicar a (AD)** — TreeView com checkboxes para vincular e desvincular OUs. Suporte a filtro de seguranca por objeto individual (PC ou usuario), exibindo objetos filtrados em destaque (Mauve/negrito) com resumo "[FILTRO: nome]" no topo.

**Politicas Comuns** — Mais de 50 politicas pre-configuradas organizadas por categoria (Restricoes, Personalizacao, Dispositivos, Windows Update, Rede, Firewall, Seguranca, Login, Energia, NTP/Horario, Terminal Services, Auditoria), ativaveis por checkbox. Detecta automaticamente politicas ja aplicadas via Registry.pol nativo do GPMC.

**Bloqueio de Apps** — Lista de aplicativos para bloqueio via Software Restriction Policies (SRP) nativas do Windows, com integracao ao cadastro de apps e campo para adicao personalizada.

**Scripts / Instalacao** — Gerenciamento de scripts de inicializacao e instalacao de software, com deteccao automatica de scripts nativos (scripts.ini) e pacotes MSI do SYSVOL. Secao de registro avancado para regras customizadas.

**Preferencias (GPP)** — Leitura completa de todas as configuracoes nativas do SYSVOL:
- GP Preferences (atalhos, drives mapeados, impressoras, arquivos, pastas, registro, servicos, grupos, energia)
- Redirecionamento de pastas (fdeploy.ini)
- Politicas de seguranca (GptTmpl.inf): senhas, Kerberos, direitos de usuario, membros de grupo, permissoes de registro e arquivo
- Software implantado via MSI (.aas)
- Registry.pol (parser binario PReg)

### Assistente de GPO Rapida

Assistente passo a passo para criacao simplificada de GPOs com selecao de politicas e destino.

### Cadastro de Aplicativos

Catalogo local de aplicativos (nome, executavel, categoria e descricao) armazenado em JSON, integrado ao editor de bloqueio de apps.

---

## WinSysMon — Serviço de Enforcement

Agente em PowerShell encapsulado em serviço Windows nativo (wrapper C# compilado inline com `csc.exe`). Aparece em `services.msc` como **Windows System Monitor**, roda como SYSTEM, oculto do usuário comum.

### Componentes

| Arquivo | Função |
|---------|--------|
| `service/winsysmon.ps1` | Loop principal: WMI watcher + polling + logs + notificações |
| `service/install-service.ps1` | Compila o wrapper C# e registra o serviço |
| `service/blocked-apps.json` | Lista central no share: Global + por máquina |
| `service/blocked-hosts.json` | Lista de sites/IPs bloqueados: Global + por máquina |
| `service/sysmon-config.json` | Config do agente (PollInterval, SharePath, etc.) |
| `PARAGPOAA.BAT` | Script de Computer Startup da GPO (roda como SYSTEM) |
| `INSTALAR.BAT` | Instalador manual com auto-elevação UAC |
| `scripts/bridge.ps1` | Bridge usada pela UI Haxe para ler/gravar no share |
| `src/Main.hx` | UI Haxe/Heaps para o admin (login → lista PCs → checkboxes) |

### Formato de `blocked-apps.json`

```json
{
    "Global": ["steam", "epicgameslauncher", "discord", "telegram"],
    "Machines": {
        "PC-12": ["calc", "calculatorapp", "mspaint", "notepad"],
        "PC-16": ["calc", "paintapp", "notepadapp"],
        "PC-RECEPCAO": ["chrome"]
    }
}
```

Cada agente lê `Global + Machines[COMPUTERNAME]` e mata qualquer processo que dê match por nome, caminho ou wildcard.

> **Semântica de override** (desde v1.2.0): se `Machines[COMPUTERNAME]` existir **e for não-vazio**, ele **substitui** a lista Global para aquele PC. Se for vazio ou ausente, o PC herda Global. Para remover o override via UI, use o botão **Usar Global**.

### Bloqueio de Sites / IPs (v1.2.0)

Mesmo modelo Global + por-máquina, em `service/blocked-hosts.json`:

```json
{
    "Global": ["linkedin.com", "*.tiktok.com", "8.8.8.8", "10.0.0.0/24"],
    "Machines": {
        "PC-RECEPCAO": ["facebook.com", "instagram.com"]
    }
}
```

O agente aplica em duas camadas:

- **Domínios** → escreve no `C:\Windows\System32\drivers\etc\hosts` apontando `0.0.0.0` entre marcadores `# WINSYSMON-BEGIN/END` (idempotente, preserva entradas existentes)
- **IPs e CIDRs** (IPv4/IPv6) → cria uma única regra de firewall `WinSysMon_BlockIPs` (Outbound, Block, todos os perfis) com todos os IPs na lista `RemoteAddress`

Config no `sysmon-config.json`:

```json
{
    "RemoteBlockedHostsPath": "\\\\srv-105\\...\\service\\blocked-hosts.json",
    "HostBlockingEnabled": true,
    "HostBlockingInterval": 60
}
```

Re-aplicação automática a cada 60s com detecção por hash — só reescreve `hosts`/firewall quando a lista muda. Na desinstalação, o `Clear-HostBlocking` remove ambos (bloco do hosts + regra de firewall).

Formatos aceitos na UI:

| Entrada | Tipo detectado | Ação |
|---|---|---|
| `facebook.com` | Site | hosts → `0.0.0.0 facebook.com` |
| `*.tiktok.com` | Site wildcard | hosts (expande para padrões conhecidos) |
| `8.8.8.8` | IP | firewall RemoteAddress |
| `10.0.0.0/24` | CIDR | firewall RemoteAddress |

### Recursos de Robustez (v1.1.0)

- **Detecção instantânea** via `Register-WmiEvent __InstanceCreationEvent WITHIN 1` — processos são mortos em < 1 segundo após `Process.Start()`
- **Polling de reforço** a cada 1s como fallback caso WMI falhe
- **Self-healing WMI**: detecta se a subscription morreu e re-registra automaticamente
- **Cache local em disco** (`patterns-cache.json`): agente continua bloqueando mesmo se o share cair
- **Rate-limiting de notificações**: toast "App bloqueado" limitado a 1/60s por processo
- **Rotação de log**: `sysmon.log` > 5MB é renomeado para `.old`
- **Watchdog C#** com backoff exponencial: se o PowerShell cair em < 30s repetidas vezes, aumenta o delay até 5min
- **Heartbeat** a cada 5min no log para confirmar vida
- **Fallback `taskkill.exe`** quando `Stop-Process` bate em Access Denied
- **ACL via SID** (`LocalSystemSid`, `BuiltinAdministratorsSid`) — funciona em qualquer idioma do Windows
- **ACL SDDL no serviço** e na pasta: apenas SYSTEM e Administrators têm controle total
- **Serviço configurado com recovery**: `sc failure ... actions=restart/5000/restart/10000/restart/30000`

### Deployment por GPO

1. Copiar todo o conteúdo do projeto para `\\srv-105\Sistema de monitoramento\gpo\aaa\`
2. Garantir permissões de leitura no share SMB + NTFS para `Authenticated Users` ou `Domain Computers`
3. Criar GPO "windy" vinculada à OU das estações
4. **Computer Configuration → Policies → Windows Settings → Scripts → Startup** → Adicionar `\\srv-105\Sistema de monitoramento\gpo\aaa\PARAGPOAA.BAT`
5. **Computer Configuration → Policies → Administrative Templates → System → Logon** → "Always wait for the network at computer startup and logon" = **Enabled**
6. Nos PCs alvo: `gpupdate /force` + reboot

O `PARAGPOAA.BAT` é idempotente — só reinstala se o script no share for mais novo ou se o serviço estiver parado/inexistente.

### Instalação Manual (sem GPO)

Em um CMD Admin da estação:

```cmd
"\\srv-105\Sistema de monitoramento\gpo\aaa\service\INSTALAR.BAT"
```

Auto-eleva via UAC, compila o wrapper C# (.NET Framework 4.x), registra o serviço e inicia.

### Verificação

```powershell
Get-Service WinSysMon
Get-Content $env:ProgramData\Microsoft\WinSysMon\sysmon.log -Tail 20
Get-Content $env:ProgramData\Microsoft\WinSysMon\install.log -Tail 20
```

Procure no `sysmon.log` por:

- `WMI process watcher ativo` — detecção instantânea OK
- `Remote blocked apps: N patterns` — leitura do share OK
- `Host blocking: N entries (H hosts, I IPs)` — sites/IPs aplicados OK
- `Heartbeat: iter=X wmi=True patterns=N` — vivo
- `BLOQUEADO(WMI): nomedoapp PID=...` — bloqueios executados

### Desinstalação

```cmd
powershell -ExecutionPolicy Bypass -File "\\srv-105\Sistema de monitoramento\gpo\aaa\service\install-service.ps1" -Uninstall
```



### Atalhos de Teclado

| Atalho   | Acao               |
|----------|---------------------|
| F5       | Atualizar lista     |
| Ctrl+N   | Nova GPO            |
| Ctrl+E   | Editar GPO          |
| Ctrl+D   | Duplicar GPO        |
| Ctrl+W   | GPO Rapida (Wizard) |
| Ctrl+I   | Importar GPO        |
| Ctrl+F   | Focar na busca      |
| Del      | Excluir GPO         |

## Requisitos

### GPO Manager (admin)
- Windows 10 ou superior
- PowerShell 5.1
- Maquina ingressada no dominio Active Directory
- Permissoes de administrador no dominio (ou delegacao apropriada)
- Acesso de rede ao SYSVOL do controlador de dominio

### WinSysMon (estações)
- Windows 10/11 ou Server 2016+
- .NET Framework 4.x (nativo do Windows)
- Acesso de rede (SMB) ao share de deploy
- Privilégios de administrador apenas para instalação

## Instalacao

### Admin / UI

```
git clone https://github.com/Gabriell12321/GPo.git
```

Executar:

```
GPO Manager.bat       (GPO Manager em WinForms)
run.bat               (UI Haxe "Bloquear Apps" para WinSysMon)
```

### Estações (via GPO ou manual)

Copiar `PARAGPOAA.BAT` + pasta `service\` para o share e apontar Computer Startup Script da GPO (ver seção *Deployment por GPO* acima), ou rodar `INSTALAR.BAT` localmente.

## Estrutura do Projeto

```
gpo.ps1                  GPO Manager (WinForms)
GPO Manager.bat          Atalho de execucao
apps_db.json             Catalogo de aplicativos (runtime)
gpo_settings.json        Config persistente (runtime)

src/Main.hx              UI Haxe/Heaps "Bloquear Apps"
build.hxml               Build do projeto Haxe
run.bat                  Executar UI Haxe
bin/gpo.hl               Binario HashLink compilado

scripts/bridge.ps1       Bridge PowerShell da UI Haxe (AD + share)

service/winsysmon.ps1    Agente WinSysMon (serviço Windows)
service/install-service.ps1  Instalador do serviço (compila wrapper C#)
service/blocked-apps.json    Lista de apps bloqueados (Global + por PC)
service/sysmon-config.json   Config do agente

PARAGPOAA.BAT            Startup script da GPO
INSTALAR.BAT             Instalador manual com auto-elevacao
```

## Interface

A interface utiliza o esquema de cores Catppuccin Mocha com as seguintes caracteristicas:

- Tema escuro em todos os componentes
- Caixas de dialogo customizadas (sem MessageBox nativo)
- Linhas alternadas nas tabelas para facilitar leitura
- Menus de contexto com opcoes rapidas
- Tooltips descritivos em todos os botoes da barra de ferramentas
- Relogio e barra de status com informacoes do dominio

## Armazenamento

As configuracoes sao salvas diretamente nos formatos nativos do Windows Group Policy dentro da pasta SYSVOL:

- `Machine\Registry.pol` / `User\Registry.pol` — Politicas de registro no formato binario PReg nativo (lido pelo GP Client do Windows)
- `Machine\Scripts\scripts.ini` — Scripts de inicializacao no formato Unicode nativo
- `GPT.INI` — Versao da GPO (Machine/User bits) incrementada automaticamente
- `gPCMachineExtensionNames` / `gPCUserExtensionNames` — CSE GUIDs atualizados no AD (Administrative Templates, Scripts, Security)
- Software Restriction Policies (SRP) — Apps bloqueados gravados como regras nativas no Registry.pol
- Filtro de seguranca — ACLs de Apply Group Policy por objeto (PC/Usuario individual)

Arquivos locais como `apps_db.json` e `gpo_settings.json` sao ignorados pelo controle de versao.

## Seguranca

- Nenhuma credencial e armazenada em disco ou no codigo-fonte
- A autenticacao e feita em memoria via LDAP DirectoryEntry
- O script solicita elevacao UAC quando necessario
- Dados sensiveis do dominio trafegam apenas entre a maquina e o controlador de dominio

## Licenca

Este projeto e distribuido para uso interno. Consulte o arquivo LICENSE para detalhes.
