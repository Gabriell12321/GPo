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
