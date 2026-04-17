# GPO Manager

Ferramenta grafica para gerenciamento de Group Policy Objects (GPOs) do Active Directory, desenvolvida em PowerShell com interface WinForms e tema escuro Catppuccin.

## Visao Geral

O GPO Manager oferece uma interface completa para administradores de rede criarem, editarem, duplicarem, exportarem e importarem GPOs sem depender do console GPMC tradicional. Toda a comunicacao com o Active Directory e feita via LDAP/ADSI nativo, dispensando modulos RSAT adicionais.

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

- Windows 10 ou superior
- PowerShell 5.1
- Maquina ingressada no dominio Active Directory
- Permissoes de administrador no dominio (ou delegacao apropriada)
- Acesso de rede ao SYSVOL do controlador de dominio

## Instalacao

1. Clone o repositorio:

```
git clone https://github.com/Gabriell12321/GPo.git
```

2. Execute o arquivo `GPO Manager.bat` ou rode diretamente pelo PowerShell:

```powershell
powershell -ExecutionPolicy Bypass -File "gpo.ps1"
```

O script solicita elevacao automatica caso nao esteja rodando como administrador.

## Estrutura do Projeto

```
gpo.ps1              Script principal da aplicacao
GPO Manager.bat      Atalho para execucao rapida
apps_db.json         Catalogo local de aplicativos (gerado em runtime)
gpo_settings.json    Configuracoes persistentes do usuario (gerado em runtime)
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
