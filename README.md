# OCVS Migration Toolkit — v0.5.1

Ferramentas para coleta, processamento e análise de dependências de rede (camada 4) durante migração de VMs OCVS para IaaS.

---

## Pré-requisitos

### 1. PowerShell 7 (recomendado)
Baixe o installer `.msi` em https://github.com/PowerShell/PowerShell/releases  
O toolkit funciona com PS5 mas PS7 é necessário para executar scripts pelo dashboard.

### 2. Módulo ImportExcel
```powershell
Install-Module ImportExcel -Scope CurrentUser
```

### 3. OpenSSH (coleta de servidores Linux)
Já incluso no Windows 10/11. Para verificar:
```powershell
Get-WindowsCapability -Online -Name OpenSSH.Client*
```

Se o resultado mostrar `State: NotPresent`, instale com:
```powershell
Add-WindowsCapability -Online -Name OpenSSH.Client~~~~0.0.1.0
```

### 4. Node.js LTS (dashboard web)
Baixe o installer `.msi` em https://nodejs.org  
Necessário apenas para rodar o dashboard de análise. As dependências (sql.js, xlsx) são instaladas automaticamente na primeira execução.

### 5. Política de execução do PowerShell
O script de inicialização tenta ajustar automaticamente. Se falhar, execute uma vez como Administrador:
```powershell
Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned
```

---

## Estrutura de pastas

```
<raiz>/
  scripts/            <- scripts PowerShell e AWK
  dashboard/          <- servidor Node + interface web
    client/           <- HTML/JS do frontend (single-page)
    db.js             <- camada SQLite (sql.js/WebAssembly)
    server.js         <- backend HTTP nativo
  dados/
    raw/              <- arquivos netstat coletados (netstat_*.txt)
    ONDAXX/
      COLETAS/        <- netstat_*.txt organizados por onda
      CONSOLIDADO/    <- arquivo consolidado intermediário
    PROCESSADOS/      <- ONDAXX_processado.txt (saída final)
    ocvs.db           <- banco SQLite (gerado automaticamente)
    config.json       <- configuração local (path do Excel, etc.)
    logs/             <- logs de execução (JSON)
  README.md
```

A planilha Excel pode estar em qualquer local da máquina ou rede. Na primeira execução, o dashboard solicita a seleção do arquivo via dialog nativo do Windows. O path é salvo em `dados/config.json` e persiste entre reinícios. Para alterar, use o botão "Alterar" na tela inicial.

A planilha precisa conter:
- Aba **vInfo**: colunas `VM`, `Primary IP Address`, `ONDA`
- Aba **Aplicacoes**: colunas `Executável`, `Aplicação` (mapeamento de processos para nomes amigáveis)
- Aba **VARIAVEIS** (opcional): colunas `Variavel`, `Valor` — configurações dinâmicas:

| Variavel | Valor | Descrição |
|----------|-------|-----------|
| IGNORAR_AD_ZABBIX | 10.0.0.1 | IPs a ignorar no filtro AD/Zabbix (uma linha por IP) |
| IGNORAR_AD_ZABBIX | 10.0.0.2 | |
| RANGES_OCVS | 10.0.0.0/22 | Ranges de rede internos em notação CIDR (uma linha por range) |
| RANGES_OCVS | 10.0.4.0/24 | |

Se a aba VARIAVEIS não existir, o dashboard usa valores padrão internos.

---

## Início rápido

```powershell
cd dashboard
.\Iniciar-Dashboard.ps1
```

O browser abre automaticamente em `http://localhost:5000`. Na primeira execução as dependências são instaladas automaticamente.

---

## Fluxo de uso

```
1. Coletar  →  2. Processar  →  3. Analisar (dashboard)
```

Todas as etapas podem ser executadas pela interface web ou por linha de comando.

### 1. Coletar netstat dos servidores

**Pela interface:** Menu lateral → Coletar Onda (ou botão na tela inicial para ondas pendentes)

**Por linha de comando:**
```powershell
cd scripts
.\Coletar-Linux.ps1 -NumeroOnda 2
```

- Lê os IPs da onda na planilha Excel
- Detecta SO via TTL do ping (Linux: TTL 50-64 ou 240-255)
- Conecta via SSH, comprime com tar+gzip e transfere
- Salva os arquivos em `dados\raw\`
- Solicita credenciais via prompt seguro do Windows (sem arquivo de senha em disco)

Para servidores com SSH legado (algoritmos antigos), o script de inicialização do dashboard cria automaticamente o arquivo `%USERPROFILE%\.ssh\config` com as opções de compatibilidade necessárias (se ainda não existir). O conteúdo gerado é:
```
Host *
    KexAlgorithms +diffie-hellman-group1-sha1
    HostKeyAlgorithms +ssh-rsa
    Ciphers +aes128-cbc
    StrictHostKeyChecking no
```

> Servidores Windows precisam ser coletados manualmente (sem SSH/WinRM disponível).  
> Copie os arquivos `netstat_*.txt` para `dados\raw\` após a coleta manual.

### 2. Processar e consolidar uma onda

**Pela interface:** Menu lateral → Processar Onda

**Por linha de comando:**
```powershell
cd scripts
.\Processar-Onda.ps1 -NumeroOnda 2
```

- Copia os arquivos da onda de `dados\raw\` para `dados\ONDA2\COLETAS\`
- Executa os scripts AWK para enriquecer e aglutinar as conexões
- Gera `dados\PROCESSADOS\ONDA2_processado.txt`
- Ingere os dados automaticamente no banco SQLite

A aglutinação reduz o volume de dados agrupando conexões semelhantes em uma única linha com contador — essencial para análise de grandes volumes.

### 3. Analisar dependências (dashboard)

Selecione a onda no dropdown do header. Todas as visões são carregadas automaticamente a partir do banco SQLite.

---

## Dashboard — interface web

### Tela inicial

Ao abrir o dashboard, a tela inicial mostra uma visão geral de todas as ondas:
- Cards com totais: ondas no Excel, processadas, ingeridas no banco, última execução
- Tabela de ondas com status e botões contextuais (Coletar / Processar / Abrir)

### Menu lateral

Acessível pelo ícone de menu (☰) no canto superior esquerdo. Empurra o conteúdo e pode ficar aberto durante a sessão.

| Seção | Ação | Descrição |
|-------|------|-----------|
| Coleta e Processamento | Coletar Onda | Coleta netstat via SSH com output em tempo real |
| | Processar Onda | Consolida dados e ingere no banco automaticamente |
| Banco de Dados | Atualizar Ondas | Lê o Excel e sincroniza a composição de todas as ondas no banco |
| | Reingerir Banco | Reconstrói o banco a partir dos arquivos .txt processados |
| Sistema | Recarregar Cache | Limpa o cache em memória e força releitura |
| | Logs de Execução | Histórico de coletas e processamentos com log completo |

### Visões de análise (por onda)

| Painel | Descrição |
|--------|-----------|
| **Servidores de Origem** | Hierarquia Servidor → Aplicação → IP Remoto → Portas. Filtros para ignorar onda atual/passadas e AD/Zabbix. Indicador de maturidade quando não há dependências críticas. |
| **Visão Geral** | Gráfico donut com distribuição: rede privada externa, IP público, ondas anteriores, mesma onda, problemas. Clique em qualquer fatia para drilldown detalhado. |
| **Top Comunicações** | 15 maiores fluxos por volume de conexões. |

### Header

- Seletor de onda e tipo de conexão (OCVS / Fora do OCVS / Todos)
- Badge de fonte dos dados: verde (SQLite) ou amarelo (.txt)
- Clique no título para voltar à tela inicial

---

## Banco de dados SQLite

A partir da v0.2.0, os dados processados são armazenados em um banco SQLite (`dados/ocvs.db`) via sql.js (WebAssembly — sem dependências nativas).

**Vantagens:**
- Alterações na composição das ondas no Excel são refletidas sem reprocessar os dados
- Consultas mais rápidas para ondas já ingeridas
- `onda_destino` resolvido via JOIN em tempo de consulta (sempre atualizado)

**Fluxo de atualização:**
- Alterou a onda de um servidor no Excel → clique em **Atualizar Ondas** no menu
- Coletou novos dados de netstat → **Processar Onda** (ingestão automática)
- Banco corrompido ou desatualizado → **Reingerir Banco**

O dashboard faz fallback automático para os arquivos `.txt` quando uma onda não está no banco.

---

## Scripts disponíveis

| Script | Descrição |
|--------|-----------|
| `Coletar-Linux.ps1` | Coleta netstat de servidores Linux via SSH |
| `Processar-Onda.ps1` | Organiza, consolida e ingere dados de uma onda |
| `Analisar-Dependencias.ps1` | Análise de dependências via linha de comando |
| `Extrair-IPs.ps1` | Extrai IPs de uma onda da planilha Excel |
| `Extrair-Hostnames.ps1` | Extrai hostnames de uma onda da planilha Excel |
| `dependencias_ocvs.awk` | Enriquece dados do netstat (chamado pelo Processar) |
| `aglutinar_ocvs.awk` | Aglutina conexões semelhantes (chamado pelo Processar) |
