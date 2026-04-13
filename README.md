# OCVS Migration Toolkit — v0.1

Ferramentas para coleta, processamento e análise de dependências de rede durante migração de VMs OCVS para IaaS.

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
Já incluso no Windows 10/11. Para verificar ou instalar:
```powershell
Get-WindowsCapability -Online -Name OpenSSH.Client*
Add-WindowsCapability -Online -Name OpenSSH.Client~~~~0.0.1.0
```

### 4. AWK (processamento dos dados coletados)
Necessário para o `Processar-Onda.ps1`. Opções:
- **Git for Windows** (recomendado — instala `awk.exe` automaticamente no PATH):  
  https://git-scm.com/download/win
- **WSL** (fallback automático se `awk.exe` não estiver no PATH)

### 5. Node.js LTS (dashboard web)
Baixe o installer `.msi` em https://nodejs.org  
Necessário apenas para rodar o dashboard de análise.

---

## Estrutura de pastas

```
<raiz>/
  scripts/          <- scripts PowerShell e AWK
  dashboard/        <- servidor Node + interface web
  dados/
    raw/            <- arquivos netstat coletados (netstat_*.txt)
    ONDAXX/
      COLETAS/      <- netstat_*.txt organizados por onda
      CONSOLIDADO/  <- arquivo consolidado intermediário
    PROCESSADOS/    <- ONDAXX_processado.txt (saída final para análise)
  SERVIDORES.xlsx   <- planilha com servidores e ondas (manter aqui)
  README.md
```

A planilha Excel deve ficar na pasta raiz e precisa conter:
- Aba **vInfo**: colunas `VM`, `Primary IP Address`, `ONDA`
- Aba **Aplicacoes**: colunas `Executável`, `Aplicação` (mapeamento de processos para nomes amigáveis)

---

## Início rápido — interface web

A forma mais simples de usar o toolkit é pelo dashboard web, que centraliza todas as operações:

```powershell
cd dashboard
.\Iniciar-Dashboard.ps1
```

O browser abre automaticamente em `http://localhost:5000`. A partir daí tudo pode ser feito pela interface:

- **Coletar** — botão no header, solicita número da onda e credenciais SSH, executa a coleta com output em tempo real
- **Processar** — botão no header, solicita número da onda, consolida os dados coletados
- **Analisar** — selecione a onda no dropdown e todas as visões são carregadas automaticamente

Os scripts de linha de comando continuam disponíveis para quem preferir automação ou integração com outros processos.

---

## Fluxo de uso

```
1. Coletar    ->  2. Processar    ->  3. Analisar (dashboard)
```

### 1. Coletar netstat dos servidores Linux

```powershell
cd scripts
.\Coletar-Linux.ps1 -NumeroOnda 2
```

- Lê os IPs da onda na planilha Excel
- Detecta SO via TTL do ping (Linux: TTL 50-64 ou 240-255)
- Conecta via SSH, comprime com tar+gzip e transfere
- Salva os arquivos em `dados\raw\`
- Solicita credenciais via prompt seguro do Windows (sem arquivo de senha em disco)

Para servidores com SSH legado (algoritmos antigos), o script já inclui as opções de compatibilidade necessárias. Se necessário, configure `%USERPROFILE%\.ssh\config`:
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

```powershell
cd scripts
.\Processar-Onda.ps1 -NumeroOnda 2
```

- Copia os arquivos da onda de `dados\raw\` para `dados\ONDA2\COLETAS\`
- Executa os scripts AWK para enriquecer e aglutinar as conexões
- Gera `dados\PROCESSADOS\ONDA2_processado.txt`

A aglutinação reduz o volume de dados agrupando conexões semelhantes em uma única linha com contador — essencial para análise de grandes volumes.

### 3. Analisar dependências (dashboard web)

```powershell
cd dashboard
.\Iniciar-Dashboard.ps1
```

Na primeira execução instala as dependências automaticamente. O browser abre em `http://localhost:5000`.

---

## Dashboard — visões disponíveis

| Painel | Descrição |
|--------|-----------|
| **Servidores de Origem** | Hierarquia Servidor → Aplicação → IP Remoto → Portas, filtrada por ESTABLISHED e SYN_SENT. Checkbox para mostrar apenas conexões sem onda agendada. |
| **Status de Migração** | Distribuição das conexões por categoria: mesma onda, onda anterior, onda futura, fora do OCVS, sem onda agendada. Clique em qualquer barra para ver o detalhe (drilldown). |
| **Top Comunicações** | 50 maiores fluxos por volume de conexões. |
| **Dependências Externas** | Servidores desta onda comunicando com IPs fora dela, com busca por hostname/IP/aplicação. |
| **Grafo de Dependências** | Visualização interativa com pan, zoom e filtros por tipo de conexão. |

O dashboard enriquece os dados automaticamente:
- Nome da aplicação via lookup na aba `Aplicacoes` da planilha
- Onda de origem e destino via lookup na aba `vInfo`

Os botões **Coletar** e **Processar** no header do dashboard permitem executar os scripts diretamente pela interface, com output em tempo real.

---

## Scripts disponíveis

| Script | Uso direto |
|--------|-----------|
| `Coletar-Linux.ps1` | Coleta netstat de servidores Linux via SSH |
| `Processar-Onda.ps1` | Organiza e consolida coletas de uma onda |
| `Analisar-Dependencias.ps1` | Análise de dependências via linha de comando |
| `Extrair-IPs.ps1` | Extrai IPs de uma onda da planilha Excel |
| `Extrair-Hostnames.ps1` | Extrai hostnames de uma onda da planilha Excel |
| `dependencias_ocvs.awk` | Enriquece dados do netstat (chamado pelo Processar) |
| `aglutinar_ocvs.awk` | Aglutina conexões semelhantes (chamado pelo Processar) |
