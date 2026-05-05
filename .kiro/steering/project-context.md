---
inclusion: auto
---

# OCVS Migration Toolkit — Contexto do Projeto

## O que é
Ferramenta para coleta, processamento e análise de dependências de rede (camada 4 — TCP) durante migração de VMs do ambiente OCVS (Oracle Cloud VMware Solution) para IaaS. Gera mapas de dependência de comunicação entre servidores para planejamento de ondas de migração.

## Arquitetura
- **Frontend:** HTML5 + Vanilla JS (single-page, sem frameworks)
- **Backend:** Node.js (http nativo, sem Express), porta 5000
- **Banco:** SQLite via sql.js (WebAssembly, sem dependências nativas)
- **Processamento:** AWK scripts + PowerShell 7
- **Excel:** Planilha externa (pode estar em qualquer local) como fonte da verdade para composição de ondas, servidores e variáveis

## Fluxo principal
```
Coletar (SSH) → raw/*.txt → Processar (AWK por servidor) → processados_servidor/*.txt → Ingerir no SQLite → Dashboard
```

## Estrutura de pastas
```
scripts/              — PowerShell e AWK
dashboard/            — Node.js server + client/index.html
  db.js              — camada SQLite
  server.js          — backend HTTP + rotas API
  client/index.html  — toda a UI (single file)
dados/
  raw/               — netstat cru coletado
  privado/           — filtrado (só IPs privados)
  publico/           — filtrado (só IPs públicos)
  processados_servidor/ — enriquecido + aglutinado por servidor
  controle/          — JSON de controle por servidor
  config.json        — path do Excel configurado
  ocvs.db            — banco SQLite
```

## Conceitos-chave
- **Onda:** grupo de servidores que migram juntos. Definida no Excel (aba vInfo, coluna ONDA)
- **Enriquecimento:** AWK adiciona direção (IN/OUT/ANALISAR), classificação OCVS, IP/porta separados
- **Aglutinação:** AWK agrupa conexões similares e conta ocorrências (reduz ~90% do volume)
- **Processamento incremental:** só processa linhas novas (por timestamp), merge com anterior
- **Pré-filtro:** separa IPs públicos dos privados antes do enriquecimento (economia de 50%+ no processamento)

## Excel — abas esperadas
- **vInfo:** VM, Primary IP Address, ONDA, SO REVISADO RESUMIDO, PROD/NÃO-PROD, Powerstate
- **Aplicacoes:** Executável, Aplicação (mapeamento processo → nome amigável)
- **VARIAVEIS:** Variavel, Valor (IGNORAR_AD_ZABBIX, RANGES_OCVS em CIDR /24)

## Telas do dashboard
- **Ondas de Migração** (home): visão geral de todas as ondas com status
- **Coletar Linux:** lista servidores Linux PoweredOn, coleta via SSH
- **Processar Servidores:** status de processamento, processar e carregar no banco
- **Inventário de Serviços:** servidores que proveem serviços por porta
- **Status SQLite:** cobertura do banco por servidor
- **Análise por onda:** Mapa de Conexões, gráficos de pizza (comunicações, SO), Top Comunicações

## Convenções de código
- Versão semântica (MAJOR.MINOR.PATCH) em 5 arquivos: README.md, index.html (title + h1), server.js, Iniciar-Dashboard.ps1
- Commits: sempre incrementar versão, mensagem descritiva em português
- Commit só após validação do usuário
- Git push direto na main
- .gitignore: *.txt, *.xlsx, *.db, dados/logs/, dados/config.json, dados/privado/, dados/publico/, dados/processados_servidor/, dados/controle/

## Ícones visuais (Lucide)
- 🛡 shield verde = servidor Produtivo
- </> code laranja = servidor Não-Produtivo
- ? circle-help azul = OCVS sem dados coletados
- ↗ external-link vermelho = Fora do OCVS

## Problemas conhecidos/resolvidos
- OneDrive causa lock no ocvs.db → salvarDB com retry (5 tentativas, delay crescente)
- SSH com chave alterada → UserKnownHostsFile /dev/null
- Excel aberto causa lock → robocopy para temp antes de ler (nos scripts PS1)
- AWK no Windows não aceita \ em -v → ConvertTo-AwkPath (troca \ por /)
