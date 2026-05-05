---
inclusion: auto
---

# Regras de Desenvolvimento

## Versionamento
- Sempre incrementar versão antes de commitar
- PATCH para fixes e melhorias pequenas
- MINOR para features novas que não quebram compatibilidade
- Atualizar versão em: README.md, index.html (title + h1 small), server.js (console.log), Iniciar-Dashboard.ps1 (synopsis)
- Não commitar sem validação do usuário — implementar, aguardar OK, depois commitar

## Git
- Push direto na main
- Mensagens de commit em português, formato: "tipo: vX.Y.Z — descrição"
- Tipos: feat, fix, docs
- Nunca commitar dados sensíveis (IPs internos, paths pessoais, senhas)

## Código
- HTML/JS: tudo em um único arquivo (dashboard/client/index.html)
- Backend: dashboard/server.js (Node.js puro, sem frameworks)
- Banco: dashboard/db.js (sql.js WebAssembly)
- Scripts: PowerShell 7 em scripts/*.ps1, AWK em scripts/*.awk
- Sem dependências externas além de sql.js e xlsx no package.json
- Ícones via Lucide CDN (unpkg)
- Chamar lucide.createIcons() após inserir HTML com data-lucide

## Interface
- Dark theme (#0f172a background)
- Sidebar com ícones (48px recolhida, 240px expandida)
- Header: hambúrguer + título clicável (volta home) + seletor de onda + badge SQLite
- Tabelas com colunas ordenáveis (clique no header)
- Seleção respeita filtro (só linhas visíveis)
- Modais de confirmação para ações destrutivas/pesadas

## Excel
- Fonte da verdade para composição de ondas e dados de servidores
- Lido via xlsx (Node) ou Import-Excel (PowerShell)
- Path configurável via dialog nativo (salvo em dados/config.json)
- PowerShell usa robocopy para copiar antes de ler (evita lock do Excel aberto)
- Colunas identificadas por nome do header (não por posição/letra)

## Processamento
- Granularidade: por servidor (não por onda)
- Incremental: só processa linhas novas (compara timestamp)
- Pré-filtro: separa público/privado antes de enriquecer
- Controle: dados/controle/{HOSTNAME}.json
- Ingestão no SQLite: após processar, via botão "Carregar no banco"
