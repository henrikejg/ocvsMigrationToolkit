/**
 * OCVS Migration Dashboard — Backend Node.js
 * Sem dependências de framework — usa apenas http nativo + xlsx para ler Excel
 */

const http    = require("http");
const fs      = require("fs");
const path    = require("path");
const { execSync, spawn } = require("child_process");

const SCRIPTS_DIR = path.join(path.resolve(__dirname, ".."), "scripts");

// Resolver path curto via symlink se disponivel (evita problemas com paths longos com espacos)
function resolverPathCurto(p) {
  try {
    const home = require("os").homedir();
    const ocvsLink = path.join(home, "ocvs", "netstat", "V2");
    if (fs.existsSync(ocvsLink)) {
      const realOcvs = fs.realpathSync(ocvsLink);
      if (p.startsWith(realOcvs)) {
        return p.replace(realOcvs, ocvsLink);
      }
    }
  } catch {}
  return p;
}

// ── Caminhos base ─────────────────────────────────────────────────────────────
// server.js fica em V2/dashboard/ — base é V2/
const BASE_DIR    = path.resolve(__dirname, "..");
const DADOS_DIR   = path.join(BASE_DIR, "dados");
const PROCESSADOS = path.join(DADOS_DIR, "PROCESSADOS");
const CLIENT_DIR  = path.join(__dirname, "client");

// Localizar Excel em V2/ — planilha deve estar na mesma pasta dos scripts
function encontrarExcel() {
  if (cache.has("excel_path")) return cache.get("excel_path");
  try {
    const files = fs.readdirSync(BASE_DIR).filter(f => f.toLowerCase().endsWith(".xlsx"));
    if (files.length > 0) {
      const found = path.join(BASE_DIR, files[0]);
      cache.set("excel_path", found);
      return found;
    }
  } catch {}
  return null;
}
// ── Cache em memória ──────────────────────────────────────────────────────────
const cache = new Map();

// ── Ler Excel com xlsx ────────────────────────────────────────────────────────
let XLSX = null;
function getXLSX() {
  if (!XLSX) XLSX = require("xlsx");
  return XLSX;
}

function lerAbaAplicacoes(excelPath) {
  if (cache.has("aplicacoes")) return cache.get("aplicacoes");
  try {
    const wb  = getXLSX().readFile(excelPath, { cellText: true, cellDates: false });
    const ws  = wb.Sheets["Aplicacoes"];
    if (!ws) { console.error("Aba 'Aplicacoes' não encontrada no Excel."); return {}; }
    const rows = getXLSX().utils.sheet_to_json(ws, { header: 1 });
    const mapa = {};
    for (let i = 1; i < rows.length; i++) {
      const [exec, app] = rows[i];
      // chave exata, lowercase — igual ao PROCV do Excel (case-insensitive)
      if (exec && app) mapa[String(exec).trim().toLowerCase()] = String(app).trim();
    }
    cache.set("aplicacoes", mapa);
    return mapa;
  } catch (e) {
    console.error("Erro ao ler aba servicos/Aplicacoes:", e.message);
    return {};
  }
}

function lerMapaOndas(excelPath) {
  if (cache.has("mapa_ondas")) return cache.get("mapa_ondas");
  try {
    const wb  = getXLSX().readFile(excelPath, { cellText: true, cellDates: false });
    const ws  = wb.Sheets["vInfo"];
    if (!ws) return {};
    const rows = getXLSX().utils.sheet_to_json(ws, { header: 1 });

    let colVM = -1, colIP = -1, colOnda = -1;
    const header = rows[0] || [];
    header.forEach((v, i) => {
      const s = String(v || "").trim();
      if (s === "VM")                          colVM   = i;
      else if (/IP|Address/i.test(s))          colIP   = i;
      else if (s === "ONDA")                   colOnda = i;
    });

    const mapa = {};
    for (let i = 1; i < rows.length; i++) {
      const row  = rows[i];
      const vm   = String(row[colVM]   || "").trim();
      const ip   = String(row[colIP]   || "").trim();
      const onda = String(row[colOnda] || "").trim();
      if (vm   && onda && vm   !== "None") mapa[vm.toUpperCase()] = onda;
      if (ip   && onda && ip   !== "None") mapa[ip]               = onda;
    }
    cache.set("mapa_ondas", mapa);
    return mapa;
  } catch (e) {
    console.error("Erro ao ler aba vInfo:", e.message);
    return {};
  }
}

// ── Listar ondas disponíveis ──────────────────────────────────────────────────
function listarOndas() {
  try {
    return fs.readdirSync(PROCESSADOS)
      .filter(f => /^ONDA\w+_processado\.txt$/.test(f))
      .map(f => f.match(/^ONDA(\w+)_processado\.txt$/)[1])
      .sort();
  } catch { return []; }
}

// ── Ler e enriquecer arquivo processado ──────────────────────────────────────
function lerProcessado(numeroOnda) {
  const key = `processado_${numeroOnda}`;
  if (cache.has(key)) return cache.get(key);

  const arquivo = path.join(PROCESSADOS, `ONDA${numeroOnda}_processado.txt`);
  if (!fs.existsSync(arquivo)) return null;

  const excelPath  = encontrarExcel();
  const aplicacoes = excelPath ? lerAbaAplicacoes(excelPath) : {};
  const mapaOndas  = excelPath ? lerMapaOndas(excelPath)     : {};

  const linhas = [];
  const conteudo = fs.readFileSync(arquivo, "utf-8");

  for (const linha of conteudo.split(/\r?\n/)) {
    if (!linha.trim()) continue;
    const c = linha.split(";");
    while (c.length < 17) c.push("");

    const processo   = (c[7]  || "").trim();
    const ipRemoto   = (c[11] || "").trim();
    const hostname   = (c[1]  || "").trim();
    const contadorRaw = (c[16] || "1").replace(/\D/g, "") || "1";

    // Enriquecer aplicação — busca exata pelo executável (igual ao PROCV do Excel)
    const procKey   = processo.trim().toLowerCase();
    const aplicacao = aplicacoes[procKey] || "Falta Identificar";

    // Enriquecer ondas
    const ondaOrigem  = mapaOndas[hostname.toUpperCase()] || "";
    const ondaDestino = mapaOndas[ipRemoto] || mapaOndas[ipRemoto.toUpperCase()] || "";

    linhas.push({
      data:        c[0],
      hostname,
      proto:       c[2],
      local:       c[3],
      remoto:      c[4],
      estado:      c[5],
      pid:         c[6],
      processo,
      aplicacao,
      ip_local:    c[9],
      porta_local: c[10],
      ip_remoto:   ipRemoto,
      porta_remota: c[12],
      direcao:     c[13],
      ocvs:        c[14],
      portas_fmt:  c[15],
      contador:    parseInt(contadorRaw, 10),
      onda_origem:  ondaOrigem,
      onda_destino: ondaDestino,
    });
  }

  cache.set(key, linhas);
  return linhas;
}

// ── Lógica de análise ─────────────────────────────────────────────────────────
function calcDependenciasExternas(dados, numeroOnda) {
  const grupos = new Map();
  for (const r of dados) {
    if (r.ocvs !== "OCVS") continue;
    if (!r.onda_destino || r.onda_destino === r.onda_origem) continue;

    const chave = `${r.hostname}|${r.ip_remoto}|${r.porta_remota}|${r.aplicacao || r.processo}`;
    if (!grupos.has(chave)) {
      grupos.set(chave, {
        hostname:     r.hostname,
        ip_remoto:    r.ip_remoto,
        porta_remota: r.porta_remota,
        aplicacao:    r.aplicacao || r.processo,
        onda_origem:  r.onda_origem,
        onda_destino: r.onda_destino,
        ocvs:         r.ocvs,
        direcao:      r.direcao,
        contador:     0,
      });
    }
    grupos.get(chave).contador += r.contador;
  }
  return [...grupos.values()].sort((a, b) => b.contador - a.contador);
}

function calcStatusMigracao(dados) {
  const status = {
    mesma_onda:    { label: "Mesma onda",      count: 0, conexoes: 0, cor: "#22c55e" },
    onda_anterior: { label: "Onda anterior",   count: 0, conexoes: 0, cor: "#3b82f6" },
    onda_futura:   { label: "Onda futura",     count: 0, conexoes: 0, cor: "#f59e0b" },
    fora_ocvs:     { label: "Fora do OCVS",    count: 0, conexoes: 0, cor: "#6b7280" },
    nao_mapeado:   { label: "Sem Onda Agendada", count: 0, conexoes: 0, cor: "#ef4444" },
  };

  for (const r of dados) {
    const c = r.contador;
    if (r.ocvs !== "OCVS") { status.fora_ocvs.count++; status.fora_ocvs.conexoes += c; continue; }

    const od = r.onda_destino;
    const oo = r.onda_origem;
    if (!od || od === "FORA DE OCVS") { status.nao_mapeado.count++; status.nao_mapeado.conexoes += c; continue; }
    if (od === oo)                    { status.mesma_onda.count++;   status.mesma_onda.conexoes   += c; continue; }

    const mDest = od.match(/Onda\s+(\d+)/i);
    const mOrig = oo.match(/Onda\s+(\d+)/i);
    if (mDest && mOrig) {
      const nDest = parseInt(mDest[1]);
      const nOrig = parseInt(mOrig[1]);
      if (nDest === nOrig)      { status.mesma_onda.count++;    status.mesma_onda.conexoes    += c; }
      else if (nDest < nOrig)   { status.onda_anterior.count++; status.onda_anterior.conexoes += c; }
      else                      { status.onda_futura.count++;   status.onda_futura.conexoes   += c; }
    } else { status.nao_mapeado.count++; status.nao_mapeado.conexoes += c; }
  }
  return status;
}

function calcTopComunicacoes(dados) {
  const grupos = new Map();
  for (const r of dados) {
    const chave = `${r.hostname}|${r.ip_remoto}|${r.porta_remota}|${r.direcao}|${r.aplicacao || r.processo}`;
    if (!grupos.has(chave)) {
      grupos.set(chave, {
        hostname:     r.hostname,
        ip_remoto:    r.ip_remoto,
        porta_remota: r.porta_remota,
        direcao:      r.direcao,
        aplicacao:    r.aplicacao || r.processo,
        ocvs:         r.ocvs,
        onda_destino: r.onda_destino,
        contador:     0,
      });
    }
    grupos.get(chave).contador += r.contador;
  }
  return [...grupos.values()].sort((a, b) => b.contador - a.contador).slice(0, 50);
}

function calcServidoresOrigem(dados, apenasSemdOnda = false, esconderDispensaveis = false) {
  const IPS_DISPENSAVEIS = new Set([
    "10.62.169.11", "10.62.169.12", "10.62.169.13", "10.62.169.14", "10.62.169.25"
  ]);

  // Filtrar apenas ESTABLISHED e SYN_SENT
  let filtrado = dados.filter(r => r.estado === "ESTABLISHED" || r.estado === "SYN_SENT");

  if (esconderDispensaveis) {
    filtrado = filtrado.filter(r => !IPS_DISPENSAVEIS.has(r.ip_remoto));
  }

  // Filtro opcional: apenas conexoes sem onda agendada no destino
  // Mesmo critério do calcStatusMigracao para categoria "nao_mapeado"
  if (apenasSemdOnda) {
    filtrado = filtrado.filter(r => {
      const od = r.onda_destino || "";
      if (!od || od === "FORA DE OCVS" || od === "A definir") return true;
      // Sem número de onda extraível = sem onda agendada
      return !/Onda\s+\d+/i.test(od);
    });
  }

  // Estrutura: hostname → aplicacao → ip_remoto → porta_fmt → { ESTABLISHED, SYN_SENT }
  const arvore = new Map();

  for (const r of filtrado) {
    const h = r.hostname;
    const a = r.aplicacao || r.processo || "-";
    const ip = r.ip_remoto;
    const pf = r.portas_fmt || `L ${r.porta_local} | R ${r.porta_remota}`;
    const est = r.estado;
    const cnt = r.contador;

    if (!arvore.has(h)) arvore.set(h, { hostname: h, ESTABLISHED: 0, SYN_SENT: 0, aplicacoes: new Map() });
    const nH = arvore.get(h);
    nH[est] = (nH[est] || 0) + cnt;

    if (!nH.aplicacoes.has(a)) nH.aplicacoes.set(a, { nome: a, ESTABLISHED: 0, SYN_SENT: 0, ips: new Map() });
    const nA = nH.aplicacoes.get(a);
    nA[est] = (nA[est] || 0) + cnt;

    if (!nA.ips.has(ip)) nA.ips.set(ip, { ip, ESTABLISHED: 0, SYN_SENT: 0, portas: new Map() });
    const nI = nA.ips.get(ip);
    nI[est] = (nI[est] || 0) + cnt;

    if (!nI.portas.has(pf)) nI.portas.set(pf, { fmt: pf, ESTABLISHED: 0, SYN_SENT: 0 });
    const nP = nI.portas.get(pf);
    nP[est] = (nP[est] || 0) + cnt;
  }

  // Serializar Maps para JSON
  return [...arvore.values()]
    .sort((a, b) => (b.ESTABLISHED + b.SYN_SENT) - (a.ESTABLISHED + a.SYN_SENT))
    .map(h => ({
      hostname:    h.hostname,
      ESTABLISHED: h.ESTABLISHED,
      SYN_SENT:    h.SYN_SENT,
      aplicacoes: [...h.aplicacoes.values()]
        .sort((a, b) => (b.ESTABLISHED + b.SYN_SENT) - (a.ESTABLISHED + a.SYN_SENT))
        .map(a => ({
          nome:        a.nome,
          ESTABLISHED: a.ESTABLISHED,
          SYN_SENT:    a.SYN_SENT,
          ips: [...a.ips.values()]
            .sort((a, b) => (b.ESTABLISHED + b.SYN_SENT) - (a.ESTABLISHED + a.SYN_SENT))
            .map(i => ({
              ip:          i.ip,
              ESTABLISHED: i.ESTABLISHED,
              SYN_SENT:    i.SYN_SENT,
              portas: [...i.portas.values()]
                .sort((a, b) => (b.ESTABLISHED + b.SYN_SENT) - (a.ESTABLISHED + a.SYN_SENT)),
            })),
        })),
    }));
}

function calcDrilldown(dados, categoria) {
  // Retorna linhas agrupadas por (hostname, ip_remoto, porta_remota, aplicacao)
  // filtradas pela categoria do status de migração
  const grupos = new Map();

  for (const r of dados) {
    let cat = null;
    if (r.ocvs !== "OCVS") {
      cat = "fora_ocvs";
    } else {
      const od = r.onda_destino;
      const oo = r.onda_origem;
      if (!od || od === "FORA DE OCVS") {
        cat = "nao_mapeado";
      } else if (od === oo) {
        cat = "mesma_onda";
      } else {
        const mDest = od.match(/Onda\s+(\d+)/i);
        const mOrig = oo.match(/Onda\s+(\d+)/i);
        if (mDest && mOrig) {
          const nDest = parseInt(mDest[1]);
          const nOrig = parseInt(mOrig[1]);
          if (nDest === nOrig)    cat = "mesma_onda";
          else if (nDest < nOrig) cat = "onda_anterior";
          else                    cat = "onda_futura";
        } else {
          cat = "nao_mapeado";
        }
      }
    }

    if (cat !== categoria) continue;

    const chave = `${r.hostname}|${r.ip_remoto}|${r.porta_remota}|${r.aplicacao || r.processo}`;
    if (!grupos.has(chave)) {
      grupos.set(chave, {
        hostname:     r.hostname,
        ip_remoto:    r.ip_remoto,
        porta_remota: r.porta_remota,
        aplicacao:    r.aplicacao || r.processo,
        processo:     r.processo,
        direcao:      r.direcao,
        onda_origem:  r.onda_origem,
        onda_destino: r.onda_destino,
        ocvs:         r.ocvs,
        estado:       r.estado,
        contador:     0,
      });
    }
    grupos.get(chave).contador += r.contador;
  }

  return [...grupos.values()].sort((a, b) => b.contador - a.contador);
}

function calcGrafo(dados) {
  const nos = new Map();
  const arestas = new Map();

  for (const r of dados) {
    const src = r.hostname;
    const dst = r.ip_remoto;
    if (!src || !dst) continue;

    if (!nos.has(src)) nos.set(src, { id: src, tipo: "origem",  onda: r.onda_origem,  ocvs: "OCVS" });
    if (!nos.has(dst)) nos.set(dst, { id: dst, tipo: "destino", onda: r.onda_destino, ocvs: r.ocvs });

    const chave = `${src}→${dst}:${r.porta_remota}`;
    if (!arestas.has(chave)) {
      arestas.set(chave, {
        source:   src,
        target:   dst,
        porta:    r.porta_remota,
        aplicacao: r.aplicacao || r.processo,
        ocvs:     r.ocvs,
        contador: 0,
      });
    }
    arestas.get(chave).contador += r.contador;
  }

  return {
    nodes: [...nos.values()],
    edges: [...arestas.values()].sort((a, b) => b.contador - a.contador).slice(0, 300),
  };
}

// ── Roteador HTTP ─────────────────────────────────────────────────────────────
function respJson(res, data, status = 200) {
  const body = JSON.stringify(data);
  res.writeHead(status, { "Content-Type": "application/json; charset=utf-8",
                           "Access-Control-Allow-Origin": "*" });
  res.end(body);
}

function respFile(res, filePath) {
  const ext = path.extname(filePath).toLowerCase();
  const mime = { ".html": "text/html", ".js": "application/javascript",
                 ".css": "text/css", ".json": "application/json",
                 ".svg": "image/svg+xml", ".ico": "image/x-icon" };
  try {
    const data = fs.readFileSync(filePath);
    res.writeHead(200, { "Content-Type": mime[ext] || "text/plain" });
    res.end(data);
  } catch {
    res.writeHead(404); res.end("Not found");
  }
}

const server = http.createServer((req, res) => {
  const [urlPath, queryStr] = req.url.split("?");
  const params = new URLSearchParams(queryStr || "");

  // ── API ──
  if (urlPath === "/api/ondas") {
    return respJson(res, listarOndas());
  }

  if (urlPath === "/api/excel") {
    const p = encontrarExcel();
    return respJson(res, { path: p, found: !!p });
  }

  if (urlPath === "/api/cache/clear") {
    cache.clear();
    return respJson(res, { ok: true });
  }

  const mOnda = urlPath.match(/^\/api\/onda\/(\w+)\/(.+)$/);
  if (mOnda) {
    const [, numero, endpoint] = mOnda;
    let dados = lerProcessado(numero);
    if (!dados) return respJson(res, { erro: "Onda não encontrada" }, 404);

    // Filtro coluna O (ocvs): "OCVS", "FORA" ou omitido (todos)
    const filtroOcvs = params.get("ocvs"); // "OCVS" | "FORA" | null
    if (filtroOcvs === "OCVS") {
      dados = dados.filter(r => r.ocvs === "OCVS");
    } else if (filtroOcvs === "FORA") {
      dados = dados.filter(r => r.ocvs !== "OCVS");
    }
    // null = todos (sem filtro)

    if (endpoint === "dependencias-externas") return respJson(res, calcDependenciasExternas(dados, numero));
    if (endpoint === "status-migracao")       return respJson(res, calcStatusMigracao(dados));
    if (endpoint === "top-comunicacoes")      return respJson(res, calcTopComunicacoes(dados));
    if (endpoint === "grafo")                 return respJson(res, calcGrafo(dados));
    if (endpoint === "servidores-origem") {
      const semOnda            = params.get("semOnda")            === "1";
      const esconderDispensaveis = params.get("esconderDispensaveis") === "1";
      return respJson(res, calcServidoresOrigem(dados, semOnda, esconderDispensaveis));
    }
    if (endpoint === "resumo") {
      const excelPath = encontrarExcel();
      let servidoresPrevistos = 0;
      if (excelPath) {
        try {
          const wb  = getXLSX().readFile(excelPath, { cellText: true, cellDates: false });
          const ws  = wb.Sheets["vInfo"];
          if (ws) {
            const rows = getXLSX().utils.sheet_to_json(ws, { header: 1 });
            const header = rows[0] || [];
            let colVM = -1, colOnda = -1;
            header.forEach((v, i) => {
              const s = String(v || "").trim();
              if (s === "VM")   colVM   = i;
              if (s === "ONDA") colOnda = i;
            });
            if (colVM >= 0 && colOnda >= 0) {
              servidoresPrevistos = rows.slice(1).filter(r => {
                const vm   = String(r[colVM]   || "").trim();
                const onda = String(r[colOnda] || "").trim();
                return vm && vm !== "None" && vm !== "A definir" &&
                       new RegExp(`Onda\\s+${numero}\\b`, "i").test(onda);
              }).length;
            }
          }
        } catch {}
      }
      const servidoresProcessados = [...new Set(dados.map(r => r.hostname))].length;
      return respJson(res, {
        total_linhas:          dados.length,
        total_conexoes:        dados.reduce((s, r) => s + r.contador, 0),
        servidores:            servidoresProcessados,
        servidores_previstos:  servidoresPrevistos,
        ips_remotos:           [...new Set(dados.map(r => r.ip_remoto))].length,
      });
    }
    // drilldown: /api/onda/99/drilldown/nao_mapeado
    const mDrill = endpoint.match(/^drilldown\/(\w+)$/);
    if (mDrill) return respJson(res, calcDrilldown(dados, mDrill[1]));
    return respJson(res, { erro: "Endpoint desconhecido" }, 404);
  }

  // Execucao de scripts PowerShell com streaming via SSE
  if (urlPath === "/api/executar" && req.method === "POST") {
    let body = "";
    req.on("data", d => body += d);
    req.on("end", () => {
      let params;
      try { params = JSON.parse(body); } catch { res.writeHead(400); res.end("JSON invalido"); return; }

      const { script, onda, usuario, senha } = params;
      const scriptMap = { "coletar": "Coletar-Linux.ps1", "processar": "Processar-Onda.ps1" };
      const scriptFile = scriptMap[script];
      if (!scriptFile) { res.writeHead(400); res.end("Script desconhecido"); return; }

      const scriptPath = resolverPathCurto(path.join(SCRIPTS_DIR, scriptFile));
      const scriptsCwd = resolverPathCurto(SCRIPTS_DIR);
      if (!fs.existsSync(scriptPath)) { res.writeHead(404); res.end("Script nao encontrado: " + scriptPath); return; }

      res.writeHead(200, {
        "Content-Type":  "text/event-stream",
        "Cache-Control": "no-cache",
        "Connection":    "keep-alive",
        "Access-Control-Allow-Origin": "*",
      });
      const sendLine = (line, tipo) => {
        try { res.write("data: " + JSON.stringify({ tipo: tipo || "log", linha: line }) + "\n\n"); } catch {}
      };

      const senhaArg   = (script === "coletar" && senha)   ? " -Senha '" + senha.replace(/'/g, "''") + "'" : "";
      const usuarioArg = (script === "coletar" && usuario) ? " -Usuario '" + usuario + "'" : "";

      const PWSH_REAL = process.env.OCVS_PWSH && fs.existsSync(process.env.OCVS_PWSH)
        ? process.env.OCVS_PWSH
        : (() => { try { return require("child_process").execSync("where pwsh", {encoding:"utf8"}).trim().split(/\r?\n/)[0]; } catch {} return "powershell.exe"; })();

      const tmpScript = path.join(require("os").tmpdir(), "ocvs_" + Date.now() + ".ps1");
      fs.writeFileSync(tmpScript,
        "& '" + scriptPath.replace(/'/g, "''") + "' -NumeroOnda '" + onda + "'" + usuarioArg + senhaArg + "\n",
        "utf8");

      sendLine("Iniciando " + scriptFile + " para Onda " + onda + "...", "info");
      console.log("[exec]", PWSH_REAL, tmpScript, "cwd=" + scriptsCwd);

      // Usar detached:true para desacoplar do processo pai (Node filho de PS7)
      const proc = spawn(PWSH_REAL,
        ["-NoProfile", "-ExecutionPolicy", "Bypass", "-File", tmpScript],
        { cwd: scriptsCwd, windowsHide: true, detached: false, stdio: ["ignore", "pipe", "pipe"] }
      );

      proc.stdout.on("data", d => d.toString().split(/\r?\n/).filter(Boolean).forEach(l => sendLine(l)));
      proc.stderr.on("data", d => d.toString().split(/\r?\n/).filter(Boolean).forEach(l => sendLine(l, "erro")));
      proc.on("error", err => sendLine("Erro: " + err.message, "erro"));
      proc.on("close", code => {
        try { fs.unlinkSync(tmpScript); } catch {}
        console.log("[close]", code);
        sendLine("Processo encerrado (exit " + code + ")", code === 0 ? "sucesso" : "erro");
        try { res.write("data: {\"tipo\":\"fim\"}\n\n"); res.end(); } catch {}
        cache.delete("processado_" + onda);
      });
      // Nao matar o processo se o browser fechar — deixar completar
      // req.on("close", () => { try { proc.kill(); } catch {} });
    });
    return;
  }

  // ── Static files ──
  if (urlPath === "/" || urlPath === "/index.html") return respFile(res, path.join(CLIENT_DIR, "index.html"));
  const staticPath = path.join(CLIENT_DIR, urlPath);
  if (fs.existsSync(staticPath) && fs.statSync(staticPath).isFile()) return respFile(res, staticPath);

  res.writeHead(404); res.end("Not found");
});

const PORT = 5000;
server.listen(PORT, "127.0.0.1", () => {
  const excelPath = encontrarExcel();
  console.log(`\n========================================`);
  console.log(` OCVS Migration Dashboard`);
  console.log(`========================================`);
  console.log(` URL:   http://localhost:${PORT}`);
  console.log(` Base:  ${BASE_DIR}`);
  console.log(` Excel: ${excelPath || "NÃO ENCONTRADO"}`);
  console.log(` Ondas: ${listarOndas().join(", ") || "nenhuma"}`);
  console.log(`========================================\n`);
  // Abrir browser automaticamente
  try { execSync(`start http://localhost:${PORT}`); } catch {}
});
