/**
 * OCVS Migration Dashboard — Backend Node.js
 * Sem dependências de framework — usa apenas http nativo + xlsx para ler Excel
 */

const http    = require("http");
const fs      = require("fs");
const path    = require("path");
const { execSync, spawn } = require("child_process");
const db      = require("./db");

const SCRIPTS_DIR = path.join(path.resolve(__dirname, ".."), "scripts");
const LOGS_DIR    = path.join(path.resolve(__dirname, ".."), "dados", "logs");

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
// ── Configuração persistente do Excel ──────────────────────────────────────────
const CONFIG_PATH = path.join(DADOS_DIR, "config.json");

function lerConfig() {
  try {
    if (fs.existsSync(CONFIG_PATH)) return JSON.parse(fs.readFileSync(CONFIG_PATH, "utf8"));
  } catch {}
  return {};
}

function salvarConfig(cfg) {
  const dir = path.dirname(CONFIG_PATH);
  if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
  fs.writeFileSync(CONFIG_PATH, JSON.stringify(cfg, null, 2), "utf8");
}

function encontrarExcel() {
  if (cache.has("excel_path")) return cache.get("excel_path");

  // 1. Path configurado pelo usuário
  const cfg = lerConfig();
  if (cfg.excelPath && fs.existsSync(cfg.excelPath)) {
    cache.set("excel_path", cfg.excelPath);
    return cfg.excelPath;
  }

  // 2. Fallback: buscar na pasta raiz do projeto
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
      if (exec && app) mapa[String(exec).trim().toLowerCase()] = String(app).trim();
    }
    cache.set("aplicacoes", mapa);
    return mapa;
  } catch (e) {
    console.error("Erro ao ler aba Aplicacoes:", e.message);
    return {};
  }
}

// ── Ler aba VARIAVEIS do Excel ────────────────────────────────────────────────
// Retorna { CHAVE: [valor1, valor2, ...], ... }
function lerVariaveis(excelPath) {
  if (cache.has("variaveis")) return cache.get("variaveis");
  try {
    const wb  = getXLSX().readFile(excelPath, { cellText: true, cellDates: false });
    const ws  = wb.Sheets["VARIAVEIS"];
    if (!ws) { console.warn("[config] Aba 'VARIAVEIS' não encontrada no Excel — usando defaults."); return cache.set("variaveis", null) || null; }
    const rows = getXLSX().utils.sheet_to_json(ws, { header: 1 });
    const mapa = {};
    for (let i = 1; i < rows.length; i++) {
      const chave = String(rows[i][0] || "").trim().toUpperCase();
      const valor = String(rows[i][1] || "").trim();
      if (!chave || !valor) continue;
      if (!mapa[chave]) mapa[chave] = [];
      mapa[chave].push(valor);
    }
    cache.set("variaveis", mapa);
    console.log(`[config] Variaveis carregadas: ${Object.keys(mapa).join(", ")}`);
    return mapa;
  } catch (e) {
    console.error("Erro ao ler aba VARIAVEIS:", e.message);
    return null;
  }
}

// Obter IPs dispensáveis (AD/Zabbix) — do Excel ou fallback hardcoded
function obterIpsDispensaveis() {
  const excelPath = encontrarExcel();
  const vars = excelPath ? lerVariaveis(excelPath) : null;
  if (vars && vars["IGNORAR_AD_ZABBIX"] && vars["IGNORAR_AD_ZABBIX"].length > 0) {
    return new Set(vars["IGNORAR_AD_ZABBIX"]);
  }
  // Fallback — será removido quando a aba VARIAVEIS estiver populada
  return new Set(["10.62.169.11", "10.62.169.12", "10.62.169.13", "10.62.169.14", "10.62.169.25"]);
}

// Obter ranges OCVS em notação CIDR — do Excel ou fallback hardcoded
function obterRangesOcvs() {
  const excelPath = encontrarExcel();
  const vars = excelPath ? lerVariaveis(excelPath) : null;
  if (vars && vars["RANGES_OCVS"] && vars["RANGES_OCVS"].length > 0) {
    return vars["RANGES_OCVS"];
  }
  // Fallback
  return ["10.62.160.0/22", "10.62.176.0/22", "10.62.184.0/24"];
}

// Verificar se um IP está dentro de um range CIDR
function ipEmCidr(ip, cidr) {
  const [rede, bits] = cidr.split("/");
  const mask = ~(0xFFFFFFFF >>> parseInt(bits)) >>> 0;
  const ipNum = ip.split(".").reduce((s, o) => (s << 8) + parseInt(o), 0) >>> 0;
  const redeNum = rede.split(".").reduce((s, o) => (s << 8) + parseInt(o), 0) >>> 0;
  return (ipNum & mask) === (redeNum & mask);
}

function ipEhOcvs(ip) {
  const ranges = obterRangesOcvs();
  return ranges.some(cidr => ipEmCidr(ip, cidr));
}

// Obter mapa hostname → ambiente (PROD / NÃO-PROD) do Excel
function obterMapaAmbiente() {
  if (cache.has("mapa_ambiente")) return cache.get("mapa_ambiente");
  const excelPath = encontrarExcel();
  if (!excelPath) return {};
  try {
    const wb  = getXLSX().readFile(excelPath, { cellText: true, cellDates: false });
    const ws  = wb.Sheets["vInfo"];
    if (!ws) return {};
    const rows = getXLSX().utils.sheet_to_json(ws, { header: 1 });
    const hdr  = rows[0] || [];
    let colVM = -1, colAmb = -1;
    hdr.forEach((v, i) => {
      const s = String(v || "").trim();
      if (s === "VM") colVM = i;
      else if (s === "PROD/NÃO-PROD" || s === "PROD/NAO-PROD") colAmb = i;
    });
    if (colVM < 0 || colAmb < 0) return {};
    const mapa = {};
    for (const row of rows.slice(1)) {
      const vm  = String(row[colVM]  || "").trim().toUpperCase();
      const amb = String(row[colAmb] || "").trim();
      if (vm && vm !== "NONE" && amb) mapa[vm] = amb;
    }
    cache.set("mapa_ambiente", mapa);
    return mapa;
  } catch { return {}; }
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

// ── Ler e enriquecer arquivo processado (async para nao bloquear event loop) ──
async function lerProcessadoAsync(numeroOnda) {
  const key = `processado_${numeroOnda}`;
  if (cache.has(key)) return cache.get(key);

  const arquivo = path.join(PROCESSADOS, `ONDA${numeroOnda}_processado.txt`);
  if (!fs.existsSync(arquivo)) return null;

  const excelPath  = encontrarExcel();
  const aplicacoes = excelPath ? lerAbaAplicacoes(excelPath) : {};
  const mapaOndas  = excelPath ? lerMapaOndas(excelPath)     : {};

  // Pre-computar mapa sem dominio para lookup rapido
  const mapaOndaSemDominio = {};
  for (const k of Object.keys(mapaOndas)) {
    const semDom = k.split(".")[0];
    if (!mapaOndaSemDominio[semDom]) mapaOndaSemDominio[semDom] = mapaOndas[k];
  }

  const buf      = fs.readFileSync(arquivo);
  const rawLinhas = buf.toString("utf-8").split("\n");
  const linhas   = [];
  const LOTE = 20000;

  for (let i = 0; i < rawLinhas.length; i++) {
    // Ceder o event loop a cada lote para nao bloquear
    if (i > 0 && i % LOTE === 0) {
      await new Promise(r => setImmediate(r));
    }

    const linha = rawLinhas[i].trimEnd();
    if (!linha) continue;
    const c = linha.split(";");
    while (c.length < 17) c.push("");

    const processo    = (c[7]  || "").trim();
    const ipRemoto    = (c[11] || "").trim();
    const hostname    = (c[1]  || "").trim();
    const contadorRaw = (c[16] || "1").replace(/\D/g, "") || "1";

    const procKey   = processo.toLowerCase();
    const aplicacao = aplicacoes[procKey] || "Falta Identificar";

    const hostnameUpper      = hostname.toUpperCase();
    const hostnameSemDominio = hostnameUpper.split(".")[0];
    const ondaOrigem = mapaOndas[hostnameUpper]
      || mapaOndas[hostnameSemDominio]
      || mapaOndaSemDominio[hostnameSemDominio]
      || "";
    const ondaDestino = mapaOndas[ipRemoto] || mapaOndas[ipRemoto.toUpperCase()] || "";

    linhas.push({
      data:         c[0],
      hostname,
      proto:        c[2],
      local:        c[3],
      remoto:       c[4],
      estado:       c[5],
      pid:          c[6],
      processo,
      aplicacao,
      ip_local:     c[9],
      porta_local:  c[10],
      ip_remoto:    ipRemoto,
      porta_remota: c[12],
      direcao:      c[13],
      ocvs:         c[14],
      portas_fmt:   c[15],
      contador:     parseInt(contadorRaw, 10) || 1,
      onda_origem:  ondaOrigem,
      onda_destino: ondaDestino,
    });
  }

  cache.set(key, linhas);
  return linhas;
}

// Wrapper sincrono para compatibilidade (usa cache se disponivel)
function lerProcessado(numeroOnda) {
  const key = `processado_${numeroOnda}`;
  return cache.get(key) || null;
}

// Map de promises em andamento para evitar processamento duplo
const _carregando = new Map();

async function lerProcessadoComCache(numeroOnda) {
  const key = `processado_${numeroOnda}`;
  if (cache.has(key)) return cache.get(key);

  // Se ja esta sendo carregado, aguardar a mesma promise
  if (_carregando.has(key)) return _carregando.get(key);

  const promise = _lerProcessadoFonte(numeroOnda).then(dados => {
    _carregando.delete(key);
    return dados;
  });
  _carregando.set(key, promise);
  return promise;
}

// Decide a fonte dos dados: banco (se onda ingerida) ou arquivo .txt (fallback)
async function _lerProcessadoFonte(numeroOnda) {
  try {
    if (db.ondaIngerida(numeroOnda)) {
      console.log(`[cache] Onda ${numeroOnda} — lendo do banco SQLite`);
      const excelPath    = encontrarExcel();
      const mapaAplic    = excelPath ? lerAbaAplicacoes(excelPath) : {};
      const linhas       = db.carregarOndaDoBanco(numeroOnda, mapaAplic);
      const key          = `processado_${numeroOnda}`;
      cache.set(key, linhas);
      return linhas;
    }
  } catch (e) {
    console.warn(`[cache] Falha ao ler banco para onda ${numeroOnda}, usando .txt: ${e.message}`);
  }
  console.log(`[cache] Onda ${numeroOnda} — lendo do arquivo .txt`);
  return lerProcessadoAsync(numeroOnda);
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
  const IPS_DISPENSAVEIS = obterIpsDispensaveis();
  const mapaAmbiente    = obterMapaAmbiente();

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

    if (!arvore.has(h)) arvore.set(h, { hostname: h, ip_local: r.ip_local, ambiente: mapaAmbiente[h.toUpperCase()] || "", ESTABLISHED: 0, SYN_SENT: 0, aplicacoes: new Map() });
    const nH = arvore.get(h);
    nH[est] = (nH[est] || 0) + cnt;

    if (!nH.aplicacoes.has(a)) nH.aplicacoes.set(a, { nome: a, ESTABLISHED: 0, SYN_SENT: 0, ips: new Map() });
    const nA = nH.aplicacoes.get(a);
    nA[est] = (nA[est] || 0) + cnt;

    if (!nA.ips.has(ip)) nA.ips.set(ip, { ip, onda_destino: r.onda_destino || "", ESTABLISHED: 0, SYN_SENT: 0, portas: new Map() });
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
      ip_local:    h.ip_local,
      ambiente:    h.ambiente,
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
              onda_destino: i.onda_destino,
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

const server = http.createServer(async (req, res) => {
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

  // Configurar path do Excel (POST com { path: "C:\\..." })
  if (urlPath === "/api/excel/configurar" && req.method === "POST") {
    let body = "";
    req.on("data", d => body += d);
    req.on("end", () => {
      try {
        const { path: excelPath } = JSON.parse(body);
        if (!excelPath) return respJson(res, { erro: "path obrigatório" }, 400);
        if (!fs.existsSync(excelPath)) return respJson(res, { erro: "Arquivo não encontrado: " + excelPath }, 404);
        if (!excelPath.toLowerCase().endsWith(".xlsx")) return respJson(res, { erro: "Arquivo deve ser .xlsx" }, 400);

        const cfg = lerConfig();
        cfg.excelPath = excelPath;
        salvarConfig(cfg);

        // Limpar caches que dependem do Excel
        cache.delete("excel_path");
        cache.delete("aplicacoes");
        cache.delete("mapa_ondas");
        cache.delete("variaveis");
        cache.delete("mapa_ambiente");

        cache.set("excel_path", excelPath);
        console.log(`[config] Excel configurado: ${excelPath}`);
        return respJson(res, { ok: true, path: excelPath });
      } catch (e) {
        return respJson(res, { erro: e.message }, 500);
      }
    });
    return;
  }

  // Abrir dialog nativo do Windows para selecionar Excel
  if (urlPath === "/api/excel/procurar" && req.method === "POST") {
    const psCmd = [
      "Add-Type -AssemblyName System.Windows.Forms",
      "$f = New-Object System.Windows.Forms.Form",
      "$f.TopMost = $true",
      "$f.WindowState = 'Minimized'",
      "$d = New-Object System.Windows.Forms.OpenFileDialog",
      "$d.Filter = 'Excel (*.xlsx)|*.xlsx'",
      "$d.Title = 'Selecionar planilha OCVS'",
      "$null = $d.ShowDialog($f)",
      "Write-Output $d.FileName",
      "$f.Dispose()",
    ].join("; ");

    const proc = spawn("powershell", ["-NoProfile", "-STA", "-Command", psCmd], {
      windowsHide: false,
      stdio: ["ignore", "pipe", "pipe"],
    });

    let stdout = "";
    proc.stdout.on("data", d => stdout += d.toString());
    proc.stderr.on("data", d => {}); // ignorar stderr
    proc.on("close", () => {
      const result = stdout.trim();
      if (result) {
        return respJson(res, { path: result });
      } else {
        return respJson(res, { path: null, cancelado: true });
      }
    });
    proc.on("error", err => {
      return respJson(res, { erro: err.message }, 500);
    });
    return;
  }

  if (urlPath === "/api/mapa-ambiente") {
    return respJson(res, obterMapaAmbiente());
  }

  if (urlPath === "/api/cache/clear") {
    cache.clear();
    _carregando.clear();
    return respJson(res, { ok: true });
  }

  // Listar IPs de uma onda sem executar coleta
  if (urlPath.match(/^\/api\/onda-servidores\/\w+$/) && req.method === "GET") {
    const numero = urlPath.split("/").pop();
    const excelPath = encontrarExcel();
    if (!excelPath) return respJson(res, { erro: "Excel não encontrado" }, 404);
    try {
      const wb   = getXLSX().readFile(excelPath, { cellText: true, cellDates: false });
      const ws   = wb.Sheets["vInfo"];
      if (!ws) return respJson(res, { erro: "Aba vInfo não encontrada" }, 404);
      const rows = getXLSX().utils.sheet_to_json(ws, { header: 1 });
      const hdr  = rows[0] || [];
      let colVM = -1, colIP = -1, colOnda = -1;
      hdr.forEach((v, i) => {
        const s = String(v || "").trim();
        if (s === "VM")                 colVM   = i;
        else if (/IP|Address/i.test(s)) colIP   = i;
        else if (s === "ONDA")          colOnda = i;
      });
      const servidores = rows.slice(1)
        .filter(r => {
          const onda = String(r[colOnda] || "").trim();
          const ip   = String(r[colIP]   || "").trim();
          return new RegExp(`Onda\\s+${numero}\\b`, "i").test(onda) &&
                 ip && ip !== "None" && ip !== "A definir";
        })
        .map(r => ({
          hostname: String(r[colVM] || "").trim(),
          ip:       String(r[colIP] || "").trim(),
        }));
      return respJson(res, servidores);
    } catch (e) {
      return respJson(res, { erro: e.message }, 500);
    }
  }
  // ── Visão geral — resumo de todas as ondas ──────────────────────────────────
  if (urlPath === "/api/visao-geral") {
    try {
      const excelPath = encontrarExcel();
      const ondasProcessadas = listarOndas();

      // Ler ondas do Excel
      const ondasExcel = {};
      if (excelPath) {
        const wb  = getXLSX().readFile(excelPath, { cellText: true, cellDates: false });
        const ws  = wb.Sheets["vInfo"];
        if (ws) {
          const rows = getXLSX().utils.sheet_to_json(ws, { header: 1 });
          const hdr  = rows[0] || [];
          let colVM = -1, colIP = -1, colOnda = -1;
          hdr.forEach((v, i) => {
            const s = String(v || "").trim();
            if (s === "VM")                 colVM   = i;
            else if (/IP|Address/i.test(s)) colIP   = i;
            else if (s === "ONDA")          colOnda = i;
          });
          for (const row of rows.slice(1)) {
            const vm   = String(row[colVM]   || "").trim();
            const onda = String(row[colOnda] || "").trim();
            if (!vm || vm === "None" || vm === "A definir") continue;
            const m = onda.match(/Onda\s+(\w+)/i);
            if (!m) continue;
            const num = m[1];
            if (!ondasExcel[num]) ondasExcel[num] = { previstos: 0 };
            ondasExcel[num].previstos++;
          }
        }
      }

      // Status do banco
      let dbStatus = { inicializado: false, ondas: [] };
      try { dbStatus = db.statusDB(); } catch {}

      // Último log
      let ultimoLog = null;
      try {
        if (fs.existsSync(LOGS_DIR)) {
          const logs = fs.readdirSync(LOGS_DIR).filter(f => f.endsWith(".json")).sort().reverse();
          if (logs.length > 0) {
            const meta = JSON.parse(fs.readFileSync(path.join(LOGS_DIR, logs[0]), "utf8"));
            ultimoLog = { script: meta.script, onda: meta.onda, inicio: meta.inicio, exitCode: meta.exitCode };
          }
        }
      } catch {}

      // Montar lista unificada de ondas
      const todasOndas = new Set([...Object.keys(ondasExcel), ...ondasProcessadas]);
      const lista = [...todasOndas].sort().map(num => {
        const excel      = ondasExcel[num] || null;
        const processada = ondasProcessadas.includes(num);
        const dbOnda     = (dbStatus.ondas || []).find(o => String(o.numero) === num);
        return {
          numero:     num,
          previstos:  excel ? excel.previstos : 0,
          processada,
          ingerida:   dbOnda ? dbOnda.ingerida : false,
          conexoesBD: dbOnda ? dbOnda.conexoes : 0,
        };
      });

      return respJson(res, {
        totalOndasExcel:      Object.keys(ondasExcel).length,
        totalOndasProcessadas: ondasProcessadas.length,
        totalOndasIngeridas:  lista.filter(o => o.ingerida).length,
        ultimoLog,
        ondas: lista,
      });
    } catch (e) {
      return respJson(res, { erro: e.message }, 500);
    }
  }

  if (urlPath === "/api/logs") {
    try {
      if (!fs.existsSync(LOGS_DIR)) return respJson(res, []);
      const logs = fs.readdirSync(LOGS_DIR)
        .filter(f => f.endsWith(".json"))
        .map(f => {
          try {
            const meta = JSON.parse(fs.readFileSync(path.join(LOGS_DIR, f), "utf8"));
            return { id: f.replace(".json",""), ...meta, linhas: undefined };
          } catch { return null; }
        })
        .filter(Boolean)
        .sort((a, b) => b.inicio.localeCompare(a.inicio));
      return respJson(res, logs);
    } catch { return respJson(res, []); }
  }

  const mLog = urlPath.match(/^\/api\/logs\/(.+)$/);
  if (mLog) {
    const logFile = path.join(LOGS_DIR, mLog[1] + ".json");
    if (!fs.existsSync(logFile)) return respJson(res, { erro: "Log não encontrado" }, 404);
    try {
      return respJson(res, JSON.parse(fs.readFileSync(logFile, "utf8")));
    } catch { return respJson(res, { erro: "Erro ao ler log" }, 500); }
  }

  const mOnda = urlPath.match(/^\/api\/onda\/(\w+)\/(.+)$/);
  if (mOnda) {
    const [, numero, endpoint] = mOnda;

    // Carregar dados (async — usa cache se ja disponivel, senao processa uma unica vez)
    let dados = await lerProcessadoComCache(numero);
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
    if (endpoint === "distribuicao-so") {
      const excelPath = encontrarExcel();
      if (!excelPath) return respJson(res, { erro: "Excel não encontrado" }, 404);
      try {
        const wb  = getXLSX().readFile(excelPath, { cellText: true, cellDates: false });
        const ws  = wb.Sheets["vInfo"];
        if (!ws) return respJson(res, { erro: "Aba vInfo não encontrada" }, 404);
        const rows = getXLSX().utils.sheet_to_json(ws, { header: 1 });
        const hdr  = rows[0] || [];
        let colVM = -1, colOnda = -1, colSO = -1;
        hdr.forEach((v, i) => {
          const s = String(v || "").trim();
          if (s === "VM")                       colVM   = i;
          else if (s === "ONDA")                colOnda = i;
          else if (s === "SO REVISADO RESUMIDO") colSO  = i;
        });
        if (colSO < 0) return respJson(res, { erro: "Coluna 'SO REVISADO RESUMIDO' não encontrada" }, 404);

        const contagem = {};
        for (const row of rows.slice(1)) {
          const vm   = String(row[colVM]   || "").trim();
          const onda = String(row[colOnda] || "").trim();
          const so   = String(row[colSO]   || "").trim() || "Desconhecido";
          if (!vm || vm === "None" || vm === "A definir") continue;
          if (!new RegExp(`Onda\\s+${numero}\\b`, "i").test(onda)) continue;
          contagem[so] = (contagem[so] || 0) + 1;
        }

        // Agrupar por família (Windows/Linux/Outro) e manter detalhe
        const familias = {};
        for (const [so, qtd] of Object.entries(contagem)) {
          let familia = "Outro";
          if (/windows/i.test(so)) familia = "Windows";
          else if (/linux|ubuntu|debian|centos|rhel|red\s*hat|suse|oracle\s*linux|alma|rocky/i.test(so)) familia = "Linux";
          if (!familias[familia]) familias[familia] = { total: 0, versoes: {} };
          familias[familia].total += qtd;
          familias[familia].versoes[so] = qtd;
        }

        return respJson(res, familias);
      } catch (e) {
        return respJson(res, { erro: e.message }, 500);
      }
    }
    // Drilldown de SO: /api/onda/90/so-drilldown/Windows%20Server%202012
    const mSO = endpoint.match(/^so-drilldown\/(.+)$/);
    if (mSO) {
      const versaoBuscada = decodeURIComponent(mSO[1]);
      const excelPath = encontrarExcel();
      if (!excelPath) return respJson(res, { erro: "Excel não encontrado" }, 404);
      try {
        const wb  = getXLSX().readFile(excelPath, { cellText: true, cellDates: false });
        const ws  = wb.Sheets["vInfo"];
        if (!ws) return respJson(res, { erro: "Aba vInfo não encontrada" }, 404);
        const rows = getXLSX().utils.sheet_to_json(ws, { header: 1 });
        const hdr  = rows[0] || [];
        let colVM = -1, colIP = -1, colOnda = -1, colSO = -1, colAmb = -1;
        hdr.forEach((v, i) => {
          const s = String(v || "").trim();
          if (s === "VM")                                              colVM   = i;
          else if (/IP|Address/i.test(s))                             colIP   = i;
          else if (s === "ONDA")                                      colOnda = i;
          else if (s === "SO REVISADO RESUMIDO")                      colSO   = i;
          else if (s === "PROD/NÃO-PROD" || s === "PROD/NAO-PROD")   colAmb  = i;
        });

        const servidores = [];
        for (const row of rows.slice(1)) {
          const vm   = String(row[colVM]   || "").trim();
          const onda = String(row[colOnda] || "").trim();
          const so   = String(row[colSO]   || "").trim() || "Desconhecido";
          if (!vm || vm === "None" || vm === "A definir") continue;
          if (!new RegExp(`Onda\\s+${numero}\\b`, "i").test(onda)) continue;
          if (so !== versaoBuscada) continue;
          servidores.push({
            hostname: vm,
            ip:       String(row[colIP]  || "").trim(),
            ambiente: colAmb >= 0 ? String(row[colAmb] || "").trim() : "",
            so,
          });
        }
        return respJson(res, servidores);
      } catch (e) {
        return respJson(res, { erro: e.message }, 500);
      }
    }
    if (endpoint === "resumo-geral") {
      // Ignora filtro de tipo de conexao — usa dados brutos completos
      const dadosBrutos = lerProcessado(numero);
      if (!dadosBrutos) return respJson(res, { erro: "Onda não encontrada" }, 404);

      const IPS_DISPENSAVEIS = obterIpsDispensaveis();

      // Verificar se IP é privado (RFC 1918 + loopback + link-local)
      function isPrivado(ip) {
        if (!ip) return false;
        const p = ip.split(".").map(Number);
        if (p[0] === 10) return true;
        if (p[0] === 172 && p[1] >= 16 && p[1] <= 31) return true;
        if (p[0] === 192 && p[1] === 168) return true;
        if (p[0] === 127) return true;
        if (p[0] === 169 && p[1] === 254) return true;
        return false;
      }

      let externoPrivado = 0, externoPublico = 0;
      let ondaAnterior = 0, mesmaOnda = 0, problema = 0;

      for (const r of dadosBrutos) {
        const c  = r.contador;
        const od = r.onda_destino || "";
        const oo = r.onda_origem  || "";

        if (r.ocvs !== "OCVS") {
          if (isPrivado(r.ip_remoto)) externoPrivado += c;
          else                        externoPublico  += c;
          continue;
        }

        const mDest = od.match(/Onda\s+(\d+)/i);
        const mOrig = oo.match(/Onda\s+(\d+)/i);
        const nDest = mDest ? parseInt(mDest[1]) : null;
        const nOrig = mOrig ? parseInt(mOrig[1]) : null;

        if (nDest !== null && nOrig !== null && nDest === nOrig) { mesmaOnda += c; continue; }
        if (nDest !== null && nOrig !== null && nDest < nOrig)   { ondaAnterior += c; continue; }

        if ((r.estado === "ESTABLISHED" || r.estado === "SYN_SENT") &&
            !IPS_DISPENSAVEIS.has(r.ip_remoto) &&
            (!od || !/Onda\s+\d+/i.test(od) || od === "A definir")) {
          problema += c;
        }
      }

      return respJson(res, { externoPrivado, externoPublico, ondaAnterior, mesmaOnda, problema });
    }

    if (endpoint === "resumo") {
      // servidores processados: sempre sem filtro de ocvs para refletir o total real
      const dadosBrutos = lerProcessado(numero) || dados;
      const servidoresProcessados = [...new Set(dadosBrutos.map(r => r.hostname.toUpperCase()))].length;
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

  // ── Status do banco ──────────────────────────────────────────────────────────
  if (urlPath === "/api/db/status") {
    return respJson(res, db.statusDB());
  }

  // ── Ingerir onda no banco (POST) ──────────────────────────────────────────────
  // Chamado automaticamente após processar uma onda
  if (urlPath === "/api/db/ingerir" && req.method === "POST") {
    let body = "";
    req.on("data", d => body += d);
    req.on("end", async () => {
      try {
        const { onda } = JSON.parse(body);
        if (!onda) return respJson(res, { erro: "onda obrigatória" }, 400);

        // Garantir que o banco está inicializado
        await db.initDB();

        // Carregar dados processados (usa cache se disponível)
        const linhas = await lerProcessadoComCache(onda);
        if (!linhas) return respJson(res, { erro: "Onda não encontrada" }, 404);

        // Sincronizar aplicações
        const excelPath = encontrarExcel();
        if (excelPath) {
          const aplicacoes = lerAbaAplicacoes(excelPath);
          db.sincronizarAplicacoes(aplicacoes);
        }

        const resultado = db.ingerirOnda(onda, linhas);
        console.log(`[db] Onda ${onda} ingerida: ${resultado.servidores} servidores, ${resultado.linhas} linhas`);
        return respJson(res, resultado);
      } catch (e) {
        console.error("[db] Erro na ingestão:", e.message);
        return respJson(res, { erro: e.message }, 500);
      }
    });
    return;
  }

  // ── Reingerir todas as ondas (POST) ──────────────────────────────────────────  // Usado para reprocessar o banco após mudanças no schema de ingestão
  if (urlPath === "/api/db/reingerir-tudo" && req.method === "POST") {
    (async () => {
      try {
        await db.initDB();
        const ondas = listarOndas();
        if (ondas.length === 0) return respJson(res, { ok: true, ondas: 0, msg: "Nenhuma onda processada encontrada" });

        const excelPath = encontrarExcel();
        const resultados = [];

        for (const onda of ondas) {
          try {
            // Forçar leitura do .txt (ignorar cache do banco)
            const linhas = await lerProcessadoAsync(onda);
            if (!linhas) { resultados.push({ onda, ok: false, erro: "Arquivo não encontrado" }); continue; }
            if (excelPath) db.sincronizarAplicacoes(lerAbaAplicacoes(excelPath));
            const r = db.ingerirOnda(onda, linhas);
            // Limpar cache para próxima leitura vir do banco atualizado
            cache.delete(`processado_${onda}`);
            _carregando.delete(`processado_${onda}`);
            resultados.push({ onda, ok: true, servidores: r.servidores, linhas: r.linhas });
            console.log(`[reingerir] Onda ${onda}: ${r.servidores} servidores, ${r.linhas} linhas`);
          } catch (e) {
            resultados.push({ onda, ok: false, erro: e.message });
          }
        }
        return respJson(res, { ok: true, resultados });
      } catch (e) {
        return respJson(res, { erro: e.message }, 500);
      }
    })();
    return;
  }

  // ── Sincronizar membros de uma onda (POST) ────────────────────────────────────  // Atualiza onda_membros sem reprocessar — usado quando composição da onda muda
  if (urlPath === "/api/db/sincronizar-onda" && req.method === "POST") {
    let body = "";
    req.on("data", d => body += d);
    req.on("end", async () => {
      try {
        const { onda } = JSON.parse(body);
        if (!onda) return respJson(res, { erro: "onda obrigatória" }, 400);

        await db.initDB();

        const excelPath = encontrarExcel();
        if (!excelPath) return respJson(res, { erro: "Excel não encontrado" }, 404);

        // Ler membros da onda do Excel (hostname + ip + ambiente)
        const wb   = getXLSX().readFile(excelPath, { cellText: true, cellDates: false });
        const ws   = wb.Sheets["vInfo"];
        if (!ws) return respJson(res, { erro: "Aba vInfo não encontrada" }, 404);

        const rows = getXLSX().utils.sheet_to_json(ws, { header: 1 });
        const hdr  = rows[0] || [];
        let colVM = -1, colIP = -1, colOnda = -1, colAmbiente = -1;
        hdr.forEach((v, i) => {
          const s = String(v || "").trim();
          if (s === "VM")                 colVM       = i;
          else if (/IP|Address/i.test(s)) colIP       = i;
          else if (s === "ONDA")          colOnda     = i;
          else if (s === "PROD/NÃO-PROD" || s === "PROD/NAO-PROD") colAmbiente = i;
        });

        const membros = rows.slice(1)
          .filter(r => {
            const ondaVal = String(r[colOnda] || "").trim();
            const vm      = String(r[colVM]   || "").trim();
            return new RegExp(`Onda\\s+${onda}\\b`, "i").test(ondaVal) &&
                   vm && vm !== "None" && vm !== "A definir";
          })
          .map(r => ({
            hostname: String(r[colVM]       || "").trim(),
            ip:       String(r[colIP]       || "").trim(),
            ambiente: String(r[colAmbiente] || "").trim(),
          }));

        const resultado = db.sincronizarOnda(onda, membros);
        // Limpar cache para forçar releitura do banco com nova composição
        cache.delete(`processado_${onda}`);
        _carregando.delete(`processado_${onda}`);
        console.log(`[db] Onda ${onda} sincronizada: ${resultado.membros} membros`);
        return respJson(res, resultado);
      } catch (e) {
        console.error("[db] Erro na sincronização:", e.message);
        return respJson(res, { erro: e.message }, 500);
      }
    });
    return;
  }

  // ── Sincronizar TODAS as ondas do Excel de uma vez (POST) ───────────────────
  if (urlPath === "/api/db/sincronizar-tudo" && req.method === "POST") {
    (async () => {
      try {
        await db.initDB();
        const excelPath = encontrarExcel();
        if (!excelPath) return respJson(res, { erro: "Excel não encontrado" }, 404);

        const wb   = getXLSX().readFile(excelPath, { cellText: true, cellDates: false });
        const ws   = wb.Sheets["vInfo"];
        if (!ws) return respJson(res, { erro: "Aba vInfo não encontrada" }, 404);

        const rows = getXLSX().utils.sheet_to_json(ws, { header: 1 });
        const hdr  = rows[0] || [];
        let colVM = -1, colIP = -1, colOnda = -1, colAmbiente = -1;
        hdr.forEach((v, i) => {
          const s = String(v || "").trim();
          if (s === "VM")                                              colVM       = i;
          else if (/IP|Address/i.test(s))                             colIP       = i;
          else if (s === "ONDA")                                      colOnda     = i;
          else if (s === "PROD/NÃO-PROD" || s === "PROD/NAO-PROD")   colAmbiente = i;
        });

        // Agrupar servidores por onda
        const porOnda = {};
        for (const row of rows.slice(1)) {
          const vm   = String(row[colVM]   || "").trim();
          const ip   = String(row[colIP]   || "").trim();
          const onda = String(row[colOnda] || "").trim();
          const amb  = String(row[colAmbiente] || "").trim();
          if (!vm || vm === "None" || vm === "A definir") continue;
          // Extrair número da onda (ex: "Onda 2" → "2")
          const mOnda = onda.match(/Onda\s+(\w+)/i);
          if (!mOnda) continue;
          const num = mOnda[1];
          if (!porOnda[num]) porOnda[num] = [];
          porOnda[num].push({ hostname: vm, ip, ambiente: amb });
        }

        const resultados = [];
        for (const [num, membros] of Object.entries(porOnda)) {
          try {
            const r = db.sincronizarOnda(num, membros);
            // Limpar cache para forçar releitura com nova composição
            cache.delete(`processado_${num}`);
            _carregando.delete(`processado_${num}`);
            resultados.push({ onda: num, ok: true, membros: r.membros });
            console.log(`[sincronizar-tudo] Onda ${num}: ${r.membros} membros`);
          } catch (e) {
            resultados.push({ onda: num, ok: false, erro: e.message });
          }
        }

        return respJson(res, { ok: true, resultados });
      } catch (e) {
        return respJson(res, { erro: e.message }, 500);
      }
    })();
    return;
  }

  // Execucao de scripts PowerShell com streaming via SSE
  if (urlPath === "/api/executar" && req.method === "POST") {
    let body = "";
    req.on("data", d => body += d);
    req.on("end", () => {
      let params;
      try { params = JSON.parse(body); } catch { res.writeHead(400); res.end("JSON invalido"); return; }

      const { script, onda, usuario, senha, servidoresSelecionados } = params;
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
      const selArg     = (script === "coletar" && servidoresSelecionados) ? " -ServidoresSelecionados '" + servidoresSelecionados + "'" : "";

      const PWSH_REAL = process.env.OCVS_PWSH && fs.existsSync(process.env.OCVS_PWSH)
        ? process.env.OCVS_PWSH
        : (() => { try { return require("child_process").execSync("where pwsh", {encoding:"utf8"}).trim().split(/\r?\n/)[0]; } catch {} return "powershell.exe"; })();

      const tmpScript = path.join(require("os").tmpdir(), "ocvs_" + Date.now() + ".ps1");
      fs.writeFileSync(tmpScript,
        "& '" + scriptPath.replace(/'/g, "''") + "' -NumeroOnda '" + onda + "'" + usuarioArg + senhaArg + selArg + "\n",
        "utf8");

      sendLine("Iniciando " + scriptFile + " para Onda " + onda + "...", "info");

      // Estrutura do log
      if (!fs.existsSync(LOGS_DIR)) fs.mkdirSync(LOGS_DIR, { recursive: true });
      const logId   = new Date().toISOString().replace(/[:.]/g,"-").slice(0,19);
      const logFile = path.join(LOGS_DIR, logId + ".json");
      const logData = {
        id:      logId,
        script:  scriptFile,
        onda,
        usuario: usuario || "",
        inicio:  new Date().toISOString(),
        fim:     null,
        exitCode: null,
        linhas:  [],
      };
      const appendLog = (linha, tipo) => {
        logData.linhas.push({ tipo: tipo || "log", linha });
        fs.writeFileSync(logFile, JSON.stringify(logData), "utf8");
      };
      appendLog("Iniciando " + scriptFile + " para Onda " + onda + "...", "info");

      // Usar detached:true para desacoplar do processo pai (Node filho de PS7)
      const proc = spawn(PWSH_REAL,
        ["-NoProfile", "-ExecutionPolicy", "Bypass", "-File", tmpScript],
        { cwd: scriptsCwd, windowsHide: true, detached: false, stdio: ["ignore", "pipe", "pipe"] }
      );

      proc.stdout.on("data", d => d.toString().split(/\r?\n/).filter(Boolean).forEach(l => { sendLine(l); appendLog(l); }));
      proc.stderr.on("data", d => d.toString().split(/\r?\n/).filter(Boolean).forEach(l => { sendLine(l, "erro"); appendLog(l, "erro"); }));
      proc.on("error", err => { sendLine("Erro: " + err.message, "erro"); appendLog("Erro: " + err.message, "erro"); });
      proc.on("close", code => {
        try { fs.unlinkSync(tmpScript); } catch {}
        logData.fim      = new Date().toISOString();
        logData.exitCode = code;
        const msg = "Processo encerrado (exit " + code + ")";
        appendLog(msg, code === 0 ? "sucesso" : "erro");
        sendLine(msg, code === 0 ? "sucesso" : "erro");

        // Se foi um processamento bem-sucedido, ingerir no banco automaticamente
        if (code === 0 && script === "processar") {
          sendLine("Ingerindo dados no banco...", "info");
          db.initDB().then(() => {
            return lerProcessadoAsync(onda); // sempre lê do .txt na ingestão
          }).then(linhas => {
            if (linhas) {
              const excelPath = encontrarExcel();
              if (excelPath) db.sincronizarAplicacoes(lerAbaAplicacoes(excelPath));
              const r = db.ingerirOnda(onda, linhas);
              sendLine(`Banco atualizado: ${r.servidores} servidores, ${r.linhas} linhas`, "sucesso");
              appendLog(`Banco atualizado: ${r.servidores} servidores, ${r.linhas} linhas`, "sucesso");
            }
          }).catch(e => {
            sendLine(`Aviso: falha na ingestão do banco — ${e.message}`, "warn");
          }).finally(() => {
            try { res.write("data: {\"tipo\":\"fim\"}\n\n"); res.end(); } catch {}
            // Limpar cache para próxima leitura vir do banco
            cache.delete("processado_" + onda);
            _carregando.delete("processado_" + onda);
          });
          return; // não fechar o SSE ainda — aguardar ingestão
        }

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
server.listen(PORT, "127.0.0.1", async () => {
  // Inicializar banco SQLite
  try {
    await db.initDB();
    const status = db.statusDB();
    console.log(` DB:    ${status.tamanhoMB}MB — ${status.servidores} servidores, ${status.conexoes} conexões`);
  } catch (e) {
    console.error(" DB:    Erro ao inicializar:", e.message);
  }
  const excelPath = encontrarExcel();
  console.log(`\n========================================`);
  console.log(` OCVS Migration Dashboard v0.3.5`);
  console.log(`========================================`);
  console.log(` URL:   http://localhost:${PORT}`);
  console.log(` Base:  ${BASE_DIR}`);
  console.log(` Excel: ${excelPath || "NÃO ENCONTRADO"}`);
  console.log(` Ondas: ${listarOndas().join(", ") || "nenhuma"}`);
  console.log(`========================================\n`);
  // Abrir browser automaticamente
  try { execSync(`start http://localhost:${PORT}`); } catch {}
});
