/**
 * OCVS Migration — Camada de banco SQLite (sql.js / WebAssembly)
 * Sem dependências nativas — portável em qualquer Windows sem build tools
 */

const fs   = require("fs");
const path = require("path");

const DB_PATH = path.join(__dirname, "..", "dados", "ocvs.db");

let _db   = null;  // instância sql.js
let _SQL  = null;  // módulo sql.js

// ── Inicializar banco ─────────────────────────────────────────────────────────
async function initDB() {
  if (_db) return _db;

  const initSqlJs = require("sql.js");
  // Apontar para o wasm incluído no pacote
  const wasmPath = path.join(__dirname, "node_modules", "sql.js", "dist", "sql-wasm.wasm");
  _SQL = await initSqlJs({ wasmBinary: fs.readFileSync(wasmPath) });

  // Carregar banco existente ou criar novo
  if (fs.existsSync(DB_PATH)) {
    const buf = fs.readFileSync(DB_PATH);
    _db = new _SQL.Database(buf);
  } else {
    _db = new _SQL.Database();
  }

  criarSchema();
  return _db;
}

// ── Persistir banco em disco ──────────────────────────────────────────────────
function salvarDB() {
  if (!_db) return;
  const dir = path.dirname(DB_PATH);
  if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
  const data = _db.export();
  const buf  = Buffer.from(data);

  // Retry com delay para lidar com lock do OneDrive
  const MAX_TENTATIVAS = 5;
  for (let i = 0; i < MAX_TENTATIVAS; i++) {
    try {
      fs.writeFileSync(DB_PATH, buf);
      return;
    } catch (e) {
      if (i < MAX_TENTATIVAS - 1 && (e.code === "EBUSY" || e.code === "EPERM")) {
        // Esperar antes de tentar novamente
        const delay = (i + 1) * 1000;
        const start = Date.now();
        while (Date.now() - start < delay) {} // busy wait (sync)
      } else {
        throw e;
      }
    }
  }
}

// ── Schema ────────────────────────────────────────────────────────────────────
function criarSchema() {
  _db.run(`
    CREATE TABLE IF NOT EXISTS servidores (
      hostname    TEXT PRIMARY KEY,
      ip          TEXT,
      ambiente    TEXT,   -- PROD / NÃO-PROD
      ingerido_em TEXT    -- ISO timestamp da última ingestão
    );

    CREATE TABLE IF NOT EXISTS ondas (
      numero TEXT PRIMARY KEY,
      nome   TEXT
    );

    CREATE TABLE IF NOT EXISTS onda_membros (
      onda_numero TEXT,
      hostname    TEXT,
      PRIMARY KEY (onda_numero, hostname)
    );

    CREATE TABLE IF NOT EXISTS aplicacoes (
      executavel TEXT PRIMARY KEY,
      nome       TEXT
    );

    CREATE TABLE IF NOT EXISTS conexoes (
      id           INTEGER PRIMARY KEY AUTOINCREMENT,
      hostname     TEXT,    -- servidor de origem
      ip_local     TEXT,
      porta_local  TEXT,
      ip_remoto    TEXT,
      porta_remota TEXT,
      direcao      TEXT,    -- IN / OUT / ANALISAR
      ocvs         TEXT,    -- OCVS / FORA DO OCVS
      estado       TEXT,    -- ESTABLISHED / SYN_SENT / etc
      protocolo    TEXT,
      processo     TEXT,
      aplicacao    TEXT,
      portas_fmt   TEXT,
      contador     INTEGER,
      data_coleta  TEXT     -- data/hora da linha original
    );

    CREATE INDEX IF NOT EXISTS idx_conexoes_hostname  ON conexoes(hostname);
    CREATE INDEX IF NOT EXISTS idx_conexoes_ip_remoto ON conexoes(ip_remoto);
    CREATE INDEX IF NOT EXISTS idx_conexoes_ocvs      ON conexoes(ocvs);
    CREATE INDEX IF NOT EXISTS idx_conexoes_estado    ON conexoes(estado);
    CREATE INDEX IF NOT EXISTS idx_onda_membros_onda  ON onda_membros(onda_numero);
  `);
}

// ── Ingerir arquivo processado de uma onda ────────────────────────────────────
// Recebe o array de linhas já enriquecidas (mesmo formato do lerProcessado)
function ingerirOnda(numeroOnda, linhas) {
  if (!_db) throw new Error("Banco não inicializado");

  // Coletar hostnames únicos desta onda com seus IPs locais
  const hostnameIpMap = {};
  for (const r of linhas) {
    if (r.hostname && r.ip_local && !hostnameIpMap[r.hostname]) {
      hostnameIpMap[r.hostname] = r.ip_local;
    }
  }
  const hostnames = Object.keys(hostnameIpMap);

  // Coletar todos os IPs remotos únicos que aparecem nas conexões
  // (para que o subquery de onda_destino consiga resolver via servidores.ip)
  const ipsRemotos = [...new Set(linhas.map(r => r.ip_remoto).filter(Boolean))];

  _db.run("BEGIN TRANSACTION");
  try {
    // Remover conexões antigas desses servidores
    for (const h of hostnames) {
      _db.run("DELETE FROM conexoes WHERE hostname = ?", [h]);
    }

    // Registrar onda
    _db.run(
      "INSERT OR IGNORE INTO ondas (numero, nome) VALUES (?, ?)",
      [numeroOnda, `Onda ${numeroOnda}`]
    );

    // Inserir/atualizar servidores de ORIGEM com IP correto
    const agora = new Date().toISOString();
    for (const h of hostnames) {
      const ip = hostnameIpMap[h] || "";
      _db.run(`
        INSERT INTO servidores (hostname, ip, ambiente, ingerido_em)
        VALUES (?, ?, ?, ?)
        ON CONFLICT(hostname) DO UPDATE SET
          ip          = CASE WHEN excluded.ip != '' THEN excluded.ip ELSE ip END,
          ingerido_em = excluded.ingerido_em
      `, [h.toUpperCase(), ip, "", agora]);

      // Registrar como membro desta onda
      _db.run(
        "INSERT OR IGNORE INTO onda_membros (onda_numero, hostname) VALUES (?, ?)",
        [numeroOnda, h.toUpperCase()]
      );
    }

    // Inserir IPs remotos na tabela servidores (sem hostname, só IP)
    // Isso permite que o subquery de onda_destino resolva via servidores.ip
    // Só insere se ainda não existe — não sobrescreve registros com hostname real
    for (const ip of ipsRemotos) {
      _db.run(`
        INSERT OR IGNORE INTO servidores (hostname, ip, ambiente, ingerido_em)
        VALUES (?, ?, '', '')
      `, [ip, ip]); // hostname = ip como placeholder quando não temos o hostname
    }

    // Inserir conexões
    const stmt = _db.prepare(`
      INSERT INTO conexoes
        (hostname, ip_local, porta_local, ip_remoto, porta_remota,
         direcao, ocvs, estado, protocolo, processo, aplicacao,
         portas_fmt, contador, data_coleta)
      VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)
    `);

    for (const r of linhas) {
      stmt.run([
        r.hostname.toUpperCase(), r.ip_local, r.porta_local, r.ip_remoto, r.porta_remota,
        r.direcao, r.ocvs, r.estado, r.proto, r.processo, r.aplicacao,
        r.portas_fmt, r.contador, r.data,
      ]);
    }
    stmt.free();

    _db.run("COMMIT");
    salvarDB();
    return { ok: true, servidores: hostnames.length, linhas: linhas.length };
  } catch (e) {
    try { _db.run("ROLLBACK"); } catch {}
    throw e;
  }
}

// ── Sincronizar membros de uma onda (sem reprocessar) ─────────────────────────
// Atualiza onda_membros e servidores.ambiente a partir da planilha
function sincronizarOnda(numeroOnda, membros) {
  // membros: array de { hostname, ip, ambiente }
  if (!_db) throw new Error("Banco não inicializado");

  _db.run("BEGIN TRANSACTION");
  try {
    // Limpar membros atuais da onda
    _db.run("DELETE FROM onda_membros WHERE onda_numero = ?", [numeroOnda]);

    // Registrar onda
    _db.run(
      "INSERT OR IGNORE INTO ondas (numero, nome) VALUES (?, ?)",
      [numeroOnda, `Onda ${numeroOnda}`]
    );

    for (const m of membros) {
      // Inserir/atualizar servidor com ambiente
      _db.run(`
        INSERT INTO servidores (hostname, ip, ambiente, ingerido_em)
        VALUES (?, ?, ?, '')
        ON CONFLICT(hostname) DO UPDATE SET
          ip = excluded.ip,
          ambiente = excluded.ambiente
      `, [m.hostname.toUpperCase(), m.ip || "", m.ambiente || ""]);

      // Adicionar à onda
      _db.run(
        "INSERT OR IGNORE INTO onda_membros (onda_numero, hostname) VALUES (?, ?)",
        [numeroOnda, m.hostname.toUpperCase()]
      );
    }

    _db.run("COMMIT");
    salvarDB();
    return { ok: true, membros: membros.length };
  } catch (e) {
    try { _db.run("ROLLBACK"); } catch {}
    throw e;
  }
}

// ── Sincronizar tabela de aplicações ─────────────────────────────────────────
function sincronizarAplicacoes(mapa) {
  // mapa: { executavel: nome }
  if (!_db) throw new Error("Banco não inicializado");
  _db.run("BEGIN TRANSACTION");
  try {
    for (const [exec, nome] of Object.entries(mapa)) {
      _db.run(
        "INSERT OR REPLACE INTO aplicacoes (executavel, nome) VALUES (?, ?)",
        [exec, nome]
      );
    }
    _db.run("COMMIT");
    salvarDB();
  } catch (e) {
    try { _db.run("ROLLBACK"); } catch {}
    throw e;
  }
}

// ── Queries de análise ────────────────────────────────────────────────────────
function queryAll(sql, params = []) {
  if (!_db) throw new Error("Banco não inicializado");
  const stmt = _db.prepare(sql);
  stmt.bind(params);
  const rows = [];
  while (stmt.step()) rows.push(stmt.getAsObject());
  stmt.free();
  return rows;
}

function queryOne(sql, params = []) {
  const rows = queryAll(sql, params);
  return rows[0] || null;
}

function exec(sql, params = []) {
  if (!_db) throw new Error("Banco não inicializado");
  _db.run(sql, params);
}

// ── Verificar se uma onda está ingerida no banco ──────────────────────────────
function ondaIngerida(numeroOnda) {
  if (!_db) return false;
  const r = queryOne(
    "SELECT COUNT(*) as n FROM conexoes c JOIN servidores s ON s.hostname = c.hostname JOIN onda_membros om ON om.hostname = s.hostname WHERE om.onda_numero = ?",
    [String(numeroOnda)]
  );
  return r && r.n > 0;
}

// ── Carregar linhas de uma onda do banco (com onda_destino via JOIN) ──────────
// Retorna array no mesmo formato que lerProcessadoAsync para compatibilidade total
function carregarOndaDoBanco(numeroOnda, mapaAplicacoes) {
  if (!_db) throw new Error("Banco não inicializado");

  // Buscar membros desta onda para filtrar conexões de origem
  const membros = queryAll(
    "SELECT hostname FROM onda_membros WHERE onda_numero = ?",
    [String(numeroOnda)]
  );
  if (membros.length === 0) return [];

  const hostnamesOnda = membros.map(m => m.hostname);

  // Buscar onda de origem de cada hostname desta onda em uma query só
  const placeholdersM = hostnamesOnda.map(() => "?").join(",");
  const ondaOrigemRows = queryAll(
    `SELECT hostname, onda_numero FROM onda_membros WHERE hostname IN (${placeholdersM})`,
    hostnamesOnda
  );
  const ondaOrigemMap = {};
  for (const r of ondaOrigemRows) {
    ondaOrigemMap[r.hostname.toUpperCase()] = r.onda_numero;
  }

  // Placeholders para IN clause
  const placeholders = hostnamesOnda.map(() => "?").join(",");

  // Buscar conexões com onda_destino resolvida via JOIN em tempo de consulta
  // Usa MIN(onda_numero) para evitar duplicatas quando IP está em múltiplas ondas
  const rows = queryAll(`
    SELECT
      c.id, c.hostname, c.ip_local, c.porta_local, c.ip_remoto, c.porta_remota,
      c.direcao, c.ocvs, c.estado, c.protocolo, c.processo, c.aplicacao,
      c.portas_fmt, c.contador, c.data_coleta,
      (
        SELECT MIN(om2.onda_numero)
        FROM servidores s2
        JOIN onda_membros om2 ON om2.hostname = s2.hostname
        WHERE s2.ip = c.ip_remoto
           OR s2.hostname = c.ip_remoto
      ) AS onda_destino
    FROM conexoes c
    WHERE c.hostname IN (${placeholders})
  `, hostnamesOnda);

  // Montar objetos no formato padrão
  return rows.map(r => {
    const hostnameUpper = (r.hostname || "").toUpperCase();
    const ondaOrigemRaw = ondaOrigemMap[hostnameUpper] || "";
    const ondaDestinoRaw = r.onda_destino || "";

    // Normalizar para o formato "Onda N" — igual ao que vem do Excel via lerMapaOndas
    const normalizarOnda = (v) => {
      if (!v) return "";
      // Se já tem o formato "Onda N", retorna como está
      if (/Onda\s+\d+/i.test(v)) return v;
      // Se é só número, formata
      if (/^\d+$/.test(String(v).trim())) return `Onda ${v}`;
      return v;
    };

    const ondaOrigem  = normalizarOnda(ondaOrigemRaw);
    const ondaDestino = normalizarOnda(ondaDestinoRaw);

    // Aplicação: usar o que está no banco, ou fazer lookup pelo processo
    let aplicacao = r.aplicacao || "";
    if ((!aplicacao || aplicacao === "Falta Identificar") && mapaAplicacoes) {
      const procKey = (r.processo || "").toLowerCase();
      aplicacao = mapaAplicacoes[procKey] || "Falta Identificar";
    }

    return {
      data:         r.data_coleta  || "",
      hostname:     r.hostname     || "",
      proto:        r.protocolo    || "",
      local:        r.ip_local ? `${r.ip_local}:${r.porta_local}` : "",
      remoto:       r.ip_remoto ? `${r.ip_remoto}:${r.porta_remota}` : "",
      estado:       r.estado       || "",
      pid:          "",
      processo:     r.processo     || "",
      aplicacao,
      ip_local:     r.ip_local     || "",
      porta_local:  r.porta_local  || "",
      ip_remoto:    r.ip_remoto    || "",
      porta_remota: r.porta_remota || "",
      direcao:      r.direcao      || "",
      ocvs:         r.ocvs         || "",
      portas_fmt:   r.portas_fmt   || "",
      contador:     r.contador     || 1,
      onda_origem:  ondaOrigem,
      onda_destino: ondaDestino,
    };
  });
}

// ── Status do banco ───────────────────────────────────────────────────────────
function statusDB() {
  if (!_db) return { inicializado: false };
  const servidores  = queryOne("SELECT COUNT(*) as n FROM servidores").n;
  const conexoes    = queryOne("SELECT COUNT(*) as n FROM conexoes").n;
  const ondas       = queryAll("SELECT numero, nome FROM ondas ORDER BY numero");
  const tamanhoMB   = fs.existsSync(DB_PATH)
    ? Math.round(fs.statSync(DB_PATH).size / 1024 / 1024 * 10) / 10
    : 0;

  // Para cada onda, verificar se tem conexões ingeridas
  const ondasComStatus = ondas.map(o => {
    const r = queryOne(
      "SELECT COUNT(*) as n FROM conexoes c JOIN onda_membros om ON om.hostname = c.hostname WHERE om.onda_numero = ?",
      [o.numero]
    );
    return { ...o, ingerida: r && r.n > 0, conexoes: r ? r.n : 0 };
  });

  return { inicializado: true, servidores, conexoes, ondas: ondasComStatus, tamanhoMB };
}

module.exports = {
  initDB, salvarDB,
  ingerirOnda, sincronizarOnda, sincronizarAplicacoes,
  ondaIngerida, carregarOndaDoBanco,
  queryAll, queryOne, exec,
  statusDB,
};
