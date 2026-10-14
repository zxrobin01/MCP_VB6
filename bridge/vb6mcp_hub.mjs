#!/usr/bin/env node
// bridge/vb6mcp_hub.mjs - standalone stdio MCP hub for VB6MCP.
//
// Zero-dependency Node MCP server.  It is the single DSH-facing entry point and it does
// NOT need any VB6 IDE to be running:
//   * tool face  = bridge/hub_tools.json (hub-only ide_* tools)
//                + add-in face: live tools/list when an IDE is up, else the GENERATED
//                  snapshot bridge/hub_schema.generated.json (tools/gen_hub_snapshot.py).
//   * tools/call = hub tools handled locally; add-in tools forwarded to the target IDE.
//
// Strictness (contract v1.3.0 invariants):
//   * tools/list  result keys are EXACTLY {tools}
//   * tools/call  result keys are EXACTLY {content,isError}
//   * `ide:{pid,project,port}` and the version stamp `mcp:{product,contract}` live in the
//     RESPONSE BODY (result.content[].text), never as top-level result keys.
//   * initialize carries `instructions` (the add-in's when an IDE is online, else a
//     snapshot self-description).
import fs from 'node:fs';
import net from 'node:net';
import os from 'node:os';
import path from 'node:path';
import { execFile, execFileSync, spawn } from 'node:child_process';
import { fileURLToPath } from 'node:url';

const HERE = path.dirname(fileURLToPath(import.meta.url));
const SNAP = path.join(HERE, 'hub_schema.generated.json');
const HUBCFG = path.join(HERE, 'hub_tools.json');
const PROTOCOLS = ['2025-11-25', '2025-06-18', '2025-03-26', '2024-11-05', '2024-10-07'];
const LATEST = PROTOCOLS[0];
const SERVER = { name: 'vb6mcp-hub', version: '0.1.0' };
// Latency budget: the DSH client must get initialize/tools/list in <1s.
const DISCOVERY_TTL_MS = 15000;     // cached (re-validated) discovery is reused for this long
const DISCOVERY_BUDGET_MS = 500;    // hard cap on one discovery pass
const PROBE_TIMEOUT_MS = 250;       // per-RPC cap for the fast scan path
// Forwarded calls may legitimately be LONG (a whole-module get_component_code
// took 22.8s directly).  The budget is configurable and a timeout is reported explicitly.
const FORWARD_TIMEOUT_MS = parseInt(process.env.VB6MCP_FORWARD_TIMEOUT_MS || '120000', 10);
const IDLE_FLUSH_MS = 600;          // a complete-but-unterminated request is flushed after this
const DISCOVERY_GRACE_MS = 400;     // initialize waits at most this long for discovery
const LIST_GRACE_MS = 400;          // discovery must never push a response towards 1s
// The add-in serves plain JSON-RPC on /rpc (this is what tools/dump*.py use).  /mcp is the
// streamable-HTTP face for MCP clients and can block on a session; the hub talks to the
// add-in over /rpc and only uses /mcp (best effort) for the instructions banner.
const ADDIN_RPC_PATH = process.env.VB6MCP_ADDIN_PATH || '/rpc';
const FACE_CACHE = path.join(HERE, 'hub_face_cache.json');
// ---- endpoint list: the local machine + a configured allow-list of remote endpoints --------
// The hub discriminates by IDE, not by machine: LOCAL and REMOTE IDEs are EQUIVALENT members
// of ONE candidate set and `remote` is only an extra marker (author ruling, task-40).
// Endpoints come from the env (VB6MCP_REMOTES) and/or bridge/hub_remotes.json; with neither
// configured, every remote code path below is inert and the behaviour is the pre-change one.
const REMOTES_FILE = path.join(HERE, 'hub_remotes.json');
const LOCAL_ENDPOINT = 'local';
const LABEL_RE = /^[A-Za-z0-9_.-]{2,32}$/;
const REMOTE_PROBE_TIMEOUT_MS = parseInt(process.env.VB6MCP_REMOTE_PROBE_TIMEOUT_MS || '1500', 10);
const REMOTE_FORWARD_TIMEOUT_MS = parseInt(process.env.VB6MCP_REMOTE_FORWARD_TIMEOUT_MS || String(FORWARD_TIMEOUT_MS), 10);
const REMOTE_TTL_MS = parseInt(process.env.VB6MCP_REMOTE_TTL_MS || String(DISCOVERY_TTL_MS), 10);
const REMOTE_GRACE_MS = parseInt(process.env.VB6MCP_REMOTE_GRACE_MS || '1500', 10);
// The canonical identity rule, restated VERBATIM in every receipt that uses it.
const TARGET_KEY_RULE = 'target key = [endpoint:]projectPath; endpoint omitted = local. Bare pid / bare project NAME stay accepted as LEGACY aliases and are never reported as the canonical key. Normalization: case-insensitive, "/" -> "\\", duplicate + trailing separators trimmed; a LOCAL path is additionally fs.realpathSync.native when the file exists; a REMOTE path is string-normalized only, because the endpoint marker is what disambiguates two machines holding the same path. The same normalized path on ONE endpoint is a conflict (VB6 cannot open the same .vbp twice); the same path on TWO endpoints is NOT a conflict.';
function readEndpointConfig() {
  const out = [], errors = [];
  const push = function (e, source) {
    if (!e || typeof e !== 'object' || Array.isArray(e)) { errors.push({ source: source, error: 'entry_not_an_object', raw: String(e) }); return; }
    const label = String(e.label === undefined || e.label === null ? '' : e.label).trim();
    const url = String(e.url === undefined || e.url === null ? '' : e.url).trim();
    if (!LABEL_RE.test(label) || label.toLowerCase() === LOCAL_ENDPOINT) { errors.push({ source: source, label: label, error: 'invalid_label', hint: '2..32 chars of [A-Za-z0-9_.-]; "local" is reserved' }); return; }
    if (!/^https?:\/\//i.test(url)) { errors.push({ source: source, label: label, error: 'invalid_url', url: url, hint: 'http(s)://host:port' }); return; }
    const rec = { label: label, url: url.replace(/\/+$/, ''), note: (e.note === undefined || e.note === null) ? '' : String(e.note), source: source };
    const dup = out.filter(function (x) { return x.label.toLowerCase() === label.toLowerCase(); })[0];
    if (dup) { out[out.indexOf(dup)] = rec; } else { out.push(rec); }
  };
  const fromFile = readJson(REMOTES_FILE, null);
  const fileArr = Array.isArray(fromFile) ? fromFile : (fromFile && Array.isArray(fromFile.remotes) ? fromFile.remotes : null);
  if (fromFile !== null && !fileArr) errors.push({ source: 'hub_remotes.json', error: 'expected_array_or_{remotes:[]}' });
  if (fileArr) for (const e of fileArr) push(e, 'hub_remotes.json');
  const envRaw = process.env.VB6MCP_REMOTES;
  if (envRaw && envRaw.trim()) {
    let parsed = null, bad = false;
    try { parsed = JSON.parse(envRaw); } catch (e) { bad = true; errors.push({ source: 'VB6MCP_REMOTES', error: 'invalid_json', message: String(e && e.message || e) }); }
    if (!bad) {
      const envArr = Array.isArray(parsed) ? parsed : (parsed && Array.isArray(parsed.remotes) ? parsed.remotes : null);
      if (!envArr) errors.push({ source: 'VB6MCP_REMOTES', error: 'expected_array_or_{remotes:[]}' });
      else for (const e of envArr) push(e, 'VB6MCP_REMOTES');
    }
  }
  return { endpoints: out, errors: errors };
}
const ENDPOINT_CFG = readEndpointConfig();
const REMOTE_ENDPOINTS = ENDPOINT_CFG.endpoints;
const ENDPOINT_CFG_ERRORS = ENDPOINT_CFG.errors;
// Deferred: dbg() reads the DEBUG const that is declared further down, so a top-level call
// here would hit the temporal dead zone.
setTimeout(function () {
  dbg('endpoints: local + ' + REMOTE_ENDPOINTS.length + ' remote(s) ' + JSON.stringify(REMOTE_ENDPOINTS.map(function (e) { return e.label + '=' + e.url; })) + (ENDPOINT_CFG_ERRORS.length ? ' configErrors=' + JSON.stringify(ENDPOINT_CFG_ERRORS) : ''));
}, 0);
function endpointByLabel(label) {
  const s = String(label === undefined || label === null ? '' : label).toLowerCase();
  if (!s) return null;
  if (s === LOCAL_ENDPOINT) return { label: LOCAL_ENDPOINT, url: null, remote: false, note: 'this machine (instance registry + port scan)' };
  for (const e of REMOTE_ENDPOINTS) if (e.label.toLowerCase() === s) return { label: e.label, url: e.url, remote: true, note: e.note, source: e.source };
  return null;
}
function endpointLabels() { return [LOCAL_ENDPOINT].concat(REMOTE_ENDPOINTS.map(function (e) { return e.label; })); }
// A configured endpoint url may carry an explicit rpc path (http://host:8796/rpc); otherwise
// the add-in's default rpc path is used.
function endpointRpcTarget(url) {
  const m = /^(https?:\/\/[^\/]+)(\/.*)?$/i.exec(String(url === undefined || url === null ? '' : url));
  if (!m) return { base: String(url || ''), path: ADDIN_RPC_PATH };
  const p = m[2] || '';
  if (!p || p === '/') return { base: m[1], path: ADDIN_RPC_PATH };
  return { base: m[1], path: p };
}

// ---- deterministic project-path identity: THE canonical target key -------------------------
// VB6 itself refuses to open the same .vbp twice, so the absolute path is the natural unique
// key; pid/port/name are NOT (measured: 38 stale registry variants of one path, one name over
// two paths, one port over two projects).
function normalizePathString(p) {
  let s = String(p === undefined || p === null ? '' : p).trim();
  if (!s) return '';
  s = s.replace(/\//g, '\\');
  const unc = s.indexOf('\\\\') === 0;
  let body = unc ? s.slice(2) : s;
  body = body.replace(/\\+/g, '\\');
  s = (unc ? '\\\\' : '') + body;
  s = s.replace(/\\+$/, '');
  if (/^[A-Za-z]:$/.test(s)) s = s + '\\';
  return s;
}
function canonicalProjectPath(raw, isRemote) {
  const base = normalizePathString(raw);
  if (!base) return { path: '', key: '', via: 'empty' };
  if (!isRemote) {
    try {
      if (fs.existsSync(base)) {
        const rp = normalizePathString(fs.realpathSync.native(base));
        if (rp) return { path: rp, key: rp.toLowerCase(), via: rp === base ? 'realpath_same' : 'realpath' };
      }
    } catch (e) { dbg('canonicalProjectPath realpath failed for ' + base + ': ' + (e && e.message)); }
    return { path: base, key: base.toLowerCase(), via: 'string' };
  }
  return { path: base, key: base.toLowerCase(), via: 'string_remote' };
}
function candidateIdentity(c) {
  const ep = String(c.endpoint || LOCAL_ENDPOINT).toLowerCase();
  if (c.pathKey) return ep + '\u0000' + c.pathKey;
  if (c.project) return ep + '\u0000name:' + String(c.project).toLowerCase();
  return ep + '\u0000pid:' + String(c.pid);
}
// One local registry/scan instance -> a candidate of the SAME set a remote IDE belongs to.
function localCandidate(i) {
  const cp = canonicalProjectPath(i.projectPath || '', false);
  const ppath = cp.path || null;
  return Object.assign({}, i, {
    endpoint: LOCAL_ENDPOINT, remote: false, endpointUrl: null,
    projectPath: ppath, pathKey: cp.key, pathVia: cp.via,
    targetKey: LOCAL_ENDPOINT + ':' + (ppath || ('pid=' + i.pid))
  });
}
function identityConflicts(insts) {
  const groups = new Map();
  for (const i of insts) {
    const k = candidateIdentity(i);
    const g = groups.get(k);
    if (g) g.push(i); else groups.set(k, [i]);
  }
  const out = [];
  for (const v of groups.values()) {
    if (v.length < 2) continue;
    const ep = v[0].endpoint || LOCAL_ENDPOINT;
    out.push('endpoint "' + ep + '": ' + v.length + ' live IDEs publish the same project identity "' + (v[0].projectPath || v[0].project || ('pid ' + v[0].pid)) + '" (pids ' + v.map(function (x) { return x.pid; }).join(',') + ', ports ' + v.map(function (x) { return x.port; }).join(',') + ') -- VB6 cannot open the same .vbp twice on one endpoint; refusing to pick one');
  }
  return out;
}
// ---- lessons (moved to the Python side: tools/lessons.py) --------------------------------
// The 5 lesson tools are HUB tools now, and the rendered lessons text is appended to
// initialize.instructions.  Both go through ONE frozen CLI:
//   python tools/lessons.py <instructions|list|add|delete|export|import-md|validate>
// which prints exactly ONE JSON object on stdout (failure = non-zero exit + {"ok":false}).
// VB6MCP_LESSONS_PY / VB6MCP_PYTHON let a test instance point at a stub without touching
// the shipped tree; the defaults are the frozen deployment.
const LESSONS_PY = process.env.VB6MCP_LESSONS_PY || path.join(HERE, '..', 'tools', 'lessons.py');
const PYTHON_BIN = process.env.VB6MCP_PYTHON || 'python';
const LESSONS_TIMEOUT_MS = parseInt(process.env.VB6MCP_LESSONS_TIMEOUT_MS || '30000', 10);
const LESSONS_RETRY_MS = 30000;     // a FAILED instructions read is retried, not cached forever
// The add-in's instructions banner is fetched from /mcp, which can take seconds
// on a busy add-in - so it is BEST EFFORT and never on the discovery critical path.  The
// first discovery returns without it (the hub then labels the downgrade), later ones carry it.
const instructionsCache = {};
// Instrumentation: VB6MCP_HUB_DEBUG=1 traces every discovery decision to stderr
// (stderr is never the MCP channel, so this cannot corrupt a response).
const DEBUG = process.env.VB6MCP_HUB_DEBUG === '1';
function dbg() {
  if (!DEBUG) return;
  try { process.stderr.write('[hub-debug] ' + Array.prototype.join.call(arguments, ' ') + '\n'); } catch (e) { }
}
// netstat-free liveness probe for a TCP port (used as a fallback, and to test the real
// failure mode with VB6MCP_HUB_NO_NETSTAT=1).
function tcpListening(port, ms) {
  return new Promise(function (resolve) {
    const s = net.connect({ host: '127.0.0.1', port: port });
    let done = false;
    const fin = function (v) { if (!done) { done = true; try { s.destroy(); } catch (e) { } resolve(v); } };
    s.setTimeout(ms || 200);
    s.on('connect', function () { fin(true); });
    s.on('timeout', function () { fin(false); });
    s.on('error', function () { fin(false); });
  });
}
async function isListening(port, listening) {
  if (listening && listening[port] && listening[port].length) return true;
  const t = await tcpListening(port, 200);
  if (t) dbg('isListening port=' + port + ' netstat=absent tcp=YES');
  return t;
}
let listeningCache = { at: 0, map: null };
function listeningPortsCached() {
  if (listeningCache.map && Date.now() - listeningCache.at < 2000) return listeningCache.map;
  const m = listeningPorts();
  listeningCache = { at: Date.now(), map: m };
  return m;
}
function saveFaceCache(d) {
  try { fs.writeFileSync(FACE_CACHE, JSON.stringify({ at: Date.now(), instances: d.instances, conflicts: d.conflicts, staleRegistry: d.staleRegistry, registryDir: d.registryDir }, null, 2), 'utf8'); } catch (e) { }
}
function loadFaceCache() {
  const c = readJson(FACE_CACHE, null);
  if (!c || !Array.isArray(c.instances) || !c.instances.length) return null;
  return c;
}
const state = { target: null };
const startedByHub = new Map(); // pid -> { project, port }

function readJson(p, dflt) { try { return JSON.parse(fs.readFileSync(p, 'utf8')); } catch (e) { return dflt; } }
function snapshot() { return readJson(SNAP, { tools: [], toolCount: 0, contractVersion: 'unknown', generatedFrom: '', breaking: [] }); }
function hubTools() { const d = readJson(HUBCFG, { tools: [] }); return Array.isArray(d.tools) ? d.tools : []; }
function hubToolNames() { return hubTools().map(function (t) { return t.name; }); }
function mcp(product, contract) { return { product: product, contract: contract }; }

function portRange(a, b) { const out = []; for (let p = a; p <= b; p++) out.push(p); return out; }
const DEFAULT_PORTS = portRange(8766, 8799);
let regPortCache = { at: 0, port: null };
function registryPort() {
  if (regPortCache.port !== null && Date.now() - regPortCache.at < 30000) return regPortCache.port;
  let port = null;
  try {
    const out = execFileSync('reg', ['query', 'HKCU\\Software\\VB6MCP', '/v', 'Port'], { encoding: 'utf8', timeout: 3000, stdio: ['ignore', 'pipe', 'ignore'] });
    const m = /Port\s+REG_\w+\s+(\d+)/i.exec(out);
    if (m) port = parseInt(m[1], 10);
  } catch (e) { /* no registry value */ }
  regPortCache = { at: Date.now(), port: port };
  return port;
}
// Candidate ports for the FALLBACK scan.  The HKCU lookup costs a process spawn, so it is
// only consulted when nothing in the default/env range is listening (the common case is an
// IDE already on 8766, which must not pay that cost) and it is cached for 30s.
function candidatePorts(listening) {
  const set = new Set(DEFAULT_PORTS);
  const env = process.env.VB6MCP_PORT;
  if (env && /^\d+$/.test(env)) set.add(parseInt(env, 10));
  let anyListening = false;
  for (const p of set) { if (listening && listening[p] && listening[p].length) { anyListening = true; break; } }
  if (!anyListening) { const rp = registryPort(); if (rp) set.add(rp); }
  return Array.from(set).sort(function (a, b) { return a - b; });
}
// ONE netstat for the whole scan (it used to run once per candidate port: the main
// reason initialize took ~9s).  Returns { port: [pid, ...] } for LISTENING sockets.
function netstatPath() {
  // NEVER guess C:\Windows here.  On this host SystemRoot is
  // B:\Windows and C:\Windows is a stale second installation whose 32-bit binaries fail
  // with 0xC0000139, so inventing a drive is worse than degrading to a PATH lookup.  Only
  // authoritative Windows-directory env vars may be used; when neither is set we say so and
  // let PATH resolve netstat (the pre-existing fallback below).
  const root = process.env.SystemRoot || process.env.windir;
  if (root) {
    const p = path.join(root, 'System32', 'netstat.exe');
    try { if (fs.existsSync(p)) return p; } catch (e) { }
  } else {
    dbg('netstatPath: neither SystemRoot nor windir is set; falling back to a PATH lookup');
  }
  return 'netstat';
}
function listeningPorts() {
  const map = {};
  if (process.env.VB6MCP_HUB_NO_NETSTAT === '1') { dbg('listeningPorts SKIPPED (VB6MCP_HUB_NO_NETSTAT=1)'); return map; }
  const exe = netstatPath();
  try {
    const out = execFileSync(exe, ['-ano', '-p', 'tcp'], { encoding: 'utf8', timeout: 8000, stdio: ['ignore', 'pipe', 'ignore'] });
    dbg('listeningPorts exe=' + exe + ' bytes=' + out.length);
    out.split(/\r?\n/).forEach(function (line) {
      const c = line.trim().split(/\s+/);
      if (c.length >= 5 && c[3] === 'LISTENING') {
        const m = /:(\d+)$/.exec(c[1] || '');
        if (m) { const p = parseInt(m[1], 10); const pid = parseInt(c[4], 10); if (p && pid) (map[p] = map[p] || []).push(pid); }
      }
    });
  } catch (e) { dbg('listeningPorts FAILED exe=' + exe + ' err=' + (e && e.code || '') + ' ' + (e && e.message || '')); }
  return map;
}
function withTimeout(p, ms) {
  return Promise.race([p, new Promise(function (_, rej) { setTimeout(function () { rej(new Error('timeout')); }, ms); })]);
}

// The add-in registers itself at %TEMP%\vb6mcp\instances\<pid>.json.  We read that
// directory first (authoritative) and keep the port scan as a FALLBACK, always labelling
// which one found an instance.  Stale entries are REPORTED, never deleted (other people's
// files, and a live restart may still be writing).
function instancesDir() {
  return process.env.VB6MCP_INSTANCES_DIR || path.join(os.tmpdir(), 'vb6mcp', 'instances');
}
function pidAlive(pid) {
  if (!pid) return false;
  try { process.kill(pid, 0); return true; }
  catch (e) {
    // Only ESRCH means "no such process"; EPERM/anything else means the process exists but we
    // may not signal it (the old version treated everything but EPERM as dead).
    const alive = !!(e && e.code && e.code !== 'ESRCH');
    if (DEBUG) dbg('pidAlive pid=' + pid + ' threw=' + (e && e.code) + ' -> ' + alive);
    return alive;
  }
}
function readRegistry() {
  const dir = instancesDir();
  const entries = [];
  let files = [];
  try { files = fs.readdirSync(dir).filter(function (f) { return f.endsWith('.json'); }); }
  catch (e) { return { dir: dir, entries: entries, unreadable: true }; }
  for (const f of files) {
    const full = path.join(dir, f);
    let j = null;
    try { j = JSON.parse(fs.readFileSync(full, 'utf8')); } catch (e) { entries.push({ file: full, parseError: true }); continue; }
    entries.push({
      file: full,
      pid: parseInt(j.pid, 10) || null,
      port: parseInt(j.port, 10) || null,
      project: j.project || j.projectName || null,
      projectPath: j.projectPath || j.path || null,
      startedAt: j.startedAt || null,
      version: j.version || null,
      dllSha256: j.dllSha256 || j.sha256 || null,
      startedByHub: !!j.startedByHub
    });
  }
  return { dir: dir, entries: entries, unreadable: false };
}

// One JSON-RPC POST against an absolute endpoint base ("http://host:port") - the SAME wire
// shape the local add-in uses, so a remote IDE is reached exactly like a local one.  An
// explicit path in the base (http://host:8796/rpc) wins over the add-in default.
async function rpcAt(base, method, params, proto, timeoutMs, rpcPath) {
  const target = endpointRpcTarget(base);
  const body = { jsonrpc: '2.0', id: (Date.now() % 100000) + 1, method: method };
  if (params !== undefined) body.params = params;
  const headers = { 'Content-Type': 'application/json', 'Accept': 'application/json, text/event-stream' };
  if (proto) headers['MCP-Protocol-Version'] = proto;
  const res = await fetch(target.base + (rpcPath || target.path), {
    method: 'POST', headers: headers, body: JSON.stringify(body), signal: AbortSignal.timeout(timeoutMs || FORWARD_TIMEOUT_MS)
  });
  const raw = await res.text();
  for (const line of raw.split(/\r?\n/)) {
    const t = line.trim();
    const s = t.startsWith('data:') ? t.slice(5).trim() : t;
    if (s.startsWith('{')) return JSON.parse(s);
  }
  return raw.trim() ? JSON.parse(raw) : null;
}
async function rpc(port, method, params, proto, timeoutMs, rpcPath) {
  return rpcAt('http://127.0.0.1:' + port, method, params, proto, timeoutMs, rpcPath);
}

// Runs in the BACKGROUND (never on a response path): a slow add-in must not delay the hub.
async function probe(port, knownProject, pids) {
  pids = pids || [];
  if (!pids.length) return null;
  let tools = [];
  try {
    const tl = await rpc(port, 'tools/list', {}, null, 8000);
    tools = (tl && tl.result && tl.result.tools) || [];
  } catch (e) { return null; }
  if (!tools.length) return null;
  let project = knownProject || null;
  if (!project) {
    try {
      const g = await rpc(port, 'tools/call', { name: 'get_project_info', arguments: {} }, null, 1500);
      const c = g && g.result && g.result.content;
      const txt = Array.isArray(c) && c[0] && c[0].text;
      if (txt) { try { const j = JSON.parse(txt); project = j.project || j.name || null; } catch (e2) { } }
    } catch (e) { }
  }
  // Best-effort banner from the MCP face; /rpc does not implement initialize.  Fetched in the
  // BACKGROUND so a slow /mcp cannot keep the instance "pending".
  const instructions = instructionsCache[port] || null;
  if (!instructions) {
    rpc(port, 'initialize', { protocolVersion: LATEST, capabilities: {}, clientInfo: { name: 'vb6mcp-hub', version: SERVER.version } }, null, 3000, '/mcp')
      .then(function (init) { const v = init && init.result && init.result.instructions; if (v) { instructionsCache[port] = v; dbg('instructions cached for port ' + port); } })
      .catch(function () { dbg('instructions fetch failed for port ' + port); });
  }
  return { port: port, pid: pids[0], pids: pids, project: project, liveTools: tools.map(function (t) { return t.name; }), liveToolDefs: tools, instructions: instructions };
}

// ---- remote endpoints: probe / cache / PER-ENDPOINT single-flight -------------------------
// Keyed per endpoint: local discovery NEVER waits on this map, so a slow or dead remote cannot
// stall a local call, while two concurrent calls to the SAME endpoint share one probe.
const remoteProbe = new Map();   // label -> { at, data, promise }
const remoteInstructionsCache = {};
function remoteCacheGet(label) {
  const st = remoteProbe.get(label);
  if (!st || !st.data) return null;
  if (Date.now() - st.at >= REMOTE_TTL_MS) return null;
  return st.data;
}
function firstText(resp) {
  const c = resp && resp.result && resp.result.content;
  return (Array.isArray(c) && c[0] && typeof c[0].text === 'string') ? c[0].text : null;
}
// Best effort, BACKGROUND: the /mcp banner of a remote endpoint (never on a critical path).
function remoteInstructions(ep, url) {
  if (remoteInstructionsCache[ep.label]) return remoteInstructionsCache[ep.label];
  rpcAt(url, 'initialize', { protocolVersion: LATEST, capabilities: {}, clientInfo: { name: 'vb6mcp-hub', version: SERVER.version } }, null, 3000, '/mcp')
    .then(function (init) { const v = init && init.result && init.result.instructions; if (v) { remoteInstructionsCache[ep.label] = v; dbg('remote instructions cached for ' + ep.label); } })
    .catch(function () { dbg('remote instructions fetch failed for ' + ep.label); });
  return null;
}
async function probeRemoteEndpoint(ep, timeoutMs) {
  const t0 = Date.now();
  const url = ep.url;
  let tl = null;
  try { tl = await rpcAt(url, 'tools/list', {}, null, timeoutMs); }
  catch (e) {
    const isTimeout = !!(e && (e.name === 'TimeoutError' || e.name === 'AbortError' || /abort|timed? ?out/i.test(String(e && e.message))));
    return { ok: false, endpoint: ep.label, endpointUrl: url, remote: true, reason: isTimeout ? 'endpoint_timeout' : 'endpoint_unreachable',
             message: String(e && e.message || e), elapsedMs: Date.now() - t0, timeoutMs: timeoutMs };
  }
  const tools = (tl && tl.result && tl.result.tools) || [];
  let ide = null, project = null, rawPath = '';
  try {
    const mi = await rpcAt(url, 'tools/call', { name: 'mcp_info', arguments: {} }, null, timeoutMs);
    const txt = firstText(mi);
    if (txt) { const j = JSON.parse(txt); if (j && j.ide) ide = j.ide; if (j && j.fileName) rawPath = String(j.fileName); if (j && j.name) project = String(j.name); }
  } catch (e) { dbg('remote ' + ep.label + ' mcp_info failed: ' + (e && e.message)); }
  if (!ide || !ide.project || !rawPath) {
    try {
      const gp = await rpcAt(url, 'tools/call', { name: 'get_project_info', arguments: {} }, null, timeoutMs);
      const txt = firstText(gp);
      if (txt) { const j = JSON.parse(txt); if (j && j.name && !project) project = String(j.name); if (j && j.fileName) rawPath = String(j.fileName); }
    } catch (e) { dbg('remote ' + ep.label + ' get_project_info failed: ' + (e && e.message)); }
  }
  if (ide && ide.project) project = String(ide.project);
  const pid = (ide && ide.pid !== undefined && ide.pid !== null) ? (parseInt(ide.pid, 10) || null) : null;
  const port = (ide && ide.port !== undefined && ide.port !== null) ? (parseInt(ide.port, 10) || null) : null;
  const cp = canonicalProjectPath(rawPath, true);
  const candidate = {
    endpoint: ep.label, remote: true, endpointUrl: url, note: ep.note,
    pid: pid, pids: pid ? [pid] : [], port: port, project: project || null,
    projectPath: cp.path || null, pathKey: cp.key, pathVia: cp.via,
    targetKey: ep.label + ':' + (cp.path || ('pid=' + pid)),
    discoveredBy: 'remote_probe', probeStatus: 'ok', pending: false,
    liveTools: tools.map(function (t) { return t.name; }), liveToolDefs: tools,
    instructions: remoteInstructions(ep, url)
  };
  return { ok: true, endpoint: ep.label, endpointUrl: url, remote: true, candidate: candidate,
           toolCount: tools.length, elapsedMs: Date.now() - t0, timeoutMs: timeoutMs };
}
async function probeEndpointCached(label, timeoutMs) {
  const ep = endpointByLabel(label);
  if (!ep || !ep.remote) return { ok: false, endpoint: String(label), reason: 'endpoint_unknown' };
  const cached = remoteCacheGet(ep.label);
  if (cached) return cached;
  const st = remoteProbe.get(ep.label);
  if (st && st.promise) return st.promise;               // per-endpoint single-flight
  const promise = probeRemoteEndpoint(ep, timeoutMs || REMOTE_PROBE_TIMEOUT_MS).then(function (d) {
    remoteProbe.set(ep.label, { at: Date.now(), data: d, promise: null });
    return d;
  }).catch(function (e) {
    const d = { ok: false, endpoint: ep.label, endpointUrl: ep.url, remote: true, reason: 'endpoint_unreachable', message: String(e && e.message || e), elapsedMs: 0 };
    remoteProbe.set(ep.label, { at: Date.now(), data: d, promise: null });
    return d;
  });
  remoteProbe.set(ep.label, { at: st ? st.at : 0, data: st ? st.data : null, promise: promise });
  return promise;
}
async function remoteCandidatesAvailable() {
  const candidates = [], endpoints = [];
  for (const ep of REMOTE_ENDPOINTS) {
    const st = remoteProbe.get(ep.label);
    const fresh = remoteCacheGet(ep.label);
    const d = fresh || (st ? st.data : null);
    endpoints.push({ label: ep.label, url: ep.url, remote: true, note: ep.note, source: ep.source,
                     cached: !!fresh, pending: !!(st && st.promise), settled: !!d,
                     reason: d ? (d.reason || 'ok') : null,
                     elapsedMs: d && d.elapsedMs !== undefined ? d.elapsedMs : null });
    if (fresh && fresh.ok && fresh.candidate) candidates.push(fresh.candidate);
  }
  return { candidates: candidates, endpoints: endpoints };
}
// A bounded pass over EVERY configured endpoint (all probes in parallel, no cross-endpoint
// serialization).  Anything that has not settled within the grace window is reported as
// pending/unreachable; the probes keep running in the background.
async function remotePass(graceMs) {
  if (!REMOTE_ENDPOINTS.length) return { candidates: [], endpoints: [] };
  await Promise.race([
    Promise.all(REMOTE_ENDPOINTS.map(function (ep) { return probeEndpointCached(ep.label).catch(function () { return null; }); })),
    new Promise(function (r) { setTimeout(r, graceMs === undefined ? REMOTE_GRACE_MS : graceMs); })
  ]);
  return remoteCandidatesAvailable();
}
// THE merged candidate list (local + remote) in ONE set.  With no configured endpoint this is
// a pass-through of the local discovery object, so every local result stays what it was.
async function withRemoteCandidates(localD, graceMs) {
  if (!REMOTE_ENDPOINTS.length) return localD;
  let r = await remoteCandidatesAvailable();
  const settled = r.endpoints.filter(function (e) { return e.settled; }).length;
  if (settled < REMOTE_ENDPOINTS.length) r = await remotePass(graceMs);
  return Object.assign({}, localD, {
    instances: (localD.instances || []).concat(r.candidates),
    endpoints: r.endpoints,
    endpointConfigErrors: ENDPOINT_CFG_ERRORS,
    targetKeyRule: TARGET_KEY_RULE,
    remoteInstanceCount: r.candidates.length
  });
}

async function discoverBounded() {
  const t0 = Date.now();
  const listening = listeningPortsCached();    // ONE netstat for the whole scan
  const insts = [], stale = [], seenPorts = new Set(), pendingCands = [];
  let scanTruncated = false;
  // The registry is the PRIORITY path: it is a handful of small files, so it is never cut
  // off by the budget (cutting it off could hide a live IDE and mislabel the face).
  // The "who owns this port" judgement lives in ONE place (scanRegistry); this
  // function must never re-implement it (the old first-wins seenPorts rule is gone).
  const scan = await scanRegistry(listening);
  dbg('discoverBounded: registry dir=' + scan.dir + ' entries=' + scan.entries + ' live=' + scan.live.length + ' stale=' + scan.stale.length);
  for (const st of scan.stale) stale.push(st);
  for (const c of scan.live) { pendingCands.push(c); seenPorts.add(c.port); }
  // Existence first (registry + pid alive + port listening); the probe only ENRICHES.
  pendingCands.forEach(function (c) { insts.push(Object.assign({ pending: true, probeStatus: 'pending' }, c)); });
  for (const c of pendingCands) {
    const it = await probe(c.port, c.project, c.pids).catch(function () { return null; });
    const idx0 = insts.findIndex(function (x) { return x.port === c.port; });
    if (!it) {
      stale.push({ file: c.registryFile, pid: c.pid, port: c.port, reason: 'registry_port_not_mcp_probe_failed' });
      // Terminal, explicit state: the probe ran and failed -> not pending any more, the tool
      // count stays UNKNOWN (null), never a misleading 0.
      if (idx0 >= 0) insts[idx0] = Object.assign({}, insts[idx0], { pending: false, probeStatus: 'failed', probeReason: 'probe_failed_or_timeout' });
      continue;
    }
    const idx = idx0;
    if (idx >= 0) insts[idx] = Object.assign({}, c, it, { pending: false, probeStatus: 'ok' });
  }
  // Port scan is the FALLBACK (only when the registry yielded nothing); its budget is what
  // remains of DISCOVERY_BUDGET_MS, and ports with no listener are skipped without a probe.
  if (insts.length === 0) {
    const ports = candidatePorts(listening);
    const scanStart = Date.now();                 // probe budget only (netstat/reg excluded)
    for (const port of ports) {
      if (Date.now() - scanStart > DISCOVERY_BUDGET_MS) { scanTruncated = true; break; }
      if (seenPorts.has(port)) continue;
      const pids = listening[port] || [];
      if (!pids.length) continue;
      let it = null;
      try { it = await withTimeout(probe(port, null, pids), 8000); } catch (e) { it = null; }
      seenPorts.add(port);
      if (it) {
        insts.push(Object.assign({}, it, { discoveredBy: 'portscan', registryFile: null, pending: false, probeStatus: 'ok' }));
      } else {
        // The port IS listening, so the instance exists even if the probe fails -
        // report it (unknown face) instead of silently dropping it.
        insts.push({ port: port, pid: pids[0], pids: pids, project: null, discoveredBy: 'portscan', registryFile: null,
                     pending: false, probeStatus: 'failed', probeReason: 'probe_failed_or_timeout' });
      }
    }
  }
  // Identity grouping is by NORMALIZED PROJECT PATH, never by project NAME (one name was
  // measured over TWO different paths, so name-grouping both missed real duplicates and
  // invented false ones).  The same path on ONE endpoint is the conflict; the same path on
  // TWO endpoints is legitimate - that is what the endpoint marker exists for.
  const candidates = insts.map(localCandidate);
  const conflicts = [];
  candidates.forEach(function (i) { if (i.pids.length > 1) conflicts.push('port ' + i.port + ' held by several processes: ' + i.pids.join(',')); });
  identityConflicts(candidates).forEach(function (m) { conflicts.push(m); });
  stale.forEach(function (s) {
    if (s.reason === 'duplicate_port_in_registry' || s.reason === 'registry_port_not_mcp') conflicts.push('registry ' + s.file + ': ' + s.reason);
  });
  return { instances: candidates, conflicts: conflicts, staleRegistry: stale, registryDir: scan.dir,
           registryUnreadable: !!scan.unreadable, scanTruncated: scanTruncated, elapsedMs: Date.now() - t0 };
}

// Cached, single-flight discovery.  A request never waits longer than
// DISCOVERY_BUDGET_MS + slack: callers use discoverSoft().
let discovery = { at: 0, data: null, promise: null };
async function discover() {
  if (discovery.data && Date.now() - discovery.at < DISCOVERY_TTL_MS) return discovery.data;
  if (discovery.promise) return discovery.promise;
  discovery.promise = discoverBounded().then(function (d) {
    discovery = { at: Date.now(), data: d, promise: null };
    // Never persist a fallback/empty result as the authoritative face.
    if (d && d.instances && d.instances.length) { try { saveFaceCache(d); } catch (e) { } }
    return d;
  }).catch(function (e) {
    discovery.promise = null;
    // Hardening (S suggestion): a FAILED discovery must never fabricate an empty
    // face that looks authoritative.  If we already have a resolved face, keep it and mark it;
    // otherwise return the empty face but say why.
    const prev = discovery.data;
    const msg = String(e && e.message || e);
    dbg('discover: FAILED (' + msg + ')' + (prev && prev.instances && prev.instances.length ? ' - keeping the previous face (' + prev.instances.length + ' instance(s))' : ''));
    if (prev && prev.instances && prev.instances.length) {
      return Object.assign({}, prev, { discoveryError: msg, staleFace: true });
    }
    return { instances: [], conflicts: [], staleRegistry: [], registryDir: instancesDir(),
             registryUnreadable: false, scanTruncated: true, error: msg, discoveryError: msg };
  });
  return discovery.promise;
}
// Wait for discovery only up to graceMs; the promise keeps running in the background, so a
// later call sees the live face.  Returns null when the answer is not ready yet.
async function discoverSoft(graceMs) {
  try {
    return await Promise.race([discover(), new Promise(function (r) { setTimeout(function () { r(null); }, graceMs === undefined ? DISCOVERY_GRACE_MS : graceMs); })]);
  } catch (e) { return null; }
}
function invalidateDiscovery() { discovery = { at: 0, data: null, promise: null }; }

// ---- target parsing + candidate selection (ONE endpoint scope at a time) --------------------
// _target syntax: "[endpoint:]selector"; the endpoint is omitted for local.  A Windows drive
// letter is a DRIVE, not an endpoint, so a prefix only qualifies when it names a CONFIGURED
// endpoint label (or the reserved "local").  A prefix that IS a valid label but is not
// configured is an explicit endpoint_unknown error - never a silent local fallback.
function parseTarget(raw) {
  const t = String(raw === undefined || raw === null ? 'auto' : raw).trim();
  if (t === '' || t.toLowerCase() === 'auto') return { kind: 'auto', raw: t };
  // A BARE configured label addresses that endpoint's IDE(s) (selector = all of them); a bare
  // anything-else stays a LOCAL selector, so legacy project NAME aliases keep working.
  const bare = endpointByLabel(t);
  if (bare) return { kind: 'endpoint', endpoint: bare, selector: '', raw: t };
  const m = /^([^:\\\/]+):(.*)$/.exec(t);
  if (m) {
    const ep = endpointByLabel(m[1]);
    if (ep) return { kind: 'endpoint', endpoint: ep, selector: m[2], raw: t };
    // "<label>:<selector>" with a VALID but unconfigured label is an explicit error; a Windows
    // drive ("D:\x.vbp") is one character and therefore never qualifies.
    if (LABEL_RE.test(m[1])) return { kind: 'endpoint_unknown', label: m[1], selector: m[2], raw: t };
  }
  return { kind: 'local', selector: t, raw: t };
}
function selectorIsPathLike(s) {
  const x = String(s === undefined || s === null ? '' : s);
  return x.indexOf('\\') >= 0 || x.indexOf('/') >= 0 || /^[A-Za-z]:/.test(x);
}
// The canonical key is matched FIRST; bare pid / bare port / bare NAME stay accepted as legacy
// aliases so every existing caller keeps working unchanged.
function matchCandidates(insts, s, isRemoteScope) {
  const sel = String(s);
  const cp = canonicalProjectPath(sel, !!isRemoteScope);
  return insts.filter(function (i) {
    if (cp.key && i.pathKey && i.pathKey === cp.key) return true;
    if (String(i.pid) === sel) return true;
    if (String(i.port) === sel) return true;
    if (i.project && i.project === sel) return true;
    return false;
  });
}
function matchedByFor(cand, s, isRemoteScope) {
  const sel = String(s);
  const cp = canonicalProjectPath(sel, !!isRemoteScope);
  if (cp.key && cand.pathKey && cand.pathKey === cp.key) return 'project_path';
  if (String(cand.pid) === sel) return 'pid';
  if (String(cand.port) === sel) return 'port';
  if (cand.project && cand.project === sel) return 'project_name';
  return 'unknown';
}
function candidateDesc(insts, isRemoteScope) {
  if (!isRemoteScope) return insts.map(function (i) { return i.port + '/' + (i.project || '?'); }).join(', ');
  return insts.map(function (i) { return (i.endpoint || '?') + ':' + (i.projectPath || i.project || ('port ' + i.port)); }).join(', ');
}
function candidateBrief(i) {
  return { pid: i.pid, port: i.port, project: i.project, projectPath: i.projectPath || null,
           endpoint: i.endpoint || LOCAL_ENDPOINT, remote: !!i.remote, targetKey: i.targetKey || null };
}
// Selection inside ONE endpoint scope.  target = a resolved selector (or null); explicit =
// the selector came from _target (a miss is an error) vs from the ide_target DEFAULT (a miss
// is ignored, never vetoing the live set).
function pickTarget(insts, target, opts) {
  opts = opts || {};
  const isRemoteScope = !!opts.remoteScope;
  const explicit = !!opts.explicit;
  if (!insts.length) return { ok: false, message: isRemoteScope ? 'the selected endpoint reports no IDE' : 'no IDE instance listening (scanned ' + DEFAULT_PORTS.length + ' default ports + registry)' };
  const desc = function () { return candidateDesc(insts, isRemoteScope); };
  if (target !== null && target !== undefined && target !== '') {
    const s = String(target);
    const hit = matchCandidates(insts, s, isRemoteScope);
    if (hit.length === 1) return { ok: true, ...hit[0], selectedBy: (explicit ? 'explicit_target' : 'effective_default'),
                                   matchedBy: matchedByFor(hit[0], s, isRemoteScope), matchedFrom: s };
    if (hit.length > 1) {
      const keySet = {};
      hit.forEach(function (i) { keySet[candidateIdentity(i)] = 1; });
      const ids = Object.keys(keySet);
      const sameIdentity = ids.length === 1;
      const ep = hit[0].endpoint || LOCAL_ENDPOINT;
      return { ok: false, conflict: true, identityConflict: sameIdentity, distinctIdentities: ids.length, matchedCount: hit.length,
        hits: hit.map(candidateBrief),
        message: sameIdentity
          ? 'target "' + s + '" is ONE project identity ("' + (hit[0].projectPath || hit[0].project || '?') + '") served by ' + hit.length + ' live IDEs on endpoint "' + ep + '": ' + hit.map(function (i) { return i.pid + '@' + i.port; }).join(',') + ' -- VB6 cannot open the same .vbp twice; refusing to pick one'
          : 'target "' + s + '" is AMBIGUOUS: it matches ' + hit.length + ' live IDEs with ' + ids.length + ' different project identities on endpoint "' + ep + '": ' + hit.map(function (i) { return i.pid + '@' + i.port + ' ' + (i.projectPath || i.project || '?'); }).join(' | ') + ' -- pick one by its projectPath' };
    }
    // A STALE DEFAULT (ide_target pointing at a project that is gone) must
    // never veto the live instances - that made every implicit call fail with no_ide_running
    // while the only live IDE was right there.
    if (explicit) {
      return { ok: false, mismatch: true, message: 'target "' + s + '" not found among ' + desc() + ' -- pick one with _target (pid|project|port) or ide_target' };
    }
    if (insts.length === 1) return { ok: true, ...insts[0], staleTarget: s, selectedBy: 'unique_live' };
    return { ok: false, ambiguous: true, staleTarget: s, message: 'the default ide_target "' + s + '" matches no live instance and several are online: ' + desc() + ' -- set _target or ide_target' };
  }
  if (insts.length === 1) return { ok: true, ...insts[0], selectedBy: 'unique_live' };
  return { ok: false, ambiguous: true, message: 'several instances online and no target set: ' + desc() + ' -- set _target (pid|project|port) or ide_target' };
}

// Explicit _target can address the add-in WITHOUT discovery - if the target
// is a listening port (or a pid in the registry, or a project name), forward straight to it.
// This is the temporary route when discovery is blind but the IDE is reachable.
async function directTarget(target) {
  const s = String(target === undefined || target === null ? '' : target);
  if (!s) return null;
  if (/^\d+$/.test(s)) {
    const n = parseInt(s, 10);
    if (n >= 1 && n <= 65535 && await tcpListening(n, 150)) {
      // An explicit PORT stays a deliberate escape hatch (works even when the registry is
      // blind), but when we DO know the owner we report its real pid/project.
      const own = await ownerOfPort(n);
      return { ok: true, port: n, pid: own ? own.pid : null, pids: own ? own.pids : [], project: own ? own.project : null, direct: true };
    }
    const e = await ownerBy(function (x) { return String(x.pid) === s; });
    if (e) return { ok: true, port: e.port, pid: e.pid, pids: e.pids, project: e.project || null, direct: true };
    return null;
  }
  const e2 = await ownerBy(function (x) { return x.project === s; });
  if (e2) return { ok: true, port: e2.port, pid: e2.pid, pids: e2.pids, project: e2.project || null, direct: true };
  // A PATH selector addresses the add-in without discovery too: registry entries are matched
  // by the SAME canonical key used everywhere else.
  const cp = canonicalProjectPath(s, false);
  if (cp.key) {
    const e3 = await ownerBy(function (x) { return canonicalProjectPath(x.projectPath || '', false).key === cp.key; });
    if (e3) return Object.assign({ ok: true }, localCandidate(e3), { direct: true, matchedBy: 'project_path', matchedFrom: s });
  }
  return null;
}

// An explicit target may only address a registry entry that scanRegistry()
// considers the OWNER of its port.  Matching raw files would let a stale file (its pid recycled
// by a live process, the port listening because another instance owns it) hijack _target.
async function ownerBy(pred) {
  const s = await scanRegistry(null);
  const hits = s.live.filter(pred);
  return hits.length ? hits[0] : null;
}
async function ownerOfPort(port) {
  const s = await scanRegistry(null);
  const hits = s.live.filter(function (x) { return String(x.port) === String(port); });
  return hits.length ? hits[0] : null;
}

// The body of every text response carries the version stamp; `ide` is added for forwards.
function bodyOf(obj, ide, contract, note) {
  const o = Object.assign({}, obj);
  if (!o.mcp) o.mcp = mcp(SERVER.version, contract);
  // (4): never overwrite an `ide` the add-in already reported.
  if (ide && !o.ide) o.ide = ide;
  if (note) {
    if (typeof note === 'string') { if (!o.targetNote) o.targetNote = note; }
    else { for (const k of Object.keys(note)) { if (o[k] === undefined) o[k] = note[k]; } }
  }
  return JSON.stringify(o, null, 2);
}

function textResult(obj, contract) {
  return { content: [{ type: 'text', text: bodyOf(obj, null, contract || snapshot().contractVersion) }], isError: false };
}

// tools/call result keys are EXACTLY {content,isError}; ide + mcp go into the body text.
function stampBody(result, ide, contract, note) {
  const content = Array.isArray(result && result.content) ? result.content.map(function (c) { return Object.assign({}, c); }) : [];
  let textItem = null;
  for (const c of content) { if (c && c.type === 'text' && typeof c.text === 'string') { textItem = c; break; } }
  if (textItem) {
    let obj = null;
    try { const j = JSON.parse(textItem.text); if (j && typeof j === 'object' && !Array.isArray(j)) obj = j; } catch (e) { }
    if (obj) { textItem.text = bodyOf(obj, ide, contract, note); }
    else { textItem.text = textItem.text + '\n' + bodyOf({}, ide, contract, note); }
  } else {
    content.push({ type: 'text', text: bodyOf({ ok: true }, ide, contract, note) });
  }
  return { content: content, isError: result && result.isError === true };
}

// ---- lessons: one spawn per lesson call, one spawn per PROCESS for the banner --------
function runLessons(cliArgs) {
  return new Promise(function (resolve) {
    // PYTHONIOENCODING pins stdout to UTF-8 (the lesson text is Chinese on this host).
    const env = Object.assign({}, process.env, { PYTHONIOENCODING: 'utf-8', PYTHONUTF8: '1' });
    execFile(PYTHON_BIN, [LESSONS_PY].concat(cliArgs),
      { encoding: 'utf8', timeout: LESSONS_TIMEOUT_MS, windowsHide: true, maxBuffer: 16 * 1024 * 1024, env: env },
      function (err, stdout, stderr) {
        const raw = (stdout === undefined || stdout === null) ? '' : String(stdout);
        let parsed = null;
        try { const t = raw.trim(); if (t) parsed = JSON.parse(t); } catch (e) { parsed = null; }
        resolve({
          parsed: parsed,
          exitCode: (err && typeof err.code === 'number') ? err.code : (err ? 1 : 0),
          error: err ? String(err.message || err) : null,
          stdout: raw,
          stderr: (stderr === undefined || stderr === null) ? '' : String(stderr)
        });
      });
  });
}
// The frozen CLI mapping: tool arguments -> tools/lessons.py argv.
function lessonCliArgs(name, args) {
  args = args || {};
  const v = function (k) { return (args[k] === undefined || args[k] === null) ? null : String(args[k]); };
  if (name === 'list_lessons') return ['list'].concat(args.includeDisabled ? ['--include-disabled'] : []);
  if (name === 'export_lessons') return ['export'].concat(args.includeDisabled ? ['--include-disabled'] : []);
  if (name === 'delete_lesson') return ['delete', '--id', v('id') === null ? '' : v('id')];
  if (name === 'import_lessons_md') return ['import-md', '--path', v('path') === null ? '' : v('path')].concat(args.dryRun ? ['--dry-run'] : []);
  if (name === 'add_lesson') {
    const a = ['add'];
    for (const k of ['id', 'scenario', 'trap', 'do', 'scope']) { const x = v(k); if (x !== null) a.push('--' + k, x); }
    if (args.enabled !== undefined) a.push('--enabled', args.enabled ? '1' : '0');
    return a;
  }
  return null;
}
// Rendered lessons text, cached per process: initialize is frequent and must never spawn
// python more than once per process (tools/list never touches python at all).  A FAILED
// read is retried after LESSONS_RETRY_MS (the module may not be deployed yet).
let lessonsCache = { text: '', ok: false, at: 0, error: null, inFlight: null };
function lessonsText() {
  if (lessonsCache.ok) return Promise.resolve(lessonsCache.text);
  if (lessonsCache.error && Date.now() - lessonsCache.at < LESSONS_RETRY_MS) return Promise.resolve('');
  if (lessonsCache.inFlight) return lessonsCache.inFlight;
  lessonsCache.inFlight = runLessons(['instructions']).then(function (r) {
    lessonsCache.inFlight = null;
    lessonsCache.at = Date.now();
    const p = r.parsed;
    if (p && p.ok !== false && typeof p.text === 'string' && p.text.length) {
      lessonsCache.ok = true; lessonsCache.text = p.text; lessonsCache.error = null;
      dbg('lessons: instructions cached (' + p.text.length + ' chars)');
    } else {
      lessonsCache.ok = false; lessonsCache.text = '';
      lessonsCache.error = (p && (p.error || p.message)) || r.error || String(r.stderr || '').trim() || 'lessons_cli_failed';
      dbg('lessons: instructions UNAVAILABLE (' + lessonsCache.error + ') - initialize continues without the lessons block');
    }
    return lessonsCache.text;
  }).catch(function (e) {
    lessonsCache.inFlight = null; lessonsCache.at = Date.now();
    lessonsCache.ok = false; lessonsCache.text = '';
    lessonsCache.error = String(e && e.message || e);
    return '';
  });
  return lessonsCache.inFlight;
}
// The lessons block is appended to EVERY branch and is never appended twice, so the composed
// instructions are byte-identical whether the add-in banner already carries the block (the
// pre-rebuild DLL still does) or not (after the move).
function composeInstructions(base, lessons) {
  if (!lessons) return base;
  if (!base) return lessons;
  const firstLine = lessons.split(/\r?\n/).filter(function (l) { return l.trim().length; })[0] || '';
  if (firstLine && base.indexOf(firstLine) >= 0) return base;      // already carries the block
  const nl = base.indexOf('\r\n') >= 0 ? '\r\n' : '\n';
  return base.replace(/[\r\n]+$/, '') + nl + nl + lessons;
}
// tools/call result keys stay EXACTLY {content,isError}; the lesson JSON goes into the body
// (with the usual mcp stamp) and {"ok":false} maps to isError.
function lessonsResult(obj, contract) {
  return { content: [{ type: 'text', text: bodyOf(obj, null, contract) }], isError: !!(obj && obj.ok === false) };
}
// `echo` carries the control-plane provenance of the request (r-a/r-b: lessons are GLOBAL;
// _target never routes them, and a supplied target is echoed back rather than silently dropped).
async function lessonsTool(name, args, contract, echo) {
  const cliArgs = lessonCliArgs(name, args);
  const r = await runLessons(cliArgs);
  const p = r.parsed;
  const pre = echo || {};
  if (p && typeof p === 'object' && !Array.isArray(p)) return lessonsResult(Object.assign({}, pre, p), contract);
  return lessonsResult(Object.assign({}, pre, {
    ok: false, reason: 'lessons_cli_failed', tool: name,
    cli: [PYTHON_BIN, LESSONS_PY].concat(cliArgs),
    exitCode: r.exitCode, error: r.error,
    stdout: r.stdout.slice(0, 2000), stderr: r.stderr.slice(0, 2000),
    hint: 'tools/lessons.py must print exactly ONE JSON object on stdout; non-zero exit + {"ok":false,...} is the documented failure report'
  }), contract);
}

async function buildInstructions(d, snap) {
  const inst = d.instances.length ? d.instances[0] : null;
  let base;
  if (inst && inst.instructions) {
    base = inst.instructions;
  } else if (inst && snap.instructions) {
    const n = (inst.liveTools && inst.liveTools.length) || snap.tools.length;
    base = 'PRODUCT vb6mcp-hub ' + SERVER.version + '  BRIDGE-TO vb6-mcp ' + snap.contractVersion + '  TOOLS ' + n + '  (faceSource=live; add-in sent no instructions)' + (inst.pending ? '  [probe pending]' : '') + '\n' + snap.instructions;
  } else {
    const mm = snap.versionToolCountMismatch ? '  MISMATCH contract-declares=' + snap.contractDeclaredToolCount + ' snapshot=' + snap.tools.length : '';
    const head = 'PRODUCT vb6mcp-hub ' + SERVER.version + '  BRIDGE-TO vb6-mcp ' + snap.contractVersion + '  TOOLS ' + snap.tools.length + '  (faceSource=snapshot; no IDE online)' + mm;
    const brk = (snap.breaking && snap.breaking.length) ? snap.breaking : ['- (snapshot has no BREAKING extract; read docs/CONTRACTS.md)'];
    base = [head, 'BREAKING (read before calling):'].concat(brk).join('\n');
  }
  // ALWAYS append the rendered lessons text when it is non-empty.
  return composeInstructions(base, await lessonsText());
}

// Last persisted face, used only when discovery has not answered within the grace window.
// Instances are re-validated cheaply (pid alive AND port listening) - no RPC.
async function cachedDiscovery() {
  const c = loadFaceCache();
  if (!c) { dbg('cachedDiscovery: no face cache file'); return null; }
  const checked = [];
  for (const i of c.instances) {
    // TCP-only: no netstat on a response path
    const okp = pidAlive(i.pid) && await isListening(i.port, null);
    if (okp) checked.push(localCandidate(i));
  }
  dbg('cachedDiscovery: file instances=' + c.instances.length + ' still-alive=' + checked.length);
  if (!checked.length) return null;
  return { instances: checked, conflicts: c.conflicts || [], staleRegistry: c.staleRegistry || [],
           registryDir: c.registryDir, registryUnreadable: false, scanTruncated: false, fromFaceCache: true };
}
// EXISTENCE must not depend on the MCP probe.  A registry entry whose pid is
// alive AND whose port is listening is a live IDE, full stop - the probe only enriches it
// (project / liveToolDefs / instructions).  This is what makes ide_list/initialize report
// faceSource=live + discoveredBy=registry even when the add-in is slow to answer.
// SINGLE SOURCE OF TRUTH for "which registry entry owns a port".
// Every consumer (discoverBounded / registryCandidates / directTarget / ide_list / faceState)
// must go through here - two implementations of this judgement is exactly the defect the card
// describes (discoverBounded used to keep the FIRST entry read, so an unrelated live pid that
// had recycled a dead entry's pid could be reported as the owner while the real IDE was pushed
// into staleRegistry as a duplicate).
//
// Credibility rule: an entry counts only when its pid is alive AND its port is listening;
// when several credible entries claim one port the one with the newest startedAt wins, and
// every loser is reported in `stale` with an explicit reason.
async function scanRegistry(listening) {
  const reg = readRegistry();
  const live = [], stale = [], byPort = new Map(), seen = {};
  for (const e of reg.entries) {
    const base = path.basename(e.file);
    if (e.parseError) { stale.push({ file: base, reason: 'unparsable_json' }); continue; }
    if (!e.pid || !e.port) { stale.push({ file: base, pid: e.pid, port: e.port, reason: 'missing_pid_or_port' }); continue; }
    const alive = pidAlive(e.pid);
    const listeningNow = await isListening(e.port, listening || null);
    dbg('registry ' + base + ' pid=' + e.pid + ' port=' + e.port + ' startedAt=' + (e.startedAt || '-') + ' pidAlive=' + alive + ' listening=' + listeningNow);
    if (!alive) { stale.push({ file: base, pid: e.pid, port: e.port, reason: 'pid_dead' }); continue; }
    if (!listeningNow) { stale.push({ file: base, pid: e.pid, port: e.port, reason: 'port_not_listening' }); continue; }
    const prev = seen[e.port];
    // The LOSER goes to staleRegistry, never the winner.
    if (prev && !(String(e.startedAt || '') > String(prev.at || ''))) {
      stale.push({ file: base, pid: e.pid, port: e.port, reason: 'duplicate_port_in_registry' });
      continue;
    }
    if (prev) {
      // we are newer: the PREVIOUS owner is the one that gets demoted
      stale.push({ file: prev.file, pid: prev.pid, port: e.port, reason: 'superseded_older_duplicate_port' });
    }
    seen[e.port] = { at: String(e.startedAt || ''), file: base, pid: e.pid };
    let pids = (listening && listening[e.port]) || [];
    if (!pids.length) pids = [e.pid];
    byPort.delete(e.port);
    byPort.set(e.port, {
      port: e.port, pid: e.pid, pids: pids,
      project: e.project || null, projectPath: e.projectPath || null,
      version: e.version || null, dllSha256: e.dllSha256 || null,
      startedAt: e.startedAt || null, discoveredBy: 'registry', registryFile: base,
      startedByHub: startedByHub.has(e.pid) || e.startedByHub,
      pending: true, probeStatus: 'pending'
    });
  }
  for (const v of byPort.values()) live.push(v);
  dbg('scanRegistry -> live=' + live.length + ' stale=' + stale.length);
  return { dir: reg.dir, unreadable: !!reg.unreadable, entries: reg.entries.length, live: live, stale: stale };
}

async function registryCandidates() {
  const s = await scanRegistry(null);
  if (s.unreadable) return null;
  if (!s.live.length && !s.stale.length) return null;
  const out = s.live.map(localCandidate);
  return { instances: out, conflicts: identityConflicts(out), staleRegistry: s.stale, registryDir: s.dir,
           registryUnreadable: false, scanTruncated: false, fromRegistryCandidates: true };
}
async function resolveDiscovery(d) {
  if (d) { dbg('resolveDiscovery: using completed discovery instances=' + d.instances.length); return d; }
  const c = await cachedDiscovery();
  if (c) return c;
  const r = await registryCandidates();
  if (r) return r;
  dbg('resolveDiscovery: NOTHING found (snapshot fallback)');
  return { instances: [], conflicts: [], staleRegistry: [], registryDir: instancesDir(), registryUnreadable: false, scanTruncated: true };
}
async function faceState(d) {
  d = await resolveDiscovery(d);
  // A cached/persisted face must NEVER freeze a dead instance: re-validate pid + port here,
  // on every request, before it is reported.  A REMOTE candidate is validated by its endpoint
  // probe instead - its pid/port live on ANOTHER machine, so a local pid test is meaningless
  // (a recycled remote pid must not silently drop a healthy remote IDE).
  if (d.instances && d.instances.length) {
    const alive = [];
    for (const i of d.instances) {
      if (i.remote) { alive.push(i); continue; }
      const okp = pidAlive(i.pid) && await isListening(i.port, null);
      if (okp) alive.push(i); else dbg('faceState: dropped dead instance pid=' + i.pid + ' port=' + i.port);
    }
    if (alive.length !== d.instances.length) d = Object.assign({}, d, { instances: alive, revalidated: true });
  }
  const snap = snapshot();
  const live = d.instances.length ? d.instances[0].liveTools : null;
  const snapNames = snap.tools.map(function (t) { return t.name; });
  let diff = null;
  if (live) {
    const onlySnap = snapNames.filter(function (n) { return live.indexOf(n) < 0; });
    const onlyLive = live.filter(function (n) { return snapNames.indexOf(n) < 0; });
    diff = { onlyInSnapshot: onlySnap, onlyInLive: onlyLive, aligned: onlySnap.length === 0 && onlyLive.length === 0 };
  }
  // ONE identity pass over the MERGED set: the same normalized path on the SAME endpoint is a
  // conflict (VB6 forbids it); the same path on TWO endpoints is legitimate.
  const localInsts = d.instances.filter(function (i) { return !i.remote; });
  const remoteInsts = d.instances.filter(function (i) { return !!i.remote; });
  const idCount = {};
  d.instances.forEach(function (i) { const k = candidateIdentity(i); idCount[k] = (idCount[k] || 0) + 1; });
  const conflicts = (d.conflicts || []).slice();
  identityConflicts(d.instances).forEach(function (m) { if (conflicts.indexOf(m) < 0) conflicts.push(m); });
  return {
    // Existence comes from the registry (pid alive + port listening), NOT from
    // whether the MCP probe has answered yet.  A pending probe is still a live IDE.
    faceSource: d.instances.length ? 'live' : 'snapshot',
    contractVersion: snap.contractVersion,
    toolCount: live ? live.length : snap.tools.length,
    snapshot: { contractVersion: snap.contractVersion, toolCount: snap.tools.length, generatedFrom: snap.generatedFrom,
                contractDeclaredToolCount: snap.contractDeclaredToolCount, versionToolCountMismatch: !!snap.versionToolCountMismatch },
    instances: d.instances.map(function (i) { return {
      pid: i.pid, port: i.port, project: i.project, projectPath: i.projectPath || null, pids: i.pids,
      startedAt: i.startedAt || null,
      probeStatus: i.probeStatus || (i.pending ? 'pending' : 'ok'),
      liveToolCount: (i.pending || i.probeStatus === 'failed') ? null : ((i.liveTools && i.liveTools.length) || null),
      probeReason: i.probeReason || null,
      startedByHub: startedByHub.has(i.pid) || !!i.startedByHub,
      discoveredBy: i.discoveredBy || 'portscan', registryFile: i.registryFile || null,
      version: i.version || null, dllSha256: i.dllSha256 || null, pending: !!i.pending,
      // ---- the ONE-set identity (ADDITIVE): a remote IDE is an equal member -----------------
      endpoint: i.endpoint || LOCAL_ENDPOINT,
      remote: !!i.remote,
      endpointUrl: i.endpointUrl || null,
      targetKey: i.targetKey || null,
      pathVia: i.pathVia || null,
      identityConflict: (idCount[candidateIdentity(i)] || 0) > 1
    }; }),
    conflicts: conflicts,
    // Observability: WHY this face looks the way it does - how many registry
    // entries were read, how many were credible, and the stale reasons by kind.
    registrySummary: {
      entries: localInsts.length + (d.staleRegistry || []).length,
      live: localInsts.length,
      stale: (d.staleRegistry || []).length,
      staleByReason: (d.staleRegistry || []).reduce(function (m, x) { const k = x.reason || 'unknown'; m[k] = (m[k] || 0) + 1; return m; }, {})
    },
    staleRegistry: d.staleRegistry || [],
    registryDir: d.registryDir || null,
    registryUnreadable: !!d.registryUnreadable,
    selectionRule: '_target (EXPLICIT) > ide_target default > the ONLY live instance; a STALE default is ignored (never vetoes); otherwise refuse with ambiguous_target / target_mismatch (never guess). The canonical selector is [endpoint:]projectPath; bare pid / bare NAME are legacy aliases.',
    targetKeyRule: TARGET_KEY_RULE,
    endpointLabels: endpointLabels(),
    target: state.target,
    liveVsSnapshot: diff,
    scanTruncated: !!d.scanTruncated,
    discoveryElapsedMs: d.elapsedMs === undefined ? null : d.elapsedMs,
    fromFaceCache: !!d.fromFaceCache,
    discoveryError: d.discoveryError || null,
    staleFace: !!d.staleFace
  };
}

// ONE VB6.EXE resolver shared by every hub entry point, sourced from
// the same places the local launcher tooling uses.
// consulted ONLY VB6MCP_VB6EXE and %ProgramFiles(x86)%, so on this host
// (ProgramFiles(x86)=B:) ide_restart answered vb6_exe_not_found by default while the
// the local launcher happily resolved C:\Program Files (x86)\Microsoft Visual Studio\VB98\VB6.EXE.
// Order: the launcher's known paths first, then both ProgramFiles flavours, then the
// standard App Paths registry entry (name-agnostic, so no install value name is guessed).
const VB6_EXE_CANDIDATES = [
  { p: 'C:\\Program Files (x86)\\Microsoft Visual Studio\\VB98\\VB6.EXE', s: 'known_install_path' },
  { p: 'C:\\Program Files\\Microsoft Visual Studio\\VB98\\VB6.EXE', s: 'known_install_path_x64' }
];
const VB6_EXE_REG_KEYS = [
  'HKLM\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\App Paths\\VB6.EXE',
  'HKLM\\SOFTWARE\\WOW6432Node\\Microsoft\\Windows\\CurrentVersion\\App Paths\\VB6.EXE'
];

function resolveVb6Exe() {
  const tried = [];
  const cands = [];
  if (process.env.VB6MCP_VB6EXE) cands.push({ p: process.env.VB6MCP_VB6EXE, s: 'VB6MCP_VB6EXE' });
  for (const c of VB6_EXE_CANDIDATES) cands.push(c);
  const pf86 = process.env['ProgramFiles(x86)'] || '';
  const pf = process.env['ProgramFiles'] || '';
  if (pf86) cands.push({ p: path.join(pf86, 'Microsoft Visual Studio', 'VB98', 'VB6.EXE'), s: 'ProgramFiles(x86)' });
  if (pf) cands.push({ p: path.join(pf, 'Microsoft Visual Studio', 'VB98', 'VB6.EXE'), s: 'ProgramFiles' });
  for (const c of cands) {
    tried.push(c.p);
    try { if (c.p && fs.existsSync(c.p)) return { exe: c.p, source: c.s, tried: tried }; } catch (e) { }
  }
  for (const k of VB6_EXE_REG_KEYS) {
    tried.push('reg:' + k);
    let v = '';
    try {
      const outReg = execFileSync('reg.exe', ['query', k, '/ve'], { windowsHide: true, timeout: 4000, encoding: 'utf8' });
      const m = /REG_SZ\s+(.+?)\s*$/m.exec(String(outReg));
      if (m) v = m[1];
    } catch (e) { v = ''; }
    if (v) {
      try { if (fs.existsSync(v)) return { exe: v, source: 'registry', tried: tried }; } catch (e) { }
    }
  }
  return { exe: '', source: '', tried: tried };
}

async function hubState() {
  // ide_list is the ONE read that shows the whole set (local + configured remote endpoints).
  // tools/list + initialize deliberately stay on the LOCAL path so a slow remote can never
  // push the MCP handshake towards its 1s budget.
  const localD = await resolveDiscovery(await discoverSoft(LIST_GRACE_MS));
  const d = await withRemoteCandidates(localD, REMOTE_GRACE_MS);
  const st = await faceState(d);
  if (REMOTE_ENDPOINTS.length) {
    st.remote = {
      configured: REMOTE_ENDPOINTS.length,
      instances: (d.instances || []).filter(function (i) { return !!i.remote; }).length,
      endpoints: (d.endpoints || []).map(function (e) { return Object.assign({}, e, { allowList: 'VB6MCP_REMOTES / bridge/hub_remotes.json (never discovered)' }); }),
      configErrors: d.endpointConfigErrors || ENDPOINT_CFG_ERRORS || []
    };
  }
  return st;
}

async function hubTool(name, args) {
  const snap = snapshot();
  if (name === 'ide_list') return textResult(await hubState(), snap.contractVersion);
  if (name === 'ide_target') {
    const t = args && args.target;
    state.target = (t === 'auto' || t === '') ? null : t;
    return textResult({ ok: true, target: state.target }, snap.contractVersion);
  }
  if (name === 'ide_start') {
    const project = (args && args.project) || process.env.VB6MCP_AUTOSTART_PROJECT;
    // The same resolver ide_restart uses, so BOTH hub starters accept
    // any VB6.EXE the launcher would accept (known install paths + both ProgramFiles
    // flavours + the standard App Paths registry entry), not only the env override.
    const rStartExe = resolveVb6Exe();
    const exe = rStartExe.exe;
    if (!project) return textResult({ ok: false, reason: 'not_implemented', message: 'ide_start does not auto-start without a project: pass project or set VB6MCP_AUTOSTART_PROJECT' }, snap.contractVersion);
    if (!exe) return textResult({ ok: false, reason: 'vb6_exe_not_found', vb6Exe: '', vb6ExeSource: '', vb6ExeTried: rStartExe.tried, message: 'no VB6.EXE was resolvable; set VB6MCP_VB6EXE or install VB6. Tried: ' + rStartExe.tried.join(' | ') }, snap.contractVersion);
    if (!fs.existsSync(project)) return textResult({ ok: false, reason: 'project_not_found', project: project }, snap.contractVersion);
    const child = spawn(exe, [project], { detached: true, stdio: 'ignore' });
    child.unref();
    startedByHub.set(child.pid, { project: project, port: null });
    invalidateDiscovery();
    return textResult({ ok: true, started: true, pid: child.pid, project: project, note: 'wait for the add-in to listen, then ide_list' }, snap.contractVersion);
  }
  if (name === 'ide_stop') {
    const pid = args && args.pid;
    if (!pid) return textResult({ ok: false, reason: 'not_implemented', message: 'ide_stop requires an explicit pid and only stops hub-started instances' }, snap.contractVersion);
    if (!startedByHub.has(pid)) return textResult({ ok: false, reason: 'not_allowed', message: 'pid ' + pid + ' was not started by this hub; not stopping it' }, snap.contractVersion);
    try { process.kill(pid); startedByHub.delete(pid); invalidateDiscovery(); return textResult({ ok: true, stopped: pid }, snap.contractVersion); }
    catch (e) { return textResult({ ok: false, reason: 'kill_failed', message: String(e && e.message || e) }, snap.contractVersion); }
  }
  if (name === 'ide_restart') {
    // The executable recovery for a wedged IDE. It only ever
    // touches OUR OWN instance (instances listening on port >= 8767; 8766 belongs to a
    // foreign process on this machine and is NEVER stopped), and EVERY path answers with
    // the same key set (discipline 98).
    const pause = function (ms) { return new Promise(function (r) { setTimeout(r, ms); }); };
    const out = {
      ok: false, tool: 'ide_restart', restarted: false,
      stoppedPid: 0, newPid: 0, port: 0, project: '',
      targetSource: '', ownership: '', refused_reason: '', stopConfirmed: false, pinPort: 0, projectVbp: '',
      saveAttempted: false, saveOk: false, registryCleared: false, lockCleared: false,
      three_way_ok: false, threeWay: { http: false, mcp_info: false, project: false },
      vb6Exe: '', vb6ExeSource: '', vb6ExeTried: [], waitMs: 0, readyWaitMs: 0, instances: [], steps: [], note: ''
    };
    const wantPid = (args && args.pid) ? Number(args.pid) : 0;
    const wantProject = (args && args.project) || process.env.VB6MCP_AUTOSTART_PROJECT || '';
    const soft = await discoverSoft(LIST_GRACE_MS);
    const d = await resolveDiscovery(soft);
    const insts = (d.instances || []).filter(function (i) { return i.pid; });
    out.instances = insts.map(function (i) { return { pid: i.pid, port: i.port, project: i.project }; });
    let target = null;
    if (wantPid) {
      target = insts.filter(function (i) { return i.pid === wantPid; })[0] || null;
      out.targetSource = 'explicit_pid';
      if (!target) { out.refused_reason = 'pid_not_found'; out.note = 'pid ' + wantPid + ' is not a live instance; see instances[]'; return textResult(out, snap.contractVersion); }
    } else {
      if (!insts.length) { out.refused_reason = 'no_ide_running'; out.note = 'nothing to restart: no live instance was discovered'; return textResult(out, snap.contractVersion); }
      if (wantProject) {
        target = insts.filter(function (i) { return i.project && String(i.project).toLowerCase() === wantProject.toLowerCase(); })[0] || null;
        if (target) out.targetSource = 'project';
      }
      if (!target) {
        if (insts.length === 1) { target = insts[0]; out.targetSource = 'only_instance'; }
        else { out.refused_reason = 'ambiguous_target'; out.note = 'several instances online; pass pid or project - ' + out.instances.map(function (i) { return i.pid + '@' + i.port; }).join(', '); return textResult(out, snap.contractVersion); }
      }
    }
    if (Number(target.port) < 8767) {
      out.refused_reason = 'foreign_instance_refused';
      out.note = 'port ' + target.port + ' is below 8767, so this instance is not ours (8766 belongs to an external process on this machine). Refusing to stop it.';
      return textResult(out, snap.contractVersion);
    }
    out.ownership = startedByHub.has(target.pid) ? 'hub_started' : 'self_instance_port>=8767';
    out.stoppedPid = target.pid;
    out.project = target.project || wantProject;
    out.port = target.port;
    // Resolve VB6.EXE through the SHARED resolver above, so the
    // recovery tool works by default and names the source it used.
    const rExe = resolveVb6Exe();
    out.vb6ExeTried = rExe.tried;
    if (rExe.exe) { out.vb6Exe = rExe.exe; out.vb6ExeSource = rExe.source; }
    if (!out.vb6Exe) {
      out.refused_reason = 'vb6_exe_not_found';
      out.note = 'no VB6.EXE was resolvable (set VB6MCP_VB6EXE, or install VB6); tried: ' + rExe.tried.join(' | ');
      return textResult(out, snap.contractVersion);
    }
    // Choose the add-in's port pin BEFORE anything is stopped, so a
    // failure here cannot leave the IDE down. The launcher pins the port through the
    // VB6MCP_PORT environment variable (acceptance = reverse lookup), and 8766 is NEVER a
    // candidate (discipline 99/107).
    for (let pp = 8767; pp <= 8796; pp++) {
      if (pp === Number(target.port)) continue;
      let busyP = true;
      try { busyP = await tcpListening(pp, 120); } catch (e) { busyP = false; }
      if (!busyP) { out.pinPort = pp; break; }
    }
    if (!out.pinPort) {
      out.refused_reason = 'no_free_port_in_8767_8796';
      out.note = 'no free port was found in 8767..8796, so NOTHING was stopped and the IDE is untouched (an instance on 8766 is never ours).';
      return textResult(out, snap.contractVersion);
    }
    // spawn() must receive a real .vbp PATH, never a project NAME, and the
    // registry entry is written BEFORE the project is loaded - the first live run passed
    // the name "VB6MCP" and the new IDE therefore opened no project at all. So the path
    // is resolved (caller arg > registry projectPath > bounded re-read of the registry)
    // and a refusal names the exact parameter to pass.
    out.projectVbp = String(target.projectPath || '');
    if (!out.projectVbp && wantProject) out.projectVbp = String(wantProject);
    if (!out.projectVbp) {
      const tP = Date.now();
      while (Date.now() - tP < 30000) {
        await pause(1500);
        invalidateDiscovery();
        let dd2 = null;
        try { dd2 = await discoverSoft(LIST_GRACE_MS); } catch (e) { dd2 = null; }
        const again = (dd2 && dd2.instances ? dd2.instances : []).filter(function (i) { return i.pid === target.pid; })[0];
        if (again && again.projectPath) { out.projectVbp = String(again.projectPath); out.project = again.project || out.project; break; }
      }
    }
    if (!out.projectVbp) {
      out.refused_reason = 'project_path_unknown';
      out.note = 'the target instance did not publish its .vbp path (the registry entry is written before the project loads); pass project=<the absolute .vbp> so the IDE can be restarted on it. Nothing was stopped.';
      return textResult(out, snap.contractVersion);
    }
    out.project = out.projectVbp;
    // 1) SAVE best effort first: a graceful quit must not throw work away.
    out.saveAttempted = true;
    try {
      const sv = await rpc(target.port, 'tools/call', { name: 'save_project', arguments: { path: out.project } }, null, 30000);
      out.saveOk = !!(sv && sv.result && !sv.result.isError);
    } catch (e) { out.saveOk = false; }
    out.steps.push({ step: 'save_project', attempted: true, ok: out.saveOk });
    // 2) GRACEFUL quit through the add-in, then confirm the process is really gone.
    let quitAnswered = false;
    try { await rpc(target.port, 'tools/call', { name: 'quit_ide', arguments: {} }, null, 6000); quitAnswered = true; } catch (e) { quitAnswered = false; }
    out.steps.push({ step: 'quit_ide', requested: true, answered: quitAnswered });
    let gone = false;
    for (let w = 0; w < 48; w++) {
      await pause(250);
      try { process.kill(target.pid, 0); } catch (e) { gone = true; break; }
    }
    if (!gone) {
      let forced = false;
      try { process.kill(target.pid); forced = true; } catch (e) { forced = false; }
      out.steps.push({ step: 'force_kill', sent: forced });
      for (let w = 0; w < 20; w++) { await pause(250); try { process.kill(target.pid, 0); } catch (e) { gone = true; break; } }
    }
    out.stopConfirmed = gone;
    // 3) clear OUR OWN registry entry + the hub's bookkeeping.
    try {
      const rp = path.join(instancesDir(), target.pid + '.json');
      if (fs.existsSync(rp)) { fs.unlinkSync(rp); out.registryCleared = true; }
    } catch (e) { out.registryCleared = false; }
    startedByHub.delete(target.pid);
    invalidateDiscovery();
    out.lockCleared = false;
    // 4) restart on the SAME project.
    const childEnv = Object.assign({}, process.env, { VB6MCP_PORT: String(out.pinPort) });
    const child = spawn(out.vb6Exe, [out.project], { detached: true, stdio: 'ignore', env: childEnv });
    child.unref();
    out.newPid = child.pid;
    startedByHub.set(child.pid, { project: out.project, port: null });
    invalidateDiscovery();
    const t0 = Date.now();
    let newInst = null;
    let sawLowPortPid = 0;
    while (Date.now() - t0 < 120000) {
      await pause(1000);
      invalidateDiscovery();
      let dd = null;
      try { dd = await discoverSoft(LIST_GRACE_MS); } catch (e) { dd = null; }
      const found = (dd && dd.instances ? dd.instances : []).filter(function (i) { return i.pid === child.pid; })[0];
      if (found && found.port) {
        if (Number(found.port) < 8767) { sawLowPortPid = found.pid; continue; }
        // readiness, not just registration: the registry file is written BEFORE the
        // add-in is listening and before the project is loaded.
        let up = false;
        try { up = await tcpListening(Number(found.port), 200); } catch (e) { up = false; }
        if (!up) continue;
        newInst = found; break;
      }
    }
    out.waitMs = Date.now() - t0;
    if (!newInst && sawLowPortPid) {
      out.refused_reason = 'new_instance_bound_a_port_below_8767';
      out.note = 'the restarted IDE (pid ' + sawLowPortPid + ') came up on a port below 8767 (the VB6MCP_PORT pin was not honoured), which is NOT an acceptable instance; it is ours, so stop it with ide_stop pid=' + sawLowPortPid + ' before retrying.';
      return textResult(out, snap.contractVersion);
    }
    if (newInst) {
      out.port = newInst.port;
      out.threeWay.http = true;
      // The registry entry is written before the add-in serves and before the project is
      // loaded, so readiness is polled with a deadline instead of judged on first sight
      // (the first live run reported three_way_ok=false for exactly this race).
      let wantName = '';
      try { wantName = String(out.project).replace(/\\/g, '/').split('/').pop().replace(/[.]vbp$/i, '').toLowerCase(); } catch (e) { wantName = ''; }
      let actualProject = '';
      let mcpOk = false;
      const tReady = Date.now();
      while (Date.now() - tReady < 120000) {
        try {
          const mi = await rpc(newInst.port, 'tools/call', { name: 'mcp_info', arguments: {} }, null, 20000);
          mcpOk = !!(mi && mi.result && !mi.result.isError);
        } catch (e) { mcpOk = false; }
        try {
          const st = await rpc(newInst.port, 'tools/call', { name: 'ide_state', arguments: {} }, null, 20000);
          if (st && st.result && st.result.content && st.result.content[0]) {
            const sb = JSON.parse(st.result.content[0].text);
            actualProject = String(sb.project || '');
          }
        } catch (e) { actualProject = ''; }
        out.threeWay.mcp_info = mcpOk;
        out.threeWay.projectActual = actualProject;
        if (mcpOk && actualProject.length > 0) break;
        await pause(2000);
      }
      out.readyWaitMs = Date.now() - tReady;
      // The IDE reports the PROJECT NAME (the .vbp Name= field), which is NOT the file
      // name: compare against the target's published project name (live run: the .vbp is
      // MCP_VB6.vbp whose project name is "VB6MCP" - comparing file names was a false
      // negative in the first verdict).
      const wantProj = String(target.project || '').toLowerCase();
      out.threeWay.project = (actualProject.length > 0) && (wantProj.length === 0 || actualProject.toLowerCase() === wantProj);
      if (!out.threeWay.project && actualProject.length === 0 && newInst.project) {
        out.threeWay.project = (String(newInst.project).toLowerCase() === String(out.project).toLowerCase());
        out.threeWay.projectActual = String(newInst.project);
      }
    }
    out.three_way_ok = out.threeWay.http && out.threeWay.mcp_info && out.threeWay.project;
    out.restarted = !!(newInst && out.stopConfirmed);
    out.ok = out.restarted && out.three_way_ok;
    if (!out.ok && !out.refused_reason) { out.refused_reason = newInst ? 'three_way_not_ready' : 'new_instance_never_came_up'; }
    out.note = out.ok
      ? 'our own IDE (pid ' + out.stoppedPid + ') was stopped and restarted on ' + out.project + ' as pid ' + out.newPid + ' on port ' + out.port + '; the add-in answers, mcp_info answers and the project matches.'
      : 'not restarted: see refused_reason / steps / threeWay. Nothing foreign was touched and the 8766 instance was never a candidate.';
    return textResult(out, snap.contractVersion);
  }
  return textResult({ ok: false, reason: 'unknown_hub_tool', name: name }, snap.contractVersion);
}

async function toolsListResult(d) {
  d = await resolveDiscovery(d);
  const snap = snapshot();
  let faceTools = snap.tools;
  if (d && d.instances.length) {
    const inst = d.instances[0];
    if (inst.pending) {
      faceTools = snap.tools;                 // probe not finished: serve the aligned snapshot
    } else if (inst.liveToolDefs && inst.liveToolDefs.length) {
      faceTools = inst.liveToolDefs;          // served from the discovery cache: no RPC
    } else {
      try { const tl = await rpc(inst.port, 'tools/list', {}, null, PROBE_TIMEOUT_MS); faceTools = (tl.result && tl.result.tools) || snap.tools; }
      catch (e) { faceTools = snap.tools; }
    }
  }
  // Hub-provided names WIN: the add-in copy of a name the hub owns is DROPPED.  The old DLL
  // still serves the 5 lesson tools until the rebuild, so without this rule the face would
  // carry each of them twice (and a duplicated name breaks strict clients).
  const merged = [];
  const seenNames = {};
  for (const t of hubTools().concat(faceTools || [])) {
    if (!t || !t.name || seenNames[t.name]) continue;
    seenNames[t.name] = true;
    merged.push(t);
  }
  // Strict: EXACTLY {tools}.
  return { tools: merged };
}

// The refusal body keeps EVERY legacy key (ok/reason/message/instances/conflicts) in the same
// order and ADDS the candidate detail the path key needs (candidates[], identityConflict).
function refusalBody(pick, insts, conflicts) {
  const reason = pick.conflict ? 'target_conflict' : (pick.mismatch ? 'target_mismatch' : (pick.ambiguous ? 'ambiguous_target' : 'no_ide_running'));
  const body = {
    ok: false, reason: reason, message: pick.message,
    instances: (insts || []).map(function (i) { return { pid: i.pid, port: i.port, project: i.project }; }),
    conflicts: conflicts
  };
  // candidates[] is emitted when it carries the identity information this change adds: a
  // multi-match refusal (the same path on one endpoint, or one NAME over several paths) and
  // ANY refusal once remote endpoints exist.  With remotes EMPTY and no multi-match, a refusal
  // stays byte-identical to the pre-change hub (proven in re_work/t1103).
  const showCandidates = REMOTE_ENDPOINTS.length > 0 || !!pick.conflict;
  if (showCandidates) {
    if (pick.identityConflict !== undefined) body.identityConflict = !!pick.identityConflict;
    if (pick.distinctIdentities !== undefined) body.distinctIdentities = pick.distinctIdentities;
    if (pick.matchedCount !== undefined) body.matchedCount = pick.matchedCount;
    body.candidates = (pick.hits && pick.hits.length) ? pick.hits : (insts || []).map(candidateBrief);
  }
  if (REMOTE_ENDPOINTS.length) { body.endpoints = endpointLabels(); body.targetKeyRule = TARGET_KEY_RULE; }
  return body;
}
// P1b (author ruling 2026-10-14: only an endpoint-qualified key may address a remote IDE).
// An UNQUALIFIED call is refused here - it is routed NOWHERE: not to the remote IDE (that silent
// machine jump is exactly what the ruling forbids) and not to a local IDE (there is none live).
// The body stays ACTIONABLE: it names the endpoint-qualified keys the caller may use verbatim.
function remoteRequiresEndpointBody(remoteInsts, conflicts) {
  const keys = [];
  for (const i of remoteInsts) { const k = i.targetKey; if (k && keys.indexOf(k) < 0) keys.push(k); }
  const labels = [];
  for (const i of remoteInsts) { const l = i.endpoint || 'remote'; if (labels.indexOf(l) < 0) labels.push(l); }
  const example = keys.length ? keys[0] : (labels.length ? labels[0] : 'remote');
  return {
    ok: false,
    reason: 'remote_requires_explicit_endpoint',
    message: 'no LOCAL IDE is live; the only live IDE(s) are on remote endpoint(s) ' + labels.join(', ')
      + ' -- an UNQUALIFIED call is never auto-routed to a remote IDE, so this call was routed NOWHERE',
    qualifiedKeys: keys,
    remoteEndpoints: labels,
    hint: 'pass _target="' + example + '" (or set ide_target{target:"' + example + '"}) to address that remote IDE explicitly; '
      + 'a live LOCAL IDE is still selected automatically, and a bare pid / bare project NAME only ever addresses THIS machine',
    instances: (remoteInsts || []).map(function (i) { return { pid: i.pid, port: i.port, project: i.project }; }),
    conflicts: conflicts || [],
    candidates: (remoteInsts || []).map(candidateBrief),
    routed: false,
    endpoints: endpointLabels(),
    targetKeyRule: TARGET_KEY_RULE
  };
}
// The ONLY forward path: local goes to 127.0.0.1:<port> over the same wire the hub always used;
// remote goes to the endpoint URL.  Per-endpoint timeout + per-endpoint probe single-flight
// means a slow remote can never stall a local call.
async function forwardPicked(pick, name, args, snap, explicit, rawRequest) {
  // r-c (PROTOCOL SYMMETRY): a lesson method must NEVER be served by an endpoint.  Normal
  // dispatch serves them at the control plane before routing is considered, so reaching this
  // point is a protocol violation - fail STRUCTURALLY and never create a per-machine copy of
  // the store (that is the rule that prevents knowledge divergence between machines).
  if (lessonCliArgs(name, args)) {
    return stampBody({ content: [{ type: 'text', text: JSON.stringify({
      ok: false, reason: 'global_state_is_control_plane_only', tool: name,
      endpoint: pick.endpoint || (pick.remote ? 'remote' : LOCAL_ENDPOINT), remote: !!pick.remote,
      message: 'lesson tools read/write the SINGLE control-plane store (mcp_lessons/); an endpoint must never serve them and must never create a local copy of that store'
    }, null, 2) }], isError: true }, null, snap.contractVersion);
  }
  const pickNote = {};
  if (pick.selectedBy) pickNote.selectedBy = pick.selectedBy;
  if (pick.staleTarget) {
    pickNote.staleTarget = pick.staleTarget;
    pickNote.staleTargetIgnored = true;
    pickNote.targetNote = 'the default ide_target "' + pick.staleTarget + '" points at no live instance (treated as unset, NOT rewritten); selected the only live instance ' + pick.port + '/' + (pick.project || '?') + ' -- clear it with ide_target{auto} if it is stale';
  }
  const isRemote = !!pick.remote;
  if (isRemote) { pickNote.remote = true; pickNote.endpoint = pick.endpoint; pickNote.endpointUrl = pick.endpointUrl; }
  if (pick.matchedBy === 'project_path') {
    pickNote.targetKey = pick.targetKey;
    pickNote.targetNormalized = { from: pick.matchedFrom || rawRequest || null, to: pick.projectPath, endpoint: pick.endpoint || LOCAL_ENDPOINT, via: pick.pathVia || null, rule: TARGET_KEY_RULE };
  }
  const ideStamp = { pid: pick.pid, project: pick.project, port: pick.port };
  if (isRemote) { ideStamp.endpoint = pick.endpoint; ideStamp.remote = true; }
  const timeoutMs = isRemote ? REMOTE_FORWARD_TIMEOUT_MS : FORWARD_TIMEOUT_MS;
  const base = isRemote ? pick.endpointUrl : ('http://127.0.0.1:' + pick.port);
  const tFwd = Date.now();
  let resp;
  try { resp = await rpcAt(base, 'tools/call', { name: name, arguments: args }, null, timeoutMs); }
  catch (e) {
    const ms = Date.now() - tFwd;
    const isTimeout = !!(e && (e.name === 'TimeoutError' || /abort/i.test(String(e && e.message))));
    const errBody = { ok: false, reason: isTimeout ? 'forward_timeout' : 'forward_failed', message: String(e && e.message || e), elapsedMs: ms, budgetMs: timeoutMs };
    if (isRemote) { errBody.endpoint = pick.endpoint; errBody.endpointUrl = pick.endpointUrl; errBody.remote = true; }
    else if (isTimeout) { errBody.hint = 'raise VB6MCP_FORWARD_TIMEOUT_MS for legitimately long calls'; }
    return stampBody({ content: [{ type: 'text', text: JSON.stringify(errBody) }], isError: true }, ideStamp, snap.contractVersion, pickNote);
  }
  dbg('forward phases: rpc=' + (Date.now() - tFwd) + 'ms tool=' + name + ' ' + (isRemote ? 'endpoint=' + pick.endpoint : 'port=' + pick.port));
  const result = (resp && resp.result) || { content: [{ type: 'text', text: JSON.stringify(resp) }] };
  return stampBody(result, ideStamp, snap.contractVersion, pickNote);
}

async function handleCall(params) {
  const name = params && params.name;
  const argsIn = (params && params.arguments) || {};
  const snap = snapshot();
  // The lesson tools are declared in hub_tools.json but SERVED by the Python module (the
  // add-in dropped them), so dispatch them before the generic hub-name branch.
  // LESSONS ARE CONTROL-PLANE STATE (r-a/r-b): _target and ide_target have NO effect on them.
  // A supplied target is ECHOED back so a caller is never left believing the write went to a
  // remote machine; the call is always served HERE from the ONE local store.
  // (r-d, deliberately NOT implemented yet: the ignored target belongs in PROVENANCE, e.g.
  //  observed_on:"<label>"/<projectPath>, so machine-specific experience is not lost while the
  //  store stays single.  That must not become a new inputSchema parameter in this batch.)
  if (lessonCliArgs(name, argsIn)) {
    const rawT = Object.prototype.hasOwnProperty.call(argsIn, '_target') ? argsIn._target : undefined;
    const ignoredT = (rawT === undefined || rawT === null || String(rawT).trim() === '' || String(rawT).trim().toLowerCase() === 'auto') ? null : String(rawT).trim();
    const defT = (state.target === undefined || state.target === null || state.target === '') ? null : String(state.target);
    const echo = { globalState: true, store: 'control-plane (mcp_lessons/)', targetIgnored: ignoredT, defaultTargetIgnored: ignoredT ? null : defT };
    if (ignoredT) {
      echo.note = 'lessons are global control-plane state (ONE store, mcp_lessons/); the supplied _target "' + ignoredT + '" was IGNORED - the call was served locally and nothing was written to any other machine';
    } else if (echo.defaultTargetIgnored) {
      echo.note = 'lessons are global control-plane state (ONE store, mcp_lessons/); the default ide_target "' + echo.defaultTargetIgnored + '" does not apply to them';
    } else {
      echo.note = 'lessons are global control-plane state (ONE store, mcp_lessons/); endpoint targeting never applies to them';
    }
    return lessonsTool(name, argsIn, snap.contractVersion, echo);
  }
  if (hubToolNames().indexOf(name) >= 0) return hubTool(name, argsIn);
  const args = Object.assign({}, argsIn);
  const hasExplicit = Object.prototype.hasOwnProperty.call(argsIn, '_target') && argsIn._target !== undefined && argsIn._target !== null;
  delete args._target;
  // Effective request: EXPLICIT _target > ide_target default > auto.  The DEFAULT stays
  // stale-tolerant (a vanished target never vetoes the live set); an explicit target is never
  // silently redirected to another endpoint.
  const defaultTarget = (state.target === undefined || state.target === null || state.target === '') ? 'auto' : String(state.target);
  const rawRequest = hasExplicit ? String(argsIn._target) : defaultTarget;
  const tgt = parseTarget(rawRequest);
  const explicit = hasExplicit && tgt.kind !== 'auto';
  const tDisc = Date.now();

  if (tgt.kind === 'endpoint_unknown') {
    return stampBody({ content: [{ type: 'text', text: JSON.stringify({
      ok: false, reason: 'endpoint_unknown', endpoint: tgt.label, endpoints: endpointLabels(),
      message: 'endpoint "' + tgt.label + '" is not configured; configured endpoints: ' + endpointLabels().join(', ')
    }, null, 2) }], isError: false }, null, snap.contractVersion);
  }

  // ---- endpoint-qualified REMOTE target: probe THAT endpoint; never fall back to local ------
  if (tgt.kind === 'endpoint' && tgt.endpoint.remote) {
    const t0 = Date.now();
    const r = await probeEndpointCached(tgt.endpoint.label, REMOTE_PROBE_TIMEOUT_MS);
    if (!r || !r.ok) {
      const dd = r || {};
      return stampBody({ content: [{ type: 'text', text: JSON.stringify({
        ok: false, endpoint: tgt.endpoint.label, endpointUrl: tgt.endpoint.url, remote: true,
        reason: dd.reason === 'endpoint_timeout' ? 'endpoint_timeout' : 'endpoint_unreachable',
        message: dd.message || 'the endpoint did not answer',
        elapsedMs: dd.elapsedMs === undefined ? (Date.now() - t0) : dd.elapsedMs,
        probeTimeoutMs: dd.timeoutMs || REMOTE_PROBE_TIMEOUT_MS,
        hint: 'a remote endpoint is only used when it is explicitly named; a local call is never redirected to it'
      }, null, 2) }], isError: true }, null, snap.contractVersion);
    }
    const pickR = pickTarget([r.candidate], tgt.selector || null, { remoteScope: true, explicit: explicit });
    if (!pickR.ok) return stampBody({ content: [{ type: 'text', text: JSON.stringify(refusalBody(pickR, [r.candidate], []), null, 2) }], isError: false }, null, snap.contractVersion);
    return forwardPicked(pickR, name, args, snap, explicit, rawRequest);
  }

  // ---- local / unqualified -------------------------------------------------------------------
  const localD = await resolveDiscovery(await discoverSoft(LIST_GRACE_MS));
  const tRes = Date.now();
  let merged = null;
  let pick;
  if (tgt.kind === 'endpoint') {                       // "local:<selector>"
    pick = pickTarget(localD.instances, tgt.selector || null, { remoteScope: false, explicit: explicit });
  } else if (tgt.kind === 'local') {                   // bare selector (path | pid | port | name)
    pick = pickTarget(localD.instances, tgt.selector, { remoteScope: false, explicit: explicit });
  } else if (localD.instances.length) {                // auto, LOCAL FIRST
    pick = pickTarget(localD.instances, null, { remoteScope: false, explicit: false });
  } else if (REMOTE_ENDPOINTS.length) {                // auto with ZERO local IDEs: LOCAL ONLY, never remote
    // P1b (author ruling): an UNQUALIFIED call may only ever pick a LOCAL IDE.  The pre-P1b rule
    // 6r ("else the ONLY live IDE overall") auto-targeted a lone REMOTE IDE - that silent machine
    // jump is what this refuses.  The remote face is still PROBED here, but only so the refusal
    // can name the endpoint-qualified keys the caller must use; nothing is routed anywhere.
    merged = await withRemoteCandidates(localD, REMOTE_GRACE_MS);
    const remoteLive = merged.instances.filter(function (i) { return !!i.remote; });
    if (remoteLive.length) {
      dbg('handleCall: REFUSED unqualified call - 0 local IDEs, ' + remoteLive.length + ' live remote IDE(s); an endpoint-qualified key is required');
      return stampBody({ content: [{ type: 'text', text: JSON.stringify(remoteRequiresEndpointBody(remoteLive, merged.conflicts), null, 2) }], isError: false }, null, snap.contractVersion);
    }
    pick = pickTarget(merged.instances, null, { remoteScope: false, explicit: false });
  } else {
    pick = pickTarget(localD.instances, null, { remoteScope: false, explicit: false });
  }
  const tPick = Date.now();
  dbg('forward phases: discoverSoft=' + (tRes - tDisc) + 'ms resolveDiscovery=' + (tPick - tRes) + 'ms instances=' + localD.instances.length + (merged ? ' merged=' + merged.instances.length : ''));
  // The direct-target escape hatch ("discovery is blind but the add-in is reachable") is NOT
  // allowed to rescue a CONFLICT: a selector that matches several live IDEs on one endpoint
  // must be reported, never silently resolved to the first registry hit (that is the exact
  // silent pick the path key exists to prevent).
  if (!pick.ok && !pick.conflict) {
    const want = explicit ? (tgt.selector || null) : (state.target || null);
    const dt = want ? await directTarget(want) : null;
    if (dt) { dbg('handleCall: directTarget bypass used for ' + want + ' -> port ' + dt.port); pick = dt; }
  }
  if (!pick.ok) {
    const insts = (merged ? merged.instances : localD.instances);
    const cf = (merged ? merged.conflicts : localD.conflicts);
    return stampBody({ content: [{ type: 'text', text: JSON.stringify(refusalBody(pick, insts, cf), null, 2) }], isError: false }, null, snap.contractVersion);
  }
  return forwardPicked(pick, name, args, snap, explicit, rawRequest);
}

function send(obj) { process.stdout.write(JSON.stringify(obj) + '\n'); }

async function handle(line) {
  let req;
  try { req = JSON.parse(line); } catch (e) {
    // A request we cannot parse gets an explicit JSON-RPC parse error.
    // Returning silently would hang any client that has no read timeout.
    let bad = null;
    const m = /"id"\s*:\s*("?[^",}\s]+"?)/.exec(line);
    if (m) { const raw = m[1].replace(/^"|"$/g, ''); bad = /^\d+$/.test(raw) ? parseInt(raw, 10) : raw; }
    return send({ jsonrpc: '2.0', id: bad, error: { code: -32700, message: 'parse_error',
      data: { bytes: line.length, hint: 'invalid JSON (and every request must end with a newline)' } } });
  }
  const id = req.id; const method = req.method;
  if (method === 'initialize') {
    const v = PROTOCOLS.indexOf(req.params && req.params.protocolVersion) >= 0 ? req.params.protocolVersion : LATEST;
    const snap = snapshot();
    const d = await discoverSoft(DISCOVERY_GRACE_MS);
    const st = await faceState(d);
    return send({ jsonrpc: '2.0', id: id, result: {
      protocolVersion: v,
      capabilities: { tools: {} },
      serverInfo: SERVER,
      instructions: await buildInstructions(await resolveDiscovery(d), snap),
      faceSource: st.faceSource,
      contractVersion: st.contractVersion,
      toolCount: st.toolCount
    } });
  }
  if (method && method.indexOf('notifications/') === 0) return;
  if (method === 'tools/list') return send({ jsonrpc: '2.0', id: id, result: await toolsListResult(await discoverSoft(LIST_GRACE_MS)) });
  if (method === 'tools/call') return send({ jsonrpc: '2.0', id: id, result: await handleCall(req.params) });
  if (method === 'ping') return send({ jsonrpc: '2.0', id: id, result: {} });
  return send({ jsonrpc: '2.0', id: id, error: { code: -32601, message: 'method not found: ' + method } });
}

// Warm the discovery cache as soon as the process starts (background; never blocks a reply).
setTimeout(function () { discover().catch(function () { }); }, 0);
// Warm the lessons cache too, so the FIRST initialize does not pay the python spawn.
setTimeout(function () { lessonsText().catch(function () { }); }, 0);

let buf = '';
let inFlight = 0;
let stdinEnded = false;
let idleTimer = null;
function maybeExit() {
  // Never process.exit() here - that truncates responses that have not been
  // flushed yet (a finite pipe got 0 bytes).  Setting exitCode lets the loop drain stdout,
  // and we only arm it once every in-flight request has answered.
  if (stdinEnded && inFlight === 0) process.exitCode = 0;
}
function dispatch(line) {
  inFlight++;
  handle(line)
    .catch(function (e) { send({ jsonrpc: '2.0', id: null, error: { code: -32603, message: String(e && e.message || e) } }); })
    .then(function () { inFlight--; maybeExit(); });
}
function consumeLines() {
  let i;
  while ((i = buf.indexOf('\n')) >= 0) {
    const line = buf.slice(0, i).replace(/\r$/, '');
    buf = buf.slice(i + 1);
    if (line.trim()) dispatch(line);
  }
}
// Idle fallback: a client that never appends '\n' must not hang forever.  If
// the buffer holds COMPLETE JSON and nothing new arrives for IDLE_FLUSH_MS, treat it as one
// line.  The timer is reset by every data chunk and unref'd so it never holds the process.
function armIdleFlush() {
  if (idleTimer) { clearTimeout(idleTimer); idleTimer = null; }
  if (stdinEnded || !buf.trim()) return;
  idleTimer = setTimeout(function () {
    idleTimer = null;
    const rem = buf.trim();
    if (!rem) return;
    try { JSON.parse(rem); } catch (e) { return; }   // incomplete: keep waiting for more
    buf = '';
    dispatch(rem);
    consumeLines();
    armIdleFlush();
  }, IDLE_FLUSH_MS);
  if (idleTimer.unref) idleTimer.unref();
}
// When stdin ends, whatever is left is HANDLED - parsed as a request, or
// answered with an explicit JSON-RPC parse error.  Silence is what hangs a client forever.
function flushResidual(reason) {
  const rem = buf.trim();
  buf = '';
  if (!rem) return;
  let id = null;
  const m = /"id"\s*:\s*("?[^",}\s]+"?)/.exec(rem);
  if (m) { const raw = m[1].replace(/^"|"$/g, ''); id = /^\d+$/.test(raw) ? parseInt(raw, 10) : raw; }
  try {
    JSON.parse(rem);
  } catch (e) {
    send({ jsonrpc: '2.0', id: id, error: { code: -32700, message: 'unterminated_request',
      data: { reason: reason, bytes: rem.length, hint: 'MCP stdio requests must end with a newline (\\n)' } } });
    return;
  }
  dispatch(rem);
}
process.stdin.setEncoding('utf8');
process.stdin.on('data', function (chunk) {
  buf += chunk;
  consumeLines();
  armIdleFlush();
});
process.stdin.on('end', function () {
  stdinEnded = true;
  if (idleTimer) { clearTimeout(idleTimer); idleTimer = null; }
  flushResidual('stdin_ended');
  maybeExit();
});
