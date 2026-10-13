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
import { execFileSync, spawn } from 'node:child_process';
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

async function rpc(port, method, params, proto, timeoutMs, rpcPath) {
  const body = { jsonrpc: '2.0', id: (Date.now() % 100000) + 1, method: method };
  if (params !== undefined) body.params = params;
  const headers = { 'Content-Type': 'application/json', 'Accept': 'application/json, text/event-stream' };
  if (proto) headers['MCP-Protocol-Version'] = proto;
  const res = await fetch('http://127.0.0.1:' + port + (rpcPath || ADDIN_RPC_PATH), {
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
  const conflicts = [];
  insts.forEach(function (i) { if (i.pids.length > 1) conflicts.push('port ' + i.port + ' held by several processes: ' + i.pids.join(',')); });
  const byProject = {};
  insts.forEach(function (i) { if (i.project) { (byProject[i.project] = byProject[i.project] || []).push(i.port); } });
  Object.keys(byProject).forEach(function (p) {
    if (byProject[p].length > 1) conflicts.push('project "' + p + '" opened by several instances on ports ' + byProject[p].join(','));
  });
  stale.forEach(function (s) {
    if (s.reason === 'duplicate_port_in_registry' || s.reason === 'registry_port_not_mcp') conflicts.push('registry ' + s.file + ': ' + s.reason);
  });
  return { instances: insts, conflicts: conflicts, staleRegistry: stale, registryDir: scan.dir,
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

function pickTarget(insts, target) {
  if (!insts.length) return { ok: false, message: 'no IDE instance listening (scanned ' + DEFAULT_PORTS.length + ' default ports + registry)' };
  const explicit = !(target === 'auto' || target === undefined || target === null || target === '');
  let t = explicit ? target : state.target;
  if (t === 'auto' || t === undefined || t === null || t === '') t = null;
  const desc = function () { return insts.map(function (i) { return i.port + '/' + (i.project || '?'); }).join(', '); };
  if (t !== null) {
    const s = String(t);
    const hit = insts.filter(function (i) { return String(i.pid) === s || (i.project && i.project === s) || String(i.port) === s; });
    if (hit.length === 1) return { ok: true, ...hit[0], selectedBy: (explicit ? 'explicit_target' : 'effective_default') };
    if (hit.length > 1) return { ok: false, conflict: true, message: 'target "' + s + '" matches several instances: ' + hit.map(function (i) { return i.port + '/' + i.pid; }).join(',') };
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

function buildInstructions(d, snap) {
  const inst = d.instances.length ? d.instances[0] : null;
  if (inst && inst.instructions) return inst.instructions;
  if (inst && snap.instructions) {
    const n = (inst.liveTools && inst.liveTools.length) || snap.tools.length;
    return 'PRODUCT vb6mcp-hub ' + SERVER.version + '  BRIDGE-TO vb6-mcp ' + snap.contractVersion + '  TOOLS ' + n + '  (faceSource=live; add-in sent no instructions)' + (inst.pending ? '  [probe pending]' : '') + '\n' + snap.instructions;
  }
  const mm = snap.versionToolCountMismatch ? '  MISMATCH contract-declares=' + snap.contractDeclaredToolCount + ' snapshot=' + snap.tools.length : '';
  const head = 'PRODUCT vb6mcp-hub ' + SERVER.version + '  BRIDGE-TO vb6-mcp ' + snap.contractVersion + '  TOOLS ' + snap.tools.length + '  (faceSource=snapshot; no IDE online)' + mm;
  const brk = (snap.breaking && snap.breaking.length) ? snap.breaking : ['- (snapshot has no BREAKING extract; read docs/CONTRACTS.md)'];
  return [head, 'BREAKING (read before calling):'].concat(brk).join('\n');
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
    if (okp) checked.push(i);
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
  const out = s.live;
  return { instances: out, conflicts: [], staleRegistry: s.stale, registryDir: s.dir,
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
  // on every request, before it is reported.
  if (d.instances && d.instances.length) {
    const alive = [];
    for (const i of d.instances) {
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
      version: i.version || null, dllSha256: i.dllSha256 || null, pending: !!i.pending
    }; }),
    conflicts: d.conflicts,
    // Observability: WHY this face looks the way it does - how many registry
    // entries were read, how many were credible, and the stale reasons by kind.
    registrySummary: {
      entries: (d.instances || []).length + (d.staleRegistry || []).length,
      live: (d.instances || []).length,
      stale: (d.staleRegistry || []).length,
      staleByReason: (d.staleRegistry || []).reduce(function (m, x) { const k = x.reason || 'unknown'; m[k] = (m[k] || 0) + 1; return m; }, {})
    },
    staleRegistry: d.staleRegistry || [],
    registryDir: d.registryDir || null,
    registryUnreadable: !!d.registryUnreadable,
    selectionRule: '_target (EXPLICIT) > ide_target default > the ONLY live instance; a STALE default is ignored (never vetoes); otherwise refuse with ambiguous_target / target_mismatch (never guess)',
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

async function hubState() { return faceState(await discoverSoft(LIST_GRACE_MS)); }

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
  // Strict: EXACTLY {tools}.
  return { tools: hubTools().concat(faceTools) };
}

async function handleCall(params) {
  const name = params && params.name;
  const argsIn = (params && params.arguments) || {};
  const snap = snapshot();
  if (hubToolNames().indexOf(name) >= 0) return hubTool(name, argsIn);
  const args = Object.assign({}, argsIn);
  const explicit = Object.prototype.hasOwnProperty.call(args, '_target') ? args._target : 'auto';
  delete args._target;
  const tDisc = Date.now();
  const soft = await discoverSoft(LIST_GRACE_MS);
  const tRes = Date.now();
  const d = await resolveDiscovery(soft);
  const tPick = Date.now();
  dbg('forward phases: discoverSoft=' + (tRes - tDisc) + 'ms resolveDiscovery=' + (tPick - tRes) + 'ms instances=' + d.instances.length);
  let pick = pickTarget(d.instances, explicit);
  if (!pick.ok) {
    const want = explicit === 'auto' ? state.target : explicit;
    const dt = want ? await directTarget(want) : null;
    if (dt) { dbg('handleCall: directTarget bypass used for ' + want + ' -> port ' + dt.port); pick = dt; }
  }
  if (!pick.ok) {
    const reason = pick.conflict ? 'target_conflict' : (pick.mismatch ? 'target_mismatch' : (pick.ambiguous ? 'ambiguous_target' : 'no_ide_running'));
    return stampBody({ content: [{ type: 'text', text: JSON.stringify({
      ok: false, reason: reason, message: pick.message,
      instances: d.instances.map(function (i) { return { pid: i.pid, port: i.port, project: i.project }; }),
      conflicts: d.conflicts
    }, null, 2) }], isError: false }, null, snap.contractVersion);
  }
  const pickNote = {};
  if (pick.selectedBy) pickNote.selectedBy = pick.selectedBy;
  if (pick.staleTarget) {
    pickNote.staleTarget = pick.staleTarget;
    pickNote.staleTargetIgnored = true;
    pickNote.targetNote = 'the default ide_target "' + pick.staleTarget + '" points at no live instance (treated as unset, NOT rewritten); selected the only live instance ' + pick.port + '/' + (pick.project || '?') + ' -- clear it with ide_target{auto} if it is stale';
  }
  const tFwd = Date.now();
  let resp;
  try { resp = await rpc(pick.port, 'tools/call', { name: name, arguments: args }, null); }
  catch (e) {
    const ms = Date.now() - tFwd;
    const isTimeout = !!(e && (e.name === 'TimeoutError' || /abort/i.test(String(e && e.message))));
    return stampBody({ content: [{ type: 'text', text: JSON.stringify({
      ok: false,
      reason: isTimeout ? 'forward_timeout' : 'forward_failed',
      message: String(e && e.message || e),
      elapsedMs: ms,
      budgetMs: FORWARD_TIMEOUT_MS,
      hint: isTimeout ? 'raise VB6MCP_FORWARD_TIMEOUT_MS for legitimately long calls' : undefined
    }) }], isError: true }, { pid: pick.pid, project: pick.project, port: pick.port }, snap.contractVersion, pickNote);
  }
  dbg('forward phases: rpc=' + (Date.now() - tFwd) + 'ms tool=' + name + ' port=' + pick.port);
  const result = (resp && resp.result) || { content: [{ type: 'text', text: JSON.stringify(resp) }] };
  return stampBody(result, { pid: pick.pid, project: pick.project, port: pick.port }, snap.contractVersion, pickNote);
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
      instructions: buildInstructions(await resolveDiscovery(d), snap),
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
