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
import crypto from 'node:crypto';

const HERE = path.dirname(fileURLToPath(import.meta.url));
const SNAP = path.join(HERE, 'hub_schema.generated.json');
const HUBCFG = path.join(HERE, 'hub_tools.json');
const PROTOCOLS = ['2025-11-25', '2025-06-18', '2025-03-26', '2024-11-05', '2024-10-07'];
const LATEST = PROTOCOLS[0];
const SERVER = { name: 'vb6mcp-hub', version: '0.1.0' };
// ---- thin shell: the Python middle service layer owns the ONE local TCP port ----------------
// Author ruling (ADR sec 0.1 three layers): the mjs side keeps the MCP wire protocol, the
// guardians and the degradation; discovery / addressing / named-pipe multiplexing move to
// bridge/vb6mcp_bridge.py, which holds exactly ONE port for every IDE on the machine.
//
//   VB6MCP_BRIDGE=python (default) -> ask the bridge for the add-in face
//   VB6MCP_BRIDGE=inproc           -> the PRE-CHANGE path, byte for byte: local discovery
//                                     (netstat + port scan) and a direct rpc() to the add-in.
//                                     EVERY bridge branch below is guarded by this mode, so
//                                     nothing new is reached and nothing old is bypassed.
const BRIDGE_MODE = (process.env.VB6MCP_BRIDGE || 'python').toLowerCase();
// ---- ZERO-CONFIG LOCAL ADDRESSING (zero-config addressing) -----------------------------------------------
// Priority: VB6MCP_BRIDGE_URL (explicit) > zero-config discovery of the LOCAL middle service >
// the legacy default.  Discovery is SELF-LIMITED by construction:
//   * loopback only: every candidate is built as http://127.0.0.1:<port> (a non-loopback endpoint
//     can only ever be reached by an EXPLICIT VB6MCP_BRIDGE_URL, never by discovery);
//   * identity before credentials: phase A sends NO Authorization header at all and accepts the
//     endpoint only when it PROVES it is ours - a 200 /health whose "server" field is
//     "vb6mcp-bridge", or a 401/403 whose HTTP "Server" header is "vb6mcp-bridge/<version>".
//     The configured token is sent ONLY to an endpoint that already passed that test;
//   * it never opens a listening port: discovery CONNECTS to ports, and when it finds nothing the
//     hub answers with a structured no_middle_service receipt instead of spawning a service.
//   * VB6MCP_BRIDGE_DISCOVERY=0 restores the old single-URL path byte-for-byte (rollback switch).
const BRIDGE_URL_ENV_RAW = String(process.env.VB6MCP_BRIDGE_URL || '').trim();
const BRIDGE_URL_EXPLICIT = !!BRIDGE_URL_ENV_RAW;
const BRIDGE_URL_LEGACY_DEFAULT = 'http://127.0.0.1:8767';
const BRIDGE_DISCOVERY_DISABLED = String(process.env.VB6MCP_BRIDGE_DISCOVERY || '').trim() === '0';
const BRIDGE_DISCOVERY_TIMEOUT_MS = parseInt(process.env.VB6MCP_BRIDGE_DISCOVERY_TIMEOUT_MS || '350', 10);
const BRIDGE_DISCOVERY_AUTH_TIMEOUT_MS = parseInt(process.env.VB6MCP_BRIDGE_DISCOVERY_AUTH_TIMEOUT_MS || '700', 10);
const BRIDGE_DISCOVERY_BUDGET_MS = parseInt(process.env.VB6MCP_BRIDGE_DISCOVERY_BUDGET_MS || '2500', 10);
const BRIDGE_DISCOVERY_RULE = 'zero-config bridge discovery (zero-config addressing): when VB6MCP_BRIDGE_URL is NOT set, the hub probes LOCAL LOOPBACK candidates only (the documented 8797 convention, ports recorded in %TEMP%\\vb6mcp\\middle.*.lock/.state.json and instance registrations, then 8767..8796) and adopts the FIRST endpoint that PROVES it is the vb6mcp middle service. Proof = a 200 /health body whose server/product field is exactly "vb6mcp-bridge", or a 401/403 whose Server header starts with "vb6mcp-bridge/" (the credential is only sent AFTER that proof). Nothing found => a structured no_middle_service receipt with a hint; discovery never spawns a listener. VB6MCP_BRIDGE_URL always wins; VB6MCP_BRIDGE_DISCOVERY=0 disables discovery and restores the legacy default.';
let BRIDGE_URL = (BRIDGE_URL_EXPLICIT ? BRIDGE_URL_ENV_RAW : BRIDGE_URL_LEGACY_DEFAULT).replace(/\/+$/, '');
let BRIDGE_DISCOVERY = {
  state: BRIDGE_URL_EXPLICIT ? 'env' : (BRIDGE_MODE === 'inproc' ? 'skipped_inproc'
         : (BRIDGE_DISCOVERY_DISABLED ? 'disabled' : 'pending')),
  source: BRIDGE_URL_EXPLICIT ? 'VB6MCP_BRIDGE_URL' : (BRIDGE_DISCOVERY_DISABLED ? 'legacy_default' : 'pending'),
  url: BRIDGE_URL, port: null, candidates: [], probes: [], probedCount: 0, elapsedMs: 0,
  evidence: null, reason: null, envOverride: BRIDGE_URL_EXPLICIT,
  legacyDefault: !BRIDGE_URL_EXPLICIT && BRIDGE_DISCOVERY_DISABLED,
  authState: 'off'
};
const BRIDGE_HEALTH_TTL_MS = parseInt(process.env.VB6MCP_BRIDGE_HEALTH_TTL_MS || '5000', 10);
const BRIDGE_LIST_TIMEOUT_MS = parseInt(process.env.VB6MCP_BRIDGE_LIST_TIMEOUT_MS || '5000', 10);
// A forwarded call may be LONG (build / make_project): the shell must never be TIGHTER than the
// pre-change budget, so its bridge read timeout is FORWARD_TIMEOUT_MS plus a margin.  It is a
// FUNCTION on purpose: FORWARD_TIMEOUT_MS is declared below this block, so reading it here at
// module-evaluation time would be a temporal-dead-zone error.
function bridgeCallTimeoutMs() {
  return parseInt(process.env.VB6MCP_BRIDGE_CALL_TIMEOUT_MS || String(FORWARD_TIMEOUT_MS + 10000), 10);
}
// (WP-22-01 / t16 g1b ruling) THE FORWARD BUDGET IS VISIBLE, NOT INVISIBLE.  An implicit target
// whose instance registered but has NOT answered a probe is still forwarded: a wedged IDE must
// stay reachable so ide_state{release_inflight:true} can recover it (recovery semantics beat
// first-call latency).  That means one call can legitimately occupy the hub for the FULL forward
// budget, so every forwarded receipt names it: forwardBudget{budgetMs,plusSlackMs,env{name,set,value},rule}.
const FORWARD_BUDGET_RULE = 'one forwarded tools/call is given budgetMs = FORWARD_TIMEOUT_MS (plus plusSlackMs for the bridge read in python mode; env overrides are named in env{name,set,value}); an implicit (no _target) call is still forwarded when its instance registered but has not answered the discovery probe (a wedged IDE must remain reachable so ide_state{release_inflight:true} can recover it), so budgetMs is the WORST-CASE wait for such a call. For a deterministic latency pass _target explicitly, or probe with ide_state first.';
const IMPLICIT_FORWARD_TIMEOUT_MS = Math.max(1000, parseInt(process.env.VB6MCP_IMPLICIT_FORWARD_TIMEOUT_MS || '15000', 10) || 15000);
// The cap applies ONLY to a call whose target was NOT given explicitly AND whose instance has
// not answered a probe (or was selected by the ladder's empty-face fallback): those are the calls
// the t16/caller timeout case is made of.  Explicit _target keeps the full budget.
function implicitForwardCapMs() { return IMPLICIT_FORWARD_TIMEOUT_MS; }
function forwardBudgetBlock(implicit, effectiveMs, decidedBy) {
  const bridgeMode = BRIDGE_MODE !== 'inproc';
  const callOv = process.env.VB6MCP_BRIDGE_CALL_TIMEOUT_MS !== undefined && String(process.env.VB6MCP_BRIDGE_CALL_TIMEOUT_MS) !== '';
  const fwdOv = process.env.VB6MCP_FORWARD_TIMEOUT_MS !== undefined && String(process.env.VB6MCP_FORWARD_TIMEOUT_MS) !== '';
  const baseBudgetMs = bridgeMode ? bridgeCallTimeoutMs() : FORWARD_TIMEOUT_MS;
  const budgetMs = (implicit === true && effectiveMs > 0) ? Math.min(baseBudgetMs, effectiveMs) : baseBudgetMs;
  // (FIX-26 / F4) The IMPLICIT cap has its OWN env variable, and when IT is what bounded this
  // call the receipt must name IT.  Before this, a call whose budget was demonstrably overridden
  // (capMs 2500, measured live) still reported env{name:null,set:false,value:null} - the receipt
  // could not say WHAT had produced the number it printed.
  const capOv = !!(process.env.VB6MCP_IMPLICIT_FORWARD_TIMEOUT_MS !== undefined
    && String(process.env.VB6MCP_IMPLICIT_FORWARD_TIMEOUT_MS).trim() !== '');
  const env = (implicit === true && capOv)
      ? { name: 'VB6MCP_IMPLICIT_FORWARD_TIMEOUT_MS', set: true, value: IMPLICIT_FORWARD_TIMEOUT_MS }
      : (callOv ? { name: 'VB6MCP_BRIDGE_CALL_TIMEOUT_MS', set: true, value: budgetMs }
                : (fwdOv ? { name: 'VB6MCP_FORWARD_TIMEOUT_MS', set: true, value: FORWARD_TIMEOUT_MS }
                         : { name: null, set: false, value: null }));
  const plusSlackMs = (bridgeMode && !callOv) ? 10000 : 0;
  return { budgetMs: budgetMs, baseBudgetMs: baseBudgetMs, plusSlackMs: plusSlackMs, implicit: implicit === true,
           decidedBy: decidedBy || null, env: env,
           // WHERE the implicit cap came from (null for an explicit call: no cap applies to it)
           capSource: implicit === true ? (capOv ? 'env' : 'default') : null,
           capEnvName: implicit === true ? 'VB6MCP_IMPLICIT_FORWARD_TIMEOUT_MS' : null,
           rule: FORWARD_BUDGET_RULE };
}
// 'registered / known, but the ADD-IN has not answered a probe yet' - the state that makes a
// forward spend the full budget (and the state a caller must see BEFORE it retries blindly).
function probePendingOf(pick) {
  if (!pick) return false;
  let live = pick.liveness || null;
  if (!live && pick.probeStatus) { try { live = livenessOf(pick, bridgeIdentitySync()); } catch (e) { live = null; } }
  if (live && live.state === 'responding') return false;
  if (pick.probeStatus === 'ok') return false;
  return true;
}
// (threat model: bridge/TOKEN_THREAT_MODEL.md).  PURELY ADDITIVE: with no credential
// configured bridgeAuthHeaders() returns {}, so every request this shell makes is byte-identical
// to before.  The VALUE is never logged - only presence and length are ever reported, and a
// value read from a file is never copied into a child process environment.
const BRIDGE_TOKEN = (process.env.VB6MCP_TOKEN || '').trim()
  || (process.env.VB6MCP_TOKEN_FILE ? readBridgeTokenFile(process.env.VB6MCP_TOKEN_FILE) : '');
// Where the LOCAL slot's credential came from - reported as a source name, never as a value.
const BRIDGE_TOKEN_SOURCE = (process.env.VB6MCP_TOKEN || '').trim() ? 'env'
  : (process.env.VB6MCP_TOKEN_FILE ? 'file' : 'none');
function readBridgeTokenFile(p) {
  try { return String(fs.readFileSync(p, 'utf8')).trim(); } catch (e) { return ''; }
}
function bridgeAuthHeaders() {
  return BRIDGE_TOKEN ? { Authorization: 'Bearer ' + BRIDGE_TOKEN } : {};
}
function bridgeAuthState() {
  return BRIDGE_TOKEN ? ('on(len=' + BRIDGE_TOKEN.length + ')') : 'off';
}
// ---- zero-config discovery (zero-config addressing): LOCAL middle service, proven identity first -----------
function parseDiscoveryPortsEnv() {
  const raw = String(process.env.VB6MCP_BRIDGE_DISCOVERY_PORTS || '').trim();
  if (!raw) return [];
  return raw.split(/[,\s]+/).map(function (s) { return parseInt(s, 10); })
    .filter(function (p) { return Number.isInteger(p) && p > 0 && p < 65536; });
}
const BRIDGE_DISCOVERY_PORTS_ENV = parseDiscoveryPortsEnv();
function readDiscoveryPortsFromRecords() {
  // The hub already persists middle.<port>.lock/state.json and the add-in publishes
  // instances/<pid>.json; both are LOCAL files, so reading them opens no socket.
  const dir = path.join(os.tmpdir(), 'vb6mcp');
  const ports = [];
  try {
    for (const name of fs.readdirSync(dir)) {
      if (!/^middle\..+\.(?:lock|state\.json)$/.test(name)) continue;
      const m = /^middle\.(\d+)\./.exec(name);
      if (m) ports.push(parseInt(m[1], 10));
      try {
        const j = JSON.parse(fs.readFileSync(path.join(dir, name), 'utf8'));
        if (j && Number(j.port) > 0) ports.push(Number(j.port));
        if (j && typeof j.bridgeUrl === 'string') {
          const pm = /:(\d+)(?:\/|$)/.exec(j.bridgeUrl);
          if (pm) ports.push(parseInt(pm[1], 10));
        }
      } catch (e) { }
    }
  } catch (e) { }
  try {
    const idir = path.join(dir, 'instances');
    for (const name of fs.readdirSync(idir)) {
      if (!name.endsWith('.json')) continue;
      try {
        const j = JSON.parse(fs.readFileSync(path.join(idir, name), 'utf8'));
        if (j && Number(j.port) > 0) ports.push(Number(j.port));
      } catch (e) { }
    }
  } catch (e) { }
  return ports;
}
function discoveryCandidatePorts() {
  const out = [], seen = new Set();
  function add(p) {
    p = Number(p);
    if (!Number.isInteger(p) || p < 1 || p > 65535 || seen.has(p)) return;
    seen.add(p); out.push(p);
  }
  // An EXPLICIT list is authoritative (ops/test knob) and suppresses the convention scan,
  // so a test can prove the "nothing found" path deterministically.
  if (BRIDGE_DISCOVERY_PORTS_ENV.length) {
    for (const p of BRIDGE_DISCOVERY_PORTS_ENV) add(p);
    return out.slice(0, 80);
  }
  add(8797);                                        // the documented local convention
  for (const p of readDiscoveryPortsFromRecords()) add(p);
  for (let p = 8767; p <= 8796; p++) add(p);        // the hub-owned middle-service range
  return out.slice(0, 80);
}
async function discoveryFetch(url, timeoutMs, headers) {
  try {
    const res = await fetch(url, { signal: AbortSignal.timeout(Math.max(50, timeoutMs)), headers: headers || {} });
    const text = await res.text().catch(function () { return ''; });
    return { ok: res.ok, status: res.status, server: String(res.headers.get('server') || ''), text: text };
  } catch (e) {
    return { ok: false, status: 0, server: '', text: '', error: String((e && e.message) || e) };
  }
}
function middleIdentityOf(body) {
  return !!(body && typeof body === 'object' && String(body.server || body.product || '') === MIDDLE_PRODUCT);
}
async function probeMiddleIdentityAt(url) {
  // PHASE A: NO credentials are sent to a candidate that has not identified itself.
  const a = await discoveryFetch(url + '/health', BRIDGE_DISCOVERY_TIMEOUT_MS, {});
  let body = null;
  try { body = a.text ? JSON.parse(a.text) : null; } catch (e) { body = null; }
  const serverHeader = String(a.server || '');
  const identityByField = a.ok && middleIdentityOf(body);
  const identityByHeader = (a.status === 401 || a.status === 403)
    && serverHeader.toLowerCase().indexOf(MIDDLE_PRODUCT + '/') === 0;
  if (!identityByField && !identityByHeader) {
    return { url: url, status: a.status, identity: false,
             reason: a.status === 0 ? 'unreachable' : 'not_our_service',
             server: serverHeader || null, authRequired: !!(body && body.auth && body.auth.required) };
  }
  const evidence = identityByField
    ? { kind: 'health_server_field', server: String(body.server || body.product), status: a.status }
    : { kind: '401_server_header', server: serverHeader, status: a.status,
        authReason: (body && body.reason) || null };
  if (identityByField && !(body.auth && body.auth.required)) {
    return { url: url, status: a.status, identity: true, ok: true, authRequired: false,
             server: MIDDLE_PRODUCT, evidence: evidence };
  }
  // PHASE B: identity is PROVEN, so the configured credential may now be sent - to this URL only.
  const b = await discoveryFetch(url + '/health', BRIDGE_DISCOVERY_AUTH_TIMEOUT_MS, bridgeAuthHeaders());
  let bbody = null;
  try { bbody = b.text ? JSON.parse(b.text) : null; } catch (e) { bbody = null; }
  const authOk = b.ok && middleIdentityOf(bbody);
  return { url: url, status: b.status, identity: true, ok: !!authOk, authRequired: true,
           authOk: !!authOk,
           authReason: authOk ? null : ((bbody && bbody.reason) || (b.status ? 'http_' + b.status : 'unreachable')),
           server: MIDDLE_PRODUCT, evidence: evidence };
}
async function discoverMiddleService() {
  const started = Date.now();
  const ports = discoveryCandidatePorts();
  const probes = [];
  let hit = null;
  for (const port of ports) {
    if (probes.length && (Date.now() - started) > BRIDGE_DISCOVERY_BUDGET_MS) break;
    const url = 'http://127.0.0.1:' + port;
    const p = await probeMiddleIdentityAt(url);
    probes.push({ port: port, url: url, status: p.status, identity: !!p.identity,
                  reason: p.reason || p.authReason || null, server: p.server || null,
                  authRequired: !!p.authRequired, authOk: p.ok === true,
                  evidence: p.evidence || null });
    if (p.identity) { hit = { port: port, url: url, probe: p }; break; }
  }
  const elapsedMs = Date.now() - started;
  if (hit) {
    BRIDGE_URL = hit.url;
    bridgeHealthCache = { at: 0, data: null, promise: null };
    bridgeIdentityCache = { at: 0, val: null, promise: null };
    middleLastEnsure = readMiddleState();
    const state = hit.probe.ok === true ? 'discovered' : 'discovered_auth_failed';
    return { state: state, source: 'discovery', url: hit.url, port: hit.port,
             candidates: ports, probes: probes, probedCount: probes.length, elapsedMs: elapsedMs,
             evidence: hit.probe.evidence || null,
             reason: state === 'discovered' ? null : (hit.probe.authReason || 'bridge_auth_failed'),
             envOverride: false, legacyDefault: false, authState: bridgeAuthState() };
  }
  return { state: 'not_found', source: 'discovery', url: BRIDGE_URL, port: null,
           candidates: ports, probes: probes, probedCount: probes.length, elapsedMs: elapsedMs,
           evidence: null, reason: 'no_middle_service_discovered',
           envOverride: false, legacyDefault: false, authState: bridgeAuthState() };
}
let bridgeDiscoveryPromise = null;
async function ensureBridgeEndpoint() {
  if (BRIDGE_MODE === 'inproc') return BRIDGE_DISCOVERY;
  if (BRIDGE_URL_EXPLICIT || BRIDGE_DISCOVERY_DISABLED) return BRIDGE_DISCOVERY;
  if (BRIDGE_DISCOVERY.state !== 'pending') return BRIDGE_DISCOVERY;
  if (bridgeDiscoveryPromise) return bridgeDiscoveryPromise;
  bridgeDiscoveryPromise = discoverMiddleService().then(function (d) {
    BRIDGE_DISCOVERY = d;
    dbg('zero-config discovery: state=' + d.state + ' url=' + d.url + ' probes=' + d.probedCount
        + ' elapsedMs=' + d.elapsedMs);
    return d;
  }).catch(function (e) {
    BRIDGE_DISCOVERY = Object.assign({}, BRIDGE_DISCOVERY, {
      state: 'error', reason: String((e && e.message) || e) });
    return BRIDGE_DISCOVERY;
  }).finally(function () { bridgeDiscoveryPromise = null; });
  return bridgeDiscoveryPromise;
}
function bridgeDiscoveryReceipt() {
  if (BRIDGE_MODE === 'inproc') return null;   // rollback path: nothing is added
  const d = BRIDGE_DISCOVERY;
  return { mode: d.state, source: d.source, url: d.url, port: d.port || bridgePortHint() || null,
           envOverride: !!BRIDGE_URL_EXPLICIT, envName: 'VB6MCP_BRIDGE_URL',
           discoveryDisabled: !!BRIDGE_DISCOVERY_DISABLED, legacyDefault: !!d.legacyDefault,
           probedCount: d.probedCount || 0, elapsedMs: d.elapsedMs || 0,
           candidates: d.candidates || [], probes: d.probes || [],
           evidence: d.evidence || null, reason: d.reason || null,
           authState: bridgeAuthState(), rule: BRIDGE_DISCOVERY_RULE };
}
let bridgeHealthCache = { at: 0, data: null, promise: null };
// WHY the middle service is not answering, kept so the discovery/control receipts can
// tell 'the credential was rejected' apart from 'there is no IDE' and from 'it did not answer'.
// Before this, a 401 was logged and then the face reported 'no instance discovered at all' - the
// caller could not tell an authorization failure from an empty machine.
let bridgeLastFailure = null;   // {kind:'auth'|'unreachable', reason, status, detail, at}

// ---- middle-service lifecycle (lazy ensure; the hub KNOWS the dependency, so it OWNS it) ------
// A pipe-only IDE is addressable ONLY through the middle service, but the hub used to assume that
// service was already running.  One restart (DSH, the machine, a crash) therefore left the whole
// tool face silent and the only clue was a null-port URL.  This block is the ONE place that makes
// the middle service EXIST, adopting a healthy one before ever spawning anything.
// Rollback: VB6MCP_BRIDGE=inproc does NO probe, NO spawn and adds NO receipt field, so the
// pre-change path stays byte-for-byte what it was.
const MIDDLE_MANAGED = !(process.env.VB6MCP_BRIDGE_MANAGED === '0');
const MIDDLE_SPAWN_WAIT_MS = parseInt(process.env.VB6MCP_BRIDGE_SPAWN_WAIT_MS || '8000', 10);
const MIDDLE_PROBE_INTERVAL_MS = 250;
const BRIDGE_PY = process.env.VB6MCP_BRIDGE_PY || path.join(HERE, 'vb6mcp_bridge.py');
const MIDDLE_PRODUCT = 'vb6mcp-bridge';
let middleServiceState = { reachable: false, url: BRIDGE_URL, pid: null, startedAt: null,
                           managedByHub: false, managedByThisHub: false,
                           reason: BRIDGE_MODE === 'inproc' ? 'inproc_mode' : 'not_checked',
                           lockFile: null, spawnAttemptedAt: null, portTakenByOther: false, foreignPid: null };
let middleServiceEnsure = null;   // single-flight: two concurrent calls never race a spawn
let middleSpawnedPid = null;     // the ONLY pid this hub process may ever signal (only-kill-what-we-spawned rule)
// ---- P1: PROVABLE code identity ------------------------------------------------
// A DSH restart can leave an OLDER hub alive (measured: pid 12172 from 17:21 served a file that
// was rewritten at 19:27), so "which code answered" must come from a receipt, never a pid guess.
const HUB_SELF_PATH = fileURLToPath(import.meta.url);
const HUB_STARTED_AT = new Date().toISOString();
let HUB_CODE_REV = 'unknown', HUB_FILE_MTIME = null;
try {
  HUB_FILE_MTIME = fs.statSync(HUB_SELF_PATH).mtime.toISOString();
  HUB_CODE_REV = crypto.createHash('sha256').update(fs.readFileSync(HUB_SELF_PATH)).digest('hex').slice(0, 16);
} catch (e) { /* a diagnostic field must never break startup */ }
// null in inproc mode: that rollback receipt stays byte-for-byte unchanged.
function hubCodeStamps() {
  if (BRIDGE_MODE === 'inproc') return null;
  return { hubCodeRev: HUB_CODE_REV, hubStartedAt: HUB_STARTED_AT, fileMtime: HUB_FILE_MTIME };
}
// ---- P2: "a forward just worked" is liveness evidence ---------------------------
const FORWARD_EVIDENCE_TTL_MS = parseInt(process.env.VB6MCP_FORWARD_EVIDENCE_TTL_MS || '120000', 10);
const lastForwardOk = new Map();   // pid -> Date.now() of the last SUCCESSFUL forward
function noteForwardSuccess(pick) {
  if (BRIDGE_MODE === 'inproc') return;   // inproc receipts must stay byte-identical
  if (pick && pick.pid) lastForwardOk.set(Number(pick.pid), Date.now());
}
// bridgeForwardCall returns a STAMPED CLIENT BODY ({content:[{type:'text',text:'<json>'}], isError})
// on BOTH the success and the failure path - only a raw JSON-RPC envelope has a ".result" member.
// Testing ".result" alone therefore scored every stamped answer as "not answered" (measured: a
// pipe-only readiness poll answered 0/87 rounds while the VERY NEXT forwarded mcp_info answered OK
// in 1 s). Accept BOTH shapes so the verdict is about the ANSWER, not about which wrapper carried it.
function forwardEnvelope(stamped) {
  if (!stamped) return null;
  return stamped.result ? stamped.result : stamped;
}
function forwardAnsweredOk(stamped) {
  const res = forwardEnvelope(stamped);
  if (!res) return false;
  if (res.isError === true) return false;
  const obj = firstBodyObject(res);
  return !(obj && obj.ok === false);
}
// WHY a forward did not count as an answer, in ONE small object. A readiness poll that answers
// "false" 89 times must be able to say what it saw on the last round: "the add-in did not answer"
// and "the forward was refused / timed out / the body said ok:false" are DIFFERENT facts.
function forwardFailureDiag(stamped) {
  try {
    if (!stamped) return { stage: 'no_response_at_all' };
    const r = forwardEnvelope(stamped);
    if (!r) return { stage: 'no_envelope' };
    const o = firstBodyObject(r) || {};
    const pick = function (k) { return o[k] === undefined ? null : o[k]; };
    return {
      stage: r.isError === true ? 'isError' : 'body_ok_false',
      reason: pick('reason'), reasonLegacy: pick('reasonLegacy'), failureKind: pick('failureKind'),
      ok: o.ok === undefined ? null : !!o.ok,
      transportReason: pick('transportReason'), budgetMs: pick('budgetMs'), elapsedMs: pick('elapsedMs'),
      message: String(o.message === undefined || o.message === null ? '' : o.message).slice(0, 240)
    };
  } catch (e) { return { stage: 'diag_error', err: String((e && e.message) || e) }; }
}
// ---- P3: lastEnsure persistence + a READ-ONLY report probe ----------------------
function middleStatePath() {
  return path.join(os.tmpdir(), 'vb6mcp', 'middle.' + (bridgePortHint() || 'default') + '.state.json');
}
function readMiddleState() { try { return JSON.parse(fs.readFileSync(middleStatePath(), 'utf8')); } catch (e) { return null; } }
function writeMiddleState(extra) {
  try {
    const d = path.dirname(middleStatePath());
    if (!fs.existsSync(d)) fs.mkdirSync(d, { recursive: true });
    const rec = Object.assign({}, readMiddleState() || {}, { at: new Date().toISOString(), bridgeUrl: BRIDGE_URL, bridgeMode: BRIDGE_MODE }, extra || {});
    fs.writeFileSync(middleStatePath(), JSON.stringify(rec, null, 2));
    return rec;
  } catch (e) { return null; }
}
let middleLastEnsure = readMiddleState();   // survives hub restarts (DSH restarts the hub often)
// ide_list calls this: /health only - NEVER a spawn, NEVER a lock write.
async function middleServiceReportProbe() {
  if (BRIDGE_MODE === 'inproc') return null;
  await ensureBridgeEndpoint();
  if (BRIDGE_DISCOVERY.state === 'not_found') {
    middleServiceState = Object.assign({}, middleServiceState, {
      reachable: false, url: BRIDGE_URL, reason: 'no_middle_service_discovered' });
    return middleServiceReceipt();
  }
  const p = await middleProbeOnce();
  if (p.ok && isOurMiddle(p.body)) {
    let pid = null;
    try { const lp = listeningPortsCached(); const hp = bridgePortHint(); pid = (hp && lp[hp] && lp[hp][0]) || null; } catch (e) { pid = null; }
    middleServiceState.reachable = true;
    if (pid) middleServiceState.pid = Number(pid);
    if (!middleServiceState.managedByThisHub) middleServiceState.reason = 'healthy_by_health_probe';
    middleServiceState.portTakenByOther = false;
    bridgeHealthCache = { at: Date.now(), data: p.body, promise: null };
    bridgeLastFailure = null;
  } else if (p.status === 401) {
    middleServiceState.reachable = false;
    middleServiceState.reason = 'bridge_auth_failed';
  } else if (middleServiceState.reachable) {
    middleServiceState.reachable = false;
    middleServiceState.reason = 'health_probe_failed_at_ide_list';
  }
  return middleServiceReceipt();
}
// null in inproc mode: every caller then OMITS the key, which is what keeps that mode identical.
function middleServiceReceipt() {
  if (BRIDGE_MODE === 'inproc') return null;
  const st = middleServiceState;
  return { reachable: !!st.reachable, url: BRIDGE_URL, pid: st.pid || null, startedAt: st.startedAt || null,
           managedByHub: !!st.managedByHub, managedByThisHub: !!st.managedByThisHub,
           reason: st.reason || null, lockFile: st.lockFile || null,
           spawnAttemptedAt: st.spawnAttemptedAt || null,
           portTakenByOther: !!st.portTakenByOther, foreignPid: st.foreignPid || null,
           discovery: bridgeDiscoveryReceipt(),
           lastEnsure: middleLastEnsure ? { at: middleLastEnsure.at || null, reachable: !!middleLastEnsure.reachable,
             pid: middleLastEnsure.pid || null, reason: middleLastEnsure.reason || null,
             by: middleLastEnsure.bridgeMode || null } : null };
}
function middleLockPath() {
  return path.join(os.tmpdir(), 'vb6mcp', 'middle.' + (bridgePortHint() || 'default') + '.lock');
}
function readMiddleLock() {
  try { return JSON.parse(fs.readFileSync(middleLockPath(), 'utf8')); } catch (e) { return null; }
}
function writeMiddleLock(obj) {
  try {
    const d = path.dirname(middleLockPath());
    if (!fs.existsSync(d)) fs.mkdirSync(d, { recursive: true });
    fs.writeFileSync(middleLockPath(), JSON.stringify(obj, null, 2));
    return true;
  } catch (e) { return false; }
}
function clearMiddleLock() { try { fs.unlinkSync(middleLockPath()); } catch (e) { } }
// ONE raw probe (bypasses the 5 s positive cache: an ensure decision must be FRESH).
async function middleProbeOnce() {
  try {
    const res = await fetch(BRIDGE_URL + '/health', { signal: AbortSignal.timeout(1200), headers: bridgeAuthHeaders() });
    if (!res.ok) {
      const body = await res.text().catch(function () { return ''; });
      return { ok: false, status: res.status, auth: authFailureEnvelope(res.status, body), detail: body.slice(0, 200) };
    }
    const j = await res.json();
    return { ok: true, status: res.status, body: j };
  } catch (e) { return { ok: false, status: 0, detail: String((e && e.message) || e) }; }
}
function isOurMiddle(j) { return !!(j && typeof j === 'object' && String(j.server || j.product || '') === MIDDLE_PRODUCT); }
function adoptMiddle(j) {
  const port = bridgePortHint();
  let pid = null;
  try { const lp = listeningPortsCached(); pid = (port && lp[port] && lp[port][0]) || null; } catch (e) { pid = null; }
  const lock = readMiddleLock();
  const lockPid = lock && lock.pid ? Number(lock.pid) : null;
  // P3: an ADOPTED service gets an ADVISORY record only (by:"adopt", pid:null so a real spawner is
  // never blocked by it).  It grants NO ownership/kill rights: managedByHub stays false and
  // managedByThisHub comes solely from middleSpawnedPid.
  if (!lock || lock.by !== 'hub') {
    writeMiddleLock({ by: 'adopt', pid: null, servicePid: pid ? Number(pid) : null, port: bridgePortHint(),
                      firstSeenAt: (lock && lock.firstSeenAt) || new Date().toISOString(), bridgeUrl: BRIDGE_URL });
  }
  middleServiceState = { reachable: true, url: BRIDGE_URL, pid: pid ? Number(pid) : lockPid, startedAt: (lock && lock.startedAt) || null,
    managedByHub: !!(lock && lock.by === 'hub'),
    // managedByThisHub means THIS process spawned it - never "the lock names the pid", which is how
    // a second hub would have claimed ownership of a service another hub started.
    managedByThisHub: !!(middleSpawnedPid && pid && Number(pid) === Number(middleSpawnedPid)),
    reason: 'healthy', lockFile: middleLockPath(), spawnAttemptedAt: middleServiceState.spawnAttemptedAt || null,
    portTakenByOther: false, foreignPid: null };
  // refresh the cached health so the very NEXT forward sees the adopted service immediately
  bridgeHealthCache = { at: Date.now(), data: j, promise: null };
  bridgeLastFailure = null;
  return middleServiceReceipt();
}
function markPortTaken(detail, foreignPid) {
  middleServiceState = { reachable: false, url: BRIDGE_URL, pid: null, startedAt: null, managedByHub: false, managedByThisHub: false,
    reason: 'port_taken_by_other', lockFile: middleLockPath(), spawnAttemptedAt: middleServiceState.spawnAttemptedAt || null,
    portTakenByOther: true, foreignPid: foreignPid || null };
  dbg('middle service: port ' + bridgePortHint() + ' is occupied by something that is NOT ' + MIDDLE_PRODUCT + ' -> fail-closed (no spawn, no forward)');
  return middleServiceReceipt();
}
function markAuthFailure(detail) {
  middleServiceState = { reachable: false, url: BRIDGE_URL, pid: null, startedAt: null, managedByHub: false, managedByThisHub: false,
    reason: 'bridge_auth_failed', lockFile: middleLockPath(), spawnAttemptedAt: middleServiceState.spawnAttemptedAt || null,
    portTakenByOther: false, foreignPid: null };
  dbg('middle service: /health answered 401 -> credential mismatch, NOT spawning a second service');
  return middleServiceReceipt();
}
async function waitForMiddle(ms) {
  const deadline = Date.now() + Math.max(0, ms || 0);
  let last = null;
  for (;;) {
    last = await middleProbeOnce();
    if (last.ok && isOurMiddle(last.body)) return last;
    if (last.status === 401) return last;
    if (Date.now() >= deadline) return last;
    await new Promise(function (r) { setTimeout(r, MIDDLE_PROBE_INTERVAL_MS); });
  }
}
function spawnMiddleService() {
  const port = bridgePortHint();
  if (!port) return { ok: false, reason: 'bridge_url_has_no_port' };
  const args = [BRIDGE_PY, '--port', String(port)];
  // The credential is inherited through the environment (VB6MCP_TOKEN_FILE / VB6MCP_TOKEN); it is
  // never copied onto a command line.  detached + unref: the service OUTLIVES this hub process.
  let child = null;
  try {
    child = spawn(PYTHON_BIN, args, { detached: true, stdio: 'ignore', windowsHide: true,
      env: Object.assign({}, process.env, { VB6MCP_BRIDGE_PORT: String(port) }) });
    child.unref();
  } catch (e) { return { ok: false, reason: 'spawn_failed: ' + String((e && e.message) || e) }; }
  const lock = { pid: child.pid || null, port: port, startedAt: new Date().toISOString(), by: 'hub', hubPid: process.pid, py: PYTHON_BIN, script: BRIDGE_PY };
  writeMiddleLock(lock);
  middleSpawnedPid = child.pid || null;
  dbg('middle service: spawned pid=' + child.pid + ' port=' + port + ' (detached; lock=' + middleLockPath() + ')');
  return { ok: true, pid: child.pid || null, lock: lock };
}
// Lazy ensure: adopt a healthy service; otherwise spawn ONE detached copy.  Never takes over a
// foreign listener (fail-closed), never spawns while another hub holds a live lock.
async function ensureMiddleService() {
  if (BRIDGE_MODE === 'inproc') return middleServiceReceipt();   // null: nothing runs, nothing is added
  // Zero-config (zero-config addressing): resolve the LOCAL middle service BEFORE any probe or spawn uses
  // BRIDGE_URL.  Nothing discovered => structured failure, and NO listener is opened here.
  await ensureBridgeEndpoint();
  if (BRIDGE_DISCOVERY.state === 'not_found') {
    middleServiceState = { reachable: false, url: BRIDGE_URL, pid: null, startedAt: null,
      managedByHub: false, managedByThisHub: false, reason: 'no_middle_service_discovered',
      lockFile: middleLockPath(), spawnAttemptedAt: middleServiceState.spawnAttemptedAt || null,
      portTakenByOther: false, foreignPid: null };
    return middleServiceReceipt();
  }
  // A service adopted/verified inside the health TTL is reused without another HTTP round trip.
  if (middleServiceState.reachable && bridgeHealthCache.data && (Date.now() - bridgeHealthCache.at) < BRIDGE_HEALTH_TTL_MS) return middleServiceReceipt();
  if (middleServiceEnsure) return middleServiceEnsure;
  middleServiceEnsure = (async function () {
    // (1) a healthy instance is ADOPTED, never restarted
    const first = await middleProbeOnce();
    if (first.ok && isOurMiddle(first.body)) return adoptMiddle(first.body);
    if (first.ok) return markPortTaken(first);
    if (first.status === 401) return markAuthFailure(first);
    // (2) is the port already occupied by something else?  A foreign listener must NEVER be
    //     forwarded to (history: a python listener was once read as a live IDE with 110 tools).
    const port = bridgePortHint();
    const listening = port ? await tcpListening(port, 300) : false;
    if (listening) {
      const second = await middleProbeOnce();
      if (second.ok && isOurMiddle(second.body)) return adoptMiddle(second.body);
      if (second.status === 401) return markAuthFailure(second);
      let foreignPid = null;
      try { const lp = listeningPortsCached(); foreignPid = (lp[port] && lp[port][0]) || null; } catch (e) { foreignPid = null; }
      return markPortTaken(second, foreignPid);
    }
    // (3) single-instance safety: a LIVE lock means another hub/operator owns the startup
    if (!MIDDLE_MANAGED) { middleServiceState.reason = 'management_disabled(VB6MCP_BRIDGE_MANAGED=0)'; return middleServiceReceipt(); }
    const lock = readMiddleLock();
    if (lock && lock.pid && pidAlive(lock.pid)) {
      dbg('middle service: lock held by live pid ' + lock.pid + ' -> waiting/adopting instead of spawning');
      const waited = await waitForMiddle(Math.min(MIDDLE_SPAWN_WAIT_MS, 5000));
      if (waited.ok && isOurMiddle(waited.body)) return adoptMiddle(waited.body);
      if (waited.status === 401) return markAuthFailure(waited);
      middleServiceState = { reachable: false, url: BRIDGE_URL, pid: Number(lock.pid), startedAt: lock.startedAt || null,
        managedByHub: true, managedByThisHub: false, reason: 'lock_held_but_service_not_answering',
        lockFile: middleLockPath(), spawnAttemptedAt: middleServiceState.spawnAttemptedAt || null,
        portTakenByOther: false, foreignPid: null };
      return middleServiceReceipt();
    }
    if (lock && lock.pid && !pidAlive(lock.pid)) { dbg('middle service: clearing stale lock pid=' + lock.pid); clearMiddleLock(); }
    // (4) spawn exactly ONE detached copy and wait for its /health
    const sp = spawnMiddleService();
    if (!sp.ok) { middleServiceState.reason = sp.reason; return middleServiceReceipt(); }
    middleServiceState.spawnAttemptedAt = new Date().toISOString();
    const after = await waitForMiddle(MIDDLE_SPAWN_WAIT_MS);
    if (after.ok && isOurMiddle(after.body)) {
      adoptMiddle(after.body);
      middleServiceState.managedByHub = true;
      middleServiceState.managedByThisHub = true;
      middleServiceState.pid = sp.pid || middleServiceState.pid;
      const lk = readMiddleLock();
      middleServiceState.startedAt = (lk && lk.startedAt) || new Date().toISOString();
      dbg('middle service: ensured pid=' + middleServiceState.pid + ' port=' + bridgePortHint() + ' managedByHub=true');
      return middleServiceReceipt();
    }
    middleServiceState.reason = after.status === 401 ? 'bridge_auth_failed' : 'spawned_but_not_answering';
    if (after.ok && !isOurMiddle(after.body)) middleServiceState.reason = 'port_taken_by_other';
    middleServiceState.portTakenByOther = (middleServiceState.reason === 'port_taken_by_other');
    return middleServiceReceipt();
  })().then(function (ms) {
    if (ms) {
      // P3: persist the OUTCOME, so a restarted hub can report the last ensure BEFORE any probe.
      middleLastEnsure = writeMiddleState({ reachable: !!ms.reachable, pid: ms.pid || null, reason: ms.reason || null,
        managedByHub: !!ms.managedByHub, managedByThisHub: !!ms.managedByThisHub, portTakenByOther: !!ms.portTakenByOther });
    }
    return ms;
  }).finally(function () { middleServiceEnsure = null; });
  return middleServiceEnsure;
}
// Startup breadcrumb, deferred like the endpoint one below: dbg()/DEBUG are declared further
// down, so calling this at module-evaluation time would be a temporal-dead-zone error.
setTimeout(function () {
  dbg('bridge mode=' + BRIDGE_MODE + ' url=' + BRIDGE_URL + ' discovery=' + BRIDGE_DISCOVERY.state
      + (BRIDGE_MODE === 'inproc' ? ' (inproc = the PRE-CHANGE path; no bridge branch is entered)' : ''));
}, 0);
// GET /health, cached.  null means "the bridge is not answering" (then the pre-change snapshot /
// local path is used, which is what keeps the tool face complete with no IDE and no bridge).
// Zero-config (zero-config addressing): the LOCAL endpoint is resolved (once) before the first probe, so every
// cached health answer belongs to the URL the hub actually discovered.
async function bridgeHealth() {
  if (BRIDGE_MODE === 'inproc') return null;
  await ensureBridgeEndpoint();
  if (BRIDGE_DISCOVERY.state === 'not_found') {
    bridgeHealthCache = { at: Date.now(), data: null, promise: null };
    bridgeLastFailure = { kind: 'unreachable', reason: 'no_middle_service_discovered', status: 0,
                          detail: (BRIDGE_DISCOVERY.reason || null), at: Date.now() };
    return null;
  }
  const now = Date.now();
  if (bridgeHealthCache.data && now - bridgeHealthCache.at < BRIDGE_HEALTH_TTL_MS) return bridgeHealthCache.data;
  if (bridgeHealthCache.promise) return bridgeHealthCache.promise;
  bridgeHealthCache.promise = (async function () {
    try {
      const res = await fetch(BRIDGE_URL + '/health', { signal: AbortSignal.timeout(800), headers: bridgeAuthHeaders() });
      // a NON-2xx answer is not a health report.  The old code cached whatever JSON came
      // back (a 401 page has no .instances, but it was truthy), so a wrong token still looked like
      // "the bridge is online" and every forward went on to hand the caller the 401 body.
      if (!res.ok) {
        const body = await res.text().catch(function () { return ''; });
        const af = authFailureEnvelope(res.status, body);
        bridgeHealthCache = { at: Date.now(), data: null, promise: null };
        bridgeLastFailure = { kind: 'auth', reason: (af && af.reason) || ('http_' + res.status),
                              status: res.status, detail: (af && af.message) || null, at: Date.now() };
        dbg('bridge health REJECTED HTTP ' + res.status + (af ? ' reason=' + af.reason : '') + ' -> bridge treated as offline (authRequired)');
        return null;
      }
      const j = await res.json();
      bridgeHealthCache = { at: Date.now(), data: j, promise: null };
bridgeLastFailure = null; //: a healthy bridge clears the recorded failure
      dbg('bridge health ok port=' + (j && j.port) + ' instances=' + ((j && j.instances) || []).length + ' faceTools=' + (j && j.faceTools) + ' auth=' + bridgeAuthState());
      return j;
    } catch (e) {
      bridgeHealthCache = { at: Date.now(), data: null, promise: null };
      bridgeLastFailure = { kind: 'unreachable', reason: 'bridge_unreachable', status: 0,
                            detail: String((e && e.message) || e), at: Date.now() };
      dbg('bridge health FAILED url=' + BRIDGE_URL + ' err=' + String((e && e.message) || e));
      return null;
    }
  })();
  return bridgeHealthCache.promise;
}
// One JSON-RPC POST to the bridge - the SAME wire shape rpcAt() uses against an add-in, so the
// bridge is just another endpoint as far as this file is concerned.
async function bridgeRpc(method, params, timeoutMs) {
  // RULE: a FOREIGN listener on the configured port is fail-closed for EVERY caller.  A
  // history bug read a python listener as a live IDE with 110 tools; no path may POST /rpc to a
  // service that did not identify itself as vb6mcp-bridge.
  if (BRIDGE_MODE !== 'inproc' && middleServiceState.portTakenByOther) {
    const err = new Error('middle service port is occupied by a foreign listener (port_taken_by_other): refusing to POST /rpc to it');
    err.code = 'middle_service_port_taken';
    throw err;
  }
  const body = { jsonrpc: '2.0', id: (Date.now() % 100000) + 1, method: method };
  if (params !== undefined) body.params = params;
  const res = await fetch(BRIDGE_URL + '/rpc', {
    method: 'POST', headers: Object.assign({ 'Content-Type': 'application/json' }, bridgeAuthHeaders()),
    body: JSON.stringify(body), signal: AbortSignal.timeout(timeoutMs || BRIDGE_LIST_TIMEOUT_MS)
  });
  const raw = await res.text();
  const authFail = authFailureEnvelope(res.status, raw);
if (authFail) throw authFailureError(authFail, res.status, BRIDGE_URL); //
  return decodeRpcBody(raw);
}
// The hub-provided names WIN over the face (the pre-change rule, restated here so the bridge
// path produces the SAME 120-name face: 10 hub + 110 add-in, each name once).
function mergeHubFace(faceTools) {
  const merged = [];
  const seenNames = {};
  for (const t of hubTools().concat(faceTools || [])) {
    if (!t || !t.name || seenNames[t.name]) continue;
    seenNames[t.name] = true;
    merged.push(t);
  }
  return merged;
}
// Forward ONE tools/call to the bridge.  Returns a stamped result, or null when the bridge is
// not answering (the caller then uses the pre-change local path, so a dead bridge degrades to
// exactly today's behaviour instead of failing).
//
// The arguments are handed over VERBATIM, _target included - stripping it and resolving it is
// the bridge's job (addressing moved to the Python layer).  stampBody() below only fills keys
// that are undefined, so the bridge's transport / transportReason / *Fallback / pipeBusy marks
// SURVIVE into the receipt and are never overwritten.
async function bridgeForwardCall(name, argsIn, snap, note, pick, implicitFwd, decidedBy) {
  // The hub OWNS the dependency.  Ensure the middle service exists before the first
  // forward.  A FOREIGN listener on the configured port is fail-closed here - it is never read
  // as an IDE and never forwarded to (that is the historical fake-IDE incident).
  const msr = (BRIDGE_MODE === 'inproc') ? null : await ensureMiddleService();
  const hcStamps = hubCodeStamps();
  if (msr && msr.portTakenByOther) {
    return stampBody({ content: [{ type: 'text', text: JSON.stringify({
      ok: false, reason: 'port_taken_by_other', failureKind: 'toolError', failureKindRule: FAILURE_KIND_RULE,
      middleService: msr, hubCode: hcStamps,
      message: 'the configured middle-service port is occupied by something that is NOT the vb6mcp middle service. Refusing to forward: a foreign listener is never read as an IDE or as the middle service.',
      hint: 'stop the foreign process on that port, or point VB6MCP_BRIDGE_URL at the real middle service, then retry. Nothing was forwarded.',
      nothingForwarded: true }, null, 2) }], isError: true }, null, snap.contractVersion);
  }
  const callBudgetMs = implicitFwd ? Math.min(bridgeCallTimeoutMs(), implicitForwardCapMs()) : bridgeCallTimeoutMs();
  const midNote = Object.assign({}, msr ? { middleService: msr } : {}, hcStamps ? { hubCode: hcStamps } : {},
    { forwardBudget: forwardBudgetBlock(!!implicitFwd, implicitFwd ? implicitForwardCapMs() : 0, decidedBy) },
    implicitFwd ? { implicitForward: true, implicitForwardCapMs: implicitForwardCapMs() } : {});
  const h = await bridgeHealth();
  if (!h) {
    // Zero-config (zero-config addressing): with discovery active, "no local middle service" is a STRUCTURED
    // answer with a hint - never a silent fall-through to a local path that would look like a
    // target problem.  An explicit VB6MCP_BRIDGE_URL keeps the legacy behaviour exactly.
    if (BRIDGE_MODE !== 'inproc' && (BRIDGE_DISCOVERY.state === 'not_found'
        || BRIDGE_DISCOVERY.state === 'discovered_auth_failed')) {
      const notFound = BRIDGE_DISCOVERY.state === 'not_found';
      const body = Object.assign({
        ok: false,
        reason: notFound ? 'no_middle_service' : 'bridge_auth_failed',
        failureKind: notFound ? 'noIde' : 'toolError', failureKindRule: FAILURE_KIND_RULE,
        message: notFound
          ? 'no LOCAL vb6mcp middle service was discovered on loopback, so this call has no addresser and was NOT forwarded'
          : 'the discovered LOCAL middle service answered and REFUSED this hub credential, so nothing was forwarded',
        nothingForwarded: true,
        bridgeMode: BRIDGE_MODE, bridgeUrl: BRIDGE_URL, bridgeLastFailure: bridgeLastFailure || null,
        discovery: bridgeDiscoveryReceipt(),
        middleService: msr, hubCode: hcStamps,
        hint: notFound
          ? 'start the middle service (python bridge/vb6mcp_bridge.py --port 8797) or set VB6MCP_BRIDGE_URL to its URL; discovery only connects to loopback and never opens a listener'
          : 'fix the credential (VB6MCP_TOKEN / VB6MCP_TOKEN_FILE) that the discovered middle service requires; a 401 is not an absent IDE',
        howToProceed: notFound
          ? ['start the middle service, then retry (the hub re-discovers once per process)',
             'or set VB6MCP_BRIDGE_URL explicitly (highest priority, no discovery)',
             'or set VB6MCP_BRIDGE_DISCOVERY=0 to use the legacy default single URL']
          : ['set VB6MCP_TOKEN or VB6MCP_TOKEN_FILE to the credential the service accepts',
             'or point VB6MCP_BRIDGE_URL at a service that accepts this hub credential']
      }, note && typeof note === 'object' ? note : {});
      return stampBody({ content: [{ type: 'text', text: JSON.stringify(nullPortUrlGuard(body, 'bridge_discovery_failed(' + String(name) + ')') || body, null, 2) }], isError: true }, null, snap.contractVersion);
    }
    return null;
  }
  const t0 = Date.now();
  let resp;
  try {
    resp = await bridgeRpc('tools/call', { name: name, arguments: argsIn }, callBudgetMs);
  } catch (e) {
    const ms = Date.now() - t0;
    const isTimeout = !!(e && (e.name === 'TimeoutError' || /abort/i.test(String(e && e.message))));
    // the peer ANSWERED and refused the credential - that is toolError (the tool never
    // ran), not "noIde"/"wedged".  Legacy reason values are untouched; the new one is additive.
    const isAuth = !!(e && e.code === 'bridge_auth_failed');
    dbg('bridge forward FAILED tool=' + name + ' ms=' + ms + (isAuth ? ' auth=' + e.reason : '') + ' err=' + String((e && e.message) || e));
    let authBlock = {};
    if (isAuth) authBlock = { authRequired: true, authReason: e.reason, authStatus: e.status, authHint: e.authHint, hubAuth: bridgeAuthState() };
    const fwBody = Object.assign({
      ok: false, reason: isAuth ? 'bridge_auth_failed' : (isTimeout ? 'forward_timeout' : 'forward_failed'),
      message: String((e && e.message) || e), elapsedMs: ms, budgetMs: callBudgetMs,
      transport: 'bridge', transportReason: isAuth ? 'bridge_auth_failed' : 'bridge_unreachable',
      bridgeUrl: BRIDGE_URL, bridgeMode: BRIDGE_MODE, bridgeLastFailure: bridgeLastFailure || null,
      middleService: msr, hubCode: hcStamps,
      failureKind: isAuth ? 'toolError' : (isTimeout ? 'wedged' : 'noIde'), failureKindRule: FAILURE_KIND_RULE,
      // the budget that was spent is part of the receipt, so a slow answer is not read as a hang
      forwardBudget: forwardBudgetBlock(!!implicitFwd, implicitFwd ? implicitForwardCapMs() : 0, decidedBy), retrySafe: !isAuth,
      ...(implicitFwd ? { implicitForward: true, implicitForwardCapMs: implicitForwardCapMs() } : {}),
      ...(probePendingOf(pick) ? { probePending: true } : {})
    }, authBlock);
    if (isTimeout) {
      // a wedged/wedged-ish peer: a blind repeat spends the SAME budget again - say so
      fwBody.howToProceed = [
        'address the instance explicitly: _target="<pid|projectPath>" (a pinned forward skips the ladder and its probe state)',
        'probe liveness first with ide_state (or ide_state{release_inflight:true} when the add-in is wedged)',
        'do NOT simply repeat the same implicit call: it will spend the same budgetMs again with the same result'
      ];
      fwBody.retryNote = 'retrySafe:true means a retry cannot corrupt state; it is NOT the fix - the peer did not answer within ' + callBudgetMs + ' ms'
        + (implicitFwd ? ' (the IMPLICIT forward cap; an explicit _target uses the longer budget)' : '') + '.';
    }
    // A timeout/failure receipt must carry the SAME evidence a success receipt does: the note
    // (ladderLiveFallback / selectedBy / targetResolution / probe state) is merged here too,
    // filling only keys the failure body does not already set.
    if (note && typeof note === 'object') { for (const k of Object.keys(note)) { if (fwBody[k] === undefined) fwBody[k] = note[k]; } }
    const fwGuard = nullPortUrlGuard(fwBody, 'bridge_forward_failed(' + String(name) + ')');
    return stampBody({ content: [{ type: 'text', text: JSON.stringify(fwGuard || fwBody, null, 2) }], isError: true }, null, snap.contractVersion);
  }
  dbg('bridge forward ok tool=' + name + ' ms=' + (Date.now() - t0));
  const result = (resp && resp.result) || { content: [{ type: 'text', text: JSON.stringify(resp) }] };
  // ---- R3-a forced refresh: the tool face may have changed under the cached tools/list --------
  // (1) an ext_* call that FORWARDED SUCCESSFULLY is exactly the event the add-in face changes on;
  // (2) an "unknown tool" refusal means the cached face and the add-in DISAGREE about the face.
  // Both are no-ops unless VB6MCP_HUB_FACE_CACHE enabled the cache (the default path calls nothing).
  if (FACE_RESULT_CACHE_ENABLED) {
    try {
      const lcName = String(name || '').toLowerCase();
      if (FACE_EXT_TOOL_NAMES.indexOf(lcName) >= 0) {
        // ext_* receipts are small, so parsing THIS one is cheap; a REFUSED ext call must not
        // invalidate (nothing changed).
        const bodyObj = firstBodyObject(result);
        const refused = !!(result && result.isError === true) || !!(bodyObj && bodyObj.ok === false);
        if (!refused) faceResultCacheForcedRefresh('ext_tool_succeeded:' + lcName);
      } else if (looksLikeUnknownTool(result)) {
        // this helper NEVER JSON-parses the body (it scans content[].text), so the common path
        // pays only a bounded string search - no parse of a possibly huge tool result.
        faceResultCacheForcedRefresh('unknown_tool_refused:' + String(name || '?'));
      }
    } catch (e) { dbg('face cache forced-refresh hook failed: ' + String((e && e.message) || e)); }
  }
// /N11: the hub's OWN liveness/resolution verdict rides along (ADDITIVE - stampBody only
  // fills keys the body does not already carry).  So a forwarded call that comes back
  // noIde/wedged can be read next to ide_list's verdict for the SAME instance, with the basis.
  // P2: this route just ANSWERED a real call - liveness evidence for ide_list.
  noteForwardSuccess(pick);
  // MULTI-BRIDGE ONLY: a local forwarded receipt names the bridge/endpoint exactly like a direct
  // one, so "which machine did this run on" never needs a second lookup.  Single-bridge unchanged.
  const localIde = MULTI_BRIDGE
    ? { pid: (pick && pick.pid) || null, project: (pick && pick.project) || null, port: (pick && pick.port === undefined ? null : pick.port),
        bridge: (pick && pick.bridge) || LOCAL_ENDPOINT, endpoint: LOCAL_ENDPOINT, remote: false }
    : null;
  return stampBody(result, localIde, snap.contractVersion, Object.assign({}, midNote, note || {}));
}
// Latency budget: the DSH client must get initialize/tools/list in <1s.
const DISCOVERY_TTL_MS = 15000;     // cached (re-validated) discovery is reused for this long
const DISCOVERY_BUDGET_MS = 500;    // hard cap on one discovery pass
const PROBE_TIMEOUT_MS = 250;       // per-RPC cap for the fast scan path
// Forwarded calls may legitimately be LONG (a whole-module get_component_code
// took 22.8s directly).  The budget is configurable and a timeout is reported explicitly.
const FORWARD_TIMEOUT_MS = parseInt(process.env.VB6MCP_FORWARD_TIMEOUT_MS || '120000', 10);
// Read-only classification of the HUB face (the add-in face is classified by
// CodeTools.ToolReadOnly). tools/check_readonly_classification.py reads BOTH tables
// and asserts that the full 131-name face (120 add-in + 11 hub) is covered exactly once.
// Exposed additively on ide_list's receipt; no tool name/argument/description is changed.
// ide_kill_and_relaunch joined the face in contract v1.13.0 and is state-changing (it stops
// the IDE process, clears its registration and relaunches it), so it belongs in the write
// table; it was missing until the read-only audit looked for an unclassified name.
const HUB_READ_ONLY_TOOLS = ['ide_list', 'ide_target', 'list_lessons', 'export_lessons'];
const HUB_WRITE_TOOLS = ['ide_start', 'ide_stop', 'ide_restart', 'ide_kill_and_relaunch', 'add_lesson', 'delete_lesson', 'import_lessons_md'];

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
// of ONE candidate set and `remote` is only an extra marker (author ruling).
// Endpoints come from the env (VB6MCP_REMOTES) and/or bridge/hub_remotes.json; with neither
// configured, every remote code path below is inert and the behaviour is the pre-change one.
const REMOTES_FILE = path.join(HERE, 'hub_remotes.json');
const LOCAL_ENDPOINT = 'local';
const LABEL_RE = /^[A-Za-z0-9_.-]{2,32}$/;
// ---- multi-bridge: VB6MCP_BRIDGES (ADDITIVE; unset => the pre-change path) ------------------
// A bridge is a middle service that owns the FACTS for its machine and forwards add-in calls.
// VB6MCP_BRIDGE_URL names the LOCAL one (label "local", token from VB6MCP_TOKEN/_FILE).
// VB6MCP_BRIDGES adds NAMED bridges, each with ITS OWN token file.  A token is presented ONLY to
// the URL it was configured for - never to another host, never on a 401-driven retry - so one
// machine's credential cannot travel to another machine.
const BRIDGES_ENV = String(process.env.VB6MCP_BRIDGES || '').trim();
const BRIDGES_ENV_PRESENT = BRIDGES_ENV.length > 0;
const REMOTES_FILE_ENV = process.env.VB6MCP_REMOTES_FILE;
// undefined => the default file; "" (explicitly empty) => the file is DISABLED (test/ops isolation).
const REMOTES_FILE_DISABLED = (typeof REMOTES_FILE_ENV === 'string' && REMOTES_FILE_ENV.trim() === '');
const REMOTES_PATH = REMOTES_FILE_DISABLED ? null
  : (typeof REMOTES_FILE_ENV === 'string' && REMOTES_FILE_ENV.trim()
     ? path.resolve(REMOTES_FILE_ENV.trim()) : REMOTES_FILE);
const BRIDGE_NAME_RESERVED = LOCAL_ENDPOINT;
const REMOTE_PROBE_TIMEOUT_MS = parseInt(process.env.VB6MCP_REMOTE_PROBE_TIMEOUT_MS || '1500', 10);
const REMOTE_FORWARD_TIMEOUT_MS = parseInt(process.env.VB6MCP_REMOTE_FORWARD_TIMEOUT_MS || String(FORWARD_TIMEOUT_MS), 10);
const REMOTE_TTL_MS = parseInt(process.env.VB6MCP_REMOTE_TTL_MS || String(DISCOVERY_TTL_MS), 10);
const REMOTE_GRACE_MS = parseInt(process.env.VB6MCP_REMOTE_GRACE_MS || '1500', 10);
// The canonical identity rule, restated VERBATIM in every receipt that uses it.
// ---- forward-failure classification (ADDITIVE;(3), evidence/-20) -----
// After a wedge the add-in still LISTENS, so a forwarded call fails in 1-68 ms with
// "forward_failed / fetch failed" - which reads like "the tool does not exist" or "the
// parameter is wrong", while the truth is a host that is there but not answering.
// failureKind turns that into a MACHINE decision; the legacy `reason` values never change.
// The rule is restated VERBATIM in every receipt that uses it (same convention as TARGET_KEY_RULE).
const FAILURE_KIND_RULE = 'failureKind = noIde (there is nothing to talk to: no live IDE, no route - unqualified call / unknown or unreachable endpoint, or the peer refused the connection) | wedged (a peer WAS found and did not answer: discovery saw it but the probe failed, or the forward timed out / the connection died mid-call - RemoteDisconnected) | toolError (the add-in ANSWERED with a structured error: a JSON-RPC error, result.isError=true, or an ok:false body). The legacy reason values are unchanged; failureKind is additive and is NOT injected when the add-in already classified the failure itself (e.g. desktop_probe failureKind=api_unavailable). The R-1021-03 envelope rides through UNTOUCHED: maxMs/maxMsSource/budgetMs/heldBy/peerState/suggestedAction/retryAfterMs/ctlAvailable/oob (and failureKind/reason when present) are UPSTREAM facts - the hub only fills a key that is absent and NEVER overwrites one, so a peer classified peer_unresponsive with oob{acked:false} cannot be relabelled by the hub. delivered/applied/peerState/suggestedAction inside oob are written by the CLIENT (the middle service), never by an add-in.';
const TARGET_KEY_RULE = 'target key = [endpoint:]projectPath; endpoint omitted = local. Bare pid / bare project NAME stay accepted as LEGACY aliases and are never reported as the canonical key. Normalization: case-insensitive, "/" -> "\\", duplicate + trailing separators trimmed; a LOCAL path is additionally fs.realpathSync.native when the file exists; a REMOTE path is string-normalized only, because the endpoint marker is what disambiguates two machines holding the same path. The same normalized path on ONE endpoint is a conflict (VB6 cannot open the same .vbp twice); the same path on TWO endpoints is NOT a conflict.';
function makeBridgeRecord(label, url, extra) {
  extra = extra || {};
  const rec = { label: label, url: url, note: extra.note ? String(extra.note) : '',
                source: 'VB6MCP_BRIDGES', bridge: true, tokenFile: null,
                token: '', tokenSource: 'none', tokenLen: 0, tokenError: null };
  if (extra.tokenFile) {
    rec.tokenFile = String(extra.tokenFile);
    try {
      const tok = String(fs.readFileSync(rec.tokenFile, 'utf8')).trim();
      if (tok) { rec.token = tok; rec.tokenSource = 'file'; rec.tokenLen = tok.length; }
      else rec.tokenError = 'token_file_empty';
    } catch (e) { rec.tokenError = 'token_file_unreadable'; }
  }
  return rec;
}
// VB6MCP_BRIDGES: JSON (preferred: [{name,url,tokenFile,note}] or {bridges:[...]}) OR the compact
// "name=url,name=url" form (no per-bridge token there - use JSON when a token is needed).
// Every rejection is REPORTED in endpointConfigErrors; nothing is silently dropped.
function parseBridgesEnv(raw, errors) {
  const out = [];
  const pushBridge = function (name, url, extra, source) {
    const label = String(name === undefined || name === null ? '' : name).trim();
    const u = String(url === undefined || url === null ? '' : url).trim().replace(/\/+$/, '');
    if (!LABEL_RE.test(label) || label.toLowerCase() === BRIDGE_NAME_RESERVED) {
      errors.push({ source: source, label: label, error: 'invalid_bridge_name',
                    hint: '2..32 chars of [A-Za-z0-9_.-]; "local" is reserved for VB6MCP_BRIDGE_URL' });
      return;
    }
    if (!/^https?:\/\//i.test(u)) {
      errors.push({ source: source, label: label, error: 'invalid_bridge_url', url: u,
                    hint: 'http(s)://host:port' });
      return;
    }
    const dup = out.filter(function (x) { return x.label.toLowerCase() === label.toLowerCase(); })[0];
    if (dup) {
      errors.push({ source: source, label: label, error: 'duplicate_bridge_name',
                    message: 'the same name appears twice in VB6MCP_BRIDGES; the LAST one wins' });
      out[out.indexOf(dup)] = makeBridgeRecord(label, u, extra);
      return;
    }
    out.push(makeBridgeRecord(label, u, extra));
  };
  const text = String(raw || '').trim();
  if (!text) return out;
  if (text.charAt(0) === '[' || text.charAt(0) === '{') {
    let parsed = null;
    try { parsed = JSON.parse(text); }
    catch (e) {
      errors.push({ source: 'VB6MCP_BRIDGES', error: 'invalid_json', message: String(e && e.message || e) });
      return out;
    }
    const arr = Array.isArray(parsed) ? parsed : (parsed && Array.isArray(parsed.bridges) ? parsed.bridges : null);
    if (!arr) {
      errors.push({ source: 'VB6MCP_BRIDGES', error: 'expected_array_or_{bridges:[]}' });
      return out;
    }
    for (const e of arr) {
      if (!e || typeof e !== 'object' || Array.isArray(e)) {
        errors.push({ source: 'VB6MCP_BRIDGES', error: 'entry_not_an_object', raw: String(e) });
        continue;
      }
      pushBridge(e.name, e.url, { note: e.note, tokenFile: e.tokenFile }, 'VB6MCP_BRIDGES');
    }
    return out;
  }
  const parts = text.split(/[,;]/);
  for (const part of parts) {
    const t = part.trim();
    if (!t) continue;
    const eq = t.indexOf('=');
    if (eq <= 0) {
      errors.push({ source: 'VB6MCP_BRIDGES', error: 'expected_name=url', raw: t.slice(0, 120) });
      continue;
    }
    pushBridge(t.slice(0, eq), t.slice(eq + 1), null, 'VB6MCP_BRIDGES');
  }
  return out;
}
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
  const fromFile = REMOTES_PATH ? readJson(REMOTES_PATH, null) : null;
  const fileArr = Array.isArray(fromFile) ? fromFile : (fromFile && Array.isArray(fromFile.remotes) ? fromFile.remotes : null);
  if (fromFile !== null && !fileArr) errors.push({ source: LEGACY_REMOTES_NAME, error: 'expected_array_or_{remotes:[]}' });
  if (fileArr) for (const e of fileArr) push(e, LEGACY_REMOTES_NAME);
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
  // VB6MCP_BRIDGES last: a named bridge REPLACES a URL-only entry with the same label (it
  // carries the per-bridge credential) and the collision is REPORTED - never silent.
  if (BRIDGES_ENV_PRESENT) {
    for (const b of parseBridgesEnv(BRIDGES_ENV, errors)) {
      const _u = function (v) { return String(v || '').replace(/\/+$/, '').toLowerCase(); };
      const dupUrl = out.filter(function (x) { return !!x.bridge && _u(x.url) === _u(b.url); })[0];
      if (dupUrl) {
        errors.push({ source: b.source, label: b.label, error: 'duplicate_bridge_url', url: b.url,
                      message: 'two bridge names point at ONE URL; each bridge owns its own credential, so the duplicate is ignored',
                      otherLabel: dupUrl.label });
        continue;
      }
      const idx = out.findIndex(function (x) { return x.label.toLowerCase() === b.label.toLowerCase(); });
      if (idx >= 0) {
        errors.push({ source: b.source, label: b.label, error: 'label_conflict',
                      message: 'the name is already used by source "' + out[idx].source + '"; the VB6MCP_BRIDGES entry WINS and the other entry is ignored',
                      otherSource: out[idx].source, otherUrl: out[idx].url });
        out[idx] = b;
      } else {
        out.push(b);
      }
    }
  }
  return { endpoints: out, errors: errors };
}
const LEGACY_REMOTES_NAME = REMOTES_PATH ? path.basename(REMOTES_PATH) : 'hub_remotes.json(disabled)';
const ENDPOINT_CFG = readEndpointConfig();
const REMOTE_ENDPOINTS = ENDPOINT_CFG.endpoints;
const ENDPOINT_CFG_ERRORS = ENDPOINT_CFG.errors;
// The bridge subset, and the ONE switch that makes every multi-bridge addition visible: with no
// VB6MCP_BRIDGES entry this is false and the single-bridge path stays byte-for-byte unchanged.
const BRIDGE_ENDPOINTS = REMOTE_ENDPOINTS.filter(function (e) { return !!e.bridge; });
const MULTI_BRIDGE = BRIDGE_ENDPOINTS.length > 0;
const BRIDGES_RULE = 'bridges[] = the bridge set this hub talks to. local = VB6MCP_BRIDGE_URL (global token); every other row is one VB6MCP_BRIDGES entry with ITS OWN token file. token{configured,source,length} NEVER carries the value. nativeMethods=true means the middle service answered the LOCAL JSON-RPC method ide_list; false means it answered -32601 method-not-found (proved unsupported, cached once per process); null means undecided (not probed yet, or the probe failed for a non-capability reason). discoveryApi names the facts source actually used for that row: bridge_native | bridge_health_fallback. A failing bridge is KEPT in bridges[] with probeStatus=failed + reason + suggestedAction - it is never dropped.';
// Deferred: dbg() reads the DEBUG const that is declared further down, so a top-level call
// here would hit the temporal dead zone.
setTimeout(function () {
  dbg('endpoints: local + ' + REMOTE_ENDPOINTS.length + ' remote(s) ' + JSON.stringify(REMOTE_ENDPOINTS.map(function (e) { return e.label + '=' + e.url; })) + (ENDPOINT_CFG_ERRORS.length ? ' configErrors=' + JSON.stringify(ENDPOINT_CFG_ERRORS) : ''));
}, 0);
// ---- per-bridge credential ------------------------------------------------------------------
// The credential belongs to a CONFIGURED URL, not to "any host that asks": the local slot uses
// the global token (unchanged), a VB6MCP_BRIDGES entry uses ITS token file, and nothing else
// receives a credential from this lookup (the legacy loopback self-limiting rule in rpcAt keeps
// applying byte-identically to every other endpoint).
function endpointAuthFor(base, label) {
  const norm = String(base === undefined || base === null ? '' : base).replace(/\/+$/, '').toLowerCase();
  // LABEL FIRST: a configured bridge is addressed by its own name, so its credential is ITS
  // token file even when two names (mis)point at one URL.  URL matching is the fallback.
  if (label !== undefined && label !== null) {
    if (String(label).toLowerCase() === LOCAL_ENDPOINT) {
      return { label: LOCAL_ENDPOINT, bridge: false, token: BRIDGE_TOKEN, source: BRIDGE_TOKEN_SOURCE,
               tokenConfigured: !!BRIDGE_TOKEN, length: BRIDGE_TOKEN.length, tokenFile: null, tokenError: null };
    }
    for (const ep of REMOTE_ENDPOINTS) {
      if (!ep.bridge || String(ep.label).toLowerCase() !== String(label).toLowerCase()) continue;
      return { label: ep.label, bridge: true, token: ep.token || '', source: ep.tokenSource || 'none',
               tokenConfigured: !!(ep.token || ''), length: (ep.token || '').length,
               tokenFile: ep.tokenFile || null, tokenError: ep.tokenError || null };
    }
  }
  if (!norm) return null;
  if (norm === BRIDGE_URL.toLowerCase()) {
    return { label: LOCAL_ENDPOINT, bridge: false, token: BRIDGE_TOKEN, source: BRIDGE_TOKEN_SOURCE,
             tokenConfigured: !!BRIDGE_TOKEN, length: BRIDGE_TOKEN.length, tokenFile: null, tokenError: null };
  }
  for (const ep of REMOTE_ENDPOINTS) {
    if (!ep.bridge) continue;
    const epNorm = String(ep.url || '').replace(/\/+$/, '').toLowerCase();
    if (epNorm === norm) {
      return { label: ep.label, bridge: true, token: ep.token || '', source: ep.tokenSource || 'none',
               tokenConfigured: !!(ep.token || ''), length: (ep.token || '').length,
               tokenFile: ep.tokenFile || null, tokenError: ep.tokenError || null };
    }
  }
  return null;
}
function authHeadersFor(base, label) {
  const a = endpointAuthFor(base, label);
  if (a && a.token) return { Authorization: 'Bearer ' + a.token };
  return {};
}
// The receipt-safe token description: NEVER the value, only presence/source/length.
function tokenFactForEndpoint(ep) {
  if (!ep || !ep.bridge) return null;
  return { configured: !!(ep.token || ''), source: ep.tokenSource || 'none', length: ep.tokenLen || 0,
           fileConfigured: !!ep.tokenFile, error: ep.tokenError || null };
}
function endpointByLabel(label) {
  const s = String(label === undefined || label === null ? '' : label).toLowerCase();
  if (!s) return null;
  if (s === LOCAL_ENDPOINT) return { label: LOCAL_ENDPOINT, url: null, remote: false, note: 'this machine (instance registry + port scan)' };
  for (const e of REMOTE_ENDPOINTS) if (e.label.toLowerCase() === s) return { label: e.label, url: e.url, remote: true, note: e.note, source: e.source, bridge: !!e.bridge, tokenFile: e.tokenFile || null, token: e.token || '', tokenSource: e.tokenSource || 'none', tokenLen: e.tokenLen || 0, tokenError: e.tokenError || null };
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
  // projectPathAsGiven (ADDITIVE): the spelling the REGISTRY published, next to the canonical
  // realpath.  On this machine the same .vbp is published through junctions
  // (a directory junction that is an alias of the same real directory), so the alias is real information the caller
  // needs to see, not noise to be swallowed by normalization.
  const rec = Object.assign({}, i, {
    endpoint: LOCAL_ENDPOINT, remote: false, endpointUrl: null,
    projectPath: ppath, pathKey: cp.key, pathVia: cp.via,
    projectPathAsGiven: i.projectPath || null,
    targetKey: LOCAL_ENDPOINT + ':' + (ppath || ('pid=' + i.pid))
  });
  // ADDITIVE, multi-bridge only: every candidate names the bridge it lives on ('local' here), so
  // ONE ide_list can be read without re-deriving which machine an instance belongs to.
  if (MULTI_BRIDGE) rec.bridge = LOCAL_ENDPOINT;
  return rec;
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
//
// OPEN-27: the banner belongs to ONE add-in binary/instance, NOT to a port.  A pipe-only instance
// PUBLISHES PORT 0 ON PURPOSE (the named pipe IS its transport; readRegistry turns 0 into null), so
// a port-keyed cache collapsed EVERY pipe-only IDE onto one key and the hub process replayed the
// FIRST DLL's banner to every later pipe IDE - a dead DLL's v1.16.4 banner kept answering for a new
// v1.16.5 DLL until the hub PROCESS was restarted.  A TCP port can be recycled the same way, so a
// port is never an identity.
//   key  = endpoint/transport + pid/port/project + the identity the add-in itself reports
//          (mcp_info contract + build.sha256)
//   slot = the instance part of the key; a CHANGED dllSha256/addinContract invalidates (and prunes)
//          the banner cached for the SAME slot, instead of keeping the old one as a second answer.
// The cache is bounded; evicting an entry only costs one more background /mcp initialize.
const BANNER_CACHE_MAX = 64;
const bannerCache = {};               // key -> { at, text, slot, identity }
function bannerIdentity(fields) {
  const f = fields || {};
  const rawPort = f.port;
  const hasPort = !(rawPort === undefined || rawPort === null || rawPort === '' || Number(rawPort) === 0);
  return {
    endpoint: String(f.endpoint === undefined || f.endpoint === null ? '' : f.endpoint),
    endpointUrl: String(f.endpointUrl === undefined || f.endpointUrl === null ? '' : f.endpointUrl),
    transport: String(f.transport === undefined || f.transport === null ? '' : f.transport),
    pid: (f.pid === undefined || f.pid === null) ? null : String(f.pid),
    port: hasPort ? String(rawPort) : null,
    project: String(f.project === undefined || f.project === null ? '' : f.project).toLowerCase(),
    addinContract: String(f.addinContract === undefined || f.addinContract === null ? '' : f.addinContract),
    dllSha256: String(f.dllSha256 === undefined || f.dllSha256 === null ? '' : f.dllSha256).toUpperCase()
  };
}
function bannerCacheKey(scope, id) {
  return scope + '|' + JSON.stringify([id.endpoint, id.endpointUrl, id.transport, id.pid, id.port, id.project, id.addinContract, id.dllSha256]);
}
function bannerCacheSlot(scope, id) {
  return scope + '|' + JSON.stringify([id.endpoint, id.endpointUrl, id.transport, id.pid, id.port, id.project]);
}
// Drop every banner cached for the same instance SLOT under a DIFFERENT banner identity: the add-in
// reported a new dllSha256/addinContract, so the old banner is invalid rather than a second answer.
function bannerCachePruneOther(slot, keepKey) {
  for (const k of Object.keys(bannerCache)) {
    const e = bannerCache[k];
    if (e && e.slot === slot && k !== keepKey) delete bannerCache[k];
  }
}
function bannerCacheGet(key) {
  const e = bannerCache[key];
  return (e && typeof e.text === 'string') ? e.text : null;
}
function bannerCachePut(key, slot, identity, text) {
  if (typeof text !== 'string' || !text.length) return null;
  bannerCachePruneOther(slot, key);
  bannerCache[key] = { at: Date.now(), text: text, slot: slot, identity: identity };
  const keys = Object.keys(bannerCache);
  if (keys.length > BANNER_CACHE_MAX) {
    keys.sort(function (a, b) { return ((bannerCache[a] && bannerCache[a].at) || 0) - ((bannerCache[b] && bannerCache[b].at) || 0); });
    for (const k of keys.slice(0, keys.length - BANNER_CACHE_MAX)) delete bannerCache[k];
  }
  return bannerCache[key];
}
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
// Liveness of ONE instance, transport aware.  A TCP instance keeps the pre-change test
// byte for byte (pid alive AND port listening).  A PIPE-only instance has NO port to listen on -
// the pipe IS the transport and the middle service owns it - so its liveness is the pid; asking
// "is port null listening?" answered "no" for every live pipe IDE and silently dropped it from
// every re-validation (cachedDiscovery, faceState), which flipped ide_list to faceSource=snapshot
// with the IDE running.  A pipe instance that ALSO has a TCP port (tcpEnabled) is still held to
// the port test.
async function instanceAlive(i, listening) {
  if (!i || !i.pid) return false;
  if (!pidAlive(i.pid)) return false;
  const isPipe = String(i.transport || '') === 'pipe';
  if (isPipe && !i.port) return true;
  if (!i.port) return false;
  return await isListening(i.port, listening || null);
}
// ============================================================================================
// ONE fact source for liveness / ownership / default target (: N11 + N12 + N14 +)
// --------------------------------------------------------------------------------------------
// Three receipts used to answer the same question three different ways:
//   * ide_list    "is the IDE alive?"      <- the REGISTRY (pid alive + a listening port/pipe)
//                   plus a probe that asked tools/list, WHICH THE MIDDLE SERVICE ANSWERS FROM ITS
//                   OWN SNAPSHOT when the add-in cannot be reached.  So a wedged IDE still got
//                   probeStatus=ok / liveToolCount=110 while a forwarded tools/call came back
//                   noIde / pipe busy 121  => the diagnostic entrance contradicted the live
//                   behaviour (N11, measured).
//   * ide_restart "is this instance OURS?" <- the PORT (port < 8767).  A pipe-only instance
//                   publishes port 0 ON PURPOSE, so 0 < 8767 refused it as "foreign" although
//                   ide_list said startedByHub=true and ide_stop acted on it (N12, measured).
//   * _target     "which instance does an omitted / auto target mean?" <- no rule: one live IDE
//                   selected itself, two refused with ambiguous_target, and 'auto' was not a
//                   selector at all, so the caller could not predict the outcome (N14).
// Everything below is the SINGLE place where those three judgements are made, and every receipt
// that reports one of them also carries the rule VERBATIM plus the evidence it was derived from.
const LIVENESS_RULE = 'liveness.state = responding (the ADD-IN ITSELF answered an MCP tools/call (mcp_info) over the SAME route a forwarded call uses) | not_responding (a peer was found but did not answer: probe failed / the pipe is not available - the SAME verdict a forwarded call reports as wedged|noIde) | unknown (no fresh verdict: probe still pending, or only a CACHED verdict older than the discovery TTL). probeStatus is DERIVED from this one verdict (ok <=> responding; not_responding => failed + probeReason), so ide_list and a forwarded call cannot disagree, and a cached ok is never re-published as a fresh one. basis names the evidence, evidenceAgeMs dates it.';
const OWNERSHIP_RULE = 'ownership = hub_started (the hub that is answering started this pid: in-memory startedByHub) | registry_started_by_hub (the instance REGISTRATION itself says startedByHub=true) | not_owned (neither). owned = the first two. The PORT is NOT an ownership criterion: a pipe-only instance publishes port 0 on purpose, so "no port" never means "not ours" (N12). The port stays a SECOND, independent safety guard: an instance whose KNOWN port is < 8767 is foreign (8766 belongs to an external process on this machine) and is refused even when owned. ide_list / ide_stop / ide_restart all call ownershipOf() and report ownership + ownershipBasis + ownershipRule.';
// ---- ONE attribution verdict ---------------------------------------------------
// THREE entrances answered "is this instance OURS?" three different ways, and the PORT was allowed
// to stand in for the answer:
//   * ide_list reported a per-instance STRING (registry_started_by_hub / not_owned);
//   * the forwarded ide_state receipt carried hubInstance.ownership = pick.ownership, which is NULL
//     for every pick the hub resolved itself => "not ours" printed next to ide_list's "ours";
//   * ide_restart read a port outside its own range as foreign.
// The SAME instance could therefore be "ours and alive" on one entrance and "foreign" on another.
// This block is the ONE resolver; its output is shipped VERBATIM by ide_list (per instance AND per
// residual registration), by the forwarded ide_state receipt, and by ide_restart.
//
// TWO AXES, deliberately separated (they answer different questions and drive different actions):
//   * legacy ownership / ownedByHub = "did THIS hub process start it?" -> the PERMISSION axis for
//     ide_stop / ide_restart.  Unchanged; its predicate stays ownershipOf().
//   * ownershipVerdict.verdict = "whose instance IS it?" -> the ATTRIBUTION axis.  It never grants
//     permission: a user-launched IDE of this machine is owned here and still NOT stoppable.
const OWNERSHIP_VERDICT_RULE = "ownershipVerdict.verdict = owned | foreign | unknown (ownershipBasis = a stable code: hub_started_in_memory | hub_started_registry | registry_started_by_hub_false | remote_bridge_instance | no_local_record). A REMOTE bridge instance is FOREIGN: it belongs to another machine and this hub cannot signal it. (ATTRIBUTION: whose instance is it? It is NOT the permission axis - legacy ownership/ownedByHub decides whether ide_stop/ide_restart may act, and that predicate is unchanged). owned: the hub-started evidence is present (THIS hub process started the pid in this lifetime, or the registration says startedByHub=true). A user-launched IDE is NOT owned: ownership is about WHO STARTED IT, which is exactly what a stop/restart decision needs. foreign: POSITIVE evidence only - (1) the registered pid is NOT alive while its port still ANSWERS, (2) the listening port is owned by an OS pid other than the pid the registration names, (3) the known port is inside the reserved range (< 8767; 8766 belongs to an external process) and ANSWERS, (4) the pid that serves the instance differs from the pid its OWN registration names. unknown: EVERYTHING else - above all portAvailable=false: a pipe-only registration publishes port 0 ON PURPOSE and an unreachable/unknown port is NOT evidence that the instance belongs to someone else. basis[] names every fact used; registryAlive and portAvailable are the raw inputs.";
// The defect this resolver exists for, as ONE named constant: a port that CANNOT be obtained must
// never be read as foreignership.  A test flips it to reproduce the pre-fix reading (mutant), so
// the guard is load-bearing instead of folklore.
const OWNERSHIP_PORT_UNAVAILABLE_IS_NOT_FOREIGN = true;

const PROCESS_KIND_RULE = 'processKind = vb6_ide (an add-in registration: %TEMP%\\vb6mcp\\instances\\<pid>.json, pid alive and its transport reachable) | middle_service (the loopback TCP listener that IS this hub\'s configured middle service: the port of VB6MCP_BRIDGE_URL, or the process that owns that port) | unknown_listener (a listening port with neither an add-in registration nor a middle-service identity). instances[] is the AUTHORITATIVE IDE view and lists ONLY vb6_ide; every excluded listener is kept in nonIdeListeners[] with the basis that excluded it, so the middle service process can never be read as an IDE .';
const TARGET_RESOLUTION_RULE = 'default target ladder, used when _target is omitted or "auto" (after the ide_target default has been tried): 1) the ONLY live IDE; 2) else the ONLY hub-owned IDE (see ownershipRule); 3) else the ONLY IDE that is RESPONDING (see livenessRule); 4) else REFUSE with reason=ambiguous_target and list every candidate + a concrete suggestion - never guess. An EXPLICIT _target that matches nothing is target_mismatch (the live set is never silently substituted for it); one that matches several is target_conflict. Every receipt that resolved a target carries resolution{rule,step,liveCount,selectedBy,candidates,ladder}.';
// The rule ide_list restates verbatim (kept in ONE constant so it cannot drift from the ladder).
const SELECTION_RULE = '_target (EXPLICIT) > ide_target default > the DOCUMENTED DEFAULT LADDER (the ONLY live IDE -> the ONLY hub-owned IDE -> the ONLY RESPONDING IDE -> REFUSE with candidates[] + suggestion{}; see targetResolutionRule); a STALE ide_target default is IGNORED (never vetoes the live set). An explicit _target that matches nothing is target_mismatch, one that matches several is target_conflict; the live set is never silently substituted. The canonical selector is [endpoint:]projectPath; bare pid / bare NAME are legacy aliases.';
// The port of the CONFIGURED middle service (no network): a portscan candidate on this port is
// the middle service itself, not an IDE.  VB6MCP_BRIDGE_URL is what the hub forwards to.
function bridgePortHint() {
  const m = /:(\d+)(?:\/|$)/.exec(String(BRIDGE_URL || ''));
  return m ? parseInt(m[1], 10) : null;
}
// GET /health, cached like bridgeHealth() itself: the middle service is authoritative about the
// TRANSPORT of every instance it can address (pipeAvailable / pipeState / pidAlive).  Those are
// the facts the hub does NOT have on its own, and they are what closes N11.
let bridgeIdentityCache = { at: 0, val: null, promise: null };
function bridgeIdentity() {
  if (bridgeIdentityCache.val && Date.now() - bridgeIdentityCache.at < BRIDGE_HEALTH_TTL_MS) return Promise.resolve(bridgeIdentityCache.val);
  if (bridgeIdentityCache.promise) return bridgeIdentityCache.promise;
  bridgeIdentityCache.promise = (async function () {
    const val = { ok: false, mode: BRIDGE_MODE, url: BRIDGE_URL, port: bridgePortHint(), pids: [], byPid: {},
                  instances: [], transport: null, faceTools: null, basis: 'VB6MCP_BRIDGE_URL' };
    if (BRIDGE_MODE !== 'inproc') {
      const h = await bridgeHealth();
      if (h) {
        val.ok = true;
        if (h.port) val.port = Number(h.port) || val.port;
        val.transport = h.transport || null;
        val.faceTools = h.faceTools || null;
        val.instances = h.instances || [];
        for (const bi of val.instances) { if (bi && bi.pid) val.byPid[String(bi.pid)] = bi; }
        val.basis = 'GET /health (' + BRIDGE_URL + ')';
      } else {
        val.basis = 'GET /health unreachable (' + BRIDGE_URL + '): the hub has NO transport verdict from the middle service';
      }
    }
    if (val.port) { try { val.pids = (listeningPortsCached()[val.port] || []).map(Number); } catch (e) { val.pids = []; } }
    bridgeIdentityCache = { at: Date.now(), val: val, promise: null };
    return val;
  })();
  return bridgeIdentityCache.promise;
}
function bridgeIdentitySync() { return bridgeIdentityCache.val || null; }
// Is this candidate an IDE, the middle service, or an unidentified listener?  Registry provenance
// WINS (an add-in registration is direct evidence), then the middle service is identified by the
// port it listens on - the exact confusion measured, where a portscan adopted the python
// service on 8797 and called it a live IDE with 110 tools.
function classifyProcess(c, ctx) {
  if (!c) return { processKind: 'unknown_listener', isIde: false, processKindBasis: 'no_candidate' };
  if (c.remote) return { processKind: 'vb6_ide', isIde: true, processKindBasis: 'remote_endpoint_probe(' + (c.endpoint || '?') + ')' };
  if (c.discoveredBy === 'registry' || c.registryFile) {
    return { processKind: 'vb6_ide', isIde: true,
      processKindBasis: 'addin_registration(' + (c.registryFile || 'registry') + ')' + (c.discoveredBy === 'bridge_health' ? ' read by the middle service (the hub\'s own registry view was empty)' : '') };
  }
  const bp = (ctx && ctx.port) || bridgePortHint();
  const bpids = ((ctx && ctx.pids) || (bridgeIdentitySync() ? bridgeIdentitySync().pids : []) || []).map(Number);
  if (bp && Number(c.port) === Number(bp)) return { processKind: 'middle_service', isIde: false, processKindBasis: 'port==bridge_port(' + bp + ')' };
  if (c.pid && bpids.indexOf(Number(c.pid)) >= 0) return { processKind: 'middle_service', isIde: false, processKindBasis: 'pid_owns_bridge_port(' + bp + ')' };
  if (c.probeStatus === 'ok' && c.probeBasis) return { processKind: 'vb6_ide', isIde: true, processKindBasis: 'portscan_addin_answered(' + c.probeBasis + ')' };
  return { processKind: 'unknown_listener', isIde: false, processKindBasis: 'portscan_no_addin_answer' };
}
// ONE ownership predicate, shared by ide_list / ide_stop / ide_restart.
function ownershipOf(inst) {
  const pid = inst && inst.pid ? Number(inst.pid) : null;
  const inMap = !!(pid && startedByHub.has(pid));
  const byReg = !!(inst && inst.startedByHub);
  const port = Number(inst && inst.port) || 0;
  const foreignPort = port > 0 && port < 8767;
  return {
    owned: inMap || byReg,
    ownership: inMap ? 'hub_started' : (byReg ? 'registry_started_by_hub' : 'not_owned'),
    ownershipBasis: inMap
      ? 'in-memory startedByHub map of THIS hub process lists pid ' + pid
      : (byReg ? 'instance registration startedByHub=true (' + (inst.registryFile || 'registry') + ')' : 'neither the startedByHub map nor the instance registration marks pid ' + pid + ' as hub-started'),
    port: port || null,
    foreignPort: foreignPort,
    portRuleApplies: port > 0,
    portBasis: port > 0
      ? ('known port ' + port + (foreignPort ? ' < 8767 => foreign (8766 belongs to an external process)' : ' >= 8767 => inside the hub-owned range'))
      : 'port null/0 (pipe-only): the port rule CANNOT apply - "no port" never means "not ours" (N12)'
  };
}
// ---- the ONE ownership/attribution resolver ------------------------------------
// Inputs, all named in basis[]: the registration (%TEMP%\vb6mcp\instances\<pid>.json), whether the
// registered pid is ALIVE, whether the port is OBTAINABLE (and, when netstat is available, WHO owns
// it), the project path, and the in-memory startedByHub map of THIS hub process.
// Output = the block below, reused by ide_list / the forwarded ide_state / ide_restart verbatim.
let ownershipRegCache = { at: 0, map: null };
function ownershipRegistrationOf(pid) {
  if (!pid) return null;
  if (!ownershipRegCache.map || Date.now() - ownershipRegCache.at > 1000) {
    const map = {};
    try { for (const e of (readRegistry().entries || [])) { if (e && e.pid) map[String(e.pid)] = e; } } catch (e) { }
    ownershipRegCache = { at: Date.now(), map: map };
  }
  return ownershipRegCache.map[String(pid)] || null;
}
// The registration the INSTANCE came from (its registryFile), which is not necessarily keyed by the
// pid the instance view ended up using: the portscan/adoption path keys an instance by the LISTENER
// pid from netstat.  That difference is ownership evidence, not noise (rule (4)).
function ownershipRegistrationByFile(file) {
  if (!file) return null;
  try {
    const base = path.basename(String(file));
    for (const e of (readRegistry().entries || [])) { if (e && e.file && path.basename(e.file) === base) return e; }
  } catch (e) { }
  return null;
}
// WHO owns a listening port, per the OPERATING SYSTEM.  Unavailable => UNKNOWN (null), never
// "nobody" and never "foreign" - an absent fact is not a negative fact.
function portOwnerPidsOf(port) {
  if (!port) return { available: false, pids: null, why: 'no_port' };
  if (process.env.VB6MCP_HUB_NO_NETSTAT === '1') return { available: false, pids: null, why: 'netstat_disabled' };
  try {
    const map = listeningPortsCached();
    if (!map) return { available: false, pids: null, why: 'netstat_failed' };
    const pids = map[port] || [];
    return { available: true, pids: pids.map(Number), why: pids.length ? 'listening' : 'not_listening' };
  } catch (e) { return { available: false, pids: null, why: 'netstat_failed' }; }
}
function ownershipVerdictBlock(f, verdict, basis, ownedEv, foreignEv) {
  // ownershipBasis = the STABLE code a caller can switch on (basis[] is the prose evidence).
  const code = f.ownershipBasis || (verdict === 'owned'
    ? 'hub_started'
    : (verdict === 'foreign' ? (f.remote ? 'remote_bridge_instance'
        : (f.registryPresent ? 'registry_started_by_hub_false' : 'no_local_record')) : 'no_local_record'));
  return { verdict: verdict, ownershipBasis: code, bridge: f.bridge || null, startedByHub: !!f.startedByHub,
           basis: basis,
           registryAlive: !!(f.registryPresent && f.pidAlive),
           portAvailable: (f.portAvailable === undefined) ? null : f.portAvailable,
           pid: (f.pid === undefined) ? null : f.pid,
           port: f.port || null,
           project: f.project || null,
           projectPath: f.projectPath || null,
           registryFile: f.registryFile || null,
           registryPresent: !!f.registryPresent,
           registryStartedByHub: !!f.registryStartedByHub,
           inMemoryStartedByHub: !!f.inMemory,
           pidAlive: !!f.pidAlive,
           probeAnswered: !!f.probeAnswered,
           registeredPid: (f.registeredPid === undefined) ? null : f.registeredPid,
           portOwnerPids: f.portOwnerPids || null,
           portOwnerWhy: f.portOwnerWhy || null,
           netstatAvailable: !!f.netstatAvailable,
           ownedEvidence: ownedEv,
           foreignEvidence: foreignEv,
           rule: OWNERSHIP_VERDICT_RULE };
}
// PURE: facts in, verdict out.  No I/O, so ide_list / ide_state / ide_restart cannot drift.
function ownershipVerdictOf(f) {
  const basis = [], ownedEv = [], foreignEv = [];
  f = f || { pid: null, port: 0 };
  if (f.remote) {
    // A remote bridge instance is a POSITIVE fact: the process is on another machine and THIS hub
    // cannot signal it.  It is therefore foreign (not "unknown"): attribution needs no local probe.
    const b = f.bridge || null;
    basis.push('the instance lives on bridge ' + (b ? '"' + b + '"' : '(unnamed)') + ': its pid/port/process belong to ANOTHER machine');
    basis.push('startedByHub=' + (f.startedByHub ? 'true' : 'false') + ' (published by the owning bridge); this hub has no in-memory record of it');
    foreignEv.push('a remote bridge instance is never attributable to THIS machine hub process');
    return ownershipVerdictBlock(Object.assign({}, f, { ownershipBasis: 'remote_bridge_instance' }), 'foreign', basis, ownedEv, foreignEv);
  }
  // ---- the FACTS, each one named (a verdict without its evidence is not auditable) ------------
  if (f.pid) basis.push('pid ' + f.pid + (f.pidAlive ? ' is ALIVE' : ' is NOT alive'));
  if (f.registryPresent) {
    basis.push('registration ' + (f.registryFile || ('instances/' + f.registeredPid + '.json')) + ' is present (registered pid ' + f.registeredPid + ')' + (f.registryStartedByHub ? ', startedByHub=true' : ', and it does NOT say startedByHub'));
  } else {
    basis.push('no instance registration is visible for pid ' + f.pid + ' (attribution can only rest on the in-memory startedByHub map or on the instance answering)');
  }
  if (f.inMemory) basis.push('THIS hub process started pid ' + f.pid + ' (in-memory startedByHub map)');
  if (f.port > 0) basis.push('known port ' + f.port);
  if (f.portAvailable === true) {
    basis.push('port ' + (f.port || '?') + ' ANSWERED a reachability test' + (f.portEvidence ? ' (' + f.portEvidence + ')' : ''));
  } else if (f.portAvailable === false) {
    basis.push('portAvailable=false: ' + (f.port ? ('port ' + f.port + ' did not answer a reachability test (' + (f.portEvidence || '?') + ')') : 'no port is obtainable at all (a pipe-only registration publishes port 0 ON PURPOSE)') + ' - an unobtainable port is NOT evidence of foreignership');
  } else {
    basis.push('port availability is UNKNOWN (no reachability test was possible)');
  }
  if (f.probeAnswered) basis.push("the ADD-IN ITSELF answered mcp_info over this instance's route (a real add-in of this machine answered)");
  if (f.netstatAvailable) {
    basis.push('OS listener table (netstat): ' + ((f.portOwnerPids && f.portOwnerPids.length) ? ('port ' + f.port + ' is owned by pid(s) ' + f.portOwnerPids.join(',')) : ('nothing is listed as listening on port ' + (f.port || '?'))));
  } else if (!f.port) {
    basis.push('OS listener table: there is NO PORT to look up (a pipe-only registration publishes port 0 ON PURPOSE), so no port owner can incriminate this instance');
  } else {
    basis.push('OS listener table unavailable (' + ((f.portOwnerWhy === 'netstat_disabled') ? 'VB6MCP_HUB_NO_NETSTAT=1' : 'netstat did not answer') + '): the port OWNER is unknown - not foreign');
  }
  // ---- OWNED (attribution) evidence -----------------------------------------------------------
  if (f.inMemory) ownedEv.push('in-memory startedByHub map of THIS hub process lists pid ' + f.pid);
  if (f.registryStartedByHub) ownedEv.push('the registration ' + (f.registryFile || '') + ' says startedByHub=true (a hub wrote it)');
  // A user-launched IDE (someone else started VB6) is NOT owned by this hub: attribution is about
  // WHO STARTED IT, which is exactly what a caller needs before any stop/restart decision.
  if (!f.inMemory && !f.registryStartedByHub && f.registryPresent) {
    foreignEv.push('the registration ' + (f.registryFile || '') + ' does NOT say startedByHub and this hub never started pid ' + f.pid + ' (a user/other process did)');
  }
  if (!f.inMemory && !f.registryStartedByHub && !f.registryPresent) {
    foreignEv.push('NO local registration and NO in-memory startedByHub record names pid ' + f.pid + ': this hub has no evidence it started it');
  }
  // ---- FOREIGN evidence: positive facts only --------------------------------------------------
  if (f.registryPresent && !f.pidAlive && f.portAvailable === true) {
    foreignEv.push('the registered pid ' + f.registeredPid + ' is NOT alive, but its port ' + f.port + ' still ANSWERS: another process holds that port');
  }
  if (f.portAvailable === true && f.portOwnerPids && f.portOwnerPids.length && f.registeredPid &&
      f.portOwnerPids.indexOf(Number(f.registeredPid)) < 0) {
    foreignEv.push('port ' + f.port + ' is LISTENING and is owned by OS pid(s) ' + f.portOwnerPids.join(',') + ', NOT by the pid the registration names (' + f.registeredPid + ')');
  }
  if (f.port > 0 && f.port < 8767 && f.portAvailable === true) {
    foreignEv.push('port ' + f.port + ' is inside the reserved foreign range (< 8767; 8766 belongs to an external process) and it ANSWERS');
  }
  if (f.pid && f.registeredPid && Number(f.registeredPid) !== Number(f.pid)) {
    foreignEv.push('the instance is served as pid ' + f.pid + ', but the registration it came from (' + (f.registryFile || '?') + ') names pid ' + f.registeredPid + ': the process that answers is NOT the registered process');
  }
  if (!OWNERSHIP_PORT_UNAVAILABLE_IS_NOT_FOREIGN && f.portAvailable === false && !f.inMemory) {
    foreignEv.push('PRE-FIX PORT-ONLY READING: no obtainable port treated as foreign (the reading this resolver exists to forbid)');
  }
  const verdict = foreignEv.length ? 'foreign' : (ownedEv.length ? 'owned' : 'unknown');
  if (verdict === 'owned' && !f.ownershipBasis) {
    f.ownershipBasis = f.inMemory ? 'hub_started_in_memory' : 'hub_started_registry';
  }
  if (!ownedEv.length && !foreignEv.length) {
    basis.push('no owned evidence and no positive foreign evidence: the honest verdict is UNKNOWN (nothing was assumed)');
  }
  return ownershipVerdictBlock(f, verdict, basis, ownedEv, foreignEv);
}
// The async half: gathers the facts the instance object does not already carry (registration, pid
// liveness, port reachability, OS port owner) and hands them to the PURE verdict function above.
async function resolveOwnership(inst, opts) {
  opts = opts || {};
  const pid = inst && inst.pid ? Number(inst.pid) : null;
  const port = Number(inst && inst.port) || 0;
  const remote = !!(inst && inst.remote);
  const inMemory = !remote && !!(pid && startedByHub.has(pid));
  // A REMOTE pid lives on another machine: the LOCAL registry/liveness facts describe a DIFFERENT
  // process that happens to share the number (the same-pid case), so they are never consulted.
  const byFile = (!remote && inst && inst.registryFile) ? ownershipRegistrationByFile(inst.registryFile) : null;
  const byPid = (!remote && pid) ? ownershipRegistrationOf(pid) : null;
  const reg = byFile || byPid;
  const registryPresent = !!reg;
  const registeredPid = reg && reg.pid ? Number(reg.pid) : pid;
  const registryStartedByHub = !!(reg && reg.startedByHub);
  const pidAliveNow = (!remote && pid) ? pidAlive(pid) : false;
  const probeAnswered = !!(inst && (inst.probeStatus === 'ok' ||
    (inst.liveness && (inst.liveness.state === 'responding' || inst.liveness.evidenceByForward))));
  let portAvailable = null, portEvidence = null;
  if (remote) { portAvailable = null; portEvidence = 'remote_endpoint(not tested on this machine)'; }
  else if (!port) { portAvailable = false; portEvidence = 'no_port_published(pipe-only)'; }
  else if (opts.portAvailable !== undefined) { portAvailable = !!opts.portAvailable; portEvidence = 'caller_supplied'; }
  else if (probeAnswered) { portAvailable = true; portEvidence = 'addin_answered_on_that_port'; }
  else {
    try { portAvailable = await tcpListening(port, 200); } catch (e) { portAvailable = false; }
    portEvidence = 'tcp_connect_test';
  }
  // A remote instance port belongs to ANOTHER machine: the OS listener table here describes a
  // DIFFERENT process that may reuse the number, so it is never consulted.
  const po = remote ? { available: false, pids: null, why: 'remote_endpoint' } : portOwnerPidsOf(port);
  return ownershipVerdictOf({
    pid: pid, port: port, remote: remote, inMemory: inMemory,
    bridge: (inst && (inst.bridge || inst.endpoint)) || null,
    startedByHub: !!(inst && inst.startedByHub),
    registryPresent: registryPresent, registryStartedByHub: registryStartedByHub,
    registeredPid: registeredPid,
    registryFile: (inst && inst.registryFile) || (reg && reg.file ? path.basename(reg.file) : null),
    pidAlive: pidAliveNow, portAvailable: portAvailable, portEvidence: portEvidence,
    portOwnerPids: po.pids, portOwnerWhy: po.why, netstatAvailable: po.available, probeAnswered: probeAnswered,
    project: (inst && inst.project) || (reg && reg.project) || null,
    projectPath: (inst && inst.projectPath) || (reg && reg.projectPath) || null
  });
}

// ONE liveness verdict per instance.  probeStatus is DERIVED from it, never taken from the
// registry alone, and a fresh add-in answer is required before "responding" is claimed.
function forwardRouteOf(i) {
  if (!i) return null;
  if (i.remote) return i.endpointUrl || ('endpoint:' + (i.endpoint || '?'));
  if (i.port) return 'http://127.0.0.1:' + i.port + ' (add-in /rpc)';
  return BRIDGE_MODE === 'inproc'
    ? 'none (inproc: a pipe-only instance is unreachable by design)'
    : (BRIDGE_URL + ' -> named pipe (the middle service is the only addresser)');
}
// A local instance is DIRECTLY addressable only when it publishes a REAL TCP port.  port 0 / null
// means the add-in registered a PIPE-ONLY instance (transport:pipe, tcpEnabled:false): the middle
// service is then the ONLY addresser, because the pipe itself is not a URL.  Building a URL from
// such a port produces "http://127.0.0.1:null/rpc" - a URL parse failure that HID the real cause
// (the middle service was not running) and surfaced as noIde.  This is the ONE place that decides
// what a pick can be reached over.
function directEndpointOf(inst) {
  if (!inst) return { kind: 'none', base: null, port: null };
  const port = Number(inst.port);
  if (inst.remote) return { kind: 'remote', base: inst.endpointUrl || null, port: (Number.isFinite(port) && port > 0) ? port : null };
  if (Number.isFinite(port) && port > 0) return { kind: 'addin', base: 'http://127.0.0.1:' + port, port: port };
  return { kind: 'pipeOnly', base: null, port: null };
}
// FAIL-CLOSED self-check: an outward payload must NEVER carry a null/undefined port URL.  A
// pipe-only IDE has no TCP endpoint, so such a URL can only be a bug - and it must not reach a
// caller even if some future code path builds one.  Returns a replacement body when it would.
const NULL_PORT_URL_RE = /(?:https?:\/\/[^\s"'<>]*?|127\.0\.0\.1):(?:null|undefined)(?![A-Za-z0-9_-])/;
function nullPortUrlGuard(body, where) {
  let text = '';
  try { text = JSON.stringify(body); } catch (e) { text = String(body); }
  if (!NULL_PORT_URL_RE.test(text)) return null;
  return { ok: false, reason: 'internal_null_port_url', failureKind: 'toolError',
           failureKindRule: FAILURE_KIND_RULE, where: String(where || 'unknown'),
           bridgeUrl: BRIDGE_URL, bridgeMode: BRIDGE_MODE,
           message: 'FAIL-CLOSED self-check: refused to emit an outward payload that contains a null/undefined port URL. A pipe-only IDE has NO TCP endpoint; the middle service (VB6MCP_BRIDGE_URL) is its only addresser.',
           hint: 'this means the hub tried to build a direct add-in URL for an instance that has no port. Nothing was forwarded; start the middle service (or expose a real add-in TCP port) and retry.' };
}
// The registration facts the pipe-only refusal must NAME (file, port, transport, tcpEnabled, pipe).
function registrySnapshotFor(pid) {
  try {
    const want = Number(pid);
    const r = readRegistry();
    const hit = (r.entries || []).filter(function (e) { return Number(e.pid) === want; })[0] || null;
    if (!hit) return null;
    // readRegistry() normalises a pipe registration's port 0 to null; report the RAW registration
    // values too, so the receipt can say "port:0 / tcpEnabled:false" instead of just "no port".
    let rawPort = hit.port, rawTcp = hit.tcpEnabled;
    try {
      const raw = JSON.parse(fs.readFileSync(hit.file, 'utf8'));
      if (raw && raw.port !== undefined) rawPort = raw.port;
      if (raw && raw.tcpEnabled !== undefined) rawTcp = raw.tcpEnabled;
    } catch (e) { }
    return { file: hit.file ? path.basename(hit.file) : null, dir: r.dir, pid: hit.pid,
             port: rawPort, portNormalized: hit.port, transport: hit.transport, tcpEnabled: rawTcp,
             pipe: hit.pipe, project: hit.project, projectPath: hit.projectPath, startedByHub: hit.startedByHub };
  } catch (e) { return null; }
}
// A picked instance with NO TCP port can only be served by the middle service.  When that service
// is not answering, the call FAILS HERE - structured.  The reason names the real cause, and the
// receipt carries the registration plus the SAME liveness evidence ide_list already reports
// (probeRoute / forwardRoute / probe reason), so a caller never sees a bare "noIde" again.
function pipeOnlyRefusal(pick, name, snap, note) {
  const reg = registrySnapshotFor(pick && pick.pid);
  let lv = null;
  try { lv = (pick && pick.liveness) || livenessOf(pick, bridgeIdentitySync() || {}); } catch (e) { lv = null; }
  const regPort = reg ? reg.port : ((pick && pick.port === undefined) ? null : pick.port);
  const regTransport = (reg && reg.transport) || (pick && pick.transport) || null;
  const regTcp = reg ? reg.tcpEnabled : ((pick && pick.tcpEnabled === undefined) ? null : pick.tcpEnabled);
  const body = {
    ok: false, tool: name,
    reason: 'middle_service_unreachable',
    reasonLegacy: 'forward_failed',
    failureKind: 'pipeOnlyNeedsMiddleService',
    failureKindRule: FAILURE_KIND_RULE,
    pipeOnly: true,
    pid: (pick && pick.pid) || null,
    port: (pick && pick.port === undefined) ? null : pick.port,
    project: (pick && pick.project) || null,
    projectPath: (pick && pick.projectPath) || (reg && reg.projectPath) || null,
    pipe: (reg && reg.pipe) || (pick && pick.pipe) || null,
    bridgeUrl: BRIDGE_URL,
    bridgeMode: BRIDGE_MODE,
    bridgePort: bridgePortHint(),
    bridgeLastFailure: bridgeLastFailure || null,
    bridgeHealthAttemptedAt: bridgeHealthCache.at || null,
    middleService: middleServiceReceipt(),
    hubCode: hubCodeStamps(),
    registry: reg,
    registryReason: (lv && lv.probeStatus === 'failed') ? 'registry_port_not_mcp_probe_failed' : 'pipe_only_registration',
    probeRoute: (lv && lv.probeRoute) || (pick && pick.probeRoute) || null,
    forwardRoute: (lv && lv.forwardRoute) || forwardRouteOf(pick),
    probeStatus: (lv && lv.probeStatus) || (pick && pick.probeStatus) || null,
    probeReason: (lv && lv.probeReason) || (pick && pick.probeReason) || null,
    liveness: lv ? { state: lv.state, basis: lv.basis, probeStatus: lv.probeStatus, probeReason: lv.probeReason,
                     probeRoute: lv.probeRoute, forwardRoute: lv.forwardRoute, pipeState: lv.pipeState,
                     pipeAvailable: lv.pipeAvailable } : null,
    message: 'this IDE is registered PIPE-ONLY (registration port ' + JSON.stringify(regPort) + ', transport ' + JSON.stringify(regTransport) + ', tcpEnabled ' + JSON.stringify(regTcp) + '): the hub cannot reach it over TCP, and the middle service (bridge/vb6mcp_bridge.py) at ' + BRIDGE_URL + ' is the ONLY addresser for it - and it is NOT answering. A direct URL would have to name a null port, which is not a real endpoint, so NO forward was attempted.',
    hint: 'start the middle service and point VB6MCP_BRIDGE_URL at its port, then retry. Nothing was forwarded. If this add-in is TCP-capable, expose a real port (tcpEnabled) and it becomes directly addressable.',
    nothingForwarded: true
  };
  const guarded = nullPortUrlGuard(body, 'pipe_only_refusal(' + String(name) + ')');
  const finalBody = guarded || body;
  const ide = { pid: (pick && pick.pid) || null, project: (pick && pick.project) || null, port: null };
  return stampBody({ content: [{ type: 'text', text: JSON.stringify(finalBody, null, 2) }], isError: true },
                   ide, snap.contractVersion, note);
}
function livenessOf(i, ctx) {
  const byPid = (ctx && ctx.byPid) || {};
  // The LOCAL middle service's per-pid facts describe THIS machine's instances only: applying them
  // to a REMOTE candidate would be cross-bridge contamination (and it WOULD collide when two
  // bridges publish the same pid, which is exactly what the multi-bridge cards must support).
  const bi = (i && !i.remote && i.pid && byPid[String(i.pid)]) || null;
  const pipeState = bi ? (bi.pipeState || null) : ((i && i.pipeState) || null);
  const pipeAvailable = bi ? !!bi.pipeAvailable : ((i && i.pipeAvailable !== undefined && i.pipeAvailable !== null) ? !!i.pipeAvailable : null);
  const rawStatus = (i && i.probeStatus) || ((i && i.pending) ? 'pending' : 'ok');
  const evidenceAt = (i && (i.probeEvidenceAt || i.evidenceAt)) || null;
  const ageMs = evidenceAt ? (Date.now() - evidenceAt) : null;
  const route = { probe: (i && i.probeRoute) || null, forward: forwardRouteOf(i) };
  let state, basis, status = rawStatus, probeReason = (i && i.probeReason) || null;
  if (i && String(i.transport || '') === 'pipe' && pipeAvailable === false) {
    // The middle service - the ONLY addresser - says this instance's pipe is NOT connectable.
    // That is the same wall a forwarded call hits, so the face must say it too (N11).
    state = 'not_responding';
    basis = 'bridge_health:pipe_not_available(pipeState=' + (pipeState || '?') + ')';
    status = 'failed';
    probeReason = 'pipe_unavailable';
  } else if (rawStatus === 'ok') {
    // An ok verdict is only a LIVE ANSWER when it is DATEABLE and FRESH.  The pre-change hub also
    // inherited ok from a persisted/cached face (hub_face_cache.json) and re-published it with a
    // full liveToolCount, which is how ide_list ended up claiming a healthy IDE that had already
    // wedged (N11).  An undated or stale ok is reported as 'stale', never as 'ok'.
    if (ageMs === null) {
      state = 'unknown'; status = 'stale';
      basis = 'ok_without_evidence_timestamp(the verdict came from a cached/persisted face and cannot be dated, so it is NOT reported as a live answer)';
      probeReason = probeReason || 'stale_probe_verdict';
    } else if (ageMs > DISCOVERY_TTL_MS) {
      state = 'unknown'; status = 'stale';
      basis = 'stale_evidence(ageMs=' + ageMs + ' > discovery_ttl=' + DISCOVERY_TTL_MS + ')';
      probeReason = probeReason || 'stale_probe_verdict';
    } else {
      state = 'responding';
      basis = (i && i.probeBasis) || ('addin_answered(' + (route.probe || route.forward) + ')');
    }
  } else if (rawStatus === 'pending') {
    state = 'unknown'; basis = 'probe_pending';
  } else {
    state = 'not_responding'; basis = 'probe_failed(' + (probeReason || 'unknown') + ')';
  }
  const contradictions = [];
  if (rawStatus === 'ok' && state !== 'responding') {
    contradictions.push({ pid: i && i.pid, port: i && i.port, raw: rawStatus, verdict: state,
      explained_by: basis,
      note: 'the cached/registry verdict says ok but it is NOT backed by a fresh add-in answer, so this receipt reports it as ' + state + ' (role: ' + basis + ') instead of repeating the ok - a verdict and a forwarded call now agree by construction' });
  }
  if (rawStatus !== 'ok' && status === 'failed' && (i && i.liveTools && i.liveTools.length)) {
    contradictions.push({ pid: i && i.pid, port: i && i.port, raw: rawStatus, verdict: state, explained_by: basis,
      note: 'the face (tools/list) was readable but the ADD-IN did not answer, so liveToolCount is NOT reported as a live count' });
  }
  return { state: state, basis: basis, evidenceAt: evidenceAt, evidenceAgeMs: ageMs,
           probeStatus: status, probeStatusRaw: rawStatus, probeReason: probeReason, cached: !!(i && i.fromFaceCache),
           probeRoute: route.probe, forwardRoute: route.forward, pipeState: pipeState, pipeAvailable: pipeAvailable,
           contradictions: contradictions };
}
// The selector handed to the middle service when the hub resolved the target itself.  Both sides
// match pid | port | project NAME | projectPath; projectPath is the canonical identity, pid is the
// most precise per-instance address, so both are offered (path first, because it survives a pid
// recycle inside the middle service's own view).
function canonicalSelectorFor(pick) {
  if (!pick || !pick.ok) return '';
  if (pick.remote) return (pick.endpoint ? pick.endpoint + ':' : '') + (pick.projectPath || pick.pid || '');
  // The middle service's selector vocabulary is pid | port | project NAME | projectPath.  The PID is
  // the PREFERRED form: it is the registry key, it is unambiguous when one project is served twice,
  // and both sides already speak it.  projectPath is only the fallback for an instance whose pid is
  // unknown, and the port/name follow (the legacy aliases).
  if (pick.pid) return String(pick.pid);
  if (pick.projectPath) return String(pick.projectPath);
  if (pick.port) return String(pick.port);
  return String(pick.project || '');
}
function selectorFormOf(sel, pick) {
  const s = String(sel === undefined || sel === null ? '' : sel);
  if (!s) return 'none';
  if (pick && pick.pid && s === String(pick.pid)) return 'pid';
  if (pick && pick.projectPath && s.toLowerCase() === String(pick.projectPath).toLowerCase()) return 'projectPath';
  if (pick && pick.port && s === String(pick.port)) return 'port';
  if (pick && pick.project && s === String(pick.project)) return 'project_name';
  return 'verbatim_explicit';
}
// The middle service ANSWERED with a target refusal: that is an EXPLICIT failure, so the hub must
// say so and must never quietly retry another form (or fall back to picking an instance itself).
function targetRefusalOf(result) {
  const obj = firstBodyObject(result);
  if (!obj) return null;
  return (obj.reason === 'target_mismatch' || obj.reason === 'target_conflict' || obj.reason === 'no_ide_running') ? obj : null;
}
function defaultTargetLabel() {
  return (state.target === undefined || state.target === null || state.target === '') ? null : String(state.target);
}
let listeningCache = { at: 0, map: null };
function listeningPortsCached() {
  if (listeningCache.map && Date.now() - listeningCache.at < 2000) return listeningCache.map;
  const m = listeningPorts();
  listeningCache = { at: Date.now(), map: m };
  return m;
}
// the persisted face cache is keyed by the FACE FINGERPRINT (the same sha256
// tools/gen_hub_snapshot.py writes into bridge/hub_schema.generated.json).  A cache without a
// key cannot be invalidated: the un-keyed version kept re-publishing a face it had captured
// under a DIFFERENT frozen face, so ide_list could answer from a face that no longer exists.
// Rule: a cached entry is used ONLY when its fingerprint equals the CURRENT snapshot's.
// When the snapshot carries no fingerprint (an older generated file) the cache is not used at
// all - a rebuild on start/reload is the documented fallback, never a silent reuse.
function faceFingerprintNow() {
  const s = snapshot();
  return (s && typeof s.faceFingerprint === 'string' && s.faceFingerprint) ? s.faceFingerprint : null;
}
function saveFaceCache(d) {
  try {
    // OPEN-27 (second half): the banner is NOT part of the persisted FACE.  initialize.instructions
    // belongs to the add-in BINARY that answered (contract + build.sha256), and a cache keyed only by
    // the frozen tool face cannot prove that identity; persisting the text let a matching fingerprint
    // replay a dead DLL's banner after a hub restart.  Only the FACE is persisted - the banner is
    // always re-fetched from the live add-in in THIS process (or the frozen snapshot banner is used
    // with the honest "add-in sent no instructions" wording).
    const persistedInstances = (Array.isArray(d.instances) ? d.instances : []).map(function (i) {
      const c = Object.assign({}, i);
      delete c.instructions;
      return c;
    });
    const rec = { at: Date.now(), faceFingerprint: faceFingerprintNow(), instances: persistedInstances, conflicts: d.conflicts, staleRegistry: d.staleRegistry, registryDir: d.registryDir };
    const prev = readJson(FACE_CACHE, null);
    if (prev && prev.faceNotify) rec.faceNotify = prev.faceNotify;   // the change record outlives a discovery write
    fs.writeFileSync(FACE_CACHE, JSON.stringify(rec, null, 2), 'utf8');
  } catch (e) { }
}
let faceCacheStatus = { reason: 'not_read' };
function loadFaceCache() {
  const c = readJson(FACE_CACHE, null);
  if (!c || !Array.isArray(c.instances) || !c.instances.length) {
    faceCacheStatus = { reason: c ? 'empty' : 'absent' };
    return null;
  }
  const want = faceFingerprintNow();
  if (!want) {
    faceCacheStatus = { reason: 'snapshot_has_no_faceFingerprint', cached: c.faceFingerprint || null,
      note: 'the snapshot cannot key a cache; the face is rebuilt on start/reload instead of reused' };
    return null;
  }
  if (c.faceFingerprint !== want) {
    faceCacheStatus = { reason: 'face_fingerprint_mismatch', cached: c.faceFingerprint || null, current: want,
      note: 'the persisted face was captured under a different face fingerprint, so it is INVALID' };
    return null;
  }
  faceCacheStatus = { reason: 'hit', faceFingerprint: want };
  return c;
}

// ---- tools face change -> notifications/tools/list_changed --------------------------------
// A client keeps its tool registry until it is TOLD that the face changed, and a notification is
// useful only to a client that is ALREADY initialized (its listChanged callback registered) and
// whose own table is stale.  Two failures shaped this code:
//   * a notification sent during hub startup - before the client had finished initialize - looked
//     successful in the hub log and was dropped by the client; the old code then recorded it as
//     "announced" and never retried, so the client stayed at an empty/stale table;
//   * a single global face record cannot describe several client generations: the baseline must
//     belong to THE CONNECTION, seeded from the cross-process record and refreshed whenever that
//     connection pulls tools/list.
// So the order is: per-connection knownFace + an explicit re-align AFTER notifications/initialized
// (the first moment the callback is guaranteed live), and the persisted record is a cross-process
// BASELINE only - never evidence that this connection has seen anything.  A record written by the
// old startup-notify shape is deliberately IGNORED as a baseline (it proves a send, not a sync).
//
// Fingerprint = the name set plus a sort-stable hash of name/description/inputSchema for every
// tool: exactly the bytes a client would have to re-learn.  In inproc mode NOTHING here runs:
// no capability is added, no timer is armed, no notification is sent and no record is written.
const FACE_PROBE_MIN_MS = Math.max(2000, parseInt(process.env.VB6MCP_FACE_PROBE_MIN_MS || '2000', 10));
const FACE_NOTIFY_MIN_MS = Math.max(2000, parseInt(process.env.VB6MCP_FACE_NOTIFY_MIN_MS || '2000', 10));
let connSeq = 0;
// ONE stdio connection per process, but the state is per-connection on purpose: a second
// initialize resets it, so a reconnect starts from a NEW baseline instead of a reused verdict.
const conn = { id: 0, initialized: false, knownFace: null };
let lastFaceFp = null;             // cross-process baseline: last face a client was known to hold
let pendingFaceChange = null;      // {fp, why}: observed, not announced to THIS connection yet
let lastFaceProbeAt = 0;
let lastFaceNotifyAt = 0;
let faceProbeInFlight = null;
let faceFlushTimer = null;
let facePeriodicTimer = null;

function canonicalJson(v) {
  if (v === null) return 'null';
  if (Array.isArray(v)) return '[' + v.map(canonicalJson).join(',') + ']';
  const t = typeof v;
  if (t === 'undefined') return 'null';
  if (t === 'string' || t === 'number' || t === 'boolean') return JSON.stringify(v);
  if (t === 'object') {
    const keys = Object.keys(v).sort();
    return '{' + keys.map(function (k) { return JSON.stringify(k) + ':' + canonicalJson(v[k]); }).join(',') + '}';
  }
  return JSON.stringify(String(v));
}
function faceFingerprintOf(tools) {
  const rows = (Array.isArray(tools) ? tools : [])
    .filter(function (t) { return t && typeof t.name === 'string'; })
    .map(function (t) { return { name: t.name,
                                 description: (typeof t.description === 'string' ? t.description : ''),
                                 inputSchema: (t.inputSchema === undefined ? null : t.inputSchema) }; })
    .sort(function (a, b) { return a.name < b.name ? -1 : (a.name > b.name ? 1 : 0); });
  const hash = crypto.createHash('sha256').update(canonicalJson(rows)).digest('hex');
  return { fingerprint: hash, toolCount: rows.length };
}
function readFaceNotifyRecord() {
  const c = readJson(FACE_CACHE, null);
  const n = (c && c.faceNotify && typeof c.faceNotify === 'object') ? c.faceNotify : null;
  if (!n) return null;
  // ONLY the new shape carries a trustworthy baseline; the old shape recorded a SEND, not a sync.
  const baseline = (n.baseline && typeof n.baseline.fingerprint === 'string') ? n.baseline : null;
  return { baseline: baseline, lastGlobalNotify: n.lastGlobalNotify || null };
}
function writeFaceNotifyRecord(fields) {
  // fields: {reason, fingerprint, toolCount, delivered}
  const prev = readFaceNotifyRecord();
  const at = new Date().toISOString();
  const rec = { at: at,
    baseline: { at: at, reason: fields.reason, fingerprint: fields.fingerprint,
                toolCount: fields.toolCount, connId: conn.id },
    lastGlobalNotify: fields.delivered
      ? { at: at, reason: fields.reason, fingerprint: fields.fingerprint, toolCount: fields.toolCount, connId: conn.id }
      : (prev && prev.lastGlobalNotify) || null };
  const tmp = FACE_CACHE + '.tmp-' + process.pid;
  try {
    const cur = readJson(FACE_CACHE, null);
    if (!cur || typeof cur !== 'object') { dbg('face notify: hub_face_cache.json unreadable -> the record was not persisted'); return; }
    cur.faceNotify = rec;
    fs.writeFileSync(tmp, JSON.stringify(cur, null, 2), 'utf8');
    fs.renameSync(tmp, FACE_CACHE);      // atomic: a concurrent reader never sees a half file
  } catch (e) {
    try { fs.unlinkSync(tmp); } catch (e2) { }
    dbg('face notify: record write failed: ' + String((e && e.message) || e));
  }
}
// bridge-mode face read: toolsListResult() answers from the middle service and NEVER
// reads its `d` argument (it returns at its own bridge branch), so awaiting discovery here bought
// nothing and could stall one face tick by up to LIST_GRACE_MS (400 ms). inproc keeps the OLD path
// byte for byte: there toolsListResult really does consume d (resolveDiscovery(d)).
async function currentFaceFingerprint() {
  const t0 = Date.now();
  const d = BRIDGE_MODE === 'inproc' ? await discoverSoft(LIST_GRACE_MS) : null;
  const res = await toolsListResult(d);
  const tools = (res && res.tools) || [];
  noteFaceProbe(Date.now() - t0, tools);
  return faceFingerprintOf(tools);
}
// rate-limited self-proof for the face read - ONE line per window, never
// one per call (per-call logging would manufacture the very traffic it is meant to measure).
//   [hub-face-probe] windowMs=60000 calls=30 ok=30 fail=0 avgMs=312 maxMs=340 tools=120 bytes=41876
// It is the hub-side evidence that a request/response pair really completed, so an add-in counter
// that stays 0 (patch-O gap) can never be misread as "no request was made".
// Window: VB6MCP_FACE_PROBE_LOG_MS (default 60000, floor 10000).
const FACE_PROBE_LOG_MS = Math.max(10000, parseInt(process.env.VB6MCP_FACE_PROBE_LOG_MS || '60000', 10));
const faceProbe = { at: 0, calls: 0, fail: 0, totalMs: 0, maxMs: 0, lastTools: 0, lastBytes: 0 };
function noteFaceProbe(ms, tools) {
  const now = Date.now();
  if (!faceProbe.at) faceProbe.at = now;
  faceProbe.calls += 1;
  if (!tools.length) faceProbe.fail += 1;
  faceProbe.totalMs += ms;
  if (ms > faceProbe.maxMs) faceProbe.maxMs = ms;
  faceProbe.lastTools = tools.length;
  try { faceProbe.lastBytes = JSON.stringify(tools).length; } catch (e) { faceProbe.lastBytes = 0; }
  const windowMs = now - faceProbe.at;
  if (windowMs < FACE_PROBE_LOG_MS) return;
  const avg = faceProbe.calls ? Math.round(faceProbe.totalMs / faceProbe.calls) : 0;
  const line = '[hub-face-probe] windowMs=' + windowMs + ' calls=' + faceProbe.calls +
    ' ok=' + (faceProbe.calls - faceProbe.fail) + ' fail=' + faceProbe.fail +
    ' avgMs=' + avg + ' maxMs=' + faceProbe.maxMs +
    ' tools=' + faceProbe.lastTools + ' bytes=' + faceProbe.lastBytes;
  try { process.stderr.write(line + '\n'); } catch (e) { }
  dbg(line);
  faceProbe.at = now; faceProbe.calls = 0; faceProbe.fail = 0; faceProbe.totalMs = 0; faceProbe.maxMs = 0;
}
// The client just RECEIVED tools/list on this connection: that face is what it now holds, so this
// is the only event that may advance the per-connection AND the cross-process baseline.
function noteFaceSynced(tools) {
  try {
    const fp = faceFingerprintOf(tools || []);
    conn.knownFace = fp.fingerprint;
    lastFaceFp = fp.fingerprint;
    pendingFaceChange = null;
    lastFaceProbeAt = Date.now();
    writeFaceNotifyRecord({ reason: 'served', fingerprint: fp.fingerprint, toolCount: fp.toolCount, delivered: false });
    return fp;
  } catch (e) { return null; }
}
function flushFaceChange() {
  if (BRIDGE_MODE === 'inproc' || !pendingFaceChange) return;
  if (!conn.initialized) return;               // callback not guaranteed live yet: keep it pending
  const wait = (lastFaceNotifyAt || 0) + FACE_NOTIFY_MIN_MS - Date.now();
  if (wait > 0) {                              // one change = one notification, >= 2 s apart
    if (!faceFlushTimer) {
      faceFlushTimer = setTimeout(function () { faceFlushTimer = null; flushFaceChange(); }, wait);
      if (faceFlushTimer.unref) faceFlushTimer.unref();
    }
    return;
  }
  const p = pendingFaceChange;
  pendingFaceChange = null;
  lastFaceNotifyAt = Date.now();
  try {
    send({ jsonrpc: '2.0', method: 'notifications/tools/list_changed' });
  } catch (e) {
    pendingFaceChange = p;                     // stdout is gone: a later probe may retry
    dbg('tools/list_changed send failed: ' + String((e && e.message) || e));
    return;
  }
  conn.knownFace = p.fp.fingerprint;
  lastFaceFp = p.fp.fingerprint;
  // faceProbe / fallbackAge say HOW this conclusion was reached and HOW OLD it is: a file-poll
  // verdict never claims the add-in answered (addInAnswers stays null / basis no_route).
  writeFaceNotifyRecord({ reason: p.why, fingerprint: p.fp.fingerprint, toolCount: p.fp.toolCount,
    delivered: true, faceProbe: p.fp.probe || 'file_poll',
    fallbackAge: p.fp.fallbackAge === undefined ? null : p.fp.fallbackAge,
    addInAnswers: null, addInAnswersBasis: 'no_route' });
  try {
    process.stderr.write('[hub-face] notifications/tools/list_changed sent conn=' + conn.id +
      ' reason=' + p.why + ' tools=' + p.fp.toolCount + ' fingerprint=' + p.fp.fingerprint.slice(0, 16) + '\n');
  } catch (e) { }
}
// ---- face change detection for LOCAL instances: registry FILE poll, ZERO pipe frames ----------
// A local instance's face can only change when (i) the binary is replaced (a NEW process => a new
// instance file / new pid), (ii) the contract moves (also a new binary), or (iii) an extension pack
// is enabled/disabled THROUGH this hub (that path marks the face dirty at the moment it happens).
// Reading the registry directory therefore detects (i) and (ii) without touching the pipe. This
// timer used to forward tools/list over the pipe every 2 s - a 145 KB response and ~330 ms of
// add-in work - just to learn whether those files had changed.
// LIMIT, stated honestly: case (iii) moves NO file this poll can see, so the poll cannot detect it,
// and this file contains NO full-hash fallback. With the R3-a result cache ENABLED
// (VB6MCP_HUB_FACE_CACHE=1) such an out-of-band change is found by that cache's 60 s max-age forced
// re-read and announced (face_changed_on_refresh -> notifications/tools/list_changed). With the
// cache OFF the hub never re-reads the face on its own: an out-of-band change becomes visible to a
// client only on the next tools/list that client makes for its own reasons (and is never announced).
const FACE_FILE_POLL = { at: 0, polls: 0, absent: 0 };
let faceDirtyNonce = 0;
// ---- 2b: dirty marker for a face change the HUB ITSELF caused ---------------------------------
// Call sites exist ONLY under the R3-a tools/list result cache (VB6MCP_HUB_FACE_CACHE=1), which is
// OFF by default: (1) a SUCCESSFULLY forwarded ext_enable/ext_disable/ext_reload/ext_install/
// ext_uninstall, (2) an "unknown tool" refusal from a forwarded tools/call, (3) an out-of-band face
// change discovered by the cache's 60 s max-age refresh. With the cache OFF this hook still has no
// call site: the file poll cannot see an out-of-band ext change and this file has NO full-hash
// fallback to hide behind. It bumps the key the cache validates on and unblocks the next poll
// immediately; the R3-a hooks then make the NEXT tools/list a real read.
function markFaceDirty(reason) {
  faceDirtyNonce += 1;
  lastFaceProbeAt = 0;          // do not let the 2 s throttle swallow the immediate re-check
  const line = '[hub-face-dirty] reason=' + (reason || 'unspecified') + ' nonce=' + faceDirtyNonce;
  try { process.stderr.write(line + '\n'); } catch (e) { }
  dbg(line);
  return faceDirtyNonce;
}
function localFaceKey() {
  const dir = instancesDir();
  let names = [];
  try {
    names = fs.readdirSync(dir).filter(function (n) { return /\.json$/i.test(n); }).sort();
  } catch (e) {
    return { key: null, rows: [], absent: true, note: 'readdir_failed:' + String((e && e.message) || e) };
  }
  const rows = [];
  for (const n of names) {
    const p = path.join(dir, n);
    try {
      const st = fs.statSync(p);
      const rec = JSON.parse(fs.readFileSync(p, 'utf8'));
      const alive = pidAlive(rec.pid);
      rows.push({ file: n, pid: rec.pid === undefined ? null : rec.pid,
                  instanceId: rec.instanceId || null, dllSha256: rec.dllSha256 || null,
                  project: rec.project || null, mtimeMs: st.mtimeMs,
                  alive: alive, absent: alive === false });
    } catch (e) {
      rows.push({ file: n, error: String((e && e.message) || e) });
    }
  }
  return { key: crypto.createHash('sha256').update(canonicalJson(rows)).digest('hex'),
           rows: rows, absent: false };
}
function hubFaceSourceKey() {
  const parts = [];
  for (const p of [HUBCFG, SNAP]) {
    try { parts.push(path.basename(p) + ':' + fs.statSync(p).mtimeMs); } catch (e) { parts.push(path.basename(p) + ':missing'); }
  }
  try { parts.push('ext:' + fs.readdirSync(path.join(HERE, 'mcp_extensions')).sort().join(',')); } catch (e) { parts.push('ext:none'); }
  return { key: crypto.createHash('sha256').update(parts.join('|')).digest('hex'), parts: parts };
}
// ---- R3-a: per-process tools/list RESULT cache (VB6MCP_HUB_FACE_CACHE, DEFAULT OFF) ----------
// MEASURED (fix/TASK_PERFIDLE_20261022_1545/MEASURED_CALL_COST.md): the add-in's own render cache
// already hits 100% (toolsCacheHits == the tools/list frame count, toolsCacheMisses == 0), so the
// cost that remains is the FRAME itself: 145,123 B written into the pipe (segPipe 9.27 ms = 92% of
// the call) + ~0.85 ms of work. The only way to remove that cost is to not send the frame - once
// this process has really built the face, the SAME bytes can be answered here until something
// that CAN change the face happens. The hub already owns every ingredient: localFaceKey() and
// hubFaceSourceKey() fingerprint the registry and the hub config/snapshot/extension files with
// ZERO pipe frames, and markFaceDirty() was left with the comment "no call site yet" - this block
// is that call site, not a second change-detection mechanism.
//
// SAFETY CONTRACT:
//   * DEFAULT OFF. VB6MCP_HUB_FACE_CACHE unset/0 => handle() runs the PRE-CHANGE expression
//     toolsListResult(await discoverSoft(LIST_GRACE_MS)); no cache state is read, no extra pipe
//     frame, no extra file write, no extra log line. That is the byte-identical rollback path.
//   * ENABLED: a cached face is served only when ALL of these hold:
//       1) this process built it from a REAL bridge face (a snapshot/fallback face is NEVER cached);
//       2) localFaceKey(), hubFaceSourceKey() and faceDirtyNonce are unchanged since that build;
//       3) it is younger than VB6MCP_HUB_FACE_CACHE_MAX_AGE_MS (default 60000 ms) - the 60 s
//          full-refresh safety net the face-change comment above used to describe;
//       4) the one-time self-check did not fail (a mismatch PERMANENTLY falls back for the process).
//   * A HIT answers with no pipe frame, no discovery pass and NO notifications/tools/list_changed
//     (the face did not change, so there is nothing to announce).
//   * FORCED REFRESH (the next tools/list goes down the real path):
//       (1) any ext_enable/ext_disable/ext_reload/ext_install/ext_uninstall that FORWARDED OK;
//       (2) an "unknown tool" refusal from a forwarded tools/call (cache and add-in disagree);
//       (3) any face-key change (IDE registration added/removed, binary replaced, hub_tools.json /
//           hub_schema.generated.json rewritten, mcp_extensions changed);
//       (4) the max-age deadline above.
//   * SELF-CHECK: the FIRST build is read back from the add-in a SECOND time and compared as
//     canonical JSON (the StrComp rule). Equal => the cache is trusted and ONE
//     [hub-face-cache] {"event":"selfcheck_ok"...} line is logged. Different => the cache is
//     DISABLED for the rest of the process and ONE selfcheck_failed line carries both sha256s.
//     Nothing is thrown and the client's response is never affected (the check runs AFTER the
//     client has been answered).
//   * Everything here is INERT in BRIDGE_MODE=inproc (that rollback path stays byte-identical).
const FACE_RESULT_CACHE_ENV = String(process.env.VB6MCP_HUB_FACE_CACHE || '').trim().toLowerCase();
const FACE_RESULT_CACHE_ON = FACE_RESULT_CACHE_ENV === '1' || FACE_RESULT_CACHE_ENV === 'on'
  || FACE_RESULT_CACHE_ENV === 'true' || FACE_RESULT_CACHE_ENV === 'yes';
const FACE_RESULT_CACHE_AGE_ENV = parseInt(process.env.VB6MCP_HUB_FACE_CACHE_MAX_AGE_MS || '', 10);
const FACE_RESULT_CACHE_MAX_AGE_MS = (isFinite(FACE_RESULT_CACHE_AGE_ENV) && FACE_RESULT_CACHE_AGE_ENV >= 0)
  ? FACE_RESULT_CACHE_AGE_ENV : 60000;
const FACE_RESULT_CACHE_LOG_ENV = parseInt(process.env.VB6MCP_HUB_FACE_CACHE_LOG_MS || '', 10);
const FACE_RESULT_CACHE_LOG_MS = (isFinite(FACE_RESULT_CACHE_LOG_ENV) && FACE_RESULT_CACHE_LOG_ENV >= 1000)
  ? FACE_RESULT_CACHE_LOG_ENV : 60000;
const FACE_RESULT_CACHE_ENABLED = FACE_RESULT_CACHE_ON && BRIDGE_MODE !== 'inproc';
const FACE_EXT_TOOL_NAMES = ['ext_enable', 'ext_disable', 'ext_reload', 'ext_install', 'ext_uninstall'];
const FACE_UNKNOWN_TOOL_RE = /unknown\s*tool|tool\s*(?:is\s*)?not\s*found|no\s+such\s*tool|unrecognized\s*tool|method\s*not\s*found|未知工具|工具不存在/i;
// ONE entry per process, NEVER persisted: a cross-process face cache is a different feature (it
// would have to prove that the OTHER process saw the same face; this one only replays the bytes
// THIS process already received).
let faceResultCache = {
  built: false, at: 0,
  result: null, body: '', bodySha256: null,
  keyLocal: null, keySource: null, nonce: 0,
  builds: 0, hits: 0, misses: 0, forced: 0,
  disabled: null,                    // {reason, at, ...}: PERMANENT fallback after a failed self-check
  verify: { state: 'not_run', attempts: 0, at: null, cachedSha256: null, realSha256: null,
            cachedTools: null, realTools: null, cachedBytes: 0, realBytes: 0 },
  throttle: {}, win: { at: 0, hits: 0, misses: 0, forced: 0, builds: 0 }
};
function faceCacheSha256(text) {
  return crypto.createHash('sha256').update(String(text), 'utf8').digest('hex');
}
function faceCacheKeysNow() {
  let local = null, source = null, err = null;
  try { local = localFaceKey(); } catch (e) { err = 'localFaceKey: ' + String((e && e.message) || e); }
  try { source = hubFaceSourceKey(); } catch (e) { err = (err ? err + '; ' : '') + 'hubFaceSourceKey: ' + String((e && e.message) || e); }
  return { local: local ? local.key : null, source: source ? source.key : null,
           localAbsent: !!(local && local.absent), localRows: local ? local.rows.length : null,
           sourceParts: source ? source.parts : null,
           nonce: faceDirtyNonce, error: err };
}
function faceCacheStateNow() {
  const now = Date.now();
  return { enabled: FACE_RESULT_CACHE_ENABLED, built: faceResultCache.built,
           ageMs: faceResultCache.built ? (now - faceResultCache.at) : null,
           maxAgeMs: FACE_RESULT_CACHE_MAX_AGE_MS,
           bytes: faceResultCache.body ? faceResultCache.body.length : 0,
           sha256: faceResultCache.bodySha256,
           builds: faceResultCache.builds, hits: faceResultCache.hits,
           misses: faceResultCache.misses, forced: faceResultCache.forced,
           selfCheck: faceResultCache.verify.state,
           disabled: faceResultCache.disabled ? faceResultCache.disabled.reason : null };
}
function faceCacheLog(event, fields) {
  if (!FACE_RESULT_CACHE_ENABLED) return;      // the default path must emit NOTHING new
  const line = '[hub-face-cache] ' + JSON.stringify(Object.assign({
    event: event, at: new Date().toISOString(), pid: process.pid, mode: BRIDGE_MODE,
    state: faceCacheStateNow() }, fields || {}));
  try { process.stderr.write(line + '\n'); } catch (e) { }
  dbg(line);
}
function faceCacheLogThrottled(key, event, fields) {
  const now = Date.now();
  const last = faceResultCache.throttle[key] || 0;
  if (last && now - last < 60000) return;
  faceResultCache.throttle[key] = now;
  faceCacheLog(event, fields);
}
// ONE rate-limited window line, never one line per call (per-call logging would manufacture the
// very traffic this feature exists to remove). Same convention as [hub-face-probe].
function faceCacheNote(kind) {
  if (!FACE_RESULT_CACHE_ENABLED) return;
  const w = faceResultCache.win;
  const now = Date.now();
  if (!w.at) w.at = now;
  w[kind] = (w[kind] || 0) + 1;
  if (now - w.at < FACE_RESULT_CACHE_LOG_MS) return;
  faceCacheLog('window', { windowMs: now - w.at,
    window: { hits: w.hits, misses: w.misses, forced: w.forced, builds: w.builds } });
  w.at = now; w.hits = 0; w.misses = 0; w.forced = 0; w.builds = 0;
}
// {hit:true,result,ageMs} | {hit:false,why,...}. A miss NEVER mutates the entry: it only says why
// the next real read is required, and the real read then REPLACES the entry (store()).
function faceResultCacheLookup() {
  if (!FACE_RESULT_CACHE_ENABLED) return { hit: false, why: 'disabled' };
  if (faceResultCache.disabled) return { hit: false, why: 'permanently_disabled', detail: faceResultCache.disabled.reason };
  if (!faceResultCache.built) return { hit: false, why: 'cold' };
  const now = Date.now();
  const age = now - faceResultCache.at;
  const miss = function (why, extra) {
    faceResultCache.misses += 1; faceCacheNote('misses');
    return Object.assign({ hit: false, why: why }, extra || {});
  };
  if (FACE_RESULT_CACHE_MAX_AGE_MS > 0 && age >= FACE_RESULT_CACHE_MAX_AGE_MS) {
    return miss('max_age', { ageMs: age, maxAgeMs: FACE_RESULT_CACHE_MAX_AGE_MS, forced: true });
  }
  const k = faceCacheKeysNow();
  if (k.error) return miss('key_error', { detail: k.error });
  if (k.nonce !== faceResultCache.nonce) return miss('face_dirty_flagged', { forced: true, nonce: k.nonce, cachedNonce: faceResultCache.nonce });
  if (k.local !== faceResultCache.keyLocal) return miss('local_face_changed', { forced: true });
  if (k.source !== faceResultCache.keySource) return miss('hub_face_changed', { forced: true });
  faceResultCache.hits += 1; faceCacheNote('hits');
  return { hit: true, result: faceResultCache.result, ageMs: age };
}
function faceResultCacheStore(result, why) {
  if (!FACE_RESULT_CACHE_ENABLED || faceResultCache.disabled) return null;
  const k = faceCacheKeysNow();
  if (k.error) { faceCacheLogThrottled('store_key', 'not_cached', { why: 'key_error', detail: k.error }); return null; }
  let body;
  try { body = canonicalJson(result); }
  catch (e) {
    faceCacheLogThrottled('store_canon', 'not_cached', { why: 'canonicalize_failed', detail: String((e && e.message) || e) });
    return null;
  }
  const sha = faceCacheSha256(body);
  faceResultCache.built = true;
  faceResultCache.at = Date.now();
  faceResultCache.result = result;
  faceResultCache.body = body;
  faceResultCache.bodySha256 = sha;
  faceResultCache.keyLocal = k.local;
  faceResultCache.keySource = k.source;
  faceResultCache.nonce = k.nonce;
  faceResultCache.builds += 1;
  faceCacheNote('builds');
  faceCacheLog('built', { why: why || null, bytes: body.length, sha256: sha,
    tools: (result && Array.isArray(result.tools)) ? result.tools.length : 0,
    keyLocal: String(k.local || '').slice(0, 16), keySource: String(k.source || '').slice(0, 16),
    nonce: k.nonce, registryRows: k.localRows, localAbsent: k.localAbsent,
    hubFaceParts: k.sourceParts });
  if (faceResultCache.verify.state !== 'ok') scheduleFaceResultCacheSelfCheck();
  return { bytes: body.length, sha256: sha };
}
// A refresh the HUB ITSELF caused: bump the nonce the cache keys on (markFaceDirty) AND say so.
// markFaceDirty() also unblocks the 2 s registry file poll, so the change is announced to an
// initialized client and the NEXT tools/list really goes down the wire.
function faceResultCacheForcedRefresh(reason) {
  if (!FACE_RESULT_CACHE_ENABLED) return 0;
  faceResultCache.forced += 1;
  faceCacheNote('forced');
  faceCacheLog('forced_refresh', { reason: String(reason || 'unspecified'),
    built: faceResultCache.built,
    ageMs: faceResultCache.built ? (Date.now() - faceResultCache.at) : null });
  return markFaceDirty(reason);
}
// A MAX-AGE refresh can discover a face change that NO hub-visible file records (an extension
// loaded out of band, a binary swapped without rewriting its registration): the bytes changed, so
// the client must be TOLD instead of silently keeping the old registry until it happens to pull.
// A cache HIT never reaches this function - it only runs on a real refresh whose bytes differ.
function faceResultCacheMaybeAnnounce(prevBody, why) {
  if (!FACE_RESULT_CACHE_ENABLED || !faceResultCache.built) return null;
  if (!prevBody || faceResultCache.body === prevBody) return null;
  if (why !== 'max_age') return null;   // key changes / ext dirty are announced by the face poll
  faceCacheLog('face_changed_on_refresh', { why: why,
    oldSha256: faceCacheSha256(prevBody), newSha256: faceResultCache.bodySha256,
    oldBytes: prevBody.length, newBytes: faceResultCache.body.length,
    action: 'notifications/tools/list_changed will be sent (the client must re-pull tools/list)' });
  const nonce = markFaceDirty('tool_face_changed_on_max_age_refresh');
  faceResultCache.nonce = faceDirtyNonce;   // keep the freshly stored entry valid (the change is REAL)
  return nonce;
}
// "the add-in says this tool does not exist" => the cached face and the add-in DISAGREE.
// A tool whose DATA merely mentions the phrase must never invalidate: a refusal marker is required.
// NEVER JSON-parses the body: it scans the already-available content[].text strings (bounded), so
// the per-call cost stays a string search even for a megabyte-sized tool receipt.
function looksLikeUnknownTool(result) {
  try {
    if (!result || typeof result !== 'object') return false;
    const chunks = [];
    if (result.error && typeof result.error === 'object') {
      chunks.push(String(result.error.message || ''));
      if (result.error.data && typeof result.error.data === 'object') chunks.push(String(result.error.data.reason || ''));
    }
    const c = result.content;
    if (Array.isArray(c)) { for (const it of c) { if (it && typeof it.text === 'string') chunks.push(it.text.slice(0, 4096)); } }
    const text = chunks.join('\n');
    if (!text || !FACE_UNKNOWN_TOOL_RE.test(text)) return false;
    if (result.isError === true) return true;
    if (result.error && typeof result.error === 'object') return true;
    if (/"reason"\s*:\s*"?(?:unknown_tool|tool_not_found|no_such_tool|method_not_found)/i.test(text)) return true;
    return /"ok"\s*:\s*false/.test(text);
  } catch (e) { return false; }
}
// The FIRST build is verified against a SECOND real read. It runs AFTER the client has been
// answered (a promise, never a blocking step), so it cannot change the response the client sees.
let faceCacheSelfCheckInFlight = null;
function scheduleFaceResultCacheSelfCheck() {
  if (!FACE_RESULT_CACHE_ENABLED || faceResultCache.disabled) return null;
  if (faceCacheSelfCheckInFlight) return faceCacheSelfCheckInFlight;
  if (faceResultCache.verify.state === 'ok' || faceResultCache.verify.attempts >= 3) return null;
  if (!faceResultCache.built) return null;
  const snap = { body: faceResultCache.body, sha: faceResultCache.bodySha256,
                 tools: (faceResultCache.result && Array.isArray(faceResultCache.result.tools)) ? faceResultCache.result.tools.length : 0,
                 bytes: faceResultCache.body.length,
                 nonce: faceResultCache.nonce, builds: faceResultCache.builds,
                 keyLocal: faceResultCache.keyLocal, keySource: faceResultCache.keySource };
  faceResultCache.verify.attempts += 1;
  faceResultCache.verify.state = 'running';
  faceCacheSelfCheckInFlight = (async function () {
    try {
      await new Promise(function (r) { setTimeout(r, 5); });   // answer the client first
      const meta = {};
      let real = null;
      try { real = await toolsListResult(undefined, meta); }
      catch (e) {
        faceResultCache.verify.state = 'inconclusive';
        faceCacheLog('selfcheck_inconclusive', { why: 'real_read_failed', attempt: faceResultCache.verify.attempts,
          detail: String((e && e.message) || e) });
        return null;
      }
      if (meta.source !== 'bridge') {
        faceResultCache.verify.state = 'inconclusive';
        faceCacheLog('selfcheck_inconclusive', { why: 'bridge_offline_or_fallback', attempt: faceResultCache.verify.attempts,
          source: meta.source || null });
        return null;
      }
      // A face change that happened WHILE the second read was in flight makes this comparison
      // meaningless (two different faces).  That is NOT a self-check failure: only a comparison
      // taken against the SAME keys can prove the cache mechanism right or wrong.
      const kDuring = faceCacheKeysNow();
      const moved = (faceResultCache.body !== snap.body)
        || (faceResultCache.nonce !== snap.nonce)
        || (faceResultCache.builds !== snap.builds)
        || !!kDuring.error
        || (kDuring.nonce !== snap.nonce)
        || (kDuring.local !== snap.keyLocal)
        || (kDuring.source !== snap.keySource);
      if (moved) {
        faceResultCache.verify.state = 'inconclusive';
        faceCacheLog('selfcheck_inconclusive', { why: 'face_changed_during_check',
          attempt: faceResultCache.verify.attempts,
          nonce: kDuring.nonce, snapNonce: snap.nonce, builds: faceResultCache.builds, snapBuilds: snap.builds,
          keyChanged: (kDuring.local !== snap.keyLocal) || (kDuring.source !== snap.keySource) || (kDuring.nonce !== snap.nonce) });
        return null;
      }
      const realBody = canonicalJson(real);
      const realSha = faceCacheSha256(realBody);
      const realTools = (real && Array.isArray(real.tools)) ? real.tools.length : 0;
      faceResultCache.verify.at = new Date().toISOString();
      faceResultCache.verify.cachedSha256 = snap.sha;
      faceResultCache.verify.realSha256 = realSha;
      faceResultCache.verify.cachedTools = snap.tools;
      faceResultCache.verify.realTools = realTools;
      faceResultCache.verify.cachedBytes = snap.bytes;
      faceResultCache.verify.realBytes = realBody.length;
      if (realBody === snap.body) {
        faceResultCache.verify.state = 'ok';
        faceCacheLog('selfcheck_ok', { sha256: realSha, bytes: realBody.length, tools: realTools,
          cachedSha256: snap.sha,
          compared: 'canonical JSON byte equality: the cache body vs a SECOND real bridge tools/list' });
        return { ok: true, sha256: realSha };
      }
      // PERMANENT fallback: the bytes this process would replay are NOT the bytes the add-in serves.
      faceResultCache.disabled = { reason: 'selfcheck_mismatch', at: new Date().toISOString(),
        cachedSha256: snap.sha, realSha256: realSha,
        cachedTools: snap.tools, realTools: realTools, cachedBytes: snap.bytes, realBytes: realBody.length };
      faceResultCache.built = false;
      faceResultCache.result = null;
      faceResultCache.body = '';
      faceResultCache.bodySha256 = null;
      faceResultCache.verify.state = 'failed';
      faceCacheLog('selfcheck_failed', { cachedSha256: snap.sha, realSha256: realSha,
        cachedTools: snap.tools, realTools: realTools, cachedBytes: snap.bytes, realBytes: realBody.length,
        action: 'cache DISABLED for the rest of this process: every tools/list takes the real path',
        rule: 'the cached body must equal, after JSON normalization, a second real bridge tools/list' });
      return { ok: false, cachedSha256: snap.sha, realSha256: realSha };
    } catch (e) {
      faceResultCache.verify.state = 'error';
      faceCacheLog('selfcheck_error', { attempt: faceResultCache.verify.attempts, detail: String((e && e.message) || e) });
      return null;
    } finally {
      faceCacheSelfCheckInFlight = null;
    }
  })();
  return faceCacheSelfCheckInFlight;
}

// The tools/list answer FOR A CLIENT. With the cache off this is byte-for-byte the pre-change
// expression; with it on, a valid cached face is answered here WITHOUT touching the pipe.
async function toolsListForClient() {
  if (!FACE_RESULT_CACHE_ENABLED) {
    return { result: await toolsListResult(await discoverSoft(LIST_GRACE_MS)), cache: null };
  }
  const look = faceResultCacheLookup();
  if (look.hit) return { result: look.result, cache: { hit: true, ageMs: look.ageMs } };
  if (look.why === 'cold' || look.why === 'permanently_disabled') {
    faceCacheLogThrottled('cold', 'miss', { why: look.why, detail: look.detail || null });
  } else {
    faceCacheLog('miss', { why: look.why, forced: !!look.forced,
      ageMs: (look.ageMs === undefined ? null : look.ageMs), detail: look.detail || null,
      maxAgeMs: FACE_RESULT_CACHE_MAX_AGE_MS });
  }
  const meta = {};
  const prevBody = faceResultCache.built ? faceResultCache.body : null;
  const result = await toolsListResult(await discoverSoft(LIST_GRACE_MS), meta);
  if (meta.source === 'bridge') {
    const stored = faceResultCacheStore(result, look.why);
    if (stored) faceResultCacheMaybeAnnounce(prevBody, look.why);
  } else {
    faceCacheLogThrottled('notcached_bridge', 'not_cached', { why: look.why, source: meta.source || 'fallback',
      note: 'a snapshot/fallback face is NEVER cached: only a REAL bridge face may be replayed' });
  }
  return { result: result, cache: { hit: false, why: look.why, stored: meta.source === 'bridge' } };
}

function faceFingerprintFromRegistry() {
  const local = localFaceKey();
  const hubFace = hubFaceSourceKey();
  const fingerprint = crypto.createHash('sha256')
    .update(String(local.key) + '|' + hubFace.key + '|' + String(faceDirtyNonce)).digest('hex');
  return { fingerprint: fingerprint, toolCount: local.rows.length, probe: 'file_poll',
           fallbackAge: null, absent: local.absent, instances: local.rows };
}
// ONE rate-limited line per window, never one per poll: the stage-1 proof that an idle hub reads
// FILES instead of forwarding tools/list (pipeCalls=0, trigger=file_poll).
function noteFacePoll(fp) {
  const now = Date.now();
  if (!FACE_FILE_POLL.at) FACE_FILE_POLL.at = now;
  FACE_FILE_POLL.polls += 1;
  if (fp && fp.absent) FACE_FILE_POLL.absent += 1;
  const windowMs = now - FACE_FILE_POLL.at;
  if (windowMs < FACE_PROBE_LOG_MS) return;
  const line = '[hub-face-probe] windowMs=' + windowMs + ' polls=' + FACE_FILE_POLL.polls +
    ' trigger=file_poll pipeCalls=0 readdirFailed=' + FACE_FILE_POLL.absent +
    ' registryRows=' + ((fp && fp.toolCount !== undefined) ? fp.toolCount : '?') +
    ' fallbackAge=null';
  try { process.stderr.write(line + '\n'); } catch (e) { }
  dbg(line);
  FACE_FILE_POLL.at = now; FACE_FILE_POLL.polls = 0; FACE_FILE_POLL.absent = 0;
}
async function checkFaceChange(why, force) {
  if (BRIDGE_MODE === 'inproc') return null;
  if (faceProbeInFlight) return faceProbeInFlight;
  const now = Date.now();
  if (!force && now - lastFaceProbeAt < FACE_PROBE_MIN_MS) return null;
  lastFaceProbeAt = now;
  faceProbeInFlight = (async function () {
    let fp = null;
    try { fp = faceFingerprintFromRegistry(); }
    catch (e) { dbg('face file poll failed: ' + String((e && e.message) || e)); }
    noteFacePoll(fp);
    if (!fp) return null;
    if (lastFaceFp === null) { lastFaceFp = fp.fingerprint; return null; }   // first probe = baseline
    if (fp.fingerprint === lastFaceFp) return null;
    lastFaceFp = fp.fingerprint;
    pendingFaceChange = { fp: fp, why: why || 'change' };
    flushFaceChange();
    return fp;
  })();
  try { return await faceProbeInFlight; } finally { faceProbeInFlight = null; }
}
// notifications/initialized is the FIRST moment the client's listChanged callback is guaranteed
// live, so a face change that happened before/around initialize is announced exactly here.
async function alignClientFace() {
  if (BRIDGE_MODE === 'inproc') return;
  let fp = null;
  try { fp = await currentFaceFingerprint(); }
  catch (e) { dbg('client_init face probe failed: ' + String((e && e.message) || e)); }
  if (!fp) return;
  if (conn.knownFace === null) {          // nothing to compare yet: the client will pull the face
    conn.knownFace = fp.fingerprint;
    lastFaceFp = fp.fingerprint;
    return;
  }
  if (conn.knownFace === fp.fingerprint) { lastFaceFp = fp.fingerprint; return; }
  pendingFaceChange = { fp: fp, why: 'client_init' };
  flushFaceChange();
}
function newConnection() {
  connSeq += 1;
  conn.id = connSeq;
  conn.initialized = false;
  // Seed ONLY from the persisted cross-process baseline: that is the face the PREVIOUS client
  // generation held.  lastFaceFp may already be the LIVE face (the startup probe advanced it),
  // so using it here would hide exactly the race this function exists to repair.
  let seeded = null;
  try {
    const rec = readFaceNotifyRecord();
    if (rec && rec.baseline && typeof rec.baseline.fingerprint === 'string') seeded = rec.baseline.fingerprint;
  } catch (e) { }
  conn.knownFace = seeded;
  pendingFaceChange = null;
}
function bootFaceNotify() {
  if (BRIDGE_MODE === 'inproc') return;
  try {
    const rec = readFaceNotifyRecord();
    if (rec && rec.baseline && typeof rec.baseline.fingerprint === 'string') lastFaceFp = rec.baseline.fingerprint;
  } catch (e) { }
  checkFaceChange('startup').catch(function () { });   // baseline only: nothing is sent before initialize
}
function startFaceNotifyTimer() {
  if (BRIDGE_MODE === 'inproc' || facePeriodicTimer) return;
  facePeriodicTimer = setInterval(function () {
    checkFaceChange('periodic').catch(function () { });
  }, FACE_PROBE_MIN_MS);
  if (facePeriodicTimer.unref) facePeriodicTimer.unref();
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
  // NEVER guess C:\Windows here.  SystemRoot may be on ANOTHER drive, in which case
  // C:\Windows can be a stale second installation whose 32-bit binaries fail
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
// (second half): the hub's own copy of the registry FIELD CONTRACT.  Kept beside
// readRegistry() so the hub and the middle service can be diffed by eye: the acknowledged
// set is what the hub actually consumes, the ignored set is the only thing it may drop.
const REGISTRY_SYNTHETIC_FIELDS = ['file', 'projectName', 'path', 'sha256', 'registryUnknownFields'];
const REGISTRY_IGNORED_FIELDS = [];
const REGISTRY_ACKNOWLEDGED_FIELDS = [
  'pid', 'port', 'transport', 'pipe', 'tcpEnabled', 'sessionId', 'instanceId', 'nonce',
  'startedAt', 'startedByHub', 'dllSha256', 'version', 'project', 'projectPath',
  'pipeProto', 'hwnd', 'tid', 'epoch',
  'ingressMode', 'ingressHwnd', 'ingressBoxHwnd', 'ingressCaption'];
function registryUnknownFields(e) {
  return Object.keys(e).filter(function (k) {
    return REGISTRY_ACKNOWLEDGED_FIELDS.indexOf(k) < 0 &&
           REGISTRY_IGNORED_FIELDS.indexOf(k) < 0 &&
           REGISTRY_SYNTHETIC_FIELDS.indexOf(k) < 0;
  }).sort();
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
    // (second half): PASS THROUGH EVERYTHING the file published; this used to be a
    // hand-written whitelist of 14 keys, so the middle service's ingress identity, the pipe
    // self-registration record (hwnd/tid/epoch) and the wire stamp (pipeProto) were INVISIBLE
    // to the hub without any warning.  Nothing is dropped except the explicit list below, and
    // anything not acknowledged is NAMED in registryUnknownFields on the entry.
    const e = Object.assign({}, j);
    for (const d of REGISTRY_IGNORED_FIELDS) delete e[d];
    e.file = full;
    e.pid = parseInt(j.pid, 10) || null;
    e.port = parseInt(j.port, 10) || null;
    // the TRANSPORT is part of an instance's identity, not decoration.  A pipe-only IDE
    // publishes port 0 ON PURPOSE (the pipe IS the transport), so "no port" must never be read as
    // "no instance" - dropping this field here is what made every live pipe IDE look stale and sent
    // discovery off to the PORT SCAN, which then adopted an unrelated listener as a fake IDE.
    e.transport = j.transport || null;
    e.pipe = j.pipe || null;
    e.tcpEnabled = j.tcpEnabled === undefined || j.tcpEnabled === null ? null : !!j.tcpEnabled;
    e.sessionId = j.sessionId || null;
    e.instanceId = j.instanceId || null;
    e.project = j.project || j.projectName || null;
    e.projectPath = j.projectPath || j.path || null;
    e.startedAt = j.startedAt || null;
    e.version = j.version || null;
    e.dllSha256 = j.dllSha256 || j.sha256 || null;
    e.startedByHub = !!j.startedByHub;
    const unk = registryUnknownFields(e);
    if (unk.length) e.registryUnknownFields = unk;
    entries.push(e);
  }
  return { dir: dir, entries: entries, unreadable: false };
}

// One JSON-RPC POST against an absolute endpoint base ("http://host:port") - the SAME wire
// shape the local add-in uses, so a remote IDE is reached exactly like a local one.  An
// explicit path in the base (http://host:8796/rpc) wins over the add-in default.
// is this endpoint base the configured middle service?  Normalised the same way
// BRIDGE_URL is (trailing slashes stripped, case-insensitive) so an operator's spelling cannot
// silently take the un-authenticated path.
function isBridgeBase(base) {
  return !!base && String(base).replace(/\/+$/, '').toLowerCase() === BRIDGE_URL.toLowerCase();
}
// ---- credentials for a LOOPBACK endpoint that ASKS for them ------------------------
// The middle service that fronts pipe-only IDEs can be reached WITHOUT any configuration - the
// legacy LOCAL path finds it by PORT SCAN - and EVERY one of its surfaces requires the token.  The
// pre-change hub attached the credential only to the CONFIGURED bridge base (isBridgeBase), so a
// scanned-but-unconfigured service answered 401 (a pretty-printed body, the second half of the
// defect) and the caller saw a misleading answer instead of the IDE.
//
// The rule is SELF-LIMITING - it is NOT an identity guess and NOT "send the token to any 401":
//   * the endpoint must be LOOPBACK, so a credential never leaves this machine, AND
//   * the answer must be HTTP 401 carrying OUR OWN envelope reason "token_missing".
// Only then is the request repeated ONCE with the credential, and the verdict is remembered per
// endpoint base so the double POST happens at most once.  An add-in (or any other service) never
// answers token_missing, so the legacy path keeps issuing exactly the same single request.
const authVerdict = new Map();   // endpoint base -> true (this endpoint asked for OUR token)
function isLoopbackBase(base) {
  try {
    const h = String(new URL(base).hostname || '').toLowerCase().replace(/^\[|\]$/g, '');
    return h === '127.0.0.1' || h === 'localhost' || h === '::1' || h === '0:0:0:0:0:0:0:1';
  } catch (e) { return false; }
}
function asksForOurToken(base, status, raw) {
  if (status !== 401) return false;
  if (!BRIDGE_TOKEN) return false;                 // nothing to retry WITH
  if (!isLoopbackBase(base)) return false;         // never send it off this machine
  const j = parseRpcBody(raw);
  return !!(j && typeof j === 'object' && !Array.isArray(j) && j.reason === 'token_missing');
}
// ---- an AUTH failure is an ERROR, never a tool answer ------------------------------
// The middle service answers 401 with OUR OWN envelope (reason = token_missing | token_mismatch |
// token_in_url_rejected - see bridge/vb6mcp_bridge.py).  On the CONFIGURED bridge path that body
// used to be handed to the caller as a NORMAL tool result: resp.result is absent, so
// bridgeForwardCall fell through to JSON.stringify(resp) and answered isError=false with the 401
// page as the "add-in answer" - a WRONG/expired token looked like a successful call (measured).
// The detector is status + our own vocabulary, so it can never misfire on an add-in (which does
// not speak it); a non-2xx body that is NOT our envelope is returned to the caller unchanged.
const AUTH_REASONS = ['token_missing', 'token_mismatch', 'token_in_url_rejected', 'token_required'];
function authFailureEnvelope(status, raw) {
  const code = Number(status);
  if (code !== 401 && code !== 403) return null;
  const j = parseRpcBody(raw);
  if (!j || typeof j !== 'object' || Array.isArray(j)) return null;
  if (j.ok !== false) return null;
  return AUTH_REASONS.indexOf(String(j.reason || '')) >= 0 ? j : null;
}
function authFailureError(j, status, base) {
  const e = new Error('the middle service REFUSED the credential (reason=' + String(j.reason)
    + '): ' + String(j.message || 'no message'));
  e.code = 'bridge_auth_failed';
  e.status = Number(status);
  e.reason = String(j.reason);
  e.base = base;
  e.authHint = 'the tool was NEVER run; set VB6MCP_TOKEN (or VB6MCP_TOKEN_FILE) for THIS endpoint and retry';
  // ADDITIVE: name the CONFIGURED bridge and its credential facts (never the value), so a named
  // bridge's 401 is diagnosable without reading the hub's environment.  The local slot and the
  // legacy endpoint path keep the exact pre-change fields above.
  const perBridge = endpointAuthFor(base);
  if (perBridge && perBridge.bridge) {
    e.bridge = perBridge.label;
    e.bridgeAuth = { configured: perBridge.tokenConfigured, source: perBridge.source,
                     length: perBridge.length, fileConfigured: !!perBridge.tokenFile,
                     fileReadError: perBridge.tokenError || null };
    e.authHint = 'the tool was NEVER run; the credential of bridge "' + perBridge.label + '" was '
      + (perBridge.tokenConfigured ? 'REJECTED by that bridge' : 'NOT CONFIGURED (tokenFile missing/empty/unreadable)')
      + ' - check that bridge\'s tokenFile, then retry.';
  }
  return e;
}

// ---- ONE decoder for every JSON-RPC body this file reads ------------------------------------
// (measured on the local SP6 box, 2026-10-16): the middle service answers with
// PRETTY-PRINTED JSON (json.dumps(..., indent=2)), so the FIRST line of its body is just "{".  The
// old line-by-line reader took the first line that STARTED with "{" and ran JSON.parse("{") on it,
// which throws the V8 message
//   "Expected property name or '}' in JSON at position 1 (line 1 column 2)"
// - and every forwarded tool reported exactly that as reason=forward_failed / failureKind=wedged
// while a direct Invoke-WebRequest against the same endpoint (reading the WHOLE body) succeeded.
// The order below is ADAPTIVE: each step either yields a value or is abandoned, never thrown.
//   1. the WHOLE body - the compact single-line add-in answer (the pre-change result, byte for
//      byte) and the whole-body fallback the old code already had as its LAST statement;
//   2. SSE: every "data:" payload joined, then each payload on its own (the streamable-HTTP /mcp
//      face frames its JSON that way);
//   3. any single line that is valid JSON on its own (a pretty body, a preamble line before the
//      JSON, or a one-line SSE frame);
//   4. null - "this body carries no JSON at all"; the CALLER decides whether that is an error
//      (decodeRpcBody below keeps the pre-change error contract).
function parseRpcBody(raw) {
  const text = String(raw === undefined || raw === null ? '' : raw);
  if (!text.trim()) return null;
  try { return JSON.parse(text); } catch (e) { }
  const lines = text.split(/\r?\n/);
  const data = [];
  for (const line of lines) {
    const t = line.trim();
    if (t.indexOf('data:') === 0) data.push(t.slice(5).trim());
  }
  if (data.length) {
    // One SSE payload may itself be PRETTY JSON framed line by line ("data: {" / 'data: "a": 1' /
    // "data: }"), so the payloads are joined first; a single-frame stream is then handled by step 3.
    try { return JSON.parse(data.join('\n')); } catch (e) { }
  }
  for (const line of lines) {
    const t = line.trim();
    const s = t.indexOf('data:') === 0 ? t.slice(5).trim() : t;
    if (!s || (s[0] !== '{' && s[0] !== '[')) continue;   // the old "starts with {" rule, both shapes
    try { return JSON.parse(s); } catch (e) { }
  }
  return null;
}
// The pre-change ERROR contract, kept on purpose: an EMPTY body was null before and stays null,
// but a NON-EMPTY body that carries no JSON still THROWS.  Callers classify a failed call on that
// throw (transportFailureKind -> wedged; probeRemoteEndpoint -> endpoint_unreachable), so quietly
// returning null would turn "the peer answered with garbage" into "the peer answered, with no
// tools" - a false success.  Only the message is the hub's own now; the class (SyntaxError) and
// therefore every downstream classification is unchanged.
function decodeRpcBody(raw) {
  const j = parseRpcBody(raw);
  const text = String(raw === undefined || raw === null ? '' : raw);
  if (j === null && text.trim()) throw new SyntaxError('no JSON in the response body (' + text.length + ' bytes)');
  return j;
}

async function rpcAt(base, method, params, proto, timeoutMs, rpcPath, authLabel) {
  const target = endpointRpcTarget(base);
  const body = { jsonrpc: '2.0', id: (Date.now() % 100000) + 1, method: method };
  if (params !== undefined) body.params = params;
  const headers = { 'Content-Type': 'application/json', 'Accept': 'application/json, text/event-stream' };
  if (proto) headers['MCP-Protocol-Version'] = proto;
  // The endpoint may BE the local middle service, whose every surface requires the token.  This
  // path (rpcAt -> rpc/probe/discovery/ide_* control) used to send NO credential, so with auth on
  // it got 401 token_missing: ide_list reported probeStatus=failed and ide_restart refused with
  // project_path_unknown.  Only attached for the bridge itself, and bridgeAuthHeaders() is {} with
  // no token configured, so a legacy add-in endpoint sees byte-identical traffic.  Nothing here
  // weakens the check - it only lets the hub present the credential it already had.
  // A configured bridge presents ITS OWN token (local slot => the global token, unchanged).
  // Any other endpoint keeps the pre-change rule exactly (loopback + 401 token_missing).
  const perBridgeAuth = endpointAuthFor(target.base, authLabel);
  const credential = !perBridgeAuth && (isBridgeBase(target.base) || authVerdict.get(target.base) === true);
  if (perBridgeAuth) {
    if (perBridgeAuth.token) Object.assign(headers, authHeadersFor(target.base));
  } else if (credential) Object.assign(headers, bridgeAuthHeaders());
  const url = target.base + (rpcPath || target.path);
  const payload = JSON.stringify(body);
  const ms = timeoutMs || FORWARD_TIMEOUT_MS;
  let res = await fetch(url, { method: 'POST', headers: headers, body: payload, signal: AbortSignal.timeout(ms) });
  let raw = await res.text();
  // A LOOPBACK endpoint that answered 401 with OUR OWN reason asked for the credential: repeat the
  // request ONCE with it (and remember the verdict per endpoint, so this costs one extra POST per
  // endpoint, not per call).  See asksForOurToken for the three conditions.
  if (!perBridgeAuth && !credential && asksForOurToken(target.base, res.status, raw)) {
    authVerdict.set(target.base, true);
    dbg('rpcAt ' + target.base + ': 401 token_missing -> repeating ONCE with the credential');
    res = await fetch(url, { method: 'POST', headers: Object.assign({}, headers, bridgeAuthHeaders()), body: payload, signal: AbortSignal.timeout(ms) });
    raw = await res.text();
  }
  // Still an auth failure AFTER presenting the credential (or when none could be presented): THROW,
  // so the caller's existing failure classification runs instead of returning the 401 page as data.
  const authFail = authFailureEnvelope(res.status, raw);
  if (authFail) {
    dbg('rpcAt ' + target.base + ': auth failure ' + authFail.reason + ' (HTTP ' + res.status + ') -> throwing');
    throw authFailureError(authFail, res.status, target.base);
  }
  return decodeRpcBody(raw);
}
async function rpc(port, method, params, proto, timeoutMs, rpcPath) {
  return rpcAt('http://127.0.0.1:' + port, method, params, proto, timeoutMs, rpcPath);
}

// The probe's ONE wire call, transport aware.  A port-bearing instance is reached
// exactly as before.  A PIPE-ONLY instance has no port and the middle service is the ONLY addresser
// for it, so the same JSON-RPC goes over the bridge; with the bridge unreachable this rejects just
// like a dead port does, so the caller's "probe failed" path is unchanged.
function instanceRpc(port, transport, method, params, proto, timeoutMs, rpcPath) {
  if (port) return rpc(port, method, params, proto, timeoutMs, rpcPath);
  if (BRIDGE_MODE === 'inproc') return Promise.reject(new Error('a pipe-only instance is addressable only through the middle service (VB6MCP_BRIDGE=python)'));
  return bridgeRpc(method, params, timeoutMs || PROBE_TIMEOUT_MS);
}
// Runs in the BACKGROUND (never on a response path): a slow add-in must not delay the hub.
// (N11): the probe's VERDICT is now an ADD-IN answer, not a face read.
//   tools/list alone is NOT liveness evidence: the middle service answers it from its OWN snapshot
//   when the add-in cannot be reached, so a wedged IDE still produced probeStatus=ok with
//   liveToolCount=110 while a forwarded tools/call answered noIde / pipe busy 121 (measured).
//   mcp_info is an ADD-IN tool, so an answer to it proves the add-in itself replied over the SAME
//   route a forwarded call uses - which is what makes ide_list and a forwarded call agree.  The
//   call is also the LIVE DLL IDENTITY: mcp_info.build.sha256 is hashed from the loaded DLL at
//   call time, so a receipt can assert which binary answered.
const PROBE_INFO_TIMEOUT_MS = parseInt(process.env.VB6MCP_PROBE_INFO_TIMEOUT_MS || '1500', 10);
// P2: a pipe-only probe is served by the MIDDLE SERVICE (which must resolve a default target and
// then cross the named pipe), so the direct-addin 1.5 s cap is too tight; it made ide_list read
// failed while an explicit-target forward succeeded (N11).  Callers can override.
const PIPE_PROBE_TIMEOUT_MS = parseInt(process.env.VB6MCP_PIPE_PROBE_TIMEOUT_MS || '4500', 10);
async function probe(port, knownProject, pids, transport) {
  pids = pids || [];
  if (!pids.length) return null;
  if (!port && BRIDGE_MODE !== 'inproc') {
    // A pipe-only probe goes THROUGH the middle service, so the service must be safely
    // ours BEFORE any /rpc is attempted.  A foreign listener is never probed as an IDE.
    const ms = await ensureMiddleService();
    if (ms && ms.portTakenByOther) {
      return { livenessFailed: true, probeReason: 'middle_service_port_taken_by_other',
               probeEvidenceAt: Date.now(), probeRoute: BRIDGE_URL + ' -> pipe',
               probeBasis: 'the configured middle-service port is occupied by a foreign listener; refusing to probe it as an IDE (no /rpc was sent)' };
    }
  }
  const route = port ? ('http://127.0.0.1:' + port + ' /rpc')
                     : (BRIDGE_MODE === 'inproc' ? 'none' : BRIDGE_URL + ' -> pipe');
  // P2: send the SAME selector a forward would use, and give a pipe-only probe pipe time.
  const pipeOnlyProbe = !port;
  const probeArgs = (pipeOnlyProbe && pids.length) ? { _target: String(pids[0]) } : {};
  const infoTimeout = pipeOnlyProbe ? PIPE_PROBE_TIMEOUT_MS : PROBE_INFO_TIMEOUT_MS;
  const probeDetail = 'infoTimeoutMs=' + infoTimeout + (pipeOnlyProbe ? ' target=' + probeArgs._target + ' (pipe-only)' : ' (direct add-in)');
  let info = null, infoReason = null;
  try {
    const mi = await instanceRpc(port, transport, 'tools/call', { name: 'mcp_info', arguments: probeArgs }, null, infoTimeout);
    const txt = firstText(mi);
    if (txt) {
      try { info = JSON.parse(txt); } catch (e) { infoReason = 'unparsable_mcp_info'; }
      // FAIL-CLOSED: a structured ERROR is not the add-in answering.  The middle service returns
      // exactly this shape when ITS forward failed (e.g. pipe_unavailable), and counting it as a
      // live answer is how a wedged IDE kept reading probeStatus=ok (N11).  It is also the
      // FAILURE_KIND_RULE's own rule: result.isError / an ok:false body = the tool did NOT run.
      if (info && info.ok === false) { infoReason = 'mcp_info_answered_not_ok(' + String(info.reason || info.message || 'unknown') + ')'; info = null; }
      else if (mi && mi.result && mi.result.isError === true) { infoReason = 'mcp_info_answered_isError'; info = null; }
    } else if (mi && mi.error) infoReason = 'mcp_info_error(' + String((mi.error && mi.error.message) || mi.error) + ')';
    else infoReason = 'empty_mcp_info_answer';
  } catch (e) { infoReason = 'addin_did_not_answer(' + String((e && e.message) || e) + ')'; }
  // NO add-in answer => NOT a live IDE, whatever the face says.  The caller turns this into
  // probeStatus=failed + probeReason, which is the SAME verdict a forward reports.  The SENTINEL
  // (not null) carries the exact reason down to the receipt instead of a generic one.
  if (!info) return { livenessFailed: true, probeReason: infoReason || 'addin_did_not_answer',
                      probeEvidenceAt: Date.now(), probeRoute: route,
                      probeBasis: 'the ADD-IN did not answer mcp_info over ' + route + ' (' + (infoReason || 'unknown') + '); ' + probeDetail };
  let tools = [];
  try {
    const tl = await instanceRpc(port, transport, 'tools/list', {}, null, 8000);
    tools = (tl && tl.result && tl.result.tools) || [];
  } catch (e) { tools = []; }
  let project = knownProject || null;
  if (!project) {
    try {
      const g = await instanceRpc(port, transport, 'tools/call', { name: 'get_project_info', arguments: {} }, null, 1500);
      const c = g && g.result && g.result.content;
      const txt = Array.isArray(c) && c[0] && c[0].text;
      if (txt) { try { const j = JSON.parse(txt); project = j.project || j.name || null; } catch (e2) { } }
    } catch (e) { }
  }
  const build = (info && info.build) || {};
  const ide = (info && info.ide) || null;
  // Best-effort banner from the MCP face; /rpc does not implement initialize.  Fetched in the
  // BACKGROUND so a slow /mcp cannot keep the instance "pending".  Cached by the INSTANCE identity
  // (endpoint/transport + pid/project + the LIVE DLL identity mcp_info just reported), NEVER by the
  // port: a pipe-only instance has port=null and used to share ONE banner with every other pipe IDE.
  const bannerId = bannerIdentity({
    endpoint: LOCAL_ENDPOINT,
    transport: transport || (port ? 'tcp' : 'pipe'),
    pid: pids[0], port: port, project: project,
    addinContract: (info && info.contract) || null,
    dllSha256: (build && build.sha256) || null
  });
  const bannerKey = bannerCacheKey('local', bannerId);
  bannerCachePruneOther(bannerCacheSlot('local', bannerId), bannerKey);
  const instructions = bannerCacheGet(bannerKey);
  if (!instructions) {
    instanceRpc(port, transport, 'initialize', { protocolVersion: LATEST, capabilities: {}, clientInfo: { name: 'vb6mcp-hub', version: SERVER.version } }, null, 3000, '/mcp')
      .then(function (init) {
        const v = init && init.result && init.result.instructions;
        if (v) {
          bannerCachePut(bannerKey, bannerCacheSlot('local', bannerId), bannerId, v);
          dbg('instructions cached for instance pid=' + bannerId.pid + ' dll=' + (bannerId.dllSha256 || '-') + ' contract=' + (bannerId.addinContract || '-'));
        }
      })
      .catch(function () { dbg('instructions fetch failed for instance pid=' + bannerId.pid); });
  }
  return { port: port, pid: pids[0], pids: pids, project: project,
           liveTools: tools.map(function (t) { return t.name; }), liveToolDefs: tools, instructions: instructions,
           // ---- the ONE liveness evidence -------------------------------------
           probeEvidenceAt: Date.now(),
           probeBasis: 'addin_answered_mcp_info via ' + route + '; ' + probeDetail,
           probeRoute: route,
           addinProduct: (info && info.product) || null,
           addinContract: (info && info.contract) || null,
           addinToolCount: (info && info.toolCount) || null,
           // the LIVE DLL identity of the binary that answered (hashed by the add-in at call time)
           dllSha256Live: build.sha256 || null,
           dllSha256LivePath: build.path || null,
           dllSha256LiveReadAt: build.readAt || null,
           sourceHash: build.sourceHash || null,
           ide: ide,
           probeReason: null,
           probeWarnings: infoReason ? ['mcp_info carried no parseable body: ' + infoReason] : [] };
}

// ---- remote endpoints: probe / cache / PER-ENDPOINT single-flight -------------------------
// Keyed per endpoint: local discovery NEVER waits on this map, so a slow or dead remote cannot
// stall a local call, while two concurrent calls to the SAME endpoint share one probe.
const remoteProbe = new Map();   // label -> { at, data, promise }
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
// Best effort, BACKGROUND: the /mcp banner of a remote endpoint (never on a critical path).  Keyed
// by the endpoint + the instance/DLL identity the probe just read, for the same reason the local
// cache is: an endpoint LABEL alone would keep serving the banner of an IDE/DLL that has since
// changed behind the same URL (the remote counterpart of OPEN-27).
function remoteInstructions(ep, url, facts) {
  facts = facts || {};
  const id = bannerIdentity({
    endpoint: ep.label, endpointUrl: url, transport: 'remote',
    pid: facts.pid, port: facts.port, project: facts.project,
    addinContract: facts.addinContract, dllSha256: facts.dllSha256
  });
  const key = bannerCacheKey('remote', id);
  const hit = bannerCacheGet(key);
  if (hit) return hit;
  rpcAt(url, 'initialize', { protocolVersion: LATEST, capabilities: {}, clientInfo: { name: 'vb6mcp-hub', version: SERVER.version } }, null, 3000, '/mcp')
    .then(function (init) {
      const v = init && init.result && init.result.instructions;
      if (v) {
        bannerCachePut(key, bannerCacheSlot('remote', id), id, v);
        dbg('remote instructions cached for ' + ep.label + ' pid=' + id.pid + ' dll=' + (id.dllSha256 || '-'));
      }
    })
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
  // the LIVE DLL identity the remote add-in reported: part of the banner cache key (OPEN-27)
  let liveContract = null, liveDllSha256 = null;
  try {
    const mi = await rpcAt(url, 'tools/call', { name: 'mcp_info', arguments: {} }, null, timeoutMs);
    const txt = firstText(mi);
    if (txt) {
      const j = JSON.parse(txt);
      if (j && j.ide) ide = j.ide;
      if (j && j.fileName) rawPath = String(j.fileName);
      if (j && j.name) project = String(j.name);
      if (j) { liveContract = j.contract || null; liveDllSha256 = (j.build && j.build.sha256) || null; }
    }
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
    projectPathAsGiven: rawPath || null,
    targetKey: ep.label + ':' + (cp.path || ('pid=' + pid)),
    discoveredBy: 'remote_probe', probeStatus: 'ok', pending: false,
    liveTools: tools.map(function (t) { return t.name; }), liveToolDefs: tools,
    instructions: remoteInstructions(ep, url, { pid: pid, port: port, project: project,
                                                addinContract: liveContract, dllSha256: liveDllSha256 })
  };
  if (MULTI_BRIDGE) candidate.bridge = ep.label;
  return { ok: true, endpoint: ep.label, endpointUrl: url, remote: true, candidate: candidate,
           toolCount: tools.length, elapsedMs: Date.now() - t0, timeoutMs: timeoutMs };
}
// ---- bridge discovery: native ide_list first, GET /health as the VISIBLE fallback ----
// A bridge (a VB6MCP_BRIDGES entry) publishes its facts in one of two ways:
//   1. the LOCAL JSON-RPC method ide_list (new middle service) -> discoveryApi=bridge_native
//   2. GET /health (every middle service, an UNCHANGED surface) -> discoveryApi=bridge_health_fallback
// A -32601 answer is PROOF that the method does not exist, so the capability is cached for the
// process and the fallback costs one round trip ever.  A TIMEOUT is NOT proof: it falls back for
// this pass only and the capability stays undecided (nativeMethods=null), so a busy bridge is not
// mislabelled forever.  Both sources are READ-ONLY and neither touches the add-in.
const BRIDGE_NATIVE_PROBE_TIMEOUT_MS = parseInt(process.env.VB6MCP_BRIDGE_NATIVE_TIMEOUT_MS || '2500', 10);
const BRIDGE_NATIVE_CAP = new Map();   // label -> {supported:true|false, at, reason, detail}


function bridgeNativeCapability(label) {
  return BRIDGE_NATIVE_CAP.get(String(label)) || null;
}


async function probeBridgeNative(ep, timeoutMs) {
  const t0 = Date.now();
  const resp = await rpcAt(ep.url, 'ide_list', {}, null, timeoutMs, undefined, ep.label);
  const el = Date.now() - t0;
  if (resp && resp.error) {
    const code = Number(resp.error.code);
    return { ok: false, unsupported: code === -32601,
             reason: code === -32601 ? 'native_method_not_found' : 'native_error',
             code: code, detail: String(resp.error.message || ''), elapsedMs: el };
  }
  const result = resp && resp.result;
  if (!result || typeof result !== 'object' || !Array.isArray(result.instances)) {
    return { ok: false, unsupported: false, reason: 'native_result_unrecognized', elapsedMs: el,
             detail: 'the answer carried no instances[] array (the method exists but the shape is unknown)' };
  }
  return { ok: true, data: result, elapsedMs: el };
}


async function probeBridgeHealth(ep, timeoutMs) {
  const t0 = Date.now();
  const res = await fetch(String(ep.url).replace(/\/+$/, '') + '/health', {
    method: 'GET',
    headers: Object.assign({ Accept: 'application/json' }, authHeadersFor(ep.url, ep.label)),
    signal: AbortSignal.timeout(timeoutMs)
  });
  const raw = await res.text();
  const authFail = authFailureEnvelope(res.status, raw);
  if (authFail) throw authFailureError(authFail, res.status, ep.url);
  const el = Date.now() - t0;
  if (!res.ok) return { ok: false, reason: 'health_http_' + res.status, message: String(raw).slice(0, 200), elapsedMs: el };
  const j = parseRpcBody(raw);
  if (!j || typeof j !== 'object' || Array.isArray(j)) {
    return { ok: false, reason: 'health_unparseable', message: String(raw).slice(0, 200), elapsedMs: el };
  }
  const server = String(j.server || j.product || '');
  if (server !== 'vb6mcp-bridge') {
    // fail-closed: the same identity rule the local middle-service probe uses.  A foreign listener
    // is NEVER read as a bridge (and its instances are never adopted).
    return { ok: false, reason: 'not_a_bridge', server: server || null, elapsedMs: el,
             message: 'the /health answer does not identify itself as vb6mcp-bridge' };
  }
  const insts = Array.isArray(j.instances) ? j.instances
    : (Array.isArray(j.liveInstances) ? j.liveInstances : []);
  return { ok: true, data: j, instances: insts, elapsedMs: el };
}


// ONE bridge instance -> ONE candidate of the SAME set local IDEs belong to.  "bridge" is the
// label the caller addresses it by ([bridge:]selector); bridgeApi names the facts source.
function bridgeCandidates(ep, instances, api) {
  const out = [];
  for (const bi of (instances || [])) {
    if (!bi || typeof bi !== 'object' || bi.pid === undefined || bi.pid === null) continue;
    const pid = Number(bi.pid) || null;
    const port = Number(bi.port) || 0;
    const rawPath = bi.projectPath || '';
    const cp = canonicalProjectPath(rawPath, true);
    const alive = (bi.alive === undefined)
      ? (bi.pidAlive === undefined ? null : !!bi.pidAlive) : !!bi.alive;
    const reach = (bi.addInReachable === undefined) ? null : !!bi.addInReachable;
    out.push({
      endpoint: ep.label, remote: true, bridge: ep.label, endpointUrl: ep.url, note: ep.note,
      bridgeApi: api,
      pid: pid, pids: pid ? [pid] : [], port: port || null,
      project: bi.project || null, projectPath: cp.path || null, pathKey: cp.key, pathVia: cp.via,
      projectPathAsGiven: rawPath || null,
      targetKey: ep.label + ':' + (cp.path || ('pid=' + pid)),
      discoveredBy: api === 'bridge_native' ? 'bridge_native_ide_list' : 'bridge_health',
      probeStatus: (alive === false || reach === false) ? 'failed' : 'ok',
      // The bridge ANSWERED for this instance, so the verdict is DATEABLE and FRESH: without
      // this the hub's livenessOf() would report it as 'stale' (an undated ok is never live).
      probeEvidenceAt: Date.now(), evidenceAt: Date.now(),
      alive: alive, addInReachable: reach,
      probeBasis: api === 'bridge_native'
        ? 'bridges[].ide_list facts (the middle service owns them; addInReachable is route-level)'
        : 'bridges[]./health facts (older middle service: no native methods)',
      transport: bi.transport || (bi.pipeAvailable ? 'pipe' : (port ? 'tcp' : null)),
      pipeAvailable: bi.pipeAvailable === undefined ? null : !!bi.pipeAvailable,
      pipeState: bi.pipeState || null,
      sessionId: bi.sessionId === undefined ? null : bi.sessionId,
      epoch: bi.epoch === undefined ? null : bi.epoch,
      tcpEnabled: bi.tcpEnabled === undefined ? null : bi.tcpEnabled,
      startedAt: bi.startedAt || null,
      startedByHub: !!bi.startedByHub,
      registrationFile: bi.registrationFile || null,
      liveTools: null, liveToolDefs: null, instructions: null
    });
  }
  return out;
}


async function probeBridge(ep, timeoutMs) {
  const t0 = Date.now();
  const cap = bridgeNativeCapability(ep.label);
  const tryNative = !(cap && cap.supported === false);
  let native = null, nativeError = null;
  if (tryNative) {
    try { native = await probeBridgeNative(ep, timeoutMs); }
    catch (e) { nativeError = e; }
    if (native) {
      if (native.unsupported) {
        BRIDGE_NATIVE_CAP.set(ep.label, { supported: false, at: Date.now(), reason: native.reason, detail: native.detail || null });
      } else if (native.ok) {
        BRIDGE_NATIVE_CAP.set(ep.label, { supported: true, at: Date.now() });
      }
    }
  }
  const token = tokenFactForEndpoint(ep);
  if (native && native.ok) {
    const cands = bridgeCandidates(ep, native.data.instances, 'bridge_native');
    return { ok: true, endpoint: ep.label, bridge: ep.label, endpointUrl: ep.url, remote: true,
             discoveryApi: 'bridge_native', nativeMethods: true, nativeMethod: 'ide_list',
             candidates: cands, bridgeFacts: native.data, elapsedMs: Date.now() - t0,
             token: token, checkedAt: new Date().toISOString() };
  }
  const nativeUnsupported = !!(cap && cap.supported === false) || !!(native && native.unsupported);
  let h = null, healthError = null;
  try { h = await probeBridgeHealth(ep, timeoutMs); }
  catch (e) { healthError = e; }
  if (h && h.ok) {
    const cands = bridgeCandidates(ep, h.instances, 'bridge_health_fallback');
    return { ok: true, endpoint: ep.label, bridge: ep.label, endpointUrl: ep.url, remote: true,
             discoveryApi: 'bridge_health_fallback',
             nativeMethods: nativeUnsupported ? false : null,
             nativeMethod: 'ide_list',
             nativeReason: nativeUnsupported ? 'native_method_not_found(-32601)'
               : (nativeError ? 'native_error' : (native ? native.reason : 'not_probed')),
             candidates: cands, health: h.data, elapsedMs: Date.now() - t0,
             token: token, checkedAt: new Date().toISOString() };
  }
  // A failing bridge is KEPT and its failure is NAMED - never a silent drop.  Auth failures are
  // their own reason (not "unreachable"): the peer answered and refused the credential.
  const err = healthError || nativeError;
  const authErr = (err && err.code === 'bridge_auth_failed') ? err : null;
  if (authErr) {
    return { ok: false, endpoint: ep.label, bridge: ep.label, endpointUrl: ep.url, remote: true,
             reason: 'endpoint_auth_failed', failureKind: 'toolError', retrySafe: false,
             suggestedAction: 'fix_params',
             authReason: authErr.reason, authStatus: authErr.status, authHint: authErr.authHint,
             message: String(authErr.message || authErr), discoveryApi: null,
             nativeMethods: nativeUnsupported ? false : null, nativeMethod: 'ide_list',
             elapsedMs: Date.now() - t0, token: token, checkedAt: new Date().toISOString() };
  }
  const msg = String((err && err.message) || (h && h.message) || 'the bridge did not answer');
  const isTimeout = !!(err && (err.name === 'TimeoutError' || err.name === 'AbortError'
    || /abort|timed? ?out/i.test(String(err && err.message))));
  return { ok: false, endpoint: ep.label, bridge: ep.label, endpointUrl: ep.url, remote: true,
           reason: isTimeout ? 'endpoint_timeout' : ((h && h.reason) || 'endpoint_unreachable'),
           failureKind: isTimeout ? 'wedged' : 'noIde', retrySafe: true, suggestedAction: 'retry_once',
           message: msg, discoveryApi: null,
           nativeReason: nativeUnsupported ? 'native_method_not_found(-32601)'
             : (nativeError ? 'native_error' : (native ? native.reason : 'not_probed')),
           nativeMethods: nativeUnsupported ? false : null, nativeMethod: 'ide_list',
           elapsedMs: Date.now() - t0, token: token, checkedAt: new Date().toISOString() };
}


async function probeBridgeCached(ep, timeoutMs) {
  const cached = remoteCacheGet(ep.label);
  if (cached) return cached;
  const st = remoteProbe.get(ep.label);
  if (st && st.promise) return st.promise;
  const promise = probeBridge(ep, timeoutMs || BRIDGE_NATIVE_PROBE_TIMEOUT_MS).then(function (d) {
    remoteProbe.set(ep.label, { at: Date.now(), data: d, promise: null });
    return d;
  }).catch(function (e) {
    const d = { ok: false, endpoint: ep.label, bridge: ep.label, endpointUrl: ep.url, remote: true,
                reason: 'bridge_probe_failed', failureKind: 'noIde', retrySafe: true,
                suggestedAction: 'retry_once', message: String((e && e.message) || e),
                discoveryApi: null, nativeMethods: null, nativeMethod: 'ide_list',
                elapsedMs: 0, token: tokenFactForEndpoint(ep), checkedAt: new Date().toISOString() };
    remoteProbe.set(ep.label, { at: Date.now(), data: d, promise: null });
    return d;
  });
  remoteProbe.set(ep.label, { at: st ? st.at : 0, data: st ? st.data : null, promise: promise });
  return promise;
}


// One row of ide_list.bridges[]: the bridge's probe state, its credential FACTS (never the value)
// and whether the native facts method exists.  A failed row keeps reason/suggestedAction/retrySafe.
function bridgeStatusRow(ep) {
  const st = remoteProbe.get(ep.label) || null;
  const fresh = remoteCacheGet(ep.label);
  const d = fresh || (st ? st.data : null);
  const cap = bridgeNativeCapability(ep.label);
  const pending = !!(st && st.promise);
  return {
    name: ep.label, url: ep.url, bridge: true, local: false, source: ep.source, note: ep.note || '',
    token: tokenFactForEndpoint(ep),
    nativeMethods: d && d.nativeMethods !== undefined ? d.nativeMethods : (cap ? !!cap.supported : null),
    nativeMethod: 'ide_list',
    nativeCapabilityBasis: cap
      ? (cap.supported ? 'the bridge answered the local ide_list method' : 'the bridge answered -32601 (method not found); /health fallback is used')
      : (d && d.nativeReason ? d.nativeReason : 'not probed yet'),
    discoveryApi: d ? (d.discoveryApi || null) : null,
    probeStatus: d ? (d.ok ? 'ok' : 'failed') : (pending ? 'pending' : 'not_probed'),
    reason: d ? (d.ok ? null : (d.reason || 'failed')) : null,
    failureKind: d && !d.ok ? (d.failureKind || null) : null,
    authReason: d && !d.ok ? (d.authReason || null) : null,
    suggestedAction: d && !d.ok ? (d.suggestedAction || null) : null,
    retrySafe: d && !d.ok ? (d.retrySafe === true) : null,
    instances: d && d.ok ? ((d.candidates || []).length) : 0,
    elapsedMs: d && d.elapsedMs !== undefined ? d.elapsedMs : null,
    checkedAt: d && d.checkedAt ? d.checkedAt : null,
    cached: !!fresh, pending: pending
  };
}


// The row of ONE configured bridge, for a failure receipt that must agree with ide_list.bridges[].
function bridgeRowForLabel(label) {
  const ep = endpointByLabel(label);
  if (!ep || !ep.bridge) return null;
  try { return bridgeStatusRow(ep); } catch (e) { return null; }
}

// The local slot (VB6MCP_BRIDGE_URL) as a bridges[] row: no native probe is needed - the hub reads
// this machine's registry directly - so nativeMethods is null with an explicit basis.
function localBridgeStatusRow() {
  const h = bridgeHealthCache.data || null;
  return {
    name: LOCAL_ENDPOINT, url: BRIDGE_URL, bridge: false, local: true, source: 'VB6MCP_BRIDGE_URL',
    note: 'this machine (instance registry + port scan)',
    token: { configured: !!BRIDGE_TOKEN, source: BRIDGE_TOKEN_SOURCE, length: BRIDGE_TOKEN.length,
             fileConfigured: !!process.env.VB6MCP_TOKEN_FILE, error: null },
    nativeMethods: null, nativeMethod: 'ide_list',
    nativeCapabilityBasis: 'not_applicable: the hub reads this machine registry directly; the local slot is not probed natively',
    discoveryApi: 'local_registry+health',
    probeStatus: middleServiceState.reachable ? 'ok' : 'failed',
    reason: middleServiceState.reachable ? null : (middleServiceState.reason || 'not_checked'),
    failureKind: null, authReason: null,
    suggestedAction: middleServiceState.reachable ? null : 'ide_start',
    retrySafe: middleServiceState.reachable ? null : true,
    instances: h && Array.isArray(h.instances) ? h.instances.length : 0,
    elapsedMs: null, checkedAt: null, cached: false, pending: false
  };
}


async function probeEndpointCached(label, timeoutMs) {
  const ep = endpointByLabel(label);
  if (!ep || !ep.remote) return { ok: false, endpoint: String(label), reason: 'endpoint_unknown' };
  if (ep.bridge) return probeBridgeCached(ep, timeoutMs);   // multi-bridge: native facts, /health fallback
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
    if (fresh && fresh.ok) {
      if (Array.isArray(fresh.candidates)) { for (const c of fresh.candidates) candidates.push(c); }
      else if (fresh.candidate) candidates.push(fresh.candidate);
    }
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
// ---- /N11: BRIDGE-ASSISTED discovery
  // The hub and the middle service do NOT necessarily share a registry view: they are different
  // USER CONTEXTS, so %TEMP%\vb6mcp\instances can be the interactive user's (where the add-in
  // registered) while the hub sees an empty directory.  In that case the hub's own registry is
  // blind, the PORTSCAN then adopts the middle service itself as a fake IDE, and
  // ide_list contradicts the very calls it is supposed to explain (N11: measured on t16 - ide_list
  // listed ONLY the bridge process pid 28240/port 8797 while ide_state answered from the REAL IDE
  // pid 21660 through the pipe).
  // The middle service IS the only addresser for a pipe-only IDE, so its /health is the
  // authoritative WITNESS of that IDE's existence.  Only used when the hub's own registry yielded
  // NOTHING credible (the registry stays the priority path), and every adopted candidate is
  // labelled discoveredBy=bridge_health with its registration file as the basis.
  if (!pendingCands.length && BRIDGE_MODE !== 'inproc') {
    let bctx = bridgeIdentitySync();
    if (!bctx) { try { bctx = await bridgeIdentity(); } catch (e) { bctx = null; } }
    for (const bi of ((bctx && bctx.instances) || [])) {
      if (!bi || !bi.pid || bi.pidAlive === false) continue;
      const isPipe = bi.pipeState === 'candidate' || String(bi.transport || '') === 'pipe';
      const bport = Number(bi.port) || 0;
      if (!isPipe && !bport) continue;
      pendingCands.push({
        pid: Number(bi.pid), pids: [Number(bi.pid)], port: bport, transport: isPipe ? 'pipe' : 'tcp',
        pipe: bi.pipe || null, project: bi.project || null, projectPath: bi.projectPath || null,
        startedAt: null, version: null, dllSha256: null, startedByHub: false,
        discoveredBy: 'bridge_health', registryFile: bi.registryFile || null,
        bridgePipeState: bi.pipeState || null, bridgePipeAvailable: bi.pipeAvailable === true,
        bridgeRegistryRoot: bi.registryRoot || null
      });
      seenPorts.add(bport);
    }
    if (pendingCands.length) dbg('discoverBounded: adopted ' + pendingCands.length + ' add-in registration(s) from the middle service /health (the hub\'s own registry view was EMPTY)');
  }
  // Existence first (registry + pid alive + port listening); the probe only ENRICHES.
  pendingCands.forEach(function (c) { insts.push(Object.assign({ pending: true, probeStatus: 'pending' }, c)); });
  for (const c of pendingCands) {
    const it = await probe(c.port, c.project, c.pids, c.transport).catch(function () { return null; });
    // Keyed by pid AS WELL as port: every pipe-only instance has port=null, so a port-only key
    // would hand one pipe IDE's probe result to another (TCP instances keep a unique port here).
    const idx0 = insts.findIndex(function (x) { return x.port === c.port && x.pid === c.pid; });
    if (!it || it.livenessFailed) {
      stale.push({ file: c.registryFile, pid: c.pid, port: c.port, reason: 'registry_port_not_mcp_probe_failed' });
      // Terminal, explicit state: the probe ran and failed -> not pending any more, the tool
      // count stays UNKNOWN (null), never a misleading 0.
      if (idx0 >= 0) insts[idx0] = Object.assign({}, insts[idx0], { pending: false, probeStatus: 'failed',
        probeReason: (it && it.probeReason) || 'probe_failed_or_timeout',
        probeEvidenceAt: (it && it.probeEvidenceAt) || Date.now(),
        probeRoute: (it && it.probeRoute) || null,
        probeBasis: (it && it.probeBasis) || 'addin_did_not_answer_mcp_info (the registry/pid/transport said the instance exists; the ADD-IN did not answer, so this is NOT a live IDE - the same verdict a forwarded call reports)' });
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
      try { it = await withTimeout(probe(port, null, pids, null), 8000); } catch (e) { it = null; }
      seenPorts.add(port);
      if (it && it.livenessFailed) {
        insts.push({ port: port, pid: pids[0], pids: pids, project: null, discoveredBy: 'portscan', registryFile: null,
                     pending: false, probeStatus: 'failed', probeReason: it.probeReason || 'probe_failed_or_timeout',
                     probeEvidenceAt: it.probeEvidenceAt || Date.now(), probeRoute: it.probeRoute || null, probeBasis: it.probeBasis || null });
      } else if (it) {
        insts.push(Object.assign({}, it, { discoveredBy: 'portscan', registryFile: null, pending: false, probeStatus: 'ok' }));
      } else {
        // The port IS listening, so the instance exists even if the probe fails -
        // report it (unknown face) instead of silently dropping it.
        insts.push({ port: port, pid: pids[0], pids: pids, project: null, discoveredBy: 'portscan', registryFile: null,
                     pending: false, probeStatus: 'failed', probeReason: 'probe_failed_or_timeout',
                     probeEvidenceAt: Date.now(), probeBasis: 'listening port without an add-in answer (NOT a live IDE)' });
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
  // TEST HOOK (never set in production): a discovery pass that NEVER settles, so the
  // regression suite can prove the ladder fallback still answers within its TOTAL budget.
  if (process.env.VB6MCP_TEST_DISCOVERY_NEVER_RESOLVE === '1') return new Promise(function () { });
  if (discovery.data && Date.now() - discovery.at < DISCOVERY_TTL_MS) return discovery.data;
  if (discovery.promise) return discovery.promise;
  const pass = discoverBounded().then(function (d) {
    // A pass that was INVALIDATED while it ran must never PUBLISH over the newer face that
    // replaced it: it was started against OLDER facts (measured: an invalidated slow pass
    // overwrote the fresh face with an empty one, re-creating the stale window the ladder
    // fallback above removes).
    if (discovery.promise !== pass) {
      dbg('discover: a SUPERSEDED pass completed; its result is NOT published (' + (((d && d.instances) || []).length) + ' instance(s))');
      return d;
    }
    discovery = { at: Date.now(), data: d, promise: null };
    // Never persist a fallback/empty result as the authoritative face.
    if (d && d.instances && d.instances.length) { try { saveFaceCache(d); } catch (e) { } }
    return d;
  }).catch(function (e) {
    if (discovery.promise === pass) discovery.promise = null;
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
  discovery.promise = pass;
  return pass;
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
           endpoint: i.endpoint || LOCAL_ENDPOINT, remote: !!i.remote, targetKey: i.targetKey || null,
           ...(MULTI_BRIDGE ? { bridge: i.bridge || i.endpoint || LOCAL_ENDPOINT } : {}) };
}
// ---- caller-spelling (ALIAS) echo for ide_list ((3)) -----------------------------
// VB6 refuses to open the same .vbp twice, so ONE project can be spelled several ways: through
// a directory junction that is an alias of the same real directory, through the junction's
// target, with "/" instead of "\", with different case, or with an 8.3
// name.  The hub normalizes ALL of them to ONE identity (realpathSync.native) - and then the
// RECEIPT hides which spelling the caller actually used.  ide_list now echoes it back next to
// the canonical path, so "the same project under three aliases" is visible instead of having to
// be re-derived.  A spelling is echoed when it resolves to a discovered candidate:
//   * project_path   - the canonical path key is equal (any separator/case/8.3/junction form);
//   * alias_directory- the spelling is an existing DIRECTORY and the candidate's realpath
//                      project path lives under it (that is exactly what a junction dir is);
//   * project_name   - the legacy bare-NAME alias.
// A bare pid/port is an address, not a project alias, and is deliberately NOT echoed.
const ALIAS_RULE = 'aliasAsGiven/projectPathAsGiven are ADDITIVE echoes, not selectors: projectPathReal (projectPath) is the canonical identity (realpathSync.native, (3)); projectPathAsGiven is the spelling the registry/probe PUBLISHED; aliasAsGiven is the caller-supplied spelling (_target / ide_target default) that resolves to the SAME candidate - by canonical path key, or as an existing directory that realpath-contains the project (a junction alias). Junction/alias spellings of one .vbp are ONE candidate, never a conflict.';
function callerSpellings(rawTarget) {
  const out = [];
  const push = function (v, source) {
    if (v === undefined || v === null) return;
    if (Array.isArray(v)) { for (const x of v) push(x, source); return; }
    const s = String(v).trim();
    if (!s || s.toLowerCase() === 'auto') return;
    if (out.some(function (x) { return x.asGiven === s && x.source === source; })) return;
    out.push({ asGiven: s, source: source });
  };
  push(rawTarget, '_target');
  push(state.target, 'ide_target');
  return out;
}
// Does ONE caller spelling resolve to ONE candidate?  Returns null (not this candidate) or the
// matchedBy evidence.  The endpoint scope is honoured: a bare spelling only ever addresses THIS
// machine, and an endpoint-qualified spelling only its own endpoint.
function aliasHitFor(sp, inst) {
  const t = sp.target;
  if (!t || t.kind === 'endpoint_unknown') return null;
  if (t.kind === 'endpoint') {
    const ep = String(inst.endpoint || LOCAL_ENDPOINT).toLowerCase();
    if (ep !== String(t.endpoint.label).toLowerCase()) return null;
  } else if (inst.remote) return null;
  const sel = String(t.selector || '');
  if (!sel) return null;
  const cp = canonicalProjectPath(sel, t.kind === 'endpoint' ? !!t.endpoint.remote : false);
  if (cp.key && inst.pathKey && cp.key === inst.pathKey) return { matchedBy: 'project_path', via: cp.via, normalized: cp.path };
  if (cp.path && inst.projectPath) {
    const dir = normalizePathString(cp.path).toLowerCase().replace(/\\+$/, '') + '\\';
    const proj = normalizePathString(inst.projectPath).toLowerCase();
    if (proj.indexOf(dir) === 0) return { matchedBy: 'alias_directory', via: cp.via, normalized: cp.path };
  }
  if (inst.project && String(inst.project) === sel) return { matchedBy: 'project_name', via: 'string', normalized: sel };
  return null;
}
// The receipt-level echo: EVERY spelling the caller supplied, what it resolved to, and whether
// all of them land on the SAME candidate (that is the property the caller needs to see).
function buildAliasEcho(spellings, insts) {
  const supplied = [];
  const candidates = new Map();
  for (const sp of spellings || []) {
    const t = parseTarget(sp.asGiven);
    const row = { asGiven: sp.asGiven, source: sp.source, targetKind: t.kind };
    if (t.kind === 'auto') { row.matched = 0; row.reason = 'auto'; supplied.push(row); continue; }
    if (t.kind === 'endpoint_unknown') { row.matched = 0; row.reason = 'endpoint_unknown'; row.endpoint = t.label; supplied.push(row); continue; }
    if (t.kind === 'endpoint') { row.endpoint = t.endpoint.label; row.remote = !!t.endpoint.remote; }
    const hits = [];
    for (const i of (insts || [])) {
      const h = aliasHitFor(Object.assign({}, sp, { target: t }), i);
      if (h) hits.push(Object.assign({ pid: i.pid, port: i.port, project: i.project, projectPathReal: i.projectPath || null,
                                       endpoint: i.endpoint || LOCAL_ENDPOINT, remote: !!i.remote, targetKey: i.targetKey || null,
                                       projectPathAsGiven: i.projectPathAsGiven || null }, h));
    }
    row.matched = hits.length;
    row.matches = hits;
    row.sameCandidate = hits.length === 1;
    if (!hits.length) row.reason = 'no_match';
    if (hits.length === 1) {
      row.normalized = hits[0].normalized; row.via = hits[0].via; row.matchedBy = hits[0].matchedBy;
      row.projectPathReal = hits[0].projectPathReal; row.resolvedPid = hits[0].pid; row.resolvedPort = hits[0].port;
      candidates.set(hits[0].endpoint + '\u0000' + String(hits[0].pid) + '\u0000' + String(hits[0].port), { pid: hits[0].pid, port: hits[0].port, endpoint: hits[0].endpoint, projectPathReal: hits[0].projectPathReal });
    }
    supplied.push(row);
  }
  const sameCandidateCount = supplied.filter(function (r) { return r.sameCandidate; }).length;
  return {
    rule: ALIAS_RULE,
    supplied: supplied,
    sameCandidateCount: sameCandidateCount,
    sameCandidate: sameCandidateCount === supplied.length && supplied.length > 0,
    resolvedCandidates: Array.from(candidates.values())
  };
}
// Selection inside ONE endpoint scope.  target = a resolved selector (or null); explicit =
// the selector came from _target (a miss is an error) vs from the ide_target DEFAULT (a miss
// is ignored, never vetoing the live set).
// The receipt-level record of HOW a target was chosen (N14).  It is attached to every
// pick - success AND refusal - so the caller can predict the outcome instead of discovering it.
function suggestionFor(insts) {
  if (!insts || !insts.length) return null;
  const c = insts[0];
  const base = { _target: c.pid ? String(c.pid) : (c.projectPath || c.project || ''),
           projectPath: c.projectPath || null, targetKey: c.targetKey || null,
           ide_target: c.pid ? String(c.pid) : null,
           note: 'a bare pid is accepted by the hub AND by the middle service; the canonical key is [endpoint:]projectPath' };
  if (MULTI_BRIDGE) {
    // A bare selector is AMBIGUOUS by design once bridges are configured (one pid can exist on
    // several machines), so the suggestion must be a QUALIFIED key the caller can paste back.
    const keys = insts.map(function (i) { return i.targetKey || ((i.bridge || i.endpoint || LOCAL_ENDPOINT) + ':' + (i.projectPath || i.pid || '')); });
    base._target = c.targetKey || base._target;
    base.candidates = keys;
    base.note = 'multiple IDEs are visible; address ONE with a qualified key such as "' + (keys[0] || '<bridge>:<pid>') + '" (<bridge>:<pid|project|projectPath>)';
  }
  return base;
}
function resolutionOf(step, insts, opts, extra) {
  const owned = (insts || []).filter(function (i) { return ownershipOf(i).owned; });
  return Object.assign({
    rule: TARGET_RESOLUTION_RULE,
    targetResolutionRule: TARGET_RESOLUTION_RULE,
    step: step,
    liveCount: (insts || []).length,
    requested: (opts && opts.requested !== undefined) ? opts.requested : null,
    explicit: !!(opts && opts.explicit),
    ideTargetDefault: defaultTargetLabel(),
    candidates: (insts || []).map(candidateBrief),
    ladder: (insts || []).length > 1 ? [
      { step: 'only_live_ide', verdict: 'skip', liveCount: insts.length },
      { step: 'sole_hub_owned', verdict: owned.length === 1 ? 'hit' : 'skip', matching: owned.map(function (i) { return i.pid; }) },
      { step: 'sole_responding', verdict: 'skip', matching: insts.filter(function (i) { return (i.liveness ? i.liveness.state === 'responding' : i.probeStatus === 'ok'); }).map(function (i) { return i.pid; }) }
    ] : undefined,
    suggestion: (insts || []).length > 1 ? suggestionFor(insts) : null
  }, extra || {});
}
function respondingOf(insts) {
  return (insts || []).filter(function (i) { return (i.liveness ? i.liveness.state === 'responding' : i.probeStatus === 'ok'); });
}
function pickTarget(insts, target, opts) {
  opts = opts || {};
  const isRemoteScope = !!opts.remoteScope;
  const explicit = !!opts.explicit;
  if (!insts.length) return { ok: false, resolution: resolutionOf('no_candidate', insts, opts),
                              message: isRemoteScope ? 'the selected endpoint reports no IDE' : 'no IDE instance listening (scanned ' + DEFAULT_PORTS.length + ' default ports + registry)' };
  const desc = function () { return candidateDesc(insts, isRemoteScope); };
  if (target !== null && target !== undefined && target !== '') {
    const s = String(target);
    const hit = matchCandidates(insts, s, isRemoteScope);
    if (hit.length === 1) return { ok: true, ...hit[0], selectedBy: (explicit ? 'explicit_target' : 'effective_default'),
                                   matchedBy: matchedByFor(hit[0], s, isRemoteScope), matchedFrom: s,
                                   resolution: resolutionOf(explicit ? 'explicit_target' : 'ide_target_default', insts, opts, { selectedBy: explicit ? 'explicit_target' : 'ide_target_default' }) };
    if (hit.length > 1) {
      const keySet = {};
      hit.forEach(function (i) { keySet[candidateIdentity(i)] = 1; });
      const ids = Object.keys(keySet);
      const sameIdentity = ids.length === 1;
      const ep = hit[0].endpoint || LOCAL_ENDPOINT;
      return { ok: false, conflict: true, identityConflict: sameIdentity, distinctIdentities: ids.length, matchedCount: hit.length,
        hits: hit.map(candidateBrief),
        candidates: hit.map(candidateBrief), suggestion: suggestionFor(hit),
        resolution: resolutionOf('refuse_target_conflict', insts, opts, { matchedCount: hit.length, distinctIdentities: ids.length }),
        message: sameIdentity
          ? 'target "' + s + '" is ONE project identity ("' + (hit[0].projectPath || hit[0].project || '?') + '") served by ' + hit.length + ' live IDEs on endpoint "' + ep + '": ' + hit.map(function (i) { return i.pid + '@' + i.port; }).join(',') + ' -- VB6 cannot open the same .vbp twice; refusing to pick one'
          : 'target "' + s + '" is AMBIGUOUS: it matches ' + hit.length + ' live IDEs with ' + ids.length + ' different project identities on endpoint "' + ep + '": ' + hit.map(function (i) { return i.pid + '@' + i.port + ' ' + (i.projectPath || i.project || '?'); }).join(' | ') + ' -- pick one by its projectPath' };
    }
    // A STALE DEFAULT (ide_target pointing at a project that is gone) must
    // never veto the live instances - that made every implicit call fail with no_ide_running
    // while the only live IDE was right there.
    if (explicit) {
      const livePids = onlinePidsOf(insts);
      const deadPid = /^\d+$/.test(s) ? Number(s) : null;
      return { ok: false, mismatch: true, candidates: insts.map(candidateBrief), suggestion: suggestionFor(insts),
        onlinePids: livePids, deadTargetPid: deadPid,
        targetDiagnostic: targetDiagnosticBlock({ reason: 'explicit_target_not_live', onlinePids: livePids, requested: s, explicit: true, deadTargetPid: deadPid }),
        resolution: resolutionOf('refuse_target_mismatch', insts, opts),
        message: 'target "' + s + '" not found among ' + desc() + ' -- pick one with _target (pid|project|port) or ide_target; onlinePids[] lists what IS live' };
    }
    if (insts.length === 1) return { ok: true, ...insts[0], staleTarget: s, selectedBy: 'unique_live',
                                     resolution: resolutionOf('stale_default_ignored_only_live_ide', insts, opts, { selectedBy: 'unique_live' }) };
  }
  // ---- the DOCUMENTED DEFAULT LADDER (N14) ------------------------------------------
  // Deterministic, predictable, and reported step by step: only live IDE -> sole hub-owned IDE
  // -> sole RESPONDING IDE -> refuse with every candidate + a suggestion.  It never guesses, and
  // the caller can predict which rung applies from the receipt fields alone.
  if (insts.length === 1) return { ok: true, ...insts[0], selectedBy: 'unique_live',
                                   resolution: resolutionOf('only_live_ide', insts, opts, { selectedBy: 'unique_live', ladderVerdict: 'hit' }) };
  const owned = insts.filter(function (i) { return ownershipOf(i).owned; });
  if (owned.length === 1) return { ok: true, ...owned[0], selectedBy: 'sole_hub_owned',
                                   resolution: resolutionOf('sole_hub_owned', insts, opts, { selectedBy: 'sole_hub_owned', ladderVerdict: 'hit' }) };
  const resp = respondingOf(insts);
  if (resp.length === 1) return { ok: true, ...resp[0], selectedBy: 'sole_responding',
                                  resolution: resolutionOf('sole_responding', insts, opts, { selectedBy: 'sole_responding', ladderVerdict: 'hit' }) };
  return { ok: false, ambiguous: true, candidates: insts.map(candidateBrief), suggestion: suggestionFor(insts),
           resolution: resolutionOf('refuse_ambiguous_target', insts, opts, { ladderVerdict: 'refused', ownedCount: owned.length, respondingCount: resp.length }),
           message: 'several instances online and none of the default ladder rungs is decisive (' + desc() + ') -- set _target (pid|project|port) or ide_target; the ladder and its per-rung verdict are in resolution' };
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
// ---- R-1021-03 C: the UPSTREAM failure envelope is PASS-THROUGH ONLY ---------------------
// The middle service (and, inside it, the client half of ctl/1) OWNS these facts.  The hub's
// job is to carry them to the caller unchanged: it may fill a key that is ABSENT (a peer that
// predates the envelope) but it must NEVER overwrite one that is present - otherwise the hub's
// own guess would hide a real verdict such as peer_unresponsive + oob{acked:false}.
const UPSTREAM_ENVELOPE_KEYS = Object.freeze([
  'failureKind', 'failureKindRule', 'peerState', 'suggestedAction', 'retryAfterMs', 'retrySafe',
  'maxMs', 'maxMsSource', 'budgetMs', 'heldBy', 'ctlAvailable', 'oob', 'limits',
  'delivered', 'applied', 'reason', 'reasonLegacy'
]);
const OOB_KEYS = Object.freeze(['attempted', 'acked', 'why', 'ackOp', 'helloMs', 'rttMs',
                                'sid', 'authBasis', 'delivered', 'applied', 'peerState',
                                'suggestedAction']);
// A merged note fills only ABSENT keys.  Named as a function so the rule is one place, not
// four inline undefined checks that could drift apart.
function mergeNote(obj, note) {
  if (!note || typeof note !== 'object') return obj;
  for (const k of Object.keys(note)) {
    if (obj[k] === undefined) obj[k] = note[k];
  }
  return obj;
}
// An OLDER / error-shaped peer can return a transport failure as a JSON-RPC ERROR whose `data`
// carries the flat envelope.  Lift ONLY the envelope keys into a body, so maxMs/oob/peerState
// are not lost inside error.data (the JSON-RPC code/message stay visible as jsonrpcError).
function errorEnvelopeOf(resp) {
  if (!resp || typeof resp !== 'object' || !resp.error || typeof resp.error !== 'object') return null;
  const d = resp.error.data;
  if (!d || typeof d !== 'object' || Array.isArray(d)) return null;
  const body = { ok: false };
  let n = 0;
  for (const k of UPSTREAM_ENVELOPE_KEYS) {
    if (d[k] !== undefined) { body[k] = d[k]; n++; }
  }
  if (!n) return null;
  if (body.message === undefined) body.message = String(resp.error.message || '');
  body.jsonrpcError = { code: resp.error.code, message: resp.error.message };
  return { content: [{ type: 'text', text: JSON.stringify(body, null, 2) }], isError: true };
}
function envelopeOf(obj, ide, contract, note) {
  const o = Object.assign({}, obj);
  if (!o.mcp) o.mcp = mcp(SERVER.version, contract);
  // (4): never overwrite an `ide` the add-in already reported.
  if (ide && !o.ide) o.ide = ide;
  // MULTI-BRIDGE ONLY: the add-in's own ide block predates the bridge layer, so the identity
  // markers are added when absent (never overwritten).  Single-bridge receipts stay unchanged.
  if (ide && o.ide && MULTI_BRIDGE) {
    if (ide.bridge && o.ide.bridge === undefined) o.ide.bridge = ide.bridge;
    if (ide.endpoint && o.ide.endpoint === undefined) o.ide.endpoint = ide.endpoint;
    if (ide.remote !== undefined && o.ide.remote === undefined) o.ide.remote = ide.remote;
  }
  if (note) {
    if (typeof note === 'string') { if (!o.targetNote) o.targetNote = note; }
    else { mergeNote(o, note); }
  }
  return o;
}
// Kept as the ONE stringifier: every body text is a single document in the same shape.
function bodyOf(obj, ide, contract, note) {
  return JSON.stringify(envelopeOf(obj, ide, contract, note), null, 2);
}

function textResult(obj, contract) {
  return { content: [{ type: 'text', text: bodyOf(obj, null, contract || snapshot().contractVersion) }], isError: false };
}

// tools/call result keys are EXACTLY {content,isError}; ide + mcp go into the body text.
// ---- WP-22-02 / MCP-1021-45: content[0].text is ONE JSON DOCUMENT ------------------------
// An ARRAY-shaped body (find_gui is the measured one) used to be emitted as
//   <array text> + "\n" + <envelope object>
// so the whole text did NOT parse, and an EMPTY result ([]) was indistinguishable from a
// parse failure at the text layer. A non-object body is now ONE document:
//   { "singleDocument": true, "data": <array|scalar>, "receipt": { mcp, ide, targetNote... } }
// A plain OBJECT body keeps its existing FLAT shape (no existing field moved), and
// structuredContent is deliberately NOT added (tools/call result keys stay {content,isError}).
// The old two-part form is kept behind VB6MCP_LEGACY_TEXTBODY=1 for ONE version.
const LEGACY_TEXTBODY = process.env.VB6MCP_LEGACY_TEXTBODY === '1';
const SINGLE_DOC_RULE = 'content[0].text is ONE JSON document. An OBJECT body keeps its flat receipt shape; an ARRAY/scalar body is {singleDocument:true, data:<value>, receipt:{mcp,ide,...}} and an empty result is data:[] (so "no hits" is distinguishable from "unparsable"); a body that is not JSON at all becomes {ok:false, reason:"content_not_json", nonJsonText:true, rawText:"<the raw text>", receipt:{...}}. The result keys stay exactly {content,isError}. VB6MCP_LEGACY_TEXTBODY=1 restores the old two-part form for one version.';
function isPlainObject(v) { return !!v && typeof v === 'object' && !Array.isArray(v); }
function isSingleDocShape(v) { return isPlainObject(v) && v.singleDocument === true && Object.prototype.hasOwnProperty.call(v, 'data') && isPlainObject(v.receipt); }
function stampBody(result, ide, contract, note) {
  const content = Array.isArray(result && result.content) ? result.content.map(function (c) { return Object.assign({}, c); }) : [];
  let textItem = null;
  for (const c of content) { if (c && c.type === 'text' && typeof c.text === 'string') { textItem = c; break; } }
  if (textItem) {
    let parsed, parseOk = false;
    try { parsed = JSON.parse(textItem.text); parseOk = true; } catch (e) { parseOk = false; }
    if (parseOk && isPlainObject(parsed) && !isSingleDocShape(parsed)) {
      textItem.text = bodyOf(parsed, ide, contract, note);
    } else if (parseOk && isSingleDocShape(parsed)) {
      // wrapped upstream (the bridge): stamp the RECEIPT, never the data
      textItem.text = JSON.stringify({ singleDocument: true, data: parsed.data, receipt: envelopeOf(parsed.receipt, ide, contract, note) }, null, 2);
    } else if (parseOk) {
      if (LEGACY_TEXTBODY) { textItem.text = textItem.text + '\n' + bodyOf({}, ide, contract, note); }
      else { textItem.text = JSON.stringify({ singleDocument: true, data: parsed, receipt: envelopeOf({}, ide, contract, note) }, null, 2); }
    } else if (LEGACY_TEXTBODY) {
      textItem.text = textItem.text + '\n' + bodyOf({}, ide, contract, note);
    } else {
      textItem.text = JSON.stringify({ ok: false, reason: 'content_not_json', nonJsonText: true, rawText: textItem.text, receipt: envelopeOf({}, ide, contract, note), singleDocumentRule: SINGLE_DOC_RULE }, null, 2);
    }
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
  if (!c) {
    // say WHY the persisted face was not used.  An invalidated cache that is
    // silently ignored is indistinguishable from a missing one.
    dbg('cachedDiscovery: face cache not used (' + faceCacheStatus.reason + ')');
    if (faceCacheStatus.reason === 'face_fingerprint_mismatch' ||
        faceCacheStatus.reason === 'snapshot_has_no_faceFingerprint') {
      process.stderr.write('vb6mcp-hub: persisted face cache INVALID (' + faceCacheStatus.reason +
        '); rebuilding the face from the live add-in.\n');
    }
    return null;
  }
  const checked = [];
  for (const i of c.instances) {
    // TCP-only: no netstat on a response path.  Transport aware: a pipe-only entry has no port
    // and used to be re-validated as dead here.
    const okp = await instanceAlive(i, null);
    if (okp) {
      // OPEN-27: a face cache file written before the banner was excluded may still carry
      // `instructions`; that text cannot be tied to the DLL running NOW, so it is never replayed.
      const clean = Object.assign({}, i);
      delete clean.instructions;
      checked.push(localCandidate(clean));
    }
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
    // (measured on the local SP6 box, 2026-10-16): a PIPE-transport instance publishes
    // port 0 ON PURPOSE - the pipe IS the transport.  Requiring a port here classified every live
    // pipe-only IDE as 'missing_pid_or_port' -> live=0, and ide_list fell back to
    // faceSource=snapshot even though the service answered (instances=1) and the add-in was
    // reachable through the pipe.  Liveness is now transport-aware: pipe => pid alive (the probe
    // still enriches it), tcp => pid alive AND port listening (unchanged).
    const isPipe = String(e.transport || '') === 'pipe';
    if (!e.pid || (!isPipe && !e.port)) { stale.push({ file: base, pid: e.pid, port: e.port, reason: 'missing_pid_or_port' }); continue; }
    const alive = pidAlive(e.pid);
    const listeningNow = e.port ? await isListening(e.port, listening || null) : true;
    dbg('registry ' + base + ' pid=' + e.pid + ' port=' + e.port + ' startedAt=' + (e.startedAt || '-') + ' pidAlive=' + alive + ' listening=' + listeningNow);
    if (!alive) { stale.push({ file: base, pid: e.pid, port: e.port, reason: 'pid_dead' }); continue; }
    if (!listeningNow) { stale.push({ file: base, pid: e.pid, port: e.port, reason: 'port_not_listening' }); continue; }
    // A pipe instance has no port, so key the duplicate check by pid (two live pipe IDEs must not
    // be mistaken for a duplicate-port pair).
    const dkey = e.port || ('pipe:' + e.pid);
    const prev = seen[dkey];
    // The LOSER goes to staleRegistry, never the winner.
    if (prev && !(String(e.startedAt || '') > String(prev.at || ''))) {
      stale.push({ file: base, pid: e.pid, port: e.port, reason: 'duplicate_port_in_registry' });
      continue;
    }
    if (prev) {
      // we are newer: the PREVIOUS owner is the one that gets demoted
      stale.push({ file: prev.file, pid: prev.pid, port: e.port, reason: 'superseded_older_duplicate_port' });
    }
    seen[dkey] = { at: String(e.startedAt || ''), file: base, pid: e.pid };
    let pids = (e.port && listening && listening[e.port]) || [];
    if (!pids.length) pids = [e.pid];
    byPort.delete(dkey);
    byPort.set(dkey, {
      port: e.port, pid: e.pid, pids: pids,
      // The TRANSPORT has to survive into the CANDIDATE, not just the registry decision: every
      // later re-validation (faceState / cachedDiscovery) asks the transport, because a pipe-only
      // instance has no port to test.  Without this the instance was re-validated as DEAD and
      // ide_list fell back to faceSource=snapshot with the IDE alive and reachable (measured).
      transport: e.transport || null,
      pipe: e.pipe || null,
      tcpEnabled: e.tcpEnabled === undefined ? null : e.tcpEnabled,
      sessionId: e.sessionId || null,
      instanceId: e.instanceId || null,
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
// instances[] is the AUTHORITATIVE IDE view.  A discovered listener that is not an
// add-in registration - above all the middle service itself, which the portscan finds on its own
// port and which answers tools/list from its snapshot - is NOT an IDE; it moves to
// nonIdeListeners[] with the basis that excluded it.  ONE place does this, and every local
// consumer goes through resolveDiscovery().
function splitByProcessKind(d, ctx) {
  if (!d || !Array.isArray(d.instances) || d.processKindSplit) return d;
  const ide = [], non = [];
  for (const i of d.instances) {
    const c = classifyProcess(i, ctx);
    const m = Object.assign({}, i, { processKind: c.processKind, isIde: c.isIde, processKindBasis: c.processKindBasis });
    (c.isIde ? ide : non).push(m);
  }
  if (non.length) dbg('splitByProcessKind: excluded ' + non.length + ' non-IDE listener(s): ' + non.map(function (x) { return x.port + '/' + x.processKind; }).join(','));
  return Object.assign({}, d, { instances: ide, nonIdeListeners: non, processKindSplit: true, excludedNonIde: non.length });
}
async function resolveDiscovery(d) {
  return splitByProcessKind(await resolveDiscoveryRaw(d), null);
}
async function resolveDiscoveryRaw(d) {
  if (d) { dbg('resolveDiscovery: using completed discovery instances=' + d.instances.length); return d; }
  const c = await cachedDiscovery();
  if (c) return c;
  const r = await registryCandidates();
  if (r) return r;
  dbg('resolveDiscovery: NOTHING found (snapshot fallback)');
  return { instances: [], conflicts: [], staleRegistry: [], registryDir: instancesDir(), registryUnreadable: false, scanTruncated: true };
}
async function faceState(d, spellings, ctx) {
  d = await resolveDiscovery(d);
  // the ONE liveness/ownership/process-kind pass, done here so EVERY receipt that
  // reports an instance carries the same fields AND the rule that produced them.
  if (!ctx) { try { ctx = await bridgeIdentity(); } catch (e) { ctx = bridgeIdentitySync() || {}; } }
  d = splitByProcessKind(d, ctx);
  const livenessContradictions = [];
  (d.instances || []).forEach(function (i) {
    const live = livenessOf(i, ctx);
    const own = ownershipOf(i);
    // P2: a SUCCESSFUL forwarded tools/call is liveness evidence too.  A pipe-only snapshot probe
    // can time out while the very same route answers a real call; without this merge ide_list
    // contradicted a working forward (N11).  probeStatusRaw stays the raw probe verdict.
    const fwdAt = lastForwardOk.get(Number(i.pid));
    if (fwdAt && (Date.now() - fwdAt) < FORWARD_EVIDENCE_TTL_MS && live.state !== 'responding') {
      live.state = 'responding';
      live.basis = String(live.basis || '') + ' + a forwarded tools/call SUCCEEDED at ' + new Date(fwdAt).toISOString();
      live.lastForwardOkAt = new Date(fwdAt).toISOString();
      live.evidenceByForward = true;
    }
    i.liveness = { state: live.state, basis: live.basis, evidenceAt: live.evidenceAt, evidenceAgeMs: live.evidenceAgeMs,
                   probeRoute: live.probeRoute, forwardRoute: live.forwardRoute, pipeState: live.pipeState, pipeAvailable: live.pipeAvailable, cached: live.cached };
    // The forward-evidence keys exist ONLY outside inproc, which must stay byte-for-byte unchanged.
    if (BRIDGE_MODE !== 'inproc' && live.evidenceByForward) {
      i.liveness.lastForwardOkAt = live.lastForwardOkAt || null;
      i.liveness.evidenceByForward = true;
    }
    i.probeStatus = live.probeStatus;
    i.probeStatusRaw = live.probeStatusRaw;
    if (live.probeReason) i.probeReason = live.probeReason;
    i.ownership = own.ownership;
    i.ownershipBasis = own.ownershipBasis;
    i.ownedByHub = own.owned;
    i.processKind = i.processKind || 'vb6_ide';
    i.processKindBasis = i.processKindBasis || 'addin_registration';
    live.contradictions.forEach(function (c) { livenessContradictions.push(c); });
  });
  // A cached/persisted face must NEVER freeze a dead instance: re-validate pid + port here,
  // on every request, before it is reported.  A REMOTE candidate is validated by its endpoint
  // probe instead - its pid/port live on ANOTHER machine, so a local pid test is meaningless
  // (a recycled remote pid must not silently drop a healthy remote IDE).
  if (d.instances && d.instances.length) {
    const alive = [];
    for (const i of d.instances) {
      if (i.remote) { alive.push(i); continue; }
      const okp = await instanceAlive(i, null);
      if (okp) alive.push(i); else dbg('faceState: dropped dead instance pid=' + i.pid + ' port=' + i.port + ' transport=' + (i.transport || '-'));
    }
    if (alive.length !== d.instances.length) d = Object.assign({}, d, { instances: alive, revalidated: true });
  }
  const snap = snapshot();
  // The live tool list is evidence ONLY when the add-in answered: a face read that the
  // middle service served from its own snapshot must not be diffed as if the IDE had replied.
  const answered = d.instances.filter(function (i) { return i.probeStatus === 'ok' && i.liveTools && i.liveTools.length; })[0] || null;
  const live = answered ? answered.liveTools : null;
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
  // ---- caller-spelling (alias) echo ----------------------------------------------------------
  // Per instance: the caller's spelling FIRST, then every spelling that lands on it.
  const aliasEcho = buildAliasEcho(spellings || [], d.instances);
  const aliasPerInstance = d.instances.map(function () { return []; });
  (aliasEcho.supplied || []).forEach(function (row) {
    if (!row.sameCandidate) return;
    d.instances.forEach(function (i, idx) {
      const hit = (row.matches || []).filter(function (m) { return m.pid === i.pid && m.port === i.port && m.endpoint === (i.endpoint || LOCAL_ENDPOINT); })[0];
      if (hit) aliasPerInstance[idx].push({ asGiven: row.asGiven, source: row.source, matchedBy: hit.matchedBy, via: hit.via, normalized: hit.normalized });
    });
  });
  const nonIde = d.nonIdeListeners || [];
  const responseCount = d.instances.filter(function (i) { return i.probeStatus === 'ok'; }).length;
  return {
// /: the face is LIVE only when an ADD-IN instance is present. Existence is
    // still independent of the probe (a pending probe is a live IDE), but a lone middle-service
    // listener is NOT an IDE and can no longer make this face look live (that was).
    faceSource: d.instances.length ? 'live' : 'snapshot',
    faceSourceBasis: d.instances.length
      ? (d.instances.length + ' add-in instance(s) discovered by registry/port scan; ' + responseCount + ' answered a fresh add-in probe')
      : (bridgeLastFailure
        ? ('NO INSTANCE VISIBLE, and this hub is NOT authorized to be sure: the middle service answered ' + bridgeLastFailure.kind + ' (' + bridgeLastFailure.reason + (bridgeLastFailure.status ? ', HTTP ' + bridgeLastFailure.status : '') + '). The empty list below is NOT evidence that no IDE is running.')
        : (nonIde.length ? 'no add-in instance; ' + nonIde.length + ' non-IDE listener(s) excluded (see nonIdeListeners)' : 'no instance discovered at all')),
    // the reason the instance view may be empty, in machine-readable form.
    bridgeFailure: bridgeLastFailure,
    authRequired: !!(bridgeLastFailure && bridgeLastFailure.kind === 'auth'),
    authReason: (bridgeLastFailure && bridgeLastFailure.kind === 'auth') ? bridgeLastFailure.reason : null,
    contractVersion: snap.contractVersion,
    toolCount: live ? live.length : snap.tools.length,
    snapshot: { contractVersion: snap.contractVersion, toolCount: snap.tools.length, generatedFrom: snap.generatedFrom,
                contractDeclaredToolCount: snap.contractDeclaredToolCount, versionToolCountMismatch: !!snap.versionToolCountMismatch },
    instances: d.instances.map(function (i, idx) { return {
      pid: i.pid, port: i.port, project: i.project, projectPath: i.projectPath || null, pids: i.pids,
      startedAt: i.startedAt || null,
      probeStatus: i.probeStatus || (i.pending ? 'pending' : 'ok'),
      probeStatusRaw: i.probeStatusRaw === undefined ? null : i.probeStatusRaw,
      liveToolCount: (i.pending || i.probeStatus !== 'ok') ? null : ((i.liveTools && i.liveTools.length) || null),
      probeReason: i.probeReason || null,
      // ---- the ONE liveness verdict + the evidence it came from ------------------
      addinAnswered: i.probeStatus === 'ok' || !!(i.liveness && i.liveness.evidenceByForward),
      ...(BRIDGE_MODE === 'inproc' ? {} : { lastForwardOkAt: (i.liveness && i.liveness.lastForwardOkAt) || null }),
      liveness: i.liveness || null,
      probeEvidenceAt: i.probeEvidenceAt || null,
      probeBasis: i.probeBasis || null,
      probeRoute: i.probeRoute || null,
      probeWarnings: (i.probeWarnings && i.probeWarnings.length) ? i.probeWarnings : null,
      // ---- the ONE ownership verdict (same predicate as ide_stop/ide_restart) -----
      ownership: i.ownership || 'not_owned',
      ownershipBasis: i.ownershipBasis || null,
      ownedByHub: !!i.ownedByHub,
      // ---- IDE vs middle service -----------------------------------------------
      processKind: i.processKind || 'vb6_ide',
      isIde: i.isIde !== false,
      processKindBasis: i.processKindBasis || null,
      startedByHub: startedByHub.has(i.pid) || !!i.startedByHub,
      discoveredBy: i.discoveredBy || 'portscan', registryFile: i.registryFile || null,
      version: i.version || null, dllSha256: i.dllSha256 || null, pending: !!i.pending,
      // the LIVE DLL identity the ADD-IN reported at call time (mcp_info.build.sha256), next to
      // the registration's own dllSha256 above - two independent readings of the same binary.
      dllSha256Live: i.dllSha256Live || null,
      dllSha256LivePath: i.dllSha256LivePath || null,
      dllSha256LiveReadAt: i.dllSha256LiveReadAt || null,
      addinProduct: i.addinProduct || null, addinContract: i.addinContract || null,
      // ---- the ONE-set identity (ADDITIVE): a remote IDE is an equal member -----------------
      endpoint: i.endpoint || LOCAL_ENDPOINT,
      // ADDITIVE, MULTI-BRIDGE ONLY: the bridge that owns this instance (the label a caller
      // addresses it by).  Single-bridge receipts keep their exact keys.
      ...(MULTI_BRIDGE ? { bridge: i.bridge || i.endpoint || LOCAL_ENDPOINT } : {}),
      remote: !!i.remote,
      endpointUrl: i.endpointUrl || null,
      targetKey: i.targetKey || null,
      pathVia: i.pathVia || null,
      // ---- (3): the canonical path NEXT TO the alias spellings (ADDITIVE) ----------
      projectPathReal: i.projectPath || null,
      projectPathAsGiven: i.projectPathAsGiven || i.projectPath || null,
      projectPathIsAlias: !!(i.projectPathAsGiven && normalizePathString(i.projectPathAsGiven).toLowerCase() !== String(i.projectPath || '').toLowerCase()),
      aliasAsGiven: (aliasPerInstance[idx] && aliasPerInstance[idx].length) ? aliasPerInstance[idx][0].asGiven : null,
      aliasSource: (aliasPerInstance[idx] && aliasPerInstance[idx].length) ? aliasPerInstance[idx][0].source : null,
      aliasMatchedBy: (aliasPerInstance[idx] && aliasPerInstance[idx].length) ? aliasPerInstance[idx][0].matchedBy : null,
      aliasesAsGiven: (aliasPerInstance[idx] || []).map(function (a) { return a.asGiven; }),
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
    // the registry side-effect policy is part of the receipt, not folklore.
    // MEASURED 2026-10-17 on this host: after ide_list the same registration file is byte-identical
    // (size/mtime/sha256 unchanged) and no .pruned/.bak residue appears. A registration is removed
    // ONLY by ide_restart, after a CONFIRMED stop of that same target (registryCleared=true).
    // WHO addresses the pipe-only IDEs, and whether a hub started/kept it.  Omitted
    // entirely in inproc mode, so that rollback receipt stays byte-for-byte unchanged.
    ...(function () { const m = middleServiceReceipt(); return m ? { middleService: m } : {}; })(),
    registryWrites: 0,
    registryPolicy: 'ide_list is READ-ONLY: it never writes, prunes or deletes a registration (%TEMP%\\vb6mcp\\instances\\<pid>.json). Cleanup happens only in ide_restart after a confirmed stop of its own target.',
// ---- /N11: WHICH source the authoritative IDE view came from
    // The hub's own registry and the middle service's /health read the SAME registration files but
    // as DIFFERENT user contexts, so they can disagree; naming the source is what makes the receipt
    // explain itself instead of "ide_list says X, a forwarded call says Y".
    instancesSource: (function () {
      const by = function (k) { return d.instances.filter(function (i) { return (i.discoveredBy || 'portscan') === k; }).length; };
      const reg = by('registry'), bh = by('bridge_health'), ps = by('portscan');
      const authoritative = reg ? 'hub_registry' : (bh ? 'middle_service_health' : (ps ? 'port_scan' : 'none'));
      return { fromRegistry: reg, fromBridgeHealth: bh, fromPortScan: ps, authoritative: authoritative,
        basis: authoritative === 'hub_registry'
          ? 'the hub read add-in registrations from its own %TEMP%\\vb6mcp\\instances'
          : (authoritative === 'middle_service_health'
            ? 'the hub\'s own registry view was EMPTY, so the add-in registrations were adopted from the middle service GET /health (it is the only addresser for a pipe-only IDE)'
            : (authoritative === 'port_scan' ? 'no registration was visible to the hub OR reported by the middle service; the candidate came from a listening-port scan' : 'no add-in instance was found at all')) };
    })(),
    // ---- the listener(s) that are NOT IDEs, with the basis that excluded them ------
    nonIdeListeners: (d.nonIdeListeners || []).map(function (i) { return {
      pid: i.pid, port: i.port, processKind: i.processKind, isIde: false, processKindBasis: i.processKindBasis,
      probeStatus: i.probeStatus || null, probeReason: i.probeReason || null,
      discoveredBy: i.discoveredBy || 'portscan', excludedFromSelection: true
    }; }),
    excludedNonIde: (d.nonIdeListeners || []).length,
    staleRegistry: d.staleRegistry || [],
    registryDir: d.registryDir || null,
    registryUnreadable: !!d.registryUnreadable,
    // ---- the RULES, restated verbatim wherever a verdict is reported --------------
    processKindRule: PROCESS_KIND_RULE,
    livenessRule: LIVENESS_RULE,
    ownershipRule: OWNERSHIP_RULE,
    targetResolutionRule: TARGET_RESOLUTION_RULE,
    livenessSummary: {
      verdict: d.instances.length === 0 ? 'no_ide' : (responseCount === d.instances.length ? 'all_responding' : (responseCount === 0 ? 'none_responding' : 'partly_responding')),
      instances: d.instances.length,
      responding: responseCount,
      notResponding: d.instances.filter(function (i) { return i.liveness && i.liveness.state === 'not_responding'; }).length,
      unknown: d.instances.filter(function (i) { return i.liveness && i.liveness.state === 'unknown'; }).length,
      basis: 'one verdict per instance: the ADD-IN must answer mcp_info over the same route a forwarded call uses (see livenessRule); the registry alone is never reported as a live answer'
    },
    livenessContradictions: livenessContradictions,
    selectionRule: SELECTION_RULE,
    targetKeyRule: TARGET_KEY_RULE,
    aliasEcho: aliasEcho,
    failureKindRule: FAILURE_KIND_RULE,
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

async function hubState(opts) {
  // P3: a READ-ONLY /health probe (never spawns, never writes the lock) so a hub that has just
  // restarted can still report the middle service before the first forward.
  if (BRIDGE_MODE !== 'inproc') { try { await middleServiceReportProbe(); } catch (e) { } }
  // ide_list is the ONE read that shows the whole set (local + configured remote endpoints).
  // tools/list + initialize deliberately stay on the LOCAL path so a slow remote can never
  // push the MCP handshake towards its 1s budget.
  const localD = await resolveDiscovery(await discoverSoft(LIST_GRACE_MS));
  // the middle service's transport verdict (pipeAvailable/pipeState per pid) is fetched
  // ONCE here and handed to faceState, which is what lets ONE liveness verdict cover both the
  // registry and the pipe reality (N11).  bridgeHealth() is cached, so this is not a new per-call
  // round trip in the steady state.
  let ctx = bridgeIdentitySync();
  if (BRIDGE_MODE !== 'inproc' && !ctx) { try { ctx = await bridgeIdentity(); } catch (e) { ctx = null; } }
  const d = await withRemoteCandidates(localD, REMOTE_GRACE_MS);
  // The caller's spelling (an explicit _target on this call, else the ide_target default) is
  // echoed back per instance ((3)); with no spelling supplied the receipt still shows
  // projectPathAsGiven next to projectPathReal.
  const st = await faceState(d, callerSpellings(opts && opts.requestTarget), ctx);
  if (REMOTE_ENDPOINTS.length) {
    st.remote = {
      configured: REMOTE_ENDPOINTS.length,
      instances: (d.instances || []).filter(function (i) { return !!i.remote; }).length,
      endpoints: (d.endpoints || []).map(function (e) { return Object.assign({}, e, { allowList: 'VB6MCP_REMOTES / bridge/hub_remotes.json (never discovered)' }); }),
      configErrors: d.endpointConfigErrors || ENDPOINT_CFG_ERRORS || []
    };
  }
  // ---- thin shell (d): the bridge's own instance / transport facts, ADDITIVE --------------
  // When the bridge is online it is authoritative about local IDEs (it is the only addresser),
  // so its transport / pipeAvailable / staleRegistry are merged in.  NEW keys only; no existing
  // key is removed or renamed (ADR sec 4.2.3 / sec 6: ide_list only ever grows).
  if (BRIDGE_MODE !== 'inproc') {
    const h = await bridgeHealth();
    if (h && h.instances) {
      st.bridge = {
        url: BRIDGE_URL, port: (h && h.port) || null, transport: (h && h.transport) || null,
        netstat: false, faceTools: (h && h.faceTools) || null,
        liveInstances: ((h && h.liveInstances) || []).length,
        instances: h.instances || [], staleRegistry: (h && h.staleRegistry) || [],
        // the middle service is a PROCESS, not an IDE.  Its address and the pid that
        // owns its port are named here, and every entry of bridge.instances[] is an add-in
        // REGISTRATION (the middle service reads the same registry the hub does).
        middleService: {
          url: BRIDGE_URL, port: (h && h.port) || ctx && ctx.port || bridgePortHint() || null,
          pids: (ctx && ctx.pids) || [], processKind: 'middle_service', isIde: false,
          basis: (ctx && ctx.basis) || 'VB6MCP_BRIDGE_URL'
        },
        processKindRule: PROCESS_KIND_RULE
      };
      const bridgeByPid = {};
      for (const bi of (h.instances || [])) { if (bi && bi.pid) bridgeByPid[String(bi.pid)] = bi; }
      st.bridge.instances = (h.instances || []).map(function (bi) {
        const isService = !!(bi && bi.pid && (ctx && ctx.pids || []).indexOf(Number(bi.pid)) >= 0);
        return Object.assign({}, bi, { processKind: isService ? 'middle_service' : 'vb6_ide', isIde: !isService,
          processKindBasis: isService ? 'pid_owns_bridge_port(' + ((ctx && ctx.port) || '?') + ')' : 'add-in instance REGISTRATION read by the middle service' });
      });
      for (const i of (st.instances || [])) {
        const bi = bridgeByPid[String(i.pid)];
        if (!bi) continue;
        if (i.transport === undefined) i.transport = bi.transport;
        if (i.pipeAvailable === undefined) i.pipeAvailable = bi.pipeAvailable;
        if (i.pipeState === undefined) i.pipeState = bi.pipeState;
      }
    }
  }
  // Zero-config (zero-config addressing): the discovery verdict is part of ide_list's additive diagnostics,
  // so "which endpoint did the hub adopt, on what evidence, with which candidates" is queryable.
  if (BRIDGE_MODE !== 'inproc') {
    const bd = bridgeDiscoveryReceipt();
    if (bd) st.bridgeDiscovery = bd;
  }
  // The ONE attribution verdict on every instance of the authoritative IDE view AND
  // on every RESIDUAL registration (staleRegistry), so ide_list / the forwarded ide_state /
  // ide_restart cannot disagree about the same instance.  ide_list is the only consumer of
  // hubState(), so this never slows the initialize/tools-list handshake.
  if (Array.isArray(st.instances)) {
    for (const r of st.instances) {
      try { r.ownershipVerdict = await resolveOwnership(r); } catch (e) { r.ownershipVerdict = null; }
    }
    st.ownershipVerdictRule = OWNERSHIP_VERDICT_RULE;
  }
  if (Array.isArray(st.staleRegistry)) {
    for (const s of st.staleRegistry) {
      if (!s || typeof s !== 'object') continue;
      try { s.ownershipVerdict = await resolveOwnership({ pid: s.pid, port: s.port, registryFile: s.file }); } catch (e) { s.ownershipVerdict = null; }
    }
  }
  // ---- multi-bridge: ide_list.bridges[] and the config errors (ADDITIVE; only when configured) ----
  if (BRIDGES_ENV_PRESENT) {
    st.multiBridge = MULTI_BRIDGE;
    st.bridgeConfig = {
      env: 'VB6MCP_BRIDGES', configured: BRIDGE_ENDPOINTS.length,
      localSlot: 'VB6MCP_BRIDGE_URL', remotesFile: REMOTES_PATH ? path.basename(REMOTES_PATH) : null,
      remotesFileDisabled: REMOTES_FILE_DISABLED, rule: BRIDGES_RULE,
      errors: (d.endpointConfigErrors || ENDPOINT_CFG_ERRORS || []),
      configuredBridges: BRIDGE_ENDPOINTS.map(function (e) { return e.label; })
    };
    st.bridges = [localBridgeStatusRow()].concat(BRIDGE_ENDPOINTS.map(bridgeStatusRow));
    st.bridgeFailures = st.bridges.filter(function (b) { return b.probeStatus === 'failed'; })
      .map(function (b) { return { name: b.name, reason: b.reason, failureKind: b.failureKind, suggestedAction: b.suggestedAction, retrySafe: b.retrySafe, authReason: b.authReason }; });
  }
  // P1: the receipt can PROVE which code answered (sha16 of the running file + start time).
  const hcStamps = hubCodeStamps();
  if (hcStamps) st.hubCode = hcStamps;
  return st;
}

// ---- Segment 4: cross-bridge CONTROL-PLANE facts --------------------------------------------
// Which bridge instances carry this pid?  A pid is only unique WITHIN one machine, so the same
// number can exist on several bridges; every control verb below handles that explicitly.
async function remoteInstancesForPid(pid) {
  if (!MULTI_BRIDGE || !pid) return [];
  try {
    const dd = await resolveDiscovery(await discoverSoft(LIST_GRACE_MS));
    const merged = await withRemoteCandidates(dd, REMOTE_GRACE_MS);
    return (merged.instances || []).filter(function (i) { return !!i.remote && Number(i.pid) === Number(pid); });
  } catch (e) { return []; }
}
function foreignRemoteVerdict(hit) {
  const bridge = (hit && (hit.bridge || hit.endpoint)) || null;
  return { verdict: 'foreign', ownershipBasis: 'remote_bridge_instance', bridge: bridge,
           startedByHub: !!(hit && hit.startedByHub), remote: true,
           pid: (hit && hit.pid) || null, port: (hit && hit.port === undefined ? null : hit.port) || null,
           project: (hit && hit.project) || null, projectPath: (hit && hit.projectPath) || null,
           registryAlive: false, portAvailable: null, registryPresent: false, pidAlive: null, probeAnswered: true,
           basis: ['the instance lives on bridge "' + bridge + '": its pid/port/process belong to ANOTHER machine',
                   'startedByHub=' + (hit && hit.startedByHub ? 'true' : 'false') + ' (published by the owning bridge); this hub has no in-memory record of it',
                   'a remote bridge instance is never attributable to THIS machine hub process'],
           rule: OWNERSHIP_VERDICT_RULE };
}
function remoteControlRefusal(tool, pid, hit) {
  const bridge = (hit && (hit.bridge || hit.endpoint)) || null;
  const common = {
    ok: false, tool: tool, pid: Number(pid), remote: true,
    endpoint: (hit && hit.endpoint) || bridge, bridge: bridge,
    ownership: 'not_owned', ownershipBasis: 'remote_bridge_instance', ownedByHub: false,
    startedByHub: !!(hit && hit.startedByHub),
    port: (hit && hit.port === undefined) ? null : hit.port,
    portRuleApplies: false,
    portBasis: 'not applicable: a remote instance port belongs to another machine and is never used as an ownership criterion here',
    ownershipRule: OWNERSHIP_RULE,
    ownershipVerdict: foreignRemoteVerdict(hit), ownershipVerdictValue: 'foreign', ownershipVerdictRule: OWNERSHIP_VERDICT_RULE,
    liveness: (hit && hit.liveness) || null, instanceSource: 'bridge_probe',
    alsoForeignPort: false,
    failureKind: 'toolError', failureKindRule: FAILURE_KIND_RULE,
    suggestedAction: 'fix_params', retrySafe: false,
    message: 'pid ' + pid + ' belongs to the IDE on bridge "' + bridge + '" (ANOTHER machine, and it was not started by this hub). ' +
      (tool === 'ide_stop' ? 'Nothing was stopped' : 'Nothing was restarted') + ': this hub only controls instances IT started on THIS machine.',
    nextStep: 'control that IDE on the machine that owns it (bridge-side control verbs are not implemented in this build); the hub never signals a process on another machine'
  };
  if (tool === 'ide_stop') return Object.assign(common, { stopped: 0, reason: 'not_allowed', nothingStopped: true });
  return Object.assign(common, { restarted: false, stopConfirmed: false, newPid: 0, refused_reason: 'not_owned', nothingRestarted: true });
}
function finishControlOut(out, snap) {
  // ADDITIVE (multi-bridge only): every control-plane failure names failureKind + the next
  // ACTION + whether a retry can help, matching the v1.12.2 failure envelope vocabulary.
  if (MULTI_BRIDGE) {
    if (!out.failureKind) out.failureKind = 'toolError';
    if (!out.failureKindRule) out.failureKindRule = FAILURE_KIND_RULE;
    if (!out.suggestedAction) {
      out.suggestedAction = (out.refused_reason === 'ambiguous_target' || out.refused_reason === 'target_conflict') ? 'fix_params'
        : (out.refused_reason === 'kill_failed' ? 'retry_once' : 'ide_start');
    }
    if (out.retrySafe === undefined) out.retrySafe = out.refused_reason === 'kill_failed';
  }
  return textResult(out, snap.contractVersion);
}
function remoteAmbiguousRefusal(tool, pid, localHit, remoteHits) {
  const cands = (localHit ? [candidateBrief(localHit)] : []).concat((remoteHits || []).map(candidateBrief));
  const insts = (localHit ? [localHit] : []).concat(remoteHits || []);
  const base = { ok: false, tool: tool, pid: Number(pid), reason: 'ambiguous_target', refused_reason: 'ambiguous_target',
                 failureKind: 'toolError', failureKindRule: FAILURE_KIND_RULE,
                 suggestedAction: 'fix_params', retrySafe: false, candidates: cands,
                 suggestion: suggestionFor(insts), conflicts: [],
                 message: 'pid ' + pid + ' exists on ' + cands.length + ' IDEs across different bridges (' + cands.map(function (c) { return (c.bridge || 'local') + ':' + c.pid; }).join(', ') + ') -- check ide_list and stop the instance on the machine that owns it; this hub signals NOTHING' };
  if (tool === 'ide_stop') return Object.assign(base, { stopped: 0, nothingStopped: true });
  return Object.assign(base, { restarted: false, stopConfirmed: false, newPid: 0, nothingRestarted: true });
}


// ============================================================================================
// (WP-22-01) TARGET RESOLUTION vs OFFLINE ROUTING
// --------------------------------------------------------------------------------------------
// The default target ladder answers ONE question: "which LIVE IDE does an unqualified call
// mean".  It must never answer "is there an IDE at all" with a STALE pid.  Before this change
// the ladder could hand the middle service a canonical selector taken from a candidate that no
// longer existed (a persisted face, a recycled pid, the hub's own bookkeeping), and the caller
// saw reason=target_mismatch plus a pid for a process that was already gone - while the same
// machine had NO IDE at all and a command-line /make build would have worked.
// Two separate rules now:
//   1) TARGET RESOLUTION is LIVE-ONLY: dead candidates are dropped BEFORE the ladder runs, and
//      every target refusal names the pids that ARE online (onlinePids) plus a next action;
//   2) OFFLINE ROUTING: a tool that does not need the add-in (build / make_project route=cli /
//      project_meta) never resolves a target at all when no live IDE exists; the hub serves it
//      itself and says so (offline:true, targetResolved:false).
// ============================================================================================
const OFFLINE_ROUTING_RULE = 'offline routing = when NO live IDE is visible, an IDE-FREE tool is served by the hub itself and the receipt carries offline:true, offlineRoute (cli|file), targetResolved:false and targetDiagnostic{reason,onlinePids,howToProceed}. A tool that cannot be served offline refuses with reason=no_live_ide (never target_mismatch, never a pid) + offlineServed:false + howToProceed. build and make_project(route != ide) fall back to an independent VB6.EXE /make subprocess; project_meta reads/writes the .vbp root keys on disk. Tools that need the add-in keep the documented target ladder.';
const NO_LIVE_IDE_RULE = 'no_live_ide = the hub could PROVE there is no live IDE to address (0 candidates after pid/transport re-validation, and the middle service reported no instance either). It is NOT target_mismatch: target_mismatch means a selector the caller passed did not match a LIVE set, and it is emitted only when at least one live IDE exists OR when the caller passed an explicit selector that names nothing. Every no_live_ide / target_mismatch receipt carries onlinePids[], howToProceed[] and suggestion{action,useTarget}.';
const TARGET_DIAGNOSIS_RULE = 'targetDiagnosis distinguishes the causes of a target failure: explicit_target_not_live (the caller named a pid/project/port that is not in the live set - the live set rides along in onlinePids/candidates) | stale_default_ignored (the ide_target DEFAULT named something that vanished; it is IGNORED and never vetoes the live set) | no_live_ide (nothing is live at all) | hub_record_stale (the bookkeeping named a pid that is gone; live discovery corrected it and the correction is reported) | ambiguous_target / target_conflict (the selector matched SEVERAL live IDEs: a live-set problem, not a missing/dead target).';
const IDE_START_READY_RULE = 'ide_start readiness: ok:true + started:true means THE PROCESS WAS CREATED. ready:true means the add-in was OBSERVED answering a tools/call (mcp_info) for that pid within ready_wait_ms (default 15000, env VB6MCP_IDE_START_READY_MS). ready:false is a truthful timeout: the receipt carries readyWaitMs, readyChecks{registryVisible,transportUp,addInAnswers,addInAnswersBasis,pidAlive,projectMatches} and pollingHint, and the pid is still valid - poll ide_list until that pid appears. A fixed sleep is never used as a readiness test.';
const READY_PROBE_RULE = 'readiness is measured on the STARTED pid ONLY: every poll round re-reads that pid registration file from disk, re-tests its transport and then asks the add-in mcp_info - the discovered instance set is NEVER consulted, because it is produced by probing every candidate and can lag the pid this call created. addInAnswersBasis is exactly one of no_route | direct_rpc | middle_service_forward; no_route means there was NO channel to ask, so addInAnswers is null (not measured) and never false. A budget-exhausted poll returns ready:false with the measured components, and the pid is still valid.';
const RELAUNCH_RULE = 'ide_kill_and_relaunch = the hub-side, add-in-independent reset primitive: it STOPS one instance that this hub owns (in-memory startedByHub OR the registration itself says startedByHub=true), clears that registration, launches VB6.EXE on the SAME .vbp and polls until the add-in answers. It never signals an instance it does not own, never touches a port below 8767 (the reserved range that includes 8766), and returns stoppedPid / newPid / ready / threeWay / elapsedMs on every path.';
const OFFLINE_TOOLS = Object.freeze({
  build:        { impl: 'cli_build', route: 'cli' },
  make_project: { impl: 'cli_build', route: 'cli', cliOnly: true },
  project_meta: { impl: 'vbp_meta',  route: 'file' },
  package_app:  { impl: null,        route: 'file' },
  verify:       { impl: null,        route: 'disk+pane' }
});
function offlineRouteFor(name, args) {
  const t = OFFLINE_TOOLS[name];
  if (!t) return null;
  const a = args || {};
  if (name === 'make_project') {
    const route = String(a.route === undefined || a.route === null || a.route === '' ? 'auto' : a.route).toLowerCase();
    if (route === 'ide') return { tool: name, offlineServed: false, reason: 'route_ide_requires_live_ide', requestedRoute: route, impl: null, route: 'ide' };
    return { tool: name, offlineServed: !!t.impl, impl: t.impl, route: 'cli', requestedRoute: route };
  }
  return { tool: name, offlineServed: !!t.impl, impl: t.impl, route: t.route, requestedRoute: null };
}
// ---- (1) LIVE-ONLY reconciliation ----------------------------------------------------------
// A candidate is alive when its pid exists AND its transport answers the SAME cheap test the
// discovery pass uses (TCP listening, or pid for a pipe-only registration).  This is deliberately
// the ONLY liveness rule here: a probe of the add-in is NOT required to pick a target (a wedged
// IDE must still be addressable so it can be recovered), but a DEAD pid must never be picked.
async function reconcileLiveInstances(insts, listening) {
  const list = (insts || []).filter(function (i) { return !!i; });
  const flags = await Promise.all(list.map(function (i) {
    return instanceAlive(i, listening || null).catch(function () { return false; });
  }));
  const keep = [], dropped = [];
  for (let k = 0; k < list.length; k++) {
    if (flags[k]) { keep.push(list[k]); continue; }
    const i = list[k];
    dropped.push({ pid: (i.pid === undefined ? null : i.pid), port: (i.port === undefined ? null : i.port),
                   project: i.project || null, projectPath: i.projectPath || null,
                   discoveredBy: i.discoveredBy || null,
                   droppedReason: pidAlive(i.pid) ? 'transport_dead' : 'pid_dead' });
  }
  return { instances: keep, dropped: dropped };
}
function onlinePidsOf(insts) {
  const out = [];
  for (const i of (insts || [])) {
    const n = Number(i && i.pid);
    if (!n) continue;
    if (out.indexOf(n) < 0) out.push(n);
  }
  return out;
}
// (35) CWD DECLARATION - the working directory a launch was spawned FROM, captured AT THE SPAWN
// POINT.  VB6 may write a .vbg into that directory during a graceful exit, so the receipt names
// the directory it is responsible for and nothing more.  If the caller passed opts.cwd that value
// is used; otherwise the child inherits THIS hub process's cwd (no cwd option is passed), so
// process.cwd() at that moment IS the spawn cwd.  The basis is exactly one of "spawn_cwd" (the
// value was read above) or "unavailable" (it could not be read, or this path started NO process);
// the value is never guessed and is never read from the child's PEB.
function hubSpawnCwd(opts) {
  try {
    // An EXPLICIT cwd wins.  A present-but-unnameable one (null/empty/non-string, e.g. a URL)
    // is NOT replaced by an inherited value: the child did not inherit, so the receipt must not
    // pretend it did - fail closed with null instead.
    if (opts && opts.cwd !== undefined && opts.cwd !== null) {
      if (typeof opts.cwd === 'string' && opts.cwd.length) return opts.cwd;
      return null;
    }
    const d = process.cwd();
    return (typeof d === 'string' && d.length) ? d : null;
  } catch (e) { return null; }
}
// ONE process launcher.  A .cmd/.bat path (a wrapper script named by VB6MCP_VB6EXE) has to be
// started through the command processor on Windows; a real VB6.EXE is spawned directly.  Both
// paths are detached-safe and never use a shell for anything except that wrapper form.
function spawnProgram(exe, args, opts) {
  const o = Object.assign({ windowsHide: true, stdio: 'ignore' }, opts || {});
  // A batch wrapper HAS to be started through the command processor; Node does the /d /s /c
  // quoting correctly for shell:true, and a manual "cmd /c \"line\"" build mangles the first
  // and last quote (measured: the fake compiler exited 1 and never wrote its log).
  if (/\.(cmd|bat)$/i.test(String(exe))) return spawn(String(exe), args || [], Object.assign({}, o, { shell: true }));
  return spawn(exe, args || [], o);
}

// ---- (2) text / .vbp helpers (byte-preserving) ---------------------------------------------
// ONE decode rule, the same one the add-in uses: valid UTF-8 wins, else system ANSI (GBK), else
// latin1 as the last resort (reported, never silent).
function decodeTextBytes(buf) {
  try { return { text: new TextDecoder('utf-8', { fatal: true }).decode(buf), encoding: 'utf-8' }; } catch (e) { }
  try { return { text: new TextDecoder('gbk').decode(buf), encoding: 'gbk' }; } catch (e) { }
  return { text: Buffer.from(buf).toString('latin1'), encoding: 'latin1' };
}
function readTextFileAny(p) {
  const buf = fs.readFileSync(p);
  const d = decodeTextBytes(buf);
  return { text: d.text, encoding: d.encoding, bytes: buf.length };
}
function encodeTextBytes(text, encoding) {
  if (encoding === 'utf-8') return Buffer.from(text, 'utf8');
  if (encoding === 'latin1') return Buffer.from(text, 'latin1');
  for (let i = 0; i < text.length; i++) if (text.charCodeAt(i) > 127) return null;
  return Buffer.from(text, 'ascii');
}
const VBP_COMPONENT_PREFIXES = ['Module=', 'Class=', 'Form=', 'UserControl=', 'RelatedDoc=', 'ResFile=',
                                'Reference=', 'Object=', 'Designer=', 'UserDocument=', 'PropertyPage='];
function vbpIsComponentLine(t) {
  for (const p of VBP_COMPONENT_PREFIXES) {
    if (t.length >= p.length && t.slice(0, p.length).toLowerCase() === p.toLowerCase()) return true;
  }
  return false;
}
function vbpScalarKeyOf(line) {
  const t = String(line === undefined || line === null ? '' : line).trim();
  if (!t) return null;
  if (t.charAt(0) === '[') return null;
  if (vbpIsComponentLine(t)) return null;
  const eq = t.indexOf('=');
  if (eq <= 1) return null;
  const k = t.slice(0, eq).trim();
  if (!k) return null;
  return { key: k, value: t.slice(eq + 1).trim() };
}
function vbpLineEnding(text) {
  if (text.indexOf('\r\n') >= 0) return '\r\n';
  if (text.indexOf('\n') >= 0) return '\n';
  return '\r\n';
}
function vbpParse(vbpPath) {
  const r = readTextFileAny(vbpPath);
  const lines = r.text.split(/\r\n|\n|\r/);
  const keys = {};
  let section = '';
  const sections = {};
  for (const raw of lines) {
    const t = String(raw).trim();
    if (t.charAt(0) === '[' && t.charAt(t.length - 1) === ']') {
      section = t.slice(1, -1).toLowerCase();
      sections[section] = sections[section] || {};
      continue;
    }
    const kv = vbpScalarKeyOf(t);
    if (!kv) continue;
    if (section) { if (sections[section][kv.key] === undefined) sections[section][kv.key] = kv.value; }
    else if (keys[kv.key] === undefined) keys[kv.key] = kv.value;
  }
  return { keys: keys, sections: sections, lineEnding: vbpLineEnding(r.text), encoding: r.encoding, bytes: r.bytes, text: r.text };
}
function sha256File(p) {
  return crypto.createHash('sha256').update(fs.readFileSync(p)).digest('hex').toUpperCase();
}
function fileStatSafe(p) {
  try { const st = fs.statSync(p); return { exists: true, bytes: st.size, mtimeMs: st.mtimeMs, mtime: new Date(st.mtimeMs).toISOString() }; }
  catch (e) { return { exists: false, bytes: 0, mtimeMs: 0, mtime: null }; }
}
// A file view WITHOUT any "is it new" heuristic.  Success may only come from a MEASURED change
// (artifactChangedIn) plus a digest read-back; sampledAt lets hashedAfterRewrites be derived from
// timestamps instead of being a constant.
function artifactView(p) {
  const st = fileStatSafe(p);
  const v = { path: p, exists: st.exists, bytes: st.bytes, mtime: st.mtime, mtimeMs: st.mtimeMs };
  v.sha256 = st.exists ? sha256File(p) : '';
  v.sampledAt = Date.now();
  return v;
}
function artifactChangedIn(pre, post) {
  if (!post || !post.exists || post.bytes <= 0) return false;
  if (!pre || !pre.exists) return true;
  if (pre.bytes !== post.bytes) return true;
  if (pre.sha256 !== post.sha256) return true;
  if (pre.mtimeMs !== post.mtimeMs) return true;
  return false;
}
// ---- (3) PE / LAA (the same bytes the add-in patches; no EditBin) --------------------------
function peLaaRead(p) {
  const buf = fs.readFileSync(p);
  if (buf.length < 64) return { ok: false, error: 'not a PE image: file is shorter than the DOS header' };
  if (buf.readUInt16LE(0) !== 0x5A4D) return { ok: false, error: 'not a PE image: the DOS signature at offset 0 is not MZ' };
  const lfanew = buf.readInt32LE(0x3C);
  if (lfanew <= 0) return { ok: false, error: 'not a PE image: e_lfanew reads ' + lfanew };
  if (lfanew + 24 > buf.length) return { ok: false, error: 'truncated PE header: e_lfanew=' + lfanew + ' but the file is only ' + buf.length + ' bytes' };
  if (buf.readUInt32LE(lfanew) !== 0x00004550) return { ok: false, error: 'not a PE image: no PE signature at e_lfanew' };
  const off = lfanew + 22;
  return { ok: true, chars: buf.readUInt16LE(off), charsOff: off };
}
function laaPatch(p, want) {
  const d = { artifact: p, requested: !!want, applicable: false, laaBefore: false, laaAfter: false, laaSet: false,
              already: false, bytesPatched: 0, peOk: false, error: '', note: '',
              laaSkippedReason: '', laaRefusedReason: '' };
  if (!want) { d.note = 'laa=false: the caller switched the LAA step off for this call; the artifact was not examined'; return d; }
  if (!fs.existsSync(p)) { d.error = 'laa: the artifact does not exist on disk yet (' + p + '), so there is no PE header to patch'; return d; }
  const rd = peLaaRead(p);
  d.peOk = !!rd.ok;
  if (!rd.ok) { d.error = 'laa: ' + rd.error; return d; }
  d.applicable = true;
  const isLaa = (rd.chars & 0x20) !== 0;
  d.laaBefore = isLaa; d.laaAfter = isLaa;
  if (isLaa) { d.already = true; d.note = 'idempotent: the artifact already carries IMAGE_FILE_LARGE_ADDRESS_AWARE (0x0020); nothing was written'; return d; }
  let fd = null;
  try {
    fd = fs.openSync(p, 'r+');
    const two = Buffer.alloc(2);
    two.writeUInt16LE((rd.chars | 0x20) & 0xFFFF, 0);
    fs.writeSync(fd, two, 0, 2, rd.charsOff);
  } catch (e) {
    d.error = 'laa: could not write the PE characteristics word: ' + String((e && e.message) || e);
    try { if (fd !== null) fs.closeSync(fd); } catch (e2) { }
    return d;
  }
  try { fs.closeSync(fd); } catch (e) { }
  const rd2 = peLaaRead(p);
  if (rd2.ok && (rd2.chars & 0x20) !== 0) {
    d.laaAfter = true; d.laaSet = true; d.bytesPatched = 2;
    d.note = 'IMAGE_FILE_LARGE_ADDRESS_AWARE (0x0020) was written into the PE header by the hub (no EditBin.exe) and the file was re-read to verify';
    return d;
  }
  try {
    const fd2 = fs.openSync(p, 'r+');
    const back = Buffer.alloc(2); back.writeUInt16LE(rd.chars & 0xFFFF, 0);
    fs.writeSync(fd2, back, 0, 2, rd.charsOff); fs.closeSync(fd2);
    d.error = 'laa: the write was not confirmed by a re-read; the original characteristics word was put back';
  } catch (e) {
    d.error = 'laa: the write was not confirmed by a re-read AND the rollback failed (' + String((e && e.message) || e) + ')';
  }
  d.laaAfter = false; d.laaSet = false;
  return d;
}
// LAA decision precedence, the same as the add-in: per-call argument > the [VB6MCP] section of
// the .vbp (LAA=1) > VB6MCP_LAA > global default ON.  An unreadable value is a CONFIGURATION
// ERROR that skips the step and says so - never a silent fallback.
function offlineResolveLaa(args, vbpSections) {
  const envRaw = process.env.VB6MCP_LAA;
  if (args && Object.prototype.hasOwnProperty.call(args, 'laa') && args.laa !== undefined && args.laa !== null) {
    return { want: !!args.laa, source: 'param', error: '' };
  }
  const sec = (vbpSections && vbpSections['vb6mcp']) || {};
  const proj = sec.LAA !== undefined ? sec.LAA : sec.laa;
  if (proj !== undefined && proj !== null && String(proj).trim() !== '') {
    const t = String(proj).trim().toLowerCase();
    if (t === '1' || t === 'true' || t === 'yes' || t === 'on') return { want: true, source: 'vb-project', error: '' };
    if (t === '0' || t === 'false' || t === 'no' || t === 'off') return { want: false, source: 'vb-project', error: '' };
    return { want: false, source: 'vb-project-error', error: 'the [VB6MCP] LAA value in the .vbp is not true/false, so the LAA setting is UNKNOWN; the LAA step was SKIPPED' };
  }
  if (envRaw !== undefined && String(envRaw).trim() !== '') {
    const t = String(envRaw).trim().toLowerCase();
    if (t === '1' || t === 'true' || t === 'yes' || t === 'on') return { want: true, source: 'env', error: '' };
    if (t === '0' || t === 'false' || t === 'no' || t === 'off') return { want: false, source: 'env', error: '' };
    return { want: false, source: 'env-error', error: 'VB6MCP_LAA is set to a value that is not true/false, so the LAA setting is UNKNOWN; the LAA step was SKIPPED' };
  }
  return { want: true, source: 'global-default', error: '' };
}

// ---- (4) VB6.EXE /make ---------------------------------------------------------------
const OFFLINE_BUILD_DEFAULT_TIMEOUT_MS = 180000;
const OFFLINE_BUILD_MAX_TIMEOUT_MS = 600000;
// EXACTLY the add-in's marker set (Build.IsCompileErrorLine): 编译错误 or the English
// 'compile error'.  翻译错误 is NOT a VB6 marker and must not be accepted here (parity).
const COMPILE_ERR_RE = /(?:编译错误|compile error)/i;
const COMPILE_FAIL_RE = /(?:编译失败|生成失败|build failed|failed)/i;
const COMPILE_BANNER_RE = /(?:Microsoft \(R\) Visual Basic|Copyright \(C\))/i;
// PARITY with the add-in (Build.BetweenQuotes / Build.LineAfter / Build.FileLineFrom):
// a file name AND a line number are taken only when BOTH exist, and the number must sit
// behind the localized 行 / English "line" marker.  Without that requirement a quoted token
// inside an error sentence plus ANY later digit became a fabricated file/line (measured trap:
// "变量 'Foo' 未定义 (位于 Form1.frm)" -> file=Foo line=1, while the add-in reports file='').
function compileLogQuoted(text) {
  let a = text.indexOf("'");
  if (a < 0) a = text.indexOf('"');
  if (a < 0) return '';
  let b = text.indexOf("'", a + 1);
  if (b < 0) b = text.indexOf('"', a + 1);
  if (b < 0) return '';
  return text.slice(a + 1, b);
}
function compileLogLineAfter(text) {
  // Build.LineAfter: 行 N first, then English "line N"; the FIRST digit run after the
  // marker wins.  A bare "(N)" grouping is NOT a line number for the add-in.
  let p = text.indexOf('行 ');
  let off = 2;
  if (p < 0) {
    p = text.toLowerCase().indexOf('line ');
    if (p < 0) return 0;
    off = 5;
  }
  let num = '';
  for (let i = p + off; i < text.length; i++) {
    const ch = text.charAt(i);
    if (ch >= '0' && ch <= '9') num += ch;
    else if (num) break;
  }
  return num ? Number(num) : 0;
}
function compileLogFileLine(text) {
  const file = compileLogQuoted(text);
  const line = compileLogLineAfter(text);
  if (file && line > 0) return { file: file, line: line };
  // HUB-ONLY EXTENSION (intentional, documented): the unquoted CLI shape `Form1.frm(7)`
  // that the add-in's parser does not cover.  It only fires when the add-in-compatible rule
  // above found nothing, so it never contradicts the product rule for the quoted shapes.
  const m2 = /((?:[A-Za-z]:)?[^\s()'"\r\n]+?\.(?:bas|cls|frm|ctl|pag|dsr))\((\d+)\)/i.exec(text);
  if (m2) return { file: m2[1], line: Number(m2[2]) };
  return { file: '', line: 0 };
}
function compileLogMsgAfter(line) {
  let p = line.indexOf(' : ');
  if (p >= 0) return line.slice(p + 3).trim();
  p = line.indexOf(':');
  if (p >= 0) return line.slice(p + 1).trim();
  return '';
}
// ONE classifier for a VB6 /make log, ported from the add-in rule set (Build.bas ClassifyLog):
// marker lines (localized or English) first; otherwise the LAST non-empty line is the structural
// failure summary and the error text is at most 3 non-empty lines above it, skipping banners.
function parseCompilerLog(raw) {
  const lines = String(raw === undefined || raw === null ? '' : raw).replace(/\r\n/g, '\n').replace(/\r/g, '\n').split('\n');
  const isBanner = function (ln) {
    const t = String(ln).trim();
    if (!t) return true;
    if (COMPILE_BANNER_RE.test(t)) return true;
    if (COMPILE_ERR_RE.test(t)) return false;
    if (COMPILE_FAIL_RE.test(t)) return true;
    if (t.length <= 3) return true;
    if (t.charAt(1) === ':') return true;
    return false;
  };
  let last = -1, markerLine = '', markerCount = 0;
  for (let i = 0; i < lines.length; i++) {
    const ln = lines[i];
    if (String(ln).trim()) last = i;
    if (COMPILE_ERR_RE.test(ln)) { markerCount++; if (!markerLine) markerLine = ln; }
  }
  let matchRule = 'none', rawLine = '', failAnchor = false;
  if (markerLine) { rawLine = markerLine; matchRule = 'err-marker'; }
  else if (last >= 0 && COMPILE_FAIL_RE.test(lines[last])) {
    failAnchor = true;
    let j = last - 1, steps = 0;
    while (j >= 0 && steps < 3) {
      if (String(lines[j]).trim()) { steps++; if (!isBanner(lines[j])) break; }
      j--;
    }
    if (j >= 0 && steps > 0 && !isBanner(lines[j])) { rawLine = lines[j]; matchRule = 'fail-marker-up'; }
  }
  let message = '', file = '', reportedLine = 0;
  if (rawLine) {
    if (matchRule === 'err-marker') { message = compileLogMsgAfter(rawLine) || String(rawLine).trim(); }
    else { message = String(rawLine).trim(); }
    // file and reportedLine stay BOTH-OR-NEITHER, exactly like Build.FileLineFrom: a line
    // number without a file name must never be published as a location.
    const fl = compileLogFileLine(rawLine);
    file = fl.file; reportedLine = fl.line;
  }
  const errorCount = markerCount > 0 ? markerCount : (failAnchor ? 1 : 0);
  let excerpt = '';
  if (!rawLine && !message) {
    const tail = [];
    for (let i = lines.length - 1; i >= 0 && tail.length < 10; i--) if (String(lines[i]).trim()) tail.unshift(lines[i]);
    excerpt = tail.join('\n');
  }
  return { errorCount: errorCount, rawLine: rawLine, matchRule: matchRule, failAnchor: failAnchor,
           message: message, file: file, reportedLine: reportedLine, excerpt: excerpt };
}
async function runVb6Make(exe, vbp, logPath, defines, timeoutMs) {
  const t0 = Date.now();
  const argv = ['/make', vbp, '/out', logPath];
  if (defines && String(defines).trim()) argv.push('/d', String(defines).trim());
  return await new Promise(function (resolve) {
    let child = null, settled = false, timer = null;
    const finish = function (v) { if (settled) return; settled = true; if (timer) clearTimeout(timer); resolve(v); };
    try { child = spawnProgram(exe, argv, { windowsHide: true, stdio: 'ignore' }); }
    catch (e) { finish({ exitCode: null, timedOut: false, spawnError: String((e && e.message) || e), elapsedMs: Date.now() - t0, childPid: null, killMethod: '' }); return; }
    timer = setTimeout(function () {
      let killed = false;
      try { child.kill(); killed = true; } catch (e) { killed = false; }
      finish({ exitCode: null, timedOut: true, spawnError: '', elapsedMs: Date.now() - t0, childPid: child.pid || null, killMethod: killed ? 'child.kill()' : 'kill_failed' });
    }, Math.max(1000, timeoutMs));
    child.on('error', function (e) { finish({ exitCode: null, timedOut: false, spawnError: String((e && e.message) || e), elapsedMs: Date.now() - t0, childPid: child.pid || null, killMethod: '' }); });
    child.on('exit', function (code) { finish({ exitCode: (code === null || code === undefined) ? null : code, timedOut: false, spawnError: '', elapsedMs: Date.now() - t0, childPid: child.pid || null, killMethod: '' }); });
  });
}
function offlineProjectInfo(raw) {
  if (raw === undefined || raw === null || String(raw).trim() === '') {
    return { ok: false, reason: 'project_required_for_offline_build', project: '',
             message: 'no live IDE is online, so there is no ACTIVE project to infer: the offline build needs the .vbp path',
             howToProceed: ['pass build{project:"<absolute path to the .vbp>"} (or make_project{route:"cli",project:"<...>"})',
                            'or start an IDE with ide_start{project:"<absolute .vbp>"} and retry; the add-in path needs no project argument'] };
  }
  let p = String(raw).trim();
  if (!path.isAbsolute(p)) p = path.resolve(process.cwd(), p);
  if (!/\.vbp$/i.test(p)) return { ok: false, reason: 'project_not_a_vbp', project: p,
    message: 'the offline build compiles a VB6 project FILE; the given path is not a .vbp',
    howToProceed: ['pass the absolute path of the .vbp (not a folder, not a component file)'] };
  if (!fs.existsSync(p)) return { ok: false, reason: 'project_not_found', project: p,
    message: 'the .vbp does not exist on THIS machine, so no offline build can run',
    howToProceed: ['check the path (the offline route builds on the machine that runs the hub)',
                   'a project on another machine needs that machine to have an IDE, or its own hub'] };
  return { ok: true, vbp: p };
}
function offlineArtifactPath(vbpPath, keys) {
  // The .vbp value is stored QUOTED (ExeName32="X.EXE") and the add-in's ArtifactPathOf strips
  // the quotes; doing the same here is what makes the artifact path the real output file.
  let nm = keys && keys.ExeName32 !== undefined ? String(keys.ExeName32) : '';
  nm = nm.replace(/[\r\n\t"]/g, '').trim();
  if (!nm) return '';
  if (nm.indexOf(':') > 0) return nm;
  return path.join(path.dirname(vbpPath), nm);
}

// ---- (5) the offline receipts ------------------------------------------------------------
async function offlineBuildReceipt(tool, args, snap, ctx) {
  const a = args || {};
  const base = {
    ok: false, tool: tool, offline: true, offlineRoute: 'cli',
    offlineRouteRule: OFFLINE_ROUTING_RULE, route: 'cli', targetResolved: false,
    selectedBy: 'offline_route_no_live_ide',
    targetDiagnostic: (ctx && ctx.targetDiagnostic) || null
  };
  const pi = offlineProjectInfo(a.project);
  if (!pi.ok) return textResult(Object.assign(base, { reason: pi.reason, project: pi.project, message: pi.message, howToProceed: pi.howToProceed,
    failureKind: 'toolError', failureKindRule: FAILURE_KIND_RULE, suggestedAction: 'fix_params', retrySafe: false }), snap.contractVersion);
  const vbp = pi.vbp;
  let parsed = null;
  try { parsed = vbpParse(vbp); } catch (e) {
    return textResult(Object.assign(base, { reason: 'project_unreadable', project: vbp, message: 'the .vbp could not be read: ' + String((e && e.message) || e),
      failureKind: 'toolError', failureKindRule: FAILURE_KIND_RULE, suggestedAction: 'fix_params', retrySafe: false }), snap.contractVersion);
  }
  const rExe = resolveVb6Exe();
  if (!rExe.exe) return textResult(Object.assign(base, { reason: 'cli_unavailable', project: vbp,
    vb6Exe: '', vb6ExeSource: '', vb6ExeTried: rExe.tried,
    message: 'no VB6.EXE is resolvable on this machine, so the offline command-line build cannot run. This is NOT a missing IDE target.',
    howToProceed: ['set VB6MCP_VB6EXE to the absolute VB6.EXE path, or install VB6 on the machine that runs the hub',
                   'then retry build{project:"' + vbp + '"}'],
    failureKind: 'noIde', failureKindRule: FAILURE_KIND_RULE, suggestedAction: 'fix_params', retrySafe: false }), snap.contractVersion);
  const artifactPath = offlineArtifactPath(vbp, parsed.keys);
  const startTime = Date.now();
  const pre = artifactPath ? artifactView(artifactPath) : null;
  const timeoutRaw = a.timeout_ms === undefined || a.timeout_ms === null ? 0 : Number(a.timeout_ms);
  let timeoutMs = timeoutRaw > 0 ? timeoutRaw : (tool === 'make_project' ? 120000 : OFFLINE_BUILD_DEFAULT_TIMEOUT_MS);
  if (timeoutMs < 1000) timeoutMs = 1000;
  if (timeoutMs > OFFLINE_BUILD_MAX_TIMEOUT_MS) timeoutMs = OFFLINE_BUILD_MAX_TIMEOUT_MS;
  const logPath = (a.out_path && String(a.out_path).trim()) ? path.resolve(String(a.out_path).trim())
    : path.join(os.tmpdir(), 'vb6mcp-offline-build-' + process.pid + '-' + startTime + '.log');
  try { if (fs.existsSync(logPath)) fs.unlinkSync(logPath); } catch (e) { }
  const run = await runVb6Make(rExe.exe, vbp, logPath, a.defines, timeoutMs);
  const logSt = fileStatSafe(logPath);
  let logText = '', logEncoding = 'absent';
  if (logSt.exists && logSt.bytes > 0) {
    try { const lt = readTextFileAny(logPath); logText = lt.text; logEncoding = lt.encoding; } catch (e) { logText = ''; logEncoding = 'unreadable'; }
  }
  const cls = parseCompilerLog(logText);
  const art = artifactPath ? artifactView(artifactPath) : { path: '', exists: false, bytes: 0, mtime: null, mtimeMs: 0, sha256: '', sampledAt: Date.now() };
  const artifactChanged = artifactChangedIn(pre, art);
  // post-processing: LAA AFTER the compile, THEN the final hash (never a pre-rewrite hash published as `artifact`)
  const laaCfg = offlineResolveLaa(a, parsed.sections);
  const rewriteSteps = [{ step: 'compile', atMs: run.elapsedMs, route: 'cli', exitCode: run.exitCode, timedOut: run.timedOut }];
  let lastRewriteAt = Date.now();
  let laa = null;
  let finalArt = art;
  if (laaCfg.error) {
    laa = { requested: laaCfg.want, applicable: false, laaSet: false, error: laaCfg.error, laaSource: laaCfg.source,
            note: 'the LAA configuration could not be resolved, so NO PE header was touched. This is NOT a silent fallback to the default.' };
    rewriteSteps.push({ step: 'laa', skipped: true, reason: 'configuration_error', source: laaCfg.source });
  } else if (laaCfg.want && artifactPath && art.exists) {
    laa = laaPatch(artifactPath, true);
    laa.laaSource = laaCfg.source;
    laa.laaApplicable = true;
    lastRewriteAt = Date.now();
    rewriteSteps.push({ step: 'laa', atMs: lastRewriteAt - startTime, laaBefore: laa.laaBefore, laaAfter: laa.laaAfter, laaSet: laa.laaSet, bytesPatched: laa.bytesPatched, error: laa.error || '' });
    finalArt = artifactView(artifactPath);
  } else {
    laa = { requested: laaCfg.want, applicable: false, laaSet: false, error: '', laaSource: laaCfg.source, laaApplicable: false,
            reason: laaCfg.want ? 'artifact_missing' : 'not_requested',
            note: laaCfg.want ? 'LAA was REQUESTED but there is no artifact to patch, so no flag was written' : 'LAA was not requested' };
    rewriteSteps.push({ step: 'laa', skipped: true, reason: laa.reason });
  }
  const hashSampledAt = finalArt.sampledAt || Date.now();
  const hashedAfterRewrites = hashSampledAt >= lastRewriteAt;
  rewriteSteps.push({ step: 'hash', of: 'artifact', atMs: hashSampledAt - startTime,
                      samplePoint: 'artifact_after_rewrites', afterRewrites: hashedAfterRewrites });
  const readbackSha = (artifactPath && finalArt.exists) ? sha256File(artifactPath) : '';
  const artifactShaReadback = { sha256: readbackSha, match: !!readbackSha && readbackSha === finalArt.sha256, atMs: Date.now() - startTime };
  const laaApplicable = !!(laaCfg.want && artifactPath && finalArt.exists);
  const laaOk = !laaCfg.want ? null : !!(laa && !laa.error && laaApplicable && (laa.laaSet || laa.already));
  const compiled = !!art.exists && art.bytes > 0 && !run.timedOut;
  const exitOk = (run.exitCode === 0);
  const verdict = run.timedOut ? 'timeout'
    : (!exitOk || cls.errorCount > 0 ? 'failed'
      : (!art.exists ? 'no_artifact' : (artifactChanged ? 'artifact' : 'no_artifact_change')));
  const ok = !run.timedOut && !run.spawnError && exitOk && cls.errorCount === 0 &&
             !!art.exists && artifactChanged && artifactShaReadback.match;
  const body = Object.assign(base, {
    ok: ok,
    verdict: verdict,
    artifactChanged: artifactChanged,
    compiled: compiled,
    vb6Exe: rExe.exe, vb6ExeSource: rExe.source, vb6ExeTried: rExe.tried,
    vbp: vbp, project: vbp,
    defines: (a.defines === undefined || a.defines === null) ? '' : String(a.defines),
    logPath: logPath, logBytes: logSt.bytes, logEncoding: logEncoding, logAvailable: !!logSt.exists,
    log: logText.length <= 65536 ? logText : logText.slice(0, 65536),
    logTruncated: logText.length > 65536,
    exitCode: run.exitCode, timedOut: run.timedOut, spawnError: run.spawnError || '', killMethod: run.killMethod || '',
    budgetMs: timeoutMs, compilerElapsedMs: run.elapsedMs, childPid: run.childPid,
    errorCount: cls.errorCount, okByKeyword: /(?:编译成功|compiled successfully|success)/i.test(logText),
    matchRule: cls.matchRule,
    firstError: (cls.errorCount > 0 && (cls.message || cls.rawLine))
      ? { file: cls.file || '', line: cls.reportedLine || 0, lineSpace: 'code-pane', text: cls.message || String(cls.rawLine).trim(), rawLine: cls.rawLine }
      : null,
    firstErrorUnavailable: (cls.errorCount > 0 && !cls.message && !cls.rawLine) ? true : false,
    preRewrite: pre ? { existed: pre.exists, bytes: pre.bytes, mtime: pre.mtime, sha256: pre.sha256 } : null,
    artifact: { path: finalArt.path, exists: finalArt.exists, bytes: finalArt.bytes, mtime: finalArt.mtime,
                sha256: finalArt.sha256, changed: artifactChanged, builtNow: artifactChanged, sampledAt: finalArt.sampledAt },
    artifactShaReadback: artifactShaReadback,
    hashedAfterRewrites: hashedAfterRewrites, hashSampledAt: hashSampledAt, rewriteSteps: rewriteSteps,
    laaRequested: !!laaCfg.want, laaSource: laaCfg.source, laaApplicable: laaApplicable,
    laaBefore: !!(laa && laa.laaBefore), laaAfter: !!(laa && laa.laaAfter), laaSet: !!(laa && laa.laaSet),
    laaAlready: !!(laa && laa.already), laaBytesPatched: (laa && laa.bytesPatched) || 0,
    laaError: (laa && laa.error) || '', laaNote: (laa && laa.note) || '', laaOk: laaOk,
    laaFlag: 'IMAGE_FILE_LARGE_ADDRESS_AWARE (0x0020)',
    laaMethod: 'the hub rewrites the PE header itself; an external EditBin.exe is NOT required',
    laa: laa,
    note: 'the build ran OFFLINE through an independent VB6.EXE /make subprocess (no add-in and no IDE target were used)'
  });
  if (tool === 'make_project' && a.output_path && String(a.output_path).trim()) {
    body.outputPathIgnored = true;
    body.outputPathNote = 'the offline route compiles the .vbp as it is on disk; output_path (a VBProject.BuildFileName setting) needs a live IDE and was NOT applied';
  }
  if (!ok) {
    body.suggestedAction = run.timedOut ? 'retry_once' : 'fix_params';
    body.retrySafe = !!run.timedOut;
    if (run.spawnError) {
      body.reason = 'cli_spawn_failed';
      body.message = 'VB6.EXE could not be started: ' + run.spawnError;
      body.howToProceed = ['check VB6MCP_VB6EXE / the VB6 install', 'retry build{project:"' + vbp + '"}'];
    } else if (run.timedOut) {
      body.reason = 'build_timeout';
      body.message = 'the offline /make subprocess did not finish within ' + timeoutMs + ' ms and was terminated; the artifact is NOT claimed as built';
      body.howToProceed = ['raise timeout_ms (max ' + OFFLINE_BUILD_MAX_TIMEOUT_MS + ') and retry'];
    } else if (!exitOk) {
      body.reason = 'build_failed';
      body.message = 'VB6.EXE exited with code ' + run.exitCode + ' - the compile is NOT claimed as a success (logBytes=' + logSt.bytes + ')';
      body.exitFailure = { exitCode: run.exitCode, logBytes: logSt.bytes, logEncoding: logEncoding };
      body.howToProceed = ['the compiler FAILED (non-zero exit): read logPath / exitFailure and fix the project',
                           'an exit code with no log usually means VB6.EXE could not start the compile (project locked/open elsewhere, missing reference or a broken install)'];
    } else if (cls.errorCount > 0) {
      body.reason = 'build_failed';
      body.message = 'the compiler reported ' + cls.errorCount + ' error(s); see firstError / logPath';
      body.howToProceed = ['read firstError{file,line,lineSpace,text} and logPath', 'fix the project and retry build'];
    } else if (!art.exists) {
      body.reason = 'build_failed';
      body.message = 'the compiler exited 0 but produced NO artifact (artifact path: ' + (artifactPath || 'unknown') + ')';
      body.howToProceed = ['check ExeName32 in the .vbp and that the output path is writable', 'see logPath for the compiler output'];
    } else if (!artifactChanged) {
      body.reason = 'artifact_not_rebuilt';
      body.message = 'the compiler exited 0 but the artifact was NOT rewritten (same bytes/mtime): this is NOT claimed as a build';
      body.howToProceed = ['treat this as "nothing was built" (artifactChanged:false), not as a success', 'see logPath; an unchanged-output situation must be resolved by the caller'];
    } else {
      body.reason = 'build_failed';
      body.message = 'the artifact digest could not be read back consistently after the build';
      body.howToProceed = ['retry the build', 'if it persists, check for an external writer/AV touching the artifact'];
    }
    body.failureKind = 'toolError';
    body.failureKindRule = FAILURE_KIND_RULE;
  } else {
    body.reason = '';
  }
  return textResult(body, snap.contractVersion);
}

// ---- (6) project_meta offline (byte-preserving text edit) --------------------------------
// The SAME text-level edit the add-in performs (root scalar keys of a .vbp, a .bak next to it,
// read-back proof), implemented in the hub.  The file is ANSI/GBK: an ASCII-only edit is written
// as ASCII bytes, so no re-encoding can corrupt it; a non-ASCII VALUE cannot be encoded safely
// here and is REFUSED instead of being mangled.
async function offlineProjectMetaReceipt(tool, args, snap, ctx) {
  const a = args || {};
  const base = { ok: false, tool: 'project_meta', offline: true, offlineRoute: 'file', offlineRouteRule: OFFLINE_ROUTING_RULE,
                 targetResolved: false, selectedBy: 'offline_route_no_live_ide',
                 targetDiagnostic: (ctx && ctx.targetDiagnostic) || null, mode: 'offline_text' };
  const action = String(a.action === undefined || a.action === null ? '' : a.action).toLowerCase();
  if (!a.path || String(a.path).trim() === '') {
    return textResult(Object.assign(base, { reason: 'path_required_without_ide', action: action,
      message: 'no live IDE is online, so there is no ACTIVE project: project_meta needs path=<the .vbp>',
      howToProceed: ['pass path:"<absolute .vbp>"', 'or start an IDE and omit path to use the open project'],
      failureKind: 'toolError', failureKindRule: FAILURE_KIND_RULE, suggestedAction: 'fix_params', retrySafe: false }), snap.contractVersion);
  }
  let vbp = String(a.path).trim();
  if (!path.isAbsolute(vbp)) vbp = path.resolve(process.cwd(), vbp);
  if (!fs.existsSync(vbp)) return textResult(Object.assign(base, { reason: 'file_not_found', action: action, path: vbp,
    failureKind: 'toolError', failureKindRule: FAILURE_KIND_RULE, suggestedAction: 'fix_params', retrySafe: false }), snap.contractVersion);
  const parsed = vbpParse(vbp);
  if (action === 'get') {
    return textResult(Object.assign(base, { ok: true, path: vbp, exists: parsed.bytes > 0, lineEnding: parsed.lineEnding === '\r\n' ? 'crlf' : 'lf',
      keys: parsed.keys, sha256: sha256File(vbp), bytes: parsed.bytes, encoding: parsed.encoding,
      note: 'read by the hub directly from the .vbp text (no IDE involved)' }), snap.contractVersion);
  }
  if (action !== 'set') return textResult(Object.assign(base, { reason: 'invalid_action', action: action, path: vbp,
    message: 'action must be get or set', failureKind: 'toolError', failureKindRule: FAILURE_KIND_RULE, suggestedAction: 'fix_params', retrySafe: false }), snap.contractVersion);
  const updates = [
    ['MajorVer', a.major], ['MinorVer', a.minor], ['RevisionVer', a.revision],
    ['CompanyName', a.company_name], ['FileDescription', a.file_description], ['ProductName', a.product_name],
    ['LegalCopyright', a.legal_copyright], ['LegalTrademark', a.legal_trademark],
    ['Comments', a.comments], ['Title', a.title], ['ExeName32', a.exe_name]
  ].filter(function (kv) { return kv[1] !== undefined && kv[1] !== null && String(kv[1]) !== ''; });
  if (Object.prototype.hasOwnProperty.call(a, 'auto_increment') && a.auto_increment !== undefined && a.auto_increment !== null) {
    updates.push(['AutoIncrementVer', a.auto_increment ? '1' : '0']);
  }
  if (Object.prototype.hasOwnProperty.call(a, 'cond_comp')) updates.push(['CondComp', a.cond_comp === null || a.cond_comp === undefined ? '' : String(a.cond_comp)]);
  if (!updates.length) return textResult(Object.assign(base, { reason: 'no_fields_supplied', path: vbp,
    message: 'action=set needs at least one metadata field', failureKind: 'toolError', failureKindRule: FAILURE_KIND_RULE, suggestedAction: 'fix_params', retrySafe: false }), snap.contractVersion);
  for (const kv of updates) {
    if (/[^\x20-\x7E]/.test(String(kv[1]))) {
      return textResult(Object.assign(base, { reason: 'non_ascii_value_unsupported_offline', path: vbp, field: kv[0],
        message: 'the offline hub route writes the .vbp as ANSI bytes and cannot ENCODE a non-ASCII value; nothing was written',
        howToProceed: ['use an ASCII value', 'or edit the .vbp with a live IDE reachable (project_meta on the add-in face)'],
        failureKind: 'toolError', failureKindRule: FAILURE_KIND_RULE, suggestedAction: 'fix_params', retrySafe: false }), snap.contractVersion);
    }
  }
  const lines = parsed.text.split(/\r\n|\n|\r/);
  const done = {};
  const changed = [], added = [];
  for (const kv of updates) {
    for (let i = 0; i < lines.length; i++) {
      const sk = vbpScalarKeyOf(lines[i]);
      if (sk && sk.key.toLowerCase() === kv[0].toLowerCase()) { lines[i] = kv[0] + '=' + kv[1]; done[kv[0].toLowerCase()] = true; changed.push(kv[0]); break; }
    }
  }
  for (const kv of updates) {
    if (done[kv[0].toLowerCase()]) continue;
    let at = lines.length;
    for (let i = 0; i < lines.length; i++) {
      const t = String(lines[i]).trim();
      if (!t) continue;
      if (t.charAt(0) === '[' || vbpIsComponentLine(t)) { at = i; break; }
    }
    lines.splice(at, 0, kv[0] + '=' + kv[1]);
    added.push(kv[0]);
  }
  const newText = lines.join(parsed.lineEnding);
  const outBytes = encodeTextBytes(newText, parsed.encoding);
  if (!outBytes) return textResult(Object.assign(base, { reason: 'encoding_unsupported_offline', path: vbp,
    message: 'the file decodes as ' + parsed.encoding + ' and the edited text is not ASCII-only, so it cannot be written back without re-encoding; nothing was written',
    failureKind: 'toolError', failureKindRule: FAILURE_KIND_RULE, suggestedAction: 'fix_params', retrySafe: false }), snap.contractVersion);
  try { fs.copyFileSync(vbp, vbp + '.bak'); } catch (e) { }
  fs.writeFileSync(vbp, outBytes);
  const back = fs.readFileSync(vbp);
  const backText = decodeTextBytes(back).text;
  if (backText !== newText) {
    return textResult(Object.assign(base, { ok: false, reason: 'write_not_visible', path: vbp, actualBytes: back.length,
      detail: 'the write did NOT become visible on disk as the intended text; the success shape is deliberately NOT returned',
      failureKind: 'toolError', failureKindRule: FAILURE_KIND_RULE, suggestedAction: 'retry_once', retrySafe: true }), snap.contractVersion);
  }
  return textResult(Object.assign(base, { ok: true, path: vbp, changed: changed.join(','), added: added.join(','),
    changedFields: changed, addedFields: added, sha256: sha256File(vbp), bytesWritten: back.length, verifiedOnDisk: true,
    encoding: parsed.encoding, backup: vbp + '.bak' }), snap.contractVersion);
}
// ---- (7) no-live / dead-target receipts --------------------------------------------------
function offlineHintList() {
  return ['build / make_project(route:"cli") / project_meta run OFFLINE (no IDE needed): pass project / path explicitly',
          'run ide_start{project:"<absolute .vbp>"} to start an IDE (ready:true/false is reported)',
          'or set ide_target{target:"auto"} to clear a stale default'];
}
function targetDiagnosticBlock(opts) {
  const o = opts || {};
  return { rule: TARGET_DIAGNOSIS_RULE, reason: o.reason || null, onlinePids: o.onlinePids || [],
           requested: o.requested === undefined ? null : o.requested, explicit: !!o.explicit,
           deadTargetPid: (o.deadTargetPid === undefined ? null : o.deadTargetPid),
           // the list matches what is ACTUALLY served offline (package_app / verify always answer
           // offlineServed:false, so advertising them as offline-capable was wrong)
           offlineServedTools: ['build', 'make_project', 'project_meta'],
           offlineUnavailableTools: ['package_app', 'verify'],
           suggestedAction: o.suggestedAction || ((o.onlinePids && o.onlinePids.length) ? 'fix_params' : 'ide_start') };
}
function noLiveIdeBody(name, localD, explicit, rawRequest) {
  const pids = onlinePidsOf(localD && localD.instances);
  const dropped = (localD && localD.droppedDeadCandidates) || [];
  return {
    ok: false, tool: name, reason: 'no_live_ide', reasonLegacy: 'no_ide_running',
    message: 'no live IDE is addressable on this machine: 0 instances after pid/transport re-validation'
      + (dropped.length ? ' (' + dropped.length + ' stale candidate(s) were dropped: ' + dropped.map(function (d) { return d.pid + '/' + d.droppedReason; }).join(',') + ')' : '')
      + '. This is NOT a target mismatch - no target was resolved, because there is nothing to resolve it against.',
    onlinePids: pids, instances: [], candidates: [], staleDroppedCandidates: dropped,
    resolution: resolutionOf('no_candidate', [], { requested: rawRequest, explicit: explicit }),
    targetResolutionRule: TARGET_RESOLUTION_RULE,
    targetDiagnostic: targetDiagnosticBlock({ reason: 'no_live_ide', onlinePids: pids, requested: rawRequest, explicit: explicit }),
    suggestion: { action: 'ide_start', useTarget: pids.length ? String(pids[0]) : null,
                  ide_start: { project: '<absolute path to the .vbp>' },
                  note: 'start an IDE, then address it with _target=<pid>; the offline tools need no IDE at all' },
    howToProceed: offlineHintList(),
    failureKind: 'noIde', failureKindRule: FAILURE_KIND_RULE,
    suggestedAction: 'ide_start', retrySafe: false,
    rule: NO_LIVE_IDE_RULE
  };
}
// ---- (B7) 'THE TARGET YOU NAMED DIED' is NOT 'this machine has no IDE' ----------------------
// An explicit _target that matches a REGISTRATION whose process is gone is peer_dead: the caller
// must be able to tell 'restart THAT pid / pick another target' from 'start an IDE at all'.  It
// used to be folded into no_live_ide because the ladder only looked at LIVE instances.
// ---- (B7) 'THE TARGET YOU NAMED DIED' is NOT 'this machine has no IDE' ----------------------
// An explicit _target that matches a REGISTRATION whose process is gone is peer_dead: the caller
// must be able to tell 'restart THAT pid / pick another target' from 'start an IDE at all'.
// SOURCES, in order - the requirement is that the RAW registration files are the primary source,
// never a filtered candidate list:
//   (1) raw %TEMP%\vb6mcp\instances\*.json read from disk RIGHT NOW (source registry_file);
//   (2) the last discovery's staleRegistry (a pid_dead/port_not_listening entry still names a
//       registration the hub WITNESSED - source discovery_stale);
//   (3) the middle service's /health view in python mode, which reads the registry as ITS user and
//       reports pidAlive per instance (source middle_service_health) - the t16 topology where the
//       hub's own tmpdir may be a different user context.
async function deadRegistrationFor(selector, localRes) {
  const raw = String(selector === undefined || selector === null ? '' : selector).trim().replace(/^local:/i, '');
  if (!raw) return null;
  const low = raw.toLowerCase();
  const numeric = /^\d+$/.test(raw) ? Number(raw) : null;
  const matches = function (pid, port, project, projectPath, file) {
    if (numeric !== null && numeric === Number(pid)) return true;
    if (numeric !== null && port && numeric === Number(port)) return true;
    if (project && String(project).toLowerCase() === low) return true;
    if (projectPath && String(projectPath).toLowerCase().indexOf(low) >= 0) return true;
    if (file && String(file).toLowerCase().replace(/\.json$/i, '') === low) return true;
    return false;
  };
  // (1) RAW FILES
  let reg = null;
  try { reg = readRegistry(); } catch (e) { reg = null; }
  const regBlind = !(reg && !reg.unreadable);
  for (const e of ((reg && reg.entries) || [])) {
    if (!e || e.parseError || !e.pid) continue;
    if (pidAlive(e.pid)) continue;
    const f = e.file ? path.basename(e.file) : null;
    if (matches(e.pid, e.port, e.project, e.projectPath, f)) return { entry: e, file: f, source: 'registry_file', regBlind: regBlind };
  }
  // (2) the discovery's WITNESSED stale entries
  const stale = (localRes && localRes.localD && localRes.localD.staleRegistry) || [];
  for (const s of stale) {
    if (!s || !s.pid) continue;
    if (s.reason !== 'pid_dead' && s.reason !== 'port_not_listening') continue;
    if (pidAlive(Number(s.pid))) continue;
    const f = s.file ? path.basename(String(s.file)) : null;
    if (!matches(s.pid, s.port, null, null, f)) continue;
    return { entry: { pid: Number(s.pid), port: s.port || null, transport: null, project: null, projectPath: null },
             file: f, source: 'discovery_stale', staleReason: s.reason, regBlind: regBlind };
  }
  // (3) the middle service's view (python mode): its registry may be readable while ours is not
  if (BRIDGE_MODE !== 'inproc') {
    let bctx = bridgeIdentitySync();
    if (!bctx) { try { bctx = await bridgeIdentity(); } catch (e) { bctx = null; } }
    for (const bi of ((bctx && bctx.instances) || [])) {
      if (!bi || !bi.pid || bi.pidAlive !== false) continue;
      const f = bi.registryFile ? path.basename(String(bi.registryFile)) : null;
      if (!matches(bi.pid, bi.port, bi.project, bi.projectPath, f)) continue;
      return { entry: { pid: Number(bi.pid), port: Number(bi.port) || null, transport: bi.transport || null,
                        project: bi.project || null, projectPath: bi.projectPath || null },
               file: f, source: 'middle_service_health', regBlind: regBlind };
    }
  }
  return null;
}
async function peerDeadBody(name, rawTarget, localRes) {
  const hit = await deadRegistrationFor(rawTarget, localRes);
  if (!hit) return null;
  const e = hit.entry;
  const pids = onlinePidsOf(localRes && localRes.localD ? localRes.localD.instances : []);
  const howRecover = [
    'the target itself is GONE: start that instance again with ide_start{project:"' + (e.projectPath || '<absolute .vbp>') + '"}' +
      ' (ide_kill_and_relaunch is NOT applicable - it needs an ALIVE pid)',
    'or address a LIVE instance instead: ide_list, then _target="<live pid>"',
    'if the registration is residue, remove ' + (hit.file ? '%TEMP%\\vb6mcp\\instances\\' + hit.file : 'its entries\\<pid>.json') + ' and retry'
  ];
  return {
    ok: false, tool: name, reason: 'peer_dead', reasonLegacy: 'no_ide_running',
    failureKind: 'peer_dead', failureKindRule: FAILURE_KIND_RULE, peerDead: true,
    message: 'the target you named is DEAD, not absent: ' + (hit.file ? 'registration ' + hit.file : 'an instance registration') +
      ' still names pid ' + e.pid + (e.projectPath ? ' (' + e.projectPath + ')' : '') + ', but that pid is not alive' +
      (e.port ? ' and its port ' + e.port + ' is not listening' : '') + '. Nothing was forwarded and no other instance was silently substituted.',
    targetRequested: String(rawTarget === undefined || rawTarget === null ? '' : rawTarget),
    peerState: { pid: Number(e.pid), pidAlive: false, registrationPresent: true,
                 port: e.port || null, transport: e.transport || null,
                 project: e.project || null, projectPath: e.projectPath || null,
                 registryFile: hit.file || null, startedAt: e.startedAt || null,
                 source: hit.source, staleReason: hit.staleReason || null,
                 hubRegistryDir: instancesDir(), hubRegistryReadable: hit.regBlind === false },
    lostContext: { debugger: null,
                   debuggerBasis: 'the process is gone, so its IDE mode (design/run/break) cannot be read; null means UNKNOWN, never design',
                   projectPath: e.projectPath || null,
                   note: 'unsaved work of that process is gone with it; no probe or forward can recover the target' },
    onlinePids: pids,
    howToRecover: howRecover,
    howToProceed: howRecover,
    suggestion: { action: 'ide_start', useTarget: String(e.pid),
                  ide_start: { project: e.projectPath || '<absolute .vbp>' },
                  note: 'the named target is dead: restart THAT project (or pick a live pid from ide_list)' },
    targetDiagnostic: targetDiagnosticBlock({ reason: 'explicit_target_not_live', onlinePids: pids,
                                             requested: String(rawTarget || ''), explicit: true, deadTargetPid: Number(e.pid) }),
    targetResolutionRule: TARGET_RESOLUTION_RULE,
    suggestedAction: 'ide_start', retrySafe: false
  };
}
function deadExplicitTargetBody(name, localD, rawRequest) {
  const s = String(rawRequest === undefined || rawRequest === null ? '' : rawRequest);
  const pids = onlinePidsOf(localD && localD.instances);
  const dropped = (localD && localD.droppedDeadCandidates) || [];
  const deadPid = /^\d+$/.test(s) ? Number(s) : null;
  const droppedHit = deadPid ? dropped.filter(function (d) { return Number(d.pid) === deadPid; })[0] : null;
  return {
    ok: false, tool: name, reason: 'target_mismatch',
    message: '_target="' + s + '" names no LIVE IDE on this machine (online: ' + (pids.length ? pids.join(',') : 'none') + ')'
      + (droppedHit ? '; the hub bookkeeping still named pid ' + deadPid + ' and live discovery CORRECTED it (' + droppedHit.droppedReason + ')' : '')
      + '. The selector was NOT rewritten and nothing was forwarded.',
    targetRequested: s, onlinePids: pids, instances: [], candidates: [],
    staleDroppedCandidates: dropped,
    targetDiagnostic: targetDiagnosticBlock({ reason: 'explicit_target_not_live', onlinePids: pids, requested: s, explicit: true, deadTargetPid: deadPid }),
    suggestion: { action: pids.length ? 'fix_params' : 'ide_start',
                  useTarget: pids.length ? String(pids[0]) : null,
                  note: pids.length ? 'retry with _target="' + pids[0] + '" (the live instance)' : 'there is no live IDE: run ide_start{project:"<absolute .vbp>"}, or use build / make_project(route:"cli") which need no IDE' },
    howToProceed: offlineHintList(),
    resolution: resolutionOf('refuse_target_mismatch', [], { requested: s, explicit: true }),
    failureKind: 'noIde', failureKindRule: FAILURE_KIND_RULE,
    suggestedAction: pids.length ? 'fix_params' : 'ide_start', retrySafe: false,
    rule: NO_LIVE_IDE_RULE
  };
}
async function serveOfflineTool(off, name, argsIn, snap, ctx) {
  // This function is only reachable when the STRICT trigger held: 0 live local instances and the
  // middle service agreed (or there was no addresser at all).  The diagnostic must therefore say
  // no_live_ide, and it must never claim an explicit non-live target (the caller passed none).
  const diag = targetDiagnosticBlock({
    reason: 'no_live_ide', onlinePids: [], requested: (ctx && ctx.rawRequest) || null, explicit: false
  });
  const c = Object.assign({}, ctx || {}, { targetDiagnostic: diag });
  const offlineLadder = (ctx && ctx.ladderFallback) ? ctx.ladderFallback : null;
  if (off.offlineServed && off.impl === 'cli_build') return await offlineBuildReceipt(name, argsIn, snap, c);
  if (off.offlineServed && off.impl === 'vbp_meta') return await offlineProjectMetaReceipt(name, argsIn, snap, c);
  const base = { ok: false, tool: name, offline: false, offlineRoute: off.route || null, offlineRouteRule: OFFLINE_ROUTING_RULE,
    offlineServed: false, targetResolved: false, targetDiagnostic: diag,
    reason: 'no_live_ide', reasonLegacy: 'no_ide_running',
    onlinePids: (ctx && ctx.onlinePids) || [],
    message: name + ' needs the add-in for its pane/IDE side, and no live IDE is online: nothing was run and no target was resolved.',
    howToProceed: offlineHintList(),
    failureKind: 'noIde', failureKindRule: FAILURE_KIND_RULE,
    suggestedAction: 'ide_start', retrySafe: false,
    rule: NO_LIVE_IDE_RULE
  };
  if (off.requestedRoute === 'ide' || off.reason === 'route_ide_requires_live_ide') {
    base.offlineUnavailableReason = 'route_ide_requires_live_ide';
    base.message = 'make_project{route:"ide"} asks for the in-IDE Make, but no live IDE is online; nothing was compiled.';
    base.howToProceed = ['use route:"cli" (or omit route) to build with an independent VB6.EXE /make subprocess - it needs no IDE',
                         'or run ide_start{project:"<absolute .vbp>"} first, then retry route:"ide"'];
    base.suggestedAction = 'fix_params';
  }
  if (offlineLadder) base.ladderLiveFallback = offlineLadder;
  return textResult(base, snap.contractVersion);
}

// ---- (8) readiness polling (no fixed sleep decides anything) ------------------------------
// ONE readiness probe, used by the REUSE branch and by the polling loop.  ready may only be true
// when the add-in ITSELF answered (direct rpc for a TCP instance, the SAME forwarded mcp_info for a
// pipe-only one).  Everything is measured; nothing is a constant.
//
// THREE MUTUALLY EXCLUSIVE ANSWER BASES.  "No probe channel" and "the probe failed" are different
// facts and must never share a value (t16 acceptance measured no_route + addInAnswers:false for a
// pid whose add-in answered 16 ms later):
//   no_route               - there was NO channel to ask at all (no registration / no port / no
//                            addressable pipe): addInAnswers is NULL ("not measured"), never false.
//   direct_rpc             - the instance's own TCP port answered (or refused) /rpc tools/call.
//   middle_service_forward - a pipe-only instance was asked through the middle service (the only
//                            addresser of a pipe); an ATTEMPTED-and-failed forward KEEPS this
//                            basis, because the channel existed.
function readinessChecks(v) {
  const a = v || {};
  return {
    registryVisible: a.registryVisible === true,
    transportUp: a.transportUp === true,
    addInAnswers: a.addInAnswers === undefined ? null : a.addInAnswers,
    addInAnswersBasis: a.addInAnswersBasis || 'no_route',
    projectMatches: a.projectMatches === undefined ? null : a.projectMatches,
    pidAlive: a.pidAlive === true
  };
}
// DIRECT, CACHED-NOWHERE read of the ONE registration file that matters.  The readiness loop may
// NOT ask "is the pid in the discovered set": discovery is produced by PROBING every candidate
// (with soft grace windows), so an instance this hub has just created can be absent from that set
// for the WHOLE budget while its own registration file is already on disk (t16: registryVisible
// false for 120077 ms on a pid whose add-in answered 16 ms later).  Every call re-reads the dir.
function readRegistrationForPid(pid) {
  const dir = instancesDir();
  const want = String(pid);
  const named = path.join(dir, want + '.json');
  const parse = function (fp) {
    try {
      const rec = JSON.parse(fs.readFileSync(fp, 'utf8'));
      if (!rec || Number(rec.pid) !== Number(pid)) return null;
      return { rec: rec, file: fp };
    } catch (e) { return null; }
  };
  let hit = null;
  try { if (fs.existsSync(named)) hit = parse(named); } catch (e) { hit = null; }
  if (!hit) {
    // The registration file is normally <pid>.json, but the bridge has published other names;
    // the pid is the identity, so any *.json naming it counts.
    try {
      for (const f of fs.readdirSync(dir)) {
        if (!/\.json$/i.test(f)) continue;
        const fp = path.join(dir, f);
        if (fp === named) continue;
        hit = parse(fp);
        if (hit) break;
      }
    } catch (e) { hit = null; }
  }
  const alive = pidAlive(pid);
  if (!hit) return { visible: false, pidAlive: alive, dir: dir, file: null, filePath: null,
                     port: 0, transport: null, pipe: null, project: null, projectPath: null,
                     startedByHub: false, tcpEnabled: null, startedAt: null };
  const rec = hit.rec || {};
  const port = parseInt(rec.port, 10) || 0;
  const transport = rec.transport ? String(rec.transport) : (port ? 'tcp' : (rec.pipe ? 'pipe' : null));
  return { visible: true, pidAlive: alive, dir: dir, file: path.basename(hit.file), filePath: hit.file,
           port: port, transport: transport, pipe: rec.pipe || null,
           project: rec.project || rec.projectName || null,
           projectPath: rec.projectPath || rec.path || null, startedByHub: !!rec.startedByHub,
           // (FIX-26 / F1) the transport FACTS a restart receipt must be able to quote back:
           // tcpEnabled is written by the add-in and is null when the record predates that field.
           tcpEnabled: (rec.tcpEnabled === undefined || rec.tcpEnabled === null) ? null : !!rec.tcpEnabled,
           startedAt: rec.startedAt || null };
}
async function probeAddinReadinessOnce(pid, inst) {
  const t0 = Date.now();
  const alive = pidAlive(pid);
  const checks = readinessChecks({ registryVisible: !!inst, addInAnswers: null, addInAnswersBasis: 'no_route', pidAlive: alive });
  let port = Number(inst && inst.port) || 0;
  let projectActual = '';
  const isPipeOnly = !!(inst && !port && (String(inst.transport || '') === 'pipe' || !!inst.pipe));
  if (inst) {
    if (isPipeOnly) checks.transportUp = true;
    else if (port) { try { checks.transportUp = await tcpListening(port, 300); } catch (e) { checks.transportUp = false; } }
  }
  if (isPipeOnly && BRIDGE_MODE !== 'inproc') {
    // A pipe-only instance has NO port the hub can probe.  The evidence is the SAME forwarded
    // call the caller would make: mcp_info over the middle service (the only addresser of a pipe).
    // Without this a live pipe-only IDE could NEVER be reported ready:false forever.
    checks.addInAnswersBasis = 'middle_service_forward';
    try {
      const ans = await bridgeForwardCall('mcp_info', { _target: String(pid) }, snapshot(), { readinessProbe: true }, null);
      checks.addInAnswers = forwardAnsweredOk(ans);
      // A NEGATIVE readiness verdict must carry the fact it is based on (never a bare false).
      if (!checks.addInAnswers) checks.addInAnswersDiag = forwardFailureDiag(ans);
    } catch (e) { checks.addInAnswers = false; checks.addInAnswersError = String((e && e.message) || e); }
    if (checks.addInAnswers) {
      try {
        const st = await bridgeForwardCall('ide_state', { _target: String(pid) }, snapshot(), { readinessProbe: true }, null);
        // SAME wrapper rule as forwardAnsweredOk: a forwarded answer is STAMPED ({content,...}), not a raw RPC envelope.
        const env = forwardEnvelope(st);
        const obj = env ? firstBodyObject(env) : null;
        if (obj) projectActual = String(obj.project || '');
      } catch (e) { projectActual = ''; }
    }
  } else if (checks.transportUp && port) {
    checks.addInAnswersBasis = 'direct_rpc';
    try {
      const mi = await rpc(port, 'tools/call', { name: 'mcp_info', arguments: {} }, null, 8000);
      checks.addInAnswers = !!(mi && mi.result && !mi.result.isError);
    } catch (e) { checks.addInAnswers = false; checks.addInAnswersError = String((e && e.message) || e); }
    if (checks.addInAnswers) {
      try {
        const st = await rpc(port, 'tools/call', { name: 'ide_state', arguments: {} }, null, 8000);
        if (st && st.result && st.result.content && st.result.content[0]) {
          try { projectActual = String(JSON.parse(st.result.content[0].text).project || ''); } catch (e2) { projectActual = ''; }
        }
      } catch (e) { projectActual = ''; }
    }
  }
  const ready = checks.addInAnswers === true;
  return { ready: ready, readyWaitMs: Date.now() - t0, readyChecks: checks,
           port: port || null, projectActual: projectActual || '', basis: checks.addInAnswersBasis, pidAlive: alive };
}
// ONE poll round for the pid THIS CALL created: the registration is re-read from DISK and the
// channel is probed from THAT record - discovery is never consulted (see readRegistrationForPid).
async function readinessRoundForPid(pid, hintPort) {
  const reg = readRegistrationForPid(pid);
  let inst = null;
  let portSource = 'none';
  if (reg.visible) {
    inst = { pid: Number(pid), port: reg.port, transport: reg.transport, pipe: reg.pipe,
             project: reg.project, projectPath: reg.projectPath, registryFile: reg.file };
    portSource = 'registration';
  } else if (hintPort > 0) {
    // (B1, t16) THE HUB PINNED THIS PORT before spawning: when the add-in listens there but did
    // NOT write its registration, this is the only honest way to reach it - and reaching it is
    // what lets the hub publish the minimal registration below.
    let up = false; try { up = await tcpListening(hintPort, 300); } catch (e) { up = false; }
    if (up) { inst = { pid: Number(pid), port: hintPort, transport: 'tcp', pipe: null,
                     project: null, projectPath: null, registryFile: null }; portSource = 'pinned_hint'; }
  }
  const pr = await probeAddinReadinessOnce(pid, inst);
  pr.registration = reg;
  pr.portSource = portSource;
  pr.readyChecks.portSource = portSource;
  pr.readyChecks.registryVisible = reg.visible;   // a HINT is not a registration
  pr.readyChecks.pidAlive = reg.pidAlive;
  return pr;
}
function ideStartReadyBudgetMs(args) {
  let v = 0;
  if (args && args.ready_wait_ms !== undefined && args.ready_wait_ms !== null) v = Number(args.ready_wait_ms);
  if (!(v > 0)) v = parseInt(process.env.VB6MCP_IDE_START_READY_MS || '15000', 10);
  if (!(v > 0)) v = 15000;
  if (v < 1000) v = 1000;
  if (v > 180000) v = 180000;
  return v;
}
const READY_POLL_INTERVAL_MS = 400;
async function waitForAddinReady(pid, opts) {
  const o = opts || {};
  const budget = o.budgetMs && o.budgetMs > 0 ? o.budgetMs : 15000;
  const t0 = Date.now();
  let rounds = 0;
  let last = null;
  let projectActual = '';
  invalidateDiscovery();   // a start/relaunch makes every cached face obsolete
  while (Date.now() - t0 < budget) {
    rounds += 1;
    // FORCED FRESH every round: fs re-read of <pid>.json + a transport test + mcp_info.  No
    // discovery cache, no soft/grace window, and above all no "is this pid in the discovered
    // set" question - the answer must be about THIS pid, not about a set built by probing others.
    last = await readinessRoundForPid(pid, Number(o.hintPort) || 0);
    if (last.projectActual) projectActual = last.projectActual;
    if (last.readyChecks.addInAnswers === true) {
      last.readyChecks.projectMatches = o.wantProjectName
        ? (projectActual ? (projectActual.toLowerCase() === String(o.wantProjectName).toLowerCase()) : null)
        : true;
    }
    if (last.ready === true) break;
    const remain = budget - (Date.now() - t0);
    if (remain <= 0) break;
    await new Promise(function (r) { setTimeout(r, Math.min(READY_POLL_INTERVAL_MS, remain)); });
  }
  const c = (last && last.readyChecks) || readinessChecks({ addInAnswers: null, addInAnswersBasis: 'no_route' });
  const elapsed = Date.now() - t0;
  const ready = !!(((c.registryVisible === true) || c.portSource === 'pinned_hint') && c.transportUp === true && c.addInAnswers === true && c.pidAlive === true);
  const body = {
    ready: ready, readyWaitMs: elapsed, readyBudgetMs: budget, readyRule: IDE_START_READY_RULE,
    readyProbeRule: READY_PROBE_RULE, readyProbeRounds: rounds,
    readyChecks: Object.assign({ registryVisible: c.registryVisible === true, transportUp: c.transportUp === true,
                   addInAnswers: c.addInAnswers === undefined ? null : c.addInAnswers,
                   projectMatches: c.projectMatches === undefined ? null : c.projectMatches,
                   addInAnswersBasis: c.addInAnswersBasis || 'no_route',
                   pidAlive: c.pidAlive === true },
                   // a NEGATIVE verdict must carry the FACT it rests on (last round's forward outcome)
                   c.addInAnswersDiag ? { addInAnswersDiag: c.addInAnswersDiag } : {}),
    port: (last && last.port) || null,
    readyRegistrationFile: (last && last.registration && last.registration.file) || null,
    readyPortSource: (last && last.portSource) || 'none',
    projectActual: projectActual || ''
  };
  if (c.addInAnswersError) body.addInAnswersError = c.addInAnswersError;
  if (ready) invalidateDiscovery();   // the next discovery must SEE the pid that just answered
  if (!ready) {
    const regVisible = c.registryVisible === true;
    body.pollingHint = 'polled the STARTED pid DIRECTLY ' + rounds + ' time(s) over ' + elapsed +
      ' ms (its own instances\\<pid>.json registration, its transport, then an mcp_info call): ' +
      (regVisible ? 'the registration exists but the add-in has not answered yet'
                  : 'no registration file for this pid has appeared yet') +
      ' (pid alive: ' + (c.pidAlive ? 'yes' : 'NO') + '). Keep polling ide_list for pid ' + pid +
      ' (or ide_start the same project again - it reuses a live instance); the process WAS created';
    body.howToProceed = ['call ide_list and look for pid ' + pid,
                         'if it never appears, the IDE may be showing a modal dialog or the add-in failed to load: check the IDE window / read_active_dialogs',
                         'a later call can address it with _target="' + pid + '" once ide_list lists it'];
  }
  return body;
}
// ---- (8b) PROCESS IDENTITY before a kill (adversarial review B5) --------------------------
// A registration saying startedByHub=true can SURVIVE its process: the pid is then recycled by an
// unrelated process and "owned" becomes a lie.  Killing an IDE this hub did not start is a
// disaster, so the registry-only ownership path is verified against the LIVE PROCESS before the
// kill: (a) the caller/pick project must agree, (b) the process command line must name that
// project, (c) the process creation time must not be later than the registration.  Anything that
// cannot be verified is REFUSED (fail closed), never guessed.
function inspectProcessIdentity(pid) {
  const n = Number(pid);
  if (!n) return { error: 'bad_pid' };
  const script = "$ErrorActionPreference='Stop'; try { $p = Get-CimInstance Win32_Process -Filter 'ProcessId=" + n + "'; if ($p -ne $null) { [pscustomobject]@{ pid=$p.ProcessId; commandLine=$p.CommandLine; creationDate=$p.CreationDate.ToString('o'); exe=$p.ExecutablePath } | ConvertTo-Json -Compress } else { '{}' } } catch { '{\"error\":\"cim_query_failed\"}' }";
  try {
    const out = execFileSync('powershell.exe', ['-NoProfile', '-NonInteractive', '-Command', script], { encoding: 'utf8', timeout: 15000, windowsHide: true });
    const txt = String(out || '').trim();
    if (!txt) return { error: 'empty_output' };
    return JSON.parse(txt);
  } catch (e) { return { error: 'powershell_failed', detail: String((e && e.message) || e) }; }
}
function verifyKillIdentity(pid, target, projectPath, own) {
  const ev = { pid: Number(pid), ownership: own && own.ownership, projectPath: projectPath || '', checks: {} };
  const wantKey = projectPath ? canonicalProjectPath(projectPath, false) : null;
  const targetKey = (target && target.projectPath) ? canonicalProjectPath(target.projectPath, false) : null;
  if (wantKey && targetKey && wantKey.key && targetKey.key && wantKey.key !== targetKey.key) {
    ev.checks.project = 'mismatch';
    return { ok: false, reason: 'project_mismatch', basis: 'canonical_project_path',
      message: 'project="' + projectPath + '" does not match the project the live registration belongs to ("' + target.projectPath + '"): refusing to stop it',
      howToProceed: ['pass no project= (use the registration path) or the EXACT .vbp the instance owns', 'if this really is a stale registration, remove it manually and start the IDE with ide_start'],
      evidence: ev };
  }
  ev.checks.project = (wantKey && targetKey) ? 'match' : 'not_compared';
  if (own && own.ownership === 'hub_started') { ev.checks.process = 'in_memory_hub_started'; return { ok: true, basis: 'in_memory_hub_started', evidence: ev }; }
  const proc = inspectProcessIdentity(pid);
  ev.process = proc;
  if (proc.error || !proc.pid) {
    ev.checks.process = 'unverifiable';
    return { ok: false, reason: 'pid_identity_unverifiable', basis: 'process_inspection_failed',
      message: 'pid ' + pid + ' is only owned by a REGISTRATION (startedByHub=true), and the live process identity could not be read (' + (proc.error || 'no process') + '): refusing to stop it. A registration can outlive its process, so this could be a RECYCLED pid.',
      howToProceed: ['verify the process yourself (project + command line) and stop it manually', 'then start a fresh IDE with ide_start{project:"' + String(projectPath || '') + '"}'],
      evidence: ev };
  }
  const cl = String(proc.commandLine || '').toLowerCase().replace(/\//g, '\\');
  const projLower = String(projectPath || (target && target.projectPath) || '').toLowerCase().replace(/\//g, '\\');
  ev.checks.commandLineContainsProject = !!(projLower && cl.indexOf(projLower) >= 0);
  if (!ev.checks.commandLineContainsProject) {
    ev.checks.process = 'command_line_mismatch';
    return { ok: false, reason: 'pid_reused', basis: 'process_command_line_mismatch',
      message: 'the LIVE process behind pid ' + pid + ' does not name the project "' + String(projectPath || (target && target.projectPath) || '?') + '" on its command line: the registration is stale and the pid was likely RECYCLED. Nothing was stopped.',
      howToProceed: ['do not kill it via this tool: confirm with Get-CimInstance Win32_Process -Filter \'ProcessId=' + pid + '\' which project it holds', 'remove the stale registration for that pid, then start the IDE with ide_start'],
      evidence: ev };
  }
  const regStart = Date.parse(String((target && target.startedAt) || ''));
  const created = proc.creationDate ? Date.parse(String(proc.creationDate)) : NaN;
  ev.checks.registrationStartedAt = (target && target.startedAt) || null;
  ev.checks.processCreationDate = proc.creationDate || null;
  if (!isNaN(regStart) && !isNaN(created) && created > regStart + 300000) {
    ev.checks.process = 'created_after_registration';
    return { ok: false, reason: 'pid_reused', basis: 'process_created_after_registration',
      message: 'the process behind pid ' + pid + ' was created AFTER the registration was written (' + proc.creationDate + ' vs ' + target.startedAt + '): the pid was RECYCLED. Nothing was stopped.',
      howToProceed: ['remove the stale registration for pid ' + pid, 'start a fresh IDE with ide_start'],
      evidence: ev };
  }
  ev.checks.process = 'verified';
  return { ok: true, basis: 'process_command_line_matches_project', evidence: ev };
}
// ---- (9) stop / relaunch (hub-side, add-in independent) ----------------------------------
function stopProcessAndWait(pid, ms) {
  return new Promise(function (resolve) {
    let sent = false;
    try { process.kill(Number(pid)); sent = true; } catch (e) { sent = false; }
    const t0 = Date.now();
    const tick = function () {
      if (!pidAlive(pid)) return resolve({ killSent: sent, gone: true, waitedMs: Date.now() - t0 });
      if (Date.now() - t0 >= (ms || 12000)) return resolve({ killSent: sent, gone: false, waitedMs: Date.now() - t0 });
      setTimeout(tick, 250);
    };
    tick();
  });
}
// ---- (B1, t16) WHO WRITES %TEMP%\vb6mcp\instances\<pid>.json? -------------------------------
// The ADD-IN does, on listen success (Util.bas:428-429 "Written on listen success, deleted on
// normal exit", Util.bas:587 startedByHub, WinsockTransport.cls:253, ServerForm.frm:357).  The HUB
// never wrote one - which is why a spawned IDE whose add-in did not finish its listen (not loaded,
// listen failed, or it wrote into another user's %TEMP%) is INVISIBLE to ide_list even though the
// hub just created it AND the add-in answered on the hub's pinned port.  In that ONE proven case
// the hub now writes a MINIMAL registration itself (writer:"hub"), after the port was proven by an
// mcp_info answer.  It NEVER overwrites an existing entry and NEVER invents an instance that did
// not answer; when the add-in did not answer, nothing is written and the receipt says so.
function hubWriteRegistration(pid, port, projectPath, projectName) {
  const out = { written: false, by: null, path: null, existed: false, readbackVisible: false, reason: null,
                owner: 'add-in (writes its own entry on listen success); the hub only fills the gap' };
  if (!pid || !(Number(port) > 0)) { out.reason = 'no_pinned_port'; return out; }
  const dir = instancesDir();
  const file = path.join(dir, String(pid) + '.json');
  out.path = file;
  try {
    if (fs.existsSync(file)) {
      const prev = JSON.parse(fs.readFileSync(file, 'utf8'));
      if (Number(prev && prev.pid) === Number(pid)) {
        out.existed = true; out.by = (prev && prev.writer === 'hub') ? 'hub' : 'add-in';
        out.readbackVisible = true; out.reason = 'an entry for this pid already exists'; return out;
      }
    }
  } catch (e) { }
  const rec = { pid: Number(pid), port: Number(port), transport: 'tcp', tcpEnabled: true,
                project: projectName || null, projectPath: projectPath || null,
                startedAt: new Date().toISOString(), startedByHub: startedByHub.has(Number(pid)),
                version: null, dllSha256: null,
                writer: 'hub', source: 'hub_registration',
                writtenBasis: 'the hub created this process, pinned its port, and the ADD-IN answered mcp_info on that port; no add-in registration file existed' };
  try {
    if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
    fs.writeFileSync(file, JSON.stringify(rec, null, 2));
    const back = JSON.parse(fs.readFileSync(file, 'utf8'));
    out.written = true; out.by = 'hub';
    out.readbackVisible = Number(back.pid) === Number(pid) && Number(back.port) === Number(port);
    invalidateDiscovery();
  } catch (e) { out.reason = String((e && e.message) || e); }
  return out;
}
function clearOwnRegistryEntry(pid) {
  // The registration file is normally <pid>.json, but the bridge has published other file
  // names; clear EVERY registration whose pid matches (the pid is the identity, not the name).
  let cleared = false;
  try {
    const dir = instancesDir();
    const named = path.join(dir, String(pid) + '.json');
    if (fs.existsSync(named)) { fs.unlinkSync(named); cleared = true; }
    for (const f of fs.readdirSync(dir)) {
      if (!/\.json$/i.test(f)) continue;
      const fp = path.join(dir, f);
      if (fp === named) continue;
      try {
        const rec = JSON.parse(fs.readFileSync(fp, 'utf8'));
        if (Number(rec && rec.pid) === Number(pid)) { fs.unlinkSync(fp); cleared = true; }
      } catch (e) { }
    }
  } catch (e) { }
  return cleared;
}
async function pickFreePort() {
  for (let pp = 8767; pp <= 8796; pp++) {
    let busy = true;
    try { busy = await tcpListening(pp, 120); } catch (e) { busy = false; }
    if (!busy) return pp;
  }
  return 0;
}

// The two registry helpers below exist at MODULE scope on purpose: hubTool also declares a
// function with the first name INSIDE its own body (block-scoped), which a module-level caller
// cannot see - calling it from here threw a TypeError and the call never answered (measured).
function registryAliveCandidates(wantPid) {
  const dir = instancesDir();
  let files = [];
  try { files = fs.readdirSync(dir); } catch (e) { return []; }
  const res = [];
  for (const f of files) {
    if (!/\.json$/i.test(f)) continue;
    let rec = null;
    try { rec = JSON.parse(fs.readFileSync(path.join(dir, f), 'utf8')); } catch (e) { continue; }
    const pid = Number(rec && rec.pid);
    if (!pid) continue;
    if (wantPid && pid !== Number(wantPid)) continue;
    let alive = false;
    try { process.kill(pid, 0); alive = true; } catch (e) { alive = !!(e && e.code === 'EPERM'); }
    if (!alive) continue;
    res.push({ pid: pid, port: Number(rec.port) || 0, transport: rec.transport || null,
               project: rec.project || null, projectPath: rec.projectPath || null,
               startedByHub: !!rec.startedByHub, dllSha256: rec.dllSha256 || null,
               registryFile: f, registryPath: path.join(dir, f) });
  }
  return res;
}
function ownershipUnknownRefusal(out, blk, legacyReason) {
  out.refused_reason = 'ownership_unknown';
  out.legacyReason = legacyReason || 'not_owned';
  out.ownershipUnknown = true;
  out.howToDisambiguate = 'nothing was stopped. To settle WHOSE instance this is: (1) pass the exact pid AND project=<the absolute .vbp> so the target cannot be a recycled pid; (2) start the middle service (VB6MCP_BRIDGE_URL) if the registration is pipe-only and retry; (3) if the registration is a residue, remove %TEMP%\\vb6mcp\\instances\\<pid>.json and retry; (4) if the port ANSWERED without the registered pid being alive, that is foreign evidence (not unknown).';
  return out;
}
// ---- (10) ide_kill_and_relaunch: the add-in-independent reset primitive -------------------
// The contract case this exists for: the add-in pipe is BUSY (a long call holds the
// single-flight gate), so ide_state{release_inflight:true} cannot even be delivered - it uses the
// same pipe.  The only recovery left is a HUB-side primitive that never speaks to the add-in:
// stop the process, clear the registration, launch VB6.EXE again, poll until the add-in answers.
async function relaunchOwnedInstance(args, snap) {
  const pause = function (ms) { return new Promise(function (r) { setTimeout(r, ms); }); };
  const t0 = Date.now();
  const a = args || {};
  const out = {
    ok: false, tool: 'ide_kill_and_relaunch', relaunchRule: RELAUNCH_RULE,
    pid: a.pid ? Number(a.pid) : 0, project: a.project ? String(a.project) : '',
    stoppedPid: 0, newPid: 0, port: 0, ready: false, readyWaitMs: 0, readyChecks: null,
    threeWay: { http: false, mcpInfo: false, project: false }, three_way_ok: false,
    ownership: '', ownershipBasis: '', ownedByHub: false, portRuleApplies: false, portBasis: '',
    ownershipVerdict: null, ownershipVerdictRule: OWNERSHIP_VERDICT_RULE,
    stopConfirmed: false, registryCleared: false, pinPort: 0, projectVbp: '',
    vb6Exe: '', vb6ExeSource: '', vb6ExeTried: [], elapsedMs: 0,
    onlinePids: [], instances: [], steps: [], note: '', refused_reason: '',
    failureKind: '', failureKindRule: FAILURE_KIND_RULE, suggestedAction: '', retrySafe: false,
    // (35) the CWD of the launch: filled AT the spawn point; a path that refuses, or whose
    // spawn throws, never started a process and keeps this fail-closed null / "unavailable".
    cwd: null, cwdBasis: 'unavailable'
  };
  dbg('relaunch: enter pid=' + out.pid + ' project=' + out.project);
  if (!out.pid) {
    out.refused_reason = 'pid_required';
    out.message = 'ide_kill_and_relaunch needs pid=<the instance to stop> (it never guesses a target)';
    out.howToProceed = ['pass the pid from ide_list (only an instance THIS hub owns can be reset)'];
    out.suggestedAction = 'fix_params';
    out.elapsedMs = Date.now() - t0;
    return textResult(out, snap.contractVersion);
  }
  if (MULTI_BRIDGE) {
    const remoteHits = await remoteInstancesForPid(out.pid);
    if (remoteHits.length) {
      const rr = remoteControlRefusal('ide_restart', out.pid, remoteHits[0]);
      for (const k of Object.keys(rr)) out[k] = rr[k];
      out.tool = 'ide_kill_and_relaunch';
      out.elapsedMs = Date.now() - t0;
      return finishControlOut(out, snap);
    }
  }
  let dd = null;
  try { dd = await resolveDiscovery(await discoverSoft(LIST_GRACE_MS)); } catch (e) { dd = null; }
  dbg('relaunch: discovery done instances=' + ((dd && dd.instances) ? dd.instances.length : -1));
  const insts = ((dd && dd.instances) || []).filter(function (i) { return i.pid; });
  out.onlinePids = onlinePidsOf(insts);
  out.instances = insts.map(function (i) {
    const o = ownershipOf(i);
    return { pid: i.pid, port: i.port, project: i.project, ownership: o.ownership, ownedByHub: o.owned };
  });
  let target = insts.filter(function (i) { return Number(i.pid) === out.pid; })[0] || null;
  if (!target) {
  dbg('relaunch: before registryAliveCandidates dir=' + instancesDir());
    const alive = registryAliveCandidates(out.pid);
    dbg('relaunch: registryAliveCandidates -> ' + alive.length);
    if (alive.length) {
      const a0 = alive[0];
      target = { pid: a0.pid, port: a0.port || 0, transport: a0.transport, project: a0.project, projectPath: a0.projectPath,
                 startedByHub: a0.startedByHub, registryFile: a0.registryFile, discoveredBy: 'registry' };
      out.steps.push({ step: 'resolve', via: 'registry_alive_candidate', pid: a0.pid });
    }
  } else {
    out.steps.push({ step: 'resolve', via: 'discovery', pid: target.pid });
  }
  dbg('relaunch: target resolved=' + (target ? target.pid : 'none'));
  if (!target) {
    out.refused_reason = 'pid_not_found';
    out.message = 'pid ' + out.pid + ' is not a live instance and no registration with that pid has an ALIVE process; nothing was stopped';
    out.howToProceed = ['call ide_list and use one of its onlinePids', 'or run ide_start{project:"<absolute .vbp>"}'];
    out.suggestedAction = 'ide_start';
    out.elapsedMs = Date.now() - t0;
    return finishControlOut(out, snap);
  }
  const own = ownershipOf(target);
  out.ownership = own.ownership; out.ownershipBasis = own.ownershipBasis; out.ownedByHub = own.owned;
  out.portRuleApplies = own.portRuleApplies; out.portBasis = own.portBasis;
  out.port = target.port || 0;
  const ov = await resolveOwnership(target);
  out.ownershipVerdict = ov;
  if (ov.verdict === 'unknown') {
    ownershipUnknownRefusal(out, ov, 'instance_found_but_ownership_unknown');
    out.refused_reason = 'ownership_unknown';
    out.message = 'pid ' + out.pid + ' is alive, but WHOSE instance it is cannot be established: ' + ov.basis.join('; ') + '. Nothing was stopped.';
    out.suggestedAction = 'fix_params';
    out.elapsedMs = Date.now() - t0;
    return finishControlOut(out, snap);
  }
  if (!own.owned) {
    out.refused_reason = 'not_allowed';
    out.message = 'pid ' + out.pid + ' is not owned by this hub: ' + own.ownershipBasis + '; nothing was stopped (this tool only resets instances the hub started or that registered startedByHub=true)';
    out.howToProceed = ['only a hub-owned instance may be reset; start it with ide_start first', 'if it is YOUR IDE, stop it yourself and call ide_start'];
    out.suggestedAction = 'fix_params';
    out.elapsedMs = Date.now() - t0;
    return finishControlOut(out, snap);
  }
  if (own.foreignPort) {
    out.refused_reason = 'foreign_instance_refused';
    out.message = own.portBasis + ' - refusing to stop it (the reserved range is never signalled by this hub)';
    out.suggestedAction = 'fix_params';
    out.elapsedMs = Date.now() - t0;
    return finishControlOut(out, snap);
  }
  // Everything that could refuse must refuse BEFORE the process is stopped: project path, exe and
  // the port pin are all resolved first, so a failure here cannot leave the IDE down.
  out.projectVbp = String((a.project ? a.project : '') || target.projectPath || '');
  if (!out.projectVbp) {
    out.refused_reason = 'project_path_unknown';
    out.message = 'the target instance did not publish its .vbp path; pass project=<the absolute .vbp> so the IDE can be relaunched on it. Nothing was stopped.';
    out.howToProceed = ['retry with project:"<absolute path to the .vbp>"'];
    out.suggestedAction = 'fix_params';
    out.elapsedMs = Date.now() - t0;
    return finishControlOut(out, snap);
  }
  const rExe = resolveVb6Exe();
  out.vb6ExeTried = rExe.tried;
  if (!rExe.exe) {
    out.refused_reason = 'vb6_exe_not_found';
    out.message = 'no VB6.EXE was resolvable (set VB6MCP_VB6EXE, or install VB6); nothing was stopped. Tried: ' + rExe.tried.join(' | ');
    out.suggestedAction = 'fix_params';
    out.elapsedMs = Date.now() - t0;
    return finishControlOut(out, snap);
  }
  out.vb6Exe = rExe.exe; out.vb6ExeSource = rExe.source;
  out.pinPort = await pickFreePort();
  if (!out.pinPort) {
    out.refused_reason = 'no_free_port_in_8767_8796';
    out.message = 'no free port was found in 8767..8796, so NOTHING was stopped and the IDE is untouched (an instance on 8766 is never ours)';
    out.suggestedAction = 'fix_params';
    out.elapsedMs = Date.now() - t0;
    return finishControlOut(out, snap);
  }
  // (B5) VERIFY THE LIVE PROCESS before anything is stopped: project agreement + process identity.
  const idv = verifyKillIdentity(out.pid, target, out.projectVbp, own);
  out.identityBasis = idv.basis || '';
  out.identityEvidence = idv.evidence || null;
  out.steps.push({ step: 'identity_check', ok: !!idv.ok, basis: idv.basis || '', reason: idv.reason || '' });
  if (!idv.ok) {
    out.refused_reason = idv.reason;
    out.message = idv.message;
    out.howToProceed = idv.howToProceed;
    out.suggestedAction = 'fix_params';
    out.elapsedMs = Date.now() - t0;
    return finishControlOut(out, snap);
  }
  // best-effort save through the add-in FIRST (it may be busy: a failure here is reported, not fatal)
  let saveOk = false;
  if (out.port) {
    try { const sv = await rpc(out.port, 'tools/call', { name: 'save_project', arguments: { path: out.projectVbp } }, null, 5000); saveOk = !!(sv && sv.result && !sv.result.isError); }
    catch (e) { saveOk = false; }
  }
  out.steps.push({ step: 'save_project', attempted: !!out.port, ok: saveOk, note: out.port ? '' : 'no TCP port: the add-in is only reachable through the middle service, so no save was attempted' });
  const stopRes = await stopProcessAndWait(out.pid, 12000);
  out.stoppedPid = out.pid;
  out.stopConfirmed = !!stopRes.gone;
  out.steps.push({ step: 'kill', pid: out.pid, killSent: stopRes.killSent, gone: stopRes.gone, waitedMs: stopRes.waitedMs });
  if (!stopRes.gone) {
    out.refused_reason = 'pid_survived_kill';
    out.message = 'the process did not exit after the kill; NOTHING was relaunched (two IDEs on one .vbp must never happen). Stop it externally and retry.';
    out.howToProceed = ['run Stop-Process -Id ' + out.pid + ' -Force', 'then retry ide_kill_and_relaunch or ide_start'];
    out.suggestedAction = 'retry_once';
    out.retrySafe = true;
    out.elapsedMs = Date.now() - t0;
    return finishControlOut(out, snap);
  }
  const clearedNow = clearOwnRegistryEntry(out.pid);
  const stillThere = registryAliveCandidates(out.pid);
  out.registryCleared = !!(clearedNow && stillThere.length === 0);
  out.registryClearedReadBack = { deleted: !!clearedNow, remainingRegistrations: stillThere.length };
  startedByHub.delete(out.pid);
  invalidateDiscovery();
  out.steps.push({ step: 'registry_cleared', pid: out.pid, cleared: out.registryCleared, readBack: out.registryClearedReadBack.remainingRegistrations });
  const childEnv = Object.assign({}, process.env, { VB6MCP_PORT: String(out.pinPort) });
  // (35) captured AT THE SPAWN POINT: no cwd option is passed, so the child inherits this hub
  // process's cwd; if it cannot be read the receipt says "unavailable" (never guessed).
  const spawnOpts = { detached: true, stdio: 'ignore', env: childEnv };
  const spawnCwd = hubSpawnCwd(spawnOpts);
  const spawnCwdBasis = spawnCwd === null ? 'unavailable' : 'spawn_cwd';
  let child = null;
  try {
    child = spawnProgram(out.vb6Exe, [out.projectVbp], spawnOpts);
    child.unref();
    // (35) a process WAS created by this call: publish the spawn cwd.  A throw above leaves the
    // fail-closed null / "unavailable" from the receipt initializer untouched.
    out.cwd = spawnCwd; out.cwdBasis = spawnCwdBasis;
  } catch (e) {
    out.refused_reason = 'launch_failed';
    out.message = 'VB6.EXE could not be started: ' + String((e && e.message) || e);
    out.howToProceed = ['check VB6MCP_VB6EXE / the VB6 install', 'retry ide_start{project:"' + out.projectVbp + '"}'];
    out.suggestedAction = 'fix_params';
    out.elapsedMs = Date.now() - t0;
    return finishControlOut(out, snap);
  }
  out.newPid = child.pid || 0;
  startedByHub.set(out.newPid, { project: out.projectVbp, port: out.pinPort });
  out.steps.push({ step: 'launch', pid: out.newPid, project: out.projectVbp, pinnedPort: out.pinPort });
  let budget = 0;
  if (a.ready_wait_ms !== undefined && a.ready_wait_ms !== null) budget = Number(a.ready_wait_ms);
  if (!(budget > 0)) budget = parseInt(process.env.VB6MCP_IDE_RELAUNCH_READY_MS || '60000', 10);
  if (!(budget > 0)) budget = 60000;
  if (budget < 1000) budget = 1000;
  if (budget > 300000) budget = 300000;
  const wantName = String(target.project || '').trim();
  const rw = await waitForAddinReady(out.newPid, { budgetMs: budget, wantProjectName: wantName, hintPort: out.pinPort || 0 });
  out.registration = rw.ready
    ? hubWriteRegistration(out.newPid, (rw.port || out.pinPort || 0), out.projectVbp, String(out.project || ''))
    : { written: false, by: null, path: path.join(instancesDir(), String(out.newPid) + '.json'), existed: false, readbackVisible: false, reason: 'the relaunched add-in did not answer on any route, so nothing was registered' };
  out.ready = !!rw.ready;
  out.readyWaitMs = rw.readyWaitMs;
  out.readyChecks = rw.readyChecks;
  out.readyProbeRounds = rw.readyProbeRounds;
  out.readyProbeRule = rw.readyProbeRule;
  out.readyRegistrationFile = rw.readyRegistrationFile || null;
  out.port = rw.port || out.pinPort;
  // threeWay components are MEASURED, and each names its own basis: http only means HTTP when a
  // TCP port exists (a pipe-only instance reaches the add-in through the middle service, which is
  // NOT http); project is a tri-state result (true/false/null=unverified), never a derived guess.
  const viaTcp = !!(out.port && rw.readyChecks.transportUp);
  out.threeWay.transport = viaTcp ? 'tcp' : 'pipe';
  // The pinned TCP port is only a readiness channel for a TCP instance; a pipe-only instance
  // reaches the add-in through the middle service, so the TCP probe is NOT APPLICABLE (not a failure).
  out.threeWay.http = viaTcp ? true : 'not_applicable';
  out.threeWay.httpBasis = viaTcp
    ? 'tcp_listening(port)'
    : 'pipe_instance: the add-in registered NO TCP endpoint (port 0 by design); the pinned port is not a channel for this instance';
  // tri-state: null = NO probe channel existed (no_route), false = the probe FAILED -
  // the aggregate below requires a true, so neither can be mistaken for an answer.
  out.threeWay.mcpInfo = rw.readyChecks.addInAnswers === true ? true
                            : (rw.readyChecks.addInAnswers === false ? false : null);
  out.threeWay.project = (rw.readyChecks.projectMatches === true) ? true : (rw.readyChecks.projectMatches === false ? false : null);
  out.threeWay.basis = {
    transport: viaTcp ? 'tcp_listening(port)' : (rw.readyChecks.addInAnswersBasis === 'middle_service_forward' ? 'pid_alive_pipe_only(middle_service_forward)' : 'not_measured'),
    mcpInfo: rw.readyChecks.addInAnswersBasis || 'not_measured',
    project: (out.threeWay.project === null) ? 'unverified(no ide_state project available)' : 'ide_state.project'
  };
  // The aggregate follows the SAME channel rule: TCP instances must pass the http probe;
  // pipe instances must pass mcp_info + project (there is no http to pass).
  out.three_way_ok = (out.threeWay.transport === 'tcp'
      ? out.threeWay.http === true
      : out.threeWay.http === 'not_applicable')
    && out.threeWay.mcpInfo === true && out.threeWay.project === true;
  out.ok = out.ready;
  if (out.ok) {
    out.note = 'pid ' + out.stoppedPid + ' was stopped and relaunched as pid ' + out.newPid + ' on ' + out.projectVbp + '; the add-in answered within ' + out.readyWaitMs + ' ms (port ' + out.port + ')';
  } else {
    out.refused_reason = 'new_instance_not_ready';
    out.message = 'the old process WAS stopped and a new VB6.EXE WAS launched (pid ' + out.newPid + '), but the add-in did not answer within ' + budget + ' ms';
    out.pollingHint = rw.pollingHint || '';
    out.howToProceed = rw.howToProceed || ['poll ide_list for pid ' + out.newPid];
    out.suggestedAction = 'retry_once';
    out.retrySafe = true;
    out.note = 'the IDE is running but not ready yet: the reset DID happen, the readiness poll timed out';
  }
  out.elapsedMs = Date.now() - t0;
  return finishControlOut(out, snap);
}

// ============================================================================================
// (WP-22-02) LONG HUB OPERATIONS: a continuable opId + a BOUNDED wait
// --------------------------------------------------------------------------------------------
// A hub operation that can take minutes (ide_restart) used to be ONE blocking call: a client
// with a 30 s request timeout got a transport error and NO receipt, while the restart had
// already succeeded - and a blind retry would restart a SECOND time. Every such operation now
// registers an op record, publishes opId + started/elapsedMs/currentStep/lastKnownState on
// EVERY path, and can be CONTINUED with ide_restart{opId} without starting anything.
// ============================================================================================
const HUB_OP_TTL_MS = 30 * 60 * 1000;
const HUB_OPS = new Map();
let HUB_OP_SEQ = 0;
const HUB_OP_RULE = 'op = a continuable long hub operation. opId names it; state is in_progress | done | refused | error; currentStep is the phase the operation is IN (or stopped in); lastKnownState is the last state it PROVED (started | stopped | relaunched | addin_answering | ready | not_ready). ide_restart{opId} answers from the registry and starts NOTHING. An in_progress op must NOT be started again: a second restart could stop the process the first one just started.';
function pruneHubOps() {
  const cut = Date.now() - HUB_OP_TTL_MS;
  for (const kv of Array.from(HUB_OPS.entries())) { const r = kv[1]; if (r.finishedAt && r.finishedAt < cut) HUB_OPS.delete(kv[0]); }
}
function startHubOp(tool, detail) {
  pruneHubOps(); HUB_OP_SEQ += 1;
  const opId = tool + '-' + process.pid.toString(36) + '-' + Date.now().toString(36) + '-' + HUB_OP_SEQ;
  const rec = { opId: opId, tool: tool, started: true, state: 'in_progress', startedAt: Date.now(), finishedAt: 0,
                currentStep: 'start', lastKnownState: 'started', stepAt: 0, steps: [], out: null, detail: detail || null };
  HUB_OPS.set(opId, rec);
  return rec;
}
function hubOpStep(rec, step, lastKnownState) {
  if (!rec) return;
  rec.currentStep = step; rec.stepAt = Date.now();
  if (lastKnownState) rec.lastKnownState = lastKnownState;
}
function markHubOpStep(out, rec, step, lastKnownState) {
  hubOpStep(rec, step, lastKnownState);
  if (out && rec) { out.currentStep = step; out.lastKnownState = rec.lastKnownState; out.elapsedMs = Date.now() - rec.startedAt; }
}
function hubOpDone(rec, state) { if (rec) { rec.state = state; rec.finishedAt = Date.now(); } }
function hubOpById(opId) { return HUB_OPS.get(String(opId)) || null; }
function hubOpInFlight(tool) {
  for (const r of HUB_OPS.values()) { if (r.tool === tool && r.state === 'in_progress') return r; }
  return null;
}
function hubOpView(rec) {
  const out = rec.out || {};
  const v = { opId: rec.opId, op: rec.tool, state: rec.state, started: rec.started, finished: rec.state !== 'in_progress',
              elapsedMs: (rec.finishedAt || Date.now()) - rec.startedAt, currentStep: rec.currentStep,
              lastKnownState: rec.lastKnownState, steps: rec.steps, opRule: HUB_OP_RULE };
  // (FIX-26 / F3) A caller that was answered pending and then follows ide_restart{opId} used to
  // get ONLY the progress scalars: the terminal RESULT of the operation it was told to poll (what
  // actually happened, what the new instance is, what to do next) was unreachable from the
  // continuation.  The same fields the blocking answer carries are published here as well.
  for (const k of ['stoppedPid', 'newPid', 'port', 'pinPort', 'project', 'projectVbp', 'three_way_ok', 'waitMs', 'readyWaitMs',
                   'readyWaitMs', 'refused_reason', 'ok', 'restarted', 'stopConfirmed', 'saveAttempted', 'saveOk',
                   'registryCleared', 'lockCleared', 'ownership', 'ownershipBasis', 'ownedByHub', 'truth',
                   'note', 'howToProceed', 'newInstanceFacts', 'probeChecks', 'manualCommand', 'nextStep',
                   'vb6Exe', 'vb6ExeSource', 'instanceCount', 'checkedRegistryDir',
                   // (35) the launch CWD travels with the terminal result of the op
                   'cwd', 'cwdBasis']) {
    if (out[k] !== undefined) v[k] = out[k];
  }
  if (out.threeWay) v.threeWay = out.threeWay;
  return v;
}
function hubOpNotFound(tool, opId) {
  return { ok: false, tool: tool, mode: 'status', statusQuery: true, opId: String(opId), reason: 'op_not_found',
           message: 'no operation with opId "' + String(opId) + '" is known to this hub (ops are kept for ' + Math.round(HUB_OP_TTL_MS / 60000) + ' minutes and do NOT survive a hub restart)',
           failureKind: 'toolError', failureKindRule: FAILURE_KIND_RULE, suggestedAction: 'fix_params', retrySafe: false,
           howToProceed: ['call ide_list to see what is actually running now', 'only start a NEW restart when ide_list shows the IDE you want replaced'],
           opRule: HUB_OP_RULE };
}
const RESTART_DEFAULT_WAIT_MS = 25000;
const RESTART_MIN_WAIT_MS = 1500;
const RESTART_MAX_WAIT_MS = 300000;
function resolveRestartWaitMs(v) {
  let ms = null;
  if (v !== undefined && v !== null && String(v).trim() !== '') { const n = Number(v); if (isFinite(n)) ms = n; }
  if (ms === null) {
    const e = Number(process.env.VB6MCP_HUB_RESTART_WAIT_MS);
    ms = isFinite(e) && e > 0 ? e : RESTART_DEFAULT_WAIT_MS;
  }
  if (ms < RESTART_MIN_WAIT_MS) ms = RESTART_MIN_WAIT_MS;
  if (ms > RESTART_MAX_WAIT_MS) ms = RESTART_MAX_WAIT_MS;
  return Math.round(ms);
}
function restartInFlightReceipt(busy) {
  const v = hubOpView(busy);
  return Object.assign({ ok: false, tool: 'ide_restart', restarted: false, stoppedPid: 0, newPid: 0,
    mode: 'restart', reason: 'restart_in_progress', refused_reason: 'restart_in_progress', inProgress: v,
    retrySafe: false, failureKind: 'toolError', failureKindRule: FAILURE_KIND_RULE, suggestedAction: 'wait',
    message: 'a restart is ALREADY in progress (opId ' + busy.opId + ', step ' + busy.currentStep + '): a SECOND restart could stop the process the first one just started. NOTHING was stopped by this call.',
    howToProceed: ['call ide_restart{opId:"' + busy.opId + '"} to poll its progress (this starts NOTHING)', 'call ide_list to see the current process fleet', 'if the operation has clearly stalled, use ide_kill_and_relaunch{pid:<pid from ide_list>}'],
    opRule: HUB_OP_RULE }, v);
}
// (FIX-26 / F3) 'still running' and 'a SECOND restart was refused' must not share one token.
// This receipt is the NORMAL, non-failure answer of a bounded call, so it says pending_op;
// restart_in_progress stays the REFUSAL of a second restart (restartInFlightReceipt).
function restartPendingReceipt(out, rec) {
  const pending = JSON.parse(JSON.stringify(out));
  const v = hubOpView(rec);
  return Object.assign(pending, {
    ok: false, pending: true, mode: 'restart', state: 'in_progress', finished: false,
    reason: 'pending_op', refused_reason: '', restarted: false,
    retrySafe: false, suggestedAction: 'wait',
    elapsedMs: v.elapsedMs, currentStep: v.currentStep, lastKnownState: v.lastKnownState,
    opId: rec.opId, opRule: HUB_OP_RULE,
    message: 'the restart was REQUESTED and is still RUNNING (opId ' + rec.opId + ', step ' + v.currentStep + ', elapsed ' + v.elapsedMs + ' ms). This is NOT a failure: the IDE process may already be stopped, relaunched or nearly ready. Do NOT start a second restart.',
    howToProceed: ['call ide_restart{opId:"' + rec.opId + '"} to continue polling (it starts NOTHING)', 'call ide_list to see the pids/ports that exist right now', 'if the operation has clearly stalled, recover with ide_kill_and_relaunch{pid:<pid from ide_list>}'],
    note: 'a client-side REQUEST TIMEOUT is not this tool\'s answer: this receipt (opId + state) is. The hub keeps the op for 30 minutes; a hub restart forgets it.'
  });
}

// WP-22-02: moved to MODULE scope (it was a function declaration nested inside hubTool) because
// runRestartSteps - the restart body the caller no longer blocks on - is a module-level function now.
// (FIX-26 / F1) The transport FACTS of the instance a restart just relaunched, quoted from ITS OWN
// registration: the difference between 'the new instance never came up' (a lie when the add-in
// wrote a pipe record one second after the relaunch) and 'it came up on the PIPE track, which the
// pinned TCP port cannot reach'.
function newInstanceFactsOf(reg, pid) {
  var r = reg || {};
  return { pid: Number(pid), port: Number(r.port) || 0,
           transport: r.transport || null,
           tcpEnabled: (r.tcpEnabled === undefined) ? null : r.tcpEnabled,
           pipe: r.pipe || null,
           project: r.project || null, projectPath: r.projectPath || null,
           registryFile: r.file || null, startedAt: r.startedAt || null,
           registered: r.visible === true,
           addressableByMiddleService: BRIDGE_MODE !== 'inproc' };
}
// A registration is PIPE-ONLY only when it really claims the pipe track: a record with port 0 that
// claims nothing is still 'not ready yet', not 'pipe'.
function pipeOnlyRegistration(reg) {
  if (!reg || reg.visible !== true) return false;
  if (Number(reg.port) !== 0) return false;
  return String(reg.transport || '') === 'pipe' || reg.tcpEnabled === false || !!reg.pipe;
}
function registryCandidateView(list) {
  return (list || []).map(function (a) {
    return { pid: a.pid, port: a.port, transport: a.transport, project: a.project,
             projectPath: a.projectPath, startedByHub: a.startedByHub, registryFile: a.registryFile };
  });
}

async function hubTool(name, args) {
  const snap = snapshot();
  if (name === 'ide_list') return textResult(await hubState({ requestTarget: args && args._target }), snap.contractVersion);
  if (name === 'ide_target') {
    const t = args && args.target;
    state.target = (t === 'auto' || t === '') ? null : t;
    // (WP-22-01) the default is VALIDATED against the LIVE set and the verdict is published:
    // a default naming a dead pid is ignored (never used) AND named, so the caller can see that
    // the hub corrected its own bookkeeping instead of silently keeping a stale pid.
    let liveInsts = [];
    try { const dt = await resolveDiscovery(await discoverSoft(LIST_GRACE_MS)); liveInsts = (dt && dt.instances) || []; } catch (e) { liveInsts = []; }
    const pids = onlinePidsOf(liveInsts);
    const defaultPid = (state.target && /^\d+$/.test(String(state.target))) ? Number(state.target) : null;
    const staleDefaultPid = (defaultPid && pids.indexOf(defaultPid) < 0) ? defaultPid : null;
    const diagReason = staleDefaultPid ? 'hub_record_stale' : (pids.length ? null : 'no_live_ide');
    return textResult({ ok: true, target: state.target,
      cleared: state.target === null,
      liveCount: liveInsts.length, onlinePids: pids,
      liveInstances: liveInsts.map(candidateBrief),
      defaultPid: defaultPid, staleDefaultPid: staleDefaultPid,
      defaultIgnored: !!staleDefaultPid,
      targetDiagnostic: targetDiagnosticBlock({ reason: diagReason, onlinePids: pids, requested: state.target, explicit: false, deadTargetPid: staleDefaultPid }),
      targetResolutionRule: TARGET_RESOLUTION_RULE,
      selectionRule: SELECTION_RULE,
      note: state.target === null
        ? 'the default target is now UNSET: an omitted _target (or _target="auto") follows the documented ladder - the ONLY live IDE, else the ONLY hub-owned IDE, else the ONLY responding IDE, else a refusal that lists every candidate and a suggestion'
        : (staleDefaultPid
          ? 'the default target "' + state.target + '" names a pid that is NOT in the live set: it is IGNORED (it never vetoes a live set) and the live pids are in onlinePids[]'
          : 'every call without an explicit _target now defaults to "' + state.target + '"; a default that matches no live instance is IGNORED (it never vetoes a live set)') }, snap.contractVersion);
  }
  if (name === 'ide_start') {
    const project = (args && args.project) || process.env.VB6MCP_AUTOSTART_PROJECT;
    if (!project) return textResult(Object.assign({ ok: false, reason: 'not_implemented', message: 'ide_start does not auto-start without a project: pass project or set VB6MCP_AUTOSTART_PROJECT' },
      MULTI_BRIDGE ? { failureKind: 'toolError', failureKindRule: FAILURE_KIND_RULE, suggestedAction: 'fix_params', retrySafe: false } : {}), snap.contractVersion);
    // ---- REUSE FIRST (Seg4) ------------------------------------------------------------------
    // Opening the same .vbp twice on one machine is impossible (VB6 refuses) and pointless; the
    // pre-change tool spawned a SECOND process regardless.  A live instance that already owns this
    // project is returned as-is, so the process count and the instance count stay flat.
    const canonP = canonicalProjectPath(project, false);
    let dStart = null;
    try { dStart = await resolveDiscovery(await discoverSoft(LIST_GRACE_MS)); } catch (e) { dStart = null; }
    const sameLocal = ((dStart && dStart.instances) || []).filter(function (i) {
      return !i.remote && i.projectPath && canonP.path &&
        canonicalProjectPath(i.projectPath, false).key === canonP.key;
    });
    if (sameLocal.length) {
      const hit = sameLocal[0];
      // A REUSED instance is NOT automatically ready: probe it with the SAME rule as a fresh start
      // (adversarial review B3: the old constant block claimed ready:true with addInAnswers:null).
      const pr = await probeAddinReadinessOnce(hit.pid, hit);
      // (FIX-26 / F2) projectMatches was null on this path: the probe was never given a project
      // name to compare against, yet ready:true was published beside it - a verdict the four
      // checks did not actually support. The comparison is made for REAL here: the add-in reported
      // its own project name (ide_state.project), and the reused registration names the project it
      // belongs to. When it cannot be made the receipt carries the REASON instead of a bare null.
      const wantProjName = String(hit.project || '').trim();
      const actualProjName = String(pr.projectActual || '').trim();
      const projMatches = actualProjName
        ? (wantProjName ? (actualProjName.toLowerCase() === wantProjName.toLowerCase()) : true)
        : null;
      pr.readyChecks.projectMatches = projMatches;
      pr.readyChecks.projectMatchesActual = actualProjName || null;
      pr.readyChecks.projectMatchesExpected = wantProjName || null;
      pr.readyChecks.projectMatchesBasis = projMatches === null
        ? 'nullReason: the add-in answered but reported no project name (ide_state.project empty), so equality could not be checked'
        : (wantProjName ? 'the add-in reported project ' + actualProjName + ' and the reused registration names ' + wantProjName
                        : 'the registration publishes no project name, so there is nothing to compare against');
      const reuseBody = { ok: true, started: false, reused: true, pid: hit.pid,
        project: hit.projectPath || canonP.path, projectPath: hit.projectPath || canonP.path,
        bridge: MULTI_BRIDGE ? LOCAL_ENDPOINT : undefined, matchedBy: 'project_path',
        ready: pr.ready === true && projMatches !== false, readyWaitMs: pr.readyWaitMs, readyBasis: 'reused_instance_probe',
        readyRule: IDE_START_READY_RULE, readyChecks: pr.readyChecks,
        processesCreated: 0, instancesAdded: 0,
        // (35) NOTHING was spawned on the reuse path: the cwd is fail-closed, not the hub's own.
        cwd: null, cwdBasis: 'unavailable',
        note: 'this .vbp is ALREADY open on this machine: reused the running instance instead of starting a second one' };
      if (!reuseBody.ready) {
        if (projMatches === false) {
          reuseBody.pollingHint = 'the instance WAS reused (no process was started) and its add-in answers, but it reports project ' + actualProjName + ' which is NOT the project its registration names (' + wantProjName + '): ready:false because the four readiness checks do not all hold. Re-run ide_list to see which .vbp that pid really has open.';
          reuseBody.howToProceed = ['call ide_list and compare project/projectPath for pid ' + hit.pid, 'or address the project you actually want explicitly with ide_start{project:<absolute .vbp>}'];
        } else {
          reuseBody.pollingHint = 'the instance WAS reused (no process was started) but the add-in did not answer the readiness probe: poll ide_list for pid ' + hit.pid + ' and retry the call itself; ready:false here is about the PROBE, not about a failed start';
          reuseBody.howToProceed = ['call ide_list and look for pid ' + hit.pid, 'if it stops answering, use ide_kill_and_relaunch{pid:' + hit.pid + '}'];
        }
      }
      return textResult(reuseBody, snap.contractVersion);
    }
    if (MULTI_BRIDGE) {
      let mergedStart = null;
      try { mergedStart = await withRemoteCandidates(dStart || { instances: [], conflicts: [], staleRegistry: [] }, REMOTE_GRACE_MS); }
      catch (e) { mergedStart = null; }
      const sameRemote = ((mergedStart && mergedStart.instances) || []).filter(function (i) {
        return !!i.remote && i.projectPath && canonP.path &&
          canonicalProjectPath(i.projectPath, true).key === canonP.key;
      });
      if (sameRemote.length) {
        const hit = sameRemote[0];
        return textResult({ ok: true, started: false, reused: true, remote: true,
          bridge: hit.bridge || hit.endpoint, pid: hit.pid,
          project: hit.projectPath || canonP.path, projectPath: hit.projectPath || canonP.path,
          matchedBy: 'project_path', processesCreated: 0, instancesAdded: 0,
          // (35) a REMOTE instance was reused: this hub spawned no process, so cwd is fail-closed.
          cwd: null, cwdBasis: 'unavailable',
          ready: false, readyWaitMs: 0, readyBasis: 'remote_endpoint_not_probed',
          readyChecks: { registryVisible: true, transportUp: null, addInAnswers: null, projectMatches: null,
                         addInAnswersBasis: 'remote_endpoint_not_probed' },
          pollingHint: 'the remote instance was REUSED (nothing was started), but this hub cannot probe the add-in on another machine: poll ide_list on that endpoint or call the tool with its endpoint-qualified target',
          note: 'this .vbp is ALREADY open on bridge "' + (hit.bridge || hit.endpoint) + '": reused that instance; NOTHING was started' }, snap.contractVersion);
      }
      if (!fs.existsSync(project)) {
        // REMOTE START: this hub cannot create a process on another machine, and this build has no
        // bridge-side launcher verb.  FAIL CLOSED with the exact configuration needed - and do NOT
        // touch any process table.
        const labels = (mergedStart && mergedStart.endpoints
          ? mergedStart.endpoints.map(function (e) { return e.label; })
          : BRIDGE_ENDPOINTS.map(function (e) { return e.label; }));
        const onlyBridge = labels.length === 1 ? labels[0] : null;
        return textResult({ ok: false, reason: 'bridge_launcher_not_configured',
          failureKind: 'toolError', failureKindRule: FAILURE_KIND_RULE,
          suggestedAction: 'fix_params', retrySafe: false,
          project: project, projectPath: canonP.path || project,
          bridge: onlyBridge, bridges: labels,
          attemptedSpawn: false, processesCreated: 0, spawnSkipped: 'remote_bridge_no_launcher',
          remoteStart: true,
          message: 'the project does not exist on THIS machine and the hub cannot start a process on another machine: a REMOTE start needs a launcher on the owning bridge, and this build implements no bridge-side launcher verb. NOTHING was started (attemptedSpawn=false).',
          howToConfigure: {
            required: 'a launcher on the bridge that owns the project' + (onlyBridge ? ' (bridge "' + onlyBridge + '")' : ''),
            bridgeSideVerb: 'ide_start (NOT implemented by vb6mcp_bridge.py in this build; the bridge is read-only by design)',
            configKey: 'VB6MCP_BRIDGES[].launcherCommand (reserved; ignored until the verb exists)',
            example: '{"name":"<bridge>","url":"http://<host>:8796","tokenFile":"<file>","launcherCommand":"<start-ide.cmd> <project>"}',
            alternative: 'start the IDE on that machine yourself (its scheduler/task), then call ide_list; a forwarded call can then use _target="<bridge>:<pid|project|projectPath>"',
            noProcessCreated: true
          },
          hint: 'pass a project path that exists on THIS machine for a local start, or start the IDE on the bridge machine first' }, snap.contractVersion);
      }
    }
    // ---- local start (the pre-change path; checks happen only after the reuse checks) --------
    const rStartExe = resolveVb6Exe();
    const exe = rStartExe.exe;
    if (!exe) return textResult(Object.assign({ ok: false, reason: 'vb6_exe_not_found', vb6Exe: '', vb6ExeSource: '', vb6ExeTried: rStartExe.tried, processesCreated: 0, attemptedSpawn: false, message: 'no VB6.EXE was resolvable; set VB6MCP_VB6EXE or install VB6. Tried: ' + rStartExe.tried.join(' | ') },
      MULTI_BRIDGE ? { failureKind: 'toolError', failureKindRule: FAILURE_KIND_RULE, suggestedAction: 'fix_params', retrySafe: false } : {}), snap.contractVersion);
    if (!fs.existsSync(project)) return textResult(Object.assign({ ok: false, reason: 'project_not_found', project: project, processesCreated: 0, attemptedSpawn: false },
      MULTI_BRIDGE ? { failureKind: 'toolError', failureKindRule: FAILURE_KIND_RULE, suggestedAction: 'fix_params', retrySafe: false } : {}), snap.contractVersion);
    const pinPort = await pickFreePort();
    const childEnv = Object.assign({}, process.env);
    if (pinPort) childEnv.VB6MCP_PORT = String(pinPort);
    // (35) captured AT THE SPAWN POINT: no cwd option is passed, so the child inherits this hub
    // process's cwd; if it cannot be read the receipt says "unavailable" (never guessed).
    const spawnOpts = { detached: true, stdio: 'ignore', env: childEnv };
    const spawnCwd = hubSpawnCwd(spawnOpts);
    const spawnCwdBasis = spawnCwd === null ? 'unavailable' : 'spawn_cwd';
    const child = spawnProgram(exe, [project], spawnOpts);
    child.unref();
    startedByHub.set(child.pid, { project: project, port: pinPort || null });
    invalidateDiscovery();
    // (WP-22-01) READINESS, not a fixed sleep: ok:true + started:true only claims the PROCESS;
    // ready:true requires the add-in to have ANSWERED mcp_info for that pid within the budget.
    const budget = ideStartReadyBudgetMs(args);
    const wantName = String(project).replace(/\\/g, '/').split('/').pop().replace(/[.]vbp$/i, '');
    const rw = await waitForAddinReady(child.pid, { budgetMs: budget, wantProjectName: wantName, hintPort: pinPort || 0 });
    // (B1, t16) an instance the hub JUST created must be visible in ide_list even when the add-in
    // wrote no registration file: publish the minimal one ONLY after the add-in proved the port.
    const regInfo = rw.ready
      ? hubWriteRegistration(child.pid, (rw.port || pinPort || 0), canonP.path || project, wantName)
      : { written: false, by: null, path: path.join(instancesDir(), String(child.pid) + '.json'), existed: false,
          readbackVisible: false, owner: 'add-in (writes its own entry on listen success)',
          reason: 'the add-in did not answer on any route, so nothing was registered: ide_list will NOT list this pid until the add-in starts and listens' };
    const startBody = { ok: true, started: true, pid: child.pid, project: project, projectPath: canonP.path || project,
      cwd: spawnCwd, cwdBasis: spawnCwdBasis,
      processesCreated: 1, pinnedPort: pinPort || null, port: rw.port || null,
      ready: !!rw.ready, readyWaitMs: rw.readyWaitMs, readyBudgetMs: rw.readyBudgetMs,
      readyChecks: rw.readyChecks, readyRule: IDE_START_READY_RULE, readyWaitSemantics: IDE_START_READY_RULE,
      readyProbeRounds: rw.readyProbeRounds, readyProbeRule: rw.readyProbeRule,
      readyRegistrationFile: rw.readyRegistrationFile || null, readyPortSource: rw.readyPortSource || null,
      registration: regInfo,
      note: rw.ready
        ? 'the IDE was created AND the add-in answered within ' + rw.readyWaitMs + ' ms: the pid is usable as _target immediately'
        : 'the IDE process WAS created (ok:true / started:true) but readiness timed out after ' + rw.readyWaitMs + ' ms: ready:false is the truth, and the pid is still valid' };
    if (rw.addInAnswersError) startBody.addInAnswersError = rw.addInAnswersError;
    if (!rw.ready) {
      startBody.pollingHint = rw.pollingHint;
      startBody.registrationNote = 'the spawned IDE did not register itself: no add-in answered on the pinned port, so the hub wrote NOTHING (it never invents an instance). Check that the add-in is loaded/enabled in that IDE (Add-In Manager), that its listen succeeded, and that it runs under the SAME user/TEMP the hub reads (' + instancesDir() + '). Until then ide_list cannot list this pid and ide_kill_and_relaunch/ide_restart cannot resolve it.';
      startBody.howToProceed = (rw.howToProceed || []).concat(['load/enable the VB6MCP add-in in the IDE that was just started (Add-In Manager), then poll ide_list', 'or start the IDE yourself and use ide_list once its add-in has registered']);
      startBody.readyNote = 'ready:false does NOT mean the start failed: poll ide_list for this pid (the receipt says how)';
    }
    return textResult(startBody, snap.contractVersion);
  }
  if (name === 'ide_kill_and_relaunch') return await relaunchOwnedInstance(args, snap);
  if (name === 'ide_stop') {
    const pid = args && args.pid;
    if (!pid) return textResult({ ok: false, reason: 'not_implemented', message: 'ide_stop requires an explicit pid and only stops hub-started instances' }, snap.contractVersion);
    // ---- Seg4: a pid that belongs to a BRIDGE instance is refused HERE ---------------------
    if (MULTI_BRIDGE) {
      const remoteHits = await remoteInstancesForPid(pid);
      if (remoteHits.length) {
        let localHit = null;
        try { const ddR = await resolveDiscovery(await discoverSoft(LIST_GRACE_MS)); localHit = (ddR.instances || []).filter(function (i) { return Number(i.pid) === Number(pid); })[0] || null; } catch (e) { localHit = null; }
        if (localHit) return textResult(remoteAmbiguousRefusal('ide_stop', pid, localHit, remoteHits), snap.contractVersion);
        return textResult(remoteControlRefusal('ide_stop', pid, remoteHits[0]), snap.contractVersion);
      }
    }
// /N12 +: the SAME ownership predicate AND the SAME resolver as ide_restart.
    // The old fallback was '{pid, startedByHub: <in-memory map>}' - it consulted NO registry, so when
    // discovery was blind (stale TCP entry / unreachable middle service) ide_stop answered 'not_owned'
    // for the very pid ide_restart called 'registry_started_by_hub', i.e. two tools disagreed about
    // one process and told the operator to leave OUR OWN wedged instance alone. The registry fallback
    // below carries the real facts (port, startedByHub, project), so the ownership verdict - and the
    // port facts that ride with it - are identical on both faces.
    let inst = null;
    let instSource = 'none';
    try {
      const dd = await resolveDiscovery(await discoverSoft(LIST_GRACE_MS));
      inst = (dd.instances || []).filter(function (i) { return Number(i.pid) === Number(pid); })[0] || null;
      if (inst) instSource = 'discovery';
    } catch (e) { inst = null; }
    if (!inst) {
      const al = registryAliveCandidates(Number(pid));
      if (al.length) {
        const a = al[0];
        inst = { pid: a.pid, port: a.port, transport: a.transport, project: a.project, projectPath: a.projectPath,
                 startedByHub: a.startedByHub, registryFile: a.registryFile, discoveredBy: 'registry' };
        instSource = 'registry';
      }
    }
    if (!inst) { inst = { pid: Number(pid), startedByHub: startedByHub.has(Number(pid)) }; instSource = 'in_memory_map'; }
    const own = ownershipOf(inst);
    // : the same ATTRIBUTION block ide_list / ide_state / ide_restart publish, so the four
    // entrances cannot describe one pid four ways.  The DECISION below still uses the permission
    // axis (own.owned) - attribution never grants the right to kill a process.
    const ovStop = await resolveOwnership(inst);
    const body = { ok: false, stopped: 0, pid: Number(pid), ownership: own.ownership, ownershipBasis: own.ownershipBasis,
                   ownedByHub: own.owned, port: own.port, portRuleApplies: own.portRuleApplies, portBasis: own.portBasis,
                   ownershipRule: OWNERSHIP_RULE, ownershipVerdict: ovStop, ownershipVerdictRule: OWNERSHIP_VERDICT_RULE,
                   liveness: inst && inst.liveness ? inst.liveness : null,
                   instanceSource: instSource, registryFile: inst.registryFile || null, startedByHub: !!inst.startedByHub,
                   alsoForeignPort: !!own.foreignPort };
    // OWNERSHIP FIRST. The port guard is a SECONDARY veto and must never mask the
    // ownership verdict: measured on this host, an instance that is simply NOT OURS on a port below
    // 8767 (any ephemeral listener - 3469/3491 here) was reported as 'foreign_instance_refused',
    // i.e. the port story stood in for 'not owned'. Both still refuse (nothing is ever stopped), but
    // the reason now names the real judgement and the port facts ride along as fields.
    body.alsoForeignPort = !!own.foreignPort;
    if (MULTI_BRIDGE) { body.failureKind = 'toolError'; body.failureKindRule = FAILURE_KIND_RULE; body.suggestedAction = 'fix_params'; body.retrySafe = false; }
    if (!own.owned) {
      body.reason = 'not_allowed';
      body.message = 'pid ' + pid + ' is not owned by this hub: ' + own.ownershipBasis + '; not stopping it (the same ownership rule ide_list and ide_restart report)' + (own.foreignPort ? ' The port guard also applies: ' + own.portBasis + '.' : '');
      return textResult(body, snap.contractVersion);
    }
    if (own.foreignPort) {
      body.reason = 'foreign_instance_refused';
      if (MULTI_BRIDGE) { body.failureKind = 'toolError'; body.failureKindRule = FAILURE_KIND_RULE; body.suggestedAction = 'fix_params'; body.retrySafe = false; }
      body.message = own.portBasis + ' => this instance is not ours; not stopping it';
      return textResult(body, snap.contractVersion);
    }
    try { process.kill(Number(pid)); startedByHub.delete(Number(pid)); invalidateDiscovery();
          return textResult(Object.assign(body, { ok: true, stopped: Number(pid), message: 'stopped pid ' + pid + ' (' + own.ownership + ')' }), snap.contractVersion); }
    catch (e) { return textResult(Object.assign(body, { reason: 'kill_failed', message: String(e && e.message || e) },
      MULTI_BRIDGE ? { failureKind: 'toolError', failureKindRule: FAILURE_KIND_RULE, suggestedAction: 'retry_once', retrySafe: true } : {}), snap.contractVersion); }
  }

  if (name === 'ide_restart') {
    const a0 = args || {};
    // ---- WP-22-02 / MCP-1021-33: a CONTINUABLE operation, and a BOUNDED wait ----------------
    // A restart is not a failure just because a client stopped waiting. ide_restart{opId} answers
    // from the op registry and touches NOTHING; a SECOND restart while one is in flight is
    // refused (it could stop the process the first one just started).
    if (a0.opId !== undefined && a0.opId !== null && String(a0.opId).trim() !== '') {
      const q = hubOpById(String(a0.opId).trim());
      if (!q) return textResult(hubOpNotFound('ide_restart', a0.opId), snap.contractVersion);
      const qv = hubOpView(q);
      return textResult(Object.assign({ ok: q.state === 'done', tool: 'ide_restart', mode: 'status', statusQuery: true }, qv), snap.contractVersion);
    }
    const busyOp = hubOpInFlight('ide_restart');
    if (busyOp) return textResult(restartInFlightReceipt(busyOp), snap.contractVersion);
    const opRec = startHubOp('ide_restart', { pid: a0.pid || null, project: a0.project || null });
    const pause = function (ms) { return new Promise(function (r) { setTimeout(r, ms); }); };
    // The executable recovery for a wedged IDE. It only ever
    // touches OUR OWN instance (instances listening on port >= 8767; 8766 belongs to a
    // foreign process on this machine and is NEVER stopped), and EVERY path answers with
    // the same key set (discipline 98).
    const out = {
      ok: false, tool: 'ide_restart', restarted: false,
      stoppedPid: 0, newPid: 0, port: 0, project: '',
      targetSource: '', ownership: '', refused_reason: '', stopConfirmed: false, pinPort: 0, projectVbp: '',
      saveAttempted: false, saveOk: false, registryCleared: false, lockCleared: false,
      three_way_ok: false, threeWay: { http: false, mcp_info: false, project: false },
      vb6Exe: '', vb6ExeSource: '', vb6ExeTried: [], waitMs: 0, readyWaitMs: 0, instances: [], steps: [], note: '',
// /N12: the ownership verdict and its basis are part of EVERY path's key set.
      ownershipBasis: '', ownedByHub: false, portRuleApplies: false, portBasis: '', ownershipRule: OWNERSHIP_RULE,
      // The ATTRIBUTION verdict (whose instance is it?) is a SEPARATE axis from the
      // permission axis above (did THIS hub start it?).  The block is present on EVERY path (null
      // when there is no candidate instance at all), so the key set stays fixed.
      ownershipVerdict: null, ownershipVerdictRule: OWNERSHIP_VERDICT_RULE,
      ownershipUnknown: false, legacyReason: '', howToDisambiguate: '',
      liveness: null, processKindRule: PROCESS_KIND_RULE,
      // the honesty keys - EVERY path carries them (discipline 98), so 'no_ide_running'
      // can no longer stand in for 'an instance is alive but did not answer'.
      instanceCount: 0, registryCandidates: [], manualCommand: '', truth: '', nextStep: '',
      checkedRegistryDir: '',
      // ONE refusal shape: an ambiguous_target from THIS tool carries the SAME canonical payload a
      // forwarded call refuses with (candidates / suggestion / resolution + the rule + failureKind),
      // so the reason does not mean one thing on one entrance and another thing on the other.
      candidates: [], suggestion: null, resolution: null, targetResolutionRule: TARGET_RESOLUTION_RULE,
      failureKind: '', failureKindRule: FAILURE_KIND_RULE, conflicts: [],
      // WP-22-02: the continuable-operation keys - present on EVERY path
      mode: 'restart', opId: '', started: true, state: 'in_progress', finished: false,
      currentStep: 'resolve_target', lastKnownState: 'started', elapsedMs: 0, opRule: HUB_OP_RULE,
      // (35) the CWD of the restart launch: filled at the spawn step; an operation that refuses
      // or aborts BEFORE spawning keeps this fail-closed null / "unavailable".
      cwd: null, cwdBasis: 'unavailable'
    };

    out.opId = opRec.opId;
    opRec.out = out; opRec.steps = out.steps;
    const runP = runRestartSteps(out, a0, opRec).then(function () {
      hubOpDone(opRec, out.ok ? 'done' : 'refused');
    }, function (e) {
      out.ok = false; out.restarted = false; out.reason = 'restart_error'; out.refused_reason = 'restart_error';
      out.note = 'the restart aborted with an internal error: ' + String((e && e.message) || e).slice(0, 300) + ' -- the process state is UNKNOWN; call ide_list before any further action and do NOT start a second restart.';
      hubOpDone(opRec, 'error');
    }).then(function () {
      out.state = opRec.state; out.finished = true; out.elapsedMs = opRec.finishedAt - opRec.startedAt;
      out.currentStep = opRec.currentStep; out.lastKnownState = opRec.lastKnownState;
      return out;
    });
    const waitMs = resolveRestartWaitMs(a0.wait_ms);
    const winner = await Promise.race([runP.then(function () { return 'done'; }), pause(waitMs).then(function () { return 'pending'; })]);
    if (winner === 'pending') return textResult(restartPendingReceipt(out, opRec), snap.contractVersion);
    return finishControlOut(out, snap);
  }
}

// (B1) how long resolve_target POLLS the add-in of a registered+alive instance before it may
// call it unresponsive.  wait_ms only bounds the CALL - the op keeps running (that is what makes
// the pending -> opId -> done path the normal one).
const RESTART_PROBE_READY_MS = Math.max(1000, parseInt(process.env.VB6MCP_RESTART_PROBE_READY_MS || '15000', 10) || 15000);
// (37) The expected PROJECT NAME for ide_restart's new-instance readiness comes from the NEW
// instance's OWN registration: the add-in writes registration.project from the project VB6
// actually opened, and VB6 may REWRITE the .vbp Name= (W44 measured "_test_proj" -> "Vb_test_proj").
// So the .vbp FILE BASENAME is never the first authority - it is only a RECORDED fallback.
// wantNameSource (published on the readiness receipt) is exactly one of:
//   "registration"      - the name came from the new instance's own registration (or its probe)
//   "basename_fallback" - the registration publishes no project name; .vbp basename used instead
//   "unavailable"       - neither exists; the expectation is empty and the verdict stays null
function restartProjectExpectation(reg, inst, projectVbp) {
  const regName = String((reg && reg.project) || '').trim();
  if (regName) return { name: regName, source: 'registration' };
  const progName = String((inst && inst.project) || '').trim();
  if (progName) return { name: progName, source: 'registration' };
  let base = '';
  try {
    base = String(projectVbp || '').replace(/\\/g, '/').split('/').pop().replace(/[.]vbp$/i, '').trim();
  } catch (e) { base = ''; }
  return base ? { name: base, source: 'basename_fallback' } : { name: '', source: 'unavailable' };
}
// (37) ONE comparison rule for the same-instance self-evidence: ide_state.project (what the
// add-in says it opened) vs registration.project (what that SAME instance registered) -
// trim + case-insensitive.  A missing side is null (UNVERIFIED), never a fabricated true.
function projectNameEqual(actual, expected) {
  const a = String(actual === undefined || actual === null ? '' : actual).trim();
  const e = String(expected === undefined || expected === null ? '' : expected).trim();
  if (!a || !e) return null;
  return a.toLowerCase() === e.toLowerCase();
}
// WP-22-02: the restart steps, moved out of hubTool so the CALLER can stop waiting without
// the operation stopping. It answers with the SAME `out` object the blocking version built.
async function runRestartSteps(out, args, rec) {
  const pause = function (ms) { return new Promise(function (r) { setTimeout(r, ms); }); };
  out.argsEcho = { pid: (args && args.pid) || null, project: (args && args.project) || null };
    const wantPid = (args && args.pid) ? Number(args.pid) : 0;
    const wantProject = (args && args.project) || process.env.VB6MCP_AUTOSTART_PROJECT || '';
    const soft = await discoverSoft(LIST_GRACE_MS);
    const d = await resolveDiscovery(soft);
    markHubOpStep(out, rec, 'resolve_target', 'target_resolved');
    const insts = (d.instances || []).filter(function (i) { return i.pid; });
    // ---- Seg4: a REMOTE instance is never restarted by this hub; a pid that exists on several
    // bridges is refused as ambiguous instead of being resolved to whichever one is local.
    if (MULTI_BRIDGE && wantPid) {
      const remoteHits = await remoteInstancesForPid(wantPid);
      if (remoteHits.length) {
        const localHit = insts.filter(function (i) { return Number(i.pid) === wantPid; })[0] || null;
        if (localHit) {
          const ambR = remoteAmbiguousRefusal('ide_restart', wantPid, localHit, remoteHits);
          Object.keys(ambR).forEach(function (k) { out[k] = ambR[k]; });
          out.truth = 'ambiguous_pid_across_bridges';
          return out;
        }
        const rr = remoteControlRefusal('ide_restart', wantPid, remoteHits[0]);
        Object.keys(rr).forEach(function (k) { out[k] = rr[k]; });
        out.truth = 'remote_bridge_instance';
        return out;
      }
    }
    const rCtx = bridgeIdentitySync();
    out.instances = insts.map(function (i) {
      const o = ownershipOf(i);
      const lv = livenessOf(i, rCtx);
      return { pid: i.pid, port: i.port, project: i.project,
               ownership: o.ownership, ownedByHub: o.owned, ownershipBasis: o.ownershipBasis,
               portRuleApplies: o.portRuleApplies, portBasis: o.portBasis,
               processKind: i.processKind || 'vb6_ide', processKindBasis: i.processKindBasis || null,
               liveness: lv.state, livenessBasis: lv.basis };
    });
    let target = null;
    if (wantPid) {
      target = insts.filter(function (i) { return i.pid === wantPid; })[0] || null;
      out.targetSource = 'explicit_pid';
      if (!target) {
        // a pid that is REGISTERED AND ALIVE but unresponsive is not 'not found'.
        const a1 = registryAliveCandidates(wantPid);
        if (a1.length) {
          const a = a1[0]; const o = ownershipOf(a);
          out.pid = a.pid; out.port = a.port || 0; out.project = a.project || ''; out.projectVbp = a.projectPath || '';
          out.ownership = o.ownership; out.ownershipBasis = o.ownershipBasis; out.ownedByHub = o.owned;
          out.portRuleApplies = o.portRuleApplies; out.portBasis = o.portBasis;
          out.instanceSource = 'registry'; out.registryFile = a.registryFile; out.startedByHub = !!a.startedByHub;
          out.alsoForeignPort = !!o.foreignPort;
          out.liveness = 'not_responding'; out.instanceCount = 1; out.registryCandidates = registryCandidateView(a1);
          out.checkedRegistryDir = instancesDir();
          out.manualCommand = 'Stop-Process -Id ' + a.pid + ' -Force';
          out.truth = 'instance_registered_and_pid_alive';
          // : ATTRIBUTION is decided by the ONE resolver, and an unknown attribution is
          // reported as such instead of being dressed up as not_owned.
          // (B1) THE SAME PROBE THE FORWARD PATH USES, POLLED - not one soft discovery read.
          // A registered+alive instance the discovery pass has not probed yet is NOT
          // 'unresponsive': wait for the add-in (wait_ms bounds the CALL, not the probe) and only
          // then refuse - with readiness-shaped probeChecks so the caller sees what failed.
          markHubOpStep(out, rec, 'probe_ready', 'probing_instance');
          const rw = await waitForAddinReady(a.pid, { budgetMs: RESTART_PROBE_READY_MS, wantProjectName: String(a.project || '') });
          out.probeChecks = rw.readyChecks; out.probeBudgetMs = RESTART_PROBE_READY_MS;
          out.readyWaitMs = rw.readyWaitMs; out.readyProbeRounds = rw.readyProbeRounds;
          out.readyProbeRule = READY_PROBE_RULE; out.readyRegistrationFile = rw.readyRegistrationFile || null;
          if (rw.ready) {
            out.targetSource = 'registry_after_probe'; out.liveness = 'responding';
            target = { pid: a.pid, port: rw.port || a.port || 0, transport: a.transport || null,
                       project: a.project || null, projectPath: a.projectPath || null,
                       startedByHub: !!a.startedByHub, registryFile: a.registryFile,
                       discoveredBy: 'registry', probeStatus: 'ok',
                       liveness: { state: 'responding', basis: 'the add-in answered mcp_info after ' + rw.readyWaitMs + ' ms' } };
          } else {
          const ovA1 = await resolveOwnership(a);
          out.ownershipVerdict = ovA1;
          if (ovA1.verdict === 'unknown') {
            ownershipUnknownRefusal(out, ovA1, o.foreignPort ? 'not_ours_by_red_line' : 'instance_found_but_unresponsive');
            out.truth = 'instance_found_but_ownership_unknown';
            out.note = 'pid ' + wantPid + ' is registered (' + a.registryFile + ') and its process is ALIVE, but WHOSE instance it is cannot be established: ' + ovA1.basis.join('; ') + '. Nothing was stopped.';
            out.nextStep = 'disambiguate first (howToDisambiguate): the manual command "' + out.manualCommand + '" is only for an instance you have confirmed is yours.';
            return out;
          }
          out.refused_reason = o.foreignPort ? 'not_ours_by_red_line' : 'instance_found_but_unresponsive';
          out.note = 'pid ' + wantPid + ' is not a LIVE instance, but it IS registered (' + a.registryFile + ') and its process is ALIVE, so it is unresponsive rather than absent. ' + (o.foreignPort ? o.portBasis + '; not ours, nothing was stopped.' : 'Ownership: ' + o.ownershipBasis + '. Nothing was stopped.');
          out.nextStep = 'run "' + out.manualCommand + '" on the machine that owns the pid and retry, or use ide_stop{pid:' + a.pid + '} when the ownership verdict above is owned.';
          out.howToProceed = [
            'retry ide_restart{pid:<pid>} - the probe is repeated; wait_ms only bounds THIS call, the op continues',
            'or check it first with ide_state{_target:<pid>}: that is the SAME probe the forward path uses',
            'if it stays unresponsive, stop it manually (' + out.manualCommand + ') and start a fresh IDE with ide_start'
          ];
          out.retrySafe = true; out.suggestedAction = 'retry_once';
          return out;
          }
        }
        if (!target) {
        out.refused_reason = 'pid_not_found';
        out.instanceCount = 0;
        out.checkedRegistryDir = instancesDir();
        out.truth = 'no_instance_at_all';
        out.note = 'pid ' + wantPid + ' is not a live instance AND no registration with that pid has an alive process; see instances[]';
        return out;
        }
      }
    } else {
      // never report 'no_ide_running' when the truth is 'this hub could not see the fleet'.
  if (!insts.length && bridgeLastFailure) {
    out.refused_reason = bridgeLastFailure.kind === 'auth' ? 'bridge_auth_failed' : 'bridge_unreachable';
    out.authRequired = bridgeLastFailure.kind === 'auth';
    out.authReason = bridgeLastFailure.kind === 'auth' ? bridgeLastFailure.reason : null;
    out.authStatus = bridgeLastFailure.status || null;
    out.authHint = bridgeLastFailure.kind === 'auth'
      ? 'the middle service rejected this hub; set VB6MCP_TOKEN (or VB6MCP_TOKEN_FILE) for the hub and retry. Nothing was restarted.'
      : 'the middle service did not answer; the instance list cannot be trusted. Nothing was restarted.';
    out.instanceCount = registryAliveCandidates(0).length;
    out.registryCandidates = registryCandidateView(registryAliveCandidates(0));
    out.checkedRegistryDir = instancesDir();
    out.truth = out.instanceCount ? 'instance_registered_and_pid_alive_but_the_middle_service_is_not_answering' : 'middle_service_not_answering';
    out.note = 'no live instance is visible' + (bridgeLastFailure.kind === 'auth' ? ' BECAUSE THE MIDDLE SERVICE REJECTED THIS HUB (' + bridgeLastFailure.reason + (bridgeLastFailure.status ? ', HTTP ' + bridgeLastFailure.status : '') + ')' : ' because the middle service did not answer') + ' - not because the machine is idle';
    return out;
  }
  if (!insts.length) {
// HONESTY LADDER. 'no_ide_running' means NO INSTANCE AT ALL. A registration whose pid
    // is ALIVE but whose probe failed IS an instance: say so, name the pid, and print the command
    // that actually recovers the machine (the card's requirement (1)).
    const alive = registryAliveCandidates(0);
    out.instanceCount = alive.length;
    out.registryCandidates = registryCandidateView(alive);
    out.checkedRegistryDir = instancesDir();
    if (alive.length) {
      const a0 = alive[0]; const o0 = ownershipOf(a0);
      out.pid = a0.pid; out.port = a0.port || 0; out.project = a0.project || ''; out.projectVbp = a0.projectPath || '';
      out.ownership = o0.ownership; out.ownershipBasis = o0.ownershipBasis; out.ownedByHub = o0.owned;
      out.portRuleApplies = o0.portRuleApplies; out.portBasis = o0.portBasis;
      out.liveness = 'not_responding';
      out.manualCommand = 'Stop-Process -Id ' + a0.pid + ' -Force';
      out.truth = 'instance_registered_and_pid_alive';
      // the SAME precedence as ide_stop/ide_restart - ownership first, the port guard second.
      out.alsoForeignPort = !!o0.foreignPort;
      const ov0 = await resolveOwnership(a0);
      out.ownershipVerdict = ov0;
      if (ov0.verdict === 'unknown') {
        ownershipUnknownRefusal(out, ov0, o0.foreignPort ? 'not_ours_by_red_line' : 'instance_found_but_unresponsive');
        out.truth = 'instance_found_but_ownership_unknown';
        out.note = 'pid ' + a0.pid + ' is registered (' + a0.registryFile + ') and its process is ALIVE, but WHOSE instance it is cannot be established: ' + ov0.basis.join('; ') + '. Nothing was stopped.';
        out.nextStep = 'disambiguate first (howToDisambiguate): the manual command "' + out.manualCommand + '" is only for an instance you have confirmed is yours.';
        return out;
      }
      if (!o0.owned) {
        out.refused_reason = 'not_owned';
        out.note = 'pid ' + a0.pid + ' IS REGISTERED AND ALIVE but is not ours (' + o0.ownershipBasis + '), so nothing was stopped'
          + (o0.foreignPort ? '. The port guard also applies: ' + o0.portBasis + '.' : '.');
      } else if (o0.foreignPort) {
        out.refused_reason = 'not_ours_by_red_line';
        out.note = 'pid ' + a0.pid + ' IS REGISTERED AND ALIVE and IS ours, but ' + o0.portBasis + ' - the red line refuses it. Nothing was stopped.';
      } else {
        out.refused_reason = 'instance_found_but_unresponsive';
        out.note = 'an instance IS running: pid ' + a0.pid + ' (' + (a0.projectPath || a0.project || 'unknown project') + ') is registered in ' + a0.registryFile + ' and its process is ALIVE, but it did not answer the probe. This is NOT an idle machine. Nothing was stopped' + (o0.owned ? '' : ' (ownership: ' + o0.ownershipBasis + ')') + '.';
      }
      out.nextStep = 'restart the IDE process: run "' + out.manualCommand + '" on the machine that owns the pid, then call ide_restart again (or use ide_stop{pid:' + a0.pid + '} when the ownership verdict above is owned).';
      return out;
    }
    out.refused_reason = 'no_ide_running';
    out.truth = 'no_instance_at_all';
    out.nextStep = 'start an IDE (ide_start) or pass a pid that exists.';
    out.note = 'nothing to restart: the registry (' + out.checkedRegistryDir + ') holds no entry whose process is alive - no instance at all';
    return out;
  }
      if (wantProject) {
        target = insts.filter(function (i) { return i.project && String(i.project).toLowerCase() === wantProject.toLowerCase(); })[0] || null;
        if (target) out.targetSource = 'project';
      }
      if (!target) {
        if (insts.length === 1) { target = insts[0]; out.targetSource = 'only_instance'; }
        else {
          // Any ambiguous_target refusal is built by the ONE builder a forwarded call uses
          // (refusalBody), so both entrances publish the same candidates[] / suggestion{} /
          // resolution{} / targetResolutionRule / failureKind.  The tool's legacy per-instance view
          // stays in out.instances; the canonical payload rides BESIDE it (additive fields).
          const ownedN = insts.filter(function (x) { return ownershipOf(x).owned; }).length;
          const pickAmb = { ok: false, ambiguous: true,
            candidates: insts.map(candidateBrief), suggestion: suggestionFor(insts),
            resolution: resolutionOf('refuse_ambiguous_target', insts, { requested: null, explicit: false },
              { ladderVerdict: 'refused', ownedCount: ownedN, respondingCount: respondingOf(insts).length }),
            message: 'several instances online; pass pid or project (or one of the candidate targetKey values) -- ' + candidateDesc(insts, false) };
          const canon = refusalBody(pickAmb, insts, (d.conflicts || []));
          out.refused_reason = canon.reason;
          out.note = canon.message;
          out.candidates = canon.candidates;
          out.suggestion = canon.suggestion;
          out.resolution = canon.resolution;
          out.targetResolutionRule = canon.targetResolutionRule;
          out.failureKind = canon.failureKind;
          out.failureKindRule = canon.failureKindRule;
          out.conflicts = canon.conflicts;
          out.instanceCount = insts.length;
          return out;
        }
      }
    }
// /N12: ONE ownership predicate, shared with ide_list and ide_stop. The port is a
    // SECONDARY guard and it can only apply when a port is actually KNOWN: a pipe-only instance
    // publishes port 0 ON PURPOSE, and reading that 0 as "below 8767 => foreign" is exactly the
    // defect that refused this recovery tool its own wedged IDE while ide_list called it ours.
    const own = ownershipOf(target);
    out.ownership = own.ownership;
    out.ownershipBasis = own.ownershipBasis;
    out.ownedByHub = own.owned;
    out.portRuleApplies = own.portRuleApplies;
    out.portBasis = own.portBasis;
    out.liveness = target.liveness || null;
    // a refusal that already KNOWS the pid must name it and print what recovers the
    // machine (the same honesty keys as the unresponsive path). 'Nothing was stopped' is only half
    // an answer - the operator still has to know WHICH pid is sitting there.
    out.pid = target.pid;
    out.port = target.port || 0;
    out.checkedRegistryDir = instancesDir();
    out.instanceCount = insts.length || (target ? 1 : 0);
    out.registryCandidates = registryCandidateView(registryAliveCandidates(0));
    out.manualCommand = 'Stop-Process -Id ' + target.pid + ' -Force';
    out.nextStep = 'this tool will not touch pid ' + target.pid + '; if you own it, run "' + out.manualCommand + '" and start a fresh IDE (ide_start).';
    // OWNERSHIP FIRST (same rule as ide_stop, so both answer the same way about the same
    // pid). The port guard stays a SECONDARY veto: it can only fire for an instance that IS ours and
    // whose KNOWN port is below 8767 (the reserved range); it can never stand in for 'not owned'.
    out.alsoForeignPort = !!own.foreignPort;
    // : the ATTRIBUTION verdict (whose instance is it?) is decided by the ONE resolver and
    // reported on every path.  It is a DIFFERENT axis from the permission decision below.
    const ovT = await resolveOwnership(target);
    out.ownershipVerdict = ovT;
    if (ovT.verdict === 'unknown') {
      ownershipUnknownRefusal(out, ovT, (!own.owned ? 'not_owned' : (own.foreignPort ? 'not_ours_by_red_line' : 'instance_found_but_unresponsive')));
      out.truth = 'instance_found_but_ownership_unknown';
      out.note = 'pid ' + target.pid + ' is alive, but WHOSE instance it is cannot be established: ' + ovT.basis.join('; ') + '. Nothing was stopped.';
      return out;
    }
    if (!own.owned) {
      out.refused_reason = 'not_owned';
      out.truth = 'instance_found_but_not_owned';
      out.note = 'this instance is not ours (' + own.ownershipBasis + '), so NOTHING was stopped. The ownership rule is shared with ide_list and ide_stop, so all three agree about pid ' + target.pid + '.' + (own.foreignPort ? ' The port guard also applies: ' + own.portBasis + '.' : '');
      return out;
    }
    if (own.foreignPort) {
      out.refused_reason = 'foreign_instance_refused';
      out.truth = 'instance_found_but_not_ours';
      out.note = own.portBasis + ': refusing to stop it.';
      return out;
    }
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
      return out;
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
      return out;
    }
    // spawn() must receive a real .vbp PATH, never a project NAME, and the
    // registry entry is written BEFORE the project is loaded - the first live run passed
    // the name "VB6MCP" and the new IDE therefore opened no project at all. So the path
    // is resolved (caller arg > registry projectPath > bounded re-read of the registry)
    // and a refusal names the exact parameter to pass.
    markHubOpStep(out, rec, 'resolve_project_path');
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
      return out;
    }
    out.project = out.projectVbp;
    markHubOpStep(out, rec, 'save_project');
    // 1) SAVE best effort first: a graceful quit must not throw work away.
    out.saveAttempted = true;
    try {
      const sv = await rpc(target.port, 'tools/call', { name: 'save_project', arguments: { path: out.project } }, null, 30000);
      out.saveOk = !!(sv && sv.result && !sv.result.isError);
    } catch (e) { out.saveOk = false; }
    out.steps.push({ step: 'save_project', attempted: true, ok: out.saveOk });
    markHubOpStep(out, rec, 'quit_ide');
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
    markHubOpStep(out, rec, 'stop_confirmed', gone ? 'stopped' : 'stop_unconfirmed');
    // 3) clear OUR OWN registry entry + the hub's bookkeeping.
    try {
      const rp = path.join(instancesDir(), target.pid + '.json');
      if (fs.existsSync(rp)) { fs.unlinkSync(rp); out.registryCleared = true; }
    } catch (e) { out.registryCleared = false; }
    startedByHub.delete(target.pid);
    invalidateDiscovery();
    out.lockCleared = false;
    markHubOpStep(out, rec, 'spawn', 'relaunching');
    // 4) restart on the SAME project.
    const childEnv = Object.assign({}, process.env, { VB6MCP_PORT: String(out.pinPort) });
    // (35) captured AT THE SPAWN POINT: no cwd option is passed, so the child inherits this hub
    // process's cwd; if it cannot be read the receipt says "unavailable" (never guessed).
    const spawnOpts = { detached: true, stdio: 'ignore', env: childEnv };
    const spawnCwd = hubSpawnCwd(spawnOpts);
    const spawnCwdBasis = spawnCwd === null ? 'unavailable' : 'spawn_cwd';
    // (FIX-26) the ONE launch that still used the BARE spawn: a .cmd/.bat VB6MCP_VB6EXE (a launcher
    // wrapper is a legitimate configuration, and the other two launches already go through
    // spawnProgram) could not be started at all - the op died as restart_error at step 'spawn'
    // instead of relaunching. Measured with a .cmd fake IDE. spawnProgram is byte-identical for an
    // .exe, so a real VB6.EXE takes the same path as before.
    const child = spawnProgram(out.vb6Exe, [out.project], spawnOpts);
    child.unref();
    // (35) a process WAS created by this operation: publish the spawn cwd.  A throw above leaves
    // the fail-closed null / "unavailable" from the receipt initializer untouched.
    out.cwd = spawnCwd; out.cwdBasis = spawnCwdBasis;
    out.newPid = child.pid;
    startedByHub.set(child.pid, { project: out.project, port: null });
    markHubOpStep(out, rec, 'wait_instance', 'relaunched');
    invalidateDiscovery();
    const t0 = Date.now();
    let newInst = null;
    let sawLowPortPid = 0;
    while (Date.now() - t0 < 120000) {
      await pause(1000);
      // (FIX-26 / F1) A RELAUNCHED IDE CAN COME UP PIPE-ONLY. The add-in DEFAULT transport is pipe
      // and this spawn inherits the HUB environment, so a hub started without VB6MCP_TRANSPORT=tcp
      // relaunches the IDE into a pipe-only instance (port 0, tcpEnabled false). The port test below
      // can never accept it, so the old loop burned the FULL 120 s and then reported
      // new_instance_never_came_up - false: the IDE WAS restarted and its own registration says so.
      const regNew = readRegistrationForPid(child.pid);
      if (pipeOnlyRegistration(regNew)) {
        // A pipe-only relaunch is a REAL restart on the PIPE track, not a failure.
        // Readiness is judged on the channel the instance ACTUALLY uses: the middle service
        // (mcp_info + ide_state.project).  The pinned TCP port does not exist for this
        // instance, so it is reported as not_applicable - never as "not ready".
        out.newInstanceFacts = newInstanceFactsOf(regNew, child.pid);
        out.restarted = true;          // the process WAS stopped and a NEW one was started
        out.port = 0;
        out.transport = 'pipe';
        out.pipe = regNew.pipe || null;
        out.pinPortUsable = false;
        // (37) expected PROJECT NAME = the NEW instance's OWN registration project (the add-in
        // writes it from the project VB6 actually opened, which may be rewritten: W44 measured
        // "_test_proj" -> "Vb_test_proj"); the .vbp basename is only a RECORDED fallback.
        const expectation = restartProjectExpectation(regNew, null, out.projectVbp);
        out.steps.push({ step: 'pipe_readiness', channel: 'middle_service', pid: child.pid,
                         wantName: expectation.name || null, wantNameSource: expectation.source });
        const rw = await waitForAddinReady(child.pid, { budgetMs: 60000, wantProjectName: expectation.name, hintPort: 0 });
        out.ready = !!rw.ready;
        out.readyWaitMs = rw.readyWaitMs;
        // (37) the project verdict is SAME-INSTANCE self-evidence: ide_state.project (what the
        // add-in says it opened) must EQUAL the registration.project of THAT SAME instance
        // (trim + case-insensitive); a missing side stays null (unverified), never a fake true.
        const pmPipe = projectNameEqual(rw.projectActual, expectation.name);
        out.readyChecks = Object.assign({}, rw.readyChecks, {
          projectMatches: pmPipe,
          projectMatchesExpected: expectation.name || null,
          projectMatchesActual: String(rw.projectActual || ''),
          projectMatchesBasis: 'ide_state.project vs new_instance_registration_project',
          wantNameSource: expectation.source
        });
        out.readyProbeRounds = rw.readyProbeRounds;
        out.readyProbeRule = rw.readyProbeRule;
        out.readyRegistrationFile = rw.readyRegistrationFile || null;
        out.readyBasis = rw.readyChecks.addInAnswersBasis || 'not_measured';
        out.threeWay = out.threeWay || { http: false, mcp_info: false, project: false };
        out.threeWay.http = 'not_applicable';
        out.threeWay.httpBasis = 'pipe_instance: the add-in registered NO TCP endpoint (port 0 by design); the pinned port ' + out.pinPort + ' is not a channel for this instance and is NOT used as a failure';
        out.threeWay.transport = 'pipe';
        out.threeWay.mcp_info = rw.readyChecks.addInAnswers === true ? true
                                  : (rw.readyChecks.addInAnswers === false ? false : null);
        out.threeWay.project = pmPipe === true ? true : (pmPipe === false ? false : null);
        out.threeWay.projectActual = String(rw.projectActual || '');
        out.threeWay.projectExpected = expectation.name || null;
        out.threeWay.projectMatchesBasis = 'ide_state.project vs new_instance_registration_project';
        out.threeWay.basis = {
          transport: 'pipe_registration(tcpEnabled=' + JSON.stringify(out.newInstanceFacts.tcpEnabled) + ', port=0)',
          mcpInfo: rw.readyChecks.addInAnswersBasis || 'not_measured',
          project: (pmPipe === null) ? 'unverified(no ide_state project available)' : 'ide_state.project == new_instance_registration_project'
        };
        out.three_way_ok = out.threeWay.mcp_info === true && out.threeWay.project === true;
        out.ok = out.restarted && out.three_way_ok;
        if (out.ok) {
          markHubOpStep(out, rec, 'done', 'ready');
          out.refused_reason = '';
          out.note = 'our own IDE (pid ' + out.stoppedPid + ') was stopped and RESTARTED on the PIPE track as pid ' +
            out.newPid + '; readiness was verified through the MIDDLE SERVICE (mcp_info answers + ide_state.project=' +
            out.threeWay.projectActual + '); the pinned TCP port ' + out.pinPort + ' is not_applicable for a pipe instance';
        } else {
          markHubOpStep(out, rec, 'pipe_not_ready', 'not_ready');
          out.refused_reason = 'new_instance_not_ready';
          out.message = 'the old process was stopped and a pipe-only instance WAS relaunched (pid ' + out.newPid +
            '), but it did not answer mcp_info through the middle service within the readiness budget';
          out.pollingHint = rw.pollingHint || '';
          out.howToProceed = rw.howToProceed || ['poll ide_list for pid ' + out.newPid];
          out.suggestedAction = 'retry_once';
          out.retrySafe = true;
          out.note = 'the reset DID happen (pipe instance registered); only its readiness was not proven - this is not a refused restart';
        }
        return out;
      }
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
      return out;
    }
    if (newInst) {
      out.port = newInst.port;
      out.transport = 'tcp';
      out.threeWay.transport = 'tcp';
      out.threeWay.http = true;
      out.threeWay.httpBasis = 'tcp_listening(port)';
      // The registry entry is written before the add-in serves and before the project is
      // loaded, so readiness is polled with a deadline instead of judged on first sight
      // (the first live run reported three_way_ok=false for exactly this race).
      // (37) expected PROJECT NAME = the NEW instance's own registration project (or its
      // discovery-probe name); the .vbp basename is only a RECORDED fallback - never the
      // pre-restart target's name and never out.project (a PATH).
      const expectation = restartProjectExpectation(readRegistrationForPid(child.pid), newInst, out.projectVbp);
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
        if (mcpOk) markHubOpStep(out, rec, 'wait_ready', 'addin_answering');
        out.threeWay.projectActual = actualProject;
        if (mcpOk && actualProject.length > 0) break;
        await pause(2000);
      }
      out.readyWaitMs = Date.now() - tReady;
      // (37) the project verdict is SAME-INSTANCE self-evidence: ide_state.project (what the
      // add-in says it opened) must EQUAL the NEW instance's registration/probe project name
      // (trim + case-insensitive).  The old code compared against target.project (the
      // PRE-restart instance) and even fell back to comparing a NAME with out.project (a PATH);
      // both are wrong authorities.  An absent side stays null (unverified) - never a fake true.
      const pmTcp = projectNameEqual(actualProject, expectation.name);
      out.threeWay.project = pmTcp === true ? true : (pmTcp === false ? false : null);
      out.threeWay.projectActual = String(actualProject || '');
      out.threeWay.projectExpected = expectation.name || null;
      out.threeWay.projectMatchesBasis = 'ide_state.project vs new_instance_registration_project';
      out.threeWay.wantNameSource = expectation.source;
      if (pmTcp === null && !String(actualProject || '').trim() && newInst.project) {
        // the add-in answered mcp_info but ide_state published no project name: quote the
        // registration's own value as the ACTUAL, keep the verdict UNVERIFIED (null).
        out.threeWay.projectActual = String(newInst.project);
      }
      // (W45 P3) the SAME project verdict travels in readyChecks on this track too (ADDITIVE:
      // threeWay keeps every field above unchanged). This only mirrors the already-decided
      // values - it does NOT re-verdict readiness (ready / three_way_ok keep their meaning).
      out.readyChecks = Object.assign({}, out.readyChecks || {}, {
        projectMatches: out.threeWay.project,
        projectMatchesExpected: expectation.name || null,
        projectMatchesActual: String(out.threeWay.projectActual || ''),
        projectMatchesBasis: 'ide_state.project vs new_instance_registration_project',
        wantNameSource: expectation.source
      });
    }
    out.three_way_ok = out.threeWay.http && out.threeWay.mcp_info && out.threeWay.project;
    markHubOpStep(out, rec, 'done', out.three_way_ok ? 'ready' : 'not_ready');
    out.restarted = !!(newInst && out.stopConfirmed);
    out.ok = out.restarted && out.three_way_ok;
    if (!out.ok && !out.refused_reason) { out.refused_reason = newInst ? 'three_way_not_ready' : 'new_instance_never_came_up'; }
    out.note = out.ok
      ? 'our own IDE (pid ' + out.stoppedPid + ') was stopped and restarted on ' + out.project + ' as pid ' + out.newPid + ' on port ' + out.port + '; the add-in answers, mcp_info answers and the project matches.'
      : 'not restarted: see refused_reason / steps / threeWay. Nothing foreign was touched and the 8766 instance was never a candidate.';
    return out;
}

// ---- schema shape on the WIRE ---------------------------------------------------------------
// Strict MCP clients validate the WHOLE tools/list result: ONE tool whose inputSchema.required is
// an EMPTY OBJECT ({} instead of []) makes them reject the entire result (INVALID_RESULT) and
// register ZERO tools of this server.  Measured 2026-10-20: six new tools carrying required:{}
// cost all 130 registrations.  Fix the SHAPE only, at the single exit every client sees; the
// inproc rollback path keeps its exact bytes.
const schemaShapeLogged = new Set();
function normalizeRequiredShape(tool) {
  const sch = (tool && typeof tool === 'object') ? tool.inputSchema : null;
  if (!sch || typeof sch !== 'object' || Array.isArray(sch)) return tool;
  if (!Object.prototype.hasOwnProperty.call(sch, 'required')) return tool;
  const req = sch.required;
  if (Array.isArray(req)) return tool;
  let out, why;
  if (typeof req === 'string') { out = [req]; why = 'string->array'; }
  else if (req && typeof req === 'object') { out = Object.keys(req).filter(function (k) { return !!req[k]; }); why = 'object->array'; }
  else { out = undefined; why = 'omitted'; }
  const copy = Object.assign({}, tool);
  const schCopy = Object.assign({}, sch);
  if (out === undefined) delete schCopy.required; else schCopy.required = out;
  copy.inputSchema = schCopy;
  const name = (tool && tool.name) ? tool.name : '?';
  const key = name + ':' + why;
  if (!schemaShapeLogged.has(key)) {
    schemaShapeLogged.add(key);
    try { process.stderr.write('[hub-schema] normalized ' + name + ' required=' + why +
      (out ? ' [' + out.join(', ') + ']' : '') + '\n'); } catch (e) { }
  }
  return copy;
}
function normalizeToolSchemas(tools) {
  if (BRIDGE_MODE === 'inproc') return tools;          // rollback path: byte-identical
  if (!Array.isArray(tools)) return tools;
  return tools.map(function (t) { return normalizeRequiredShape(t); });
}
async function toolsListResult(d, meta) {
  // ---- thin shell (b): the bridge owns the add-in face ---------------------------------
  // Bridge ONLINE  -> its 110-tool face + the 10 hub tools (each name once) = the SAME 120.
  // Bridge OFFLINE -> fall THROUGH to the pre-change path below, which publishes the COMPLETE
  // face from bridge/hub_schema.generated.json (110) + hub_tools.json (10): a dead Python
  // layer must never cost the caller the tool list.
  if (BRIDGE_MODE !== 'inproc') {
    const h = await bridgeHealth();
    if (h) {
      try {
        const tl = await bridgeRpc('tools/list', {}, BRIDGE_LIST_TIMEOUT_MS);
        const tools = (tl && tl.result && tl.result.tools) || null;
        if (tools && tools.length) {
          dbg('bridge tools/list tools=' + tools.length);
          if (meta) { meta.source = 'bridge'; meta.toolCount = tools.length; }
          return { tools: normalizeToolSchemas(mergeHubFace(tools)) };     // strict: EXACTLY {tools}
        }
      } catch (e) { dbg('bridge tools/list FAILED: ' + String((e && e.message) || e)); }
    }
    dbg('bridge offline: serving the complete face from ' + SNAP + ' + ' + HUBCFG);
  }
  if (meta) meta.source = 'fallback';
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
  return { tools: normalizeToolSchemas(merged) };
}

// ---- failureKind: the machine-readable distinction of a forward failure ----------------------
// See FAILURE_KIND_RULE.  Everything here is ADDITIVE: reason / ok / message keep their values.
function transportFailureKind(e) {
  if (e && (e.name === 'TimeoutError' || e.name === 'AbortError')) return 'wedged';
  const code = String((e && (e.code || (e.cause && e.cause.code))) || '').toUpperCase();
  const msg = String((e && (e.message || e)) || '');
  if (/timed?\s?out|abort/i.test(msg)) return 'wedged';
  // Only a REFUSED/failed lookup proves there is nobody there.  A reset / EPIPE / closed socket
  // proves the opposite: something accepted the connection and then went away -> wedged.
  if (code === 'ECONNREFUSED' || code === 'ENOTFOUND' || code === 'EAI_AGAIN' || code === 'EHOSTUNREACH' || code === 'ENETUNREACH' || code === 'ERR_INVALID_URL') return 'noIde';
  return 'wedged';
}
// A refusal means no candidate could be ROUTED to.  When every discovered candidate is a
// failed probe (registry alive + port listening, but the add-in never answered) this is not
// "there is no IDE" - it is the wedge /-20 described, and it is reported as wedged.
function refusalFailureKind(insts) {
  const live = (insts || []).filter(function (i) { return !!i && i.pending !== true; });
  if (live.length && live.every(function (i) { return i.probeStatus === 'failed' || !!i.probeReason; })) return 'wedged';
  return 'noIde';
}
function firstBodyObject(result) {
  const c = result && result.content;
  const txt = (Array.isArray(c) && c[0] && typeof c[0].text === 'string') ? c[0].text : null;
  if (!txt) return null;
  try { const j = JSON.parse(txt); return (j && typeof j === 'object' && !Array.isArray(j)) ? j : null; } catch (e) { return null; }
}
// Did the add-in ANSWER with a structured error?  Only then is the host known to be alive.
function answerFailureKind(resp) {
  if (!resp || typeof resp !== 'object') return null;
  if (resp.error) return 'toolError';
  const r = resp.result;
  if (!r || typeof r !== 'object') return null;
  if (r.isError === true) return 'toolError';
  const obj = firstBodyObject(r);
  if (obj && obj.ok === false) return 'toolError';
  return null;
}
// The refusal body keeps EVERY legacy key (ok/reason/message/instances/conflicts) in the same
// order and ADDS the candidate detail the path key needs (candidates[], identityConflict).
// /N14: candidates[] + suggestion + resolution are now emitted on EVERY refusal (not only
// on a multi-match or when remote endpoints exist): a refusal must say WHICH instances exist and
// HOW to address one, otherwise "ambiguous_target" is a dead end instead of a predictable rule.
function refusalBody(pick, insts, conflicts) {
  const reason = pick.conflict ? 'target_conflict' : (pick.mismatch ? 'target_mismatch' : (pick.ambiguous ? 'ambiguous_target' : 'no_ide_running'));
  const cands = (pick.hits && pick.hits.length) ? pick.hits : (insts || []).map(candidateBrief);
  const body = {
    ok: false, reason: reason, message: pick.message,
    instances: (insts || []).map(function (i) { return { pid: i.pid, port: i.port, project: i.project }; }),
    conflicts: conflicts,
    candidates: pick.candidates || cands,
    suggestion: pick.suggestion || ((insts || []).length ? suggestionFor(insts) : null),
    resolution: pick.resolution || resolutionOf('refused', insts, { requested: null, explicit: false }),
    targetResolutionRule: TARGET_RESOLUTION_RULE
  };
  const showIdentity = REMOTE_ENDPOINTS.length > 0 || !!pick.conflict;
  if (showIdentity) {
    if (pick.identityConflict !== undefined) body.identityConflict = !!pick.identityConflict;
    if (pick.distinctIdentities !== undefined) body.distinctIdentities = pick.distinctIdentities;
    if (pick.matchedCount !== undefined) body.matchedCount = pick.matchedCount;
  }
  if (REMOTE_ENDPOINTS.length) { body.endpoints = endpointLabels(); body.targetKeyRule = TARGET_KEY_RULE; }
  body.failureKind = refusalFailureKind(insts);
  body.failureKindRule = FAILURE_KIND_RULE;
  if (MULTI_BRIDGE) {
    // Same vocabulary as the WP-4-07 failure envelope: an ambiguous/conflicting target needs its
    // PARAMETERS fixed, not a retry; a missing IDE needs a start.
    body.suggestedAction = (reason === 'ambiguous_target' || reason === 'target_conflict' || reason === 'target_mismatch')
      ? 'fix_params' : 'ide_start';
    body.retrySafe = false;
    // failureKind must agree with suggestedAction: an ambiguity/conflict is a CALLER-PARAMETER
    // problem (toolError), not "there is no IDE" (noIde).
    if (reason === 'ambiguous_target' || reason === 'target_conflict' || reason === 'target_mismatch') {
      body.failureKind = 'toolError';
    }
    if (reason === 'ambiguous_target') {
      body.candidatesRule = 'each candidate names the bridge it lives on (bridge) and its canonical targetKey; pass _target="<bridge>:<pid|project|projectPath>" to pick one';
    }
  }
  // (WP-22-01) a target failure must be DIAGNOSABLE: which pids ARE online, why the selector
  // failed, and the one action that fixes it.  target_mismatch vs no_live_ide are distinguished
  // by the LIVE SET, never by a pid the hub used to know.
  body.onlinePids = onlinePidsOf(insts);
  const diagReason = (reason === 'no_ide_running' || reason === 'no_target') ? 'no_live_ide'
    : ((reason === 'ambiguous_target' || reason === 'target_conflict') ? reason : 'explicit_target_not_live');
  body.targetDiagnostic = targetDiagnosticBlock({
    reason: diagReason,
    onlinePids: body.onlinePids,
    requested: (pick && pick.matchedFrom) ? pick.matchedFrom : ((pick && pick.resolution && pick.resolution.requested) || null),
    explicit: !!(pick && pick.resolution && pick.resolution.explicit)
  });
  if (reason === 'no_ide_running' || reason === 'no_target') {
    body.reason = 'no_live_ide';
    body.reasonLegacy = reason;
    body.rule = NO_LIVE_IDE_RULE;
    body.failureKind = 'noIde';
    body.suggestedAction = 'ide_start';
    body.retrySafe = false;
  }
  if (!body.howToProceed) body.howToProceed = offlineHintList();
  if (!body.suggestion) {
    body.suggestion = body.onlinePids.length
      ? suggestionFor(insts)
      : { action: 'ide_start', useTarget: null, ide_start: { project: '<absolute path to the .vbp>' },
          note: 'there is no live IDE: run ide_start{project:"<absolute .vbp>"}, or use build / make_project(route:"cli") / project_meta which need no IDE at all' };
  }
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
    targetKeyRule: TARGET_KEY_RULE,
    failureKind: 'noIde',
    failureKindRule: FAILURE_KIND_RULE
  };
}
// The ONLY forward path: local goes to 127.0.0.1:<port> over the same wire the hub always used;
// remote goes to the endpoint URL.  Per-endpoint timeout + per-endpoint probe single-flight
// means a slow remote can never stall a local call.
async function forwardPicked(pick, name, args, snap, explicit, rawRequest, ladderBlock, implicitFwd, decidedBy) {
  // r-c (PROTOCOL SYMMETRY): a lesson method must NEVER be served by an endpoint.  Normal
  // dispatch serves them at the control plane before routing is considered, so reaching this
  // point is a protocol violation - fail STRUCTURALLY and never create a per-machine copy of
  // the store (that is the rule that prevents knowledge divergence between machines).
  if (lessonCliArgs(name, args)) {
    return stampBody({ content: [{ type: 'text', text: JSON.stringify({
      ok: false, reason: 'global_state_is_control_plane_only', tool: name,
      endpoint: pick.endpoint || (pick.remote ? 'remote' : LOCAL_ENDPOINT), remote: !!pick.remote,
      failureKind: 'toolError', failureKindRule: FAILURE_KIND_RULE,
      message: 'lesson tools read/write the SINGLE control-plane store (mcp_lessons/); an endpoint must never serve them and must never create a local copy of that store'
    }, null, 2) }], isError: true }, null, snap.contractVersion);
  }
  const pickNote = {};
  const midRec = middleServiceReceipt();
  if (midRec) pickNote.middleService = midRec;   // every forwarded receipt names its addresser
  const hcPick = hubCodeStamps();               // P1: same code-proof on the direct path
  if (hcPick) pickNote.hubCode = hcPick;
  if (pick.selectedBy) pickNote.selectedBy = pick.selectedBy;
  // the budget a forwarded call may spend, and the probe state of the instance it was given to
  pickNote.forwardBudget = forwardBudgetBlock(!!implicitFwd, implicitFwd ? implicitForwardCapMs() : 0, decidedBy);
  if (implicitFwd) { pickNote.implicitForward = true; pickNote.implicitForwardCapMs = implicitForwardCapMs();
                     if (decidedBy) pickNote.implicitForwardBasis = decidedBy; }
  if (probePendingOf(pick)) pickNote.probePending = true;
  // when the LADDER's empty-face fallback supplied the candidate, the receipt says so
  // (the read sources ride in ladderLiveFallback.sources; the probe state in probePending)
  if (ladderBlock && ladderBlock.used) pickNote.selectedBy = 'ladder_live_fallback';
  if (pick.staleTarget) {
    pickNote.staleTarget = pick.staleTarget;
    pickNote.staleTargetIgnored = true;
    pickNote.targetNote = 'the default ide_target "' + pick.staleTarget + '" points at no live instance (treated as unset, NOT rewritten); selected the only live instance ' + pick.port + '/' + (pick.project || '?') + ' -- clear it with ide_target{auto} if it is stale';
  }
  const isRemote = !!pick.remote;
// /N11: the SAME verdict block ide_list reports (liveness + its basis, ownership,
  // process kind, resolution) rides on the forwarded receipt too, so a noIde/wedged answer from a
  // forwarded tool can never be read as contradicting ide_list - both name one fact source.
  const hv = hubVerdictNote(pick, rawRequest);
  if (hv) { for (const k of Object.keys(hv)) { if (pickNote[k] === undefined) pickNote[k] = hv[k]; } }
  // : the SAME attribution block ide_list publishes, on the forwarded diagnostic receipt
  // (stampBody only fills keys the add-in did not set, so an add-in envelope is never overwritten).
  const ovNote = await ownershipNoteFor(pick, name);
  for (const k of Object.keys(ovNote)) { if (pickNote[k] === undefined) pickNote[k] = ovNote[k]; }
  if (isRemote) { pickNote.remote = true; pickNote.endpoint = pick.endpoint; pickNote.endpointUrl = pick.endpointUrl; }
  if (pick.matchedBy === 'project_path') {
    pickNote.targetKey = pick.targetKey;
    pickNote.targetNormalized = { from: pick.matchedFrom || rawRequest || null, to: pick.projectPath, endpoint: pick.endpoint || LOCAL_ENDPOINT, via: pick.pathVia || null, rule: TARGET_KEY_RULE };
  }
  const ideStamp = { pid: pick.pid, project: pick.project, port: pick.port };
  if (isRemote) {
    ideStamp.endpoint = pick.endpoint; ideStamp.remote = true;
    ideStamp.bridge = pick.bridge || pick.endpoint;
  } else if (MULTI_BRIDGE) {
    ideStamp.bridge = pick.bridge || LOCAL_ENDPOINT;
    ideStamp.endpoint = LOCAL_ENDPOINT;
    ideStamp.remote = false;
  }
  const baseTimeoutMs = isRemote ? REMOTE_FORWARD_TIMEOUT_MS : FORWARD_TIMEOUT_MS;
  // (g1) the implicit cap: an unqualified call whose instance has not answered a probe must
  // answer inside the CALLER's patience; an explicit _target keeps the long budget.
  const timeoutMs = implicitFwd ? Math.min(baseTimeoutMs, implicitForwardCapMs()) : baseTimeoutMs;
  // ONE decision about what this pick can be reached over.  A pipe-only pick has NO direct
  // endpoint: it is the middle service's job (or nothing), and a URL built from its port would
  // be "...:null/rpc" - the parse failure that hid the real cause.  Refuse structurally instead.
  const fe = directEndpointOf(pick);
  if (fe.kind === 'pipeOnly') {
    dbg('forward REFUSED: pipe-only instance pid=' + pick.pid + ' port=' + pick.port + ' -> the middle service is required at ' + BRIDGE_URL);
    return pipeOnlyRefusal(pick, name, snap, pickNote);
  }
  if (!fe.base) {
    const nb = { ok: false, tool: name, reason: 'no_forward_route',
      failureKind: isRemote ? 'noIde' : 'pipeOnlyNeedsMiddleService', failureKindRule: FAILURE_KIND_RULE,
      remote: !!isRemote, endpoint: isRemote ? pick.endpoint : null, endpointUrl: isRemote ? (pick.endpointUrl || null) : null,
      pid: pick.pid || null, port: (pick.port === undefined) ? null : pick.port,
      forwardRoute: forwardRouteOf(pick), probeRoute: pick.probeRoute || null,
      bridgeUrl: isRemote ? null : BRIDGE_URL, bridgeMode: BRIDGE_MODE,
      message: isRemote ? 'the remote endpoint has no URL, so nothing was forwarded' : 'the selected instance has no usable forward route, so nothing was forwarded',
      nothingForwarded: true };
    const nbg = nullPortUrlGuard(nb, 'no_forward_route(' + String(name) + ')');
    return stampBody({ content: [{ type: 'text', text: JSON.stringify(nbg || nb, null, 2) }], isError: true }, ideStamp, snap.contractVersion, pickNote);
  }
  const base = fe.base;
  const tFwd = Date.now();
  let resp;
  try { resp = await rpcAt(base, 'tools/call', { name: name, arguments: args }, null, timeoutMs); }
  catch (e) {
    const ms = Date.now() - tFwd;
    const isTimeout = !!(e && (e.name === 'TimeoutError' || /abort/i.test(String(e && e.message))));
    const errBody = { ok: false, reason: isTimeout ? 'forward_timeout' : 'forward_failed', message: String(e && e.message || e), elapsedMs: ms, budgetMs: timeoutMs,
                      forwardBudget: forwardBudgetBlock(!!implicitFwd, implicitFwd ? implicitForwardCapMs() : 0, decidedBy), retrySafe: true,
                      ...(implicitFwd ? { implicitForward: true, implicitForwardCapMs: implicitForwardCapMs() } : {}),
                      ...(probePendingOf(pick) ? { probePending: true } : {}),
                      failureKind: transportFailureKind(e), failureKindRule: FAILURE_KIND_RULE,
                      forwardRoute: forwardRouteOf(pick), probeRoute: pick.probeRoute || null,
                      port: fe.port, bridgeUrl: isRemote ? null : BRIDGE_URL, bridgeMode: BRIDGE_MODE,
                      registry: isRemote ? null : registrySnapshotFor(pick.pid),
                      liveness: pick.liveness || null };
    if (isRemote) { errBody.endpoint = pick.endpoint; errBody.endpointUrl = pick.endpointUrl; errBody.remote = true; }
    else if (isTimeout) {
      errBody.hint = 'raise VB6MCP_FORWARD_TIMEOUT_MS for legitimately long calls';
      errBody.howToProceed = [
        'address the instance explicitly: _target="<pid|projectPath>" (a pinned forward skips the ladder and its probe state)',
        'probe liveness first with ide_state (or ide_state{release_inflight:true} when the add-in is wedged)',
        'do NOT simply repeat the same implicit call: it will spend the same budgetMs again with the same result'
      ];
      errBody.retryNote = 'retrySafe:true means a retry cannot corrupt state; it is NOT the fix - the peer did not answer within ' + timeoutMs + ' ms';
    }
    const guardedErr = nullPortUrlGuard(errBody, 'forward_failed(' + String(name) + ')');
    return stampBody({ content: [{ type: 'text', text: JSON.stringify(guardedErr || errBody, null, 2) }], isError: true }, ideStamp, snap.contractVersion, pickNote);
  }
  dbg('forward phases: rpc=' + (Date.now() - tFwd) + 'ms tool=' + name + ' ' + (isRemote ? 'endpoint=' + pick.endpoint : 'port=' + pick.port));
  // A body-shaped result wins; otherwise an error.data envelope (older peer shape) is lifted so
  // the caller still sees the envelope; only then the raw response is shown verbatim.
  const result = (resp && resp.result) || errorEnvelopeOf(resp) ||
    { content: [{ type: 'text', text: JSON.stringify(resp) }] };
  // The add-in ANSWERED.  If that answer is a structured error it is toolError (the host is
  // alive; the tool/params are the story) - injected additively, and NEVER over an envelope the
  // add-in already classified itself (bodyOf only adds keys that are undefined).
  const answeredKind = answerFailureKind(resp);
  const note = answeredKind ? Object.assign({}, pickNote, { failureKind: answeredKind, failureKindRule: FAILURE_KIND_RULE }) : pickNote;
  return stampBody(result, ideStamp, snap.contractVersion, note);
}

// ---- ONE local resolution path ---------------------------------------------------
// discovery -> directTarget escape -> the documented ladder.  Used by BOTH transports, so the
// middle service and the direct path can never apply two different default-target rules (that
// divergence is the "sometimes it works, sometimes ambiguous_target" of N14).
// ---- (WP-22-01 / t16) "I DO NOT KNOW YET" is NOT "NOTHING IS LIVE" -------------------------
// discoverSoft() gives up after LIST_GRACE_MS (400 ms) so a response never waits long, while ONE
// discovery pass legitimately takes seconds: it reads the registry, asks the middle service
// /health and probes every port-scan candidate (a probe can take 1.5 s, a pipe probe 4.5 s).
// MEASURED on t16 (hub f8dab602): run_project{observe:false} answered "no IDE instance listening
// (scanned 34 default ports + registry)" TWICE, then ide_state{} answered from pid 9252 241 ms
// later, and every explicit _target call worked throughout.  The IDE was alive the whole time:
// the two calls landed while the ONE discovery pass was still running, and the fallback sources
// the ladder used (the persisted face, then the hub's own registry) were blind on that machine.
// The ladder then said "nothing is live" on the strength of an EMPTY CACHE.
//
// Both tools take the SAME path (handleCall -> resolveLocalPick -> pickTarget); run_project has no
// hub-side branch (grep: 0 hits).  So this is a TIMING property of the ladder, not a per-tool one,
// and it is fixed HERE, once: when the face is EMPTY the ladder does two bounded things before it
// may answer no_candidate -
//   (a) WAIT for the discovery pass that is ALREADY RUNNING (bounded by LADDER_EMPTY_FACE_WAIT_MS),
//   (b) still empty -> read the DIRECT sources a cache cannot represent: the hub's registry files
//       from disk right now, then ONE forced fresh pass (registry + middle-service /health + port
//       scan).  What happened is reported in resolution.ladderLiveFallback{sources,waitedMs,...}.
const LADDER_EMPTY_FACE_WAIT_MS = Math.max(200, parseInt(process.env.VB6MCP_LADDER_EMPTY_FACE_WAIT_MS || '2500', 10) || 2500);
const LADDER_LIVE_FALLBACK_RULE = 'an EMPTY discovery face is NOT evidence that nothing is live: discoverSoft() only waits 400 ms for a pass that can take seconds, and a cached empty face stays cached for the discovery TTL. Before answering no_candidate the default-target ladder (a) waits, bounded, for a discovery pass that is already running, then (b) reads the DIRECT sources a cache cannot represent - the hub registry files from disk RIGHT NOW, then ONE forced fresh pass (registry + middle-service /health + port scan) - and reports what it did in ladderLiveFallback{used,sources,waitedMs,elapsedMs,budgetMs,complete,timedOut,rule}. ONE TOTAL BUDGET (LADDER_EMPTY_FACE_WAIT_MS) covers ALL of those stages: when it is exhausted the fallback RETURNS with complete:false + timedOut:true (used:false if nothing was found) instead of waiting for a pass that may never finish. An EXPLICIT _target is never affected by this fallback; it is handed down verbatim.';
async function ladderLiveFallback(current) {
  const t0 = Date.now();
  const budget = LADDER_EMPTY_FACE_WAIT_MS;          // ONE total budget for ALL stages below
  const deadline = t0 + budget;
  const left = function () { return Math.max(0, deadline - Date.now()); };
  const out = { sources: [], instances: [], waitedMs: 0, elapsedMs: 0, budgetMs: budget, timedOut: false, complete: false };
  const done = function () { out.elapsedMs = Date.now() - t0; out.complete = !out.timedOut; return out; };
  // (a) a pass that is ALREADY running will answer this question: wait for it, bounded by the
  //     REMAINING total budget (its own wait never gets a fresh allowance).
  const inflight = discovery.promise;
  if (inflight) {
    const tW = Date.now();
    let d2 = null;
    try { d2 = await withTimeout(inflight, Math.max(1, left())); } catch (e) { d2 = null; out.timedOut = true; }
    out.waitedMs = Date.now() - tW;
    out.sources.push(d2 ? 'in_flight_discovery_pass' : 'in_flight_discovery_pass:timed_out');
    const rd2 = d2 ? await resolveDiscovery(d2) : null;
    if (rd2 && rd2.instances && rd2.instances.length) { out.instances = rd2.instances; return done(); }
    if (left() <= 0) { out.sources.push('budget_exhausted:after_in_flight'); return done(); }
  } else out.sources.push('in_flight_discovery_pass:none');
  // (b1) the hub's OWN registration files, read from disk NOW (no face cache, no TTL).
  let rc = null;
  try { rc = await withTimeout(registryCandidates(), Math.max(1, left())); } catch (e) { rc = null; out.timedOut = true; }
  const rcr = rc ? await resolveDiscovery(rc) : null;
  if (rcr && rcr.instances && rcr.instances.length) {
    out.instances = rcr.instances; out.sources.push('direct_registry_read'); return done();
  }
  out.sources.push('direct_registry_read:no_live_instance');
  if (left() <= 0) { out.sources.push('budget_exhausted:after_registry_read'); return done(); }
  // (b2) ONE forced fresh pass.  This is the ONLY thing that can see an IDE the hub's own registry
  //      does not contain (port scan) or one only the middle service can see (/health adoption).
  //      The cached empty face is invalidated first, otherwise discover() would just hand it back.
  //      The pass itself CANNOT be cancelled, so we stop WAITING at the deadline and report
  //      complete:false - a pass that never finishes must never hold the tool call.
  invalidateDiscovery();
  let fresh = null;
  try { fresh = await withTimeout(discover(), Math.max(1, left())); } catch (e) { fresh = null; out.timedOut = true; }
  const fr = fresh ? await resolveDiscovery(fresh) : null;
  if (fr && fr.instances && fr.instances.length) {
    out.instances = fr.instances; out.sources.push('forced_fresh_discovery_pass');
  } else out.sources.push('forced_fresh_discovery_pass:timed_out_or_empty');
  return done();
}
function ladderFallbackBlock(out) {
  if (!out) return null;
  return { used: !!(out.instances && out.instances.length), sources: out.sources || [],
           waitedMs: out.waitedMs || 0, elapsedMs: out.elapsedMs || 0,
           budgetMs: out.budgetMs || LADDER_EMPTY_FACE_WAIT_MS,
           complete: out.complete === true, timedOut: out.timedOut === true,
           rule: LADDER_LIVE_FALLBACK_RULE };
}
// The fallback's evidence must be on EVERY receipt that travelled the empty-face path, whatever
// the envelope looks like.  It is attached as a TOP-LEVEL key (mergeNote fills only ABSENT
// top-level keys, so a body assembled downstream cannot shadow it) AND inside resolution when a
// resolution block exists.
function withLadderFallback(body, localRes) {
  const blk = (localRes && localRes.ladderFallback) ? localRes.ladderFallback : null;
  if (!blk || !body || typeof body !== 'object') return body;
  if (body.ladderLiveFallback === undefined) body.ladderLiveFallback = blk;
  if (body.resolution && typeof body.resolution === 'object' && body.resolution.ladderLiveFallback === undefined) {
    body.resolution.ladderLiveFallback = blk;
  }
  return body;
}
async function resolveLocalPick(tgt, explicit, rawRequest) {
  let localD0 = await resolveDiscovery(await discoverSoft(LIST_GRACE_MS));
  // (t16) an EMPTY face is not proof that nothing is live - see LADDER_LIVE_FALLBACK_RULE.
  let ladderFallback = null;
  if (!(localD0 && localD0.instances && localD0.instances.length)) {
    ladderFallback = await ladderLiveFallback(localD0);
    if (ladderFallback.instances.length) {
      localD0 = Object.assign({}, localD0 || {}, { instances: ladderFallback.instances,
                                                    conflicts: (localD0 && localD0.conflicts) || [] });
      invalidateDiscovery();   // the cached face was blind: never keep serving it
    }
  }
  // (WP-22-01) LIVE-ONLY reconciliation: a candidate carried by a persisted face, a recycled pid
  // or a registration whose transport is gone must NEVER become the default target.  Dropping it
  // BEFORE pickTarget is what stops a dead pid from being handed down to the middle service as a
  // canonical selector (the reported target_mismatch for a pid that had already exited).
  const rec = await reconcileLiveInstances((localD0 && localD0.instances) || []);
  const localD = Object.assign({}, localD0, {
    instances: rec.instances,
    droppedDeadCandidates: ((localD0 && localD0.droppedDeadCandidates) || []).concat(rec.dropped),
    liveReconciled: true
  });
  const tRes = Date.now();
  const popts = { remoteScope: false, explicit: explicit, requested: rawRequest };
  let merged = null;
  let pick = null;
  let remoteOnly = null;
  if (tgt.kind === 'endpoint') {                       // "local:<selector>"
    pick = pickTarget(localD.instances, tgt.selector || null, popts);
  } else if (MULTI_BRIDGE && (tgt.kind === 'auto' || tgt.kind === 'local')) {
    // MULTI-BRIDGE ONLY: `auto` and a BARE selector are resolved over the FULL candidate set
    // (local + every configured bridge), because only the merged set can see that one pid exists
    // on two bridges.  Single-bridge mode never enters this branch (byte-identical path).
    merged = await withRemoteCandidates(localD, REMOTE_GRACE_MS);
    const sel = (tgt.kind === 'local') ? tgt.selector : null;
    pick = pickTarget(merged.instances, sel, popts);
    // A match spanning DIFFERENT bridges is AMBIGUITY, not a one-endpoint conflict (VB6's
    // "same .vbp twice" rule is per machine).  It is named as such so a caller can pick a side.
    if (pick && !pick.ok && pick.hits && pick.hits.length > 1) {
      const eps = {};
      pick.hits.forEach(function (h) { eps[h.endpoint || LOCAL_ENDPOINT] = 1; });
      if (Object.keys(eps).length > 1) {
        pick = { ok: false, ambiguous: true, conflict: false,
                 candidates: pick.hits, hits: pick.hits,
                 suggestion: pick.suggestion || suggestionFor(pick.hits),
                 resolution: resolutionOf('refuse_ambiguous_target', merged.instances, popts,
                   { matchedCount: pick.hits.length, distinctBridges: Object.keys(eps) }),
                 message: 'target "' + sel + '" matches ' + pick.hits.length + ' live IDEs on DIFFERENT bridges (' +
                   pick.hits.map(function (h) { return (h.endpoint || LOCAL_ENDPOINT) + ':' + h.pid + (h.project ? ' ' + h.project : ''); }).join(' | ') +
                   ') -- address ONE explicitly with <bridge>:<pid|project|projectPath>' };
      }
    }
    const remoteLiveM = merged.instances.filter(function (i) { return !!i.remote; });
    if (pick && !pick.ok && !pick.ambiguous && !pick.mismatch && remoteLiveM.length) {
      remoteOnly = { remoteLive: remoteLiveM, merged: merged };
    }
  } else if (tgt.kind === 'local') {                   // bare selector (path | pid | port | name)
    pick = pickTarget(localD.instances, tgt.selector, popts);
  } else if (localD.instances.length) {                // auto, LOCAL FIRST
    pick = pickTarget(localD.instances, null, popts);
  } else if (REMOTE_ENDPOINTS.length) {                // auto with ZERO local IDEs: LOCAL ONLY, never remote
    // P1b (author ruling): an UNQUALIFIED call may only ever pick a LOCAL IDE.  The pre-P1b rule
    // 6r ("else the ONLY live IDE overall") auto-targeted a lone REMOTE IDE - that silent machine
    // jump is what this refuses.  The remote face is still PROBED here, but only so the refusal
    // can name the endpoint-qualified keys the caller must use; nothing is routed anywhere.
    merged = await withRemoteCandidates(localD, REMOTE_GRACE_MS);
    const remoteLive = merged.instances.filter(function (i) { return !!i.remote; });
    if (remoteLive.length) remoteOnly = { remoteLive: remoteLive, merged: merged };
    else pick = pickTarget(merged.instances, null, popts);
  } else {
    pick = pickTarget(localD.instances, null, popts);
  }
  // The direct-target escape hatch ("discovery is blind but the add-in is reachable") is NOT
  // allowed to rescue a CONFLICT: a selector that matches several live IDEs on one endpoint
  // must be reported, never silently resolved to the first registry hit (that is the exact
  // silent pick the path key exists to prevent).
  if (pick && !pick.ok && !pick.conflict) {
    const want = explicit ? (tgt.selector || null) : (state.target || null);
    const dt = want ? await directTarget(want) : null;
    if (dt) {
      dbg('resolveLocalPick: directTarget bypass used for ' + want + ' -> port ' + dt.port);
      pick = Object.assign({}, dt, { resolution: resolutionOf('direct_target_escape', localD.instances, popts, { selectedBy: 'direct_target', matchedFrom: want }) });
    }
  }
  // the ladder may have had to read a DIRECT source (see LADDER_LIVE_FALLBACK_RULE): say so
  // in the resolution block, so a receipt shows WHY it did not trust the empty face.
  if (ladderFallback && pick && pick.resolution) pick.resolution.ladderLiveFallback = ladderFallbackBlock(ladderFallback);
  return { localD: localD, merged: merged, pick: pick, remoteOnly: remoteOnly, elapsedMs: Date.now() - tRes,
           ladderFallback: ladderFallback ? ladderFallbackBlock(ladderFallback) : null };
}
// The hub's verdict about the resolved instance, attached to a forwarded receipt so a
// noIde/wedged answer is never bare: WHICH instance, its liveness + basis, and how it was chosen.
function hubVerdictNote(pick, rawRequest, step) {
  if (!pick) return null;
  const live = pick.liveness || (pick.probeStatus ? livenessOf(pick, bridgeIdentitySync()) : null);
  // : hubInstance.ownership used to be pick.ownership, which is UNDEFINED for every pick
  // the hub resolves itself - so the forwarded receipt printed "no ownership" next to ide_list's
  // "registry_started_by_hub" for the SAME pid.  It is now the same predicate ide_list publishes.
  let legacyOwn = pick.ownership || null;
  if (!legacyOwn) { try { legacyOwn = ownershipOf(pick).ownership; } catch (e) { legacyOwn = null; } }
  return { hubInstance: { pid: pick.pid || null, port: pick.port || null, project: pick.project || null,
                          projectPath: pick.projectPath || null, targetKey: pick.targetKey || null,
                          processKind: pick.processKind || 'vb6_ide', ownership: legacyOwn },
           hubLiveness: live ? { state: live.state, basis: live.basis, probeStatus: live.probeStatus, forwardRoute: live.forwardRoute, pipeState: live.pipeState } : null,
           livenessRule: LIVENESS_RULE,
           targetResolution: pick.resolution || null, targetRequested: rawRequest, targetResolutionRule: TARGET_RESOLUTION_RULE,
           targetNote: step || null };
}
// The attribution block for the forwarded DIAGNOSTIC entrance.  ide_state is
// the tool whose answer was read as a second opinion about the SAME instance ide_list listed, so it
// carries the SAME ownershipVerdict block, computed by the SAME resolver, from the same facts.
async function ownershipNoteFor(pick, name) {
  if (name !== 'ide_state' || !pick || !pick.pid) return {};
  let blk = null;
  try { blk = await resolveOwnership(pick); } catch (e) { blk = null; }
  return blk ? { ownershipVerdict: blk, ownershipVerdictRule: OWNERSHIP_VERDICT_RULE } : {};
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
      failureKind: 'noIde', failureKindRule: FAILURE_KIND_RULE,
      message: 'endpoint "' + tgt.label + '" is not configured; configured endpoints: ' + endpointLabels().join(', ')
    }, null, 2) }], isError: false }, null, snap.contractVersion);
  }

  // ---- thin shell (c): a LOCAL / unqualified call is handed to the bridge ------------------
//, three EXPLICIT rules (reviewer-confirmed):
  //   * EXPLICIT _target  -> handed down VERBATIM (hubTargetNotRewritten=true).  The hub never
  //     rewrites a caller's selector; it only ATTACHES its own verdict/resolution as diagnostics,
  //     and the middle service stays the resolver, exactly as before this change.
  //   * OMITTED / auto    -> the HUB resolves it with the documented ladder and hands down a
  //     canonical selector (pid preferred).  If the ladder is not decisive the hub REFUSES with
  //     candidates[] + suggestion{} instead of letting a second rule guess.
  //   * HUB BLIND (0 add-in instances visible) -> the caller's args go down verbatim with
  //     targetResolution.step=blind_delegation, because delegation is the only route left.
  // If the middle service REFUSES the target/selector, that refusal is surfaced with
  // noSilentFallback=true: no other form is tried and no instance is picked by the hub itself.
  const localScope = !(tgt.kind === 'endpoint' && tgt.endpoint && tgt.endpoint.remote);
  const localRes = localScope ? await resolveLocalPick(tgt, explicit, rawRequest) : null;
  // (g1) THE IMPLICIT CAP: a call that did NOT name _target and whose instance has not answered a
  // probe (or was selected by the ladder's empty-face fallback) is capped by
  // VB6MCP_IMPLICIT_FORWARD_TIMEOUT_MS so the hub ANSWERS inside the caller's patience.  An
  // EXPLICIT _target keeps the long budget (that is the wedged-IDE release_inflight channel).
  // (g1, t16 live) EVERY unqualified call is capped: a fresh hub's FIRST call picks the healthy
  // instance through the normal ladder (probePending false, no fallback), so gating the cap on
  // those inputs left exactly the reported case on the 130 s budget.  An implicit target is
  // cheap to retry; an EXPLICIT _target keeps the long budget (wedged-IDE release_inflight).
  const implicitFwd = !explicit;
  const fwdDecidedBy = { explicit: !!explicit, probePending: probePendingOf(localRes && localRes.pick),
                         ladderFallbackUsed: !!(localRes && localRes.ladderFallback && localRes.ladderFallback.used),
                         capApplied: !explicit, capMs: implicitFwd ? implicitForwardCapMs() : 0, requested: rawRequest };

  // (B7) an EXPLICIT selector whose only match is a DEAD registration is peer_dead, not
  // no_live_ide: the caller must know which pid to restart vs whether to start an IDE at all.
  if (explicit && localScope && localRes && localRes.pick && !localRes.pick.ok) {
    const pd = await peerDeadBody(name, rawRequest, localRes);
    if (pd) return stampBody({ content: [{ type: 'text', text: JSON.stringify(withLadderFallback(pd, localRes), null, 2) }], isError: false }, null, snap.contractVersion);
  }
  // ---- MULTI-BRIDGE addressing (only when VB6MCP_BRIDGES is configured) ----------------------
  if (MULTI_BRIDGE && localScope && localRes) {
    const mp = localRes.pick || null;
    if (mp && mp.ok && mp.remote) {
      // (D2) `auto` is the ONLY unqualified form allowed to reach another bridge when bridges are
      // explicitly configured; a BARE selector still obeys N15 and is refused with the key to use.
      if (tgt.kind === 'auto') return forwardPicked(mp, name, args, snap, explicit, rawRequest, localRes && localRes.ladderFallback, implicitFwd, fwdDecidedBy);
      const q = remoteRequiresEndpointBody([mp], []);
      q.message = 'the selector "' + rawRequest + '" matches an IDE that lives on bridge "' + (mp.bridge || mp.endpoint) +
        '" -- a bare selector only ever addresses THIS machine, so the call was routed NOWHERE';
      q.qualifiedKeys = [mp.targetKey].filter(Boolean);
      q.suggestedAction = 'fix_params'; q.retrySafe = false; withLadderFallback(q, localRes);
      q.hint = 'pass _target="' + (mp.targetKey || (mp.bridge || 'bridge') + ':<pid>') + '" (or set ide_target{target:"..."}) to address that IDE explicitly';
      return stampBody({ content: [{ type: 'text', text: JSON.stringify(q, null, 2) }], isError: false }, null, snap.contractVersion);
    }
    // A cross-bridge ambiguous match is refused by the HUB before any delegation: delegating it to
    // this machine's middle service would silently pick the local side (the invisible same-pid bug).
    if (mp && !mp.ok && mp.ambiguous && mp.hits && mp.hits.length > 1) {
      const instsM = localRes.merged ? localRes.merged.instances : localRes.localD.instances;
      const cfM = localRes.merged ? localRes.merged.conflicts : localRes.localD.conflicts;
      return stampBody({ content: [{ type: 'text', text: JSON.stringify(withLadderFallback(refusalBody(mp, instsM, cfM), localRes), null, 2) }], isError: false }, null, snap.contractVersion);
    }
  }
  // ---- (WP-22-01) OFFLINE ROUTING: an IDE-FREE tool never needs a target ------------------
  // STRICT trigger (adversarial review B1): the hub must have ZERO live LOCAL instances AND the
  // implicit ladder must have failed for the "nothing is live" reason (resolution.step =
  // no_candidate).  An explicit selector, ambiguity, conflict or a remote-only set is NEVER
  // re-routed: those are refusals the caller must see, not a reason to run offline.
  if (localScope) {
    const off = offlineRouteFor(name, argsIn);
    if (off) {
      const localInsts = (localRes && localRes.localD) ? localRes.localD.instances : [];
      const pick = localRes && localRes.pick;
      const pickStep = (pick && pick.resolution && pick.resolution.step) || '';
      const noLiveLocal = localInsts.length === 0 && !(localRes && localRes.remoteOnly);
      const nothingToResolve = noLiveLocal && !explicit && (!pick || (!pick.ok && !pick.ambiguous && !pick.conflict && !pick.mismatch && pickStep === 'no_candidate'));
      let bridgeAgreesNoLive = true;
      if (nothingToResolve && BRIDGE_MODE !== 'inproc') {
        let bh = null;
        try { bh = await bridgeHealth(); } catch (e) { bh = null; }
        if (bh && Array.isArray(bh.instances)) bridgeAgreesNoLive = (bh.instances.length === 0);
        else bridgeAgreesNoLive = !middleServiceState.reachable;   // no addresser at all: offline is the only route
      }
      if (nothingToResolve && bridgeAgreesNoLive) {
        return await serveOfflineTool(off, name, argsIn, snap, {
          localD: (localRes && localRes.localD) || null,
          explicit: false, rawRequest: rawRequest, onlinePids: [],
          middleServiceAgreesNoLive: bridgeAgreesNoLive,
          droppedDeadCandidates: ((localRes && localRes.localD && localRes.localD.droppedDeadCandidates) || []),
          ladderFallback: (localRes && localRes.ladderFallback) || null
        });
      }
    }
  }
  if (BRIDGE_MODE !== 'inproc' && localScope) {
    const hubBlind = !localRes || localRes.localD.instances.length === 0;
    // (WP-22-01) 0 LIVE INSTANCES AND THE MIDDLE SERVICE AGREES: nothing to resolve a target
    // against, so an IDE-facing call answers no_live_ide (or the explicit-selector diagnosis)
    // WITHOUT delegating.  This is the path that used to hand a vanished pid down as a selector.
    const bctxOffline = bridgeIdentitySync();
    // "The middle service agrees there is no IDE" is only true when it ANSWERED (ok:true).  An
    // OFFLINE middle service has the same empty instances[] but says NOTHING about IDEs - and with
    // zero-config discovery (zero-config addressing) that difference decides whether the caller gets the legacy
    // no_live_ide or the structured no_middle_service receipt.
    const bridgeKnowsNoIde = !!(bctxOffline && bctxOffline.ok === true
      && Array.isArray(bctxOffline.instances) && bctxOffline.instances.length === 0);
    if (hubBlind && bridgeKnowsNoIde && localRes) {
      const body = explicit
        ? deadExplicitTargetBody(name, localRes.localD, rawRequest)
        : noLiveIdeBody(name, localRes.localD, explicit, rawRequest);
      return stampBody({ content: [{ type: 'text', text: JSON.stringify(withLadderFallback(body, localRes), null, 2) }], isError: false }, null, snap.contractVersion);
    }
    if (explicit) {
      // (1) AN EXPLICIT _target IS NEVER REWRITTEN.  The hub only ATTACHES its own verdict and its
      // own resolution attempt as ADDITIVE diagnostics; the caller's spelling goes down VERBATIM and
      // the middle service stays the resolver (exactly the pre-change behaviour for an explicit
      // selector).  The receipt says so: hubTargetNotRewritten=true.
      // "local:<selector>" is hub-only syntax: the middle service speaks pid|port|name|path.
      const localEpQualified = tgt.kind === 'endpoint' && tgt.endpoint && !tgt.endpoint.remote;
      const argsForBridge = localEpQualified ? Object.assign({}, argsIn, { _target: String(tgt.selector || '') }) : argsIn;
      const note = Object.assign({ hubTargetNotRewritten: !localEpQualified, targetRequested: rawRequest,
                                   hubSelectorSent: localEpQualified ? String(tgt.selector || '') : String(argsIn._target),
                                   hubSelectorForm: localEpQualified ? 'local_prefix_stripped' : 'verbatim_explicit',
                                   ...(localEpQualified ? { hubLocalPrefixStripped: 'local:' } : {}) },
        (localRes && localRes.pick && localRes.pick.ok) ? hubVerdictNote(localRes.pick, rawRequest,
          'explicit _target "' + String(argsIn._target) + '" was handed to the middle service VERBATIM (the hub never rewrites an explicit target)') : {},
        await ownershipNoteFor(localRes && localRes.pick, name));
      if (localRes && localRes.ladderFallback) note.ladderLiveFallback = localRes.ladderFallback;
      if (probePendingOf(localRes && localRes.pick)) note.probePending = true;
      if (fwdDecidedBy) note.forwardBudgetDecidedBy = fwdDecidedBy;
      const stampedEx = await bridgeForwardCall(name, argsForBridge, snap, note, localRes && localRes.pick, implicitFwd, fwdDecidedBy);
  if (forwardAnsweredOk(stampedEx)) noteForwardSuccess(localRes && localRes.pick);
      if (stampedEx) {
        const refusal = targetRefusalOf(stampedEx);
        if (refusal) {
          return stampBody(stampedEx, null, snap.contractVersion, {
            targetRefusedByMiddleService: true, noSilentFallback: true,
            targetRefusalReason: refusal.reason, targetRefusalMessage: refusal.message || null,
            hubSelectorSent: localEpQualified ? String(tgt.selector || '') : String(argsIn._target),
            hubSelectorForm: localEpQualified ? 'local_prefix_stripped' : 'verbatim_explicit',
            note2: 'the middle service refused the target itself: the hub did NOT retry another selector form and did NOT pick an instance on its own (no silent fallback)'
          });
        }
        return stampedEx;
      }
    } else if (localRes && localRes.pick && localRes.pick.ok) {
      const sel = canonicalSelectorFor(localRes.pick);
      const form = selectorFormOf(sel, localRes.pick);
      const argsOut = Object.assign({}, argsIn, { _target: sel });
      const note = Object.assign({}, hubVerdictNote(localRes.pick, rawRequest,
        'the hub resolved _target (step=' + ((localRes.pick.resolution && localRes.pick.resolution.step) || '?')
          + ') and handed the middle service the canonical selector "' + sel + '" (form=' + form + ')'),
        await ownershipNoteFor(localRes.pick, name),
        { targetResolved: sel, targetResolvedBy: 'hub', hubSelectorSent: sel, hubSelectorForm: form,
          hubTargetNotRewritten: false });
      if (localRes && localRes.ladderFallback) note.ladderLiveFallback = localRes.ladderFallback;
      if (localRes && localRes.ladderFallback && localRes.ladderFallback.used) note.selectedBy = 'ladder_live_fallback';
      if (probePendingOf(localRes && localRes.pick)) note.probePending = true;
      if (fwdDecidedBy) note.forwardBudgetDecidedBy = fwdDecidedBy;
      const stamped = await bridgeForwardCall(name, argsOut, snap, note, localRes && localRes.pick, implicitFwd, fwdDecidedBy);
  if (forwardAnsweredOk(stamped)) noteForwardSuccess(localRes && localRes.pick);
      if (stamped) {
        // (3) NO SILENT FALLBACK: if the middle service does not accept the form the hub handed it,
        // the call FAILS with a distinguishable reason, and the receipt records that nothing else
        // was tried and no instance was picked on the hub's own authority.
        const refusal = targetRefusalOf(stamped);
        if (refusal) {
          return stampBody(stamped, null, snap.contractVersion, {
            targetRefusedByMiddleService: true, noSilentFallback: true,
            targetRefusalReason: refusal.reason, targetRefusalMessage: refusal.message || null,
            hubSelectorSent: sel, hubSelectorForm: form,
            note2: 'the middle service refused the selector the hub resolved: NO other form was tried and NO instance was picked on the hub\'s own authority (no silent fallback)'
          });
        }
        return stamped;
      }
    } else if (!hubBlind && localRes) {
      const insts2 = (localRes.merged ? localRes.merged.instances : localRes.localD.instances);
      const cf2 = (localRes.merged ? localRes.merged.conflicts : localRes.localD.conflicts);
      if (localRes.remoteOnly) {
        dbg('handleCall: REFUSED unqualified call - 0 local IDEs, ' + localRes.remoteOnly.remoteLive.length + ' live remote IDE(s); an endpoint-qualified key is required');
        return stampBody({ content: [{ type: 'text', text: JSON.stringify(withLadderFallback(remoteRequiresEndpointBody(localRes.remoteOnly.remoteLive, localRes.remoteOnly.merged.conflicts), localRes), null, 2) }], isError: false }, null, snap.contractVersion);
      }
      dbg('handleCall: the hub REFUSED the target (reason=' + ((localRes.pick && (localRes.pick.conflict ? 'target_conflict' : (localRes.pick.mismatch ? 'target_mismatch' : (localRes.pick.ambiguous ? 'ambiguous_target' : 'no_ide_running')))) || 'no_ide_running') + '); NOT delegating to a second rule');
      return stampBody({ content: [{ type: 'text', text: JSON.stringify(withLadderFallback(refusalBody(localRes.pick || { resolution: resolutionOf('no_candidate', insts2, { requested: rawRequest, explicit: explicit }) }, insts2, cf2), localRes), null, 2) }], isError: false }, null, snap.contractVersion);
    } else {
      const stamped = await bridgeForwardCall(name, argsIn, snap, {
        targetRequested: rawRequest,
        targetResolution: { rule: TARGET_RESOLUTION_RULE, step: 'blind_delegation',
          basis: 'the hub\'s own discovery saw 0 add-in instances, so the target was resolved by the middle service (the only addresser left)' }
      });
      if (stamped) return stamped;
    }
    dbg('bridge offline: falling back to the pre-change local path for tool=' + name);
  }

  // ---- endpoint-qualified REMOTE target: probe THAT endpoint; never fall back to local ------
  if (tgt.kind === 'endpoint' && tgt.endpoint.remote) {
    const t0 = Date.now();
    const r = await probeEndpointCached(tgt.endpoint.label, REMOTE_PROBE_TIMEOUT_MS);
    if (!r || !r.ok) {
      const dd = r || {};
      // A remote endpoint that accepts TCP but never answers is a WEDGED peer; one that refuses
      // the connection / cannot be resolved is simply not there.
      const isAuthFail = dd.reason === 'endpoint_auth_failed';
      const epKind = dd.reason === 'endpoint_timeout' ? 'wedged' : (isAuthFail ? 'toolError' : 'noIde');
      return stampBody({ content: [{ type: 'text', text: JSON.stringify(Object.assign({
        ok: false, endpoint: tgt.endpoint.label, endpointUrl: tgt.endpoint.url, remote: true,
        reason: MULTI_BRIDGE ? (dd.reason || 'endpoint_unreachable')
          : (dd.reason === 'endpoint_timeout' ? 'endpoint_timeout' : 'endpoint_unreachable'),
        failureKind: epKind, failureKindRule: FAILURE_KIND_RULE,
        message: dd.message || 'the endpoint did not answer',
        elapsedMs: dd.elapsedMs === undefined ? (Date.now() - t0) : dd.elapsedMs,
        probeTimeoutMs: dd.timeoutMs || REMOTE_PROBE_TIMEOUT_MS,
        hint: isAuthFail
          ? 'the bridge answered and REFUSED this bridge credential: fix that bridge tokenFile and retry; nothing was forwarded'
          : 'a remote endpoint is only used when it is explicitly named; a local call is never redirected to it'
      }, MULTI_BRIDGE ? { suggestedAction: dd.suggestedAction || (isAuthFail ? 'fix_params' : 'retry_once'),
                          retrySafe: dd.retrySafe === true,
                          authReason: dd.authReason || null, authHint: dd.authHint || null,
                          token: dd.token || null,
                          bridgeFailures: (bridgeRowForLabel(tgt.endpoint.label)) } : {}), null, 2) }], isError: true }, null, snap.contractVersion);
    }
    // A bridge probe answers candidates[] (one per IDE); a legacy URL-only endpoint answers ONE
    // candidate.  Both shapes are accepted here, so the endpoint-qualified path works for either.
    const epCands = Array.isArray(r.candidates) ? r.candidates
      : (r.candidate ? [r.candidate] : []);
    const pickR = pickTarget(epCands, tgt.selector || null, { remoteScope: true, explicit: explicit });
    if (!pickR.ok) return stampBody({ content: [{ type: 'text', text: JSON.stringify(refusalBody(pickR, epCands, []), null, 2) }], isError: false }, null, snap.contractVersion);
    return forwardPicked(pickR, name, args, snap, explicit, rawRequest, null, false, null);
  }

  // ---- local / unqualified (the pre-change direct path) ---------------------------------------
  // The pick was already made by resolveLocalPick() above, with the SAME ladder the bridge path
  // used, so the two transports cannot disagree (checked:).
  dbg('forward phases: resolveLocal=' + localRes.elapsedMs + 'ms instances=' + localRes.localD.instances.length + (localRes.merged ? ' merged=' + localRes.merged.instances.length : ''));
  if (localRes.remoteOnly) {
    dbg('handleCall: REFUSED unqualified call - 0 local IDEs, ' + localRes.remoteOnly.remoteLive.length + ' live remote IDE(s); an endpoint-qualified key is required');
    return stampBody({ content: [{ type: 'text', text: JSON.stringify(withLadderFallback(remoteRequiresEndpointBody(localRes.remoteOnly.remoteLive, localRes.remoteOnly.merged.conflicts), localRes), null, 2) }], isError: false }, null, snap.contractVersion);
  }
  const pick = localRes.pick;
  if (!pick.ok) {
    const insts = (localRes.merged ? localRes.merged.instances : localRes.localD.instances);
    const cf = (localRes.merged ? localRes.merged.conflicts : localRes.localD.conflicts);
    return stampBody({ content: [{ type: 'text', text: JSON.stringify(withLadderFallback(refusalBody(pick, insts, cf), localRes), null, 2) }], isError: false }, null, snap.contractVersion);
  }
  return forwardPicked(pick, name, args, snap, explicit, rawRequest, localRes && localRes.ladderFallback, implicitFwd, fwdDecidedBy);
}

function send(obj) {
  // FAIL-CLOSED last line of defence: no JSON-RPC payload may carry a null/undefined
  // port URL.  The per-receipt guards make this unreachable in normal operation; if some future
  // path still builds one, the caller gets a structured replacement instead of a URL parse error.
  let out = JSON.stringify(obj);
  if (NULL_PORT_URL_RE.test(out)) {
    const id = (obj && obj.id !== undefined) ? obj.id : null;
    dbg('SEND GUARD: replaced a payload containing a null/undefined port URL (id=' + String(id) + ')');
    out = JSON.stringify({ jsonrpc: '2.0', id: id, result: { content: [{ type: 'text', text: JSON.stringify({
      ok: false, reason: 'internal_null_port_url', failureKind: 'toolError', failureKindRule: FAILURE_KIND_RULE,
      where: 'jsonrpc_send', bridgeUrl: BRIDGE_URL, bridgeMode: BRIDGE_MODE,
      middleService: middleServiceReceipt(),
      message: 'FAIL-CLOSED self-check: refused to emit an outward payload that contains a null/undefined port URL.',
      hint: 'a pipe-only IDE has no TCP endpoint; the middle service (VB6MCP_BRIDGE_URL) is its only addresser. Nothing was forwarded.'
    }, null, 2) }], isError: true } });
  }
  process.stdout.write(out + '\n');
}

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
    // listChanged is advertised ONLY outside inproc: the rollback path keeps its exact bytes.
    const capabilities = BRIDGE_MODE === 'inproc' ? { tools: {} } : { tools: { listChanged: true } };
    send({ jsonrpc: '2.0', id: id, result: {
      protocolVersion: v,
      capabilities: capabilities,
      serverInfo: SERVER,
      instructions: await buildInstructions(await resolveDiscovery(d), snap),
      faceSource: st.faceSource,
      contractVersion: st.contractVersion,
      toolCount: st.toolCount
    } });
    if (BRIDGE_MODE !== 'inproc') {
      newConnection();
      startFaceNotifyTimer();
    }
    return;
  }
  if (method && method.indexOf('notifications/') === 0) {
    if (method === 'notifications/initialized' && BRIDGE_MODE !== 'inproc') {
      conn.initialized = true;
      alignClientFace().catch(function () { });
    }
    return;
  }
  if (method === 'tools/list') {
    // R3-a: with VB6MCP_HUB_FACE_CACHE=1 a valid cached face is answered here (no pipe frame);
    // otherwise toolsListForClient() runs the PRE-CHANGE expression byte for byte.
    const faceOut = await toolsListForClient();
    send({ jsonrpc: '2.0', id: id, result: faceOut.result });
    // A cache HIT announces NOTHING (the face did not change) and does not re-write the
    // cross-process record: the build that produced these bytes already recorded that face.
    if (BRIDGE_MODE !== 'inproc' && !(faceOut.cache && faceOut.cache.hit)) noteFaceSynced(faceOut.result.tools);
    return;
  }
  if (method === 'tools/call') {
    const result = await handleCall(req.params);
    if (BRIDGE_MODE !== 'inproc') checkFaceChange('call').catch(function () { });
    return send({ jsonrpc: '2.0', id: id, result: result });
  }
  if (method === 'ping') return send({ jsonrpc: '2.0', id: id, result: {} });
  return send({ jsonrpc: '2.0', id: id, error: { code: -32601, message: 'method not found: ' + method } });
}

// Warm the discovery cache as soon as the process starts (background; never blocks a reply).
setTimeout(function () { discover().catch(function () { }); }, 0);
// First face probe at startup: a restarted hub compares the face the PREVIOUS client generation
// saw (hub_face_cache.json) with the current one and, if a client then initializes, announces it.
setTimeout(function () { bootFaceNotify(); }, 0);
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
// The middle service must be RESIDENT even when NO VB IDE exists at
// all, so the lazy ensure is ALSO kicked at HUB STARTUP (background - it must never delay
// initialize/tools/list, hence the timer and the ignored promise).  A later IDE needs no restart
// of the service: the bridge caches "no instance" for VB6MCP_DISCOVERY_TTL_MS (2 s) and a failed
// pipe probe for 2 s, so the next call after the add-in registers is already served.
// ---- R-1021-03 C: an EXECUTABLE check of the pass-through rule ----------------------------
// `node bridge/vb6mcp_hub.mjs --check-envelope` runs the pure envelope functions and exits.
// No network, no bridge, no stdin: it is the offline proof for the hub half of the client.
if (process.argv.indexOf('--check-envelope') >= 0) {
  const bad = [];
  const base = { ok: false, reason: 'timeout', reasonLegacy: 'forward_timeout',
    elapsedMs: 2100, maxMs: 2000, maxMsSource: 'param_slack+hard_cap',
    failureKind: 'peer_unresponsive', failureKindRule: 'rule', peerState: 'peer_unresponsive',
    suggestedAction: 'kill_and_restart', ctlAvailable: false,
    heldBy: { tool: 'dependency_graph', pid: 1, sinceMs: 2100 },
    oob: { lane: 'ctl', attempted: true, acked: false, why: 'ctl_no_ack', ackOp: null,
           helloMs: 0, rttMs: 0, delivered: 'queued_unacknowledged', applied: false,
           peerState: 'peer_unresponsive', suggestedAction: 'kill_and_restart' } };
  const clobber = { failureKind: 'toolError', peerState: 'peer_busy',
    suggestedAction: 'retry_once', maxMs: 999999, maxMsSource: 'hub',
    oob: { lane: 'business' }, reason: 'forward_failed', heldBy: null };
  const merged = mergeNote(Object.assign({}, base), clobber);
  for (const k of Object.keys(base)) {
    if (JSON.stringify(merged[k]) !== JSON.stringify(base[k])) bad.push('overwritten:' + k);
  }
  if (merged.failureKind !== 'peer_unresponsive') bad.push('failureKind_overwritten');
  if (merged.oob.why !== 'ctl_no_ack') bad.push('oob_overwritten');
  const filled = mergeNote({ ok: false, reason: 'x' },
    { suggestedAction: 'retry_once', maxMs: 2000 });
  if (filled.suggestedAction !== 'retry_once' || filled.maxMs !== 2000) bad.push('absent_not_filled');
  const lifted = errorEnvelopeOf({ jsonrpc: '2.0', id: 1, error: { code: -32000,
    message: 'forward failed', data: { maxMs: 2000, maxMsSource: 'param_slack',
      failureKind: 'peer_unresponsive', oob: { acked: false, why: 'ctl_no_ack' } } } });
  if (!lifted || !lifted.isError) bad.push('error_data_not_lifted');
  else {
    const lb = JSON.parse(lifted.content[0].text);
    if (lb.maxMs !== 2000 || lb.oob.why !== 'ctl_no_ack' || lb.jsonrpcError.code !== -32000) {
      bad.push('error_data_content');
    }
  }
  if (errorEnvelopeOf({ jsonrpc: '2.0', id: 2, error: { code: -32601, message: 'no data' } }) !== null) {
    bad.push('error_without_data_not_null');
  }
  if (!UPSTREAM_ENVELOPE_KEYS.length || !OOB_KEYS.length) bad.push('key_lists');
  process.stdout.write(JSON.stringify({ ok: bad.length === 0, checks: 6, failures: bad,
    envelopeKeys: UPSTREAM_ENVELOPE_KEYS.length, oobKeys: OOB_KEYS.length }, null, 2) + '\n');
  process.exit(bad.length === 0 ? 0 : 1);
}
if (BRIDGE_MODE !== 'inproc') {
  setTimeout(function () {
    ensureMiddleService().then(function (ms) {
      if (ms) dbg('middle ensure at startup: reachable=' + ms.reachable + ' reason=' + ms.reason + ' pid=' + ms.pid + ' managedByThisHub=' + ms.managedByThisHub);
    }).catch(function (e) { dbg('middle ensure at startup failed: ' + String((e && e.message) || e)); });
  }, 50);
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
