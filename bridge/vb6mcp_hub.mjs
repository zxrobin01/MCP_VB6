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
const BRIDGE_URL = (process.env.VB6MCP_BRIDGE_URL || 'http://127.0.0.1:8767').replace(/\/+$/, '');
const BRIDGE_HEALTH_TTL_MS = parseInt(process.env.VB6MCP_BRIDGE_HEALTH_TTL_MS || '5000', 10);
const BRIDGE_LIST_TIMEOUT_MS = parseInt(process.env.VB6MCP_BRIDGE_LIST_TIMEOUT_MS || '5000', 10);
// A forwarded call may be LONG (build / make_project): the shell must never be TIGHTER than the
// pre-change budget, so its bridge read timeout is FORWARD_TIMEOUT_MS plus a margin.  It is a
// FUNCTION on purpose: FORWARD_TIMEOUT_MS is declared below this block, so reading it here at
// module-evaluation time would be a temporal-dead-zone error.
function bridgeCallTimeoutMs() {
  return parseInt(process.env.VB6MCP_BRIDGE_CALL_TIMEOUT_MS || String(FORWARD_TIMEOUT_MS + 10000), 10);
}
// (threat model: bridge/TOKEN_THREAT_MODEL.md).  PURELY ADDITIVE: with no credential
// configured bridgeAuthHeaders() returns {}, so every request this shell makes is byte-identical
// to before.  The VALUE is never logged - only presence and length are ever reported, and a
// value read from a file is never copied into a child process environment.
const BRIDGE_TOKEN = (process.env.VB6MCP_TOKEN || '').trim()
  || (process.env.VB6MCP_TOKEN_FILE ? readBridgeTokenFile(process.env.VB6MCP_TOKEN_FILE) : '');
function readBridgeTokenFile(p) {
  try { return String(fs.readFileSync(p, 'utf8')).trim(); } catch (e) { return ''; }
}
function bridgeAuthHeaders() {
  return BRIDGE_TOKEN ? { Authorization: 'Bearer ' + BRIDGE_TOKEN } : {};
}
function bridgeAuthState() {
  return BRIDGE_TOKEN ? ('on(len=' + BRIDGE_TOKEN.length + ')') : 'off';
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
function forwardAnsweredOk(stamped) {
  if (!stamped || !stamped.result) return false;
  if (stamped.result.isError === true) return false;
  const obj = firstBodyObject(stamped.result);
  return !(obj && obj.ok === false);
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
  dbg('bridge mode=' + BRIDGE_MODE + ' url=' + BRIDGE_URL
      + (BRIDGE_MODE === 'inproc' ? ' (inproc = the PRE-CHANGE path; no bridge branch is entered)' : ''));
}, 0);
// GET /health, cached.  null means "the bridge is not answering" (then the pre-change snapshot /
// local path is used, which is what keeps the tool face complete with no IDE and no bridge).
function bridgeHealth() {
  if (BRIDGE_MODE === 'inproc') return Promise.resolve(null);
  const now = Date.now();
  if (bridgeHealthCache.data && now - bridgeHealthCache.at < BRIDGE_HEALTH_TTL_MS) return Promise.resolve(bridgeHealthCache.data);
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
async function bridgeForwardCall(name, argsIn, snap, note, pick) {
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
  const midNote = Object.assign({}, msr ? { middleService: msr } : {}, hcStamps ? { hubCode: hcStamps } : {});
  const h = await bridgeHealth();
  if (!h) return null;
  const t0 = Date.now();
  let resp;
  try {
    resp = await bridgeRpc('tools/call', { name: name, arguments: argsIn }, bridgeCallTimeoutMs());
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
      message: String((e && e.message) || e), elapsedMs: ms, budgetMs: bridgeCallTimeoutMs(),
      transport: 'bridge', transportReason: isAuth ? 'bridge_auth_failed' : 'bridge_unreachable',
      bridgeUrl: BRIDGE_URL, bridgeMode: BRIDGE_MODE, bridgeLastFailure: bridgeLastFailure || null,
      middleService: msr, hubCode: hcStamps,
      failureKind: isAuth ? 'toolError' : (isTimeout ? 'wedged' : 'noIde'), failureKindRule: FAILURE_KIND_RULE
    }, authBlock);
    const fwGuard = nullPortUrlGuard(fwBody, 'bridge_forward_failed(' + String(name) + ')');
    return stampBody({ content: [{ type: 'text', text: JSON.stringify(fwGuard || fwBody, null, 2) }], isError: true }, null, snap.contractVersion);
  }
  dbg('bridge forward ok tool=' + name + ' ms=' + (Date.now() - t0));
  const result = (resp && resp.result) || { content: [{ type: 'text', text: JSON.stringify(resp) }] };
// /N11: the hub's OWN liveness/resolution verdict rides along (ADDITIVE - stampBody only
  // fills keys the body does not already carry).  So a forwarded call that comes back
  // noIde/wedged can be read next to ide_list's verdict for the SAME instance, with the basis.
  // P2: this route just ANSWERED a real call - liveness evidence for ide_list.
  noteForwardSuccess(pick);
  return stampBody(result, null, snap.contractVersion, Object.assign({}, midNote, note || {}));
}
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
// of ONE candidate set and `remote` is only an extra marker (author ruling).
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
// ---- forward-failure classification (ADDITIVE;(3), evidence/-20) -----
// After a wedge the add-in still LISTENS, so a forwarded call fails in 1-68 ms with
// "forward_failed / fetch failed" - which reads like "the tool does not exist" or "the
// parameter is wrong", while the truth is a host that is there but not answering.
// failureKind turns that into a MACHINE decision; the legacy `reason` values never change.
// The rule is restated VERBATIM in every receipt that uses it (same convention as TARGET_KEY_RULE).
const FAILURE_KIND_RULE = 'failureKind = noIde (there is nothing to talk to: no live IDE, no route - unqualified call / unknown or unreachable endpoint, or the peer refused the connection) | wedged (a peer WAS found and did not answer: discovery saw it but the probe failed, or the forward timed out / the connection died mid-call - RemoteDisconnected) | toolError (the add-in ANSWERED with a structured error: a JSON-RPC error, result.isError=true, or an ok:false body). The legacy reason values are unchanged; failureKind is additive and is NOT injected when the add-in already classified the failure itself (e.g. desktop_probe failureKind=api_unavailable).';
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
  // projectPathAsGiven (ADDITIVE): the spelling the REGISTRY published, next to the canonical
  // realpath.  On this machine the same .vbp is published through junctions
  // (a directory junction that is an alias of the same real directory), so the alias is real information the caller
  // needs to see, not noise to be swallowed by normalization.
  return Object.assign({}, i, {
    endpoint: LOCAL_ENDPOINT, remote: false, endpointUrl: null,
    projectPath: ppath, pathKey: cp.key, pathVia: cp.via,
    projectPathAsGiven: i.projectPath || null,
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
const OWNERSHIP_VERDICT_RULE = "ownershipVerdict.verdict = owned | foreign | unknown (ATTRIBUTION: whose instance is it? It is NOT the permission axis - legacy ownership/ownedByHub decides whether ide_stop/ide_restart may act, and that predicate is unchanged). owned: the instance is attributable to THIS machine's add-in installation - the hub-started evidence is present (this hub process started the pid, or its registration says startedByHub=true), OR the pid is registered, ALIVE and the ADD-IN ITSELF answered mcp_info over its own route (a user-launched IDE is owned, NOT foreign: it is this machine's IDE, someone else simply started it). foreign: POSITIVE evidence only - (1) the registered pid is NOT alive while its port still ANSWERS, (2) the listening port is owned by an OS pid other than the pid the registration names, (3) the known port is inside the reserved range (< 8767; 8766 belongs to an external process) and ANSWERS, (4) the pid that serves the instance differs from the pid its OWN registration names. unknown: EVERYTHING else - above all portAvailable=false: a pipe-only registration publishes port 0 ON PURPOSE and an unreachable/unknown port is NOT evidence that the instance belongs to someone else. basis[] names every fact used; registryAlive and portAvailable are the raw inputs.";
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
  return { verdict: verdict,
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
    basis.push("remote endpoint: the pid/port/registration belong to ANOTHER machine, so attribution is not judged here");
    return ownershipVerdictBlock(f, 'unknown', basis, ownedEv, foreignEv);
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
  if (!f.inMemory && !f.registryStartedByHub && f.registryPresent && f.pidAlive && f.probeAnswered) {
    ownedEv.push('pid ' + f.pid + ' is registered, ALIVE and the ADD-IN ITSELF answered over its own route: this machine add-in installation, whoever started the process');
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
  const inMemory = !!(pid && startedByHub.has(pid));
  // the registration the instance CAME FROM wins; falling back to the one keyed by its pid.
  const byFile = (inst && inst.registryFile) ? ownershipRegistrationByFile(inst.registryFile) : null;
  const byPid = pid ? ownershipRegistrationOf(pid) : null;
  const reg = byFile || byPid;
  const registryPresent = !!reg;
  const registeredPid = reg && reg.pid ? Number(reg.pid) : pid;
  const registryStartedByHub = !!(reg && reg.startedByHub);
  const pidAliveNow = pid ? pidAlive(pid) : false;
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
  const po = portOwnerPidsOf(port);
  return ownershipVerdictOf({
    pid: pid, port: port, remote: remote, inMemory: inMemory,
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
  const bi = (i && i.pid && byPid[String(i.pid)]) || null;
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
    const rec = { at: Date.now(), faceFingerprint: faceFingerprintNow(), instances: d.instances, conflicts: d.conflicts, staleRegistry: d.staleRegistry, registryDir: d.registryDir };
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
async function currentFaceFingerprint() {
  const res = await toolsListResult(await discoverSoft(LIST_GRACE_MS));
  return faceFingerprintOf((res && res.tools) || []);
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
  writeFaceNotifyRecord({ reason: p.why, fingerprint: p.fp.fingerprint, toolCount: p.fp.toolCount, delivered: true });
  try {
    process.stderr.write('[hub-face] notifications/tools/list_changed sent conn=' + conn.id +
      ' reason=' + p.why + ' tools=' + p.fp.toolCount + ' fingerprint=' + p.fp.fingerprint.slice(0, 16) + '\n');
  } catch (e) { }
}
async function checkFaceChange(why, force) {
  if (BRIDGE_MODE === 'inproc') return null;
  if (faceProbeInFlight) return faceProbeInFlight;
  const now = Date.now();
  if (!force && now - lastFaceProbeAt < FACE_PROBE_MIN_MS) return null;
  lastFaceProbeAt = now;
  faceProbeInFlight = (async function () {
    let fp = null;
    try { fp = await currentFaceFingerprint(); }
    catch (e) { dbg('face probe failed: ' + String((e && e.message) || e)); }
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

async function rpcAt(base, method, params, proto, timeoutMs, rpcPath) {
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
  const credential = isBridgeBase(target.base) || authVerdict.get(target.base) === true;
  if (credential) Object.assign(headers, bridgeAuthHeaders());
  const url = target.base + (rpcPath || target.path);
  const payload = JSON.stringify(body);
  const ms = timeoutMs || FORWARD_TIMEOUT_MS;
  let res = await fetch(url, { method: 'POST', headers: headers, body: payload, signal: AbortSignal.timeout(ms) });
  let raw = await res.text();
  // A LOOPBACK endpoint that answered 401 with OUR OWN reason asked for the credential: repeat the
  // request ONCE with it (and remember the verdict per endpoint, so this costs one extra POST per
  // endpoint, not per call).  See asksForOurToken for the three conditions.
  if (!credential && asksForOurToken(target.base, res.status, raw)) {
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
  // Best-effort banner from the MCP face; /rpc does not implement initialize.  Fetched in the
  // BACKGROUND so a slow /mcp cannot keep the instance "pending".
  const instructions = instructionsCache[port] || null;
  if (!instructions) {
    instanceRpc(port, transport, 'initialize', { protocolVersion: LATEST, capabilities: {}, clientInfo: { name: 'vb6mcp-hub', version: SERVER.version } }, null, 3000, '/mcp')
      .then(function (init) { const v = init && init.result && init.result.instructions; if (v) { instructionsCache[port] = v; dbg('instructions cached for port ' + port); } })
      .catch(function () { dbg('instructions fetch failed for port ' + port); });
  }
  const build = (info && info.build) || {};
  const ide = (info && info.ide) || null;
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
    projectPathAsGiven: rawPath || null,
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
  return { _target: c.pid ? String(c.pid) : (c.projectPath || c.project || ''),
           projectPath: c.projectPath || null, targetKey: c.targetKey || null,
           ide_target: c.pid ? String(c.pid) : null,
           note: 'a bare pid is accepted by the hub AND by the middle service; the canonical key is [endpoint:]projectPath' };
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
      return { ok: false, mismatch: true, candidates: insts.map(candidateBrief), suggestion: suggestionFor(insts),
        resolution: resolutionOf('refuse_target_mismatch', insts, opts),
        message: 'target "' + s + '" not found among ' + desc() + ' -- pick one with _target (pid|project|port) or ide_target' };
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
  // P1: the receipt can PROVE which code answered (sha16 of the running file + start time).
  const hcStamps = hubCodeStamps();
  if (hcStamps) st.hubCode = hcStamps;
  return st;
}

async function hubTool(name, args) {
  const snap = snapshot();
  if (name === 'ide_list') return textResult(await hubState({ requestTarget: args && args._target }), snap.contractVersion);
  if (name === 'ide_target') {
    const t = args && args.target;
    state.target = (t === 'auto' || t === '') ? null : t;
// /N14: 'auto' is not a selector name, it is the ABSENCE of a default - and the
    // receipt now states what an absent default resolves to, so the caller can predict it.
    return textResult({ ok: true, target: state.target,
      cleared: state.target === null,
      targetResolutionRule: TARGET_RESOLUTION_RULE,
      selectionRule: SELECTION_RULE,
      note: state.target === null
        ? 'the default target is now UNSET: an omitted _target (or _target="auto") follows the documented ladder - the ONLY live IDE, else the ONLY hub-owned IDE, else the ONLY responding IDE, else a refusal that lists every candidate and a suggestion'
        : 'every call without an explicit _target now defaults to "' + state.target + '"; a default that matches no live instance is IGNORED (it never vetoes a live set)' }, snap.contractVersion);
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
    if (!own.owned) {
      body.reason = 'not_allowed';
      body.message = 'pid ' + pid + ' is not owned by this hub: ' + own.ownershipBasis + '; not stopping it (the same ownership rule ide_list and ide_restart report)' + (own.foreignPort ? ' The port guard also applies: ' + own.portBasis + '.' : '');
      return textResult(body, snap.contractVersion);
    }
    if (own.foreignPort) {
      body.reason = 'foreign_instance_refused';
      body.message = own.portBasis + ' => this instance is not ours; not stopping it';
      return textResult(body, snap.contractVersion);
    }
    try { process.kill(Number(pid)); startedByHub.delete(Number(pid)); invalidateDiscovery();
          return textResult(Object.assign(body, { ok: true, stopped: Number(pid), message: 'stopped pid ' + pid + ' (' + own.ownership + ')' }), snap.contractVersion); }
    catch (e) { return textResult(Object.assign(body, { reason: 'kill_failed', message: String(e && e.message || e) }), snap.contractVersion); }
  }
// an INSTANCE that is REGISTERED and whose PROCESS IS ALIVE, but which did not answer the
// probe. 'no_ide_running' is reserved for the case where NO such entry exists. Measured (
// item 10): the wedged IDE held the only pipe instance and ide_list listed it as probeStatus=failed,
// while ide_restart answered 'no_ide_running' - a falsehood that hid the one pid to kill.
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
// The refusal an UNKNOWN attribution produces.  It must NOT be reported as not_owned: "not ours"
// and "cannot tell" are different answers (that conflation is exactly the original complaint).
// The legacy string the pre-fix code would have used rides along as legacyReason, so an existing
// consumer that only knows the old vocabulary is not broken by the honesty.
function ownershipUnknownRefusal(out, blk, legacyReason) {
  out.refused_reason = 'ownership_unknown';
  out.legacyReason = legacyReason || 'not_owned';
  out.ownershipUnknown = true;
  out.howToDisambiguate = 'nothing was stopped. To settle WHOSE instance this is: (1) pass the exact pid AND project=<the absolute .vbp> so the target cannot be a recycled pid; (2) start the middle service (VB6MCP_BRIDGE_URL) if the registration is pipe-only, or make its port reachable, so the add-in can ANSWER mcp_info - a registered+alive pid that answers is attributable (owned); (3) if the registration is a residue, remove %TEMP%\\vb6mcp\\instances\\<pid>.json and retry; (4) if the port ANSWERED without the registered pid being alive, that is foreign evidence (not unknown) and the verdict says so.';
  return out;
}
function registryCandidateView(list) {
  return (list || []).map(function (a) {
    return { pid: a.pid, port: a.port, transport: a.transport, project: a.project,
             projectPath: a.projectPath, startedByHub: a.startedByHub, registryFile: a.registryFile };
  });
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
      failureKind: '', failureKindRule: FAILURE_KIND_RULE, conflicts: []
    };
    const wantPid = (args && args.pid) ? Number(args.pid) : 0;
    const wantProject = (args && args.project) || process.env.VB6MCP_AUTOSTART_PROJECT || '';
    const soft = await discoverSoft(LIST_GRACE_MS);
    const d = await resolveDiscovery(soft);
    const insts = (d.instances || []).filter(function (i) { return i.pid; });
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
          const ovA1 = await resolveOwnership(a);
          out.ownershipVerdict = ovA1;
          if (ovA1.verdict === 'unknown') {
            ownershipUnknownRefusal(out, ovA1, o.foreignPort ? 'not_ours_by_red_line' : 'instance_found_but_unresponsive');
            out.truth = 'instance_found_but_ownership_unknown';
            out.note = 'pid ' + wantPid + ' is registered (' + a.registryFile + ') and its process is ALIVE, but WHOSE instance it is cannot be established: ' + ovA1.basis.join('; ') + '. Nothing was stopped.';
            out.nextStep = 'disambiguate first (howToDisambiguate): the manual command "' + out.manualCommand + '" is only for an instance you have confirmed is yours.';
            return textResult(out, snap.contractVersion);
          }
          out.refused_reason = o.foreignPort ? 'not_ours_by_red_line' : 'instance_found_but_unresponsive';
          out.note = 'pid ' + wantPid + ' is not a LIVE instance, but it IS registered (' + a.registryFile + ') and its process is ALIVE, so it is unresponsive rather than absent. ' + (o.foreignPort ? o.portBasis + '; not ours, nothing was stopped.' : 'Ownership: ' + o.ownershipBasis + '. Nothing was stopped.');
          out.nextStep = 'run "' + out.manualCommand + '" on the machine that owns the pid and retry, or use ide_stop{pid:' + a.pid + '} when the ownership verdict above is owned.';
          return textResult(out, snap.contractVersion);
        }
        out.refused_reason = 'pid_not_found';
        out.instanceCount = 0;
        out.checkedRegistryDir = instancesDir();
        out.truth = 'no_instance_at_all';
        out.note = 'pid ' + wantPid + ' is not a live instance AND no registration with that pid has an alive process; see instances[]';
        return textResult(out, snap.contractVersion);
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
    return textResult(out, snap.contractVersion);
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
        return textResult(out, snap.contractVersion);
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
      return textResult(out, snap.contractVersion);
    }
    out.refused_reason = 'no_ide_running';
    out.truth = 'no_instance_at_all';
    out.nextStep = 'start an IDE (ide_start) or pass a pid that exists.';
    out.note = 'nothing to restart: the registry (' + out.checkedRegistryDir + ') holds no entry whose process is alive - no instance at all';
    return textResult(out, snap.contractVersion);
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
          return textResult(out, snap.contractVersion);
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
    out.instanceCount = insts.length;
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
      return textResult(out, snap.contractVersion);
    }
    if (!own.owned) {
      out.refused_reason = 'not_owned';
      out.truth = 'instance_found_but_not_owned';
      out.note = 'this instance is not ours (' + own.ownershipBasis + '), so NOTHING was stopped. The ownership rule is shared with ide_list and ide_stop, so all three agree about pid ' + target.pid + '.' + (own.foreignPort ? ' The port guard also applies: ' + own.portBasis + '.' : '');
      return textResult(out, snap.contractVersion);
    }
    if (own.foreignPort) {
      out.refused_reason = 'foreign_instance_refused';
      out.truth = 'instance_found_but_not_ours';
      out.note = own.portBasis + ': refusing to stop it.';
      return textResult(out, snap.contractVersion);
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
async function toolsListResult(d) {
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
          return { tools: normalizeToolSchemas(mergeHubFace(tools)) };     // strict: EXACTLY {tools}
        }
      } catch (e) { dbg('bridge tools/list FAILED: ' + String((e && e.message) || e)); }
    }
    dbg('bridge offline: serving the complete face from ' + SNAP + ' + ' + HUBCFG);
  }
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
async function forwardPicked(pick, name, args, snap, explicit, rawRequest) {
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
  if (isRemote) { ideStamp.endpoint = pick.endpoint; ideStamp.remote = true; }
  const timeoutMs = isRemote ? REMOTE_FORWARD_TIMEOUT_MS : FORWARD_TIMEOUT_MS;
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
                      failureKind: transportFailureKind(e), failureKindRule: FAILURE_KIND_RULE,
                      forwardRoute: forwardRouteOf(pick), probeRoute: pick.probeRoute || null,
                      port: fe.port, bridgeUrl: isRemote ? null : BRIDGE_URL, bridgeMode: BRIDGE_MODE,
                      registry: isRemote ? null : registrySnapshotFor(pick.pid),
                      liveness: pick.liveness || null };
    if (isRemote) { errBody.endpoint = pick.endpoint; errBody.endpointUrl = pick.endpointUrl; errBody.remote = true; }
    else if (isTimeout) { errBody.hint = 'raise VB6MCP_FORWARD_TIMEOUT_MS for legitimately long calls'; }
    const guardedErr = nullPortUrlGuard(errBody, 'forward_failed(' + String(name) + ')');
    return stampBody({ content: [{ type: 'text', text: JSON.stringify(guardedErr || errBody, null, 2) }], isError: true }, ideStamp, snap.contractVersion, pickNote);
  }
  dbg('forward phases: rpc=' + (Date.now() - tFwd) + 'ms tool=' + name + ' ' + (isRemote ? 'endpoint=' + pick.endpoint : 'port=' + pick.port));
  const result = (resp && resp.result) || { content: [{ type: 'text', text: JSON.stringify(resp) }] };
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
async function resolveLocalPick(tgt, explicit, rawRequest) {
  const localD = await resolveDiscovery(await discoverSoft(LIST_GRACE_MS));
  const tRes = Date.now();
  const popts = { remoteScope: false, explicit: explicit, requested: rawRequest };
  let merged = null;
  let pick = null;
  let remoteOnly = null;
  if (tgt.kind === 'endpoint') {                       // "local:<selector>"
    pick = pickTarget(localD.instances, tgt.selector || null, popts);
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
  return { localD: localD, merged: merged, pick: pick, remoteOnly: remoteOnly, elapsedMs: Date.now() - tRes };
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
  if (BRIDGE_MODE !== 'inproc' && localScope) {
    const hubBlind = !localRes || localRes.localD.instances.length === 0;
    if (explicit) {
      // (1) AN EXPLICIT _target IS NEVER REWRITTEN.  The hub only ATTACHES its own verdict and its
      // own resolution attempt as ADDITIVE diagnostics; the caller's spelling goes down VERBATIM and
      // the middle service stays the resolver (exactly the pre-change behaviour for an explicit
      // selector).  The receipt says so: hubTargetNotRewritten=true.
      const note = Object.assign({ hubTargetNotRewritten: true, targetRequested: rawRequest,
                                   hubSelectorSent: String(argsIn._target), hubSelectorForm: 'verbatim_explicit' },
        (localRes && localRes.pick && localRes.pick.ok) ? hubVerdictNote(localRes.pick, rawRequest,
          'explicit _target "' + String(argsIn._target) + '" was handed to the middle service VERBATIM (the hub never rewrites an explicit target)') : {},
        await ownershipNoteFor(localRes && localRes.pick, name));
      const stampedEx = await bridgeForwardCall(name, argsIn, snap, note, localRes && localRes.pick);
  if (forwardAnsweredOk(stampedEx)) noteForwardSuccess(localRes && localRes.pick);
      if (stampedEx) {
        const refusal = targetRefusalOf(stampedEx);
        if (refusal) {
          return stampBody(stampedEx, null, snap.contractVersion, {
            targetRefusedByMiddleService: true, noSilentFallback: true,
            targetRefusalReason: refusal.reason, targetRefusalMessage: refusal.message || null,
            hubSelectorSent: String(argsIn._target), hubSelectorForm: 'verbatim_explicit',
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
      const stamped = await bridgeForwardCall(name, argsOut, snap, note, localRes && localRes.pick);
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
        return stampBody({ content: [{ type: 'text', text: JSON.stringify(remoteRequiresEndpointBody(localRes.remoteOnly.remoteLive, localRes.remoteOnly.merged.conflicts), null, 2) }], isError: false }, null, snap.contractVersion);
      }
      dbg('handleCall: the hub REFUSED the target (reason=' + ((localRes.pick && (localRes.pick.conflict ? 'target_conflict' : (localRes.pick.mismatch ? 'target_mismatch' : (localRes.pick.ambiguous ? 'ambiguous_target' : 'no_ide_running')))) || 'no_ide_running') + '); NOT delegating to a second rule');
      return stampBody({ content: [{ type: 'text', text: JSON.stringify(refusalBody(localRes.pick || { resolution: resolutionOf('no_candidate', insts2, { requested: rawRequest, explicit: explicit }) }, insts2, cf2), null, 2) }], isError: false }, null, snap.contractVersion);
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
      const epKind = dd.reason === 'endpoint_timeout' ? 'wedged' : 'noIde';
      return stampBody({ content: [{ type: 'text', text: JSON.stringify({
        ok: false, endpoint: tgt.endpoint.label, endpointUrl: tgt.endpoint.url, remote: true,
        reason: dd.reason === 'endpoint_timeout' ? 'endpoint_timeout' : 'endpoint_unreachable',
        failureKind: epKind, failureKindRule: FAILURE_KIND_RULE,
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

  // ---- local / unqualified (the pre-change direct path) ---------------------------------------
  // The pick was already made by resolveLocalPick() above, with the SAME ladder the bridge path
  // used, so the two transports cannot disagree (checked:).
  dbg('forward phases: resolveLocal=' + localRes.elapsedMs + 'ms instances=' + localRes.localD.instances.length + (localRes.merged ? ' merged=' + localRes.merged.instances.length : ''));
  if (localRes.remoteOnly) {
    dbg('handleCall: REFUSED unqualified call - 0 local IDEs, ' + localRes.remoteOnly.remoteLive.length + ' live remote IDE(s); an endpoint-qualified key is required');
    return stampBody({ content: [{ type: 'text', text: JSON.stringify(remoteRequiresEndpointBody(localRes.remoteOnly.remoteLive, localRes.remoteOnly.merged.conflicts), null, 2) }], isError: false }, null, snap.contractVersion);
  }
  const pick = localRes.pick;
  if (!pick.ok) {
    const insts = (localRes.merged ? localRes.merged.instances : localRes.localD.instances);
    const cf = (localRes.merged ? localRes.merged.conflicts : localRes.localD.conflicts);
    return stampBody({ content: [{ type: 'text', text: JSON.stringify(refusalBody(pick, insts, cf), null, 2) }], isError: false }, null, snap.contractVersion);
  }
  return forwardPicked(pick, name, args, snap, explicit, rawRequest);
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
    const face = await toolsListResult(await discoverSoft(LIST_GRACE_MS));
    send({ jsonrpc: '2.0', id: id, result: face });
    if (BRIDGE_MODE !== 'inproc') noteFaceSynced(face.tools);
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
