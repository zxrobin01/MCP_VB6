#!/usr/bin/env python3
# -*- coding: utf-8 -*-
r"""vb6mcp_bridge.py - the VB6MCP middle service layer.

WHAT THIS IS (ADR docs/design/TRANSPORT_NP_DUAL.md sec 0.1, the 2026-10-15 three-layer ruling)

    MCP client --stdio/HTTP--> mjs thin shell (MCP wire protocol + guardian + degradation)
                                  |  (the ONLY TCP port on this machine; owned by THIS file)
                         Python middle service layer (HTTP forward / discovery / addressing /
                                                       named-pipe multiplexing / policy)
                                  |  (one named pipe PER IDE instance)
                         VB6 add-in (no TCP listener; JSON-RPC over the pipe)

This file is the middle layer.  It is the single TCP listener for every IDE on the
machine: ONE port no matter how many VB6 instances are running.  It keeps TODAY's client
surface so that no existing tool changes:

    POST /rpc        plain JSON-RPC (what tools/dump_tools_schema.py and
                     tests/battery_test.py / tests/smoke_test.py use)
    POST /mcp        the MCP streamable-HTTP face (what tools/mcp_strict_audit.py uses)
    --stdio          line-delimited JSON-RPC on stdin/stdout (the stdio entry)

    GET /health|/ping, GET /, OPTIONS, DELETE /mcp, 404/405/413        - mirrors the add-in

TRACKS (env VB6MCP_TRANSPORT)

    auto (default)   prefer the pipe; fall back to the add-in's HTTP /rpc, VISIBLY
    pipe             pipe only; a missing/unusable pipe is an explicit error
    tcp              the add-in's HTTP /rpc only (spike control track, rollback, transition)

  NOTE ON THE DEFAULT (recorded deviation, see the report): ADR sec 0.2 writes "default pipe",
  ADR sec 2.1 writes "default auto".  This implementation treats UNSET as "prefer pipe, fallback
  allowed" and honours an EXPLICIT VB6MCP_TRANSPORT=pipe as strict (sec 2.3's
  pipe_forced_but_absent error).  Reason: during the transition an old add-in registry entry has
  no pipe field at all, so a strict-by-default pipe track would refuse to serve a working machine.

NAMED PIPE WIRE PROTOCOL (np/1) - fixed by the spike, agreed with the VB side

    pipe name   \\\\.\\pipe\\vb6mcp-<sid>-<pid>-<slug>-<nonce16hex>
    framing     [4-byte LITTLE-ENDIAN length][UTF-8 payload], one frame <= 8 MiB,
                PIPE_TYPE_BYTE + PIPE_READMODE_BYTE (spike: message mode does NOT remove
                re-framing - one ReadFile returned a message containing 3 frames)
    handshake   C->S {"vb6mcp":"hello","proto":"np/1","nonce":...,
                      "client":{"pid":..,"session":..,"user":".."}}
                S->C {"vb6mcp":"welcome","proto":"np/1","instanceId":"<pid>-<nonce>",
                      "pid":..,"session":..,"project":..,"dllSha256":.., ...}
                S->C {"vb6mcp":"deny","reason":"nonce_mismatch"|"bad_hello"|"frame_too_large"}
    traffic     every later frame payload is EXACTLY the JSON-RPC message /rpc carries
    progress    optional {"vb6mcp":"progress","seq":n,"elapsedMs":n,"tool":"<maybe empty>"}
    heartbeat   optional {"vb6mcp":"ping"} / {"vb6mcp":"pong"}

AUTHENTICATION (threat model = bridge/TOKEN_THREAT_MODEL.md)

    VB6MCP_TOKEN              the bearer credential (preferred: it never touches disk)
    VB6MCP_TOKEN_FILE         owner-only file holding it (used only when VB6MCP_TOKEN is unset)

    Every HTTP surface - /rpc, /mcp, /health and the rest - requires it once configured, and
    answers 401 with a structured reason otherwise (token_missing / token_mismatch /
    token_bad_scheme / token_in_url_rejected).  A credential in the URL is REJECTED, never
    honoured; it is never echoed and never logged (the log keeps length + source only, and a
    redact() guard scrubs every logged field).  Send it as `Authorization: Bearer <token>` or
    `X-VB6MCP-Token`.

    With NO credential configured the service behaves exactly as before, but /health says
    auth.enabled=false and prints the shared-machine warning, and the startup line warns on
    BOTH stderr and the log - an unauthenticated face must never be invisible.

    BINDING: the DEFAULT tier is loopback-only - a non-loopback host is refused with exit 6,
    exactly as before.  REMOTE MODE (docs/design/REMOTE_ACCESS.md) is the explicit,
    fail-closed exception: it needs --allow-remote + --tls-cert/--tls-key + a token source + a
    writable --audit-log (missing any of them refuses to start with exit 7 and names them), a
    non-loopback caller is READ-ONLY unless --remote-write, the source is re-checked in process
    (403), tokens are compared in constant time with a previous token kept valid for rotation,
    failures back off exponentially, every response is audited once (JSONL) and /health is
    redacted for remote callers.  A bearer token authenticates an ENDPOINT, not a user's intent:
    a client that legitimately holds the token is outside this control - which is exactly why
    remote callers are read-only by default.

BOUNDED WAITS (ADR sec 7.4) - this client can ALWAYS free itself

    VB6MCP_PIPE_CONNECT_MS   300    connect + hello/welcome only
    VB6MCP_FORWARD_TIMEOUT_MS 120000  the CALL budget: identical on both tracks, never tighter
    VB6MCP_PIPE_PROGRESS_MS  2000   expected progress cadence
    VB6MCP_PIPE_IDLE_MS      6000   (= 3 x progress cadence) dead-connection threshold, ARMED
                                    ONLY AFTER THE FIRST progress frame is seen.  A server that
                                    never sends progress (old add-in) degrades to
                                    FORWARD_TIMEOUT_MS, i.e. never tighter than TCP.

HONESTY / KNOWN GAPS (also printed by --selftest and listed in the report)

  * the LEGACY MCP SSE transport (GET /sse + POST /message) is NOT implemented here; it
    answers 501 with a stated reason.  No in-scope client uses it.
  * a request that was FULLY WRITTEN to a pipe and then lost its answer is NEVER silently
    replayed on another track (that could double-apply a write tool).  VB6MCP_PIPE_REPLAY=1
    opts in.  A request that failed BEFORE it was fully written is safely retried.
  * discovery is registry + WaitNamedPipe/CreateFile only: this file never runs netstat and
    never scans a port range.

FROZEN FORM (PyInstaller onefile) AND ITS ACCEPTANCE

  Under onefile, __file__ lives in the temporary _MEIxxxx unpack directory, so the two frozen
  artifacts are resolved through DATA_DIRS (VB6MCP_BRIDGE_DATA_DIR -> sys._MEIPASS -> the EXE's
  own directory -> bridge/).  In SOURCE form (no env set, not frozen) the resolution is exactly
  bridge/, i.e. unchanged.

  A frozen build is accepted with --smoke, NOT with --selftest: --selftest runs tools/*.py and
  tests/*.py out of the repository, which a bundle does not carry, so it refuses (exit 2) when
  frozen.  --smoke carries its own stub np/1 add-in and serves the real /rpc + /mcp face, so the
  SAME binary can be exercised anywhere; every line it prints is one JSON object.

  REPRODUCIBLE PACKAGING (run from the REPOSITORY ROOT; outputs outside the tree on purpose):

      $dist = "$env:TEMP\vb6mcp_bridge_dist"
      python -m PyInstaller --noconfirm --onefile --clean --name vb6mcp_bridge \
        --distpath $dist --workpath "$env:TEMP\vb6mcp_bridge_build" \
        --specpath "$env:TEMP\vb6mcp_bridge_build" \
        --add-data "<REPO>\bridge\hub_schema.generated.json;." \
        --add-data "<REPO>\bridge\hub_tools.json;." \
        "<REPO>\bridge\vb6mcp_bridge.py"

  * --specpath RE-ROOTS the relative source path of --add-data to the SPEC directory, not the
    CWD.  Use ABSOLUTE source paths (as above) or omit --specpath; relative paths silently
    produce a bundle with no assets and the failure looks like a code bug.
  * the two JSON artifacts MUST be bundled: without them tools/list falls back to an empty face.
  * acceptance, two shells:
        & "$dist\vb6mcp_bridge.exe" --smoke --smoke-seconds 45 --port 8791
        python tools/mcp_strict_audit.py --url http://127.0.0.1:8791/mcp
        python tools/dump_tools_schema.py --url http://127.0.0.1:8791/rpc --out $env:TEMP\exe_dump.json
    and the source form of the same three commands; the two runs must agree on
    snapshotFingerprint / liveFingerprint / liveTools (== 110) and the dumped tool array must
    equal docs/tools_schema.json field for field.

Exit codes: 0 ok | 2 usage | 3 refused/occupied port | 4 bind failure | 5 unreadable snapshot
"""

from __future__ import annotations

import argparse
import hashlib
import ipaddress
import ssl
import base64
import ctypes
import hashlib
import hmac
import http.server
import json
import os
import re
import shutil
import socket
import subprocess
import sys
import tempfile
import threading
import time
import urllib.error
import urllib.request
import uuid
from ctypes import wintypes

PRODUCT = "vb6mcp-bridge"
VERSION = "0.1.0"
NP_PROTO = "np/1"
FRAME_MAX = 8 * 1024 * 1024          # sec 7.3 single-frame ceiling
HTTP_BODY_MAX = 8 * 1024 * 1024      # request body ceiling of this listener
DEFAULT_PORT = 8767                  # >= tools/mcp_strict_audit.py MIN_SELF_OWNED_PORT
RESERVED_PORT = 8766                 # never bound by this repository's tooling

#: WHEN this process started.  ide_list/ide_status publish uptimeMs from it - a pure fact
#: about THIS middle service (never persisted, never a schema).
_BOOT_AT = time.time()
_BOOT_ISO = time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime(_BOOT_AT))

#: the directory THIS FILE lives in, in SOURCE form.
_SRC_DIR = os.path.dirname(os.path.abspath(__file__))


def _bridge_home():
    """Where the bundled data is looked for.

    Under PyInstaller onefile, `__file__` lives in the temporary `_MEIxxxx` unpack directory,
    so a frozen build must resolve its assets from there (or from next to the EXE, where an
    operator can drop an override copy).  In source form this is exactly `bridge/`, so the
    source behaviour is unchanged.
    """
    if getattr(sys, "frozen", False):
        return getattr(sys, "_MEIPASS", None) or os.path.dirname(os.path.abspath(sys.executable))
    return _SRC_DIR


HERE = _bridge_home()
#: the repository root - a SOURCE-form concept; the selftest refuses to run when frozen.
REPO_ROOT = os.path.dirname(_SRC_DIR)


def _data_dirs():
    """Ordered candidates for the frozen artifacts, most specific first.

    `VB6MCP_BRIDGE_DATA_DIR` is an explicit override (also useful in source form to test a
    different snapshot without touching the tree).  With no override the source form resolves
    to `bridge/` only - byte-identical behaviour to before this change.
    """
    dirs = []
    env = os.environ.get("VB6MCP_BRIDGE_DATA_DIR")
    if env:
        dirs.append(env)
    if getattr(sys, "frozen", False):
        meipass = getattr(sys, "_MEIPASS", None)
        if meipass:
            dirs.append(meipass)
        dirs.append(os.path.dirname(os.path.abspath(sys.executable)))
    dirs.append(_SRC_DIR)
    out = []
    for d in dirs:
        if d and d not in out:
            out.append(d)
    return out


DATA_DIRS = _data_dirs()


def _resolve_data_file(name):
    for d in DATA_DIRS:
        p = os.path.join(d, name)
        if os.path.isfile(p):
            return p
    return os.path.join(_SRC_DIR, name)   # keep the miss message pointing at the source copy


SNAP_PATH = _resolve_data_file("hub_schema.generated.json")
HUB_TOOLS_PATH = _resolve_data_file("hub_tools.json")


# ---------------------------------------------------------------- configuration ----

def _env_str(name, default=None):
    v = os.environ.get(name)
    return default if v is None or v == "" else v


def _env_int(name, default):
    v = os.environ.get(name)
    if v is None or v == "":
        return default
    try:
        return int(str(v).strip())
    except ValueError:
        return default


def _env_bool(name, default):
    v = os.environ.get(name)
    if v is None or v == "":
        return default
    return str(v).strip().lower() in ("1", "true", "yes", "on")


class Config(object):
    """Process-level policy.  Read once at start, never per call (sec 2.1)."""

    def __init__(self):
        self.transport = (_env_str("VB6MCP_TRANSPORT", "") or "").strip().lower()
        self.port = _env_int("VB6MCP_BRIDGE_PORT", _env_int("VB6MCP_PORT", DEFAULT_PORT))
        self.host = _env_str("VB6MCP_BRIDGE_HOST", "127.0.0.1")
        self.pipe_connect_ms = _env_int("VB6MCP_PIPE_CONNECT_MS", 300)
        self.forward_timeout_ms = _env_int("VB6MCP_FORWARD_TIMEOUT_MS", 120000)
        # R-1021-03 A4: TWO independent layers (mirror of CodeTools.ToolBudgetMs).
        #   max_call_ms      generic ceiling: tools WITHOUT a budget of their own only.
        #   max_call_hard_ms absolute cap: every derived budget, including build /
        #                    make_project and an explicit caller timeout_ms.
        self.max_call_ms = _env_int("VB6MCP_MAX_CALL_MS", 120000)
        self.max_call_hard_ms = _env_int("VB6MCP_MAX_CALL_HARD_MS", 600000)
        if self.max_call_hard_ms < self.max_call_ms:
            self.max_call_hard_ms = self.max_call_ms
        # one immediate retry when the pipe reports ERROR_PIPE_BUSY (a just-started or
        # just-abandoned instance is a TRANSIENT busy, not "no IDE")
        self.pipe_busy_retry_ms = _env_int("VB6MCP_PIPE_BUSY_RETRY_MS", 300)
        # R-1021-03 A4 client half: the ctl/1 CONTROL LANE (bridge/CTL_WIRE_CONTRACT.md).
        # Feature-detected from the instance record; this switch only exists so an operator can
        # prove a business-only path still works after a rollback.  Disabling it can never make
        # the bridge fall back to an UNAUTHENTICATED control channel - it simply does not use
        # the lane (the business lane is the only other route, and it is nonce-paired).
        self.ctl_enabled = _env_bool("VB6MCP_CTL", True)
        # the hello deadline is the SERVER's VB6MCP_PIPE_HELLO_MS window; the ack deadline bounds
        # ONE op on the lane (4.7: every ctl op must be fast - a slow op delays release_gate).
        self.ctl_hello_ms = _env_int("VB6MCP_CTL_HELLO_MS", _env_int("VB6MCP_PIPE_HELLO_MS", 5000))
        self.ctl_ack_ms = _env_int("VB6MCP_CTL_ACK_MS", 1500)
        # the out-of-process read-only fallback (section B).  Only used when the ctl lane is
        # UNAVAILABLE and the peer was classified peer_unresponsive.
        self.oop_enabled = _env_bool("VB6MCP_OOP_FALLBACK", True)
        self.progress_ms = _env_int("VB6MCP_PIPE_PROGRESS_MS", 2000)
        # NB: deliberately NOT "VB6MCP_PIPE_IDLE_MS" - the VB side uses that name for the
        # SERVER's 120 s idle close.  Sharing the name would turn a client stall threshold of
        # 6 s into 120 s by accident (a silent loss of the fast escape), so the client's
        # threshold is bridge-namespaced.
        self.pipe_idle_ms = _env_int("VB6MCP_BRIDGE_PIPE_IDLE_MS", 0) or (3 * self.progress_ms)
        self.pipe_fallback = _env_bool("VB6MCP_PIPE_FALLBACK", True)
        self.pipe_replay = _env_bool("VB6MCP_PIPE_REPLAY", False)
        self.debug = _env_bool("VB6MCP_HUB_DEBUG", False)
#: step 3 - the INBOUND ingress. "pipe" (the default) is byte-for-byte
        #: the pre-existing path; "textbox" pushes the request into the add-in's hidden
        #: IngressForm textbox with SendMessageTimeout (which returns in microseconds)
        #: so the tool runs AFTER the write returned instead of inside it.  An
        #: unrecognised value never silently enables a new path.
        self.ingress = (_env_str("VB6MCP_INGRESS", "") or "pipe").strip().lower()
        if self.ingress not in ("pipe", "textbox"):
            self.ingress = "pipe"
        #: an APPEND-ONLY diagnostic log.  Without it a failed joint run leaves a 0-byte file
        #: and no way to tell "the IDE was invisible" from "the handshake was rejected".
        self.log_path = _env_str("VB6MCP_BRIDGE_LOG", "")
        self.target = _env_str("VB6MCP_TARGET", "") or ""
        self.instances_dir = _env_str("VB6MCP_INSTANCES_DIR", "") or os.path.join(
            tempfile.gettempdir(), "vb6mcp", "instances")
        self.discovery_ttl_ms = _env_int("VB6MCP_DISCOVERY_TTL_MS", 2000)
        # ---- authentication (threat model = bridge/TOKEN_THREAT_MODEL.md) -----------
        # M-5: environment first (never touches disk), then an owner-only file.  With neither,
        # the service still runs on loopback and REPORTS that it is unauthenticated - a weak
        # configuration must be visible, not silent.
        self.token_env = _env_str("VB6MCP_TOKEN", "")
        self.token_file = _env_str("VB6MCP_TOKEN_FILE", "")
        self.token = ""
        self.token_source = "none"          # env | file | none
        self.token_note = ""

    def resolve_token(self):
        """(token, source, note).  Called once by main(); never logs the value (M-4)."""
        if self.token_env:
            self.token, self.token_source = self.token_env, "env"
            if self.token_file:
                self.token_note = "VB6MCP_TOKEN_FILE ignored because VB6MCP_TOKEN is set"
            return self.token, self.token_source, self.token_note
        if self.token_file:
            try:
                with open(self.token_file, "r", encoding="utf-8-sig") as fh:
                    self.token = fh.read().strip()
            except OSError as exc:
                self.token_note = "VB6MCP_TOKEN_FILE unreadable: %s" % exc
                self.token, self.token_source = "", "none"
                return self.token, self.token_source, self.token_note
            if self.token:
                self.token, self.token_source = self.token, "file"
            else:
                self.token_note = "VB6MCP_TOKEN_FILE is empty"
        return self.token, self.token_source, self.token_note

    def auth_enabled(self):
        return bool(self.token)

    def preference(self):
        """'pipe' | 'tcp'.  Unset and 'auto' both mean pipe-first (see the module docstring)."""
        if self.transport == "tcp":
            return "tcp"
        return "pipe"

    def strict(self):
        """True only for an EXPLICIT pipe/tcp force (sec 2.3 hard semantics)."""
        return self.transport in ("pipe", "tcp")


CFG = Config()
_DEBUG_LOCK = threading.Lock()


def dbg(*parts):
    if not CFG.debug:
        return
    line = "[bridge-debug] " + " ".join(str(p) for p in parts) + "\n"
    with _DEBUG_LOCK:
        try:
            sys.stderr.write(line)
            sys.stderr.flush()
        except Exception:
            pass


_LOG_LOCK = threading.Lock()


def is_loopback_host(host):
    """True only for a literal loopback address (M-0).

    The bridge refuses to listen on anything else.  This is a WHITELIST, not a denylist: an
    operator cannot invent a host that accidentally exposes the 120-tool face, and 0.0.0.0 is
    impossible rather than merely discouraged (ADR sec 0.5 forbids the exposure outright).
    """
    h = ("" if host is None else str(host)).strip().lower()
    if h in ("localhost", "::1", "[::1]"):
        return True
    if h.startswith("127."):
        parts = h.split(".")
        if len(parts) == 4:
            try:
                return all(0 <= int(p) <= 255 for p in parts)
            except ValueError:
                return False
    return False


def _token_digest(value):
    return hashlib.sha256(("" if value is None else str(value)).encode("utf-8")).digest()


def token_matches(provided, expected):
    """Constant-time comparison (M-2).

    Both sides are SHA-256 digests first, so the comparison is over equal-length buffers:
    neither the content nor the LENGTH of the supplied credential leaks through timing, and
    an empty/absent credential cannot short-circuit.
    """
    if not expected:
        return True                      # auth disabled: nothing to check
    return hmac.compare_digest(_token_digest(provided), _token_digest(expected))


def redact(text):
    """Defence in depth for M-4: no diagnostic may ever carry the token value.

    Applies to every field bridge_log writes, so an accidental token=... can never reach the
    log.  The log records length and source instead (a fingerprint of the secret would be a
    guess-confirmation oracle, so we deliberately do NOT emit one).
    """
    s = "" if text is None else str(text)
    if CFG.token and CFG.token in s:
        s = s.replace(CFG.token, "<redacted>")
    return s


def auth_state():
    """The auth block that /health and the startup report publish - no secret, no fingerprint."""
    return {
        "enabled": CFG.auth_enabled(),
        "required": CFG.auth_enabled(),
        "scheme": "bearer",
        "source": CFG.token_source,
        "tokenLen": len(CFG.token),
        "note": CFG.token_note or None,
        "warning": None if CFG.auth_enabled() else
                   ("no token configured: the face is reachable by ANY local user on loopback "
                    "- set VB6MCP_TOKEN (or VB6MCP_TOKEN_FILE) before any non-local use"),
    }


def bridge_log(event, **fields):
    """One diagnostic record, to stderr AND to VB6MCP_BRIDGE_LOG when it is set.

    Every record carries a monotonically increasing sequence so a reader can tell a truncated
    tail from a full log.  This exists because a joint run once produced a 0-byte bridge.log
    while the real story ("the registry file could not be decoded") was invisible.
    """
    global _LOG_SEQ
    with _LOG_LOCK:
        _LOG_SEQ += 1
        seq = _LOG_SEQ
    parts = []
    for k in sorted(fields):
        v = redact(fields[k])
        if isinstance(v, str) and (" " in v or v == ""):
            v = '"%s"' % v
        parts.append("%s=%s" % (k, v))
    line = "%s seq=%d event=%s %s\n" % (time.strftime("%Y-%m-%d %H:%M:%S"), seq, event,
                                       " ".join(parts))
    with _DEBUG_LOCK:
        try:
            sys.stderr.write("[bridge] " + line)
            sys.stderr.flush()
        except Exception:
            pass
    if CFG.log_path:
        try:
            with open(CFG.log_path, "a", encoding="utf-8") as fh:
                fh.write(line)
        except Exception as exc:
            with _DEBUG_LOCK:
                try:
                    sys.stderr.write("[bridge] LOG_WRITE_FAILED %s %r\n" % (CFG.log_path, exc))
                except Exception:
                    pass


def nonce_diag(expected, sending):
    """The pre-send self-check the VB writer asked for (C7): len + hex + equality.

    A length or content difference is the ONE failure the add-in reports as an opaque
    deny reason=nonce_mismatch, so it is checked and logged BEFORE the frame is written.
    """
    e = "" if expected is None else str(expected)
    s = "" if sending is None else str(sending)
    return {
        "expectedLen": len(e), "expectedHex": e.encode("utf-8", "replace").hex(),
        "sendingLen": len(s), "sendingHex": s.encode("utf-8", "replace").hex(),
        "equal": e == s,
    }


_LOG_SEQ = 0


# ---------------------------------------------------------------- win32 kernel32 ----

kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)

GENERIC_READ = 0x80000000
GENERIC_WRITE = 0x40000000
OPEN_EXISTING = 3
FILE_FLAG_OVERLAPPED = 0x40000000
INVALID_HANDLE_VALUE = ctypes.c_void_p(-1).value

PIPE_ACCESS_DUPLEX = 0x00000003
PIPE_TYPE_BYTE = 0x00000000
PIPE_READMODE_BYTE = 0x00000000
PIPE_WAIT = 0x00000000

WAIT_OBJECT_0 = 0x00000000
WAIT_TIMEOUT = 0x00000102
WAIT_FAILED = 0xFFFFFFFF

ERROR_FILE_NOT_FOUND = 2
ERROR_ACCESS_DENIED = 5
ERROR_BROKEN_PIPE = 109
ERROR_PIPE_BUSY = 231
ERROR_NO_DATA = 232
ERROR_PIPE_NOT_CONNECTED = 233
ERROR_MORE_DATA = 234
ERROR_PIPE_CONNECTED = 535
ERROR_IO_PENDING = 997
ERROR_IO_INCOMPLETE = 996
ERROR_SEM_TIMEOUT = 121

PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
STILL_ACTIVE = 259


class OVERLAPPED(ctypes.Structure):
    _fields_ = [
        ("Internal", ctypes.c_void_p),
        ("InternalHigh", ctypes.c_void_p),
        ("Offset", wintypes.DWORD),
        ("OffsetHigh", wintypes.DWORD),
        ("hEvent", wintypes.HANDLE),
    ]


kernel32.CreateFileW.restype = wintypes.HANDLE
kernel32.CreateFileW.argtypes = [wintypes.LPCWSTR, wintypes.DWORD, wintypes.DWORD,
                                 ctypes.c_void_p, wintypes.DWORD, wintypes.DWORD,
                                 wintypes.HANDLE]
kernel32.WaitNamedPipeW.restype = wintypes.BOOL
kernel32.WaitNamedPipeW.argtypes = [wintypes.LPCWSTR, wintypes.DWORD]
kernel32.ReadFile.restype = wintypes.BOOL
kernel32.ReadFile.argtypes = [wintypes.HANDLE, ctypes.c_void_p, wintypes.DWORD,
                              ctypes.POINTER(wintypes.DWORD), ctypes.POINTER(OVERLAPPED)]
kernel32.WriteFile.restype = wintypes.BOOL
kernel32.WriteFile.argtypes = [wintypes.HANDLE, ctypes.c_void_p, wintypes.DWORD,
                               ctypes.POINTER(wintypes.DWORD), ctypes.POINTER(OVERLAPPED)]
kernel32.PeekNamedPipe.restype = wintypes.BOOL
kernel32.PeekNamedPipe.argtypes = [wintypes.HANDLE, ctypes.c_void_p, wintypes.DWORD,
                                   ctypes.POINTER(wintypes.DWORD), ctypes.POINTER(wintypes.DWORD),
                                   ctypes.POINTER(wintypes.DWORD)]
kernel32.CreateEventW.restype = wintypes.HANDLE
kernel32.CreateEventW.argtypes = [ctypes.c_void_p, wintypes.BOOL, wintypes.BOOL, wintypes.LPCWSTR]
kernel32.WaitForSingleObject.restype = wintypes.DWORD
kernel32.WaitForSingleObject.argtypes = [wintypes.HANDLE, wintypes.DWORD]
kernel32.GetOverlappedResult.restype = wintypes.BOOL
kernel32.GetOverlappedResult.argtypes = [wintypes.HANDLE, ctypes.POINTER(OVERLAPPED),
                                         ctypes.POINTER(wintypes.DWORD), wintypes.BOOL]
kernel32.CloseHandle.restype = wintypes.BOOL
kernel32.CloseHandle.argtypes = [wintypes.HANDLE]
kernel32.CreateNamedPipeW.restype = wintypes.HANDLE
kernel32.CreateNamedPipeW.argtypes = [wintypes.LPCWSTR, wintypes.DWORD, wintypes.DWORD,
                                      wintypes.DWORD, wintypes.DWORD, wintypes.DWORD,
                                      wintypes.DWORD, ctypes.c_void_p]
kernel32.ConnectNamedPipe.restype = wintypes.BOOL
kernel32.ConnectNamedPipe.argtypes = [wintypes.HANDLE, ctypes.POINTER(OVERLAPPED)]
kernel32.DisconnectNamedPipe.restype = wintypes.BOOL
kernel32.DisconnectNamedPipe.argtypes = [wintypes.HANDLE]
kernel32.FlushFileBuffers.restype = wintypes.BOOL
kernel32.FlushFileBuffers.argtypes = [wintypes.HANDLE]
kernel32.CancelIo.restype = wintypes.BOOL
kernel32.CancelIo.argtypes = [wintypes.HANDLE]
kernel32.SetEvent.restype = wintypes.BOOL
kernel32.SetEvent.argtypes = [wintypes.HANDLE]
kernel32.ResetEvent.restype = wintypes.BOOL
kernel32.ResetEvent.argtypes = [wintypes.HANDLE]
kernel32.OpenProcess.restype = wintypes.HANDLE
kernel32.OpenProcess.argtypes = [wintypes.DWORD, wintypes.BOOL, wintypes.DWORD]
kernel32.GetExitCodeProcess.restype = wintypes.BOOL
kernel32.GetExitCodeProcess.argtypes = [wintypes.HANDLE, ctypes.POINTER(wintypes.DWORD)]


def pid_alive(pid):
    """True when a process with this pid is alive.

    OpenProcess + GetExitCodeProcess, NOT kill(pid,0): a pid that was recycled by an
    unrelated process must not read as 'our IDE' (an earlier fix), and this is only ever used
    together with the nonce pairing, which is what actually proves identity.
    """
    try:
        pid = int(pid)
    except (TypeError, ValueError):
        return False
    if pid <= 0:
        return False
    h = kernel32.OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, False, pid)
    if not h:
        return False
    try:
        code = wintypes.DWORD()
        if not kernel32.GetExitCodeProcess(h, ctypes.byref(code)):
            return False
        return code.value == STILL_ACTIVE
    finally:
        kernel32.CloseHandle(h)


# ---------------------------------------------------------------- framing ---------

class FrameError(Exception):
    def __init__(self, reason, detail=""):
        Exception.__init__(self, reason + ((": " + detail) if detail else ""))
        self.reason = reason
        self.detail = detail


def encode_frame(obj):
    """One np/1 frame: 4-byte little-endian length + UTF-8 JSON payload.

    COMPACT separators, deliberately, but NOT as a fix for anything: measured against the LIVE
    add-in on t16, a spaced payload (123 B) and a compact one (111 B) BOTH get a welcome, so the
    add-in really parses the JSON and whitespace is tolerated.  Compact is kept because it is
    deterministic and slightly smaller.  The hypothesis that spacing caused the handshake
    rejections was TESTED AND IS FALSE - the rejections came from the registry DECODE below.
    """
    payload = json.dumps(obj, ensure_ascii=False, separators=(",", ":")).encode("utf-8")
    if len(payload) > FRAME_MAX:
        raise FrameError("frame_too_large", "%d > %d" % (len(payload), FRAME_MAX))
    return len(payload).to_bytes(4, "little") + payload


def decode_frame_payload(payload):
    try:
        return json.loads(payload.decode("utf-8"))
    except Exception as exc:
        raise FrameError("bad_json", str(exc))


class FrameBuffer(object):
    """Accumulates raw pipe bytes and pops complete frames.

    Short reads are NORMAL (sec 7.3) - a ReadFile that returns fewer bytes than requested
    must never be mistaken for the end of a frame.
    """

    def __init__(self):
        self.buf = bytearray()

    def feed(self, data):
        self.buf.extend(data)

    def pop(self):
        """Return the next complete payload (bytes) or None; raises on an over-long frame."""
        if len(self.buf) < 4:
            return None
        n = int.from_bytes(bytes(self.buf[:4]), "little")
        if n > FRAME_MAX:
            raise FrameError("frame_too_large", "declared=%d" % n)
        if len(self.buf) < 4 + n:
            return None
        payload = bytes(self.buf[4:4 + n])
        del self.buf[:4 + n]
        return payload


class PipeError(Exception):
    def __init__(self, code, detail="", stage="io", partial=b""):
        Exception.__init__(self, "pipe error %s %s (%s)" % (code, detail, stage))
        self.code = code
        self.detail = detail
        self.stage = stage
        self.partial = partial


def _pipe_reason(code):
    return {
        ERROR_FILE_NOT_FOUND: "absent",
        ERROR_PIPE_BUSY: "busy",
        # WaitNamedPipe reports ERROR_SEM_TIMEOUT when every instance is CONNECTED,
        # which is the same situation as CreateFile's ERROR_PIPE_BUSY (spike 4.3).
        ERROR_SEM_TIMEOUT: "busy",
        ERROR_NO_DATA: "closed",
        ERROR_PIPE_NOT_CONNECTED: "closed",
        ERROR_BROKEN_PIPE: "closed",
        ERROR_ACCESS_DENIED: "denied",
    }.get(code, "error")


def wait_named_pipe(name, ms):
    """(state, code).  state in ok|busy|absent|error.  NO netstat, NO port scan (sec 1.2.4)."""
    ok = kernel32.WaitNamedPipeW(ctypes.c_wchar_p(name), int(ms))
    if ok:
        return "ok", 0
    code = ctypes.get_last_error()
    return _pipe_reason(code) if code in (ERROR_FILE_NOT_FOUND, ERROR_PIPE_BUSY,
                                          ERROR_ACCESS_DENIED, ERROR_SEM_TIMEOUT) else "error", code


def wait_named_pipe_bounded(name, budget_ms, tries=4):
    """WaitNamedPipe with a few IMMEDIATE-'absent' retries inside one connect budget.

    An error_ms-free ERROR_FILE_NOT_FOUND is not always 'the IDE is gone': the VB side writes a
    deny frame, waits for the client to read it (or 750 ms), closes the handle and rebuilds the
    same pipe within ~250 ms (measurement).  Retrying costs nothing when the pipe is
    genuinely gone because WaitNamedPipe returns immediately in that case.
    """
    deadline = time.monotonic() + budget_ms / 1000.0
    last = ("error", 0)
    for _ in range(max(1, tries)):
        remaining = max(1, int((deadline - time.monotonic()) * 1000))
        last = wait_named_pipe(name, remaining)
        if last[0] != "absent":
            return last
        if time.monotonic() >= deadline:
            break
        time.sleep(min(0.05, max(0.0, deadline - time.monotonic())))
    return last


# ---------------------------------------------------------------- pipe client ------

class PipeClient(object):
    """A bounded, fully overlapped np/1 client.  Every wait has a deadline, so this class can
    always free itself (ADR sec 7.4: 'the client must be able to escape on its own')."""

    def __init__(self, name):
        self.name = name
        self.handle = None
        self.ev = None
        self.rbuf = FrameBuffer()
        self.last_payload = None
        self.last_progress_ms = 0
        self.progress_seq = 0

    # -- lifecycle ---------------------------------------------------------------
    def connect(self, timeout_ms):
        deadline = time.monotonic() + max(1, timeout_ms) / 1000.0
        state, code = wait_named_pipe_bounded(self.name, max(1, timeout_ms))
        if state != "ok":
            raise PipeError(code, "WaitNamedPipe=%s" % state, "connect")
        while True:
            h = kernel32.CreateFileW(ctypes.c_wchar_p(self.name), GENERIC_READ | GENERIC_WRITE,
                                     0, None, OPEN_EXISTING, FILE_FLAG_OVERLAPPED, None)
            if h != INVALID_HANDLE_VALUE and h:
                self.handle = h
                break
            code = ctypes.get_last_error()
            if code != ERROR_PIPE_BUSY:
                raise PipeError(code, "CreateFile", "connect")
            if time.monotonic() >= deadline:
                raise PipeError(ERROR_PIPE_BUSY, "CreateFile busy", "connect")
            time.sleep(0.005)
        self.ev = kernel32.CreateEventW(None, True, False, None)
        if not self.ev:
            self.close()
            raise PipeError(ctypes.get_last_error(), "CreateEvent", "connect")
        return self

    def close(self):
        try:
            if self.handle:
                kernel32.CloseHandle(self.handle)
        except Exception:
            pass
        try:
            if self.ev:
                kernel32.CloseHandle(self.ev)
        except Exception:
            pass
        self.handle = None
        self.ev = None

    @property
    def closed(self):
        return self.handle is None

    # -- raw overlapped I/O ------------------------------------------------------
    def _overlapped(self):
        ov = OVERLAPPED()
        ov.hEvent = self.ev
        ctypes.memset(ctypes.byref(ov), 0, ctypes.sizeof(ov))
        ov.hEvent = self.ev
        return ov

    def _wait_io(self, ov, deadline, what, buf=None):
        """Wait for THIS overlapped op inside the deadline.

        Loops (never recurses) on ERROR_IO_INCOMPLETE: the spike measured that
        GetOverlappedResult(bWait=0) reports 996, not 997, and recursing there could run
        once per spin until the I/O finished.
        """
        while True:
            remaining = deadline - time.monotonic()
            if remaining <= 0:
                raise PipeError(ERROR_SEM_TIMEOUT, "%s budget exhausted" % what, "timeout",
                                self._abandon(ov, buf))
            wr = kernel32.WaitForSingleObject(self.ev, int(remaining * 1000) or 1)
            if wr == WAIT_TIMEOUT:
                raise PipeError(ERROR_SEM_TIMEOUT, "%s timed out" % what, "timeout",
                                self._abandon(ov, buf))
            if wr == WAIT_FAILED:
                raise PipeError(ctypes.get_last_error(), "WaitForSingleObject", "io")
            got = wintypes.DWORD()
            if kernel32.GetOverlappedResult(self.handle, ctypes.byref(ov),
                                            ctypes.byref(got), False):
                return got.value
            code = ctypes.get_last_error()
            if code in (ERROR_IO_INCOMPLETE, ERROR_IO_PENDING):
                time.sleep(0.001)
                continue
            raise PipeError(code, "GetOverlappedResult", "io")

    def _abandon(self, ov, buf=None):
        """Cancel a pending op and REAP it, so a deadline never leaks an I/O request.

        The handle is deliberately NOT closed here: a benign idle poll (a read that simply
        found no frame yet) must not tear down a healthy persistent connection.  Any bytes
        that completed before the cancellation are handed back so framing cannot desync.
        """
        try:
            kernel32.CancelIo(self.handle)
        except Exception:
            pass
        got_n = 0
        try:
            got = wintypes.DWORD()
            if kernel32.GetOverlappedResult(self.handle, ctypes.byref(ov), ctypes.byref(got), True):
                got_n = got.value
        except Exception:
            pass
        if buf is not None and got_n:
            return buf.raw[:got_n]
        return b""

    def write_all(self, data, deadline):
        """(fully_written, sent).  fully_written=False means a retry is safe (sec 6.3)."""
        view = memoryview(data)
        sent = 0
        while sent < len(data):
            kernel32.ResetEvent(self.ev)   # the wait must actually wait for THIS write
            chunk = bytes(view[sent:sent + 65536])
            buf = ctypes.create_string_buffer(chunk, len(chunk))
            ov = self._overlapped()
            nwritten = wintypes.DWORD()
            ok = kernel32.WriteFile(self.handle, ctypes.cast(buf, ctypes.c_void_p), len(chunk),
                                    ctypes.byref(nwritten), ctypes.byref(ov))
            if not ok:
                code = ctypes.get_last_error()
                if code != ERROR_IO_PENDING:
                    return False, sent
                try:
                    n = self._wait_io(ov, deadline, "WriteFile")
                except PipeError:
                    return False, sent
            else:
                n = nwritten.value or len(chunk)
            if not n:
                return False, sent
            sent += n
        return True, sent

    def read_some(self, want, deadline):
        """(data, None) | (b'', 'timeout') | (None, PipeError).  One bounded overlapped ReadFile."""
        kernel32.ResetEvent(self.ev)
        buf = ctypes.create_string_buffer(want)
        ov = self._overlapped()
        nread = wintypes.DWORD()
        ok = kernel32.ReadFile(self.handle, ctypes.cast(buf, ctypes.c_void_p), want,
                               ctypes.byref(nread), ctypes.byref(ov))
        if not ok:
            code = ctypes.get_last_error()
            if code != ERROR_IO_PENDING:
                if code in (ERROR_BROKEN_PIPE, ERROR_NO_DATA, ERROR_PIPE_NOT_CONNECTED):
                    return None, PipeError(code, "read", "read")
                return None, PipeError(code, "read", "read")
            try:
                n = self._wait_io(ov, deadline, "ReadFile", buf)
            except PipeError as exc:
                if exc.stage == "timeout":
                    return (exc.partial or b""), "timeout"
                return None, exc
        else:
            n = nread.value
        if n == 0:
            return None, PipeError(ERROR_BROKEN_PIPE, "eof", "read")
        return buf.raw[:n], None

    def bytes_available(self):
        try:
            avail = wintypes.DWORD()
            if kernel32.PeekNamedPipe(self.handle, None, 0, None, ctypes.byref(avail), None):
                return avail.value
        except Exception:
            pass
        return None

    # -- np/1 frames -------------------------------------------------------------
    def send_frame(self, obj, deadline):
        return self.write_all(encode_frame(obj), deadline)

    def recv_frame(self, deadline):
        """Next frame as a dict.  (None, 'timeout') when the deadline passes with no frame."""
        while True:
            try:
                payload = self.rbuf.pop()
            except FrameError as exc:
                raise PipeError(ERROR_MORE_DATA, "bad frame: %s" % exc.reason, "frame")
            if payload is not None:
                self.last_payload = payload          # exact bytes, for byte-pass-through
                return decode_frame_payload(payload), None
            want = max(4096, min(65536, FRAME_MAX))
            data, err = self.read_some(want, deadline)
            if err == "timeout":
                if data:
                    self.rbuf.feed(data)             # never lose bytes reaped from a cancelled read
                return None, "timeout"
            if err is not None:
                return None, err
            self.rbuf.feed(data)


# ---------------------------------------------------------------- ctl lane (ctl/1) ----
# bridge/CTL_WIRE_CONTRACT.md.  The CONTROL LANE is a SECOND named pipe with exactly three
# jobs, all of which must work while a long call holds the single-flight BUSINESS gate:
#   1. liveness/state questions (ping, state),
#   2. recovery of a stuck gate (release_gate, cancel, abandon),
#   3. a SMALL read-only, FAST tool subset (tool, 4.7).
# It is NOT a second business lane.  This block is the CLIENT half only.  Three client rules
# are load bearing and are enforced here rather than by convention:
#   * feature-detect from the instance record; NEVER fall back to an unauthenticated channel;
#   * every frame carries sid + a strictly increasing seq; a mismatch kills the connection;
#   * an effect is claimed ONLY when the server acknowledged it (oob.acked) - a wedged peer
#     produces delivered:"queued_unacknowledged" + applied:false, never a success.

CTL_PROTO = "ctl/1"
CTL_OPS = ("ping", "state", "release_gate", "cancel", "abandon", "tool")

# 4.7 CTL read-only subset = (the add-in's ToolReadOnly table) INTERSECTED WITH (fast: well
# under one second, never opens a dialog or a modal host loop).  A WHITELIST on purpose: a new
# read-only-but-slow tool must not leak onto the recovery lane, where a slow call would sit in
# front of release_gate/cancel (the lane serves ONE connection at a time, sec 1).
CTL_FAST_TOOLS = (
    "ide_state", "execution_state", "desktop_probe", "read_active_dialogs", "find_gui",
    "ui_text", "last_error", "list_windows", "current_position", "mcp_info",
    "get_codepane", "get_project_info",
)

# 4 preamble: the server's static honesty block.  The client REPORTS a block a server actually
# sent (under limits) and never invents one: fabricating it would present the server's own
# stated limits as if the server had spoken them.
CTL_RESULT_PASS = (
    "released", "releasedTool", "heldMs", "depthAfter", "gateReleaseCount", "statementInterrupted",
    "requested", "honoredAtNextPump", "abandoned", "abandonedTool", "responseLost", "businessPipe",
    "inFlightTool", "howToProceed", "note", "cancel", "release", "inflight", "gate",
    "mode", "project", "pid", "uptimeMs", "ackedAt", "pump", "ctl", "source",
)

#: oob.why - SIX distinct failure grades.  Collapsing them into one "ctl unavailable" would hide
#: exactly the distinction an operator needs: a lane that was never published, one that could
#: not be opened, one that rejected us, one that locked us out, and one that accepted the bytes
#: but never answered.
CTL_WHY = ("ctl_pipe_absent", "ctl_connect_failed", "ctl_hello_timeout",
           "ctl_unauthorized", "ctl_lockout", "ctl_no_ack")

#: Lead ruling 3: a fix_params verdict must NAME the fix, not just classify the failure.
CTL_HOWTOFIX_UNAUTHORIZED = (
    "the instance record and the add-in disagree (or a bearer is required): check "
    "ctlPipe/ctlNonce/ctlProto in the instance record, and the VB6MCP_TOKEN / VB6MCP_TOKEN_FILE "
    "the ADD-IN process was started with")


class CtlError(Exception):
    """One ctl exchange failure.  why is always one of CTL_WHY (never a bare string)."""

    def __init__(self, why, detail="", stage="", retry_after_ms=None, response=None):
        Exception.__init__(self, "%s%s" % (why, ((": " + str(detail)) if detail else "")))
        self.why = why if why in CTL_WHY else "ctl_connect_failed"
        self.detail = str(detail or "")
        self.stage = stage
        self.retry_after_ms = retry_after_ms
        self.response = response or {}
        self.hello_ms = 0
        self.rtt_ms = 0
        self.sid = None


def ctl_problem(inst):
    """(why, detail) when the lane is NOT usable, else None.  FEATURE DETECT ONLY (sec 6)."""
    if inst is None:
        return "ctl_pipe_absent", "no_instance"
    if not CFG.ctl_enabled:
        return "ctl_pipe_absent", "disabled_by_env"
    if not getattr(inst, "ctlPipe", None):
        return "ctl_pipe_absent", "no_ctl_pipe_field"
    if not getattr(inst, "ctlNonce", None):
        return "ctl_pipe_absent", "no_ctl_nonce_field"
    if getattr(inst, "ctlAvailable", None) is False:
        return "ctl_pipe_absent", "ctl_available_false"
    proto = getattr(inst, "ctlProto", None)
    if proto and str(proto) != CTL_PROTO:
        # an unsupported proto is NOT usable and is never coerced into ctl/1
        return "ctl_pipe_absent", "ctl_proto_unsupported:%s" % proto
    return None


def ctl_available(inst):
    return ctl_problem(inst) is None


def _ctl_tag(value):
    """A nonce/secret-safe log tag: first 4 chars + ... (sec 2 factor 1: never log it in full)."""
    s = "" if value is None else str(value)
    return (s[:4] + "...") if len(s) > 4 else "..."


def ctl_sid4(sid):
    return _ctl_tag(sid)


class CtlClient(object):
    """One ctl/1 connection: hello -> welcome -> monotonic seq -> one response per frame.

    Reuses PipeClient (same np/1 framing, fully overlapped, every wait bounded), so the control
    lane cannot introduce a second, differently-behaved transport implementation.
    """

    def __init__(self, inst):
        self.inst = inst
        self.pc = None
        self.sid = None
        self.seq = None
        self.seq_base = 0
        self.ops = []
        self.auth_basis = []
        self.max_frame = 65536
        self.welcome = {}
        self.hello_ms = 0
        self.last_rtt_ms = 0
        self.ctl_tools = []

    # -- lifecycle ---------------------------------------------------------------
    def open(self, hello_ms=None):
        why, detail = ctl_problem(self.inst) or (None, None)
        if why:
            raise CtlError(why, detail, "detect")
        budget = max(200, int(hello_ms or CFG.ctl_hello_ms))
        t_hello = time.monotonic()
        pc = None
        for attempt in (0, 1):
            pc = PipeClient(self.inst.ctlPipe)
            try:
                pc.connect(min(CFG.pipe_connect_ms, budget))
                break
            except PipeError as exc:
                code = getattr(exc, "code", 0)
                if code in (ERROR_PIPE_BUSY, ERROR_SEM_TIMEOUT) and attempt == 0:
                    # sec 6: ONE retry after VB6MCP_PIPE_BUSY_RETRY_MS before reporting busy
                    bridge_log("ctl_connect_busy", pid=self.inst.pid,
                               retryMs=CFG.pipe_busy_retry_ms)
                    time.sleep(max(0, CFG.pipe_busy_retry_ms) / 1000.0)
                    continue
                if _pipe_reason(code) == "absent":
                    raise CtlError("ctl_pipe_absent", "connect_absent", "connect")
                raise CtlError("ctl_connect_failed",
                               "connect_%s_%d" % (_pipe_reason(code), code), "connect")
        self.pc = pc
        hello = {
            "vb6mcp": "ctlhello", "proto": CTL_PROTO,
            "nonce": self.inst.ctlNonce,
            "instanceId": "%s-%s" % (self.inst.pid, self.inst.ctlNonce),
            "client": {"pid": os.getpid(), "session": _session_id(),
                       "user": _user_name(), "version": VERSION},
        }
        if CFG.auth_enabled():
            # sec 2 optional fourth factor.  The token VALUE never reaches a log or a receipt;
            # tokenFp is an in-memory compare aid for the server and is never printed here.
            hello["auth"] = {"scheme": "bearer", "token": CFG.token,
                             "tokenFp": hashlib.sha256(
                                 CFG.token.encode("utf-8")).hexdigest()[:8]}
        bridge_log("ctl_hello", pid=self.inst.pid, nonce=_ctl_tag(self.inst.ctlNonce),
                   bearer=bool(CFG.auth_enabled()), proto=CTL_PROTO)
        deadline = t_hello + budget / 1000.0
        sent, _n = self.pc.send_frame(hello, deadline)
        if not sent:
            self.close()
            raise CtlError("ctl_connect_failed", "hello_write_failed", "hello")
        frame, err = self.pc.recv_frame(deadline)
        if err == "timeout":
            self.close()
            e = CtlError("ctl_hello_timeout", "no_welcome_in_%dms" % budget, "hello")
            e.hello_ms = _msf(t_hello)
            raise e
        if err is not None:
            self.close()
            raise CtlError("ctl_connect_failed",
                           "hello_read_%s" % getattr(err, "code", err), "hello")
        self.hello_ms = _msf(t_hello)
        self._validate_welcome(frame)
        return self

    def close(self):
        if self.pc is not None:
            try:
                self.pc.close()
            except Exception:
                pass
        self.pc = None

    # -- handshake validation ----------------------------------------------------
    def _validate_welcome(self, frame):
        kind = str(frame.get("vb6mcp") or "")
        if kind == "deny":
            self.close()
            raise _ctl_deny_error(frame, "hello")
        if kind != "ctlwelcome":
            self.close()
            raise CtlError("ctl_connect_failed",
                           "unexpected_first_frame:%s" % (kind or "<empty>"), "hello",
                           response=frame)
        proto = str(frame.get("proto") or "")
        if proto != CTL_PROTO:
            self.close()
            raise CtlError("ctl_connect_failed", "welcome_proto:%s" % (proto or "<empty>"),
                           "hello", response=frame)
        want = "%s-%s" % (self.inst.pid, self.inst.ctlNonce)
        if str(frame.get("instanceId") or "") != want:
            self.close()
            raise CtlError("ctl_unauthorized", "welcome_instance_mismatch", "hello",
                           response=frame)
        try:
            wpid = int(frame.get("pid"))
        except (TypeError, ValueError):
            wpid = -1
        if wpid != int(self.inst.pid or 0):
            self.close()
            raise CtlError("ctl_unauthorized", "welcome_pid_mismatch", "hello", response=frame)
        sid = frame.get("sid")
        if not isinstance(sid, str) or not sid:
            self.close()
            raise CtlError("ctl_connect_failed", "welcome_without_sid", "hello", response=frame)
        try:
            base = int(frame.get("seqBase") or 0)
        except (TypeError, ValueError):
            base = 0
        self.sid = sid
        self.seq_base = base
        # "a strictly increasing seq (starting at seqBase)": the FIRST frame is seqBase itself.
        self.seq = base - 1
        self.welcome = frame
        ops = frame.get("ops")
        self.ops = [str(o) for o in ops] if isinstance(ops, list) else []
        basis = frame.get("authBasis")
        self.auth_basis = [str(b) for b in basis] if isinstance(basis, list) else []
        tools = frame.get("ctlTools")
        self.ctl_tools = [str(t) for t in tools] if isinstance(tools, list) else []
        try:
            self.max_frame = int(frame.get("maxFrame") or 65536)
        except (TypeError, ValueError):
            self.max_frame = 65536
        bridge_log("ctl_welcome", pid=wpid, sid=_ctl_tag(sid), ops=",".join(self.ops) or "-",
                   authBasis=",".join(self.auth_basis) or "-", helloMs=self.hello_ms)

    # -- one op ------------------------------------------------------------------
    def call(self, op, args=None, ack_ms=None):
        """Send one op and return the ctlresult dict (ok:false refusals are RETURNED, not raised).

        Raised: every transport/auth failure (CtlError).  An op-level refusal such as
        nothing_in_flight / readonly_lane_refused is an ANSWER and is returned to the caller.
        """
        if self.pc is None:
            raise CtlError("ctl_connect_failed", "not_open", "op")
        if op not in CTL_OPS:
            raise CtlError("ctl_connect_failed", "op_not_in_protocol:%s" % op, "op")
        if self.ops and op not in self.ops:
            raise CtlError("ctl_connect_failed", "op_not_published:%s" % op, "op")
        self.seq += 1
        frame = {"vb6mcp": "ctl", "sid": self.sid, "seq": self.seq, "op": op}
        if args:
            frame["args"] = args
        budget = max(100, int(ack_ms or CFG.ctl_ack_ms))
        t0 = time.monotonic()
        deadline = t0 + budget / 1000.0
        try:
            encoded = encode_frame(frame)
        except FrameError:
            raise CtlError("ctl_connect_failed", "frame_over_maxFrame_op", "op")
        if len(encoded) - 4 > self.max_frame:
            raise CtlError("ctl_connect_failed", "frame_over_maxFrame:%d>%d"
                           % (len(encoded) - 4, self.max_frame), "op")
        sent, _n = self.pc.write_all(encoded, deadline)
        if not sent:
            raise CtlError("ctl_connect_failed", "op_write_failed:%s" % op, "op")
        frame, err = self.pc.recv_frame(deadline)
        if err == "timeout":
            e = CtlError("ctl_no_ack", "no_result:%s" % op, "op")
            e.hello_ms = self.hello_ms
            e.rtt_ms = _msf(t0)
            e.sid = self.sid
            raise e
        if err is not None:
            raise CtlError("ctl_connect_failed",
                           "op_read_failed:%s" % getattr(err, "code", err), "op")
        self.last_rtt_ms = _msf(t0)
        kind = str(frame.get("vb6mcp") or "")
        if kind == "deny":
            raise _ctl_deny_error(frame, "op")
        if kind != "ctlresult":
            raise CtlError("ctl_connect_failed",
                           "unexpected_op_frame:%s" % (kind or "<empty>"), "op", response=frame)
        try:
            seq = int(frame.get("seq"))
        except (TypeError, ValueError):
            seq = -1
        if str(frame.get("sid") or "") != str(self.sid):
            raise CtlError("ctl_unauthorized", "result_sid_mismatch", "op", response=frame)
        if seq != self.seq:
            raise CtlError("ctl_unauthorized", "result_seq_mismatch:%s!=%s" % (seq, self.seq),
                           "op", response=frame)
        if frame.get("ok") is False:
            reason = str(frame.get("reason") or "")
            if reason in ("ctl_unauthorized", "not_authed", "ctl_lockout"):
                raise _ctl_deny_error(frame, "op")
            # an op-level refusal (nothing_in_flight / readonly_lane_refused / unknown op)
            return frame
        return frame


def _msf(t0):
    return int((time.monotonic() - t0) * 1000)


def _ctl_deny_error(frame, stage):
    """Map a structured ctl rejection to one of the six reasons (sec 5).  No secret is echoed."""
    reason = str(frame.get("reason") or "")
    detail = str(frame.get("detail") or "")
    if reason == "ctl_lockout":
        try:
            retry = int(frame.get("retryAfterMs") or 60000)
        except (TypeError, ValueError):
            retry = 60000
        return CtlError("ctl_lockout", detail or "lockout", stage, retry_after_ms=retry,
                        response=frame)
    if reason in ("ctl_unauthorized", "not_authed"):
        return CtlError("ctl_unauthorized", detail or reason, stage, response=frame)
    if reason == "hello_timeout":
        return CtlError("ctl_hello_timeout", detail or reason, stage, response=frame)
    if reason in ("bad_hello", "ctl_bad_frame"):
        return CtlError("ctl_unauthorized", "%s:%s" % (reason, detail or "-"), stage,
                        response=frame)
    return CtlError("ctl_connect_failed", "%s:%s" % (reason or "deny", detail or "-"), stage,
                    response=frame)


def ctl_request(inst, op, args=None, ack_ms=None):
    """One complete ctl exchange.  Returns (result_dict, info).  Raises CtlError.

    info is the client-side evidence an oob report is built from: helloMs (connect+hello+welcome),
    rttMs (one op round trip), sid (for the log only), authBasis and the published op set.
    """
    ctl = CtlClient(inst)
    try:
        ctl.open()
        res = ctl.call(op, args, ack_ms)
        info = {"helloMs": ctl.hello_ms, "rttMs": ctl.last_rtt_ms, "sid": ctl.sid,
                "authBasis": list(ctl.auth_basis), "ops": list(ctl.ops),
                "ctlTools": list(ctl.ctl_tools), "seq": ctl.seq}
        return res, info
    finally:
        ctl.close()


# -- oob: the client-side delivery report (sec 6 + sec 7) ---------------------------------
# The server never emits these four fields (it cannot observe a frame it never received);
# THEY BELONG TO THE CLIENT.  A receipt claims an effect only when oob.acked is true.

def ctl_oob_ok(op, info, applied):
    # The SHAPE is uniform with the failure case (peerState/suggestedAction present but null):
    # a consumer can read oob["peerState"] without branching on acked first.
    return {"lane": "ctl", "attempted": True, "acked": True, "why": None, "ackOp": op,
            "helloMs": int(info.get("helloMs") or 0), "rttMs": int(info.get("rttMs") or 0),
            "sid": ctl_sid4(info.get("sid")), "authBasis": list(info.get("authBasis") or []),
            "delivered": "acked", "applied": bool(applied),
            "peerState": None, "suggestedAction": None}


def ctl_oob_error(exc, inst):
    """The client-side verdict for an attempt that was NOT acknowledged.  Six grades, kept apart."""
    oob = {"lane": "ctl", "attempted": True, "acked": False, "why": exc.why, "ackOp": None,
           "helloMs": int(getattr(exc, "hello_ms", 0) or 0),
           "rttMs": int(getattr(exc, "rtt_ms", 0) or 0),
           "sid": ctl_sid4(getattr(exc, "sid", None)),
           "detail": exc.detail or None,
           "delivered": "queued_unacknowledged", "applied": False,
           "peerState": "peer_unresponsive", "suggestedAction": "kill_and_restart"}
    alive = bool(getattr(inst, "pidAlive", False))
    if exc.why in ("ctl_lockout", "ctl_unauthorized"):
        # The peer ANSWERED.  It is not unresponsive: it is refusing us for a stated reason, so
        # "kill_and_restart" would be the wrong remedy and would hide a registry/auth mismatch.
        oob["delivered"] = "rejected_by_server"
        oob["peerState"] = "peer_busy"
        if exc.why == "ctl_lockout":
            oob["suggestedAction"] = "retry_once"
            oob["retryAfterMs"] = int(exc.retry_after_ms or 60000)
        else:
            oob["suggestedAction"] = "fix_params"
            oob["howToFix"] = CTL_HOWTOFIX_UNAUTHORIZED
    elif exc.why == "ctl_pipe_absent":
        oob["delivered"] = "not_attempted"
        oob["peerState"] = "noIde"
        oob["suggestedAction"] = "ide_start"
    elif not alive:
        oob["peerState"] = "noIde"
        oob["suggestedAction"] = "ide_start"
    return oob


# -- inflight registry: what this bridge currently has on the wire -------------------------
# A ctl cancel is COOPERATIVE (sec 4.4).  When it is acknowledged, the bridge records the request
# against the call still on the wire, so that call's eventual failure receipt keeps
# cancelled_by_request (with timeoutAlsoDue when the budget expired first) instead of being
# flattened into a plain timeout - the caller must be able to tell "I cancelled" from "it died".
_INFLIGHT = {}
_INFLIGHT_LOCK = threading.Lock()
_INFLIGHT_SEQ = [0]


def inflight_begin(tool, rid):
    with _INFLIGHT_LOCK:
        _INFLIGHT_SEQ[0] += 1
        token = _INFLIGHT_SEQ[0]
        _INFLIGHT[token] = {"tool": tool, "rid": rid, "t0": time.monotonic(),
                            "cancelRequested": False, "cancelledAt": None, "pid": None}
        return token


def inflight_end(token):
    with _INFLIGHT_LOCK:
        _INFLIGHT.pop(token, None)


def inflight_cancelled(token):
    with _INFLIGHT_LOCK:
        rec = _INFLIGHT.get(token)
        return bool(rec and rec.get("cancelRequested"))


def inflight_mark_cancel(target):
    """Record an ACKED cancel against every matching inflight call.  Returns the tools matched."""
    hit = []
    tgt = str(target if target not in (None, "") else "*").strip()
    with _INFLIGHT_LOCK:
        for rec in _INFLIGHT.values():
            if tgt == "*" or tgt == str(rec.get("tool") or ""):
                rec["cancelRequested"] = True
                rec["cancelledAt"] = time.monotonic()
                if rec.get("tool") not in hit:
                    hit.append(rec.get("tool"))
    return hit


def inflight_snapshot():
    with _INFLIGHT_LOCK:
        return [{"tool": r.get("tool"), "rid": r.get("rid"),
                 "forMs": int((time.monotonic() - r.get("t0")) * 1000),
                 "cancelRequested": bool(r.get("cancelRequested"))} for r in _INFLIGHT.values()]


# -- routing the two EXISTING ide_state intents onto the lane ------------------------------
def ctl_intent_of(req_obj):
    """(op, args, intentName) for an existing tool parameter, else None.

    SEC 6: the lane adds NO new tool parameters.  Exactly two existing intents are routed:
      ide_state{release_inflight:true}      -> release_gate
      ide_state{cancel_inflight:true}       -> cancel
    """
    if not isinstance(req_obj, dict) or req_obj.get("method") != "tools/call":
        return None
    params = req_obj.get("params") if isinstance(req_obj.get("params"), dict) else {}
    if params.get("name") != "ide_state":
        return None
    args = params.get("arguments") if isinstance(params.get("arguments"), dict) else {}
    if args.get("cancel_inflight") is True:
        target = args.get("cancel_target")
        return ("cancel", {"target": (str(target) if target else "*")}, "cancel_inflight")
    if args.get("release_inflight") is True:
        try:
            min_held = int(args.get("release_min_held_ms") or 0)
        except (TypeError, ValueError):
            min_held = 0
        return ("release_gate", {"minHeldMs": max(0, min_held)}, "release_inflight")
    return None


def ctl_intent_receipt(rid, inst, tool, intent):
    """Route one intent through the lane and build the receipt.  Returns (text, isError, meta).

    An UNAVAILABLE lane is NOT handled here: the caller falls back to the pre-existing business
    path unchanged (sec 6: only fall back to the business lane when the lane is not available).
    """
    op, op_args, intent_name = intent
    base = {"tool": tool, "intent": intent_name, "lane": "ctl", "source": "ctl_lane",
            "ctlOp": op, "proto": CTL_PROTO}
    try:
        res, info = ctl_request(inst, op, op_args)
    except CtlError as exc:
        oob = ctl_oob_error(exc, inst)
        body = dict(base)
        body.update({
            "ok": False,
            "reason": exc.why,
            "message": str(exc),
            "ctlDetail": exc.detail or None,
            "oob": oob,
            "ctlAvailable": True,
            "failureKind": oob["peerState"],
            "failureKindRule": FAILURE_KIND_RULE,
            "peerState": oob["peerState"],
            "suggestedAction": oob["suggestedAction"],
            "delivered": oob["delivered"],
            "applied": False,
        })
        if oob.get("retryAfterMs"):
            body["retryAfterMs"] = oob["retryAfterMs"]
        if exc.why == "ctl_unauthorized":
            # the SAME text as oob.howToFix (one source, two places a caller may look)
            body["howToFix"] = oob.get("howToFix") or CTL_HOWTOFIX_UNAUTHORIZED
        body["note"] = ("the control lane did not acknowledge this intent; nothing here is a "
                        "claimed effect (oob.acked=false).  A VB6 statement is never preempted, and "
                        "a call that never returns to the add-in's pump can neither be cancelled nor "
                        "served (ctl/1 sec 7).")
        bridge_log("ctl_intent_unacked", tool=tool, intent=intent_name, why=exc.why,
                   detail=exc.detail, pid=getattr(inst, "pid", None))
        return _jsonrpc_result(rid, error_result(body, True)), True, {
            "ctl": True, "lane": "ctl", "why": exc.why, "acked": False, "tool": tool}
    ok = res.get("ok") is not False
    if op == "release_gate":
        applied = bool(res.get("released"))
    elif op == "cancel":
        applied = bool(res.get("requested"))
    else:
        applied = ok
    if applied and op == "cancel":
        inflight_mark_cancel(op_args.get("target"))
    oob = ctl_oob_ok(op, info, applied)
    body = dict(base)
    body["ok"] = ok
    body["oob"] = oob
    for key in CTL_RESULT_PASS:
        if key in res:
            body[key] = res[key]
    if "limits" in res:
        # sec 4 preamble: relayed VERBATIM from the server, never synthesised here.
        body["limits"] = res["limits"]
    if not ok:
        body["reason"] = res.get("reason")
        body["ctlReason"] = res.get("reason")
        body["ctlDetail"] = res.get("detail")
        if "inFlightTool" in res:
            body["inFlightTool"] = res["inFlightTool"]
    bridge_log("ctl_intent", tool=tool, intent=intent_name, op=op, ok=ok, applied=applied,
               ackOp=op, helloMs=info.get("helloMs"), rttMs=info.get("rttMs"))
    return _jsonrpc_result(rid, error_result(body, not ok)), (not ok), {
        "ctl": True, "lane": "ctl", "why": None, "acked": True, "tool": tool,
        "applied": applied}


def ctl_cancel_override(body, token, inst=None):
    """A receipt for a call whose cancel was ACKED but not honoured before its budget expired.

    Keeps cancelled_by_request visible (the caller asked for it) AND states that the wire budget
    also expired (timeoutAlsoDue) - so it is never a silent timeout and never a false "stopped".
    """
    if not inflight_cancelled(token):
        return body
    b = dict(body)
    prev = b.get("reason")
    b["cancelRequested"] = True
    b["cancelHonored"] = False
    # Only a WEDGE-shaped failure is relabelled: the caller asked to cancel a call that never
    # came back, so cancelled_by_request is the more informative verdict - but the budget facts
    # are KEPT under timeoutAlsoDue/timeoutReason, never overwritten.
    # Any other failure (deny, refused, ambiguous) already says what happened; the cancel is
    # attached as an additional fact and its reason is NOT replaced.
    if prev in ("timeout", "pipe_stalled"):
        b["timeoutAlsoDue"] = (prev == "timeout")
        b["timeoutReason"] = prev
        b["reasonLegacy"] = b.get("reasonLegacy") or (
            "forward_timeout" if prev == "timeout" else prev)
        b["reason"] = "cancelled_by_request"
        b["cancelNote"] = ("cancel was acknowledged on the control lane but a VB6 statement is "
                           "never preempted (ctl/1 sec 7): the call never reached a pump point, so "
                           "the original budget failure also stands")
    else:
        b["timeoutAlsoDue"] = False
        b["cancelNote"] = ("cancel was acknowledged on the control lane; the failure below has a "
                           "different cause and its reason is preserved")
    return b


# ---------------------------------------------------------------- out-of-process ----
# Section B of the R-1021-03 client half: the LAST-resort read-only fallback.
#
# It is used ONLY when (a) the ctl lane is UNAVAILABLE for this instance and (b) the business
# forward was classified peer_unresponsive (there is a pid, but nothing answers).  In that state
# six read-only tools can still be answered from PROCESS-LEVEL facts that need no cooperation
# from the add-in at all: EnumWindows/EnumChildWindows + IsHungAppWindow.  This measures; it
# never drives, never sends input, and never fabricates an add-in receipt: every receipt says
# source:"out_of_process" + degraded:true, and a fact that could not be read is listed in
# missing[] with the reason instead of being guessed.
user32 = ctypes.WinDLL("user32", use_last_error=True)
ENUM_PROC = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)
GW_OWNER = 4

user32.EnumWindows.argtypes = [ENUM_PROC, wintypes.LPARAM]
user32.EnumWindows.restype = wintypes.BOOL
user32.EnumChildWindows.argtypes = [wintypes.HWND, ENUM_PROC, wintypes.LPARAM]
user32.EnumChildWindows.restype = wintypes.BOOL
user32.GetWindowThreadProcessId.argtypes = [wintypes.HWND, ctypes.POINTER(wintypes.DWORD)]
user32.GetWindowThreadProcessId.restype = wintypes.DWORD
user32.GetWindowTextLengthW.argtypes = [wintypes.HWND]
user32.GetWindowTextLengthW.restype = ctypes.c_int
user32.GetWindowTextW.argtypes = [wintypes.HWND, wintypes.LPWSTR, ctypes.c_int]
user32.GetWindowTextW.restype = ctypes.c_int
user32.GetClassNameW.argtypes = [wintypes.HWND, wintypes.LPWSTR, ctypes.c_int]
user32.GetClassNameW.restype = ctypes.c_int
user32.IsWindowVisible.argtypes = [wintypes.HWND]
user32.IsWindowVisible.restype = wintypes.BOOL
user32.IsWindowEnabled.argtypes = [wintypes.HWND]
user32.IsWindowEnabled.restype = wintypes.BOOL
user32.IsHungAppWindow.argtypes = [wintypes.HWND]
user32.IsHungAppWindow.restype = wintypes.BOOL
user32.GetForegroundWindow.argtypes = []
user32.GetForegroundWindow.restype = wintypes.HWND
user32.GetWindow.argtypes = [wintypes.HWND, wintypes.DWORD]
user32.GetWindow.restype = wintypes.HWND
user32.DestroyWindow.argtypes = [wintypes.HWND]
user32.DestroyWindow.restype = wintypes.BOOL

OOP_TOOLS = ("ide_state", "execution_state", "read_active_dialogs", "find_gui", "ui_text",
             "last_error")
DIALOG_CLASS = "#32770"


def _hwnd_i(v):
    try:
        return int(v)
    except (TypeError, ValueError):
        return 0


def _win_text(hwnd):
    n = int(user32.GetWindowTextLengthW(hwnd) or 0)
    if n <= 0:
        return ""
    buf = ctypes.create_unicode_buffer(n + 1)
    got = int(user32.GetWindowTextW(hwnd, buf, n + 1) or 0)
    return buf.value[:got] if got > 0 else ""


def _win_class(hwnd):
    buf = ctypes.create_unicode_buffer(256)
    got = int(user32.GetClassNameW(hwnd, buf, 256) or 0)
    return buf.value[:got] if got > 0 else ""


def _win_node(hwnd, fg=0):
    h = _hwnd_i(hwnd)
    owner = _hwnd_i(user32.GetWindow(h, GW_OWNER))
    node = {"hwnd": h, "class": _win_class(hwnd), "title": _win_text(hwnd),
            "visible": bool(user32.IsWindowVisible(hwnd)),
            "enabled": bool(user32.IsWindowEnabled(hwnd)),
            "hung": bool(user32.IsHungAppWindow(hwnd)),
            "foreground": bool(fg and h == fg),
            "owner": owner or None,
            "ownerEnabled": (bool(user32.IsWindowEnabled(owner)) if owner else None)}
    return node


def _enum_top_windows(pid):
    """Top-level windows owned by pid.  Read-only; no message is ever sent to the target."""
    found = []

    @ENUM_PROC
    def _cb(hwnd, _lp):
        try:
            wpid = wintypes.DWORD()
            user32.GetWindowThreadProcessId(hwnd, ctypes.byref(wpid))
            if int(wpid.value) == int(pid):
                found.append(_hwnd_i(hwnd))
        except Exception:
            pass
        return True

    ok = bool(user32.EnumWindows(_cb, 0))
    return found, ok


def _enum_children(hwnd):
    found = []

    @ENUM_PROC
    def _cb(child, _lp):
        try:
            found.append(_hwnd_i(child))
        except Exception:
            pass
        return True

    ok = bool(user32.EnumChildWindows(hwnd, _cb, 0))
    return found, ok


def _proc_session(pid):
    try:
        kernel32.ProcessIdToSessionId.restype = wintypes.BOOL
        kernel32.ProcessIdToSessionId.argtypes = [wintypes.DWORD,
                                                  ctypes.POINTER(wintypes.DWORD)]
        sid = wintypes.DWORD()
        if kernel32.ProcessIdToSessionId(int(pid), ctypes.byref(sid)):
            return sid.value
    except Exception:
        pass
    return None


def oop_measure(pid):
    """(facts, missing[]) measured for ONE pid.  Every unreadable fact is NAMED, never guessed."""
    facts = {"pid": pid, "pidAlive": pid_alive(pid), "session": _proc_session(pid),
             "ourSession": _session_id(), "windows": [], "dialogs": [],
             "hung": None, "foregroundHwnd": _hwnd_i(user32.GetForegroundWindow()),
             "enumOk": False}
    missing = []
    if not facts["pidAlive"]:
        missing.append({"what": "process", "why": "pid_not_alive"})
        return facts, missing
    if (facts["session"] is not None and facts["ourSession"]
            and facts["session"] != facts["ourSession"]):
        # EnumWindows only sees the CALLING desktop; a window station/session boundary is a
        # hard limit of out-of-process observation, not an error.
        missing.append({"what": "windows", "why": "cross_session_windows_invisible",
                        "targetSession": facts["session"], "ourSession": facts["ourSession"]})
        return facts, missing
    hwnds, ok = _enum_top_windows(pid)
    facts["enumOk"] = ok
    if not hwnds:
        missing.append({"what": "windows", "why": "no_top_level_windows"})
        return facts, missing
    fg = facts["foregroundHwnd"]
    nodes = [_win_node(h, fg) for h in hwnds]
    facts["windows"] = nodes
    hung = [n["hung"] for n in nodes]
    if hung:
        facts["hung"] = bool(all(hung)) if len(hung) > 1 else bool(hung[0])
    for node in nodes:
        if node["class"] != DIALOG_CLASS:
            continue
        kids, kok = _enum_children(node["hwnd"])
        buttons = []
        texts = []
        for k in kids:
            kn = _win_node(k)
            if kn["class"].lower() == "button":
                buttons.append({"hwnd": kn["hwnd"], "text": kn["title"],
                                "enabled": kn["enabled"], "visible": kn["visible"]})
            elif kn["class"].lower() in ("static", "edit"):
                texts.append({"hwnd": kn["hwnd"], "class": kn["class"], "text": kn["title"]})
        readable = [b for b in buttons if b["text"]] + [t for t in texts if t["text"]]
        if kids and not readable:
            # a control exists but no text came back: UIPI (elevated target) or a themed owner
            # draw.  Say so explicitly instead of reporting an empty dialog body.
            missing.append({"what": "dialog_text", "hwnd": node["hwnd"],
                            "why": "child_text_unavailable"})
        node = dict(node)
        node["childrenOk"] = kok
        node["buttons"] = buttons
        node["texts"] = texts
        facts["dialogs"].append(node)
    return facts, missing


def _oop_base(tool, why, facts, missing):
    return {
        "ok": bool(facts.get("pidAlive")),
        "tool": tool,
        "source": "out_of_process",
        "degraded": True,
        "addInReceipt": False,
        "oob": {"lane": "oop", "attempted": True, "acked": False, "why": why,
                "delivered": "not_delivered", "applied": False,
                "peerState": "peer_unresponsive", "suggestedAction": "kill_and_restart"},
        "measured": facts,
        "missing": missing,
        "note": ("the add-in did not answer, so this receipt is built ONLY from process-level "
                 "measurements (EnumWindows/IsHungAppWindow).  No add-in receipt was produced and "
                 "none is claimed; a fact that could not be read is listed in missing[]."),
    }


def oop_body(tool, args, why, body):
    """The tool-shaped body for one of OOP_TOOLS (or None when tool is not served here)."""
    args = args or {}
    pid = _as_int(args.pop("__pid", None))
    if not pid:
        return None
    facts, missing = oop_measure(pid)
    out = _oop_base(tool, why, facts, missing)
    under = body or {}
    out["failureUnderneath"] = {"reason": under.get("reason"),
                                "reasonLegacy": under.get("reasonLegacy"),
                                "elapsedMs": under.get("elapsedMs"),
                                "maxMs": under.get("maxMs"),
                                "maxMsSource": under.get("maxMsSource"),
                                "failureKind": under.get("failureKind"),
                                "peerState": under.get("peerState")}
    wins = facts.get("windows") or []
    if tool in ("ide_state", "execution_state"):
        main = None
        for w in wins:
            if w["visible"] and w["title"]:
                main = w
                break
        out["pid"] = pid
        out["hung"] = facts.get("hung")
        out["mode"] = None
        out["modeNote"] = "the IDE mode lives inside the add-in/engine and is NOT observable from outside"
        out["mainWindow"] = main
        out["windowCount"] = len(wins)
        out["dialogCount"] = len(facts.get("dialogs") or [])
        if tool == "execution_state":
            out["topLevelWindows"] = wins
            out["modalDialogs"] = [{"hwnd": d["hwnd"], "title": d["title"],
                                    "buttons": [b["text"] for b in d["buttons"]]}
                                   for d in (facts.get("dialogs") or [])]
            out["runningState"] = ("hung" if facts.get("hung") else
                                   ("unknown" if not wins else "responsive_windows"))
            out["runningStateBasis"] = "IsHungAppWindow on the pid's top-level windows"
    elif tool == "read_active_dialogs":
        out["dialogs"] = facts.get("dialogs") or []
        out["count"] = len(out["dialogs"])
    elif tool == "find_gui":
        kw = str(args.get("keyword") or "").lower()
        cls = str(args.get("cls") or "")
        showing = args.get("showing")
        hits = []
        for w in wins:
            if kw and kw not in (w["title"] or "").lower() and kw not in (w["class"] or "").lower():
                continue
            if cls and w["class"] != cls:
                continue
            if showing is True and not w["visible"]:
                continue
            if showing is False and w["visible"]:
                continue
            hits.append(w)
        for d in (facts.get("dialogs") or []):
            for k in (d.get("buttons") or []) + (d.get("texts") or []):
                if kw and kw not in (k.get("text") or "").lower():
                    continue
                hits.append({"hwnd": k["hwnd"], "class": k.get("class") or "Button",
                             "title": k.get("text") or "", "visible": k.get("visible", True),
                             "enabled": k.get("enabled", True), "hung": False,
                             "child": True, "dialog": d["hwnd"]})
        out["keyword"] = args.get("keyword") or ""
        out["cls"] = cls or ""
        out["showing"] = showing
        out["count"] = len(hits)
        out["hits"] = hits
        out["ok"] = bool(facts.get("pidAlive"))
    elif tool == "ui_text":
        want = _as_int(args.get("hwnd"))
        cls = str(args.get("cls") or "")
        target = None
        if want:
            target = _win_node(want)
        else:
            for w in wins:
                if cls and w["class"] == cls:
                    target = w
                    break
            if target is None and wins:
                target = wins[0]
        if target is None:
            out["ok"] = False
            out["reason"] = "window_not_found"
            out["text"] = ""
        else:
            kids, _kok = _enum_children(target["hwnd"])
            child_nodes = [_win_node(k) for k in kids]
            pieces = [target["title"]] + [c["title"] for c in child_nodes if c["title"]]
            out["hwnd"] = target["hwnd"]
            out["class"] = target["class"]
            out["text"] = "\n".join([p for p in pieces if p])
            out["source"] = "wm_gettext"
            out["children"] = child_nodes
            if not any(c["title"] for c in child_nodes) and child_nodes:
                missing.append({"what": "child_text", "why": "child_text_unavailable"})
    elif tool == "last_error":
        dlg = None
        for d in (facts.get("dialogs") or []):
            alltext = " ".join([d["title"]] + [t.get("text") or "" for t in
                                               (d.get("texts") or [])])
            if "run-time error" in alltext.lower() or "runtime error" in alltext.lower():
                dlg = d
                break
        if dlg is None:
            out["ok"] = bool(facts.get("pidAlive"))
            out["errorDialog"] = {"present": False}
            out["errNumber"] = None
            out["errDescription"] = None
        else:
            alltext = "\n".join([dlg["title"]] + [t.get("text") or ""
                                                  for t in (dlg.get("texts") or [])])
            out["errorDialog"] = {"present": True, "hwnd": dlg["hwnd"], "title": dlg["title"],
                                  "body": alltext,
                                  "buttons": [b["text"] for b in dlg["buttons"]],
                                  "bodySource": "wm_gettext"}
            out["errNumber"] = None
            out["errDescription"] = None
            m = re.search(r"run-?time error '?(-?\d+)'?", alltext, re.IGNORECASE)
            if m:
                out["errNumber"] = int(m.group(1))
            dm = re.search(r"run-?time error '?-?\d+'?\s*:\s*(.+)", alltext, re.IGNORECASE)
            if dm:
                out["errDescription"] = dm.group(1).strip().split("\n")[0][:400]
            out["parsed"] = bool(out["errNumber"] is not None)
            out["parseBasis"] = "regex over the measured dialog text (not an add-in field)"
            if not alltext.strip():
                missing.append({"what": "error_text", "why": "child_text_unavailable"})
    return out


def oop_fallback_body(tool, args, sel, body):
    """Build the out-of-process receipt, or None when the fallback does not apply.

    Gate (Lead ruling): ONLY when the ctl lane is unavailable AND the peer was classified
    peer_unresponsive.  A live ctl lane is never bypassed by this path.
    """
    if not CFG.oop_enabled or tool not in OOP_TOOLS:
        return None
    body = body or {}
    if body.get("peerState") != "peer_unresponsive":
        return None
    # WEDGE-SHAPED failures only.  A DENY / handshake failure also classifies as
    # peer_unresponsive (a peer was found and did not answer the call), but there the add-in
    # actually SPOKE - the denial itself is the diagnosis - so measuring windows would bury the
    # real reason (and would have silently changed the S23 receipt).  Only a call that was
    # written and never answered is eligible.
    if str(body.get("reason") or "") not in ("timeout", "pipe_stalled"):
        return None
    if body.get("ctlAvailable"):
        return None
    inst, _ref = select_target(sel)
    pid = getattr(inst, "pid", None) or _as_int((body.get("heldBy") or {}).get("pid"))
    if not pid:
        return None
    why = "ctl_no_ack" if _has_ctl(inst) else "ctl_unavailable"
    a = dict(args or {})
    a["__pid"] = pid
    out = oop_body(tool, a, why, body)
    if out is None:
        return None
    out["pid"] = pid
    return out


# -- the ADDITIVE second trigger (Lead ruling 2) -------------------------------------------
# A wedged host is one case; a host-modal box (VB6's own Make dialog, a runtime-error box, a
# host-modal loop) is another: business AND ctl both stop answering, yet the process is alive
# and the window tree shows exactly what is stuck.  So there is a SECOND trigger that only
# APPENDS a measured segment - it never replaces the failure fields, and it is never used to
# re-diagnose a peer that SPOKE (a deny/ctl_unauthorized keeps its own reason verbatim).
OOP_ADDENDUM_STATES = ("peer_busy", "peer_unresponsive")
OOP_ADDENDUM_WHY = ("ctl_unavailable", "ctl_unauthorized")


def oop_addendum_eligible(body, inst):
    """("ctl_unavailable"|"ctl_unauthorized"|None) - whether the additive segment may be built.

    Pure decision (no measurement), so tools/check_failure_envelope.py can assert every branch.
    """
    if not CFG.oop_enabled:
        return None
    body = body or {}
    if body.get("outOfProcess"):
        return None
    if body.get("peerState") not in OOP_ADDENDUM_STATES:
        return None
    # The PRIMARY replacement path owns the wedge shape (timeout / pipe_stalled); the additive
    # segment must not duplicate it.
    if str(body.get("reason") or "") in ("timeout", "pipe_stalled"):
        return None
    # A peer that ANSWERED (deny / handshake refusal) is NEVER re-diagnosed: its reason is the
    # diagnosis and the receipt stays verbatim.  This is the S23 lesson as a rule.
    if body.get("serverReason") or body.get("serverDetail"):
        return None
    ctl_why = ((body.get("oob") or {}).get("why"))
    if ctl_why in ("ctl_unauthorized", "ctl_lockout"):
        return "ctl_unauthorized"
    if ctl_why == "ctl_pipe_absent":
        return "ctl_unavailable"
    if _has_ctl(inst):
        # the lane is published and usable: an operator can ask it directly, so measuring the
        # window tree is not needed (and would be a second opinion on a live control lane).
        return None
    return "ctl_unavailable"


def oop_addendum(tool, args, sel, body):
    """The additive outOfProcess segment, or None.  NEVER touches body's own fields.

    Evidence is REQUIRED: a #32770 dialog is visible, or IsHungAppWindow says the pid is not
    responding.  Without evidence nothing is appended - an unobserved claim is worse than none.
    """
    inst, _ref = select_target(sel)
    why = oop_addendum_eligible(body, inst)
    if why is None:
        return None
    pid = getattr(inst, "pid", None) or _as_int((body or {}).get("heldBy", {}).get("pid"))
    if not pid or not pid_alive(pid):
        return None
    facts, missing = oop_measure(pid)
    dialogs = [d for d in (facts.get("dialogs") or []) if d.get("class") == DIALOG_CLASS]
    hung = facts.get("hung") is True
    if not dialogs and not hung:
        return None
    return {
        "source": "out_of_process",
        "degraded": True,
        "addInReceipt": False,
        "additive": True,
        "why": why,
        "tool": tool,
        "evidence": {"modalDialogs": len(dialogs), "hung": hung,
                     "dialogTitles": [d.get("title") for d in dialogs][:4],
                     "basis": ("a #32770 modal dialog is visible" if dialogs else
                               "IsHungAppWindow reports the pid is not responding")},
        "measured": facts,
        "missing": missing,
        "note": ("ADDITIVE process-level observation only (source=out_of_process, degraded): "
                 "the failure fields above are the middle service's own verdict and were NOT "
                 "rewritten - a peer that spoke is never re-diagnosed, and a wedged host is "
                 "reported HERE, not instead of the real reason"),
    }


# ---------------------------------------------------------------- registry --------

def read_json(path, default):
    try:
        with open(path, "r", encoding="utf-8") as fh:
            return json.load(fh)
    except Exception:
        return default


def normalize_path(p):
    """The hub's path normalization, restated (case handled by the caller)."""
    s = ("" if p is None else str(p)).strip()
    if not s:
        return ""
    s = s.replace("/", "\\")
    unc = s.startswith("\\\\")
    body = s[2:] if unc else s
    while "\\\\" in body:
        body = body.replace("\\\\", "\\")
    s = ("\\\\" if unc else "") + body
    while s.endswith("\\"):
        s = s[:-1]
    return s


class Instance(object):
    """One registry entry, enriched with the pipe identity verdict (ADR sec 3.3)."""

    def __init__(self, entry):
        self.file = entry.get("file")
        self.pid = entry.get("pid")
        self.port = entry.get("port")
        self.project = entry.get("project")
        self.projectPath = entry.get("projectPath")
        self.startedAt = entry.get("startedAt")
        self.version = entry.get("version")
        self.dllSha256 = entry.get("dllSha256")
        self.startedByHub = entry.get("startedByHub")
        self.transport = entry.get("transport")
        self.pipe = entry.get("pipe")
        self.nonce = entry.get("nonce")
        self.sessionId = entry.get("sessionId")
        self.instanceId = entry.get("instanceId")
        self.tcpEnabled = entry.get("tcpEnabled")
        # R-1021-03 A4: out-of-band control lane identity (absent on older add-ins).
        self.ctlPipe = entry.get("ctlPipe")
        self.ctlNonce = entry.get("ctlNonce")
        self.ctlProto = entry.get("ctlProto")
        self.ctlAuth = entry.get("ctlAuth")
        self.ctlAvailable = entry.get("ctlAvailable")
        self.ctlTools = entry.get("ctlTools")
        # fields the pass-through surfaces that the OLD whitelist dropped.  The
        # self-registration record (hwnd/tid/epoch) and the wire-protocol stamp (pipeProto)
        # are published by PipeTransport.PublishInstance and were invisible to this reader.
        self.pipeProto = entry.get("pipeProto")
        self.hwnd = _as_int(entry.get("hwnd"))
        self.tid = _as_int(entry.get("tid"))
        self.epoch = _as_int(entry.get("epoch"))
        #: the raw registry object, so a caller can read a field that is not modelled here
        #: without another parse - and so an unregistered field is never unreachable.
        self.raw = entry
#: step 3 - the OPTIONAL textbox doorbell identity. Absent on every
        #: add-in that predates it and on every "pipe" deployment, so a missing field
        #: means "not available", never an error.
        self.ingressMode = entry.get("ingressMode")
        self.ingressHwnd = entry.get("ingressHwnd")
        self.ingressBoxHwnd = entry.get("ingressBoxHwnd")
        self.ingressCaption = entry.get("ingressCaption")
        #: diagnostics: which encoding the registry file needed and which root it came from
        self.jsonEncoding = entry.get("jsonEncoding")
        self.registryRoot = entry.get("registryRoot")
        self.parseError = entry.get("parseError", False)
        self.pipeState, self.pipeStateReason = self._pipe_identity()
        self.pidAlive = pid_alive(self.pid) if self.pid else False

    # -- identity (sec 3.3) ------------------------------------------------------
    def _pipe_identity(self):
        if self.parseError:
            return "none", "unreadable"
        if not self.pipe:
            return "none", "no_pipe_field"
        if self.transport and str(self.transport).lower() != "pipe":
            return "none", "transport_not_pipe"
        if not self.nonce:
            return "none", "no_nonce"
        name = str(self.pipe)
        tail = name.rsplit("-", 1)[-1].strip()
        if tail.lower() != str(self.nonce).strip().lower():
            return "mismatch", "pipe_identity_mismatch"
        expect = "%s-%s" % (self.pid, self.nonce)
        if self.instanceId and str(self.instanceId) != expect:
            return "mismatch", "pipe_identity_mismatch"
        return "candidate", "ok"

    @property
    def pipeAvailable(self):
        """Registry-side verdict only; the live WaitNamedPipe check happens on forward."""
        return self.pipeState == "candidate"

    @property
    def has_route(self):
        return bool(self.pipeAvailable or self.port)

    @property
    def key(self):
        return "%s-%s" % (self.pid, self.nonce or "")

    def brief(self):
        return {
            "pid": self.pid,
            "port": self.port,
            "project": self.project,
            "projectPath": self.projectPath,
            "transport": "pipe" if self.pipeAvailable else "tcp",
            "pipeAvailable": self.pipeAvailable,
            "pipeState": self.pipeState,
            "pidAlive": self.pidAlive,
            "registryFile": self.file,
            "registryRoot": self.registryRoot,
            "pipeProto": self.pipeProto,
            "hwnd": self.hwnd,
            "tid": self.tid,
            "epoch": self.epoch,
            "jsonEncoding": self.jsonEncoding,
        }


# ---------------------------------------------------------------- registry fields ----
# (second half): THE FIELD CONTRACT OF %TEMP%\vb6mcp\instances\*.json.
#
# Why this block exists: read_registry() used to build the entry as a HAND-WRITTEN
# whitelist of 21 keys, so every field the WRITER published was silently swallowed until
# somebody remembered to add it here. That is not hypothetical - the textbox
# identity (ingressMode / ingressHwnd / ingressBoxHwnd / ingressCaption) was lost that
# way, and the same dict still dropped pipeProto / hwnd / tid / epoch, which
# PipeTransport.PublishInstance writes as the self-registration record "the bridge needs
# to bind to the INGRESS window without guessing".  A reader that cannot see a field the
# writer publishes is a disagreement nobody is told about.
#
# The rule is now: PASS EVERYTHING THROUGH, and name only what is deliberately dropped.
#   * our own synthetic keys (file / jsonEncoding / registryRoot) are added on top;
#   * REGISTRY_IGNORED_FIELDS is the EXPLICIT, reviewed drop list (empty by design);
#   * REGISTRY_ACKNOWLEDGED_FIELDS is the set this middle service actually consumes;
#   * anything else is returned by registry_unknown_fields() and LOGGED by discover(),
#     so a new field is visible on the first run instead of in a later investigation.
REGISTRY_SYNTHETIC_FIELDS = frozenset(("file", "jsonEncoding", "registryRoot",
                                       "projectName", "path", "sha256"))
REGISTRY_IGNORED_FIELDS = frozenset(())
REGISTRY_ACKNOWLEDGED_FIELDS = frozenset((
    # identity / liveness
    "pid", "port", "transport", "pipe", "tcpEnabled", "sessionId", "instanceId",
    "nonce", "startedAt", "startedByHub", "dllSha256", "version",
    # project identity (the normalised names; the legacy aliases are SYNTHETIC above)
    "project", "projectPath",
    # pipe wire protocol + the self-registration record
    "pipeProto", "hwnd", "tid", "epoch",
# textbox doorbell identity
    "ingressMode", "ingressHwnd", "ingressBoxHwnd", "ingressCaption",
    # R-1021-03: the ctl/1 control lane (additive; absent => lane unavailable)
    "ctlPipe", "ctlNonce", "ctlProto", "ctlAuth", "ctlAvailable", "ctlTools",
))


def registry_unknown_fields(entry):
    """Registry fields neither ours, nor deliberately ignored, nor ACKNOWLEDGED.

    Empty is the healthy state.  Non-empty means a writer publishes a field this reader was
    never told about: it still PASSES THROUGH (nothing is lost), but it is NAMED here and
    logged by discover(), so "a new field must be explicitly registered" is enforced by a
    report instead of by memory.
    """
    return sorted(k for k in entry
                  if k not in REGISTRY_ACKNOWLEDGED_FIELDS
                  and k not in REGISTRY_IGNORED_FIELDS
                  and k not in REGISTRY_SYNTHETIC_FIELDS)


def read_registry():
    """Tolerant read of %TEMP%\\vb6mcp\\instances\\*.json (sec 4.2.2: never drop an entry)."""
    d = CFG.instances_dir
    entries, stale, roots_seen, candidates = [], [], [], []
    unknown_fields = []
    for root in registry_roots():
        try:
            names = sorted(os.listdir(root))
        except OSError:
            roots_seen.append({"root": root, "exists": False, "entries": 0})
            continue
        jsons = [n for n in names if n.lower().endswith(".json")]
        roots_seen.append({"root": root, "exists": True, "entries": len(jsons)})
        for name in jsons:
            candidates.append((root, name))
    for root, name in candidates:
        full = os.path.join(root, name)
        try:
            with open(full, "rb") as fh:
                j, enc = _loads_registry_bytes(fh.read())
        except Exception as exc:
            stale.append({"file": full, "reason": "parse_error", "detail": str(exc)[:200]})
            continue
        if not isinstance(j, dict):
            stale.append({"file": full, "reason": "not_an_object"})
            continue
        # PASS-THROUGH first - the file's own fields are the entry - and the
        # normalisations are written ON TOP.  Only REGISTRY_IGNORED_FIELDS is dropped, and
        # anything not ACKNOWLEDGED is reported below instead of vanishing.
        entry = dict(j)
        for _drop in REGISTRY_IGNORED_FIELDS:
            entry.pop(_drop, None)
        entry["file"] = full
        entry["pid"] = _as_int(j.get("pid"))
        entry["port"] = _as_int(j.get("port"))
        entry["project"] = j.get("project") or j.get("projectName")
        entry["projectPath"] = j.get("projectPath") or j.get("path")
        entry["startedAt"] = j.get("startedAt")
        entry["version"] = j.get("version")
        entry["dllSha256"] = j.get("dllSha256") or j.get("sha256")
        entry["startedByHub"] = bool(j.get("startedByHub"))
        entry["transport"] = j.get("transport")
        entry["pipe"] = j.get("pipe")
        entry["nonce"] = j.get("nonce")
        entry["sessionId"] = _as_int(j.get("sessionId"))
        entry["instanceId"] = j.get("instanceId")
        entry["tcpEnabled"] = j.get("tcpEnabled")
        entry["ctlPipe"] = j.get("ctlPipe")
        entry["ctlNonce"] = j.get("ctlNonce")
        entry["ctlProto"] = j.get("ctlProto")
        entry["ctlAuth"] = j.get("ctlAuth")
# step 3 - the textbox doorbell identity. Now a pass-through like every
        # other writer field; kept explicit so the normalisation stays readable.
        entry["ingressMode"] = j.get("ingressMode")
        entry["ingressHwnd"] = j.get("ingressHwnd")
        entry["ingressBoxHwnd"] = j.get("ingressBoxHwnd")
        entry["ingressCaption"] = j.get("ingressCaption")
        entry["jsonEncoding"] = enc
        entry["registryRoot"] = root
        _unknown = registry_unknown_fields(entry)
        if _unknown:
            unknown_fields.append({"file": full, "fields": _unknown})
        inst = Instance(entry)
        if any(e.pid == inst.pid and e.key == inst.key for e in entries):
            continue                    # the same instance registered under two roots
        entries.append(inst)
        if not inst.pidAlive:
            stale.append({"file": full, "pid": inst.pid, "reason": "pid_not_alive"})
        elif inst.pipeState == "mismatch":
            stale.append({"file": full, "pid": inst.pid, "reason": "pipe_identity_mismatch"})
    return {"dir": d, "unreadable": False, "instances": entries, "stale": stale,
            "roots": roots_seen, "unknownFields": unknown_fields,
            "ignoredFields": sorted(REGISTRY_IGNORED_FIELDS),
            "acknowledgedFieldCount": len(REGISTRY_ACKNOWLEDGED_FIELDS)}


def _loads_registry_bytes(raw):
    """Decode a registry file the way VB6 actually writes it.  Returns (obj, encoding).

    The add-in writes these files with its OWN ANSI encoding, and the VB6 DEFAULT project name
    is literally the two characters 工程1, so a strict UTF-8 read raises UnicodeDecodeError.
    That used to drop the WHOLE entry and surface as "no_ide_running": the IDE was there and
    healthy, the bridge simply could not read its own registry file.  Measured on t16
    (instances\18464.json): decodes as gbk, fails as utf-8.

    Order: utf-8-sig (BOM tolerant) -> mbcs (the Windows ANSI code page) -> gbk -> latin-1
    (never raises).  A JSON error AFTER a successful decode still propagates, so a genuinely
    broken file is still reported as parse_error - these fallbacks only ever fix DECODING.
    """
    last = None
    for enc in ("utf-8-sig", "mbcs", "gbk", "latin-1"):
        try:
            text = raw.decode(enc)
        except (UnicodeDecodeError, LookupError) as exc:
            last = exc
            continue
        return json.loads(text), enc
    try:
        return json.loads(raw.decode("utf-8", "replace")), "utf-8/replace"
    except Exception:
        raise last if last is not None else ValueError("undecodable registry file")


def registry_roots():
    """Every directory that may hold instances/<pid>.json, most specific first.

    A bridge deployed as a service (SYSTEM) must still find the registrations written by the
    INTERACTIVE user's IDE: the TEMP of a SYSTEM process is C:/Windows/TEMP, which never holds
    them.  That trap cost a whole joint test run on t16 (my own t61 frozen run also had
    TEMP=C:/Windows/TEMP), so it is scanned by default.  Bounded and read-only;
    VB6MCP_INSTANCES_SCAN=0 leaves only the configured directory.
    """
    roots = [CFG.instances_dir]
    if _env_bool("VB6MCP_INSTANCES_SCAN", True):
        roots.append(os.path.join(HERE, "instances"))
        roots.append(os.path.join(os.environ.get("SystemRoot") or "C:/Windows", "TEMP",
                                  "vb6mcp", "instances"))
        try:
            base = "C:/Users"
            for user in sorted(os.listdir(base))[:32]:
                p = os.path.join(base, user, "AppData", "Local", "Temp", "vb6mcp", "instances")
                if os.path.isdir(p):
                    roots.append(p)
        except OSError:
            pass
    out = []
    for d in roots:
        if d and d not in out:
            out.append(d)
    return out


def _as_int(v):
    try:
        return int(v)
    except (TypeError, ValueError):
        return None


# ---------------------------------------------------------------- discovery -------

_DISC_LOCK = threading.Lock()
_DISC_CACHE = {"at": 0.0, "data": None}


def discover(force=False):
    """Registry-only discovery.  NO netstat, NO port scan (ADR sec 1.2 ad 4)."""
    with _DISC_LOCK:
        now = time.monotonic()
        if (not force and _DISC_CACHE["data"] is not None
                and (now - _DISC_CACHE["at"]) * 1000.0 < CFG.discovery_ttl_ms):
            return _DISC_CACHE["data"]
        raw = read_registry()
        data = {
            "dir": raw["dir"],
            "unreadable": raw["unreadable"],
            "instances": raw["instances"],
            "stale": raw["stale"],
            "roots": raw.get("roots") or [],
            "unknownFields": raw.get("unknownFields") or [],
            "live": [i for i in raw["instances"] if i.pidAlive and i.has_route],
            "at": now,
        }
        _DISC_CACHE["at"] = now
        _DISC_CACHE["data"] = data
        dbg("discover: entries=%d live=%d stale=%d" % (len(raw["instances"]), len(data["live"]),
                                                       len(raw["stale"])))
        _unknown = raw.get("unknownFields") or []
        bridge_log("discover", configuredDir=raw.get("dir"), entries=len(raw["instances"]),
                   live=len(data["live"]),
                   roots=";".join("%s:%d" % (r["root"], r["entries"]) for r in (raw.get("roots") or [])),
                   staleReasons=",".join(sorted(set(s.get("reason", "?") for s in raw["stale"]))) or "-",
                   encodings=",".join(sorted(set(str(i.jsonEncoding) for i in raw["instances"]))) or "-",
                   # an unregistered registry field is NAMED, never swallowed.
                   unknownRegistryFields=";".join(
                       "%s:%s" % (os.path.basename(u["file"]), ",".join(u["fields"]))
                       for u in _unknown) or "-")
        return data


# ---------------------------------------------------------------- target pick -----

def _match_selector(inst, sel):
    s = str(sel).strip()
    if not s:
        return False
    if s.isdigit():
        n = int(s)
        if inst.pid == n or inst.port == n:
            return True
    if inst.project and inst.project == s:
        return True
    if inst.projectPath and normalize_path(inst.projectPath).lower() == normalize_path(s).lower():
        return True
    return False


def select_target(selector):
    """(instance|None, refusal|None).  Rule (ADR sec 6, unchanged): explicit target >
    live default > the ONLY live instance; ambiguity is REFUSED, never guessed."""
    d = discover()
    live = d["live"]
    sel = (selector or "").strip()
    if sel and sel.lower() != "auto":
        hits = [i for i in live if _match_selector(i, sel)]
        if len(hits) == 1:
            bridge_log("select", via="explicit", selector=sel, pid=hits[0].pid,
                       port=hits[0].port, pipe=_pipe_leaf(hits[0].pipe))
            return hits[0], None
        if len(hits) == 0:
            return None, _refusal("target_mismatch", "no live instance matches _target=%r" % sel,
                                  d)
        return None, _refusal("target_conflict",
                              "_target=%r matches %d live instances" % (sel, len(hits)), d)
    if len(live) == 1:
        bridge_log("select", via="unique_live", pid=live[0].pid, port=live[0].port,
                   pipe=_pipe_leaf(live[0].pipe), pipeState=live[0].pipeState)
        return live[0], None
    if not live:
        return None, _refusal("no_ide_running", "no live VB6MCP instance", d)
    return None, _refusal("ambiguous_target",
                          "%d live instances; pass _target=<pid|port|project|projectPath>" % len(live), d)


def _pipe_leaf(name):
    """The pipe name without the \\.\\pipe\\ prefix, for compact diagnostic lines."""
    s = "" if name is None else str(name)
    idx = s.rfind("\\")
    return s[idx + 1:] if idx >= 0 else s


def _refusal(reason, message, d):
    bridge_log("no_target", reason=reason, message=message,
               configDir=d.get("dir"), entries=len(d.get("instances") or []),
               live=len(d.get("live") or []),
               staleReasons=",".join(sorted(set(s.get("reason", "?") for s in d.get("stale") or []))) or "-")
    # The registry facts ride along so a refusal is NEVER ambiguous: "no_ide_running" next to
    # registryDir + entries + staleRegistry says immediately WHICH directory was read and what
    # was in it (the t16 defect was a registry file the bridge could not DECODE - the IDE was
    # perfectly healthy, and the old receipt could not say so).
    return {
        "ok": False,
        "reason": reason,
        "message": message,
        "registryDir": d.get("dir"),
        "registryRoots": d.get("roots"),
        "registryEntries": len(d.get("instances") or []),
        "liveInstances": len(d.get("live") or []),
        "instances": [{"pid": i.pid, "port": i.port, "project": i.project} for i in d["live"]],
        "candidates": [i.brief() for i in d["instances"]],
        "staleRegistry": d["stale"],
        "selectionRule": ("explicit _target > live default > the ONLY live instance; "
                          "ambiguity is refused, never guessed"),
        "failureKind": "noIde",
        "failureKindRule": FAILURE_KIND_RULE,
    }


# ---------------------------------------------------------------- transport -------

FAILURE_KIND_RULE = (
    "failureKind = noIde (there is nothing to talk to) | wedged (a peer WAS found and did "
    "not answer: probe failed / forward timed out / the connection died mid-call) | toolError "
    "(the add-in ANSWERED with a structured error).  Additive; legacy reason values are unchanged.  "
    "The R-1021-03 failure envelope adds peerState (noIde | peer_busy | peer_unresponsive | peer_dead), "
    "suggestedAction (retry_once | release_gate | kill_and_restart | ide_start | fix_params), "
    "maxMs/maxMsSource/heldBy/retryAfterMs, ctlAvailable, and - on every control-lane attempt - "
    "oob{attempted,acked,why,ackOp,helloMs,rttMs,delivered,applied}.  The four delivery fields "
    "(delivered/applied/peerState/suggestedAction) are written by the CLIENT and are NEVER emitted "
    "by an add-in: a server states limits{servedFromSamePump,statementInterrupted,"
    "nonPumpingCallServiceable} instead, because it cannot observe a frame it never received.")


class TransportError(Exception):
    def __init__(self, reason, message, code=0, stage="", wrote=False, flags=None):
        Exception.__init__(self, message)
        self.reason = reason
        self.message = message
        self.code = code
        self.stage = stage
        self.wrote = wrote
        self.flags = flags or {}


# What the live WaitNamedPipe probe last said, per instance key (sec 2.3 note 3).
_PIPE_PROBE = {}
_PIPE_PROBE_LOCK = threading.Lock()
_PIPE_PROBE_OK_TTL_MS = 60000.0
_PIPE_PROBE_BAD_TTL_MS = 2000.0


def pipe_probe(inst, force=False):
    """Live pipe availability for this instance.  Cached: a SUCCESS is remembered (no per-call
    WaitNamedPipe); a failure is re-probed after a short TTL so recovery is possible."""
    if not inst.pipeAvailable:
        return inst.pipeState, inst.pipeStateReason, 0
    if not force and POOL.has_live(inst):
        return "ok", "own_live_session", 0
    key = inst.key
    now = time.monotonic() * 1000.0
    with _PIPE_PROBE_LOCK:
        hit = _PIPE_PROBE.get(key)
        if hit and not force and now - hit[0] < hit[3]:
            return hit[1], hit[2], hit[4]
    state, code = wait_named_pipe_bounded(inst.pipe, CFG.pipe_connect_ms)
    why = "waitnamedpipe_%s_%d" % (state, code)
    ttl = _PIPE_PROBE_OK_TTL_MS if state == "ok" else _PIPE_PROBE_BAD_TTL_MS
    with _PIPE_PROBE_LOCK:
        _PIPE_PROBE[key] = (now, state, why, ttl, code)
    return state, why, code


def _drop_pipe_probe(inst):
    with _PIPE_PROBE_LOCK:
        _PIPE_PROBE.pop(inst.key, None)


def choose_track(inst):
    """(track, reason, flags) - the ONE place a track is decided (ADR sec 2.2)."""
    flags = {}
    pref = CFG.preference()
    if pref == "tcp":
        return "tcp", "forced_tcp", flags
    state, why, _code = pipe_probe(inst)
    if state == "ok":
        return "pipe", "forced_pipe" if CFG.strict() else "auto_pipe_ok", flags
    if state == "busy":
        flags["pipeBusy"] = True
        reason = "auto_pipe_busy"
    elif state == "mismatch":
        flags["pipeIdentityMismatch"] = True
        reason = "auto_pipe_error"
    elif state in ("absent", "none"):
        reason = "auto_pipe_absent"
    else:
        reason = "auto_pipe_error"
    if CFG.strict():
        raise TransportError("pipe_unavailable",
                             "VB6MCP_TRANSPORT=pipe but the pipe is %s (%s)" % (state, why),
                             flags=flags)
    if not CFG.pipe_fallback or not inst.port:
        raise TransportError("pipe_unavailable",
                             "the pipe is %s (%s) and no TCP fallback is possible "
                             "(port=%r)" % (state, why, inst.port), flags=flags)
    flags["transportFallback"] = True
    return "tcp", reason, flags


# ---------------------------------------------------------------- pipe sessions ---

class PipeSession(object):
    def __init__(self, inst):
        self.inst = inst
        self.client = None
        self.lock = threading.Lock()
        self.ready = False
        self.last_used = time.monotonic()
        self.welcome = None

    def handshake(self, budget_ms):
        """hello/welcome inside the CONNECT budget only (sec 7.4 first class)."""
        deadline = time.monotonic() + budget_ms / 1000.0
        client = PipeClient(self.inst.pipe)
        self.client = client
        client.connect(budget_ms)
        hello = {
            "vb6mcp": "hello",
            "proto": NP_PROTO,
            "nonce": self.inst.nonce,
            "client": {"pid": os.getpid(), "session": _session_id(), "user": _user_name()},
        }
        if CFG.auth_enabled():
            # M-6: bridge-side only; server-side validation is a separate card in the VB write
            # scope.  Purely additive - a server that ignores `auth` behaves exactly as before,
            # and this bridge never REQUIRES the server to check it (a `welcome` is accepted
            # either way).  tokenFp is for the server's own compare/audit; the server must NOT
            # log it (a published fingerprint is a guess-confirmation oracle).
            hello["auth"] = {
                "scheme": "bearer",
                "token": CFG.token,
                "tokenFp": hashlib.sha256(CFG.token.encode("utf-8")).hexdigest()[:8],
            }
        diag = nonce_diag(self.inst.nonce, hello["nonce"])
        bridge_log("pipe_hello", pipe=self.inst.pipe, pid=self.inst.pid,
                   jsonEncoding=getattr(self.inst, "jsonEncoding", None),
                   payloadBytes=len(encode_frame(hello)), **diag)
        if not diag["equal"]:
            # never send a hello whose nonce is not byte-identical to the registry value: the
            # add-in answers the SAME opaque deny for "missing" / "wrong case" / "whitespace",
            # so failing here keeps the diagnosis on our side.
            raise PipeError(ERROR_ACCESS_DENIED,
                            "local nonce self-check FAILED (expected len=%d hex=%s, sending "
                            "len=%d hex=%s) - refusing to send the hello"
                            % (diag["expectedLen"], diag["expectedHex"], diag["sendingLen"],
                               diag["sendingHex"]), "identity")
        ok, _sent = client.send_frame(hello, deadline)
        if not ok:
            raise PipeError(ERROR_NO_DATA, "hello not written", "handshake")
        while True:
            frame, err = client.recv_frame(deadline)
            if err == "timeout":
                raise PipeError(ERROR_SEM_TIMEOUT, "no welcome within %dms" % budget_ms, "handshake")
            if err is not None:
                raise err
            kind = frame.get("vb6mcp")
            if kind == "deny":
                # the RAW reason and detail are logged here and surfaced on the receipt; a deny
                # must never collapse into a generic "no ide running".
                bridge_log("pipe_deny", pipe=self.inst.pipe, stage="handshake",
                           reason=frame.get("reason"), detail=frame.get("detail"), **diag)
                raise PipeError(ERROR_ACCESS_DENIED,
                                "server denied the handshake: reason=%s detail=%r"
                                % (frame.get("reason"), frame.get("detail")),
                                "handshake",
                                partial=json.dumps(frame, ensure_ascii=False).encode("utf-8"))
            if kind == "welcome":
                self._verify_welcome(frame)
                self.welcome = frame
                bridge_log("pipe_welcome", pipe=self.inst.pipe, pid=frame.get("pid"),
                           session=frame.get("session"), instanceId=frame.get("instanceId"),
                           project=frame.get("project"))
                self.ready = True
                self.last_used = time.monotonic()
                return
            # anything else before the welcome is a protocol violation
            raise PipeError(ERROR_ACCESS_DENIED,
                            "expected welcome, got %r" % (kind or frame.get("method")), "handshake")

    def _verify_welcome(self, frame):
        """Sec 3.4: every one of these must hold, otherwise the pipe is NOT the registered IDE."""
        inst = self.inst
        if frame.get("proto") != NP_PROTO:
            raise PipeError(ERROR_ACCESS_DENIED, "welcome.proto=%r" % frame.get("proto"), "identity")
        expect_iid = "%s-%s" % (inst.pid, inst.nonce)
        if str(frame.get("instanceId") or "") != expect_iid:
            raise PipeError(ERROR_ACCESS_DENIED,
                            "welcome.instanceId=%r expected %r" % (frame.get("instanceId"), expect_iid),
                            "identity")
        if _as_int(frame.get("pid")) != inst.pid:
            raise PipeError(ERROR_ACCESS_DENIED, "welcome.pid=%r" % frame.get("pid"), "identity")
        if inst.sessionId is not None and _as_int(frame.get("session")) != inst.sessionId:
            raise PipeError(ERROR_ACCESS_DENIED,
                            "welcome.session=%r expected %r" % (frame.get("session"),
                                                                inst.sessionId), "identity")
        if inst.dllSha256 and frame.get("dllSha256") and \
                str(frame.get("dllSha256")).upper() != str(inst.dllSha256).upper():
            raise PipeError(ERROR_ACCESS_DENIED, "welcome.dllSha256 mismatch", "identity")

    def close(self):
        try:
            if self.client:
                self.client.close()
        finally:
            self.client = None
            self.ready = False


class PipePool(object):
    """One persistent, serialized connection PER instance (sec 7.5 model A: nMaxInstances=1)."""

    def __init__(self):
        self.lock = threading.Lock()
        self.sessions = {}

    def get(self, inst):
        with self.lock:
            sess = self.sessions.get(inst.key)
            if sess is None:
                sess = PipeSession(inst)
                self.sessions[inst.key] = sess
            return sess

    def has_live(self, inst):
        """True when this process already holds a READY session for that instance.

        WaitNamedPipe answers "busy" when every instance is CONNECTED - and our own persistent
        connection is exactly that.  Without this check a probe-cache expiry would make the
        bridge mistake its own healthy pipe for someone else's and fall back (measured: two
        spurious 'pipe_unavailable' failures in one battery run).
        """
        with self.lock:
            sess = self.sessions.get(inst.key)
        if sess is None or not sess.ready:
            return False
        client = sess.client
        return client is not None and not client.closed

    def drop(self, inst):
        with self.lock:
            sess = self.sessions.pop(inst.key, None)
        if sess is not None:
            sess.close()
        _drop_pipe_probe(inst)

    def close_all(self):
        with self.lock:
            sessions = list(self.sessions.values())
            self.sessions = {}
        for sess in sessions:
            sess.close()


POOL = PipePool()


def _user_name():
    for key in ("USERNAME", "USER"):
        v = os.environ.get(key)
        if v:
            return v
    return "?"


_SESSION_ID = None


def _session_id():
    """Windows session id.  GetCurrentProcessId's session via kernel32; 0 when unavailable."""
    global _SESSION_ID
    if _SESSION_ID is not None:
        return _SESSION_ID
    _SESSION_ID = 0
    try:
        kernel32.ProcessIdToSessionId.restype = wintypes.BOOL
        kernel32.ProcessIdToSessionId.argtypes = [wintypes.DWORD, ctypes.POINTER(wintypes.DWORD)]
        sid = wintypes.DWORD()
        if kernel32.ProcessIdToSessionId(kernel32.GetCurrentProcessId(), ctypes.byref(sid)):
            _SESSION_ID = sid.value
    except Exception:
        pass
    return _SESSION_ID


kernel32.GetCurrentProcessId.restype = wintypes.DWORD


# ---------------------------------------------------------------- forwarding ------

def frame_wrap(payload_bytes):
    """4-byte little-endian length prefix + payload, refusing an over-long frame.

    The ceiling is enforced HERE as well as in FrameBuffer: a request that exceeds it must
    fail STRUCTURALLY instead of putting an over-long length prefix on the wire.
    """
    if len(payload_bytes) > FRAME_MAX:
        raise FrameError("frame_too_large", "%d > %d" % (len(payload_bytes), FRAME_MAX))
    return len(payload_bytes).to_bytes(4, "little") + payload_bytes


# ------------------------------------------------------------- textbox doorbell ---
# (step 3): the INBOUND doorbell.
#
# WHAT PROBLEM THIS SOLVES: the pipe round trip writes the request and then WAITS for the
# answer, so a 30 s tool holds the bridge for 30 s.  With the doorbell the bridge pushes
# the request into the add-in's hidden textbox with SendMessageTimeout - measured to
# return in microseconds - and the add-in executes it afterwards on its own watchdog
# tick.  The reply needs no new channel: it still arrives on the pipe.
#
# MEASURED TRAPS THIS CODE ENCODES (step 1) - do not "simplify" them away:
#   * SetWindowText on another process's child control fails SILENTLY;
#   * PostMessage(WM_SETTEXT) never arrives (the text is not marshalled);
#   * SendMessageTimeout returns NULL on timeout and ctypes with restype=c_void_p turns
#     that into None, while `None != 0` is True - so a timeout once read as a successful
#     delivery.  Every check below goes through _smto_ok();
#   * a write TIMEOUT does not mean "not delivered", so the same payload is never
#     written twice and never replayed on the pipe.

WM_GETTEXTLENGTH = 0x000E
WM_SETTEXT = 0x000C
SMTO_ABORTIFHUNG = 0x0002
INGRESS_MAX_B64 = 12288
INGRESS_WRITE_TIMEOUT_MS = 2000
INGRESS_PROBE_TIMEOUT_MS = 2000

_user32 = None
_ingress_seq_lock = threading.Lock()
_ingress_seq = [0]
_ENUMPROC = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)


def _next_ingress_seq():
    with _ingress_seq_lock:
        _ingress_seq[0] += 1
        return _ingress_seq[0]


def _u32():
    global _user32
    if _user32 is None:
        u = ctypes.WinDLL("user32", use_last_error=True)
        u.SendMessageTimeoutW.restype = ctypes.c_void_p
        u.SendMessageTimeoutW.argtypes = [ctypes.c_void_p, ctypes.c_uint, ctypes.c_size_t,
                                          ctypes.c_void_p, ctypes.c_uint, ctypes.c_uint,
                                          ctypes.POINTER(ctypes.c_size_t)]
        u.IsWindow.restype = wintypes.BOOL
        u.IsWindow.argtypes = [wintypes.HWND]
        u.FindWindowW.restype = wintypes.HWND
        u.FindWindowW.argtypes = [wintypes.LPCWSTR, wintypes.LPCWSTR]
        u.GetClassNameW.restype = ctypes.c_int
        u.GetClassNameW.argtypes = [wintypes.HWND, wintypes.LPWSTR, ctypes.c_int]
        u.EnumChildWindows.restype = wintypes.BOOL
        u.EnumChildWindows.argtypes = [wintypes.HWND, _ENUMPROC, wintypes.LPARAM]
        _user32 = u
    return _user32


def _smto(hwnd, msg, wparam, lparam, timeout_ms):
    """(return value, lpdwResult).  The return value is None on TIMEOUT."""
    u = _u32()
    res = ctypes.c_size_t(0)
    r = u.SendMessageTimeoutW(wintypes.HWND(hwnd), msg, wparam, lparam,
                              SMTO_ABORTIFHUNG, timeout_ms, ctypes.byref(res))
    return r, res.value


def _smto_ok(r):
    """THE trap: NULL (None through c_void_p) is a FAILURE, not a success."""
    return not (r is None or r == 0)


def _is_textbox(hwnd):
    u = _u32()
    buf = ctypes.create_unicode_buffer(256)
    if u.GetClassNameW(wintypes.HWND(hwnd), buf, 255) <= 0:
        return False
    return "textbox" in buf.value.lower()


def _find_by_caption(caption):
    if not caption:
        return 0
    u = _u32()
    fw = u.FindWindowW(None, caption)
    if not fw:
        return 0
    found = []

    def _cb(h, _lp):
        if _is_textbox(h):
            found.append(int(h))
            return False
        return True

    u.EnumChildWindows(wintypes.HWND(fw), _ENUMPROC(_cb), 0)
    return found[0] if found else 0


def _ingress_box(inst):
    """(hwnd, how).  The handle published in the instance record first - it is exact and
    needs no search - with the caption as the fallback for a recycled/stale handle."""
    u = _u32()
    hwnd = _as_int(inst.ingressBoxHwnd) or 0
    if hwnd and u.IsWindow(wintypes.HWND(hwnd)) and _is_textbox(hwnd):
        return hwnd, "recorded"
    bycap = _find_by_caption(inst.ingressCaption)
    if bycap:
        return bycap, "caption"
    return 0, "missing"


def doorbell_write(inst, req_text, budget_ms=None):
    """Push one request into the add-in's hidden textbox.

    Returns (seq, reason, detail); reason == "" means DELIVERED.
    "write_timeout" is the one reason the caller must NOT answer with a pipe fallback:
    the payload may already be in the box.  Every other reason is a clean "not
    available" and the caller may safely write the request on the pipe instead.
    """
    if CFG.ingress != "textbox":
        return None, "switch_pipe", "VB6MCP_INGRESS=%s" % CFG.ingress
    if str(inst.ingressMode or "").lower() != "textbox":
        return None, "instance_not_textbox", "ingressMode=%r" % inst.ingressMode
    mine = _session_id()
    theirs = _as_int(inst.sessionId)
    if theirs and mine and theirs != mine:
        # Cross-session: a session-1 window is not drivable from session 0.
        return None, "cross_session", "instance session=%s bridge session=%s" % (theirs, mine)
    hwnd, how = _ingress_box(inst)
    if not hwnd:
        return None, "hwnd_missing", "box=%r caption=%r (%s)" % (
            inst.ingressBoxHwnd, inst.ingressCaption, how)
    r, boxlen = _smto(hwnd, WM_GETTEXTLENGTH, 0, None, INGRESS_PROBE_TIMEOUT_MS)
    if not _smto_ok(r):
        return None, "probe_timeout", "WM_GETTEXTLENGTH hwnd=%s" % hwnd
    if boxlen != 0:
        # ONE IN FLIGHT, ALWAYS: writing now would overwrite a request whose reply is
        # still on its way, and nothing downstream could report that loss.
        return None, "box_not_empty", "windowLength=%d hwnd=%s" % (boxlen, hwnd)
    seq = _next_ingress_seq()
    raw = "vb6mcp1|%d|%s|%s" % (seq, inst.nonce or "", req_text)
    b64 = base64.b64encode(raw.encode("utf-8")).decode("ascii")
    if len(b64) > INGRESS_MAX_B64:
        # Small payloads ride the window message; a big one goes back to the pipe.
        return None, "payload_too_large", "b64Len=%d limit=%d" % (len(b64), INGRESS_MAX_B64)
    buf = ctypes.create_unicode_buffer(b64)
    r, _res = _smto(hwnd, WM_SETTEXT, 0, ctypes.cast(buf, ctypes.c_void_p),
                    INGRESS_WRITE_TIMEOUT_MS)
    if not _smto_ok(r):
        return seq, "write_timeout", ("WM_SETTEXT hwnd=%s windowLengthAtWrite=%d "
                                      "b64Len=%d seq=%d" % (hwnd, boxlen, len(b64), seq))
    return seq, "", "hwnd=%s via=%s seq=%d b64Len=%d" % (hwnd, how, seq, len(b64))

def pipe_roundtrip(inst, req_text, req_id, budget_ms, ingress_flags=None):
    """Send one JSON-RPC request over the pipe and return (raw_resp_text|None, saw_progress).

    Raises TransportError.  Every wait is deadline-bounded (sec 7.4).
    """
    sess = POOL.get(inst)
    with sess.lock:
        deadline = time.monotonic() + budget_ms / 1000.0
        if not sess.ready:
            try:
                sess.handshake(CFG.pipe_connect_ms)
            except PipeError as exc:
                POOL.drop(inst)
                flags = {"pipeIdentityMismatch": True} if exc.stage == "identity" else {}
                # A server deny carries the ONLY diagnosis the add-in offers, so its raw
# reason/detail are lifted onto the receipt here as well (item 2 of :
                # a deny must never collapse into an anonymous failure).
                deny_reason = ""
                try:
                    deny = json.loads((exc.partial or b"").decode("utf-8"))
                    if isinstance(deny, dict) and deny.get("vb6mcp") == "deny":
                        flags["serverReason"] = deny.get("reason") or "unspecified"
                        deny_reason = str(deny.get("reason") or "")
                        if deny.get("detail"):
                            flags["serverDetail"] = deny.get("detail")
                        flags["serverDenyRaw"] = json.dumps(deny, ensure_ascii=False)
                except Exception:
                    pass
                # auth-deny mapping: a HANDSHAKE AUTH deny must not stay "pipe_connect_failed" with the
                # real cause buried in the message text.  ONLY the two auth reasons are remapped;
                # every other deny (nonce_mismatch / bad_hello / frame_too_large / identity) and
                # every non-deny failure keeps the pre-change reason + flags byte-for-byte.
                if deny_reason in ("token_missing", "token_mismatch"):
                    flags["authReason"] = deny_reason
                    flags["authRequired"] = True
                    flags["authBasis"] = ("np/1 handshake deny frame (server said reason=%s)"
                                          % deny_reason)
                    flags["failureKind"] = "toolError"
                    flags["suggestedAction"] = "fix_params"
                    flags["retrySafe"] = False
                    flags["nothingForwarded"] = True
                    flags["howToProceed"] = [
                        "the add-in ANSWERED the pipe handshake and REFUSED the credential: this is an auth mismatch, not a missing IDE",
                        ("configure VB6MCP_TOKEN / VB6MCP_TOKEN_FILE on the bridge with the token that IDE expects"
                         if deny_reason == "token_missing" else
                         "the supplied token does not match: set VB6MCP_TOKEN / VB6MCP_TOKEN_FILE to the token that IDE expects"),
                        "the token is compared per IDE instance; a bridge serving several IDEs must carry the one they all accept",
                    ]
                    raise TransportError("pipe_auth_failed", str(exc), exc.code,
                                         exc.stage or "handshake", wrote=False, flags=flags)
                raise TransportError("pipe_connect_failed", str(exc), exc.code,
                                     exc.stage or "connect", wrote=False, flags=flags)
        client = sess.client
        _mark("pool_and_handshake")
        payload = req_text.encode("utf-8")
        try:
            framed = frame_wrap(payload)
        except FrameError as exc:
            raise TransportError(exc.reason, str(exc), 0, "frame", wrote=False)
# step 3: the inbound doorbell. Switched OFF by default, so a "pipe"
        # deployment takes exactly the pre-existing path below (byte for byte).
        wrote_on_pipe = True
        if CFG.ingress == "textbox":
            _igseq, why, detail = doorbell_write(inst, req_text, budget_ms)
            if why == "":
                wrote_on_pipe = False
                if ingress_flags is not None:
                    ingress_flags["ingress"] = "textbox"
                    ingress_flags["ingressSeq"] = _igseq
                    ingress_flags["ingressDetail"] = detail
                bridge_log("ingress", pid=inst.pid, result="delivered", seq=_igseq, detail=detail)
            elif why == "write_timeout":
                POOL.drop(inst)
                raise TransportError("ingress_timeout",
                                     "the textbox doorbell write timed out (%s); the payload "
                                     "is NOT rewritten and NOT replayed on the pipe, because "
                                     "a timeout does not mean it was not delivered" % detail,
                                     0, "ingress", wrote=True,
                                     flags={"ingressMode": "textbox", "ingressReason": why,
                                            "ingressDetail": detail, "ingressNoReplay": True})
            else:
                if ingress_flags is not None:
                    ingress_flags["ingressFallback"] = why
                    ingress_flags["ingressDetail"] = detail
                bridge_log("ingress", pid=inst.pid, result="fallback", reason=why, detail=detail)
        if wrote_on_pipe:
            ok, _sent = client.write_all(framed, deadline)
            if not ok:
                POOL.drop(inst)
                raise TransportError("pipe_write_failed", "request not fully written", 0, "write",
                                     wrote=False)
        _mark("transport_write_done")
        if req_id is None:
            return None, False            # a notification gets NO response frame
        seen_progress = False
        last_activity = time.monotonic()
        while True:
            now = time.monotonic()
            remaining = deadline - now
            if remaining <= 0:
                POOL.drop(inst)
                raise TransportError("forward_timeout",
                                     "no answer within FORWARD_TIMEOUT_MS=%d" % budget_ms,
                                     0, "timeout", wrote=True)
            wait = remaining
            if seen_progress:
                idle_left = (last_activity + CFG.pipe_idle_ms / 1000.0) - now
                if idle_left <= 0:
                    POOL.drop(inst)
                    raise TransportError("pipe_stalled",
                                         "no frame for PIPE_IDLE_MS=%d after progress"
                                         % CFG.pipe_idle_ms, 0, "stall", wrote=True)
                wait = min(wait, idle_left)
            frame, err = client.recv_frame(time.monotonic() + wait)
            if err == "timeout":
                if seen_progress and (time.monotonic() - last_activity) * 1000.0 >= CFG.pipe_idle_ms:
                    POOL.drop(inst)
                    raise TransportError("pipe_stalled",
                                         "no frame for PIPE_IDLE_MS=%d after progress"
                                         % CFG.pipe_idle_ms, 0, "stall", wrote=True)
                continue
            if err is not None:
                POOL.drop(inst)
                code = getattr(err, "code", 0)
                reason = "pipe_stalled" if code in (ERROR_BROKEN_PIPE, ERROR_NO_DATA,
                                                    ERROR_PIPE_NOT_CONNECTED) else "pipe_read_failed"
                raise TransportError(reason, str(err), code, "read", wrote=True)
            kind = frame.get("vb6mcp")
            if kind == "progress":
                seen_progress = True
                last_activity = time.monotonic()
                seq = _as_int(frame.get("seq")) or 0
                if seq and seq <= client.progress_seq:
                    dbg("pipe: non-increasing progress seq", seq)
                client.progress_seq = max(client.progress_seq, seq)
                continue
            if kind == "pong":
                last_activity = time.monotonic()
                continue
            if kind == "deny":
                POOL.drop(inst)
                server_reason = frame.get("reason") or "unspecified"
                flags = {"serverReason": server_reason}
                if frame.get("detail"):
                    flags["serverDetail"] = frame.get("detail")
                raise TransportError("pipe_denied", "server denied: %s" % server_reason, 0,
                                     "deny", wrote=True, flags=flags)
            if kind in ("hello", "welcome"):
                continue
            if frame.get("id") == req_id:
                sess.last_used = time.monotonic()
                _mark("response_frame")
                return client.last_payload.decode("utf-8"), seen_progress
            # an unexpected frame: additive frames are allowed, keep waiting inside the budget
            last_activity = time.monotonic()
            seen_progress = seen_progress or False


def tcp_roundtrip(inst, req_text, budget_ms):
    """The spike control / rollback / transition track: the add-in's own HTTP /rpc."""
    url = "http://127.0.0.1:%s/rpc" % inst.port
    req = urllib.request.Request(url, data=req_text.encode("utf-8"),
                                 headers={"Content-Type": "application/json"})
    try:
        with urllib.request.urlopen(req, timeout=max(1.0, budget_ms / 1000.0)) as resp:
            return resp.read().decode("utf-8", "replace")
    except urllib.error.HTTPError as exc:
        try:
            return exc.read().decode("utf-8", "replace")
        except Exception:
            raise TransportError("forward_failed", "HTTP %s" % exc.code, 0, "tcp")
    except urllib.error.URLError as exc:
        raise TransportError("forward_failed", str(exc.reason), 0, "tcp")
    except socket.timeout:
        raise TransportError("forward_timeout", "tcp read timeout", 0, "tcp")


def pipe_roundtrip_busy_retry(inst, req_text, req_id, budget_ms, igf):
    """One immediate retry on ERROR_PIPE_BUSY (A4: transient busy, not no-IDE)."""
    try:
        return pipe_roundtrip(inst, req_text, req_id, budget_ms, igf)
    except TransportError as exc:
        if _peer_state(exc.reason, exc.flags, inst) == "peer_busy" and not exc.wrote:
            time.sleep(max(0, CFG.pipe_busy_retry_ms) / 1000.0)
            return pipe_roundtrip(inst, req_text, req_id, budget_ms, igf)
        raise


def forward_raw(req_text, req_obj, selector, timeout_ms=None):
    """Forward one JSON-RPC message.  Returns a dict:
       {text|None, transport, transportReason, flags{}, elapsedMs, ide{}, failure{}}"""
    # PER-MESSAGE budget: None keeps the configured value, so the default forward path is
    # unchanged.  The only caller that passes it is the bounded ide_status live probe.
    tool_for_budget = _tool_of(req_obj) if isinstance(req_obj, dict) else None
    if timeout_ms:
        budget, budget_src = timeout_ms, "explicit"
    elif tool_for_budget and isinstance(req_obj, dict) and req_obj.get("method") == "tools/call":
        budget, budget_src = tool_budget(tool_for_budget, _args_of(req_obj))
    else:
        budget, budget_src = CFG.forward_timeout_ms, "env_forward"
    inst, refusal = select_target(selector)
    _mark("select_target")
    if inst is None:
        return {"text": None, "failure": refusal, "elapsedMs": 0}
    ide = {"pid": inst.pid, "project": inst.project, "port": inst.port}
    t0 = time.monotonic()
    try:
        track, reason, flags = choose_track(inst)
    except TransportError as exc:
        body = _transport_failure_body(exc.reason, exc.message, t0, budget,
                                       tool_for_budget, budget_src, inst, exc.flags)
        body.update(exc.flags)
        return {"text": None, "failure": body, "elapsedMs": _ms(t0), "ide": ide}
    _mark("track_chosen")
    req_id = req_obj.get("id") if isinstance(req_obj, dict) else None
    bridge_log("track", tool=_tool_of(req_obj), pid=inst.pid, port=inst.port,
               pipe=_pipe_leaf(inst.pipe), track=track, reason=reason,
               flags=",".join(sorted(flags)) or "-")
    if track == "pipe":
        try:
            igf = {}
            text, saw_progress = pipe_roundtrip(inst, req_text, req_id,
                                               budget, igf)
            _mark("roundtrip_done")
            flags = dict(flags)
            flags.update(igf)
            return {"text": text, "transport": "pipe", "transportReason": reason,
                    "flags": flags, "elapsedMs": _ms(t0), "ide": ide,
                    "sawProgress": saw_progress}
        except TransportError as exc:
            flags = dict(flags)
            flags.update(exc.flags)
            if exc.reason == "pipe_stalled":
                flags["pipeStalled"] = True
            retryable = (not exc.wrote) or CFG.pipe_replay
            if (not CFG.strict()) and inst.port and retryable:
                flags["transportFallback"] = True
                dbg("pipe failed (%s) -> tcp fallback (wrote=%s)" % (exc.reason, exc.wrote))
                try:
                    text = tcp_roundtrip(inst, req_text, budget)
                    return {"text": text, "transport": "tcp", "transportReason": reason,
                            "flags": flags, "elapsedMs": _ms(t0), "ide": ide,
                            "fellBackFrom": exc.reason}
                except TransportError as exc2:
                    body = _transport_failure_body(exc2.reason, exc2.message, t0, budget,
                                                   tool_for_budget, budget_src, inst, flags)
                    body.update(flags)
                    body["transport"] = "tcp"
                    body["transportReason"] = reason
                    return {"text": None, "failure": body, "elapsedMs": _ms(t0), "ide": ide}
            reason2 = "auto_pipe_busy" if flags.get("pipeBusy") else (
                "auto_pipe_error" if flags.get("pipeIdentityMismatch") else "auto_pipe_error")
            body = _transport_failure_body(exc.reason, exc.message, t0, budget,
                                           tool_for_budget, budget_src, inst, flags)
            body.update(flags)
            body["transport"] = "pipe"
            body["transportReason"] = reason2
            if exc.wrote and not CFG.pipe_replay:
                body["noReplay"] = True
                body["hint"] = ("the request WAS fully written to the pipe, so it is never "
                                "silently replayed (that could double-apply a write tool); "
                                "set VB6MCP_PIPE_REPLAY=1 to opt in")
            return {"text": None, "failure": body, "elapsedMs": _ms(t0), "ide": ide}
    try:
        text = tcp_roundtrip(inst, req_text, budget)
        return {"text": text, "transport": "tcp", "transportReason": reason,
                "flags": dict(flags), "elapsedMs": _ms(t0), "ide": ide}
    except TransportError as exc:
        body = _transport_failure_body(exc.reason, exc.message, t0, budget,
                                       tool_for_budget, budget_src, inst, flags)
        body["transport"] = "tcp"
        body["transportReason"] = reason
        body.update(flags)
        return {"text": None, "failure": body, "elapsedMs": _ms(t0), "ide": ide}


def _tool_of(req_obj):
    try:
        return (req_obj.get("params") or {}).get("name") or req_obj.get("method")
    except Exception:
        return None


def _ms(t0):
    return int((time.monotonic() - t0) * 1000)


def _args_of(req_obj):
    try:
        return (req_obj.get("params") or {}).get("arguments") or {}
    except Exception:
        return {}


def tool_budget(tool, params=None):
    """(maxMs, source) - the bridge-side mirror of CodeTools.ToolBudgetMs."""
    p = params or {}
    g = CFG.max_call_ms
    hard = CFG.max_call_hard_ms
    t = (tool or "").lower()
    def _i(k, d):
        try:
            return int(p.get(k, d))
        except Exception:
            return d
    if t in ("immediate_eval", "immediate_run"):
        v = _i("timeout_ms", 5000)
        src = "param" if "timeout_ms" in p else "tool_default"
    elif t == "make_project":
        v, src = 140000, "tool_default"
    elif t == "build":
        v, src = _i("timeout_ms", 180000) + 15000, "param_slack"
    elif t in ("debug_session", "run_until_break"):
        v, src = _i("timeout_ms", 30000) + 8000, "param_slack"
    elif t == "dependency_graph":
        b = max(_i("budget_ms", 0), 25000)
        v, src = max(b + 8000, 33000), "param_slack"
    elif t == "eval_at_frame":
        v, src = _i("timeoutMs", 5000) + 2000, "param_slack"
    elif t == "get_locals":
        v, src = 8000, "tool_default"
    else:
        v, src = g, "env_ceiling"
    v = max(v, 1000)
    if v > hard:
        v, src = hard, src + "+hard_cap"
    return v, src


def _peer_state(reason, flags=None, inst=None):
    """peer_busy (wait) | peer_unresponsive (do not wait) | noIde (nothing there)."""
    flags = flags or {}
    r = (reason or "").lower()
    if flags.get("pipeBusy") or "busy" in r:
        return "peer_busy"
    if r in ("pipe_unavailable", "absent", "not_found", "denied", "no_ide"):
        return "noIde"
    return "peer_unresponsive"


def _pid_alive(pid):
    """True / False when the liveness of `pid` can be PROVED, None when it cannot.

    A DEAD peer is the one fact that changes the remedy: "peer_unresponsive" says wait or
    release the gate, but a process that is GONE can never answer and its debug context is
    lost.  Windows: OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION) + GetExitCodeProcess ==
    STILL_ACTIVE (no signal is ever sent).  POSIX: kill(pid, 0).  ERROR_ACCESS_DENIED and
    EPERM mean UNKNOWN, never "dead" - a live process owned by someone else must not be
    reported as gone.
    """
    try:
        pid = int(pid)
    except (TypeError, ValueError):
        return None
    if pid <= 0:
        return None
    if os.name == "nt":
        try:
            import ctypes
            from ctypes import wintypes
            k32 = ctypes.WinDLL("kernel32", use_last_error=True)
            PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
            STILL_ACTIVE = 259
            h = k32.OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, False, pid)
            if not h:
                # 5 = ERROR_ACCESS_DENIED: the process may be alive and unqueryable -> UNKNOWN.
                return None if ctypes.get_last_error() == 5 else False
            try:
                code = wintypes.DWORD()
                if not k32.GetExitCodeProcess(h, ctypes.byref(code)):
                    return None
                return code.value == STILL_ACTIVE
            finally:
                k32.CloseHandle(h)
        except Exception:
            return None
    try:
        os.kill(pid, 0)
        return True
    except ProcessLookupError:
        return False
    except PermissionError:
        return None
    except OSError:
        return None


def _inst_pid(inst):
    """The pid of a registry record whether it is an object or a dict (None when absent)."""
    if inst is None:
        return None
    if isinstance(inst, dict):
        return inst.get("pid")
    return getattr(inst, "pid", None)


def _has_ctl(inst):
    """Feature-detect the ctl/1 lane from the instance record (sec 6): no fields => unavailable.

    This is the SAME predicate the client uses, so a receipt that says ctlAvailable:true and a
    client that refuses to use the lane can never disagree.
    """
    return ctl_available(inst)


def _transport_failure_body(reason, message, t0, budget_ms, tool=None, budget_src=None,
                            inst=None, flags=None):
    """R-1021-03 A4: one structured envelope for every transport failure.
    reason is kept for compatibility; reasonClass/timeout say what a caller must branch on."""
    m = message or ""
    elapsed = _ms(t0)
    peer = _peer_state(reason, flags, inst)
    # MCP-1021-43 (WP-22-02): a WEDGE-shaped failure is only "unresponsive" while the process
    # still EXISTS.  Re-check liveness ONCE here, so a pipe_stalled/timeout whose IDE has
    # actually exited answers peer_dead (+ the lost debug context + a recovery path) instead of
    # sending the caller to release_flight on a process that is gone.
    peer_dead = False
    dead_pid = None
    if peer == "peer_unresponsive":
        dead_pid = _inst_pid(inst)
        if _pid_alive(dead_pid) is False:
            peer = "peer_dead"
            peer_dead = True
    ctl = _has_ctl(inst)
    body = {
        "ok": False,
        "reason": "timeout" if reason == "forward_timeout" else reason,
        "reasonLegacy": reason,
        "message": m,
        "elapsedMs": elapsed,
        "maxMs": budget_ms,
        "maxMsSource": budget_src or "env_forward",
        "budgetMs": budget_ms,
        "failureKind": peer,
        "failureKindRule": FAILURE_KIND_RULE,
        "peerState": peer,
    }
    if peer_dead:
        # A dead peer is a DIFFERENT fact from a wedged one: there is nothing left to release.
        debug_tools = ("run_project", "run_until_break", "debug_session", "get_locals",
                       "eval_at_frame", "last_error", "current_position", "get_callstack")
        body["peerDead"] = True
        body["suggestedAction"] = "kill_and_restart"
        body["lostContext"] = {
            "peerDead": True,
            "pid": dead_pid,
            "debugger": bool(tool in debug_tools),
            "note": ("the IDE process is GONE (proved dead): its add-in pipe, the running program, "
                     "the break state and the locals died with it; nothing can be read from it any more"),
        }
        body["howToRecover"] = [
            "the peer is provably DEAD, so release_inflight cannot help - there is nothing left to release",
            "recover out-of-band with the hub tool ide_kill_and_relaunch{pid, project?} (or ide_start on the same .vbp); taskkill from the caller is NOT the standard path",
            "the debug context (break position, locals, callstack) is LOST - re-run to reproduce it",
        ]
    elif peer == "peer_busy":
        body["suggestedAction"] = "retry_once"
        body["retryAfterMs"] = CFG.pipe_busy_retry_ms
    elif peer == "peer_unresponsive":
        body["suggestedAction"] = "release_gate" if ctl else "kill_and_restart"
    # A published lane is a FACT about this instance, not a property of the failure class: report
    # it on every envelope (a caller that hits peer_busy or noIde still needs to know whether a
    # recovery lane exists).  Absent means "no lane"; false is only stated for the unresponsive
    # case, which is the one where the absence changes the remedy.
    if ctl:
        body["ctlAvailable"] = True
    elif peer == "peer_unresponsive":
        body["ctlAvailable"] = False
    if tool:
        body["heldBy"] = {"tool": tool, "pid": getattr(inst, "pid", None),
                          "sinceMs": elapsed}
    return body


# ---------------------------------------------------------------- receipts --------

def contract_version():
    snap = read_json(SNAP_PATH, None)
    if isinstance(snap, dict) and snap.get("contractVersion"):
        return snap["contractVersion"]
    return "unknown"


def stamp_body(result, ide, contract, note):
    """The hub's stamping contract, restated: the JSON-RPC result keys stay EXACTLY
    {content,isError}; mcp/ide/transport markers live in the BODY TEXT only (ADR sec 2.5).

    content[0].text is ONE JSON document (v1.14.0): an ARRAY/scalar body becomes
    {singleDocument:true, data, receipt} (an EMPTY result is data:[] - distinguishable from
    an unparsable body), a non-JSON body becomes a structured content_not_json document and
    an OBJECT body keeps its FLAT shape.  VB6MCP_LEGACY_TEXTBODY=1 restores the OLD
    two-document form for ONE version (migration switch, byte-for-byte unchanged)."""
    content = []
    for c in (result.get("content") or []):
        content.append(dict(c))
    text_item = None
    for c in content:
        if c.get("type") == "text" and isinstance(c.get("text"), str):
            text_item = c
            break
    if text_item is not None:
        raw = text_item["text"]
        if (os.environ.get("VB6MCP_LEGACY_TEXTBODY") or "") == "1":
            # ONE VERSION ONLY: the pre-v1.14.0 two-document form, kept byte-for-byte unchanged.
            obj = None
            try:
                legacy_parsed = json.loads(raw)
                if isinstance(legacy_parsed, dict):
                    obj = legacy_parsed
            except Exception:
                obj = None
            if obj is not None:
                text_item["text"] = _body_of(obj, ide, contract, note)
            else:
                text_item["text"] = raw + "\n" + _body_of({}, ide, contract, note)
        else:
            parsed = None
            parse_ok = False
            try:
                parsed = json.loads(raw)
                parse_ok = True
            except Exception:
                parse_ok = False
            already_single = (parse_ok and isinstance(parsed, dict)
                              and parsed.get("singleDocument") is True
                              and "data" in parsed and isinstance(parsed.get("receipt"), dict))
            if parse_ok and isinstance(parsed, dict) and not already_single:
                text_item["text"] = _body_of(parsed, ide, contract, note)
            elif already_single:
                text_item["text"] = json.dumps(
                    {"singleDocument": True, "data": parsed["data"],
                     "receipt": json.loads(_body_of(parsed["receipt"], ide, contract, note))},
                    ensure_ascii=False, indent=2)
            elif parse_ok:
                text_item["text"] = json.dumps(
                    {"singleDocument": True, "data": parsed,
                     "receipt": json.loads(_body_of({}, ide, contract, note))},
                    ensure_ascii=False, indent=2)
            else:
                text_item["text"] = json.dumps(
                    {"ok": False, "reason": "content_not_json", "nonJsonText": True,
                     "rawText": raw, "receipt": json.loads(_body_of({}, ide, contract, note))},
                    ensure_ascii=False, indent=2)
    else:
        content.append({"type": "text", "text": _body_of({"ok": True}, ide, contract, note)})
    return {"content": content, "isError": result.get("isError") is True}


def _body_of(obj, ide, contract, note):
    o = dict(obj)
    if not o.get("mcp"):
        o["mcp"] = {"product": PRODUCT, "contract": contract}
    if ide and not o.get("ide"):
        o["ide"] = ide
    if note:
        for k, v in note.items():
            if k not in o:
                o[k] = v
    return json.dumps(o, ensure_ascii=False, indent=2)


def error_result(body, is_error=True):
    return {"content": [{"type": "text", "text": json.dumps(body, ensure_ascii=False, indent=2)}],
            "isError": bool(is_error)}


# ---------------------------------------------------------------- dispatch --------

SUPPORTED_PROTOCOLS = ["2025-11-25", "2025-06-18", "2025-03-26", "2024-11-05", "2024-10-07"]
LATEST_PROTOCOL = SUPPORTED_PROTOCOLS[0]

_SNAP_CACHE = {"mtime": None, "tools": [], "instructions": "", "contract": "unknown"}


def snapshot():
    """The frozen add-in face (bridge/hub_schema.generated.json).  Read-only, cached by mtime.

    This is what makes the thin shell (and this service) able to publish the COMPLETE tool
    list with no IDE and no engine - criterion 4 of the card.
    """
    try:
        mtime = os.path.getmtime(SNAP_PATH)
    except OSError:
        return dict(_SNAP_CACHE)
    if _SNAP_CACHE["mtime"] != mtime:
        j = read_json(SNAP_PATH, {})
        if isinstance(j, dict):
            tools = j.get("tools") if isinstance(j.get("tools"), list) else []
            _SNAP_CACHE.update({"mtime": mtime, "tools": tools,
                                "instructions": j.get("instructions") or "",
                                "contract": j.get("contractVersion") or "unknown"})
    return dict(_SNAP_CACHE)


def negotiate_protocol(client_version):
    if client_version and client_version in SUPPORTED_PROTOCOLS:
        return client_version
    return LATEST_PROTOCOL


def local_initialize(req):
    """A LOCAL initialize for the degraded path (no IDE live).  When an IDE IS live the
    add-in's own answer is forwarded byte-for-byte instead - this is only the fallback."""
    params = req.get("params") if isinstance(req.get("params"), dict) else {}
    snap = snapshot()
    return {
        "protocolVersion": negotiate_protocol(params.get("protocolVersion")),
        "capabilities": {"tools": {}, "resources": {"listChanged": False},
                         "prompts": {"listChanged": False}},
        "serverInfo": {"name": PRODUCT, "version": VERSION},
        "contractVersion": snap["contract"] or "unknown",
        "instructions": snap["instructions"] or "",
    }


def _jsonrpc_result(rid, result):
    return json.dumps({"jsonrpc": "2.0", "id": rid, "result": result}, ensure_ascii=False)


def _jsonrpc_error(rid, code, message, data=None):
    err = {"code": code, "message": message}
    if data is not None:
        err["data"] = data
    return json.dumps({"jsonrpc": "2.0", "id": rid, "error": err}, ensure_ascii=False)


def _transport_note(out):
    note = {}
    if out.get("transport"):
        note["transport"] = out["transport"]
    if out.get("transportReason"):
        note["transportReason"] = out["transportReason"]
    for k, v in (out.get("flags") or {}).items():
        note[k] = v
    if out.get("elapsedMs") is not None:
        note["forwardElapsedMs"] = out["elapsedMs"]
    if out.get("fellBackFrom"):
        note["transportFallbackFrom"] = out["fellBackFrom"]
    return note


def _extract_target(req, default):
    """_target may be supplied per call (hub-compatible) - it is STRIPPED before forwarding."""
    params = req.get("params")
    if not isinstance(params, dict):
        return default, req
    args = params.get("arguments")
    if not isinstance(args, dict) or "_target" not in args:
        return default, req
    raw = args.get("_target")
    sel = default
    if raw is not None and str(raw).strip() and str(raw).strip().lower() != "auto":
        sel = str(raw).strip()
    req2 = dict(req)
    params2 = dict(params)
    args2 = dict(args)
    args2.pop("_target", None)
    params2["arguments"] = args2
    req2["params"] = params2
    return sel, req2


# ------------------------------------------------------------- timing probe (147) ---
# ONE-SHOT per-request stage timing.  OFF unless VB6MCP_TIMING=1, so a normal
# deployment pays nothing, writes nothing and behaves byte-for-byte as before.  Remove
# the whole block (or leave it switched off) once the attribution is recorded - a
# permanently-on log would pollute the CPU measurements this project relies on.
TIMING_ON = (os.environ.get("VB6MCP_TIMING") or "") == "1"
TIMING_FILE = os.environ.get("VB6MCP_TIMING_FILE") or ""
_TL = threading.local()
_T_LOCK = threading.Lock()


def _t0():
    if not TIMING_ON:
        return
    _TL.t = {"t0": time.perf_counter(), "marks": []}


def _mark(name):
    if not TIMING_ON:
        return
    t = getattr(_TL, "t", None)
    if t is None:
        return
    t["marks"].append((name, time.perf_counter()))


def _timing_emit(extra=None):
    if not TIMING_ON:
        return
    t = getattr(_TL, "t", None)
    if not t:
        return
    base = t["t0"]
    row = {"stages": []}
    prev, prevname = base, "start"
    for name, ts in t["marks"]:
        row["stages"].append([prevname + ">" + name, round((ts - prev) * 1000.0, 3)])
        prev, prevname = ts, name
    row["totalMs"] = round((prev - base) * 1000.0, 3)
    if extra:
        row.update(extra)
    if TIMING_FILE:
        with _T_LOCK:
            with open(TIMING_FILE, "a", encoding="utf-8") as fh:
                fh.write(json.dumps(row) + "\n")
    _TL.t = None


# ---------------------------------------------------------------- local facts RPC -----
# Three-layer ownership (ADR sec 0.1): the mjs side keeps the MCP wire protocol; this Python
# middle service is the FACTS OWNER for the IDE set on THIS machine.
#
# Two READ-ONLY methods are served LOCALLY, in the JSON-RPC METHOD namespace:
#
#     POST /rpc {"jsonrpc":"2.0","id":1,"method":"ide_list"}
#     POST /rpc {"jsonrpc":"2.0","id":2,"method":"ide_status","params":{"target":"<pid|project>","probe":true}}
#
# They can never shadow an add-in tool: the add-in's names live in tools/list + tools/call, a
# DIFFERENT namespace entirely.  Every OTHER method is forwarded exactly as before, so an older
# middle service answers -32601 "Method not found" for these two names - which is the capability
# signal the hub uses to fall back to GET /health (one probe, then cached for the process).
#
# Invariants (why this lives here): READ-ONLY.  Registry read only, no spawn, no lock write, no
# registration pruning, and ide_list NEVER touches the add-in (no pipe session, no contention
# with a running tool).  ide_status{probe:true} is the ONE opt-in live check and it is bounded
# by probeTimeoutMs (hard-capped at 3000 ms).
LOCAL_RPC_METHODS = ("ide_list", "ide_status")
LOCAL_METHOD_PRIORITY = (
    "a bridge-local JSON-RPC METHOD wins over forwarding for exactly these names; the add-in "
    "TOOL face (tools/list + tools/call) is a separate namespace and is never shadowed or "
    "rewritten by this interception. Unknown methods keep the pre-existing forward behaviour.")
ADDIN_REACHABLE_RULE = (
    "addInReachable = the registered pid is ALIVE and the instance has a route (a consistent np/1 "
    "pipe registration or a TCP port). It is a ROUTE-level verdict derived from the registry and "
    "the OS - ide_list never touches the add-in (no pipe session, no contention). "
    "ide_status{probe:true} performs ONE bounded live mcp_info over the same route a forwarded "
    "call uses and reports it in liveProbe{attempted,ok,reason,elapsedMs}.")
REDACTION_RULE = (
    "a non-loopback caller gets the SAME redacted view /health gives it: no instance path, pipe "
    "name, pid, registration file or registry directory. Loopback callers (and the stdio owner) "
    "get the full facts. Redaction is decided by the PEER ADDRESS only, never by a parameter.")


def _iso_utc_now():
    return time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime())


def _is_remote_peer(peer):
    """True only for a caller that is NOT loopback; None (= the stdio owner) is local."""
    if peer is None:
        return False
    try:
        return not peer_is_loopback(str(peer))
    except Exception:
        return True                     # fail closed: an address we cannot judge is remote


def _bridge_facts():
    auth = auth_state()
    snap = snapshot() or {}
    return {
        "product": PRODUCT,
        "version": VERSION,
        "pid": os.getpid(),
        "host": CFG.host,
        "port": CFG.port,
        "startedAt": _BOOT_ISO,
        "uptimeMs": int(max(0.0, time.time() - _BOOT_AT) * 1000.0),
        "transport": CFG.transport or "auto(pipe-first)",
        "frozen": bool(getattr(sys, "frozen", False)),
        "auth": {"enabled": bool(auth.get("enabled")),
                 "required": bool(auth.get("required")),
                 "scheme": auth.get("scheme"),
                 "source": auth.get("source"),
                 "tokenLen": auth.get("tokenLen")},
        "contractVersion": contract_version(),
        "snapshotTools": len(snap.get("tools") or []),
        "localMethods": list(LOCAL_RPC_METHODS),
        "readOnly": True,
    }


def _reach_basis(inst):
    if not inst.pidAlive:
        return "pid_not_alive"
    if inst.pipeAvailable:
        return "pid_alive_and_pipe_registration_consistent(route=pipe)"
    if inst.port:
        return "pid_alive_and_tcp_port_present(route=tcp)"
    return "pid_alive_but_no_route_pipe_or_tcp"


def _redact_instance_facts(facts):
    keep = ("project", "port", "transport", "alive", "addInReachable", "addInReachableBasis")
    out = {k: facts.get(k) for k in keep}
    # the ctl pipe name embeds the ctlNonce, so a non-loopback caller gets the VERDICT only
    # (available + why) and never the name or the credential.
    ctl = facts.get("ctl") or {}
    out["ctl"] = {"available": ctl.get("available"), "reason": ctl.get("reason"),
                  "detail": ctl.get("detail"), "redacted": True}
    out["redacted"] = True
    return out


def _instance_facts(inst, redact):
    facts = {
        "pid": inst.pid,
        "project": inst.project,
        "projectPath": inst.projectPath,
        "port": inst.port,
        "pipe": inst.pipe,
        "tcpEnabled": inst.tcpEnabled,
        "epoch": inst.epoch,
        "sessionId": inst.sessionId,
        "startedAt": inst.startedAt,
        "alive": bool(inst.pidAlive),
        "addInReachable": bool(inst.pidAlive and inst.has_route),
        "registrationFile": inst.file,
        # additive transport detail (the SAME vocabulary /health already publishes)
        "transport": "pipe" if inst.pipeAvailable else "tcp",
        "pipeAvailable": bool(inst.pipeAvailable),
        "pipeState": inst.pipeState,
        "instanceId": inst.instanceId,
        "version": inst.version,
        "dllSha256": inst.dllSha256,
        "registryRoot": inst.registryRoot,
        "hwnd": inst.hwnd,
        "tid": inst.tid,
        "addInReachableBasis": _reach_basis(inst),
        "addInReachableRule": ADDIN_REACHABLE_RULE,
    }
    # R-1021-03: the control lane is FEATURE-DETECTED evidence, never a promise: a lane that is
    # unpublished is reported with the reason instead of being silently absent.
    _ctl_why = ctl_problem(inst)
    facts["ctl"] = {
        "available": _ctl_why is None,
        "reason": _ctl_why[0] if _ctl_why else None,
        "detail": _ctl_why[1] if _ctl_why else None,
        "proto": inst.ctlProto or None,
        "authBasis": inst.ctlAuth or None,
        "tools": inst.ctlTools if isinstance(inst.ctlTools, list) else None,
        "pipe": inst.ctlPipe,
        "nonce": inst.ctlNonce,
        "pipeTag": _ctl_tag(inst.ctlNonce),
    }
    return _redact_instance_facts(facts) if redact else facts


def _stale_facts(d, redact):
    if redact:
        return []
    out = []
    for s in (d.get("stale") or []):
        f = s.get("file")
        pid = s.get("pid")
        try:
            age = int((time.time() - os.path.getmtime(f)) * 1000.0) if f else None
        except OSError:
            age = None
        pid_dead = None
        if pid is not None:
            try:
                pid_dead = not pid_alive(pid)
            except Exception:
                pid_dead = None
        row = {"pid": pid, "file": f, "pidDead": pid_dead,
               "reason": s.get("reason"), "ageMs": age}
        if s.get("detail"):
            row["detail"] = s.get("detail")
        out.append(row)
    return out


def _selection_facts(d, target, redact):
    live = d.get("live") or []
    cands = [{"pid": i.pid, "port": i.port, "project": i.project} for i in live]
    if redact:
        cands = [{"project": c.get("project"), "port": c.get("port"), "redacted": True}
                 for c in cands]
    return {
        "requested": target,
        "rule": ("explicit target > live default > the ONLY live instance; ambiguity is "
                 "refused, never guessed"),
        "liveCount": len(live),
        "candidates": cands,
    }


def _redact_refusal(ref):
    out = dict(ref or {})
    for k in ("registryDir", "registryRoots", "candidates", "staleRegistry"):
        out.pop(k, None)
    out["instances"] = [{"project": (c or {}).get("project"),
                         "port": (c or {}).get("port"), "redacted": True}
                        for c in (out.get("instances") or [])]
    out["remoteRedacted"] = True
    return out


def _probe_addin_bounded(inst, selector, timeout_ms):
    """ONE bounded live mcp_info over the SAME route a forwarded call uses.  Never raises."""
    probe = {"attempted": True, "ok": False, "reason": None, "elapsedMs": None,
             "timeoutMs": timeout_ms,
             "route": "pipe" if inst.pipeAvailable else ("tcp" if inst.port else "none")}
    req = {"jsonrpc": "2.0", "id": 991001, "method": "tools/call",
           "params": {"name": "mcp_info", "arguments": {}}}
    try:
        out = forward_raw(json.dumps(req), req, selector, timeout_ms=timeout_ms)
    except Exception as exc:
        probe["reason"] = "probe_error"
        probe["message"] = str(exc)[:300]
        return probe
    probe["elapsedMs"] = out.get("elapsedMs")
    probe["transport"] = out.get("transport")
    probe["transportReason"] = out.get("transportReason")
    text = out.get("text")
    if not text:
        failure = out.get("failure") or {}
        probe["reason"] = failure.get("reason") or "no_answer"
        probe["failureKind"] = failure.get("failureKind")
        return probe
    try:
        env = json.loads(text)
        res = (env or {}).get("result") or {}
        content = res.get("content") or []
        body = {}
        if content and isinstance(content[0], dict) and isinstance(content[0].get("text"), str):
            body = json.loads(content[0]["text"])
        probe["ok"] = bool(body.get("ok", True))
        probe["reason"] = "answered" if probe["ok"] else "answered_error"
        probe["liveToolCount"] = body.get("toolCount")
        probe["contractVersion"] = ((body.get("mcp") or {}).get("contract")
                                    or body.get("contract"))
        probe["ide"] = body.get("ide")
    except Exception as exc:
        probe["reason"] = "unparseable_answer"
        probe["message"] = str(exc)[:200]
    return probe


def local_ide_list(peer=None):
    """The instance set + stale registrations + this bridge's identity.  READ-ONLY."""
    d = discover()
    redact = _is_remote_peer(peer)
    body = {
        "ok": True,
        "readOnly": True,
        "generatedAt": _iso_utc_now(),
        "instances": [_instance_facts(i, redact) for i in (d.get("instances") or [])],
        "stale": _stale_facts(d, redact),
        "registry": {"dir": d.get("dir"),
                     "count": len(d.get("instances") or []),
                     "roots": d.get("roots") or [],
                     "unreadable": bool(d.get("unreadable"))},
        "bridge": _bridge_facts(),
        "counts": {"instances": len(d.get("instances") or []),
                   "live": len(d.get("live") or []),
                   "stale": len(d.get("stale") or [])},
        "addInReachableRule": ADDIN_REACHABLE_RULE,
        "registryPolicy": ("ide_list is READ-ONLY: it never writes, prunes or deletes a "
                           "registration and it never opens a pipe session."),
    }
    if redact:
        body["registry"] = {"redacted": True}
        body["remoteRedacted"] = True
        body["redactionRule"] = REDACTION_RULE
    return body


def local_ide_status(params=None, peer=None):
    """ONE instance's full facts; probe=true adds ONE bounded live add-in check.  READ-ONLY."""
    params = params if isinstance(params, dict) else {}
    target = params.get("target")
    redact = _is_remote_peer(peer)
    probe_wanted = bool(params.get("probe"))
    try:
        probe_timeout = int(params.get("probeTimeoutMs") or 3000)
    except (TypeError, ValueError):
        probe_timeout = 3000
    probe_timeout = max(250, min(probe_timeout, 3000))    # HARD cap: a facts call never hangs
    d = discover()
    inst, refusal = select_target(target)
    body = {
        "ok": True,
        "readOnly": True,
        "generatedAt": _iso_utc_now(),
        "instance": None,
        "selection": _selection_facts(d, target, redact),
        "bridge": _bridge_facts(),
        "counts": {"instances": len(d.get("instances") or []),
                   "live": len(d.get("live") or []),
                   "stale": len(d.get("stale") or [])},
        "stale": _stale_facts(d, redact),
        "probeRequested": probe_wanted,
        "liveProbe": {"attempted": False, "ok": None, "reason": "not_requested",
                      "timeoutMs": probe_timeout,
                      "addInReachableRule": ADDIN_REACHABLE_RULE},
        "addInReachableRule": ADDIN_REACHABLE_RULE,
    }
    if redact:
        body["remoteRedacted"] = True
        body["redactionRule"] = REDACTION_RULE
    if inst is None:
        body["ok"] = False
        facts = _redact_refusal(refusal) if redact else dict(refusal or {})
        for k, v in facts.items():
            body.setdefault(k, v)
        return body
    body["instance"] = _instance_facts(inst, redact)
    if probe_wanted:
        selector = str(inst.pid) if inst.pid else (target or "")
        body["liveProbe"] = _probe_addin_bounded(inst, selector, probe_timeout)
        body["liveProbe"]["addInReachableRule"] = ADDIN_REACHABLE_RULE
    return body


LOCAL_METHODS = {
    "ide_list": lambda params, peer: local_ide_list(peer),
    "ide_status": lambda params, peer: local_ide_status(params, peer),
}


def serve_local_method(method, params, peer):
    body = LOCAL_METHODS[method](params or {}, peer)
    body["localMethod"] = method
    body["localMethods"] = list(LOCAL_RPC_METHODS)
    body["localMethodPriority"] = LOCAL_METHOD_PRIORITY
    return body


def handle_jsonrpc(req_text, selector=None, peer=None):
    """Serve ONE JSON-RPC message.  Returns (response_text|None, is_error, meta).

    peer is the CALLER address (the HTTP handler passes client_address[0]); None means the
    local stdio owner.  It decides ONLY the redaction of the two local facts methods - never
    dispatch, never a target - so a non-loopback caller cannot get facts /health hides.

    Contract face rules that must not drift:
      * tools/list result keys are EXACTLY {tools}   (mcp_strict_audit check 3b)
      * the JSON-RPC envelope carries NO _meta / nextCursor / ttlMs / cacheScope
      * a notification gets NO response body (HTTP 202, stdio none)
    """
    sel0 = selector if selector is not None else CFG.target
    try:
        req = json.loads(req_text)
    except Exception:
        _mark("json_parse_failed")
        return _jsonrpc_error(None, -32700, "Parse error"), True, {"reason": "parse_error"}
    _mark("json_parsed")
    if not isinstance(req, dict):
        return _jsonrpc_error(None, -32600, "Invalid Request"), True, {"reason": "not_object"}
    method = req.get("method")
    rid = req.get("id", None)
    if not isinstance(method, str) or not method:
        if rid is None:
            return None, False, {"notification": True}
        return _jsonrpc_error(rid, -32600, "Invalid Request"), True, {"reason": "no_method"}

    # LOCAL FACTS METHODS (JSON-RPC method namespace, NOT the add-in tool namespace).
    # Exactly the two reserved names are intercepted; every other method keeps the pre-existing
    # forward path byte-for-byte, which is what lets an older middle service answer -32601 for
    # these names (the hub's capability signal) without changing anything else.
    if method in LOCAL_METHODS:
        try:
            body = serve_local_method(method, req.get("params"), peer)
        except Exception as exc:
            dbg("local method %s failed: %r" % (method, exc))
            return _jsonrpc_error(rid, -32603, "local method failed: %s" % exc), True, {"local": True, "method": method}
        return _jsonrpc_result(rid, body), False, {"local": True, "method": method}

    if method.startswith("notifications/"):
        sel, req2 = _extract_target(req, sel0)
        out = forward_raw(json.dumps(req2, ensure_ascii=False), req2, sel)
        return None, False, {"notification": True, "transport": out.get("transport"),
                             "transportReason": out.get("transportReason"),
                             "flags": out.get("flags") or {}}

    if method == "initialize":
        out = forward_raw(req_text, req, sel0)
        if out.get("text") is not None:
            return out["text"], False, out
        body = local_initialize(req)
        meta = dict(out)
        meta["faceTools"] = "snapshot"
        meta["degraded"] = out.get("failure", {}).get("reason")
        return _jsonrpc_result(rid, body), False, meta

    if method == "tools/list":
        out = forward_raw(req_text, req, sel0)
        if out.get("text") is not None:
            return out["text"], False, out
        tools = snapshot()["tools"]
        reason = (out.get("failure") or {}).get("reason")
        if tools and reason in ("no_ide_running", "ambiguous_target", "target_mismatch",
                                "no_target"):
            meta = dict(out)
            meta["faceTools"] = "snapshot"
            return _jsonrpc_result(rid, {"tools": tools}), False, meta
        failure = out.get("failure") or {"ok": False, "reason": reason or "forward_failed"}
        return _jsonrpc_error(rid, -32000, "no tool face available: %s" % (reason or "unknown"),
                              failure), True, out

    if method == "tools/call":
        sel, req2 = _extract_target(req, sel0)
        tool_name = _tool_of(req2)
        args2 = _args_of(req2)
        # R-1021-03 A4 / ctl/1 sec 6: the TWO existing ide_state intents go to the control lane
        # FIRST when it is published; the business lane is the fallback and ONLY the fallback.
        # No new tool parameter exists - these are the pre-existing release_inflight /
        # cancel_inflight knobs, exactly as the add-in already defines them.
        intent = ctl_intent_of(req2)
        if intent is not None:
            inst_ctl, _ctl_refusal = select_target(sel)
            if inst_ctl is not None and ctl_available(inst_ctl):
                return ctl_intent_receipt(rid, inst_ctl, tool_name, intent)
            _mark("ctl_intent_business_fallback")
        req2_text = json.dumps(req2, ensure_ascii=False)
        # An in-flight call is REGISTERED here, before the wire, so an acknowledged ctl cancel can
        # land on it from another thread and survive into this receipt (ctl/1 sec 4.4).
        tok = inflight_begin(tool_name, rid)
        try:
            out = forward_raw(req2_text, req2, sel)
            if out.get("text") is None:
                body = dict(out.get("failure") or {"ok": False, "reason": "forward_failed"})
                note = _transport_note(out)
                for k, v in note.items():
                    if k not in body:
                        body[k] = v
                if out.get("ide"):
                    body["ide"] = out["ide"]
                body = ctl_cancel_override(body, tok)
                oop = oop_fallback_body(tool_name, args2, sel, body)
                if oop is not None:
                    meta = dict(out)
                    meta.update({"oop": True, "source": "out_of_process", "degraded": True,
                                 "why": (oop.get("oob") or {}).get("why")})
                    return _jsonrpc_result(rid, error_result(oop, True)), True, meta
                # Lead ruling 2: a SECOND, purely ADDITIVE trigger.  A host-modal box wedges the
                # business AND the ctl lane while the process stays alive; that state is worth
                # reporting, but it must never rewrite the failure that caused it.  So the
                # measured segment is APPENDED as outOfProcess and every field above stays as-is.
                add = oop_addendum(tool_name, args2, sel, body)
                if add is not None:
                    body["outOfProcess"] = add
                meta = out
                if body.get("cancelRequested"):
                    meta = dict(out)
                    meta["cancelRequested"] = True
                    meta["timeoutAlsoDue"] = body.get("timeoutAlsoDue")
                return _jsonrpc_result(rid, error_result(body, True)), True, meta
            try:
                resp_obj = json.loads(out["text"])
            except Exception:
                return out["text"], True, out
            res = resp_obj.get("result") if isinstance(resp_obj, dict) else None
            if isinstance(res, dict):
                note = _transport_note(out)
                parsed = None
                for c in (res.get("content") or []):
                    if c.get("type") == "text" and isinstance(c.get("text"), str):
                        try:
                            p = json.loads(c["text"])
                            if isinstance(p, dict):
                                parsed = p
                        except Exception:
                            parsed = None
                        break
                if parsed is not None and parsed.get("ok") is False:
                    note.setdefault("failureKind", "toolError")
                    note.setdefault("failureKindRule", FAILURE_KIND_RULE)
                resp_obj["result"] = stamp_body(res, out.get("ide"), contract_version(), note)
            return json.dumps(resp_obj, ensure_ascii=False), False, out
        finally:
            inflight_end(tok)

    out = forward_raw(req_text, req, sel0)
    if out.get("text") is not None:
        return out["text"], False, out
    failure = out.get("failure") or {"ok": False, "reason": "forward_failed"}
    return _jsonrpc_error(rid, -32000, failure.get("message") or "forward failed", failure), True, out


# ---------------------------------------------------------------- health ----------

def health_body():
    d = discover()
    face = "snapshot"
    if d["live"]:
        face = "live" if snapshot()["tools"] else "live(unsnapshotted)"
    auth = auth_state()
    if not auth["enabled"]:
        # Lead ruling (item 3a): the weak configuration must be visible HERE too, not
        # only in the startup log.
        auth["guidance"] = ("multi-user machine: set VB6MCP_TOKEN (or VB6MCP_TOKEN_FILE) - on a "
                            "shared box ANY local user can reach this loopback port and drive "
                            "the whole tool face")
    return {
        "ok": bool(d["live"]),
        "server": PRODUCT,
        "version": VERSION,
        "port": CFG.port,
        "auth": auth,
        "netstat": False,
        "transport": CFG.transport or "auto(pipe-first)",
        "pipeFallback": CFG.pipe_fallback,
        "pipeConnectMs": CFG.pipe_connect_ms,
        "forwardTimeoutMs": CFG.forward_timeout_ms,
        "progressMs": CFG.progress_ms,
        "pipeIdleMs": CFG.pipe_idle_ms,
        # R-1021-03: the control lane is a CLIENT capability here; whether an individual
        # instance publishes it is per-instance facts (ide_status.ctl).
        "ctlLane": {"proto": CTL_PROTO, "enabled": bool(CFG.ctl_enabled),
                    "helloMs": CFG.ctl_hello_ms, "ackMs": CFG.ctl_ack_ms,
                    "fastTools": list(CTL_FAST_TOOLS),
                    "outOfProcessFallback": bool(CFG.oop_enabled),
                    "maxMs": CFG.max_call_ms, "maxCallHardMs": CFG.max_call_hard_ms},
        "contractVersion": contract_version(),
        "faceTools": face,
        "snapshotTools": len(snapshot()["tools"]),
        "registryDir": d["dir"],
        "instances": [i.brief() for i in d["instances"]],
        "liveInstances": [i.brief() for i in d["live"]],
        "staleRegistry": d["stale"],
    }


# ---------------------------------------------------------------- HTTP face -------

CORS = {
    "Access-Control-Allow-Origin": "*",
    "Access-Control-Allow-Methods": "POST, GET, OPTIONS, DELETE",
    "Access-Control-Allow-Headers": "Content-Type, MCP-Protocol-Version, MCP-Session-Id",
    "Access-Control-Expose-Headers": "Mcp-Session-Id",
}


class BridgeHandler(http.server.BaseHTTPRequestHandler):
    protocol_version = "HTTP/1.1"
    server_version = PRODUCT + "/" + VERSION
    sys_version = ""

    # -- plumbing -------------------------------------------------------------
    def log_message(self, fmt, *args):
        dbg("http", fmt % args)

    def setup(self):
# probe: t0 is the moment this connection's handler started, i.e.
        # BEFORE the request line and headers are parsed by http.server.
        http.server.BaseHTTPRequestHandler.setup(self)
        _t0()
        _mark("socket_setup")

    def _send_raw(self, status, ctype, body, extra=None, content_length=True):
        if isinstance(body, str):
            body = body.encode("utf-8")
        self.send_response(status)
        if ctype:
            self.send_header("Content-Type", ctype)
        for k, v in CORS.items():
            self.send_header(k, v)
        for k, v in (extra or {}).items():
            self.send_header(k, v)
        if content_length:
            self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        if body:
            try:
                self.wfile.write(body)
            except Exception:
                pass

    def _read_body(self):
        """(body_bytes, error_response_or_None).  Content-Length or chunked."""
        te = (self.headers.get("Transfer-Encoding") or "").lower()
        if "chunked" in te:
            return self._read_chunked()
        try:
            n = int(self.headers.get("Content-Length") or 0)
        except ValueError:
            n = 0
        if n > HTTP_BODY_MAX:
            return None, (413, "text/plain", "Request Entity Too Large")
        if n <= 0:
            return b"", None
        return self.rfile.read(n), None

    def _read_chunked(self):
        chunks = []
        total = 0
        while True:
            line = self.rfile.readline(65536).strip()
            if not line:
                break
            try:
                size = int(line.split(b";")[0], 16)
            except ValueError:
                return None, (400, "text/plain", "Bad chunked encoding")
            if size == 0:
                self.rfile.readline(65536)
                break
            total += size
            if total > HTTP_BODY_MAX:
                return None, (413, "text/plain", "Request Entity Too Large")
            chunks.append(self.rfile.read(size))
            self.rfile.read(2)
        return b"".join(chunks), None

    # -- verbs ----------------------------------------------------------------
    # -- authentication (TOKEN_THREAT_MODEL.md M-1..M-4) ---------------------------
    def _url_tokens(self):
        """Query parameters that a caller must never put a credential in (M-3)."""
        if "?" not in self.path:
            return []
        q = self.path.split("?", 1)[1]
        bad = []
        for pair in q.split("&"):
            if not pair:
                continue
            k = pair.split("=", 1)[0].strip().lower()
            if k in ("token", "access_token", "api_key", "apikey", "auth", "authorization"):
                bad.append(k)
        return bad

    def _remote_gate(self):
        """None when the request may proceed, else (status, reason, message, extra).

        The LOOPBACK tier returns None immediately, so its behaviour stays byte-for-byte the
        pre-existing one.  Remote mode adds: source allowlist (403), an auditable log (503) and
        an auth lockout (429) - all BEFORE the token check, so a locked-out source learns
        nothing about the credential.
        """
        if not _REMOTE["enabled"]:
            return None
        peer = self.client_address[0] if self.client_address else ""
        if peer_is_loopback(peer):
            return None
        self._remote_peer = peer
        if not source_allowed(peer):
            return (403, "source_not_allowed",
                    "this source (%s) is not inside --allow-remote %r" % (peer, _REMOTE["allowRaw"]), {})
        if not _audit_writable():
            return (503, "audit_unavailable",
                    "remote requests require an append-only audit log and it is not writable", {})
        remain = auth_lock_remaining(peer)
        if remain > 0:
            return (429, "auth_locked_out",
                    "too many failed authentications from %s; retry in %ss" % (peer, remain),
                    {"Retry-After": str(remain)})
        return None

    def _reject_gate(self, status, reason, message, extra):
        body = json.dumps({"ok": False, "reason": reason, "message": message},
                          ensure_ascii=False, indent=2)
        self.close_connection = True
        self._send(status, "application/json", body, dict(extra or {}))
        return

    def _auth_problem(self):
        """None when the request may proceed, else (reason, message).  Never returns a secret."""
        if not CFG.auth_enabled():
            return None
        bad = self._url_tokens()
        if bad:
            return ("token_in_url_rejected",
                    "a credential must not travel in the URL (%s); send it as "
                    "Authorization: Bearer <token> or X-VB6MCP-Token" % ",".join(bad))
        provided = ""
        authz = self.headers.get("Authorization") or ""
        if authz:
            parts = authz.split(None, 1)
            if len(parts) == 2 and parts[0].lower() == "bearer":
                provided = parts[1].strip()
            else:
                return ("token_bad_scheme",
                        "Authorization must be 'Bearer <token>'")
        if not provided:
            provided = (self.headers.get("X-VB6MCP-Token") or "").strip()
        if not provided:
            return ("token_missing",
                    "this bridge requires a token; send Authorization: Bearer <token> "
                    "or X-VB6MCP-Token")
        if not token_ok(provided):
            # count the failure; after AUTH_FAIL_THRESHOLD the source is locked out
            # with a doubling wait (auth_locked_out), so an online brute force is not free.
            count = auth_record_failure(self.client_address[0] if self.client_address else "")
            bridge_log("auth_failure", source=self.client_address[0] if self.client_address else "",
                       count=count)
            return ("token_mismatch", "the supplied token does not match")
        auth_record_success(self.client_address[0] if self.client_address else "")
        return None

    def _reject_auth(self, reason, message):
        bridge_log("auth_reject", reason=reason, path=self.path.split("?", 1)[0],
                   method=self.command, source=CFG.token_source, tokenLen=len(CFG.token))
        body = json.dumps({
            "ok": False, "reason": reason, "message": message,
            "auth": {"required": True, "scheme": "bearer", "source": CFG.token_source,
                     "header": "Authorization: Bearer <token>"},
            "hint": "the bridge authenticates the ENDPOINT, not a user's intent; a client that "
                    "legitimately holds the token is outside this control",
        }, ensure_ascii=False, indent=2)
        self.close_connection = True
        self._send(401, "application/json", body, {"WWW-Authenticate": 'Bearer realm="vb6mcp"'})

    def _send(self, code, ctype, payload, headers=None):
        """: the ONE response exit. Every response (refusals included) is audited once."""
        try:
            if not getattr(self, "_audited", False):
                self._audited = True
                peer = self.client_address[0] if self.client_address else ""
                if _REMOTE["enabled"] and not peer_is_loopback(peer) and _REMOTE["auditPath"]:
                    audit_write("request",
                                source=peer,
                                method=getattr(self, "command", ""),
                                path=(getattr(self, "path", "") or "").split("?", 1)[0],
                                status=code,
                                tool=getattr(self, "_audit_tool", None),
                                result=getattr(self, "_audit_result", None),
                                tokenId=("current" if getattr(self, "_audit_token", "") == "current"
                                         else (getattr(self, "_audit_token", "") or None)))
        except Exception:
            pass
        return self._send_raw(code, ctype, payload, headers)

    def do_OPTIONS(self):
        # CORS preflight carries no Authorization header by design, so it is never gated.
        self._send(204, "text/plain", b"")

    def do_GET(self):
        gate = self._remote_gate()
        if gate is not None:
            self._reject_gate(*gate)
            return
        problem = self._auth_problem()
        if problem is not None:
            self._reject_auth(*problem)
            return
        path = self.path.split("?", 1)[0]
        if path in ("/health", "/ping"):
            hb = health_body()
            peer = self.client_address[0] if self.client_address else ""
            if not peer_is_loopback(peer):
                # a remote caller never learns instance paths or pipe names.
                hb = _redact_health(hb)
            self._send(200, "application/json", json.dumps(hb, ensure_ascii=False, indent=2))
        elif path == "/":
            self._send(200, "text/plain", "%s %s is running on port %d (middle service layer).\n"
                       % (PRODUCT, VERSION, CFG.port))
        elif path == "/mcp":
            self._send(405, "text/plain", "Streamable HTTP endpoint requires POST (or DELETE)",
                       {"Allow": "POST, DELETE"})
        elif path == "/sse":
            self._send(501, "application/json", json.dumps({
                "ok": False, "reason": "sse_transport_not_implemented",
                "message": ("the legacy MCP SSE transport (GET /sse + POST /message) is not "
                            "implemented by the Python middle service layer; use POST /rpc or "
                            "POST /mcp, which are the faces the in-scope tools use"),
            }, ensure_ascii=False, indent=2))
        else:
            self._send(404, "text/plain", "Not Found")

    def do_DELETE(self):
        gate = self._remote_gate()
        if gate is not None:
            self._reject_gate(*gate)
            return
        problem = self._auth_problem()
        if problem is not None:
            self._reject_auth(*problem)
            return
        path = self.path.split("?", 1)[0]
        if path == "/mcp":
            self._send(200, "application/json", b"")
        else:
            self._send(404, "text/plain", "Not Found")

    def do_POST(self):
        gate = self._remote_gate()
        if gate is not None:
            self._reject_gate(*gate)
            return
        problem = self._auth_problem()
        if problem is not None:
            self._reject_auth(*problem)
            return
        path = self.path.split("?", 1)[0]
        if path == "/message":
            self._send(501, "application/json", json.dumps({
                "ok": False, "reason": "sse_transport_not_implemented",
                "message": "POST /message belongs to the legacy SSE transport, which this "
                           "service does not implement; use POST /rpc or POST /mcp",
            }, ensure_ascii=False, indent=2))
            return
        if path not in ("/rpc", "/mcp"):
            self._send(404, "text/plain", "Not Found")
            return
        _mark("path_routed")
        body, err = self._read_body()
        _mark("body_read")
        if err is not None:
            # the request body was NOT drained: never keep this connection alive
            self.close_connection = True
            self._send(err[0], err[1], err[2], {"Connection": "close"})
            return
        req_text = body.decode("utf-8-sig", "replace")
        # a non-loopback caller is READ-ONLY unless the service was started with
        # --remote-write.  Deny-FIRST: anything not on the read-only allowlist is refused.
        _peer = self.client_address[0] if self.client_address else ""
        if _REMOTE["enabled"] and not peer_is_loopback(_peer):
            try:
                _req = json.loads(req_text)
                _tool = ((_req.get("params") or {}).get("name") if isinstance(_req, dict) else None)
            except Exception:
                _tool = None
            self._audit_tool = _tool
            if _req_is_tool_call(req_text):
                _ok, _reason, _msg = remote_tool_allowed(_tool, True)
                if not _ok:
                    self._audit_result = _reason
                    _rid = None
                    try:
                        _rid = json.loads(req_text).get("id")
                    except Exception:
                        _rid = None
                    refusal = _jsonrpc_result(_rid, {"content": [{"type": "text", "text": json.dumps(
                        {"ok": False, "reason": _reason, "message": _msg,
                         "remoteMode": "read-only" if not _REMOTE["write"] else "read-write",
                         "hint": "start the bridge with --remote-write if the network really must "
                                 "drive write/execute tools"},
                        ensure_ascii=False)}]})
                    self._send(200, "application/json", refusal.encode("utf-8"))
                    return
        target = self.headers.get("X-VB6MCP-Target") or CFG.target
        try:
            _peer_addr = self.client_address[0] if self.client_address else None
            resp_text, _is_error, _meta = handle_jsonrpc(req_text, target, _peer_addr)
        except Exception as exc:
            dbg("dispatch failed:", repr(exc))
            resp_text = _jsonrpc_error(None, -32603, "Internal error: %s" % exc)
        _mark("dispatch_done")
        if resp_text is None:
            _mark("send_begin")
            self._send(202, "application/json", b"")
            _mark("send_done")
            _timing_emit({"path": path, "reqLen": len(req_text), "respLen": 0})
            return
        accept = (self.headers.get("Accept") or "").lower()
        if "text/event-stream" in accept and "application/json" not in accept:
            sse = "event: message\r\ndata: " + resp_text + "\r\n\r\n"
            self._send(200, "text/event-stream", sse.encode("utf-8"))
            return
        extra = {}
        if '"method":"initialize"' in req_text.replace(" ", ""):
            extra["Mcp-Session-Id"] = str(uuid.uuid4())
        _mark("send_begin")
        self._send(200, "application/json", resp_text.encode("utf-8"), extra)
        _mark("send_done")
        _timing_emit({"path": path, "reqLen": len(req_text), "respLen": len(resp_text)})


class BridgeServer(http.server.ThreadingHTTPServer):
    daemon_threads = True
    # deliberately NOT allow_reuse_address: this process must be the ONE listener on the port
    allow_reuse_address = False


# ---------------------------------------------------------------- stdio -----------

def run_stdio(selector):
    out = sys.stdout.buffer
    for raw in sys.stdin.buffer:
        line = raw.strip()
        if not line:
            continue
        try:
            req_text = line.decode("utf-8", "replace")
        except Exception:
            continue
        try:
            resp_text, _is_error, _meta = handle_jsonrpc(req_text, selector)
        except Exception as exc:
            resp_text = _jsonrpc_error(None, -32603, "Internal error: %s" % exc)
        if resp_text:
            out.write(resp_text.encode("utf-8") + b"\n")
            out.flush()
    return 0


# ---------------------------------------------------------------- main ------------

# =====================================================================================
# REMOTE MODE (the specification is docs/design/REMOTE_ACCESS.md).
#
# The DEFAULT tier is unchanged: loopback only, no TLS requirement, and a non-loopback
# host still exits with EXIT_LOOPBACK_REFUSED (6) unless remote mode is explicitly set up.
# Remote mode is FAIL-CLOSED and needs ALL THREE of: --allow-remote <cidr[,cidr]>, a TLS
# certificate+key, and a token source.  On top of that:
#   * the source address is re-checked INSIDE the process (403 source_not_allowed),
#   * a non-loopback caller is READ-ONLY by default (remote_read_only) unless --remote-write,
#   * tokens are compared in constant time, a SECOND (previous) token stays valid for rotation,
#   * auth failures back off exponentially per source (then auth_locked_out),
#   * every response (refusals included) appends exactly ONE JSONL audit line, and an
#     unwritable audit log refuses remote requests outright (audit_unavailable),
#   * /health never shows instance paths or pipe names to a remote caller.
# =====================================================================================
EXIT_LOOPBACK_REFUSED = 6
EXIT_REMOTE_MISCONFIGURED = 7

#: Deny-first: a remote caller may only use tools on this READ-ONLY list.  A tool that is not
#: listed (including anything new) is refused with remote_read_only, so a forgetful allowlist
#: fails CLOSED instead of silently handing the write face to the network.
REMOTE_READ_ONLY_TOOLS = {
    "ide_state", "ide_list", "ide_target", "mcp_info", "audit_tail", "todo_status",
    "current_position", "desktop_probe", "execution_state", "read_active_dialogs",
    "list_windows", "list_addins", "list_component_parts", "list_files", "list_lessons",
    "get_project_info", "get_component_list", "get_component_code", "get_component_properties",
    "get_codepane", "get_selection", "get_device", "get_device_at", "get_device_properties",
    "get_engine_status", "get_ic", "get_model_path", "get_param_value", "get_sim_mode",
    "get_sim_time", "get_simulations", "get_status", "get_symbol_info", "get_display_mode",
    "read_file", "read_source", "read_point", "read_monitor", "read_bytes",
    "search_in_files", "search_point", "search_model_object", "search_procedures", "search_strings",
    "find_in_component", "find_gui", "find_xrefs_to_name",
    "check_component", "check_references", "check_public_api", "check_write_paths", "check_schema_baseline",
    "validate_code", "verify", "verify_scope", "verify_reconstruction",
    "export_public_api", "export_component", "export_evidence_bundle", "export_lessons",
    "dependency_graph", "tree_list_nodes", "tree_node_info", "tree_find_node",
    "units_of", "grep", "glob", "comment", "inline_comment", "list_bookmarks", "list_breakpoints",
    "list_documents", "list_names", "list_procedures", "list_segments", "list_strings",
    "list_architectures", "list_unknowns", "binary_overview", "binary_session",
    "analysis_view", "inspect_analysis_view", "procedure_info", "procedure_assemblY",
    "procedure_pseudo_code", "procedure_callers", "procedure_callees", "procedure_references",
    "current_address", "current_document", "current_procedure", "address_name",
    "read_function_instructions", "get_evidence_bundle", "get_navigation_context",
}
#: Even when the mode is --remote-write these stay refused from a remote peer (defence in depth).
REMOTE_NEVER_TOOLS = {"quit_ide"}

_REMOTE = {
    "allow": (),            # tuple of ipaddress networks
    "allowRaw": "",
    "tlsCert": "", "tlsKey": "",
    "write": False,
    "tokenPrev": "",
    "auditPath": "",
    "enabled": False,
}
_AUTH_LOCK = threading.Lock()
_AUTH_FAILS = {}            # source -> {"count": n, "until": epoch_seconds}
AUTH_FAIL_THRESHOLD = 5     # then back off: 30 s, doubling, capped at 300 s
AUTH_BACKOFF_BASE_S = 30
AUTH_BACKOFF_MAX_S = 300


def remote_configure(allow_remote, tls_cert, tls_key, remote_write, token_prev, audit_path):
    """Set the remote-mode policy once, at startup.  Returns the list of MISSING prerequisites.

    Nothing here enables anything by accident: an empty allow_remote keeps the loopback tier.
    """
    nets = []
    raw = (allow_remote or "").strip()
    if raw:
        for part in raw.split(","):
            part = part.strip()
            if not part:
                continue
            try:
                nets.append(ipaddress.ip_network(part, strict=False))
            except ValueError as exc:
                _REMOTE["allowError"] = "bad cidr %r: %s" % (part, exc)
    _REMOTE["allow"] = tuple(nets)
    _REMOTE["allowRaw"] = raw
    _REMOTE["tlsCert"] = (tls_cert or "").strip()
    _REMOTE["tlsKey"] = (tls_key or "").strip()
    _REMOTE["write"] = bool(remote_write)
    _REMOTE["tokenPrev"] = (token_prev or "").strip()
    _REMOTE["auditPath"] = (audit_path or os.environ.get("VB6MCP_AUDIT_LOG", "") or "").strip()
    _REMOTE["enabled"] = bool(raw)
    return remote_missing()


def remote_missing():
    """Every prerequisite the CURRENT configuration lacks (empty list = remote is complete)."""
    if not _REMOTE["enabled"]:
        return []
    missing = []
    if _REMOTE.get("allowError"):
        missing.append("--allow-remote (%s)" % _REMOTE["allowError"])
    if not _REMOTE["allow"]:
        missing.append("--allow-remote <cidr[,cidr]>")
    if not _REMOTE["tlsCert"] or not _REMOTE["tlsKey"]:
        missing.append("--tls-cert/--tls-key")
    elif not (os.path.isfile(_REMOTE["tlsCert"]) and os.path.isfile(_REMOTE["tlsKey"])):
        missing.append("--tls-cert/--tls-key (file not found)")
    if not CFG.auth_enabled():
        missing.append("token source (VB6MCP_TOKEN or VB6MCP_TOKEN_FILE)")
    if not _REMOTE["auditPath"]:
        missing.append("--audit-log/VB6MCP_AUDIT_LOG (remote requests MUST be auditable)")
    elif not _audit_writable():
        missing.append("--audit-log (path is not writable)")
    return missing


def remote_enabled():
    return bool(_REMOTE["enabled"])


def peer_is_loopback(peer):
    return is_loopback_host(str(peer or ""))


def source_allowed(peer):
    """True when a REMOTE peer is inside --allow-remote.  Loopback is always allowed."""
    if peer_is_loopback(peer):
        return True
    if not _REMOTE["enabled"]:
        return False
    try:
        addr = ipaddress.ip_address(str(peer))
    except ValueError:
        return False
    for net in _REMOTE["allow"]:
        if addr in net:
            return True
    return False


def remote_tool_allowed(tool, remote_peer):
    """(allowed, reason, message) for one tool call from a (possibly remote) peer."""
    if not remote_peer:
        return True, "", ""
    name = str(tool or "")
    if name in REMOTE_NEVER_TOOLS:
        return False, "remote_read_only", ("%s is never executable from a remote caller" % name)
    if _REMOTE["write"]:
        return True, "", ""
    if name in REMOTE_READ_ONLY_TOOLS:
        return True, "", ""
    return False, "remote_read_only", (
        "%r is not on the remote READ-ONLY allowlist: a non-loopback caller never gets the "
        "write/execute face unless the service was started with --remote-write" % name)


def token_ok(provided):
    """Constant-time check against the CURRENT token and, during rotation, the PREVIOUS one."""
    if token_matches(provided, CFG.token):
        return True
    prev = _REMOTE["tokenPrev"]
    if prev and token_matches(provided, prev):
        return True
    return False


def _audit_writable():
    p = _REMOTE["auditPath"]
    if not p:
        return False
    try:
        d = os.path.dirname(os.path.abspath(p))
        if d and not os.path.isdir(d):
            os.makedirs(d, exist_ok=True)
        with open(p, "a", encoding="utf-8") as fh:
            fh.write("")
        return True
    except OSError:
        return False


def audit_write(event, **fields):
    """Append ONE JSONL line.  Never raises, never records a token value (M-4)."""
    p = _REMOTE["auditPath"]
    if not p:
        return False
    rec = {"ts": time.strftime("%Y-%m-%dT%H:%M:%S"), "event": event}
    for k, v in fields.items():
        if k.lower() in ("token", "authorization", "x-vb6mcp-token", "secret", "credential"):
            rec[k] = "<redacted>"
        else:
            rec[k] = v
    try:
        with open(p, "a", encoding="utf-8") as fh:
            fh.write(json.dumps(rec, ensure_ascii=False) + chr(10))
        return True
    except OSError:
        return False


def auth_lock_remaining(peer):
    """Seconds a source is still locked out (0 = it may try now)."""
    now = time.time()
    with _AUTH_LOCK:
        st = _AUTH_FAILS.get(str(peer or ""))
        if not st:
            return 0
        if st["until"] <= now:
            return 0
        return int(st["until"] - now) + 1


def auth_record_failure(peer):
    """Record one auth failure; after AUTH_FAIL_THRESHOLD they are locked out with a doubled wait."""
    now = time.time()
    with _AUTH_LOCK:
        st = _AUTH_FAILS.setdefault(str(peer or ""), {"count": 0, "until": 0, "wait": 0})
        st["count"] += 1
        if st["count"] >= AUTH_FAIL_THRESHOLD:
            wait = st.get("wait") or AUTH_BACKOFF_BASE_S
            st["until"] = now + wait
            st["wait"] = min(int(wait) * 2, AUTH_BACKOFF_MAX_S)
        return st["count"]


def auth_record_success(peer):
    with _AUTH_LOCK:
        _AUTH_FAILS.pop(str(peer or ""), None)


def _req_is_tool_call(req_text):
    try:
        j = json.loads(req_text)
    except Exception:
        return False
    return isinstance(j, dict) and j.get("method") == "tools/call"


def _jsonrpc_result(req_id, result):
    return json.dumps({"jsonrpc": "2.0", "id": req_id, "result": result}, ensure_ascii=False)


def _redact_health(body):
    """Strip discovery detail (instance paths, pipe names, pids) for a non-loopback caller."""
    try:
        out = dict(body or {})
        insts = out.get("instances") or []
        out["instances"] = [{"project": (i or {}).get("project"), "port": (i or {}).get("port"),
                             "transport": (i or {}).get("transport"), "redacted": True} for i in insts]
        for k in ("pipe", "pipeName", "registryFile", "registryRoot", "registryDir", "projectPath",
                  "pid", "instanceId", "dllSha256", "logPath", "dataDirs", "sourceDir"):
            out.pop(k, None)
        out["remoteRedacted"] = True
        return out
    except Exception:
        return {"ok": True, "remoteRedacted": True}


def _tls_fingerprint(cert_path):
    """sha256 of the DER certificate - the value a client pins (never a secret)."""
    try:
        with open(cert_path, "rb") as fh:
            pem = fh.read()
        if b"-----BEGIN CERTIFICATE-----" in pem:
            import base64 as _b64
            body = pem.split(b"-----BEGIN CERTIFICATE-----", 1)[1].split(b"-----END CERTIFICATE-----", 1)[0]
            der = _b64.b64decode(b"".join(body.split()))
        else:
            der = pem
        return hashlib.sha256(der).hexdigest()
    except Exception:
        return ""


def serve(port, host, selector):
    if not is_loopback_host(host):
        # M-0 is now the DEFAULT tier, not the only one.  Without --allow-remote the
        # old invariant holds EXACTLY (exit 6, no TLS requirement for loopback).  With it, remote
        # mode must be COMPLETE (allow-remote + TLS cert/key + token source + an auditable log);
        # any gap refuses to start and names every missing item - never a silent downgrade.
        missing = remote_missing()
        if not _REMOTE["enabled"]:
            sys.stderr.write(
                "vb6mcp_bridge: REFUSED - %r is not a loopback address.\n"
                "  This service exposes the whole tool face (110 tools, including WRITE tools that\n"
                "  edit files and run code in the IDE).  The DEFAULT tier is loopback-only\n"
                "  (ADR sec 0.5 / TOKEN_THREAT_MODEL.md M-0); remote binding requires ALL of\n"
                "  --allow-remote <cidr> + --tls-cert/--tls-key + a token source (see\n"
                "  docs/design/REMOTE_ACCESS.md).  Missing: --allow-remote.\n" % (host,))
            bridge_log("bind_refused", host=host, reason="non_loopback_host")
            return EXIT_LOOPBACK_REFUSED
        if missing:
            sys.stderr.write(
                "vb6mcp_bridge: REFUSED - remote mode is INCOMPLETE for %r.\n"
                "  Missing: %s\n"
                "  Remote mode is fail-closed: it needs --allow-remote, --tls-cert/--tls-key, a\n"
                "  token source and an auditable --audit-log (docs/design/REMOTE_ACCESS.md).\n"
                % (host, "; ".join(missing)))
            bridge_log("bind_refused", host=host, reason="remote_misconfigured", missing=missing)
            return EXIT_REMOTE_MISCONFIGURED
    if port == RESERVED_PORT:
        sys.stderr.write("vb6mcp_bridge: REFUSED - port %d is reserved by this repository.\n"
                         % RESERVED_PORT)
        return 3
    try:
        httpd = BridgeServer((host, port), BridgeHandler)
    except OSError as exc:
        sys.stderr.write("vb6mcp_bridge: cannot bind %s:%d -> %s\n"
                         "  (this process must be the ONLY listener on its port; no silent "
                         "port hopping)\n" % (host, port, exc))
        return 4
    CFG.port = port
    scheme = "http"
    if _REMOTE["enabled"] and _REMOTE["tlsCert"]:
        try:
            ctx = ssl.SSLContext(ssl.PROTOCOL_TLS_SERVER)
            ctx.minimum_version = ssl.TLSVersion.TLSv1_2
            ctx.load_cert_chain(_REMOTE["tlsCert"], _REMOTE["tlsKey"])
            httpd.socket = ctx.wrap_socket(httpd.socket, server_side=True)
            scheme = "https"
            fp = _tls_fingerprint(_REMOTE["tlsCert"])
        except Exception as exc:
            sys.stderr.write("vb6mcp_bridge: REFUSED - TLS could not be enabled: %r\n" % (exc,))
            bridge_log("bind_refused", host=host, reason="tls_setup_failed", error=str(exc))
            return EXIT_REMOTE_MISCONFIGURED
    sys.stderr.write("vb6mcp_bridge %s listening on %s://%s:%d  (/rpc, /mcp, /health)\n"
                     % (VERSION, scheme, host, port))
    if _REMOTE["enabled"]:
        # LOUD: an operator must never wonder whether the write face is exposed to the network.
        _mode = "READ-WRITE (--remote-write)" if _REMOTE["write"] else "READ-ONLY"
        sys.stderr.write("vb6mcp_bridge REMOTE MODE ENABLED: allow=%s tls=%s fingerprint=%s\n"
                         "  remote callers: %s; audit=%s; token rotation=%s\n"
                         % (_REMOTE["allowRaw"], _REMOTE["tlsCert"], fp or "(none)",
                            _mode, _REMOTE["auditPath"],
                            "previous token also valid" if _REMOTE["tokenPrev"] else "single token"))
        bridge_log("remote_mode", allow=_REMOTE["allowRaw"], write=_REMOTE["write"],
                   audit=_REMOTE["auditPath"], tlsFingerprint=fp,
                   tokenRotation=bool(_REMOTE["tokenPrev"]))
    # ---- auth state, to stderr AND the log (Lead ruling item 3b) -------------------
    _a = auth_state()
    _line = ("vb6mcp_bridge auth: enabled=%s required=%s source=%s tokenLen=%d"
             % (_a["enabled"], _a["required"], _a["source"], _a["tokenLen"]))
    sys.stderr.write(_line + "\n")
    bridge_log("auth_state", enabled=_a["enabled"], source=_a["source"], tokenLen=_a["tokenLen"],
               note=_a["note"])
    if not _a["enabled"]:
        _warn = ("vb6mcp_bridge WARNING: no token configured.  This is acceptable ONLY on a "
                 "single-user loopback setup.  On a multi-user machine ANY local account can "
                 "reach 127.0.0.1:%d and drive the whole tool face (including write tools) - "
                 "set VB6MCP_TOKEN or VB6MCP_TOKEN_FILE." % port)
        sys.stderr.write(_warn + "\n")
        bridge_log("auth_warning", reason="no_token_configured", multiUser="configure_a_token")
    # ALWAYS-ON startup diagnostics (never a 0-byte log again): which registry directories were
    # read, how many entries each held, how many were live, and why the rest were dropped.
    try:
        d0 = discover(force=True)
        roots = "; ".join("%s (%d entries)" % (r["root"], r["entries"]) for r in (d0.get("roots") or []))
        sys.stderr.write("vb6mcp_bridge diagnostics: configuredDir=%s entries=%d live=%d\n"
                         "  scanned: %s\n"
                         "  stale: %s\n  log=%s\n"
                         % (d0.get("dir"), len(d0.get("instances") or []), len(d0.get("live") or []),
                            roots or "(none)",
                            json.dumps(d0.get("stale") or [], ensure_ascii=False)[:400],
                            CFG.log_path or "(stderr only; set VB6MCP_BRIDGE_LOG)"))
    except Exception as exc:
        sys.stderr.write("vb6mcp_bridge diagnostics FAILED: %r\n" % (exc,))
    bridge_log("start", version=VERSION, host=host, port=port,
               frozen=bool(getattr(sys, "frozen", False)), pid=os.getpid())
    sys.stderr.flush()
    try:
        httpd.serve_forever(poll_interval=0.2)
    except KeyboardInterrupt:
        pass
    finally:
        httpd.server_close()
        POOL.close_all()
    return 0


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0],
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--port", type=int, default=None, help="TCP port (default %d)" % DEFAULT_PORT)
    ap.add_argument("--host", default=None)
    ap.add_argument("--target", default=None, help="_target default for every call")
    ap.add_argument("--transport", default=None, choices=["auto", "pipe", "tcp"])
    ap.add_argument("--stdio", action="store_true", help="serve JSON-RPC line-delimited on stdio")
    ap.add_argument("--selftest", action="store_true", help="run the in-process acceptance harness")
    ap.add_argument("--with-battery", action="store_true",
                    help="--selftest also runs tests/battery_test.py (slow)")
    ap.add_argument("--smoke", action="store_true",
                    help="RUNTIME acceptance for a frozen build: bundle-own stub + the real "
                         "HTTP face, one JSON object per line (never uses the repo scripts)")
    ap.add_argument("--smoke-seconds", type=int, default=60,
                    help="how long --smoke keeps serving (default 60)")
    ap.add_argument("--print-config", action="store_true")
# ---- remote mode (docs/design/REMOTE_ACCESS.md)
    ap.add_argument("--allow-remote", default=None,
                    help="comma-separated CIDRs allowed to reach a NON-loopback bind; remote mode "
                         "also needs --tls-cert/--tls-key, a token source and --audit-log")
    ap.add_argument("--tls-cert", default=None, help="PEM certificate for the HTTPS face")
    ap.add_argument("--tls-key", default=None, help="PEM private key for the HTTPS face")
    ap.add_argument("--remote-write", action="store_true",
                    help="allow a remote caller to use write/execute tools (default: read-only)")
    ap.add_argument("--token-prev", default=None,
                    help="previous token that stays valid during rotation (or VB6MCP_TOKEN_PREV)")
    ap.add_argument("--audit-log", default=None,
                    help="append-only JSONL audit log; REQUIRED for remote mode")
    ap.add_argument("--stub-addin", action="store_true",
                    help="DEV ONLY: run a stub np/1 pipe add-in in its own process "
                         "(used by --selftest; env VB6MCP_STUB_PIPE/NONCE/...)")
    args = ap.parse_args(argv)

    CFG.resolve_token()
    remote_configure(args.allow_remote, args.tls_cert, args.tls_key, args.remote_write,
                     args.token_prev or os.environ.get("VB6MCP_TOKEN_PREV", ""),
                     args.audit_log or os.environ.get("VB6MCP_AUDIT_LOG", ""))
    if args.port is not None:
        CFG.port = args.port
    if args.host:
        CFG.host = args.host
    if args.target is not None:
        CFG.target = args.target
    if args.transport:
        CFG.transport = args.transport

    if args.print_config:
        print(json.dumps({
            "product": PRODUCT, "version": VERSION, "port": CFG.port, "host": CFG.host,
            "transport": CFG.transport or "auto(pipe-first)", "strict": CFG.strict(),
            "pipeConnectMs": CFG.pipe_connect_ms, "forwardTimeoutMs": CFG.forward_timeout_ms,
            "progressMs": CFG.progress_ms, "pipeIdleMs": CFG.pipe_idle_ms,
            "instancesDir": CFG.instances_dir, "snapshot": SNAP_PATH,
            "hubTools": HUB_TOOLS_PATH, "frozen": bool(getattr(sys, "frozen", False)),
            "dataDirs": DATA_DIRS, "sourceDir": _SRC_DIR,
            "logPath": CFG.log_path or None, "instancesScan": _env_bool("VB6MCP_INSTANCES_SCAN", True),
            "registryRoots": registry_roots(),
            "authEnabled": CFG.auth_enabled(), "authSource": CFG.token_source,
            "tokenLen": len(CFG.token), "tokenNote": CFG.token_note or None,
            "loopbackOnly": (not remote_enabled()),
            "remoteMode": remote_enabled(),
            "allowRemote": _REMOTE["allowRaw"] or None,
            "remoteWrite": _REMOTE["write"],
            "tlsCert": _REMOTE["tlsCert"] or None,
            "auditLog": _REMOTE["auditPath"] or None,
            "tokenRotation": bool(_REMOTE["tokenPrev"]),
            "remoteMissing": remote_missing(),
        }, ensure_ascii=False, indent=2))
        return 0
    if args.stub_addin:
        return run_stub_addin()
    if args.smoke:
        return run_smoke(args.port, args.smoke_seconds)
    if args.selftest:
        if getattr(sys, "frozen", False):
            sys.stderr.write(
                "vb6mcp_bridge: --selftest needs the SOURCE tree (it runs tools/*.py and\n"
                "  tests/*.py out of the repository, which a frozen bundle does not carry).\n"
                "  Accept a frozen build with --smoke instead: it bundles its own stub np/1\n"
                "  add-in and serves the real /rpc + /mcp face for the external tools.\n")
            return 2
        return run_selftest(with_battery=args.with_battery)
    if args.stdio:
        return run_stdio(CFG.target)
    return serve(CFG.port, CFG.host, CFG.target)


# (the __main__ entry point lives at the very END of this file, after the test harness)


# =====================================================================================
# DEV / TEST HARNESS  --  reached ONLY by --selftest
#
# This is the stub add-in the card asks for: it speaks the SAME np/1 protocol the VB side
# will speak, so the Python service layer can be proven end to end WITHOUT a VB6 IDE and
# WITHOUT touching a single VB source file.  Nothing here is on the product path.
# =====================================================================================

def _blocking_read_frame(h, deadline):
    buf = FrameBuffer()
    while True:
        avail = wintypes.DWORD()
        peek_ok = kernel32.PeekNamedPipe(h, None, 0, None, ctypes.byref(avail), None)
        if not peek_ok:
            # A CLOSED peer must be noticed IMMEDIATELY, not at the idle deadline: a stub that
            # holds a dead handle keeps its single pipe instance "connected" and the next
            # client's WaitNamedPipe legitimately reports busy (measured: ctl reconnect got
            # 121 for the full connect budget before this check existed).
            code = ctypes.get_last_error()
            if code in (ERROR_BROKEN_PIPE, ERROR_NO_DATA, ERROR_PIPE_NOT_CONNECTED):
                return None
        if peek_ok and avail.value:
            n = avail.value
            chunk = ctypes.create_string_buffer(n)
            got = wintypes.DWORD()
            if not kernel32.ReadFile(h, ctypes.cast(chunk, ctypes.c_void_p), n,
                                     ctypes.byref(got), None):
                return None
            buf.feed(chunk.raw[:got.value])
        try:
            payload = buf.pop()
        except FrameError:
            return None
        if payload is not None:
            return payload
        if time.monotonic() > deadline:
            return None
        time.sleep(0.003)


def _blocking_write_frame(h, obj):
    data = encode_frame(obj)
    sent = 0
    while sent < len(data):
        chunk = data[sent:]
        cbuf = ctypes.create_string_buffer(chunk, len(chunk))
        got = wintypes.DWORD()
        if not kernel32.WriteFile(h, ctypes.cast(cbuf, ctypes.c_void_p), len(chunk),
                                  ctypes.byref(got), None):
            return False
        if not got.value:
            return False
        sent += got.value
    return True


WS_CHILD = 0x40000000
WS_VISIBLE = 0x10000000
WS_POPUP = 0x80000000
WS_EX_TOOLWINDOW = 0x00000080
WS_EX_NOACTIVATE = 0x08000000


def _make_oop_window_fixture():
    """A #32770 dialog + Button/Static children owned by THIS process.

    Section B measures an EXTERNAL pid.  This fixture makes that measurable deterministically
    WITHOUT starting VB6 (red line) and without touching any other live application: the pid
    being measured is just this test process.  The dialog is a 1x1 off-screen toolwindow with
    WS_EX_NOACTIVATE, so it never steals focus and never appears on screen.
    """
    try:
        user32.CreateWindowExW.restype = wintypes.HWND
        user32.CreateWindowExW.argtypes = [
            wintypes.DWORD, wintypes.LPCWSTR, wintypes.LPCWSTR, wintypes.DWORD,
            ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.c_int,
            wintypes.HWND, wintypes.HMENU, wintypes.HINSTANCE, ctypes.c_void_p]
    except Exception:
        return {"dialog": 0, "button": 0, "debug": 0, "static": 0}

    def _mk(parent, cls, text, style, x, y, w, h):
        return _hwnd_i(user32.CreateWindowExW(
            0, cls, text, style, x, y, w, h, parent or None, None, None, None))

    dlg = _mk(None, "#32770", "Microsoft Visual Basic",
              WS_POPUP | WS_VISIBLE, 3000, 3000, 320, 120)
    btn = _mk(dlg, "Button", "End", WS_CHILD | WS_VISIBLE, 4, 4, 80, 24)
    dbg = _mk(dlg, "Button", "Debug", WS_CHILD | WS_VISIBLE, 90, 4, 80, 24)
    st = _mk(dlg, "Static",
             "Run-time error '91': Object variable or With block variable not set",
             WS_CHILD | WS_VISIBLE, 4, 40, 300, 60)
    return {"dialog": dlg, "button": btn, "debug": dbg, "static": st}


class StubPipeAddIn(object):
    """A fake VB6 add-in: np/1 framing, hello/welcome/deny, JSON-RPC, progress frames."""

    #: the stub's idle-close window.  ADR 7.2's real server closes after 120 s idle; this is
    #: deliberately much shorter but still far longer than any gap inside one battery run, so
    #: the recorded counts are REPRODUCIBLE instead of racing a 1 s timeout.
    IDLE_CLOSE_S = 30.0
    #: the add-in's handshake deadline (writer-np-vb: 5000 ms; measured 5.5 s eviction on t16).
    #: A client that connects and then says nothing must be evicted, or it would hold the single
    #: npipe instance (nMaxInstances=1) for ever.
    HELLO_DEADLINE_S = 5.0

    def __init__(self, pipe_name, nonce, pid, project="StubProj", tools=None,
                 behaviour="normal", slow_ms=0, progress_ms=200, dll_sha="STUBDLL",
                 deny_reason="nonce_mismatch", deny_detail=""):
        self.deny_reason = deny_reason
        self.deny_detail = deny_detail
        self.pipe_name = pipe_name
        self.nonce = nonce
        self.pid = pid
        self.project = project
        self.tools = tools or []
        self.behaviour = behaviour
        self.slow_ms = slow_ms
        self.progress_ms = progress_ms
        self.dll_sha = dll_sha
        self.stop_evt = threading.Event()
        self.thread = None
        self.sessions = 0
        self.requests = 0
        self.denied = 0
        self.helloTimeouts = 0
        self.last_args = None
        self.last_error = None
        #: the EXACT bytes of the first frame this session received (the hello), so a test can
        #: assert on the wire - not on the code - that the nonce was byte-identical/compact.
        self.last_hello = None

    def start(self):
        self.thread = threading.Thread(target=self._run, name="stub-pipe", daemon=True)
        self.thread.start()
        # give the first CreateNamedPipe a moment so the very next WaitNamedPipe sees it
        for _ in range(200):
            if wait_named_pipe(self.pipe_name, 1)[0] != "absent":
                break
            time.sleep(0.01)
        return self

    def stop(self):
        self.stop_evt.set()
        try:
            wait_named_pipe(self.pipe_name, 50)
        except Exception:
            pass

    def _run(self):
        while not self.stop_evt.is_set():
            h = kernel32.CreateNamedPipeW(
                ctypes.c_wchar_p(self.pipe_name), PIPE_ACCESS_DUPLEX,
                PIPE_TYPE_BYTE | PIPE_READMODE_BYTE | PIPE_WAIT, 1, 65536, 65536, 0, None)
            if not h or h == INVALID_HANDLE_VALUE:
                time.sleep(0.05)
                continue
            try:
                if not kernel32.ConnectNamedPipe(h, None):
                    code = ctypes.get_last_error()
                    if code != ERROR_PIPE_CONNECTED:
                        continue
                self._serve(h)
            except Exception as exc:
                self.last_error = repr(exc)
            finally:
                kernel32.DisconnectNamedPipe(h)
                kernel32.CloseHandle(h)

    def _serve(self, h):
        first = _blocking_read_frame(h, time.monotonic() + self.HELLO_DEADLINE_S)
        if first is None:
            self.helloTimeouts += 1
            return
        self.last_hello = first
        try:
            hello = json.loads(first.decode("utf-8"))
        except Exception:
            hello = {}
        if self.behaviour == "deny_hello" or hello.get("vb6mcp") != "hello"                 or str(hello.get("nonce")) != str(self.nonce):
            self.denied += 1
            _blocking_write_frame(h, {"vb6mcp": "deny", "reason": self.deny_reason,
                                      "detail": self.deny_detail})
            # the real add-in writes the deny, waits for the client to read it (or a 750 ms
            # grace) and only then closes the handle.  Closing immediately would DISCARD the
            # buffered frame and turn every deny into a transport error (233).
            kernel32.FlushFileBuffers(h)
            time.sleep(0.75)
            return
        self.sessions += 1
        _blocking_write_frame(h, {
            "vb6mcp": "welcome", "proto": NP_PROTO,
            "instanceId": "%s-%s" % (self.pid, self.nonce), "pid": self.pid, "session": 1,
            "project": self.project, "dllSha256": self.dll_sha,
            "transport": "pipe", "tcpEnabled": True})
        while not self.stop_evt.is_set():
            payload = _blocking_read_frame(h, time.monotonic() + self.IDLE_CLOSE_S)
            if payload is None:
                return
            try:
                req = json.loads(payload.decode("utf-8"))
            except Exception:
                return
            self.requests += 1
            params = req.get("params") if isinstance(req.get("params"), dict) else {}
            if isinstance(params.get("arguments"), dict):
                self.last_args = params["arguments"]
            if not self._answer(h, req):
                return

    def _answer(self, h, req):
        rid = req.get("id")
        method = req.get("method")
        params = req.get("params") if isinstance(req.get("params"), dict) else {}
        tool = params.get("name") if isinstance(params.get("name"), str) else ""
        if self.behaviour == "die":
            return False
        if self.behaviour == "deny_frame_too_large":
            _blocking_write_frame(h, {"vb6mcp": "deny", "reason": "frame_too_large",
                                      "detail": self.deny_detail})
            kernel32.FlushFileBuffers(h)
            time.sleep(0.75)
            return False
        if self.behaviour == "silent":
            time.sleep(60)
            return False
        if self.behaviour == "stall":
            _blocking_write_frame(h, {"vb6mcp": "progress", "seq": 1, "elapsedMs": 1,
                                      "tool": tool or ""})
            time.sleep(60)
            return False
        result = self._result_for(method, tool, rid, params)
        if rid is None:
            return True
        if self.behaviour == "slow" and self.slow_ms:
            t0 = time.monotonic()
            n = 0
            while (time.monotonic() - t0) * 1000.0 < self.slow_ms and not self.stop_evt.is_set():
                n += 1
                _blocking_write_frame(h, {"vb6mcp": "progress", "seq": n,
                                          "elapsedMs": int((time.monotonic() - t0) * 1000),
                                          "tool": tool or ""})
                time.sleep(self.progress_ms / 1000.0)
        return _blocking_write_frame(h, result)

    def _result_for(self, method, tool, rid, params):
        if method == "initialize":
            return {"jsonrpc": "2.0", "id": rid,
                    "result": local_initialize({"params": params})}
        if method == "tools/list":
            return {"jsonrpc": "2.0", "id": rid, "result": {"tools": self.tools}}
        if method == "tools/call":
            if tool == "__bigframe":
                n = params.get("arguments", {}).get("n") or 0
                return {"jsonrpc": "2.0", "id": rid,
                        "result": {"content": [{"type": "text", "text": json.dumps(
                            {"ok": True, "stub": True, "tool": tool, "filler": "x" * int(n)})}],
                            "isError": False}}
            body = {"ok": True, "stub": True, "tool": tool, "project": self.project}
            if tool == "get_project_info":
                body = {"ok": True, "stub": True, "tool": tool, "projectType": "StandardExe",
                        "project": self.project, "ideState": "design"}
            return {"jsonrpc": "2.0", "id": rid,
                    "result": {"content": [{"type": "text", "text": json.dumps(body)}],
                               "isError": False}}
        return {"jsonrpc": "2.0", "id": rid, "result": {"ok": True, "stub": True,
                                                        "method": method}}


class StubCtlAddIn(object):
    """A fake add-in CONTROL lane (ctl/1) - STRICT server semantics, never a permissive pass.

    It implements the parts of bridge/CTL_WIRE_CONTRACT.md the client half must be proven
    against: hello/welcome with authBasis, sid + strictly increasing seq, the six ops, the two
    readonly-lane refusal details, the 3-attempts-per-60s lockout, the nested tool receipt
    (outer ok = transport verdict, receipt.ok = the tool's own verdict), the static limits block,
    and an abandon that must be preceded by a release (or confirmed).
    """

    IDLE_CLOSE_S = 20.0
    HELLO_DEADLINE_S = 5.0
    LOCKOUT_S = 60.0
    LOCKOUT_AFTER = 3
    LIMITS = {"servedFromSamePump": True, "statementInterrupted": False,
              "nonPumpingCallServiceable": False}
    WRITE_TOOLS = frozenset(("build", "make_project", "write_file", "replace_code",
                             "control_execution", "add_control", "add_component"))

    def __init__(self, pipe_name, nonce, pid, token="", gate_tool="immediate_eval",
                 gate_held_ms=4210, ops=None):
        self.pipe_name = pipe_name
        self.nonce = nonce
        self.pid = pid
        self.token = token
        self.ops = list(ops or CTL_OPS)
        self.gate = {"depth": 1, "mutBusy": True, "tool": gate_tool, "heldMs": gate_held_ms}
        self.gate_release_count = 0
        self.cancel = {"requested": False, "applied": False}
        self.business_pipe_closed = False
        self.stop_evt = threading.Event()
        self.thread = None
        self.conns = 0
        self.denied = 0
        self.served = 0
        self.helloTimeouts = 0
        self.lockouts = 0
        self.denials = []
        self.last_hello = None
        self.sids = []
        self.pings = 0
        self.states = 0
        self.release_calls = []
        self.cancel_calls = []
        self.abandon_calls = []
        self.tool_calls = []
        self.last_error = None
        self._auth_window = []
        self._lock_until = 0.0

    def start(self):
        self.thread = threading.Thread(target=self._run, name="stub-ctl", daemon=True)
        self.thread.start()
        for _ in range(200):
            if wait_named_pipe(self.pipe_name, 1)[0] != "absent":
                break
            time.sleep(0.01)
        return self

    def stop(self):
        self.stop_evt.set()
        try:
            wait_named_pipe(self.pipe_name, 50)
        except Exception:
            pass

    # -- plumbing ----------------------------------------------------------------
    def _send(self, h, obj):
        try:
            return _blocking_write_frame(h, obj)
        except Exception:
            return False

    def _run(self):
        while not self.stop_evt.is_set():
            h = kernel32.CreateNamedPipeW(
                ctypes.c_wchar_p(self.pipe_name), PIPE_ACCESS_DUPLEX,
                PIPE_TYPE_BYTE | PIPE_READMODE_BYTE | PIPE_WAIT, 2, 65536, 65536, 0, None)
            if not h or h == INVALID_HANDLE_VALUE:
                time.sleep(0.05)
                continue
            try:
                if not kernel32.ConnectNamedPipe(h, None):
                    if ctypes.get_last_error() != ERROR_PIPE_CONNECTED:
                        continue
                self._serve(h)
            except Exception as exc:
                self.last_error = repr(exc)
            finally:
                kernel32.DisconnectNamedPipe(h)
                kernel32.CloseHandle(h)

    def _deny_unauth(self, h, detail):
        """Count ONE unauthenticated attempt; the 3rd inside the window locks the lane for 60 s."""
        self.denied += 1
        now = time.monotonic()
        self._auth_window = [t for t in self._auth_window if now - t <= self.LOCKOUT_S]
        self._auth_window.append(now)
        self.denials.append({"detail": detail, "attempt": len(self._auth_window)})
        if len(self._auth_window) >= self.LOCKOUT_AFTER:
            self._lock_until = now + self.LOCKOUT_S
            self.lockouts += 1
            self._send(h, {"vb6mcp": "deny", "proto": CTL_PROTO, "reason": "ctl_lockout",
                           "retryAfterMs": int(self.LOCKOUT_S * 1000)})
        else:
            self._send(h, {"vb6mcp": "deny", "proto": CTL_PROTO, "reason": "ctl_unauthorized",
                           "detail": detail})
        try:
            kernel32.FlushFileBuffers(h)
        except Exception:
            pass
        time.sleep(0.05)

    def _serve(self, h):
        first = _blocking_read_frame(h, time.monotonic() + self.HELLO_DEADLINE_S)
        if first is None:
            self.helloTimeouts += 1
            return
        self.last_hello = first
        try:
            hello = json.loads(first.decode("utf-8"))
        except Exception:
            hello = {}
        self.conns += 1
        if time.monotonic() < self._lock_until:
            self._send(h, {"vb6mcp": "deny", "proto": CTL_PROTO, "reason": "ctl_lockout",
                           "retryAfterMs": int(self.LOCKOUT_S * 1000)})
            try:
                kernel32.FlushFileBuffers(h)
            except Exception:
                pass
            time.sleep(0.05)
            return
        if not isinstance(hello, dict) or str(hello.get("vb6mcp")) != "ctlhello":
            self._send(h, {"vb6mcp": "deny", "proto": CTL_PROTO, "reason": "bad_hello",
                           "detail": "first_frame_not_ctlhello"})
            try:
                kernel32.FlushFileBuffers(h)
            except Exception:
                pass
            time.sleep(0.05)
            return
        if str(hello.get("proto")) != CTL_PROTO:
            return self._deny_unauth(h, "bad_proto")
        if (str(hello.get("nonce")) != str(self.nonce)
                or str(hello.get("instanceId")) != "%s-%s" % (self.pid, self.nonce)):
            return self._deny_unauth(h, "nonce_mismatch")
        if self.token:
            auth = hello.get("auth") if isinstance(hello.get("auth"), dict) else {}
            supplied = str(auth.get("token") or "")
            same = hmac.compare_digest(hashlib.sha256(supplied.encode("utf-8")).digest(),
                                       hashlib.sha256(self.token.encode("utf-8")).digest())
            if not supplied or not same:
                return self._deny_unauth(h, "token_mismatch")
        sid = uuid.uuid4().hex[:16]
        self.sids.append(sid)
        basis = ["nonce", "acl"] + (["token"] if self.token else [])
        self._send(h, {"vb6mcp": "ctlwelcome", "proto": CTL_PROTO,
                       "instanceId": "%s-%s" % (self.pid, self.nonce), "pid": self.pid,
                       "sid": sid, "seqBase": 0, "ops": list(self.ops), "authBasis": basis,
                       "maxFrame": 65536, "ackedAt": int(time.time() * 1000),
                       "gate": dict(self.gate),
                       "ctl": {"conns": self.conns, "denied": self.denied, "served": self.served},
                       "ctlTools": list(CTL_FAST_TOOLS), "limits": dict(self.LIMITS)})
        # Lead ruling 1 (pinned in the contract): the FIRST frame must be EXACTLY seqBase and
        # every later frame exactly +1.  A first frame of seqBase+1 is a REJECTION - S29g asserts
        # that direction (the client's own first frame is seqBase, S29e covers the accept side).
        expected_seq = 0
        while not self.stop_evt.is_set():
            payload = _blocking_read_frame(h, time.monotonic() + self.IDLE_CLOSE_S)
            if payload is None:
                return
            try:
                frame = json.loads(payload.decode("utf-8"))
            except Exception:
                return
            if not isinstance(frame, dict) or str(frame.get("vb6mcp")) != "ctl":
                self._send(h, {"vb6mcp": "deny", "proto": CTL_PROTO, "reason": "bad_hello",
                               "detail": "not_a_ctl_frame"})
                return
            if str(frame.get("sid") or "") != sid:
                self._send(h, {"vb6mcp": "ctlresult", "sid": frame.get("sid"),
                               "seq": frame.get("seq"), "ok": False,
                               "reason": "ctl_unauthorized", "detail": "bad_session"})
                return
            try:
                seq = int(frame.get("seq"))
            except (TypeError, ValueError):
                seq = -1
            if seq != expected_seq:
                self._send(h, {"vb6mcp": "ctlresult", "sid": sid, "seq": seq if seq >= 0 else 0,
                               "ok": False, "reason": "ctl_unauthorized", "detail": "bad_seq"})
                return
            expected_seq = seq + 1
            self.served += 1
            self._dispatch(h, sid, seq, str(frame.get("op") or ""),
                           frame.get("args") if isinstance(frame.get("args"), dict) else {})

    # -- ops ---------------------------------------------------------------------
    def _result(self, sid, seq):
        return {"vb6mcp": "ctlresult", "sid": sid, "seq": seq, "limits": dict(self.LIMITS)}

    def _dispatch(self, h, sid, seq, op, args):
        d = self._result(sid, seq)
        if op == "ping":
            self.pings += 1
            d.update({"ok": True, "pid": self.pid, "uptimeMs": 12345,
                      "ackedAt": int(time.time() * 1000), "gate": dict(self.gate),
                      "businessPipe": {"state": "busy" if self.gate["mutBusy"] else "idle",
                                       "heldMs": self.gate["heldMs"]},
                      "pump": {"calls": self.served, "gateSkips": 0}})
            return self._send(h, d)
        if op == "state":
            self.states += 1
            d.update({"ok": True, "pid": self.pid, "mode": "run", "project": "StubCtl",
                      "gate": dict(self.gate),
                      "cancel": dict(self.cancel),
                      "release": {"attempted": self.gate_release_count > 0,
                                  "released": self.gate_release_count > 0,
                                  "heldMs": self.gate["heldMs"]},
                      "ctl": {"conns": self.conns, "denied": self.denied, "served": self.served}})
            return self._send(h, d)
        if op == "release_gate":
            self.release_calls.append(dict(args))
            try:
                min_held = int(args.get("minHeldMs") or 0)
            except (TypeError, ValueError):
                min_held = 0
            if not self.gate["mutBusy"]:
                d.update({"ok": False, "reason": "nothing_in_flight"})
                return self._send(h, d)
            if self.gate["heldMs"] < min_held:
                # declined, not refused: the op ran, the gate stays held (no invented reason code)
                d.update({"ok": True, "released": False, "releasedTool": None,
                          "heldMs": self.gate["heldMs"], "depthAfter": self.gate["depth"],
                          "gateReleaseCount": self.gate_release_count,
                          "statementInterrupted": False,
                          "note": "heldMs %d < minHeldMs %d" % (self.gate["heldMs"], min_held)})
                return self._send(h, d)
            tool = self.gate["tool"]
            held = self.gate["heldMs"]
            self.gate_release_count += 1
            self.gate = {"depth": 0, "mutBusy": False, "tool": None, "heldMs": 0}
            d.update({"ok": True, "released": True, "releasedTool": tool, "heldMs": held,
                      "depthAfter": 0, "gateReleaseCount": self.gate_release_count,
                      "statementInterrupted": False,
                      "note": "gate accounting released; the running statement still finishes"})
            return self._send(h, d)
        if op == "cancel":
            self.cancel_calls.append(dict(args))
            tgt = str(args.get("target") or "*")
            in_flight = self.gate["tool"] if self.gate["mutBusy"] else None
            if not in_flight:
                d.update({"ok": False, "reason": "nothing_in_flight"})
                return self._send(h, d)
            if tgt not in ("*", in_flight):
                d.update({"ok": False, "reason": "nothing_in_flight", "inFlightTool": in_flight})
                return self._send(h, d)
            self.cancel = {"requested": True, "applied": False}
            d.update({"ok": True, "requested": True, "honoredAtNextPump": True})
            return self._send(h, d)
        if op == "abandon":
            self.abandon_calls.append(dict(args))
            confirm = args.get("confirm") is True
            try:
                min_held = int(args.get("minHeldMs") or 0)
            except (TypeError, ValueError):
                min_held = 0
            if self.gate_release_count <= 0:
                if confirm and self.gate["mutBusy"] and self.gate["heldMs"] >= min_held:
                    self.gate_release_count += 1
                    self.gate = {"depth": 0, "mutBusy": False, "tool": None, "heldMs": 0}
                else:
                    d.update({"ok": False, "reason": "gate_not_released",
                              "howToProceed": "call release_gate first"})
                    return self._send(h, d)
            abandoned_tool = "immediate_eval"
            self.business_pipe_closed = True
            self.gate = {"depth": 0, "mutBusy": False, "tool": None, "heldMs": 0}
            d.update({"ok": True, "abandoned": True, "responseLost": True,
                      "abandonedTool": abandoned_tool, "heldMs": 4210,
                      "businessPipe": {"closed": True, "state": "idle"},
                      "note": "the abandoned call's response will be discarded"})
            return self._send(h, d)
        if op == "tool":
            name = str(args.get("name") or "")
            self.tool_calls.append(name)
            allowed = list(CTL_FAST_TOOLS)
            if name in self.WRITE_TOOLS:
                d.update({"ok": False, "reason": "readonly_lane_refused",
                          "detail": "not_read_only", "allowed": allowed,
                          "failureKind": "toolError"})
                return self._send(h, d)
            if name not in CTL_FAST_TOOLS:
                d.update({"ok": False, "reason": "readonly_lane_refused",
                          "detail": "slow_tool_not_on_ctl_lane", "allowed": allowed,
                          "failureKind": "toolError"})
                return self._send(h, d)
            # 4.6: the tool's receipt is NESTED; the outer ok is the frame verdict only.
            receipt = {"ok": True, "tool": name, "stub": True, "ctlLane": True,
                       "arguments": args.get("arguments") or {}}
            d.update({"ok": True, "source": "ctl_lane", "receipt": receipt})
            return self._send(h, d)
        d.update({"ok": False, "reason": "ctl_unknown_op", "supported": list(self.ops)})
        return self._send(h, d)


class StubHttpAddIn(object):
    """The transition/rollback track target: an add-in that still listens on HTTP /rpc."""

    def __init__(self, tools, project="StubHttp"):
        self.tools = tools
        self.project = project
        self.httpd = None
        self.port = None
        self.requests = 0

    def start(self):
        outer = self

        class Handler(http.server.BaseHTTPRequestHandler):
            protocol_version = "HTTP/1.1"

            def log_message(self, *a):
                pass

            def do_POST(self):
                n = int(self.headers.get("Content-Length") or 0)
                raw = self.rfile.read(n)
                try:
                    req = json.loads(raw.decode("utf-8"))
                except Exception:
                    req = {}
                outer.requests += 1
                rid = req.get("id")
                method = req.get("method")
                name = ""
                p = req.get("params")
                if isinstance(p, dict) and isinstance(p.get("name"), str):
                    name = p["name"]
                if method == "initialize":
                    result = local_initialize(req)
                elif method == "tools/list":
                    result = {"tools": outer.tools}
                elif method == "tools/call":
                    result = {"content": [{"type": "text", "text": json.dumps(
                        {"ok": True, "stub": "http", "tool": name, "project": outer.project})}],
                        "isError": False}
                else:
                    result = {"ok": True, "stub": "http", "method": method}
                body = json.dumps({"jsonrpc": "2.0", "id": rid, "result": result}).encode("utf-8")
                self.send_response(200)
                self.send_header("Content-Type", "application/json")
                self.send_header("Content-Length", str(len(body)))
                self.end_headers()
                self.wfile.write(body)

        self.httpd = http.server.ThreadingHTTPServer(("127.0.0.1", 0), Handler)
        self.httpd.daemon_threads = True
        self.port = self.httpd.server_address[1]
        threading.Thread(target=self.httpd.serve_forever, kwargs={"poll_interval": 0.05},
                         daemon=True).start()
        return self

    def stop(self):
        try:
            self.httpd.shutdown()
            self.httpd.server_close()
        except Exception:
            pass


def _reset_registry(regdir):
    try:
        for name in os.listdir(regdir):
            if name.lower().endswith(".json"):
                os.remove(os.path.join(regdir, name))
    except OSError:
        pass
    discover(force=True)


def _write_instance(regdir, pid, port=None, project=None, pipe=None, nonce=None,
                    instance_id=None, transport=None, project_path=None, **extra):
    if not os.path.isdir(regdir):
        os.makedirs(regdir)
    payload = {"pid": pid, "port": port, "project": project, "projectPath": project_path,
               "startedAt": "2026-10-15 00:00:00", "version": "stub-selftest",
               "dllSha256": "STUBDLL", "sessionId": 1}
    if transport:
        payload["transport"] = transport
    if pipe:
        payload["pipe"] = pipe
    if nonce:
        payload["nonce"] = nonce
    if instance_id:
        payload["instanceId"] = instance_id
    payload.update(extra)
    name = (nonce or uuid.uuid4().hex[:12]) + ".json"
    path = os.path.join(regdir, name)
    with open(path, "w", encoding="utf-8") as fh:
        json.dump(payload, fh)
    return path


def _jreq(method, params=None, rid=1):
    d = {"jsonrpc": "2.0", "id": rid, "method": method}
    if params is not None:
        d["params"] = params
    return json.dumps(d, ensure_ascii=False)


def _resp_body(text):
    try:
        j = json.loads(text)
    except Exception:
        return {"_unparsable": text}
    res = j.get("result") if isinstance(j, dict) else None
    if isinstance(res, dict):
        for c in (res.get("content") or []):
            if c.get("type") == "text":
                try:
                    return json.loads(c["text"])
                except Exception:
                    return {"_raw": c["text"]}
        return res
    if isinstance(j, dict) and j.get("error"):
        return {"_error": j["error"]}
    return j if isinstance(j, dict) else {"_type": type(j).__name__}


def _free_port():
    s = socket.socket()
    s.bind(("127.0.0.1", 0))
    port = s.getsockname()[1]
    s.close()
    return port


def _child_env(extra=None):
    env = dict(os.environ)
    env["PYTHONIOENCODING"] = "utf-8"
    if extra:
        env.update(extra)
    return env


def _run_tool(args, timeout=180):
    return subprocess.run([sys.executable] + args, cwd=REPO_ROOT, capture_output=True,
                          text=True, encoding="utf-8", errors="replace",
                          env=_child_env(), timeout=timeout)


def _tools_sha256(tools):
    """Stable fingerprint of a tools array, so two runs (source vs frozen) can be diffed."""
    import hashlib
    blob = json.dumps(tools, ensure_ascii=False, sort_keys=True).encode("utf-8")
    return hashlib.sha256(blob).hexdigest().upper()


def _free_port_selfowned():
    """A free port inside the self-owned range (tools/mcp_strict_audit.py refuses < 8767)."""
    for _ in range(50):
        p = _free_port()
        if p >= DEFAULT_PORT:
            return p
    return DEFAULT_PORT + 1


def run_smoke(port, seconds):
    """Runtime acceptance for a FROZEN build (and a useful source-form check).

    A frozen EXE must NOT be accepted by --selftest: that harness runs tools/*.py and
    tests/*.py out of the repository, which a onefile bundle does not carry.  This mode
    carries its own stub np/1 add-in and serves the real HTTP face, so the SAME binary can be
    exercised on any box with nothing but the EXE.  Every line it prints is one JSON object,
    so two runs can be compared mechanically (the point of the acceptance).

    It never writes to the repository: the registry fixture lives in a temp directory.
    """
    if not is_loopback_host(CFG.host):
        print(json.dumps({"smoke": "fatal", "reason": "non_loopback_host_refused",
                          "host": CFG.host,
                          "message": "M-0: non-loopback binding is disabled unconditionally"}))
        return 6
    frozen = bool(getattr(sys, "frozen", False))
    exe = sys.executable
    print(json.dumps({"smoke": "start", "product": PRODUCT, "version": VERSION,
                      "frozen": frozen, "executable": exe,
                      "dataDirs": DATA_DIRS, "snapshotPath": SNAP_PATH,
                      "hubToolsPath": HUB_TOOLS_PATH,
                      "snapshotExists": os.path.isfile(SNAP_PATH)}, ensure_ascii=False))
    snap = snapshot()
    tools = snap.get("tools") or []
    print(json.dumps({"smoke": "assets", "snapshotTools": len(tools),
                      "snapshotFingerprint": _tools_sha256(tools),
                      "contractVersion": contract_version()}, ensure_ascii=False))
    if not tools:
        print(json.dumps({"smoke": "fatal",
                          "reason": "snapshot_unreadable", "path": SNAP_PATH}))
        return 5

    tmpdir = tempfile.mkdtemp(prefix="vb6mcp_bridge_smoke_")
    CFG.instances_dir = os.path.join(tmpdir, "instances")
    nonce = uuid.uuid4().hex[:16]
    pipe = "\\\\.\\pipe\\vb6mcp-smoke-%d-%s" % (os.getpid(), nonce)
    stub = StubPipeAddIn(pipe, nonce, os.getpid(), project="SmokeStub", tools=tools).start()
    _write_instance(CFG.instances_dir, os.getpid(), project="SmokeStub", pipe=pipe,
                    nonce=nonce, instance_id="%d-%s" % (os.getpid(), nonce), transport="pipe")
    discover(force=True)

    text, iserr, meta = handle_jsonrpc(_jreq("tools/list", {}))
    got = []
    try:
        got = (json.loads(text or "{}").get("result") or {}).get("tools") or []
    except Exception:
        got = []
    print(json.dumps({"smoke": "pipe_roundtrip", "transport": meta.get("transport"),
                      "transportReason": meta.get("transportReason"),
                      "liveTools": len(got), "liveFingerprint": _tools_sha256(got),
                      "isError": bool(iserr), "stubSessions": stub.sessions,
                      "faceMatchesSnapshot": _tools_sha256(got) == _tools_sha256(tools)},
                     ensure_ascii=False))

    chosen = port or _free_port_selfowned()
    try:
        httpd = BridgeServer((CFG.host, chosen), BridgeHandler)
    except OSError as exc:
        print(json.dumps({"smoke": "fatal", "reason": "bind_failed", "port": chosen,
                          "error": str(exc)}))
        stub.stop()
        shutil.rmtree(tmpdir, ignore_errors=True)
        return 4
    CFG.port = chosen
    threading.Thread(target=httpd.serve_forever, kwargs={"poll_interval": 0.1},
                     daemon=True).start()
    print(json.dumps({"smoke": "ready", "port": chosen,
                      "url_rpc": "http://%s:%d/rpc" % (CFG.host, chosen),
                      "url_mcp": "http://%s:%d/mcp" % (CFG.host, chosen),
                      "health": "http://%s:%d/health" % (CFG.host, chosen),
                      "stubPipe": pipe, "seconds": seconds}, ensure_ascii=False))
    sys.stdout.flush()
    deadline = time.monotonic() + max(1, seconds)
    try:
        while time.monotonic() < deadline:
            time.sleep(0.2)
    except KeyboardInterrupt:
        pass
    httpd.shutdown()
    httpd.server_close()
    POOL.close_all()
    stub.stop()
    shutil.rmtree(tmpdir, ignore_errors=True)
    print(json.dumps({"smoke": "stopped"}))
    return 0


def run_stub_addin():
    """DEV ONLY: run a stub np/1 pipe server as its OWN OS process.

    Used by --selftest (scenario S15) so the DEPLOYABLE entry point
    ("python vb6mcp_bridge.py --port N", the real CLI, its own process) can be exercised
    against a pipe whose server lives in a different process, exactly as in production.
    """
    name = os.environ.get("VB6MCP_STUB_PIPE")
    nonce = os.environ.get("VB6MCP_STUB_NONCE")
    if not name or not nonce:
        sys.stderr.write("--stub-addin needs VB6MCP_STUB_PIPE and VB6MCP_STUB_NONCE\n")
        return 2
    project = os.environ.get("VB6MCP_STUB_PROJECT", "StubChild")
    tools_path = os.environ.get("VB6MCP_STUB_TOOLS") or os.path.join(REPO_ROOT, "docs",
                                                                    "tools_schema.json")
    tools = read_json(tools_path, {}).get("tools") or []
    regdir = os.environ.get("VB6MCP_STUB_REGISTRY_DIR")
    if regdir:
        _write_instance(regdir, os.getpid(), project=project, pipe=name, nonce=nonce,
                        instance_id="%d-%s" % (os.getpid(), nonce), transport="pipe")
    stub = StubPipeAddIn(name, nonce, os.getpid(), project=project, tools=tools,
                         behaviour=os.environ.get("VB6MCP_STUB_BEHAVIOUR", "normal")).start()
    sys.stderr.write("stub-addin pid=%d pipe=%s tools=%d\n" % (os.getpid(), name, len(tools)))
    sys.stderr.flush()
    try:
        while True:
            time.sleep(0.5)
    except KeyboardInterrupt:
        pass
    finally:
        stub.stop()
    return 0


def run_selftest(with_battery=False):
    results = []

    def check(name, ok, detail=""):
        results.append((name, bool(ok), detail))
        print("%s %-64s %s" % ("PASS" if ok else "FAIL", name, detail))

    print("=" * 100)
    print("vb6mcp_bridge selftest - stub np/1 pipe add-in, NO VB6 IDE, NO VB source touched")
    print("=" * 100)
    print("python=%s  port-policy=bridge owns one port  netstat=never" % sys.version.split()[0])

    tmpdir = tempfile.mkdtemp(prefix="vb6mcp_bridge_selftest_")
    regdir = os.path.join(tmpdir, "instances")
    saved = (CFG.instances_dir, CFG.transport, CFG.pipe_idle_ms, CFG.forward_timeout_ms,
             CFG.pipe_fallback, CFG.pipe_replay, CFG.pipe_connect_ms, CFG.target,
             os.environ.get("VB6MCP_INSTANCES_SCAN"))
    # Hermetic: the multi-root registry scan is ON in production (a SYSTEM service must still
    # find the interactive user's registrations), but it would let FOREIGN entries into these
    # assertions - so the harness pins the single temp root.
    os.environ["VB6MCP_INSTANCES_SCAN"] = "0"
    CFG.instances_dir = regdir
    CFG.transport = ""
    CFG.target = ""
    CFG.pipe_fallback = True
    CFG.pipe_replay = False
    CFG.pipe_connect_ms = 300
    CFG.pipe_idle_ms = 900
    CFG.forward_timeout_ms = 120000
    stubs = []
    http_stubs = []
    httpd = None
    ok = False
    try:
        frozen = read_json(os.path.join(REPO_ROOT, "docs", "tools_schema.json"), {})
        tools = frozen.get("tools") or []
        check("P0 frozen add-in face loaded", len(tools) == 110, "docs/tools_schema.json tools=%d"
              % len(tools))

        mine = os.getpid()
        _reset_registry(regdir)

        # ---- S1: pipe track round-trip -------------------------------------------------
        n1 = uuid.uuid4().hex[:16]
        p1 = "\\\\.\\pipe\\vb6mcp-1-%d-selftest-%s" % (mine, n1)
        s1 = StubPipeAddIn(p1, n1, mine, project="StubOne", tools=tools).start()
        stubs.append(s1)
        _write_instance(regdir, mine, project="StubOne", pipe=p1, nonce=n1,
                        instance_id="%d-%s" % (mine, n1), transport="pipe")
        discover(force=True)
        text, iserr, meta = handle_jsonrpc(_jreq("tools/list", {}))
        tl = json.loads(text or "{}")
        check("S1 pipe track is chosen (auto_pipe_ok)",
              meta.get("transport") == "pipe" and meta.get("transportReason") == "auto_pipe_ok",
              "transport=%s reason=%s" % (meta.get("transport"), meta.get("transportReason")))
        check("S1 tools/list result keys == {tools} (strict client)",
              set((tl.get("result") or {}).keys()) == {"tools"},
              "keys=%s" % sorted((tl.get("result") or {}).keys()))
        check("S1 live 110-tool face over the pipe",
              len((tl.get("result") or {}).get("tools") or []) == 110,
              "count=%d" % len((tl.get("result") or {}).get("tools") or []))
        check("S1 stub observed exactly 1 pipe session", s1.sessions == 1,
              "sessions=%d requests=%d" % (s1.sessions, s1.requests))
        src = open(__file__, "r", encoding="utf-8").read()
        cut = src.find("# DEV / TEST HARNESS")
        product_src = src[:cut] if cut > 0 else src
        needles = ["net" + "stat.exe", "-a" + "no", "8799", "subprocess."]
        check("S1 discovery is registry + WaitNamedPipe/CreateFile only (no netstat, no scan)",
              not any(n in product_src for n in needles),
              "product region %d bytes, forbidden tokens absent" % len(product_src))

        # ---- S2: the HTTP face, driven by the EXISTING tools, unchanged ---------------
        port = _free_port()
        httpd = BridgeServer(("127.0.0.1", port), BridgeHandler)
        threading.Thread(target=httpd.serve_forever, kwargs={"poll_interval": 0.05},
                         daemon=True).start()
        CFG.port = port
        url_rpc = "http://127.0.0.1:%d/rpc" % port
        url_mcp = "http://127.0.0.1:%d/mcp" % port
        print("--- S2: middle service listening on http://127.0.0.1:%d ---" % port)

        r = _run_tool(["tools/mcp_strict_audit.py", "--url", url_mcp])
        check("S2 tools/mcp_strict_audit.py (UNCHANGED) passes every strict check",
              r.returncode == 0 and "strict_failures=0" in r.stdout,
              "rc=%d" % r.returncode)
        print((r.stdout or "").rstrip())

        dumped_path = os.path.join(tmpdir, "dumped.json")
        r2 = _run_tool(["tools/dump_tools_schema.py", "--url", url_rpc, "--out", dumped_path])
        dumped = read_json(dumped_path, {})
        check("S2 tools/dump_tools_schema.py (UNCHANGED) rc=0", r2.returncode == 0,
              "rc=%d %s" % (r2.returncode, (r2.stdout or "").replace("\n", " | ")[:120]))
        check("S2 the SERVICE face is byte-identical to the frozen 110-tool contract",
              dumped.get("tools") == tools,
              "tools=%d identical=%s" % (len(dumped.get("tools") or []),
                                         dumped.get("tools") == tools))

        r3 = _run_tool(["tests/smoke_test.py", "--port", str(port)])
        check("S2 tests/smoke_test.py (UNCHANGED) rc=0", r3.returncode == 0,
              "rc=%d" % r3.returncode)
        if r3.returncode != 0:
            print((r3.stderr or "").rstrip())
        else:
            print((r3.stdout or "").rstrip().splitlines()[0])

        if with_battery:
            r4 = _run_tool(["tests/battery_test.py", str(port)], timeout=300)
            exc = (r4.stdout or "").count("[EXC]")
            tgt = (r4.stdout or "").count("[??]")
            err = (r4.stdout or "").count("[ERR]")
            check("S2 tests/battery_test.py (UNCHANGED) runs the WHOLE battery with 0 transport exceptions",
                  exc == 0, "rc=%d [EXC]=%d | informational, NOT asserted (they vary run "
                            "to run because a stub is not an IDE): [??]=%d [ERR]=%d" % (
                      r4.returncode, exc, tgt, err))
            for line in (r4.stdout or "").splitlines()[-4:]:
                print("    " + line)
            for line in [l for l in (r4.stdout or "").splitlines() if "[ERR]" in l][:12]:
                print("    ERR " + line.strip()[:150])

        # ---- S13: no IDE at all -> the COMPLETE face still comes out of the snapshot ---
        _reset_registry(regdir)
        text, iserr, meta = handle_jsonrpc(_jreq("tools/list", {}))
        tl = json.loads(text or "{}")
        check("S13 no IDE: complete 110-tool face from bridge/hub_schema.generated.json",
              (meta.get("faceTools") == "snapshot"
               and len((tl.get("result") or {}).get("tools") or []) == 110),
              "faceTools=%s count=%d degraded=%s" % (
                  meta.get("faceTools"), len((tl.get("result") or {}).get("tools") or []),
                  meta.get("degraded")))
        text, iserr, meta = handle_jsonrpc(_jreq("initialize", {
            "protocolVersion": "2025-06-18", "capabilities": {}, "clientInfo": {}}))
        init = json.loads(text or "{}").get("result") or {}
        check("S13 no IDE: initialize still echoes the client protocolVersion",
              init.get("protocolVersion") == "2025-06-18",
              "protocolVersion=%s" % init.get("protocolVersion"))
        text, iserr, meta = handle_jsonrpc(_jreq("tools/call", {"name": "mcp_info", "arguments": {}}))
        body = _resp_body(text)
        check("S13 no IDE: tools/call fails STRUCTURALLY (never a fake success)",
              iserr is True and body.get("reason") == "no_ide_running"
              and body.get("failureKind") == "noIde",
              "reason=%s failureKind=%s" % (body.get("reason"), body.get("failureKind")))

        # ---- S4: registry identity mismatch -> never use that pipe, fall back visibly ---
        _reset_registry(regdir)
        hs = StubHttpAddIn(tools, project="HttpOne").start()
        http_stubs.append(hs)
        nb = uuid.uuid4().hex[:16]
        pb = "\\\\.\\pipe\\vb6mcp-1-%d-bad-%s" % (mine, uuid.uuid4().hex[:16])
        sbad = StubPipeAddIn(pb, nb, mine, project="BadEntry", tools=tools,
                             behaviour="die").start()
        stubs.append(sbad)
        _write_instance(regdir, mine, port=hs.port, project="BadEntry", pipe=pb, nonce=nb,
                        instance_id="WRONG-ID", transport="pipe")
        discover(force=True)
        text, iserr, meta = handle_jsonrpc(_jreq("tools/call",
                                                 {"name": "mcp_info", "arguments": {}}))
        d = discover()
        check("S4 pipe identity mismatch: pipe is REFUSED, tcp is used, flag is visible",
              meta.get("transport") == "tcp" and (meta.get("flags") or {}).get("pipeIdentityMismatch") is True,
              "transport=%s flags=%s" % (meta.get("transport"), meta.get("flags")))
        check("S4 mismatch is reported through staleRegistry",
              any(x.get("reason") == "pipe_identity_mismatch" for x in d["stale"]),
              "staleReasons=%s" % [x.get("reason") for x in d["stale"]])
        check("S4 the mismatched pipe server was NEVER contacted", sbad.sessions == 0,
              "sessions=%d denied=%d" % (sbad.sessions, sbad.denied))
        hs.stop()
        http_stubs.remove(hs)

        # ---- S5: ERROR_PIPE_BUSY -> pipeBusy + visible fallback -----------------------
        _reset_registry(regdir)
        hs = StubHttpAddIn(tools, project="HttpBusy").start()
        http_stubs.append(hs)
        n5 = uuid.uuid4().hex[:16]
        p5 = "\\\\.\\pipe\\vb6mcp-1-%d-busy-%s" % (mine, n5)
        s5 = StubPipeAddIn(p5, n5, mine, project="BusyOne", tools=tools).start()
        stubs.append(s5)
        _write_instance(regdir, mine, port=hs.port, project="BusyOne", pipe=p5, nonce=n5,
                        instance_id="%d-%s" % (mine, n5), transport="pipe")
        discover(force=True)
        occupier = PipeClient(p5).connect(600)
        occupier.send_frame({"vb6mcp": "hello", "proto": NP_PROTO, "nonce": n5,
                             "client": {"pid": mine, "session": 1, "user": "selftest"}},
                            time.monotonic() + 1.0)
        welcome, _err = occupier.recv_frame(time.monotonic() + 2.0)
        check("S5 the single pipe instance is occupied (nMaxInstances=1)",
              isinstance(welcome, dict) and welcome.get("vb6mcp") == "welcome",
              "occupier welcome=%s" % (welcome or {}).get("vb6mcp"))
        _drop_pipe_probe(discover()["live"][0])
        text, iserr, meta = handle_jsonrpc(_jreq("tools/call",
                                                 {"name": "mcp_info", "arguments": {}}))
        check("S5 the second client gets pipeBusy + transportFallback (auto only)",
              (meta.get("flags") or {}).get("pipeBusy") is True
              and (meta.get("flags") or {}).get("transportFallback") is True
              and meta.get("transportReason") == "auto_pipe_busy",
              "flags=%s reason=%s" % (meta.get("flags"), meta.get("transportReason")))
        occupier.close()
        hs.stop()
        http_stubs.remove(hs)
        # ---- S5b: pipe busy WITHOUT a tcp leg -> structured peer_busy envelope -------
        _reset_registry(regdir)
        n5b = uuid.uuid4().hex[:16]
        p5b = "\\\\.\\pipe\\vb6mcp-1-%d-busy2-%s" % (mine, n5b)
        s5b = StubPipeAddIn(p5b, n5b, mine, project="BusyTwo", tools=tools).start()
        stubs.append(s5b)
        _write_instance(regdir, mine, port=None, project="BusyTwo", pipe=p5b, nonce=n5b,
                        instance_id="%d-%s" % (mine, n5b), transport="pipe")
        discover(force=True)
        occ2 = PipeClient(p5b).connect(600)
        occ2.send_frame({"vb6mcp": "hello", "proto": NP_PROTO, "nonce": n5b,
                         "client": {"pid": mine, "session": 1, "user": "selftest"}},
                        time.monotonic() + 1.0)
        occ2.recv_frame(time.monotonic() + 2.0)
        _drop_pipe_probe(discover()["live"][0])
        text5b, iserr5b, meta5b = handle_jsonrpc(_jreq("tools/call",
                                                       {"name": "mcp_info", "arguments": {}}))
        body5b = _resp_body(text5b)
        check("S5b peer_busy envelope (failureKind/peerState/suggestedAction/retryAfterMs/heldBy)",
              iserr5b is True and body5b.get("failureKind") == "peer_busy"
              and body5b.get("peerState") == "peer_busy"
              and body5b.get("suggestedAction") == "retry_once"
              and body5b.get("retryAfterMs") == CFG.pipe_busy_retry_ms
              and isinstance(body5b.get("heldBy"), dict)
              and body5b.get("maxMsSource") == "env_ceiling",
              "reason=%s kind=%s action=%s retry=%s" % (body5b.get("reason"),
              body5b.get("failureKind"), body5b.get("suggestedAction"),
              body5b.get("retryAfterMs")))
        occ2.close()
        s5b.stop()

        # ---- S6: progress then silence -> pipe_stalled, NOT a 120 s wait ---------------
        _reset_registry(regdir)
        n6 = uuid.uuid4().hex[:16]
        p6 = "\\\\.\\pipe\\vb6mcp-1-%d-stall-%s" % (mine, n6)
        s6 = StubPipeAddIn(p6, n6, mine, project="StallOne", tools=tools,
                           behaviour="stall").start()
        stubs.append(s6)
        _write_instance(regdir, mine, project="StallOne", pipe=p6, nonce=n6,
                        instance_id="%d-%s" % (mine, n6), transport="pipe")
        discover(force=True)
        t0 = time.monotonic()
        text, iserr, meta = handle_jsonrpc(_jreq("tools/call",
                                                 {"name": "build", "arguments": {}}))
        body = _resp_body(text)
        elapsed = (time.monotonic() - t0) * 1000.0
        check("S6 progress then silence -> pipe_stalled inside PIPE_IDLE_MS (fast escape)",
              body.get("reason") == "pipe_stalled" and elapsed < 4000,
              "reason=%s elapsedMs=%.0f budgetMs=%s" % (body.get("reason"), elapsed,
                                                        body.get("budgetMs")))
        check("S6 pipeStalled flag rides on the receipt",
              body.get("pipeStalled") is True, "pipeStalled=%s" % body.get("pipeStalled"))

        # ---- S7: server dies mid-call, request already written -> no silent replay -----
        _reset_registry(regdir)
        n7 = uuid.uuid4().hex[:16]
        p7 = "\\\\.\\pipe\\vb6mcp-1-%d-die-%s" % (mine, n7)
        s7 = StubPipeAddIn(p7, n7, mine, project="DieOne", tools=tools,
                           behaviour="die").start()
        stubs.append(s7)
        _write_instance(regdir, mine, project="DieOne", pipe=p7, nonce=n7,
                        instance_id="%d-%s" % (mine, n7), transport="pipe")
        discover(force=True)
        text, iserr, meta = handle_jsonrpc(_jreq("tools/call",
                                                 {"name": "write_file", "arguments": {}}))
        body = _resp_body(text)
        check("S7 a fully-written request is NEVER silently replayed on another track",
              iserr is True and body.get("noReplay") is True and body.get("failureKind") == "peer_unresponsive",
              "reason=%s noReplay=%s" % (body.get("reason"), body.get("noReplay")))

        # ---- S8: no answer at all -> forward_timeout at the budget, not earlier --------
        _reset_registry(regdir)
        n8 = uuid.uuid4().hex[:16]
        p8 = "\\\\.\\pipe\\vb6mcp-1-%d-silent-%s" % (mine, n8)
        s8 = StubPipeAddIn(p8, n8, mine, project="SilentOne", tools=tools,
                           behaviour="silent").start()
        stubs.append(s8)
        _write_instance(regdir, mine, project="SilentOne", pipe=p8, nonce=n8,
                        instance_id="%d-%s" % (mine, n8), transport="pipe")
        discover(force=True)
        CFG.max_call_ms = 1200
        CFG.max_call_hard_ms = 1200
        t0 = time.monotonic()
        text, iserr, meta = handle_jsonrpc(_jreq("tools/call", {"name": "mcp_info", "arguments": {}}))
        body = _resp_body(text)
        elapsed = (time.monotonic() - t0) * 1000.0
        check("S8 no frame at all -> forward_timeout at FORWARD_TIMEOUT_MS (never tighter than tcp)",
              body.get("reason") == "timeout" and body.get("reasonLegacy") == "forward_timeout"
              and 1100 <= elapsed < 4000 and body.get("maxMs") == 1200
              and body.get("maxMsSource") == "env_ceiling"
              and body.get("failureKind") == "peer_unresponsive"
              and body.get("suggestedAction") in ("release_gate", "kill_and_restart")
              and isinstance(body.get("heldBy"), dict),
              "reason=%s legacy=%s elapsed=%.0f maxMs=%s src=%s kind=%s act=%s" %
              (body.get("reason"), body.get("reasonLegacy"), elapsed, body.get("maxMs"),
               body.get("maxMsSource"), body.get("failureKind"), body.get("suggestedAction")))
        CFG.max_call_ms = 120000
        CFG.max_call_hard_ms = 600000
        # ---- S28: two-layer budget mirror (pure function) ----------------------------
        check("S28 build budget >= 180000+15000 and <= hard cap",
              tool_budget("build", {})[0] >= 195000 and tool_budget("build", {})[1] == "param_slack",
              "build=%s" % (tool_budget("build", {}),))
        check("S28 dependency_graph = max(budget_ms+8000,33000)",
              tool_budget("dependency_graph", {})[0] == 33000,
              "dep=%s" % (tool_budget("dependency_graph", {}),))
        check("S28 immediate default is 5000/tool_default",
              tool_budget("immediate_eval", {}) == (5000, "tool_default"),
              "imm=%s" % (tool_budget("immediate_eval", {}),))
        CFG.max_call_ms = 120000
        CFG.max_call_hard_ms = 600000
        check("S28 explicit timeout_ms between generic ceiling and hard cap is NOT reduced",
              tool_budget("immediate_eval", {"timeout_ms": 150000}) == (150000, "param"),
              "imm150=%s" % (tool_budget("immediate_eval", {"timeout_ms": 150000}),))
        check("S28 explicit timeout_ms above the hard cap is capped and names +hard_cap",
              tool_budget("immediate_eval", {"timeout_ms": 700000}) == (600000, "param+hard_cap"),
              "imm700=%s" % (tool_budget("immediate_eval", {"timeout_ms": 700000}),))
        CFG.forward_timeout_ms = 120000

        # ---- S9: a >=3 s call is NOT cut (progress keeps the idle timer alive) --------
        _reset_registry(regdir)
        n9 = uuid.uuid4().hex[:16]
        p9 = "\\\\.\\pipe\\vb6mcp-1-%d-slow-%s" % (mine, n9)
        s9 = StubPipeAddIn(p9, n9, mine, project="SlowOne", tools=tools, behaviour="slow",
                           slow_ms=3000, progress_ms=300).start()
        stubs.append(s9)
        _write_instance(regdir, mine, project="SlowOne", pipe=p9, nonce=n9,
                        instance_id="%d-%s" % (mine, n9), transport="pipe")
        discover(force=True)
        t0 = time.monotonic()
        text, iserr, meta = handle_jsonrpc(_jreq("tools/call",
                                                 {"name": "make_project", "arguments": {}}))
        body = _resp_body(text)
        elapsed = (time.monotonic() - t0) * 1000.0
        check("S9 a long call completes over the pipe (progress frames keep it alive)",
              iserr is False and body.get("ok") is True and elapsed >= 3000,
              "elapsedMs=%.0f ok=%s" % (elapsed, body.get("ok")))

        # ---- S10: explicit transport=tcp (spike control / rollback track) -------------
        _reset_registry(regdir)
        hs = StubHttpAddIn(tools, project="HttpOnly").start()
        http_stubs.append(hs)
        _write_instance(regdir, mine, port=hs.port, project="HttpOnly")
        discover(force=True)
        CFG.transport = "tcp"
        text, iserr, meta = handle_jsonrpc(_jreq("tools/list", {}))
        tl = json.loads(text or "{}")
        check("S10 VB6MCP_TRANSPORT=tcp -> forced_tcp and the same 110-tool face",
              meta.get("transport") == "tcp" and meta.get("transportReason") == "forced_tcp"
              and len((tl.get("result") or {}).get("tools") or []) == 110,
              "transport=%s reason=%s count=%d" % (meta.get("transport"),
                                                   meta.get("transportReason"),
                                                   len((tl.get("result") or {}).get("tools") or [])))

        # ---- S11: VB6MCP_TRANSPORT=pipe but no pipe at all -> explicit, visible error --
        CFG.transport = "pipe"
        text, iserr, meta = handle_jsonrpc(_jreq("tools/call",
                                                 {"name": "mcp_info", "arguments": {}}))
        body = _resp_body(text)
        check("S11 explicit pipe with no pipe is an EXPLICIT error (sec 2.3), never a silent tcp hop",
              iserr is True and body.get("reason") == "pipe_unavailable",
              "reason=%s transport=%s" % (body.get("reason"), body.get("transport")))
        CFG.transport = ""

        # ---- S12: N instances multi-plexed + addressing + _target stripping -----------
        _reset_registry(regdir)
        insts = []
        for proj in ("Alpha", "Beta"):
            nn = uuid.uuid4().hex[:16]
            pp = "\\\\.\\pipe\\vb6mcp-1-%d-%s-%s" % (mine, proj.lower(), nn)
            st = StubPipeAddIn(pp, nn, mine, project=proj, tools=tools).start()
            stubs.append(st)
            _write_instance(regdir, mine, project=proj, pipe=pp, nonce=nn,
                            instance_id="%d-%s" % (mine, nn), transport="pipe")
            insts.append(st)
        discover(force=True)
        text, iserr, meta = handle_jsonrpc(_jreq("tools/call",
                                                 {"name": "get_project_info", "arguments": {}}))
        body = _resp_body(text)
        check("S12 two live instances -> ambiguous_target (never guessed)",
              iserr is True and body.get("reason") == "ambiguous_target",
              "reason=%s instances=%d" % (body.get("reason"), len(body.get("instances") or [])))
        text, iserr, meta = handle_jsonrpc(_jreq("tools/call", {
            "name": "get_project_info", "arguments": {"_target": "Beta"}}))
        body = _resp_body(text)
        check("S12 _target=<project> addresses exactly that instance over its own pipe",
              body.get("project") == "Beta" and meta.get("transport") == "pipe",
              "project=%s transport=%s" % (body.get("project"), meta.get("transport")))
        check("S12 _target is STRIPPED before it reaches the add-in",
              isinstance(insts[1].last_args, dict) and "_target" not in insts[1].last_args,
              "stub saw %s" % insts[1].last_args)
        text, iserr, meta = handle_jsonrpc(_jreq("tools/call", {
            "name": "get_project_info", "arguments": {"_target": "Alpha"}}))
        body = _resp_body(text)
        check("S12 both instances hold their OWN pipe session (N pipes multiplexed)",
              all(st.sessions == 1 for st in insts) and body.get("project") == "Alpha",
              "sessions=%s projects=[%s,%s]" % ([st.sessions for st in insts],
                                                insts[0].project, insts[1].project))

        # ---- S14: the stdio entry ----------------------------------------------------
        # a FRESH registry dir + a FRESH stub: the parent process must not already hold this
        # pipe's single instance (nMaxInstances=1), or the child would legitimately get busy
        stdiodir = os.path.join(tmpdir, "instances_stdio")
        ns = uuid.uuid4().hex[:16]
        ps = "\\\\.\\pipe\\vb6mcp-1-%d-stdio-%s" % (mine, ns)
        ss = StubPipeAddIn(ps, ns, mine, project="StdioProj", tools=tools).start()
        stubs.append(ss)
        _write_instance(stdiodir, mine, project="StdioProj", pipe=ps, nonce=ns,
                        instance_id="%d-%s" % (mine, ns), transport="pipe")
        env = _child_env({"VB6MCP_INSTANCES_DIR": stdiodir})
        proc = subprocess.run(
            [sys.executable, os.path.join(HERE, "vb6mcp_bridge.py"), "--stdio"],
            input=_jreq("tools/list", {}) + "\n" + _jreq("tools/call",
                                                          {"name": "get_project_info",
                                                           "arguments": {}}, rid=2) + "\n",
            capture_output=True, text=True, encoding="utf-8", errors="replace",
            env=env, timeout=60, cwd=REPO_ROOT)
        lines = [l for l in (proc.stdout or "").splitlines() if l.strip()]
        got_tools = 0
        got_proj = None
        for line in lines:
            j = json.loads(line)
            res = j.get("result") or {}
            if isinstance(res.get("tools"), list):
                got_tools = len(res["tools"])
            for c in (res.get("content") or []):
                try:
                    got_proj = json.loads(c["text"]).get("project")
                except Exception:
                    pass
        check("S14 the --stdio entry serves the same face (line-delimited JSON-RPC)",
              len(lines) == 2 and got_tools == 110 and got_proj == "StdioProj",
              "lines=%d tools=%d project=%s stdout=%s stderr=%s" % (
                  len(lines), got_tools, got_proj, (proc.stdout or "")[:300],
                  (proc.stderr or "")[:160]))

        # ---- S16: server-side deny frames are surfaced, never a hang --------------------
        _reset_registry(regdir)
        n16 = uuid.uuid4().hex[:16]
        p16 = "\\\\.\\pipe\\vb6mcp-1-%d-deny-%s" % (mine, n16)
        s16 = StubPipeAddIn(p16, n16, mine, project="DenyBig", tools=tools,
                            behaviour="deny_frame_too_large").start()
        stubs.append(s16)
        _write_instance(regdir, mine, project="DenyBig", pipe=p16, nonce=n16,
                        instance_id="%d-%s" % (mine, n16), transport="pipe")
        discover(force=True)
        t0 = time.monotonic()
        text, iserr, meta = handle_jsonrpc(_jreq("tools/call", {"name": "build", "arguments": {}}))
        body = _resp_body(text)
        check("S16a a server deny frame (frame_too_large) reaches the receipt in <3 s",
              iserr is True and body.get("reason") == "pipe_denied"
              and body.get("serverReason") == "frame_too_large"
              and (time.monotonic() - t0) < 3.0,
              "reason=%s serverReason=%s elapsedMs=%.0f" % (
                  body.get("reason"), body.get("serverReason"), (time.monotonic() - t0) * 1000.0))

        _reset_registry(regdir)
        n16b = uuid.uuid4().hex[:16]
        p16b = "\\\\.\\pipe\\vb6mcp-1-%d-badhello-%s" % (mine, n16b)
        s16b = StubPipeAddIn(p16b, n16b, mine, project="BadHello", tools=tools,
                             behaviour="deny_hello", deny_reason="bad_hello").start()
        stubs.append(s16b)
        _write_instance(regdir, mine, project="BadHello", pipe=p16b, nonce=n16b,
                        instance_id="%d-%s" % (mine, n16b), transport="pipe")
        discover(force=True)
        t0 = time.monotonic()
        text, iserr, meta = handle_jsonrpc(_jreq("tools/call", {"name": "build", "arguments": {}}))
        body = _resp_body(text)
        check("S16b a bad_hello deny fails the handshake with the server reason, no hang",
              iserr is True and body.get("reason") == "pipe_connect_failed"
              and "bad_hello" in str(body.get("message"))
              and (time.monotonic() - t0) < 3.0,
              "reason=%s message=%s" % (body.get("reason"), str(body.get("message"))[:70]))

        # ---- S18: the 8 MiB frame ceiling ------------------------------------------------
        raised8, raised9 = False, False
        exact = frame_wrap(b"x" * FRAME_MAX)
        try:
            frame_wrap(b"x" * (FRAME_MAX + 1))
        except FrameError as exc:
            raised9 = exc.reason == "frame_too_large"
        buf18 = FrameBuffer()
        buf18.feed((FRAME_MAX + 1).to_bytes(4, "little"))
        try:
            buf18.pop()
        except FrameError as exc:
            raised8 = exc.reason == "frame_too_large"
        check("S18a exactly 8 MiB is a legal frame; 8 MiB + 1 is refused (writer side)",
              len(exact) == 4 + FRAME_MAX and raised9 is True,
              "header=%d raised=%s" % (len(exact) - FRAME_MAX, raised9))
        check("S18b a peer DECLARING > 8 MiB is refused by the reader (never buffered)",
              raised8 is True, "raised=%s" % raised8)

        # ---- S16c: an over-8MiB REQUEST is refused structurally, nothing goes on the wire
        _reset_registry(regdir)
        n16c = uuid.uuid4().hex[:16]
        p16c = "\\\\.\\pipe\\vb6mcp-1-%d-over-%s" % (mine, n16c)
        s16c = StubPipeAddIn(p16c, n16c, mine, project="OverBig", tools=tools).start()
        stubs.append(s16c)
        _write_instance(regdir, mine, project="OverBig", pipe=p16c, nonce=n16c,
                        instance_id="%d-%s" % (mine, n16c), transport="pipe")
        discover(force=True)
        before16 = s16c.requests
        big_req = _jreq("tools/call", {"name": "write_file",
                                       "arguments": {"content": "x" * (FRAME_MAX + 16)}})
        text, iserr, meta = handle_jsonrpc(big_req)
        body = _resp_body(text)
        check("S16c an over-8MiB request fails structurally and is never written to the pipe",
              iserr is True and body.get("reason") == "frame_too_large"
              and s16c.requests == before16,
              "reason=%s stubRequests=%d->%d" % (body.get("reason"), before16, s16c.requests))

        # ---- S18c: a real multi-megabyte frame round-trips (short reads are normal) -----
        big_n = 7500000
        t0 = time.monotonic()
        text, iserr, meta = handle_jsonrpc(_jreq("tools/call", {
            "name": "__bigframe", "arguments": {"n": big_n}}))
        body = _resp_body(text)
        check("S18c a 7.5 MB response frame round-trips intact over the pipe",
              iserr is False and body.get("ok") is True
              and len(body.get("filler") or "") == big_n,
              "filler=%d elapsedMs=%.0f" % (len(body.get("filler") or ""),
                                            (time.monotonic() - t0) * 1000.0))

        # ---- S17: VB6MCP_PIPE_REPLAY=1 opts into replaying an already-written request ----
        _reset_registry(regdir)
        hs17 = StubHttpAddIn(tools, project="HttpReplay").start()
        http_stubs.append(hs17)
        n17 = uuid.uuid4().hex[:16]
        p17 = "\\\\.\\pipe\\vb6mcp-1-%d-replay-%s" % (mine, n17)
        s17 = StubPipeAddIn(p17, n17, mine, project="ReplayOne", tools=tools,
                            behaviour="die").start()
        stubs.append(s17)
        _write_instance(regdir, mine, port=hs17.port, project="ReplayOne", pipe=p17, nonce=n17,
                        instance_id="%d-%s" % (mine, n17), transport="pipe")
        discover(force=True)
        CFG.pipe_replay = True
        text, iserr, meta = handle_jsonrpc(_jreq("tools/call",
                                                 {"name": "mcp_info", "arguments": {}}))
        body = _resp_body(text)
        check("S17 VB6MCP_PIPE_REPLAY=1 replays a fully-written request on the tcp track",
              iserr is False and body.get("ok") is True and meta.get("transport") == "tcp"
              and (meta.get("flags") or {}).get("transportFallback") is True
              and meta.get("fellBackFrom") == "pipe_stalled",
              "transport=%s from=%s ok=%s" % (meta.get("transport"), meta.get("fellBackFrom"),
                                              body.get("ok")))
        CFG.pipe_replay = False
        hs17.stop()
        http_stubs.remove(hs17)

        # ---- S19: our OWN live session must never look like pipe_busy -------------------
        _reset_registry(regdir)
        n19 = uuid.uuid4().hex[:16]
        p19 = "\\\\.\\pipe\\vb6mcp-1-%d-own-%s" % (mine, n19)
        s19 = StubPipeAddIn(p19, n19, mine, project="OwnOne", tools=tools).start()
        stubs.append(s19)
        _write_instance(regdir, mine, project="OwnOne", pipe=p19, nonce=n19,
                        instance_id="%d-%s" % (mine, n19), transport="pipe")
        discover(force=True)
        text, iserr, meta = handle_jsonrpc(_jreq("tools/list", {}))
        first_transport = meta.get("transport")
        _drop_pipe_probe(discover()["live"][0])   # expire the probe cache WHILE we hold it
        text, iserr, meta = handle_jsonrpc(_jreq("tools/list", {}))
        check("S19 our own live pipe session is never mistaken for pipe_busy",
              first_transport == "pipe" and meta.get("transport") == "pipe"
              and meta.get("transportReason") == "auto_pipe_ok" and s19.sessions == 1,
              "first=%s afterProbeCacheExpiry=%s reason=%s sessions=%d" % (
                  first_transport, meta.get("transport"), meta.get("transportReason"),
                  s19.sessions))

        # ---- S20: welcome.project is NOT an identity input ------------------------------
# measured that welcome.project is the LIVE ActiveVBProject.Name at handshake
        # time: it may be "" (no project) and it may CHANGE after the add-in starts.  It is a
        # snapshot, never the authoritative project id - that is the registry project/projectPath.
        for idx, (welcome_project, reg_project) in enumerate([("WELCOME_OTHER", "RegName"),
                                                              ("", "RegNameNoProj")]):
            _reset_registry(regdir)
            n20 = uuid.uuid4().hex[:16]
            p20 = "\\\\.\\pipe\\vb6mcp-1-%d-wproj%d-%s" % (mine, idx, n20)
            s20s = StubPipeAddIn(p20, n20, mine, project=welcome_project, tools=tools).start()
            stubs.append(s20s)
            _write_instance(regdir, mine, project=reg_project, pipe=p20, nonce=n20,
                            instance_id="%d-%s" % (mine, n20), transport="pipe")
            discover(force=True)
            text, iserr, meta = handle_jsonrpc(_jreq("tools/call",
                                                     {"name": "mcp_info", "arguments": {}}))
            body = _resp_body(text)
            check("S20 welcome.project=%r is neither identity nor the stamped project"
                  % welcome_project,
                  iserr is False and meta.get("transport") == "pipe"
                  and body.get("ide", {}).get("project") == reg_project
                  and s20s.sessions == 1,
                  "registryProject=%s stamped=%s transport=%s sessions=%d" % (
                      reg_project, body.get("ide", {}).get("project"), meta.get("transport"),
                      s20s.sessions))

        # ---- S21: a registry file written in ANSI (the add-in's own encoding) -------------
        # The VB6 default project name is literally 工程1 and the add-in writes the registry in
        # ANSI.  A strict UTF-8 read raised, the WHOLE entry was dropped, and the bridge answered
        # no_ide_running while the IDE was healthy - the t16 joint-run defect.
        _reset_registry(regdir)
        n21 = uuid.uuid4().hex[:16]
        p21 = "\\\\.\\pipe\\vb6mcp-1-%d-ansi-%s" % (mine, n21)
        s21 = StubPipeAddIn(p21, n21, mine, project="StubAnsi", tools=tools).start()
        stubs.append(s21)
        reg21 = os.path.join(regdir, n21 + ".json")
        with open(reg21, "w", encoding="gbk") as fh:
            # ensure_ascii=False is ESSENTIAL: the default escapes 工程1 to \uXXXX and the
            # fixture would be pure ASCII - it would not reproduce the defect at all.
            # projectPath carries non-ASCII too, and the decode must cover the WHOLE file.
            json.dump({"pid": mine, "port": 0, "transport": "pipe", "pipe": p21,
                       "sessionId": 1, "nonce": n21, "instanceId": "%d-%s" % (mine, n21),
                       "tcpEnabled": False, "project": "工程1",
                       "projectPath": "C:\\测试\\_test_proj.vbp"}, fh, ensure_ascii=False)
        discover(force=True)
        d21 = discover()
        raw21 = open(reg21, "rb").read()
        strict_utf8_ok = True
        try:
            raw21.decode("utf-8")
        except UnicodeDecodeError:
            strict_utf8_ok = False
        text, iserr, meta = handle_jsonrpc(_jreq("tools/list", {}))
        check("S21 an ANSI/GBK registry entry (project=工程1) is NOT dropped",
              len(d21["instances"]) == 1 and len(d21["live"]) == 1
              and meta.get("transport") == "pipe"
              and d21["instances"][0].registryRoot == regdir
              and d21["instances"][0].project == "工程1"
              and "测试" in (d21["instances"][0].projectPath or ""),
              "strictUtf8Decodes=%s entries=%d live=%d transport=%s encoding=%s project=%r" % (
                  strict_utf8_ok, len(d21["instances"]), len(d21["live"]),
                  meta.get("transport"),
                  d21["instances"][0].jsonEncoding if d21["instances"] else "-",
                  d21["instances"][0].project if d21["instances"] else None))

        # ---- S22: the hello ON THE WIRE is compact and its nonce is byte-identical --------
        hello_raw = s21.last_hello or b""
        try:
            hello_obj = json.loads(hello_raw.decode("utf-8"))
        except Exception:
            hello_obj = {}
        check("S22 the hello frame is compact and carries the byte-identical nonce",
              hello_raw.startswith(b'{"vb6mcp":"hello"') and b'"nonce":"' in hello_raw
              and b'"nonce": "' not in hello_raw and hello_obj.get("nonce") == n21,
              "payload=%s" % hello_raw[:110].decode("utf-8", "replace"))

        # ---- S27: the NEW writer escapes non-ASCII -> the file is pure ASCII ------
        # Writer-side fix: Util.AsciiEscapeJson turns every >127 code unit into \uXXXX, so the
        # registry file is always pure ASCII.  The Lead requires the TOLERANT CHAIN TO STAY
        # (legacy GBK entries still exist on disk), so both forms must be readable in the SAME
        # pass - that is the real transition state, not either form alone.  VB6 Hex$ emits
        # UPPERCASE escapes (\u5DE5), which is what this fixture reproduces verbatim.
        n27 = uuid.uuid4().hex[:16]
        p27 = "\\\\.\\pipe\\vb6mcp-1-%d-ascii-%s" % (mine, n27)
        s27 = StubPipeAddIn(p27, n27, mine, project="StubAscii", tools=tools).start()
        stubs.append(s27)
        reg27 = os.path.join(regdir, n27 + ".json")
        import re as _re
        blob = json.dumps({"pid": mine, "port": 0, "transport": "pipe", "pipe": p27,
                           "sessionId": 1, "nonce": n27,
                           "instanceId": "%d-%s" % (mine, n27), "tcpEnabled": False,
                           "project": "工程1", "projectPath": "C:\\测试\\_test_proj.vbp"},
                          ensure_ascii=True)
        blob = _re.sub(r"\\u([0-9a-fA-F]{4})", lambda m: "\\u" + m.group(1).upper(), blob)
        with open(reg27, "w", encoding="ascii") as fh:
            fh.write(blob)
        pure_ascii = all(b < 128 for b in blob.encode("ascii"))
        d27 = discover(force=True)
        by_key = dict((i.key, i) for i in d27["instances"])
        new27 = by_key.get("%d-%s" % (mine, n27))
        legacy27 = [i for i in d27["instances"] if i.jsonEncoding not in (None, "utf-8-sig")]
        check("S27 the NEW pure-ASCII \\uXXXX form decodes, alongside the legacy GBK entry",
              pure_ascii and new27 is not None and new27.project == "工程1"
              and "测试" in (new27.projectPath or "")
              and len(d27["instances"]) == 2 and len(legacy27) == 1,
              "pureAscii=%s entries=%d newProject=%r newPath=%r legacyEncoding=%s" % (
                  pure_ascii, len(d27["instances"]),
                  new27.project if new27 else None,
                  new27.projectPath if new27 else None,
                  legacy27[0].jsonEncoding if legacy27 else "-"))

        # ---- S23: a deny surfaces reason + detail RAW, never as no_ide_running -------------
        _reset_registry(regdir)
        n23 = uuid.uuid4().hex[:16]
        p23 = "\\\\.\\pipe\\vb6mcp-1-%d-denyx-%s" % (mine, n23)
        s23 = StubPipeAddIn(p23, n23, mine, project="DenyDetail", tools=tools,
                            behaviour="deny_hello", deny_reason="nonce_mismatch",
                            deny_detail="DEBUG-KIND").start()
        stubs.append(s23)
        _write_instance(regdir, mine, project="DenyDetail", pipe=p23, nonce=n23,
                        instance_id="%d-%s" % (mine, n23), transport="pipe")
        discover(force=True)
        text, iserr, meta = handle_jsonrpc(_jreq("tools/call",
                                                 {"name": "ide_state", "arguments": {}}))
        body = _resp_body(text)
        check("S23 a deny surfaces reason+detail raw and never collapses to no_ide_running",
              body.get("reason") != "no_ide_running"
              and body.get("serverReason") == "nonce_mismatch"
              and body.get("serverDetail") == "DEBUG-KIND",
              "reason=%s serverReason=%s serverDetail=%s" % (
                  body.get("reason"), body.get("serverReason"), body.get("serverDetail")))

        # ---- S24: a refusal says WHICH registry directory was read and what was in it ------
        _reset_registry(regdir)
        text, iserr, meta = handle_jsonrpc(_jreq("tools/call",
                                                 {"name": "ide_state", "arguments": {}}))
        body = _resp_body(text)
        check("S24 a refusal carries registryDir / entries / roots (self-diagnosing)",
              body.get("reason") == "no_ide_running" and bool(body.get("registryDir"))
              and body.get("registryEntries") == 0
              and isinstance(body.get("registryRoots"), list),
              "reason=%s dir=%s entries=%s roots=%d" % (
                  body.get("reason"), body.get("registryDir"), body.get("registryEntries"),
                  len(body.get("registryRoots") or [])))

        # ---- S25: VB6MCP_BRIDGE_LOG records the story (never a 0-byte log) ----------------
        _reset_registry(regdir)
        logf = os.path.join(tmpdir, "bridge_diag.log")
        CFG.log_path = logf
        n25 = uuid.uuid4().hex[:16]
        p25 = "\\\\.\\pipe\\vb6mcp-1-%d-log-%s" % (mine, n25)
        s25 = StubPipeAddIn(p25, n25, mine, project="LogOne", tools=tools).start()
        stubs.append(s25)
        _write_instance(regdir, mine, project="LogOne", pipe=p25, nonce=n25,
                        instance_id="%d-%s" % (mine, n25), transport="pipe")
        discover(force=True)
        handle_jsonrpc(_jreq("tools/list", {}))
        logtext = open(logf, encoding="utf-8").read() if os.path.isfile(logf) else ""
        check("S25 the bridge log carries discover/select/track/hello + the nonce self-check",
              ("event=discover" in logtext and "event=select" in logtext
               and "event=track" in logtext and "event=pipe_hello" in logtext
               and "expectedHex=" in logtext and "sendingHex=" in logtext
               and "equal=True" in logtext),
              "logLines=%d bytes=%d" % (logtext.count("\n"), len(logtext)))
        CFG.log_path = ""

        # ---- S26: a client that connects and says NOTHING must be evicted (5 s deadline) --
        _reset_registry(regdir)
        n26 = uuid.uuid4().hex[:16]
        p26 = "\\\\.\\pipe\\vb6mcp-1-%d-hd-%s" % (mine, n26)
        s26 = StubPipeAddIn(p26, n26, mine, project="HalfDead", tools=tools).start()
        stubs.append(s26)
        _write_instance(regdir, mine, project="HalfDead", pipe=p26, nonce=n26,
                        instance_id="%d-%s" % (mine, n26), transport="pipe")
        discover(force=True)
        hd = PipeClient(p26).connect(600)         # connect, then never write and never read
        t0 = time.monotonic()
        t_evict = None
        while time.monotonic() - t0 < 12:
            try:
                probe = PipeClient(p26).connect(200)
                t_evict = time.monotonic() - t0
                probe.close()
                break
            except PipeError:
                time.sleep(0.25)
        hd.close()
        check("S26 a half-dead client is evicted by the handshake deadline (pipe not held)",
              t_evict is not None and 3.0 <= t_evict <= 9.0 and s26.helloTimeouts >= 1,
              "evictedAfter=%.1fs helloTimeouts=%d" % (t_evict if t_evict else -1.0,
                                                       s26.helloTimeouts))

        # ---- S28: token authentication (verification plan V1-V8 of TOKEN_THREAT_MODEL.md) ----
        import urllib.error as _uerr
        _reset_registry(regdir)
        n28 = uuid.uuid4().hex[:16]
        p28 = "\\\\.\\pipe\\vb6mcp-1-%d-auth-%s" % (mine, n28)
        s28 = StubPipeAddIn(p28, n28, mine, project="AuthOne", tools=tools).start()
        stubs.append(s28)
        _write_instance(regdir, mine, project="AuthOne", pipe=p28, nonce=n28,
                        instance_id="%d-%s" % (mine, n28), transport="pipe")
        discover(force=True)
        SECRET = "SELFTEST-SECRET-TOKEN-0123456789"
        CFG.token, CFG.token_source = "", "none"          # phase 1: auth DISABLED (M-5)
        CFG.log_path = os.path.join(tmpdir, "auth_diag.log")
        port28 = _free_port()
        httpd28 = BridgeServer(("127.0.0.1", port28), BridgeHandler)
        threading.Thread(target=httpd28.serve_forever, kwargs={"poll_interval": 0.05},
                         daemon=True).start()
        CFG.port = port28

        def post28(path, payload_text, token=None, header="Authorization"):
            headers = {"Content-Type": "application/json"}
            if token is not None:
                headers[header] = ("Bearer " + token) if header == "Authorization" else token
            req = urllib.request.Request("http://127.0.0.1:%d%s" % (port28, path),
                                         data=payload_text.encode("utf-8"), headers=headers)
            try:
                with urllib.request.urlopen(req, timeout=20) as resp:
                    return resp.status, resp.read().decode("utf-8", "replace")
            except _uerr.HTTPError as exc:
                return exc.code, exc.read().decode("utf-8", "replace")

        def get28(path, token=None):
            headers = {}
            if token is not None:
                headers["Authorization"] = "Bearer " + token
            req = urllib.request.Request("http://127.0.0.1:%d%s" % (port28, path),
                                         headers=headers)
            try:
                with urllib.request.urlopen(req, timeout=20) as resp:
                    return resp.status, resp.read().decode("utf-8", "replace")
            except _uerr.HTTPError as exc:
                return exc.code, exc.read().decode("utf-8", "replace")

        st, b1 = post28("/rpc", _jreq("tools/list", {}))
        st_h, b_h = get28("/health")
        h1 = json.loads(b_h or "{}")
        check("V8a with NO token configured the face behaves exactly as before (M-5)",
              st == 200 and len((json.loads(b1 or "{}").get("result") or {}).get("tools") or []) == 110,
              "status=%s tools=%d" % (st, len((json.loads(b1 or "{}").get("result") or {}).get("tools") or [])))
        check("V8b an unauthenticated /health SAYS it is unauthenticated and warns about shared machines",
              st_h == 200 and h1.get("auth", {}).get("enabled") is False
              and bool(h1.get("auth", {}).get("guidance")),
              "auth=%s" % json.dumps(h1.get("auth"), ensure_ascii=False)[:150])

        CFG.token, CFG.token_source = SECRET, "env"       # phase 2: auth ENABLED
        write_call = _jreq("tools/call", {"name": "build", "arguments": {}})
        before28 = s28.requests
        st, body = post28("/rpc", write_call)
        r1 = _resp_body(body)
        check("V1 no token -> 401 token_missing AND the write tool did NOT run",
              st == 401 and r1.get("reason") == "token_missing" and s28.requests == before28,
              "status=%s reason=%s stubRequests=%d->%d" % (st, r1.get("reason"), before28,
                                                           s28.requests))
        st, body_wrong = post28("/rpc", write_call, token="wrong-token")
        r2 = _resp_body(body_wrong)
        check("V2 wrong token -> 401 token_mismatch AND the write tool did NOT run",
              st == 401 and r2.get("reason") == "token_mismatch" and s28.requests == before28,
              "status=%s reason=%s stubRequests=%d" % (st, r2.get("reason"), s28.requests))
        st, body_url = post28("/rpc?token=" + SECRET, write_call)
        r3 = _resp_body(body_url)
        check("V3 a credential in the URL -> 401 token_in_url_rejected (never honoured)",
              st == 401 and r3.get("reason") == "token_in_url_rejected" and s28.requests == before28,
              "status=%s reason=%s stubRequests=%d" % (st, r3.get("reason"), s28.requests))
        check("V6 a 401 echoes neither the supplied nor the expected credential",
              SECRET not in body_wrong and "wrong-token" not in body_wrong
              and SECRET not in body_url and SECRET not in body,
              "wrongBodyLen=%d urlBodyLen=%d" % (len(body_wrong), len(body_url)))
        st, body_ok = post28("/rpc", _jreq("tools/list", {}), token=SECRET)
        check("V4 correct token -> the face is UNCHANGED (110 tools)",
              st == 200 and len((json.loads(body_ok or "{}").get("result") or {}).get("tools") or []) == 110,
              "status=%s tools=%d" % (st, len((json.loads(body_ok or "{}").get("result") or {}).get("tools") or [])))
        before28 = s28.requests
        st, body_call = post28("/rpc", write_call, token=SECRET)
        check("V4b correct token -> the call DOES reach the add-in (auth is not a wall)",
              st == 200 and s28.requests == before28 + 1,
              "status=%s stubRequests=%d->%d" % (st, before28, s28.requests))
        st, body_hdr = post28("/rpc", _jreq("tools/list", {}), token=SECRET,
                              header="X-VB6MCP-Token")
        check("V4c the alternative header X-VB6MCP-Token is accepted too",
              st == 200 and "tools" in body_hdr, "status=%s" % st)
        st, body_health2 = get28("/health", token=SECRET)
        check("V4d /health requires the token and reports auth when authorised",
              st == 200 and json.loads(body_health2 or "{}").get("auth", {}).get("enabled") is True,
              "status=%s" % st)
        st_noauth_health, _bh = get28("/health")
        check("V4e /health without a token leaks NOTHING (401, no instances/paths)",
              st_noauth_health == 401, "status=%s" % st_noauth_health)
        logtext28 = open(CFG.log_path, encoding="utf-8").read() if os.path.isfile(CFG.log_path) else ""
        check("V5 the bridge log records the rejection but NEVER the credential",
              SECRET not in logtext28 and "auth_reject" in logtext28
              and ("tokenLen=%d" % len(SECRET)) in logtext28,
              "logBytes=%d hasSecret=%s hasReject=%s" % (len(logtext28), SECRET in logtext28,
                                                         "auth_reject" in logtext28))
        # The non-loopback examples are TEST-NET reservations (RFC 5737): 203.0.113.0/24 and
        # 198.51.100.0/24.  They keep the "a non-loopback host is refused" semantics while making
        # the shipped tree carry no real intranet address.
        check("V7a only literal loopback hosts are accepted (0.0.0.0 impossible)",
              all(is_loopback_host(h) for h in ("127.0.0.1", "127.9.9.9", "localhost", "::1"))
              and not any(is_loopback_host(h) for h in ("0.0.0.0", "203.0.113.7", "198.51.100.7",
                                                        "example.com", "127.0.0.1.example.com")),
              "whitelist ok")
        check("V7b DEFAULT tier: serve() refuses a non-loopback host with exit 6 (no --allow-remote)",
              serve(8799, "0.0.0.0", "") == 6 and serve(8799, "203.0.113.7", "") == 6,
              "refused both")
        httpd28.shutdown()
        httpd28.server_close()
        CFG.token, CFG.token_source, CFG.log_path = "", "none", ""

        # =================================================================================
        # S29-S36: the ctl/1 CONTROL LANE, CLIENT HALF (R-1021-03 A4)
        # The SERVER half below is StubCtlAddIn, which implements the contract STRICTLY (hello/
        # welcome, sid+seq, six ops, nested tool receipt, 4.7 fast subset, 60 s lockout, limits
        # block, abandon preconditions).  Every assertion is therefore "the client obeys the
        # contract against a strict stub" - NOT a measurement against a live add-in.
        # =================================================================================
        def _ctl_pair(tag, token="", behaviour="normal", gate_tool="immediate_eval",
                      gate_held_ms=4210):
            nb = uuid.uuid4().hex[:16]
            pb = "\\\\.\\pipe\\vb6mcp-1-%d-%s-%s" % (mine, tag, nb)
            biz = StubPipeAddIn(pb, nb, mine, project="Ctl_" + tag, tools=tools,
                                behaviour=behaviour).start()
            stubs.append(biz)
            nc = uuid.uuid4().hex[:16]
            pc = "\\\\.\\pipe\\vb6mcp-ctl-1-%d-%s-%s" % (mine, tag, nc)
            ctl = StubCtlAddIn(pc, nc, mine, token=token, gate_tool=gate_tool,
                               gate_held_ms=gate_held_ms).start()
            stubs.append(ctl)
            return {"biz": biz, "pb": pb, "nb": nb, "ctl": ctl, "pc": pc, "nc": nc}

        def _ctl_register(tag, pair, ctl_nonce=None, ctl_proto="ctl/1", ctl_available=True):
            _reset_registry(regdir)
            _write_instance(regdir, mine, project="Ctl_" + tag, pipe=pair["pb"],
                            nonce=pair["nb"], instance_id="%d-%s" % (mine, pair["nb"]),
                            transport="pipe", ctlPipe=pair["pc"],
                            ctlNonce=(ctl_nonce or pair["nc"]), ctlProto=ctl_proto,
                            ctlAvailable=ctl_available, ctlAuth=["nonce", "acl"])
            discover(force=True)
            live = discover()["live"]
            return live[0] if live else None

        def _raw_hello(pair, proto=None, nonce=None, sid=None, extra=None):
            fr = {"vb6mcp": "ctlhello", "proto": (proto or CTL_PROTO),
                  "nonce": (nonce or pair["nc"]),
                  "instanceId": "%d-%s" % (mine, nonce or pair["nc"])}
            fr.update(extra or {})
            return fr

        # ---- S29: ctl AUTHENTICATION - the six failure grades stay APART -----------------
        pair = _ctl_pair("auth")
        _ctl_register("auth", pair, ctl_nonce="0" * 16)   # registry nonce != add-in nonce
        text, iserr, meta = handle_jsonrpc(_jreq("tools/call", {
            "name": "ide_state", "arguments": {"release_inflight": True}}))
        body = _resp_body(text)
        oob = body.get("oob") or {}
        check("S29a wrong ctl nonce -> ctl_unauthorized, oob.acked/applied=false, NO business fallback",
              iserr is True and body.get("reason") == "ctl_unauthorized"
              and oob.get("lane") == "ctl" and oob.get("why") == "ctl_unauthorized"
              and oob.get("acked") is False and oob.get("applied") is False
              and oob.get("delivered") == "rejected_by_server"
              and pair["biz"].requests == 0,
              "reason=%s why=%s bizRequests=%d" % (body.get("reason"), oob.get("why"),
                                                   pair["biz"].requests))

        pair = _ctl_pair("proto")
        _ctl_register("proto", pair, ctl_proto="ctl/2")
        text, iserr, meta = handle_jsonrpc(_jreq("tools/call", {
            "name": "ide_state", "arguments": {"release_inflight": True}}))
        body = _resp_body(text)
        check("S29b an UNSUPPORTED ctlProto is never coerced: lane unavailable -> the "
              "pre-existing business path (unchanged)",
              iserr is False and body.get("stub") is True and pair["biz"].requests == 1,
              "stub=%s bizRequests=%d" % (body.get("stub"), pair["biz"].requests))

        pair = _ctl_pair("rawproto")
        pc29 = PipeClient(pair["pc"]).connect(500)
        pc29.send_frame(_raw_hello(pair, proto="ctl/9"), time.monotonic() + 2.0)
        fr29, _e29 = pc29.recv_frame(time.monotonic() + 2.0)
        pc29.close()
        check("S29c bad proto -> ctl_unauthorized/bad_proto and the connection is CLOSED",
              (fr29 or {}).get("reason") == "ctl_unauthorized"
              and (fr29 or {}).get("detail") == "bad_proto" and pair["ctl"].denied == 1,
              "frame=%s denied=%d" % (fr29, pair["ctl"].denied))

        pair = _ctl_pair("rawsid")
        pc29 = PipeClient(pair["pc"]).connect(500)
        pc29.send_frame(_raw_hello(pair), time.monotonic() + 2.0)
        wel29, _e29 = pc29.recv_frame(time.monotonic() + 2.0)
        pc29.send_frame({"vb6mcp": "ctl", "sid": "deadbeefdeadbeef", "seq": 0, "op": "ping"},
                        time.monotonic() + 2.0)
        fr29, _e29 = pc29.recv_frame(time.monotonic() + 2.0)
        pc29.close()
        check("S29d an unknown sid -> ctl_unauthorized/bad_session (out-of-sync channel untrusted)",
              (wel29 or {}).get("vb6mcp") == "ctlwelcome"
              and (fr29 or {}).get("reason") == "ctl_unauthorized"
              and (fr29 or {}).get("detail") == "bad_session",
              "frame=%s" % (fr29,))

        pair = _ctl_pair("rawseq")
        pc29 = PipeClient(pair["pc"]).connect(500)
        pc29.send_frame(_raw_hello(pair), time.monotonic() + 2.0)
        wel29, _e29 = pc29.recv_frame(time.monotonic() + 2.0)
        sid29 = (wel29 or {}).get("sid")
        pc29.send_frame({"vb6mcp": "ctl", "sid": sid29, "seq": 0, "op": "ping"},
                        time.monotonic() + 2.0)
        ok29, _e29 = pc29.recv_frame(time.monotonic() + 2.0)
        pc29.send_frame({"vb6mcp": "ctl", "sid": sid29, "seq": 0, "op": "ping"},
                        time.monotonic() + 2.0)
        fr29, _e29 = pc29.recv_frame(time.monotonic() + 2.0)
        pc29.close()
        check("S29e a REPLAYED seq -> ctl_unauthorized/bad_seq (the first seq=seqBase was accepted)",
              (ok29 or {}).get("ok") is True and (fr29 or {}).get("reason") == "ctl_unauthorized"
              and (fr29 or {}).get("detail") == "bad_seq",
              "first=%s replay=%s" % ((ok29 or {}).get("ok"), (fr29 or {}).get("detail")))

        # Lead ruling 1: the contract pins the FIRST frame to EXACTLY seqBase, so seqBase+1 as a
        # FIRST frame is a rejection - the reverse direction of S29e, asserted on the wire.
        pair = _ctl_pair("rawseq0")
        pc29 = PipeClient(pair["pc"]).connect(500)
        pc29.send_frame(_raw_hello(pair), time.monotonic() + 2.0)
        wel29, _e29 = pc29.recv_frame(time.monotonic() + 2.0)
        sid29 = (wel29 or {}).get("sid")
        base29 = int((wel29 or {}).get("seqBase") or 0)
        pc29.send_frame({"vb6mcp": "ctl", "sid": sid29, "seq": base29 + 1, "op": "ping"},
                        time.monotonic() + 2.0)
        fr29b, _e29 = pc29.recv_frame(time.monotonic() + 2.0)
        pc29.close()
        check("S29g the FIRST frame must be EXACTLY seqBase: seqBase+1 -> ctl_unauthorized/bad_seq "
              "(never silently accepted)",
              (fr29b or {}).get("reason") == "ctl_unauthorized"
              and (fr29b or {}).get("detail") == "bad_seq",
              "seqBase=%s firstSent=%s frame=%s" % (base29, base29 + 1, fr29b))

        pair = _ctl_pair("lockout")
        for i in range(3):
            pcl = PipeClient(pair["pc"]).connect(500)
            pcl.send_frame(_raw_hello(pair, nonce="f" * 16), time.monotonic() + 2.0)
            last = pcl.recv_frame(time.monotonic() + 2.0)[0]
            pcl.close()
        _ctl_register("lockout", pair)      # correct nonce this time
        inst29 = discover()["live"][0]
        biz_before = pair["biz"].requests
        text, iserr, meta = handle_jsonrpc(_jreq("tools/call", {
            "name": "ide_state", "arguments": {"release_inflight": True}}))
        body = _resp_body(text)
        oob = body.get("oob") or {}
        check("S29f 3 unauthenticated attempts -> ctl_lockout for the NEXT attempt, 60 s, "
              "peer_busy (NOT unresponsive), no business fallback",
              iserr is True and body.get("reason") == "ctl_lockout"
              and oob.get("why") == "ctl_lockout" and oob.get("acked") is False
              and oob.get("peerState") == "peer_busy"
              and oob.get("suggestedAction") == "retry_once"
              and oob.get("retryAfterMs") == 60000
              and pair["ctl"].lockouts >= 1 and pair["biz"].requests == biz_before,
              "lastDeny=%s why=%s state=%s retry=%s lockouts=%d" % (
                  last, oob.get("why"), oob.get("peerState"), oob.get("retryAfterMs"),
                  pair["ctl"].lockouts))

        # ---- S30: bearer is OPTIONAL - 2 factors off, 3 factors on ----------------------
        SECRET30 = "SELFTEST-CTL-BEARER-0123456789"
        pair = _ctl_pair("bearer-off")
        inst30 = _ctl_register("bearer-off", pair)
        CFG.token, CFG.token_source = "", "none"
        res30, info30 = ctl_request(inst30, "ping")
        hello30 = json.loads((pair["ctl"].last_hello or b"{}").decode("utf-8"))
        check("S30a bearer OFF: welcome.authBasis == [nonce,acl] and the hello carries NO auth block",
              res30.get("ok") is True and info30.get("authBasis") == ["nonce", "acl"]
              and "auth" not in hello30,
              "basis=%s helloKeys=%s" % (info30.get("authBasis"), sorted(hello30.keys())))

        pair = _ctl_pair("bearer-on", token=SECRET30)
        inst30 = _ctl_register("bearer-on", pair)
        CFG.token, CFG.token_source = SECRET30, "env"
        res30, info30 = ctl_request(inst30, "ping")
        hello30 = json.loads((pair["ctl"].last_hello or b"{}").decode("utf-8"))
        auth30 = hello30.get("auth") or {}
        check("S30b bearer ON (both sides): 3 factors, hello carries scheme+token+tokenFp(8 hex)",
              res30.get("ok") is True
              and info30.get("authBasis") == ["nonce", "acl", "token"]
              and auth30.get("scheme") == "bearer" and auth30.get("token") == SECRET30
              and len(str(auth30.get("tokenFp") or "")) == 8,
              "basis=%s fpLen=%d" % (info30.get("authBasis"),
                                     len(str(auth30.get("tokenFp") or ""))))

        pair = _ctl_pair("bearer-noclient", token=SECRET30)   # add-in REQUIRES a bearer
        inst30 = _ctl_register("bearer-noclient", pair)
        CFG.token, CFG.token_source = "", "none"
        try:
            ctl_request(inst30, "ping")
            denied30 = False
        except CtlError as exc:
            denied30 = (exc.why == "ctl_unauthorized" and "token_mismatch" in exc.detail)
        check("S30c the add-in REQUIRES a bearer the bridge does not have -> ctl_unauthorized/"
              "token_mismatch (never a silent downgrade)",
              denied30 is True, "denied=%s" % denied30)

        pair = _ctl_pair("bearer-bridgeonly")
        inst30 = _ctl_register("bearer-bridgeonly", pair)
        CFG.token, CFG.token_source = SECRET30, "env"
        res30, info30 = ctl_request(inst30, "ping")
        check("S30d bridge HAS a token, add-in needs none: accepted, and the receipt claims NO "
              "bearer factor (authBasis stays 2)",
              res30.get("ok") is True and info30.get("authBasis") == ["nonce", "acl"],
              "basis=%s" % (info30.get("authBasis"),))
        CFG.token, CFG.token_source = "", "none"

        # ---- S31: ping / state / release_gate SEMANTICS ---------------------------------
        pair = _ctl_pair("ops")
        inst31 = _ctl_register("ops", pair)
        res31, info31 = ctl_request(inst31, "ping")
        check("S31a ping proves liveness and reports gate + businessPipe state",
              res31.get("ok") is True and res31.get("pid") == mine
              and isinstance(res31.get("gate"), dict)
              and (res31.get("businessPipe") or {}).get("state") == "busy"
              and res31.get("limits") == StubCtlAddIn.LIMITS,
              "pid=%s gate=%s" % (res31.get("pid"), res31.get("gate")))
        res31, info31 = ctl_request(inst31, "state")
        check("S31b state reports gate.depth/mutBusy/tool + cancel/release snapshots",
              res31.get("ok") is True and res31["gate"]["mutBusy"] is True
              and res31["gate"]["tool"] == "immediate_eval"
              and res31["gate"]["depth"] == 1 and isinstance(res31.get("cancel"), dict),
              "gate=%s" % (res31.get("gate"),))
        res31, info31 = ctl_request(inst31, "release_gate", {"minHeldMs": 5000})
        check("S31c release_gate(minHeldMs=5000) with heldMs=4210 DECLINES (released=false) "
              "without inventing a new reason code",
              res31.get("ok") is True and res31.get("released") is False
              and res31.get("heldMs") == 4210 and not res31.get("reason"),
              "released=%s heldMs=%s" % (res31.get("released"), res31.get("heldMs")))
        res31, info31 = ctl_request(inst31, "release_gate", {"minHeldMs": 0})
        check("S31d release_gate(0) releases the gate; statementInterrupted is ALWAYS false",
              res31.get("released") is True
              and res31.get("releasedTool") == "immediate_eval"
              and res31.get("statementInterrupted") is False
              and res31.get("gateReleaseCount") == 1 and res31.get("depthAfter") == 0,
              "released=%s tool=%s count=%s" % (res31.get("released"),
                                                res31.get("releasedTool"),
                                                res31.get("gateReleaseCount")))
        res31, info31 = ctl_request(inst31, "release_gate", {"minHeldMs": 0})
        check("S31e release_gate with nothing in flight -> nothing_in_flight (NOT an auth failure)",
              res31.get("ok") is False and res31.get("reason") == "nothing_in_flight",
              "reason=%s" % (res31.get("reason"),))

        pair = _ctl_pair("intent")
        inst31 = _ctl_register("intent", pair)
        text, iserr, meta = handle_jsonrpc(_jreq("tools/call", {
            "name": "ide_state", "arguments": {"release_inflight": True,
                                               "release_min_held_ms": 0}}))
        body = _resp_body(text)
        oob = body.get("oob") or {}
        check("S31f ide_state{release_inflight} is served by the CTL LANE: oob acked/ackOp/applied, "
              "limits relayed, business pipe untouched",
              iserr is False and body.get("source") == "ctl_lane"
              and body.get("intent") == "release_inflight"
              and body.get("released") is True and body.get("statementInterrupted") is False
              and oob.get("lane") == "ctl" and oob.get("acked") is True
              and oob.get("ackOp") == "release_gate" and oob.get("applied") is True
              and isinstance(oob.get("helloMs"), int) and isinstance(oob.get("rttMs"), int)
              and body.get("limits") == StubCtlAddIn.LIMITS
              and pair["biz"].requests == 0,
              "acked=%s ackOp=%s applied=%s bizRequests=%d" % (
                  oob.get("acked"), oob.get("ackOp"), oob.get("applied"),
                  pair["biz"].requests))

        # ---- S32: cancel SEMANTICS -----------------------------------------------------
        pair = _ctl_pair("cancel")
        inst32 = _ctl_register("cancel", pair)
        res32, info32 = ctl_request(inst32, "cancel", {"target": "mcp_info"})
        check("S32a cancel for a tool that is NOT in flight -> nothing_in_flight + inFlightTool",
              res32.get("ok") is False and res32.get("reason") == "nothing_in_flight"
              and res32.get("inFlightTool") == "immediate_eval",
              "reason=%s inFlightTool=%s" % (res32.get("reason"), res32.get("inFlightTool")))
        res32, info32 = ctl_request(inst32, "cancel", {"target": "*"})
        check("S32b cancel '*' -> requested + honoredAtNextPump (cooperative)",
              res32.get("ok") is True and res32.get("requested") is True
              and res32.get("honoredAtNextPump") is True,
              "requested=%s" % (res32.get("requested"),))
        ctl_request(inst32, "release_gate", {"minHeldMs": 0})
        res32, info32 = ctl_request(inst32, "cancel", {"target": "*"})
        check("S32c cancel with nothing in flight -> nothing_in_flight",
              res32.get("ok") is False and res32.get("reason") == "nothing_in_flight",
              "reason=%s" % (res32.get("reason"),))

        # ---- S33: abandon SEMANTICS (incl. args.minHeldMs) -----------------------------
        pair = _ctl_pair("abandon")
        inst33 = _ctl_register("abandon", pair)
        res33, info33 = ctl_request(inst33, "abandon", {"confirm": False})
        check("S33a abandon before any release -> gate_not_released + howToProceed",
              res33.get("ok") is False and res33.get("reason") == "gate_not_released"
              and "release_gate" in str(res33.get("howToProceed")),
              "reason=%s how=%s" % (res33.get("reason"), res33.get("howToProceed")))
        res33, info33 = ctl_request(inst33, "abandon", {"confirm": True, "minHeldMs": 999999})
        check("S33b abandon confirm=true with minHeldMs above heldMs is STILL refused",
              res33.get("ok") is False and res33.get("reason") == "gate_not_released",
              "reason=%s" % (res33.get("reason"),))
        res33, info33 = ctl_request(inst33, "abandon", {"confirm": True, "minHeldMs": 0})
        check("S33c abandon confirm=true forces the release first, then abandons "
              "(responseLost + businessPipe.closed)",
              res33.get("ok") is True and res33.get("abandoned") is True
              and res33.get("responseLost") is True
              and (res33.get("businessPipe") or {}).get("closed") is True,
              "abandoned=%s pipe=%s" % (res33.get("abandoned"), res33.get("businessPipe")))

        # ---- S34: ABANDON THEN RECONNECT (the Lead asked for the measured cost + reason) --
        pair = _ctl_pair("reconnect")
        inst34 = _ctl_register("reconnect", pair)
        ctl_request(inst34, "release_gate", {"minHeldMs": 0})
        first34, info34a = ctl_request(inst34, "state")
        t34 = time.monotonic()
        res34, info34 = ctl_request(inst34, "abandon", {"confirm": False})
        abandon_ms = int((time.monotonic() - t34) * 1000)
        t34 = time.monotonic()
        re34 = {"ok": False, "why": "not_attempted"}
        try:
            res34b, info34b = ctl_request(inst34, "ping")
            re34 = {"ok": res34b.get("ok") is True, "why": "ok",
                    "ms": int((time.monotonic() - t34) * 1000),
                    "sid": info34b.get("sid"), "helloMs": info34b.get("helloMs")}
        except CtlError as exc:
            re34 = {"ok": False, "why": exc.why, "ms": int((time.monotonic() - t34) * 1000)}
        check("S34 abandon then RECONNECT: a fresh ctl connection with a NEW sid succeeds, "
              "cost measured (ms) and the reason code reported",
              res34.get("abandoned") is True and re34.get("ok") is True
              and re34.get("why") == "ok" and re34.get("sid") and re34.get("sid") != info34.get("sid")
              and re34.get("ms", 99999) < 3000,
              "abandonMs=%d reconnectMs=%s reason=%s sidChanged=%s" % (
                  abandon_ms, re34.get("ms"), re34.get("why"),
                  bool(re34.get("sid") and re34.get("sid") != info34.get("sid"))))
        print("    S34 abandon->reconnect: abandonMs=%d reconnectMs=%s reason=%s helloMs=%s"
              % (abandon_ms, re34.get("ms"), re34.get("why"), re34.get("helloMs")))

        # ---- S35: the 4.7 READ-ONLY FAST SUBSET + the NESTED tool receipt --------------
        pair = _ctl_pair("lane")
        inst35 = _ctl_register("lane", pair)
        res35, info35 = ctl_request(inst35, "tool", {"name": "ide_state", "arguments": {}})
        check("S35a tool op returns a NESTED receipt (outer ok = frame verdict; receipt.ok = "
              "tool verdict) and tool fields are NEVER flattened",
              res35.get("ok") is True and res35.get("source") == "ctl_lane"
              and isinstance(res35.get("receipt"), dict)
              and res35["receipt"].get("ok") is True
              and res35["receipt"].get("tool") == "ide_state"
              and "stub" not in res35 and "tool" not in res35,
              "outerKeys=%s receiptKeys=%s" % (sorted(res35.keys()),
                                               sorted((res35.get("receipt") or {}).keys())))
        res35, info35 = ctl_request(inst35, "tool", {"name": "dependency_graph", "arguments": {}})
        check("S35b a SLOW read-only tool is refused: readonly_lane_refused/"
              "slow_tool_not_on_ctl_lane (+ allowed list, failureKind=toolError, an ANSWER not a raise)",
              res35.get("ok") is False and res35.get("reason") == "readonly_lane_refused"
              and res35.get("detail") == "slow_tool_not_on_ctl_lane"
              and res35.get("failureKind") == "toolError"
              and CTL_FAST_TOOLS[0] in (res35.get("allowed") or []),
              "detail=%s allowed=%d" % (res35.get("detail"),
                                        len(res35.get("allowed") or [])))
        res35, info35 = ctl_request(inst35, "tool", {"name": "build", "arguments": {}})
        check("S35c a WRITE tool is refused: readonly_lane_refused/not_read_only",
              res35.get("ok") is False and res35.get("reason") == "readonly_lane_refused"
              and res35.get("detail") == "not_read_only",
              "detail=%s" % (res35.get("detail"),))

        # ---- S36: the SWALLOWED-RAISE fixture (max budget + cancel interaction) ---------
        # dependency_graph's own budget is 33000 by table; the HARD cap is what decides here, so
        # VB6MCP_MAX_CALL_MS / VB6MCP_MAX_CALL_HARD_MS = 2000 makes the whole assertion fast
        # while exercising the REAL two-layer arithmetic (maxMsSource names +hard_cap).
        _reset_registry(regdir)
        n36 = uuid.uuid4().hex[:16]
        p36 = "\\\\.\\pipe\\vb6mcp-1-%d-raise-%s" % (mine, n36)
        s36 = StubPipeAddIn(p36, n36, mine, project="RaiseOne", tools=tools,
                            behaviour="silent").start()
        stubs.append(s36)
        _write_instance(regdir, mine, project="RaiseOne", pipe=p36, nonce=n36,
                        instance_id="%d-%s" % (mine, n36), transport="pipe")
        discover(force=True)
        CFG.max_call_ms = 2000
        CFG.max_call_hard_ms = 2000
        t36 = time.monotonic()
        text, iserr, meta = handle_jsonrpc(_jreq("tools/call", {
            "name": "dependency_graph", "arguments": {}}))
        body = _resp_body(text)
        el36 = (time.monotonic() - t36) * 1000
        check("S36a swallowed raise: dependency_graph at VB6MCP_MAX_CALL_MS=2000 -> "
              "reason=timeout with maxMs/maxMsSource/elapsedMs",
              body.get("reason") == "timeout" and body.get("maxMs") == 2000
              and str(body.get("maxMsSource") or "").endswith("hard_cap")
              and body.get("elapsedMs") is not None and 1900 <= el36 < 6000
              and body.get("failureKind") == "peer_unresponsive"
              and isinstance(body.get("heldBy"), dict),
              "reason=%s maxMs=%s src=%s elapsed=%d" % (
                  body.get("reason"), body.get("maxMs"), body.get("maxMsSource"), el36))

        pair = _ctl_pair("raise", behaviour="silent", gate_tool="dependency_graph",
                         gate_held_ms=4210)
        _ctl_register("raise", pair)
        cancel_out = {}

        def _cancel_worker():
            time.sleep(0.35)
            try:
                t_c, ie_c, _m_c = handle_jsonrpc(_jreq("tools/call", {
                    "name": "ide_state",
                    "arguments": {"cancel_inflight": True,
                                  "cancel_target": "dependency_graph"}}, rid=771))
                cancel_out["body"] = _resp_body(t_c)
                cancel_out["isError"] = ie_c
            except Exception as exc:
                cancel_out["error"] = repr(exc)

        th = threading.Thread(target=_cancel_worker, name="selftest-cancel", daemon=True)
        th.start()
        text, iserr, meta = handle_jsonrpc(_jreq("tools/call", {
            "name": "dependency_graph", "arguments": {}}))
        th.join(15)
        body = _resp_body(text)
        cb = cancel_out.get("body") or {}
        coob = cb.get("oob") or {}
        check("S36b a ctl-ACKED cancel is NOT overwritten by the timeout: receipt keeps "
              "cancelled_by_request + timeoutAlsoDue and the budget facts",
              cancel_out.get("isError") is False and coob.get("acked") is True
              and coob.get("applied") is True
              and body.get("reason") == "cancelled_by_request"
              and body.get("cancelRequested") is True
              and body.get("cancelHonored") is False
              and body.get("timeoutAlsoDue") is True
              and body.get("timeoutReason") == "timeout"
              and body.get("maxMs") == 2000
              and (body.get("elapsedMs") or 0) >= 1900,
              "cancelAcked=%s cancelApplied=%s reason=%s timeoutAlsoDue=%s maxMs=%s elapsed=%s"
              % (coob.get("acked"), coob.get("applied"), body.get("reason"),
                 body.get("timeoutAlsoDue"), body.get("maxMs"), body.get("elapsedMs")))
        CFG.max_call_ms = 120000
        CFG.max_call_hard_ms = 600000

        # ---- S37: OUT-OF-PROCESS read-only fallback (client section B) -------------------
        # A #32770 dialog + Button/Static children owned by THIS process make the measurement
        # deterministic without starting VB6 (red line) or touching another live app.
        fx = _make_oop_window_fixture()
        try:
            ok37 = _hwnd_i(fx.get("dialog")) != 0
            facts37, miss37 = oop_measure(os.getpid())
            check("S37a oop_measure sees the pid's top-level windows and its #32770 dialog "
                  "with button/static text (read-only, no input sent)",
                  ok37 and facts37.get("pidAlive") is True
                  and any(w.get("title") == "Microsoft Visual Basic"
                          for w in facts37.get("windows") or [])
                  and len(facts37.get("dialogs") or []) == 1
                  and any(b.get("text") == "End" for b in facts37["dialogs"][0]["buttons"])
                  and isinstance(facts37.get("hung"), bool),
                  "windows=%d dialogs=%d missing=%s" % (
                      len(facts37.get("windows") or []), len(facts37.get("dialogs") or []),
                      [m.get("why") for m in miss37]))

            b37 = oop_body("ide_state", {"__pid": os.getpid()}, "ctl_unavailable", {})
            check("S37b ide_state fallback: source=out_of_process + degraded=true + "
                  "addInReceipt=false, and only measured fields",
                  b37.get("source") == "out_of_process" and b37.get("degraded") is True
                  and b37.get("addInReceipt") is False
                  and b37.get("mode") is None and b37.get("windowCount") >= 1
                  and (b37.get("oob") or {}).get("lane") == "oop",
                  "source=%s windows=%s mode=%s" % (b37.get("source"),
                                                    b37.get("windowCount"), b37.get("mode")))

            b37 = oop_body("read_active_dialogs", {"__pid": os.getpid()}, "ctl_unavailable", {})
            d37 = (b37.get("dialogs") or [{}])[0]
            check("S37c read_active_dialogs fallback reports the #32770 title + buttons",
                  b37.get("count") == 1 and d37.get("class") == "#32770"
                  and d37.get("title") == "Microsoft Visual Basic"
                  and sorted(str(x.get("text")) for x in d37.get("buttons") or [])
                  == ["Debug", "End"],
                  "buttons=%s" % ([x.get("text") for x in d37.get("buttons") or []],))

            b37 = oop_body("find_gui", {"__pid": os.getpid(), "keyword": "Visual Basic"},
                           "ctl_unavailable", {})
            check("S37d find_gui fallback filters by keyword (title/class match only)",
                  b37.get("count") >= 1
                  and any(h.get("title") == "Microsoft Visual Basic"
                          for h in b37.get("hits") or []),
                  "hits=%d" % (b37.get("count"),))

            b37 = oop_body("ui_text", {"__pid": os.getpid(), "hwnd": fx["dialog"]},
                           "ctl_unavailable", {})
            check("S37e ui_text fallback returns the measured window text (WM_GETTEXT), "
                  "never an invented string",
                  "Run-time error" in str(b37.get("text") or "")
                  and b37.get("source") == "wm_gettext",
                  "textLen=%d" % len(str(b37.get("text") or "")))

            b37 = oop_body("last_error", {"__pid": os.getpid()}, "ctl_unavailable", {})
            check("S37f last_error fallback finds the runtime-error dialog and PARSES the "
                  "measured text (errNumber 91) with parseBasis stated",
                  (b37.get("errorDialog") or {}).get("present") is True
                  and b37.get("errNumber") == 91
                  and "Object variable" in str(b37.get("errDescription") or "")
                  and b37.get("parseBasis"),
                  "errNumber=%s desc=%r" % (b37.get("errNumber"),
                                            (b37.get("errDescription") or "")[:40]))

            # A ctl-less live instance is the ONLY state in which the fallback may fire, so the
            # gate is probed against exactly that state (and the same entry serves S37h).
            _reset_registry(regdir)
            n37 = uuid.uuid4().hex[:16]
            p37 = "\\\\.\\pipe\\vb6mcp-1-%d-oop-%s" % (mine, n37)
            s37 = StubPipeAddIn(p37, n37, mine, project="OopOne", tools=tools,
                                behaviour="silent").start()
            stubs.append(s37)
            _write_instance(regdir, mine, project="OopOne", pipe=p37, nonce=n37,
                            instance_id="%d-%s" % (mine, n37), transport="pipe")
            discover(force=True)

            gate37a = oop_fallback_body("ide_state", {}, None,
                                        {"peerState": "peer_busy", "ctlAvailable": False})
            gate37b = oop_fallback_body("ide_state", {}, None,
                                        {"peerState": "peer_unresponsive", "ctlAvailable": True})
            gate37c = oop_fallback_body("build", {}, None,
                                        {"peerState": "peer_unresponsive", "ctlAvailable": False})
            gate37d = oop_fallback_body("ide_state", {}, None,
                                        {"peerState": "peer_unresponsive", "reason": "timeout",
                                         "ctlAvailable": False})
            # a DENY-shaped failure is peer_unresponsive too, but the add-in SPOKE: measuring
            # windows would bury the denial itself (this is the S23 regression this gate prevents).
            gate37e = oop_fallback_body("ide_state", {}, None,
                                        {"peerState": "peer_unresponsive",
                                         "reason": "pipe_connect_failed", "ctlAvailable": False})
            check("S37g the fallback GATE: never for peer_busy, never over a LIVE ctl lane, "
                  "never for a non-listed tool, never for a DENY-shaped failure; yes for "
                  "ctl-unavailable + a WEDGE-shaped peer_unresponsive",
                  gate37a is None and gate37b is None and gate37c is None and gate37e is None
                  and gate37d is not None
                  and (gate37d.get("oob") or {}).get("why") == "ctl_unavailable",
                  "a=%s b=%s c=%s d=%s e=%s" % (gate37a is None, gate37b is None,
                                                 gate37c is None, gate37d is not None,
                                                 gate37e is None))

            # end to end: the same silent business pipe + NO ctl fields on the entry -> the
            # receipt that would have been a bare timeout becomes an out-of-process reading.
            CFG.max_call_ms = 1200
            CFG.max_call_hard_ms = 1200
            text, iserr, meta = handle_jsonrpc(_jreq("tools/call", {
                "name": "ide_state", "arguments": {}}))
            body = _resp_body(text)
            check("S37h END TO END: a timed-out ide_state on a ctl-less instance is answered "
                  "OUT OF PROCESS (degraded, why=ctl_unavailable) instead of a bare timeout",
                  iserr is True and body.get("source") == "out_of_process"
                  and body.get("degraded") is True
                  and (body.get("oob") or {}).get("why") == "ctl_unavailable"
                  and (body.get("failureUnderneath") or {}).get("reason") == "timeout"
                  and (meta or {}).get("source") == "out_of_process",
                  "source=%s why=%s under=%s" % (
                      body.get("source"), (body.get("oob") or {}).get("why"),
                      (body.get("failureUnderneath") or {}).get("reason")))

            # ---- S37i/S37j: the ADDITIVE second trigger + the DENY is never re-diagnosed ----
            # The failure fields must be IDENTICAL with the process-out-of-process machinery ON
            # and OFF.  A digest over the semantic fields is the comparison, not a human read.
            def _biz_digest(b):
                keys = ("reason", "reasonLegacy", "failureKind", "peerState",
                        "suggestedAction", "serverReason", "serverDetail", "retryAfterMs",
                        "ctlAvailable", "maxMsSource")
                snap = dict((k, (b or {}).get(k)) for k in keys)
                return hashlib.sha256(json.dumps(snap, sort_keys=True).encode("utf-8")).hexdigest()

            # S37i: peer_busy (a HOST-MODAL-shaped stall on the business pipe) + ctl lane absent
            # + process alive + a visible #32770 => the segment is APPENDED, nothing rewritten.
            _reset_registry(regdir)
            n37b = uuid.uuid4().hex[:16]
            p37b = "\\\\.\\pipe\\vb6mcp-1-%d-oopbusy-%s" % (mine, n37b)
            s37b = StubPipeAddIn(p37b, n37b, mine, project="OopBusy", tools=tools).start()
            stubs.append(s37b)
            _write_instance(regdir, mine, project="OopBusy", pipe=p37b, nonce=n37b,
                            instance_id="%d-%s" % (mine, n37b), transport="pipe")
            discover(force=True)
            occ37 = PipeClient(p37b).connect(600)
            occ37.send_frame({"vb6mcp": "hello", "proto": NP_PROTO, "nonce": n37b,
                              "client": {"pid": mine, "session": 1, "user": "selftest"}},
                             time.monotonic() + 1.0)
            occ37.recv_frame(time.monotonic() + 2.0)
            _drop_pipe_probe(discover()["live"][0])
            CFG.oop_enabled = False
            text, iserrA, metaA = handle_jsonrpc(_jreq("tools/call", {
                "name": "ide_state", "arguments": {}}))
            bodyA = _resp_body(text)
            CFG.oop_enabled = True
            text, iserrB, metaB = handle_jsonrpc(_jreq("tools/call", {
                "name": "ide_state", "arguments": {}}))
            bodyB = _resp_body(text)
            add37 = bodyB.get("outOfProcess") or {}
            check("S37i ADDITIVE second trigger: peer_busy + ctl lane absent + pid alive + a "
                  "visible #32770 -> outOfProcess APPENDED, failure fields byte-identical "
                  "(digest ON == digest OFF)",
                  iserrB is True and bodyB.get("peerState") == "peer_busy"
                  and bodyB.get("failureKind") == "peer_busy"
                  and bodyB.get("suggestedAction") == "retry_once"
                  and "outOfProcess" not in bodyA
                  and add37.get("source") == "out_of_process"
                  and add37.get("degraded") is True and add37.get("additive") is True
                  and (add37.get("evidence") or {}).get("modalDialogs") >= 1
                  and _biz_digest(bodyA) == _biz_digest(bodyB),
                  "reason=%s kind=%s add=%s dialogs=%s digestEqual=%s" % (
                      bodyB.get("reason"), bodyB.get("failureKind"), add37.get("source"),
                      (add37.get("evidence") or {}).get("modalDialogs"),
                      _biz_digest(bodyA) == _biz_digest(bodyB)))
            occ37.close()

            # S37j: a DENY (the add-in SPOKE) must keep its reason verbatim: the segment must
            # NOT appear even though the pid is alive, the ctl lane is absent and a #32770 is
            # visible - the exact state that fires S37i.
            _reset_registry(regdir)
            n37c = uuid.uuid4().hex[:16]
            p37c = "\\\\.\\pipe\\vb6mcp-1-%d-oopdeny-%s" % (mine, n37c)
            s37c = StubPipeAddIn(p37c, n37c, mine, project="OopDeny", tools=tools,
                                 behaviour="deny_hello", deny_reason="nonce_mismatch").start()
            stubs.append(s37c)
            _write_instance(regdir, mine, project="OopDeny", pipe=p37c, nonce=n37c,
                            instance_id="%d-%s" % (mine, n37c), transport="pipe")
            discover(force=True)
            CFG.oop_enabled = True
            _drop_pipe_probe(discover()["live"][0])
            text, iserrD, metaD = handle_jsonrpc(_jreq("tools/call", {
                "name": "ide_state", "arguments": {}}))
            bodyD = _resp_body(text)
            # A deny makes the add-in REBUILD its pipe, and the bridge caches that transient
            # busy for 2 s - so the comparison run needs a fresh probe and the rebuild window,
            # otherwise the second call legitimately reports pipe_unavailable instead.
            time.sleep(0.5)
            CFG.oop_enabled = False
            _drop_pipe_probe(discover()["live"][0])
            text, iserrE, _metaE = handle_jsonrpc(_jreq("tools/call", {
                "name": "ide_state", "arguments": {}}))
            bodyE = _resp_body(text)
            CFG.oop_enabled = True
            _live37 = discover()["live"]
            eligD = oop_addendum_eligible(bodyD, _live37[0] if _live37 else None)
            check("S37j a DENY is NEVER re-diagnosed: reason/serverReason/peerState byte-identical "
                  "(digest ON == digest OFF) and NO outOfProcess segment is appended",
                  bodyD.get("reason") == "pipe_connect_failed"
                  and bodyD.get("serverReason") == "nonce_mismatch"
                  and bodyD.get("peerState") == "peer_unresponsive"
                  and "outOfProcess" not in bodyD
                  and eligD is None
                  and _biz_digest(bodyD) == _biz_digest(bodyE),
                  "reason=%s serverReason=%s outOfProcess=%s eligible=%s digestEqual=%s" % (
                      bodyD.get("reason"), bodyD.get("serverReason"),
                      "outOfProcess" in bodyD, eligD,
                      _biz_digest(bodyD) == _biz_digest(bodyE)))
            CFG.max_call_ms = 120000
            CFG.max_call_hard_ms = 600000
            CFG.oop_enabled = True
        finally:
            if fx.get("dialog"):
                user32.DestroyWindow(fx["dialog"])

        # ---- S15: the DEPLOYABLE entry point, as its own OS process ----------------------
        live_dir = os.path.join(tmpdir, "instances_live")
        os.makedirs(live_dir, exist_ok=True)
        n15 = uuid.uuid4().hex[:16]
        p15 = "\\\\.\\pipe\\vb6mcp-1-child-stub-%s" % n15
        stub_env = _child_env({
            "VB6MCP_STUB_PIPE": p15, "VB6MCP_STUB_NONCE": n15,
            "VB6MCP_STUB_PROJECT": "LiveChild", "VB6MCP_STUB_REGISTRY_DIR": live_dir,
            "VB6MCP_STUB_TOOLS": os.path.join(REPO_ROOT, "docs", "tools_schema.json")})
        stub_proc = subprocess.Popen(
            [sys.executable, os.path.join(HERE, "vb6mcp_bridge.py"), "--stub-addin"],
            cwd=REPO_ROOT, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, env=stub_env)
        reg_file = None
        for _ in range(200):
            try:
                files = [f for f in os.listdir(live_dir) if f.endswith(".json")]
            except OSError:
                files = []
            if files:
                reg_file = os.path.join(live_dir, files[0])
                break
            time.sleep(0.05)
        port15 = _free_port()
        srv_env = _child_env({"VB6MCP_INSTANCES_DIR": live_dir})
        srv_proc = subprocess.Popen(
            [sys.executable, os.path.join(HERE, "vb6mcp_bridge.py"), "--port", str(port15)],
            cwd=REPO_ROOT, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, env=srv_env)
        online = False
        for _ in range(120):
            try:
                s = socket.create_connection(("127.0.0.1", port15), 0.2)
                s.close()
                online = True
                break
            except OSError:
                time.sleep(0.05)
        r15 = None
        if online:
            r15 = _run_tool(["tools/mcp_strict_audit.py",
                             "--url", "http://127.0.0.1:%d/mcp" % port15])
        check("S15 the deployable CLI (--port, own process) serves the pipe end to end",
              bool(r15) and r15.returncode == 0 and "strict_failures=0" in r15.stdout,
              "registry=%s online=%s rc=%s" % (os.path.basename(reg_file or "-"), online,
                                               r15.returncode if r15 else "-"))
        for proc in (srv_proc, stub_proc):
            try:
                proc.terminate()
            except Exception:
                pass
        for proc in (srv_proc, stub_proc):
            try:
                proc.wait(timeout=10)
            except Exception:
                try:
                    proc.kill()
                except Exception:
                    pass

        # ---- S16 (WP-22-02): a DEAD peer is NOT "unresponsive" ---------------------------
        # MCP-1021-43: after pipe_stalled the IDE process was gone; the old envelope still said
        # peer_unresponsive, so the caller was told to release a gate that no longer existed.
        dead_inst = Instance({"pid": 4294967294, "port": 0, "transport": "pipe",
                              "project": "DeadProbe", "startedAt": "2026-10-22 09:00:00"})
        live_inst = Instance({"pid": os.getpid(), "port": 0, "transport": "pipe",
                              "project": "LiveProbe", "startedAt": "2026-10-22 09:00:00"})
        dead_body = _transport_failure_body("pipe_stalled", "pipe error 109 GetOverlappedResult (io)",
                                            time.monotonic(), 120000, "run_project",
                                            "param_slack", dead_inst, {})
        live_body = _transport_failure_body("pipe_stalled", "pipe error 109 GetOverlappedResult (io)",
                                            time.monotonic(), 120000, "run_project",
                                            "param_slack", live_inst, {})
        check("S16 a DEAD pid upgrades peer_unresponsive -> peer_dead (+peerDead/lostContext/howToRecover)",
              dead_body.get("peerState") == "peer_dead"
              and dead_body.get("peerDead") is True
              and dead_body.get("failureKind") == "peer_dead"
              and (dead_body.get("lostContext") or {}).get("debugger") is True
              and any("ide_kill_and_relaunch" in x for x in (dead_body.get("howToRecover") or [])),
              "peerState=%s debugger=%s" % (dead_body.get("peerState"),
                                            (dead_body.get("lostContext") or {}).get("debugger")))
        check("S16 a LIVE pid keeps the ordinary peer_unresponsive envelope (no false peer_dead)",
              live_body.get("peerState") == "peer_unresponsive"
              and "peerDead" not in live_body
              and live_body.get("suggestedAction") in ("release_gate", "kill_and_restart"),
              "peerState=%s" % live_body.get("peerState"))

        ok = all(r[1] for r in results)
    finally:
        if httpd is not None:
            try:
                httpd.shutdown()
                httpd.server_close()
            except Exception:
                pass
        POOL.close_all()
        for st in stubs:
            st.stop()
        for hs in http_stubs:
            hs.stop()
        (CFG.instances_dir, CFG.transport, CFG.pipe_idle_ms, CFG.forward_timeout_ms,
         CFG.pipe_fallback, CFG.pipe_replay, CFG.pipe_connect_ms,
         CFG.target, _saved_scan) = saved
        if _saved_scan is None:
            os.environ.pop("VB6MCP_INSTANCES_SCAN", None)
        else:
            os.environ["VB6MCP_INSTANCES_SCAN"] = _saved_scan
        shutil.rmtree(tmpdir, ignore_errors=True)
        discover(force=True)

    failed = [r for r in results if not r[1]]
    print("=" * 100)
    print("RESULT %d/%d checks passed%s" % (len(results) - len(failed), len(results),
                                            "" if not failed else
                                            "  FAILED: " + ", ".join(r[0] for r in failed)))
    print("=" * 100)
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())



