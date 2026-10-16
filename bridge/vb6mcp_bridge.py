#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""vb6mcp_bridge.py - the VB6MCP middle service layer.

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

NAMED PIPE WIRE PROTOCOL (np/1) - fixed by the P1-065 spike, agreed with the VB side

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

AUTHENTICATION (task-71; threat model = bridge/TOKEN_THREAT_MODEL.md)

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

    BINDING IS LOOPBACK-ONLY, UNCONDITIONALLY (threat model M-0): a non-loopback host is
    refused with exit 6 and there is deliberately NO switch to override it.  A bearer token
    authenticates an ENDPOINT, not a user's intent: a client that legitimately holds the token
    is outside this control.

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
import ctypes
import hashlib
import hmac
import http.server
import json
import os
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
        self.progress_ms = _env_int("VB6MCP_PIPE_PROGRESS_MS", 2000)
        # NB: deliberately NOT "VB6MCP_PIPE_IDLE_MS" - the VB side uses that name for the
        # SERVER's 120 s idle close.  Sharing the name would turn a client stall threshold of
        # 6 s into 120 s by accident (a silent loss of the fast escape), so the client's
        # threshold is bridge-namespaced.
        self.pipe_idle_ms = _env_int("VB6MCP_BRIDGE_PIPE_IDLE_MS", 0) or (3 * self.progress_ms)
        self.pipe_fallback = _env_bool("VB6MCP_PIPE_FALLBACK", True)
        self.pipe_replay = _env_bool("VB6MCP_PIPE_REPLAY", False)
        self.debug = _env_bool("VB6MCP_HUB_DEBUG", False)
        #: an APPEND-ONLY diagnostic log.  Without it a failed joint run leaves a 0-byte file
        #: and no way to tell "the IDE was invisible" from "the handshake was rejected".
        self.log_path = _env_str("VB6MCP_BRIDGE_LOG", "")
        self.target = _env_str("VB6MCP_TARGET", "") or ""
        self.instances_dir = _env_str("VB6MCP_INSTANCES_DIR", "") or os.path.join(
            tempfile.gettempdir(), "vb6mcp", "instances")
        self.discovery_ttl_ms = _env_int("VB6MCP_DISCOVERY_TTL_MS", 2000)
        # ---- authentication (task-71; threat model = bridge/TOKEN_THREAT_MODEL.md) -----------
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
    unrelated process must not read as 'our IDE' (TASK-P3-034), and this is only ever used
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
    same pipe within ~250 ms (task-57 measurement).  Retrying costs nothing when the pipe is
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
            "jsonEncoding": self.jsonEncoding,
        }


def read_registry():
    """Tolerant read of %TEMP%\\vb6mcp\\instances\\*.json (sec 4.2.2: never drop an entry)."""
    d = CFG.instances_dir
    entries, stale, roots_seen, candidates = [], [], [], []
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
        entry = {
            "file": full,
            "pid": _as_int(j.get("pid")),
            "port": _as_int(j.get("port")),
            "project": j.get("project") or j.get("projectName"),
            "projectPath": j.get("projectPath") or j.get("path"),
            "startedAt": j.get("startedAt"),
            "version": j.get("version"),
            "dllSha256": j.get("dllSha256") or j.get("sha256"),
            "startedByHub": bool(j.get("startedByHub")),
            "transport": j.get("transport"),
            "pipe": j.get("pipe"),
            "nonce": j.get("nonce"),
            "sessionId": _as_int(j.get("sessionId")),
            "instanceId": j.get("instanceId"),
            "tcpEnabled": j.get("tcpEnabled"),
            "jsonEncoding": enc,
            "registryRoot": root,
        }
        inst = Instance(entry)
        if any(e.pid == inst.pid and e.key == inst.key for e in entries):
            continue                    # the same instance registered under two roots
        entries.append(inst)
        if not inst.pidAlive:
            stale.append({"file": full, "pid": inst.pid, "reason": "pid_not_alive"})
        elif inst.pipeState == "mismatch":
            stale.append({"file": full, "pid": inst.pid, "reason": "pipe_identity_mismatch"})
    return {"dir": d, "unreadable": False, "instances": entries, "stale": stale,
            "roots": roots_seen}


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
            "live": [i for i in raw["instances"] if i.pidAlive and i.has_route],
            "at": now,
        }
        _DISC_CACHE["at"] = now
        _DISC_CACHE["data"] = data
        dbg("discover: entries=%d live=%d stale=%d" % (len(raw["instances"]), len(data["live"]),
                                                       len(raw["stale"])))
        bridge_log("discover", configuredDir=raw.get("dir"), entries=len(raw["instances"]),
                   live=len(data["live"]),
                   roots=";".join("%s:%d" % (r["root"], r["entries"]) for r in (raw.get("roots") or [])),
                   staleReasons=",".join(sorted(set(s.get("reason", "?") for s in raw["stale"]))) or "-",
                   encodings=",".join(sorted(set(str(i.jsonEncoding) for i in raw["instances"]))) or "-")
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
    """The pipe name without the \\.\pipe\ prefix, for compact diagnostic lines."""
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
    "(the add-in ANSWERED with a structured error).  Additive; legacy reason values are unchanged.")


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


def pipe_roundtrip(inst, req_text, req_id, budget_ms):
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
                # reason/detail are lifted onto the receipt here as well (item 2 of task-64:
                # a deny must never collapse into an anonymous failure).
                try:
                    deny = json.loads((exc.partial or b"").decode("utf-8"))
                    if isinstance(deny, dict) and deny.get("vb6mcp") == "deny":
                        flags["serverReason"] = deny.get("reason") or "unspecified"
                        if deny.get("detail"):
                            flags["serverDetail"] = deny.get("detail")
                        flags["serverDenyRaw"] = json.dumps(deny, ensure_ascii=False)
                except Exception:
                    pass
                raise TransportError("pipe_connect_failed", str(exc), exc.code,
                                     exc.stage or "connect", wrote=False, flags=flags)
        client = sess.client
        payload = req_text.encode("utf-8")
        try:
            framed = frame_wrap(payload)
        except FrameError as exc:
            raise TransportError(exc.reason, str(exc), 0, "frame", wrote=False)
        ok, _sent = client.write_all(framed, deadline)
        if not ok:
            POOL.drop(inst)
            raise TransportError("pipe_write_failed", "request not fully written", 0, "write",
                                 wrote=False)
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


def forward_raw(req_text, req_obj, selector):
    """Forward one JSON-RPC message.  Returns a dict:
       {text|None, transport, transportReason, flags{}, elapsedMs, ide{}, failure{}}"""
    inst, refusal = select_target(selector)
    if inst is None:
        return {"text": None, "failure": refusal, "elapsedMs": 0}
    ide = {"pid": inst.pid, "project": inst.project, "port": inst.port}
    t0 = time.monotonic()
    try:
        track, reason, flags = choose_track(inst)
    except TransportError as exc:
        body = _transport_failure_body(exc.reason, exc.message, t0, CFG.forward_timeout_ms)
        body.update(exc.flags)
        body["failureKind"] = "noIde" if exc.reason == "pipe_unavailable" else "wedged"
        return {"text": None, "failure": body, "elapsedMs": _ms(t0), "ide": ide}
    req_id = req_obj.get("id") if isinstance(req_obj, dict) else None
    bridge_log("track", tool=_tool_of(req_obj), pid=inst.pid, port=inst.port,
               pipe=_pipe_leaf(inst.pipe), track=track, reason=reason,
               flags=",".join(sorted(flags)) or "-")
    if track == "pipe":
        try:
            text, saw_progress = pipe_roundtrip(inst, req_text, req_id, CFG.forward_timeout_ms)
            return {"text": text, "transport": "pipe", "transportReason": reason,
                    "flags": dict(flags), "elapsedMs": _ms(t0), "ide": ide,
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
                    text = tcp_roundtrip(inst, req_text, CFG.forward_timeout_ms)
                    return {"text": text, "transport": "tcp", "transportReason": reason,
                            "flags": flags, "elapsedMs": _ms(t0), "ide": ide,
                            "fellBackFrom": exc.reason}
                except TransportError as exc2:
                    body = _transport_failure_body(exc2.reason, exc2.message, t0,
                                                   CFG.forward_timeout_ms)
                    body.update(flags)
                    body["transport"] = "tcp"
                    body["transportReason"] = reason
                    return {"text": None, "failure": body, "elapsedMs": _ms(t0), "ide": ide}
            reason2 = "auto_pipe_busy" if flags.get("pipeBusy") else (
                "auto_pipe_error" if flags.get("pipeIdentityMismatch") else "auto_pipe_error")
            body = _transport_failure_body(exc.reason, exc.message, t0, CFG.forward_timeout_ms)
            body.update(flags)
            body["transport"] = "pipe"
            body["transportReason"] = reason2
            body["failureKind"] = "wedged"
            if exc.wrote and not CFG.pipe_replay:
                body["noReplay"] = True
                body["hint"] = ("the request WAS fully written to the pipe, so it is never "
                                "silently replayed (that could double-apply a write tool); "
                                "set VB6MCP_PIPE_REPLAY=1 to opt in")
            return {"text": None, "failure": body, "elapsedMs": _ms(t0), "ide": ide}
    try:
        text = tcp_roundtrip(inst, req_text, CFG.forward_timeout_ms)
        return {"text": text, "transport": "tcp", "transportReason": reason,
                "flags": dict(flags), "elapsedMs": _ms(t0), "ide": ide}
    except TransportError as exc:
        body = _transport_failure_body(exc.reason, exc.message, t0, CFG.forward_timeout_ms)
        body["transport"] = "tcp"
        body["transportReason"] = reason
        body["failureKind"] = "wedged" if exc.reason == "forward_timeout" else "noIde"
        body.update(flags)
        return {"text": None, "failure": body, "elapsedMs": _ms(t0), "ide": ide}


def _tool_of(req_obj):
    try:
        return (req_obj.get("params") or {}).get("name") or req_obj.get("method")
    except Exception:
        return None


def _ms(t0):
    return int((time.monotonic() - t0) * 1000)


def _transport_failure_body(reason, message, t0, budget_ms):
    m = message or ""
    kind = "wedged" if reason in ("forward_timeout", "pipe_stalled", "pipe_read_failed") else "noIde"
    return {
        "ok": False,
        "reason": reason,
        "message": m,
        "elapsedMs": _ms(t0),
        "budgetMs": budget_ms,
        "failureKind": kind,
        "failureKindRule": FAILURE_KIND_RULE,
    }


# ---------------------------------------------------------------- receipts --------

def contract_version():
    snap = read_json(SNAP_PATH, None)
    if isinstance(snap, dict) and snap.get("contractVersion"):
        return snap["contractVersion"]
    return "unknown"


def stamp_body(result, ide, contract, note):
    """The hub's stamping contract, restated: the JSON-RPC result keys stay EXACTLY
    {content,isError}; mcp/ide/transport markers live in the BODY TEXT only (ADR sec 2.5)."""
    content = []
    for c in (result.get("content") or []):
        content.append(dict(c))
    text_item = None
    for c in content:
        if c.get("type") == "text" and isinstance(c.get("text"), str):
            text_item = c
            break
    if text_item is not None:
        obj = None
        try:
            parsed = json.loads(text_item["text"])
            if isinstance(parsed, dict):
                obj = parsed
        except Exception:
            obj = None
        if obj is not None:
            text_item["text"] = _body_of(obj, ide, contract, note)
        else:
            text_item["text"] = text_item["text"] + "\n" + _body_of({}, ide, contract, note)
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


def handle_jsonrpc(req_text, selector=None):
    """Serve ONE JSON-RPC message.  Returns (response_text|None, is_error, meta).

    Contract face rules that must not drift:
      * tools/list result keys are EXACTLY {tools}   (mcp_strict_audit check 3b)
      * the JSON-RPC envelope carries NO _meta / nextCursor / ttlMs / cacheScope
      * a notification gets NO response body (HTTP 202, stdio none)
    """
    sel0 = selector if selector is not None else CFG.target
    try:
        req = json.loads(req_text)
    except Exception:
        return _jsonrpc_error(None, -32700, "Parse error"), True, {"reason": "parse_error"}
    if not isinstance(req, dict):
        return _jsonrpc_error(None, -32600, "Invalid Request"), True, {"reason": "not_object"}
    method = req.get("method")
    rid = req.get("id", None)
    if not isinstance(method, str) or not method:
        if rid is None:
            return None, False, {"notification": True}
        return _jsonrpc_error(rid, -32600, "Invalid Request"), True, {"reason": "no_method"}

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
        req2_text = json.dumps(req2, ensure_ascii=False)
        out = forward_raw(req2_text, req2, sel)
        if out.get("text") is None:
            body = dict(out.get("failure") or {"ok": False, "reason": "forward_failed"})
            note = _transport_note(out)
            for k, v in note.items():
                if k not in body:
                    body[k] = v
            if out.get("ide"):
                body["ide"] = out["ide"]
            return _jsonrpc_result(rid, error_result(body, True)), True, out
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
        # Lead ruling (task-71 item 3a): the weak configuration must be visible HERE too, not
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

    def _send(self, status, ctype, body, extra=None, content_length=True):
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
    # -- authentication (task-71 / TOKEN_THREAT_MODEL.md M-1..M-4) ---------------------------
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
        if not token_matches(provided, CFG.token):
            return ("token_mismatch", "the supplied token does not match")
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

    def do_OPTIONS(self):
        # CORS preflight carries no Authorization header by design, so it is never gated.
        self._send(204, "text/plain", b"")

    def do_GET(self):
        problem = self._auth_problem()
        if problem is not None:
            self._reject_auth(*problem)
            return
        path = self.path.split("?", 1)[0]
        if path in ("/health", "/ping"):
            self._send(200, "application/json", json.dumps(health_body(), ensure_ascii=False, indent=2))
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
        body, err = self._read_body()
        if err is not None:
            # the request body was NOT drained: never keep this connection alive
            self.close_connection = True
            self._send(err[0], err[1], err[2], {"Connection": "close"})
            return
        req_text = body.decode("utf-8-sig", "replace")
        target = self.headers.get("X-VB6MCP-Target") or CFG.target
        try:
            resp_text, _is_error, _meta = handle_jsonrpc(req_text, target)
        except Exception as exc:
            dbg("dispatch failed:", repr(exc))
            resp_text = _jsonrpc_error(None, -32603, "Internal error: %s" % exc)
        if resp_text is None:
            self._send(202, "application/json", b"")
            return
        accept = (self.headers.get("Accept") or "").lower()
        if "text/event-stream" in accept and "application/json" not in accept:
            sse = "event: message\r\ndata: " + resp_text + "\r\n\r\n"
            self._send(200, "text/event-stream", sse.encode("utf-8"))
            return
        extra = {}
        if '"method":"initialize"' in req_text.replace(" ", ""):
            extra["Mcp-Session-Id"] = str(uuid.uuid4())
        self._send(200, "application/json", resp_text.encode("utf-8"), extra)


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

def serve(port, host, selector):
    if not is_loopback_host(host):
        # M-0, UNCONDITIONAL: no override switch exists on purpose.
        sys.stderr.write(
            "vb6mcp_bridge: REFUSED - %r is not a loopback address.\n"
            "  This service exposes the whole tool face (110 tools, including WRITE tools that\n"
            "  edit files and run code in the IDE).  Non-loopback binding is disabled outright\n"
            "  (ADR sec 0.5 / TOKEN_THREAT_MODEL.md M-0) and there is deliberately NO switch to\n"
            "  override it.  Use 127.0.0.1, ::1 or localhost.\n" % (host,))
        bridge_log("bind_refused", host=host, reason="non_loopback_host")
        return 6
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
    sys.stderr.write("vb6mcp_bridge %s listening on http://%s:%d  (/rpc, /mcp, /health)\n"
                     % (VERSION, host, port))
    # ---- auth state, to stderr AND the log (Lead ruling task-71 item 3b) -------------------
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
    ap.add_argument("--stub-addin", action="store_true",
                    help="DEV ONLY: run a stub np/1 pipe add-in in its own process "
                         "(used by --selftest; env VB6MCP_STUB_PIPE/NONCE/...)")
    args = ap.parse_args(argv)

    CFG.resolve_token()
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
            "loopbackOnly": True,
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
        if kernel32.PeekNamedPipe(h, None, 0, None, ctypes.byref(avail), None) and avail.value:
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
              iserr is True and body.get("noReplay") is True and body.get("failureKind") == "wedged",
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
        CFG.forward_timeout_ms = 1200
        t0 = time.monotonic()
        text, iserr, meta = handle_jsonrpc(_jreq("tools/call", {"name": "build", "arguments": {}}))
        body = _resp_body(text)
        elapsed = (time.monotonic() - t0) * 1000.0
        check("S8 no frame at all -> forward_timeout at FORWARD_TIMEOUT_MS (never tighter than tcp)",
              body.get("reason") == "forward_timeout" and 1100 <= elapsed < 4000,
              "reason=%s elapsedMs=%.0f budgetMs=%s" % (body.get("reason"), elapsed,
                                                        body.get("budgetMs")))
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
        # task-57 measured that welcome.project is the LIVE ActiveVBProject.Name at handshake
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

        # ---- S27: the NEW writer (task-68) escapes non-ASCII -> the file is pure ASCII ------
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
        check("V7a only literal loopback hosts are accepted (0.0.0.0 impossible)",
              all(is_loopback_host(h) for h in ("127.0.0.1", "127.9.9.9", "localhost", "::1"))
              and not any(is_loopback_host(h) for h in ("0.0.0.0", "10.10.10.88", "192.168.1.5",
                                                        "example.com", "127.0.0.1.example.com")),
              "whitelist ok")
        check("V7b serve() refuses a non-loopback host with exit 6 (no override switch)",
              serve(8799, "0.0.0.0", "") == 6 and serve(8799, "10.10.10.88", "") == 6,
              "refused both")
        httpd28.shutdown()
        httpd28.server_close()
        CFG.token, CFG.token_source, CFG.log_path = "", "none", ""

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



