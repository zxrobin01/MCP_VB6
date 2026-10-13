"""Strict-schema audit of a Streamable HTTP MCP endpoint.

Derived from strict-client post-mortems: the official @modelcontextprotocol/sdk
client validates responses with STRICT zod schemas, so extra/malformed keys make it
drop the response and register NO tools (configured but does not take effect).

Checks:
  1. no top-level _meta on the JSON-RPC envelope
  2. initialize.protocolVersion is echoed from the client when supported
  3. tools/list result carries no nextCursor=null / ttlMs / cacheScope
  4. tools/call returns a content[] array

Usage: python mcp_strict_audit.py --url <endpoint>
       python mcp_strict_audit.py <endpoint>            # legacy positional form

THE TARGET HAS NO DEFAULT, AND IT IS NEVER GUESSED.  This audit performs a
tools/call, i.e. it DRIVES whichever IDE owns the endpoint.  Port 8766 is RESERVED
and is always refused here, and only the self-owned port range is accepted, so:
  * no --url and no positional URL -> usage error, exit 2, NOTHING is sent
  * port 8766                      -> refused, exit 3, NOTHING is sent
  * any port below 8767            -> refused, exit 3, NOTHING is sent
  * port 8767 or higher            -> accepted; if it cannot be reached the
                                      script reports CONNECT_FAILED (exit 4) and
                                      NEVER falls back to another port
The effective target is always printed as the first line (URL=...).

Order in the toolchain regression: dump_tools_schema -> export_tool_catalog ->
check_schema_baseline -> check_live_tools -> mcp_strict_audit (this).  The script
never writes a file and never rewrites the FROZEN tools/schema_baseline.json.

Exit codes: 0 all checks passed | 1 strict check(s) failed | 2 usage error |
            3 refused target | 4 connection failure
"""
import json
import sys
import urllib.error
import urllib.parse
import urllib.request

CLIENT_SUPPORTED = ["2025-11-25", "2025-06-18", "2025-03-26", "2024-11-05", "2024-10-07"]
STRICT_ONLY_KEYS = ["nextCursor", "ttlMs", "cacheScope"]
FORBIDDEN_PORT = 8766
MIN_SELF_OWNED_PORT = 8767
HELP_FLAGS = ("-h", "--help")

fails = []


def parse_target(argv):
    """(url, error).  --url <u> | --url=<u> | a bare positional URL; no default."""
    url = None
    i = 0
    while i < len(argv):
        a = argv[i]
        if a in HELP_FLAGS:
            return "HELP", None
        if a == "--url":
            if i + 1 >= len(argv):
                return None, "--url needs a value"
            url = argv[i + 1]
            i += 2
            continue
        if a.startswith("--url="):
            url = a.split("=", 1)[1]
            i += 1
            continue
        if a.startswith("-"):
            return None, "unknown option: %s" % a
        if url is None:
            url = a
            i += 1
            continue
        return None, "unexpected extra argument: %s" % a
    if url is None:
        return None, "no target given"
    return url, None


def target_port(url):
    """(port, "") or (None, reason).  Only http(s) URLs are accepted."""
    parts = urllib.parse.urlsplit(url)
    if parts.scheme not in ("http", "https") or not parts.hostname:
        return None, "not an http(s) URL"
    return (parts.port or (443 if parts.scheme == "https" else 80)), ""


def post(url, method, params=None, rid=1, proto=None):
    body = {"jsonrpc": "2.0", "id": rid, "method": method}
    if params is not None:
        body["params"] = params
    headers = {"Content-Type": "application/json",
               "Accept": "application/json, text/event-stream"}
    if proto:
        headers["MCP-Protocol-Version"] = proto
    req = urllib.request.Request(url, data=json.dumps(body).encode("utf-8"), headers=headers)
    with urllib.request.urlopen(req, timeout=60) as r:
        raw = r.read().decode("utf-8-sig")
    for line in raw.splitlines():
        line = line.strip()
        if line.startswith("data:"):
            line = line[5:].strip()
        if line.startswith("{"):
            return json.loads(line)
    return json.loads(raw) if raw.strip() else None


def check(name, ok, detail=""):
    print("%-46s %s %s" % (name, "PASS" if ok else "FAIL", detail))
    if not ok:
        fails.append(name)


def usage_error(reason):
    sys.stderr.write("mcp_strict_audit: %s\n" % reason)
    sys.stderr.write("  NOTHING WAS SENT.  This audit has no default target because it\n")
    sys.stderr.write("  drives whichever IDE owns the endpoint, and port 8766 is reserved.\n")
    sys.stderr.write("  Usage: python tools/mcp_strict_audit.py --url http://127.0.0.1:<your own port>/mcp\n")


def main():
    url, err = parse_target(sys.argv[1:])
    if url == "HELP":
        print(__doc__)
        return 0
    if err:
        usage_error(err)
        return 2
    port, why = target_port(url)
    if port is None:
        usage_error("target %r is not usable (%s)" % (url, why))
        return 3
    if port == FORBIDDEN_PORT:
        usage_error("REFUSED: port %d is RESERVED and is never touched by this script" % FORBIDDEN_PORT)
        return 3
    if port < MIN_SELF_OWNED_PORT:
        usage_error("REFUSED: port %d is below the self-owned range (>= %d)" % (port, MIN_SELF_OWNED_PORT))
        return 3
    print("URL=" + url)
    try:
        _audit(url)
    except (urllib.error.URLError, OSError) as exc:
        sys.stderr.write("CONNECT_FAILED %s -> %s: %s\n" % (url, type(exc).__name__, exc))
        sys.stderr.write("  no fallback target is tried (this audit has no default).\n")
        return 4
    print("SUMMARY strict_failures=%d" % len(fails))
    print("EXIT=" + ("0" if not fails else "1"))
    return 0 if not fails else 1


def _audit(url):
    init = post(url, "initialize", {"protocolVersion": "2025-06-18",
                                   "capabilities": {},
                                   "clientInfo": {"name": "strict-audit", "version": "1"}})
    check("1. no top-level _meta on envelope", "_meta" not in init,
          "keys=" + ",".join(init.keys()))
    sv = (init.get("result") or {}).get("protocolVersion")
    check("2. protocolVersion supported by client", sv in CLIENT_SUPPORTED, "server=%r" % sv)
    # Every official version must be echoed back, otherwise a client that only supports
    # a newer version gets a reply it does not recognise and registers no tools.
    for v in CLIENT_SUPPORTED:
        got = (post(url, "initialize", {"protocolVersion": v, "capabilities": {},
                                       "clientInfo": {"name": "strict-audit", "version": "1"}},
                    rid=10, proto=None).get("result") or {}).get("protocolVersion")
        check("2b. echoes client %s" % v, got == v, "server=%r" % got)

    tl = post(url, "tools/list", {}, 2, proto=sv)
    tres = tl.get("result") or {}
    extra = [k for k in STRICT_ONLY_KEYS if k in tres]
    check("3. tools/list has no strict-only keys", not extra, "extra=" + ",".join(extra))
    check("3b. tools/list result keys == {tools}", set(tres.keys()) <= {"tools"},
          "keys=" + ",".join(tres.keys()))
    n = len(tres.get("tools") or [])
    check("3c. tools present", n > 0, "count=%d" % n)
    check("3d. every tool has inputSchema",
          all(t.get("inputSchema") for t in (tres.get("tools") or [])), "")

    name = (tres.get("tools") or [{}])[0].get("name")
    tc = post(url, "tools/call", {"name": name, "arguments": {}}, 3, proto=sv)
    cres = tc.get("result") or {}
    check("4. tools/call returns content[]", isinstance(cres.get("content"), list),
          "tool=%s keys=%s" % (name, ",".join(cres.keys())))


raise SystemExit(main())
