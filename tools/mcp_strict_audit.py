"""Strict-schema audit of a Streamable HTTP MCP endpoint.

Derived from the dynsim MCP post-mortems: the official @modelcontextprotocol/sdk
client validates responses with STRICT zod schemas, so extra/malformed keys make it
drop the response and register NO tools ("configured but does not take effect").

Checks:
  1. no top-level _meta on the JSON-RPC envelope
  2. initialize.protocolVersion is echoed from the client when supported
  3. tools/list result carries no nextCursor=null / ttlMs / cacheScope
  4. tools/call returns a content[] array

Usage: python mcp_strict_audit.py [url]

Order in the toolchain regression: dump_tools_schema -> export_tool_catalog ->
check_schema_baseline -> check_live_tools -> mcp_strict_audit (this).  Read-only;
never rewrites the FROZEN tools/schema_baseline.json.
"""
import json
import sys
import urllib.request

URL = sys.argv[1] if len(sys.argv) > 1 else "http://127.0.0.1:8766/mcp"
CLIENT_SUPPORTED = ["2025-11-25", "2025-06-18", "2025-03-26", "2024-11-05", "2024-10-07"]
STRICT_ONLY_KEYS = ["nextCursor", "ttlMs", "cacheScope"]


def post(method, params=None, rid=1, proto=None):
    body = {"jsonrpc": "2.0", "id": rid, "method": method}
    if params is not None:
        body["params"] = params
    headers = {"Content-Type": "application/json",
               "Accept": "application/json, text/event-stream"}
    if proto:
        headers["MCP-Protocol-Version"] = proto
    req = urllib.request.Request(URL, data=json.dumps(body).encode("utf-8"), headers=headers)
    with urllib.request.urlopen(req, timeout=60) as r:
        raw = r.read().decode("utf-8-sig")
    for line in raw.splitlines():
        line = line.strip()
        if line.startswith("data:"):
            line = line[5:].strip()
        if line.startswith("{"):
            return json.loads(line)
    return json.loads(raw) if raw.strip() else None


fails = []


def check(name, ok, detail=""):
    print("%-46s %s %s" % (name, "PASS" if ok else "FAIL", detail))
    if not ok:
        fails.append(name)


print("URL=" + URL)
init = post("initialize", {"protocolVersion": "2025-06-18",
                           "capabilities": {},
                           "clientInfo": {"name": "strict-audit", "version": "1"}})
check("1. no top-level _meta on envelope", "_meta" not in init,
      "keys=" + ",".join(init.keys()))
sv = (init.get("result") or {}).get("protocolVersion")
check("2. protocolVersion supported by client", sv in CLIENT_SUPPORTED, "server=%r" % sv)
# Every official version must be echoed back, otherwise a client that only supports
# a newer version gets a reply it does not recognise and registers no tools.
for v in CLIENT_SUPPORTED:
    got = (post("initialize", {"protocolVersion": v, "capabilities": {},
                             "clientInfo": {"name": "strict-audit", "version": "1"}},
                rid=10, proto=None).get("result") or {}).get("protocolVersion")
    check("2b. echoes client %s" % v, got == v, "server=%r" % got)

tl = post("tools/list", {}, 2, proto=sv)
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
tc = post("tools/call", {"name": name, "arguments": {}}, 3, proto=sv)
cres = tc.get("result") or {}
check("4. tools/call returns content[]", isinstance(cres.get("content"), list),
      "tool=%s keys=%s" % (name, ",".join(cres.keys())))

print("SUMMARY strict_failures=%d" % len(fails))
print("EXIT=" + ("0" if not fails else "1"))
sys.exit(0 if not fails else 1)
