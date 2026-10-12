#!/usr/bin/env python3
"""Quick smoke test for the VB6MCP HTTP endpoint (no stdio involved).

Run while the VB6 IDE is open with the VB6MCP add-in loaded:

    python tests/smoke_test.py [--port 8766]
"""
import argparse
import json
import urllib.request


def rpc(port, payload):
    req = urllib.request.Request(
        "http://127.0.0.1:%d/rpc" % port,
        data=json.dumps(payload).encode("utf-8"),
        headers={"Content-Type": "application/json"},
    )
    with urllib.request.urlopen(req, timeout=60) as resp:
        return json.loads(resp.read().decode("utf-8"))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--port", type=int, default=8766)
    args = ap.parse_args()

    init = rpc(args.port, {
        "jsonrpc": "2.0", "id": 1, "method": "initialize",
        "params": {"protocolVersion": "2024-11-05", "capabilities": {},
                   "clientInfo": {"name": "smoke-test", "version": "0.0.1"}},
    })
    print("initialize ->", json.dumps(init["result"]["serverInfo"],
                                      ensure_ascii=False))
    print("capabilities:", sorted(init["result"]["capabilities"].keys()))

    tools = rpc(args.port, {"jsonrpc": "2.0", "id": 2,
                            "method": "tools/list", "params": {}})
    names = [t["name"] for t in tools["result"]["tools"]]
    print("tools (%d): %s" % (len(names), ", ".join(names)))

    info = rpc(args.port, {"jsonrpc": "2.0", "id": 3,
                           "method": "tools/call",
                           "params": {"name": "get_project_info",
                                      "arguments": {}}})
    item = info["result"]["content"][0]
    print("get_project_info isError=%s" % info["result"].get("isError"))
    print(item["text"])


if __name__ == "__main__":
    main()
