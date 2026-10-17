#!/usr/bin/env python3
"""
mcp_stdio_bridge.py - stdio <-> HTTP bridge for the VB6MCP add-in.

MCP clients speak JSON-RPC over stdio; the VB6MCP add-in exposes the same
JSON-RPC over HTTP on 127.0.0.1. This bridge proxies between the two using
only the Python standard library.

LEGACY / ROLLBACK ONLY.  This is the PRE-bridge stdio entry point: it talks to
the add-in's own HTTP port directly (default 8766) and has NO token support.  The current
entry point is the Node MCP shell

    node bridge/vb6mcp_hub.mjs

which talks to the middle service bridge/vb6mcp_bridge.py (ONE local TCP port for every IDE
on the machine, bearer-token auth via VB6MCP_TOKEN / VB6MCP_TOKEN_FILE).  Keep this file to
roll back a deployment that has not moved to the middle service yet.

Usage:
    python mcp_stdio_bridge.py [--port 8766] [--host 127.0.0.1]

Port resolution order: --port, VB6MCP_PORT env var, %TEMP%\\vb6mcp_port.txt
(written by the add-in at startup), default 8766.
"""
import argparse
import http.client
import json
import os
import sys
import tempfile


def port_from_file():
    try:
        p = os.path.join(tempfile.gettempdir(), "vb6mcp_port.txt")
        with open(p, "r", encoding="utf-8", errors="replace") as f:
            text = f.read().strip()
        return int(text) if text else None
    except (OSError, ValueError):
        return None


def resolve_port(cli_port):
    if cli_port:
        return cli_port
    env = os.environ.get("VB6MCP_PORT")
    if env:
        try:
            return int(env)
        except ValueError:
            pass
    return port_from_file() or 8766


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--port", type=int, default=None)
    ap.add_argument("--host", default="127.0.0.1")
    args = ap.parse_args()

    port = resolve_port(args.port)
    conn = None

    def send(payload):
        nonlocal conn
        body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        last_error = None
        for _attempt in range(2):
            try:
                if conn is None:
                    conn = http.client.HTTPConnection(args.host, port, timeout=60)
                conn.request(
                    "POST", "/rpc", body=body,
                    headers={"Content-Type": "application/json"})
                resp = conn.getresponse()
                data = resp.read()
                if resp.status in (200, 202):
                    return data, None
                try:
                    conn.close()
                except Exception:
                    pass
                conn = None
                last_error = "HTTP %s" % resp.status
            except OSError as exc:
                try:
                    conn.close()
                except Exception:
                    pass
                conn = None
                last_error = str(exc)
        return None, last_error

    out = sys.stdout.buffer
    for raw in sys.stdin.buffer:
        line = raw.strip()
        if not line:
            continue
        req_id = None
        try:
            req = json.loads(line.decode("utf-8"))
            if isinstance(req, dict):
                req_id = req.get("id")
        except ValueError:
            req = None
        if req is None:
            out.write(
                b'{"jsonrpc":"2.0","id":null,"error":'
                b'{"code":-32700,"message":"Parse error"}}\n')
            out.flush()
            continue
        data, err = send(req)
        if err is not None:
            if req_id is None:
                continue  # notifications get no response
            payload = {
                "jsonrpc": "2.0",
                "id": req_id,
                "error": {
                    "code": -32000,
                    "message": ("VB6 MCP server unreachable (%s). Is the VB6 "
                                "IDE running with the VB6MCP add-in loaded?") % err,
                },
            }
            data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        if data:
            out.write(data + b"\n")
            out.flush()


if __name__ == "__main__":
    main()
