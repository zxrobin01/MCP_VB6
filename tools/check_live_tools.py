"""Independent cross-check: a captured tools/list snapshot vs the frozen baseline.

This is the authoritative drift detector (argv snapshot in, ADDED/REMOVED/CHANGED +
ORDER_MATCH out); check_schema_baseline.py compares the two repo files instead.

Order: dump_tools_schema -> export_tool_catalog -> check_schema_baseline ->
check_live_tools (this) -> mcp_strict_audit.

Usage: check_live_tools.py <live.json> [tools/schema_baseline.json]
tools/schema_baseline.json is FROZEN (only 维护者的契约导出脚本 --freeze writes it,
after a contract broadcast)."""
import json, pathlib, sys

live_path = sys.argv[1]
base_path = sys.argv[2] if len(sys.argv) > 2 else str(pathlib.Path(__file__).resolve().parent / "schema_baseline.json")

live = json.loads(pathlib.Path(live_path).read_text(encoding="utf-8-sig"))
base = json.loads(pathlib.Path(base_path).read_text(encoding="utf-8-sig"))
def tools_of(doc):
    """Accept both the JSON-RPC envelope and the bare {"tools": [...]} snapshot."""
    if isinstance(doc, dict) and "result" in doc:
        return doc["result"]["tools"]
    return doc["tools"]

lt = tools_of(live)
bt = tools_of(base)
ln = [t["name"] for t in lt]
bn = [t["name"] for t in bt]
bs = {t["name"]: json.dumps(t.get("inputSchema"), sort_keys=True) for t in bt}
changed = [t["name"] for t in lt if bs.get(t["name"]) != json.dumps(t.get("inputSchema"), sort_keys=True)]
noschema = [t["name"] for t in lt if not t.get("inputSchema")]
print("LIVE=%d BASELINE=%d" % (len(lt), len(bt)))
print("ADDED=" + (",".join(x for x in ln if x not in bn) or "none"))
print("REMOVED=" + (",".join(x for x in bn if x not in ln) or "none"))
print("CHANGED=" + (",".join(changed) or "none"))
print("NO_SCHEMA=" + (",".join(noschema) or "none"))
print("ORDER_MATCH=" + str(ln == bn))
ok = set(ln) == set(bn) and not changed and not noschema
print("SUMMARY live_tool_changes=%d" % (len(set(ln) ^ set(bn)) + len(changed)))
print("EXIT=" + ("0" if ok else "1"))
sys.exit(0 if ok else 1)
