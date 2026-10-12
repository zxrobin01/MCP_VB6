"""Compare the captured live snapshot (docs/tools_schema.json) with the FROZEN baseline.

Order: dump_tools_schema -> export_tool_catalog -> check_schema_baseline (this)
-> check_live_tools (independent) -> mcp_strict_audit.

tools/schema_baseline.json is a FROZEN contract artifact: it is only rewritten by
维护者的契约导出脚本 --freeze, after a contract broadcast.  Never freeze to force
schema_changes=0.
"""
import json, os, sys

def index(d):
    return {t.get('name'): t for t in d.get('tools', [])}

def main():
    root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    base_p = os.path.join(root, 'tools', 'schema_baseline.json')
    cur_p = os.path.join(root, 'docs', 'tools_schema.json')
    base = json.load(open(base_p, encoding='utf-8'))
    cur = json.load(open(cur_p, encoding='utf-8'))
    b = index(base)
    c = index(cur)
    added = sorted(set(c) - set(b))
    removed = sorted(set(b) - set(c))
    changed = sorted(n for n in (set(b) & set(c)) if b[n] != c[n])
    print('BASELINE=%d CURRENT=%d' % (len(b), len(c)))
    print('ADDED=' + (','.join(added) if added else 'none'))
    print('REMOVED=' + (','.join(removed) if removed else 'none'))
    print('CHANGED=' + (','.join(changed) if changed else 'none'))
    bad = bool(added or removed or changed)
    print('SUMMARY schema_changes=%d' % (len(added) + len(removed) + len(changed)))
    return 1 if bad else 0

raise SystemExit(main())
