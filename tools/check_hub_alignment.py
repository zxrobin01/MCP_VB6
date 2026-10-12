"""Machine-check that the hub tool face matches the frozen tool face.

TASK-P3-020 (内部独立宿主方案 §8 rule 3).  Makes "someone forgot to
regenerate the hub snapshot" fail loudly instead of drifting silently.

Checks:
  1. bridge/hub_schema.generated.json tools == docs/tools_schema.json tools, EXACTLY and
     in order (name / description / inputSchema); the first difference is reported.
  2. hub-only tools (bridge/hub_tools.json) have NO name collision with the add-in face.
  3. no omissions (snapshot tools == schema tools).

Usage:
  python tools/check_hub_alignment.py [--snapshot f] [--schema f] [--hub-tools f]
Exit 0 = aligned, 1 = drift.
"""
import json, os, sys


def opt(name, default=None):
    if name in sys.argv:
        i = sys.argv.index(name)
        if i + 1 < len(sys.argv) and not sys.argv[i + 1].startswith('--'):
            return sys.argv[i + 1]
    return default


def tools_of(doc):
    if isinstance(doc, dict) and 'result' in doc:
        return doc['result']['tools']
    return doc['tools']


def norm(t):
    return {'name': t.get('name'), 'description': t.get('description', ''),
            'inputSchema': t.get('inputSchema', {})}


def main():
    root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    snapf = opt('--snapshot', os.path.join(root, 'bridge', 'hub_schema.generated.json'))
    schemaf = opt('--schema', os.path.join(root, 'docs', 'tools_schema.json'))
    hubf = opt('--hub-tools', os.path.join(root, 'bridge', 'hub_tools.json'))
    with open(snapf, encoding='utf-8-sig') as f:
        snap = json.load(f)
    with open(schemaf, encoding='utf-8-sig') as f:
        schema = json.load(f)
    with open(hubf, encoding='utf-8-sig') as f:
        hub = json.load(f)
    a, b = [norm(t) for t in snap['tools']], [norm(t) for t in tools_of(schema)]
    errs = []
    if len(a) != len(b):
        errs.append('toolCount: snapshot=%d schema=%d' % (len(a), len(b)))
    for i, (x, y) in enumerate(zip(a, b)):
        if x != y:
            for k in ('name', 'description', 'inputSchema'):
                if x[k] != y[k]:
                    errs.append('index %d field %s: snapshot=%r schema=%r' % (
                        i, k, str(x[k])[:120], str(y[k])[:120]))
                    break
            if len(errs) >= 5:
                break
    anames = [t['name'] for t in a]
    bnames = [t['name'] for t in b]
    if set(anames) != set(bnames):
        errs.append('name set mismatch: only_in_snapshot=%s only_in_schema=%s' % (
            sorted(set(anames) - set(bnames))[:5], sorted(set(bnames) - set(anames))[:5]))
    hnames = [t['name'] for t in hub.get('tools', [])]
    coll = sorted(set(hnames) & set(bnames))
    if coll:
        errs.append('hub-only tool name collides with add-in face: %s' % coll)
    def rel(p):
        try:
            return os.path.relpath(p, root).replace('\\', '/')
        except ValueError:  # e.g. --snapshot on another drive (test fixtures)
            return p.replace('\\', '/')
    print('SNAPSHOT=%s count=%d version=%s' % (rel(snapf), len(a), snap.get('contractVersion')))
    print('SCHEMA=%s count=%d' % (rel(schemaf), len(b)))
    print('HUB_ONLY=%s' % ','.join(hnames))
    if errs:
        for e in errs:
            print('FAIL ' + e)
        print('SUMMARY hub_alignment_failures=%d' % len(errs))
        return 1
    print('ORDER_MATCH=True  DESCRIPTIONS_MATCH=True  INPUT_SCHEMA_MATCH=True  NO_NAME_COLLISION=True')
    print('SUMMARY hub_alignment_failures=0')
    return 0


raise SystemExit(main())
