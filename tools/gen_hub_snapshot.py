"""Generate bridge/hub_schema.generated.json from docs/tools_schema.json.

Rule: the hub tool face is GENERATED, never
hand-written.  This is the single producer; docs/TOOL_CATALOG.md rides the same pipeline,
so any contract-driven tool-surface change reaches the hub automatically.

contractVersion is EXTRACTED from docs/CONTRACTS.md (the newest row of the 版本记录 table),
NOT inferred from the tool count: contracts v1.3.1/v1.3.2 changed the version without
changing the tool count, so count-matching under-reported the contract.  The
declared tool count is used only as a CROSS-CHECK and a mismatch is reported loudly
(`versionToolCountMismatch`) instead of silently picking one of the two.

Usage:
  python tools/gen_hub_snapshot.py [--schema docs/tools_schema.json]
                                   [--contract docs/CONTRACTS.md]
                                   [--out bridge/hub_schema.generated.json]
"""
import json, os, re, sys


def opt(name, default=None):
    if name in sys.argv:
        i = sys.argv.index(name)
        if i + 1 < len(sys.argv) and not sys.argv[i + 1].startswith('--'):
            return sys.argv[i + 1]
    return default


def _vkey(s):
    return tuple(int(x) for x in s.split('.'))


def version_table(contract_path):
    """[(version, declared_target_count_or_None)] in file order, from the 版本记录 table."""
    try:
        t = open(contract_path, encoding='utf-8', errors='replace').read()
    except OSError:
        return []
    rows, in_table = [], False
    for line in t.splitlines():
        if line.startswith('## 版本记录'):
            in_table = True
            continue
        if not in_table:
            continue
        if line.startswith('## '):
            break
        if not line.startswith('|'):
            continue
        if '版本' in line and '日期' in line:
            continue
        if set(line.replace('|', '').strip()) <= set('-: '):
            continue
        vs = re.findall(r'v(\d+\.\d+\.\d+)', line)
        if not vs:
            continue
        v = max(vs, key=_vkey)
        m = re.search(r'(\d+)\s*(?:->|\u2192)\s*(\d+)\b', line)
        rows.append((v, int(m.group(2)) if m else None))
    return rows


def contract_version(contract_path, tool_count):
    """Explicitly-extracted current contract version + (declared count, mismatch flag)."""
    rows = version_table(contract_path)
    if not rows:
        return ('unknown', None, False)
    order = sorted({r[0] for r in rows}, key=_vkey)
    byv = {r[0]: r[1] for r in rows}
    cur = order[-1]
    declared = byv.get(cur)
    if declared is None:
        # a version that did not change the count inherits the nearest older declared count
        for v in reversed([x for x in order if _vkey(x) < _vkey(cur)]):
            if byv.get(v) is not None:
                declared = byv[v]
                break
    mismatch = declared is not None and declared != tool_count
    return ('v' + cur, declared, mismatch)


def instructions_of(root):
    """Pull the contract's fixed instructions block (and its BREAKING list) out of
    docs/CONTRACTS.md so the hub never hand-writes version/breaking prose."""
    p = os.path.join(root, 'docs', 'CONTRACTS.md')
    try:
        t = open(p, encoding='utf-8', errors='replace').read()
    except OSError:
        return '', []
    i = t.find('## v1.3.0')
    seg = t[i:] if i >= 0 else t
    m = re.search(r'```text\s*\n(.*?)\n```', seg, re.S)
    block = m.group(1).strip() if m else ''
    breaking = []
    if block:
        lines = block.splitlines()
        k = -1
        for idx, l in enumerate(lines):
            if l.startswith('BREAKING'):
                k = idx
                break
        if k >= 0:
            breaking = [l for l in lines[k + 1:] if l.strip()]
    return block, breaking


def main():
    root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    schema = opt('--schema', os.path.join(root, 'docs', 'tools_schema.json'))
    contract = opt('--contract', os.path.join(root, 'docs', 'CONTRACTS.md'))
    out = opt('--out', os.path.join(root, 'bridge', 'hub_schema.generated.json'))
    with open(schema, encoding='utf-8-sig') as f:
        doc = json.load(f)
    tools = doc['result']['tools'] if isinstance(doc, dict) and 'result' in doc else doc['tools']
    face = [{'name': t.get('name'), 'description': t.get('description', ''),
             'inputSchema': t.get('inputSchema', {})} for t in tools]
    cv, declared, mismatch = contract_version(contract, len(face))
    block, breaking = instructions_of(root)

    def rel(p):
        try:
            return os.path.relpath(p, root).replace('\\', '/')
        except ValueError:  # e.g. --out on another drive (test fixtures)
            return p.replace('\\', '/')

    snap = {
        'generatedFrom': rel(schema),
        'contractVersion': cv,
        'contractDeclaredToolCount': declared,
        'versionToolCountMismatch': bool(mismatch),
        'toolCount': len(face),
        'instructions': block,
        'breaking': breaking,
        'tools': face,
    }
    os.makedirs(os.path.dirname(out), exist_ok=True)
    with open(out, 'w', encoding='utf-8') as f:
        json.dump(snap, f, ensure_ascii=False, indent=2)
        f.write('\n')
    if mismatch:
        print('WARNING versionToolCountMismatch: contract %s declares %s tools, snapshot has %d '
              '(snapshot NOT silently reconciled)' % (cv, declared, len(face)))
    print('HUB_SNAPSHOT=%s TOOLS=%d VERSION=%s DECLARED=%s MISMATCH=%s' % (
        rel(out), len(face), cv, declared, bool(mismatch)))
    return 0


raise SystemExit(main())
