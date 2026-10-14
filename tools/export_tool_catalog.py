"""Regenerate docs/TOOL_CATALOG.md from the live tools/list snapshot + the hub's own tools.

Source of truth: docs/tools_schema.json (the live snapshot captured by the maintainer's
snapshot tool) MERGED with bridge/hub_tools.json - the tools the HUB owns (ide_* plus the
5 lessons tools that moved from the add-in to tools/lessons.py in contract v1.11.0).
We deliberately do NOT parse integration/CodeTools.cls: that sandbox copy lags the product
and once made the catalog under-count the tool surface (the catalog said 96 while the
snapshot and the baseline said 99).

Merge rule (identical to the hub's tools/list de-duplication): a name present in BOTH files
is taken from hub_tools.json (the hub copy WINS) and the add-in copy is dropped; a hub name
absent from the snapshot is appended.  That keeps the moved lesson tools documented after
they leave the add-in face, and keeps the DSH-facing total stable across the move.
`TOOLS=` keeps its historical meaning (the ADD-IN face count, the number the other gates
print); the merged DSH-facing count is printed as `CATALOG_TOOLS=`.

Usage:
  python tools/export_tool_catalog.py [--schema <json>] [--hub-tools <json>] [--out <md>]
    # defaults: docs/tools_schema.json / bridge/hub_tools.json / docs/TOOL_CATALOG.md
  (--out writes elsewhere; use it for a non-destructive preview of a regeneration)

Order (after any tool-surface/description change):
  1. dump_tools_schema.py -> docs/tools_schema.json
  2. export_tool_catalog.py (this file) -> docs/TOOL_CATALOG.md
  3. check_schema_baseline.py (snapshot vs FROZEN baseline -> schema_changes=0)
  4. check_live_tools.py docs/tools_schema.json (independent cross-check -> ORDER_MATCH=True)
  5. mcp_strict_audit.py (strict_failures=0)

tools/schema_baseline.json is a FROZEN contract artifact: only the snapshot tool's
--freeze writes it, and only after a contract broadcast (a version bump plus an ack from
every impacted process).  Never freeze just to force 0.
"""
import json, os, sys


def opt_value(name, default=None):
    if name in sys.argv:
        i = sys.argv.index(name)
        if i + 1 < len(sys.argv) and not sys.argv[i + 1].startswith('--'):
            return sys.argv[i + 1]
    return default


def tools_of(doc):
    if isinstance(doc, dict) and 'result' in doc:
        return doc['result']['tools']
    return doc['tools']


def load_hub_tools(path):
    """The hub's own tool face, or ([], False) when the file cannot be read."""
    try:
        with open(path, encoding='utf-8-sig') as f:
            hub = json.load(f)
    except (OSError, ValueError):
        return [], False
    tools = hub.get('tools', []) if isinstance(hub, dict) else []
    if not isinstance(tools, list):
        return [], False
    return [t for t in tools if isinstance(t, dict) and t.get('name')], True


def merge_hub_tools(schema_tools, hub_tools):
    """Merge the hub face into the add-in face.  Returns (merged, added, overridden).

    Hub-provided names WIN (the add-in copy is replaced in place, keeping snapshot order);
    hub names the snapshot does not carry are appended in hub-file order.
    """
    index = {}
    for i, t in enumerate(schema_tools):
        index[str(t.get('name', ''))] = i
    merged = list(schema_tools)
    added, overridden, seen_hub = 0, 0, set()
    for t in hub_tools:
        n = str(t.get('name', ''))
        if n in seen_hub:                      # a duplicate inside hub_tools.json: keep the first
            continue
        seen_hub.add(n)
        if n in index:
            merged[index[n]] = t               # hub copy wins
            overridden += 1
        else:
            index[n] = len(merged)
            merged.append(t)
            added += 1
    return merged, added, overridden


def main():
    root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    schema = opt_value('--schema', os.path.join(root, 'docs', 'tools_schema.json'))
    hubf = opt_value('--hub-tools', os.path.join(root, 'bridge', 'hub_tools.json'))
    out = opt_value('--out', os.path.join(root, 'docs', 'TOOL_CATALOG.md'))
    with open(schema, encoding='utf-8-sig') as f:
        tools = tools_of(json.load(f))
    schema_count = len(tools)
    hub_tools, hub_ok = load_hub_tools(hubf)
    merged, hub_added, hub_overridden = merge_hub_tools(tools, hub_tools)
    try:
        rel = os.path.relpath(schema, root).replace('\\', '/')
    except ValueError:  # e.g. --schema on another drive (test fixtures)
        rel = schema.replace('\\', '/')
    try:
        relhub = os.path.relpath(hubf, root).replace('\\', '/')
    except ValueError:
        relhub = hubf.replace('\\', '/')
    lines = []
    lines.append('# VB6MCP 工具总览（工具面契约沙箱）')
    lines.append('')
    lines.append('- 来源：%s（由维护者侧的快照工具从运行中的 add-in 抓取；本文件是生成物，勿手改）'
                 ' + %s（hub 自有工具：ide_* 以及在 v1.11.0 迁到 tools/lessons.py 的 lessons 工具；'
                 '同名时以 hub 定义为准）' % (rel, relhub))
    lines.append('- 数量：%d（快照 %d + hub 追加 %d；hub 覆盖同名 %d）'
                 % (len(merged), schema_count, hub_added, hub_overridden))
    if not hub_ok:
        lines.append('- 警告：未能读取 %s，hub 工具未并入本目录。' % relhub)
    lines.append('')
    lines.append('## 使用顺序（工具集/描述变化后，按序执行并原样回贴输出）')
    lines.append('')
    lines.append('```text')
    lines.append('rem 1. 抓 live 快照并重冻：需要运行中的 add-in（自有实例；--url 必须显式给出），且需先经契约广播授权')
    lines.append('rem --freeze 会写 docs/tools_schema.json 与冻结基线 tools/schema_baseline.json；只想预览就改用 --out 指向临时路径。')
    lines.append('python tools/dump_tools_schema.py --freeze --url http://127.0.0.1:<your own port>/rpc')
    lines.append('rem 2. 从 docs/tools_schema.json + bridge/hub_tools.json 重新生成本文件')
    lines.append('python tools/export_tool_catalog.py')
    lines.append('rem 3-4. 离线只读，可直接运行')
    lines.append('python tools/check_schema_baseline.py')
    lines.append('python tools/check_live_tools.py docs/tools_schema.json')
    lines.append('rem 5. 严格 schema 审计：需要活着的 IDE（自有实例，端口 8767 及以上；8766 留给维护者自己的交互会话）')
    lines.append('python tools/mcp_strict_audit.py --url http://127.0.0.1:<your own port>/mcp')
    lines.append('```')
    lines.append('')
    lines.append('## 冻结件与授权前提')
    lines.append('')
    lines.append('- `tools/schema_baseline.json` 是**冻结件**：快照默认**不写**它；只有显式 `--freeze` 才写，且必须**先经契约广播**（版本号递增 + 每个受影响方的确认）。')
    lines.append('- 禁止为了得到 `schema_changes=0` 就 `--freeze`——那会让基线比对退化为恒真。')
    lines.append('- 数据源以 `docs/tools_schema.json`（live 快照）+ `bridge/hub_tools.json`（hub 自有工具）为准；**不要**解析 `integration/CodeTools.cls`（沙箱副本会滞后，曾致数量失真）。')
    lines.append('')
    lines.append('| 工具 | 说明 |')
    lines.append('|---|---|')
    for t in merged:
        n = str(t.get('name', ''))
        d = str(t.get('description', '')).replace('|', '/').replace('\r', ' ').replace('\n', ' ').strip()
        lines.append('| %s | %s |' % (n, d))
    with open(out, 'w', encoding='utf-8') as f:
        f.write('\n'.join(lines) + '\n')
    # TOOLS= keeps its historical meaning (the add-in face count the other gates print).
    print('TOOLS=%d OUT=%s' % (schema_count, out))
    print('CATALOG_TOOLS=%d (schema=%d hub_added=%d hub_override=%d HUB=%s)'
          % (len(merged), schema_count, hub_added, hub_overridden, relhub))
    return 0


raise SystemExit(main())
