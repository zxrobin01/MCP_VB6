"""Regenerate docs/TOOL_CATALOG.md from the live tools/list snapshot.

Source of truth: docs/tools_schema.json (the live snapshot captured by the maintainer's
snapshot tool).  We deliberately do NOT parse integration/CodeTools.cls: that sandbox copy
lags the product and once made the catalog under-count the tool surface (the catalog said 96
while the snapshot and the baseline said 99).

Usage:
  python tools/export_tool_catalog.py [--schema <json>]   # default docs/tools_schema.json

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


def main():
    root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    schema = opt_value('--schema', os.path.join(root, 'docs', 'tools_schema.json'))
    with open(schema, encoding='utf-8-sig') as f:
        tools = tools_of(json.load(f))
    try:
        rel = os.path.relpath(schema, root).replace('\\', '/')
    except ValueError:  # e.g. --schema on another drive (test fixtures)
        rel = schema.replace('\\', '/')
    lines = []
    lines.append('# VB6MCP 工具总览（工具面契约沙箱）')
    lines.append('')
    lines.append('- 来源：%s（由维护者侧的快照工具从运行中的 add-in 抓取；本文件是生成物，勿手改）' % rel)
    lines.append('- 数量：%d' % len(tools))
    lines.append('')
    lines.append('## 使用顺序（工具集/描述变化后，按序执行并原样回贴输出）')
    lines.append('')
    lines.append('```text')
    lines.append('rem 1. 抓 live 快照并重冻：需要运行中的 add-in（自有实例；--url 必须显式给出），且需先经契约广播授权')
    lines.append('rem --freeze 会写 docs/tools_schema.json 与冻结基线 tools/schema_baseline.json；只想预览就改用 --out 指向临时路径。')
    lines.append('python tools/dump_tools_schema.py --freeze --url http://127.0.0.1:<your own port>/rpc')
    lines.append('rem 2. 从 docs/tools_schema.json 重新生成本文件')
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
    lines.append('- 数据源以 `docs/tools_schema.json`（live 快照）为准；**不要**解析 `integration/CodeTools.cls`（沙箱副本会滞后，曾致数量失真）。')
    lines.append('')
    lines.append('| 工具 | 说明 |')
    lines.append('|---|---|')
    for t in tools:
        n = str(t.get('name', ''))
        d = str(t.get('description', '')).replace('|', '/').replace('\r', ' ').replace('\n', ' ').strip()
        lines.append('| %s | %s |' % (n, d))
    out = os.path.join(root, 'docs', 'TOOL_CATALOG.md')
    with open(out, 'w', encoding='utf-8') as f:
        f.write('\n'.join(lines) + '\n')
    print('TOOLS=%d OUT=%s' % (len(tools), out))
    return 0


raise SystemExit(main())
