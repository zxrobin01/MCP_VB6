"""Comprehensive tool test battery against the live VB6MCP add-in.

WHAT IT NEEDS
  * a LIVE, OWNED VB6 IDE whose add-in port is passed as argv[1] (default 8766).
    NEVER point it at 8766 when 8766 belongs to another process/project - start your
    own IDE on 8767 or above.
  * the IDE in DESIGN state. No particular starting project is required: the first
    call is add_project, which REPLACES whatever is open (VBProjects.Add with
    Exclusive=True) with a fresh StandardExe, so the battery builds its own fixture.

WHAT IT CREATES / WRITES
  * a project named 工程1 (the IDE's default new-project name) whose components are
    the built-in Form1 plus a class TestCls1 and a module TestMod1 that add_component
    creates and rename_component later turns into TestModRenamed.
  * it calls save_project({path: <repo root>/_test_proj.vbp}): the scratch project at
    the repo root is this battery's OUTPUT (the public tree does not ship it), and the
    battery also writes <repo root>/_test_export.bas and <repo root>/_mcp_test.txt.
  * it starts and stops the program (run / break / restart / full_compile cycles), so
    the instance must be disposable and owned by the caller.

EXIT CODE
  * 0 when every check passed, 1 when any check failed, so a caller/CI can branch.
  * stdout keeps its original format: one line per check, then the totals.
"""
import json
import os
import sys
import time
import urllib.request

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__))).replace('\\', '/')
PORT = int(sys.argv[1]) if len(sys.argv) > 1 else 8766
PASS = 0
FAIL = 0
FAILED_NAMES = []


def rpc(payload):
    req = urllib.request.Request(
        "http://127.0.0.1:%d/rpc" % PORT,
        data=json.dumps(payload).encode("utf-8"),
        headers={"Content-Type": "application/json"})
    with urllib.request.urlopen(req, timeout=180) as r:
        return json.loads(r.read().decode("utf-8"))


def call_tool(name, args=None):
    resp = rpc({"jsonrpc": "2.0", "id": 999, "method": "tools/call",
                "params": {"name": name, "arguments": args or {}}})
    res = resp.get("result", {})
    content = res.get("content", [])
    text = content[0]["text"] if content else "(no content)"
    return res.get("isError", False), text


def t(name, args=None, expect=None, show=140):
    global PASS, FAIL
    try:
        iserr, text = call_tool(name, args)
        if iserr:
            print("[ERR] %s | %s" % (name, text[:300].replace("\n", " ")))
            FAIL += 1
            FAILED_NAMES.append(name)
            return False, text
        if expect is not None and expect not in text:
            print("[??]  %s | missing %r in: %s" % (
                name, expect, text[:200].replace("\n", " ")))
            FAIL += 1
            FAILED_NAMES.append(name)
            return False, text
        print("[OK]  %s | %s" % (name, text[:show].replace("\n", " ")))
        PASS += 1
        return True, text
    except Exception as e:
        print("[EXC] %s | %s" % (name, e))
        FAIL += 1
        FAILED_NAMES.append(name)
        return False, str(e)


print("=" * 78)
print("VB6MCP 功能测试  port=%d" % PORT)
print("=" * 78)

# ---- 1. 工程信息 ----
t("add_project", {"type": "StandardExe"}, expect='"created"')
t("get_component_list", expect='"Form1"')

t("get_project_info", expect='"projectType"')
t("get_component_list", expect='"components"')

# ---- 2. 组件增删 ----
t("add_component", {"type": "module", "name": "TestMod1",
                    "code": "Option Explicit\r\n\r\nPublic Function AddOne(ByVal n As Long) As Long\r\n    AddOne = n + 1\r\nEnd Function\r\n"},
  expect='"created"')
t("add_component", {"type": "class", "name": "TestCls1",
                    "code": "Option Explicit\r\n\r\nPublic Property Get Greeting() As String\r\n    Greeting = \"hello from class\"\r\nEnd Property\r\n"},
  expect='"created"')
t("get_component_list", expect='"TestMod1"')
t("get_component_code", {"name": "TestMod1"}, expect="AddOne")
t("get_component_code", {"name": "TestMod1", "include_line_numbers": True},
  expect="AddOne")

# ---- 3. 代码编辑 ----
t("replace_code", {"name": "TestMod1", "operation": "append",
                   "code": "Public Function Mul2(ByVal n As Long) As Long\r\n    Mul2 = n * 2\r\nEnd Function"},
  expect='"insertedLines":3')
t("replace_code", {"name": "TestMod1", "operation": "insert",
                   "start_line": 1,
                   "code": "' header comment line"},
  expect='"insertedLines":1')
# Contract v1.8.0 made expected_sha256 REQUIRED for replace/delete; without it the
# add-in answers a structured {ok:false, reason:"expected_sha256_required",
# currentSha256:...} refusal. Fetch the CURRENT pane hash from get_component_code
# first and pass it, so this case still exercises a REAL replace and its
# replacedLines receipt instead of testing a refusal.
_ok, _code = call_tool("get_component_code", {"name": "TestMod1"})
_mod_sha = json.loads(_code).get("textSha256", "")
t("replace_code", {"name": "TestMod1", "operation": "replace",
                   "start_line": 2, "end_line": 2,
                   "code": "' replaced comment",
                   "expected_sha256": _mod_sha},
  expect='"replacedLines"')
t("add_code", {"component": "TestMod1",
               "code": "Public Function SubOne(ByVal n As Long) As Long\r\n    SubOne = n - 1\r\nEnd Function"},
  expect='"newTotalLines"')
t("get_component_code", {"name": "TestMod1"}, expect="SubOne")
t("find_in_component", {"component": "TestMod1", "target": "SubOne"},
  expect='"found":true')
t("find_in_component", {"component": "TestMod1", "target": "NoSuchProc"},
  expect='"found":false')

# ---- 4. 重命名/导出/重载 ----
t("rename_component", {"name": "TestMod1", "new_name": "TestModRenamed"},
  expect='"newName"')
t("get_component_list", expect='"TestModRenamed"')
t("export_component", {"name": "TestModRenamed",
                       "path": f"{ROOT}/_test_export.bas"},
  expect='"saved"')

# ---- 5. 文件操作 ----
t("list_files", {"path": f"{ROOT}",
                 "pattern": "*.bas;*.vbp", "max_entries": 20}, expect='"files"')
t("write_file", {"path": f"{ROOT}/_mcp_test.txt",
                 "content": "hello from MCP test\r\nline2 中文内容\r\n"},
  expect='"bytesWritten"')
t("read_file", {"path": f"{ROOT}/_mcp_test.txt"},
  expect="hello from MCP test")
t("search_in_files", {"query": "Option Explicit",
                      "path": f"{ROOT}",
                      "pattern": "*.bas;*.cls;*.frm", "max_results": 10},
  expect='"results"')

# ---- 6. 窗体设计器 ----
t("get_component_properties", {"name": "Form1"}, expect='"properties"')
t("set_component_property", {"name": "Form1", "property": "Caption",
                             "value": "MCP Test Form"}, expect='"property"')
t("form_designer", {"form": "Form1", "action": "info"}, expect='"controls"')
t("add_control", {"form": "Form1", "prog_id": "VB.CommandButton"},
  expect='"added"')
t("form_designer", {"form": "Form1", "action": "info"},
  expect='"controlCount":1')
t("set_control_property", {"form": "Form1", "index": 1,
                           "property": "Caption", "value": "HelloBtn"},
  expect='"property"')
t("create_event_proc", {"component": "Form1", "object": "Command1",
                        "event": "Click"}, expect='"startLine"')
t("get_component_code", {"name": "Form1"}, expect="Command1_Click")
t("form_designer", {"form": "Form1", "action": "select_all"},
  expect='"selectedAll"')

# ---- 7. 成员/属性 ----
t("list_members", {"component": "TestCls1"}, expect='"members"')
t("set_member_attribute", {"component": "TestCls1", "member": "Greeting",
                           "attribute": "description", "value": "test greeting"},
  expect='"attribute"')
t("property_bag", {"target": "project", "section": "MCPTest", "key": "k1",
                   "value": "v1", "has_value": True}, expect='"wrote"')
t("property_bag", {"target": "project", "section": "MCPTest", "key": "k1"},
  expect='"v1"')

# ---- 8. 窗口/选择 ----
t("list_windows", expect='"windows"')
t("activate_component", {"name": "TestModRenamed", "window": "code"})
t("get_selection", expect='"activeCodeModule"')
t("get_codepane", {"component": "TestModRenamed"}, expect='"component"')
t("set_codepane", {"component": "TestModRenamed",
                   "selection": "2,1,2,1", "top_line": 1}, expect='"applied"')

# ---- 9. 引用 ----
t("add_reference", {"mode": "guid",
                    "guid": "{420B2830-E718-11CF-893D-00A0C9054228}",
                    "major": 1, "minor": 0}, expect='"name"')
t("remove_reference", {"name": "Scripting"}, expect='"removed"')

# ---- 10. 插件/菜单 ----
t("list_addins", expect='"addins"')
t("set_addin_state", {"prog_id": "VB6MCP.Connect", "connect": True},
  expect='"connect"')
t("command_bars", {"action": "list", "bar": "菜单条"}, expect='"bars"')

# ---- 11. 保存/编译 ----
t("save_project", {"path": f"{ROOT}/_test_proj.vbp"}, expect='"project"')
t("reload_component", {"name": "TestModRenamed"}, expect='"reloaded"')

print("=" * 78)
print("运行控制测试 (make_project + run/break/end)")
print("=" * 78)
t("make_project", expect='"compiled":true', show=200)
time.sleep(1)
t("control_execution", {"mode": "run"}, show=120)


def wait_state(want, tries, label):
    global PASS, FAIL
    for _ in range(tries):
        iserr, text = call_tool("get_project_info")
        if not iserr and ('"ideState":"%s"' % want) in text:
            print("[OK]  %s | ideState=%s reached" % (label, want))
            PASS += 1
            return True
        time.sleep(1)
    print("[??]  %s | never reached ideState=%s" % (label, want))
    FAIL += 1
    FAILED_NAMES.append(label)
    return False


wait_state("run", 20, "wait_state run")
t("control_execution", {"mode": "end"}, expect='"executed":true', show=120)
wait_state("design", 10, "wait_state design")
t("get_project_info", show=60)

print("=" * 78)
print("break/restart/full_compile 模式测试")
print("=" * 78)
# break cycle: run -> break -> end
t("control_execution", {"mode": "run"}, show=120)
wait_state("run", 20, "wait_state run-b")
t("control_execution", {"mode": "break"}, expect='"executed":true', show=120)
wait_state("break", 10, "wait_state break")
t("control_execution", {"mode": "end"}, expect='"executed":true', show=120)
wait_state("design", 10, "wait_state design-b")

# restart cycle: run -> restart (deferred) -> still running -> end
t("control_execution", {"mode": "run"}, show=120)
wait_state("run", 20, "wait_state run-r")
t("control_execution", {"mode": "restart"}, expect='"restartPending":true', show=120)
wait_state("run", 20, "wait_state run-after-restart")
t("control_execution", {"mode": "end"}, expect='"executed":true', show=120)
wait_state("design", 10, "wait_state design-r")

# full_compile cycle: full compile + run -> end
t("control_execution", {"mode": "full_compile"}, show=120)
wait_state("run", 20, "wait_state run-fc")
t("control_execution", {"mode": "end"}, expect='"executed":true', show=120)
wait_state("design", 10, "wait_state design-fc")

print("=" * 78)
print(" 回归: 写前块结构守卫 / 歧义候选 / 多行锚点 / 过程回显")
print("=" * 78)

T99_FIX = ("Option Explicit\r\n\r\n"
           "Private x As Long\r\nPrivate a As Long\r\n\r\n"
           "Public Sub T99Dup()\r\n    a = 1\r\n    a = 1\r\nEnd Sub\r\n\r\n"
           "Public Sub T99Victim()\r\n    x = 1\r\nEnd Sub\r\n\r\n"
           "Public Sub T99Frag()\r\n    x = 1\r\nEnd Sub\r\n")


def t99(label, tool, args, judge):
    """judge(obj) -> bool; a FAIL prints the receipt so the miss is diagnosable."""
    global PASS, FAIL
    _e, _txt = call_tool(tool, args)
    try:
        obj = json.loads(_txt)
    except Exception:
        obj = None
    if (not _e) and isinstance(obj, dict) and judge(obj):
        print("[OK]  %s | %s" % (label, json.dumps(obj, ensure_ascii=False)[:150]))
        PASS += 1
    else:
        print("[??]  %s | %s" % (label, _txt[:220].replace("\n", " ")))
        FAIL += 1
        FAILED_NAMES.append("task99:" + label)
    return obj


def _t99_pane():
    return (json.loads(call_tool("get_component_code", {"name": "ModT99"})[1]).get("text") or "").split("\n")


def _t99_sha():
    return json.loads(call_tool("get_component_code", {"name": "ModT99"})[1]).get("textSha256", "")


t("add_component", {"type": "module", "name": "ModT99", "code": T99_FIX}, expect='"created"')

# --- the pasted block's OWN pairing, with the offending line numbers ---------
t99("06 extra end refused", "replace_code",
    {"name": "ModT99", "operation": "append",
     "code": "Public Function T99Bad() As Long\r\n    T99Bad = 1\r\nEnd Function\r\nEnd Function"},
    lambda o: o.get("reason") == "extra_procedure_end" and o.get("writePerformed") is False
              and bool(o.get("extraEndLines")) and bool(o.get("pairs")) and bool(o.get("suggestedText")))
t99("06 normal block accepted", "replace_code",
    {"name": "ModT99", "operation": "append",
     "code": "Public Function T99Good() As Long\r\n    T99Good = 2\r\nEnd Function"},
    lambda o: o.get("ok") is not False and "newTotalLines" in o)
t99("06 missing end refused (reverse)", "replace_code",
    {"name": "ModT99", "operation": "append",
     "code": "Public Function T99Open() As Long\r\n    T99Open = 3"},
    lambda o: o.get("reason") == "missing_procedure_end" and o.get("writePerformed") is False
              and bool(o.get("unclosedHeaderLines")))
# NOTE: the fragment case uses T99Frag (NOT T99Dup) on purpose - the ambiguity case
# below needs T99Dup's two identical lines to stay untouched.
_l = _t99_pane()
_frag = next((i + 1 for i, x in enumerate(_l) if "Sub T99Frag()" in x), 0)
t99("06 fragment accepted (no false positive)", "replace_code",
    {"name": "ModT99", "operation": "replace", "start_line": _frag, "end_line": _frag + 1,
     "code": "Public Sub T99Frag()\r\n    x = 2", "expected_sha256": _t99_sha()},
    lambda o: o.get("ok") is not False and "appliedRange" in o)

# --- ambiguity returns candidates and never writes; block anchors work ------
t99("05 ambiguous anchor -> candidates", "edit_by_anchor",
    {"name": "ModT99", "anchor": "    a = 1", "new_text": "    a = 9"},
    lambda o: o.get("reason") == "anchor_not_unique" and len(o.get("hits") or []) == 2
              and all("line" in h and "proc" in h and "text" in h for h in o["hits"]))
t99("05 multi-line anchor replaces block", "edit_by_anchor",
    {"name": "ModT99", "anchor": "Public Sub T99Dup()\r\n    a = 1\r\n    a = 1\r\nEnd Sub",
     "new_text": "Public Sub T99Dup()\r\n    a = 7\r\nEnd Sub", "expected_sha256": _t99_sha()},
    lambda o: o.get("ok") is not False and o.get("matchedSpanLines") == 4)
t99("05/06 preview predicts the refusal", "preview_edit",
    {"name": "ModT99", "edits": [{"op": "R",
                                  "anchor": "Public Function T99Good() As Long\r\n    T99Good = 2\r\nEnd Function",
                                  "text": "Public Function T99Good() As Long\r\n    T99Good = 3"}]},
    lambda o: o.get("reason") == "predicted_procedure_imbalance" and o.get("tool") == "preview_edit")

# --- a stale sha refuses the line-number edit; the receipt echoes the proc ---
_stale = _t99_sha()
call_tool("replace_code", {"name": "ModT99", "operation": "insert", "start_line": 1, "code": "' drift"})
_l2 = _t99_pane()
_gl = next((i + 1 for i, x in enumerate(_l2) if "T99Good = 2" in x), 0)
t99("07 stale sha refuses drift", "replace_code",
    {"name": "ModT99", "operation": "replace", "start_line": _gl, "end_line": _gl,
     "code": "    T99Good = 99", "expected_sha256": _stale},
    lambda o: o.get("reason") == "sha256_mismatch" and bool(o.get("actualSha256")))
t99("07 proc echo on replace", "replace_code",
    {"name": "ModT99", "operation": "replace", "start_line": _gl, "end_line": _gl,
     "code": "    T99Good = 5", "expected_sha256": _t99_sha()},
    lambda o: (o.get("procBefore") or {}).get("proc") == "T99Good"
              and (o.get("procAfter") or {}).get("proc") == "T99Good"
              and (o.get("procBefore") or {}).get("startLine", 0) > 0)

print("=" * 78)
print(" 回归: dependency_graph 同输入三次必须给出同一份答案（确定性配额）")
print("=" * 78)

# A guaranteed CALL GRAPH for the fixture. Without it the fixture's procedures
# never call each other, so procEdgesFound would be 0 and this regression could
# not tell "the procedure-edge phase ran" from "it was cut". T95A -> T95B is
# asserted verbatim below.
t("add_component", {"type": "module", "name": "ModT95",
                    "code": "Option Explicit\r\n\r\nPublic Sub T95A()\r\n    T95B\r\nEnd Sub\r\n\r\n"
                            "Public Sub T95B()\r\n    T95C\r\nEnd Sub\r\n\r\n"
                            "Public Sub T95C()\r\n    T95A\r\nEnd Sub\r\n"},
  expect='"created"')

# ENVELOPE MASK LIST
# The ONLY fields a receipt may differ on between two calls with the SAME input.
# Every entry names the field AND why it can never be part of the answer; every
# other byte must be identical. Masking anything not listed here - or adding a
# mask without a reason - silently turns this regression into a no-op.
ENVELOPE_MASK = [
    (r'"ms":\d+',           "self wall time of THIS call (top level ms)"),
    (r'"elapsedMs":\d+',    "self wall time, second spelling (budget suffix)"),
    (r'"stampedCount":\d+', "session receipt counter, +1 per sealed receipt (inflight{})"),
    (r'"sinceMs":\d+',      "age of the in-flight call at stamp time (inflight{})"),
    (r'"pid":\d+',          "IDE process id: session identity, changes per IDE restart"),
    (r'"port":\d+',         "add-in TCP port: session identity"),
]


def _mask_envelope(text):
    import re
    for _pat, _why in ENVELOPE_MASK:
        text = re.sub(_pat, lambda m: m.group(0).split(":")[0] + ":0", text)
    return text


def depgraph_identity(label, strict_partial):
    """3 x dependency_graph: after masking ONLY the enumerated envelope the answer
    must be byte-identical, partial must agree with the *StageCompleted flags, and
    the default path must still yield procedure edges (the defect this guards)."""
    global PASS, FAIL
    runs = []
    for _i in range(3):
        _iserr, _txt = call_tool("dependency_graph", {})
        if _iserr:
            print("[ERR] %s dependency_graph run %d | %s" % (label, _i + 1, _txt[:200]))
            FAIL += 1
            FAILED_NAMES.append("dependency_graph-repeat:" + label)
            return None
        runs.append(_txt)
    masked = [_mask_envelope(x) for x in runs]
    if masked[0] == masked[1] == masked[2]:
        print("[OK]  %s: 三次逐字节相同（只掩码 %d 个信封字段），bytes=%d"
              % (label, len(ENVELOPE_MASK), len(runs[0])))
        PASS += 1
    else:
        print("[??]  %s: 三次不一致（掩码后仍不同）" % label)
        FAIL += 1
        FAILED_NAMES.append("dependency_graph-repeat:" + label)
    o = json.loads(runs[0])
    flags = {k: bool(o.get(k)) for k in ("edgeStageCompleted", "purityStageCompleted",
                                         "cyclesStageCompleted", "procEdgeStageCompleted")}
    partial = bool(o.get("partial"))
    all_done = all(flags.values())
    print("      coverage: partial=%s quotaHit=%s timing=%s chars=%s/%s comps=%s/%s "
          "procEdges=%s procsScanned=%s/%s stages=%s"
          % (partial, o.get("budgetExhausted"), o.get("timingTruncated"),
             o.get("charsProcessed"), o.get("charQuota"),
             o.get("componentsScanned"), o.get("componentsTotal"),
             o.get("procEdgesFound"), o.get("proceduresScanned"), o.get("procCount"),
             flags))
    if partial == (not all_done):
        print("[OK]  %s: partial 与 *StageCompleted 自洽（阶段被截断 => 其 flag 为 false）" % label)
        PASS += 1
    else:
        print("[??]  %s: partial=%s 但 all_stages_done=%s" % (label, partial, all_done))
        FAIL += 1
        FAILED_NAMES.append("dependency_graph-consistency:" + label)
    if partial and not (o.get("budgetExhausted") or o.get("timingTruncated")):
        print("[??]  %s: partial=true 却没有任何截止标志" % label)
        FAIL += 1
        FAILED_NAMES.append("dependency_graph-cutflag:" + label)
    else:
        PASS += 1
    if o.get("procEdgesFound", 0) <= 0:
        print("[??]  %s: procEdgesFound=0（默认档必须产出边）" % label)
        FAIL += 1
        FAILED_NAMES.append("dependency_graph-noedges:" + label)
    else:
        PASS += 1
    if strict_partial and not partial:
        print("[??]  %s: 期望部分覆盖（本段覆盖截断路径）却得到完整结果" % label)
        FAIL += 1
        FAILED_NAMES.append("dependency_graph-partial:" + label)
    if not partial:
        if o.get("budgetExhausted") or o.get("timingTruncated"):
            print("[??]  %s: 全部阶段完成却仍有截止标志" % label)
            FAIL += 1
            FAILED_NAMES.append("dependency_graph-cutflag2:" + label)
        else:
            print("[OK]  %s: 完整结果（budgetExhausted=false 且四阶段 completed）同样三次一致" % label)
            PASS += 1
    return o


o_fixture = depgraph_identity("fixture", False)
if o_fixture is not None:
    # NOTE: procedure names in the graph are LOWERCASED by the extractor (pre-existing
    # behaviour: the declaration line is matched case-insensitively and m_pn keeps the
    # lowercased name - e.g. "ServerForm.startlistening"), so compare case-insensitively.
    _pe = o_fixture.get("procEdges") or []
    _got = {(str(e.get("from", "")).lower(), str(e.get("to", "")).lower())
            for e in _pe if isinstance(e, dict)}
    _want = ("modt95.t95a", "modt95.t95b")
    if _want in _got:
        print("[OK]  fixture: 过程边 %s -> %s 已解析（边阶段真的扫过过程体）" % _want)
        PASS += 1
    else:
        print("[??]  fixture: 未找到过程边 %s；实际=%s" % (str(_want), sorted(_got)[:8]))
        FAIL += 1
        FAILED_NAMES.append("dependency_graph-procedge")

# PART 2: the SAME contract on the SHIPPED project, where the default quota is
# really exceeded, so this is the PARTIAL/truncated path (PART 1 above is a tiny
# fixture and is therefore normally complete). Skipped with a note when the shipped
# project is not next to this script or does not open (missing references).
_shipped = os.path.join(ROOT, "MCP_VB6.vbp")
if os.path.exists(_shipped):
    t("save_project", {"path": f"{ROOT}/_test_proj.vbp"}, show=60)
    _iserr2, _txt2 = call_tool("open_project", {"path": _shipped, "exclusive": True})
    if _iserr2:
        print("      (skip shipped-project 段: open_project 失败 | %s)"
              % _txt2[:160].replace("\n", " "))
    else:
        depgraph_identity("shipped", True)
else:
    print("      (skip shipped-project 段: %s 不存在)" % _shipped)

print("=" * 78)
print("结果: 通过 %d, 失败 %d" % (PASS, FAIL))
if FAILED_NAMES:
    print("失败项:", ", ".join(FAILED_NAMES))
print("=" * 78)
# Packaging gap: the old script always exited 0, so a caller could not tell a green
# run from a red one. stdout format is unchanged - only the exit code is.
sys.exit(1 if FAIL else 0)
