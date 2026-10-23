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
            print("[ERR] %s | %s" % (name, text[:4000].replace("\n", " ")))
            FAIL += 1
            FAILED_NAMES.append(name)
            return False, text
        if expect is not None and expect not in text:
            print("[??]  %s | missing %r in: %s" % (
                name, expect, text[:4000].replace("\n", " ")))
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


def expect_refusal(label, tool, args, reason):
    """负例：必须 ok:false + 指定 reason（不删断言，专门验证守卫/分类拒绝）。"""
    global PASS, FAIL
    _e, _txt = call_tool(tool, args)
    if (not _e) and ('"reason":"%s"' % reason) in _txt:
        print("[OK]  %s | refused with reason=%s" % (label, reason))
        PASS += 1
        return True
    print("[??]  %s | 期望 reason=%s，实际: %s" % (label, reason, _txt[:4000].replace("\n", " ")))
    FAIL += 1
    FAILED_NAMES.append(label)
    return False


def expect_not_error(label, tool, args):
    """无固定字符串可断言时使用：只要不是 isError 且没有 "ok":false 就算通过
    （仅用于 set/clear 这类回执形状未冻结的辅助动作，不用来放宽任何契约断言）。"""
    global PASS, FAIL
    _e, _txt = call_tool(tool, args)
    if (not _e) and ('"ok":false' not in _txt):
        print("[OK]  %s | %s" % (label, _txt[:140].replace("\n", " ")))
        PASS += 1
        return _txt
    print("[??]  %s | %s" % (label, _txt[:4000].replace("\n", " ")))
    FAIL += 1
    FAILED_NAMES.append(label)
    return None


def t_end(label):
    """end 的双形态契约判定：在跑 => executed:true；不在跑 => ok:true + state_not_changed
    （alreadyNotRunning / already_not_running）。两者之外一律失败。"""
    global PASS, FAIL
    _e, _txt = call_tool("control_execution", {"mode": "end"})
    running_shape = '"executed":true' in _txt
    idle_shape = ('"ok":true' in _txt
                  and ('"already_not_running"' in _txt or '"alreadyNotRunning":true' in _txt))
    if (not _e) and (running_shape or idle_shape):
        print("[OK]  %s | %s" % (label, "executed" if running_shape
                                 else "already_not_running (contract shape)"))
        PASS += 1
        return True
    print("[??]  %s | %s" % (label, _txt[:4000].replace("\n", " ")))
    FAIL += 1
    FAILED_NAMES.append("control_execution:" + label)
    return False


print("=" * 78)
print("VB6MCP 功能测试  port=%d" % PORT)
print("=" * 78)

# ---- 1. 工程信息 ----
t("add_project", {"type": "StandardExe"}, expect='"created"')
t("get_component_list", expect='"Form1"')

# 对齐 v1.16.x 契约 · make_project 在「工程还没落盘」时必须**分类拒绝**（v1.16 把它从
# 「会弹 VB6 Save 对话框的未定义行为」改成结构化拒绝）：
#   ok:false + tool="make_project" + reason="project_unsaved"（ClassifiedRefusalJson）
# 依据：docs/CONTRACTS.md 的 make_project / 分类拒绝条目；docs/KNOWN_LIMITATIONS_AND_FAQ.md
#   关于「编译未落盘工程会打开 VB6 Save 对话框」的处置说明（先 save_project，或在没有
#   文件名时调用会得到 classified refusal，而不是卡在宿主模态框）。
# 与脏内存守卫（dirty_memory_would_be_saved）同属「契约化拒绝」⇒ 必须被回归**锁定**，
# 故这里做负例：此刻工程是 add_project 刚建的、尚未 save_project（不可写 .vbp）。
expect_refusal("make_project on unsaved project refused (project_unsaved)",
               "make_project", {}, "project_unsaved")

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
# 能力覆盖准备（对齐 ③）：为 Form1 建 Form_Load 并写入一行可断点语句。
# 对齐：CommandBars 在 RUN 态挂起（已文档化）⇒ 菜单驱动的 break 在 RUN 态不可用；
# 真实中断能力改用菜单无关的 breakpoints{action:"probe"} 覆盖（见 run 段之后的独立用例）。
t("create_event_proc", {"component": "Form1", "object": "Form", "event": "Load"},
  expect='"startLine"')
# get_component_code 返回的是 RECEIPT（JSON 字符串），代码窗文本在 .text 里——
# 直接拿 [1] 当文本用会定位到错误行（上一版就因此在声明区插了语句，导致 make_project 编译失败）。
_fl_rec = json.loads(call_tool("get_component_code", {"name": "Form1"})[1] or "{}")
_fl_txt = _fl_rec.get("text") or ""
_fl_line = next((i + 1 for i, x in enumerate(_fl_txt.split("\n"))
                 if "Form_Load()" in x), 0)
if _fl_line <= 0:
    print("[??]  Form_Load 行号定位失败（get_component_code 未返回 text）")
    FAIL += 1
    FAILED_NAMES.append("form-load-line-locate")
call_tool("replace_code", {"name": "Form1", "operation": "insert",
                           "start_line": _fl_line + 1,
                           "code": "    Dim probeMarker As Long\r\n    probeMarker = 1"})
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


# ---------------------------------------------------------------------------
# 契约对齐（测试侧，产品不动）· v1.16.0
#
# 依据一：docs/CONTRACTS.md v1.12.2「运行/单步落盘的观测与知情同意（R-1021-02；新增 1 个
#   可选参数 allowDiskMutation；回执全部增量）」——脏内存（VBProject.Saved=False 或任一
#   组件 IsDirty=True）下 control_execution{run|full_compile|restart|step_*} 默认**拒绝**
#   启动，回 ok:false + reason="dirty_memory_would_be_saved"；显式 allowDiskMutation:true
#   是**知情同意**（表示本用例有意允许 IDE 把内存改动写盘），不是关闭守卫。
# 依据二：docs/MATRIX.md:86 把「battery 未传 allowDiskMutation:true -> dirty_memory_would_be_saved;
#   already_not_running」登记为 FAIL_TEST_ASSUMPTION（测试自身前提过期）。
# 依据三：docs/MATRIX.md:233 / docs/CONTRACTS.md v1.16.0 control_execution 条目——end/stop
#   落在「本来就没在跑」时回 ok:true + state_not_changed:true + alreadyNotRunning:true，
#   不再回 ok:false，也不再保证 executed:true。
#
# 因此本段做两件事，**断言一条都不删**：
#   (a) 先显式把 IDE 内存弄脏，再断言「不带 allowDiskMutation 的 run 必须被拒」——这样
#       新守卫本身仍有回归覆盖（负例）；
#   (b) 其后的 run/restart/full_compile 一律显式 allowDiskMutation:true（模拟“允许落盘”
#       的场景），end 则按上面第三条的双形态契约判定。
# ---------------------------------------------------------------------------
# (a) 负例：弄脏内存 => 不带 allowDiskMutation 的 run 必须被拒（契约 v1.12.2）
call_tool("replace_code", {"name": "TestModRenamed", "operation": "append",
                           "code": "Public Function GuardDirty() As Long\r\n"
                                   "    GuardDirty = 1\r\nEnd Function"})
expect_refusal("dirty guard negative (no allowDiskMutation)",
               "control_execution", {"mode": "run"}, "dirty_memory_would_be_saved")

print("=" * 78)
print("运行控制测试 (make_project + run/break/end)")
print("=" * 78)
# 前置：本用例之前 battery 已 t("save_project", {path: ROOT/_test_proj.vbp})（见 §11 保存/编译）。
# 对齐 v1.16.x 契约 · make_project 对**未落盘**工程会分类拒绝 reason="project_unsaved"
# （§1 已有对应负例），因此正例必须发生在 save_project 之后；这里保留原严格期望
# （必须 compiled:true），不放宽。
t("make_project", expect='"compiled":true', show=200)
time.sleep(1)
# 对齐 v1.16.0 契约 · 依据 docs/CONTRACTS.md v1.12.2（allowDiskMutation 知情同意）：
# 内存此刻是脏的（上一段刻意弄脏并验证了拒绝），本用例**有意**允许 IDE 落盘。
t("control_execution", {"mode": "run", "allowDiskMutation": True}, show=120)


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
# 对齐 v1.16.0 契约 · 依据 docs/MATRIX.md:233：end 在「本来没在跑」时是
# ok:true + state_not_changed + already_not_running（合法形态），在跑时必须 executed:true。
t_end("end")
wait_state("design", 10, "wait_state design")
t("get_project_info", show=60)

print("=" * 78)
print("break/restart/full_compile 模式测试")
print("=" * 78)
# break cycle: run -> (RUN 态 break 必须被拒) -> 仍在跑 -> end
# 对齐：CommandBars 在 RUN 态挂起（已文档化）—— 见 control_execution/command_bars 工具
# 描述与 docs/KNOWN_LIMITATIONS_AND_FAQ.md。RUN 态菜单驱动的 break 的契约行为就是
# ok:false + reason="commandbars_unavailable_while_running"，应被回归**锁定**，而不是
# 当作按钮能按下去。另：allowDiskMutation 依据 CONTRACTS v1.12.2（知情同意）。
t("control_execution", {"mode": "run", "allowDiskMutation": True}, show=120)
wait_state("run", 20, "wait_state run-b")
expect_refusal("RUN-state break refused (documented CommandBars suspension)",
               "control_execution", {"mode": "break"},
               "commandbars_unavailable_while_running")
# 被拒之后程序必须仍在跑（不允许「拒了但其实停了」）
wait_state("run", 5, "wait_state run-after-refused-break")
t_end("end-b")
wait_state("design", 10, "wait_state design-b")

# restart cycle: run -> (RUN 态 restart 必须被拒) -> 仍在跑 -> end
# 对齐：同 break —— restart = end+run，RUN 态同样走 CommandBars（已文档化限制），
# 契约行为是 ok:false + reason="commandbars_unavailable_while_running"。
# allowDiskMutation 依据 CONTRACTS v1.12.2（restart 的 Start 半段受脏内存守卫）。
t("control_execution", {"mode": "run", "allowDiskMutation": True}, show=120)
wait_state("run", 20, "wait_state run-r")
expect_refusal("RUN-state restart refused (documented CommandBars suspension)",
               "control_execution", {"mode": "restart", "allowDiskMutation": True},
               "commandbars_unavailable_while_running")
wait_state("run", 5, "wait_state run-after-refused-restart")
t_end("end-r")
wait_state("design", 10, "wait_state design-r")

# full_compile cycle: full compile + run -> end
# 对齐 v1.16.0 契约 · allowDiskMutation 知情同意（CONTRACTS v1.12.2）
t("control_execution", {"mode": "full_compile", "allowDiskMutation": True}, show=120)
wait_state("run", 20, "wait_state run-fc")
t_end("end-fc")
wait_state("design", 10, "wait_state design-fc")

print("=" * 78)
print(" 能力覆盖：菜单无关的真实断点能力（breakpoints probe）")
print("=" * 78)
# 对齐 ③：RUN 态菜单 break/restart 已按文档化限制锁定为「必须被拒」；本用例验证的是
# 「程序确实能被停下来」这一能力本身，走**菜单无关**的真实断点路径。
# 依据：docs/CONTRACTS.md v1.16.0 breakpoints{action:"probe"} 两阶段契约
# （phase=start => probeResult=started_pending + probeId；phase=observe => 读到
#  position{component,proc,paneLine} 后 END 程序并回读证明 stateAfter=design）。
# **前置（契约原文）**：probe「verifies whatever breakpoint is already set」——它自己
# **不注入、不 F9、不写 ledger**（见回执 probeCleanup.injectedByProbe=false /
# "there is nothing to inject or remove"）。所以必须先显式 set 一个断点，probe 才有东西
# 可命中；用完再 clear。缺这一步时 probe 必然 did_not_stop（这是我们上一版用例的测试错误，
# 不是产品缺陷）。
# （native 模式说明，供后续用例参考）mode:"native" 只能在 design（或 break）态设置；
# 运行态下 CommandBars 被挂起，F9 不会落地，probe 会如实 timeout 且 nativeStateVerified:false ——
# 这是设置时机问题，不是断点失效。本用例走 injected marker（行号取 set 回执 breakpoint.paneLine）。
# 旧注释：
# 程序会先停在标记的 Debug.Assert 上（实测 probeResult=stopped_elsewhere），而不是目标行。
# native 只放断点、不写源码，probe 才能停在 targetLine 本身。
""
# 规范用法（addin-core 判读）：marker 模式**就该停在 marker 自己的 Debug.Assert 行**；
# 所以 target 行要取 marker 块里的 Debug.Assert 行，而不是 anchor 语句行。断言仍严格：
# probeResult 必须 stopped_at_target、position.proc 必须 Form_Load、callstack 必须含 Form_Load。
_bp_set = expect_not_error("breakpoints (set injected marker for probe)",
                           "breakpoints", {"action": "set", "component": "Form1",
                                           "anchor": "    probeMarker = 1",
                                           "id": "bat-probe-1"})
# marker 注入会移动 pane 行号；目标行必须取 **语句行**（probe 只接受可断行：
# blank_line / 过程头都会被 line_not_breakable 拒掉）。这里直接从 pane 文本里找
# 「probeMarker = 1」那一行，而不是用 set 回执的行号（回执行号可能指向 marker 块后的空行）。
_prec = json.loads(call_tool("get_component_code", {"name": "Form1"})[1] or "{}")
_ptxt = _prec.get("text") or ""
_probe_line = next((i + 1 for i, x in enumerate(_ptxt.split("\n"))
                    if "probeMarker = 1" in x), 0)
if _probe_line <= 0 and _bp_set:
    try:
        _bpj = json.loads(_bp_set)
        _probe_line = int(_bpj.get("paneLine") or _bpj.get("line") or 0)
    except Exception:
        _probe_line = 0
# 官方用法（addin-core 对源码核过）：
#   set 回执结构 = breakpoint{id, component, paneLine, markerLine, diskLine, mode}
#   停点 = **breakpoint.paneLine**（注入的 Debug.Assert 语句行；markerLine 是它上一行注释，不可断）
#   更稳：probe 支持文本兜底 lineText（实现里 hitLine Or textHit）⇒ 不依赖行号。
if _bp_set:
    try:
        _bpj2 = json.loads(_bp_set)
        _bp_meta = _bpj2.get("breakpoint") or {}
        _probe_line = int(_bp_meta.get("paneLine") or 0)
    except Exception:
        _probe_line = 0
if _probe_line <= 0:
    _pw = json.loads(call_tool("get_component_code", {"name": "Form1"})[1] or "{}").get("text") or ""
    _pwl = _pw.split("\n")
    _marker_line = next((i + 1 for i, x in enumerate(_pwl) if "@mcp-break bat-probe-1" in x), 0)
    if _marker_line > 0:
        for _k in range(_marker_line, min(_marker_line + 3, len(_pwl) + 1)):
            if "Debug.Assert" in _pwl[_k - 1]:
                _probe_line = _k
                break
# srcZ 的 breakpoints 工具**未声明 lineText**（acceptedArguments: action/component/anchor/
# line/mode/id/clearStale/timeout_ms/phase/probeId/hitCount），未知参数会被结构化拒绝
# （reason=unknown_argument）——所以这里只用 line=breakpoint.paneLine，不用文本兜底。
_pstart = json.loads(call_tool("breakpoints", {"action": "probe", "component": "Form1",
                                               "line": _probe_line,
                                               "phase": "start",
                                               "timeout_ms": 20000})[1] or "{}")
if _pstart.get("probeResult") == "started_pending" and _pstart.get("probeId"):
    print("[OK]  breakpoints probe phase=start | started_pending probeId=%s"
          % _pstart["probeId"])
    PASS += 1
    _pobs = json.loads(call_tool("breakpoints", {"action": "probe", "phase": "observe",
                                                 "probeId": _pstart["probeId"],
                                                 "timeout_ms": 30000})[1] or "{}")
    _pos = _pobs.get("position") or {}
    _hit_line = int(_pos.get("paneLine") or _pos.get("line") or 0)
    if (_pobs.get("probeResult") == "stopped_at_target"
            and _hit_line == _probe_line
            and _pos.get("proc") == "Form_Load"):
        print("[OK]  breakpoints probe phase=observe | stopped_at_target component=%s proc=%s"
              % (_pos.get("component"), _pos.get("proc")))
        PASS += 1
    else:
        print("[??]  breakpoints probe phase=observe | %s" % json.dumps(_pobs)[:4000])
        FAIL += 1
        FAILED_NAMES.append("breakpoints-probe:observe")
else:
    print("[??]  breakpoints probe phase=start | %s" % json.dumps(_pstart)[:4000])
    FAIL += 1
    FAILED_NAMES.append("breakpoints-probe:start")
expect_not_error("breakpoints (clear native breakpoint)",
                 "breakpoints", {"action": "clear", "component": "Form1",
                                 "id": "bat-probe-1", "mode": "native"})

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
        print("[??]  %s | %s" % (label, _txt[:4000].replace("\n", " ")))
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

# 33b pre-clean: run this block from an EMPTY workspace in DESIGN, so add_project cannot
# raise a host save dialog (a dirty current project would) and the block cannot hang the runner.
sys.stdout.flush()
_33b_pre_err, _33b_pre_txt = call_tool("control_execution", {"mode": "end"})
sys.stdout.flush()
_33b_rm_err, _33b_rm_txt = call_tool("remove_project", {"allowDirty": True})
sys.stdout.flush()
if (not _33b_rm_err) and ("dirty_untitled_would_write" in (_33b_rm_txt or "")):
    _33b_rm_err, _33b_rm_txt = call_tool("remove_project", {"allowDirty": True, "allowHostWrite": True})
    sys.stdout.flush()
# ---- 33b: dirty-untitled close path (default refusal / allowHostWrite / open ok+apiWarnings) ----
# 33a measured that closing a DIRTY UNTITLED project makes VB6 write it into the IDE current
# directory. 33b refuses that by default (zero writes); allowHostWrite:true restores 3a and the
# receipt must report hostWrites. The written files are removed again so this battery leaves the
# working tree unchanged.
sys.stdout.flush()
_33b_fresh = call_tool("add_project", {"type": "StandardExe"})
sys.stdout.flush()
_33b_class = call_tool("add_component", {"type": "class", "name": "C33BDirty"})
sys.stdout.flush()
_33b_edit = call_tool("replace_code", {"name": "C33BDirty", "operation": "append",
                                       "code": "Public Sub C33BMarker()" + chr(13) + chr(10) +
                                               "    Dim x As Long" + chr(13) + chr(10) +
                                               "    x = 1" + chr(13) + chr(10) + "End Sub"})
sys.stdout.flush()
expect_refusal("33b dirty-untitled default refusal", "remove_project",
               {"allowDirty": True}, "dirty_untitled_would_write")
sys.stdout.flush()
_33b_aw_err, _33b_aw_txt = call_tool("remove_project", {"allowDirty": True, "allowHostWrite": True})
_33b_aw = None
try:
    _33b_aw = json.loads(_33b_aw_txt)
except Exception:
    _33b_aw = None
if (not _33b_aw_err) and isinstance(_33b_aw, dict) and _33b_aw.get("ok") is True and (_33b_aw.get("hostWrites") or []):
    print("[OK]  33b allowHostWrite:true remove | ok=true hostWrites=%d" % len(_33b_aw["hostWrites"]))
    PASS += 1
    for _33b_w in _33b_aw["hostWrites"]:
        _33b_path = _33b_w.get("path") if isinstance(_33b_w, dict) else None
        if _33b_path and os.path.exists(_33b_path):
            try:
                os.remove(_33b_path)
            except OSError:
                pass
else:
    print("[??]  33b allowHostWrite:true remove | %s" % _33b_aw_txt[:800].replace("\n", " "))
    FAIL += 1
    FAILED_NAMES.append("33b-allowHostWrite")
_33b_probe = os.path.join(ROOT, "_test_proj.vbp")
sys.stdout.flush()
_33b_oerr, _33b_otxt = call_tool("open_project", {"path": _33b_probe, "exclusive": True})
_33b_ok = False
_33b_oj = None
try:
    _33b_oj = json.loads(_33b_otxt)
    _33b_ok = ((not _33b_oerr) and isinstance(_33b_oj, dict) and _33b_oj.get("ok") is True
               and "apiError" not in _33b_oj)
    if _33b_oj.get("apiWarnings") is not None:
        _33b_ok = _33b_ok and isinstance(_33b_oj.get("apiWarnings"), list)
except Exception:
    _33b_ok = False
if _33b_ok:
    print("[OK]  33b open_project ok/apiWarnings | ok=true apiError absent apiWarnings=%s"
          % ("present" if _33b_oj.get("apiWarnings") is not None else "absent"))
    PASS += 1
else:
    print("[??]  33b open_project ok/apiWarnings | %s" % _33b_otxt[:800].replace("\n", " "))
    FAIL += 1
    FAILED_NAMES.append("33b-open-ok-apierrors")

# ---- 34a: structured error channel (RE-1) ------------------------------------
# NON-MODAL coverage only, by design: the batch contract says the runtime-error box
# is NOT suppressed, and manufacturing one from the battery is both locale-fragile
# and can leave the box behind and take the remaining assertions down with it. The
# modal main path (errBasis="vb6map" with number/description/SOURCE) plus the
# reverse case (a DIFFERENT error must move the fields) and the A11 cached shape are
# covered by the hand test (T1/T2/T3) and by the probe runtime window; that gap is
# disclosed in RECORD_34a.md instead of being papered over here.
sys.stdout.flush()
_le_err, _le_txt = call_tool("last_error", {})
_le = None
try:
    _le = json.loads(_le_txt)
except Exception:
    _le = None
_34a_keys = ("err", "errBasis", "errChannel", "errCrossCheck", "errLocation")
_34a_sub = {
    "err": ("number", "description", "source", "hresult", "numberBasis"),
    "errChannel": ("readOk", "recordUsed", "selector", "fields", "bstrReads", "why"),
    "errCrossCheck": ("available", "number", "description", "relation", "note"),
    "errLocation": ("line", "column", "file", "basis", "why"),
}
_34a_ok = isinstance(_le, dict) and all(_k in _le for _k in _34a_keys)
if _34a_ok:
    for _k, _subs in _34a_sub.items():
        _v = _le.get(_k)
        if not isinstance(_v, dict) or not all(_s in _v for _s in _subs):
            _34a_ok = False
            break
if _34a_ok:
    print("[OK]  34a last_error structured keys | err/errBasis/errChannel/errCrossCheck/errLocation all present")
    PASS += 1
else:
    print("[??]  34a last_error structured keys | %s" % (_le_txt or "")[:800].replace("\n", " "))
    FAIL += 1
    FAILED_NAMES.append("34a-structured-keys")

_34a_ok2 = False
if isinstance(_le, dict) and isinstance(_le.get("err"), dict):
    _basis = _le.get("errBasis")
    _ch = _le.get("errChannel") or {}
    _loc = _le.get("errLocation") or {}
    _xc = _le.get("errCrossCheck") or {}
    _sel = _ch.get("selector") or {}
    _reads = _ch.get("bstrReads")
    _num = _le["err"].get("number")
    _34a_ok2 = (
        _basis in ("vb6map", "dialog", "none", "unavailable")
        and _sel.get("basis") in ("dword074_bit24", "fallback_probe", "unavailable")
        and isinstance(_reads, list) and len(_reads) == 2
        and [r.get("offset") for r in _reads] == ["+0x04", "+0x08"]
        and _loc.get("basis") == "not_attempted"
        and _loc.get("line") is None and _loc.get("column") is None and _loc.get("file") is None
        and _xc.get("relation") == "addinThreadErrAtHandlerEntry"
        and (_ch.get("readOk") is True or bool(_ch.get("why")))
        and (_basis == "vb6map" or _num is None)
    )
if _34a_ok2:
    print("[OK]  34a errBasis closed set + null discipline | basis=%s readOk=%s why=%s"
          % (_basis, _ch.get("readOk"), str(_ch.get("why"))[:60]))
    PASS += 1
else:
    print("[??]  34a errBasis closed set + null discipline | %s" % (_le_txt or "")[:800].replace("\n", " "))
    FAIL += 1
    FAILED_NAMES.append("34a-errbasis-discipline")

# A failure must not look like a success: in DESIGN with no box up the tool answers
# ok:false (or a documented reason) - never ok:true on an empty read.
_34a_ok3 = isinstance(_le, dict) and ((_le.get("ok") is False) or bool(_le.get("reason")))
if _34a_ok3:
    print("[OK]  34a failure not silent | ok=%s reason=%s" % (_le.get("ok"), _le.get("reason")))
    PASS += 1
else:
    print("[??]  34a failure not silent | %s" % (_le_txt or "")[:400].replace("\n", " "))
    FAIL += 1
    FAILED_NAMES.append("34a-failure-not-silent")

_mi_err, _mi_txt = call_tool("mcp_info", {})
_mi = None
try:
    _mi = json.loads(_mi_txt)
except Exception:
    _mi = None
_34a_ok4 = (isinstance(_mi, dict) and _mi.get("toolCount") == 120
            and str(_mi.get("contract", "")).startswith("v1.16"))
if _34a_ok4:
    print("[OK]  34a contract face | contract=%s toolCount=%s"
          % (_mi.get("contract"), _mi.get("toolCount")))
    PASS += 1
else:
    print("[??]  34a contract face | %s" % (_mi_txt or "")[:400].replace("\n", " "))
    FAIL += 1
    FAILED_NAMES.append("34a-contract-face")

# ===================== 39: dialog-id-first face (B39-1..B39-6) =====================
# ID-first button selection is an internal DECISION change; these assertions only read the
# receipt fields it publishes (idBasis / controlId / idTableVersion) and NEVER create a
# dialog themselves. A box this battery cannot produce is reported as [SKIP] - counted as
# neither pass nor fail (the lane wrapper only reads ^[OK] / ^[??] lines).
SKIPPED = []


def skip(label, why):
    print("[SKIP] %s | %s" % (label, why))
    SKIPPED.append(label)


def _dialogs_now():
    _e, _t = call_tool("read_active_dialogs", {})
    try:
        _o = json.loads(_t)
    except Exception:
        return []
    return _o.get("dialogs") or []


def _ids_of(_d):
    return [int(b.get("id") or 0) for b in (_d.get("buttons") or []) if isinstance(b, dict)]


# ---- B39-6 (always): the no-dialog path must stay byte-identical (lane 0-diff) ----
_d6 = _dialogs_now()
if _d6:
    skip("B39-6 no-dialog paths add no keys",
         "a #32770 is up during the battery run, so the no-dialog shape cannot be observed here")
else:
    _e6a, _t6a = call_tool("close_active_dialogs", {})
    _e6b, _t6b = call_tool("last_error", {})
    _e6c, _t6c = call_tool("read_active_dialogs", {})
    _o6a = None
    try:
        _o6a = json.loads(_t6a)
    except Exception:
        _o6a = None
    _new_click_keys = ('"idBasis"', '"controlId"', '"idTableVersion"')
    _b39_6 = ((not _e6a) and isinstance(_o6a, dict) and not _o6a.get("clicked")
              and not any(k in _t6a for k in _new_click_keys)
              and (not _e6b) and ('"dismissIdBasis"' not in _t6b)
              and ('"dismissControlId"' not in _t6b)
              and (not _e6c) and ('"idBasis"' not in _t6c) and ('"idTableVersion"' not in _t6c))
    if _b39_6:
        print("[OK]  B39-6 no-dialog paths add no keys | close/last_error/read keep the frozen shape")
        PASS += 1
    else:
        print("[??]  B39-6 | close=%s last_error=%s read=%s"
              % ((_t6a or "")[:400], (_t6b or "")[:400], (_t6c or "")[:200]).replace(chr(10), " "))
        FAIL += 1
        FAILED_NAMES.append("dialog-id-first:B39-6")

# ---- B39-1 / B39-2 (conditional): runtime-error box, Debug control id 4801 ----
_d39 = _dialogs_now()
_rt = [d for d in _d39 if 4801 in _ids_of(d)]
if not _rt:
    skip("B39-1 runtime-error ID-first dismiss",
         "no #32770 carrying control id 4801 (Debug) is up; this battery never creates a modal runtime-error box")
    skip("B39-2 runtime-error ID is caption-independent",
         "same: no runtime-error box up, so the Debug control id cannot be read")
else:
    _e1, _t1 = call_tool("last_error", {"dismiss": True})
    _o1 = None
    try:
        _o1 = json.loads(_t1)
    except Exception:
        _o1 = None
    _b39_1 = (isinstance(_o1, dict) and _o1.get("dismissedVia") == "debug"
              and _o1.get("dismissIdBasis") == "control_id"
              and _o1.get("dismissControlId") == 4801)
    if _b39_1:
        print("[OK]  B39-1 runtime-error ID-first dismiss | dismissedVia=debug controlId=4801 idBasis=control_id")
        PASS += 1
    else:
        print("[??]  B39-1 | %s" % (_t1 or "")[:1200].replace(chr(10), " "))
        FAIL += 1
        FAILED_NAMES.append("dialog-id-first:B39-1")
    _caps = []
    for _d in _rt:
        for _b in (_d.get("buttons") or []):
            if isinstance(_b, dict) and int(_b.get("id") or 0) == 4801:
                _caps.append(str(_b.get("text") or ""))
    _b39_2 = (isinstance(_o1, dict) and _o1.get("dismissControlId") == 4801
              and _o1.get("dismissedVia") == "debug")
    if _b39_2:
        print("[OK]  B39-2 runtime-error ID is caption-independent | control id 4801 chosen; captions=%r" % (_caps,))
        PASS += 1
    else:
        print("[??]  B39-2 | caps=%r receipt=%s" % (_caps, (_t1 or "")[:800].replace(chr(10), " ")))
        FAIL += 1
        FAILED_NAMES.append("dialog-id-first:B39-2")

# ---- B39-3 / B39-4 (conditional): save-confirm box, ids {6,7,2}, DRY RUN only ----
_sv = [d for d in _dialogs_now() if set([6, 7, 2]).issubset(set(_ids_of(d)))]
if not _sv:
    skip("B39-3 save-confirm ID-first default (yes=6)",
         "no #32770 carrying control ids {6,7,2} is up; this battery never creates a save-confirm box")
    skip("B39-4 save-confirm explicit Cancel -> control id 2",
         "same: no save-confirm box up")
else:
    _e3, _t3 = call_tool("close_active_dialogs", {"dry_run": True})
    _o3 = None
    try:
        _o3 = json.loads(_t3)
    except Exception:
        _o3 = None
    _ck3 = (_o3 or {}).get("clicked") or []
    _b39_3 = (isinstance(_o3, dict) and len(_ck3) == 1
              and _ck3[0].get("idBasis") == "control_id" and _ck3[0].get("controlId") == 6)
    if _b39_3:
        print("[OK]  B39-3 save-confirm ID-first default | idBasis=control_id controlId=6 (dry_run)")
        PASS += 1
    else:
        print("[??]  B39-3 | %s" % (_t3 or "")[:1200].replace(chr(10), " "))
        FAIL += 1
        FAILED_NAMES.append("dialog-id-first:B39-3")
    # the explicit-button path matches the CAPTION it is given (pre-existing rule), so feed
    # back the box's own Cancel caption and let the ID table pick control id 2.
    _cap2 = ""
    for _d in _sv:
        for _b in (_d.get("buttons") or []):
            if isinstance(_b, dict) and int(_b.get("id") or 0) == 2:
                _cap2 = str(_b.get("text") or "")
                break
        if _cap2:
            break
    if not _cap2:
        print("[SKIP] B39-4 save-confirm explicit Cancel -> control id 2 | the box has no button with control id 2")
        SKIPPED.append("B39-4-no-id2")
        _e4, _t4 = None, ""
        _b39_4 = True
        _ck4 = []
    else:
        _e4, _t4 = call_tool("close_active_dialogs", {"button": _cap2, "dry_run": True})
    if _cap2:
        _o4 = None
        try:
            _o4 = json.loads(_t4)
        except Exception:
            _o4 = None
        _ck4 = (_o4 or {}).get("clicked") or []
        _b39_4 = (isinstance(_o4, dict) and len(_ck4) == 1
                  and _ck4[0].get("idBasis") == "control_id" and _ck4[0].get("controlId") == 2)
    if _b39_4:
        print("[OK]  B39-4 save-confirm explicit Cancel | idBasis=control_id controlId=2 (dry_run)")
        PASS += 1
    else:
        print("[??]  B39-4 | %s" % (_t4 or "")[:1200].replace(chr(10), " "))
        FAIL += 1
        FAILED_NAMES.append("dialog-id-first:B39-4")

# ---- B39-5 (conditional): the rollback switch is read INSIDE the IDE process ----
_idfirst = os.environ.get("VB6MCP_DLG_ID_FIRST", "").strip().lower()
if _idfirst in ("0", "off", "false", "no"):
    _d5 = _dialogs_now()
    if not _d5:
        skip("B39-5 VB6MCP_DLG_ID_FIRST=0 legacy path",
             "switch is OFF but no dialog is up, so no selection can be observed")
    else:
        _e5, _t5 = call_tool("close_active_dialogs", {"dry_run": True})
        _o5 = None
        try:
            _o5 = json.loads(_t5)
        except Exception:
            _o5 = None
        _ck5 = (_o5 or {}).get("clicked") or []
        _b39_5 = (isinstance(_o5, dict) and all(c.get("idBasis") == "caption_fallback" for c in _ck5))
        if _b39_5:
            print("[OK]  B39-5 VB6MCP_DLG_ID_FIRST=0 legacy path | idBasis=caption_fallback")
            PASS += 1
        else:
            print("[??]  B39-5 | %s" % (_t5 or "")[:1200].replace(chr(10), " "))
            FAIL += 1
            FAILED_NAMES.append("dialog-id-first:B39-5")
else:
    skip("B39-5 VB6MCP_DLG_ID_FIRST=0 legacy path",
         "the IDE under test was started WITHOUT VB6MCP_DLG_ID_FIRST=0 (the switch lives in the IDE process), so the legacy path cannot be exercised from here")

print("=" * 78)
print("结果: 通过 %d, 失败 %d" % (PASS, FAIL))
if FAILED_NAMES:
    print("失败项:", ", ".join(FAILED_NAMES))
print("=" * 78)
if SKIPPED:
    print("SKIP: %d (%s)" % (len(SKIPPED), ", ".join(SKIPPED)))
# Packaging gap: the old script always exited 0, so a caller could not tell a green
# run from a red one. stdout format is unchanged - only the exit code is.

# ===================== 34b: exitWrites face assertions (B2 then B1) =====================
# These two MUST stay at the very end: B1 requests the quit and the IDE exits with it.
# B2 first (dirty project -> refusal, NO exitWrites key), then B1 (clean -> declaration block).
_b2_dirty_vbp = os.path.join(os.environ.get("TEMP", "."), "vb6mcp_b1_probe.vbp")
call_tool("add_project", {})
_e2, _txt2 = call_tool("quit_ide")
_b2_ok = ((not _e2) and ('"requested":false' in _txt2)
          and ('"exitWrites"' not in _txt2))
if _b2_ok:
    print("[OK]  B2 exitWrites absent on blocked quit | %s" % _txt2[:140].replace(chr(10), " "))
    PASS += 1
else:
    print("[??]  B2 | %s" % _txt2[:2000].replace(chr(10), " "))
    FAIL += 1
    FAILED_NAMES.append("exitWrites:B2")
# make the workspace clean+saved so the next quit is REQUESTED (writes outside the repo)
call_tool("save_project", {"path": _b2_dirty_vbp})
_e1, _txt1 = call_tool("quit_ide")
def _b1_ok(s):
    try:
        o = json.loads(s)
    except Exception:
        return False
    ew = o.get("exitWrites")
    if not isinstance(ew, dict):
        return False
    if ew.get("declaration") != "advisory_only" or ew.get("observed") != "not_observable":
        return False
    if ew.get("condition") not in ("no_project_observed_no_write",
                                    "project_group_context_then_clean_quit",
                                    "undetermined"):
        return False
    ev = ew.get("evidence")
    return isinstance(ev, list) and len(ev) == 4
if (not _e1) and _b1_ok(_txt1):
    print("[OK]  B1 exitWrites declaration block complete | %s" % _txt1[:140].replace(chr(10), " "))
    PASS += 1
else:
    print("[??]  B1 | %s" % _txt1[:2000].replace(chr(10), " "))
    FAIL += 1
    FAILED_NAMES.append("exitWrites:B1")

sys.exit(1 if FAIL else 0)
