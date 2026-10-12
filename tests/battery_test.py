"""Comprehensive tool test battery against the live VB6MCP add-in."""
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
t("replace_code", {"name": "TestMod1", "operation": "replace",
                   "start_line": 2, "end_line": 2,
                   "code": "' replaced comment"},
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
print("结果: 通过 %d, 失败 %d" % (PASS, FAIL))
if FAILED_NAMES:
    print("失败项:", ", ".join(FAILED_NAMES))
print("=" * 78)
