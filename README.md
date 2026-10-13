# VB6MCP — 让 VB6 IDE 成为 MCP 服务器的插件

> **中文** | [English](README.en.md)

VB6MCP 是一个 **VB6 插件（Add-In）**：把它加载到 VB6 IDE 之后，IDE 就变成一个
[MCP (Model Context Protocol)](https://modelcontextprotocol.io) 服务器。
Claude Desktop / Cursor / Codex 等 AI 编程助手即可直接读取、搜索、修改当前打开的
VB6 工程，并控制编译与运行。

```
 MCP 客户端 (Claude Desktop / Cursor / ...)
        │  stdio (JSON-RPC, 每行一条)
        ▼
 bridge/mcp_stdio_bridge.py   (Python 3, 仅标准库)
        │  HTTP POST /rpc  (http://127.0.0.1:8766)
        ▼
 VB6MCP.dll  ← 运行在 VB6.EXE 进程内
   ├─ Winsock 控制: 本地 HTTP 服务器 (也支持 /sse + /message)
   ├─ 自研 JSON 解析/序列化 (JsonLib.bas, 零外部依赖)
   └─ VBIDE 扩展对象模型: 读取/修改代码、保存、编译、运行控制
```

## 这是一个 AI 辅助构建的项目

本项目**由 AI 辅助构建，并非逐行人工编写**。把这一点讲清楚，对使用者与贡献者都重要。

- **代码**：绝大部分源码由 AI 编程助手在人工监督下生成与修改。人类负责需求、取舍、验收，
  以及**真机验证**——本项目所有关键结论都要求在真实 VB6 上取到可复现的证据。
- **方法**：开发过程刻意按「可验证」组织——任务以卡片下发、每条结论都须附机器可读证据、
  关键改动由**独立验证方**复核（实现者与验证者分开）、工具面有**冻结基线门禁**、工具做不到时
  必须给出**机器可读的原因**而不是悄悄降级。
- **人工复核程度不作保证**：请不要假设每一行都被人类逐行审阅过。用于生产或二次开发前请自行
  评估；发现问题请开 issue（见 CONTRIBUTING）。
- **已知局限**：见「七、已知限制」，尤其是**本项目在中文版 VB6（zh-CN）上开发与验证**这一点。
## 为什么做这个（项目背景）

VB6 在 2008 年就停止演进了，但它写的系统还在跑：工厂、财务、医院、政府内部系统里有大量
陈旧 VB6 工程仍需要维护，而维护者面对的是一套**没有任何现代工具链**的环境——没有语言
服务器、没有可靠的静态检查、编译报错只能手动点、改代码靠自己不手抖。

AI 编程助手已经能很好地处理现代语言，却**够不到 VB6**：它们既看不见 IDE 里打开的工程，
也点不了那个「生成工程」对话框。本项目就是补这个缺口。

**做法**：不另造编辑器，而是**把 VB6 IDE 本身变成一个 MCP 服务器**——加载一个插件，IDE
就对外提供 111 个工具：读/搜/改代码、静态检查、锚点编辑与安全事务、编译、运行与单步、
弹窗处理、窗体与项目元数据、快照与回滚、打包。**因为工具就住在 IDE 进程里**，它能做
进程外方案做不到的事：真的去点那个对话框、真的按运行、真的读断点处的局部变量。

**目标**：让维护陈旧 VB6 的人（包括作者自己）用上现代 AI 工作流。项目采用 Apache-2.0 并
允许任意扩充——**别人做得越好，我能用到的功能也越好**。

## 与同类项目的关系

VB6 的现代化工具并非空白。已知的几条路线：

| 项目 | 路线 | 与本项目的关系 |
|---|---|---|
| Agwstin/vb6-lsp（含 mcp-bridge.mjs） | VB6 语言服务器 + MCP 桥 | **进程外**：编辑期分析强，但无法驱动 IDE 编译/调试/弹窗 |
| xlide / xlide-mcp（VS Code 扩展，已迭代到 v10） | 在 VS Code 里支持 VB6/VBA | 编辑器内辅助，不操作真实 VB6 IDE |
| wieslawsoltes/VB6（VB6 Studio Web） | 浏览器里重建 VB6 IDE | 另一条路线，面向重写而非维护既有工程 |
| 各类 VB6 静态分析 MCP | 源码只读分析 | 只读，不能改、不能编译、不能运行 |

**本项目的不同**：把**真实 VB6 IDE 的完整操作面**（编译、运行、断点、对话框、窗体、打包）
暴露成 MCP，并把**回执诚实性**当作产品的一部分——工具做不到时会给出机器可读的原因，
而不是假装成功。

## 目录结构

```
MCP_VB6/
├─ MCP_VB6.vbp            VB6 工程文件（ActiveX DLL，输出 VB6MCP.dll）
├─ Connect.cls            插件入口，实现 IDTExtensibility2
├─ ServerHost.cls         组装服务器 + Add-Ins 菜单项
├─ McpServer.cls          MCP/JSON-RPC 协议分发 + HTTP 路由
├─ HttpConn.cls           HTTP/1.1 请求缓冲与解析（支持 keep-alive）
├─ ServerForm.frm         隐藏窗体，承载 Winsock 监听 + 端口探测定时器
├─ CodeTools.cls          111 个 MCP 工具的实现（完整 VBIDE 对象模型 + 文件系统 + 语法门禁/编辑安全/编译权威）
├─ JsonLib.bas            零依赖 JSON 解析/序列化
├─ Utf8.bas               UTF-8/ANSI/UTF-16 转换（kernel32 API）
├─ Util.bas               日志、端口文件
├─ register.reg           在 VB6 中注册该插件的注册表项（导入一次即可）
├─ unregister.reg         卸载插件
├─ build.bat              一键编译 + regsvr32 注册 + 导入插件注册表
├─ mcp_config.example.json    Claude Desktop 配置示例
├─ bridge/mcp_stdio_bridge.py stdio 桥接器（Python 3 标准库）
└─ tests/smoke_test.py     直接打 HTTP 端点的冒烟测试
```

## 一、构建

**前提**：装有 VB6 IDE（SP6）。`MSWINSCK.OCX`（Winsock 控件）随 VB6 安装并已注册。

1. 双击运行 `build.bat`（若 VB6 装在其它目录，编辑其中的 `VB6` 变量）。
   脚本会先注册 Add-In Designer 类型库（`MSADDNDR.DLL`），再编译。
2. 成功后生成 `VB6MCP.dll`，并自动执行 `regsvr32` 与 `register.reg` 导入。
3. 启动 VB6 IDE → 菜单 **外接程序(Add-Ins) → 外接程序管理器**，勾选
   **"VB6 MCP Server"**（注册表 LoadBehavior=3，通常会自动加载）。

也可以在 IDE 里直接打开 `MCP_VB6.vbp` 编译（需引用：VB6 Extensibility、
Microsoft Add-In Designer；若编译报 `IDTExtensibility2` 相关错误，见下文 FAQ）。

> ⚠️ 不要把 VB6MCP 工程自身作为当前活动工程去"生成"——插件 DLL 正被 IDE 加载占用，
> 会因文件锁导致编译失败。请关掉 IDE 用 build.bat 编译，或用另一个工程测试。

## 二、端口与多实例

- 默认监听 `http://127.0.0.1:8766`（MSWINSCK 的 LocalIP 只读，监听器绑定全部本机接口，
  但接受连接时校验对端 IP 必须为 127.0.0.1，非回环连接立即拒绝）。
- 端口占用时自动向上探测（最多 +40）。
- 实际端口写入 `%TEMP%\vb6mcp_port.txt`，桥接器自动读取。
- 自定义端口：环境变量 `VB6MCP_PORT`，或注册表 `HKCU\Software\VB6MCP\Port` (DWORD)。
- 同时开多个 VB6 实例时，每个实例各占一个端口；端口文件记录最后启动的实例。
- 运行状态查询：IDE 菜单 **Add-Ins → VB6 MCP Status...**；
  或浏览器访问 `http://127.0.0.1:8766/health`。
- 日志：`%TEMP%\vb6mcp.log`。

## 三、接入 MCP 客户端

### Claude Desktop

把 `mcp_config.example.json` 里的 `mcpServers.vb6` 合并进
`%APPDATA%\Claude\claude_desktop_config.json`（`command` 用你机器上的
python/py 路径），重启 Claude Desktop 即可。

### Cursor（SSE / streamable HTTP 模式）

Cursor 可直接连 HTTP 端点，无需桥接器：

```
类型: sse     URL: http://127.0.0.1:8766/sse
```

### 其它 stdio 客户端

```json
{
  "mcpServers": {
    "vb6": {
      "command": "python",
      "args": ["D:/path/to/VB6MCP/bridge/mcp_stdio_bridge.py"]
    }
  }
}
```

验证链路是否打通：

```
python tests/smoke_test.py --port 8766
```

## 四、工具列表（111 个）

> 权威清单见 **`docs/TOOL_CATALOG.md`**（由运行中的 add-in 经 `tools/list` 导出），
> 离线基线为 `tools/schema_baseline.json`，可用 `tools/check_live_tools.py` 做在线/基线比对。

下表为最初 45 个工具（完整 VBIDE 对象模型部分）；其余 50 个为后续新增的
语法门禁 / 编辑安全 / 编译权威 / 快照一致性 / 测试宿主 / 运行时观测 / 调试控制 /
VBA6 探针 / 治理项（`validate_code` · `check_component` · `read_source` · `write_source` ·
`edit_by_anchor` · `apply_edits` · `preview_edit` · `build` · `verify` · `inject_trace` ·
`remove_trace` · `fault_inject` · `clear_faults` · `set_breakpoint` · `clear_breakpoints` ·
`debug_commands` · `execution_state` · `native_breakpoint` · `project_changes` · `vba6_probe_plan` · `audit_tail` · `export_public_api` ·
`check_public_api` · `dependency_graph` · `purity_check` · `generate_stub` · `record_fixture` ·
`compare_fixture`），外加 **经验库**（`list_lessons` · `add_lesson` · `delete_lesson` ·
`export_lessons` · `import_lessons_md`）与 **扩展库**（`ext_list` · `ext_enable` · `ext_disable` ·
`ext_reload` · `ext_install` · `ext_uninstall`）。

### 经验库与扩展库

- **经验库** `mcp_lessons/`：`defaults.json`（随插件发布，勿改）+ `user.json`（用户层，按 id 覆盖/禁用/新增）。
  合并结果经 `initialize.instructions` 发布给客户端，agent 调工具前即可读到；`add/delete/import` 写 `user.json` 时带 `.bak` 备份。
- **扩展库** `mcp_extensions/`：每个 `*.json` 是一个**声明式工具包**（`name` / `tools[].inputSchema` / `tools[].handler`）。
  启用的包其工具自动并入 `tools/list` 并由 `CodeTools.CallTool` 兜底分发；`state.json` 记录禁用列表（默认启用）。
  处理器只开放少量安全原语（`http_get` / `read_file` / `write_file`），**不从磁盘加载 VB6 代码**。

完整对象模型清单与差距分析见 `内部 VBIDE 对象模型笔记`（由 VB6EXT.OLB 5.3
类型库二进制解析生成，原始数据见 `内部类型库转储`）。

| 工具 | 说明 | 关键参数 |
|------|------|----------|
| `get_project_info` | 工程/工作区全貌：类型、IDE 状态、启动对象、编译路径、组件、引用、已打开工程 | - |
| `set_project_options` | 工程选项：描述、帮助文件/上下文、启动对象、启动工程、编译输出路径、活动工程、只读模式 | `description`, `startup_object`, `build_file_name` ... |
| `open_project` / `add_project` / `remove_project` | 打开 .vbp / 新建工程（或从模板）/ 关闭工程 | `path`, `type`, `template_path`, `name` |
| `get_component_list` | 组件列表：过程、行号、脏状态、设计器信息 | - |
| `get_component_code` | 从 IDE 内存读取组件全部代码 | `name`, `include_line_numbers` |
| `replace_code` | 行级修改：replace/insert/append/delete/**set(整模块替换)** | `name`, `operation`, `start_line`, `end_line`, `code` |
| `add_component` / `remove_component` / `import_component` / `export_component` / `rename_component` / `reload_component` | 组件全生命周期（reload 配合 write_file 使用） | `type`, `name`, `path`, `new_name` ... |
| `add_code` / `create_event_proc` / `find_in_component` | 追加代码（字符串/文件）/ 生成事件过程桩 / 模块内查找 | `component`, `code`, `file`, `object`, `event`, `target` |
| `get_codepane` / `set_codepane` | 代码窗格：光标/选区/滚动/视图模式 | `component`, `selection`, `top_line`, `view` |
| `get_selection` / `activate_component` | IDE 当前选区 / 激活组件（代码或设计器窗口） | `name`, `window` |
| `list_files` / `read_file` / `write_file` / `search_in_files` | 磁盘文件操作（UTF-8/UTF-16/ANSI 自动识别） | `path`, `pattern`, `query` ... |
| `save_project` | 保存工程与全部脏组件 | - |
| `make_project` | **静默编译**（VBProject.MakeCompiledFile，可指定输出路径；失败回退菜单方式）；编译成功后默认给产物打 LAA；`diagnose:true` 用隔离副本拿真编译日志 | `output_path`, `laa`, `diagnose` |
| `control_execution` | 运行控制：run / full_compile / break / end / restart | `mode` |
| `add_reference` / `remove_reference` | 引用：按 GUID 或文件路径添加 / 按名称移除 | `mode`, `guid`, `path`, `name` |
| `get_component_properties` / `set_component_property` | 设计期属性（Properties 窗口）：读取/设置（含多维索引） | `name`, `property`, `value`, `indices` |
| `list_members` / `set_member_attribute` | 类成员：类型/作用域/属性；设置 Hidden/Browsable/Category 等 | `component`, `member`, `attribute`, `value` |
| `list_windows` / `control_window` | IDE 窗口：枚举/聚焦/关闭/显隐/最小化/移动/改名 | `window`, `action` ... |
| `form_designer` / `add_control` / `remove_control` / `set_control_property` | 窗体设计器：控件枚举/选区/剪贴板/增删控件/设控件属性 | `form`, `action`, `prog_id`, `index` |
| `property_bag` | 工程/组件属性包读写（ReadProperty/WriteProperty） | `target`, `key`, `value` |
| `list_addins` / `set_addin_state` | 插件枚举 / 加载卸载插件 | `prog_id`, `connect` |
| `command_bars` | IDE 菜单：枚举或按名称执行命令 | `action`, `bar`, `control` |
| `quit_ide` | 退出 IDE（响应发出后延迟执行，未保存会丢失） | - |

## 五、两条修改代码的路径（重要）

1. **`replace_code`（推荐，作用于 IDE 内存）**：修改立即生效于 IDE 中的代码，
   组件被标记为未保存；之后用 `save_project` 落盘。
2. **`write_file`（写磁盘）**：适合生成新文件或非 VB 文件。若目标是工程内已被
   IDE 打开的文件，用 **`reload_component`** 从磁盘重载即可生效（无需重启 IDE）。

## 六、编码说明

- VB6 源码文件（`.bas/.frm/.cls/.vbp/...`）按 **ANSI（系统代码页，如 GBK）** 读写；
  其它文件默认 UTF-8。`write_file` 可用 `encoding` 参数显式指定
  `auto / utf-8 / ansi / utf-16`。
- MCP 线上协议一律 UTF-8；桥接器 `ensure_ascii=False`，中文内容可正常往返。

## 七、已知限制

- VB6 是单线程 STA：大数据量搜索时会短暂阻塞 IDE 界面（已设 2000 文件上限）。
- 断点状态（break mode）下 `make_project` / `control_execution` 可能失败，
  工具会返回错误文本。
- `make_project` 对当前活动工程执行；无法静默指定输出路径（跟随 IDE 的工程设置）。
- 插件自身 DLL 被 IDE 占用时无法原地重建（见 FAQ）。
- 服务器接受连接时强制校验对端为 127.0.0.1（MSWINSCK.OCX 的 LocalIP 只读，
  无法在监听前绑定回环地址，故用接受时过滤实现），无鉴权——只信任本机进程。

- **本地化**：本项目在**中文版 VB6（zh-CN）**上开发与验证。IDE 标题的模式标记、「立即」
  窗口名、部分对话框按钮标题按中文匹配；英文版 VB6 显示 [Design]/Immediate 等，**模式识别
  与部分弹窗流程可能失效**。宿主 bridge/ 侧是本地化无关的（只假设 VB98 目录名，各语言版本
  一致）。详见本文件「七、已知限制」。

## 八、FAQ / 故障排查

**编译报 "User-defined type not defined: IDTExtensibility2"**
→ 需要 **Microsoft Add-In Designer** 类型库：`MSADDNDR.DLL`（Typelib GUID
`{AC0714F2-3D04-11D1-AE7D-00A0C90F26F4}`，通常位于
`C:\Program Files (x86)\Common Files\designer\MSADDNDR.DLL`）。
`build.bat` 会自动注册它（内部用 `%SystemRoot%\SysWOW64\regsvr32.exe`）；手动注册：
`"%SystemRoot%\SysWOW64\regsvr32.exe" "%CommonProgramFiles(x86)%\designer\MSADDNDR.DLL"`。

**编译报 "Method or data member not found"（如 Member/Reload/MakeCompiledFile）**
→ 工程引用的是 Office 的 VBA 扩展性类型库（{0002E157}，64 个类型，缺 VB6 专属对象）。
必须在 Project → References 中勾选 **Microsoft Visual Basic 6.0 Extensibility**
（`VB98\VB6EXT.OLB`，GUID `{EF404E00-EDA6-101A-8DAF-00DD010F7EBB}`，112 个类型）。
本工程 vbp 已写死该引用；若 IDE 又改回 Office 版，请手动改回。

**本机 VB6 适配说明（已内化到代码）**
- `MSWINSCK.OCX` 的 `LocalIP` 只读：监听绑定全部本机接口，接受连接时校验对端为 127.0.0.1；
  若 OCX 未注册，`build.bat` 前先 `"%SystemRoot%\SysWOW64\regsvr32.exe" "%SystemRoot%\SysWOW64\MSWINSCK.OCX"`
  ★**32 位控件的注册必须用「本机」`%SystemRoot%\SysWOW64\regsvr32.exe`，绝不硬编码 `C:\Windows`**：
  本机 `SystemRoot=B:\Windows`，而 `C:\Windows` 是一份**陈旧的第二安装** ⇒ 用它会让**任何** 32 位 DLL
  注册失败（`0xC0000139`）且 `/s` **静默**（纪律 29；`build.bat` 已用 `%REGSVR%` 显式指向本机 SysWOW64）。
  （或在工程目录放一份 OCX 文件）。
- VB6EXT 中 `CodePane.CodePaneView`、`Window.Caption` 为只读，相应设置请求会被忽略并说明。
- `VBProjects.StartProject` 只有 Put 访问器，需用无 Set 的对象赋值。
- VB6 保留字坑：过程不能叫 `Stop`，参数不能叫 `attribute`；数组参数必须 ByRef。

**编译报 "procedure declaration does not match"（IDTExtensibility2_OnConnection）**
→ `Connect.cls` 已按类型库原生声明使用 `ext_ConnectMode` / `ext_DisconnectMode`
（来自 Add-In Designer）。若个别环境的类型库版本不同，可改为
`vbext_ConnectMode` / `vbext_DisconnectMode`（来自 VB6 Extensibility）。

**加载时报 MSWINSCK.OCX 缺失**
→ 从 VB6 安装源找到 `MSWINSCK.OCX` 放到 `%SystemRoot%\SysWOW64` 并
`"%SystemRoot%\SysWOW64\regsvr32.exe" MSWINSCK.OCX`（★不要硬编码 `C:\Windows`，理由见上文纪律 29）。

**桥接器报 "VB6 MCP server unreachable"**
→ 确认 VB6 IDE 已打开且插件已加载（Add-Ins → VB6 MCP Status... 查看端口），
端口与 `%TEMP%\vb6mcp_port.txt` 一致；防火墙若拦截本机回环，放行 python。

**写文件后 IDE 里没变化**
→ 属于预期行为，见第五节；执行 File → Reload Project。

**想在未打开 VB6 IDE 时也能用**
→ 不行：本插件依赖 VBIDE 对象模型，必须随 IDE 运行。日常做法是让 VB6 常驻后台。

## 九、协议说明

- 实现 MCP 核心：`initialize` / `ping` / `tools/list` / `tools/call` / 通知。
- 协议版本**协商**：服务端支持 2025-11-25 等版本，按客户端 `initialize` 回显其请求的版本
  （未带则用最新；见 `McpServer.NegotiateProtocol`）。
- 传输：`POST /rpc`、`POST /mcp`（streamable HTTP）、`POST /message`；请求头带
  `Accept: text/event-stream` 时用 SSE 帧返回单条消息。**不提供** `GET /sse` 长连订阅，
  因此不使用服务端主动通知（notifications）——客户端按需 `tools/list` 拉取即可。
- 实现了 `resources/list|read`（`vb6://project` + 每个组件的 `vb6://component/<name>`）
  与 `prompts/list|get`（`edit-checklist` / `release-checklist`）；未知 uri/prompt 回 `-32602`。
- JSON-RPC 通知返回 HTTP 202 空响应；工具错误按 MCP 规范以
  `isError: true` 的 content 返回，而非 JSON-RPC error。

**`open_project`（D5 已破案并修复，2026-09-21）**

- `VBProjects.AddFromFile` 的签名是 **`AddFromFile(PathName, Exclusive)`**（见 `内部 VBIDE 对象模型笔记` §3）。
  早先“VB6 该 API 只接受文件名”的说法是错的；实测**两种调用形式都会抛 `类型不匹配(13)`，但工程确实被加进去了**。
- 因此现在**以 IDE 的工程列表判定成功**，返回值只作参考：成功时响应里会带
  `apiError: "0xD 类型不匹配"` 与一句“这是虚警、打开其实成功”的说明。
- `exclusive: true` 需要当前工程可关闭。有未保存改动时 VB6 会报 60082（工程已加载），
  所以现在会**先落盘（响应带 `savedBeforeClose:true`）再关闭**，实测 `closedProjects:1` 后切换成功。
- 逐步日志：`%TEMP%\vb6mcp_openproject.log`；方法级 begin 标记：`%TEMP%\vb6mcp_audit.log`。
- ⚠ 早先记录的“调用后 MCP 端点失联”与 `open_project` **无关**：那是 **VB6 进程崩溃**，
  触发态是 `.vbp` 引用了磁盘上不存在的组件文件。详见 `内部已知限制与 FAQ` §10。

---

## `inject_trace` 会自动附加 `modMcpTrace`，我需要手工导入吗？

**不需要。** 从 `TASK-P1-129` 起，目标工程里没有 `modMcpTrace` 时，`inject_trace` 会**自己把该模块生成并添加到你的工程**，添加后还会复验它确实在工程里；模块已存在时直接跳过（幂等）。

- **自动附加**：模块源码内嵌在插件 VB6MCP.dll 里，不再作为仓库文件发布。回执会带 `moduleAutoAttached` / `moduleAttachAction` / `moduleAttachMs`；附加失败时带 `moduleAttachError`（失败一律如实报告，不静默降级）。
- **注释语言怎么选（界面优先）**：先看 **IDE 自己的界面语言**（主窗口菜单文案 文件/编辑/视图/帮助，其次标题里的模式标记 `[设计]`/`[运行]`）——能可靠读出就用它；读不出来再回退到 VB6.EXE 的版本资源语言 ID（`\VarFileInfo\Translation`，主语言中文 LANG_CHINESE ⇒ 中文）；两者都拿不到 ⇒ **英文**（英文在任何代码页都可读，中文在非中文系统会乱码）。回执里并排给出三个信号：`ideLanguage`（版本资源）· `ideMenuLanguage`（界面）· `commentLanguage`（最终选择），以及各自的 `*Source` 依据。
- **发布提醒**：注入块调用的 `McpTrace` / `McpTraceOn` 就定义在 `modMcpTrace` 里，所以**保留注入调用并编译/发布 EXE 或 DLL 时，必须把该模块一起编进去**（回执里的 `shipNote` 同样说明）。
- **怎么清理**：`remove_trace` 会移除注入块，**并且**把自动附加的 `modMcpTrace` 从工程中一并移除（回执 `moduleRemoved` / `moduleRemoveReason`）。只有带本插件自动生成标记（注释行 `' VB6MCP-AUTOGEN modMcpTrace`）的模块才会被删；同名的、你自己写的模块一律不动。注意 `VBComponents.Remove` 只从工程里移除组件，**不会删除磁盘上的 .bas 文件**（回执用 `moduleFileLeftOnDisk` 如实标注）。
- **日志落点**：追踪由 VB6 进程的环境变量控制 —— `MCP_TRACE_ON=1` 或 `MCP_TRACE_FILE=<路径>` 打开；默认写 `%TEMP%\mcp_trace.jsonl`。写失败记到 `MCP_TRACE_ERR_FILE`（默认 `%TEMP%\mcp_trace_err.log`），不会静默吞掉。
- **其它工具都不需要它**：编辑、检查、编译、运行、断点、弹窗、窗体、快照、打包，全部与它无关。

## 编译产物默认带上 LAA（大地址感知）——`make_project` / `build`

`make_project` 与 `build` 在**编译成功后**由插件自己改写产物的 PE 头，置 `IMAGE_FILE_LARGE_ADDRESS_AWARE`（`0x0020`）——即 EditBin `/LARGEADDRESSAWARE` 的等价物，**不依赖外部 `EditBin.exe`**（用户机器不一定有它）。开关参数是 `laa`（两个工具都接受），**默认开**。

- **谁做决定（优先级从高到低）**：
  1. **本次调用的 `laa` 参数**（`laa:true|false`）——传了就由它决定（`laaSource:"param"`）；
  2. **工程 `.vbp` 的 `[VB6MCP]` 命名节**（`LAA=1`，可选 `LAANote=<说明>`；键名大小写不敏感，`1/true/yes/on` 与 `0/false/no/off` 都认）——`laaSource:"vb-project"`；
  3. **工程目录旁的 `.vb6mcp.json`**（`{"laa": true|false, "laaNote": "..."}`）——`laaSource:"project-file"`；
  4. **环境变量 `VB6MCP_LAA`**——`laaSource:"env"`；
  5. 都没有 ⇒ **全局默认（开）**——`laaSource:"global-default"`。
- **回执**：`laaSource` / `laaProjectFile`（旁挂文件路径）/ `laaBefore` / `laaAfter` / `laaSet`，另附 `laaSourceNote` 与 `laa{}` 明细；产物已带该标志时**幂等**（`laaSet:false`、`laaAlready:true`，一个字节都不写）。
- **配置写错不会被静默吞掉**：`[VB6MCP]` 里 `LAA=` 的值非法，或旁挂文件不是合法 JSON 对象 / 没有 `laa` 键 / `laa` 不是布尔 ⇒ LAA 步骤被**跳过**并回 `laaConfigError`（此时 `laaSource` 为 `vb-project-error` / `project-file-error`），**不会**悄悄退回默认值。
- **两条护栏（都拒绝写盘并给出原因）**：
  - `signed_image`：产物带 Authenticode 证书表（`IMAGE_DIRECTORY_ENTRY_SECURITY` 非空）⇒ **拒绝改写**，因为改 PE 头会让签名失效。正确顺序是**先打标志、后签名**（签名必须是最后一步）；确实要非 LAA 产物就显式传 `laa:false`。
  - `own_addin_dll`：目标是插件自己的 `VB6MCP.dll` 时**故意不打**（ActiveX DLL 的 LAA 由宿主进程的地址空间决定，打在 DLL 上本身没有意义）。
- 写失败或复读校验不过 ⇒ 把原特性字回写（回滚），产物保持不变；只有在复读确认新值之后才报 `laaSet:true`（`laaBytesPatched:2`）。
- `laaMachineEffective`：该标志只在 64 位 Windows（或 32 位 Windows 开 `/3GB`）上才真的扩大 32 位进程的用户地址空间；否则回执如实标注它在本机不生效。
- **运维要点（重要）**：`[VB6MCP]` 节必须在**工程在 IDE 里关闭时**写进 `.vbp`。工程**打开期间**从外部追加的节不在 IDE 的内存工程模型里，会被 VB6 那一次**保存工程**覆盖掉；而**打开时就已经在**的节，此后每次保存都会原样保留。需要随时可写，就改用旁挂 `.vb6mcp.json`（VB6 只重写 `.vbp`，不碰它）。

## 编译说成功却没产出？先看 `unbuiltReasonUnknown`，再用 `diagnose:true` 拿真编译日志

- **IDE 路由（默认 `route:"ide"`）没有编译器子进程**：`VBProject.MakeCompiledFile` 只回报成功/失败，所以这条路由**结构上**就写不出 `/out` 日志 ⇒ 回执带 `logAvailable:false` 与 `logUnavailableReason` / `diagnosisUnavailableReason`；`exitCode:0` 只表示「没有编译器子进程跑过」，**不是**「编译器没有发现错误」。
- **`compiled:true` 不是证据**：回执永远给 `compiled_means`、`artifactChanged`、`verdict` 与 `output{exists,size,mtime,sha256}`。当 `compiled:true` 而 `artifactChanged:false` 时，回执会**自动**置 `unbuiltReasonUnknown:true`，并写明**未产出、原因未知、可能编译失败**——绝不粉饰成成功。
- **`diagnose:true`（也会自动运行）**：把 `.vbp` 及其组件文件复制到临时目录，对**副本**跑隔离的 `VB6.EXE /make /out <log>`，回**真日志 + `firstError{file,line,text}`**，来源标为 `firstErrorSource:"isolated_copy_make"`；明细在 `diagnosis{}`（`route:"isolated_copy_make"`、`tmpDir`、`copiedFiles`、`missingFiles`、`errorCount`、`artifactExists` 等）。在「`compiled=true` 但 `artifactChanged=false`」时它会**自动运行**，因为那条判决否则与编译失败不可区分。
  - 副本反映的是**磁盘**上的状态，所以**未保存的窗格改动不在副本里**。
  - 已知边界：另一个 VB6 实例正持有同一个 `.vbp` 时，隔离 `/make` 会**退出 0 且不写日志**——先关掉那个实例，或改对工程副本编译，再重试。
- 编译失败时还会尝试读 VB6 自己的编译错误框（`firstErrorSource:"ide_compile_error_dialog"`；`errorCount:1` 只表示「看到过一个框」，不是「恰好一个错误」）；读不到就如实说读不到。
- `project_meta{action:"set", cond_comp:""}`——**清空宏是被接受的**（生产构建要清掉测试宏）：**参数存在**就触发写入，空串不等于「没给字段」。

## 断点（break）模式下，代码窗滚轮也能用了

- 插件自带的代码窗滚轮（**默认关闭**，可在 MCP 设置界面里开启）在**中断态**下也按窗格工作：消息落在哪个代码窗格就滚哪个。
- 以前在断点态失效的根因：`VBE.CodePanes` 在断点态**可能返回空表**（调试器打开的那个模块窗格在设计期从未存在过），于是每个 `VbaWindow` 都被判成 `not_a_code_pane`、从不挂钩 ⇒ 滚轮消息根本进不来。现在只要曾经确认识别过一个代码窗格，**空表时仍会挂钩**；这次表是从哪来的由 `captionsSource` 如实给出（`all` / `active_fallback` / `empty`）。
- **失败原因如实分流**（不再共用一句含糊的失败）：
  - `lastTopStatus:"pane_not_found"`：标题没能解析成 CodePane ⇒ 滚轮原样转发，什么都不滚；
  - `lastTopStatus:"pane_not_writable_in_break"`：窗格**解析到了**，但 IDE 拒绝了 `TopLine` 的读/写（断点态实测）⇒ 回执带**真实 VB 错误**（`lastTopError`，形如「VB error N (…)」），并把这次滚动**排队**，由**定时器在窗口过程（WndProc）之外重试**（`pendingScroll` / `pendingScrollResult`）。
- **观测字段**：`lastResolveVia`（`caption` / `component` / `active_name` / `none`，这次靠哪条通道找到窗格）、`lastTopStatus`、`lastTopError`、`captionsSource` / `captionsCount`，以及原有的 `lastTopLineBefore/After`、`handled` / `forwarded`、`msgSeen` 等（都在 `ide_state` 与 `mcp_info` 的 `wheel{}` 里）。
- **立即 / 监视 / 局部变量窗不会被误映射**：它们与代码窗格共用窗口类，但唯一那条「映射到活动窗格」的兜底要求窗口标题里**确实含有**该窗格的组件名，因此这三个非代码窗口结构上不可能被当作代码窗格去滚（只会出现在 `skipped[]` 里）。
- `lastTopLineBefore/After` 是 VBIDE `CodePane.TopLine`（1 基）；**`-1` 表示「没有值」，不是第 0 行、也不是「没有滚动」**——原因看 `lastTopStatus` / `lastTopError`。要看当前行请读 `paneTopLineNow`（活动窗格的 `TopLine`）。

## 许可证

**Apache License 2.0**（见 LICENSE）。可自由使用、修改、分发、商用，也允许闭源衍生；
只需保留版权与许可声明，并说明你修改了什么。

**平台前提**：本项目是 VB6 的插件，**使用与构建都需要你自己合法的 VB6 安装**；本仓库
不分发任何微软文件（见 THIRD_PARTY_NOTICES.md）。本项目与 Microsoft 无隶属关系。

## 关于代码注释里引用的文档

源码与注释中会提到 内部逆向记录（多篇）、内部交接说明、内部方案文档、
内部已知限制与 FAQ、内部 VBIDE 对象模型笔记、内部类型库转储、
内部命令栏清单、内部禁用路由说明 等文件。它们属于维护者的内部设计与
逆向研究记录，不随源码发布（本仓库只发布代码，以及构建与使用所需的最小集合）。这些引用不影响
编译与运行；如果你确实需要其中某一篇的内容，请开一个 issue 说明用途。

docs/tools_schema.json（工具面契约）与 tools/schema_baseline.json（冻结基线）已包含在本源码包中，
因此 tools/ 下的门禁脚本可以直接运行。

## 如何参与

见 CONTRIBUTING.zh-CN.md（构建、注册、门禁、编码规则）。

**第一优先是本地化**：本项目在**中文版 VB6** 上开发与验证，已知耦合点可枚举（见「七、
已知限制」）。如果你在用英文版或其它语言版本，哪怕只报告「能用/不能用」也是极有价值
的贡献。
