# VB6MCP — 让 VB6 IDE 成为 MCP 服务器的插件

> **中文** | [English](README.en.md)

> **变更记录**：[CHANGELOG.md](CHANGELOG.md)

VB6MCP 是一个 **VB6 插件（Add-In）**：把它加载到 VB6 IDE 之后，IDE 就变成一个
[MCP (Model Context Protocol)](https://modelcontextprotocol.io) 服务器。
Claude Desktop / Cursor / Codex 等 AI 编程助手即可直接读取、搜索、修改当前打开的
VB6 工程，并控制编译与运行。

```
 MCP 客户端 (Claude Desktop / Cursor / ...)
        │  stdio (JSON-RPC, 每行一条)
        ▼
 bridge/vb6mcp_hub.mjs      (Node，零依赖；MCP 线协议 + ide_* / 经验库共 10 个工具)
        │  HTTP  http://127.0.0.1:8767  (配了凭据时需 bearer token)
        ▼
 bridge/vb6mcp_bridge.py    (Python 3，仅标准库；本机唯一的 TCP 监听者)
        │  命名管道 (每个 IDE 实例一条)
        ▼
 VB6MCP.dll  ← 运行在 VB6.EXE 进程内 (自身不再监听 TCP)
   ├─ 命名管道服务端
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
- **已知局限**：见「七、已知限制」，尤其是**端到端验证只在中文版 VB6（zh-CN）上做过；英文版能力上支持但未闭环**这一点。
## 为什么做这个（项目背景）

VB6 在 2008 年就停止演进了，但它写的系统还在跑：工厂、财务、医院、政府内部系统里有大量
陈旧 VB6 工程仍需要维护，而维护者面对的是一套**没有任何现代工具链**的环境——没有语言
服务器、没有可靠的静态检查、编译报错只能手动点、改代码靠自己不手抖。

AI 编程助手已经能很好地处理现代语言，却**够不到 VB6**：它们既看不见 IDE 里打开的工程，
也点不了那个「生成工程」对话框。本项目就是补这个缺口。

**做法**：不另造编辑器，而是**把 VB6 IDE 本身变成一个 MCP 服务器**——加载一个插件，IDE
就对外提供 120 个工具：读/搜/改代码、静态检查、锚点编辑与安全事务、编译、运行与单步、
弹窗处理、窗体与项目元数据、快照与回滚、打包。此外还有 10 个 **hub 面**工具（IDE 实例管理与
经验库）在 IDE 进程外提供，客户端合计可见 130 个。
**因为工具就住在 IDE 进程里**，它能做进程外方案做不到的事：真的去点那个对话框、真的按运行、
真的读断点处的局部变量。

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
├─ CodeTools.cls          120 个 MCP 工具的实现（完整 VBIDE 对象模型 + 文件系统 + 语法门禁/编辑安全/编译权威）
├─ JsonLib.bas            零依赖 JSON 解析/序列化
├─ Utf8.bas               UTF-8/ANSI/UTF-16 转换（kernel32 API）
├─ Util.bas               日志、端口文件
├─ register.reg           在 VB6 中注册该插件的注册表项（导入一次即可）
├─ unregister.reg         卸载插件
├─ build.bat              一键编译 + regsvr32 注册 + 导入插件注册表
├─ mcp_config.example.json    Claude Desktop 配置示例
├─ bridge/vb6mcp_hub.mjs   MCP 入口（Node，stdio；ide_* 与经验库工具）
├─ bridge/vb6mcp_bridge.py 中间服务：本机唯一的 TCP 监听者（默认 8767），转发到命名管道
└─ tests/smoke_test.py     直接打 HTTP 端点的冒烟测试
```

## 用 AI 助手一键安装（推荐）

不想照下面一步步点？把这一段**整段**复制进你的 AI 编程助手（Claude Code / Cursor / Codex / 任何能开终端的 agent）。
唯一要你补的是仓库地址：把 `<REPO_URL>` 换成这个仓库的 URL（先填上更好；留空也行，它会回头问你）。
它会自己取源码、编译、注册、启动 IDE、**真的验证 MCP 服务能用**，再把客户端接好，最后把「做了什么、看到什么、哪一步没做到」原样报给你。

```text
你是我的安装助手。请在一台 Windows 机器上，从零把 VB6MCP 装好、启动，并通过真实验证。VB6MCP 是一个 VB6 IDE 插件（Add-In，ActiveX DLL，编译产物 VB6MCP.dll）：把它加载进 VB6 IDE 之后，IDE 本身就成了一个 MCP 服务器，AI 客户端可以直接读写我当前打开的 VB6 工程、编译并运行它。

源码仓库：<REPO_URL>
如果这个地址还是占位符（空的、或仍写着 <REPO_URL>），先问我要仓库地址；不要猜，也不要去找一个同名的别的仓库。

本次交付：源码取到 -> 编译 -> 注册 -> 启动 IDE -> 真实验证 MCP 服务可用 -> 接好客户端 -> 跑冒烟测试 -> 如实汇报。

边界（请先读完再动手）：
- 只做这一件事。不要顺手改我别的工程、别的插件或别的注册表项。
- 禁止破坏性或不可逆的操作：不要删除/覆盖非你创建的文件和目录，不要改动我已有的 VB6 工程，不要卸载或禁用我其它插件，不要"清理"进程或注册表。
- 不要杀进程。不要结束 VB6.EXE 或任何进程；需要关闭 IDE 时先问我（我可能有没保存的改动）。
- 端口只碰你探测到的那个。这一代里**中间服务**是唯一的 TCP 监听者：默认 127.0.0.1:8767（用 --port 或 VB6MCP_BRIDGE_PORT 改），插件自身不再监听 TCP。真实端口以 http://127.0.0.1:<端口>/health 返回的 port 字段为准；python bridge/vb6mcp_bridge.py --print-config 也会打印生效端口、传输与 authEnabled。不要扫描、占用或改动其它端口，也不要把这个服务暴露到回环地址之外（非回环绑定会被直接拒绝并以退出码 6 结束，没有覆盖开关）。
- 任何可能弄坏我现有 VB6 环境的动作（机器级重新注册 OCX/类型库、改 HKCU\Software\VB6MCP、改 VB6 的引用设置等），先停下来问我：你要改什么、为什么、怎么撤回。
- 所有产物只放在你新建的工作目录里；每个结论都要有可复现的命令和原始输出，不要用"应该成功了"下结论。需要我手工点的地方，直接告诉我。

步骤（按顺序执行，每步给出证据）：

1) 查前提。确认是 Windows（32 位或 64 位），然后检查：
   - 有没有装 VB6 IDE（SP6）：找 VB6.EXE（标准安装目录是 VB98 文件夹，各机器可能不同）。
   - MSWINSCK.OCX（Winsock 控件）是否可用/已注册；MSADDNDR.DLL（Add-In Designer 类型库）是否可用。
   - 32 位注册工具：64 位 Windows 上用 %SystemRoot%\SysWOW64\regsvr32.exe，不要用裸 regsvr32，也不要把 Windows 目录写死（用 %SystemRoot% 变量）。
   - 如果没有 VB6：立刻停下并如实告诉我——VB6MCP 是 VB6 IDE 的插件，没有 VB6 就既不能编译也不能运行，没有绕过办法；仓库也不分发任何微软文件，必须用我自己合法的 VB6 安装。不要试图用别的工具链"伪造"VB6，也不要在没有 VB6 的机器上继续后面的步骤。
   - 如果缺 MSWINSCK.OCX / MSADDNDR.DLL，按仓库 README 的 FAQ 处理；要写系统目录先问我。

2) 取源码。git clone <REPO_URL> 到你/我约定的工作目录（或下载仓库 ZIP 解压）。确认目录里有 MCP_VB6.vbp、build.bat、register.reg、bridge/vb6mcp_bridge.py、bridge/vb6mcp_hub.mjs、tests/smoke_test.py。

3) 编译。优先直接运行 build.bat（它会先注册 MSADDNDR.DLL，再编译出 VB6MCP.dll）。如果 VB6 不在默认位置，改 build.bat 里的 VB6 变量，不要另处硬编码路径。
   注意：不要在 IDE 已经加载了这个插件 DLL 时原地重编译——DLL 被占用会让编译失败。先关掉 IDE（关之前先问我有没有没保存的东西），或用另一个 VB6 实例构建。
   也可以在 IDE 里打开 MCP_VB6.vbp 编译，但不要把 VB6MCP 工程自己当成当前活动工程去"生成"（同样是文件锁）。
   编译失败时把编译器原始报错贴出来（例如工程引用是否勾选了 VB6 Extensibility 与 Add-In Designer，见 README 的 FAQ）；不要凭猜测改源码。

4) 注册并启动 IDE。用步骤 1 确认的 32 位 regsvr32 注册 VB6MCP.dll，导入 register.reg，然后自己查一遍注册表确认：
   reg query "HKCU\Software\Microsoft\Visual Basic\6.0\Addins\VB6MCP.Connect" /v LoadBehavior
   应得到 0x3（=3，表示自动加载），FriendlyName 应为 VB6 MCP Server。build.bat 一般会替你做注册和导入，但仍要自己查证。
   然后启动 VB6 IDE（它本来就开着且没加载插件时，先问我要不要重开，不要自己杀进程），在菜单 外接程序(Add-Ins) -> 外接程序管理器 里确认 "VB6 MCP Server" 已勾选（LoadBehavior=3 通常会自动加载）。

5) 真实验证服务器（关键一步，不能只看"进程起来了"）：
   - GET http://127.0.0.1:8767/health，直到返回 {"ok":true,...}；以响应里的 port 为准，并与 python bridge/vb6mcp_bridge.py --print-config 的 port 对照。下面把这个端口记作 <端口>。**若你（或用户）配置了 token，这个请求也要带 `Authorization: Bearer <token>`**，否则会得到 401。
   - 对它发一次真正的 MCP 调用：POST http://127.0.0.1:<端口>/rpc，Content-Type: application/json
     {"jsonrpc":"2.0","id":1,"method":"initialize","params":{"protocolVersion":"2025-11-25","capabilities":{},"clientInfo":{"name":"install-check","version":"1.0"}}}
     然后 {"jsonrpc":"2.0","id":2,"method":"tools/list","params":{}}
   - 数 tools/list 里的工具个数，并与仓库自己的权威清单对比：docs/TOOL_CATALOG.md 里记录的数量，以及 tools/schema_baseline.json 里 tools 数组的长度。不要写死任何数字——就用你刚克隆下来的仓库里的数字。也可以调用 mcp_info 工具，看服务端自报的 contract 与 toolCount。
   - 三者一致就通过并记录数字；不一致就如实报告差多少、多了/少了哪些，不要改仓库、也不要"凑数字"。如果 tools/list 根本拿不到，把原始响应和 HTTP 状态码贴出来。

6) 接客户端（新一代推荐 stdio 启动 hub）：command 指向 node，args 指向 <仓库>/bridge/vb6mcp_hub.mjs，并给它两个环境变量 —— VB6MCP_BRIDGE_URL=http://127.0.0.1:<中间服务端口>（hub 源码默认 http://127.0.0.1:8767，服务端口不同时必须显式给出）与 token（VB6MCP_TOKEN 或 VB6MCP_TOKEN_FILE，仅当服务端配置了凭据时）。Claude Desktop 是 %APPDATA%\Claude\claude_desktop_config.json；Cursor 用它的 MCP 配置（也可以 streamable HTTP 直连 http://127.0.0.1:<端口>/mcp，配了 token 时带 Authorization: Bearer <token> 头）。把占位路径换成这台机器上的真实路径；JSON 里用 / 或写成 \\。改配置前先备份原文件。

7) 跑冒烟测试：python tests/smoke_test.py --port <端口>（<端口> 指中间服务端口）。这一步必须在 VB6 IDE 已打开且插件已加载的前提下跑，它不是离线自检。注意该脚本**不发送 token**，仅在服务端未配置凭据时可用；配了 token 时请用 python bridge/vb6mcp_bridge.py --print-config 查看状态，并改用 --selftest 做进程内自检。

8) 汇报（最重要）。给我：
   - 你实际执行过的命令（原样）与实际观察到的结果（关键输出、health 响应、工具数、冒烟测试结果）；
   - 每一步的判定（通过/失败/跳过），失败的给出原始报错；
   - 你没能做到的事，直说（例如"没装 VB6，所以无法编译"、"需要我手工点某个对话框"、"这一步没权限"）。
   不要为了让报告好看而编造成功；做不到就说做不到。

再次强调：不要杀进程；不要动你没被要求动的端口、服务、注册表、工程文件；任何可能破坏我现有环境的事，先问再做。我需要的是"确实成功了"，或者"诚实地说清卡在哪里"。
```

想手动来一遍 → 见下面「一、构建」「二、连接模型、端口与多实例」「三、接入 MCP 客户端」。两条路任选其一即可。

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

## 二、连接模型、端口与多实例

**新一代：一个监听端口 + 每个 IDE 实例一条命名管道。**

- 客户端不再直连插件。本机由**中间服务** `bridge/vb6mcp_bridge.py` 持有**唯一一个 TCP 端口**：
  默认 `http://127.0.0.1:8767`（用 `--port` 或环境变量 `VB6MCP_BRIDGE_PORT` 改；`VB6MCP_BRIDGE_HOST` 只接受回环地址）。
- **非回环绑定不可能**：绑定到非回环地址会被拒绝，并以**退出码 6** 结束，**没有覆盖开关**。
- MCP 客户端经 **hub** `bridge/vb6mcp_hub.mjs`（Node，零依赖，stdio）接入：hub 提供 `ide_*` 与经验库共 10 个工具，
  并把插件面的 120 个工具转发给中间服务；hub 用环境变量 `VB6MCP_BRIDGE_URL` 找中间服务（源码默认 `http://127.0.0.1:8767`）。
- 插件只接受**命名管道**（每个 IDE 实例一条；同一条管道**不允许并发连接**）⇒ 客户端应在**同一条连接上复用请求 id**，
  或把调用串行化；管道忙时回执给 `pipeBusy`，而不是「没有 IDE」。
- **传输选择** `VB6MCP_TRANSPORT=auto|pipe|tcp`（默认 `auto`：优先管道，不可用时**可见地**回退到插件的 HTTP；
  `pipe` 为严格管道、缺管道是显式错误；`tcp` 只走插件 HTTP，用于回滚与过渡）。
- **自定义端口**：`--port N`、`VB6MCP_BRIDGE_PORT`（或同一台机器上的 `VB6MCP_PORT`）。
  同时开多个 VB6 实例时**仍然只有一个服务端口**；多实例靠 `_target` 寻址（见第三节）。
- 运行状态查询：IDE 菜单 **Add-Ins → VB6 MCP Status...**；或 `GET http://127.0.0.1:<端口>/health`
  （配了凭据时该请求也要带 token）；`python bridge/vb6mcp_bridge.py --print-config` 打印生效端口 / 传输 / `authEnabled`。
- 日志：中间服务的日志路径见 `--print-config` 的 `logPath`；插件自身日志仍是 `%TEMP%\vb6mcp.log`。

### 哪些工具免 IDE、哪些必须有活 IDE

**插件面 120 个工具：必须有活的 IDE。** 它们全部由运行在 `VB6.EXE` 进程内的插件实现（读/改代码、静态检查、编译、断点、窗体、快照……），没有可执行的实例就没有执行者。**列出**工具可以在没有 IDE 时进行（hub 会用随包快照回 `tools/list`，标 `faceSource=snapshot`），但**调用**它们需要本机有活实例，或显式寻址一个有活实例的远端端点；没有活实例时的拒绝码见下一小节。

**hub 面 11 个工具：在 IDE 进程之外运行，调用它们不需要加载插件**，但依赖各不相同：

- **完全不需要 IDE**（只读本机状态，或只读写本地 JSON 存储）：
  - `ide_list` —— 读本机实例登记表 + 端口扫描 + **已配置远端端点的探测状态**；一个实例都没有也照常回执（并说明原因与过期登记项）；
  - `ide_target` —— 只改 hub 自己的默认目标（本地状态，不碰 IDE）；
  - `list_lessons` / `add_lesson` / `delete_lesson` / `export_lessons` / `import_lessons_md` —— 只读写 `mcp_lessons/` 下的 JSON（写操作带 `.bak`）。
- **需要一个目标实例才动手**：`ide_stop`（按 pid 停掉**本 hub 自己启动的**实例，绝不动外来实例）、`ide_restart`（恢复自有实例并验证就绪）、`ide_kill_and_relaunch`（管道 busy/wedged 时的带外复位：停自有实例、清注册、原工程重启并轮询就绪；绝不碰 8766/远端实例）。没有任何实例时它们回结构化拒绝，而不是假装成功（`ide_restart` 在「一个实例都没有」时回 `no_ide_running`）。
- **不需要现有 IDE，但需要能启动 IDE 的条件**：`ide_start`（显式 `project` 或 `VB6MCP_AUTOSTART_PROJECT`，加一个可解析的 `VB6.EXE`；否则回 `not_implemented` / `vb6_exe_not_found` / `project_not_found`）。

> 可离线运行的是 `tools/check_*` 那一组门禁；`tests/smoke_test.py` 与 `mcp_strict_audit` **必须**先打开 IDE 并加载插件 —— 它们不是离线自检。

### 两个「没有 IDE」类拒绝码，各自的前置条件

**`no_ide_running`**：本机**一个活实例都没有** —— 登记表里没有任何「pid 还活着」的条目，也没有实例能被探测到。它**不**用于下面这些情况（各有更精确的形状）：

- 实例**已登记、进程还活着**，只是没有应答探测 ⇒ 回执给出该 pid、判定字段与恢复用的 `nextStep`，而不是 `no_ide_running`；
- **只有远端**端点上有活 IDE、而这次调用没有显式寻址 ⇒ `remote_requires_explicit_endpoint`；
- hub 看不到机群是因为**中间服务不可达**或**被鉴权拒绝** ⇒ `bridge_unreachable` / `bridge_auth_failed`（不是「本机空闲」）；
- 有多个活实例、无法唯一确定 ⇒ `ambiguous_target`（附 `candidates[]` / `suggestion` / `resolution`）。
  另外：没有 `_target` 且本机**只有一个**活实例时，它会自动被选中，不报错。

**`remote_requires_explicit_endpoint`**：前置条件必须**同时**成立 —— (a) **本机没有活实例**；(b) **至少一个已配置的远端端点**上有活 IDE；(c) 这次调用**没有**给带 endpoint 的目标（`_target` 不带 endpoint，也没有通过 `ide_target` 设过默认）。此时调用**不会被自动路由到远端**（本机也没有实例）⇒ 回执是 `ok:false` + `routed:false` + `qualifiedKeys[]`（可直接照用的 `endpoint:projectPath`）+ `hint`。显式给 `_target="<endpoint>:<path>"`（或先 `ide_target{target:"<endpoint>:<path>"}`）即可寻址。**裸 pid / 裸工程名只寻址本机**，永远不会被解释成远端。

### 工程身份：任意 junction / 别名拼写 = 同一个工程

VB6 不允许同一个 `.vbp` 被打开两次，而同一个工程**可以有很多种拼写**：目录 junction（别名目录）、junction 指向的真实目录、`/` 与 `\`、不同大小写、8.3 短名。hub 用规范实路径把它们**全部归一成同一个身份** —— 这些拼写是**同一个候选，不是冲突**（**同一个身份被两个活实例发布**才是冲突）。

归一化不会抹掉「你用的是哪种拼写」：`ide_list` 对每个实例同时给出

- `projectPathReal` —— **规范身份**（realpath 归一后）；
- `projectPathAsGiven` —— **登记表/探测发布出来的**拼写（`projectPathIsAlias` 标出它与规范路径不同）；
- `aliasAsGiven` / `aliasesAsGiven[]` / `aliasSource` / `aliasMatchedBy` —— **调用方传入的**拼写（`_target` 或 `ide_target` 默认值），以及它解析到哪个候选、依据是什么（规范路径相同 / 别名目录包含该工程 / 工程名）；
- 顶层 `aliasEcho` —— 逐条列出这次调用传入的每个拼写、命中几个候选、是否**全部落到同一个候选**（`sameCandidate`），并附上这条身份规则原文。

**这些是加法回显，不是选择器**：解析目标仍按规范身份与寻址规则（`targetKey = [endpoint:]projectPath`）。**同一个规范路径在同一个端点上出现两次才算冲突，出现在两个端点上则不算**。收益是双向的：任何拼写（junction、真实路径、分隔符差异、大小写、8.3 短名）都落到同一个工程；而「它为什么认成了另一个身份」可以直接从 `aliasEcho` 读出依据，不必猜。

### 常驻中间服务：安装、升级、健康与卸载

中间服务（`bridge/vb6mcp_bridge.py`）是**常驻层**，不是按需进程：它**占住一个 TCP 端口**、等待并管理所有 add-in 的命名管道、转发 hub 的调用。**即使本机没有任何 VB6 IDE 在运行，它也必须活着**；IDE 后起时会被**自动接管**（add-in 注册管道后立即可转发，无需重启服务）。hub 侧按需拉起（lazy-ensure）只是**兜底**，常驻服务才是主路径；两者并存不冲突（同一端口只能有一个监听者，服务持有这个单实例）。

- **安装（需管理员）**：`powershell -File tools\install_middle_service.ps1 -Action install`。它写一个启动器（环境变量 + 日志重定向）、注册**开机自启**的计划任务（`AtStartup`、`SYSTEM` 账户、失败按 `RestartCount/RestartInterval` 自动重启），启动后**轮询 `/health`** 直到 200。**先看计划**：加 `-DryRun` 只打印将执行的命令与所需权限、不做任何改动。脚本**从不自行提权**；非管理员运行会返回结构化 `needs_elevation`。
- **默认值**：`-Port 8797`、`-ExePath dist\vb6mcp_bridge.exe`（**缺失时自动回退**到 `python bridge\vb6mcp_bridge.py`）、`-TokenFile` 默认**当前用户目录**下的 `.vb6mcp\token.txt`；另有 `-TaskName`、`-AuditLog`、`-UseService`（检测到 nssm 时改走 Windows 服务）。
- **升级**：先停任务 → 替换产物（`dist\vb6mcp_bridge.exe` 或源码）→ 再跑一次 `-Action install`（幂等覆盖注册）→ `-Action status` 应报 `verdict:UP`。
- **卸载（需管理员）**：`-Action uninstall`：停止并移除任务/服务、停止本服务的进程；**保留 token 文件与日志/审计**（脚本不会删它们）。
- **健康检查（无需管理员）**：`-Action status` 输出任务/服务、进程、端口是否 `LISTENING`、`/health` **原文**（存在 token 文件时自动带 `Authorization: Bearer`），并给出 `verdict`：`up` / `up-health-unverified` / `degraded` / `down`。**语义**：`/health` 的 **HTTP 200 表示中间服务本身在跑**；body 里的 `ok` 表示**当前是否有活 add-in** —— 没有 IDE 时 `ok:false` 属正常（服务在等），**不是**服务故障。`-Action smoke` 用**临时端口**启一份新进程跑 `--smoke`（不需要管理员、不碰线上端口），输出结构化结果。
- **日志与审计**：服务日志由启动器写到 `%ProgramData%\vb6mcp\logs\middle-service.log`；远端模式（`--allow-remote` + TLS + token）要求 `--audit-log`，安装器默认指向 `%ProgramData%\vb6mcp\logs\remote-audit.jsonl`（JSONL，仅记录外网来源的请求）。两者在卸载时**保留**。
- **token 轮换**：token 来源优先 `VB6MCP_TOKEN`（不落盘），否则 `VB6MCP_TOKEN_FILE`。轮换时把旧 token 传给 `--token-prev`（或 `VB6MCP_TOKEN_PREV`）⇒ 新旧在一个窗口内同时有效，客户端切换完成后去掉旧值。安装器只**引用** token 文件，不创建也不删除。
- **端口必须一致**：`VB6MCP_BRIDGE_URL` 必须与安装器的 `-Port` 指向**同一个**服务端口（安装器默认 **8797**；`bridge` 源码自身默认是 `8767`，用安装器时以 `-Port` 为准）。改端口要**同时**改安装器与客户端配置，否则 hub 连不上。
- **故障排查**：`status` 报 `listener.listening:false` ⇒ 任务没起或进程已退（看日志与任务上次结果 `lastTaskResult`）；`health.httpStatus:401` ⇒ 服务在跑但 token 不匹配（核对 `-TokenFile` 与客户端 token）；hub 报 `127.0.0.1:null/rpc`、`RemoteDisconnected` 或所有插件工具静默失败 ⇒ **没有服务在监听该端口**（本轮曾因此让 120 个插件工具全部静默）；`verdict:degraded` ⇒ 任务存在但端口未监听，重启任务并复看日志。

## 三、接入 MCP 客户端

### 连接与鉴权（新一代）

- **推荐**：客户端以 **stdio** 启动 hub —— `command: node`，`args: ["<仓库>/bridge/vb6mcp_hub.mjs"]`，
  环境变量给 `VB6MCP_BRIDGE_URL=http://127.0.0.1:8767`（中间服务端口不同时必须显式给出）。
- **HTTP 直连**（streamable HTTP）：`POST http://127.0.0.1:<端口>/mcp`（`/rpc` 是同一面）。
- **鉴权**：一旦服务端配置了凭据（`VB6MCP_TOKEN`，或 `VB6MCP_TOKEN_FILE` 指向一个仅属主可读的文件），
  **所有 HTTP 面**（`/rpc`、`/mcp`、`/health` …）都需要 bearer token：`Authorization: Bearer <token>`，
  或 `X-VB6MCP-Token: <token>`。**把凭据放进 URL 会被拒绝**（401 `token_in_url_rejected`）；
  **未配置凭据时服务照常运行**，但 `/health` 会明确标注 `auth.enabled=false` —— 多用户机器上不要这样长期运行。
- **寻址**：调用用 `_target` 指定目标工程，规范写法 `[endpoint:]projectPath`（例如 `C:/work/MyProj.vbp`），
  省略 `endpoint` 即本机；**没有 `_target` 的调用不会被送到远端**。
- 旧代的直连端口（8766）在新一代**不再存在** —— 仍指向它的客户端会直接连接失败。

### 远端端点：可达性来自转发层，凭据只发回环端点

- **中间服务只绑回环**（非回环绑定被直接拒绝、退出码 6、无覆盖开关）⇒ **跨机可达性不由它提供**。
  要让另一台机器上的 hub 连到本机的中间服务，需要在它**前面**放一层**中继 / 端口转发**，把「只绑回环的」
  服务暴露成一个 hub 可访问的 base：`http://<host>:<port>`（RPC 路径即 `…/rpc`）。
- **hub 只对回环端点自动发送凭据**：唯一的自动机制要求「端点是回环地址（`127.0.0.1` / `localhost` / `::1`）」
  且「该端点用 `401 token_missing` 明确索要」两者同时成立 —— 因此 token 不会经这条路径离开本机。
- **对白名单里的远端端点，hub 不带 token** ⇒ 远端端点要么**自身不鉴权**，要么把鉴权做在**转发层**
  （中继 / 端口转发那一层）。
- ★例外（要清楚自己在做什么）：`VB6MCP_BRIDGE_URL` 指到哪个 base，hub 就把它当作**自己的中间服务**，
  不论它是否回环 —— 所以只有当你**显式**把一个远端 base 配成 bridge URL 时，hub 才会给它带 token。
- **远端端点来自白名单**：环境变量 `VB6MCP_REMOTES`，或 hub 目录下的 `hub_remotes.json`（仓库附带
  `hub_remotes.example.json` 模板），每条是 `{label, url, note?}`。**只探白名单、绝不自动发现主机**；
  连不上的端点**照样列出**并带**探测状态**；未知 label / 连不上 / 超时分别是结构化拒绝
  `endpoint_unknown` / `endpoint_unreachable` / `endpoint_timeout`。

### Claude Desktop

让 Claude Desktop 以 **stdio** 启动 hub（路径换成你自己的）：

```json
{
  "mcpServers": {
    "vb6": {
      "command": "node",
      "args": ["C:/path/to/VB6MCP/bridge/vb6mcp_hub.mjs"],
      "env": { "VB6MCP_BRIDGE_URL": "http://127.0.0.1:8767" }
    }
  }
}
```

服务端配了 token 时，在 `env` 里再加 `VB6MCP_TOKEN`（或 `VB6MCP_TOKEN_FILE`）。

> ★**配置里的路径都要换成你自己机器上的**：示例（包括 `mcp_config.example.json`）用的是**占位路径** —— `command` 的 node/python、`args` 指向的仓库脚本（如 `bridge/vb6mcp_hub.mjs`）都要改成你的实际安装位置。JSON 里建议统一用正斜杠 `/`；若要写 Windows 反斜杠，必须写成 `\\`。

### Cursor（streamable HTTP）

Cursor 也可以直接连中间服务的 streamable HTTP 面：

```
类型: streamable HTTP
URL: http://127.0.0.1:8767/mcp
请求头: Authorization: Bearer <token>    # 仅当服务端配置了 token
```

（用 stdio 启动 hub 也可以，见上。）

### 其它 stdio 客户端

与 Claude Desktop 相同：`command: node`、`args: ["<仓库>/bridge/vb6mcp_hub.mjs"]`，按需设置 `VB6MCP_BRIDGE_URL` 与 `VB6MCP_TOKEN` / `VB6MCP_TOKEN_FILE`。

验证链路是否打通：`python tests/smoke_test.py --port <中间服务端口>`（该脚本**不发送 token**，只在服务端未配置凭据时可用；且需要**先打开 VB6 IDE 并加载插件** —— 这不是离线自检）。

## 四、工具列表（插件面 120 + hub 面 11 = 131 个）

> 权威清单见 **`docs/TOOL_CATALOG.md`**（由运行中的 add-in 经 `tools/list` 导出，再并入 hub 自有工具），
> 离线基线为 `tools/schema_baseline.json`，可用 `tools/check_live_tools.py docs/tools_schema.json` 做离线交叉比对。
>
> 工具来自**两个面**：**插件面 120 个**（`docs/tools_schema.json`，住在 VB6 IDE 进程内）与
> **hub 面 11 个**（`bridge/hub_tools.json`，在 IDE 进程外）。下面按功能把两个面共 **131 个**工具分组，
> 每个工具只出现一次：插件面分 16 类，hub 面见第 17 节；`关键参数` 一列直接取自各自的
> `inputSchema`（必填参数在前），不是凭记忆写的。hub 面的 5 个经验库工具**无需 IDE 在线即可用**。

### 1. 工程与工作区

| 工具 | 能干什么 | 关键参数 |
|------|----------|----------|
| `get_project_info` | 读工程全貌：类型、IDE 状态、启动对象、编译路径、组件与引用 | - |
| `set_project_options` | 设置工程选项：描述、启动对象、编译输出路径、只读模式 | `description`, `startup_object`, `build_file_name` |
| `open_project` | 打开 .vbp 工程，可独占替换当前工作区 | `path`, `exclusive` |
| `add_project` | 新建工程，或按模板创建工程 | `type`, `template_path` |
| `remove_project` | 从 IDE 工作区关闭一个工程 | `name` |
| `save_project` | 保存工程与全部未保存组件 | `path` |
| `vbg_group` | 读写 .vbg 工程组：列出、添加、移除工程 | `action`, `path`, `project` |
| `workspace_status` | 对比未保存组件与 git 工作区状态 | `repo` |

### 2. 组件生命周期

| 工具 | 能干什么 | 关键参数 |
|------|----------|----------|
| `get_component_list` | 列出全部组件及其过程、行号、脏状态 | - |
| `add_component` | 新建模块/类/窗体等组件 | `type`, `name`, `code` |
| `remove_component` | 从工程移除组件，不可撤销 | `name` |
| `import_component` | 把源文件或模板导入为工程组件 | `path`, `mode` |
| `export_component` | 把组件代码或整个设计器导出到文件 | `name`, `path` |
| `rename_component` | 重命名工程内的组件 | `name`, `new_name` |
| `reload_component` | 从磁盘上绑定的文件重新加载组件 | `name` |
| `activate_component` | 打开组件代码窗格或设计器并置前 | `name`, `window` |

### 3. 代码读取、检索与光标

| 工具 | 能干什么 | 关键参数 |
|------|----------|----------|
| `get_component_code` | 从 IDE 内存读组件代码，可只要某个过程或符号；`proc` / `symbol` 的范围同时给 IDE 行与磁盘行（`startLineDisk` / `endLineDisk`） | `name`, `include_line_numbers`, `coords` |
| `read_source` | 读磁盘源码：编码识别 + 字节级 sha256 | `path` |
| `read_file` | 读任意文本文件，UTF-8/UTF-16/ANSI 自动识别 | `path`, `max_lines` |
| `find_in_component` | 在组件内查找文本：整词、大小写、模式匹配；回 `hits[]` / `count` / `nextStartLine`，走完报 `no_more_hits` 且**不回绕** | `component`, `target`, `start_line` |
| `get_selection` | 读 IDE 当前选中的组件、活动窗口与光标 | - |
| `get_codepane` | 读代码窗格的选区、顶行与可见行数 | `component` |
| `set_codepane` | 设置代码窗格选区与滚动，并可置前 | `component`, `selection`, `top_line`, `activate` |
| `current_position` | 读断点停在哪：模块、过程、行号、源码行 | - |

### 4. 代码编辑与安全事务

| 工具 | 能干什么 | 关键参数 |
|------|----------|----------|
| `replace_code` | 行级改代码：替换、插入、追加、删除、整模块替换 | `name`, `operation`, `start_line` |
| `add_code` | 向组件追加代码，来自字符串或文件主体；落点由**写后回读**给出（`insertedAt{ide,disk}` / `inDeclarationArea`） | `component`, `code`, `file` |
| `create_event_proc` | 生成事件过程桩，如 Command1_Click | `component`, `object`, `event` |
| `write_file` | 写文本文件，可选编码、建目录与备份；回执含 `createdDirs[]` / `exists` / `size`，父目录建不出来报 `parent_dir_missing` | `path`, `content`, `encoding` |
| `write_source` | 写 VB 源码文件：显式编码 + 字节级 sha256；回执含 `createdDirs[]` / `exists` / `size`，父目录建不出来报 `parent_dir_missing` | `path`, `text`, `encoding` |
| `edit_by_anchor` | 按唯一内容锚点替换整行，不依赖行号 | `name`, `anchor`, `new_text` |
| `apply_edits` | 原子批量锚点编辑，任一处失败则整批不改 | `name`, `edits`, `expected_sha256` |
| `preview_edit` | 只回统一 diff 预览，不写任何内容 | `name`, `edits` |
| `replace_in_files` | 跨组件字面/正则替换，默认只计划不写 | `path`, `query`, `replacement` |
| `rename_symbol` | 按词法重命名标识符，跳过字符串与注释 | `path`, `old_name`, `new_name` |
| `move_procedure` | 事务化移动整个过程块（源删除 + 目标插入，任一侧失败双侧回滚） | `from_component`, `proc`, `to_component` |
| `delete_procedures` | 一次事务删除多个过程（每个受影响组件先写 undo 快照，失败全部恢复） | `procedures`, `component` |
| `edit_by_byte_anchor` | 按 GBK/ANSI **原始字节**锚点改文件（base64 进出；拒绝切开双字节字符，失败不写盘） | `path`, `old_bytes_b64`, `new_bytes_b64` |

### 5. 语法门禁与工程校验

| 工具 | 能干什么 | 关键参数 |
|------|----------|----------|
| `validate_code` | 写前静态校验代码，回带稳定指纹的诊断 | `code`, `name`, `kind` |
| `check_component` | 用同一套规则校验 IDE 里的组件文本；可选 `baseline` 做公开 API 比对（不可解析报 `baseline_unusable`） | `name`, `baseline` |
| `verify` | 组件总判定：IDE 与磁盘一致 + 诊断计数；**仅隐藏 Attribute 行**的差异单列 `attributeOnlyDiff` / `attributeLineCount` | `name`, `disk_path`, `baseline` |
| `verify_scope` | 批量核验多个组件的 IDE 与磁盘一致性；UNKNOWN 单列**不算 fail**（`pass + fail + unknown == checked`），build 证据按**项目**判定 | `components`, `path` |
| `format_code` | 按块结构重排缩进，并证明只改前导空白 | `path`, `indent`, `apply` |
| `check_references` | 列出工程引用并标出文件缺失的引用 | `include_resolved` |

### 6. 依赖、纯度与公开 API

| 工具 | 能干什么 | 关键参数 |
|------|----------|----------|
| `dependency_graph` | 组件与过程级依赖图：边、环、纯度分类；主截止是**确定性字符配额**（`VB6MCP_CHAR_QUOTA`，默认 900000），墙钟只作最后手段（`VB6MCP_BUDGET_MS`，默认 30000）。回执带 `partial` / `*StageCompleted` / `budgetExhausted` / `timingTruncated` / `procEdgesFound` 等完整性字段 —— **默认档是确定性的部分结果** | - |
| `purity_check` | 判定组件能否进离线测试宿主及原因 | `name` |
| `export_public_api` | 把组件公开 API 冻结成可比对规格；省略 `name` 时做有界整工程导出 | `name`, `membersOnly` |
| `check_public_api` | 与冻结规格比对，报移除、变更、新增成员；**给了 `name` 就永远比对**（`membersOnly` 只影响输出大小），不可解析的 baseline 报 `baseline_unusable`；省略 `name` 时做有界整工程核验 | `baseline`, `name`, `membersOnly` |
| `generate_stub` | 生成可编译桩模块：留声明、清过程体 | `name` |

### 7. 编译与交付产物

| 工具 | 能干什么 | 关键参数 |
|------|----------|----------|
| `make_project` | 编译工程并回权威产物判定，可选真日志诊断；构建状态可查询（`buildState`） | `output_path`, `laa`, `diagnose`, `wait_for_build` |
| `build` | 用隔离的 VB6.EXE /make 编译并回编译器日志 | `project`, `defines`, `laa` |
| `package_app` | 组装交付目录：exe、清单、依赖表、安装脚本 | `path`, `exe_path`, `out_dir` |
| `app_manifest` | 生成或读取 exe/dll 旁清单（UAC 与 DPI） | `action`, `path`, `level` |

### 8. 运行、调试与断点

| 工具 | 能干什么 | 关键参数 |
|------|----------|----------|
| `control_execution` | 控制运行：启动、全编译、中断、单步、结束、重启 | `mode` |
| `execution_state` | 读进程窗口、是否无响应、模态框数量 | `pid` |
| `debug_commands` | 列出本服务器认识的调试命令目录（固定表） | - |
| `set_breakpoint` | 注入 Debug.Assert False 断点，编译版自动省略 | `name`, `anchor`, `id` |
| `clear_breakpoints` | 按 id 移除注入的断点块 | `name`, `id`, `expected_sha256` |
| `list_breakpoints` | 列出注入的断点及其行号与启用状态 | - |
| `native_breakpoint` | 经调试菜单切换真实 IDE 断点或全部清除 | `action`, `component`, `line` |
| `fault_inject` | 在过程入口注入受控故障：抛错、返回、延时 | `name`, `procedure`, `mode`, `id` |
| `clear_faults` | 移除先前注入的故障块 | `name`, `id`, `expected_sha256` |
| `debug_session` | 一次调用跑完整个调试会话：设断点 → 不保存工程启动 → 等命中 → 按序执行动作 | `runArgs`, `actions` |
| `run_until_break` | 运行到断点 / 进程退出 / 未处理错误 / 超时（两段式：Start 是排队命令） | `timeout_ms`, `attach`, `breakpoints`, `watch` |
| `breakpoints` | 管理真实 IDE 断点：设/清/启用/禁用/列出（组件 + 源码锚点或 pane 行） | `action`, `component`, `anchor`, `line` |
| `get_locals` | 断点态读当前过程的参数与局部变量（名字来自源码，值来自引擎） | `frame_index` |
| `eval_at_frame` | 在 IDE 立即上下文求值（显式安全契约；对象/数组表达式先拒绝） | `expr` |
| `run_project` | 从 IDE 启动当前工程（可设环境变量 / 命令行参数；两段式） | `commandLineArgs`, `env`, `start_mode`, `observe`, `timeout_ms` |
| `last_error` | 读未处理错误的上下文：对话框正文 + 错误号 + 调用栈 + 局部变量名字 | `dismiss` |

### 9. 追踪、调用栈与进程内执行

| 工具 | 能干什么 | 关键参数 |
|------|----------|----------|
| `inject_trace` | 在过程入口插 McpTrace 调用，缺模块自动生成 | `name`, `procedure`, `id` |
| `remove_trace` | 移除注入的追踪块，自动生成的模块保留 | `name`, `id`, `expected_sha256` |
| `project_changes` | 轮询 IDE 变更日志：组件与控件增删改名 | `since`, `clear` |
| `get_callstack` | 读引擎自己的调用栈，逐帧给出过程名 | `max_frames` |
| `vba6_probe` | 进程内探 VBA6 调试引擎：模式、调用栈、立即执行 | `exec_probe`, `exec_line`, `run_break_probes` |
| `vba6_probe_plan` | 给出探针的有序执行计划与前置条件 | - |
| `immediate_eval` | 在立即上下文求值一个表达式并取回结果 | `expression`, `expression_file`, `timeout_ms` |
| `immediate_run` | 在立即上下文执行语句并取回打印输出 | `code`, `code_file`, `timeout_ms` |

### 10. 窗口、对话框与 UI 自动化

| 工具 | 能干什么 | 关键参数 |
|------|----------|----------|
| `list_windows` | 枚举 IDE 窗口：标题、可见性、位置与大小 | - |
| `control_window` | 操作单个 IDE 窗口：聚焦、关闭、显隐、移动、缩放 | `window`, `action` |
| `find_gui` | 递归枚举本进程窗口，按关键字或类名过滤 | `keyword`, `cls`, `showing` |
| `ui_text` | 读窗口及子树的可见文本 | `hwnd`, `cls`, `depth` |
| `read_active_dialogs` | 列出阻塞中对话框及其按钮，只读不点 | `cls`, `include_children`, `max` |
| `close_active_dialogs` | 只点安全按钮关掉阻塞对话框，可预演 | `hwnd`, `button`, `dry_run` |
| `send_window_message` | 向本进程窗口发消息：点击、按键或原始消息号 | `hwnd`, `action`, `message` |
| `command_bars` | 列出或执行 IDE 菜单命令，按名称或 id | `action`, `bar`, `control` |

### 11. 设计器与控件

| 工具 | 能干什么 | 关键参数 |
|------|----------|----------|
| `form_designer` | 窗体设计器：控件清单、属性、选区与剪贴板 | `form`, `action`, `control` |
| `add_control` | 按 ProgID 在窗体添加控件，先校验已注册 | `form`, `prog_id`, `name` |
| `remove_control` | 按名称或序号从窗体移除控件 | `form`, `index`, `control` |
| `set_control_property` | 设置控件的设计期属性 | `form`, `property`, `value` |
| `list_component_parts` | 列出工程注册的 ActiveX 控件部件与解析路径 | - |

### 12. 引用、成员与属性

| 工具 | 能干什么 | 关键参数 |
|------|----------|----------|
| `add_reference` | 按 GUID 或文件路径添加类型库引用 | `guid`, `path`, `mode` |
| `remove_reference` | 按名称或 GUID 移除工程引用 | `name` |
| `get_component_properties` | 读组件设计期属性，即属性窗口内容 | `name` |
| `set_component_property` | 设置组件属性，支持多维索引属性 | `name`, `property`, `value` |
| `list_members` | 列出类组件的公开成员及代码位置 | `component` |
| `set_member_attribute` | 改类成员属性：说明、Hidden、Category 等 | `component`, `member`, `attribute`, `value` |
| `property_bag` | 读写工程或组件的属性包键值 | `target`, `key`, `name` |

### 13. 磁盘文件与工程元数据

| 工具 | 能干什么 | 关键参数 |
|------|----------|----------|
| `list_files` | 按通配符递归列出目录下的文件 | `path`, `pattern`, `recursive` |
| `search_in_files` | 跨文件文本或正则搜索，回文件名与行号 | `query`, `path`, `pattern` |
| `project_meta` | 读写 .vbp 版本资源与构建元数据 | `action`, `path`, `major` |
| `git` | 只读 git：状态、diff --stat、提交日志 | `action`, `repo`, `file` |

### 14. 快照、回滚与测试宿主

| 工具 | 能干什么 | 关键参数 |
|------|----------|----------|
| `project_snapshot` | 工程级快照与回滚，带逐文件清单 | `action`, `path`, `dir` |
| `undo_edit` | 组件级单步撤销，只恢复窗格文本不动磁盘 | `component`, `action`, `ts` |
| `test_host` | 生成离线测试宿主 .vbp，或统计 PASS/FAIL | `action`, `vbp_path`, `module_name` |
| `record_fixture` | 以原始 ANSI 字节记录回归夹具 | `path`, `text` |
| `compare_fixture` | 与夹具比对：相同、新增、删除、首个差异行 | `path`, `text` |

### 15. 扩展库

- **扩展库** `mcp_extensions/`：每个 `*.json` 是一个**声明式工具包**（`name` / `tools[].inputSchema` / `tools[].handler`）。
  启用的包其工具自动并入 `tools/list` 并由 `CodeTools.CallTool` 兜底分发；`state.json` 记录禁用列表（默认启用）。
  处理器只开放少量安全原语（`http_get` / `read_file` / `write_file`），**不从磁盘加载 VB6 代码**。

| 工具 | 能干什么 | 关键参数 |
|------|----------|----------|
| `ext_list` | 列出扩展包及其启用状态与工具数 | - |
| `ext_enable` | 启用扩展包，其工具并入 tools/list | `name` |
| `ext_disable` | 禁用扩展包，其工具从 tools/list 消失 | `name` |
| `ext_reload` | 重扫 mcp_extensions 并报告包与工具数 | - |
| `ext_install` | 把扩展 JSON 安装进 mcp_extensions | `path` |
| `ext_uninstall` | 移除扩展包，文件改名保留 | `name` |

### 16. 治理、自检与 IDE 环境

| 工具 | 能干什么 | 关键参数 |
|------|----------|----------|
| `mcp_info` | 读服务器版本戳、契约、工具数与构建哈希；`build.sourceHash` / `sourceAlgo` 是源码输入集合的指纹与算法标识 | - |
| `ide_state` | 即时读 IDE 模式与在飞、取消、释放闸门状态（`cancel_inflight` / `release_inflight` 为**尽力而为**：不能中止陷在宿主模态里的调用） | `cancel_inflight`, `release_inflight`, `release_min_held_ms` |
| `desktop_probe` | 测量宿主进程桌面与可交互性，且不抢焦点 | - |
| `ide_quiet_mode` | 开或关安静模式：不抢前台，回执附提示 | `on`, `ttl_ms` |
| `todo_status` | 只读扫描仓库内的任务板与报告 | - |
| `audit_tail` | 读最近的工具调用审计条目 | `count` |
| `list_addins` | 列出已注册插件及其连接状态 | - |
| `set_addin_state` | 加载或卸载一个 IDE 插件 | `prog_id`, `connect` |
| `quit_ide` | 退出 IDE，响应后延迟执行，未保存会丢 | - |

### 17. hub 面：IDE 实例管理与经验库（11 个，在 IDE 进程之外运行）

这 11 个工具由 MCP hub（`bridge/vb6mcp_hub.mjs`）提供，**在 IDE 进程之外**运行：客户端不必先加载插件。其中 `ide_list`、`ide_target` 与 5 个经验库工具**不依赖任何 IDE**（经验库只读写本地 JSON）；`ide_stop` / `ide_restart` 需要**一个目标实例**，`ide_start` 需要**能启动 VB6 的条件**。逐项清单与两个拒绝码的前置条件见「二、连接模型」里的「哪些工具免 IDE、哪些必须有活 IDE」。

- **经验库** `mcp_lessons/`：`defaults.json`（随包发布，勿改）+ `user.json`（用户层，按 id 覆盖/禁用/新增）。
  合并结果经 `initialize.instructions` 发布给客户端，agent 调工具前即可读到；`add/delete/import` 写 `user.json` 时带 `.bak` 备份。
- **IDE 实例管理**：`ide_list` 列出 hub 能寻址的全部 VB6 IDE：本机实例（含同一工程被两个实例打开的冲突），`ide_target` 记住默认目标，
  `ide_start` 负责启动；`ide_stop` / `ide_restart` 只操作**本 hub 自己启动的**实例；同时列出**已配置的远端端点**（只探白名单、**绝不自动发现**；连不上也照样列出并带探测状态），并回显实例的**唯一身份**（endpoint / targetKey / pathVia）、真实路径与登记表拼写、以及**解析到同一身份的调用者别名**（junction / 别名写法）。

| 工具 | 能干什么 | 关键参数 |
|------|----------|----------|
| `ide_list` | 列出 hub 能寻址的**全部** VB6 IDE：本机实例 + **已配置的远端端点**（只探白名单、绝不自动发现；连不上也列出并带探测状态）；回报 pid/端口/工程、唯一身份（endpoint / targetKey / pathVia）、真实路径与登记表拼写，并回显解析到同一身份的**调用者别名**；含过期登记项与「同一身份两个实例」冲突 | - |
| `ide_target` | 设置或清除默认 IDE 目标：pid、工程名或 auto | `target` |
| `ide_start` | 按 .vbp 启动一个 VB6 IDE 实例，并轮询到插件就绪（`ready` 才是真就绪） | `project`, `ready_wait_ms` |
| `ide_stop` | 停止本 hub 启动的 IDE 实例（需要一个目标实例） | `pid`, `force` |
| `ide_restart` | 重启卡死的自有 IDE 实例并证明已就绪（需要一个目标实例）；长重启可 `wait_ms` 到点回 `pending`+`opId`，再用 `opId` 续查（只查不重启） | `pid`, `project`, `opId`, `wait_ms` |
| `ide_kill_and_relaunch` | 管道 busy/wedged 时的带外复位：只停本 hub 自有实例、清注册项、原 .vbp 重启并轮询就绪（不依赖插件管道） | `pid`, `project`, `ready_wait_ms` |
| `list_lessons` | 列出经验条目：场景、陷阱、做法（hub 面，无需 IDE 在线） | `includeDisabled` |
| `add_lesson` | 新增或覆盖一条用户经验，带备份（无需 IDE 在线） | `id`, `scenario`, `trap`, `do` |
| `delete_lesson` | 删除用户经验，或禁用随包经验（无需 IDE 在线） | `id` |
| `export_lessons` | 把合并后的经验库导出为 Markdown（无需 IDE 在线） | `includeDisabled` |
| `import_lessons_md` | 从 Markdown 导入经验，可只解析不写（无需 IDE 在线） | `path`, `dryRun` |

> `get_codepane` / `set_codepane` **没有** `view` 参数：VB6 的可扩展模型把代码窗格的视图模式
> （`CodePaneView`）定为只读、命令栏里也没有视图切换项，插件无法在 full_module / procedure 之间切换；
> 传入 `view` 会被记入 `ignoredArgs`。

完整对象模型清单与差距分析见 `内部 VBIDE 对象模型笔记`（由 VB6EXT.OLB 5.3
类型库二进制解析生成，原始数据见 `内部类型库转储`）。

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
- 插件是**单线程**：一次调用还在飞行时，**任何第二次调用（包括只读）都不会被执行**，而是立即收到 **busy 拒绝回执**；回执字段 `readonly_serviceable` **恒为 `false`**。`cancel_inflight` / `release_inflight` 只是**尽力而为**——陷在**宿主模态循环**（例如 VB6 自己的对话框）里的调用**无法被中止**；**唯一可靠出路是等它结束，或从外部终止 VB6 进程**。
- 超时参数**不是强制预算**：`build` / `make_project` 的 `timeout_ms`（以及 `immediate_*` 的 `timeout_ms`）**不会**在到点时中断正在执行的调用——它们只决定调用方等多久、或作为闸门年龄阈值。实测传入远小于真实耗时的值仍可成功返回（此时 `budgetMs` 可能为 `null`）。要中止正在跑的调用，请**从外部终止 VB6 进程**。
- 断点状态（break mode）下 `make_project` / `control_execution` 可能失败，
  工具会返回错误文本。
- `make_project` 对当前活动工程执行；无法静默指定输出路径（跟随 IDE 的工程设置）。
- 插件自身 DLL 被 IDE 占用时无法原地重建（见 FAQ）。
- 服务器接受连接时强制校验对端为 127.0.0.1（MSWINSCK.OCX 的 LocalIP 只读，
  无法在监听前绑定回环地址，故用接受时过滤实现），无鉴权——只信任本机进程。

### 本地化与英文版 IDE 支持

- **结论**：英文版（en-US）VB6 在**能力上支持**，但**未做端到端实测**（我们的验证机是中文版 zh-CN，见本节末尾）。
- **与语言无关的部分**：菜单/调试命令按 caption 匹配、忽略本地化 bar 名（`Start`/`Continue`/`End`、`Toggle Breakpoint`、`Step Into/Over`、`Run To Cursor`、`Start With Full Compile`、`Clear All Breakpoints` 等）；断点开关与调试命令的中英标题都在匹配表内；IDE 模式标记中英同集且大小写不敏感（`[设计]`/`[运行]`/`[Design]`/`[Run]`）；对话框按钮**控件 ID 优先、中英文本兜底**；运行时错误框的英文措辞（`RUN-TIME ERROR` / `RUNTIME ERROR`）已识别；注释/诊断语言**界面优先**（英文 IDE 得到英文注释）；hub/bridge 侧只假设 `VB98` 目录名，与语言无关。
- **仍需注意**：个别流程的**窗口标题**匹配与少数弹窗仍以中文为基准；英文版首次使用建议先跑 `mcp_info` / `ide_state` / `command_bars{action:"list"}` 三条只读调用，确认工具面与状态。
- **欢迎反馈**：若你在英文版（或其它语言版本）VB6 上使用，请报告 **tool 名 + 回执 + IDE 界面语言**——即使只是「可用/不可用」也是极有价值的贡献。

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
  在部分机器上 `SystemRoot` **不是** `C:\Windows`，而 `C:\Windows` 可能是**一份陈旧的第二安装** ⇒ 用它会让
  **任何** 32 位 DLL 注册失败（`0xC0000139`）且 `/s` **静默**（`build.bat` 已用 `%REGSVR%` 显式指向本机的 SysWOW64）。
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
`"%SystemRoot%\SysWOW64\regsvr32.exe" MSWINSCK.OCX`（★不要硬编码 `C:\Windows`，理由见上文）。

**桥接器报 "VB6 MCP server unreachable"**
→ 确认 VB6 IDE 已打开且插件已加载（Add-Ins → VB6 MCP Status... 查看端口），
端口与 `%TEMP%\vb6mcp_port.txt` 一致；防火墙若拦截本机回环，放行 python。

**写文件后 IDE 里没变化**
→ 属于预期行为，见第五节；执行 File → Reload Project。

**想在未打开 VB6 IDE 时也能用**
→ 不行：本插件依赖 VBIDE 对象模型，必须随 IDE 运行。日常做法是让 VB6 常驻后台。

**Agent 正在操作 IDE 时，我能在同一台机器的 IDE 里打字吗？**
→ **不要。** 这不是建议而是纪律：插件驱动 IDE 时会设置光标、滚动、打开窗格，有的动作还会把窗口提到前台
（`control_window{action:"activate"}`；`native_breakpoint` 会先把光标移到目标行）。**这时你的按键会落到「当时前台的那个窗口」**——
本项目实测过一次真实事故：用户在 Agent 操作期间打字，击键落进了代码窗，源码里多出一行游离字符
（那是**击键**写进去的，不是工具写的）。要在同一台机器上手工编辑，请**先让 Agent 停手**，或用**另一台机器 / 另一个实例**。

写类工具为此加了护栏（回执里都能看到）：

- **写后回读断言**：编辑类工具本来就有 `expected_sha256` / `sha256Before` / `sha256After`；**不写文本**的断点与窗口类工具
  （`native_breakpoint` / `control_window` / `set_codepane`）现在也在操作前后各取一次窗格文本的 sha256，**两者必须相同**：
  一旦不同，回执被翻成 `{ok:false, reason:"paneTextMismatch"}` 并附**逐行差异**（`paneDiffLines`）。
  窗格读不到时如实标注 `paneAssertion:"notApplicable"`，**不会**当成「通过」。
- `set_member_attribute` 写后**回读**该成员属性，回执给 `readback` 与 `applied`（★那是**内存**里的值；能否落成磁盘上的 `Attribute` 行由 VB6 在保存组件时决定，见 FAQ §50）。
- **焦点安全**：`native_breakpoint` / `control_window` 的回执带 `foregroundBefore` / `foregroundAfter` / `foregroundTookFocus`，
  且**非 activate 动作会把前台还给原来的窗口**。
- **断点类没有第三条路**（如实写在回执里）：VB6 **没有断点 API**，Debug 菜单项只作用于**光标**，另一条路 `set_breakpoint` 是**写文本**注入标记块。
  所以 `native_breakpoint` 采用**补偿**：操作后把光标放回原处（`cursorRestored`）。`breakpointWithoutCursorMove:false` 就是这个事实的机器可读标记。

**文件里改的 `.vbp` 为什么没生效？`open_project` 现在会告诉你**
→ 已打开的工程以前只回 `alreadyOpen:true`，读起来像「你指名的文件被加载了」，其实**没有重读磁盘**。现在回执区分
`reloaded` / `keptInMemory`，回显 IDE **内存里**的工程属性（`inMemory{}`），并把 `.vbp` 与内存做比较（`diskVsMemory{}`）。
两者真的不一致时**一律结构化拒绝**（`ok:false` + 结构化 `reason`），绝不静默忽略，也**不会为了让磁盘生效而去关工程**：
- 内存里有未保存改动 ⇒ `reason:"unsaved_changes_would_be_lost"`；
- 内存干净 ⇒ `reason:"disk_vs_memory_mismatch"`；
- `VBProject.Saved` 读不到 ⇒ `reason:"project_saved_flag_unreadable"`。

★**为什么不做"真重载"**：本机实测（2026-10-13）关掉再重开（`VBProjects.Remove` → `AddFromFile`）会抛 **59999**「方法 '~' 作用于对象 '~' 失败」，
**把工作区丢成 0 个工程**，紧接着再 `AddFromFile` 仍然失败 ⇒ 工具**不会为了刷新而摧毁工作区**，回执带 `reloadNotAttempted:true` 与说明。
要采用磁盘版本，请自己走 `remove_project` → `open_project`。

★**明确不提供的能力（按设计拒绝，不是缺陷）**：本插件**没有**「把 `.vbp` 层面的工程属性写进 IDE 内存」或「原地重载一个已打开的工程」的通道 —— 唯一能实现后者的机制就是上面那条会毁工作区的路径，所以它**从不尝试**：一旦「磁盘与内存不一致」，回执就是**结构化拒绝**（`ok:false` + `reason` + `reloadNotAttempted:true` + `sideEffects:"none: ..."`），而不是静默 no-op。**能**写进 IDE 的工程选项（描述、帮助文件/上下文、启动工程、编译输出路径、活动工程、只读模式）由 `set_project_options` 提供；该构建做不到的（如 `startup_object`）回**结构化的** `startupObjectError` + `startupObjectWorkaround`，**不会假装成功**。`.vbp` 层面的版本资源等键由 `project_meta{action:"set"}` 写**文件**（带 `.bak` 与写入后 sha256），不会注入到 IDE 内存。

**`coords:"disk"` 的行号为什么不是连续的？**
→ 因为 VBA6 会把**成员 Attribute 行**（形如 `Attribute Foo.VB_UserMemId = -4`）留在磁盘上、却**不显示在代码窗格里**。
磁盘行号因此是**逐行累计**出来的（`diskLine = ideLine + headerLines + 该行之上被隐藏的 Attribute 行数`），
跳过隐藏行时行号自然不连续。回执带 `mappingBasis` / `mappingResolved` / `hiddenInlineAttributeLines` 与
`hiddenInlineAttributeDiskLines`（被隐藏的磁盘行号），对齐失败时如实标 `mappingResolved:false`，**不假装精确**。

**`build` 的 `/d` 宏为什么偶尔不生效？**
→ 已修：**仅当值里含空格才加引号**（无条件加引号会让编译器把引号折进宏名，宏静默失效而构建仍报成功）。
回执给 `commandLine`（**实际命令行**）、`conditionalCompileArgs`（结构化的 `name` / `value` / `quoted`）、
`conditionalCompileArgsRaw` 与 `conditionalCompileQuoteRule`。

**`vba6_probe{exec_probe}` 报错时，怎么区分「语句抛错」和「通道挂了」？**
→ 两种原因已分开：**`exec_error`** = 插件**答了**（语句抛错，或 EbExecuteLine 报了 VB 错误）⇒ 回执带结构化的
`exec_error{vbaErrNumber, description}`；**`forward_timeout` / `forward_failed`** = 由**hub**（`bridge/vb6mcp_hub.mjs`）报告，
含义是插件**根本没答**（通道挂住或超出转发预算），此时不存在插件回执，也就不可能有 `exec_error`。

## 九、协议说明

- 实现 MCP 核心：`initialize` / `ping` / `tools/list` / `tools/call` / 通知。
- 协议版本**协商**：服务端支持 2025-11-25 等版本，按客户端 `initialize` 回显其请求的版本
  （未带则用最新；见 `McpServer.NegotiateProtocol`）。
- 传输：中间服务提供 `POST /rpc` 与 `POST /mcp`（streamable HTTP）；请求头带 `Accept: text/event-stream` 时用 SSE 帧返回单条消息。旧代的 `GET /sse` + `POST /message` 长连订阅**不再实现**（返回 501），因此不使用服务端主动通知（notifications）——客户端按需 `tools/list` 拉取即可。
- 鉴权：配置了凭据后**所有 HTTP 面**都要 `Authorization: Bearer <token>` 或 `X-VB6MCP-Token`；把凭据放进 URL 一律 401 `token_in_url_rejected`。
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

## hub 侧恢复原语与可靠性开关（v1.13 ~ v1.16）

> 这一节从**用户视角**说明这批能力解决什么故障、什么时候用、回执里该看哪些字。逐字段契约见 `docs/CONTRACTS.md` 的 v1.13.0 ~ v1.16.5 版本行。

1. **`ide_kill_and_relaunch`（v1.13.0，hub 侧、不依赖插件管道）** —— 管道 busy / wedged 时**唯一**的恢复原语：结束**本 hub 自有**的实例、清掉它的注册项、在原 `.vbp` 上重启并轮询就绪；只停 hub 自有实例，**绝不碰 8766、远端实例或别人启动的 IDE**。回执看 `ready` / `threeWay{http,mcpInfo,project}` / `ownership`；`ready:false` 只表示"进程已拉起但插件还没应答"，照 `howToProceed` 轮询 `ide_list`，不要重试 kill。
   **为什么要它**：插件是单线程的，管道 busy 时连 `ide_state{release_inflight:true}` 都送不进去；以前唯一出路是人工 `Stop-Process` 再重开 IDE，现在一次调用即可，而且不动别人的实例。

2. **`ide_restart` 的 `wait_ms` + `opId`（v1.14.0）** —— 重启可能比客户端超时长：`wait_ms` 限制**单次调用**阻塞多久，到点回 `pending:true + state:"in_progress" + opId`；拿 `opId` 再调 `ide_restart{opId}` 只查询、**不启动任何东西**（操作注册表保留 30 分钟，hub 重启即忘）。同一个 hub **同一时刻只允许一个重启**，进行中再调会被 `restart_in_progress` 拒绝。
   **为什么要它**：重启实际已完成、客户端却只收到 `Request timed out`（无回执），调用方只能去查进程；盲目重试又会二次重启。

3. **`ide_restart` 的 content 单 JSON 文档（v1.14.0）** —— `content[0].text` 现在是**一个** JSON 文档：数组/标量 body 变为 `{singleDocument:true, data, receipt}`（空结果 = `data: []`，与解析失败可区分），非 JSON body 给结构化 `content_not_json`，对象 body 维持扁平；`VB6MCP_LEGACY_TEXTBODY=1` 可临时回到旧两段式（只保留一个版本）。传输失败且 peer 可证已死时给 `peerState:"peer_dead"` + `lostContext` + `howToRecover[]`。
   **为什么要它**：旧版把"JSON 数组 + 追加回执对象"拼成两段，严格客户端 `JSON.parse` 直接失败 —— `find_gui` 在无命中（`[]`）时甚至无法区分"空结果"与"解析失败"。

4. **`build` / `make_project` 的离线路由（v1.13.0）** —— 没有活 IDE 时，`build` / `make_project(route!="ide")` / `project_meta` 由 hub 自己服务：`offline:true + offlineRoute:"cli"|"file" + targetResolved:false`，走独立的 `VB6.EXE /make`；不能启动编译器时回 `reason:"cli_unavailable"` + `howToProceed`，**不再**把"没有 IDE"当成 `target_mismatch`。0 实例时 `package_app`/`verify` 明确 `offlineServed:false + no_live_ide`，不假装离线可用。
   **为什么要它**：离线编译是唯一能拿到真编译器日志的通道；以前 IDE 一关，`build` 仍被陈旧目标挡住，只能绕过 MCP 手跑 `/make`。

5. **`verify` / `verify_scope` 的 `requireCompiler`（v1.16.0）** —— 默认 false 时回执逐字节不变；true 时要么给出**编译器权威**（`authority:"compiler"` + `errorCount` + `artifactChanged` + `compiler{...}`），要么结构化拒绝 `compiler_not_configured` + `compilerProbe{...}` + `howToProceed`，**永不用 lint 顶替编译**。
   **为什么要它**：lint 只做词法与结构，曾对大模块报 49 个 error，而同一时刻 `/make` 是 0 error；把"疑似"当"编译结论"会给出与事实相反的假红。

6. **`apply_edits` / `preview_edit` 多组件批量（v1.16.0，仍是同一个工具）** —— `edits[].component` 让一批编辑跨最多 40 个组件；`expected_shas[]` 任一不符即整批拒绝、一字节不写；顺序语义 = 按 `edits[]` 顺序、每条锚点在该组件**已更新文本**上解析；失败**全或无**回滚（`rolledBack` + `restoreVerified` + `restoredComponents[]`）；`max_ops` 默认 200 / 上限 1000。`preview_edit` 用同一形状按组件分组 hunks 且**什么都不写**。
   **为什么要它**：跨文件改名/搬迁原来要十几次单组件调用，中间失败会留下"改了一半"的状态；多组件批量把一批编辑变成一个事务，且先 preview 能看清每处落地。

7. **`breakpoints{action:"probe"}` 两阶段行为验证（v1.16.0）** —— 唯一能回答"这个断点到底生效吗"的通道：`phase:"start"`（默认）检查设计态 / 无模态框 / 可断行后**启动程序**（不保存、不写源码），回 `probeId + startedPending + howToContinue`；`phase:"observe"` 用同一 `probeId` 等停、读 `position` + `callstack`、End 并用回读证明 `stateAfter=design`。`probeId` 单次使用、TTL 600000 ms；`probeResult` 是闭合枚举。同一版里 `control_execution{mode:"end"|"stop"}` 在 RUN 态**不再假拒绝**（旧的 `commandbars_unavailable_while_running` 是伪拒绝），`programEnded:true` 只在 design 回读之后出现。
   **为什么要它**：native 断点"记录成功但从不命中"，injected 断点写盘并两次搞死 IDE；只有真的跑一次、停在目标行，才能证明它有效。

8. **`run_project` 的 observe 预算与完成标记（v1.15.0）** —— `observe_max_wait_ms`（默认 `timeout_ms`，夹取 1000..120000）到点**返回**：`state` 是封闭五值 `done-by-marker | exited | break | error | timeout`；`done_marker{path,contains}` 可在 IDE 机器（允许 UNC）上等一个文件标记，只读前 65536 字节、不写文件。★`timeout` 是 `ok:true + timeout:true + retrySafe:true + suggestedAction:"observe_again"`（**不是失败**）；break 态 observe 明确拒绝（`observe_in_break_not_supported`）；IDE 自己的查找框会被分类成 `observe_blocked_by_ide_dialog`，不再误报成程序错误。
   **为什么要它**：回放 7 秒就结束，工具却报 `Request timed out` 并把管道留在 busy；完成标记让"已经跑完"可判定，而不是靠外部轮询。

9. **`get_locals{allowUnsafeInBreak}`（v1.15.0）与 `last_error` 结构化错误（v1.16.3）** —— break 态 `get_locals` 默认**拒绝**（`break_eval_requires_opt_in`）；`true` 是知情同意，**不保证**不挂，恢复走 hub 的 `ide_kill_and_relaunch`。`last_error`（以及 `ide_state.lastError`）在错误框**仍然存在**时新增只读结构化字段 `err{number,description,source,hresult,numberBasis}` / `errBasis` / `errChannel{...}` / `errCrossCheck` / `errLocation{...}`；它**不抑制、不自动关闭、不无框读取**错误框，读不到的字段是 null + 原因，绝不猜。
   **为什么要它**：break 态读局部变量可能把单线程管道 wedged，必须让调用方显式选择并知道恢复路径；错误框的本地化文本以前是唯一来源，读不到就只能外部截图。

10. **`quit_ide` 的 `exitWrites`（v1.16.4）** —— `requested:true` 的回执新增 `exitWrites{declaration,observed,condition,evidence,basis}`：它是**声明**，不是测量（进程无法观测自身退出阶段）。`condition` 只按当前可见工程数判定：0 个工程 ⇒ `no_project_observed_no_write`（实测未写）；≥2 个工程 ⇒ `project_group_context_then_clean_quit`（约 250 ms 后出现 50~53 B 的 `.vbg`）；恰好 1 个 ⇒ `undetermined`（两个实测形状相反）。`sideEffects.filesWritten` 只覆盖调用期间，**不覆盖退出阶段**。
   **为什么要它**：IDE 退出阶段可能写 `.vbg`，而工具既不能确认也不能阻止；与其给一个"没写"的假保证，不如把能观察到的条件、证据和不确定都写出来。

> 以上能力**都不改变既有默认行为**（不传新参数 / 不用 probe / 不 requireCompiler 时，回执与旧版一致）；契约与冻结面变化见 `docs/CONTRACTS.md`。

---

## `modMcpTrace` 是干什么的？必须要有吗？

**① 它是干什么用的**：一个很小的标准模块，里面只有 `McpTrace` / `McpTraceOn` 两个过程（外加开关变量）。它只服务于 `inject_trace` —— 被注入的追踪调用最终落在它上面，负责把一行 JSON 写进追踪文件。

**② 是否必须**：**VB6MCP 自身不需要它。** 插件本身（编辑、检查、编译、运行、断点、弹窗、窗体、快照、打包等）与它完全无关；**没有用过 `inject_trace`、也没有保留注入调用的工程里本来就不该有它**。

**③ 只有什么情况需要**：只有当你**用过 `inject_trace` 并且保留了那些注入调用**时才需要 —— 注入块调用的 `McpTrace` / `McpTraceOn` 就定义在它里面。★**保留注入调用并编译/发布 EXE 或 DLL 时，必须把该模块一起编进去**（回执里的 `shipNote` 同样说明）。

**④ MCP 怎么添加它**：`inject_trace` 发现目标工程里没有 `modMcpTrace` 时，会**自己按 IDE 语言生成并附加**（模块源码内嵌在 VB6MCP.dll 里，**免手工导入**），附加后还会复验它确实在工程里；**模块已存在则直接跳过**（幂等）。回执带 `moduleAutoAttached` / `moduleAttachAction` / `moduleAttachMs`；附加失败时带 `moduleAttachError`（失败一律如实报告，不静默降级）。

- **注释语言怎么选（界面优先）**：先看 **IDE 自己的界面语言**（主窗口菜单文案 文件/编辑/视图/帮助，其次标题里的模式标记 `[设计]`/`[运行]`）——能可靠读出就用它；读不出来再回退到 VB6.EXE 的版本资源语言 ID（`\VarFileInfo\Translation`，主语言中文 LANG_CHINESE ⇒ 中文）；两者都拿不到 ⇒ **英文**（英文在任何代码页都可读，中文在非中文系统会乱码）。回执里并排给出三个信号：`ideLanguage`（版本资源）· `ideMenuLanguage`（界面）· `commentLanguage`（最终选择），以及各自的 `*Source` 依据。

**⑤ 什么情况会添加／移除它的「引用」**：

- **添加**：首次注入且工程内没有该模块时（见 ④）。
- **移除**：**新版不再自动移除**。`remove_trace` 只删除**该 id 的注入块**，`modMcpTrace` **一律保留**；回执如实给出 `moduleRemoved:false` · `moduleLeftInPlace:true` · `moduleRemoveReason:"kept_by_design"`（工程里根本没有该模块时为 `module_absent`），并由 `moduleRemoveNote` 写明下面两条理由与手动移除指引。
- **旧行为已废除（及原因）**：早期版本曾让 `remove_trace` 顺带删除**带自动生成标记**的 `modMcpTrace`（`moduleRemoveReason:"removed_autogen_module"`）。该分支已**删除**，因为：(a) 用户可能在自动生成的模块里加了自己的代码，删模块会把那些代码一起毁掉；(b) 工程别处可能仍有 `@mcp-trace` 注入块，模块一旦被删，那些调用找不到 `McpTrace`/`McpTraceOn`，**工程将编译不过**。
- **需要手动清理时**：先确认全工程**没有** `@mcp-trace` 残留，再自行调用 `remove_component{name:"modMcpTrace"}`。注意 `VBComponents.Remove` 只从工程里移除组件，**不会删除磁盘上的 .bas 文件**（回执用 `moduleFileOnDisk` 如实标注它绑定在哪）。
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

## 工具回执的失败面与行号口径（v1.11.6 ~ v1.11.7）

这一代给「**没做成的事必须说出来**」补了几处此前会把失败静默掉的地方。

- **未声明的参数名一律拒绝**：`export_public_api` / `check_public_api` / `validate_code` / `check_component` 先查参数名，未声明 ⇒ `ok:false` + `reason:"unknown_argument"` + `unknownArguments[]` + `acceptedArguments[]`（`_target` 是传输上下文，恒接受）。整个工具面做同一件事之后，才不会有「这个工具查、那个工具不查」的盲区；`reason` 还细分 `unknown_argument` / `component_not_found` / `target_has_no_project` / `ide_not_ready`（旧值保留在 `reasonLegacy`）。
- **比对永不跳过**：`check_public_api` 只要给了 `name` 就**永远比对**，`membersOnly` 只影响输出大小；给了一个不是规格文本的 `baseline` ⇒ `ok:false` + `reason:"baseline_unusable"`（不再回「没有差异」）。组件级与整工程级两个入口回同一组 `comparedAgainstBaseline` / `baselineLines`；`check_component` 也接受可选 `baseline`。
- **写文件回读**：`write_file` / `write_source` 回执带 `createdDirs[]` / `createdDirsCount` / `exists` / `size`；父目录建不出来 ⇒ `reason:"parent_dir_missing"`（`missingPath` + `reasonLegacy:"parent_not_creatable"`），不再返回一个「写了多少字节」的自洽成功形状。
- **分页有终点**：`find_in_component` 回 `hits[]` / `count` / `nextStartLine`；走到末尾 ⇒ `found:false` + `reason:"no_more_hits"`，**不回绕**（连续取页不会又把前面的行交回来）。
- **隐藏 Attribute 行单列**：`verify` 对「只差隐藏 Attribute 行」的面板/文件差异给 `attributeOnlyDiff` / `attributeLineCount`，与真实源码差异**分开计**，比较时跳过这些行。注意：经 VBIDE `SaveAs` 保存会把隐藏 Attribute 行**物化到磁盘**。
- **两个行号空间**：`get_component_code{proc|symbol}` 的范围字段是 **IDE 行**，另给 `startLineDisk` / `endLineDisk`；`add_code` 的落点由**写后回读**量出，给出 `insertedAt{ide,disk}` 与 `inDeclarationArea`（真正的声明区：`insertedAt.ide` 小于第一个过程头行；**末尾追加为 `false`**）。
- **范围形状自带模式**：`requestedRange` 每个形状都带 `mode` —— `range`（`replace_code` / `add_code` / 守卫拒绝）、`anchor`（`edit_by_anchor`）、`anchor-batch`（`apply_edits`），调用方一段代码即可判定。
- **返回信封**：`result.content[0].text` 是**一个 JSON 字符串**（这一条已写入契约）。

## `verify_scope` 的计数口径：UNKNOWN 不是 fail；build 证据是【项目级】

- **UNKNOWN 单列**：`verify_scope` 的三个桶互斥 —— `pass` / `fail` / `unknown`，且不变式 **`pass + fail + unknown == checked`** 恒成立（回执带 `countsConsistent`）；`failures[]` 只装**已证实**的 IDE 与磁盘不一致。`unknown` 的含义是「**没有足够证据**」，不是「不一致」。
- **build 证据按【项目】判定**：`buildOk` / `buildSource` / `buildNote` 来自**项目级** build 缓存（key 与 hash 都是项目级），**单个组件不可能「仅自己缺 build 证据」**。因此 `unknown` 是**项目级**结论，`pass` 只在**有新鲜 build 缓存的工程**上可达 —— 从未构建过的工程里，任何 IDE 与磁盘一致的组件都会是 `unknown`。
- **为什么值得单独写**：把 UNKNOWN 当 fail，会让一次「没测到」看起来像一次「不一致」；把项目级证据当组件级用，会读出一个不可能的结论。两种误读都会让调用方去修一个不存在的问题。

## 传输可靠性：间歇断连与大于 64 KiB 的响应截断已修复

- 曾经有两个看起来像「服务偶发不稳定」的症状：**间歇性的 `RemoteDisconnected`**，以及**大响应在 64/96 KiB 处被截断**。它们不是环境噪声，而是三个**确定性**缺陷叠在一起：
  1. **连接对象未被持久化** ⇒ 一条请求跨 TCP 段到达时会丢掉前半段，要等 30 s 空闲清扫才断开；
  2. **close-after-send 的判定顺序**放在队列已经清空之后 ⇒ `Connection: close` 的响应会滞留到空闲清扫才被发出；
  3. **flush 的累积 `Mid$` 起点越过 64 KiB**，且 `send()=-1` 伴随 `WSAGetLastError()=0` 被误判为致命错误 ⇒ 正文被截断。
- 三处都已修好。**工具面与参数表没有因此变化**（当时插件 110 + hub 10 = 120，不变）：这是一次**行为修复**，升级后不需要改任何调用方式。

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
因此 tools/ 下的离线门禁脚本可以直接运行（`check_schema_baseline` / `check_hub_alignment` / `check_live_tools docs/tools_schema.json`）；需要活 IDE 的 `mcp_strict_audit` 与 `smoke_test` 要先把 VB6 跑起来。

## 如何参与

见 CONTRIBUTING.zh-CN.md（构建、注册、门禁、编码规则）。

**本地化**：英文版 VB6 **能力上已支持**（菜单/调试命令按 caption 匹配、模式标记中英同集、对话框按钮控件 ID 优先、运行时错误框英文措辞已识别），但**尚未做端到端实测**（验证机为中文版 zh-CN）。欢迎英文版用户报告「哪条工具可用/哪条失败（附回执与 IDE 界面语言）」——即使只是「能用/不能用」也极有价值。
