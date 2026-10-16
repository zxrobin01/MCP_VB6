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
- **已知局限**：见「七、已知限制」，尤其是**本项目在中文版 VB6（zh-CN）上开发与验证**这一点。
## 为什么做这个（项目背景）

VB6 在 2008 年就停止演进了，但它写的系统还在跑：工厂、财务、医院、政府内部系统里有大量
陈旧 VB6 工程仍需要维护，而维护者面对的是一套**没有任何现代工具链**的环境——没有语言
服务器、没有可靠的静态检查、编译报错只能手动点、改代码靠自己不手抖。

AI 编程助手已经能很好地处理现代语言，却**够不到 VB6**：它们既看不见 IDE 里打开的工程，
也点不了那个「生成工程」对话框。本项目就是补这个缺口。

**做法**：不另造编辑器，而是**把 VB6 IDE 本身变成一个 MCP 服务器**——加载一个插件，IDE
就对外提供 110 个工具：读/搜/改代码、静态检查、锚点编辑与安全事务、编译、运行与单步、
弹窗处理、窗体与项目元数据、快照与回滚、打包。此外还有 10 个 **hub 面**工具（IDE 实例管理与
经验库）在 IDE 进程外提供，客户端合计可见 120 个。
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
├─ CodeTools.cls          110 个 MCP 工具的实现（完整 VBIDE 对象模型 + 文件系统 + 语法门禁/编辑安全/编译权威）
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
  并把插件面的 110 个工具转发给中间服务；hub 用环境变量 `VB6MCP_BRIDGE_URL` 找中间服务（源码默认 `http://127.0.0.1:8767`）。
- 插件只接受**命名管道**（每个 IDE 实例一条；同一条管道**不允许并发连接**）⇒ 客户端应在**同一条连接上复用请求 id**，
  或把调用串行化；管道忙时回执给 `pipeBusy`，而不是「没有 IDE」。
- **传输选择** `VB6MCP_TRANSPORT=auto|pipe|tcp`（默认 `auto`：优先管道，不可用时**可见地**回退到插件的 HTTP；
  `pipe` 为严格管道、缺管道是显式错误；`tcp` 只走插件 HTTP，用于回滚与过渡）。
- **自定义端口**：`--port N`、`VB6MCP_BRIDGE_PORT`（或同一台机器上的 `VB6MCP_PORT`）。
  同时开多个 VB6 实例时**仍然只有一个服务端口**；多实例靠 `_target` 寻址（见第三节）。
- 运行状态查询：IDE 菜单 **Add-Ins → VB6 MCP Status...**；或 `GET http://127.0.0.1:<端口>/health`
  （配了凭据时该请求也要带 token）；`python bridge/vb6mcp_bridge.py --print-config` 打印生效端口 / 传输 / `authEnabled`。
- 日志：中间服务的日志路径见 `--print-config` 的 `logPath`；插件自身日志仍是 `%TEMP%\vb6mcp.log`。

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

## 四、工具列表（插件面 110 + hub 面 10 = 120 个）

> 权威清单见 **`docs/TOOL_CATALOG.md`**（由运行中的 add-in 经 `tools/list` 导出，再并入 hub 自有工具），
> 离线基线为 `tools/schema_baseline.json`，可用 `tools/check_live_tools.py docs/tools_schema.json` 做离线交叉比对。
>
> 工具来自**两个面**：**插件面 110 个**（`docs/tools_schema.json`，住在 VB6 IDE 进程内）与
> **hub 面 10 个**（`bridge/hub_tools.json`，在 IDE 进程外）。下面按功能把两个面共 **120 个**工具分组，
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
| `get_component_code` | 从 IDE 内存读组件代码，可只要某个过程或符号 | `name`, `include_line_numbers`, `coords` |
| `read_source` | 读磁盘源码：编码识别 + 字节级 sha256 | `path` |
| `read_file` | 读任意文本文件，UTF-8/UTF-16/ANSI 自动识别 | `path`, `max_lines` |
| `find_in_component` | 在组件内查找文本：整词、大小写、模式匹配 | `component`, `target`, `start_line` |
| `get_selection` | 读 IDE 当前选中的组件、活动窗口与光标 | - |
| `get_codepane` | 读代码窗格的选区、顶行与可见行数 | `component` |
| `set_codepane` | 设置代码窗格选区与滚动，并可置前 | `component`, `selection`, `top_line`, `activate` |
| `current_position` | 读断点停在哪：模块、过程、行号、源码行 | - |

### 4. 代码编辑与安全事务

| 工具 | 能干什么 | 关键参数 |
|------|----------|----------|
| `replace_code` | 行级改代码：替换、插入、追加、删除、整模块替换 | `name`, `operation`, `start_line` |
| `add_code` | 向组件追加代码，来自字符串或文件主体 | `component`, `code`, `file` |
| `create_event_proc` | 生成事件过程桩，如 Command1_Click | `component`, `object`, `event` |
| `write_file` | 写文本文件，可选编码、建目录与备份 | `path`, `content`, `encoding` |
| `write_source` | 写 VB 源码文件：显式编码 + 字节级 sha256 | `path`, `text`, `encoding` |
| `edit_by_anchor` | 按唯一内容锚点替换整行，不依赖行号 | `name`, `anchor`, `new_text` |
| `apply_edits` | 原子批量锚点编辑，任一处失败则整批不改 | `name`, `edits`, `expected_sha256` |
| `preview_edit` | 只回统一 diff 预览，不写任何内容 | `name`, `edits` |
| `replace_in_files` | 跨组件字面/正则替换，默认只计划不写 | `path`, `query`, `replacement` |
| `rename_symbol` | 按词法重命名标识符，跳过字符串与注释 | `path`, `old_name`, `new_name` |

### 5. 语法门禁与工程校验

| 工具 | 能干什么 | 关键参数 |
|------|----------|----------|
| `validate_code` | 写前静态校验代码，回带稳定指纹的诊断 | `code`, `name`, `kind` |
| `check_component` | 用同一套规则校验 IDE 里的组件文本 | `name` |
| `verify` | 组件总判定：IDE 与磁盘一致 + 诊断计数 | `name`, `disk_path`, `baseline` |
| `verify_scope` | 批量核验多个组件的 IDE 与磁盘一致性 | `components`, `path` |
| `format_code` | 按块结构重排缩进，并证明只改前导空白 | `path`, `indent`, `apply` |
| `check_references` | 列出工程引用并标出文件缺失的引用 | `include_resolved` |

### 6. 依赖、纯度与公开 API

| 工具 | 能干什么 | 关键参数 |
|------|----------|----------|
| `dependency_graph` | 组件与过程级依赖图：边、环、纯度分类；主截止是**确定性字符配额**（`VB6MCP_CHAR_QUOTA`，默认 900000），墙钟只作最后手段（`VB6MCP_BUDGET_MS`，默认 30000）。回执带 `partial` / `*StageCompleted` / `budgetExhausted` / `timingTruncated` / `procEdgesFound` 等完整性字段 —— **默认档是确定性的部分结果** | - |
| `purity_check` | 判定组件能否进离线测试宿主及原因 | `name` |
| `export_public_api` | 把组件公开 API 冻结成可比对规格；省略 `name` 时做有界整工程导出 | `name`, `membersOnly` |
| `check_public_api` | 与冻结规格比对，报移除、变更、新增成员；省略 `name` 时做有界整工程核验 | `baseline`, `name`, `membersOnly` |
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
| `ide_state` | 即时读 IDE 模式与在飞、取消、释放闸门状态 | `cancel_inflight`, `release_inflight`, `release_min_held_ms` |
| `desktop_probe` | 测量宿主进程桌面与可交互性，且不抢焦点 | - |
| `ide_quiet_mode` | 开或关安静模式：不抢前台，回执附提示 | `on`, `ttl_ms` |
| `todo_status` | 只读扫描仓库内的任务板与报告 | - |
| `audit_tail` | 读最近的工具调用审计条目 | `count` |
| `list_addins` | 列出已注册插件及其连接状态 | - |
| `set_addin_state` | 加载或卸载一个 IDE 插件 | `prog_id`, `connect` |
| `quit_ide` | 退出 IDE，响应后延迟执行，未保存会丢 | - |

### 17. hub 面：IDE 实例管理与经验库（10 个，无需 IDE 在线）

这 10 个工具由 MCP hub（`bridge/vb6mcp_hub.mjs`）提供，**在 IDE 进程之外**运行：客户端不必先加载插件，
VB6 也不必在跑。其中 5 个经验库工具只依赖本地 JSON 存储，**无需 IDE 在线即可用**。

- **经验库** `mcp_lessons/`：`defaults.json`（随包发布，勿改）+ `user.json`（用户层，按 id 覆盖/禁用/新增）。
  合并结果经 `initialize.instructions` 发布给客户端，agent 调工具前即可读到；`add/delete/import` 写 `user.json` 时带 `.bak` 备份。
- **IDE 实例管理**：`ide_list` 列出 hub 能寻址的全部 VB6 IDE：本机实例（含同一工程被两个实例打开的冲突），`ide_target` 记住默认目标，
  `ide_start` / `ide_stop` / `ide_restart` 只操作**本 hub 自己启动的**实例；同时列出**已配置的远端端点**（只探白名单、**绝不自动发现**；连不上也照样列出并带探测状态），并回显实例的**唯一身份**（endpoint / targetKey / pathVia）、真实路径与登记表拼写、以及**解析到同一身份的调用者别名**（junction / 别名写法）。

| 工具 | 能干什么 | 关键参数 |
|------|----------|----------|
| `ide_list` | 列出 hub 能寻址的**全部** VB6 IDE：本机实例 + **已配置的远端端点**（只探白名单、绝不自动发现；连不上也列出并带探测状态）；回报 pid/端口/工程、唯一身份（endpoint / targetKey / pathVia）、真实路径与登记表拼写，并回显解析到同一身份的**调用者别名**；含过期登记项与「同一身份两个实例」冲突 | - |
| `ide_target` | 设置或清除默认 IDE 目标：pid、工程名或 auto | `target` |
| `ide_start` | 按 .vbp 启动一个 VB6 IDE 实例 | `project` |
| `ide_stop` | 停止本 hub 启动的 IDE 实例 | `pid`, `force` |
| `ide_restart` | 重启卡死的自有 IDE 实例并证明已就绪 | `pid`, `project` |
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

**第一优先是本地化**：本项目在**中文版 VB6** 上开发与验证，已知耦合点可枚举（见「七、
已知限制」）。如果你在用英文版或其它语言版本，哪怕只报告「能用/不能用」也是极有价值
的贡献。
