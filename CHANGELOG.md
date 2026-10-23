# 变更记录

> **中文** | [English](CHANGELOG.en.md)

这里记录**用户能感知到的变更**：每次改了什么、**为什么改**。
原则只有一条：**只写能从代码或提交里核实的事**；无法确认的历史宁可不写。文中不出现任何外部未开源产品的名字、私有工程名或维护者本机路径。

**当前：产品 `0.3.0` · 契约 `v1.16.5` · 工具面 `120`（插件）+ `11`（hub）= `131`（客户端可见）** —— 产品版本来自 `Util.bas`（`initialize` 与 `mcp_info` 都会回显）；插件面住在 VB6 IDE 进程内，hub 面在进程外、**不开 IDE 也能用**；同一台机器上由**中间服务**（`bridge/vb6mcp_bridge.py`）持有**唯一一个 TCP 端口**（默认 `127.0.0.1:8767`，插件改走**命名管道**），MCP 客户端经 **hub**（`bridge/vb6mcp_hub.mjs`，stdio）接入，**配了凭据时 HTTP 面一律要 token**。契约版本是**工具面与回执的兼容性契约**（工具增删/描述变更会升它，`mcp_info` 的 `breaking[]`/`migrations[]` 列出会改变调用方式的变更）。本仓库**只发布源码**，`VB6MCP.dll` 需自行运行 `build.bat` 生成。另：`.vbp` 工程文件自身还有一个 `1.0.0` 的版本资源，与上面两个无关。

本文件只记录**源码 / 功能**方面的改动；打包、发布与其它基础设施调整不在此列。

---

## 2026-10-24

### 01:35 · 契约 v1.16.5 · 版本说明补正：instructions 不再写死版本与工具数 · README 补 v1.13~v1.16（工具面零变化）

**改了什么**：① `initialize.instructions` 的末条由硬编码 `v1.11.4 current (110 tools)` 改为动态拼接 `VB6MCP_CONTRACT_VERSION` 与 `Owner.ToolCount()`；同段标题由 `BREAKING (read before calling):` 改为 `HISTORICAL CONTRACT NOTES (oldest first; the authoritative breaking[] is in mcp_info):`（那 4 条是 v1.6~v1.11 时代的历史条款，不是当前 breaking）。② README 中英新增「hub 侧恢复原语与可靠性开关（v1.13 ~ v1.16）」章节（10 项），并把工具面口径修正为 120 + 11 = 131，补上 `ide_kill_and_relaunch` 与 `ide_start.ready_wait_ms` / `ide_restart.opId`·`wait_ms`。③ `Util.bas` 的契约常量 v1.16.4 → v1.16.5。④ README 中英新增「本地化与英文版 IDE 支持」小节，并把原「英文版可能失效 / 第一优先是本地化」的口径更正为「能力支持、未端到端验证」（只改文档）。工具面零变化：add-in 120 → 120、hub 11 → 11、DSH 面 131 → 131；无工具/参数/描述/schema 变化 ⇒ **冻结面不变、不重冻**。
**为什么**：客户端在 `initialize` 里看到 "v1.11.4 current (110 tools)" 与当前 v1.16.4 / 120 自相矛盾；把版本与数量写成动态值，就不会再随版本漂移而陈旧。README 缺的这批能力恰恰是"出故障时怎么自救"的说明，不写出来用户只能自己试错。

## 2026-10-23

### 20:59 · 契约 v1.16.4 · quit_ide 退出阶段写盘改成「声明」· 契约戳横幅校正（无工具变化）

**改了什么**：① `quit_ide` 在 `requested:true` 回执新增 `exitWrites{declaration,observed,condition,evidence,basis}`：它是**声明**不是测量（进程无法观测自身退出阶段）。`condition` 仅按当前可见工程数判定：0 个 ⇒ `no_project_observed_no_write`（实测未写）；≥2 个 ⇒ `project_group_context_then_clean_quit`（约 250 ms 后出现 50~53 B 的 `.vbg`）；恰好 1 个 ⇒ `undetermined`（两个实测形状相反）。`evidence` 为固定 4 项，`basis` 标注来源；`sideEffects.filesWritten` 只覆盖调用期间、**不覆盖退出阶段**。② 修正文档→生成物里的契约戳横幅（旧 v1.16.0 → v1.16.4）。工具面 add-in 120 / hub 11 / DSH 131 不变；仅 `quit_ide` 描述变更 ⇒ 冻结面显式重冻。
**为什么**：IDE 退出阶段可能写 `.vbg`，工具既不能确认也不能阻止；与其给一个"没写"的假保证，不如把能观察到的条件、证据和不确定写出来。

### 19:43 · 契约 v1.16.3 · last_error 的结构化错误通道（错误框还在时即可读）· 无工具变化

**改了什么**：① `last_error` 与 `ide_state.lastError` 新增只读结构化字段 `err{number,description,source,hresult,numberBasis}` / `errBasis` / `errChannel{...}` / `errCrossCheck` / `errLocation`；数据来自纯进程内内存读取（`GetModuleHandleW` + `CopyMemory` 读异常记录），**不调用** `EbExceptionHandled`（实测它会清错误状态）。② 本批**不抑制、不自动关闭、不无框读取**错误框；读不到的字段给 null + 原因，`errLocation` 统一 `basis:"not_attempted"`（不发明行列）。工具面 ADDITIVE：add-in 120→120、hub 11→11、DSH 面 131→131；仅 1 个 description 变更 ⇒ 冻结面显式重冻（`tools_schema.json` == `schema_baseline.json` = `D2A3EF0D…`）。
**为什么**：错误框的本地化文本以前是唯一来源，读不到就只能外部截图；结构化记录能在框还挂着时给出 number/description/source，同时明确划出"不做什么"的边界（不关框、不无框读、不猜行列）。

### 18:38 · 契约 v1.16.2 · dirty-untitled 关闭默认拒绝 + open_project 成功面归一 + 断点台账落点固定

**改了什么**：① `open_project{exclusive:true}` 与 `remove_project` 在"工程无盘上 `.vbp` 且未保存"时默认拒绝且**零写盘**（`dirty_untitled_would_write` + `zeroWriteBasis` 事后二次扫描自证）；组件级 dirty 必须读 `VBComponent.IsDirty`（`VBComponent.Saved` 对本机每个组件都抛 438）；opt-in 为两个工具各新增 `allowHostWrite`（默认 false）。② `open_project` 成功即 `ok:true`，非致命 IDE API 噪音只进 `apiWarnings[]`，不再占用 `apiError`。③ 断点注入台账默认写 `%TEMP%\vb6mcp\breakpoints.json`（可用 `VB6MCP_BP_REG_FILE` 覆盖），写入路径必须出现在 `sideEffects.filesWritten[]`。工具面 ADDITIVE：+2 可选参数，无工具增删、无 required 变化，2 个工具描述与 1 个参数描述变更 ⇒ 显式重冻。
**为什么**：VB6 关闭"脏且无标题"工程时会把未保存工程写进 IDE 当前目录并覆盖同名文件，而 `VBProject.Saved` 只读 ⇒ 只能默认拒绝、让调用方显式 opt-in；`AddFromFile` 抛错但工程实际已打开会让成功面说谎；断点台账落在工程 CWD 会污染工作区。

## 2026-10-22

### 23:36 · 契约 v1.16.1 · 静息 CPU 降 84% · TIMERPROC 宿主 tick · add_component 回执不再说谎（工具面零变化）

**改了什么**：① `add_component` 回执新增 `fileSyncSkipped` / `fileSyncSkippedDetail` / `failedAt` / `adoptedAfterError` / `adoptedVbError` / `adoptedNote`，新增失败原因 `add_component_failed`（旧 `add_component_vbp_failed` 仅保留在 component_file/vbp_sync 阶段），并修正 `vbError` 语义（旧值恒为 `"0:"`，现为真实 `Number:Description`）。工具面 120/11/131 不变、schema 未变（不重冻）。② 功能侧：静息 CPU 0.281% 总 / 0.231% 稳态（改造前 1.797% / 1.781%）；默认宿主 tick 由 VB6 Timer 控件改为 Win32 `TIMERPROC`（长调用期间 progress 帧 3–15）；bridge 对数组/标量 body 改为单文档 `{singleDocument,data,receipt}`（修 `find_gui` / `vba6_probe_plan` / `audit_tail`）；hub 面检测改为注册表文件轮询。
**为什么**：空闲轮询把宿主 CPU 拉到近 2%，长调用还会饿死进度帧；add_component 把"跳过文件同步"写成成功、并保留恒为 `"0:"` 的 vbError，会让调用方把失败读成成功。

### 14:01 · 契约 v1.16.0 · 断点行为验证（两阶段 probe）· RUN 态 End 去伪拒绝 · verify 编译器权威 · 多组件批量编辑

**改了什么**：① `breakpoints{action:"probe"}` 两阶段：`phase:"start"` 判设计态/无模态框/可断行后**启动程序**（不保存、不写源码）回 `probeId + startedPending`；`phase:"observe"` 用同一 `probeId` 等停、读 `position`+`callstack`、End 并用回读证明 `stateAfter=design`；`probeId` 单次、TTL 600000，`probeResult` 是闭合枚举，`phase:"auto"` 只是 start 别名。② `control_execution{mode:"end"|"stop"}` 在 RUN 态**不再假拒绝**（旧的 `commandbars_unavailable_while_running` 是伪拒绝），`programEnded:true` 只在 design 回读之后出现。③ `verify` / `verify_scope` 新增可选 `requireCompiler`（默认 false 逐字节不变；true 要么给编译器权威，要么结构化拒绝 `compiler_not_configured`，**绝不用 lint 顶替**）。④ `apply_edits` / `preview_edit` 多组件批量：`edits[].component`、`expected_shas[]`、40 组件 / 1000 ops 上限、全或无回滚，`preview_edit` 按组件分组 hunks 且**什么都不写**。工具数 120/11/131 不变；+7 可选参数、`breakpoints.action` 枚举 +probe、9 个描述变更、`apply_edits`/`preview_edit` 的 required 由 `["name","edits"]` 变为 `["edits"]`；本版已重冻。
**为什么**：native 断点"记录成功但从不命中"、injected 断点写盘并两次搞死 IDE —— 只有真跑一次、停在目标行，才能证明它有效；lint 是词法/结构，曾对大模块报 49 error 而同刻 `/make` 0 error，必须让编译器做权威；跨文件改名原来要十几次单组件调用、中途失败会留下"改了一半"的状态。

### 10:02 · 契约 v1.15.0 · run_project 的 observe 预算与完成标记 · get_locals 知情同意（插件面 +3 可选参数）

**改了什么**：① `run_project.observe_max_wait_ms`（默认 `timeout_ms`，夹取 1000..120000）到点**返回**：`state` 是封闭五值 `done-by-marker|exited|break|error|timeout`；`timeout` 是 `ok:true + timeout:true + retrySafe:true + suggestedAction:"observe_again"`（**不是失败**），`pipeReleasedByReturn:true` 不再长期持门。② `run_project.done_marker{path,contains}`：路径在 IDE 所在机器（可 UNC），只读前 65536 字节、不写；`contains` 大小写敏感，非 ASCII 同时尝试 UTF-8 与 ANSI；`done-by-marker` 仅当 `marker.matched:true`，请求了但未命中给 `marker.unmet:true`。③ break 态 observe 拒绝（`observe_in_break_not_supported`）；`#32770` 分类：只有运行时错误框算 `program_error`，IDE 自己的查找/替换框给 `observe_blocked_by_ide_dialog`。④ `get_locals.allowUnsafeInBreak`（默认 false）：break 态默认拒绝 `break_eval_requires_opt_in`；true 是知情同意、**不保证**不挂，恢复走 hub 的 `ide_kill_and_relaunch`，失败回执带 `gateReleased`+`howToRecover`。工具名/工具数不变（120/11/131）；本版冻结面未改，v1.16.0 关闭该滞后。
**为什么**：回放 7 秒结束却报 `Request timed out` 并把管道留在 busy；完成标记让"已经跑完"可判定；break 态读局部变量可能 wedged 单线程管道，必须让调用方显式同意并知道恢复路径。

### 10:02 · 契约 v1.14.0 · ide_restart 可续查 + content 单 JSON 文档 + peer_dead

**改了什么**：① `ide_restart` 新增 `wait_ms`（单次调用阻塞上限，默认 25000，夹取 1500..300000）与 `opId`（续查）；**超时不是失败**：回 `pending:true + state:"in_progress" + opId + currentStep + lastKnownState`，重启继续在后台进行；`ide_restart{opId}` 只查进程内操作注册表（保留 30 分钟）且**不启动任何东西**；同一时刻只允许一个重启，进行中再调回 `restart_in_progress + inProgress`。每条回执新增 `opId/started/state/finished/elapsedMs/currentStep/lastKnownState/opRule`。② `content[0].text` 改为**单个 JSON 文档**：数组/标量 body 变 `{singleDocument:true,data,receipt}`（空 = `data: []`），非 JSON body 给 `content_not_json`，对象 body 维持扁平；`VB6MCP_LEGACY_TEXTBODY=1` 只保留旧两段式一个版本。③ 传输失败且 peer 可证已死时给 `peer_dead` / `lostContext` / `howToRecover`。工具面仅 hub 的 `ide_restart` 变化（+2 可选参数），add-in 120 / hub 11 / DSH 131 不变。
**为什么**：重启实际完成、客户端却只收到无回执的 30 s timeout，调用方只能外部查 pid；盲目重试会二次重启。"JSON 数组 + 追加回执对象"的两段式让严格客户端 `JSON.parse` 失败，空结果与解析失败也无法区分。

### 10:02 · 契约 v1.13.0 · 目标解析与离线路由分离 + ide_start 就绪 + hub 带外复位（新增 1 个 hub 工具）

**改了什么**：① 默认目标阶梯**只从活实例选**：discovery 后复核 pid/transport 丢弃已死候选，hub 陈旧记录不再能成为默认目标；显式 `_target` 指向死 pid 时 `target_mismatch` 带 `onlinePids[]` / `targetDiagnostic` / `howToProceed` / `suggestion`；0 活实例时 IDE 面工具回 `no_live_ide`（`reasonLegacy:"no_ide_running"`），不再出现 pid 级 target_mismatch。② **离线路由**：`build` / `make_project(route!="ide")` / `project_meta` 在无活 IDE 时由 hub 自服务（`offline:true + offlineRoute:"cli"|"file" + targetResolved:false`），独立 `VB6.EXE /make`，编译后补 LAA 并回读、**所有改写之后**才取 `artifact{exists,bytes,mtime,sha256}`+`artifactChanged`（`hashedAfterRewrites`+`rewriteSteps[]`）；不能启动编译器回 `cli_unavailable`；`package_app`/`verify` 明确 `offlineServed:false + no_live_ide`。③ `ide_start`：`ok:true+started:true` 只声明**进程已创建**，`ready:true` 才表示观测到 add-in 对同一 pid 应答 `mcp_info`（有上限的轮询，非固定 sleep）；超时回 `ready:false + readyWaitMs + readyChecks{...} + pollingHint`；同 `.vbp` 已打开则复用。④ 新增 hub 工具 `ide_kill_and_relaunch`：不依赖插件管道，结束**本 hub 自有**实例、清注册、原工程重启并轮询就绪；回执 `stoppedPid/newPid/ready/readyWaitMs/threeWay{...}`；非自有 / 已知端口 <8767（含 8766）/ 远端实例一律拒绝。工具面：add-in 120→120，hub 10→11，DSH 面 130→131。
**为什么**：IDE 全关时 `build` 仍被陈旧 pid 挡住（离线编译是唯一能拿真编译器日志的通道）；`ide_start` 返回 pid 后立即调用却 `target_mismatch`（约 10~15 s 注册空窗）；管道 busy 时连 `release_inflight` 都送不进去，只能人工 Stop-Process —— 这三件事都需要 hub 侧有"只认活实例 + 带外复位"的能力。

## 2026-10-20

### 20:24 · 契约 v1.12.0 · 调试能力 + 写/验/构加固 · 插件面 110 → 120（新增 10 个工具）

**改了什么**：① **新增 10 个工具**：`debug_session` / `run_until_break` / `breakpoints` / `get_locals` / `eval_at_frame` / `run_project` / `last_error` / `move_procedure` / `delete_procedures` / `edit_by_byte_anchor`（插件面 **110 → 120**，hub 面不变 10，客户端可见 **130**）。② **Start/End 是排队命令（STA）**：一次调用覆盖不了 start→break ⇒ 两段式（start 回 `startedPending` + `howToContinue`，再 attach 观测）；`debug_session` 不保存工程（`preRunSave=false`）；排队中的 End 报 `endQueued` 而不是假 `wedged`（`wedged` 仅指调用内无法恢复）。③ 断点带 `paneLine` + `diskLine` + 来源声明；手工设的原生断点不可枚举/不可验证（`nativeStateVerified=false`）；条件断点用 `hitCount` 等价表达。④ `get_locals` 只给**名字**（立即窗上下文不是断点帧 ⇒ 值不可读，回 `localsUnavailable` + 原因）；`eval_at_frame` 只在 IDE 立即上下文求值，break/run 下默认拒绝、`allowUnsafeEval` 才执行，**对象/数组表达式求值前静态拒绝**（`object_expression_unsupported`，否则会挂死 IDE）；`timeoutMs` 只是调用方预算提示。⑤ `run_project` 两段式、Start 前设环境变量；Standard EXE 从 IDE 启动拒绝命令行参数；不能 attach 已运行的 `-Embedding` 进程；`last_error` 给错误对话框正文 + 调用栈 + 局部变量名字。⑥ `move_procedure` / `delete_procedures` **事务化**（快照 + 回滚）；`edit_by_byte_anchor` 按 **ANSI 原始字节**编辑、GBK lead/trail 边界 + roundtrip 校验，**失败绝不写盘**（`fileUntouched` + sha 相等）。⑦ `build` / `make_project` 新增 `sourceFreshness`（逐组件 disk-vs-pane，只比**规范化文本**）、`final{}`（LAA 补丁后采样 hash/PE/FileVersion/`verifiedOnDisk`）、`project`（仅 cli 路线）与 CLSID `InprocServer32` 守卫（前后读回/恢复）；**build 的 cache-hit 不再解析任何日志**（`logAvailable=false`、`logSuppressed`、`logOwnership.verified=false`，旧日志只报告不解析 ⇒ 修掉跨工程日志串台），并给 `compileLogHint` + `diagnoseRecommended`。⑧ `verify` 新增 `semanticDiff`（大小写 / Attribute / 尾空格归一 + 逐行 `paneLine/diskLine/category`）、`failureLayers{ideDiskDiff,lint,build,modal}`、`overallBasis`；`verify_scope` 只把**已证 IDE≠disk** 计 fail（`failBasis=ide_disk_diff_proven`）并支持 `baseline`，不变量 `pass+fail+unknown==checked` 恒发布。⑨ `add_component` / `remove_component` 的 `.vbp` 行与组件文件**原子同步**（组件文件先落盘，`.vbp` 恰好一个行尾 CRLF；失败**双侧回滚**），`Form` / `UserControl` 的新增显式拒绝（`pairedBinaryUnsupported`）。⑩ `write_file` / `write_source` 对「已存在非空文件 + 空内容」拒绝（`empty_content_would_truncate`，`confirm_truncate=true` 才允许），VB 源文件永不写空。⑪ `dependency_graph` 在 `partial` 时发布 `minQuotaHint{charQuota,budgetMs,basis}`（按建议重跑即可完整）与指向确切参数的 `completeHint`；不可达时报 `cannotComplete`，不编造。
**为什么**：这一版把「**调试**」从只能看状态推进到能驱动：Start/End 是 STA 上的**排队命令**，一次调用不可能同时覆盖 start 与 break，所以把它做成**两段式**并明确回报「已排队/下一步怎么继续」，比让调用方等一个不会返回的调用可靠得多。断点与局部变量这边，能**证实**的只有名字与位置（立即窗上下文不是断点帧），所以值不可读时明确回 `localsUnavailable`，而不是给一个看似可用的假值；对象/数组表达式会挂死 IDE，就在求值前静态拒绝。写类工具这一批统一了「**要么整笔成功、要么原样不动**」：程序块移动/批量删除带事务与回滚，字节锚点编辑在 GBK 边界与 roundtrip 校验通过前**绝不落盘**。构建侧把「**缓存命中**」和「**这次真的编译了**」分开：cache-hit 不再解析上一次的旧日志（那会把别的工程的错误当成这次的结果），并把耗时预算做成**强制**（到点杀子进程且事后无残留）。**限制（如实）**：断点帧局部变量的**值**仍不可读、对象/数组 inspect 不可达、任何立即窗求值都有挂死风险（可用 `release_inflight` 恢复）、`run_project` 的环境变量**程序内回读未证实**、UNC 路径未验证。

### 01:50 · 契约 v1.11.12 · 诊断一致性 + 写类行序 + 挂死可观测（无工具增删）

**改了什么**：① **归属判据统一**：`ide_list`、转发的 `ide_state` 回执、`ide_restart`（`ide_stop` 亦附）共用**唯一**解析器；回执新增 `ownershipVerdict{verdict,basis[],...}` + 规则原文，两轴并列：旧 `ownership` / `ownedByHub`（＝「本 hub 启动的」→ `ide_stop` / `ide_restart` 的**许可轴**，不变）与新 `ownershipVerdict`（**归属身份轴**）；**归属不授予许可**（用户手工启动但能应答的 IDE ⇒ owned，仍 `ownedByHub=false`、仍不可被 `ide_restart` 动）。判定 **owned / foreign / unknown**：foreign **只认正面证据**（注册 pid 已死但端口仍应答 / 监听端口的 OS owner ≠ 注册 pid / 保留区端口且应答 / 实例监听 pid ≠ 其注册记录 pid）；unknown 含 `portAvailable=false`（pipe-only 注册故意发布 port 0）；归属未知时 `ide_restart` 用 `refused_reason:"ownership_unknown"` + `legacyReason` + `howToDisambiguate`（默认拒绝重启）。`hubInstance.ownership` 不再为 null。② **写类行序**：`apply_edits` 新增 `deletedLines` / `deletedText`（截断）/ `deletedTextLines` / `deletedTextTruncated` / `deletedTextLimit`；`apply_edits` 与 `preview_edit` 用**同一 reason** 拒绝 `delete_with_text` / `would_empty_component`（preview 不报 `deleted*` 键）；`reload_component` / `save_project` 新增 `vbpOrder*` / `vbpRevision*` / `inMemoryOrder*` 一组字段 —— **保存后 `.vbp` 行序变了必须恢复**（只重排行、不丢 VB6 写入内容；恢复失败 `ok=false` + `reason:"vbp_order_restore_failed"`）；锁定上限：`vbpOrderLinesCapped` = 400 行、`vbpOrderChanges[]` ≤ 50、`deletedTextLimit` = 4000 字符。③ **挂死可观测**：`make_project` 新增 `timedOut` / `budgetEnforced` / `budgetMs` / `budgetSource`（`timeout_ms` / `default` / `clamped_*` / `not_applicable`）/ `stopReason`（四值）/ `stopReasonPriority` / `host_modal_blocking` / `host_modal{...}` / `isolatedMake{...}` / `budgetNote` / `timeoutNote`；`build` 同机制（超时 `ok:false` + `verdict:"timeout"`，不再抛「build process failed to run」）。**`ide_state` 零工具面变化**（无参数/描述/schema 改动），只改行为：隔离编译的等待循环泵消息（`Util.PumpNow`），使 `depth>0` 的**只读探测在长调用期间被派发**（`serviceable` 按实测 = 该只读调用确实被服务）。
**为什么**：这三批修的都是「**同一个事实在不同入口给出相反答案**」与「**卡住时说不清卡在哪**」。归属判据分散在端口 / 注册表 / 来源三处时，同一个 pid 可能在一处是「我们的」、另一处是「外来的」，于是唯一为「卡死恢复」设计的工具反而在最需要它的时候拒绝自己人；统一成一个解析器、两轴分开（身份 ≠ 许可）、未知就明确拒绝并给消歧步骤，才让它可预测。`.vbp` 行序在保存/重载后若被重排，工程能编译但组件顺序悄悄变了；把「保存后恢复原序」写成契约（含恢复失败即 `ok=false`）比事后排查便宜。挂死类问题最难的是**区分**「超时被杀」「缓存命中」「宿主模态框挡住」「预检拒绝」——把这些做成互不覆盖的字段（并且 cache-hit 不冒充编译），读回执的人才能一眼判断该重跑还是该改预算。

## 2026-10-17

### 18:36 · 契约 v1.11.11 · 版本说明补正 · `vba6_probe` 参数描述 hygiene

**改了什么**：① **版本说明补正**：`mcp_info.changelog` 里 v1.11.10 的条目此前写成「wording only, no behaviour change」，**与事实不符** —— 本版如实重述为**行为修复**（三个 winsock 传输缺陷，见上一条），并把同批的**描述级改动逐项**列出：**3 个工具描述**（`edit_by_anchor` / `apply_edits` / `preview_edit`：去内部编号前缀）、**1 个工具描述**（`verify_scope`：build 证据按项目判定）、**2 个参数描述**（`replace_code.expected_sha256` / `make_project.wait_for_build`）、**1 处运行期回执文本**（`make_project.buildStateNote` 去内部编号前缀，字段与语义不变）。② **参数描述 hygiene**：`vba6_probe` 的 **10 个参数描述**（`exec_probe` / `exec_set_frame` / `frame_probe` / `frame_set` / `frame_set_a1` / `frame_set_a2` / `worker_call` / `worker_p2_engine` / `mem_dump` / `mem_scan`）去掉内部实验编号与逆向轮次号，**语义不变**；**无参数增删、无类型 / required 变化**；发布树注释里的同类编号一并清除。导出审计 regex **不扩**（这是一次性 hygiene，不是审计规则）。**面变化构成**：插件 110 → 110、hub 10 → 10、客户端可见 120 → 120，ADDED / REMOVED = 0；生成件按新面重冻。
**为什么**：changelog 的唯一用途是让人判断「**该不该重新验证**」。把一个**行为修复**写成「只是措辞」，会让使用者在升级后跳过复验 —— 那他遇到的断连与截断会被当成「新版的新问题」，而实际上那正是被修掉的东西。所以补正比新增功能更优先：**说错比不说更糟**。参数描述里的实验编号同理：对调用方无用，却会把内部实验过程带进公开面；只改措辞、不动语义与参数表，所以除这次补正外不需要额外处理。

### 17:50 · 契约 v1.11.10 · 传输修复（行为修复）· build 证据按【项目】判定 · 一次性 hygiene

**改了什么**：① `verify_scope` 的 **build 证据按【项目】判定**：`buildOk` / `buildSource` / `buildNote` 来自**项目级** build 缓存（key 与 hash 都是项目级），**单个组件不可能「仅自己缺 build 证据」** ⇒ `unknown` 是**项目级**结论，`pass` 只在**有新鲜 build 缓存的工程**上可达（从未构建过的工程里，任何 IDE 与磁盘一致的组件都是 `unknown`）。② **传输修复（这是行为修复，不是措辞）**：修掉三个 winsock 传输缺陷 —— (a) **连接对象未被持久化** ⇒ 一条请求跨 TCP 段到达时**丢掉前半段**，要等 30 s 空闲清扫才断开；(b) **close-after-send 的判定顺序**放在队列已经清空之后 ⇒ `Connection: close` 的响应会滞留到空闲清扫才被发出；(c) **flush 的累积 `Mid$` 起点越过 64 KiB**，且 `send()=-1` 伴随 `WSAGetLastError()=0` 被误判为致命错误 ⇒ 正文在 64/96 KiB 处被截断。三者合起来，修掉了**间歇性的 `RemoteDisconnected` 与大于 64 KiB 的响应截断**。诊断侧的失败回执同时给出结构化 `reason`（`transport_idle_closed` / `remote_disconnected` / `truncated_response` / `transport_timeout` / `connection_reset` / `connection_refused` / `http_error`），并对**一次**传输失败自动重试一次。③ 一次性 hygiene：若干族的内部编号从注释、散文、工具/参数描述与运行期回执文本中清除；审计 regex **不扩**（这些族与错误码、协议头、版本号不可区分），发布树里只保留协议头名这类合法标识。**面变化构成**：插件 110 → 110、hub 10 → 10、客户端可见 120 → 120，ADDED / REMOVED = 0；**description 变更 4 处**（`edit_by_anchor` / `apply_edits` / `preview_edit` / `verify_scope`）+ **仅参数描述文本 2 处**（`replace_code.expected_sha256` / `make_project.wait_for_build`）+ **运行期回执文本 1 处**（`make_project.buildStateNote` 去掉内部编号前缀，字段与语义不变）。
**为什么**：① build 证据是**项目级**的事实（一次构建产出一个产物），把它当组件级用会读出一个不可能的结论 ——「这个组件自己没有 build 证据」。写清它是项目级，`unknown` 才不会被误当成这个组件的问题。② 那两个症状（偶发断连、大响应被截断）看起来像「服务偶发不稳定」，实际是**三个确定性缺陷**叠在一起：请求头在分段到达时被丢、连接关闭顺序反了、`flush` 用一个会累积的偏移。它们的共同后果是**同一次调用时好时坏**，最难排查；只有把根因写成清单，才谈得上「已修」。③ 内部编号与错误码、协议头在字面上不可区分，机械扩展 regex 会误伤（把合法的协议头名判成泄漏），所以这一族只做一次性清理，并以「发布树里还剩什么」作为验收。

### 16:24 · 契约 v1.11.9 · 声明区的真实语义 · 每个范围形状都带 mode · UNKNOWN 不再算 fail

**改了什么**：① `add_code.inDeclarationArea` 改为**真正的声明区**判定：`insertedAt.ide` 小于**第一个过程头行**才算声明区，因此**末尾追加 = `false`**（键名不变，**值语义变了**）。② `requestedRange` **每个形状都带 `mode`**：`mode:"range"`（`replace_code` / `add_code` / 守卫拒绝 = `{operation,start_line,end_line}`）、`mode:"anchor"`（`edit_by_anchor`）、`mode:"anchor-batch"`（`apply_edits`）。③ `verify_scope`：**UNKNOWN 不再计入 `fail`**，`unknown` 单列且与 `failures[]` **互斥**，不变式 **`pass + fail + unknown == checked`** 恒成立，回执新增 `countsConsistent`。④ 纯 hygiene（**无面变化**）：已发布树里的**内部任务号 token 与裸短号**全部移除（后者只出现在注释里）；导出审计的 regex 由「仅首族前缀」扩为**覆盖全部内部任务号族**（非数字后缀同样命中），并新增**显式可见的夹具豁免** `-InternalIdFixtures`（默认空 ⇒ 发布路径零豁免）；裸短号**不纳入** regex（与错误码、版本号不可区分）。面变化构成：4 个 `description` 变更，插件 110 → 110、hub 10 → 10、客户端可见 120 → 120，ADDED / REMOVED = 0。
**为什么**：① `inDeclarationArea` 原来的口径是「不落在某个未闭合的过程体里」，于是在**模块末尾追加**也算 `true` —— 那不是声明区；改成按「第一个过程头行」判定后，这个字段才真的能回答「我是不是写进了声明区」。② 三个范围形状原来只有两种带 `mode`，调用方要靠「是哪个工具回的」去猜；每个形状自带 `mode` 才能被同一段代码解析。③ UNKNOWN 的语义是「**没有足够证据**」（典型情形：工程从未构建过，没有 build 证据），把它算进 `fail` 等于把「没测到」判成「不一致」；单列并保持 `pass + fail + unknown == checked`，让调用方能自己核对这三个数加起来对不对。④ 导出审计原来只认一族前缀，其余族的内部编号照样能进公开树；扩成全族、同时给出默认空的夹具豁免，才既守住规则又不妨碍测试。

### 14:16 · 契约 v1.11.8 · 内部任务标识从工具描述、回执文本与源码注释中全部移除（重建窗口）

**改了什么**：这是一次**重建窗口**：把工具描述、参数描述、运行期回执文本与发布源码注释里的**内部任务标识**全部移除，**行为不变，只是措辞**。面变化构成（实测）：插件 110 → 110、hub 10 → 10、客户端可见 120 → 120，**ADDED / REMOVED = 0**；**9 处工具描述**（`add_code` / `apply_edits` / `check_component` / `edit_by_anchor` / `ide_state` / `make_project` / `preview_edit` / `replace_code` / `validate_code`）+ **3 处参数描述**（`export_public_api.name` / `ide_state.cancel_inflight` / `make_project.wait_for_build`），**无参数增删、无类型或 required 变化**。生成件（`docs/tools_schema.json` / `tools/schema_baseline.json` / `bridge/hub_schema.generated.json` 与工具目录）按清理后的描述重新生成并**显式重冻**。
**为什么**：内部工作编号对使用者**没有任何用处**：它既不能帮人判断该不该升级，也无法在阅读回执时提供线索，只会把维护者的内部流程顺带带进公开面。清理它们的风险主要在「**顺手改到语义**」，所以这一版被单独列为一个重建窗口：只动措辞、不动字段，并把面变化逐项列出来，让基线比对能把这批差异解释成预期内的文本变化，而不是一个说不清的面变更。

### 13:45 · 契约 v1.11.7 · 写类回执与行号空间 · 全工具面未知参数门 · 隐藏 Attribute 行

**改了什么**：① `add_code` 回执新增 `insertedAt{ide,disk}` 与 `inDeclarationArea` —— **落点由写后回读测量**，IDE 行与磁盘行两个坐标同时给出。② **未知参数门覆盖整个工具面**：以 `tools/list` 声明的参数集为合同，未声明的键一律 `ok:false` + `reason:"unknown_argument"` + `unknownArguments[]` + `acceptedArguments[]`（`_target` 除外）；`reason` 同时细分 `unknown_argument` / `component_not_found` / `target_has_no_project` / `ide_not_ready`（旧的 `no_project` / `ide_model_not_ready` 保留在 `reasonLegacy`）。③ `verify` 对「**只有隐藏 Attribute 行**」的面板/文件差异给出结构化标注 `attributeOnlyDiff` / `attributeLineCount`，与真实源码差异**分开计**；「经 VBIDE `SaveAs` 保存会把隐藏 Attribute 行物化到磁盘」这一点写进了工具描述。④ `get_component_code{proc|symbol}` 明确范围字段的行号空间（**IDE 行**）并新增 `startLineDisk` / `endLineDisk`。⑤ `check_public_api.membersOnly` 写明作用与局限（文本级轻解析、参数名丢弃、**比对永不跳过**）。⑥ **返回信封写入契约**：`result.content[0].text` 是一个 JSON 字符串。⑦ `check_public_api` 接受「按 JSON 字符串字段回灌」的规格：入口把字面 `\n` / `\t` / `\r` 解码回 TAB/LF/CR，并回 `baselineNormalized:true`。⑧ `dependency_graph` 的完整性变成**可判定**：新增 `edgesComplete` + `componentsUnscanned` / `edgeComponentsUnscanned` + `quotaSource`，并允许用 `char_quota` / `budget_ms` 为**单次调用**要完整图（每次调用独立，不继承上一次的配额）。⑨ `verify` / `verify_scope` / `get_component_code` 共用**同一个** `headerOffset`（含隐藏内联 Attribute 行的映射），`verify` 的比较跳过隐藏 Attribute 行 —— 两个入口不再对同一份文本给出相反结论。**面变化构成**：插件 110 → 110、hub 10 → 10、客户端可见 120 → 120，无工具增删；`inputSchema` 恰新增 **6 个参数**（`preview_edit.match` / `set_codepane.view` / `vba6_probe.step` / `vba6_probe.frame_index`，后三者描述里注明「被忽略 / 仅兼容保留」；`dependency_graph.char_quota` / `budget_ms`），**无 required 变化、无既有参数类型变化**；工具描述变更 2 处；其余为回执字段。冻结面随之重冻。
**为什么**：① 写类操作最怕「说了写在哪、其实在别处」；`insertedAt` 由**写后回读**量出来而不是推算，IDE 行与磁盘行分开给，是因为两者之间还隔着隐藏 Attribute 行。② 未声明参数被静默忽略是最便宜也最贵的错：便宜在于调用方看不出来，贵在于他以为自己配置过的行为从没生效；把门开到整个工具面，才不会有「这个工具查、那个工具不查」的盲区。③ 面板与磁盘的差异混在一起时，「只是多了几行隐藏 Attribute」会被报成源码不一致，反过来也让人不再相信 `verify`；把这类差异**单列**并计数，真实差异才有意义。④ `get_component_code` 与 `verify` 各自算一套头部偏移，就会出现**同一份文本两个入口结论相反**——这是一个纯内部不一致，对使用者却是「工具在骗我」；共用同一个偏移并统一跳过隐藏行，才消掉这种自相矛盾。⑤ `dependency_graph` 的截断原来没法从回执里判定，调用方只能猜这次的图是否完整；把 `edgesComplete` 与未扫描组件数写出来，并允许单次调用自己指定配额，完整性才成为一个可核对的事实。

### 13:17 · 契约 v1.11.6 · 诚实失败面 · 未知参数门 · 写入与分页回执补齐

**改了什么**：① `check_public_api` 只要给了 `name` 就**永远比对**（`membersOnly` 现在**只影响输出大小**，不再顺带关掉比对）；给一个**不是规格文本**的 `baseline` ⇒ `ok:false` + `reason:"baseline_unusable"`（不再回「没有差异」）。两个入口（给了 `name` 的组件级、不给 `name` 的整工程级）回**同一组字段** `comparedAgainstBaseline` / `baselineLines`，不可解析的 baseline 走**同一个拒绝形状**（`scope:"component"|"project"`）。② `check_component` 新增**可选参数** `baseline`，真正声明并生效（无法解析 ⇒ 同样 `baseline_unusable`）。③ `export_public_api` / `check_public_api` / `validate_code` / `check_component` 对**未声明的参数名**一律 `ok:false` + `reason:"unknown_argument"` + `unknownArguments[]` + `acceptedArguments[]`（`_target` 是传输上下文，恒接受）。④ `write_file` / `write_source` 回执新增 `createdDirs[]` / `createdDirsCount` / `exists` / `size`；父目录建不出来 ⇒ `reason:"parent_dir_missing"`（`missingPath` + `reasonLegacy:"parent_not_creatable"` + `reasonAliases[]`），不再是一句笼统失败。⑤ `find_in_component` 回执新增 `hits[]` / `count` / `nextStartLine`；分页走到末尾 ⇒ `found:false` + `reason:"no_more_hits"`（**不回绕**：连续取页不会突然又交回前面的行）。⑥ `make_project{diagnose}` 的**隔离副本**现在纳入二进制兼容引用（`CompatibleEXE32` / `CompatibleEXE`）与 PropertyPage 的 `.pgx`，回执新增 `compatibilityReferenceCopied` / `compatibilityReference`；缺失的引用写进 `missingFiles`，不再静默重铸 class ID。**面变化构成**：插件 110 → 110、hub 10 → 10、客户端可见 120 → 120，**无工具增删**；`inputSchema` 恰新增 **1 个可选参数**（`check_component.baseline`），另有 **4 处工具描述**；其余全部是**回执字段**（不影响 `inputSchema`）。冻结面随之显式重冻。
**为什么**：这一批针对的都是「**把没测到读成没问题**」的形状。一个解析不了的 baseline 原来会让比对**静默跳过**并回「没有差异」，`membersOnly` 原来还能顺带关掉比对 —— 两者都会给调用方一个绿色的假结论；改成「永远比对、比不了就明确拒绝」，`ok:false` 才是唯一能表达「这次没比成」的结果。未声明的参数名原来被**静默忽略**（调用方以为某个开关生效了，其实根本没传进去），现在逐个列进 `unknownArguments[]` 并附 `acceptedArguments[]`，拼错的参数名当场可见。写文件原来只报「写了多少字节」，父目录不存在时那次写入**根本没发生**却仍回一个自洽的成功形状；现在回读 `exists` / `size`，建不出父目录就明确说 `parent_dir_missing`。分页原来走到末尾会**回绕**，连续取页会把同一批行再交一遍；`no_more_hits` 让「到了末尾」成为可判定的终点。隔离副本漏带兼容引用时会**静默重铸 class ID**，编出来的产物与真实工程不是一回事；把漏掉的文件写进 `missingFiles`，诊断才谈得上可信。

### 11:02 · 契约 v1.11.5 · 撤回「在飞期间只读调用仍被服务」· 行数口径统一 · hub 归属判据统一

**改了什么**：① **撤回一条旧口径**：本代之前的文档/回执曾称「写飞行期间只读调用仍被服务」—— **现在不再成立**。插件是**单线程**：一次调用还在飞行时，**任何第二次调用都不会被执行（包括只读）**，被服务的是 **busy 拒绝回执**；`readonly_serviceable` 现**恒为 `false`**（键保留，旧客户端仍能读到）；`cancel_inflight` 与 `release_inflight` 降为 **尽力而为（best effort）**——陷在**宿主模态循环**里的调用**永不中止**（消息泵被 VB6 自己占着）。**唯一可靠出路 = 等它结束，或从外部终止 VB6 进程**。② **行数口径统一**：`newTotalLines` 改为**调用方口径**（与 `get_component_code` 的 `splitlines()` 一致，**末尾空行不计**），四条写路径（`replace_code` / `add_code` / `ChangeResult2` / `BreakpointEdit`）共用**同一个计数函数**；新增回执字段 **`trailingBlankLine`**；`syncWarning` 改为与**面板真值**比较（此前两个内部计数互比 ⇒ 永远相等 ⇒ 永不触发）。③ **hub 归属判据统一**：`ide_stop` / `ide_restart` 共用**一个所有权解析器**（同一个 pid 不可能得到相反结论），所有拒绝路径携带 `alsoForeignPort` / `ownershipBasis` / `portRuleApplies` / `portBasis`；`ide_list` 自述 `registryWrites:0` 与 `registryPolicy`。**面变化构成**：工具面**无增删**（插件 110 / hub 10 / 客户端可见 120），`inputSchema` **无参数增删、无类型变更**；本版自身 **2 处描述文本**（`ide_state` 的描述，及其 `cancel_inflight` 参数的描述）+ 若干**新增回执字段**（`trailingBlankLine` 等，不影响 `inputSchema`）。**同日并入同一广播的增量**：`command_bars{action:list}` 每条控件新增 `executable` 与 `enabledReason`（口径 = `action:execute` **实际会不会执行**——按同一枚举顺序取**第一个**该 caption 的控件看它是否 enabled；caption 会跨 bar/工具栏重复，单看 `enabled` **推断不出**可执行性）、`replace_code` / `add_code` 的描述写明**尾随换行不是 op 的属性**（结果尾部取决于**基底是否以换行结尾**与**插入位置**：中间插入继承基底、末尾插入重建尾部；**以回执 `trailingBlankLine` 配合 `newTotalLines` 为准**）⇒ 面变化共 **+3 处描述文本**（合计本批 **5 处**），GENERATED 三件套随本次广播重冻。
**为什么**：① 旧口径是**判断错误**：插件单线程，第二次调用根本不会被执行，能拿到的只有 busy 拒绝——继续宣称「读是安全的」会让调用方反复试探、只多收几次拒绝，更糟的是在真需要自救时误以为读能通而错过唯一有效出路；把 `cancel_inflight` / `release_inflight` 标成尽力而为，是因为**陷在宿主模态循环里的调用无法被打断**，把它们当可靠自救通道会误导。② 四个写路径各算各的 `newTotalLines`，「改了多少行」随路径不同而不同；`splitlines()` 口径（末尾空行不计）让调用方用**同一个公式**核对。`syncWarning` 原先两个内部计数互比 ⇒ **永不触发**，改成与面板真值比较后才可能报出来。③ hub 的归属判据原先分散在端口/注册表/来源三处，同一个 pid 能同时得到「外来实例」与「自己的实例」两种结论——这会让唯一为「卡在宿主模态」设计的恢复工具，在最需要它的时候被自己的护栏拒绝。

### 01:57 · 契约补记 · 两条工具描述由乱码修正并重冻（工具集零变化）

**改了什么**：两条工具描述此前是**乱码**，现已修正为正确中文并**显式重冻**：`command_bars` 的 `bar` 参数说明（旧 `璋冭瘯` → 新 `调试`）与 `vba6_probe` 的 `worker_p2_engine` 参数说明（旧 `搂5` → 新 `§5`）。工具面**零变化**：插件 110 → 110、hub 10 → 10、客户端可见 120 → 120，**ADDED / REMOVED 均无**；110 个工具逐字段全量 diff **恰好 2 处**，且方向已证——**旧串 == 新串的 UTF-8 字节按 CP936 读**（纯乱码，无语义变化）。冻结面（`docs/tools_schema.json` / `tools/schema_baseline.json` / `bridge/hub_schema.generated.json` / `docs/TOOL_CATALOG.md`）随之**显式重冻**，门禁 `schema_changes=0`、`hub_alignment_failures=0`、`live_tool_changes=0`。**未升契约版本**：仅 2 处 description 且方向已证为「还原意图文本」，无工具/参数增删 ⇒ 不 bump；将来出现**真正的**面语义变化（增删工具 / 参数 / 枚举值）**必须** bump。
**为什么**：源文件当时以 UTF-8 存储，而 VB6 按系统 ANSI 代码页读取 ⇒ 中文变成乱码。修正后描述语义与写作者原意一致；但**描述属于冻结面**，必须显式重冻并留证，否则基线比对会把它当成真实的面变更。不升版本，是因为它不是行为或接口变化，只是把同一段文本**还原成它本来该有的样子**。

## 2026-10-16

### 22:31 · 契约补记 · 加法回执字段 + 「缺席也是一种断言」

**改了什么**：**不升契约版本、不动冻结件**（`tools/list` 的 name / description / inputSchema **逐字节不变**；插件 110 / hub 10 / 客户端可见 120）。`build` / `make_project` / `diagnose` 等新增并明确若干**回执字段**：`matchRule` / `firstErrorKind` / `rawLen` / `rawSha256` / `logEncoding` / `logBytes` / `rawLine` / `logExcerpt`；`firstError` 取值有**三态**（`file-line` / `project-level` / `log-unparsed`），**未能解析出首错时 `firstError` 整体缺席**——不是 `{file:"",line:0,lineSpace:"disk",text:""}` 这样的空壳。同一条补记还含三项（细节见契约表对应版本行）：`quit_ide` 的三段契约（回执含 `requested` / `pendingExit` / `exited` / `blockedBy` / `evidence` / `howToConfirm` / `waitedMs`，其中 **`exited` 恒为 `false`**——进程无法观测自身退出，这是契约不是缺陷）、`add_code` 的 string 分支改为**末尾追加**（回执新增 `ok` / `okReason` / `sha256Before` / `sha256After` / `lineDelta` / `appliedRange` / `landedLine` / `landedVerified`；**需要按位置插入请改用 `edit_by_anchor` 或 `replace_code`**）、以及插件启用门的环境契约（**决定插件是否被实例化的是 `%WINDIR%\VBADDIN.INI` 的 `[Add-Ins32]` 段**，不是注册表 `LoadBehavior`：该值为 `0` 表示**曾被一次真实加载失败禁用、此后永不重试**，`1`/`3` 为启用、`3` 为稳态，**缺失条目**则退回注册表、可加载）。
**为什么**：「字段缺席」本身就是**断言**。拿不到编译器原话时，回执必须**缺**这个字段，而不是给一个填零的空壳——否则调用方会把「没测到」读成「测到的是没有错误」，这是最危险的一类假绿。同理 `rawLine` / `logExcerpt` 能读到就给原文、读不到就不写；三态 `firstErrorKind` 让「文件行号级错误」「工程级错误」「日志没解析出来」不再混成一个。`quit_ide` 把 `exited` 恒 false 写进契约，是因为**进程无法观测自身退出**——与其返回一个猜的值，不如明确说「这一项不由我判定」。`add_code` 改成末尾追加，是为了让 `requestedRange == appliedRange`、落点可由**写后回读**验证（`landedVerified` / `landedLine`）；需要精确位置时用锚点编辑工具。
### 11:37 · 换代 · 中间服务成为唯一 TCP 监听者 · 插件改走命名管道 · HTTP 面一律要 token

**改了什么**：客户端**不再直连插件**。本机由**中间服务**（`bridge/vb6mcp_bridge.py`）持有**唯一一个 TCP 端口**（默认 `127.0.0.1:8767`，可用 `--port` / `VB6MCP_BRIDGE_PORT` 改），插件与它之间走**命名管道**（每个 IDE 实例一条管道；同一条管道**不允许并发连接**）；MCP 客户端经 **hub**（`bridge/vb6mcp_hub.mjs`，零依赖，stdio）接入 —— hub 提供 `ide_*` 与经验库这 10 个工具，并把插件面的 110 个工具转发给中间服务。**插件自身的 HTTP 监听不再是默认路径**（`tcp` 轨道只用于回滚与过渡）。**鉴权**：一旦配置了凭据，**所有 HTTP 面**（`/rpc`、`/mcp`、`/health` 及其余）都要 bearer token，用 `Authorization: Bearer <token>` 或 `X-VB6MCP-Token` 发送；凭据来源 `VB6MCP_TOKEN`（优先，不落盘）或 `VB6MCP_TOKEN_FILE`（仅在前者未设时读取）；**未配置凭据时服务照常运行**，但 `/health` 报 `auth.enabled=false`、启动日志**同时告警** —— 未鉴权的监听面**不允许看不见**。**URL 里带凭据一律拒绝**（401 `token_in_url_rejected`）。**绑定只允许回环**：非回环地址被拒绝并以**退出码 6** 结束，**没有覆盖开关**。**传输选择** `VB6MCP_TRANSPORT=auto|pipe|tcp`（默认 `auto`：优先管道、不可用时**可见地**回落到插件的 HTTP；`pipe` 为严格管道，缺管道是显式错误；`tcp` 只走插件 HTTP）。**寻址**：`_target = [endpoint:]projectPath`（未限定的调用**不会被送到远端**）；单实例管道忙时回执给 `pipeBusy`，而不是「没有 IDE」。**客户端配置**：hub 需 `VB6MCP_BRIDGE_URL`（源码默认 `http://127.0.0.1:8767`；服务实际端口不同时必须显式给出）与上面两个 token 环境变量之一。可用 `python bridge/vb6mcp_bridge.py --print-config` 查看生效的端口、传输与 `authEnabled`。
**为什么**：旧代由插件自己监听 HTTP 并被直连 —— 本机回环口对任何本地用户都是可驱动 IDE 的入口，而且「多实例」意味着多个端口。新代把对外监听收敛为**一个**由中间服务持有的端口、让插件只接受**本机命名管道**，HTTP 面再统一加 token；并**从实现层禁止非回环绑定**（拒绝而非只警告），避免「以为只在本机、其实暴露在局域网」。`auto` 默认保证过渡期不打断仍按旧寻址工作的机器；严格 `pipe` 档让「以为走了管道、实际落到 HTTP」不再可能。管道单实例是**传输定义**而非缺陷：客户端应在**同一条连接上复用请求 id**，或把调用串行化。

## 2026-10-15

### 23:20 · 契约 v1.11.4 · 写前结构守卫 · 歧义锚点与多行预览 · 过程范围回显 · 构建状态可查询

**改了什么**：① **写前结构守卫** —— 写类工具在**写入前**检查新块首尾关键字配对：新块自带多余尾部 `End Sub/Function/Property` ⇒ `extra_procedure_end`（给出具体块行号、已配对项与 `suggestedText`）；`set`/`append`/`add_code` 的块不自足 ⇒ `missing_procedure_end`；结果级 `predicted_procedure_imbalance` **只在「改前配平、改后将失衡」时**触发（片段式合法编辑不受影响，已失衡的模块仍可被修复）。② **歧义锚点**：未命中或歧义时回候选 `hits[]{line,proc,text}` 与 `coords`/`lineSpace` 且**不写入**（用 `occurrence`/`scope` 指定）；`preview_edit` 补齐**多行块锚点**，与 `apply_edits` 走**同一条**解析路径（此前只替换首行 ⇒ 预览与写入分歧）。③ **过程范围回显**：`replace_code`/`apply_edits` 回执带 `procBefore`/`procAfter{proc,startLine,endLine,firstLine,lastLine}`；`replace_code` 的 replace/delete **保留** `expected_sha256` 硬前置。④ **构建状态可查询**：`make_project` 新增**可选**参数 `wait_for_build`（毫秒，默认 0 = 既有 15 秒，上限 120 秒）与 `waitRequestedMs`/`waitUsedMs`，并恒带单一状态 `buildState`（`blocked_by_host_modal|pending|completed|completed_without_artifact_change|failed|no_artifact`）；**隔离副本诊断已证编译失败时不再报 `pending`**，改为 `verdict=compile_failed` + `buildState=failed` + `pendingOverriddenByDiagnosis=true`。⑤ **表述层**：`validate_code`/`check_component` 明确 **linter ≠ 编译器**（只做词法 + 结构，**不解析类型与符号**），真编译结论用 `build`。工具面**无增删**（插件 110 / hub 10 / 客户端可见 120）：`inputSchema` 只有 `make_project` 多一个**可选**参数，另有 **7 个工具的描述变化** ⇒ 冻结面显式重冻。
**为什么**：① 带多余收尾的块原先要等到**编译期**才炸成 `E105`，真实原因与位置都更难读 —— 现在在写入前就拒绝并指出位置；② 锚点歧义时的旧行为是静默空操作或改错行，现在先给候选由调用方定夺，预览也不再与实际写入分歧；③ 行号漂移下按范围替换会打到别的过程 —— 回显改写前后的过程范围让「到底改了哪个过程」可核对（保留整面 sha 前置，是因为它能发现**任何**漂移）；④ 编译的「成功」原先只看 `compiled`，而「未产出」与「编译失败」分不开 —— 现在给了单一可查询状态，并修掉了本批验收自己抓到的缺陷：诊断已经证明编译失败、回执却还叫调用方「再等等、重读 output」；⑤ 调用方曾把 `validate_code` 当离线类型检查用 —— 描述必须说出真实能力边界。

### 22:13 · 契约 v1.11.3 · 依赖图：确定性配额 + 诚实截断

**改了什么**：`dependency_graph` 的主截止从**墙钟**改为**确定性字符配额**（`VB6MCP_CHAR_QUOTA`，默认 **900000**，按已处理字符数计），每个阶段各留**保留份额**（分词 / 组件边 / 纯度 / 环与过程边）⇒ 默认档**必定产出过程边**；墙钟降为**最后手段**（`VB6MCP_BUDGET_MS`，默认 30000）并用**独立字段** `timingTruncated` 标注 —— `budgetExhausted` 只表示配额到点，两者语义不同。**诚实化**：新增 `procEdgeStageCompleted`；`partial` 改为「**有阶段没做完**」；四个 `*StageCompleted` 在**该阶段或其输入阶段被截断**时一律 `false`；`componentsTotal` 改为**工程真实组件数**；`componentsScanned` 只计**完整分词过**的组件。**确定性口径**：同一输入重复调用，**答案部分逐字节相同**，允许变化的只有自计时 / 会话身份这类**信封字段**。工具面**无增删**、`inputSchema` **逐字节不变**，唯一变化是 `dependency_graph` 的**描述** ⇒ 冻结面显式重冻。
**为什么**：旧实现按墙钟决定截到哪 —— 同一份未改动源码连续三次调用，过程边数量会漂移，大工程甚至会在「零边」和「上千边」之间跳；更糟的是阶段被截断**仍报「已完成」**，默认档交付的是「有节点、零边」的图。改成确定性配额后同一输入必停在同一处，并且「没做完」会被如实标注。**已知限制**：超大工程的**完整图单次调用交付不了**（要完整图需同时提高字符配额与墙钟预算，分页能力见后续版本）。

### 16:34 · 契约 v1.11.2 · 源码印记（SOURCE STAMP）

**改了什么**：`mcp_info` 的 `build` 下新增两个**加法**字段：`sourceHash`（源码输入集合的确定性指纹）与 `sourceAlgo`（算法标识）。**工具面零变化**：`tools/list` 的名称 / 描述 / 参数逐字节不变，四个冻结面**不重冻**。语义：**相同代码 ⇒ 相同指纹**（跨次、跨机；二进制本身可因构建环境不同而不同），**不承诺反向**（源码改动若对编译器无效，产物理应相同）。算法标识由 `vb6src/1` 演进为 `vb6src/2`：对 `.vbp` 取指纹前**屏蔽四个版本键的值**（`MajorVer`/`MinorVer`/`RevisionVer`/`AutoIncrementVer`，键保留），使编译器把自增后的版本号写回 `.vbp` 时指纹**不再漂移**；产物版本改由检查工具**单独输出并比较**，版本不同**不影响**判定。
**为什么**：构建判定原先靠对产物做字符串刮取 —— 同一份源码在不同环境构建出的二进制可以不同，「源码到底变没变」因此无法复核。把**源码输入集合**的指纹固化成回执字段后，调用方可以用它做**源码身份**核对；把版本键从指纹里摘出去，是因为一次 `/make` 会**改写自己的输入**（版本号自增），旧算法会因此把「没改源码」误判成「源码变了」。
## 2026-10-14

### 23:30 · 契约 v1.11.1 · 有界化与引导修正 · 依赖图与公开 API 工具不再可能拖死宿主

**改了什么**：① **`dependency_graph` 与两个公开 API 工具（`export_public_api` / `check_public_api`）跑在同一道墙钟预算下**（默认 **8000 毫秒**，可用环境变量 `VB6MCP_BUDGET_MS` 覆盖）；预算到点即**停止扫描并返回已有结果**，回执恒定带 `budgetExhausted` / `budgetMs` / `elapsedMs` / `stage`（`dependency_graph` 另带 `tokenizedComponents` / `edgeScannedComponents`，两个 API 工具另带 `scanned` / `componentCount`）。② 两个 API 工具各新增**可选**参数 **`membersOnly`**（默认 `false`）：为真时只回成员表（`kind` / `name` / `sig`），**跳过规格文本 / 违规文本的生成以及随之而来的第二次解析**。③ 两个 API 工具的 **`name` 现在可以省略**：省略即对**整工程**做**有界**导出 / 比对（逐组件计数 + `scanned` / `componentCount` + `budgetExhausted`），不再是无边界扫描。④ `dependency_graph` 的组件边阶段由「**对每一对模块各做一次全文匹配**」改为「**一趟式成员判定**」——每个模块只分词一次，再用一张小写名表做字典式查表。⑤ 工具面 **插件 110 → 110**、hub **10**、客户端可见合计仍 **120**：**没有增删工具**，新参数为纯增量，默认行为不变。⑥ **非功能性措辞修正**：`validate_code` 的说明不再自称"本地规则引擎"，改为明确**它在插件内运行、需要活的 IDE、不是离线检查**。⑦ hub 侧的**转发失败分类**随本批落地：`noIde`（没有可对话的 IDE / 没有可用路由）、`wedged`（连接已建立但对方超时或断开，或所有已发现的实例探活都失败）、`toolError`（IDE 已经应答，而应答本身是结构化错误）—— 调用方不必再从原始传输错误里猜。
**为什么**：原先组件边阶段的开销随**模块数的平方**增长 —— 模块一多，每一次配对都对整段模块文本做一次匹配，即上万次全文扫描；在 **200 组件级别**的工程上实测**调用不返回**且持续占用 CPU，把宿主一起拖住，之后连只读的状态查询都要排队。改成一趟式判定后，同一工程约 **8.5–11.7 秒**返回，随后的只读状态查询约 **47 毫秒**、窗口保持响应。两个 API 工具在省略组件名时原本是**无边界**的整工程扫描，同样可能越过客户端的调用超时；给它们同一道预算、并让 `name` 成为可选项，是让"整工程导出"变成一个**有界、有进度可看**的能力，而不是一个可能挂死的调用。`membersOnly` 是纯增量：默认关，只给"只要成员表"的调用方省掉一次解析与一大段文本。措辞修正的动机来自实测：调用方曾把 `validate_code` 当成**离线校验**使用 —— 它的描述必须说出真实前提（需要活的 IDE），否则调用方的预期与事实分叉。

### 20:31 · 契约 v1.11.0 · 工具面 115 -> 110（插件面）· 5 个经验库工具移到 hub 面 · IDE 寻址键改为 `.vbp` 路径 · 远端端点显式白名单

**改了什么**：① **5 个经验库工具移到 hub 面**（`list_lessons` / `add_lesson` / `delete_lesson` / `export_lessons` / `import_lessons_md`）：工具名、描述与参数**一字未改**，调用方式不变；实现移到 Python 侧（`tools/lessons.py`），数据仍是 `mcp_lessons/defaults.json` + `mcp_lessons/user.json`（schema 不变，已有的 `user.json` 原样沿用）。因此**插件面 115 -> 110**、**hub 面 5 -> 10**，**客户端可见总数仍是 120**。② **IDE 寻址键从 pid 改为 `.vbp` 路径**：规范写法 `_target = "[endpoint:]<drive>:\<dir>\project.vbp"`（省略 `endpoint` 即本机）；裸 `pid`、裸工程名与 `auto` 仍然可用，但只作**兼容别名**、不再作为规范键回报。路径按「大小写不敏感、`\` 与 `/` 等价、去尾部分隔符」归一化，本机存在的路径还会做一次真实路径解析；同一个选择器匹配到多个身份时**结构化拒绝**（`target_conflict`）并列出候选，**绝不替调用方挑一个**。③ **不带 `_target` 的调用不再自动选到远端**：只有「本机没有任何活 IDE、而远端在线」时返回结构化的 `remote_requires_explicit_endpoint`（`routed:false`，并给出可直接复制的端点限定键）；本机有活 IDE 时照旧自动选中。④ **远端端点只能来自显式白名单**（环境变量 `VB6MCP_REMOTES` 或 `bridge/hub_remotes.json`）：hub **只探白名单里列出的端点、绝不自动发现主机**；没有配置时所有远端代码路径都不生效（行为与升级前一致）。失败一律结构化：`endpoint_unknown`（写了未配置的端点名）/ `endpoint_unreachable` / `endpoint_timeout`。
**为什么**：经验库是**全局单例**，不属于任何一台机器上的 IDE —— 放在插件里会让「没有 IDE 在线」时整块能力消失，也让「改一条经验」必须**重新编译插件**；移到进程外后它们**不开 IDE 也能用**，改经验只是改一个运行时读取的 JSON 数据文件。pid 是**每次启动都会变**的身份，默认目标一旦记住旧 pid 就会指向不存在的实例；`.vbp` 路径才是工程的稳定身份，也是唯一能同时表达「同一台机器上两个不同工程」与「两台机器上同一路径」的键（后者用 `endpoint` 前缀区分）。原先「不带 `_target` 就自动挑唯一活实例」在加入远端后会把调用**静默发到另一台机器**上；改成只在本机范围内自动选、远端必须显式指名，并把「只有远端在线」写成可读的拒绝码。远端主机既然需要显式配置，就不允许被自动发现：白名单之外不存在可发地址，探活只针对白名单，失败也给出结构化原因与耗时。

## 2026-10-13

### 16:26 · 工具面 114 -> 115 · 实时桌面状态 · 长代码可走文件 · 失败信封统一

**改了什么**：新增只读工具 **`desktop_probe`**：报告承载插件的进程的**实时**桌面状态（窗口站 / 桌面 / 会话 / 前台），任何一个 API 读不到就**如实拒绝**（`ok:false` + `failureKind`），**不再**返回一份填零的假象；`vba6_probe` / `immediate_eval` / `immediate_run` 各加一个**可选**的「从文件读入」参数（`exec_line_file` / `expression_file` / `code_file`）；`vba6_probe` 的失败统一成**信封**（`ok` / `failureKind` / `reason`），并回报**请求长度与实际截断后的长度**，越界参数给出结构化的 `arg_out_of_range` 而不是一句原始溢出错误；自绘面板读不到时，`ui_text` 给出更清楚的诊断字段；**公开的工具描述里清掉了内部标签**。
**为什么**：桌面 / 会话状态原先**没有任何工具可查**，调用方只能猜；长代码或带引号的代码经常被命令行的长度与引号规则弄坏；失败原因不统一，调用方无法区分「参数不对」和「引擎不可用」；描述里的内部标签会让公开面出现读者无法核对的字眼。另外，一处「前台值在本机恒为 0」的旧说法**已被实测推翻**并改正：桌面状态一律按**实时值**回报，且**绝不用窗口站名字推断「能否发送输入」**。★**立即窗口的正文仍然读不到**（该窗格不暴露可读正文）⇒ **没有**为它新增工具，这条限制如实保留。

### 15:08 · 编译不再卡住 · 快照与审计脚本的默认行为收紧

**改了什么**：模态框守卫的时限改为「**看到框就重新计时**」，且编译路径的时限覆盖**整个调用** ⇒ 编译不再卡住，而是返回结构化结果（IDE 的框由安全按钮清掉）；抓取工具面快照的脚本**默认不再写任何文件**，严格审计脚本**必须显式指定目标**并**拒绝本机保留端口**。
**为什么**：编译时 IDE 会弹出它自己的「生成工程」框，而该框出现得**比守卫时限还晚** ⇒ 调用一直不返回；快照脚本原先默认会**覆写契约基线**；审计脚本的默认目标恰好是**本机保留端口**。
### 14:50 · 断点前先判定该行是否允许 · 模态框安全守卫

**改了什么**：切换断点前**先判定该行是否允许**（不允许就**结构化拒绝**，不去触发 IDE 的报错框）；新增**模态框安全守卫** —— 调用被 IDE 自身弹出的模态框挡住时，**自动只点安全按钮**并**如实报告**（框的内容、点了哪个按钮、耗时、副作用）。
**为什么**：对不允许的行切换断点会让调用**挂住约一分钟**，并让后续的写操作全部判忙；一次误触发的抛错语句也会弹出模态框并**卡住整条工具链**。
### 14:23 · 新开代码窗格后滚轮立即可用

**改了什么**：打开**新的代码窗格**后**鼠标滚轮立即生效**（不再需要先点一下别处）；配套新增一个 **250 毫秒**的轻量挂接定时器与**逐窗挂接诊断字段**。
**为什么**：原先挂接新窗格的唯一驱动是**5 秒周期的看门狗**（实测 10 秒只扫 2 次）—— 新开窗口后的几秒内**滚轮必然无响应**，几秒后才自行恢复。
### 14:04 · 进程内执行与取值 · 会话级安静模式 · 闸门强制释放

**改了什么**：新增**进程内执行与取值**能力（**直接返回值并捕获输出**，调用方不必再自己写文件把数据带出来）；新增**会话级安静模式**（抑制工具抢前台）；调用被 IDE 自身的模态框卡住时，**闸门现在可以强制释放**，使后续调用恢复。
**为什么**：原先运行态验收必须**人工去看 IDE 立即窗口**；工具抢前台会让用户**正在进行的击键落进代码窗**；一次误触发的抛错语句会让**整条工具链锁死**且没有恢复手段。
### 13:29 · 行号累计 · 操作回执与拒绝 · 前台与窗格保护 · 条件编译参数

**改了什么**：组件的磁盘行号改为**按行累计**（存在隐藏的内联 Attribute 行时不再偏低）；「删除控件」与「打开工程」的返回补上**成功标志、前后计数和内存/磁盘差异**，不安全时**结构化拒绝**；涉及光标与窗格的工具会**报告并恢复前台**，并在操作前后**校验窗格文本未变**；条件编译参数的值**不再被加引号**，回执给出实际命令行。
**为什么**：行号偏低会让按行号改动静默改错行；「看起来成功」或「静默不生效」会让调用方误判；用户正在同一台机器上打字时，抢焦点可能让击键落进代码窗、改坏源码；加引号会让条件编译宏静默失效。
### 10:05 · `remove_trace` 不再自动移除 `modMcpTrace`

**改了什么**：`remove_trace` 只删除注入的跟踪代码块，不再删除 `modMcpTrace` 模块。
**为什么**：那个模块里可能有你自己加的代码，工程里其它地方也可能还在用它 —— 删掉会丢代码或让工程编译不过。

### 09:53 · 公开文档不再引用无法验证的外部产品

**改了什么**：README 与代码注释里去掉了对某个未开源外部产品的溯源；示例配置改用占位路径。
**为什么**：读者无法核对那个外部参照，只会困惑；写死的本机路径对别人也直接用不了。

### 03:16 · 首个公开版本

以源码形式开源（Apache-2.0）：111 个 MCP 工具，把真实的 VB6 IDE 变成可编程接口。
