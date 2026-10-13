# MCP 工具契约（docs/CONTRACTS.md）

本文件是 VB6MCP **工具面（共享事实）** 的契约正本。工具集/描述/参数的任何变化都必须：
S 预写本节并 `bus.mjs contract` 广播（主卡 + 受影响进程确认卡 + 版本递增）→ 实现方落地 →
`python 维护者的契约导出脚本（未随源发布）` 出 live 快照 → **`--freeze` 显式重冻** `tools/schema_baseline.json` →
`check_schema_baseline.py` / `check_live_tools.py` / `mcp_strict_audit.py` 全绿 → P4 独立复测。

## 版本记录

> **改契约的必做一步**：本表新增/变更版本后，**必须同步 add-in 的唯一常量**
> `Util.bas` 的 `VB6MCP_CONTRACT_VERSION` —— 它同时供 `initialize.instructions` 首行、
> `mcp_info.contract` 与每次 `tools/call` 的 `mcp` 戳使用。**禁止在别处再硬编码契约版本**：
> 横幅必须是**永不陈旧的自称**（与 `BUILT` 戳同理；hub 转发**活体**横幅，不一致会直接暴露给用户）。

| 版本 | 日期 | 变更 |
|---|---|---|
| **v1.10.0** | 2026-10-13 | **BATCH-1：能力簇落地 + 描述标签中性化**。Tool face: **add-in 114 -> 115**（新增只读 `desktop_probe`），hub 5 -> 5，DSH-facing total **119 -> 120**。**ADDED** = `desktop_probe`（只读：进程的实时窗口站/桌面/sessionId/可交互性/前台状态；**fail-closed** ⇒ 任一 API 失败即 `ok:false` + `failureKind:"api_unavailable"`）。**CHANGED（12）** = `build` / `close_active_dialogs` / `ide_quiet_mode` / `ide_state` / `immediate_eval` / `immediate_run` / `inject_trace` / `make_project` / `remove_trace` / `todo_status` / `validate_code` / `vba6_probe`：其中 10 个只改**描述文本**（内部标签中性化，行为与回执不变），`ide_state` 与 `vba6_probe` 描述未变、只加参数。**新增参数（3，纯增量）** = `vba6_probe.exec_line_file` / `immediate_eval.expression_file` / `immediate_run.code_file`。**R13 失败信封**：`vba6_probe` 七类 `failureKind`（`not_loaded` / `alloc_failed` / `arg_out_of_range` / `address_invalid` / `not_readable` / `engine_zero` / `refused_by_design`），`immediate_*` 的 `exec_error` 包装新增两类（`channel_error` / `statement_raised`），且 `immediate_*` **自身**的 `exec_error` **刻意不带** `failureKind`。**修复**：静默截断改为结构化（`lenRequested` / `lenClamped`、`maxHits*`）；`arg_out_of_range` 取代裸 VB 溢出。**未落地**：`immediate_window_text`（R1 NO-GO ⇒ 不加该工具）。**D1 有界残留**：`todo_status` 保留功能性自动发现（仓库内**可选任务文件夹**）并新增环境变量 `VB6MCP_TASKS_DIR` 覆盖；契约与公开面不再出现该目录名。DELIBERATE, S-authorized re-freeze（四冻件全变；hub 快照随之重生成）。 |
| **v1.9.0** | 2026-10-13 | **工具面扩容 + 卡死可恢复**。Tool face: **add-in 111 -> 114**（新增 `immediate_eval` / `immediate_run` / `ide_quiet_mode`），hub 5 -> 5，DSH-facing total 116 -> 119。**Add-in CHANGED** = `ide_state`（新增 `release_inflight` / `release_min_held_ms`；正文**恒带** `release{}` 固定键集）；只读回执仅在**发生过强释**后才追加 `gate*` 键；`ide_quiet_mode` ON 时**每次** `tools/call` 正文带 `quiet{...,advisory}`、`control_window{action:activate}` 回执带 `quietModeSuppressedFocus`。**纯增量**：quiet OFF / 未强释路径的既有回执逐字节不变。DELIBERATE,S-authorized re-freeze（四冻件全变）。|
| **v1.8.0** | 2026-10-12 | P1 fix pack. Tool face: add-in 111 -> 111,hub 4 -> 5 (new `ide_restart`),DSH-facing total 115 -> 116. Add-in CHANGED = `replace_code` (replace/delete now REQUIRE `expected_sha256`; the refusal returns `currentSha256`),`validate_code` (+optional `baseline`; every diagnostic gains `proc`/`procOffset`/`fingerprint`/`stableKey`; always-present `baseline_diff`),`ide_state` (+`cancel_inflight`/`cancel_target`; every read-only receipt gains `inflight{}`). Behaviour fixes with no schema change: PRE-WRITE block-pairing refusal in replace_code/add_code/apply_edits/edit_by_anchor; multi-line (block) anchor + duplicate-anchor hits + procedure first/last echo in edit_by_anchor; close_active_dialogs emits ONE key set on dry_run and real clicks; build/make_project stop publishing an EMPTY digest and re-read logs in their own encoding. DELIBERATE,S-authorized re-freeze. |
| **v1.7.0** | 2026-10-09 | **版本号补齐（additive，不改工具面 ⇒ 免重冻）**：为 `replace_code` / `apply_edits` / `edit_by_anchor` 新增的**可选** `snapshot`（`type=boolean`，**默认 off ⇒ 默认行为不变**，跨构建 A/B `ab_differences=0` 为证）补上契约版本号与迁移条目 ⇒ **工具数不变（111 → 111）**；对应冻结件 sha **`1EE3BFA9E917AB6E5D148D409FB6B33ACA91EF2B20C7988E39ADE589536108BE`**（== `docs/tools_schema.json`，本卡**不再重冻**）；`mcp_info.breaking` 本版**不新增条目**（非破坏性），迁移说明见 **§9** |
| **v1.6.0** | 2026-10-09 | **工具面扩容**：新增 `replace_in_files`（跨组件字面/正则替换，`dry_run` 默认 true）与 `undo_edit`（组件级显式单步撤销；写入前把 pane 文本快照到 `<工程>\_undo\`）⇒ **109 → 111，必须显式重冻**；同一次重冻并落地 `send_window_message` 的**描述补全 + 高危消息护栏**（新增参数 `confirm` ⇒ 本版 `CHANGED=send_window_message`）；另**登记**（不改 schema）`add_component{type:"form"}` 改由 IDE 自建 + `createdVia`/`ideRouteNote`、`control_execution.mode` 的 12 值 enum、以及两条工具面局限（`command_bars{list}` 不含 Add-Ins 弹层；`AddFromTemplate` 窗体模板路线本机不可用） |
| v1.0.0（隐含） | 2026-09-21 | 既有 96 个工具（此前只有 live 快照与冻结基线，没有书面契约） |
| **v1.5.0** | 2026-10-09 | **消息型窗口原语 + 可选滚轮 + 写入回执的相位/分块字段**：新增 `read_active_dialogs`（只读）/ `close_active_dialogs`（**只显式调用**）/ `send_window_message`（**只发窗口消息**），**106 → 109，必须显式重冻**；`command_bars` 在**既有 `control` 参数内**接受 `#<id>` 寻址（**不改 inputSchema**）；写入回执新增 `hunks` 与相位字段（**纯增量**，见下）；可选代码窗滚轮（env/注册表，**默认关**，**不新工具、不改 schema**） |
| **v1.4.0** | 2026-10-09 | **OCX 防挂死**：`add_control` 预校验（未登记部件 ⇒ `control_part_not_registered`，绝不进入会挂死的调用）+ 新只读工具 `list_component_parts`（**105 → 106，需显式重冻**） |
| **v1.3.2** | 2026-10-09 | **结构化错误号**（响应侧，无需重冻）：失败回执增可选 `vbErrorCode`/`hresult`（8 位 hex）/`errorLine`；`mcp_info` 可增 `errorDecoding`/`errorTableEntries` |
| **v1.3.1** | 2026-10-08 | **多 IDE 归属确定化**：`SO_EXCLUSIVEADDRUSE` + 动态端口 + 实例注册（`%TEMP%\vb6mcp\instances\<pid>.json`）+ 响应正文 `ide:{pid,project,port}`（不触 inputSchema、无需重冻；不变式不变） |
| **v1.3.0** | 2026-10-08 | **版本变更可见化**：新增只读 `mcp_info`（104 -> 105）+ `initialize.instructions` 固定版块 + 每次 `tools/call` 正文带 `mcp:{product,contract}` 版本戳 + 产品版本 0.2.0 -> 0.3.0 |
| **v1.2.0** | 2026-10-08 | **工具加固**（从真实工程实战反推）：α 回执诚实性（纯增量）+ β 行为加固（**破坏性**）+ γ 输入参数（触 inputSchema）+ δ 3 个新工具；分 A 批（α+β）与 B+C 批（γ+δ）两次落地，见下节 |
| ~~v1.1.0~~ **v1.1.0（含勘误）** | 2026-09-22 | **新增 5 个只读工具**：`ide_state` / `current_position` / `list_breakpoints` / `find_gui` / `ui_text`（96 → 101）；另有 1 条非 schema 变更（断点态写权限错误的 `reason`，见下） |

## v1.1.0 新增工具

### 1. `ide_state`（G-1）
- 参数：无
- 返回：`{state:"design"|"run"|"break", ebMode:0|1|2（design=0 / run=1 / break=2）, title:"<主窗口标题>", project:"<工程名>", running:bool}`
- 语义：**实时**读取 IDE 模式；不得返回上一次动作前的陈旧值（design/run/break 切换后立即调用必须正确）。
- 只读。用途：所有调试类判据的前置条件。

### 2. `current_position`（G-3）
- 参数：无
- 返回：断点态 `{available:true, project, module, proc, line, text, ebMode}`；非断点态 `{available:false, reason:"not_in_break"}`（**不报错**）。
- `line` 为模块内 **1-based** 行号；`text` 为该行原文（不含行号前缀）。
- 只读。用途：单步/继续的「位置变化」取证。

### 3. `list_breakpoints`（G-4）
- 参数：无
- 返回：`{injected:[{component, line, kind:"injected", enabled}], native:null, note:"..."}`
- 语义：枚举**本工具注入的**断点标记（`set_breakpoint` / `native_breakpoint` 写进源码的标记行）；
  IDE 原生断点**不可枚举**时必须返回 `native:null` 并如实说明（VBIDE 对象模型不提供断点 API），**不得伪造**。
- 只读。用途：断点类证据从「自述」变「可枚举」。

### 4. `find_gui`（G-5）
- 参数：`keyword`（子串，可空）、`cls`（窗口类名，可空）、`showing`（bool，可空）；三者可组合。
- 返回：`[{hwnd, cls, text, showing, enabled, rect:{l,t,w,h}, path:"<父链标题>"}]`
- 语义：**只在当前进程内**（VB6 IDE 进程）枚举顶层窗口与子控件；`keyword` 为 contains 匹配。
- 只读。用途：定位对话框/菜单/弹窗文本。

### 5. `ui_text`（G3 未闭环的 UI 观测方向，/ R47 方向2）
- 参数：`hwnd`（可空）、`cls`（可空，取第一个匹配）、`depth`（默认 3）。
- 返回：`{ok:true, source:"uia|msaa|wm_gettext", text:"..."}` 或 `{ok:false, reason:"no_accessible_text", evidence:"<可复现负向证据>"}`。
- 语义：读取指定窗口及其子树的**可访问文本**；优先 UI Automation，回退 MSAA（`AccessibleObjectFromWindow`），再回退 `WM_GETTEXT`。
  目标场景是**自绘面板**（断点态的 Locals / Watch / Immediate 无子控件）。
- 只读。**负向结果同样算交付**，但必须给可复现证据与结论，不得伪造。

## v1.1.0 非 schema 变更（G-7）

断点态下**写源码类**工具（`set_breakpoint` / `clear_breakpoints` / `apply_edits` / `write_source` 等）返回的裸
VB6 权限错误（`拒绝的权限`）必须追加 `reason:"ide_busy_in_break_mode"` 与恢复建议（先 `end` 退出断点态）。

**v1.1.0 勘误 2（2026-10-08，真机实测）**：**运行态**写源码同样会被拒（裸 `拒绝的权限`），追加值为 `reason:"ide_busy_in_run_mode"` + `ideState`。属**新增取值**（不改既有值语义）。

**v1.1.0 勘误 3（2026-10-09，源码复核）**：§3 的语义里「`set_breakpoint` / `native_breakpoint` 写进源码的标记行」**不成立**。
实现（`CodeTools.cls` `ToolListBreakpoints`，只扫描 `' @mcp-break <id> begin` 标记行）⇒ **只有 `set_breakpoint` 注入的断点会被列出**；
`native_breakpoint` 走 Debug 菜单的 toggle，**不写任何源码标记**，因此**不会**出现在 `injected[]` 里。
正确口径：`list_breakpoints` 只枚举 **set_breakpoint 注入物**；**IDE 原生断点全局不可枚举** ⇒ 一律 `native:null` + `note`（description 早已如此，本次只修契约正文）。
该勘误**不改 inputSchema/输出字段** ⇒ **不需要重冻**。

## v1.2.0 工具加固（从真实工程实战反推）

来源：`现场反馈记录（内部文档）`、`工具加固计划（内部文档）`。
**两批落地**：**A 批 = α + β**（不改 inputSchema，只加响应字段 + 改两处默认行为）；**B+C 批 = γ + δ**（改 inputSchema + 新增工具）。
每批都必须走：实现 → `维护者的契约导出脚本` → `check_schema_baseline.py`（**必须先看到 `schema_changes≠0` 才准重冻**）→ `--freeze` → `check_live_tools.py`（ORDER_MATCH=True）→ `mcp_strict_audit.py`（strict_failures=0）→ P4 独立复测。

### A 批 · α 回执诚实性（纯增量，不破坏调用方）

| 工具 | 新增响应字段 | 语义 |
|---|---|---|
| `get_component_code` / `find_in_component` / `replace_code` / `apply_edits` / `edit_by_anchor` / `preview_edit` / `verify` | `coords: {headerOffset, ideLine, diskLine, default}` | 明确每个行号属于哪套坐标系；`headerOffset = diskLine - ideLine`；`default` **按工具如实给**（`get_component_code` = `"disk"`（它把 header 拼回了，编号就是磁盘行号）；编辑类 = `"ide"`）；另带 `resolved`（精确头部行数是否解析成功；`false` 时 offset 不可信，见勘误 3） |
| `replace_code` / `apply_edits` / `edit_by_anchor` | `requestedRange` / `appliedRange` / `clamped` | 区间被裁剪时必须可见（本批仅标记，拒绝行为见 β-1） |
| `apply_edits` / `replace_code` | `ideTotalLines`（`SetCompText` 后重读 `cm.CountOfLines`）、`syncWarning` | 与文本算出的 `newTotalLines` 并列；两者不一致时给 `syncWarning` |
| `build` / `make_project` / `verify` | `artifactPath` / `artifactMtime` / `artifactChanged` / `artifactLockedHint`；`ok:true` 但产物未变化时 `verdict:"compiled-no-artifact-change"` | 编译成功 ≠ 产物已更新（被锁场景） |
| 编辑类 | `diskSha256` / `paneSha256` / `onDisk` | 自证「报的哈希是盘上还是代码窗格」 |
| `edit_by_anchor`（拒绝时） | `hits: [{line, proc, text}]` | 锚点不唯一时列出候选，而不是只说 `anchor not unique (N)` |
| `preview_edit` | `diff` + `hunks[{kind,start,count,oldStart,oldCount,newStart,newCount,contextLines}]` + `contextLines`（默认 3）+ `contextApplied` + `algorithm` + 可选 `algorithmNote` + `lineDelta` | ★**勘误（F7）**：本行旧文写「diff 仍是**前后缀裁剪**」**与实现不符** —— `preview_edit` 走 `EditTxn.DiffPlan` 的**逐行 LCS**（行级 LCS + hunk 分组 + 默认 3 行上下文），且 `diff` 与 `hunks` **同源**（contract invariant 8，二者不可能互相矛盾）；`algorithm` 如实报 `lcs` / `fallback-too-large`（输入过大时降级，回执带 `algorithmNote` 明说）/ `none`。纯预览，**不写盘也不写 IDE** |

### A 批 · β 行为加固（**破坏性**，调用方必须知情）

1. `replace_code` 的 `replace`/`delete`：`end_line > CountOfLines` 或 `start_line > CountOfLines` ⇒ **拒绝**，返回 `{ok:false, reason:"range_beyond_end", requestedRange, countOfLines}`；**不得静默钳制、不得退化成 insert、IDE 内容不得改变**。
2. 写类工具（`replace_code` / `apply_edits` / `edit_by_anchor` / `add_code` 的插入路径）：插入点或替换区间落在 **VB 续行**（上一逻辑行以 `_` 结尾）内部 ⇒ 拒绝，返回 `reason:"inside_continuation"` + 建议；不得写入。
3. 写类工具在**写入前**对本批结果跑 `Vb6Lint`（E101–E116 + E107），把 `diagnostics[]` 与 `lintSummary{errors,warnings}` 放进响应；**默认 `warn`，不阻断**（`strict` 见 γ）。

**迁移说明（A 批）**：旧行为「超界静默钳制 + 回执 `deletedLines:0`」被「结构化拒绝」取代 —— 任何把 `deletedLines:0` 当成功判据的脚本都必须改判 `ok`/`reason`。

### B+C 批 · γ 输入参数（**触 inputSchema**）

| 工具 | 新参数 | 语义 |
|---|---|---|
| `get_component_code` | `coords: "ide"|"disk"`（默认 **`"ide"`**） | `"disk"` = 现有行为（把 header 拼回，磁盘坐标系）；`"ide"` = **纯代码区文本**。`include_line_numbers` 与 `coords` 必须一致 |
| `get_component_code` | `proc: "<过程名>"` / `symbol: "<标识符>"` | 只返回该过程/符号的范围：`{startLine, endLine, text, sha256, coords}`；大模块不必整 dump |
| `replace_code` | `expected_sha256` | 与 `edit_by_anchor`/`apply_edits` 共用 `GuardSha`；不一致 ⇒ 拒绝且**不改** |
| 写类工具 | `validate: "off"|"warn"|"strict"`（默认 `"warn"`） | `strict`：有 `severity=error` 的 structural 诊断 ⇒ **拒绝写入、IDE 与磁盘都不变** |
| `edit_by_anchor` | `occurrence: N`（1-based 第 N 次命中）、`scope: {proc?, start_line?, end_line?}` | 允许处理重复行；不足 ⇒ `reason:"anchor_not_found"` |
| `edit_by_anchor` / `apply_edits`（anchor） | `match: "normalized"|"exact"`（默认 `"normalized"`） | `normalized` = 现状（忽略缩进与大小写）；`exact` = 按原始行比较 |
| 锚点响应（成功路径） | `matchedText` / `matchedLine` | 回显**实际命中的原始行** |

### B+C 批 · δ 新工具（一律只读）

1. `workspace_status()` → `{dirtyComponents:[...], git:{branch, ahead, behind, files:[{path,status}]}, diffStat:"..."}`：IDE 未保存组件 ↔ git 未提交清单。
2. `verify_scope({components?, path?})` → 逐组件 `{component, ideEqualsDisk, lintErrors, lintWarnings, firstDiagnostic}` + **汇总**（只回 FAIL 明细）。
3. `todo_status()`（优先级最低，可选）→ 从 git + 证据文件生成可机读待办。

### v1.2.0 追加不变式

4. `coords` 必须出现在**所有**坐标/编辑类响应里；`default` 在 B+C 批之后为 `"ide"`。
5. **`get_component_code` 的默认口径在 B+C 批改为 `coords:"ide"`（破坏性：旧调用方拿到的文本会少掉 header）**；需要磁盘坐标必须显式 `coords:"disk"`。迁移：把 `include_line_numbers` 的读数视为 IDE 行号，或显式传 `coords:"disk"`。
6. 新增字段不得移除或改语义既有字段；`ok` / `reason` 语义在 v1.2.0 内保持稳定。
7. `validate:"strict"` 的拒绝必须发生在**写入之前**（IDE 与磁盘都不变），并附 `diagnostics[]`。
8. `preview_edit` 的 diff 在 B+C 批替换为**行级 LCS + hunk 分组**；替换前（A 批）必须在 description 里写明其算法局限。

### v1.2.0 A 批落地勘误与裁定（2026-10-08，真机实测；S 裁定）

- **勘误 1（断言被证伪）**：本契约与判据里的「`.bas` 的 header=0」**不成立**。VB6（连 `import_component` 在内）会往 `.bas` 写 `Attribute VB_Name` 行 —— **代码窗格没有、文件有**，故实测 `headerOffset=1`；只有**手工写成、文件内无 Attribute 行**的 `.bas` 才是 0。`.frm` 实测 17。
- **勘误 2（`default` 语义）**：见上表 α-1 —— `default` 是**每工具如实标注**，不是全局 `ide`。
- **裁定 1（`get_component_code` 响应改 JSON 信封）**：**接受**。理由：α 表要求该响应带 `coords`，裸文本无法承载字段；P1 已证 `text` 字段与改动前**字节逐一相同**（含 `include_line_numbers` 的编号）。列入迁移说明：调用方需从 `result.text` 取文本（旧的裸文本响应不再返回）。
- **裁定 2（`resolved` 字段）**：**接受**为 α 的组成部分（`coords.resolved`）。窗格首行在磁盘上找不到时（例如刚在首行插入）必须 `false` 且注明不可信，**不得**把回落值当精确值。
### v1.2.0 落地补记 · 缺陷批次（2026-10-08，；同一棵树 101 工具，DLL `8795F5B7…` @18:57:26）

缺陷修复带来的**行为与描述变化**（`dump` 显示 `CHANGED=set_breakpoint,set_project_options`，已显式重冻）：

1. `write_file`：写入**必须截断**（旧行为短内容残留旧尾部）。
2. `SetCompText`（内部）：写入后 `cm.CountOfLines` **不得**比文本多一行；`ideTotalLines == newTotalLines`，不一致才给 `syncWarning`。
3. `set_breakpoint`：锚点落在**过程体外**（声明行/模块级）⇒ 拒绝，`reason:"anchor_outside_procedure"`，**IDE 与磁盘 sha256 均不变**；判定用 `Vb6Lint.ProcBodyOwner`（VBIDE 的 `ProcOfLine`/`ProcBodyLine` 会放行声明行，不可用）。
4. `set_project_options(startup_object)`：VB6 侧**不可行**（`Sub Main`/模块名/窗体名均报 `无效的过程调用或参数`）⇒ 结构化返回 `startupObjectError` + `startupObjectWorkaround`（可行路线：直接写 `.vbp` 的 `Startup=` 再 `save_project`，实测起动后 `run` 断在 `modMain.Main`）。结论见 `内部已知限制与 FAQ` §16。
5. `onDisk`：`.frm`/`.bas` 连续两次调用取值稳定（正文与窗格一致时 `true`）。

**S 记：流程偏差（如实记录）** —— 本批的契约条目**未能在实现之前写进本文件**（缺陷批次直接从「立卡」走到了「落地」）。事后补记如上，事实口径已与实现一致。**规则照旧**：此后任何工具集/描述/参数/行为变化，必须**先**写契约广播、**再**实现；B+C 批（γ+δ）按此执行。

- **裁定 3（两个既存缺陷）**：**都立卡**（`write_file` 不截断 = 数据完整性级；`SetCompText` 写后多一行 = 「回执行数撒谎」的机制本体）。另有 §7 报出的 3 条既存缺陷（`set_breakpoint` 允许过程外锚点导致编译失败、`set_project_options(Sub Main)` 失败、`onDisk` 抖动）一并立卡。
- **裁定 4（失败响应的 `coords`）**（2026-10-08，提出）：**任何含行号的响应都必须带 `coords`**，包括**失败**响应（`range_beyond_end`/`inside_continuation`/`anchor_not_found`/`anchor_outside_procedure` 等）：其中 `requestedRange`/`appliedRange`/`hits[].line`/`countOfLines` 的行号一律按 `coords.default` 解释。无行号的失败（如纯 `reason`+`evidence`）可不带。
- **口径 5（`textSha256` 不是文件哈希）**（2026-10-08）：`get_component_code` 的 `textSha256` 是**返回文本**的哈希（文件末尾 CRLF 等差异会让它与 `read_source` 的字节 sha256 不同）。判据里不得把两者当同一个值比对；要文件哈希请用 `read_source`。
- **勘误 3（`.frm` 的 headerOffset 不是常量）**（2026-10-08，实测）：**工具生成**的 `.frm` headerOffset=**17**；**手写**的 `.frm` = **16** 且 `coords.resolved:false`（窗格首行在盘上定位不到）。故 `17` 是「工具生成布局」的性质，**不是** `.frm` 的通用常量；判据里不得写死 `.frm=17`。同理 `.bas`：含 `Attribute` 行=1、手工无 Attribute=0。

## v1.3.0 版本变更可见化（避免调用方按旧思路调用）

### 问题
工具面在 v1.1.0（+5 工具）与 v1.2.0（A 批语义加固 + B+C 批 7 参数 + 3 工具）之间发生了**多处破坏性/语义变化**，但 add-in 自报的 `serverInfo.version` 仍是 **0.2.0**，调用方无从察觉。已实测的踩坑类型：行号坐标默认值变了、超界从静默钳制变拒绝、`write_file` 变截断、`compiled:true` 不再等于产物更新 —— **调用方按旧思路调用会拿到「看起来成功、实际错」的结果**。

### 三个机制（本版一并落地）

| # | 机制 | 位置 | 是否触 inputSchema |
|---|---|---|---|
| 1 | **`initialize` 的 `instructions` 顶部固定版块** | JSON-RPC `initialize.result.instructions` | 否 |
| 2 | **每次 `tools/call` 的文本载荷里带版本戳** `mcp:{product, contract}` | 各工具返回的 JSON 正文（**不是**顶层结果键） | 否 |
| 3 | **新增只读工具 `mcp_info()`**：机器可读的版本与破坏性变更清单 | 工具表 | **是（104 -> 105）** |

### 1. `instructions` 版块（首行必须可一行读完）
```text
PRODUCT vb6-mcp <semver>  CONTRACT v1.10.0  TOOLS 115  BUILT <sha256 前 8 位>
BREAKING (read before calling):
- v1.10.0 +1 tool: desktop_probe (read-only live desktop/interactivity identity; fail-closed: ok:false + failureKind="api_unavailable" when an API fails). +3 OPTIONAL file-fed params (vba6_probe.exec_line_file / immediate_eval.expression_file / immediate_run.code_file, ADDITIVE). 12 tool DESCRIPTIONS were neutralized (internal label text only - no behaviour change); vba6_probe gains a 7-value structured failure envelope and the immediate_* exec_error wrappers add channel_error / statement_raised (ADDITIVE)
- v1.9.0 +3 tools: immediate_eval / immediate_run / ide_quiet_mode; ide_state gains release_inflight / release_min_held_ms and a release{} block; ide_quiet_mode ON adds quiet{...advisory} to every receipt (ADDITIVE: quiet OFF and the no-release path are byte-for-byte unchanged)
- v1.6.0 send_window_message: WM_CLOSE/WM_QUIT/WM_SYSCOMMAND need confirm=true; the parameter is message (not msg) and raw numbers are accepted
- v1.6.0 add_component(type=form): the form is built by the IDE itself; the receipt adds createdVia (ide_dialog|synthesized) and the call takes ~3.5-4.9 s
- v1.6.0 +2 tools: replace_in_files (dry_run defaults to true; only the OPEN project) and undo_edit (pane-only, explicit)
- v1.7.0 snapshot param (ADDITIVE, non-breaking): replace_code / apply_edits / edit_by_anchor accept snapshot=true (on/yes/1) to guarantee exactly ONE pre-edit pane snapshot that undo_edit can restore; default off (set / apply_edits / edit_by_anchor already snapshot by default)
- get_component_code: coords default disk -> ide (pass coords="disk" for the old text/numbering); response is a JSON envelope, text in result.text
- replace_code: out-of-range/continuation now REJECTED (range_beyond_end / inside_continuation); deletedLines=0 is no longer a success signal
- write-class tools: validate=off|warn|strict; edit-side failures carry ok:false + reason
- write_file now truncates; build/make_project/verify: judge by artifactChanged/verdict (compiled-no-artifact-change = NOT built)
- v1.1.0 added ide_state/current_position/list_breakpoints/find_gui/ui_text; v1.2.0 added workspace_status/verify_scope/todo_status
```

### 2. 版本戳（键固定为 `mcp`，只增不改）
`"mcp":{"product":"<semver>","contract":"v1.3.0"}` 与既有响应字段并列；**不得**为此新增顶层 JSON-RPC 结果键（保住 `tools/call` 结果键 = `{content,isError}` 的严格性）。

### 3. 新工具 `mcp_info`（只读）
- 参数：无
- 返回（示例）：
```json
{"ok":true,"product":"0.3.0","contract":"v1.3.0","toolCount":105,
 "build":{"sha256":"<64hex>","builtAt":"YYYY-MM-DD HH:mm:ss"},
 "breaking":[{"since":"v1.2.0","what":"get_component_code 的 coords 默认由 disk 改为 ide","instead":"需要旧行为请显式传 coords=disk"},
             {"since":"v1.2.0","what":"超界不再静默钳制，改为 range_beyond_end 拒绝","instead":"用 expected_sha256 + 现取行号，或改用锚点编辑"},
             {"since":"v1.2.0","what":"write_file 改为截断写","instead":"无需变更，旧「残留尾部」行为已修"},
             {"since":"v1.3.0","what":"新增 mcp_info 工具与版本戳","instead":"—"}],
 "migrations":["行号混用防护：读响应里的 coords.lineNumbersAre / editToolsTake / convertToEditLine"],
 "changelog":["v1.1.0 +5 只读工具","v1.2.0 A 批回执诚实性+守卫","v1.2.0 B+C 批 7 参数+3 工具","v1.3.0 版本可见化"]}
```
- 只读、无副作用；只报自身构建信息，不读盘上任意文件。

### 4. 版本号与不变式
- **产品版本必须随语义变化递增**（本次 0.2.0 -> **0.3.0**）；`initialize` 同时给出 `contractVersion`。
- 机制 1/2 为**纯增量**；机制 3 是新工具 ⇒ 必须**显式重冻**（104 -> 105）并同步 catalog 与 hub 快照。
- **不变式**：`tools/list` 结果键仍**恰为** `{tools}`；`tools/call` 结果键仍**恰为** `{content,isError}`；版本戳只出现在正文里。


## v1.3.1 多 IDE 归属确定化（传输侧硬化 + 响应侧 `ide` 回显）

### 背景（当前缺陷）
add-in 用 `SO_REUSEADDR` 绑 8766 ⇒ 两个 IDE 都能绑上、**端口归属未定义**；hub 只能扫端口「猜」目标。

### 本版内容（**不触 inputSchema、无需重冻**）
1. **传输侧**：
   - 监听改 **`SO_EXCLUSIVEADDRUSE`**（或不设 REUSEADDR 并检查 bind 结果）：第二个实例必须**绑不上并结构化报冲突**，不得静默共享端口；
   - 端口被占时**动态探测空闲端口**（`VB6MCP_PORT` / 注册表覆盖仍优先），并把实际端口写进启动日志与回执；
   - **实例注册**：`%TEMP%\vb6mcp\instances\<pid>.json`（`pid/port/project/projectPath/dllSha256/startedAt/startedByHub`），退出注销；陈旧项由**读方**按「pid 是否存活 + 端口是否有人听」双重校验处理（**写方不清理别人的**）。
2. **响应侧**：每次 `tools/call` 的**正文**增 `ide:{pid,project,port}`（与既有 `mcp` 戳并列）。

#
## v1.3.2 结构化错误号（响应侧；采纳 P2 的 R60–R71 研究）

### 背景（静态层已具备条件，见 的 R60–R71）
- `sub_FA9A05A`（**321 xref**）= 全局 **HRESULT → VB 错误码** 转换器：`if (a1>=0) return 0;`，否则按 facility 分流 + **85 项二分表**（`dword_FA9A0E0` RVA `0xA0E0`，步长 8B，终止 `[84] 0xFFFFFFFF→440`）；未命中 ⇒ `dwMessageId = 原始 HRESULT`。
- 错误号在 `word_FC0D18A`；描述两条路（全局 ANSI `word_FC0D18C`，或 `sub_FBC53CD(号)` 按号现造）；行/列由 `sub_FB8749A` **两阶段解析**（位置全局 `word_FC0D188` / `dword_FC0D190` / `word_FC0D194-198`）。
- **消费方**：(a) 回执错误诊断（VB 码 → HRESULT → 根因类别；如 429 覆盖 **6** 个不同 HRESULT）；(b) 失败模式归因；(c) **测试断言可用数值码，跨中英文 IDE 稳定**；(d) 知识沉淀。

### 内容（**纯响应侧 ⇒ 不触 inputSchema、无需重冻**）
在**失败回执**里就地增以下**可选**字段（能取到才给）：
- `vbErrorCode`（十进制数字）：VB6 数值错误号（add-in 自身 `Err.Number`，或被测工程侧 `word_FC0D18A`）；
- `hresult`（**8 位十六进制字符串**，如 `0x80040154`）：原始 HRESULT（`sub_FA9A05A` 未命中表时写入的 `dwMessageId`）；
- `errorLine`（数字，`coords:"ide"` 口径）：能解析出的出错行（位置协议）。
- （可选）`mcp_info` 增 `errorDecoding:true` 与 `errorTableEntries:85`。

### 不变式
- **绝不改变**既有 `ok`/`reason`/文案的语义；新字段是**增量**。
- `tools/call` 顶层键仍**恰为** `{content,isError}`；`tools/list` 仍**恰为** `{tools}`。
- **取不到就不出现该字段**（不得用 0/空串假装有值）；读取失败**不得**影响原错误路径。

### 勘误/落地补记（2026-10-09，真机实测后）

- **`hresult` 在本工具面下不可实现 ⇒ 永不出现**：`dwMessageId`（VA `0x0FC0D10C`）会被 **IDE 自己消费并清零**（`sub_FBC53CD`），采样是竞态；且只读实测显示**弹着 `#32770` 时三处采样全为 0**（连 `word_FC0D18A` 都不含 VB6 自己显示的 429）⇒ 按「只在可判定时出现」的不变式，该字段**永不出现**（证据： /）。
- **`errorLine` 已撤回实现**：在 add-in 内**裸读**引擎位置全局（`RtlMoveMemory`）曾**导致 IDE 进程死亡**（A/B 实证，，已记 FAQ §19）；是否属「**绝对 VA 被再次加基准**」的混淆待判定 ⇒ 在判定与正确地址下复测之前，该字段**不实现**。
- **`vbErrorCode` 是唯一已交付字段**：仅来自 add-in 自身的 `Err.Number`（在 `CallTool` 的 Fail 入口**一次性**读取并 `Err.Clear`）；我们自己发的结构化拒绝**没有**该字段（合法缺失）。
- **有模态框时引擎错误全局一律不可用**：替代来源按可靠度 —— **对话框 Static 文本（`find_gui` + `ui_text`）> `vb6mcp.log`/工具回执 > `<Component>.log`**；引擎全局（含 `word_FC0D18A`）在有模态框时 **❌**（见 FAQ §15.3）。

## v1.4.0 OCX 控件部件：防挂死 + 只读清单（**+1 工具 ⇒ 必须显式重冻 105→106**）

### 背景（根因已由真机钉死）
把某个 OCX 当「控件部件」用时，若该部件**未登记进工程**（`.vbp` 无对应 `Object=` 行），`VBControls.Add(ProgID)` 会进入 **VBIDE 内部静默等待**：**不弹任何对话框**（五格 ProgID 矩阵新增窗口数全 0）、**整条 UI 线程被占**（第二个客户端的 `ide_state`/`get_project_info` 双双 5s 超时）、观测 **683s 不返回**、审计零行、无 `<form>.log`；**触发与 OCX 是否注册无关**（虚构 `Foo.Bar.Baz` 同样挂死；CLSID 已注册的 `MSComDlg.CommonDialog` 也挂死），正向对照 `VB.CommandButton` 3s 成功。⇒ **进程内不可修复**（单线程 STA 无法中断已阻塞调用）。另：**手写 `.vbp` 的 `Object=` 行在本机不成立**（`open_project` 正常但重开后工程 components 为空、窗体不加载）。

### 内容
1. **`add_control` 预校验（防挂死）**：调用 `VBControls.Add` **之前**判定 `prog_id` 是否属于**本项目已登记的控件部件**；不属于 ⇒ **结构化拒绝** `{ok:false, reason:"control_part_not_registered", progId, knownParts:[...], advice}`，**绝不进入那次可能挂死的调用**。
2. **新只读工具 `list_component_parts`**：列出当前工程**已登记的控件部件**（ProgID / 部件名 / 文件路径 / GUID，能取多少给多少），供调用方在加控件前自查。

### 不变式
- `list_component_parts` **只读**；工具数 **105 → 106**，**必须显式重冻**并同步 catalog 与 hub 快照。
- `add_control` 的**成功路径行为不变**（合法 ProgID 照旧）；只**新增**拒绝分支与结构化 reason。
- `tools/list` 键恰为 `{tools}`、`tools/call` 键恰为 `{content,isError}` 不变。

### 书面边界（必须写进文档）
- **进程内无法中断**已阻塞的 `VBControls.Add` ⇒ 修法只有**预防**（预校验）与**文档化绕过**（已挂死时：终止并重启该 IDE 实例）；
- ★ 「让 OCX 成为控件部件」的正规途径，在**现有工具面下已证不可行**（只读探路）：
  - 菜单项**可触发**（`command_bar=Toolbox`、**command id 746**，但 `command_bars execute` **只能按题注**匹配 —— 该工具没有 id 参数）；
  - 触发后打开 `#32770 «Components»`；其**列表项不通过 `WM_GETTEXT` 暴露**（`ui_text` 只读到外框元素 ⇒ 读不到 ListBox 项）；
  - **缺三个原语**：**对话框列表项枚举/选择**（`LB_GETTEXT`/`LB_SETCURSEL` 或 MSAA/UIA）· **按 hwnd 点按或发送按键**（`control_window` 只认与语言无关的类型名，标题与 `#32770` 均被拒）· **合成鼠标**（连已知坐标也点不了）；
  - **驱动它会锁死 IDE**，唯一脱困 = **终止该 IDE 实例**（已演练；只终止自己的实例、不碰他人端口）。

### 版本关系
v1.3.2（响应侧，无需重冻）与 v1.4.0（+1 工具，需重冻）**可同批也可分两批**；若分两批，先 v1.3.2 再 v1.4.0。
## v1.5.0 消息型窗口原语 + 写入回执的分块/相位字段 + 可选代码窗滚轮

> 本版是**纯增量**：**不改**任何既有工具的 inputSchema、不删/不改任何既有回执键的语义。

### 1. 新增三个工具（106 → 109，**必须显式重冻**）

| 工具 | 权限 | 边界 |
|---|---|---|
| `read_active_dialogs` | **只读** | 枚举**本进程**可见的顶层对话框（默认 class `#32770`；`cls:"*"` = 本进程所有可见顶层窗），返回 `{hwnd,cls,title,rect,owner,ownerTitle,ownerDisabled,foreground,childCount,buttons[],controls[]}`。★`ownerDisabled=true` 才是「这个框**挡住了**宿主（模态）」的可靠信号 —— 「有窗口在显示」本身不是。**只枚举 + 读文本，不发任何消息** |
| `close_active_dialogs` | 变更 | **只允许显式调用**（产品内任何路径都不会自动调用它）。只对**安全白名单**内的按钮标题发 **BM_CLICK**（OK/Yes/Cancel/Close/Continue/Ignore + 中文等价），**危险黑名单优先**（delete/remove/uninstall/format/overwrite/discard/reset/清空/删除/移除/卸载/覆盖/丢弃/重置/格式化）；禁用或非白名单按钮**只报告 skipped，绝不点击**；`dry_run=true` 只规划。回执含 `clicked[]`/`skipped[]`/`remaining[]`，**如实回报是否真的关掉** |
| `send_window_message` | 变更 | 向**本进程**的 hwnd 发窗口消息：`action=click`（BM_CLICK）/ `action=key`（WM_KEYDOWN+WM_KEYUP）/ `action=message`（具名消息或数字）。默认 SendMessage，`post=true` 用 PostMessage。**只发消息**：不写内存、不注入、不改 IAT；**跨进程 hwnd 一律结构化拒绝**（`reason:"not_this_process"`） |

`send_window_message` 的具名消息（可扩）：`click`/`bm_click`、`wm_command`、`wm_keydown`、`wm_keyup`、`wm_char`、`wm_close`、`wm_mousewheel`、`wm_vscroll`、`wm_syscommand`、`wm_null`、`bm_setstate`、`cb_showdropdown`、`em_linescroll`；`wparam`/`lparam` 接受十进制或 `0x` 十六进制。

### 2. `command_bars` 支持按 id（**不改 inputSchema**）

- 复用**既有** `control` 参数：`"#746"` / `"id:746"` / 裸数字 = 按 id；
- ★VB6 的 CommandBar **Id 不唯一**（仅 id=746 就覆盖 986 个节点）⇒ id 路径**只在恰好一个 Enabled 控件命中时执行**；0 个命中或 ≥2 个命中都**结构化拒绝**（后者回执带候选 `bar>caption` 列表，用 `bar=` 或题注再收窄）。**拒绝猜**是设计，不是缺陷；
- 回执新增 `byId` / `id` /（歧义时）`idCandidates`。

### 3. 写入回执：`hunks`（**纯增量字段**）

- 所有走最小写入的写工具回执在 `writeCalls`/`writeMs` 旁**新增整数 `hunks`**；
- **口径（不变式重新界定）**：
  - **既有单块路径 / 回退路径**：`writeCalls <= 2` **不变**，此时 `hunks = 1`；
  - **hunk 路径**：`writeCalls = 2 x hunks`（每个 hunk 至多 1 次 DeleteLines + 至多 1 次 InsertLines）；
  - **`hunks <= MAX_HUNKS = 16`**；超过上限（或任一 hunk 超过「最小跨度单块」规模）则**回退单块最小化**，回执以 `hunks = 1` 体现；
  - `writeMs` **窗口未变**（仍是该次写的最外层计时）。
- **观测（仅现象，未作机制定论）**：写入耗时对**块规模**呈超线性（历史实测 1/50/500/5,000/26,999 行 ⇒ 0/15/31/844/21,593 ms 量级）；该现象在 的分相 / 双因素矩阵判定前，**不得**写成「慢在 COM」或「慢在我们」的结论。

### 4. 写入回执：相位字段（**纯增量字段**）

- 报 `writeMs` 的写工具回执新增：`diffMs` / `joinMs` / `delMs` / `insMs` / `maxBlockLines` / `phaseNote`；
- **语义（`phaseNote` 逐字写在回执里）**：`diffMs` = hunk 生成 + 校验；`joinMs` = 我方其余准备（切行、公共前后缀扫描、块数组构建、Join）；`delMs`/`insMs` = **在 CodeModule 调用内部**的时间；`maxBlockLines` = 单次 CodeModule 调用交给引擎的最大行数；
- ★这些是**一次调用的区间测量**，**不构成因果结论**；归属判定属 ，其真机矩阵结论落定前本节**不写机制**。

### 5. 可选代码窗滚轮—— **不是工具、不改 schema**

- **默认关闭**。开关：环境变量 `VB6MCP_WHEEL=1`（`true/on/yes` 亦可）或注册表 `HKCU\Software\VB6MCP\Wheel`；`VB6MCP_WHEEL_LINES`（1..50，默认 3）调每格行数；
- 实现 = **子类化本进程的代码窗窗口**（class `VbaWindow`，`SetWindowLongW(GWLP_WNDPROC)` + `CallWindowProcW`），把 `WM_MOUSEWHEEL` 翻译成 `WM_VSCROLL`；★**绝不安装键盘钩子、绝不安装全局钩子（含 WH_MOUSE）**；
- 行为：仅当**指针位于该代码窗矩形内**时生效（`Ctrl` = 翻页，`Shift` = 3 倍行数，不改光标与选择）；窗体设计器 / 其它窗格不受影响；
- 卸载：`OnDisconnection` 恢复原窗口过程（**仅当该窗口当前过程仍是我们的** —— 若第三方插件在我们之后子类化，则**保留链条、只停用**，并在回执里如实记 deferred）；**同一 hwnd 重复扫描不会堆叠**；
- 详细开关 / 行为 / 与第三方插件共存说明见 `内部已知限制与 FAQ` §34。
- ★**更正（终修，2026-10-09；上一条的「翻译成 `WM_VSCROLL`」已过期，**追加**更正、不改 schema）**：实测 `WM_VSCROLL` 对准**真代码窗**的 `SB_LINEDOWN×2 / SB_PAGEDOWN / SB_LINEUP×2` **5/5 不动**；现改为**在进程内直接设该窗格的 `CodePane.TopLine`**（读回值进回执），并且**只子类化标题能映射到 `CodePane.Window.Caption` 的窗口**（立即 / 监视 / 本地窗口同属 `VbaWindow` 类，按标题过滤并列入 `ide_state.wheel.skipped[]`）。on 臂 / 默认关 / 多实例 / 卸载退出码 0 的真机证据见 FAQ §34.6。

### 6. 本版不变式

- `tools/list` 结果键**恰为** `{tools}`；`tools/call` 顶层键**恰为** `{content,isError}`；
- **既有 106 个工具的 inputSchema 零改动**；本版**只新增**三个工具；
- 门禁四件套：**pre-freeze 必须能看到** `ADDED=read_active_dialogs,close_active_dialogs,send_window_message`，`--freeze` 后 `schema_changes=0`、`ORDER_MATCH=True`、`strict_failures=0`、`check_hub_alignment` 0 失败（hub 快照同步重生成）。

### v1.5.0 非 schema 增量（2026-10-09）：工程目录写入边界 + 写入清单 `projectDirWrites`

- **不改任何 inputSchema、不加参数、不改任何 description** ⇒ `tools/list` 仍 **109**、`schema_changes=0`、**无需重冻**；
- 三个会写工程目录的写工具的**回执恒带 `projectDirWrites`**：本次调用**实际写到磁盘的文件绝对路径数组**（写序、去重）；**没有写任何文件时是空数组 `[]`（字段恒在）**。★**限定（独立复测指出，补正）**：「恒在」只对**非拒绝回执**成立 —— **结构化拒绝**（`GuardReject` / 预校验路径）由**另一条**回执构建器产生，**不带** `projectDirWrites`。
 口径：`import_component` = 工程局部副本（契约）；`add_component` = 仅 `type:"form"` 走表单回退时非空，其余一律 `[]`；`save_project` = 本次写入的工程文件 + 组件文件
  （是「工程目录内」的**超集**：绑定在工程目录之外的组件文件也会列出 ⇒ 任何写入都不会藏在字段名后面）；
- **无工程目录（没有工程 / 工程从未存盘）⇒ 前置拒绝、零副作用**：`import_component` 与 `add_component {type:"form"}` 回
  `ok:false reason:"project_dir_unavailable"`（字段 `suggestedAction` / `detail` / `sideEffects:"none"`）；`save_project` 仍按其正常职责工作并把 `projectDirWrites` 置 `[]`；
- ★**旧的两个回退已删除**：`FreeImportPath` 的 `App.Path` 与表单回退的 `%TEMP%`/`C:`。旧行为会把导入产物写成 **`<产品目录>\Module1.bas`**（物证 `内部工作区`；本卡改前实测复现）
  或 `%TEMP%\Fm1.frm`（本卡改前实测），即写到**工程之外**；
- 环境开关 `VB6MCP_ALLOW_PROJECTDIR_WRITES`（**不是工具参数**）：`0`/`false`/`no`/`off` ⇒ 表单回退回 `ok:false reason:"project_dir_write_not_allowed"`；
  未设或 `1`/`true`/`yes`/`on` ⇒ 允许（`add_component` 的显式 `type:"form"` 请求即 opt-in）；
- 详述与改前/改后实测：`内部已知限制与 FAQ` §37。

### 非 schema 说明（2026-10-09）：F4 勘误 / F6 不存在的工具 / 写入回执字段覆盖清单

- **F4 已勘误**（见上文 **v1.1.0 勘误 3**）：`list_breakpoints` 只枚举 **`set_breakpoint` 注入物**（`' @mcp-break` 标记行）；`native_breakpoint` 不写源码标记 ⇒ **不会**出现；IDE 原生断点一律 `native:null`。
- **F6 / `read_monitor` 不存在**：109 工具面**没有** `read_monitor`。查 IDE **Message Monitor** 的现有手段只有 `show_message_monitor`（**打开**窗口）+ `find_gui` / `ui_text`（读窗口文本）+ `read_active_dialogs`/`close_active_dialogs`/`send_window_message`（模态框原语）。
  ⇒ **能开窗、能读文本，但没有「结构化条目 + severity 过滤 + 条数上限」的读取接口**。研究草稿 `内部工作区`（含建议 description/inputSchema/验收）**留档、未实现**（新增工具须 109→110 全量重冻；S 裁定暂不做）。详见 FAQ §39。
- **写入回执字段覆盖清单**（`writeMs` / `writeCalls` / `hunks` / 相位字段；语义见本版 §3、§4 与 FAQ §35。★本表只是把 的覆盖范围**登记**进契约，**不改任何 schema**）：

| 工具 / 分支 | `writeMs` | `writeCalls` | `hunks` | 相位字段（`diffMs/joinMs/delMs/insMs/maxBlockLines/phaseNote`） |
|---|---|---|---|---|
| `replace_code` op=replace（最小写入） | ✅ | ✅ | ✅（单块/回退 = 1；hunk 路径 = 2×hunks） | ✅ |
| `replace_code` op=set 的**退化路径** | ✅ | ✅ | ✅（=1） | ✅ |
| `replace_code` op=insert / append / delete | ✅ | ✅（0 或 1） | ✅（=1） | ❌（该分支无相位字段） |
| `add_code` | ✅ | ✅（=1） | ✅（=1） | ❌ |
| `edit_by_anchor` / `apply_edits` | ✅ | ✅ | ✅ | ✅ |
| `set_breakpoint` / `clear_breakpoints` / 其余注入类写路径 | ✅ | ✅ | ✅ | ✅ |
| `preview_edit`（**只预览、不写**） | ❌ | ❌ | ✅（与 `diff` 同源） | ❌ |

★口径以**回执实际字段**为准：写路径回执**恒带** `hunks`（最小写入路径），`preview_edit` 只回 `hunks`。

## 不变式（必须保持）
- `tools/call` 顶层结果键**恰为** `{content,isError}`；`tools/list` 结果键**恰为** `{tools}`；工具数与顺序不变（105）；既有字段语义**零变化**。
- `schema_changes=0`、`check_live_tools ORDER_MATCH=True`、`mcp_strict_audit strict_failures=0`、`check_hub_alignment` 0 失败。
- **同一工程双开必须被拒**并在 `ide_list` 里标红（不得静默）。

## 不变式

1. 既有 96 个工具的**参数与语义不变**；本版只增不改。
2. 新工具一律**只读**，不得有副作用。
3. 契约变更必须走 `contract` 广播（主卡 + 确认卡 + 版本递增），并**显式重冻** `tools/schema_baseline.json`；
   重冻前 `check_schema_baseline.py` 必须先报出 `schema_changes≠0`（证明变更可见），重冻后才回到 0。
**v1.5.0 勘误（2026-10-09，/ F7）**：§A 批 `preview_edit` 一行原写「**本批不改算法**：diff 仍是**前后缀裁剪**」——与实现不符：`preview_edit` 早已走 `EditTxn.DiffPlan` 的**逐行 LCS**（行级 LCS + hunk 分组 + `contextLines` 默认 3），且 `diff` 与 `hunks` **同源**（contract invariant 8），回执另有 `algorithm`（`lcs` / `fallback-too-large` / `none`）与可选 `algorithmNote`。已把该行改成与实际一致；属**同一版本内的事实勘误**，不递增版本、不改任何 inputSchema/返回字段名或工具行为。工具 description 早已写明 LINE-LEVEL LCS（`CodeTools.cls` 源码注释同），因此**本次不涉及任何描述变更、无需重冻**（`schema_changes=0`）。

**v1.1.0 勘误（2026-09-22，真机实测）**：§1 `ide_state` 的 `ebMode` 原写 `0|2`，实测 **run 态 = 1**（design=0 / run=1 / break=2，原样透传 `Vba6Probe.IdeEbMode()`，未做映射）。已在 §1 更正为 `0|1|2`；属**同一版本内的事实勘误**，不递增版本、不改任何工具的参数名或返回字段名。`内部已知限制与 FAQ` 里「EbMode 只返回 0 或 2」的旧说法同步勘误。

## v1.6.0 工具面扩容：`replace_in_files` + `undo_edit`（**109 → 111，必须显式重冻**）

> 授权：用户裁定（2026-10-09）「可以扩展工具面」。本版**同一次重冻**落地两个新工具与 `send_window_message` 的**描述补全 + 高危消息护栏**（后者新增参数 `confirm` ⇒ 本版 `CHANGED=send_window_message`，如实记录）。

### 1. 新增工具 `replace_in_files`（跨组件替换；消 S4 卡点）

- **入参**：`path`（**必须是当前打开工程的 .vbp 或其目录**）· `query` · `replacement` · `regex`（默认 false）· ★`dry_run`（**默认 true**）· `expected_shas[]`（逐组件守卫 `[{file, sha256}]`，`file` 可写**组件名**或**组件文件路径**）· `validate`（off|warn|strict，默认 warn）· `max_files`（默认 200，上限 1000）
- **语义**：逐组件读 **pane 文本** → 字面/正则替换 →（`dry_run=false`）经**既有最小写入通道** `SetCompTextMinimal` 写入（单组件原子：hunk 化、`writeCalls`/`hunks` 语义与既有编辑工具一致）
- **回执**：`{ok, dryRun, regex, files:[{comp, file?, hits, firstHitLine, changed, skipped, reason?, shaBefore, shaAfter, lintErrors, lintWarnings, writeCalls?, hunks?}], summary:{files, filesChanged, wouldChange, totalHits, applied[], stoppedAt?, rolledBack?, note?}, totalMs}`
- ★**失败语义（诚实，不假称原子）**：任一组件失败（写入失败 / 写后 pane 校验失败 / `validate=strict` 拒绝）⇒ **停止**；对**已应用**组件用运行前文本**回滚**（`rolledBack:true`）；回滚没能完成则 `rolledBack:false` 并**如实列出** `applied[]`；此时 `ok:false` + `reason:"write_stopped"`
- **安全**：`dry_run` 默认 true；`expected_shas` 不匹配 ⇒ 该组件**跳过**（`skipped:true, reason:"expected_sha_mismatch"`）**绝不写**；`path` 不是当前工程 ⇒ 结构化拒绝 `not_the_open_project`；**只写 pane，不碰磁盘**（盘上文件由 `save_project` 或 IDE 保存落盘）

### 2. 新增工具 `undo_edit`（组件级撤销；消 S10 卡点）

- **入参**：`component`（缺省 = 最近一次快照的组件）· `action`（`list`|`restore`，默认 restore）· `ts`（精确快照时间戳）· `to_disk`（默认 false；**true 目前结构化拒绝** `to_disk_not_supported` ⇒ 落盘请用 `save_project`）
- **快照**：**每次组件文本写入前**把**写入前的 pane 文本**以 **UTF-8** 存到 `<工程>\_undo\<component>.<yyyymmdd-hhnnss-ttttt>.txt`；与其它工程目录写入**共用门禁** `VB6MCP_ALLOW_PROJECTDIR_WRITES`（关掉则不写快照，`undo_edit` 如实报 `no_snapshot`）
- **回执**：`{ok, action, component?, ts?, shaBefore, shaAfter, snapshotSha256, restored, diskUntouched:true, writeCalls, hunks, snapshots?[]}`；无快照 ⇒ `{ok:false, reason:"no_snapshot"}`；`ts` 不存在 ⇒ `{ok:false, reason:"snapshot_not_found", available:[...]}`
- ★**显式单步、绝不自动回滚、默认不碰盘上文件**。

### 3. `send_window_message`：描述补全 + 高危消息护栏（本版 `CHANGED`）

- 入参名是 **`message`**（**不是 `msg`**）：写成 `msg` ⇒ 参数为空 ⇒ `reason:"bad_message"`（的实测症状，PLAN §7 已 A/B/C/D 四项对照）；`message` 同时收**名字**与**原始数字**（`0x0111` / `522` / `&H20A`）
- ★**新增 `confirm`（boolean，默认 false）**：`WM_CLOSE(0x0010)` / `WM_QUIT(0x0012)` / `WM_SYSCOMMAND(0x0112)` 默认**拒绝**（`ok:false, reason:"high_risk_message", confirmRequired:true`，**什么都不发**）；仅 `confirm=true` 才发（回执带 `highRiskConfirmed:true`）。其余消息行为**不变**。

### 4. 随本版一并登记（**不改工具面**）

- `add_component{type:"form"}`：窗体**改由 IDE 自身**创建（见 FAQ §41）；回执新增 `createdVia`（`ide_dialog`|`synthesized`）与兜底时的 `ideRouteNote`；IDE 路线 `projectDirWrites` 为 `[]`，耗时约 **3.5–4.9 s**
- `control_execution.mode` 的 `enum` 已在 v1.5.0（描述级重冻）落到实现真正接受的 **12 个值**（含 `resume`/`stop`/`stepinto`/`stepover`/`stepout`/`runtocursor`/`fullcompile` 兼容别名）—— 本版仅**登记**，不改 schema
- ★**工具面局限（登记）**：① `command_bars{action:"list"}` **不含 Add-Ins 弹层**（这些菜单项只能 execute-by-caption 或按 id 触达）；② `AddFromTemplate` 的**窗体模板路线在本机不可用**（本仓库的模板目录与 VB6 `Template\Forms` 都没有普通窗体模板 `.frm`），且它会把组件绑定到模板文件本身（同类隐患）⇒ **不实现**（S 裁定）

### 5. 不变式与冻结（本版复核）

- `tools/list` 结果键**恰为** `{tools}` 且 **111**；`tools/call` 结果键**恰为** `{content,isError}`
- 既有 109 工具的回执形状**不变**（新增键只出现在两个新工具，以及 `add_component` 的 `createdVia`/`ideRouteNote` 纯增量）；`writeCalls ≤2`（单块路径）与 `hunks` 语义**不变**
- 冻结：`docs/tools_schema.json` 与 `tools/schema_baseline.json` = **111**；hub 快照 = 111 + 4 hub-only = **115**；`hub_alignment_failures=0`
- pre-freeze 实测：`ADDED=replace_in_files,undo_edit` / `CHANGED=send_window_message`（`REMOVED=none`）—— 与卡面预期 `CHANGED=none` 的差异源于**同一次重冻并入了第三项的 `confirm` 参数与描述补全**，如实记录

### 6. v1.6.0 之后的非 schema 增量（构建 4F19D1A5…）

**`_undo` 快照并入 `projectDirWrites`（可枚举写入一致性）** —— 不改工具面、不改 inputSchema、**不重冻**：

- **落盘位置与命名**：`<工程目录>\_undo\<组件>.<yyyymmdd-hhnnss-ttttt>.txt`（**UTF-8**；`<组件>` 经文件名安全化，非字母数字字符替换为 `_`）。快照内容是**写入前的 pane 文本**。
- **何时写**：任何**经单块/多 hunk 文本写入通道**（`SetCompTextMinimal`）的写工具，在**改动 pane 之前**写一份；与其它工程目录写入**共用同一把门禁** `VB6MCP_ALLOW_PROJECTDIR_WRITES`（`0/false/no/off` ⇒ 不写、也不列出）。
- **可枚举性（本增量）**：产生快照的那一次调用的回执，其 `projectDirWrites` **列出本次实际写出的快照文件绝对路径**（写序、去重、**仅在文件确实落盘时列出**）。覆盖：`replace_code` / `apply_edits` / `edit_by_anchor` / `inject_trace` / `remove_trace` / `fault_inject` / `clear_faults` / `set_breakpoint` / `clear_breakpoints` / `replace_in_files` / `undo_edit`（以及本来就带该字段的 `add_component` / `import_component` / `save_project`）。**没有写入时是空数组 `[]`，字段恒在。**
- **与「不碰盘上源码」的关系**：快照文件在**我们自己的快照区**，**不改动任何 `.bas/.cls/.frm`**；`undo_edit` 默认只恢复 **pane**（`diskUntouched:true`），要落盘仍须 `save_project`。
- **实现口径**：该字段是**本次顶层调用**的写入清单 —— 清单在**唯一工具漏斗 `CallTool`** 里按调用重置（避免跨调用继承，与 同一纪律）。

---

### 7. v1.6.0 之后的非 schema 增量（构建 12F8D5FE…）

**回执可观测性与开销可见性（F1 诊断上限 / F2 回退标注 / F3 回执尾部计时 / F4 add_code 区间）** —— 不改工具面、不改 inputSchema、**不重冻**（`schema_changes=0`、冻结件 sha `9CEA9D1F…` 不变）：

- **F1 诊断上限（显式截断，信息不静默丢失）**：所有带 lint 的回执，其 `diagnostics` 数组**上限 300 条**（可用环境变量 `VB6MCP_LINT_MAX_DIAG` 改，形状参照既有 `VB6MCP_LINT_REGION_THRESHOLD`），并**恒带**三个新键：
  - `diagnosticsTotal` = 本次 lint 的**全量**诊断条数；
  - `lintTruncated` = `diagnostics` 是否被截断（`true` 时数组短于全量）；
  - `lintDiagnosticCap` = 本次生效的上限值。
  `lintSummary{errors,warnings}` **仍统计全部诊断** ⇒ `validate:"strict"` 的拒绝判据（读 `lintSummary("errors")`）**不被上限削弱**。真机回归：27,005 条 error 的夹具上 `validate:"strict"` 仍以 `ok:false, reason:"lint_errors", lintErrors=27004…27005` 拒绝、**磁盘逐字节不变**，同时回执 `diagnostics` 只有 300 条、`lintTruncated=true`。
- **F2 整模块回退显式化**：调用方给了 `targetLine>0`（要求区间）但**没有任何过程包含该行**时，回执现在标 `lintScope:"whole"` + `lintFallback:"no_procedure_contains_line"`（此前**完全静默**，与区间通过的形状无法区分）。`lintNote` 的**取值不变**；`targetLine=0`（小模块 / 工具本就没有单一目标行）时仍**不出现** `lintScope`（与 v1.5 逐位一致）。
- **F3 回执尾部计时（新增两键，`totalMs` 语义不变）**：
  - `serializeMs` = 本调用**全部序列化轮次**的累计耗时（序列化 + 注入 `mcp/ide` 戳 + UTF-8 字节计数）。★诚实边界：**产生返回串的那一次无法把自己写进自己**，故该值**少算最后一轮**；
  - `receiptBytes` = **客户端收到的 `content[0].text` 的 UTF-8 字节数**（含 dispatcher 注入的 `mcp/ide` 戳）。**可外部证伪**：实测每一例都满足 `len(content[0].text.encode('utf-8')) == receiptBytes`（A/B/C/S + `edit_by_anchor` + `apply_edits`）。
  - ★**传输时间不可测**：JSON-RPC 信封、HTTP/socket、hub 转发、客户端解析全部在 add-in **之外**，`totalMs` / `serializeMs` / `receiptBytes` 都**不覆盖**它们。
  - 适用范围：**所有已经带 `totalMs` 的回执根**（`replace_code` / `add_code` / `ChangeResult2` 全家 / `set_breakpoint`·`clear_breakpoints` / `replace_in_files` / `verify_scope`）。**写拒绝回执**（`StrictRefusal`，本就无 `totalMs`）**不带**这两键 —— 这是有意的口径边界。
- **F4 `add_code` 已满足（无源码改动）**：`add_code` **自 起**就把 `targetLine` 传为**追加行的物理行号**（`LintPackForEdit(predicted,kind,cm.CountOfLines,LineCount(cur)+1)`），因此大模块上**不**是整模块 lint。真机（12F8D5FE，27,000 行**过程化**模块，自有夹具）：`add_code` 回执 `lintScope:"region"`、`lintMs=172`、`diagnosticsTotal=0`、**磁盘未变**。第 4 条「`add_code` 永远 `targetLine=0`」系**源码误读**：其**效果**在当轮的**无过程**夹具上确实成立（无可归属过程 ⇒ 回退整模块），但**机制不是** `targetLine=0`。
- **实测开销（同一夹具、同一调用：A = 27,005 行的「模块级编辑」场景）**：墙钟 **11,244 → 1,019 ms**；`content[0].text` **2,690,816 → 30,872 B**（**87×**）；`lintMs` **3,812 → 672 ms**（差的 3,140 ms 是**上限之前**为 27,005 条诊断逐个 `CreateObject` 的开销，被 F1 一并消除）；`diagnosticsTotal=27,005`、`lintTruncated=true`、`lintScope=whole`、`lintFallback=no_procedure_contains_line`；`lintSummary` 与改前**逐位相同**。键集合差异 = **纯新增**（A 场景 +7 键、B/C/S +5 键，`REMOVED=none`）。
---

### 8. v1.6.0 之后的**参数增量**（构建 `4F661FE1…`）：opt-in `snapshot` + `undo_edit` 结构化

> 本卡**触 `inputSchema`**（三个最热编辑工具各新增一个**可选**参数）⇒ **已显式重冻**（`python 维护者的契约导出脚本（未随源发布） --freeze`）。
> 冻结件 sha：**`9CEA9D1F9A699703…` → `1EE3BFA9E917AB6E…`**（`tools/schema_baseline.json` == `docs/tools_schema.json`）。
> pre-freeze 实测：`ADDED=none` / `REMOVED=none` / `CHANGED=apply_edits,edit_by_anchor,replace_code`（`schema_changes=3`）—— ★脚本按**字母序**输出 CHANGED，**集合与卡面预期完全相同**。
> post：`schema_changes=0`、`ORDER_MATCH=True`、`live_tool_changes=0`、`strict_failures=0`、`tools/list` 键恰 `{tools}` 且 **111**（新增的是**参数**不是工具 ⇒ 数量不变）、hub face **111 + 4 hub-only = 115**、`hub_alignment=0`。
> ★**本卡未改版本号**（仍 `v1.6.0`，`Util.bas` 常量未动）：卡面只授权「一次显式重冻」⇒ 如实登记。★版本号已由后续版本补齐为 `v1.7.0`（见 **§9**）。

#### 8.1 新参数 `snapshot`（boolean，默认 `false` = off）

| 工具 / 分支 | `snapshot=false`（缺省） | `snapshot=true` |
|---|---|---|
| `replace_code` **op=replace/insert/append/delete** | **不写快照**（与 v1.6.0 逐字节一致） | 写入**写前 pane 文本**快照，`undo_edit` 可撤销该次编辑 |
| `replace_code` **op=set** | 走写通道 ⇒ **默认已写**一份（不变） | 同一份（**不会写第二份**） |
| `apply_edits` | 走写通道 ⇒ **默认已写**一份（不变） | 同一份（**不会写第二份**） |
| `edit_by_anchor` | 走写通道 ⇒ **默认已写**一份（不变） | 同一份（**不会写第二份**） |

- **口径**：`snapshot=true` = 「**保证**本次调用在写之前落下**恰好一份**写前快照」。对**本来就默认写快照**的工具是**幂等**的：快照在工具内先取，写通道被显式告知「已取」（`SetCompTextMinimal` 的 `snapshotAlreadyTaken`），因此**不会**出现第二份同名语义的文件。
- **取值**：JSON 布尔 `true` 为准；为兼容卡面/FAQ 里的 `snapshot=on`，**字符串 `"on"` / `"yes"` / `"true"` / `"1"`（大小写不敏感）同样算开**；其余（含缺省）一律**关**。
- **落盘**：`<工程>\_undo\<组件>.<yyyymmdd-hhnnss-ttttt>.txt`（UTF-8），并在**同一次调用的回执** `projectDirWrites` 里列出；与其它工程目录写入**共用门禁** `VB6MCP_ALLOW_PROJECTDIR_WRITES`（`0/false/no/off` ⇒ **即使 `snapshot=true` 也不写、回执为 `[]`**）。
- **零回归**：缺省（off）时**不新增任何文件、不新增任何回执键**；`writeCalls`/`hunks`/相位字段/`projectDirWrites` 的恒在性与取值均不变（见 §8.3 的 A/B）。★**限定**：`projectDirWrites` 的「恒在」只对**非拒绝回执**成立，**结构化拒绝回执不带**该字段（见 §6 的补正）。
- **拒绝路径不写快照**：所有守卫拒绝（`range_beyond_end` / `inside_continuation` / `validate:"strict"` / sha 守卫）都发生在快照点**之前**，被拒调用**不产生**快照文件。

#### 8.2 `undo_edit` 的结构化「无快照」

| 场景 | 旧 | 新 |
|---|---|---|
| **点名组件**且 `_undo` 目录不存在 / 该组件没有任何快照（`list` 或 `restore`） | `reason:"no_snapshot"` | **`reason:"no_snapshot_for_component"`** + `component` + `hint` |
| **未点名组件**（全局）无快照 | `reason:"no_snapshot"` | **不变**（`no_snapshot`） |
| 点名组件**有**快照但 `ts` 不存在 | `reason:"snapshot_not_found"` + `available[]` | **不变** |

`hint` 是固定一句英文：`pass snapshot=true ("on"/"yes"/1 also accepted) to replace_code / apply_edits / edit_by_anchor to write a pre-edit pane snapshot; then undo_edit can restore it. set, apply_edits and edit_by_anchor already snapshot by default; replace_code's replace/insert/append/delete do NOT unless snapshot=true`。任何「无快照」路径都**不**再落到 `unexpected_error`。

#### 8.3 实测（构建 `4F661FE1`，自有夹具 `内部工作区`，干净 IDE `list_addins=21`）

- **默认行为零回归（A/B，跨构建）**：把归档的改前源码（`内部工作区`）重新编译成**功能等价**的改前构建，跑**同一段共享调用序列**：**6 个采样点的 pane 文本逐字节相同**、8 个共享步骤的**回执键集合相同**（唯二差异是 `undo_edit` 无快照的 S4/S5：`+{component,hint}`，即本卡预期）、`writeCalls`/`hunks`/`projectDirWrites` 计数相同、`_undo` 文件数相同（3）。`ab_differences=0`。
- **opt-in 生效**：`replace_code op=replace` + `snapshot=true` ⇒ `undo_edit restore` 后 pane **逐字节回到写前**（CR 归一化后相等，sha `e989698c…`）；快照文件内容与写前 pane **逐字节相同**。
- **恰好一份**：`replace_code(op=replace/set)`、`apply_edits`、`edit_by_anchor` 在 `snapshot=true` 下每次调用 `_undo` 文件数 **+1**（无重复）。
- **字符串开关**：`snapshot:"on"` 同样生效。
- **被拒调用不写**：`range_beyond_end` 的调用 `_undo` 文件数 **±0**。
- ★**构建不可字节复现（新事实）**：同一份源码连续 3 次 `build.bat` 得到 **3 个不同的 DLL sha**（`d7846AA2…`/`0CC374F9…`/`C508A4F6…`，size 均 1,921,024）。⇒ 「重建后核 `build.status` 的 sha 变化」这个门**每次构建都会通过**；**不能**用 sha 相等来论证「未重建」或「源码一致」，源码一致性必须看源码 sha（本卡 `CodeTools.cls` = `FC1EB2DA…`）。

---

### 9. v1.7.0：版本号补齐 + 迁移条目（additive，**免重冻**）

> 本版**不改工具面**：`tools/schema_baseline.json` == `docs/tools_schema.json` == **`1EE3BFA9E917AB6E5D148D409FB6B33ACA91EF2B20C7988E39ADE589536108BE`**（已重冻的那一份，**本卡不再重冻**）；工具数 **111** 不变。
> 变更内容 = **仅版本号与文档**：`Util.bas` 的 `VB6MCP_CONTRACT_VERSION` = `v1.7.0`（单源常量；`initialize.contractVersion` / 每次响应的 `mcp` 戳 / `mcp_info.contract` / instructions 首行都读它）+ hub 的 `contractVersion`（由本文件的版本记录表经 `tools/gen_hub_snapshot.py` 生成）。

#### 9.1 迁移条目（如实：additive）

- **新增参数**：`replace_code` / `apply_edits` / `edit_by_anchor` 各新增**可选** `snapshot`（`type=boolean`，**默认 off**）。
- **默认行为不变**：参数缺省时回执**不新增键、不新增文件**；跨构建 A/B（改前源码重建的功能等价构建 vs 交付构建）在**同一段共享调用序列**上 `ab_differences=0`、6 个 pane 采样点**逐字节相同**（证据：`内部证据{ab_compare.txt,a087_old.json,a087_new.json}`）。
- **新增语义**：`snapshot=true`（`"on"`/`"yes"`/`"1"` 亦认）**保证恰好一份写前 pane 快照**；`undo_edit` 在**点名组件且该组件无快照**时返回 `reason:"no_snapshot_for_component"` + `hint`（**全局**仍是 `no_snapshot`；**有**快照但 `ts` 不存在仍是 `snapshot_not_found` + `available[]`）。
- **落盘仍受门禁**：`VB6MCP_ALLOW_PROJECTDIR_WRITES=0/false/no/off` ⇒ 即使 `snapshot=true` 也不写快照、回执 `[]`。

#### 9.2 `mcp_info.breaking` 的处置（以实际语义为准）

- **本版不新增任何 breaking 条目**：新增的是**可选参数**，不破坏任何既有调用（默认 off ⇒ 行为不变，有 A/B 为证）⇒ `breaking[]` 里**不编造**条目。
- `breaking[]` **保留历史条目**（v1.2.0–v1.6.0）：删除它们等于篡改记录。
- v1.7.0 的说明放在 **`mcp_info.migrations`**（两条：参数语义 + 冻结件 sha）与 **`mcp_info.changelog`**（一条）里。★若 S 的原意是「`breaking` 字段整体置空」，那属**破坏历史记录**的另一种改动，请明示后再做。

#### 9.3 三处一致（实测口径）

| 位置 | 读法 | 期望 |
|---|---|---|
| `Util.bas` `VB6MCP_CONTRACT_VERSION` | 源码 | `v1.7.0` |
| add-in `initialize.contractVersion` / `mcp_info.contract` / 每次响应的 `mcp.contract` | 活体（**需重建后**） | `v1.7.0` |
| hub `bridge/hub_schema.generated.json` 的 `contractVersion` | `gen_hub_snapshot.py` 生成 | `v1.7.0` |



---

### 10. v1.7.0 之后的非 schema 增量：VB6 菜单里的 MCP 设置界面

**不改工具面**（无新工具、不改任何 `inputSchema`/描述）⇒ `tools/list` 仍 **111**、`schema_changes=0`、**不重冻**（冻结件 `1EE3BFA9…` 不变）。

#### 10.1 新增的 IDE 菜单项与设置窗体（**产品内非工具能力**）

- **VB6 IDE → Add-Ins → `VB6 MCP Settings...`**（add-in 注册的**第二个** Add-Ins 项；原「VB6 MCP Server」项不变）⇒ `SettingsForm.Show vbModeless`（**非模态**：表单开着时 MCP 继续应答，实测 `read_active_dialogs=0`）。
- 窗体 `SettingsForm` **由 IDE 自己生成**（`add_component{type:"form"}` + `form_designer`/`add_control`/`set_control_property`/`create_event_proc` + `replace_code{operation:"set"}` 写代码窗）——**未**手写 `.frm` 头部/布局块（`内部禁用路由说明` 禁令 2）；文案**纯 ASCII**。

#### 10.2 设置项与优先级

| 项 | HKCU 叶名（`HKCU\Software\VB6MCP`） | 默认 | 环境变量（优先） |
|---|---|---|---|
| 代码窗滚轮开关 | `Wheel` | off | `VB6MCP_WHEEL` |
| 每格滚动行数（1–50） | `WheelLines` | 3 | `VB6MCP_WHEEL_LINES` |
| SHA-256 强制纯 VB6 回退 | `ShaFallback` | off（CryptoAPI） | `VB6MCP_SHA_FALLBACK` |

- **优先级恒为 environment > HKCU\Software\VB6MCP > 内置默认**；表单对每项显示**有效值 + 来源**（`environment`/`HKCU`/`default` 三态）。`environment` 来源的项**不可**从表单改（表单**只写 HKCU**，不改 VB6 进程环境）。
- **只读状态区**：build sha256 · product 版本 · contract 版本 · **工具数**（`CodeTools.ToolCount()`，不硬编码）· 当前实例端口 · 注册表位置。

#### 10.3 即时生效与重启

- **三项都不需要重启 IDE**：`Apply` 写 HKCU 后**现场**调 `WinOps.WheelInit` / `WinOps.WheelShutdownJson`（滚轮子类化装上/卸下）；`WheelLines` 由 `WheelInit` 每次重读；`ShaFallback` 由 `Sha256.ShaForceFallback()` **按调用**读取。

#### 10.4 实现事实（边界）

- 注册表读写走 **advapi32**（`RegCreateKeyExA`/`RegSetValueExA`/`RegQueryValueExA`，包装在 `WinOps.RegGetString`/`RegSetString` + `McpSettingGet`/`McpSettingSet`）；★`WScript.Shell` 的 `RegRead`/`RegWrite` 在该 add-in 进程内**不可靠**（实测写不进/读回为空）。
- **不碰用户工程**；**无全局钩子**（未用 `SetWindowsHookEx`）；滚轮仍是窗口子类化。新增 `McpServer.ToolCount()`（供表单状态区）。
- 详细 FAQ：`内部已知限制与 FAQ` **§46**。

---

### 11. v1.7.0 之后的**描述级重冻**：`close_active_dialogs` 补破坏性正文护栏

> **本版如实定为 additive/changed（仅描述变更）**：只改 `close_active_dialogs` 的 **description**，
> **不改任何 inputSchema、不增删工具、不改任何回执字段或取值** ⇒ 工具数仍 **111**、`tools/call` 顶层键仍恰为 `{content,isError}`。
> ★**契约版本不递增**（仍 **v1.7.0**）：与描述级重冻（版本不递增，S 已验收）同一先例；
> 本卡授权只有「**一次显式重冻**」。若日后要求版本递增，按既定处置补一条
> `v1.7.1` 即可（一处常量 + 一行版本表）。
> ★**冻结件已更新（旧 → 新）**：`1EE3BFA9E917AB6E5D148D409FB6B33ACA91EF2B20C7988E39ADE589536108BE` →
> **`043B242862F88743C871C7274C4DF99E8E696598BC85314BC1B5C82C3B4B6B2B`**
> （`tools/schema_baseline.json` == `docs/tools_schema.json`，105,696 B）⇒
> ★**新冻结 sha 对应的契约版本 = v1.7.0**（同一个版本内的**描述级重冻**，不是版本递增）。
> pre-freeze 实测：`ADDED=none` / `REMOVED=none` / `CHANGED=close_active_dialogs`（`schema_changes=1`）
> —— ★CHANGED 集合**恰为**授权的那一件，**未抹平**。
> post 实测（同一构建的**活体 IDE**）：`schema_changes=0` · `ORDER_MATCH=True` · `live_tool_changes=0` ·
> `strict_failures=0` · `tools/list` 键恰 `{tools}` 且 **111** · hub face **111+4 = 115** · `hub_alignment_failures=0`。

#### 11.1 变更内容（`close_active_dialogs` 的 description：464 → 858 字符）

行为实现体在构建 `d4f12489…` 时就已落地；本卡把**当时只写在实现与契约正文里**的护栏
补进**工具描述**（此前调用方**看不到**）：

1. **点击前先读对话框正文** = 标题 + **每个子控件的文本**；
2. 正文命中**破坏性模式**（`remove`/`delete`/`uninstall`/`erase`/`purge`/`destroy`/`format`/`reset` + 中文孪生）
   ⇒ **默认拒绝点击**，结构化返回 `reason:"dangerous_dialog_body"` + `bodyHit` + `dialogTitle` + `advice`，
   ★**一个窗口消息都不发**（根上仅在触发时多 `refusedCount`/`bodyGuard`）；
3. **逃生口 = 显式 `button="No"`**：`No` **不在默认白名单**（「保存吗？」这类框点 No 会丢工作），
   但它**永不具破坏性** ⇒ 只有**精确题注** `No`/`否` 被**显式点名**时才点；
4. **`dry_run=true` 预演**该判定（规划、不发送任何东西）。

#### 11.2 与 v1.5.0 条目表的关系（**不篡改历史**）

v1.5.0 的工具表里 `close_active_dialogs` 那一行只写了「安全白名单 + 危险按钮黑名单」。
**按钮级判定至今仍然成立**，但它**不足以描述** 之后的**正文级**护栏
（真实事故：破坏性动词在 STATIC 正文里，按钮只有 `Yes`/`No`/`Help`）⇒ 以**本节**为准；
v1.5.0 原文**保留不动**（改写/删除历史条目等于篡改记录）。

#### 11.3 行为与回归（本卡**未改实现**）

护栏的真机 A/B（真实同形弹框：BEFORE 默认点 `&Yes` 且把 AddIns 注册项**整条删掉**；AFTER 默认**拒绝**、
显式 `button="No"` 可点、拒绝路径不改注册表；安全框 3/3 逐项不变）见 `内部证据`
⇒ **不重跑**（卡面判据 5 明示）。本卡的回归 = **活体 tools/list 描述回归**
（`内部证据`：只有这一个工具的 description 变化、其余 110 个工具逐字节不变、
**所有 inputSchema 不变**、顺序不变）。

#### 11.4 相关 FAQ

`内部已知限制与 FAQ` **§47**（护栏用法 + ★两条新事实：**VB6 按 AddIns 名降序枚举** ⇒
坏条目排在 `VB6MCP.Connect` **之后**会「**弹框先出、8766 起不来**」；该弹框的 `Yes` 会**整条删除**注册项、
且 `LoadBehavior` 是 **VB6 自己**置 0 的）。
### 12. v1.7.0 之后的非 schema 增量：失败语义迁移 + `preview_edit` 与 `apply_edits` 解析对齐

> 本节**不改 `inputSchema`、不改工具面** ⇒ 工具仍 **111**、契约版本仍 **v1.7.0**、冻结件 **`043b2428…`** 未变（免重冻）。
> 变更全部在**响应侧语义**与**解析语义**；实现证据见 `内部证据`。

#### 12.1 迁移条目（★调用方必须知情：这几条路径**不再**报 `unexpected_error`）

| # | 触发 | 旧回执（迁移前） | 新回执（现在） | 卡 |
|---|---|---|---|---|
| 1 | **组件名不存在**（`FindComponent` 未命中） | `isError:true`、`reason:"unexpected_error"`、`vbErrorCode:-2147221484` | `isError:false`、`reason:"component_not_found"`（**不带 `vbErrorCode`**）；附 `triedComponent` / `componentCount` / `availableComponents` / `nearMatches` / `trimmedMatch` / `whitespaceCollapsedMatch` / `lookupIsCaseInsensitive` / `caseDifferenceIsNotTheCause` | |
| 2 | `preview_edit` 锚点 0 命中 / 命中 >1 / `op` 不是 R/I/D / `edits` 为空 | `isError:true`、`reason:"unexpected_error"`、`vbErrorCode:-2147221420` | `isError:false`、`reason` = `anchor_not_found` / `anchor_not_unique` / `unknown_op` / `edits_required` | |
| 3 | `apply_edits` 锚点不解析 / `op` 非法 | 原本就是结构化回执（`ok:false` + `reason`） | **未变** | — |

- **影响面（实测口径）**：第 1 类由**共享 helper**（`FindComponent` 的 `+20` 出口）产生 ⇒ **37 个调用点 / 41 个工具**（`内部证据`）。
  ★卡面原写「37 调用点 / ~25 工具」，**实测修正为 41 个工具**（同一 helper 被 4 个工具在多条路径上复用）。
- ★**成功路径与其它既有拒绝路径零影响**：跨臂差分实测 **22/22 逐字节不变**；在本卡构建上**复测同样 22/22**（见 12.3）。
- ★**需要改判的调用方**：把 `reason:"unexpected_error"` + `vbErrorCode:-2147221484` 当「未知内部错误」重试/上报的脚本，改为按
  `reason:"component_not_found"` 分支；`availableComponents` / `nearMatches` 已给可用名与近似名。
  ★**不要**按大小写去排查 —— 该查找用 `LCase$` 比较，**大小写差异永远不触发**该拒绝（回执里 `lookupIsCaseInsensitive` / `caseDifferenceIsNotTheCause` 明说）。

#### 12.2 `preview_edit` 与 `apply_edits` **共用同一解析**（缺陷修复）

- **问题（同会话实测，`内部证据内部对照数据`）**：同一输入下两工具**方向相反**地分歧 ——
  per-op `match:"exact"`：preview **成功** vs apply **`anchor_not_found`**；per-op `occurrence:2`：preview **`anchor_not_unique`** vs apply **成功并写入**。
- **判定 = 缺陷（不是有意差异）**：`docs/tools_schema.json:1488`（apply_edits）与 `:1530`（preview_edit）的 description **都没有**声明解析规则不同；
  两者的 `edits` 形状**完全相同**（`:1496-1499` vs `:1538-1541`）；`CONTRACTS.md:110` 把 `match` 只声明给 `edit_by_anchor`/`apply_edits`；
  `:96` 把续行拒绝声明给**写类工具**；FAQ §48.4 明确把它记为**既有缺陷（未修）** ⇒ **判据 (4) 的逃生门不成立** ⇒ 修。
- **修法**：`ToolPreviewEdit` 不再调 `EditTxn.ApplyEdits`，改为与 `ToolApplyEdits` **同一套三步**：整批一次 `VbAnchor.BuildNormIndex` →
  逐 op `ResolveAnchor`（**认 per-op `match`/`occurrence`**，`scope` 仍传 `Nothing`）→ **同一条续行守卫**（`inside_continuation`）→ `EditTxn.ApplyLineOps`。
- **`preview_edit` 的 per-op honoured 集合 = `{op, anchor, text, match, occurrence}`**（与 `apply_edits` **完全相同**）：只有 per-op `scope` 仍被
  **任何工具**忽略，回执照旧用 `ignoredOpKeys` / `opKeysHonoured` / `ignoredOpKeysNote` 显式报告（不静默）。
- **失败回执形状未变**：仍走 `MCPREASON:` ⇒ `isError:false` + 结构化 `reason`（12.1 第 2 类），只是现在**与写入结论一致**。
- ★**唯一保留且已文档化的差异**：`preview_edit` 的**顶层** `match` 仍**不生效**（该工具 `inputSchema` 只有 `name`/`edits`，本卡**不新增参数** ⇒ 免重冻）：
  传了会回 `ignoredArgs:["match"]` + `ignoredArgsNote` 明说被忽略。

#### 12.3 回归与等价证明（本卡实测）

- **等价证明（语义 oracle 双证）**：同一输入分别给两工具（各自**独立的一模一样的夹具副本**，`sameInputSha:true` 自证）⇒ `verdictsAgree`；
  且把 preview 的 `diff` 在 **Python 里独立重放**（每条上下文/删除行都与旧文本逐行断言）得到的文本 == `apply_edits` **写后窗格文本**（`replayEqualsWrittenPane:true`）。
- **既有路径逐字节不变**：`EQ`（10）+ `EQF`（12）= **22/22**，跨臂（修复前构建 `C6DB23A0` vs 修复后构建）掩码后回执**逐字节相同**；
  归一化集合按纪律 36 用**同一构建双跑**实测逼出（`内部证据vol1104_*.json`）。
- **续行夹具**：卡面夹具**无唯一续行锚点**（`P/BigB.bas` 的 1078 条续行**全同行**；`X1/X2/A` 零续行）⇒ 本卡**自造** `内部工作区`（在 `MX1.bas` 末尾追加**唯一**的 `csthead000001 = 1 + _ / csttail000001`），
  真机复现「预览成功、写入拒绝」并修复。

#### 12.4 本卡不变式（实测）

| 项 | 值 |
|---|---|
| 工具面 / `tools/list` 顶层键 | **111** / 恰 `{tools}` |
| 契约版本 | **v1.7.0**（未递增） |
| `docs/tools_schema.json` == `tools/schema_baseline.json` | **`043b2428…`**（未变） |
| hub 面 / hub_alignment | **115** / **0** |
| `schema_changes` / `live_tool_changes` / `strict_failures` | **0** / **0** / **0** |
| 槽位 `VB6MCP.dll` sha | == `build.status` 的 declared sha |
---

### 13. v1.7.0 之后的 **schema 变更**（构建 `B764F07D…`）：`check_public_api.baseline` 声明为必填

> 本卡**触 `inputSchema`**（`check_public_api` 由「无 `required` 键」变为 `required:["baseline"]`）
> ⇒ **已显式重冻**（`python 维护者的契约导出脚本（未随源发布） --freeze`）。
> 冻结件 sha：**`043B242862F88743…` → `2D4503C0FC923645E06D8FDBECA1B25A70B5BC7413A5993C50DFE5915FAE3903`**
> （`tools/schema_baseline.json` == `docs/tools_schema.json`，105,696 → 105,753 B）。
> pre-freeze 实测（**先看到变化才准重冻**，见本文 §1 的流程）：`ADDED=none` / `REMOVED=none` /
> **`CHANGED=check_public_api`** / `schema_changes=1` ✅；post-freeze：`schema_changes=0` ✅。

#### 13.1 方向裁定（缺陷 D-1107-02；为什么是「声明必填」而不是「接受缺省」）

| 证据 | 位置 |
|---|---|
| 旧 schema 里 `check_public_api` 有 `baseline` 属性，但**整块没有 `required` 键** | `docs/tools_schema.json` 旧件（`043B2428`）**1841–1855**；新件 **1841–1859**（`required` 在 1855–1857） |
| 111 个工具里 **56 个**显式声明了 `required`，`check_public_api` 是漏网之一；对照 `read_file` **有** `required:["path"]` | `docs/tools_schema.json:614–630`（`read_file`）/ 全文 `required` 键共 56 处 |
| 源码：`Set pCa("baseline") = Prop(...)`，随后 **`SchemaWith(pCa)`**（**没有** `Req(...)`） | `CodeTools.cls` 旧件（`ffcb27ad`）**841–847**，尤其 **847** |
| 源码：工具描述写着 *against a **frozen spec***；实现**强制**非空 baseline | `CodeTools.cls` **846** / **12132–12133**（旧行号） |
| 分类器**已经**把该条件当作调用方输入校验：`reason:"required_argument_missing"` + `argument:"baseline"` | `CodeTools.cls` **1463–1464**（`ClassifiedRefusalJson`） |
| `CONTRACTS.md` / `内部已知限制与 FAQ` **都没有**任何文字承诺 `baseline` 可选（0 命中） | 文档检索 |
| `docs/TOOL_CATALOG.md:99` 描述同样是 *Compare the current component public API against a frozen spec* | 生成物 |

**裁定：声明必填。** 理由：① 该工具**唯一**功能就是「与冻结 spec 比对」，没有 spec 就没有比对对象；
② 「接受缺省 baseline」必须**发明一个成功语义**（空 baseline ⇒ 当前所有公开成员都被报成 `added[]`），
这正是本项目列为缺陷的**误导性成功**（同族 D-1107-04/05）；③ 该做法还要**删掉** 刚落地的
分类（`required_argument_missing`）；④ **行为不变**——缺 `baseline` 以前被拒、现在仍被拒，**只有声明**移动。

#### 13.2 与既有行为的差分（零行为变更）

- `check_public_api {name}` **仍**回 `ok:false, reason:"required_argument_missing", argument:"baseline"`
  （同前）；`check_public_api {name, baseline}` 仍成功（`violations:[]`）。
- schema 侧唯一差异 = `inputSchema.required = ["baseline"]`；`properties` 与**全部 `description` 未改**
  ⇒ `docs/TOOL_CATALOG.md` **逐字节不变**（`525A9809…`，以旧快照确定性重建证明）。

#### 13.3 回归：`schema.required` == 实现真正强制的集合（本卡实测）

- **本工具 PASS**：`schema.required=["baseline"]` == 真机强制 `["baseline"]`（`argument:"baseline"`）。
- 声明侧对照：`read_file`（声明 `path`；缺参 ⇒ `required_argument_missing/path`）✅；
  `replace_code`（声明 `operation`；缺参 ⇒ **被拒**，但归类为 `invalid_action`+ `allowedValues`）。
- ★**同类不符（本卡只列出、未修）**——真机缺参实测 `required_argument_missing` 而 schema 未声明：

  | 工具 | 未声明的必填参数 | 触发方式 |
  |---|---|---|
  | `app_manifest` | `path` | 无条件 |
  | `compare_fixture` | `path` | 无条件 |
  | `ext_install` | `path` | 无条件 |
  | `record_fixture` | `path` | 无条件 |
  | `vbg_group` | `path` | 无条件 |
  | `project_snapshot` | `path` | **按 action**（`list`/`create`） |
  | `add_reference` | `guid`（或 `path`） | **按 mode**（`guid`/`file`） |
  | `add_code` | `code` | **二选一**（`code or file`） |

- ★**静态候选的假阳性**同样登记（真机未强制）：`export_component` / `git` / `list_files` /
  `project_meta` / `search_in_files` / `verify_scope`（默认值或工程目录兜底 ⇒ 声明可选**正确**）；
  `import_lessons_md` 缺 `path` 时**确实被拒**，但回执只有 `error`、**没有 `reason`**（另一类，登记）。

#### 13.4 本案不变式（实测）

| 项 | 值 |
|---|---|
| 工具面 / `tools/list` 顶层键 | **111** / 恰 `{tools}` |
| 契约版本 | **v1.7.0**（未递增） |
| `docs/tools_schema.json` == `tools/schema_baseline.json` | **`2D4503C0…`**（新冻结） |
| pre-freeze / post-freeze `schema_changes` | **1（`CHANGED=check_public_api`）** / **0** |
| hub 面 / `hub_alignment` / `ORDER_MATCH` / `strict_failures` | **115** / **0** / **True** / **0** |
| 槽位 `VB6MCP.dll` sha | == `build.status` declared（`B764F07D…`） |
| `docs/TOOL_CATALOG.md` | **未变**（`525A9809…`） |

### 14. v1.7.0 之后的 **schema 变更**（构建 `E4A7C124…`）：`property_bag.has_value` 声明 + 9 条 required/条件声明（**两卡合并为一次显式重冻**）

> 两卡**都触 `inputSchema`**，按这两卡的授权**合并为一次显式重冻**
> （`python 维护者的契约导出脚本（未随源发布） --freeze`）。
> 冻结件 sha：**`2D4503C0FC923645E06D8FDBECA1B25A70B5BC7413A5993C50DFE5915FAE3903`
> → `6808B1DB7F2FE9546A20DFA4900F544CFDA6AFBE28ACB9007624018A434CF97D`**
> （`tools/schema_baseline.json` == `docs/tools_schema.json`，105,753 → 106,571 B；工具数 **111 → 111**）。
> pre-freeze 实测（**先看到变化才准重冻**，见本文 §1 的流程）：`ADDED=none` / `REMOVED=none` /
> **`CHANGED=add_code,add_reference,app_manifest,compare_fixture,ext_install,import_lessons_md,project_snapshot,property_bag,record_fixture,vbg_group`**
> / `schema_changes=10`（rc=1）；post-freeze：`schema_changes=0` ✅。
> `docs/TOOL_CATALOG.md` **逐字节未变**（`525A9809…`）—— 本版**没有**改动任何**工具级 description**。
> ★**契约版本号仍 `v1.7.0`**：两卡只授权「一次显式重冻」，**未**授权版本递增
> （与既定处置相同；若 S 要求递增，请派版本卡）。
> ★**级联（纪律 42）**：**此后所有卡的「冻结件不得变」一律以 `6808B1DB…` 为准**（旧 `2D4503C0…` 已结案）。

#### 14.1 方向裁定：**让文档化的 `value` 生效**（不是删 `value`，也不是只声明 `has_value`）

| 证据 | 位置 |
|---|---|
| schema 自己的 `value` 描述写着 *Value to write (**omit to read**)* ⇒ **presence 即写** | `docs/tools_schema.json`（旧件 `2D4503C0`）`property_bag.properties.value` |
| 实现真正决定读/写的是**未声明**的 `has_value` | `CodeTools.cls` 旧件 **2076**（dispatch：`JsonGetBool(args,"has_value",False)`） |
| `hasValue` 为真才 `WriteProperty`，否则 `ReadProperty` | `CodeTools.cls` **7895–7925**（`ToolPropertyBag`） |
| 另一处文档又写「写需 `has_value:true`」⇒ **两处文档互相矛盾**；只按 schema 传 `value` 的调用方**静默拿到一次 READ** | `内部交接说明:96` |

**裁定：保留 `value` 并让它真正生效（presence ⇒ write）**，**并且**把实现真正 honour 的 `has_value` 补进声明。理由：
① **只声明 `has_value` 不修缺陷** —— 按文档只传 `value` 的调用方**仍然**会静默读（这正是被报的缺陷，纪律 38）；
② `has_value` 已被实现 honour 且被 `内部交接说明` 文档化 ⇒ **不声明它**就留下**另一个**「声明≠实现」；
③ 两者语义不冲突：`value` 的**存在**只能**新增**一次写，`has_value` 用于写**空值**（`value:""` 亦可）。
**行为差分（真机实测）**：`{target:"component",name:"mb",key:"k1112a",value:"7"}` 旧 = 静默 READ（回 `value:""`）
→ 新 = `{wrote:true,component:"mb",key:"k1112a"}`，**随后回读 `value:"7"`** —— 文档化的输入不再被忽略。

#### 14.2 8 条「声明与实现不符」逐条复核（每条给 schema 位置 + 实现行号 + 实际强制行为）

| # | 工具.`键` | 旧 schema | 实现行号 | **实际强制行为（真机）** | 判定 | 本次改动 |
|---|---|---|---|---|---|---|
| 1 | `app_manifest.path` | 无 `required` | **6293** `If Len(p)=0 Then Err.Raise … "path is required (the .exe or .dll …)"` | **无条件** ⇒ `required_argument_missing/argument=path` | **应声明 required** | `required:["action"] → ["action","path"]` |
| 2 | `compare_fixture.path` | 无 `required` | dispatch **2222** `ResolvePath(JsonGetStr(args,"path"))`；`ResolvePath` **8729** 抛 `"path is required"` | **无条件** | **应声明** | `required:[] → ["path"]` |
| 3 | `ext_install.path` | 无 `required` | dispatch **2254** 同上 | **无条件** | **应声明** | `required:[] → ["path"]` |
| 4 | `record_fixture.path` | 无 `required` | dispatch **2220** 同上 | **无条件** | **应声明** | `required:[] → ["path"]` |
| 5 | `vbg_group.path` | 无 `required` | **6134** `"path is required (the .vbg project group)"` | **无条件**（在 `Select Case action` **之前**） | **应声明** | `required:["action"] → ["action","path"]` |
| 6 | `project_snapshot.path` | 无 `required` | **6163**（create）/ **6174**（list）要 `path`；**6169**（restore）要 `dir` 而**不要** `path` | **按 action** | **JSON-Schema 无法表达条件必填 ⇒ 由 property 描述声明条件**（真机两侧都断言） | `path`/`dir` 描述改写（`required` 保持 `["action"]`） |
| 7 | `add_reference.guid` | 无 `required` | **6597–6603**：`mode="file"` ⇒ 要 `path`；**否则**（默认 `guid`）⇒ 要 `guid` | **按 mode** | 同上 ⇒ 描述声明条件 | `guid`/`path` 描述改写 |
| 8 | `add_code` 的 `code`/`file` | `required:["component"]` | **4990** `Err.Raise … "code or file is required"`（**先查 component**） | **二选一** | 同上 ⇒ 描述声明条件（`ArgNameOf` 对含 `or` 的消息**故意不回 `argument`**） | `code`/`file` 描述改写 |
| 9 | `import_lessons_md.path`（判据 (3)，§13.3 已登记） | 无 `required` | dispatch **2236** 空 `path` ⇒ 旧回 `{ok:false,error}` **无 reason** | **无条件** | **应声明 + 补 reason** | `required:[] → ["path"]`；回执补 `reason/argument/message/suggestedAction/sideEffects`（**保留 `error`**） |

★**判定方法学**：结论**一律来自真机缺参**（§14.3），静态调用图只用于**生成候选**。
★**6 个静态候选已判假阳性、不再处理**（`export_component`/`git`/`list_files`/`project_meta`/`search_in_files`/`verify_scope` —— 有默认值或工程目录兜底，声明可选**正确**）。

#### 14.3 真机正向断言：**声明集合 == 实现强制集合**（`内部工作区*.jsonl`）

| 工具 | 真机缺参实测 | 与声明一致？ |
|---|---|---|
| `app_manifest` | `{action:"read"}` ⇒ `required_argument_missing/argument=path` | ✅ 新声明 `["action","path"]` |
| `compare_fixture` | `{}` ⇒ `required_argument_missing/argument=path` | ✅ `["path"]` |
| `ext_install` | `{}` ⇒ `required_argument_missing/argument=path` | ✅ `["path"]` |
| `record_fixture` | `{}` ⇒ `required_argument_missing/argument=path` | ✅ `["path"]` |
| `vbg_group` | `{action:"list"}` ⇒ `required_argument_missing/argument=path` | ✅ `["action","path"]` |
| `import_lessons_md` | `{}` ⇒ `required_argument_missing/argument=path` + `error` | ✅ `["path"]` |
| `project_snapshot` | create/list ⇒ `argument=path`；**restore ⇒ `argument=dir`**（**不是 path**） | ✅ 与描述一致（两侧都断言） |
| `add_reference` | 无 `mode` / `mode:"guid"` ⇒ `argument=guid`；`mode:"file"` ⇒ `argument=path` | ✅ 与描述一致（两侧都断言） |
| `add_code` | `{component:"mb"}` ⇒ `required_argument_missing`（**无 `argument`**，因为是二选一）；`{component:"mb",code:"'x"}` ⇒ **成功** | ✅ 与描述一致（两侧都断言） |
| `property_bag` | `{target,key,value}` ⇒ **写**；`{target,key}` ⇒ **读**；`{…,has_value:true}` ⇒ **写** | ✅ 声明 `value`/`has_value` 与实现同集合 |

#### 14.4 本案不变式（实测）

| 项 | 值 |
|---|---|
| 工具面 / `tools/list` 顶层键 | **111** / 恰 `{tools}` |
| 契约版本 | **v1.7.0**（未递增） |
| `docs/tools_schema.json` == `tools/schema_baseline.json` | **`6808B1DB…`**（新冻结，106,571 B） |
| pre-freeze / post-freeze `schema_changes` | **10（`CHANGED=` 上列 10 工具）** / **0** |
| hub 面 / `hub_alignment` / `ORDER_MATCH` / `strict_failures` | **115** / **0** / **True** / **0** |
| 槽位 `VB6MCP.dll` sha | == `build.status` declared（`E4A7C124…`） |
| `docs/TOOL_CATALOG.md` | **未变**（`525A9809…`） |

### 15. v1.7.0 之后的 **schema 变更**（构建 `8F6C0522…`）：`set_codepane.view` 删声明 + `caption`/`section` 回审计（**一次显式重冻**）

> 本卡授权**一次**显式重冻（`python 维护者的契约导出脚本（未随源发布） --freeze`）。
> 冻结件 sha：**`6808B1DB7F2FE9546A20DFA4900F544CFDA6AFBE28ACB9007624018A434CF97D`（106,571 B）
> → `871FE83226677FB9E85F96A8851D05D1626F045E969D9A4DB4AEF622BFF9E196`（106,715 B）**
> （`tools/schema_baseline.json` == `docs/tools_schema.json`；工具数 **111 → 111**，未变）。
> pre-freeze 实测（**先看到变化才准重冻**，§1 流程）：`ADDED=none` / `REMOVED=none` /
> **`CHANGED=control_window,property_bag,set_codepane`** / `schema_changes=3`（rc=1）；post-freeze：`schema_changes=0` ✅。
> `docs/TOOL_CATALOG.md` **变了**（`525A9809… → D44C5ED1…`）：本版**改了工具级 description**
> （删掉 `set_codepane` 的 "switch view mode"）—— 这一点与 112/113（catalog 未变）不同。
> ★**契约版本仍 `v1.7.0`**：本卡只授权一次重冻，**未**授权版本递增（与 v1.7.0 内各次重冻同一处置）。
> ★**级联（纪律 42）**：此后所有卡的「冻结件不得变」一律以 **`871FE832…`** 为准（旧 `6808B1DB…` 已结案）。

#### 15.1 裁定总表

| # | 工具.`键` | 旧声明 | 判定 | 处置 |
|---|---|---|---|---|
| 1 | `set_codepane.view` | `enum {full_module,procedure}` + 工具级描述 "switch view mode" | **实现不可能**（三条独立证据见 15.2）⇒ 声明**永不生效** = 契约谎言 | ★**删声明**（移除 property + 工具级描述删去 "switch view mode"） |
| 2 | `control_window.caption` | "read-only in VB6EXT; action accepted but ignored" | 描述**正确**，但**忽略无审计**（纪律 38） | 保留声明 ⇒ **回审计** + 描述补一句 |
| 3 | `property_bag.section` | "Section (project property bag only)" | `target=component` 时**静默丢弃** | 保留声明 ⇒ **回审计** + 描述补一句 |

#### 15.2 `set_codepane.view` 为什么只能**删声明**（三条互相独立的证据）

1. ★**类型库是 GET-only**：`VB6EXT.OLB` v5.3（`{EF404E00-EDA6-101A-8DAF-00DD010F7EBB}`）经 pywin32 makepy 展开后，
   `_CodePane._prop_map_get_` 含 `CodePaneView`，而 `_prop_map_put_` **只有 `TopLine`**；
   并且 `_CodePane` **根本没有 `View` 属性**（证据：`内部工作区`）。
2. ★**真机命令面没有该开关**：真机 `command_bars{action:"list"}` 枚举 **34** 个 bar，控制项**自身** caption/tag
   匹配 `全模块|过程视图|procedure view|full module` 的命中 **0** 条（只有 `视图(&V)` 菜单本身 + 一个 disabled 的 `修改自定义视图`）。
3. ★**真机请求无效**：`get_codepane.view` 在 `set_codepane{view:"procedure"}` **前后都是 `full_module`**
   （旧臂 S1/S2 · 新臂 S1/S2 · smoke，共 5 次运行，全部 `unchanged=True`）。

⇒ **不存在任何 add-in 可用路径**（「合成鼠标/按键」属用户已列但**未裁定**的提案 A/B，不在本卡范围）。
调用方若仍传 `view`：**不再静默** —— 回执带 `ignoredArgs:["view"]` + `ignoredArgsNote`
（历史标记 `applied:["view-ignored"]` 保留，行为不变）。

#### 15.3 `caption` / `section` 回审计的形状（纪律 38）

* 只在**参数确实出现**时回审计 ⇒ 不带该参数的调用**逐字节不变**（`sp_noargs` / `cw_cap_pos` / `pb_nosec_neg` 全 SAME）；
* **无假阳性**：真正被 honour 的输入**不得**出现审计字段 —— `set_codepane{selection}` ⇒ `applied:["selection"]`（无 ignoredArgs）、
  `cw_hide`/`cw_show`（无）、`pb_sec_pos`（`target=project` 的 `section` **确实生效**、回读 `"2"`，无）；
* 审计字段成对出现：`ignoredArgs`（数组）+ `ignoredArgsNote`（可执行说明：为什么、以及替代做法）。

#### 15.4 等价（纪律 36 + 43：各臂独立逼出后取并集）

| 同构建对 | **独立逼出**的移动路径 | 条数 |
|---|---|---|
| 旧臂 `E4A7C124` S1 vs S2 | `build.readAt` · `ide.pid` · `ms` · `wheel.skipped[].hwnd` | 4 |
| 新臂 `8F6C0522` S1 vs S2 | 上列 + **`receiptBytes`** + **`serializeMs`** | 6 |
| ★**并集（本卡采用）** | — | **6** |

★**两臂集合不一致 ⇒ 追根因、未掩盖**（纪律 43）：新臂对里 `receiptBytes`/`serializeMs` 只在 `ud_ac_pos` 移动。
对该臂逐字段枚举（纪律 44）：两轮回执**只**差 `ide`(pid 19796/11232 —— **同为 5 位**) · `receiptBytes`(1207/1206) · `serializeMs`(**15/0**)。
两个数字字段宽度差 **0**，`serializeMs` 宽度差 **+1** ⇒ **`receiptBytes` 的 +1 字节由 `serializeMs` 的宽度完全解释**
（`receiptBytes` 是**回执自身的字节数**，回执里又回显这些数字）。旧臂对同两轮 `serializeMs` 都是 `0` ⇒ 该字段**有时不动**，
所以只出现在新臂集合（纪律 43 ③）。**并集之外零移动路径。**

结果：同构建对 **66/66 SAME**（新、旧各一对，`CMP_EXIT=0`）；跨臂 **58/66 SAME**，8 个 DIFF =
**6 个有意臂**（`sp_view_neg` · `sp_view_neg2` · `cw_cap_neg` · `cw_cap_neg2` · `pb_sec_neg` · `sc_sp_badview`）
+ `mcp_info` + `eq_addins`；后两者**逐字段枚举后全部是构建身份**（纪律 44）：

| 臂 | 逐字段差异（并集归一化之后） | 是否预期 |
|---|---|---|
| `mcp_info` | `build.sha256/builtAt/path`（`build.readAt` / `ide.pid` / `wheel.skipped[].hwnd` 已被并集规则归一） | ✅ 构建身份；`build.path` 还**自证加载的是哪一条臂**（`…arm\VB6MCP.old1114.dll` vs `…new1114.dll`） |
| `eq_addins` | `.addins[3].guid`：`{D5629954-6CAC-452F-A503-50CEB9B111F9}`（旧构建 CLSID）vs `{EB6F878B-15C8-49C3-B47C-B62B03F35B69}`（新构建 CLSID） | ✅ 构建身份（纪律 28：每个构建注册**构建专属 CLSID**） |

★**所有成功臂与所有既有拒绝逐字节不变**：`eq_*` · `ref_*`（7 条既有拒绝）· `ud_*`（8 条缺参拒绝）·
的 `property_bag` value/has_value 回归臂 —— 全部 SAME。

#### 15.5 判据 (5)：全工具面扫描（口径 + 命中；**只列出不修**）

**口径**：
(a) **静态全 face** —— 把 `docs/tools_schema.json` 里 111 个工具的**每一个 declared property** 与 dispatch
`Select Case`（`CodeTools.cls` 1965–2290）里**实际读取**的键逐一对照 ⇒ **declared-but-never-read = 0**
（脚本 `内部工作区`，输出 `out/scan_static_1114.txt`）。
(b) **条件 honour 族** —— 键**被读**、但只在**某些条件下**才生效且**无审计** ⇒ 每条**上真机**。

| 命中 | 真机证据（新构建 `8F6C0522`） | 性质 | 处置 |
|---|---|---|---|
| `control_window.left/top`（`action≠move`） | `activate` + `left:123,top:45` ⇒ CodeWindow 仍 `(111,88)`；回执**无** `ignoredArgs` | 条件键被静默忽略 | **只列出** |
| `control_window.left/top`（字面 `0`） | `move{left:0,top:0}` ⇒ 仍 `(111,88)`（`If left <> 0 Then` 丢弃） | ★**合法值 0 被静默丢弃** | **只列出** |
| `control_window.width/height`（`action≠resize`） | `activate` + `width:300,height:200` ⇒ 仍 `700x420`；回执**无** `ignoredArgs` | 条件键被静默忽略 | **只列出** |
| `property_bag.name`（`target=project`） | `{target:"project",name:"mb",section:"sc2",key:"k2a",value:"5"}` ⇒ 写入成功、回读 `"5"`，`name` 无审计 | 条件键被静默忽略 | **只列出** |
| `add_code` 的 `code`（与 `file` 同给） | `{code:"'SCAN_CODE_DROPPED", file:…}` ⇒ 展开后**只有** file 的 `'SCAN_FILE_WINS`，**没有** code 的标记；回执无审计 | ★两个 declared 键**后者覆盖前者**、前者静默丢 | **只列出** |

**正例对照（纪律 40：正反例结论必须不同）**：`move{111,88}` 真的移到 `(111,88)`；`resize{700,420}` 真的变 `700x420`；
`add_code{code}` **单独**给时标记**确实出现**；`property_bag{target:"component",name:"mb"}` 的 `name` **确实生效**。
⇒ 5 条命中都是「**只有条件不满足时才丢**」，与「该键永远不生效」不同，但**同样没有审计**。

#### 15.6 门禁（重冻后，clean IDE）

`内部证据` —— **14/14 PASS, ALL_GATES=True**：
**111** · 键恰 `{tools}` · `schema_changes=0` · `ORDER_MATCH=True` · `live_tool_changes=0` · `strict_failures=0` ·
**冻结件两份 == `871FE832…`** · hub face **115**（111 add-in + 4 hub-only）· `hub_alignment=0` ·
槽位 sha == `build.status` declared（`8F6C0522…`）· `InprocServer32`==SLOT · `LoadBehavior=3` ·
clean IDE（`list_addins count=21`，`connect=true` 仅 `VB6MCP.Connect`）。
---

### 16. v1.7.0 之后的 **描述级重冻 + 响应侧字段**（构建 `AB1F2057…`）：`reload_component` / `import_component` 的「绑定文件」语义与回执如实性

> 本节改 **两个工具的 `description`**（触 schema ⇒ **一次显式重冻**，S 预先授权）+ **`reload_component` 的三个响应字段**（响应侧，不触 schema）。
> 工具面仍 **111**、契约版本仍 **v1.7.0**（未递增）；`inputSchema` 的**结构与属性名集合未变**（`scope_diff_1119.txt` 证明唯一差异是 description 文本）。

#### 16.1 语义（用户最容易踩的坑，来自一次稳定复现）

* `import_component{mode:"file"}` **把源文件复制进工程目录**并把组件绑定到**那个副本**（源文件**不被链接**，返回时校验其 sha256 未变）。
* `reload_component{name}` 重载的是**组件绑定的文件**（`VBComponent.FileNames(1)`）—— **不是「你刚改的那个文件」**。改导入源再 reload ⇒ **回执照样成功，而代码窗格保持旧内容**（H 臂 3/3；在新夹具 `P119R` 上复现 3/3）。

#### 16.2 变更（`CHANGED` 恰为 2 个工具）

| 工具 | 旧 description | 新 description（要点） |
|---|---|---|
| `import_component` | `Add an existing source file (or a template) to the project as a component.` | + `mode=file` **会复制进工程目录并绑定副本**；`componentFile` / `projectDirWrites` 报的就是那个副本 |
| `reload_component` | `Reload a component from its file on disk (VBComponent.Reload). Use after write_file ...` | + 重载的是 **THE FILE IT IS BOUND TO**（`FileNames(1)`），**不一定是你改的那个文件**；回执报 `reloaded` 与 `file`；或改用 IDE 内编辑工具 |

#### 16.3 `reload_component` 的响应字段（纯增量，调用方按需读）

| 字段 | 含义 |
|---|---|
| `reloaded` | 组件的**真名**（**在 `Reload` 之前**读取 —— `FindComponent` 刚用 `LCase$(comp.Name)` 匹配过，所以该值可证为真名） |
| `requested` | 调用方传入的 `name` 原文 |
| `file` | 实际被重载的**绑定路径**（`FileNames(1)`；取不到时为空串） |

#### 16.4 缺陷处置（D-1115-01 / D-1115-02）

* **D-1115-02（回执假名）⇒ 已修**：此前 `reloaded` 读的是 `Reload` **之后**的 `comp.Name`，而目标组件**代码窗格从未打开**时 VB6 返回**默认名**（标准模块 = `Module1`）。改为**重载前**读取。
  * 同口径对照（同夹具 `P119R`、同步步骤）：`J1-pane-never-opened` 由 **`Module1` → `mNew2`**；H/I 臂由 `Module1` → `mNew`。
* **D-1115-01（改错文件却报成功）⇒ 行为未改、回执改为可判定**：`reload` 仍只重载绑定文件（这是 `VBComponent.Reload` 的语义，**不是缺陷**），但现在回执给出 `file`，调用方可直接与「我改的路径」比对。
  * H 臂改后仍 3/3 `paneHas778=false`（**行为如实保留**），但 3/3 回执带 `file=…\Module1_2.bas` ⇒ 调用方**能判定**重载覆盖不到自己改的文件。
* ★**未做**：`reload_component` **不做窗格前后校验**（不加 `paneShaBefore/After`）；D-1115-01 的「建议」仍只是建议。

#### 16.5 不变式与冻结（本卡实测，`内部证据`）

| 项 | 值 |
|---|---|
| 工具面 / `tools/list` 顶层键 | **111** / 恰 `{tools}` |
| 契约版本 | **v1.7.0**（未递增） |
| 重冻前 `check_schema_baseline` | `CHANGED=import_component,reload_component` · `schema_changes=2`（**先可见**） |
| 重冻后 `schema_changes` / `live_tool_changes` / `strict_failures` | **0** / **0** / **0** |
| `docs/tools_schema.json` == `tools/schema_baseline.json` | **`C9C96324A1A19304…`**（108,019 B）—— 旧 `FAC01EC70…`（107,400 B） |
| hub 快照 / `hub_alignment` | **`6EA40C0C…`**（111,116 B）/ **0** |
| 槽位 `VB6MCP.dll` | **`AB1F2057…`**（2,076,672 B）== `build.status` declared；旧 `5C5AB249…`（2,072,576 B） |
| 改前副本 | `CodeTools.cls.before119`（695,804 B `677CF8AB…`）· `内部工作区`（`5C5AB249…`） |

★**级联（纪律 42/58）**：基线自此为 **`C9C96324…`**；此后各卡判据中「冻结件不得变」**一律以该 sha 为准**。
---

### 17. 【追加记录 · append-only】描述级重冻与 `wnd_probe` 语义补登（2026-10-10 补）

> ★**本节是迟到的补登**：当时**没有**在本文里留条目（`内部证据` 写的「迁移条目正本见 `docs/CONTRACTS.md` §15」实际指向的是另一节的 §15）⇒ 按纪律 12（已验收件只追加、不篡改）在此**追加**；内容全部取自该卡**已验收**的报告与证据，**未改动任何已验收产物**。

#### 17.1 变更（该卡唯一一次显式重冻；改的是 `vba6_probe` 的**参数级**描述）

* **`CHANGED` 恰为 1 项**：`vba6_probe` 的 `inputSchema.properties.wnd_probe.description`，**264 → 949 字符**；`added=[]` / `removed=[]`，顶层键相同、`inputSchema` 属性名集合相同、唯一差异是该描述（`scope_diff_1117.txt`）。
* 描述写定的**语义**：「**本进程顶层窗的直接子窗（深度恰好 2）**；顶层窗自身与孙及更深**永不入列**；无可见性 / owner / rect / style 过滤」+ **NOT a complete window enumeration**（覆盖率 **27/76 ≈ 36 %**）+ **STRUCTURAL BLIND SPOT**（对话框通常是顶层 `#32770` ⇒ 找对话框用 `find_gui`）+ 源码出处 `Vba6Probe.bas` 的 `WindowProbeJson`（L937-1004）。
* ★该语义在 被**升级为实测精确等式**（见 §55 的更正段）：`wnd_probe` 报的集合 **==** 宿主侧独立枚举算出的 `{depth == 1}`（27 == 27，双向差集皆空）。

#### 17.2 指纹与不变式（实测）

| 项 | 旧 → 新 |
|---|---|
| 冻结件（两份） | `871FE832…`（106,715 B）→ **`FAC01EC70…`（107,400 B）** |
| 交付 DLL | `8F6C0522…`（2,072,576 B）→ **`5C5AB249…`（2,072,576 B）** |
| hub 快照 | `5F70FAF9…` → **`085D18C5…`（110,497 B）** |
| 工具面 / 契约 | **111 → 111** / **v1.7.0（不递增）** |
| `schema_changes` / `live_tool_changes` / `hub_alignment` | 重冻前 `1`（先可见）→ 重冻后 **0 / 0 / 0** |
| `docs/TOOL_CATALOG.md` | 重生成后**逐字节未变**（26,749 B）—— 它只收录**工具级**描述，而本案改的是**参数级** |
| 构建后冒烟（纪律 41） | **15/15 PASS**（`内部回执`） |
| 文档 | FAQ 新增 **§55**（三口径不可互引）与 **§56**（`reload_component` 只重载绑定文件） |

* 证据：`内部证据{内部窗口探针说明,gates1117.txt,内部回执,scope_diff_1117.txt,内部工具面快照}`。
* ★**环境事实**：入场时 `build.bat` 因**外部实例（pid 19784 / 8766）加载了我们的 add-in 并锁住交付 DLL** 而报「权限被拒绝」⇒ 用「**改名挪走 + 同字节复制回槽位**」重建（★**未 kill 外部实例**）；该配方后来固化为**纪律 60**。
* ★**级联（纪律 42/58）**：该卡的基线为 `FAC01EC70…`；★该值已被后续重冻再次变更，**当前基线见 §16**（`C9C96324…`）。
---

### 18. `find_gui` 集合稳定性修复（D-1118-01）—— ★**零 schema 变更、零重冻**

> 本节**不改 `inputSchema`、不改工具面、不改任何描述** ⇒ 工具仍 **111**、契约仍 **v1.7.0**、冻结件 **`C9C96324…`（108,019 B）跑前 == 跑后**、hub 快照 **`6EA40C0C…`（111,116 B）未变**、`hub_alignment=0` ⇒ ★**免重冻**（S 预先授权的第三次重冻**未使用**）。
> ★理由：`find_gui` 的**现有**描述本来就是「*Enumerate the windows of THIS process (top-level windows plus all their children)…*」—— **那正是修好之后的行为**，改前的实现做不到自己声明的事。

#### 18.1 缺陷（D-1118-01）

`WinProbe.bas` 的 `FindGuiJson`（以及同模式的 `UiFindByClass`）在子窗阶段**把 `m_fgHwnd` 同时当「源列表」与「目的缓冲」**：`m_fgCount` 在进入子窗循环前被归零，而 `FgCollect` 从**索引 0** 起写同一个数组 ⇒ **第 1 次迭代后源列表即被子窗覆盖**，后续迭代取到的已不是原顶层窗。可观测后果三条：

1. **不稳定**：走到本进程子树的概率取决于当时桌面 Z 序（同一实例不同时刻实测 **26 / 76 / 80 / 132 项**）；
2. **重复**：同一 hwnd 被反复收集（实测 132 项 / 93 唯一 = **+39 重复**；80 / 57 = **+23**）；
3. **漏报 ⇒ 静默假否定**：可以**只返回顶层窗**（对 `vba6_probe{wnd_probe:true}` 的覆盖 **0/31**），于是 `find_gui{cls:"VbaWindow"}` 等过滤查询回 **0**，使调用方得出「该窗/类不存在」的错误结论（实测 **6/6 类假否定**）。

#### 18.2 修法（最小改动：4 处编辑 / 103 行 diff）

| # | 改动 | 作用 |
|---|---|---|
| 1 | 新增私有助手 `FgSnapshotTops(ByRef topsArr() As Long) As Long`：`FgEnumWindows` 后把顶层窗**拷贝进局部动态数组**再返回个数 | ★**断开别名**（两处调用点共用） |
| 2 | `FindGuiJson` 改为遍历 `topsArr`（顶层拷贝处补 `total <= UBound(all)` 守卫） | 消除「源被目的覆盖」 |
| 3 | `FindGuiJson` 发出循环增加**按 hwnd 唯一**（内联 O(n²)，n ≤ 1600，不引新依赖；边界先判后取） | ★**唯一性由构造保证**，不靠运气 |
| 4 | `UiFindByClass`（`ui_text` 的按类找窗助手）改为同一套快照遍历 | ★同类模式一并修（纪律 64） |

#### 18.3 改后同口径对照（同一夹具、同一步骤、复用当时的原脚本）

| 阶段 | findGuiN | 唯一 hwnd | 重复附加项 | `wnd_probe` n | 覆盖率 | 过滤假否定 |
|---|---|---|---|---|---|---|
| `before_two` | 80 | **57** | **+23** | 27 | **51.9 %** | — |
| `after_two` | 70 | **70** | **0** | 27 | **100 %** | — |
| `before_single` | 26 | 26 | 0 | 31 | **0 %** | **6 / 6** |
| `after_single` | 95 | **95** | **0** | 31 | **100 %** | **0 / 6** |

* ★连打 20 次：改后 `n` 与 **hwnd 集合逐次相同**（`varying=0`，`zeros=0`）；★`hammer` 相位实测改前 b 实例 **132 项 / 93 唯一**。
* ★过滤命中改后**逐类精确相等**：`VbaWindow 4` · `PROJECT 1` · `MDIClient 1` · `ToolsPalette 1` · `VBSlider 4` · `wndclass_flex_gsk 8`（改前全 0）。

#### 18.4 独立宿主侧 oracle（纪律 63：先自证仪器）

普通 libpython + 纯 Win32（`EnumWindows` + 对每个桌面顶层窗 `EnumChildWindows` + 按 pid 过滤）独立算出本进程窗集合：桌面顶层 **628**、本进程顶层 **22**、后代 **48** ⇒ **宿主集合 70**；同一时刻 `find_gui` = **70 项 / 70 唯一** ⇒ ★**集合相等（双向差集皆空）**。证据 `内部证据`。

#### 18.5 不变式与门禁（`内部证据`，★只读、纪律 61 + 62）

| 项 | 值 |
|---|---|
| 工具面 / `tools/list` 顶层键 | **111** / 恰 `["tools"]` |
| 契约版本 | **v1.7.0**（未递增） |
| `schema_changes` / `live_tool_changes` / `strict_failures` / `hub_alignment` | **0 / 0 / 0 / 0** |
| 冻结件（两份） | **`C9C96324A1A193047B1A7B85E9E0038D15CE798D6B2B76460241CBF1B17950B0`（108,019 B）** —— 跑前 == 跑后 |
| hub 快照 | **`6EA40C0C…`（111,116 B）** 未变 |
| 槽位 `VB6MCP.dll` | **`894447B3C209EE9AB4616CFE6340CA28C3F39A61E5562B8A91DC94F41EDD2DA8`（2,076,672 B）** == `build.status` declared；旧 `AB1F2057…` |
| 构建后冒烟 | **11/11 PASS**（`内部回执`） |

#### 18.6 行为变更（调用方需知情）

* `find_gui` 现在**按 hwnd 唯一**且**完整覆盖本进程窗**（含全部子窗）。★**只有把「项数」或「某窗不在列表里」当判据的调用方会受影响** —— 那类判据本来就不可靠。
* `ui_text` 的按类找窗助手同步修复（只可能**找得到**以前找不到的窗）。
* ★**级联：无**（冻结件与 hub 未变，工具数与契约未变）。
* ★**遗留边界（不是缺陷）**：子窗聚合仍受 `FG_MAX = 800` 上限约束（实测本夹具 70 / 95 窗，远未触及）。

### 19. `v1.9.0`：`ide_quiet_mode`（A）+ `immediate_eval`/`immediate_run`（B）+ `ide_state{release_inflight}` 与卡死自愈（C）—— 两项

> 基线（本版之前，实测）：**add-in 111 工具 / contract v1.8.0 / DLL `79609DBC…`（2,568,192 B）**。
> 本版：**add-in 114 工具 / contract v1.9.0 / DLL `158BB063D6ED4A91…`（2,637,824 B）**，**四冻件全部重冻**。

#### 19.1 A · `ide_quiet_mode`（会话级安静模式）
- 参数：`on`（boolean，**必填**）、`ttl_ms`（integer，可选，默认 600000 = 10 分钟；TTL 每次读取时重新判定，到期自动失效）。
- 返回：`{ok, quietMode, previousQuietMode, changed, since, ttlMs, quietExpiresAt, expiresInMs, suppressedFocusCount, foregroundBefore, foregroundAfter, foregroundTookFocus, foregroundReadable, advisory, suppressedPoints, note}`。
- 语义（ON）：**`control_window{action:"activate"}` 不调 `Window.SetFocus`**（该回执带 `quietModeSuppressedFocus:true` + `suppressedAction` + `quietModeNote`，**其余动作照做**）；且**每次** `tools/call` 正文带 `quiet{quietMode, since, ttlMs, quietExpiresAt, expiresInMs, suppressedFocusCount, advisory, note}`，`advisory` = 「quiet mode is ON - do not type in this IDE while an agent is driving it」。
- `native_breakpoint` 是 **Debug 菜单命令**，**无法在不激活 IDE 的前提下发出** ⇒ 安静模式只能在那里**交还前台**（回执 `quietModeSuppressedFocus:false` + `quietModeForegroundRestored`），**如实标注**，不谎称抑制。
- ★**不得用 0/0 声称「没抢前台」**：前台值是**实时读出**的 —— 交互会话里 `GetForegroundWindow()` 返回**有效非零句柄**；**锁屏/非交互桌面**上它可能是 **0**，此时 `foregroundReadable=false` 只表示**那一次读不到**（如实回报，不猜）。判据必须写成 `fgAfter == fgBefore`（**值可为 0**）**且** `quietModeSuppressedFocus:true`；工具**绝不抢焦点**去取该值，也**绝不**用窗口站名字推断「能否发送输入」。
- 分类为**只读**（只改插件自身状态）⇒ 写类在飞时**仍可被服务**。
- ★**quiet OFF ⇒ 逐字节回到 v1.8.0**：不抑制、不加 `quiet{}`。

#### 19.2 B · `immediate_eval` / `immediate_run`（立即求值 / 运行并取回输出）
- `immediate_eval{expression, timeout_ms=5000}` ⇒ `{ok, value, typeName, valueParsed, output, outputTrimmed, outputLines, outputBytes, outputTruncated, debugPrintRewrites, line, ret, retHex, execAnswered, elapsedMs, sinkFile, sinkDeleted, ...}`。表达式**只求值一次**（先赋给 `v511`，再 `Print #511, TypeName(v511) & "|" & v511`）。
- `immediate_run{code, timeout_ms=5000}` ⇒ 同形状；`output`/`outputTrimmed` 只含**本次**打印的文本；**语句边界**的 `Debug.Print` 被内部改写到 sink（`debugPrintRewrites` 如实计数）。
- ★**调用方不再需要写 `Open`/`Print #`/`Close`**：工具内部把值/输出经**私有临时文件**带出（`%TEMP%\vb6mcp_im\*.tmp`，读完即删，`sinkDeleted` 回读证明）。
- ★**解析失败必须如实结构化**：`{ok:false, reason:"value_not_parsed"}` + 原始 `output`/`outputTrimmed`，**不许猜**；引擎拒行时 `reason:"engine_refused_line"`（带 `ret`/`retHex`）。
- ★**实测通道事实**（本轮真机，原始回执在 未随源发布的内部证据）：
  `Open "…" For Output As #511: Print #511, 1+1: Close #511` = `ret=0`（sink `" 2 "`）；
  `v511 = 3/4: Print #511, TypeName(v511) & "|" & v511: Close #511` = `Double|0.75`；
  **`On Error Resume Next`（单独或冒号连写）被拒 `ret=40036`**；**行内 CR/LF 亦被拒 `ret=40036`** ⇒ 多语句以「冒号 + 空格」连接，`?` 简写与 `_` 续行**不支持**。
- ★**抛错语句仍会挂住 IDE**（VB6 弹出它自己的模态「实时错误」框，见 19.3）：因此 `timeout_ms` **不是硬超时** —— 回执恒带 `timeoutMsEnforcedBy:"gate_age_threshold"`、`hardTimeoutEnforced:false`、`statementInterrupted:false`；它是**闸门年龄阈值**，到达后由**下一次工具调用**自动释放闸门（见 19.3 的 armed release）。
- ★既有 `vba6_probe{exec_probe}` 与既有回执**逐字节不变**（本版未改其实现）。

#### 19.3 C · 缺陷 #10：单飞闸门可强释 + 卡死自愈
- **实测根因**（本机 v1.8.0 构建，原始回执 未随源发布的内部证据）：抛错语句使 `EbExecuteLine` **不返回**，IDE 弹出**它自己的模态框**（`read_active_dialogs` 报 `ownerDisabled=true`，正文「实时错误 '13': 类型不匹配」），`m_mutBusy` 不再释放 ⇒ 此后**所有写类调用**一律 `{ok:false,reason:"busy"}`，而**只读仍被服务**（`ide_state` 19 ms 应答）。★**把模态框点掉后闸门自行恢复**（实测：`set_codepane` 立刻不再 busy）⇒ 根因是那个模态框，不是「线程楔死」。
- **强释入口**：**`ide_state{release_inflight:true, release_min_held_ms?}`**（默认阈值 60000）—— **当且仅当**在飞调用**年龄 > 阈值**时才动手；只读 ⇒ 卡死期间**仍可被服务**。
- **回执 `release{}`（固定键集，每条 ide_state 都有）**：`{attempted, released, releasedTool, releasedKind, heldMs, minHeldMs, retryAfterMs, depthBefore, depthAfter, targetDepth, mutBusyBefore, abandonedSeq, reason, gateReleaseCount, statementInterrupted:false, timeoutMsEnforcedBy:"gate_age_threshold", note, recoveryAction, recoveryHint}`。
  `released:true` 时 `note` 明写「**被阻塞的调用仍在 IDE 内运行；闸门是被强释的；它做的事没有被撤销；没有语句被打断**」。
  `reason` ∈ `not_requested` | `nothing_in_flight` | `age_below_threshold`（带 `retryAfterMs`）| `forced_release` | `auto_released_armed_timeout`。
- ★**弃帧语义（防二次扣减）**：每次调用带 token（`m_callSeq`），**最外层帧**的 token 记在 `m_outerSeq`（在 depth 变 1 时捕获，这样嵌套读**不会**把自己当成在飞者）。强释时把**被弃帧**的 token 记入 `m_abandonSeq` 并把 `m_callDepth` 归零。被弃帧日后返回时，其 unwind 在 `CallUnwind` 里**被吞掉恰好一次**；并且**只有持有者**（`m_mutBusyTool = name`）才能清 `m_mutBusy` ⇒ **不可能出现「归零后旧帧再减一次」**（那会让后续调用被错误拒绝，且偶发难查）。
- ★**自愈（armed release）**：`immediate_*` 在调用 `EbExecuteLine` **之前** arm（`timeout_ms`），正常返回即解除；若没返回，**下一次任意工具调用**在入场处按年龄强释 ⇒ **一次坏语句不会把会话永久卡死**，无需人工干预。
- ★**完整恢复配方（机器可读，写在 `recoveryAction`/`recoveryHint`）**：①`ide_state{release_inflight:true}` 强释闸门 → ②`close_active_dialogs` 把那个模态框点掉（被阻塞的语句才可能返回）→ ③ 正常继续。★注意 `close_active_dialogs` 自身是**写类**，卡死期间会被 `busy` 拒绝，所以**必须先强释**。
- 只读回执在**发生过强释之后**才追加 `gateReleased` / `gateReleaseCount` / `gateReleasedBy` / `gateReleasedTool` / `gateReleasedHeldMs` / `gateReleasedAt` / `gateReleaseReason` / `gateReleaseNote`；**未发生过则键集不变**（零漂移）。
- ★`gateReleasedBy` 指明**是谁/什么**释放的：**`ide_state`** = 有一次显式 `release_inflight=true` 请求；**`armed_release`** = `immediate_*` 的 **armed timeout 自动释放**（**没有**任何调用带过 `release_inflight`）。两者的 `gateReleaseReason` 分别是 `forced_release` 与 `auto_released_armed_timeout`。★这一区分是本卡真机跑出来的：初版把自动释放也标成 `ide_state`，会让回执声称一个**没人提过**的请求，已改。

#### 19.4 不变式
- `tools/list` 结果键仍**恰为** `{tools}`；`tools/call` 结果键仍**恰为** `{content,isError}`；版本戳仍只在正文里。
- **四冻件（本版全部重冻）**：`docs/tools_schema.json` == `tools/schema_baseline.json` = **`B0C18528845DFB58…`**（123,357 B，114 工具）；`bridge/hub_schema.generated.json` = **`2AC4036B2F619B12…`**（127,021 B）；`docs/CONTRACTS.md` = 本文件。
- **产品版本 `0.3.0` 不变**（本版只改契约面，不改产品语义版本）。

---

## v1.10.0 · BATCH-1 变更记录（2026-10-13）

> hub 快照的 `contractVersion` 取自 `## 版本记录` 的**最新行**；`initialize.instructions` 的横幅取自 `## v1.3.0` 的 text 代码块 —— 两者都已随本版更新（`CONTRACT v1.10.0` / `TOOLS 115`）。

**工具面**：add-in **114 -> 115**（新增只读 `desktop_probe`），hub 5 -> 5，DSH-facing total **119 -> 120**。

### ADDED
- `desktop_probe`（只读）：报告承载 add-in 的进程的**实时**窗口站 / 桌面 / sessionId / 可交互性 / 前台状态与取证依据；**fail-closed** —— 任一 API 失败即 `ok:false` + `failureKind:"api_unavailable"`，绝不猜。前台值是**实时读出**的：交互会话里非零，锁屏/非交互桌面可能为 0（如实回报）；工具**不抢焦点**去取它，也**不**用窗口站名字推断「能否发送输入」。

### CHANGED（12 个；其中 10 个只改描述文本）
`build` · `close_active_dialogs` · `ide_quiet_mode` · `ide_state` · `immediate_eval` · `immediate_run` · `inject_trace` · `make_project` · `remove_trace` · `todo_status` · `validate_code` · `vba6_probe`
- **10 个**只做**描述文本中性化**：去掉描述里的内部任务号 / 内部目录标签；**行为与回执字段不变**。
- `ide_state` 与 `vba6_probe` 的**描述未变**，只新增参数（见下）。

### 新增参数（3，纯增量：缺省时不改变任何既有行为）
| 工具 | 新增参数 |
|---|---|
| `vba6_probe` | `exec_line_file` |
| `immediate_eval` | `expression_file` |
| `immediate_run` | `code_file` |

### R13 结构化失败信封
- `vba6_probe` 七类 `failureKind`：`not_loaded` / `alloc_failed` / `arg_out_of_range` / `address_invalid` / `not_readable` / `engine_zero` / `refused_by_design`；
- `immediate_eval` / `immediate_run` 的 `exec_error` 包装新增两类：`channel_error` / `statement_raised`；
- ★**刻意的例外**：`immediate_*` **自身**的 `exec_error` **不带** `failureKind`（那是引擎侧原始错误，不属于本信封的判定）。

### 修复
- **静默截断 -> 结构化**：请求长度与实际截断都回报（`lenRequested` / `lenClamped`，以及 `maxHits*` 一族）；
- **`arg_out_of_range`** 取代裸 VB 溢出：越界参数现在是结构化 `failureKind`，不再是一句 VB 运行时错误。

### 未落地（如实）
- **`immediate_window_text` 未创建**：R1 判定 **NO-GO**（MSAA 只暴露窗格**标题**，读不到正文）⇒ 工具面**不加**该工具；该缺口保持公开、可追踪，而不是承诺。

### D1 有界残留（`todo_status`）
- `todo_status` 的**功能性自动发现**保留：它在仓库里自动发现一个**可选的任务文件夹**，并用其中的任务/报告标题生成机器可读的 todo；
- 新增环境变量 **`VB6MCP_TASKS_DIR`** 覆盖该路径（显式给出时优先）；
- 契约与公开面**不再出现该内部目录名**：描述只说「自动发现的任务文件夹 + 环境变量覆盖」。

### 不变式
- `tools/list` 结果键仍**恰为** `{tools}`；`tools/call` 结果键仍**恰为** `{content,isError}`；版本戳仍只在正文里。
- 四冻件在本版**全部重冻**；产品版本仍 **0.3.0**（本版只改契约面）。
