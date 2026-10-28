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
| **v1.17.0** | 2026-10-28 | **MINOR 级发布（add-in 119 -> 120、hub 11 -> 11、DSH 面 130 -> 131；新增 1 个只读工具 `trace_query`；`inject_trace` 新增 2 个可选参数 `condition`/`recordCondition`（logpoint）；`remove_trace` 描述补 S-1 句；**本批显式重冻**（工具面/参数/描述变化，禁止用 schema_changes=0 代替重冻）：冻结面 211,499 B / `37564E39…` -> **221,252 B / `4A201C0C…`**（canonical CRLF；仓库 LF blob `A46F2E1B…` -> `5C3030AB…` / 217,919 B）、catalog 95,714 B / `D2EC667E…` -> **100,446 B / `1966CA14…`**、hub snapshot 236,335 B / `E839F6BC…` -> 重新生成（VERSION=**v1.17.0** / DECLARED=**120** / MISMATCH=False / 新 FACEFP=`59f8b4cb…`）；**有 VB 源码改动**（`TraceQuery.bas` 新增 **2,153 行 / 84,234 B / `957655B8…`**、`McpFace.bas` 202,841 B / `DE80BC79…` / +38 −2、`CodeTools.cls` 1,714,383 B / `722B1485…` / +109 −9、`TraceGen.bas` 67,325 B / `00E32839…` / +170 −10、`MCP_VB6.vbp` +1 模块、`Util.bas` 常量 v1.16.37 -> v1.17.0，101,541 B / `18FC7944…`）⇒ 源码印记 I=`cef99abb…`（算法仍 `vb6src/3`，FILES=55）、DLL 在 t16 重新 `/make`（**5,533,747 B / `EAE527B1…`**，builtAt 2026-10-28 15:00:37）并重签部署）** ① **`trace_query`（`b04be88` + `ec10130`，新只读工具）**：`TraceQuery.bas` 纯函数 + 薄入口 —— 字节游标 `cursor.nextOffset`（半行不推进）、严格 UTF-8 解码、8 值失败闭合集 / 7 值 warning 闭合集、头 64 KiB + range 双 sha 取证、真 UTC mtime、零写盘、B-only `bytes_budget` 安全层、`format:"jsonl"` 原文旁路；13 个参数、required 为空；与写侧 `inject_trace` 同批发布。② **`inject_trace.condition`/`recordCondition`（`b7d4414`，logpoint）**：5 行块 -> 7 行（外层 `If McpTraceOn() Then` 恒最外 ⇒ TRACE-OFF 时条件不求值）；`@cond` 走既有 kv ⇒ 生成的 modMcpTrace 模板一行不改；条件词法门 7 项 + 与 `expressions` 共用同一可见符号集，**全部在 EnsureTraceModule 之前**（拒绝时零写入、无模态框）；回执带 `conditionGuarantee:"necessary_not_sufficient"` 诚实边界；**省略 condition 时 5 行块与今天逐字节相同**。③ **`remove_trace` 描述增量（同批）**：结尾补 S-1 句（工程已无 `'@mcp-trace'` 块时可手工 `remove_component{name:"modMcpTrace"}`），**不改任何行为代码**。④ **重冻证据（t16 活体，非恒真）**：pre-freeze preview = `ADDED=trace_query` / `CHANGED=inject_trace,remove_trace` / `REMOVED=none`；活体 contractVersion=**v1.17.0** / `TOOLS 120` / `BUILT EAE527B1` / `loadedModule.mismatch=false` / `registration.dllSha256=EAE527B1…`。⑤ **批次提交**：`b04be88`（trace_query）+ `ec10130`（key 解码修复）+ `b7d4414`（logpoint + S-1）；`debug_eval_frame` 不在本批（另批，A 形态研究脚本 `tools/debug_eval_frame_probe.py` 不在 120 工具面）。⑥ **T2 窗口内修正（文本，非行为）**：首发冻结被 `check_shipped_tree.py` 的 **UNC 规则**拒绝（`trace_query.path` 描述里的字面 UNC 示例；该门 fail-closed、无白名单），改述为 `UNC/network paths (two leading backslashes, e.g. a server-share path)` 后**重跑 /make + 重新冻结**（首遍数值作废；最终值见本节与 pin）。同批 KPI M2 `tools/*.py` 计数 pin 86 -> 90；派生件 `BuildStamp.bas` 296 B / `6CFE7EA1…`、`MCP_VB6.vbp` 2,820 B / `D77851B8…`。 |
| **v1.16.37** | 2026-10-28 | **patch 级发布（工具面零变化：add-in 119 -> 119、hub 11 -> 11、DSH 面 130 -> 130；无工具增删、无参数/描述/inputSchema 变化 ⇒ 冻结面不变、**不重冻**；**本批无 VB 实现源码改动**（`.bas`/`.cls` 字节不变，仅 `Util.bas` 契约版本常量 v1.16.36 → v1.16.37，101,542 B 不变、sha `0B308A83…` -> `BD0DB4BC…`）⇒ 源码印记 I 变（算法仍 `vb6src/3`，I=`141b2a085646e0494e4b7ca4fdcf8a396f8d819b625358c5669175d7acb01ac6`）、DLL 在 t16 重新 `/make` 并重签部署）** ① **M5 wave-30（`0e88275`，纯 tools/tests/docs，5 文件 / +695 −34）**：**FF:36 数组表达式模型** —— FC:B0/FF:36 handler + FF:02 `join_array`，`ary_join_array()` 唯一提交点；**整链可证才 emit 既有 `assign`+`call` 节点**（**无新 AST 节点类型**），`ary_model{form,builtin,elements|value_args,evidence,chain}` 为 additive 元数据；**不可证链保持 raw**。（**19 条数组赋值实际提交、2 条拒绝**）⇒ **文本层 1,895 -> 1,911/3,121（+16）**、**判分 1,924 -> 1,940/3,121（+16，EXACT 1,869 -> 1,885）**、**M5-67 563 -> 565/732（+2）**；**0 pass→fail**、分母 3,121/732 不变、extras 1,191 -> 1,171（−20）/ placeholders 493 -> 472（−21）。pins：`tools/vb6_decompile.py` 600,850/`69573A98…` -> **618,432 / `75BB9F1D…`**、`tools/test_vb6_decompile.py` 298,694/`2BFF72B3…`（1,139）-> **307,324 / `BC9B4AE5…`（1,161 checks）**、`tests/test_g10_impad.py` 69,091/`FA882EA4…` -> **72,227 / `F22B1124…`**、`docs/RE_TOOLS_INTEGRATION.md` 890,507/`0779A27B…` -> **907,489 LF / `325425DD…`**（`9.134`，总 5,252 行；`9.134.5` 登记 5 条 DEFER）。② **docs**：`docs/PROJECT_STATUS_20261028.md` 刷新至 v1.16.36（`6eb3592`，单文件 / +117 −65，37,996 B）；`docs/M5_MODEL_RESEARCH_20261028.md` 预研与 Lead 裁决（`8868708`/`4c8f6a6`，26,101 B：**授权 FF:36 数组表达式模型（wave-30，fail-closed）**、E4 OnGoSub 延后、late/vcall 文本诚实度列为可选维护、Debug.Print 永久关闭且分母保留）。③ **契约 bump + re-emit**：本批**无 VB 实现源码改动** ⇒ I=`141b2a085646e0494e4b7ca4fdcf8a396f8d819b625358c5669175d7acb01ac6`，t16 重新 /make + 重签部署；工具面 119/11/130 不变、schema_changes=0 ⇒ 冻结面（**`37564E39…` / 211,499 B**）与 catalog（**`D2EC667E…` / 95,714 B**）字节不变、**不重冻**；hub snapshot 仅 contractVersion 串变（FACEFP 仍 `cacb1bd4…`）。**M5 双口径（Lead 裁定）**：文本层 **1,911/3,121**；判分 **1,940/3,121（EXACT 1,885）**；差 **+29 不变 = 8**（v1.16.31 卡 A measurement-side）**+ 10**（wave-28 结构化面）**+ 7**（caliber-2 measurement-only）**+ 4**（wave-29 结构化面）；wave-30 的 **+16 为文本与判分同步真实增益**（非 measurement-side）。M5-67 **565/732**。 |
| **v1.16.36** | 2026-10-28 | **patch 级发布（工具面零变化：add-in 119 -> 119、hub 11 -> 11、DSH 面 130 -> 130；无工具增删、无参数/描述/inputSchema 变化 ⇒ 冻结面不变、不重冻；**本批无 VB 实现源码改动**（`.bas`/`.cls` 字节不变，仅 `Util.bas` 契约版本常量 v1.16.35 -> v1.16.36，101,542 B 不变、sha `7AB43C9A…` -> `0B308A83…`）⇒ 源码印记 I 变（算法仍 `vb6src/3`，I=`477816ac…`）、DLL 重新 `/make` 并重签部署）** ① **M5 wave-29（`0e11a73`，纯 tools/tests/docs，5 文件 / +434 −38）**：两条**产品层**修复 —— (a) **同语句 staged 字面量折叠**（decode，文本与 JSON 同时移动）：VB6 把 Variant 字面量 stage 进 frame temp 再传地址，shipped import 调用因此渲染成匿名临时量（`v1 = Format$(x_-A8, "0.00")` / `{"e":"ref","slot":-168}`）；现在把**上一语句之后**才 stage 的常量折回实参（`staged_pc[slot] > last_emit_pc` = 同语句实参求值），**只对 shipped function 行**、**排除 out_first/out_param**，**负控保留**（`f06_variant` 的 `vdate=` 行保持 ref，stale staged 值绝不代入）；(b) **结构化默认参裁剪**（post-pass，仅结构化面）：`import_value_text()` 一直丢弃有文档的默认参（Format day/week、FormatNumber/Currency/Percent 负值可选参、`Replace$` R7 默认 `[1,-1,0]`），JSON 节点此前保留全部操作数，现在用**同一规则** fail-closed 裁剪（`args_dropped_defaults` 计数），**out 惯例不裁剪**（实测节点级 out 裁剪会把诚实的 unsupported import 注释变成 Call）⇒ **判分 1,915 -> 1,924/3,121（+9，EXACT 1,860 -> 1,869）**、**文本层 1,890 -> 1,895/3,121（+5）**（+5 文本可见：4 Format 家族字面量 + `VarType(True)`；另 +4 是产品 JSON 结构化面：3 Replace + 1 f04 Format，**不是 measurement-side**）、**M5-67 558 -> 563/732（+5）**、**0 pass→fail**、分母 3,121/732 不变、extras 1,191 / placeholders 493 不变、320 text A/B = 6 fixture 变化 / +12 −12 / unexpected=0；pins：`tools/vb6_decompile.py` 595,370/`FDB18D1F…` -> **600,850 B / `69573A98…`**、`tools/test_vb6_decompile.py` 291,080/`04692ED6…`（1,124）-> **298,694 B / `2BFF72B3…`（1,139 checks）**、`tests/test_g10_impad.py` 67,426/`663A6392…` -> **69,091 B / `FA882EA4…`**、`docs/RE_TOOLS_INTEGRATION.md` 871,283/`3F7D1900…` -> **890,507 LF B / `0779A27B…`（`9.133`，总 5,160 行）**。② **docs：property-sheet 域 P1 修复复测（`aea865f`，2 文件 / +174 −1）**：新增 `docs/OPEN42_PS_P1_RETEST_20261028.md`（**22,362 B**）—— **ACC-P1R = R1–R8 全 PASS（8/8）**，**D-P1-1 闭环**：`activePage` 三态（hwnd == 外部可见页 6620174/3343134/328796、`tabIndex` == TCM curSel 0/2/4、`controlCountTotal` 24/11/10 == 外部 EnumChildWindows、`controls[]` 逐项相等且 `controlsTruncated:false`）与**同刻外部**证据逐项相等；无回归：pages==materializedPages==external 1/3/5、close 三态零点击 + `property_sheet_hold` + `refusedPropertySheetCount=1`、write 入口 top-level `classification=property_sheet`、非属性表 `#32770` 仍 unknown + 零点击 hold；治理 **252/17**（失败集合逐项相同、`NEW_FAILS=0`、两条 D-P1-1 赋值目标断言 PASS）+ honesty 73/0，**HOLD 零点击**不变；工具面/schema 零变化。③ **契约 bump + 重签发**：I=`477816ac…`（`vb6src/3`），t16 重新 `/make` + 重签部署；hub snapshot 仅 contractVersion 串变（`FACEFP` 仍 `cacb1bd4…`）。④ **M5 双口径（Lead 裁定）**：文本层 **1,895/3,121**、判分 **1,924/3,121（EXACT 1,869）**；差 **+29 = 8**（v1.16.31 卡 A measurement-side）**+ 10**（wave-28 结构化面）**+ 7**（caliber-2 measurement-only）**+ 4**（wave-29 结构化面：3 Replace$ + 1 Format 默认参裁剪）；**不得**把 +29 当文本层增益；M5-67 **563/732**。 |
| **v1.16.35** | 2026-10-28 | **patch 级发布（工具面零变化：add-in 119 -> 119、hub 11 -> 11、DSH 面 130 -> 130；无工具增删、无参数/描述/inputSchema 变化 ⇒ 冻结面不变、不重冻；本批**有 VB 实现源码改动**（`WinOps.bas` 1 字节对，属性表域 P1 缺陷修复 D-P1-1）⇒ 源码印记 I 变（算法仍 `vb6src/3`，I=`9a148f9a…`）、DLL 重新 `/make` 并重签部署）** ① **属性表域 P1 缺陷修复 D-P1-1（`ad58250`，2 文件 / +14 −1）**：v1.16.34 P1 活体验收（ACC-P1 7 PASS / 2 FAIL）的两条 FAIL **同源**于 `activePage` **从未真正挂载** —— `WinOps.bas:1412` 把 `Set ps("activePage") = ap` 写成 `Set ap("activePage") = ap`（**赋值目标**写成源对象）；修复为**一个字节对**（`ap` -> `ps`），`WinOps.bas` 保持 **323,279 B**（sha `F609FAF8…` -> `7A337C14…`）；**additive 行为**：`propertySheet.activePage` 回执非 null（`tabIndex` == TCM curSel、`controlCountTotal` 真总数、`controls[]` 非 null），P0/P1 其余不变、策略仍 **HOLD 零点击**；`tools/check_modal_governance.py` 新增 **2 条赋值目标断言**（**252/17**（`tools/check_modal_governance.py` **47,112 B / `F402E229…`**；原 46,109 B / `09D542A3…` 为 250/17），17 条既有红集合不变；**NEW gate vs pre-fix 源 = 252/19** 证明能抓住该 bug）。② **M5 口径第三批（`d80c9cb`，纯 tools/tests/docs，3 文件 / +191 −3）**：**measurement-only** 呈现层规范化 —— (a) `aclib.flatten_decomp` 为 `For` 节点渲染 `step`（产品 JSON **本来就带** step，28 个 step-less for 节点字节不变）；(b) judge **单行 If 规则 A** 补救（编译器抹除源码形态；post-align 用「最小 if-header + 一个**已解码** then 语句 + `Else`/`End If` 终止符」重建，两半仍走**未改动**的 `judge_pair`，`Else`/`End If` 仍计 extras；**规则 B 未实现**）⇒ **判分 1,908 -> 1,915/3,121（+7 = 3 Step-For + 4 SLIF-A）**、**M5-67 551 -> 558/732（+7）**、**文本层 1,890 不变**、**0 pass→fail**、分母 3,121/732 不变、extras 1,199 -> 1,191、placeholders 493 不变；pins：`re_work/minerAC/judge.py` **27,676 B / `1AC0E2ED…`**（untracked，sha 记账）、`re_work/minerAC/aclib.py` **60,907 B / `CA5C92DB…`**、`docs/RE_TOOLS_INTEGRATION.md` LF **871,283 B / `3F7D1900…`**（`9.132`，总 5,041 行）；`tools/vb6_decompile.py` **595,370 B / `FDB18D1F…`**、`tools/test_vb6_decompile.py` **291,080 B / `04692ED6…`（1,124 checks）**、`tests/test_g10_impad.py` **67,426 B / `663A6392…` 不变**。③ **docs（`6cc8f79`，2 文件 / +200 −1）**：新增 `docs/OPEN42_PS_P1_VERIFY_20261028.md`（P1 活体验收 ACC-P1 7 PASS / 2 FAIL，FAIL 同源 D-P1-1 待复测）+ `docs/OPEN_ITEMS.md`；caliber 预研（`53909bc`）与授权（`17b3bd6`，含**备份纪律修正**）。④ **契约 bump + 重签发**：I=`9a148f9a…`（`vb6src/3`）、DLL t16 重 `/make` 重签；工具面零变化 ⇒ 冻结面 `37564E39…`/211,499 与 catalog `D2EC667E…`/95,714 字节不变、不重冻；hub snapshot 仅 contractVersion 串变（`FACEFP` 仍 `cacb1bd4…`）。**M5 双口径（Lead 裁定）**：文本 1,890/3,121、判分 1,915/3,121，差 **+25** = **18**（8 v1.16.31 卡 A measurement-side + 10 wave-28 结构化输出面）+ **7**（本批 measurement-only）；不得把本批 +7 当产品行为改变，也不得把文本层写成判分−25；M5-67 **558/732**（卡面历史 555 是 SLIF-A **单臂**，已更正为**合臂** 551+3+4=558）。 |
| **v1.16.34** | 2026-10-28 | **patch 级发布（工具面零变化：add-in 119 -> 119、hub 11 -> 11、DSH 面 130 -> 130；无工具增删、无参数/描述/inputSchema 变化 ⇒ 冻结面不变、不重冻；本批**有 VB 实现源码改动**（`WinOps.bas` +281 −44、`CodeTools.cls` +19，属性表域 P1）⇒ 源码印记 I 变（算法仍 `vb6src/3`，I=`509e488f…`）、DLL 重新 `/make` 并重签部署）** ① **属性表域 P1（`38f112a`，VB 源码，3 文件 / +396 −44）**：`read_active_dialogs` 的 additive `propertySheet` 描述符新增 `pages[]`（每页页标题与页内控件；`materializedPages` 真总数、512 cap 只限单次返回）、`activePage`/`activePageState`、活动页 `controls[]`（512 cap + 真总数）；页标题新增 **`TCM_GETITEMA` 回退**（实测钉 **`0x1305`**；任务书草拟的 `0x133D` 已实测证伪、不采用），与既有 `TCM_GETITEMW=0x133C` 互补、各自 fail-closed（读不到给 `stateReadable:false`+reason）；**上层 classification 与逐对话框分类对齐**（根级/嵌套/`preExistingHandled` 同一判据）；`close_active_dialogs` 新增只读 `refusedPropertySheetCount`；策略仍 **HOLD**（零点击/零关闭/零切页，无 BM_CLICK/WM_COMMAND/TCM_SETCURSEL/focus/input 路径）；`tools/check_modal_governance.py` 断言 **226/17 -> 250/17**（+24 条 P1 断言，失败集合不变）。② **M5 wave-28（`ee4f7e2`，纯 tools/tests/docs，4 文件 / +317 −12）**：**结构化 JSON callee 身份修复**（shipped `import_map` 两个 evidence 类在**表达式位置**重拼写为 VB 内建；语句位置/负控不动）⇒ **判分 1,898 -> 1,908/3,121（+10，EXACT 1,845）**、**M5-67 549 -> 551/732（+2）**、**文本层 0 变化（1,890/3,121）**、0 pass→fail、分母不变；pins：`tools/vb6_decompile.py` **595,370 B / `FDB18D1F…`**、`tools/test_vb6_decompile.py` **291,080 B / `04692ED6…`（1,124 checks）**、`tests/test_g10_impad.py` **67,426 B / `663A6392…` 不变**、`docs/RE_TOOLS_INTEGRATION.md` LF **853,095 B / `818A2CF7…`**（`9.131`，总 4,923 行）。③ **docs（`20d293f`，2 文件 / +264 −1）**：新增 `docs/OPEN42_PS_VERIFY_20261028.md`（P0 活体验证 ACC-1..12，deviation 7 / UNKNOWN 5 如实记录）；`docs/OPEN_ITEMS.md` 更新。④ **契约 bump + 重签发**：I=`509e488f…`（`vb6src/3`）、DLL t16 重 `/make` 重签；工具面零变化 ⇒ 冻结面 `37564E39…`/211,499 与 catalog `D2EC667E…`/95,714 字节不变、不重冻；hub snapshot 仅 contractVersion 串变（`FACEFP` 仍 `cacb1bd4…`）。**M5 双口径（Lead 裁定）**：文本 1,890/3,121、判分 1,908/3,121，差 **+18** = 8（v1.16.31 卡 A measurement-side）+ **10**（wave-28 结构化输出面：产品 JSON callee 身份真实改变、文本 0 移动）；不得把 +10 当 measurement-side 抹掉，也不得把文本层写成判分−8；M5-67 551/732。pins：`WinOps.bas` 323,279/`F609FAF8…`、`CodeTools.cls` 1,708,375/`A917A5C7…`、`tools/check_modal_governance.py` 250/17（+24 断言、失败集合不变）、`tools/vb6_decompile.py` 595,370/`FDB18D1F…`、`tools/test_vb6_decompile.py` 291,080/`04692ED6…`（1,124 checks）、`tests/test_g10_impad.py` 67,426/`663A6392…` 不变、`docs/RE_TOOLS_INTEGRATION.md` LF 853,095/`818A2CF7…`（`9.131`）、`re_work/minerAC/judge.py` 22,904/`93EB2EAB…`、`aclib.py` 60,464/`D0B35B9B…` 不变、`tests/test_g10_udt.py` 69,182/`88231A14…` 不变。 |
| **v1.16.33** | 2026-10-28 | **patch 级发布（工具面零变化：add-in 119 -> 119、hub 11 -> 11、DSH 面 130 -> 130；无工具增删、无参数/描述/inputSchema 变化 ⇒ 冻结面不变、不重冻；本批**有 VB 实现源码改动**（`WinOps.bas` +449 −7，属性表域 P0）⇒ 源码印记 I 变（算法仍 `vb6src/3`，I=`95b16be5…`）、DLL 重新 `/make` 并重签部署）** ① **属性表域 P0（`4a1ac2e`，VB 源码，1 文件 / `WinOps.bas` +449 −7）**：结构性识别 `property_sheet` = 本进程可见 `#32770` + **直接子** `SysTabControl32` id=12320 + frame 按钮 {1,2,9,12321} 全在且为按钮；缺任一条 fail-closed 回 `unknown`（标题只作 `titleCorroborated` 佐证）；`read_active_dialogs.dialogs[]` additive `classification:"property_sheet"` + `propertySheet{readOnly/basis/titleCorroborated/tabControl{hwnd,id,tabCount,curSel,stateReadable,stateBasis,stateReason}/tabs[]/tabsTruncated/frameButtons[]/materializedPages/limits/mappingNote/readOnlyBasis}` + 根 `propertySheetCount`；**同批零点击闸门**：`close_active_dialogs` 对属性表框一律 `skipped[].reason="property_sheet_hold"`、`clickedCount=0`/`plannedCount=0`（dry 与非 dry 同），入口 pass / guard / `preExistingHandled` 一律 hold（action=hold），嵌套 ledger `held=true/result=held_policy/clickedId=0`；结构分支位于 host_modal 之后、所有文本 token 分支之前（正文增长不可能读成 `click_ok`）；**`TCM_GETITEMW=0x133C`**（设计文 0x1373 有误：commctrl.h `TCM_GETITEMW=TCM_FIRST+60=0x133C`，32-bit comctl32 5.82 实测 0x133C 可读、0x1373 返回 FALSE）、`TCM_GETITEMCOUNT=0x1304`/`TCM_GETCURSEL=0x130B` 经 `SendMessageTimeoutW(SMTO_ABORTIFHUNG|SMTO_BLOCK,250ms)`；只读、零点击、零 tab 切换、无 BM_CLICK/WM_COMMAND/TCM_SETCURSEL/focus/input 路径。② **M5 wave-27（`2e833eb`，纯 tools/tests/docs，5 文件 / +341 −23）**：Variant-迭代器 For Each 家族 `FE:89`/`FE:8C` 接入既有 `FE:72`/`FE:73` collection 栈机；对象 For-Each `Exit For` `FD:CB` 经既有 exit For 节点建模；`0x56` 只在 For-Each collection feed 时保留 ref（fail-closed）。**文本层 M5-320 1,881 → 1,890/3,121（+9）**、**判分 1,889 → 1,898/3,121（EXACT 1,845）**、M5-67 **548 → 549/732（+1）**、0 pass→fail、分母 3,121/732 不变；W28 归因常量（`W28_CHANGED_FIXTURES`/`V28_CHANGED_SHA`）。pins：`WinOps.bas` **312,626 B / `04C8F0AA…`**、`tools/vb6_decompile.py` **593,463 B / `631B0791…`**、`tools/test_vb6_decompile.py` **283,454 B / `4D003A11…`（**1,107** checks）**、`tests/test_g10_impad.py` **67,426 B / `663A6392…`**、`docs/RE_TOOLS_INTEGRATION.md` LF **834,294 B / `61B3EDFE…`**（§9.130，总 4,807 行）、`re_work/minerAC/judge.py` **22,904 B / `93EB2EAB…`**（untracked）、`re_work/minerAC/aclib.py` **60,464 B / `D0B35B9B…`（不变）**；`CodeTools.cls` 1,707,208 / `65DAEB69…`、`tests/test_g10_udt.py` 69,182 / `88231A14…` 均不变。 |
| **v1.16.32** | 2026-10-28 | **patch 级发布（工具面零变化：add-in 119 -> 119、hub 11 -> 11、DSH 面 130 -> 130；无工具增删、无参数/描述/inputSchema 变化 ⇒ 冻结面不变、不重冻；本批**无 VB 实现源码改动**（`.bas`/`.cls` 字节不变，只有 `Util.bas` 契约版本常量 v1.16.31 → v1.16.32）⇒ 源码印记 I 变（算法仍 `vb6src/3`）、DLL 重新 `/make` 并重签部署）** ① **M5 wave-26（`7f66226`，纯 tools/tests/docs）**：i2/r8/Variant 计数器 For 接入**既有** i4 栈机（8 个 p-code 键：`FE:63`/`0x64`、`FE:6B`/`0x65`、`FE:6E`/`0x69`、`FE:68`/`FE:7E`）+ **FE:7E K 由导出表 3 修正为字节预言 4**（真实 terminal ret 藏在相邻 `00 14` line-marker 对里 ⇒ 循环无出口边；4 fixture 逐字节一致）；fail-closed 双门（同族 init 配对 + Variant 回跳必须精确落已解码指令边界）；**文本层 M5-320 1,869 → 1,881/3,121（+12，全部真实结构行）**、**判分 1,877 → 1,889/3,121（EXACT 1,836 / SEM 53 / DIFF 563 / MISS 609 / UNK 60）**、M5-67 **544 → 548/732（+4）**、0 pass→fail、分母 3,121/732 逐 fixture 不变、7 fixture/+28−25、unexpected=0；+8 仍归因 v1.16.31 卡 A 的 **measurement-side** 映射解移（本批未动 module_vars/映射），**禁止 1,889 单独外报**；4 条 Step-For/For-global **MISSING→DIFF**（冻结 aclib 的 for 渲染不含 Step / 计数器为 module datum ⇒ 判分结构性不可达，0 分损失）。② **docs**：U-C3a 旧口径 → **HOLD** + `check_modal_governance` **226/17 基线**（`docs/TOOLING_MAINTENANCE_20261028.md` 附录 D，本批不修）；`docs/PROJECT_STATUS_20261028.md`（发布链一页式状态）；`docs/PROPERTY_SHEET_DOMAIN_DESIGN_20261028.md`（`#32770` + `SysTabControl32` id=12320 + frame ids 1/2/9/12321；仅 additive 输出键、策略 HOLD、零点击/零消息、10–13 h）；RE 分诊 **§19**（`026170d`，0 条新请求；v1.16.31 窗口的外部并发提交）。pins：`tools/vb6_decompile.py` **589,583 B / `F7F3166D…`**、`tools/test_vb6_decompile.py` **278,458 B / `1E6546FC…`**（**1,092** checks）、`tests/test_g10_impad.py` **64,782 B / `92B27291…`**、`docs/RE_TOOLS_INTEGRATION.md` LF **817,832 B / `92A2D1E8…`**（§9.129 = L4357–4485，总 4,693 行）、`re_work/minerAC/judge.py` **22,904 B / `93EB2EAB…`**（untracked）、`re_work/minerAC/aclib.py` **60,464 B / `D0B35B9B…`（不变）**；`CodeTools.cls` 1,707,208 / `65DAEB69…`、`WinOps.bas` 294,099 / `A27C3D18…`、`tests/test_g10_udt.py` 69,182 / `88231A14…` 均不变。 |
| **v1.16.31** | 2026-10-28 | **patch 级发布（工具面零变化：add-in 119 -> 119、hub 11 -> 11、DSH 面 130 -> 130；无工具增删、无参数/描述/inputSchema 变化 ⇒ 冻结面不变、不重冻；本批**无 VB 实现源码改动**（`.bas`/`.cls` 字节不变，只有 `Util.bas` 契约版本常量 v1.16.30 → v1.16.31）⇒ 源码印记 I 变（算法仍 `vb6src/3`）、DLL 重新 `/make` 并重签部署）** ① **M5 wave-25 批（`319c051`，纯 tools/tests/docs）**：`recognize_select` 两处 fail-closed 边界延伸（`_sel_key` 接受 ref 槽同槽身份；`_chain_arms` 仅放行「单臂 + 空 body」），证据仍是既有 selector temp-copy，0 新增证据/节点/渲染分支；**文本层 M5-320 1,855 → 1,869/3,121（+14）**、M5-67 **544/732 不变**、0 pass→fail、分母不变、End Select MISSING 4→0。② **M5 卡 A（`19b2485`，纯 tools/tests/docs）**：UDT/帧基址 fmem 成员 store/load 用 **base slot** 作归属键（per-proc `(base_slot, off)`）；伪造全局 `g1/g2/g3` 归零，改为 `fmem_member` 语句节点 + 诚实审计注释（`(field name UNKNOWN)`，0 可配对源行）；`module_vars` 伪声明消失 ⇒ **文本层 EXACT ±0（1,869 不变）**、**判分 1,877/3,121（+8）**，+8 **全部**为 judge 的 module-vars 映射解移（MVBASE-SIM 反证回到 1,869）；文本层 EXACT **1,816** / SEM 53 / DIFF 569 / MISS 620 / UNK 63，判分 EXACT **1,824** / SEM 53 / DIFF 559 / MISS 625 / UNK 60，+8 = **measurement-side** 映射解移），**禁止 1,877 单独外报**；5 条原 DIFF 诚实化为 MISSING，0 pass→fail；另修 `recover_declarations` 的 ByRef-UDT 证据门。③ **docs**：OPEN-42 段 C 结案（D1–D4 PASS / U5=HOLD；232 全段 A/B1/B2/C 关闭）+ With 建模预研与 Lead 裁决（aclib shim 被拒、结构框单做不做、卡 C 后置）+ UDT 元数据预研（320 p-code Standard EXE 无字段名 ⇒ 只给偏移 + UNKNOWN）。pins：`tools/vb6_decompile.py` 582,413 B / `0EAB1945…`、`tools/test_vb6_decompile.py` 270,951 B / `66AD3C97…`（1,072 checks）、`tests/test_g10_impad.py` 62,367 B / `0B37FD4D…`、`docs/RE_TOOLS_INTEGRATION.md` LF 799,513 B / `F7078021…`（§9.128 = L4222–4355，总 4,564 行）。 |
| **v1.16.30** | 2026-10-28 | **patch 级发布（工具面零变化：add-in 119 -> 119、hub 11 -> 11、DSH 面 130 -> 130；无工具增删、无参数/描述/inputSchema 变化 ⇒ 冻结面不变、不重冻；本批**无 VB 实现源码改动**（`.bas`/`.cls` 字节不变，只有 `Util.bas` 契约版本常量 v1.16.29 -> v1.16.30）⇒ 源码印记 I 变（仍 `vb6src/3`）、DLL 重签）** ① **M5 wave-24（`d385b71`）**：`Select Case` / `Case Is` 识别门与 W5 temp-copy 证据统一（负控制保留）；M5-320 **1,801 -> 1,806**/3,121、M5-67 **523 -> 528**/732，0 pass→fail，分母不变。② **M5 judge R1（`72c5aa4`，只动 judge，`aclib.py` 不动）**：judge 合成 End 行按**过程种类**拼合法 `End Property`；**新基线 M5-320 1,855/3,121、M5-67 544/732**（+49/+16，0 pass→fail，分母不变，49 行归因）；**这是一次测量口径修正（judge synthetic End spelling），不是产品改进**；旧值 1,806/528、1,801/523 记为 **historical**。③ **OPEN-42 活体记录（docs-only）**：breakpoint-fix1 复验 R1-R8（`f4fc8eb`，全 PASS）、B2（`43d3ef2`）、judge 预研与授权（`7ed3363`/`5fa2efb`）。pins：`tools/vb6_decompile.py` **572,938 / `9611DFEB…`**、`tools/test_vb6_decompile.py` **255,317 / `AA2689EB…`**（**1,033** checks）、`tests/test_g10_impad.py` **59,099 / `0581C3A4…`**、`docs/RE_TOOLS_INTEGRATION.md` LF **762,939 / `CE431343…`**（§9.126 = L4008-4111，总 4,319 行）、`re_work/minerAC/judge.py` **22,904 / `93EB2EAB…`**、`re_work/minerAC/aclib.py` **60,464 / `D0B35B9B…`（不变）**；`CodeTools.cls` 1,707,208 / `65DAEB69…`、`WinOps.bas` 294,099 / `A27C3D18…`、`tests/test_g10_udt.py` 69,182 / `88231A14…` 不变 |
| **v1.16.29** | 2026-10-28 | **patch 级发布（工具面零变化：add-in 119 -> 119、hub 11 -> 11、DSH 面 130 -> 130；无工具增删、无参数/描述/inputSchema 变化 ⇒ 冻结面不变、不重冻；本批含一个 VB 实现源码补丁（`CodeTools.cls`）⇒ 源码印记 I 必变（算法仍 `vb6src/3`）、DLL 必须重新 `/make` 签发）** ① **breakpoint-fix1（`ae2a3a2`，1 文件 / +43 −1）**：修 D-B4a——`breakpoints{action:"probe"}` 的 `phase=observe` 在 **BREAK 态**误报 `lineClass="component_not_found"` + `targetLineText=""`（同一目标在 **DESIGN 态** phase=start 被正确分类为 `statement_line` 且行文本可读）；现在 observe 复用 **start 阶段（design）实测**的分类缓存覆盖 BREAK 解析结果，**回执键集合不变（78 → 78）**、判定仍以 `hitLine` 为准（不放宽）、缓存不可用仍如实 `component_not_found`；`CodeTools.cls` 1,704,109 B / `E758FF20…` → **1,707,208 B / `65DAEB69…`**。② **M5 wave-23（`fe577d5`，纯 tools/tests/docs）**：13 个过程已解码出**非终止** ret 但结构器漏 emit 其块 ⇒ 过程内出口丢失；纯静态 BODY 层后置修复在 pc 顺序位置插同一 `Exit Sub/Function`（复用既有 exit 节点/渲染器；0 新节点类型、0 渲染器分支、0 指令表与 ins_count 变化、0 K 表与 CFG/结构器改动）；M5-320 **1,801/3,121**（EXACT 1,748 / SEM 53 / DIFF 567 / MISS 690 / UNK 63；+13 EXACT、0 pass→fail）、M5-67 **523/732**；pins：`tools/vb6_decompile.py` **572,182 B / `B1D80C13…`**、`tools/test_vb6_decompile.py` **248,075 B / `089A1EC6…`（1,017 checks）**、`tests/test_g10_impad.py` **58,083 B / `2D8F09DE…`**、`docs/RE_TOOLS_INTEGRATION.md` LF **733,804 B / `BAEAD26D…`**（§9.124 = L3796–3912）、`tests/test_g10_udt.py` 69,182 B / `88231A14…`（151）不变、`WinOps.bas` 294,099 B / `A27C3D18…` 不变。③ **工具侧维护（`b1f10ab`）**：`tools/check_breakpoints_honesty.py` **57 checked / 12 failed → 73/0**（按 C-3 模块拆分重定位，断言未弱化），pin **15,891 B / `5D90FC9C…`**；新增 `docs/TOOLING_MAINTENANCE_20261028.md`；`check_modal_governance.py` **226/17 基线记录**（未修）。④ **OPEN-42 活体记录（docs-only，2 条提交）**：`43d3ef2` OPEN-42 232-B2 C1–C7 全 PASS、`a0e4e2a` dialog-fix4 复验 R-1..R-6；新增 `docs/OPEN42_232_B2_20261028.md`、`docs/OPEN42_DIALOG4_VERIFY_20261028.md`（存在/关键 marker 断言，不做尺寸 pin）。⑤ **行为变化**：probe observe 在 BREAK 的 `lineClass`→`statement_line`、`targetLineText`→start 缓存文本（键集合不变、判定不放宽、缓存不可用仍诚实）；`check_breakpoints_honesty.py` 强化到 73/0（工具侧，非产品面）。⑥ 契约 bump + re-emit（同代新 I + DLL 重签）。 |
| **v1.16.28** | 2026-10-28 | **patch 级发布（工具面零变化：add-in 119 -> 119、hub 11 -> 11、DSH 面 130 -> 130；无工具增删、无参数/描述/inputSchema 变化 ⇒ 冻结面不变、不重冻；本批含两个 VB 实现源码补丁（`CodeTools.cls` + `WinOps.bas`）⇒ 源码印记 I 必变（算法仍 `vb6src/3`）、DLL 必须重新 `/make` 签发）** ① **dialog-fix4（`3a6197f`，2 文件 / +54 −5 + +26 −0）**：**D-11** `make_project` 的 IDE 内路线此前在宿主模态框 preflight **之前**先 `save_project` 与 `Reference=` 修复（两次写），回 `modal_blocked`/held 的调用仍重写脏工程（实测 autoSaved=true / autoSaveVia="save_project" / v27a.vbp 371 → 529 B）；现在只读 `HostModalRiskProbe` 先判 hold，hold 时跳过两次写 ⇒ **held 调用零写盘**（autoSaved:false、autoSaveVia:""、autoSaveReceipt 不出现、.vbp/组件 sha256 前后相等；模态框中途消失仍保持 hold=C4）；`CodeTools.cls` 1,700,955 B / `CEF017B6…` → **1,704,109 B / `E758FF20…`**。**D-12** `WoMgRecordDialog` 对同一真实 hwnd 只保留 1 条 dialogs[]（hwnd=0 不作身份；首条优先、后续只补缺字段；不同框不合并）⇒ held 回执不再出现同框 2 条；`WinOps.bas` 292,685 B / `47C83921…` → **294,099 B / `A27C3D18…`**。② **M5 wave-22（`9a0bbe3`，纯 tools/tests/docs，4 文件 / +529 −10）**：25 个过程已解码终止 ret 但结构器未 emit 其块，过程唯一终止符消失、End 行 MISSING；纯静态 post-pass 在 BODY 补同一隐式 End（0 新节点类型 / 0 渲染器分支 / 0 指令表与 ins_count 变化 / 0 K 表与 CFG 改动；fail-closed 前置 5 条 + 负控制 12 条）；M5-320 **1,788/3,121**（+17 EXACT，0 pass→fail）、M5-67 **515/732**；pins：`tools/vb6_decompile.py` **566,574 B / `767BAC3C…`**、`tools/test_vb6_decompile.py` **230,065 B / `5B363BFC…`（973 checks）**、`tests/test_g10_impad.py` 56,960 B / `933D6167…`（125）、`tests/test_g10_udt.py` 69,182 B / `88231A14…`（151）不变、`docs/RE_TOOLS_INTEGRATION.md` LF **714,112 B / `42F11850…`**（§9.123 = L3677–3795）。③ **OPEN-42 活体记录（docs-only，3 条提交）**：OPEN-42 232-A2（`6d845dd`，V1–V9 + Lead 裁决）、OPEN-42 232-B1（`c5ab0ac`，B1–B6 全 PASS）、OPEN-42 174（`e1b726a`，F1–F8 全 PASS）；新增 3 个 docs 只做存在/关键 marker 断言。④ **行为变化**：held/modal_blocked 分支零写盘（autoSaved:false / autoSaveVia:"" / autoSaveReceipt 不出现）；held 回执 dialogs[] 按 hwnd 去重（依赖重复条数计数的旧脚本需同步）。⑤ 契约 bump + re-emit（同代新 I + DLL 重签）。 |
| **v1.16.27** | 2026-10-28 | **patch 级发布（工具面零变化：add-in 119 -> 119、hub 11 -> 11、DSH 面 130 -> 130；无工具增删、无参数/描述/inputSchema 变化 ⇒ 冻结面不变、不重冻；本批含**两个 VB 实现源码补丁**（`CodeTools.cls` + `WinOps.bas`）⇒ 源码印记 I 必变（算法仍 `vb6src/3`）、DLL 必须重新 `/make` 签发）** ① **dialog-fix2（`80fd8d8`，VB 源码，2 文件 / +254 −19）**：修 D-1..D-5——held 分支回执键补全（`dialogs[]` / `awaiting_user` / `howToRecover` / `recovery` / `guaranteedUnsafe`）、unknown 类 #32770 改为 HOLD（不再自动 `click_cancel`）、dry 语义与计划目标分离、入口 pass 的 result/closedNow 早采样、`make_project` 补 `buildState`；`CodeTools.cls` 1,695,000 B / `977B6F85…` → **1,700,955 B / `CEF017B6…`**、`WinOps.bas` 272,396 B / `F389E5EB…`（fix2 中间态）。② **dialog-fix3（`4ea94f3`，仅 `WinOps.bas`，+255 −10）**：修 D-6..D-9——`WO_KIDS_MAX` 80→512 且两趟 `EnumChildWindows` 真计数（新增 `childCountTotal` / `childrenTruncated` / `childLimit`）、dry 计划不再计入 `clicked`（新增 `planned[]` / `plannedCount`，dry `clickedCount=0`）、dry 计划目标禁止 确定/OK 类（`dry_plan_*` 拒绝）、`dialogsAfter` 有界 settle（新增 `afterSettleMs`）；最终 **`WinOps.bas` 292,685 B / `47C83921…`**。③ **M5 wave-21（`f40f195`，纯 tools/tests/docs，5 文件 / +500 −17）**：尾随类型化出口字节被 K=2 挡在 pad 里 → 隐式过程 End 补回；M5-320 **1,771/3,121**（+9 EXACT，0 pass→fail）、M5-67 **507/732**；pins：`tools/vb6_decompile.py` **561,723 B / `C864CF42…`**、`tools/test_vb6_decompile.py` **214,309 B / `FCF21069…`**（**920** checks）、`tests/test_g10_impad.py` **56,960 B / `933D6167…`**、`docs/RE_TOOLS_INTEGRATION.md` LF **692,969 B / `7F40C3E3…`**（§9.122 = L3549–3676）、`tests/test_g10_udt.py` 69,182 B / `88231A14…`（151）不变。④ **OPEN-42 活体记录（docs-only，4 条提交）**：`5015bb9` OPEN-42 popup-governance A1–A7 活体验收 + `9fa4702` Lead 裁决、`d343fa2` OPEN-42 Project Properties 取证 E1–E9、`bd5c9f0` scoping Lead 裁决；`docs/OPEN42_RUNTIME_ERROR_VERIFY_20261027.md` **95,796 B / `AC251E42…` 不变**；新增 `docs/OPEN42_232_A_20261027.md`、`docs/OPEN42_234_20261027.md`、`docs/OPEN42_REMAINING_SCOPING_20261027.md`（均只做存在/关键 marker 断言，不做尺寸 pin）。⑤ **行为变化（外部脚本可能受影响，CHANGELOG 已写）**：unknown 类 #32770 → HOLD（入口 pass/guard/close 一律不点、只报；`preExistingHandled{plannedAction:"hold",clickedId:0}`；`close_active_dialogs` 对 unknown 返回 `refused/skipped reason="unknown_class_hold"`；普通 unknown 消息框不再被自动关闭）；dry 回执 `clicked[]` 恒空、`clickedCount=0`，计划目标改在 `planned[]`（新增 `plannedCount`）；>80 后代框可枚举（上限 512，超限如实 `childCountTotal/childrenTruncated/childLimit`），`dialogsAfter` 改为点击后 ≤1.2 s settle 实测并新增 `afterSettleMs`（dry=0）；`make_project` modal_blocked 分支补 `buildState:"blocked_by_host_modal"`，held 回执补 `dialogs[]/awaiting_user/howToRecover/recovery/guaranteedUnsafe`。⑥ 契约 bump + re-emit（同代新 I + **DLL 重签**）。 |
| **v1.16.26** | 2026-10-27 | **patch 级发布（工具面零变化：add-in 119 -> 119、hub 11 -> 11、DSH 面 130 -> 130；无工具增删、无参数/描述/inputSchema 变化 ⇒ 冻结面不变、不重冻；本批**无 VB 实现源码改动**（`.bas`/`.cls` 自 v1.16.25 起未动；仅 `Util.bas` 契约版本常量 bump）⇒ 源码印记 I 必变（算法仍 `vb6src/3` 未升代）、DLL 必须重新 `/make` 签发）** ① **M5 wave-20（§9.121，实现批，`fb3501d`，5 文件 / +477 −33）**：两处纯静态修复——（a）`OPS` 补 3 行（`0x16/0x17/0x18` → 与 `0x14` 同一 `ret` handler；类型化出口 Integer/Single/Double/Currency 见 `RE_RETSLOT_SPEC` §3.1/A1，shipped consensus K=2，实测 raw 各 3 B），此前渲染成 `[unhandled p-code 17 …]` **占位注释行**并把源码 `End Function` 挤成 MISSING；（b）新增 post-pass `fold_terminal_exit_to_end()`：顶层尾随 `0x13/0x15` 出口（旧渲染成品结尾前一行 `Exit Sub`/`Exit Function`）改写为与 `0x14` 同一 `ret` 节点；两处都只在 `M5_R1_RETHOME` shipped 分支生效（嵌套 exit / `Exit For/Do` / 非尾随 / legacy 分支逐字节不变，0 新节点类型、0 渲染器分支）。收益：M5-320 **1,735/3,121 → 1,762/3,121**（EXACT 1,682→1,709 / SEM 53 / DIFF 567→568 / MISS 756→728 / UNK 63；**分母 3,121 不变**；**0 pass→fail**；27 条 MISSING→EXACT = 22 fold + 5 typed；320 文本动 34 fixture = 40 行 Exit + 5 占位行，unexpected=0）、M5-67 **503/732 → 507/732**（EXACT 475→479，0 pass→fail）；`tools/vb6_decompile.py` 554,531 B / `B5B441CD…` → **557,135 B / `EA1EF37B…`**；`tools/test_vb6_decompile.py` 194,006 B / `AA0BC5C0…` → **202,824 B / `A3C2662B…`**（**877** checks，新增 `check_w20_terminal_exit()` 21 checks）；`tests/test_g10_impad.py` 53,522 B / `AEF41A67…` → **55,819 B / `C2DF6CCE…`**（S0/S1 重 pin `78934453…` / `517ca83e…`、7 个 per-fixture TEXT_SHA 重签、3 个 V4 pin、`W20_CHANGED_FIXTURES` 34，125/0，断言未弱化）；`tests/test_g10_udt.py` 69,182 B / `88231A14…`（151）不变；`docs/RE_TOOLS_INTEGRATION.md` 新增 **§9.121 = L3397–3547**（总 3,756 行，LF 归一 674,049 B / `9B15EEA3…`）。② **OPEN-42 §13 活体复验记录（`99eafbd`，docs-only，2 文件 / +104 −1）**：v1.16.25 在 t16 上 **R1–R8 全 PASS**（host-modal D1/D2/D3 + D1-r：预存在宿主模态框不被自动取消、只读工具不清框、`dry_run` 真 dry、`make_project` `modal_blocked=true` + `leave_alone / gone_before_report`），**R9/D4 未复现（仅观察）**；`docs/OPEN42_RUNTIME_ERROR_VERIFY_20261027.md` 72,089 B / `C9092020…` / 473 行 → **95,796 B / `AC251E42…` / 576 行**；`docs/OPEN_ITEMS.md` OPEN-42 行同步。③ 契约 bump + re-emit（同代新 I；DLL 重签；hub snapshot 因 contractVersion 串重生成 3 行差异，faceFingerprint 不变） |
| **v1.16.25** | 2026-10-27 | **patch 级发布（工具面零变化：add-in 119 -> 119、hub 11 -> 11、DSH 面 130 -> 130；无工具增删、无参数/描述/inputSchema 变化 ⇒ 冻结面不变、不重冻；但本批**含 VB 源变更**（`McpRe.bas`/`CodeTools.cls`/`WinOps.bas`）⇒ 源码印记 I 必变（算法仍 `vb6src/3` 未升代）、DLL 必须重新签发）** ① **OPEN-42 D1-r（`c62f1ee`，3 文件 / +13 −36）**：`McpRe.ReParseErrorDescription` 改为委托 `McpBpCore.H3ParseErrDescription`，删除第二套旧解析（`McpRe.bas` 90,648 B → **89,868 B / `7DEE3BDA…`**）；`error.description` 与顶层 `errDescription` 同源（先剥引号再剥冒号、纯标点拒绝、绝不伪造），纯标点 body 不再回 `"':"`；冻结面键集不变。② **host-modal 治理（`2e6c42d`，2 文件 / +180 −11）**：`CodeTools.cls` 1,688,454 B / `30B64E2F…` → **1,695,000 B / `977B6F85…`**、`WinOps.bas` 267,946 B → **272,396 B / `F389E5EB…`**；入口 pass 对只读工具与 `close_active_dialogs` 跳过（`preSkip`）；`WoModalClearPreexisting` 的 `host_modal_risk_known` → `hold`（只报不点）；`WoModalGuardTick` 不按预存在宿主模态框；`make_project` 预存在分支不再自动 Cancel，回 `modal_blocked=true` + hwnd/title/controls + `modalDismissPolicy=leave_alone / gone_before_report`；`close_active_dialogs` 顶层 `clickedCount/host_modal_risk/dialogsBefore` 与嵌套 `dialogs[]` 同源；`diagnose` 修 VB error 450（对象安全读 `artifactPath`）、清 `logUnavailableReason`、回填 `compilerElapsedMs`。**行为变化**：预存在宿主模态框不再自动取消、只读工具不再顺带清框。③ **M5 wave-19（`b8bb420`，5 文件 / +332 −22）**：44-CT phantom-param pass slot-aware（frame slot 命中参数槽即保留，`resolve_names()` 前不再误删真实 store）——M5-320 **1,724/3,121 → 1,735/3,121**（+11 EXACT，0 pass→fail、11 fail→pass、22 条语句变化 / 16 verdict + 20 text fixture、extras +7/−0；分母不变），M5-67 **499/732 → 503/732**；`tools/vb6_decompile.py` 553,283 B / `E83936D9…` → **554,531 B / `B5B441CD…`**；`tools/test_vb6_decompile.py` 188,269 B / `8A3A7AB5…` → **194,006 B / `AA0BC5C0…`**（**856** checks，新增 `check_w19_phantom_param_slots` 12 checks）；`tests/test_g10_impad.py` 51,777 B / `F985B0FA…` → **53,522 B / `AEF41A67…`**（S0/S1 重 pin `566fd9f8…` / `7f95775c…`、`W19_CHANGED_FIXTURES` 20 fixture、`W19_V4_REPINNED = ()`，125/0）；`tests/test_g10_udt.py` 69,182 B / `88231A14…`（151）不变；`docs/RE_TOOLS_INTEGRATION.md` §9.120 = **L3267–3395**（总 3,604 行；LF 归一 653,822 B / `5608337B…`）。④ **docs**：`docs/OPEN42_RUNTIME_ERROR_VERIFY_20261027.md` 36,790 B / `92C49FFC…` / 298 行 → **72,089 B / `C9092020…` / 473 行**（§11 v1.16.24 活体复测 D1 PASS / D2 核心 PASS + D2-a / 残余 D1-r；§11.8 Lead 裁定；§12 host-modal 活体验收 C2/C3 FAIL + D1–D6 偏差）；`docs/OPEN_ITEMS.md` OPEN-42 行更新。⑤ **发布面**：契约串 v1.16.24 → **v1.16.25**、I 按新源码重算、DLL 重签（工具面/冻结面/catalog 字节不变、**不重冻**）。 |
| **v1.16.24** | 2026-10-27 | **patch 级发布（工具面零变化：add-in 119 -> 119、hub 11 -> 11、DSH 面 130 -> 130；无工具增删、无参数/描述/inputSchema 变化 ⇒ 冻结面不变、不重冻；但本批**含 VB 源变更**（`McpBpCore.bas`/`CodeTools.cls`）⇒ 源码印记 I 必变（算法仍 `vb6src/3` 未升代）、DLL 必须重新签发）** ① **M5 wave-18（§9.119，实现批，`9c613c5`）**：4 文件 / +275 −86；`tools/vb6_decompile.py` 552,189 B / `86445DE3…` → **553,283 B / `E83936D9…`**；`tools/test_vb6_decompile.py` 183,642 B / `4BAEC17C…` → **188,269 B / `8A3A7AB5…`**（**773** checks；`check_w17_fb_fe_k()` 升级为 `check_w18_fb_zero_k()` 21 checks + 2 处既有断言随 K 修正更新）；`tests/test_g10_impad.py` 50,225 B / `83D79F49…` → **51,777 B / `F985B0FA…`**（S0/S1 重 pin `2042c4e0…` / `aeeb2ef4…` + `W18_CHANGED_FIXTURES` 4 fixture + `W18_V4_REPINNED = ()`；断言未弱化，125/0）；`tests/test_g10_udt.py` 69,182 B / `88231A14…`（151）不变；`docs/RE_TOOLS_INTEGRATION.md` §9.119 = **L3180–3266**（总 3,474 行；LF 归一 635,441 B / `B6448A2A…`）。② **四槽 K=0**：EXT `FB_ZERO_K` 覆盖 FC/FD/FE/FF 四槽（三个 handler 实测读 0 操作数字节，= 共识 K2=0/high + 语料字节见证），`apply_fb_zero_k()` 幂等、只碰 4 槽，wave-17 历史名保留同值别名；**FB:FD 渲染器归位为纯 int2str**（删除 44-CT 的 s&amp;CStr(n) 链模型，否则 0x2A 二次消费成重复拼接）。③ **收益（新口径，Lead 预授权）**：M5-320 **1,702/3,120 → 1,724/3,121**（EXACT 1,671 / SEM 53 / DIFF 556 / MISS 778 / UNK 63；**0 pass→fail**、21 fail→pass、55 条语句变化 / 3 verdict + 1 text-only fixture），分母 +1 = f04 新配对语句（**测量口径升级**，旧口径最后值 1,702/3,120 保留历史基线）；**M5-67 499/732 不变**；320 语料文本只动 4 fixture（f01/f04/f06/f10）、**无 fixture 净变差**、**f10 由 −3 反转为 +6**（残留 1 个 FB:FF 无渲染器 fail-closed 占位，不伪造）；M1 30/30、M4 20/20、channels 14/14。④ **OPEN-42 D1/D2（`5db83e5`，3 文件 / +193 −9，产品 patch）**：`McpBpCore.bas` **15,580 B / `EF96A9A2…`**（D1：先剥收尾引号再剥冒号 + 纯标点候选拒绝 + 解析失败 `errParsed.description=false`，**不伪造**；`errSource` 标准框保持 null）；`CodeTools.cls` **1,688,463 B / `5BF91FFA…`**（D2：30 行**纯加法**，仅回填 dismiss 回执缺失的 `errorDialog.buttons[]`（点击前快照），不新增键、不改字段语义）；`docs/OPEN42_RUNTIME_ERROR_VERIFY_20261027.md` 25,504 B → **36,790 B / `92C49FFC…` / 298 行**（§10 修复记录；12 例夹具 **12/12 PASS**：Python 镜像 + t16 真实 VB6 行为；隔离 `/make` **rc=0**（产物 `364CC568…` / 5,324,856 B），注册副作用已恢复；D2 行为验收交本批 t16 活体）。⑤ **契约 bump + re-emit**：源码 I 变、DLL 重新签发；工具面 119 + 11 = 130 零变化、冻结面/目录不变 ⇒ **不重冻**。 |
| **v1.16.23** | 2026-10-27 | **patch 级发布（工具面零变化：add-in 119 -> 119、hub 11 -> 11、DSH 面 130 -> 130；无工具增删、无参数/描述/inputSchema 变化 ⇒ 冻结面不变、不重冻；DLL 因契约版本串变化重建）** ① **M5 wave-17 批（§9.118，实现批，`e9fd3a0`）**：本批 = `e9fd3a0`（4 文件 / +255 −15：`tools/vb6_decompile.py` +36、`docs/RE_TOOLS_INTEGRATION.md` +99、`tools/test_vb6_decompile.py` +71、`tests/test_g10_impad.py` +64/−15）；**`tools/vb6_decompile.py` 549,888 B / `E5D15D3F…` → 552,189 B / `86445DE3…`**；**`tools/test_vb6_decompile.py` 179,499 B / `E8748CDB…` → 183,642 B / `4BAEC17C…`**（**762** checks，新增 `check_w17_fb_fe_k()` 9 checks）；**`tests/test_g10_impad.py` 47,628 B / `827CA82B…` → 50,225 B / `83D79F49…`**（S0/S1 两个 320 聚合 digest 重 pin 为 `de03e7c0…` / `3f52cc18…` + 新增 `W17_CHANGED_FIXTURES`（12 fixture）归因常量与 `W17_V4_REPINNED`（2 fixture）重 pin；断言不弱化，125/0 不变）；`tests/test_g10_udt.py` 69,182 B / `88231A14…`（151）不变；`docs/RE_TOOLS_INTEGRATION.md` §9.118 = **L3081–3179**（总 3,387 行，§9.94–§9.117 原文未动；LF 归一 618,906 B / `8B56AC70…`）。② **证伪 + 根因 + 最小子集**：recensus #4 的 54 `fmem.store.str` / FF:50 **不是指令**，而是出厂 EXT 表把 **FB:FE int2str 的 K 记成 1**（handler 0x6610911B 实读 0 个操作数字节）后**吞掉下一 opcode 字节**造成的 1 字节错位伪影（54 站点中 26 个伪影 base_slot=0x2AFF 非法、FF:50 的 sub 0x50 > 0x46 不可 dispatch，前一条恒为 K=1 的 FB:FE）；只落 **FB:FE K=0**（`FB_FE_ZERO_K` + `apply_fb_fe_operand()` 在 `ps.load_ext()` 后装运行期覆盖，幂等），**FB:FC/FD/FF 保持出厂 K=1（fail-closed，另卡）**；字节见证 = `minerE/f08_params` Main 的 FB FE + `23 54 FF 2A 23 50 FF`。③ **收益（t15 隔离树 `repo_w26` 同窗口两臂）**：M5-320 **1,658 → 1,702/3,120**（**+44**）、M5-67 **480 → 499/732**（**+19**）、**0 pass→fail**；输出 diff **12 fixture**（10 个有 judge verdict 变化）逐条归因；**f10_udt 的 2 条 DIFF→MISSING 如实标注为「对齐修好后暴露的其它未实现 op」，不是已修复**；M1 30/30、M4 20/20、channels 14/14；**wave-18 候选**（4 槽全改 K=0 → 1,722）已登记未落（judge 分母将变 **3,121**，需先登记口径）。④ **OPEN-42 活体验证（`b75c55a`，docs-only，产品零改动）**：t16 活 IDE 上预注册判据 **R1–R9 = 9/9 PASS**；新增 `docs/OPEN42_RUNTIME_ERROR_VERIFY_20261027.md`（**25,504 B / `218B2669…`**，215 行）：错误框三通道识别一致、`last_error{dismiss:true}` 点 **Debug**（`dismissedVia=debug` / control id **4801** / `endClicked=false` / `endButtonPresent=true`，**绝不 End**）、进 **break**（`ebMode=2`）且程序未结束、`get_callstack`/`current_position` 可用、`get_locals` 未给 opt-in 时拒绝（`break_eval_requires_opt_in`）、`read_active_dialogs` 的 `ownerDisabled`/按钮列表与 session-1 独立 Win32 枚举 **5/5 全等**、收车干净且 **8796/8797 常驻 pid 未动**；**发现 D1**：legacy **文本**解析 `errDescription`/`errSource` 有偏差（结构化异常记录通道正确）→ `docs/OPEN_ITEMS.md` OPEN-42 行已标注、**另卡修**。⑤ **本批自身**：contract `v1.16.22` → `v1.16.23`；re-`emit`（算法仍 `vb6src/3`；I 换新）；`bridge/hub_schema.generated.json` 重生成（仅 contractVersion / instructions 三行差异，faceFingerprint 不变；`bridge --selftest` 保持 **129/129**）；**冻结面不变、不重冻** |
| **v1.16.22** | 2026-10-27 | **patch 级发布（工具面零变化：add-in 119 -> 119、hub 11 -> 11、DSH 面 130 -> 130；无工具增删、无参数/描述/inputSchema 变化 ⇒ 冻结面不变、不重冻；DLL 因契约版本串变化重建）** ① **M5 wave-16 C 批（§9.117，实现批，`f892467`）**：本批 = `f892467`（4 文件 / +353 −31：`tools/vb6_decompile.py` +221/−29、`docs/RE_TOOLS_INTEGRATION.md` +79、`tools/test_vb6_decompile.py` +40、`tests/test_g10_impad.py` +13/−2）；**`tools/vb6_decompile.py` 539,659 B / `8046DFA5…` → 549,888 B / `E5D15D3F…`**；**`tools/test_vb6_decompile.py` 177,622 B / `EFB994E3…` → 179,499 B / `E8748CDB…`**（**753** checks，新增 `check_w16_select_arms()` 12 checks）；**`tests/test_g10_impad.py` 46,831 B / `D1B88850…` → 47,628 B / `827CA82B…`**（S0/S1 两个 320 聚合 digest 重 pin 为 `f38509dd…` / `c3bf333f…` + 新增 `W16_CHANGED_FIXTURES`（4 fixture）归因常量；断言不弱化，125/0 不变）；`tests/test_g10_udt.py` 69,182 B / `88231A14…`（151）不变；`docs/RE_TOOLS_INTEGRATION.md` §9.117 = **L3002–3079**（总 3,288 行，§9.94–§9.116 原文未动；LF 归一 600,058 B / `8ED750E7…`）。② **Select Case 范围/区间臂（+19 EXACT）**：`FB:86 between.i4` / `FB:84 between.u8`（K=0；栈 `[value,lo,hi]`，语义 lo<=value<=hi）与**前向** `1D branch.true`（Select 臂真跳；后向 1D = Do/Loop latch 逐字节走 W16 前 raw 路径）；结构器新增 `_jt` term（`_jcc` 镜像）；`recognize_select` 扩展 eq/range/Is 三类臂，并按「前一臂 `jt` + then 体结构相等」合并多值臂；**range 臂是唯一 Select-only 证据门**——纯 `Case Is` 比较链与被编译掉的 `Case Else` 保持 fail-closed。③ **收益（t15 隔离树 `repo_w25` 同窗口两臂）**：M5-320 **1,639 → 1,658/3,120**（EXACT **1,591 → 1,610** / SEM 48 / DIFF 541 / MISS **848 → 829** / UNK 92）、M5-67 **468 → 480/732**（EXACT 442→454 / MISS 188→176）；**0 pass→fail**；**19 条 fail→pass** 全部 Select 结构行；输出 diff **4 fixture / 19 条语句 / extras +0/−31**、unexpected=0；M1 30/30、M4 20/20（19 IDENTICAL + 1 SEMANTIC）、channels 14/14。④ **OPEN-33 P1（`a6ce0a3`，4 文件 / +1992 −3）**：`bridge/vb6mcp_bridge.py` +1,148 行 → **419,584 B / `7B89891A…`**，新增版本化**非 MCP 管理面** `GET|POST /ide/v1/status`（只读，`envelope:"vb6mcp.ide/1"`）与 `/ide/v1/op/<id>` 404 壳；写动词 **policy-before-capability**（默认 `status` ⇒ 403 `ide_control_disabled`；`full` ⇒ 501 `not_implemented`；effective-remote 未开开关 ⇒ 403 `ide_control_remote_forbidden`）；`--ide-control=off|status|full`（`off` = 普通 404、不审计）、`--ide-control-remote`（`full` + 远程三件套缺失 ⇒ 启动拒绝 exit 8）、**45 条封闭错误码**（设计表 41 + 2 复用鉴权词 + 2 additive 路由/兜底；`ide_failure()` 未登记 reason 直接 `KeyError`）、**审计 JSONL**（一请求一行、只记 `tokenId` 永不记 token 值）、**代理对策**（`Forwarded`/`X-Forwarded-For`/`X-Real-IP` + 非本地 `Host` 升级 `peerBasis`，`--trust-proxy` 仅在 raw peer 回环时信任；`proxy_hint` 按远端并脱敏）；**MCP 工具面零变化**（add-in 119 / hub 11 / DSH 130）；新增 `tests/test_ide_control_face.py` **34,584 B / `7595B2F6…` / 27 个 stdlib unittest**；`bridge --selftest` **109 → 129/129**（新增 `S38a–S38s` 20 项）；`docs/OPEN33_DESIGN_20261027.md` 追加 **§14 P1 实现注记**（**84,003 B / `4C8E6E42…`**，927 行）。⑤ **本批自身**：contract `v1.16.21` → `v1.16.22`；仓库内 re-`emit`（算法标识不变，仍 `vb6src/3`；I 换新；`BuildStamp.bas` / `.vbp VersionComments` / `dist/MCP_VB6.srcmanifest.txt` 三处同值）；`bridge/hub_schema.generated.json` 重生成（仅 contractVersion / instructions 三行差异，faceFingerprint 不变）；**冻结面不变、不重冻** |
| **v1.16.21** | 2026-10-27 | **patch 级发布（工具面零变化：add-in 119 -> 119、hub 11 -> 11、DSH 面 130 -> 130；无工具增删、无参数/描述/inputSchema 变化 ⇒ 冻结面不变、不重冻；DLL 因契约版本串变化重建）** ① **M5 wave-15 B 批（§9.116，实现批，`c251ae7`）**：本批 = `c251ae7`（3 文件 / +110 −2：`tools/vb6_decompile.py` +28、`tests/test_g10_impad.py` +15/−2、`docs/RE_TOOLS_INTEGRATION.md` +69）；**`tools/vb6_decompile.py` 537,790 B / `FE7B5442…` → 539,659 B / `8046DFA5…`**；**`tests/test_g10_impad.py` 46,101 B / `0C6D12C0…` → 46,831 B / `D1B88850…`**（S0/S1 两个 320 fixture 聚合 digest 重 pin + 新增 `W15_CHANGED_FIXTURES`（6 fixture）归因常量；断言不弱化，125/0 不变）；`tools/test_vb6_decompile.py` 177,622 B / `EFB994E3…`（741）与 `tests/test_g10_udt.py` 69,182 B / `88231A14…`（151）不变；`docs/RE_TOOLS_INTEGRATION.md` §9.116 = **L2933–3000**（总 3,209 行，§9.94–§9.115 原文未动；LF 归一 584,480 B / `8791ED7F…`）。② **`62 store.darg`（6 站点，+6 EXACT）**：操作数 = u16 frame offset（字节见证 `28 <slot> <val> | 62 0C 00 | 14`；共享 K 表 K=2）；6/6 站点同形 = staged Variant 拷入第一个 ByRef Variant 参数数据区（0x000C），源全为 `v = <lit>`；注册 `_op('62','store.darg','darg_store', size=16)` 并**复用既有 `assign` 节点**（param slot 按正序**提前命名** `p<N>`，因 `sanitize_phantom_targets` 先于 `resolve_names`）；异常形态 ⇒ `stack=[]` + 原样 raw（W15 前逐字节相同）。③ **5F/33/FD:F0/FD:CE 复核后 0 合入（fail-closed）**：5F 改共享 K 会漂移全语料 decode 边界且静态长整型链不闭合；33/FD:F0 的 UDT 成员名 / 定长数组名**不落盘**（§4.3）⇒ 只能伪造命名；FD:CE 需新 `Exit For` 语句节点 + 结构器核心改动。④ **收益（t15 隔离树 `repo_w24` 同窗口两臂）**：M5-320 **1,633 → 1,639/3,120**（EXACT **1,585 → 1,591** / SEM 48 / DIFF 541 / MISS **854 → 848** / UNK 92）、M5-67 **467 → 468/732**（EXACT 441→442 / MISS 189→188）；**0 pass→fail**；输出 diff **6 fixture / 6 条语句**（None → `p1 = 1` / `p1 = 5`）、unexpected=0；6 条 fail→pass 逐条归因；M1 30/30、M4 20/20、channels 14/14（冷跑 13/14 伪影、暖跑权威）。⑤ **OPEN-33 远程 IDE 管理设计稿（docs-only，`97aefff`，775 行）**：`/ide/v1/{status,start,stop,restart,op}` 设计 + **24 条错误码** + 安全边界 + **P0 portproxy 发现**（`8796 → 127.0.0.1:8797` ⇒ peer 地址判据可绕过）+ P1/P2/P3 拆分；**零 MCP 新工具、零 add-in 改动**。⑥ **本批自身**：contract `v1.16.20` → `v1.16.21`；仓库内 re-`emit`（算法标识不变，仍 `vb6src/3`；EOL 归一定义未变 ⇒ 同代新 I）。|
| **v1.16.20** | 2026-10-27 | **patch 级发布（工具面零变化：add-in 119 -> 119、hub 11 -> 11、DSH 面 130 -> 130；无工具增删、无参数/描述/inputSchema 变化 ⇒ 冻结面不变、不重冻；DLL 因契约版本串变化重建）** ① **M5 wave-14 A 批（§9.115，实现批，`1c9d613`）**：本批 = `1c9d613`（4 文件 / +268 −17：`tools/vb6_decompile.py` +151、`tools/test_vb6_decompile.py` +26/−2、`tests/test_g10_impad.py` +31/−2、`docs/RE_TOOLS_INTEGRATION.md` +77）；**`tools/vb6_decompile.py` 529,156 B / `B84084DE…` → 537,790 B / `FE7B5442…`**；**`tests/test_g10_impad.py` 44,367 B / `FEAA5B18…` → 46,101 B / `0C6D12C0…`（S0/S1 聚合 digest + `minerE/f12_multi3` V4 单 fixture pin 重 pin，新增 `W14_CHANGED_FIXTURES`/`W14_V4_REPINNED` 归因常量；断言不弱化，125/0 不变）**；`tools/test_vb6_decompile.py` 176,505 B → **177,622 B / `EFB994E3…`**（`open_only` 两条旧断言替换为 4 条更强 operand-first 断言 ⇒ 门 739 → 741）；`tests/test_g10_udt.py` 69,182 B / `88231A14…`（151 checks）不变；`docs/RE_TOOLS_INTEGRATION.md` §9.115 = **L2856–2931**（总 3,140 行，§9.94–§9.114 原文未动；LF 归一 572,000 B / `4D3A54E7…`）。② **#1 FF:0D `Printer.Print`（14 站点，+8 EXACT）**：操作数 `FF 0D | u16_0 | u16_1`，`u16_1 = argbytes = 8 + Σ(item_width)`（空 Print=8 / 1 串=12 / 2 串=16 / `"a";1;"b"`=20；与 FF:0E/FF:0F 同契约）；栈序 = 值先 push、**接收者对象最后 push** ⇒ 实现先弹接收者再按 argbytes 逐值按宽度弹，**不归零则 fail-closed 回 raw**；接收者身份只从 **predeclared-object 链** `04 ref | 05 push.import.ref | 24 new.if.null.pr | 0D vcall {mid,tag}` 取，`mid ∈ {0x14 App, 0x18 Screen, 0x20 Printer}`（14/14 站点 mid=0x20，源码↔字节校准 lb_global/lb_printfn/lb_printer_ids），表外 mid 一律不命名；渲染复用现有 call 节点 `Printer.Print "a", "b"`（judge/aclib 无需改）+ 工具 `render_text` 极窄「点号 no-parens」分支；6 条 MISS→DIFF（Spc/Tab/分号，**不伪造**）。③ **#2 Open 模式 operand-first（32 站点，+5 EXACT）**：`file_open` 读 **FE:5D 操作数**（VB6 OpenMode 常量 {1:4, 2:17, 4:2, 8:2, 32:6}），合法即 `modeSource='pcode_operand'`（`modeOperand` 可追溯）；表外值（唯一 `0x3001`，`x_file_lock_emit`，非本族形态）与其余形态**回退 44-CF usage-scan**（fail-closed，不猜 mode）；修掉 old handler 的 `Append→Output` / `Random→Binary` / `Input→Binary` 误判。④ **收益（t15 隔离树 `repo_w23` 同窗口两臂）**：M5-320 **1,620 → 1,633/3,120**（EXACT **1,572 → 1,585** / SEM 48 / DIFF 541 / MISS **867 → 854** / UNK 92）、M5-67 **460 → 467/732**（EXACT 434→441 / MISS 191→189）；**0 pass→fail**（statement 级 + fixture 级）；输出 diff **19 fixture / 24 条语句文本变化 + extras +2/−1，unexpected=0**（逐条归因；M4 20/20、channels 14/14、M1 30/30 全绿）。⑤ **U7-S1 翻默认前评估（docs-only，`1fcc126` + `61c2d4b` §7）**：ON 臂 M1 **30/30**、M4 **20/20**、channels **14/14**、M5 **0 delta**、35/35 声明行正确、18 fixture 逐行归因 0 unexpected；**Lead 裁决 D1 = 不翻默认**（`VB6_G10_ARRAY_DECL` 维持 OFF；JSON `moduleData.arrayDeclarations[]` 常开；翻的 4 条硬前置见 §7.2），**不重钉任何 pin**。⑥ **M5 残差 recensus（docs-only，`6c91659`）**：当前 pin 非 consistent **= 1,500**（541 DIFF + 867 MISS + 92 UNK），其中 **243 条（16.2%）是 judge 自带 `not_decompiler_fault`**（`End*` 139 + `Debug.Print` 104，编译器擦除/合成行，不是解码缺陷）⇒ **有效缺陷面 = 1,257**；wave-13 后唯一移动簇 = print 参数（`print_argument_dropped` 12→1），其余 reason/构造/家族逐项不变；wave-14 Top 6 候选见 §5。⑦ **ccmets Lead 裁决（docs-only，`61c2d4b` §6）**：采纳 **R1/R2**（已落地 v1.16.19 §9.114）、**不采纳 R3/R4**、**暂缓 R5/R6**、**R7 未被点名 ⇒ 状态不变（不得写成「不采纳」）**、**R8 维持「无 TypeInfo ⇒ 名字 UNKNOWN」口径**；IAT 静态定名「新增 = 0」登记结案（不再重复评估）；`docs/CCMETA_REVIEW_20261027.md` 追加 §6 后重钉 **32,876 B / `9FAA8594…`**（原 v1.16.19 pin 21,689 B 已 stale）。⑧ **本批自身**：contract `v1.16.19` → `v1.16.20`；仓库内 re-`emit`（**算法标识不变**，仍 `vb6src/3`；EOL 归一定义未变 ⇒ 同代新 I）。|| **v1.16.19** | 2026-10-27 | **patch 级发布（工具面零变化：add-in 119 -> 119、hub 11 -> 11、DSH 面 130 -> 130；无工具增删、无参数/描述/inputSchema 变化 ⇒ 冻结面不变、不重冻；DLL 因契约版本串变化重建）** ① **G10 wave-21 P2 硬化 R1+R2（§9.114，实现批，`ef7861c`）**：本批 = `ef7861c`（3 文件 / +261 −5：`tools/vb6_decompile.py` +55/−5、`tests/test_g10_udt.py` +136、`docs/RE_TOOLS_INTEGRATION.md` +75）；**`tools/vb6_decompile.py` 526,221 B / `683DF7F8…` → 529,156 B / `B84084DE…`**；**`tests/test_g10_udt.py` 62,578 B / `70C68D51…`（139 checks）→ 69,182 B / `88231A14…`（151 checks，+12，append-only、断言不弱化）**；`tests/test_g10_impad.py` 44,367 B / `FEAA5B18…` 不变；`docs/RE_TOOLS_INTEGRATION.md` §9.114 = **L2781–2855**（总 3,063 行，§9.94–§9.113 未动）。② **R1（entry-stub VBHeader 语义校验）**：`_entry_stub_vbheader()` 只判 `bytes[eoff]==0x68` 的旧实现会把**非 VBHeader 的 imm32** 当 VBHeader；AX DLL 入口 = `5A 68 <imm32>`（10/10 imm32 ≠ VBHeader）证明「入口 0x68」不唯一。新实现要求 imm32 通过 `_vb_container_ok()`（`VB5!` + PD→OT→POD 链 + POD↔ObjectInfo 回指 + 模块名可读，与 `VB5!` 扫描同一套），否则 `None` ⇒ 回退 `'VB5!'` 扫描（**U6-P1 AX DLL 回退保留**）。③ **R2（ref 描述符 owner-ObjectInfo 门 C7b）**：C7 只校验 desc ∈ `meta.lpMethods`，而 meta 是候选自述 ⇒ 新门要求 `source='ref'` 且已知 `owner_oi` 时 `u32(d)==owner_oi`，否则 `rejected / C7b / reject_desc_not_owned_by_oi`；**门位在 C0–C7 之后、解码预算之前** ⇒ 既有 gate/reason 不变，本批 C7b 拒绝 **0**；命名偏离申报（C8 已占用为预算门，故 C7b）。④ **零输出变化证据（t15 `repo_w21`）**：320 全文 pin 同 digest / 0 fixture 变化；441 表面 `render_text` **text_changed=0**；R1 普查 441 fixture **mismatch=0**（entry-push-0x68 440 / vb5-magic 1）＋三例 imm32 反例全拒且 `vb` 同值；R2 **1,095/1,095** 到达 C7b 的 ref 槽位全部通过、**C7b 拒绝 0 / 分类变化 0**；M5-67 **460/732**、M5-320 **1,620/3,120** 两臂同值、**0 pass→fail**。⑤ **ccmets 来料复核（docs-only，`f654fdb`）**：`docs/CCMETA_REVIEW_20261027.md`（200 行 / 21,689 B / `6D0A6CA5…`）在 716 文件（563 p-code + 153 非 p-code）上独立重导容器链 ⇒ **C1 PASS 仅限 EXE 706/706**（AX DLL 0/10 不成立）、**C2 PASS 结构面**（1,925 描述符 100% `desc[0]==OI`、0 重叠）+ `+0x12==0x19` **不是常量**（42 模块 6.4% 有 `+0x12==0`）、**C3 PASS**（95/563 文件 104 非描述符槽，对我方风险 0）；**IAT 定名新增 = 0**；采纳 R1/R2、不采纳 R3/R4、暂缓 R5/R6、**R8 维持 UNKNOWN**。⑥ **本批自身**：contract `v1.16.18` → `v1.16.19`；仓库内 re-`emit`（**算法标识不变**，仍 `vb6src/3`；EOL 归一定义未变 ⇒ 同代新 I）。| **v1.16.18** | 2026-10-27 | **patch 级发布（工具面零变化：add-in 119 -> 119、hub 11 -> 11、DSH 面 130 -> 130；无工具增删、无参数/描述/inputSchema 变化 ⇒ 冻结面不变、不重冻；DLL 因契约版本串变化重建）** ① **G10 wave-20 U7-S1 数组声明面（§9.113，实现批）**：本批 = `aefcf71`（3 文件 / +430 −2：`tools/vb6_decompile.py` +206/−2、`tests/test_g10_udt.py` +114、`docs/RE_TOOLS_INTEGRATION.md` +110）；**`tools/vb6_decompile.py` 515,875 B / `9B561B07…` → 526,221 B / `683DF7F8…`**；**`tests/test_g10_udt.py` 55,459 B / `8B43C443…`（126 checks）→ 62,578 B / `70C68D51…`（139 checks，+13，append-only、断言不弱化）**；`tests/test_g10_impad.py` 44,367 B / `FEAA5B18…` 不变；`docs/RE_TOOLS_INTEGRATION.md` §9.113 = **L2671–2779**（总 2,988 行，§9.94–§9.112 未动）。② **声明面规则（窄口、fail-closed、"先证明后声明"）**：静态模块数组（既有 `modules[].staticArrays`）与**模块 datum 的 ReDim**（`kind=datum + module + off`）在简单 VT（VARTYPE 与 elem_size 字节一致：Integer 2 / Long 4 / Single 4 / Double 8 / Currency 8 / Date 8 / String 4（BSTR 指针）/ Boolean 2 / Variant 16）时发 `Public g_<off>() As <T>`；**永不声明维度上界**（模块数组 datum = 4B 前缀 + 内联 SAFEARRAY 描述符、数据指针在运行期 ⇒ 只发 `()`）；UDT 元素（`vt=None + udtFxDesc` / vt 缺失 / vt-elem_size 冲突）**SKIPPED、不发 Type**；datum 若有 scalar fmem/impad/cross 访问证据 ⇒ SKIPPED `datum_also_accessed_as_scalar`；落在 `typeBlocks` whole-UDT witness span 内 ⇒ SKIPPED `inside_whole_udt_witness_span`；同一 datum 多个 ReDim 站点 vt/elem_size 冲突 ⇒ SKIPPED `conflicting_redim_evidence`；**局部 slot 的 ReDim 不是声明面 row**（320 语料 30 个 ReDim 全部是局部 slot、0 个 module datum 目标）。③ **开关纪律**：JSON `moduleData.arrayDeclarations[]` 无条件生成（证据/可测试）；**文本行默认 OFF**（`VB6_G10_ARRAY_DECL=1` 开启，与 U7-S0 的 `VB6_G10_ARRAY_AUDIT` 同纪律）——默认路径 345 fixture 的 `render_text` 逐字节不变、M5 判定面 0 变化。④ **M5 零敏感（本地 + t15 两层同向）**：同窗口两臂（OFF=默认 / ON=开关开启）M5-67 **460/732**、M5-320 **1,620/3,120** 两臂逐值相同；逐 fixture delta=0（67/67、320/320）、**pass→fail=[]**、errors=[]；ON 相对 OFF 的 JSON 只变 `procedures[].moduleDecls`（18 fixture 声明行载体）、顶层 `diagnostics`（20 条）、`decodeBudget.lastHit.steps`（1 条既有 NONDET）；judge 消费键（`module_vars` + `procedures[].name/params/locals/body`）两臂逐字节相同。⑤ **ON 增量实测（320 全量文本 diff，独立脚本 `surface_s1.py`）**：18 fixture 变化（全部在 `minerH`）；声明 **35 行** + `' module` 承载头 **8 行** + 诊断行替换 **18 行**；unified-diff **added 61 / removed 18 = 净增 43 行**；**unexpected=0**；外部 25 fixture 0 变化；预估与实测 `corpus_changed [18,18]` / `decl_lines [35,35]` / `header_lines [8,8]` 全一致。⑥ **独立 VT 表验证（D2 前置②）**：简单 VT 静态数组 **17/17** declared 且类型/宽度一致、UDT 元素 **7/7** SKIP、module datum ReDim **1/1** SKIP；VARTYPE→类型映射 3→Long / 5→Double / 8→String / 12→Variant 各唯一（`table_ambiguous={}`）、mismatch=0；`--selftest` 把 3→Integer 破坏后 comparator 变红（自证伪）。⑦ **OPEN-34 空闲地板决定性复测（docs-only）**：`docs/OPEN34_IDLE_FLOOR_20261027.md`（344 行，`b14ed32`）在 t16 活 IDE 设计态按 §1 预注册判据采测 ⇒ **712 s 有效窗口 / 0 quantum**（VB6 CPU Δ=0、median/p95 `0.0000%`、总量 <0.0022%/核）⇒ 1.5–2.5% 不复现、判低地板；**OPEN-34 → FIXED**。⑧ **backlog 状态同步（docs-only，`9c488f8`）**：OPEN-38（逐子项 8 落地 / U7-S1 在飞 / U9 BLOCKED）、OPEN-37（`r5_final/` 69 文件）、OPEN-41（commit `5a95f74`）。⑨ **本批自身**：contract `v1.16.17` → `v1.16.18`；仓库内 re-`emit`（**算法标识不变**，仍 `vb6src/3`；EOL 归一定义未变 ⇒ 同代新 I）。 || **v1.16.17** | 2026-10-27 | **patch 级发布（工具面零变化：add-in 119 -> 119、hub 11 -> 11、DSH 面 130 -> 130；无工具增删、无参数/描述/inputSchema 变化 ⇒ 冻结面不变、不重冻；DLL 因契约版本串变化重建）** ① **M5 wave-13 W11_other 残余（§9.112，实现批）**：本批 = `baa93a8`（3 文件 / +256 −33：`tools/vb6_decompile.py` 143 行、`tests/test_g10_impad.py` 25 行、`docs/RE_TOOLS_INTEGRATION.md` 121 行）；**`tools/vb6_decompile.py` 511,632 B / `70A43C3A…` → 515,875 B / `9B561B07…`**；**`tests/test_g10_impad.py` 44,367 B / `FEAA5B18…`**（4 个 320 文本 pin 重 pin + 新增 `W13_CHANGED_FIXTURES`（11 fixture）/`W13_V4_REPINNED` 归因常量；**断言未弱化、未删除**，125/0 不变）；`tests/test_g10_udt.py` **55,459 B / `8B43C443…`（126 checks）不变**；`docs/RE_TOOLS_INTEGRATION.md` §9.112 = **L2550–2670**（总 2,878 行，§9.94–§9.111 未动）。② **FF:0E/FF:0F `Print#`/`Write#` 操作数契约**：布局 `FF 0E | u16_0 | u16_1`（FF:0F 同构），**u16_0 = 打印列表描述符（语义未定，保留为节点 `argc` 供 `fold_print_continuations`）**、**u16_1 = argbytes = 8 + Σ(item_width)**（4 / 8，Double·Currency·Date 记 8）；旧 handler 把 u16_0 当 argc ⇒ u16_0=0 时**一个值都不弹**（minerE 10 条 `Print #fnum, s` 值全丢）、u16_0>实际值数时靠 leading-underflow 兜底（多值顺序错位，`Write #f, "a", 1` → `Write #v1, 1, "a"`）；新实现**先弹 channel 再按 need=ab-8 逐值按宽度弹栈**，**仅当 need 恰好归零才采用**（paired），否则**回 legacy u16_0 路径、行为与基线逐字节一致**（fail-closed：不发明操作数、不猜宽度）；paired 节点 additive 增加 `printArgbytes`、未配对节点**逐键不变**。③ **参数拷贝归位**：`print_copy_source()` 仅在「相邻前驱对同一 slot 的 assign + 源值同类型 param/local + 该 slot stores==1 ∧ loads==1」时还原为唯一前驱源表达式（赋值语句保留 ⇒ 渲染等价、语句数不变）。④ **收益（t15 同窗口两臂 w18→w19a）**：M5-320 **1,609 → 1,620/3,120**（EXACT 1561→1572 / SEM 48 / DIFF 552→541 / MISS 867 / UNK 92）、M5-67 **455 → 460/732**（EXACT 429→434 / DIFF 81→76）；**0 pass→fail**（statement 级与 fixture 级）；**11 条语句文本变化全部归因、unexpected=0**（`constructs3/file_modes` S_OutputWrite + minerE `f01_arith`…`f09_onerror`/`f13_form` 共 10 条 `Print #g_0, p1`）；本地同口径 A/B CHANGES=11 / PASS2FAIL=0 与 t15 逐条一致。⑤ **证伪清单**：P2 Nothing/Empty/Null（11）证伪不做（aclib 对任何 str 值加引号 ⇒ 只能伪造节点种类）；P3 date 字面量拼写（6）证伪不做（FA/FE:C5 无源拼写）；P4 import→builtin（11）暂缓（仅 4 站点纯改名）；P5 Replace 默认实参（2）不做（字节不可区分）；P6 For…Step（3）暂缓（非本批 op 契约范围）。⑥ **残余风险**：u16_0 语义未解（续行判据不动、channels `print_semicolon` 14/14）；argbytes 不归零必回 legacy；`lb_print2` 类 FF:0D/别名槽 `FE:B2-B4` **未同步**（RE_M4DECODE_SPEC C6）属已知缺口。⑦ **本批自身**：contract 串 `v1.16.16` → `v1.16.17`，仓库内 re-`emit`（**algo 仍 `vb6src/3`，未升代** ⇒ 同代新 I）。 || **v1.16.16** | 2026-10-27 | **patch 级发布（工具面零变化：add-in 119 -> 119、hub 11 -> 11、DSH 面 130 -> 130；无工具增删、无参数/描述/inputSchema 变化 ⇒ 冻结面不变、不重冻；DLL 因契约版本串变化重建）** ① **G5 R5 收口（docs-only，§11）**：`docs/G5_RUNTIME_SAMPLING_20261027.md` 追加 **§11**（t16 活 IDE、add-in **v1.16.15**（`7A8FB9D2…`）/ IDE pid 27128 只读补测）：**6 类对话框（Find/Replace/References/Components/Add-In Manager/Options）/ 9 次打开 / 23 个状态 / 470 条控件行**（每状态 `--repeat 2` 全 stable）⇒ **页内无子窗口增删**（单页状态变化的结构 diff **added/removed 全 0**、id 稳定、hwnd 全复用；真正新增只有**页级**（Components 每 tab +8 / Options 每页 +18，旧页隐藏不销毁、hwnd 复用）与**对话框级**（Find→Replace 顶层销毁+新建，模板 `res4029`→`res4030`、17→21 子窗口））⇒ **10 条指纹规则**（活动页指纹 / 可见性按 parent 链过滤 / 排除 hwnd·text·enabled·style（含 `BS_DEFPUSHBUTTON` 位）·rect·z_order / 同模板多 tab 必须带状态键 / 何时重采）；`docs/OPEN_ITEMS.md` 的 **OPEN-37 由 PARTIAL → FIXED**（R1–R4 已闭环、R5 由 §11 关闭）；**未改任何产品代码/DLL/注册**。② **G10 wave-12 U5-S1/S2 class/form 实例 datum 图 + 实例布局**（§9.111，`68777e2`）：`tools/vb6_decompile.py` **488,857 B / `F83B01FA…`** → **511,632 B / `70A43C3A…`**；`tests/test_g10_udt.py` 48,366 B / 116 checks → **55,459 B / `8B43C443…`（126 checks，+10，只追加不弱化）**；`classModules[]` 追加 **11 个 additive 键**（`var_area`/`var_area_class`/`instance_tail`/`controls`/`members[]`/`instance_spans[]`/`instance_witnesses{W1c,W2c,W6c}`/`instance_unresolved[]`/`instance_failclosed[]`/`layout_state`/`source`；`g10/module-data/1` **schema 不升版**、旧键语义不变），判据全部 **fail-closed**（作用域门 `POD+0x10==0`、变量区 `[0x34, instance_size-0xC)`、尾 12B UNKNOWN、`08 08 00`+PR-relative `mem.*` datum 编码、W1c/W2c/W6c witness、`members[]` **无 name 键**、form `controls.count` null、D6 `base_slot>8` 永久 legacy）；独立验证器（**不 import 产品**）**11/11 PASS**；444 表面 **文本 0 变化**、JSON changed=39 fixture 全归因 / **unexpected=0**；320 语料文本 **0 变化**、class/form 组 601 statements 语句面 0 变化、**0 pass→fail**（M5-320 **1,609/3,120**、M5-67 **455/732** 不变）。③ 本批自身：contract 串 `v1.16.15` → `v1.16.16`，仓库内 re-`emit`（**算法标识不变**，仍 `vb6src/3`；EOL 归一定义未变 ⇒ **同一代内的新 I**）。 || **v1.16.15** | 2026-10-27 | **patch 级发布（工具面零变化：add-in 119 -> 119、hub 11 -> 11、DSH 面 130 -> 130；无工具增删、无参数/描述/inputSchema 变化 ⇒ 冻结面不变、不重冻；DLL 因契约版本串变化重建）** ① **G5 runtime R1–R5 采样（docs-only）**：新增 `docs/G5_RUNTIME_SAMPLING_20261027.md`；`docs/OPEN_ITEMS.md` 的 **OPEN-37 由 BLOCKED → PARTIAL**（R1 属性表外层 frame ids = comctl32 运行时内存模板（4 个 ids 不在 81 corpus 内、跨 2 进程 3 次开启全稳）；R2 ThunderRT6 控件 id 由 VB 运行时按控件序确定性分配 ⇒ 按窗体域作键；R3 Find `res4029` **15/15** + References `res4040` **14/14** 模板 id==GetDlgCtrlID；R4 三构建 **81/81** 模板 id 多重集相同、0 id 差异；R5 属性页宿主动态创建已复现、**单页内动态控件 NOT-COVERED**；t16 活 IDE **222** 控件行 / 跨进程 / 跨版本证据；**未改任何产品代码/DLL/注册**）。② **G10 fx4 / 44-DC 局部嵌套 UDT 的 Type 分组**（§9.110，`a292d90`）：`tools/vb6_decompile.py` **472,709 B / `2B2487B4…`** → **488,857 B / `F83B01FA…`**；`tests/test_g10_udt.py` 39,493 B / 96 checks → **48,366 B / `E3E458B7…`（116 checks）**；窄口全门 fail-closed（standard 模块 + 帧被 local 槽完整覆盖 + 1 个 0x47 定长串 + 1 个 FF:06 数组 + 精确游标 + Len 常量 witness），嵌套拆分 inferred、数组元素类型仍 UNKNOWN（REQ-G10-F2 不消除）；全库双臂 **changed=2 / unexpected=0**（1,198 文件）、320/67 语料语句 diff **0**（M5-320 **1,609/3,120**、M5-67 **455/732** 不变、**0 pass→fail**）。③ 本批自身：contract 串 `v1.16.14` → `v1.16.15`，仓库内 re-`emit`（**算法标识不变**，仍 `vb6src/3`；EOL 归一定义未变 ⇒ **同一代内的新 I**）。 || **v1.16.14** | 2026-10-27 | **patch 级发布（工具面零变化：add-in 119 -> 119、hub 11 -> 11、DSH 面 130 -> 130；无工具增删、无参数/描述/inputSchema 变化 ⇒ 冻结面不变、不重冻；DLL 因契约版本串变化重建）** ① **G10-F1：FF:06 数组元素访问落表**（§9.109，`c81322f`，工具 **472,709 B / `2B2487B4…`**；FF:06 操作数布局 = **4 B 偏移 + 总长 6 B**，元素类型无法静态确证 ⇒ **fail-closed 输出 UNKNOWN 元素**；`tests/test_g10_udt.py` 39,493 B / 96 checks / 0 fail；REQ-G10-F1/F2 **ANSWERED**；fx4 解锁）。② **M5 wave-12 W5_u8local**（§9.108，`7a7d690`；FC:E0 `load.u8.local` + u8 家族 + **无 Case-Else 的 `Select` 识别**；**+20 EXACT / 0 pass→fail**，M5-320 **1,609/3,120**、M5-67 **455/732**；`tests/test_g10_impad.py` 43,138 B pin 同步）。③ **本批自身**：contract 串 `v1.16.13` → `v1.16.14`，仓库内 re-`emit` 源码印记（**算法标识不变**，仍 `vb6src/3`；EOL 归一定义未变 ⇒ **同一代内的新 I**）。 |
| **v1.16.13** | 2026-10-27 | **patch 级发布（工具面零变化：add-in 119 -> 119、hub 11 -> 11、DSH 面 130 -> 130；无工具增删、无参数/描述/inputSchema 变化 ⇒ 冻结面不变、不重冻；DLL 因契约版本串变化重建）** ① **G10 反编译覆盖批（RE 工具，不属 add-in 工具面）**：wave-10（U2/U3 staticrec 记录语法 + 定长串布局，§9.103，417,939 B）→ wave-11a（U4 `vars_size` fail-closed + U6-P1/P2，§9.104，426,454 B）→ wave-12 impad（S0 默认 ON / S1 `VB6_G10_IMPAD_RESTORE=1` 默认 OFF，§9.105，440,117 B / `026765D7…`）→ **W11-1 classMember V4**（Gate A/B/C fail-closed，§9.106，`27d4fee`，446,709 B / `894CBAC2…`；320 文本 diff 3 fixture / 3 过程 / −3 行 + 1 仅标签，canonical `obj_class_prop` 逐字节，0 pass→fail）→ **W11-2 U7-S0 + U5-S0**（§9.107，`92c66c6`，工具 **460,110 B / `42159E44…`**；`moduleData.arrayAudit[]` additive 默认 ON（320 static 35 / dynamic 30；全库 63 fixture / 92 行）、`modules[].staticArrays` 增 source/vt_name/udt_layout_size/desc_va（redim elem_size 仅审计）；class/form 模块改由新键 `classModules[]` 承载（320：28 模块 / 27 fixture）、删除假 `vars_size`、kind 解析不出 ⇒ null + `CLASS_OR_FORM_UNRESOLVED`，既有 `modules` 语义不变；320 文本逐字节不变（`ba740853…`））；测试：`tests/test_g10_udt.py` 35,010 B（86 checks）、`tests/test_g10_impad.py` 42,402 B（125/0）、`tests/test_g10_classmember.py` 18,197 B（21 tests / 0 fail）。② **M5 wave-9 W1_memload_field_expr：评估 0 合入**（§9.102；三个候选子模式全部证伪，无可证收益，如实记录）。③ **REQ-G10-F1/F2/F3 登记**（`docs/RE_RESEARCH_REQUESTS.md`；F3 = FE:23 K=4 操作数语义，不含计数声明）。④ **v1.16.12 部署/回滚文档刷新（只读，未执行部署）**：`docs/DEPLOY_v11612.md` + `deploy/verify_release.ps1`（泛化只读校验器，默认期望值由 v1.16.11 刷新到 v1.16.12）+ `deploy/README.md`；**本批不刷 v1.16.13 默认值、不新增 DEPLOY_v11613**（部署文档滞后一批）。⑤ **只读报告**：OPEN-38 scoping、G10 U6 验证、U4/U8、cross-module census、classMember 影响面、U7/U5 scoping、impad scoping。⑥ **本批自身**：contract 串 `v1.16.12` → `v1.16.13`，仓库内 re-`emit` 源码印记（**算法标识不变**，仍 `vb6src/3`；EOL 归一定义未变 ⇒ **同一代内的新 I**）。 |
| **v1.16.12** | 2026-10-27 | **patch 级发布（工具面零变化：add-in 119 -> 119、hub 11 -> 11、DSH 面 130 -> 130；无工具增删、无参数/描述/inputSchema 变化 ⇒ 冻结面不变、不重冻；DLL 因契约版本串变化重建）** ① **M5 wave-8 W2（p-code 反编译器 `CreateObject` 变体 `Set` 折叠，不属 add-in 工具面）**：`tools/vb6_decompile.py` = **404,173 B / `0921058D…`**；新增 `FE:4E var.set.var.func` 渲染 + 复用既有 verified `import_map_w4.json` 的 CreateObject 行做「创建对象 → 变体 Set」折叠（fail-closed：仅当紧邻上一条语句是 verified CreateObject 调用且 temp slot 一致才折叠，否则保持 raw placeholder，不新增 inferred 行、不猜名）；实测 **+18 consistent**（M5-320：1,571 → **1,589/3,120**；M5-67：**445/732 不变**），**0 pass→fail**；六门通过（decompile 739/0、M1 30/30、M4 20/20 + 14/14、M5-67、M5-320、320 输出 diff 12 fixture / 25 statement 全部归因、unexpected=0）；逐条证据见 `docs/RE_TOOLS_INTEGRATION.md` §9.101。② **M5 wave-7 judge 对齐（measurement-only，不改产品）**：§9.100 + `tests/test_m5_judge_align.py` / `tests/test_m5_judge_aclib.py`；新 judge pin `72127882…` / aclib `D0B35B9B…`，两份基线副本**只在 `fix/`、不入库**。③ **v1.16.11 部署/回滚文档（只读，未执行部署）**：`docs/DEPLOY_v11611.md` + `deploy/verify_release.ps1`（泛化只读校验器：DLL sha / contract / 工具数 / banner `BUILT` 前缀全部参数化，可校验任意版本与回滚点）+ `deploy/README.md`。④ **本批自身**：contract 串 `v1.16.11` → `v1.16.12`，仓库内 re-`emit` 源码印记（**算法标识不变**，仍 `vb6src/3`；EOL 归一定义未变 ⇒ **同一代内的新 I**）。 |
| **v1.16.11** | 2026-10-27 | **patch 级发布（工具面零变化：add-in 119 -> 119、hub 11 -> 11、DSH 面 130 -> 130；无工具增删、无参数/描述/inputSchema 变化 ⇒ 冻结面不变、不重冻；DLL 因契约版本串变化重建）** ① **M5 wave-5 W11（p-code 反编译器字面量渲染，不属 add-in 工具面）**：`tools/vb6_decompile.py` 新增 3 个低风险渲染组合（Boolean `True`、`Chr$`、Currency/Date 字面量），实测 **+17 consistent**（M5-320：1,569/3,120；M5-67：443/732）；新增 `tools/pcode_data/import_map_w5.json`；逐条证据与 1 条 DIFF→UNKNOWN（无 pass→fail）见 `docs/RE_TOOLS_INTEGRATION.md` §9.98。 ② **M5 wave-6 W8**：仅文档评估（§9.99 折叠/对齐，**0 合入**），风险分析如实记录。 ③ **hub 面 OPEN-28（客户端可见面修复）**：hub 自带一份编译日志分类器（`bridge/vb6mcp_hub.mjs` 的 `compileLog*`），与 VB 产品侧 `Build.bas ClassifyLog` 是两份实现且 hub 侧此前无夹具；本批按 `Build.BetweenQuotes/LineAfter/FileLineFrom` 逐条移植，修掉实测 **3 处分叉**（引号 token 误当 file/line、err-marker 集多出「翻译错误」、裸 `(N)` 兜底路径），新增 `tests/hub_compilelog_parity_test.mjs`（同一 10 夹具经**真实 hub 离线回执**逐字段比对：t16 **PASS=135 / FAIL=0**；未修 hub **FAIL=3** 红基线证明夹具抓得住漂移）。**face 未变**。 ④ **运行时产物去跟踪**：`bridge/hub_face_cache.json` 是 hub 的运行时 face 缓存（时间戳/指纹/实例记录），每次运行都重写 ⇒ 改入 `.gitignore` 并停止跟踪，不再让工作树永久变脏、不污染「与 HEAD 逐字节一致」判定。 ⑤ **文档/夹具卫生**：`docs/OPEN_ITEMS.md` 记 OPEN-28 **FIXED** / OPEN-30 **DECIDED**（可选套件保持手工执行、不进 27 门）；`tests/FIXTURE_SUITE.md` 更正 hub 侧覆盖口径。 ⑥ **部署文档（只读）**：新增 `docs/DEPLOY_v11610.md`（本机部署/回滚手册）、`deploy/verify_v11610.ps1`（只读校验器，**本批未执行任何部署**）、`deploy/README.md`。 ⑦ 本批 **27 门**全绿、`--selftest` 109/109；`FROZEN_SHA` 不变（`37564E39…`）；srcstamp **算法标识不变**（`vb6src/3`，EOL 归一定义未变），仓库内 re-emit 得新 I。 |
| **v1.16.10** | 2026-10-27 | **patch 级发布（工具面零变化：add-in 119 -> 119、hub 11 -> 11、DSH 面 130 -> 130；无工具增删、无参数/描述/inputSchema 变化 ⇒ 冻结面不变、不重冻；DLL 因 srcstamp 算法升代与契约版本串重建）** ① **源码印记升代 `vb6src/2` -> `vb6src/3`**：strict 模式的被哈希字节对文本工程文件做**行尾归一**（CRLF / 裸 CR → LF；二进制伴生件 `.frx/.res/.dsr` 仍原样、源文件不改写）⇒ 按 `docs/SOURCE_STAMP.md` §4/§11.3「算法标识随定义变化」升标识；**旧 I 全部作废**、跨代不可比；仓库内已重新 `emit`（`BuildStamp.bas` 与 `.vbp` 的 `VersionComments` 同步）。详见 `docs/SOURCE_STAMP.md` §13。 ② **p-code 离线工具两处修复（不属 add-in 工具面）**：`tools/pcode_scan.py` 改**有界 DP decode**（挂死修复，全量 3,242 过程 100% 覆盖、旧可终止集逐步等价 0 差异，T-5）、新增 `tools/pcode_k_ext_merge.py`（K 表 ext 键合并器，T-6；`tools/pcode_scan.py`/`tools/pcode_dis.py` 兼容）。 ③ **M5 W3-W4**：`tools/vb6_decompile.py` 新基线 1,552/3,120 + `tools/pcode_data/import_map_w4.json`；W10 step-1 仅记录口径修正（`tools/pcode_data/opcode_semantics_observed.json`，**无行为收益，如实记录**）。 ④ **hub 面 OPEN-27（客户端可见面修复）**：`bridge/vb6mcp_hub.mjs` 的 banner 缓存键改为**按实例身份**（此前可能把别的实例缓存的 banner 当自己的 ⇒ 客户端会看到过期 `CONTRACT` 串）；face 未变。 ⑤ **smallfix 批（OPEN-29/32/43③）**：KPI 的 `tool_blob` 每次从 `tools/*.py` **派生**（84 与旧 blob 降级为显式历史 pin，不再静默漂移）、bridge 非 loopback 拒绝文案的工具数改为派生、`srcstamp` 行尾归一（即 ① 的算法侧）。 ⑥ **gates OPEN-41**：3 项路径敏感断言归一化（parse-only 语料改仓库内；decompile / KPI M2 的 `re_work/minerP/selection.json` 作者绝对路径改按 checkout 映射），每项带**等长篡改自证伪** ⇒ 任意 checkout 布局的红集合一致。 ⑦ **冻结面/文档 pin 的 EOL 口径**：`docs/tools_schema.json` 与 `tools/schema_baseline.json` 的"冻结字节"明确定义为 **CRLF 归一后的字节**（历史 pin `37564E39…`/211,499 B 按该口径成立；仓库 blob 是 LF），契约钉改为「内容 → EOL 归一 → sha256」，等长篡改仍红；`docs/RE_VB6_CONTROL_REFERENCE.md` 的 control-reference pin 同口径。 ⑧ 本批 **27 门**全绿、`--selftest` 109/109；`FROZEN_SHA` 不变（`37564E39…`）。 |
| **v1.16.9** | 2026-10-26 | **回执可判定性批（工具面零变化：add-in 119 -> 119、hub 11 -> 11、DSH 面 130 -> 130；无工具增删、无参数增删、无 inputSchema/描述变化 ⇒ 冻结面不变、不重冻；两条回执 ADDITIVE + 指令/文档口径卫生）** ① **`command_bars` 回执回显实际命中身份**：`action=execute` 成功回执新增 `matched{bar,path,id,caption,enabled,executable,matchRule,uid}`（`matchRule` 为闭集：`id_exact_one_enabled` ／ `caption_first` ／ `caption_prefix_scan`）；`action=list` 的**每一条**目录项（含 popup 子项）新增全局唯一键 **`uid = "<bar> > <id> > <caption> @<pos>"`** —— `pos` 是该控件在所在 bar 内的 1-based 枚举位置（子项为 `"<top>.<sub>"`），`list` 另回 `uidRule` 说明该键（**不透明，勿解析**）。**为什么必须带位置**：实测真 IDE 全量 2,918 条控件中，(bar, id, caption) 三元组有 **259 组重复**（终态；pre-exec 快照 258）、**17 条控件同 id=865** ⇒ 三元组本身不唯一；位置后缀是 `list` 与 `execute` **两侧都能不经二次搜索算出**的唯一判别项，因此同一控件在两侧的 `uid` 必然逐字相同。② **id 歧义拒绝的结构化候选**：`execute` 的 `menu_control_ambiguous` 失败侧新增 `idCandidates` / `idCandidateCount` / `matchRule`（成功侧保留 `idCandidates` 并新增计数）；`reason` 与回执其余形状不变（仍走 `ClassifiedRefusalJson`，仍是 `isError:false` 的结构化拒绝）。③ **`mcp_info.loadedModule`**：新增 `loadedModule{available,path,sha256,shaBasis,shaBasisReliable,hashCapturedAt,moduleHandle,handleBasis,disk{available,path,sha256,hashedAt,basis},mismatch,mismatchNote,basis}` —— `path` 由 `GetModuleHandleW("VB6MCP.dll")` + `GetModuleFileNameW` 从**当前进程**取得（不是部署记录）；`sha256` 的哈希基准点是 **add-in 初始化（`CodeTools.Init` 连接时）**（`Init` 播种路径+哈希并把时间写进 `m_buildShaAtText`；`loadedModule.sha256` 只用这份 memo 并回显 `hashCapturedAt`，**不是「第一次调 `mcp_info` 时现算」**——否则先换盘再调用会错误地相等）⇒ 磁盘 DLL 被替换而 IDE 未重启时**仍是旧值**；`disk.sha256` 是同一路径**此刻**的哈希；两者不同 ⇒ `mismatch:true` + 一句提示。API 失败**或初始化基准缺失** ⇒ `available:false + reason`（后者为 `reason:basis_unavailable`，既不给 sha 也不给 mismatch），不编造。**`build.sha256` / `build.path` 的既有语义未改动**。④ **指令/文档口径卫生（独立审计 FAIL-7/8/12 收口）**：另含一处**文档笔误更正（C-1）**——v1.16.0 冻结台账行把活体 DLL 字节数写成 3,939,828 B，实为 **4,939,828 B**（sha `BB20AD87…` 不变；证据 `fix/CONTRACT_V116_FREEZE_20261022_134809/REPORT.md:5` 与同目录 `S1_mcp_info.txt` 的 `build.onDiskSignature:"4939828|…"`），**无工具面/构建变化**；`initialize.instructions` 固定块的 `CONTRACT v1.16.7 ` 与 `v1.16.7 current` 改为 v1.16.9（源在 `docs/CONTRACTS.md`，hub 快照随 contractVersion 重生成 ⇒ 客户端可见 instructions 与 `contractVersion` 一致）；`skills/vb6mcp/SKILL.md` 的 120/131 当前口径改为 **119 / 130**；FAIL-8 涉及的描述批工作文档仅在文件头加「历史快照」行（不重写正文）。⑤ **冻结面**：本批无描述/参数/schema 变化 ⇒ `docs/tools_schema.json` == `tools/schema_baseline.json` 保持 **37564E39…** / 211,499 B 不变、`docs/TOOL_CATALOG.md` 保持 **D2EC667E…** / 95,714 B 不变；`bridge/hub_schema.generated.json` 仅因 contractVersion/instructions 文本重生成（生成物，非逐字节冻结件）。Util.bas 契约常量 v1.16.8 -> v1.16.9。 |
| **v1.16.8** | 2026-10-26 | **add-in 工具描述批（79 张卡；工具面 add-in 119 -> 119、hub 11 -> 11、DSH 面 130 -> 130；无工具增删、无参数增删、无参数类型/枚举变化）** ① 按实现回读逐条改写 63 个 add-in 工具的 description 与参数说明：删除过量承诺（add_component 声称 form 只在 Add 失败时回退、实际 form 无条件走合成 .frm；control_window 声称能 rename、实际 Window.Caption 只读；breakpoints 说 DESIGN or BREAK、实际 break 态被拒；debug_session 说 env 会传给被启动程序、实际不生效；immediate_eval / immediate_run 承诺的 eval_sink_error / run_sink_error 从不出现），补上此前未记载的真实拒绝与回执键（range_beyond_end / inside_continuation / unencodable_text / project_held_by_ide / menu_control_ambiguous / commandbars_unavailable_while_in_break / vbp_order_restore_failed / autoSaved / caseOnly 等），并修正参数语义（command_bars 的 id 寻址与歧义、write_source 的 encoding=auto、set_control_property 的 Picture/Icon 文件路径、search_in_files 的每行 1 条与 2000 文件上限、read_file 的纯文本返回与 UTF-16BE 缺口等）。② **7 张卡改 inputSchema 的 required（非 breaking）**：ext_enable / ext_disable / ext_uninstall / generate_stub / purity_check 把实现本来就有的必填 name 显式列入 required（required=[] -> required=["name"]）；immediate_eval / immediate_run 把 required 由 ["expression"] / ["code"] 清空为 []（expression 与 expression_file、code 与 code_file 是互斥替代，缺参由 ResolveTextInput 回结构化拒绝）。**没有删除或新增任何参数，没有改参数类型、顺序或枚举**（逐工具断言参数集/顺序/类型/enum 不变）。③ **B 类 6 张**：A-02 / A-03 / A-05 / D-17 采 B2（描述如实记录真实行为与已知缺口，实现修复另立）；C-03 / C-04（ext_enable / ext_disable 对不存在的 pack 名恒回 ok:true 的假成功）**只落 schema 部分**，实现修复仍 BLOCKED-BY-CODE（新增 reason=unknown_extension 属后续契约增量）。④ 冻结面：**必须**在远程机按 dump_tools_schema.py --freeze 重生成 docs/tools_schema.json 与 tools/schema_baseline.json（两者逐字节相同），并刷新 docs/TOOL_CATALOG.md 与 bridge/hub_schema.generated.json。判据：TOOLS=119、CATALOG_TOOLS=130、hub DECLARED=119 / MISMATCH=False、contract v1.16.8。⑤ Util.bas 的唯一契约常量 v1.16.7 -> **v1.16.8**。**FACE_DELTA = NONE（工具数/工具名/参数集不变）；TEXT_DELTA = 63 个工具的 description 与参数说明；SCHEMA_DELTA = 7 个 required 变更。** |
| **v1.16.7** | 2026-10-26 | **GIT 查询能力移除（**REMOVAL**：add-in 120 -> 119、hub 11 -> 11、DSH 面 131 -> 130；删除 1 个工具 `git`；`workspace_status` 的 inputSchema 去掉 `repo` 且描述改写；`todo_status` 不变）** ① **`git` 工具删除**：插件面不再有 `action:status|diff|log` 的只读 git 查询（内部用 `git.exe` 捕获 stdout 的 `ToolGit` 实现一并删除）；原因见 `docs/AGENT_WORKFLOW_POLICY.md` W-1/W-2：大仓库上 `git status` 全树扫描会长时间卡住，而真实工作流里 agent 会反复调用。② **`workspace_status` 去掉 git 半边**：不再接受 `repo`、不再启动任何外部进程，回执只保留 IDE 组件模型 `{dirtyComponents[{name,file,isDirty}], dirtyCount, ideOpenComponents, readOnly}`。③ **`mcp_info.breaking[]` 新增 v1.16.7**（加入 git 工具被移除、`workspace_status` 不再有 git 半边；instead: 删除 git 调用、去掉 repo 参数、改用调用方自有 VCS）；**`migrations[]`** 新增 v1.16.7 迁移说明；**`changelog[]`** 新增 v1.16.7 条目。④ 冻结面**显式重冻**（这是删除，`tools/list` 必然变化）：`docs/tools_schema.json` == `tools/schema_baseline.json`（**119** 工具）、`bridge/hub_schema.generated.json`（contractVersion `v1.16.7` / toolCount **119** / DECLARED=119 / MISMATCH=False）、`docs/TOOL_CATALOG.md` 全部重生成；`initialize.instructions` 的 `release-checklist` 不再提 `git status / git diff`。⑤ `Util.bas` 的唯一契约常量 v1.16.6 → **v1.16.7**。⑥ **同批 `workspace_status` 的 dirty 归因修正（ADDITIVE 回执字段；工具数不变）**：VB6 自动纠正标识符/关键字大小写会把仅大小写规范化的组件标成 `IsDirty`；每个 dirty 组件现带 `paneVsDisk{classification: identical|case_only|whitespace_only|eol_only|real|unknown, firstDiffLine?, caseOnlyLineCount, whitespaceOnlyLineCount, realLineCount, basis, caseFoldingScope:"identifiers_only", truncated}`，顶层新增 `realDirtyCount / caseOnlyCount / whitespaceOnlyCount / identicalCount / eolOnlyCount / unknownCount` 与 `note`（`dirtyCount` 仍是 `IsDirty` 计数、旧 `dirtyComponents` 保留）。口径：逐行比较、bare LF 归一、跳过 header 与隐藏内联 `Attribute` 行；**大小写折叠只作用于字符串字面量/注释之外的 code token**（字面量与注释逐字比较 ⇒ 只改字面量大小写必为 `real`），空白折叠同范围；读盘失败/新未保存组件 ⇒ `unknown`；pane+file > 4,000,000 字符 ⇒ `unknown + truncated:true`。只读：不改窗格、不保存、不写盘。 **FACE_DELTA = REMOVED `git`；CHANGED `workspace_status`（inputSchema 少 1 个属性 + 描述改写）；add-in 119 / hub 11 / DSH 130。** 活体判据：`tools/list` = 119 且不含 `git`；`workspace_status` 回执无 `git` 字段、无参数；`todo_status` 仍在。冻结面实测（v1.16.7 重冻，2026-10-26，t16 活体）：`docs/tools_schema.json` == `tools/schema_baseline.json` = **`03A669F81C1FCB6A05CCAB59A18A843580AFFA80DFC6135ECA6E73D699D1A26B`**（190,891 B / **119** 工具）；`docs/TOOL_CATALOG.md` = `EFF454F0C53B130F5E286DFEB0BACB39EF2815B2C7B42A2B47574F88DA2897D2`（77,857 B / CATALOG_TOOLS=130）；`bridge/hub_schema.generated.json` = `56E6970F01DDE350CD2C996288D3D0181F1C5E531ECCE4D20E89FCB1CA277FA6`（contractVersion v1.16.7 / toolCount 119 / DECLARED=119 / MISMATCH=False / FACE_FP=`65a70480…`）；活体 DLL = `C5F6FB432436C60976D79781257AB7AD63A3578422E5A46757D78E03131F61DD`（5,259,310 B，contract v1.16.7 / toolCount 119）。 |
| **v1.16.6** | 2026-10-26 | **诚实性/口径批（工具面零变化：add-in 120 -> 120、hub 11 -> 11、DSH 面 131 -> 131；无工具增删、无参数增删、无 inputSchema/工具描述变化 ⇒ 冻结面不变、不重冻；只改行为语义与回执字段，全部为 ADDITIVE 字段或「不再谎报」的行为修正）** ① **verify / verify_scope 只在能证明对齐时才报 MISMATCH**：pane 块在盘上按调用方 headerOffset 或内容搜索逐字找到 ⇒ `ideDiskAligned:true`；否则用内容锚点（pane 首/末非空行，分数 = 大小写不敏感相等的非空行数）取**唯一最佳偏移**，证明后在该偏移上算差异；证明不出 ⇒ 新判决 **`ideDisk:"NOT-ALIGNED"`** + `ideDiskAligned:false` + `ideDiskAlignedBasis`（点名原因：分数不足 / 并列 / 工作预算），**不产出行级 diff**，`semanticDiff.available:false`，`overall` = **UNKNOWN**（不是 FAIL），`overallBasis:"ideDiskDiffUnproven"`；`verify_scope` 把它放进 `unknowns[]` 且 `reason:"ide_disk_diff_unproven"`，**绝不进 failures[]**（fail = 已证 IDE≠disk）。已证明对齐时的 MISMATCH 语义、正/反样例与 `ideEqualsDisk` 不变。② **`get_component_code{proc|symbol}` 的 `startLineDisk`/`endLineDisk` 改用累积映射**：旧值 = IDE 行 + **标量** `headerOffset`，在模块体内含隐藏内联 `Attribute` 行的 .cls 上与**同一回执自己打印的行号**（`coords.diskLine` 早已走累积映射）不一致（活体实测差 1 行）；现优先取 `ExactDiskLineOf` 并新增 `startLineDiskExact`/`endLineDiskExact`（false = 标量回退，**不可信**，配 `startLineDiskBasis` 说明）；`startLine`/`endLine` 仍是 **IDE code-pane 行**，不随 `coords` 变化。③ **build cache hit 回执补发布子进程预算组**：`cached:true + path:"cache"` 的回执此前**整组缺席** `timedOut`/`budgetEnforced`/`budgetMs`/`budgetSource`/`childPid`/`childLeftover`；现按结构性不可用发布 `false/false/0/not_applicable/0/false`（与预检拒绝同形），**非 cache 与 timeout 路径逐字节不变**。④ contract v1.16.5 -> v1.16.6（`Util.bas` 唯一常量；instructions 末条/`mcp_info.contract`/每次回执 `mcp{}` 随之自证）。⑤ 本版同时**活体复核并文档化**（未改代码）三项此前报告的「假绿/口径」项：`check_public_api{membersOnly:true, baseline}` **确实比较**（`comparedAgainstBaseline:true` + `counts{}`/`added[]`/`violations[]`；坏 baseline ⇒ `baseline_unusable` + `comparedAgainstBaseline:false`）、`check_component{baseline}` 发布 `baseline_diff{enabled:true, baselineKeyCount, addedCount, removedCount, unchangedCount}`（干净 baseline = 0 key 是合法输入）、按名工具的三类可区分拒绝 `unknown_argument`（含 `unknownArguments[]`/`acceptedArguments[]`）/ `component_not_found`（含 `availableComponents[]` 与大小写不可达说明）/ `target_has_no_project`。⑥ 尾随换行矩阵与返回信封字段口径见文末新增节（**实测**；`fix/TASK_V166_HONESTY_20261026/evidence/`）。 |
| **v1.16.5** | 2026-10-24 | **版本说明补正（工具面零变化；add-in 120 → 120、hub 11 → 11、DSH 面 131 → 131；无工具增删、无参数增删、无 inputSchema/描述变化 ⇒ 冻结面不变、不重冻）** ① 客户端可见的 initialize.instructions 末条由硬编码 v1.11.4 current (110 tools) 改为动态：VB6MCP_CONTRACT_VERSION + CStr(Owner.ToolCount())（不再写死版本与数量）；② 同一 instructions 的第 2 段标题由 BREAKING (read before calling): 改为 HISTORICAL CONTRACT NOTES (oldest first; the authoritative breaking[] is in mcp_info): —— 那 4 条是 v1.6~v1.11 时代的历史条款，不是当前 breaking；③ README.md / README.en.md 新增「hub 侧恢复原语与可靠性开关（v1.13 ~ v1.16）」章节，并把工具面口径修正为 120 + 11 = 131；④ Util.bas 的 VB6MCP_CONTRACT_VERSION 由 v1.16.4 改为 v1.16.5（只改这一个字符串）；⑤ README 中英新增「本地化与英文版 IDE 支持」小节，并把原「英文版可能失效 / 第一优先是本地化」的口径更正为「能力支持、未端到端验证」（只改文档）。判据：mcp_info.contract = v1.16.5、toolCount 仍 120；tools/list 的 name/description/inputSchema 逐字节不变。 |
| **v1.16.4** | 2026-10-23 | **34b 退出阶段写盘声明（ADDITIVE，回执 only）：`quit_ide` 在 requested=true 回执新增 `exitWrites{declaration,observed,condition,evidence,basis}` 声明块；工具面 add-in 120->120、hub 11->11、DSH 面 131->131；无工具增删、无参数增删；1 处 description 变更（`quit_ide`）故冻结面必须活体重冻；另修文档->生成物的合同戳横幅一致性（旧横幅 v1.16.0 -> v1.16.4）。** |
| **v1.16.3** | 2026-10-23 | **34a 结构化错误通道（RE-1）：`last_error` 与 `ide_state.lastError` 新增只读的结构化错误字段（框仍存在）。工具面 ADDITIVE：add-in 120→120、hub 11→11、DSH 面 131→131；无工具增删、无参数增删、无 required 变化；唯 1 个 description 变更（`last_error` 追加 2 段 RE-1 文案）⇒ 冻结面显式重冻（`docs/tools_schema.json` == `tools/schema_baseline.json` = `D2A3EF0DFBCF22D3B8F586FC31266E8C79FED37E6BEE49067B8E6ED4EA7805B4`，190,535 B / 120 工具，两件逐字节相同；`check_schema_baseline: schema_changes=0`；`check_live_tools: live_tool_changes=0 ORDER_MATCH=True`）；`docs/TOOL_CATALOG.md` = `46550B46F7676D71CBA2C01BD3895D5DA7694886D195EB6B7E7AB4578A957C7B`（76,260 B / CATALOG_TOOLS=131）；`bridge/hub_schema.generated.json` 重生成（`contractVersion:"v1.16.3"` / DECLARED=120 / sha `1409F34B90F9967F81907A86399035234EBCD418FF0F7214B6EEEE7B23E75871`）；`Util.bas` 契约常量 v1.16.2→v1.16.3（1 行 / 99,324 B / `743655621D5A41CA6D7E981D6B69C229ED6884CA74828A474662AD08244E0174`→`A5DB10B72E204FC39CE28E6075527B01D32E05806A833FB12A80E6963330C542`）。① **新增字段（只增不改）**：`err{number,description,source,hresult,numberBasis}`、`errBasis`（封闭四值 `vb6map|dialog|none|unavailable`）、`errChannel{readOk,recordUsed,selector{raw010,dword074,bit24,chosen,basis},fields,bstrReads[{offset,role,error,chars,bytes}],why}`、`errCrossCheck{available,number,description,relation,note}`、`errLocation{line,column,file,basis,why}`；既有键（`errorDialog` / `errNumber` / `errDescription` / `errSource` / `errParsed` / `callstack` / `localsNames` / `position` / `stopped_at` / `howToRecover`）语义不变（ADDITIVE；缺省参数与动作路径全部不变）。② **数据来源**：零 VBA6 导出调用 —— 纯进程内 `GetModuleHandleW("VBA6.DLL")` + `CopyMemory` 读**异常记录**（record A RVA `0x17D078` / record B `0x17D0A8`；选择子 = `dword_FC0D074 & 0x1000000`，不可读时回退为对两条记录各自按空判定 = `basis:"fallback_probe"`；`+0x04`=Source BSTR、`+0x08`=Description BSTR、`+0x1C`=HRESULT，facility 命中 `(hr & 0x1FFF0000)==0x000A0000` 时低 16 位即 VB errnum，否则 `number:null + numberBasis:"facility_not_vb"`），**不调用 `EbExceptionHandled`**（实测它会清 `dword_FC0D010` 并吞掉错误状态）。③ **诚实边界（禁止误读）**：本批不抑制、不自动关闭、不无框读取错误框 —— 框仍然存在且仍然模态；`last_error{dismiss:true}` 沿用既有按钮投递通道；**自动关框的前置是 BREAK→DESIGN 恢复通道，本批不做**；行/列/文件位置仍不可得（`errLocation.basis:"not_attempted"`，需虚拟调用，超出本批）；`+0x0C`（默认帮助文件 `VBENLR98.CHM`）/ `+0x10`（=1005555）/ `+0x18` 不发布。④ **失败显式化（不伪造）**：设计态/无框 → `ok:false + stopped_at:"no_error_dialog" + errBasis:"none" + err.number:null + errChannel.readOk:false + errChannel.why`（具名四值 `vba6_not_loaded|selector_unreadable|record_unreadable|record_empty`）；**绝不用 `Err` 对象或猜测值冒充记录值**，`errCrossCheck` 只作对照（add-in 线程 `Err` 在处置入口快照，`relation:"addinThreadErrAtHandlerEntry"`，且仅当可得）。⑤ **证据/验收**：跨进程 2/2 命中异常记录映射（G2 探针 `fix/TASK_RE_ERROR_CHANNEL_20261023/` 的 fixture + probe 与手测回执 JSON，按 receiptSha 对账）；**产品代码不含任何可切换的假偏移/假记录开关（C10）**；完整验收、battery 覆盖缺口披露与回执索引见 `fix/TASK_RE_ERROR_CHANNEL_20261023/RECORD_34a.md`。 ⑥ **errChannel.selector.basis="unavailable" 的语义（Lead 裁定）**：它表示**选择子未参与选择**（record A 在 modal 与 settled 两种状态下均为当前记录；实测 bit24 只跟随 modal 对话框状态、不是 record 选择子，选择子未参与），**不是读不到/读取失败**；A5 枚举无“record A 恒选”项，故用 unavailable 表示“未据此判据选择”，**未新增枚举值、未改 schema 形状**（last_error 描述已同步这句）。 |
| **v1.16.2** | 2026-10-23 | **33b 语义批：dirty-untitled 关闭路径默认拒绝 + 显式 opt-in + `open_project` 成功面归一 + 断点台账落点固定。工具面 ADDITIVE：add-in 120→120、hub 11→11、DSH 面 131→131；`inputSchema` 恰 +2 个可选参数（`open_project.allowHostWrite`、`remove_project.allowHostWrite`），**无工具增删、无 required 变化、无既有参数类型变化**；另有 2 个工具 description 与 1 个参数 description（`remove_project.allowDirty`）文本变更 ⇒ 冻结面**显式重冻**（`docs/tools_schema.json` == `tools/schema_baseline.json` = `0F0CC82D0E72D7CBB99255FBD5C31255EC4298D44BEF28159F3C2DFE6B50560B`，120 工具；`check_live_tools: live_tool_changes=0`）。① **S1 脏且无标题（dirty ∧ untitled）的关闭路径默认拒绝、零写盘**：工程**无盘上 `.vbp`**（`VBProject.FileName` 为空）且 `VBProject.Saved=False`（或组件 `IsDirty=True`）时，VB6 在关闭（`VBProjects.Remove` / `open_project{exclusive:true}` 的换工程路径）时会把未保存工程**写进 IDE 当前目录**（33a 实测：`工程1.vbp` + `工程1.vbw` + 组件文件，并覆盖同名 `Form1.frm`），而 `VBProject.Saved` 在 VBIDE 里**只读** ⇒ 无法抑制。现默认**拒绝且零写盘**：`ok:false + reason:"dirty_untitled_would_write" + tool + projectName + fileName:"" + untitled:true + dirtySource + evidenceSource + dirtyUnknown + componentProbe + dirtyComponents[{name,isDirty,lines?}]（**组件级不可读时为 `null`**，此时 `dirtyUnknown:true` + `componentProbe:"unreadable"` + `componentSavedReadErrors[]` 共同说明，**绝不**以空数组冒充「没有脏组件」）+ componentSavedReadErrors[] + cwdCandidateFiles[]（CWD 深度 1 的现有文件 = 宿主可能覆盖的同名文件）+ cwdCandidateFilesScanned/Truncated/ScanFailed + **zeroWriteBasis + zeroWriteScanPre/Post**（33b.1：拒绝回执装配完成后**再扫一次** IDE 当前目录，`pre_cwd_scan_equals_post_cwd_scan` = 零写盘自证）+ why + howToProceed + note`。opt-in = 两工具各自的 **`allowHostWrite`**（boolean，默认 false）置 true 才恢复 3a 行为，并**如实报告**写入：`hostWrites[] + hostWritesBasis + hostWritesTruncated + hostWritesScanned` 与 `sideEffects.filesWritten[]`。`open_project` 的闸门只在 `exclusive:true` 的关闭路径、且只判定**将被关闭的其它工程**；`remove_project` 的闸门在 `allowDirty:true`（已选择丢弃改动）**之后**追加 ⇒ `allowDirty:false` 对**有** `.vbp` 的脏工程**沿用旧 reason** `unsaved_changes_would_be_discarded`，**不是** `dirty_untitled_would_write`（E5 回归保持）。★**实测（2026-10-23，srcBB→srcBC 手测 T1/T1b/T3）**：组件级 dirty 必须读 **`VBComponent.IsDirty`** —— **`VBComponent.Saved` 不是该接口的成员**，本机对每个组件都抛 **438（对象不支持该属性或方法）**，这正是 33b 初版 `dirtyComponents:[]` + `dirtyUnknown:false` 的根因（同一作用域内 `VBProject.Saved` 是合法成员，两者不可混用）。33b.1 起 `open_project{exclusive:true}` 闸门与 `remove_project` 共用**同一个** `ComponentDirtyProbe`（读 `IsDirty` 并附 `lines`），实测输出形如 `[{name:"UntitledDirty1",isDirty:true,lines:6}]`；组件级**读不出**时两条路径**一致**给 `dirtyComponents:null` + `dirtyUnknown:true` + `componentProbe:"unreadable"` + `componentSavedReadErrors[{name,errorNumber,errorDescription}]`。**拦截判定只依赖项目级证据**（`VBProject.Saved=false` ∧ 无盘上 `.vbp`），与组件列表是否可读无关。② **S2 `open_project` 成功面归一**：此前 `AddFromFile` 抛 IDE API 错误但工程实际已打开时会回 `ok:false + apiError`（33a 判为诚实性缺陷，3/3 会话复现）。现**打开成功即 `ok:true`**，非致命 IDE API 噪音只进 `apiWarnings[]{code,message,raw,phase,spurious}` + `note`，**不再占用 `apiError`**（`apiError` 只表示真的没打开；真正的失败仍抛错）。③ **S3 断点台账落点固定**：`breakpoints{action:"set"}` 的注入台账默认写 **`%TEMP%\vb6mcp\breakpoints.json`**（不再落在工程 CWD / 深度 2），可用环境变量 **`VB6MCP_BP_REG_FILE`** 覆盖；实际写入路径必须出现在回执 `sideEffects.filesWritten[]`（写失败或未写则不列，不发明路径）。★**实测（2026-10-23，srcBB→srcBC 手测 T6/T7）**：`BP_REG_DIR` 由 33b 开发期的 **`.vb6mcp`** 统一为 **`vb6mcp`**（与 `%TEMP%\vb6mcp\instances`、与 `prune` 回执 note 的写法一致），旧位置 `%TEMP%\.vb6mcp` **不再读写**；实测 `set` 把 `%TEMP%\vb6mcp\breakpoints.json` 列进 `sideEffects.filesWritten[]`，`VB6MCP_BP_REG_FILE` 覆盖**实测有效**（T7：写入指定文件并在 `filesWritten[]` 列出，默认路径 mtime 不变）；**工程 CWD 及其父目录均不出现 `.vb6mcp`/`vb6mcp` 目录**。④ **验收（33b.1 重测，2026-10-23，srcBC）**：`docs/tools_schema.json` == `tools/schema_baseline.json` = `8AE015325F577466CE9E4A0CFB36AEA11C5240FC46BD4B2D3659CFC08D19951A`（189,227 B / 120 工具；`check_live_tools: live_tool_changes=0`）；`bridge/hub_schema.generated.json` = `86BE1C1B5EA4957DB847B9F12EEFA04887230539E14E59169B6065E8B92B9AF4`（VERSION=v1.16.2 / DECLARED=120 / MISMATCH=False / FACEFP=`5b37d606…`；`LIVE_FACE_FINGERPRINT … FRESH`、`hub_alignment_failures=0`）；`docs/TOOL_CATALOG.md` = `E3E65030A986C0F49D1208A66FC2C32A74096A3575CEF96CA2798030E80FA173`；构建产物 `_vb6build\srcBC\VB6MCP.dll` = 5,152,814 B / `4359CDB1632725614C931E2DB4265E8EE9F931A9EDCD52BAE3DC1B1FB8E79922`（`/make` 日志 `'VB6MCP.dll' 编译成功`、errorCount=0）；`SOURCE_HASH=5c17fba90928f5f9ce8e5b36e67cf8907966348c450d1c8d94899c42a680dff0`（FILES=54）；`check_skill_truth` 4/4、`check_schema_baseline: schema_changes=0`；gates 12/12（×3）。手测（srcBC，disposable IDE + 诱饵/全清单口径）：**T1/T1b/T3 PASS**（`dirtyComponents:["UntitledDirty1"]` + `dirtyUnknown:false` + `componentProbe:"readable"` + `zeroWriteBasis:"pre_cwd_scan_equals_post_cwd_scan"` + 零写盘）、**T2/T4/T5 PASS**、**T6 PASS**（ledger 落 `%TEMP%\vb6mcp\breakpoints.json` 且在 `sideEffects.filesWritten[]`，旧 `%TEMP%\.vb6mcp` 不再被触碰，CWD/父目录无 `.vb6mcp`/`vb6mcp`）、**T7 PASS**（`VB6MCP_BP_REG_FILE` 覆盖写到指定文件、列进 `filesWritten`、默认 ledger mtime 不变）。 |
| **v1.16.1** | 2026-10-22 | **工具面零变化（add-in 120 / hub 11 / DSH 131）；契约版本号 bump v1.16.0→v1.16.1；`add_component` 回执**只增不改**：新增 `fileSyncSkipped` / `fileSyncSkippedDetail` / `failedAt` / `adoptedAfterError` / `adoptedVbError` / `adoptedNote`，新增失败原因 `add_component_failed`（旧 `add_component_vbp_failed` 保留，仅用于 component_file/vbp_sync 阶段），并修正 `vbError` 语义（旧值恒为 `"0:"`，现为真实 `Number:Description`）；`tools_schema.json` / `schema_baseline.json` **未变**（`check_schema_baseline: CHANGED=none`，不需重冻）。功能侧：静息 CPU（私有拓扑 300 s）0.281% 总/0.231% 稳态/0-of-30（改造前 1.797%/1.781%）；默认宿主 tick 由 VB6 Timer 控件改为 Win32 `TIMERPROC`（长调用期间 progress 帧实测 3–15 帧）；bridge `stamp_body` 对数组/标量 body 改为单文档 `{singleDocument:true,data,receipt}`（`find_gui`/`vba6_probe_plan`/`audit_tail` 修复）；hub 面检测改为注册表文件轮询。基线：`srcZ` = 5,083,181 B / sha256 `c48962e6…` / `SOURCE_HASH=a67e873c…` / FILES=42 / gates 12/12 / HEAD `6a83646`。验收：battery 92/0、tick 臂 5/5、私有静息按预先注册判据 PASS。 |
| **v1.16.0** | 2026-10-22 | **add-in 面：断点行为验证（两阶段 probe）+ RUN 态 End 去伪拒绝 + 编译器权威开关 + 多组件批量编辑；工具数不变（add-in 120、hub 11、DSH 面 131），全部 ADDITIVE（缺省参数/动作路径不变）；本版已重冻冻结面** ① **`breakpoints{action:"probe"}`（行为验证 —— 唯一能回答「这个断点有效吗」的通道）必须两阶段**：`phase:"start"`（默认）判定 design 态 / 无模态框 / 可断行后**启动程序**（**不快照、不保存工程、不写源码**），回执 `probePhase:"start" + probeResult:"started_pending" + probeId + startedPending:true + startQueued:true + howToContinue`（Start 在单线程 STA IDE 里是**排队命令**，程序在本调用返回**之后**才真正开跑）；`phase:"observe"`（带同一 `probeId`）等待至多 `timeout_ms`，命中后读 `position{component,proc,paneLine}` + `callstack`，再 **End** 程序并按**回读**证明 `stateAfter=design`；End 收尾可能需再确认一次 —— 尚未观测到时回执 `endPending:true`（不谎报已停），直到回读到 `stateAfter=design`。`phase:"auto"` 是 `start` 的显式别名（**不是**一次调用跑完两段：一次调用无法观测它自己启动的运行）。`probeResult` 是**封闭枚举**：`started_pending | stopped_at_target | stopped_elsewhere | did_not_stop | timeout | did_not_stop_unobserved | line_not_breakable | not_attempted | run_failed`。`probeId` **单次使用**且有 TTL（600000 ms）；未知/过期 ⇒ `ok:false + reason:"probe_not_found" + howToProceed`。★代价：它**真的跑程序**（observe 命中 = 等停 ≤timeout + 可能一段确认；实测 6.2 s / 12.1 s），且**不做 pre-run flush**。 ② **`control_execution{mode:"end"|"stop"}` 在 RUN 态不再假拒绝**：旧的 `reason:"commandbars_unavailable_while_running"` 对 end/stop 是**伪拒绝**（实测 End 命令条在 RUN 态返回 true 且真停机）；现走 `runStateEndPath:"dbclick_end_while_running"`，回执新增 `endCommandDelivered / endPending / stateAfterVerifiedBy:"ide_state" / stateAfterReadBack / programEnded / deliveryVsStateNote / endSettleMs` —— **`programEnded:true` 只在 design 回读之后出现**；已发出但尚不可观测时回 `ok:false + endPending:true + howToContinue`。**其它 mode 的 RUN 态拒绝不变**。 ③ **`verify` / `verify_scope` 新增可选 `requireCompiler`（boolean，默认 false）**：false 时回执**逐字节不变**；true 时取**编译器**证据（`authority:"compiler"` + `errorCount` + `artifactChanged` + `compiler{path,logPath,errorCount,artifactChanged}`；`verify_scope` 取得的是**一次项目级**编译判定），拿不到编译器 ⇒ **结构化拒绝** `ok:false + reason:"compiler_not_configured" + compilerProbe{resolved,path,why} + howToProceed`，**绝不用 lint 顶替编译**（该路径不保存工程、不碰真产物）。 ④ **`apply_edits` / `preview_edit` 多组件批量（同一工具，不新增工具）**：`edits[].component`（工具级 `name` 变**选填**；两者都缺 ⇒ `reason:"component_required" + opIndex`）；`expected_shas[]`（`[{component,sha256}]`，任一不符 ⇒ **整批拒绝、一字节不写**，回 `guards[]`）；**顺序语义** = 按 `edits[]` 顺序，每条的锚点在**该组件已更新文本**上解析（前面的 op 可移动后面的锚点），组件之间互不影响；**全或无**：任一失败即把**已写组件**恢复到批前 pane 文本并用**回读**证明（`rolledBack / restoreVerified / restoredComponents[]`），回 `ok:false + failedOpIndex`（**全局**下标）+ `hits[]`；上限 `max_ops` 默认 **200 / 上限 1000**、单批最多 **40 个组件** ⇒ 超限 `reason:"batch_too_large" + limits{ops,components,max_ops}` 且不写；成功回执新增 `appliedOps / totalOps / components[] / batchElapsedMs`；`preview_edit` 用**同一** `edits[]` 形状，按组件分组 hunks（`components[{component,ops,hunks,hunkCount,algorithm}] + totalHunks`）且**仍然什么都不写**（`writes:false`）。 ⑤ **工具数量不变**：add-in **120 → 120**、hub **11 → 11**、DSH 面 **131 → 131**；ADDED/REMOVED = 0；**本版已重冻**（`dump_tools_schema --freeze` + `gen_hub_snapshot` + `export_tool_catalog`；活体车道 = 最终 DLL `BB20AD87…`，contract v1.16.0 / toolCount 120）。 ⑥ 本版一并落冻三处**此前已声明但未进冻结面**的纯描述变更（`immediate_eval` / `immediate_run` 的 sink 失败分类措辞；`run_project` 的 STA 排队说明 + `diag{}`），无参数变化。 **FACE_DELTA = ADDED 7 optional params**（breakpoints.phase / breakpoints.probeId / apply_edits.expected_shas / apply_edits.max_ops / preview_edit.max_ops / verify.requireCompiler / verify_scope.requireCompiler）+ **breakpoints.action 枚举 +probe**；**CHANGED descriptions ×9**（apply_edits / breakpoints / control_execution / immediate_eval / immediate_run / preview_edit / run_project / verify / verify_scope）；`apply_edits`/`preview_edit` 的 `required` 由 `["name","edits"]` 变为 `["edits"]`。 |
| **v1.16.0+冻结面台账补登（hub 面哈希 + 快照语义）** | 2026-10-22 | **只改台账文字：登记 v1.16.0 重冻后的实测哈希，并把 `bridge/hub_schema.generated.json` 的冻结语义写清（防止下一次漂移被误判）。工具面零变化（add-in 120→120、hub 11→11、DSH 面 131→131），不新增版本、不改任何散列值、不动冻结件。** **实测（本机工作区文件，Get-FileHash -Algorithm SHA256，2026-10-22）**：① `docs/tools_schema.json` == `tools/schema_baseline.json` = **`7FCC00AE6B6C8CE5BB9A08E64FA0D35B6A228E50EA26482D631DB260FD645355`**（120 工具；两份逐字节相等）；② `bridge/hub_schema.generated.json` = **`D95A85F6678838C640CCF7F1117D8B60102AD32296181DAFCDA2AE7BF6F0A250`**（VERSION=v1.16.0 / DECLARED=120 / MISMATCH=False）；③ `docs/TOOL_CATALOG.md` = **`3D6A3CC1E246227BC72C65BD8E42F7DA39A80950ACBF92C39A23CFAF2566154A`**；④ 活体 DLL = **`BB20AD876C152260C61A6B1953A5BC21A66B40EF98AE0181721C1AC66592A94C`**（4,939,828 B；contract v1.16.0 / toolCount 120；活体自证见该版冻结执行报告）。**冻结语义（关键，必须按此判读）**：**(a) 逐字节冻结件 = `docs/tools_schema.json` 与 `tools/schema_baseline.json` 这一对** —— 两者必须逐字节相等，任何工具面变化只能走 freeze 流程；**(b) `bridge/hub_schema.generated.json` 是「生成物 / 缓存」，不是逐字节冻结件**：唯一生产者 `tools/gen_hub_snapshot.py`（输入 = `docs/tools_schema.json` + `docs/CONTRACTS.md` 的版本表与 instructions 文本块）；它的正确性判据是**关系**而不是自身 sha —— ① `faceFingerprint` 必须等于由当前 `docs/tools_schema.json` 重算的指纹（`tools/check_hub_alignment.py` ⇒ `LIVE_FACE_FINGERPRINT … FRESH`），② `tools[]` 的 name/description/inputSchema 与冻结面**逐字段、同顺序**相等，③ hub-only 名称不冲突；⇒ **重新生成必然改变它自己的 sha（这是生成物的预期行为）**，上面登记的 `D95A85F6…` 只是「本版重冻那一次」的证据锚点，**不是**要求逐字节保持不变的目标；**不得**为保持旧 sha 而拒绝运行 `gen_hub_snapshot.py`（那样只会让快照 STALE）。**(c) 历史补记**：提交 `82129c2`（live 面追加 ROUTING/ENCODING 指引，+14/−2）曾使该生成物在 git blob（LF 计）下为 `2E1277073A3281B8…` 且**未新增版本行 / 未登记哈希**；该值已被 v1.16.0 重冻取代（现 `D95A85F6…`），工具面自 v1.11.4 起零变化（`tools/check_schema_baseline.py` ⇒ `BASELINE=120 CURRENT=120`、`schema_changes=0`）。**(d) 生效边界**：本行登记值必须与 `Get-FileHash` 实测**逐字符相等**；哈希值本身不会变，变的只是「哪一次重冻」。 |
| **v1.15.0** | 2026-10-22 | **add-in 面参数变更：`run_project` +2 可选参数、`get_locals` +1 可选参数；add-in 面 120→120、hub 11→11、DSH 面 131→131；冻结面待活体重冻（本版本未重冻）** ① **`run_project.observe_max_wait_ms`**（可选，integer，毫秒，默认 `timeout_ms`，**clamp 1000..120000**）：observe 预算。到点即**返回**（**ok:true**，超时**不是失败**）+ `state:"timeout" + timeout:true + retrySafe:true + suggestedAction:"observe_again" + observedState + elapsedMs`（elapsedMs 为**实测**值）+ `marker.unmetReason="budget reached before the marker appeared"`（请求了 marker 但未命中时）+ `pipeReleasedByReturn:true` —— 单飞门不被长期持有（单线程 STA 不能抢占已执行的 VB 语句；"返回超时"是调用方决定，不是杀语句）。 ② **`run_project.done_marker`**（可选，object `{path, contains}`）：`path` 是 **IDE 所在机器**的路径（远端时由 middle service 侧判读；允许 UNC），只读**前 65536 字节**、不写任何文件；`contains` **大小写敏感**，纯 ASCII 按字节匹配，非 ASCII 同时尝试 UTF-8 与系统 ANSI 字节（`marker.matchedBy = bytes|ansi|utf8`）。命中即返回 `state:"done-by-marker" + marker{path,contains,matched,bytesRead,matchedBy}`。`state` 是**封闭五值**（互斥、可判定）：`done-by-marker`（**仅当** `marker.matched:true`，真命中文件标记）| `exited`（程序**正常结束**：run/进程已退出）| `break`（停在断点/调试上下文）| `error`（未处理错误框在场）| `timeout`（预算到点仍未终止，**不是**程序失败）。`done-by-marker` **恒蕴含** `marker.matched:true`；`exited` + `marker.requested:true` + `marker.matched:false` 时回执带 `marker.unmet:true`（"程序结束了但没等到要的证据"，**不得**读成目标达成）。 ③ **`observe` 在 break 态显式拒绝**：`reason:"observe_in_break_not_supported" + howToProceed`（先 `control_execution{mode:"continue"|"end"}` 再 observe；一次性读断点上下文用 `debug_session`）。 ④ **对话框预检/分类**：observe 前先分类 #32770 —— 只有 VB6 **运行时错误框**算 `program_error`；IDE 自身对话框（查找/替换/选项/引用/组件/工程属性）是 `ide_own_dialog`，返回 `reason:"observe_blocked_by_ide_dialog" + dialogPrecheck{dialogs[],isProgramError:false}`，**不再被误报成程序错误**。 ⑤ **`get_locals.allowUnsafeInBreak`**（可选，boolean，默认 false）：break 态 get_locals **默认拒绝**（`reason:"break_eval_requires_opt_in" + howToWorkAround + howToRecover`）；`true` 是知情同意——**不保证语句不挂**，仍可能 wedge 管道，恢复走 hub 的 `ide_kill_and_relaunch`（**不是** taskkill）。任何失败回执带 `gateReleased:true/false` + `howToRecover`。 **FACE_DELTA = ADDED 2 optional params to run_project, 1 optional param to get_locals**；工具名/工具数不变（120/11/131）。**冻结面滞后（事实）**：`docs/tools_schema.json` 与 `tools/schema_baseline.json` **本版本未改**，上述三个参数"声明有、冻结无"，须在新 DLL 部署后由活体车道 `dump_tools_schema --freeze` + `gen_hub_snapshot` + `export_tool_catalog` 重冻。 |
| **v1.14.0** | 2026-10-22 | **hub 面工具变更：可续查的重启操作 + 单 JSON 文档内容通道 + peer_dead 判定；add-in 面 120→120、hub 11→11、DSH 面 131→131** ① **`ide_restart` 新增两个可选参数**：`wait_ms`（限制**单次调用**阻塞时长，默认 25000，env `VB6MCP_HUB_RESTART_WAIT_MS`，夹取 1500..300000）与 `opId`（**续查**上一次调用返回的操作）。**等待超时不是失败**：到点返回 `ok:false + pending:true + state:"in_progress" + opId + currentStep + lastKnownState`，重启本身继续在后台进行；`ide_restart{opId}` 从进程内操作注册表回答（保留 30 分钟，hub 重启即忘）且**不启动任何东西**；同一时刻只允许一个重启，进行中再调返回 `reason:"restart_in_progress" + inProgress{opId,...}`（第二次重启可能杀掉第一次刚拉起的进程）。 ② 每条 `ide_restart` 回执新增 `opId/started/state(done|refused|error|in_progress)/finished/elapsedMs/currentStep/lastKnownState/opRule`。 ③ **`content[0].text` 是单个 JSON 文档**：数组/标量型 body 变为 `{singleDocument:true, data, receipt}`（空结果 = `data:[]`，与解析失败可区分），非 JSON body 变为结构化 `content_not_json` 文档（保留 `rawText`），对象型 body **维持扁平**；`VB6MCP_LEGACY_TEXTBODY=1` 只保留旧两段式**一个版本**。 ④ 传输失败信封在 peer 进程**可证明已死**时给出 `peerState:"peer_dead"` + `peerDead:true` + `lostContext{}` + `howToRecover[]`（指向 hub 侧 `ide_kill_and_relaunch`，不再让调用方去 release 一个已不存在的门）。 **FACE_DELTA = CHANGED hub tool ide_restart（+2 可选参数）**；add-in face 无变化，`toolCount` 仍 **120**（hub 11，DSH 面 131）。 |
| **v1.13.0** | 2026-10-22 | **目标解析与离线路由分离 + 就绪等待 + hub 级带外复位（新增 1 个 hub 工具 `ide_kill_and_relaunch`；`ide_start` 新增可选 `ready_wait_ms`；add-in 工具面 120→120，hub 10→11，DSH 面 130→131）** ① **目标阶梯只从活实例选**：discovery 后先按 pid/transport 复核并丢弃已死候选，hub 自己的陈旧记录不再能成为默认目标；显式 `_target` 指向死 pid 时返回结构化 `target_mismatch`，回执带 `onlinePids[]` / `targetDiagnostic{rule,reason,deadTargetPid,...}` / `howToProceed[]` / `suggestion{action,useTarget}`；**0 个活实例**时 IDE 面工具返回 `reason:"no_live_ide"`（`reasonLegacy:"no_ide_running"`）+ `howToProceed` + `suggestion`，不再出现 pid 级 target_mismatch。② **离线路由**：`build` / `make_project(route!="ide")` / `project_meta` 在**无活 IDE** 时由 hub 自己服务（`offline:true` + `offlineRoute:"cli"|"file"` + `targetResolved:false`）：独立 `VB6.EXE /make <vbp> /out <log>` 子进程（回执含 `logPath/logBytes/logEncoding/exitCode/errorCount/firstError{file,line,lineSpace,text}`）、**编译后**补 `IMAGE_FILE_LARGE_ADDRESS_AWARE (0x0020)` 并回读（`laaBefore/laaAfter/laaSet/laaSource`）、**所有改写之后**才取 `artifact{exists,bytes,mtime,sha256}` 与 `artifactChanged`（`hashedAfterRewrites:true` + `rewriteSteps[]`）；VB6.EXE 不可解析时 `reason:"cli_unavailable"` + howToProceed（不是 target_mismatch）；`package_app` / `verify` 在 0 实例时明确 `offlineServed:false` + `no_live_ide`（不假装离线可用）。③ **ide_start 就绪语义**：`ok:true+started:true` 只声明**进程已创建**；`ready:true` 才表示**观测到 add-in 对同一 pid 应答 mcp_info**（轮询、有上限，绝不固定 sleep）；超时返回 `ready:false` + `readyWaitMs` + `readyChecks{registryVisible,transportUp,addInAnswers,addInAnswersBasis,pidAlive,projectMatches}` + `readyProbeRounds` + `pollingHint`（就绪**只测刚拉起的那个 pid**：每轮直读它的注册文件 + transport + `mcp_info`，不用 discovery 集合；`addInAnswersBasis` 三态互斥 `no_route|direct_rpc|middle_service_forward`，`no_route` ⇒ `addInAnswers:null`）；同一 .vbp 已打开时**复用**（`started:false,reused:true,processesCreated:0`）。④ **新 hub 工具 `ide_kill_and_relaunch{pid, project?, ready_wait_ms?}`**：不依赖 add-in 管道——结束**本 hub 拥有**的实例（内存 `startedByHub` 或注册 `startedByHub=true`）、清其注册项、在原 `.vbp` 上重启并轮询就绪；回执固定 `stoppedPid/newPid/ready/readyWaitMs/threeWay{http,mcpInfo,project}/elapsedMs/steps[]`，拒绝路径给 `refused_reason` + `howToProceed`；**非本 hub 实例、已知端口 <8767（含 8766）、远程 bridge 实例一律拒绝**。⑤ 诊断：所有最终失败回执保留 `failureKind/suggestedAction/retrySafe`；`targetDiagnostic` 区分 `explicit_target_not_live` / `stale_default_ignored` / `no_live_ide` / `hub_record_stale`。**FACE_DELTA = ADDED hub tool ide_kill_and_relaunch；CHANGED ide_start（新增可选 `ready_wait_ms`）**；add-in face 无变化，`toolCount` 仍 **120**（hub 11，DSH 面 131）。 |
| **v1.12.2** | 2026-10-21 | **运行/单步落盘的观测与知情同意（R-1021-02；新增 1 个可选参数 `allowDiskMutation`；回执全部增量）** ① **归因更正**：运行/单步把 IDE 内存改动写盘**不是 MCP 工具的行为**，而是 VB6 IDE 的全局选项 **工具(T) → 选项(O)... → 环境 → 组「启动程序时」**（保存更改 / 提示 / 不保存更改）；注册表载体经 t16 实测为 `HKEY_CURRENT_USER\Software\Microsoft\Visual Basic\6.0\SaveBeforeRun`（REG_SZ，**0=不保存更改 / 1=提示 / 2=保存更改**，三值各跑一轮定案）。**本工具不改该选项**，只**只读**读出并在回执给 `ideAutoSaveMode`/`ideAutoSaveReadFrom`。② `control_execution`/`run_project` 新增 **`allowDiskMutation`**（boolean，默认 false）：从**脏内存**（`VBProject.Saved=False` 或某组件 `IsDirty=True`）启动时默认**拒绝启动**，`reason:"dirty_memory_would_be_saved"` + 脏组件清单（>64 截断并标 `truncated`）+ `howToProceed`（save_project / 关 IDE 开关 / `allowDiskMutation:true`）；`dirtyUnknown`（标志读不出）**不拒绝**但在回执写明判据。③ 两工具新增**磁盘快照**：调用前（**早于 run/full_compile 已有的 pre-run flush**）与调用后各一次，覆盖当前工程 `.vbp` + 全部组件文件 + 成对 `.frx/.res`；回执固定 `diskWrites:[{file,beforeBytes,afterBytes,beforeSha,afterSha,reason:"pre-run-save",attribution}]`（`attribution` ∈ `mcp_pre_run_flush` / `ide_auto_save` / `unknown`；**step_* 路径绝不归因 flush**，**run/full_compile 的 flush 写盘绝不归因 ide_auto_save**，归属按 pre-run flush 自己回执的 `projectDirWrites` 交集判定）、`touchedFiles[]`（mtime/size 变而 sha 相同，**不进 diskWrites**）、`diskSnapshot{hashMode,budgetBytes,checkedFiles,snapshotMs,taken}`、`checkedFiles`。工程源文件合计 ≤32 MiB 时**两侧全量 SHA**（`hashMode:"full"`）；超预算降级 `"gated"`（只哈希 .vbp + 脏组件 + stat 变化的文件，未取的摘要一律 `""` + `beforeShaNote/afterShaNote`，**不许编**）。④ **`saveOnRun` 语义**：= 「**本次调用自身**会把 IDE 内存写盘」（**仅** run/full_compile 的 pre-run flush 为 true，并配 `saveOnRunReason:"mcp_pre_run_flush"`），与**恒存在**的 `ideAutoSaveHint:"may_save_on_start"` **并列**；`saveOnRun:false` **不是**「磁盘不变」的保证。⑤ **restart**：回执 `saveOnRun:false` + `deferredStart.pending:true`；deferred Start 的落盘在**下一次 `tools/call` 回执**里以 `deferredStart{startedAt,diskSnapshotRef,emittedInThisReceipt,diskWrites,touchedFiles}` 补齐（**不静默丢**）。⑥ `debug_session`/`run_until_break` 也会 Start 程序：**只加提示**（`saveOnRun:false` + `ideAutoSaveHint/ideAutoSaveMode`），**不加参数、不加拒绝**。⑦ **值 1（提示）可达**：脏内存 Start 会弹**模态**保存框；本次回执先给 `savePromptExpected:true`，下一次回执给 `savePromptPresent/startDialog`，并把「只点取消 / 改开关」写进 `howToProceedSavePrompt`。⑧ 文档：`docs/KNOWN_LIMITATIONS_AND_FAQ.md` §51（开关位置/三取值/人工关闭步骤/「仍以运行后核对 sha 为硬门禁」/edit→run 现在会被拒）与 `docs/T16_SANDBOX_RECIPE.md`（t16 实测默认值 0，**不能反推作者机**）。**FACE_DELTA = CHANGED control_execution,run_project**（各自 `inputSchema.properties` 仅新增布尔 `allowDiskMutation`，不新增 required），`toolCount` 仍 **120**。 |
| **v1.12.1** | 2026-10-20 | **行为修复：E112 保留字判定改为实测分类（工具面/参数/inputSchema 逐字节不变 ⇒ 不重冻）** ① 旧实现用一张约 110 词的平表，把内建函数名、类型名也当成「任何位置都非法」的关键字，`Dim name` / `Dim left` / `Dim mid` / `Dim time` / `Sub Property()` 等合法写法被误报 E112；本版换成分位置实测的三类：**HARD_FORBIDDEN**（lint 会检查的每个位置都被编译器拒绝 ⇒ 仍报）、**POSITION_RULES**（仅部分位置非法，如 `property` 只在模块变量/公有类字段非法，而这两个位置 lint 不检查 ⇒ 不再报）；★实测另发现 **Type 成员（UDT 字段）与模块名这两个位置根本不校验关键字**（`If As Long` 作 UDT 成员、`If` 作模块名都能编译），故矩阵中这两列**不计入**合法性判定，未来复测不要把它们读成「合法」；③ **第二轮按区域复筛**（24 个更细位置 × 同一 264 词 = 6,336 例，全部实测 + 3,739 例二次重编 0 mismatch）确认：**lint 可达的过程类/声明类/参数类三个保守交集仍都是同一 123 词集**（added/removed 均为空）⇒ **不改代码、不升版、不重冻**。同轮实测补充：`property` 作 `Dim` 合法但作 `Static` 被拒（lint 本就不报 `property` ⇒ 属**保守漏报**，非误报）；`.vbp` 的 `Name=` 项目名对非法/关键字值会被编译器**静默改名**（`9abc`→`Vb9abc`、`if`→`vb4projectif`）⇒ 记 `AUTO_RENAMED`、按不可用处理；窗体控件名**只有被代码引用时**才校验（纯声明不校验）；窗体级变量还会与窗体隐式成员（`left/line/name/width`）冲突——这些均属**位置命名空间效应**，不进规则、**INTRINSIC_OK**（编译器接受的内建名 ⇒ 永不报）。② 判定取自 VB6 编译器实测矩阵（264 词 × 11 个语法位置 = 2904 个独立极小工程，证据 `evidence/kw/matrix.tsv` + `evidence/kw/summary.md`），E112 只在该词被实测拒绝的位置触发；词表对比：去掉 17 个误报词、纳入 47 个原先漏掉的真保留字（后者都无法通过编译，故不会命中可编译代码）。③ 无工具/参数/描述变更，`tools/list` 不变，**FACE_DELTA = none**。 |
| **v1.12.0** | 2026-10-20 | **调试能力 + 写/验/构加固（工具 110→120，新增 10 个；既有工具新增参数；回执全部**增量**）** ① 新增 `debug_session`/`run_until_break`/`breakpoints`/`get_locals`/`eval_at_frame`/`run_project`/`last_error`/`move_procedure`/`delete_procedures`/`edit_by_byte_anchor`。② **Start/End 是排队命令**（STA）：一次调用无法覆盖 start→break ⇒ 两段式（start 回 `startedPending`+`howToContinue`，attach 观测）；`debug_session` **不保存工程**（`preRunSave=false`），排队中的 End 报 `endQueued` 而非假 `wedged`（`wedged` 仅指调用内无法恢复）。③ 断点带 **`paneLine`+`diskLine`+来源声明**；手工设的原生断点不可枚举/不可验证（`nativeStateVerified=false`）；条件断点用 `hitCount` 等价表达。④ `get_locals` 只给**名字**（立即窗上下文不是断点帧 ⇒ 值 `localsUnavailable`+原因）；`eval_at_frame` 仅 IDE 立即上下文，break/run 下**默认拒绝**，`allowUnsafeEval` 才执行，**对象/数组表达式求值前静态拒绝**（`object_expression_unsupported`，会挂死 IDE），每次求值闸门兜底、`timeoutMs` 只是调用方预算提示。⑤ `run_project` 两段式、Start 前设 env；Standard EXE 从 IDE 启动**拒绝命令行参数**；**不能 attach 已运行的 -Embedding 进程**；`last_error` 给错误对话框正文+调用栈+局部变量**名字**。⑥ `move_procedure`/`delete_procedures` **事务化**（快照+回滚）；`edit_by_byte_anchor` 按 **ANSI 原始字节**编辑、GBK lead/trail 边界+roundtrip 校验，**失败绝不写盘**（`fileUntouched`+sha 相等）。⑦ `build`/`make_project` 增 `sourceFreshness`（逐组件 disk-vs-pane，只比**规范化文本**，sha 仅供参考）、`final{}`（**LAA 补丁后**采样：hash/PE/FileVersion/`verifiedOnDisk`）、`project`（仅 cli 路线）与 **CLSID `InprocServer32` 守卫**（前后读回/恢复）、以及 cache-hit 的 `compileLogHint`+`diagnoseRecommended`（**cache hit 不解析旧日志**）。⑧ `verify` 增 `semanticDiff`（大小写/Attribute/尾空格归一 + 逐行 `paneLine/diskLine/category`）、`failureLayers{ideDiskDiff,lint,build,modal}`、`overallBasis`；`verify_scope` 仅把**已证 IDE≠disk** 计 fail（`failBasis=ide_disk_diff_proven`）并支持 `baseline`，不变量 `pass+fail+unknown==checked` 恒发布。⑨ `add_component`/`remove_component` 的 `.vbp` 行与组件文件**原子同步**（**文件先落盘**再改 `.vbp`；`.vbp` 恰好一个行尾 CRLF；失败**双侧回滚**），`Form`/`UserControl` add **显式拒绝**（`pairedBinaryUnsupported`）。⑩ `write_file`/`write_source` 对**已存在非空文件 + 空内容**拒绝（`empty_content_would_truncate`，`confirm_truncate=true` 才允许），VB 源文件**永不写空**。⑪ `dependency_graph` 在 `partial` 时发布 `minQuotaHint{charQuota,budgetMs,basis}`，按建议一次重跑即可完整；否则报 `cannotComplete` 而不编造。⑫ **schema 形状**：每个工具的 `inputSchema` 的 `properties` **恒为对象**（无参时 `{}`）、`required` **恒为数组**（无必填时 `[]`）；六个新工具先前误发 `"required":{}`（把空 Dictionary 当成列表传），已在发布前修正，schema 构造器现会把这种形状**强制转换或拒绝**；`dump_tools_schema --freeze` 与严格审计（3e/3f）都会拒绝非数组 `required`。**限制（如实）**：断点帧局部变量**值**不可读、对象/数组 inspect 不可达、任何立即窗求值都可能挂（`release_inflight` 恢复）、`run_project` 的 env 程序内回读**未证实**、UNC 路径未验证。 |
| **v1.11.12** | 2026-10-20 | **诊断一致性 + 写类行序 + 挂死可观测（三张缺陷卡同批；无工具增删；inputSchema 无结构变化 ⇒ 不重冻）** ① **归属判据统一（作者编号 114）**：`ide_list`、转发回执 `ide_state`、`ide_restart`（`ide_stop` 亦附）共用**唯一**解析器 `resolveOwnership()`，回执新增 `ownershipVerdict{verdict, basis[], registryAlive, portAvailable, pid, port, project, projectPath, registryFile, registeredPid, portOwnerPids, portOwnerWhy, netstatAvailable, ownedEvidence[], foreignEvidence[], rule}` + `ownershipVerdictRule`；**两轴并列写清**：legacy `ownership` / `ownedByHub`（＝「本 hub 启动的」→ `ide_stop` / `ide_restart` 的**许可轴**，**不变**）vs 新 `ownershipVerdict`（＝**归属身份轴**）；**归属不授予许可**（用户手工启动但能应答的 IDE ⇒ verdict=owned，仍 `ownedByHub=false`、仍不可被 `ide_restart` 动）。判定：**owned** = 可归因本机 add-in 安装（hub 启动证据 或 注册表 + pid 存活 + add-in 亲自应答 `mcp_info`）；**foreign** = **仅正面证据**（注册 pid 已死但端口仍应答 / 监听端口的 OS owner ≠ 注册 pid / 保留区端口 < 8767 且应答 / 实例监听 pid ≠ 其注册记录 pid）；**unknown** = 其余一切，**含 `portAvailable=false`**（端口不可得 ≠ 别人的：pipe-only 注册故意发布 port 0）。`ide_restart` 的 `not_owned` 拒绝**保留**；**归属未知**改用 `refused_reason:"ownership_unknown"` + `legacyReason`（改造前该路径的旧 reason，供旧消费者比对）+ `ownershipUnknown:true` + `howToDisambiguate`（默认拒绝重启 + 消歧步骤）。`hubInstance.ownership` 不再为 `null`（此前 `ide_list` 说 ours、`ide_state` 回执说 null 的**字面矛盾**）。 ② **写类行序（作者编号 113）**：**`apply_edits`** 新增 `deletedLines` / `deletedText`（截断）/ `deletedTextLines` / `deletedTextTruncated` / `deletedTextLimit`；**`apply_edits` 与 `preview_edit` 都**新增拒绝 `reason:"delete_with_text"`、`reason:"would_empty_component"`（两者**逐字同 reason**），但 **`preview_edit` 不报 `deleted*` 键**（preview 的删除体现在 diff 中；读 `preview_edit.deletedLines` 会得到 `undefined`，不要依赖）；`reload_component` 新增 `vbpPath` / `vbpOrderChecked` / `vbpOrderPreserved` / `vbpOrderRestored` / `vbpOrderChanges[]{key, line, beforeLine, afterLine, beforeOrder, afterOrder}` / `vbpOrderLinesBefore[]` / `vbpOrderLinesAfter[]` / `vbpOrderLinesCapped` / `vbpRevisionBefore` / `vbpRevisionAfter` / `vbpRevisionBumped` / `vbpOrderNote` / `inMemoryOrderMoved` / `inMemoryOrderChanges[]{component, beforeIndex, afterIndex}` / `inMemoryOrderBefore` / `inMemoryOrderAfter`；`save_project` 新增 `vbpOrderChecked` / `vbpOrderPreserved` / `vbpOrderRestored` / `vbpOrderFinalPreserved` / `vbpOrderChanges[]` / `vbpOrderChangeCount` / `vbpOrderLinesBefore[]` / `vbpOrderLinesAfter[]` / `vbpOrderLinesCapped` / `vbpRevisionBefore` / `vbpRevisionAfter` / `vbpRevisionBumped` / `vbpOrderNote` —— **保存后行序变了必须恢复原序**（只重排行、不丢 VB6 写入内容；**恢复失败 ⇒ `ok=false` + `reason:"vbp_order_restore_failed"`，说明文本在 `okReason`**），`vbpRevisionBumped` 是**实测**值；**上限（锁定值）**：`vbpOrderLinesCapped` = **400** 行（`VBP_ORDER_MAX_LINES`；`vbpOrderLinesBefore[]` / `vbpOrderLinesAfter[]` 各最多 400 行，超则截断）、`vbpOrderChanges[]` 最多 **50** 条（`VBP_ORDER_MAX_CHANGES`）、`deletedTextLimit` = **4000** 字符（`VBP_DELETED_MAX_CHARS`）。 ③ **挂死可观测（作者编号 115，字段已锁定）**：`make_project` 新增 `timedOut`（子进程到预算被**杀**）、`budgetEnforced`（预算真的执行了终止；cache hit / 未起子进程 = false）、`budgetMs`（实际预算；未起编译器 = 0）、`budgetSource`（`timeout_ms` / `default` / `clamped_min` / `clamped_max` / `not_applicable`；默认与上下界：`build` **180000 / 1000 / 600000 ms**，`make_project(diagnose)` **120000 / 5000 / 600000 ms**；`budgetMs` 报 **clamp 之后**的实际生效值，clamp 时不叠加原来源名；`not_applicable` = in-IDE Make / 缓存命中 / 预检拒绝，此时 `budgetMs=0`、`budgetEnforced=false`）、`stopReason`（**四值写死**：`completed` / `timeout` / `refused` / `host_modal_blocking`）、`stopReasonPriority`（两者同真时 `timedOut` 是**终止原因**、`host_modal_blocking` 描述**宿主状态**）、`host_modal_blocking`（与 `timedOut` **独立**）、`host_modal{blocking, stage, count, risk, riskReason, riskClass, ownerDisabled, hwnd, title, controls[], dialogs[], observedDuringCall, source}`（`observedDuringCall`/`source` = 调用期间出现过模态框、收尾时已被守卫取消，活体已看不到框）、`isolatedMake{attempted, pid, exitCode, killed, killMethod, elapsedMs, leftover, budgetMs}`（**`exitCode=-1` = 被 kill / 未真正运行**）、`budgetNote` / `timeoutNote`（说明字符串，**不参与判定**）；预算到点**杀子进程且事后无泄漏**。 **`build` 同机制**：`timedOut` / `budgetEnforced` / `budgetMs` / `killMethod` / `childPid` / `compilerElapsedMs` / `childLeftover`，**超时时 `ok:false` + `verdict:"timeout"`**（不再抛「build process failed to run」）—— **本版一并关闭该缺陷卡（`build` 的 `timeout_ms` 现为强制预算）**。 `ide_state` **零工具面变化**（无参数/描述/schema 改动，`inflight{}` 字段集不变），只改行为：隔离编译等待循环泵消息（`Util.PumpNow`），使 `depth>0` 的**只读探测在长调用期间被派发**，`serviceable` 语义按实测 = 该只读调用**确实被服务**。 Tool face：add-in **110 → 110**、hub **10 → 10**、DSH-facing **120 → 120**；ADDED / REMOVED 参数：**无**；`tools/check_schema_baseline.py` ⇒ `schema_changes=0`（**不重冻**）；闸门集 **8/8**。 |
| **v1.11.11** | 2026-10-17 | **版本说明补正 + 一次参数描述 hygiene（10 个参数描述文本变更；无工具增删；inputSchema 无结构变化）** ① `mcp_info.changelog` 里 v1.11.10 的条目此前写成「wording only, no behaviour change」，与事实不符 —— 本版**如实重述**：v1.11.10 修了**三个 winsock 传输缺陷**（a 每连接对象未持久化 ⇒ 跨 TCP 段到达的请求丢前半段、30 s 空闲清扫后才断；b close-after-send 在队列已清空后才判定 ⇒ `Connection: close` 响应滞留至空闲清扫；c flush 累积 `Mid$` 起点越过 64 KiB，且 `send()=-1` 且 `WSAGetLastError()=0` 被误判为致命 ⇒ 正文在 64/96 KiB 处截断），合起来**修掉了间歇 `RemoteDisconnected` 与 >64 KiB 截断**（属**行为修复**）；同批**描述级改动逐项**：**3 个工具描述**（`edit_by_anchor` / `apply_edits` / `preview_edit`：去内部号前缀）、**1 个工具描述**（`verify_scope`：build 证据按项目判定）、**2 个参数描述**（`replace_code.expected_sha256` / `make_project.wait_for_build`）、**1 处运行期回执文本**（`make_project.buildStateNote` 去内部号前缀，字段与语义不变）。② 本版新增：**vba6_probe 的 10 个参数描述**（`exec_probe` / `exec_set_frame` / `frame_probe` / `frame_set` / `frame_set_a1` / `frame_set_a2` / `worker_call` / `worker_p2_engine` / `mem_dump` / `mem_scan`）去掉内部实验编号与逆向轮次号（**语义不变**；无参数增删、无类型/required 变化），发布树注释里的同类编号一并清除；导出审计 regex **不扩**（一次性 hygiene，非审计规则）。面变化构成：add-in **110 -> 110**、hub 10 -> 10、DSH-facing **120 -> 120**，ADDED/REMOVED = 0；GENERATED 三件套按新面重冻。`tools/dump_tools_schema.py` 沿用 v1.11.10 更正后的表述：**仓库内冻结用脚本、不在公开导出树**。 |
| **v1.11.10** | 2026-10-17 | **证据尺度澄清 + 传输修复 + 一次性 hygiene（6 个工具的文本变更；无工具增删；inputSchema 无结构变化）** ① `verify_scope` 描述明确 **build 证据按【项目】判定**：`buildOk`/`buildSource`/`buildNote` 来自**项目级** build 缓存（`Build.CacheLoad`，key/hash 都是项目级），**单个组件不可能“仅自己缺 build 证据”**，因此 `unknown` 是**项目级**结论、`pass` 只在**有 fresh build cache 的工程**上可达（从未构建的工程里任何一致组件都是 `unknown`）；详见下方 v1.11.10 字段语义表。② 传输修复（`tools/dump_tools_schema.py` 是**仓库内冻结用脚本、不在公开导出树**；加上插件 winsock 传输；**无工具面变化**）：`tools/dump_tools_schema.py` 失败回执给出结构化 `reason`（`transport_idle_closed` / `remote_disconnected` / `truncated_response` / `transport_timeout` / `connection_reset` / `connection_refused` / `http_error`），并对**一次**传输失败自动重试一次；插件侧修三个 winsock 传输缺陷（① 连接对象未持久化 ⇒ 跨 TCP 段到达的请求丢前半段，30 s 空闲清扫后才断；② close-after-send 判定顺序；③ flush 的 64 KiB 截断 + `n=-1, WSAGetLastError()=0` 被误判为致命错误）。**冻结语义不变**（只有成功取到 tools/list 才写冻件）。③ 一次性 hygiene（**面变化见末尾**）：`D-nnn` / `MCP-nn` / `REQ-…` 族内部号从**注释、shipped 文档散文、工具/参数描述、运行期回执文本**全部清除，生成件重冻后自然归零；审计 regex **不扩**（这些族与错误码/协议头/版本号不可区分）⇒ **一次性 hygiene，非审计规则**；发布树里只剩 `MCP-Protocol-Version` / `MCP-Session-Id` 这类**协议头名**（非内部号）。面变化构成：ADDED/REMOVED **0**（110 -> 110，顺序相同）；**description 变更 4**（`edit_by_anchor` / `apply_edits` / `preview_edit` / `verify_scope`）；**仅参数描述文本变更 2**（`replace_code.expected_sha256` / `make_project.wait_for_build`，参数/类型/required 均未变）；**运行期回执文本变更 1 处**（`make_project.buildStateNote` 去掉内部号前缀 —— 属行为说明的文字变化，字段与语义不变）。 |
| **v1.11.9** | 2026-10-17 | **命名/形状 + 计数语义批（4 个 description 变更；无工具增删；inputSchema 未变）** ① `add_code.inDeclarationArea` 语义改为**真正的声明区**判定（`insertedAt.ide < 第一个过程头行`）—— 末尾追加 = `false`、首过程前插入 = `true`（键名不变，值语义变）；② `requestedRange` **每个形状都带 `mode`**：`mode:"range"`（`replace_code` / `add_code` / 守卫拒绝 = `{operation,start_line,end_line}`）、`mode:"anchor"`（`edit_by_anchor`）、`mode:"anchor-batch"`（`apply_edits`）；③ `verify_scope`：**UNKNOWN 不再计入 `fail`**，`unknown` 单列且与 `failures[]` 互斥，**不变式 `pass + fail + unknown == checked`**，回执新增 `countsConsistent`；④ 纯 hygiene（**无面变化**）：已发布树内部任务号 **30 个 `TASK-` 族 token（4 个文件）** 与 **36 个裸短号（8 个文件，全为注释）** 全部移除；`Util.bas` 3 行裸 LF→CRLF；导出审计 regex 由「仅首族前缀」扩为**全 `TASK-` 族**（`(?-i:TASK-[A-Za-z0-9]+(?:-[A-Za-z0-9]+)*)`，非数字后缀亦命中），并新增**显式可见的夹具豁免** `-InternalIdFixtures`（默认空 = 发布路径零豁免）；裸短号**不纳入 regex**（与错误码/版本号不可区分）。GENERATED 三件套按新面重冻；add-in **110 -> 110**、hub **10 -> 10**、DSH-facing **120 -> 120**，ADDED/REMOVED = 0（CHANGED 4 = 上述描述）。 |
| **v1.11.8** | 2026-10-17 | **内部任务号清理（重建窗口）—— 行为不变，仅措辞**：工具描述与回执文本里的内部任务标识（各内部任务号族与内部进程目录名）全部移除，源码注释同步清理；面变化 = 若干 `description` 文本、无参数增删、无类型/required 变化；GENERATED 三件套由清理后的描述重新生成/重冻。**面变化构成（实测）**：add-in 110 -> 110、hub 10 -> 10、DSH-facing 120 -> 120，ADDED/REMOVED = 0；**工具描述 9 处**（add_code / apply_edits / check_component / edit_by_anchor / ide_state / make_project / preview_edit / replace_code / validate_code）+ **参数描述 3 处**（export_public_api.name / ide_state.cancel_inflight / make_project.wait_for_build），**无参数增删/无类型或 required 变化**。冻件：`docs/tools_schema.json` == `tools/schema_baseline.json` = **`320CE2DA0067AB2B6E0D68D5A4F8F1A736A7810DFA5C3972E4A5E3619FA9CC26`**、`bridge/hub_schema.generated.json` = **`6665BBC771AD976F5446A19953D132235656C54BB9DFA9D0D46EA1F71C39D92B`**（VERSION=v1.11.8 / 110 / MISMATCH=False）、`docs/TOOL_CATALOG.md` = **`A1BF5BA61B50129442DCA2AA9485A7738D4DC2B1D4E63F43B4BA633B5B77613C`**。导出审计（真实 84 文件树）：`IDENTIFIER_HITS 0` · `INTERNAL_ID_HITS 0` · `MACHINE_PATH_HITS 0` · `INTRANET_IP_HITS 0` · `AUDIT_OK` · `SELFCHECK_OK`。 |
| **v1.11.7** | 2026-10-17 | **写类回执 + 诊断口径批（无工具增删；+4 个此前未声明的参数，其中 1 个仅兼容保留）** ① `add_code` 回执新增 `insertedAt{ide,disk}` 与 `inDeclarationArea`（落点由写后回读测量，双坐标显式）；② **全工具面未知参数门**：以 `tools/list` 的声明参数集为合同，未声明键一律 `ok:false + reason=unknown_argument + unknownArguments[] + acceptedArguments[]`（`_target` 除外）；reason 细分 `unknown_argument` / `component_not_found` / `target_has_no_project` / `ide_not_ready`（`no_project` 与 `ide_model_not_ready` 保留在 `reasonLegacy`）；③ `verify` 对「仅隐藏 Attribute 行」的面板/文件差异给出结构化标注 `attributeOnlyDiff` / `attributeLineCount`（与真实源码差异分开计；保存路径经 VBIDE `SaveAs` 物化隐藏 Attribute 行这一点在描述中写明）；④ `get_component_code{proc|symbol}` 明确范围字段的行号空间（IDE 行）并新增 `startLineDisk`/`endLineDisk`（= IDE 行 + `coords.headerOffset`）；⑤ `check_public_api.membersOnly` 写明作用与局限（文本级轻解析、参数名丢弃、比对永不跳过）；⑥ 返回信封（`result.content[0].text` 是 JSON 字符串）写入契约；⑦ `check_public_api` 接受「按 JSON 字符串字段回灌」的 spec：入口把字面 `\n`/`\t`/`\r` 解码回 TAB/LF 并回 `baselineNormalized:true`（不再把 spec 字段本身指回去当 hint）；⑧ `dependency_graph` 的完整性**可判定**：新增 `edgesComplete` + `componentsUnscanned`/`edgeComponentsUnscanned` + `quotaSource`，并允许调用方用 `char_quota`/`budget_ms` 为单次调用要完整图（每次调用独立，不继承上一次的配额）；⑨ `verify`/`verify_scope`/`get_component_code` 共用**同一个** headerOffset（= headerLines，含隐藏内联 Attribute 行映射），verify 的比较跳过隐藏 Attribute 行，两入口结论不再相反。**面变化构成（对旧冻结面 `1B1F6A7B` 的逐字段 diff，实测）**：add-in **110 -> 110**、hub 10 -> 10、DSH-facing **120 -> 120**，ADDED/REMOVED = 0；`inputSchema` 恰新增 **6 个参数**（`preview_edit.match` / `set_codepane.view` / `vba6_probe.step` / `vba6_probe.frame_index`（后三者描述里注明「被忽略/仅兼容保留」）+ `dependency_graph.char_quota` / `dependency_graph.budget_ms`），**无 required 变化、无既有参数类型变化**；**描述变更 2 处**（`check_public_api` 工具描述、`get_component_code.proc` 参数描述）；其余全部是**回执字段**（不影响 inputSchema）。GENERATED 三件套随本次广播重冻（最终代次）：`docs/tools_schema.json` == `tools/schema_baseline.json` = **`9BCCAFC576E6EA63B5CAE8D92C9622DD38E794EC258A685208E5231167A5A153`**、`bridge/hub_schema.generated.json` = **`BB5D3DEB7588E438E086D2627B9F0F05BC12994092E513CDB846526EC0A12FE5`**（VERSION=v1.11.7 / DECLARED=110 / MISMATCH=False）、`docs/TOOL_CATALOG.md` = **`70ED5B690018DA4B33A766A789743CE7CDA9C038575446E918AB4DE46CEE0C27`**。 |
| **v1.11.6** | 2026-10-17 | **行为语义变更（诚实失败面 + 回执字段）—— 无工具增删；inputSchema 恰 +1 个可选参数** ① `check_public_api`：给了 `name` 时**永远比对**（`membersOnly` 只是输出大小：仍比对，并把成员表加进回执）；baseline 不是 spec 文本 ⇒ `ok:false + reason=baseline_unusable`（不再回“无差异”）。 ①b **两个入口同一字段集**：给了 `name` 与不给 `name`（整工程）都带 `comparedAgainstBaseline`（true/false）与 `baselineLines`，且不可解析的 baseline 走**同一拒绝形状**（`scope:"component"|"project"`）。 ② `check_component` **新增可选参数 `baseline`**（真正声明并生效；无法解析 ⇒ `baseline_unusable`）。 ③ `export_public_api`/`check_public_api`/`validate_code`/`check_component` 对**未声明参数名**一律 `ok:false + reason=unknown_argument` + `unknownArguments[]` + `acceptedArguments[]`（`_target` 是传输上下文，恒接受）。 ④ `write_file`/`write_source` 回执新增 `createdDirs[]`/`createdDirsCount`/`exists`/`size`；父目录建不出来 ⇒ `reason=parent_dir_missing`（`missingPath` + `reasonLegacy=parent_not_creatable` + `reasonAliases[]`）。 ⑤ `find_in_component` 回执新增 `hits[]`/`count`/`nextStartLine`；分页走完 ⇒ `found:false + reason=no_more_hits`（**不回绕**，连续取页不会返回其它过程的行）；`start_line` 越过末行不再是拒绝而是分页终点。 ⑥ `make_project{diagnose}` 的隔离副本纳入 `CompatibleEXE32`/`CompatibleEXE` 引用文件与 PropertyPage 的 `.pgx`，回执新增 `compatibilityReferenceCopied`/`compatibilityReference`；缺失的引用写进 `missingFiles`（不再静默重铸 class ID）。 **面变化构成（对旧冻结面 `720256E2` 的逐字段 diff，实测）**：add-in **110 -> 110**、hub 10 -> 10、DSH-facing **120 -> 120**，**ADDED/REMOVED = 0**；差异恰为：**工具描述 4 处** = `check_component` / `check_public_api` / `find_in_component` / `write_source`；`inputSchema` **+1 个可选参数**（`check_component.baseline`，**无既有参数类型变更、无 required 变化**）+ **参数描述 4 处**（`check_public_api.membersOnly` / `export_public_api.membersOnly` / `find_in_component.start_line` / `write_file.create_dirs`）；其余全部是**回执字段**（不影响 inputSchema）。GENERATED 三件套随本次广播重冻：`docs/tools_schema.json` == `tools/schema_baseline.json` = **`1B1F6A7BF65F7888D552593250E7D3D2A4678F13749954EA3F6B228BD0DBBD1B`**（110 工具）、`bridge/hub_schema.generated.json`（VERSION=v1.11.6 / DECLARED=110 / MISMATCH=False）、`docs/TOOL_CATALOG.md`。活体面来自 t16 沙盒（DLL `6639248D…`，本机不许起 IDE ⇒ 这是唯一可行的活体来源）；重冻后门禁 **7/7**。 |
| **v1.11.5** | 2026-10-17 | **行为语义变更（撤回 + 口径统一 + 判定统一）—— 面变化极小，但消费者必须知道** ① **（方案 B）撤回**「写飞行期间只读调用仍被服务」这一承诺：单线程 add-in 在一次调用在飞时**不执行**任何第二次调用（含只读），被服务的是 **busy 拒绝回执**；`readonly_serviceable` 现**恒 false**（键保留），`cancel_inflight`/`release_inflight` 降为 **BEST EFFORT**（陷宿主模态循环的调用永不中止）；**唯一可靠出路 = 等待或从外部终止 VB6 进程**。 ② **** `newTotalLines` 统一为**调用方口径**（`get_component_code` + `splitlines`，末尾空行不计），四条写路径（`replace_code`/`add_code`/`ChangeResult2`/`BreakpointEdit`）同一计数函数；新增回执字段 **`trailingBlankLine`**；`syncWarning` 改为**与面板真值比较**（原先两个内部计数互比，永不触发）。 ③ **** hub：`ide_stop`/`ide_restart` 共用**一个所有权解析器**（同一 pid 不得相反 verdict），所有拒绝路径携带 `alsoForeignPort`/`ownershipBasis`/`portRuleApplies`/`portBasis`；`ide_list` 自述 `registryWrites:0`+`registryPolicy`。 **面变化构成（实测 diff）**：add-in **110 -> 110**（**无增删工具**）、hub 10 -> 10、DSH-facing **120 -> 120**；`inputSchema` **无参数增删/无类型变更**，**仅 2 处描述文本**：`ide_state.description` 与 `ide_state.cancel_inflight.description`；新增的 `trailingBlankLine` 是**回执字段**（不影响 inputSchema）。**增量（同日并入同一条广播）**：④ **** `command_bars{action:list}` 的每条控件新增 `executable` 与 `enabledReason` 两个**回执字段**，口径 = **action=execute 实际会怎样**（按同一枚举顺序取**第一个**该 caption 的控件，看它是否 enabled；captions 跨 bar/工具栏重复，所以`enabled` 单独**不能**推断可执行性）——**只加回执字段，inputSchema 不变**；⑤ **** 把实测到的「尾随换行规则」写进 `replace_code`/`add_code` 的**工具描述**：★**它不是 op 的属性** —— 结果尾部取决于**基底是否以换行结尾**与**插入位置**（**中间**插入不触碰尾部 ⇒ 继承基底；**末尾**插入会重建尾部 ⇒ 规范化）；**禁止按 op 归类**，以回执 `trailingBlankLine`（配合 `newTotalLines` 的 `splitlines` 口径）为准 ⇒ **面变化 +3 处描述文本**（`command_bars`/`replace_code`/`add_code`），**仍无参数增删/无类型变更**。GENERATED 三件套随本次广播重冻。 |
| **v1.11.4+台账补记（描述乱码修正）** | 2026-10-17 | **两条工具描述由乱码修正为正确中文（）+ 显式重冻；工具集零变化**。工具面：add-in **110 -> 110**、hub 10 -> 10、DSH-facing 120 -> 120；**ADDED/REMOVED 均无**；110 个工具**逐字段全量 diff 恰好 2 处**：`command_bars.inputSchema.properties.bar.description`（旧 `璋冭瘯` -> 新 `调试`）与 `vba6_probe.inputSchema.properties.worker_p2_engine.description`（旧 `搂5` -> 新 `§5`），且**方向已证**：旧串 == 新串 UTF-8 字节按 CP936 读（纯乱码，无语义变化）。根因 = `CodeTools.cls` 以 UTF-8 存储而 VB6 按 ACP 读；已改源（ASCII + `ChrW$`）并重建（产物 `0A96D2A1` -> `EF0830A9`，CLSID 不变），t16 同步为同字节。**已显式重冻**：`docs/tools_schema.json` 与 `tools/schema_baseline.json`（131,953 -> 129,112 B，.bak 保留）、`bridge/hub_schema.generated.json`、`docs/TOOL_CATALOG.md`；门禁 `schema_changes=0`、`hub_alignment_failures=0`、`check_live_tools` **live_tool_changes=0**。注意 **未改 `Util.bas` 的 `VB6MCP_CONTRACT_VERSION`**（仍 v1.11.4）—— 本次是描述文本修正，未走版本递增；**已裁定：不 bump**（S 主卡：递归 diff 覆盖 110 工具、恰好 2 处且都是 description、方向已证为纯乱码 ⇒ 属「还原意图文本」，无工具/参数增删；将来出现真正的面语义变化——增删工具/参数/枚举值——则**必须** bump）。 另：**b 的 hwndActive 门控项在 t16 交互会话实测「与 GetForegroundWindow 冗余」**（hwndActive == 前台顶层窗口；MDI 子窗格只体现在 hwndFocus），故该项对 MDI 子窗格切换零增益；建议改用 hwndFocus 或回退，**待 S 裁定**。工具面零变化，本广播与冻结不受影响。同次台账补记：`TraceGen.bas` 的 236 个 GBK 字面量字符按 `--strict` 为 WARN ⇒ **记录为可移植性债务**（产物串正确、乱码不在，ACP=936 下实证），**有意保留**：改写为 `ChrW$` 链会毁掉一纸中文模板的可读性。另：`CodeTools.cls` 等 7 个组件已由 UTF-8 统一为 GBK 无 BOM，源印记 `288afca0` -> `13bf3be1` -> `417e8377`（FILES=41）。|
| **v1.11.4+台账补记** | 2026-10-16 | **台账补记（非版本变更 · 工具面零变化 · 不重冻）**。不新增版本、不改 `VB6MCP_CONTRACT_VERSION`、不动任何冻结件；`tools/list` 的 name/description/inputSchema 逐字节不变（add-in 110 / hub 10 / DSH 120）。 **(1) · 加法回执字段 + 「缺席也是一种断言」**：`build`/`make_project`/diagnose/cache 新增/明确 `matchRule` · `firstErrorKind` · `rawLen` · `rawSha256` · `logEncoding` · `logBytes` · `rawLine` · `logExcerpt`；`firstError` 三态 = `file-line` / `project-level` / `log-unparsed`，**未解析出首错时 `firstError` 整体缺席**（不是 `{file:"",line:0,lineSpace:"disk",text:""}` 空壳）。 **(2) · `quit_ide` 三段契约**：回执含 `requested` / `pendingExit` / `exited` / `blockedBy` / `evidence` / `howToConfirm` / `waitedMs`；**`exited` 恒为 `false`**（进程无法观测自身退出，是契约不是缺陷）；其 description 按卡面 (iv) 走完整重冻链（`docs/tools_schema.json` + `tools/schema_baseline.json` + `docs/TOOL_CATALOG.md` 各改 1 行，提交 `f244365`），工具数仍 110/110、无增删 ⇒ 属**描述/回执语义**变更而非工具面变更。 **(3) · `add_code` 行为变更 + 参数描述重冻**：string 分支由「按**插入符**落点」改为「**末尾追加**」（file 分支本就追加）⇒ `requestedRange == appliedRange`，落点由**写后回读**验证（`landedVerified`/`landedLine`）；回执新增 `ok` / `okReason` / `sha256Before` / `sha256After` / `lineDelta` / `appliedRange` / `landedLine` / `landedVerified`；**迁移**：需按位置插入者改用 `edit_by_anchor` 或 `replace_code`。同批 `replace_code.expected_sha256` 的**参数描述**改动已重冻（`docs/tools_schema.json` + `tools/schema_baseline.json`，`CHANGED=['replace_code']`；`TOOL_CATALOG.md` 不变是正确的——它记的是工具**描述**而非参数描述）。 **(4) 环境契约 · VB6 插件启用门（，本机实测）**：插件是否被实例化由 **`%WINDIR%\VBADDIN.INI` 的 `[Add-Ins32]` 段 `<ProgID>`** 决定，**不是**注册表 `LoadBehavior`：`0` = **曾被一次真实加载失败禁用、VB6 此后永不重试**（`regsvr32`/`reg import`/`LoadBehavior=3` 都救不回，且 VB6 不改写注册表 ⇒ 链上全绿而插件是死的）；`1`/`3` = 启用，**`3` 为稳态**；**缺失条目** ⇒ 退回注册表、可加载。实测口径两处：**成功加载时**才改写为 `3`（**不是 IDE 退出时**），且**原地只改「值」那一个字节**；`VB6.EXE /make` 会把**已启用**的值归一化为 `1`，**不会**把 `0` 变回 `1`（⇒「重建救不回」成立）。**判据**：`tools/verify_registration.bat` 以该文件为**闸门 A（决定性）**，`=0` ⇒ `EXIT=1` 并打印处方；缺失条目 ⇒ `EXIT=0` + 显式提示。 **(5) 补记口径**：(1)(2)(3) 为既有卡的终版记录、本轮**未重测**（字段名与数值以各卡留存的证据文件为准）；(4) 为本轮**实测**（五格矩阵 + 一次真实加载循环 + `regsvr32` 与 `/make` 的分离实验）。 |
| **v1.11.4** | 2026-10-15 | **既有问题·台账批（同一批缺陷）** —— 来源 = 外部项目执行处工具问题台账（28 条）的分诊 §2/§5。Tool face: **add-in 110 -> 110**、**hub 10 -> 10**、DSH-facing **120 -> 120**；**无工具增删**，`inputSchema` 仅 `make_project` 新增一个**可选**参数 `wait_for_build`，另有 **7 个工具的 description 变化** ⇒ 冻结面**显式重冻**：`docs/tools_schema.json` == `tools/schema_baseline.json` = **`E1A53E42816526E3AC66B57D79D4013273D39463DD11D49340C0C7B81C5F248B`**（110 工具，130,710 B）、`bridge/hub_schema.generated.json` = **`12270E2784B4B4ED6A695271FEC19BCE26425CC4BE78911933F4B0D3243021AC`**、`docs/TOOL_CATALOG.md` 同步重生成。**写前结构守卫（最高价值）**：写类工具在**写入前**检查新块首尾关键字配对 —— `extra_procedure_end`（块自带多余尾部 `End Sub|Function|Property`：报 `extraEndLines[]` 行号 + 已配对的 `pairs[]` + `suggestedText`，不再变成编译期 E105）、`missing_procedure_end`（`replace_code` 的 set/append 与 `add_code` 必须自足，否则拒绝）、以及**结果级**守卫 `predicted_procedure_imbalance`（对 `replace_code`/`edit_by_anchor`/`apply_edits`/`add_code`/`preview_edit`：**仅当改写前面板配平、改写后将失衡**才拒绝 ⇒ 片段式合法编辑不受影响，且已失衡的模块仍可被修复）。**歧义候选 + 多行锚点**：锚点未命中/歧义的回执带 `hits[]{line,proc,text}` + `coords`/`lineSpace`（可用 `occurrence`/`scope`）；`preview_edit` 补齐**多行块锚点**（此前只替换首行 ⇒ 预览与写入分歧；现与 `apply_edits` 同一解析路径、同一条结构拒绝）。**行号漂移**：`replace_code`/`apply_edits` 回执新增 `procBefore`/`procAfter{proc,startLine,endLine,firstLine,lastLine}`；`replace_code` 的 replace/delete **保留** `expected_sha256` 硬前置（**选它而非按过程寻址**：整面 sha 能发现**任何**漂移，过程内相对寻址只能发现跨过程漂移、且它本身要从被钉住的文本重新推导）。**构建状态**：`make_project` 新增 `wait_for_build`（ms，默认 0 = 既有 15000，上限 120000）与 `waitRequestedMs`/`waitUsedMs`，并恒带单一可查询状态 `buildState`（`blocked_by_host_modal|pending|completed|completed_without_artifact_change|failed|no_artifact`，与 `verdict` 同源）；**本批验收自己抓到的缺陷**：当隔离副本 `/make` 诊断已证编译失败时，回执**不再报 `pending`**（那会叫调用方再等等、重读 output{}，掩盖已可证实的失败），改为 `verdict=compile_failed` + `buildState=failed` + `pendingOverriddenByDiagnosis=true`。**表述层**：`validate_code`/`check_component` 描述写明 **linter ≠ compiler**（只做词法+结构，**不解析类型/符号** ⇒ 永远不能报调用了不存在的方法，是**明确的非目标**），真编译结论用 `build`（隔离 /make，真日志 + `firstError`）。**实测（t16；I=`2d376202a7116374…`，DLL=`9DBD0C840A3329B8314FEEE200AA9B84D3EBCB1B864B0C3B691C8EFC8B9378DB`，check=VERDICT=PASS）**：验收脚本 `<scratch>/t99_tests.py` **19/19 通过**（含 06 多余收尾被拒 + 正常块被接受 + 反向缺收尾被拒 + 片段式编辑不误拒；05 同文本两次返回 2 条候选且**不写入** + 多行锚点整块替换；07 陈旧 sha 拒绝漂移 + 过程回显；03 正例 `completed` / 反例 `failed` 而非 `pending`）；battery 全绿 **92 通过 / 0 失败（EXIT=0，含新增 段）**。**mcp_info.breaking 本版不新增条目**（无签名破坏；新增参数可选、默认行为不变）。另：`Util.bas` 的 `VB6MCP_CONTRACT_VERSION` 同步 `v1.11.3 -> v1.11.4`（四个 `skills/*/SKILL.md` 的 `skill-truth` 契约行随闸门同步）。 |
| **v1.11.3** | 2026-10-15 | **`dependency_graph` 确定性截止 + 诚实截断（回执语义修正；无工具增删）**。Tool face: **add-in 110 -> 110**、**hub 10 -> 10**、DSH-facing **120 -> 120**；`inputSchema` **逐字节不变**，`tools/list` 的**唯一**变化是 `dependency_graph` 的 **description**（写明确定性配额 / 默认部分覆盖 / 该看哪些字段 / 如何取完整结果）⇒ 冻结面**显式重冻**：`docs/tools_schema.json` == `tools/schema_baseline.json` = **`B8CFE5828A86C7E9C1F3FD2359F4B0FB9225FC6DEF5F692B3DA6D12CA5440495`**（110 工具，126,167 B）、`bridge/hub_schema.generated.json` = **`E46AB51D7F324E51C2290CED0F91CE0106F62E6D2AEE44F4470BF439E31D4A68`**、`docs/TOOL_CATALOG.md` 同步重生成；除此之外工具面零变化。**缺陷（本卡 ）**：旧实现用**墙钟**（`DEP_BUDGET_DEFAULT_MS=8000`）决定截到哪，同一份未改动源码连续三次调用 `procEdges` 会漂移（3,474 → 3,915 → 4,168），大工程在 `0` 与 `811` 之间跳；更糟的是**阶段被截断仍报 `*StageCompleted=true`**，且 `m_exhausted` 的全局闩锁让**过程边阶段实际只跑了 1 个过程**（`proceduresScanned=1 / procCount=1400`）⇒ 默认档交付的是"有节点、零边"的图谱。**修复**：① 主截止改为**确定性字符配额** `VB6MCP_CHAR_QUOTA`（默认 **900000**），按**已处理字符数**计 ⇒ 同一输入必停在同一处；② 每阶段有**保留份额**（tokenize 40% / 组件边 65% / purity 75% / cycles+procEdges 100%）⇒ 默认档**必然产出过程边**；③ 墙钟降为**最后手段**（`VB6MCP_BUDGET_MS`，默认 30000）并用**独立**字段 `timingTruncated` 标注（`budgetExhausted` 只表示配额截止，二者语义不同）；④ **诚实化**：**ADDED** `procEdgeStageCompleted`；`partial` 改为"**有阶段没做完**"；`edgeStageCompleted` / `purityStageCompleted` / `cyclesStageCompleted` / `procEdgeStageCompleted` 在**该阶段或其输入阶段被截断**时一律 `false`（不再谎报完成）；`componentsTotal` 改为**工程真实组件数**（由采集侧显式回填，不再等于"恰好采集到的数"）；`componentsScanned` 只计**完整 tokenize 过**的组件。**同名字段语义变化（非签名破坏）**：`partial` / `componentsScanned` / `componentsTotal` / 四个 `*StageCompleted`；并且**默认覆盖率显著变小**（默认档 = 确定性的部分覆盖，取全图请同时提高 `VB6MCP_CHAR_QUOTA` 与 `VB6MCP_BUDGET_MS`）。**确定性口径（契约）**：同一输入重复调用，**答案部分逐字节相同**；允许变化的只有**信封字段** `ms` / `elapsedMs` / `inflight.stampedCount` / `inflight.sinceMs` / `ide.pid` / `ide.port`（自计时 / 会话计数 / 会话身份），`tests/battery_test.py` 的 段带**逐字段理由的显式掩码清单**，并断言 `partial` 与 `*StageCompleted` 自洽、默认档 `procEdgesFound>0`。**实测（v1.11.3 构建，t16，DLL `B8FE31B0…`）**：默认配额下 SMALL `procEdgesFound=1217`、约 1.5 s，LARGE `procEdgesFound=6929`、约 14 s，两者 `partial=true / budgetExhausted=true / timingTruncated=false`，**三次运行掩码后逐字节相同、跨 IDE 重启亦相同**；提高配额后 SMALL 完整图 `procEdgesFound=9690 / partial=false / 约 35 s`，同样三次逐字节相同。**已知限制（另立新卡）**：LARGE 的**完整图单次调用交付不了**（约 300 s 处连接被关闭）⇒ 需分页 / 可续 + 单次上限。**mcp_info.breaking 本版不新增条目**（无签名或调用面破坏）。另：`Util.bas` 的 `VB6MCP_CONTRACT_VERSION` 同步 `v1.11.2 -> v1.11.3`（四个 `skills/*/SKILL.md` 的 `skill-truth` 契约行随闸门一起同步）。 |
| **v1.11.2** | 2026-10-15 | **源码印记（SOURCE STAMP）回执字段（纯 ADDITIVE，工具面零变化）**。Tool face: **add-in 110 -> 110**、**hub 10 -> 10**、DSH-facing **120 -> 120**；`tools/list` 的 name/description/inputSchema **逐字节不变** ⇒ `docs/tools_schema.json` / `tools/schema_baseline.json` / `bridge/hub_schema.generated.json` / `bridge/hub_tools.json` **不重冻、必须逐字节不变**（若变即视为工具面被意外改动，须停手排查）。**ADDED 回执字段（2，加法）**：`mcp_info.build.sourceHash`（源码输入集合的确定性指纹）与 `mcp_info.build.sourceAlgo`（算法标识，首版 `vb6src/1`）。语义：**相同代码 ⇒ 相同 sourceHash**（跨次、跨机；二进制本身可因构建环境不同而不同），**不承诺反向**（源码改动对编译器可能无效，产物理应相同）。 另：`Util.bas` 的 `VB6MCP_CONTRACT_VERSION` 同步 `v1.11.1 -> v1.11.2`。 **算法标识演进（`sourceAlgo` 值变化，字段/结构不变 ⇒ 不 bump 契约）**：2026-10-15 作者裁定"版本号允许自增，把版本号作为单独比较项"，`tools/srcstamp.py` 升级为 **`vb6src/2`** —— 在对 `.vbp` 取哈希前**屏蔽四个版本键的值**（`MajorVer`/`MinorVer`/`RevisionVer`/`AutoIncrementVer`，**键保留**），使 `/make` 把自增后的 `RevisionVer` 写回 `.vbp` 时 **I 不再漂移**；产物版本改由 `check` **单独输出并比较**（`ARTIFACT_VERSION`/`VBP_VERSION`/`VERSION_COMPARE`/`VERSION_RECORDED*`），**版本不同不影响 `VERDICT=PASS`**。定义变更 ⇒ 全部 I 作废并重算（MCP_VB6 需重新 `emit`）；`mcp_info.build.sourceAlgo` 由 `vb6src/1` 变为 `vb6src/2`。 |
| **v1.11.1** | 2026-10-14 | **有界化与引导修正（ADDITIVE 工具面变更 + 非功能性修正）**。Tool face: **add-in 110 -> 110**、hub 快照随之重生成；**无工具增删**。**ADDED 参数（2，可选、纯增量）**：`export_public_api.membersOnly`（轻量模式：只回成员表、跳过 spec 文本与第二次解析）与 `check_public_api.membersOnly`（只回成员表、跳过 violations/added 文本生成）；两者默认 false ⇒ 旧调用方行为不变。**ADDED 回执字段**：`dependency_graph` 与两个 API 工具**新增** `budgetExhausted` / `budgetMs` / `elapsedMs` / `stage`（前者另加 `tokenizedComponents` / `edgeScannedComponents`，后者加 `scanned` / `componentCount`）；默认预算 **8000 ms**，可用环境变量 `VB6MCP_BUDGET_MS` 覆盖。**修复（功能性，见 /-20）**：`dependency_graph` 的组件边阶段由 O(模块数²) 全文匹配改为**一趟式成员判定**；两者不再可能越过预算把宿主拖入 wedge（实测 204 组件工程：旧 DLL 无返回且持续烧 CPU ⇒ 新 DLL 8.5–11.7 s 返回、后续 `ide_state` 47 ms、`Responding=True`）。**非功能性**：`initialize.instructions` 里 `validate_code` 的措辞由 "a local rule engine" 改为 "runs INSIDE the add-in (a live IDE is required - NOT an offline check)"（消除 的误读）；`ide_list` 描述与别名回显、转发失败分类 `noIde`/`wedged`/`toolError` 随本批在 hub 侧落地。DELIBERATE re-freeze（四冻件全变）。 |
| **v1.11.0** | 2026-10-14 | **经验教训模块迁移到 Python 侧（架构归位；工具名/描述/参数零变化）**。Tool face: **add-in 115 -> 110**（5 个 lesson 工具移出插件），**hub 5 -> 10**（同 5 个工具改由 hub 提供），**DSH-facing total 120 -> 120**。**MOVED（5）** = `list_lessons` / `add_lesson` / `delete_lesson` / `export_lessons` / `import_lessons_md`：name / description / inputSchema **逐字不变**，调用方无感；差别是它们现在**没有 IDE 在线时也能用**。数据仍为 `mcp_lessons/defaults.json` + `mcp_lessons/user.json`（schema 不变，既有 user.json 原样沿用）。实现：新 `tools/lessons.py`（Python 权威实现 + 数据 owner），hub 薄转发；VB 侧删除 `Lessons.bas` 的全部引用（`CodeTools.cls` 的注入 + 5 个注册/分发、`McpServer.cls` 的 prompt 注入），`Lessons.bas` 文件保留但不再编入工程。**对调用方行为零变化**（`initialize.instructions` 的 lessons 段逐字节一致，金标准 `re_work/t1102/instructions_full_golden.txt`）。DELIBERATE re-freeze（四冻件全变）。 |
| **v1.10.1** | 2026-10-14 | **引导可见化（NON-FUNCTIONAL：不改任何工具行为、不增删工具）**。动机：调用日志实测显示「依赖图/影响面」族工具在 2026-10-09 被跨工程试用后即无人再用，根因之一是**注册描述没告诉调用方过程级调用边的存在**（`dependency_graph` 实际返回 `procNodes` / `procEdges` / `procCount`，描述却只写组件级）。内容：① `initialize.instructions` 的**两个生产者**（add-in `BuildInstructions()` + 本节 text 块）新增 **CAPABILITY INDEX**（问题到工具的路由）与 **FILE ENCODING** 规则；② `dependency_graph` 描述补过程级调用图；③ `write_file` / `write_source` / `project_meta` 描述补「.vbp/.frm/.cls/.bas/.ctl 一律 ANSI 无 BOM」规则（BOM 事故的产品侧引导修复）。**行为零变化**（只改字符串常量 + `Util.bas` 的 `VB6MCP_CONTRACT_VERSION`）。Tool face：**add-in 115 -> 115**、hub 快照随之重生成 ⇒ **DELIBERATE re-freeze（四冻件全变）**。 |
| **v1.10.0** | 2026-10-13 | **BATCH-1：能力簇落地 + 描述标签中性化**。Tool face: **add-in 114 -> 115**（新增只读 `desktop_probe`），hub 5 -> 5，DSH-facing total **119 -> 120**。**ADDED** = `desktop_probe`（只读：进程的实时窗口站/桌面/sessionId/可交互性/前台状态；**fail-closed** ⇒ 任一 API 失败即 `ok:false` + `failureKind:"api_unavailable"`）。**CHANGED（12）** = `build` / `close_active_dialogs` / `ide_quiet_mode` / `ide_state` / `immediate_eval` / `immediate_run` / `inject_trace` / `make_project` / `remove_trace` / `todo_status` / `validate_code` / `vba6_probe`：其中 10 个只改**描述文本**（内部标签中性化，行为与回执不变），`ide_state` 与 `vba6_probe` 描述未变、只加参数。**新增参数（3，纯增量）** = `vba6_probe.exec_line_file` / `immediate_eval.expression_file` / `immediate_run.code_file`。**R13 失败信封**：`vba6_probe` 七类 `failureKind`（`not_loaded` / `alloc_failed` / `arg_out_of_range` / `address_invalid` / `not_readable` / `engine_zero` / `refused_by_design`），`immediate_*` 的 `exec_error` 包装新增两类（`channel_error` / `statement_raised`），且 `immediate_*` **自身**的 `exec_error` **刻意不带** `failureKind`。**修复**：静默截断改为结构化（`lenRequested` / `lenClamped`、`maxHits*`）；`arg_out_of_range` 取代裸 VB 溢出。**未落地**：`immediate_window_text`（R1 NO-GO ⇒ 不加该工具）。**D1 有界残留**：`todo_status` 保留功能性自动发现（仓库内**可选任务文件夹**）并新增环境变量 `VB6MCP_TASKS_DIR` 覆盖；契约与公开面不再出现该目录名。DELIBERATE, S-authorized re-freeze（四冻件全变；hub 快照随之重生成）。 |
| **v1.9.0** | 2026-10-13 | **工具面扩容 + 卡死可恢复**。Tool face: **add-in 111 -> 114**（新增 `immediate_eval` / `immediate_run` / `ide_quiet_mode`），hub 5 -> 5，DSH-facing total 116 -> 119。**Add-in CHANGED** = `ide_state`（新增 `release_inflight` / `release_min_held_ms`；正文**恒带** `release{}` 固定键集）；只读回执仅在**发生过强释**后才追加 `gate*` 键；`ide_quiet_mode` ON 时**每次** `tools/call` 正文带 `quiet{...,advisory}`、`control_window{action:activate}` 回执带 `quietModeSuppressedFocus`。**纯增量**：quiet OFF / 未强释路径的既有回执逐字节不变。DELIBERATE,S-authorized re-freeze（四冻件全变）。|
| **v1.8.0** | 2026-10-12 | P1 fix pack. Tool face: add-in 111 -> 111,hub 4 -> 5 (new `ide_restart`),DSH-facing total 115 -> 116. Add-in CHANGED = `replace_code` (replace/delete now REQUIRE `expected_sha256`; the refusal returns `currentSha256`),`validate_code` (+optional `baseline`; every diagnostic gains `proc`/`procOffset`/`fingerprint`/`stableKey`; always-present `baseline_diff`),`ide_state` (+`cancel_inflight`/`cancel_target`; every read-only receipt gains `inflight{}`). Behaviour fixes with no schema change: PRE-WRITE block-pairing refusal in replace_code/add_code/apply_edits/edit_by_anchor; multi-line (block) anchor + duplicate-anchor hits + procedure first/last echo in edit_by_anchor; close_active_dialogs emits ONE key set on dry_run and real clicks; build/make_project stop publishing an EMPTY digest and re-read logs in their own encoding. DELIBERATE,S-authorized re-freeze. |
| **v1.7.0** | 2026-10-09 | **版本号补齐（additive，不改工具面 ⇒ 免重冻）**：为 `replace_code` / `apply_edits` / `edit_by_anchor` 新增的**可选** `snapshot`（`type=boolean`，**默认 off ⇒ 默认行为不变**，跨构建 A/B `ab_differences=0` 为证）补上契约版本号与迁移条目 ⇒ **工具数不变（111 → 111）**；对应冻结件 sha **`1EE3BFA9E917AB6E5D148D409FB6B33ACA91EF2B20C7988E39ADE589536108BE`**（== `docs/tools_schema.json`，本卡**不再重冻**）；`mcp_info.breaking` 本版**不新增条目**（非破坏性），迁移说明见 **§9** |
| **v1.6.0** | 2026-10-09 | **工具面扩容**：新增 `replace_in_files`（跨组件字面/正则替换，`dry_run` 默认 true）与 `undo_edit`（组件级显式单步撤销；写入前把 pane 文本快照到 `<工程>\_undo\`）⇒ **109 → 111，必须显式重冻**；同一次重冻并落地 `send_window_message` 的**描述补全 + 高危消息护栏**（新增参数 `confirm` ⇒ 本版 `CHANGED=send_window_message`）；另**登记**（不改 schema）`add_component{type:"form"}` 改由 IDE 自建 + `createdVia`/`ideRouteNote`、`control_execution.mode` 的 12 值 enum、以及两条工具面局限（`command_bars{list}` 不含 Add-Ins 弹层；`AddFromTemplate` 窗体模板路线本机不可用） |
| v1.0.0（隐含） | 2026-09-21 | 既有 96 个工具（此前只有 live 快照与冻结基线，没有书面契约） |

> **v1.16.4 后补（hub ADDITIVE 输出字段；不新增版本、不重冻、不重生成 hub snapshot、不影响 G13）**：2026-10-23 批次 35 —— `ide_start` / `ide_restart` / `ide_kill_and_relaunch` 的**成功回执**新增
> `cwd`（绝对路径字符串，或 `null`）与 `cwdBasis`（`"spawn_cwd"` | `"unavailable"`，固定两值）。**语义**：`cwd` = 该次启动在 **spawn 点**捕获的工作目录（未显式传 `cwd` 选项时即 hub 进程当时的 cwd）；
> 该次调用**没有 spawn**（例如 `ide_start` 命中已有实例的复用路径）或取不到时 ⇒ `cwd:null` + `cwdBasis:"unavailable"`（fail-closed，不猜、不读目标进程 PEB）。**追加边界（34b 同源实测）**：
> VB6 在**优雅退出阶段**可能向该 CWD 写工程组文件 `.vbg`（已实测条件：会话中建立过工程组上下文且退出时工程干净 —— W37.T6 53 B / +250 ms、W38b.B1b 50 B / +250 ms；单工程 clean quit 与无工程 quit 不写；
> hub stop/kill 族**未测**，不声明）。**冻结面零变化**：工具面 add-in 120 / hub 11 / DSH 131；`docs/tools_schema.json` == `tools/schema_baseline.json` = `AAFA1C6C…`；`bridge/hub_schema.generated.json` = `828B4DAD…`（未重生成）。
> 证据目录：`fix/TASK_HUB_CWD_35_20261023/`（`PREREGISTER_35.md` / `STATIC_IMPL_35.md` / `BATCH35_RUNTIME_RESULT.md`）。

> **批次 37 修复（同属 v1.16.4 后补；行为修复，不改工具面/冻结面）**：`ide_restart` 的新实例就绪判定**改用「新实例注册的 `project` 名」为期望值**（W44 根因：此前 pipe 分支用 `.vbp` 路径 basename、TCP 分支用 pre-restart `target.project`，都会把「工程名 vs 工程路径/旧实例」错配 ⇒ `projectMatches` 恒 false ⇒ `three_way_ok=false` ⇒ 误回 `refused new_instance_not_ready`）。现在 `wantNameSource = "registration" | "basename_fallback" | "unavailable"`，比较为 trim + 大小写不敏感、缺失侧保持 `null`（unverified）；W45 复测 pipe/TCP 两分支 `ide_restart` 均 `ok:true / restarted:true / three_way_ok:true`。证据：`fix/TASK_RUNTIME_PROBES/WINDOW_44_RESULT.md`（根因）/ `WINDOW_45_RESULT.md`（复测）/ `fix/TASK_HUB_CWD_35_20261023/BATCH37_IMPL.md`。

> **批次 39（同属 v1.16.4 后补；add-in 行为改进，不改工具面/描述/参数/冻结面，不重冻）**：`last_error` / `close_active_dialogs` 的对话框按钮选择改为 **控件 ID 优先、文本兜底**。按对话框种类**分域**（runtime_error：Debug=4801 / End=4800 / Continue=4802 / Help=4902 / Body=4803；save_confirm：Yes=6 / No=7 / Cancel=2；eject_refused：OK=2；load_fail：Yes=6 / No=7 / Help=9），必须 `Button` + `enabled` + **文本二次校验**（冲突 ⇒ `id_text_conflict` 拒绝）才点击；`continue`/`end` **永不点击**，`last_error{dismiss}` 保持 **Debug-only**，`close_active_dialogs` 对 runtime_error 框按 kind 拒绝；无 ID / 未知 ID 时回退原文本路径（行为与旧一致）。仅在真正发生 ID 解析/点击时，回执新增 `idBasis` / `controlId` / `idTableVersion`（`last_error` 为 `dismissIdBasis` / `dismissControlId`）；无框/无错误路径**零新增键**（lane 27-call 0-diff，verify_39 PASS 仅 build_identity×3）。回滚开关 `VB6MCP_DLG_ID_FIRST=0`（运行期，无需重编）。G5 现场：runtime_error 4 次独立冷启 ID 集 `[4800,4801,4802,4803,4902]`、pairs=6 / text-varies=4 / unstable=0（PASS_P15=TRUE）；save-confirm / References 触发路径 **UNOBSERVABLE**、英文 IDE 未测（如实披露）。冻结面不变：schema 两件 `AAFA1C6C…`、hub snapshot `828B4DAD…`、工具数 120/11/131。证据：`fix/TASK_DIALOG_IDFIRST_20261024/`（`PREREGISTER_39.md` / `STATIC_IMPL_39.md` / `RECORD_39.md` / `verify_39_result.json`）、`fix/TASK_RUNTIME_PROBES/WINDOW_46_RESULT.md` / `WINDOW_47_RESULT.md`。

> 注（2026-10-17）：上表历史行中的**机器专用路径**已替换为占位符（如 `<scratch>`），原文语义与结论不变；公共树不得含本机路径（导出审计要求）。
| **v1.5.0** | 2026-10-09 | **消息型窗口原语 + 可选滚轮 + 写入回执的相位/分块字段**：新增 `read_active_dialogs`（只读）/ `close_active_dialogs`（**只显式调用**）/ `send_window_message`（**只发窗口消息**），**106 → 109，必须显式重冻**；`command_bars` 在**既有 `control` 参数内**接受 `#<id>` 寻址（**不改 inputSchema**）；写入回执新增 `hunks` 与相位字段（**纯增量**，见下）；可选代码窗滚轮（env/注册表，**默认关**，**不新工具、不改 schema**） |
| **v1.4.0** | 2026-10-09 | **OCX 防挂死**：`add_control` 预校验（未登记部件 ⇒ `control_part_not_registered`，绝不进入会挂死的调用）+ 新只读工具 `list_component_parts`（**105 → 106，需显式重冻**） |
| **v1.3.2** | 2026-10-09 | **结构化错误号**（响应侧，无需重冻）：失败回执增可选 `vbErrorCode`/`hresult`（8 位 hex）/`errorLine`；`mcp_info` 可增 `errorDecoding`/`errorTableEntries` |
| **v1.3.1** | 2026-10-08 | **多 IDE 归属确定化**：`SO_EXCLUSIVEADDRUSE` + 动态端口 + 实例注册（`%TEMP%\vb6mcp\instances\<pid>.json`）+ 响应正文 `ide:{pid,project,port}`（不触 inputSchema、无需重冻；不变式不变） |
| **v1.3.0** | 2026-10-08 | **版本变更可见化**：新增只读 `mcp_info`（104 -> 105）+ `initialize.instructions` 固定版块 + 每次 `tools/call` 正文带 `mcp:{product,contract}` 版本戳 + 产品版本 0.2.0 -> 0.3.0 |
| **v1.2.0** | 2026-10-08 | **工具加固**（从真实工程实战反推）：α 回执诚实性（纯增量）+ β 行为加固（**破坏性**）+ γ 输入参数（触 inputSchema）+ δ 3 个新工具；分 A 批（α+β）与 B+C 批（γ+δ）两次落地，见下节 |
| ~~v1.1.0~~ **v1.1.0（含勘误）** | 2026-09-22 | **新增 5 个只读工具**：`ide_state` / `current_position` / `list_breakpoints` / `find_gui` / `ui_text`（96 → 101）；另有 1 条非 schema 变更（断点态写权限错误的 `reason`，见下） |

## v1.16.0 变更说明（add-in 面；纯增量，冻结面已重冻）

> **结论先行**：**工具数不变**（add-in **120** / hub **11** / DSH 面 **131**），**ADDED/REMOVED 工具 = 0**；新增 **7 个可选参数** + **breakpoints.action 多一个枚举值 `probe`**，**9 个工具描述**变长；**缺省路径行为不变**（不传新参数、不用 `probe` 时，回执与旧版一致）。本版在活体车道（最终 DLL `BB20AD87…`，contract v1.16.0、toolCount 120）上执行 `dump_tools_schema --freeze` → `export_tool_catalog` → `gen_hub_snapshot` 重冻冻结面。

### 0. 冻结面哈希与冻结语义（台账补登，2026-10-22）

- **逐字节冻结件（唯一一对）**：`docs/tools_schema.json` == `tools/schema_baseline.json` = `7FCC00AE6B6C8CE5BB9A08E64FA0D35B6A228E50EA26482D631DB260FD645355`（120 工具）。工具面任何变化必须走 freeze 流程并显式重冻。
- **生成物（不是逐字节冻结件）**：`bridge/hub_schema.generated.json` = `D95A85F6678838C640CCF7F1117D8B60102AD32296181DAFCDA2AE7BF6F0A250`（本次重冻锚点）。判据见 `tools/check_hub_alignment.py`：`faceFingerprint` 必须 FRESH + `tools[]` 逐字段同顺序相等；重新生成会改变自身 sha，属预期。
- `docs/TOOL_CATALOG.md` = `3D6A3CC1E246227BC72C65BD8E42F7DA39A80950ACBF92C39A23CFAF2566154A`；活体 DLL = `BB20AD876C152260C61A6B1953A5BC21A66B40EF98AE0181721C1AC66592A94C`（contract v1.16.0 / toolCount 120）。
- 细则与历史补记见版本表 `v1.16.0+冻结面台账补登` 行。

### 1. `breakpoints{action:"probe"}`：行为验证，且必须**两阶段**

VBIDE 对象模型没有断点 API，"这个断点有效吗"只能**行为验证**。一次调用**无法**观测它自己启动的运行（Start 是 STA 里的**排队命令**），所以本工具是两段：

| 阶段 | 入参 | 回执 |
|---|---|---|
| `phase:"start"`（**默认**） | `breakpoints{action:"probe", component, line, timeout_ms?}` | 先判定 design 态 / 无模态框 / 行可断；通过后**启动程序**（**不保存工程、不写源码、不做 pre-run flush**），回 `probePhase:"start"`、`probeResult:"started_pending"`、`probeId`、`startedPending:true`、`startQueued:true`、`howToContinue` |
| `phase:"observe"` | `breakpoints{action:"probe", phase:"observe", probeId, timeout_ms?}` | 等至多 `timeout_ms`，命中后读 `position{component,proc,paneLine}` + `callstack`，再 **End** 程序，并按**回读**证明 `stateAfter=design`；尚未观测到停机时回 `endPending:true`（不谎报已停），直到读回 `stateAfter=design` |

- `phase:"auto"` 是 `start` 的**显式别名**（不是"一次调用跑完两段"）。
- `probeResult` 是**封闭枚举**：`started_pending | stopped_at_target | stopped_elsewhere | did_not_stop | timeout | did_not_stop_unobserved | line_not_breakable | not_attempted | run_failed`。
- `probeId` **单次使用**、有 TTL（**600000 ms**）；未知/过期 ⇒ `ok:false + reason:"probe_not_found" + howToProceed`（不猜、不复用旧运行）。
- **代价（如实）**：它**真的跑程序**——observe 命中 = 等停 ≤`timeout_ms` + 可能一段 End 确认；实测一次完整验证 **6.2 s / 12.1 s**。它是唯一诚实通道，但**不是**廉价只读操作。

### 2. `control_execution{mode:"end"|"stop"}`：RUN 态不再**假拒绝**

旧实现在程序 RUN 时对 `end`/`stop` 回 `reason:"commandbars_unavailable_while_running"` —— 实测这是**伪拒绝**（End 命令条在 RUN 态返回 true 且真的停止程序）。现走 `runStateEndPath:"dbclick_end_while_running"`，回执新增：

`endCommandDelivered` / `endPending` / `stateAfterVerifiedBy:"ide_state"` / `stateAfterReadBack` / `programEnded` / `deliveryVsStateNote` / `endSettleMs`。

★`programEnded:true` **只在 design 回读之后**出现；End 已发出但本次调用内尚不可观测 ⇒ `ok:false + endPending:true + howToContinue`。**其它 mode 在 RUN 态的拒绝不变**。

### 3. `verify` / `verify_scope`：`requireCompiler`（默认关）

- `requireCompiler:false`（默认）⇒ 回执**逐字节不变**。
- `requireCompiler:true` ⇒ 必须拿**编译器**证据：对**盘上工程的副本**跑隔离 `VB6.EXE /make`，成功回执新增 `authority:"compiler"` + `errorCount` + `artifactChanged` + `compiler{path,logPath,errorCount,artifactChanged}`；`verify_scope` 取得的是**一次项目级**编译判定。
- **拿不到编译器**（解析不到 / 起不来 / 没写日志 / 预算到点被杀）⇒ **结构化拒绝** `ok:false + reason:"compiler_not_configured" + compilerProbe{resolved,path,why} + howToProceed`；**绝不用 lint 结论顶替编译结论**。该路径**不保存工程、不碰真产物**。

### 4. `apply_edits` / `preview_edit`：多组件批量（**同一接口**，不新增工具）

- `edits[].component`：每条 op 可带自己的组件名；工具级 `name` 变**选填**（作为默认值）。**两者都缺** ⇒ `ok:false + reason:"component_required" + opIndex`。
- `expected_shas[]`：`[{component, sha256}]` 前置守卫；**任一不符 ⇒ 整批拒绝、一字节不写**，回 `guards[{component,expected,actual,ok}]`（旧的 `expected_sha256` 继续守默认/单组件）。
- **顺序语义**：op 按 `edits[]` **数组顺序**应用，且**每个锚点在"该组件已更新文本"上解析**（前面的 op 可移动后面的锚点）；组件之间互相独立。
- **全或无**：任一 op 失败 ⇒ 把**已写入的组件**恢复到批前 pane 文本，并用**回读**证明恢复（`rolledBack` + `restoreVerified` + `restoredComponents[]`），然后 `ok:false + failedOpIndex`（**全局**下标）+ `hits[]`。
- **上限**：`max_ops` 默认 **200**、**上限 1000**；单批最多 **40 个组件** ⇒ 超限 `reason:"batch_too_large" + limits{ops,components,max_ops}` 且**不写**。
- **成功回执**新增 `appliedOps` / `totalOps` / `components[]` / `batchElapsedMs`。
- `preview_edit`：**同一** `edits[]` 形状，按组件分组 hunks（`components[{component,ops,hunks,hunkCount,algorithm}]` + `totalHunks`），**仍然什么都不写**（`writes:false`）。
- **实测收益只能表述为"省往返"**：12 ops / 4 组件 **812 → 376 ms**；40 ops / 8 组件 **3013 → 847 ms**。add-in 是**单线程 STA，没有并行** —— 变快来自一次调用完成多次往返，不是并发执行。

### 5. 随本版落冻的纯描述变更（此前已声明、未进冻结面）

`immediate_eval` / `immediate_run`（sink 失败分类措辞：`engine_refused_line` 与 `eval_sink_error` / `run_sink_error` 分开）与 `run_project`（STA 排队语义 + `diag{}` 块说明）的描述变长；**无参数变化**，属"声明有、冻结无"的滞后在本版一次补齐。

## v1.1.0 新增工具

### 1. `ide_state`
- 参数：无
- 返回：`{state:"design"|"run"|"break", ebMode:0|1|2（design=0 / run=1 / break=2）, title:"<主窗口标题>", project:"<工程名>", running:bool}`
- 语义：**实时**读取 IDE 模式；不得返回上一次动作前的陈旧值（design/run/break 切换后立即调用必须正确）。
- 只读。用途：所有调试类判据的前置条件。

### 2. `current_position`
- 参数：无
- 返回：断点态 `{available:true, project, module, proc, line, text, ebMode}`；非断点态 `{available:false, reason:"not_in_break"}`（**不报错**）。
- `line` 为模块内 **1-based** 行号；`text` 为该行原文（不含行号前缀）。
- 只读。用途：单步/继续的「位置变化」取证。

### 3. `list_breakpoints`
- 参数：无
- 返回：`{injected:[{component, line, kind:"injected", enabled}], native:null, note:"..."}`
- 语义：枚举**本工具注入的**断点标记（`set_breakpoint` / `native_breakpoint` 写进源码的标记行）；
  IDE 原生断点**不可枚举**时必须返回 `native:null` 并如实说明（VBIDE 对象模型不提供断点 API），**不得伪造**。
- 只读。用途：断点类证据从「自述」变「可枚举」。

### 4. `find_gui`
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

## v1.1.0 非 schema 变更

断点态下**写源码类**工具（`set_breakpoint` / `clear_breakpoints` / `apply_edits` / `write_source` 等）返回的裸
VB6 权限错误（`拒绝的权限`）必须追加 `reason:"ide_busy_in_break_mode"` 与恢复建议（先 `end` 退出断点态）。

**v1.1.0 勘误 2（2026-10-08，真机实测）**：**运行态**写源码同样会被拒（裸 `拒绝的权限`），追加值为 `reason:"ide_busy_in_run_mode"` + `ideState`。属**新增取值**（不改既有值语义）。

**v1.1.0 勘误 3（2026-10-09，源码复核）**：§3 的语义里「`set_breakpoint` / `native_breakpoint` 写进源码的标记行」**不成立**。
实现（`CodeTools.cls` `ToolListBreakpoints`，只扫描 `' @mcp-break <id> begin` 标记行）⇒ **只有 `set_breakpoint` 注入的断点会被列出**；
`native_breakpoint` 走 Debug 菜单的 toggle，**不写任何源码标记**，因此**不会**出现在 `injected[]` 里。
正确口径：`list_breakpoints` 只枚举 **set_breakpoint 注入物**；**IDE 原生断点全局不可枚举** ⇒ 一律 `native:null` + `note`（description 早已如此，本次只修契约正文）。
该勘误**不改 inputSchema/输出字段** ⇒ **不需要重冻**。

### v1.11.9 字段语义（本轮新增/澄清，调用方按此读回执）

| 字段 | 工具 | 语义 |
|---|---|---|
| `inDeclarationArea` | `add_code` | **声明区**判定：落点 `insertedAt.ide < 第一个过程头行`（VB6 只允许 `Option Explicit` / 模块级 `Dim`/`Const`/`Type`/`Enum` 出现在第一个过程之前）。因此**模块末尾追加 = `false`**、**首个过程之前插入 = `true`**；模块无任何过程时恒 `true`。旧语义（“不在未闭合过程体内”）已废弃。 |
| `requestedRange` | `replace_code` / `add_code` / 守卫拒绝 | `{mode:"range", operation, start_line, end_line}` |
| `requestedRange` | `edit_by_anchor` | `{mode:"anchor", anchor, occurrence, match, scope?}` |
| `requestedRange` | `apply_edits` | `{mode:"anchor-batch", ops, match}` |
| `countsConsistent` | `verify_scope` | 回执自证算术：`checked == pass + fail + unknown`。 |
| `pass` / `fail` / `unknown` | `verify_scope` | 三者**互斥**：`pass` = `verify` 判 PASS；`fail` = 判 FAIL（**已证不一致**）；`unknown` = 判不出（如仅 build 证据缺失）——**不计入 `fail`**。`failures[]` 只含已证不一致，`unknowns[]` 只含 unknown，**同一组件不会同时出现**。 |
### v1.11.10 字段语义（证据尺度澄清，不改行为）

| 字段 | 工具 | 语义 |
|---|---|---|
| `buildOk` / `buildSource` / `buildNote` | `verify` / `verify_scope` | ★**build 证据按【项目】判定，与组件无关**：三者来自**项目级** build 缓存（`ToolVerify` 读 `Build.CacheLoad`，key/hash 都是**项目级**），同一工程内**所有组件共享同一份** build 证据。因此 `buildSource: "none"` / `buildNote: "no build has been recorded for this project; run build for a real compile verdict"` 是**项目级**结论 —— **单个组件不可能“仅自己缺 build 证据”**；在**从未构建**的工程里，**任何**一致组件都会是 `unknown`（`pass` 不可达）；`pass` 只能在**有 fresh build cache 的工程**上出现。`buildOk` = `-1` 表示 unknown、`0` 表示 FAIL、`1` 表示 PASS（同样是**项目级**判定）。 |

★**判据提醒（核对口径，不是新行为）**：在**从未构建**的工程上，`verify_scope` 的 `pass` **恒为 0**（`pass + fail + unknown == checked` 依然成立，全部落在 `unknown`）—— 这是**已文档化的正确行为**（项目级 build 证据缺失），**不是回归**；要拿到 `pass`，先在该工程上跑一次 `build` 生成项目级 build 缓存。

### v1.15.0 回执字段补充（无 schema 变化、无版本 bump）：新建组件的初始声明区 / `add_code` 的「成败」与重复声明 / `immediate_eval` 的值传输

> 本批**只做两件事**：新增**回执字段**，以及在**写入前**多做一次检查。没有新增/改名工具，没有新增或修改参数，`inputSchema` 与工具描述未变 ⇒ 契约版本保持 **v1.15.0**，冻结面**不需要**重冻。

#### 1) `add_component`：新建组件的初始内容可预知

- 固定字段（**每一次成功回执都有**，含 `name` 缺省与设计器类型这两条早期返回路径）：
  - `component`：组件名（与既有 `created` 同值，供按通用字段名读取的调用方使用）；
  - `initialDeclarations[]`：**新建组件代码窗格**在**第一个过程头之前**的声明区行（去首尾空白、按窗格顺序、跳过空行）；
  - `initialDeclarationsCount`：上面的条数；
  - `initialDeclarationsRule`：声明区口径的完整文字；
  - `filePath` / `componentFile`：组件成员文件路径（工程没有盘上文件时为空串）；
  - `fileMaterialized`：回执生成时该文件是否已在盘上；
  - `fileSizeAfterWrite`：该文件的**实测字节数**（`FileLen`）；`0` 且 `fileMaterialized:false` 表示“文件当时不在盘上”，**不是**“没测”；
  - `fileSizeAfterWriteBasis`：上一字段的口径原文。
- ★**默认内容**：VB6 的 `VBComponents.Add` 会给新模块种一条 `Option Explicit`，因此 `add_component{type:"module"}` 的 `initialDeclarations` 通常是 `["Option Explicit"]`。
- ★**口径**：`initialDeclarations` 只描述**代码窗格**。`Attribute VB_Name = "..."` 是**磁盘组件文件**的头行，**不出现在代码窗格**，因此**不会**出现在该数组里 —— 它随工程保存写入组件文件。
- 调用方因此**不需要先读回组件**就能决定“要不要再补一条 Option Explicit”。

#### 2) `add_code`：`ok` == “是否写入”；重复声明在**写入之前**检查

- ★**`ok` 只有一个含义：本次调用是否改动了代码窗格。** 新增字段：
  - `written`（boolean）、`writeOutcome`（`appended` | `not_written`）；
  - `ok = written`，`okBasis` 把这条口径写进回执；
  - 未写入 ⇒ `ok:false` **且一定有 `reason`**（当前取值 `no_write`），不再出现 `reason` 缺省。
- ★**内容校验与写入成败分离**：`landedVerified` 仍是“结果**恰好**等于旧文本 + 追加块”的逐字节判定（语义未变）；当**写已发生**而该判定失败时，`ok` 仍为 `true`，并用
  - `landingDifference`：`trailing_line_break_only`（IDE 只规范化了尾行断行）| `content_mismatch`（内容确有差异）；
  - `landingNote`：说明“写已发生、不要重复追加、用 `sha256After` / `newTotalLines` / `appliedRange` 核对”。
  - 依据（实测）：`InsertLines` 追加后 IDE 可能规范化尾行断行，旧的 `ok = landedVerified And …` 会把**已写入**报成 `ok:false`，调用方因此重复追加。
- ★**重复声明前置检查（写入之前）**：扫描“当前窗格文本 + 本次提交块”的声明区，检查
  - `Option Explicit`（模块只允许一条）；
  - 模块级 `Dim` / `Const` / `Type` / `Enum`（以及 `Declare`）的**同名**冲突（大小写不敏感）。
  回执固定带 `declaredCheck`（`performed` | `not_performed`）、`declaredCheckBasis`（口径与限制）、`duplicateDeclarations[]`、`duplicateDeclarationCount`、`duplicatePolicy`。
  - `validate:"strict"` 且有命中 ⇒ **不写**：`ok:false`、`reason` 为 `duplicate_option_explicit` 或 `duplicate_declaration`，并给 `sha256Before`/`sha256After`（两者相等）、`newTotalLines`、`readbackUnchanged:true` 作为“内容未变”的证据；
  - `validate:"warn"` / `off` ⇒ **照常写入**（`ok:true`），每个命中作为 **`severity:"warning"`** 的诊断进入既有 `diagnostics[]`（`ruleId` = `duplicate_option_explicit` | `duplicate_declaration`，`source:"add_code_precheck"`）；它**不会**变成阻断写入的 error。
  - ★**限制（如实）**：文本级检查，不展开 `#If` 条件编译分支，不判“声明落在了过程之后”（那是另一个编译错误），也不是编译器。
- ★`newTotalLines` / `insertedAt{ide,disk}` / `sha256Before` / `sha256After` 语义未变。

#### 3) `immediate_eval`：多行字符串**无损**传输

- 值现在走**三条** sink 行：①旧有的 Print 格式化值；②`<<MCPLEN>>Len(CStr(值))`（**完整**长度，先打印，故 sink 被截断也能读到）；③`<<MCPV>>` + 把反斜杠 / CR / LF 转义后的**完整**值（单行，Print 无法把它拆行）。
- 新增字段：
  - `valueEncoding`：`raw`（单行标量，`value` 保持旧的 Print 语义与首行取值，**既有单行行为逐字节不变**）| `json`（**多行**结果：`value` 是**完整**文本，回执 JSON 把 CR/LF 转义，`JSON.parse` 后即为原文，**不要**按行切）；
  - `valueBytes`：**完整**值的长度（`Len(CStr(值))`）；`valueBytesReturned`：本次实际交付的 `value` 长度；
  - `valueLines`：完整值按 CR/LF 切分的行数（尾行断行也算一行）；被截断时为 `-1`（未知）；
  - `valueComplete`：`value` 是否就是完整值；
  - `valueEscaped`：转义形式（反斜杠 `\\`、CR `\r`、LF `\n`），供只能处理单行 ASCII 的调用方；截断时带 `valueEscapedPartial:true`；
  - `valueTruncated`（boolean）+ `valueTruncatedReason`（截断原因）+ `valueSinkTruncated`（sink 读取是否触顶）；
  - `outputDescribes`：★`output` / `outputBytes` / `outputLines` 描述的是**第一阶段 sink（TypeName 通道）**，**不是返回值**；返回值只看 `value*` 字段。
- ★**截断语义（绝不给半个值假装完整）**：确认真截断时 `value` = `null`，`valueBytes` 仍是**完整**长度，`valueTruncated:true` 且给原因（sink 读取上限 65536 字节，或转义行比独立打印的完整长度短）。
- ★`Null` 结果：`CStr(Null)` 会抛错并弹 IDE 模态框，因此 `Null` 走**常量标记**路径（`"Null"` + 长度 4），与旧的 `Print #` 行为一致；`Empty` 同理为长度 0。


### v1.15.0 hub 侧回执字段补充（无 schema 变化、无版本 bump）：默认目标阶梯的 live fallback

> 只改 **hub 的默认目标解析行为** 与回执里**新增**的诊断字段；hub 工具面（11）、add-in 面（120）、DSH 面（131）与所有 `inputSchema` **不变** ⇒ 契约版本保持 **v1.15.0**，冻结面**不需要**重冻。

- **缺陷（t16 实测）**：同一 hub 同一时刻，`run_project{observe:false}`（不带 `_target`）两次回 `reason:"no IDE instance listening (scanned 34 default ports + registry)"` + `suggestedAction:"ide_start"`，241 ms 后 `ide_state{}` 却成功（pid 9252）；带 `_target` 的调用始终正常。两者走**同一条**路径（`handleCall → resolveLocalPick → pickTarget`，`run_project` 在 hub **没有**专门分支），差别是**时机**：`discoverSoft()` 只等 400 ms，而一次发现可能几秒；空发现结果还会按 `DISCOVERY_TTL_MS`（15 s）缓存 ⇒ 阶梯把「我还不知道」当成「没有活实例」。
- **修复（行为）**：阶梯在回答 `no_candidate` **之前**做两件有界的事：① 等**正在运行**的那一次发现（等待时长报在 `waitedMs`，上限默认 2500 ms，env `VB6MCP_LADDER_EMPTY_FACE_WAIT_MS`）；② 仍为空则读**直连源**（hub 自己的注册文件**当场**从盘读 → 仍为空再强制一次**全新**发现：注册表 + middle service `/health` + 端口扫描）。显式 `_target` 不受影响（原样下传）。
- **新增回执字段（ADDITIVE）**：`resolution.ladderLiveFallback{used,sources[],waitedMs,elapsedMs,rule}`（转发回执里同一块在 `targetResolution.ladderLiveFallback`）。`sources` 取值：`in_flight_discovery_pass[:timed_out|:none]`、`direct_registry_read[:no_live_instance]`、`forced_fresh_discovery_pass[:timed_out_or_empty]`。`used:false` 表示两条直连源都查过且为空（`no_live_ide` 拒绝照旧，**不会**凭空造出目标）。
- 附带硬化：被 `invalidateDiscovery()` 取代的旧发现**不再发布**其结果（旧实现会让一次被作废的慢发现覆盖掉更新的 face）。
- 回归：`tests/hub_default_target_live_fallback_test.mjs`（13 条，无需 VB6；假 add-in 经**端口扫描**被发现）——修前 hub 同测试 5 PASS / 8 FAIL，失败形状即 t16 原文。

- **第五轮补正（t16 复议）**：
  - `ladderLiveFallback` 现在**无条件**出现在走过「空 face → fallback」路径的**任何**回执上：顶层 `ladderLiveFallback`（转发回执亦同；`mergeNote` 只填缺省顶层键，所以不会被上游 body 遮蔽），并在有 `resolution` 时带同内容的 `resolution.ladderLiveFallback`。覆盖转发/各类拒绝（含 python 模式 0 IDE 的 `no_live_ide`、remote-only、MULTI_BRIDGE 跨桥）/离线回执。
  - 字段扩充（ADDITIVE）：`ladderLiveFallback{used, sources[], waitedMs, elapsedMs, budgetMs, complete, timedOut, rule}`。`budgetMs` = 本次 fallback 的**总**预算（`VB6MCP_LADDER_EMPTY_FACE_WAIT_MS`，默认 2500 ms）；三段（等 in-flight 发现 / 直连注册表 / 强制 fresh pass）**共用**该预算；用尽即返回 `complete:false + timedOut:true`（找到实例则 `used:true`，否则 `used:false`），**不会**等一个不完成的发现。
  - 观测口径：`used:false + complete:true` = 直连源都读完了、确实没有实例；`used:false + complete:false + timedOut:true` = 预算用尽、无法确认（回执 `reason` 仍是 `no_live_ide`，`howToProceed` 照旧）。
  - 测试钩子（仅测试）：`VB6MCP_TEST_DISCOVERY_NEVER_RESOLVE=1` 让发现永不完成，用于证明总预算仍回执；生产不得设置。
  - 第六轮补正（转发预算可见）：`forwardBudget{budgetMs,plusSlackMs,env{name,set,value},rule}` 出现在**任何**转发回执上（成功与 `forward_timeout` 都有）——`budgetMs` 默认 python 模式 `FORWARD_TIMEOUT_MS + 10000`（inproc 直连为 `FORWARD_TIMEOUT_MS`），env 覆盖时 `plusSlackMs:0` 并在 `env{name,set,value}` 里点名；`probePending:true` 表示选中的实例**还没应答过 probe**（正是会吃满预算的状态）；`selectedBy:"ladder_live_fallback"` 表示候选来自空 face fallback 的直连源（读取来源仍见 `ladderLiveFallback.sources`）。`forward_timeout` 另带 `failureKind:"wedged"`、`retrySafe:true`、`howToProceed[]`（`_target` 直达 / 先 `ide_state` 探活 / **不要**盲目重试）与 `retryNote`。**恢复语义不变**：implicit 目标在 probe pending 时仍会转发（那是 wedged IDE 唯一的恢复通道），本批只把预算与状态变成可见的。
  - 第七轮补正（B1/B7/g1，仍 v1.15.0，无 schema/工具面变化）：
    - **`ide_restart` 的 `resolve_target` 用与转发同一套探针**（`waitForAddinReady`，与 `ide_start` 同构）：注册表里有条目、pid 活、但 add-in 未应答时**先轮询**（上限 `VB6MCP_RESTART_PROBE_READY_MS`，默认 15000），不再从一次软发现就判 `instance_found_but_unresponsive`。`wait_ms` 只约束**本次调用**，因此 `pending → opId → done/refused` 成为常规路径（`pending` 可达）。等不到才拒绝，且必带 `probeChecks`（与 `readyChecks` 同构：`registryVisible/transportUp/addInAnswers/addInAnswersBasis/pidAlive/projectMatches`）+ `probeBudgetMs` + `readyWaitMs` + `howToProceed[]` + `retrySafe:true`/`suggestedAction:"retry_once"`。
    - **显式 `_target` 命中「注册表有条目但 pid 已死」→ `reason:"peer_dead"`**（`failureKind:"peer_dead"`，`reasonLegacy:"no_ide_running"` 保留兼容）：`peerDead:true` + `peerState{pid,pidAlive:false,registrationPresent:true,port,transport,project,projectPath,registryFile,startedAt}` + `lostContext{debugger:null,debuggerBasis,projectPath,note}`（`debugger:null` = 未知，不是 design）+ `howToRecover[]`（重启该工程 / 换活实例 / 清残留注册）；只有**确实没有任何候选**才回 `no_live_ide`。
    - **隐式目标的转发上限**：`VB6MCP_IMPLICIT_FORWARD_TIMEOUT_MS`（默认 15000）。仅当调用**没有显式 `_target`** 且（选中实例 `probePending` 或候选来自 ladder 空 face fallback）时生效；到点必回执 `forward_timeout` + `forwardBudget{budgetMs,baseBudgetMs,implicit:true,...}` + `implicitForward:true`/`implicitForwardCapMs` + `probePending` + `retrySafe:true` + `howToProceed[]`（显式 `_target` 重试 / 先 `ide_state` 探活）。**显式 `_target` 保持长预算**（wedged `release_inflight` 恢复通道）。
  - 第八轮补正（按活体裁定收口，仍 v1.15.0、无 schema 变化）：
    - 隐式转发 cap 判定：!explicit 即 cap（不再要求 probePending/ladder fallback）；forwardBudget.decidedBy{explicit,probePending,ladderFallbackUsed,capApplied,capMs} + 顶层 forwardBudgetDecidedBy 记录判定输入，forwardBudget.baseBudgetMs 是未 cap 的原预算。
    - peer_dead 的三路来源（peerState.source）：registry_file（直读原始注册文件，默认）/ discovery_stale / middle_service_health；另给 peerState.hubRegistryDir、hubRegistryReadable。
  - 第九轮补正（B1 注册可见性，仍 v1.15.0、无 schema 变化）：
    - `ide_start` / `ide_kill_and_relaunch` 回执新增 `registration{written,by,path,existed,readbackVisible,owner,reason}`（hub 只在 **add-in 已在 pinned 端口应答** 且 `<pid>.json` 不存在时写一条最小注册，`by:"hub"`；否则 `written:false` + 原因），`ide_start` 另加 `registrationNote`（add-in 未注册时的处置说明）与 `readyPortSource`（`registration|pinned_hint|none`）。
    - readiness 现接受 **pinned-port hint**：注册文件缺失时用 hub 自己 pin 的端口探 `mcp_info`；`readyChecks.portSource` 标明通道，`registryVisible` 仍如实为 false。
  - 第十轮补正（FIX-26：活体验收挖出的 hub 缺陷收口，仍 v1.15.0、无 schema/工具面变化）：
    - **`ide_restart` 的 relaunch 命中 PIPE-ONLY 新实例**：新实例自己的注册项若 `port===0` 且（`transport==="pipe"` 或 `tcpEnabled===false` 或带 `pipe`），`wait_instance` 不再空等满 120 s 再报 `new_instance_never_came_up`（那是**假话**：IDE 已经重启），而是**快速**返回 `refused_reason:"new_instance_is_pipe_only"` + `restarted:true` + `ok:false` + `newInstanceFacts{pid,port,transport,tcpEnabled,pipe,project,projectPath,registryFile,startedAt,registered,addressableByMiddleService}` + `note`（明说 "the IDE WAS RESTARTED"）+ `howToProceed[]`（起 hub 时带 `VB6MCP_TRANSPORT=tcp`，或改用 python 模式让中间服务走 pipe）。成因：add-in 默认轨是 pipe（本机无 `HKCU\\Software\\VB6MCP\\Transport`），而 relaunch 继承 **hub 的**环境。
    - **`ide_restart` 的 relaunch 改用 `spawnProgram`**：`.cmd/.bat` 形态的 `VB6MCP_VB6EXE` 之前被裸 `spawn()` 直接失败（op 死在 step `spawn`，`reason:"restart_error"`），现在与 `ide_start` / `ide_kill_and_relaunch` 一致（对 `.exe` 字节级等价）。
    - **`ide_restart{opId}` 续查回执带终态结果**：除进度标量外，另含 `ok/restarted/stopConfirmed/saveOk/registryCleared/pinPort/projectVbp/ownership/ownershipBasis/ownedByHub/truth/note/howToProceed/newInstanceFacts/probeChecks/manualCommand/nextStep` 等（从 op 的 `out` 取，缺省即不出现）——否则「先 pending 再续查」的调用方永远看不到结果。
    - **pending 回执的 `reason` 改为 `pending_op`**（`restart_in_progress` 仅保留给「同一 hub 同一时刻已有 restart 在跑 ⇒ 第二次 restart 被拒」），两者不再同词。
    - **`ide_start` reuse 的 `readyChecks.projectMatches` 现在是真比对**（add-in 的 `ide_state.project` vs 复用注册项的 `project`），另带 `projectMatchesActual/projectMatchesExpected/projectMatchesBasis`；比不了时 `null` + `projectMatchesBasis` 以 `nullReason:` 开头；比到**不相等**时 `ready:false`（四项不全就不发 ready:true）。
    - **`ide_start` 的两条拒绝回执带 `processesCreated:0` + `attemptedSpawn:false`**（`vb6_exe_not_found` / `project_not_found`），不再让 `processesCreated` 为 null。
    - **`forwardBudget` 新增 `capSource`（`env|default`，仅 implicit）与 `capEnvName`；`VB6MCP_IMPLICIT_FORWARD_TIMEOUT_MS` 生效时 `env{name,set,value}` 点名它**（此前 capMs 已变但 env 仍 null）。
    - 断言：`tests/hub_default_target_live_fallback_test.mjs` R0–R8 / S1–S4 / T1–T3，其中 R6/R7 是对 pinned 旧 hub（0528f0a）的**变异对照**（旧代码 6 s 时仍在跑，且永不产出 `new_instance_is_pipe_only`）。



### v1.15.0 参数归属澄清（无 schema 变化、无行为变更）：`diagnose` 只在 `make_project` 上，`build` 没有

- `make_project{diagnose:true}` = 在 **IDE 之外**做一次**副本隔离编译**（`VB6.EXE /make` 打到临时目录），拿真编译器日志与 `firstError{file,line,text}`；它还会在「in-IDE Make 说 `compiled:true` 但 `artifactChanged:false`」时**自动**跑一次（那种回执否则无法区分「编译失败」与「编译了但产物没变」）。
- `build` **没有** `diagnose` 参数：它自己的正常路线本来就是隔离 `VB6.EXE /make` + `/out` 日志，不存在第二层「诊断」开关。传 `build{diagnose:true}` 得到的是**结构化拒绝** `reason:"unknown_argument"` + `acceptedArguments[]`（+ hint），**不是**静默忽略，也**不是** `unexpected_error`。
- 判据：**某工具接受哪些参数，以该工具自己的 `acceptedArguments` / `inputSchema` 为准，不要按兄弟工具猜** —— `build` 与 `make_project` 是两个工具，参数集本来就不同（例如 `output_path` / `route` / `wait_for_build` / `diagnose` 只在 `make_project` 上）。
- 选路：要诊断一个 **IDE 正持有**的工程 ⇒ `make_project{diagnose:true}`（或对该工程的**副本**调用 `build`）；要拿被构建工程的真编译器日志 ⇒ 直接用 `build`。

### v1.15.0 参数/能力边界（FIX-29，无 schema 变化）：immediate_* 的 1000 字符上限与恢复语义

- `immediate_eval` / `immediate_run` 只接受**一行**；**上限 900 字符，量的是真正执行的那一行**
  （表达式 + `Open <sink> For Output As #511` / `Print #511,` / `Close #511` 包裹，约 +123 字符），常量 `IM_MAX_EXEC_LINE_CHARS`。
  本机实测：**执行行 ≤900 稳定 33/33**；表达式 901–998（执行行 1024–1131）**14/15 失败**（10 次 `value_not_parsed`、
  4 次 `engine_refused_line` ret 40036，且**非单调**——905 有一次 OK），符合立即窗口/EbExecuteLine 行缓冲 ~1024 截断；
  表达式 1000/2000/2040/2047 → `ret:0` + 空 sink；**2048 → VB6 进程消失并重启**，遗留 `#32770` 运行时错误 55「文件已打开」。
  超限时本工具**执行前**拒绝：`reason: statement_too_large` + `limitChars` / `limitBasis: full_exec_line` /
  `actualChars`（=执行行长度）/ `actualBytes` / `expressionLimitHint` / `howToProceed`，并同样发布 `sinkEntryClose*`。
  1000 / 2000 / 2040 / 2047 → `ret:0` + 空 sink（表现成 `value_not_parsed`，**不是**显式截断）；**2048 → VB6 进程消失并重启**，
  遗留 `#32770` 运行时错误 55「文件已打开」。超限时本工具**执行前**拒绝：`reason: statement_too_large` + 
  `limitChars` / `actualChars` / `actualBytes` + `howToProceed`。
- 每次调用**入口无条件 `Close #511`**（不再依赖上一次的记录），结果见 `sinkEntryCloseAttempted` / `sinkEntryCloseRet` /
  `sinkEntryCloseClosed`；语句行尾仍保证 `Close #511`。
- **书面边界**：若超长语句通过别的途径发出（或本限制被绕过），**IDE 可能需要重启**才能恢复 immediate 通道；
  重启后通道恢复，代价是该 IDE 会话里未保存的内存改动会丢失（先 `save_project`）。

### v1.11.10 传输修复（根因记录 — 为什么过去会「间歇失败」）

活体面（HTTP/winsock）过去会间歇性地「**无响应**」或「**响应被截断**」。根因在插件侧 winsock 传输，共三处，均已修复（**只动 winsock 传输；PipeTransport 与冻结流程语义未动**）：

| # | 缺陷 | 机制 | 实测形状 |
|---|---|---|---|
| 1 | **连接对象未持久化** | `DoAccept` 把 socket 句柄(Long)放进 `m_conns`，而 `GetConn` 用 `Set c = m_conns(key)` 取对象（Long 不是对象 ⇒ 失败）⇒ 每次 FD_READ 都新建 `HttpConn` ⇒ **跨 TCP 段到达的请求丢掉前半段**，永不成完整请求 | 请求拆两段发送 ⇒ **0/6 成功**，30 s 后被空闲清扫关闭 ⇒ 客户端 `transport_idle_closed`（RemoteDisconnected @ ~35 s）；修后 6/6 |
| 2 | **close-after-send 判定顺序** | `DoFlush` 先 `ClearOut`（会删 `m_closeAfterSend`）再 `IsCloseAfterSend` ⇒ `Connection: close` 的连接不立即关，靠 30 s 清扫兜底 | 连接滞留 + 每 30 s `closing idle connection` 日志 |
| 3 | **发送路径两处** | ① 旧 flush 用累计 `pos` 做 `Mid$(pending, pos+1, 32768)`（起点随累计增长，>64 KiB 后拿不到可用块）；② `send` 返回 `n=-1` 且 `WSAGetLastError()=0` 时被当成致命错误直接关连接 | 响应稳定截断在 **65536 / 98304 B**（2-3 个 32 KiB 块），客户端 `IncompleteRead` / `truncated_response`；修后回放 12/12、裸 socket 12/12、真实客户端 15/15（零重试） |

修复后行为：flush 以「切掉已发前缀」推进（`Mid$` 起点恒 ≤ 32769）；**无 winsock 错误（`werr=0`，含 `n=-1`）一律视为 DEFER**（保留剩余、等 FD_WRITE 重试），只有真实 winsock 错误才关连接；异常路径落 `flush defer` / `flush error` 日志。客户端 `tools/dump_tools_schema.py` 失败给结构化 `reason` 且最多自动重试一次。

跨机 live 冻结的权威路径仍是：**t16 本机 loopback 抓 `tools/list` 原始字节 → 本机回放 → 产品脚本 `--freeze`**（本轮传输修复后**直连 `--freeze` 也已成功**，两种路径等价；见 docs/LIVE_FACE_TRANSPORT.md 的 runbook）。

### v1.11.11 发布后修复：bridge 源在 v1.11.8–v1.11.11 之间不可启动（已修 + 已加闸门）

`bridge/vb6mcp_bridge.py` 自 `1cf4bec`（v1.11.8 的内部号清理）起是**不可导入**的：`_send` 的 docstring 行被拽到列 0（`IndentationError: expected an indented block after function definition`），文件另有 UTF-8 BOM（BOM 更早就在，不是启动阻断者）。**症状只有一条：中继服务起不来**——VB 工程照常编译、闸门全绿（当时为 7 道）、导出审计 `AUDIT_OK`（bridge 源不在公开导出树），所以 v1.11.8→v1.11.11 四个提交都没发现。项目侧启动失败后自行修复并重启，本仓库采纳其修复（docstring 归位 + 去 BOM）。

**为什么漏了**：闸门里没有任何一步编译 `.py`/`.mjs`；既有「source encoding」门只覆盖 VB 构建组件。

**已堵**（`tools/check_source_encoding.py` 扩展；并新增**第 8 道门禁** `tools/check_source_encoding_selftest.py`，闸门集为 **8/8**）：① 编译每个 `bridge/*.py` 与 `tools/*.py`；② 对每个 `bridge/*.mjs` 跑 `node --check`；③ `.py`/`.mjs` 上的 UTF-8 BOM 判 **ERROR**（markdown 上的 BOM 判 WARN）；④ 第 8 道 = 编码守卫的**负样本自证**：BOM / UTF-8 组件 / GBK 字面量在 `--strict` 下必须判 **FAIL**，合规树必须 **PASS**（10 cases，随 `tools/run_gates.py` 一起跑）。负向臂实测：插入缩进破坏 ⇒ `VERDICT=FAIL rc=1`（报 `文件:行`）；插入 BOM ⇒ `VERDICT=FAIL rc=1`。

### v1.11.12 字段语义（归属两轴 + 写类行序 + 挂死可观测）

**归属的两条轴（作者编号 114）—— 同一个实例、两个不同的问题，必须分开读**

| 回执字段 | 回答的问题 | 取值 | 谁用它 |
|---|---|---|---|
| `ownership` / `ownedByHub`（legacy，**未变**） | 「**是不是本 hub 启动的**」 | `hub_started` / `registry_started_by_hub` / `not_owned`；bool | **许可轴**：只有它决定 `ide_stop` / `ide_restart` 敢不敢动手 |
| `ownershipVerdict`（新） | 「**这个实例到底是谁的**（归属身份）」 | `verdict: owned / foreign / unknown` + `basis[]` + 证据字段 | 诊断轴：`ide_list`、转发回执 `ide_state`、`ide_restart`（`ide_stop` 亦附）**同一块、逐字一致** |

★ 两轴**不是**同义词：**归属不授予许可**。用户手工启动的本机 IDE ⇒ `ownershipVerdict.verdict=owned`，但 `ownedByHub=false` ⇒ `ide_restart` 仍以 `not_owned` 拒绝（这是权限问题，不是身份问题）。

**`ownershipVerdict` 逐字段**

| 字段 | 语义 |
|---|---|
| `verdict` | `owned` = 可归因**本机 add-in 安装**（hub 启动证据 或 注册表 + pid 存活 + **add-in 亲自应答** `mcp_info`）；`foreign` = **仅正面证据**（见下四类）；`unknown` = 其余一切 |
| `basis[]` | 每条**正面事实**一句话（pid 存活 / 注册文件 / 端口是否可达 / netstat 端口 owner / 是否应答），**可审计**：判 `unknown` 时它会明说「没有正面证据」 |
| `registryAlive` | 注册条目存在**且**该 pid 存活 |
| `portAvailable` | 端口**是否可得**；`false` **不等于** `foreign`（pipe-only 注册故意发布 port 0，端口拿不到不是别人的证据） |
| `pid` / `port` / `project` / `projectPath` / `registryFile` / `registeredPid` | 被判定实例的身份，以及**它自己那条注册记录**（两者可能不同，见 `foreign` 第 4 类） |
| `portOwnerPids` / `netstatAvailable` | 监听端口的 OS owner；`null` = **未知**（netstat 不可用），**不是**「没人」 |
| `ownedEvidence[]` / `foreignEvidence[]` | 支撑 verdict 的正面证据列表（分开列，便于机器判定与人工复核） |
| `rule` | 判据全文（`OWNERSHIP_VERDICT_RULE`），随回执下发 |
| `registryPresent` / `pidAlive` / `probeAnswered` / `registryStartedByHub` / `inMemoryStartedByHub` | 构成 `basis[]` 的原始布尔事实（注册是否可见 / pid 是否活 / add-in 是否亲自应答 / 注册是否说 hub 启动 / 内存表是否记录） |
| `portOwnerWhy` | **端口所有权探测的"为什么"**（5 个取值，见下表）；它是 `foreignEvidence` 的解释器，不是独立的判定依据 |

**`portOwnerWhy` 的 5 个取值（逐一含义 + 与 `portAvailable` / `netstatAvailable` / `verdict` 的关系）**

| `portOwnerWhy` | 何时出现 | `portAvailable` | `netstatAvailable` | 对 `verdict` 的关系 |
|---|---|---|---|---|
| `no_port` | 实例**没有端口**（pipe-only 注册故意发布 port 0） | `false` | `false` | **不构成 foreign**：端口不可得 ≠ 别人的 |
| `netstat_disabled` | 环境设了 `VB6MCP_HUB_NO_NETSTAT=1`（端口可能有，但 OS owner 不可查） | 由 TCP 探测决定（`true` / `false` 都可能） | `false` | **不能据此判 foreign**：这是**缺证据**，不是负证据 |
| `listening` | 端口**在监听**，`portOwnerPids` 给出 OS owner —— **判 foreign 时最常见的取值** | `true` | `true` | 是 **foreign 第 2 类证据**（OS owner ≠ 注册 pid）的载体；`foreignEvidence` 在此点名 OS pid |
| `not_listening` | 注册端口**无人监听**（`ide_list` 会把该注册放进 `staleRegistry{reason:"port_not_listening"}`） | `false` | `true` | 本身**不构成** foreign（除非另有正面证据，如 pid 已死 + 端口在别处应答） |
| `netstat_failed` | netstat **抛错 / 返回 null**（**异常态**；本版**未实测**，代码路径存在） | 由 TCP 探测决定 | `false` | 缺证据 ⇒ **不得判 foreign** |

**`foreign` 的四类正面证据**（只在这四类时给出，其余一律不得判 foreign）：① 注册 pid 已死但端口**仍应答**；② 监听端口的 **OS owner ≠ 注册 pid**；③ 已知端口落在**保留区**（< 8767，8766 属外部进程）**且应答**；④ 实例的监听 pid ≠ 它自己注册记录里的 pid。

**`unknown` 与 `ide_restart` 的处理**：`ide_restart` 的 `not_owned` 拒绝**保留**（键集固定）；归属**未知**时改用 `refused_reason:"ownership_unknown"` + `legacyReason`（改造前该路径会给出的旧 reason，供只认旧词表的消费者比对）+ `ownershipUnknown:true` + `howToDisambiguate`（**默认拒绝重启** + 消歧步骤：精确 pid 与 project / 让实例可被应答 / 清理残留注册文件）。**把「不知道」说成「不是我们的」，正是本版要消除的缺陷。**

`legacyReason` = 该拒绝路径在改造前会给出的**旧词表值**（`not_owned` / `not_ours_by_red_line` / `instance_found_but_unresponsive` 之一，**视路径而定**）；只认旧字符串的消费者用它比对，新代码用 `refused_reason`。

**写类行序（作者编号 113）**：`reload_component` / `save_project` 在保存后若 `.vbp` 行序变了**必须恢复原序**（只重排行、不丢 VB6 写入的内容）并报 `vbpOrderRestored=true`；**恢复失败** ⇒ `ok=false` + `reason:"vbp_order_restore_failed"`（说明文本在 `okReason`）；`vbpRevisionBumped` 是**实测**值（前后 `RevisionVer` 从文件读；VB6 不 bump 就是 false + `vbpOrderNote` 说明）；`vbpOrderLinesCapped` 是行序快照的条数上限 = **400 行**（`VBP_ORDER_MAX_LINES`；`vbpOrderLinesBefore[]` / `vbpOrderLinesAfter[]` 各最多 400 行，超则截断），`vbpOrderChanges[]` 最多 **50** 条（`VBP_ORDER_MAX_CHANGES`），`deletedTextLimit` = **4000** 字符（`VBP_DELETED_MAX_CHARS`）。`apply_edits` 的块删除（`op:"D"`）现在报 `deletedLines` / `deletedText`（截断口径：`deletedTextTruncated` + `deletedTextLimit`）；`apply_edits` 与 `preview_edit` **都**新增两条**逐字同 reason** 的拒绝（`delete_with_text` / `would_empty_component`），但 **`preview_edit` 不报 `deleted*` 键** —— 它的删除体现在 diff 中，读 `preview_edit.deletedLines` 只会得到 `undefined`。

**挂死可观测（作者编号 115，字段已锁定）**

| 字段 | 工具 | 语义 |
|---|---|---|
| `timedOut` | `make_project` / `build` | 子进程**到预算被终止**。`build` 超时时 `ok:false` + `verdict:"timeout"`（不再抛「build process failed to run」） |
| `budgetEnforced` | `make_project` / `build` | 预算**真的执行了终止**；cache hit / 未起编译器 = `false` |
| `budgetMs` | `make_project` / `build` | 实际预算；未起编译器 = `0` |
| `budgetSource` | `make_project` | `timeout_ms` / `default` / `clamped_min` / `clamped_max` / `not_applicable` |
| `stopReason` / `stopReasonPriority` | `make_project` | `stopReason` **四值写死**：`completed` / `timeout` / `refused` / `host_modal_blocking`；两者同真时 `timedOut` 是**终止原因**、`host_modal_blocking` 描述**宿主状态**（顺序固定） |
| `budgetNote` / `timeoutNote` | `make_project` / `build` | **说明字符串，不参与判定**（人读；机器判定只看 `timedOut` / `budgetEnforced` / `stopReason` 等布尔与枚举字段） |
| `host_modal_blocking` + `host_modal{blocking, stage, count, risk, riskReason, riskClass, ownerDisabled, hwnd, title, controls[], dialogs[], observedDuringCall, source}` | `make_project` | 宿主模态框挡住编译；与 `timedOut` **独立**，可同时为真。**`observedDuringCall`**(bool) + **`source`**(string) 表达「调用期间出现过模态框、收尾时已被守卫取消」（这时活体已看不到框，只看 `blocking` 会漏报） |
| `isolatedMake{attempted, pid, exitCode, killed, killMethod, elapsedMs, leftover, budgetMs}` | `make_project` | 隔离进程证据：被杀 + `leftover` 证明事后无泄漏；**`exitCode=-1` 表示子进程被 kill / 未真正运行** |
| `killMethod` / `childPid` / `compilerElapsedMs` / `childLeftover` | `build` | 同一机制的收窄字段集 |

**预算的默认值与上下界（锁定值）**

| 工具 | 默认 | min | max | `budgetSource` 取值 |
|---|---|---|---|---|
| `build`（隔离 `/make`） | **180000 ms**（既有默认不变） | **1000** | **600000** | `timeout_ms` / `default` / `clamped_min` / `clamped_max` |
| `make_project`（`diagnose` 隔离副本） | **120000 ms** | **5000** | **600000** | 同上 + `not_applicable` |

**三条语义（按此读回执）**

1. **`budgetMs` 报实际生效值（clamp 之后）**，不是请求值；clamp 触发时 `budgetSource` 写 `clamped_min` / `clamped_max`（**不叠加**原来源名，避免不可判读）。
2. **`not_applicable` 的三种情形**：in-IDE Make（无编译器子进程） / 预检拒绝 ⇒ `budgetMs=0`、`budgetSource:"not_applicable"`、`budgetEnforced=false`（两者均**活体实测**）；`make_project` 的 **in-IDE 路由结构性不受 `timeout_ms` 约束**（没有子进程）。★**命中缓存时回执也发布这组字段**（v1.16.6 起；`path:"cache"`、`cached:true` 可判）：cache hit **不起编译器子进程**，故按结构性不可用发布 `timedOut:false` / `budgetEnforced:false` / `budgetMs:0` / `budgetSource:"not_applicable"` / `childPid:0` / `childLeftover:false`（与预检拒绝同形）；**非 cache 与 timeout 路径逐字节不变**。补之前（≤ v1.16.5）该组整组缺席，旧读方以 `cached:true` 判断即可。
   ★另：`make_project` 的预算明细在 **`diagnosis` 子对象**里（`diagnosis{attempted, ok, reason, preflightError, budgetMs, budgetSource, isolatedMake{...}}`），**根回执没有 `diagnosisAttempted`**；复制失败时 `diagnosis.attempted=false` 且 `isolatedMake.attempted=false`。
3. **`build` 的兼容性保证**：非超时路径的 `ok` / `verdict` / `exitCode` / `log` / `firstError` **一字不变**，新增字段全 **additive**（`timedOut=false` / `budgetEnforced=false` / `budgetMs` / `budgetSource` / `killMethod:""` / `childPid` / `compilerElapsedMs` / `childLeftover=false`）；**只有** `timedOut=true` 时覆盖为 `ok:false` + `verdict:"timeout"` + `buildState:"failed"` + `timeoutNote`，**不再** raise「build process failed to run」。

`isolatedMake.killed` = `budgetEnforced`；`isolatedMake.leftover` = kill 后 tasklist 是否仍列出子进程（**`false` 才表示确认无残留**）。

**本版关闭该缺陷卡** —— 关闭口径**要实测、不是声明**：① 小预算 ⇒ 到点 `timedOut:true` / `budgetEnforced:true`，且子进程经 tasklist 确认**无残留**（★clamp 下限**按工具不同**：`build{timeout_ms:300}` ⇒ **1000**；`make_project{diagnose,timeout_ms:300}` ⇒ **5000**）；② 同一工程另跑一次正常预算 ⇒ 非超时路径的 `ok` / `verdict` 与基线**一致**。**两条都成立才算关闭。**
   ✅ **2026-10-20 已按此关闭（该缺陷卡）**：独立验证者在实机一手实测 —— `build` 1000ms 与 `make_project diagnose` 5000ms 均 `timedOut=true/budgetEnforced=true/killMethod=WshScriptExec.Terminate/childLeftover=false` + 独立 `tasklist` 复查无残留；同工程 180000ms 重跑 `ok=true/verdict=artifact/cached=false/artifactChanged=true`；小工程默认预算 `180000/default`。证据 `<evidence-dir>\v6_live\115_arms.txt`、`115_big.txt`。**`ide_state` 零工具面变化**（无参数 / 描述 / schema 改动，`inflight{}` 字段集不变）：只改行为 —— 隔离编译的等待循环泵消息（`Util.PumpNow`），使 `depth>0` 的**只读探测在长调用期间被派发**；`serviceable` 语义按实测 = 该只读调用**确实被服务**。

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











### BREAK 态的 control_execution：**有界交付**（新增稳定 reason）

- BREAK 态是唯一允许走 CommandBars 的例外，但它必须**有界**：交付期限常量 `CB_CMD_DEADLINE_MS = 2000`，交付后的状态等待用同一个预算 `CB_CMD_STATE_WAIT_MS = 2000`（此前是 12,000 ms，实测把 single-flight 门占了 11.7 s，后续调用全 busy）。
- 交付前做**有界 liveness 探活**（2000 ms 上限）；不响应则不执行、直接结构化失败。
- 交付未生效（状态在预算内没变化）⇒ `ok:false, reason:"commandbars_command_not_delivered"` + `commandBarsTouched:true` + `attemptedSteps[]` + `boundedMs` + `howToProceed`。
- 存在**模态运行时错误框**时不再走"CommandBars 驱动解决"（那会排队，实测 12 s）：直接返回结构化 `commandbars_command_not_delivered`（`blockedBy:"runtime_error_dialog"`），先让调用方用 `last_error{dismiss:true}` / `send_window_message` 关框再重试。
- 交付期间 `ArmGateRelease "control_execution", CB_CMD_DEADLINE_MS` + `DisarmGateRelease`：即使 CommandBars.Execute 本身挂住，单飞门也会在预算内被释放。
### 运行期间的 CommandBars：**做不到**（实测 APPCRASH，架构限制）

- **事实**：VB6 在程序 RUN 期间使 CommandBars（mscomctl.ocx）进入 suspended 状态，**任何触碰都可能让 VB6.EXE 硬崩**。WER 1001 APPCRASH：`VB6.EXE c0000005 @000a250f`、`VBA6.dll c0000005 @000b90a9`（与 FIX-16 同一偏移）、`mscomctl.ocx c000041d/c0000094`。
- **拒绝**：需要 CommandBars 的工具在 RUN 态直接返回 `ok:false, reason:"commandbars_unavailable_while_running"`，带 `commandBarsTouched:false` + `howToProceed`（从 IDE 界面停止程序或等它结束）。BREAK 态下非 debug-control 的 CommandBars 工具返回 `reason:"commandbars_unavailable_while_in_break"`。
- **只读安全路径**：`ide_state` 在 RUN 态只用 VBA6 `EbMode` 回答（`stateSource:"ebMode"`，`commandBarsSuspended:true`），**不触碰 CommandBars**；只有 DESIGN 态才用 CommandBars 增强信息。
- **唯一例外**：BREAK 态的 `control_execution`（continue/step/end）仍走 CommandBars（否则无法恢复；69fa3f7f 实测该路径可用）。
### `run_project{observe:true}` 在 RUN 态的 STA 排队语义（实测，ADDITIVE）

**dispatch 字段口径（第二轮修正）**：`programStateAtEntry` = **工具真正开始执行那一刻** 的程序状态（不是客户端发请求时）；`queuedBeforeToolEntryMs` 在 STA 下**结构性不可测**（add-in 只有在被调度之后才开始跑），因此固定为 `null` + `queuedBeforeToolEntryUnavailableReason:"STA: the add-in cannot observe the interval before its own dispatch"` —— **不许报 0 冒充测量**；`dispatchedLate` 是可判定定义：**入口状态 ≠ "run"** 即 true（对 `observe:true`），message 前置说明"本次调用被调度时程序已不在运行；无法证明期间有实时观察，marker 命中的是调度时刻的状态"。


- **事实**：VB6 IDE 是单线程 STA。程序 RUN 期间，消息泵**不派发新的工具调用**；`run_project{observe:true}` 只有在程序**让出控制权**（自身调用 `DoEvents`）、**命中断点**、或**退出**之后才被服务。活体数据：client 20,533 ms，而 add-in 内部 `addInObserveMs=0` / `toolEntryToObserveMs=78` / `observeEntryTickMs==clockAtLoopEntry==clockAtExit` / `iterations=1`。
- **因此**：`observe` 在 RUN 态**不是实时观察**，语义 = "等程序让出控制权/结束"；`state="timeout"` 在 RUN 态**结构性不可达**（除非程序自己 DoEvents）。B3 那次"7 s 命中 marker"实际是**退出那一刻**才返回，不是实时观察。
- **正确做法**：用断点 + `run_until_break`（BREAK 态可被服务），或让程序自身 DoEvents。
- **回执新增** `diag.dispatch{toolEntryToObserveMs, addInObserveMs, queuedBeforeToolEntryMs, dispatchedLate, programStateAtEntry}`（取不到时写 `null` + `queuedBeforeToolEntryUnavailableReason`）。当 `programStateAtEntry="exited"` 时 `dispatchedLate:true`，`state` 仍保持 `exited`，但 message 明说"本次观察在程序结束后才被调度，期间没有实时观察，marker 命中的是退出那一刻的状态"。
### `get_locals` 在 BREAK 的默认拒绝 + opt-in（稳定 reason，不 bump 契约）

- **BREAK 且未给 `allowUnsafeInBreak:true`** ⇒ 立即返回 `ok:false + reason:"break_eval_requires_opt_in"` + `failureKind:"toolError"` + `refusedBeforeExecution:true` + 回显 `allowUnsafeInBreak:false` + `howToWorkAround[]`（① 传 `allowUnsafeInBreak:true`（知情同意；可能 wedge，恢复走 hub `ide_kill_and_relaunch`）② 先 `control_execution{mode:"continue"}` 离开 BREAK 再取）+ `howToRecover`；**不执行任何取值**，且 `localsUnavailable:false`（它不是"读不到值"）。
- **带 `allowUnsafeInBreak:true`** ⇒ 走取值路径，回执 `optInAccepted:true` + `howToRecover`；若仍取不到值，用**另一套**字段区分：`localsUnavailable:true + localsUnavailableReason:"values_unreadable: …"`（**没有** `reason:break_eval_requires_opt_in`）。
- **两种 ok:false 必须可区分**：`break_eval_requires_opt_in`（未同意）vs `values_unreadable`（同意但读不到）——活体 FAIL 正是把两者混成了一个 ok:false。
### ADDITIVE reason 值：`immediate_eval` / `immediate_run` 的 sink 失败分类（不 bump 契约）

- `engine_refused_line` 只表示**表达式/语句被引擎拒绝**（`ret=40036` 一类）。
- 新增 `eval_sink_error` / `run_sink_error`：**sink 行自身**在运行期失败（文件号/路径/权限等 VB 文件错误码 52/53/54/55/59/61/62/63/64/67/68/70/71/75/76），回执同时带 `errNumber`/`errDescription`/`lineText`/`sinkFile`/`ret`/`retHex`，且 `sinkErrorKind` 明说"表达式没有被拒绝"。
- phase-2（取值阶段）失败时另带 `phase2Ret`/`phase2ErrNumber`/`phase2Reason`。
- 背景（实测）：执行行曾以 `Close #511:` 开头，而该文件号此时尚未打开 ⇒ VB6 在该语句处运行期报错 ⇒ 整行被 `EbExecuteLine` 拒绝 ⇒ 被误报 `engine_refused_line`。现在 sink 行用 `FreeFile` 自取文件号，且**不再前导 Close**；旧注释"a leading Close #511 is harmless"已被实测推翻，源码里保留为**历史误记**说明。
### ADDITIVE 回执字段（不 bump 契约）：`run_project(observe)` 的 `diag{}` 与 `clockSuspect`

观察回执新增一个**只读诊断块**，用于让"这一轮到底谁在耗墙钟、时钟是否可信"可判定，而不是靠猜：

```text
diag = {
  blockingProc,            // 实测耗墙钟最多的那个被测过程（raw clock 跨度最大者）
  blockingProcMs,          // 它的累计实测跨度（ms）
  clockImpl,               // 选中的时钟：codetools_gettickcount | util_tickms | vb_timer | calendar_now | none_available
  clockUsable,             // 该时钟是否通过了入口的 Sleep 自检（delta > 0）
  clockStart, clockAtLoopEntry, clockAtExit, tickDelta,   // 真实读数（不可得时 null + 相应 unavailableReason）
  iterations,              // 等待循环体内真实执行次数
  errNumber, errDescription, errProc,   // 被吞掉的最后一次 Err（每个被测调用后立即采样）
  waitSpans = { <proc> = {proc, ms, calls}, ... },   // 至少覆盖 ReCaptureRuntimeError / GetIdeState / WinOps.HasVisibleHostDialog / DbOfWait / DbBriefWait / DoEvents 中被测到的几项
  clockSelfTest = { sleepMs:50, before{}, after{}, delta{}, rule },  // 入口用 Sleep 50 前后各读一次，零 delta 即"本构建时钟不前进"
  rule
}
clockSuspect                 // true = 本构建没有任何时钟跨 Sleep 50 前进；message 会明说"预算不可信"
```

约束：`iterations` 与 `clockStart/clockAtLoopEntry/clockAtExit/tickDelta` 都是**真实读数/计数**

**诊断轮 2 修正（同一 ADDITIVE 块，字段名以这里为准）**
**诊断轮 3 修正（同一 ADDITIVE 块；根因是 error 91）**：
- `diag.clockSelfTest` 的 `start/end/delta` 必须走 `RpSetInto`（Set 失败写 `<key>Error*` 而不是丢字段）；**不得**出现 `cst("start") = Nothing` 这种漏 `Set` 的对象赋值（VB6 抛 error 91，整段自检丢 start）。
- 新增墙钟字段：`diag.observeEntryTickMs`、`diag.addInObserveMs`（observe 内部实测墙钟）、`diag.toolEntryToObserveMs` / `diag.toolTotalMsAtObserveReturn`（用 `CallTool` 的 `m_callStartAt`，可判"客户端的 20 s 是否根本没进入 observe"）。
- `waitSpans` 现在也覆盖**循环外的等待**：`ReFindRuntimeErrorDialog`、`ReRuntimeErrorText`、`ReFindDebugButton`、`ReFindEndButtonInDialog`、`DebugClickSend`（同步 BM_CLICK）、`RePollBreakMs`（有界 EbMode 轮询）、`ReFindRuntimeErrorDialogAfterClick`；由 `m_diagSpans/m_diagClock + DgNow/DgSpan` 采集。
- `diag.blockingProc` **只报耗时最长的 span**；所有 span 都是 0 ms 时它必须是 `null` + `blockingProcUnavailableReason`（"报第一个 0 ms span 是撒谎"）。





**immediate_eval 第四轮（定案：即时上下文属于"当前打开的项目"）**：活体 `line="Util.ImSinkVal = 1+1"` ⇒ `ret=40036` —— 因为即时窗口在 **打开的项目** 里执行，而该项目没有 add-in 的 `Util` 模块，任何 `Util.*` 标识符都是编译错误。**执行行只能使用字面量文件号 `#511` + 标准语句（隐式变量/`FreeFile` 在 Option Explicit 的打开工程里是编译错，实测 ret=40036/errNumber=0）**。因此 eval 走 `immediate_run` 已验证可用的通路：执行 `Debug.Print <expr>`（body 由 `ImRewriteDebugPrint` 改写进 add-in 的 Debug.Print sink），add-in 读 sink 文本作为 `value`，`valueTransport="debug_print"`；多行值按 sink 原样保留（`valueLines/valueBytes` 照旧）；`TypeName` 在这条路上不可得，回执给 `typeNameUnavailableReason`。对象表达式仍先被 `object_expression_unsupported` 拦下；空输出是 `value_not_parsed`（不猜）。
**immediate_eval 第三轮（实测：`EbExecuteLine` 不接受任何一个 sink 行）**：`Open/Print #/Close #` 组成的执行行在活体上**一律被拒**（即使只用已声明变量）。执行行现在**只有一条赋值**：`Util.ImSinkVal = <expr>`；`TypeName/Len/LenB/行数/转义` 全部由 **add-in 代码**从该 scratch Variant 直接读取（`valueTransport="direct_variant"`），**没有文件 sink、没有 MCPLEN/MCPV 标记**。表达式本身非法时 `EbExecuteLine` 的非零返回才是真正的 `engine_refused_line`。`immediate_run` 仍用临时文件 sink（它要搬运 `Debug.Print` 的输出来）。
**即时上下文允许的语句形态（`immediate_eval`/`immediate_run` 口径）**：执行行里**只允许**出现工程已声明的标识符（本项目固定为 `Util.ImSinkFn` / `Util.ImSinkVal`）+ 字面量 + 标准函数；**不得**引入隐式新变量（工程开了 `Option Explicit`，隐式变量是**编译期**错误，整行 `ret<>0` 且 `errNumber=0`，看起来像"引擎拒绝表达式"）。行内不使用 `On Error Resume Next`、不嵌 CRLF、不以 `Close #n` 开头（文件号未打开时该语句会运行期报错并让整行被拒）。
：`diag` 必须**先落盘再判断**——`clockSelfTest` 在自检**运行之前**就写入 `diag`（自检把 start/end/delta 与失败的 `clockSelfTestError{errNumber,errDescription,errProc}` 写进同一个对象），`clocksAtStart` / `clockEndToStart` / `waitSpans` / `preLoopSpans` 同理；**任何被文案引用到的 `diag.X` 键必须真实存在**（否则视为缺陷）。新增：
- `clockSelfTest{start,end,delta,sleepRequestedMs,nowDeltaMs,sleepApiErr,sleepIneffective,timePassProven,busyWait{attempts,clockImpl,deltaMs,unavailableReason}}` —— 区分"**睡眠本身没生效**"与"**时间过去了但钟不动**"；
- `clockVerdict` = `clock_advancing` | `time_passed_no_clock_advanced` | `unknown_time_pass_not_proven`；`clockSuspect` **只有在前两者（证明确实测到时间流逝）时才允许为 true**，`unknown_time_pass_not_proven` 时不得声称"时钟死了"；
- `budgetClockImpl` / `budgetClockTrusted` / `spanClockImpl`（预算与跨度各用哪个钟、是否被自检证明）；
- `clockEndToStart`（整次调用前后的四钟真实差值，**不需要 Sleep** 即可判读）；
- `preLoopSpans[{proc,ms,calls}]` + `preLoopMs` + `preLoopTopProc`（解释"iterations=1 但墙钟 20 s"：时间是否花在**进循环之前**）；
- 每条被测调用后的 `errNumber/errDescription/errProc`，以及 `blockingProc/blockingProcMs/blockingProcUnavailableReason`。
收据还新增 `preObserveGetTickCountMs` / `preObserveClockImpl` / `preObserveGetIdeStateBlocked`（observe 入口之前那次 `GetIdeState` 的耗时，用 CodeTools 自己的 `GetTickCount()` 量）。
；拿不到就写 `null` + `unavailableReason`，**不写 0 冒充**。工具面不变（无新参数、无新工具、不 bump 契约、不改冻结面）。

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
PRODUCT vb6-mcp <semver>  CONTRACT v1.17.0  TOOLS 120  BUILT <sha256 前 8 位>
BREAKING (read before calling):
- v1.17.0 MINOR ADD-IN face change: +1 READ-ONLY tool `trace_query` (add-in 119 -> 120, hub 11 -> 11, DSH-facing 130 -> 131) and two OPTIONAL `inject_trace` parameters `condition`/`recordCondition` (logpoint; the 5-line block is byte-for-byte unchanged when `condition` is omitted); `remove_trace` only gains a description sentence (S-1). The frozen face was RE-FROZEN (221,252 B / `4A201C0C…`). Existing callers that never pass the new parameters are unaffected.
- v1.16.0 ADD-IN face, purely ADDITIVE (no tool added or removed; add-in 120, hub 11, DSH-facing 131): breakpoints{action:"probe"} is the BEHAVIOURAL breakpoint verifier and it is TWO-PHASE - phase="start" (default) only STARTS the program (QUEUED in the STA IDE, so the run begins AFTER the call returns) and answers probePhase=start + probeResult=started_pending + probeId + startedPending:true + startQueued:true + howToContinue; phase="observe" (the SAME probeId, single-use, TTL 600000 ms) waits up to timeout_ms for the stop, reads position{component,proc,paneLine} + callstack, ENDs the program and reports endPending:true until stateAfter=design is PROVEN by read-back (unknown/expired probeId answers reason=probe_not_found). probeResult is a CLOSED set: started_pending | stopped_at_target | stopped_elsewhere | did_not_stop | timeout | did_not_stop_unobserved | line_not_breakable | not_attempted | run_failed. The probe REALLY runs the program and never saves it (no pre-run flush; measured 6.2 s / 12.1 s). control_execution{mode:"end"|"stop"} in RUN state no longer false-refuses (measured End path; endCommandDelivered/endPending/stateAfterVerifiedBy/stateAfterReadBack/programEnded/deliveryVsStateNote/endSettleMs; programEnded:true ONLY after a design read-back). verify and verify_scope gain OPTIONAL requireCompiler (default false = receipt unchanged): true demands a REAL isolated VB6.EXE /make verdict (authority=compiler + compiler{path,logPath,errorCount,artifactChanged}) and REFUSES with reason=compiler_not_configured + compilerProbe{resolved,path,why} instead of ever substituting lint. apply_edits and preview_edit accept a MULTI-COMPONENT batch: edits[].component (tool-level name becomes OPTIONAL; neither present = reason=component_required), expected_shas[] guard, ops applied in edits[] order with each anchor resolved on its component UPDATED text, all-or-nothing rollback (rolledBack/restoreVerified/restoredComponents[]), max_ops default 200 / cap 1000 and at most 40 components (batch_too_large + limits), success adds appliedOps/totalOps/components[]/batchElapsedMs, and preview_edit groups hunks per component with writes:false
- v1.15.0 ADD-IN face parameter change (ADDITIVE for callers that do not pass them): run_project.observe_max_wait_ms (OPTIONAL integer ms, default timeout_ms, clamped 1000..120000) and run_project.done_marker (OPTIONAL object {path,contains}; path is ON THE MACHINE THAT RUNS THE IDE, UNC allowed; the first 65536 bytes are read as BYTES and nothing is written; contains is CASE-SENSITIVE, a pure-ASCII value is byte-exact and a non-ASCII value is tried as UTF-8 and as system ANSI). observe RETURNS with state = done-by-marker | exited | break | error | timeout (a CLOSED five-value set: done-by-marker IMPLIES marker.matched:true, exited means the program ended normally, and a requested-but-unmatched marker reports marker.unmet:true - the run finished without the requested evidence), observedState, elapsedMs, marker{...} and pipeReleasedByReturn:true, so a timeout no longer holds the single-flight gate (a timeout receipt is ok:true + timeout:true + retrySafe:true + suggestedAction=observe_again with the measured elapsedMs - NOT a failure - and marker.unmetReason names the budget when a requested marker never appeared; the budget is enforced INSIDE the wait loop, not only in a loop condition); observe in BREAK is REFUSED with reason="observe_in_break_not_supported"; a #32770 dialog is CLASSIFIED (only a runtime-error box is program_error; an IDE Find/Replace dialog answers reason="observe_blocked_by_ide_dialog" with dialogPrecheck.isProgramError:false). get_locals.allowUnsafeInBreak (OPTIONAL boolean, default false): BREAK-state get_locals is REFUSED with reason="break_eval_requires_opt_in" unless true - the opt-in is INFORMED CONSENT and does NOT guarantee the statement will not hang; recovery is the hub-side ide_kill_and_relaunch, not taskkill; failure receipts carry gateReleased + howToRecover. The FROZEN FACE lag was closed at v1.16.0: docs/tools_schema.json and tools/schema_baseline.json now carry these declarations (toolCount stays 120; hub 11; DSH-facing 131).
- v1.14.0 hub tool face (ADDITIVE for the add-in face): ide_restart gains OPTIONAL wait_ms (bound how long ONE call blocks) and opId (continue/poll an operation a previous call returned). A restart that outlives wait_ms answers pending:true + state="in_progress" + opId + currentStep + lastKnownState instead of looking like a failure, and ide_restart{opId} answers from the in-process operation registry and starts NOTHING; a second restart while one is in flight is refused with reason="restart_in_progress". Every ide_restart receipt carries opId/started/state/finished/elapsedMs/currentStep/lastKnownState/opRule. content[0].text is ONE JSON document: array/scalar bodies are {singleDocument:true, data, receipt} (empty = data:[]), a non-JSON body is a structured content_not_json document, object bodies stay FLAT; VB6MCP_LEGACY_TEXTBODY=1 restores the old two-part form for one version. Transport failures add peerState="peer_dead" (+peerDead/lostContext/howToRecover) when the peer process is provably gone. add-in face unchanged (120 tools; hub 11; DSH-facing 131)
- v1.13.0 +1 HUB tool: ide_kill_and_relaunch{pid, project?, ready_wait_ms?} (hub-side, add-in independent: stops ONLY a hub-owned instance, clears its registration, relaunches on the same .vbp and polls for readiness) and an OPTIONAL ide_start.ready_wait_ms. Receipts gain ready/readyWaitMs/readyChecks{registryVisible,transportUp,addInAnswers,addInAnswersBasis,pidAlive,projectMatches}/readyProbeRounds/pollingHint (ide_start; readiness is measured on the STARTED pid only - its own registration, its transport, mcp_info - never on the discovered set; addInAnswersBasis is exactly one of no_route|direct_rpc|middle_service_forward, and no_route means no channel existed, so addInAnswers is null, never false), offline/offlineRoute/targetResolved (build/make_project/project_meta with no live IDE), no_live_ide + onlinePids[]/targetDiagnostic/howToProceed (target refusals; the old reason no_ide_running is preserved as reasonLegacy). ADDITIVE for the ADD-IN face: no add-in tool or parameter changed (toolCount 120; hub 10 -> 11; DSH-facing 130 -> 131)
- v1.12.2 +1 OPTIONAL parameter: control_execution.allowDiskMutation / run_project.allowDiskMutation (boolean, default false). Starting from DIRTY IDE memory is now REFUSED with reason=dirty_memory_would_be_saved unless true; receipts add diskWrites/touchedFiles/diskSnapshot/saveOnRun/ideAutoSaveHint (ADDITIVE: existing fields unchanged; saveOnRun:false is NOT a promise that the disk did not change)
- v1.11.0 the 5 lessons tools (list_lessons / add_lesson / delete_lesson / export_lessons / import_lessons_md) MOVED from the add-in face to the hub face: names/descriptions/inputSchema are unchanged and they now also work with NO IDE online; the add-in face is 110 tools, the hub face 5 -> 10, DSH-facing total stays 120
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
ROUTING (use the tool; never hand-edit VB project files):
- .vbp version + build metadata (MajorVer / MinorVer / RevisionVer / AutoIncrementVer / CompanyName / FileDescription / ProductName / Title / ExeName32 / CondComp) -> project_meta (text edit, keeps a .bak; close the project first)
- project / IDE options (description / help file / startup object / startup project / output path / read-only) -> set_project_options
- UAC level + DPI awareness -> app_manifest (writes <exe>.manifest; do not hand-write one)
- components: add / remove / import / rename / reload / export -> add_component / remove_component / import_component / rename_component / reload_component / export_component
- COM references -> add_reference / remove_reference / check_references
- forms and controls: list / add / remove / design-time properties -> form_designer / add_control / remove_control / set_control_property / list_component_parts
- lint before writing / static checks / formatting -> validate_code / check_component / format_code / verify / verify_scope
- component text edits (surgical or bulk) -> edit_by_anchor / apply_edits / replace_code / replace_in_files; preview_edit first; snapshot=true or undo_edit to roll back
- project-wide checkpoint / rollback -> project_snapshot (whole project: .vbp + component files)
- build verdict -> build (offline /make) or make_project (in-IDE): judge by artifactChanged + verdict; compiled:true alone is NOT built
- ENCODING: VB project files (.vbp / .frm / .cls / .bas / .ctl / .pag / .dsr / .res / .inc) must stay system ANSI (GBK) WITHOUT a BOM - never hand-edit or re-encode them (a UTF-8 BOM makes VB6 refuse the project: "invalid keyword"); the tools above already write ANSI for those extensions.
- v1.17.0 current (120 add-in tools): the bullets above are historical - call mcp_info once after initialize for the authoritative breaking[] / migrations[] / changelog[]
CAPABILITY INDEX (question -> tool; read this before improvising):
- impact / who calls whom -> dependency_graph  (component edges + PROCEDURE-level procNodes/procEdges/procCount + cycles + purity)
- who references a symbol -> rename_symbol  (dry_run=true by default; skips strings and comments)
- public API broken? -> check_public_api / export_public_api ; missing files or OCX parts -> check_references / list_component_parts
- runtime window tree / control text / modal dialog -> find_gui / ui_text / read_active_dialogs
- click a button or send keys to a window of THIS VB6 process -> send_window_message
- offline single-module test -> test_host ; stub a procedure or trace calls -> fault_inject / inject_trace
- NOT available: screenshots; driving a standalone built EXE (message tools are in-process only)
FILE ENCODING: .vbp/.frm/.cls/.bas/.ctl are ANSI WITHOUT BOM (write_file / write_source already do this); never write them as UTF-8-with-BOM, or VB6 reads a mangled first keyword and the project stops loading; use project_meta for .vbp version metadata (keeps a .bak).
```


**同步规则（lane 门禁 G13）**：hub 快照的 contractVersion 取「## 版本记录」最新行，而 instructions 首行与 BREAKING 列表取「## v1.3.0」的 text 块 —— **契约升级时必须两处同步**（横幅字面 CONTRACT <ver> 改到与最新版本行一致），否则快照自相矛盾；G13 断言 contractVersion == instructions 首行 CONTRACT <ver> == add-in 侧契约值。

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

#### 13.1 方向裁定（一项既有缺陷；为什么是「声明必填」而不是「接受缺省」）

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
这正是本项目列为缺陷的**误导性成功**（同族的相邻缺陷）；③ 该做法还要**删掉** 刚落地的
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

#### 16.4 缺陷处置（重载目标 / 重载回执）

* **重载回执的假名 ⇒ 已修**：此前 `reloaded` 读的是 `Reload` **之后**的 `comp.Name`，而目标组件**代码窗格从未打开**时 VB6 返回**默认名**（标准模块 = `Module1`）。改为**重载前**读取。
  * 同口径对照（同夹具 `P119R`、同步步骤）：`J1-pane-never-opened` 由 **`Module1` → `mNew2`**；H/I 臂由 `Module1` → `mNew`。
* **重载目标不可判定（改错文件却报成功）⇒ 行为未改、回执改为可判定**：`reload` 仍只重载绑定文件（这是 `VBComponent.Reload` 的语义，**不是缺陷**），但现在回执给出 `file`，调用方可直接与「我改的路径」比对。
  * H 臂改后仍 3/3 `paneHas778=false`（**行为如实保留**），但 3/3 回执带 `file=…\Module1_2.bas` ⇒ 调用方**能判定**重载覆盖不到自己改的文件。
* ★**未做**：`reload_component` **不做窗格前后校验**（不加 `paneShaBefore/After`）；该缺陷的「建议」仍只是建议。

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

### 18. `find_gui` 集合稳定性修复 —— ★**零 schema 变更、零重冻**

> 本节**不改 `inputSchema`、不改工具面、不改任何描述** ⇒ 工具仍 **111**、契约仍 **v1.7.0**、冻结件 **`C9C96324…`（108,019 B）跑前 == 跑后**、hub 快照 **`6EA40C0C…`（111,116 B）未变**、`hub_alignment=0` ⇒ ★**免重冻**（S 预先授权的第三次重冻**未使用**）。
> ★理由：`find_gui` 的**现有**描述本来就是「*Enumerate the windows of THIS process (top-level windows plus all their children)…*」—— **那正是修好之后的行为**，改前的实现做不到自己声明的事。

#### 18.1 缺陷

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

## 返回信封与坐标系（v1.11.7 明确化，实现未变）

· **返回信封**：每次 `tools/call` 的 `result.content[0].text` 是 **一整串 JSON 字符串**（工具回执），不是已解析的对象；客户端必须对该字符串再解码一次。此前该事实只体现在实现里（`McpServer.cls` 的 envelope 构造），现作为契约写明。
· **`get_component_code{proc|symbol}` 的行号语义**：`startLine`/`endLine` 是 **IDE code-pane 行**（与所有写类工具同一坐标空间，**不随 `coords` 变化**），`startLineDisk`/`endLineDisk` 给出同一范围的**文件行**；被编号的文本本身仍跟随 `coords`（`coords:"disk"` 时编号为文件行），因此范围字段与编号**不是同一空间**。★**v1.16.6 起** `startLineDisk`/`endLineDisk` 取**累积** pane→disk 映射（header 行 + 该行之上的隐藏内联 `Attribute` 行），不再用标量 `headerOffset`；回执新增 `startLineDiskExact`/`endLineDiskExact`（true = 累积映射命中；false = 标量回退，**不可信**）与 `startLineDiskBasis` 说明。活体实测（`Class1.cls`：9 行 header + 1 行隐藏 `Attribute Foo.VB_UserMemId`；`proc:"Baz"`、`coords:"disk"`）：修复前 `startLineDisk:19`/`endLineDisk:22` 而同一回执打印的首/末行号为 20/23；修复后 `startLineDisk:20`/`endLineDisk:23`/`*Exact:true`，与编号一致。
· **`check_public_api{membersOnly:true}` 的作用与局限**：作用 = 在一次调用里同时拿到**成员表**（kind/name/signature）与 `violations[]/added[]/counts`；局限 = 成员表是**文本级轻解析**（参数名按设计丢弃、无类型解析，不是编译器符号表），且**比对永远执行**（`membersOnly` 只是输出大小，不能用来跳过比对）。

## 2026-10-23 · Build 1（a+b+c+d）

- `vba6_probe{engine_items}` 回执**语义收窄（非破坏）**：`items` 与 `available` 均置 **null**，声明「不支持项号」；`available` 是 `items` 的**历史别名**（同一含义，两个键同时保留只为显式说明）。
- 新增（附加）：`bitmap_len`(72) / `bitmap_hex`(144 hex) / `byte_offsets[]`(5 段，`confidence=inferred`) / `confidence` / `note`；`hex` 与 `availableCount` 保留，其中 `availableCount` 的语义明确为**非零字节数**（不是项数）。
- `get_project_info` 回执**恒发** `noProject`（`true`/`false`），此前仅在无工程时出现。
- 描述变更（4 处，输入 schema 零变更）：`vba6_probe.engine_items` 参数、`add_project`（独占替换）、`breakpoints`（三条规则）、`get_project_info`（`noProject` 恒发）。
- 工具数不变：add-in 120（hub 11 / DSH 131 不涉及）。

### 冻结时发现的两项事实（必须记录，非本批功能）

1. **carried-in drift（继承漂移）**：冻结前三方对照证明 `tools/schema_baseline.json` 已**落后于源码**——
   `find_in_component` 与 `search_in_files` 的输入 schema 各有 2 个参数（`include_procedure`、`include_leading_comment`）
   已经存在于**基线二进制 srcAE0 的实况面**里，但从未进入冻结面（旧 `docs/tools_schema.json` 是更早的捕获）。
   本次 `--freeze` 一并把它们纳入冻结面；**这不是本批新增的功能**，属把契约面补成事实。工具数仍 120。
2. **`engine_items` 的位语义实测**：两态（design vs RUN）对照显示 `+0x07/+0x08` **在两种状态下都是 0**，
   规格里「受 EbMode==1 门控」的推断**未获确认**；实测发生变化的字节是 `+0x09/+0x14/+0x15/+0x1D`。
   因此 `byte_offsets[]` 的条目统一标注 `INFERRED_STATIC`，并新增回执字段 `observedTwoStateChanged:[9,20,21,29]`
   与 `twoStateProbe`（明确写出「gating claim NOT confirmed」）。

## 2026-10-23 · Build 2（e）：`open_project` 同文件多种拼写短路（**无 untitled 守卫**）

- **行为变更（仅 `open_project`）**：目标 `.vbp` 已经打开时，如果调用方给出的路径只是**另一种拼写**
  （8.3 短名、正斜杠、大小写差异、目录别名），现在也会被识别为「同一个文件」，走**既有**的
  already-open 分支：返回 `alreadyOpen:true` ＋ `diskVsMemory`，**不**关闭再重开工程。
  原来的精确比较 `StrComp(VBProject.FileName, full, vbTextCompare)` **一个字符都没改**，只在它没命中后
  追加一轮规范化比较（文本归一化 → `GetAbsolutePathName` → 规范路径 `GetFinalPathNameByHandleW`；
  任何一步证不出来就保持 `False`，绝不猜）。
- **回执键集合未变**：短路复用 already-open 分支既有字段（`alreadyOpen` / `diskVsMemory` /
  `keptInMemory` / `reloadNotAttempted` …），**没有新增回执键**；**inputSchema 零变更**；工具数 120。
- **描述变更 1 处**：`open_project`。
- **不做 untitled 守卫**：本项目原先假设的「exclusive 分支先 `save_project("")`、对 untitled 工程会弹
  VB6 宿主 Save-As 模态框并阻塞 add-in」已被实测**反驳** —— 在 srcAE（Build 1 二进制）上
  A / A2 / B / B2 四个实验全部**正常返回**（33–69 ms，外部 `#32770` 前后均为 0），
  `remove_project{allowDirty:true}` 与 `open_project{exclusive:true}` 都不会卡住。
  真正的副作用是「未保存的 untitled 工程在关闭时被 VB6 **静默写进 IDE 当前目录**（并可覆盖同名组件文件）」，
  已单列新卡，实测口径见 `docs/KNOWN_LIMITATIONS_AND_FAQ.md` §62。
- 验收出证：`fix/TASK_CODETOOLS_SPLIT_20261022/handtest_build2/measure_untitled.json`（srcAE 七实验：
  A/B/B2/A2/S/H2-old/H6）与同目录 `handtest_build2_v2.json`（srcAG 手测矩阵：2b/H2/H2b×2/2c/H6/3a/3b，
  八步全部按预登记形状通过）。

## 2026-10-23 · 批次 3a：隐式写盘的**回执诚实性**（行为不变）

实测（`fix/TASK_CODETOOLS_SPLIT_20261022/handtest_build2/card_silent_write_measure.json`，M1–M4）：
`VBProject.Saved` 在 VBIDE 里**只读**（隔离 diag 构建编译失败于 cannot assign to a read-only property），
因此 add-in **无法**阻止宿主写盘。`VBProjects.Remove` 对**未保存**工程（untitled 与 named 皆然）会把工程
静默写进 IDE 当前目录（named 还可能写进它自己的工程目录），并可覆盖同名组件文件。故本批**不改行为**，
只让回执如实。

**`remove_project`（成功路径）**

- `discardedUnsavedChanges`：**语义更正** —— 旧值只表示「请求了丢弃」；新值表示「确实丢弃」。
  `dirtyFound And allowDirty` 且宿主写盘 ⇒ `false`，原因见 `discardedUnsavedChangesBasis`
  （取值：`host_persists_changes_on_remove` / `nothing_to_discard` / `no_host_write_observed`）。
- 新增：`discardedUnsavedChangesBasis`、`hostWritePrevented`、`hostWritePreventedWhy`、
  `hostWrites[]`、`hostWritesBasis`、`hostWritesTruncated`、`hostWritesScanned`。
- `hostWrites[i]` = `{path,sizeBefore,sizeAfter,existedBefore,overwritten}`：depth-1 快照、
  覆盖 IDE 当前目录 ＋（named 时）被删工程自己的目录、上限 2000、fail-open；
  **同尺寸且 mtime 相同的改写不可区分**（写在 `hostWritesBasis` 原文里）。
- `sideEffects.filesWritten[]` **并入** `hostWrites[].path`（本调用确实写了这些文件；provenance 看 `hostWrites`）。

**`open_project`（关闭旧工程路径）**

- `savedBeforeClose`：**语义更正** —— 旧值只表示「尝试过 preSave」；新值 = preSave 之后重读被关闭工程的
  `Saved`，全 True 才 true（untitled 的 no-op preSave ⇒ **false**）。
- 新增：`preCloseSave{attempted,result,saved,projectSavedAfter,basis}`、
  `hostWrites[]` / `hostWritesBasis` / `hostWritesTruncated`、
  `hostWritePrevented` / `hostWritePreventedWhy`。
- `hostWrites` 在**所有** `open_project` 回执里恒在；already-open / 非 exclusive 分支为空且
  `hostWritesBasis="not_applicable: no project was closed in this call"`；`preCloseSave` 只在关闭路径出现。
- 行为（`closedProjects` / `alreadyOpen` / 短路 / 不拒绝）**一律不变**；close 路径**绝不**写 `Saved`。

## 2026-10-23 · P1 #1：`ide_state` 的成功回执**恒定**给出 `mode`（纯加法）

**先纠正卡片前提（实测更正）**：`ide_state` 只有**一种**成功结构。`ToolIdeState` 在**每条**成功路径都写 `state`
（`design|run|break|runtime_error`）；所谓「commandBars 降级形状」并不是另一种结构，而是同一结构下 `stateSource` 的差异
（`ebMode` / `commandBars` / `ebMode+commandBars` / `runtime_error_dialog`）。`ideState` **从来不是 `ide_state` 的字段**
（它属于 `get_project_info` / `run_project` 等其它工具）；FAQ §61 的「形状 A 含 `ideState`」是误记。

**ADDITIVE 回执字段（3；工具数 120 不变、`inputSchema` 逐字节不变）**

- `mode`：`design | run | break | unknown`，**每条成功回执都有**（= 「该读哪个字段」的唯一答案）；`state` 保留原始四态，
  `runtime_error` 归一为 `unknown`（模态框是独立状态，不冒充三态）。
- `modeBasis`：`ebMode | commandBars | ebMode+commandBars | runtime_error_dialog | unrecognized_state`，**每条成功回执都有**
  （= `mode` 的判定来源；与 `stateSource` 同词表，但语义是「mode 的源」而不是「判定器的源」）。
- `modeConflict`：**只在矛盾时出现**（object）：`{state, ebMode, expectedFromEbMode}`。权威口径：`mode` 永远镜像 `state`，
  `modeConflict` 只报告「原始 EbMode 与已判状态不一致」，不把 `mode` 降为 `unknown`（避免双重语义）。

**不变**：`state` / `stateSource` / `ebMode` / `notRunning` / `running` / `commandBarsSuspended` / `autoRecover` / `pump` / `wheel` /
`cancel` / `release` / `inflight` / `lastError` 全部原样；**拒绝信封**（`ok:false` + `reason:ide_not_ready` + `vbErrorCode:438` +
`ideState:"design"`）**不加 `mode`**（拒绝不是成功形状，「先看 `ok`」的规则不变）。

**实现位置**：`CodeTools.cls` 的 `IdeModeBlock()`（唯一推导点）+ `ToolIdeState` 末尾唯一调用点（在 `autoRecover` 重算之后，
保证与最终 `state` 一致）；共享判定器 `GetIdeStateEx` **未改**。

**schema 重冻结**：仅 `ide_state.description` 变化 ⇒ `docs/tools_schema.json` == `tools/schema_baseline.json` =
`BD570A792925835EE206763E5036F7A2D2B7AF7E6A2E3E45A357DF84D68F54E6`（120 工具）；`bridge/hub_schema.generated.json` 重生成
（VERSION=v1.16.1、DECLARED=120、MISMATCH=False、FACEFP=`ff5cacae…`）；`hub_alignment_failures=0` / `live_tool_changes=0`。

**实测（srcAK，四形状）**：design（`stateSource=commandBars`）⇒ `mode=design`/`modeBasis=commandBars`；run（EbMode=1）⇒ `mode=run`/`ebMode`；
break（EbMode=2，注入断点命中）⇒ `mode=break`/`ebMode`；`runtime_error` **本轮不可靠复现、未测**（代码路径按上表映射为
`unknown` + `runtime_error_dialog`）。另：post-run 的 `ebMode+commandBars` 在 `ide_state` 回执里**不可观测**——该工具在返回前会
`autoRecover` 清一次悬挂窗口并**重读** ⇒ 降级源只存在于 `autoRecover` 之前的内部状态；`control_execution{end}` 的
`commandBarsSuspendedAfter=true` 是这条路径的外部证据。
## v1.16.6 诚实性/口径批（活体实测；字段级契约）

本节是 v1.16.6 的行为与字段口径正本。全部判据来自活体回执（新 DLL `D2148775…`，contract v1.16.6）与离线复现，证据在 `fix/TASK_V166_HONESTY_20261026/evidence/`。

### A. verify / verify_scope：没有证明对齐就不许报 MISMATCH（行号空间/对齐缺陷）

`verify.ideDisk` 是一个**封闭枚举**，v1.16.6 新增第六值 `NOT-ALIGNED`：`OK` / `MISMATCH` / `MISMATCH-FORMATTING` / `NOT-SYNCED` / `READ-FAIL` / **`NOT-ALIGNED`**。

判定顺序（`Snapshot.bas` 的 `VerifyTextVsDisk`，**行号空间/对齐缺陷 的实际写域是 Snapshot.bas，不是 Verify.bas**；Verify.bas 只负责把结果折进 `overall`/`failureLayers`）：

1. 调用方 `headerOffset` 处**逐字块匹配**成功 ⇒ `OK`，`ideDiskAligned:true`，basis 点名「caller headerOffset located ... VERBATIM」；
2. 失败 ⇒ 在整份盘上做**内容搜索**（`LocateBlock`，跳过 Attribute 行）找逐字块；找到 ⇒ `OK`，basis 点名「located VERBATIM at disk offset N by content search」；
3. 失败 ⇒ 先做**整文本**判等（只是去 paren-less `Call X()` 括号的正常化）：相等 ⇒ `MISMATCH-FORMATTING`，`ideDiskAligned:true`；
4. 失败 ⇒ **内容锚点对齐**：候选偏移来自 pane 的**首个非空行**（再试**末个非空行**），分数 = 该偏移下大小写不敏感相等的**非空**行数；只有**唯一最佳**且 `>= 2` 才算证明。证明成功 ⇒ 在该偏移上计算差异并回 `MISMATCH` / `NOT-SYNCED`，`ideDiskAligned:true` + `ideDiskAlignScore` + basis（含「matched N non-blank pane line(s) ... next best M」）；
5. 仍失败 ⇒ **`NOT-ALIGNED`**：`ideDiskAligned:false` + `ideDiskAlignedBasis`（三种可判原因：锚点分数不足 / 多个偏移并列 / 搜索到工作预算），**不产出行级 diff**（不回 `differing line(s)`），`semanticDiff = {available:false, reason}`，`overall = UNKNOWN`，`overallBasis = "ideDiskDiffUnproven"`。

**相关字段（additive）**：

| 字段 | 恒在？ | 语义 |
|---|---|---|
| `ideDiskAligned` | **仅当本次真的做了盘比较**（`diskPath` 非空）；`SKIPPED` 时整组缺席 | true = 对齐已证明（上列 1-4 任一）；false = 未证明（只有 `NOT-ALIGNED`） |
| `ideDiskAlignedBasis` | 同上（true/false 都带） | 可判的人读依据；**不是**日志 |
| `ideDiskAlignScore` | 仅 `ideDiskAligned:true` 时（且只有第 1/2/4 类给出） | 对齐证明命中的非空行数；`MISMATCH-FORMATTING` 时为整文本行数 |
| `failureLayers.ideDiskDiff.aligned` / `.alignedBasis` | 恒在（在 ideDiskDiff 对象内） | 与顶层同源；`failed:false` 的 `NOT-ALIGNED` **不是** proved diff |

**`verify_scope` 桶语义**：`NOT-ALIGNED` **属于 `unknowns[]`**（`reason:"ide_disk_diff_unproven"` + `alignedBasis` + `howToProceed`），**绝不进 `failures[]`**；不变量 `pass + fail + unknown == checked` 与 `countsConsistent` 不变。活体：3 组件工程中把 Module1 的盘上正文整体改写 ⇒ `checked:3 / pass:0 / fail:0 / unknown:3 / countsConsistent:true / failures:[]`，Module1 的 unknown 项 reason=`ide_disk_diff_unproven`；还原后同工程 ⇒ `pass:3 / fail:0 / unknown:0`。

**正/反样例（活体）**：

| 场景 | 盘上改动 | 判决 | 关键字段 |
|---|---|---|---|
| P 正例：只差首部行、正文相同 | 在 `.bas` 头部插入 1 行 `Attribute VB_Description` | `OK` | `ideDiskAligned:true`，`diskBodyOffsetUsed:2` |
| N1 真差异（且 header 也 +1） | 正文 `Public Sub Main()` -> `Public Sub MainRenamed()` | `MISMATCH` | `ideDiskAligned:true` + `ideDiskAlignScore:2`，`first diff at ide line 3 (disk line 5)`（偏移已证明，行号可信） |
| N2 无法对齐 | 盘上正文整体改写成另一个模块 | `NOT-ALIGNED` | `ideDiskAligned:false` + basis「matched only 1 non-blank ... minimum 2」，`overall:UNKNOWN`，`semanticDiff.available:false` |

### B. get_component_code：`startLineDisk`/`endLineDisk` 改为可证明的累积映射（行号空间/口径缺陷）

`get_component_code{proc|symbol}` 的 `startLine`/`endLine` **永远是 IDE code-pane 行**（写类工具的坐标空间），`coords` 只决定**被编号文本**的空间；`startLineDisk`/`endLineDisk` 是同一范围在**组件文件**里的行。v1.16.6 修复：旧实现用**标量** `headerOffset`（IDE 行 + header 行数）推导文件行，遇到模块体内**隐藏内联 `Attribute` 行**就比同回执打印的编号小 1（编号走累积映射，`coords.diskLine` 也走累积映射）。

新字段：`startLineDiskExact` / `endLineDiskExact`（boolean；true = 累积映射命中，false = 标量回退，**不可信**）、`startLineDiskBasis`（一句话说明）。实测（`Class1.cls` 9 行 header + 1 行隐藏 `Attribute Foo.VB_UserMemId`；`proc:"Baz"`、`coords:"disk"`、`include_line_numbers:true`）：

| 项 | v1.16.5（修复前） | v1.16.6（修复后） |
|---|---|---|
| 回执打印的首/末行号 | 20 / 23 | 20 / 23 |
| `startLineDisk` / `endLineDisk` | **19 / 22**（错 1 行） | **20 / 23** |
| `startLineDiskExact` / `endLineDiskExact` | 不存在 | **true / true** |

### C. build cache hit：预算组六字段（预算字段缺陷）

`build` 的 **cache hit** 回执（`cached:true`、`path:"cache"`、`route:"cache"`、`buildState:"cached_not_built"`）从 v1.16.6 起**恒定发布**：`timedOut:false`、`budgetEnforced:false`、`budgetMs:0`、`budgetSource:"not_applicable"`、`childPid:0`、`childLeftover:false`。语义 = 「本轮没有编译器子进程，所以子进程预算结构性不适用」，与预检拒绝同形；**其它键与值逐字节不变**（非 cache / timeout 路径完全未触及）。活体（同一工程连做 make_project 后 `build`）：六字段齐全且值如上，`cachedArtifactMatches:true`、`artifactChanged:false`、`advisedReason:"project fingerprint matches the last successful build"`。

### D. 尾随换行矩阵（实测 10 格）

口径：**不是 op 的属性**，也不是「哪种写法更好」；这里只记录**实际结果**，调用方以回执 `trailingBlankLine` + pane 文本的 `ends_with_crlf` 为准（`newTotalLines` 按 `splitlines` 口径，末尾换行不算额外一行）。矩阵（同一 4 行 `.bas`，提交块内容相同，只改「带/不带尾随 CRLF」）：

| 操作 | 提交带尾随 CRLF | pane 结果 `ends_with_crlf` | `trailingBlankLine` | newTotalLines / ideTotalLines |
|---|---|---|---|---|
| `replace_code{set}` | 是 | **false**（被规范化） | false | 5 / 5 |
| `replace_code{set}` | 否 | false | false | 5 / 5 |
| `replace_code{insert}` | 是 | **false**（被规范化） | false | 5 / 5 |
| `replace_code{insert}` | 否 | false | false | 5 / 5 |
| `replace_code{append}` | 是 | **false**（被规范化） | false | 5 / 5 |
| `replace_code{append}` | 否 | false | false | 5 / 5 |
| `replace_code{replace}` | 是 | **false**（被规范化） | false | 4 / 4 |
| `replace_code{replace}` | 否 | false | false | 4 / 4 |
| `add_code` | 是 | **true**（保留） | **true** | 5 / 6（IDE 计入尾随空行） |
| `add_code` | 否 | false | false | 5 / 5 |

**结论（本版只文档化，不统一行为）**：`add_code` 保留提交的尾随 CRLF，四个 `replace_code` op 一律规范化掉；v1.16.5 的回执/描述里曾写「set 保留 / insert 规范化」，与本次实测**不符**（set 也规范化），以本表为准。若要统一（改变任一行为）需要更强的理由并按破坏性变更流程走，本版不做。

### E. 返回信封字段口径：恒在 vs 条件缺席（行号空间/口径缺陷 J）

原则：**能证明的才发布**；不能取得的字段**整键缺席**（不是 0/空串/空对象），并在同族的 `note`/`basis`/`reason` 里给出「为什么没有」。据此分三类：

1. **恒在（每个 tools/call 回执）**：`mcp{product,contract}`、`ide{pid,project,port}`（无 IDE 时为 null/0，并在外层给 `no_live_ide`/`targetMismatch`）、`ok`；`tool` 与 `reason` 在**拒绝**回执里恒在（`ok:false`）；`inflight{}`/`dialogs[]`/`awaiting_user` 等传输与并发上下文按其文档恒在。
2. **按工具恒在**：`verify` 的 `baseline_diff` 恒在（无 baseline 时 `enabled:false` + 0 计数的固定形状）；`check_component` 的 `baseline_diff` 同规则；`build`/`make_project` 的预算组在**所有**路径恒在（v1.16.6 起含 cache hit）；`ideEqualsDisk` 恒在（`SKIPPED` 时 false）。
3. **条件缺席（明确不保证存在）**：`firstError`（只在日志解析出首错时；未解析出 ⇒ 整键缺席，不是空壳）、`file`/`reportedLine`/`diskLine`（只在日志带文件行结构时）、`ideDiskAligned`/`ideDiskAlignedBasis`/`ideDiskAlignScore`（只在做了盘比较时；`SKIPPED` 无）、`startLineDisk*`（只在 `proc`/`symbol` 切片时）、`exitWrites`（只在 quit 被请求时）、`cachedCapture`（只在有缓存捕获时）。**读方不得把「字段不存在」读成 false/0**；判「没做」与「做了但失败」必须用同族的理由字段。

该口径不改变任何既有字段的存在性（纯文档化 + v1.16.6 新增的三个字段族按上述规则），亦不重冻工具面。

### F. 假绿与 reason 语义：活体复核（本版未改代码）

| 项 | 活体结果（v1.16.6） |
|---|---|
| `check_public_api{membersOnly:true, baseline}` | `comparedAgainstBaseline:true`、`baselineNormalized:true`、`baselineLines:3`、`violations:[]`、`added:[]`、`counts{before:3,after:3,removed:0,changed:0,added:0}`，并附 `members[]`（membersOnly 只是**追加**成员表，**不跳过比较**）；追加一个公开成员后同调用 ⇒ `added:[{kind:"function",name:"extrapublic"}]`、`counts.added:1` |
| `check_public_api` 坏 baseline | `ok:false + reason:"baseline_unusable" + comparedAgainstBaseline:false + baselineProblem`，不报任何 added/removed |
| `check_component{baseline}` | `baseline_diff{enabled:true, baselineKeyCount:1, addedCount:0, removedCount:1, unchangedCount:0, removed:[...]}`，证明比较确实发生；干净 baseline（0 key）是合法输入，此时「所有当前诊断都是新增」的读法成立；完全不可解析的文本 ⇒ `ok:false + reason:"baseline_unusable" + lintPerformed:false` |
| 未知参数（`add_code` / `replace_code` / `get_component_code` / `ide_state` 等） | `ok:false + reason:"unknown_argument"` + `tool` + `unknownArguments[]` + `acceptedArguments[]`（`_target` 是传输参数恒接受），**不静默忽略** |
| 组件不存在 | `ok:false + reason:"component_not_found"` + `triedComponent` + `availableComponents[]` + `lookupIsCaseInsensitive:true` + `caseDifferenceIsNotTheCause`（大小写差异按设计**可解析**，故「大小写不匹配」不是可区分的失败类；回执把这条事实写出来）+ `nearMatches[]`/`trimmedMatch` 等线索 |
| 项目未打开 | `ok:false + reason:"target_has_no_project"` + `note`（在读取/写入任何东西之前拒绝）；`get_project_info.noProject:true` 为该状态的正本 |
| 名字大小写 | `add_code{component:"module1"}` 成功解析到 `Module1`（大小写不敏感；因此该场景不产生错误回执，属于**可区分失败类之外的合法成功**） |

以上均为 §「`_meta`/返回信封」之外的**回执字段/行为事实**，不新增工具、不新增参数、不改 inputSchema。
