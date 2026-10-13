# 为 VB6MCP 做贡献

> [English](CONTRIBUTING.md) | **中文**

感谢帮忙。VB6MCP 把 Microsoft Visual Basic 6.0 的 IDE 变成一个
[MCP](https://modelcontextprotocol.io) 服务器：一个 VB6 插件（工具面）加一个 Node MCP
宿主（`bridge/`）。

## 一、前提（请先读这一段）

VB6MCP 是 **VB6 的插件**，所以你需要自己的一套可用 VB6 安装：

| 需要 | 为什么 | 从哪来 |
|---|---|---|
| **Visual Basic 6.0 IDE**（`VB6.EXE`、`VB98`） | 插件必须加载进它；每个工具都在驱动一个活着的 IDE。 | 你自己的正版 VB6 安装。本项目**不**分发它，也不可能分发。 |
| **`MSADDNDR.DLL`** | VB6 的 Add-In Designer 类型库，编译插件需要它。 | 从同一套 VB6 安装里复制到 `tests/_msaddndr.dll`（已被 git 忽略）。 |
| **Node.js 18+** | 运行 MCP 宿主 `bridge/vb6mcp_hub.mjs`。 | nodejs.org |
| **Python 3.8+** | 运行 `tools/` 下的维护与验证脚本。 | python.org |

**不需要任何第三方包**：宿主只用 Node 标准库，工具只用 Python 标准库（见 `THIRD_PARTY_NOTICES.md`）。

## 二、构建

```bat
build.bat
```

构建产出 `VB6MCP.dll`（已被 git 忽略）。**每次构建后都必须重新注册插件**，因为每次构建都会生成
新的类型库/GUID：

```bat
%SystemRoot%\SysWOW64\regsvr32.exe /s VB6MCP.dll
reg import register.reg /reg:32
```

> 注意：**必须用本机的 `%SystemRoot%\SysWOW64\regsvr32.exe`，绝不要硬编码 `C:\Windows`**。
> 有些机器上 `SystemRoot` 不是 `C:\Windows`，而 `C:\Windows` 可能是一份**陈旧的第二安装**——
> 用它会让**任何** 32 位 DLL 注册失败（`0xC0000139`），且带 `/s` 时**静默失败**。

然后验证注册（这一步专门抓那个经典故障：注册返回成功、插件却从不出现）：

```bat
tools\verify_registration.bat VB6MCP.dll
```

仓库根目录的 `unregister.reg` 可以卸载注册。

## 三、开 PR 之前请先跑这些检查

这些就是本项目真正使用的门禁。

**离线、只读**（不需要 IDE、不需要构建，只读仓库里的文件）：

```bat
rem 仓库内快照与冻结基线一致
python tools\check_schema_baseline.py
rem 宿主面与插件面一致
python tools\check_hub_alignment.py
rem 仓库内快照与冻结基线交叉比对
python tools\check_live_tools.py docs\tools_schema.json
```

**需要活着的 IDE**（你自己的实例，端口 **8767 及以上**——永不用 8766）：

```bat
rem 对整个工具面做严格审计 —— 需要活着的 IDE
python tools\mcp_strict_audit.py --url http://127.0.0.1:8767/mcp
rem 端到端冒烟测试 —— 需要活着的 IDE
python tests\smoke_test.py --port 8767
```

`tools\verify_registration.bat VB6MCP.dll`（见第二节）是读注册表，需要先构建出 DLL。

对工具面的改动必须是**有意的**：冻结件是 `tools/schema_baseline.json`、`docs/tools_schema.json`、
`bridge/hub_schema.generated.json`、`bridge/hub_tools.json` 与 `docs/CONTRACTS.md`。如果你的改动是
有意的，就在**同一个提交**里更新它们并在说明里写清；否则 `check_schema_baseline.py` 会失败——
**这是设计如此**。注意 `python tools\gen_hub_snapshot.py` 会**改写冻结件**（`bridge/hub_schema.generated.json`），所以它属于那次有意的改动，不属于上面的只读检查。

## 四、测试

* `tests/battery_test.py` —— 端到端测试组。它需要一个活着的 IDE（在某个工程上运行 `VB6.EXE`）并
  驱动真实工具；模式说明见文件头。
* **一个端口一个实例**。端口 **8766** 保留给维护者自己的交互会话；测试实例必须绑定 **8767 及以上**。

## 五、项目规矩

* **源码编码**：VB6 的 `.bas`/`.cls`/`.frm` 以 **ANSI（中文 Windows 上是 GBK）** 存放。
  不要把它们转成 UTF-8——VB6 会把其中的非 ASCII 字符串字面量弄乱。
* **源码或文本文件里绝不允许出现 NUL 字节**。我们的一部分工具会因此拒绝处理。
* **VB6 源码单行有硬上限：1023 个字符。** 更长的行会编译失败（实测：1023 能编译，1028 报语法错误）。
  长单引号字符串请用 ` & _` 续行拆开——字符串的值不变。★编译器报「第 N 行语法错误」时，还要看 N
  及其之后的行：对**续行语句**，VB6 报的是**语句起始行**，所以被报的那一行可能看起来完全正常。
* **版本号是有承载意义的**。包版本在 `bridge/hub_tools.json`，契约版本在 `docs/CONTRACTS.md`；
  **不要**为了让构建看起来更新而改动它们。
* **每一份回执都必须诚实**。工具做不到某事时，必须给出**机器可读的原因**，而不是悄悄降级。
  这是本项目最重要的一条规矩；以诚实换取更好看的绿灯状态的 PR 会被拒绝。
* **不要提交** `VB6MCP.dll`、构建日志、`*.before*`/`*.after*` 备份 —— 公开的 `.gitignore` 已经覆盖
  这些**通用**产物。维护者的私有工作目录与一次性脚本不属于公开树，请勿随 PR 提交；它们只在维护者
  本地通过 `.git/info/exclude` 排除，**不在**公开发布的 `.gitignore` 里。

## 六、报告缺陷

请附上：

1. `mcp_info` 的输出（契约版本、工具数、`build.sha256`）——它标识了你正在运行的**确切二进制**。
2. 失败调用的**完整回执**（不是摘要）。
3. IDE 当时是设计/运行/中断模式中的哪一种，以及它在哪个端口上。

## 七、许可证

提交贡献即表示你同意你的贡献以 Apache License 2.0 授权（见 `LICENSE`）。
