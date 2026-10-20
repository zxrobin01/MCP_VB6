# Releasing VB6MCP

This is the checklist this project actually uses. Do not skip step 4 or 6: they are the two
checks that have caught real, shipped defects.

## 0. Decide the version

* Package version: `bridge/hub_tools.json` -> `version`.
* Protocol contract: `docs/CONTRACTS.md` - its version table states the current contract (`v1.11.1` as of this writing; always read the table, never a copy).
* Add-in face count and hub-only count are recorded together; the total is the DSH-visible
  tool count.

Change them **deliberately**, not to make a build look newer.

## 1. Clean tree

```bat
git status --porcelain
```

Everything you intend to ship is either committed or covered by `.gitignore`. No
`*.before*`/`*.after*` backup, no build log, no `VB6MCP.dll`.

## 2. Build

```bat
build.bat
```

Record `build.status` and the produced `VB6MCP.dll` SHA-256. **The SHA is the identity of the
release** - every verification item in a release report must name the DLL SHA it was
measured against.

## 3. Re-register the add-in (mandatory)

Every build creates a new typelib/GUID, and VB6 rewrites its own `LoadBehavior` on teardown.
So after every build:

```bat
%SystemRoot%\SysWOW64\regsvr32.exe /s VB6MCP.dll
reg import register.reg /reg:32
tools\verify_registration.bat VB6MCP.dll
```

`verify_registration.bat` must print `EXIT=0`. It reads **two independent gates, and both must be green:**

* **Gate A (DECISIVE) - `%WINDIR%\VBADDIN.INI`, section `[Add-Ins32]`, key `<ProgID>`.** This is the
  switch VB6 consults *before* it ever calls `CoCreateInstance`. `0` means the add-in is dead: VB6 does
  not load it and does not touch `LoadBehavior` (which can still read `0x3`), so every other check stays
  green while the add-in is silently gone. The `0` is written after ONE failed instantiation and **VB6
  never retries** - no rebuild, no `regsvr32` and no `reg import` clears it. To recover, set the key to
  `1` and start the IDE once; **on a successful load VB6 rewrites the value to `3` by itself** (measured:
  the rewrite happens when the add-in loads, not at IDE exit; it then stays `3`). `1` and `3` are both
  enabled; `3` is the steady state. A **missing** entry is not an error - VB6 then falls back to the
  registry and loads - but it is not protection either: one failure creates the entry as `0` and the
  add-in is dead from then on. **This applies to every box, the t16 sandbox included** (t16 has no entry
  today, so it is unrestricted for now, but it would die silently on its first failed load).
  Change the value **in place** (WritePrivateProfileString semantics): measured, VB6's own rewrite touches
  only the value byte, while rewriting the whole file perturbs unrelated bytes. Export a backup before any
  experiment on that file and restore it by sha256 afterwards.
* **Gate B (SECONDARY) - `ProgID -> CLSID -> InprocServer32` plus `AddIns LoadBehavior = 0x3`.** A stale
  or mismatched registration is exactly what makes VB6 write `LoadBehavior=0` as a *consequence*, so this
  gate can be red while the add-in still loads - which is why gate A is the one that decides.

## 4. Tool-face gates (must all be green)

```bat
python tools\check_schema_baseline.py
python tools\check_hub_alignment.py
python tools\check_live_tools.py docs\tools_schema.json
python tools\check_skill_truth.py
python tools\check_private_tokens.py     # no tracked file may carry a private host/project identifier
python tools\run_gates.py                 # THE WHOLE SET in one command (includes the write-call-site ratchet)
rem the next line needs a running IDE (your own instance, 8767 or above - never 8766)
python tools\mcp_strict_audit.py --url http://127.0.0.1:8767/mcp
```

Expected: no schema change, no live-tool change, no hub-alignment failure, no strict failure. The strict audit drives a live IDE, so it needs an instance already running on your own port (8767 or above - never 8766).
If a gate fails because the tool surface intentionally changed, update the frozen artifacts
(`tools/schema_baseline.json`, `docs/tools_schema.json`, `bridge/hub_schema.generated.json`,
`bridge/hub_tools.json`, `docs/CONTRACTS.md`) in the same commit.

## 4.5 Lesson-store audit (mandatory before freezing)

A fixed defect that is still described in the lesson store teaches every reader that the tool is broken - and the banner is
shown to **every** client at initialize. So: audit the store, disable only the entries that assert a defect fixed in this
release (shipped entry -> `enabled:false`; user entry -> removed), keep language/protocol facts and usage rules, and attach
the before/after `validate` plus `instructions` outputs. Full procedure: skills/README.md, "Keeping the lesson store current".


## 4.6 Stamp freshness (mandatory for any batch that changed a component)

A batch that edits component source **must** re-emit the stamp in the repository and prove it, because a stale stamp row makes
every locally built artifact fail `check`:  `python tools/srcstamp.py emit . --vbp MCP_VB6.vbp --out BuildStamp.bas --vbp-patch MCP_VB6.vbp`
then `python tools/check_stamp_current.py --root . --vbp MCP_VB6.vbp` **must exit 0** (all three agree: computed I == `.vbp` src == `BuildStamp.bas`).
`emit` must NOT change the computed hash - if it does, stop and report (it means an input was pulled into or pushed out of the set).

## 5. Freeze check

Re-hash the frozen artifacts and confirm they are byte-identical to what you reported. A
generated cache (for example `bridge/hub_face_cache.json`) is a live artifact and may change;
the frozen files may not.

## 6. End-to-end verification on a real IDE

Start one instance on a port **>= 8767** (never 8766), then re-run the release-critical
scenarios **against the exact DLL you are shipping**, each labelled with that SHA:

* a failing build produces a real compiler log and a first-error location;
* a build that would overwrite a locked artifact does not delete it, and restores it
  byte-identically when it can;
* a dialog that would start a long host-modal operation is reported and not clicked;
* a stale log is labelled as stale and never presented as this call's output;
* a diagnostic baseline survives line insertions (no pre-existing warning is reported as new);
* `ide_restart` stops and restarts only this project's own instance (port >= 8767) and proves
  three-way readiness afterwards.

## 7. Machine hygiene

Release only from an idle machine: every VB6 IDE closed, no listener left on the ports you
tested with, the build lock released if you use one, and the build artifact re-hashed and
unchanged after all testing.

## 8. Tag and publish

```bat
git tag -a v<version> -m "v<version>"
git push origin main --tags
```

Then package the source-only distribution with the export script. It copies the allow-list in
`PUBLIC_TREE.md` and audits the result (75 files at the time of writing; no build products, no
backups, no logs, no internal working directories). The output folder's leaf name must start with
`vb6mcp_pub` - the script refuses anything else, and it refuses a folder inside the repository:

```bat
powershell -NoProfile -ExecutionPolicy Bypass -File tools\make_public_export.ps1 -OutDir C:\Temp\vb6mcp_public -Force
```

It must print `AUDIT_OK` with `FORBIDDEN_NAMES 0` and `NUL_IN_TEXT 0`; if it does not, do not
attach the archive.
---

## 9. Issue 同步口径（长期规则，作者 2026-10-20 明确）

远端（GitHub）issue **只同步"远端版本本身有"的问题**。判据是**远端代码**，不是"我们这次碰到了什么"。

### 必须同步（满足任一条即可）

1. 该缺陷**存在于当前远端版本**（用户拿到的那个 commit 就能复现）；
2. 新发现的问题，但在**远端代码里同样存在**（即：未推送的本地修复之外，远端仍有同样的形态）。

### 不同步

1. **自上次推送以来新增的工具**所引发的问题（远端根本没有这些工具）—— 本地修即可；
2. **过程中遇到的问题**：环境/部署/调试过程、本地构建与冻结流程、内部编号与随包卫生、协作与派单过程等；
3. 仅在本地未推送分支/工作树上存在的回归；
4. 会被**下一次推送直接消除**的问题（写内部记录即可，不开 issue）。

### 操作口径

- 开 issue 前先问一句：**"远端 HEAD 的那个 commit，用户能不能踩到这个？"** 答"不能"就不开；
- 判定要给证据：远端 commit 号 + 该版本中可复现的具体落点（文件/工具/字段）写进 issue 正文；
- 已被下一次推送修掉的，**不补开** issue；若已有 issue，则在推送后回复并关闭，附修复 commit；
- 内部问题一律落在仓库内文档（`docs/HANDOFF_*`、`docs/PROMPT_PACK_*`、`mcp_lessons/`），不占远端 issue 位。

### 本规则的一次应用记录（2026-10-20，v1.12.0 批次）

| 问题 | 存在于远端？ | 处置 |
|---|---|---|
| 6 个新工具 `inputSchema.required` 写成 `{}`（客户端整批拒绝） | ❌ 这些工具**不在远端**（远端 110 工具） | 不开 issue；本地修 + hub 出口规范化 + 新增门禁 |
| hub `list_changed` 时序/基线竞态 | ❌ 该能力**未推送** | 不开 issue；本地修 |
| 中间服务非常驻 / `127.0.0.1:null/rpc` | ❌ 已在推送 6（`11a3e94`）修复 | 不开 issue |
| `CHANGELOG` 顶部数字滞后（远端写 `v1.11.11` / 110，实际 v1.11.12） | ✅ 远端仍是旧数字 | **会被下一次推送消除** ⇒ 不补开 issue |
| `/health` 的 `ok` 语义歧义（`ok` 表示"有活 add-in"而非"服务健康"） | ✅ 远端代码同样如此 | 若下一次推送**未**覆盖，则开 issue（当前计划 v1.12.x 覆盖） |
