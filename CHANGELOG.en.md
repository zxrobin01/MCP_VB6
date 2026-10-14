# Changelog

> [中文](CHANGELOG.md) | **English**

User-visible changes: what changed and **why**. One rule: only what can be verified from the code or the history; unverifiable history stays unwritten. No names of unopened external products, no private project names, no maintainer machine paths appear here.

**Current: product `0.3.0` · contract `v1.11.0` · tool face: `110` (add-in) + `10` (hub) = `120` (client-visible)** - the product version comes from `Util.bas` (echoed by `initialize` and `mcp_info`). The add-in face lives inside the VB6 IDE process; the hub face runs outside it and **works with no IDE online**. The contract version is the compatibility contract of the tool face and receipts (it rises when tools are added/removed or descriptions change; `mcp_info.breaking[]` / `migrations[]` list changes that alter how you must call). This repository ships **source only**: build `VB6MCP.dll` yourself with `build.bat`. Separately, the `.vbp` project file carries its own `1.0.0` version resource, unrelated to the two above.

This file records **source and feature** changes only; packaging, release and other infrastructure work is out of scope.

---

## 2026-10-14

### 20:31 · Contract v1.11.0 · tool face 115 -> 110 (add-in) · the five lesson tools move to the hub face · the IDE addressing key is the .vbp path · explicit remote-endpoint allow-list

**What changed**: (1) **the five lesson tools moved to the hub face** (`list_lessons` / `add_lesson` / `delete_lesson` / `export_lessons` / `import_lessons_md`): name, description and parameters are **unchanged word for word**, so calls look the same; the implementation moved to the Python side (`tools/lessons.py`) and the data stays `mcp_lessons/defaults.json` + `mcp_lessons/user.json` (same schema, an existing `user.json` carries over as-is). Hence **add-in face 115 -> 110**, **hub face 5 -> 10**, and the **client-visible total is still 120**. (2) **the IDE addressing key is now the `.vbp` path instead of the pid**: the canonical form is `_target = "[endpoint:]D:\path\x.vbp"` (endpoint omitted = local); a bare `pid`, a bare project name and `auto` still work, but only as **legacy aliases** and are no longer reported as the canonical key. Paths are normalized case-insensitively, `\` and `/` are equivalent, trailing separators are trimmed, and a local path that exists is additionally resolved to its real path; when one selector matches several identities the call is **refused structurally** (`target_conflict`) with the candidates listed - it is **never resolved to one for you**. (3) **a call without `_target` is no longer auto-routed to a remote**: only when **no local IDE is live while a remote one is** does it return the structured `remote_requires_explicit_endpoint` (`routed:false`, with endpoint-qualified keys you can paste verbatim); with a live local IDE it still selects automatically. (4) **remote endpoints come only from an explicit allow-list** (the `VB6MCP_REMOTES` environment variable or `bridge/hub_remotes.json`): the hub **probes only the listed endpoints and never discovers hosts on its own**; with nothing configured every remote code path is inert (behaviour identical to before the change). Failures are always structured: `endpoint_unknown` (a name that is not configured) / `endpoint_unreachable` / `endpoint_timeout`.
**Why**: the lesson store is a **global singleton** that does not belong to the IDE on any one machine - keeping it in the add-in made the whole capability disappear whenever **no IDE was online**, and made **changing one lesson require rebuilding the add-in**; moving it outside means it **works with no IDE online**, and changing a lesson is just editing a JSON data file that is read at runtime. A pid is an identity that **changes on every launch**, so a remembered default target could point at an instance that no longer exists; the `.vbp` path is the project's stable identity and the only key that can express both "two different projects on one machine" and "the same path on two machines" (the latter is separated by the `endpoint` prefix). The old rule "with no target, pick the only live instance" would, once remotes existed, **silently send a call to another machine**; now automatic selection only ever stays local, a remote must be named explicitly, and "only a remote is online" became a readable refusal code. Since a remote host has to be configured explicitly, it must not be auto-discovered: no address exists outside the allow-list, probing targets only the allow-list, and failures carry a structured reason and elapsed time.

## 2026-10-13

### 16:26 · Tool face 114 -> 115 · live desktop state · long code can come from a file · one failure envelope

**What changed**: added a read-only tool **`desktop_probe`** that reports the **live** desktop state of the process hosting the add-in (window station / desktop / session / foreground) and **refuses honestly** (`ok:false` + `failureKind`) when an API cannot be read, instead of returning a zero-filled picture; `vba6_probe` / `immediate_eval` / `immediate_run` each gained one **optional** "read it from a file" parameter (`exec_line_file` / `expression_file` / `code_file`); `vba6_probe` failures now use a **uniform envelope** (`ok` / `failureKind` / `reason`) and report the **requested vs clamped length**, and an out-of-range argument is a structured `arg_out_of_range` instead of a raw overflow; `ui_text` gives clearer diagnostics when a self-drawn panel cannot be read; **internal labels were purged from the public tool descriptions**.
**Why**: the desktop / session state had **no tool at all**, so callers could only guess; long or quote-heavy code was regularly broken by command-line length and quoting rules; without a uniform failure reason a caller cannot tell "bad argument" from "engine unavailable"; internal labels in descriptions put unverifiable text on the public surface. A previously reported claim that the foreground value "always reads 0 on this machine" was **disproved by measurement** and is corrected everywhere: desktop state is now reported as a **live value**, and send-input capability is **never** inferred from a window-station name. Note that the **Immediate-window body text still cannot be read** (that pane exposes no readable body), so **no** tool was added for it and the limitation stays documented.

### 15:08 · Compilation no longer hangs · tighter defaults in the snapshot and audit scripts

**What changed**: the modal guard now **restarts its deadline whenever it sees a box**, and the compile path's deadline covers the **whole call** => compilation no longer hangs but returns a structured result (the IDE box is dismissed with a safe button); the tool-face snapshot script now **writes no file by default**, and the strict audit script **requires an explicit target** and **refuses the machine's reserved port**.
**Why**: while compiling, the IDE pops its own "Make Project" box, and that box appears **later than the guard deadline** => the call never returned; the snapshot script used to **overwrite the contract baseline** by default; the audit script's default target happened to be the **machine's reserved port**.
### 14:50 · Breakpoint line pre-check · modal safety guard

**What changed**: before toggling a breakpoint the tool now **checks whether that line allows it** (if not it **refuses structurally** instead of triggering the IDE error box); added a **modal safety guard** - when a call is blocked by a modal box the IDE itself popped up, it **clicks only safe buttons** and **reports honestly** (the box content, which button was clicked, elapsed time, side effects).
**Why**: toggling a breakpoint on a line that does not allow it made the call **hang for about a minute** and left every later write operation answering busy; an accidentally triggered throwing statement could also pop a modal box and **stall the whole tool chain**.
### 14:23 · The wheel works immediately in a newly opened code pane

**What changed**: after a **new code pane** opens, the **mouse wheel now works immediately** (no need to click elsewhere first); this comes with a **250 ms lightweight hook timer** and per-window hook diagnostics.
**Why**: the only driver that used to hook new panes was a **5-second watchdog** (measured: 2 scans in 10 seconds) - so for a few seconds after a new window opened the **wheel was bound to do nothing**, then recovered on its own.
### 14:04 · In-process execution and read-back · session-level quiet mode · forced gate release

**What changed**: added an **in-process execution and read-back** capability (**returns values directly and captures output**, so callers no longer have to write a file to get data out); added a **session-level quiet mode** (suppresses tools stealing the foreground); when a call is stuck on one of the IDE's own modal boxes, the **gate can now be force-released** so later calls recover.
**Why**: runtime acceptance used to require **a human watching the IDE Immediate window**; a tool stealing the foreground can drop the user's **in-progress keystrokes into the code pane**; a single accidentally triggered throwing statement could **lock the whole tool chain** with no way to recover.
### 13:29 · Cumulative line numbers · operation receipts and refusals · foreground and pane protection · conditional-compilation arguments

**What changed**: component disk line numbers now **accumulate per line** (no longer too low when hidden inline Attribute lines exist); the "delete control" and "open project" receipts now carry a **success flag, before/after counts and the memory-vs-disk difference**, and **refuse structurally** when it is unsafe; tools that touch the caret or a pane now **report and restore the foreground**, and **verify the pane text is unchanged** across the operation; conditional-compilation argument values are **no longer quoted**, and the receipt shows the actual command line.
**Why**: a low line number can silently edit the wrong line; "looks successful" or "silently ineffective" makes callers misjudge; when the user is typing on the same machine, stealing focus can drop keystrokes into the code pane and corrupt the source; quoting silently disables the conditional-compilation macro.
### 10:05 · `remove_trace` no longer removes `modMcpTrace`

**What changed**: `remove_trace` deletes only the injected trace block; it no longer deletes the `modMcpTrace` module.
**Why**: that module may contain code you added yourself, and other places in the project may still use it - deleting it would lose code or break the build.

### 09:53 · Public docs no longer cite an unverifiable external product

**What changed**: the READMEs and a code comment no longer cross-reference an unopened external product; the example config now uses a placeholder path.
**Why**: readers cannot check that external reference, so it only confuses; a hard-coded local path cannot be used by anyone else.

### 03:16 · First public release

Open-sourced as source only (Apache-2.0): 111 MCP tools that turn a real VB6 IDE into a programmable interface.
