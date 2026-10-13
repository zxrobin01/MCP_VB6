# Changelog

> [中文](CHANGELOG.md) | **English**

User-visible changes: what changed and **why**. One rule: only what can be verified from the code or the history; unverifiable history stays unwritten. No names of unopened external products, no private project names, no maintainer machine paths appear here.

**Current: product `0.3.0` · contract `v1.8.0`** - the product version comes from `Util.bas` (echoed by `initialize` and `mcp_info`); the contract version is the compatibility contract of the tool face and receipts (it rises when tools are added/removed or descriptions change; `mcp_info.breaking[]` / `migrations[]` list changes that alter how you must call). This repository ships **source only**: build `VB6MCP.dll` yourself with `build.bat`. Separately, the `.vbp` project file carries its own `1.0.0` version resource, unrelated to the two above.

This file records **source and feature** changes only; packaging, release and other infrastructure work is out of scope.

---

## 2026-10-13

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
