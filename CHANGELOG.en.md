# Changelog

> [中文](CHANGELOG.md) | **English**

User-visible changes: what changed, **why**, and **what it means for you**. One rule: only what can be verified from the code or the history; unverifiable history stays unwritten. No names of unopened external products, no private project names, no maintainer machine paths appear here.

**Current: product `0.3.0` · contract `v1.8.0`** - the product version comes from `Util.bas` (echoed by `initialize` and `mcp_info`); the contract version is the compatibility contract of the tool face and receipts (it rises when tools are added/removed or descriptions change; `mcp_info.breaking[]` / `migrations[]` list changes that alter how you must call). This repository ships **source only**: build `VB6MCP.dll` yourself with `build.bat`. Separately, the `.vbp` project file carries its own `1.0.0` version resource, unrelated to the two above.

---

## 2026-10-13

### 10:05 · `remove_trace` no longer removes `modMcpTrace` automatically

**Type**: change (tool description rewritten; tool-face contract re-frozen)

- **What changed**: `remove_trace` now deletes **only the `@mcp-trace` block of the id you name**; the `modMcpTrace` component is **always kept in the project**.
- **Why**:
  1. the auto-generated module lives in **your project** - you may have added your own code to it, and removing it would destroy that code too;
  2. other `@mcp-trace` blocks may still exist elsewhere in the project, and they all need `McpTrace` / `McpTraceOn` from that module to compile - removing it would break the build.
- **What it means for you**:
  - the receipt always reports `moduleRemoved: false`, plus `moduleLeftInPlace`, `moduleRemoveReason` (`kept_by_design` when it is kept by design, `module_absent` when the project never had one), `moduleKeptNote` and `moduleFileOnDisk` (the file it is bound to on disk);
  - to clean it up **by hand**: first confirm that **no** `@mcp-trace` block remains anywhere, then call `remove_component{name:"modMcpTrace"}` yourself. Removing a component does **not** delete the `.bas` file on disk;
  - the tool **description was rewritten** and the **tool-face baseline was re-frozen**: treat `tools/list` and `docs/tools_schema.json` as authoritative;
  - the matching README section (both languages) was rewritten as: (1) what it is for, (2) whether it is required (VB6MCP itself does not need it), (3) when you do need it, (4) how MCP generates and attaches it, (5) when the reference is added or removed.

### 09:53 · Public documentation and example-config cleanup

**Type**: fixes / change / docs

- **Fix - machine-absolute path in the example config**: `mcp_config.example.json` used to contain the maintainer machine path; copying it could not work for anyone else, and it exposed the author directory layout. It now uses the placeholder `C:/path/to/MCP_VB6/bridge/mcp_stdio_bridge.py`, and both READMEs state clearly that **every path in the config must be replaced with one from your own machine** (both the python/py `command` and the `args` path to the bridge script).
  - Effect: a new user changes one place and it works; the published tree contains no machine-specific path.
- **Fix - machine-path matching in the release audit**: the audit now matches machine paths **separator-insensitively (`\` and `/` normalised) and case-insensitively**.
  - Why: the old check only understood backslashes, while the config file used **forward slashes**, so a tree that clearly contained the machine path reported "0" (a **false negative**).
  - Effect: this class of leak can no longer slip through because of separator style; the pre-publish self-check also hard-asserts that local-only files never enter the published tree (a hit fails the run).
- **Change - removed unverifiable external provenance**: both READMEs and one script comment no longer carry a cross-reference to an **unopened external product**; the public docs no longer cite something a reader cannot check.
- **Docs**: the pre-publish checks (including the deny words that must be passed explicitly, and the post-export self-check) are now written down in the maintainer notes.

### 03:16 · First public release

**Type**: added

- Published for the first time under **Apache-2.0**, **source only** (no binaries). Build `VB6MCP.dll` with the README "Building" section, register it, and you are ready.
- Ships **111 MCP tools** that turn a real VB6 IDE into a programmable interface: read/search/edit code, compile and compile diagnostics, breakpoints and run control, forms and design-time properties, project snapshots, the lesson store and the extension store.
- Effect: AI coding assistants (Claude Desktop / Cursor / Codex, ...) can drive the **VB6 project you have open** through the stdio bridge.

---

## How to read these entries

- **Time**: sections are per day; entries within a day are ordered by **time** (newest first).
- **Type**: added / fixed / changed / docs / removed.
- **Scope**: this file covers the first public release and everything after it; earlier internal history is out of scope.
- **Fact boundary**: only what can be verified from the code or the history is written down - unverifiable history stays silent rather than being filled in with a plausible-sounding guess.
