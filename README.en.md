# VB6MCP — turn the VB6 IDE into an MCP server

> [中文](README.md) | **English**

VB6MCP is a **VB6 add-in**: once loaded into the VB6 IDE, the IDE itself becomes a
[MCP (Model Context Protocol)](https://modelcontextprotocol.io) server. AI coding assistants
such as Claude Desktop, Cursor or Codex can then read, search and modify the VB6 project that
is open in the IDE, and drive compilation and execution.

```
 MCP client (Claude Desktop / Cursor / ...)
        |  stdio (JSON-RPC, one message per line)
        v
 bridge/mcp_stdio_bridge.py   (Python 3, standard library only)
        |  HTTP POST /rpc  (http://127.0.0.1:8766)
        v
 VB6MCP.dll  <- runs inside the VB6.EXE process
   |-- Winsock control: local HTTP server (/sse + /message also supported)
   |-- home-grown JSON parser/serializer (JsonLib.bas, zero external dependencies)
   +-- VBIDE extensibility object model: read/modify code, save, compile, control execution
```

## This is an AI-assisted project

This project is **AI-assisted and was not written line by line by hand**. Saying so plainly
matters to both users and contributors.

- **The code**: the large majority of the source was generated and modified by AI coding
  assistants under human supervision. Humans own the requirements, the trade-offs, the
  acceptance decisions and the **real-machine verification** - every key claim in this project
  is required to have reproducible evidence taken on a real VB6 installation.
- **The method**: the development process is deliberately organised around verifiability -
  work is issued as task cards, every claim carries machine-readable evidence, critical changes
  are re-checked by an **independent verifier** (implementer and verifier are separate roles),
  the tool surface is guarded by **frozen-baseline gates**, and a tool that cannot do something
  must return a **machine-readable reason** instead of degrading silently.
- **No guarantee of human line-by-line review**: do not assume a human read every line. Assess
  it yourself before using it in production or building on it, and open an issue if you find a
  problem (see CONTRIBUTING).
- **Known limitations**: see section 7 - in particular that **this project is developed and
  verified on a Chinese (zh-CN) VB6**.
## Why this exists (background)

VB6 stopped evolving in 2008, but the systems it wrote are still running: factories, finance,
hospitals and government back offices are full of old VB6 projects that still need maintenance,
and the people maintaining them face an environment with **no modern tooling at all** - no
language server, no dependable static checking, compiler errors that can only be clicked
through by hand, and edits made by hoping your hand does not slip.

AI coding assistants handle modern languages well, but they **cannot reach VB6**: they can
neither see the project open in the IDE nor click that "Make Project" dialog. This project
fills that gap.

**The approach**: rather than building another editor, **turn the VB6 IDE itself into an MCP
server**. Load one add-in and the IDE exposes 111 tools: read/search/edit code, static
checking, anchor-based editing with safe transactions, compilation, run and single-step,
dialog handling, form and project metadata, snapshots and rollback, packaging. **Because the
tools live inside the IDE process**, they can do what an out-of-process approach cannot:
actually click that dialog, actually press Run, actually read the local variables at a
breakpoint.

**The goal**: let people who maintain old VB6 (including the author) use a modern AI workflow.
The project is Apache-2.0 and allows arbitrary extension - **the better others make it, the
better the features I get to use**.

## Relationship to similar projects

Modern tooling for VB6 is not a blank slate. The known routes are:

| Project | Route | Relationship to this project |
|---|---|---|
| Agwstin/vb6-lsp (includes mcp-bridge.mjs) | VB6 language server + MCP bridge | **Out of process**: strong edit-time analysis, but it cannot drive IDE compilation, debugging or dialogs |
| xlide / xlide-mcp (VS Code extension, now at v10) | VB6/VBA support inside VS Code | Editor-side assistance; it does not operate a real VB6 IDE |
| wieslawsoltes/VB6 (VB6 Studio Web) | Rebuilding the VB6 IDE in the browser | A different route, aimed at rewriting rather than maintaining existing projects |
| Various VB6 static-analysis MCPs | Read-only source analysis | Read-only: cannot modify, compile or run |

**What is different here**: it exposes the **complete operating surface of a real VB6 IDE**
(compile, run, breakpoints, dialogs, forms, packaging) as MCP, and treats **receipt honesty**
as part of the product: when a tool cannot do something it returns a machine-readable reason
instead of pretending to succeed.

## Directory layout

```
MCP_VB6/
|-- MCP_VB6.vbp            VB6 project file (ActiveX DLL, output VB6MCP.dll)
|-- Connect.cls            add-in entry point, implements IDTExtensibility2
|-- ServerHost.cls         assembles the server + Add-Ins menu items
|-- McpServer.cls          MCP/JSON-RPC dispatch + HTTP routing
|-- HttpConn.cls           HTTP/1.1 request buffering and parsing (keep-alive supported)
|-- ServerForm.frm         hidden form hosting the Winsock listener + port-probe timer
|-- CodeTools.cls          the 111 MCP tools (full VBIDE object model + file system + syntax gate / edit safety / compile authority)
|-- JsonLib.bas            zero-dependency JSON parser/serializer
|-- Utf8.bas               UTF-8/ANSI/UTF-16 conversion (kernel32 APIs)
|-- Util.bas               logging, port file
|-- register.reg           registry entries that register the add-in with VB6 (import once)
|-- unregister.reg         remove the add-in registration
|-- build.bat              one-click build + regsvr32 registration + import the add-in registry file
|-- mcp_config.example.json    Claude Desktop configuration example
|-- bridge/mcp_stdio_bridge.py stdio bridge (Python 3 standard library)
+-- tests/smoke_test.py     smoke test that talks to the HTTP endpoint directly
```

## 1. Building

**Prerequisites**: a VB6 IDE installation (SP6). `MSWINSCK.OCX` (the Winsock control) is
installed and registered by VB6.

1. Run `build.bat` (if VB6 lives elsewhere, edit the `VB6` variable inside it).
   The script first registers the Add-In Designer type library (`MSADDNDR.DLL`), then compiles.
2. On success it produces `VB6MCP.dll` and automatically runs `regsvr32` and imports `register.reg`.
3. Start the VB6 IDE, open the **Add-Ins -> Add-In Manager** menu, and tick
   **"VB6 MCP Server"** (the registry holds LoadBehavior=3, so it usually loads automatically).

You can also open `MCP_VB6.vbp` in the IDE and compile it there (references needed: VB6
Extensibility and Microsoft Add-In Designer; if compilation reports anything about
`IDTExtensibility2`, see the FAQ below).

> Note: do not make the VB6MCP project itself the active project and use "Make" - the add-in
> DLL is loaded and held by that very IDE, so the file lock makes the build fail. Close the IDE
> and use build.bat, or test from a different project.

## 2. Ports and multiple instances

- Listens on `http://127.0.0.1:8766` by default (MSWINSCK's LocalIP is read-only and the
  listener binds every local interface, but incoming connections are validated: the peer IP
  must be 127.0.0.1, and non-loopback connections are rejected immediately).
- If the port is taken it probes upwards (up to +40).
- The actual port is written to `%TEMP%\vb6mcp_port.txt`, which the bridge reads automatically.
- Custom port: the `VB6MCP_PORT` environment variable, or the registry value
  `HKCU\Software\VB6MCP\Port` (DWORD).
- With several VB6 instances open, each takes its own port; the port file records the most
  recently started instance.
- Status: IDE menu **Add-Ins -> VB6 MCP Status...**, or `http://127.0.0.1:8766/health`.
- Log: `%TEMP%\vb6mcp.log`.

## 3. Connecting an MCP client

### Claude Desktop

Merge the `mcpServers.vb6` entry from `mcp_config.example.json` into
`%APPDATA%\Claude\claude_desktop_config.json` (point `command` at the python/py path on your
machine) and restart Claude Desktop.

> **Replace every path with one from your own machine**: the paths in `mcp_config.example.json` (and in the "Other stdio clients" example below) are **placeholders** - both `command` (python/py) and the `args` path to `bridge/mcp_stdio_bridge.py` must point at your own installation. In JSON prefer forward slashes `/`; a Windows backslash must be written as `\\`.

### Cursor (SSE / streamable HTTP)

Cursor can talk to the HTTP endpoint directly, with no bridge:

```
type: sse     url: http://127.0.0.1:8766/sse
```

### Other stdio clients

```json
{
  "mcpServers": {
    "vb6": {
      "command": "python",
      "args": ["D:/path/to/VB6MCP/bridge/mcp_stdio_bridge.py"]
    }
  }
}
```

To check the whole chain:

```
python tests/smoke_test.py --port 8766
```

## 4. Tool list (111 tools)

> The authoritative list is `docs/TOOL_CATALOG.md` (exported from a running add-in via
> `tools/list`); the offline baseline is `tools/schema_baseline.json`, and
> `tools/check_live_tools.py` compares the live face against that baseline.

The table below covers the original 45 tools (the complete VBIDE object-model part). The other
66 were added later for syntax gating / edit safety / compile authority / snapshot consistency
/ test hosting / runtime observation / debug control / VBA6 probing / governance
(`validate_code`, `check_component`, `read_source`, `write_source`, `edit_by_anchor`,
`apply_edits`, `preview_edit`, `build`, `verify`, `inject_trace`, `remove_trace`,
`fault_inject`, `clear_faults`, `set_breakpoint`, `clear_breakpoints`, `debug_commands`,
`execution_state`, `native_breakpoint`, `project_changes`, `vba6_probe_plan`, `audit_tail`,
`export_public_api`, `check_public_api`, `dependency_graph`, `purity_check`, `generate_stub`,
`record_fixture`, `compare_fixture`), plus the **lesson store** (`list_lessons`, `add_lesson`,
`delete_lesson`, `export_lessons`, `import_lessons_md`) and the **extension store** (`ext_list`,
`ext_enable`, `ext_disable`, `ext_reload`, `ext_install`, `ext_uninstall`).

### Lesson store and extension store

- **Lesson store** `mcp_lessons/`: `defaults.json` (shipped with the add-in, do not edit) plus
  `user.json` (the user layer, which may override, disable or add entries by id). The merged
  result is published to clients through `initialize.instructions`, so an agent can read it
  before calling any tool; `add`/`delete`/`import` write `user.json` with a `.bak` backup.
- **Extension store** `mcp_extensions/`: every `*.json` is a **declarative tool pack**
  (`name` / `tools[].inputSchema` / `tools[].handler`). An enabled pack's tools are merged into
  `tools/list` automatically and dispatched by `CodeTools.CallTool`; `state.json` records the
  disabled list (enabled by default). Handlers expose only a few safe primitives (`http_get`,
  `read_file`, `write_file`) and **never load VB6 code from disk**.

The full object-model inventory and gap analysis live in the maintainer's internal VBIDE
object-model notes (produced by parsing the VB6EXT.OLB 5.3 type-library binary; the raw data is
in the internal type-library dump).

| Tool | Purpose | Key parameters |
|------|---------|----------------|
| `get_project_info` | Project/workspace overview: type, IDE state, startup object, build path, components, references, open projects | - |
| `set_project_options` | Project options: description, help file/context, startup object, startup project, build output path, active project, read-only mode | `description`, `startup_object`, `build_file_name` ... |
| `open_project` / `add_project` / `remove_project` | Open a .vbp / create a project (or from a template) / close a project | `path`, `type`, `template_path`, `name` |
| `get_component_list` | Component list: procedures, line numbers, dirty state, designer info | - |
| `get_component_code` | Read a component's full code from IDE memory | `name`, `include_line_numbers` |
| `replace_code` | Line-level edit: replace/insert/append/delete/**set (replace the whole module)** | `name`, `operation`, `start_line`, `end_line`, `code` |
| `add_component` / `remove_component` / `import_component` / `export_component` / `rename_component` / `reload_component` | Full component lifecycle (use reload together with write_file) | `type`, `name`, `path`, `new_name` ... |
| `add_code` / `create_event_proc` / `find_in_component` | Append code (string or file) / create an event-procedure stub / search inside a module | `component`, `code`, `file`, `object`, `event`, `target` |
| `get_codepane` / `set_codepane` | Code pane: caret/selection/scroll/view mode | `component`, `selection`, `top_line`, `view` |
| `get_selection` / `activate_component` | Current IDE selection / activate a component (code or designer window) | `name`, `window` |
| `list_files` / `read_file` / `write_file` / `search_in_files` | Disk file operations (UTF-8/UTF-16/ANSI auto-detected) | `path`, `pattern`, `query` ... |
| `save_project` | Save the project and every dirty component | - |
| `make_project` | **Silent compile** (VBProject.MakeCompiledFile, output path optional; falls back to the menu route on failure); the artifact gets LAA by default; `diagnose:true` compiles an isolated copy for a real compiler log | `output_path`, `laa`, `diagnose` |
| `control_execution` | Execution control: run / full_compile / break / end / restart | `mode` |
| `add_reference` / `remove_reference` | References: add by GUID or file path / remove by name | `mode`, `guid`, `path`, `name` |
| `get_component_properties` / `set_component_property` | Design-time properties (the Properties window): read/set, including multi-dimensional indices | `name`, `property`, `value`, `indices` |
| `list_members` / `set_member_attribute` | Class members: type/scope/attributes; set Hidden/Browsable/Category and friends | `component`, `member`, `attribute`, `value` |
| `list_windows` / `control_window` | IDE windows: enumerate/focus/close/show/hide/minimize/move/rename | `window`, `action` ... |
| `form_designer` / `add_control` / `remove_control` / `set_control_property` | Form designer: enumerate controls, selection, clipboard, add/remove controls, set control properties | `form`, `action`, `prog_id`, `index` |
| `property_bag` | Project/component property bag read/write (ReadProperty/WriteProperty) | `target`, `key`, `value` |
| `list_addins` / `set_addin_state` | Enumerate add-ins / load or unload an add-in | `prog_id`, `connect` |
| `command_bars` | IDE menus: enumerate, or execute a command by name | `action`, `bar`, `control` |
| `quit_ide` | Quit the IDE (executed shortly after the response; unsaved work is lost) | - |

## 5. Two ways to modify code (important)

1. **`replace_code` (recommended; operates on IDE memory)**: changes take effect in the IDE's
   code immediately and the component is marked dirty; use `save_project` afterwards to write it
   to disk.
2. **`write_file` (writes to disk)**: good for new files or non-VB files. If the target is a
   file already open in the project, use **`reload_component`** to pick up the disk version (no
   IDE restart needed).

## 6. Encoding

- VB6 source files (`.bas/.frm/.cls/.vbp/...`) are read and written as **ANSI (the system code
  page, e.g. GBK)**; other files default to UTF-8. `write_file` accepts an explicit `encoding`
  of `auto / utf-8 / ansi / utf-16`.
- The MCP wire protocol is always UTF-8; the bridge uses `ensure_ascii=False`, so non-ASCII
  content round-trips intact.

## 7. Known limitations

- VB6 is single-threaded STA: large searches briefly block the IDE UI (a 2000-file cap is
  enforced).
- In break mode `make_project` / `control_execution` may fail; the tools return the error text.
- `make_project` acts on the active project and cannot silently redirect the output path (it
  follows the IDE's project settings).
- The add-in DLL cannot be rebuilt in place while the IDE holds it (see the FAQ).
- Incoming connections are forced to 127.0.0.1 (MSWINSCK's LocalIP is read-only, so the loopback
  address cannot be bound before listening and the check is done on accept instead). There is no
  authentication - only local processes are trusted.
- **Localization**: this project is developed and verified on a **Chinese (zh-CN) VB6**. The IDE
  title's mode markers, the "Immediate" window name and some dialog button captions are matched
  in Chinese; an English VB6 shows `[Design]`/`Immediate` and so on, which means **mode
  detection and some dialog flows may fail**. The hub side (`bridge/`) is locale-independent
  (it only assumes the `VB98` folder name, which is the same in every language). See section 1
  of the known-gaps list in section 7 of this file.

## 8. FAQ / troubleshooting

**Compile error "User-defined type not defined: IDTExtensibility2"**
-> You need the **Microsoft Add-In Designer** type library: `MSADDNDR.DLL` (typelib GUID
`{AC0714F2-3D04-11D1-AE7D-00A0C90F26F4}`, usually at
`C:\Program Files (x86)\Common Files\designer\MSADDNDR.DLL`). `build.bat` registers it
for you (via `%SystemRoot%\SysWOW64\regsvr32.exe`); to do it by hand:
`"%SystemRoot%\SysWOW64\regsvr32.exe" "%CommonProgramFiles(x86)%\designer\MSADDNDR.DLL"`.

**Compile error "Method or data member not found" (e.g. Member/Reload/MakeCompiledFile)**
-> The project is referencing Office's VBA extensibility type library ({0002E157}, 64 types,
missing the VB6-specific objects). Tick **Microsoft Visual Basic 6.0 Extensibility** in
Project -> References (`VB98\VB6EXT.OLB`, GUID `{EF404E00-EDA6-101A-8DAF-00DD010F7EBB}`, 112
types). This project's .vbp pins that reference; if the IDE switches it back to the Office
version, change it back by hand.

**Notes on adapting to this machine (already internalised in the code)**
- MSWINSCK.OCX's `LocalIP` is read-only: the listener binds all local interfaces and validates
  the peer on accept. If the OCX is not registered, run
  `"%SystemRoot%\SysWOW64\regsvr32.exe" "%SystemRoot%\SysWOW64\MSWINSCK.OCX"` before
  `build.bat`. **32-bit control registration must use the machine's own
  `%SystemRoot%\SysWOW64\regsvr32.exe`; never hardcode `C:\Windows`**: on some machines
  `SystemRoot` is not `C:\Windows`, and `C:\Windows` may be a stale second installation, which
  makes **any** 32-bit DLL registration fail (`0xC0000139`) and silently so with `/s`.
  (`build.bat` therefore points `%REGSVR%` at the machine's own SysWOW64.)
- In VB6EXT, `CodePane.CodePaneView` and `Window.Caption` are read-only; requests to set them
  are ignored and explained in the response.
- `VBProjects.StartProject` has a Put accessor only, so object assignment without Set is
  required.
- VB6 reserved-word traps: a procedure cannot be called `Stop`, a parameter cannot be called
  `attribute`, and array parameters must be ByRef.

**Compile error "procedure declaration does not match" (IDTExtensibility2_OnConnection)**
-> `Connect.cls` already uses the type library's native declarations `ext_ConnectMode` /
`ext_DisconnectMode` (from Add-In Designer). If a particular environment ships a different
type-library version, switch them to `vbext_ConnectMode` / `vbext_DisconnectMode` (from VB6
Extensibility).

**Loading reports a missing MSWINSCK.OCX**
-> Find `MSWINSCK.OCX` in the VB6 installation source, put it in `%SystemRoot%\SysWOW64` and run
`"%SystemRoot%\SysWOW64\regsvr32.exe" MSWINSCK.OCX` (do not hardcode `C:\Windows`; the reason
is above).

**The bridge reports "VB6 MCP server unreachable"**
-> Make sure the VB6 IDE is open and the add-in is loaded (Add-Ins -> VB6 MCP Status... shows the
port), and that the port matches `%TEMP%\vb6mcp_port.txt`. If a firewall blocks loopback traffic,
allow python through.

**Nothing changes in the IDE after writing a file**
-> Expected: see section 5. Run File -> Reload Project.

**Can I use this without a VB6 IDE running?**
-> No: the add-in depends on the VBIDE object model, so it must run with the IDE. The usual
practice is to leave VB6 running in the background.

## 9. Protocol notes

- Implements the MCP core: `initialize` / `ping` / `tools/list` / `tools/call` / notifications.
- Protocol version **negotiation**: the server understands 2025-11-25 and others, and echoes
  whatever version the client requests in `initialize` (the newest one when none is given; see
  `McpServer.NegotiateProtocol`).
- Transports: `POST /rpc`, `POST /mcp` (streamable HTTP) and `POST /message`; when the request
  carries `Accept: text/event-stream`, a single message is returned as an SSE frame. There is
  **no** `GET /sse` long-lived subscription, so server-initiated notifications are not used -
  clients pull `tools/list` when they need it.
- Implements `resources/list|read` (`vb6://project` plus `vb6://component/<name>` per component)
  and `prompts/list|get` (`edit-checklist` / `release-checklist`); unknown uris/prompts return
  `-32602`.
- JSON-RPC notifications return an empty HTTP 202; tool errors follow the MCP spec and come back
  as content with `isError: true` rather than as a JSON-RPC error.

**`open_project` (solved and fixed on 2026-09-21)**

- `VBProjects.AddFromFile` is declared **`AddFromFile(PathName, Exclusive)`** (see section 3 of
  the internal VBIDE object-model notes). The earlier claim that "this VB6 API only accepts a
  file name" was wrong; in practice **both call forms raise `Type mismatch (13)` and the project
  really is added anyway**.
- Success is therefore judged by the IDE's project list, and the return value is advisory: on
  success the response carries `apiError: "0xD Type mismatch"` plus a note saying this is a false
  alarm and the open actually succeeded.
- `exclusive: true` requires the current project to be closable. With unsaved changes VB6 raises
  60082 (project already loaded), so the tool now **saves first (the response carries
  `savedBeforeClose: true`) and then closes**; switching succeeds after `closedProjects: 1`.
- Step-by-step log: `%TEMP%\vb6mcp_openproject.log`; method-level begin markers:
  `%TEMP%\vb6mcp_audit.log`.
- Note: the earlier report that "the MCP endpoint goes away after the call" was **not** related to
  `open_project` - that was a **VB6 process crash**, triggered by a `.vbp` referencing component
  files that no longer exist on disk. See section 10 of the internal known-limitations document.

---

## What is `modMcpTrace` for, and is it required?

**1. What it is for**: a tiny standard module holding just `McpTrace` / `McpTraceOn` (plus a switch variable). It exists only to serve `inject_trace` - the injected tracing calls land on it and it writes one JSON line per event to the trace file.

**2. Is it required?**: **No - VB6MCP itself does not need it.** The add-in (editing, checking, compiling, running, breakpoints, dialogs, forms, snapshots, packaging) is independent of it, and **a project that never used `inject_trace` and keeps no injected calls should not have this module at all**.

**3. When is it needed?**: only when you **used `inject_trace` and kept the injected calls**, because those calls resolve to `McpTrace` / `McpTraceOn`, which live in this module. *If you keep the injected calls and build/ship an EXE or DLL, the module must be compiled into it* (the receipt's `shipNote` says the same).

**4. How MCP adds it**: when the target project has no `modMcpTrace`, `inject_trace` **generates and attaches it itself, choosing the comment language from the IDE** (the module source is embedded in VB6MCP.dll, so **there is nothing to import by hand**); it then re-reads the component list to prove it landed, and **skips the step entirely when the module already exists** (idempotent). The receipt carries `moduleAutoAttached` / `moduleAttachAction` / `moduleAttachMs`, plus `moduleAttachError` when the attach fails (a failure is always reported, never silently downgraded).

- **Which comment language (UI first)**: the **IDE's own UI language** decides first (main-window menu captions File/Edit/View/Help, else the mode marker `[design]`/`[run]` in the title); if that cannot be read it falls back to the VB6.EXE version resource (`\VarFileInfo\Translation`; LANG_CHINESE => Chinese); if neither can be read the answer is **English** (readable on every code page, while Chinese becomes mojibake on a non-Chinese one). The receipt reports all three signals side by side - `ideLanguage` (version resource), `ideMenuLanguage` (UI) and `commentLanguage` (the final choice) - each with its own `*Source` explanation.

**5. When its *reference* is added or removed**:

- **Added**: on the first injection when the project does not contain it (see 4).
- **Removed**: **the new version never removes it automatically** (`TASK-P1-149`). `remove_trace` deletes only **that id's injected block** and always keeps `modMcpTrace`; the receipt states `moduleRemoved:false` / `moduleLeftInPlace:true` / `moduleRemoveReason:"kept_by_design"` (or `module_absent` when the project has no such component), and `moduleRemoveNote` carries the two reasons below plus the manual-removal steps.
- **The old behaviour is withdrawn (and why)**: `TASK-P1-129` used to have `remove_trace` take an auto-generated `modMcpTrace` out of the project as well (`moduleRemoveReason:"removed_autogen_module"`). That branch has been **deleted**, because (a) you may have added your own code to an auto-generated module and deleting the module would have destroyed that code, and (b) other `@mcp-trace` blocks may still exist elsewhere, so with the module gone those calls could no longer find `McpTrace`/`McpTraceOn` and **the project would stop compiling**.
- **Removing it by hand**: first confirm that **no** `@mcp-trace` block remains anywhere in the project, then call `remove_component{name:"modMcpTrace"}` yourself. Note that `VBComponents.Remove` only removes the component from the PROJECT - it **does not delete the .bas file on disk** (the receipt reports the bound path as `moduleFileOnDisk`).
- **Where the log goes**: tracing is enabled by environment variables of the VB6 process - `MCP_TRACE_ON=1` or `MCP_TRACE_FILE=<path>`; the default sink is `%TEMP%\mcp_trace.jsonl`. A write failure is recorded in `MCP_TRACE_ERR_FILE` (default `%TEMP%\mcp_trace_err.log`) instead of being swallowed.
- **No other tool needs it**: editing, checking, compiling, running, breakpoints, dialogs, forms, snapshots and packaging are all independent of it.

## Compile artifacts get LAA (large-address-aware) by default - `make_project` / `build`

After a successful compile, `make_project` and `build` set `IMAGE_FILE_LARGE_ADDRESS_AWARE` (`0x0020`) in the artifact's PE header themselves - the equivalent of EditBin `/LARGEADDRESSAWARE`, with **no external `EditBin.exe`** required (a user machine may not have it). The switch is the `laa` argument (both tools accept it) and it is **ON by default**.

- **Who decides (highest precedence first)**:
  1. the per-call **`laa`** argument (`laa:true|false`) - pass it and it decides (`laaSource:"param"`);
  2. the project **`.vbp` `[VB6MCP]` named section** (`LAA=1`, optional `LAANote=...`; key names are case-insensitive and `1/true/yes/on` as well as `0/false/no/off` are accepted) - `laaSource:"vb-project"`;
  3. the **`.vb6mcp.json` sidecar** next to the project (`{"laa": true|false, "laaNote": "..."}`) - `laaSource:"project-file"`;
  4. the **`VB6MCP_LAA` environment variable** - `laaSource:"env"`;
  5. none of them => the **global default (ON)** - `laaSource:"global-default"`.
- **Receipt**: `laaSource` / `laaProjectFile` (the sidecar path) / `laaBefore` / `laaAfter` / `laaSet`, plus `laaSourceNote` and a `laa{}` detail object. An artifact that already carries the flag is **idempotent** (`laaSet:false`, `laaAlready:true`, not a single byte written).
- **A bad setting is never silently swallowed**: an illegal `LAA=` value in `[VB6MCP]`, or a sidecar that is not a JSON object, has no `laa` key, or carries a non-boolean `laa` => the LAA step is **skipped** and `laaConfigError` is returned (with `laaSource` `vb-project-error` / `project-file-error`); it does **not** quietly fall back to the default.
- **Two guards (both refuse to write and say why)**:
  - `signed_image`: the artifact carries an Authenticode certificate table (`IMAGE_DIRECTORY_ENTRY_SECURITY` non-empty) => the rewrite is **refused**, because it would invalidate that signature. The right order is **flag first, sign last** (signing must be the final step); if you really want a non-LAA artifact, pass `laa:false` explicitly.
  - `own_addin_dll`: when the target artifact is this add-in's own `VB6MCP.dll` the flag is deliberately **not** written (an ActiveX DLL's LAA is decided by the host process address space, so it is meaningless on the DLL itself).
- If the write or the read-back verification fails, the original characteristics word is put back (rollback) and the artifact is left unchanged; `laaSet:true` (`laaBytesPatched:2`) is only reported once the re-read confirms the new value.
- `laaMachineEffective`: the flag only actually enlarges a 32-bit process's user address space on 64-bit Windows (or 32-bit Windows booted with `/3GB`); otherwise the receipt says so instead of implying it always works.
- **Operational note (important)**: write the `[VB6MCP]` section into the `.vbp` while the project is **closed in the IDE**. A section **appended from outside while the project is open** is not in the IDE's in-memory project model, so VB6's **next project save overwrites it**; a section that was **already there when the project was loaded** is re-emitted verbatim on every save. If you need to change it at any time, use the `.vb6mcp.json` sidecar instead (VB6 only rewrites the `.vbp`, never the sidecar).

## Compiled but nothing produced? Read `unbuiltReasonUnknown`, then use `diagnose:true`

- **The in-IDE route (default `route:"ide"`) has no compiler subprocess**: `VBProject.MakeCompiledFile` reports only success or failure, so this route **structurally** cannot write an `/out` log => the receipt carries `logAvailable:false` with `logUnavailableReason` / `diagnosisUnavailableReason`, and `exitCode:0` only means "no compiler subprocess ran", **not** "the compiler found no errors".
- **`compiled:true` is not evidence**: every receipt carries `compiled_means`, `artifactChanged`, `verdict` and `output{exists,size,mtime,sha256}`. When `compiled:true` meets `artifactChanged:false`, the receipt **automatically** sets `unbuiltReasonUnknown:true` and states plainly that the artifact was **not produced and the reason is unknown - a compile failure is possible**; it is never dressed up as a success.
- **`diagnose:true` (it also runs automatically)**: it copies the `.vbp` and its component files into a temp directory and runs an isolated `VB6.EXE /make /out <log>` over the **copy**, returning the **real log plus `firstError{file,line,text}`**, marked `firstErrorSource:"isolated_copy_make"`; the detail lives in `diagnosis{}` (`route:"isolated_copy_make"`, `tmpDir`, `copiedFiles`, `missingFiles`, `errorCount`, `artifactExists`, ...). It runs **automatically** when the in-IDE Make reports `compiled=true` but `artifactChanged=false`, because that verdict is otherwise indistinguishable from a compile failure.
  - The copy is the state **on disk**, so **unsaved pane edits are not in it**.
  - Known edge: while another VB6 instance holds the same `.vbp`, the isolated `/make` **exits 0 and writes no log** - close that instance, or compile a copy of the project, and retry.
- On a compile failure the add-in also tries to read VB6's own compile-error box (`firstErrorSource:"ide_compile_error_dialog"`; `errorCount:1` means "one box was seen", not "exactly one error exists"); if it cannot read the box it says so.
- `project_meta{action:"set", cond_comp:""}` - **clearing the macros is accepted** (a production build wants the test macros gone): **presence** of the argument selects the write, so an empty string is not "no field supplied".

## The code-pane wheel now works in break mode too

- The add-in's own code-pane wheel (**off by default**, switchable in the MCP settings form) scrolls the pane the message lands on, **in break mode as well**.
- Why it used to fail in break: `VBE.CodePanes` can come back **empty** while the debugger is stopped (the pane the debugger opened need never have existed in design mode), so every `VbaWindow` was classified `not_a_code_pane` and never subclassed => the wheel message never reached the add-in. Now, once a code pane has been positively identified at least once, the add-in **still hooks when the table is empty**; where the last table came from is reported honestly in `captionsSource` (`all` / `active_fallback` / `empty`).
- **Failure reasons are separated honestly** (no single vague failure any more):
  - `lastTopStatus:"pane_not_found"` - the caption did not resolve to a CodePane => the wheel is forwarded untouched and nothing scrolls;
  - `lastTopStatus:"pane_not_writable_in_break"` - the pane **was** resolved but the IDE refused the `TopLine` read/write (measured in break mode) => the receipt carries the **real VB error** in `lastTopError` (e.g. "VB error N (...)" and the scroll is **queued and retried by the timer outside the window procedure** (`pendingScroll` / `pendingScrollResult`).
- **Observability**: `lastResolveVia` (`caption` / `component` / `active_name` / `none` - which channel found the pane), `lastTopStatus`, `lastTopError`, `captionsSource` / `captionsCount`, plus the existing `lastTopLineBefore/After`, `handled` / `forwarded`, `msgSeen`, ... (all inside the `wheel{}` object of `ide_state` and `mcp_info`).
- **Immediate / Watch / Locals windows cannot be mis-mapped**: they share the window class with a code pane, but the only fallback that maps a window onto the ACTIVE pane requires the window title to actually **contain** that pane's component name, so these three non-code windows are structurally unable to be treated as (or scroll) a code pane - they only show up in `skipped[]`.
- `lastTopLineBefore/After` are VBIDE `CodePane.TopLine` (1-based); **`-1` means "no value", not line 0 and not "it did not move"** - the cause is in `lastTopStatus` / `lastTopError`. For the current line read `paneTopLineNow` (the active pane's `TopLine`).

## License

**Apache License 2.0** (see LICENSE). You may use, modify, distribute and commercialise it
freely, including in closed-source derivatives; you only have to keep the copyright and licence
notices and state what you changed.

**Platform prerequisite**: this project is a VB6 add-in, so **both using and building it require
your own licensed VB6 installation**; this repository distributes no Microsoft files (see
THIRD_PARTY_NOTICES.md). The project is not affiliated with Microsoft.

## About the documents referenced in code comments

The source and its comments mention files such as internal reverse-engineering records, an
internal handover note, internal plan documents, an internal known-limitations/FAQ, internal
VBIDE object-model notes, an internal type-library dump, an internal command-bar inventory and
an internal forbidden-routes note. They are the maintainer's internal design and
reverse-engineering records and are **not shipped with the source** (the repository publishes
code plus the minimal set needed to build and use it). Those references do not affect
compilation or runtime; if you genuinely need the content of one of them, open an issue
explaining what for.

`docs/tools_schema.json` (the tool-face contract) and `tools/schema_baseline.json` (the frozen
baseline) **are** included, so the gate scripts under `tools/` run out of the box.

## Contributing

See CONTRIBUTING.md (English; Chinese: CONTRIBUTING.zh-CN.md) and see section 7 above for the known gaps.

**Localization comes first**: this project is developed and verified on a **Chinese VB6**, and
the known coupling points are enumerable (see section 7 of this file). If you run an English or other
language build, even a report of "works / does not work" is a valuable contribution.
