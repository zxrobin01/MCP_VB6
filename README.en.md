# VB6MCP — turn the VB6 IDE into an MCP server

> [中文](README.md) | **English**

> **Changelog**: [CHANGELOG.en.md](CHANGELOG.en.md)

VB6MCP is a **VB6 add-in**: once loaded into the VB6 IDE, the IDE itself becomes a
[MCP (Model Context Protocol)](https://modelcontextprotocol.io) server. AI coding assistants
such as Claude Desktop, Cursor or Codex can then read, search and modify the VB6 project that
is open in the IDE, and drive compilation and execution.

```
 MCP client (Claude Desktop / Cursor / ...)
        |  stdio (JSON-RPC, one message per line)
        v
 bridge/vb6mcp_hub.mjs      (Node, zero dependencies; MCP wire protocol + 10 ide_* / lesson tools)
        |  HTTP  http://127.0.0.1:8767  (bearer token required once a credential is configured)
        v
 bridge/vb6mcp_bridge.py    (Python 3, standard library only; the machine's ONLY TCP listener)
        |  named pipe (one per IDE instance)
        v
 VB6MCP.dll  <- runs inside the VB6.EXE process (no TCP listener of its own)
   |-- named-pipe server
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
server**. Load one add-in and the IDE exposes 110 tools: read/search/edit code, static
checking, anchor-based editing with safe transactions, compilation, run and single-step,
dialog handling, form and project metadata, snapshots and rollback, packaging. Ten more
**hub-face** tools (IDE instance management plus the lesson store) run outside the IDE
process, so a client sees 120 tools in total. **Because the tools live inside the IDE process**,
they can do what an out-of-process approach cannot: actually click that dialog, actually press
Run, actually read the local variables at a breakpoint.

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
|-- CodeTools.cls          the 110 MCP tools (full VBIDE object model + file system + syntax gate / edit safety / compile authority)
|-- JsonLib.bas            zero-dependency JSON parser/serializer
|-- Utf8.bas               UTF-8/ANSI/UTF-16 conversion (kernel32 APIs)
|-- Util.bas               logging, port file
|-- register.reg           registry entries that register the add-in with VB6 (import once)
|-- unregister.reg         remove the add-in registration
|-- build.bat              one-click build + regsvr32 registration + import the add-in registry file
|-- mcp_config.example.json    Claude Desktop configuration example
|-- bridge/vb6mcp_hub.mjs   MCP entry point (Node, stdio; ide_* and lesson tools)
|-- bridge/vb6mcp_bridge.py middle service: the machine's only TCP listener (default 8767), forwards to the named pipe
+-- tests/smoke_test.py     smoke test that talks to the HTTP endpoint directly
```

## Let an AI agent install it for you (recommended)

Not in the mood to click through the steps below? Copy this **whole block** into your AI coding assistant
(Claude Code / Cursor / Codex / any agent with a terminal). The only thing you have to supply is the
repository address: replace `<REPO_URL>` with this repository's URL (filling it in is better; leaving it
blank is fine too - the agent will ask you for it). The agent then fetches the source, builds, registers,
starts the IDE, **actually verifies the MCP server works**, wires up your client, and finally reports back
what it did, what it saw and what it could not do.

```text
You are my installation assistant. On one Windows machine, install VB6MCP from scratch, start it, and verify that it really works. VB6MCP is a VB6 IDE add-in (an ActiveX DLL that produces VB6MCP.dll): load it into the VB6 IDE and the IDE itself becomes an MCP server, so an AI client can read and modify the VB6 project I have open, compile it and run it.

Source repository: <REPO_URL>
If that address is still a placeholder (empty, or still literally <REPO_URL>), ask me for the repository address first. Do not guess, and do not go looking for a different repository with a similar name.

Deliverable: obtain the source -> build -> register -> start the IDE -> really verify the MCP server -> wire up the client -> run the smoke test -> report honestly.

Boundaries (read these before touching anything):
- Do this one job only. Do not also change my other projects, other add-ins or other registry entries.
- No destructive or irreversible actions: do not delete or overwrite files and directories you did not create, do not modify any VB6 project I already have, do not uninstall or disable my other add-ins, do not "clean up" processes or registry keys.
- Do not kill processes. Never terminate VB6.EXE or anything else; if the IDE has to be closed, ask me first (I may have unsaved work).
- Touch only the port you detect. In this generation the **middle service** is the only TCP listener: default 127.0.0.1:8767 (change it with --port or VB6MCP_BRIDGE_PORT); the add-in no longer listens on TCP at all. The real port is the port field returned by http://127.0.0.1:<port>/health; python bridge/vb6mcp_bridge.py --print-config also prints the effective port, transport and authEnabled. Do not scan, occupy or change any other port, and never expose this service beyond the loopback address (a non-loopback bind is refused outright and the process exits with code 6 - there is no override).
- Stop and ask me before anything that could break my existing VB6 setup (machine-wide re-registration of OCX files or type libraries, changes under HKCU\Software\VB6MCP, changes to VB6's references, ...): tell me what you would change, why, and how to undo it.
- Keep every artifact inside a working directory you create. Every conclusion must rest on a reproducible command and its raw output; "it should have worked" is not a conclusion. If something needs me to click it, just tell me.

Steps (in order; give evidence for each):

1) Check the prerequisites. Confirm this is Windows (32-bit or 64-bit), then check:
   - Is a VB6 IDE (SP6) installed? Look for VB6.EXE (the standard install folder is VB98; it differs between machines).
   - Is MSWINSCK.OCX (the Winsock control) available/registered, and is MSADDNDR.DLL (the Add-In Designer type library) available?
   - The 32-bit registration tool: on 64-bit Windows use %SystemRoot%\SysWOW64\regsvr32.exe. Do not use a bare regsvr32 and do not hardcode the Windows folder path (use the %SystemRoot% variable).
   - If VB6 is absent: stop immediately and tell me the truth - VB6MCP is an add-in for the VB6 IDE, so without VB6 it can neither be built nor run, and there is no workaround; the repository redistributes no Microsoft files, so my own licensed VB6 installation is required. Do not try to fake VB6 with another toolchain, and do not continue with the later steps on a machine that has no VB6.
   - If MSWINSCK.OCX or MSADDNDR.DLL is missing, follow the repository README's FAQ; ask me before writing into system directories.

2) Get the source. git clone <REPO_URL> into the working directory we agree on (or download and unpack the repository ZIP). Confirm the tree contains MCP_VB6.vbp, build.bat, register.reg, bridge/vb6mcp_bridge.py, bridge/vb6mcp_hub.mjs and tests/smoke_test.py.

3) Build. Prefer running build.bat directly (it registers MSADDNDR.DLL first, then compiles VB6MCP.dll). If VB6 is not in the default location, edit the VB6 variable inside build.bat rather than hardcoding a path anywhere else.
   Note: do not rebuild in place while the IDE already has this add-in DLL loaded - the file lock makes the build fail. Close the IDE first (ask me whether anything is unsaved), or build from another VB6 instance.
   You may also open MCP_VB6.vbp in the IDE and compile there, but do not make the VB6MCP project itself the active project and use "Make" (same file lock).
   If the build fails, quote the compiler's original error text (for example whether the project references VB6 Extensibility and Add-In Designer are ticked - see the README FAQ). Do not guess-edit source.

4) Register and start the IDE. Register VB6MCP.dll with the 32-bit regsvr32 you confirmed in step 1, import register.reg, then check the registration yourself:
   reg query "HKCU\Software\Microsoft\Visual Basic\6.0\Addins\VB6MCP.Connect" /v LoadBehavior
   It should read 0x3 (=3, load automatically) and FriendlyName should be VB6 MCP Server. build.bat normally does the registration and the import for you, but verify it anyway.
   Then start the VB6 IDE (if it is already open but the add-in is not loaded, ask me whether to restart it - do not kill the process yourself) and confirm "VB6 MCP Server" is ticked under Add-Ins -> Add-In Manager (LoadBehavior=3 usually loads it automatically).

5) Really verify the server (the key step: "the process is running" is not evidence):
   - GET http://127.0.0.1:8767/health until it returns {"ok":true,...}; trust the port field in that response and compare it with the port printed by python bridge/vb6mcp_bridge.py --print-config. Call that port <port> below. **If a token is configured, this request needs `Authorization: Bearer <token>` too** - otherwise it returns 401.
   - Make a real MCP call to it: POST http://127.0.0.1:<port>/rpc with Content-Type: application/json
     {"jsonrpc":"2.0","id":1,"method":"initialize","params":{"protocolVersion":"2025-11-25","capabilities":{},"clientInfo":{"name":"install-check","version":"1.0"}}}
     then {"jsonrpc":"2.0","id":2,"method":"tools/list","params":{}}
   - Count the tools in tools/list and compare against the repository's own authoritative catalog: the number recorded in docs/TOOL_CATALOG.md, and the length of the tools array in tools/schema_baseline.json. Do not hardcode any number - use the numbers from the repository you just cloned. You may also call the mcp_info tool and read the contract and toolCount the server reports about itself.
   - If all three agree, pass the step and record the number. If they disagree, report the exact difference and which tools are extra or missing; do not edit the repository to make it match and do not fudge the number. If tools/list cannot be obtained at all, paste the raw response and the HTTP status code.

6) Wire up the client (this generation recommends starting the hub over stdio): point command at node, args at <repo>/bridge/vb6mcp_hub.mjs, and give it two environment variables - VB6MCP_BRIDGE_URL=http://127.0.0.1:<middle-service port> (the hub's source default is http://127.0.0.1:8767; give it explicitly whenever the service runs on another port) and a token (VB6MCP_TOKEN or VB6MCP_TOKEN_FILE, only when the server has a credential configured). Claude Desktop: %APPDATA%\Claude\claude_desktop_config.json; Cursor: its MCP configuration (streamable HTTP against http://127.0.0.1:<port>/mcp also works, with an Authorization: Bearer <token> header when a token is configured). Replace placeholder paths with the real paths on this machine; in JSON use / or write \\. Back up the configuration file before changing it.

7) Run the smoke test: python tests/smoke_test.py --port <port> (the middle service's port). It requires the VB6 IDE to be open with the add-in loaded; it is not an offline self-check. Note that this script sends **no token**, so it only works while the server has no credential configured; with a token configured use python bridge/vb6mcp_bridge.py --print-config to read the state, and --selftest for the in-process self-check.

8) Report back (this is the most important step). Give me:
   - the commands you actually ran (verbatim) and what you actually observed (key output, the health response, the tool count, the smoke-test result);
   - your verdict per step (pass/fail/skipped), with the raw error for anything that failed;
   - plainly, what you could not do (for example "VB6 is not installed, so I could not build", "this step needs me to click a dialog", "I lacked permission here").
   Never invent a success to make the report look better; if it did not work, say it did not work.

One more time: do not kill processes; do not touch ports, services, registry keys or project files you were not asked to touch; ask before anything that could damage my existing setup. I want a real success, or an honest statement of where it is stuck.
```

Prefer to do it by hand → see "1. Building", "2. Connection model, ports and multiple instances" and "3. Connecting an MCP
client" below. Either route works; you only need one.

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

## 2. Connection model, ports and multiple instances

**This generation: one listening port + one named pipe per IDE instance.**

- Clients no longer connect to the add-in directly. The **middle service** `bridge/vb6mcp_bridge.py`
  owns **the single TCP port**: default `http://127.0.0.1:8767` (change it with `--port` or the
  `VB6MCP_BRIDGE_PORT` environment variable; `VB6MCP_BRIDGE_HOST` accepts loopback addresses only).
- **Non-loopback binding is impossible**: binding to a non-loopback address is refused and the process
  exits with **code 6** - there is **no override switch**.
- MCP clients come in through the **hub** `bridge/vb6mcp_hub.mjs` (Node, zero dependencies, stdio): the
  hub serves the 10 `ide_*` and lesson tools and forwards the add-in face (110 tools) to the middle
  service; it finds the service through `VB6MCP_BRIDGE_URL` (source default `http://127.0.0.1:8767`).
- The add-in accepts only a **named pipe** (one per IDE instance; a single pipe **does not allow
  concurrent connections**) => clients should multiplex by request id **on one connection** or
  serialize their calls; a busy pipe answers `pipeBusy` instead of "no IDE".
- **Transport selection** `VB6MCP_TRANSPORT=auto|pipe|tcp` (default `auto`: prefer the pipe and fall back
  to the add-in's HTTP **visibly**; `pipe` is strict, so a missing pipe is an explicit error; `tcp` uses
  only the add-in's HTTP, for rollback and transition).
- **Custom port**: `--port N`, `VB6MCP_BRIDGE_PORT` (or `VB6MCP_PORT` on the same machine). With several
  VB6 instances open there is **still only one service port**; instances are addressed with `_target`
  (see section 3).
- Status: IDE menu **Add-Ins -> VB6 MCP Status...**, or `GET http://127.0.0.1:<port>/health` (that request
  needs the token too once a credential is configured); `python bridge/vb6mcp_bridge.py --print-config`
  prints the effective port / transport / `authEnabled`.
- Log: the middle service's log path is `logPath` in `--print-config`; the add-in's own log stays `%TEMP%\vb6mcp.log`.

## 3. Connecting an MCP client

### Connection and authentication (this generation)

- **Recommended**: start the hub over **stdio** - `command: node`, `args: ["<repo>/bridge/vb6mcp_hub.mjs"]`,
  with `VB6MCP_BRIDGE_URL=http://127.0.0.1:8767` in the environment (give it explicitly whenever the
  middle service runs on another port).
- **Direct HTTP** (streamable HTTP): `POST http://127.0.0.1:<port>/mcp` (`/rpc` is the same face).
- **Authentication**: once the server has a credential configured (`VB6MCP_TOKEN`, or `VB6MCP_TOKEN_FILE`
  pointing at an owner-only file), **every HTTP surface** (`/rpc`, `/mcp`, `/health`, ...) requires a
  bearer token: `Authorization: Bearer <token>`, or `X-VB6MCP-Token: <token>`. **A credential inside the
  URL is rejected** (401 `token_in_url_rejected`). **With no credential configured the service still runs**,
  but `/health` reports `auth.enabled=false` - do not run it that way on a multi-user machine.
- **Addressing**: a call names its target project with `_target`, canonically `[endpoint:]projectPath`
  (for example `C:/work/MyProj.vbp`); omitting `endpoint` means local. **A call without `_target` is never
  sent to a remote.**
- The old generation's direct port (8766) **no longer exists** - a client still pointing at it fails to connect.

### Claude Desktop

Let Claude Desktop start the hub over **stdio** (replace the path with your own):

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

When the server has a token configured, add `VB6MCP_TOKEN` (or `VB6MCP_TOKEN_FILE`) to `env`.

> **Replace every path with one from your own machine**: the examples (including `mcp_config.example.json`) use **placeholder paths** - both `command` (node/python) and the `args` path to the repository script (for example `bridge/vb6mcp_hub.mjs`) must point at your own installation. In JSON prefer forward slashes `/`; a Windows backslash must be written as `\\`.

### Cursor (streamable HTTP)

Cursor can also talk to the middle service's streamable-HTTP face directly:

```
type: streamable HTTP
url: http://127.0.0.1:8767/mcp
header: Authorization: Bearer <token>    # only when a token is configured
```

(Starting the hub over stdio works too - see above.)

### Other stdio clients

The same as Claude Desktop: `command: node`, `args: ["<repo>/bridge/vb6mcp_hub.mjs"]`, plus `VB6MCP_BRIDGE_URL` and `VB6MCP_TOKEN` / `VB6MCP_TOKEN_FILE` as needed.

To check the whole chain: `python tests/smoke_test.py --port <middle-service port>` (that script sends **no token**, so it only works while the server has no credential configured; and it **needs the VB6 IDE open with the add-in loaded** - this is not an offline self-check).

## 4. Tool list (add-in face 110 + hub face 10 = 120 tools)

> The authoritative list is `docs/TOOL_CATALOG.md` (exported from a running add-in via
> `tools/list` and merged with the hub's own tools); the offline baseline is
> `tools/schema_baseline.json`, and `tools/check_live_tools.py docs/tools_schema.json`
> compares the shipped snapshot against that baseline.
>
> Tools come from **two faces**: the **add-in face, 110 tools** (`docs/tools_schema.json`, inside
> the VB6 IDE process) and the **hub face, 10 tools** (`bridge/hub_tools.json`, outside it). Below,
> all **120** tools are grouped: 16 add-in categories plus the hub section (17); each tool
> appears exactly once. The key-parameter column is taken straight from each tool's own
> `inputSchema` (required parameters first), never typed from memory. The five hub lesson tools
> **work with no IDE online**.

### 1. Project & workspace

| Tool | What it does | Key parameters |
|------|--------------|----------------|
| `get_project_info` | Read project overview: type, IDE state, startup object, build path, components, references | - |
| `set_project_options` | Set project options: description, startup object, build path, read-only mode | `description`, `startup_object`, `build_file_name` |
| `open_project` | Open a .vbp project, optionally replacing the workspace | `path`, `exclusive` |
| `add_project` | Create a new project, or one from a template | `type`, `template_path` |
| `remove_project` | Close a project from the IDE workspace | `name` |
| `save_project` | Save the project and every dirty component | `path` |
| `vbg_group` | Read or edit a .vbg project group: list, add, remove | `action`, `path`, `project` |
| `workspace_status` | Show unsaved components beside the git working tree | `repo` |

### 2. Component lifecycle

| Tool | What it does | Key parameters |
|------|--------------|----------------|
| `get_component_list` | List components with procedures, line ranges and dirty state | - |
| `add_component` | Add a component: module, class, form and more | `type`, `name`, `code` |
| `remove_component` | Remove a component from the project, cannot be undone | `name` |
| `import_component` | Import a source file or template as a project component | `path`, `mode` |
| `export_component` | Export a component's code or whole designer to a file | `name`, `path` |
| `rename_component` | Rename a component inside the project | `name`, `new_name` |
| `reload_component` | Reload a component from the file it is bound to | `name` |
| `activate_component` | Open a component's code pane or designer window | `name`, `window` |

### 3. Reading code, search & caret

| Tool | What it does | Key parameters |
|------|--------------|----------------|
| `get_component_code` | Read component code from IDE memory, optionally one proc or symbol | `name`, `include_line_numbers`, `coords` |
| `read_source` | Read a source file from disk with encoding detection and sha256 | `path` |
| `read_file` | Read any text file, UTF-8, UTF-16 or ANSI auto-detected | `path`, `max_lines` |
| `find_in_component` | Find text in a component: whole word, case, pattern | `component`, `target`, `start_line` |
| `get_selection` | Read the IDE's selected component, active window and caret | - |
| `get_codepane` | Read code-pane selection, top line and visible line count | `component` |
| `set_codepane` | Set code-pane selection and scroll, optionally activating it | `component`, `selection`, `top_line`, `activate` |
| `current_position` | Read where the engine stopped: module, procedure, line, text | - |

### 4. Editing with safety transactions

| Tool | What it does | Key parameters |
|------|--------------|----------------|
| `replace_code` | Line-level edit: replace, insert, append, delete or set module | `name`, `operation`, `start_line` |
| `add_code` | Append code to a component from a string or a file | `component`, `code`, `file` |
| `create_event_proc` | Create an event-procedure stub such as Command1_Click | `component`, `object`, `event` |
| `write_file` | Write a text file with optional encoding, directory creation and backup | `path`, `content`, `encoding` |
| `write_source` | Write a VB source file with explicit encoding and byte-level sha256 | `path`, `text`, `encoding` |
| `edit_by_anchor` | Replace one uniquely matched line by content anchor, not line numbers | `name`, `anchor`, `new_text` |
| `apply_edits` | Apply a batch of anchor edits atomically, all or nothing | `name`, `edits`, `expected_sha256` |
| `preview_edit` | Preview anchor edits as a unified diff, writing nothing | `name`, `edits` |
| `replace_in_files` | Cross-component literal or regex replace, dry_run by default | `path`, `query`, `replacement` |
| `rename_symbol` | Rename an identifier lexically, skipping strings and comments | `path`, `old_name`, `new_name` |

### 5. Syntax gating & project checks

| Tool | What it does | Key parameters |
|------|--------------|----------------|
| `validate_code` | Validate code before writing, with stable-key diagnostics | `code`, `name`, `kind` |
| `check_component` | Lint the current in-IDE text of one component | `name` |
| `verify` | Aggregate verdict: IDE equals disk plus lint counts | `name`, `disk_path`, `baseline` |
| `verify_scope` | Check many components for IDE/disk equality and diagnostics | `components`, `path` |
| `format_code` | Re-indent by block structure, proving only leading whitespace changed | `path`, `indent`, `apply` |
| `check_references` | List project references and flag ones whose file is missing | `include_resolved` |

### 6. Dependencies, purity & public API

| Tool | What it does | Key parameters |
|------|--------------|----------------|
| `dependency_graph` | Component and procedure dependency graph: edges, cycles, purity; the main deadline is a **deterministic character quota** (`VB6MCP_CHAR_QUOTA`, default 900000) and the wall clock is only the last resort (`VB6MCP_BUDGET_MS`, default 30000). The receipt carries completeness fields such as `partial` / `*StageCompleted` / `budgetExhausted` / `timingTruncated` / `procEdgesFound` - **the default setting delivers a deterministic partial result** | - |
| `purity_check` | Classify whether a component can enter an offline test host | `name` |
| `export_public_api` | Freeze a component's public API into a comparable spec; omit `name` for a bounded whole-project export | `name`, `membersOnly` |
| `check_public_api` | Compare against a frozen spec: removed, changed, added members; omit `name` for a bounded whole-project check | `baseline`, `name`, `membersOnly` |
| `generate_stub` | Generate a compilable stub: declarations kept, bodies emptied | `name` |

### 7. Build & delivery artifacts

| Tool | What it does | Key parameters |
|------|--------------|----------------|
| `make_project` | Compile and return the authoritative artifact verdict; diagnose gives a real log; the build state is queryable (`buildState`) | `output_path`, `laa`, `diagnose`, `wait_for_build` |
| `build` | Compile via an isolated VB6.EXE /make and return the log | `project`, `defines`, `laa` |
| `package_app` | Assemble a delivery folder: exe, manifest, dependencies, installer | `path`, `exe_path`, `out_dir` |
| `app_manifest` | Generate or read the app manifest beside an exe or dll | `action`, `path`, `level` |

### 8. Run, debug & breakpoints

| Tool | What it does | Key parameters |
|------|--------------|----------------|
| `control_execution` | Control execution: run, full compile, break, step, end, restart | `mode` |
| `execution_state` | Read a process's windows, responsiveness and modal-dialog count | `pid` |
| `debug_commands` | Catalog of known debug commands, a fixed table | - |
| `set_breakpoint` | Inject a Debug.Assert False breakpoint, omitted in compiled builds | `name`, `anchor`, `id` |
| `clear_breakpoints` | Remove an injected breakpoint block by id | `name`, `id`, `expected_sha256` |
| `list_breakpoints` | List injected breakpoints with line and enabled state | - |
| `native_breakpoint` | Toggle a real IDE breakpoint through the Debug menu | `action`, `component`, `line` |
| `fault_inject` | Inject a guarded fault at a procedure start: raise, return, delay | `name`, `procedure`, `mode`, `id` |
| `clear_faults` | Remove a previously injected fault block | `name`, `id`, `expected_sha256` |

### 9. Tracing, call stack & in-process execution

| Tool | What it does | Key parameters |
|------|--------------|----------------|
| `inject_trace` | Insert a runtime-guarded McpTrace call, generating the module if absent | `name`, `procedure`, `id` |
| `remove_trace` | Remove an injected trace block, keeping the generated module | `name`, `id`, `expected_sha256` |
| `project_changes` | Poll the IDE change journal: component and control churn | `since`, `clear` |
| `get_callstack` | Read the engine's own call stack with per-frame procedure names | `max_frames` |
| `vba6_probe` | Probe the VBA6 debug engine in-process: mode, callstack, execute | `exec_probe`, `exec_line`, `run_break_probes` |
| `vba6_probe_plan` | Ordered probe plan with preconditions, read-only steps first | - |
| `immediate_eval` | Evaluate one expression in the immediate context and return its value | `expression`, `expression_file`, `timeout_ms` |
| `immediate_run` | Run statements in the immediate context and return what they printed | `code`, `code_file`, `timeout_ms` |

### 10. Windows, dialogs & UI automation

| Tool | What it does | Key parameters |
|------|--------------|----------------|
| `list_windows` | List IDE windows with caption, visibility, position and size | - |
| `control_window` | Control one IDE window: focus, close, show, hide, move, resize | `window`, `action` |
| `find_gui` | Enumerate this process's windows, filtered by keyword or class | `keyword`, `cls`, `showing` |
| `ui_text` | Read a window's accessible text via UIA, MSAA or WM_GETTEXT | `hwnd`, `cls`, `depth` |
| `read_active_dialogs` | List active dialogs and their buttons without clicking anything | `cls`, `include_children`, `max` |
| `close_active_dialogs` | Dismiss blocking dialogs by clicking only whitelisted safe buttons | `hwnd`, `button`, `dry_run` |
| `send_window_message` | Send a window message to this process: click, key, raw number | `hwnd`, `action`, `message` |
| `command_bars` | List IDE menu bars or execute one command by name or id | `action`, `bar`, `control` |

### 11. Form designer & controls

| Tool | What it does | Key parameters |
|------|--------------|----------------|
| `form_designer` | Form designer: control inventory, properties, selection and clipboard | `form`, `action`, `control` |
| `add_control` | Add a control by ProgID after pre-validating it is registered | `form`, `prog_id`, `name` |
| `remove_control` | Remove a control from a form by name or index | `form`, `index`, `control` |
| `set_control_property` | Set a control's design-time property value | `form`, `property`, `value` |
| `list_component_parts` | List the project's registered ActiveX parts and resolved paths | - |

### 12. References, members & properties

| Tool | What it does | Key parameters |
|------|--------------|----------------|
| `add_reference` | Add a type-library reference by GUID or file path | `guid`, `path`, `mode` |
| `remove_reference` | Remove a project reference by name or GUID | `name` |
| `get_component_properties` | Read a component's design-time properties, the Properties window content | `name` |
| `set_component_property` | Set a component property, including multi-dimensional indexed ones | `name`, `property`, `value` |
| `list_members` | List a class component's public members with code locations | `component` |
| `set_member_attribute` | Change class-member attributes: description, Hidden, Browsable, Category | `component`, `member`, `attribute`, `value` |
| `property_bag` | Read or write named values in a project or component property bag | `target`, `key`, `name` |

### 13. Files on disk & project metadata

| Tool | What it does | Key parameters |
|------|--------------|----------------|
| `list_files` | List files under a directory by wildcard, recursively | `path`, `pattern`, `recursive` |
| `search_in_files` | Search text or regex across files, returning file and line | `query`, `path`, `pattern` |
| `project_meta` | Read or write .vbp version resources and build metadata | `action`, `path`, `major` |
| `git` | Read-only Git: status, diff --stat and commit log | `action`, `repo`, `file` |

### 14. Snapshots, rollback & test host

| Tool | What it does | Key parameters |
|------|--------------|----------------|
| `project_snapshot` | Project-level snapshot and rollback with a per-file manifest | `action`, `path`, `dir` |
| `undo_edit` | Component-level single-step undo of pane text only | `component`, `action`, `ts` |
| `test_host` | Generate an offline test host .vbp, or score PASS/FAIL results | `action`, `vbp_path`, `module_name` |
| `record_fixture` | Record a deterministic fixture as raw ANSI bytes | `path`, `text` |
| `compare_fixture` | Compare text against a fixture: same, added, removed, first diff | `path`, `text` |

### 15. Extension store

- **Extension store** `mcp_extensions/`: every `*.json` is a **declarative tool pack**
  (`name` / `tools[].inputSchema` / `tools[].handler`). An enabled pack's tools are merged into
  `tools/list` automatically and dispatched by `CodeTools.CallTool`; `state.json` records the
  disabled list (enabled by default). Handlers expose only a few safe primitives (`http_get`,
  `read_file`, `write_file`) and **never load VB6 code from disk**.

| Tool | What it does | Key parameters |
|------|--------------|----------------|
| `ext_list` | List extension packs with enabled state and tool count | - |
| `ext_enable` | Enable an extension pack, its tools join tools/list | `name` |
| `ext_disable` | Disable an extension pack, its tools leave tools/list | `name` |
| `ext_reload` | Rescan mcp_extensions and report pack and tool counts | - |
| `ext_install` | Install an extension pack JSON into mcp_extensions | `path` |
| `ext_uninstall` | Remove an extension pack, renaming its file aside | `name` |

### 16. Governance, self-check & IDE environment

| Tool | What it does | Key parameters |
|------|--------------|----------------|
| `mcp_info` | Read the server version stamp, contract, tool count and build hash; `build.sourceHash` / `sourceAlgo` are the source input set's fingerprint and algorithm id | - |
| `ide_state` | Read IDE mode plus in-flight, cancel and gate-release state | `cancel_inflight`, `release_inflight`, `release_min_held_ms` |
| `desktop_probe` | Measure the host's desktop and interactivity without stealing focus | - |
| `ide_quiet_mode` | Toggle session quiet mode: no focus stealing, advisory appended | `on`, `ttl_ms` |
| `todo_status` | Read-only snapshot of an optional in-repo task board | - |
| `audit_tail` | Return the most recent tool-call audit entries | `count` |
| `list_addins` | List registered add-ins and their connection state | - |
| `set_addin_state` | Load or unload an IDE add-in by ProgID | `prog_id`, `connect` |
| `quit_ide` | Quit the IDE after responding, unsaved changes are lost | - |

### 17. Hub face: IDE instance management & lesson store (10 tools, no IDE online required)

These 10 tools come from the MCP hub (`bridge/vb6mcp_hub.mjs`) and run **outside the IDE
process**: the client does not have to load the add-in first, and VB6 does not have to be
running. The five lesson tools only touch a local JSON store, so they **work with no IDE online**.

- **Lesson store** `mcp_lessons/`: `defaults.json` (shipped, do not edit) plus `user.json` (the user layer,
  which may override, disable or add entries by id). The merged result is published to clients through
  `initialize.instructions`, so an agent can read it before calling any tool; `add`/`delete`/`import` write
  `user.json` with a `.bak` backup.
- **IDE instance management**: `ide_list` lists every VB6 IDE the hub can address: local instances (including the conflict where two instances
  hold the same project), `ide_target` remembers the default target, and `ide_start` / `ide_stop` /
  `ide_restart` only ever act on **instances this hub started**. It also lists every configured **remote endpoint** (allow-list only, never discovered; unreachable ones are still listed with their probe status) and echoes the one-set identity (endpoint / targetKey / pathVia), the canonical real path beside the registry spelling, and the caller's alias spelling that resolves to the same identity.

| Tool | What it does | Key parameters |
|------|--------------|----------------|
| `ide_list` | List **every** VB6 IDE the hub can address: local instances plus each **configured remote endpoint** (allow-list only, never discovered; unreachable ones still listed with their probe status). Reports pid/port/project, the one-set identity (endpoint / targetKey / pathVia), the canonical real path beside the registry spelling, and echoes the caller's alias that resolves to the same identity; includes stale entries and same-identity conflicts | - |
| `ide_target` | Set or clear the default IDE target: pid, project or auto | `target` |
| `ide_start` | Start a VB6 IDE instance for a .vbp project | `project` |
| `ide_stop` | Stop an IDE instance that this hub started | `pid`, `force` |
| `ide_restart` | Restart a wedged own IDE instance and prove readiness | `pid`, `project` |
| `list_lessons` | List lesson entries: scenario, trap, do (hub face, no IDE online) | `includeDisabled` |
| `add_lesson` | Add or overwrite a user lesson, with backup (no IDE online) | `id`, `scenario`, `trap`, `do` |
| `delete_lesson` | Remove a user lesson, or disable a shipped one (no IDE online) | `id` |
| `export_lessons` | Export the merged lesson library as Markdown (no IDE online) | `includeDisabled` |
| `import_lessons_md` | Import lessons from Markdown, optionally parse-only (no IDE online) | `path`, `dryRun` |

> `get_codepane` / `set_codepane` take **no** `view` parameter: VB6 exposes `CodePaneView` as read-only
> and its command bars carry no view toggle, so the add-in cannot switch between full_module and procedure
> view; a `view` argument you pass is reported in `ignoredArgs`.

The full object-model inventory and gap analysis live in the maintainer's internal VBIDE
object-model notes (produced by parsing the VB6EXT.OLB 5.3 type-library binary; the raw data is
in the internal type-library dump).

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

**Can I type in the IDE on the same machine while an agent is driving it?**
-> **No.** This is a rule, not advice: while a tool drives the IDE it moves the caret, scrolls,
opens panes, and some actions bring their window to the foreground (`control_window{action:"activate"}`;
`native_breakpoint` first moves the caret to the target line). **Your keystrokes go to whatever window
has the foreground at that moment** - this project has seen a real incident where the user typed during
an agent operation, the keystrokes landed in the code pane, and a stray line of characters appeared in
the source (typed by the keystrokes, not written by a tool). To edit by hand on the same machine, stop
the agent first, or use another machine / another instance.

Write-class tools therefore carry guards (all of them visible in the receipts):

- **Post-write read-back assertions**: editing tools already carry `expected_sha256` / `sha256Before` /
  `sha256After`; the breakpoint and window tools that do **not** write text (`native_breakpoint` /
  `control_window` / `set_codepane`) now also hash the code-pane text before and after and require the
  two to be equal - a difference turns the receipt into `{ok:false, reason:"paneTextMismatch"}` with the
  line-level differences (`paneDiffLines`). When the pane cannot be read it is reported as
  `paneAssertion:"notApplicable"`, never as a pass.
- `set_member_attribute` reads the member attribute back after writing (`readback` / `applied`); note
  that this is the **in-memory** value - whether it lands on disk as an `Attribute` line is decided by
  VB6 when the component is saved.
- **Focus safety**: `native_breakpoint` / `control_window` receipts carry `foregroundBefore` /
  `foregroundAfter` / `foregroundTookFocus`, and any non-`activate` action gives the foreground back to
  the window that had it. The value is reported **live**: it is a valid non-zero handle in an interactive
  session and may read `0` on a locked or non-interactive desktop, in which case it is reported as-is -
  the tool never steals focus to obtain it, and send-input capability is never inferred from a
  window-station name.
- **Breakpoints have no third route** (stated in the receipt): VB6 has **no breakpoint API**, the Debug
  menu items act on the **caret**, and the other route (`set_breakpoint`) **writes text** marker blocks.
  `native_breakpoint` therefore **compensates**: it puts the caret back where it was (`cursorRestored`).
  `breakpointWithoutCursorMove:false` is the machine-readable mark of that fact.

## 9. Protocol notes

- Implements the MCP core: `initialize` / `ping` / `tools/list` / `tools/call` / notifications.
- Protocol version **negotiation**: the server understands 2025-11-25 and others, and echoes
  whatever version the client requests in `initialize` (the newest one when none is given; see
  `McpServer.NegotiateProtocol`).
- Transports: the middle service provides `POST /rpc` and `POST /mcp` (streamable HTTP); when the request carries `Accept: text/event-stream`, a single message is returned as an SSE frame. The old generation's `GET /sse` + `POST /message` long-lived subscription is **no longer implemented** (it answers 501), so server-initiated notifications are not used - clients pull `tools/list` when they need it.
- Authentication: once a credential is configured, **every HTTP surface** needs `Authorization: Bearer <token>` or `X-VB6MCP-Token`; a credential inside the URL is always 401 `token_in_url_rejected`.
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

**What is deliberately NOT provided (a refusal by design, not a defect)**: this add-in has **no**
channel that writes `.vbp`-level project properties into the IDE's in-memory project model, and
**no** channel that reloads an already-open project from disk in place. The only mechanism that
could do the latter was measured to raise `59999` and leave the workspace with **zero projects**,
so it is never attempted: `open_project` answers a disk-vs-memory divergence with a **structured
refusal** (`ok:false` + `reason` + `reloadNotAttempted:true` + `sideEffects:"none: ..."`) instead
of a silent no-op. To adopt the disk copy, call `remove_project` and then `open_project`
yourself. The project options the VBIDE model does expose (description, help file/context,
startup project, build output path, active project, read-only mode) are writable through
`set_project_options`; what this build cannot set (e.g. `startup_object`) comes back as a
structured `startupObjectError` plus a `startupObjectWorkaround`, never as a silent no-op.
`.vbp`-level keys the object model does not expose (the version resource and friends) are
written to the **file** by `project_meta{action:"set"}` (with a `.bak` and a post-write
sha256) - they are not injected into IDE memory.

---

## What is `modMcpTrace` for, and is it required?

**1. What it is for**: a tiny standard module holding just `McpTrace` / `McpTraceOn` (plus a switch variable). It exists only to serve `inject_trace` - the injected tracing calls land on it and it writes one JSON line per event to the trace file.

**2. Is it required?**: **No - VB6MCP itself does not need it.** The add-in (editing, checking, compiling, running, breakpoints, dialogs, forms, snapshots, packaging) is independent of it, and **a project that never used `inject_trace` and keeps no injected calls should not have this module at all**.

**3. When is it needed?**: only when you **used `inject_trace` and kept the injected calls**, because those calls resolve to `McpTrace` / `McpTraceOn`, which live in this module. *If you keep the injected calls and build/ship an EXE or DLL, the module must be compiled into it* (the receipt's `shipNote` says the same).

**4. How MCP adds it**: when the target project has no `modMcpTrace`, `inject_trace` **generates and attaches it itself, choosing the comment language from the IDE** (the module source is embedded in VB6MCP.dll, so **there is nothing to import by hand**); it then re-reads the component list to prove it landed, and **skips the step entirely when the module already exists** (idempotent). The receipt carries `moduleAutoAttached` / `moduleAttachAction` / `moduleAttachMs`, plus `moduleAttachError` when the attach fails (a failure is always reported, never silently downgraded).

- **Which comment language (UI first)**: the **IDE's own UI language** decides first (main-window menu captions File/Edit/View/Help, else the mode marker `[design]`/`[run]` in the title); if that cannot be read it falls back to the VB6.EXE version resource (`\VarFileInfo\Translation`; LANG_CHINESE => Chinese); if neither can be read the answer is **English** (readable on every code page, while Chinese becomes mojibake on a non-Chinese one). The receipt reports all three signals side by side - `ideLanguage` (version resource), `ideMenuLanguage` (UI) and `commentLanguage` (the final choice) - each with its own `*Source` explanation.

**5. When its *reference* is added or removed**:

- **Added**: on the first injection when the project does not contain it (see 4).
- **Removed**: **the new version never removes it automatically**. `remove_trace` deletes only **that id's injected block** and always keeps `modMcpTrace`; the receipt states `moduleRemoved:false` / `moduleLeftInPlace:true` / `moduleRemoveReason:"kept_by_design"` (or `module_absent` when the project has no such component), and `moduleRemoveNote` carries the two reasons below plus the manual-removal steps.
- **The old behaviour is withdrawn (and why)**: an earlier version used to have `remove_trace` take an auto-generated `modMcpTrace` out of the project as well (`moduleRemoveReason:"removed_autogen_module"`). That branch has been **deleted**, because (a) you may have added your own code to an auto-generated module and deleting the module would have destroyed that code, and (b) other `@mcp-trace` blocks may still exist elsewhere, so with the module gone those calls could no longer find `McpTrace`/`McpTraceOn` and **the project would stop compiling**.
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
baseline) **are** included, so the offline gate scripts under `tools/` (`check_schema_baseline`, `check_hub_alignment`, `check_live_tools docs/tools_schema.json`) run out of the box; the live audit (`mcp_strict_audit`) and `tests/smoke_test.py` need a running IDE.

## Contributing

See CONTRIBUTING.md (English; Chinese: CONTRIBUTING.zh-CN.md) and see section 7 above for the known gaps.

**Localization comes first**: this project is developed and verified on a **Chinese VB6**, and
the known coupling points are enumerable (see section 7 of this file). If you run an English or other
language build, even a report of "works / does not work" is a valuable contribution.
