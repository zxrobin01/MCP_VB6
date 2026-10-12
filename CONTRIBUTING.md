# Contributing to VB6MCP

> **English** | [中文](CONTRIBUTING.zh-CN.md)

Thanks for helping. VB6MCP turns the Microsoft Visual Basic 6.0 IDE into an
[MCP](https://modelcontextprotocol.io) server: a VB6 add-in (the tool surface) plus a
Node MCP hub (`bridge/`).

## 1. Prerequisites (read this first)

VB6MCP is an **add-in for VB6**, so you need a working VB6 installation of your own:

| Need | Why | Where it comes from |
|---|---|---|
| **Visual Basic 6.0 IDE** (`VB6.EXE`, `VB98`) | The add-in must load into it; every tool drives a live IDE. | Your own licensed VB6 install. It is **not** distributed here and cannot be. |
| **`MSADDNDR.DLL`** | VB6's add-in designer type library. The build needs it to compile the add-in. | Copy it from the same VB6 install into `tests/_msaddndr.dll` (git-ignored). |
| **Node.js 18+** | Runs the MCP hub `bridge/vb6mcp_hub.mjs`. | nodejs.org |
| **Python 3.8+** | Runs the maintenance/verification scripts in `tools/`. | python.org |

No third-party packages are needed: the hub uses only the Node standard library and the
tools use only the Python standard library (see `THIRD_PARTY_NOTICES.md`).

## 2. Build

```bat
build.bat
```

The build produces `VB6MCP.dll` (git-ignored). **After every build you must re-register
the add-in**, because each build creates a new typelib/GUID:

```bat
%SystemRoot%\SysWOW64\regsvr32.exe /s VB6MCP.dll
reg import register.reg /reg:32
```

Then prove the registration (this is the check that catches the classic failure where the
add-in never appears):

```bat
tools\verify_registration.bat VB6MCP.dll
```

`unregister.reg` removes the registration again.

## 3. Run the checks before you open a pull request

These are the gates this project actually uses. All of them are offline and read-only:

```bat
python tools\check_schema_baseline.py      :: tool schema identical to the frozen baseline
python tools\check_hub_alignment.py       :: hub face and add-in face agree
python tools\check_live_tools.py          :: live tool list matches the declared face
python tools\mcp_strict_audit.py          :: strict audit over the whole tool surface
python tools\gen_hub_snapshot.py          :: regenerate bridge/hub_schema.generated.json
tools\verify_registration.bat VB6MCP.dll  :: add-in registration self-check
```

A change to the tool surface must be **deliberate**: the frozen artifacts are
`tools/schema_baseline.json`, `docs/tools_schema.json`, `bridge/hub_schema.generated.json`,
`bridge/hub_tools.json` and `docs/CONTRACTS.md`. If your change is intended, update them in
the same commit and say so in the description; `check_schema_baseline.py` will fail otherwise
— that is by design.

## 4. Testing

* `tests/battery_test.py` — the end-to-end battery. It needs a live IDE (`VB6.EXE` on a
  project) and drives real tools; see the file header for the modes.
* Keep one IDE instance per port. Port **8766** is reserved for the maintainer's own
  interactive session; test instances must bind **8767 or above**.

## 5. House rules

* **Source encoding:** VB6 `.bas`/`.cls`/`.frm` files are stored as ANSI (GBK on a Chinese
  Windows). Do not convert them to UTF-8 — VB6 will garble non-ASCII string literals.
* **Never ship a NUL byte** in a source or text file. Some of our tooling refuses it.
* **Version numbers are load-bearing.** `bridge/hub_tools.json` carries the package
  version and the contract version lives in `docs/CONTRACTS.md`; do not edit either just to
  make a build look newer.
* **Every receipt must be honest.** If a tool cannot do something, it must say so with a
  machine-readable reason instead of silently degrading. This is the single most important
  rule in this project, and pull requests that trade honesty for a greener status will be
  rejected.
* **Do not commit** `VB6MCP.dll`, build logs, `*.before*`/`*.after*` backups, the internal
  workspaces (`_bus/`, `re_work/`, `docs/evidence/`) or the one-off `tests/fix_round*.py`
  archaeology scripts; `.gitignore` already covers them.

## 6. Reporting a bug

Include:

1. `mcp_info` output (contract version, tool count, `build.sha256`) — it identifies the exact
   binary you are running.
2. The full tool receipt of the failing call (not a summary).
3. Whether the IDE was in design/run/break mode, and the port it was on.

## 7. License

By contributing you agree that your contribution is licensed under the Apache License 2.0
(see `LICENSE`).