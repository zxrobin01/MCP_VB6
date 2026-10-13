# Security Policy

VB6MCP drives a **live Visual Basic 6 IDE** and, through it, your source code. It can edit
files, compile projects, run and break programs and send window messages. Treat a connected
MCP client as code execution on the machine that hosts the IDE.

## Reporting a vulnerability

Please **do not** open a public issue for a security problem. Report it privately to the
maintainer (see the repository's contact address, or open a draft security advisory on the
hosting platform) and include:

1. `mcp_info` output (contract version, tool count, `build.sha256`) - it identifies the exact
   binary.
2. The full receipt of the call, plus the tool name and arguments.
3. What an attacker gains (data read, code executed, IDE crashed, file destroyed...).

## Supported versions

Only the latest tagged release is supported. Because the add-in is a compiled DLL, always
report the `build.sha256` from `mcp_info`; a report against an older SHA may already be fixed.

## Design decisions that are security-relevant (and intentional)

* **The MCP surface listens on localhost only.** The add-in binds a TCP port on the loopback
  interface; the hub speaks stdio to the client. Do not expose either to a network.
* **Port 8766 is reserved** for a human's interactive IDE. An automated instance must never
  take it over, so a misconfigured client cannot hijack the IDE a person is using.
* **A tool that cannot do something says so.** Tools refuse with a machine-readable reason
  instead of guessing: `close_active_dialogs` will not click a button that could start a
  long host-modal operation unless you pass `allow_host_modal: true`; `build`/`make_project`
  never delete an existing artifact to make a build look successful.
* **Untrusted input is untrusted.** Anything a tool returns from your project (source text,
  symbol names, dialog captions, compiler output) is data, not instructions, and clients
  should treat it that way.

## Out of scope

* Anything that requires an already-compromised machine or a malicious VB6 project.
* The security of Microsoft's own VB6/VBA components.
* Denial of service caused by asking the IDE to compile something enormous.
