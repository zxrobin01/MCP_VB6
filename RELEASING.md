# Releasing VB6MCP

This is the checklist this project actually uses. Do not skip step 4 or 6: they are the two
checks that have caught real, shipped defects.

## 0. Decide the version

* Package version: `bridge/hub_tools.json` -> `version`.
* Protocol contract: `docs/CONTRACTS.md` - its version table states the current contract, `v1.10.0` at the time of writing.
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

`verify_registration.bat` must print `EXIT=0` **and** report `LoadBehavior=0x3`. A `0x1` means
the add-in will not appear even though registration returned success.

## 4. Tool-face gates (must all be green)

```bat
python tools\check_schema_baseline.py
python tools\check_hub_alignment.py
python tools\check_live_tools.py docs\tools_schema.json
rem the next line needs a running IDE (your own instance, 8767 or above - never 8766)
python tools\mcp_strict_audit.py --url http://127.0.0.1:8767/mcp
```

Expected: no schema change, no live-tool change, no hub-alignment failure, no strict failure. The strict audit drives a live IDE, so it needs an instance already running on your own port (8767 or above - never 8766).
If a gate fails because the tool surface intentionally changed, update the frozen artifacts
(`tools/schema_baseline.json`, `docs/tools_schema.json`, `bridge/hub_schema.generated.json`,
`bridge/hub_tools.json`, `docs/CONTRACTS.md`) in the same commit.

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
