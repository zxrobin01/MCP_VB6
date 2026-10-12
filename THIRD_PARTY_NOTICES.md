# Third-Party Notices

VB6MCP itself is licensed under the Apache License, Version 2.0 (see LICENSE).

## Runtime dependencies

None. The MCP server surface (bridge/) uses only the Node.js standard library,
and the maintenance and verification scripts (tools/) use only the Python 3
standard library. No third-party package is vendored or required at runtime.

## Build-time / platform dependencies (NOT distributed)

| Component | Owner | How it is obtained |
|---|---|---|
| Microsoft Visual Basic 6.0 IDE (VB6.EXE) | Microsoft Corporation | The user's own licensed VB6 installation. Not distributed here. |
| `MSADDNDR.DLL` (VB6 add-in designer type library) | Microsoft Corporation | Copied from the user's own VB6 installation at build time (`tests/_msaddndr.dll` is git-ignored). Not distributed here. |
| VBA6.DLL / VBE6.DLL (VB6 debug engine) | Microsoft Corporation | Present in the user's own VB6 installation. Not distributed here. |
| Visual Basic 6.0 runtime (MSVBVM60.DLL) | Microsoft Corporation | Present on the target machine as part of the user's VB6 installation. Not distributed here. |

## Generated artifacts

`VB6MCP.dll` is a build product of this project's own sources. It is git-ignored
and is not part of the source distribution.

## Trademarks

Visual Basic, VB6, VBA, Windows, Microsoft and MSADDNDR are trademarks of
Microsoft Corporation. This project claims no rights in them and is not
affiliated with or endorsed by Microsoft.
