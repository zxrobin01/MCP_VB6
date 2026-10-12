"""Final audit of CodeTools.cls / McpServer.cls / ServerForm.frm."""
import re

s = open("CodeTools.cls", encoding="utf-8").read()
print("tools.Add MakeTool count:", len(re.findall(r"tools\.Add MakeTool", s)))
calltool = s.split("Public Function CallTool")[1].split("End Function")[0]
print("CallTool cases:", len(re.findall(r'Case "', calltool)))

m = open("McpServer.cls", encoding="utf-8").read()
print("McpServer RequestQuit:", "RequestQuit" in m,
      "| ConsumeQuit:", "ConsumeQuit" in m,
      "| DoQuit:", "DoQuit" in m,
      "| SetServer Me:", "SetServer Me" in m)

f = open("ServerForm.frm", encoding="utf-8").read()
print("ServerForm quitArmed:", "m_quitArmed" in f,
      "| ConsumeQuit:", "ConsumeQuit" in f)

# duplicate Dim check inside procedures
for fn in ("CodeTools.cls", "McpServer.cls", "ServerForm.frm"):
    src = open(fn, encoding="utf-8").read()
    # crude: find repeated 'Dim X As' with different types in one procedure
    procs = re.split(r"(?:Private|Public) (?:Sub|Function|Property Get)", src)
    for idx, p in enumerate(procs):
        dims = re.findall(r"Dim\s+(\w+)\s+As\s+(\w+)", p)
        seen = {}
        for name, typ in dims:
            if name in seen and seen[name] != typ:
                print(fn, "proc", idx, "dup Dim", name, seen[name], "vs", typ)
            seen[name] = typ

# line length / ascii on all vb files
import glob
for fn in ["CodeTools.cls", "McpServer.cls", "ServerForm.frm", "Connect.cls",
           "ServerHost.cls", "HttpConn.cls", "JsonLib.bas", "Utf8.bas", "Util.bas"]:
    lines = open(fn, encoding="utf-8").read().splitlines()
    mx = max(len(l) for l in lines)
    nonascii = [i + 1 for i, l in enumerate(lines) if any(ord(c) > 127 for c in l)]
    print("%-16s lines=%4d maxlen=%3d nonascii=%s" %
          (fn, len(lines), mx, nonascii[:3] if nonascii else "ok"))
