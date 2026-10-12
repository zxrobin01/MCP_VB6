Attribute VB_Name = "Extensions"
Option Explicit
' Extension library (扩展库) - modelled on the dynsim MCP mcp_extensions.
'
' An extension is a DECLARATIVE tool pack: <dll dir>\mcp_extensions\*.json
'   { "name": "net-tools", "version": "1.0", "description": "...",
'     "tools": [ { "name": "http_get", "description": "...",
'                  "inputSchema": { ... },
'                  "handler": { "kind": "http_get", "url": "{url}" } } ] }
'
' Enabled packs are merged into tools/list and dispatched from CallTool.
' Capability is deliberately limited to a few safe primitives; the add-in never
' loads VB6 code from disk (VB6 cannot do that safely).
'
' state.json beside the packs: {"disabled": ["pack-name", ...]} (default enabled)

Private Const DIR_NAME As String = "mcp_extensions"
Private Const STATE_FILE As String = "state.json"

Public Function ExtDir() As String
    Dim b As String
    b = App.Path
    If Right$(b, 1) <> "\" Then b = b & "\"
    ExtDir = b & DIR_NAME
End Function

Public Function ExtPath(ByVal leaf As String) As String
    ExtPath = ExtDir() & "\" & leaf
End Function

Public Sub EnsureDir()
    On Error Resume Next
    If Len(Dir$(ExtDir(), vbDirectory)) = 0 Then MkDir ExtDir()
End Sub

' ---------- io helpers ----------

Private Function Q(ByVal s As String) As String
    Q = Chr$(34) & s & Chr$(34)
End Function

Private Function JStr(ByVal s As String) As String
    Dim t As String
    t = Replace(s, Chr$(92), Chr$(92) & Chr$(92))
    t = Replace(t, Chr$(34), Chr$(92) & Chr$(34))
    t = Replace(t, vbCrLf, Chr$(92) & "n")
    t = Replace(t, vbCr, Chr$(92) & "n")
    t = Replace(t, vbLf, Chr$(92) & "n")
    t = Replace(t, vbTab, Chr$(92) & "t")
    JStr = Chr$(34) & t & Chr$(34)
End Function

Private Function Trunc(ByVal s As String, ByVal n As Long) As String
    If Len(s) <= n Then Trunc = s Else Trunc = Left$(s, n) & "...[truncated]"
End Function

Public Function ReadText(ByVal path As String) As String
    Dim bb() As Byte, enc As String, s As String
    On Error Resume Next
    Encoding.ReadAllBytesTo path, bb
    If Encoding.ByteCount(bb) <= 0 Then Exit Function
    enc = Encoding.DetectEncoding(bb)
    Encoding.DecodeBytes bb, enc, s
    ReadText = s
End Function

Public Function WriteText(ByVal path As String, ByVal text As String) As Boolean
    Dim bb() As Byte
    On Error GoTo Bad
    Encoding.EncodeText text, "utf-8", bb
    Encoding.WriteAllBytesTo path, bb
    WriteText = True
    Exit Function
Bad:
    WriteText = False
End Function

Private Function LoadJson(ByVal path As String) As Object
    Dim txt As String, v As Variant
    txt = ReadText(path)
    If Len(txt) = 0 Then Exit Function
    On Error Resume Next
    If JsonParse(txt, v) Then
        If IsObject(v) Then
            If TypeName(v) = "Dictionary" Then Set LoadJson = v
        End If
    End If
End Function

' ---------- state ----------

Public Function LoadState() As Object
    Dim st As Object
    Set st = LoadJson(ExtPath(STATE_FILE))
    If st Is Nothing Then Set st = CreateObject("Scripting.Dictionary")
    Set LoadState = st
End Function

Private Function DisabledList(ByVal st As Object) As Collection
    Dim c As Variant, isObj As Boolean, out As New Collection, i As Long
    If JsonGet(st, "disabled", c, isObj) Then
        If isObj Then
            If TypeName(c) = "Collection" Then
                For i = 1 To c.Count
                    out.Add LCase$(CStr(c(i)))
                Next
            End If
        End If
    End If
    Set DisabledList = out
End Function

Private Function IsDisabled(ByVal st As Object, ByVal nm As String) As Boolean
    Dim c As Collection, i As Long
    Set c = DisabledList(st)
    For i = 1 To c.Count
        If c(i) = LCase$(nm) Then
            IsDisabled = True
            Exit Function
        End If
    Next
End Function

Private Function SaveState(ByVal st As Object) As Boolean
    EnsureDir
    SaveState = WriteText(ExtPath(STATE_FILE), JsonStringify(st))
End Function

' ---------- packs ----------

Public Function PackFiles() As Collection
    Dim out As New Collection, f As String
    On Error Resume Next
    f = Dir$(ExtDir() & "\*.json")
    Do While Len(f) > 0
        MaybePump
        If LCase$(f) <> LCase$(STATE_FILE) Then out.Add ExtPath(f)
        f = Dir$()
    Loop
    Set PackFiles = out
End Function

Public Function ToolSpecs() As Collection
    Dim out As New Collection, files As Collection, i As Long
    Dim doc As Object, st As Object, nm As String
    Dim arr As Variant, isObj As Boolean, j As Long, t As Object, item As Object
    EnsureDir
    Set st = LoadState()
    Set files = PackFiles()
    For i = 1 To files.Count
        MaybePump
        Set doc = LoadJson(files(i))
        If Not doc Is Nothing Then
            nm = JsonGetStr(doc, "name", "")
            If Len(nm) > 0 Then
                If Not IsDisabled(st, nm) Then
                    If JsonGet(doc, "tools", arr, isObj) Then
                        If isObj Then
                            If TypeName(arr) = "Collection" Then
                                For j = 1 To arr.Count
                                    If IsObject(arr(j)) Then
                                        Set t = arr(j)
                                        If TypeName(t) = "Dictionary" Then
                                            If Len(JsonGetStr(t, "name", "")) > 0 Then
                                                Set item = CreateObject("Scripting.Dictionary")
                                                item("pack") = nm
                                                item("name") = JsonGetStr(t, "name", "")
                                                item("description") = JsonGetStr(t, "description", "")
                                                Set item("spec") = t
                                                out.Add item
                                            End If
                                        End If
                                    End If
                                Next
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Next
    Set ToolSpecs = out
End Function

' ---------- handlers ----------

Private Function Template(ByVal tpl As String, ByVal args As Object) As String
    Dim s As String, p As Long, q As Long, key As String
    s = tpl
    p = InStr(s, "{")
    Do While p > 0
        q = InStr(p, s, "}")
        If q = 0 Then Exit Do
        key = Mid$(s, p + 1, q - p - 1)
        s = Left$(s, p - 1) & JsonGetStr(args, key, "") & Mid$(s, q + 1)
        p = InStr(s, "{")
    Loop
    Template = s
End Function

Private Function HttpGet(ByVal url As String) As String
    Dim http As Object, status As Long, body As String
    On Error Resume Next
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    If http Is Nothing Then Set http = CreateObject("MSXML2.ServerXMLHTTP")
    If http Is Nothing Then Set http = CreateObject("MSXML2.XMLHTTP")
    If http Is Nothing Then
        HttpGet = "{" & Q("ok") & ":false," & Q("error") & ":" & JStr("no HTTP component available") & "}"
        Exit Function
    End If
    On Error GoTo Bad
    http.Open "GET", url, False
    http.Send
    status = http.status
    body = http.responseText
    HttpGet = "{" & Q("ok") & ":true," & Q("status") & ":" & CStr(status) & "," & Q("body") & ":" & JStr(Trunc(body, 8000)) & "}"
    Exit Function
Bad:
    HttpGet = "{" & Q("ok") & ":false," & Q("error") & ":" & JStr(Err.Description) & "}"
End Function

Public Function Dispatch(ByVal toolName As String, ByVal args As Object) As String
    Dim specs As Collection, i As Long, item As Object, spec As Object, t As Object, h As Object
    Dim isObj As Boolean, kind As String, res As String
    Set specs = ToolSpecs()
    For i = 1 To specs.Count
        MaybePump
        Set item = specs(i)
        If LCase$(item("name")) = LCase$(toolName) Then
            Set spec = item("spec")
            Set h = Nothing
            If JsonGet(spec, "handler", t, isObj) Then
                If isObj Then
                    If TypeName(t) = "Dictionary" Then Set h = t
                End If
            End If
            If h Is Nothing Then
                Dispatch = "{" & Q("ok") & ":false," & Q("error") & ":" & JStr("extension tool has no handler") & "}"
                Exit Function
            End If
            kind = LCase$(JsonGetStr(h, "kind", ""))
            Select Case kind
                Case "http_get"
                    res = HttpGet(Template(JsonGetStr(h, "url", ""), args))
                Case "read_file"
                    Dispatch = "{" & Q("ok") & ":true," & Q("text") & ":" & JStr(Trunc(ReadText(Template(JsonGetStr(h, "path", ""), args)), 20000)) & "}"
                    Exit Function
                Case "write_file"
                    If WriteText(Template(JsonGetStr(h, "path", ""), args), Template(JsonGetStr(h, "text", ""), args)) Then
                        res = "{" & Q("ok") & ":true}"
                    Else
                        res = "{" & Q("ok") & ":false," & Q("error") & ":" & JStr("write failed") & "}"
                    End If
                Case Else
                    res = "{" & Q("ok") & ":false," & Q("error") & ":" & JStr("unknown handler kind: " & kind) & "}"
            End Select
            Dispatch = "{" & Q("tool") & ":" & JStr(item("name")) & "," & Q("pack") & ":" & JStr(item("pack")) & "," & Q("result") & ":" & res & "}"
            Exit Function
        End If
    Next
    Dispatch = ""
End Function
' ---------- management API (used by the ext_* tools) ----------

Private Function SetDisabled(ByVal nm As String, ByVal disabled As Boolean) As String
    Dim st As Object, cur As Collection, c As New Collection, i As Long, hit As Boolean
    Set st = LoadState()
    Set cur = DisabledList(st)
    For i = 1 To cur.Count
        If cur(i) <> LCase$(nm) Then c.Add cur(i)
    Next
    hit = False
    For i = 1 To c.Count
        If c(i) = LCase$(nm) Then hit = True
    Next
    If disabled And Not hit Then c.Add LCase$(nm)
    Set st("disabled") = c
    If SaveState(st) Then
        SetDisabled = "{" & Q("ok") & ":true," & Q("name") & ":" & JStr(nm) & "," & Q("enabled") & ":" & LCase$(CStr(Not disabled)) & "}"
    Else
        SetDisabled = "{" & Q("ok") & ":false," & Q("error") & ":" & JStr("cannot write state.json") & "}"
    End If
End Function

Public Function EnableJson(ByVal nm As String) As String
    EnableJson = SetDisabled(nm, False)
End Function

Public Function DisableJson(ByVal nm As String) As String
    DisableJson = SetDisabled(nm, True)
End Function

Public Function ListJson() As String
    Dim files As Collection, st As Object, i As Long, doc As Object, nm As String
    Dim s As String, arr As Variant, isObj As Boolean, n As Long, cnt As Long
    EnsureDir
    Set st = LoadState()
    Set files = PackFiles()
    For i = 1 To files.Count
        MaybePump
        Set doc = LoadJson(files(i))
        If Not doc Is Nothing Then
            nm = JsonGetStr(doc, "name", "")
            If Len(nm) > 0 Then
                cnt = cnt + 1
                n = 0
                If JsonGet(doc, "tools", arr, isObj) Then
                    If isObj Then
                        If TypeName(arr) = "Collection" Then n = arr.Count
                    End If
                End If
                If Len(s) > 0 Then s = s & ","
                s = s & "{" & Q("name") & ":" & JStr(nm) _
                    & "," & Q("version") & ":" & JStr(JsonGetStr(doc, "version", "")) _
                    & "," & Q("description") & ":" & JStr(JsonGetStr(doc, "description", "")) _
                    & "," & Q("tools") & ":" & CStr(n) _
                    & "," & Q("enabled") & ":" & LCase$(CStr(Not IsDisabled(st, nm))) _
                    & "," & Q("file") & ":" & JStr(files(i)) & "}"
            End If
        End If
    Next
    Dim etCount As Long
    etCount = ToolSpecs().Count
    ListJson = "{" & Q("dir") & ":" & JStr(ExtDir()) & "," & Q("count") & ":" & CStr(cnt) & "," & Q("packs") & ":[" & s & "]," & Q("enabledTools") & ":" & CStr(etCount) & "}"
End Function

Public Function ReloadJson() As String
    Dim pc As Long, tc As Long
    EnsureDir
    pc = PackFiles().Count
    tc = ToolSpecs().Count
    ReloadJson = "{" & Q("ok") & ":true," & Q("packs") & ":" & CStr(pc) & "," & Q("enabledTools") & ":" & CStr(tc) & "}"
End Function

Public Function InstallJson(ByVal path As String) As String
    Dim doc As Object, nm As String, dest As String, txt As String
    EnsureDir
    If Len(Dir$(path)) = 0 Then
        InstallJson = "{" & Q("ok") & ":false," & Q("error") & ":" & JStr("not found: " & path) & "}"
        Exit Function
    End If
    Set doc = LoadJson(path)
    If doc Is Nothing Then
        InstallJson = "{" & Q("ok") & ":false," & Q("error") & ":" & JStr("not a valid extension JSON") & "}"
        Exit Function
    End If
    nm = JsonGetStr(doc, "name", "")
    If Len(nm) = 0 Then
        InstallJson = "{" & Q("ok") & ":false," & Q("error") & ":" & JStr("extension needs a name") & "}"
        Exit Function
    End If
    dest = ExtPath(nm & ".json")
    txt = ReadText(path)
    If WriteText(dest, txt) Then
        InstallJson = "{" & Q("ok") & ":true," & Q("name") & ":" & JStr(nm) & "," & Q("path") & ":" & JStr(dest) & "}"
    Else
        InstallJson = "{" & Q("ok") & ":false," & Q("error") & ":" & JStr("cannot write: " & dest) & "}"
    End If
End Function

Public Function UninstallJson(ByVal nm As String) As String
    Dim files As Collection, i As Long, doc As Object
    Set files = PackFiles()
    For i = 1 To files.Count
        MaybePump
        Set doc = LoadJson(files(i))
        If Not doc Is Nothing Then
            If LCase$(JsonGetStr(doc, "name", "")) = LCase$(nm) Then
                On Error Resume Next
                Name files(i) As files(i) & ".removed"
                On Error GoTo 0
                UninstallJson = "{" & Q("ok") & ":true," & Q("name") & ":" & JStr(nm) & "," & Q("action") & ":" & JStr("renamed to .removed") & "}"
                Exit Function
            End If
        End If
    Next
    UninstallJson = "{" & Q("ok") & ":false," & Q("error") & ":" & JStr("no such extension: " & nm) & "}"
End Function
