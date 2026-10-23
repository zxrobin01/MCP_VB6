VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "McpServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' MCP (Model Context Protocol) JSON-RPC dispatcher + HTTP routing.
' All heavy lifting lives in CodeTools; this class is the wire protocol.

' Protocol negotiation. NEVER advertise a custom / stateless version here: the
' official @modelcontextprotocol/sdk client validates initialize.protocolVersion
' against its SUPPORTED_PROTOCOL_VERSIONS and DROPS the entire handshake when it
' does not match, so it registers ZERO tools while the server looks healthy
' (configured but does not take effect). Mirror the client when it asks for a
' version we implement, else fall back to the latest official version.
Private Const SUPPORTED_PROTOCOLS As String = "2025-11-25,2025-06-18,2025-03-26,2024-11-05,2024-10-07"
Private Const LATEST_PROTOCOL As String = "2025-11-25"
Private Const SERVER_NAME As String = "vb6-mcp"
' Product version: single source of truth is VB6MCP_PRODUCT_VERSION in Util.bas.

Private m_vbe As Object
Private m_tools As CodeTools
Private m_port As Long
Private m_startedAt As Date
Private m_quitPending As Boolean
Private m_restartPending As Boolean

Public Sub Init(ByVal vbe As Object)
    Set m_vbe = vbe
    Set m_tools = New CodeTools
    m_tools.Init vbe
    m_tools.SetServer Me
    ' hand the VBIDE holder to WinOps as a late-bound sink, so the
    ' optional code-pane wheel can move a pane through the object model (WM_VSCROLL
    ' to the pane was measured to do nothing 5/5). No new type library is needed:
    ' WinOps.bas is a standard module and calls the sink late-bound.
    WinOps.WheelSetSink m_tools
    m_startedAt = Now
End Sub

' deferred quit: the HTTP response must be sent before the IDE dies
Public Sub RequestQuit()
    m_quitPending = True
End Sub

Public Function ConsumeQuit() As Boolean
    If m_quitPending Then
        m_quitPending = False
        ConsumeQuit = True
    End If
End Function

Public Sub RequestRestart()
    m_restartPending = True
End Sub

Public Function ConsumeRestart() As Boolean
    If m_restartPending Then
        m_restartPending = False
        ConsumeRestart = True
    End If
End Function

Public Sub DoRestart()
    If Not m_tools Is Nothing Then m_tools.DoDeferredStart
End Sub

Public Sub DoQuit()
    On Error Resume Next
    If Not m_vbe Is Nothing Then m_vbe.Quit
End Sub

Public Property Let Port(ByVal p As Long)
    m_port = p
End Property

Public Property Get Port() As Long
    Port = m_port
End Property

Public Function ProjectName() As String
    If m_tools Is Nothing Then Exit Function
    ProjectName = McpProjTools.ProjectNameSafe(m_tools)
End Function

' the instance registry needs the project file and the DLL hash.
Public Function ProjectFileName() As String
    On Error Resume Next
    If m_tools Is Nothing Then Exit Function
    ProjectFileName = McpCls.ProjectFileNameSafe(m_tools)
End Function

Public Function ProductBuildSha256() As String
    On Error Resume Next
    If m_tools Is Nothing Then Exit Function
    ProductBuildSha256 = McpBuildTools.BuildSha256(m_tools)
End Function

' the settings form shows the tool count; forward it from the tool table so it
' can never drift from tools/list.
Public Function ToolCount() As Long
    On Error Resume Next
    If m_tools Is Nothing Then Exit Function
    ToolCount = m_tools.ToolCount()
End Function

' ---------- JSON-RPC (MCP) core ----------

Public Function HandleJsonRpc(ByVal requestText As String) As String
    Dim id As Variant
    Dim hasId As Boolean
    Dim obj As Variant
    If Not JsonParse(requestText, obj) Then
        HandleJsonRpc = JsonRpcError(Null, -32700, "Parse error", "")
        Exit Function
    End If
    On Error GoTo InternalFail
    If Not IsObject(obj) Then
        HandleJsonRpc = JsonRpcError(Null, -32600, "Invalid Request", "request must be a JSON object")
        Exit Function
    End If
    hasId = obj.Exists("id")
    If hasId Then id = obj.Item("id")
    Dim method As String
    method = JsonGetStr(obj, "method")
    If Len(method) = 0 Then
        If hasId Then
            HandleJsonRpc = JsonRpcError(id, -32600, "Invalid Request", "missing method")
        Else
            HandleJsonRpc = ""
        End If
        Exit Function
    End If
    Select Case method
        Case "initialize"
            HandleJsonRpc = JsonRpcResult(id, BuildInitializeResult(ClientProtocolVersion(obj)))
        Case "ping"
            Dim emptyDict As Object
            Set emptyDict = CreateObject("Scripting.Dictionary")
            HandleJsonRpc = JsonRpcResult(id, emptyDict)
        Case "tools/list"
            HandleJsonRpc = JsonRpcResultRaw(id, m_tools.ToolsListResultJson())
        Case "tools/call"
            HandleJsonRpc = HandleToolCall(obj, id)
        Case "resources/list"
            HandleJsonRpc = JsonRpcResult(id, BuildResourcesList())
        Case "resources/read"
            HandleJsonRpc = HandleResourceRead(obj, id)
        Case "prompts/list"
            HandleJsonRpc = JsonRpcResult(id, BuildPromptsList())
        Case "prompts/get"
            HandleJsonRpc = HandlePromptGet(obj, id)
        Case "notifications/initialized", "notifications/cancelled", "notifications/roots/list_changed", "notifications/tools/list_changed"
            HandleJsonRpc = ""
        Case Else
            If hasId Then
                HandleJsonRpc = JsonRpcError(id, -32601, "Method not found: " & method, "")
            Else
                HandleJsonRpc = ""
            End If
    End Select
    Exit Function
ParseFail:
    HandleJsonRpc = JsonRpcError(Null, -32700, "Parse error: " & Err.Description, "")
    Exit Function
InternalFail:
    HandleJsonRpc = JsonRpcError(Null, -32603, "Internal error: " & Err.Description, "")
End Function

' ---------- resources / prompts (capabilities beyond tools) ----------

Private Function ParamObj(ByVal obj As Object) As Object
    Dim params As Object
    Set params = Nothing
    If obj Is Nothing Then Exit Function
    On Error Resume Next
    If obj.Exists("params") Then Set params = obj.Item("params")
    On Error GoTo 0
    Set ParamObj = params
End Function

Private Function ParamStr(ByVal obj As Object, ByVal key As String) As String
    Dim params As Object
    Set params = ParamObj(obj)
    If params Is Nothing Then Exit Function
    ParamStr = JsonGetStr(params, key)
End Function

Private Function BuildResourcesList() As Object
    Dim items As Object
    Set items = New Collection
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("uri") = "vb6://project"
    d("name") = "project"
    d("title") = "The .vbp of the active VB6 project"
    d("mimeType") = "text/plain"
    items.Add d
    Dim names As Object
    Set names = Nothing
    On Error Resume Next
    Set names = m_tools.ComponentNamesList()
    On Error GoTo 0
    If Not names Is Nothing Then
        Dim v As Variant
        For Each v In names
            Dim e As Object
            Set e = CreateObject("Scripting.Dictionary")
            e("uri") = "vb6://component/" & CStr(v)
            e("name") = CStr(v)
            e("title") = "Source text of component " & CStr(v)
            e("mimeType") = "text/plain"
            items.Add e
        Next
    End If
    Dim root As Object
    Set root = CreateObject("Scripting.Dictionary")
    Set root("resources") = items
    Set BuildResourcesList = root
End Function

Private Function HandleResourceRead(ByVal obj As Object, ByVal id As Variant) As String
    Dim uri As String
    uri = ParamStr(obj, "uri")
    If Len(uri) = 0 Then
        HandleResourceRead = JsonRpcError(id, -32602, "params.uri is required", "")
        Exit Function
    End If
    Dim text As String
    text = ""
    Dim compPrefix As String
    compPrefix = "vb6://component/"
    If uri = "vb6://project" Then
        On Error Resume Next
        text = McpProjTools.ProjectVbpText(m_tools)
        On Error GoTo 0
    ElseIf Left$(LCase$(uri), Len(compPrefix)) = compPrefix Then
        On Error Resume Next
        text = m_tools.ComponentCodeText(Mid$(uri, Len(compPrefix) + 1))
        On Error GoTo 0
    Else
        HandleResourceRead = JsonRpcError(id, -32602, "unknown resource uri: " & uri, "")
        Exit Function
    End If
    Dim item As Object
    Set item = CreateObject("Scripting.Dictionary")
    item("uri") = uri
    item("mimeType") = "text/plain"
    item("text") = text
    Dim contents As Object
    Set contents = New Collection
    contents.Add item
    Dim root As Object
    Set root = CreateObject("Scripting.Dictionary")
    Set root("contents") = contents
    HandleResourceRead = JsonRpcResult(id, root)
End Function

Private Function PromptDef(ByVal nm As String, ByVal desc As String) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("name") = nm
    d("description") = desc
    Dim args As Object
    Set args = New Collection
    Dim a As Object
    Set a = CreateObject("Scripting.Dictionary")
    a("name") = "component"
    a("description") = "Optional component name to focus the checklist on"
    a("required") = False
    args.Add a
    Set d("arguments") = args
    Set PromptDef = d
End Function

Private Function BuildPromptsList() As Object
    Dim items As Object
    Set items = New Collection
    items.Add PromptDef("edit-checklist", "How to change VB6 code safely with this server: read, validate, anchor-edit, verify, compile.")
    items.Add PromptDef("release-checklist", "What to run before shipping a VB6 build from this server.")
    Dim root As Object
    Set root = CreateObject("Scripting.Dictionary")
    Set root("prompts") = items
    Set BuildPromptsList = root
End Function

Private Function HandlePromptGet(ByVal obj As Object, ByVal id As Variant) As String
    Dim nm As String
    nm = ParamStr(obj, "name")
    If Len(nm) = 0 Then
        HandlePromptGet = JsonRpcError(id, -32602, "params.name is required", "")
        Exit Function
    End If
    Dim text As String
    If nm = "edit-checklist" Then
        text = "VB6 safe-edit checklist" & vbCrLf & _
               "1) read_source / get_component_code before editing." & vbCrLf & _
               "2) validate_code (or check_component) on the intended text; fix every error first." & vbCrLf & _
               "3) preview_edit, then edit_by_anchor / apply_edits - anchors, never line numbers." & vbCrLf & _
               "4) verify (IDE == DISK + lint), then build for the authoritative compile." & vbCrLf & _
               "5) project open in the IDE -> make_project; project closed -> build." & vbCrLf & vbCrLf
    ElseIf nm = "release-checklist" Then
        text = "VB6 release checklist" & vbCrLf & _
               "1) validate_code + build: the tree must compile clean." & vbCrLf & _
               "2) verify every touched component (IDE == DISK)." & vbCrLf & _
               "3) project_meta set: MajorVer / MinorVer / RevisionVer / CompanyName / FileDescription." & vbCrLf & _
               "4) app_manifest write: choose the UAC level and DPI awareness the app needs." & vbCrLf & _
               "5) make_project or build; then git status / git diff to review the change set." & vbCrLf & _
               "6) project_snapshot create before handing the build over." & vbCrLf
    Else
        HandlePromptGet = JsonRpcError(id, -32602, "unknown prompt: " & nm, "")
        Exit Function
    End If
    Dim content As Object
    Set content = CreateObject("Scripting.Dictionary")
    content("type") = "text"
    content("text") = text
    Dim msg As Object
    Set msg = CreateObject("Scripting.Dictionary")
    msg("role") = "user"
    Set msg("content") = content
    Dim msgs As Object
    Set msgs = New Collection
    msgs.Add msg
    Dim root As Object
    Set root = CreateObject("Scripting.Dictionary")
    root("description") = nm
    Set root("messages") = msgs
    HandlePromptGet = JsonRpcResult(id, root)
End Function

Private Function HandleToolCall(ByVal obj As Object, ByVal id As Variant) As String
    Dim params As Object
    Set params = Nothing
    Dim args As Object
    Set args = Nothing
    If obj.Exists("params") Then
        ' objects must be read with Set (object-typed target); a scalar
        ' params value is trapped and treated as missing
        On Error Resume Next
        Set params = obj.Item("params")
        On Error GoTo 0
        If Not params Is Nothing Then
            If params.Exists("arguments") Then
                On Error Resume Next
                Set args = params.Item("arguments")
                On Error GoTo 0
            End If
        End If
    End If
    If params Is Nothing Then
        HandleToolCall = JsonRpcResult(id, ToolErrorResult("params object is required for tools/call"))
        Exit Function
    End If
    Dim name As String
    name = JsonGetStr(params, "name")
    If Len(name) = 0 Then
        HandleToolCall = JsonRpcResult(id, ToolErrorResult("params.name is required"))
        Exit Function
    End If
    WinOps.WheelRequestAttach "rpc_begin"
    HandleToolCall = JsonRpcResult(id, m_tools.CallTool(name, args))
    WinOps.WheelRequestAttach "rpc_end"
End Function

' ---------- HTTP routing ----------

Public Function HandleHttpRequest(ByVal conn As HttpConn) As String
    Dim method As String
    method = conn.Method
    Dim path As String
    path = conn.Path
    Dim q As Long
    q = InStr(path, "?")
    If q > 0 Then path = Left$(path, q - 1)
    If method = "OPTIONS" Then
        HandleHttpRequest = BuildHttpResponse("204 No Content", "text/plain", "", CorsHeaders(), conn.KeepAlive)
    ElseIf method = "GET" Then
        If path = "/health" Or path = "/ping" Then
            HandleHttpRequest = BuildHttpResponse("200 OK", "application/json", JsonStringify(BuildHealth()), CorsHeaders(), conn.KeepAlive)
        ElseIf path = "/" Then
            HandleHttpRequest = BuildHttpResponse("200 OK", "text/plain", "VB6 MCP server is running. Project: " & ProjectName() & ", port: " & CStr(m_port), "", conn.KeepAlive)
        ElseIf path = "/mcp" Then
            ' Streamable HTTP: GET is only for standalone SSE streams,
            ' which this server does not use - clients must POST
            HandleHttpRequest = BuildHttpResponse("405 Method Not Allowed", "text/plain", "Streamable HTTP endpoint requires POST (or DELETE)", "Allow: POST, DELETE", conn.KeepAlive)
        Else
            HandleHttpRequest = BuildHttpResponse("404 Not Found", "text/plain", "Not Found", "", conn.KeepAlive)
        End If
    ElseIf method = "POST" And (path = "/rpc" Or path = "/mcp" Or path = "/message") Then
        Dim bodyText As String
        bodyText = Utf8DecodeStr(conn.Body)
        Dim resp As String
        resp = HandleJsonRpc(bodyText)
        If Len(resp) = 0 Then
            ' JSON-RPC notification: no response body
            HandleHttpRequest = BuildHttpResponse("202 Accepted", "application/json", "", CorsHeaders(), conn.KeepAlive)
        Else
            Dim acceptHdr As String
            acceptHdr = LCase$(conn.Header("accept"))
            Dim wantSse As Boolean
            wantSse = (InStr(1, acceptHdr, "text/event-stream") > 0) And (InStr(1, acceptHdr, "application/json") = 0)
            If wantSse Then
                ' Streamable HTTP: client requested SSE-only responses
                Dim sseBody As String
                sseBody = "event: message" & vbCrLf & "data: " & resp & vbCrLf & vbCrLf
                HandleHttpRequest = BuildHttpResponse("200 OK", "text/event-stream", sseBody, CorsHeaders(), conn.KeepAlive)
            Else
                Dim extraHdr As String
                extraHdr = CorsHeaders()
                ' Streamable HTTP: hand the client a session id on
                ' initialize; requests are otherwise stateless
                Dim compact As String
                compact = Replace(bodyText, " ", "")
                If InStr(1, compact, """method"":""initialize""") > 0 Then
                    extraHdr = extraHdr & vbCrLf & "Mcp-Session-Id: " & NewHttpSessionId()
                End If
                HandleHttpRequest = BuildHttpResponse("200 OK", "application/json", resp, extraHdr, conn.KeepAlive)
            End If
        End If
    ElseIf method = "DELETE" And path = "/mcp" Then
        ' Streamable HTTP session termination (server is stateless)
        HandleHttpRequest = BuildHttpResponse("200 OK", "application/json", "", CorsHeaders(), conn.KeepAlive)
    Else
        HandleHttpRequest = BuildHttpResponse("405 Method Not Allowed", "text/plain", "Method Not Allowed", "", conn.KeepAlive)
    End If
End Function

Private Function NewHttpSessionId() As String
    On Error Resume Next
    Dim g As Object
    Set g = CreateObject("Scriptlet.TypeLib")
    ' the GUID is a fixed buffer: keep only up to the closing brace
    NewHttpSessionId = g.Guid
    Dim gb As Long
    gb = InStr(1, NewHttpSessionId, "}")
    If gb > 0 Then NewHttpSessionId = Left$(NewHttpSessionId, gb)
    If Len(NewHttpSessionId) = 0 Then
        Randomize
        NewHttpSessionId = "vb6mcp-" & Format$(Timer * 1000, "0") & "-" & CStr(Int(Rnd * 1000000))
    End If
End Function

Public Function BuildHttpResponse(ByVal status As String, ByVal contentType As String, ByVal body As String, ByVal extraHeaders As String, ByVal keepAlive As Boolean) As String
    ' returns a byte-per-char string: ASCII header + UTF-8 body
    Dim bodyBytes() As Byte
    bodyBytes = Utf8Encode(body)
    Dim bodyLen As Long
    bodyLen = 0
    On Error Resume Next
    bodyLen = UBound(bodyBytes) - LBound(bodyBytes) + 1
    If Err.Number <> 0 Then bodyLen = 0
    On Error GoTo 0
    If bodyLen < 0 Then bodyLen = 0
    Dim head As String
    head = "HTTP/1.1 " & status & vbCrLf
    If Len(contentType) > 0 Then
        head = head & "Content-Type: " & contentType & "; charset=utf-8" & vbCrLf
    End If
    head = head & "Content-Length: " & CStr(bodyLen) & vbCrLf
    If keepAlive Then
        head = head & "Connection: keep-alive" & vbCrLf
    Else
        head = head & "Connection: close" & vbCrLf
    End If
    If Len(extraHeaders) > 0 Then head = head & extraHeaders & vbCrLf
    head = head & vbCrLf
    If bodyLen > 0 Then
        BuildHttpResponse = head & BytesToStr(bodyBytes)
    Else
        BuildHttpResponse = head
    End If
End Function

' ---------- helpers ----------

Private Function ClientProtocolVersion(ByVal obj As Object) As String
    On Error Resume Next
    Dim p As Variant
    ClientProtocolVersion = ""
    ' NB: JsonGet uses Set internally. Writing `p = obj.Item("params")` instead
    ' raises error 450 (an object needs Set), which On Error Resume Next would
    ' silently swallow - p stays Empty and the version is never read.
    Dim isObj As Boolean
    If Not JsonGet(obj, "params", p, isObj) Then Exit Function
    If Not isObj Then Exit Function
    ClientProtocolVersion = JsonGetStr(p, "protocolVersion", "")
End Function

' Echo the version the client asked for when we implement it; otherwise answer with
' the latest official version. Returning anything else makes a strict client discard
' the whole handshake and register no tools at all.
Private Function NegotiateProtocol(ByVal clientVersion As String) As String
    If Len(clientVersion) > 0 Then
        If InStr(1, "," & SUPPORTED_PROTOCOLS & ",", "," & clientVersion & ",", vbTextCompare) > 0 Then
            NegotiateProtocol = clientVersion
            Exit Function
        End If
    End If
    NegotiateProtocol = LATEST_PROTOCOL
End Function

Private Function BuildInitializeResult(ByVal clientVersion As String) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("protocolVersion") = NegotiateProtocol(clientVersion)
    Dim caps As Object
    Set caps = CreateObject("Scripting.Dictionary")
    Dim toolsCap As Object
    Set toolsCap = CreateObject("Scripting.Dictionary")
    Set caps("tools") = toolsCap
    ' resources: the project file and every component source, readable by URI.
    ' prompts: reusable checklists so a client can surface them without a tool call.
    Dim resCap As Object
    Set resCap = CreateObject("Scripting.Dictionary")
    resCap("listChanged") = False
    Set caps("resources") = resCap
    Dim prCap As Object
    Set prCap = CreateObject("Scripting.Dictionary")
    prCap("listChanged") = False
    Set caps("prompts") = prCap
    Set d("capabilities") = caps
    Dim info As Object
    Set info = CreateObject("Scripting.Dictionary")
    info("name") = SERVER_NAME
    info("version") = VB6MCP_PRODUCT_VERSION
    Set d("serverInfo") = info
    ' contract v1.3.0: which tool-face contract this server implements
    d("contractVersion") = VB6MCP_CONTRACT_VERSION
    ' The instructions banner is published through initialize.instructions so the
    ' agent reads it before calling anything.
    On Error Resume Next
    d("instructions") = McpBuildTools.BuildInstructions(m_tools)
    On Error GoTo 0
    Set BuildInitializeResult = d
End Function

Private Function BuildHealth() As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("ok") = True
    d("server") = SERVER_NAME
    d("version") = VB6MCP_PRODUCT_VERSION
    d("port") = m_port
    d("project") = ProjectName()
    d("uptimeSeconds") = CLng(DateDiff("s", m_startedAt, Now))
    Set BuildHealth = d
End Function

Private Function JsonRpcResult(ByVal id As Variant, ByVal result As Object) As String
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("jsonrpc") = "2.0"
    d("id") = id
    If result Is Nothing Then
        d("result") = Null
    Else
        Set d("result") = result
    End If
    JsonRpcResult = JsonStringify(d)
End Function

' patch O: same envelope as JsonRpcResult, but the RESULT is already serialized.
' Used only by tools/list, whose result string is cached (see CodeTools).
Private Function JsonRpcResultRaw(ByVal id As Variant, ByVal resultJson As String) As String
    JsonRpcResultRaw = "{""jsonrpc"":""2.0"",""id"":" & JsonStringify(id) & ",""result"":" & resultJson & "}"
End Function

Private Function JsonRpcError(ByVal id As Variant, ByVal code As Long, ByVal message As String, ByVal data As String) As String
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("jsonrpc") = "2.0"
    d("id") = id
    Dim e As Object
    Set e = CreateObject("Scripting.Dictionary")
    e("code") = code
    e("message") = message
    If Len(data) > 0 Then e("data") = data
    Set d("error") = e
    JsonRpcError = JsonStringify(d)
End Function

Private Function ToolErrorResult(ByVal msg As String) As Object
    ' this transport-level failure path used to answer a bare "Error: <text>",
    ' which the caller cannot branch on. It now carries the same structured envelope as the
    ' add-in Fail handler ({ok:false, reason, message[, vbErrorCode]}).
    Dim vbErr As Long
    vbErr = Err.Number
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("ok") = False
    d("reason") = "unexpected_error"
    d("message") = msg
    If vbErr <> 0 Then d("vbErrorCode") = vbErr
    Dim res As Object
    Set res = CreateObject("Scripting.Dictionary")
    Dim content As Object
    Set content = New Collection
    Dim item As Object
    Set item = CreateObject("Scripting.Dictionary")
    item("type") = "text"
    item("text") = JsonStringify(d)
    content.Add item
    Set res("content") = content
    res("isError") = True
    Set ToolErrorResult = res
End Function

Private Function CorsHeaders() As String
    CorsHeaders = "Access-Control-Allow-Origin: *" & vbCrLf & _
                  "Access-Control-Allow-Methods: POST, GET, OPTIONS, DELETE" & vbCrLf & _
                  "Access-Control-Allow-Headers: Content-Type, MCP-Protocol-Version, MCP-Session-Id"
End Function


' R-1021-03 A4: ctl lane dispatch surface (PipeTransport owns the wire/auth).
Public Function CtlGateObj() As Object
    On Error Resume Next
    If m_tools Is Nothing Then Exit Function
    Set CtlGateObj = m_tools.CtlGateObj()
End Function
Public Function CtlStateObj() As Object
    On Error Resume Next
    If m_tools Is Nothing Then Exit Function
    Set CtlStateObj = m_tools.CtlStateObj()
End Function
Public Function CtlReleaseObj(ByVal minHeldMs As Long) As Object
    On Error Resume Next
    If m_tools Is Nothing Then Exit Function
    Set CtlReleaseObj = m_tools.CtlReleaseObj(minHeldMs)
End Function
Public Function CtlCancelObj(ByVal target As String) As Object
    On Error Resume Next
    If m_tools Is Nothing Then Exit Function
    Set CtlCancelObj = m_tools.CtlCancelObj(target)
End Function
Public Function CtlAbandonObj(ByVal confirm As Boolean, ByVal minHeldMs As Long) As Object
    On Error Resume Next
    If m_tools Is Nothing Then Exit Function
    Set CtlAbandonObj = m_tools.CtlAbandonObj(confirm, minHeldMs)
End Function
Public Function CtlToolsJson() As String
    On Error Resume Next
    If m_tools Is Nothing Then Exit Function
    CtlToolsJson = m_tools.CtlToolsJson()
End Function
Public Function CtlToolAllowed(ByVal name As String) As Boolean
    On Error Resume Next
    If m_tools Is Nothing Then Exit Function
    CtlToolAllowed = m_tools.CtlToolFastAllowed(name)
End Function
Public Function CtlToolIsReadOnly(ByVal name As String) As Boolean
    On Error Resume Next
    If m_tools Is Nothing Then Exit Function
    CtlToolIsReadOnly = m_tools.CtlToolIsReadOnly(name)
End Function
Public Function CtlToolRefusedObj() As Object
    Dim d As Object
    On Error Resume Next
    Set d = CreateObject("Scripting.Dictionary")
    d("ok") = False
    d("reason") = "readonly_lane_refused"
    d("detail") = "slow_tool_not_on_ctl_lane"
    d("allowed") = m_tools.CtlToolsJson()
    d("failureKind") = "toolError"
    Set CtlToolRefusedObj = d
End Function
Public Function CtlReadOnlyText(ByVal name As String, ByVal argsJson As String) As String
    On Error Resume Next
    If m_tools Is Nothing Then Exit Function
    CtlReadOnlyText = m_tools.CtlReadOnlyText(name, argsJson)
End Function

Public Function CtlToolsColl() As Object
    On Error Resume Next
    If m_tools Is Nothing Then Exit Function
    Set CtlToolsColl = m_tools.CtlToolCollection()
End Function

Public Function CtlInflightObj() As Object
    On Error Resume Next
    If m_tools Is Nothing Then Exit Function
    Set CtlInflightObj = m_tools.CtlInflightObj()
End Function
