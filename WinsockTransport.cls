VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0  'vbNone
  MTSTransactionMode = 0  'NotAnMTSObject
END
Attribute VB_Name = "WinsockTransport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ---------------------------------------------------------------------------
' Pure Winsock API transport - the MSWINSCK.OCX-free counterpart of ServerForm.
' Same public surface (StartListening/StopListening) so ServerHost can pick
' either one. Reuses the EXISTING, transport-agnostic pipeline untouched:
'     McpServer.HandleHttpRequest(conn) / HandleJsonRpc(text) / BuildHttpResponse
' Differences from the OCX path:
'   * the listener binds 127.0.0.1 explicitly (the OCX exposed LocalIP read-only,
'     so ServerForm had to accept-then-filter);
'   * events arrive via WSAAsyncSelect on a message-only window, not via a
'     control array;
'   * the watchdog is a Win32 SetTimer, not a Timer control.
' ---------------------------------------------------------------------------

Private Const MAX_CONNECTIONS As Integer = 16
Private Const DEFAULT_PORT As Long = 8766
' TASK-P1-037 (contract v1.3.1 H2): the dynamic range the hub scans (8766..8799).
Private Const FIRST_PORT As Long = 8766
Private Const MAX_PORT_TRIES As Long = 34
Private Declare Function WsaGetCurrentProcessId Lib "kernel32" Alias "GetCurrentProcessId" () As Long
Private Const RECV_CHUNK As Long = 65536
Private Const IDLE_TIMEOUT_SEC As Long = 30

Private m_mcp As McpServer
Private m_listen As Long
Private m_port As Long
Private m_tries As Long
Private m_listening As Boolean
Private m_explicitPort As Boolean    ' the caller NAMED this port - never silently move
Private m_tried As String            ' ports we attempted, in order
Private m_started As Boolean
Private m_quitArmed As Boolean
Private m_restartArmed As Boolean

Private m_conns As New Collection        ' CStr(sock) -> HttpConn
Private m_out As New Collection          ' CStr(sock) -> pending bytes (byte-per-char string)
Private m_sendPos As New Collection      ' CStr(sock) -> next offset
Private m_lastActive As New Collection   ' CStr(sock) -> Date
Private m_closeAfterSend As New Collection ' CStr(sock) -> True
Private m_sessions As New Collection     ' sessionId -> sock
' TASK-P1-064: the identity actually published, so the watchdog can re-publish it when it changes.
Private m_instProject As String
Private m_instProjectPath As String
Private m_instSha As String
Private m_instStartedAt As String

' --- public surface --------------------------------------------------------

Public Sub StartListening(ByVal mcp As McpServer, Optional ByVal basePort As Long = DEFAULT_PORT, Optional ByVal explicitPort As Boolean = False)
    Set m_mcp = mcp
    m_port = basePort
    m_explicitPort = explicitPort
    m_tried = ""
    m_tries = 0
    m_listening = False
    If Not WsaStart() Then
        LogLine "winsock transport: WSAStartup failed"
        Exit Sub
    End If
    If WsaMessageWindow() = 0 Then
        LogLine "winsock transport: could not create message window"
        Exit Sub
    End If
    WsaSetSink Me
    m_started = True
    TryListen
End Sub

Public Sub StopListening()
    On Error Resume Next
    WsaKillTimer
    WsaSetSink Nothing
    If m_listen <> 0 Then WsaClose m_listen
    m_listen = 0
    ' TASK-P1-037: normal exit must deregister the instance file
    If m_listening Then DeleteInstanceFile WsaGetCurrentProcessId()
    m_listening = False
    Dim i As Long
    For i = m_conns.Count To 1 Step -1
        WsaClose CLng(m_conns.Item(i))
    Next
    Set m_conns = New Collection
    Set m_out = New Collection
    Set m_sendPos = New Collection
    Set m_lastActive = New Collection
    Set m_closeAfterSend = New Collection
    Set m_sessions = New Collection
    If m_started Then
        ClearPortFileIfOurs m_port
        WsaDestroyWindow
        WsaStop
        m_started = False
    End If
    m_listening = False
    Set m_mcp = Nothing
End Sub

Public Property Get Port() As Long
    Port = m_port
End Property

' --- event entry point (called by WsaWndProc) -------------------

Public Function OnWsaMessage(ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    On Error GoTo Fail
    If uMsg = WM_TIMER_MSG Then
        TimerTick
    ElseIf uMsg = WM_SOCKET Then
        Dim s As Long
        Dim ev As Long
        Dim werr As Long
        s = wParam
        ev = lParam And &HFFFF&
        werr = (lParam \ &H10000) And &HFFFF&
        If werr <> 0 Then
            HandleSockError s, werr
        ElseIf s = m_listen Then
            If (ev And FD_ACCEPT) <> 0 Then DoAccept
        Else
            If (ev And FD_READ) <> 0 Then DoRead s
            If (ev And FD_WRITE) <> 0 Then DoFlush s
            If (ev And FD_CLOSE) <> 0 Then CloseConn s
        End If
    End If
    OnWsaMessage = 0
    Exit Function
Fail:
    LogLine "winsock transport error: " & CStr(Err.Number) & " " & Err.Description
    OnWsaMessage = 0
End Function

Private Sub HandleSockError(ByVal s As Long, ByVal werr As Long)
    If s = m_listen Then
        LogLine "winsock listen socket error " & CStr(werr)
        If Not m_listening Then NextPort
    Else
        LogLine "winsock socket error " & CStr(werr) & " on " & CStr(s)
        CloseConn s
    End If
End Sub

' --- listening -------------------------------------------------------------

Private Sub TryListen()
    On Error Resume Next
    If m_listen <> 0 Then WsaClose m_listen
    Dim werr As Long
    werr = 0
    m_listen = WsaListenSocket(m_port, werr)
    m_tries = m_tries + 1
    If Len(m_tried) = 0 Then
        m_tried = CStr(m_port)
    Else
        m_tried = m_tried & "," & CStr(m_port)
    End If
    If m_listen <> 0 Then
        If WsaRegister(m_listen, FD_ACCEPT) Then
            m_listening = True
            If Not m_mcp Is Nothing Then m_mcp.Port = m_port
            WritePortFile m_port
            RegisterInstance
            LogLine "winsock transport listening on http://127.0.0.1:" & CStr(m_port) & " (explicitPortOverride=" & CStr(m_explicitPort) & ")"
            WsaSetTimer 5000
            Exit Sub
        End If
        WsaClose m_listen
        m_listen = 0
    End If
    WsaSetTimer 350
End Sub

' TASK-P1-037: write %TEMP%\vb6mcp\instances\<pid>.json so the hub can resolve
' ownership instead of scanning ports and guessing.
Private Sub RegisterInstance()
    On Error Resume Next
    Dim sha As String, proj As String, ppath As String
    sha = "" : proj = "" : ppath = ""
    If Not m_mcp Is Nothing Then
        sha = m_mcp.ProductBuildSha256()
        proj = m_mcp.ProjectName()
        ppath = m_mcp.ProjectFileName()
    End If
    If Len(m_instStartedAt) = 0 Then m_instStartedAt = Format$(Now, "yyyy-mm-dd hh:nn:ss")
    m_instProject = proj
    m_instProjectPath = ppath
    m_instSha = sha
    WriteInstanceFile WsaGetCurrentProcessId(), m_port, proj, ppath, sha, m_instStartedAt
End Sub

' TASK-P1-064: the entry is published when the port comes up, which can precede VB6 opening
' the command-line project; nothing re-published it afterwards, so project=""/projectPath=""
' stayed for the whole process lifetime. Re-publish only when the identity really changes.
Private Sub RefreshInstanceIfChanged()
    On Error Resume Next
    If m_mcp Is Nothing Then Exit Sub
    Dim proj As String, ppath As String, sha As String
    proj = m_mcp.ProjectName()
    ppath = m_mcp.ProjectFileName()
    sha = m_mcp.ProductBuildSha256()
    If StrComp(proj, m_instProject, vbBinaryCompare) <> 0 _
       Or StrComp(ppath, m_instProjectPath, vbBinaryCompare) <> 0 _
       Or StrComp(sha, m_instSha, vbBinaryCompare) <> 0 Then
        LogLine "Re-publishing instance identity: project='" & proj & "' path='" & ppath & "'"
        RegisterInstance
    End If
End Sub

Private Sub NextPort()
    ' TASK-P1-037: an EXPLICIT port must never be silently replaced by another one -
    ' that is how two IDEs ended up sharing 8766 with undefined ownership.
    If m_explicitPort Then
        WsaKillTimer
        WriteBindConflict "port_in_use", m_tried, "the requested port " & CStr(m_port) & " is already owned by another process; this instance will NOT share it. Close the other server, set VB6MCP_PORT to a free port, or clear the override so the add-in can pick a free port in 8766..8799."
        LogLine "winsock transport: explicit port " & CStr(m_port) & " is in use - refusing to share it (wrote %TEMP%\vb6mcp\bind_conflict.json)"
        Exit Sub
    End If
    If m_tries < MAX_PORT_TRIES Then
        m_port = m_port + 1
        TryListen
    Else
        WsaKillTimer
        WriteBindConflict "no_free_port", m_tried, "the whole range " & CStr(FIRST_PORT) & ".." & CStr(FIRST_PORT + MAX_PORT_TRIES - 1) & " is occupied; free a port or set VB6MCP_PORT to a free one."
        LogLine "winsock transport could not bind any port in range (" & m_tried & ")"
    End If
End Sub

Private Sub TimerTick()
    On Error Resume Next
    ' TASK-P1-084: this transport has its own watchdog, and its WM_TIMER is dispatched
    ' by modal loops exactly like ServerForm's. The OCX transport already ticks WinOps
    ' from ServerForm.tmrWatch; do the same here so the armed Add Form confirm (and the
    ' optional wheel rescan) also work when VB6MCP_TRANSPORT=winsock is selected.
    WinOps.WheelTick
    If m_restartArmed Then
        m_restartArmed = False
        WsaSetTimer 5000
        If Not m_mcp Is Nothing Then m_mcp.DoRestart
        Exit Sub
    End If
    If m_quitArmed Then
        m_quitArmed = False
        WsaKillTimer
        If Not m_mcp Is Nothing Then m_mcp.DoQuit
        Exit Sub
    End If
    If Not m_listening Then
        NextPort
        Exit Sub
    End If
    SweepIdle
    RefreshInstanceIfChanged
End Sub

Private Sub DoAccept()
    On Error Resume Next
    Dim cs As Long
    Dim lasterr As Long
    Do
        cs = WsaAccept(m_listen)
        If cs = WSA_INVALID Then
            lasterr = WsaLastError()
            Exit Do
        End If
        If Not WsaRegister(cs, FD_READ Or FD_WRITE Or FD_CLOSE) Then
            WsaClose cs
        Else
            m_conns.Add cs, CStr(cs)
            TouchActive cs
        End If
    Loop
End Sub

' --- per-connection I/O ----------------------------------------------------

Private Function GetConn(ByVal s As Long) As HttpConn
    On Error Resume Next
    Dim c As HttpConn
    Set c = m_conns(CStr(s))
    On Error GoTo 0
    If c Is Nothing Then
        Set c = New HttpConn
        On Error Resume Next
        m_conns.Add c, CStr(s)
        On Error GoTo 0
    End If
    Set GetConn = c
End Function

Private Sub DoRead(ByVal s As Long)
    On Error GoTo Fail
    TouchActive s
    Dim b() As Byte
    Dim n As Long
    ReDim b(0 To RECV_CHUNK - 1)
    n = WsaRecv(s, VarPtr(b(0)), RECV_CHUNK)
    If n <= 0 Then
        If n = 0 Then CloseConn s
        Exit Sub
    End If
    If n < RECV_CHUNK Then ReDim Preserve b(0 To n - 1)
    Dim conn As HttpConn
    Set conn = GetConn(s)
    conn.AppendBytes b
    If conn.Ready Then HandleRequest s, conn
    Exit Sub
Fail:
    ' Card TASK-P1-010 root cause: this handler used to CloseConn with NO response,
    ' so an error raised while handling an already-received request was invisible to
    ' the caller (a bare "closed connection without response" while the add-in stayed
    ' alive) - the shape of both the native_breakpoint comment-line incident and the
    ' reported first-RPC drop. Answer with a 500 instead of dropping the socket.
    LogLine "winsock DoRead error #" & CStr(s) & ": " & CStr(Err.Number)
    If Not m_mcp Is Nothing Then QueueSend s, m_mcp.BuildHttpResponse("500 Internal Server Error", "text/plain", "internal error while handling the request", "", False)
    CloseConn s
End Sub

Private Sub DoFlush(ByVal s As Long)
    On Error Resume Next
    Dim pending As String
    Dim pos As Long
    pending = OutGet(s, pos)
    Do While Len(pending) > 0 And pos < Len(pending)
        Dim chunk As String
        chunk = Mid$(pending, pos + 1, 32768)
        Dim n As Long
        Dim bb() As Byte
        bb = StrToBytes(chunk)
        n = WsaSend(s, VarPtr(bb(0)), UBound(bb) - LBound(bb) + 1)
        If n <= 0 Then
            If WsaLastError() = 10035 Then Exit Do   ' WSAEWOULDBLOCK
            CloseConn s
            Exit Sub
        End If
        pos = pos + n
    Loop
    If pos >= Len(pending) Then
        ClearOut s
        If IsCloseAfterSend(s) Then CloseConn s
    Else
        OutSet s, pending, pos
    End If
End Sub

Private Function OutGet(ByVal s As Long, ByRef pos As Long) As String
    On Error Resume Next
    Dim v As Variant
    v = m_out(CStr(s))
    If IsEmpty(v) Then
        pos = 0
        OutGet = ""
    Else
        OutGet = CStr(v)
        v = m_sendPos(CStr(s))
        If IsEmpty(v) Then pos = 0 Else pos = CLng(v)
    End If
End Function

Private Sub OutSet(ByVal s As Long, ByVal text As String, ByVal pos As Long)
    On Error Resume Next
    m_out.Remove CStr(s)
    m_out.Add text, CStr(s)
    m_sendPos.Remove CStr(s)
    m_sendPos.Add pos, CStr(s)
End Sub

Private Sub ClearOut(ByVal s As Long)
    On Error Resume Next
    m_out.Remove CStr(s)
    m_sendPos.Remove CStr(s)
    m_closeAfterSend.Remove CStr(s)
End Sub

Private Sub QueueSend(ByVal s As Long, ByVal text As String)
    Dim existing As String
    Dim pos As Long
    existing = OutGet(s, pos)
    If Len(existing) > 0 And pos > 0 Then
        existing = Mid$(existing, pos + 1)
        pos = 0
    End If
    OutSet s, existing & text, pos
    DoFlush s
End Sub

Private Function IsCloseAfterSend(ByVal s As Long) As Boolean
    On Error Resume Next
    Dim v As Variant
    v = m_closeAfterSend(CStr(s))
    IsCloseAfterSend = Not IsEmpty(v)
End Function

Private Sub MarkCloseAfterSend(ByVal s As Long)
    On Error Resume Next
    m_closeAfterSend.Remove CStr(s)
    m_closeAfterSend.Add True, CStr(s)
End Sub

Private Sub CloseConn(ByVal s As Long)
    On Error Resume Next
    LogLine "winsock CloseConn " & CStr(s)
    WsaClose s
    m_conns.Remove CStr(s)
    ClearOut s
    m_lastActive.Remove CStr(s)
    Dim k As Long
    For k = m_sessions.Count To 1 Step -1
        If CLng(m_sessions.Item(k)) = s Then m_sessions.Remove k
    Next
End Sub

Private Sub TouchActive(ByVal s As Long)
    On Error Resume Next
    m_lastActive.Remove CStr(s)
    m_lastActive.Add Now, CStr(s)
End Sub

Private Function GetLastActive(ByVal s As Long) As Date
    On Error Resume Next
    GetLastActive = m_lastActive(CStr(s))
    If Err.Number <> 0 Then GetLastActive = Now
End Function

Private Function IsSseSock(ByVal s As Long) As Boolean
    On Error Resume Next
    Dim k As Long
    For k = 1 To m_sessions.Count
        If CLng(m_sessions.Item(k)) = s Then
            IsSseSock = True
            Exit Function
        End If
    Next
End Function

Private Sub SweepIdle()
    On Error Resume Next
    Dim i As Long
    For i = m_conns.Count To 1 Step -1
        Dim s As Long
        s = CLng(m_conns.Item(i))
        If IsCloseAfterSend(s) Then
            If DateDiff("s", GetLastActive(s), Now) > 15 Then CloseConn s
        ElseIf Not IsSseSock(s) Then
            If DateDiff("s", GetLastActive(s), Now) > IDLE_TIMEOUT_SEC Then
                LogLine "winsock closing idle connection"
                CloseConn s
            End If
        End If
    Next
End Sub

' --- request handling (same shape as ServerForm) ---------------------------

Private Sub HandleRequest(ByVal s As Long, ByVal conn As HttpConn)
    On Error GoTo Failed
    If conn.ErrorText <> "" Then
        QueueSend s, m_mcp.BuildHttpResponse("413 Request Entity Too Large", "text/plain", "Request Entity Too Large", "", False)
        CloseConn s
        Exit Sub
    End If
    Dim method As String, rawPath As String
    method = conn.Method
    rawPath = conn.Path
    If method = "GET" And Left$(rawPath, 4) = "/sse" Then
        Dim sessionId As String
        sessionId = NewSessionId()
        RegisterSession sessionId, s
        conn.MarkSse
        QueueSend s, BuildSseHeaders() & "event: endpoint" & vbCrLf & "data: /message?sessionId=" & sessionId & vbCrLf & vbCrLf
        Exit Sub
    End If
    If method = "DELETE" And Left$(rawPath, 8) = "/message" Then
        Dim delSession As String
        delSession = QueryParam(rawPath, "sessionId")
        If delSession <> "" Then RemoveSession delSession
        QueueSend s, m_mcp.BuildHttpResponse("200 OK", "application/json", "", "Access-Control-Allow-Origin: *", False)
        MarkCloseAfterSend s
        DoFlush s
        Exit Sub
    End If
    Dim isMessage As Boolean
    isMessage = (method = "POST" And Left$(rawPath, 8) = "/message")
    Dim sseSession As String
    If isMessage Then
        sseSession = QueryParam(rawPath, "sessionId")
        If sseSession = "" Then
            QueueSend s, m_mcp.BuildHttpResponse("400 Bad Request", "text/plain", "Missing sessionId", "", False)
            MarkCloseAfterSend s
            DoFlush s
            Exit Sub
        End If
        If Not SessionExists(sseSession) Then
            QueueSend s, m_mcp.BuildHttpResponse("404 Not Found", "text/plain", "Unknown session", "", False)
            MarkCloseAfterSend s
            DoFlush s
            Exit Sub
        End If
    End If
    Dim response As String
    If isMessage Then
        Dim rpcResp As String
        rpcResp = m_mcp.HandleJsonRpc(Utf8DecodeStr(conn.Body))
        If Len(rpcResp) > 0 Then PushSse sseSession, rpcResp
        response = m_mcp.BuildHttpResponse("202 Accepted", "application/json", "", "Access-Control-Allow-Origin: *", True)
    Else
        response = m_mcp.HandleHttpRequest(conn)
    End If
    QueueSend s, response
    CheckDeferred
    If conn.IsSse Then
        conn.ClearAll
    ElseIf conn.KeepAlive Then
        conn.AdvanceToNext
        If conn.Ready Then HandleRequest s, conn
    Else
        MarkCloseAfterSend s
        DoFlush s
    End If
    Exit Sub
Failed:
    QueueSend s, m_mcp.BuildHttpResponse("500 Internal Server Error", "text/plain", "Internal Server Error: " & Err.Description, "", False)
    MarkCloseAfterSend s
    DoFlush s
End Sub

Private Sub CheckDeferred()
    On Error Resume Next
    If m_mcp Is Nothing Then Exit Sub
    If m_mcp.ConsumeQuit() Then
        m_quitArmed = True
        WsaSetTimer 150
    End If
    If m_mcp.ConsumeRestart() Then
        m_restartArmed = True
        WsaSetTimer 500
    End If
End Sub

' --- SSE -------------------------------------------------------------------

Private Sub RegisterSession(ByVal sessionId As String, ByVal s As Long)
    On Error Resume Next
    m_sessions.Add s, sessionId
End Sub

Private Function SessionExists(ByVal sessionId As String) As Boolean
    On Error Resume Next
    Dim v As Variant
    v = m_sessions(sessionId)
    SessionExists = (Err.Number = 0) And Not IsEmpty(v)
End Function

Private Sub RemoveSession(ByVal sessionId As String)
    On Error Resume Next
    Dim s As Variant
    s = m_sessions(sessionId)
    m_sessions.Remove sessionId
    If Not IsEmpty(s) Then CloseConn CLng(s)
End Sub

Private Sub PushSse(ByVal sessionId As String, ByVal jsonText As String)
    On Error Resume Next
    Dim s As Variant
    s = m_sessions(sessionId)
    If IsEmpty(s) Then Exit Sub
    QueueSend CLng(s), "event: message" & vbCrLf & "data: " & jsonText & vbCrLf & vbCrLf
End Sub

Private Function BuildSseHeaders() As String
    BuildSseHeaders = "HTTP/1.1 200 OK" & vbCrLf & _
        "Content-Type: text/event-stream" & vbCrLf & _
        "Cache-Control: no-cache" & vbCrLf & _
        "Connection: keep-alive" & vbCrLf & _
        "Access-Control-Allow-Origin: *" & vbCrLf
End Function

Private Function NewSessionId() As String
    On Error Resume Next
    Dim g As Object
    Set g = CreateObject("Scriptlet.TypeLib")
    NewSessionId = g.Guid
    Dim gb As Long
    gb = InStr(1, NewSessionId, "}")
    If gb > 0 Then NewSessionId = Left$(NewSessionId, gb)
    If Len(NewSessionId) = 0 Then
        NewSessionId = "s" & Format$(Timer * 1000, "0") & "_" & CStr(Int(Rnd * 99999))
    End If
End Function

Private Function QueryParam(ByVal path As String, ByVal name As String) As String
    Dim q As Long
    q = InStr(path, "?")
    If q = 0 Then Exit Function
    Dim query As String
    query = Mid$(path, q + 1)
    Dim pairs() As String
    pairs = Split(query, "&")
    Dim i As Long
    For i = 0 To UBound(pairs)
        Dim kv() As String
        kv = Split(pairs(i), "=")
        If UBound(kv) >= 1 Then
            If LCase$(kv(0)) = LCase$(name) Then
                QueryParam = kv(1)
                Exit Function
            End If
        End If
    Next
End Function
