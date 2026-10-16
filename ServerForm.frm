VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.1#0"; "MSWINSCK.OCX"
Begin VB.Form ServerForm 
   BorderStyle     =   0  'None
   Caption         =   "VB6MCP"
   ClientHeight    =   300
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   300
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   300
   ScaleWidth      =   300
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Timer tmrWatch 
      Interval        =   0
      Left            =   0
      Top             =   0
   End
   Begin MSWinsockLib.Winsock sck 
      Index           =   0
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393217
   End
End
Attribute VB_Name = "ServerForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Hidden form hosting the Winsock listener (a Winsock control array) and a
' watchdog timer used to scan for a free port. Handles plain HTTP JSON-RPC
' (POST /rpc) plus a minimal MCP streamable-HTTP transport (GET /sse +
' POST /message).

Private Const MAX_CONNECTIONS As Integer = 16
Private Const DEFAULT_PORT As Long = 8766
' TASK-P1-037 (contract v1.3.1 H2): same dynamic range as the winsock transport / hub.
Private Const FIRST_PORT As Long = 8766
Private Const MAX_PORT_TRIES As Long = 34
Private Declare Function FormProcessId Lib "kernel32" Alias "GetCurrentProcessId" () As Long

Private m_mcp As McpServer
Private m_port As Long
Private m_tries As Long
Private m_conns As New Collection
Private m_sessions As New Collection
Private m_listening As Boolean
Private m_explicitPort As Boolean    ' a NAMED port is never silently replaced
Private m_tried As String
Private m_quitArmed As Boolean
Private m_restartArmed As Boolean
Private m_lastActive As New Collection
Private m_closeAfterSend As New Collection
' TASK-P1-064: the identity ACTUALLY published for this instance, so the watchdog can
' re-publish it the moment it changes (a project opened after listen was never re-published
' => empty project/projectPath stayed in the registry forever).
Private m_instProject As String
Private m_instProjectPath As String
Private m_instSha As String
Private m_instStartedAt As String
' TASK-57 (ADR section 0 / 7): the named-pipe track. PipeTransport is kernel32
' only on purpose - WSACleanup is process-wide and, as measured in the P1-065
' spike, it killed the HTTP listener of this very add-in.
Private m_pipe As PipeTransport
Private m_mode As String
Private m_portPublished As Boolean
Private Const IDLE_TIMEOUT_SEC As Long = 30

' TASK-49: in-flight (BUSY) request tracking. A connection whose request is being
' SERVED must never be closed by the idle sweeper: the serve runs on this same
' single thread, and every Util.MaybePump -> DoEvents inside a long tool call lets
' the 5 s watchdog fire SweepIdleConnections. Before this, a call that legitimately
' ran longer than IDLE_TIMEOUT_SEC lost its response (task-47: client EOF at
' 32.236 s with ZERO bytes). BUSY is set only when a COMPLETE request is in hand,
' and cleared on the OUTERMOST return of the serve, on the error path, and on every
' teardown, so a slot can never become permanently unsweepable.
Private m_busy(1 To MAX_CONNECTIONS) As Boolean
Private m_busySince(1 To MAX_CONNECTIONS) As Date
' Leaked-flag backstop: a BUSY slot older than this becomes sweepable again AND
' logs (unlike a normal busy skip, which is silent). Overridable for testing.
Private Const IDLE_BUSY_CAP_SEC_DEF As Long = 300
Private m_idleBusyCapSec As Long

Public Sub StartListening(ByVal mcp As McpServer, Optional ByVal basePort As Long = DEFAULT_PORT, Optional ByVal explicitPort As Boolean = False)
    Randomize
    Set m_mcp = mcp
    m_port = basePort
    m_explicitPort = explicitPort
    m_tried = ""
    m_tries = 0
    m_listening = False
    Dim i As Integer
    For i = 1 To MAX_CONNECTIONS
        Load sck(i)
    Next
    m_idleBusyCapSec = IdleBusyCapFromEnv()
    ' TASK-57 transport switch (ADR section 0.2 rule 3: pipe | tcp | auto, DEFAULT
    ' pipe). ResolveMode runs on PipeTransport's predeclared instance, so reading
    ' the setting does not create a pipe object.
    m_mode = PipeTransport.ResolveMode()
    If m_mode = "pipe" Or m_mode = "auto" Then
        Set m_pipe = New PipeTransport
        ' "auto" keeps the TCP listener below as the hub's fallback leg; "pipe"
        ' does not, so this IDE then occupies no TCP port at all (rule 1).
        m_pipe.StartServer m_mcp, m_mode, (m_mode = "auto")
        ' Publish now: in auto mode the TCP leg may still be probing for a port,
        ' and the hub must be able to find the pipe either way.
        RegisterInstance
    End If
    If m_mode = "pipe" Then
        tmrWatch.Interval = m_pipe.TickMs
        tmrWatch.Enabled = True
        LogLine "transport=pipe reason=forced_pipe tcpEnabled=false pipe=" & m_pipe.PipeNameRedacted() & " nonce=" & m_pipe.NoncePrefix() & "... tickMs=" & CStr(m_pipe.TickMs)
        Exit Sub
    End If
    TryListen
End Sub

Public Sub StopListening()
    tmrWatch.Enabled = False
    ' TASK-57: only clear the port file if THIS instance actually published one -
    ' in pipe-only mode m_port is just the requested default, and clearing it
    ' could delete a live instance's file.
    If m_portPublished Then ClearPortFileIfOurs m_port
    ' TASK-P1-037: a normal exit must deregister the instance file. TASK-57: the
    ' pipe-only track never publishes a port, so it has to be covered separately.
    If m_listening Or (Not m_pipe Is Nothing) Then DeleteInstanceFile FormProcessId()
    If Not m_pipe Is Nothing Then
        m_pipe.StopServer
        Set m_pipe = Nothing
    End If
    m_listening = False
    Dim i As Integer
    On Error Resume Next
    sck(0).Close
    For i = 1 To MAX_CONNECTIONS
        ClearBusy i
        sck(i).Close
    Next
    Set m_conns = Nothing
    Set m_sessions = Nothing
    Set m_mcp = Nothing
End Sub

Private Sub TryListen()
    On Error Resume Next
    sck(0).Close
    sck(0).LocalPort = m_port
    ' NOTE: MSWINSCK.OCX exposes LocalIP as read-only, so the listener binds
    ' all local interfaces; connections are filtered to 127.0.0.1 in
    ' sck_ConnectionRequest instead.
    sck(0).Listen
    m_tries = m_tries + 1
    If Len(m_tried) = 0 Then
        m_tried = CStr(m_port)
    Else
        m_tried = m_tried & "," & CStr(m_port)
    End If
    tmrWatch.Interval = 350
    tmrWatch.Enabled = True
End Sub

' TASK-P1-037: register this IDE in %TEMP%\vb6mcp\instances\<pid>.json so the hub
' resolves ownership instead of scanning ports and guessing.
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
    If Not m_pipe Is Nothing Then
        If m_pipe.Active() Then
            ' TASK-57: ADR section 4.1 - transport/pipe/sessionId/nonce/instanceId/
            ' tcpEnabled must be in the SAME entry, or the hub has nothing to pair
            ' the instance against. In TCP mode the legacy writer below keeps the
            ' entry byte-for-byte as before (ADR section 4.2 rule 1: a missing
            ' transport/pipe means tcp).
            m_pipe.PublishInstance PipeRegistryPort(), m_pipe.TcpEnabled(), proj, ppath, sha, m_instStartedAt
            PruneDeadInstanceFiles FormProcessId()
            Exit Sub
        End If
    End If
    WriteInstanceFile FormProcessId(), m_port, proj, ppath, sha, m_instStartedAt
    ' TASK-48: a successful registration proves THIS process is alive - the right
    ' moment to drop the registrations left behind by instances that died without
    ' running their exit path (crash / Stop-Process -Force / taskkill).
    PruneDeadInstanceFiles FormProcessId()
End Sub

' TASK-P1-064: the registry entry is published the moment the port is up, which can be
' BEFORE VB6 has finished opening the command-line project - and a project opened later
' (or reopened) was never re-published, so the entry kept project=""/projectPath="" for the
' rest of the process lifetime. That is the registry shape the hub reported for a NON-FIRST
' instance (a second IDE started with a .vbp another IDE already had open opens with NO
' project at all). Re-publish whenever the identity actually changes; identical identities
' are NOT rewritten (no file churn).
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

' TASK-57: 0 means "this instance has no TCP leg" (pipe-only); the hub decides
' from tcpEnabled, not from the port number.
Private Function PipeRegistryPort() As Long
    If m_listening Then
        PipeRegistryPort = m_port
    Else
        PipeRegistryPort = 0
    End If
End Function

Private Sub tmrWatch_Timer()
    On Error GoTo TimerFail
    ' TASK-57: exactly one bounded pipe servicing pass per tick, on this same UI
    ' thread / same watchdog. PipeTransport handles re-entrancy internally (a
    ' nested tick, reached from DoEvents during a long tool call, only emits a
    ' progress frame - it never re-enters request handling).
    If Not m_pipe Is Nothing Then m_pipe.Tick
    If m_restartArmed Then
        m_restartArmed = False
        tmrWatch.Interval = 5000
        If Not m_mcp Is Nothing Then m_mcp.DoRestart
        Exit Sub
    End If
    If m_quitArmed Then
        m_quitArmed = False
        tmrWatch.Enabled = False
        If Not m_mcp Is Nothing Then m_mcp.DoQuit
        Exit Sub
    End If
    ' TASK-57: in pipe-only mode no winsock listener was ever created, so the
    ' port-probing path below must not run (sck() was only Loaded, never Listen-ed).
    If m_mode = "pipe" Then
        WinOps.WheelTick
        RefreshInstanceIfChanged
        Exit Sub
    End If
    If m_listening Then
        ' TASK-P1-067: thriftily (re)install the optional code-pane wheel on code
        ' panes that appeared after startup. Returns at once while the wheel is
        ' OFF (the default), so the timer's cost is unchanged.
        WinOps.WheelTick
        SweepIdleConnections
        RefreshInstanceIfChanged
        Exit Sub
    End If
    If sck(0).State = sckListening Then
        m_listening = True
        m_mcp.Port = m_port
        WritePortFile m_port
        m_portPublished = True
        RegisterInstance
        LogLine "VB6MCP listening on http://127.0.0.1:" & CStr(m_port) & " (explicitPortOverride=" & CStr(m_explicitPort) & ")"
        ' TASK-57: keep the watchdog fast enough to service the pipe track; the TCP
        ' track alone keeps the historical 5000 ms cadence.
        If m_pipe Is Nothing Then
            tmrWatch.Interval = 5000
        Else
            tmrWatch.Interval = m_pipe.TickMs
        End If
        tmrWatch.Enabled = True
    ElseIf m_explicitPort Then
        ' TASK-P1-037: the port was NAMED (VB6MCP_PORT / registry) - never take
        ' another one behind the caller@s back; report a structured conflict.
        tmrWatch.Enabled = False
        WriteBindConflict "port_in_use", m_tried, "the requested port " & CStr(m_port) & " is already owned by another process; this instance will NOT share it. Close the other server, set VB6MCP_PORT to a free port, or clear the override so the add-in can pick a free port in 8766..8799."
        LogLine "VB6MCP: explicit port " & CStr(m_port) & " is in use - refusing to share it (wrote %TEMP%\vb6mcp\bind_conflict.json)"
    ElseIf m_tries < MAX_PORT_TRIES Then
        m_port = m_port + 1
        TryListen
    Else
        tmrWatch.Enabled = False
        WriteBindConflict "no_free_port", m_tried, "the whole range " & CStr(FIRST_PORT) & ".." & CStr(FIRST_PORT + MAX_PORT_TRIES - 1) & " is occupied; free a port or set VB6MCP_PORT to a free one."
        LogLine "VB6MCP could not bind any port in range (" & m_tried & ")"
    End If
    Exit Sub
TimerFail:
    LogLine "Timer error: " & CStr(Err.Number) & " " & Err.Description
    tmrWatch.Enabled = True
End Sub

Private Sub SweepIdleConnections()
    ' Close keep-alive connections that are idle too long so the 16-socket
    ' pool cannot be exhausted by clients that never send FIN. SSE streams
    ' are exempt (they legitimately idle between events).
    ' Also force-close sockets whose deferred close never completed.
    On Error Resume Next
    Dim i As Integer
    For i = 1 To MAX_CONNECTIONS
        ' TASK-P1-064: a request socket can be left in sckClosing (8) - the documented
        ' "Close while sending => RST => stuck in sckClosing" path (HANDOVER 39) - and such
        ' a slot is in NO map, so nothing ever reclaimed it. Sixteen of them wedge the whole
        ' pool and the instance stops serving /rpc at once (the hub sees ECONNRESET ->
        ' forward_failed, and once the probe fails it reports registry_port_not_mcp).
        If sck(i).State = sckClosing Then
            If DateDiff("s", GetLastActive(i), Now) > 5 Then
                LogLine "Reclaiming socket #" & CStr(i) & " stuck in sckClosing"
                ResetSocket i
            End If
        End If
        Dim marked As Variant
        marked = Empty
        marked = m_closeAfterSend(CStr(i))
        If Not IsEmpty(marked) Then
            If (Not IsBusy(i)) And (sck(i).State <> sckConnected Or DateDiff("s", GetLastActive(i), Now) > 15) Then
                LogLine "Force closing deferred-close socket #" & CStr(i)
                m_closeAfterSend.Remove CStr(i)
                CloseConn i
            End If
        End If
        If sck(i).State = sckConnected Then
            If Not IsSseIndex(i) Then
                If IsBusy(i) Then
                    ' TASK-49: a normal busy skip is SILENT. Only the leaked-flag
                    ' backstop logs, because it closes a socket we believe in flight.
                    If DateDiff("s", m_busySince(i), Now) > m_idleBusyCapSec Then
                        LogLine "Idle-sweep: busy slot #" & CStr(i) & " exceeded hard cap; reclaiming"
                        ClearBusy i
                        CloseConn i
                    End If
                Else
                    Dim last As Date
                    last = GetLastActive(i)
                    If DateDiff("s", last, Now) > IDLE_TIMEOUT_SEC Then
                        LogLine "Closing idle connection #" & CStr(i)
                        CloseConn i
                    End If
                End If
            End If
        End If
    Next
End Sub

' TASK-P1-064: hard-reset one control-array slot. A plain Close() does not reliably move
' a wedged (sckClosing) slot back to sckClosed, so the slot is torn down and re-created;
' the fresh element copies the design-time properties of sck(0) and keeps its Index, so no
' event wiring is lost.
Private Sub ResetSocket(ByVal idx As Integer)
    On Error Resume Next
    If idx <= 0 Then Exit Sub
    ClearBusy idx
    m_conns.Remove CStr(idx)
    m_closeAfterSend.Remove CStr(idx)
    Dim k As Long
    For k = m_sessions.Count To 1 Step -1
        If m_sessions.Item(k) = idx Then m_sessions.Remove k
    Next
    Unload sck(idx)
    Load sck(idx)
    LogLine "Reset socket #" & CStr(idx) & " state=" & CStr(sck(idx).State)
End Sub

Private Function IsSseIndex(ByVal idx As Integer) As Boolean
    On Error Resume Next
    Dim k As Long
    For k = 1 To m_sessions.Count
        If m_sessions.Item(k) = idx Then
            IsSseIndex = True
            Exit Function
        End If
    Next
End Function

Private Function GetLastActive(ByVal idx As Integer) As Date
    On Error Resume Next
    GetLastActive = m_lastActive(CStr(idx))
    If Err.Number <> 0 Then GetLastActive = Now
End Function

Private Sub TouchActive(ByVal idx As Integer)
    On Error Resume Next
    m_lastActive.Remove CStr(idx)
    m_lastActive.Add Now, CStr(idx)
End Sub

' TASK-49: BUSY bookkeeping. MarkBusy/ClearBusy are cheap and idempotent so they can
' be called from every teardown path; IsBusy is what the sweeper consults.
Private Sub MarkBusy(ByVal idx As Integer)
    On Error Resume Next
    If idx < 1 Or idx > MAX_CONNECTIONS Then Exit Sub
    m_busy(idx) = True
    m_busySince(idx) = Now
End Sub

Private Sub ClearBusy(ByVal idx As Integer)
    On Error Resume Next
    If idx < 1 Or idx > MAX_CONNECTIONS Then Exit Sub
    m_busy(idx) = False
End Sub

Private Function IsBusy(ByVal idx As Integer) As Boolean
    On Error Resume Next
    If idx < 1 Or idx > MAX_CONNECTIONS Then Exit Function
    IsBusy = m_busy(idx)
End Function

' TASK-49: run ONE request/response serve. It exists so the caller can clear BUSY on
' BOTH paths (VB6 has no finally) AND so an error that ESCAPES HandleRequest is still
' reported: HandleRequest answers ordinary errors itself with a 500.
' WHY Err IS NOT READ AFTER THIS RETURNS: MEASURED - guarded code inside the serve
' (Collection.Remove on a missing key in MarkCloseAfterSend/TouchActive) leaves Err
' set BY DESIGN, so a caller that inspects Err after a normal return reads a PHANTOM
' error 5 and aborts a response that SendData had already queued.
Private Function ServeRequest(ByVal Index As Integer, ByRef conn As HttpConn) As Boolean
    On Error GoTo ServeFail
    HandleRequest Index, conn
    ServeRequest = True
    Exit Function
ServeFail:
    ServeRequest = False
End Function

' TASK-49: the busy hard cap, overridable for tests (VB6MCP_IDLE_BUSY_CAP_SEC, e.g.
' 35) so the backstop can be exercised without waiting 5 minutes. A leaked BUSY flag
' must never be able to wedge a pool slot forever.
Private Function IdleBusyCapFromEnv() As Long
    On Error Resume Next
    IdleBusyCapFromEnv = IDLE_BUSY_CAP_SEC_DEF
    Dim s As String
    s = Trim$(Environ$("VB6MCP_IDLE_BUSY_CAP_SEC"))
    If Len(s) = 0 Then Exit Function
    If Not IsNumeric(s) Then Exit Function
    Dim v As Long
    v = CLng(Val(s))
    If v > 0 Then IdleBusyCapFromEnv = v
End Function

' ---------- Winsock events ----------

Private Sub sck_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    On Error GoTo ConnFail
    If Index = 0 Then
        Dim i As Integer
        For i = 1 To MAX_CONNECTIONS
            If sck(i).State = sckClosed Then
                sck(i).Accept requestID
                ClearBusy i
                TouchActive i
                ' loopback-only policy: LocalIP is read-only in MSWINSCK.OCX,
                ' so the listener binds all interfaces; reject any peer that
                ' is not 127.0.0.1 right after accepting.
                If sck(i).RemoteHostIP <> "127.0.0.1" Then
                    LogLine "Rejected non-loopback connection from " & sck(i).RemoteHostIP
                    sck(i).Close
                End If
                Exit Sub
            End If
        Next
        Dim dbg As String
        dbg = ""
        Dim j As Integer
        For j = 1 To MAX_CONNECTIONS
            dbg = dbg & CStr(sck(j).State) & ","
        Next
        LogLine "Connection rejected: all sockets busy. states: " & dbg
    End If
    Exit Sub
ConnFail:
    LogLine "ConnectionRequest error: " & CStr(Err.Number) & " " & Err.Description
End Sub

Private Sub sck_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    On Error GoTo DataFail
    TouchActive Index
    Dim b() As Byte
    sck(Index).GetData b, vbArray + vbByte
    Dim conn As HttpConn
    Set conn = GetConn(Index)
    ' do NOT mark SSE here: a plain keep-alive POST must keep its
    ' pipelined bytes via AdvanceToNext; /sse marks explicitly
    conn.AppendBytes b
    If conn.Ready Then
        ' TASK-49: from here until the OUTERMOST return of the serve this slot is IN
        ' FLIGHT and the idle sweeper must not touch it. VB6 has no finally, so the
        ' serve runs through ServeRequest() and BUSY is cleared unconditionally after it.
        ' NB: Err is NOT inspected after a normal return - see ServeRequest.
        MarkBusy Index
        Dim serveOk As Boolean
        serveOk = ServeRequest(Index, conn)
        ClearBusy Index
        TouchActive Index
        If Not serveOk Then
            LogLine "DataArrival error #" & CStr(Index) & ": " & CStr(Err.Number) & " " & Err.Description
            CloseConn Index
        End If
    End If
    Exit Sub
DataFail:
    ClearBusy Index
    LogLine "DataArrival error #" & CStr(Index) & ": " & CStr(Err.Number) & " " & Err.Description
    CloseConn Index
End Sub

Private Sub sck_Close(Index As Integer)
    On Error Resume Next
    ClearBusy Index
    m_conns.Remove CStr(Index)
    Dim k As Long
    For k = m_sessions.Count To 1 Step -1
        If m_sessions.Item(k) = Index Then
            m_sessions.Remove k
        End If
    Next
End Sub

Private Sub sck_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    If Index = 0 Then
        If Not m_listening Then
            ' bind failure - the watchdog timer will advance to the next port
            Exit Sub
        End If
    End If
    LogLine "Winsock error (" & CStr(Index) & "): " & CStr(Number) & " " & Description
    CloseConn Index
End Sub

' ---------- request handling ----------

Private Sub HandleRequest(ByVal Index As Integer, ByVal conn As HttpConn)
    On Error GoTo Failed
    If conn.ErrorText <> "" Then
        SendResponse Index, m_mcp.BuildHttpResponse("413 Request Entity Too Large", "text/plain", "Request Entity Too Large", "", False), False
        CloseConn Index
        Exit Sub
    End If
    Dim method As String
    method = conn.Method
    Dim rawPath As String
    rawPath = conn.Path
    If method = "GET" And Left$(rawPath, 4) = "/sse" Then
        ' open an MCP streamable-HTTP SSE stream
        Dim sessionId As String
        sessionId = NewSessionId()
        RegisterSession sessionId, Index
        conn.MarkSse
        Dim firstEvent As String
        firstEvent = "event: endpoint" & vbCrLf & "data: /message?sessionId=" & sessionId & vbCrLf & vbCrLf
        SendResponse Index, BuildSseHeaders() & firstEvent, True
        Exit Sub
    End If
    If method = "DELETE" And Left$(rawPath, 8) = "/message" Then
        Dim delSession As String
        delSession = QueryParam(rawPath, "sessionId")
        If delSession <> "" Then RemoveSession delSession
        SendResponse Index, m_mcp.BuildHttpResponse("200 OK", "application/json", "", CorsOnly(), False), False
        CloseConn Index
        Exit Sub
    End If
    Dim isMessage As Boolean
    isMessage = (method = "POST" And Left$(rawPath, 8) = "/message")
    Dim sseSession As String
    sseSession = ""
    If isMessage Then
        sseSession = QueryParam(rawPath, "sessionId")
        If sseSession = "" Then
            SendResponse Index, m_mcp.BuildHttpResponse("400 Bad Request", "text/plain", "Missing sessionId", "", False), False
            CloseConn Index
            Exit Sub
        End If
        If Not SessionExists(sseSession) Then
            SendResponse Index, m_mcp.BuildHttpResponse("404 Not Found", "text/plain", "Unknown session", "", False), False
            CloseConn Index
            Exit Sub
        End If
    End If
    Dim response As String
    If isMessage Then
        Dim rpcResp As String
        rpcResp = m_mcp.HandleJsonRpc(Utf8DecodeStr(conn.Body))
        If Len(rpcResp) > 0 Then
            PushSse sseSession, rpcResp
        End If
        response = m_mcp.BuildHttpResponse("202 Accepted", "application/json", "", CorsOnly(), True)
    Else
        response = m_mcp.HandleHttpRequest(conn)
    End If
    SendResponse Index, response, conn.KeepAlive
    If conn.IsSse Then
        conn.ClearAll
    ElseIf conn.KeepAlive Then
        conn.AdvanceToNext
        If conn.Ready Then HandleRequest Index, conn
    Else
        MarkCloseAfterSend Index
    End If
    Exit Sub
Failed:
    SendResponse Index, m_mcp.BuildHttpResponse("500 Internal Server Error", "text/plain", "Internal Server Error: " & Err.Description, "", False), False
    MarkCloseAfterSend Index
End Sub

' ---------- SSE helpers ----------

Private Sub RegisterSession(ByVal sessionId As String, ByVal Index As Integer)
    On Error Resume Next
    m_sessions.Add Index, sessionId
End Sub

Private Function SessionExists(ByVal sessionId As String) As Boolean
    On Error Resume Next
    Dim v As Variant
    v = m_sessions(sessionId)
    SessionExists = (Err.Number = 0) And Not IsEmpty(v)
End Function

Private Sub RemoveSession(ByVal sessionId As String)
    On Error Resume Next
    Dim idx As Variant
    idx = m_sessions(sessionId)
    m_sessions.Remove sessionId
    If Not IsEmpty(idx) Then
        If sck(CInt(idx)).State = sckConnected Then sck(CInt(idx)).Close
        m_conns.Remove CStr(CInt(idx))
    End If
End Sub

Private Sub PushSse(ByVal sessionId As String, ByVal jsonText As String)
    On Error Resume Next
    Dim idx As Variant
    idx = m_sessions(sessionId)
    If IsEmpty(idx) Then Exit Sub
    Dim Index As Integer
    Index = CInt(idx)
    If sck(Index).State <> sckConnected Then
        m_sessions.Remove sessionId
        Exit Sub
    End If
    Dim ev As String
    ev = "event: message" & vbCrLf & "data: " & jsonText & vbCrLf & vbCrLf
    Dim bytes() As Byte
    bytes = StrToBytes(ev)
    sck(Index).SendData bytes
End Sub

Private Function BuildSseHeaders() As String
    BuildSseHeaders = "HTTP/1.1 200 OK" & vbCrLf & _
        "Content-Type: text/event-stream" & vbCrLf & _
        "Cache-Control: no-cache" & vbCrLf & _
        "Connection: keep-alive" & vbCrLf & _
        CorsOnly() & vbCrLf
End Function

Private Function CorsOnly() As String
    CorsOnly = "Access-Control-Allow-Origin: *"
End Function

Private Function NewSessionId() As String
    On Error Resume Next
    Dim g As Object
    Set g = CreateObject("Scriptlet.TypeLib")
    ' the GUID is a fixed buffer: keep only up to the closing brace
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

' ---------- connection bookkeeping ----------

Private Function GetConn(ByVal Index As Integer) As HttpConn
    On Error Resume Next
    Dim c As HttpConn
    Set c = m_conns(CStr(Index))
    On Error GoTo 0
    If c Is Nothing Then
        Set c = New HttpConn
        m_conns.Add c, CStr(Index)
    End If
    Set GetConn = c
End Function

Private Sub sck_SendComplete(Index As Integer)
    On Error Resume Next
    m_closeAfterSend.Remove CStr(Index)
    CloseConn Index
End Sub

Private Sub MarkCloseAfterSend(ByVal Index As Integer)
    On Error Resume Next
    m_closeAfterSend.Remove CStr(Index)
    m_closeAfterSend.Add True, CStr(Index)
End Sub

Private Sub CloseConn(ByVal Index As Integer)
    If Index <= 0 Then Exit Sub
    On Error Resume Next
    ClearBusy Index
    LogLine "CloseConn #" & CStr(Index) & " state before=" & CStr(sck(Index).State)
    sck(Index).Close
    m_conns.Remove CStr(Index)
End Sub

Private Sub SendResponse(ByVal Index As Integer, ByVal response As String, ByVal keepAlive As Boolean)
    Dim bytes() As Byte
    bytes = StrToBytes(response)
    On Error Resume Next
    Dim nBytes As Long
    nBytes = UBound(bytes) - LBound(bytes) + 1
    If Err.Number <> 0 Or nBytes <= 0 Then Exit Sub
    On Error GoTo 0
    On Error Resume Next
    sck(Index).SendData bytes
    If Not m_mcp Is Nothing Then
        If m_mcp.ConsumeQuit() Then
            ' deferred IDE quit: response has been sent
            m_quitArmed = True
            tmrWatch.Interval = 150
            tmrWatch.Enabled = True
        End If
        If m_mcp.ConsumeRestart() Then
            ' deferred restart: wait for the run-mode teardown
            m_restartArmed = True
            tmrWatch.Interval = 500
            tmrWatch.Enabled = True
        End If
    End If
End Sub

Private Sub Form_Terminate()
    ' TASK-48: belt-and-braces deregistration. StopListening already deletes our
    ' <pid>.json on a clean exit; this covers a teardown that skips it. The delete
    ' is for OUR OWN pid and DeleteInstanceFile is idempotent (FileExists check +
    ' swallowed errors), so a double delete is harmless and no other instance's
    ' registration can be touched.
    On Error Resume Next
    DeleteInstanceFile FormProcessId()
End Sub
