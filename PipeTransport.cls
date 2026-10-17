VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PipeTransport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ===========================================================================
' (ADR docs/design/TRANSPORT_NP_DUAL.md section 0 / 7): the named-pipe
' server end of the add-in. One instance per VB6 process, serviced from the
' SAME UI/EDT thread and the SAME message pump the TCP transports already use
' (ServerForm.tmrWatch / WinsockTransport's Win32 timer call Pipe.Tick).
'
' HARD RULES THIS FILE OBEYS (each one is a measured failure in the
' spike report, docs/evidence/np_spike/REPORT.md):
'   1. NO WSAStartup / WSACleanup anywhere in this file, and no ws2_32 declare.
'      WSACleanup is PROCESS-WIDE: in the spike it killed the HTTP listener of
'      the product add-in running in the same VB6.EXE (spike trap 6). The pipe
'      is kernel32-only, which is also why the VB side can occupy no TCP port
'      (ADR section 0.2 rule 1).
'   2. Framing is 4-byte little-endian length prefix + payload, with
'      PIPE_TYPE_BYTE / PIPE_READMODE_BYTE. Message mode does NOT remove the
'      need to de-frame: the spike measured one ReadFile receiving a single
'      message holding THREE frames (msgSticky=4). Byte mode + PeekNamedPipe
'      is deterministic: we only ever read the bytes the kernel says are there.
'   3. nMaxInstances = 1 (ADR section 7.5 model A). A second client gets
'      ERROR_PIPE_BUSY(231) from its own CreateFile/WaitNamedPipe. That refusal
'      happens on the CLIENT side, so this server never sees it - the visible
'      busy/fallback marker therefore lives in the hub receipt, and this file
'      reports the server-side half (connect/teardown counters + a log line).
'   4. EVERY call in the servicing path is bounded. No blocking ReadFile, no
'      blocking ConnectNamedPipe (both are FILE_FLAG_OVERLAPPED and polled with
'      GetOverlappedResult(bWait = 0)), FlushFileBuffers is deliberately NEVER
'      called here because it blocks until the client has read the data, and
'      every wait has a deadline: hello (VB6MCP_PIPE_HELLO_MS), a partially
'      received frame (VB6MCP_PIPE_READ_DEADLINE_MS), an idle connection
'      (VB6MCP_PIPE_IDLE_MS).
'   5. GetOverlappedResult(bWait=0) reports ERROR_IO_INCOMPLETE(996), NOT
'      ERROR_IO_PENDING(997) (spike trap 3). Both are accepted everywhere, and
'      a 0 error code is ALSO treated as "not finished yet" rather than as a
'      fatal error (spike trap 4: recv returned SOCKET_ERROR with err=0).
'   6. The receive accumulator is a FLAT 1-D Byte array, never a 2-D array.
'      VB6 arrays are COLUMN-MAJOR, so using arr(i, 0) as a raw buffer pointer
'      interleaves the buffers of two pipes (spike trap 1 - it cost the spike
'      several rounds). There is exactly one pipe here, so the whole class of
'      bug is structurally absent; the flat array is also what lets CopyMemory
'      (RtlMoveMemory, memmove semantics) extract and shift frames in one call.
'
' NOT DONE / NOT CLAIMED (must not be read as verified):
'   * The progress frame of ADR section 7.4 class 3 is emitted, but only while
'     a tool call is actually running AND that call pumps messages through the
'     existing Util.MaybePump -> DoEvents model. It is NOT a timer thread and
'     it cannot fire during a call that never pumps.
'   * The explicit SDDL ACL of ADR section 5.1 is ATTEMPTED (see
'     BuildSecurityDescriptor) and falls back to the process default DACL with
'     a loud log line when anything about it fails. The access-control effect
'     itself was NOT verified on the test bed (spike section 9.3 lists it as
'     unverified for the same reason). Set VB6MCP_PIPE_ACL=0 to skip it.
'   * PIPE_TYPE_MESSAGE and the ERROR_MORE_DATA(234) branch are NOT used at all
'     (byte mode was chosen by the spike), so that path cannot be exercised.
' ===========================================================================

' --- framing / sizes -------------------------------------------------------

' ADR section 7.3: one frame is at most 8 MiB; an over-long frame is refused
' loudly (deny frame + teardown), never silently truncated.
Private Const NP_MAX_FRAME As Long = 8388608
Private Const NP_ACC_CAP As Long = NP_MAX_FRAME + 4
Private Const NP_PIPE_BUF As Long = 262144
Private Const NP_DEFAULT_TICK_MS As Long = 15
Private Const NP_DEFAULT_IDLE_MS As Long = 120000
Private Const NP_DEFAULT_READ_DEADLINE_MS As Long = 2000
Private Const NP_DEFAULT_HELLO_MS As Long = 5000
Private Const NP_DEFAULT_PROGRESS_MS As Long = 2000
Private Const NP_CREATE_RETRY_MS As Long = 250
' Grace window between "the deny frame is finished on the wire" and closing the
' handle. DisconnectNamedPipe was MEASURED (t16) to lose the buffered deny frame:
' the client saw EOF with no frame at all, i.e. the reason was silently dropped.
' The handle therefore survives until the client has read it, or until this
' deadline expires, and DisconnectNamedPipe is never used on this path.
Private Const NP_FLUSH_GRACE_MS As Long = 750

' --- win32 constants (kernel32 only; deliberately no ws2_32 here) -----------

Private Const NP_INVALID_HANDLE As Long = -1
Private Const NP_FILE_FLAG_OVERLAPPED As Long = &H40000000
Private Const NP_PIPE_ACCESS_DUPLEX As Long = &H3
Private Const NP_PIPE_TYPE_BYTE As Long = &H0
Private Const NP_PIPE_READMODE_BYTE As Long = &H0
Private Const NP_PIPE_WAIT As Long = &H0
Private Const NP_ERROR_IO_PENDING As Long = 997
Private Const NP_ERROR_IO_INCOMPLETE As Long = 996
Private Const NP_ERROR_PIPE_CONNECTED As Long = 535
Private Const NP_ERROR_PIPE_BUSY As Long = 231
Private Const NP_ERROR_NO_DATA As Long = 232
Private Const NP_ERROR_PIPE_NOT_CONNECTED As Long = 233
Private Const NP_ERROR_BROKEN_PIPE As Long = 109
Private Const NP_TOKEN_USER_CLASS As Long = 1
Private Const NP_TOKEN_QUERY As Long = &H8
Private Const NP_SDDL_REVISION_1 As Long = 1

' EVERY kernel32/advapi32 entry below carries an explicit Alias. A Declare whose
' VB name differs from the exported name WITHOUT an Alias compiles fine and then
' fails at RUN time with VB error 453 "cannot find entry point <name> in
' kernel32" - measured on t16 on the first call to NpGetCurrentProcessId. That is
' a silent-until-called trap, so the alias is mandatory here, not stylistic.
Private Type NpOverlapped
    Internal As Long
    InternalHigh As Long
    Offset As Long
    OffsetHigh As Long
    hEvent As Long
End Type

Private Type NpSecurityAttributes
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

Private Declare Function NpCreateNamedPipe Lib "kernel32" Alias "CreateNamedPipeA" (ByVal lpName As String, ByVal dwOpenMode As Long, ByVal dwPipeMode As Long, ByVal nMaxInstances As Long, ByVal nOutBufferSize As Long, ByVal nInBufferSize As Long, ByVal nDefaultTimeOut As Long, ByVal lpSecurityAttributes As Long) As Long
Private Declare Function NpConnectNamedPipe Lib "kernel32" Alias "ConnectNamedPipe" (ByVal hNamedPipe As Long, ByRef lpOverlapped As NpOverlapped) As Long
Private Declare Function NpDisconnectNamedPipe Lib "kernel32" Alias "DisconnectNamedPipe" (ByVal hNamedPipe As Long) As Long
Private Declare Function NpPeekNamedPipe Lib "kernel32" Alias "PeekNamedPipe" (ByVal hNamedPipe As Long, ByVal lpBuffer As Long, ByVal nBufferSize As Long, ByRef lpBytesRead As Long, ByRef lpTotalBytesAvail As Long, ByRef lpBytesLeftThisMessage As Long) As Long
Private Declare Function NpReadFile Lib "kernel32" Alias "ReadFile" (ByVal hFile As Long, ByRef lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, ByRef lpNumberOfBytesRead As Long, ByRef lpOverlapped As NpOverlapped) As Long
Private Declare Function NpWriteFile Lib "kernel32" Alias "WriteFile" (ByVal hFile As Long, ByRef lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, ByRef lpNumberOfBytesWritten As Long, ByRef lpOverlapped As NpOverlapped) As Long
Private Declare Function NpGetOverlappedResult Lib "kernel32" Alias "GetOverlappedResult" (ByVal hFile As Long, ByRef lpOverlapped As NpOverlapped, ByRef lpNumberOfBytesTransferred As Long, ByVal bWait As Long) As Long
Private Declare Function NpCreateEvent Lib "kernel32" Alias "CreateEventA" (ByVal lpEventAttributes As Long, ByVal bManualReset As Long, ByVal bInitialState As Long, ByVal lpName As Long) As Long
Private Declare Function NpSetEvent Lib "kernel32" Alias "SetEvent" (ByVal hEvent As Long) As Long
Private Declare Function NpResetEvent Lib "kernel32" Alias "ResetEvent" (ByVal hEvent As Long) As Long
Private Declare Function NpCloseHandle Lib "kernel32" Alias "CloseHandle" (ByVal hObject As Long) As Long
Private Declare Function NpCancelIo Lib "kernel32" Alias "CancelIo" (ByVal hFile As Long) As Long
Private Declare Function NpGetTickCount Lib "kernel32" Alias "GetTickCount" () As Long
Private Declare Function NpGetCurrentProcessId Lib "kernel32" Alias "GetCurrentProcessId" () As Long
Private Declare Function NpProcessIdToSessionId Lib "kernel32" Alias "ProcessIdToSessionId" (ByVal dwProcessId As Long, ByRef pSessionId As Long) As Long
Private Declare Function NpGetCurrentProcess Lib "kernel32" Alias "GetCurrentProcess" () As Long
Private Declare Function NpOpenProcessToken Lib "advapi32" Alias "OpenProcessToken" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, ByRef TokenHandle As Long) As Long
Private Declare Function NpGetTokenInformation Lib "advapi32" Alias "GetTokenInformation" (ByVal TokenHandle As Long, ByVal TokenInformationClass As Long, ByRef TokenInformation As Any, ByVal TokenInformationLength As Long, ByRef ReturnLength As Long) As Long
Private Declare Function NpConvertSidToStringSid Lib "advapi32" Alias "ConvertSidToStringSidA" (ByVal Sid As Long, ByRef StringSid As Long) As Long
Private Declare Function NpConvertStringSecurityDescriptorToSecurityDescriptor Lib "advapi32" Alias "ConvertStringSecurityDescriptorToSecurityDescriptorA" (ByVal StringSecurityDescriptor As String, ByVal StringSDRevision As Long, ByRef SecurityDescriptor As Long, ByVal SecurityDescriptorSize As Long) As Long
Private Declare Function NpLocalFree Lib "kernel32" Alias "LocalFree" (ByVal hMem As Long) As Long
Private Declare Function NpLstrLenA Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Long) As Long
Private Declare Function NpRtlGenRandom Lib "advapi32" Alias "SystemFunction036" (ByVal RandomBuffer As Long, ByVal RandomBufferLength As Long) As Long
Private Declare Sub NpCopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Sub NpCopyMemoryToLong Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Long, ByVal Source As Long, ByVal Length As Long)
Private Declare Sub NpCopyMemoryFromLong Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)

' --- per-instance state ----------------------------------------------------

Private m_mcp As McpServer
Private m_active As Boolean
Private m_mode As String                 ' "pipe" | "auto" | "tcp"
Private m_tcpEnabled As Boolean           ' is the TCP leg of the SAME instance live?
Private m_pipeName As String              ' \\.\pipe\vb6mcp-<sid>-<pid>-<slug>-<nonce16>
Private m_nonce As String                 ' 16 hex chars, RtlGenRandom, never persisted outside the registry
Private m_instanceId As String            ' <pid>-<nonce>
Private m_sid As Long                     ' Windows session id
Private m_pid As Long
Private m_slug As String
Private m_startedAt As String

Private m_hPipe As Long
Private m_ov As NpOverlapped
Private m_state As Long                   ' 0=create 1=connect-pending 2=connected 3=read-pending 4=write-pending
Private m_acc() As Byte
Private m_accLen As Long
Private m_out() As Byte
Private m_outLen As Long
Private m_outSent As Long
Private m_outQ As Collection
Private m_authed As Boolean
Private m_closeAfterFlush As Boolean
Private m_closeAfterFlushAt As Long
Private m_aclReason As String
Private m_inTick As Boolean
Private m_serving As Boolean
Private m_tears As Long
Private m_createFails As Long
Private m_lastCreateAt As Long
Private m_connAt As Long
Private m_idleAt As Long
Private m_frameAt As Long
Private m_pendingAt As Long
Private m_conns As Long
Private m_frames As Long
Private m_bytes As Long
Private m_errs As Long
Private m_lastReason As String
Private m_curTool As String
Private m_serveStart As Long
Private m_progSeq As Long
Private m_lastProgAt As Long
' diagnostic (Lead-authorized narrow round): count every nested re-entry so
' "was the pipe ever ticked while a tool call was running?" is a FACT in the log,
' not a guess.
Private m_nestedTicks As Long
' the REAL reason the last instance publish failed.  PublishInstance runs under
' On Error Resume Next (it always has), so without this the retry loop could report only the
' symptom ("the instance file is missing") and never the cause.  Measured with the injection
' test: the cause was a path whose parent was a FILE, so "cannot write" and "cannot write
' BECAUSE the path is not a directory" differ exactly here.
Private m_lastPublishErr As String
Private m_progFrames As Long
Private m_nestedLogged As Long
Private m_useAcl As Boolean

' --- knobs -----------------------------------------------------------------

Private m_tickMs As Long
Private m_idleMs As Long
Private m_readDeadlineMs As Long
Private m_helloMs As Long
Private m_progressMs As Long
Private m_progressOn As Boolean

' --- public entry points ---------------------------------------------------

' transport switch. ADR section 0.2 rule 3: pipe | tcp | auto, DEFAULT
' pipe. "winsock" is the pre-existing value for the pure-Winsock TCP track and
' is kept working (it maps onto tcp).
Public Function ResolveMode() As String
    Dim s As String
    s = LCase$(Trim$(Environ$("VB6MCP_TRANSPORT")))
    If Len(s) = 0 Then
        ' wrapper: WScript.Shell registry reads were unreliable here.
        s = LCase$(Trim$(WinOps.McpSettingGet("Transport")))
    End If
    ResolveMode = NormalizeMode(s)
End Function

Public Function NormalizeMode(ByVal s As String) As String
    Select Case LCase$(Trim$(s))
        Case "pipe"
            NormalizeMode = "pipe"
        Case "auto"
            NormalizeMode = "auto"
        Case "tcp", "winsock", "http", "ocx"
            NormalizeMode = "tcp"
        Case Else
            ' ADR section 0.2 rule 3: pipe is the new default.
            NormalizeMode = "pipe"
    End Select
End Function

' The timer period the HOST transport must run at so the pipe stays responsive.
' The TCP transports used 5000 ms once listening; that would make a pipe call
' take up to five seconds, so the host asks this instead.
Public Function TickMs() As Long
    TickMs = m_tickMs
End Function

Public Function Active() As Boolean
    Active = m_active
End Function

Public Function TcpEnabled() As Boolean
    TcpEnabled = m_tcpEnabled
End Function

Public Function PipeName() As String
    PipeName = m_pipeName
End Function

' Redacted name for the log: ADR section 5.2 keeps the full nonce out of the
' log/terminal history (the registry file is the only place it is written).
Public Function PipeNameRedacted() As String
    If Len(m_pipeName) <= 4 Then
        PipeNameRedacted = m_pipeName
    Else
        PipeNameRedacted = Left$(m_pipeName, Len(m_pipeName) - 16) & Left$(m_nonce, 4) & "..."
    End If
End Function

Public Function NoncePrefix() As String
    NoncePrefix = Left$(m_nonce, 4)
End Function

Public Function InstanceId() As String
    InstanceId = m_instanceId
End Function

Public Function SessionIdValue() As Long
    SessionIdValue = m_sid
End Function

Public Sub StartServer(ByVal mcp As McpServer, ByVal mode As String, ByVal tcpEnabled As Boolean)
    On Error GoTo Fail
    Set m_mcp = mcp
    m_mode = NormalizeMode(mode)
    m_tcpEnabled = tcpEnabled
    m_pid = NpGetCurrentProcessId()
    m_sid = SessionOfProcess(m_pid)
    m_startedAt = Format$(Now, "yyyy-mm-dd hh:nn:ss")
    If Len(m_nonce) <> 16 Then m_nonce = MakeNonce16()
    m_instanceId = CStr(m_pid) & "-" & m_nonce
    m_slug = Slugify(mcp.ProjectName())
    If Len(m_slug) = 0 Then m_slug = "noproj"
    ' ADR section 3.2 name: \\.\pipe\vb6mcp-<sid>-<pid>-<slug>-<nonce16>.
    ' The nonce is what makes a stale/reused pid unable to name the same pipe
    ' (ADR section 3.5); sid isolates Windows sessions; slug is human-readable
    ' only and plays no part in any security decision.
    m_pipeName = "\\.\pipe\vb6mcp-" & CStr(m_sid) & "-" & CStr(m_pid) & "-" & m_slug & "-" & m_nonce
    m_tickMs = Util.EnvLongOr("VB6MCP_PIPE_TICK_MS", NP_DEFAULT_TICK_MS)
    m_idleMs = Util.EnvLongOr("VB6MCP_PIPE_IDLE_MS", NP_DEFAULT_IDLE_MS)
    m_readDeadlineMs = Util.EnvLongOr("VB6MCP_PIPE_READ_DEADLINE_MS", NP_DEFAULT_READ_DEADLINE_MS)
    m_helloMs = Util.EnvLongOr("VB6MCP_PIPE_HELLO_MS", NP_DEFAULT_HELLO_MS)
    m_progressMs = Util.EnvLongOr("VB6MCP_PIPE_PROGRESS_MS", NP_DEFAULT_PROGRESS_MS)
    m_progressOn = EnvFlag("VB6MCP_PIPE_PROGRESS", True)
    m_useAcl = EnvFlag("VB6MCP_PIPE_ACL", True)
    ReDim m_acc(0 To NP_ACC_CAP - 1)
    ReDim m_out(0 To 3)
    Set m_outQ = New Collection
    m_state = 0
    m_active = True
    NpLog "transport=pipe event=start mode=" & m_mode & " tcpEnabled=" & BoolText(m_tcpEnabled) & " pipe=" & PipeNameRedacted() & " nonce=" & NoncePrefix() & "... pid=" & CStr(m_pid) & " session=" & CStr(m_sid) & " tickMs=" & CStr(m_tickMs) & " acl=" & CStr(m_useAcl)
    Exit Sub
Fail:
    m_active = False
    NpLog "transport=pipe event=start_failed err=" & CStr(Err.Number) & " " & Err.Description
End Sub

Public Sub StopServer()
    On Error Resume Next
    m_active = False
    If m_hPipe <> 0 Then Teardown "stop"
    If m_ov.hEvent <> 0 Then
        NpCloseHandle m_ov.hEvent
        m_ov.hEvent = 0
    End If
    Set m_outQ = New Collection
    NpLog "transport=pipe event=stop frames=" & CStr(m_frames) & " bytes=" & CStr(m_bytes) & " conns=" & CStr(m_conns) & " errors=" & CStr(m_errs) & " tears=" & CStr(m_tears) & " nestedTicks=" & CStr(m_nestedTicks) & " progressFrames=" & CStr(m_progFrames)
    Set m_mcp = Nothing
End Sub

' One bounded servicing pass. The host (ServerForm's tmrWatch / the Winsock
' transport's Win32 timer) calls this from the UI thread only.
Public Sub Tick()
    If Not m_active Then Exit Sub
    If m_inTick Then
        ' NESTED RE-ENTRY. This is reached only through the existing pump model:
        ' a long tool call runs Util.MaybePump -> DoEvents, which dispatches the
        ' host timer again while we are still inside ServeRpc. NEVER process
        ' frames here (that would re-enter HandleJsonRpc underneath itself);
        ' only emit the ADR section 7.4 progress frame.
        m_nestedTicks = m_nestedTicks + 1
        MaybeProgress
        Exit Sub
    End If
    m_inTick = True
    On Error GoTo Done
    Dim guard As Long
    guard = 0
    Do
        guard = guard + 1
        If guard > 128 Then Exit Do
        Select Case m_state
            Case 0
                If Not EnsurePipe() Then Exit Do
            Case 1
                If Not PollConnect() Then Exit Do
            Case 2
                If Not ServiceConnected() Then Exit Do
            Case 3
                If Not PollRead() Then Exit Do
            Case 4
                If Not PollWrite() Then Exit Do
            Case Else
                Exit Do
        End Select
    Loop
Done:
    m_inTick = False
End Sub

Public Function StatsJson() As String
    StatsJson = "{""transport"":""pipe"",""mode"":""" & JsonEsc(m_mode) & """,""active"":" & BoolText(m_active) _
        & ",""tcpEnabled"":" & BoolText(m_tcpEnabled) & ",""pipe"":""" & JsonEsc(PipeNameRedacted()) _
        & """,""nonce4"":""" & NoncePrefix() & """,""instanceId"":""" & JsonEsc(m_instanceId) _
        & """,""pid"":" & CStr(m_pid) & ",""session"":" & CStr(m_sid) _
        & ",""conns"":" & CStr(m_conns) & ",""frames"":" & CStr(m_frames) _
        & ",""bytes"":" & CStr(m_bytes) & ",""errors"":" & CStr(m_errs) _
        & ",""teardowns"":" & CStr(m_tears) & ",""nestedTicks"":" & CStr(m_nestedTicks) & ",""progressFrames"":" & CStr(m_progFrames) & ",""lastReason"":""" & JsonEsc(m_lastReason) & """}"
End Function

' ADR section 4.1: the registration entry the hub pairs against. It is written
' HERE rather than through Util.WriteInstanceFile because the pipe identity
' triple (pipe/sessionId/nonce/instanceId) has to be added to the same file and
' the legacy writer must stay byte-for-byte unchanged for the TCP track.
' ServerForm owns the retry loop but NOT the transport log; this hook lets the
' loop write registration_retry / registration_failed / registration_ok into the same log
' the transport already uses, so the accounting is READABLE (a counter nobody can read is
' not accounting).
Public Function LastPublishError() As String
    LastPublishError = m_lastPublishErr
End Function

Public Sub LogRegistration(ByVal msg As String)
    NpLog msg
End Sub

Public Sub PublishInstance(ByVal port As Long, ByVal tcpEnabled As Boolean, ByVal project As String, ByVal projectPath As String, ByVal dllSha As String, ByVal startedAt As String, Optional ByVal hwnd As Long = 0, Optional ByVal tid As Long = 0, Optional ByVal epoch As Long = 0, Optional ByVal ingressOn As Boolean = False, Optional ByVal ingressHwnd As Long = 0, Optional ByVal ingressBoxHwnd As Long = 0, Optional ByVal ingressCaption As String = "")
    On Error Resume Next
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso Is Nothing Then Exit Sub
    EnsureFolder fso, Util.Vb6mcpDir()
    EnsureFolder fso, Util.InstancesDir()
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("pid") = m_pid
    d("port") = port
    d("transport") = "pipe"
    d("pipe") = m_pipeName
    d("sessionId") = m_sid
    d("nonce") = m_nonce
    d("instanceId") = m_instanceId
    d("tcpEnabled") = tcpEnabled
    d("pipeProto") = "np/1"
    d("project") = project
    d("projectPath") = projectPath
    d("dllSha256") = dllSha
    If Len(startedAt) = 0 Then
        d("startedAt") = Format$(Now, "yyyy-mm-dd hh:nn:ss")
    Else
        d("startedAt") = startedAt
    End If
    d("startedByHub") = (Environ$("VB6MCP_STARTED_BY_HUB") = "1")
    ' the self-registration record the bridge needs to bind to the INGRESS window
    ' (textbox) without guessing: the window handle, its thread, and a monotonic
    ' epoch that makes a stale record distinguishable from a fresh one.  Written by the same
    ' attempt that publishes pid/pipe/nonce, so retry and HWND reporting are one pass.
    d("hwnd") = hwnd
    d("tid") = tid
    d("epoch") = epoch
    ' (step 2): the OPTIONAL textbox-ingress identity. Written ONLY when the
    ' ingress switch is on, so the default (pipe) instance file keeps exactly the keys
    ' it has today - the fallback path must not visibly drift just because a new field
    ' exists.  These are Additions on the textbox track, not a widening of the default.
    If ingressOn Then
        d("ingressMode") = "textbox"
        d("ingressHwnd") = ingressHwnd
        d("ingressBoxHwnd") = ingressBoxHwnd
        d("ingressCaption") = ingressCaption
    End If
    Dim ts As Object
    m_lastPublishErr = ""
    Err.Clear
    Set ts = fso.CreateTextFile(Util.InstancesDir() & "\" & CStr(m_pid) & ".json", True)
    ' name the cause (0x46 permission, 0x4B/0x4C path shapes) instead of letting the
    ' caller infer it from a missing file.
    If Err.Number <> 0 Then
        m_lastPublishErr = "CreateTextFile 0x" & Hex$(Err.Number) & " " & Err.Description
        Err.Clear
        Exit Sub
    End If
    ' same ASCII-only rule as Util.WriteInstanceFile (pipe is the DEFAULT track,
    ' so without this the GBK defect would survive on the default path).
    ts.WriteLine Util.AsciiEscapeJson(JsonStringify(d))
    ts.Close
End Sub

' --- state 0: create the pipe instance -------------------------------------

Private Function EnsurePipe() As Boolean
    EnsurePipe = False
    Dim now As Long
    now = NpGetTickCount()
    If m_lastCreateAt <> 0 Then
        If (now - m_lastCreateAt) < NP_CREATE_RETRY_MS Then Exit Function
    End If
    m_lastCreateAt = now
    If m_ov.hEvent = 0 Then
        m_ov.hEvent = NpCreateEvent(0, 1, 0, 0)
        If m_ov.hEvent = 0 Then
            NpLog "transport=pipe event=createfail reason=no_event"
            Exit Function
        End If
    End If
    Dim sa As NpSecurityAttributes
    Dim pSA As Long
    Dim pSD As Long
    pSA = 0
    pSD = 0
    If m_useAcl Then
        m_aclReason = ""
        pSD = BuildSecurityDescriptor(m_aclReason)
        If pSD <> 0 Then
            sa.nLength = Len(sa)
            sa.lpSecurityDescriptor = pSD
            sa.bInheritHandle = 0
            pSA = VarPtr(sa)
        Else
            ' LOUD, not silent, and it names the step that failed: a silent ACL
            ' fallback would look identical to a working one.
            NpLog "transport=pipe event=acl result=fallback_default_dacl reason=" & m_aclReason & " (ADR 5.1 SDDL not applied)"
        End If
    End If
    ' PIPE_TYPE_BYTE | PIPE_READMODE_BYTE | PIPE_WAIT, nMaxInstances = 1.
    Dim h As Long
    h = NpCreateNamedPipe(m_pipeName, NP_PIPE_ACCESS_DUPLEX Or NP_FILE_FLAG_OVERLAPPED, _
        NP_PIPE_TYPE_BYTE Or NP_PIPE_READMODE_BYTE Or NP_PIPE_WAIT, _
        1, NP_PIPE_BUF, NP_PIPE_BUF, 0, pSA)
    Dim e As Long
    e = Err.LastDllError
    If pSD <> 0 Then NpLocalFree pSD
    If h = NP_INVALID_HANDLE Then
        m_createFails = m_createFails + 1
        If m_createFails <= 3 Or (m_createFails Mod 20) = 0 Then
            NpLog "transport=pipe event=createfail err=" & CStr(e) & " fails=" & CStr(m_createFails)
        End If
        Exit Function
    End If
    m_hPipe = h
    m_createFails = 0
    ResetOv
    ' Overlapped connect: it returns at once (ERROR_IO_PENDING) and is polled in
    ' state 1. ERROR_PIPE_CONNECTED means a client was already waiting.
    Dim ok As Long
    ok = NpConnectNamedPipe(h, m_ov)
    e = Err.LastDllError
    If ok <> 0 Then
        OnConnected
        EnsurePipe = True
    ElseIf e = NP_ERROR_IO_PENDING Or e = NP_ERROR_IO_INCOMPLETE Then
        m_state = 1
    ElseIf e = NP_ERROR_PIPE_CONNECTED Then
        OnConnected
        EnsurePipe = True
    Else
        NpLog "transport=pipe event=connectfail err=" & CStr(e)
        NpCloseHandle h
        m_hPipe = 0
        m_state = 0
    End If
End Function

' --- state 1: overlapped ConnectNamedPipe completion ----------------------

Private Function PollConnect() As Boolean
    PollConnect = False
    Dim ok As Long, br As Long, e As Long
    ok = NpGetOverlappedResult(m_hPipe, m_ov, br, 0)
    e = Err.LastDllError
    If ok <> 0 Then
        OnConnected
        PollConnect = True
    ElseIf e = NP_ERROR_IO_PENDING Or e = NP_ERROR_IO_INCOMPLETE Then
        Exit Function
    ElseIf e = NP_ERROR_PIPE_CONNECTED Then
        OnConnected
        PollConnect = True
    Else
        NpLog "transport=pipe event=connectfail err=" & CStr(e)
        Teardown "connect_error"
    End If
End Function

Private Sub OnConnected()
    m_state = 2
    m_conns = m_conns + 1
    m_connAt = NpGetTickCount()
    m_idleAt = m_connAt
    m_accLen = 0
    m_authed = False
    m_outLen = 0
    m_outSent = 0
    m_closeAfterFlush = False
    Set m_outQ = New Collection
    ' (step 3): a new authenticated connection means a new doorbell seq space.
    Ingress.SessionReset
    NpLog "transport=pipe event=connect conns=" & CStr(m_conns) & " (handshake deadline " & CStr(m_helloMs) & " ms)"
End Sub

' --- state 2: connected ---------------------------------------------------

' =====================================================================
' (step 3): the TEXTBOX DOORBELL service path.
'
' The request does NOT arrive on the pipe: it is queued by the IngressForm Change
' handler and handed in HERE from the watchdog tick. The reply is written by the
' ordinary QueueOut path onto the ALREADY AUTHENTICATED connection the bridge is
' holding open, so the wire protocol is unchanged - the request simply arrives
' out of band.
' =====================================================================
Public Function IngressReady() As Boolean
    IngressReady = False
    If m_hPipe = 0 Then Exit Function
    If Not m_authed Then Exit Function
    ' Never re-enter a tool call: a nested tick (DoEvents inside a long tool) must
    ' not start a second request on the same connection.
    If m_serving Then Exit Function
    If m_state <> 2 And m_state <> 3 Then Exit Function
    ' A partially received pipe frame means the client is talking on the pipe too;
    ' serving out of band at that moment would interleave two requests.
    If m_accLen <> 0 Then Exit Function
    IngressReady = True
End Function

Public Sub IngressServe(ByVal reqText As String)
    If Not IngressReady() Then Exit Sub
    ServeRpc reqText
End Sub

' The per-process pairing credential (16 hex chars) the payload of the textbox
' doorbell must carry. It is the SAME value the pipe handshake uses, so there is
' one identity and one compare - the caption's 12-char locator only FINDS the
' window, it never authorises anything.
Public Function NonceValue() As String
    NonceValue = m_nonce
End Function

Private Function ServiceConnected() As Boolean
    ServiceConnected = False
    Dim now As Long
    now = NpGetTickCount()

    ' (a) handshake deadline: a squatter may not hold the single instance
    '     (nMaxInstances = 1) without presenting hello.
    If Not m_authed Then
        If (now - m_connAt) > m_helloMs Then
            NpLog "transport=pipe result=teardown reason=hello_timeout"
            Teardown "hello_timeout"
            Exit Function
        End If
    End If

    ' (b) flush whatever is queued before reading more. A deny frame closes the
    '     connection only AFTER it has been read by the client (or after a bounded
    '     grace window), so the refusal reason always reaches the caller.
    If m_outLen = 0 Then
        If m_outQ.Count > 0 Then
            BeginNextOut
        ElseIf m_closeAfterFlush Then
            If m_closeAfterFlushAt = 0 Then
                m_closeAfterFlushAt = now
            End If
            If NpPeerGone() Or (now - m_closeAfterFlushAt) >= NP_FLUSH_GRACE_MS Then
                Teardown "deny_sent"
                Exit Function
            End If
        End If
    End If
    If m_outLen > 0 Then
        If NpStartWrite() Then ServiceConnected = True
        Exit Function
    End If

    ' (c) bounded read deadline for a PARTIALLY received frame (ADR section 7.4
    '     S ruling 2). It constrains bytes only - never the tool execution.
    If m_accLen > 0 Then
        If (now - m_frameAt) > m_readDeadlineMs Then
            NpLog "transport=pipe result=teardown reason=read_deadline_exceeded accLen=" & CStr(m_accLen)
            Teardown "read_deadline_exceeded"
            Exit Function
        End If
    End If

    ' (d) non-blocking availability probe, then a read of EXACTLY the bytes the
    '     kernel reports. Never a blocking ReadFile.
    Dim pr As Long, avail As Long, lft As Long, ok As Long
    ok = NpPeekNamedPipe(m_hPipe, 0, 0, pr, avail, lft)
    If ok = 0 Then
        Dim pe As Long
        pe = Err.LastDllError
        If pe = NP_ERROR_BROKEN_PIPE Or pe = NP_ERROR_PIPE_NOT_CONNECTED Then
            ' The CLIENT closed its end - a normal end of session, not a failure.
            NpLog "transport=pipe event=peer_closed err=" & CStr(pe)
            Teardown "peer_closed"
        Else
            NpLog "transport=pipe event=peekfail err=" & CStr(pe)
            Teardown "peek_failed"
        End If
        Exit Function
    End If
    If avail = 0 Then
        If (now - m_idleAt) > m_idleMs Then
            NpLog "transport=pipe result=teardown reason=idle_close idleMs=" & CStr(m_idleMs)
            Teardown "idle_close"
        End If
        Exit Function
    End If
    Dim toRead As Long
    toRead = avail
    If toRead > (NP_ACC_CAP - m_accLen) Then toRead = NP_ACC_CAP - m_accLen
    If toRead <= 0 Then
        ' Safety net only: the header check in ProcessFrames refuses over-long
        ' frames long before the accumulator can fill up. Never truncate.
        SendDenyFrame "frame_too_large", "accumulator_full"
        Exit Function
    End If
    Dim br As Long, e As Long
    ResetOv
    ok = NpReadFile(m_hPipe, m_acc(m_accLen), toRead, br, m_ov)
    e = Err.LastDllError
    If ok <> 0 Then
        m_accLen = m_accLen + br
        m_bytes = m_bytes + br
        m_idleAt = now
        m_frameAt = now
        ProcessFrames
        ServiceConnected = True
    ElseIf e = NP_ERROR_IO_PENDING Or e = NP_ERROR_IO_INCOMPLETE Or e = 0 Then
        ' 996 is what GetOverlappedResult(bWait=0) reports (spike trap 3) and a 0
        ' error code must NOT be read as fatal (spike trap 4: recv returned
        ' SOCKET_ERROR with WSAGetLastError()==0 on a perfectly healthy socket).
        m_state = 3
        m_pendingAt = now
    Else
        m_errs = m_errs + 1
        NpLog "transport=pipe event=readfail err=" & CStr(e)
        Teardown "read_error"
    End If
End Function

' --- state 3: overlapped read completion ----------------------------------

Private Function PollRead() As Boolean
    PollRead = False
    Dim ok As Long, br As Long, e As Long, now As Long
    ok = NpGetOverlappedResult(m_hPipe, m_ov, br, 0)
    e = Err.LastDllError
    If ok <> 0 Then
        m_state = 2
        m_accLen = m_accLen + br
        m_bytes = m_bytes + br
        now = NpGetTickCount()
        m_idleAt = now
        m_frameAt = now
        ProcessFrames
        PollRead = True
    ElseIf e = NP_ERROR_IO_PENDING Or e = NP_ERROR_IO_INCOMPLETE Then
        now = NpGetTickCount()
        If (now - m_pendingAt) > m_readDeadlineMs Then
            NpLog "transport=pipe result=teardown reason=read_deadline_exceeded state=read"
            Teardown "read_deadline_exceeded"
        End If
    Else
        m_errs = m_errs + 1
        NpLog "transport=pipe event=readfail err=" & CStr(e) & " state=complete"
        Teardown "read_complete_error"
    End If
End Function

' --- state 4: overlapped write completion ---------------------------------

Private Function PollWrite() As Boolean
    PollWrite = False
    Dim ok As Long, br As Long, e As Long
    ok = NpGetOverlappedResult(m_hPipe, m_ov, br, 0)
    e = Err.LastDllError
    If ok <> 0 Then
        m_outSent = m_outSent + br
        If m_outSent >= m_outLen Then
            m_outLen = 0
            m_outSent = 0
        End If
        m_state = 2
        PollWrite = True
    ElseIf e = NP_ERROR_IO_PENDING Or e = NP_ERROR_IO_INCOMPLETE Then
        Exit Function
    Else
        m_errs = m_errs + 1
        NpLog "transport=pipe event=writefail err=" & CStr(e)
        Teardown "write_error"
    End If
End Function

' --- framing --------------------------------------------------------------

Private Sub ProcessFrames()
    Dim guard As Long
    guard = 0
    Do While m_accLen >= 4
        guard = guard + 1
        If guard > 64 Then Exit Do
        ' 4-byte little-endian length. Read as Double first so a hostile or
        ' corrupted prefix cannot overflow a Long before it is range-checked.
        Dim flen As Double
        flen = CDbl(m_acc(0)) + CDbl(m_acc(1)) * 256# + CDbl(m_acc(2)) * 65536# + CDbl(m_acc(3)) * 16777216#
        If flen > CDbl(NP_MAX_FRAME) Then
            m_errs = m_errs + 1
            NpLog "transport=pipe event=badframe len=" & CStr(flen)
            SendDenyFrame "frame_too_large", CStr(flen)
            Exit Sub
        End If
        Dim n As Long
        n = CLng(flen)
        ' ADR section 7.3: read until the WHOLE payload is present. A short read
        ' is normal and must never be mistaken for the end of a frame.
        If m_accLen < 4 + n Then Exit Do
        Dim txt As String
        txt = TakeFrame(n)
        m_frames = m_frames + 1
        If m_hPipe = 0 Then Exit Sub
        HandleFrame txt
        If m_hPipe = 0 Then Exit Sub
    Loop
End Sub

Private Function TakeFrame(ByVal n As Long) As String
    Dim b() As Byte
    If n > 0 Then
        ReDim b(0 To n - 1)
        ' One CopyMemory instead of a VB loop: RtlMoveMemory is memmove, so it is
        ' safe for the overlapping shift below too. The accumulator is flat (see
        ' the header note about VB6 column-major 2-D arrays).
        NpCopyMemory b(0), m_acc(4), n
        TakeFrame = Utf8DecodeBytes(b)
    Else
        TakeFrame = ""
    End If
    Dim remaining As Long
    remaining = m_accLen - 4 - n
    If remaining > 0 Then NpCopyMemory m_acc(0), m_acc(4 + n), remaining
    m_accLen = remaining
    m_frameAt = NpGetTickCount()
End Function

Private Sub HandleFrame(ByVal txt As String)
    If Not m_authed Then
        HandleHello txt
        Exit Sub
    End If
    ' JSON-RPC payloads never carry a "vb6mcp" member, so the control frames of
    ' ADR section 7.3 are recognised by that member alone. The cheap InStr guard
    ' keeps the common path from being parsed twice.
    Dim kind As String
    kind = ""
    If InStr(1, txt, """vb6mcp""", vbBinaryCompare) > 0 Then kind = ControlKind(txt)
    If kind = "progress" Or kind = "welcome" Or kind = "pong" Or kind = "deny" Then
        ' Server-to-client kinds arriving here are ignored: never echoed, never
        ' treated as a request.
        Exit Sub
    ElseIf kind = "ping" Then
        QueueOut "{""vb6mcp"":""pong""}"
        Exit Sub
    ElseIf kind = "hello" Then
        ' A second hello on an already authenticated connection is a protocol
        ' error, not a re-handshake (the client is told to reconnect instead).
        SendDenyFrame "unexpected_hello", ""
        Exit Sub
    End If
    ServeRpc txt
End Sub

Private Sub HandleHello(ByVal txt As String)
    Dim kind As String
    kind = ControlKind(txt)
    If kind <> "hello" Then
        ' Fail-closed: the FIRST frame must be the handshake (ADR section 3.4).
        SendDenyFrame "bad_hello", kind
        Exit Sub
    End If
    If Not HelloNonceMatches(txt) Then
        NpLog "transport=pipe event=deny reason=nonce_mismatch"
        SendDenyFrame "nonce_mismatch", ""
        Exit Sub
    End If
    m_authed = True
    QueueOut WelcomeJson()
    NpLog "transport=pipe event=welcome instanceId=" & m_instanceId & " tcpEnabled=" & BoolText(m_tcpEnabled)
End Sub

Private Function HelloNonceMatches(ByVal txt As String) As Boolean
    HelloNonceMatches = False
    On Error Resume Next
    Dim v As Variant
    If Not JsonParse(txt, v) Then Exit Function
    If Not IsObject(v) Then Exit Function
    Dim o As Object
    Set o = v
    Dim got As String
    got = JsonGetStr(o, "nonce", "")
    ' Constant-ish comparison is not worth a crypto framing here: the nonce is a
    ' pairing credential, not a password (ADR section 5.2 - the ACL is the access
    ' control). An empty or absent nonce never matches a 16-hex nonce.
    If Len(got) = Len(m_nonce) Then
        If StrComp(got, m_nonce, vbBinaryCompare) = 0 Then HelloNonceMatches = True
    End If
End Function

Private Function WelcomeJson() As String
    Dim proj As String, sha As String
    proj = ""
    sha = ""
    On Error Resume Next
    If Not m_mcp Is Nothing Then
        proj = m_mcp.ProjectName()
        sha = m_mcp.ProductBuildSha256()
    End If
    ' ADR section 3.4 welcome + two ADDITIVE fields so the hub can see which leg
    ' of this instance is live (the visible-transport requirement, section 0.2
    ' rule 3 / section 2.5). instanceId echoes pid + nonce, which is the pairing
    ' check the client performs.
    WelcomeJson = "{""vb6mcp"":""welcome"",""proto"":""np/1"",""instanceId"":""" & JsonEsc(m_instanceId) & """,""pid"":" & CStr(m_pid) & ",""session"":" & CStr(m_sid) & ",""project"":""" & JsonEsc(proj) & """,""dllSha256"":""" & JsonEsc(sha) & """,""transport"":""pipe"",""tcpEnabled"":" & BoolText(m_tcpEnabled) & "}"
End Function

' --- request servicing ----------------------------------------------------

Private Sub ServeRpc(ByVal txt As String)
    Dim resp As String
    m_serving = True
    m_serveStart = NpGetTickCount()
    m_curTool = ExtractToolName(txt)
    m_progSeq = 0
    m_lastProgAt = m_serveStart
    On Error Resume Next
    resp = m_mcp.HandleJsonRpc(txt)
    If Err.Number <> 0 Then
        ' The TCP path answers a 500 here; a pipe must answer SOMETHING or the
        ' client waits for its whole forward budget.
        NpLog "transport=pipe event=rpcraise err=" & CStr(Err.Number) & " " & Err.Description
        resp = InternalErrorJson(txt, Err.Number)
        Err.Clear
    End If
    On Error GoTo 0
    m_serving = False
    m_idleAt = NpGetTickCount()
    ' An empty response is a JSON-RPC notification: send no frame (matching the
    ' /rpc semantics, where notifications produce no body).
    If Len(resp) > 0 Then QueueOut resp
End Sub

' ADR section 7.4 class 3: progress frames let the client tell "busy" from
' "dead" without shortening the 120 s forward budget. Emitted ONLY from a
' nested tick (i.e. inside DoEvents of a long call), never from a timer thread,
' and never while a frame is still queued.
' Transmit whatever is queued RIGHT NOW, from inside a nested tick. Used only by
' MaybeProgress, whose guards already prove that no write is in flight.
Private Sub FlushProgressNow()
    On Error Resume Next
    If m_outLen <> 0 Then Exit Sub
    If m_outQ.Count = 0 Then Exit Sub
    If Not BeginNextOut() Then Exit Sub
    ' The return value is deliberately ignored: False here means "went pending"
    ' (state 4) OR "hard error, torn down" - both are handled by the outer loop.
    NpStartWrite
End Sub

Private Sub MaybeProgress()
    ' diagnostic: every guard is reported as a machine-readable reason. The
    ' nested_tick line is rate limited (first 10, then every 100th) so a storm cannot
    ' flood the log while the decisive first occurrences stay visible.
    Dim why As String
    why = ""
    If Not m_progressOn Then
        why = "progress_off"
    ElseIf Not m_serving Then
        why = "not_serving"
    ElseIf Not m_authed Then
        why = "not_authed"
    ElseIf m_state <> 2 Then
        why = "state_" & CStr(m_state)
    ElseIf m_outLen > 0 Then
        why = "out_in_flight"
    ElseIf m_outQ.Count > 0 Then
        why = "queue_nonempty"
    Else
        Dim now As Long
        now = NpGetTickCount()
        If (now - m_lastProgAt) < m_progressMs Then
            why = "too_soon_" & CStr(now - m_lastProgAt)
        Else
            m_progSeq = m_progSeq + 1
            QueueOut "{""vb6mcp"":""progress"",""seq"":" & CStr(m_progSeq) & ",""elapsedMs"":" & CStr(now - m_serveStart) & ",""tool"":""" & JsonEsc(m_curTool) & """}"
            m_lastProgAt = now
            m_progFrames = m_progFrames + 1
            ' MEASURED FIX (t16): the outer service loop is BLOCKED inside
            ' HandleJsonRpc for the whole tool call, so a merely QUEUED frame is not
            ' transmitted until the call returns. Measured before this fix, under the
            ' Win32-timer host: progress seq=1 with elapsedMs=2000 reached the client
            ' at wallMs=10922 - i.e. only when the 10.9 s call had already finished,
            ' which defeats the whole point (ADR 7.4 class 3 needs it DURING the
            ' call). m_outLen = 0 is guaranteed by the guards above, so starting the
            ' write from here cannot disturb an in-flight transfer; if it goes
            ' pending, m_state becomes 4 and the outer loop completes it.
            FlushProgressNow
            why = "EMITTED_seq_" & CStr(m_progSeq)
            NpLog "transport=pipe event=progress seq=" & CStr(m_progSeq) & " elapsedMs=" & CStr(now - m_serveStart) & " nestedTicks=" & CStr(m_nestedTicks)
        End If
    End If
    m_nestedLogged = m_nestedLogged + 1
    If m_nestedLogged <= 10 Or (m_nestedLogged Mod 100) = 0 Then
        NpLog "transport=pipe event=nested_tick n=" & CStr(m_nestedTicks) & " serving=" & CStr(m_serving) & " authed=" & CStr(m_authed) & " state=" & CStr(m_state) & " outLen=" & CStr(m_outLen) & " q=" & CStr(m_outQ.Count) & " sinceProgMs=" & CStr(NpGetTickCount() - m_lastProgAt) & " => " & why
    End If
End Sub

Private Function ExtractToolName(ByVal txt As String) As String
    ExtractToolName = ""
    On Error Resume Next
    Dim v As Variant
    If Len(txt) = 0 Then Exit Function
    If InStr(1, txt, """tools/call""", vbBinaryCompare) = 0 Then Exit Function
    If Not JsonParse(txt, v) Then Exit Function
    If Not IsObject(v) Then Exit Function
    Dim o As Object
    Set o = v
    Dim pv As Variant
    Dim pIsObj As Boolean
    pIsObj = False
    If Not JsonGet(o, "params", pv, pIsObj) Then Exit Function
    If Not pIsObj Then Exit Function
    Dim po As Object
    Set po = pv
    ExtractToolName = JsonGetStr(po, "name", "")
End Function

Private Function InternalErrorJson(ByVal requestText As String, ByVal errNo As Long) As String
    InternalErrorJson = "{""jsonrpc"":""2.0"",""id"":" & ExtractIdJson(requestText) & ",""error"":{""code"":-32603,""message"":""internal error in the add-in (VB error " & CStr(errNo) & ") over the pipe transport""}}"
End Function

Private Function ExtractIdJson(ByVal requestText As String) As String
    ExtractIdJson = "null"
    On Error Resume Next
    Dim v As Variant
    If Not JsonParse(requestText, v) Then Exit Function
    If Not IsObject(v) Then Exit Function
    Dim o As Object
    Set o = v
    If Not o.Exists("id") Then Exit Function
    If IsObject(o.Item("id")) Then Exit Function
    Dim iv As Variant
    iv = o.Item("id")
    If IsNull(iv) Then Exit Function
    ExtractIdJson = JsonStringify(iv)
    If Len(ExtractIdJson) = 0 Then ExtractIdJson = "null"
End Function

Private Function ControlKind(ByVal txt As String) As String
    ControlKind = ""
    On Error Resume Next
    Dim v As Variant
    If Not JsonParse(txt, v) Then Exit Function
    If Not IsObject(v) Then Exit Function
    Dim o As Object
    Set o = v
    ControlKind = LCase$(JsonGetStr(o, "vb6mcp", ""))
End Function

' --- output queue ---------------------------------------------------------
' Frames are serialised one at a time: m_out is only ever ReDim'd while
' m_outLen = 0, so a pending overlapped write can never have its buffer moved
' underneath it. A frame that is queued while another is in flight waits in
' m_outQ (this is also the backpressure that stops us reading a second request
' while a big response is still draining).

Private Sub QueueOut(ByVal payload As String)
    On Error Resume Next
    If Len(payload) = 0 Then Exit Sub
    m_outQ.Add payload
End Sub

Private Function BeginNextOut() As Boolean
    BeginNextOut = False
    If m_outQ.Count = 0 Then Exit Function
    Dim payload As String
    payload = m_outQ.Item(1)
    m_outQ.Remove 1
    Dim b() As Byte
    Dim n As Long
    b = Utf8Encode(payload)
    n = 0
    On Error Resume Next
    n = UBound(b) - LBound(b) + 1
    If Err.Number <> 0 Then n = 0
    Err.Clear
    On Error GoTo 0
    If n > NP_MAX_FRAME Then
        m_errs = m_errs + 1
        NpLog "transport=pipe event=response_too_large bytes=" & CStr(n) & " (frame dropped, connection kept)"
        Exit Function
    End If
    ReDim m_out(0 To 3 + n)
    m_out(0) = n And &HFF
    m_out(1) = (n \ 256) And &HFF
    m_out(2) = (n \ 65536) And &HFF
    m_out(3) = (n \ 16777216) And &HFF
    If n > 0 Then NpCopyMemory m_out(4), b(0), n
    m_outLen = 4 + n
    m_outSent = 0
    BeginNextOut = True
End Function

Private Function NpStartWrite() As Boolean
    NpStartWrite = False
    Dim n As Long
    n = m_outLen - m_outSent
    If n <= 0 Then
        m_outLen = 0
        m_outSent = 0
        NpStartWrite = True
        Exit Function
    End If
    ResetOv
    Dim ok As Long, bw As Long, e As Long
    ok = NpWriteFile(m_hPipe, m_out(m_outSent), n, bw, m_ov)
    e = Err.LastDllError
    If ok <> 0 Then
        m_outSent = m_outSent + bw
        If m_outSent >= m_outLen Then
            m_outLen = 0
            m_outSent = 0
        End If
        NpStartWrite = True
    ElseIf e = NP_ERROR_IO_PENDING Or e = NP_ERROR_IO_INCOMPLETE Then
        m_state = 4
    Else
        m_errs = m_errs + 1
        NpLog "transport=pipe event=writefail err=" & CStr(e)
        Teardown "write_error"
    End If
End Function

' --- teardown -------------------------------------------------------------

Private Sub SendDenyFrame(ByVal reason As String, ByVal detail As String)
    QueueOut "{""vb6mcp"":""deny"",""reason"":""" & JsonEsc(reason) & """,""detail"":""" & JsonEsc(detail) & """}"
    ' The connection is closed only AFTER the frame has actually drained, so the
    ' client always sees the reason rather than a broken pipe.
    m_closeAfterFlush = True
End Sub

Private Sub Teardown(ByVal reason As String)
    On Error Resume Next
    If m_hPipe <> 0 Then
        ' CancelIo (this thread's pending overlapped I/O) + CloseHandle.
        ' Two calls are deliberately ABSENT here:
        '   * FlushFileBuffers blocks until the CLIENT has read the buffered data -
        '     exactly the unbounded blocking call the ADR forbids in this loop;
        '   * DisconnectNamedPipe DISCARDED the buffered deny frame (measured on
        '     t16: the client got EOF and never saw reason=nonce_mismatch).
        ' Closing the handle keeps already-written bytes readable by the client.
        NpCancelIo m_hPipe
        NpCloseHandle m_hPipe
    End If
    On Error GoTo 0
    m_hPipe = 0
    m_state = 0
    m_accLen = 0
    m_outLen = 0
    m_outSent = 0
    m_authed = False
    m_closeAfterFlush = False
    m_closeAfterFlushAt = 0
    m_serving = False
    m_tears = m_tears + 1
    m_lastReason = reason
    Set m_outQ = New Collection
    If m_ov.hEvent <> 0 Then
        NpCloseHandle m_ov.hEvent
        m_ov.hEvent = 0
    End If
    NpLog "transport=pipe event=teardown reason=" & reason & " tears=" & CStr(m_tears) & " conns=" & CStr(m_conns) & " frames=" & CStr(m_frames) & " errors=" & CStr(m_errs) & " nestedTicks=" & CStr(m_nestedTicks) & " progressFrames=" & CStr(m_progFrames)
    ' ADR section 7.2: an idle close re-creates the instance on the next tick.
    m_lastCreateAt = 0
End Sub

' Has the client closed its end? Used only to release a deny connection early.
' The error codes matter: PeekNamedPipe reports ERROR_BROKEN_PIPE(109) or
' ERROR_PIPE_NOT_CONNECTED(233) once the peer is gone.
Private Function NpPeerGone() As Boolean
    NpPeerGone = False
    If m_hPipe = 0 Then
        NpPeerGone = True
        Exit Function
    End If
    Dim pr As Long, avail As Long, lft As Long, e As Long
    If NpPeekNamedPipe(m_hPipe, 0, 0, pr, avail, lft) = 0 Then
        e = Err.LastDllError
        If e = NP_ERROR_BROKEN_PIPE Or e = NP_ERROR_PIPE_NOT_CONNECTED Then NpPeerGone = True
    End If
End Function

Private Sub ResetOv()
    m_ov.Internal = 0
    m_ov.InternalHigh = 0
    m_ov.Offset = 0
    m_ov.OffsetHigh = 0
    If m_ov.hEvent <> 0 Then NpResetEvent m_ov.hEvent
End Sub

' --- ADR section 5.1: SDDL ACL (best effort, loud fallback) ---------------

Private Function BuildSecurityDescriptor(ByRef why As String) As Long
    ' Returns a pointer to a self-relative SD, or 0. Every step is checked AND
    ' named in 'why'; any failure returns 0 and the caller falls back to the
    ' default DACL and reports the failing step.
    On Error GoTo Fail
    BuildSecurityDescriptor = 0
    Dim hTok As Long
    hTok = 0
    If NpOpenProcessToken(NpGetCurrentProcess(), NP_TOKEN_QUERY, hTok) = 0 Then
        why = "open_process_token err=" & CStr(Err.LastDllError)
        Exit Function
    End If
    If hTok = 0 Then
        why = "open_process_token null"
        Exit Function
    End If
    Dim buf() As Byte
    ReDim buf(0 To 255)
    Dim need As Long
    need = 0
    If NpGetTokenInformation(hTok, NP_TOKEN_USER_CLASS, buf(0), 256, need) = 0 Then
        why = "get_token_information err=" & CStr(Err.LastDllError)
        NpCloseHandle hTok
        Exit Function
    End If
    NpCloseHandle hTok
    hTok = 0
    ' TOKEN_USER is { PSID Sid; DWORD Attributes; } - the SID pointer is the
    ' first 4 bytes of the buffer.
    Dim sidPtr As Long
    sidPtr = 0
    NpCopyMemoryToLong sidPtr, VarPtr(buf(0)), 4
    If sidPtr = 0 Then
        why = "sid_ptr_null"
        Exit Function
    End If
    Dim pStr As Long
    pStr = 0
    If NpConvertSidToStringSid(sidPtr, pStr) = 0 Then
        why = "sid_to_string err=" & CStr(Err.LastDllError)
        Exit Function
    End If
    If pStr = 0 Then
        why = "sid_to_string null"
        Exit Function
    End If
    Dim sidText As String
    sidText = PtrToStr(pStr)
    NpLocalFree pStr
    If Len(sidText) = 0 Then
        why = "sid_text_empty"
        Exit Function
    End If
    ' ADR section 5.1: SYSTEM full, the CURRENT USER full, nothing for Everyone
    ' or Authenticated Users.
    Dim sddl As String
    sddl = "D:(A;;GA;;;SY)(A;;GA;;;" & sidText & ")"
    Dim pSD As Long
    pSD = 0
    If NpConvertStringSecurityDescriptorToSecurityDescriptor(sddl, NP_SDDL_REVISION_1, pSD, 0) = 0 Then
        why = "sddl_convert err=" & CStr(Err.LastDllError) & " sddl=" & sddl
        Exit Function
    End If
    BuildSecurityDescriptor = pSD
    Exit Function
Fail:
    why = "raised err=" & CStr(Err.Number) & " " & Err.Description
    BuildSecurityDescriptor = 0
End Function

' LPSTR -> VB6 String, the CORRECT way. A VB6 string is UTF-16, so copying n ANSI
' bytes straight into StrPtr(s) produces one code unit per PAIR of bytes - which
' is exactly how the first version turned the user SID into mojibake and made
' ConvertStringSecurityDescriptor fail with ERROR_INVALID_SID(1337) (spike trap 2,
' measured again here on t16). Collect the bytes, then StrConv to Unicode.
Private Function PtrToStr(ByVal p As Long) As String
    PtrToStr = ""
    On Error Resume Next
    If p = 0 Then Exit Function
    Dim n As Long
    n = NpLstrLenA(p)
    If n <= 0 Then Exit Function
    Dim b() As Byte
    ReDim b(0 To n - 1)
    NpCopyMemoryFromLong VarPtr(b(0)), p, n
    PtrToStr = StrConv(b, vbUnicode)
End Function

' --- small helpers --------------------------------------------------------

Private Function SessionOfProcess(ByVal pid As Long) As Long
    Dim sid As Long
    sid = 0
    If NpProcessIdToSessionId(pid, sid) <> 0 Then
        SessionOfProcess = sid
    Else
        SessionOfProcess = 0
    End If
End Function

Private Function MakeNonce16() As String
    ' ADR section 5.2: RtlGenRandom (SystemFunction036), NOT rand(), and never
    ' persisted anywhere except the registry entry.
    On Error Resume Next
    Dim b(0 To 7) As Byte
    If NpRtlGenRandom(VarPtr(b(0)), 8) <> 0 Then
        Dim i As Long, s As String
        s = ""
        For i = 0 To 7
            s = s & Right$("0" & Hex$(b(i)), 2)
        Next i
        If Len(s) = 16 Then
            MakeNonce16 = LCase$(s)
            Exit Function
        End If
    End If
    ' Fallback (not expected): GUID hex digits, still 16 hex chars.
    Dim g As String
    g = ""
    g = CreateObject("Scriptlet.TypeLib").Guid
    g = LCase$(Replace(Replace(Replace(g, "-", ""), "{", ""), "}", ""))
    If Len(g) < 16 Then g = g & String$(16 - Len(g), "0")
    MakeNonce16 = Left$(g, 16)
End Function

Private Function Slugify(ByVal s As String) As String
    Dim i As Long, c As String, out As String
    out = ""
    s = Trim$(s)
    For i = 1 To Len(s)
        c = Mid$(s, i, 1)
        If (c >= "0" And c <= "9") Or (c >= "a" And c <= "z") Or (c >= "A" And c <= "Z") Or c = "_" Then
            out = out & c
        Else
            out = out & "_"
        End If
        If Len(out) >= 16 Then Exit For
    Next i
    Slugify = out
End Function

Private Function EnvFlag(ByVal name As String, ByVal dflt As Boolean) As Boolean
    Dim s As String
    s = LCase$(Trim$(Environ$(name)))
    If Len(s) = 0 Then
        EnvFlag = dflt
        Exit Function
    End If
    If s = "0" Or s = "off" Or s = "no" Or s = "false" Then
        EnvFlag = False
    Else
        EnvFlag = True
    End If
End Function

Private Function BoolText(ByVal v As Boolean) As String
    If v Then BoolText = "true" Else BoolText = "false"
End Function

Private Function JsonEsc(ByVal s As String) As String
    JsonEsc = Replace(Replace(Replace(s, "\", "\\"), """", "\"""), vbCrLf, " ")
End Function

' All pipe logging goes through Util.LogLine, the same %TEMP%\vb6mcp.log the
' TCP transports write to, so one file shows the selected track and every
' teardown reason (the "fallback/busy must be visible" requirement).
Private Sub NpLog(ByVal msg As String)
    On Error Resume Next
    Util.LogLine msg
End Sub

Private Sub EnsureFolder(ByVal fso As Object, ByVal path As String)
    On Error Resume Next
    If Len(path) = 0 Then Exit Sub
    If fso.FolderExists(path) Then Exit Sub
    Dim parent As String, i As Long
    i = InStrRev(path, "\")
    If i > 1 Then
        parent = Left$(path, i - 1)
        If Not fso.FolderExists(parent) Then EnsureFolder fso, parent
    End If
    fso.CreateFolder path
End Sub
