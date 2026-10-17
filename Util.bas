Attribute VB_Name = "Util"
Option Explicit

' Small shared helpers: logging, port file, temp dir.

' ---------------------------------------------------------------------------
' Contract v1.3.0 product metadata - the SINGLE source of truth for the version
' stamp. McpServer reports it in initialize (serverInfo.version +
' contractVersion); CodeTools reports it in mcp_info and in the per-response
' mcp stamp. ASCII only: VB6 compiles these sources as ANSI(GBK).
' ---------------------------------------------------------------------------
Public Const VB6MCP_PRODUCT As String = "vb6-mcp"

' this module needs the pid for the instance registry; declare it with
' a module-unique name (CodeTools declares GetCurrentProcessId privately too).
Private Declare Function UtilProcessId Lib "kernel32" Alias "GetCurrentProcessId" () As Long
' the window thread of this process (the UI thread that owns the form and will own
' the ingress window).  The bridge needs it in the self-registration record.
Public Declare Function UtilThreadId Lib "kernel32" Alias "GetCurrentThreadId" () As Long
' registration accounting state (module-level declarations MUST stay in the
' declaration region - VB6 refuses a Dim/Private after any procedure).
Private m_regState As String              ' "" | "pending" | "registered" | "failed"
Private m_regAttempts As Long
Private m_regLastError As String
Private m_regLastSuccessAt As String
Private m_regLastAttemptAt As String
Private m_regElapsedMs As Long
Private m_regHwnd As Long
Private m_regTid As Long
Private m_regEpoch As Long
Private m_regPipe As String
Private m_regNonce As String
Private m_regDllSha As String
Private m_regSession As Long

Private Declare Function PumpTickCount Lib "kernel32" Alias "GetTickCount" () As Long

' process-liveness probes for the instance-registry sweep (see
' PruneDeadInstanceFiles at the bottom of this module).
'
' WHY Toolhelp32 IS THE PRIMARY: OpenProcess + GetLastError was MEASURED broken in
' this host - a dead pid gives OpenProcess = NULL but the thread last-error reads 0
' (not ERROR_INVALID_PARAMETER 87), because VB6''s Declare epilogue clobbers it. That
' made every dead entry look alive and the sweep pruned nothing. A Toolhelp32 process
' snapshot needs NO error code and also sees processes owned by OTHER users, which
' OpenProcess can refuse with ERROR_ACCESS_DENIED. The OpenProcess path is kept as a
' fallback for the case where the snapshot is unusable, and it can only ever make the
' verdict MORE conservative.
Private Type UTIL_PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szExeFile As String * 260
End Type
Private Const UTIL_TH32CS_SNAPPROCESS As Long = &H2
Private Const UTIL_INVALID_HANDLE As Long = -1
Private Const UTIL_STILL_ACTIVE As Long = 259
Private Const UTIL_ERROR_INVALID_PARAMETER As Long = 87
Private Const UTIL_PROCESS_QUERY_LIMITED As Long = &H1000
Private Declare Function UtilCreateToolhelp32Snapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal dwFlags As Long, ByVal th32ProcessID As Long) As Long
Private Declare Function UtilProcess32First Lib "kernel32" Alias "Process32First" (ByVal hSnapshot As Long, ByRef lppe As UTIL_PROCESSENTRY32) As Long
Private Declare Function UtilProcess32Next Lib "kernel32" Alias "Process32Next" (ByVal hSnapshot As Long, ByRef lppe As UTIL_PROCESSENTRY32) As Long
Private Declare Function UtilOpenProcess Lib "kernel32" Alias "OpenProcess" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function UtilCloseHandle Lib "kernel32" Alias "CloseHandle" (ByVal hObject As Long) As Long
Private Declare Function UtilGetExitCodeProcess Lib "kernel32" Alias "GetExitCodeProcess" (ByVal hProcess As Long, ByRef lpExitCode As Long) As Long
Public Const VB6MCP_PRODUCT_VERSION As String = "0.3.0"

' THE contract banner - a single source of truth. The first line of
' initialize.instructions and mcp_info.contract BOTH read this one constant, so bumping
' the contract is a ONE-LINE change here. NEVER hardcode a contract version anywhere
' else in the product: a stale self-report is the exact defect this card fixed (the
' banner still claimed v1.3.0 after the contract text had reached v1.4.0, and because
' the hub forwards the LIVE banner, users saw the old version).
' the snapshot parameter is an inputSchema change, so the contract
' version moves with it. It is ADDITIVE (optional parameter, default off), but a tool-face
' change still has to be named by a version - the rule above is "bump it HERE, in one line".
' (Lead batch ruling 2026-10-15): mcp_info.build{} gains two ADDITIVE keys,
' sourceHash/sourceAlgo (the SOURCE STAMP of docs/SOURCE_STAMP.md). No tool name,
' description or inputSchema changes, and no existing receipt key changes - the bump
' exists because a NEW receipt field is a contract-visible addition, exactly as v1.11.1
' was for its ADDED receipt fields.
' (v1.11.3): dependency_graph gained ADDITIVE coverage fields (partial /
' charQuota / charsProcessed / componentsScanned / componentsTotal / procEdgesFound /
' edgeStageCompleted / purityStageCompleted / cyclesStageCompleted / timingTruncated /
' completeHint) and its cutoff became DETERMINISTIC (a character quota) so the same
' input returns the same receipt. No existing key changed.
' (v1.11.4, MCP-台账批): the write-class and build tools gain behaviour and
' receipts: a PRE-WRITE structural pairing gate on every block-writing tool plus a
' result-level guard (extra_procedure_end / missing_procedure_end /
' predicted_procedure_imbalance, with block/pane line numbers), candidate hits[] for an
' ambiguous anchor on the batch path, MULTI-LINE block anchors in preview_edit (already
' in edit_by_anchor/apply_edits), procBefore/procAfter on replace_code and apply_edits,
' make_project buildState + wait_for_build, and the linter!=compiler clarification in
' validate_code/check_component. No tool added or removed; inputSchema gains ONE optional
' parameter (make_project.wait_for_build), so the frozen face must be re-frozen.
' /170/171 (v1.11.5): a REAL semantics change, so the contract version is bumped at
' patch level and the change is broadcast before the face is re-frozen:
'   * (plan B) RETRACTS the promise that a read-only call is served while another call is
'     in flight (readonly_serviceable is now always false; cancel/release are BEST EFFORT);
' * newTotalLines is the CALLER's splitlines line count on all four write paths, and
'     a new trailingBlankLine field explains the difference from the IDE's own count;
' * the hub's ide_stop/ide_restart share one ownership resolver, so the same pid can no
'     longer get two verdicts, and every refusal carries alsoForeignPort/ownershipBasis.
' v1.11.6 (P0 honesty batch): a REAL semantics change, broadcast before the face is re-frozen:
'   * check_public_api ALWAYS compares when a component name is given (membersOnly is output
'     size only) and refuses a baseline it cannot parse with reason=baseline_unusable instead
'     of reporting no differences; export_public_api/check_public_api/validate_code/
'     check_component refuse undeclared argument names with reason=unknown_argument +
'     acceptedArguments.
'   * check_component gains the OPTIONAL baseline parameter (declared and honoured; an unusable
'     baseline is refused with baseline_unusable). This is the one inputSchema addition.
'   * write_file/write_source publish createdDirs[]/createdDirsCount/exists/size, and a parent
'     that cannot be created is refused as parent_dir_missing (missingPath, reasonLegacy=
'     parent_not_creatable) instead of a generic failure.
'   * find_in_component publishes hits[]/count/nextStartLine and answers found:false +
'     reason=no_more_hits at the end of a pagination walk (it never wraps back to line 1).
'   * make_project{diagnose} isolated copies now carry the binary-compatibility reference
'     (CompatibleEXE32/CompatibleEXE) and a PropertyPage .pgx, and report
'     compatibilityReferenceCopied; a missing reference is named in missingFiles.
' v1.11.7 (write-receipt + diagnosis batch): behaviour-lit changes, broadcast before re-freeze:
'   * unknown-argument gate for the WHOLE tool face (declared property set is the contract; _target excepted);
'     reasons split into unknown_argument / component_not_found / target_has_no_project / ide_not_ready
'     (legacy no_project / ide_model_not_ready are kept in reasonLegacy);
'   * add_code publishes insertedAt{ide,disk} + inDeclarationArea; every write-class receipt keeps
'     ok + sha256Before/sha256After + lineDelta + the measured landing;
'   * verify annotates an ATTRIBUTE-ONLY pane/disk difference (attributeOnlyDiff/attributeLineCount);
'   * get_component_code{proc|symbol} states the range line space (IDE) and adds startLineDisk/endLineDisk;
'   * schema gains the declared-only parameters that were honoured but undeclared (preview_edit.match,
'     set_codepane.view, vba6_probe.step/frame_index).
' v1.11.8 (internal-identifier cleanup, rebuild window): tool descriptions and receipts no longer
'   carry internal task identifiers; the DELIVERED behaviour is unchanged (wording only, plus the
'   generated artifacts re-derived from the cleaned descriptions).
' v1.11.9 (naming/shape + count semantics batch), broadcast before the face is re-frozen:
'   * add_code.inDeclarationArea now means the DECLARATION AREA (insertedAt.ide < the first
'     procedure header line) instead of "outside an unclosed procedure body", so an END
'     append on a module that already has procedures reports false;
'   * requestedRange has ONE shape per mode: mode="range" {operation,start_line,end_line} on
'     replace_code / add_code / guard refusals, mode="anchor" / "anchor-batch" on
'     edit_by_anchor and apply_edits;
'   * verify_scope: UNKNOWN is NOT a fail (it is counted once in `unknown` and listed in
'     unknowns[] only), so pass + fail + unknown == checked always holds (countsConsistent).
' v1.11.11 (changelog catch-up + one description-hygiene pass; tool count unchanged 110):
'   * BEHAVIOUR FIX, restated honestly: the v1.11.10 batch fixed three winsock transport defects
'     (the per-connection object was never persisted, so a request split across TCP segments lost
'     its head; close-after-send was tested after the queue had already been cleared; the flush
'     path grew a cumulative Mid$ start past 64 KiB and treated send()=-1 with
'     WSAGetLastError()=0 as fatal). Together they closed the intermittent RemoteDisconnected
'     and the >64 KiB truncation - a behaviour fix, NOT a wording change.
'   * TEXT-ONLY edits in the same batch: 3 tool descriptions (edit_by_anchor / apply_edits /
'     preview_edit: internal-id prefixes removed), 1 tool description (verify_scope: build evidence
'     is PROJECT-level), 2 parameter descriptions (replace_code.expected_sha256 /
'     make_project.wait_for_build), 1 runtime receipt string (make_project.buildStateNote lost its
'     internal-id prefix; fields and semantics unchanged).
'   * v1.11.11: 10 vba6_probe PARAMETER descriptions no longer carry internal experiment ids or
'     reverse-engineering round numbers (same meaning, no parameter added/removed/retyped), and the
'     remaining internal experiment ids and round numbers are gone from the shipped comments.
Public Const VB6MCP_CONTRACT_VERSION As String = "v1.11.11"
' message-pump period and state. The DECLARATIONS live here (VB6 only accepts
' module-level declarations before the first procedure); the pump itself is documented at
' the bottom of this module next to PumpReset / MaybePump / PumpNow.
Public Const PUMP_EVERY_ITER As Long = 1000
Public Const PUMP_EVERY_MS As Long = 150
Private Const PUMP_TICK_MASK As Long = 63

Private m_pumpIter As Long
Private m_pumpTimeIter As Long
Private m_pumpLastAt As Long
Private m_pumpDepth As Long
Private m_pumpCount As Long
Private m_pumpMs As Long
Private m_pumpMaxMs As Long
' the ONE gate for the whole pump class. A host modal box that is UP means a
' DoEvents here can dispatch a POSTED host command and enter the host's modal loop -
' measured in: the add-in then never returns, no watchdog tick is delivered
' inside that loop, and the single pipe instance stays held for ever. So: no DoEvents
' while a visible #32770 of this process is up. The cached probe keeps it cheap.
Private Const PUMP_GATE_CACHE_MS As Long = 100
Private m_gateAt As Long
Private m_gateBlocked As Boolean
Private m_gateSkips As Long
Private m_gateTitle As String

' item 6: the side-effects AUDIT BLOCK is accumulated per top-level
' call in CodeTools.cls, but VbFmt.bas (format_code / rename_symbol) writes .bak files as
' well and is a different module. This ONE collection is therefore shared: CodeTools
' resets it once per call and reads it when it builds the additive `sideEffects` block;
' VbFmt contributes to it through SeNoteBackup below. Additive only - nothing else in
' this module changed. (The other five collections live in CodeTools.cls because only
' CodeTools creates components / renames them / removes them / writes undo snapshots.)
Public g_seBackups As Collection

Public Sub LogLine(ByVal msg As String)
    On Error Resume Next
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim path As String
    path = TempDir() & "\vb6mcp.log"
    Dim ts As Object
    Set ts = fso.OpenTextFile(path, 8, True)
    ts.WriteLine Format$(Now, "yyyy-mm-dd hh:nn:ss") & "  " & msg
    ts.Close
End Sub

' ---------------------------------------------------------------------------
' Capture the value of AddressOf.
' VB6's grammar rejects "p = AddressOf Foo" (syntax error), and AddressOf may not
' be written into a UDT member either. It IS accepted as an argument to an
' ordinary VB function, so this one-liner is the supported way to obtain the
' address as a Long:
'     Dim p As Long
'     p = FuncAddr(AddressOf MyProc)
' Verified here by compiling AND running a minimal Standard EXE: the captured
' value was 0x4016B0 (a real code address), with the target still callable.
' ---------------------------------------------------------------------------
Public Function FuncAddr(ByVal ptr As Long) As Long
    FuncAddr = ptr
End Function

Public Sub WritePortFile(ByVal port As Long)
    On Error Resume Next
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim path As String
    path = TempDir() & "\vb6mcp_port.txt"
    Dim ts As Object
    Set ts = fso.CreateTextFile(path, True)
    ts.WriteLine CStr(port)
    ts.Close
End Sub

' ---------------------------------------------------------------------------
' (contract v1.3.1 H2): instance registration
' %TEMP%\vb6mcp\instances\<pid>.json  {pid,port,project,projectPath,dllSha256,
' startedAt,startedByHub}. Written on listen success, deleted on normal exit.
' Stale entries are the READER@s business (pid alive + port listening); a writer
' never deletes another instance@s file.
' ---------------------------------------------------------------------------

Public Function Vb6mcpDir() As String
    Vb6mcpDir = TempDir() & "\vb6mcp"
End Function

' ===================== registration accounting =====================
' The registration RETRY loop lives in ServerForm (it owns the publisher); the STATE lives
' here so a read-only receipt (mcp_info.registration) can report it without reaching into
' the form.  Before this card every publish error was swallowed by On Error Resume Next, so
' a registration that never landed was indistinguishable from a successful one (the same
' silent-failure family as). Nothing here blocks: the loop is driven from the
' existing watchdog tick and never waits for a bridge.

Public Function RegState() As String
    If Len(m_regState) = 0 Then RegState = "pending" Else RegState = m_regState
End Function

Public Function RegAttempts() As Long
    RegAttempts = m_regAttempts
End Function

Public Function RegLastError() As String
    RegLastError = m_regLastError
End Function

Public Sub RegNoteAttempt(ByVal attempts As Long, ByVal errText As String)
    m_regAttempts = attempts
    m_regLastError = errText
    m_regLastAttemptAt = Format$(Now, "yyyy-mm-dd hh:nn:ss")
    If m_regState <> "registered" Then m_regState = "pending"
End Sub

Public Sub RegNoteRecord(ByVal hwnd As Long, ByVal tid As Long, ByVal epoch As Long, _
                         ByVal pipeName As String, ByVal nonce As String, ByVal dllSha As String, ByVal sessionId As Long)
    m_regHwnd = hwnd
    m_regTid = tid
    m_regEpoch = epoch
    m_regPipe = pipeName
    m_regNonce = nonce
    m_regDllSha = dllSha
    m_regSession = sessionId
End Sub

Public Sub RegNoteSuccess(ByVal elapsedMs As Long)
    m_regState = "registered"
    m_regLastError = ""
    m_regLastSuccessAt = Format$(Now, "yyyy-mm-dd hh:nn:ss")
    m_regElapsedMs = elapsedMs
End Sub

Public Sub RegNoteFailed(ByVal elapsedMs As Long)
    m_regState = "failed"
    m_regElapsedMs = elapsedMs
End Sub

' the full self-registration record, so one retry does BOTH jobs (retry and
' HWND reporting) instead of building the report twice.
Public Function RegistrationInfo() As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("state") = RegState()
    d("attempts") = m_regAttempts
    d("lastError") = m_regLastError
    d("lastSuccessAt") = m_regLastSuccessAt
    d("lastAttemptAt") = m_regLastAttemptAt
    d("elapsedMs") = m_regElapsedMs
    d("pid") = UtilProcessId()
    d("tid") = m_regTid
    d("hwnd") = m_regHwnd
    d("epoch") = m_regEpoch
    d("sessionId") = m_regSession
    d("dllSha256") = m_regDllSha
    d("pipeName") = m_regPipe
    d("sessionNonce") = m_regNonce
    d("backoffMs") = "500/1000/2000/4000, cap 15000"
    Set RegistrationInfo = d
End Function

Public Function InstancesDir() As String
    InstancesDir = Vb6mcpDir() & "\instances"
End Function

' the instance registration MUST stay pure ASCII. This writer uses FSO
' CreateTextFile (ANSI / system ACP), so a project name such as 工程1 was written as
' GBK bytes; a strict UTF-8 reader (the Python bridge) then failed to decode the whole
' entry and dropped it. Escaping every non-ASCII UTF-16 code unit as \uXXXX makes the
' file ASCII-only, so the ANSI writer, a strict UTF-8 reader and JSON.parse all decode
' the SAME characters.
' It is applied to the SERIALIZED text, never to JsonStringify itself: JsonStringify's
' structural characters are ASCII and any quote/backslash/control character inside a
' string value has already been escaped by it, so a >127 character in its output is
' always inside a string literal and a blanket per-code-unit rewrite is correct
' (astral characters become the standard \uD83D\uDE00 surrogate-pair form).
' Pure-ASCII input is returned byte-for-byte unchanged (zero behaviour change).
Public Function AsciiEscapeJson(ByVal s As String) As String
    Dim n As Long
    n = Len(s)
    If n = 0 Then Exit Function
    Dim i As Long
    Dim cc As Long
    Dim need As Boolean
    need = False
    For i = 1 To n
        ' AscW returns a SIGNED 16-bit value, so mask to get the true code unit.
        cc = AscW(Mid$(s, i, 1)) And &HFFFF&
        If cc > 127 Then
            need = True
            Exit For
        End If
    Next i
    If Not need Then
        AsciiEscapeJson = s
        Exit Function
    End If
    Dim out As String
    Dim j As Long
    out = Space$(n * 6)
    j = 0
    For i = 1 To n
        cc = AscW(Mid$(s, i, 1)) And &HFFFF&
        If cc > 127 Then
            Mid$(out, j + 1, 6) = "\u" & Right$("0000" & Hex$(cc), 4)
            j = j + 6
        Else
            Mid$(out, j + 1, 1) = Chr$(cc)
            j = j + 1
        End If
    Next i
    AsciiEscapeJson = Left$(out, j)
End Function

Public Sub WriteInstanceFile(ByVal pid As Long, ByVal port As Long, ByVal project As String, _
                             ByVal projectPath As String, ByVal dllSha As String, _
                             Optional ByVal startedAt As String = "")
    On Error Resume Next
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso Is Nothing Then Exit Sub
    EnsureFolder fso, Vb6mcpDir()
    EnsureFolder fso, InstancesDir()
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("pid") = pid
    d("port") = port
    d("project") = project
    d("projectPath") = projectPath
    d("dllSha256") = dllSha
    ' the entry may be RE-published when the identity changes; startedAt must
    ' keep meaning "when this instance started", so the caller passes the first value back in.
    If Len(startedAt) = 0 Then
        d("startedAt") = Format$(Now, "yyyy-mm-dd hh:nn:ss")
    Else
        d("startedAt") = startedAt
    End If
    d("startedByHub") = (Environ$("VB6MCP_STARTED_BY_HUB") = "1")
    Dim ts As Object
    Set ts = fso.CreateTextFile(InstancesDir() & "\" & CStr(pid) & ".json", True)
    ' ASCII-only output so every reader (ANSI / strict UTF-8 / JSON.parse) agrees.
    ts.WriteLine AsciiEscapeJson(JsonStringify(d))
    ts.Close
End Sub

Public Sub DeleteInstanceFile(ByVal pid As Long)
    On Error Resume Next
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso Is Nothing Then Exit Sub
    Dim path As String
    path = InstancesDir() & "\" & CStr(pid) & ".json"
    If fso.FileExists(path) Then fso.DeleteFile path, True
End Sub

' ---------------------------------------------------------------------------
' prune DEAD instance registrations.
'
' StopListening already deletes our own <pid>.json on a CLEAN exit, but when the
' IDE crashes or is killed (Stop-Process -Force / taskkill) that code never runs
' and the entry stays in %TEMP%\vb6mcp\instances forever. MEASURED before this
' change: 50 entries, registrySummary {entries:50, live:0, stale:50,
' staleByReason:{pid_dead:50}}. The cleanup must therefore not depend on the
' dying process: the next instance that REGISTERS sweeps the directory.
'
' CONSERVATIVE BY DESIGN - an entry is deleted ONLY when its pid is PROVABLY
' gone. Everything else is KEPT:
'   * our own pid is always skipped;
'   * only file names that are digits + .json are considered at all;
'   * an entry whose pid IS alive is never touched;
'   * an entry that cannot be parsed, has no digits-only "pid", or whose pid
'     disagrees with its own <pid>.json file name is KEPT (counted and reported);
'   * a probe that cannot prove the pid is gone (no usable process snapshot, a
'     failed OpenProcess, an empty enumeration) counts as ALIVE;
'   * a delete that fails (file held open) is swallowed and does not fail the
'     registration path - the sweep is best effort.
' ---------------------------------------------------------------------------
Public Sub PruneDeadInstanceFiles(ByVal ownPid As Long)
    On Error Resume Next
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso Is Nothing Then Exit Sub
    Dim dir As String
    dir = InstancesDir()
    If Not fso.FolderExists(dir) Then Exit Sub
    ' Pass 1 decides from a SNAPSHOT of the folder: deleting while enumerating a
    ' live Files collection is not defined behaviour.
    Dim doomedPath As New Collection
    Dim doomedName As New Collection
    Dim badName As String
    Dim f As Object
    Dim pid As Long
    Dim kept As Long
    Dim unreadable As Long
    kept = 0
    unreadable = 0
    badName = ""
    For Each f In fso.GetFolder(dir).Files
        If UtilInstanceFileName(fso, f.Name) Then
            pid = 0
            If Not InstanceFilePid(fso, f.Path, pid) Then
                ' KEEP it: we do not understand this entry, so we never delete it.
                unreadable = unreadable + 1
                If Len(badName) < 300 Then
                    If Len(badName) > 0 Then badName = badName & ", "
                    badName = badName & f.Name
                End If
            ElseIf pid = ownPid Then
                kept = kept + 1
            ElseIf UtilPidAlive(pid) Then
                kept = kept + 1
            Else
                doomedPath.Add f.Path
                doomedName.Add f.Name
            End If
        End If
    Next
    ' Pass 2 deletes only what pass 1 proved dead; each delete is best effort.
    Dim i As Long
    Dim n As Long
    Dim prunedName As String
    n = 0
    prunedName = ""
    For i = 1 To doomedPath.Count
        If Len(prunedName) < 400 Then
            If Len(prunedName) > 0 Then prunedName = prunedName & ", "
            prunedName = prunedName & doomedName.Item(i)
        End If
        Err.Clear
        fso.DeleteFile doomedPath.Item(i), True
        If Err.Number = 0 Then n = n + 1 Else Err.Clear
    Next
    LogLine "Pruned " & CStr(n) & " stale instance registration(s): " & prunedName
    If unreadable > 0 Then LogLine "Kept " & CStr(unreadable) & " unreadable instance registration(s): " & badName
End Sub

' TRUE only for an instance-registry FILE NAME: digits only, .json suffix. Any
' other file in the folder is ignored entirely.
Private Function UtilInstanceFileName(ByVal fso As Object, ByVal fname As String) As Boolean
    On Error Resume Next
    If fso Is Nothing Then Exit Function
    If LCase$(fso.GetExtensionName(fname)) <> "json" Then Exit Function
    Dim b As String
    b = fso.GetBaseName(fname)
    If Len(b) = 0 Then Exit Function
    Dim i As Long
    Dim ch As String
    For i = 1 To Len(b)
        ch = Mid$(b, i, 1)
        If ch < "0" Or ch > "9" Then Exit Function
    Next
    UtilInstanceFileName = True
End Function

' Reads the "pid" field of ONE registry file. Returns False whenever the content
' cannot be trusted: unreadable, empty, not JSON, no digits-only "pid", a
' non-positive pid, or a pid that disagrees with the <pid>.json file name. The
' caller MUST treat False as "KEEP this file".
Private Function InstanceFilePid(ByVal fso As Object, ByVal path As String, ByRef pidOut As Long) As Boolean
    On Error Resume Next
    pidOut = 0
    If fso Is Nothing Then Exit Function
    Dim ts As Object
    Set ts = fso.OpenTextFile(path, 1)
    If ts Is Nothing Then Exit Function
    Dim s As String
    s = ""
    Do While Not ts.AtEndOfStream
        s = s & ts.ReadLine
        If Len(s) > 100000 Then Exit Do
    Loop
    ts.Close
    If Len(s) = 0 Then Exit Function
    Dim root As Variant
    If Not JsonParse(s, root) Then Exit Function
    If Not IsObject(root) Then Exit Function
    Dim d As Object
    Set d = root
    Dim v As Variant
    Dim isObj As Boolean
    isObj = False
    If Not JsonGet(d, "pid", v, isObj) Then Exit Function
    If isObj Then Exit Function
    If IsNull(v) Then Exit Function
    ' The pid field must be DIGITS ONLY: Val() alone is lenient ("12abc" -> 12), and
    ' the promise above is that anything we cannot parse exactly is KEPT.
    Dim sv As String
    sv = Trim$(CStr(v))
    If Len(sv) = 0 Then Exit Function
    Dim i As Long
    Dim ch As String
    For i = 1 To Len(sv)
        ch = Mid$(sv, i, 1)
        If ch < "0" Or ch > "9" Then Exit Function
    Next
    Dim p As Long
    p = 0
    p = CLng(Val(sv))
    If p <= 0 Then Exit Function
    ' WriteInstanceFile always publishes <pid>.json with d("pid") = pid, so a name
    ' that disagrees is a corrupt or foreign file - never deleted.
    If StrComp(fso.GetBaseName(path), CStr(p), vbTextCompare) <> 0 Then Exit Function
    pidOut = p
    InstanceFilePid = True
End Function

' TRUE only when the pid is PROVABLY gone. Every inconclusive outcome counts as
' ALIVE (fail closed): a live registration must never be deleted on doubt.
Private Function UtilPidAlive(ByVal pid As Long) As Boolean
    On Error Resume Next
    UtilPidAlive = True
    If pid <= 0 Then
        UtilPidAlive = False
        Exit Function
    End If
    Dim r As Long
    r = UtilPidInSnapshot(pid)
    If r = 1 Then Exit Function                    ' in the snapshot => alive
    If r = 0 Then                                  ' snapshot read OK and pid absent => gone
        UtilPidAlive = False
        Exit Function
    End If
    ' r = -1: no trustworthy snapshot. Fall back to OpenProcess, which can only
    ' make the verdict MORE conservative (any failure we cannot attribute to a
    ' missing pid keeps the entry).
    Dim h As Long
    h = UtilOpenProcess(UTIL_PROCESS_QUERY_LIMITED, 0, pid)
    If h = 0 Then
        Dim le As Long
        le = Err.LastDllError
        If le = UTIL_ERROR_INVALID_PARAMETER Then UtilPidAlive = False
        Exit Function
    End If
    Dim code As Long
    code = 0
    If UtilGetExitCodeProcess(h, code) <> 0 Then
        UtilPidAlive = (code = UTIL_STILL_ACTIVE)
    End If
    UtilCloseHandle h
End Function

' 1 = the pid is in the Toolhelp32 process snapshot (ALIVE)
' 0 = the snapshot was read and the pid is NOT in it (provably GONE)
' -1 = the snapshot could not be read (UNKNOWN => the caller KEEPS the entry)
Private Function UtilPidInSnapshot(ByVal pid As Long) As Long
    On Error Resume Next
    UtilPidInSnapshot = -1
    Dim snap As Long
    snap = UtilCreateToolhelp32Snapshot(UTIL_TH32CS_SNAPPROCESS, 0)
    If snap = 0 Or snap = UTIL_INVALID_HANDLE Then Exit Function
    Dim pe As UTIL_PROCESSENTRY32
    pe.dwSize = Len(pe)
    Dim n As Long
    Dim hit As Boolean
    n = 0
    hit = False
    If UtilProcess32First(snap, pe) <> 0 Then
        Do
            n = n + 1
            If pe.th32ProcessID = pid Then hit = True
        Loop While UtilProcess32Next(snap, pe) <> 0
    End If
    UtilCloseHandle snap
    ' An enumeration that returned fewer than two processes is not evidence that
    ' anything is dead (a struct-layout mistake must never turn into a delete).
    If n < 2 Then Exit Function
    If hit Then
        UtilPidInSnapshot = 1
    Else
        UtilPidInSnapshot = 0
    End If
End Function

' Structured bind conflict: a second IDE on the SAME explicit port must fail loudly.
Public Sub WriteBindConflict(ByVal reason As String, ByVal tried As String, ByVal advice As String)
    On Error Resume Next
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso Is Nothing Then Exit Sub
    EnsureFolder fso, Vb6mcpDir()
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("ok") = False
    d("reason") = reason
    d("pid") = UtilProcessId()
    d("tried") = tried
    d("advice") = advice
    d("at") = Format$(Now, "yyyy-mm-dd hh:nn:ss")
    Dim ts As Object
    Set ts = fso.CreateTextFile(Vb6mcpDir() & "\bind_conflict.json", True)
    ts.WriteLine JsonStringify(d)
    ts.Close
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

Public Sub ClearPortFileIfOurs(ByVal port As Long)
    ' delete the port file only when it still points at our port
    ' (another VB6 instance may have written its own port meanwhile)
    On Error Resume Next
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim path As String
    path = TempDir() & "\vb6mcp_port.txt"
    If fso.FileExists(path) Then
        Dim ts As Object
        Set ts = fso.OpenTextFile(path, 1)
        Dim s As String
        s = Trim$(ts.ReadLine)
        ts.Close
        If s = CStr(port) Then fso.DeleteFile path, True
    End If
End Sub

Public Function TempDir() As String
    TempDir = Environ$("TEMP")
    If TempDir = "" Then TempDir = Environ$("TMP")
    If TempDir = "" Then TempDir = "C:\"
End Function

Public Function FindBarControl(ByVal vbe As Object, ByVal barCands As String, ByVal ctrlCands As String) As Object
    ' Two passes, because the menu bar NAME is localized: on a Chinese IDE it
    ' is the local word for "menu bar", so a bar-name filter never matches and
    ' the whole lookup died before reaching any control - which surfaced as
    ' "Menu command not found: Start|..." even when the item existed.
    '   1) restrict to bars whose name matches barCands (precise, fast)
    '   2) ignore the bar name entirely
    Set FindBarControl = FindBarControlPass(vbe, barCands, ctrlCands, True)
    If Not FindBarControl Is Nothing Then Exit Function
    Set FindBarControl = FindBarControlPass(vbe, barCands, ctrlCands, False)
End Function

Public Function FindBarControlPass(ByVal vbe As Object, ByVal barCands As String, ByVal ctrlCands As String, ByVal filterBar As Boolean) As Object
    ' Returns the matching CommandBarControl (searching bars and one level of
    ' popup sub-menus) or Nothing.
    ' Candidates are pipe-separated names covering English and localized
    ' (Chinese) captions; "(&x)" accelerators are ignored and candidates
    ' also prefix-match (e.g. "Make" matches "Make Project1.exe...").
    On Error Resume Next
    Dim bars As Object
    Set bars = vbe.CommandBars
    If bars Is Nothing Then Exit Function
    Dim nb As Long
    nb = bars.Count
    Dim bi As Long
    Dim b As Object
    For bi = 1 To nb
        Set b = bars.Item(bi)
        If Not b Is Nothing Then
            Dim barOk As Boolean
            barOk = MatchCand(b.Name, barCands)
            ' bar= names a BAR *or a menu-path segment*. The id path already
            ' honoured an ancestor on the path; the caption path did not, so a caller who
            ' copied the 'path' the list published ('菜单条>文件(&F)>...') could not use
            ' its middle segment. Matching power only GROWS here.
            Dim gateBar As Boolean
            gateBar = (filterBar And Len(barCands) > 0)
            If (Not filterBar) Or barOk Then
                Dim nc As Long
                nc = b.Controls.Count
                Dim ci As Long
                Dim c As Object
                For ci = 1 To nc
                    Set c = b.Controls.Item(ci)
                    If Not c Is Nothing Then
                        If MatchCand(c.Caption, ctrlCands) Then
                            Set FindBarControlPass = c
                            Exit Function
                        End If
                        Dim underOk As Boolean
                        underOk = barOk Or MatchCand(c.Caption, barCands)
                        ' one level of popup sub-menus (buttons have no
                        ' child controls: touching .Controls on them
                        ' raises 438 - guarded access)
                        Dim subc As Object
                        Dim subCol As Object
                        Err.Clear
                        Set subCol = c.Controls
                        If Not subCol Is Nothing Then
                            Dim ns As Long
                            ns = subCol.Count
                            Dim si As Long
                            For si = 1 To ns
                                Set subc = subCol.Item(si)
                                If Not subc Is Nothing Then
                                    If (Not gateBar) Or underOk Or MatchCand(subc.Caption, barCands) Then
                                        If MatchCand(subc.Caption, ctrlCands) Then
                                            Set FindBarControlPass = subc
                                            Exit Function
                                        End If
                                    End If
                                End If
                            Next
                        End If
                    End If
                Next
            End If
        End If
    Next
End Function

' THE normaliser. Both the caption and every candidate go through this, so
' 'what the list prints' and 'what execute accepts' cannot drift apart again.
'   * the accelerator group '(&X)' is decoration: everything from the first '(' is dropped;
'   * a trailing ellipsis ('...' or U+2026) is decoration too, NOT identity:
'     '调用堆栈...' names exactly the command '调用堆栈(&K)...' does;
'   * everything is trimmed.
' Public so the list side can group controls by the SAME key the matcher uses
' (a second copy of this normalisation is exactly how 'listed' and 'executable' drift apart).
Public Function CandBase(ByVal s As String) As String
    Dim t As String
    t = Replace(s, "&", "")
    Dim p As Long
    p = InStr(t, "(")
    If p > 0 Then t = Left$(t, p - 1)
    t = Trim$(t)
    Do While Len(t) >= 3
        If Right$(t, 3) <> "..." Then Exit Do
        t = Trim$(Left$(t, Len(t) - 3))
    Loop
    Do While Len(t) >= 1
        If Right$(t, 1) <> ChrW$(&H2026) Then Exit Do
        t = Trim$(Left$(t, Len(t) - 1))
    Loop
    CandBase = t
End Function

Public Function MatchCand(ByVal caption As String, ByVal cands As String) As Boolean
    ' ====================================================================
    ' (C4) MEASURED DEFECT: the CAPTION was normalised (accelerator
    ' "&" removed, everything from the first "(" dropped) but the CANDIDATE
    ' was not. command_bars{action=list} prints the RAW caption, e.g.
    '     生成工程1.exe(&K)...   or   调用堆栈(&K)...
    ' so passing back the very string the list published could never match -
    ' '列得出来' but '打不中' (menu_control_not_found). Both sides now go
    ' through the SAME normalisation, so the listed caption matches verbatim
    ' and every shorter form that used to work keeps working.
    ' ====================================================================
    Dim base As String
    base = CandBase(caption)
    If Len(base) = 0 Then Exit Function
    Dim pos As Long
    Dim sep As Long
    Dim cand As String
    pos = 1
    Do While pos <= Len(cands)
        sep = InStr(pos, cands, "|")
        If sep = 0 Then sep = Len(cands) + 1
        cand = Trim$(Mid$(cands, pos, sep - pos))
        If Len(cand) > 0 Then
            ' BOTH sides go through the ONE normaliser - the accelerator group
            ' AND a trailing ellipsis. Measured residual: '调用堆栈...' (accelerator removed,
            ' ellipsis kept) was still unmatchable, because the ellipsis was only ever
            ' dropped as a side effect of cutting at '(' - which a caller who strips the
            ' accelerator does not trigger.
            cand = CandBase(cand)
            If StrComp(base, cand, vbTextCompare) = 0 Then
                MatchCand = True
                Exit Function
            End If
            If Len(base) >= Len(cand) Then
                If StrComp(Left$(base, Len(cand)), cand, vbTextCompare) = 0 Then
                    MatchCand = True
                    Exit Function
                End If
            End If
        End If
        pos = sep + 1
    Loop
End Function

' ---------------------------------------------------------------------------
' (i18n): Id-first command lookup.
'
' Until now a menu item could only be identified by its LOCALIZED caption, so a
' VB6 IDE in any language other than English/Chinese lost every menu command.
' The CommandBars model also exposes a language-independent Id - but
' measured that VB6 Ids are NOT unique: id=746 alone covers 986 nodes (including
' both "Make" and "Start With Full Compile"), and Start/Continue share id=186.
' A bare Id query would therefore hit the wrong control.
'
' Rule enforced here (per the Id-usage ruling): the Id is NEVER used
' on its own. It is AND-ed with the shortcut text when one is known, and
' otherwise with the caption candidates. If no Id-verified candidate is found we
' fall back to the original localized-caption lookup, so the pre-existing
' candidate lists keep working and nothing that used to be found is lost.
' ---------------------------------------------------------------------------
Public Function FindBarControlById(ByVal vbe As Object, ByVal barCands As String, _
                                   ByVal ctrlCands As String, ByVal wantId As Long, _
                                   ByVal wantShortcut As String) As Object
    ' Id-first, using the SAME two passes as the caption matcher: first scoped to
    ' bars matching barCands (cheap), then across every bar. Scope matters - an
    ' unscoped walk has to touch all ~2300 CommandBars nodes on every call.
    If wantId <> 0 Then
        Dim v As Object
        Set v = FindByIdPass(vbe, barCands, ctrlCands, wantId, wantShortcut, True)
        If Not v Is Nothing Then
            Set FindBarControlById = v
            Exit Function
        End If
        Set v = FindByIdPass(vbe, barCands, ctrlCands, wantId, wantShortcut, False)
        If Not v Is Nothing Then
            Set FindBarControlById = v
            Exit Function
        End If
    End If
    ' last fallback: the localized-caption matcher, unchanged
    Set FindBarControlById = FindBarControl(vbe, barCands, ctrlCands)
End Function

Private Function FindByIdPass(ByVal vbe As Object, ByVal barCands As String, ByVal ctrlCands As String, _
                              ByVal wantId As Long, ByVal wantShortcut As String, _
                              ByVal filterBar As Boolean) As Object
    On Error Resume Next
    Dim bars As Object
    Set bars = vbe.CommandBars
    If bars Is Nothing Then Exit Function
    Dim nb As Long
    Dim bi As Long
    Dim b As Object
    nb = bars.Count
    For bi = 1 To nb
        Set b = bars.Item(bi)
        If Not b Is Nothing Then
            If (Not filterBar) Or MatchCand(b.Name, barCands) Then
                Dim nc As Long
                Dim ci As Long
                Dim c As Object
                nc = b.Controls.Count
                For ci = 1 To nc
                    Set c = b.Controls.Item(ci)
                    If Not c Is Nothing Then
                        If CtrlMatches(c, wantId, wantShortcut, ctrlCands) Then
                            Set FindByIdPass = c
                            Exit Function
                        End If
                        ' one level of popup sub-menus; buttons raise 438 on .Controls
                        Dim subCol As Object
                        Set subCol = Nothing
                        Err.Clear
                        Set subCol = c.Controls
                        If Not subCol Is Nothing Then
                            Dim ns As Long
                            Dim si As Long
                            Dim subc As Object
                            ns = subCol.Count
                            For si = 1 To ns
                                Set subc = subCol.Item(si)
                                If Not subc Is Nothing Then
                                    If CtrlMatches(subc, wantId, wantShortcut, ctrlCands) Then
                                        Set FindByIdPass = subc
                                        Exit Function
                                    End If
                                End If
                            Next
                        End If
                    End If
                Next
            End If
        End If
    Next
End Function

Private Function CtrlMatches(ByVal c As Object, ByVal wantId As Long, _
                             ByVal wantShortcut As String, ByVal ctrlCands As String) As Boolean
    On Error Resume Next
    Dim gId As Long
    gId = -1
    gId = c.Id
    If gId <> wantId Then Exit Function
    If Len(wantShortcut) > 0 Then
        Dim sh As String
        sh = ""
        sh = c.ShortcutText
        If StrComp(Trim$(sh), Trim$(wantShortcut), vbTextCompare) <> 0 Then Exit Function
    Else
        ' No shortcut is known for this command, so the caption must still agree.
        ' This is what makes an ambiguous Id such as 746 safe to refuse.
        If Not MatchCand(c.Caption, ctrlCands) Then Exit Function
    End If
    CtrlMatches = True
End Function

' ---------------------------------------------------------------------------
' Id-ONLY command lookup, used by command_bars when the caller
' addresses an item as "#746" / "id:746" / "746".
'
' The existing FindBarControlById is Id-AND-(shortcut or caption), because a
' bare Id is ambiguous. A command_bars caller who only knows the Id has no
' caption to AND with, so this variant answers the ambiguity openly: it returns
' a control ONLY when exactly ONE ENABLED control matches, and otherwise leaves
' matchCount and a short candidate list for the caller to report. Refusing to
' guess is the feature, not a limitation.
' ---------------------------------------------------------------------------
Public Function FindBarControlByIdOnly(ByVal vbe As Object, ByVal barCands As String, _
        ByVal capCands As String, ByVal wantId As Long, _
        ByRef matchCount As Long, ByRef candsOut As String, _
        Optional ByRef disabledCount As Long) As Object
    On Error Resume Next
    matchCount = 0
    candsOut = ""
    disabledCount = 0
    Dim v As Object
    If Len(barCands) > 0 Then
        ' pass 1: the caller's bar argument must match the BAR NAME or a node on
        ' the control's menu PATH (e.g. bar="工程" for 工程>部件).
        Set v = FindByIdOnlyPass(vbe, barCands, capCands, wantId, matchCount, candsOut, disabledCount)
        If Not v Is Nothing Then
            Set FindBarControlByIdOnly = v
            Exit Function
        End If
        If matchCount > 0 Then Exit Function   ' ambiguous inside that scope: do not widen
        ' pass 2: the bar name is localized, so retry with no bar scope
        Set v = FindByIdOnlyPass(vbe, "", capCands, wantId, matchCount, candsOut, disabledCount)
    Else
        Set v = FindByIdOnlyPass(vbe, "", capCands, wantId, matchCount, candsOut, disabledCount)
    End If
    If Not v Is Nothing Then Set FindBarControlByIdOnly = v
End Function

Private Function FindByIdOnlyPass(ByVal vbe As Object, ByVal barCands As String, ByVal capCands As String, _
        ByVal wantId As Long, ByRef matchCount As Long, ByRef candsOut As String, _
        ByRef disabledCount As Long) As Object
    On Error Resume Next
    Dim bars As Object
    Set bars = vbe.CommandBars
    If bars Is Nothing Then Exit Function
    Dim nb As Long, bi As Long, b As Object, hit As Object, under As Boolean
    nb = bars.Count
    Set hit = Nothing
    For bi = 1 To nb
        Set b = bars.Item(bi)
        If Not b Is Nothing Then
            under = True
            If Len(barCands) > 0 Then under = MatchCand(b.Name, barCands)
            FindByIdOnlyIn b.Controls, b.Name, 0, wantId, barCands, capCands, under, hit, matchCount, candsOut, disabledCount
        End If
    Next
    If matchCount = 1 Then Set FindByIdOnlyPass = hit
End Function

' Qualify = Id matches AND enabled AND (the caller's bar argument matched the bar
' name or an ancestor on the menu PATH) AND (when the caller also gave a caption,
' the caption matches). The caption conjunct is REQUIRED: measured
' that id=746 has 525 ENABLED matches across the menus and still 6 inside the
' 工程 submenu alone, so an id on its own cannot name a control.
Private Sub FindByIdOnlyIn(ByVal col As Object, ByVal barName As String, ByVal depth As Long, _
        ByVal wantId As Long, ByVal barCands As String, ByVal capCands As String, ByVal under As Boolean, _
        ByRef hit As Object, ByRef matchCount As Long, ByRef candsOut As String, _
        Optional ByRef disabledCount As Long)
    On Error Resume Next
    If col Is Nothing Then Exit Sub
    If depth > 2 Then Exit Sub
    Dim n As Long, k As Long, c As Object, subCol As Object
    Dim selfOk As Boolean, cap As String
    n = col.Count
    For k = 1 To n
        Set c = col.Item(k)
        If Not c Is Nothing Then
            cap = ""
            cap = c.Caption
            selfOk = under
            If Len(barCands) = 0 Then selfOk = True
            If Not selfOk Then selfOk = MatchCand(cap, barCands)
            Dim gid As Long, en As Boolean, qual As Boolean
            gid = -1
            en = False
            gid = c.Id
            en = c.Enabled
            qual = (gid = wantId) And en And selfOk
            If qual And Len(capCands) > 0 Then qual = MatchCand(cap, capCands)
            If qual Then
                matchCount = matchCount + 1
                If hit Is Nothing Then Set hit = c
                If Len(candsOut) < 300 Then
                    If Len(candsOut) > 0 Then candsOut = candsOut & "; "
                    candsOut = candsOut & barName & ">" & cap
                End If
            Else
                ' the control may EXIST and simply be DISABLED. Counting it separately
                ' lets command_bars answer 'disabled' instead of the false 'not found'.
                Dim qualOff As Boolean
                qualOff = (gid = wantId) And Not en And selfOk
                If qualOff And Len(capCands) > 0 Then qualOff = MatchCand(cap, capCands)
                If qualOff Then disabledCount = disabledCount + 1
            End If
            Set subCol = Nothing
            Err.Clear
            Set subCol = c.Controls
            If Not subCol Is Nothing Then
                FindByIdOnlyIn subCol, barName, depth + 1, wantId, barCands, capCands, selfOk, hit, matchCount, candsOut, disabledCount
            End If
        End If
    Next
End Sub


' (kernel32 GetTickCount is declared in the declarations section above - VB6
'  only accepts Declare statements before the first procedure.)
' ===========================================================================
' THE unified message-pump point.
'
' WHY. VB6 hosts this add-in on the IDE's single UI thread (STA). While a tool
' call computes, the IDE message queue is not serviced, so the IDE UI - and
' every other MCP request - is frozen until the call returns. Measured BEFORE
' this fix on a 3 x 26,999-line fixture: a 3.2 s verify_scope made a concurrent
' ide_state wait 2.9 s and return at the very instant the long call ended
' (内部证据). Long loops therefore hand the
' queue back periodically.
'
' WHAT. Exactly ONE helper. Rule enforced by grep: no module may call DoEvents
' directly (the only other site is the pre-existing control_execution wait,
' which is deliberate and documented there). Every pump point is the literal
' call MaybePump, so the set of pump points is auditable in one grep.
'
' PERIOD. Every PUMP_EVERY_ITER iterations OR every PUMP_EVERY_MS milliseconds,
' whichever comes first - the user asked for a PERIOD, not a call per tick.
' Both are named constants so the policy can be retuned in one place. The tick
' is only sampled every 64 iterations, so the per-iteration cost is two Long
' increments.
'
' RE-ENTRANCY. A pump never nests (m_pumpDepth): once DoEvents is running, a
' second DoEvents would recurse the message loop. The tool-level guard lives in
' CodeTools.CallTool, which refuses a second state-changing call while one is
' in flight.
'
' NOT FOR. Error handlers, and critical sections covered by On Error Resume
' Next whose state must not be disturbed.
' ===========================================================================


' Start a fresh period window (called once per top-level tool call).
Public Sub PumpReset()
    m_pumpIter = 0
    m_pumpTimeIter = 0
    m_pumpLastAt = PumpTickCount()
End Sub

Public Sub MaybePump()
    ' A pump inside a pump must never happen.
    If m_pumpDepth > 0 Then Exit Sub
    m_pumpIter = m_pumpIter + 1
    If m_pumpIter >= PUMP_EVERY_ITER Then
        PumpNow
        Exit Sub
    End If
    m_pumpTimeIter = m_pumpTimeIter + 1
    If (m_pumpTimeIter And PUMP_TICK_MASK) = 0 Then
        If (PumpTickCount() - m_pumpLastAt) >= PUMP_EVERY_MS Then PumpNow
    End If
End Sub

Public Sub PumpNow()
    m_pumpIter = 0
    m_pumpTimeIter = 0
    m_pumpDepth = m_pumpDepth + 1
    m_pumpCount = m_pumpCount + 1
    Dim t0 As Long
    t0 = PumpTickCount()
    m_pumpLastAt = t0
    ' ==== THE ONE GATE ====
    ' A host modal box that is UP turns DoEvents into the wedge path: the pump dispatches
    ' whatever the host has POSTED (measured VB6's File>Make command arriving that
    ' way), the host opens its modal box, this single-threaded add-in is suspended inside
    ' that modal loop and NEVER returns - and inside it NO watchdog tick is delivered, so
    ' neither the one-shot confirm nor the generic modal guard can rescue it.
    ' The TICK IS STILL RUN: WinOps.WheelTick below keeps driving the guard (it presses
    ' synchronously, it does not need a message pump), so withholding DoEvents costs
    ' OBSERVATION of the dismissal, never the ability to act.
    If PumpGateBlocked() Then
        m_gateSkips = m_gateSkips + 1
    Else
        DoEvents
    End If
    ' throttle-rescan the optional code-pane wheel. It returns
    ' immediately while the wheel is OFF (the default), so the pump cost is
    ' unchanged by default and no hook is ever installed.
    WinOps.WheelTick
    Dim dt As Long
    dt = PumpTickCount() - t0
    m_pumpMs = m_pumpMs + dt
    If dt > m_pumpMaxMs Then m_pumpMaxMs = dt
    m_pumpDepth = m_pumpDepth - 1
End Sub

Public Function PumpCount() As Long
    PumpCount = m_pumpCount
End Function

Public Function PumpMsTotal() As Long
    PumpMsTotal = m_pumpMs
End Function

Public Function PumpMsMax() As Long
    PumpMsMax = m_pumpMaxMs
End Function

' THE gate. True = a visible top-level #32770 of THIS process is up, so a pump
' would be unsafe. Cached for PUMP_GATE_CACHE_MS because the pump runs roughly every
' PUMP_EVERY_MS during long loops and a window walk is not free.
Public Function PumpGateBlocked() As Boolean
    Dim now As Long
    now = PumpTickCount()
    If m_gateAt <> 0 Then
        If (now - m_gateAt) < PUMP_GATE_CACHE_MS Then
            PumpGateBlocked = m_gateBlocked
            Exit Function
        End If
    End If
    m_gateAt = now
    m_gateTitle = ""
    m_gateBlocked = WinOps.HasVisibleHostDialog(m_gateTitle)
    PumpGateBlocked = m_gateBlocked
End Function

' The same verdict, phrased for the DIRECT DoEvents call sites (there are six outside this
' module): they must ask before pumping. Rule: no module calls DoEvents unguarded.
Public Function CanPump() As Boolean
    CanPump = Not PumpGateBlocked()
End Function

Public Function PumpGateSkipped() As Long
    PumpGateSkipped = m_gateSkips
End Function

' ONE source of truth for the pump/gate numbers. The object is what a receipt should
' carry (a nested BLOCK, so a caller does not have to parse a JSON string inside JSON).
Public Function PumpStatsObj() As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("calls") = m_pumpCount
    d("msInPump") = m_pumpMs
    d("maxPumpMs") = m_pumpMaxMs
    d("iterPeriod") = PUMP_EVERY_ITER
    d("msPeriod") = PUMP_EVERY_MS
    ' the gate is part of the measurement, not a hidden policy.
    d("gateSkips") = m_gateSkips
    d("gateBlockedNow") = m_gateBlocked
    d("gateTitle") = m_gateTitle
    d("gateNote") = "gateSkips counts the pumps WITHHELD because a visible #32770 of this process was up: with a host modal box up, DoEvents would dispatch the host's POSTED command and enter its modal loop, from which this single-threaded add-in never returns (measured). The tick still runs, so the guard can still act; only the DoEvents is withheld."
    Set PumpStatsObj = d
End Function

Public Function PumpStatsJson() As String
    PumpStatsJson = JsonStringify(PumpStatsObj())
End Function

' Read-and-reset, for a measurement run.
Public Function PumpStatsConsume() As String
    Dim s As String
    s = PumpStatsJson()
    m_pumpCount = 0
    m_gateSkips = 0
    m_pumpMs = 0
    m_pumpMaxMs = 0
    PumpStatsConsume = s
End Function

' item 6: record a .bak this call created (VbFmt.bas calls it). The collection
' itself is declared in the declarations section above; CodeTools resets it per call.
Public Sub SeNoteBackup(ByVal path As String)
    On Error Resume Next
    If Len(path) = 0 Then Exit Sub
    If g_seBackups Is Nothing Then Set g_seBackups = New Collection
    g_seBackups.Add path
    Err.Clear
End Sub

' ===========================================================================
' (consumer ledger -20): SHARED WALL-CLOCK BUDGET.
'
' WHY. dependency_graph and export_public_api/check_public_api scan the whole
' project on the IDE's single UI thread. On a real 110-component project the
' component-edge phase was O(components^2) whole-module string scans, so the call
' ran past the 60 s MCP client timeout and left the add-in permanently wedged
' (VB6 alive, Responding=False, every later call failing in 1-2 ms). A READ-ONLY
' tool must never be able to do that, so every whole-project scan now runs under
' a deadline and RETURNS WHAT IT HAS when the deadline expires - partial, with
' budgetExhausted / budgetMs / elapsedMs / stage in the receipt.
'
' OVERRIDE. VB6MCP_BUDGET_MS (a positive integer of milliseconds) overrides the
' per-tool default for one process; the default lives in DepGraph.bas.
'
' The clock is GetTickCount (already declared above, milliseconds since boot).
' Nothing here runs unless a budgeted scan starts, so no existing behaviour moves.
' ===========================================================================

Public Function TickMs() As Long
    TickMs = PumpTickCount()
End Function

' A positive Long from an environment variable, else the caller's default.
' Defensive on purpose: a malformed value must fall back, never raise.
Public Function EnvLongOr(ByVal name As String, ByVal dflt As Long) As Long
    Dim s As String, d As Double
    EnvLongOr = dflt
    On Error Resume Next
    s = Trim$(Environ$(name))
    If Len(s) = 0 Then Exit Function
    d = Val(s)
    If d < 1# Then Exit Function
    If d > 2147483000# Then d = 2147483000#
    EnvLongOr = CLng(d)
    Err.Clear
End Function
