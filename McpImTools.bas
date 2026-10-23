Attribute VB_Name = "McpImTools"
Option Explicit

' =======================================================================
' C-3 facade batch: McpImTools (moved bodies). The CLASS keeps same-signature Private
' forwarders, so no class call site changes; each facade delegates to this module.
' Owner convention: Owner As CodeTools is the FIRST parameter of every procedure here.
' =======================================================================

Private Declare Function WoGetTickCount Lib "kernel32" Alias "GetTickCount" () As Long

Private Const IM_FILE_NUM As Long = 511
Private Const IM_MAX_EXEC_LINE_CHARS As Long = 900
Private Const IM_SINK_MAX_BYTES As Long = 65536

Public Function ImAnsiByteCount(Owner As CodeTools, ByVal s As String) As Long
    ' FIX-25 (D2): the REAL size of the value in ANSI bytes (a GBK char is 2 bytes). Kept
    ' separate from valueBytes so an existing consumer of valueBytes (a CHARACTER count) does
    ' not silently change meaning.
    ImAnsiByteCount = 0
    If Len(s) = 0 Then Exit Function
    On Error Resume Next
    ' use the SAME encoder the write paths use (Encoding.EncodeText + ByteCount). StrConv could
    ' not be used: measured, StrConv(s, vbFromUnicode) gave 3 bytes for an ASCII 3-char string
    ' and even 0 when read through a Variant - a byte count that lies is worse than none.
    Dim b() As Byte
    Encoding.EncodeText s, "ansi", b
    If Err.Number <> 0 Then
        Err.Clear
        ImAnsiByteCount = Len(s)
        Exit Function
    End If
    ImAnsiByteCount = Encoding.ByteCount(b)
    Err.Clear
End Function

Public Function ImCountLines(Owner As CodeTools, ByVal s As String) As Long
    Dim t As String
    If Len(s) = 0 Then Exit Function
    t = ImJoinLines(Owner, s)
    If Right$(t, 1) = vbLf Then t = Left$(t, Len(t) - 1)
    If Len(t) = 0 Then Exit Function
    ImCountLines = UBound(Split(t, vbLf)) + 1
End Function

Public Function ImDeleteSink(Owner As CodeTools, ByVal p As String) As Boolean
    ImDeleteSink = False
    On Error Resume Next
    If Len(Dir$(p)) > 0 Then Kill p
    ImDeleteSink = (Len(Dir$(p)) = 0)
    Err.Clear
End Function

Public Function ImEscapeValue(Owner As CodeTools, ByVal s As String) As String
    ' The add-in escapes the value itself (no marker transport): backslash FIRST, then CR, then LF.
    Dim bs As String
    bs = Chr$(92)
    ImEscapeValue = Replace(Replace(Replace(s, bs, bs & bs), vbCr, bs & "r"), vbLf, bs & "n")
End Function

Public Function ImExecuteImmediate(Owner As CodeTools, ByVal mode As String, ByVal payload As String, ByVal timeoutMs As Long) As Object
    Dim o As Object
    Set o = CreateObject("Scripting.Dictionary")
    o("tool") = IIf(mode = "eval", "immediate_eval", "immediate_run")
    o("mode") = mode
    o("ok") = False
    o("reason") = ""
    o("valueParsed") = False
    o("valueKind") = "scalar"
    o("valueEncoding") = ""
    o("valueBytes") = 0
    o("valueBytesReturned") = 0
    o("valueLines") = 0
    o("valueComplete") = False
    o("valueEscaped") = ""
    o("valueTruncated") = False
    o("valueTruncatedReason") = ""
    o("valueSinkTruncated") = False
    o("valueTransport") = ""
    o("valueTransportNote") = ""
    o("outputDescribes") = "immediate_eval and immediate_run both carry output through the Debug.Print sink; value* describes the carried VALUE (valueTransport=debug_print for eval)"
    o("value") = ""
    o("typeName") = ""
    o("output") = ""
    o("outputTrimmed") = ""
    o("outputBytes") = 0
    o("outputTruncated") = False
    o("outputLines") = 0
    o("debugPrintRewrites") = 0
    o("sinkFile") = ""
    o("sinkDeleted") = False
    o("line") = ""
    o("elapsedMs") = 0
    o("execAnswered") = False
    o("ret") = 0
    o("retHex") = ""
    o("timeoutMs") = timeoutMs
    o("timeoutMsEnforcedBy") = "gate_age_threshold"
    o("hardTimeoutEnforced") = False
    o("statementInterrupted") = False
    o("armedReleaseMs") = timeoutMs
    o("recoveryAction") = "ide_state{release_inflight:true, release_min_held_ms:<timeout_ms>}"
    o("recoveryHint") = "the add-in cannot preempt a statement, so there is NO hard timeout here. If this call never returns, the ARMED release frees the single-flight gate on the next tool call."
    o("note") = ""
    If timeoutMs <= 0 Then timeoutMs = 5000
    o("timeoutMs") = timeoutMs
    o("armedReleaseMs") = timeoutMs
    If Len(payload) = 0 Then
        o("reason") = "empty_input"
        o("note") = "pass a non-empty " & IIf(mode = "eval", "expression", "code") & "; nothing was executed"
        Set ImExecuteImmediate = o
        Exit Function
    End If
    ' EVAL SAFETY SAFETY NET: if a runtime-error box from an EARLIER abandoned evaluation is
    ' STILL holding the IDE, do not add another evaluation to it. Read it, dismiss it (OK when
    ' it has no Debug caption - the evaluator box shape) and answer with the error IT carried.
    ' Such a call can never be rescued by the call that raised it: resolving the box unwinds the
    ' stack under it (measured: the call never returns).
    Dim stuckDlg As Object
    Set stuckDlg = McpRe.ReFindRuntimeErrorDialog(Owner)
    If Not stuckDlg Is Nothing Then
        Dim stuckRes As Object
        Set stuckRes = McpRe.ReCaptureRuntimeError(Owner, True)
        o("ok") = False
        o("stopped_at") = "previous_call_raised_and_blocked"
        o("reason") = "previous_call_raised_and_blocked"
        o("failureKind") = "statement_raised"
        o("refusedBeforeExecution") = True
        o("recoveredBox") = True
        o("note") = "a runtime-error box from an EARLIER add-in evaluation was holding the IDE: this call READ it and dismissed it, and did NOT execute its own input. The earlier call cannot return by itself - resolving the box abandons it."
        o("howToProceed") = "the box is resolved: re-run the same call. To avoid producing such a box at all, an evaluation whose root object is Empty/Nothing is now REFUSED before execution (reason=refused_would_raise)."
        If IsObject(stuckRes) Then
            If stuckRes.Exists("error") Then Set o("previousCallError") = stuckRes("error")
            If stuckRes.Exists("outcome") Then o("previousCallOutcome") = CStr(stuckRes("outcome"))
            If stuckRes.Exists("dismissedVia") Then o("dismissedVia") = CStr(stuckRes("dismissedVia"))
            If stuckRes.Exists("proof") Then Set o("previousCallProof") = stuckRes("proof")
        End If
        Set ImExecuteImmediate = o
        Exit Function
    End If
    If mode = "eval" Then
        Dim objProblem As String
        objProblem = ImObjectExprProblem(Owner, payload)
        If Len(objProblem) > 0 Then
            o("reason") = "object_expression_unsupported"
            o("rejectedToken") = objProblem
            o("guidance") = "object/array EXPRESSIONS are not supported by the immediate channel (measured: a plain object assignment hangs the call). Assign the object to a PROJECT variable first and read scalar members."
            o("howToRecover") = "nothing was executed by this call, so no IDE recovery is needed"
            o("note") = "REFUSED BEFORE EXECUTION: the expression was never sent to the immediate context."
            Set ImExecuteImmediate = o
            Exit Function
        End If
        ' EVAL SAFETY PRIMARY FIX: a raise in THIS context is not catchable (the box belongs to
        ' the IDE expression evaluator and resolving it abandons the call), so PREVENT it: probe
        ' the ROOT of a dotted expression with TypeName (cannot raise) and refuse a member
        ' access that is a GUARANTEED VB6 error.
        Dim risk As Object
        Dim riskErr As Object
        Set risk = ImWouldRaiseRisk(Owner, payload, timeoutMs)
        If Not risk Is Nothing Then
            Set riskErr = CreateObject("Scripting.Dictionary")
            riskErr("vbaErrNumber") = CLng(risk("errNumber"))
            riskErr("description") = CStr(risk("description"))
            riskErr("raisedIn") = "addin_precheck: NOT executed (a raise in this context is not catchable and abandons the call)"
            riskErr("errorAuthority") = "VB6 semantics + a MEASURED probe (" & CStr(risk("probeExpr")) & " -> " & CStr(risk("typeName")) & ")"
            o("ok") = False
            o("stopped_at") = "refused_would_raise"
            o("reason") = "refused_would_raise"
            o("failureKind") = "statement_raised"
            o("refusedBeforeExecution") = True
            o("wouldRaise") = CStr(risk("description"))
            Set o("error") = riskErr
            Set o("precheck") = risk
            o("note") = "REFUSED BEFORE EXECUTION: member access on " & CStr(risk("root")) & " is a guaranteed VB6 error here (probe " & CStr(risk("probeExpr")) & " = " & CStr(risk("typeName")) & "), and a raise in this evaluation context escapes as a modal IDE box that the in-flight call can never return from. Nothing was sent to the immediate context."
            o("howToProceed") = "assign a real object first (Set " & CStr(risk("root")) & " = ...) or evaluate a scalar that does not dereference it"
            o("howToRecover") = "nothing was executed by this call, so no IDE recovery is needed"
            Set ImExecuteImmediate = o
            Exit Function
        End If
    End If
    Dim flat As String
    Dim rewrites As Long
    Dim body As String
    Dim sink As String
    flat = ImFlatLine(Owner, payload)
    o("statementChars") = Len(flat)
    o("statementBytes") = ImAnsiByteCount(Owner, flat)
    rewrites = 0

    Dim execLine As String
    Dim fn As String
    sink = ImSinkPath(Owner, mode)
    o("sinkFile") = sink
    On Error Resume Next
    If Len(Dir$(sink)) > 0 Then Kill sink
    On Error GoTo 0
    fn = "#511"
    ' FIX-25 (D4): a RAISED statement leaves the sink handle OPEN (its trailing Close #511 never
    ' ran), and every later call then failed on Open #511 for ~19 s. Probe-close a remembered
    ' stale handle HERE - a Close on a handle that IS open is valid - and report what happened.
    Dim staleNum As Long, staleRet As Long, staleClosed As Boolean, staleProbeRaw As String
    staleNum = Owner.m_imStaleSink
    staleRet = -1
    staleClosed = False
    staleProbeRaw = ""
    ' FIX-29: the entry ALWAYS closes #511 first - it no longer depends on OUR record of
    ' what happened last time. Measured failure it prevents: a 2048-char call made the VB6 process
    ' disappear and restart, leaving #511 open with a modal runtime error 55 ("file already open");
    ' every later immediate_* call then failed until the handle was closed by hand. A Close on a
    ' handle that is NOT open raises, which is tolerated here (we only need the attempt + result).
    Dim entryCloseRaw As String
    Dim entryCloseRet As Long
    entryCloseRaw = ""
    entryCloseRet = -1
    On Error Resume Next
    Err.Clear
    entryCloseRaw = Vba6Probe.ExecLineJson("Close #511", 0, 0, 0, False, False, False)
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
    entryCloseRet = 0
    If Len(entryCloseRaw) > 0 Then entryCloseRet = CLng(Val(Owner.JsonFieldRaw(entryCloseRaw, "ret")))
    Dim entryCloseAttempted As Boolean
    Dim entryCloseClosed As Boolean
    entryCloseAttempted = True
    entryCloseClosed = (entryCloseRet = 0)
    If entryCloseRet = 0 Then
        staleClosed = True
        Owner.m_imStaleSink = 0
        Owner.m_imStaleAt = 0
    End If
    staleRet = entryCloseRet
    o("sinkEntryCloseAttempted") = entryCloseAttempted
    o("sinkEntryCloseRet") = entryCloseRet
    o("sinkEntryCloseClosed") = entryCloseClosed
    o("sinkEntryCloseNote") = "every immediate_* call closes the sink handle #511 BEFORE opening it (FIX-29). ret=0 means it was open and is now closed; a non-zero ret means it was not open (harmless) OR the line was refused - the call then proceeds to open its own sink."
    o("sinkStaleDetected") = (staleNum > 0)
    o("sinkStaleNumber") = staleNum
    o("sinkStaleClosed") = staleClosed
    o("sinkStaleCleanupRet") = staleRet
    o("sinkStaleAgeMs") = IIf(staleNum > 0, WoGetTickCount() - Owner.m_imStaleAt, 0)
    If staleNum > 0 Then
        If staleClosed Then
            o("sinkStaleNote") = "a previous immediate_* line raised and left the sink handle #" & CStr(staleNum) & " open. This call closed it (probe Close #" & CStr(staleNum) & " returned ret=0) BEFORE opening its own sink, so the channel is usable again."
        Else
            o("sinkStaleNote") = "a previous immediate_* line raised and may have left the sink handle #" & CStr(staleNum) & " open. The probe Close returned ret=" & CStr(staleRet) & ", so it could not be confirmed closed; this call still proceeded and its own Open may fail (engine_refused_line)."
            o("howToRecover") = "if this call fails with engine_refused_line, the sink handle is still open: use ide_state{release_inflight:true} / restart the IDE"
        End If
    End If
    ' WP-22-02: the immediate context belongs to the OPEN project, which has NO add-in modules.
    ' Any "Util.*" identifier in the exec line is therefore a COMPILE error (measured: ret=40036).
' the line uses ONLY the literal file number #511 and the DEBUG.PRINT channel that
    ' immediate_run already proves works; the add-in reads its own Debug.Print sink as the value.
    If mode = "eval" Then
        body = ImRewriteDebugPrint(Owner, "Debug.Print " & flat, rewrites, fn)
    Else
        body = ImRewriteDebugPrint(Owner, flat, rewrites, fn)
    End If
    o("debugPrintRewrites") = rewrites
    ' LITERAL file number, no variable: the immediate window executes in the OPEN project, which has
    ' Option Explicit and no add-in modules - an implicit variable (fn511) or a Util.* identifier is a
    ' COMPILE error there (measured: ret=40036 with errNumber=0, i.e. it never reached run time).
    ' FreeFile is unusable for the same reason: its result must be assigned to a variable.
    ' This path cannot re-enter (the add-in runs one tool call at a time), so a fixed #511 cannot
    ' collide with itself; do NOT prepend Close #511 - a leading Close on a number that is not open
    ' raises at that statement and refuses the whole line (measured).
    execLine = "Open """ & sink & """ For Output As #511: " & body & ": Close #511"
    ' FIX-31: the guard measures the FULL EXEC LINE (expression + wrapper). Measured on this IDE:
    ' <=900 characters stable (33/33); 901-998 failed 14/15 with 10x value_not_parsed and 4x
    ' engine_refused_line ret 40036 (non-monotonic: one 905 OK), which is what a ~1024-character
    ' immediate-window line buffer truncation looks like. Refusing here means the refusal SHAPE is
    ' the same one the execution path publishes (sinkEntry* included).
    o("execLineChars") = Len(execLine)
    o("expressionChars") = Len(flat)
    o("wrapperChars") = Len(execLine) - Len(flat)
    If Len(execLine) > IM_MAX_EXEC_LINE_CHARS Then
        o("reason") = "statement_too_large"
        o("limitChars") = IM_MAX_EXEC_LINE_CHARS
        o("limitBasis") = "full_exec_line"
        o("actualChars") = Len(execLine)
        o("actualBytes") = ImAnsiByteCount(Owner, execLine)
        o("limitNote") = "the limit applies to the line that is EXECUTED (Open <sink> For Output As #511: Print #511, <expr>: Close #511), which is " & CStr(Len(execLine)) & " characters; the wrapper alone costs " & CStr(Len(execLine) - Len(flat)) & ". Measured on this IDE: <=900 characters stable 33/33; 901-998 failed 14/15 (value_not_parsed / engine_refused_line ret 40036), so the safe LINE width is 900."
        o("expressionLimitHint") = "the usable EXPRESSION length in this environment is about " & CStr(IM_MAX_EXEC_LINE_CHARS - (Len(execLine) - Len(flat))) & " characters (the wrapper costs " & CStr(Len(execLine) - Len(flat)) & ")"
        o("note") = "the executed line would be " & CStr(Len(execLine)) & " characters, over the " & CStr(IM_MAX_EXEC_LINE_CHARS) & "-character line limit of this channel, so it is REFUSED BEFORE EXECUTION - nothing was sent to the immediate context."
        o("howToProceed") = "shorten the expression, or build the long literal INSIDE the project (assign it to a module variable) and read it back. NOTE: if a LONG line is attempted anyway, the IDE may need a restart afterwards - measured: a 2048-character line made the VB6 process disappear and restart and left a modal runtime error 55 (file already open) with the sink handle still held."
        o("sinkEntryCloseAttempted") = entryCloseAttempted
        o("sinkEntryCloseRet") = entryCloseRet
        o("sinkEntryCloseClosed") = entryCloseClosed
        o("sinkState") = "closed"
        Set ImExecuteImmediate = o
        Exit Function
    End If
    o("line") = execLine
    Owner.m_armTool = o("tool")
    Owner.m_armMs = timeoutMs
    Owner.m_armAt = WoGetTickCount()
    Dim t0 As Long
    t0 = WoGetTickCount()
    Dim raw As String
    Dim raised As Boolean
    Dim rNum As Long
    Dim rDesc As String
    raw = ""
    raised = False
    rNum = 0
    rDesc = ""
    On Error Resume Next
    Err.Clear
    raw = Vba6Probe.ExecLineJson(execLine, 0, 0, 0, False, False, False)
    If Err.Number <> 0 Then
        raised = True
        rNum = Err.Number
        rDesc = Err.Description
        Err.Clear
    End If
    On Error GoTo 0
    o("elapsedMs") = WoGetTickCount() - t0
    Owner.m_armMs = 0
    Owner.m_armTool = ""
    Owner.m_armAt = 0
    If raised Or Len(raw) = 0 Then
        o("reason") = "exec_error"
        o("execError") = "vbErrNumber=" & CStr(rNum) & "; description=" & rDesc
        o("note") = "the add-in DID answer, but the call into the exec probe raised: " & rDesc
        o("raw") = Left$(raw, 400)
    Else
        o("execAnswered") = True
        o("ret") = CLng(Val(Owner.JsonFieldRaw(raw, "ret")))
        o("retHex") = Owner.JsonFieldRaw(raw, "retHex")
        o("errNumber") = CLng(Val(Owner.JsonFieldRaw(raw, "vbaError")))
        o("errDescription") = Owner.JsonFieldRaw(raw, "vbaErrorDesc")
        ' EVAL SAFETY: a line that RAISED and WAS CAUGHT here is a statement_raised (the real
        ' vbaErrNumber is carried so a caller can act on it) - not a channel/refusal failure.
        If CLng(Val(Owner.JsonFieldRaw(raw, "vbaError"))) <> 0 Then
            Dim caughtB As Object
            Set caughtB = CreateObject("Scripting.Dictionary")
            caughtB("vbaErrNumber") = CLng(Val(Owner.JsonFieldRaw(raw, "vbaError")))
            caughtB("description") = Owner.JsonFieldRaw(raw, "vbaErrorDesc")
            caughtB("raisedIn") = "the executed line in the project immediate context (caught by the add-in error handler)"
            Set o("error") = caughtB
            o("failureKind") = "statement_raised"
            If Len(CStr(o("reason"))) = 0 Then o("reason") = "statement_raised"
        End If
        o("lineText") = execLine
        If o("ret") <> 0 Then
            o("reason") = "engine_refused_line"
            o("note") = "EbExecuteLine returned " & CStr(o("retHex")) & " (ret=" & CStr(o("ret")) & "), so the line was NOT executed. On this channel ret=40036 (0x9C64) means the line was REFUSED - most often a syntax/compile error inside the line (e.g. an undeclared implicit variable under Option Explicit) or On Error Resume Next."
        End If
    End If
    ' FIX-25 (D4): if the exec did not complete, the trailing Close #511 may not have run - record
    ' the handle so the NEXT call probe-closes it (and says so).
    Dim sinkExecOk As Boolean
    sinkExecOk = False
    If Not raised Then
        If Len(raw) > 0 Then
            If CLng(Val(Owner.JsonFieldRaw(raw, "ret"))) = 0 Then sinkExecOk = True
        End If
    End If
    If sinkExecOk Then
        If Owner.m_imStaleSink = IM_FILE_NUM Then
            Owner.m_imStaleSink = 0
            Owner.m_imStaleAt = 0
        End If
        o("sinkState") = "closed"
    Else
        Owner.m_imStaleSink = IM_FILE_NUM
        Owner.m_imStaleAt = WoGetTickCount()
        o("sinkState") = "left_open"
        o("sinkStateNote") = "the exec line did not complete (no answer, a VB raise, or a non-zero ret), so its trailing Close #511 may not have run. The handle is remembered and probe-closed on the NEXT immediate_* call, which reports sinkStale*."
    End If
    Dim txt As String
    Dim trunc As Boolean
    Dim garbageN As Long
    Dim nulPos As Long
    trunc = False
    txt = ImReadSink(Owner, sink, trunc)
    ' FIX-25 (D2): a NUL character in the decoded sink is a BYTE-coordinate artifact, never part
    ' of a VB6 value. Cut it at the first NUL and remember how much was cut; the receipt may not
    ' call the value complete while garbage was present (measured before the fix:
    ' value="中文值" + 3 NULs, valueBytes:7, valueLines:2, valueComplete:true).
    garbageN = 0
    nulPos = InStr(txt, Chr$(0))
    If nulPos > 0 Then
        garbageN = Len(txt) - nulPos + 1
        txt = Left$(txt, nulPos - 1)
    End If
    o("sinkTrailingNulBytes") = garbageN
    o("outputBytes") = Len(txt)
    o("outputTruncated") = trunc
    o("output") = ImJoinLines(Owner, txt)
    o("outputTrimmed") = ImTrimLines(Owner, txt)
    o("outputLines") = ImCountLines(Owner, txt)
    o("sinkDeleted") = ImDeleteSink(Owner, sink)
    If mode = "eval" Then
        Dim joined As String
        joined = ImJoinLines(Owner, txt)
        o("typeName") = ""
        o("typeNameUnavailableReason") = "Debug.Print transports the VALUE only: the immediate context runs in the OPEN project (no add-in types), so TypeName is unavailable on this route"
        o("valueTransport") = "debug_print"
        o("valueTransportNote") = "the expression is executed as Debug.Print <expr> through the SAME channel immediate_run uses; the add-in reads that sink. The exec line opens a sink file under an immediate-context scratch handle, so no add-in identifier appears in it."
        If Len(Trim$(joined)) = 0 Then
            If CStr(o("reason")) = "" Then
                o("reason") = "value_not_parsed"
                o("note") = "the statement completed (ret=0) but Debug.Print produced no output; output/outputTrimmed hold the RAW sink text - nothing is guessed"
            End If
        Else
            o("valueParsed") = True
            o("value") = ImTrimLines(Owner, txt)
            o("valueEscaped") = ImEscapeValue(Owner, CStr(o("value")))
            o("valueBytes") = Len(CStr(o("value")))
            o("valueBytesReturned") = Len(CStr(o("value")))
            o("valueLines") = ImValueLineCount(Owner, CStr(o("value")))
            ' FIX-25 (D2): valueComplete is a MEASUREMENT now: complete only when nothing was cut
            ' off the tail and the sink was not truncated. It used to be the constant True, which
            ' made a value followed by NUL garbage read as complete.
            o("valueComplete") = (garbageN = 0 And Not trunc)
            If Not CBool(o("valueComplete")) Then
                If garbageN > 0 Then
                    o("valueCompleteReason") = "trailing_nul_bytes"
                    o("valueCompleteNote") = "the sink carried " & CStr(garbageN) & " NUL character(s) after the value; they are a byte-coordinate artifact of the sink FILE, not data, and they were CUT at the first NUL. The value shown is the part before it."
                ElseIf trunc Then
                    o("valueCompleteReason") = "sink_truncated"
                    o("valueCompleteNote") = "the sink read stopped at IM_SINK_MAX_BYTES (" & CStr(IM_SINK_MAX_BYTES) & " bytes), so the value is a PREFIX"
                End If
            End If
            o("valueChars") = Len(CStr(o("value")))
            o("valueAnsiBytes") = ImAnsiByteCount(Owner, CStr(o("value")))
            o("valueEncoding") = "raw"
            If CLng(o("valueLines")) > 1 Then o("valueEncoding") = "json"
            o("valueSinkTruncated") = trunc
            If trunc Then
                o("valueTruncated") = True
                o("valueTruncatedReason") = "the sink read stopped at IM_SINK_MAX_BYTES (" & CStr(IM_SINK_MAX_BYTES) & " bytes)"
            End If
            If CStr(o("reason")) = "" Then o("ok") = True
            o("note") = "the expression was executed as Debug.Print in the ACTIVE project and its printed output was carried back through the Debug.Print sink (valueTransport=debug_print)."
        End If
    End If
    If CStr(o("reason")) = "" Then
        If mode <> "eval" Then
            o("ok") = True
            o("note") = "the code ran in the ACTIVE project's immediate context with each statement-boundary Debug.Print redirected into a temp file the tool reads back; only what THIS call printed is returned."
        End If
    End If
    Set ImExecuteImmediate = o
End Function

Public Function ImFlatLine(Owner As CodeTools, ByVal code As String) As String
    Dim s As String
    s = Replace(code, vbCrLf, ": ")
    s = Replace(s, vbLf, ": ")
    s = Replace(s, vbCr, ": ")
    ImFlatLine = Trim$(s)
End Function

Public Function ImJoinLines(Owner As CodeTools, ByVal s As String) As String
    Dim t As String
    t = Replace(s, vbCrLf, vbLf)
    t = Replace(t, vbCr, vbLf)
    ImJoinLines = t
End Function

Public Function ImObjectExprProblem(Owner As CodeTools, ByVal expr As String) As String
    ImObjectExprProblem = ""
    Dim t As String
    t = LCase$(expr)
    If Len(t) = 0 Then Exit Function
    If InStr(1, t, "createobject(", vbTextCompare) > 0 Then
        ImObjectExprProblem = "CreateObject("
    ElseIf InStr(1, t, "getobject(", vbTextCompare) > 0 Then
        ImObjectExprProblem = "GetObject("
    ElseIf InStr(1, t, "array(", vbTextCompare) > 0 Then
        ImObjectExprProblem = "Array("
    ElseIf InStr(1, t, " new ", vbTextCompare) > 0 Then
        ImObjectExprProblem = "New"
    ElseIf Left$(t, 4) = "new " Then
        ImObjectExprProblem = "New"
    End If
End Function

Public Function ImReadSink(Owner As CodeTools, ByVal p As String, ByRef wasTruncated As Boolean) As String
    Dim h As Integer
    Dim n As Long
    Dim s As String
    ImReadSink = ""
    wasTruncated = False
    On Error Resume Next
    If Len(p) = 0 Then Exit Function
    If Len(Dir$(p)) = 0 Then Exit Function
    n = FileLen(p)
    If n <= 0 Then Exit Function
    If n > IM_SINK_MAX_BYTES Then
        wasTruncated = True
        n = IM_SINK_MAX_BYTES
    End If
    h = FreeFile
    Open p For Binary As #h
    If Err.Number <> 0 Then
        Err.Clear
        Exit Function
    End If
    ' FIX-25 (D2): the sink is a BYTE file. The old code sized a STRING buffer with the BYTE
    ' count (s = Space$(n), n = FileLen) and then Get-copied the bytes into it, so a GBK value
    ' (2 bytes per char) came back as the value PLUS NUL characters while the receipt still said
    ' valueComplete:true (measured: value="中文值" + 3 NULs, valueBytes:7, valueLines:2).
    ' Read the BYTES as bytes and decode them as ANSI - the encoding the Debug.Print channel writes.
    Dim b() As Byte
    ReDim b(0 To n - 1)
    Get #h, , b
    Close #h
    If Err.Number <> 0 Then
        Err.Clear
        Exit Function
    End If
    Encoding.DecodeBytes b, "ansi", s
    ImReadSink = s
    Err.Clear
End Function

Public Function ImRewriteDebugPrint(Owner As CodeTools, ByVal code As String, ByRef rewrites As Long, ByVal fileToken As String) As String
    rewrites = 0
    ImRewriteDebugPrint = code
    If Len(code) = 0 Then Exit Function
    Dim rx As Object
    Dim col As Object
    On Error Resume Next
    Set rx = CreateObject("VBScript.RegExp")
    If Err.Number <> 0 Or rx Is Nothing Then
        Err.Clear
        Exit Function
    End If
    rx.Global = True
    rx.IgnoreCase = True
    rx.Pattern = "(^|:)\s*Debug\s*\.\s*Print\b"
    Set col = rx.Execute(code)
    If Err.Number = 0 Then
        If Not col Is Nothing Then rewrites = col.Count
    End If
    Err.Clear
    ImRewriteDebugPrint = rx.Replace(code, "$1 Print " & fileToken & ",")
    If Err.Number <> 0 Then
        ImRewriteDebugPrint = code
        rewrites = 0
        Err.Clear
    End If
    On Error GoTo 0
End Function

Public Function ImSinkDir(Owner As CodeTools) As String
    Dim p As String
    p = Util.TempDir() & "\vb6mcp_im"
    On Error Resume Next
    If Len(Dir$(p, vbDirectory)) = 0 Then MkDir p
    On Error GoTo 0
    ImSinkDir = p
End Function

Public Function ImSinkErrorCode(Owner As CodeTools, ByVal ret As Long, ByVal vbaErr As Long) As Boolean
    ImSinkErrorCode = False
    Dim codes As Variant
    Dim k As Long
    codes = Array(52, 53, 54, 55, 59, 61, 62, 63, 64, 67, 68, 70, 71, 75, 76)
    For k = LBound(codes) To UBound(codes)
        If ret = CLng(codes(k)) Then ImSinkErrorCode = True
        If vbaErr = CLng(codes(k)) Then ImSinkErrorCode = True
    Next k
End Function

Public Function ImSinkPath(Owner As CodeTools, ByVal tag As String) As String
    ImSinkPath = ImSinkDir(Owner) & "\" & tag & "_" & CStr(Owner.m_callSeq) & "_" & CStr(WoGetTickCount()) & ".tmp"
End Function

Public Function ImTrimLines(Owner As CodeTools, ByVal s As String) As String
    Dim parts() As String
    Dim i As Long
    Dim out As String
    out = ""
    parts = Split(ImJoinLines(Owner, s), vbLf)
    For i = LBound(parts) To UBound(parts)
        If Len(Trim$(parts(i))) > 0 Then
            If Len(out) > 0 Then out = out & vbLf
            out = out & Trim$(parts(i))
        End If
    Next i
    ImTrimLines = out
End Function

Public Function ImValueLineCount(Owner As CodeTools, ByVal s As String) As Long
    ' Lines the COMPLETE value occupies when split on CR/LF (a trailing break adds one, the
    ' same convention ImCountLines uses for the output sink).
    Dim t2 As String
    ImValueLineCount = 1
    If Len(s) = 0 Then Exit Function
    t2 = Replace(s, vbCrLf, vbLf)
    t2 = Replace(t2, vbCr, vbLf)
    ImValueLineCount = UBound(Split(t2, vbLf)) + 1
End Function

Public Function ImWouldRaiseRisk(Owner As CodeTools, ByVal expr As String, ByVal timeoutMs As Long) As Object
    Set ImWouldRaiseRisk = Nothing
    On Error Resume Next
    Dim t As String
    t = Trim$(expr)
    If Len(t) < 3 Then Exit Function
    Dim n As Long
    Dim ch As String
    Dim root As String
    root = ""
    For n = 1 To Len(t)
        ch = Mid$(t, n, 1)
        If (ch >= "A" And ch <= "Z") Or (ch >= "a" And ch <= "z") Or (ch >= "0" And ch <= "9") Or ch = "_" Then
            root = root & ch
        Else
            Exit For
        End If
    Next n
    If Len(root) = 0 Then Exit Function
    ch = Mid$(t, 1, 1)
    If Not ((ch >= "A" And ch <= "Z") Or (ch >= "a" And ch <= "z") Or ch = "_") Then Exit Function
    ' member access only: the very next character is a dot and something follows it
    If Mid$(t, Len(root) + 1, 1) <> "." Then Exit Function
    If Len(t) <= Len(root) + 1 Then Exit Function
    Dim probeMs As Long
    probeMs = timeoutMs
    If probeMs <= 0 Then probeMs = 5000
    If probeMs > 8000 Then probeMs = 8000
    Dim probeExpr As String
    probeExpr = "TypeName(" & root & ")"
    Dim pr As Object
    Set pr = ImExecuteImmediate(Owner, "eval", probeExpr, probeMs)
    If pr Is Nothing Then Exit Function
    If pr.Exists("ok") Then
        If Not CBool(pr("ok")) Then Exit Function
    End If
    Dim tn As String
    tn = ""
    If pr.Exists("outputTrimmed") Then tn = Trim$(CStr(pr("outputTrimmed")))
    If Len(tn) = 0 Then
        If pr.Exists("value") Then tn = Trim$(CStr(pr("value")))
    End If
    If Len(tn) = 0 Then Exit Function
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("root") = root
    d("probeExpr") = probeExpr
    d("typeName") = tn
    d("probeMs") = probeMs
    d("memberAccess") = Mid$(t, Len(root) + 2)
    If StrComp(tn, "Empty", vbTextCompare) = 0 Then
        d("errNumber") = 424
        d("description") = "Object required"
        Set ImWouldRaiseRisk = d
    ElseIf StrComp(tn, "Nothing", vbTextCompare) = 0 Then
        d("errNumber") = 91
        d("description") = "Object variable or With block variable not set"
        Set ImWouldRaiseRisk = d
    End If
    Err.Clear
End Function

