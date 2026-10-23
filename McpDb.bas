Attribute VB_Name = "McpDb"
Option Explicit

' =======================================================================
' C-3 facade batch: McpDb (moved bodies). The CLASS keeps same-signature Private
' forwarders, so no class call site changes; each facade delegates to this module.
' Owner convention: Owner As CodeTools is the FIRST parameter of every procedure here.
' =======================================================================

Public Sub DbAddTranscript(Owner As CodeTools, ByRef tr As Object, ByVal kind As String, ByVal posJson As String, ByVal stackJson As String)
    Dim e As Object
    Dim sj As String
    sj = ""
    On Error Resume Next
    sj = Vba6Probe.CallstackJson(32)
    If Err.Number <> 0 Then
        sj = ""
        Err.Clear
    End If
    On Error GoTo 0
    Set e = CreateObject("Scripting.Dictionary")
    e("seq") = tr.Count + 1
    e("kind") = kind
    e("position") = posJson
    e("callstack") = sj
    If Len(sj) = 0 Then e("callstackUnavailable") = True
    Set e("locals") = Owner.EmptyJsonArray()
    e("localsUnavailable") = True
    e("output") = ""
    e("err") = ""
    ' W-H2: watch expressions evaluated at THIS stop, same per-value shape as get_locals.
    If Not Owner.m_watchExprs Is Nothing Then
        If Owner.m_watchExprs.Count > 0 Then Set e("watch") = DbWatchArray(Owner)
    End If
    tr.Add e
End Sub

Public Function DbArmBreakpoints(Owner As CodeTools, ByVal args As Object) As Object
    Dim out As Object
    Set out = New Collection
    ' FIX-7: validate the WHOLE batch BEFORE arming anything: a half-armed batch is worse
    ' than a refusal, and a refused position can leave a host modal over the IDE.
    Owner.m_bpArmRefusal = ""
    If Not Owner.BpArmValidateBatch(args, Owner.m_bpArmRefusal) Then
        Set DbArmBreakpoints = out
        Exit Function
    End If
    Dim bps As Object
    Set bps = DbObjOf(Owner, args, "breakpoints")
    If bps Is Nothing Then
        Set DbArmBreakpoints = out
        Exit Function
    End If
    Dim i As Long
    For i = 1 To bps.Count
        Dim one As Object
        Set one = Nothing
        If IsObject(bps(i)) Then Set one = bps(i)
        Dim comp As String, lineN As Long
        comp = ""
        lineN = 0
        If Not one Is Nothing Then
            If one.Exists("component") Then comp = CStr(one("component"))
            If one.Exists("line") Then lineN = CLng(one("line"))
        End If
        Dim rec As Object
        Set rec = CreateObject("Scripting.Dictionary")
        rec("component") = comp
        rec("paneLine") = lineN
        rec("diskLine") = lineN
        rec("diskLineExact") = False
        rec("mappingResolved") = False
        On Error Resume Next
        Dim cdA As Object
        Set cdA = Owner.CoordsDictOfName(comp, lineN, "ide")
        If Err.Number = 0 Then
            If Not cdA Is Nothing Then
                If cdA.Exists("diskLine") Then rec("diskLine") = cdA("diskLine")
                If cdA.Exists("diskLineExact") Then rec("diskLineExact") = cdA("diskLineExact")
                If cdA.Exists("mappingResolved") Then rec("mappingResolved") = cdA("mappingResolved")
            End If
        End If
        Err.Clear
        On Error GoTo 0
        rec("armed") = False
        rec("hitCount") = 0
        ' W-H1: stop only on the Nth hit (the engine has no conditional breakpoints, so this is
        ' the equivalent capability: the driver re-runs until the target hit is reached).
        rec("hitCountTarget") = 1
        If Not one Is Nothing Then
            If one.Exists("hitCount") Then rec("hitCountTarget") = CLng(one("hitCount"))
        End If
        If CLng(rec("hitCountTarget")) < 1 Then rec("hitCountTarget") = 1
        If lineN > 0 Then
            Dim bpa As Object
            Set bpa = CreateObject("Scripting.Dictionary")
            bpa("action") = "toggle"
            bpa("component") = comp
            bpa("line") = lineN
            Dim rj As String
            rj = ""
            On Error Resume Next
            rj = Owner.ToolNativeBreakpoint(bpa)
            If Err.Number <> 0 Then
                rj = ""
                Err.Clear
            End If
            On Error GoTo 0
            rec("toggleReceipt") = Left$(rj, 400)
            ' armed = the toggle receipt reports the breakpoint ON (or at least no refusal)
            If InStr(1, rj, """breakpointOn"":true", vbTextCompare) > 0 Then
                rec("armed") = True
            ElseIf Len(rj) > 0 And InStr(1, rj, """reason""", vbTextCompare) = 0 Then
                rec("armed") = True
            End If
        End If
        out.Add rec
    Next
    Set DbArmBreakpoints = out
End Function

Public Function DbBreakGate(Owner As CodeTools, ByRef root As Object, ByRef tr As Object) As String
    DbBreakGate = "stop"
    On Error Resume Next
    Dim posJ As String
    posJ = DbPositionOrFallback(Owner, root)
    root("state") = "break"
    root("position") = posJ
    root("positionSource") = "paneSelection"
    root("proc") = DbPosField(Owner, posJ, "proc")
    Dim sj As String
    sj = Vba6Probe.CallstackJson(32)
    root("callstack") = sj
    If Len(sj) = 0 Then root("callstackUnavailable") = True
    Set root("locals") = Owner.EmptyJsonArray()
    root("localsUnavailable") = True
    root("hitCountTotal") = CLng(root("hitCountTotal")) + 1
    Dim modName As String
    modName = DbPosField(Owner, posJ, "module")
    Dim lineN As Long
    lineN = DbPosLine(Owner, posJ)
    Dim bps As Object
    Set bps = Nothing
    If root.Exists("breakpoints") Then
        If IsObject(root("breakpoints")) Then Set bps = root("breakpoints")
    End If
    If Not bps Is Nothing Then
        Dim i As Long
        Dim rec As Object
        Dim tgt As Long
        For i = 1 To bps.Count
            Set rec = Nothing
            If IsObject(bps(i)) Then Set rec = bps(i)
            If Not rec Is Nothing Then
                If StrComp(CStr(rec("component")), modName, vbTextCompare) = 0 Then
                    If CLng(rec("paneLine")) = lineN Or CLng(rec("paneLine")) = lineN - 1 Then
                        rec("hitCount") = CLng(rec("hitCount")) + 1
                        tgt = 1
                        If rec.Exists("hitCountTarget") Then tgt = CLng(rec("hitCountTarget"))
                        If tgt < 1 Then tgt = 1
                        If CLng(rec("hitCount")) < tgt Then
                            root("autoContinues") = CLng(root("autoContinues")) + 1
                            root("hitCountGateNote") = "the engine has NO conditional breakpoints; the driver auto-continues each hit until the requested hitCount is reached (expression conditions are not supported)"
                            DbBreakGate = "continue"
                            Exit Function
                        End If
                        Set root("lastBreakpointHit") = rec
                    End If
                End If
            End If
        Next
    End If
    DbAddTranscript Owner, tr, "break", posJ, ""
End Function

Public Function DbBreakNow(Owner As CodeTools) As Boolean
    DbBreakNow = False
    On Error Resume Next
    If IdeEbMode() = 2 Then DbBreakNow = True
    Err.Clear
    On Error GoTo 0
    If DbBreakNow Then Exit Function
    If Owner.GetIdeState() = "break" Then DbBreakNow = True
End Function

Public Sub DbBriefWait(Owner As CodeTools, ByVal ms As Long)
    Dim w As Long
    w = Util.TickMs()
    Do While (Util.TickMs() - w) < ms
    Loop
End Sub

Public Function DbClickEnd(Owner As CodeTools) As Boolean
    DbClickEnd = False
    On Error Resume Next
    Dim ctrl As Object
    Dim why As String
    why = ""
    Set ctrl = McpFindTools.FindCmd(Owner, "end", "Debug|Run|" & ChrW$(&H8C03) & ChrW$(&H8BD5) & "|" & ChrW$(&H8FD0) & ChrW$(&H884C), "End|" & ChrW$(&H7ED3) & ChrW$(&H675F))
    If ctrl Is Nothing Then Set ctrl = Owner.m_cachedEnd Else Set Owner.m_cachedEnd = ctrl
    If ctrl Is Nothing Then Exit Function
    DbClickEnd = (Owner.ExecIfEnabledWhy(ctrl, "end", why) > 0)
    Err.Clear
End Function

Public Function DbClickStart(Owner As CodeTools, ByRef whyOut As String) As Boolean
    whyOut = ""
    DbClickStart = False
    Dim barCands As String, cands As String
    barCands = "Debug|Run|" & ChrW$(&H8C03) & ChrW$(&H8BD5) & "|" & ChrW$(&H8FD0) & ChrW$(&H884C)
    If Owner.GetIdeState() = "break" Then
        cands = "Continue|" & ChrW$(&H7EE7) & ChrW$(&H7EED)
    Else
        cands = "Start|" & ChrW$(&H542F) & ChrW$(&H52A8) & "|Continue|" & ChrW$(&H7EE7) & ChrW$(&H7EED)
    End If

    ' R-1021-05: a Start here begins a NEW run, so the previous runtime-error capture is
    ' STALE. The Continue branch above is NOT a new run and deliberately keeps it.
    If Owner.GetIdeState() <> "break" Then McpRe.ReInvalidateRuntimeError Owner, "new_run_started"
    Dim ctrl As Object
    Set ctrl = McpFindTools.FindCmd(Owner, "start", barCands, cands)
    If ctrl Is Nothing Then Set ctrl = Owner.m_cachedStart
    If ctrl Is Nothing Then Set ctrl = McpFindTools.FindBarControlDeep(Owner, cands)
    If ctrl Is Nothing Then
        whyOut = "start menu command not found (searched " & cands & ")"
        Exit Function
    End If
    Set Owner.m_cachedStart = ctrl
    Dim n As Long
    n = Owner.ExecIfEnabledWhy(ctrl, "start", whyOut)
    ' (D3) the post-run suspended CommandBars window leaves the item DISABLED although nothing
    ' is running (measured). Clear it ONCE, re-resolve and retry ONCE instead of reporting a
    ' start failure that is really an IDE-state artefact.
    If n = 0 And InStr(1, whyOut, "disabled", vbTextCompare) > 0 Then
        Dim nr As Object
        Set nr = DbNudgeBarsOut(Owner)
        If CBool(nr("attempted")) Then
            Set ctrl = Nothing
            Set ctrl = McpFindTools.FindCmd(Owner, "start", barCands, cands)
            If ctrl Is Nothing Then Set ctrl = McpFindTools.FindBarControlDeep(Owner, cands)
            If Not ctrl Is Nothing Then
                whyOut = ""
                n = Owner.ExecIfEnabledWhy(ctrl, "start", whyOut)
                Set Owner.m_cachedStart = ctrl
            End If
        End If
    End If
    DbClickStart = (n > 0)
End Function

Public Function DbErrorDialogBlock(Owner As CodeTools, ByVal title As String) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("present") = True
    d("title") = title
    On Error Resume Next
    Dim body As String
    body = ""
    body = WinOps.DialogBodiesText("#32770")
    If Len(body) > 600 Then body = Left$(body, 600)
    d("body") = body
    d("howToFinish") = "the runtime-error box was consumed by the capture path: the add-in READ it and clicked DEBUG (NEVER End). This block is the dialog identity only; a still-visible box is reported by runtimeErrorDismiss{reason,buttons,measuredEbMode}. END IS NEVER CLICKED on this path"
    Set DbErrorDialogBlock = d
End Function

Public Function DbFileExists(Owner As CodeTools, ByVal p As String) As Boolean
    On Error Resume Next
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    DbFileExists = fso.FileExists(p)
    If Err.Number <> 0 Then
        DbFileExists = False
        Err.Clear
    End If
    On Error GoTo 0
End Function

Public Sub DbFinalize(Owner As CodeTools, ByRef root As Object, ByRef tr As Object, ByVal vbpPath As String, ByVal diskBefore As String, ByVal leaveRunning As Boolean)
    On Error Resume Next
    Dim stNow As String
    stNow = Owner.GetIdeState()
    root("ideStateAfter") = stNow
    If Not leaveRunning Then
        ' request the end; the command is QUEUED (STA) and is executed only after this call
        ' returns, so the wait below is a bounded confirmation - never proof of failure.
        root("endRequested") = True
        root("endClicked") = DbClickEnd(Owner)
        Dim tEnd As Long
        tEnd = Util.TickMs()
        Do While (Util.TickMs() - tEnd) < 3000
            If Util.CanPump() Then DoEvents
            stNow = Owner.GetIdeState()
            If stNow = "design" Then Exit Do
        Loop
        root("ideStateAfter") = stNow
        If stNow = "break" Or stNow = "run" Then
            root("endQueued") = True
            Set root("howToFinish") = DbHowToFinish(Owner)
            If IdeEbMode() = 2 And (Not Owner.IdeBarsUsable()) Then
                Set root("autoRecover") = DbNudgeBarsOut(Owner)
                If IdeEbMode() = 2 And (Not Owner.IdeBarsUsable()) Then
                    root("wedged") = True
                    root("wedgedNote") = "an EbMode=2 break context is held while the IDE CommandBars model is gone, and one activate_component recovery attempt did not bring it back"
                End If
            End If
        End If
    Else
        If stNow = "break" Or stNow = "run" Then
            root("stillRunning") = True
            Set root("howToFinish") = DbHowToFinish(Owner)
        End If
    End If
    DbFinalizeDiskAndHits Owner, root, vbpPath, diskBefore, tr
    Err.Clear
End Sub

Public Sub DbFinalizeDiskAndHits(Owner As CodeTools, ByRef root As Object, ByVal vbpPath As String, ByVal diskBefore As String, ByRef tr As Object)
    On Error Resume Next
    Dim diskAfter As String
    diskAfter = ""
    If Len(vbpPath) > 0 Then
        If DbFileExists(Owner, vbpPath) Then diskAfter = Sha256HexFile(vbpPath)
    End If
    root("diskVbpShaAfter") = diskAfter
    root("vbpShaBefore") = diskBefore
    root("vbpShaAfter") = diskAfter
    root("savedDuringCall") = (Len(diskBefore) > 0 And StrComp(diskBefore, diskAfter, vbTextCompare) <> 0)
    root("reorderedDuringCall") = False
    root("vbpRewrittenByIde") = (Len(diskBefore) > 0 And StrComp(diskBefore, diskAfter, vbTextCompare) <> 0)
    Set root("transcript") = tr
    root("transcriptCount") = tr.Count
    Dim sr As String
    sr = CStr(root("stopReason"))
    If Len(sr) = 0 Then
        sr = "error"
        root("stopReason") = "error"
    End If
    If sr = "initial_break" Or sr = "process_exit" Or sr = "actions_done" Or sr = "started" Or sr = "stopped" Or sr = "break" Or sr = "unhandled_error" Or sr = "runtime_error" Or sr = "runtime_error_run_ended" Then
        root("ok") = True
    Else
        root("ok") = False
    End If
    ' R-1021-02 (HINT ONLY): debug_session and run_until_break ALSO start the
    ' program, so the same IDE option can save the in-memory edits on that Start.
    ' No parameter and no refusal was added on this path - the receipt only names
    ' the mode and the hint, so a caller can not read saveOnRun:false as "no write".
    Dim rdkHint As Object
    Set rdkHint = Owner.RunDiskAutoSaveBlock()
    Owner.RunDiskMergePrim root, rdkHint
    root("saveOnRun") = False
    root("saveOnRunReason") = ""
    root("runDiskGateApplied") = False
    root("runDiskGateNote") = "HINT ONLY on this tool: debug_session / run_until_break start the program, so the IDE option may save on Start, but this path has no allowDiskMutation parameter and no dirty-memory refusal. Re-check the disk yourself (verify / get_component_code) when it matters."
    Err.Clear
End Sub

Public Function DbHowToFinish(Owner As CodeTools) As Object
    Dim hf As Object
    Set hf = CreateObject("Scripting.Dictionary")
    hf("tool") = "control_execution"
    Dim hfa As Object
    Set hfa = CreateObject("Scripting.Dictionary")
    hfa("mode") = "end"
    Set hf("arguments") = hfa
    Set DbHowToFinish = hf
End Function

Public Sub DbNoteHit(Owner As CodeTools, ByRef root As Object, ByVal bps As Object)
    On Error Resume Next
    root("hitCountTotal") = CLng(root("hitCountTotal")) + 1
    If bps Is Nothing Then Exit Sub
    Dim i As Long, rec As Object
    For i = 1 To bps.Count
        Set rec = bps(i)
        If Not rec Is Nothing Then
            If CBool(rec("armed")) Then
                rec("hitCount") = CLng(rec("hitCount")) + 1
                rec("hitCountNote") = "attributed to the first armed breakpoint (no native per-breakpoint hit API)"
                Exit For
            End If
        End If
    Next
    Err.Clear
End Sub

Public Function DbNudgeBarsOut(Owner As CodeTools) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("attempted") = False
    d("recovered") = False
    d("component") = ""
    d("via") = "activate_component"
    On Error Resume Next
    Dim nm As String
    nm = DbNudgeComponentName(Owner)
    d("component") = nm
    If Len(nm) = 0 Then
        d("note") = "no component to activate"
    Else
        d("attempted") = True
        Dim rj As String
        rj = ""
        rj = Owner.ToolActivateComponent(nm, "code")
        If Err.Number <> 0 Then
            d("error") = Err.Description
            Err.Clear
        End If
        d("thisCallAnswered") = (Len(rj) > 0)
        d("recovered") = Owner.IdeBarsUsable()
    End If
    Err.Clear
    On Error GoTo 0
    Set DbNudgeBarsOut = d
End Function

Public Function DbNudgeComponentName(Owner As CodeTools) As String
    On Error Resume Next
    DbNudgeComponentName = ""
    Dim c As VBIDE.VBComponent
    Set c = Owner.m_vbe.ActiveVBComponent
    If Err.Number = 0 Then
        If Not c Is Nothing Then
            DbNudgeComponentName = c.Name
            Exit Function
        End If
    End If
    Err.Clear
    Dim p As VBIDE.VBProject
    Set p = Owner.GetProject()
    If p Is Nothing Then Exit Function
    If p.VBComponents.Count = 0 Then Exit Function
    Set c = p.VBComponents.Item(1)
    If Err.Number = 0 Then DbNudgeComponentName = c.Name
    Err.Clear
End Function

Public Function DbObjOf(Owner As CodeTools, ByVal args As Object, ByVal key As String) As Object
    Set DbObjOf = Nothing
    If args Is Nothing Then Exit Function
    On Error Resume Next
    If Not args.Exists(key) Then Exit Function
    If Not IsObject(args(key)) Then Exit Function
    Set DbObjOf = args(key)
    Err.Clear
End Function

Public Function DbObserveUntilBreak(Owner As CodeTools, ByRef root As Object, ByRef tr As Object, ByVal timeoutMs As Long) As String
    On Error Resume Next
    DbObserveUntilBreak = "timeout"
    Dim t0 As Long
    Dim st As String
    Dim dlgTitle As String
    t0 = Util.TickMs()
    root("state") = "run"
    Do While (Util.TickMs() - t0) < timeoutMs
        ' R-1021-05: probe the MODAL runtime-error box FIRST - before any CommandBars or
        ' IDE-object-model call in this iteration. With the box up those calls QUEUE
        ' (measured: an End issued then took ~12 s and reported state_not_changed), and the
        ' reverse-engineered engine entry (EbSetMode) crashes the IDE. The probe itself is
        ' Win32-only (this process's own windows); on a hit the rule runs here and returns.
        Dim reRes As Object
        Set reRes = McpRe.ReCaptureRuntimeError(Owner, True)
        If CBool(reRes("found")) Then
            root("err") = CStr(reRes("dialogTitle"))
            Set root("errorDialog") = DbErrorDialogBlock(Owner, CStr(reRes("dialogTitle")))
            DbObserveUntilBreak = McpRe.ReMergeRuntimeError(Owner, root, reRes)
            Exit Function
        End If
        If DbBreakNow(Owner) Then
            Dim gate As String
            gate = DbBreakGate(Owner, root, tr)
            If gate = "continue" Then
                ' auto-continue towards the requested hit count
                root("state") = "run"
                Dim g1 As Long
                g1 = Util.TickMs()
                Owner.ToolControlExecution "continue"
                Do While (Util.TickMs() - g1) < 8000
                    If DbBreakNow(Owner) Then Exit Do
                    If Owner.GetIdeState() = "design" Then Exit Do
                    If Util.CanPump() Then DoEvents
                    DbBriefWait Owner, 40
                Loop
            Else
                DbObserveUntilBreak = "break"
                Exit Function
            End If
        Else
            st = Owner.GetIdeState()
            If st = "design" Then
                root("state") = "design"
                McpRe.ReInvalidateRuntimeError Owner, "run_ended_normally"
                DbObserveUntilBreak = "process_exit"
                Exit Function
            End If
            dlgTitle = ""
            If WinOps.HasVisibleHostDialog(dlgTitle) Then
                root("state") = "break"
                root("err") = dlgTitle
                Set root("errorDialog") = DbErrorDialogBlock(Owner, dlgTitle)
                DbObserveUntilBreak = "unhandled_error"
                Exit Function
            End If
        End If
        If Util.CanPump() Then DoEvents
        DbBriefWait Owner, 40
    Loop
    ' R-1021-05: never consult the IDE object model while the modal box is (still) up.
    If McpRe.ReFindRuntimeErrorDialog(Owner) Is Nothing Then
        root("state") = Owner.GetIdeState()
    Else
        root("state") = "runtime_error"
        root("stateSource") = "runtime_error_dialog (timeout reached with the box up: the IDE object model was NOT consulted)"
    End If
End Function

Public Function DbPosField(Owner As CodeTools, ByVal posJson As String, ByVal key As String) As String
    On Error Resume Next
    DbPosField = ""
    Dim v As Variant
    If Not JsonParse(posJson, v) Then Exit Function
    Dim o As Object
    Set o = v
    If o Is Nothing Then Exit Function
    DbPosField = JsonGetStr(o, key, "")
    Err.Clear
End Function

Public Function DbPosLine(Owner As CodeTools, ByVal posJson As String) As Long
    On Error Resume Next
    DbPosLine = 0
    Dim v As Variant
    If Not JsonParse(posJson, v) Then Exit Function
    Dim o As Object
    Set o = v
    If o Is Nothing Then Exit Function
    DbPosLine = JsonGetLong(o, "line", 0)
    Err.Clear
End Function

Public Function DbPositionOrFallback(Owner As CodeTools, ByRef root As Object) As String
    DbPositionOrFallback = ""
    On Error Resume Next
    DbPositionOrFallback = Owner.ToolCurrentPosition()
    If Err.Number <> 0 Then
        DbPositionOrFallback = ""
        Err.Clear
    End If
    On Error GoTo 0
    If Len(DbPositionOrFallback) = 0 Then
        root("positionUnavailable") = True
        DbPositionOrFallback = "{""positionSource"":""breakpoint""}"
    End If
End Function

Public Function DbRubHowToContinue(Owner As CodeTools, ByVal timeoutMs As Long) As Object
    Dim hc As Object
    Set hc = CreateObject("Scripting.Dictionary")
    hc("tool") = "run_until_break"
    Dim ha As Object
    Set ha = CreateObject("Scripting.Dictionary")
    ha("attach") = True
    ha("timeout_ms") = timeoutMs
    Set hc("arguments") = ha
    Set DbRubHowToContinue = hc
End Function

Public Sub DbRunActions(Owner As CodeTools, ByVal args As Object, ByRef tr As Object, ByVal timeoutMs As Long, ByRef root As Object)
    Dim acts As Object
    Set acts = DbObjOf(Owner, args, "actions")
    If acts Is Nothing Then
        root("stopReason") = "actions_done"
        Exit Sub
    End If
    Dim i As Long
    For i = 1 To acts.Count
        Dim a As Object
        Set a = Nothing
        If IsObject(acts(i)) Then Set a = acts(i)
        Dim op As String
        op = ""
        If Not a Is Nothing Then
            If a.Exists("op") Then op = LCase$(CStr(a("op")))
        End If
        If op = "stop" Then
            root("stopReason") = "stopped"
            Exit Sub
        ElseIf op = "eval" Then
            Dim expr As String
            expr = ""
            If a.Exists("expr") Then expr = CStr(a("expr"))
            Dim ev As String
            ev = ""
            On Error Resume Next
            ev = Owner.ToolImmediateEval(expr, "", 5000)
            If Err.Number <> 0 Then
                ev = ""
                Err.Clear
            End If
            On Error GoTo 0
            Dim e2 As Object
            Set e2 = CreateObject("Scripting.Dictionary")
            e2("seq") = tr.Count + 1
            e2("kind") = "eval"
            e2("position") = Owner.ToolCurrentPosition()
            e2("expr") = expr
            e2("output") = Left$(ev, 2000)
            e2("callstack") = ""
            Set e2("locals") = Owner.EmptyJsonArray()
            e2("err") = ""
            tr.Add e2
        ElseIf op = "continue" Then
            On Error Resume Next
            Owner.ToolControlExecution "continue"
            If Err.Number <> 0 Then Err.Clear
            On Error GoTo 0
            Dim t1 As Long
            ' same asynchronous-start problem as after DbClickStart
            Dim tC As Long
            tC = Util.TickMs()
            Dim stC As String
            stC = Owner.GetIdeState()
            Do While (Util.TickMs() - tC) < 8000
                If Util.CanPump() Then DoEvents
                stC = Owner.GetIdeState()
                If Not DbBreakNow(Owner) Then Exit Do
            Loop
            t1 = Util.TickMs()
            Dim st As String
            st = Owner.GetIdeState()
            Do While (Util.TickMs() - t1) < timeoutMs
                If Util.CanPump() Then DoEvents
                st = Owner.GetIdeState()
                If DbBreakNow(Owner) Then Exit Do
                If st = "design" Then Exit Do
            Loop
            If st = "break" Then
                DbAddTranscript Owner, tr, "break", DbPositionOrFallback(Owner, root), ""
            ElseIf st = "design" Then
                root("stopReason") = "process_exit"
                Exit Sub
            Else
                root("stopReason") = "timeout"
                Exit Sub
            End If
        ElseIf op = "step_over" Or op = "step_into" Or op = "step_out" Then
            On Error Resume Next
            Owner.ToolControlExecution op
            If Err.Number <> 0 Then Err.Clear
            On Error GoTo 0
            Dim t2 As Long
            t2 = Util.TickMs()
            Dim st2 As String
            st2 = Owner.GetIdeState()
            Do While (Util.TickMs() - t2) < 10000
                If Util.CanPump() Then DoEvents
                st2 = Owner.GetIdeState()
                If DbBreakNow(Owner) Then Exit Do
                If st2 = "design" Then Exit Do
            Loop
            If st2 = "break" Then
                DbAddTranscript Owner, tr, "break", DbPositionOrFallback(Owner, root), ""
            ElseIf st2 = "design" Then
                root("stopReason") = "process_exit"
                Exit Sub
            End If
        End If
    Next
    If Len(CStr(root("stopReason"))) = 0 Then root("stopReason") = "actions_done"
End Sub

Public Function DbWatchArray(Owner As CodeTools) As Object
    Dim out As Object
    Set out = New Collection
    On Error Resume Next
    If Owner.m_watchExprs Is Nothing Then Set DbWatchArray = out : Exit Function
    If Owner.m_watchExprs.Count = 0 Then Set DbWatchArray = out : Exit Function
    Dim diag As Object
    Set diag = CreateObject("Scripting.Dictionary")
    Dim ctxErr3 As String
    ctxErr3 = McpWh2.Wh2FrameContext(Owner, Owner.m_watchFrame, diag)
    If Len(ctxErr3) > 0 Then Set DbWatchArray = out : Exit Function
    Dim i As Long
    For i = 1 To Owner.m_watchExprs.Count
        Dim ex As String
        ex = CStr(Owner.m_watchExprs(i))
        If McpWh2.Wh2NameAllowed(Owner, ex) Then
            Owner.ArmGateRelease "watch", 6000
            out.Add McpWh2.Wh2ValueShape(Owner, ex, Owner.m_watchFrame, diag)
            Owner.DisarmGateRelease
        Else
            Dim refused As Object
            Set refused = CreateObject("Scripting.Dictionary")
            refused("expr") = ex
            refused("error") = "unknown_identifier_refused"
            refused("note") = "transcript[].watch evaluates ONLY identifiers enumerated by get_locals in this IDE session; an unknown identifier is NOT evaluated because it can hang the in-flight call (measured)"
            out.Add refused
        End If
    Next
    Err.Clear
    Set DbWatchArray = out
End Function

