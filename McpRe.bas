Attribute VB_Name = "McpRe"
Option Explicit

' =======================================================================
' C-3 facade batch: McpRe (moved bodies). The CLASS keeps same-signature Private
' forwarders, so no class call site changes; each facade delegates to this module.
' Owner convention: Owner As CodeTools is the FIRST parameter of every procedure here.
' =======================================================================

Private Declare Function WoGetTickCount Lib "kernel32" Alias "GetTickCount" () As Long

Private Declare Function DbGetCurrentProcessId Lib "kernel32" Alias "GetCurrentProcessId" () As Long
Private Declare Function SetEnvironmentVariableW Lib "kernel32" (ByVal lpName As Long, ByVal lpValue As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Const ReBreakPollMs As Long = 100
Private Const ReBreakWaitMs As Long = 5000
Private Const ReEndWaitMs As Long = 5000

Public Function ReButtonsJson(Owner As CodeTools, ByVal dlg As Object) As Object
    Dim out As Object
    Set out = New Collection
    On Error Resume Next
    If dlg Is Nothing Then
        Set ReButtonsJson = out
        Exit Function
    End If
    Dim btns As Object
    Set btns = Nothing
    If dlg.Exists("buttons") Then Set btns = dlg("buttons")
    If btns Is Nothing Then
        Set ReButtonsJson = out
        Exit Function
    End If
    Dim i As Long
    Dim b As Object
    Dim e As Object
    For i = 1 To btns.Count
        Set b = btns(i)
        If Not b Is Nothing Then
            Set e = CreateObject("Scripting.Dictionary")
            e("hwnd") = JsonGetLong(b, "hwnd", 0)
            e("text") = CStr(b("text"))
            e("cls") = CStr(b("cls"))
            e("enabled") = CBool(b("enabled"))
            out.Add e
        End If
    Next
    Err.Clear
    Set ReButtonsJson = out
End Function

Public Function ReCachedPositionBlock(Owner As CodeTools, ByRef blk As Object) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("available") = True
    d("positionSource") = "cached_capture"
    d("module") = blk("module")
    d("proc") = blk("proc")
    d("line") = blk("line")
    d("paneLine") = blk("paneLine")
    d("diskLine") = blk("diskLine")
    d("capturedAt") = blk("capturedAt")
    d("text") = "(cached runtime-error capture: the pane selection is no longer the stop point)"
    Set ReCachedPositionBlock = d
End Function

Public Function ReCaptureRuntimeError(Owner As CodeTools, ByVal doClick As Boolean) As Object
    Dim res As Object
    Set res = ReEmptyResult(Owner)
    On Error Resume Next
    Dim dgT As Double
    Dim dlg As Object
    dgT = McpCls.DgNow(Owner)
    Set dlg = ReFindRuntimeErrorDialog(Owner)
    McpCls.DgSpan Owner, "ReFindRuntimeErrorDialog", dgT
    If dlg Is Nothing Then
        Err.Clear
        Set ReCaptureRuntimeError = res
        Exit Function
    End If
    res("found") = True
    res("dialogTitle") = CStr(dlg("title"))
    Set res("dialogInfo") = ReDialogInfoBlock(Owner, dlg)
    ' ---- (2) READ THE BODY BEFORE ANY CLICK -----------------------------
    Dim body As String
    dgT = McpCls.DgNow(Owner)
    body = ReRuntimeErrorText(Owner, dlg)
    McpCls.DgSpan Owner, "ReRuntimeErrorText", dgT
    Dim readAt As String
    readAt = Format$(Now, "yyyy-mm-dd hh:nn:ss")
    Dim numText As String
    Dim desc As String
    Dim src As String
    numText = ReParseErrorNumber(Owner, body)
    desc = ReParseErrorDescription(Owner, body)
    src = McpBpCore.H3ParseErrSource(Owner, body)
    ' (duplicate Dim dgT removed)
    Dim dbg As Object
    Dim endBtn As Object
    Dim okBtn As Object
    Dim clickTarget As Object
    dgT = McpCls.DgNow(Owner)
    Set dbg = ReIdFirstButton(Owner, dlg, "debug", False)
    If Not dbg Is Nothing Then
        res("dismissIdBasis") = "control_id"
        res("dismissControlId") = JsonGetLong(dbg, "id", 0)
    End If
    If dbg Is Nothing Then
        Set dbg = ReFindDebugButton(Owner, dlg)
        If Not dbg Is Nothing Then res("dismissIdBasis") = "caption_fallback"
        If Not dbg Is Nothing Then res("dismissControlId") = JsonGetLong(dbg, "id", 0)
    End If
    McpCls.DgSpan Owner, "FindDebugButton", dgT
    dgT = McpCls.DgNow(Owner)
    Set endBtn = ReIdFirstButton(Owner, dlg, "end", True)
    If endBtn Is Nothing Then Set endBtn = ReFindEndButtonInDialog(Owner, dlg)
    McpCls.DgSpan Owner, "ReFindEndButtonInDialog", dgT
    dgT = McpCls.DgNow(Owner)
    Set okBtn = ReFindOkButton(Owner, dlg)
    McpCls.DgSpan Owner, "ReFindOkButton", dgT
    Dim btns As Object
    Set btns = ReButtonsJson(Owner, dlg)
    Dim ebBefore As Long
    ebBefore = IdeEbMode()
    ' ---- (3) click DEBUG ONLY ------------------------------------------
    Dim clickJson As String
    Dim clicked As Boolean
    clickJson = ""
    clicked = False
    Dim viaClick As String
    viaClick = "debug"
    If doClick Then
        Set clickTarget = dbg
        If dbg Is Nothing Then
            ' EVAL SAFETY: OK fallback. An EVALUATOR box has no Debug caption, so refusing here
            ' left the user's IDE holding a box forever. OK is the correct dismissal for it.
            ' The END caption is still NEVER clicked.
            Set clickTarget = okBtn
            viaClick = "ok"
            If clickTarget Is Nothing Then
                ' STRUCTURED FAILURE: neither Debug nor OK. NOTHING is clicked, above all no End.
                Set res("dismiss") = ReDismissFailBlock(Owner, "no_debug_or_ok_button", dlg, btns, endBtn, ebBefore, "", True, False, ebBefore)
                Err.Clear
                Set ReCaptureRuntimeError = res
                Exit Function
            End If
        End If
        res("dismissedVia") = viaClick
        dgT = McpCls.DgNow(Owner)
        clickJson = WinOps.SendWindowMessageJson(JsonGetLong(clickTarget, "hwnd", 0), "click", "", "", "", "", 1, False, False)
        McpCls.DgSpan Owner, "DismissClickSend", dgT
        If viaClick = "debug" Then
            res("debugClickReceipt") = Left$(clickJson, 800)
        Else
            res("okClickReceipt") = Left$(clickJson, 800)
        End If
        ' The sender receipt is READ: a message that never reached the button must never be
        ' reported as a click. clicked stays False unless the sender itself says ok:true.
        Dim cVariant As Variant
        Dim cobj As Object
        Dim sentOk As Boolean
        sentOk = False
        If JsonParse(clickJson, cVariant) Then
            If IsObject(cVariant) Then
                Set cobj = cVariant
                If cobj.Exists("ok") Then sentOk = CBool(cobj("ok"))
            End If
        End If
        res("dismissClickSent") = sentOk
        If viaClick = "debug" Then
            res("debugClickSent") = sentOk
        Else
            res("okClickSent") = sentOk
        End If
        If Not sentOk Then
            ' STRUCTURED FAILURE: the BM_CLICK never reached the dismissal button. Nothing was
            ' measured and nothing else is sent - above all no End.
            res("outcome") = "blocked"
            Set res("postDebugClick") = RePostDebugBlock(Owner, IdeEbMode(), ReRawModeFromEb(Owner, IdeEbMode()), False, False, 0, False, True, 0)
            Set res("dismiss") = ReDismissFailBlock(Owner, IIf(viaClick = "debug", "debug_click_send_failed", "ok_click_send_failed"), dlg, btns, endBtn, IdeEbMode(), clickJson, True, False, ebBefore)
            Err.Clear
            Set ReCaptureRuntimeError = res
            Exit Function
        End If
        clicked = True
    End If
    ' ---- (4) MEASURE what the click actually DID - never assume break -----
    ' FIX-5: the proof is a BOUNDED POLL, never a single sample. The modal loop unwinds
    ' asynchronously (live, with the FORMER 3 s single-shot wait: EbMode still read 1; the
    ' next call read 2). During the wait only IdeEbMode() and a sleep are used.
    Dim waited As Long
    Dim ebNow As Long
    Dim pollAttempts As Long
    ebNow = IdeEbMode()
    waited = 0
    pollAttempts = 0
    If clicked Then
        dgT = McpCls.DgNow(Owner)
        waited = RePollBreakMs(Owner, True, ebNow, pollAttempts)
        McpCls.DgSpan Owner, "RePollBreakMs", dgT
    End If
    Dim boxGone As Boolean
    Dim d2 As Object
    boxGone = False
    If clicked Then
        dgT = McpCls.DgNow(Owner)
        Set d2 = ReFindRuntimeErrorDialog(Owner)
        McpCls.DgSpan Owner, "ReFindRuntimeErrorDialogAfterClick", dgT
        boxGone = (d2 Is Nothing)
    End If
    ' enteredBreak is TRUE only when the ENGINE is measured in break mode (EbMode = 2).
    res("enteredBreak") = False
    If ebNow = 2 Then res("enteredBreak") = True
    res("runEnded") = (ebNow = 0)
    Set res("postDebugClick") = RePostDebugBlock(Owner, ebNow, ReRawModeFromEb(Owner, ebNow), (ebNow = 2), (ebNow = 0), waited, clicked, boxGone, pollAttempts)
    If clicked Then
        If ebNow = 2 Then
            res("outcome") = "break"
        ElseIf ebNow = 0 Then
            res("outcome") = "run_ended"
        Else
            If boxGone Then
                ' FIX-5: the click WAS sent and the box IS gone - that part succeeded. Only
                ' the break PROOF is late, so this is a PENDING measurement (never a click
                ' failure, never a failed dismissal, never an End fallback).
                res("outcome") = "break_pending"
                res("breakPending") = True
                ' FIX-5b: the break transition is NOT observable from inside the call that
                ' triggers it (measured: 47 polls over 5 s all read EbMode=1; the NEXT call
                ' read 2). The numbers are DERIVED, never hard-coded, and the wording must
                ' not imply "just wait longer".
                res("breakPendingReason") = "the break transition is not observable from inside the call that triggers it: " & CStr(pollAttempts) & " polls over " & CStr(ReBreakWaitMs \ 1000) & " s read EbMode=" & CStr(ebNow) & " and the NEXT call reads 2. The click DID work (box gone) - re-probe with ide_state (or control_execution{end} again) and it will report break."
                res("howToConfirm") = "ide_state"
                Dim pendFin As Object
                Set pendFin = New Collection
                pendFin.Add "call ide_state (read-only) - it will report state=break / ebMode=2"
                pendFin.Add "do NOT click End on the dialog; nothing is left to dismiss (box is gone)"
                Set res("howToFinish") = pendFin
            Else
                ' STRUCTURED FAILURE: the click was sent but the box is STILL THERE - a real
                ' dismissal failure. Report it and never fall back to End.
                res("outcome") = "blocked"
                Set res("dismiss") = ReDismissFailBlock(Owner, IIf(viaClick = "debug", "dialog_still_present_after_debug_click", "dialog_still_present_after_ok_click"), dlg, btns, endBtn, ebNow, clickJson, True, True, ebBefore)
                Err.Clear
                Set ReCaptureRuntimeError = res
                Exit Function
            End If
        End If
    Else
        res("outcome") = "read_only"
    End If
    Set res("proof") = ReProofBlock(Owner, ebBefore, ebNow, waited, clicked, boxGone, pollAttempts)
    ' ---- position AFTER the break (the pane selection only means something there)
    Dim posJ As String
    posJ = ""
    If ebNow = 2 Then posJ = Owner.ToolCurrentPosition()
    res("position") = posJ
    Dim pModule As String
    Dim pProc As String
    Dim pLine As Long
    Dim pDisk As Long
    pModule = McpDb.DbPosField(Owner, posJ, "module")
    pProc = McpDb.DbPosField(Owner, posJ, "proc")
    pLine = McpDb.DbPosLine(Owner, posJ)
    pDisk = 0
    If Len(pModule) > 0 Then pDisk = pLine + McpCls.SharedHeaderOffset(Owner, pModule)
    res("proc") = pProc
    res("module") = pModule
    res("line") = pLine
    res("paneLine") = pLine
    res("diskLine") = pDisk
    ' ---- (5) the capture block + the cache ------------------------------
    Dim blk As Object
    Set blk = CreateObject("Scripting.Dictionary")
    If Len(numText) > 0 Then
        blk("number") = CLng(Val(numText))
    Else
        blk("number") = Null
    End If
    blk("description") = desc
    blk("source") = src
    blk("dialogText") = body
    blk("capturedAt") = readAt
    blk("proc") = pProc
    blk("module") = pModule
    blk("line") = pLine
    blk("paneLine") = pLine
    blk("diskLine") = pDisk
    blk("source_") = "live_error_dialog"
    blk("stillInBreak") = (ebNow = 2)
    blk("dialogTitle") = CStr(res("dialogTitle"))
    blk("dialogHwnd") = JsonGetLong(dlg, "hwnd", 0)
    Set blk("buttons") = btns
    blk("debugButtonPresent") = (Not dbg Is Nothing)
    blk("debugButtonHwnd") = 0
    If Not dbg Is Nothing Then blk("debugButtonHwnd") = JsonGetLong(dbg, "hwnd", 0)
    blk("okButtonPresent") = (Not okBtn Is Nothing)
    blk("okButtonHwnd") = 0
    If Not okBtn Is Nothing Then blk("okButtonHwnd") = JsonGetLong(okBtn, "hwnd", 0)
    If doClick Then blk("dismissedVia") = CStr(res("dismissedVia"))
    blk("endButtonPresent") = (Not endBtn Is Nothing)
    blk("endClicked") = False
    blk("readBeforeClick") = True
    blk("bodyChars") = Len(body)
    Set res("error") = blk
    If doClick Then
        ' The error TEXT is real whatever the click did, but only a COMPLETED resolution
        ' (measured break OR measured ended run) is cached. FIX-5: a pending break proof is
        ' NOT a failed click, so ok is not downgraded for it.
        If CStr(res("outcome")) = "break" Or CStr(res("outcome")) = "run_ended" Then
            Set Owner.m_lastRuntimeError = blk
            Owner.m_lastRuntimeErrorAt = readAt
            res("cacheStored") = True
            Set res("cachedCapture") = ReErrorCacheBlock(Owner)
        End If
    End If
    res("ok") = ((Not doClick) Or (CStr(res("outcome")) = "break") Or (CStr(res("outcome")) = "run_ended") Or (CStr(res("outcome")) = "break_pending"))
    If Err.Number <> 0 Then
        res("ok") = False
        res("errorMessage") = Err.Description
        Err.Clear
    End If
    Set ReCaptureRuntimeError = res
End Function

Public Function ReClassifyRuntimeDialog(Owner As CodeTools, ByVal dlg As Object) As Boolean
    ReClassifyRuntimeDialog = False
    Owner.m_reDetectMarker = ""
    Owner.m_reDetectReason = ""
    On Error Resume Next
    If dlg Is Nothing Then Exit Function
    Dim body As String
    body = UCase$(ReRuntimeErrorText(Owner, dlg))
    If Len(body) = 0 Then
        Err.Clear
        Exit Function
    End If
    ' FIX-4: BOTH Simplified-Chinese wordings are real. The LIVE one is the 4-character
    ' wording built from 5B9E 65F6 9519 8BEF; the OLDER 5-character wording
    ' (8FD0 884C 65F6 9519 8BEF) is kept for older builds. Source stays 7-bit ASCII.
    Dim zhShi As String
    Dim zhRun As String
    zhShi = ChrW$(&H5B9E) & ChrW$(&H65F6) & ChrW$(&H9519) & ChrW$(&H8BEF)
    zhRun = ChrW$(&H8FD0) & ChrW$(&H884C) & ChrW$(&H65F6) & ChrW$(&H9519) & ChrW$(&H8BEF)
    If InStr(1, body, zhShi) > 0 Then
        Owner.m_reDetectMarker = "zh_shishi"
        Owner.m_reDetectReason = "the dialog body matched the Simplified Chinese runtime-error wording built from code points 5B9E 65F6 9519 8BEF (the live wording on this IDE)"
        ReClassifyRuntimeDialog = True
        Err.Clear
        Exit Function
    End If
    If InStr(1, body, zhRun) > 0 Then
        Owner.m_reDetectMarker = "zh_yunxing"
        Owner.m_reDetectReason = "the dialog body matched the Simplified Chinese runtime-error wording built from code points 8FD0 884C 65F6 9519 8BEF (kept for older builds; this wording is NOT what the live box shows)"
        ReClassifyRuntimeDialog = True
        Err.Clear
        Exit Function
    End If
    If InStr(1, body, "RUN-TIME ERROR") > 0 Or InStr(1, body, "RUNTIME ERROR") > 0 Then
        Owner.m_reDetectMarker = "en_runtime"
        Owner.m_reDetectReason = "the dialog body matched the English wording RUN-TIME ERROR / RUNTIME ERROR"
        ReClassifyRuntimeDialog = True
        Err.Clear
        Exit Function
    End If
    If InStr(1, body, "RUN TIME ERROR") > 0 Then
        Owner.m_reDetectMarker = "en_run_time"
        Owner.m_reDetectReason = "the dialog body matched the English wording RUN TIME ERROR (spaces, no hyphen)"
        ReClassifyRuntimeDialog = True
        Err.Clear
        Exit Function
    End If
    ' STRUCTURAL fallback: the text did not match any known wording, so fall back to
    ' the STRUCTURE of a VB6 runtime-error box. A compile-error box has only one
    ' button (OK, the 2-character Chinese caption), so requiring BOTH the DEBUG and
    ' the END button excludes it;
    ' the quoted error number ("'11'") is the third independent signal.
    Dim dbg As Object
    Dim endBtn As Object
    Dim dgT As Double
    dgT = McpCls.DgNow(Owner)
    Set dbg = ReFindDebugButton(Owner, dlg)
    McpCls.DgSpan Owner, "ReFindDebugButton", dgT
    dgT = McpCls.DgNow(Owner)
    Set endBtn = ReFindEndButtonInDialog(Owner, dlg)
    McpCls.DgSpan Owner, "ReFindEndButtonInDialog", dgT
    If (Not (dbg Is Nothing)) And (Not (endBtn Is Nothing)) Then
        If ReHasQuotedNumber(Owner, body) Then
            Owner.m_reDetectMarker = "structural_debug_end_buttons"
            Owner.m_reDetectReason = "the text matched no known wording, but the dialog is this process's own #32770 carrying BOTH a DEBUG and an END button and a quoted error number in its body - that is the STRUCTURE of a VB6 runtime-error box, not a compile-error box (which has only OK)"
            ReClassifyRuntimeDialog = True
        End If
    End If
    Err.Clear
End Function

Public Function ReDialogInfoBlock(Owner As CodeTools, ByVal dlg As Object) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("present") = True
    d("hwnd") = 0
    d("title") = ""
    Set d("buttons") = New Collection
    If dlg Is Nothing Then
        Set ReDialogInfoBlock = d
        Exit Function
    End If
    On Error Resume Next
    d("hwnd") = JsonGetLong(dlg, "hwnd", 0)
    d("title") = CStr(dlg("title"))
    d("cls") = CStr(dlg("cls"))
    d("ownerTitle") = CStr(dlg("ownerTitle"))
    d("ownerDisabled") = CBool(dlg("ownerDisabled"))
    Set d("buttons") = ReButtonsJson(Owner, dlg)

    ' FIX-4: echo WHY this dialog is believed to be a runtime-error box.

    ' detectMarker is one of the 5 tags written by ReClassifyRuntimeDialog; the body

    ' length/head prove that text was really READ (a synthetic box cannot fake 0 bytes).

    Dim detectOk As Boolean

    detectOk = ReClassifyRuntimeDialog(Owner, dlg)

    Dim detectBody As String

    detectBody = ReRuntimeErrorText(Owner, dlg)

    d("detectMarker") = Owner.m_reDetectMarker

    d("detectReason") = Owner.m_reDetectReason

    d("detectBodyLen") = Len(detectBody)

    d("detectTextHead") = Left$(detectBody, 200)

    Set ReDialogInfoBlock = d
    Err.Clear
End Function

Public Function ReDialogIsRuntimeError(Owner As CodeTools, ByVal dlg As Object) As Boolean
    ' Boolean signature kept: callers (and the static checker) rely on it. The
    ' classification itself lives in ReClassifyRuntimeDialog, which also records
    ' detectMarker / detectReason for the receipt.
    ReDialogIsRuntimeError = ReClassifyRuntimeDialog(Owner, dlg)
End Function

Public Function ReDismissFailBlock(Owner As CodeTools, ByVal reason As String, ByVal dlg As Object, ByVal btns As Object, ByVal endBtn As Object, ByVal ebMeasured As Long, ByVal clickJson As String, ByVal stillPresent As Boolean, ByVal clicked As Boolean, ByVal ebBefore As Long) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("reason") = reason
    d("dialogTitle") = ""
    d("dialogHwnd") = 0
    If Not dlg Is Nothing Then
        d("dialogTitle") = CStr(dlg("title"))
        d("dialogHwnd") = JsonGetLong(dlg, "hwnd", 0)
    End If
    Set d("buttons") = btns
    d("measuredEbMode") = ebMeasured
    d("ebModeBeforeClick") = ebBefore
    d("endButtonPresent") = (Not endBtn Is Nothing)
    d("endClicked") = False
    d("clicked") = clicked
    d("clickReceipt") = Left$(clickJson, 800)
    d("dialogStillPresent") = stillPresent
    d("note") = "the runtime-error box was LEFT UNTOUCHED: NO End click was sent (End can leave the IDE unstable - handles may not be released). Read reason + buttons + measuredEbMode and decide yourself; this call dismissed nothing."
    d("measuredAt") = Format$(Now, "yyyy-mm-dd hh:nn:ss")
    Set ReDismissFailBlock = d
End Function

Public Function ReEmptyResult(Owner As CodeTools) As Object
    Dim r As Object
    Set r = CreateObject("Scripting.Dictionary")
    r("found") = False
    r("ok") = False
    r("enteredBreak") = False
    r("runEnded") = False
    r("outcome") = "none"
    r("dialogTitle") = ""
    Set r("error") = CreateObject("Scripting.Dictionary")
    Set r("proof") = CreateObject("Scripting.Dictionary")
    Set r("dismiss") = CreateObject("Scripting.Dictionary")
    r("cacheStored") = False
    r("position") = ""
    r("proc") = ""
    r("module") = ""
    r("line") = 0
    r("paneLine") = 0
    r("diskLine") = 0
    Set ReEmptyResult = r
End Function

Public Function ReEndWithRuntimeErrorDialog(Owner As CodeTools, ByVal m As String) As String
    Dim cbR As Object
    If Owner.CmdBarsRunRefusal("last_error:resolve", cbR) Then
        ReEndWithRuntimeErrorDialog = JsonStringify(cbR)
        Exit Function
    End If
    Dim root As Object
    Set root = CreateObject("Scripting.Dictionary")
    root("tool") = "control_execution"
    root("mode") = m
    root("ok") = False
    root("executed") = False
    root("stateBefore") = "runtime_error"
    root("state") = "runtime_error"
    root("resolvedRuntimeErrorDialog") = False
    root("enteredBreak") = False
    root("runEnded") = False
    Dim t0 As Long
    t0 = WoGetTickCount()
    Dim res As Object
    Set res = ReCaptureRuntimeError(Owner, True)
    If Not CBool(res("found")) Then
        root("reason") = "runtime_error_dialog_gone"
        root("state") = "unknown"
        root("advice") = "the runtime-error box was no longer there when the call started resolving it; nothing was executed - re-probe with ide_state / last_error"
        Owner.RunDiskAttach root, Nothing, False, "unknown"
        ReEndWithRuntimeErrorDialog = JsonStringify(root)
        Exit Function
    End If
    Set root("error") = res("error")
    Set root("runtimeError") = res("error")
    If res.Exists("dismissIdBasis") Then root("dismissIdBasis") = CStr(res("dismissIdBasis"))
    If res.Exists("dismissControlId") Then root("dismissControlId") = CLng(res("dismissControlId"))
    root("cacheStored") = CBool(res("cacheStored"))
    If res.Exists("postDebugClick") Then
        If IsObject(res("postDebugClick")) Then Set root("postDebugClick") = res("postDebugClick")
    End If
    If res.Exists("dialogInfo") Then
        If IsObject(res("dialogInfo")) Then Set root("runtimeErrorDialog") = res("dialogInfo")
    End If
    Dim ebAfter As Long
    Dim breakAttempts As Long
    Dim pdcObj As Object
    ebAfter = 0
    breakAttempts = 0
    Set pdcObj = Nothing
    If res.Exists("postDebugClick") Then
        If IsObject(res("postDebugClick")) Then Set pdcObj = res("postDebugClick")
    End If
    If Not pdcObj Is Nothing Then
        ebAfter = JsonGetLong(pdcObj, "ebMode", 0)
        breakAttempts = JsonGetLong(pdcObj, "pollAttempts", 0)
        If JsonGetBool(pdcObj, "pollAttemptsSuspect", False) Then root("breakPollAttemptsSuspect") = True
    End If
    If breakAttempts > 0 Then root("breakPollAttempts") = breakAttempts
    Dim outcome As String
    outcome = ""
    If res.Exists("outcome") Then outcome = CStr(res("outcome"))
    If outcome = "break" Then
        ' ---- proven break: the IDE's own End command is safe to execute NOW ----
        Dim ctrl As Object
        Dim why As String
        Dim n As Long
        why = ""
        Set ctrl = McpFindTools.FindCmd(Owner, "end", "Debug|Run|" & ChrW$(&H8C03) & ChrW$(&H8BD5) & "|" & ChrW$(&H8FD0) & ChrW$(&H884C), "End|" & ChrW$(&H7ED3) & ChrW$(&H675F))
        If ctrl Is Nothing Then Set ctrl = Owner.m_cachedEnd Else Set Owner.m_cachedEnd = ctrl
        If ctrl Is Nothing Then Set ctrl = McpFindTools.FindBarControlDeep(Owner, "End|" & ChrW$(&H7ED3) & ChrW$(&H675F))
        If Not ctrl Is Nothing Then n = Owner.ExecIfEnabledWhy(ctrl, m, why)
        root("executed") = (n > 0)
        ' The End command is DELAYED (measured: state_not_changed at return, design ~2 s later),
        ' so the confirmation is a bounded poll of ide_state, and a still-not-design result is
        ' reported as a structured failure instead of being claimed as success.
        Dim tEnd As Long
        Dim stNow As String
        Dim stImmediate As String
        Dim endAttempts As Long
        tEnd = WoGetTickCount()
        stNow = Owner.GetIdeState()
        stImmediate = stNow
        endAttempts = 1
        Do While stNow <> "design" And (WoGetTickCount() - tEnd) < ReEndWaitMs
            If Util.CanPump() Then DoEvents
            Owner.MaybeHonorCancel
            If Util.PumpAbortDue() Then Exit Do
            McpDb.DbBriefWait Owner, 100
            stNow = Owner.GetIdeState()
            endAttempts = endAttempts + 1
        Loop
        root("state") = stNow
        root("finalState") = stNow
        root("stateImmediatelyAfterEnd") = stImmediate
        ' endDeferred = the End command did NOT take effect by the time it returned and the
        ' bounded poll then confirmed design. endConfirmed is the bare measured outcome.
        root("endDeferred") = (stImmediate <> "design" And stNow = "design")
        root("endConfirmed") = (stNow = "design")
        Dim endWaitMs As Long
        endWaitMs = WoGetTickCount() - tEnd
        root("endElapsedMs") = endWaitMs
        root("endPollMs") = endWaitMs
        If RePollSane(Owner, endAttempts, endWaitMs) Then
            root("endPollAttempts") = endAttempts
        Else
            root("endPollAttemptsSuspect") = True
            root("endPollAttemptsRaw") = endAttempts
            root("endPollNote") = "counter failed sanity check (average < 10 ms per poll): do not trust endPollAttempts"
        End If
        root("resolvedRuntimeErrorDialog") = True
        root("enteredBreak") = True
        root("via") = "debug_then_end"
        root("boxDismissed") = True
        If m = "restart" Then
            If Not Owner.m_server Is Nothing Then Owner.m_server.RequestRestart
            root("restartPending") = True
        End If
        Set root("runtimeErrorDialogResolved") = ReResolvedBlock(Owner, "debug_then_end", ebAfter, WoGetTickCount() - t0)
        If stNow = "design" Then
            root("ok") = True
        Else
            root("ok") = False
            root("reason") = "end_not_confirmed_after_runtime_error_dialog"
            ' FIX-6: same additive pending layer as the generic end path; ok/reason kept.
            root("endQueued") = True
            root("confirmed") = False
            root("confirmedBy") = ""
            root("howToConfirm") = "ide_state"
            Dim reFin As Object
            Set reFin = New Collection
            reFin.Add "call ide_state (read-only) - it will report the state AFTER the IDE command dispatch; measured: design appears ~2 s after this call returns, never inside it"
            Set root("howToFinish") = reFin
            root("advice") = "the box was resolved (DEBUG clicked, EbMode=2 measured) and the IDE End command was issued, but ide_state did not reach design within the bounded " & CStr(ReEndWaitMs \ 1000) & " s poll; re-probe with ide_state"
        End If
    ElseIf outcome = "run_ended" Then
        ' ---- the DEBUG click ended the run: nothing is left to End ----
        root("state") = "design"
        root("runEnded") = True
        root("resolvedRuntimeErrorDialog") = True
        root("ok") = True
        root("alreadyNotRunning") = True
        root("state_not_changed") = True
        root("reason") = "already_not_running"
        root("via") = "debug_ended_run"
        root("boxDismissed") = True
        If m = "restart" Then
            root("restartPending") = False
            root("restartNote") = "the run had already ended when the box was resolved, so nothing was restarted - call control_execution{run} to start it again"
        End If
        Set root("runtimeErrorDialogResolved") = ReResolvedBlock(Owner, "debug_ended_run", ebAfter, WoGetTickCount() - t0)
    ElseIf outcome = "break_pending" Then
        ' ---- FIX-5: the box was READ and DEBUG clicked (box gone) but the break proof
        '      is still pending. The END command is NOT sent blindly; the caller confirms
        '      with ide_state. This is a PENDING resolution, not a refusal.
        root("state") = "run"
        root("enteredBreak") = False
        root("breakPending") = True
        root("breakPendingReason") = CStr(res("breakPendingReason"))
        root("breakPollMs") = ReBreakPollMs
        root("breakWaitMs") = ReBreakWaitMs
        root("howToConfirm") = "ide_state"
        Dim endFin As Object
        Set endFin = New Collection
        endFin.Add "call control_execution{end} again - it will find the measured break and issue the IDE End command"
        endFin.Add "or call ide_state to see the break first"
        Set root("howToFinish") = endFin
        root("boxDismissed") = True
        root("resolvedRuntimeErrorDialog") = False
        root("ok") = True
        root("executed") = False
        root("via") = "debug_click_break_pending"
        root("ebMode") = ebAfter
        root("elapsedMs") = WoGetTickCount() - t0
        root("advice") = "the runtime-error box was READ and DEBUG was clicked (the box is GONE). The break transition is not observable from inside the call that triggers it, so no break proof is available here; the NEXT call observes it. No End click and no IDE End command were sent (never blindly). Follow howToFinish: call control_execution{end} again, or ide_state first."
        Set root("runtimeErrorDialogResolved") = ReResolvedBlock(Owner, "debug_click_break_pending", ebAfter, WoGetTickCount() - t0)
    Else
        ' ---- neither break nor design: structured failure, the box stays untouched ----
        root("state") = "error_dialog"
        root("reason") = "blocked_by_runtime_error_dialog"
        root("resolvedRuntimeErrorDialog") = False
        root("ebMode") = ebAfter
        root("elapsedMs") = WoGetTickCount() - t0
        root("runtimeErrorDialogDismissFailed") = True
        If res.Exists("dismiss") Then
            If IsObject(res("dismiss")) Then Set root("runtimeErrorDismiss") = res("dismiss")
        End If
        Set root("runtimeErrorDialogResolved") = ReResolvedBlock(Owner, "none", ebAfter, WoGetTickCount() - t0)
        root("advice") = "the runtime-error box could not be turned into a break or a finished run: NO End click was sent on the DIALOG and the IDE End command was NOT executed. Read ebMode + runtimeErrorDismiss{buttons,measuredEbMode} + elapsedMs and decide yourself"
    End If
    If root.Exists("endPollAttemptsSuspect") Then
        If CBool(root("endPollAttemptsSuspect")) Then
            If root.Exists("advice") Then root("advice") = CStr(root("advice")) & " " & CStr(root("endPollNote"))
        End If
    End If
    Owner.RunDiskAttach root, Nothing, CBool(root("executed")), CStr(root("state"))
    ReEndWithRuntimeErrorDialog = JsonStringify(root)
End Function

Public Function ReErrorCacheBlock(Owner As CodeTools) As Object
    Set ReErrorCacheBlock = Nothing
    If Owner.m_lastRuntimeError Is Nothing Then Exit Function
    On Error Resume Next
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("number") = Owner.m_lastRuntimeError("number")
    d("description") = Owner.m_lastRuntimeError("description")
    d("source") = Owner.m_lastRuntimeError("source")
    d("dialogText") = Owner.m_lastRuntimeError("dialogText")
    d("capturedAt") = Owner.m_lastRuntimeError("capturedAt")
    d("proc") = Owner.m_lastRuntimeError("proc")
    d("module") = Owner.m_lastRuntimeError("module")
    d("line") = Owner.m_lastRuntimeError("line")
    d("paneLine") = Owner.m_lastRuntimeError("paneLine")
    d("diskLine") = Owner.m_lastRuntimeError("diskLine")
    d("source_") = "cached_runtime_error"
    d("stillInBreak") = (IdeEbMode() = 2)
    Set ReErrorCacheBlock = d
    Err.Clear
End Function

Public Function ReFindButton(Owner As CodeTools, ByVal dlg As Object, ByVal eng As String, ByVal zh As String) As Object
    Set ReFindButton = Nothing
    On Error Resume Next
    If dlg Is Nothing Then Exit Function
    Dim btns As Object
    Set btns = Nothing
    If dlg.Exists("buttons") Then Set btns = dlg("buttons")
    If btns Is Nothing Then Exit Function
    Dim i As Long
    Dim b As Object
    Dim n As String
    For i = 1 To btns.Count
        Set b = btns(i)
        If Not b Is Nothing Then
            n = ReNormCaption(Owner, CStr(b("text")))
            If Len(n) > 0 Then
                If StrComp(n, UCase$(eng), vbTextCompare) = 0 Then
                    Set ReFindButton = b
                    Exit Function
                End If
                If Len(zh) > 0 Then
                    If InStr(1, n, zh) > 0 Then
                        Set ReFindButton = b
                        Exit Function
                    End If
                End If
            End If
        End If
    Next
    Err.Clear
End Function

Public Function ReIdFirstButton(Owner As CodeTools, ByVal dlg As Object, ByVal want As String, ByVal allowNotClickable As Boolean) As Object
    ' (39) ID-FIRST lookup for the runtime-error box: measured Debug=4801, End=4800,
    ' Help=4902, Continue=4802 (WinOps.WoDlgIdSemantic). Text is only the fallback and
    ' the secondary cross-check; a non-Button / unknown id / text conflict returns Nothing
    ' so the caller keeps the exact legacy caption behaviour.
    Set ReIdFirstButton = Nothing
    On Error Resume Next
    If dlg Is Nothing Then Exit Function
    If Not WinOps.WoDlgIdFirstOn() Then Exit Function
    Dim kind As String
    kind = WinOps.WoDlgKindOf(dlg)
    If kind <> "runtime_error" Then Exit Function
    Dim btns As Object
    Set btns = Nothing
    If dlg.Exists("buttons") Then Set btns = dlg("buttons")
    If btns Is Nothing Then Exit Function
    Dim i As Long
    Dim b As Object
    Dim r As Object
    For i = 1 To btns.Count
        Set b = btns(i)
        If Not b Is Nothing Then
            If CBool(b("enabled")) Then
                Set r = WinOps.WoDlgIdResolve(kind, CLng(b("id")), CStr(b("cls")), CBool(b("enabled")), CStr(b("text")), "")
                If Not r Is Nothing Then
                    If StrComp(CStr(r("semantics")), want, vbTextCompare) = 0 Then
                        If CBool(r("ok")) Or allowNotClickable Then
                            Set ReIdFirstButton = b
                            Exit Function
                        End If
                    End If
                End If
            End If
        End If
    Next
    Err.Clear
End Function

Public Function ReFindDebugButton(Owner As CodeTools, ByVal dlg As Object) As Object
    ' the ChrW$ codes below are the localized DEBUG caption; the SOURCE stays 7-bit
    ' ASCII on purpose (string-literal policy in docs/SOURCE_ENCODING.md)
    Set ReFindDebugButton = ReFindButton(Owner, dlg, "Debug", ChrW$(&H8C03) & ChrW$(&H8BD5))
End Function

Public Function ReFindEndButtonInDialog(Owner As CodeTools, ByVal dlg As Object) As Object
    ' SEARCH ONLY. This caption is NEVER clicked on the runtime-error path; it is found
    ' so the receipt can report endButtonPresent without ever sending it a message.
    Set ReFindEndButtonInDialog = ReFindButton(Owner, dlg, "End", ChrW$(&H7ED3) & ChrW$(&H675F))
End Function

Public Function ReFindOkButton(Owner As CodeTools, ByVal dlg As Object) As Object
    Set ReFindOkButton = ReFindButton(Owner, dlg, "OK", ChrW$(&H786E) & ChrW$(&H5B9A))
End Function

Public Function ReFindRuntimeErrorDialog(Owner As CodeTools) As Object
    Set ReFindRuntimeErrorDialog = Nothing
    On Error Resume Next
    Dim dialogs As Object
    Set dialogs = WinOps.DialogObjectsByClass("#32770", True, 16)
    If dialogs Is Nothing Then Exit Function
    Dim i As Long
    Dim d As Object
    For i = 1 To dialogs.Count
        Set d = dialogs(i)
        If Not d Is Nothing Then
            If ReDialogIsRuntimeError(Owner, d) Then
                Set ReFindRuntimeErrorDialog = d
                Exit Function
            End If
        End If
    Next
    Err.Clear
End Function

Public Function ReHasQuotedNumber(Owner As CodeTools, ByVal body As String) As Boolean
    ReHasQuotedNumber = False
    Dim n As Long
    Dim i As Long
    Dim j As Long
    Dim ch As String
    n = Len(body)
    i = 1
    Do While i <= n
        If Mid$(body, i, 1) = "'" Then
            If i + 1 <= n Then
                ch = Mid$(body, i + 1, 1)
                If ch >= "0" And ch <= "9" Then
                    j = i + 1
                    Do While j <= n
                        ch = Mid$(body, j, 1)
                        If ch >= "0" And ch <= "9" Then
                            j = j + 1
                        Else
                            Exit Do
                        End If
                    Loop
                    If j <= n Then
                        If Mid$(body, j, 1) = "'" Then
                            ReHasQuotedNumber = True
                            Exit Function
                        End If
                    End If
                End If
            End If
        End If
        i = i + 1
    Loop
End Function

Public Sub ReInvalidateRuntimeError(Owner As CodeTools, ByVal why As String)
    Set Owner.m_lastRuntimeError = Nothing
    Owner.m_lastRuntimeErrorAt = ""
    Owner.m_lastRuntimeErrorInvalidatedBy = why
    Owner.m_lastRuntimeErrorInvalidatedAt = Format$(Now, "yyyy-mm-dd hh:nn:ss")
End Sub

Public Function ReMergeRuntimeError(Owner As CodeTools, ByRef root As Object, ByRef res As Object) As String
    On Error Resume Next
    ReMergeRuntimeError = "runtime_error_dialog_dismiss_failed"
    root("enteredBreak") = CBool(res("enteredBreak"))
    root("runEnded") = CBool(res("runEnded"))
    Set root("error") = res("error")
    Set root("runtimeError") = res("error")
    Set root("runtimeErrorBreakProof") = res("proof")
    If res.Exists("postDebugClick") Then
        If IsObject(res("postDebugClick")) Then Set root("postDebugClick") = res("postDebugClick")
    End If
    If res.Exists("dialogInfo") Then
        If IsObject(res("dialogInfo")) Then Set root("runtimeErrorDialog") = res("dialogInfo")
    End If
    root("cacheStored") = CBool(res("cacheStored"))
    If res.Exists("cachedCapture") Then
        If IsObject(res("cachedCapture")) Then Set root("cachedCapture") = res("cachedCapture")
    End If
    Dim outcome As String
    outcome = ""
    If res.Exists("outcome") Then outcome = CStr(res("outcome"))
    If outcome = "break" Then
        ReMergeRuntimeError = "runtime_error"
        root("state") = "break"
        root("proc") = res("proc")
        root("module") = res("module")
        root("line") = res("line")
        root("paneLine") = res("paneLine")
        root("diskLine") = res("diskLine")
        If Len(CStr(res("position"))) > 0 Then
            root("position") = res("position")
            root("positionSource") = "paneSelection"
        End If
        root("stopReason") = "runtime_error"
        root("note") = "the runtime-error box was READ, then DEBUG was clicked and break was PROVEN by measuring VBA6 EbMode = 2; the dialog's End button was NOT clicked"
    ElseIf outcome = "run_ended" Then
        ReMergeRuntimeError = "runtime_error_run_ended"
        root("state") = "design"
        root("stopReason") = "runtime_error_run_ended"
        root("note") = "the runtime-error box was READ, then DEBUG was clicked and the run ENDED instead of stopping at the error line (measured VBA6 EbMode = 0 / design; a queued End can also fire when the box disappears). The captured error is returned and cached; there is NO break position and no locals - see postDebugClick"
    ElseIf outcome = "break_pending" Then
        ReMergeRuntimeError = "runtime_error_break_pending"
        root("state") = "unknown"
        root("breakPending") = True
        root("breakPendingReason") = res("breakPendingReason")
        root("howToConfirm") = "ide_state"
        If res.Exists("howToFinish") Then
            If IsObject(res("howToFinish")) Then Set root("howToFinish") = res("howToFinish")
        End If
        root("stopReason") = "runtime_error_break_pending"
        root("note") = "the runtime-error box was READ and DEBUG was clicked (the box is GONE) but the break PROOF did not arrive inside the bounded EbMode poll. This is a PENDING MEASUREMENT, not a failed dismissal and not a click failure. Re-probe with ide_state: the engine usually reaches break shortly after the modal loop unwinds. No End click was sent."
    Else
        root("state") = "error_dialog"
        root("runtimeErrorDialogDismissFailed") = True
        Set root("runtimeErrorDismiss") = res("dismiss")
        root("stopReason") = "runtime_error_dialog_dismiss_failed"
        root("note") = "the runtime-error box was LEFT UNTOUCHED: its End button was NOT clicked (End can leave the IDE unstable). Read runtimeErrorDismiss{reason,buttons,measuredEbMode} and decide yourself"
    End If
    Err.Clear
End Function

Public Function ReNormCaption(Owner As CodeTools, ByVal t As String) As String
    Dim s As String
    Dim p As Long
    s = Trim$(t)
    s = Replace(s, "&", "")
    p = InStr(s, "(")
    If p = 0 Then p = InStr(s, ChrW$(&HFF08))
    If p > 0 Then s = Trim$(Left$(s, p - 1))
    Do While Len(s) > 0
        If Right$(s, 1) = "." Then s = Left$(s, Len(s) - 1) Else Exit Do
    Loop
    ReNormCaption = UCase$(Trim$(s))
End Function

Public Function ReParseErrorDescription(Owner As CodeTools, ByVal body As String) As String
    ' D1-r (OPEN-42 section 11 live retest): this function used to be a SECOND, weaker parser -
    ' it split the header at the first colon and never stripped the QUOTED error number, so a
    ' punctuation-only body produced the fragment "':" which the cached_runtime_error channel
    ' then re-published as errDescription. There is now ONE parser: the H3 implementation
    ' strips the closing quote BEFORE the colon and rejects a punctuation-only candidate, so a
    ' failed parse honestly stays "". Kept as a thin wrapper because the legacy call paths
    ' (ReMergeRuntimeError / McpRe line 96) address it by name.
    ReParseErrorDescription = McpBpCore.H3ParseErrDescription(Owner, body)
End Function

Public Function ReParseErrorNumber(Owner As CodeTools, ByVal body As String) As String
    ReParseErrorNumber = ""
    On Error Resume Next
    Dim zh As String
    zh = ChrW$(&H8FD0) & ChrW$(&H884C) & ChrW$(&H65F6) & ChrW$(&H9519) & ChrW$(&H8BEF)
    Dim i As Long
    i = InStr(1, body, "Run-time error", vbTextCompare)
    If i = 0 Then i = InStr(1, body, "Runtime error", vbTextCompare)
    If i = 0 Then i = InStr(1, body, zh)
    If i = 0 Then
        ReParseErrorNumber = McpBpCore.H3ParseErrNumber(Owner, body)
        Exit Function
    End If
    Dim j As Long
    Dim ch As String
    Dim digits As String
    digits = ""
    j = i
    Do While j <= Len(body)
        ch = Mid$(body, j, 1)
        If ch >= "0" And ch <= "9" Then
            digits = digits & ch
        ElseIf Len(digits) > 0 Then
            Exit Do
        End If
        j = j + 1
    Loop
    ReParseErrorNumber = digits
End Function

Public Function RePollBreakMs(Owner As CodeTools, ByVal clicked As Boolean, ByRef ebNow As Long, ByRef pollAttempts As Long) As Long
    ' FIX-5: bounded EbMode poll after the DEBUG BM_CLICK. The modal runtime-error loop
    ' unwinds ASYNCHRONOUSLY, so ONE sample is a lie (measured false negative). While
    ' waiting this function touches NOTHING except IdeEbMode() and a short sleep:
    ' no CommandBars, no Swing, no window messages.
    Dim t0 As Long
    pollAttempts = 0
    If Not clicked Then
        RePollBreakMs = 0
        Exit Function
    End If
    t0 = Util.TickMs()
    Do
        ebNow = IdeEbMode()
        pollAttempts = pollAttempts + 1
        If ebNow = 2 Or ebNow = 0 Then Exit Do
        If (Util.TickMs() - t0) >= ReBreakWaitMs Then Exit Do
        If Util.CanPump() Then DoEvents
        McpDb.DbBriefWait Owner, ReBreakPollMs
    Loop
    RePollBreakMs = Util.TickMs() - t0
End Function

Public Function RePollSane(Owner As CodeTools, ByVal attempts As Long, ByVal elapsedMs As Long) As Boolean
    RePollSane = (attempts <= (elapsedMs \ 10) + 2)
End Function

Public Function RePostDebugBlock(Owner As CodeTools, ByVal eb As Long, ByVal modeRaw As String, ByVal enteredBreak As Boolean, ByVal runEnded As Boolean, ByVal elapsedMs As Long, ByVal clicked As Boolean, ByVal boxGone As Boolean, ByVal pollAttempts As Long) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("ebMode") = eb
    d("mode") = modeRaw
    d("enteredBreak") = enteredBreak
    d("runEnded") = runEnded
    d("elapsedMs") = elapsedMs
    d("clicked") = clicked
    d("boxGone") = boxGone
    If RePollSane(Owner, pollAttempts, elapsedMs) Then
        d("pollAttempts") = pollAttempts
    Else
        d("pollAttemptsSuspect") = True
        d("pollAttemptsRaw") = pollAttempts
    End If
    If enteredBreak Then
        d("confirmedBy") = "ebMode_poll"
    Else
        d("confirmedBy") = ""
    End If
    d("signal") = "measured AFTER the DEBUG click: VBA6.DLL EbMode (0=design, 1=run, 2=break)"
    Set RePostDebugBlock = d
End Function

Public Function ReProofBlock(Owner As CodeTools, ByVal ebBefore As Long, ByVal ebAfter As Long, ByVal waitedMs As Long, ByVal clicked As Boolean, ByVal boxGone As Boolean, ByVal pollAttempts As Long) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("signal") = "VBA6.DLL EbMode (2 = break context); measured, never inferred"
    d("ebModeBeforeClick") = ebBefore
    d("ebModeAfterClick") = ebAfter
    d("ebMode") = ebAfter
    d("verified") = (ebAfter = 2)
    d("clicked") = clicked
    d("waitedMs") = waitedMs
    d("boxGoneAfterClick") = boxGone
    If RePollSane(Owner, pollAttempts, waitedMs) Then
        d("pollAttempts") = pollAttempts
    Else
        d("pollAttemptsSuspect") = True
        d("pollAttemptsRaw") = pollAttempts
    End If
    If ebAfter = 2 Then
        d("confirmedBy") = "ebMode_poll"
    Else
        d("confirmedBy") = ""
    End If
    Set ReProofBlock = d
End Function

Public Function ReRawModeFromEb(Owner As CodeTools, ByVal eb As Long) As String
    ReRawModeFromEb = "run"
    If eb = 0 Then ReRawModeFromEb = "design"
    If eb = 2 Then ReRawModeFromEb = "break"
End Function

Public Function ReResolvedBlock(Owner As CodeTools, ByVal how As String, ByVal ebAfter As Long, ByVal elapsedMs As Long) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("how") = how
    d("ebModeAfterClick") = ebAfter
    d("elapsedMs") = elapsedMs
    d("capturedBy") = "READ box -> click DEBUG -> measure VBA6 EbMode"
    Set ReResolvedBlock = d
End Function

Public Function ReRuntimeErrorText(Owner As CodeTools, ByVal dlg As Object) As String
    ReRuntimeErrorText = ""
    On Error Resume Next
    If dlg Is Nothing Then Exit Function
    Dim s As String
    s = CStr(dlg("title")) & vbLf
    Dim ctrls As Object
    Set ctrls = Nothing
    If dlg.Exists("controls") Then Set ctrls = dlg("controls")
    If Not ctrls Is Nothing Then
        Dim i As Long
        For i = 1 To ctrls.Count
            s = s & CStr(ctrls(i)("text")) & vbLf
        Next
    End If
    ReRuntimeErrorText = s
    Err.Clear
End Function

Public Sub RpBusyWaitForSpan(Owner As CodeTools, ByRef out As Object)
    out("attempts") = 0
    out("clockImpl") = ""
    out("deltaMs") = Null
    out("unavailableReason") = ""
    Dim impl As String
    impl = RpPickAnyClock(Owner)
    out("clockImpl") = impl
    If Len(impl) = 0 Then
        out("unavailableReason") = "no clock read succeeded, so a pump wait cannot prove that time passed"
        Exit Sub
    End If
    Dim v0 As Double
    v0 = RpClockValue(Owner, RpClockRead(Owner, impl))
    Dim tries As Long
    Do While tries < 200000
        tries = tries + 1
        If Util.CanPump() Then DoEvents
        Dim v As Double
        v = RpClockValue(Owner, RpClockRead(Owner, impl))
        If v <> v0 Then
            out("attempts") = tries
            out("deltaMs") = v - v0
            Exit Sub
        End If
    Loop
    out("attempts") = tries
    out("unavailableReason") = "the chosen clock did not change across " & CStr(tries) & " pump iterations"
End Sub

Public Function RpBytesContain(Owner As CodeTools, ByVal hay As String, ByVal needle As String) As Boolean
    RpBytesContain = False
    If Len(needle) = 0 Then Exit Function
    If Len(hay) = 0 Then Exit Function
    RpBytesContain = (InStrB(1, hay, needle, vbBinaryCompare) > 0)
End Function

Public Function RpClassifyDialogTitle(Owner As CodeTools, ByVal s As String) As String
    On Error Resume Next
    Dim h As String
    h = LCase$(s)
    RpClassifyDialogTitle = "unrecognized_host_dialog"
    If InStr(1, h, "run-time error", vbTextCompare) > 0 Then
        RpClassifyDialogTitle = "program_error"
        Exit Function
    End If
    If InStr(1, h, "runtime error", vbTextCompare) > 0 Then
        RpClassifyDialogTitle = "program_error"
        Exit Function
    End If
    If InStr(1, s, ChrW(&H8FD0) & ChrW(&H884C) & ChrW(&H65F6) & ChrW(&H9519) & ChrW(&H8BEF)) > 0 Then
        RpClassifyDialogTitle = "program_error"
        Exit Function
    End If
    If InStr(1, h, "find", vbTextCompare) > 0 Then
        RpClassifyDialogTitle = "ide_own_dialog"
        Exit Function
    End If
    If InStr(1, h, "replace", vbTextCompare) > 0 Then
        RpClassifyDialogTitle = "ide_own_dialog"
        Exit Function
    End If
    If InStr(1, h, "options", vbTextCompare) > 0 Then
        RpClassifyDialogTitle = "ide_own_dialog"
        Exit Function
    End If
    If InStr(1, h, "references", vbTextCompare) > 0 Then
        RpClassifyDialogTitle = "ide_own_dialog"
        Exit Function
    End If
    If InStr(1, h, "components", vbTextCompare) > 0 Then
        RpClassifyDialogTitle = "ide_own_dialog"
        Exit Function
    End If
    If InStr(1, h, "object browser", vbTextCompare) > 0 Then
        RpClassifyDialogTitle = "ide_own_dialog"
        Exit Function
    End If
    If InStr(1, h, "project properties", vbTextCompare) > 0 Then
        RpClassifyDialogTitle = "ide_own_dialog"
        Exit Function
    End If
    If InStr(1, s, ChrW(&H67E5) & ChrW(&H627E)) > 0 Then
        RpClassifyDialogTitle = "ide_own_dialog"
        Exit Function
    End If
    If InStr(1, s, ChrW(&H66FF) & ChrW(&H6362)) > 0 Then
        RpClassifyDialogTitle = "ide_own_dialog"
        Exit Function
    End If
    If InStr(1, s, ChrW(&H9009) & ChrW(&H9879)) > 0 Then
        RpClassifyDialogTitle = "ide_own_dialog"
        Exit Function
    End If
    If InStr(1, s, ChrW(&H5F15) & ChrW(&H7528)) > 0 Then
        RpClassifyDialogTitle = "ide_own_dialog"
        Exit Function
    End If
    If InStr(1, s, ChrW(&H5DE5) & ChrW(&H7A0B) & ChrW(&H5C5E) & ChrW(&H6027)) > 0 Then
        RpClassifyDialogTitle = "ide_own_dialog"
        Exit Function
    End If
End Function

Public Function RpClockMs(Owner As CodeTools) As Long
    On Error Resume Next
    Dim s As Single
    s = Timer
    If Err.Number <> 0 Then
        Err.Clear
        RpClockMs = 0
        Exit Function
    End If
    RpClockMs = CLng(s * 1000#)
End Function

Public Function RpClockRead(Owner As CodeTools, ByVal kind As String) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("impl") = kind
    d("ok") = False
    d("value") = Null
    d("errNumber") = 0
    d("errDescription") = ""
    On Error GoTo RD_ERR
    Dim v As Double
    If kind = "codetools_gettickcount" Then
        v = CDbl(WoGetTickCount())
    ElseIf kind = "util_tickms" Then
        v = CDbl(Util.TickMs())
    ElseIf kind = "vb_timer" Then
        v = CDbl(Timer) * 1000#
    ElseIf kind = "calendar_now" Then
        v = CDbl(Now) * 86400000#
    Else
        d("errDescription") = "unknown clock kind"
        Set RpClockRead = d
        Exit Function
    End If
    On Error GoTo 0
    d("ok") = True
    d("value") = v
    Set RpClockRead = d
    Exit Function
RD_ERR:
    d("errNumber") = Err.Number
    d("errDescription") = Err.Description
    Err.Clear
    On Error GoTo 0
    Set RpClockRead = d
End Function

Public Function RpClockReadAll(Owner As CodeTools) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    Set d("codetools_gettickcount") = RpClockRead(Owner, "codetools_gettickcount")
    Set d("util_tickms") = RpClockRead(Owner, "util_tickms")
    Set d("vb_timer") = RpClockRead(Owner, "vb_timer")
    Set d("calendar_now") = RpClockRead(Owner, "calendar_now")
    Set RpClockReadAll = d
End Function

Public Function RpClockSelfTest(Owner As CodeTools) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    RpSelfTestInto Owner, d
    Set RpClockSelfTest = d
End Function

Public Function RpClockValue(Owner As CodeTools, ByRef r As Object) As Double
    RpClockValue = 0#
    If r Is Nothing Then Exit Function
    If Not r.Exists("ok") Then Exit Function
    If Not CBool(r("ok")) Then Exit Function
    If IsNull(r("value")) Then Exit Function
    RpClockValue = CDbl(r("value"))
End Function

Public Function RpDeltaAll(Owner As CodeTools, ByRef a As Object, ByRef b As Object) As Object
    Dim d As Object
    Dim kinds As Object
    Dim k As Variant
    Set d = CreateObject("Scripting.Dictionary")
    Set kinds = New Collection
    kinds.Add "codetools_gettickcount"
    kinds.Add "util_tickms"
    kinds.Add "vb_timer"
    kinds.Add "calendar_now"
    On Error Resume Next
    For Each k In kinds
        If a Is Nothing Or b Is Nothing Then
            d(CStr(k)) = Null
        ElseIf a.Exists(CStr(k)) And b.Exists(CStr(k)) Then
            Dim ra As Object
            Dim rb As Object
            Set ra = a(CStr(k))
            Set rb = b(CStr(k))
            If CBool(ra("ok")) And CBool(rb("ok")) Then
                Dim dv As Double
                dv = CDbl(rb("value")) - CDbl(ra("value"))
                If dv < 0# Then dv = dv + 86400000#
                d(CStr(k)) = dv
            Else
                d(CStr(k)) = Null
            End If
        Else
            d(CStr(k)) = Null
        End If
    Next
    Err.Clear
    Set RpDeltaAll = d
End Function

Public Function RpDialogBlockFromTitle(Owner As CodeTools, ByVal title As String, ByVal cls As String) As Object
    On Error Resume Next
    Dim blk As Object
    Dim one As Object
    Dim list As Object
    Set blk = CreateObject("Scripting.Dictionary")
    Set list = New Collection
    Set one = CreateObject("Scripting.Dictionary")
    one("title") = title
    one("classifiedAs") = cls
    one("isProgramError") = (cls = "program_error")
    list.Add one
    blk("present") = True
    blk("count") = 1
    Set blk("dialogs") = list
    blk("rule") = "a #32770 dialog is CLASSIFIED before it can be reported as an error: only a VB6 runtime-error box is program_error; an IDE dialog (Find/Replace/Options/References/Components/Project Properties) is ide_own_dialog and is NOT a program error"
    Set RpDialogBlockFromTitle = blk
End Function

Public Function RpDialogPrecheck(Owner As CodeTools, ByRef root As Object) As String
    On Error Resume Next
    RpDialogPrecheck = ""
    Dim re As Object
    Set re = ReFindRuntimeErrorDialog(Owner)
    If Not re Is Nothing Then
        root("observePrecheck") = "a VB6 runtime-error dialog is already up"
        root("runningState") = "error"
        root("observedState") = "error"
        root("state") = "error"
        root("ok") = False
        root("err") = "a VB6 runtime-error dialog is up"
        Set root("errorDialog") = McpBpCore.H3DialogBlock(Owner, "runtime error")
        RpDialogPrecheck = "program_error"
        Exit Function
    End If
    Dim rj As String
    rj = WinOps.ReadActiveDialogsJson("#32770", True, 3)
    Dim v As Variant
    If Not JsonParse(rj, v) Then Exit Function
    Dim top As Object
    Set top = v
    If top Is Nothing Then Exit Function
    If Not top.Exists("dialogs") Then Exit Function
    Dim arr As Object
    Set arr = top("dialogs")
    If arr Is Nothing Then Exit Function
    Dim list As Object
    Dim blocked As Long
    Set list = New Collection
    blocked = 0
    Dim dlg As Variant
    For Each dlg In arr
        If IsObject(dlg) Then
            Dim t As String
            Dim cc As String
            Dim k As String
            Dim d As Object
            t = JsonGetStr(dlg, "title", "")
            cc = JsonGetStr(dlg, "cls", "")
            k = RpClassifyDialogTitle(Owner, t)
            Set d = CreateObject("Scripting.Dictionary")
            d("title") = t
            d("cls") = cc
            d("hwnd") = JsonGetLong(dlg, "hwnd", 0)
            d("classifiedAs") = k
            d("isProgramError") = (k = "program_error")
            list.Add d
            If k <> "program_error" Then blocked = blocked + 1
        End If
    Next
    If list.Count = 0 Then Exit Function
    Dim blk As Object
    Set blk = CreateObject("Scripting.Dictionary")
    blk("present") = True
    blk("count") = list.Count
    Set blk("dialogs") = list
    blk("rule") = "a #32770 dialog is CLASSIFIED before it can be reported as an error: only a VB6 runtime-error box is program_error; an IDE dialog is ide_own_dialog and is NOT a program error"
    Set root("dialogPrecheck") = blk
    If blocked > 0 Then
        root("ok") = False
        root("reason") = "observe_blocked_by_ide_dialog"
        root("observedState") = "modal_dialog"
        root("isProgramError") = False
        root("retrySafe") = True
        root("failureKind") = "toolError"
        root("howToProceed") = "a #32770 dialog is up and it is NOT a program error. Close or answer it (close_active_dialogs for an answerable box), then observe again; observe never clicks a dialog."
        RpDialogPrecheck = "blocked"
    End If
End Function

Public Function RpElapsedMs(Owner As CodeTools, ByVal t0 As Long) As Long
    On Error Resume Next
    Dim d As Long
    d = RpClockMs(Owner) - t0
    If d < 0 Then d = d + 86400000
    If d < 0 Then d = 0
    RpElapsedMs = d
End Function

Public Sub RpErrNote(Owner As CodeTools, ByRef diag As Object, ByVal procName As String)
    On Error Resume Next
    If diag Is Nothing Then Exit Sub
    If Err.Number <> 0 Then
        diag("errNumber") = Err.Number
        diag("errDescription") = Err.Description
        diag("errProc") = procName
        Err.Clear
    End If
End Sub

Public Function RpMarkerProbe(Owner As CodeTools, ByVal p As String, ByVal contains As String, ByRef info As Object) As Boolean
    On Error Resume Next
    RpMarkerProbe = False
    If info Is Nothing Then Exit Function
    info("path") = p
    info("contains") = contains
    info("matched") = False
    info("matchedBy") = ""
    info("bytesRead") = 0
    info("error") = ""
    info("caseSensitive") = True
    info("readLimitBytes") = 65536
    info("pathSemantics") = "path is on the machine that runs this VB6 IDE (a remote middle service judges it there); only the first 65536 bytes are read and NOTHING is written"
    If Len(p) = 0 Then
        info("error") = "path_empty"
        Exit Function
    End If
    If Len(Dir$(p)) = 0 Then
        info("error") = "file_not_found_or_unreadable"
        Exit Function
    End If
    Dim hay As String
    hay = RpReadFilePrefixBytes(Owner, p, 65536)
    info("bytesRead") = Len(hay)
    If Len(hay) = 0 Then
        info("error") = "file_empty_or_unreadable"
        Exit Function
    End If
    Dim ansiNeedle As String
    Dim utf8Needle As String
    ansiNeedle = StrConv(contains, vbFromUnicode)
    utf8Needle = RpUtf8Bytes(Owner, contains)
    If RpBytesContain(Owner, hay, ansiNeedle) Then
        RpMarkerProbe = True
        info("matched") = True
        If ansiNeedle = utf8Needle Then
            info("matchedBy") = "bytes"
        Else
            info("matchedBy") = "ansi"
        End If
    ElseIf utf8Needle <> ansiNeedle Then
        If RpBytesContain(Owner, hay, utf8Needle) Then
            RpMarkerProbe = True
            info("matched") = True
            info("matchedBy") = "utf8"
        End If
    End If
End Function

Public Function RpObserve(Owner As CodeTools, ByRef root As Object, ByVal timeoutMs As Long, ByVal maxWaitMs As Long, ByVal markerPath As String, ByVal markerContains As String) As String
    On Error Resume Next
    Dim t0 As Long
    t0 = RpClockMs(Owner)
    root("clockSource") = "vb_timer_seconds_since_midnight"
    Dim budgetMs As Long
    budgetMs = timeoutMs
    root("budgetSource") = "timeout_ms"
    If maxWaitMs > 0 Then
        budgetMs = maxWaitMs
        root("budgetSource") = "observe_max_wait_ms"
    End If
    If budgetMs < 1000 Then budgetMs = 1000
    If budgetMs > 120000 Then budgetMs = 120000
    root("budgetMs") = budgetMs
    root("observeMaxWaitMs") = maxWaitMs
    If maxWaitMs > 0 Then
        root("budgetBasis") = "observe_max_wait_ms"
    Else
        root("budgetBasis") = "timeout_ms"
    End If
    root("timeoutMs") = timeoutMs
    root("state") = ""
    root("observedState") = ""
    root("elapsedMs") = 0
    root("timeout") = False
    Dim diag As Object
    Set diag = CreateObject("Scripting.Dictionary")
    Set root("diag") = diag
    diag("blockingProc") = Null
    diag("blockingProcMs") = Null
    diag("clockImpl") = "none_available"
    diag("clockStart") = Null
    diag("clockAtLoopEntry") = Null
    diag("clockAtExit") = Null
    diag("tickDelta") = Null
    diag("iterations") = 0
    diag("errNumber") = 0
    diag("errDescription") = ""
    diag("errProc") = ""
    diag("rule") = "every value here is a REAL reading or Null+<field>UnavailableReason; no constant is reported as a measurement"
    Dim cst As Object
    Set cst = CreateObject("Scripting.Dictionary")
    RpSetInto Owner, diag, "clockSelfTest", cst
    RpSelfTestInto Owner, cst
    RpSetInto Owner, diag, "clocksAtStart", RpClockReadAll(Owner)
    diag("clockVerdict") = "unknown_time_pass_not_proven"
    Dim selfImpl As String
    selfImpl = RpPickClock(Owner, cst)
    Dim timePassProven As Boolean
    timePassProven = False
    If cst.Exists("timePassProven") Then timePassProven = CBool(cst("timePassProven"))
    If timePassProven Then
        If selfImpl = "none_available" Then
            diag("clockVerdict") = "time_passed_no_clock_advanced"
        Else
            diag("clockVerdict") = "clock_advancing"
        End If
    End If
    Dim budgetTrusted As Boolean
    budgetTrusted = (CStr(diag("clockVerdict")) = "clock_advancing")
    diag("budgetClockTrusted") = budgetTrusted
    Dim spanImpl As String
    spanImpl = RpPickAnyClock(Owner)
    diag("spanClockImpl") = spanImpl
    Dim clockUsable As Boolean
    clockUsable = (Len(spanImpl) > 0)
    Owner.m_diagClock = spanImpl
    Set Owner.m_diagSpans = Nothing
    diag("clockUsable") = clockUsable
    Dim spans As Object
    Set spans = CreateObject("Scripting.Dictionary")
    RpSetInto Owner, diag, "waitSpans", spans
    Set Owner.m_diagSpans = spans
    Dim preSpans As Object
    Set preSpans = CreateObject("Scripting.Dictionary")
    RpSetInto Owner, diag, "preLoopSpans", preSpans
    Dim tObserveEntry As Long
    tObserveEntry = WoGetTickCount()
    diag("observeEntryTickMs") = tObserveEntry
    ' WP-22-02: the STA queues new calls while a VB6 program is RUNNING, so the call can be
    ' DISPATCHED only after the program yields/exits (measured: client 20.5 s, add-in 0 ms
    ' inside observe and ONE iteration). This block publishes that fact instead of hiding it.
    Dim disp As Object
    Set disp = CreateObject("Scripting.Dictionary")
    RpSetInto Owner, diag, "dispatch", disp
    disp("programStateAtEntry") = CStr(root("ideStateBeforeObserve"))
    disp("programStateAtEntryDefinition") = "the program state RIGHT WHEN the tool began executing - NOT the state when the client sent the request"
    disp("addInObserveMs") = 0
    disp("toolEntryToObserveMs") = 0
    disp("queuedBeforeToolEntryMs") = Null
    disp("queuedBeforeToolEntryUnavailableReason") = "STA: the add-in cannot observe the interval before its own dispatch (it only starts running after the IDE message pump hands it the call)"
    If Owner.m_callStartAt > 0 Then disp("toolEntryToObserveMs") = tObserveEntry - Owner.m_callStartAt
    ' dispatchedLate is DECIDABLE: for observe:true the call was NOT dispatched while the program
    ' was running when the state seen at entry is anything other than "run".
    disp("dispatchedLate") = (CStr(disp("programStateAtEntry")) <> "run")
    Dim iterations As Long
    iterations = 0
    Dim spA As Double
    Dim spB As Double
    Dim lockEntry As Double
    Dim initSpanStart As Double
    Dim budgetImpl As String
    lockEntry = 0#
    initSpanStart = 0#
    If Len(spanImpl) > 0 Then initSpanStart = RpClockValue(Owner, RpClockRead(Owner, spanImpl))
    If budgetTrusted Then
        budgetImpl = selfImpl
    Else
        budgetImpl = spanImpl
    End If
    diag("budgetClockImpl") = budgetImpl
    If Len(budgetImpl) > 0 Then diag("clockStart") = RpClockValue(Owner, RpClockRead(Owner, budgetImpl))
    If Not budgetTrusted Then
        If CStr(diag("clockVerdict")) = "time_passed_no_clock_advanced" Then
            diag("clockUnavailableReason") = "a measurable span WAS proven but no clock advanced across it (see diag.clockSelfTest.delta and diag.clockEndToStart): the clocks do not advance in this build"
        Else
            diag("clockUnavailableReason") = "time passing could not be proven in this build (Sleep produced no measurable span and the pump wait measured no clock change): the clock verdict is UNKNOWN, NOT dead; see diag.clockSelfTest"
        End If
    End If
    root("programExited") = False
    root("pipeReleasedByReturn") = True
    root("pipeReleaseNote") = "observe RETURNS at the budget or at the done marker, so the normal unwind releases the single-flight gate. It cannot preempt a VB statement that is already executing (single-threaded STA): a returned timeout is a caller-side decision, NOT a killed statement."
    root("stateRule") = "state is a CLOSED set of FIVE values, mutually exclusive and decidable: done-by-marker (ONLY when marker.matched:true - the done_marker file was actually matched), exited (the program ended normally: the run/process is gone), break (stopped at a breakpoint / debug context), error (an unhandled runtime-error dialog is up) and timeout (the observe budget expired while the run had not terminated). done-by-marker IMPLIES marker.matched:true. If a done_marker was REQUESTED but never matched, marker.unmet:true says the run ended without producing the evidence the caller asked for. runningState keeps its own historical vocabulary (exited | break | error | timeout | run) and is a DIFFERENT axis."
    If Len(markerPath) > 0 Then
        Dim mk As Object
        Set mk = CreateObject("Scripting.Dictionary")
        Set root("marker") = mk
        mk("requested") = True
        mk("matched") = False
        mk("unmet") = False
        mk("unmetReason") = ""
    End If
    Dim pre As String
    If Len(spanImpl) > 0 Then spA = RpClockValue(Owner, RpClockRead(Owner, spanImpl))
    pre = RpDialogPrecheck(Owner, root)
    If Len(spanImpl) > 0 Then
        spB = RpClockValue(Owner, RpClockRead(Owner, spanImpl))
        RpSpanAdd Owner, preSpans, "RpDialogPrecheck", spB - spA
    End If
    RpErrNote Owner, diag, "RpDialogPrecheck"
    If Len(pre) > 0 Then
        If Len(markerPath) > 0 Then
            Dim mkp As Object
            Set mkp = root("marker")
            If Not mkp Is Nothing Then
                mkp("unmet") = True
                mkp("unmetReason") = "observe was blocked before the done_marker could be read (reason=" & CStr(root("reason")) & ")"
            End If
        End If
        root("elapsedMs") = RpElapsedMs(Owner, t0)
        RpObserve = JsonStringify(root)
        Exit Function
    End If
    root("observePrecheck") = "no blocking #32770 dialog was up when observe started"
    Dim st As String
    Dim dlgTitle As String
    Dim seenBreak As Boolean
    Dim markerHit As Boolean
    Dim reR As Object
    Dim m As Object
    seenBreak = False
    markerHit = False
    If Len(spanImpl) > 0 Then
        diag("preLoopMs") = RpClockValue(Owner, RpClockRead(Owner, spanImpl)) - initSpanStart
        diag("clockAtLoopEntry") = RpClockValue(Owner, RpClockRead(Owner, spanImpl))
    End If
    If Len(budgetImpl) > 0 Then lockEntry = RpClockValue(Owner, RpClockRead(Owner, budgetImpl))
    Dim timedOut As Boolean
    Dim elapsedSoFar As Long
    timedOut = False
    elapsedSoFar = 0
    Do
        iterations = iterations + 1
        diag("iterations") = iterations
        If iterations = 1 Then
            If clockUsable Then diag("clockAtLoopEntry") = RpClockValue(Owner, RpClockRead(Owner, spanImpl))
        End If
        ' THE BUDGET IS ENFORCED INSIDE THE WAIT LOOP (not only in a loop condition and not only
        ' after the fact): when it expires, state=timeout is assigned HERE, the flag is set and
        ' the loop is left immediately - the program is NOT touched and the IDE state is NOT changed.
        If Len(budgetImpl) > 0 Then
            elapsedSoFar = CLng(RpClockValue(Owner, RpClockRead(Owner, budgetImpl)) - lockEntry)
            If elapsedSoFar < 0 Then elapsedSoFar = 0
        Else
            elapsedSoFar = RpElapsedMs(Owner, t0)
        End If
        If elapsedSoFar >= budgetMs Then
            timedOut = True
            root("timeout") = True
            root("state") = "timeout"
            Exit Do
        End If
        If Len(markerPath) > 0 Then
            If RpMarkerProbe(Owner, markerPath, markerContains, root("marker")) Then
                markerHit = True
                Exit Do
            End If
        End If
        If clockUsable Then spA = RpClockValue(Owner, RpClockRead(Owner, spanImpl))
        Set reR = ReCaptureRuntimeError(Owner, True)
        If clockUsable Then
            spB = RpClockValue(Owner, RpClockRead(Owner, spanImpl))
            RpSpanAdd Owner, spans, "ReCaptureRuntimeError", spB - spA
        End If
        RpErrNote Owner, diag, "ReCaptureRuntimeError"
        If CBool(reR("found")) Then
            root("err") = CStr(reR("dialogTitle"))
            root("runningState") = "error"
            Set root("errorDialog") = McpBpCore.H3DialogBlock(Owner, CStr(reR("dialogTitle")))
            ReMergeRuntimeError Owner, root, reR
            Exit Do
        End If
        If clockUsable Then spA = RpClockValue(Owner, RpClockRead(Owner, spanImpl))
        If McpDb.DbBreakNow(Owner) Then
            seenBreak = True
            root("runningState") = "break"
            Set root("firstBreakpointHit") = McpBpCore.H3PositionObject(Owner)
            Exit Do
        End If
        If clockUsable Then
            spB = RpClockValue(Owner, RpClockRead(Owner, spanImpl))
            RpSpanAdd Owner, spans, "DbBreakNow", spB - spA
        End If
        RpErrNote Owner, diag, "DbBreakNow"
        If clockUsable Then spA = RpClockValue(Owner, RpClockRead(Owner, spanImpl))
        st = Owner.GetIdeState()
        If clockUsable Then
            spB = RpClockValue(Owner, RpClockRead(Owner, spanImpl))
            RpSpanAdd Owner, spans, "GetIdeState", spB - spA
        End If
        RpErrNote Owner, diag, "GetIdeState"
        If st = "design" Then
            root("runningState") = "exited"
            ReInvalidateRuntimeError Owner, "run_ended_normally"
            Exit Do
        End If
        dlgTitle = ""
        If clockUsable Then spA = RpClockValue(Owner, RpClockRead(Owner, spanImpl))
        If WinOps.HasVisibleHostDialog(dlgTitle) Then
            Dim dlgClass As String
            dlgClass = RpClassifyDialogTitle(Owner, dlgTitle)
            Set root("dialogPrecheck") = RpDialogBlockFromTitle(Owner, dlgTitle, dlgClass)
            If dlgClass = "program_error" Then
                root("runningState") = "error"
                root("err") = dlgTitle
                Set root("errorDialog") = McpBpCore.H3DialogBlock(Owner, dlgTitle)
            Else
                root("observedState") = "modal_dialog"
                root("reason") = "observe_blocked_by_ide_dialog"
                root("dialogClass") = dlgClass
                root("isProgramError") = False
                root("retrySafe") = True
                root("failureKind") = "toolError"
                root("howToProceed") = "a #32770 dialog from the IDE (or an unrecognized host dialog) is up; it is NOT a program error. Close or answer it (close_active_dialogs for an answerable box) and observe again. observe never clicks a dialog."
            End If
            Exit Do
        End If
        If clockUsable Then
            spB = RpClockValue(Owner, RpClockRead(Owner, spanImpl))
            RpSpanAdd Owner, spans, "WinOps.HasVisibleHostDialog", spB - spA
        End If
        RpErrNote Owner, diag, "WinOps.HasVisibleHostDialog"
        If clockUsable Then spA = RpClockValue(Owner, RpClockRead(Owner, spanImpl))
        If Util.CanPump() Then DoEvents
        If clockUsable Then
            spB = RpClockValue(Owner, RpClockRead(Owner, spanImpl))
            RpSpanAdd Owner, spans, "DoEvents", spB - spA
        End If
        RpErrNote Owner, diag, "DoEvents"
        If clockUsable Then spA = RpClockValue(Owner, RpClockRead(Owner, spanImpl))
        McpDb.DbBriefWait Owner, 40
        If clockUsable Then
            spB = RpClockValue(Owner, RpClockRead(Owner, spanImpl))
            RpSpanAdd Owner, spans, "DbBriefWait", spB - spA
        End If
        RpErrNote Owner, diag, "DbBriefWait"
    Loop
    If (Not markerHit) And Len(markerPath) > 0 Then
        If RpMarkerProbe(Owner, markerPath, markerContains, root("marker")) Then
            markerHit = True
        End If
    End If
    Dim rs As String
    rs = CStr(root("runningState"))
    If markerHit Then
        If Len(rs) = 0 Then root("runningState") = "run"
    ElseIf Len(rs) = 0 Then
        root("runningState") = "timeout"
    End If
    rs = CStr(root("runningState"))
    If timedOut And Not markerHit Then
        root("state") = "timeout"
        root("runningState") = "timeout"
    ElseIf rs = "error" Then
        root("state") = "error"
    ElseIf rs = "break" Then
        root("state") = "break"
    ElseIf rs = "exited" Then
        root("programExited") = True
        root("state") = "exited"
    ElseIf markerHit Then
        root("state") = "done-by-marker"
    Else
        root("state") = "timeout"
    End If
    If Len(markerPath) > 0 Then
        Set m = root("marker")
        If Not m Is Nothing Then
            If Not CBool(m("matched")) Then
                m("unmet") = True
                If timedOut Then
                    m("unmetReason") = "budget reached before the marker appeared"
                Else
                    m("unmetReason") = "the observation ended with state=" & CStr(root("state")) & " but the done_marker never matched: the run may have finished WITHOUT producing the evidence the caller asked for"
                End If
            End If
        End If
    End If
    root("observedState") = rs
    root("startedPid") = CLng(DbGetCurrentProcessId())
    root("startedPidNote") = "an IDE-run program executes INSIDE the VB6 process, so this is the IDE pid; a separate EXE pid exists only for an out-of-process component"
    If Not seenBreak Then root("firstBreakpointHit") = Null
    If rs <> "run" Then RpRestoreEnv Owner, root
    If ReFindRuntimeErrorDialog(Owner) Is Nothing Then
        root("runningStateNow") = Owner.GetIdeState()
    Else
        root("runningStateNow") = "runtime_error"
        root("runningStateNowSource") = "runtime_error_dialog (the box is still up: the IDE object model was NOT consulted)"
    End If
    If CStr(root("state")) = "timeout" Then
        root("ok") = True
        root("timeout") = True
        root("reason") = "observe_budget_reached"
        root("retrySafe") = True
        root("suggestedAction") = "observe_again"
        root("message") = "the observe budget (" & CStr(budgetMs) & " ms) was reached and the call RETURNED; the program may still be running. This is a retryable end of the observation WINDOW - NOT a failure of the program and NOT the goal being reached. Observe again (a new call) or inspect with ide_state / read_active_dialogs / last_error."
    Else
        root("ok") = True
        If CBool(root("programExited")) Then
            root("message") = "the program ended normally (state=exited); the observation is complete."
            Set m = root("marker")
            If Not m Is Nothing Then
                If CBool(m("unmet")) Then
                    root("message") = "the program ended normally (state=exited) BUT the done_marker never matched (marker.unmet:true): the run finished without producing the evidence you asked for - do NOT read this as the goal being reached."
                End If
            End If
        ElseIf markerHit Then
            Set m = root("marker")
            If Not m Is Nothing Then
                root("message") = "the done_marker matched (bytesRead=" & CStr(m("bytesRead")) & ", matchedBy=" & CStr(m("matchedBy")) & "); observe returned immediately and the gate is released."
            Else
                root("message") = "the done_marker matched."
            End If
        ElseIf CStr(root("state")) = "break" Then
            root("message") = "the program stopped at a breakpoint. Resume/step/end it with control_execution; observe only WATCHES."
        Else
            root("message") = "a VB6 runtime-error dialog is up; the receipt carries errorDialog and the captured error text. last_error can re-read it."
        End If
    End If
    diag("iterations") = iterations
    diag("preLoopTopProc") = RpTopSpan(Owner, preSpans)
    If Len(spanImpl) > 0 Then
        diag("clockAtExit") = RpClockValue(Owner, RpClockRead(Owner, spanImpl))
        If Not IsNull(diag("clockAtLoopEntry")) Then
            diag("tickDelta") = CDbl(diag("clockAtExit")) - CDbl(diag("clockAtLoopEntry"))
        End If
        If diag.Exists("clocksAtStart") Then
            RpSetInto Owner, diag, "clockEndToStart", RpDeltaAll(Owner, diag("clocksAtStart"), RpClockReadAll(Owner))
        End If
    End If
    diag("addInObserveMs") = WoGetTickCount() - tObserveEntry
    If Not disp Is Nothing Then
        disp("addInObserveMs") = WoGetTickCount() - tObserveEntry
    End If
    If Owner.m_callStartAt > 0 Then
        diag("toolEntryToObserveMs") = tObserveEntry - Owner.m_callStartAt
        diag("toolTotalMsAtObserveReturn") = WoGetTickCount() - Owner.m_callStartAt
    End If
    diag("blockingProc") = RpTopSpan(Owner, spans)
    If Len(CStr(diag("blockingProc"))) = 0 Then
        diag("blockingProcUnavailableReason") = "no wait spans were measured (no clock read succeeded; see diag.clockSelfTest and diag.clockEndToStart)"
    ElseIf spans.Exists(CStr(diag("blockingProc"))) Then
        diag("blockingProcMs") = CDbl(spans(CStr(diag("blockingProc")))("ms"))
        If CDbl(diag("blockingProcMs")) <= 0# Then
            diag("blockingProc") = Null
            diag("blockingProcMs") = Null
            diag("blockingProcUnavailableReason") = "every measured span was 0 ms: the add-in-side wall time did NOT sit inside the instrumented calls (see diag.addInObserveMs / diag.toolEntryToObserveMs) - reporting the first span would be a lie"
        End If
    End If
    Dim clockSuspect As Boolean
    clockSuspect = False
    If Not budgetTrusted Then clockSuspect = True
    root("clockSuspect") = clockSuspect
    If clockSuspect Then
        root("clockSource") = "none_available"
        root("message") = "clockSuspect:true - verdict=" & CStr(diag("clockVerdict")) & "; see diag.clockSelfTest / diag.clockEndToStart. The observe budget may not be enforceable in this build. " & CStr(root("message"))
    End If
    If Not disp Is Nothing Then
        If CBool(disp("dispatchedLate")) Then
            root("message") = "DISPATCHED LATE: when this call began executing the program was NOT running (" & CStr(disp("programStateAtEntry")) & "). The STA cannot prove that any real-time observation happened before this call was served; a marker matched here reflects the state AT DISPATCH TIME, not a live observation. " & CStr(root("message"))
        End If
    End If
    root("elapsedMs") = RpElapsedMs(Owner, t0)
    Owner.RunDiskAttachPhase2 root
    RpObserve = JsonStringify(root)
End Function

Public Function RpPickAnyClock(Owner As CodeTools) As String
    RpPickAnyClock = ""
    Dim order As Object
    Dim k As Variant
    Set order = New Collection
    order.Add "codetools_gettickcount"
    order.Add "util_tickms"
    order.Add "vb_timer"
    order.Add "calendar_now"
    For Each k In order
        Dim r As Object
        Set r = RpClockRead(Owner, CStr(k))
        If CBool(r("ok")) Then
            RpPickAnyClock = CStr(k)
            Exit Function
        End If
    Next
End Function

Public Function RpPickClock(Owner As CodeTools, ByRef st As Object) As String
    RpPickClock = "none_available"
    If st Is Nothing Then Exit Function
    If Not st.Exists("delta") Then Exit Function
    Dim deltas As Object
    Set deltas = st("delta")
    If deltas Is Nothing Then Exit Function
    Dim order As Object
    Dim k As Variant
    Set order = New Collection
    order.Add "codetools_gettickcount"
    order.Add "util_tickms"
    order.Add "vb_timer"
    order.Add "calendar_now"
    For Each k In order
        If deltas.Exists(CStr(k)) Then
            If Not IsNull(deltas(CStr(k))) Then
                If CDbl(deltas(CStr(k))) > 0# Then
                    RpPickClock = CStr(k)
                    Exit Function
                End If
            End If
        End If
    Next
End Function

Public Function RpReadFilePrefixBytes(Owner As CodeTools, ByVal p As String, ByVal maxBytes As Long) As String
    On Error Resume Next
    RpReadFilePrefixBytes = ""
    If Len(p) = 0 Then Exit Function
    Dim f As Integer
    Dim n As Long
    Dim buf As String
    f = FreeFile
    Open p For Binary Access Read As #f
    If Err.Number <> 0 Then
        Err.Clear
        Exit Function
    End If
    n = LOF(f)
    If n > maxBytes Then n = maxBytes
    If n > 0 Then
        buf = String$(n, vbNullChar)
        Get #f, 1, buf
        If Err.Number = 0 Then RpReadFilePrefixBytes = buf
        Err.Clear
    End If
    Close #f
    Err.Clear
End Function

Public Sub RpRestoreEnv(Owner As CodeTools, ByRef root As Object)
    On Error Resume Next
    root("envRestored") = False
    If Owner.m_envRestore Is Nothing Then
        root("envRestored") = True
        Exit Sub
    End If
    Dim i As Long
    For i = 1 To Owner.m_envRestore.Count
        Dim rec As Object
        Set rec = Owner.m_envRestore(i)
        Dim nm As String
        nm = CStr(rec("name"))
        If CBool(rec("previousValueWasPresent")) Then
            SetEnvironmentVariableW StrPtr(nm), StrPtr(CStr(rec("previousValue")))
        Else
            SetEnvironmentVariableW StrPtr(nm), 0
        End If
    Next
    Set Owner.m_envRestore = Nothing
    root("envRestored") = True
    root("envRestoreNote") = "each variable was restored to its previous value (or DELETED when it did not exist before)"
End Sub

Public Sub RpSelfTestInto(Owner As CodeTools, ByRef cst As Object)
    On Error GoTo CST_ERR
    cst("sleepRequestedMs") = 50
    cst("sleepApiErr") = 0
    cst("sleepIneffective") = False
    cst("timePassProven") = False
    cst("nowDeltaMs") = Null
    RpSetInto Owner, cst, "start", RpClockReadAll(Owner)
    Dim beforeNow As Double
    beforeNow = RpClockValue(Owner, RpClockRead(Owner, "calendar_now"))
    Sleep 50
    Dim afterNow As Double
    afterNow = RpClockValue(Owner, RpClockRead(Owner, "calendar_now"))
    If beforeNow > 0# And afterNow > 0# Then
        Dim nd As Double
        nd = afterNow - beforeNow
        If nd < 0# Then nd = nd + 86400000#
        cst("nowDeltaMs") = nd
        If nd > 0# Then cst("timePassProven") = True
    Else
        cst("nowUnavailableReason") = "the calendar clock could not be read before/after the sleep"
    End If
    If Not CBool(cst("timePassProven")) Then
        cst("sleepIneffective") = True
        ' Sleep produced no measurable span; try a bounded pump wait that does NOT trust Sleep.
        Dim bw As Object
        Set bw = CreateObject("Scripting.Dictionary")
        RpBusyWaitForSpan Owner, bw
        RpSetInto Owner, cst, "busyWait", bw
        If Not IsNull(bw("deltaMs")) Then
            If CDbl(bw("deltaMs")) > 0# Then cst("timePassProven") = True
        End If
    End If
    RpSetInto Owner, cst, "end", RpClockReadAll(Owner)
    RpSetInto Owner, cst, "delta", RpDeltaAll(Owner, cst("start"), cst("end"))
    Exit Sub
CST_ERR:
    cst("selfTestErrNumber") = Err.Number
    cst("selfTestErrDescription") = Err.Description
    cst("selfTestErrProc") = "RpSelfTestInto"
    Err.Clear
    On Error Resume Next
    If Not cst.Exists("end") Then RpSetInto Owner, cst, "end", RpClockReadAll(Owner)
    If Not cst.Exists("delta") Then RpSetInto Owner, cst, "delta", RpDeltaAll(Owner, cst("start"), cst("end"))
    Err.Clear
End Sub

Public Sub RpSetInto(Owner As CodeTools, ByRef target As Object, ByVal key As String, ByRef value As Object)
    On Error Resume Next
    If target Is Nothing Then Exit Sub
    Set target(key) = value
    If Err.Number <> 0 Then
        target(key & "ErrorNumber") = Err.Number
        target(key & "ErrorDescription") = Err.Description
        Err.Clear
    End If
    Err.Clear
End Sub

Public Sub RpSpanAdd(Owner As CodeTools, ByRef spans As Object, ByVal procName As String, ByVal ms As Double)
    On Error Resume Next
    If spans Is Nothing Then Exit Sub
    Dim d As Object
    If spans.Exists(procName) Then
        Set d = spans(procName)
    Else
        Set d = CreateObject("Scripting.Dictionary")
        d("proc") = procName
        d("ms") = 0#
        d("calls") = 0
        Set spans(procName) = d
    End If
    d("ms") = CDbl(d("ms")) + CDbl(ms)
    d("calls") = CLng(d("calls")) + 1
    Err.Clear
End Sub

Public Function RpTopSpan(Owner As CodeTools, ByRef spans As Object) As String
    RpTopSpan = ""
    On Error Resume Next
    If spans Is Nothing Then Exit Function
    Dim k As Variant
    Dim d As Object
    Dim bestMs As Double
    bestMs = -1#
    For Each k In spans.Keys
        Set d = spans(k)
        If CDbl(d("ms")) > bestMs Then
            bestMs = CDbl(d("ms"))
            RpTopSpan = CStr(k)
        End If
    Next
    Err.Clear
End Function

Public Function RpUtf8Bytes(Owner As CodeTools, ByVal s As String) As String
    On Error Resume Next
    Dim out As String
    Dim i As Long
    Dim cp As Long
    out = ""
    For i = 1 To Len(s)
        cp = AscW(Mid$(s, i, 1))
        If cp < 0 Then cp = cp + 65536
        If cp < 128 Then
            out = out & Chr$(cp)
        ElseIf cp < 2048 Then
            out = out & Chr$(&HC0 + (cp \ 64)) & Chr$(&H80 + (cp And &H3F))
        Else
            out = out & Chr$(&HE0 + (cp \ 4096)) & Chr$(&H80 + ((cp \ 64) And &H3F)) & Chr$(&H80 + (cp And &H3F))
        End If
    Next
    RpUtf8Bytes = out
End Function

