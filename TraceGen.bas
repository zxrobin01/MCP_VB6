Attribute VB_Name = "TraceGen"
Option Explicit

' ===========================================================================
' TASK-P1-129: modMcpTrace is now MATERIALIZED BY THE ADD-IN.
' The guard inject_trace emits is 'If McpTraceOn() Then ... End If' and it has
' to resolve at COMPILE time, so the target project must contain modMcpTrace.
' Instead of shipping a file the user has to import by hand, the module source
' lives below as VB6 string data and inject_trace adds the component itself.
'
' The comment language comes from the IDE's OWN version resource:
'   GetModuleFileNameW(0) -> VB6.EXE (the add-in runs INSIDE it)
'   GetFileVersionInfoW + VerQueryValueW(\VarFileInfo\Translation) -> LANGID
'   LANG_CHINESE (primary language 0x04) -> Chinese header, else English.
' A FAILED detection ALWAYS means English: English is readable on every code
' page while Chinese becomes mojibake on a non-Chinese one. The IDE's own menu
' captions are read as a cross-check; an explicitly ENGLISH menu on a Chinese
' resource also downgrades to English.
' ===========================================================================

Private Declare Function GetModuleFileNameW Lib "kernel32" (ByVal hModule As Long, ByVal lpFileName As Long, ByVal nSize As Long) As Long
Private Declare Function GetFileVersionInfoSizeW Lib "version.dll" (ByVal lptstrFilename As Long, ByRef lpdwHandle As Long) As Long
Private Declare Function GetFileVersionInfoW Lib "version.dll" (ByVal lptstrFilename As Long, ByVal dwHandle As Long, ByVal dwLen As Long, ByRef lpData As Any) As Long
Private Declare Function VerQueryValueW Lib "version.dll" (ByRef pBlock As Any, ByVal lpSubBlock As Long, ByRef lplpBuffer As Long, ByRef puLen As Long) As Long
Private Declare Sub TgCopy Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByVal Source As Long, ByVal Length As Long)
Private Declare Function TgFindWindowW Lib "user32" Alias "FindWindowW" (ByVal lpClassName As Long, ByVal lpWindowName As Long) As Long
Private Declare Function GetMenu Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function GetMenuStringW Lib "user32" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As Long, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Private Declare Function TgGetWindowTextW Lib "user32" Alias "GetWindowTextW" (ByVal hWnd As Long, ByVal lpString As Long, ByVal nMaxCount As Long) As Long

Private Const TG_MARKER As String = "VB6MCP-AUTOGEN modMcpTrace"
Private Const TG_MF_BYPOSITION As Long = &H400
Private Const TG_VB6_WNDCLASS As String = "wndclass_desked_gsk"

' ---------------- module materialization ----------------

Public Function GenerateTraceModule(ByVal templatePath As String, ByVal outPath As String) As Boolean
    Dim b() As Byte
    On Error GoTo fail
    ReadAllBytesTo templatePath, b
    If ByteCount(b) <= 0 Then
        GenerateTraceModule = False
        Exit Function
    End If
    WriteAllBytesTo outPath, b
    GenerateTraceModule = True
    Exit Function
fail:
    GenerateTraceModule = False
End Function

' ---------------- trace ----------------

Public Function InjectTrace(ByVal text As String, ByVal procName As String, ByVal exprsCsv As String, ByVal id As String, ByRef newText As String, ByRef errMsg As String) As Boolean
    Dim lines() As String, di As Long, cnt As Long, body As String
    lines = Split(AsLf(text), vbLf)
    FindProc lines, procName, di, cnt
    If cnt = 0 Then
        errMsg = "procedure not found: " & procName
        InjectTrace = False
        Exit Function
    End If
    If cnt > 1 Then
        errMsg = "procedure not unique (" & CStr(cnt) & "): " & procName
        InjectTrace = False
        Exit Function
    End If
    body = BuildCall(id, exprsCsv)
    Dim ins As Long, ind As String
    ins = BodyInsertIndex(lines, di, ind)
    ' TASK-P1-028: the guard is a RUNTIME check on purpose. "#If MCP_TRACE Then" is a
    ' COMPILE-TIME condition: the VBIDE object model does not expose CondComp, our
    ' project_meta tool edits the .vbp TEXT, so an already-LOADED project keeps its old
    ' in-memory CondComp and the guarded block is compiled out - the trace silently
    ' produces nothing (measured by S/P4: the same block works when it is present at
    ' project LOAD time). "If McpTraceOn() Then" has no load-order dependency; the cost
    ' is that the call is always compiled (one boolean check per traced line).
    newText = InsertMarkedBlock(lines, ins, "@mcp-trace ", id, "If McpTraceOn() Then", "End If", body, ind)
    errMsg = ""
    InjectTrace = True
End Function

Public Function RemoveTrace(ByVal text As String, ByVal id As String, ByRef newText As String, ByRef errMsg As String) As Boolean
    RemoveTrace = RemoveBlockByMarkers(text, "@mcp-trace ", id, newText, errMsg)
End Function

' =====================================================================
' TASK-P1-032: lexical symbol visibility at the injection point.
' A MODAL compile-error dialog is much worse than a rejection (it blocks
' end/remove_trace too), so inject_trace must never emit code that cannot
' compile. Everything here is decided from the TEXT - VB6 is never asked to
' compile, because that is exactly the path that raises the modal dialog.
' =====================================================================

' Leading identifier token of an expression fragment; "" when the fragment
' does not start with a letter/underscore (a literal, '(' , '"', a digit ...).
Public Function LeadingIdent(ByVal s As String) As String
    ' NB: digits are part of an identifier AFTER the first character (Form1, n2).
    ' Truncating at the first digit would turn 'Form1' into 'Form' and could both
    ' false-reject a legal name and false-ACCEPT an undeclared one (measured).
    Dim t As String, i As Long, ch As String, r As String
    t = LTrim$(StripCommentLite(s))
    For i = 1 To Len(t)
        ch = Mid$(t, i, 1)
        If IsIdentChar(ch, Len(r) > 0) Then
            r = r & ch
        ElseIf Len(r) > 0 Then
            Exit For
        ElseIf ch = " " Or ch = vbTab Then
            ' leading whitespace: keep looking
        Else
            Exit Function
        End If
    Next
    LeadingIdent = r
End Function

' Names DECLARED on one source line: Dim/Static/Const/ReDim/Type/Enum/Declare/
' Private/Public/Global/WithEvents, Sub/Function/Property headers, and For /
' For Each loop counters (an implicit declaration on that line).
Public Function DeclNamesOnLine(ByVal ln As String) As String
    Dim t As String, w As String
    t = Trim$(StripCommentLite(ln))
    If Len(t) = 0 Then Exit Function
    w = LCase$(FirstWordOf(t))
    Select Case w
        Case "dim", "static", "const", "redim", "private", "public", "global", "type", "enum", "declare"
            DeclNamesOnLine = NamesAfterKeyword(t, w)
        Case "sub", "function", "property"
            DeclNamesOnLine = FirstIdentOf(t)
        Case "for"
            DeclNamesOnLine = FirstIdentOf(Trim$(Mid$(t, 4)))
    End Select
End Function

' Parameters of the procedure whose declaration starts at lines(di) (continuation aware).
Public Function ProcParamNames(ByRef lines() As String, ByVal di As Long) As String
    Dim j As Long, decl As String
    j = di
    Do While j <= UBound(lines)
        MaybePump
        decl = decl & " " & lines(j)
        If Not EndsWithContinuation(lines(j)) Then Exit Do
        j = j + 1
        If j > di + 20 Then Exit Do
    Loop
    Dim p1 As Long, p2 As Long
    p1 = InStr(decl, "(")
    If p1 = 0 Then Exit Function
    p2 = InStrRev(decl, ")")
    If p2 <= p1 Then Exit Function
    Dim inner As String, parts() As String, i As Long, out As String, nm As String
    inner = Mid$(decl, p1 + 1, p2 - p1 - 1)
    SplitTop inner, ",", parts
    For i = 0 To UBound(parts)
        MaybePump
        nm = FirstIdentOf(parts(i))
        If Len(nm) > 0 Then out = out & "," & nm
    Next
    If Len(out) > 0 Then ProcParamNames = Mid$(out, 2)
End Function

' Every name that is legal to reference at the injection point of procName:
'   (a) the procedure's parameters,
'   (b) its locals/Statics/Consts declared at or ABOVE the insertion line
'       (the block is inserted after 0-based line insIndex, so lines di..insIndex
'        are textually above it - that is exactly what the compiler sees),
'   (c) module-level declarations (lines that are not inside any procedure body).
Public Function VisibleSymbolsAt(ByVal text As String, ByVal procName As String, _
                                 ByRef insIndex As Long, ByRef afterIdeLine As Long, ByRef errMsg As String) As String
    Dim lines() As String, di As Long, cnt As Long
    lines = Split(AsLf(text), vbLf)
    FindProc lines, procName, di, cnt
    If cnt = 0 Then
        errMsg = "procedure not found: " & procName
        Exit Function
    End If
    If cnt > 1 Then
        errMsg = "procedure not unique (" & CStr(cnt) & "): " & procName
        Exit Function
    End If
    Dim ind As String
    insIndex = BodyInsertIndex(lines, di, ind)
    If insIndex < di Then insIndex = di
    afterIdeLine = insIndex + 1
    ' TASK-P1-049: was "names = names & ..." ONCE PER MODULE LINE (quadratic on
    ' the module being traced). Collect the fragments and Join ONCE; the old code
    ' put a "," BEFORE every fragment, which is exactly Join(frags, ",") - empty
    ' fragments (DeclNamesOnLine can return "") are preserved, so ",," still
    ' appears exactly as before.
    Dim i As Long, k As Long
    Dim ub As Long
    ub = UBound(lines)
    Dim frags() As String
    ReDim frags(0 To 2 * (ub + 1) + 2)
    k = 0
    frags(k) = ProcParamNames(lines, di)
    k = k + 1
    For i = di To insIndex
        MaybePump
        frags(k) = DeclNamesOnLine(lines(i))
        k = k + 1
    Next
    On Error Resume Next
    Vb6Lint.Lint text, "bas"
    Err.Clear
    On Error GoTo 0
    For i = 0 To ub
        MaybePump
        If Len(Vb6Lint.ProcBodyOwner(i + 1)) = 0 Then
            frags(k) = DeclNamesOnLine(lines(i))
            k = k + 1
        End If
    Next
    Dim names As String
    If k > 0 Then
        ReDim Preserve frags(0 To k - 1)
        names = Join(frags, ",")
    Else
        names = ""
    End If
    VisibleSymbolsAt = DedupCsv(names)
End Function

' ---------------- lexical helpers ----------------

Private Function StripCommentLite(ByVal s As String) As String
    Dim i As Long
    i = InStr(s, "'")
    If i > 0 Then StripCommentLite = Left$(s, i - 1) Else StripCommentLite = s
End Function

Private Function IsAlphaChar(ByVal ch As String) As Boolean
    If ch >= "a" And ch <= "z" Then IsAlphaChar = True
    If ch >= "A" And ch <= "Z" Then IsAlphaChar = True
    If ch = "_" Then IsAlphaChar = True
End Function

Private Function IsIdentChar(ByVal ch As String, ByVal started As Boolean) As Boolean
    If IsAlphaChar(ch) Then IsIdentChar = True: Exit Function
    If started Then
        If ch >= "0" And ch <= "9" Then IsIdentChar = True
    End If
End Function

Private Function FirstWordOf(ByVal t As String) As String
    Dim i As Long, ch As String, r As String
    For i = 1 To Len(t)
        ch = Mid$(t, i, 1)
        If IsIdentChar(ch, Len(r) > 0) Then r = r & ch Else Exit For
    Next
    FirstWordOf = r
End Function

' tokens that are syntax rather than a declared name
Private Function IsSkipWord(ByVal w As String) As Boolean
    Select Case LCase$(w)
        Case "each", "get", "let", "set", "function", "sub", "property", "declare", _
             "byval", "byref", "optional", "paramarray", "lib", "alias", "withevents", _
             "as", "to", "in", "step", "new", "public", "private", "global", "dim", "static", "const"
            IsSkipWord = True
    End Select
End Function

Private Function FirstIdentOf(ByVal chunk As String) As String
    Dim t As String, i As Long, ch As String, r As String, guard As Long
    t = Trim$(chunk)
    guard = 0
    Do While Len(t) > 0 And guard < 6
        guard = guard + 1
        r = ""
        For i = 1 To Len(t)
            MaybePump
            ch = Mid$(t, i, 1)
            If IsIdentChar(ch, Len(r) > 0) Then r = r & ch Else Exit For
        Next
        If Len(r) = 0 Then Exit Function
        If IsSkipWord(r) Then
            t = Trim$(Mid$(t, Len(r) + 1))
        Else
            FirstIdentOf = r
            Exit Function
        End If
    Loop
End Function

' all declared names on a declaration line, after the leading keyword
Private Function NamesAfterKeyword(ByVal t As String, ByVal kw As String) As String
    Dim rest As String, parts() As String, i As Long, out As String, nm As String
    rest = Trim$(Mid$(t, Len(kw) + 1))
    SplitTop rest, ",", parts
    For i = 0 To UBound(parts)
        MaybePump
        nm = FirstIdentOf(parts(i))
        If Len(nm) > 0 Then out = out & "," & nm
    Next
    If Len(out) > 0 Then NamesAfterKeyword = Mid$(out, 2)
End Function

' split on a separator that is not inside parentheses and not inside a string.
' VB6 cannot RETURN an array from a function (measured pitfall), hence ByRef.
Private Sub SplitTop(ByVal s As String, ByVal sep As String, ByRef out() As String)
    ' NB: the flag must NOT be called inStr - VB6 is case-insensitive and would
    ' shadow the builtin InStr, giving 'syntax error' at compile time (measured).
    Dim n As Long, i As Long, ch As String, depth As Long, inStrLit As Boolean
    Dim cur As String
    ReDim out(0 To 0)
    depth = 0
    inStrLit = False
    For i = 1 To Len(s)
        MaybePump
        ch = Mid$(s, i, 1)
        If ch = Chr$(34) Then inStrLit = Not inStrLit
        If Not inStrLit Then
            If ch = "(" Then depth = depth + 1
            If ch = ")" Then If depth > 0 Then depth = depth - 1
        End If
        If ch = sep And depth = 0 And Not inStrLit Then
            out(n) = cur
            n = n + 1
            ReDim Preserve out(0 To n)
            cur = ""
        Else
            cur = cur & ch
        End If
    Next
    out(n) = cur
End Sub

Private Function DedupCsv(ByVal s As String) As String
    Dim d As Object, parts() As String, i As Long, out As String, k As String
    Set d = CreateObject("Scripting.Dictionary")
    parts = Split(s, ",")
    For i = 0 To UBound(parts)
        MaybePump
        k = Trim$(parts(i))
        If Len(k) > 0 Then
            If Not d.Exists(LCase$(k)) Then
                d(LCase$(k)) = k
                out = out & "," & k
            End If
        End If
    Next
    If Len(out) > 0 Then DedupCsv = Mid$(out, 2)
End Function

' ---------------- fault injection ----------------

Public Function InjectFault(ByVal text As String, ByVal procName As String, ByVal mode As String, ByVal value As String, ByVal id As String, ByRef newText As String, ByRef errMsg As String) As Boolean
    Dim lines() As String, di As Long, cnt As Long, body As String
    lines = Split(AsLf(text), vbLf)
    FindProc lines, procName, di, cnt
    If cnt = 0 Then
        errMsg = "procedure not found: " & procName
        InjectFault = False
        Exit Function
    End If
    If cnt > 1 Then
        errMsg = "procedure not unique (" & CStr(cnt) & "): " & procName
        InjectFault = False
        Exit Function
    End If
    body = FaultBody(procName, lines(di), mode, value, id, errMsg)
    If Len(errMsg) > 0 Then
        InjectFault = False
        Exit Function
    End If
    Dim ins As Long, ind As String
    ins = BodyInsertIndex(lines, di, ind)
    newText = InsertMarkedBlock(lines, ins, "@mcp-fault ", id, "#If MCP_FAULT Then", "#End If", body, ind)
    errMsg = ""
    InjectFault = True
End Function

Public Function RemoveFault(ByVal text As String, ByVal id As String, ByRef newText As String, ByRef errMsg As String) As Boolean
    RemoveFault = RemoveBlockByMarkers(text, "@mcp-fault ", id, newText, errMsg)
End Function

' ---------------- shared helpers ----------------

Private Function BuildCall(ByVal id As String, ByVal exprsCsv As String) As String
    Dim e() As String, i As Long, s As String, nm As String
    s = "McpTrace " & Chr$(34) & id & Chr$(34)
    e = Split(exprsCsv, ",")
    For i = 0 To UBound(e)
        MaybePump
        nm = Trim$(e(i))
        If Len(nm) > 0 Then s = s & ", " & Chr$(34) & nm & Chr$(34) & ", " & nm
    Next
    BuildCall = s
End Function

Private Function FaultBody(ByVal procName As String, ByVal declLine As String, ByVal mode As String, ByVal value As String, ByVal id As String, ByRef errMsg As String) As String
    Dim k As String, ms As String, stmt As String
    k = ProcKindOf(declLine)
    If mode = "raise" Then
        FaultBody = "Err.Raise vbObjectError + 9999, " & Chr$(34) & "MCP" & Chr$(34) & ", " & Chr$(34) & "injected fault: " & id & Chr$(34)
        Exit Function
    End If
    If mode = "delay" Then
        ms = Trim$(value)
        If Len(ms) = 0 Then ms = "200"
        FaultBody = "McpFaultDelay " & ms
        Exit Function
    End If
    If mode = "return" Then
        If k = "sub" Then
            FaultBody = "Exit Sub"
            Exit Function
        End If
        If k = "function" Or k = "property get" Then
            stmt = ""
            If Len(Trim$(value)) > 0 Then stmt = procName & " = " & Trim$(value) & ": "
            If k = "function" Then
                FaultBody = stmt & "Exit Function"
            Else
                FaultBody = stmt & "Exit Property"
            End If
            Exit Function
        End If
        errMsg = "return mode not supported for kind: " & k
        Exit Function
    End If
    errMsg = "unknown fault mode: " & mode
End Function

Private Function ProcKindOf(ByVal line As String) As String
    Dim s As String
    s = StripMods2(LCase$(Trim$(line)))
    If Left$(s, 4) = "sub " Then ProcKindOf = "sub" : Exit Function
    If Left$(s, 9) = "function " Then ProcKindOf = "function" : Exit Function
    If Left$(s, 9) = "property " Then ProcKindOf = "property " & TokN(s, 2) : Exit Function
End Function

Private Sub FindProc(ByRef lines() As String, ByVal procName As String, ByRef di As Long, ByRef cnt As Long)
    Dim i As Long
    di = -1
    cnt = 0
    For i = 0 To UBound(lines)
        MaybePump
        If IsProcDeclOf(lines(i), procName) Then
            cnt = cnt + 1
            di = i
        End If
    Next
End Sub

Private Function InsertMarkedBlock(ByRef lines() As String, ByVal di As Long, ByVal prefix As String, ByVal id As String, ByVal gOpen As String, ByVal gClose As String, ByVal body As String, ByVal indent As String) As String
    Dim blk(0 To 4) As String, out() As String
    Dim i As Long, j As Long, n As Long
    ' P0-4: indentation now comes from the target line instead of a hardcoded 4
    ' spaces, and the markers/guards are indented with it (RemoveBlockByMarkers
    ' compares Trim$() so that is still safe).
    blk(0) = indent & "' " & prefix & id & " begin"
    blk(1) = indent & gOpen
    blk(2) = indent & body
    blk(3) = indent & gClose
    blk(4) = indent & "' " & prefix & id & " end"
    n = (UBound(lines) + 1) + 5
    ReDim out(0 To n - 1)
    For i = 0 To di
        MaybePump
        out(i) = lines(i)
    Next
    For j = 0 To 4
        out(di + 1 + j) = blk(j)
    Next
    For i = di + 1 To UBound(lines)
        MaybePump
        out(i + 5) = lines(i)
    Next
    InsertMarkedBlock = Join(out, vbCrLf)
End Function

Private Function RemoveBlockByMarkers(ByVal text As String, ByVal prefix As String, ByVal id As String, ByRef newText As String, ByRef errMsg As String) As Boolean
    Dim lines() As String, out() As String
    Dim i As Long, k As Long, n As Long, bi As Long, ei As Long, cb As Long, ce As Long
    Dim mb As String, mkEnd As String, nl As String
    lines = Split(AsLf(text), vbLf)
    mb = LCase$("' " & prefix & id & " begin")
    mkEnd = LCase$("' " & prefix & id & " end")
    bi = -1
    ei = -1
    cb = 0
    ce = 0
    For i = 0 To UBound(lines)
        MaybePump
        nl = LCase$(Trim$(lines(i)))
        If nl = mb Then
            cb = cb + 1
            bi = i
        End If
        If nl = mkEnd Then
            ce = ce + 1
            ei = i
        End If
    Next
    If cb = 0 Or bi < 0 Then
        errMsg = "marker begin not found: " & prefix & id
        RemoveBlockByMarkers = False
        Exit Function
    End If
    If ce = 0 Or ei < 0 Then
        errMsg = "marker end not found: " & prefix & id
        RemoveBlockByMarkers = False
        Exit Function
    End If
    If cb > 1 Or ce > 1 Then
        errMsg = "markers not unique: " & prefix & id
        RemoveBlockByMarkers = False
        Exit Function
    End If
    If ei <= bi Then
        errMsg = "markers out of order: " & prefix & id
        RemoveBlockByMarkers = False
        Exit Function
    End If
    n = (UBound(lines) + 1) - (ei - bi + 1)
    If n <= 0 Then
        newText = ""
        errMsg = ""
        RemoveBlockByMarkers = True
        Exit Function
    End If
    ReDim out(0 To n - 1)
    k = 0
    For i = 0 To UBound(lines)
        MaybePump
        If i < bi Or i > ei Then
            out(k) = lines(i)
            k = k + 1
        End If
    Next
    newText = Join(out, vbCrLf)
    errMsg = ""
    RemoveBlockByMarkers = True
End Function

Private Function StripMods2(ByVal s As String) As String
    Dim t As String, again As Boolean
    t = s
    Do
        again = False
        If Left$(t, 7) = "public " Then t = Mid$(t, 8) : again = True
        If (Not again) And Left$(t, 8) = "private " Then t = Mid$(t, 9) : again = True
        If (Not again) And Left$(t, 7) = "friend " Then t = Mid$(t, 8) : again = True
        If (Not again) And Left$(t, 7) = "static " Then t = Mid$(t, 8) : again = True
        If Not again Then Exit Do
    Loop
    StripMods2 = t
End Function

Private Function TokN(ByVal s As String, ByVal n As Long) As String
    Dim parts() As String, r As String, p As Long
    parts = Split(Trim$(s), " ")
    If n - 1 > UBound(parts) Then
        TokN = ""
        Exit Function
    End If
    r = parts(n - 1)
    p = InStr(r, "(")
    If p > 0 Then r = Left$(r, p - 1)
    TokN = r
End Function

Private Function IsProcDeclOf(ByVal line As String, ByVal procName As String) As Boolean
    Dim s As String, p As String
    s = StripMods2(LCase$(Trim$(line)))
    p = LCase$(Trim$(procName))
    If Left$(s, 4) = "sub " Then IsProcDeclOf = (TokN(s, 2) = p) : Exit Function
    If Left$(s, 9) = "function " Then IsProcDeclOf = (TokN(s, 2) = p) : Exit Function
    If Left$(s, 9) = "property " Then IsProcDeclOf = (TokN(s, 3) = p) : Exit Function
End Function

' ---------------- breakpoints (injected Debug.Assert False) ----------------

Public Function SetBreakpoint(ByVal text As String, ByVal anchor As String, ByVal id As String, ByRef newText As String, ByRef errMsg As String) As Boolean
    Dim lines() As String, blk(0 To 2) As String, out() As String
    Dim i As Long, j As Long, n As Long, idx As Long, cnt As Long
    lines = Split(AsLf(text), vbLf)
    idx = -1
    cnt = 0
    For i = 0 To UBound(lines)
        MaybePump
        If LCase$(Trim$(lines(i))) = LCase$(Trim$(anchor)) Then
            cnt = cnt + 1
            idx = i
        End If
    Next
    If cnt = 0 Then
        errMsg = "anchor not found"
        SetBreakpoint = False
        Exit Function
    End If
    If cnt > 1 Then
        errMsg = "anchor not unique (" & CStr(cnt) & ")"
        SetBreakpoint = False
        Exit Function
    End If
    Dim ind As String
    ind = LeadingSpace(lines(idx))
    blk(0) = ind & "' @mcp-break " & id & " begin"
    blk(1) = ind & "Debug.Assert False"
    blk(2) = ind & "' @mcp-break " & id & " end"
    n = (UBound(lines) + 1) + 3
    ReDim out(0 To n - 1)
    For i = 0 To idx - 1
        MaybePump
        out(i) = lines(i)
    Next
    For j = 0 To 2
        out(idx + j) = blk(j)
    Next
    For i = idx To UBound(lines)
        MaybePump
        out(i + 3) = lines(i)
    Next
    newText = Join(out, vbCrLf)
    errMsg = ""
    SetBreakpoint = True
End Function

Public Function ClearBreakpoint(ByVal text As String, ByVal id As String, ByRef newText As String, ByRef errMsg As String) As Boolean
    ClearBreakpoint = RemoveBlockByMarkers(text, "@mcp-break ", id, newText, errMsg)
End Function

' ---------------------------------------------------------------------------
' P0-4: injection pre-checks, adapted from the algorithms in EventTracer v223
' (Planet Source Code 33472, frmAddIn.frm) which is the same kind of tool:
'   AdjustForLineContinuation / IsFirstLineOfProcedure / InsertionPointFor / FindIndent.
'
' Why it matters:
'   * a multi-line procedure DECLARATION ("Sub Foo(a, _" / "b)") means di+1 is
'     still inside the parameter list - inserting there produced broken code;
'   * the body must be indented like the line it lands next to;
'   * blank/comment lines after the header should be stepped over so the trace
'     lands on the first executable line.
'
' CONTRACT: returns the index AFTER WHICH to insert, because InsertMarkedBlock
' copies lines 0..di inclusive before writing the block.
' ---------------------------------------------------------------------------

Private Function BodyInsertIndex(ByRef lines() As String, ByVal di As Long, ByRef indent As String) As Long
    Dim i As Long, s As String
    i = di + 1
    ' 1) walk past the rest of a continued declaration
    Do While i <= UBound(lines)
        MaybePump
        If EndsWithContinuation(lines(i - 1)) Then
            i = i + 1
        Else
            Exit Do
        End If
    Loop
    ' 2) step over blank and comment-only lines
    Do While i <= UBound(lines)
        MaybePump
        s = Trim$(lines(i))
        If Len(s) = 0 Or Left$(s, 1) = "'" Then
            i = i + 1
        Else
            Exit Do
        End If
    Loop
    If i > UBound(lines) Then i = UBound(lines)
    indent = LeadingSpace(lines(i))
    If Len(indent) = 0 Then
        ' fall back to the declaration's own indentation plus one level
        indent = LeadingSpace(lines(di)) & "    "
    End If
    ' InsertMarkedBlock inserts AFTER the index it is given, so hand back the
    ' line BEFORE the first executable one. Measured: returning the first
    ' executable line itself placed the block one line too late.
    BodyInsertIndex = i - 1
End Function

' VB6 continues a logical line with " _" (space, underscore, end of line).
Private Function EndsWithContinuation(ByVal ln As String) As Boolean
    Dim s As String
    s = RTrim$(ln)
    If Len(s) < 2 Then Exit Function
    If Right$(s, 1) <> "_" Then Exit Function
    If Mid$(s, Len(s) - 1, 1) <> " " Then Exit Function
    EndsWithContinuation = True
End Function

Private Function LeadingSpace(ByVal ln As String) As String
    Dim i As Long, ch As String, r As String
    For i = 1 To Len(ln)
        ch = Mid$(ln, i, 1)
        If ch = " " Or ch = vbTab Then
            r = r & ch
        Else
            Exit For
        End If
    Next
    LeadingSpace = r
End Function

Private Function AsLf(ByVal text As String) As String
    AsLf = Replace(Replace(text, vbCrLf, vbLf), vbCr, vbLf)
End Function


' ---------------- TASK-P1-129: generated modMcpTrace ----------------

Public Function TraceModuleMarker() As String
    TraceModuleMarker = TG_MARKER
End Function

' True only for a module THIS add-in generated: both header variants emit the
' marker, so a hand-written modMcpTrace is never mistaken for ours.
Public Function IsAutogenTraceModule(ByVal code As String) As Boolean
    IsAutogenTraceModule = (InStr(1, code, TG_MARKER, vbTextCompare) > 0)
End Function

' The module source for one language. "zh" is the ONLY value that selects the
' Chinese header; anything else gives English.
Public Function TraceModuleSource(ByVal lang As String) As String
    If LCase$(Trim$(lang)) = "zh" Then
        TraceModuleSource = TgModuleZh()
    Else
        TraceModuleSource = TgModuleEn()
    End If
End Function

' TASK-P1-129 (final ruling): the comment language is chosen UI-FIRST.
'   1) the IDE's own UI language (menu captions, else the title mode marker) wins;
'   2) if the UI cannot be read, the running VB6.EXE's version resource decides
'      (PRIMARYLANGID 0x04 = LANG_CHINESE -> zh, anything else -> en);
'   3) if neither can be read -> English (readable on every code page, while Chinese
'      becomes mojibake on a non-Chinese one).
' All three signals are handed back so the receipt can show them side by side: the
' development machine measures UI=zh with a resource of en, and the receipt must make
' that contradiction visible instead of hiding it.
Public Function DetectIdeLanguage(ByRef src As String, ByRef resLang As String, ByRef resWhy As String, _
                                 ByRef uiLang As String, ByRef uiWhy As String) As String
    resLang = TgResourceLanguage(resWhy)
    uiLang = TgIdeUiLanguage(uiWhy)
    If Len(uiLang) > 0 Then
        src = "UI-first: IDE UI language = " & uiLang & " (" & uiWhy & "); VB6.EXE version resource = " & TgSignalText(resLang, resWhy) & " (NOT used while the UI is readable)"
        DetectIdeLanguage = uiLang
        Exit Function
    End If
    If Len(resLang) > 0 Then
        src = "UI unavailable (" & uiWhy & "), so the VB6.EXE version resource decided: " & resWhy
        DetectIdeLanguage = resLang
        Exit Function
    End If
    src = "UI unavailable (" & uiWhy & ") and the VB6.EXE version resource is unreadable -> English"
    DetectIdeLanguage = "en"
End Function

Private Function TgSignalText(ByVal lang As String, ByVal why As String) As String
    If Len(lang) = 0 Then
        TgSignalText = "unreadable"
    Else
        TgSignalText = lang & " (" & why & ")"
    End If
End Function

' The language of the running VB6.EXE according to its OWN version resource:
' "zh" when PRIMARYLANGID = 0x04 (LANG_CHINESE), "en" otherwise, "" when unreadable.
Private Function TgResourceLanguage(ByRef why As String) As String
    Dim langId As Long
    langId = TgVersionLangId()
    If langId = 0 Then
        why = "VB6.EXE version resource unreadable (GetFileVersionInfo/VerQueryValue failed)"
        TgResourceLanguage = ""
        Exit Function
    End If
    If (langId And &H3FF&) = &H4& Then
        why = "VB6.EXE version resource LANGID 0x" & Hex$(langId) & " (LANG_CHINESE)"
        TgResourceLanguage = "zh"
    Else
        why = "VB6.EXE version resource LANGID 0x" & Hex$(langId) & " (not Chinese)"
        TgResourceLanguage = "en"
    End If
End Function

' The language of the IDE USER INTERFACE, read from the IDE's own window: the menu
' captions (a broad set of localised words) first, the title's mode marker second.
' "" means it could not be read and the caller falls back to the version resource.
' Pure Win32 READS of a native window, so this is safe even in break mode.
Private Function TgIdeUiLanguage(ByRef why As String) As String
    why = ""
    Dim menuLang As String
    menuLang = TgIdeMenuLanguage()
    If Len(menuLang) > 0 Then
        If menuLang = "zh" Then
            why = "IDE main-menu captions (" & ChrW$(&H6587) & ChrW$(&H4EF6) & "/" & ChrW$(&H7F16) & ChrW$(&H8F91) & ")"
        Else
            why = "IDE main-menu captions (File/Edit/View/Help)"
        End If
        TgIdeUiLanguage = menuLang
        Exit Function
    End If
    Dim titleLang As String
    titleLang = TgTitleLanguage()
    If Len(titleLang) > 0 Then
        If titleLang = "zh" Then
            why = "IDE window title mode marker ([" & ChrW$(&H8BBE) & ChrW$(&H8BA1) & "])"
        Else
            why = "IDE window title mode marker ([design]/[run]/[break])"
        End If
        TgIdeUiLanguage = titleLang
        Exit Function
    End If
    why = "no IDE window found, or its menus and title carried no recognisable marker"
End Function

' \VarFileInfo\Translation of the running VB6.EXE; 0 when anything fails.
Private Function TgVersionLangId() As Long
    On Error GoTo fail
    TgVersionLangId = 0
    Dim buf As String, n As Long
    buf = String$(520, vbNullChar)
    n = GetModuleFileNameW(0, StrPtr(buf), 520)
    If n <= 0 Then Exit Function
    Dim exePath As String
    exePath = Left$(buf, n)
    Dim h As Long, sz As Long
    sz = GetFileVersionInfoSizeW(StrPtr(exePath), h)
    If sz <= 0 Then Exit Function
    Dim blk() As Byte
    ReDim blk(0 To sz - 1)
    If GetFileVersionInfoW(StrPtr(exePath), 0, sz, blk(0)) = 0 Then Exit Function
    Dim subBlk As String, lp As Long, ln As Long
    subBlk = "\VarFileInfo\Translation"
    If VerQueryValueW(blk(0), StrPtr(subBlk), lp, ln) = 0 Then Exit Function
    If lp = 0 Or ln < 2 Then Exit Function
    Dim langId As Long
    TgCopy langId, lp, 2
    TgVersionLangId = langId And &HFFFF&
    Exit Function
fail:
    TgVersionLangId = 0
End Function

' Menu captions of the IDE's main window: the ChrW-built Chinese words (wen-jian / bian-ji /
' shi-tu / bang-zhu) versus File / Edit / View / Help.
' "" means undecidable.
Private Function TgIdeMenuLanguage() As String
    On Error GoTo fail
    TgIdeMenuLanguage = ""
    Dim cls As String
    cls = TG_VB6_WNDCLASS
    Dim hWnd As Long
    hWnd = TgFindWindowW(StrPtr(cls), 0)
    If hWnd = 0 Then Exit Function
    Dim hMenu As Long
    hMenu = GetMenu(hWnd)
    If hMenu = 0 Then Exit Function
    Dim cnt As Long, i As Long, cap As String
    cnt = GetMenuItemCount(hMenu)
    If cnt <= 0 Then Exit Function
    Dim sawEn As Boolean, sawZh As Boolean
    For i = 0 To cnt - 1
        cap = TgMenuCaption(hMenu, i)
        If Len(cap) > 0 Then
            If TgIsKnownEn(cap) Then sawEn = True
            If TgIsKnownZh(cap) Then sawZh = True
        End If
    Next
    If sawZh And Not sawEn Then TgIdeMenuLanguage = "zh"
    If sawEn And Not sawZh Then TgIdeMenuLanguage = "en"
    Exit Function
fail:
    TgIdeMenuLanguage = ""
End Function

' The mode marker in the IDE title, e.g. "P129 - Microsoft Visual Basic [X]" where X is the
' ChrW-built Chinese design/run/break word (see TgTitleLanguage) or its English counterpart.
Private Function TgTitleLanguage() As String
    On Error GoTo fail
    TgTitleLanguage = ""
    Dim cls As String, hWnd As Long, n As Long, buf As String
    cls = TG_VB6_WNDCLASS
    hWnd = TgFindWindowW(StrPtr(cls), 0)
    If hWnd = 0 Then Exit Function
    buf = String$(520, vbNullChar)
    n = TgGetWindowTextW(hWnd, StrPtr(buf), 520)
    If n <= 0 Then Exit Function
    Dim t As String
    t = LCase$(Left$(buf, n))
    If InStr(t, "[" & ChrW$(&H8BBE) & ChrW$(&H8BA1) & "]") > 0 Then TgTitleLanguage = "zh": Exit Function
    If InStr(t, "[" & ChrW$(&H8FD0) & ChrW$(&H884C) & "]") > 0 Then TgTitleLanguage = "zh": Exit Function
    If InStr(t, "[" & ChrW$(&H4E2D) & ChrW$(&H65AD) & "]") > 0 Then TgTitleLanguage = "zh": Exit Function
    If InStr(t, "[design]") > 0 Then TgTitleLanguage = "en": Exit Function
    If InStr(t, "[run]") > 0 Then TgTitleLanguage = "en": Exit Function
    If InStr(t, "[break]") > 0 Then TgTitleLanguage = "en"
    Exit Function
fail:
    TgTitleLanguage = ""
End Function

Private Function TgMenuCaption(ByVal hMenu As Long, ByVal idx As Long) As String
    On Error Resume Next
    TgMenuCaption = ""
    Dim buf As String, n As Long
    buf = String$(260, vbNullChar)
    n = GetMenuStringW(hMenu, idx, StrPtr(buf), 260, TG_MF_BYPOSITION)
    If n > 0 Then TgMenuCaption = Left$(buf, n)
End Function

' Localised words are built with ChrW$ so this module's own source stays pure ASCII
' and can never itself become mojibake.
Private Function TgIsKnownEn(ByVal cap As String) As Boolean
    Dim t As String
    t = LCase$(cap)
    If InStr(t, "file") > 0 Then TgIsKnownEn = True
    If InStr(t, "edit") > 0 Then TgIsKnownEn = True
    If InStr(t, "view") > 0 Then TgIsKnownEn = True
    If InStr(t, "help") > 0 Then TgIsKnownEn = True
End Function

Private Function TgIsKnownZh(ByVal cap As String) As Boolean
    Dim wenjian As String, bianji As String, bangzhu As String, shitu As String
    wenjian = ChrW$(&H6587) & ChrW$(&H4EF6)
    bianji = ChrW$(&H7F16) & ChrW$(&H8F91)
    bangzhu = ChrW$(&H5E2E) & ChrW$(&H52A9)
    shitu = ChrW$(&H89C6) & ChrW$(&H56FE)
    If InStr(cap, wenjian) > 0 Then TgIsKnownZh = True
    If InStr(cap, bianji) > 0 Then TgIsKnownZh = True
    If InStr(cap, bangzhu) > 0 Then TgIsKnownZh = True
    If InStr(cap, shitu) > 0 Then TgIsKnownZh = True
End Function
Private Function TgModuleEn() As String
    Dim s As String
    s = Join(Array( _
        "Option Explicit", _
        "", _
        "' ===========================================================================", _
        "' THIS MODULE IS GENERATED AND ATTACHED TO YOUR PROJECT AUTOMATICALLY BY THE", _
        "' VB6MCP ADD-IN when the MCP tool inject_trace runs.", _
        "' VB6MCP-AUTOGEN modMcpTrace (lang=en)", _
        "' Do not hand-edit it. Re-running inject_trace is IDEMPOTENT: once this", _
        "' component exists it will not be added a second time.", _
        "'", _
        "' What it does:", _
        "'   Defines McpTrace() / McpTraceOn() / McpTraceStatus(). McpTrace writes", _
        "'   timestamped trace events to a log file; McpTraceOn() is the RUNTIME guard", _
        "'   used by the injected blocks and reads the environment at call time", _
        "'   (MCP_TRACE_FILE / MCP_TRACE_ON), so injected calls stay compiled even while", _
        "'   tracing is off.", _
        "'"), vbCrLf)
    s = s & vbCrLf & Join(Array( _
        "' Why it is here:", _
        "'   inject_trace inserts 'If McpTraceOn() Then ... End If' into your procedure.", _
        "'   That call must resolve at COMPILE time, which is exactly why this module has", _
        "'   to exist in your project.", _
        "'", _
        "' Shipping note:", _
        "'   If you keep the injected calls and build/ship an EXE or DLL, this module must", _
        "'   be compiled into it. If you no longer need tracing, the MCP tool remove_trace", _
        "'   removes both the injected blocks and this module.", _
        "' ===========================================================================", _
        "", _
        "' A trace sink is only useful if it can actually be written, and a", _
        "' silent failure here is worse than no trace at all. This module now (1) creates", _
        "' the sink directory, (2) falls back to %TEMP% when the configured sink is not", _
        "' usable, (3) NEVER hides a write failure - it appends it to MCP_TRACE_ERR_FILE", _
        "' (default %TEMP%\mcp_trace_err.log), and (4) can be asked what it resolved via"), vbCrLf)
    s = s & vbCrLf & Join(Array( _
        "' McpTraceStatus(). MCP_TRACE_DBGWIN=0 suppresses the OutputDebugStringW call for", _
        "' diagnosis (the file sink stays primary).", _
        "", _
        "Private Declare Function OutputDebugStringW Lib ""kernel32"" (ByVal lpString As Long) As Long", _
        "Private Declare Function GetCurrentProcessId Lib ""kernel32"" () As Long", _
        "", _
        "Private m_path As String", _
        "Private m_ready As Boolean", _
        "Private m_lastErr As String", _
        "Private m_fallback As Boolean", _
        "", _
        "Private Sub EnsureInit()", _
        "    If m_ready Then Exit Sub", _
        "    Dim want As String", _
        "    want = Environ$(""MCP_TRACE_FILE"")", _
        "    If Len(want) = 0 Then want = Environ$(""TEMP"") & ""\mcp_trace.jsonl"""), vbCrLf)
    s = s & vbCrLf & Join(Array( _
        "    m_path = want", _
        "    m_fallback = False", _
        "    m_lastErr = """"", _
        "    If Not PrepareSink(m_path) Then", _
        "        Dim alt As String", _
        "        alt = Environ$(""TEMP"") & ""\mcp_trace.jsonl""", _
        "        If LCase$(alt) <> LCase$(m_path) And PrepareSink(alt) Then", _
        "            m_lastErr = ""configured sink unusable: "" & m_path & "" -> fell back to "" & alt", _
        "            m_path = alt", _
        "            m_fallback = True", _
        "        Else", _
        "            m_lastErr = ""configured sink unusable and no usable fallback: "" & m_path", _
        "        End If", _
        "        TraceErr m_lastErr", _
        "    End If", _
        "    m_ready = True"), vbCrLf)
    s = s & vbCrLf & Join(Array( _
        "End Sub", _
        "", _
        "' Create the parent directory when it is missing, then PROVE the directory is", _
        "' writable by creating and deleting a probe file. Never assume.", _
        "Private Function PrepareSink(ByVal p As String) As Boolean", _
        "    On Error Resume Next", _
        "    PrepareSink = False", _
        "    Dim i As Long, dirPath As String, fn As Integer", _
        "    i = InStrRev(p, ""\"")", _
        "    If i <= 1 Then Exit Function", _
        "    dirPath = Left$(p, i - 1)", _
        "    EnsureDir dirPath", _
        "    err.Clear", _
        "    fn = FreeFile", _
        "    Open dirPath & ""\_mcp_trace_probe.tmp"" For Output As #fn", _
        "    If err.number <> 0 Then err.Clear: Exit Function"), vbCrLf)
    s = s & vbCrLf & Join(Array( _
        "    Close #fn", _
        "    Kill dirPath & ""\_mcp_trace_probe.tmp""", _
        "    err.Clear", _
        "    PrepareSink = True", _
        "End Function", _
        "", _
        "' MkDir only makes ONE level, so walk the path and create every missing level.", _
        "' NB: the variable must NOT be called 'dir' - VB6 is case-insensitive and would", _
        "' shadow the built-in Dir$, turning 'Dir$(dir, vbDirectory)' into an array access", _
        "' (measured: compile error 'Expected array').", _
        "Private Function EnsureDir(ByVal dirPath As String) As Boolean", _
        "    On Error Resume Next", _
        "    EnsureDir = False", _
        "    If Len(dirPath) = 0 Then Exit Function", _
        "    If Len(Dir$(dirPath, vbDirectory)) > 0 Then EnsureDir = True: Exit Function", _
        "    Dim parts() As String, i As Long, acc As String"), vbCrLf)
    s = s & vbCrLf & Join(Array( _
        "    parts = Split(dirPath, ""\"")", _
        "    acc = parts(0)", _
        "    If Right$(acc, 1) <> ""\"" Then acc = acc & ""\""", _
        "    For i = 1 To UBound(parts)", _
        "        If Len(parts(i)) > 0 Then", _
        "            If Right$(acc, 1) <> ""\"" Then acc = acc & ""\""", _
        "            acc = acc & parts(i)", _
        "            If Len(Dir$(acc, vbDirectory)) = 0 Then MkDir acc", _
        "        End If", _
        "    Next", _
        "    err.Clear", _
        "    EnsureDir = (Len(Dir$(dirPath, vbDirectory)) > 0)", _
        "End Function", _
        "", _
        "' Append a diagnostic line somewhere observable. This is the ONLY place that is", _
        "' allowed to fail silently, and even then it tries %TEMP%."), vbCrLf)
    s = s & vbCrLf & Join(Array( _
        "Private Sub TraceErr(ByVal msg As String)", _
        "    On Error Resume Next", _
        "    Dim p As String, fn As Integer", _
        "    p = Environ$(""MCP_TRACE_ERR_FILE"")", _
        "    If Len(p) = 0 Then p = Environ$(""TEMP"") & ""\mcp_trace_err.log""", _
        "    fn = FreeFile", _
        "    Open p For Append As #fn", _
        "    Print #fn, Format$(Now, ""yyyy-mm-dd hh:nn:ss"") & ""  "" & msg", _
        "    Close #fn", _
        "    err.Clear", _
        "End Sub", _
        "", _
        "' The injected guard calls this instead of relying on ""#If MCP_TRACE"".", _
        "' A compile-time #If only takes effect when the project is LOADED with that CondComp;", _
        "' a runtime check has no such load-order dependency.", _
        "'   ON  when MCP_TRACE_ON=1  or MCP_TRACE_FILE is set"), vbCrLf)
    s = s & vbCrLf & Join(Array( _
        "'   OFF when MCP_TRACE_OFF=1, or neither of the above is set", _
        "Public Function McpTraceOn() As Boolean", _
        "    On Error Resume Next", _
        "    EnsureInit", _
        "    If Environ$(""MCP_TRACE_OFF"") = ""1"" Then", _
        "        McpTraceOn = False", _
        "        Exit Function", _
        "    End If", _
        "    If Environ$(""MCP_TRACE_ON"") = ""1"" Then", _
        "        McpTraceOn = True", _
        "        Exit Function", _
        "    End If", _
        "    McpTraceOn = (Len(Environ$(""MCP_TRACE_FILE"")) > 0)", _
        "End Function", _
        "", _
        "' What did the module resolve, and did anything fail? (diagnosis hook)"), vbCrLf)
    s = s & vbCrLf & Join(Array( _
        "Public Function McpTraceStatus() As String", _
        "    On Error Resume Next", _
        "    EnsureInit", _
        "    McpTraceStatus = ""path="" & m_path & "" ; fallback="" & CStr(m_fallback) & "" ; lastErr="" & m_lastErr", _
        "End Function", _
        "", _
        "Public Sub McpTrace(ByVal id As String, ParamArray kv() As Variant)", _
        "    On Error Resume Next", _
        "    EnsureInit", _
        "    Dim q As String", _
        "    q = Chr$(34)", _
        "    Dim s As String", _
        "    s = ""{"" & q & ""id"" & q & "":"" & q & Esc(id) & q & "","" & q & ""pid"" & q & "":"" & CStr(GetCurrentProcessId())", _
        "    Dim i As Long", _
        "    For i = LBound(kv) To UBound(kv) Step 2", _
        "        If i + 1 <= UBound(kv) Then"), vbCrLf)
    s = s & vbCrLf & Join(Array( _
        "            s = s & "","" & q & Esc(CStr(kv(i))) & q & "":"" & q & Esc(SafeStr(kv(i + 1))) & q", _
        "        End If", _
        "    Next", _
        "    s = s & ""}""", _
        "    Dim fnum As Integer", _
        "    fnum = FreeFile", _
        "    On Error GoTo wrFail", _
        "    Open m_path For Append As #fnum", _
        "    Print #fnum, s", _
        "    Close #fnum", _
        "    On Error GoTo 0", _
        "    If Environ$(""MCP_TRACE_DBGWIN"") <> ""0"" Then OutputDebugStringW StrPtr(s)", _
        "    Exit Sub", _
        "wrFail:", _
        "    Dim em As String", _
        "    em = err.Description"), vbCrLf)
    s = s & vbCrLf & Join(Array( _
        "    On Error Resume Next", _
        "    Close #fnum", _
        "    TraceErr ""McpTrace write FAILED id="" & id & "" path="" & m_path & "" err="" & em & "" line="" & s", _
        "End Sub", _
        "", _
        "Private Function SafeStr(ByRef v As Variant) As String", _
        "    On Error Resume Next", _
        "    If IsObject(v) Then", _
        "        If v Is Nothing Then SafeStr = ""<Nothing>"" Else SafeStr = TypeName(v)", _
        "    ElseIf IsNull(v) Then", _
        "        SafeStr = ""<Null>""", _
        "    ElseIf IsEmpty(v) Then", _
        "        SafeStr = ""<Empty>""", _
        "    Else", _
        "        SafeStr = CStr(v)", _
        "    End If"), vbCrLf)
    s = s & vbCrLf & Join(Array( _
        "End Function", _
        "", _
        "Public Sub McpFaultDelay(ByVal ms As Long)", _
        "    Dim t As Double", _
        "    t = Timer", _
        "    Do While (Timer - t) * 1000# < CDbl(ms)", _
        "        DoEvents", _
        "    Loop", _
        "End Sub", _
        "", _
        "Private Function Esc(ByVal s As String) As String", _
        "    Dim bs As String, qt As String, t As String", _
        "    bs = Chr$(92)", _
        "    qt = Chr$(34)", _
        "    t = Replace(s, bs, bs & bs)", _
        "    t = Replace(t, qt, bs & qt)"), vbCrLf)
    s = s & vbCrLf & Join(Array( _
        "    t = Replace(t, vbCrLf, "" "")", _
        "    t = Replace(t, vbCr, "" "")", _
        "    t = Replace(t, vbLf, "" "")", _
        "    Esc = t", _
        "End Function"), vbCrLf)
    TgModuleEn = s
End Function

Private Function TgModuleZh() As String
    Dim s As String
    s = Join(Array( _
        "Option Explicit", _
        "", _
        "' ===========================================================================", _
        "' 本模块由 VB6MCP 插件在运行 MCP 工具 inject_trace 时自动生成并添加到你的工程。", _
        "' VB6MCP-AUTOGEN modMcpTrace (lang=zh)", _
        "' 请勿手工修改。重复调用 inject_trace 是幂等的：一旦本模块存在就不会再加第二份。", _
        "'", _
        "' 作用：", _
        "'   定义 McpTrace() / McpTraceOn() / McpTraceStatus()。McpTrace 把带时间戳的追踪", _
        "'   事件写入日志文件；McpTraceOn() 是注入块使用的运行时守卫，在调用时读取环境变量", _
        "'   （MCP_TRACE_FILE / MCP_TRACE_ON），所以即使追踪关闭，插入的调用也仍然能编译。", _
        "'", _
        "' 为什么在这里：", _
        "'   inject_trace 会往你的过程里插入 'If McpTraceOn() Then ... End If'，这个调用必须", _
        "'   在编译期就能解析，所以本模块必须存在于你的工程中。", _
        "'"), vbCrLf)
    s = s & vbCrLf & Join(Array( _
        "' 发布提醒：", _
        "'   如果你保留了插入的调用并要编译/发布 EXE 或 DLL，必须把本模块一起编进去。", _
        "'   不再需要追踪时，用 MCP 工具 remove_trace 可以一并移除插入块与本模块。", _
        "' ===========================================================================", _
        "", _
        "' A trace sink is only useful if it can actually be written, and a", _
        "' silent failure here is worse than no trace at all. This module now (1) creates", _
        "' the sink directory, (2) falls back to %TEMP% when the configured sink is not", _
        "' usable, (3) NEVER hides a write failure - it appends it to MCP_TRACE_ERR_FILE", _
        "' (default %TEMP%\mcp_trace_err.log), and (4) can be asked what it resolved via", _
        "' McpTraceStatus(). MCP_TRACE_DBGWIN=0 suppresses the OutputDebugStringW call for", _
        "' diagnosis (the file sink stays primary).", _
        "", _
        "Private Declare Function OutputDebugStringW Lib ""kernel32"" (ByVal lpString As Long) As Long", _
        "Private Declare Function GetCurrentProcessId Lib ""kernel32"" () As Long", _
        ""), vbCrLf)
    s = s & vbCrLf & Join(Array( _
        "Private m_path As String", _
        "Private m_ready As Boolean", _
        "Private m_lastErr As String", _
        "Private m_fallback As Boolean", _
        "", _
        "Private Sub EnsureInit()", _
        "    If m_ready Then Exit Sub", _
        "    Dim want As String", _
        "    want = Environ$(""MCP_TRACE_FILE"")", _
        "    If Len(want) = 0 Then want = Environ$(""TEMP"") & ""\mcp_trace.jsonl""", _
        "    m_path = want", _
        "    m_fallback = False", _
        "    m_lastErr = """"", _
        "    If Not PrepareSink(m_path) Then", _
        "        Dim alt As String", _
        "        alt = Environ$(""TEMP"") & ""\mcp_trace.jsonl"""), vbCrLf)
    s = s & vbCrLf & Join(Array( _
        "        If LCase$(alt) <> LCase$(m_path) And PrepareSink(alt) Then", _
        "            m_lastErr = ""configured sink unusable: "" & m_path & "" -> fell back to "" & alt", _
        "            m_path = alt", _
        "            m_fallback = True", _
        "        Else", _
        "            m_lastErr = ""configured sink unusable and no usable fallback: "" & m_path", _
        "        End If", _
        "        TraceErr m_lastErr", _
        "    End If", _
        "    m_ready = True", _
        "End Sub", _
        "", _
        "' Create the parent directory when it is missing, then PROVE the directory is", _
        "' writable by creating and deleting a probe file. Never assume.", _
        "Private Function PrepareSink(ByVal p As String) As Boolean", _
        "    On Error Resume Next"), vbCrLf)
    s = s & vbCrLf & Join(Array( _
        "    PrepareSink = False", _
        "    Dim i As Long, dirPath As String, fn As Integer", _
        "    i = InStrRev(p, ""\"")", _
        "    If i <= 1 Then Exit Function", _
        "    dirPath = Left$(p, i - 1)", _
        "    EnsureDir dirPath", _
        "    err.Clear", _
        "    fn = FreeFile", _
        "    Open dirPath & ""\_mcp_trace_probe.tmp"" For Output As #fn", _
        "    If err.number <> 0 Then err.Clear: Exit Function", _
        "    Close #fn", _
        "    Kill dirPath & ""\_mcp_trace_probe.tmp""", _
        "    err.Clear", _
        "    PrepareSink = True", _
        "End Function", _
        ""), vbCrLf)
    s = s & vbCrLf & Join(Array( _
        "' MkDir only makes ONE level, so walk the path and create every missing level.", _
        "' NB: the variable must NOT be called 'dir' - VB6 is case-insensitive and would", _
        "' shadow the built-in Dir$, turning 'Dir$(dir, vbDirectory)' into an array access", _
        "' (measured: compile error 'Expected array').", _
        "Private Function EnsureDir(ByVal dirPath As String) As Boolean", _
        "    On Error Resume Next", _
        "    EnsureDir = False", _
        "    If Len(dirPath) = 0 Then Exit Function", _
        "    If Len(Dir$(dirPath, vbDirectory)) > 0 Then EnsureDir = True: Exit Function", _
        "    Dim parts() As String, i As Long, acc As String", _
        "    parts = Split(dirPath, ""\"")", _
        "    acc = parts(0)", _
        "    If Right$(acc, 1) <> ""\"" Then acc = acc & ""\""", _
        "    For i = 1 To UBound(parts)", _
        "        If Len(parts(i)) > 0 Then", _
        "            If Right$(acc, 1) <> ""\"" Then acc = acc & ""\"""), vbCrLf)
    s = s & vbCrLf & Join(Array( _
        "            acc = acc & parts(i)", _
        "            If Len(Dir$(acc, vbDirectory)) = 0 Then MkDir acc", _
        "        End If", _
        "    Next", _
        "    err.Clear", _
        "    EnsureDir = (Len(Dir$(dirPath, vbDirectory)) > 0)", _
        "End Function", _
        "", _
        "' Append a diagnostic line somewhere observable. This is the ONLY place that is", _
        "' allowed to fail silently, and even then it tries %TEMP%.", _
        "Private Sub TraceErr(ByVal msg As String)", _
        "    On Error Resume Next", _
        "    Dim p As String, fn As Integer", _
        "    p = Environ$(""MCP_TRACE_ERR_FILE"")", _
        "    If Len(p) = 0 Then p = Environ$(""TEMP"") & ""\mcp_trace_err.log""", _
        "    fn = FreeFile"), vbCrLf)
    s = s & vbCrLf & Join(Array( _
        "    Open p For Append As #fn", _
        "    Print #fn, Format$(Now, ""yyyy-mm-dd hh:nn:ss"") & ""  "" & msg", _
        "    Close #fn", _
        "    err.Clear", _
        "End Sub", _
        "", _
        "' The injected guard calls this instead of relying on ""#If MCP_TRACE"".", _
        "' A compile-time #If only takes effect when the project is LOADED with that CondComp;", _
        "' a runtime check has no such load-order dependency.", _
        "'   ON  when MCP_TRACE_ON=1  or MCP_TRACE_FILE is set", _
        "'   OFF when MCP_TRACE_OFF=1, or neither of the above is set", _
        "Public Function McpTraceOn() As Boolean", _
        "    On Error Resume Next", _
        "    EnsureInit", _
        "    If Environ$(""MCP_TRACE_OFF"") = ""1"" Then", _
        "        McpTraceOn = False"), vbCrLf)
    s = s & vbCrLf & Join(Array( _
        "        Exit Function", _
        "    End If", _
        "    If Environ$(""MCP_TRACE_ON"") = ""1"" Then", _
        "        McpTraceOn = True", _
        "        Exit Function", _
        "    End If", _
        "    McpTraceOn = (Len(Environ$(""MCP_TRACE_FILE"")) > 0)", _
        "End Function", _
        "", _
        "' What did the module resolve, and did anything fail? (diagnosis hook)", _
        "Public Function McpTraceStatus() As String", _
        "    On Error Resume Next", _
        "    EnsureInit", _
        "    McpTraceStatus = ""path="" & m_path & "" ; fallback="" & CStr(m_fallback) & "" ; lastErr="" & m_lastErr", _
        "End Function", _
        ""), vbCrLf)
    s = s & vbCrLf & Join(Array( _
        "Public Sub McpTrace(ByVal id As String, ParamArray kv() As Variant)", _
        "    On Error Resume Next", _
        "    EnsureInit", _
        "    Dim q As String", _
        "    q = Chr$(34)", _
        "    Dim s As String", _
        "    s = ""{"" & q & ""id"" & q & "":"" & q & Esc(id) & q & "","" & q & ""pid"" & q & "":"" & CStr(GetCurrentProcessId())", _
        "    Dim i As Long", _
        "    For i = LBound(kv) To UBound(kv) Step 2", _
        "        If i + 1 <= UBound(kv) Then", _
        "            s = s & "","" & q & Esc(CStr(kv(i))) & q & "":"" & q & Esc(SafeStr(kv(i + 1))) & q", _
        "        End If", _
        "    Next", _
        "    s = s & ""}""", _
        "    Dim fnum As Integer", _
        "    fnum = FreeFile"), vbCrLf)
    s = s & vbCrLf & Join(Array( _
        "    On Error GoTo wrFail", _
        "    Open m_path For Append As #fnum", _
        "    Print #fnum, s", _
        "    Close #fnum", _
        "    On Error GoTo 0", _
        "    If Environ$(""MCP_TRACE_DBGWIN"") <> ""0"" Then OutputDebugStringW StrPtr(s)", _
        "    Exit Sub", _
        "wrFail:", _
        "    Dim em As String", _
        "    em = err.Description", _
        "    On Error Resume Next", _
        "    Close #fnum", _
        "    TraceErr ""McpTrace write FAILED id="" & id & "" path="" & m_path & "" err="" & em & "" line="" & s", _
        "End Sub", _
        "", _
        "Private Function SafeStr(ByRef v As Variant) As String"), vbCrLf)
    s = s & vbCrLf & Join(Array( _
        "    On Error Resume Next", _
        "    If IsObject(v) Then", _
        "        If v Is Nothing Then SafeStr = ""<Nothing>"" Else SafeStr = TypeName(v)", _
        "    ElseIf IsNull(v) Then", _
        "        SafeStr = ""<Null>""", _
        "    ElseIf IsEmpty(v) Then", _
        "        SafeStr = ""<Empty>""", _
        "    Else", _
        "        SafeStr = CStr(v)", _
        "    End If", _
        "End Function", _
        "", _
        "Public Sub McpFaultDelay(ByVal ms As Long)", _
        "    Dim t As Double", _
        "    t = Timer", _
        "    Do While (Timer - t) * 1000# < CDbl(ms)"), vbCrLf)
    s = s & vbCrLf & Join(Array( _
        "        DoEvents", _
        "    Loop", _
        "End Sub", _
        "", _
        "Private Function Esc(ByVal s As String) As String", _
        "    Dim bs As String, qt As String, t As String", _
        "    bs = Chr$(92)", _
        "    qt = Chr$(34)", _
        "    t = Replace(s, bs, bs & bs)", _
        "    t = Replace(t, qt, bs & qt)", _
        "    t = Replace(t, vbCrLf, "" "")", _
        "    t = Replace(t, vbCr, "" "")", _
        "    t = Replace(t, vbLf, "" "")", _
        "    Esc = t", _
        "End Function"), vbCrLf)
    TgModuleZh = s
End Function
