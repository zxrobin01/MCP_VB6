Attribute VB_Name = "Vb6Lint"
Option Explicit

Public Type TDiag
    ruleId As String
    severity As String
    physLine As Long
    msg As String
End Type

Private m_raw() As String
Private m_code() As String
Private m_endCont() As Boolean
Private m_n As Long
Private m_lc() As String
Private m_ls() As Long
Private m_le() As Long
Private m_nlog As Long
Private m_inProc() As Boolean
Private m_procStart() As Long
Private m_procEnd() As Long
Private m_procName() As String
' item F-d/D1 (E116): the KIND of each procedure (sub/function/property).
' Filled in ComputeProcs from the SAME StripMods-based detection IsProcStart already
' uses, so the two can never disagree.
Private m_procKind() As String
Private m_nproc As Long
Private m_firstProcLine As Long
Private m_diag() As TDiag
Private m_nd As Long
' (F2): m_nd stays the FULL diagnostic count (DiagCount's meaning is
' unchanged). m_nerr/m_nwarn count every accepted diagnostic by severity so the
' summary no longer has to scan the array; m_nst is how many diagnostics were
' actually materialised and m_nstoreMax is the caller's cap (0 = store everything,
' i.e. the historical behaviour, so a caller that does not opt in is bit-identical).
Private m_nerr As Long
Private m_nwarn As Long
Private m_nst As Long
Private m_nstoreMax As Long
Private m_hasOptExp As Boolean
Private m_silent As Boolean
Private m_baseKeys As Object
Private m_baseHits As Long
Private m_regR0 As Long
Private m_regR1 As Long
' item 8 (F-b): how many rule ids this engine can emit. The receipt used to
' report only how many diagnostics FIRED, which a caller reads as coverage. Counted from
' the source: E101..E115 (15) + W301 + W304 + W305 + E116 + E117 (F-d) = 20.
' One constant, in the DECLARATIONS section (VB6 requires Const/Declare/Dim to precede
' every procedure); adding a rule moves exactly this number.
Private Const RULE_COUNT As Long = 20

Public Sub Lint(ByVal text As String, ByVal compKind As String)
    m_regR0 = 0
    m_regR1 = 0
    ResetScan
    ResetRun
    LoadLines text
    BuildLogical
    ComputeProcs
    RunRules compKind
End Sub

' per-run state that must be cleared before the rules execute.
Private Sub ResetRun()
    m_nd = 0
    m_nerr = 0
    m_nwarn = 0
    m_nst = 0
    ReDim m_diag(0 To 63)
    m_silent = False
    m_baseHits = 0
    ' m_nstoreMax is deliberately NOT reset here: it is a CALLER setting (SetStoreCap).
End Sub

' state produced by the structural pre-scan (LoadLines/BuildLogical/
' ComputeProcs), cleared before that pre-scan runs.
Private Sub ResetScan()
    m_firstProcLine = -1
    m_hasOptExp = False
End Sub

Private Sub RunRules(ByVal compKind As String)
    RuleBlocks
    RuleProcTerm
    RuleDupProcs
    RuleStray
    RuleDecls compKind
    RuleOption
    RuleLex
    RuleAgent
    RuleSemantic
    RuleUndeclaredConst
    ' F-d (an earlier lint batch): the two NEW rules run last, so every existing rule's
    ' diagnostic order and output is untouched.
    RuleExitKind
    RuleConstAssign
End Sub

' region scope: is this physical (1-based) line inside the linted region?
' m_regR1 = 0 means "no region" (whole module).
Private Function InReg(ByVal ln As Long) As Boolean
    If m_regR1 <= 0 Then
        InReg = True
    Else
        InReg = (ln >= m_regR0 And ln <= m_regR1)
    End If
End Function

Private Function ProcOverlapsRegion(ByVal p As Long) As Boolean
    If m_regR1 <= 0 Then
        ProcOverlapsRegion = True
        Exit Function
    End If
    Dim s0 As Long, s1 As Long
    s0 = m_procStart(p) + 1
    s1 = m_procEnd(p) + 1
    ProcOverlapsRegion = (s0 <= m_regR1 And s1 >= m_regR0)
End Function

' lint ONLY the procedure that contains physLine, so a large module
' does not pay a whole-module pass for a small edit. The rule set, rule order and
' the physical/IDE line space are THE SAME as Lint's: every rule still runs (so
' global context - duplicate procedure names, the module-wide declared-name set,
' Option Explicit, the block stack carried from earlier code - is EXACT), but the
' expensive per-line scans are skipped outside the region and AddDiag drops any
' diagnostic outside it. The caller can therefore compare this result with a
' whole-module Lint filtered to the same range and get an identical sequence.
' Returns True when a procedure contained physLine (diagnostics are region-scoped,
' and procName/r0/r1 describe it); False when no procedure contains the line, in
' which case the WHOLE module was linted and the caller must NOT claim region scope.
Public Function LintRegionAt(ByVal text As String, ByVal compKind As String, ByVal physLine As Long, ByRef procName As String, ByRef r0 As Long, ByRef r1 As Long) As Boolean
    Dim i As Long, ln As Long
    procName = ""
    r0 = 0
    r1 = 0
    m_regR0 = 0
    m_regR1 = 0
    ResetScan
    ResetRun
    LoadLines text
    BuildLogical
    ComputeProcs
    ln = physLine - 1
    If ln >= 0 Then
        For i = 0 To m_nproc - 1
            MaybePump
            If ln >= m_procStart(i) And ln <= m_procEnd(i) Then
                procName = m_procName(i)
                r0 = m_procStart(i) + 1
                r1 = m_procEnd(i) + 1
                Exit For
            End If
        Next i
    End If
    If r0 = 0 Then
        ' No procedure contains the line: the whole-module pass is already complete
        ' (ComputeProcs has emitted its E112 diagnostics), so just run the rules.
        RunRules compKind
        LintRegionAt = False
        Exit Function
    End If
    m_regR0 = r0
    m_regR1 = r1
    ResetRun
    ' ComputeProcs emitted its E112 (reserved word as a procedure name) diagnostics
    ' before the region was known. Re-emit exactly those whose line is inside the
    ' region, in ComputeProcs's own order, instead of paying a second whole-module
    ' scan. The output is identical to running ComputeProcs with the region active.
    For i = 0 To m_nproc - 1
        MaybePump
        If IsReserved(m_procName(i)) Then
            AddDiag "VB6-E112", "error", m_procStart(i) + 1, "reserved word as procedure name: " & m_procName(i)
        End If
    Next i
    RunRules compKind
    LintRegionAt = True
End Function

Public Property Get DiagCount() As Long
    DiagCount = m_nd
End Property

' (F2): DiagCount is the FULL count (unchanged semantics). DiagStored is
' how many diagnostics are really in m_diag and is <= DiagCount whenever a store cap is
' in force, so every ARRAY reader must iterate DiagStored - never DiagCount. The two
' counters are the whole-summary sources: they count every diagnostic AddDiag accepted.
Public Property Get DiagStored() As Long
    DiagStored = m_nst
End Property
Public Property Get ErrorCount() As Long
    ErrorCount = m_nerr
End Property
Public Property Get WarningCount() As Long
    WarningCount = m_nwarn
End Property

' (F2): 0 = materialise every diagnostic (the historical behaviour);
' n > 0 = materialise at most n. The FULL counters are unaffected either way.
Public Sub SetStoreCap(ByVal n As Long)
    If n < 0 Then n = 0
    m_nstoreMax = n
End Sub
' (F2): the four raw array readers are bounded by what is STORED, not by
' DiagCount. A caller that iterates 0..DiagCount()-1 while a store cap is in force would
' otherwise read uninitialised slots; these guards turn that into "" / 0.
Public Function RuleCount() As Long
    RuleCount = RULE_COUNT
End Function

Public Function DiagRule(ByVal i As Long) As String
    If i < 0 Or i >= m_nst Then Exit Function
    DiagRule = m_diag(i).ruleId
End Function
Public Function DiagSev(ByVal i As Long) As String
    If i < 0 Or i >= m_nst Then Exit Function
    DiagSev = m_diag(i).severity
End Function
Public Function DiagLine(ByVal i As Long) As Long
    If i < 0 Or i >= m_nst Then Exit Function
    DiagLine = m_diag(i).physLine
End Function
Public Function DiagMsg(ByVal i As Long) As String
    If i < 0 Or i >= m_nst Then Exit Function
    DiagMsg = m_diag(i).msg
End Function

Private Sub AddDiag(ByVal rid As String, ByVal sev As String, ByVal ln As Long, ByVal msg As String)
    If m_silent Then Exit Sub
    ' region scope: a diagnostic outside the linted region is not reported.
    If m_regR1 > 0 Then
        If ln < m_regR0 Or ln > m_regR1 Then Exit Sub
    End If
    ' the dedup key is only needed when a baseline was installed. It used
    ' to be built for EVERY diagnostic (one NormCtx = LCase + Trim + a Replace loop,
    ' 27,005 times on the module-level fixture) even though m_baseKeys was Nothing.
    If Not m_baseKeys Is Nothing Then
        Dim kk As String
        kk = rid & ":" & NormCtx(ln)
        If m_baseKeys.Exists(kk) Then
            m_baseHits = m_baseHits + 1
            Exit Sub
        End If
    End If
    ' (F2): COUNT every diagnostic, MATERIALISE at most m_nstoreMax of them.
    ' The receipt cap (CodeTools.LintMaxDiag, default 300) is handed in by the caller via
    ' SetStoreCap, so a 27,005-diagnostic module no longer pays 27,005 UDT+BSTR writes for
    ' an array whose consumers keep at most 300 entries. m_nd keeps the FULL count, so
    ' DiagCount / diagnosticsTotal / lintTruncated / strict refusal are unchanged.
    m_nd = m_nd + 1
    If sev = "error" Then
        m_nerr = m_nerr + 1
    Else
        m_nwarn = m_nwarn + 1
    End If
    If m_nstoreMax > 0 Then
        If m_nst >= m_nstoreMax Then Exit Sub
    End If
    ' GEOMETRIC growth. Growing by a fixed +64 was QUADRATIC: the
    ' module-level fixture emits 27,005 diagnostics, i.e. 422 ReDim Preserve calls
    ' copying an average of ~13,500 UDT-with-BSTR elements each (~5.7M copies).
    ' Only the CAPACITY changes: the diagnostics, their order, every counter and
    ' every receipt key are identical (DiagCount returns m_nd, not the bound).
    If m_nst > UBound(m_diag) Then
        Dim ncap As Long
        ncap = (UBound(m_diag) + 1) * 2 + 63
        If m_nstoreMax > 0 Then
            If ncap > m_nstoreMax - 1 Then ncap = m_nstoreMax - 1
        End If
        ReDim Preserve m_diag(0 To ncap)
    End If
    m_diag(m_nst).ruleId = rid
    m_diag(m_nst).severity = sev
    m_diag(m_nst).physLine = ln
    m_diag(m_nst).msg = msg
    m_nst = m_nst + 1
End Sub

' ---------------- lexer ----------------

Private Sub LoadLines(ByVal text As String)
    Dim t As String
    t = Replace(text, vbCrLf, vbLf)
    t = Replace(t, vbCr, vbLf)
    m_raw = Split(t, vbLf)
    ' an EMPTY pane (CountOfLines = 0 - e.g. the empty module a
    ' name-colliding import creates) is a LEGAL input. VB6's Split("") returns an
    ' EMPTY array (UBound = -1), so the "0 To m_n - 1" Redims below were
    ' ReDim x(0 To -1) => VB error 9 (下标越界); that is the subscript error that
    ' 's empty-component guard then exposed on verify. Model the empty
    ' module as exactly ONE blank line so every downstream rule sees a valid array.
    If UBound(m_raw) < 0 Then
        ReDim m_raw(0 To 0)
        m_raw(0) = ""
    End If
    m_n = UBound(m_raw) + 1
    ReDim m_code(0 To m_n - 1)
    ReDim m_endCont(0 To m_n - 1)
    Dim i As Long
    For i = 0 To m_n - 1
        MaybePump
        m_code(i) = StripCode(m_raw(i))
        m_endCont(i) = EndsCont(m_code(i))
    Next
End Sub

Public Function StripCode(ByVal s As String) As String
    ' same output as before, but O(n) per line. The old body appended
    ' one character at a time (buf = buf & ch), which is O(n^2) on a BSTR; a
    ' 1.5 MB module paid that on every one of its ~27k lines. Here quoted runs are
    ' DROPPED by slicing around them, so the surviving text is built in a handful
    ' of Mid$ slices instead of one concatenation per character. Semantics
    ' (doubled "" escape, trailing comment cut, Trim) are unchanged.
    Dim i As Long, n As Long, inside As Boolean, ch As String, nxt As String
    Dim segStart As Long, buf As String, hitComment As Boolean
    ' Fast path: with no double quote and no apostrophe the whole line IS the code,
    ' so there is nothing to strip. This is the common case in real modules and it
    ' avoids the per-character walk entirely (two C-level InStr scans instead).
    If InStr(s, Chr$(34)) = 0 Then
        If InStr(s, Chr$(39)) = 0 Then
            StripCode = Trim$(s)
            Exit Function
        End If
    End If
    n = Len(s)
    i = 1
    segStart = 1
    Do While i <= n
        ch = Mid$(s, i, 1)
        If ch = Chr$(34) Then
            If Not inside Then
                buf = buf & Mid$(s, segStart, i - segStart)
                inside = True
                i = i + 1
                segStart = i
            Else
                nxt = Mid$(s, i + 1, 1)
                If nxt = Chr$(34) Then
                    i = i + 2
                    segStart = i
                Else
                    inside = False
                    i = i + 1
                    segStart = i
                End If
            End If
        ElseIf ch = Chr$(39) And Not inside Then
            buf = buf & Mid$(s, segStart, i - segStart)
            hitComment = True
            Exit Do
        Else
            i = i + 1
        End If
    Loop
    If Not hitComment And Not inside Then
        If segStart <= n Then buf = buf & Mid$(s, segStart, n - segStart + 1)
    End If
    StripCode = Trim$(buf)
End Function

Private Function EndsCont(ByVal s As String) As Boolean
    Dim t As String, pc As String
    t = RTrim$(s)
    If Len(t) < 2 Then Exit Function
    If Right$(t, 1) <> "_" Then Exit Function
    pc = Mid$(t, Len(t) - 1, 1)
    If pc = " " Or pc = vbTab Then EndsCont = True
End Function

Private Sub BuildLogical()
    ReDim m_lc(0 To m_n)
    ReDim m_ls(0 To m_n)
    ReDim m_le(0 To m_n)
    m_nlog = 0
    Dim i As Long, inLog As Boolean, buf As String, sp As Long, lp As Long
    Dim piece As String
    inLog = False
    For i = 0 To m_n - 1
        MaybePump
        If Not inLog Then
            sp = i
            buf = ""
            inLog = True
        End If
        piece = m_code(i)
        If m_endCont(i) Then
            If Len(piece) > 0 Then piece = Left$(piece, Len(piece) - 1)
        End If
        piece = Trim$(piece)
        If Len(buf) > 0 Then
            If Len(piece) > 0 Then buf = buf & " " & piece
        Else
            buf = piece
        End If
        lp = i
        If Not m_endCont(i) Then
            m_lc(m_nlog) = LCase$(buf)
            m_ls(m_nlog) = sp
            m_le(m_nlog) = lp
            m_nlog = m_nlog + 1
            inLog = False
        End If
    Next
    If inLog Then
        m_lc(m_nlog) = LCase$(buf)
        m_ls(m_nlog) = sp
        m_le(m_nlog) = lp
        m_nlog = m_nlog + 1
    End If
End Sub

Private Sub ComputeProcs()
    ReDim m_inProc(0 To m_n - 1)
    ReDim m_procStart(0 To m_nlog)
    ReDim m_procEnd(0 To m_nlog)
    ReDim m_procName(0 To m_nlog)
    ReDim m_procKind(0 To m_nlog)
    m_nproc = 0
    Dim k As Long, isOpen As Boolean, st As Long, nm As String
    ' F-d/D1: the KIND has to be captured on the DECLARATION line. Reading it
    ' when the procedure CLOSES would read "end sub" and yield "" - which then mismatched
    ' every Exit statement (the first F-d run reported G01..G04/G11 as E116 false
    ' positives for exactly this reason).
    Dim kindCur As String
    isOpen = False
    nm = ""
    For k = 0 To m_nlog - 1
        MaybePump
        If Not isOpen Then
            If IsProcStart(m_lc(k), nm) Then
                isOpen = True
                st = m_ls(k)
                kindCur = ProcKindOf(m_lc(k))
                If m_firstProcLine < 0 Then m_firstProcLine = m_ls(k)
                If IsReserved(nm) Then AddDiag "VB6-E112", "error", m_ls(k) + 1, "reserved word as procedure name: " & nm
            End If
        Else
            If IsProcEnd(m_lc(k)) Then
                m_procStart(m_nproc) = st
                m_procEnd(m_nproc) = m_le(k)
                m_procName(m_nproc) = nm
                m_procKind(m_nproc) = kindCur
                m_nproc = m_nproc + 1
                isOpen = False
            End If
        End If
    Next
    If isOpen Then
        m_procStart(m_nproc) = st
        m_procEnd(m_nproc) = m_n - 1
        m_procName(m_nproc) = nm
        m_procKind(m_nproc) = kindCur
        m_nproc = m_nproc + 1
    End If
    Dim p As Long, j As Long
    For p = 0 To m_nproc - 1
        MaybePump
        For j = m_procStart(p) To m_procEnd(p)
            MaybePump
            If j >= 0 And j < m_n Then m_inProc(j) = True
        Next
    Next
End Sub

Private Function StripMods(ByVal s As String) As String
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
    StripMods = t
End Function

Private Function IsProcStart(ByVal L As String, ByRef nm As String) As Boolean
    Dim s As String
    s = StripMods(L)
    nm = ""
    If Left$(s, 4) = "sub " Then nm = Tok(s, 2) : IsProcStart = True : Exit Function
    If Left$(s, 9) = "function " Then nm = Tok(s, 2) : IsProcStart = True : Exit Function
    If Left$(s, 9) = "property " Then nm = Tok(s, 3) : IsProcStart = True : Exit Function
End Function

Private Function IsProcEnd(ByVal L As String) As Boolean
    If L = "end sub" Then IsProcEnd = True
    If L = "end function" Then IsProcEnd = True
    If L = "end property" Then IsProcEnd = True
End Function

Private Function Tok(ByVal s As String, ByVal n As Long) As String
    Dim parts() As String, r As String, p As Long
    parts = Split(Trim$(s), " ")
    If n - 1 > UBound(parts) Then Tok = "" : Exit Function
    r = parts(n - 1)
    p = InStr(r, "(")
    If p > 0 Then r = Left$(r, p - 1)
    Tok = r
End Function

Private Function HasWord(ByVal L As String, ByVal w As String) As Boolean
    HasWord = (" " & L & " ") Like ("* " & w & " *")
End Function

' ---------------- rules: blocks ----------------

Private Sub RuleBlocks()
    Dim sk() As String, sv() As String, sl() As Long, ns As Long
    Dim ck() As String, cl() As Long, nc As Long
    ReDim sk(0 To m_nlog)
    ReDim sv(0 To m_nlog)
    ReDim sl(0 To m_nlog)
    ReDim ck(0 To m_nlog)
    ReDim cl(0 To m_nlog)
    ns = 0
    nc = 0
    Dim k As Long, L As String, ln As Long, fv As String
    For k = 0 To m_nlog - 1
        MaybePump
        L = m_lc(k)
        ln = m_ls(k) + 1
        If m_regR1 > 0 Then
            If ln > m_regR1 Then Exit For
        End If
        If Left$(L, 1) = "#" Then
            If Left$(L, 3) = "#if" Then
                ck(nc) = "#if"
                cl(nc) = ln
                nc = nc + 1
            ElseIf L = "#end if" Then
                If nc > 0 Then nc = nc - 1
            End If
        ElseIf IsBlockIfStart(L) Then
            sk(ns) = "if" : sv(ns) = "" : sl(ns) = ln : ns = ns + 1
        ElseIf L = "end if" Then
            PopB sk, sv, sl, ns, "if", "", ln, "VB6-E101", "End If"
        ElseIf IsForStart(L, fv) Then
            sk(ns) = "for" : sv(ns) = fv : sl(ns) = ln : ns = ns + 1
        ElseIf L = "next" Or Left$(L, 5) = "next " Then
            PopB sk, sv, sl, ns, "for", NextVar(L), ln, "VB6-E102", "Next"
        ElseIf L = "do" Or Left$(L, 3) = "do " Then
            sk(ns) = "do" : sv(ns) = "" : sl(ns) = ln : ns = ns + 1
        ElseIf L = "loop" Or Left$(L, 5) = "loop " Then
            PopB sk, sv, sl, ns, "do", "", ln, "VB6-E103", "Loop"
        ElseIf Left$(L, 6) = "while " Then
            sk(ns) = "while" : sv(ns) = "" : sl(ns) = ln : ns = ns + 1
        ElseIf L = "wend" Then
            PopB sk, sv, sl, ns, "while", "", ln, "VB6-E103", "Wend"
        ElseIf Left$(L, 5) = "with " Then
            sk(ns) = "with" : sv(ns) = "" : sl(ns) = ln : ns = ns + 1
        ElseIf L = "end with" Then
            PopB sk, sv, sl, ns, "with", "", ln, "VB6-E104", "End With"
        ElseIf Left$(L, 11) = "select case" Then
            sk(ns) = "select" : sv(ns) = "" : sl(ns) = ln : ns = ns + 1
        ElseIf L = "end select" Then
            PopB sk, sv, sl, ns, "select", "", ln, "VB6-E104", "End Select"
        ElseIf IsTypeStart(L) Then
            sk(ns) = "type" : sv(ns) = "" : sl(ns) = ln : ns = ns + 1
        ElseIf L = "end type" Or L = "end enum" Then
            PopB sk, sv, sl, ns, "type", "", ln, "VB6-E106", "End Type/Enum"
        End If
    Next
    Dim i As Long
    For i = ns - 1 To 0 Step -1
        AddDiag RuleForKind(sk(i)), "error", sl(i), "unclosed block: " & sk(i)
    Next
    For i = 0 To nc - 1
        MaybePump
        AddDiag "VB6-E106", "error", cl(i), "unclosed #If"
    Next
End Sub

Private Sub PopB(ByRef sk() As String, ByRef sv() As String, ByRef sl() As Long, ByRef n As Long, ByVal kind As String, ByVal extra As String, ByVal ln As Long, ByVal rid As String, ByVal nm As String)
    If n = 0 Then
        AddDiag rid, "error", ln, "extra " & nm
        Exit Sub
    End If
    n = n - 1
    If sk(n) <> kind Then
        AddDiag rid, "error", ln, "block mismatch: open " & sk(n) & " (line " & CStr(sl(n)) & ") vs end " & nm
    Else
        If Len(extra) > 0 And Len(sv(n)) > 0 Then
            If LCase$(extra) <> LCase$(sv(n)) Then
                AddDiag "VB6-E102", "error", ln, "loop variable mismatch: For " & sv(n) & " / Next " & extra
            End If
        End If
    End If
End Sub

Private Function RuleForKind(ByVal k As String) As String
    If k = "if" Then RuleForKind = "VB6-E101"
    If k = "for" Then RuleForKind = "VB6-E102"
    If k = "do" Or k = "while" Then RuleForKind = "VB6-E103"
    If k = "with" Or k = "select" Then RuleForKind = "VB6-E104"
    If k = "type" Then RuleForKind = "VB6-E106"
End Function

Private Function IsBlockIfStart(ByVal L As String) As Boolean
    Dim p As Long, rest As String
    If Left$(L, 3) <> "if " Then Exit Function
    p = InStr(L, " then")
    If p = 0 Then Exit Function
    rest = Trim$(Mid$(L, p + 5))
    If Len(rest) = 0 Then IsBlockIfStart = True
End Function

Private Function IsForStart(ByVal L As String, ByRef var As String) As Boolean
    var = ""
    If Left$(L, 4) <> "for " Then Exit Function
    If Left$(L, 8) = "for each" Then
        var = Tok(L, 3)
        IsForStart = True
        Exit Function
    End If
    If HasWord(L, "to") Or InStr(L, "=") > 0 Then
        var = Tok(L, 2)
        IsForStart = True
    End If
End Function

Private Function NextVar(ByVal L As String) As String
    NextVar = Tok(L, 2)
End Function

Private Function IsTypeStart(ByVal L As String) As Boolean
    Dim s As String
    s = StripMods(L)
    If Left$(s, 5) = "type " Then IsTypeStart = True
    If Left$(s, 5) = "enum " Then IsTypeStart = True
End Function

' ---------------- rules: procedure terminators ----------------

Private Sub RuleProcTerm()
    Dim depth As Long, k As Long, L As String, nm As String
    depth = 0
    For k = 0 To m_nlog - 1
        MaybePump
        L = m_lc(k)
        If m_regR1 > 0 Then
            If m_ls(k) + 1 > m_regR1 Then Exit For
        End If
        If IsProcStart(L, nm) Then
            depth = depth + 1
        ElseIf IsProcEnd(L) Then
            If depth = 0 Then
                AddDiag "VB6-E105", "error", m_ls(k) + 1, "extra " & L & " (no matching procedure)"
            Else
                depth = depth - 1
            End If
        End If
    Next
    If depth > 0 Then AddDiag "VB6-E105", "error", m_n, CStr(depth) & " procedure(s) missing End"
End Sub

' ---------------- rules: duplicate procedures (E108) ----------------

Private Sub RuleDupProcs()
    Dim decl As Object, k As Long, nm As String, kind As String, key As String, prev As String
    Set decl = CreateObject("Scripting.Dictionary")
    For k = 0 To m_nlog - 1
        MaybePump
        If IsProcStart(m_lc(k), nm) Then
            kind = ProcKindOfLine(m_lc(k))
            If Len(nm) > 0 Then
                key = LCase$(nm)
                If decl.Exists(key) Then
                    prev = decl(key)
                    If Not (IsPropKind(prev) And IsPropKind(kind) And prev <> kind) Then
                        AddDiag "VB6-E108", "error", m_ls(k) + 1, "duplicate procedure name: " & nm
                    End If
                Else
                    decl(key) = kind
                End If
            End If
        End If
    Next
End Sub

Private Function IsPropKind(ByVal k As String) As Boolean
    If Left$(k, 9) = "property " Then IsPropKind = True
End Function

Private Function ProcKindOfLine(ByVal L As String) As String
    Dim s As String
    s = StripMods(L)
    If Left$(s, 4) = "sub " Then ProcKindOfLine = "sub" : Exit Function
    If Left$(s, 9) = "function " Then ProcKindOfLine = "function" : Exit Function
    If Left$(s, 9) = "property " Then ProcKindOfLine = "property " & Tok(s, 2) : Exit Function
End Function

' ---------------- rules: stray statements outside procedures (E107) ----------------

Private Sub RuleStray()
    ' E107 is emitted only for a line OUTSIDE every procedure
    ' (Not m_inProc), and a region is exactly one procedure, so no in-region
    ' diagnostic can come from this rule.
    If m_regR1 > 0 Then Exit Sub
    Dim k As Long, L As String, ln As Long, inEnum As Boolean, inBegin As Boolean
    inEnum = False
    inBegin = False
    For k = 0 To m_nlog - 1
        MaybePump
        L = m_lc(k)
        ln = m_ls(k) + 1
        If L = "begin" Or Left$(L, 6) = "begin " Then inBegin = True
        If inBegin Then
            If L = "end" Then inBegin = False
        Else
            ' F-d/D3 (an earlier lint batch): the line still carries its modifiers
            ' ("public enum ...", "private enum ..."), so the un-stripped compare never
            ' matched and the whole Enum body was scanned as if it were module code -
            ' every member declaration was then reported as a stray statement (E107).
            ' StripMods is what the rest of this file already uses.
            If Left$(StripMods(L), 5) = "enum " Then inEnum = True
            If StripMods(L) = "end enum" Then inEnum = False
            If (Not inEnum) And (Not m_inProc(m_ls(k))) Then
                If Left$(L, 1) <> "#" Then
                    If IsStrayStmt(L) Then AddDiag "VB6-E107", "error", ln, "statement outside a procedure"
                End If
            End If
        End If
    Next
End Sub

Private Function IsStrayStmt(ByVal L As String) As Boolean
    Dim s As String, first As String
    If Len(L) = 0 Then Exit Function
    If Left$(L, 8) = "version " Then Exit Function
    If Left$(L, 7) = "object " Then Exit Function
    s = StripMods(L)
    first = Tok(s, 1)
    If IsDeclStartWord(first) Then Exit Function
    ' F-d/D4 (an earlier lint batch): Tok splits on spaces, so "debug.print 1"
    ' yields the first token "debug.print" and the old test first="debug" MISSED it -
    ' a plain Debug.Print outside a procedure was silently green. Any "debug.<member>"
    ' belongs to the same statement family.
    If first = "call" Or first = "debug" Or first = "print" Then IsStrayStmt = True : Exit Function
    If Left$(first, 6) = "debug." Then IsStrayStmt = True : Exit Function
    If first = "if" And Right$(L, 4) = "then" Then IsStrayStmt = True : Exit Function
    If first = "do" Or first = "for" Or first = "while" Or first = "select" Then IsStrayStmt = True : Exit Function
    If InStr(L, "=") > 0 Then IsStrayStmt = True : Exit Function
End Function

Private Function IsDeclStartWord(ByVal w As String) As Boolean
    Select Case w
        Case "option", "attribute", "declare", "type", "enum", "dim", "redim", "const", "event", "implements"
            IsDeclStartWord = True
        Case "sub", "function", "property", "end", "lib", "defint", "deflng", "defdbl", "defsng", "defstr"
            IsDeclStartWord = True
        Case "defbool", "defbyte", "defcur", "defdate", "defobj", "defvar", "note"
            IsDeclStartWord = True
    End Select
End Function

' ---------------- rules: declarations ----------------

Private Sub RuleDecls(ByVal compKind As String)
    Dim p As Long, k As Long, L As String, nm As String
    For p = 0 To m_nproc - 1
        MaybePump
        If ProcOverlapsRegion(p) Then
        Dim decl As Object
        Set decl = CreateObject("Scripting.Dictionary")
        For k = 0 To m_nlog - 1
            MaybePump
            If m_ls(k) >= m_procStart(p) And m_ls(k) <= m_procEnd(p) Then
                L = m_lc(k)
                If IsProcStart(L, nm) Then
                    CollectParams L, m_ls(k) + 1, decl
                ElseIf IsLocalDecl(L) Then
                    CollectDecl L, m_ls(k) + 1, decl
                End If
            End If
        Next
        End If
    Next
    For k = 0 To m_nlog - 1
        MaybePump
        L = m_lc(k)
        ' E110 requires Not m_inProc, i.e. a MODULE-LEVEL line, and a
        ' region is exactly one procedure, so E110 can never be an in-region
        ' diagnostic. Skip its scan in region mode.
        If m_regR1 <= 0 Then
            If m_firstProcLine >= 0 Then
                If m_ls(k) > m_firstProcLine Then
                    If Not m_inProc(m_ls(k)) Then
                        If IsModuleDecl(L) Then
                            AddDiag "VB6-E110", "error", m_ls(k) + 1, "module-level declaration after first procedure"
                        End If
                    End If
                End If
            End If
        End If
        If InReg(m_ls(k) + 1) Then
            If LCase$(compKind) <> "bas" Then
                If InStr(L, "public const ") = 1 Or InStr(L, "public declare ") = 1 Or InStr(L, "public type ") = 1 Or InStr(L, "public enum ") = 1 Then
                    AddDiag "VB6-E111", "error", m_ls(k) + 1, "public module-level declaration not allowed in class/form module"
                End If
            End If
        End If
    Next
End Sub

Private Function IsLocalDecl(ByVal L As String) As Boolean
    If Left$(L, 4) = "dim " Then IsLocalDecl = True
    If Left$(L, 7) = "static " Then IsLocalDecl = True
    If Left$(L, 6) = "const " Then IsLocalDecl = True
End Function

Private Function IsModuleDecl(ByVal L As String) As Boolean
    Dim s As String
    s = StripMods(L)
    If Left$(s, 6) = "const " Then IsModuleDecl = True
    If Left$(s, 8) = "declare " Then IsModuleDecl = True
    If Left$(s, 5) = "type " Then IsModuleDecl = True
    If Left$(s, 5) = "enum " Then IsModuleDecl = True
    If Left$(s, 4) = "dim " Then IsModuleDecl = True
End Function

Private Sub CollectDecl(ByVal L As String, ByVal ln As Long, ByVal decl As Object)
    Dim s As String, parts() As String, i As Long, nm As String
    s = L
    If Left$(s, 4) = "dim " Then s = Mid$(s, 5)
    If Left$(s, 7) = "static " Then s = Mid$(s, 8)
    If Left$(s, 6) = "const " Then s = Mid$(s, 7)
    parts = Split(s, ",")
    For i = 0 To UBound(parts)
        MaybePump
        nm = LeadingIdent(Trim$(parts(i)))
        Register nm, ln, decl
    Next
End Sub

Private Sub CollectParams(ByVal L As String, ByVal ln As Long, ByVal decl As Object)
    Dim p1 As Long, p2 As Long, s As String, parts() As String, i As Long, t As String
    p1 = InStr(L, "(")
    If p1 = 0 Then Exit Sub
    p2 = InStrRev(L, ")")
    If p2 <= p1 Then Exit Sub
    s = Mid$(L, p1 + 1, p2 - p1 - 1)
    If Len(Trim$(s)) = 0 Then Exit Sub
    parts = Split(s, ",")
    For i = 0 To UBound(parts)
        t = StripParamMods(Trim$(parts(i)))
        Register LeadingIdent(t), ln, decl
    Next
End Sub

Private Sub Register(ByVal nm As String, ByVal ln As Long, ByVal decl As Object)
    If Len(nm) = 0 Then Exit Sub
    If decl.Exists(LCase$(nm)) Then
        AddDiag "VB6-E109", "error", ln, "duplicate declaration in procedure: " & nm
    Else
        decl(LCase$(nm)) = True
    End If
    If IsReserved(nm) Then AddDiag "VB6-E112", "error", ln, "reserved word used as identifier: " & nm
End Sub

Private Function LeadingIdent(ByVal s As String) As String
    Dim i As Long, ch As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If (ch >= "a" And ch <= "z") Or (ch >= "A" And ch <= "Z") Or ch = "_" Then
        ElseIf (ch >= "0" And ch <= "9") And i > 1 Then
        Else
            Exit For
        End If
    Next
    LeadingIdent = Left$(s, i - 1)
End Function

Private Function StripParamMods(ByVal s As String) As String
    Dim t As String, again As Boolean
    t = Trim$(s)
    Do
        again = False
        If Left$(t, 6) = "byval " Then t = Mid$(t, 7) : again = True
        If (Not again) And Left$(t, 6) = "byref " Then t = Mid$(t, 7) : again = True
        If (Not again) And Left$(t, 9) = "optional " Then t = Mid$(t, 10) : again = True
        If (Not again) And Left$(t, 11) = "paramarray " Then t = Mid$(t, 12) : again = True
        If Not again Then Exit Do
    Loop
    StripParamMods = t
End Function

Private Function ReservedList() As String
    ReservedList = " name len time date string object left right mid error print loop line input output append base text currency type set get let str val call close open write read stop end sub function property dim const static private public friend option if then else elseif for next do while wend with select case on goto resume exit rem new nothing empty null me true false and or not xor mod is like to step each in as byval byref optional paramarray redim erase lbound ubound variant boolean byte integer long single double "
End Function

Private Function IsReserved(ByVal nm As String) As Boolean
    If Len(nm) = 0 Then Exit Function
    If InStr(1, ReservedList(), " " & LCase$(nm) & " ", vbBinaryCompare) > 0 Then IsReserved = True
End Function

' ---------------- rules: option / lexical / agent ----------------

Private Sub RuleOption()
    Dim k As Long, cnt As Long
    For k = 0 To m_nlog - 1
        MaybePump
        If m_lc(k) = "option explicit" Then
            cnt = cnt + 1
            m_hasOptExp = True
            If cnt > 1 Then AddDiag "VB6-E114", "error", m_ls(k) + 1, "duplicate Option Explicit"
        End If
    Next
End Sub

Private Sub RuleLex()
    Dim i As Long, k As Long, L As String
    For i = 0 To m_n - 1
        MaybePump
        If InReg(i + 1) Then
            If Not StringClosed(m_raw(i)) Then AddDiag "VB6-E113", "error", i + 1, "unterminated string"
        End If
    Next
    For k = 0 To m_nlog - 1
        MaybePump
        If Not InReg(m_ls(k) + 1) Then GoTo LexNext
        L = m_lc(k)
        If Left$(L, 1) <> "#" Then
            If CountCh(L, "(") <> CountCh(L, ")") Then AddDiag "VB6-E113", "error", m_ls(k) + 1, "unbalanced parentheses"
            If Left$(L, 5) = "endif" Then AddDiag "VB6-E115", "error", m_ls(k) + 1, "EndIf must be End If"
            If Left$(L, 6) = "endsub" Then AddDiag "VB6-E115", "error", m_ls(k) + 1, "EndSub must be End Sub"
            If Left$(L, 11) = "endfunction" Then AddDiag "VB6-E115", "error", m_ls(k) + 1, "EndFunction must be End Function"
            If Left$(L, 7) = "endwith" Then AddDiag "VB6-E115", "error", m_ls(k) + 1, "EndWith must be End With"
            If Left$(L, 9) = "endselect" Then AddDiag "VB6-E115", "error", m_ls(k) + 1, "EndSelect must be End Select"
            If Left$(L, 7) = "else if" Then AddDiag "VB6-E115", "error", m_ls(k) + 1, "Else If should be ElseIf"
            If Left$(L, 3) = "if " Then
                If InStr(L, " then") = 0 Then AddDiag "VB6-E115", "error", m_ls(k) + 1, "If missing Then"
            End If
        End If
LexNext:
    Next
End Sub

Private Function StringClosed(ByVal raw As String) As Boolean
    Dim i As Long, inside As Boolean, ch As String, nxt As String
    i = 1
    Do While i <= Len(raw)
        MaybePump
        ch = Mid$(raw, i, 1)
        If ch = Chr$(34) Then
            If inside Then
                nxt = Mid$(raw, i + 1, 1)
                If nxt = Chr$(34) Then
                    i = i + 2
                Else
                    inside = False
                    i = i + 1
                End If
            Else
                inside = True
                i = i + 1
            End If
        ElseIf ch = Chr$(39) And Not inside Then
            Exit Do
        Else
            i = i + 1
        End If
    Loop
    StringClosed = Not inside
End Function

Private Function CountCh(ByVal s As String, ByVal c As String) As Long
    Dim i As Long, n As Long
    For i = 1 To Len(s)
        MaybePump
        If Mid$(s, i, 1) = c Then n = n + 1
    Next
    CountCh = n
End Function

Private Sub RuleAgent()
    Dim k As Long, L As String, ln As Long, s2 As String, inBegin As Boolean
    inBegin = False
    For k = 0 To m_nlog - 1
        MaybePump
        L = m_lc(k)
        ln = m_ls(k) + 1
        If inBegin Then
            If L = "end" Then inBegin = False
            GoTo NextLine
        End If
        If Left$(L, 8) = "version " Or Left$(L, 7) = "object " Then GoTo NextLine
        If L = "begin" Or Left$(L, 6) = "begin " Then
            inBegin = True
            GoTo NextLine
        End If
        If Not InReg(ln) Then GoTo NextLine
        If InStr(L, "==") > 0 Or InStr(L, "!=") > 0 Or InStr(L, "&&") > 0 Or InStr(L, "||") > 0 Or InStr(L, "++") > 0 Or InStr(L, "->") > 0 Or InStr(L, "=>") > 0 Then
            AddDiag "VB6-X201", "error", ln, "non-VB6 C-style operator"
        End If
        If InStr(L, "//") > 0 Or InStr(L, "/*") > 0 Or InStr(L, "*/") > 0 Or InStr(L, "{") > 0 Or InStr(L, "}") > 0 Then
            AddDiag "VB6-X202", "error", ln, "non-VB6 comment/block"
        End If
        If Right$(L, 1) = ";" Then AddDiag "VB6-X202", "error", ln, "trailing semicolon"
        If Left$(L, 1) = "#" Then
            If Left$(L, 3) <> "#if" And Left$(L, 6) <> "#else" And Left$(L, 3) <> "#en" Then
                AddDiag "VB6-X202", "error", ln, "non-VB6 # comment"
            End If
        End If
        If HasForeignKeyword(L) Then AddDiag "VB6-X203", "error", ln, "non-VB6 keyword"
        If InStr(L, "print(") > 0 Or InStr(L, ".length") > 0 Or InStr(L, ".count(") > 0 Then
            AddDiag "VB6-X204", "error", ln, "non-VB6 call/member syntax"
        End If
        If Left$(L, 9) = "for each " Then
            If Not HasWord(L, "in") Then AddDiag "VB6-X205", "error", ln, "For Each without In"
        End If
        s2 = StripMods(L)
        If Left$(s2, 4) = "dim " Or Left$(s2, 7) = "static " Then
            If Right$(s2, 3) = " as" Then AddDiag "VB6-X206", "error", ln, "declaration missing type after As"
        End If
NextLine:
    Next
End Sub

Private Function HasForeignKeyword(ByVal L As String) As Boolean
    Dim toks() As String, i As Long, t As String
    toks = Split(Tokenize(L), " ")
    For i = 0 To UBound(toks)
        MaybePump
        t = toks(i)
        If t = "def" Or t = "import" Or t = "null" Or t = "undefined" Or t = "none" Or t = "self" Or t = "this" Or t = "elif" Or t = "async" Or t = "await" Or t = "try" Or t = "catch" Or t = "finally" Or t = "throw" Then
            HasForeignKeyword = True
            Exit Function
        End If
    Next
End Function

Private Function Tokenize(ByVal s As String) As String
    ' O(n) via a pre-sized buffer + Mid$ assignment (the old form
    ' appended one character at a time = O(n^2)). Output is byte-identical.
    Dim i As Long, n As Long, ch As String, t As String
    n = Len(s)
    If n = 0 Then Exit Function
    t = String$(n, " ")
    For i = 1 To n
        MaybePump
        ch = Mid$(s, i, 1)
        If (ch >= "a" And ch <= "z") Or (ch >= "A" And ch <= "Z") Or (ch >= "0" And ch <= "9") Or ch = "_" Then
            Mid$(t, i, 1) = ch
        End If
    Next
    Tokenize = t
End Function

Private Function RemoveHex(ByVal s As String) As String
    ' O(n) via a pre-sized buffer + Mid$ assignment (the old form
    ' appended one character at a time = O(n^2)). Output is byte-identical.
    Dim i As Long, n As Long, o As Long, ch As String, j As Long, t As String
    n = Len(s)
    If n = 0 Then Exit Function
    t = String$(n, " ")
    i = 1
    Do While i <= n
        ch = Mid$(s, i, 1)
        If ch = "&" And i < n Then
            If LCase$(Mid$(s, i + 1, 1)) = "h" Then
                j = i + 2
                Do While j <= n
                    If InStr("0123456789abcdefABCDEF", Mid$(s, j, 1)) > 0 Then
                        j = j + 1
                    Else
                        Exit Do
                    End If
                Loop
                i = j
            Else
                o = o + 1
                Mid$(t, o, 1) = ch
                i = i + 1
            End If
        Else
            o = o + 1
            Mid$(t, o, 1) = ch
            i = i + 1
        End If
    Loop
    If o = 0 Then RemoveHex = "" Else RemoveHex = Left$(t, o)
End Function

Private Function IsAllCaps(ByVal t As String) As Boolean
    Dim i As Long, ch As String, hasA As Boolean
    hasA = False
    For i = 1 To Len(t)
        MaybePump
        ch = Mid$(t, i, 1)
        If ch >= "A" And ch <= "Z" Then
            hasA = True
        ElseIf ch >= "a" And ch <= "z" Then
            Exit Function
        End If
    Next
    IsAllCaps = hasA
End Function

Private Function CapsKeyword(ByVal t As String) As Boolean
    Dim lst As String
    lst = " AND OR NOT XOR MOD IS LIKE TO STEP EACH IN AS NEW NOTHING EMPTY NULL TRUE FALSE OPTION EXPLICIT IF THEN ELSE ELSEIF FOR NEXT DO WHILE WEND WITH SELECT CASE ON GOTO RESUME EXIT REM DIM CONST STATIC PRIVATE PUBLIC FRIEND SUB FUNCTION PROPERTY SET GET LET CALL END BYVAL BYREF OPTIONAL PARAMARRAY REDIM ERASE LBOUND UBOUND VARIANT BOOLEAN BYTE INTEGER LONG SINGLE DOUBLE STRING OBJECT ME APP CLIPBOARD SCREEN ERR "
    If InStr(1, lst, " " & t & " ", vbBinaryCompare) > 0 Then CapsKeyword = True
End Function

Private Function LogicalRaw(ByVal k As Long) As String
    Dim i As Long, buf As String, piece As String
    For i = m_ls(k) To m_le(k)
        MaybePump
        piece = m_code(i)
        If i < m_le(k) Then
            piece = RTrim$(piece)
            If Right$(piece, 1) = "_" Then piece = Left$(piece, Len(piece) - 1)
        End If
        If Len(Trim$(piece)) > 0 Then
            If Len(buf) > 0 Then buf = buf & " " & Trim$(piece) Else buf = Trim$(piece)
        End If
    Next
    LogicalRaw = buf
End Function

Private Sub RuleUndeclaredConst()
    If Not m_hasOptExp Then Exit Sub
    Dim decl As Object
    Set decl = CreateObject("Scripting.Dictionary")
    m_silent = True
    AddAllDeclared decl
    m_silent = False
    Dim k As Long, i As Long, L As String, toks() As String, t As String
    For k = 0 To m_nlog - 1
        MaybePump
        If Not InReg(m_ls(k) + 1) Then GoTo UcNext
        L = LogicalRaw(k)
        If Left$(L, 1) <> "#" Then
            L = RemoveHex(L)
            toks = Split(Tokenize(L), " ")
            For i = 0 To UBound(toks)
                MaybePump
                t = toks(i)
                If Len(t) >= 3 Then
                    If IsAllCaps(t) Then
                        If Not decl.Exists(LCase$(t)) Then
                            If Not CapsKeyword(t) Then
                                AddDiag "VB6-W301", "warning", m_ls(k) + 1, "undeclared identifier (ALL_CAPS?): " & t
                                Exit For
                            End If
                        End If
                    End If
                End If
            Next
        End If
UcNext:
    Next
End Sub

Private Sub AddAllDeclared(ByVal decl As Object)
    Dim k As Long, L As String, nm As String
    For k = 0 To m_nlog - 1
        MaybePump
        L = m_lc(k)
        If IsProcStart(L, nm) Then
            If Len(nm) > 0 Then decl(LCase$(nm)) = True
            CollectParamsSilent L, decl
        ElseIf IsLocalDecl(L) Or IsModuleDecl(L) Then
            CollectDeclSilent L, decl
        End If
    Next
End Sub

Private Sub CollectDeclSilent(ByVal L As String, ByVal decl As Object)
    Dim s As String, parts() As String, i As Long, nm As String
    s = L
    If Left$(s, 4) = "dim " Then s = Mid$(s, 5)
    If Left$(s, 7) = "static " Then s = Mid$(s, 8)
    If Left$(s, 6) = "const " Then s = Mid$(s, 7)
    parts = Split(s, ",")
    For i = 0 To UBound(parts)
        MaybePump
        nm = LeadingIdent(Trim$(parts(i)))
        If Len(nm) > 0 Then decl(LCase$(nm)) = True
    Next
End Sub

Private Sub CollectParamsSilent(ByVal L As String, ByVal decl As Object)
    Dim p1 As Long, p2 As Long, s As String, parts() As String, i As Long, nm As String
    p1 = InStr(L, "(")
    If p1 = 0 Then Exit Sub
    p2 = InStrRev(L, ")")
    If p2 <= p1 Then Exit Sub
    s = Mid$(L, p1 + 1, p2 - p1 - 1)
    If Len(Trim$(s)) = 0 Then Exit Sub
    parts = Split(s, ",")
    For i = 0 To UBound(parts)
        MaybePump
        nm = LeadingIdent(StripParamMods(Trim$(parts(i))))
        If Len(nm) > 0 Then decl(LCase$(nm)) = True
    Next
End Sub

' ---------------- semantic warnings + output helpers ----------------

Private Sub RuleSemantic()
    Dim k As Long, L As String, ln As Long, nm2 As String
    For k = 0 To m_nlog - 1
        MaybePump
        L = m_lc(k)
        ln = m_ls(k) + 1
        If Not InReg(ln) Then GoTo SemNext
        If IsProcStart(L, nm2) Then CheckArrayByVal L, ln
        If InStr(L, "iif(") > 0 Then AddDiag "VB6-W305", "warning", ln, "IIf evaluates both branches"
        If InStr(L, "on error resume next") > 0 Then AddDiag "VB6-W305", "warning", ln, "On Error Resume Next swallows errors"
SemNext:
    Next
End Sub

Private Sub CheckArrayByVal(ByVal L As String, ByVal ln As Long)
    Dim p1 As Long, p2 As Long, s As String, parts() As String, i As Long, t As String, ty As String, p As Long
    p1 = InStr(L, "(")
    If p1 = 0 Then Exit Sub
    p2 = InStrRev(L, ")")
    If p2 <= p1 Then Exit Sub
    s = Mid$(L, p1 + 1, p2 - p1 - 1)
    parts = Split(s, ",")
    For i = 0 To UBound(parts)
        MaybePump
        t = Trim$(parts(i))
        If InStr(t, "byval ") > 0 And InStr(t, "()") > 0 Then
            p = InStrRev(t, " as ")
            ty = ""
            If p > 0 Then ty = Trim$(Mid$(t, p + 4))
            If ty <> "variant" Then AddDiag "VB6-W304", "warning", ln, "typed array parameter cannot be ByVal"
        End If
    Next
End Sub

Public Function HasErrors() As Boolean
    ' (F2): the counter, not an array scan. Identical answer (m_nerr counts
    ' every diagnostic AddDiag accepted) and it stays correct when storage is capped.
    HasErrors = (m_nerr > 0)
End Function

Public Function DiagnosticsJson() As String
    ' this used to accumulate "s = s & ..." once per diagnostic, which is
    ' O(n^2) on a BSTR - the SAME defect class as JsonLib.Quote / SliceLines /
    ' NormalizeCode / SetCompText / Toks / Sanitize. Measured on the 27,000-line fixture
    ' (1,079 diagnostics) it was 1,516-1,734 ms and was the SINGLE LARGEST segment of
    ' verify/verify_scope. The pieces are now collected into a String() and joined ONCE;
    ' the output is byte-identical (see docs/evidence).
    Dim i As Long, parts() As String
    ' (F2): serialise what is actually STORED. With no store cap (every caller
    ' in this project that uses this function) that is all of them, so the output is
    ' byte-identical to before.
    If m_nst <= 0 Then
        DiagnosticsJson = "[]"
        Exit Function
    End If
    ReDim parts(0 To m_nst - 1)
    For i = 0 To m_nst - 1
        MaybePump
        parts(i) = "{" & Chr$(34) & "ruleId" & Chr$(34) & ":" & Jr(m_diag(i).ruleId) _
            & "," & Chr$(34) & "severity" & Chr$(34) & ":" & Jr(m_diag(i).severity) _
            & "," & Chr$(34) & "line" & Chr$(34) & ":" & CStr(m_diag(i).physLine) _
            & "," & Chr$(34) & "message" & Chr$(34) & ":" & Jr(m_diag(i).msg) _
            & "," & Chr$(34) & "key" & Chr$(34) & ":" & Jr(DiagKey(i)) & "}"
    Next
    DiagnosticsJson = "[" & Join(parts, ",") & "]"
End Function

Private Function Jr(ByVal s As String) As String
    Dim t As String
    t = Replace(s, Chr$(92), Chr$(92) & Chr$(92))
    t = Replace(t, Chr$(34), Chr$(92) & Chr$(34))
    Jr = Chr$(34) & t & Chr$(34)
End Function

' the owner of the procedure BODY that contains physical line ln
' (1-based), or "" when ln is module level or part of a procedure's declaration.
' m_procStart is the procedure's first logical line and m_procEnd its End line, so a
' body line is strictly after the start and at or before the end. This is exact,
' unlike VBIDE's ProcOfLine (which attributes a blank line before a procedure to that
' procedure) and ProcBodyLine (which returns the declaration line here).
Public Function ProcBodyOwner(ByVal ln As Long) As String
    Dim i As Long
    ProcBodyOwner = ""
    ln = ln - 1
    If ln < 0 Then Exit Function
    For i = 0 To m_nproc - 1
        MaybePump
        If ln > m_procStart(i) And ln <= m_procEnd(i) Then
            ProcBodyOwner = m_procName(i)
            Exit Function
        End If
    Next
End Function

Public Function DiagKey(ByVal i As Long) As String
    If i < 0 Or i >= m_nst Then Exit Function
    DiagKey = m_diag(i).ruleId & ":" & NormCtx(m_diag(i).physLine)
End Function

Private Function NormCtx(ByVal physLine As Long) As String
    Dim i As Long, s As String
    i = physLine - 1
    If i < 0 Or i >= m_n Then Exit Function
    s = LCase$(Trim$(m_code(i)))
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    NormCtx = s
End Function

Public Function IdeDiskOffset(ByVal diskLineCount As Long, ByVal ideLineCount As Long) As Long
    IdeDiskOffset = diskLineCount - ideLineCount
End Function

Public Sub SetBaseline(ByVal lines As String)
    If m_baseKeys Is Nothing Then Set m_baseKeys = CreateObject("Scripting.Dictionary")
    m_baseKeys.RemoveAll
    Dim arr() As String, i As Long, t As String
    arr = Split(Replace(lines, vbCrLf, vbLf), vbLf)
    For i = 0 To UBound(arr)
        MaybePump
        t = Trim$(arr(i))
        If Len(t) > 0 Then m_baseKeys(t) = True
    Next
End Sub

Public Function BaselineHits() As Long
    BaselineHits = m_baseHits
End Function

' ===================== F-d (an earlier lint batch) =====================
' D1 (E116) + D2 (E117). Both are NEW diagnostics; no existing rule, message or
' description changed and RULE_COUNT moved by exactly two.
' ========================================================================

' The kind of the procedure a declaration line opens, using IsProcStart's own prefix
' vocabulary (after StripMods, so Public/Private/Friend/Static do not matter).
Private Function ProcKindOf(ByVal L As String) As String
    Dim s As String
    s = StripMods(L)
    If Left$(s, 4) = "sub " Then ProcKindOf = "sub"
    If Left$(s, 9) = "function " Then ProcKindOf = "function"
    If Left$(s, 9) = "property " Then ProcKindOf = "property"
End Function

' Which procedure contains this PHYSICAL line? -1 = none. VB6 has NO nested procedures,
' so this is a total function - that is what makes E116 statically decidable.
Private Function ProcIndexAt(ByVal physLine As Long) As Long
    Dim p As Long
    ProcIndexAt = -1
    For p = 0 To m_nproc - 1
        If physLine >= m_procStart(p) And physLine <= m_procEnd(p) Then
            ProcIndexAt = p
            Exit Function
        End If
    Next
End Function

' ---------------- D1: E116 - Exit <kind> must match the enclosing procedure --------
Private Sub RuleExitKind()
    If m_nproc = 0 Then Exit Sub
    Dim k As Long, s As String, what As String, p As Long
    For k = 0 To m_nlog - 1
        MaybePump
        s = StripMods(m_lc(k))
        If Left$(s, 5) = "exit " Then
            what = Trim$(Mid$(s, 6))
            ' Exit For / Exit Do are loop exits and always legal; anything else is not a
            ' statement this rule knows about.
            If what = "sub" Or what = "function" Or what = "property" Then
                p = ProcIndexAt(m_ls(k))
                If p >= 0 Then
                    If m_procKind(p) <> what Then
                        AddDiag "VB6-E116", "error", m_ls(k) + 1, _
                            "Exit " & UCase$(Left$(what, 1)) & Mid$(what, 2) & " is not valid inside a " & UCase$(Left$(m_procKind(p), 1)) & Mid$(m_procKind(p), 2) & " (" & m_procName(p) & "): a procedure may only be left with its own Exit statement"
                    End If
                End If
            End If
        End If
    Next
End Sub

' Reads the assignee of a plain assignment or of a For-loop initialiser. Returns ""
' when the line is not one of those shapes. Deliberately narrow: the left-hand side
' must be a single bare identifier, so "Foo.Bar = 1", "a(1) = 2" and every comparison
' (<, >, <=, >=, =, <>) is rejected instead of guessed at.
Private Function AssignTarget(ByVal s As String, ByRef isFor As Boolean) As String
    isFor = False
    AssignTarget = ""
    Dim p As Long, lhs As String, c As String
    If Left$(s, 4) = "for " Then
        p = InStr(s, "=")
        If p <= 4 Then Exit Function
        lhs = Trim$(Mid$(s, 5, p - 5))
        If InStr(lhs, " ") > 0 Then Exit Function
        If InStr(lhs, "(") > 0 Then Exit Function
        If InStr(lhs, ".") > 0 Then Exit Function
        If Len(lhs) = 0 Then Exit Function
        isFor = True
        AssignTarget = lhs
        Exit Function
    End If
    p = InStr(s, "=")
    If p <= 1 Then Exit Function
    c = Mid$(s, p - 1, 1)
    If c = "<" Or c = ">" Or c = "=" Or c = "!" Then Exit Function
    If Mid$(s, p + 1, 1) = "=" Then Exit Function
    lhs = Trim$(Left$(s, p - 1))
    If Len(lhs) = 0 Then Exit Function
    If InStr(lhs, " ") > 0 Then Exit Function
    If InStr(lhs, "(") > 0 Then Exit Function
    If InStr(lhs, ".") > 0 Then Exit Function
    AssignTarget = lhs
End Function

' Conservative shadow collector: it records EVERY declarator it can see, because an
' unnoticed shadow would become a false positive.
Private Sub CollectDeclaredNames(ByVal s As String, ByVal into As Object)
    Dim rest As String, parts() As String, i As Long, pn As String, q As Long
    If Left$(s, 4) = "dim " Then rest = Mid$(s, 5)
    If Left$(s, 7) = "static " Then rest = Mid$(s, 8)
    If Left$(s, 6) = "redim " Then rest = Mid$(s, 7)
    If Len(rest) = 0 Then Exit Sub
    parts = Split(rest, ",")
    For i = 0 To UBound(parts)
        pn = Trim$(parts(i))
        q = InStr(pn, "(")
        If q > 0 Then pn = Left$(pn, q - 1)
        q = InStr(pn, " ")
        If q > 0 Then pn = Left$(pn, q - 1)
        pn = Trim$(pn)
        If Len(pn) > 0 Then
            If IsIdentName(pn) Then into(pn) = 1
        End If
    Next
End Sub

Private Function IsIdentName(ByVal s As String) As Boolean
    Dim i As Long, ch As String
    IsIdentName = False
    If Len(s) = 0 Then Exit Function
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If InStr("abcdefghijklmnopqrstuvwxyz0123456789_", ch) = 0 Then Exit Function
    Next
    IsIdentName = True
End Function

' Parameter names of a procedure declaration line ("sub f(a as long, b)").
' F-d: named CollectParamNames, NOT CollectParams - the module already has a
' CollectParams(L, ln, decl) for the E111 declaration audit and VB6 refuses two procedures
' with one name in the same module (build error: variable definition conflict).
Private Sub CollectParamNames(ByVal s As String, ByVal into As Object)
    Dim op As Long, cl As Long, args As String, parts() As String, i As Long, pn As String, q As Long
    op = InStr(s, "(")
    If op = 0 Then Exit Sub
    cl = InStrRev(s, ")")
    If cl <= op Then Exit Sub
    args = Mid$(s, op + 1, cl - op - 1)
    parts = Split(args, ",")
    For i = 0 To UBound(parts)
        pn = Trim$(parts(i))
        q = InStr(pn, "(")
        If q > 0 Then pn = Left$(pn, q - 1)
        q = InStr(pn, " ")
        If q > 0 Then pn = Left$(pn, q - 1)
        pn = Trim$(pn)
        If Len(pn) > 0 Then
            If IsIdentName(pn) Then into(pn) = 1
        End If
    Next
End Sub

' ---------------- D2: E117 - assignment to a read-only constant ----------------
' All FOUR narrowings that an earlier lint batch measured are present, in this order:
'   1. the Enum body is SKIPPED (a member line is a declaration, not an assignment),
'      but the member NAMES are collected, so assigning to one is still caught;
'   2. module-level AND procedure-level Const names are collected;
'   3. For <const> = ... is matched as well as <const> = ...;
'   4. a local Dim/Static/parameter that shadows the name suppresses the report.
' Only a single bare identifier on the left is ever reported, so property/array and
' comparison forms are never guessed at.
Private Sub RuleConstAssign()
    Dim modConsts As Object
    Set modConsts = CreateObject("Scripting.Dictionary")
    Dim procConsts() As Object
    Dim shadows() As Object
    ReDim procConsts(0 To m_nproc)
    ReDim shadows(0 To m_nproc)
    Dim p As Long
    For p = 0 To m_nproc - 1
        Set procConsts(p) = CreateObject("Scripting.Dictionary")
        Set shadows(p) = CreateObject("Scripting.Dictionary")
    Next
    Dim k As Long, L As String, s As String, inEnum As Boolean, cn As String
    inEnum = False
    ' ---- pass 1: the name sets ----
    For k = 0 To m_nlog - 1
        MaybePump
        L = m_lc(k)
        s = StripMods(L)
        If inEnum Then
            If s = "end enum" Then
                inEnum = False
            Else
                ' narrowing 1: a member DECLARATION is a name, never an assignment
                cn = Tok(s, 1)
                If Len(cn) > 0 Then
                    If IsIdentName(cn) Then modConsts(cn) = 1
                End If
            End If
        ElseIf Left$(s, 5) = "enum " Then
            inEnum = True
        ElseIf Left$(s, 6) = "const " Then
            ' narrowing 2: Const is collected wherever it appears (module or procedure)
            cn = Tok(s, 2)
            If Len(cn) > 0 Then
                If IsIdentName(cn) Then
                    p = ProcIndexAt(m_ls(k))
                    If p >= 0 Then
                        procConsts(p)(cn) = 1
                    Else
                        modConsts(cn) = 1
                    End If
                End If
            End If
        Else
            p = ProcIndexAt(m_ls(k))
            If p >= 0 Then
                ' narrowing 4: every name the procedure introduces hides a module const
                If Left$(s, 4) = "dim " Or Left$(s, 7) = "static " Or Left$(s, 6) = "redim " Then
                    CollectDeclaredNames s, shadows(p)
                End If
                If m_ls(k) = m_procStart(p) Then CollectParamNames s, shadows(p)
            End If
        End If
    Next
    ' ---- pass 2: the assignments ----
    inEnum = False
    Dim isFor As Boolean, tgt As String, hit As Boolean
    For k = 0 To m_nlog - 1
        MaybePump
        L = m_lc(k)
        s = StripMods(L)
        If inEnum Then
            If s = "end enum" Then inEnum = False
        ElseIf Left$(s, 5) = "enum " Then
            inEnum = True
        ElseIf Left$(s, 6) = "const " Then
            ' the declaration itself is not an assignment
        Else
            isFor = False
            tgt = AssignTarget(s, isFor)
            If Len(tgt) > 0 Then
                hit = False
                p = ProcIndexAt(m_ls(k))
                If p >= 0 Then
                    If procConsts(p).Exists(tgt) Then hit = True
                    If shadows(p).Exists(tgt) Then hit = False
                End If
                If Not hit Then
                    If modConsts.Exists(tgt) Then
                        If p >= 0 Then
                            If Not shadows(p).Exists(tgt) Then hit = True
                        Else
                            hit = True
                        End If
                    End If
                End If
                If hit Then
                    AddDiag "VB6-E117", "error", m_ls(k) + 1, _
                        "assignment to a read-only constant '" & tgt & "' (a Const or an Enum member cannot be assigned" & IIf(isFor, "; use a local variable for the loop counter", "") & ")"
                End If
            End If
        End If
    Next
End Sub

' ===========================================================================
' (the counting review / the wording review / the structural-gate review)
' Two evidence-driven helpers.
'   LintProcMap          - the caller-supplied STABLE identity of a diagnostic:
'                          enclosing procedure + offset inside it (the counting review).
'   BlockPairingProblem  - the PRE-WRITE structural gate: a block carrying its
'                          own procedure header AND an extra "End Function" was
'                          accepted and written, and only the linter complained
'                          (VB6-E105). The writers now ask this BEFORE any write.
' ===========================================================================

' Strip a trailing VB comment from a code line (quote aware).
Private Function LintStripComment(ByVal s As String) As String
    Dim i As Long, inQ As Boolean, ch As String
    inQ = False
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch = Chr$(34) Then
            inQ = Not inQ
        ElseIf ch = "'" And Not inQ Then
            LintStripComment = Left$(s, i - 1)
            Exit Function
        End If
    Next i
    LintStripComment = s
End Function

' The procedure a HEADER line declares ("" when the line is not a header).
Public Function LintProcHeaderName(ByVal line As String) As String
    Dim s As String, u As String, t As String, p As Long
    s = Trim$(LintStripComment(line))
    If Len(s) = 0 Then Exit Function
    u = UCase$(s)
    Do While Len(s) > 0
        If Left$(u, 7) = "PUBLIC " Then
            s = Mid$(s, 8): u = Mid$(u, 8)
        ElseIf Left$(u, 8) = "PRIVATE " Then
            s = Mid$(s, 9): u = Mid$(u, 9)
        ElseIf Left$(u, 7) = "FRIEND " Then
            s = Mid$(s, 8): u = Mid$(u, 8)
        ElseIf Left$(u, 7) = "STATIC " Then
            s = Mid$(s, 8): u = Mid$(u, 8)
        ElseIf Left$(u, 7) = "PUBLIC " Then
            s = Mid$(s, 8): u = Mid$(u, 8)
        Else
            Exit Do
        End If
    Loop
    If Len(s) = 0 Then Exit Function
    u = UCase$(s)
    If Left$(u, 4) = "SUB " Then
        t = Mid$(s, 5)
    ElseIf Left$(u, 9) = "FUNCTION " Then
        t = Mid$(s, 10)
    ElseIf Left$(u, 9) = "PROPERTY " Then
        t = Trim$(Mid$(s, 10))
        u = UCase$(t)
        If Left$(u, 4) = "GET " Then
            t = Trim$(Mid$(t, 5))
        ElseIf Left$(u, 4) = "LET " Then
            t = Trim$(Mid$(t, 5))
        ElseIf Left$(u, 4) = "SET " Then
            t = Trim$(Mid$(t, 5))
        End If
    Else
        Exit Function
    End If
    t = Trim$(t)
    p = InStr(t, "(")
    If p > 0 Then t = Left$(t, p - 1)
    p = InStr(t, " ")
    If p > 0 Then t = Left$(t, p - 1)
    LintProcHeaderName = Trim$(t)
End Function

' "Sub" / "Function" / "Property" when the line TERMINATES a procedure, else "".
Public Function LintProcEndKind(ByVal line As String) As String
    Dim u As String
    u = UCase$(Trim$(LintStripComment(line)))
    If u = "END SUB" Then
        LintProcEndKind = "Sub"
    ElseIf u = "END FUNCTION" Then
        LintProcEndKind = "Function"
    ElseIf u = "END PROPERTY" Then
        LintProcEndKind = "Property"
    End If
End Function

' One pass over the text: for every line, names(i)/starts(i) are the enclosing
' procedure and the 1-based line it starts on (0 / "" before the first header).
Public Sub LintProcMap(ByVal text As String, ByRef names() As String, ByRef starts() As Long)
    Dim lines() As String
    Dim i As Long, n As Long, cur As String, curStart As Long, hn As String
    lines = Split(Replace(Replace(text, vbCrLf, vbLf), vbCr, vbLf), vbLf)
    n = UBound(lines) + 1
    If n < 1 Then n = 1
    ReDim names(0 To n - 1)
    ReDim starts(0 To n - 1)
    cur = ""
    curStart = 0
    For i = 0 To n - 1
        hn = LintProcHeaderName(lines(i))
        If Len(hn) > 0 Then
            cur = hn
            curStart = i + 1
        End If
        names(i) = cur
        starts(i) = curStart
        If Len(LintProcEndKind(lines(i))) > 0 Then
            cur = ""
            curStart = 0
        End If
    Next i
End Sub

' The PRE-WRITE pairing gate. "" = the block is structurally acceptable.
' Only a block that OPENS at least one procedure can carry a redundant closer: a
' bare fragment (0 headers) is deliberately left alone, because closing a
' procedure opened by an EARLIER call is a legal flow and refusing it would break
' writes that work today (negative control).
Public Function BlockPairingProblem(ByVal blockText As String, ByRef headers As Long, ByRef ends As Long, _
                                    ByRef extraEnds As Long, ByRef suggestedText As String) As String
    Dim lines() As String
    Dim i As Long, n As Long
    headers = 0
    ends = 0
    extraEnds = 0
    suggestedText = ""
    BlockPairingProblem = ""
    lines = Split(Replace(Replace(blockText, vbCrLf, vbLf), vbCr, vbLf), vbLf)
    n = UBound(lines) + 1
    If n < 1 Then Exit Function
    Dim nh As Long, ne As Long
    nh = 0
    ne = 0
    For i = 0 To n - 1
        If Len(LintProcHeaderName(lines(i))) > 0 Then
            nh = nh + 1
        ElseIf Len(LintProcEndKind(lines(i))) > 0 Then
            ne = ne + 1
        End If
    Next i
    headers = nh
    ends = ne
    If nh = 0 Then Exit Function
    If ne <= nh Then Exit Function
    extraEnds = ne - nh
    BlockPairingProblem = "extra_procedure_end"
    ' The executable fix: drop the LAST extraEnds terminator lines.
    Dim skip() As Boolean, drop As Long, k As Long
    ReDim skip(0 To n - 1)
    drop = extraEnds
    For i = n - 1 To 0 Step -1
        If drop > 0 Then
            If Len(LintProcEndKind(lines(i))) > 0 Then
                skip(i) = True
                drop = drop - 1
            End If
        End If
    Next i
    Dim keep() As String
    ReDim keep(0 To n - 1)
    k = 0
    For i = 0 To n - 1
        If Not skip(i) Then
            keep(k) = lines(i)
            k = k + 1
        End If
    Next i
    If k > 0 Then
        ReDim Preserve keep(0 To k - 1)
        suggestedText = Join(keep, vbCrLf)
    End If
End Function
