Attribute VB_Name = "EditTxn"
Option Explicit

' Middle blocks larger than this are diffed with the classic block replace instead
' of the LCS DP (bounded work in the IDE process).
Private Const LCS_MAX_LINES As Long = 1000

' TASK-P1-098: the diff TEXT is emitted through ONE place so it can be CAPPED while it is
' built (never materialising a megabyte the caller cannot use). mDiffCap = 0 means
' unlimited, which is what every caller that does not pass a cap gets.
Private mDiffCap As Long
Private mDiffLines As Long
Private mDiffTotal As Long

' TASK-P1-024 (B+C batch, contract invariant 8): preview_edit's diff is a
' LINE-LEVEL LCS with hunk grouping. The A batch only trimmed a common
' prefix/suffix, which renders a middle insertion as "delete a region + add a
' longer region" - that is not a diff. A pure middle insertion now renders as
' ONE hunk of +N with no delete hunk at all.
'
' Anchor ops are resolved by the caller (anchor + occurrence + scope + match ->
' a concrete 1-based IDE line) and applied BOTTOM-UP so that an earlier line
' number stays valid while later lines grow or shrink. The whole batch is
' applied to one text and written once, so apply_edits stays all-or-nothing.

' ops entries (resolved): "R" & TAB & lineNo & TAB & newLine
'                         "I" & TAB & lineNo & TAB & inserted   (Chr(1) = newline)
'                         "D" & TAB & lineNo
' TASK-P1-082: BuildOps (CodeTools) escapes a literal TAB inside a payload as Chr$(2)
' so the TAB-separated record keeps its field count. This is the matching decode for
' the preview_edit path, which parses that record right here.
Private Function OpUnescTxn(ByVal s As String) As String
    OpUnescTxn = Replace(s, Chr$(2), vbTab)
End Function
Public Function ApplyLineOps(ByVal text As String, ByRef ops() As String, ByRef errMsg As String) As String
    ' TASK-P1-097: ONE line array, ONE Join.
    '
    ' The old body rebuilt the WHOLE text once per op - ReplaceLineAt / InsertAfterLine /
    ' DeleteLineAt each did Split(AsLf(text), vbLf) ... Join(parts, vbCrLf) - and the range
    ' check called CountLines(cur) TWICE per op (two more whole-text splits). Measured on
    ' the 27,011-line fixture: 84-92 ms per op (s=50 -> 4,244 ms, s=74 -> 6,796 ms).
    '
    ' EQUIVALENCE. For every op the old code did
    '     parts = Split(AsLf(cur), vbLf)   ...   cur = Join(parts, vbCrLf)
    ' and Join(parts, vbCrLf) -> AsLf -> Split(vbLf) is the IDENTITY on parts: no element of
    ' parts can contain CR or LF (parts came from a vbLf split of an LF-only string, and the
    ' payload lines come from the same kind of split). The old chain is therefore exactly a
    ' sequence of in-place edits on ONE array, wrapped in one Split and one Join of dead
    ' weight per op. Keeping that array and joining ONCE cannot change a byte. The value
    ' CountLines(cur) measured is the array length, which is now read off the arrays (the
    ' count formula below is the document length at that point in the sequence).
    '
    ' The ops are already sorted DESCENDING by line (same stable bubble sort, now over
    ' pre-parsed line numbers), and in that order every op only touches the line at its own
    ' position plus the lines it inserts immediately after it. Same-line groups keep their
    ' input order, which is what makes D-spill (D, then another op after the region ran
    ' empty) behave exactly like the old per-op rebuild - see REP-TASK-P1-097 for the case
    ' table (R/R, R/I, I/R, I/D, D/D, D-spill).
    Dim i As Long, k As Long, j As Long, n As Long, lb As Long
    Dim order() As Long, tmp As Long, parts() As String
    Dim oOp() As String, oLn() As Long, oPay() As String
    Dim op As String, ln As Long, payload As String
    Dim o As Long
    errMsg = ""
    n = 0
    On Error Resume Next
    lb = LBound(ops)
    n = UBound(ops) - lb + 1
    If Err.Number <> 0 Then n = 0
    On Error GoTo 0
    If n <= 0 Then
        ApplyLineOps = text
        Exit Function
    End If
    ' ---- 1. parse every op ONCE (the old sort re-split both records on every compare) ---
    ReDim oOp(0 To n - 1)
    ReDim oLn(0 To n - 1)
    ReDim oPay(0 To n - 1)
    For i = 0 To n - 1
        MaybePump
        ' TASK-P1-082: limit 3 keeps the payload - the LAST field - intact when it holds
        ' literal TABs, so edit_by_anchor / apply_edits no longer truncate at a TAB.
        parts = Split(ops(lb + i), vbTab, 3)
        If UBound(parts) < 1 Then
            oOp(i) = Chr$(1)          ' malformed: reported in SORTED order, as before
            oLn(i) = 0
            oPay(i) = ""
        Else
            oOp(i) = parts(0)
            oLn(i) = 0
            On Error Resume Next
            oLn(i) = CLng(parts(1))
            On Error GoTo 0
            oPay(i) = ""
            If UBound(parts) >= 2 Then oPay(i) = Replace(parts(2), Chr$(1), vbCrLf)
        End If
    Next
    ' ---- 2. descending STABLE sort (byte-identical to the old bubble sort) -------------
    ReDim order(0 To n - 1)
    For i = 0 To n - 1
        order(i) = i
    Next
    For i = 0 To n - 2
        MaybePump
        For j = 0 To n - 2 - i
            If oLn(order(j)) < oLn(order(j + 1)) Then
                tmp = order(j)
                order(j) = order(j + 1)
                order(j + 1) = tmp
            End If
        Next
    Next
    ' ---- 3. the working arrays (one original array, one reverse accumulator, one region) -
    Dim orig() As String, nOrig As Long, p As Long
    orig = Split(AsLf(text), vbLf)
    nOrig = LineUbound(orig) + 1
    If nOrig <= 0 Then
        nOrig = 0
        ReDim orig(0 To 0)
    End If
    p = nOrig - 1
    Dim cap As Long, acc() As String, nAcc As Long
    cap = nOrig + n + 8
    For i = 0 To n - 1
        cap = cap + Len(ops(lb + i)) + 1
    Next
    ReDim acc(0 To cap)
    nAcc = 0
    Dim reg() As String, nReg As Long
    ReDim reg(0 To 7)
    nReg = 0
    Dim groupLn As Long, curCount As Long
    groupLn = -1
    For k = 0 To n - 1
        MaybePump
        o = order(k)
        op = oOp(o)
        ln = oLn(o)
        payload = oPay(o)
        If op = Chr$(1) Then
            errMsg = "bad op at " & CStr(lb + o)
            ApplyLineOps = ""
            Exit Function
        End If
        If ln <> groupLn Then
            ' close the previous group, then open this one at ORIGINAL line ln
            If groupLn >= 0 Then
                CommitRev acc, nAcc, reg, nReg
                nReg = 0
            End If
            Do While p >= ln
                acc(nAcc) = orig(p)
                nAcc = nAcc + 1
                p = p - 1
            Loop
            curCount = (p + 1) + nAcc
            If ln < 1 Or ln > curCount Then
                errMsg = "line " & CStr(ln) & " out of range (1.." & CStr(curCount) & ")"
                ApplyLineOps = ""
                Exit Function
            End If
            reg(0) = orig(p)
            nReg = 1
            p = p - 1
            groupLn = ln
        End If
        If nReg = 0 Then
            ' the region at this line ran empty: the next line of the document is the tail
            If nAcc > 0 Then
                nAcc = nAcc - 1
                reg(0) = acc(nAcc)
                nReg = 1
            Else
                curCount = (p + 1) + nAcc
                errMsg = "line " & CStr(ln) & " out of range (1.." & CStr(curCount) & ")"
                ApplyLineOps = ""
                Exit Function
            End If
        End If
        If op = "R" Then
            RegReplace reg, nReg, payload
        ElseIf op = "I" Then
            RegInsertAfter reg, nReg, payload
        ElseIf op = "D" Then
            For i = 1 To nReg - 1
                reg(i - 1) = reg(i)
            Next
            nReg = nReg - 1
        Else
            errMsg = "unknown op " & op
            ApplyLineOps = ""
            Exit Function
        End If
    Next
    If groupLn >= 0 Then CommitRev acc, nAcc, reg, nReg
    Do While p >= 0
        acc(nAcc) = orig(p)
        nAcc = nAcc + 1
        p = p - 1
    Loop
    If nAcc <= 0 Then
        ApplyLineOps = ""
        Exit Function
    End If
    Dim fwd() As String
    ReDim fwd(0 To nAcc - 1)
    For i = 0 To nAcc - 1
        fwd(i) = acc(nAcc - 1 - i)
    Next
    ApplyLineOps = Join(fwd, vbCrLf)
End Function

' TASK-P1-097: replace the first line of the current region with the payload's lines (R).
' An EMPTY payload drops the line, exactly like the old ReplaceLineAt (whose inner loop
' ran zero times when the payload split came back empty).
Private Sub RegReplace(ByRef reg() As String, ByRef nReg As Long, ByVal payload As String)
    Dim ins() As String, ni As Long, i As Long
    ins = Split(AsLf(payload), vbLf)
    ni = LineUbound(ins)
    If ni < 0 Then
        For i = 1 To nReg - 1
            reg(i - 1) = reg(i)
        Next
        nReg = nReg - 1
        Exit Sub
    End If
    RegEnsure reg, nReg + ni
    For i = nReg - 1 To 1 Step -1
        reg(i + ni) = reg(i)
    Next
    For i = 0 To ni
        reg(i) = ins(i)
    Next
    nReg = nReg + ni
End Sub

' TASK-P1-097: insert the payload's lines right AFTER the region's first line (I). The
' tail moves first so the payload keeps its own order; an empty payload inserts nothing.
Private Sub RegInsertAfter(ByRef reg() As String, ByRef nReg As Long, ByVal payload As String)
    Dim ins() As String, ni As Long, i As Long
    ins = Split(AsLf(payload), vbLf)
    ni = LineUbound(ins)
    If ni < 0 Then Exit Sub
    RegEnsure reg, nReg + ni + 1
    For i = nReg - 1 To 1 Step -1
        reg(i + ni + 1) = reg(i)
    Next
    For i = 0 To ni
        reg(1 + i) = ins(i)
    Next
    nReg = nReg + ni + 1
End Sub

Private Sub RegEnsure(ByRef reg() As String, ByVal need As Long)
    Dim u As Long
    u = LineUbound(reg)
    If u >= need Then Exit Sub
    If u < 7 Then u = 7
    Do While u < need
        u = u * 2 + 1
    Loop
    ReDim Preserve reg(0 To u)
End Sub

' TASK-P1-097: push the region onto the reverse accumulator. acc is kept in REVERSE
' document order (acc(0) is the LAST line of the result), so reg's last line goes first.
Private Sub CommitRev(ByRef acc() As String, ByRef nAcc As Long, ByRef reg() As String, ByVal nReg As Long)
    Dim i As Long
    For i = nReg - 1 To 0 Step -1
        acc(nAcc) = reg(i)
        nAcc = nAcc + 1
    Next
End Sub

' TASK-P1-049: UBound of a Split result, safely -1 for an empty array.
Private Function LineUbound(ByRef arr() As String) As Long
    Dim u As Long
    u = -1
    On Error Resume Next
    u = UBound(arr)
    On Error GoTo 0
    LineUbound = u
End Function

' ---- line-level LCS diff (contract invariant 8) ----------------------------
' Returns a plan: algorithm, diff (unified-ish text), hunks (Collection of
' {kind,start,count,oldStart,oldCount,newStart,newCount,contextLines}).
Public Function DiffPlan(ByVal oldText As String, ByVal newText As String, ByVal contextLines As Long, Optional ByVal maxDiffLines As Long = 0) As Object
    Dim a() As String, b() As String, plan As Object, hunks As Collection
    mDiffCap = maxDiffLines
    mDiffLines = 0
    mDiffTotal = 0
    Dim pre As Long, suf As Long, na As Long, nb As Long, am As Long, bm As Long
    a = Split(AsLf(oldText), vbLf)
    b = Split(AsLf(newText), vbLf)
    na = UBound(a) + 1
    nb = UBound(b) + 1
    TrimCommon a, b, pre, suf, na, nb
    am = na - pre - suf
    bm = nb - pre - suf
    If am < 0 Then am = 0
    If bm < 0 Then bm = 0
    Set plan = CreateObject("Scripting.Dictionary")
    Set hunks = New Collection
    If am = 0 And bm = 0 Then
        plan("algorithm") = "none"
        plan("diff") = ""
        Set plan("hunks") = hunks
        Set DiffPlan = plan
        Exit Function
    End If
    If am > LCS_MAX_LINES Or bm > LCS_MAX_LINES Then
        ' Honest fallback for a huge middle block: the classic block replace.
        If am > 0 Then hunks.Add HunkItem2("delete", pre + 1, am, pre + 1, am, pre + 1, 0, contextLines)
        If bm > 0 Then hunks.Add HunkItem2("add", pre + 1, bm, pre + 1, 0, pre + 1, bm, contextLines)
        plan("algorithm") = "fallback-too-large"
        plan("diff") = BlockDiff(a, b, pre, am, bm)
        DiffMark plan
        plan("note") = "middle block exceeds " & CStr(LCS_MAX_LINES) & " lines; fell back to a block replace"
        Set plan("hunks") = hunks
        Set DiffPlan = plan
        Exit Function
    End If
    If am = 0 Or bm = 0 Then
        ' Pure insert / pure delete: one hunk, no DP needed.
        If am > 0 Then hunks.Add HunkItem2("delete", pre + 1, am, pre + 1, am, pre + 1, 0, contextLines)
        If bm > 0 Then hunks.Add HunkItem2("add", pre + 1, bm, pre + 1, 0, pre + 1, bm, contextLines)
        plan("algorithm") = "lcs"
        plan("diff") = BlockDiff(a, b, pre, am, bm)
        DiffMark plan
        Set plan("hunks") = hunks
        Set DiffPlan = plan
        Exit Function
    End If
    Dim c() As Long, i As Long, j As Long
    ReDim c(0 To am, 0 To bm)
    For i = am - 1 To 0 Step -1
        For j = bm - 1 To 0 Step -1
            If a(pre + i) = b(pre + j) Then
                c(i, j) = c(i + 1, j + 1) + 1
            ElseIf c(i + 1, j) >= c(i, j + 1) Then
                c(i, j) = c(i + 1, j)
            Else
                c(i, j) = c(i, j + 1)
            End If
        Next
    Next
    Dim steps() As String, sol() As Long, snl() As Long, ns As Long
    ReDim steps(0 To am + bm)
    ReDim sol(0 To am + bm)
    ReDim snl(0 To am + bm)
    ns = 0
    i = 0
    j = 0
    ' TASK-P1-099: the bounds MUST be tested in their own branches. VB6's And/Or are
    ' BITWISE, not short-circuit, so the old single condition
    '     If i < am And j < bm And a(pre + i) = b(pre + j) Then
    ' evaluated a(pre + i) = b(pre + j) even when i = am or j = bm. When the middle block
    ' reaches the end of the text (suf = 0) pre + am = na = UBound(a) + 1 and
    ' pre + bm = nb = UBound(b) + 1, so that evaluation read one element past the array and
    ' raised VB error 9 (subscript out of range). MEASURED on instrumented build 4383e5cf:
    '   stage=S:walk.if nErr=9 desc=subscript out of range
    '   am=1 bm=1 pre=1001 suf=0 na=1002 nb=1002 ns=1 i=0 j=1 ubA=1001 ubB=1001
    '   -> b(pre + j) = b(1002) while UBound(b) = 1001.
    ' The branches below are the SAME decision tree with the bounds tested first, so every
    ' in-range input takes exactly the branch it took before and emits the same steps
    ' (byte-identical diffs); the out-of-range read is now unreachable.
    Do While i < am Or j < bm
        If i >= am Then
            steps(ns) = "+"
            sol(ns) = 0
            snl(ns) = pre + j + 1
            ns = ns + 1
            j = j + 1
        ElseIf j >= bm Then
            steps(ns) = "-"
            sol(ns) = pre + i + 1
            snl(ns) = 0
            ns = ns + 1
            i = i + 1
        ElseIf a(pre + i) = b(pre + j) Then
            steps(ns) = "="
            sol(ns) = pre + i + 1
            snl(ns) = pre + j + 1
            ns = ns + 1
            i = i + 1
            j = j + 1
        ElseIf c(i + 1, j) <= c(i, j + 1) Then
            steps(ns) = "+"
            sol(ns) = 0
            snl(ns) = pre + j + 1
            ns = ns + 1
            j = j + 1
        Else
            steps(ns) = "-"
            sol(ns) = pre + i + 1
            snl(ns) = 0
            ns = ns + 1
            i = i + 1
        End If
    Loop
    Dim outC As Collection, s As Long, st As Long, en As Long, t As Long
    Dim dCount As Long, aCount As Long, oStart As Long, nStart As Long
    Dim lastCtx As Long, lead As Long, trail As Long, tc As Long
    Set outC = New Collection
    lastCtx = -1
    s = 0
    Do While s < ns
        If steps(s) = "=" Then
            s = s + 1
        Else
            st = s
            Do While s < ns And steps(s) <> "="
                s = s + 1
            Loop
            en = s - 1
            dCount = 0
            aCount = 0
            oStart = 0
            nStart = 0
            For t = st To en
                MaybePump
                If steps(t) = "-" Then
                    dCount = dCount + 1
                    If oStart = 0 Then oStart = sol(t)
                Else
                    aCount = aCount + 1
                    If nStart = 0 Then nStart = snl(t)
                End If
            Next
            If oStart = 0 Then oStart = PrevEqualOld(steps, sol, ns, st)
            If nStart = 0 Then nStart = PrevEqualNew(steps, snl, ns, st)
            If dCount > 0 Then hunks.Add HunkItem2("delete", oStart, dCount, oStart, dCount, nStart, 0, contextLines)
            If aCount > 0 Then hunks.Add HunkItem2("add", nStart, aCount, oStart, 0, nStart, aCount, contextLines)
            DiffAdd outC, "@@ -" & CStr(oStart) & "," & CStr(dCount) & " +" & CStr(nStart) & "," & CStr(aCount) & " @@"
            lead = st - contextLines
            If lead < 0 Then lead = 0
            If lead <= lastCtx Then lead = lastCtx + 1
            For t = lead To st - 1
                If steps(t) = "=" Then
                    DiffAdd outC, " " & a(sol(t) - 1)
                    lastCtx = t
                End If
            Next
            For t = st To en
                If steps(t) = "-" Then DiffAdd outC, "-" & a(sol(t) - 1)
            Next
            For t = st To en
                If steps(t) = "+" Then DiffAdd outC, "+" & b(snl(t) - 1)
            Next
            trail = en + 1
            tc = 0
            Do While trail < ns And steps(trail) = "=" And tc < contextLines
                DiffAdd outC, " " & a(sol(trail) - 1)
                lastCtx = trail
                trail = trail + 1
                tc = tc + 1
            Loop
        End If
    Loop
    plan("algorithm") = "lcs"
    DiffFinish plan, outC
    Set plan("hunks") = hunks
    Set DiffPlan = plan
End Function

' TASK-P1-098: every DiffPlan exit that builds a diff text ends here, so the truncation
' markers are reported on exactly the paths that cap. Keys are ONLY added when the cap
' actually cut something: an uncapped diff keeps its historical key set byte for byte.
Private Sub DiffFinish(ByRef plan As Object, ByRef c As Collection)
    plan("diff") = JoinLines(c)
    DiffMark plan
End Sub

Private Sub DiffMark(ByRef plan As Object)
    If mDiffCap > 0 And mDiffTotal > mDiffLines Then
        plan("diffTruncated") = True
        plan("diffLinesShown") = mDiffLines
        plan("diffLinesTotal") = mDiffTotal
        plan("diffLimitLines") = mDiffCap
    End If
End Sub

Public Function PreviewDiff(ByVal oldText As String, ByVal newText As String, Optional ByVal contextLines As Long = 0) As String
    Dim p As Object
    Set p = DiffPlan(oldText, newText, contextLines)
    PreviewDiff = p("diff")
End Function

Public Function PreviewHunks(ByVal oldText As String, ByVal newText As String, ByVal contextLines As Long) As Object
    Dim p As Object
    Set p = DiffPlan(oldText, newText, contextLines)
    Set PreviewHunks = p("hunks")
End Function

Public Function PreviewAlgorithm(ByVal oldText As String, ByVal newText As String) As String
    Dim p As Object
    Set p = DiffPlan(oldText, newText, 0)
    PreviewAlgorithm = p("algorithm")
End Function

Private Sub TrimCommon(ByRef a() As String, ByRef b() As String, ByRef pre As Long, ByRef suf As Long, ByRef na As Long, ByRef nb As Long)
    pre = 0
    Do While pre < na And pre < nb
        If a(pre) <> b(pre) Then Exit Do
        pre = pre + 1
    Loop
    suf = 0
    Do While suf < (na - pre) And suf < (nb - pre)
        If a(na - 1 - suf) <> b(nb - 1 - suf) Then Exit Do
        suf = suf + 1
    Loop
End Sub

Private Function PrevEqualOld(ByRef steps() As String, ByRef sol() As Long, ByVal ns As Long, ByVal st As Long) As Long
    Dim t As Long
    For t = st - 1 To 0 Step -1
        If steps(t) = "=" Then
            PrevEqualOld = sol(t) + 1
            Exit Function
        End If
    Next
    PrevEqualOld = 1
End Function

Private Function PrevEqualNew(ByRef steps() As String, ByRef snl() As Long, ByVal ns As Long, ByVal st As Long) As Long
    Dim t As Long
    For t = st - 1 To 0 Step -1
        If steps(t) = "=" Then
            PrevEqualNew = snl(t) + 1
            Exit Function
        End If
    Next
    PrevEqualNew = 1
End Function

Private Sub DiffAdd(ByRef c As Collection, ByVal line As String)
    mDiffTotal = mDiffTotal + 1
    If mDiffCap <= 0 Or mDiffLines < mDiffCap Then
        c.Add line
        mDiffLines = mDiffLines + 1
    End If
End Sub

Private Function BlockDiff(ByRef a() As String, ByRef b() As String, ByVal pre As Long, ByVal am As Long, ByVal bm As Long) As String
    Dim outC As Collection, i As Long
    Set outC = New Collection
    For i = pre To pre + am - 1
        MaybePump
        DiffAdd outC, "-" & a(i)
    Next
    For i = pre To pre + bm - 1
        MaybePump
        DiffAdd outC, "+" & b(i)
    Next
    BlockDiff = JoinLines(outC)
End Function

Private Function HunkItem2(ByVal kind As String, ByVal startLine As Long, ByVal count As Long, _
                           ByVal oldStart As Long, ByVal oldCount As Long, _
                           ByVal newStart As Long, ByVal newCount As Long, _
                           ByVal contextLines As Long) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("kind") = kind
    d("start") = startLine
    d("count") = count
    d("oldStart") = oldStart
    d("oldCount") = oldCount
    d("newStart") = newStart
    d("newCount") = newCount
    d("contextLines") = contextLines
    Set HunkItem2 = d
End Function

Private Function JoinLines(ByVal c As Collection) As String
    ' TASK-P1-049: collect + ONE Join (the old "s = s & ..." is quadratic). The
    ' trailing CRLF is preserved by appending "" as a final part, so
    ' Join(parts, vbCrLf) == the old concatenation byte for byte.
    '
    ' TASK-P1-098: the collection was walked BY INDEX ("parts(i-1) = c(i)") and VB6
    ' Collection index access is a linear walk, so this loop was O(n^2) in the number of
    ' diff lines. MEASURED on the 27,011-line fixture: 7,321 diff lines -> 578 ms, 49,921
    ' diff lines -> 4,738 ms; on the 25,899-line procedure fixture a 1.20 MB diff cost
    ' 4,760 ms even at s=50. For Each uses the Collection's enumerator (O(n)) and produces
    ' the identical part list, so every byte of every diff is unchanged.
    If c.Count <= 0 Then
        JoinLines = ""
        Exit Function
    End If
    Dim parts() As String, i As Long, v As Variant
    ReDim parts(0 To c.Count)
    i = 0
    For Each v In c
        MaybePump
        parts(i) = CStr(v)
        i = i + 1
    Next
    parts(c.Count) = ""
    JoinLines = Join(parts, vbCrLf)
End Function

Private Function AsLf(ByVal text As String) As String
    AsLf = Replace(Replace(text, vbCrLf, vbLf), vbCr, vbLf)
End Function

' Backward-compatible entry point (the A batch call shape): resolves every anchor
' op through VbAnchor (normalized match, uniqueness required) and applies the batch
' through ApplyLineOps. The B+C batch layers occurrence/scope/match on top by
' resolving before it calls in; this keeps any existing caller working unchanged.
Public Function ApplyEdits(ByVal text As String, ByRef ops() As String, ByVal expectedSha As String, ByRef errMsg As String) As String
    Dim cur As String, i As Long, parts() As String, lineOps() As String
    Dim op As String, anchor As String, payload As String, hits() As Long, cnt As Long
    Dim ln As Long, k As Long, kk As Long
    cur = text
    errMsg = ""
    If Len(expectedSha) > 0 Then
        If Sha256HexTextUtf8(cur) <> UCase$(expectedSha) Then
            errMsg = "sha mismatch"
            ApplyEdits = ""
            Exit Function
        End If
    End If
    k = 0
    On Error Resume Next
    k = UBound(ops) - LBound(ops) + 1
    If Err.Number <> 0 Then k = 0
    On Error GoTo 0
    If k <= 0 Then
        ApplyEdits = cur
        Exit Function
    End If
    ReDim lineOps(0 To k - 1)
    ' TASK-P1-094: ONE O(N) normalisation pass for the whole batch. Each op used to call
    ' VbAnchor.FindHits with the full text, i.e. Split() + a NormLine per line PER OP
    ' (O(s x N) allocations/normalisations: 16,452 ms at s=50 and 33,236 ms at s=500 -
    ' the latter killed the connection). The index answers a unique anchor in O(1); an
    ' ambiguous one still uses the exact linear scan, so hit lists stay identical.
    Dim nraw() As String, nnorm() As String, idxFirst As Object, idxCnt As Object
    Dim useIdx As Boolean
    useIdx = (k >= 4)   ' below that the one-pass index cannot pay for itself; keep the
                        ' historical single-op path byte-for-byte
    If useIdx Then VbAnchor.BuildNormIndex cur, nraw, nnorm, idxFirst, idxCnt
    kk = 0
    For i = 0 To k - 1
        MaybePump
        If Len(Trim$(ops(LBound(ops) + i))) = 0 Then GoTo NextOp
        parts = Split(ops(LBound(ops) + i), vbTab)
        If UBound(parts) < 1 Then
            errMsg = "bad op at " & CStr(i)
            ApplyEdits = ""
            Exit Function
        End If
        op = parts(0)
        ' TASK-P1-082: decode the escaped free-text fields produced by BuildOps.
        anchor = OpUnescTxn(parts(1))
        payload = ""
        If UBound(parts) >= 2 Then payload = OpUnescTxn(parts(2))
        If useIdx Then
            VbAnchor.FindHitsIndexed nnorm, idxFirst, idxCnt, anchor, hits, cnt
        Else
            VbAnchor.FindHits cur, anchor, hits, cnt
        End If
        If cnt = 0 Then
            errMsg = "op " & CStr(i) & " failed: anchor not found"
            ApplyEdits = ""
            Exit Function
        End If
        If cnt > 1 Then
            errMsg = "op " & CStr(i) & " failed: anchor not unique (" & CStr(cnt) & ")"
            ApplyEdits = ""
            Exit Function
        End If
        ln = hits(0)
        lineOps(kk) = op & vbTab & CStr(ln) & vbTab & payload
        kk = kk + 1
NextOp:
    Next
    If kk = 0 Then
        ApplyEdits = cur
        Exit Function
    End If
    If kk < k Then ReDim Preserve lineOps(0 To kk - 1)
    ApplyEdits = ApplyLineOps(cur, lineOps, errMsg)
End Function
